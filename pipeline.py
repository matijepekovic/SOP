from __future__ import annotations

import logging
import re
import threading
from dataclasses import dataclass
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any

from sop_reporter.email_client import EmailRecord, GmailIMAPClient
from sop_reporter.extractor import ExtractionEngine
from sop_reporter.printer import ExcelPrinter
from sop_reporter.report_builder import ReportBuilder
from sop_reporter.state_store import ProcessedStateStore


LOGGER = logging.getLogger(__name__)
INVALID_FILENAME_CHARACTERS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


class RunStatus(str, Enum):
    SUCCESS = "success"
    GENERATED = "generated"
    NO_MESSAGES = "no_messages"
    NO_DATA = "no_data"
    ALREADY_RUNNING = "already_running"
    FAILED = "failed"


@dataclass(frozen=True)
class RunResult:
    status: RunStatus
    message: str
    report_path: Path | None = None
    report_paths: tuple[Path, ...] = ()
    email_count: int = 0
    attachment_count: int = 0
    row_count: int = 0
    error: Exception | None = None


class JobRunner:
    def __init__(
        self,
        *,
        email_client: GmailIMAPClient,
        state_store: ProcessedStateStore,
        extractor: ExtractionEngine,
        report_builder: ReportBuilder,
        printer: ExcelPrinter,
        downloads_dir: Path,
        reports_dir: Path,
        report_filename: str,
    ) -> None:
        self.email_client = email_client
        self.state_store = state_store
        self.extractor = extractor
        self.report_builder = report_builder
        self.printer = printer
        self.downloads_dir = Path(downloads_dir)
        self.reports_dir = Path(reports_dir)
        self.report_filename = report_filename
        self._run_lock = threading.Lock()

    @property
    def is_running(self) -> bool:
        return self._run_lock.locked()

    def run_once(self, trigger: str = "manual") -> RunResult:
        if not self._run_lock.acquire(blocking=False):
            LOGGER.warning("Ignored %s trigger because a run is already active", trigger)
            return RunResult(
                RunStatus.ALREADY_RUNNING,
                "SOP Reporter is already running.",
            )
        try:
            return self._run_once_locked(trigger)
        except Exception as exc:
            LOGGER.exception("SOP Reporter %s run failed", trigger)
            return RunResult(
                status=RunStatus.FAILED,
                message=f"Run failed: {exc}",
                error=exc,
            )
        finally:
            self._run_lock.release()

    def _run_once_locked(self, trigger: str) -> RunResult:
        started_at = datetime.now()
        LOGGER.info("Starting SOP Reporter run (trigger=%s)", trigger)
        self.downloads_dir.mkdir(parents=True, exist_ok=True)
        self.reports_dir.mkdir(parents=True, exist_ok=True)

        records = self.email_client.fetch_messages(
            self.downloads_dir,
            handled_message_ids=self.state_store.handled_message_ids(),
        )
        if not records:
            LOGGER.info("No new matching Gmail messages")
            return RunResult(RunStatus.NO_MESSAGES, "No new matching email was found.")

        metadata = self._state_metadata(records)
        empty_records = [record for record in records if not record.attachments]
        attachment_records = [record for record in records if record.attachments]
        if empty_records:
            self.state_store.mark_processed(
                [record.message_id for record in empty_records],
                outcome="no_matching_attachment",
                metadata=metadata,
            )

        if not attachment_records:
            LOGGER.info("New messages had no matching Excel attachments")
            return RunResult(
                RunStatus.NO_DATA,
                "New email was checked, but it had no matching Excel attachment.",
                email_count=len(records),
            )

        attachments = [
            attachment
            for record in attachment_records
            for attachment in record.attachments
        ]
        partitions = self.extractor.extract_partitions(
            attachment.path for attachment in attachments
        )
        attachment_message_ids = list(
            dict.fromkeys(record.message_id for record in attachment_records)
        )
        total_rows = sum(len(partition.data.rows) for partition in partitions)
        if not partitions or total_rows == 0:
            self.state_store.mark_processed(
                attachment_message_ids,
                outcome="no_matching_rows",
                metadata=metadata,
            )
            LOGGER.info("Attachments contained no rows matching extraction filters")
            return RunResult(
                RunStatus.NO_DATA,
                "The attachment was processed, but no rows matched the rules.",
                email_count=len(records),
                attachment_count=len(attachments),
            )

        report_paths: list[Path] = []
        for partition in partitions:
            report_path = self._unique_report_path(started_at, partition.label)
            self.report_builder.build(
                partition.data,
                report_path,
                generated_at=started_at,
                title=self.report_builder.title_for_partition(partition.label),
            )
            report_paths.append(report_path)

        self.state_store.claim_for_print(
            attachment_message_ids,
            report_paths=[str(path) for path in report_paths],
            metadata=metadata,
        )
        completed_print_jobs = 0
        try:
            for report_path in report_paths:
                self.printer.print_workbook(report_path)
                completed_print_jobs += 1
        except Exception as exc:
            # If PrintOut was never invoked, a retry is safe. Once Excel has
            # accepted PrintOut, retaining the claim prevents a possible duplicate
            # after an ambiguous spooler/COM error or abrupt machine shutdown.
            if (
                completed_print_jobs == 0
                and getattr(exc, "print_invoked", False) is False
            ):
                self.state_store.release_claims(attachment_message_ids)
            raise

        outcome = "printed" if self.printer.config.enabled else "generated_not_printed"
        self.state_store.complete_claims(
            attachment_message_ids,
            outcome=outcome,
        )
        status = RunStatus.SUCCESS if self.printer.config.enabled else RunStatus.GENERATED
        action = "printed" if self.printer.config.enabled else "generated"
        LOGGER.info(
            "SOP Reporter run completed: %d email(s), %d attachment(s), "
            "%d report(s), %d row(s)",
            len(records),
            len(attachments),
            len(report_paths),
            total_rows,
        )
        report_word = "report" if len(report_paths) == 1 else "separate reports"
        return RunResult(
            status=status,
            message=(
                f"{len(report_paths)} {report_word} {action} "
                f"with {total_rows} total row(s)."
            ),
            report_path=report_paths[0],
            report_paths=tuple(report_paths),
            email_count=len(records),
            attachment_count=len(attachments),
            row_count=total_rows,
        )

    def _unique_report_path(self, started_at: datetime, partition_label: str) -> Path:
        filename = started_at.strftime(self.report_filename)
        source_path = Path(filename)
        suffix = ""
        if partition_label:
            safe_label = self._filename_token(partition_label)
            raw_suffix = self.report_builder.filename_suffix_for_partition(safe_label)
            suffix = INVALID_FILENAME_CHARACTERS.sub("_", raw_suffix)
            suffix = re.sub(r"\s+", "_", suffix).strip(" .")
            if not suffix:
                suffix = f"_{safe_label}"
        resolved_filename = f"{source_path.stem}{suffix}{source_path.suffix}"
        candidate = self.reports_dir / resolved_filename
        if not candidate.exists():
            return candidate
        counter = 2
        while True:
            candidate = self.reports_dir / (
                f"{Path(resolved_filename).stem}_{counter}"
                f"{Path(resolved_filename).suffix}"
            )
            if not candidate.exists():
                return candidate
            counter += 1

    @staticmethod
    def _filename_token(value: str) -> str:
        token = INVALID_FILENAME_CHARACTERS.sub("_", value)
        token = re.sub(r"\s+", "_", token).strip(" ._")
        return token[:80] or "No_Sub_Status"

    @staticmethod
    def _state_metadata(records: list[EmailRecord]) -> dict[str, dict[str, Any]]:
        return {
            record.message_id: {
                "subject": record.subject,
                "sender": record.sender,
                "sent_date": record.sent_date,
                "attachments": [
                    attachment.original_filename for attachment in record.attachments
                ],
            }
            for record in records
        }
