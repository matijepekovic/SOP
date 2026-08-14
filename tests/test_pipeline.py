from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

from sop_reporter.email_client import DownloadedAttachment, EmailRecord
from sop_reporter.exceptions import PrintError
from sop_reporter.extractor import ExtractedData, ExtractedPartition
from sop_reporter.config import MatchConfig, ReportDefinition
from sop_reporter.pipeline import JobRunner, ReportJob, RunStatus
from sop_reporter.state_store import ProcessedStateStore


class FakeEmailClient:
    def __init__(self, record: EmailRecord) -> None:
        self.record = record

    def fetch_messages(self, _download_dir, handled_message_ids=()):
        return [] if self.record.message_id in handled_message_ids else [self.record]


class FakeExtractor:
    def extract_partitions(self, paths):
        source_files = tuple(paths)
        data = ExtractedData(
            headers=("Representative", "Total"),
            rows=({"Representative": "Alice", "Total": 2500},),
            number_formats={"Total": "$#,##0.00"},
            widths={"Representative": 20, "Total": 14},
            source_files=source_files,
        )
        return (ExtractedPartition("", data),)


class MultiFakeExtractor(FakeExtractor):
    def extract_partitions(self, paths):
        source_files = tuple(paths)
        first = ExtractedData(
            headers=("Job Number",),
            rows=({"Job Number": "J001"}, {"Job Number": "J002"}),
            number_formats={},
            widths={"Job Number": 14},
            source_files=source_files,
        )
        second = ExtractedData(
            headers=("Job Number",),
            rows=({"Job Number": "J003"},),
            number_formats={},
            widths={"Job Number": 14},
            source_files=source_files,
        )
        return (
            ExtractedPartition("Item Notification", first),
            ExtractedPartition("Install Issue", second),
        )


class FakeBuilder:
    def build(self, _data, destination, generated_at=None, title=None):
        destination.write_bytes(b"report")
        return destination

    def title_for_partition(self, label):
        return f"OLYMPIA — {label}" if label else "SOP REPORT"

    def filename_suffix_for_partition(self, safe_label):
        return f"_{safe_label}" if safe_label else ""


class FakePrinter:
    def __init__(self) -> None:
        self.config = SimpleNamespace(enabled=True)
        self.printed: list[Path] = []

    def print_workbook(self, path):
        self.printed.append(path)


class FailingPrinter(FakePrinter):
    def __init__(self, print_invoked: bool) -> None:
        super().__init__()
        self.print_invoked = print_invoked

    def print_workbook(self, path):
        self.printed.append(path)
        raise PrintError("simulated print failure", print_invoked=self.print_invoked)


class FailingOnSecondPrinter(FakePrinter):
    def print_workbook(self, path):
        self.printed.append(path)
        if len(self.printed) == 2:
            raise PrintError("second print failed before PrintOut", print_invoked=False)



def make_job(
    name: str,
    extractor,
    *,
    patterns: tuple[str, ...] = ("*.xlsx", "*.xlsm"),
    enabled: bool = True,
    report_filename: str | None = None,
    builder=None,
) -> ReportJob:
    """Build a ReportJob around fake collaborators for pipeline tests."""
    return ReportJob(
        definition=ReportDefinition(
            name=name,
            match=MatchConfig(filename_patterns=patterns, enabled=enabled),
            extraction=None,  # pipeline never reads this; the engine is faked
            report_filename=report_filename,
        ),
        extractor=extractor,
        report_builder=builder or FakeBuilder(),
    )


class PipelineTests(unittest.TestCase):
    def _make_runner(self, root: Path, printer, extractor=None):
        attachment_path = root / "input.xlsx"
        attachment_path.write_bytes(b"input")
        record = EmailRecord(
            sequence_id="1",
            message_id="<one@example>",
            subject="Daily SOP",
            sender="reports@example.com",
            sent_date="Thu, 13 Aug 2026 07:00:00 -0700",
            attachments=(
                DownloadedAttachment(
                    "input.xlsx",
                    attachment_path,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    5,
                ),
            ),
        )
        store = ProcessedStateStore(root / "state.json")
        runner = JobRunner(
            email_client=FakeEmailClient(record),
            state_store=store,
            jobs=[make_job("Test Report", extractor or FakeExtractor())],
            printer=printer,
            downloads_dir=root / "downloads",
            reports_dir=root / "reports",
            report_filename="Report_%Y%m%d_%H%M%S.xlsx",
        )
        return runner, store

    def test_successful_run_prints_and_second_run_deduplicates(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            printer = FakePrinter()
            runner, store = self._make_runner(root, printer)
            first = runner.run_once()
            second = runner.run_once()

            self.assertEqual(first.status, RunStatus.SUCCESS)
            self.assertTrue(first.report_path.is_file())
            self.assertEqual(len(printer.printed), 1)
            self.assertTrue(store.is_handled("<one@example>"))
            self.assertEqual(second.status, RunStatus.NO_MESSAGES)
            self.assertEqual(len(printer.printed), 1)

    def test_each_sub_status_builds_and_prints_separately(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            printer = FakePrinter()
            runner, store = self._make_runner(
                root,
                printer,
                extractor=MultiFakeExtractor(),
            )
            result = runner.run_once()

            self.assertEqual(result.status, RunStatus.SUCCESS)
            self.assertEqual(result.row_count, 3)
            self.assertEqual(len(result.report_paths), 2)
            self.assertEqual(len(printer.printed), 2)
            self.assertIn("Item_Notification", result.report_paths[0].name)
            self.assertIn("Install_Issue", result.report_paths[1].name)
            state_record = store.snapshot()["messages"]["<one@example>"]
            self.assertEqual(len(state_record["report_paths"]), 2)

    def test_failure_before_printout_releases_claim_for_retry(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            printer = FailingPrinter(print_invoked=False)
            runner, store = self._make_runner(Path(directory), printer)
            with self.assertLogs("sop_reporter.pipeline", level="ERROR"):
                result = runner.run_once()
            self.assertEqual(result.status, RunStatus.FAILED)
            self.assertFalse(store.is_handled("<one@example>"))

    def test_ambiguous_failure_after_printout_retains_claim(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            printer = FailingPrinter(print_invoked=True)
            runner, store = self._make_runner(Path(directory), printer)
            with self.assertLogs("sop_reporter.pipeline", level="ERROR"):
                result = runner.run_once()
            self.assertEqual(result.status, RunStatus.FAILED)
            self.assertTrue(store.is_handled("<one@example>"))

    def test_failure_before_second_print_keeps_claim_for_first_print(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            printer = FailingOnSecondPrinter()
            runner, store = self._make_runner(
                Path(directory),
                printer,
                extractor=MultiFakeExtractor(),
            )
            with self.assertLogs("sop_reporter.pipeline", level="ERROR"):
                result = runner.run_once()
            self.assertEqual(result.status, RunStatus.FAILED)
            self.assertEqual(len(printer.printed), 2)
            self.assertTrue(store.is_handled("<one@example>"))


if __name__ == "__main__":
    unittest.main()
