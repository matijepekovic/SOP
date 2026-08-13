from __future__ import annotations

import logging
from pathlib import Path

from sop_reporter.config import (
    load_app_config,
    load_extraction_config,
    update_email_account,
)
from sop_reporter.credentials import CredentialManager
from sop_reporter.email_client import GmailIMAPClient
from sop_reporter.exceptions import CredentialsCancelledError
from sop_reporter.extractor import ExtractionEngine
from sop_reporter.logging_setup import setup_logging
from sop_reporter.paths import AppPaths
from sop_reporter.pipeline import JobRunner
from sop_reporter.printer import ExcelPrinter
from sop_reporter.report_builder import ReportBuilder
from sop_reporter.scheduler import JobScheduler
from sop_reporter.single_instance import AlreadyRunningError, SingleInstance
from sop_reporter.state_store import ProcessedStateStore
from sop_reporter.tray import TrayApp


LOGGER = logging.getLogger(__name__)


def _show_error(title: str, message: str) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showerror(title, message, parent=root)
        root.destroy()
    except Exception:
        LOGGER.exception("Could not display startup error dialog")


def _build_runner(paths: AppPaths) -> tuple[JobRunner, object]:
    app_config = load_app_config(paths.app_config_path)
    credentials = CredentialManager().ensure_credentials(app_config.email.account)
    if credentials.account != app_config.email.account:
        update_email_account(paths.app_config_path, credentials.account)
        app_config = load_app_config(paths.app_config_path)

    extraction_config = load_extraction_config(paths.extraction_config_path)
    downloads_dir = paths.resolve_runtime_directory(
        app_config.output.downloads_directory
    )
    reports_dir = paths.resolve_runtime_directory(app_config.output.reports_directory)
    downloads_dir.mkdir(parents=True, exist_ok=True)
    reports_dir.mkdir(parents=True, exist_ok=True)

    email_client = GmailIMAPClient(
        app_config.email,
        account=credentials.account,
        app_password=credentials.app_password,
    )
    printer = ExcelPrinter(app_config.printer)
    runner = JobRunner(
        email_client=email_client,
        state_store=ProcessedStateStore(paths.state_path),
        extractor=ExtractionEngine(extraction_config),
        report_builder=ReportBuilder(extraction_config),
        printer=printer,
        downloads_dir=downloads_dir,
        reports_dir=reports_dir,
        report_filename=app_config.output.report_filename,
    )
    return runner, app_config


def main() -> int:
    paths: AppPaths | None = None
    instance: SingleInstance | None = None
    scheduler: JobScheduler | None = None
    log_path: Path | None = None
    try:
        paths = AppPaths.discover()
        paths.ensure_layout()
        log_path = setup_logging(paths.logs_dir)
        instance = SingleInstance.acquire()

        initial_config = load_app_config(paths.app_config_path)
        log_path = setup_logging(
            paths.logs_dir,
            level=initial_config.logging.level,
            max_bytes=initial_config.logging.max_bytes,
            backup_count=initial_config.logging.backup_count,
        )
        LOGGER.info("Starting SOP Reporter from %s", paths.runtime_root)
        runner, app_config = _build_runner(paths)

        scheduler = JobScheduler(
            app_config.schedule,
            callback=lambda: runner.run_once(trigger="scheduled"),
        )
        reports_dir = paths.resolve_runtime_directory(
            app_config.output.reports_directory
        )
        tray = TrayApp(
            job_runner=runner,
            scheduler=scheduler,
            tooltip=app_config.tray.tooltip,
            icon_path=paths.default_icon_path,
            config_path=paths.config_dir,
            logs_dir=paths.logs_dir,
            reports_dir=reports_dir,
        )
        scheduler.start()
        tray.run()
        return 0
    except AlreadyRunningError:
        return 0
    except CredentialsCancelledError:
        if log_path:
            LOGGER.info("First-run Gmail setup was cancelled")
        return 1
    except Exception as exc:
        LOGGER.exception("SOP Reporter could not start")
        detail = f"SOP Reporter could not start:\n\n{exc}"
        if log_path:
            detail += f"\n\nLog: {log_path}"
        _show_error("SOP Reporter", detail)
        return 1
    finally:
        if scheduler is not None:
            scheduler.stop()
        if instance is not None:
            instance.close()


if __name__ == "__main__":
    raise SystemExit(main())
