from __future__ import annotations

import logging
import sys
from pathlib import Path

from sop_reporter import __version__
from sop_reporter.config import (
    load_app_config,
    load_report_definitions,
    update_email_account,
)
from sop_reporter.credentials import CredentialManager
from sop_reporter.email_client import GmailIMAPClient
from sop_reporter.exceptions import ConfigurationError, CredentialsCancelledError
from sop_reporter.extractor import ExtractionEngine
from sop_reporter.logging_setup import setup_logging
from sop_reporter.paths import AppPaths
from sop_reporter.pipeline import JobRunner, ReportJob
from sop_reporter.printer import ExcelPrinter
from sop_reporter.report_builder import ReportBuilder
from sop_reporter.scheduler import JobScheduler
from sop_reporter.single_instance import AlreadyRunningError, SingleInstance
from sop_reporter.state_store import ProcessedStateStore
from sop_reporter.tray import TrayApp
from sop_reporter.updater import RELAUNCH_FLAG, Updater


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

    definitions = load_report_definitions(paths.config_dir)
    enabled = [d for d in definitions if d.match.enabled]
    LOGGER.info(
        "Loaded %d report definition(s); enabled: %s",
        len(definitions),
        ", ".join(d.name for d in enabled) or "(none)",
    )
    if not enabled:
        raise ConfigurationError(
            "No report is enabled. Set match.enabled to true in at least one "
            f"file under {paths.config_dir / 'rules'}."
        )
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
    jobs = [
        ReportJob(
            definition=definition,
            extractor=ExtractionEngine(definition.extraction),
            report_builder=ReportBuilder(definition.extraction),
        )
        for definition in definitions
        if definition.match.enabled
    ]
    runner = JobRunner(
        email_client=email_client,
        state_store=ProcessedStateStore(paths.state_path),
        jobs=jobs,
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
        # A build started by a self-update waits for the outgoing process to
        # release the single-instance mutex before claiming it.
        relaunched = RELAUNCH_FLAG in sys.argv[1:]
        instance = SingleInstance.acquire(wait_seconds=45.0 if relaunched else 0.0)

        initial_config = load_app_config(paths.app_config_path)
        log_path = setup_logging(
            paths.logs_dir,
            level=initial_config.logging.level,
            max_bytes=initial_config.logging.max_bytes,
            backup_count=initial_config.logging.backup_count,
        )
        LOGGER.info(
            "Starting SOP Reporter %s from %s", __version__, paths.runtime_root
        )
        runner, app_config = _build_runner(paths)

        updater: Updater | None = None
        if app_config.update.enabled:
            updater = Updater(
                repository=app_config.update.repository,
                current_version=__version__,
                include_prereleases=app_config.update.include_prereleases,
            )
            # Clear the build the previous update renamed aside; it is
            # unlocked now that the older process has exited.
            try:
                updater.cleanup_previous_builds()
            except Exception:
                LOGGER.debug("Could not clean up superseded builds", exc_info=True)

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
            extraction_config_path=paths.extraction_config_path,
            logs_dir=paths.logs_dir,
            reports_dir=reports_dir,
            updater=updater,
            current_version=__version__,
            check_updates_on_startup=(
                updater is not None and app_config.update.check_on_startup
            ),
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
