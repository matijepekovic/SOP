from __future__ import annotations

import logging
import os
import threading
from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw

from sop_reporter.pipeline import JobRunner, RunStatus
from sop_reporter.scheduler import JobScheduler


LOGGER = logging.getLogger(__name__)


class TrayApp:
    def __init__(
        self,
        *,
        job_runner: JobRunner,
        scheduler: JobScheduler,
        tooltip: str,
        icon_path: Path,
        config_path: Path,
        extraction_config_path: Path,
        logs_dir: Path,
        reports_dir: Path,
        updater=None,
        current_version: str = "",
        check_updates_on_startup: bool = False,
        change_credentials=None,
    ) -> None:
        self.job_runner = job_runner
        self.scheduler = scheduler
        self.tooltip = tooltip
        self.icon_path = icon_path
        self.config_path = config_path
        self.extraction_config_path = extraction_config_path
        self.logs_dir = logs_dir
        self.reports_dir = reports_dir
        self.updater = updater
        self.current_version = current_version
        self.check_updates_on_startup = check_updates_on_startup
        self.change_credentials = change_credentials
        self._icon: Any | None = None

    def run(self) -> None:
        try:
            import pystray
        except ImportError as exc:
            raise RuntimeError(
                "The pystray package is not installed; reinstall SOP Reporter"
            ) from exc

        menu = pystray.Menu(
            pystray.MenuItem("Open SOP Reporter", self._show_control_window, default=True),
            pystray.MenuItem("Run Now", self._run_now),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Open Reports", self._open_reports),
            pystray.MenuItem("Open Settings", self._open_settings),
            pystray.MenuItem("Open Logs", self._open_logs),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Change Gmail Sign-in", self._change_gmail),
            pystray.MenuItem("Check for Updates", self._check_updates),
            pystray.MenuItem("Exit", self._exit),
        )
        self._icon = pystray.Icon(
            "SOPReporter",
            self._load_icon(),
            self.tooltip,
            menu,
        )
        LOGGER.info("Tray application started")
        self._icon.run_detached()

        from sop_reporter.gui.control_window import ControlWindow

        self._control_window = ControlWindow(
            job_runner=self.job_runner,
            run_now=self._run_now,
            exit_app=self._exit,
            app_config_path=self.config_path / "app_config.yaml",
            extraction_config_path=self.extraction_config_path,
            reports_dir=self.reports_dir,
            logs_dir=self.logs_dir,
            updater=self.updater,
            current_version=self.current_version,
            check_updates_on_startup=self.check_updates_on_startup,
            change_credentials=self.change_credentials,
        )
        self._control_window.run()

    def _show_control_window(self, _icon=None, _item=None) -> None:
        window = getattr(self, "_control_window", None)
        if window is not None:
            window.show()

    def _change_gmail(self, _icon=None, _item=None) -> None:
        window = getattr(self, "_control_window", None)
        if window is None:
            return
        window.show()
        window.change_gmail_signin()

    def _check_updates(self, _icon=None, _item=None) -> None:
        """Surface the update flow, which lives in the control window."""
        window = getattr(self, "_control_window", None)
        if window is None:
            return
        if self.updater is None:
            self._notify(
                "Updates are disabled in the configuration file.",
                "SOP Reporter",
            )
            return
        window.show()
        window.check_for_updates()

    def _load_icon(self) -> Image.Image:
        try:
            if self.icon_path.is_file():
                return Image.open(self.icon_path).convert("RGBA")
        except Exception:
            LOGGER.exception("Could not load tray icon; using generated fallback")
        image = Image.new("RGBA", (64, 64), "#172033")
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle((5, 5, 59, 59), radius=10, fill="#1F4E78")
        draw.text((15, 19), "SOP", fill="white")
        return image

    def _run_now(self, _icon=None, _item=None) -> None:
        threading.Thread(
            target=self._run_now_worker,
            name="SOP-Manual-Run",
            daemon=True,
        ).start()

    def _run_now_worker(self) -> None:
        window = getattr(self, "_control_window", None)
        if window is not None:
            window.set_running()
        result = self.job_runner.run_once(trigger="manual")
        if result.status == RunStatus.FAILED:
            title = "SOP Reporter failed"
        elif result.status == RunStatus.SUCCESS:
            title = "SOP Reporter printed"
        else:
            title = "SOP Reporter"
        self._notify(result.message, title)
        if window is not None:
            window.set_result(result)

    def _notify(self, message: str, title: str) -> None:
        try:
            if self._icon is not None:
                self._icon.notify(message, title)
        except Exception:
            LOGGER.debug("Tray notification failed", exc_info=True)

    def _open_reports(self, _icon=None, _item=None) -> None:
        self.reports_dir.mkdir(parents=True, exist_ok=True)
        self._open_path(self.reports_dir)

    def _open_settings(self, _icon=None, _item=None) -> None:
        self._open_path(self.config_path)

    def _open_logs(self, _icon=None, _item=None) -> None:
        self.logs_dir.mkdir(parents=True, exist_ok=True)
        self._open_path(self.logs_dir)

    @staticmethod
    def _open_path(path: Path) -> None:
        try:
            if os.name != "nt":
                raise OSError("Shell opening is supported only on Windows")
            os.startfile(str(path))  # type: ignore[attr-defined]
        except OSError:
            LOGGER.exception("Could not open %s", path)

    def _exit(self, icon=None, _item=None) -> None:
        if self.job_runner.is_running:
            self._notify(
                "Wait for the current fetch/report/print run to finish, then choose Exit again.",
                "SOP Reporter is busy",
            )
            return
        LOGGER.info("Exit requested from tray")
        self.scheduler.stop()
        if self.job_runner.is_running:
            self.scheduler.start()
            self._notify(
                "A scheduled run just started. Wait for it to finish, then choose Exit again.",
                "SOP Reporter is busy",
            )
            return
        active_icon = icon or self._icon
        if active_icon is not None:
            active_icon.stop()
        window = getattr(self, "_control_window", None)
        if window is not None:
            window.close()
