from __future__ import annotations

import logging
import threading
from collections.abc import Callable

from sop_reporter.config import ScheduleConfig


LOGGER = logging.getLogger(__name__)


class JobScheduler:
    def __init__(self, config: ScheduleConfig, callback: Callable[[], object]) -> None:
        self.config = config
        self._callback = callback
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self._scheduler = None

    def start(self) -> None:
        if not self.config.enabled:
            LOGGER.info("Automatic schedule is disabled")
            return
        if self._thread and self._thread.is_alive():
            return
        try:
            import schedule
        except ImportError as exc:
            raise RuntimeError(
                "The schedule package is not installed; reinstall SOP Reporter"
            ) from exc

        self._scheduler = schedule.Scheduler()
        for day in self.config.days:
            day_job = getattr(self._scheduler.every(), day)
            day_job.at(self.config.time).do(self._dispatch)
        self._stop_event.clear()
        self._thread = threading.Thread(
            target=self._loop,
            name="SOP-Scheduler",
            daemon=True,
        )
        self._thread.start()
        LOGGER.info(
            "Scheduler started for %s at %s",
            ", ".join(self.config.days),
            self.config.time,
        )

    def stop(self) -> None:
        self._stop_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=max(2, self.config.polling_seconds + 1))
        LOGGER.info("Scheduler stopped")

    def _loop(self) -> None:
        assert self._scheduler is not None
        while not self._stop_event.wait(self.config.polling_seconds):
            try:
                self._scheduler.run_pending()
            except Exception:
                LOGGER.exception("Scheduled-job dispatch failed")

    def _dispatch(self) -> None:
        threading.Thread(
            target=self._run_callback,
            name="SOP-Scheduled-Run",
            daemon=True,
        ).start()

    def _run_callback(self) -> None:
        try:
            self._callback()
        except Exception:
            LOGGER.exception("Scheduled SOP Reporter run failed")

