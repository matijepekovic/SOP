from __future__ import annotations

import logging
import sys
import threading
from logging.handlers import RotatingFileHandler
from pathlib import Path


LOG_FILENAME = "sop_reporter.log"


def setup_logging(
    logs_dir: Path,
    level: str = "INFO",
    max_bytes: int = 2_097_152,
    backup_count: int = 5,
) -> Path:
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_path = logs_dir / LOG_FILENAME

    numeric_level = getattr(logging, level.upper(), logging.INFO)
    handler = RotatingFileHandler(
        log_path,
        maxBytes=max_bytes,
        backupCount=backup_count,
        encoding="utf-8",
    )
    handler.setFormatter(
        logging.Formatter(
            "%(asctime)s %(levelname)s [%(threadName)s] %(name)s: %(message)s"
        )
    )

    root = logging.getLogger()
    for existing_handler in root.handlers:
        try:
            existing_handler.close()
        except Exception:
            pass
    root.handlers.clear()
    root.setLevel(numeric_level)
    root.addHandler(handler)
    if sys.stdout is not None:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(handler.formatter)
        root.addHandler(console_handler)

    def handle_exception(exc_type, exc_value, exc_traceback) -> None:
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        logging.getLogger(__name__).critical(
            "Unhandled exception",
            exc_info=(exc_type, exc_value, exc_traceback),
        )

    def handle_thread_exception(args: threading.ExceptHookArgs) -> None:
        logging.getLogger(__name__).critical(
            "Unhandled thread exception in %s",
            args.thread.name if args.thread else "unknown thread",
            exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
        )

    sys.excepthook = handle_exception
    threading.excepthook = handle_thread_exception
    return log_path
