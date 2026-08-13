from __future__ import annotations

import json
import os
import threading
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Mapping

from sop_reporter.exceptions import StateStoreError


STATE_VERSION = 1
HANDLED_STATUSES = {"printing", "processed"}


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


class ProcessedStateStore:
    def __init__(self, path: Path) -> None:
        self.path = path
        self._lock = threading.RLock()
        self._state = self._load()

    def _load(self) -> dict[str, Any]:
        if not self.path.exists():
            return {"version": STATE_VERSION, "messages": {}}
        try:
            with self.path.open("r", encoding="utf-8") as handle:
                state = json.load(handle)
        except (OSError, json.JSONDecodeError) as exc:
            raise StateStoreError(f"Cannot read state file {self.path}: {exc}") from exc
        if not isinstance(state, dict) or state.get("version") != STATE_VERSION:
            raise StateStoreError(f"Unsupported or invalid state file: {self.path}")
        if not isinstance(state.get("messages"), dict):
            raise StateStoreError(f"State file has no valid messages map: {self.path}")
        return state

    def snapshot(self) -> dict[str, Any]:
        with self._lock:
            return deepcopy(self._state)

    def handled_message_ids(self) -> set[str]:
        with self._lock:
            return {
                message_id
                for message_id, record in self._state["messages"].items()
                if isinstance(record, dict) and record.get("status") in HANDLED_STATUSES
            }

    def is_handled(self, message_id: str) -> bool:
        return message_id in self.handled_message_ids()

    def mark_processed(
        self,
        message_ids: Iterable[str],
        *,
        outcome: str,
        report_path: str | None = None,
        metadata: Mapping[str, Mapping[str, Any]] | None = None,
    ) -> None:
        with self._lock:
            previous = deepcopy(self._state)
            try:
                timestamp = _utc_now()
                for message_id in message_ids:
                    record = {
                        "status": "processed",
                        "outcome": outcome,
                        "processed_at": timestamp,
                    }
                    if report_path:
                        record["report_path"] = report_path
                    if metadata and message_id in metadata:
                        record.update(dict(metadata[message_id]))
                    self._state["messages"][message_id] = record
                self._write_locked()
            except Exception:
                self._state = previous
                raise

    def claim_for_print(
        self,
        message_ids: Iterable[str],
        *,
        report_path: str | None = None,
        report_paths: Iterable[str] | None = None,
        metadata: Mapping[str, Mapping[str, Any]] | None = None,
    ) -> None:
        ids = list(dict.fromkeys(message_ids))
        paths = list(dict.fromkeys(report_paths or ()))
        if report_path and report_path not in paths:
            paths.insert(0, report_path)
        if not paths:
            raise StateStoreError("At least one report path is required for a print claim")
        with self._lock:
            already_handled = [message_id for message_id in ids if self.is_handled(message_id)]
            if already_handled:
                raise StateStoreError(
                    "Refusing to print already-handled message(s): "
                    + ", ".join(already_handled)
                )
            previous = deepcopy(self._state)
            try:
                timestamp = _utc_now()
                for message_id in ids:
                    record = {
                        "status": "printing",
                        "outcome": "print_claimed",
                        "claimed_at": timestamp,
                        "report_path": paths[0],
                        "report_paths": paths,
                    }
                    if metadata and message_id in metadata:
                        record.update(dict(metadata[message_id]))
                    self._state["messages"][message_id] = record
                self._write_locked()
            except Exception:
                self._state = previous
                raise

    def complete_claims(
        self, message_ids: Iterable[str], *, outcome: str = "printed"
    ) -> None:
        with self._lock:
            previous = deepcopy(self._state)
            try:
                timestamp = _utc_now()
                for message_id in message_ids:
                    record = self._state["messages"].get(message_id)
                    if not isinstance(record, dict) or record.get("status") != "printing":
                        raise StateStoreError(
                            f"Message does not have an active print claim: {message_id}"
                        )
                    record["status"] = "processed"
                    record["outcome"] = outcome
                    record["processed_at"] = timestamp
                self._write_locked()
            except Exception:
                self._state = previous
                raise

    def release_claims(self, message_ids: Iterable[str]) -> None:
        with self._lock:
            previous = deepcopy(self._state)
            changed = False
            for message_id in message_ids:
                record = self._state["messages"].get(message_id)
                if isinstance(record, dict) and record.get("status") == "printing":
                    del self._state["messages"][message_id]
                    changed = True
            if changed:
                try:
                    self._write_locked()
                except Exception:
                    self._state = previous
                    raise

    def _write_locked(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_name(
            f"{self.path.name}.{os.getpid()}.{threading.get_ident()}.tmp"
        )
        try:
            with temporary.open("w", encoding="utf-8", newline="\n") as handle:
                json.dump(self._state, handle, indent=2, sort_keys=True)
                handle.write("\n")
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary, self.path)
        except OSError as exc:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass
            raise StateStoreError(f"Cannot persist state file {self.path}: {exc}") from exc
