from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from sop_reporter.exceptions import StateStoreError
from sop_reporter.state_store import ProcessedStateStore


class ProcessedStateStoreTests(unittest.TestCase):
    def test_add_save_and_reload_round_trip(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "state.json"
            store = ProcessedStateStore(path)
            store.mark_processed(
                ["<one@example>"],
                outcome="no_matching_rows",
                metadata={"<one@example>": {"subject": "Daily"}},
            )
            reloaded = ProcessedStateStore(path)
            self.assertTrue(reloaded.is_handled("<one@example>"))
            record = reloaded.snapshot()["messages"]["<one@example>"]
            self.assertEqual(record["subject"], "Daily")
            self.assertEqual(record["outcome"], "no_matching_rows")

    def test_print_claim_is_handled_and_can_be_completed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "state.json"
            store = ProcessedStateStore(path)
            store.claim_for_print(["<two@example>"], report_path="report.xlsx")
            self.assertTrue(store.is_handled("<two@example>"))
            store.complete_claims(["<two@example>"])
            record = ProcessedStateStore(path).snapshot()["messages"]["<two@example>"]
            self.assertEqual(record["status"], "processed")
            self.assertEqual(record["outcome"], "printed")

    def test_partial_temporary_write_does_not_corrupt_primary_state(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "state.json"
            store = ProcessedStateStore(path)
            store.mark_processed(["<safe@example>"], outcome="printed")
            with patch("sop_reporter.state_store.os.replace", side_effect=OSError("disk failure")):
                with self.assertRaises(StateStoreError):
                    store.mark_processed(["<partial@example>"], outcome="printed")
            self.assertFalse(store.is_handled("<partial@example>"))
            reloaded = ProcessedStateStore(path)
            self.assertTrue(reloaded.is_handled("<safe@example>"))
            self.assertFalse(reloaded.is_handled("<partial@example>"))


if __name__ == "__main__":
    unittest.main()
