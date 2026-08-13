from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from sop_reporter.config import load_app_config, load_extraction_config
from sop_reporter.exceptions import ConfigurationError


PROJECT_ROOT = Path(__file__).resolve().parent.parent


class ConfigurationTests(unittest.TestCase):
    def test_default_configs_load(self) -> None:
        app = load_app_config(PROJECT_ROOT / "config" / "app_config.default.yaml")
        extraction = load_extraction_config(
            PROJECT_ROOT / "config" / "extraction_rules.default.yaml"
        )
        self.assertEqual(app.schedule.time, "07:00")
        self.assertEqual(app.schedule.days[:2], ("monday", "tuesday"))
        self.assertEqual(app.printer.paper_size, "tabloid")
        self.assertIsNone(extraction.input.sheet_name)
        self.assertTrue(extraction.split.enabled)
        self.assertEqual(extraction.split.by, "Sub Status")
        self.assertEqual(extraction.filters.rules[0].value, "Olympia")
        self.assertFalse(extraction.grouping.enabled)

    def test_invalid_schedule_time_has_clear_error(self) -> None:
        invalid = """
version: 1
email: {}
schedule:
  time: seven
output: {}
printer: {}
logging: {}
tray: {}
"""
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "app.yaml"
            path.write_text(invalid, encoding="utf-8")
            with self.assertRaisesRegex(ConfigurationError, "24-hour HH:MM"):
                load_app_config(path)


if __name__ == "__main__":
    unittest.main()
