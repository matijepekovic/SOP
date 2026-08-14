from __future__ import annotations

import tempfile
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from sop_reporter.config import load_extraction_config
from sop_reporter.extractor import ExtractionEngine
from sop_reporter.report_builder import ReportBuilder
from tests.fixtures.make_salesforce_sample_xlsx import (
    make_salesforce_sample_workbook,
)


PROJECT_ROOT = Path(__file__).resolve().parent.parent


class ReportBuilderTests(unittest.TestCase):
    def test_report_style_dimensions_formats_and_page_setup(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = make_salesforce_sample_workbook(root / "input.xlsx")
            source_before = source.read_bytes()
            config = load_extraction_config(
                PROJECT_ROOT / "config" / "extraction_rules.default.yaml"
            )
            partition = ExtractionEngine(config).extract_partitions([source])[0]
            destination = root / "report.xlsx"
            builder = ReportBuilder(config)
            builder.build(
                partition.data,
                destination,
                generated_at=datetime(2026, 8, 13, 7, 0),
                title=builder.title_for_partition(partition.label),
            )

            self.assertEqual(source.read_bytes(), source_before)
            workbook = load_workbook(destination, data_only=False)
            worksheet = workbook["Olympia SOP"]
            self.assertEqual(worksheet["A1"].value, "OLYMPIA — Item Notification")
            self.assertIn("08/13/2026 07:00 AM", worksheet["A2"].value)
            self.assertEqual(worksheet["A4"].value, "Product")
            self.assertEqual(worksheet["F4"].value, "Amount")
            self.assertEqual(worksheet["A4"].fill.fgColor.rgb, "001F4E78")
            self.assertTrue(worksheet["A4"].font.bold)
            self.assertEqual(worksheet["F5"].number_format, "$#,##0.00")
            self.assertEqual(worksheet["F5"].alignment.horizontal, "right")
            self.assertEqual(worksheet["H5"].alignment.horizontal, "center")
            self.assertEqual(worksheet.column_dimensions["A"].width, 12)
            self.assertEqual(worksheet.freeze_panes, "A5")
            self.assertEqual(worksheet.page_setup.paperSize, 3)
            self.assertEqual(worksheet.page_setup.orientation, "landscape")
            self.assertEqual(worksheet.page_setup.fitToWidth, 1)
            self.assertEqual(worksheet.auto_filter.ref, "A4:K7")
            workbook.close()


if __name__ == "__main__":
    unittest.main()
