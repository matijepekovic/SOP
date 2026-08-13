from __future__ import annotations

import tempfile
import unittest
from dataclasses import replace
from pathlib import Path

from sop_reporter.config import (
    AggregationRule,
    ColumnRule,
    FiltersConfig,
    FilterRule,
    GroupingConfig,
    InputConfig,
    SortRule,
    SplitConfig,
    load_extraction_config,
)
from sop_reporter.extractor import ExtractionEngine
from tests.fixtures.make_sample_xlsx import make_sample_workbook
from tests.fixtures.make_salesforce_sample_xlsx import (
    make_salesforce_sample_workbook,
)


PROJECT_ROOT = Path(__file__).resolve().parent.parent


class ExtractionEngineTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.workbook_path = make_sample_workbook(
            Path(self.temporary.name) / "sample.xlsx"
        )
        default = load_extraction_config(
            PROJECT_ROOT / "config" / "extraction_rules.default.yaml"
        )
        self.columns = (
            ColumnRule("Sales Rep", "Representative", "text"),
            ColumnRule("Branch", "Branch", "text"),
            ColumnRule("Job Number", "Job", "text"),
            ColumnRule("Status", "Status", "text"),
            ColumnRule("Amount", "Amount", "currency", number_format="$#,##0.00"),
            ColumnRule("Sale Date", "Date", "date", number_format="mm/dd/yyyy"),
            ColumnRule("Notes", "Notes", "text", required=False),
            ColumnRule("Score", "Score", "integer"),
            ColumnRule("Active", "Active", "boolean"),
        )
        self.base_config = replace(
            default,
            input=InputConfig("Data", 1, 2, False, True),
            columns=self.columns,
            filters=FiltersConfig(),
            split=SplitConfig(),
            grouping=GroupingConfig(),
            sort=(),
        )

    def _jobs_for(self, rule: FilterRule) -> list[str]:
        config = replace(
            self.base_config,
            filters=FiltersConfig(mode="all", rules=(rule,)),
        )
        data = ExtractionEngine(config).extract_file(self.workbook_path)
        return [row["Job"] for row in data.rows]

    def test_column_rename_and_type_conversion(self) -> None:
        data = ExtractionEngine(self.base_config).extract_file(self.workbook_path)
        self.assertEqual(data.headers[0], "Representative")
        self.assertEqual(data.rows[2]["Amount"], 1500.0)
        self.assertEqual(data.rows[5]["Amount"], -300.0)
        self.assertTrue(data.rows[0]["Active"])
        self.assertFalse(data.rows[1]["Active"])

    def test_all_supported_filter_operators(self) -> None:
        cases = [
            (FilterRule("Status", "equals", value="APPROVED"), ["J001", "J003", "J004", "J006"]),
            (FilterRule("Status", "not_equals", value="approved"), ["J002", "J005"]),
            (FilterRule("Notes", "contains", value="priority"), ["J001"]),
            (FilterRule("Notes", "not_contains", value="budget"), ["J001", "J003", "J004", "J005", "J006"]),
            (FilterRule("Branch", "in", values=("Tacoma", "Olympia")), ["J001", "J003", "J004", "J005"]),
            (FilterRule("Branch", "not_in", values=("Tacoma", "Olympia")), ["J002", "J006"]),
            (FilterRule("Score", "range", minimum=2, maximum=4), ["J002", "J003", "J004"]),
            (FilterRule("Amount", "gt", value=1000), ["J003", "J004"]),
            (FilterRule("Amount", "gte", value=1000), ["J001", "J003", "J004"]),
            (FilterRule("Score", "lt", value=2), ["J005", "J006"]),
            (FilterRule("Score", "lte", value=1), ["J005", "J006"]),
            (FilterRule("Notes", "is_blank"), ["J005"]),
            (FilterRule("Notes", "not_blank"), ["J001", "J002", "J003", "J004", "J006"]),
        ]
        for rule, expected in cases:
            with self.subTest(operator=rule.operator):
                self.assertEqual(self._jobs_for(rule), expected)

    def test_case_sensitive_filter(self) -> None:
        self.assertEqual(
            self._jobs_for(
                FilterRule(
                    "Status",
                    "equals",
                    value="approved",
                    case_sensitive=True,
                )
            ),
            ["J004"],
        )

    def test_grouping_aggregates_across_complete_input(self) -> None:
        config = replace(
            self.base_config,
            filters=FiltersConfig(
                rules=(FilterRule("Status", "equals", value="Approved"),)
            ),
            grouping=GroupingConfig(
                enabled=True,
                by=("Representative", "Branch"),
                aggregations=(
                    AggregationRule("Job", "Approved Jobs", "count"),
                    AggregationRule("Amount", "Total Sales", "sum"),
                    AggregationRule("Amount", "Average Sale", "avg"),
                    AggregationRule("Date", "First Sale", "min"),
                    AggregationRule("Date", "Last Sale", "max"),
                ),
            ),
            sort=(
                SortRule("Total Sales", "desc"),
                SortRule("Representative", "asc"),
            ),
        )
        data = ExtractionEngine(config).extract_file(self.workbook_path)
        self.assertEqual(
            data.headers,
            (
                "Representative",
                "Branch",
                "Approved Jobs",
                "Total Sales",
                "Average Sale",
                "First Sale",
                "Last Sale",
            ),
        )
        alice = next(row for row in data.rows if row["Representative"] == "Alice")
        self.assertEqual(alice["Approved Jobs"], 2)
        self.assertEqual(alice["Total Sales"], 2500.0)
        self.assertEqual(alice["Average Sale"], 1250.0)
        self.assertEqual(str(alice["First Sale"]), "2026-08-10")
        self.assertEqual(str(alice["Last Sale"]), "2026-08-11")
        self.assertEqual(data.rows[0]["Representative"], "Alice")
        self.assertEqual(data.rows[1]["Representative"], "Carol")

    def test_olympia_filter_fill_down_and_sub_status_partitions(self) -> None:
        source = make_salesforce_sample_workbook(
            Path(self.temporary.name) / "salesforce.xlsx"
        )
        config = load_extraction_config(
            PROJECT_ROOT / "config" / "extraction_rules.default.yaml"
        )
        partitions = ExtractionEngine(config).extract_partitions([source])

        self.assertEqual(
            [partition.label for partition in partitions],
            ["Item Notification", "Install Issue", "On Hold"],
        )
        self.assertEqual(
            [len(partition.data.rows) for partition in partitions],
            [3, 2, 1],
        )
        self.assertNotIn("Market", partitions[0].data.headers)
        self.assertNotIn("Sub Status", partitions[0].data.headers)
        self.assertEqual(
            partitions[0].data.rows[0]["Job Number"],
            "274803",
        )
        all_jobs = {
            row["Job Number"]
            for partition in partitions
            for row in partition.data.rows
        }
        self.assertFalse({"900001", "900002"} & all_jobs)


if __name__ == "__main__":
    unittest.main()
