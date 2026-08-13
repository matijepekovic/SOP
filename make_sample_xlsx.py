from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import Workbook


HEADERS = [
    "Sales Rep",
    "Branch",
    "Job Number",
    "Status",
    "Amount",
    "Sale Date",
    "Notes",
    "Score",
    "Active",
]

ROWS = [
    ["Alice", "Tacoma", "J001", "Approved", 1000, date(2026, 8, 10), "Priority roof", 5, "yes"],
    ["Bob", "Seattle", "J002", "Rejected", 900, date(2026, 8, 10), "Budget", 3, "no"],
    ["Alice", "Tacoma", "J003", "Approved", "$1,500.00", date(2026, 8, 11), "Standard", 2, "yes"],
    ["Carol", "Olympia", "J004", "approved", 2500, date(2026, 8, 12), "Referral", 4, "yes"],
    ["Dave", "Tacoma", "J005", "Pending", 100, date(2026, 8, 12), None, 1, "no"],
    ["Eve", "Seattle", "J006", "Approved", "(300.00)", date(2026, 8, 13), "Callback", 0, "yes"],
]


def make_sample_workbook(path: Path) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    worksheet.append(HEADERS)
    for row in ROWS:
        worksheet.append(row)
    worksheet.append([None] * len(HEADERS))
    workbook.save(path)
    workbook.close()
    return path


if __name__ == "__main__":
    make_sample_workbook(Path(__file__).with_name("sample_input.xlsx"))

