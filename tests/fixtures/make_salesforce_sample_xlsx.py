from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import Workbook


# Mirrors the real GM SOP export: ten banner rows above the header, an empty
# leading column, arrows appended to whichever columns the report is sorted
# by, and a colon in the contact column's name.
BANNER_ROWS = [
    [],
    [None, "GM SOP"],
    [None, "As of 2026-08-13 14:00:09 Pacific Standard Time/PST • Sorted by Close Date (Descending)"],
    [],
    [],
    [None, "Filtered By"],
    [None, "Show: All opportunities"],
    [None, "Status not equal to Cancel Outside Recission,Canceled,Completed"],
    [None, "Service Project equals False"],
    [],
]
HEADER_ROW_NUMBER = len(BANNER_ROWS) + 1

HEADERS = [
    "Market  \u2191",
    "Sub Status  \u2193",
    "Product  \u2191",
    "Job Number",
    "Contact: Full Name",
    "Opportunity Name",
    "Next Step-WO",
    "Amount",
    "Days in Current Sub Status",
    "Close Date",
    "Assigned Service Resource: Name",
]

ROWS = [
    [
        "Olympia",
        "Item Notification",
        "Roofing",
        "273140",
        "Ivan Sanchez",
        "8409 Spruce Street Southwest Lakewood Washington 98498",
        "08/13 FU MB 08/10 Need PO for whole roof, need plywood diagram MB",
        29_598.91,
        1,
        date(2026, 8, 17),
        "Leona Miles",
    ],
    [
        None,
        None,
        "Roofing",
        "274803",
        "Robert Gillis",
        "13027 Kapowsin Highlands Dr E Graham Washington 98338",
        "08/13 FU. CB. 08/12 FU. CB. 08/11 FU CB. Sent email on approval.",
        15_475.00,
        5,
        date(2026, 8, 22),
        "Jose Perez",
    ],
    [
        None,
        None,
        "Siding",
        "274000",
        "Tom Randles",
        "9949 Beth Court Southeast Yelm Washington 98597",
        "08/12 Need HO to confirm new color or siding type and hover re-built.",
        71_417.91,
        4,
        date(2026, 8, 26),
        "Dale Porter",
        None,
    ],
    [None, "Subtotal", None, None, None, None, None, 116_491.82, None, None, None, None],
    [None, "Install Issue", None, None, None, None, None, None, None, None, None, None],
    [
        None,
        None,
        "Windows",
        "275101",
        "Erin Walsh",
        "4207 Cooper Point Road Olympia Washington 98502",
        "Confirm replacement sash delivery and installer availability.",
        18_750.00,
        9,
        date(2026, 8, 20),
        "Leona Miles",
    ],
    [
        None,
        None,
        "Roofing",
        "275112",
        "Paul Young",
        "1700 East Bay Drive Northeast Olympia Washington 98506",
        "Verify flashing correction photos before closeout.",
        24_200.00,
        3,
        date(2026, 8, 21),
        "Jose Perez",
        None,
    ],
    [None, "Subtotal", None, None, None, None, None, 42_950.00, None, None, None, None],
    [None, "On Hold", None, None, None, None, None, None, None, None, None, None],
    [
        None,
        None,
        "Siding",
        "275150",
        "Maria Lee",
        "510 Capitol Way South Olympia Washington 98501",
        "Waiting for HOA color authorization.",
        32_800.00,
        12,
        date(2026, 9, 1),
        "Dale Porter",
    ],
    [
        "Tacoma",
        "Item Notification",
        "Roofing",
        "900001",
        "Not Olympia One",
        "Tacoma Washington",
        "This row must be filtered out.",
        99_000.00,
        20,
        date(2026, 9, 2),
        "Other Rep",
        None,
    ],
    [
        None,
        None,
        "Roofing",
        "900002",
        "Not Olympia Two",
        "Tacoma Washington",
        "Fill-down must keep this row in Tacoma and filter it out.",
        88_000.00,
        19,
        date(2026, 9, 3),
        "Other Rep",
        None,
    ],
    ["Grand Total", None, None, None, None, None, None, 379_241.82, None, None, None, None],
]


def make_salesforce_sample_workbook(path: Path) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "GM SOP"
    for banner in BANNER_ROWS:
        worksheet.append(banner)
    # A leading empty column, exactly as the export produces.
    worksheet.append([None] + HEADERS)
    for row in ROWS:
        worksheet.append([None] + list(row))
    workbook.save(path)
    workbook.close()
    return path


if __name__ == "__main__":
    make_salesforce_sample_workbook(
        Path(__file__).with_name("salesforce_olympia_sample.xlsx")
    )
