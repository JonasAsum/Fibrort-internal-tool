"""Tells which layout a sheet uses. The file name settles the two templates that
carry a name in it (Customer Visits in Detail, Weekly Sales Planning); otherwise
the header position does."""
import os

from .headers import find_visit_header, normalize
from .sheet import Sheet

LAYOUTS = ("mock", "weekly2026", "lucas", "detail")


def label_cell(sheet: Sheet, text: str, max_row: int = 20):
    """(row, col) of the first cell whose normalized text equals `text`."""
    for r in range(1, min(max_row, sheet.max_row) + 1):
        for c in range(1, sheet.max_col + 1):
            if normalize(sheet.value(r, c)) == text:
                return r, c
    return None


def _mock_header(sheet: Sheet) -> bool:
    """visits.reader's layout: the Date header sits in column A."""
    return any(normalize(sheet.value(r, 1)) == "date" for r in range(1, min(20, sheet.max_row) + 1))


def detect(sheet: Sheet, path: str) -> str | None:
    name = os.path.basename(path).lower()
    if "customer visits in detail" in name:
        return "detail" if label_cell(sheet, "customer name") else None
    if "weekly sales planning" in name:
        return "lucas" if find_visit_header(sheet) else None
    if _mock_header(sheet):
        return "mock"
    found = find_visit_header(sheet)
    if found is None:
        return None
    return "lucas" if "follow_up" in found[1] else "weekly2026"
