"""Finds a visit sheet's header row and maps its labels to the fields the readers
use. A label matches a synonym when its normalized text equals it or starts with
it. The first column carrying a field wins."""
import re

from .sheet import Sheet

SCAN_ROWS = 30
SCAN_COLS = 40

SYNONYMS = [
    ("date", ("date", "visiting date")),
    ("cw", ("cw",)),
    ("sales", ("sales",)),
    ("plan", ("plan",)),
    ("city", ("location/city", "location", "city")),
    ("customer", ("customer",)),
    ("potential", ("potential",)),
    ("contact", ("contact name", "name")),
    ("phone", ("phone no.", "phone")),
    ("office", ("office",)),
    ("wechat", ("wechat",)),
    ("email", ("email",)),
    ("industry", ("industries", "industry")),
    ("purpose", ("purpose_topics visit", "purpose for visiting", "activity", "purpose")),
    ("competitor", ("competiton products", "competitor")),
    ("communication", ("communication content",)),
    ("next_step", ("next step / action", "next step")),
    ("follow_up", ("follow up",)),
    ("due", ("due date", "deadline")),
    ("status", ("status",)),
]


def normalize(label) -> str:
    """Lowercase, newlines and runs of spaces to one space, trailing colon dropped."""
    text = "" if label is None else str(label)
    text = re.sub(r"\s+", " ", text.replace("\n", " ")).strip().lower()
    return text.rstrip("：:").strip()


def field_for(label) -> str | None:
    t = normalize(label)
    if not t:
        return None
    if "new, exist" in t:
        return "customer_type"
    for field, synonyms in SYNONYMS:
        if any(t == s or t.startswith(s) for s in synonyms):
            return field
    return None


def find_header(sheet: Sheet):
    """(row, {field: col}) for the first row mapping at least four fields,
    one of them the date; None when the sheet has no such row."""
    for r in range(1, min(SCAN_ROWS, sheet.max_row) + 1):
        cols = {}
        for c in range(1, min(SCAN_COLS, sheet.max_col) + 1):
            field = field_for(sheet.value(r, c))
            if field and field not in cols:
                cols[field] = c
        if "date" in cols and len(cols) >= 4:
            return r, cols
    return None
