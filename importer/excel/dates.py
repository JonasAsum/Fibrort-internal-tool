"""Turns the date and due cells of the sheets into ISO dates.

parse_date resolves a visit date against the sheet's week (ref) when the text is
ambiguous (05/06 is 5 June or 6 May); parse_due reads a deadline, where a bare
decimal such as 5.25 means 25 May. Anything that is not a date comes back as the
original phrase so visits.when can still try it later.
"""
import re
from datetime import date, datetime, timedelta

from openpyxl.utils import column_index_from_string
from openpyxl.utils.datetime import from_excel

from .split import is_placeholder

MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
ORD = r"(?:st|nd|rd|th)?"
ISO_LIKE = re.compile(r"^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})$")
NUM_SLASH = re.compile(r"^(\d{1,2})/(\d{1,2})(?:/(\d{4}))?$")
DAY_MONTH = re.compile(r"^(\d{1,2})" + ORD + r"[\s,./-]*([A-Za-z]{3,})\.?(?:[\s,]*(\d{4}))?$")
MONTH_DAY = re.compile(r"^([A-Za-z]{3,})\.?[\s,]*(\d{1,2})" + ORD + r"(?:[\s,]*(\d{4}))?$")
FORMULA = re.compile(r"^=\s*\$?([A-Za-z]{1,3})\$?(\d+)\s*\+\s*(\d+)$")
DECIMAL = re.compile(r"^(\d{1,2})\.(\d{1,2})$")
RANGE = re.compile(r"\s*(?:--|–)\s*")
NEAR_DAYS = 10
OLD_YEAR = 1990  # Excel stores a "d/mmm" typed without a year as 1900
ROLL_DAYS = 60


def _month(name: str) -> int | None:
    key = name.lower()[:3]
    return MONTHS.index(key) + 1 if key in MONTHS else None


def _mk(y: int, m: int, d: int) -> date | None:
    try:
        return date(y, m, d)
    except (ValueError, TypeError):
        return None


def _year(ref, year):
    return ref.year if ref else year


def _stamp(value, ref, year):
    """A date cell. A 1900 date means the year was never typed: keep month and day,
    take the year from the ref week, else the workbook year."""
    day = value.date() if isinstance(value, datetime) else value
    if day.year >= OLD_YEAR:
        return day.isoformat(), ""
    yr = _year(ref, year)
    if yr is None:
        return None, "no_year"
    fixed = _mk(yr, day.month, day.day)
    return (fixed.isoformat(), "") if fixed else (None, "unparsed")


def _parse_text(text, ref, year):
    if m := ISO_LIKE.match(text):
        day = _mk(*(int(g) for g in m.groups()))
        return (day.isoformat(), "") if day else (None, "unparsed")
    if m := NUM_SLASH.match(text):
        a, b, yy = m.groups()
        yr = int(yy) if yy else _year(ref, year)
        if yr is None:
            return None, "no_year"
        cands = []
        for day, mon in ((int(a), int(b)), (int(b), int(a))):  # day/month first
            d = _mk(yr, mon, day)
            if d and d not in cands:
                cands.append(d)
        if not cands:
            return None, "unparsed"
        if len(cands) == 1:
            return cands[0].isoformat(), ""
        if ref is None:
            return cands[0].isoformat(), "ambiguous"
        best = min(cands, key=lambda d: abs((d - ref).days))
        flag = "" if abs((best - ref).days) <= NEAR_DAYS else "ambiguous"
        return best.isoformat(), flag
    for pattern, groups in ((DAY_MONTH, ("d", "m", "y")), (MONTH_DAY, ("m", "d", "y"))):
        if m := pattern.match(text):
            parts = dict(zip(groups, m.groups()))
            mon = _month(parts["m"])
            yr = int(parts["y"]) if parts["y"] else _year(ref, year)
            if mon is None:
                return None, "unparsed"
            if yr is None:
                return None, "no_year"
            day = _mk(yr, mon, int(parts["d"]))
            return (day.isoformat(), "") if day else (None, "unparsed")
    return None, "unparsed"


def parse_date(value, ref: date | None, year: int | None = None, lookup=None):
    """(iso or None, flag). Flags: "" clean, "ambiguous", "range" (first day taken),
    "formula" (=D9+N resolved), "empty", "unparsed", "no_year".
    lookup(row, col) gives the value of another cell, for "=D9+1" text."""
    if isinstance(value, (datetime, date)):
        return _stamp(value, ref, year)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if value < 1:
            return None, "empty"
        try:
            stamp = from_excel(value)
        except (OverflowError, ValueError, TypeError):  # 1e30 and the like are not a date
            return None, "unparsed"
        return _stamp(stamp, ref, year)
    text = "" if value is None else re.sub(r"\s+", " ", str(value)).strip()
    if text.startswith("="):
        return _formula(text, ref, year, lookup)
    if text in ("", "--") or is_placeholder(text):
        return None, "empty"
    if "--" in text or "–" in text:
        first, second = RANGE.split(text, maxsplit=1)
        if not first:
            return None, "unparsed"
        if not re.search(r"[A-Za-z]{3,}", first) and (mon := re.search(r"[A-Za-z]{3,}", second)):
            first = f"{first} {mon.group(0)}"
        if not re.search(r"\d{4}", first) and (yr := re.search(r"\d{4}", second)):
            first = f"{first} {yr.group(0)}"
        iso, _ = _parse_text(first, ref, year)
        return (iso, "range") if iso else (None, "unparsed")
    return _parse_text(text, ref, year)


def _formula(text, ref, year, lookup):
    m = FORMULA.match(text)
    if not m or lookup is None:
        return None, "formula"
    base = lookup(int(m.group(2)), column_index_from_string(m.group(1).upper()))
    base_iso, _ = parse_date(base, ref, year)  # one level only, no chains
    if base_iso is None:
        return None, "formula"
    day = date.fromisoformat(base_iso) + timedelta(days=int(m.group(3)))
    return day.isoformat(), "formula"


def parse_due(value, visit_day: date):
    """(iso or None, phrase, flag). Flag "decimal" marks a month.day number; the
    phrase is the original text when no date was read."""
    if isinstance(value, (datetime, date)):
        iso, _ = _stamp(value, visit_day, None)
        return (iso, "", "") if iso else (None, "", "")
    if isinstance(value, float) and value % 1:
        return _decimal(f"{value}", visit_day, str(value))
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        iso, flag = parse_date(value, visit_day)
        return (iso, "", flag) if iso else (None, str(value), "phrase")
    text = "" if value is None else str(value).strip()
    if not text or is_placeholder(text):
        return None, "", ""
    if DECIMAL.match(text):
        return _decimal(text, visit_day, text)
    iso, flag = parse_date(text, visit_day)
    if iso:
        return iso, "", flag
    return None, text, "phrase"


def _decimal(text, visit_day, phrase):
    m = DECIMAL.match(text)
    if not m:
        return None, phrase, "phrase"
    month, dd = int(m.group(1)), m.group(2)
    if not 1 <= month <= 12:
        return None, phrase, "phrase"
    days = [int(dd)]
    if len(dd) == 1 and int(dd) * 10 <= 31:  # 5.1 may have been 5.10
        days.append(int(dd) * 10)
    cands = []
    for day in days:
        d = _mk(visit_day.year, month, day)
        if d is None:
            continue
        if (visit_day - d).days > ROLL_DAYS:
            d = _mk(visit_day.year + 1, month, day)
        cands.append(d)
    if not cands:
        return None, phrase, "phrase"
    ahead = [d for d in cands if d >= visit_day]
    pick = min(ahead, key=lambda d: (d - visit_day).days) if ahead else max(cands)
    return pick.isoformat(), "", "decimal"
