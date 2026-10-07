"""Turns what a visit list says about time into instants the CRM can sort and
remind on.

The sheet carries a date per visit and nothing finer, so a visit's time of day
is a placeholder that keeps the sheet's own order within the day; the activity
body says the time was not recorded. Deadlines are free text ("Reply to
customer in 3 days.", "Follow up in CW 26.") resolved against the visit's
date; a phrase none of the rules below recognise stays text and gets no due
date rather than a guessed one.
"""
import calendar
import re
from datetime import date, datetime, time, timedelta
from zoneinfo import ZoneInfo

from .reader import MONTHS, Visit, cw_monday

FIRST_SLOT = time(9, 0)
SLOT_HOURS = 2  # 09:00, 11:00, 13:00 ... in sheet order

_IN_DAYS = re.compile(r"\bin\s+(\d{1,2})\s+(?:working\s+)?days?\b", re.I)
_CW = re.compile(r"\b(?:CW|week)\s*(\d{1,2})\b", re.I)
_MONTH = re.compile(r"\bin\s+(" + "|".join(m for m in MONTHS) + r")[a-z]*\b", re.I)
_FOLLOW_UP_BY = re.compile(r"^\s*([A-Z][a-z]+(?:\s?[A-Z][a-z]+)?)\s+will\s+follow\s+up", re.I)


def visit_instant(visit: Visit, tz: ZoneInfo) -> datetime | None:
    if not visit.date_iso:
        return None
    day = date.fromisoformat(visit.date_iso)
    slot = datetime.combine(day, FIRST_SLOT, tzinfo=tz) + timedelta(hours=SLOT_HOURS * visit.order_in_day)
    return slot


def day_instant(day_iso: str, tz: ZoneInfo, at: time = FIRST_SLOT) -> datetime:
    return datetime.combine(date.fromisoformat(day_iso), at, tzinfo=tz)


def resolve_due(phrase: str, visit_day: date) -> date | None:
    """The date a deadline phrase points at, counted from the visit's day."""
    text = phrase.strip()
    if not text:
        return None
    lower = text.lower()
    if m := _IN_DAYS.search(text):
        return visit_day + timedelta(days=int(m.group(1)))
    if "tomorrow" in lower:
        return visit_day + timedelta(days=1)
    if "next week" in lower:
        return visit_day + timedelta(days=7 - visit_day.weekday())
    if "this week" in lower:
        return visit_day + timedelta(days=max(0, 4 - visit_day.weekday()))  # that week's Friday
    if "this month" in lower:
        return visit_day.replace(day=calendar.monthrange(visit_day.year, visit_day.month)[1])
    if m := _CW.search(text):
        cw = int(m.group(1))
        year = visit_day.year
        # "CW 2" written in a December sheet means next January.
        if cw < visit_day.isocalendar().week - 26:
            year += 1
        try:
            return cw_monday(year, cw)
        except ValueError:
            return None
    if m := _MONTH.search(text):
        month = MONTHS.index(m.group(1)[:3].lower()) + 1
        year = visit_day.year + (1 if month < visit_day.month else 0)
        return date(year, month, 1)
    return None


def follow_up_owner(phrase: str) -> str | None:
    """"LiuYang will follow up this customer." -> "LiuYang"; the caller folds
    it against the rep names it knows."""
    m = _FOLLOW_UP_BY.match(phrase)
    return m.group(1) if m else None


def fold_person(name: str) -> str:
    return re.sub(r"[^a-z]", "", name.lower())
