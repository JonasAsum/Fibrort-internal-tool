"""Parses one visit-list workbook into raw visit rows and one week record per
sheet.

Layout (docs/README.md in mock_visits, confirmed against the generator and the
real sheets it mimics): row 1 names the owner and the week, rows 2-7 hold the
weekday "Key Plan" (column B) beside the week's "Summary" lines (column J), a
header row says `Date` in its first cell, and every visit afterwards is a block
of merged rows A-J with one note per row in column K (colour = action/waiting)
and a deadline in column L. The year is not on the sheet; it is the last
`_YYYY` token in the workbook's file name.
"""
import os
import re
from dataclasses import dataclass, field, replace
from datetime import date, timedelta

import openpyxl

MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]

DATE_PATTERNS = [
    re.compile(r"^(\d{1,2})-([A-Za-z]{3,})$"),  # 13-May
    re.compile(r"^(\d{1,2})(?:st|nd|rd|th),\s*([A-Za-z]{3,})$", re.I),  # 26th,May
    re.compile(r"^([A-Za-z]{3,})\s+(\d{1,2})$"),  # May 13
    re.compile(r"^(\d{1,2})/(\d{1,2})$"),  # 13/05 (day/month)
]

SHEET_CW = re.compile(r"(?:CW|Week|W)\s*(\d+)", re.I)
WORKBOOK_YEAR = re.compile(r"_(\d{4})\.xlsx$", re.I)

GREEN_FILL = "1E8070"
YELLOW_FILL = "FFFF00"

# Rows 3-7 label their day "Monday", "Tue", ... ; the value is the offset from
# that week's Monday.
WEEKDAY_OFFSETS = {"mon": 0, "tue": 1, "wed": 2, "thu": 3, "fri": 4, "sat": 5, "sun": 6}


def fold_name(name: str) -> str:
    return re.sub(r"[\W_]+", "", (name or "").lower())


def assign_ordinals(visits: list) -> None:
    """Numbers visits that share owner, date and customer 0, 1, ... in the order given (one sheet)."""
    seen: dict = {}
    for v in visits:
        n = seen.get(v.content_key, 0)
        seen[v.content_key] = n + 1
        v.key_ordinal = n


@dataclass
class Note:
    text: str
    highlight: str | None  # "green", "yellow", or None
    deadline: str
    excel_row: int = 0


@dataclass
class Visit:
    workbook: str
    sheet: str
    owner: str
    year: int
    cw: int
    excel_row_start: int
    excel_row_end: int
    date_raw: str
    date_iso: str | None
    city: str
    customer_raw: str
    potential: str
    customer_type_raw: str
    contact_raw: str
    phone_raw: str
    office_phone: str
    email: str
    purpose: str
    notes: list[Note] = field(default_factory=list)
    # Position among this sheet's visits on the same date, from 0. The sheet
    # states the order a rep went in but never a clock time.
    order_in_day: int = 0
    # Columns the mock layout does not have (WeChat, industry, competitor,
    # status ...), kept so the activity's raw row still holds every cell.
    extra: dict = field(default_factory=dict)
    # Which of the visits with the same owner, date and customer this is, in sheet
    # order from 0 (set by assign_ordinals once a whole sheet is read).
    key_ordinal: int = 0

    @property
    def visit_key(self) -> str:
        """Where the visit sits in the file; identifies it inside one run, never across files."""
        return f"{self.workbook}:{self.sheet}:{self.excel_row_start}"

    @property
    def content_key(self) -> tuple:
        return (fold_name(self.owner), self.date_iso, fold_name(self.customer_raw))

    @property
    def source_id(self) -> str:
        """The same visit in a later export of the file gets the same id, whatever the file is called."""
        owner, day, customer = self.content_key
        return f"visit:{owner}:{day}:{customer}:{self.key_ordinal}"

    @property
    def kind(self) -> str:
        t = (self.customer_type_raw or "").strip().lower()
        if "exhibit" in t:
            return "exhibition"
        if t.startswith("dealer"):
            return "dealer"
        # Neither "internal" nor blank-for-an-internal-row is written to the
        # type column (the generator leaves it empty); the only reliable tell
        # is that none of the per-contact columns carry anything either.
        if not t and not any((self.potential, self.contact_raw, self.phone_raw, self.email)):
            return "internal"
        return "customer"

    @property
    def customer_type(self) -> str:
        """The type column folded to new / existing / dealer, or "" when the
        cell says something else (an exhibition, an internal row)."""
        t = (self.customer_type_raw or "").strip().lower()
        for prefix, folded in (("dealer", "dealer"), ("exist", "existing"), ("new", "new")):
            if t.startswith(prefix):
                return folded
        return ""

    @property
    def status(self) -> str:
        if self.kind in ("internal", "exhibition"):
            return self.kind
        if any("cancel" in n.text.lower() for n in self.notes):
            return "cancelled"
        return "done"


@dataclass
class DayPlan:
    weekday: str  # as the sheet spells it: "Monday", "Tue", ...
    date_iso: str
    plan: str


@dataclass
class Week:
    workbook: str
    sheet: str
    owner: str
    year: int
    cw: int
    monday_iso: str
    days: list[DayPlan]
    summary: list[str]

    @property
    def week_key(self) -> str:
        """Where the week sits in the file; identifies it inside one run only."""
        return f"{self.workbook}:{self.sheet}:week"

    @property
    def source_id(self) -> str:
        return f"week:{fold_name(self.owner)}:{self.year}-W{self.cw}"


def merge_weeks(weeks: list) -> list:
    """One week per (owner, ISO year, CW): the plan and summary lines of later copies are
    added to the first one's."""
    merged: dict = {}
    for w in weeks:
        first = merged.get(w.source_id)
        if first is None:
            merged[w.source_id] = replace(w, days=[replace(d) for d in w.days], summary=list(w.summary))
            continue
        by_date = {d.date_iso: d for d in first.days}
        for d in w.days:
            mine = by_date.get(d.date_iso)
            if mine is None:
                first.days.append(replace(d))
                by_date[d.date_iso] = first.days[-1]
            elif d.plan and d.plan not in mine.plan:
                mine.plan = f"{mine.plan}; {d.plan}" if mine.plan else d.plan
        first.days.sort(key=lambda d: d.date_iso)
        first.summary += [line for line in w.summary if line not in first.summary]
    return list(merged.values())


@dataclass
class Parsed:
    visits: list[Visit] = field(default_factory=list)
    weeks: list[Week] = field(default_factory=list)


def parse_date(raw: str, year: int) -> str | None:
    raw = (raw or "").strip()
    for pattern in DATE_PATTERNS:
        m = pattern.match(raw)
        if not m:
            continue
        a, b = m.groups()
        day, month_name = (a, b) if a.isdigit() and not b.isdigit() else (b, a)
        if a.isdigit() and b.isdigit():  # 13/05 = day/month
            day, month = a, b
        else:
            try:
                month = MONTHS.index(month_name[:3].lower()) + 1
            except ValueError:
                return None
        try:
            return date(year, int(month), int(day)).isoformat()
        except ValueError:
            return None
    return None


def cw_monday(year: int, cw: int) -> date:
    """Monday of ISO calendar week `cw` - the "CW" every sheet is named by."""
    return date.fromisocalendar(year, cw, 1)


def _fill_highlight(cell) -> str | None:
    color = cell.fill.fgColor.rgb if cell.fill and cell.fill.fgColor else None
    if not color or cell.fill.patternType != "solid":
        return None
    suffix = str(color)[-6:].upper()
    if suffix == GREEN_FILL:
        return "green"
    if suffix == YELLOW_FILL:
        return "yellow"
    return None


def _header_row(ws) -> int:
    for r in range(1, 20):
        if str(ws.cell(row=r, column=1).value or "").strip().lower() == "date":
            return r
    raise ValueError(f"{ws.title}: no 'Date' header row found in the first 20 rows")


def _text(value) -> str:
    return str(value).strip() if value is not None else ""


def parse_week(ws, workbook_name: str, year: int, cw: int, header_row: int) -> Week:
    monday = cw_monday(year, cw)
    days, summary = [], []
    for r in range(1, header_row):
        label = _text(ws.cell(row=r, column=1).value)
        offset = WEEKDAY_OFFSETS.get(label[:3].lower())
        if offset is not None:
            days.append(DayPlan(weekday=label, date_iso=(monday + timedelta(days=offset)).isoformat(),
                                plan=_text(ws.cell(row=r, column=2).value)))
        line = _text(ws.cell(row=r, column=10).value)
        if r > 1 and line:  # row 1 holds the "Summary" heading itself
            summary.append(line)
    return Week(workbook=workbook_name, sheet=ws.title, owner=_text(ws.cell(row=1, column=2).value),
                year=year, cw=cw, monday_iso=monday.isoformat(), days=days, summary=summary)


def parse_sheet(ws, workbook_name: str, year: int) -> tuple[list[Visit], Week]:
    owner = _text(ws.cell(row=1, column=2).value)
    cw_match = SHEET_CW.search(ws.title)
    if not cw_match:
        raise ValueError(f"{workbook_name}/{ws.title}: sheet name carries no CW/Week number")
    cw = int(cw_match.group(1))

    header_row = _header_row(ws)
    starts = [r for r in range(header_row + 1, ws.max_row + 1) if ws.cell(row=r, column=1).value is not None]

    visits = []
    seen_per_day: dict[str, int] = {}
    for i, start in enumerate(starts):
        end = (starts[i + 1] - 1) if i + 1 < len(starts) else ws.max_row
        cells = [ws.cell(row=start, column=c).value for c in range(1, 11)]
        date_raw, city, customer_raw, potential, ctype_raw, contact_raw, phone_raw, office, email, purpose = (
            _text(v) for v in cells
        )
        notes = []
        for r in range(start, end + 1):
            k, l = ws.cell(row=r, column=11), ws.cell(row=r, column=12)
            if k.value is None or not str(k.value).strip():
                continue
            notes.append(Note(text=str(k.value).strip(), highlight=_fill_highlight(k),
                              deadline=_text(l.value), excel_row=r))
        date_iso = parse_date(date_raw, year)
        day = date_iso or date_raw
        order = seen_per_day.get(day, 0)
        seen_per_day[day] = order + 1
        visits.append(Visit(
            workbook=workbook_name, sheet=ws.title, owner=owner, year=year, cw=cw,
            excel_row_start=start, excel_row_end=end,
            date_raw=date_raw, date_iso=date_iso, city=city,
            customer_raw=customer_raw, potential=potential, customer_type_raw=ctype_raw,
            contact_raw=contact_raw, phone_raw=phone_raw, office_phone=office, email=email,
            purpose=purpose, notes=notes, order_in_day=order,
        ))
    assign_ordinals(visits)
    return visits, parse_week(ws, workbook_name, year, cw, header_row)


def parse_workbook(path: str) -> Parsed:
    name = os.path.basename(path)
    year_match = WORKBOOK_YEAR.search(name)
    year = int(year_match.group(1)) if year_match else date.today().year
    wb = openpyxl.load_workbook(path)
    out = Parsed()
    for sheet_name in wb.sheetnames:
        visits, week = parse_sheet(wb[sheet_name], name, year)
        out.visits.extend(visits)
        out.weeks.append(week)
    return out


def parse_paths(paths: list[str]) -> Parsed:
    out = Parsed()
    for p in paths:
        one = parse_workbook(p)
        out.visits.extend(one.visits)
        out.weeks.extend(one.weeks)
    out.weeks = merge_weeks(out.weeks)
    return out
