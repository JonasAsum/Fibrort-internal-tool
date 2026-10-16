"""Independent oracle for the to-do / sales-topic workbooks (folder 07). Built from the cell ledger only; no
importer code is imported.

For every sheet it states: whether it is a task list (a header row naming topic / to-do columns) or a value list
(the Dropdown sheet), the header row, which column is which field, the title / owner label cells above the header,
and for every row below the header its cells with the field each one belongs to. A row with no topic and no
to-do text is a `no_text` row (the stray status / date cells of a blank template).
"""
import os
import re
from dataclasses import dataclass, field

from tests.quality import fam3_common as oc

LABELS = {  # normalised header text -> field
    "no.": "seq", "date": "date", "categorie": "category", "category": "category", "topics": "topic", "topic": "topic",
    "to do / next steps": "todo", "responsible": "owner", "due date": "due", "status": "status", "prio": "prio",
    "remarks": "remarks", "documents or websites (hyperlink)": "link", "date provided": "provided",
}
PLACEHOLDER_DATE = re.compile(r"^[\dx\-/.]+$", re.I)  # 2026-x-xx: a date nobody filled in
TITLE = re.compile(r"^topic\s*[-–:]\s*(.*)$", re.I)
OWNER_LABEL = re.compile(r"^(?:project responsible|name)\s*[:：]\s*(.*)$", re.I)
PLACEHOLDER_NAME = re.compile(r"^(?:x+|_+|\.+|-+|n/?a|tbd|tbc)$", re.I)
REPS = ("Jason Zhao", "Devin Guo", "Mark Wang", "Sean Wang", "Ethan Lv", "Jason Yi", "Andrew Zhang", "Lucas Guo")


@dataclass
class Row:
    file: str
    sheet: str
    row: int
    cells: dict  # field -> ledger Cell (non-empty only)

    @property
    def has_text(self):
        return "topic" in self.cells or "todo" in self.cells


@dataclass
class Sheet:
    file: str
    sheet: str
    kind: str  # todo | topics | value_list | other
    header_row: int = 0
    columns: dict = field(default_factory=dict)  # field -> col
    title: str = ""
    owner_label: str = ""  # text after "Project Responsible:" / "Name:" ('' when empty)
    owner_is_placeholder: bool = False
    above: list = field(default_factory=list)
    header_cells: list = field(default_factory=list)
    rows: list = field(default_factory=list)  # Row
    other_cells: list = field(default_factory=list)  # cells of a sheet that is not a task list


@dataclass
class FileFacts:
    path: str
    file: str
    sheets: dict = field(default_factory=dict)
    cells: list = field(default_factory=list)
    is_it_list: bool = False


def _header(rows):
    for r in sorted(rows):
        if r > 30:
            break
        cols = {}
        for c in rows[r].values():
            f = LABELS.get(oc.norm(c.text).lower())
            if f and f not in cols:
                cols[f] = c.col
        if len(cols) >= 3 and ("topic" in cols or "todo" in cols):
            return r, cols
    return None


def analyse(path) -> FileFacts:
    cells = oc.workbook_cells(path)
    facts = FileFacts(path, os.path.basename(path), cells=cells)
    for sheet, rows in oc.rows_by_sheet(cells).items():
        found = _header(rows)
        if not found:
            texts = {oc.norm(c.text).lower() for r in rows.values() for c in r.values()}
            value_list = bool(texts) and all(t in {"status", "prio", "open", "wip", "done", "1", "2", "3"} for t in texts)
            sh = Sheet(facts.file, sheet, "value_list" if value_list else "other")
            sh.other_cells = [c for r in rows.values() for c in r.values()]
            facts.sheets[sheet] = sh
            continue
        hrow, cols = found
        layout = "todo" if {"todo", "owner", "status", "prio"} & cols.keys() else "topics"
        sh = Sheet(facts.file, sheet, layout, hrow, cols)
        by_col = {c: f for f, c in cols.items()}
        for r in sorted(rows):
            if r < hrow:
                sh.above += list(rows[r].values())
                for c in rows[r].values():
                    t = oc.norm(c.text)
                    if m := TITLE.match(t):
                        sh.title = sh.title or m.group(1).strip()
                    elif m := OWNER_LABEL.match(t):
                        sh.owner_label = m.group(1).strip()
                continue
            if r == hrow:
                sh.header_cells = list(rows[r].values())
                continue
            rcells = {by_col[c.col]: c for c in rows[r].values() if c.col in by_col}
            unmapped = [c for c in rows[r].values() if c.col not in by_col]
            if unmapped:
                sh.other_cells += unmapped
            sh.rows.append(Row(facts.file, sheet, r, rcells))
        sh.owner_is_placeholder = (not sh.owner_label) or bool(PLACEHOLDER_NAME.match(sh.owner_label))
        if PLACEHOLDER_NAME.match(sh.title or ""):
            sh.title = ""
        facts.sheets[sheet] = sh
    facts.is_it_list = bool(re.search(r"(?<![A-Za-z])IT(?![A-Za-z])", os.path.splitext(facts.file)[0])) or any(
        re.search(r"jet-?it", row.cells["owner"].text, re.I) for s in facts.sheets.values() for row in s.rows if "owner" in row.cells)
    return facts


def all_facts(sample_dir=oc.SAMPLE_DIR):
    return [analyse(p) for p in oc.folder_files("07", sample_dir)]
