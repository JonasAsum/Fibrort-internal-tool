"""Independent oracle for the AR / financial-status workbooks (folder 09). Built from the cell ledger and
openpyxl (for the sheet visibility only); no importer code is imported.

What it states, from the raw cells:

  * for every sheet of every workbook: an invoice-level AR list, or not; and if not, why (hidden, bank cash flow,
    accounts payable, fixed cost, finance summary, customer down payment, aging ledger),
  * for every AR list: the day it was counted from (the `Today` cell, = TODAY() as last saved), the totals the
    sheet states itself (Total Account Receivables, Total Amount Overdue), every row, and the rows that are not
    invoices (placeholders `#N/A` / `No Invoice`, zero amounts),
  * the OVERDUE RULE, read from the file's own columns: `Due days` = Due Date - Today, so an invoice is overdue
    when `Due days` < 0 (due today = 0 = not overdue). The same rule applied to the due date and the Today cell
    is checked row by row against the `Due days` column the file computed,
  * per-customer open and overdue sums as Decimal, grouped by the Chinese customer name written in the row.
"""
import datetime as dt
import os
import re
from collections import defaultdict
from dataclasses import dataclass, field
from decimal import Decimal

import openpyxl

from tests.quality import fam3_common as oc

LABELS = {  # normalised header text -> role
    "customer name(en)": "en", "customer name": "cn", "sales invoice no.": "invoice_no", "invoice date": "invoice_date",
    "due date": "due_date", "payment term in days": "term", "due days": "due_days", "days since invoice": "age",
    "remark": "remark", "币种": "currency_cn", "original amt": "original", "currency": "currency_code",
}
AMOUNT_LABEL = "invoice value local currency"  # startswith, then "（RMB）"
SHEET_KINDS = (  # own patterns; first match on the sheet name wins
    ("down_payment", re.compile(r"advance|pre-?paid|downpayment|down payment", re.I)),
    ("payable", re.compile(r"^AP\b|payable|^AP-", re.I)),
    ("bank", re.compile(r"\bcash\b|\bEBA\b|\bBOC\b|\bABC\b", re.I)),
    ("fixed", re.compile(r"fix(ed)? cost", re.I)),
    ("summary", re.compile(r"overview|mapping|offset", re.I)),
    ("aging", re.compile(r"\bAR\b.*\d{8}|pure RT AR", re.I)),
)
DATE = re.compile(r"^\d{4}-\d{2}-\d{2}")


@dataclass
class ArRow:
    row: int
    en: str
    cn: str
    invoice_no: str
    invoice_date: str
    due_date: str
    amount: Decimal
    due_days: int | None
    remark: str
    coords: dict
    merged_from: list = field(default_factory=list)  # roles whose value comes from a merged range's top-left cell


@dataclass
class ArSheet:
    file: str
    sheet: str
    header_row: int
    columns: dict  # role -> col
    today: str  # YYYY-MM-DD
    claimed_open: Decimal | None
    claimed_overdue: Decimal | None
    rows: list = field(default_factory=list)
    dropped: list = field(default_factory=list)  # (row, reason)
    side_cells: list = field(default_factory=list)  # ledger Cells beside the list (columns after the list's own)
    above_header: list = field(default_factory=list)
    pivot: dict = field(default_factory=dict)  # customer -> Decimal, the file's own per-customer pivot beside the list

    def overdue(self, r: ArRow) -> bool:
        """The file's own rule: `Due days` < 0 (Due Date - Today). A row whose Due days cell is empty (it sits under
        a merged invoice) takes the same difference from its due date; a row with neither is not overdue."""
        if r.due_days is not None:
            return r.due_days < 0
        return bool(r.due_date) and r.due_date < self.today

    def customers(self):
        """{cn name (en when no cn): {open, overdue, invoices, overdue_invoices, oldest_due}}"""
        out = defaultdict(lambda: {"open": Decimal(0), "overdue": Decimal(0), "invoices": 0, "overdue_invoices": 0,
                                   "oldest_due": ""})
        for r in self.rows:
            c = out[r.cn or r.en]
            c["open"] += r.amount
            c["invoices"] += 1
            if self.overdue(r):
                c["overdue"] += r.amount
                c["overdue_invoices"] += 1
            if r.due_date and (not c["oldest_due"] or r.due_date < c["oldest_due"]):
                c["oldest_due"] = r.due_date
        return dict(out)

    @property
    def open_total(self):
        return sum((r.amount for r in self.rows), Decimal(0))

    @property
    def overdue_total(self):
        return sum((r.amount for r in self.rows if self.overdue(r)), Decimal(0))


@dataclass
class FileFacts:
    path: str
    file: str
    display: str
    sheets: dict  # sheet -> state ("visible" / "hidden")
    ar: dict = field(default_factory=dict)  # sheet -> ArSheet
    other: dict = field(default_factory=dict)  # sheet -> kind ("hidden", "bank", ...)
    cells: list = field(default_factory=list)
    name_date: str = ""


def _find_header(rows):
    for r in sorted(rows):
        if r > 25:
            break
        roles = {}
        for c in rows[r].values():
            t = oc.norm(c.text).lower()
            if t in LABELS:
                roles.setdefault(LABELS[t], c.col)
            elif t.startswith(AMOUNT_LABEL):
                roles.setdefault("amount", c.col)
        if {"invoice_no", "due_date", "amount"} <= set(roles) and ({"en", "cn"} & set(roles)):
            return r, roles
    return None


def _num_below(rows, hrow, col_hint_label):
    """The number under a label like 'Total Account Receivables' (same column, next row)."""
    for r in sorted(rows):
        if r >= hrow:
            break
        for c in rows[r].values():
            if col_hint_label in oc.norm(c.text).lower():
                below = rows.get(r + 1, {}).get(c.col)
                if below is not None and oc.dec(below.text) is not None:
                    return oc.dec(below.text)
    return None


def _read_ar(file, sheet, rows, found, origin):
    """`origin` maps (row, col) inside a merged range to the (row, col) of its top-left cell: Excel shows the
    top-left value in every cell of the range (an invoice that spans two amount rows has one number)."""
    hrow, roles = found
    today = ""
    for r in sorted(rows):
        if r >= hrow:
            break
        if any("today" in oc.norm(c.text).lower() or "今天" in c.text for c in rows[r].values()):
            today = next((c.text[:10] for c in sorted(rows[r].values(), key=lambda c: c.col) if DATE.match(c.text)), "")
            break
    sh = ArSheet(file, sheet, hrow, roles, today, _num_below(rows, hrow, "total account receivable"),
                 _num_below(rows, hrow, "total amount overdue"))
    own_cols = set(roles.values())
    last_own = max(own_cols)
    side = {}
    for r in sorted(rows):
        if r < hrow:
            sh.above_header += list(rows[r].values())
            continue
        if r == hrow:
            continue
        rc = rows[r]
        sh.side_cells += [c for c in rc.values() if c.col > last_own]
        for c in rc.values():
            if c.col > last_own:
                side.setdefault(r, {})[c.col] = c
        def cell_of(col, _r=r):
            if col in rows[_r]:
                return rows[_r][col]
            o = origin.get((_r, col))
            return rows.get(o[0], {}).get(o[1]) if o else None

        def get(role, _r=r):
            c = cell_of(roles[role]) if role in roles else None
            return c.text if c is not None else ""

        amount_text = get("amount")
        en, cn, no = oc.norm(get("en")), oc.norm(get("cn")), oc.norm(get("invoice_no"))
        mine = {k: v for k, v in rc.items() if k <= last_own}
        if not mine:
            continue
        if no.startswith("#") or en.startswith("#") or cn.startswith("#"):
            if not amount_text:
                sh.dropped.append((r, "placeholder row (#N/A / No Invoice), no amount"))
                continue
        amount = oc.dec(amount_text) if amount_text else None
        if amount is None:
            sh.dropped.append((r, "no amount"))
            continue
        if amount == 0:
            sh.dropped.append((r, "zero amount"))
            continue
        dd = oc.dec(get("due_days")) if get("due_days") else None
        sh.rows.append(ArRow(r, en, cn, no, get("invoice_date")[:10], get("due_date")[:10] if DATE.match(get("due_date")) else "",
                             amount, int(dd) if dd is not None else None, oc.norm(get("remark")),
                             {role: rc[col].coord for role, col in roles.items() if col in rc},
                             merged_from=sorted(role for role, col in roles.items() if col not in rc and cell_of(col) is not None)))
    # the pivot the 0410 workbook keeps beside the list: label / customer, total
    for r, cols in side.items():
        lo = min(cols)
        if lo in cols and lo + 1 in cols and oc.dec(cols[lo + 1].text) is not None and not DATE.match(cols[lo].text):
            if cols[lo].text not in ("总计",):
                sh.pivot[oc.norm(cols[lo].text)] = oc.dec(cols[lo + 1].text)
    return sh


def sheet_kind(name, state, head):
    if state != "visible":
        return "hidden"
    for kind, rx in SHEET_KINDS:
        if rx.search(name):
            return kind
    for kind, rx in SHEET_KINDS:
        if rx.search(head):
            return kind
    return "other"


def analyse(path) -> FileFacts:
    wb = openpyxl.load_workbook(path, read_only=False)
    states = {ws.title: ws.sheet_state for ws in wb.worksheets}
    wb.close()
    merged = {}
    cells = oc.workbook_cells(path, merged)
    by = oc.rows_by_sheet(cells)
    m = re.search(r"(?<!\d)(20\d{2})(\d{2})(\d{2})(?!\d)", os.path.basename(path))
    facts = FileFacts(path, os.path.basename(path), oc.display_name(path), states, cells=cells,
                      name_date=f"{m[1]}-{m[2]}-{m[3]}" if m else "")
    for sheet, state in states.items():
        rows = by.get(sheet, {})
        found = _find_header(rows) if state == "visible" else None
        if found:
            origin = {(r, c): (r0, c0) for r0, c0, r1, c1 in merged.get(sheet, [])
                      for r in range(r0, r1 + 1) for c in range(c0, c1 + 1) if (r, c) != (r0, c0)}
            facts.ar[sheet] = _read_ar(facts.file, sheet, rows, found, origin)
            continue
        head = " ".join(c.text for r in sorted(rows)[:4] for c in rows[r].values() if isinstance(c.text, str))
        facts.other[sheet] = sheet_kind(sheet, state, head)
    return facts


def all_facts(sample_dir=oc.SAMPLE_DIR):
    return [analyse(p) for p in oc.folder_files("09", sample_dir)]


def hidden_sheet_has_invoice_list(path) -> list:
    """Hidden sheets whose header looks like an AR list (they would be a second, silent source of invoices)."""
    wb = openpyxl.load_workbook(path, read_only=False)
    hidden = [ws.title for ws in wb.worksheets if ws.sheet_state != "visible"]
    wb.close()
    by = oc.rows_by_sheet(oc.workbook_cells(path))
    return [s for s in hidden if _find_header(by.get(s, {}))]
