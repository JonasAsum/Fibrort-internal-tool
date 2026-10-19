"""Independent oracle for the sales-order workbooks (folder 08). Built from the cell ledger only:
no importer code is imported. It states, from the raw cells,

  * what every file is (orders / invoice-packing / stock list / formula workbook) and why,
  * for every order sheet: header row, which column is which field, every data row, and every row that is
    not an order line with the reason (blank, banner, total row, column-totals row, repeated header),
  * for every order line: the raw text of each field cell and of every other column,
  * per document (and per file): line count, quantity, net, tax and gross as Decimal sums of the raw rows,
  * the sheets' own total rows, so the sums can be checked against what the file states itself.

Column meaning is decided by the header text with the regular expressions in FIELD_HEADERS. They are written
for the labels that really occur (bilingual Chinese / English ERP exports) and are deliberately strict: the
tax column is the LINE tax (表体税额 / Tax Amount), never the repeated document total (税额 Tax Amount).
An unmapped header is carried as an "other column"; nothing is silently dropped.
"""
import os
import re
from dataclasses import dataclass, field
from decimal import Decimal

from tests.quality import fam3_common as oc

FIELDS = ("document_no", "date", "customer", "salesperson", "status", "product_code", "product_name",
          "quantity", "unit", "net_amount", "tax_amount", "gross_amount")
NUMERIC = ("quantity", "net_amount", "tax_amount", "gross_amount")

# field -> regex over the header text with whitespace collapsed. First column that matches wins; a column
# serves one field.
FIELD_HEADERS = {
    "document_no": r"^(单据编号|Document No\.?)",
    "date": r"^(单据日期|Document Date)",
    "customer": r"^(客户\s+Customer(?: Name)?|Customer)$",
    "salesperson": r"^(销售业务员|Salesperson)",
    "status": r"^(订单状态|Order Status)",
    "product_code": r"^(商品编码|Product Code)",
    "product_name": r"^(商品名称|Product Name)",
    "quantity": r"^(销售数量|Sales Quantity)",
    "unit": r"^(销售单位|Sales Unit)",
    "net_amount": r"^(无税金额|Amount Excl\. Tax$)",
    "tax_amount": r"^(表体税额|Tax Amount in Detail)",
    "gross_amount": r"^(含税金额|Amount Incl\. Tax$)",
}
ENTITY_HEADER = r"^(销售组织|Sales Organization)"
# the document total the Chinese entity repeats on every line (not a line amount)
LOCAL_LINE_NET = r"^表体本币无税金额"
LOCAL_LINE_GROSS = r"^表体本币含税金额"
LOCAL_DOC_NET = r"^(本币无税金额|Amount Excl\. Tax in Functional)"
LOCAL_DOC_GROSS = r"^(本币含税金额|Amount Incl\. Tax in Functional)"
TOTAL_WORDS = ("合计", "总计", "小计", "total", "subtotal", "grand total")
EXPORT_STAMP_WORDS = ("导出时间", "output time")
INVOICE = re.compile(r"invoice\s*#|packing\s*list|commercial invoice|proforma", re.I)
STOCK = re.compile(r"\binventory\b|\bstock\b", re.I)
READER_COLUMN_CAP = 40  # excel/sheet.py reads at most 40 columns


@dataclass
class Line:
    file: str
    sheet: str
    row: int
    text: dict  # field -> raw text ("" when the cell is empty)
    cells: dict  # field -> coordinate
    other: dict  # header label -> raw text, every non-empty cell of an unmapped column
    other_coords: dict  # header label -> coordinate
    num: dict = field(default_factory=dict)  # numeric field -> Decimal | None
    bad_numbers: list = field(default_factory=list)  # (field, text) that is not a plain number
    entity: str = ""
    local_line_net: Decimal | None = None
    local_line_gross: Decimal | None = None
    local_doc_net: Decimal | None = None
    local_doc_gross: Decimal | None = None


@dataclass
class SheetFacts:
    file: str
    sheet: str
    header_row: int
    columns: dict  # field -> (col, label)
    other_columns: dict  # col -> label (a label is repeated with "(2)" when it occurs twice)
    lines: list = field(default_factory=list)
    dropped: list = field(default_factory=list)  # (row, reason)
    above_header: list = field(default_factory=list)  # ledger Cells
    export_stamp: str = ""
    totals_rows: list = field(default_factory=list)  # {row, label, values: {col: Decimal}}
    beyond_cap: list = field(default_factory=list)  # ledger Cells in columns > READER_COLUMN_CAP below the header
    header_cells: list = field(default_factory=list)


@dataclass
class FileFacts:
    path: str
    file: str  # base name as on disk
    display: str
    kind: str
    reason: str
    sheets: dict = field(default_factory=dict)  # sheet -> SheetFacts (order sheets only)
    other_sheets: dict = field(default_factory=dict)  # sheet -> why it is not an order list
    cells: list = field(default_factory=list)
    export_stamp: str = ""  # "YYYY-MM-DD HH:MM:SS" or ""
    newest_order_day: str = ""


def _h(text):
    return oc.norm(text)


def _map_columns(header_cells):
    """({field: (col, label)}, entity_col|None, {col: label} of every other header)"""
    cols, taken = {}, set()
    for fld in FIELDS:
        rx = re.compile(FIELD_HEADERS[fld])
        for c in header_cells:
            if c.col not in taken and rx.search(_h(c.text)):
                cols[fld] = (c.col, _h(c.text))
                taken.add(c.col)
                break
    return cols, taken


def _header_row(rows):
    for r in sorted(rows):
        if r > 30:
            break
        cells = [c for c in rows[r].values() if isinstance(c.text, str)]
        cols, _ = _map_columns(cells)
        if "document_no" in cols and len(cols) >= 4:
            return r, cols, cells
    return None


def _date_part(text):
    return text[:10] if re.match(r"\d{4}-\d{2}-\d{2}", text) else ""


def _stamp(text):
    m = re.match(r"(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2}:\d{2})", text)
    return f"{m[1]} {m[2]}" if m else ""


def _read_sheet(file, sheet, rows, found):
    hrow, cols, hcells = found
    taken = {c for c, _ in cols.values()}
    used, other = {}, {}
    for c in sorted(hcells, key=lambda c: c.col):
        label = _h(c.text)
        if c.col in taken or not label:
            continue
        used[label] = used.get(label, 0) + 1
        other[c.col] = label if used[label] == 1 else f"{label} ({used[label]})"
    entity_col = next((c.col for c in hcells if re.search(ENTITY_HEADER, _h(c.text))), None)
    lc = {k: next((c.col for c in hcells if re.search(rx, _h(c.text))), None)
          for k, rx in (("ln", LOCAL_LINE_NET), ("lg", LOCAL_LINE_GROSS), ("dn", LOCAL_DOC_NET), ("dg", LOCAL_DOC_GROSS))}
    sf = SheetFacts(file, sheet, hrow, cols, other, header_cells=hcells)
    for r in sorted(rows):
        if r < hrow:
            sf.above_header += list(rows[r].values())
            for c in rows[r].values():
                if isinstance(c.text, str) and any(w in _h(c.text).lower() for w in EXPORT_STAMP_WORDS):
                    for c2 in sorted(rows[r].values(), key=lambda x: x.col):
                        if c2.col > c.col and _stamp(c2.text):
                            sf.export_stamp = _stamp(c2.text)
                            break
            continue
        if r == hrow:
            continue
        rc = rows[r]
        sf.beyond_cap += [c for c in rc.values() if c.col > READER_COLUMN_CAP]
        within = {k: v for k, v in rc.items() if k <= READER_COLUMN_CAP}
        doc_col = cols["document_no"][0]
        first = min(within.values(), key=lambda c: c.col) if within else None
        label = _h(first.text).lower() if first is not None else ""
        if first is not None and any(label == w or label.startswith(w + " ") or (not w.isascii() and label.startswith(w))
                                     for w in TOTAL_WORDS):
            sf.dropped.append((r, "total row"))
            sf.totals_rows.append({"row": r, "label": first.text,
                                   "values": {c.col: oc.dec(c.text) for c in within.values() if oc.dec(c.text) is not None}})
            continue
        if doc_col not in rc:
            numeric = all(oc.dec(c.text) is not None for c in within.values())
            if numeric and len(within) <= 3:
                sf.dropped.append((r, "column totals row (numbers only, no label)"))
                sf.totals_rows.append({"row": r, "label": "", "values": {c.col: oc.dec(c.text) for c in within.values()}})
            else:
                sf.dropped.append((r, "no document number"))
            continue
        again, _ = _map_columns([c for c in rc.values() if isinstance(c.text, str)])
        if "document_no" in again and len(again) >= 3:
            sf.dropped.append((r, "repeated header row"))
            continue
        text, coords, num, bad = {}, {}, {}, []
        for fld in FIELDS:
            c = rc.get(cols[fld][0]) if fld in cols else None
            text[fld] = c.text if c is not None else ""
            coords[fld] = c.coord if c is not None else ""
            if fld in NUMERIC:
                num[fld] = oc.dec(c.text) if c is not None else None
                if c is not None and num[fld] is None:
                    bad.append((fld, c.text))
        oth = {label: rc[col].text for col, label in other.items() if col in rc}
        oth_c = {label: rc[col].coord for col, label in other.items() if col in rc}
        ln = Line(file, sheet, r, text, coords, oth, oth_c, num, bad)
        ln.entity = oc.norm(rc[entity_col].text) if entity_col and entity_col in rc else ""
        get = lambda k: oc.dec(rc[lc[k]].text) if lc[k] and lc[k] in rc else None  # noqa: E731
        ln.local_line_net, ln.local_line_gross, ln.local_doc_net, ln.local_doc_gross = get("ln"), get("lg"), get("dn"), get("dg")
        sf.lines.append(ln)
    return sf


def analyse(path) -> FileFacts:
    # declared rule: a cell holding only white space is blank
    cells = [c for c in oc.workbook_cells(path) if oc.norm(c.text) or c.formula]
    by_sheet = oc.rows_by_sheet(cells)
    facts = FileFacts(path=path, file=os.path.basename(path), display=oc.display_name(path), kind="", reason="", cells=cells)
    for sheet, rows in by_sheet.items():
        found = _header_row(rows)
        if found:
            facts.sheets[sheet] = _read_sheet(facts.file, sheet, rows, found)
    others = [s for s in by_sheet if s not in facts.sheets]
    # which workbook is a calculation over a copy of the order list: other sheets hold formulas that read an order sheet
    refs = tuple(f"'{s}'!" for s in facts.sheets) + tuple(f"{s}!" for s in facts.sheets if re.fullmatch(r"\w+", s))
    formula_hits = sum(1 for c in cells if c.sheet in others and c.formula and any(r in c.formula for r in refs))
    for s in others:
        head = " ".join([s] + [c.text for c in cells if c.sheet == s and c.row <= 12 and isinstance(c.text, str)])
        facts.other_sheets[s] = ("invoice/packing list" if INVOICE.search(head) else
                                 "inventory/stock list" if STOCK.search(head) else "not a sales order list")
    if facts.sheets and formula_hits >= 3:
        facts.kind, facts.reason = "formula_workbook", "formula workbook (calculations over a copy of the order list)"
    elif facts.sheets:
        facts.kind = "orders"
    elif any(v == "invoice/packing list" for v in facts.other_sheets.values()):
        facts.kind, facts.reason = "invoice_packing", "invoice/packing list"
    elif any(v == "inventory/stock list" for v in facts.other_sheets.values()):
        facts.kind, facts.reason = "stock_list", "inventory/stock list"
    else:
        facts.kind, facts.reason = "unrecognised", "no sheet looks like a sales order list"
    stamps = [s.export_stamp for s in facts.sheets.values() if s.export_stamp]
    facts.export_stamp = max(stamps) if stamps else ""
    days = [_date_part(ln.text["date"]) for s in facts.sheets.values() for ln in s.lines if _date_part(ln.text["date"])]
    facts.newest_order_day = max(days) if days else ""
    return facts


# ---------- totals ----------

def doc_totals(lines):
    """{(entity, document_no): {lines, quantity, net, tax, gross}} Decimal sums of the raw rows."""
    out = {}
    for ln in lines:
        t = out.setdefault((ln.entity, ln.text["document_no"]),
                           {"lines": 0, "quantity": Decimal(0), "net": Decimal(0), "tax": Decimal(0), "gross": Decimal(0)})
        t["lines"] += 1
        for key, fld in (("quantity", "quantity"), ("net", "net_amount"), ("tax", "tax_amount"), ("gross", "gross_amount")):
            t[key] += ln.num.get(fld) or Decimal(0)
    return out


def line_key(ln):
    """What makes two lines of one document 'the same line' on two sheets that repeat each other."""
    return (ln.text["product_code"], ln.text["product_name"], ln.num["quantity"], ln.num["net_amount"], ln.num["gross_amount"])


def merged_lines(lines):
    """The lines of ONE document in ONE file when its sheets repeat each other (a 'completed' / 'open' cut of
    the total list): every distinct line counted at the highest multiplicity any one sheet has."""
    by_sheet = {}
    for ln in lines:
        by_sheet.setdefault(ln.sheet, []).append(ln)
    if len(by_sheet) == 1:
        return list(lines)
    wanted = {}
    for sl in by_sheet.values():
        counts = {}
        for ln in sl:
            counts[line_key(ln)] = counts.get(line_key(ln), 0) + 1
        for k, n in counts.items():
            wanted[k] = max(wanted.get(k, 0), n)
    kept, seen = [], {}
    for sl in by_sheet.values():
        for ln in sl:
            k = line_key(ln)
            if seen.get(k, 0) < wanted[k]:
                seen[k] = seen.get(k, 0) + 1
                kept.append(ln)
    return kept


def file_order_totals(facts: FileFacts):
    """{(entity, doc): totals} for an order file, sheets that repeat each other counted once."""
    groups = {}
    for s in facts.sheets.values():
        for ln in s.lines:
            groups.setdefault((ln.entity, ln.text["document_no"]), []).append(ln)
    return {k: doc_totals(merged_lines(v))[k] for k, v in groups.items()}


def all_facts(sample_dir=oc.SAMPLE_DIR):
    return [analyse(p) for p in oc.folder_files("08", sample_dir)]
