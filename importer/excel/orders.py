"""Read-only reader for Sales Order list workbooks (one row per order line). Pure parsing:
nothing here knows about the CRM. Headers are found by label (English, Chinese, German, any
column order); rows that are not order lines are skipped with a reason, never imported."""
import datetime as dt
import hashlib
import os
import re
import unicodedata
from dataclasses import dataclass, field

from .companies import norm as company_norm
from .readers import ReviewItem
from .files import file_display_name
from .split import cell_text
from .sheet import Sheet, load, open_workbook

SCAN_ROWS = 30
FIELDS = ("document_no", "date", "customer", "salesperson", "status", "product_code", "product_name",
          "quantity", "unit", "net_amount", "tax_amount", "gross_amount", "currency")
# What a sheet should carry besides the document number; missing ones are reported, not fatal.
EXPECTED = ("date", "customer", "product_code", "quantity", "gross_amount")
CONTENT = ("customer", "product_code", "product_name", "quantity", "net_amount", "gross_amount")


class OrdersError(ValueError):
    """The file cannot be read as a workbook at all (wrong type, empty, corrupt)."""


def label_key(label) -> str:
    """NFKC (full-width to ASCII), lowercase, every run of punctuation/space to one space."""
    text = unicodedata.normalize("NFKC", "" if label is None else str(label)).lower()
    return re.sub(r"[^\w]+|_", " ", text).strip()


# field -> [(lang, synonym)] in priority order. A leading "=" means the whole label must equal it.
_RAW = {
    "document_no": [("en", "document no"), ("zh", "单据编号"), ("en", "sales order no"), ("en", "order no"),
                    ("en", "order number"), ("zh", "销售订单号"), ("zh", "订单编号"), ("de", "belegnummer"),
                    ("de", "beleg nr"), ("de", "auftragsnummer"), ("de", "auftrags nr"), ("de", "auftrag nr"),
                    ("en", "doc no")],
    "date": [("en", "document date"), ("zh", "单据日期"), ("en", "order date"), ("zh", "订单日期"),
             ("de", "belegdatum"), ("de", "auftragsdatum"), ("de", "bestelldatum"), ("en", "=date"),
             ("de", "=datum")],
    "customer": [("en", "customer name"), ("en", "customer"), ("zh", "客户"), ("de", "kundenname"), ("de", "kunde")],
    "salesperson": [("en", "salesperson"), ("en", "sales person"), ("en", "sales rep"), ("zh", "销售业务员"),
                    ("zh", "业务员"), ("de", "vertriebsmitarbeiter"), ("de", "vertreter"),
                    ("de", "verkäufer"), ("de", "verkaeufer")],
    "status": [("en", "order status"), ("en", "order stats"), ("zh", "订单状态"), ("de", "auftragsstatus"),
               ("en", "=status"), ("de", "=status")],
    "product_code": [("en", "product code"), ("zh", "商品编码"), ("zh", "产品编码"), ("zh", "物料编码"),
                     ("en", "artikle"), ("en", "article no"), ("en", "article code"), ("en", "item code"),
                     ("de", "artikelnummer"), ("de", "artikel nr"), ("de", "materialnummer")],
    "product_name": [("en", "product name"), ("zh", "商品名称"), ("zh", "产品名称"), ("de", "artikelbezeichnung"),
                     ("de", "bezeichnung"), ("en", "item description"), ("en", "=description")],
    "quantity": [("en", "sales quantity"), ("en", "sales qty"), ("zh", "销售数量"), ("de", "verkaufsmenge"),
                 ("en", "=quantity"), ("en", "=qty"), ("zh", "数量"), ("de", "=menge")],
    "unit": [("en", "sales unit"), ("zh", "销售单位"), ("de", "verkaufseinheit"), ("en", "primary uom"),
             ("zh", "主计量"), ("en", "=unit"), ("en", "=uom"), ("zh", "单位"), ("de", "=einheit")],
    "net_amount": [("en", "=amount excl tax"), ("en", "amount excl tax"), ("en", "amount excluding tax"),
                   ("zh", "=无税金额"), ("zh", "无税金额"), ("en", "net amount"), ("en", "net value"),
                   ("de", "nettobetrag"), ("de", "nettowert"), ("de", "=netto"), ("zh", "本币无税金额")],
    # The LINE tax first: "表体税额" / "Tax Amount in Detail". The plain "Tax Amount" / "税额" column of the ERP
    # exports is the document's total tax repeated on every line (summing it counts the tax once per line).
    "tax_amount": [("zh", "表体税额"), ("en", "tax amount in detail"), ("en", "=tax amount"), ("en", "tax amount"),
                   ("en", "tax value"),
                   ("zh", "税额"), ("de", "steuerbetrag"), ("de", "mwst betrag"), ("de", "=mwst")],
    "gross_amount": [("en", "=amount incl tax"), ("zh", "含税金额"), ("en", "amount incl tax"),
                     ("en", "gross amount"), ("en", "gross value"), ("de", "bruttobetrag"),
                     ("de", "bruttowert"), ("de", "=brutto"), ("zh", "本币含税金额")],
    "currency": [("en", "=currency"), ("zh", "币种"), ("zh", "币别"), ("de", "währung"), ("de", "waehrung")],
}
SYNONYMS = {f: [(lang, label_key(s[1:]) if s.startswith("=") else label_key(s), s.startswith("=")) for lang, s in lst]
            for f, lst in _RAW.items()}

# Labels containing these never fill a field (unless the synonym itself says so).
_REJECT = {
    "document_no": r"purchase|采购|\bpo\b",
    "date": r"planned|delivery|ship|expected|desired|计划|期望|\bfrom\b",
    "customer": r"date|日期|期望|expected|desired|type|分类|category|short|code|编码|\bno\b|号",
    "status": r"payment|收款",
    "quantity": r"price|单价|amount|金额",
    "unit": r"price|单价",
    "net_amount": r"price|单价|rate|税率|税额",
    "tax_amount": r"rate|率|category|分类|code|编码",
    "gross_amount": r"price|单价|rate|税率|税额",
    "product_code": r"tax|税收|分类",
    "product_name": r"tax|税收",
}
# Columns that hold a bank account or a tax id are never read, not even into a line's `extra`.
_SENSITIVE = re.compile(r"bank|iban|swift|银行|账号|帐号|账户|纳税人|税号|\btax\s*(id|no|number)\b|\bvat\s*(id|no|number)\b|"
                        r"ust\W*id|steuer\W*(nr|nummer|id)|account\s*(no|number)", re.I)
_LOCAL = ("本币", "local", "functional", "表体", "fc")
_TOTAL_WORDS = ("合计", "总计", "小计", "总额", "total", "subtotal", "sub total", "grand total", "summe",
                "gesamt", "gesamtsumme", "zwischensumme")


def _has(token, text):
    return re.search(r"(?<![a-z])%s(?![a-z])" % re.escape(token), text) if token.isascii() else token in text


def map_headers(cells):
    """cells: [(col, label)] -> {field: (col, header_text, lang)}; each column serves one field."""
    normed = [(c, label, label_key(label)) for c, label in cells if label_key(label)]
    taken, out = set(), {}
    for fld in FIELDS:
        reject = re.compile(_REJECT[fld]) if fld in _REJECT else None
        for lang, syn, exact in SYNONYMS[fld]:
            hit = None
            for c, label, t in normed:
                if c in taken or not (t == syn if exact else syn in t):
                    continue
                if reject and reject.search(t) and not reject.search(syn):
                    continue
                if any(_has(tok, t) and tok not in syn for tok in _LOCAL):
                    continue
                hit = (c, label, lang)
                break
            if hit:
                taken.add(hit[0])
                out[fld] = hit
                break
    return out


def find_order_header(sheet: Sheet):
    """(row, {field: (col, label, lang)}) of the first row naming a document number column and
    at least two more recognised columns, one of them describing the line; else None."""
    for r in range(1, min(SCAN_ROWS, sheet.max_row) + 1):
        cells = [(c, sheet.value(r, c)) for c in range(1, sheet.max_col + 1)
                 if not sheet.is_merged_continuation(r, c) and isinstance(sheet.value(r, c), str)]
        cols = map_headers(cells)
        if "document_no" in cols and len(cols) >= 3 and any(f in cols for f in CONTENT):
            return r, cols
    return None


def looks_like_orders(sheet: Sheet) -> bool:
    return find_order_header(sheet) is not None


@dataclass
class OrderLine:
    document_no: str
    date_iso: str | None
    customer: str | None
    salesperson: str | None
    status: str | None
    product_code: str | None
    product_name: str | None
    quantity: float | None
    unit: str | None
    net_amount: float | None
    tax_amount: float | None
    gross_amount: float | None
    currency: str | None
    row: int
    sheet: str = ""
    issues: list = field(default_factory=list)  # problems found in this line's cells
    extra: dict = field(default_factory=dict)  # every other column, by header label


@dataclass
class OrderDocument:
    document_no: str
    date_iso: str | None = None
    customer: str | None = None
    line_count: int = 0
    quantity: float = 0.0
    net_amount: float = 0.0
    tax_amount: float = 0.0
    gross_amount: float = 0.0
    sheets: list = field(default_factory=list)
    rows: list = field(default_factory=list)


@dataclass
class SkippedRow:
    sheet: str
    row: int
    reason: str


@dataclass
class OrdersResult:
    path: str = ""
    lines: list = field(default_factory=list)
    documents: dict = field(default_factory=dict)  # document_no -> OrderDocument
    skipped: list = field(default_factory=list)
    skipped_sheets: list = field(default_factory=list)  # (sheet name, reason)
    headers_found: dict = field(default_factory=dict)  # sheet name -> {field: header label}
    layout: str | None = None  # "en", "zh" or "de"; None when no sheet was recognised
    warnings: list = field(default_factory=list)
    kind: str = "orders"  # what the file is, see classify(): only "orders" is CRM data
    reason: str = ""  # why a file of another kind is not CRM data
    excluded_sheets: list = field(default_factory=list)  # (sheet name, reason): sheets of an order file that are not order lists
    export_time: dt.datetime | None = None  # the export stamp above the header ("导出时间 Output Time"), when there is one


# ---------- value parsing ----------

_DATE_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
                 "%d.%m.%Y %H:%M:%S", "%Y年%m月%d日", "%d-%b-%Y", "%d %b %Y", "%Y%m%d")


def parse_date(v):
    """ISO date string, or None when v is empty or not a date (caller tells them apart)."""
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        if 20000 <= v <= 80000:
            return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date().isoformat()
        return None
    if isinstance(v, str):
        s = v.strip()
        for fmt in _DATE_FORMATS:
            try:
                return dt.datetime.strptime(s, fmt).date().isoformat()
            except ValueError:
                pass
    return None


def parse_number(v):
    """float, or None when v is empty or not a number."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if not isinstance(v, str):
        return None
    s = unicodedata.normalize("NFKC", v).replace("\xa0", "").replace(" ", "")
    s = re.sub(r"(?i)rmb|cny|eur|usd|¥|€|\$|元|￥", "", s)
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1]
    if s.endswith("-"):
        neg, s = True, s[:-1]
    if s.startswith("-"):
        neg, s = not neg, s[1:]
    s = s.lstrip("+")
    if not s or not re.fullmatch(r"[\d.,']+", s):
        return None
    s = s.replace("'", "")
    if "," in s and "." in s:
        dec = "," if s.rfind(",") > s.rfind(".") else "."
        s = s.replace("." if dec == "," else ",", "").replace(dec, ".")
    elif "," in s:
        s = s.replace(",", "") if re.fullmatch(r"\d{1,3}(,\d{3})+", s) else s.replace(",", ".")
    elif s.count(".") > 1:
        if not re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
            return None
        s = s.replace(".", "")
    try:
        n = float(s)
    except ValueError:
        return None
    return -n if neg else n


def _doc_no(v):
    return (cell_text(v) or None)


# ---------- reading ----------

def _is_total_row(sheet, r, ncols):
    seen = 0
    for c in range(1, ncols + 1):
        v = sheet.value(r, c)
        if v is None or (isinstance(v, str) and not v.strip()):
            continue
        seen += 1
        if isinstance(v, str):
            t = label_key(v)
            if any(t == w or t.startswith(w + " ") or (w.isascii() is False and t.startswith(w)) for w in _TOTAL_WORDS):
                return True
        if seen >= 3:
            break
    return False


def _row_class(sheet, r, ncols):
    """None for a candidate line, else the reason the row is not one."""
    own, cont = [], []
    for c in range(1, ncols + 1):
        if sheet.value(r, c) is None:
            continue
        (cont if sheet.is_merged_continuation(r, c) else own).append(c)
    if not own and not cont:
        return "blank row"
    if not own:
        return "merged cells without line data of its own"
    if len(own) == 1 and cont and all(sheet.top.get((r, c)) == (r, own[0]) for c in cont):
        return "merged banner row"
    return None


def _read_sheet(sheet, result, seen_formula):
    found = find_order_header(sheet)
    if found is None:
        result.skipped_sheets.append((sheet.name, "no header row with a document number column"))
        return None
    hrow, cols = found
    stamp = _export_time(sheet, hrow)
    if stamp and (result.export_time is None or stamp > result.export_time):
        result.export_time = stamp
    result.headers_found[sheet.name] = {f: (cell_text(label) or None) for f, (_, label, _) in cols.items()}
    missing = [f for f in EXPECTED if f not in cols]
    if missing:
        result.warnings.append("sheet %r: no column found for %s" % (sheet.name, ", ".join(missing)))
    if sheet.cut:
        result.warnings.append("sheet %r: %s" % (sheet.name, sheet.cut))
    mapped = {v[0]: f for f, v in cols.items()}
    ncols = sheet.max_col
    names, used = {}, {}
    for c in range(1, ncols + 1):
        label = (cell_text(sheet.value(hrow, c)) or None) if not sheet.is_merged_continuation(hrow, c) else None
        if c in mapped or not label:
            continue
        if _SENSITIVE.search(label):
            result.warnings.append("sheet %r: column %r not read (bank account or tax id)" % (sheet.name, label))
            continue
        used[label] = used.get(label, 0) + 1
        names[c] = label if used[label] == 1 else "%s (%d)" % (label, used[label])
    langs = [v[2] for v in cols.values()]
    n_before = len(result.lines)
    for r in range(hrow + 1, sheet.max_row + 1):
        why = _row_class(sheet, r, ncols)
        if why is None and _is_total_row(sheet, r, ncols):
            why = "total or subtotal row"
        if why is None:
            again = map_headers([(c, sheet.value(r, c)) for c in range(1, ncols + 1)
                                 if isinstance(sheet.value(r, c), str) and not sheet.is_merged_continuation(r, c)])
            if "document_no" in again and len(again) >= 3:
                why = "repeated header row"
        if why:
            result.skipped.append(SkippedRow(sheet.name, r, why))
            continue
        issues = []

        def get(fld):
            c = cols[fld][0] if fld in cols else None
            if c is None:
                return None, None
            v = sheet.value(r, c)
            if v is None or (isinstance(v, str) and not v.strip()):
                try:
                    formula = sheet.formula(r, c) if fld in ("document_no", "date", "quantity", "net_amount",
                                                            "tax_amount", "gross_amount") else None
                except Exception:
                    formula = None
                if formula:
                    seen_formula.append((sheet.name, r))
                    return None, "%s: formula without cached value" % fld
                return None, None
            return v, None

        raw, problem = get("document_no")
        doc = _doc_no(raw)
        if doc is None:
            result.skipped.append(SkippedRow(sheet.name, r, problem or "no document number"))
            continue
        vals = {}
        for fld in FIELDS[1:]:
            v, problem = get(fld)
            if problem:
                issues.append(problem)
            if fld == "date":
                vals[fld] = parse_date(v)
                if v is not None and vals[fld] is None:
                    issues.append("date: not a date")
            elif fld in ("quantity", "net_amount", "tax_amount", "gross_amount"):
                vals[fld] = parse_number(v)
                if v is not None and vals[fld] is None:
                    issues.append("%s: not a number" % fld)
            else:
                vals[fld] = (cell_text(v) or None)
        extra = {}
        for c, name in names.items():
            v = sheet.value(r, c)
            if v is not None and not (isinstance(v, str) and not v.strip()):
                extra[name] = v
        result.lines.append(OrderLine(document_no=doc, date_iso=vals["date"], customer=vals["customer"],
                                      salesperson=vals["salesperson"], status=vals["status"],
                                      product_code=vals["product_code"], product_name=vals["product_name"],
                                      quantity=vals["quantity"], unit=vals["unit"], net_amount=vals["net_amount"],
                                      tax_amount=vals["tax_amount"], gross_amount=vals["gross_amount"],
                                      currency=vals["currency"], row=r, sheet=sheet.name, issues=issues,
                                      extra=extra))
    return langs, len(result.lines) - n_before


def _group(result):
    docs = result.documents
    customers = {}
    for ln in result.lines:
        d = docs.get(ln.document_no)
        if d is None:
            d = docs[ln.document_no] = OrderDocument(ln.document_no, ln.date_iso, ln.customer)
        d.date_iso = d.date_iso or ln.date_iso
        d.customer = d.customer or ln.customer
        d.line_count += 1
        d.quantity += ln.quantity or 0.0
        d.net_amount += ln.net_amount or 0.0
        d.tax_amount += ln.tax_amount or 0.0
        d.gross_amount += ln.gross_amount or 0.0
        if ln.sheet not in d.sheets:
            d.sheets.append(ln.sheet)
        d.rows.append((ln.sheet, ln.row))
        if ln.customer:
            customers.setdefault(ln.document_no, set()).add(ln.customer)
    clash = sum(1 for s in customers.values() if len(s) > 1)
    if clash:
        result.warnings.append("%d document(s) have lines with different customers" % clash)
    multi = sum(1 for d in docs.values() if len(d.sheets) > 1)
    if multi:
        result.warnings.append("%d document(s) appear on more than one sheet; their totals count every sheet" % multi)


# ---------- what kind of file it is ----------
# The folder name is no evidence: files called "..._Order_List" are invoice and packing lists. Only the
# content decides, and only the ERP sales-order exports are CRM data.
NOT_CRM = "not CRM data: "
INVOICE_REASON = NOT_CRM + "invoice/packing list"
STOCK_REASON = NOT_CRM + "inventory/stock list"
FORMULA_REASON = NOT_CRM + "formula workbook (calculations over a copy of the order list)"
_INVOICE_WORDS = re.compile(r"invoice\s*#|packing\s*list|commercial invoice|proforma|packliste|装箱单|发票号码", re.I)
_STOCK_WORDS = re.compile(r"\binventory\b|\bstock\b|bestand|库存盘点", re.I)
_EXPORT_STAMP = ("导出时间", "output time", "export time", "exported at", "export date")
_FORMULA_HITS = 3  # formulas on other sheets that read an order sheet before the workbook counts as a calculation


def _head_text(sheet, rows=12) -> str:
    cells = (sheet.value(r, c) for r in range(1, min(rows, sheet.max_row) + 1) for c in range(1, sheet.max_col + 1))
    return " ".join([sheet.name] + [str(v) for v in cells if isinstance(v, str)])


def _sheet_reason(sheet) -> str:
    head = _head_text(sheet)
    if _INVOICE_WORDS.search(head):
        return INVOICE_REASON
    if _STOCK_WORDS.search(head):
        return STOCK_REASON
    return "not a sales order list"


def _calculates_over(path, order_sheets, other_sheets) -> bool:
    """True when sheets other than the order lists hold formulas that read an order sheet (a report
    built on a copy of the list). Cached values are not enough: formulas are read in a second,
    read-only pass, only for .xlsx workbooks that have both kinds of sheet."""
    if not other_sheets or not path.lower().endswith((".xlsx", ".xlsm")):
        return False
    try:
        wb = open_workbook(path, read_only=True)
    except Exception:
        return False
    refs = tuple(["'%s'!" % n for n in order_sheets] + ["%s!" % n for n in order_sheets if re.fullmatch(r"\w+", n)])
    hits = 0
    try:
        for name in other_sheets:
            for row in wb[name].iter_rows(max_row=500, max_col=60, values_only=True):
                for v in row:
                    if isinstance(v, str) and v.startswith("=") and any(ref in v for ref in refs):
                        hits += 1
                        if hits >= _FORMULA_HITS:
                            return True
    except Exception:
        return False
    finally:
        wb.close()
    return False


def classify(path, sheets, result) -> None:
    """Sets result.kind and reason: "orders" (an ERP sales-order export), "formula_workbook",
    "invoice_packing", "stock_list" or "unrecognised"; lists the sheets of an order file that are not
    order lists in result.excluded_sheets. The lines stay in the result either way; callers import
    only kind == "orders"."""
    ordered = set(result.headers_found)
    hidden = {n for n, why in result.skipped_sheets if why.startswith("sheet is ")}
    others = [sh for sh in sheets if sh.name not in ordered and sh.name not in hidden]
    reasons = [(sh.name, _sheet_reason(sh)) for sh in others]
    if ordered:
        if _calculates_over(path, ordered, [sh.name for sh in others]):
            result.kind, result.reason = "formula_workbook", FORMULA_REASON
            return
        result.kind, result.reason, result.excluded_sheets = "orders", "", reasons
        return
    for kind, why in (("invoice_packing", INVOICE_REASON), ("stock_list", STOCK_REASON)):
        if any(r == why for _, r in reasons):
            result.kind, result.reason, result.excluded_sheets = kind, why, reasons
            return
    result.kind, result.reason, result.excluded_sheets = "unrecognised", "no sheet looks like a sales order list", reasons


def _stamp(v):
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, str):
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(v.strip(), fmt)
            except ValueError:
                pass
    return None


def _export_time(sheet, header_row):
    """The export stamp in the rows above the header ("导出时间 Output Time": 2026-04-16 22:37:33)."""
    for r in range(1, header_row):
        for c in range(1, sheet.max_col + 1):
            v = sheet.value(r, c)
            if isinstance(v, str) and any(w in label_key(v) for w in _EXPORT_STAMP) and not sheet.is_merged_continuation(r, c):
                for c2 in range(c + 1, sheet.max_col + 1):
                    cell = sheet.value(r, c2)
                    if cell is not None:
                        return _stamp(cell)
    return None


def read_orders(path) -> OrdersResult:
    """All order lines of every visible sheet that looks like an order list. Raises OrdersError
    when the file is not a readable .xlsx/.xls workbook; everything else is reported in the result."""
    path = str(path)
    if not os.path.isfile(path):
        raise OrdersError("file not found")
    if os.path.getsize(path) == 0:
        raise OrdersError("file is empty")
    if not path.lower().endswith((".xlsx", ".xlsm", ".xls")):
        raise OrdersError("not an Excel workbook (.xlsx or .xls expected)")
    try:
        sheets = load(path)
    except Exception as e:  # corrupt, wrong content behind the extension, password protected ...
        raise OrdersError("cannot open as a workbook: %s" % type(e).__name__) from e
    result = OrdersResult(path=path)
    if not sheets:
        result.warnings.append("workbook has no sheets")
    seen_formula, all_langs = [], []
    for sh in sheets:
        state = getattr(sh.ws, "sheet_state", "visible") if sh.ws is not None else "visible"
        if state != "visible":
            result.skipped_sheets.append((sh.name, "sheet is %s" % state))
            continue
        out = _read_sheet(sh, result, seen_formula)
        if out:
            all_langs += out[0]
    if seen_formula:
        result.warnings.append("%d cell(s) hold a formula without a cached value (open and save the file in Excel)"
                               % len(seen_formula))
    if all_langs:
        counts = {k: all_langs.count(k) for k in ("zh", "de", "en")}
        result.layout = max(("zh", "de", "en"), key=lambda k: counts[k])
    else:
        result.warnings.append("no sheet looks like a sales order list")
    for name, why in result.skipped_sheets:
        result.warnings.append("sheet %r skipped: %s" % (name, why))
    _group(result)
    classify(path, sheets, result)
    return result


# ---------- orders: the lines of all exports rolled up per document ----------
# Plain records first (OrdersBook); the CRM mapping lives in excel/orders_margince.py.
FAMILY = "orders"
COMPANY_OK = 0.9
STATUS_WORDS = (  # (key, words); the first row with a hit wins, so "partially received" is not "received"
    ("cancelled", ("cancel", "取消", "作废", "已关闭")),
    ("in_progress", ("partially", "to receive", "部分", "待收货")),
    ("completed", ("completed", "已完成")),
    ("to_ship", ("to ship", "pending delivery", "待发货")),
    ("open", ("open", "开立")),
)
LOCAL_ENTITY = "有限公司"  # a Chinese legal entity: its local / functional currency columns are RMB


def status_key(text) -> str:
    t = unicodedata.normalize("NFKC", text or "").lower()
    return next((k for k, words in STATUS_WORDS if any(w in t for w in words)), "other")


@dataclass
class OrderItem:
    product_code: str | None
    product_name: str | None
    quantity: float | None
    unit: str | None
    net_amount: float | None
    gross_amount: float | None


@dataclass
class Order:
    """One sales order (document number) as the newest export that holds it states it."""
    document_no: str
    date_iso: str | None
    customer: str
    salesperson: str | None
    status: str | None  # as written
    status_key: str  # open | to_ship | in_progress | completed | cancelled | other
    entity: str | None  # the selling company ("Sales Organization")
    net: float  # sums of the lines, in the currency the export states them in (not always named)
    tax: float
    gross: float
    currency: str | None  # "CNY" only when the export proves it; None = not stated
    cny_net: float | None  # the Chinese entity's local-currency totals; None for other entities
    cny_gross: float | None
    payment_status: str | None
    collected: float | None
    planned_ship: str | None
    items: list
    file: str  # the export it was read from (hash prefix removed)
    sheet: str
    export_iso: str  # when that export was made (its stamp, else the day of its newest order)
    row: int = 0
    company_key: str = ""
    issues: list = field(default_factory=list)
    superseded: list = field(default_factory=list)  # older exports that held the same document: [file]
    id_prefix: str = ""  # set only when two selling companies use the same document number

    @property
    def source_id(self) -> str:
        return "order:" + self.id_prefix + self.document_no


@dataclass
class Org:
    key: str
    name: str  # the name the CRM gets: the master's when it knows the company, else as written
    written_names: list = field(default_factory=list)
    match: dict = field(default_factory=dict)  # code, cn, en, score, method
    queued: bool = False  # a person must answer "same company?" before its orders go in


@dataclass
class FileInfo:
    path: str
    name: str
    kind: str = ""
    reason: str = ""
    lines: int = 0
    documents: int = 0
    export_iso: str = ""
    excluded_sheets: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    kept_documents: int = 0  # documents this export is the newest holder of


@dataclass
class OrdersBook:
    orders: list = field(default_factory=list)
    orgs: dict = field(default_factory=dict)  # company key -> Org
    review: list = field(default_factory=list)  # ReviewItem
    files: list = field(default_factory=list)  # FileInfo of every file read
    skipped_files: list = field(default_factory=list)  # [{"file", "error"}]: unreadable or identical copies

    def accepted(self) -> list:
        """Orders that can be written: dated, their company is known or answered, and nothing in their lines is in doubt."""
        return [o for o in self.orders if o.date_iso and o.company_key and not self.orgs[o.company_key].queued
                and not o.issues]  # an order whose lines do not add up cleanly is wrong data: held out

    def accepted_orgs(self) -> dict:
        used = {o.company_key for o in self.accepted()}
        return {k: g for k, g in self.orgs.items() if k in used}

    def excluded(self) -> list:
        return [f for f in self.files if f.kind != "orders"]


def file_kind(path) -> tuple[str, str]:
    """(kind, reason) of a workbook: "orders" is CRM data; "invoice_packing", "formula_workbook",
    "stock_list" and "unrecognised" are not, and the reason says why."""
    res = read_orders(path)
    return res.kind, res.reason


def _extra(line, *prefixes, avoid=()):
    """The value of the first extra column whose label starts with one of the prefixes."""
    for label, v in line.extra.items():
        t = label_key(label)
        if any(t.startswith(p) for p in prefixes) and not any(a in t for a in avoid):
            return v
    return None


def _local_total(lines, line_prefixes, doc_prefixes):
    """Sum of the lines' own local-currency amounts, else the document total the export repeats on every line."""
    own = [parse_number(_extra(ln, *line_prefixes)) for ln in lines]
    if own and all(v is not None for v in own):
        return round(sum(own), 2)
    doc = {round(v, 2) for ln in lines if (v := parse_number(_extra(ln, *doc_prefixes, avoid=("表体",)))) is not None}
    return doc.pop() if len(doc) == 1 else None


def _line_key(ln):
    return (ln.product_code, ln.product_name, ln.quantity, ln.net_amount, ln.gross_amount)


def _merge_sheets(lines):
    """The lines of one document across the sheets of one file: a sheet that repeats another's lines
    (a "completed" or "open" cut of the total list) adds none; lines only one sheet holds are kept."""
    by_sheet = {}
    for ln in lines:
        by_sheet.setdefault(ln.sheet, []).append(ln)
    if len(by_sheet) == 1:
        return lines
    wanted = {}
    for sheet_lines in by_sheet.values():
        counts = {}
        for ln in sheet_lines:
            counts[_line_key(ln)] = counts.get(_line_key(ln), 0) + 1
        for k, n in counts.items():
            wanted[k] = max(wanted.get(k, 0), n)
    kept, seen = [], {}
    for sheet_lines in by_sheet.values():
        for ln in sheet_lines:
            k = _line_key(ln)
            if seen.get(k, 0) < wanted[k]:
                seen[k] = seen.get(k, 0) + 1
                kept.append(ln)
    return kept


def _richness(lines) -> int:
    return sum(1 for ln in lines for v in (ln.customer, ln.salesperson, ln.status, ln.product_code, ln.product_name,
                                           ln.quantity, ln.net_amount, ln.gross_amount, ln.date_iso, *ln.extra.values())
               if v not in (None, ""))


def _blank(lines) -> bool:
    """A copy that names no customer or states no amount cannot stand for the order."""
    return (not any(ln.customer for ln in lines)
            or all(ln.net_amount is None and ln.gross_amount is None for ln in lines))


def _order(doc, lines, info) -> Order:
    first = lines[0]
    customers = [c for c in dict.fromkeys(ln.customer for ln in lines) if c]
    issues = []
    if len(customers) > 1:
        issues.append("lines name different customers; the first is used")
    entity = (cell_text(_extra(first, "销售组织", "sales organization")) or None)
    net = round(sum(ln.net_amount or 0.0 for ln in lines), 2)
    tax = round(sum(ln.tax_amount or 0.0 for ln in lines), 2)
    gross = round(sum(ln.gross_amount or 0.0 for ln in lines), 2)
    if any(ln.gross_amount is None for ln in lines):
        issues.append("a line states no amount")
    amounts = [ln.gross_amount for ln in lines if ln.gross_amount]
    if len(amounts) >= 3 and any(abs(a - (sum(amounts) - a)) <= 0.05 for a in amounts):
        issues.append("a line equals the sum of the other lines (possible double count)")
    cny_net = cny_gross = None
    if entity and LOCAL_ENTITY in entity:
        cny_net = _local_total(lines, ("表体本币无税金额",), ("本币无税金额", "amount excl tax in functional currency"))
        cny_gross = _local_total(lines, ("表体本币含税金额",), ("本币含税金额", "amount incl tax in functional currency"))
    named = {c for ln in lines if (c := ln.currency)}
    currency = named.pop() if len(named) == 1 else None
    if currency is None and cny_gross is not None and abs(cny_gross - gross) <= 0.05:
        currency = "CNY"  # ordered in the local currency: the transaction amount is the local amount
    collected = parse_number(_extra(first, "已收款金额", "cumulative amount collected"))
    return Order(
        document_no=doc, date_iso=next((ln.date_iso for ln in lines if ln.date_iso), None),
        customer=customers[0] if customers else "", salesperson=first.salesperson, status=first.status,
        status_key=status_key(first.status), entity=entity, net=net, tax=tax, gross=gross, currency=currency,
        cny_net=cny_net, cny_gross=cny_gross, payment_status=(cell_text(_extra(first, "收款状态", "payment status")) or None),
        collected=collected or None,
        planned_ship=parse_date(_extra(first, "计划发货日期", "planned shipping date")),
        items=[OrderItem(ln.product_code, ln.product_name, ln.quantity, ln.unit, ln.net_amount, ln.gross_amount)
               for ln in lines],
        file=info.name, sheet=first.sheet, export_iso=info.export_iso, row=first.row, issues=issues)


def _unique_files(paths, skipped):
    seen, keep = {}, []
    for p in sorted(os.fspath(p) for p in paths):
        try:
            with open(p, "rb") as f:
                digest = hashlib.sha256(f.read()).hexdigest()
        except OSError as exc:
            skipped.append({"file": p, "error": f"unreadable: {exc}"})
            continue
        if digest in seen:
            skipped.append({"file": p, "error": f"identical to {os.path.basename(seen[digest])}"})
        else:
            seen[digest] = p
            keep.append(p)
    return keep


def _read_one(path, skipped):
    try:
        return read_orders(path)
    except OrdersError as exc:
        skipped.append({"file": path, "error": str(exc)})
    except Exception as exc:  # one bad workbook must not stop the run
        skipped.append({"file": path, "error": f"could not be read: {type(exc).__name__}: {exc}"})
    return None


GROUP_NAME = re.compile(r"fibro|菲博", re.I)  # the group's own companies: an order, but not a customer's


def _match_orgs(orders, master, review):
    matches, orgs = {}, {}
    for o in orders:
        if o.customer not in matches:
            m = master.match(o.customer)
            matches[o.customer] = m
            if m.score < COMPANY_OK:
                review.append(ReviewItem(FAMILY, o.file, o.sheet, o.row, "customer", o.customer, "company",
                                         [list(c) for c in (m.candidates or [])]))
        m = matches[o.customer]
        known = m.score >= COMPANY_OK and (m.cn or m.en)
        if known and o.entity and LOCAL_ENTITY not in o.entity and m.method != "code":
            # another sales entity (the US company) sells to its own customers: a China customer of the same name is not proof
            o.issues.append("order of another sales entity matched to a China customer by name only")
        name = (m.cn or m.en) if known else o.customer
        o.company_key = company_norm(name) or name.lower()
        org = orgs.get(o.company_key)
        if org is None:
            org = orgs[o.company_key] = Org(o.company_key, name)
            org.match = {"code": m.code, "cn": m.cn, "en": m.en, "score": m.score, "method": m.method}
            org.queued = m.score < COMPANY_OK
        if o.customer not in org.written_names:
            org.written_names.append(o.customer)
            if GROUP_NAME.search(o.customer):
                review.append(ReviewItem(FAMILY, o.file, o.sheet, o.row, "customer", o.customer, "group_company"))
        if GROUP_NAME.search(o.customer) and not any("group company" in i for i in o.issues):
            o.issues.append("group company: intercompany, not a customer")
    return orgs


def build_book(paths, master=None) -> OrdersBook:
    """Every ERP sales-order export among `paths` rolled up to one Order per document number.
    Byte-identical copies are read once; files that are not order lists (invoice/packing lists,
    stock lists, calculation workbooks) are named with the reason and contribute nothing. When several
    exports hold a document, the newest export's version wins (a hash-prefixed copy of the same export
    ties; the richer one wins); a copy that names no customer or amount never replaces a complete one."""
    if master is None:
        from .run import IMPORTER_DIR, _master
        master = _master(IMPORTER_DIR)
    book = OrdersBook()
    reads = []
    for path in _unique_files(paths, book.skipped_files):
        res = _read_one(path, book.skipped_files)
        if res is None:
            continue
        info = FileInfo(path=path, name=file_display_name(path), kind=res.kind, reason=res.reason, lines=len(res.lines),
                        documents=len(res.documents), excluded_sheets=list(res.excluded_sheets),
                        warnings=list(res.warnings))
        book.files.append(info)
        if res.kind != "orders" or not res.lines:
            if res.kind == "orders":
                info.kind, info.reason = "unrecognised", "no order lines"
            continue
        days = [ln.date_iso for ln in res.lines if ln.date_iso]
        stamp = res.export_time or (dt.datetime.fromisoformat(max(days) + "T23:59:59") if days else dt.datetime.min)
        info.export_iso = stamp.isoformat(sep=" ", timespec="seconds")
        reads.append((info, res, stamp, os.path.getmtime(path)))
    holders = {}  # document -> [(rank, info, lines)]
    for info, res, stamp, mtime in reads:
        per_doc = {}
        for ln in res.lines:
            entity = (cell_text(_extra(ln, "销售组织", "sales organization")) or None) or ""
            per_doc.setdefault((entity, ln.document_no), []).append(ln)
        for key, lines in per_doc.items():
            lines = _merge_sheets(lines)
            holders.setdefault(key, []).append(((stamp, _richness(lines), mtime, info.name), info, lines))
    for (_, doc), found in holders.items():
        found.sort(key=lambda t: t[0], reverse=True)
        usable = [t for t in found if not _blank(t[2])]
        if not usable:
            _, info, lines = found[0]
            book.review.append(ReviewItem(FAMILY, info.name, lines[0].sheet, lines[0].row, "document_no", doc,
                                          "blank_order"))
            continue
        _, info, lines = usable[0]
        order = _order(doc, lines, info)
        order.superseded = [t[1].name for t in found if t[1] is not info]
        others = [round(sum(ln.gross_amount or 0.0 for ln in t[2]), 2) for t in usable[1:]]
        if any(abs(g - order.gross) > 0.5 for g in others):  # two exports give two amounts: which one is right is not known
            order.issues.append("exports disagree on the amount: " + ", ".join(f"{g:,.2f}" for g in [order.gross, *others]))
        info.kept_documents += 1
        if not order.date_iso:
            book.review.append(ReviewItem(FAMILY, info.name, order.sheet, order.row, "date", doc, "no_date"))
        if order.currency is None:
            book.review.append(ReviewItem(FAMILY, info.name, order.sheet, order.row, "currency", doc,
                                          "currency_not_stated"))
        book.orders.append(order)
    numbers = {}
    for o in book.orders:
        numbers[o.document_no] = numbers.get(o.document_no, 0) + 1
    for o in book.orders:  # one number used by two selling companies: two orders, told apart by the seller
        if numbers[o.document_no] > 1:
            o.id_prefix = hashlib.sha1((o.entity or "").encode("utf-8")).hexdigest()[:6] + ":"
    book.orders.sort(key=lambda o: (o.date_iso or "", o.document_no))
    book.orgs = _match_orgs(book.orders, master, book.review)
    book.review += [ReviewItem(FAMILY, o.file, o.sheet, o.row, "order", o.document_no, "order_doubtful")
                    for o in book.orders if o.issues]
    return book


def summarize(book: OrdersBook) -> dict:
    accepted = book.accepted()
    reasons = {}
    for item in book.review:
        reasons[item.reason] = reasons.get(item.reason, 0) + 1
    read = [f for f in book.files if f.kind == "orders"]
    return {
        "files": len(book.files), "files_read": len(read), "files_excluded": len(book.excluded()),
        "lines": sum(f.lines for f in read), "orders": len(book.orders), "orders_accepted": len(accepted),
        "orders_queued": len(book.orders) - len(accepted),
        "companies": len(book.orgs), "companies_accepted": len(book.accepted_orgs()),
        "companies_queued": sum(1 for g in book.orgs.values() if g.queued),
        "currency_not_stated": sum(1 for o in book.orders if o.currency is None),
        "review_by_reason": reasons,
        "excluded": [{"file": f.name, "kind": f.kind, "reason": f.reason} for f in book.excluded()],
        "skipped_files": [{"file": os.path.basename(s["file"]), "error": s["error"]} for s in book.skipped_files],
    }
