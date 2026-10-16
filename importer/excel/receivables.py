"""Read-only reader for the receivables (AR) list of the finance workbooks (folder "09 AR and Financial
Status"). Pure parsing: nothing here knows the CRM. Only ONE kind of sheet is CRM data: the invoice-level
list of open sales invoices (customer, invoice no., invoice date, due date, value in local currency).
From it come two numbers per customer, open and overdue, and the date they are true for.

Everything else in these workbooks is excluded by code, with a reason, and never read into a record:
bank cash flows, accounts payable, fixed costs, customer down payments received in advance (a liability,
not a receivable), aging-bucket ledger snapshots (the legacy Läpple AR/AP files), finance summaries and
mappings, hidden sheets. Only the named columns of the receivables list are read, so a bank account, tax
id or remark that sits beside them never reaches a record."""
import datetime as dt
import os
import re
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal

from .companies import norm as company_norm
from .files import file_display_name
from .orders import GROUP_NAME, label_key, parse_date, parse_number, _unique_files
from .readers import ReviewItem
from .sheet import load

FAMILY = "receivables"
COMPANY_OK = 0.9
SCAN_ROWS = 25
NOT_CRM = "not CRM data: "
BANK = NOT_CRM + "bank cash flows"
PAYABLE = NOT_CRM + "accounts payable"
FIXED = NOT_CRM + "fixed costs"
SUMMARY = NOT_CRM + "finance summary or mapping"
DOWN = "not receivables: customer down payments received in advance (a liability)"
AGING = NOT_CRM + "aging-bucket ledger snapshot (legacy Läpple AR/AP), not an invoice list"
HIDDEN = "hidden sheet"
OTHER = "not an invoice-level receivables list"
# first match wins, on the sheet name and its first rows
_SHEET_KINDS = (
    (re.compile(r"(received|rec\.?) in advance|down ?payment|pre-?paid|预收|预付款", re.I), DOWN),
    (re.compile(r"\bcash\b|\beba\b|\bboc\b|\babc\b|bank|银行", re.I), BANK),
    (re.compile(r"fix(ed)? cost|固定成本", re.I), FIXED),
    (re.compile(r"overview|mapping|offset|summary", re.I), SUMMARY),
    (re.compile(r"\bAP\b|payable|应付|supplier", re.I), PAYABLE),
    (re.compile(r"within payment terms|账期内|days overdue|逾期", re.I), AGING),
)
ANNOTATION = re.compile(r"\s*[（(][^）)]*\btransfer\b[^）)]*[）)]?\s*$", re.I)


@dataclass
class Invoice:
    customer_en: str
    customer_cn: str
    invoice_no: str | None
    invoice_date: str | None
    due_date: str | None
    amount: Decimal  # in the local currency named in the list's header
    row: int
    note: str = ""  # why this row is counted the way it is


@dataclass
class ArSheet:
    """One invoice-level list with the date its overdue figure was measured."""
    file: str
    sheet: str
    currency: str | None  # from the header of the amount column ("Invoice Value Local Currency (RMB)")
    as_of: str | None  # the day "due days" were counted from, else the report's date, else the file name's
    as_of_source: str = ""
    report_date: str | None = None
    invoices: list = field(default_factory=list)
    skipped_rows: int = 0  # "No Invoice" placeholder rows and blanks
    claimed_open: Decimal | None = None  # the totals the sheet states itself
    claimed_overdue: Decimal | None = None
    warnings: list = field(default_factory=list)

    @property
    def open_total(self) -> Decimal:
        return sum((i.amount for i in self.invoices), Decimal(0))

    @property
    def overdue_total(self) -> Decimal:
        return sum((i.amount for i in self.invoices if _overdue(i, self.as_of)), Decimal(0))


@dataclass
class Customer:
    key: str
    name: str  # the name the CRM gets
    written_names: list = field(default_factory=list)
    open: Decimal = Decimal(0)
    overdue: Decimal = Decimal(0)
    invoices: int = 0
    overdue_invoices: int = 0
    oldest_due: str | None = None  # earliest due date among the customer's open invoices (audit aid), None when none has one
    group_company: bool = False
    match: dict = field(default_factory=dict)
    queued: bool = False
    doubts: list = field(default_factory=list)  # why the totals may be wrong; a customer with doubts is held out


@dataclass
class FileInfo:
    path: str
    name: str
    kind: str = ""  # receivables | excluded
    reason: str = ""
    sheets_read: list = field(default_factory=list)
    excluded_sheets: list = field(default_factory=list)  # (sheet name, reason)
    as_of: str | None = None
    invoices: int = 0
    superseded_by: str = ""  # an older snapshot is replaced by a newer file
    warnings: list = field(default_factory=list)


@dataclass
class ReceivablesBook:
    as_of: str | None = None  # the snapshot that counts: the newest
    source: str = ""  # its file
    currency: str | None = None
    customers: dict = field(default_factory=dict)  # key -> Customer
    review: list = field(default_factory=list)
    files: list = field(default_factory=list)
    skipped_files: list = field(default_factory=list)
    sheet: ArSheet | None = None

    def accepted(self) -> list:
        return [c for c in self.customers.values() if not c.queued and not c.doubts]

    def excluded(self) -> list:
        return [f for f in self.files if f.kind == "excluded"]


def _overdue(inv: Invoice, as_of) -> bool:
    return bool(inv.due_date and as_of and inv.due_date < as_of)


def _money(v) -> Decimal | None:
    n = parse_number(v)
    return None if n is None else Decimal(str(n)).quantize(Decimal("0.01"), ROUND_HALF_UP)


# ---------- the list's header ----------

def _field(label) -> str | None:
    t = label_key(label)
    if t.startswith("customer name en"):
        return "customer_en"
    if t in ("customer name", "客户名称", "客户"):
        return "customer_cn"
    if t.startswith(("sales invoice no", "invoice no", "invoice number")) or t == "发票号":
        return "invoice_no"
    if t in ("invoice date", "发票日期"):
        return "invoice_date"
    if t in ("due date", "到期日"):
        return "due_date"
    if t.startswith(("invoice value local currency", "本币金额")):
        return "amount"
    return None


def find_invoice_header(sheet):
    """(row, {field: col}) of the invoice list's header, else None. It needs the invoice number, the due
    date, the local-currency amount and a customer name: aging tables and ledgers have none of that."""
    for r in range(1, min(SCAN_ROWS, sheet.max_row) + 1):
        cols = {}
        for c in range(1, sheet.max_col + 1):
            v = sheet.value(r, c)
            f = _field(v) if isinstance(v, str) and not sheet.is_merged_continuation(r, c) else None
            if f and f not in cols:
                cols[f] = c
        if {"invoice_no", "due_date", "amount"} <= set(cols) and ({"customer_en", "customer_cn"} & set(cols)):
            return r, cols
    return None


def looks_like_receivables(sheet) -> bool:
    return find_invoice_header(sheet) is not None


def file_kind(path) -> tuple[str, str]:
    """("receivables", "") when the workbook holds an invoice-level receivables list, else
    ("excluded", why): bank cash flows, accounts payable, fixed costs, down payments, aging ledgers ..."""
    info = FileInfo(path=path, name=file_display_name(path))
    if read_workbook(path, info):
        return "receivables", ""
    return "excluded", "; ".join(dict.fromkeys(r for _, r in info.excluded_sheets if r != HIDDEN)) or OTHER


def _currency_of(label) -> str | None:
    t = str(label)
    return "CNY" if re.search(r"rmb|cny|人民币", t, re.I) else None


def _date_near(sheet, header_row, words):
    """The date beside a label (a cell holding one of `words`) in the rows above the header."""
    for r in range(1, header_row):
        cells = [(c, sheet.value(r, c)) for c in range(1, sheet.max_col + 1)]
        if any(isinstance(v, str) and any(w in label_key(v) for w in words) for _, v in cells):
            for _, v in cells:
                if isinstance(v, (dt.datetime, dt.date)):
                    return parse_date(v)
    return None


def _claimed(sheet, header_row, words):
    """The figure the sheet itself states under a label like "Total Account Receivables"."""
    for r in range(1, header_row):
        for c in range(1, sheet.max_col + 1):
            v = sheet.value(r, c)
            if isinstance(v, str) and any(w in label_key(v) for w in words):
                for r2 in range(r + 1, min(r + 3, header_row)):
                    n = _money(sheet.value(r2, c))
                    if n is not None:
                        return n
    return None


def _report_date(sheets):
    """The date typed on the workbook's Overview sheet ("Date:" beside a date)."""
    for sh in sheets:
        if label_key(sh.name) != "overview":
            continue
        for r in range(1, min(10, sh.max_row) + 1):
            for c in range(1, sh.max_col):
                v = sh.value(r, c)
                if isinstance(v, str) and label_key(v) == "date":
                    got = parse_date(sh.value(r, c + 1))
                    if got:
                        return got
    return None


def _name_date(path):
    m = re.search(r"(?<!\d)(20\d{2})(\d{2})(\d{2})(?!\d)", os.path.basename(path))
    if not m:
        return None
    try:
        return dt.date(int(m[1]), int(m[2]), int(m[3])).isoformat()
    except ValueError:
        return None


def _clean_en(name) -> str:
    return ANNOTATION.sub("", name or "").strip()


def read_ar_sheet(sheet, path, report_date) -> ArSheet:
    hrow, cols = find_invoice_header(sheet)
    amount_label = sheet.value(hrow, cols["amount"])
    today = _date_near(sheet, hrow, ("today", "今天"))
    as_of, source = (today, "the list's own 'today' cell") if today else \
        (report_date, "the workbook's Overview date") if report_date else \
        (_name_date(path), "the date in the file name")
    out = ArSheet(file=path, sheet=sheet.name, currency=_currency_of(amount_label), as_of=as_of,
                  as_of_source=source if as_of else "", report_date=report_date,
                  claimed_open=_claimed(sheet, hrow, ("total account receivable",)),
                  claimed_overdue=_claimed(sheet, hrow, ("total amount overdue",)))
    if sheet.cut:
        out.warnings.append(sheet.cut)

    def cell(r, f):
        return sheet.value(r, cols[f]) if f in cols else None

    def text(v):
        return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""

    for r in range(hrow + 1, sheet.max_row + 1):
        amount = _money(cell(r, "amount"))
        en, cn, no = text(cell(r, "customer_en")), text(cell(r, "customer_cn")), text(cell(r, "invoice_no"))
        if no.startswith("#") or en.startswith("#") or cn.startswith("#"):
            out.skipped_rows += 1  # a formula error or "No Invoice" placeholder
            continue
        if amount is None or not (en or cn):
            out.skipped_rows += 1
            continue
        if amount == 0:
            out.skipped_rows += 1  # nothing open
            continue
        due, issued = parse_date(cell(r, "due_date")), parse_date(cell(r, "invoice_date"))
        note = ""
        if not no:
            note = "no invoice number or due date: counted as open, never as overdue"
        elif due is None:
            note = "no due date: counted as open, never as overdue"
        out.invoices.append(Invoice(_clean_en(en), cn, no or None, issued, due, amount, r, note))
    for label, got, claimed in (("open", out.open_total, out.claimed_open),
                                ("overdue", out.overdue_total, out.claimed_overdue)):
        if claimed is not None and got != claimed:
            out.warnings.append(f"{label} total {got} differs from the {claimed} the sheet states itself")
    if out.currency is None:
        out.warnings.append("the amount column does not name its currency")
    if out.as_of is None:
        out.warnings.append("no date for the figures (no 'today' cell, Overview date or date in the file name)")
    return out


# ---------- what the other sheets are ----------

def sheet_reason(sheet, hidden: bool) -> str:
    if hidden:
        return HIDDEN
    head = " ".join(str(sheet.value(r, c)) for r in range(1, min(4, sheet.max_row) + 1)
                    for c in range(1, min(sheet.max_col, 14) + 1) if isinstance(sheet.value(r, c), str))
    for text in (sheet.name, head):  # the name is the better evidence, so it is tried alone first
        for pattern, reason in _SHEET_KINDS:
            if pattern.search(text):
                return reason
    return OTHER


def read_workbook(path, info: FileInfo):
    """The ArSheet of every visible invoice list in the workbook; every other sheet is named with its reason."""
    sheets = load(path)
    report = _report_date(sheets)
    found = []
    for sh in sheets:
        state = getattr(sh.ws, "sheet_state", "visible") if sh.ws is not None else "visible"
        if state == "visible" and find_invoice_header(sh):
            found.append(read_ar_sheet(sh, path, report))
            info.sheets_read.append(sh.name)
        else:
            info.excluded_sheets.append((sh.name, sheet_reason(sh, state != "visible")))
    return found


# ---------- the book ----------

def _match_name(master, en, cn):
    """The master's answer for a customer written in two languages: the Chinese name first."""
    results = [master.match(n) for n in (cn, en) if n]
    for m in results:
        if m.score >= COMPANY_OK and (m.cn or m.en):
            return m
    for m in results:
        if m.method == "new":
            return m
    return max(results, key=lambda m: m.score)


def build_book(paths, master=None) -> ReceivablesBook:
    """The newest receivables snapshot among `paths`, as open and overdue per customer. Older snapshots,
    byte-identical copies, and every sheet that is not an invoice list are named and left out."""
    if master is None:
        from .run import IMPORTER_DIR, _master
        master = _master(IMPORTER_DIR)
    book = ReceivablesBook()
    lists = []
    for path in _unique_files(paths, book.skipped_files):
        info = FileInfo(path=path, name=file_display_name(path))
        book.files.append(info)
        try:
            found = read_workbook(path, info)
        except Exception as exc:  # corrupt, protected or unsupported: report, go on
            book.skipped_files.append({"file": path, "error": f"could not be read: {type(exc).__name__}: {exc}"})
            info.kind, info.reason = "excluded", "could not be read"
            continue
        reasons = [r for _, r in info.excluded_sheets if r != HIDDEN]
        if not found:
            info.kind = "excluded"
            info.reason = "; ".join(dict.fromkeys(reasons)) or OTHER
            continue
        info.kind = "receivables"
        for sh in found:
            info.warnings += sh.warnings
            info.as_of = max(filter(None, [info.as_of, sh.as_of]), default=None)
            info.invoices += len(sh.invoices)
            lists.append((sh, os.path.getmtime(path), info))
    usable = [t for t in lists if t[0].as_of and t[0].currency]
    for sh, _, info in lists:
        if not (sh.as_of and sh.currency):
            book.review.append(ReviewItem(FAMILY, info.name, sh.sheet, 0, "as_of" if not sh.as_of else "currency",
                                          "", "undated_or_unlabelled_snapshot"))
    if not usable:
        return book
    usable.sort(key=lambda t: (t[0].as_of, len(t[0].invoices), t[1], t[2].name), reverse=True)
    chosen, _, chosen_info = usable[0]
    for sh, _, info in usable[1:]:
        info.superseded_by = chosen_info.name
    book.as_of, book.source, book.currency, book.sheet = chosen.as_of, chosen_info.name, chosen.currency, chosen
    _customers(book, chosen, master)
    return book


def _customers(book, sheet, master):
    matches, by_key, seen_numbers = {}, book.customers, {}
    for inv in sheet.invoices:
        ident = (inv.customer_en, inv.customer_cn)
        if ident not in matches:
            m = matches[ident] = _match_name(master, inv.customer_en, inv.customer_cn)
            if m.score < COMPANY_OK:
                book.review.append(ReviewItem(FAMILY, book.source, sheet.sheet, inv.row, "customer",
                                              inv.customer_cn or inv.customer_en, "company",
                                              [list(c) for c in (m.candidates or [])]))
        m = matches[ident]
        written = inv.customer_cn or inv.customer_en
        known = m.score >= COMPANY_OK and (m.cn or m.en)
        name = (m.cn or m.en) if known else written
        key = company_norm(name) or name.lower()
        cust = by_key.get(key)
        if cust is None:
            cust = by_key[key] = Customer(key=key, name=name, queued=m.score < COMPANY_OK,
                                          match={"code": m.code, "cn": m.cn, "en": m.en, "score": m.score,
                                                 "method": m.method},
                                          group_company=bool(GROUP_NAME.search(f"{inv.customer_en} {inv.customer_cn}")))
            if cust.group_company:
                book.review.append(ReviewItem(FAMILY, book.source, sheet.sheet, inv.row, "customer", written,
                                              "group_company"))
        for n in (inv.customer_cn, inv.customer_en):
            if n and n not in cust.written_names:
                cust.written_names.append(n)
        numbers = seen_numbers.setdefault(key, set())
        for doubt in (["invoice without a number"] if not inv.invoice_no else
                      ["invoice number %s twice" % inv.invoice_no] if inv.invoice_no in numbers else []):
            cust.doubts.append(doubt)
            book.review.append(ReviewItem(FAMILY, book.source, sheet.sheet, inv.row, "invoice", written, "invoice_doubtful"))
        numbers.add(inv.invoice_no)
        cust.open += inv.amount
        cust.invoices += 1
        if inv.due_date and (cust.oldest_due is None or inv.due_date < cust.oldest_due):
            cust.oldest_due = inv.due_date
        if _overdue(inv, sheet.as_of):
            cust.overdue += inv.amount
            cust.overdue_invoices += 1
        if inv.note:
            book.review.append(ReviewItem(FAMILY, book.source, sheet.sheet, inv.row, "invoice", written, "no_due_date"))


def _by_reason(sheets) -> dict:
    out = {}
    for _, reason in sheets:
        out[reason] = out.get(reason, 0) + 1
    return out


def summarize(book: ReceivablesBook) -> dict:
    reasons = {}
    for item in book.review:
        reasons[item.reason] = reasons.get(item.reason, 0) + 1
    accepted = book.accepted()
    return {
        "files": len(book.files), "files_with_receivables": sum(f.kind == "receivables" for f in book.files),
        "files_excluded": len(book.excluded()), "snapshot": book.source, "as_of": book.as_of,
        "currency": book.currency, "invoices": len(book.sheet.invoices) if book.sheet else 0,
        "customers": len(book.customers), "customers_accepted": len(accepted),
        "customers_queued": len(book.customers) - len(accepted),
        "open_total": str(sum((c.open for c in book.customers.values()), Decimal(0))),
        "overdue_total": str(sum((c.overdue for c in book.customers.values()), Decimal(0))),
        "superseded": [{"file": f.name, "by": f.superseded_by} for f in book.files if f.superseded_by],
        "excluded": [{"file": f.name, "reason": f.reason, "sheets": len(f.excluded_sheets)} for f in book.excluded()],
        "excluded_sheets": {f.name: _by_reason(f.excluded_sheets) for f in book.files},
        "review_by_reason": reasons,
        "skipped_files": [{"file": os.path.basename(s["file"]), "error": s["error"]} for s in book.skipped_files],
    }
