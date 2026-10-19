"""Reader for the Quote Log workbook (one row per quote revision). Pure parsing: nothing here knows
about the CRM. Rows that share a quote number are the revisions of one enquiry (one deal); a row with
no usable quote number gets a stable key from its row so a re-run finds it again. Money is read
as written: the log names no currency, so none is invented here (the adapter asks)."""
import datetime as dt
import hashlib
import os
import re
from collections import Counter
from dataclasses import dataclass, field

from .companies import norm as company_norm
from .files import file_display_name
from .orders import COMPANY_OK, Org, label_key, parse_date, parse_number
from .readers import ReviewItem
from .sheet import load
from .split import cell_text

FAMILY = "quotes"
HEADER_ROW = 1
ERROR_CELLS = ("#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?")
QUOTE_NO = re.compile(r"^QU\d{7,9}$")
LAST_YEAR = dt.date.today().year + 1

# field -> exact label_key of the header cell
COLUMNS = {
    "quote_no": "quote", "product_group": "product group", "article": "article", "product_type": "product type",
    "company": "company", "end_customer": "end customer", "contact": "contact", "clerk": "clerk",
    "responsible": "responsible", "representative": "representative", "rfq_date": "rfq date",
    "quote_date": "quote date", "revision": "quote id", "quantity": "qty", "est_value": "estimated quote value",
    "est_cost": "est manufacturing costs", "price": "sales price", "margin": "db ii", "total": "total sales volume",
    "lead_time": "lead time", "order_date": "order date", "erp_order": "erp order", "status": "status sales chance",
    "lost_index": "lost index", "remark": "follow up remark",
}
REQUIRED = ("quote_no", "company", "status", "price")
WON = ("won order", "order won", "ordered ")
OPEN = ("open sales chance", "budget offer", "customer has not received", "in process", "in progress", "on hold",
        "paused", "plan project")
LOST = ("cancel", "lost", "outbid", "no technical", "no economic", "no feas", "no replacement", "no feedback",
        "merged", "didn't get", "rejected", "declined", "no demand", "no budget")


def status_bucket(text: str) -> str:
    """won | open | lost | unknown, from the sales-chance wording. Unknown stays open and is flagged."""
    t = (text or "").lower()
    if not t:
        return "unknown"
    for bucket, words in (("won", WON), ("lost", LOST), ("open", OPEN)):
        if any(w in t for w in words):
            return bucket
    return "unknown"


@dataclass
class Quote:
    """One row: a quote number at one revision."""
    deal_key: str  # the quote number, or "row<N>" when the number is empty, a placeholder or reused
    revision: int
    quote_no: str
    row: int
    sheet: str
    file: str
    customer: str
    company_key: str = ""
    product_group: str = ""
    article: str = ""
    product_type: str = ""
    end_customer: str = ""
    contact: str = ""
    clerk: str = ""
    responsible: str = ""
    representative: str = ""
    rfq_date: str | None = None
    quote_date: str | None = None
    order_date: str | None = None
    quantity: float | None = None
    price: float | None = None  # unit sales price as written
    total: float | None = None  # price x quantity as written
    est_cost: float | None = None
    margin: float | None = None
    lead_time: str = ""
    erp_order: str = ""
    status: str = ""
    bucket: str = "unknown"
    lost_index: str = ""
    remark: str = ""
    issues: list = field(default_factory=list)

    @property
    def amount(self) -> float | None:
        """What the offer is worth: the total volume the log states, else price x quantity, else the price."""
        if self.total:
            return self.total
        if self.price:
            return round(self.price * (self.quantity or 1), 2)
        return None

    @property
    def source_id(self) -> str:
        return f"quote:{self.deal_key}"


@dataclass
class QuotesBook:
    quotes: list = field(default_factory=list)
    orgs: dict = field(default_factory=dict)  # company key -> Org
    review: list = field(default_factory=list)
    files: list = field(default_factory=list)  # [{"path", "name", "rows"}]
    skipped_files: list = field(default_factory=list)

    def accepted(self) -> list:
        """Quotes that can be written: their company is known or answered."""
        return [q for q in self.quotes if q.company_key and not self.orgs[q.company_key].queued]

    def accepted_orgs(self) -> dict:
        used = {q.company_key for q in self.accepted()}
        return {k: g for k, g in self.orgs.items() if k in used}

    def deals(self) -> dict:
        """deal key -> its accepted quotes, revisions in order."""
        out = {}
        for q in sorted(self.accepted(), key=lambda q: (q.deal_key, q.revision, q.row)):
            out.setdefault(q.deal_key, []).append(q)
        return out


def header_map(sheet) -> dict | None:
    """{field: column} when row 1 carries the quote log's columns, else None."""
    labels = {label_key(sheet.value(HEADER_ROW, c)): c for c in range(1, sheet.max_col + 1)
              if sheet.value(HEADER_ROW, c) is not None}
    found = {f: labels[k] for f, k in COLUMNS.items() if k in labels}
    return found if all(f in found for f in REQUIRED) and "product group" in labels else None


def looks_like_quotes(sheet) -> bool:
    return header_map(sheet) is not None


def _text(v) -> str:
    t = cell_text(v) if v is not None else ""
    return "" if t in ERROR_CELLS else t


def _date(v, issues, label):
    if v is None or _text(v) == "":
        return None
    iso = parse_date(v)
    if iso is None:
        m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", _text(v))  # the 2026 rows are typed dd/mm/yyyy
        if m:
            try:
                iso = dt.date(int(m[3]), int(m[2]), int(m[1])).isoformat()
            except ValueError:
                pass
    if iso is None:
        issues.append(f"{label} is not a date: {_text(v)!r}")
    elif not 2015 <= int(iso[:4]) <= LAST_YEAR:
        issues.append(f"{label} {iso} is outside the plausible years")
        return None
    return iso


def _money(v, issues, label):
    if v is None or _text(v) == "":
        return None
    n = parse_number(v)
    if n is None:
        issues.append(f"{label} is not an amount: {_text(v)!r}")
        return None
    return round(n, 2)


def _revision(v) -> int:
    n = parse_number(v)
    return int(n) if n and n >= 1 else 1


def read_sheet(sheet, path, cols, review) -> list:
    name = file_display_name(path)

    def get(r, f):
        return sheet.value(r, cols[f]) if f in cols else None

    last = max((r for (r, _c) in sheet.grid), default=HEADER_ROW)
    out = []
    for r in range(HEADER_ROW + 1, last + 1):
        customer = _text(get(r, "company"))
        if not customer and not _text(get(r, "quote_no")):
            continue
        issues = []
        raw_no = _text(get(r, "quote_no")).replace(" ", "")
        key = raw_no if QUOTE_NO.match(raw_no) else f"row{r}"
        if key != raw_no:
            review.append(ReviewItem(FAMILY, name, sheet.name, r, "quote_no", raw_no, "quote_number_unusable"))
        q = Quote(deal_key=key, revision=_revision(get(r, "revision")), quote_no=raw_no, row=r, sheet=sheet.name,
                  file=name, customer=customer.strip())
        for f in ("product_group", "article", "product_type", "end_customer", "contact", "clerk", "responsible",
                  "representative", "lead_time", "erp_order", "status", "lost_index"):
            setattr(q, f, _text(get(r, f)))
        if isinstance(get(r, "product_type"), dt.datetime):  # Excel turned a code like "1.10" into a date
            q.product_type = ""
            issues.append("product type was turned into a date by Excel")
        q.rfq_date = _date(get(r, "rfq_date"), issues, "RFQ date")
        q.quote_date = _date(get(r, "quote_date"), issues, "quote date")
        q.order_date = _date(get(r, "order_date"), issues, "order date")
        q.quantity = parse_number(get(r, "quantity"))
        q.price = _money(get(r, "price"), issues, "sales price")
        q.total = _money(get(r, "total"), issues, "total sales volume")
        q.est_cost = _money(get(r, "est_cost"), issues, "manufacturing cost")
        q.margin = parse_number(get(r, "margin"))
        q.remark = _text(get(r, "remark"))
        q.bucket = status_bucket(q.status)
        if q.bucket == "unknown":
            issues.append(f"sales-chance status not understood: {q.status!r}" if q.status else "no sales-chance status")
        if not q.customer:
            issues.append("no company")
        q.issues = issues
        out.append(q)
    return out


def _fix_reused_numbers(quotes, review):
    """The same quote number and revision twice is two enquiries sharing a number: the later row is
    kept apart under its own row key."""
    seen = set()
    for q in quotes:
        ident = (q.deal_key, q.revision)
        if ident in seen:
            review.append(ReviewItem(FAMILY, q.file, q.sheet, q.row, "quote_no", q.quote_no, "quote_number_reused"))
            q.deal_key = f"row{q.row}"
            q.revision = 1
        seen.add(ident)


def _match_orgs(quotes, master, review) -> dict:
    matches, orgs = {}, {}
    for q in quotes:
        if not q.customer:
            continue
        if q.customer not in matches:
            m = matches[q.customer] = master.match(q.customer)
            if m.score < COMPANY_OK:
                review.append(ReviewItem(FAMILY, q.file, q.sheet, q.row, "customer", q.customer, "company",
                                         [list(c) for c in (m.candidates or [])]))
        m = matches[q.customer]
        known = m.score >= COMPANY_OK and (m.cn or m.en)
        name = (m.cn or m.en) if known else q.customer
        q.company_key = company_norm(name) or name.lower()
        org = orgs.get(q.company_key)
        if org is None:
            org = orgs[q.company_key] = Org(q.company_key, name)
            org.match = {"code": m.code, "cn": m.cn, "en": m.en, "score": m.score, "method": m.method}
            org.queued = m.score < COMPANY_OK
        if q.customer not in org.written_names:
            org.written_names.append(q.customer)
    return orgs


def read_paths(paths, master=None) -> QuotesBook:
    """Every quote log among `paths`. A bad file is recorded and the run goes on; identical copies
    are read once."""
    if master is None:
        from .run import IMPORTER_DIR, _master
        master = _master(IMPORTER_DIR)
    book, digests = QuotesBook(), set()
    for path in sorted(os.fspath(p) for p in paths):
        try:
            with open(path, "rb") as f:
                digest = hashlib.sha256(f.read()).hexdigest()
            if digest in digests:
                book.skipped_files.append({"file": path, "error": "identical to another file in this upload"})
                continue
            digests.add(digest)
            sheets = load(path)
        except Exception as exc:
            book.skipped_files.append({"file": path, "error": f"could not be read: {type(exc).__name__}: {exc}"})
            continue
        rows = 0
        for sheet in sheets:
            cols = header_map(sheet)
            if cols is None:
                continue
            got = read_sheet(sheet, path, cols, book.review)
            rows += len(got)
            book.quotes += got
        book.files.append({"path": path, "name": file_display_name(path), "rows": rows})
        if not rows:
            book.skipped_files.append({"file": path, "error": "no sheet has the quote log's columns"})
    _fix_reused_numbers(book.quotes, book.review)
    book.orgs = _match_orgs(book.quotes, master, book.review)
    book.review += [ReviewItem(FAMILY, q.file, q.sheet, q.row, "quote", q.quote_no or f"row {q.row}", "quote_doubtful")
                    for q in book.quotes if q.issues]
    return book


def summarize(book: QuotesBook) -> dict:
    accepted, deals = book.accepted(), book.deals()
    buckets = Counter(q.bucket for q in accepted)
    reasons = Counter(i.reason for i in book.review)
    return {"quotes": len(book.quotes), "quotes_accepted": len(accepted), "quotes_queued": len(book.quotes) - len(accepted),
            "deals": len(deals), "companies": len(book.orgs), "companies_accepted": len(book.accepted_orgs()),
            "companies_queued": sum(o.queued for o in book.orgs.values()),
            "won": buckets["won"], "open": buckets["open"], "lost": buckets["lost"], "unknown": buckets["unknown"],
            "without_price": sum(q.amount is None for q in accepted), "review_by_reason": dict(reasons)}
