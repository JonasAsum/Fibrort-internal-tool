"""Read-only reader for contact-list workbooks (folder "06 Contact Lists"). Pure parsing: nothing
here knows the CRM. Two layouts, found by header label (English or Chinese, any column order,
header row anywhere in the first rows):

  list      a company column (公司名称 / Company ...): one row per company and contact; unlabelled
            type columns (应收/应付 and 客户/供应商) are recognised by their values.
  sections  no company column: column A holds a team or supplier label that carries down to the
            rows below it, and rows with only that label are section headings. Under a heading that
            names the group (FIBRO) the people are internal colleagues of that company; under any
            other heading the label is a supplier or service provider.

Rows that are not contacts are skipped with a reason, never silently. A bad file or sheet is
recorded and the run goes on. Company names go through the same master match as the visit
importer (excel.companies.Master); a name it does not know well enough is queued ("company")
and its contacts stay out until a person answers."""
import os
import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field

from visits.match import Contact, _person_key
from visits.pipeline import contact_emails, normalize_phone

from .companies import norm
from .headers import field_for as visit_field
from .headers import label_field
from .readers import ReviewItem
from .files import unique_contents
from .sheet import load
from .split import cell_text, phone_numbers

FAMILY = "contacts"
SCAN_ROWS = 15
COMPANY_OK = 0.9
BLOCKING = {"company"}  # review reasons that keep a company and its contacts out of the commit
UNRECOGNISED_MIN_CELLS = 10
GROUP = "FIBRO"  # a heading that names it marks the people below as internal colleagues
GROUP_FALLBACK = "FIBRO GmbH"
INTERNAL_SUFFIX = " (internal)"  # the CRM account of the group's own colleagues is named so
# Where a person came from, as the CRM shows it (a picklist, so the owner can filter or delete by it).
ORIGINS = {("list", "customer"): "Finance list: customer", ("list", "supplier"): "Finance list: supplier",
           ("list", ""): "Finance list", ("sections", "internal"): "Handover list: internal colleague",
           ("sections", "supplier"): "Handover list: supplier", ("sections", ""): "Handover list"}

# field -> labels; a leading "=" means the whole label must equal it, otherwise it may start with it.
SYNONYMS = {
    "ledger": ("应收/应付",),
    "kind": ("客户/供应商",),
    "seq": ("=序号", "=no", "=no.", "=#", "=nr"),
    "company": ("公司名称", "company", "客户名称", "供应商名称", "单位名称", "=customer", "=supplier", "=公司"),
    "address": ("地址", "address", "anschrift"),
    "contact": ("contact person", "contact name", "contact", "联系人", "姓名", "ansprechpartner", "=name"),
    "email": ("e-mail", "email", "邮箱", "电子邮件", "邮件"),
    "phone": ("phone", "tel", "mobile", "电话", "手机", "telefon"),
    "position": ("position", "job title", "title", "role", "职位", "职务", "funktion"),
    "remark": ("remark", "note", "comment", "备注", "bemerkung"),
}
LEDGER = {"应收": "customer", "预收": "customer", "其他应收": "customer",
          "应付": "supplier", "预付": "supplier", "其他应付": "supplier"}
KIND_WORDS = {"客户": "customer", "供应商": "supplier", "customer": "customer", "supplier": "supplier",
              "vendor": "supplier"}
CJK = re.compile(r"[一-鿿]")
PAREN = re.compile(r"\(([^()一-鿿]*)\)")  # a bracket in Latin only; one inside a Chinese name is part of it
LEGAL_LATIN = re.compile(r"\b(?:gmbh|ag|ltd|llc|inc|co\.?|corp|limited)\b", re.I)
# A bracket after a supplier label either says what the supplier does for us ("DHL Express (Urgent
# Parts)") or names the supplier of a part ("Motors (SEW China)"): these words mark the first kind.
SERVICE_WORDS = re.compile(
    r"\b(?:orders?|parts?|docs?|files?|urgent|exhibit\w*|trips?|clearance|customs|transport\w*|service\w*|"
    r"dealer|normal|delivery|import|export|shipping|forwarder\w*)\b", re.I)
# Bank and tax details never reach the CRM: no column for them is read, and free text that holds some is withheld.
SENSITIVE = re.compile(r"税号|纳税|开户|银行[账帐]|[账帐]号|\b(?:iban|swift|bic|vat|ust-?id|tax)\b|bank\s*(?:acc|detail)", re.I)
CREDIT_CODE = re.compile(r"\b(?=[0-9A-Z]*\d)[0-9A-Z]{18}\b")  # a Chinese unified social credit code (tax ID)
FAX = re.compile(r"^\s*(?:fax|f)\b\s*[.:]?\s*(?=[+\d(])", re.I)
PHONE_LABEL = re.compile(r"^\s*(?:tel|telefon|phone|mobile|mob|t|m|p)\b\s*[.:]?\s*(?=[+\d(])", re.I)


@dataclass
class Entry:
    """One company, or one person at a company, as read from one row (then merged across copies)."""
    company: str
    full_name: str = ""
    emails: list = field(default_factory=list)
    phones: list = field(default_factory=list)  # E.164 numbers
    role: str = ""
    remark: str = ""
    department: str = ""  # the team label of an internal colleague, or what a supplier supplies ("Motors (SEW China)")
    section: str = ""  # the heading above the row
    kind: str = ""  # customer | supplier | internal | ""
    ledger: str = ""  # 应收 / 应付 ... as written
    address: str = ""
    file: str = ""
    sheet: str = ""
    row: int = 0
    layout: str = ""
    company_key: str = ""
    source_id: str = ""
    copies: int = 1  # rows (in any file) that said the same person

    @property
    def origin(self) -> str:
        return ORIGINS.get((self.layout, self.kind), "")

    def richness(self) -> int:
        return sum(bool(v) for v in (self.full_name, self.emails, self.phones, self.role, self.remark,
                                     self.department, self.address, self.ledger))


@dataclass
class Org:
    key: str
    name: str  # the name the CRM gets: the master's when it knows the company, else as written
    kind: str = ""  # the first kind seen; kinds holds all of them (a company can be both customer and supplier)
    kinds: list = field(default_factory=list)
    ledgers: list = field(default_factory=list)  # 应收 / 应付 ... as written
    origin: str = ""
    address: str = ""
    location: str = ""  # the remark of a row with no person ("Shanghai", "Germany"): the only trace of where the supplier is
    written_names: list = field(default_factory=list)
    match: dict = field(default_factory=dict)  # code, cn, en, score, method
    queued: bool = False
    source_id: str = ""


@dataclass
class ContactsFileStat:
    path: str
    layout: str = "none"
    sheets: int = 0
    read_sheets: int = 0
    rows: int = 0  # data rows (a company or a person)
    blank_rows: int = 0
    error: str = ""
    unrecognised: list = field(default_factory=list)
    sheet_errors: list = field(default_factory=list)


@dataclass
class ContactsResult:
    contacts: list  # merged people (Entry with a full_name)
    orgs: dict  # company key -> Org
    review: list  # ReviewItem: skipped rows, rows imported with a gap, companies to answer
    skipped_files: list  # [{"file", "error"}]
    per_file_stats: list = field(default_factory=list)

    def accepted_contacts(self) -> list:
        return [c for c in self.contacts if not self.orgs[c.company_key].queued]

    def accepted_orgs(self) -> dict:
        return {k: o for k, o in self.orgs.items() if not o.queued}


def _present(value) -> bool:
    return value is not None and str(value).strip() != ""


def find_contacts_header(sheet):
    """(row, {field: column}) of the first row that labels a contact or company column and at least
    two more fields; a visit list's header (it has a date) never counts."""
    for r in range(1, min(SCAN_ROWS, sheet.max_row) + 1):
        cols, visit = {}, False
        for c in range(1, sheet.max_col + 1):
            value = sheet.value(r, c)
            visit = visit or visit_field(value) in ("date", "cw")
            name = label_field(value, SYNONYMS)
            if name and name not in cols:
                cols[name] = c
        if not visit and len(cols) >= 3 and ("contact" in cols or "company" in cols):
            return r, cols
    return None


def detect(sheet) -> str | None:
    found = find_contacts_header(sheet)
    return None if found is None else ("list" if "company" in found[1] else "sections")


def _sniff_type_columns(sheet, header_row, cols):
    """Unlabelled columns whose every value is a ledger word (应收 ...) or a kind word (客户 / 供应商)."""
    for c in range(1, sheet.max_col + 1):
        if c in cols.values():
            continue
        values = {cell_text(sheet.value(r, c)).casefold() for r in range(header_row + 1, sheet.max_row + 1)} - {""}
        if not values:
            continue
        if "ledger" not in cols and values <= set(LEDGER):
            cols["ledger"] = c
        elif "kind" not in cols and values <= set(KIND_WORDS):
            cols["kind"] = c


def _clean_phone_cell(cell):
    """The cell without a leading "T" / "Tel." label; None for a fax number."""
    text = cell_text(cell)
    return None if FAX.match(text) else PHONE_LABEL.sub("", text, count=1)


def _reach(file, sheet, r, phone_cell, email_cell, internal, items):
    """(E.164 numbers, addresses) of a row. What cannot be placed becomes a review item and stays
    out of the record: a local number without area code is not guessed at."""
    phones, emails = [], []
    cleaned = _clean_phone_cell(phone_cell)
    for raw in [n for part in re.split(r"[/,;，；、\n]", cleaned or "") if part.strip() for n in phone_numbers(part)]:
        number = normalize_phone(raw)
        # normalize_phone reads a number with a leading 0 as a China number; the group's own
        # (German) numbers are only taken when they carry their country prefix.
        if not number.startswith("+") or (internal and not raw.lstrip().startswith("+")):
            items.append(ReviewItem(FAMILY, file, sheet, r, "phone", raw, "phone_no_country"))
        elif number not in phones:
            phones.append(number)
    text = cell_text(email_cell)
    if text:
        emails = [e["email"] for e in contact_emails(Contact(key="", company_key="", full_name="", emails=[text]))]
        if not emails:
            items.append(ReviewItem(FAMILY, file, sheet, r, "email", text, "bad_email"))
    return phones, emails


def company_from_label(label: str) -> str:
    """The company a supplier label names. A registered Chinese name wins ("Chenliang 太仓…有限公司" ->
    "太仓…有限公司"); else a bracket naming a legal form, or a bracket that does not describe a
    service, is the company ("Motors (SEW China)" -> "SEW China"); else the text before the bracket."""
    text = unicodedata.normalize("NFKC", label).strip().rstrip("?？ ")
    brackets = PAREN.findall(text)
    head = " ".join(PAREN.sub(" ", text).split())
    if CJK.search(head):
        return head[CJK.search(head).start():] if "公司" in head else head
    paren = next((b.strip() for b in brackets if b.strip()), "")
    if not paren or LEGAL_LATIN.search(head) and not LEGAL_LATIN.search(paren):
        return head or paren
    if not head or LEGAL_LATIN.search(paren) or not SERVICE_WORDS.search(paren):
        return paren
    return head


def group_name(section: str) -> str:
    """"Order Management _ FIBRO Rundtische GmbH" -> "FIBRO Rundtische GmbH"."""
    m = re.search(rf"(?<![A-Za-z]){GROUP}\b[^_/()]*", section, re.I)
    return " ".join(m.group(0).split()) if m else GROUP_FALLBACK


def _free(ctx, sheet, r, label, value, items) -> str:
    """A free-text cell (position, remark, address); "" and a review item when it holds bank or tax details."""
    text = cell_text(value)
    if text and (SENSITIVE.search(text) or CREDIT_CODE.search(text)):
        items.append(ReviewItem(FAMILY, ctx["file"], sheet.name, r, label, "", "withheld_sensitive"))
        return ""
    return text


def _entry(ctx, sheet, r, name, phone_cell, email_cell, internal, items, **kw) -> Entry:
    phones, emails = _reach(ctx["file"], sheet.name, r, phone_cell, email_cell, internal, items)
    e = Entry(full_name=name, phones=phones, emails=emails, file=ctx["file"], sheet=sheet.name, row=r,
              layout=ctx["layout"], **kw)
    if not name and (phones or emails):  # the CRM needs a name; the reachability is reported, not kept
        items.append(ReviewItem(FAMILY, e.file, e.sheet, r, "contact", e.company, "no_contact_name"))
        e.phones, e.emails = [], []
    return e


def _read_list(sheet, header_row, cols, ctx, items, stat):
    out = []
    _sniff_type_columns(sheet, header_row, cols)
    for r in range(header_row + 1, sheet.max_row + 1):
        cells = {f: sheet.value(r, c) for f, c in cols.items() if f != "seq"}
        if not any(_present(v) for v in cells.values()):
            stat.blank_rows += 1
            continue
        company = cell_text(cells.get("company"))
        if not company:
            items.append(ReviewItem(FAMILY, ctx["file"], sheet.name, r, "company", "", "no_company"))
            continue
        ledger = cell_text(cells.get("ledger"))
        kind = KIND_WORDS.get(cell_text(cells.get("kind")).casefold()) or LEDGER.get(ledger, "")
        out.append(_entry(ctx, sheet, r, cell_text(cells.get("contact")), cells.get("phone"), cells.get("email"), False,
                          items, company=company, role=_free(ctx, sheet, r, "position", cells.get("position"), items),
                          remark=_free(ctx, sheet, r, "remark", cells.get("remark"), items), kind=kind,
                          ledger=ledger, address=_free(ctx, sheet, r, "address", cells.get("address"), items)))
    return out


def _read_sections(sheet, header_row, cols, ctx, items, stat):
    out = []
    label_col = 1 if 1 not in cols.values() else None
    section = label = ""
    for r in range(header_row + 1, sheet.max_row + 1):
        a = cell_text(sheet.value(r, label_col)) if label_col else ""
        cells = {f: sheet.value(r, c) for f, c in cols.items()}
        has_data = any(_present(v) for v in cells.values())
        if not a and not has_data:
            stat.blank_rows += 1
            label = ""  # a blank row ends the team a label was carried down for
            continue
        if a and not has_data:
            section, label = a, ""
            items.append(ReviewItem(FAMILY, ctx["file"], sheet.name, r, "label", a, "section"))
            continue
        label = a or label
        internal = bool(re.search(GROUP, section, re.I))
        company = group_name(section) if internal else company_from_label(label)
        if not company:
            items.append(ReviewItem(FAMILY, ctx["file"], sheet.name, r, "label", "", "no_company"))
            continue
        out.append(_entry(ctx, sheet, r, cell_text(cells.get("contact")), cells.get("phone"), cells.get("email"),
                          internal, items, company=company,
                          role=_free(ctx, sheet, r, "position", cells.get("position"), items),
                          remark=_free(ctx, sheet, r, "remark", cells.get("remark"), items),
                          department=label,
                          section=section, kind="internal" if internal else ("supplier" if section else "")))
    return out


def read_contacts_sheet(sheet, path):
    """(layout, entries, review items, blank row count); layout None when the sheet is not a contact list."""
    found = find_contacts_header(sheet)
    if found is None:
        return None, [], [], 0
    header_row, cols = found
    layout = "list" if "company" in cols else "sections"
    ctx = {"file": os.path.basename(path), "layout": layout}
    items, stat = [], ContactsFileStat(path)
    reader = _read_list if layout == "list" else _read_sections
    entries = reader(sheet, header_row, cols, ctx, items, stat)
    return layout, entries, items, stat.blank_rows




def _merge(entries, items):
    """One record per (company, person). The richest copy is the base: a copy that only has a name
    never wins, and it fills nothing in that the base already holds. Blanks in the base are filled
    from the others; a value that differs is reported."""
    groups = defaultdict(list)
    for e in entries:
        groups[(e.company_key, _person_key(e.full_name) if e.full_name else "")].append(e)
    out = []
    for (ckey, pkey), group in groups.items():
        group.sort(key=lambda e: (-e.richness(), e.file, e.row))
        base = group[0]
        base.copies = len(group)
        for other in group[1:]:
            for f in ("role", "remark", "department", "address", "ledger", "kind"):
                if not getattr(base, f):
                    setattr(base, f, getattr(other, f))
                elif getattr(other, f) and getattr(other, f) != getattr(base, f) and f in ("role", "address"):
                    items.append(ReviewItem(FAMILY, other.file, other.sheet, other.row, f, getattr(other, f),
                                            "differs_between_copies"))
            for f in ("emails", "phones"):
                for v in getattr(other, f):
                    if v.lower() not in (x.lower() for x in getattr(base, f)):
                        getattr(base, f).append(v)
        base.source_id = f"contact:{ckey}:{pkey}" if pkey else f"company:{ckey}"
        out.append(base)
    return out


def _orgs(entries, master, items):
    """Match every distinct company name once; entries are re-keyed on the name the CRM will get."""
    kinds = defaultdict(set)
    first = {}
    for e in entries:
        kinds[e.company].add(e.kind)
        first.setdefault(e.company, e)
    canonical, matches, queued = {}, {}, set()
    for name, e in first.items():
        if "internal" in kinds[name]:
            canonical[name] = name + INTERNAL_SUFFIX
            continue
        m = master.match(name)
        matches[name] = {"code": m.code, "cn": m.cn, "en": m.en, "score": m.score, "method": m.method}
        canonical[name] = (m.cn or m.en) if m.score >= COMPANY_OK and (m.cn or m.en) else name
        if m.score < COMPANY_OK:
            queued.add(name)
            items.append(ReviewItem(FAMILY, e.file, e.sheet, e.row, "company", name, "company",
                                    [list(c) for c in (m.candidates or [])]))
    orgs = {}
    for e in entries:
        e.company_key = norm(canonical[e.company]) or canonical[e.company].lower()
        org = orgs.get(e.company_key)
        if org is None:
            org = orgs[e.company_key] = Org(e.company_key, canonical[e.company], source_id=f"company:{e.company_key}")
        org.kind = org.kind or e.kind
        org.origin = org.origin or e.origin
        for value, seen in ((e.kind, org.kinds), (e.ledger, org.ledgers)):
            if value and value not in seen:
                seen.append(value)
        org.address = org.address or e.address
        if not e.full_name and e.remark:
            org.location = org.location or e.remark
        # The group's name stays out of the lookup names: a customer of that name must not take its colleagues.
        written = canonical[e.company] if e.kind == "internal" else e.company
        if written not in org.written_names:
            org.written_names.append(written)
        if e.company in matches and not org.match:
            org.match = matches[e.company]
        org.queued = org.queued or e.company in queued
    return orgs


def read_paths(paths, master=None) -> ContactsResult:
    if master is None:
        from .run import IMPORTER_DIR, _master
        master = _master(IMPORTER_DIR)
    skipped, stats, entries, items = [], [], [], []
    for path in unique_contents(paths, skipped):
        stat = ContactsFileStat(path=path)
        stats.append(stat)
        try:
            sheets = load(path)
        except Exception as exc:  # a corrupt or unsupported workbook must not stop the run
            stat.error = f"{type(exc).__name__}: {exc}"
            skipped.append({"file": path, "error": f"could not be read: {stat.error}"})
            continue
        stat.sheets = len(sheets)
        layouts = Counter()
        for sheet in sheets:
            if sheet.cut:
                stat.sheet_errors.append((sheet.name, sheet.cut))
            try:
                layout, got, found, blank = read_contacts_sheet(sheet, path)
            except Exception as exc:
                stat.error = (stat.error + "; " if stat.error else "") + f"{sheet.name}: {exc}"
                stat.sheet_errors.append((sheet.name, str(exc) or type(exc).__name__))
                continue
            if layout is None:
                if len(sheet.grid) >= UNRECOGNISED_MIN_CELLS:
                    stat.unrecognised.append(sheet.name)
                continue
            layouts[layout] += 1
            stat.read_sheets += 1
            stat.rows += len(got)
            stat.blank_rows += blank
            entries += got
            items += found
        stat.layout = layouts.most_common(1)[0][0] if layouts else "none"
    orgs = _orgs(entries, master, items)
    merged = _merge(entries, items)
    return ContactsResult(contacts=[e for e in merged if e.full_name], orgs=orgs, review=items,
                          skipped_files=skipped, per_file_stats=stats)


def summarize(result: ContactsResult) -> dict:
    """Counts for a CLI line, the web UI and the every-file matrix. No cell values."""
    accepted = result.accepted_contacts()
    return {
        "files": len(result.per_file_stats),
        "sheets": sum(s.sheets for s in result.per_file_stats),
        "read_sheets": sum(s.read_sheets for s in result.per_file_stats),
        "rows": sum(s.rows for s in result.per_file_stats),
        "blank_rows": sum(s.blank_rows for s in result.per_file_stats),
        "contacts": len(result.contacts),
        "contacts_accepted": len(accepted),
        "contacts_queued": len(result.contacts) - len(accepted),
        "companies": len(result.orgs),
        "companies_accepted": len(result.accepted_orgs()),
        "companies_queued": sum(o.queued for o in result.orgs.values()),
        "queued_by_reason": dict(sorted(Counter(i.reason for i in result.review if i.reason in BLOCKING).items())),
        "skipped_by_reason": dict(sorted(Counter(i.reason for i in result.review if i.reason not in BLOCKING).items())),
        "skipped_files": len(result.skipped_files),
    }
