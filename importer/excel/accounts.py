"""Read-only reader for the customer-list workbooks (folders "04 ERP Customer Lists by Region" and
"05 Customer Lists"). Pure parsing: nothing here knows the CRM. Five layouts, found by header label
(any column order, header row anywhere in the first rows; a sheet can hold several lists):

  erp           ERP customer export (客户编码 + 客户名称). ONLY the whitelisted columns of ERP_TABLE are
                read: bank accounts, tax IDs, the legal representative and the other finance columns are
                never looked at, so they cannot reach a record or the CRM. The contact person of the row
                (联系人) with his phone, mobile and e-mail (电话, 手机, Email地址) becomes a contact of
                the account, tied to it by the customer code.
  master        Customer_List (Customer Code + CN name): region, product area, dealer, supporting sales,
                first order. A code that is not an ERP code ("New") means the customer has none yet.
  supporting    Customer lists handed over from a colleague (Customer + FIBRO Existing Customer).
  effective     Regional analysis workbook: several lists on one sheet, each under a title row.
                Old / new customers carry application, competitor and expected annual volume;
                competitors of the exhibition become competitor accounts; machine-tool makers who attended
                become leads.
  registration  A customer-registration form (one lead per row).

Rows that are not usable are skipped with a reason, never silently. A bad file or sheet is recorded and
the run goes on. The same customer in several lists is merged on its ERP code (a row without one is
matched by name with excel.companies.Master); the richest row wins each field, and a value that differs
becomes a review item instead of an overwrite. A name the master cannot place with confidence is queued
("company") and stays out of the commit until a person answers."""
import os
import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime

from .companies import CJK_RE, CODE_RE, Master, _canon_code, lookalike, norm
from .readers import ReviewItem
from .files import unique_contents
from .sheet import load
from .split import is_placeholder, phone_e164, phone_numbers

FAMILY = "accounts"
SCAN_ROWS = 8
COMPANY_OK = 0.9
BLOCKING = {"company"}  # review reasons that keep an account or lead out of the commit
UNRECOGNISED_MIN_CELLS = 10

# field -> labels. "=x": the whole label must equal x, otherwise x may be anywhere in the label.
# Labels are NFKC, lowercase, one line. A column serves one field, the first in table order.
ERP_TABLE = [
    ("code", ("=客户编码",)), ("cn", ("=客户名称",)), ("short", ("=客户简称",)), ("mnemonic", ("=助记码",)),
    ("area", ("=地区名称",)), ("industry", ("=所属行业",)), ("address", ("=地址",)),
    ("rep", ("=专营业务员名称",)), ("province", ("=省直辖市",)), ("note", ("=purpose_topics visit",)),
    ("contact", ("=联系人",)), ("phone", ("=电话",)), ("mobile", ("=手机",)), ("email", ("=email地址",)),
]
MASTER_TABLE = [
    ("code", ("=customer code",)), ("cn", ("=cn name",)), ("en", ("=en name",)), ("province", ("=province",)),
    ("region", ("=region",)), ("area", ("=area",)), ("supporting", ("supporting",)), ("dealer", ("=dealer",)),
    ("first_order", ("first order",)),
]
SUPPORT_TABLE = [
    ("cn", ("=客户 customer",)), ("en", ("company english",)), ("city", ("=city",)), ("province", ("=province",)),
    ("industry", ("industry segment",)), ("product_area", ("fibro tables category",)),
    ("existing", ("fibro existing customer",)), ("active", ("active status",)),
    ("u8", ("u8 as",)), ("supported", ("supported by",)),
]
REG_TABLE = [
    ("name", ("customer name", "客户名称")), ("location", ("location", "地点")),
    ("industry", ("industry", "所属行业")), ("application", ("application", "应用")),
    ("brand", ("current brand", "现使用")), ("division", ("division", "分度")),
    ("speed", ("indexing speed", "旋转速度")), ("accuracy", ("accuracy", "精度")), ("inertia", ("mtm", "最大惯量")),
]
EFF_TABLE = [
    ("name", ("客户名称", "厂家名称")), ("en", ("english name",)), ("place", ("地址", "产地")),
    ("contact", ("联系人", "销售代表")), ("phone", ("电话",)), ("application", ("主要应用型号",)),
    ("competitor", ("竞争对手",)), ("forecast", ("预测全年采购数量",)), ("possible", ("可能需求转台型号", "产品系列")),
    ("region", ("所属区域",)), ("intel", ("已建立联系",)),
]
REG_LABELS = {"brand": "Current brand/type", "division": "Division", "speed": "Indexing speed",
              "accuracy": "Accuracy", "inertia": "Max inertia"}

# Tie-break between rows of equal richness: the ERP export is the source of truth.
PRIORITY = {"erp": 0, "master": 1, "supporting": 2, "old": 3, "new": 3, "competitors": 4, "exhibition": 4,
            "registration": 5}
SCALARS = ("cn", "en", "short", "province", "city", "region", "industry", "address", "erp_rep", "supporting",
           "dealer", "product_area", "first_order", "application", "competitor")
NAME_FIELDS = ("en", "short")  # other spellings are aliases, not conflicts
PLACE_FIELDS = ("province", "city", "region", "industry")
PERSON_FIELDS = ("erp_rep", "supporting")
RICHER_WINS = PLACE_FIELDS + PERSON_FIELDS + ("first_order",)
LEGAL_WORDS = re.compile(r"\b(?:co|ltd|limited|inc|gmbh|corp|group|company)\b", re.I)
# A tax number (18 digits and capitals) or a bank account typed into a name column: dropped, never stored.
ID_LIKE = re.compile(r"(?=[0-9A-Z]*\d)[0-9A-Z]{15,20}|\d{10,}")
INSTRUCTION = re.compile(r"\s*pl(?:ea)?se?\b.*\bfill\b", re.I)  # "Pls Xulu fill in necessary": a request in a cell, no data
UNIT = re.compile(r"\s*(\d+(?:\.\d+)?)\s*(?:台|pcs?|sets?|units?)?\s*", re.I)


@dataclass
class Row:
    """One customer, competitor or prospect as one row of one list says it."""
    layout: str  # erp | master | supporting | old | new | competitors | exhibition | registration
    kind: str  # account | competitor | lead
    file: str
    sheet: str
    row: int
    code: str = ""  # canonical ERP code (see companies._canon_code)
    cn: str = ""
    en: str = ""
    short: str = ""
    aliases: list = field(default_factory=list)  # more names written in the same cell
    province: str = ""
    city: str = ""
    region: str = ""  # sales region (East, North ...)
    industry: str = ""
    address: str = ""
    erp_rep: str = ""
    supporting: str = ""
    dealer: str = ""
    product_area: str = ""
    first_order: str = ""  # YYYY or YYYY-MM
    existing: bool | None = None  # "FIBRO existing customer"
    do_not_use: bool = False
    application: str = ""
    competitor: str = ""
    forecast: float | None = None  # expected annual volume (units)
    contact: str = ""
    phone: str = ""  # the phone cell as written
    mobile: str = ""  # the mobile cell as written (ERP lists)
    email: str = ""  # the e-mail cell as written (ERP lists)
    location: str = ""  # a lead's place
    details: str = ""  # a lead's technical requirements
    notes: list = field(default_factory=list)

    def name(self) -> str:
        return self.cn or self.en or self.short

    def richness(self) -> int:
        # the contact columns of an ERP row must not make it richer than a master row: that would change which row wins a field
        return sum(bool(getattr(self, f)) for f in SCALARS) + (self.forecast is not None) + (bool(self.contact) and self.layout != "erp")


@dataclass
class Person:
    """A contact person of an account: key = account (by customer code) + name. Several numbers and
    addresses are kept, each in its own entry."""
    name: str
    phone: str = ""  # the number cells as written, joined by " / " (every row that names this person)
    phones: list = field(default_factory=list)  # E.164, in the order written
    emails: list = field(default_factory=list)
    other: list = field(default_factory=list)  # text of a phone / e-mail cell that is no number or address we can place

    def __eq__(self, other):
        if isinstance(other, tuple):  # the old shape of Account.contacts: (name, phone as written)
            return (self.name, self.phone) == other
        return isinstance(other, Person) and (self.name, self.phone, self.phones, self.emails, self.other) == \
            (other.name, other.phone, other.phones, other.emails, other.other)

    __hash__ = None


@dataclass
class Account:
    key: str  # erp:<code> or name:<normalized name>
    source_id: str  # content-based, never a file position
    code: str = ""
    cn: str = ""
    en: str = ""
    short: str = ""
    names: list = field(default_factory=list)  # every real name written, for finding it in the CRM
    province: str = ""
    city: str = ""
    region: str = ""
    industry: str = ""
    address: str = ""
    erp_rep: str = ""
    supporting: str = ""
    dealer: str = ""
    product_area: str = ""
    first_order: str = ""
    application: str = ""
    competitor: str = ""
    forecast: float | None = None
    lifecycle: str = ""  # customer | prospect | ""
    relationships: list = field(default_factory=list)  # customer | competitor
    contacts: list = field(default_factory=list)  # [Person]; compares equal to (name, phone as written)
    notes: list = field(default_factory=list)
    sources: list = field(default_factory=list)  # [(file, sheet, row, layout)]
    match: dict = field(default_factory=dict)  # how a row without a code was placed
    queued: bool = False

    def display(self) -> str:
        return self.cn or self.en or self.short

    def other_language_name(self) -> str:
        """The name in the other script than display(): shown in the CRM next to it, never translated."""
        shown = self.display()
        for name in self.names:
            if bool(CJK_RE.search(name)) != bool(CJK_RE.search(shown)):
                return name
        return ""


@dataclass
class Lead:
    key: str
    source_id: str
    company: str
    contact: str = ""
    phone: str = ""
    location: str = ""
    industry: str = ""
    application: str = ""
    details: str = ""
    origin: str = ""  # registration | exhibition
    sources: list = field(default_factory=list)
    queued: bool = False


@dataclass
class AccountsFileStat:
    path: str
    layout: str = "none"
    sheets: int = 0
    read_sheets: int = 0
    rows: int = 0  # data rows read
    account_rows: int = 0
    competitor_rows: int = 0
    lead_rows: int = 0
    with_competitor: int = 0
    with_forecast: int = 0
    with_application: int = 0
    blank_rows: int = 0
    error: str = ""
    unrecognised: list = field(default_factory=list)
    sheet_errors: list = field(default_factory=list)
    unread: list = field(default_factory=list)  # header labels in a recognised sheet that no field reads


@dataclass
class AccountsResult:
    accounts: list
    leads: list
    review: list  # ReviewItem: skipped rows, conflicts, companies to answer
    skipped_files: list  # [{"file", "error"}]
    per_file_stats: list = field(default_factory=list)

    def accepted_accounts(self) -> list:
        return [a for a in self.accounts if not a.queued]

    def accepted_leads(self) -> list:
        return [l for l in self.leads if not l.queued]


# ---- cell helpers ----

def _flat(value) -> str:
    """A cell as one line: lines joined by " / ", placeholders ("-", "none") gone, floats without ".0"."""
    if value is None or isinstance(value, (bool, datetime, date)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    lines = [" ".join(line.split()) for line in str(value).splitlines()]
    text = " / ".join(l for l in lines if l and not is_placeholder(l))
    return "" if is_placeholder(text) else text


def _spaced(text: str) -> str:
    """"江苏省Jiangsu" -> "江苏省 Jiangsu": a bilingual cell with a space at the script change."""
    return re.sub(r"([A-Za-z])([㐀-鿿])", r"\1 \2", re.sub(r"([㐀-鿿）)])([A-Za-z])", r"\1 \2", text))


def latin_tail(text: str) -> str:
    """The Latin words after the last Chinese run ("机械行业 Machinery" -> "Machinery"); "" when there are none."""
    last = None
    for last in CJK_RE.finditer(text):
        pass
    tail = (text[last.end():] if last else text).strip(" |/,;:()（）-")
    return tail if re.search(r"[A-Za-z]{2}", tail) else ""


def _label(value) -> str:
    s = unicodedata.normalize("NFKC", str(value if value is not None else "")).replace("\n", " ").lower()
    return " ".join(s.split()).strip("：: ")


def _map(sheet, r, table) -> dict:
    labels = {c: _label(sheet.value(r, c)) for c in range(1, sheet.max_col + 1)}
    cols: dict[str, int] = {}
    for name, pats in table:
        for c, lab in labels.items():
            if not lab or c in cols.values():
                continue
            if any(lab == p[1:] if p.startswith("=") else p in lab for p in pats):
                cols[name] = c
                break
    return cols


def parse_code(cell) -> tuple[str, bool]:
    """(canonical ERP code, whether other text stood in the cell). "00CR991038" -> "CR991038"; "New" -> ""."""
    text = _flat(cell)
    m = CODE_RE.search(text)
    if not m:
        return "", False
    return _canon_code(m.group(1)), bool(CODE_RE.sub(" ", text).strip(" /"))


def parse_forecast(value) -> tuple[float | None, str]:
    """(units per year, the text when it is not a plain number)."""
    if value is None or isinstance(value, bool):
        return None, ""
    if isinstance(value, (int, float)):
        return float(value), ""
    text = _flat(value)
    m = UNIT.fullmatch(text) if text else None
    return (float(m.group(1)), "") if m else (None, text)


def parse_first_order(value) -> str:
    """2021.12 -> "2021-12"; 2022.9 -> "2022-09"; 2020.1 -> "2020": January or October, the number cannot say."""
    if value is None or isinstance(value, (bool, datetime, date)):
        return ""
    text = repr(float(value)) if isinstance(value, (int, float)) else str(value).strip()
    m = re.fullmatch(r"((?:19|20)\d\d)(?:\.(\d{1,2}))?(?:\.0)?", text)
    if not m:
        return ""
    year, month = m.group(1), m.group(2)
    if month in (None, "0", "00", "1") or not 1 <= int(month) <= 12:
        return year
    return f"{year}-{month.zfill(2)}"


def _looks_romanized(en: str, cn: str) -> bool:
    """"wuhanjingbo" beside a Chinese name: a pinyin tag, not an English name."""
    return bool(cn) and en == en.lower() and " " not in en and not LEGAL_WORDS.search(en)


def _is_english_name(text: str) -> bool:
    """The ERP "mnemonic" is an English name for most customers and a run of pinyin capitals for the rest."""
    return bool(re.search(r"[A-Za-z]{3}", text)) and (" " in text or bool(re.search(r"[a-z]", text)))


def _person(text: str) -> str:
    """"Richard.Yu" -> "Richard Yu"."""
    return re.sub(r"(?<=[a-z])\.(?=[A-Z])", " ", text)


def _industry(value) -> tuple[str, bool]:
    text = _flat(value)
    if not text:
        return "", False
    if re.search(r"do\s*not\s*use", text, re.I):
        return "", True
    return latin_tail(text) or _spaced(text), False


def _place(text: str) -> tuple[str, str]:
    """(province, city) of a one-cell place: a province or autonomous region is a province."""
    text = _spaced(_flat(text))
    if not text or latin_tail(text).lower().startswith("other"):
        return "", ""
    cjk = "".join(CJK_RE.findall(text))
    return (text, "") if cjk.endswith(("省", "自治区")) else ("", text)


def _city_of_area(area: str, province: str) -> str:
    text = _spaced(_flat(area))
    if not text or latin_tail(text).lower().startswith("other"):
        return ""
    cjk = "".join(CJK_RE.findall(text))
    return "" if cjk.endswith(("省", "自治区")) or text == province else text


# ---- contact numbers and addresses ----

_EMAIL = re.compile(r"[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+")  # same rule as visits.pipeline.contact_emails
_DASHES = dict.fromkeys(map(ord, "\u2010\u2011\u2012\u2013\u2014\u2015\u2212"), "-")
_PLACED = re.compile(r"\+\d{8,15}")
_SERVICE_NUMBER = re.compile(r"[48]00\d{7}")  # 400 / 800 service lines


def _number(token: str) -> str:
    """E.164 for a token that is exactly one phone number (excel.split.phone_e164 rules, plus 400/800
    service lines), else "": a number without area code or country code is never guessed at."""
    token = unicodedata.normalize("NFKC", token).translate(_DASHES).strip()
    e164 = phone_e164(token)
    if _PLACED.fullmatch(e164):
        return e164
    digits = re.sub(r"[\s\-().]", "", token)
    return "+86" + digits if _SERVICE_NUMBER.fullmatch(digits) else ""


def parse_numbers(text: str) -> tuple[list, list]:
    """(E.164 numbers, texts that could not be placed) of a phone cell. One cell can hold several numbers
    (split on / , ; and spaces). An extension, a name written after the number or a shorthand "/65" places the
    number and keeps the whole text as well. A local number, a word ("call later") or a bad number is kept
    as text only. An e-mail address typed into a phone cell is left to parse_emails."""
    text = _EMAIL.sub(" ", unicodedata.normalize("NFKC", text or "")).translate(_DASHES)
    placed, left = [], []
    groups = []
    for part in re.split(r"[/,;、\n]", text):  # explicit separators first: split.phone_numbers would glue a short tail to the number before it
        if not part.strip():
            continue
        if groups and re.fullmatch(r"\s*\d{1,3}\s*", part):  # "010-82241866/65": the last digits of a second number
            groups[-1] = (groups[-1][0], f"{groups[-1][1]}/{part.strip()}")
            continue
        groups += [(g.strip(), g.strip()) for g in phone_numbers(part)]
    for group, shown in groups:
        if n := _number(group):
            placed.append(n)
            if shown != group:
                left.append(shown)
            continue
        rest = group
        if m := re.fullmatch(r"(.*\d)\s*([^\d\s+\-/.].*)", rest):  # a name or word after the number
            rest = m.group(1)
        heads = [rest]
        if m := re.fullmatch(r"(.*\d)\s*-\s*(\d{1,5})", rest):  # an extension
            heads.append(m.group(1))
        if n := next((n for n in map(_number, heads) if n), ""):
            placed.append(n)
        left.append(shown)
    return list(dict.fromkeys(placed)), left


def parse_emails(text: str) -> tuple[list, list]:
    """(addresses, text left over) of an e-mail cell, or of a phone cell that holds an address."""
    found = list(dict.fromkeys(_EMAIL.findall(text or "")))
    rest = _EMAIL.sub(" ", text or "")
    return found, ([" ".join(rest.split())] if rest.strip(" /,;、\n\t") else [])


def _person_key(name: str) -> str:
    return "".join(unicodedata.normalize("NFKC", name).casefold().split())


def _people(rows) -> tuple[list, list]:
    """([Person], [text]) of one account: one Person per name (the name is the key within the account, so two
    people of the same name at two companies stay two entries), numbers and addresses of every row that names
    him joined. The texts are contact cells of rows that name nobody."""
    people: dict[str, Person] = {}
    written: dict[str, list] = {}  # person key -> the number cells as written, each once
    nameless = []
    for r in rows:
        cells = [("phone", r.phone), ("mobile", r.mobile), ("e-mail", r.email)]
        if not r.contact:
            nameless += [f"{label} {text}" for label, text in cells if text]
            continue
        p = people.setdefault(_person_key(r.contact), Person(r.contact))
        cells_written = written.setdefault(_person_key(r.contact), [])
        cells_written += [t for t in (r.phone, r.mobile) if t and t not in cells_written]
        p.phone = " / ".join(cells_written)
        numbers = []
        for text in (r.mobile, r.phone):  # a mobile first: it is the number that reaches the person
            got, left = parse_numbers(text)
            numbers += got
            p.other += [t for t in left if t not in p.other]
            if "@" in text:  # an address typed into a phone cell
                p.emails += [e for e in parse_emails(text)[0] if e not in p.emails]
        p.phones += [n for n in dict.fromkeys(numbers) if n not in p.phones]
        mails, junk = parse_emails(r.email)
        p.emails += [e for e in mails if e not in p.emails]
        p.other += [t for t in junk if t not in p.other]
    return list(people.values()), nameless


# ---- layouts ----

def _find(sheet, table, need):
    for r in range(1, min(SCAN_ROWS, sheet.max_row) + 1):
        cols = _map(sheet, r, table)
        if set(need) <= cols.keys():
            return r, cols
    return None


def _effective_headers(sheet) -> list:
    out = []
    for r in range(1, sheet.max_row + 1):
        labels = [_label(sheet.value(r, c)) for c in range(1, sheet.max_col + 1)]
        if any("english name" in l for l in labels) and any("客户名称" in l or "厂家名称" in l for l in labels):
            out.append((r, _map(sheet, r, EFF_TABLE)))
    return out


def detect(sheet) -> str | None:
    """erp | master | supporting | registration | effective, or None for a sheet that is not a customer list."""
    if _find(sheet, ERP_TABLE, ("code", "cn")):
        return "erp"
    if _find(sheet, MASTER_TABLE, ("code", "cn")):
        return "master"
    if _find(sheet, SUPPORT_TABLE, ("cn", "existing")):
        return "supporting"
    if _effective_headers(sheet):
        return "effective"
    return "registration" if _find(sheet, REG_TABLE, ("name", "location", "application")) else None


class _Reader:
    def __init__(self, sheet, path):
        self.sheet, self.path, self.rows, self.items, self.blank = sheet, path, [], [], 0

    def get(self, r, cols, name) -> str:
        return _flat(self.sheet.value(r, cols[name])) if name in cols else ""

    def row(self, r, layout, kind, **kw) -> Row:
        return Row(layout=layout, kind=kind, file=self.path, sheet=self.sheet.name, row=r, **kw)

    def skip(self, r, fieldname, value, reason):
        self.items.append(ReviewItem(FAMILY, self.path, self.sheet.name, r, fieldname, value, reason))

    def add(self, row: Row):
        for f in SCALARS:
            if ID_LIKE.fullmatch(getattr(row, f).replace(" ", "")):
                setattr(row, f, "")
                self.skip(row.row, f, "", "id_number_in_text_field")
        self.rows.append(row)


def _read_erp(rd: _Reader, h, cols):
    sheet = rd.sheet
    # a column with no header text, inside the table, holds a free sales note about the row (seen between 联系人 and 电话)
    unlabelled = [c for c in range(1, max(cols.values()) + 1) if not _label(sheet.value(h, c)) and c not in cols.values()]
    for r in range(h + 1, sheet.max_row + 1):
        raw, name = rd.get(r, cols, "code"), rd.get(r, cols, "cn")
        if not (raw or name or rd.get(r, cols, "short")):
            rd.blank += 1
            continue
        code, polluted = parse_code(sheet.value(r, cols["code"]))
        if not name:
            rd.skip(r, "cn", raw, "no_name")
            continue
        if not code:
            rd.skip(r, "code", raw, "no_code")
            continue
        if polluted:
            rd.skip(r, "code", code, "code_cell_has_text")
        industry, dnu = _industry(sheet.value(r, cols["industry"]) if "industry" in cols else None)
        province = _spaced(rd.get(r, cols, "province"))
        mnemonic = rd.get(r, cols, "mnemonic")
        row = rd.row(r, "erp", "account", code=code, cn=name, short=rd.get(r, cols, "short"),
                     en=mnemonic if _is_english_name(mnemonic) else "", province=province,
                     city=_city_of_area(rd.get(r, cols, "area"), province), industry=industry,
                     address=rd.get(r, cols, "address"), erp_rep=_person(rd.get(r, cols, "rep")), do_not_use=dnu,
                     contact=rd.get(r, cols, "contact"), phone=rd.get(r, cols, "phone"),
                     mobile=rd.get(r, cols, "mobile"), email=rd.get(r, cols, "email"))
        if dnu:
            row.notes.append("Marked \"Do Not use\" in the ERP customer list.")
            rd.skip(r, "industry", name, "do_not_use")
        if note := rd.get(r, cols, "note"):
            row.notes.append(f"Purpose / topics of visit (contacted-customers list): {note}")
        for c in unlabelled:
            if text := _flat(sheet.value(r, c)):
                row.notes.append(f"Sales note (unlabelled column of the ERP list): {text}")
        rd.add(row)


def _read_master(rd: _Reader, h, cols):
    sheet = rd.sheet
    for r in range(h + 1, sheet.max_row + 1):
        raw, name = rd.get(r, cols, "code"), rd.get(r, cols, "cn")
        if not (raw or name or rd.get(r, cols, "en")):
            rd.blank += 1
            continue
        if not (name or rd.get(r, cols, "en")):
            rd.skip(r, "cn", raw, "no_name")
            continue
        code, polluted = parse_code(sheet.value(r, cols["code"]))  # "New": the customer has no ERP code yet
        if polluted:
            rd.skip(r, "code", code, "code_cell_has_text")
        first = parse_first_order(sheet.value(r, cols["first_order"])) if "first_order" in cols else ""
        rd.add(rd.row(r, "master", "account", code=code, cn=name, en=rd.get(r, cols, "en"),
                      province=_spaced(rd.get(r, cols, "province")), region=rd.get(r, cols, "region"),
                      product_area=rd.get(r, cols, "area"), supporting=rd.get(r, cols, "supporting"),
                      dealer=rd.get(r, cols, "dealer"), first_order=first))


def _read_supporting(rd: _Reader, h, cols):
    for r in range(h + 1, rd.sheet.max_row + 1):
        name = rd.get(r, cols, "cn")
        if not (name or rd.get(r, cols, "en")):
            rd.blank += 1
            continue
        existing = {"y": True, "n": False}.get(rd.get(r, cols, "existing").lower())
        industry, _ = _industry(rd.sheet.value(r, cols["industry"]) if "industry" in cols else None)
        row = rd.row(r, "supporting", "account", cn=name, en=rd.get(r, cols, "en"), existing=existing,
                     city=_spaced(rd.get(r, cols, "city")), province=_spaced(rd.get(r, cols, "province")),
                     industry=industry, product_area=rd.get(r, cols, "product_area"),
                     erp_rep=_person(rd.get(r, cols, "u8")), supporting=rd.get(r, cols, "supported"))
        if rd.sheet.name.strip().casefold() != "supporting":  # the default group needs no label; "Migrated" says the customer was migrated
            row.notes.append(f"Customer hand-over list, sheet '{rd.sheet.name.strip()}'.")
        active = rd.get(r, cols, "active")
        if active.lower() in ("y", "n"):
            row.notes.append(f"Order placed 2024-2025: {active.upper()} (customer hand-over list).")
        elif INSTRUCTION.match(active):  # a reminder to the list's owner, not a value
            rd.skip(r, "active", active, "instruction_not_data")
        elif active:  # free text where Y / N was asked ("Projekt on hand WK52,2025")
            row.notes.append(f"Order placed 2024-2025 (customer hand-over list): {active}")
        rd.add(row)


def _read_registration(rd: _Reader, h, cols):
    for r in range(h + 1, rd.sheet.max_row + 1):
        name = rd.get(r, cols, "name")
        if not name:
            if any(rd.sheet.value(r, c) not in (None, "") for c in range(1, rd.sheet.max_col + 1)):
                rd.skip(r, "name", "", "no_name")
            else:
                rd.blank += 1
            continue
        details = "; ".join(f"{label}: {rd.get(r, cols, f)}" for f, label in REG_LABELS.items() if rd.get(r, cols, f))
        rd.add(rd.row(r, "registration", "lead", cn=name if CJK_RE.search(name) else "",
                      en="" if CJK_RE.search(name) else name, location=_spaced(rd.get(r, cols, "location")),
                      industry=rd.get(r, cols, "industry"), application=rd.get(r, cols, "application"),
                      details=details))


def _title(sheet, header_row) -> str:
    for r in range(header_row - 1, max(header_row - 4, 0), -1):
        text = _flat(sheet.value(r, 1))
        if text and not re.fullmatch(r"[\d.]+", text):
            return text
    return ""


def _read_effective(rd: _Reader, headers):
    sheet = rd.sheet
    for i, (h, cols) in enumerate(headers):
        end = headers[i + 1][0] - 1 if i + 1 < len(headers) else sheet.max_row
        title = latin_tail(_title(sheet, h)) or _title(sheet, h)
        if "competitor" in cols:
            kind, layout = "account", "new" if re.search(r"新客户|new customers", _title(sheet, h), re.I) else "old"
        elif "intel" in cols:
            kind, layout = "competitor", "competitors"
        elif "possible" in cols:
            kind, layout = "lead", "exhibition"
        else:
            rd.skip(h, "section", title, "unrecognised_section")
            continue
        intel = _flat(sheet.value(h, cols["intel"])) if "intel" in cols else ""
        for r in range(h + 1, end + 1):
            if sheet.is_merged_continuation(r, cols["name"]):  # a title row, merged over the whole width
                continue
            names = [n for n in (_flat(p) for p in str(sheet.value(r, cols["name"]) or "").splitlines()) if n]
            if not names:
                rd.blank += 1
                continue
            cn = names[0]
            en = rd.get(r, cols, "en")
            if _looks_romanized(en, cn):
                en = ""
            province, city = _place(rd.get(r, cols, "place"))
            row = rd.row(r, layout, kind, cn=cn, en=en, aliases=names[1:], province=province, city=city,
                         contact=rd.get(r, cols, "contact"), phone=rd.get(r, cols, "phone"))
            if not CJK_RE.search(cn):  # an English-only name is the English name
                row.cn, row.en = "", en or cn
            if kind == "account":
                row.application, row.competitor = rd.get(r, cols, "application"), rd.get(r, cols, "competitor")
                row.forecast, text = parse_forecast(sheet.value(r, cols["forecast"]) if "forecast" in cols else None)
                if text:
                    rd.skip(r, "forecast", text, "forecast_not_a_number")
            elif kind == "competitor":
                row.notes.append(f"{title}. {intel}" if intel else title)
            else:
                row.location = " ".join(p for p in (row.city, row.province) if p)
                row.application = rd.get(r, cols, "possible")
                region = rd.get(r, cols, "region")
                row.details = f"Exhibition region: {latin_tail(region) or region}" if region else ""
                row.notes.append(f"{title}." + (f" Possible requirement: {row.application}" if row.application else ""))
            rd.add(row)


def unread_labels(sheet, layout) -> list:
    """Header labels of a recognised sheet that no field of the layout reads (for the notes on the page)."""
    table = {"erp": ERP_TABLE, "master": MASTER_TABLE, "supporting": SUPPORT_TABLE, "registration": REG_TABLE}.get(layout)
    if table is None:
        return []
    need = {"erp": ("code", "cn"), "master": ("code", "cn"), "supporting": ("cn", "existing"), "registration": ("name", "location", "application")}[layout]
    found = _find(sheet, table, need)
    if found is None:
        return []
    h, cols = found
    used = set(cols.values())
    out = []
    for c in range(1, sheet.max_col + 1):
        text = " ".join(str(sheet.value(h, c) or "").split())
        if text and c not in used:
            out.append(text)
    return out


def read_accounts_sheet(sheet, path):
    """(layout or None, rows, review items, blank rows)."""
    layout = detect(sheet)
    if layout is None:
        return None, [], [], 0
    rd = _Reader(sheet, path)
    if layout == "effective":
        _read_effective(rd, _effective_headers(sheet))
    else:
        table, need, reader = {
            "erp": (ERP_TABLE, ("code", "cn"), _read_erp), "master": (MASTER_TABLE, ("code", "cn"), _read_master),
            "supporting": (SUPPORT_TABLE, ("cn", "existing"), _read_supporting),
            "registration": (REG_TABLE, ("name", "location", "application"), _read_registration)}[layout]
        h, cols = _find(sheet, table, need)
        reader(rd, h, cols)
    return layout, rd.rows, rd.items, rd.blank


# ---- merge ----



def _letters(text: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFKC", text).casefold() if ch.isalnum())


def _same(name: str, a: str, b: str) -> bool:
    """Two values of one field that say the same thing, written two ways."""
    ka, kb = norm(a), norm(b)
    if ka == kb:
        return True
    if name == "cn":
        return lookalike(a, b)
    if name in PLACE_FIELDS:
        ca, cb = "".join(CJK_RE.findall(a)), "".join(CJK_RE.findall(b))
        if ca and cb and not (ca in cb or cb in ca):  # two Chinese spellings of different places: "广东省" is not "江苏省"
            return False
        ta, tb = norm(latin_tail(a)), norm(latin_tail(b))
        return not ta or not tb or ta == tb or ta in tb or tb in ta  # a Chinese-only value is not comparable to an English one
    if name in PERSON_FIELDS:  # "Wang Yao" = "Yao Wang"; "Richard" = "Richard Yu"
        ta, tb = set(re.findall(r"[a-z]+", a.lower())), set(re.findall(r"[a-z]+", b.lower()))
        return bool(ka and kb) and (ka in kb or kb in ka or ta == tb)
    if name == "first_order":
        return bool(ka and kb) and (ka in kb or kb in ka)
    if name == "product_area":  # Motion = Rotomotion, Cutting = Rotocutting
        return ka.removeprefix("roto") == kb.removeprefix("roto")
    return False


def _master_with(master: Master, rows) -> Master:
    """The company master of the visit importer plus every ERP-coded row of this read, answers applied."""
    # a CRM company without an ERP code carries a "crm:<id>" stand-in; it is no customer code, and parse_code would
    # read the tail of the id as one
    recs = [{"source": "master", "code": "" if c.code.startswith("crm:") else parse_code(c.code)[0], "cn": c.cn, "en": c.en,
             "short": c.short} for c in master.companies]
    recs += [{"source": "lists", "code": r.code, "cn": r.cn, "en": r.en, "short": r.short} for r in rows if r.code]
    merged = Master.from_records(recs)
    from .decisions import DecisionStore, default_path
    from .run import IMPORTER_DIR
    merged.skipped_decisions = merged.apply_decisions(DecisionStore(default_path(IMPORTER_DIR)).as_dict(), strict=False)
    return merged


def _known(m) -> bool:
    return m.score >= COMPANY_OK and bool(m.cn or m.en)


def _group(rows, master, items):
    """{group key: rows}; also {key: info} for the groups a row without code could only be placed by name."""
    master = _master_with(master, rows)
    groups: dict[str, list] = defaultdict(list)
    info: dict[str, dict] = {}
    placed: dict[str, tuple] = {}
    for row in rows:
        if row.code:
            groups[f"erp:{row.code}"].append(row)
            continue
        name = row.cn or row.en or row.short
        if not name:
            continue
        if name not in placed:
            m = master.match(row.cn or row.en)
            if not _known(m) and row.cn and row.en:  # the English name may find what the Chinese one did not
                m2 = master.match(row.en)
                m = m2 if _known(m2) else m
            placed[name] = (m, f"name:{norm(name) or name.lower()}")
        m, name_key = placed[name]
        if _known(m):  # a company the master knows, with or without an ERP code
            key = f"erp:{_canon_code(m.code)}" if m.code else f"name:{norm(m.cn or m.en)}"
            info.setdefault(key, {"hit": True, "code": m.code, "cn": m.cn, "en": m.en, "score": m.score,
                                  "method": m.method})
        elif m.method in ("new", "none"):
            key = name_key
        else:  # a question for a person: this row alone, kept out until answered
            key = name_key
            if key not in info:
                items.append(ReviewItem(FAMILY, row.file, row.sheet, row.row, "company", name, "company",
                                        [list(c) for c in (m.candidates or [])]))
                info[key] = {"queued": True, "score": m.score, "method": m.method}
        groups[key].append(row)
    return groups, info


def _account(key, rows, info, items) -> Account:
    rows.sort(key=lambda r: (-r.richness(), PRIORITY[r.layout], r.file, r.row))
    code = key[4:] if key.startswith("erp:") else ""
    acc = Account(key=key, source_id=f"account:{code or key[5:]}", code=code)
    acc.match = {k: v for k, v in info.items() if k not in ("queued", "hit")}
    acc.queued = bool(info.get("queued"))
    for f in SCALARS:
        vals = [(r, getattr(r, f)) for r in rows if getattr(r, f)]
        if not vals:
            continue
        winner = vals[0][1]
        if f in RICHER_WINS:  # the same value written more fully in another row ("广东省" / "广东省 Guangdong") keeps the fuller text
            for r, v in vals[1:]:
                if _same(f, winner, v) and not (Counter(_letters(winner)) - Counter(_letters(v))) and Counter(_letters(v)) != Counter(_letters(winner)):
                    winner = v
        setattr(acc, f, winner)
        for r, v in vals[1:]:
            if f not in NAME_FIELDS and not _same(f, winner, v):
                items.append(ReviewItem(FAMILY, r.file, r.sheet, r.row, f, v,
                                        "name_differs" if f == "cn" else "conflict",
                                        [[winner, os.path.basename(vals[0][0].file)]]))
    nums = [(r, r.forecast) for r in rows if r.forecast is not None]
    if nums:
        acc.forecast = nums[0][1]
        items += [ReviewItem(FAMILY, r.file, r.sheet, r.row, "forecast", str(v), "conflict",
                             [[str(acc.forecast), os.path.basename(nums[0][0].file)]])
                  for r, v in nums[1:] if v != acc.forecast]
    primary = [info[k] for k in ("cn", "en") if info.get(k)] if info and not any(r.code for r in rows) else []
    names = [n for n in [*primary, *(n for r in rows for n in (r.cn, r.en, *r.aliases))] if n]
    acc.names = list(dict.fromkeys(names))
    if primary:  # a name the master knows is filed under the master's spelling, as the visit importer does
        acc.cn, acc.en = (info.get("cn") or acc.cn), (info.get("en") or acc.en)
    ranks = {"": 0, "prospect": 1, "customer": 2}
    for r in rows:
        life = ("customer" if r.layout == "erp" and not r.do_not_use else
                "customer" if r.layout == "master" and r.first_order else
                "customer" if r.layout == "supporting" and r.existing else
                "customer" if r.layout == "old" else "prospect" if r.layout == "new" else "")
        if ranks[life] > ranks[acc.lifecycle]:
            acc.lifecycle = life
    if any(r.do_not_use for r in rows):  # one list says "Do Not use", another says it buys: not known, so not a customer
        acc.lifecycle = ""
    acc.relationships = (["customer"] if acc.lifecycle == "customer" else []) + \
                        (["competitor"] if any(r.kind == "competitor" for r in rows) else [])
    acc.contacts, nameless = _people(rows)
    notes = [n for r in rows for n in r.notes]
    notes += [f"Contact {p.name}: other text of the phone / e-mail cells (extension, local number, remark): {' / '.join(p.other)}"
              for p in acc.contacts if p.other]
    if nameless:
        notes.append("Contact details of the customer list that name no person: " + "; ".join(dict.fromkeys(nameless)))
    acc.notes = list(dict.fromkeys(notes))
    acc.sources = [(r.file, r.sheet, r.row, r.layout) for r in rows]
    return acc


def _lead(key, rows, queued) -> Lead:
    rows.sort(key=lambda r: (-r.richness(), PRIORITY[r.layout], r.file, r.row))
    first = rows[0]
    pick = lambda f: next((getattr(r, f) for r in rows if getattr(r, f)), "")
    return Lead(key=key, source_id=f"lead:{key[5:]}", company=first.name(), contact=pick("contact"),
                phone=pick("phone"), location=pick("location"), industry=pick("industry"),
                application=pick("application"), details=pick("details"),
                origin=first.layout, sources=[(r.file, r.sheet, r.row, r.layout) for r in rows], queued=queued)


def read_paths(paths, master=None) -> AccountsResult:
    if master is None:
        from .run import IMPORTER_DIR, _master
        master = _master(IMPORTER_DIR)
    skipped, stats, rows, items = [], [], [], []
    for path in unique_contents(paths, skipped):
        stat = AccountsFileStat(path=path)
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
                layout, got, found, blank = read_accounts_sheet(sheet, path)
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
            stat.unread += [t for t in unread_labels(sheet, layout) if t not in stat.unread]
            stat.rows += len(got)
            stat.blank_rows += blank
            stat.account_rows += sum(r.kind == "account" for r in got)
            stat.competitor_rows += sum(r.kind == "competitor" for r in got)
            stat.lead_rows += sum(r.kind == "lead" for r in got)
            stat.with_competitor += sum(bool(r.competitor) for r in got)
            stat.with_forecast += sum(r.forecast is not None for r in got)
            stat.with_application += sum(bool(r.application) for r in got)
            rows += got
            items += found
        stat.layout = layouts.most_common(1)[0][0] if layouts else "none"
    groups, info = _group(rows, master, items)
    accounts, leads = [], []
    for key, members in groups.items():
        queued = bool(info.get(key, {}).get("queued"))
        if all(r.kind == "lead" for r in members) and not info.get(key, {}).get("hit"):
            leads.append(_lead(key, members, queued))
        else:
            accounts.append(_account(key, members, info.get(key, {}), items))
    twins = defaultdict(list)  # one company under two ERP codes
    for a in accounts:
        if a.code and a.cn:
            twins[norm(a.cn)].append(a)
    for group in twins.values():
        for a in group[1:]:
            file, sheet, row, _ = a.sources[0]
            items.append(ReviewItem(FAMILY, file, sheet, row, "code", a.code, "same_name_other_code",
                                    [[g.code, ""] for g in group if g is not a]))
    return AccountsResult(accounts=accounts, leads=leads, review=items, skipped_files=skipped, per_file_stats=stats)


def file_kind(path) -> tuple[str, str]:
    """("accounts", layouts) when the workbook holds a customer list (any sheet), else ("unrecognised", why)."""
    try:
        sheets = load(path)
    except Exception as exc:
        return "unrecognised", f"could not be read: {type(exc).__name__}"
    found = sorted({d for d in (detect(sh) for sh in sheets) if d})
    return ("accounts", ", ".join(found)) if found else ("unrecognised", "no customer-list header on any sheet")


def summarize(result: AccountsResult) -> dict:
    """Counts for a CLI line, the web UI and the every-file matrix. No cell values."""
    accepted = result.accepted_accounts()
    return {
        "files": len(result.per_file_stats),
        "sheets": sum(s.sheets for s in result.per_file_stats),
        "read_sheets": sum(s.read_sheets for s in result.per_file_stats),
        "rows": sum(s.rows for s in result.per_file_stats),
        "blank_rows": sum(s.blank_rows for s in result.per_file_stats),
        "accounts": len(result.accounts),
        "accounts_accepted": len(accepted),
        "accounts_queued": len(result.accounts) - len(accepted),
        "customers": sum(a.lifecycle == "customer" for a in accepted),
        "prospects": sum(a.lifecycle == "prospect" for a in accepted),
        "competitors": sum("competitor" in a.relationships for a in accepted),
        "with_competitor": sum(bool(a.competitor) for a in accepted),
        "with_forecast": sum(a.forecast is not None for a in accepted),
        "with_application": sum(bool(a.application) for a in accepted),
        "with_contacts": sum(bool(a.contacts) for a in accepted),
        "notes": sum(len(a.notes) for a in accepted),
        "leads": len(result.leads),
        "leads_accepted": len(result.accepted_leads()),
        "conflicts": sum(i.reason == "conflict" for i in result.review),
        "queued_by_reason": dict(sorted(Counter(i.reason for i in result.review if i.reason in BLOCKING).items())),
        "skipped_by_reason": dict(sorted(Counter(i.reason for i in result.review if i.reason not in BLOCKING).items())),
        "skipped_files": len(result.skipped_files),
    }


def file_report(result: AccountsResult) -> list[dict]:
    """Counts per file (basename), for the report and the every-file matrix."""
    conflicts = Counter(os.path.basename(i.file) for i in result.review if i.reason == "conflict")
    out = []
    for s in result.per_file_stats:
        name = os.path.basename(s.path)
        out.append({"file": name, "layout": s.layout, "rows": s.rows, "account_rows": s.account_rows,
                    "competitor_rows": s.competitor_rows, "lead_rows": s.lead_rows,
                    "with_competitor": s.with_competitor, "with_forecast": s.with_forecast,
                    "with_application": s.with_application, "conflicts": conflicts[name], "error": s.error})
    return out
