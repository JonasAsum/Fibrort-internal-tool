"""Quality checks for the account readers (folders 04 and 05), run on every cell of every workbook.

The independent oracle (oracle_accounts.py, built from the cell ledger only) says which cells exist and
what header stands over each. These rules then check, for the whole population and without sampling:

  cell-lost            every non-empty data cell is found in a reader record (letters and digits, using
                       invariants.chars) or is dropped by a DECLARED rule below, with its reason
  wrong-field          a cell under a header whose field is known landed in that field (or is declared)
  mixed-cells          a text field holds the text of one cell, never of two cells joined
  merge-loses-value    a value of a row survives the merge across files or is reported as a conflict
  shape-*              codes, provinces, names, phone numbers and e-mail addresses look like what they are
  untied-record        every record names the file / sheet / row it came from, and that row exists

A rule listed in KNOWN_BUGS is a bug we know about and have not fixed: it is strict-xfail, so it keeps
failing loudly in the report, and the day it is fixed the marker must be removed.
"""
import collections
import os
import re
import unicodedata

import pytest

from excel import accounts
from excel.files import unique_contents
from excel.sheet import load
from tests.quality import oracle_accounts, snapshot_accounts
from tests.quality.invariants import chars, is_placeholder

KNOWN_BUGS = {
}

# ---- declared rules: a cell may be dropped only for one of these reasons ----

# ERP / list columns that are kept out on purpose (label, casefolded, whole text) -> reason.
KEPT_OUT = {
    "序号": "row counter of the list",
    "序号 item": "row counter of the list",
    "序号 no.": "row counter of the list",
    "纳税人登记号": "tax number: never imported (finance)",
    "开户银行": "bank name: never imported (finance)",
    "银行账号": "bank account: never imported (finance)",
    "所属银行名称": "bank name: never imported (finance)",
    "法人": "legal representative: never imported",
    "销售默认收付款协议编码": "payment terms code: finance",
    "币种": "currency: finance",
    "税率%": "tax rate: finance",
    "潜在客户编码": "empty column in the ERP export",
    "发货地址": "shipping address: logistics (see decisions D-ACC-03)",
    "发运方式名称": "shipping method: logistics",
    "发货仓库名称": "shipping warehouse: logistics",
    "分管部门名称": "ERP department: always the product line of the export (see D-ACC-03)",
    "客户分类编码": "ERP customer class code",
    "客户分类名称": "ERP customer class name",
    "传真": "fax number: not a contact number (see D-ACC-04)",
}
PINYIN_TAG = re.compile(r"^[A-Z0-9]{2,}$")  # an ERP mnemonic of capitals only is a pinyin tag, not an English name
OTHER_WORDS = re.compile(r"other", re.I)  # area / province cell "其他Others": no place to file
INDUSTRY_LABELS = {"所属行业", "行业 industry segment", "industry所属行业", "所属区域 region"}  # bilingual cells keep their English tail
INSTRUCTION = re.compile(r"\s*pl(?:ea)?se?\b.*\bfill\b", re.I)
PLACE_LABELS = {"province", "省直辖市", "city", "地区名称"}
CJK = re.compile(r"[\u4e00-\u9fff]")
ID_LIKE = re.compile(r"(?=[0-9A-Z]*\d)[0-9A-Z]{15,20}|\d{10,}")
DO_NOT_USE = re.compile(r"do\s*not\s*use", re.I)

# header label (casefolded) -> record fields in which the cell's text must be found (the right CRM field).
# (a label matches when it equals the key, or the key ends with "*" and the label starts with it)
FIELDS = {
    "erp": {"客户编码": ("code",), "客户名称": ("cn",), "客户简称": ("short",), "助记码": ("en",),
            "地区名称": ("city", "province"), "所属行业": ("industry",), "地址": ("address",),
            "专营业务员名称": ("erp_rep",), "省直辖市": ("province",), "purpose_topics visit": ("notes",),
            "联系人": ("contact",), "电话": ("phone",), "手机": ("mobile",), "email地址": ("email",)},
    "master": {"customer code": ("code",), "cn name": ("cn",), "en name": ("en",), "province": ("province",),
               "region": ("region",), "area": ("product_area",), "previous / current supporting": ("supporting",),
               "dealer": ("dealer",), "first order time from 2020": ("first_order",)},
    "supporting": {"客户 customer": ("cn",), "company english short name": ("en",), "company english name": ("en",),
                   "city": ("city",), "province": ("province",), "行业 industry segment": ("industry",),
                   "fibro tables category": ("product_area",), "active status/order placed 2024-2025": ("notes",),
                   "u8 asigned to (previously)": ("erp_rep",)},
    "registration": {"customer name客户名称": ("cn", "en"), "location地点": ("location",),
                     "industry所属行业": ("industry",), "application应用": ("application",)},
    "effective": {"客户名称 customer name": ("cn", "en", "aliases"), "厂家名称 manufacturer's name": ("cn", "en", "aliases"),
                  "english name": ("en",), "地址 adress": ("province", "city"), "产地 adress": ("province", "city"),
                  "联系人 contact name": ("contact",), "销售代表 sales representative": ("contact",),
                  "电话 phone": ("phone",), "主要应用型号 main application models": ("application",),
                  "竞争对手 competitor": ("competitor",)},
}
SCALARS = ("cn", "en", "short", "province", "city", "region", "industry", "address", "erp_rep", "supporting",
           "dealer", "product_area", "first_order", "application", "competitor", "location")
COMPOSITE = ("details", "notes", "aliases")  # fields made from several cells on purpose
COMPANY_SUFFIX = re.compile(r"有限公司|有限责任公司|股份|\bco\.?,? ?ltd\b|\bltd\b|\bgmbh\b|\binc\b|\bcorp\b|\bllc\b|集团", re.I)
EMAIL = re.compile(r"^[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$")
E164 = re.compile(r"^\+\d{8,15}$")


def lab(text) -> str:
    return unicodedata.normalize("NFKC", " ".join(str(text).split())).casefold()


def strings(obj):
    """Every text held anywhere inside a record (fields, lists, tuples, nested records)."""
    if isinstance(obj, str):
        yield obj
    elif isinstance(obj, dict):
        for v in obj.values():
            yield from strings(v)
    elif isinstance(obj, (list, tuple, set)):
        for v in obj:
            yield from strings(v)
    elif hasattr(obj, "__dataclass_fields__"):
        for f in obj.__dataclass_fields__:
            yield from strings(getattr(obj, f))
    elif isinstance(obj, (int, float)) and not isinstance(obj, bool):
        yield str(int(obj)) if float(obj).is_integer() else str(obj)


def subset(small, big) -> bool:
    return not (chars(small) - chars(big))


def latin_tail(text) -> str:
    """The text after the last Chinese character (a bilingual cell's English half), trimmed."""
    last = None
    for last in CJK.finditer(text):
        pass
    return (text[last.end():] if last else text).strip(" |/,;:()（）-")


def CJK_ONLY(text) -> str:
    return "".join(CJK.findall(text))


DASH = dict.fromkeys(map(ord, "\u2010\u2011\u2012\u2013\u2014\u2015\u2212"), "-")
CONTACT_LABELS = {"联系人", "联系人 contact name", "销售代表 sales representative"}
PHONE_LABELS = {"电话", "手机", "电话 phone"}
EMAIL_LABELS = {"email地址"}
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+")


def expected_numbers(text):
    """The E.164 numbers a phone cell plainly holds (independent of the reader): a Chinese mobile, a landline with its
    area code (with or without an extension after a dash), a 400/800 line, a number with its + prefix."""
    t = unicodedata.normalize("NFKC", text).translate(DASH)
    out = []
    for tok in re.split(r"[/,;、\n]|\s{2,}", t):
        tok = tok.strip()
        ext = re.fullmatch(r"(.*\d)\s*-\s*\d{1,5}", tok)
        for cand in ([tok] + ([ext.group(1)] if ext else [])):
            d = re.sub(r"[\s\-().]", "", cand)
            if m := re.fullmatch(r"(?:\+?86|0086)?(1[3-9]\d{9})", d):
                out.append("+86" + m.group(1))
            elif m := re.fullmatch(r"0(\d{9,11})", d):
                out.append("+86" + m.group(1))
            elif re.fullmatch(r"[48]00\d{7}", d):
                out.append("+86" + d)
            elif re.fullmatch(r"\+\d{8,15}", d):
                out.append(d)
            else:
                continue
            break
    out += ["+86" + m for m in re.findall(r"(?<!\d)(1[3-9]\d{9})(?!\d)", re.sub(r"[\s\-]", "", t))]
    return list(dict.fromkeys(out))


def person_key(name):
    return "".join(unicodedata.normalize("NFKC", name).casefold().split())


def canon_code(text) -> str:
    m = re.search(r"([A-Za-z]{2,4}\d{5,})", text)
    return m.group(1).upper() if m else ""


class World:
    """Everything read once: the oracle rows, the pre-merge rows, the merged records, the review items."""

    def __init__(self, sample_dir):
        self.paths = snapshot_accounts.files(sample_dir)
        self.kept = {os.path.basename(p) for p in unique_contents(self.paths, [])}
        self.orows = oracle_accounts.oracle(self.paths)
        self.rows = {}
        for p in unique_contents(self.paths, []):
            for sheet in load(p):
                layout, got, _, _ = accounts.read_accounts_sheet(sheet, p)
                for r in got:
                    self.rows[(os.path.basename(p), sheet.name, r.row)] = r
        self.result = accounts.read_paths(self.paths)
        self.owner = collections.defaultdict(list)  # (file, sheet, row) -> merged records that name it
        for rec in [*self.result.accounts, *self.result.leads]:
            for f, s, r, _ in rec.sources:
                self.owner[(os.path.basename(f), s, r)].append(rec)
        self.review = collections.defaultdict(list)
        for i in self.result.review:
            self.review[(os.path.basename(i.file), i.sheet, i.row)].append(i)


@pytest.fixture(scope="session")
def world(sample_dir):
    return World(sample_dir)


def data_cells(world):
    for r in world.orows:
        if r.kind == "data":
            for c in r.cells:
                yield r, c


def declared_drop(world, orow, cell):
    """The reason a data cell may be absent from every record, or None. Each reason is a narrow, stated rule."""
    label = lab(cell.label)
    if cell.file not in world.kept:
        return "identical copy of another workbook (read once)"
    if label in KEPT_OUT:
        return KEPT_OUT[label]
    if is_placeholder(cell.text):
        return "placeholder (empty stand-in)"
    row = world.rows.get(orow.key)
    if row is None:
        skips = [i.reason for i in world.review.get(orow.key, ()) if i.reason in ("no_name", "no_code")]
        if skips:
            return f"row skipped with a stated reason ({skips[0]})"
        return None
    layout = orow.layout
    if layout == "erp" and label == "助记码" and PINYIN_TAG.match(cell.text.strip()):
        return "pinyin mnemonic, not an English name"
    if label in ("客户编码", "customer code") and cell.text.strip().lower() == "new":
        return "code 'New': the customer has no ERP code yet"
    if layout == "erp" and label in ("地区名称", "省直辖市") and OTHER_WORDS.search(cell.text):
        return "area 'Others' has no place to file"
    if layout == "erp" and label == "地区名称" and re.search(r"(省|自治区)\s*[A-Za-z ]*$", cell.text) and \
            subset("".join(CJK.findall(cell.text)), row.province):
        return "area cell names the province: the province column holds it"
    if ID_LIKE.fullmatch(cell.text.replace(" ", "")) and any(
            i.reason == "id_number_in_text_field" for i in world.review.get(orow.key, ())):
        return "tax number typed into a name column: dropped and reported"
    if layout == "erp" and label == "所属行业" and DO_NOT_USE.search(cell.text):
        return "industry cell 'Do Not use': flag and note instead"
    if layout == "supporting" and label.startswith("fibro existing customer") and cell.text.strip().lower() in ("y", "n"):
        return "Y/N flag becomes `existing` and the lifecycle"
    if layout == "supporting" and cell.col and label.startswith("active status") and cell.text.strip().lower() in ("y", "n"):
        return None
    if layout == "supporting" and label.startswith("active status") and INSTRUCTION.match(cell.text):
        return "request to the list's owner in a Y/N cell, not data (reported as instruction_not_data; see D-ACC-09)"
    if layout == "effective" and label == "english name" and " " not in cell.text.strip() and row.cn:
        return "one-word pinyin tag beside a Chinese name, not an English name"
    return None


BOOKKEEPING = {"key", "source_id", "sources", "match", "lifecycle", "relationships", "queued", "origin", "layout", "kind"}


def content(rec):
    """The texts of a merged record that carry list content (not its keys, its source references or its flags)."""
    for f in rec.__dataclass_fields__:
        if f not in BOOKKEEPING:
            yield from strings(getattr(rec, f))


def pool(world, orow):
    """Texts of the final record(s) built from this row (account or lead) and the review items about it.
    The pre-merge row is deliberately not in the pool: a value must survive into the merged record."""
    out = []
    for rec in world.owner.get(orow.key, ()):
        out.append(list(content(rec)))
    out.append([s for i in world.review.get(orow.key, ()) for s in (i.value, *(c[0] for c in i.candidates if c))])
    return out


def violation(rule, orow, cell, field, got=""):
    return {"rule": rule, "file": orow.file, "sheet": orow.sheet, "row": orow.row, "coord": cell.coord if cell else "",
            "field": field, "source": cell.text if cell else "", "got": got}


@pytest.fixture(scope="session")
def violations(world):
    return collect_violations(world)


def collect_violations(world):
    out = []
    # ---- cell-lost / wrong-field ----
    for orow, cell in data_cells(world):
        if declared_drop(world, orow, cell):
            continue
        texts = pool(world, orow)
        label = lab(cell.label)
        fields = FIELDS.get(orow.layout, {}).get(label)
        row = world.rows.get(orow.key)
        if fields and row is not None:
            if label in ("客户编码", "customer code"):
                ok = canon_code(cell.text) == row.code
            elif label in INDUSTRY_LABELS:
                ok = any(subset(latin_tail(cell.text) or cell.text, str(getattr(row, f))) for f in fields)
            elif fields == ("first_order",):
                ok = bool(row.first_order) and subset(row.first_order.split("-")[0], cell.text)
            else:
                ok = subset(cell.text, " ".join(" ".join(strings(getattr(row, f, ""))) for f in fields))
            if not ok:
                out.append(violation("wrong-field", orow, cell, "/".join(fields),
                                     " | ".join(" ".join(strings(getattr(row, f, ""))) for f in fields)))
                continue
        forms = [cell.text]
        if label in PLACE_LABELS and CJK.search(cell.text):
            forms.append(CJK_ONLY(cell.text))  # declared: a bilingual place keeps at least its Chinese half (see D-ACC-06)
        if label in INDUSTRY_LABELS and re.search(r"[A-Za-z]{2}", cell.text):
            forms.append(latin_tail(cell.text))  # declared: a bilingual industry / region keeps its English half
        if label in ("客户编码", "customer code"):
            forms.append(canon_code(cell.text))  # declared: the ERP code is filed without its leading zeros ("00CR991038" = "CR991038")
        if label == "fibro tables category":
            forms.append(re.sub(r"(?i)^roto", "", cell.text))  # declared: Cutting = Rotocutting, Motion = Rotomotion
        if label.startswith("first order time"):
            forms.append(re.match(r"\d{4}", cell.text).group(0) if re.match(r"\d{4}", cell.text) else cell.text)  # declared: "2023.1" is January or October: only the year is kept
        every = [s for pl in texts for s in pl]

        def found(text):
            return any(subset(f, s) for f in [text, *forms] for s in every)
        lines = [ln for ln in cell.text.split("\n") if chars(ln)]
        if label in EMAIL_LABELS:  # an e-mail cell is split into its addresses and the text around them
            lines = [*EMAIL_RE.findall(cell.text), *([EMAIL_RE.sub(" ", cell.text)] if chars(EMAIL_RE.sub(" ", cell.text)) else [])]
        if label == "city" and not CJK.search(cell.text) and any(CJK.search(rec.city) and not re.search(r"[A-Za-z]", rec.city)
                                                                    for rec in world.owner.get(orow.key, ()) if hasattr(rec, "city")):
            continue  # declared: an English city beside a Chinese-only city of the same account cannot be compared (see D-ACC-06)
        # a multi-line cell may be split at its line breaks (names, aliases): every line must then be found by itself
        if not (found(cell.text) or (len(lines) > 1 and all(found(ln) for ln in lines))):
            out.append(violation("cell-lost", orow, cell, cell.label))
    # ---- mixed-cells: a scalar field of a row is the text of one cell ----
    for orow in world.orows:
        row = world.rows.get(orow.key)
        if orow.kind != "data" or row is None:
            continue
        for f in SCALARS:
            value = getattr(row, f)
            if not value or f == "first_order":
                continue
            if not any(subset(value, c.text) for c in orow.cells):
                out.append(violation("mixed-cells", orow, None, f, value))
    # ---- merge-loses-value ----
    for key, row in world.rows.items():
        recs = world.owner.get(key, ())
        if not recs:
            out.append({"rule": "untied-record", "file": key[0], "sheet": key[1], "row": key[2], "coord": "",
                        "field": "", "source": "row is in no account or lead", "got": ""})
            continue
        rec = recs[0]
        for f in SCALARS:
            value = getattr(row, f, "")
            if not value or not hasattr(rec, f):
                continue
            kept = getattr(rec, f)
            names = list(getattr(rec, "names", []))
            reported = [i for i in world.review.get(key, ()) if i.field == f]
            if subset(value, kept) or reported or (f in ("en", "cn") and any(subset(value, n) for n in names)):
                continue
            if f == "product_area" and value.casefold().removeprefix("roto") == kept.casefold().removeprefix("roto"):
                continue  # declared: Motion = Rotomotion, Cutting = Rotocutting
            if f in ("province", "city") and CJK.search(value) != CJK.search(kept):
                continue  # declared: a Chinese and an English spelling of a place cannot be compared (see D-ACC-06)
            if f in ("province", "city") and subset(CJK_ONLY(value), kept):
                continue  # declared: the kept value is the Chinese half of the bilingual place name
            out.append({"rule": "merge-loses-value", "file": key[0], "sheet": key[1], "row": key[2], "coord": "",
                        "field": f, "source": value, "got": kept})
    # ---- contacts of the account: the person, his numbers and addresses ----
    for orow in world.orows:
        if orow.kind != "data" or orow.file not in world.kept:
            continue
        cs = {lab(c.label): c for c in orow.cells}
        name_cell = next((c for l, c in cs.items() if l in CONTACT_LABELS), None)
        reach = [c for l, c in cs.items() if l in PHONE_LABELS | EMAIL_LABELS]
        if not (name_cell or reach):
            continue
        owners = world.owner.get(orow.key, ())
        leads = [r for r in owners if not hasattr(r, "contacts")]
        if leads and name_cell and not is_placeholder(name_cell.text):  # a lead keeps one contact and one phone cell
            if not any(subset(name_cell.text, l.contact) for l in leads):
                out.append(violation("contact-missing", orow, name_cell, "lead.contact", [l.contact for l in leads]))
            continue
        people = [p for r in owners if hasattr(r, "contacts") for p in r.contacts]
        mine = [p for p in people if name_cell and getattr(p, "name", None) and person_key(p.name) == person_key(name_cell.text)]
        if name_cell and not is_placeholder(name_cell.text) and not mine:
            out.append(violation("contact-missing", orow, name_cell, "contacts", [getattr(p, "name", p) for p in people]))
            continue
        for c in reach:
            want = expected_numbers(c.text) if lab(c.label) in PHONE_LABELS else []
            have = {n for p in mine for n in getattr(p, "phones", [])}
            for n in want:
                if mine and n not in have:
                    out.append(violation("number-not-placed", orow, c, "contacts.phones", sorted(have)))
            mails = EMAIL_RE.findall(c.text)
            have_mail = {e.lower() for p in mine for e in getattr(p, "emails", [])}
            for e in mails:
                if mine and e.lower() not in have_mail:
                    out.append(violation("email-not-placed", orow, c, "contacts.emails", sorted(have_mail)))
    # ---- a record is tied to its cells: its code, its people and its flag come from its own rows ----
    cells_of = {r.key: r for r in world.orows if r.kind == "data"}
    for a in world.result.accounts:
        src = [cells_of[(os.path.basename(f), sh, rw)] for f, sh, rw, _ in a.sources if (os.path.basename(f), sh, rw) in cells_of]
        if a.code and not any(canon_code(c.text) == a.code for r in src for c in r.cells if lab(c.label) in ("客户编码", "customer code")):
            out.append({"rule": "code-not-in-source", "file": os.path.basename(a.sources[0][0]), "sheet": a.sources[0][1],
                        "row": a.sources[0][2], "coord": "", "field": "code", "source": "no source row carries it", "got": a.code})
        names = {person_key(c.text) for r in src for c in r.cells if lab(c.label) in CONTACT_LABELS}
        digits = {re.sub(r"\D", "", c.text) for r in src for c in r.cells if lab(c.label) in PHONE_LABELS}
        for p in a.contacts:
            pname = getattr(p, "name", "")
            if person_key(pname) not in names:
                out.append({"rule": "contact-not-from-source", "file": os.path.basename(a.sources[0][0]), "sheet": a.sources[0][1],
                            "row": a.sources[0][2], "coord": "", "field": "contacts", "source": "no source row names him", "got": pname})
            for n in getattr(p, "phones", []):
                if not any(n.lstrip("+86").lstrip("0") in d or n[1:] in d for d in digits):
                    out.append({"rule": "contact-not-from-source", "file": os.path.basename(a.sources[0][0]), "sheet": a.sources[0][1],
                                "row": a.sources[0][2], "coord": "", "field": "contacts.phones",
                                "source": "no phone cell of the account's rows holds it", "got": n})
    for orow in world.orows:  # the Y/N flag of a hand-over list is the account's `existing` flag and its lifecycle
        if orow.kind == "data" and orow.layout == "supporting" and orow.file in world.kept:
            flag = next((c for c in orow.cells if lab(c.label).startswith("fibro existing customer")), None)
            row = world.rows.get(orow.key)
            if flag and row is not None and flag.text.strip().lower() in ("y", "n") and \
                    row.existing is not (flag.text.strip().lower() == "y"):
                out.append(violation("wrong-field", orow, flag, "existing", row.existing))
    # ---- a field of a merged account is the value one of its own rows says (not invented, not another company's) ----
    for a in world.result.accounts:
        rows = [world.rows[(os.path.basename(f), sh, rw)] for f, sh, rw, _ in a.sources if (os.path.basename(f), sh, rw) in world.rows]
        for f in SCALARS:
            value = getattr(a, f, "")
            if not value or any(chars(value) == chars(getattr(r, f, "")) for r in rows):
                continue
            if f in ("cn", "en") and a.match.get("method") and not a.code and value in (a.match.get("cn"), a.match.get("en")):
                continue  # declared: a company without code found in the master is filed under the master's spelling
            out.append({"rule": "field-not-from-source", "file": os.path.basename(a.sources[0][0]), "sheet": a.sources[0][1],
                        "row": a.sources[0][2], "coord": "", "field": f, "source": a.key, "got": value})
    # ---- the sheet of a hand-over list is a label ("Migrated"): it must reach the account's notes ----
    for orow in world.orows:
        if orow.kind == "data" and orow.layout == "supporting" and orow.file in world.kept and orow.key in world.rows \
                and orow.sheet.strip().casefold() != "supporting":  # the default group "Supporting" needs no label (see D-ACC-12)
            notes = [n for r in world.owner.get(orow.key, ()) for n in getattr(r, "notes", [])]
            if not any(orow.sheet.strip() in n for n in notes):
                out.append(violation("sheet-label-lost", orow, None, "notes", notes))
    # ---- untied records ----
    seen = {(r.file, r.sheet, r.row) for r in world.orows if r.kind == "data"}
    for rec in [*world.result.accounts, *world.result.leads]:
        if not rec.sources:
            out.append({"rule": "untied-record", "file": "", "sheet": "", "row": 0, "coord": "", "field": "",
                        "source": "no source", "got": getattr(rec, "key", "")})
        for f, s, r, _ in rec.sources:
            if (os.path.basename(f), s, r) not in seen:
                out.append({"rule": "untied-record", "file": os.path.basename(f), "sheet": s, "row": r, "coord": "",
                            "field": "", "source": "source row not a data row of the workbook", "got": rec.key})
    out += shape_violations(world)
    return out


def shape_violations(world):
    out = []

    def bad(rule, rec, field, value):
        f, s, r, _ = rec.sources[0]
        out.append({"rule": rule, "file": os.path.basename(f), "sheet": s, "row": r, "coord": "", "field": field,
                    "source": getattr(rec, "key", ""), "got": value})
    for a in world.result.accounts:
        if a.code and not re.fullmatch(r"[A-Z]{2,4}\d{5,}", a.code):
            bad("shape-code", a, "code", a.code)
        for f in ("province", "city"):
            v = getattr(a, f)
            if v and (re.search(r"\d|@", v) or COMPANY_SUFFIX.search(v)):
                bad("shape-place", a, f, v)
        for p in a.contacts:
            name = getattr(p, "name", p[0] if isinstance(p, tuple) else "")
            if not name or COMPANY_SUFFIX.search(name) or re.search(r"\d|@", name):
                bad("shape-person-name", a, "contacts", name)
            for n in getattr(p, "phones", []):
                if not E164.match(n):
                    bad("shape-phone", a, "contacts", n)
            for e in getattr(p, "emails", []):
                if not EMAIL.match(e):
                    bad("shape-email", a, "contacts", e)
    for l in world.result.leads:
        if l.phone and not re.fullmatch(r"[\d+\s\-()/.]+", l.phone):
            bad("shape-lead-phone", l, "phone", l.phone)
    return out


ALL_RULES = sorted({"cell-lost", "wrong-field", "mixed-cells", "merge-loses-value", "untied-record", "shape-code",
                    "shape-place", "shape-person-name", "shape-phone", "shape-email", "shape-lead-phone", "contact-missing",
                    "number-not-placed", "email-not-placed", "code-not-in-source", "contact-not-from-source", "sheet-label-lost", "field-not-from-source"} | set(KNOWN_BUGS))


def _param(rule):
    if rule in KNOWN_BUGS:
        return pytest.param(rule, marks=pytest.mark.xfail(strict=True, reason=KNOWN_BUGS[rule]))
    return rule


@pytest.mark.parametrize("rule", [_param(r) for r in ALL_RULES])
def test_rule_has_no_violations(rule, violations):
    bad = [v for v in violations if v["rule"] == rule]
    assert not bad, f"{len(bad)} break '{rule}':\n" + "\n".join(
        f"  {v['file'][:30]} {v['sheet'].strip()[:14]} row {v['row']} {v['coord']} [{v['field']}]: "
        f"{v['source'][:50]!r} -> {str(v['got'])[:50]!r}" for v in bad[:6])


def test_every_cell_of_every_workbook_is_checked(world):
    """No sampling: the oracle sees every workbook and every non-empty cell, and every data row is accounted for."""
    files = {os.path.basename(p) for p in world.paths}
    assert len(files) == 7 and {r.file for r in world.orows} == files
    data_rows = [r for r in world.orows if r.kind == "data"]
    assert len(data_rows) == 2006
    assert sum(len(r.cells) for r in world.orows) == 19571
    unread = [r.key for r in data_rows if r.key not in world.rows and
              not any(i.reason in ("no_name", "no_code") for i in world.review.get(r.key, ()))]
    assert not unread, f"{len(unread)} data rows were neither read nor skipped with a reason: {unread[:5]}"


def test_declared_drops_are_few_and_listed(world):
    """The declared rules must not swallow the data: report what each rule drops."""
    count = collections.Counter()
    for orow, cell in data_cells(world):
        why = declared_drop(world, orow, cell)
        if why:
            count[why] += 1
    total = sum(len(r.cells) for r in world.orows if r.kind == "data")
    assert count, "no declared drops at all is suspicious"
    # kept-out columns and placeholders are most of it; everything else must stay small
    other = sum(n for why, n in count.items() if why not in KEPT_OUT.values() and not why.startswith("placeholder"))
    assert other < total * 0.05, count.most_common(8)
