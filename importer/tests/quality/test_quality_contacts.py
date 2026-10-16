"""Quality checks for the contact-list reader (folder 06), run on every cell of every workbook.

The independent oracle (oracle_contacts.py, built from the cell ledger only) says which cells exist and what
header stands over each. These rules then check, for the whole population and without sampling:

  cell-lost            every non-empty data cell is found in a reader record (letters and digits, using
                       invariants.chars) or is dropped by a DECLARED rule below, with its reason
  wrong-field          a cell under a header whose field is known landed in that field
  mixed-cells          a text field holds the text of one cell, never of two cells joined
  number-not-placed    a plain phone number of a cell is among the person's E.164 numbers (several are kept)
  email-not-placed     an address of an e-mail cell is among the person's addresses
  merge-loses-value    a value of a row survives the merge of copies (across files) or is reported
  shape-*              names, numbers and addresses look like what they are
  untied-record        every record names the file / sheet / row it came from, and that row exists

A rule listed in KNOWN_BUGS is a bug we know about and have not fixed: it is strict-xfail, so it keeps
failing loudly in the report, and the day it is fixed the marker must be removed.
"""
import collections
import os
import re
import unicodedata

import pytest

from excel import contacts
from excel.files import unique_contents
from excel.sheet import load
from tests.quality import oracle_contacts, snapshot_contacts
from tests.quality.invariants import chars, is_placeholder

KNOWN_BUGS = {
}

# header label (casefolded) -> record fields in which the cell's text must be found (the right CRM field)
FIELDS = {
    "list": {"公司名称": ("company",), "地址": ("address",), "联系人": ("full_name",), "电话": ("phones",)},
    "sections": {"contact person": ("full_name",), "e-mail addr.": ("emails",), "phone no.": ("phones",),
                 "position": ("role",), "remark": ("remark",), "": ("company", "department", "section")},
}
KEPT_OUT = {"序号": "row counter of the list"}
LEDGER_WORDS = {"应收": "customer", "预收": "customer", "其他应收": "customer", "应付": "supplier", "预付": "supplier",
                "其他应付": "supplier", "客户": "customer", "供应商": "supplier"}
COMPANY_SUFFIX = re.compile(r"有限公司|有限责任公司|股份|\bco\.?,? ?ltd\b|\bltd\b|\bgmbh\b|\binc\b|\bcorp\b|\bllc\b|集团", re.I)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+")
EMAIL_FULL = re.compile(r"^[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$")
E164 = re.compile(r"^\+\d{8,15}$")
DASH = dict.fromkeys(map(ord, "‐‑‒–—―−"), "-")
SCALARS = ("company", "full_name", "role", "remark", "department", "section", "address")


def lab(text) -> str:
    return unicodedata.normalize("NFKC", " ".join(str(text).split())).casefold()


def fold(text) -> str:
    return "".join(unicodedata.normalize("NFKC", str(text)).casefold().split())


def strings(obj):
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


def subset(small, big) -> bool:
    return not (chars(small) - chars(big))


def expected_numbers(text):
    """The E.164 numbers a phone cell plainly holds: a Chinese mobile, a landline with its area code, a number with its
    + prefix (a leading T / Tel label ignored). Independent of the reader."""
    t = unicodedata.normalize("NFKC", text).translate(DASH)
    out = []
    for tok in re.split(r"[/,;、\n]|\s{2,}", t):
        tok = re.sub(r"^\s*(?:tel|t|m|mob|mobile|phone)\b[.:]?\s*", "", tok.strip(), flags=re.I)
        ext = re.fullmatch(r"(.*\d)\s*-\s*\d{1,5}", tok)
        for cand in ([tok] + ([ext.group(1)] if ext else [])):
            d = re.sub(r"[\s\-().]", "", cand)
            if m := re.fullmatch(r"(?:\+?86|0086)?(1[3-9]\d{9})", d):
                out.append("+86" + m.group(1))
            elif m := re.fullmatch(r"0(\d{9,11})", d):
                out.append("+86" + m.group(1))
            elif re.fullmatch(r"\+\d{8,15}", d):
                out.append(d)
            else:
                continue
            break
    out += ["+86" + m for m in re.findall(r"(?<!\d)(1[3-9]\d{9})(?!\d)", re.sub(r"[\s\-]", "", t))]
    return list(dict.fromkeys(out))


class World:
    def __init__(self, sample_dir):
        self.paths = snapshot_contacts.files(sample_dir)
        self.kept = {os.path.basename(p) for p in unique_contents(self.paths, [])}
        self.orows = oracle_contacts.oracle(self.paths)
        self.entries = {}
        for p in unique_contents(self.paths, []):
            for sheet in load(p):
                layout, got, _, _ = contacts.read_contacts_sheet(sheet, p)
                for e in got:
                    self.entries[(os.path.basename(p), sheet.name, e.row)] = e
        self.result = contacts.read_paths(self.paths)
        self.review = collections.defaultdict(list)
        for i in self.result.review:
            self.review[(os.path.basename(i.file), i.sheet, i.row)].append(i)
        # company name as written -> org(s); (org key, person) -> merged contact
        self.org_of_name = collections.defaultdict(list)
        for k, o in self.result.orgs.items():
            for n in o.written_names:
                self.org_of_name[fold(n)].append(o)
        self.person = collections.defaultdict(list)
        for c in self.result.contacts:
            self.person[(c.company_key, fold(c.full_name))].append(c)

    def contacts_of(self, key):
        """The merged contacts a pre-merge entry was folded into."""
        e = self.entries.get(key)
        if e is None or not e.full_name:
            return []
        out = []
        for o in self.orgs_of(key):
            out += self.person.get((o.key, fold(e.full_name)), [])
        return out

    def orgs_of(self, key):
        e = self.entries.get(key)
        if e is None:
            return []
        return list(self.org_of_name.get(fold(e.company), [])) + list(self.org_of_name.get(fold(e.company + contacts.INTERNAL_SUFFIX), []))


@pytest.fixture(scope="session")
def world(sample_dir):
    return World(sample_dir)


def data_cells(world):
    for r in world.orows:
        if r.kind == "data":
            for c in r.cells:
                yield r, c


def declared_drop(world, orow, cell):
    """The reason a data cell may be absent from every record, or None."""
    label = lab(cell.label)
    if cell.file not in world.kept:
        return "identical copy of another workbook (read once)"
    if label in KEPT_OUT:
        return KEPT_OUT[label]
    if is_placeholder(cell.text) or cell.text.strip() == "无":
        return "placeholder (empty stand-in)"
    if orow.layout == "list" and label == "" and cell.text.strip() in LEDGER_WORDS:
        return "ledger / kind word: becomes the ledger type and the customer / supplier kind"
    return None


def pool(world, orow):
    """Texts of the final records about this row: the merged contact(s), the companies, the review items."""
    out = []
    for c in world.contacts_of(orow.key):
        out.append([s for f in c.__dataclass_fields__ if f not in ("source_id", "company_key", "file", "sheet", "layout", "kind", "copies")
                    for s in strings(getattr(c, f))])
    for o in world.orgs_of(orow.key):
        out.append([s for f in ("name", "address", "location", "written_names", "ledgers") for s in strings(getattr(o, f))])
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
    for orow, cell in data_cells(world):
        why = declared_drop(world, orow, cell)
        if why:
            continue
        label = lab(cell.label)
        e = world.entries.get(orow.key)
        fields = FIELDS.get(orow.layout, {}).get(label)
        if fields and e is not None:
            if label in ("phone no.", "电话"):
                ok = True  # checked below by number
            elif label == "e-mail addr.":
                ok = True
            else:
                ok = subset(cell.text, " ".join(" ".join(strings(getattr(e, f, ""))) for f in fields)) or \
                    (label == "" and orow.layout == "sections" and subset(contacts.company_from_label(cell.text), e.company))
            if not ok:
                out.append(violation("wrong-field", orow, cell, "/".join(fields),
                                     " | ".join(" ".join(strings(getattr(e, f, ""))) for f in fields)))
                continue
        every = [s for pl in pool(world, orow) for s in pl]
        lines = [ln for ln in cell.text.split("\n") if chars(ln)]
        if label in ("e-mail addr.",):
            lines = [*EMAIL_RE.findall(cell.text), *([EMAIL_RE.sub(" ", cell.text)] if chars(EMAIL_RE.sub(" ", cell.text)) else [])]

        def found(text):
            return any(subset(text, s) for s in every)
        if label in ("phone no.", "电话"):
            # a phone cell: its numbers are the person's E.164 numbers, or the cell is reported (phone_no_country)
            reported = any(i.field == "phone" for i in world.review.get(orow.key, ()))
            if not (found(cell.text) or reported or expected_numbers(cell.text)):
                out.append(violation("cell-lost", orow, cell, cell.label))
            continue
        if not (found(cell.text) or (len(lines) > 1 and all(found(ln) for ln in lines))):
            if label == "" and orow.layout == "sections" and not world.contacts_of(orow.key) and \
                    found(contacts.company_from_label(cell.text)):
                continue  # declared: a supplier row with no person is filed under the company its label names; what it supplies has no CRM field (D-CON-03)
            out.append(violation("cell-lost", orow, cell, cell.label))
    # ---- numbers and addresses of each person ----
    for orow in world.orows:
        if orow.kind != "data" or orow.file not in world.kept:
            continue
        cs = {lab(c.label): c for c in orow.cells}
        people = world.contacts_of(orow.key)
        for c in orow.cells:
            if lab(c.label) in ("phone no.", "电话") and people:
                have = {n for p in people for n in p.phones}
                for n in expected_numbers(c.text):
                    if n not in have:
                        out.append(violation("number-not-placed", orow, c, "phones", sorted(have)))
            if lab(c.label) in ("phone no.", "电话"):
                # every number of the cell is placed, or reported as a number the reader would not guess at
                have = {re.sub(r"\D", "", n) for p in people for n in p.phones}
                said = [re.sub(r"\D", "", i.value) for i in world.review.get(orow.key, ()) if i.field == "phone"]
                for tok in re.split(r"[/,;、\n]|\s{2,}", unicodedata.normalize("NFKC", c.text).translate(DASH)):
                    d = re.sub(r"\D", "", re.sub(r"^\s*(?:tel|t|m|mob|mobile|phone)\b[.:]?", "", tok.strip(), flags=re.I))
                    if len(d) >= 6 and not (any(h.endswith(d.lstrip("0")) for h in have) or any(d in x for x in said)):
                        out.append(violation("number-unaccounted", orow, c, "phones", tok))
            if lab(c.label) == "e-mail addr." and people:
                have = {e.lower() for p in people for e in p.emails}
                for e in EMAIL_RE.findall(c.text):
                    if e.lower() not in have:
                        out.append(violation("email-not-placed", orow, c, "emails", sorted(have)))
    # ---- mixed-cells: a scalar field of an entry is the text of one cell ----
    for orow in world.orows:
        e = world.entries.get(orow.key)
        if orow.kind != "data" or e is None:
            continue
        for f in ("full_name", "role", "remark", "address"):
            value = getattr(e, f)
            if value and not any(subset(value, c.text) for c in orow.cells):
                out.append(violation("mixed-cells", orow, None, f, value))
    # ---- merge-loses-value: a copy's value survives the merge, or is reported ----
    for key, e in world.entries.items():
        merged = world.contacts_of(key)
        if e.full_name and not merged:
            out.append({"rule": "untied-record", "file": key[0], "sheet": key[1], "row": key[2], "coord": "", "field": "",
                        "source": "row's person is in no merged contact", "got": e.full_name})
            continue
        for c in merged[:1]:
            for f in ("role", "address", "remark", "department"):
                v, kept = getattr(e, f), getattr(c, f)
                if v and not subset(v, kept) and not any(i.field == f for i in world.review.get(key, ())):
                    out.append({"rule": "merge-loses-value", "file": key[0], "sheet": key[1], "row": key[2], "coord": "",
                                "field": f, "source": v, "got": kept})
            for f in ("emails", "phones"):
                lost = [x for x in getattr(e, f) if x.lower() not in {y.lower() for y in getattr(c, f)}]
                if lost:
                    out.append({"rule": "merge-loses-value", "file": key[0], "sheet": key[1], "row": key[2], "coord": "",
                                "field": f, "source": str(lost), "got": str(getattr(c, f))})
    # ---- every contact is tied to a data row of a workbook ----
    data_rows = {r.key for r in world.orows if r.kind == "data"}
    for c in world.result.contacts:
        if (c.file, c.sheet, c.row) not in data_rows:
            out.append({"rule": "untied-record", "file": c.file, "sheet": c.sheet, "row": c.row, "coord": "", "field": "",
                        "source": "source row is not a data row", "got": c.full_name})
        if c.company_key not in world.result.orgs:
            out.append({"rule": "untied-record", "file": c.file, "sheet": c.sheet, "row": c.row, "coord": "", "field": "company",
                        "source": "company of the contact is unknown", "got": c.company_key})
    # ---- shapes ----
    for c in world.result.contacts:
        def bad(rule, field, value):
            out.append({"rule": rule, "file": c.file, "sheet": c.sheet, "row": c.row, "coord": "", "field": field,
                        "source": c.source_id, "got": value})
        if COMPANY_SUFFIX.search(c.full_name) or re.search(r"\d|@", c.full_name):
            bad("shape-person-name", "full_name", c.full_name)
        for n in c.phones:
            if not E164.match(n):
                bad("shape-phone", "phones", n)
        for m in c.emails:
            if not EMAIL_FULL.match(m):
                bad("shape-email", "emails", m)
        if len(set(c.phones)) != len(c.phones) or len({m.lower() for m in c.emails}) != len(c.emails):
            bad("shape-duplicate-entry", "phones/emails", str(c.phones + c.emails))
    for k, o in world.result.orgs.items():
        if not o.name.strip() or re.search(r"@", o.name):
            out.append({"rule": "shape-company", "file": "", "sheet": "", "row": 0, "coord": "", "field": "name",
                        "source": k, "got": o.name})
    return out


ALL_RULES = sorted({"cell-lost", "wrong-field", "mixed-cells", "merge-loses-value", "untied-record", "number-not-placed",
                    "email-not-placed", "number-unaccounted", "shape-person-name", "shape-phone", "shape-email", "shape-duplicate-entry",
                    "shape-company"} | set(KNOWN_BUGS))


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
    files = {os.path.basename(p) for p in world.paths}
    assert len(files) == 4 and {r.file for r in world.orows} == files
    assert sum(len(r.cells) for r in world.orows) == 297
    data_rows = [r for r in world.orows if r.kind == "data"]
    assert len(data_rows) == 57
    unread = [r.key for r in data_rows if r.key not in world.entries]
    assert not unread, f"{len(unread)} data rows were not read: {unread[:5]}"


def test_headings_are_kept_as_sections(world):
    """A row with only a column-A cell is a heading: its text is the `section` of the rows below or a reported label."""
    heading = [r for r in world.orows if r.kind == "heading"]
    assert heading
    for r in heading:
        text = r.cells[0].text
        kept = any(i.reason == "section" and i.value == contacts.cell_text(text) for i in world.review.get(r.key, ()))
        assert kept, f"heading not reported: {r.key} {text!r}"
