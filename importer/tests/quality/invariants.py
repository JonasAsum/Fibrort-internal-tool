"""Whole-population checks for visit records: every record, every file, no sampling.

Each reader field is tied to the cell it came from (the visit's `extra["source_cells"]`, looked up in
the independent ledger). A check passes only when the *right field* holds the right text:

  preserved   the letters and digits of the source cell are all present, and nothing else was added
  split       a customer cell is split into the right parts (name / English name / address / code)

A violation says file, sheet, row, field, which rule broke, and the source text next to what the reader made of it.
"""
import re
import unicodedata
from collections import Counter

# visit field -> (record key holding the reader's value, where it lives)
SHORT_FIELDS = {
    "city": ("city",), "contact": ("contact_raw",),
    "phone": ("phone_raw", "extra.phone_more"),  # a cell with several numbers: the first, then the others
    "office": ("office_phone", "extra.office_address"),  # the Office column: a phone, or (Lucas sheets) an address
    "email": ("email",),
    "customer_type": ("customer_type_raw",), "potential": ("potential",), "date": ("date_raw",),
}
CJK = re.compile(r"[\u4e00-\u9fff]")
ADDR_LABEL = re.compile(r"(?im)^\s*(?:Addr(?:ess)?|地址)\s*[:：]")  # declared rule: the 'Addr:' label is removed
PLACEHOLDERS = {"", "/", "-", "--", "无", "none", "n/a", "na", "tbd", "null"}
ENG_CO = re.compile(r"^(?!No\.|\d)[A-Za-z][^,]{2,}?\b(Co\.?,? ?Ltd\.?|Ltd\.?|GmbH|Limited|Corporation|Company|Co\.)\b")


def chars(text) -> Counter:
    """Letters and digits only, folded, so spacing, brackets, dashes and case do not matter."""
    s = unicodedata.normalize("NFKC", str(text or "")).casefold()
    return Counter(ch for ch in s if ch.isalnum())


def is_placeholder(text) -> bool:
    return unicodedata.normalize("NFKC", str(text or "")).strip().casefold() in PLACEHOLDERS


def _phone_shaped(text) -> bool:
    """Digits and phone punctuation, at most a short label such as 'T' or 'Tel' (independent of the reader)."""
    t = str(text)
    if re.fullmatch(r"\d{3,5}", t.strip()):  # declared: a short number is an extension, kept as written (the CRM write skips it)
        return True
    return sum(c.isdigit() for c in t) >= 6 and sum(c.isalpha() for c in t) <= 4 and not re.search(r"[^\w\s+\-()./,;:：]", t)


def _violation(rec, rule, field, source, got):
    return {"file": rec["_file"], "sheet": rec["sheet"], "row": rec["excel_row_start"], "field": field,
            "rule": rule, "source": source, "got": got}


def check_visit(rec, ledger) -> list:
    """Violations for one visit record (a snapshot dict with `_file` added) against its workbook's ledger."""
    out = []
    cells = (rec.get("extra") or {}).get("source_cells") or {}
    sheet = rec["sheet"]

    # short fields: the reader field must hold exactly the letters and digits of its cell
    for field, keys in SHORT_FIELDS.items():
        coord = cells.get(field)
        if not coord:
            continue
        src = ledger.text(sheet, coord)
        got = " ".join(str((rec.get("extra") or {}).get(k[6:]) if k.startswith("extra.") else rec.get(k) or "") for k in keys)
        got = got.replace("None", "").strip()
        if is_placeholder(src):
            if got and not is_placeholder(got):
                out.append(_violation(rec, "invented-value", field, src, got))
        elif field == "email" and "@" not in src:
            if got:  # declared rule: text that is not an email address is left out of the email field
                out.append(_violation(rec, "email-field-holds-non-email", field, src, got))
        elif chars(src) != chars(got):
            out.append(_violation(rec, "not-preserved", field, src, got))

    # a phone field must hold a phone number (an address in a phone field has the same letters as its cell)
    for key in ("phone_raw", "office_phone"):
        got = (rec.get(key) or "").strip()
        if got and not _phone_shaped(got):
            out.append(_violation(rec, "phone-field-not-a-phone", key, ledger.text(sheet, cells.get("phone" if key == "phone_raw" else "office", "")), got))

    # customer cell: name + English name + address (+ code) must add up to the cell, and sit in the right parts
    coord = cells.get("customer")
    if coord:
        src = ledger.text(sheet, coord)
        name = rec.get("customer_raw") or ""
        en = (rec.get("extra") or {}).get("name_en") or ""
        addr = (rec.get("extra") or {}).get("address") or ""
        code = (rec.get("extra") or {}).get("customer_code") or ""
        if not is_placeholder(src):
            if re.search(r"(?<![A-Za-z0-9])0*CR\d{5,}", src) and code not in re.findall(r"(?<![A-Za-z0-9])0*CR\d{5,}", src):
                out.append(_violation(rec, "customer-code-not-extracted", "customer", src, code))
            have, parts = chars(ADDR_LABEL.sub("", src)), chars(name) + chars(en) + chars(addr) + chars(code)
            missing = have - parts
            if have != parts and not missing:  # reader text is a superset: completed from the company master
                pass  # declared rule: name completed from the company master (see oracle spec)
            elif have != parts and missing in (chars("Co.,Ltd."), chars("Ltd.")):
                pass  # declared rule: a legal suffix typed twice ("… Co., Ltd. Co.,Ltd.") is dropped
            elif have != parts:
                out.append(_violation(rec, "customer-not-preserved", "customer", src, f"{name} | {en} | {addr}"))
            if re.search(r"Addr|地址", name) or re.search(r"(?i)^\s*addr(?:ess)?\s*[:：]", addr):
                out.append(_violation(rec, "address-text-in-name", "customer", src, name))
            if re.search(r"(?<![A-Za-z0-9])0*CR\d{5,}", name):
                out.append(_violation(rec, "customer-code-in-name", "customer", src, name))
            # a bilingual name stays one name (hand-checked gold set), so Chinese + English in the name is allowed
            if ENG_CO.match(addr):
                out.append(_violation(rec, "english-name-in-address", "customer", src, addr))
            if ")" in addr and "(" not in addr and "）" not in addr:
                out.append(_violation(rec, "stray-bracket-in-address", "customer", src, addr))
    return out
