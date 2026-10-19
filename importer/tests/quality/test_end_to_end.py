"""L5: the whole folder goes into the in-memory CRM, then every phone number, e-mail address and company the CRM
holds is traced back to a cell of the Excel files (nothing invented), and the contacts the visit sheets name are
looked up in the CRM (nothing lost). The CRM here is tests/fake_site_crm.py, driven through the same web endpoints the
page uses; the real CRM is never touched.
"""
import json
import os
import re
import unicodedata

import pytest

from tests.quality.ledger import cells as ledger_cells
from tests.site_fixture import site  # noqa: F401  (pytest fixture)
from tests.test_all_files_matrix import SITE_FILES, real_master  # noqa: F401

SYSTEM_COMPANIES = {"FIBRO Rundtische GmbH (internal)"}  # the importer's own label for internal meetings, not in any file
DIGITS = re.compile(r"\d{7,}")
EMAIL = re.compile(r"[A-Za-z0-9._%+'-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+")


def _key(text):
    return re.sub(r"[\W_]+", "", unicodedata.normalize("NFKC", str(text or "")).casefold())


def _tail(digits: str) -> str:
    """A phone's national digits: no 0086 / 86 prefix, no trunk 0."""
    d = re.sub(r"\D", "", digits)
    for prefix in ("0086", "86"):
        if d.startswith(prefix) and len(d) > 9:
            d = d[len(prefix):]
            break
    return d.lstrip("0")


def _excel_facts():
    """Everything the 59 workbooks say, read with the independent ledger: all digit runs, all e-mail addresses, all text."""
    digit_runs, emails, text = set(), set(), []
    for _, _, path in SITE_FILES:
        try:
            got = list(ledger_cells(path))
        except Exception:  # noqa: BLE001  (an unreadable workbook has no facts; the matrix test reports it)
            continue
        for c in got:
            for m in DIGITS.finditer(re.sub(r"[\s\-()（）.]", "", c.text)):
                digit_runs.add(_tail(m.group()))
            emails.update(e.lower() for e in EMAIL.findall(c.text))
            text.append(_key(c.text))
    return digit_runs, emails, "".join(text)


def test_nothing_in_the_crm_is_invented(site, real_master):  # noqa: F811
    if not SITE_FILES:
        pytest.skip("sample folder not found")
    site.upload(*[(absolute, rel) for rel, _, absolute in SITE_FILES])
    job = site.commit()
    assert job["status"] == "done"
    crm = site.crm
    if os.environ.get("E2E_DUMP"):  # keep the imported CRM for study without running the import again
        json.dump({"companies": crm.companies, "contacts": crm.contacts, "employment": crm.employment},
                  open(os.environ["E2E_DUMP"], "w", encoding="utf-8"), ensure_ascii=False, default=str)
    digit_runs, emails, all_text = _excel_facts()
    assert len(crm.contacts) > 100 and len(crm.companies) > 100, (len(crm.contacts), len(crm.companies))

    bad = []
    n_phones = n_emails = 0
    for cid, c in crm.contacts.items():
        for p in c.get("phones") or []:
            number = p["phone"] if isinstance(p, dict) else str(p)
            n_phones += 1
            if not re.fullmatch(r"\+\d{8,15}", number):
                bad.append(("phone not E.164", c.get("full_name"), number))
            elif _tail(number) not in digit_runs and not any(_tail(number) in r or r in _tail(number) for r in digit_runs):
                bad.append(("phone not in any Excel cell", c.get("full_name"), number))
        for e in c.get("emails") or []:
            address = (e["email"] if isinstance(e, dict) else str(e)).lower()
            n_emails += 1
            if address not in emails:
                bad.append(("email not in any Excel cell", c.get("full_name"), address))
    for cid, co in crm.companies.items():
        name = co.get("display_name") or ""
        if name and _key(name) not in all_text and name not in SYSTEM_COMPANIES:
            bad.append(("company name not in any Excel cell", name, ""))
    # nothing lost: a contact a visit sheet names, and the CRM holds, carries every phone number the sheet gives
    from tests.quality import snapshot
    from tests.quality.ledger import Ledger
    by_name = {}
    for c in crm.contacts.values():
        by_name.setdefault(_key(c.get("full_name")), []).append({_tail(p["phone"] if isinstance(p, dict) else p)
                                                                 for p in c.get("phones") or []})
    paths = {os.path.basename(p): p for p in snapshot.visit_files()}
    ledgers, checked, absent, lost, skipped = {}, 0, 0, [], {"same name at several companies": 0, "not a valid mobile number": 0, "company waits for the owner's answer": 0}
    records = snapshot.load()
    companies_of = {}  # contact name -> the customer cells it appears under (a name at several companies is not one person)
    for key, rec in records.items():
        src = (rec.get("extra") or {}).get("source_cells") or {}
        if rec.get("type") == "visit" and "contact" in src and "customer" in src:
            base = key.split("|")[0]
            if base not in ledgers:
                ledgers[base] = Ledger(paths[base])
            led = ledgers[base]
            companies_of.setdefault(_key(led.text(rec["sheet"], src["contact"])), set()).add(_key(led.text(rec["sheet"], src["customer"])))
    for key, rec in records.items():
        src = (rec.get("extra") or {}).get("source_cells") or {}
        if rec.get("type") != "visit" or "contact" not in src or "phone" not in src:
            continue
        base = key.split("|")[0]
        if base not in ledgers:  # (setdefault would build a new ledger, and reload the workbook, for every visit)
            ledgers[base] = Ledger(paths[base])
        led = ledgers[base]
        name, phone_cell = led.text(rec["sheet"], src["contact"]), led.text(rec["sheet"], src["phone"])
        runs = [_tail(m.group()) for m in DIGITS.finditer(re.sub(r"[\s\-()（）.]", "", phone_cell))]
        if not name.strip() or not runs:
            continue
        runs = [r for r in runs if re.fullmatch(r"1[3-9]\d{9}", r)]  # only a valid mobile number is expected in the CRM
        if not runs:
            skipped["not a valid mobile number"] += 1
            continue
        if ((rec.get("extra") or {}).get("company_match") or {}).get("method") in ("review", "ambiguous"):
            skipped["company waits for the owner's answer"] += 1  # not imported until a person answers "same company?"
            continue
        if len(companies_of.get(_key(name), ())) > 1:
            skipped["same name at several companies"] += 1
            continue
        checked += 1
        holders = by_name.get(_key(name))
        if not holders:
            absent += 1  # that contact is not in the CRM: its company waits for an answer, or the contact is not a person
            continue
        for r in runs:
            if not any(r in numbers or any(r in n or n in r for n in numbers) for numbers in holders):
                lost.append((base, rec["sheet"], name, r))
    print(f"phones named by visit sheets: {checked} cells; contact in CRM: {checked - absent}; contact not imported: {absent}; phone lost: {len(lost)}; skipped: {skipped}")
    assert not lost, f"{len(lost)} phone numbers lost for contacts that are in the CRM:\n" + "\n".join(map(str, lost[:10]))
    print(f"\nCRM: {len(crm.companies)} companies, {len(crm.contacts)} contacts, {n_phones} phones, {n_emails} e-mails checked")
    assert not bad, f"{len(bad)} CRM values with no source cell:\n" + "\n".join(map(str, bad[:15]))
