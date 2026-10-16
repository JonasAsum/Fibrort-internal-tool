"""The company master read from the CRM, so a visit's company is matched against what the CRM holds now.

The CRM keeps a company's name in `display_name` and the same name in the other language in `legal_name`, the
ERP code and (when the field exists) the USCC in custom fields. Short names and other spellings are not stored in
the CRM: they are taken from the customer list files for the companies the CRM already holds, by ERP code.
The CRM is the source: where it holds a company (by ERP code, or by an identical name) its record wins. A company
that is only in the files stays matchable until it has been imported, so a CRM that holds only a few companies does not
turn every visit into a new company; CRM_MASTER_ONLY=1 drops those. When the CRM cannot be read, or holds no companies,
the caller falls back to the files (Master.load).
"""
from __future__ import annotations

import os
import time

from .companies import CJK_RE, Master, USCC_RE, _canon_code, norm

CODE_LABEL = "ERP customer code"
USCC_LABEL = "Unified social credit code"
CACHE_SECONDS = 60  # answering several questions re-reads the upload; the CRM is not asked every time
_CACHE: dict = {"at": 0.0, "records": None}


def _custom_columns(client) -> dict:
    out = {}
    for f in client.custom_fields("company"):
        label = (f.get("label") or "").strip().lower()
        if label == CODE_LABEL.lower():
            out["code"] = f["column_name"]
        elif label == USCC_LABEL.lower():
            out["uscc"] = f["column_name"]
    return out


def record_of(company: dict, columns: dict) -> dict:
    """A master record for one CRM company: Chinese name, other-language name, further names as the short name."""
    names = [n.strip() for n in (company.get("display_name"), company.get("legal_name")) if n and n.strip()]
    cn = next((n for n in names if CJK_RE.search(n)), "")
    en = next((n for n in names if not CJK_RE.search(n)), "")
    rest = [n for n in names if n not in (cn, en)]
    code = str(company.get(columns.get("code")) or "").strip() if columns.get("code") else ""
    uscc = str(company.get(columns.get("uscc")) or "").strip().upper() if columns.get("uscc") else ""
    return {"source": "crm", "code": code or f"crm:{company['id']}", "cn": cn, "en": en,
            "short": rest[0] if rest else "", "uscc": uscc if USCC_RE.fullmatch(uscc) else ""}


def crm_records(make_client, use_cache: bool = True) -> list[dict]:
    """The CRM's companies as master records. `make_client` is only called (a login) when the last read is older
    than CACHE_SECONDS, so answering several questions in a row does not log in again each time."""
    if use_cache and _CACHE["records"] is not None and time.time() - _CACHE["at"] < CACHE_SECONDS:
        return _CACHE["records"]
    client = make_client()
    columns = _custom_columns(client)
    records = [record_of(c, columns) for c in client.paginate("/v1/companies", {"limit": 200})]
    if records:  # an empty CRM is asked again next time
        _CACHE.update(at=time.time(), records=records)
    return records


def load(make_client, importer_dir) -> Master:
    """The master of the CRM's companies. Raises when the CRM cannot be read (the caller falls back to the files)."""
    crm = crm_records(make_client)
    files, skipped, _ = Master.read_files(importer_dir)
    held = {_canon_code(r["code"]): r["code"] for r in crm if not r["code"].startswith("crm:")}
    by_name = {norm(n): r["code"] for r in crm if r["code"].startswith("crm:") for n in (r["cn"], r["en"]) if norm(n)}
    rows, only = [], 0
    for r in files:
        if not r.get("code"):
            continue
        code = held.get(_canon_code(r["code"])) or next((by_name[k] for k in (norm(r.get("cn")), norm(r.get("en"))) if k in by_name), None)
        if code is None:  # not in the CRM yet
            only += 1
            if os.environ.get("CRM_MASTER_ONLY") == "1":
                continue
            code = r["code"]
        rows.append({**r, "source": "erp", "code": code})
    master = Master.from_records(crm + rows)
    crm_codes = {r["code"] for r in crm}
    master.sources = {"crm": sum(1 for c in master.companies if c.code in crm_codes),
                      "lists": len({r["code"] for r in rows if r["code"] in crm_codes}),  # CRM companies the lists added to
                      "lists_only": sum(1 for c in master.companies if c.code not in crm_codes)}
    master.skipped = skipped
    master.origin = "crm"
    return master.with_decisions(importer_dir)
