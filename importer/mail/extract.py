"""Layer 1: archive -> staging/extracted.json (deterministic, no network, no AI).

    .venv/bin/python -m mail.extract                     # whole archive
    .venv/bin/python -m mail.extract --only Acme      # folders whose path contains this text
"""
import argparse
import re
import json
import os
import sys
from collections import Counter

from . import checks
from . import companies as companies_mod
from . import documents
from . import people
from . import relevance
from .reader import read_corpus

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_ROOT = os.path.join(HERE, "..", "..", "Email Data", "6vugs2jt.default-release", "ImapMail", "outlook.office365.com")


def finalize(msgs, contacts, folder_companies, comps, rows) -> list[dict]:
    """Applies the relevance decisions: company names onto contacts and messages, a keep flag on every record."""
    by_folder_name = {co["folder_name"]: co for co in comps.values() if co.get("folder_name")}
    for key, co in comps.items():
        r = rows.get(key)
        if key.startswith("folder:"):  # a folder with no outside domain: you filed it, so it is kept
            r = rows.setdefault(key, {"key": key, "keep": True, "reason": "filed folder without a domain",
                                      "relationship": relevance.TYPE_TO_REL.get(co.get("customer_type"), "Customer"),
                                      "evidence": {}, "subjects": [], "overridden": False})
        co["keep"] = bool(r and r["keep"])
        co["relationship"] = r["relationship"] if r else None
        if re.match(r"(?i)fibro|laepple|läpple", co["name"] or ""):  # the FIBRO USA folder is our own group
            co["relationship"], co["group"] = "FIBRO group", True
            if r:
                r["relationship"] = "FIBRO group"
        co["reason"] = r["reason"] if r else "no mail from or to this company"
    internal = [{**fc, "key": "internal:" + fc["name"], "keep": True, "relationship": "FIBRO group",
                 "reason": "our own company", "legal_name_en": None, "name_conflicts": {}, "name_evidence": {},
                 "group": True} for fc in folder_companies if fc.get("customer_type") == "internal"]
    name_of_key = {k: co["name"] for k, co in comps.items()}
    for c in contacts:
        if c["internal"]:
            c["keep"] = True
            continue
        if c["company_source"] == "folder" and c["company"] in by_folder_name:
            co = by_folder_name[c["company"]]
            c["company"], c["keep"] = co["name"], co["keep"]
            continue
        k = c["org"] or c["key"]
        r = rows.get(k)
        c["keep"] = bool(r and r["keep"])
        c["company"] = name_of_key.get(c["org"]) if c["org"] else c["company"]
        c["company_source"] = "domain" if c["org"] else c["company_source"]
    for c in contacts:  # a nameless address that never wrote to us (CC on a distribution) is no contact
        if c["keep"] and c["mailbox"] and not c["sent"] and not c["internal"]:
            c["keep"] = False
    kept_keys = {k for k, r in rows.items() if r["keep"]}
    kept_names = {co["name"] for co in comps.values() if co["keep"]}
    for m in msgs:
        parts = [p["email"] for p in [m["from"]] + m["to"] + m["cc"] if p["email"]]
        ks = [relevance.key_of(a) for a in parts]
        ks = [k for k in ks if k]
        fc = by_folder_name.get(m["folder"].get("company"))
        names = [name_of_key.get(k) for k in ks if k in kept_keys and name_of_key.get(k)]
        m["company"] = fc["name"] if fc else (Counter(names).most_common(1)[0][0] if names else None)
        mass = len({a.split("@")[-1] for a in parts if relevance.key_of(a)}) >= relevance.MASS_DOMAINS
        if fc or m["folder"]["customer_type"] == "internal":
            m["import"] = True  # you filed it
        elif mass or (m.get("bulk") and m["direction"] == "inbound"):
            m["import"] = False
        else:  # internal-only mail stays; mail only with dropped outsiders goes
            m["import"] = not ks or any(k in kept_keys for k in ks)
        if m["company"] and m["company"] not in kept_names:
            m["company"] = None
    return sorted(list(comps.values()) + internal, key=lambda c: (not c["keep"], c["name"].lower()))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--root", default=DEFAULT_ROOT, help="the Thunderbird ImapMail account directory")
    ap.add_argument("--staging", default=os.path.join(HERE, "staging"))
    ap.add_argument("--only", help="only mailbox folders whose path contains this text (case-insensitive)")
    args = ap.parse_args()

    files_dir = os.path.join(args.staging, "files")
    os.makedirs(args.staging, exist_ok=True)
    corpus = read_corpus(os.path.abspath(args.root), files_dir, args.only)
    msgs = corpus.messages
    print(f"{corpus.stats['read']} messages read, {len(msgs)} unique "
          f"({corpus.stats['duplicate_across_folders']} copies across folders), "
          f"{corpus.stats['parse_errors']} parse errors")
    for e in corpus.errors[:10]:
        print("  ", e, file=sys.stderr)

    contacts, folder_companies = people.build_people(msgs)

    docs: dict[str, dict] = {}
    for m in msgs:
        for a in m["attachments"]:
            if a["stored_as"] and a["filename"].lower().endswith(".pdf") and a["sha256"] not in docs:
                docs[a["sha256"]] = {"filename": a["filename"], "stored_as": a["stored_as"],
                                     **documents.read_document(os.path.join(files_dir, a["stored_as"]))}
    for m in msgs:
        m["refs"] = documents.find_refs(m["subject"], m["top"])
        for a in m["attachments"]:
            d = docs.get(a["sha256"])
            if d:
                for kind, vals in d["refs"].items():
                    m["refs"].setdefault(kind, [])
                    m["refs"][kind] = sorted(set(m["refs"][kind]) | set(vals))

    overrides = {}
    dp = os.path.join(args.staging, "decisions.json")
    if os.path.exists(dp):
        with open(dp) as f:
            overrides = json.load(f)
    comps = companies_mod.build_companies(msgs, contacts, folder_companies, docs)
    for key, o in overrides.items():  # "same_as": two records that are one company (a spelling variant)
        target = o.get("same_as") if isinstance(o, dict) else None
        if target and key in comps and target in comps:
            src, dst = comps.pop(key), comps[target]
            dst["parties"] = sorted(set(dst.get("parties") or []) | set(src.get("parties") or []))
            for f in ("legal_name", "legal_name_en"):
                dst[f] = dst.get(f) or src.get(f)
            dst["domains"] = sorted(set(dst["domains"]) | set(src["domains"]))
    for key, o in overrides.items():  # your corrections win over the extracted names
        if key in comps and isinstance(o, dict):
            for field in ("name", "legal_name", "legal_name_en", "parent"):
                if o.get(field):
                    comps[key][field] = o[field]
    rows = relevance.decide(msgs, comps, overrides)
    companies = finalize(msgs, contacts, folder_companies, comps, rows)
    relevance.write_csv(rows, comps, os.path.join(args.staging, "relevance.csv"))

    out = {
        "summary": {
            "messages": len(msgs), "contacts": len(contacts), "companies": len(companies),
            "documents_read": len(docs), "documents_parsed": sum(1 for d in docs.values() if "parsed" in d),
            "documents_unreadable": sum(1 for d in docs.values() if not d["readable"]),
            "by_folder_type": dict(Counter(m["folder"]["customer_type"] for m in msgs)),
            "contacts_with_phone": sum(1 for c in contacts if c["phones"]),
            "companies_kept": sum(1 for c in companies if c["keep"]),
            "companies_dropped": sum(1 for c in companies if not c["keep"]),
            "contacts_kept": sum(1 for c in contacts if c["keep"]),
            "mailbox_contacts": sum(1 for c in contacts if c["keep"] and c["mailbox"]),
            "messages_kept": sum(1 for m in msgs if m["import"]),
            "name_conflicts": sum(1 for c in companies if c["keep"] and any(c["name_conflicts"].values())),
            "parse_errors": corpus.errors,
        },
        "companies": companies, "contacts": contacts, "documents": docs, "messages": msgs,
        "relevance": rows,
    }
    path = os.path.join(args.staging, "extracted.json")
    with open(path, "w") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(json.dumps(out["summary"], ensure_ascii=False, indent=1))
    print("wrote", path, "and", os.path.join(args.staging, "relevance.csv"))
    failed = checks.run(out)
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
