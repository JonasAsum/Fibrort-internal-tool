"""Case files for the reading layer: per company, everything a reader needs to understand its mails.

    .venv/bin/python -m mail.cases                     # every company, plus the unfiled threads
    .venv/bin/python -m mail.cases --company ZETA

Writes staging/reading/<case>/case.json (the data, with the real Message-IDs behind every alias) and case.md
(what a reader reads). Mails are cited as M1, M2 …, attachments as D1, D2 …; reading.py turns those back into
Message-IDs and file hashes and checks every quote against the original mail. Built from the code-only plan
(no enrichment), so the same extraction always gives the same case files.
"""
import argparse
import hashlib
import json
import os
import re
from collections import Counter, defaultdict

from .plan import build_plan, fold
from .reader import _QUOTE_START

HERE = os.path.dirname(os.path.abspath(__file__))
FORWARD = re.compile(r"^\s*(fw|fwd|wg|tr|转发|转寄)\s*[:：]", re.I)
HISTORY_CHARS = 4000  # quoted history shown for a thread's first mail and for forwards
EXCERPT_CHARS = 1500  # text of a readable PDF that code did not parse
UNFILED_CHARS = 45_000  # the unfiled threads are split into parts of about this much text
REF_TYPES = ("sales_order", "quotation", "order_confirmation", "purchase_order", "fibro_order_no", "supplier_contract")


def slug(name: str) -> str:
    s = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")
    return s or "c-" + hashlib.sha1(name.encode()).hexdigest()[:8]


def split_history(body: str) -> str:
    """The quoted part of a body: everything from the first quote marker on (top_reply keeps the rest)."""
    lines = body.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    for i, line in enumerate(lines):
        if i > 0 and _QUOTE_START.match(line):
            return re.sub(r"\n{3,}", "\n\n", "\n".join(lines[i:])).strip()
    return ""


def person(p: dict) -> str:
    return f"{p['name']} <{p['email']}>" if p.get("name") else (p.get("email") or "?")


def build_cases(extracted: dict) -> dict[str, dict]:
    plan = build_plan(extracted, None)
    docs = extracted["documents"]
    kept = sorted(c["name"] for c in plan.companies)
    contact_of = {e: c for c in plan.contacts for e in c["emails"]}

    threads = defaultdict(list)
    for m in plan.emails:
        threads[m["thread_key"]].append(m)
    order = sorted(threads, key=lambda k: (min(m["date_iso"] or "" for m in threads[k]), k))
    owner = {}
    for k in order:
        votes = Counter(m["company"] for m in threads[k] if m["company"])
        owner[k] = votes.most_common(1)[0][0] if votes else None

    # the unfiled threads, split into parts a reader can take in one go
    unfiled, part, size = {}, 1, 0
    for k in order:
        if owner[k]:
            continue
        chars = sum(len(m["top"]) for m in threads[k])
        if size and size + chars > UNFILED_CHARS:
            part, size = part + 1, 0
        unfiled[k] = f"_unfiled-{part}"
        size += chars

    cases: dict[str, dict] = {}
    for k in order:
        name = owner[k] or unfiled[k]
        case = cases.setdefault(name, {"case": name, "company": owner[k], "messages": {}, "documents": {},
                                       "threads": {}, "_doc_alias": {}})
        t_alias = f"T{len(case['threads']) + 1}"
        case["threads"][t_alias] = []
        for i, m in enumerate(sorted(threads[k], key=lambda m: (m["date_iso"] or "", m["message_id"]))):
            alias = f"M{len(case['messages']) + 1}"
            atts = []
            for a in m["attachments"]:
                if a["decoration"] or not a.get("stored_as"):
                    continue
                d_alias = case["_doc_alias"].get(a["sha256"])
                if not d_alias:
                    d_alias = case["_doc_alias"][a["sha256"]] = f"D{len(case['documents']) + 1}"
                    doc = docs.get(a["sha256"]) or {}
                    is_pdf = a["filename"].lower().endswith(".pdf")
                    case["documents"][d_alias] = {
                        "sha256": a["sha256"], "filename": a["filename"], "file": f"mail/staging/files/{a['stored_as']}",
                        "kind": ("scan" if is_pdf and doc and not doc.get("readable") else
                                 "pdf" if is_pdf and doc else "file"),
                        "parsed": doc.get("parsed"),
                        "excerpt": (doc.get("excerpt") or "")[:EXCERPT_CHARS] if doc.get("readable") else "",
                    }
                atts.append(d_alias)
            history = split_history(m["body"]) if i == 0 or FORWARD.match(m["subject"] or "") else ""
            case["messages"][alias] = {
                "message_id": m["message_id"], "thread": t_alias, "date": m["date_iso"], "direction": m["direction"],
                "from": person(m["from"]), "to": [person(p) for p in m["to"]], "cc": [person(p) for p in m["cc"]],
                "subject": m["subject"], "folder": m["folder"]["path"], "text": m["top"],
                "history": history[:HISTORY_CHARS], "history_cut": len(history) > HISTORY_CHARS,
                "attachments": atts, "refs": {t: v for t, v in (m.get("refs") or {}).items() if t in REF_TYPES},
            }
            case["threads"][t_alias].append(alias)

    for name, case in cases.items():
        del case["_doc_alias"]
        emails = {e for msg in case["messages"].values() for e in re.findall(r"<([^<>]+)>", " ".join(
            [msg["from"]] + msg["to"] + msg["cc"]))}
        people = [contact_of[e] for e in sorted(emails) if e in contact_of]
        people = list({c["key"]: c for c in people}.values())
        refs = defaultdict(set)
        for msg in case["messages"].values():
            for t, v in msg["refs"].items():
                refs[t] |= set(v)
        case["known"] = {
            "companies": kept,
            "contacts": [{"name": c["full_name"], "emails": c["emails"], "title": c.get("title"),
                          "company": c.get("company"), "internal": c["internal"]} for c in people],
            "deals_from_documents": [
                {"key": d["key"], "company": d["company"], "amount_minor": d["amount_minor"], "currency": d["currency"],
                 "documents": [x["filename"] for x in d["docs"]]}
                for d in plan.deals.values() if d["docs"] and (case["company"] is None or fold(d["company"] or "") == fold(case["company"]))],
            "numbers": {t: sorted(v) for t, v in sorted(refs.items())},
        }
        body = json.dumps({k: case[k] for k in ("messages", "documents", "known")}, ensure_ascii=False, sort_keys=True)
        case["digest"] = hashlib.sha256(body.encode()).hexdigest()[:16]
    return cases


def render(case: dict) -> str:
    """The case as a reader sees it."""
    out = [f"# Case: {case['case']}", "",
           "Read this with mail/READING.md. Cite a mail as M<n> and an attachment as D<n>. Quote word for word.", ""]
    k = case["known"]
    out += ["## Known facts", "", "Companies you may name: " + ", ".join(k["companies"]), ""]
    if k["contacts"]:
        out.append("People in these mails (the only names you may use):")
        for c in k["contacts"]:
            out.append(f"- {c['name']} <{', '.join(c['emails'])}>" + (f" — {c['title']}" if c["title"] else "")
                       + (f" — {c['company']}" if c["company"] else ""))
        out.append("")
    if k["deals_from_documents"]:
        out.append("Deals code already made from parsed contracts (refer to them by number, do not re-create):")
        for d in k["deals_from_documents"]:
            amount = f"{d['currency']} {d['amount_minor'] / 100:,.2f}" if d["amount_minor"] else "no amount"
            out.append(f"- {d['key']} — {d['company']} — {amount} — from {', '.join(d['documents'])}")
        out.append("")
    if k["numbers"]:
        out.append("Numbers code found in these mails: " + "; ".join(f"{t}: {', '.join(v)}" for t, v in k["numbers"].items()))
        out.append("")
    for t_alias, aliases in case["threads"].items():
        first = case["messages"][aliases[0]]
        out += [f"## {t_alias} — {first['subject'] or '(no subject)'} ({len(aliases)} mail{'s' if len(aliases) > 1 else ''})", ""]
        for alias in aliases:
            m = case["messages"][alias]
            out.append(f"### {alias} · {(m['date'] or '?')[:16].replace('T', ' ')} · {m['direction']} · from {m['from']}")
            out.append(f"To: {', '.join(m['to']) or '-'}" + (f" · Cc: {', '.join(m['cc'])}" if m["cc"] else ""))
            out.append(f"Subject: {m['subject']} · Folder: {m['folder']}")
            if m["attachments"]:
                out.append("Attachments: " + ", ".join(f"{d} {case['documents'][d]['filename']} ({case['documents'][d]['kind']})"
                                                     for d in m["attachments"]))
            out += ["", m["text"].strip() or "(no new text)", ""]
            if m["history"]:
                out += ["--- quoted history (older mails in this thread, as quoted below this one) ---", m["history"].strip()]
                if m["history_cut"]:
                    out.append("[… quoted history cut here …]")
                out.append("")
    if case["documents"]:
        out += ["## Attachments", ""]
        for d_alias, d in case["documents"].items():
            head = f"### {d_alias} · {d['filename']} · "
            if d["kind"] == "scan":
                out += [head + f"SCAN, no text layer: open {d['file']} to read it", ""]
            elif d["kind"] == "file":
                out += [head + "not read by code (not a PDF)", ""]
            else:
                out.append(head + ("read by code: " + json.dumps(d["parsed"], ensure_ascii=False) if d["parsed"] else "text excerpt:"))
                if d["excerpt"]:
                    out += ["", d["excerpt"].strip()]
                out.append("")
    return "\n".join(out) + "\n"


def write_cases(cases: dict[str, dict], staging: str, only: str | None = None) -> list[str]:
    written = []
    for name, case in cases.items():
        if only and fold(only) not in (fold(name), slug(name)):
            continue
        folder = os.path.join(staging, "reading", slug(name))
        os.makedirs(folder, exist_ok=True)
        with open(os.path.join(folder, "case.json"), "w") as fh:
            json.dump(case, fh, ensure_ascii=False, indent=1)
        with open(os.path.join(folder, "case.md"), "w") as fh:
            fh.write(render(case))
        written.append(name)
    return written


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--staging", default=os.path.join(HERE, "staging"))
    ap.add_argument("--company", help="only this company (name or folder slug)")
    args = ap.parse_args()
    with open(os.path.join(args.staging, "extracted.json")) as f:
        extracted = json.load(f)
    cases = build_cases(extracted)
    written = write_cases(cases, args.staging, args.company)
    print(f"{'case':42} threads  mails  scans  chars")
    for name in sorted(written, key=lambda n: -len(cases[n]["messages"])):
        c = cases[name]
        chars = sum(len(m["text"]) + len(m["history"]) for m in c["messages"].values())
        scans = sum(1 for d in c["documents"].values() if d["kind"] == "scan")
        print(f"{name[:42]:42} {len(c['threads']):7} {len(c['messages']):6} {scans:6} {chars:6}")
    print(f"\n{len(written)} case files in {os.path.join(args.staging, 'reading')}")


if __name__ == "__main__":
    main()
