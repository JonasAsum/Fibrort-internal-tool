#!/usr/bin/env python3
"""Pick a stratified sample of the Thunderbird export and write it with a ground-truth manifest.

Read-only on the export. Output (default ./sample):
  NNNN.eml       one raw message per file, NNNN is its IMAP UID
  manifest.json  expected sender, recipients, attachments, language and strata per UID
"""
import argparse
import collections
import email
import json
import mailbox
import os
import random
import re
from email import policy
from email.utils import getaddresses, parsedate_to_datetime

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_ROOT = os.path.join(HERE, "..", "Email Data", "6vugs2jt.default-release", "ImapMail", "outlook.office365.com")
OWNER = os.environ.get("MAIL_OWNER_ADDRESS")  # the mailbox owner's primary address (see importer/mail/owner.example.json)
if not OWNER:
    _cfg = os.path.join(HERE, "..", "importer", "mail", "owner.json")
    OWNER = json.load(open(_cfg))["addresses"][0] if os.path.exists(_cfg) else None
if not OWNER:
    raise SystemExit("set MAIL_OWNER_ADDRESS or create importer/mail/owner.json")
SKIP_FOLDERS = ("junk", "deleted", "draft", "template", "outbox", "journal", "notes")
ROLE_LOCALS = {"info", "sales", "service", "support", "office", "contact", "admin", "marketing", "noreply", "no-reply"}
VI_MARKS = re.compile(r"[ăâđêôơưĂÂĐÊÔƠƯạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợụủứừửữự]")
HAN = re.compile(r"[一-鿿]")
DE_WORDS = re.compile(r"\b(sehr geehrte|mit freundlichen|vielen dank|anfrage|angebot|bestellung|guten tag)\b", re.I)


def is_mbox(path):
    try:
        with open(path, "rb") as f:
            return f.read(5) == b"From "
    except OSError:
        return False


def mbox_files(root):
    for dirpath, _, names in os.walk(root):
        for n in sorted(names):
            if n.endswith(".msf") or n.startswith("."):
                continue
            rel = os.path.relpath(os.path.join(dirpath, n), root)
            if any(s in rel.lower() for s in SKIP_FOLDERS):
                continue
            if "sent" in rel.lower():
                continue
            p = os.path.join(dirpath, n)
            if is_mbox(p):
                yield rel, p


def text_of(msg):
    body = msg.get_body(preferencelist=("plain", "html"))
    try:
        return (body.get_content() if body else "")[:4000]
    except (LookupError, UnicodeError):
        return ""


def language(subject, text):
    blob = f"{subject}\n{text}"
    if HAN.search(blob):
        return "zh"
    if VI_MARKS.search(blob):
        return "vi"
    if DE_WORDS.search(blob) or re.search(r"[äöüß]", blob):
        return "de"
    return "en"


def addresses(msg, *headers):
    out = []
    for name, addr in getaddresses([str(h) for hdr in headers for h in msg.get_all(hdr, [])]):
        if addr:
            out.append({"name": name, "email": addr.lower()})
    return out


def describe(folder, key, mb):
    raw = mb.get_bytes(key)
    msg = email.message_from_bytes(raw, policy=policy.default)
    subject = str(msg.get("Subject", ""))
    sender = addresses(msg, "From")
    attachments = [p.get_filename() for p in msg.iter_attachments() if p.get_filename()]
    try:
        when = parsedate_to_datetime(str(msg.get("Date"))).isoformat()
    except (TypeError, ValueError):
        return None
    mid = str(msg.get("Message-ID", "")).strip()
    if not mid or not sender:
        return None
    local = sender[0]["email"].split("@")[0]
    strata = [language(subject, text_of(msg))]
    strata.append("attachment" if attachments else "no-attachment")
    low = subject.lower()
    strata.append("forward" if re.match(r"\s*(fw|fwd|wg)\s*:", low) else "reply" if re.match(r"\s*(re|aw)\s*:", low) else "fresh")
    if local.split("-")[0] in ROLE_LOCALS or local in ROLE_LOCALS:
        strata.append("shared-inbox")
    return {
        "raw": raw,
        "meta": {
            "message_id": mid,
            "folder": folder,
            "date": when,
            "subject": subject,
            "from": sender[0],
            "to": addresses(msg, "To"),
            "cc": addresses(msg, "Cc"),
            "attachments": attachments,
            "from_owner": sender[0]["email"] == OWNER,
            "strata": strata,
        },
    }


def stratified(items, n, rng):
    buckets = collections.defaultdict(list)
    for it in items:
        buckets["|".join(it["meta"]["strata"])].append(it)
    for b in buckets.values():
        rng.shuffle(b)
    order = sorted(buckets, key=lambda k: len(buckets[k]))
    picked = []
    while len(picked) < n and any(buckets.values()):
        for k in order:
            if buckets[k] and len(picked) < n:
                picked.append(buckets[k].pop())
    return picked


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--root", default=DEFAULT_ROOT)
    ap.add_argument("--out", default=os.path.join(HERE, "sample"))
    ap.add_argument("--n", type=int, default=200)
    ap.add_argument("--seed", type=int, default=7)
    args = ap.parse_args()

    seen, items = set(), []
    for folder, path in mbox_files(args.root):
        mb = mailbox.mbox(path)
        for key in mb.keys():
            d = describe(folder, key, mb)
            if d and d["meta"]["message_id"] not in seen:
                seen.add(d["meta"]["message_id"])
                items.append(d)
    print(f"eligible messages: {len(items)}")

    chosen = sorted(stratified(items, args.n, random.Random(args.seed)), key=lambda it: it["meta"]["date"])
    os.makedirs(args.out, exist_ok=True)
    manifest = {}
    for uid, it in enumerate(chosen, start=1):
        with open(os.path.join(args.out, f"{uid:04d}.eml"), "wb") as f:
            f.write(it["raw"])
        manifest[str(uid)] = it["meta"]
    with open(os.path.join(args.out, "manifest.json"), "w", encoding="utf-8") as f:
        json.dump({"owner": OWNER, "messages": manifest}, f, ensure_ascii=False, indent=1)

    tally = collections.Counter(s for it in chosen for s in it["meta"]["strata"])
    print(f"sampled {len(chosen)} -> {args.out}")
    for k, v in sorted(tally.items()):
        print(f"  {k}: {v}")


if __name__ == "__main__":
    main()
