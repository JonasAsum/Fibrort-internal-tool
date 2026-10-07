"""Fact check for the email import. Two checks, both read-only:

  source - every value the import uses (names, emails, phones, titles, legal names, order numbers, amounts, which
           company and deal a mail belongs to, attachments) must be found again in the ORIGINAL archive.
  crm    - the CRM must hold exactly what the import plan says: nothing missing, changed, extra or doubled.

    .venv/bin/python -m mail.verify            # both
    .venv/bin/python -m mail.verify --source   # originals -> import plan, no network
    .venv/bin/python -m mail.verify --crm      # import plan -> CRM, reads only

Writes staging/verify_report.json (every finding) and staging/verify_report.csv (everything not VERIFIED, opens
in Excel). Exits 1 when anything is NOT FOUND, MISSING, CHANGED, EXTRA or DUPLICATE.

The source check decodes the mbox files with its own few lines of stdlib code and shares no parsing code with
reader.py or people.py (only the folder walk), so a bug there cannot confirm itself.
"""
import argparse
import csv
import email
import email.policy
import email.utils
import hashlib
import html
import io
import json
import logging
import mailbox
import os
import pickle
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime

from . import pipeline
from .extract import DEFAULT_ROOT
from .plan import EMAIL_RE, PHONE_RE, REF_RE, build_plan, fold, load_enrichment
from .reader import SKIP_FOLDERS, _walk_mboxes

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE_VERSION = 3
logging.getLogger("pypdf").setLevel(logging.ERROR)  # broken-xref chatter on scanned PDFs

VERIFIED, WEAK, MANUAL, NOT_FOUND = "VERIFIED", "WEAK", "MANUAL", "NOT FOUND"
MISSING, CHANGED, EXTRA, DUPLICATE = "MISSING", "CHANGED", "EXTRA", "DUPLICATE"
WAITING, DECLINED = "WAITING", "DECLINED"  # a close the AI agent asked for: in your CRM inbox / you said no
FAILING = {NOT_FOUND, MISSING, CHANGED, EXTRA, DUPLICATE}
INTERNAL = {"fibrort.com", "fibrort.org"}


# ---------------------------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------------------------

def norm(s: str | None) -> str:
    """Compare text as printed: full-width forms folded (，→ ,), case and all whitespace ignored."""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", s or "")).casefold()


PHONE_CHUNK = re.compile(r"\+?\d[\d\s\-–().+/]{5,}\d")


def phone_chunks(text: str) -> set[str]:
    return {re.sub(r"\D", "", c) for c in PHONE_CHUNK.findall(unicodedata.normalize("NFKC", text or ""))}


def phone_tail(phone: str) -> str:
    """The last 9 digits: the same number written +86 138…, 138-…, or 0512-… (country code and trunk 0 differ)."""
    d = re.sub(r"\D", "", phone)
    return d[-9:] if len(d) >= 9 else d


CJK = re.compile("[\u4e00-\u9fff]")


def words(text: str) -> set[str]:
    """Latin words, so 'Baka' is not found inside 'Bakar' (Chinese has no spaces and is matched as a substring)."""
    return {w for w in re.findall(r"[^\W\d_]+", unicodedata.normalize("NFKC", text or "").casefold()) if not CJK.search(w)}


def name_key(name: str) -> str:
    """Word order does not matter: 'Last, First' and 'First Last' are one display name."""
    return "".join(sorted(re.findall(r"[^\W\d_]+", unicodedata.normalize("NFKC", name or "").casefold())))


def amount_forms(minor: int) -> list[str]:
    v = minor / 100
    forms = [f"{v:,.2f}", f"{v:.2f}"]
    if minor % 100 == 0:
        forms += [f"{v:,.0f}", f"{int(v)}"]
    return [norm(f) for f in forms]


def domain_of(addr: str) -> str:
    return addr.rsplit("@", 1)[-1].lower() if "@" in addr else ""


def on_domain(addr: str, domains) -> str | None:
    d = domain_of(addr)
    return next((x for x in domains if d == x or d.endswith("." + x)), None)


def ref_forms(key: str) -> list[str]:
    forms = [norm(key)]
    if key.upper().startswith("SO"):  # Chinese order confirmations print SO as S0
        forms.append(norm("S0" + key[2:]))
    return forms


class Findings:
    def __init__(self):
        self.rows: list[dict] = []

    def add(self, check: str, record: str, field: str, value, status: str, detail: str = "") -> None:
        self.rows.append({"check": check, "record": record, "field": field,
                          "value": value if isinstance(value, str) else json.dumps(value, ensure_ascii=False),
                          "status": status, "detail": detail})


# ---------------------------------------------------------------------------------------------
# the archive, read independently
# ---------------------------------------------------------------------------------------------

def _decode_part(part) -> str:
    payload = part.get_payload(decode=True) or b""
    cs = (part.get_content_charset() or "utf-8").lower()
    if cs in {"gb2312", "gbk", "x-gbk", "gb_2312-80", "cp936"}:
        cs = "gb18030"  # mail clients label GBK text as gb2312
    for enc in (cs, "utf-8", "gb18030"):
        try:
            text = payload.decode(enc)
            break
        except (LookupError, UnicodeDecodeError):
            continue
    else:
        text = payload.decode("utf-8", "replace")
    if part.get_content_subtype() == "html":
        text = re.sub(r"(?is)<(script|style)\b.*?</\1>", " ", text)
        text = html.unescape(re.sub(r"(?s)<[^>]+>", " ", re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>", "\n", text)))
    return text


def _header(msg, name: str) -> str:
    try:
        return " ".join(str(v) for v in (msg.get_all(name) or []))
    except Exception:  # a malformed header must not hide the rest of the mail
        return ""


def _pdf_text(data: bytes) -> str:
    try:
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(data))
        return "\n".join((p.extract_text() or "") for p in reader.pages[:40])
    except Exception:
        return ""


def _read_message(msg, folder: str, src: str, pdfs: dict) -> dict:
    parties = {}
    for h in ("From", "To", "Cc"):
        parties[h] = [(n.strip(), a.strip().lower()) for n, a in email.utils.getaddresses([_header(msg, h)]) if a]
    mid = _header(msg, "Message-ID")
    found = re.search(r"<([^<>]+)>", mid)
    mid = ((found.group(1) if found else mid).strip().lower()) or None
    try:
        date = email.utils.parsedate_to_datetime(_header(msg, "Date")).isoformat()
    except Exception:
        date = None
    texts, atts = [], {}
    for part in msg.walk():
        name = part.get_filename()
        if name:
            data = part.get_payload(decode=True) or b""
            if not data:
                continue
            sha = hashlib.sha256(data).hexdigest()
            atts[sha] = name
            if sha not in pdfs and (name.lower().endswith(".pdf") or part.get_content_type() == "application/pdf"):
                pdfs[sha] = norm(_pdf_text(data))
        elif part.get_content_maintype() == "text":
            texts.append(_decode_part(part))
    subject = re.sub(r"\s+", " ", _header(msg, "Subject")).strip()
    body = "\n".join(texts)
    names = " ".join(n for ps in parties.values() for n, _ in ps)
    addrs = " ".join(a for ps in parties.values() for _, a in ps)
    # prose: the same text with every email address taken out, so a name is never "found" inside
    # wenhao.zhang@… - that would only prove someone spelled a name from the address
    prose = EMAIL_RE.sub(" ", " ".join([subject, names, body, " ".join(atts.values())]))
    return {"id": mid, "src": src, "folders": [folder], "subject": subject, "date": date,
            "from": (parties["From"] or [("", "")])[0], "to": parties["To"], "cc": parties["Cc"],
            "text": norm(" ".join([prose, addrs, body])), "prose": norm(prose), "words": words(prose),
            "phones": phone_chunks(body), "atts": atts}


class Archive:
    """Every message in the account, indexed by Message-ID (and by mailbox#index for the few without one)."""

    def __init__(self, root: str, cache_path: str | None = None):
        files = [(p, segs) for p, segs in _walk_mboxes(root) if segs[0] not in SKIP_FOLDERS
                 and segs[-1] not in SKIP_FOLDERS]
        stamp = [CACHE_VERSION] + [(p, os.path.getsize(p), os.path.getmtime(p)) for p, _ in files]
        if cache_path and os.path.exists(cache_path):
            with open(cache_path, "rb") as fh:
                cached = pickle.load(fh)
            if cached.get("stamp") == stamp:
                self.mails, self.pdfs = cached["mails"], cached["pdfs"]
                self._index()
                return
        self.mails, self.pdfs = [], {}
        by_id: dict[str, dict] = {}
        for path, segs in files:
            folder = "/".join(segs)
            box = mailbox.mbox(path, factory=lambda f: email.message_from_binary_file(f, policy=email.policy.default))
            for idx, msg in enumerate(box):
                try:
                    rec = _read_message(msg, folder, f"{os.path.basename(path)}#{idx}", self.pdfs)
                except Exception as e:
                    print(f"  could not read {folder}#{idx}: {e!r}", file=sys.stderr)
                    continue
                prev = by_id.get(rec["id"]) if rec["id"] else None
                if prev:
                    prev["folders"].append(folder)
                    prev.setdefault("also", []).append(rec["src"])
                    continue
                if rec["id"]:
                    by_id[rec["id"]] = rec
                self.mails.append(rec)
        if cache_path:
            with open(cache_path, "wb") as fh:
                pickle.dump({"stamp": stamp, "mails": self.mails, "pdfs": self.pdfs}, fh)
        self._index()

    def _index(self) -> None:
        self.by_id = {m["id"]: m for m in self.mails if m["id"]}
        self.by_src = defaultdict(list)
        for m in self.mails:
            for s in [m["src"]] + m.get("also", []):
                self.by_src[s].append(m)
        # Exchange writes some senders as IMCEAEX-… legacy addresses; a display name used with exactly one
        # real address elsewhere in the archive says who that is
        names = defaultdict(set)
        for m in self.mails:
            for n, a in [m["from"]] + m["to"] + m["cc"]:
                if n and "@" in a and not a.startswith("imceaex"):
                    names[name_key(n)].add(a)
        self.names = names
        self.exchange = {k: next(iter(v)) for k, v in names.items() if len(v) == 1}
        self.from_addr = {}
        self.sent_by = defaultdict(list)
        self.with_addr = defaultdict(list)
        for m in self.mails:
            n, a = m["from"]
            if a.startswith("imceaex") or not a:
                a = self.exchange.get(name_key(n), a)
            self.from_addr[id(m)] = a
            self.sent_by[a].append(m)
            for _, x in [m["from"]] + m["to"] + m["cc"]:
                self.with_addr[x].append(m)
            self.with_addr[a].append(m)
        self.header_addrs = set(self.with_addr)
        self.all_text = None

    def find(self, message_id: str, source: str) -> dict | None:
        if message_id in self.by_id:
            return self.by_id[message_id]
        cands = self.by_src.get(source, [])
        return cands[0] if len(cands) == 1 else None

    def anywhere(self, form: str) -> dict | None:
        """The first mail whose text, or one of whose PDFs, contains form."""
        for m in self.mails:
            if form in m["text"] or any(form in self.pdfs.get(s, "") for s in m["atts"]):
                return m
        return None

    def name_in(self, token: str, mails) -> dict | None:
        """A mail whose prose has this name part: a whole word for Latin, a substring for Chinese."""
        t = norm(token)
        for m in mails:
            if (t in m["prose"]) if CJK.search(t) else (t in m["words"]):
                return m
        return None

    def in_mails(self, form: str, mails, field: str = "text") -> dict | None:
        for m in mails:
            if form in m[field] or any(form in self.pdfs.get(s, "") for s in m["atts"]):
                return m
        return None


def where(m: dict | None) -> str:
    if not m:
        return ""
    return f"{(m['date'] or '')[:10]} '{m['subject'][:70]}' ({m['folders'][0]})"


# ---------------------------------------------------------------------------------------------
# check 1: archive -> import plan
# ---------------------------------------------------------------------------------------------

def check_source(plan, extracted: dict, enrichment: dict, arc: Archive, files_dir: str, out: Findings) -> None:
    F = lambda *a: out.add("source", *a)  # noqa: E731
    companies = {c["name"]: c for c in plan.companies}
    routing = (enrichment or {}).get("routing", {})
    hand_deal_msgs = {mid: e["key"] for e in (enrichment or {}).get("deals", []) for mid in e.get("message_ids", [])}
    contact_company = {e: c["company"] for c in plan.contacts for e in c["emails"]}

    def company_mails(co: dict) -> list[dict]:
        out_ = [m for m in arc.mails if set(m["folders"]) & set(co.get("folders") or [])]
        for d in co.get("domains") or []:
            out_ += [m for a, ms in arc.with_addr.items() if on_domain(a, [d]) for m in ms]
        if co.get("origin") == "internal":  # a FIBRO entity is named in its own staff's signatures
            out_ += [m for a, ms in arc.sent_by.items() if domain_of(a) in INTERNAL for m in ms]
        for d in plan.deals.values():  # a contract party's name is in the contract it signed
            if d.get("company") == co["name"]:
                out_ += [arc.by_id[mid] for mid in d["messages"] if mid in arc.by_id]
        return list({id(m): m for m in out_}.values())

    co_mails = {name: company_mails(co) for name, co in companies.items()}

    # --- contacts ---
    for ct in plan.contacts:
        rec = f"contact {ct['emails'][0] if ct['emails'] else ct['key']}"
        own = [m for e in ct["emails"] for m in arc.sent_by.get(e, [])]
        part = list({id(m): m for e in ct["emails"] for m in arc.with_addr.get(e, [])}.values())
        if not part:  # only on a forwarded thread's quoted Cc/To line: that line is where the name is written
            part = [m for m in arc.mails if any(norm(e) in m["text"] for e in ct["emails"])]
        for e in ct["emails"]:
            if e in arc.header_addrs:
                F(rec, "email", e, VERIFIED, f"in the headers of {len(arc.with_addr[e])} mails")
            elif arc.anywhere(norm(e)):
                F(rec, "email", e, WEAK, "only in a mail text, never in a From/To/Cc header: " + where(arc.anywhere(norm(e))))
            else:
                F(rec, "email", e, NOT_FOUND, "appears nowhere in the archive")
        if not ct["mailbox"] and "@" not in ct["full_name"]:
            shown = EMAIL_RE.sub(" ", " ".join(n for m in part for n, a in [m["from"]] + m["to"] + m["cc"]
                                               if a in ct["emails"]))
            for tok in re.findall(r"[一-鿿]+|[^\W\d_]+", ct["full_name"]):
                t = norm(tok)
                if len(t) < 2:
                    continue  # an initial proves nothing either way
                if (t in norm(shown)) if CJK.search(t) else (t in words(shown)):
                    F(rec, "name", tok, VERIFIED, "the display name used with this address")
                elif arc.name_in(tok, own):
                    F(rec, "name", tok, VERIFIED, "in their own mail: " + where(arc.name_in(tok, own)))
                elif arc.name_in(tok, part):
                    F(rec, "name", tok, WEAK, "only in a mail they took part in, not their own: "
                      + where(arc.name_in(tok, part)))
                elif any(t in norm(e.split("@")[0]) for e in ct["emails"]):
                    # first.last@ with no display name anywhere: a reading of the address, accepted but marked
                    F(rec, "name", tok, WEAK, "spelled from the address; no mail shows this name")
                else:
                    F(rec, "name", tok, NOT_FOUND, f"not in any of their {len(part)} mails")
        for p in ct["phones"]:
            tail = phone_tail(p["phone"])
            hit = next((m for m in own if any(tail in c for c in m["phones"])), None)
            if hit:
                F(rec, "phone", p["phone"], VERIFIED, "in their own mail: " + where(hit))
                continue
            hit = next((m for m in part if any(tail in c for c in m["phones"])), None) or next(
                (m for m in arc.mails if any(tail in c for c in m["phones"])), None)
            pdf = next((s for s, t in arc.pdfs.items() if tail in t), None)
            if hit:
                F(rec, "phone", p["phone"], WEAK, "only in someone else's mail: " + where(hit))
            elif pdf:
                F(rec, "phone", p["phone"], WEAK, "only in a PDF")
            else:
                F(rec, "phone", p["phone"], NOT_FOUND, "this number appears nowhere in the archive")
        if ct.get("title"):
            t = norm(ct["title"])
            if arc.in_mails(t, own):
                F(rec, "title", ct["title"], VERIFIED, "in their own mail: " + where(arc.in_mails(t, own)))
            elif arc.anywhere(t):
                F(rec, "title", ct["title"], WEAK, "not in their own mails, only in: " + where(arc.anywhere(t)))
            else:
                F(rec, "title", ct["title"], NOT_FOUND, "this title appears nowhere in the archive")
        if ct.get("company"):
            co = companies.get(ct["company"])
            doms = (co or {}).get("domains") or []
            if co is None:
                F(rec, "company", ct["company"], NOT_FOUND, "the contact's company is not in the import")
            elif any(on_domain(e, doms) for e in ct["emails"]):
                F(rec, "company", ct["company"], VERIFIED, "their address is on the company's domain")
            elif ct["internal"] and all(domain_of(e) in INTERNAL for e in ct["emails"]):
                F(rec, "company", ct["company"], VERIFIED, "FIBRO address; the entity is named in their signature"
                  if arc.in_mails(norm(ct["company"]), own) else "FIBRO address")
            elif any(set(m["folders"]) & set(co.get("folders") or []) for m in part):
                F(rec, "company", ct["company"], VERIFIED, "their mails are filed in the company's folder")
            elif arc.in_mails(norm(ct["company"]), own):
                F(rec, "company", ct["company"], VERIFIED, "the company is named in their own mail")
            else:
                F(rec, "company", ct["company"], NOT_FOUND, "no domain, folder or signature ties them to it")

    # --- the same mistakes anywhere: a colleague's name on an outside address, one name twice ---
    staff = {c["full_name"] for c in plan.contacts if c["internal"] and not c["mailbox"]}
    seen = defaultdict(list)
    for ct in plan.contacts:
        rec = f"contact {ct['emails'][0] if ct['emails'] else ct['key']}"
        if not ct["internal"] and ct["full_name"] in staff:
            F(rec, "not a colleague", ct["full_name"], NOT_FOUND, "an outside address carries a FIBRO colleague's name")
        if not ct["mailbox"]:
            seen[ct["full_name"]].append(rec)
    for name, recs in seen.items():
        if len(recs) > 1:
            F(recs[0], "same name", name, WEAK, f"{len(recs)} contacts have this name: "
              + ", ".join(r.split(" ", 1)[1] for r in recs) + " (one person, or a shared inbox?)")

    # --- companies ---
    for name, co in companies.items():
        rec = f"company {name}"
        mails = co_mails[name]
        for d in co.get("domains") or []:
            users = [a for a in arc.header_addrs if on_domain(a, [d])]
            if users:
                F(rec, "domain", d, VERIFIED, f"used by {', '.join(sorted(users)[:3])}")
            elif arc.anywhere("@" + norm(d)):
                F(rec, "domain", d, WEAK, "only on a forwarded thread's quoted header lines: " + where(arc.anywhere("@" + norm(d))))
            else:
                F(rec, "domain", d, NOT_FOUND, "no sender or recipient uses it")
        for field, value in (("legal name", co.get("legal_name")), ("legal name (EN)", co.get("legal_name_en"))):
            if not value:
                continue
            t = norm(value)
            hit = arc.in_mails(t, mails)
            if hit:
                F(rec, field, value, VERIFIED, "in the company's own mail or PDF: " + where(hit))
            elif arc.anywhere(t):
                F(rec, field, value, WEAK, "only in a mail not tied to the company: " + where(arc.anywhere(t)))
            else:
                F(rec, field, value, NOT_FOUND, "this name appears nowhere in the archive")
        letters = re.sub(r"[^a-z0-9]", "", name.lower())
        leafs = {f.rstrip("/").split("/")[-1] for f in co.get("folders") or []}
        if name in leafs:
            F(rec, "name", name, VERIFIED, "the name of its mail folder")
        elif letters and any(letters in re.sub(r"[^a-z0-9]", "", d) for d in co.get("domains") or []):
            F(rec, "name", name, VERIFIED, "spelled from its domain")
        elif arc.in_mails(norm(name), mails):
            F(rec, "name", name, VERIFIED, "in the company's own mail: " + where(arc.in_mails(norm(name), mails)))
        elif any(norm(name) in norm(v) for v in (co.get("legal_name"), co.get("legal_name_en")) if v):
            F(rec, "name", name, VERIFIED, "a part of its legal name")
        else:
            F(rec, "name", name, NOT_FOUND, "not its folder, not its domain, not in its mails")
        for field, label in (("customer type", co.get("customer_type_label")), ("business line", co.get("business_line"))):
            if not label:
                continue
            ok = any(label.lower() in f.lower().split("/") for f in co.get("folders") or [])
            F(rec, field, label, VERIFIED if ok else NOT_FOUND,
              "from its folder path" if ok else "no folder of this company says so")

    # --- deals ---
    for key, d in plan.deals.items():
        rec = f"deal {key}"
        manual_note = f"hand-read; evidence: {d.get('evidence') or 'none given'}"
        if key.startswith("RFQ-"):
            F(rec, "number", key, MANUAL, "a name given by the reader, not a number from the mails")
        else:
            hit = next((arc.anywhere(f) for f in ref_forms(key) if arc.anywhere(f)), None)
            F(rec, "number", key, VERIFIED if hit else NOT_FOUND, where(hit) if hit else "in no mail or PDF")
        if d.get("origin") == "reading" and isinstance(d.get("evidence"), list):
            continue  # a deal found by reading: its quotes are checked below with every other read fact
        doc_texts = [arc.pdfs.get(x["sha256"], "") for x in d["docs"]]
        readable = [t for t in doc_texts if t]
        if d.get("amount_minor"):
            forms = amount_forms(d["amount_minor"])
            amount = f"{d['amount_minor'] / 100:,.2f} {d['currency']}"
            if any(f in t for t in readable for f in forms):
                F(rec, "amount", amount, VERIFIED, "in its document: " + ", ".join(x["filename"] for x in d["docs"]))
            elif readable:
                F(rec, "amount", amount, NOT_FOUND, "its PDF has text, but not this amount")
            elif any(arc.anywhere(f) for f in forms):
                F(rec, "amount", amount, WEAK, "the document is a scan; the amount is written in a mail: "
                  + where(next(arc.anywhere(f) for f in forms if arc.anywhere(f))))
            else:
                F(rec, "amount", amount, MANUAL, manual_note)
        for i, ln in enumerate(d.get("lines") or [], 1):
            label = f"line {i} {ln.get('model') or ln.get('text', '')[:30]}"
            if not readable:
                F(rec, label, f"{ln.get('qty')} x {ln.get('unit_price_minor', 0) / 100:,.2f}", MANUAL, manual_note)
                continue
            price_ok = any(f in t for t in readable for f in amount_forms(ln.get("unit_price_minor") or 0))
            model_ok = not ln.get("model") or any(norm(ln["model"]) in t for t in readable)
            F(rec, label, f"{ln.get('qty')} x {ln.get('unit_price_minor', 0) / 100:,.2f}",
              VERIFIED if price_ok and model_ok else NOT_FOUND,
              "price and model in its PDF" if price_ok and model_ok else
              f"not in its PDF: {'price' if not price_ok else ''} {'model' if not model_ok else ''}".strip())
        if d.get("company"):
            co = companies.get(d["company"], {})
            names = [norm(v) for v in (co.get("legal_name"), co.get("legal_name_en"), d["company"], *(co.get("parties") or []))
                     if v and len(norm(v)) >= 4]
            msgs = [arc.by_id[m] for m in d["messages"] if m in arc.by_id]
            if any(n in t for t in readable for n in names):
                F(rec, "company", d["company"], VERIFIED, "named in its document")
            elif any(set(m["folders"]) & set(co.get("folders") or []) or any(
                    on_domain(a, co.get("domains") or []) for _, a in [m["from"]] + m["to"] + m["cc"]) for m in msgs):
                F(rec, "company", d["company"], VERIFIED, "its mails are filed with / exchanged with the company")
            elif not d["docs"]:
                F(rec, "company", d["company"], MANUAL, manual_note)
            else:
                F(rec, "company", d["company"], NOT_FOUND, "neither its document nor its mails name the company")

    # --- emails ---
    for m in plan.emails:
        rec = f"email {m['date_iso'][:10] if m['date_iso'] else '?'} {m['subject'][:50]!r}"
        raw = arc.find(m["message_id"], m["source"])
        if not raw:
            F(rec, "exists", m["message_id"], NOT_FOUND, "no mail with this Message-ID in the archive")
            continue
        problems = []
        if norm(raw["subject"]) != norm(m["subject"]):
            problems.append(f"subject in the archive: {raw['subject'][:80]!r}")
        if raw["date"] and m["date_iso"] and datetime.fromisoformat(raw["date"]) != datetime.fromisoformat(m["date_iso"]):
            problems.append(f"date in the archive: {raw['date']}")
        sender = arc.from_addr[id(raw)]
        mine = (m["from"]["email"] or "").lower()
        if mine != sender:
            if raw["from"][1].startswith("imceaex") and mine in arc.names.get(name_key(raw["from"][0]), ()):
                pass  # an Exchange path; this display name is used with this address elsewhere
            elif not raw["from"][1] or raw["from"][1].startswith("imceaex"):
                F(rec, "sender", m["from"]["email"], WEAK, f"the archive has an Exchange/empty sender ({raw['from'][0]!r})")
            else:
                problems.append(f"sender in the archive: {sender}, the import has {mine or 'none'}")
        F(rec, "headers", m["message_id"], NOT_FOUND if problems else VERIFIED, "; ".join(problems) or where(raw))
        if re.search(r"LearnAboutSenderIdentification|erhalten nicht häufig E-Mails|don't often get email|"
                     r"thường không nhận được email|通常不会收到", m["top"] or ""):
            F(rec, "text", m["subject"], NOT_FOUND, "starts with Outlook's sender warning instead of the mail's text")
        if m.get("company"):
            co = companies.get(m["company"], {})
            doms = co.get("domains") or []
            ev, status = None, VERIFIED
            dom = next((d for _, a in [raw["from"]] + raw["to"] + raw["cc"] for d in [on_domain(a, doms)] if d), None)
            if dom:
                ev = f"a participant writes from {dom}"
            elif set(raw["folders"]) & set(co.get("folders") or []):
                ev = "filed in the company's folder"
            elif any(norm(d) in raw["text"] for d in doms):
                ev = "its domain appears in the mail (forwarded / quoted)"
            elif any(v and len(norm(v)) >= 4 and (norm(v) in raw["text"] or any(norm(v) in arc.pdfs.get(s, "") for s in raw["atts"]))
                     for v in (co.get("legal_name"), co.get("legal_name_en"), m["company"], *(co.get("parties") or []))):
                ev = "the company is named in the mail or its PDF"
            elif m.get("deal_key") and plan.deals.get(m["deal_key"], {}).get("company") == m["company"]:
                ev, status = f"belongs to deal {m['deal_key']}", VERIFIED if any(
                    f in raw["text"] for f in ref_forms(m["deal_key"])) else MANUAL
            elif m["message_id"] in routing:
                ev, status = "moved here by a reading (its quote is checked below)", MANUAL
            else:
                via = [a for _, a in [raw["from"]] + raw["to"] + raw["cc"] if contact_company.get(a) == m["company"]]
                if via:
                    ev, status = f"only because {via[0]} is filed under this company", WEAK
            F(rec, "company link", m["company"], status if ev else NOT_FOUND, ev or "nothing in the mail ties it to "
              "this company")
        if m.get("deal_key"):
            pdf_hit = any(f in arc.pdfs.get(s, "") for s in raw["atts"] for f in ref_forms(m["deal_key"]))
            if any(f in raw["text"] for f in ref_forms(m["deal_key"])) or pdf_hit:
                F(rec, "deal link", m["deal_key"], VERIFIED, "the number is in the mail or its PDF")
            elif m["message_id"] in hand_deal_msgs or m["message_id"] in plan.deals[m["deal_key"]]["messages"]:
                F(rec, "deal link", m["deal_key"], MANUAL, "listed for this deal by the reader / its scanned document")
            else:
                F(rec, "deal link", m["deal_key"], NOT_FOUND, "the deal number is not in this mail")

    # --- attachments ---
    for a in plan.attachments:
        rec = f"attachment {a['filename']}"
        path = os.path.join(files_dir, a["stored_as"])
        raw = arc.by_id.get(a["message_id"])
        try:
            with open(path, "rb") as fh:
                sha = hashlib.sha256(fh.read()).hexdigest()
        except OSError:
            F(rec, "file", a["stored_as"], NOT_FOUND, "the staged file is missing")
            continue
        if sha != a["sha256"]:
            F(rec, "file", a["stored_as"], NOT_FOUND, "the staged file's content differs from what was recorded")
        elif not raw or a["sha256"] not in raw["atts"]:
            F(rec, "file", a["stored_as"], NOT_FOUND, "the mail it is filed on has no attachment with this content")
        else:
            F(rec, "file", a["filename"], VERIFIED, "byte-identical to the attachment in the original mail")

    # --- facts read from the mails (READING.md): every quote must be in the mail or PDF it cites ---
    imported = {m["message_id"] for m in extracted["messages"] if m["import"]}
    source = {m["message_id"]: m["source"] for m in extracted["messages"]}

    def quotes(rec: str, evidence) -> None:
        if not isinstance(evidence, list) or not evidence:
            F(rec, "quote", str(evidence or "-")[:80], MANUAL,
              "hand-written evidence, no quote to check" if evidence else "no evidence given")
            return
        for e in evidence:
            q = norm(e.get("quote"))
            if e.get("scan"):
                F(rec, "quote", e["quote"][:70], MANUAL, "read from a scan (no text layer): both readings agreed or Jonas checked")
            elif e.get("message_id"):
                raw = arc.find(e["message_id"], source.get(e["message_id"], ""))
                if not raw or e["message_id"] not in source:
                    F(rec, "quote", e["quote"][:70], NOT_FOUND, "the cited mail is not in the archive any more")
                elif q in raw["text"] or any(q in arc.pdfs.get(s, "") for s in raw["atts"]):
                    F(rec, "quote", e["quote"][:70], VERIFIED if e["message_id"] in imported else WEAK,
                      where(raw) + ("" if e["message_id"] in imported else " (a mail that is not imported)"))
                else:
                    F(rec, "quote", e["quote"][:70], NOT_FOUND, "not in the cited mail: " + where(raw))
            elif e.get("sha256"):
                ok = q in arc.pdfs.get(e["sha256"], "")
                F(rec, "quote", e["quote"][:70], VERIFIED if ok else NOT_FOUND, "in the cited PDF" if ok else "not in the cited PDF")

    for key, d in plan.deals.items():
        if d.get("origin") != "reading":
            continue
        quotes(f"deal {key}", d.get("evidence"))
        ev = d.get("evidence") if isinstance(d.get("evidence"), list) else []
        if d.get("amount_minor") and ev:
            ok = any(f in norm(e["quote"]) for e in ev for f in amount_forms(d["amount_minor"]))
            F(f"deal {key}", "amount", f"{d['amount_minor'] / 100:,.2f} {d['currency']}", VERIFIED if ok else NOT_FOUND,
              "inside its quotes" if ok else "not inside any of its quotes")
    for r in plan.routing:
        quotes(f"routing {r['message_id'][:40]}", r.get("evidence"))
    for u in plan.contact_updates:
        quotes(f"title {u['email']}", u.get("evidence"))
    for kind, rows in (("task", plan.tasks), ("note", plan.notes)):
        for r in rows:
            rec = f"{kind} {r['id']}"
            quotes(rec, r.get("evidence"))
            text = " ".join(str(r.get(k) or "") for k in ("subject", "body"))
            for tok in set(REF_RE.findall(text)) | set(EMAIL_RE.findall(text)):
                hit = next((arc.anywhere(f) for f in ref_forms(tok) if arc.anywhere(f)), None)
                F(rec, "reference", tok, VERIFIED if hit else NOT_FOUND, where(hit) if hit else "in no mail or PDF")
            for tok in set(PHONE_RE.findall(text)):
                hit = any(phone_tail(tok) in c for m in arc.mails for c in m["phones"])
                F(rec, "phone", tok, VERIFIED if hit else NOT_FOUND, "" if hit else "in no mail")


# ---------------------------------------------------------------------------------------------
# check 2: import plan -> CRM
# ---------------------------------------------------------------------------------------------

def check_crm(plan, client: pipeline.MailClient, out: Findings, emit=print) -> None:
    F = lambda *a: out.add("crm", *a)  # noqa: E731
    ws = lambda s: re.sub(r"\s+", " ", (s or "").strip())  # noqa: E731

    emit("  reading companies")
    companies = list(client.paginate("/v1/companies", {"limit": 200}))
    by_fold = defaultdict(list)
    for c in companies:
        by_fold[fold(c["display_name"])].append(c)
    for k, cs in by_fold.items():
        if len(cs) > 1:
            F(f"company {cs[0]['display_name']}", "record", len(cs), DUPLICATE, f"{len(cs)} companies with this name")
    columns = {f["label"].lower(): f["column_name"] for f in client.custom_fields("company")}
    company_id, company_name = {}, {c["id"]: c["display_name"] for c in companies}
    for co in plan.companies:
        rec = f"company {co['name']}"
        found = by_fold.get(fold(co["name"]))
        if not found:
            F(rec, "record", co["name"], MISSING, "not in the CRM")
            continue
        cur = client.get_company(found[0]["id"])
        company_id[co["name"]] = cur["id"]
        want = {"name": co["name"], "legal name": co.get("legal_name") or co.get("legal_name_en")}
        have = {"name": cur.get("display_name"), "legal name": cur.get("legal_name")}
        for attr, label, *_ in pipeline.COMPANY_FIELDS:
            value = co.get(f"{attr}_label") if attr in ("relationship", "customer_type") else co.get(attr)
            want[label] = value
            have[label] = cur.get(columns.get(label.lower(), "?"))
        for f, v in want.items():
            if v and have[f] != v:
                F(rec, f, v, CHANGED, f"the CRM has {have[f]!r}")
            elif v:
                F(rec, f, v, VERIFIED)
        doms = {d["domain"] for d in cur.get("domains") or []}
        for d in co.get("domains") or []:
            F(rec, "domain", d, VERIFIED if d in doms else MISSING, "" if d in doms else "not on the company")

    emit("  reading contacts")
    contacts = client.all_contacts()
    by_email = defaultdict(list)
    for ct in contacts:
        for e in ct.get("emails") or []:
            by_email[e["email"].lower()].append(ct)
    for e, cts in by_email.items():
        if len({c["id"] for c in cts}) > 1:
            F(f"contact {e}", "record", e, DUPLICATE, f"this address is on {len(cts)} contacts: "
              + ", ".join(c["full_name"] for c in cts))
    contact_id, contact_name = {}, {c["id"]: c.get("primary_email") or c["full_name"] for c in contacts}
    for ct in plan.contacts:
        rec = f"contact {ct['emails'][0] if ct['emails'] else ct['key']}"
        found = next((by_email[e][0] for e in ct["emails"] if e in by_email), None)
        if not found:
            F(rec, "record", ct["full_name"], MISSING, "no contact with any of its addresses")
            continue
        for e in ct["emails"]:
            contact_id[e] = found["id"]
        F(rec, "name", ct["full_name"], VERIFIED if found["full_name"] == ct["full_name"] else CHANGED,
          "" if found["full_name"] == ct["full_name"] else f"the CRM has {found['full_name']!r}")
        have_e = {e["email"].lower() for e in found.get("emails") or []}
        for e in ct["emails"]:
            F(rec, "email", e, VERIFIED if e in have_e else MISSING, "" if e in have_e else "not on the contact")
        have_p = {p["phone"] for p in found.get("phones") or []}
        want_p = {p["phone"] for p in ct["phones"]}
        for p in sorted(want_p):
            F(rec, "phone", p, VERIFIED if p in have_p else MISSING, "" if p in have_p else "not on the contact")
        for p in sorted(have_p - want_p):
            F(rec, "phone", p, EXTRA, "in the CRM but not from the archive (added by hand or by another import?)")
        if ct.get("title"):
            want_t, have_t = ws(ct["title"][:200]), ws(found.get("title"))
            F(rec, "title", want_t, VERIFIED if want_t == have_t else CHANGED, "" if want_t == have_t else
              f"the CRM has {have_t!r}")
        if ct.get("company") and ct["company"] in company_id:
            emp = (found.get("employer") or {}).get("company_name")
            ok = emp and fold(emp) == fold(ct["company"])
            F(rec, "employer", ct["company"], VERIFIED if ok else CHANGED, "" if ok else f"the CRM has {emp!r}")

    def agent_check(rec: str, record: dict, text: str | None, item: dict) -> None:
        """A read record must say it was read by AI (agent passport, or the "[AI-read]" marker) and carry its quotes."""
        by = record.get("captured_by") or ""
        marked = "[AI-read]" in ((record.get("subject") or "") + (record.get("title") or "") + (text or ""))
        ok_by = by.startswith("agent:") or marked
        F(rec, "marked as AI-read", "AI agent or [AI-read]", VERIFIED if ok_by else CHANGED,
          "" if ok_by else f"written by {by.split(':')[0] or 'unknown'} with no [AI-read] marker")
        if isinstance(item.get("evidence"), list) and item["evidence"]:
            ok = "Read from the mails by AI" in (text or "")
            F(rec, "quotes shown", len(item["evidence"]), VERIFIED if ok else CHANGED, "" if ok else "the record does not show its quotes")

    closes = {}
    for status in ("pending", "approved", "rejected"):
        for ap in client.paginate("/v1/approvals", {"kind": "advance_deal", "status": status, "limit": 100}):
            if ap.get("proposed_by", "").startswith("agent:") and ap["created_at"] > closes.get(ap["target_entity_id"], {}).get("created_at", ""):
                closes[ap["target_entity_id"]] = ap

    emit("  reading deals")
    deals = list(client.paginate("/v1/deals", {"limit": 200}))
    deal_by_name = defaultdict(list)
    for d in deals:
        deal_by_name[d["name"]].append(d)
    deal_id = {}
    for key, d in plan.deals.items():
        rec = f"deal {key}"
        found = deal_by_name.get(d["name"])
        if not found:
            F(rec, "record", d["name"], MISSING, "not in the CRM (imported before it was read or reviewed?)")
            continue
        if len(found) > 1:
            F(rec, "record", d["name"], DUPLICATE, f"{len(found)} deals with this name")
        cur = found[0]
        deal_id[key] = cur["id"]
        checks = [("company", d.get("company"), company_name.get(cur.get("company_id")))]
        if d.get("origin") == "reading":
            agent_check(rec, cur, client.get(f"/v1/deals/{cur['id']}").get("description"), d)
            close = closes.get(cur["id"])
            if d["status"] != "open" and cur.get("status") == "open" and close and close["status"] == "pending":
                F(rec, "status", d["status"], WAITING, "the AI agent asked to close it; waiting in your CRM approval inbox")
            elif d["status"] != "open" and cur.get("status") == "open" and close and close["status"] == "rejected":
                F(rec, "status", d["status"], DECLINED, "you declined the close in the CRM")
            else:
                checks.append(("status", d["status"], cur.get("status")))
        else:
            checks.append(("status", d["status"], cur.get("status")))
        if d.get("amount_minor"):
            checks += [("amount", d["amount_minor"], cur.get("amount_minor")), ("currency", d["currency"], cur.get("currency"))]
        for f, want, have in checks:
            if want is None:
                continue
            F(rec, f, want, VERIFIED if want == have else CHANGED, "" if want == have else f"the CRM has {have!r}")
        if d.get("lines"):
            offer = next(client.paginate(f"/v1/deals/{cur['id']}/offers", {"limit": 5}), None)
            if not offer:
                F(rec, "offer", len(d["lines"]), MISSING, "the deal has no offer")
            else:
                items = sorted(offer.get("line_items") or [], key=lambda x: x["position"])
                for i, ln in enumerate(d["lines"]):
                    it = items[i] if i < len(items) else None
                    want = (ln["qty"], ln["unit_price_minor"])
                    have = (it["quantity"], it["unit_price_minor"]) if it else None
                    F(rec, f"offer line {i + 1}", f"{want[0]} x {want[1] / 100:,.2f}",
                      VERIFIED if have == want else CHANGED if have else MISSING,
                      "" if have == want else f"the CRM has {have[0]} x {have[1] / 100:,.2f}" if have else "no such line")
    for name, ds in deal_by_name.items():
        if name not in {d["name"] for d in plan.deals.values()} and ds[0].get("source") == "import" and re.match(
                r"(SO|QU|OC|PO|RFQ)", name):
            F(f"deal {name}", "record", name, EXTRA, "an imported deal the archive no longer gives")

    emit("  reading emails, tasks and notes")
    acts = defaultdict(list)
    for kind in ("email", "task", "note"):
        for a in client.paginate("/v1/activities", {"kind": kind, "limit": 200}):
            if a.get("source_system") == pipeline.SOURCE_SYSTEM:
                acts[a["source_id"]].append(a)
    for sid, as_ in acts.items():
        if len(as_) > 1:
            F(f"activity {sid[:60]}", "record", sid, DUPLICATE, f"{len(as_)} activities with this source id")
    planned = {m["message_id"]: m for m in plan.emails if m["date_iso"]}
    planned_ids = set(planned) | {f"task:{r['id']}" for r in plan.tasks} | {f"note:{r['id']}" for r in plan.notes}
    for sid, as_ in acts.items():
        if sid not in planned_ids:
            F(f"activity {as_[0]['kind']} {as_[0].get('subject', '')[:50]!r}", "record", sid, EXTRA,
              "imported from the archive earlier, but not part of the import any more")
    linked_by_mail = set()
    routed = {r["message_id"]: r for r in plan.routing}
    for i, (mid, m) in enumerate(planned.items(), 1):
        rec = f"email {m['date_iso'][:10]} {m['subject'][:50]!r}"
        if i % 100 == 0:
            emit(f"  emails {i}/{len(planned)}")
        if mid not in acts:
            F(rec, "record", mid, MISSING, "not in the CRM")
            continue
        a = acts[mid][0]
        diffs = []
        if a.get("subject") != (m["subject"] or "(no subject)"):
            diffs.append(f"subject {a.get('subject')!r}")
        if a.get("direction") != m["direction"]:
            diffs.append(f"direction {a.get('direction')}")
        if datetime.fromisoformat(a["occurred_at"]) != datetime.fromisoformat(m["date_iso"]):
            diffs.append(f"date {a['occurred_at']}")
        F(rec, "fields", mid, CHANGED if diffs else VERIFIED, "the CRM has " + ", ".join(diffs) if diffs else "")
        want = set()
        for p in [m["from"]] + m["to"] + m["cc"]:
            if p["email"] in contact_id and not pipeline._internal(p["email"]):
                want.add(("contact", contact_id[p["email"]]))
        route = routed.get(mid) or {}
        company = route.get("company") if route.get("company") in company_id else m.get("company")
        if company in company_id:
            want.add(("company", company_id[company]))
        for k in [m.get("deal_key")] + list(route.get("deals") or []):
            if k in deal_id:
                want.add(("deal", deal_id[k]))
        have = {(lk["entity_type"], lk["entity_id"]) for lk in client.get(f"/v1/activities/{a['id']}").get("links") or []}
        linked_by_mail |= have
        names = {**{("company", v): k for k, v in company_id.items()}, **{("deal", v): k for k, v in deal_id.items()},
                 **{("contact", v): k for k, v in contact_id.items()}}
        label = lambda x: f"{x[0]} {names.get(x) or company_name.get(x[1]) or contact_name.get(x[1], x[1])}"  # noqa: E731
        # The CRM lists an address once per mail (a sender who is also on Cc shows only as sender) and adds
        # the importing seat to some inbound mails it displays, so the question is: is every address of the
        # original there, and is the sender the sender
        shown = client.get(f"/v1/activities/{a['id']}/email-presentation")
        addrs = lambda xs: {(x.get("address") or "").lower() for x in xs or []} - {"", pipeline.UNKNOWN_SENDER}  # noqa: E731
        sender = (m["from"]["email"] or "").lower()
        have_all = addrs(shown.get("from")) | addrs(shown.get("to")) | addrs(shown.get("cc"))
        lost = sorted({p["email"].lower() for p in m["to"] + m["cc"]} - have_all)
        bad = ([f"sender: the CRM has {', '.join(sorted(addrs(shown.get('from')))) or 'nobody'}"]
               if sender and sender not in addrs(shown.get("from")) else []) + ([f"missing: {', '.join(lost)}"] if lost else [])
        F(rec, "participants", sender or "-", CHANGED if bad else VERIFIED, "; ".join(bad))
        if want == have:
            F(rec, "links", len(want), VERIFIED)
        else:
            F(rec, "links", ", ".join(sorted(label(x) for x in want)) or "-", CHANGED,
              "; ".join(filter(None, ["missing: " + ", ".join(sorted(label(x) for x in want - have)) if want - have else "",
                                      "extra: " + ", ".join(sorted(label(x) for x in have - want)) if have - want else ""])))
    for kind, rows in (("task", plan.tasks), ("note", plan.notes)):
        for r in rows:
            rec, sid = f"{kind} {r['id']}", f"{kind}:{r['id']}"
            if sid not in acts:
                F(rec, "record", r["subject"], MISSING, "not in the CRM (imported before it was read or reviewed?)")
                continue
            a = acts[sid][0]
            agent_check(rec, a, a.get("body"), r)
            checks = [("subject", r["subject"], a.get("subject"))]
            if kind == "task":
                checks += [("due", r.get("due"), (a.get("due_at") or "")[:10] or None), ("done", bool(r.get("done")), a.get("is_done"))]
            for f, want, have in checks:
                if want is None:
                    continue
                F(rec, f, want, VERIFIED if want == have else CHANGED, "" if want == have else f"the CRM has {have!r}")

    for u in plan.contact_updates:
        rec = f"title {u['email']}"
        cid = contact_id.get(u["email"])
        if not cid or not u.get("title"):
            continue
        have = (client.get_contact(cid).get("title") or "").strip()
        if have == u["title"][:200].strip():
            F(rec, "title", u["title"], VERIFIED)
        elif not have:
            F(rec, "title", u["title"], WAITING, "not on the contact yet: asked for approval, or not imported")
        else:
            F(rec, "title", u["title"], WEAK, f"the CRM already had {have!r}; a read title never overwrites")

    emit("  reading attachments")
    act_id = {mid: acts[mid][0]["id"] for mid in planned if mid in acts}
    expected = defaultdict(set)
    for a in plan.attachments:
        if a["message_id"] in act_id:
            expected[("activity", act_id[a["message_id"]], a["message_id"])].add((a["filename"], a["size"]))
        if a.get("deal_key") in deal_id:
            expected[("deal", deal_id[a["deal_key"]], a["deal_key"])].add((a["filename"], a["size"]))
    for (et, eid, ref), want in expected.items():
        have = {(x["filename"], x.get("byte_size")) for x in client.attachments_of(et, eid)}
        for fn, size in sorted(want):
            rec = f"attachment {fn}"
            if (fn, size) in have:
                F(rec, f"on {et}", ref[:60], VERIFIED)
            elif fn in {h[0] for h in have}:
                F(rec, f"on {et}", ref[:60], CHANGED, f"a file with this name is there, but not {size} bytes")
            else:
                F(rec, f"on {et}", ref[:60], MISSING, "not uploaded (was the import run with attachments?)")

    planned_companies = set(company_id.values())
    for c in companies:
        if c["id"] not in planned_companies and c.get("source") == "import" and (
                c.get("domains") or ("company", c["id"]) in linked_by_mail):
            F(f"company {c['display_name']}", "record", c["display_name"], EXTRA,
              "imported, but the archive check no longer keeps it")
    planned_contacts = set(contact_id.values())
    for ct in contacts:
        if ct["id"] not in planned_contacts and ("contact", ct["id"]) in linked_by_mail:
            F(f"contact {ct.get('primary_email') or ct['full_name']}", "record", ct["full_name"], EXTRA,
              "linked to imported mails, but not part of the import")


# ---------------------------------------------------------------------------------------------
# report
# ---------------------------------------------------------------------------------------------

TITLES = {"source": "source check (original archive -> import plan)", "crm": "CRM check (import plan -> CRM)"}


def field_counts(rows: list[dict]) -> dict[str, Counter]:
    """'contact phone' -> Counter(VERIFIED=97, MISSING=2); deal lines fold into one 'deal line' row."""
    by_field = defaultdict(Counter)
    for r in rows:
        kind = r["record"].split(" ", 1)[0]
        field = "line" if kind == "deal" and r["field"].startswith("line ") else r["field"]
        by_field[f"{kind} {field}"][r["status"]] += 1
    return dict(sorted(by_field.items()))


def summarize(out: Findings, show: int = 25) -> bool:
    failed = False
    for check, title in TITLES.items():
        rows = [r for r in out.rows if r["check"] == check]
        if not rows:
            continue
        print(f"\n{title}")
        for f, cnt in field_counts(rows).items():
            bad = any(s in FAILING for s in cnt)
            failed |= bad
            parts = [f"{n} {s.lower() if s == VERIFIED else s}" for s, n in sorted(cnt.items(), key=lambda x: x[0] != VERIFIED)]
            print(f"  {'FAIL' if bad else 'ok  '}  {f:<28} {', '.join(parts)}")
        problems = [r for r in rows if r["status"] in FAILING]
        for r in problems[:show]:
            print(f"    {r['status']:<9} {r['record'][:60]} | {r['field']}: {r['value'][:50]} | {r['detail'][:110]}")
        if len(problems) > show:
            print(f"    … and {len(problems) - show} more in verify_report.csv")
    return failed


def write(out: Findings, staging: str) -> None:
    with open(os.path.join(staging, "verify_report.json"), "w") as f:
        counts = Counter((r["check"], r["status"]) for r in out.rows)
        fields = {c: {k: dict(v) for k, v in field_counts([r for r in out.rows if r["check"] == c]).items()}
                  for c in TITLES if any(r["check"] == c for r in out.rows)}
        json.dump({"generated": datetime.now().isoformat(timespec="seconds"),
                   "counts": {f"{c}:{s}": n for (c, s), n in sorted(counts.items())}, "fields": fields,
                   "findings": out.rows}, f, ensure_ascii=False, indent=1)
    with open(os.path.join(staging, "verify_report.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["check", "status", "record", "field", "value", "detail"])
        w.writeheader()
        order = {s: i for i, s in enumerate([NOT_FOUND, MISSING, CHANGED, EXTRA, DUPLICATE, WAITING, DECLINED, WEAK, MANUAL])}
        for r in sorted((r for r in out.rows if r["status"] != VERIFIED), key=lambda r: (r["check"], order[r["status"]])):
            w.writerow({k: r[k] for k in w.fieldnames})


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--source", action="store_true", help="only the archive check (no network)")
    ap.add_argument("--crm", action="store_true", help="only the CRM check")
    ap.add_argument("--staging", default=os.path.join(HERE, "staging"))
    ap.add_argument("--root", default=DEFAULT_ROOT, help="the Thunderbird account directory")
    ap.add_argument("--cache", default=os.path.join(HERE, "staging", ".verify_cache.pickle"),
                    help="where the decoded archive is kept between runs (rebuilt when a mailbox file changes)")
    ap.add_argument("--base-url", default=os.environ.get("MARGINCE_BASE_URL", "http://localhost:8080"))
    ap.add_argument("--email", default=os.environ.get("MARGINCE_EMAIL", "admin@demo.test"))
    ap.add_argument("--password", default=os.environ.get("MARGINCE_PASSWORD", "demo-password-123"))
    args = ap.parse_args()
    both = not args.source and not args.crm

    with open(os.path.join(args.staging, "extracted.json")) as f:
        extracted = json.load(f)
    enrichment, reading = load_enrichment(args.staging)
    plan = build_plan(extracted, enrichment)
    print(f"import plan: {len(plan.companies)} companies, {len(plan.contacts)} contacts, {len(plan.deals)} deals, "
          f"{len(plan.emails)} emails, {len(plan.attachments)} attachments"
          + (f", read from the mails: {reading}" if reading else " (nothing read from the mails yet)"))
    out = Findings()
    if args.source or both:
        print("source check")
        print("  reading the archive (cached after the first run)")
        arc = Archive(args.root, args.cache)
        print(f"  {len(arc.mails)} mails, {len(arc.pdfs)} PDFs")
        check_source(plan, extracted, enrichment, arc, os.path.join(args.staging, "files"), out)
    if args.crm or both:
        print("CRM check")
        client = pipeline.MailClient(args.base_url)
        client.login(args.email, args.password)
        check_crm(plan, client, out)
    write(out, args.staging)
    failed = summarize(out)
    print(f"\nreport: {os.path.join(args.staging, 'verify_report.csv')}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
