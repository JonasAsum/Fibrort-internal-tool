"""Layer 1a: walk the Thunderbird mbox tree and turn every message into a plain dict.

Everything here is mechanical: headers, thread keys, direction, the top reply
with the quoted trail cut off, attachments hashed and (for documents) copied to
staging. The folder a message sits in is the first piece of classification
(customer type, business line, company), see classify_folder().
"""
import email
import email.policy
import email.utils
import hashlib
import html
import json
import mailbox
import os
import re
from collections import Counter
from dataclasses import dataclass, field



def _load_owner() -> tuple[str, list[str]]:
    """The mailbox owner is a real person, so name and addresses live in a gitignored owner.json
    (copy owner.example.json) or the MAIL_OWNER_NAME / MAIL_OWNER_ADDRESSES (comma-separated) env vars."""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "owner.json")
    cfg = json.load(open(path)) if os.path.exists(path) else {}
    name = os.environ.get("MAIL_OWNER_NAME") or cfg.get("name")
    addrs = [a.strip().lower() for a in (os.environ.get("MAIL_OWNER_ADDRESSES") or "").split(",") if a.strip()] \
        or [a.lower() for a in cfg.get("addresses", [])]
    if not name or not addrs:
        raise RuntimeError("mailbox owner not configured: copy mail/owner.example.json to mail/owner.json "
                           "and fill it in, or set MAIL_OWNER_NAME and MAIL_OWNER_ADDRESSES")
    return name, addrs


OWNER_NAME, _OWNER_LIST = _load_owner()
OWNER_ADDRESSES = set(_OWNER_LIST)
OWNER_EMAIL = _OWNER_LIST[0]  # the primary address
INTERNAL_DOMAINS = {"fibrort.com", "fibrort.org"}

SKIP_FOLDERS = {"Deleted Items", "Junk", "Trash"}
GENERIC_PURCHASE = {"Invoice", "Logistics", "Oder Confirmation", "Local Invoice", "Purchase", "Purchase Sent",
                    "Order Confirmation Sent", "Invoice Sent"}
CUSTOMER_TYPES = {"Dealer": "dealer", "End User": "end_user", "Manufacturer": "manufacturer"}
BUSINESS_LINES = {"Motion", "Cutting"}
INTERNAL_TOPICS = {"Finance", "IT", "Marketing", "Laepple"}

# Attachments worth storing and uploading. Everything else is recorded in the
# message's attachment list but not copied (logos, inline photos, videos).
DOCUMENT_EXT = {".pdf", ".xls", ".xlsx", ".xlsm", ".csv", ".doc", ".docx", ".ppt", ".pptx", ".txt", ".zip",
                ".dwg", ".dxf", ".step", ".stp", ".msg", ".eml"}

_QUOTE_START = re.compile(
    r"^\s*(?:>|_{5,}|-{3,}\s*(?:Original|Ursprüngliche|Forwarded|原始)|"
    r"(?:From|Von|De|发件人|寄件者)\s*[:：]|"
    r"On .{5,120}wrote\s*:|Am .{5,120}schrieb .{0,80}:|"
    r".{0,120}在\s*\d{4}\s*(?:年|-\d{1,2}-\d{1,2}).{0,80}写道\s*[:：]|-{3,}\s*原始邮件|"  # QQ / Foxmail / 163 quote lines
    r"\*{0,2}From:\*{0,2})", re.I)


# A forwarded or quoted mail keeps its header block in the body ("From: ... To: ... Cc: ...", in English,
# German or Chinese). Only these blocks are read for addresses - never free text, where "cid:image001.png@
# 01DCB885.077DFA50" and Teams ids look like addresses too.
_Q_FROM = re.compile(r"^\s*\*{0,2}(?:From|Von|De|发件人|寄件者|發件人)\s*\*{0,2}\s*[:：]\s*(.*)$", re.I)
_Q_TO = re.compile(r"^\s*\*{0,2}(?:To|An|À|收件人|收件者)\s*\*{0,2}\s*[:：]\s*(.*)$", re.I)
_Q_CC = re.compile(r"^\s*\*{0,2}(?:Cc|Kopie|抄送|副本)\s*\*{0,2}\s*[:：]\s*(.*)$", re.I)
_Q_OTHER = re.compile(r"^\s*\*{0,2}(?:Sent|Date|Gesendet|Datum|Subject|Betreff|Objet|Envoyé|发送时间|已发送|日期|主题|主旨|"
                      r"Importance|Wichtigkeit)\s*\*{0,2}\s*[:：]", re.I)
ADDRESS_RE = re.compile(r"[A-Za-z0-9._%+\-]+@((?:[A-Za-z0-9\-]+\.)+[A-Za-z]{2,24})(?![A-Za-z0-9\-])")


# Outlook puts a warning above mail from rare senders; it is not the sender's text, and the CRM would show it
# as the message ("They wrote last, saying 'Sie erhalten nicht häufig E-Mails von …'")
SENDER_BANNER = re.compile(
    r"(?:Sie erhalten nicht häufig E-Mails von|You don't often get email from|Some people who received this message "
    r"don't often get email from|Bạn thường không nhận được email từ|你通常不会收到来自|Vous ne recevez pas souvent)"
    r".{0,300}?(?:(?:warum dies wichtig ist|why this is important|quan trọng|很重要|pourquoi c'est important)[.。]?"
    r"\s*(?:<https?://aka\.ms/LearnAboutSenderIdentification>)?|<?https?://aka\.ms/LearnAboutSenderIdentification>?)"
    r"[ \t]*\n?", re.I | re.S)


def strip_banners(body: str) -> str:
    return SENDER_BANNER.sub("", body)


def _addrs_in(text: str) -> list[str]:
    """Addresses whose domain ends in a real (alphabetic) TLD; mailto: duplicates collapse."""
    return sorted({m.group(0).lower().strip(".") for m in ADDRESS_RE.finditer(text)})


def _names_in(text: str) -> dict[str, str]:
    """'Tom Smith <Tom.Smith@x.com>; Doe, John <s.e@y.com>' -> {addr: name} as written."""
    out = {}
    text = re.sub(r"<mailto:[^>]*>", "", text.replace("&lt;", "<").replace("&gt;", ">"))
    for part in re.split(r"[;；]", text):
        m = re.match(r"\s*\"?([^\"<>@]{2,80}?)\"?\s*<\s*([^<>\s]+@[^<>\s]+?)\s*>", part)
        if m:
            out[m.group(2).lower().strip(".")] = m.group(1).strip()
    return out


def quoted_headers(body: str) -> list[dict]:
    """The header blocks of quoted/forwarded mails in a body: [{from: [addr], to: [...], cc: [...],
    names: {addr: display name as written}}]."""
    lines = body.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    blocks, cur, field_ = [], None, None
    for i, line in enumerate(lines):
        mf = _Q_FROM.match(line)
        mq = re.match(r"^(.{0,120}?)在\s*\d{4}\s*(?:年|-\d{1,2}-\d{1,2}).{0,80}写道\s*[:：]", line)  # 'Tammy<t@x.com> 在 2026年5月7日 写道：'
        if mq and i > 0 and _addrs_in(mq.group(1).replace("&gt;", ">").replace("&lt;", "<")):
            blocks.append({"from": _addrs_in(mq.group(1).replace("&gt;", ">")), "to": [], "cc": [], "_rest": 99,
                           "names": _names_in(mq.group(1))})
            cur, field_ = None, None
            continue
        if mf and i > 0:
            cur = {"from": [], "to": [], "cc": [], "_rest": 0, "names": {}}
            blocks.append(cur)
            field_ = "from"
            cur["from"] += _addrs_in(mf.group(1))
            cur["names"].update(_names_in(mf.group(1)))
            continue
        if cur is None:
            continue
        mt, mc = _Q_TO.match(line), _Q_CC.match(line)
        if mt or mc:
            field_ = "to" if mt else "cc"
            cur[field_] += _addrs_in((mt or mc).group(1))
            cur["names"].update(_names_in((mt or mc).group(1)))
        elif _Q_OTHER.match(line):
            field_ = None
        elif field_ and line.strip():
            cur[field_] += _addrs_in(line)  # a header value wrapped onto the next line
            cur["names"].update(_names_in(line))
        cur["_rest"] += 1
        if cur["_rest"] > 12:  # past the header block, into the quoted text
            cur, field_ = None, None
    out = []
    for b in blocks:
        b.pop("_rest")
        if b["from"]:
            b["from"] = b["from"][:1]
            out.append(b)
    return out


@dataclass
class FolderInfo:
    path: str
    company: str | None = None
    customer_type: str | None = None  # dealer | end_user | manufacturer | supplier | internal | unfiled
    business_line: str | None = None
    topic: str | None = None
    sent: bool = False


def classify_folder(segments: list[str]) -> FolderInfo:
    """segments: folder names from the account root, e.g.
    ['INBOX','Customer','Dealer','Motion','Acme','Acme Sent']."""
    segs = list(segments)
    sent = False
    if segs and (segs[-1] == "Sent Items" or segs[-1].endswith(" Sent")):
        sent = True
        if segs[-1] != "Sent Items":
            segs = segs[:-1]
    info = FolderInfo(path="/".join(segments), sent=sent)
    if segs and segs[0] == "INBOX":
        segs = segs[1:]
    if not segs or segs == ["Sent Items"]:
        info.customer_type = "unfiled"
        return info
    if segs[0] == "Customer" and len(segs) >= 2:
        info.customer_type = CUSTOMER_TYPES.get(segs[1], "customer")
        if len(segs) >= 3 and segs[2] in BUSINESS_LINES:
            info.business_line = segs[2]
            if len(segs) >= 4:
                info.company = re.sub(r"[0-9a-f]{8}$", "", segs[3]).strip() or segs[3]
        elif len(segs) >= 3:
            info.company = segs[2]
        return info
    if segs[0] == "FIBRO RT":
        if len(segs) >= 3 and segs[1] == "Purchase":
            info.customer_type = "supplier"
            leaf = segs[2]
            if leaf not in GENERIC_PURCHASE:
                info.company = leaf
            else:
                info.topic = leaf
            return info
        info.customer_type = "internal"
        info.topic = segs[1] if len(segs) > 1 else None
        if info.topic == "Purchase":
            info.customer_type = "supplier"
        return info
    info.customer_type = "unfiled"
    return info


def _walk_mboxes(root: str):
    for dp, dn, fn in os.walk(root):
        dn.sort()
        for f in sorted(fn):
            if f.endswith((".msf", ".json", ".DS_Store", ".dat")):
                continue
            p = os.path.join(dp, f)
            if os.path.getsize(p) == 0:
                continue
            with open(p, "rb") as fh:
                if fh.read(5) != b"From ":
                    continue
            rel = os.path.relpath(p, root).split(os.sep)
            segs = [s[:-4] if s.endswith(".sbd") else s for s in rel[:-1]] + [rel[-1]]
            yield p, segs


def _decode(value) -> str:
    return str(value).strip() if value is not None else ""


def _addresses(msg, header: str) -> list[dict]:
    out = []
    try:
        raw = msg.get_all(header) or []
    except Exception:
        raw = []
    for h in raw:
        try:
            parsed = email.utils.getaddresses([str(h)])
        except Exception:
            continue
        for name, addr in parsed:
            addr = addr.strip().lower()
            if "@" in addr:
                out.append({"name": name.strip().strip('"'), "email": addr})
    return out


def _body_text(msg) -> str:
    plain, html_ = None, None
    for part in msg.walk():
        if part.get_content_maintype() != "text" or part.get_filename():
            continue
        try:
            content = part.get_content()
        except Exception:
            payload = part.get_payload(decode=True) or b""
            content = payload.decode(part.get_content_charset() or "utf-8", "replace")
        if part.get_content_type() == "text/plain" and plain is None:
            plain = content
        elif part.get_content_type() == "text/html" and html_ is None:
            html_ = content
    if plain and plain.strip():
        return plain
    if html_:
        h = re.sub(r"(?is)<(style|script|head).*?</\1>", "", html_)
        h = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>|</li>", "\n", h)
        return html.unescape(re.sub(r"<[^>]+>", "", h))
    return ""


def top_reply(text: str) -> str:
    """The new text of a message: everything before the first quoted-reply marker."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    kept = []
    for i, line in enumerate(lines):
        if i > 0 and _QUOTE_START.match(line):
            break
        kept.append(line)
    out = "\n".join(kept)
    return re.sub(r"\n{3,}", "\n\n", out).strip()


def _msg_id(raw: str | None) -> str | None:
    if not raw:
        return None
    m = re.search(r"<([^<>]+)>", raw)
    return (m.group(1) if m else raw).strip().lower() or None


def _is_decoration(name: str, ctype: str, size: int) -> bool:
    ext = os.path.splitext(name)[1].lower()
    if ctype.startswith("image/") and size < 80_000:
        return True
    return bool(re.fullmatch(r"image\d{3}\.(png|jpg|jpeg|gif)", name.lower())) and size < 150_000


def _safe_name(name: str) -> str:
    return re.sub(r"[^\w.\-一-鿿]+", "_", name)[:120] or "file"


@dataclass
class Corpus:
    messages: list[dict] = field(default_factory=list)
    stats: Counter = field(default_factory=Counter)
    errors: list = field(default_factory=list)


def read_corpus(root: str, files_dir: str | None = None, only_folder: str | None = None) -> Corpus:
    """Reads every mailbox under `root` (the account directory, e.g. .../ImapMail/outlook.office365.com).
    Messages that appear in several folders are merged: the deepest, most specific folder wins."""
    corpus = Corpus()
    by_key: dict[str, dict] = {}
    stored: dict[str, str] = {}  # sha256 -> relative path of the one copy on disk
    if files_dir:
        os.makedirs(files_dir, exist_ok=True)
    for path, segs in _walk_mboxes(root):
        if segs[-1] in SKIP_FOLDERS or segs[0] in SKIP_FOLDERS:
            continue
        folder = classify_folder(segs)
        if only_folder and only_folder.lower() not in folder.path.lower():
            continue
        box = mailbox.mbox(path, factory=lambda f: email.message_from_binary_file(f, policy=email.policy.default))
        for idx, msg in enumerate(box):
            corpus.stats["read"] += 1
            try:
                rec = _parse_message(msg, folder, path, idx, files_dir, stored)
            except Exception as e:  # keep going: one broken message must not stop 1000 others
                corpus.stats["parse_errors"] += 1
                corpus.errors.append(f"{folder.path}#{idx}: {e!r}")
                continue
            key = rec["message_id"]
            prev = by_key.get(key)
            if prev:
                corpus.stats["duplicate_across_folders"] += 1
                prev["folders"].append(folder.path)
                if _specificity(folder) > _specificity(prev["folder_info"]):
                    rec["folders"] = prev["folders"]
                    by_key[key] = rec
                continue
            by_key[key] = rec
    corpus.messages = sorted(by_key.values(), key=lambda r: r["date_iso"] or "")
    for r in corpus.messages:
        r["folder"] = r["folder_info"].__dict__
        del r["folder_info"]
    _assign_threads(corpus.messages)
    return corpus


def _specificity(f: FolderInfo) -> int:
    return (2 if f.company else 0) + (1 if f.topic or f.business_line else 0) + (0 if f.customer_type == "unfiled" else 1)


def _parse_message(msg, folder: FolderInfo, path: str, idx: int, files_dir, stored: dict) -> dict:
    date_hdr = _decode(msg["Date"])
    try:
        dt = email.utils.parsedate_to_datetime(date_hdr)
        date_iso = dt.isoformat()
    except Exception:
        date_iso = None
    sender = (_addresses(msg, "From") or [{"name": "", "email": ""}])[0]
    to, cc = _addresses(msg, "To"), _addresses(msg, "Cc")
    subject = re.sub(r"\s+", " ", _decode(msg["Subject"]))
    body = strip_banners(_body_text(msg))
    message_id = _msg_id(_decode(msg["Message-ID"]))
    if not message_id:
        message_id = "synthetic-" + hashlib.sha1(f"{date_iso}|{sender['email']}|{subject}|{body[:300]}".encode(
            "utf-8", "replace")).hexdigest()
    refs = [r for r in (_msg_id(x) for x in re.findall(r"<[^<>]+>", _decode(msg["References"]))) if r]
    in_reply = _msg_id(_decode(msg["In-Reply-To"]))
    outbound = folder.sent or sender["email"] in OWNER_ADDRESSES
    attachments = []
    for part in msg.walk():
        name = part.get_filename()
        if not name:
            continue
        payload = part.get_payload(decode=True)
        if not payload:
            continue
        ctype = part.get_content_type()
        sha = hashlib.sha256(payload).hexdigest()
        ext = os.path.splitext(name)[1].lower()
        deco = _is_decoration(name, ctype, len(payload))
        rel = None
        if files_dir and not deco and ext in DOCUMENT_EXT:
            if sha not in stored:  # same content under another name points at the first copy
                stored[sha] = os.path.join(sha[:2], f"{sha[:12]}_{_safe_name(name)}")
                full = os.path.join(files_dir, stored[sha])
                os.makedirs(os.path.dirname(full), exist_ok=True)
                with open(full, "wb") as fh:
                    fh.write(payload)
            rel = stored[sha]
        attachments.append({"filename": name, "content_type": ctype, "size": len(payload), "sha256": sha,
                            "decoration": deco, "stored_as": rel})
    return {
        "message_id": message_id,
        "source": f"{os.path.basename(path)}#{idx}",
        "folder_info": folder,
        "folders": [folder.path],
        "date": date_hdr,
        "date_iso": date_iso,
        "from": sender, "to": to, "cc": cc,
        "subject": subject,
        "direction": "outbound" if outbound else "inbound",
        "in_reply_to": in_reply,
        "references": refs,
        "body": body[:60_000],
        "top": top_reply(body)[:20_000],
        "attachments": attachments,
        # newsletter / automated mail: says nothing about a business relationship
        "bulk": bool(msg["List-Unsubscribe"] or msg["List-Id"]
                     or re.search(r"bulk|list|junk", _decode(msg["Precedence"]), re.I)
                     or re.search(r"auto-(generated|replied)", _decode(msg["Auto-Submitted"]), re.I)),
        "quoted": quoted_headers(body)[:20],
    }


def _assign_threads(messages: list[dict]) -> None:
    """thread_key = the root Message-ID of the conversation. Falls back to the
    normalised subject when a client dropped the reference headers."""
    root_of: dict[str, str] = {}
    by_id = {m["message_id"]: m for m in messages}

    def root(mid: str, seen=()) -> str:
        if mid in root_of:
            return root_of[mid]
        m = by_id.get(mid)
        if not m or mid in seen:
            return mid
        parent = (m["references"][0] if m["references"] else m["in_reply_to"])
        r = root(parent, seen + (mid,)) if parent and parent != mid else mid
        root_of[mid] = r
        return r

    subject_root: dict[str, str] = {}
    for m in messages:
        r = root(m["message_id"])
        if r == m["message_id"] and not (m["references"] or m["in_reply_to"]):
            norm = re.sub(r"^((re|fw|fwd|aw|wg|回复|答复)\s*[:：]\s*)+", "", m["subject"], flags=re.I).strip().lower()
            if norm and re.search(r"^(re|fw|fwd|aw|wg|回复|答复)\s*[:：]", m["subject"], re.I):
                r = subject_root.get(norm, r)
            if norm:
                subject_root.setdefault(norm, r)
        m["thread_key"] = r
