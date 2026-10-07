"""Checks the two readings of a case (READING.md) and merges them into staging/enrichment/<case>.json.

    .venv/bin/python -m mail.reading --company Acme   # check pass A and B, merge, say what needs review
    .venv/bin/python -m mail.reading --all              # every case that has a reading
    .venv/bin/python -m mail.reading --status           # one line per case

Check: every quote must be found in the cited mail as the original archive has it (verify.Archive, its own
decoder), every number in a fact inside one of its quotes, every name and company from the case file. An item
that fails is rejected, with the reason. Scans have no text layer, so their quotes cannot be checked by code.

Merge: an item both readings agree on is `auto`. Money (deal amounts and lines), scan facts, and every
disagreement (different values, or found by one reading only) are `needs_review` and stay out of the CRM until
Jonas decides in the web UI's review tab. His decisions are kept when the readings are merged again, as long as
the two readings of that item are unchanged.
"""
import argparse
import hashlib
import json
import os
import re
import sys
from datetime import datetime

from .cases import slug
from .verify import Archive, amount_forms, norm, ref_forms
from .extract import DEFAULT_ROOT

HERE = os.path.dirname(os.path.abspath(__file__))
KINDS = ("deal", "task", "note", "routing", "contact", "document")
COMPARE = {"deal": ("status", "amount_minor", "currency", "expected_close", "lines"), "task": ("due", "done"),
           "note": ("type",), "routing": ("company", "deal"), "contact": ("title",),
           "document": ("type", "number", "date", "total_minor", "currency")}
MONTHS = "jan feb mar apr may jun jul aug sep oct nov dec".split()


def h8(s: str) -> str:
    return hashlib.sha1(s.encode()).hexdigest()[:8]


def date_forms(iso: str) -> list[str]:
    """2026-04-08 as mails write it: 2026/4/8, 4月8日, 08.04.2026, 8 Apr, Apr 8 …"""
    y, m, d = iso[:10].split("-")
    mi, di, mon = int(m), int(d), MONTHS[int(m) - 1]
    forms = [f"{y}-{m}-{d}", f"{y}/{m}/{d}", f"{y}.{m}.{d}", f"{y}/{mi}/{di}", f"{y}-{mi}-{di}", f"{y}年{mi}月{di}日",
             f"{mi}月{di}日", f"{d}.{m}.{y}", f"{di}.{mi}.{y}", f"{d}.{m}.", f"{di} {mon}", f"{mon} {di}", f"{m}/{d}",
             f"{mi}/{di}"]
    return [norm(f) for f in forms]


class Checker:
    """Resolves aliases and checks quotes for one case."""

    def __init__(self, case: dict, arc: Archive, extracted: dict):
        self.case, self.arc = case, arc
        self.slug = slug(case["case"])
        self.kept = set(case["known"]["companies"])
        self.emails = {e.lower() for c in extracted["contacts"] for e in c["emails"]}
        self.source = {m["message_id"]: m["source"] for m in extracted["messages"]}

    def evidence(self, items, problems: list, field: str) -> list[dict]:
        out = []
        for ev in items or []:
            alias, quote = str(ev.get("m") or "").strip(), str(ev.get("q") or "").strip()
            row = {"field": field, "alias": alias, "quote": quote, "message_id": None, "sha256": None,
                   "verified": False, "scan": False}
            if not quote or len(quote) > 400:
                problems.append(f"{field}: {'empty' if not quote else 'over-long'} quote from {alias or '?'}")
            elif alias in self.case["messages"]:
                mid = self.case["messages"][alias]["message_id"]
                row["message_id"] = mid
                raw = self.arc.find(mid, self.source.get(mid, ""))
                q = norm(quote)
                row["verified"] = bool(raw) and (q in raw["text"] or any(q in self.arc.pdfs.get(s, "") for s in raw["atts"]))
                if not row["verified"]:
                    problems.append(f"{field}: quote not found in {alias}: {quote[:60]!r}")
            elif alias in self.case["documents"]:
                doc = self.case["documents"][alias]
                row["sha256"] = doc["sha256"]
                if doc["kind"] == "scan":
                    row["scan"] = True  # no text layer: checked by the second reading and by Jonas
                else:
                    row["verified"] = norm(quote) in self.arc.pdfs.get(doc["sha256"], "")
                    if not row["verified"]:
                        problems.append(f"{field}: quote not found in {alias} {doc['filename']}: {quote[:60]!r}")
            else:
                problems.append(f"{field}: {alias or '(no alias)'} is not a mail or attachment of this case")
            out.append(row)
        return out

    def messages(self, aliases, problems: list) -> list[str]:
        out = []
        for a in aliases or []:
            if a in self.case["messages"]:
                out.append(self.case["messages"][a]["message_id"])
            else:
                problems.append(f"messages: {a} is not a mail of this case")
        return out

    def message_date(self, mid: str) -> str:
        return next((m["date"] or "" for m in self.case["messages"].values() if m["message_id"] == mid), "")

    def thread_head(self, mid: str) -> str:
        t = next((m["thread"] for m in self.case["messages"].values() if m["message_id"] == mid), None)
        return self.case["messages"][self.case["threads"][t][0]]["message_id"] if t else mid

    @staticmethod
    def has_number(evidence: list[dict], forms: list[str]) -> bool:
        return any(f in norm(e["quote"]) for e in evidence for f in forms)

    # ---- one item per kind; returns (id, resolved item, problems) ----

    def company(self, raw: dict, p: list) -> str | None:
        """The company an item is about: the case's, or another one from the list when the item names it."""
        named = raw.get("company")
        if named and named not in self.kept:
            p.append(f"company {named!r} is not in the list")
        return named if named in self.kept else self.case["company"]

    def deal(self, d: dict):
        p: list[str] = []
        ev = {f: self.evidence((d.get("evidence") or {}).get(f), p, f) for f in ("status", "amount", "expected_close", "lines")}
        key = (d.get("key") or "").strip()
        if key:
            if not any(self.arc.anywhere(f) for f in ref_forms(key)):
                p.append(f"the number {key} appears in no mail or PDF")
        else:
            anchor = d.get("anchor")
            if anchor not in self.case["messages"]:
                p.append("an inquiry without a number needs an anchor mail")
                key = f"INQ-{self.slug}-?"
            else:
                key = f"INQ-{self.slug}-{(self.case['messages'][anchor]['date'] or '')[:10].replace('-', '')}"
        status = d.get("status") or "open"
        if status not in ("open", "won", "lost"):
            p.append(f"status {status!r} is not open/won/lost")
        if status != "open" and not ev["status"]:
            p.append(f"status {status} has no quote")
        amount = d.get("amount_minor")
        if amount:
            if not ev["amount"]:
                p.append("an amount needs a quote")
            elif not self.has_number(ev["amount"], amount_forms(int(amount))):
                p.append(f"the amount {int(amount) / 100:,.2f} is not in its quotes")
        if d.get("expected_close") and not ev["expected_close"]:
            p.append("expected_close needs a quote")
        lines = d.get("lines") or []
        for i, ln in enumerate(lines, 1):
            if ln.get("unit_price_minor") and not self.has_number(ev["lines"], amount_forms(int(ln["unit_price_minor"]))):
                p.append(f"line {i}: the unit price is not in the line quotes")
        item = {"key": key, "company": self.case["company"], "status": status, "amount_minor": amount,
                "currency": d.get("currency") if amount else None, "expected_close": d.get("expected_close"),
                "lines": lines, "messages": self.messages(d.get("messages"), p), "summary": d.get("summary") or "",
                "evidence": [e for f in ev.values() for e in f], "confidence": d.get("confidence"),
                "scan": any(e["scan"] for f in ev.values() for e in f)}
        return f"deal:{key}", item, p

    def task(self, t: dict):
        p: list[str] = []
        ev = {f: self.evidence((t.get("evidence") or {}).get(f), p, f) for f in ("task", "due", "done")}
        if not ev["task"]:
            p.append("a task needs a quote")
        anchor = next((e["message_id"] for e in ev["task"] if e["message_id"]), None)
        due = t.get("due")
        literal = None
        if due:
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(due)):
                p.append(f"due {due!r} is not YYYY-MM-DD")
            elif not ev["due"]:
                p.append("a due date needs a quote")
            else:
                literal = self.has_number(ev["due"], date_forms(due))
        done = bool(t.get("done"))
        if done:
            later = [e for e in ev["done"] if e["message_id"] and anchor and self.message_date(e["message_id"]) > self.message_date(anchor)]
            if not later:
                p.append("done needs a quote from a later mail")
        item = {"subject": (t.get("subject") or "").strip(), "body": t.get("body") or "", "due": due, "done": done,
                "due_literal": literal, "company": self.company(t, p), "anchor": anchor,
                "evidence": [e for f in ev.values() for e in f], "confidence": t.get("confidence")}
        if not item["subject"]:
            p.append("a task needs a subject")
        return f"task:{self.slug}:{h8(anchor or item['subject'])}", item, p  # same mail = same task; a differing due date is a dispute

    def note(self, n: dict):
        p: list[str] = []
        ev = self.evidence((n.get("evidence") or {}).get("note"), p, "note")
        if not ev:
            p.append("a note needs a quote")
        kind = n.get("type") or "case"
        if kind not in ("case", "purchase"):
            p.append(f"note type {kind!r} is not case/purchase")
        msgs = self.messages(n.get("messages"), p)
        head = self.thread_head(msgs[0] if msgs else next((e["message_id"] for e in ev if e["message_id"]), ""))
        item = {"type": kind, "subject": (n.get("subject") or "").strip(), "body": n.get("body") or "",
                "company": self.company(n, p), "messages": msgs, "date": self.message_date(msgs[0]) if msgs else None,
                "evidence": ev}
        return f"note:{self.slug}:{h8(head)}:{kind}", item, p

    def routing(self, r: dict):
        p: list[str] = []
        ev = self.evidence(r.get("evidence"), p, "routing")
        alias = r.get("m")
        mid = self.case["messages"].get(alias, {}).get("message_id")
        if not mid:
            p.append(f"{alias} is not a mail of this case")
        company = r.get("company")
        if company not in self.kept:
            p.append(f"company {company!r} is not in the list")
        deal = r.get("deal")
        if deal and not any(self.arc.anywhere(f) for f in ref_forms(deal)):
            p.append(f"deal {deal} appears in no mail")
        if not ev:
            p.append("routing needs a quote")
        return f"routing:{h8(mid or str(alias))}", {"message_id": mid, "company": company, "deal": deal, "evidence": ev}, p

    def contact(self, c: dict):
        p: list[str] = []
        ev = self.evidence(c.get("evidence"), p, "title")
        email, title = (c.get("email") or "").lower(), (c.get("title") or "").strip()
        if email not in self.emails:
            p.append(f"{email} is not a known contact")
        if not title or not any(norm(title) in norm(e["quote"]) for e in ev):
            p.append("the title must be inside its quote")
        return f"contact:{email}:title", {"email": email, "title": title, "evidence": ev}, p

    def document(self, d: dict):
        p: list[str] = []
        ev = self.evidence(d.get("evidence"), p, "document")
        doc = self.case["documents"].get(d.get("d"))
        if not doc:
            p.append(f"{d.get('d')} is not an attachment of this case")
        total = d.get("total_minor")
        if total and not self.has_number(ev, amount_forms(int(total))):
            p.append(f"the total {int(total) / 100:,.2f} is not in its quotes")
        if d.get("number") and not any(f in norm(e["quote"]) for e in ev for f in ref_forms(str(d["number"]))):
            p.append(f"the number {d['number']} is not in its quotes")
        item = {"sha256": doc["sha256"] if doc else None, "filename": doc["filename"] if doc else None,
                "type": d.get("type"), "number": d.get("number"), "date": d.get("date"), "total_minor": total,
                "currency": d.get("currency") if total else None, "summary": d.get("summary") or "",
                "company": self.case["company"], "evidence": ev}
        return f"document:{(doc or {}).get('sha256', 'x')[:12]}", item, p

    def check(self, reading: dict) -> dict[str, dict]:
        """{item id: {"kind", "item", "problems"}} for one reading."""
        out = {}
        for kind, key in (("deal", "deals"), ("task", "tasks"), ("note", "notes"), ("routing", "routing"),
                          ("contact", "contacts"), ("document", "documents")):
            for raw in reading.get(key) or []:
                iid, item, problems = getattr(self, kind)(raw)
                # anything that rests on a scan (no text layer, so code cannot check its quote) goes to review
                item["scan"] = item.get("scan") or any(e.get("scan") for e in item.get("evidence") or [])
                while iid in out:  # two tasks on one mail with one due date
                    iid += "+"
                out[iid] = {"kind": kind, "item": item, "problems": problems}
        return out


def _fields(kind: str, item: dict, code_deals: set) -> tuple:
    # a deal code made from its contract keeps the contract's values whatever a reading says: only its status
    # is compared (plan.build_plan takes nothing else from a reading for it)
    return ("status",) if kind == "deal" and item.get("key") in code_deals else COMPARE[kind]


def _comparable(kind: str, item: dict, fields: tuple) -> dict:
    out = {}
    for f in fields:
        v = item.get(f)
        if f == "lines":  # only priced lines are money; a bare "2 x VR.NC.17" belongs in the summary
            v = sorted((ln.get("qty"), ln.get("unit_price_minor")) for ln in v or [] if ln.get("unit_price_minor"))
        elif f == "title":
            v = norm(v)
        out[f] = v
    return out


def _mails(kind: str, item: dict) -> set:
    ev = {e["message_id"] for e in item.get("evidence") or [] if e.get("message_id")}
    return ev | (set(item.get("messages") or []) if kind == "note" else set())


def _pair_up(a: dict, b: dict) -> dict:
    """Two readers often anchor the same task or note on different mails of one case. A B item without an
    exact counterpart takes the id of the A item of its kind (and note type) it shares the most mails with."""
    taken = set(b) & set(a)
    out = {}
    for iid, vb in sorted(b.items()):
        if iid in a or vb["kind"] not in ("task", "note"):
            out[iid] = vb
            continue
        mine = _mails(vb["kind"], vb["item"])
        best, overlap = None, 0
        for aid, va in a.items():
            if aid in taken or va["kind"] != vb["kind"] or va["item"].get("type") != vb["item"].get("type"):
                continue
            n = len(mine & _mails(va["kind"], va["item"]))
            if n > overlap:
                best, overlap = aid, n
        if best:
            taken.add(best)
            out[best] = vb
        else:
            out[iid] = vb
    return out


def merge(case: dict, a: dict | None, b: dict | None, previous: dict | None) -> dict:
    old = {it["id"]: it for it in (previous or {}).get("items", [])}
    code_deals = {d["key"] for d in case["known"]["deals_from_documents"]}
    if a and b:
        b = _pair_up(a, b)
    items = []
    for iid in sorted(set(a or {}) | set(b or {})):
        va, vb = (a or {}).get(iid), (b or {}).get(iid)
        kind = (va or vb)["kind"]
        ok_a = va and not va["problems"]
        ok_b = vb and not vb["problems"]
        reasons = []
        if not ok_a and not ok_b:
            state, agreement, value = "rejected", "rejected", None
        elif ok_a and ok_b:
            fields = _fields(kind, va["item"], code_deals)
            diffs = [f for f in fields if _comparable(kind, va["item"], fields).get(f) != _comparable(kind, vb["item"], fields).get(f)]
            agreement = "disputed" if diffs else "agreed"
            if diffs:
                reasons.append("the readings differ on " + ", ".join(diffs))
            value = None if diffs else va["item"]
            state = "needs_review" if diffs else "auto"
        else:
            agreement = "single"
            reasons.append(f"only reading {'A' if ok_a else 'B'} found this" + (
                f" (reading {'B' if ok_a else 'A'} was rejected)" if (va and vb) else ""))
            value, state = None, "needs_review"
        if state != "rejected":
            chosen = (va if ok_a else vb)["item"]
            if kind == "deal" and chosen.get("key") not in code_deals and (
                    chosen.get("amount_minor") or any(ln.get("unit_price_minor") for ln in chosen.get("lines") or [])):
                reasons.append("money")
            if kind == "document" or chosen.get("scan"):
                reasons.append("read from a scan")
            if reasons and state == "auto":
                state = "needs_review"
        fp = h8(json.dumps([va and va["item"], vb and vb["item"]], ensure_ascii=False, sort_keys=True, default=str))
        review = old.get(iid, {}).get("review") if old.get(iid, {}).get("fp") == fp else None
        items.append({"id": iid, "kind": kind, "state": state, "agreement": agreement, "reasons": reasons,
                      "a": va and va["item"], "b": vb and vb["item"],
                      "a_problems": (va or {}).get("problems", []), "b_problems": (vb or {}).get("problems", []),
                      "value": value, "fp": fp, "review": review})
    return {"version": 2, "case": case["case"], "company": case["company"], "case_digest": case["digest"],
            "merged_at": datetime.now().isoformat(timespec="seconds"), "items": items}


def final_value(item: dict) -> dict | None:
    """What goes into the CRM: auto items as merged, reviewed items as decided, nothing else."""
    r = item.get("review")
    if r:
        if r["choice"] == "reject":
            return None
        if r["choice"] == "edit":
            return r["value"]
        return item[r["choice"]]
    return item["value"] if item["state"] == "auto" else None


def save_review(path: str, item_id: str, choice: str, value: dict | None, by: str) -> dict:
    with open(path) as fh:
        doc = json.load(fh)
    item = next(it for it in doc["items"] if it["id"] == item_id)
    if choice not in ("a", "b", "edit", "reject") or (choice in ("a", "b") and not item[choice]):
        raise ValueError(f"cannot choose {choice!r} for {item_id}")
    item["review"] = {"choice": choice, "value": value if choice == "edit" else None, "by": by,
                      "at": datetime.now().isoformat(timespec="seconds")}
    tmp = path + ".tmp"
    with open(tmp, "w") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=1)
    os.replace(tmp, path)
    return item


def enrichment_dir(staging: str) -> str:
    return os.path.join(staging, "enrichment")


def load_for_plan(staging: str) -> tuple[dict | None, dict]:
    """The merged readings in the shape build_plan takes, plus counts. None when nothing has been read yet."""
    folder = enrichment_dir(staging)
    files = sorted(f for f in os.listdir(folder) if f.endswith(".json")) if os.path.isdir(folder) else []
    if not files:
        return None, {}
    out = {"deals": [], "tasks": [], "notes": [], "routing": {}, "contacts": {}}
    counts = {"auto": 0, "accepted": 0, "rejected_by_review": 0, "waiting_for_review": 0, "rejected_by_check": 0}
    for f in files:
        with open(os.path.join(folder, f)) as fh:
            doc = json.load(fh)
        for it in doc["items"]:
            if it["state"] == "rejected":
                counts["rejected_by_check"] += 1
                continue
            v = final_value(it)
            if v is None:
                counts["rejected_by_review" if it.get("review") else "waiting_for_review"] += 1
                continue
            counts["accepted" if it.get("review") else "auto"] += 1
            reading = {"case": doc["case"], "item": it["id"], "agreement": it["agreement"],
                       "reviewed_by": (it.get("review") or {}).get("by"), "reviewed_at": (it.get("review") or {}).get("at")}
            ev = [{k: e[k] for k in ("message_id", "sha256", "quote", "scan")} for e in v.get("evidence", [])]
            if it["kind"] == "deal":
                out["deals"].append({"key": v["key"], "company": v["company"], "status": v["status"],
                                     "amount_minor": v["amount_minor"], "currency": v["currency"] or "CNY",
                                     "expected_close": v["expected_close"], "lines": v["lines"], "note": v["summary"],
                                     "message_ids": v["messages"], "evidence": ev, "reading": reading})
            elif it["kind"] == "task":
                out["tasks"].append({"id": it["id"], "subject": v["subject"], "body": v["body"], "due": v["due"],
                                     "done": v["done"], "company": v["company"], "evidence": ev, "reading": reading})
            elif it["kind"] == "note":
                out["notes"].append({"id": it["id"], "type": v["type"], "subject": v["subject"], "body": v["body"],
                                     "company": v["company"], "date": v["date"], "evidence": ev, "reading": reading})
            elif it["kind"] == "document":
                total = f" — total {v['currency'] or ''} {v['total_minor'] / 100:,.2f}".rstrip() if v["total_minor"] else ""
                out["notes"].append({"id": it["id"], "type": "document",
                                     "subject": f"{(v['type'] or 'document').replace('_', ' ').capitalize()} {v['number'] or v['filename']}",
                                     "body": f"{v['summary']}{total} ({v['filename']}, read from the scan)",
                                     "company": v["company"], "date": v["date"], "evidence": ev, "reading": reading})
            elif it["kind"] == "routing":
                out["routing"][v["message_id"]] = {"company": v["company"], "deal": v["deal"], "evidence": ev,
                                                   "reading": reading}
            elif it["kind"] == "contact":
                out["contacts"][v["email"]] = {"title": v["title"], "phones": [], "evidence": ev, "reading": reading}
    return out, counts


def run_case(staging: str, name_or_slug: str, arc: Archive, extracted: dict) -> dict:
    folder = os.path.join(staging, "reading", slug(name_or_slug) if not os.path.isdir(
        os.path.join(staging, "reading", name_or_slug)) else name_or_slug)
    with open(os.path.join(folder, "case.json")) as fh:
        case = json.load(fh)
    checker = Checker(case, arc, extracted)
    readings = {}
    for p in ("a", "b"):
        path = os.path.join(folder, f"pass_{p}.json")
        if os.path.exists(path):
            with open(path) as fh:
                readings[p] = checker.check(json.load(fh))
    os.makedirs(enrichment_dir(staging), exist_ok=True)
    out_path = os.path.join(enrichment_dir(staging), f"{slug(case['case'])}.json")
    previous = None
    if os.path.exists(out_path):
        with open(out_path) as fh:
            previous = json.load(fh)
    doc = merge(case, readings.get("a"), readings.get("b"), previous)
    doc["readers"] = sorted(readings)
    with open(out_path, "w") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=1)
    return doc


def summary_line(doc: dict) -> str:
    c = {}
    for it in doc["items"]:
        key = "reviewed" if it.get("review") else it["state"]
        c[key] = c.get(key, 0) + 1
    return ", ".join(f"{n} {k.replace('_', ' ')}" for k, n in sorted(c.items())) or "nothing found"


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--company", help="a case name or its folder name")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--status", action="store_true")
    ap.add_argument("--pending", metavar="CASE", help="print the items of a merged case that still need a decision")
    ap.add_argument("--decide", nargs=3, metavar=("CASE", "ITEM_ID", "CHOICE"), help="record a decision: a, b, edit or reject")
    ap.add_argument("--value", help="with --decide edit: the item's corrected value as JSON")
    ap.add_argument("--by", default="Claude (adjudicator)", help="who decided (shown in the CRM)")
    ap.add_argument("--staging", default=os.path.join(HERE, "staging"))
    ap.add_argument("--root", default=DEFAULT_ROOT)
    args = ap.parse_args()
    reading_dir = os.path.join(args.staging, "reading")
    folders = sorted(os.listdir(reading_dir)) if os.path.isdir(reading_dir) else []

    if args.pending or args.decide:
        name = args.pending or args.decide[0]
        path = os.path.join(enrichment_dir(args.staging), f"{slug(name) if not os.path.exists(os.path.join(enrichment_dir(args.staging), name + '.json')) else name}.json")
        if args.decide:
            item = save_review(path, args.decide[1], args.decide[2], json.loads(args.value) if args.value else None, args.by)
            print("saved:", item["id"], item["review"]["choice"])
            return
        with open(path) as fh:
            doc = json.load(fh)
        todo = [it for it in doc["items"] if it["state"] == "needs_review" and not it.get("review")]
        print(json.dumps([{k: it[k] for k in ("id", "kind", "agreement", "reasons", "a", "b", "a_problems", "b_problems")}
                          for it in todo], ensure_ascii=False, indent=1))
        print(f"{len(todo)} item(s) need a decision", file=sys.stderr)
        return

    if args.status:
        print(f"{'case':42} {'readings':9} merged")
        for f in folders:
            with open(os.path.join(reading_dir, f, "case.json")) as fh:
                case = json.load(fh)
            passes = "".join(p.upper() for p in "ab" if os.path.exists(os.path.join(reading_dir, f, f"pass_{p}.json")))
            out = os.path.join(enrichment_dir(args.staging), f"{f}.json")
            merged = "-"
            if os.path.exists(out):
                with open(out) as fh:
                    doc = json.load(fh)
                merged = summary_line(doc) + ("  (STALE: the case changed, merge again)" if doc["case_digest"] != case["digest"] else "")
            print(f"{case['case'][:42]:42} {passes or '-':9} {merged}")
        return

    with open(os.path.join(args.staging, "extracted.json")) as f:
        extracted = json.load(f)
    arc = Archive(args.root, os.path.join(args.staging, ".verify_cache.pickle"))
    targets = [f for f in folders if any(os.path.exists(os.path.join(reading_dir, f, f"pass_{p}.json")) for p in "ab")] \
        if args.all else [args.company] if args.company else []
    if not targets:
        sys.exit("name a case with --company, or use --all / --status")
    bad = False
    for t in targets:
        doc = run_case(args.staging, t, arc, extracted)
        print(f"\n{doc['case']}: readings {', '.join(doc['readers']) or 'none'} -> {summary_line(doc)}")
        for it in doc["items"]:
            probs = [f"A: {p}" for p in it["a_problems"]] + [f"B: {p}" for p in it["b_problems"]]
            if it["state"] != "auto" or probs:
                label = (it["a"] or it["b"] or {}).get("subject") or (it["a"] or it["b"] or {}).get("key") or ""
                print(f"  {it['state']:12} {it['id'][:48]:48} {str(label)[:40]:40} {'; '.join(it['reasons'] + probs)[:160]}")
                bad |= it["state"] == "rejected"
    sys.exit(0)


if __name__ == "__main__":
    main()
