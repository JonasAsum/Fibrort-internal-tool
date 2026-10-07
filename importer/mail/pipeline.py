"""Layer 3: writes a Plan into Margince. Every record carries source_system + a stable source_id, and
every lookup happens before a create, so running it twice converges instead of duplicating.

Order: vocabulary -> companies -> contacts -> products -> deals (+ offers) -> emails -> routing -> titles
-> tasks/notes -> attachments.

What code copied out of the mails is written with the human login. What the reading layer understood (deals
found by reading, tasks, notes, mails moved to another company or deal, titles) is written with the AI
agent's passport (common/passport.py), so the CRM stamps it agent:<id>. Every such record ends with the
quotes it was read from. A deal the agent closes as won or lost waits in the CRM approval inbox; an approved
close is completed on the next run, because only the agent that asked may complete it.
"""
import hashlib
import json
import mimetypes
import os
from dataclasses import dataclass, field
from datetime import date

from common.client import ApiError, MargenceClient
from common.fields import ensure_field
from .plan import Plan, fold

SOURCE_SYSTEM = "mailarchive"
BODY_LIMIT = 20_000
UNKNOWN_SENDER = "unknown-sender@unknown.invalid"  # bounces/notifications with no usable From; the API requires one

COMPANY_FIELDS = [
    # "Relationship" is refused as a label (structural keyword); the old "Partner type" field held other options
    ("relationship", "Partner category", "picklist",
     ["Customer", "Supplier", "Service provider", "Lead", "Third party", "Partner", "FIBRO group"]),
    ("customer_type", "Customer type", "picklist", ["Dealer", "End user", "Manufacturer"]),
    ("business_line", "Business line", "picklist", ["Motion", "Cutting"]),
    ("legal_name_en", "Legal name (EN)", "text", None),
]


class MailClient(MargenceClient):
    def post_h(self, path: str, json_: dict, key: str | None = None) -> tuple[int, dict]:
        if key and not key.isascii():  # HTTP headers are latin-1: a Chinese company name becomes a hash
            key = "h:" + hashlib.sha256(key.encode()).hexdigest()
        headers = {"Idempotency-Key": key[:200]} if key else {}
        r = self._request("POST", path, json=json_, headers=headers)
        return r.status_code, (r.json() if r.content else {})

    def all_products(self) -> dict[str, dict]:
        return {p.get("sku") or p["name"]: p for p in self.paginate("/v1/products", {"limit": 200})}

    def pipelines(self) -> list[dict]:
        out = list(self.paginate("/v1/pipelines", {"limit": 100}))
        for p in out:
            if not p.get("stages"):
                p["stages"] = list(self.paginate(f"/v1/pipelines/{p['id']}/stages", {"limit": 100}))
        return out

    def all_deals(self) -> dict[str, dict]:
        return {d["name"]: d for d in self.paginate("/v1/deals", {"limit": 200})}

    def all_contacts(self) -> list[dict]:
        return list(self.paginate("/v1/contacts", {"limit": 200}))

    def upload_attachment(self, entity_type: str, entity_id: str, path: str, filename: str) -> dict:
        ctype = mimetypes.guess_type(filename)[0] or "application/octet-stream"
        with open(path, "rb") as fh:
            r = self._request("POST", "/v1/attachments", data={"entity_type": entity_type, "entity_id": entity_id},
                              files={"file": (filename, fh, ctype)})
        return r.json()

    def attachments_of(self, entity_type: str, entity_id: str) -> list[dict]:
        return list(self.paginate("/v1/attachments", {"entity_type": entity_type, "entity_id": entity_id,
                                                      "limit": 200}))


@dataclass
class Report:
    created: dict = field(default_factory=dict)
    existing: dict = field(default_factory=dict)
    failed: list = field(default_factory=list)
    approvals: dict = field(default_factory=dict)  # won/lost closes and title fills that wait for a human

    def note(self, what: str) -> None:
        self.approvals[what] = self.approvals.get(what, 0) + 1

    def made(self, kind: str, new: bool) -> None:
        bucket = self.created if new else self.existing
        bucket[kind] = bucket.get(kind, 0) + 1

    def fail(self, kind: str, ref: str, err: Exception) -> None:
        self.failed.append({"kind": kind, "ref": ref, "error": str(err)[:500]})


class Run:
    def __init__(self, client: MailClient, plan: Plan, files_dir: str, attachments: str, emit=print,
                 agent: MailClient | None = None, agent_id: str | None = None):
        self.c, self.plan, self.files_dir, self.attachments, self.emit = client, plan, files_dir, attachments, emit
        # read facts go in under the agent; without a passport they are still marked in their text
        self.ai, self.agent_id = agent or client, agent_id if agent else None
        self.ai_marker = "" if agent else "[AI-read] "
        self.msg_info = {m["message_id"]: m for m in plan.emails}
        self.doc_names = {a["sha256"]: a["filename"] for m in plan.emails for a in m["attachments"]}
        self.report = Report()
        self.company_ids: dict[str, str] = {}
        self.contact_ids: dict[str, str] = {}  # email -> id
        self.product_ids: dict[str, str] = {}
        self.deal_ids: dict[str, str] = {}
        self.activity_ids: dict[str, str] = {}  # message_id -> id
        self.fields: dict[str, str] = {}
        self.me = client.get("/v1/me")["user"]["id"]

    def provenance(self, evidence, reading) -> str:
        """The quotes a read fact stands on, as the record's last lines in the CRM."""
        lines = []
        if isinstance(evidence, str):
            lines.append(f"Read from: {evidence}")
        for e in (evidence if isinstance(evidence, list) else [])[:3]:
            if e.get("message_id"):
                m = self.msg_info.get(e["message_id"])
                src = f"{(m['date_iso'] or '')[:10]} {m['from']['name'] or m['from']['email']}" if m else "a mail"
            else:
                src = self.doc_names.get(e.get("sha256"), "an attachment") + (" (scan)" if e.get("scan") else "")
            lines.append(f"Read from {src}: “{e['quote']}”")
        if reading and reading.get("reviewed_by"):
            lines.append(f"Checked by {reading['reviewed_by']} on {(reading.get('reviewed_at') or '')[:10]}.")
        elif reading and reading.get("agreement") == "agreed":
            lines.append("Two independent readings agreed.")
        return ("\n\n— Read from the mails by AI —\n" + "\n".join(lines)) if lines else ""

    def vocabulary(self) -> None:
        for attr, label, type_, options in COMPANY_FIELDS:
            f = ensure_field(self.c, "company", label, type_, options, "mail")
            self.fields[attr] = f["column_name"]

    def companies(self) -> None:
        have = {fold(n): i for n, i in self.c.company_id_by_name().items()}
        for co in self.plan.companies:
            name = co["name"]
            try:
                cid = have.get(fold(name))
                if not cid:
                    body = {"display_name": name, "source": "import", "source_system": SOURCE_SYSTEM,
                            "domains": [{"domain": d, "is_primary": i == 0} for i, d in enumerate(co.get("domains", []))]}
                    # the Chinese registered name: Margince search finds any part of it (示例, 智能)
                    if co.get("legal_name") or co.get("legal_name_en"):
                        body["legal_name"] = co.get("legal_name") or co["legal_name_en"]
                    _, created = self.c.post_h("/v1/companies", body, f"co:{fold(name)}")
                    cid = created["id"]
                    self.report.made("company", True)
                else:
                    self.report.made("company", False)
                self.company_ids[name] = cid
                patch = {}
                if co.get("relationship_label"):
                    patch[self.fields["relationship"]] = co["relationship_label"]
                if co.get("customer_type_label"):
                    patch[self.fields["customer_type"]] = co["customer_type_label"]
                if co.get("legal_name_en"):
                    patch[self.fields["legal_name_en"]] = co["legal_name_en"]
                if co.get("business_line"):
                    patch[self.fields["business_line"]] = co["business_line"]
                legal = co.get("legal_name") or co.get("legal_name_en")
                if patch or legal:
                    current = self.c.get_company(cid)
                    patch = {k: v for k, v in patch.items() if current.get(k) != v}
                    if legal and not current.get("legal_name"):  # filled in on a re-run, never overwritten
                        patch["legal_name"] = legal
                    if patch:
                        self.c.patch(f"/v1/companies/{cid}", patch)
            except ApiError as e:
                self.report.fail("company", name, e)

    def contacts(self) -> None:
        by_email: dict[str, dict] = {}
        for ct in self.c.all_contacts():
            for e in ct.get("emails") or []:
                by_email[e["email"].lower()] = ct
        for ct in self.plan.contacts:
            try:
                found = next((by_email[e] for e in ct["emails"] if e in by_email), None)
                if found:
                    cid = found.get("id") or found.get("contact_id")
                    self.report.made("contact", False)
                    self._add_reachability(cid, ct)
                else:
                    body = {"full_name": ct["full_name"], "source": "import", "source_system": SOURCE_SYSTEM,
                            "emails": [{"email": e, "email_type": "work", "is_primary": i == 0, "position": i}
                                       for i, e in enumerate(ct["emails"])],
                            "phones": ct["phones"]}
                    if ct.get("title"):
                        body["title"] = ct["title"][:200]
                    _, created = self.c.post_h("/v1/contacts", body, f"ct:{ct['key']}")
                    cid = created["id"]
                    self.report.made("contact", True)
                    if ct.get("company") and ct["company"] in self.company_ids:
                        self.c.link_employment(cid, self.company_ids[ct["company"]])
                for e in ct["emails"]:
                    self.contact_ids[e] = cid
            except ApiError as e:
                self.report.fail("contact", ct["key"], e)

    def _add_reachability(self, cid: str, ct: dict) -> None:
        current = self.c.get_contact(cid)
        have_p = [{k: p[k] for k in ("phone", "phone_type", "is_primary", "position") if k in p}
                  for p in current.get("phones") or []]
        new_p = [p for p in ct["phones"] if p["phone"] not in {h["phone"] for h in have_p}]
        have_e = [{k: e[k] for k in ("email", "email_type", "is_primary", "position") if k in e}
                  for e in current.get("emails") or []]
        known = {h["email"].lower() for h in have_e}
        new_e = [e for e in ct["emails"] if e not in known]
        patch = {}
        if ct.get("title") and not (current.get("title") or "").strip():  # filled in, never overwritten
            patch["title"] = ct["title"][:200]
        if new_p:
            patch["phones"] = have_p + [{**p, "is_primary": not have_p and i == 0, "position": len(have_p) + i}
                                        for i, p in enumerate(new_p)]
        if new_e:
            patch["emails"] = have_e + [{"email": e, "email_type": "work", "is_primary": not have_e and i == 0,
                                         "position": len(have_e) + i} for i, e in enumerate(new_e)]
        if patch:
            self.c.patch(f"/v1/contacts/{cid}", patch)

    def products(self) -> None:
        have = self.c.all_products()
        for sku, p in self.plan.products.items():
            try:
                if sku in have:
                    self.product_ids[sku] = have[sku]["id"]
                    self.report.made("product", False)
                    continue
                api = self.ai if p.get("origin") == "reading" else self.c
                _, created = api.post_h("/v1/products", {
                    "name": p["name"], "sku": sku, "unit_price_minor": p["unit_price_minor"], "currency": p["currency"],
                    "default_tax_rate": 13 if p["currency"] == "CNY" else 0, "source": "import"}, f"pr:{sku}")
                self.product_ids[sku] = created["id"]
                self.report.made("product", True)
            except ApiError as e:
                self.report.fail("product", sku, e)

    def deals(self) -> None:
        if not self.plan.deals:
            return
        pipeline = next((p for p in self.c.pipelines() if p.get("is_default")), None) or self.c.pipelines()[0]
        stages = sorted(pipeline["stages"], key=lambda s: s["position"])
        first_open = next(s for s in stages if s["semantic"] == "open")
        won = next((s for s in stages if s["semantic"] == "won"), None)
        lost = next((s for s in stages if s["semantic"] == "lost"), None)
        have = self.c.all_deals()
        today = date.today().isoformat()
        staged = self._my_approvals() if self.agent_id else {}
        for key, d in self.plan.deals.items():
            read = d.get("origin") == "reading"
            api = self.ai if read else self.c
            try:
                existing = have.get(d["name"])
                if existing:
                    deal_id, status = existing["id"], existing.get("status")
                    self.report.made("deal", False)
                else:
                    body = {"name": d["name"], "pipeline_id": pipeline["id"], "stage_id": first_open["id"],
                            "source": "import", "source_system": SOURCE_SYSTEM}
                    # The server refuses a currency without an amount (amount_currency_pair).
                    if d.get("amount_minor"):
                        body["amount_minor"] = d["amount_minor"]
                        body["currency"] = d["currency"]
                    if d.get("company") in self.company_ids:
                        body["company_id"] = self.company_ids[d["company"]]
                    # An open deal may not claim a past close date, and a close is stamped with
                    # today, so a historical date is kept in the description instead.
                    desc = d.get("note") or ""
                    if d.get("expected_close"):
                        if d["status"] == "open" and d["expected_close"] >= today:
                            body["expected_close_date"] = d["expected_close"]
                        else:
                            desc = f"{desc}\n\nClose date per source: {d['expected_close']}".strip()
                    if read:
                        desc = (self.ai_marker + desc + self.provenance(d.get("evidence"), d.get("reading"))).strip()
                    if desc:
                        body["description"] = desc
                    _, created = api.post_h("/v1/deals", body, f"deal:{key}")
                    deal_id, status = created["id"], "open"
                    self.report.made("deal", True)
                self.deal_ids[key] = deal_id
                # Close and offer are checked on existing deals too, so the next run completes a
                # deal that an earlier run left half-written.
                target = won if d["status"] == "won" else lost if d["status"] == "lost" else None
                if target and status == "open":
                    adv = {"to_stage_id": target["id"], "status": d["status"]}
                    if d["status"] == "won":
                        adv["won_without_contract_reason"] = "imported"
                    else:  # the CRM insists on a reason to close as lost
                        adv["lost_reason"] = (d.get("note") or "Lost, as read from the mails")[:300]
                    if read and self.agent_id:
                        self._close_by_agent(deal_id, adv, staged.get(deal_id))
                    else:
                        api.post(f"/v1/deals/{deal_id}/advance", json=adv)
                if not existing or not next(self.c.paginate(f"/v1/deals/{deal_id}/offers", {"limit": 1}), None):
                    self._offer(key, d, api)
            except ApiError as e:
                self.report.fail("deal", key, e)

    def _my_approvals(self) -> dict[str, dict]:
        """Per deal, the latest won/lost approval this agent asked for (pending, approved or rejected)."""
        out: dict[str, dict] = {}
        for status in ("pending", "approved", "rejected"):
            for a in self.c.paginate("/v1/approvals", {"kind": "advance_deal", "status": status, "limit": 100}):
                if a.get("proposed_by") == f"agent:{self.agent_id}":
                    cur = out.get(a["target_entity_id"])
                    if not cur or a["created_at"] > cur["created_at"]:
                        out[a["target_entity_id"]] = a
        return out

    def _close_by_agent(self, deal_id: str, adv: dict, staged: dict | None) -> None:
        if staged and staged["status"] == "approved":
            # approved since the last run: only this agent may complete it, with the approval's id
            self.ai._request("POST", f"/v1/deals/{deal_id}/advance", json=staged["proposed_change"]["body"],
                             headers={"X-Approval-Token": staged["id"]})
            self.report.note("closes completed after your approval")
        elif staged and staged["status"] == "pending":
            self.report.note("closes waiting for your approval")
        elif staged and staged["status"] == "rejected":
            self.report.note("closes you declined")
        else:
            try:
                self.ai.post(f"/v1/deals/{deal_id}/advance", json=adv)
                self.report.note("closes applied without asking")  # only if the CRM stops gating won/lost
            except ApiError as e:
                if e.response.status_code != 403 or "approval_required" not in e.response.text:
                    raise
                self.report.note("closes waiting for your approval")

    def _offer(self, key: str, d: dict, api: MailClient | None = None) -> None:
        if not d["lines"] or not d.get("currency") or any(not ln.get("unit_price_minor") or not ln.get("qty") for ln in d["lines"]):
            return  # an offer needs a currency, a quantity and a price on every line; anything less stays in the summary
        items = []
        for i, ln in enumerate(d["lines"], 1):
            pid = self.product_ids.get(ln.get("model"))
            item = {"position": i, "quantity": ln["qty"], "unit_price_minor": ln["unit_price_minor"],
                    "description": ln.get("text") or ln.get("model")}
            if pid:
                item["product_id"] = pid
            if d["currency"] == "CNY":
                item["tax_rate"] = 13
            items.append(item)
        docs = ", ".join(x["filename"] for x in d["docs"])
        intro = f"Imported from {docs}." if docs else (self.ai_marker + "Lines as read from the mails." + self.provenance(
            [e for e in d.get("evidence") or [] if isinstance(e, dict)], d.get("reading"))).strip()
        (api or self.c).post_h(f"/v1/deals/{self.deal_ids[key]}/offers", {
            "currency": d["currency"], "line_items": items, "source": "import", "intro_text": intro}, f"offer:{key}")
        self.report.made("offer", True)

    def emails(self, limit: int | None = None) -> None:
        rows = [m for m in self.plan.emails if m["date_iso"]]
        if limit:
            rows = rows[:limit]
        for i, m in enumerate(rows, 1):
            try:
                links = self._links(m)
                payload = {
                    "kind": "email", "source": "import", "source_system": SOURCE_SYSTEM, "source_id": m["message_id"],
                    "direction": m["direction"], "subject": m["subject"] or "(no subject)",
                    "body": (m["top"] or m["body"] or "")[:BODY_LIMIT], "occurred_at": m["date_iso"],
                    "rfc_message_id": m["message_id"], "thread_key": m["thread_key"],
                    "participants": {"from": m["from"]["email"] or UNKNOWN_SENDER,
                                     "to": [p["email"] for p in m["to"]], "cc": [p["email"] for p in m["cc"]]},
                    "links": links,
                    "source_author_name": m["from"]["name"] or None,
                    "raw": {"folder": m["folder"]["path"], "folders": m["folders"], "refs": m.get("refs", {}),
                            "attachments": [{k: a[k] for k in ("filename", "size", "sha256")} for a in m["attachments"]
                                            if not a["decoration"]]},
                }
                created, act = self._log(payload)
                self.activity_ids[m["message_id"]] = act["id"]
                self.report.made("email", created)
            except ApiError as e:
                self.report.fail("email", m["message_id"], e)
            if i % 50 == 0:
                self.emit(f"  emails {i}/{len(rows)}")

    def _log(self, payload: dict, api: MailClient | None = None) -> tuple[bool, dict]:
        api = api or self.c
        try:
            return api.log_activity(payload)
        except ApiError as e:
            if e.response.status_code == 422 and payload["links"]:
                # a link target the server refuses must not lose the message itself
                return api.log_activity({**payload, "links": []})
            raise

    def routing(self) -> None:
        """Mails a reading tied to another company or to a deal, moved by the agent (the CRM keeps who did it)."""
        for r in self.plan.routing:
            aid = self.activity_ids.get(r["message_id"])
            if not aid:
                continue  # not written in this run (no date, or a trial limit)
            try:
                links = {(lk["entity_type"], lk["entity_id"]) for lk in self.c.get(f"/v1/activities/{aid}").get("links") or []}
                targets = [("company", self.company_ids[r["company"]], True)] if r.get("company") in self.company_ids else []
                targets += [("deal", self.deal_ids[k], False) for k in r.get("deals") or [] if k in self.deal_ids]
                for et, eid, move in targets:
                    if (et, eid) in links:
                        self.report.made("routing", False)
                        continue
                    self.ai._request("POST", f"/v1/activities/{aid}/relink", headers={"Idempotency-Key": f"relink:{aid}:{eid}"},
                                     json={"entity_type": et, "entity_id": eid, "replace_existing_of_type": move})
                    self.report.made("routing", True)
            except ApiError as e:
                self.report.fail("routing", r["message_id"], e)

    def titles(self) -> None:
        """Titles a reading found, filled in by the agent where the CRM has none (never overwritten)."""
        for u in self.plan.contact_updates:
            cid = self.contact_ids.get(u["email"])
            if not cid or not u.get("title"):
                continue
            try:
                if (self.c.get_contact(cid).get("title") or "").strip():
                    self.report.made("title", False)
                    continue
                answer = self.ai.patch(f"/v1/contacts/{cid}", {"title": u["title"][:200]})
                if answer.get("staged_approval"):
                    self.report.note("titles waiting for your approval")
                else:
                    self.report.made("title", True)
            except ApiError as e:
                if e.response.status_code == 403 and "approval_required" in e.response.text:
                    self.report.note("titles waiting for your approval")
                else:
                    self.report.fail("title", u["email"], e)

    def _links(self, m: dict) -> list[dict]:
        links = []
        for p in [m["from"]] + m["to"] + m["cc"]:
            cid = self.contact_ids.get(p["email"])
            if cid and not _internal(p["email"]) and {"entity_type": "contact", "entity_id": cid} not in links:
                links.append({"entity_type": "contact", "entity_id": cid})
        if m.get("company") in self.company_ids:
            links.append({"entity_type": "company", "entity_id": self.company_ids[m["company"]]})
        if m.get("deal_key") in self.deal_ids:
            links.append({"entity_type": "deal", "entity_id": self.deal_ids[m["deal_key"]]})
        return links

    def tasks_and_notes(self) -> None:
        for kind, rows in (("task", self.plan.tasks), ("note", self.plan.notes)):
            for r in rows:
                try:
                    links = []
                    if r.get("company") in self.company_ids:
                        links.append({"entity_type": "company", "entity_id": self.company_ids[r["company"]]})
                    if r.get("deal") in self.deal_ids:
                        links.append({"entity_type": "deal", "entity_id": self.deal_ids[r["deal"]]})
                    when = r.get("date") or (r.get("due") and f"{r['due']}T09:00:00+08:00")
                    if when and len(when) == 10:
                        when = f"{when}T09:00:00+08:00"
                    payload = {"kind": kind, "source": "import", "source_system": SOURCE_SYSTEM,
                               "source_id": f"{kind}:{r['id']}", "subject": self.ai_marker + r["subject"],
                               "body": ((r.get("body") or "") + self.provenance(r.get("evidence"), r.get("reading"))).strip(),
                               "occurred_at": when, "links": links,
                               "raw": {"evidence": r.get("evidence"), "reading": r.get("reading"), "type": r.get("type")}}
                    if kind == "task" and r.get("due"):
                        payload["due_at"] = f"{r['due']}T09:00:00+08:00"
                        payload["assignee_id"] = self.me
                    created, act = self._log(payload, self.ai)
                    if kind == "task" and created and r.get("done"):
                        self.ai.patch(f"/v1/activities/{act['id']}", {"is_done": True})
                    self.report.made(kind, created)
                except ApiError as e:
                    self.report.fail(kind, r["id"], e)

    def files(self) -> None:
        if self.attachments == "none":
            return
        done_entities: dict[tuple, set] = {}
        for a in self.plan.attachments:
            path = os.path.join(self.files_dir, a["stored_as"])
            targets = []
            if a["message_id"] in self.activity_ids:
                targets.append(("activity", self.activity_ids[a["message_id"]]))
            if a.get("deal_key") in self.deal_ids:
                targets.append(("deal", self.deal_ids[a["deal_key"]]))
            for et, eid in targets:
                try:
                    if (et, eid) not in done_entities:
                        done_entities[(et, eid)] = {(x["filename"], x.get("byte_size")) for x in
                                                    self.c.attachments_of(et, eid)}
                    if (a["filename"], a["size"]) in done_entities[(et, eid)]:
                        self.report.made("attachment", False)
                        continue
                    self.c.upload_attachment(et, eid, path, a["filename"])
                    done_entities[(et, eid)].add((a["filename"], a["size"]))
                    self.report.made("attachment", True)
                except (ApiError, OSError) as e:
                    self.report.fail("attachment", f"{a['filename']}@{et}", e)


def _internal(addr: str) -> bool:
    return addr.split("@")[-1] in {"fibrort.com", "fibrort.org"}


def run(client: MailClient, plan: Plan, files_dir: str, attachments: str = "documents", limit: int | None = None,
        emit=print, agent: MailClient | None = None, agent_id: str | None = None) -> Report:
    r = Run(client, plan, files_dir, attachments, emit, agent, agent_id)
    for name in ("vocabulary", "companies", "contacts", "products", "deals"):
        emit(name)
        getattr(r, name)()
    emit("emails")
    r.emails(limit)
    emit("routing")
    r.routing()
    emit("titles")
    r.titles()
    emit("tasks and notes")
    r.tasks_and_notes()
    emit("attachments")
    r.files()
    return r.report


def preview(client: MailClient, plan: Plan) -> dict:
    """What a commit would create versus what the CRM already holds. Writes nothing."""
    have_co = {fold(n) for n in client.company_id_by_name()}
    emails = {e["email"].lower() for ct in client.all_contacts() for e in (ct.get("emails") or [])}
    deals = set(client.all_deals())
    products = set(client.all_products())
    return {
        "companies_new": sum(1 for c in plan.companies if fold(c["name"]) not in have_co),
        "contacts_new": sum(1 for c in plan.contacts if not any(e in emails for e in c["emails"])),
        "products_new": sum(1 for s in plan.products if s not in products),
        "deals_new": sum(1 for d in plan.deals.values() if d["name"] not in deals),
        "emails": "idempotent on Message-ID: re-runs skip what is already there",
    }


def write_report(report: Report, path: str) -> None:
    with open(path, "w") as f:
        json.dump({"created": report.created, "existing": report.existing, "failed": report.failed,
                   "approvals": report.approvals}, f, indent=1, ensure_ascii=False)
