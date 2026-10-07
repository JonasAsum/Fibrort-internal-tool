"""Layer 1d: which companies we actually do business with. Decided by counting, never by guessing.

Evidence is collected per organisation domain (per address for free-mail senders). These messages give none:
  mass mail  - 8 or more outside domains on To/Cc (the "former employee" apology went to 30 firms)
  automated  - List-Unsubscribe / List-Id / Precedence: bulk headers, or a noreply/newsletter/... sender
  bills      - inbound mail whose subject only announces an invoice/statement (travelsite.com, didi, telecom)
Keep rules, the first that fires is recorded as the reason:
  filed        mails filed in this company's own Customer/Purchase folder
  purchase     a sender in the Purchase / Order Confirmation folders (supplier)
  logistics    a sender in the Logistics folder (service provider)
  two-way      we wrote to them and they wrote to us, one-to-one
  we wrote     we wrote to them about business (quote, order, contract, product)
  contract     a party in a parsed contract / order confirmation
  inquiry      they asked us about our products and nobody answered in this archive (lead)
  quoted       sender of a forwarded business mail (Alpha, Beta inside the Acme warranty case)
Everything else is dropped, with the reason. staging/decisions.json overrides any row:
  {"travelsite.com": {"keep": false}, "fairgrounds.de": {"keep": true, "relationship": "Service provider"},
   "gamma.com": {"name": "Gamma", "legal_name": "...", "legal_name_en": "...", "parent": "..."}}
"""
import csv
import re
from collections import Counter, defaultdict

from .people import FREEMAIL, is_group, is_internal, org_domain

MASS_DOMAINS = 8
AUTOMATED_SENDER = re.compile(r"^(no[-_.]?reply|do[-_.]?not[-_.]?reply|notifications?|newsletter|news|marketing|"
                              r"mailer|bounce|system|alert|edm|promotion|e-?mail)\b", re.I)
BILL = re.compile(r"发票|对账|账单|结算|报销|invoice|statement|receipt|fapiao|dunning|mahnung", re.I)
PRODUCT = re.compile(r"转台|分割器|工作台|rotary|rundtisch|FIBROTOR|FIBROPLAN|FIBROTAKT|VR\.?\s?NC|\bER\.?\s?\d{2}|"
                     r"\bEM\.?\s?\d{2}|\btables?\b|spare\s?parts?|备件|配件", re.I)
INQUIRY = re.compile(r"询价|报价|RFQ|inquiry|enquiry|quotation|\bquote\b|angebot|repair|维修|保养|maintenance|"
                     r"service request", re.I)
DEAL = re.compile(r"\bcontract\b|合同|order confirmation|auftragsbest|\bSO\s?\d{6,}|\bOC\s?\d{6,}|\bQU\d{6,}|"
                  r"\bPO\s?\d{6,}|订单|payment|付款|delivery|交货|发货", re.I)
OUR_PO = re.compile(r"(?<![A-Za-z0-9])PO\s?2\d{6}(?!\d)")  # FIBRO Shanghai's own purchase order numbers: PO2603011
ORDER_REF = re.compile(r"(?<![A-Za-z0-9])(?:PO|SO|OC|QU)\s?\d{6,}", re.I)  # po0000000001: Example grinding's order to us
SERVICE = re.compile(r"shipment|shipping|订舱|pre-?alert|报关|customs|air export|sea freight|freight|货运|logistic|"
                     r"pick ?up|ticket|audit|审计|\btax\b|税|exhibition|展会|展位|booth|registration|symposium|"
                     r"consult|咨询|translation|balloon", re.I)
RELATIONSHIPS = ["Customer", "Supplier", "Service provider", "Lead", "Third party", "Partner", "FIBRO group"]
TYPE_TO_REL = {"dealer": "Customer", "end_user": "Customer", "manufacturer": "Customer", "customer": "Customer",
               "supplier": "Supplier"}


def _ext(addr: str) -> bool:
    d = addr.split("@")[-1]
    return bool(addr) and not is_internal(addr) and not is_group(d)


def key_of(addr: str) -> str | None:
    if not addr or not _ext(addr):
        return None
    d = addr.split("@")[-1].lower()
    return addr.lower() if d in FREEMAIL else org_domain(d)


OUR_NAMES = re.compile(r"fibro\W*(rotary\W*tables?|rundtische?|precision)\w*|rotary\W*tables?\W*(us|sea|inc|gmbh)\b",
                       re.I)


def _text(m: dict) -> str:
    """Subject + document names, with our own company names removed ('FIBRO Rotary Tables US Inc' on an audit
    form is not a product)."""
    t = m["subject"] + " " + " ".join(a["filename"] for a in m["attachments"] if not a["decoration"])
    return OUR_NAMES.sub(" ", t)


def decide(messages: list[dict], companies: dict[str, dict], overrides: dict | None = None) -> dict[str, dict]:
    """Returns {key: row}; row = {key, keep, relationship, reason, evidence, subjects}."""
    overrides = overrides or {}
    ev: dict[str, Counter] = defaultdict(Counter)
    subjects: dict[str, list] = defaultdict(list)
    folder_type: dict[str, Counter] = defaultdict(Counter)
    accepted = {d: co for co in companies.values() if co.get("folder_name") for d in co["domains"]}
    party_keys = {k for k, co in companies.items() if co.get("parties")}

    for m in messages:
        sender = m["from"]["email"]
        outbound = m["direction"] == "outbound" or (sender and not _ext(sender))
        rcpt_domains = {p["email"].split("@")[-1] for p in m["to"] + m["cc"]
                        if p["email"] and _ext(p["email"])}
        mass = len(rcpt_domains) >= MASS_DOMAINS
        automated = m.get("bulk") or bool(sender and AUTOMATED_SENDER.match(sender.split("@")[0]))
        text = _text(m)
        product, inquiry, deal = bool(PRODUCT.search(text)), bool(INQUIRY.search(text)), bool(DEAL.search(text))
        bill_only = bool(BILL.search(text)) and not (product or inquiry)
        business = product or inquiry or deal
        our_po = bool(OUR_PO.search(text))
        folder = m["folder"]
        own_group_folder = bool(re.match(r"(?i)fibro|laepple", folder.get("company") or ""))
        keys = ({key_of(p["email"]) for p in m["to"] + m["cc"]} if outbound else {key_of(sender)}) - {None}
        for k in keys:
            e = ev[k]
            if len(subjects[k]) < 3 and m["subject"] not in subjects[k]:
                subjects[k].append(m["subject"][:70])
            own_folder = folder.get("company") and accepted.get(k) and accepted[k]["folder_name"] == folder["company"]
            if own_folder:
                e["filed"] += 1
                folder_type[k][folder["customer_type"]] += 1
            elif folder.get("company"):
                e["filed_other"] += 1
            if mass:
                e["mass"] += 1
                continue
            if automated and not outbound:
                e["automated"] += 1
                continue
            if outbound:
                e["out"] += 1
                # one-way "we wrote about business" only counts for mail the mailbox owner wrote herself; a
                # colleague's product mailing with her on CC is his relationship, not evidence of ours
                e["out_business"] += business and m["direction"] == "outbound"
                e["we_buy"] += our_po
            else:
                if bill_only:
                    e["bill"] += 1
                    continue
                e["in"] += 1
                e["in_inquiry"] += product or inquiry
                e["in_order"] += bool(ORDER_REF.search(text)) and not our_po
                # the Purchase folders also hold telecom bills and shipping notes: only mail about an order,
                # a price or our PO number makes a supplier
                if (folder.get("topic") in ("Purchase", "Oder Confirmation") or (
                        folder["customer_type"] == "supplier" and not folder.get("topic"))) and (
                        deal or our_po or product or re.search(r"price|preis|surcharge|价格|涨价|报价|credit term|account application|开户|账期", text, re.I)):
                    e["purchase_folder"] += 1
                if folder.get("topic") == "Logistics":
                    e["logistics_folder"] += 1
            e["product"] += product or inquiry
            e["our_product"] += product
            e["service"] += bool(SERVICE.search(text))
            e["deal"] += deal
        # forwarded / quoted mails: only the SENDER of a quoted block, only in a business mail
        # a forwarded payment or admin mail (bank, accountant) is not a business relation; one about our
        # products, an inquiry, or anything in a customer's folder is
        if mass or (automated and not outbound) or not (
                product or inquiry or (folder.get("company") and not own_group_folder)):
            continue
        for b in m.get("quoted", []):
            qdomains = {a.split("@")[-1] for a in b["from"] + b["to"] + b["cc"] if _ext(a)}
            if len(qdomains) >= MASS_DOMAINS:
                continue
            # the quoted sender, and on a small targeted mail (<= 3 outside domains) also its recipients:
            # Alpha wrote, Beta (the end user with the broken table) was on CC
            quoted_addrs = b["from"] + (b["to"] + b["cc"] if len(qdomains) <= 3 else [])
            for k in {key_of(a) for a in quoted_addrs} - {None} - keys:
                ev[k]["quoted"] += 1
                ev[k]["quoted_inquiry"] += (product or inquiry) and key_of(b["from"][0]) == k
                if len(subjects[k]) < 3 and m["subject"] not in subjects[k]:
                    subjects[k].append(m["subject"][:70])

    rows: dict[str, dict] = {}
    for k in set(ev) | {c for c in companies if c.startswith("party:")} | set(party_keys):
        e = ev[k]
        two_way = e["out"] > 0 and e["in"] > 0
        if any(is_group(d) for d in [k.split("@")[-1]]):
            keep, reason, rel = True, "FIBRO group domain", "FIBRO group"
        elif e["filed"]:
            t = folder_type[k].most_common(1)[0][0]
            keep, reason, rel = True, f"filed in its own folder ({e['filed']} mails)", TYPE_TO_REL.get(t, "Customer")
        elif e["purchase_folder"]:
            keep, reason, rel = True, f"sender in Purchase folders ({e['purchase_folder']})", "Supplier"
        elif e["logistics_folder"]:
            keep, reason, rel = True, f"sender in Logistics folder ({e['logistics_folder']})", "Service provider"
        elif two_way:
            keep, reason, rel = True, f"two-way conversation ({e['out']} out / {e['in']} in)", None
        elif e["out_business"]:
            keep, reason, rel = True, f"we wrote to them about business ({e['out_business']})", None
        elif k in party_keys or k.startswith("party:"):
            keep, reason, rel = True, "party in a contract / order document", None
        elif e["in_inquiry"]:  # an inquiry quoting a project/PO number ('询价PO44016724') is still an inquiry
            keep, reason, rel = True, f"inbound product inquiry, unanswered here ({e['in_inquiry']})", "Lead"
        elif e["in_order"]:
            keep, reason, rel = True, f"sent us mail about an order number ({e['in_order']})", "Customer"
        elif e["quoted_inquiry"]:
            keep, reason, rel = True, f"asked about our products in a forwarded mail ({e['quoted_inquiry']})", "Lead"
        elif e["quoted"]:
            keep, reason, rel = True, f"in a forwarded business mail ({e['quoted']})", "Third party"
        else:
            keep, rel = False, None
            if e["mass"] and not (e["in"] or e["out"]):
                reason = f"only in mass mailings ({e['mass']})"
            elif e["automated"] and not (e["in"] or e["out"]):
                reason = f"newsletter / automated only ({e['automated']})"
            elif e["bill"] and not e["out"]:
                reason = f"only sends bills, never answered ({e['bill']})"
            elif e["out"] and not e["in"]:
                reason = "we wrote to them, not about business, no answer"
            else:
                reason = "inbound only, no business or product content"
        if keep and rel is None:
            if e["we_buy"]:
                rel = "Supplier"
            elif e["our_product"] and (e["out"] or e["deal"] or k in party_keys or k.startswith("party:")):
                rel = "Customer"
            elif e["service"]:
                rel = "Service provider"
            elif k in party_keys or k.startswith("party:"):
                rel = "Customer"
            elif e["filed_other"] or e["quoted"]:
                rel = "Third party"
            else:
                rel = "Partner"
        rows[k] = {"key": k, "keep": keep, "relationship": rel, "reason": reason, "evidence": dict(e),
                   "subjects": subjects[k], "overridden": False}

    for k, o in overrides.items():
        if not isinstance(o, dict) or o.get("same_as") or k.startswith("_"):
            continue
        r = rows.setdefault(k, {"key": k, "keep": False, "relationship": None, "reason": "", "evidence": {},
                                "subjects": [], "overridden": False})
        if "keep" in o:
            r["keep"] = bool(o["keep"])
        if o.get("relationship"):
            r["relationship"] = o["relationship"]
        r["reason"] = (r["reason"] + "; " if r["reason"] else "") + "overridden in decisions.json"
        r["overridden"] = True
    return rows


def write_csv(rows: dict[str, dict], companies: dict[str, dict], path: str) -> None:
    cols = ["decision", "relationship", "key", "display_name", "legal_name_zh", "legal_name_en", "name_conflicts",
            "reason", "out", "in", "filed", "quoted", "mass", "automated", "bill", "subjects"]
    with open(path, "w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig: Excel shows the Chinese
        w = csv.writer(f)
        w.writerow(cols)
        for r in sorted(rows.values(), key=lambda r: (not r["keep"], r["relationship"] or "", r["key"])):
            co = companies.get(r["key"], {})
            conf = co.get("name_conflicts") or {}
            e = r["evidence"]
            w.writerow(["KEEP" if r["keep"] else "DROP", r["relationship"] or "", r["key"], co.get("name") or "",
                        co.get("legal_name") or "", co.get("legal_name_en") or "",
                        " | ".join(conf.get("zh", []) + conf.get("en", [])), r["reason"],
                        e.get("out", 0), e.get("in", 0), e.get("filed", 0), e.get("quoted", 0), e.get("mass", 0),
                        e.get("automated", 0), e.get("bill", 0), " || ".join(r["subjects"])])
