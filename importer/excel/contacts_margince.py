"""Margince adapter for contact lists (excel/contacts.py): the only place that knows Margince.

The accepted companies and people go through the visit pipeline's company and contact writes (a
contact is reused on a duplicate email only when the name matches). What the list knows beyond that
is filled in where the CRM has nothing and never overwritten:

  company  relationship type (customer / supplier; the group's own colleagues sit on an account of
           type "other"), lifecycle "customer" for a receivables company, address, the other-language
           name, custom fields "Contact list origin" and "Ledger type" (应收 / 应付 as written)
  contact  job title, custom fields "Contact origin", "Department", "Remark"

Contacts take no source id in Margince, so a re-run is idempotent by company + name (and the
duplicate-email rule); a second run writes nothing. Preview only reads the CRM.

  contacts_plan(result)                 -> (MatchResult, {contact key: extras})
  contacts_preview(client, match_result) -> counts of what a commit would add (reads only)
  contacts_commit(client, result)       -> counts of what was written
"""
import re

from common.fields import ensure_field
from excel.split import fold_key
from visits import pipeline
from visits.match import Company, Contact, MatchResult
from visits.reader import Parsed

from . import contacts as contact_lists

SOURCE = "contact lists"  # suffix for a custom field whose label another importer already uses
RELATIONSHIP = {"customer": "customer", "supplier": "supplier", "internal": "other"}  # CRM relationship_types by list kind
MOBILE = re.compile(r"\+861[3-9]\d{9}")
ORIGINS = sorted(set(contact_lists.ORIGINS.values()))
# attribute, label, type, options. Labels differ per object so one name never stands for two fields.
COMPANY_FIELDS = [("origin", "Contact list origin", "picklist", ORIGINS), ("ledger", "Ledger type", "text", None)]
CONTACT_FIELDS = [("origin", "Contact origin", "picklist", ORIGINS), ("department", "Department", "text", None),
                  ("remark", "Remark", "text", None)]


def contacts_plan(result) -> tuple[MatchResult, dict]:
    """The accepted companies and people of a contact-list read as the visit pipeline's MatchResult,
    plus {contact key: {title, origin, department, remark}} (empty values left out). Companies and
    contacts still in the review queue are left out."""
    match_result, extras = MatchResult(), {}
    for key, org in result.accepted_orgs().items():
        known = org.match if org.match.get("score", 0) >= contact_lists.COMPANY_OK else {}
        match_result.companies[key] = Company(
            key=key, display_name=org.name, city="", written_names=list(org.written_names),
            list_names=[n for n in (known.get("cn"), known.get("en")) if n])
    for c in result.accepted_contacts():
        mobiles = [p for p in c.phones if MOBILE.fullmatch(p)]
        match_result.contacts[c.source_id] = Contact(
            key=c.source_id, company_key=c.company_key, full_name=c.full_name, phones=mobiles,
            office_phones=[p for p in c.phones if p not in mobiles], emails=list(c.emails),
            spellings=[c.full_name])
        extra = {"title": c.role, "origin": c.origin, "department": c.department, "remark": c.remark}
        extras[c.source_id] = {k: v for k, v in extra.items() if v}
    return match_result, extras


def contacts_preview(client, match_result: MatchResult) -> dict:
    """What a commit would add, from reads only: companies and people the CRM does not hold yet.
    A person it holds under another company (found by email at commit) counts as new here."""
    by_name = {fold_key(n): i for n, i in client.company_id_by_name().items()}
    ids = {}
    for key, company in match_result.companies.items():
        hit = next((n for n in pipeline._crm_names(company) if fold_key(n) in by_name), None)
        if hit:
            ids[key] = by_name[fold_key(hit)]
    people, present = {}, 0
    for contact in match_result.contacts.values():
        cid = ids.get(contact.company_key)
        if cid is not None and cid not in people:
            people[cid] = {fold_key(c["full_name"]) for c in client.contacts_for_company(cid)}
        present += cid is not None and fold_key(contact.full_name) in people[cid]
    return {"companies_new": len(match_result.companies) - len(ids), "companies_present": len(ids),
            "contacts_new": len(match_result.contacts) - present, "contacts_present": present}


class _ContactsRun(pipeline._Run):
    """The visit pipeline's company and contact writes for people with no visit behind them."""

    def __init__(self, client, match_result, orgs, extras, opts, emit):
        super().__init__(client, Parsed(visits=[], weeks=[]), match_result, opts, emit)
        self.orgs, self.extras = orgs, extras
        self.columns = {}  # (object, attribute) -> custom field column
        self.titles_set = self.fields_set = 0

    def ensure_fields(self) -> None:
        for object_, specs in (("company", COMPANY_FIELDS), ("contact", CONTACT_FIELDS)):
            for attr, label, type_, options in specs:
                self.columns[object_, attr] = ensure_field(self.client, object_, label, type_, options, SOURCE)["column_name"]

    def enrich_company(self, company_id, company) -> int:
        org, current, patch = self.orgs[company.key], self.client.get_company(company_id), {}
        have = current.get("relationship_types") or []
        new = [t for t in dict.fromkeys(RELATIONSHIP[k] for k in org.kinds if k in RELATIONSHIP) if t not in have]
        if new:
            patch["relationship_types"] = [*have, *new]
        if "customer" in org.kinds and current.get("lifecycle") in (None, "unknown"):
            patch["lifecycle"] = "customer"
        address = current.get("address") or {}
        where = org.address or org.location  # a supplier row with no person: its remark names the place ("Shanghai")
        if where and not address.get("line1"):
            patch["address"] = {**address, "line1": where}
        other = pipeline.other_language_name(current.get("display_name") or "", company.written_names)
        if other and not current.get("legal_name"):
            patch["legal_name"] = other
        for attr, value in (("origin", org.origin), ("ledger", ", ".join(org.ledgers))):
            column = self.columns["company", attr]
            if value and not current.get(column):
                patch[column] = value
        if not patch:
            return 0
        self.client.patch(f"/v1/companies/{company_id}", patch)
        return 1

    def import_contacts(self) -> None:
        super().import_contacts()
        for key, extra in self.extras.items():
            contact_id = self.contact_ids.get(key)
            if not contact_id:
                continue
            current, patch = self.client.get_contact(contact_id), {}
            if extra.get("title") and not (current.get("title") or "").strip():
                patch["title"] = extra["title"][:200]
            for attr in ("origin", "department", "remark"):
                column = self.columns["contact", attr]
                if extra.get(attr) and not current.get(column):
                    patch[column] = extra[attr]
            if patch:
                self.client.patch(f"/v1/contacts/{contact_id}", patch)
                self.titles_set += "title" in patch
                self.fields_set += len(patch) - ("title" in patch)


def contacts_commit(client, result, source_system="contactlist", on_progress=None) -> dict:
    """Writes the accepted companies and people of a contact-list read; a second run writes nothing."""
    match_result, extras = contacts_plan(result)
    emit = on_progress or (lambda step, **info: None)
    run = _ContactsRun(client, match_result, result.accepted_orgs(), extras,
                       pipeline.Options(source_system=source_system), emit)
    run.ensure_fields()
    run.import_companies()
    run.import_contacts()
    r = run.report
    return {"companies": r.companies_report["disposition"], "companies_enriched": r.companies_updated,
            "contacts_created": r.contacts_created, "contacts_updated": r.contacts_updated,
            "titles_set": run.titles_set, "fields_set": run.fields_set,
            "contacts_without_company": len(match_result.contacts) - len(run.contact_ids)}
