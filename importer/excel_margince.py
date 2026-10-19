"""Margince adapter for the excel/ reader: the only place that knows Margince.

excel/ turns workbooks into plain visit and week records; this module hands the
auto-accepted ones to the existing visit pipeline (visits.match + visits.pipeline),
which writes companies, contacts, meetings, tasks and week notes idempotently
(source_system + source_id = workbook:sheet:row, so a re-run writes nothing new).
Visits still waiting in the review queue are left out until they are resolved.

  preview - log in, preview the company CSV import, count what a commit would write.
  commit  - write. Only ever point it at a local test CRM unless told otherwise.

--family contacts reads contact lists (excel/contacts.py) instead of visit lists. The people and
their companies go through the same company and contact writes as the visits (a contact is reused
on a duplicate email only when the name matches); preview there only reads the CRM. Contacts take
no source id in Margince, so a re-run is idempotent by company + name (and the duplicate-email rule).

Usage (from importer/):
  .venv/bin/python -m excel_margince --mode preview "<glob>" ...
  .venv/bin/python -m excel_margince --family contacts --mode preview "<contact lists glob>"
"""
import argparse
import os

from common.client import MargenceClient
from excel import contacts as contact_lists
from excel.contacts_margince import contacts_commit, contacts_plan, contacts_preview  # noqa: F401  (re-exported)
from excel.cli import _expand
from excel.run import read_paths
from visits import pipeline
from visits.match import match
from visits.reader import Parsed


def accepted(result) -> Parsed:
    """Only the auto-accepted visits, and the weeks they belong to."""
    visits = [v for v in result.parsed.visits if v.visit_key in result.auto_accepted]
    for v in visits:
        # A customer the master list knows is filed under the master's name, so every
        # spelling of it lands on one company; the sheet's spelling stays in raw.
        m = v.extra.get("company_match") or {}
        if m.get("score", 0) >= 0.9 and (m.get("cn") or m.get("en")):
            v.extra.setdefault("customer_as_written", v.customer_raw)
            v.customer_raw = m.get("cn") or m.get("en")
    sheets = {(v.workbook, v.sheet) for v in visits}
    return Parsed(visits=visits, weeks=[w for w in result.parsed.weeks if (w.workbook, w.sheet) in sheets])


def _main_contacts(args):
    result = contact_lists.read_paths(_expand(args.excel))
    s = contact_lists.summarize(result)
    print(f"{s['contacts_accepted']} of {s['contacts']} contacts and {s['companies_accepted']} of "
          f"{s['companies']} companies ready ({s['companies_queued']} companies wait for an answer: "
          f"{s['contacts_queued']} contacts); {s['rows']} rows read from {s['files']} files, "
          f"{s['blank_rows']} blank, skipped: {s['skipped_by_reason'] or 'none'}")
    for f in result.skipped_files:
        print(f"file skipped: {os.path.basename(f['file'])}: {f['error']}")
    match_result, extras = contacts_plan(result)
    client = MargenceClient(args.base_url)
    client.login(args.email, args.password)
    if args.mode == "preview":
        p = contacts_preview(client, match_result)
        print(f"would write: {p['companies_new']} companies ({p['companies_present']} already there), "
              f"{p['contacts_new']} contacts ({p['contacts_present']} already there), {sum('title' in e for e in extras.values())} job titles")
        return
    r = contacts_commit(client, result, args.source_system)
    print("companies:", r["companies"], f"- {r['companies_enriched']} enriched")
    print(f"contacts: {r['contacts_created']} created, {r['contacts_updated']} given new phones/emails; "
          f"{r['titles_set']} job titles and {r['fields_set']} other fields set")
    if r["contacts_without_company"]:
        print(f"contacts without a company in the CRM: {r['contacts_without_company']}")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("excel", nargs="+")
    ap.add_argument("--mode", choices=["preview", "commit"], default="preview")
    ap.add_argument("--base-url", default=os.environ.get("MARGINCE_BASE_URL", "http://localhost:8080"))
    ap.add_argument("--email", default=os.environ.get("MARGINCE_EMAIL", "admin@demo.test"))
    ap.add_argument("--password", default=os.environ.get("MARGINCE_PASSWORD", "demo-password-123"))
    ap.add_argument("--family", choices=["visits", "contacts"], default="visits")
    ap.add_argument("--source-system")
    args = ap.parse_args(argv)
    args.source_system = args.source_system or ("contactlist" if args.family == "contacts" else "visitlist")
    if args.family == "contacts":
        return _main_contacts(args)

    result = read_paths(_expand(args.excel))
    parsed = accepted(result)
    matched = match(parsed.visits)
    counts = pipeline.plan_counts(parsed, matched)
    print(f"{len(parsed.visits)} auto-accepted of {len(result.parsed.visits)} visits "
          f"({len(result.review)} review items stay out); {len(matched.companies)} companies, "
          f"{counts['contacts']} contacts, {counts['tasks']} action lines, {counts['weeks']} weeks")

    client = MargenceClient(args.base_url)
    client.login(args.email, args.password)
    if args.mode == "preview":
        report = pipeline.preview_companies(client, matched)
        print(f"companies: {report['rows_read']} new rows -> {report['disposition']}")
        print(f"would write: {counts['meetings']} meetings, {counts['notes']} visit notes, "
              f"{counts['tasks']} tasks, {counts['weeks']} weekly reports")
        return
    r = pipeline.run(client, parsed, matched, pipeline.Options(source_system=args.source_system))
    print("companies:", r.companies_report["disposition"], f"- {r.companies_updated} enriched")
    print(f"contacts: {r.contacts_created} created, {r.contacts_updated} given new phones/emails")
    print(f"visits: {r.activities_created} created, {r.activities_idempotent} already present")
    print(f"tasks: {r.tasks_created} created ({r.tasks_closed} already done), {r.tasks_idempotent} already present")
    print(f"weekly reports: {r.weeks_created} created, {r.weeks_idempotent} already present")
    if r.activities_skipped_no_date:
        print(f"skipped (no parseable date): {len(r.activities_skipped_no_date)}")


if __name__ == "__main__":
    main()
