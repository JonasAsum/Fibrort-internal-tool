"""Margince adapter for the excel/ reader: the only place that knows Margince.

excel/ turns workbooks into plain visit and week records; this module hands the
auto-accepted ones to the existing visit pipeline (visits.match + visits.pipeline),
which writes companies, contacts, meetings, tasks and week notes idempotently
(source_system + source_id = workbook:sheet:row, so a re-run writes nothing new).
Visits still waiting in the review queue are left out until they are resolved.

  preview - log in, preview the company CSV import, count what a commit would write.
  commit  - write. Only ever point it at a local test CRM unless told otherwise.

Usage (from importer/):
  .venv/bin/python -m excel_margince --mode preview "<glob>" ...
"""
import argparse
import os

from common.client import MargenceClient
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


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("excel", nargs="+")
    ap.add_argument("--mode", choices=["preview", "commit"], default="preview")
    ap.add_argument("--base-url", default=os.environ.get("MARGINCE_BASE_URL", "http://localhost:8080"))
    ap.add_argument("--email", default=os.environ.get("MARGINCE_EMAIL", "admin@demo.test"))
    ap.add_argument("--password", default=os.environ.get("MARGINCE_PASSWORD", "demo-password-123"))
    ap.add_argument("--source-system", default="visitlist")
    args = ap.parse_args(argv)

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
