"""Entry point. Three modes, cheapest first:

  parse    - read the workbook(s), run the customer/contact matching, print
             counts and the judgement calls matching made. No network call.
  preview  - also log in and run the company CSV import's preview step
             (uploads, previews, reads the report) and count everything else
             a commit would write. Never approves, never writes a record.
  commit   - the real thing: companies, their fields, contacts with phones
             and emails, one activity per visit, one task per action line and
             one note per week.

Credentials: --email/--password, or MARGINCE_EMAIL/MARGINCE_PASSWORD. Falls
back to the seeded local-dev admin (admin@demo.test) documented in
margince/docs/tutorials/getting-started.md, which only exists on a `make
dev` + `make seed-dev` install - never use that fallback against anything
but your own localhost.
"""
import argparse
import glob
import os
import sys
from collections import Counter

from . import pipeline
from common.client import MargenceClient
from .match import match
from .reader import parse_paths


def _expand(patterns: list[str]) -> list[str]:
    paths = []
    for p in patterns:
        matches = sorted(glob.glob(p))
        paths.extend(matches if matches else [p])
    missing = [p for p in paths if not os.path.isfile(p)]
    if missing:
        sys.exit(f"not a file: {missing[0]}")
    return paths


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("excel", nargs="+", help="workbook path(s) or glob(s), e.g. 'excel/*.xlsx'")
    ap.add_argument("--mode", choices=["parse", "preview", "commit"], default="parse")
    ap.add_argument("--base-url", default=os.environ.get("MARGINCE_BASE_URL", "http://localhost:8080"))
    ap.add_argument("--email", default=os.environ.get("MARGINCE_EMAIL", "admin@demo.test"))
    ap.add_argument("--password", default=os.environ.get("MARGINCE_PASSWORD", "demo-password-123"))
    ap.add_argument("--source-system", default="visitlist",
                    help="non-reserved source_system stamped on every record for idempotent re-runs")
    ap.add_argument("--timezone", default="Asia/Shanghai", help="where the visits happened (IANA name)")
    ap.add_argument("--no-tasks", action="store_true", help="keep action lines in the visit body only")
    ap.add_argument("--close-tasks-due-before", metavar="YYYY-MM-DD",
                    help="tasks due before this day are written as done "
                         "(default: Monday of the newest week in the files)")
    ap.add_argument("--dump-csv", metavar="DIR",
                    help="write companies.csv here (e.g. to upload by hand in Postman) and exit")
    args = ap.parse_args()

    parsed = parse_paths(_expand(args.excel))
    result = match(parsed.visits)
    counts = pipeline.plan_counts(parsed, result)
    print(f"{counts['visits']} visit rows, {len(parsed.weeks)} weeks, {len(result.companies)} companies, "
          f"{counts['contacts']} contacts ({counts['contacts_with_phone']} with a phone, "
          f"{counts['contacts_with_email']} with an email), {counts['tasks']} action lines")
    no_date = sum(1 for v in parsed.visits if not v.date_iso)
    if no_date:
        print(f"WARNING: {no_date} rows have an unparsed date and will be skipped on commit", file=sys.stderr)

    if args.dump_csv:
        os.makedirs(args.dump_csv, exist_ok=True)
        path = os.path.join(args.dump_csv, "companies.csv")
        with open(path, "w") as f:
            f.write(pipeline.companies_csv(result))
        print("wrote", path)
        return

    if args.mode == "parse":
        decisions = result._namer.decisions
        print("short company names resolved:", dict(Counter(d["why"] for d in decisions)))
        for d in decisions:
            if d["why"].startswith("kept separate"):
                print(f"  kept separate: {d['written']!r} in {d['city'] or '?'} (candidates: {d['candidates']})")
        return

    client = MargenceClient(args.base_url)
    client.login(args.email, args.password)

    if args.mode == "preview":
        report = pipeline.preview_companies(client, result)
        print(f"companies: {report['rows_read']} new rows -> {report['disposition']}")
        if report["issues"]:
            print("  near-duplicates the CRM will queue for review:", len(report["issues"]))
        print(f"would write: {counts['meetings']} meetings, {counts['notes']} visit notes, "
              f"{counts['tasks'] if not args.no_tasks else 0} tasks, {counts['weeks']} weekly reports")
        return

    opts = pipeline.Options(source_system=args.source_system, timezone=args.timezone, tasks=not args.no_tasks,
                            close_tasks_due_before=args.close_tasks_due_before)
    r = pipeline.run(client, parsed, result, opts)
    print("custom fields created:", r.fields_created or "none (already there)")
    print("reps with a seat:", list(r.reps_with_seat) or "none", "| without:", r.reps_without_seat or "none")
    print("companies:", r.companies_report["disposition"], f"- {r.companies_updated} enriched")
    print(f"contacts: {r.contacts_created} created, {r.contacts_updated} given new phones/emails")
    print(f"visits: {r.activities_created} created, {r.activities_idempotent} already present")
    print(f"tasks: {r.tasks_created} created ({r.tasks_closed} already done), {r.tasks_idempotent} already present")
    print(f"weekly reports: {r.weeks_created} created, {r.weeks_idempotent} already present")
    if r.activities_skipped_no_date:
        print(f"skipped (no parseable date): {len(r.activities_skipped_no_date)}")


if __name__ == "__main__":
    main()
