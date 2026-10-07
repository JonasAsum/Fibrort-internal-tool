"""Entry point. Same three modes as visits/cli.py, cheapest first:

  parse    - read staging/extracted.json (+ enrichment.json), print what a commit would write. No network.
  preview  - log in and compare with what the CRM already holds. Writes nothing.
  commit   - companies, contacts (phones, emails), products, deals + offers, one email activity per
             message, tasks/notes from enrichment.json, then attachments. Safe to re-run.

Run `python -m mail.extract` first to produce staging/extracted.json.

Credentials: --email/--password or MARGINCE_EMAIL/MARGINCE_PASSWORD; defaults are the seeded local-dev
admin (`make dev` + `make seed-dev`). Production = the same command with its --base-url and login.
"""
import argparse
import json
import os
import sys

from . import pipeline
from .plan import build_plan, load_enrichment
from common.passport import agent_client

HERE = os.path.dirname(os.path.abspath(__file__))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--mode", choices=["parse", "preview", "commit"], default="parse")
    ap.add_argument("--staging", default=os.path.join(HERE, "staging"))
    ap.add_argument("--base-url", default=os.environ.get("MARGINCE_BASE_URL", "http://localhost:8080"))
    ap.add_argument("--email", default=os.environ.get("MARGINCE_EMAIL", "admin@demo.test"))
    ap.add_argument("--password", default=os.environ.get("MARGINCE_PASSWORD", "demo-password-123"))
    ap.add_argument("--attachments", choices=["documents", "none"], default="documents")
    ap.add_argument("--only-company", help="restrict the run to one company (e.g. Acme), for a trial")
    ap.add_argument("--limit-emails", type=int, help="write at most this many emails (trial run)")
    ap.add_argument("--agent", action="store_true", help="write AI-read records under an agent passport (the CRM stamps them AI, but caps it at 200 writes per 24 h)")
    args = ap.parse_args()

    with open(os.path.join(args.staging, "extracted.json")) as f:
        extracted = json.load(f)
    enrichment, reading = load_enrichment(args.staging)
    plan = build_plan(extracted, enrichment, args.only_company)
    plan.reading = reading
    counts = plan.counts()
    if reading:
        counts["reading"] = reading
    print(json.dumps(counts, indent=1))
    for w in plan.warnings:
        print("WARNING:", w, file=sys.stderr)
    if args.mode == "parse":
        return

    client = pipeline.MailClient(args.base_url)
    client.login(args.email, args.password)
    if args.mode == "preview":
        print(json.dumps(pipeline.preview(client, plan), indent=1))
        return

    agent = agent_id = None
    read = plan.tasks or plan.notes or plan.routing or plan.contact_updates or any(
        d.get("origin") == "reading" for d in plan.deals.values())
    if read and args.agent:  # default: the human login with an "[AI-read]" marker. The agent passport the CRM
        # stamps as AI is capped at 200 writes per 24 h, far below what a full reading writes.
        agent, agent_id = agent_client(client, cls=pipeline.MailClient)
        print("AI agent passport:", agent_id)
    report = pipeline.run(client, plan, os.path.join(args.staging, "files"), args.attachments, args.limit_emails,
                          agent=agent, agent_id=agent_id)
    pipeline.write_report(report, os.path.join(args.staging, "report.json"))
    print("created:", report.created)
    print("already there:", report.existing)
    if report.approvals:
        print("approvals (CRM inbox):", report.approvals)
    print("failed:", len(report.failed), "(details in staging/report.json)")


if __name__ == "__main__":
    main()
