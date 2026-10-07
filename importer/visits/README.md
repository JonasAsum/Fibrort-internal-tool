# visits

Imports the weekly Excel visit lists (layout documented in
`mock_visits/README.md`) into a running Margince CRM, keeping every cell:
customers become companies with their potential and type, the people on
each visit become contacts with their phones and emails, every visit row
becomes a meeting on its date, every action line becomes a task with a due
date, and every week's key plan and summary becomes a note. The table below
says where each column lands.

## Setup

One `.venv` in `importer/` serves both importers and the web UI; see
`importer/README.md`. Every command below runs from `importer/`.

## The easy way: web UI

```sh
cd importer
.venv/bin/python -m webui.server
```
then open **http://localhost:5055**. Drop a workbook on the page, review
the companies/contacts it found, click "Import to CRM", confirm what's
about to be written, done. Connection settings (server URL, login) are
behind the gear icon in the top right — it connects with the local dev
admin by default, so you only need to touch that if you're pointing it at
a different installation.

State (the uploaded files, the matched companies/contacts) lives in the
server process while it's running, and `.webui_config.json` /
`.uploads/` next to `webui/server.py` persist between restarts — both are
local, throwaway files, not something to commit anywhere.

## The scriptable way: CLI + Postman

```sh
# 1. Parse only - no network call, sanity-checks the file(s) and prints counts.
.venv/bin/python -m visits.cli --mode parse "../mock_visits/excel/*.xlsx"

# 2. Preview - logs in, previews the company CSV import and counts the
#    contacts, visits, tasks and week notes a commit would write. Writes nothing.
.venv/bin/python -m visits.cli --mode preview "../mock_visits/excel/*.xlsx"

# 3. Commit - the real run: approves both CSV imports and posts one
#    activity per visit, idempotently (safe to re-run on the same file).
.venv/bin/python -m visits.cli --mode commit "/path/to/the/real/workbooks/*.xlsx"
```

Point it at one workbook, several, or a glob - everything is consolidated
into one set of companies/contacts/activities before anything is written,
so a customer visited by two owners still lands as one company.

Credentials default to the seeded local-dev admin from
`margince/docs/tutorials/getting-started.md` (`admin@demo.test`, only valid
on a `make dev && make seed-dev` install). Point `--base-url`,
`--email`/`--password` (or `MARGINCE_BASE_URL`/`MARGINCE_EMAIL`/
`MARGINCE_PASSWORD`) at a real installation before running `--mode commit`
against one.

## Calling the endpoints by hand, in Postman

`postman/postman_collection.json` + `postman/postman_environment.json` walk through the
same calls `common/client.py` makes, one request per step, so you can inspect
each response instead of trusting the script: login, upload the companies
CSV, preview the import run, read its report and approve it, then add a
company field, create a contact with phones and emails and link it to its
employer, and log one visit, one task and one week plan. Import both files into
Postman, select the "Margince local dev" environment, and run the "1.
Auth / Login" request first — its Tests script pulls the session cookie
out of the response and stores it as a collection variable that every
other request sends back as a literal `Cookie` header (`crm_session` is
marked `Secure`, so Postman's own cookie jar may otherwise refuse to
resend it over plain `http://localhost`, exactly like `requests.Session`
did until `client.py` worked around it the same way).

For the CSV-upload request, Postman can't carry a file inside an exported
collection — pick the file by hand in the `file` form-data row. Generate it
from your workbook(s) first:

```sh
.venv/bin/python -m visits.cli --dump-csv visits/sample_csv "/path/to/workbooks/*.xlsx"
```

Each folder is numbered in the order to run it in. Everything from the
"Approve" request onwards writes.

## Where every cell of the sheet ends up

| Sheet | In Margince |
|---|---|
| Customer, Location/City | company (`display_name`, `address.city`) |
| Potential (A, B, C) | company field **Potential** (newest value by visit date); every older rating stays on the visit that recorded it |
| Customer (new, exist, dealer) | company field **Customer type** (new / existing / dealer) and `lifecycle` (existing → customer, new → prospect) when the CRM still says unknown; the cell as written stays in the visit |
| Owner (the rep) | author of every record ("Liu Yang via visitlist"); meeting host, task assignee and company owner when the rep has a seat with the same name; company field **Sales reps** |
| Name, Phone, Office, Email | contact with every distinct number (mobile + office, normalized to +86…) and email, employed at the company; numbers a re-run finds are added, never replaced |
| Date | `occurred_at` on the visit's day in `--timezone` (default Asia/Shanghai). The sheet has no clock time: visits on one day are spaced 09:00, 11:00, 13:00 … in sheet order and the body says "time of day not recorded" |
| Purpose, all note lines, deadline column | the visit's body, with `[action]` / `[waiting]` for the green/yellow fills and the resolved due date next to the deadline as written |
| Green/yellow note lines, deadlines | one **task** each, due on the date the phrase names ("in 3 days", "this week", "this month", "CW 26", "in October"), assigned to whoever "X will follow up" names. Tasks due before the newest week in the file are written already done (`--close-tasks-due-before`) |
| Key Plan (Mon-Fri) | a "Week plan: planned for Thu 31 Jul 2025" line on each visit to a company the plan names (plus "visited …" when the visit happened on another day) |
| Key Plan + Summary | one **weekly report** note per sheet, dated Friday 17:00 of that week and linked to no company, so it never shows on a company's page as if it were about that company |
| The whole row | the activity's `raw` (`visit_list_row` / `visit_list_week`), so nothing the sheet said is lost even where no field fits |
| First/last visit | company fields **First visit**, **Last visit** |

Every write carries `source_system=visitlist` and a stable `source_id`
(`workbook:sheet:row`, `…#K<row>` for a task, `…:week` for a week), so a
second run of the same files writes nothing.

## How it works

1. `reader.py` turns one workbook into raw visit rows (date in 4 spellings,
   CW in 4 sheet-name spellings, every column, the note lines with their
   fill colour and Excel row) and one week record per sheet (key plan per
   weekday with its date, summary lines).
2. `match.py` folds the spellings a file uses for one company or person into
   one record. A name written out in full is compared word for word, so
   `Hengrui Intelligent Equipment` and `Hengrui Intelligent Manufacturing`
   stay two companies. A shortened form (`Hengrui Auto.`, `HENGRUI`,
   `Hengrui Co.`) is tied to a full name only on evidence from its own row:
   the same contact met at exactly one candidate in that city, or exactly one
   candidate in that city. Otherwise it becomes its own company and the review
   screen lists it, because a duplicate is one "Merge company" click, while two
   customers' visits under one name is a silent error. Contacts group by
   company plus surname, longest surname first (`Lin` is not `Li`).
3. `when.py` turns dates and deadline phrases into instants.
4. `pipeline.py` writes, in order: the five company fields (created once,
   which needs an admin login); companies the CRM does not already hold
   exactly, through the normal CSV import with `on_duplicate=create`, then
   their fields; contacts with phones and emails plus their employment link;
   one meeting per visit (a note when no contact was met, because a meeting
   cannot link a company directly); one task per action line; and one note per
   week.

Why `on_duplicate=create`: the CSV importer's duplicate check is a fuzzy name
ladder meant to queue pairs for a human. On Chinese company names, which
share "Intelligent Equipment", "Tech" and "Automation", it reported 61 of the
91 mock companies as duplicates of the others and skipped them. Exact names
already in the CRM are filtered out beforehand, so a re-run creates no twins,
and the near-misses still go to the CRM's duplicate review queue.

## Measured on the mock files

Against `mock_visits/ground_truth`, on a fresh `make dev` + `make seed-dev`
workspace, reading every record back through the API:

- 309/309 visits written with the right date, author, purpose, city, contact,
  phone, office, email, potential and type; all 763 note lines and 45
  deadlines appear in the body; cancelled visits show as `canceled`.
- 91 companies for 91 real ones. One visit is filed under the wrong company
  (a bare `Huaxin` in Guangzhou, where two Huaxin companies both are) and one
  real company is split in two (`Yuxin Co.` / `Yuxin Automation`).
- 126 contacts for 125 real people, none merged wrongly; every ground-truth
  phone and email is on the right contact.
- 155 tasks (112 with a due date, 126 closed as historical); the 15
  "Cancelled by customer." lines are outcomes, not tasks.
- 56/56 weekly reports with every plan entry and summary line.
- A second run writes nothing: 0 created, everything "already present".

## Known limitations

- Matching is evidence-based, not perfect: see the numbers above. The review
  screen and `--mode parse` list every short name that was kept separate,
  which is the list to check for "Merge company" afterwards.
- Reps are matched to seats by display name. Without a seat they are still
  the named author of every record, but the meeting host, task assignee and
  company owner fall back to whoever ran the import.
- The meaning of the green/yellow fills (action / waiting) is still a guess
  from the printed sample.
- Due dates come from a handful of phrase rules. A phrase they do not know
  stays text on the task and gets no date rather than a guessed one.
- Every visit needs a parseable date; a row whose date cell matches none of
  the four known styles is skipped and named at the end of the run.
