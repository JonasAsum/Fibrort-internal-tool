# importer

Gets data into the local Margince CRM (`../margince`). One site, one venv, one API client:

```
importer/
  webui/      the site, http://localhost:5055: Excel lists, Email archive, Reading review, Checks
    entry.py      the front door: tells what each uploaded workbook is, by its content
    families.py   one adapter per kind of list: read, counts, questions, flags, dry run, commit
    checks.py, selftest.py   the Checks page
  excel/      one reader per kind of Excel list (plain records first) and its Margince adapter (*_margince.py)
  visits/     the visit pipeline (match, companies, contacts, meetings, tasks, notes) behind folders 01-03
  mail/       Thunderbird email archive -> companies, contacts, deals, emails, attachments
    staging/  extracted.json, enrichment.json, decisions.json, relevance.csv, files/
  common/     client.py: login and REST calls to Margince, used by all of it
```

## Setup (once)

```sh
cd importer
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
```

## Run

Every command runs from `importer/`. The CRM must be up (`cd ../margince && make dev && make seed-dev`).

```sh
.venv/bin/python -m webui.server                          # the site, http://localhost:5055

.venv/bin/python -m visits.cli --mode parse "../mock_visits/excel/*.xlsx"   # parse | preview | commit
.venv/bin/python -m excel_margince --family accounts --mode preview "<glob>"  # also contacts | todos | orders | receivables

.venv/bin/python -m mail.extract                          # archive -> mail/staging/extracted.json
.venv/bin/python -m mail.checks                           # re-check extracted.json
.venv/bin/python -m mail.cli --mode parse                 # parse | preview | commit
.venv/bin/python -m mail.verify                           # fact check (see below)

.venv/bin/python -m mail.cases                            # case files for the reading layer (see below)
.venv/bin/python -m mail.reading --company Acme         # check + merge the two readings of a case
.venv/bin/python -m mail.reading --status                 # which cases are read, merged, waiting
```

Details: `visits/README.md`, `excel/SPEC.md` and `mail/HANDOFF.md`.

## The site

Open the site, press the gear for the CRM address and login (saved in `webui/.webui_config.json`, never in git).

**Excel lists** (Upload -> Review -> Import) is the front door for every spreadsheet. Drop one file, many files or a
whole folder (also `Choose a folder`; only the `.xlsx` and `.xls` in it count). Each file is recognised by what is
inside it, not by its name or folder, and the review starts with a table that says, per file, what it is, where it goes
in the CRM, how many rows it holds and what will happen:

| Folder of the CRM review | List | Goes into the CRM as |
|---|---|---|
| 01-03 | weekly sales reports, weekly planning, customer visits in detail | meetings, notes, contacts, tasks, weekly notes |
| 04, 05 | ERP customer lists, customer lists, registrations | accounts; competitor, forecast, application on the account; registrations as leads |
| 06 | contact lists (Chinese finance lists, FIBRO handover) | contacts on their accounts, tagged by origin |
| 07 | to-do and sales topics | tasks |
| 08 | ERP sales order lists | read-only order history notes on the account |
| 09 | AR and financial status | open and overdue receivables as a number on the account |

A file the folder review says does not belong in the CRM (invoice and packing lists, the `Actual_mit_Formeln` workbook,
bank cash flows, payables, fixed costs, Läpple 2023, stock lists) is shown as *recognised, kept out of the CRM on
purpose* with the review's reason. It is never an error. A file nobody knows says so and why. A blank template says it
holds nothing.

Below the table there is one tab per kind of list in the upload: counts, notes, the "same company?" questions, and
the **flagged** groups the owner asked to be able to filter: the JET-IT list, the topics template, finance-list
contacts and internal FIBRO colleagues. They are imported with a visible tag unless you untick them (one click).

**Test import (dry run)** shows, per kind of record, exactly what an import would write now, and writes nothing.
**Import to CRM** shows the same sheet as a last look, then writes in a fixed order: accounts (04/05), contacts (06),
visits (01-03), tasks (07), orders (08), receivables (09). One list failing never blocks the others; the result page
lists, per list, what was written, was already there, was skipped or failed. Importing the same files again says
*Nothing new* and writes nothing (ids are made from the content).

**Checks** (header) lists PASS / WARN / FAIL rows, each failing one with how to fix it: CRM reachable, login with the
saved settings, Python modules, writable folders, company master, saved answers, the custom fields the receivable,
order and opportunity data need, and a built-in self-test that builds a tiny workbook per kind of list and reads it
with the real reader. Checking never writes to the CRM.

API behind the page: `POST /api/upload` (any number of workbooks), `GET /api/batch`, `POST /api/company_answer`,
`/api/company_pair_answer`, `/api/flag`, `/api/dry_run`, `/api/commit` (+ `GET /api/commit/<job>`), `GET /api/checks`.

## Tests

```sh
.venv/bin/python -m pytest -q                      # about 6 minutes; skips what needs the real sample folder
.venv/bin/python -m pytest -q -m e2e               # browser tests (Playwright) against throw-away servers
REGEN_MATRIX=1 .venv/bin/python -m pytest -q tests/test_all_files_matrix.py   # after a deliberate change, then read the diff
```

`tests/test_all_files_matrix.py` walks the real sample folder (`EXCEL_SAMPLE_DIR`, default
`../../CRM Excel Lists 2026-10-11`) and gives every workbook in folders 01-09 its own test case: uploaded through
`/api/upload` like the page does, dry-run, committed into an in-memory CRM (`tests/fake_site_crm.py`) and committed
again. The expectations (counts only) are in `tests/golden_files_matrix.json`; a new file without a row fails.
Browser tests use `python -m tests.e2e.site_server`: the real site with that in-memory CRM, on a spare port, with its
own upload and settings folders. They never touch port 5055 or the real CRM; only `tests/e2e/test_import_local_crm.py`
talks to a CRM, a local one, and only when it is up.

## Quality checks: how we know the Excel import is right

Everything the readers produce is compared with the Excel files themselves, not with the readers (`tests/quality/`):

| Check | What it does | Run it |
|---|---|---|
| Cell ledger | Every non-empty cell of every workbook, read with openpyxl/xlrd only (`ledger.py`). | used by the others |
| Rules on every record | Per family (`invariants.py`, `test_quality_<family>.py`): every cell is found in the right field or dropped by a declared rule; numbers match to the cent; a phone field holds a phone. | `pytest tests/quality` |
| Two AI readers | Two independent labellers split every distinct customer cell; every word must be copied from the cell; the importer is compared with both (`l2_compare.py`, `test_l2_customers.py`). After a reader change run `python -m tests.quality.l2_refresh`. | `pytest tests/quality/test_l2_customers.py` |
| Snapshots | Every field of every record, before and after a change (`snapshot*.py`, `tests/golden_records/`, git-ignored: real customer data). | `python -m tests.quality.snapshot diff` |
| Into a test CRM and back | All 59 files imported into the in-memory CRM; every phone, e-mail and company in it traced to an Excel cell (`test_end_to_end.py`). | `pytest tests/quality/test_end_to_end.py` |
| Test the tests | Old bugs are put back on purpose; the checks must fail (`test_quality_mutations.py`, `test_the_checks_catch_the_old_bugs`). | in `pytest tests/quality` |

The decisions the owner can change, each with the recommended answer already applied: `docs/decisions-queue.md`
(detail: `tests/quality/decisions/`). Hand-checked rows that were overridden by a decision: `excel/gold/overrides.json`.

## Reading layer: what only understanding the mails can give

Deals, tasks, case notes, purchase notes, mails filed under the wrong company, and what scanned PDFs say are
read by Claude, under rules that code enforces (`mail/READING.md`):

1. `python -m mail.cases` writes one case file per company (`mail/staging/reading/<company>/case.md`). It holds
   every mail with an alias (M1, M2 …), the attachments (D1 …), and the known people, companies and numbers.
2. Two **independent readings** per case, `pass_a.json` and `pass_b.json`; the second reader never sees the
   first. Every fact carries a word-for-word quote from a specific mail or PDF.
3. `python -m mail.reading --company X` checks every quote against the original mail (the fact check's own
   decoder) and every number against its quote, then merges both readings into
   `mail/staging/enrichment/<company>.json`:
   - **auto**: both agree, and nothing needs a person.
   - **needs review**: deal amounts, anything read from a scan, and every disagreement.
   - **rejected**: a quote not found, a name not in the case. These never go in.
4. **Reading review tab** in the web UI: both readings side by side, with the quoted sentence highlighted in
   its mail and a link to open each scan. Accept A / Accept B / Edit / Reject. Decisions are kept in the
   company file and survive re-merging.
5. The import writes the read facts under the **AI agent's passport** (`importer/.agent_passport.json`,
   readable only by you, revocable in the CRM):
   - the CRM marks them as created by AI;
   - each record ends with the quotes it was read from and who checked it;
   - a deal closed as won or lost waits in the CRM approval inbox; once approved, the next import completes it.

Keep `mail/staging/enrichment/`: like `decisions.json`, it holds review decisions, which can't be rebuilt.

## Fact check (email import)

`python -m mail.verify` (or **Fact check** in the web UI's Email archive tab) checks two things and writes
nothing:

1. **Original archive → import:** it re-reads the Thunderbird mailboxes with its own decoder and looks up
   every value the import uses. That covers names, emails, phones, titles, legal names, domains, order numbers,
   amounts, offer lines, which company and deal each mail belongs to, and the attachments' bytes. Each value
   must be found where it should come from, e.g. a phone number in a mail that person sent.
2. **Import → CRM:** it reads every company, contact, deal, email, task, note and attachment back from
   the CRM and compares them field by field. It reports anything missing, changed, extra or doubled.

Every value gets one status:
- **VERIFIED**: found where it should be.
- **WEAK**: found, but only somewhere else.
- **MANUAL**: read by hand from a scan; only its evidence note can be checked.
- **NOT FOUND**, **MISSING**, **CHANGED**, **EXTRA**, **DUPLICATE**: problems. The exit code is then 1.

Output:
- `mail/staging/verify_report.csv`: one row per item that isn't VERIFIED; opens in Excel.
- `verify_report.json`: every item.

The first run decodes the 1.7 GB archive in about a minute. Later runs take seconds, using
`mail/staging/.verify_cache.pickle`, which is rebuilt when a mailbox file changes. `--source` or `--crm` runs
only one half.

## Clear database (web UI, gear icon → Danger zone)

Stops the dev stack, rebuilds the `margince` database with `make dev-fresh`, runs `make seed-dev`,
and re-adds the local CNY→EUR rate that won CNY deals need. This takes a few minutes. The full
output is in `webui/.uploads/reset_db.log`.

**AI-read records:** written under the normal login with an `[AI-read]` marker and quote lines. The agent passport (`--agent`) is capped by the CRM at 200 writes per 24 h, so it is off by default. See `mail/HANDOFF.md` Update 7.
