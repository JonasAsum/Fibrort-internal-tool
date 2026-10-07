# importer

Gets data into the local Margince CRM (`../margince`). Two importers share one venv, one
API client and one web UI:

```
importer/
  webui/      web UI for both importers: http://localhost:5055
  visits/     weekly Excel visit lists  -> companies, contacts, meetings, tasks, notes
  mail/       Thunderbird email archive -> companies, contacts, deals, emails, attachments
    staging/  extracted.json, enrichment.json, decisions.json, relevance.csv, files/
  common/     client.py: login and REST calls to Margince, used by both
```

## Setup (once)

```sh
cd importer
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
```

## Run

Every command runs from `importer/`. The CRM must be up (`cd ../margince && make dev && make seed-dev`).

```sh
.venv/bin/python -m webui.server                          # web UI, http://localhost:5055

.venv/bin/python -m visits.cli --mode parse "../mock_visits/excel/*.xlsx"   # parse | preview | commit

.venv/bin/python -m mail.extract                          # archive -> mail/staging/extracted.json
.venv/bin/python -m mail.checks                           # re-check extracted.json
.venv/bin/python -m mail.cli --mode parse                 # parse | preview | commit
.venv/bin/python -m mail.verify                           # fact check (see below)

.venv/bin/python -m mail.cases                            # case files for the reading layer (see below)
.venv/bin/python -m mail.reading --company Acme         # check + merge the two readings of a case
.venv/bin/python -m mail.reading --status                 # which cases are read, merged, waiting
```

Details: `visits/README.md` and `mail/HANDOFF.md`.

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
