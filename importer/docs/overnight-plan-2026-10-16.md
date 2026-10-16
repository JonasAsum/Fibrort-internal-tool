# Overnight plan: Excel importer, delivery-ready (started 2026-10-16)

Owner's brief: one clear version A where the owner only decides special single cases (each with my best
recommended fix already applied as the default); full end-to-end tests; an executive-level visual
(graph, workflow) whose focus is WHY the data is correct and trustworthy. Reliability is priority 1.
Owner is asleep: all permissions granted, work autonomously, report in the morning.

## Decisions already made by the owner (apply, do not re-ask)
- Companies that are not in the ERP lists but appear in action files are imported as new companies.
- A contact can have several phone numbers (the CRM holds many).
- Phone numbers without a country code are Chinese: 11 digits starting 13-19 are mobiles (+86), numbers
  starting 0 are landlines (+86, drop the 0). `+49`, `+86` etc. kept. Helper: `excel/split.py:phone_e164`.
- People with the same short name at different companies are separate entries (company + name is the key).
- Fuzzy company matching quality is checked at the latest step.
- Contact details (name/phone/email) in the folder 04 ERP lists must be matched to their customers (TODO.md).
- Sampling is not acceptable: every cell of every file is checked.

## Quality system (the thing that makes it trustworthy), layers
L0 ledger (every cell, independent) | L1 invariants on every record | L2 two independent AI labellers on
distinct cell texts, slices must be literal substrings | L3 record snapshot + diff on every change |
L4 synthetic adversarial workbooks | L5 import into the in-memory CRM and compare to the ledger |
mutation checks (inject faults, tests must fail).
Code: `importer/tests/quality/` (`ledger.py`, `snapshot.py`, `invariants.py`, `run_visits.py`,
`test_quality_visits.py`), snapshots in `importer/tests/golden_records/` (git-ignored, real customer data).

## State (final, 16 Oct 2026 morning)
- [x] visits (folders 01-03): 1,267 records, 0 rule violations (167 before); 398 of 404 distinct customer cells agree with both AI readers, 4 open cases (V-05/06/07/13), 2 harmless.
- [x] accounts 04-05, contacts 06, todos 07, orders 08, receivables 09: oracle + rules + snapshot + fixes (accounts/contacts/orders bugs fixed, see decisions/*.md).
- [x] phones: one shared rule (`excel/split.py:phone_e164`, used by `visits/pipeline.py:normalize_phone` and the accounts/contacts mapping); several numbers per contact.
- [x] end to end: all 59 files into the in-memory CRM, every phone/e-mail/company traced to an Excel cell; second import writes nothing.
- [x] decisions queue (`docs/decisions-queue.md`), trust report (`docs/trust-report.html`, artifact https://claude.ai/artifact/NfnsYdttTdDxC79eX2XB5p).
- [x] tests: full suite 1,080 passed + the one timing test passes alone; 97 browser tests pass (`pytest -m e2e tests/e2e`), including importing twice into the local test CRM.
- [ ] fuzzy company matching quality on real data (owner: last step), first import into the real CRM starting with the dry run.
- Nothing is committed to git: no commit was requested. About 90 files are modified or new in the working tree.

## Known facts to keep
- Real sample folder: `/Users/jonas/Mac Documents/Shanghai Company/CRM Excel Lists 2026-10-11` (folders 01-09, 59 workbooks).
- Machine is slow when loaded; full suite takes 6-10 min; read all visit files takes ~1 min.
- `scratchpad/gt/` holds the old sampled ground truth (select.py there shadows the stdlib: never put it on sys.path).
- Old golden count tests (`tests/test_all_files_matrix.py`) were regenerated on 16 Oct 2026 after reviewing the diff (more accounts and contacts written, fewer company questions, second import writes nothing).
- Published artifacts: map https://claude.ai/artifact/NkYWjZNL28wkKq5ToPosz7, progress board https://claude.ai/artifact/LNRsJa7kmVm56S1EXBrheQ
