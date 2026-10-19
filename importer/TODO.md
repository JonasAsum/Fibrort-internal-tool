# Importer to-do

- [x] Match contact details to customers: the name, phone and email from the folder 04 ERP lists are read and attached to their
  account by customer code (done 16 Oct 2026, `tests/quality/decisions/accounts.md` D-ACC-14).
- [ ] Check fuzzy company matching on real data at the last step (owner's request): how often a visit's company is matched to the
  wrong company, or asked about although it is obvious. The matcher is `excel/companies.py`; the visit side is `excel/run.py:best_match`.
- [ ] A first import into the real CRM starts with the preview ("Test import (dry run)") and the 17 decision cases in `docs/decisions-queue.md`.
