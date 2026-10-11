import unittest

from visits.match import Contact
from visits.pipeline import contact_emails


class Emails(unittest.TestCase):
    def test_only_ascii_addresses_count(self):
        c = Contact(key="k", company_key="c", full_name="X", emails=["名字@test.example", "a.b@test.example", "x@例え.example"])
        self.assertEqual([e["email"] for e in contact_emails(c)], ["a.b@test.example"])


# ---- idempotency across files and re-exports ----
import os
from datetime import datetime

from excel.run import read_paths
from tests.excel_book import WEEKLY_HEADER, TmpCase, write_book
from tests.fake_crm import FakeCRM
from visits import pipeline
from visits.match import match


def weekly_file(path, sheets):
    """sheets: {name: [(date, customer, next step)]}: visits of one rep."""
    out = {}
    for name, rows in sheets.items():
        cells = dict(WEEKLY_HEADER)
        cells.update({"E4": "Owner：", "F4": "Jason Zhao"})
        for i, (day, customer, step) in enumerate(rows):
            cells.update({f"B{9 + i}": "Jason Zhao", f"D{9 + i}": day, f"F{9 + i}": customer, f"P{9 + i}": step})
        out[name] = {"cells": cells}
    return write_book(path, out)


def load(crm, *paths):
    parsed = read_paths(list(paths)).parsed
    return pipeline.run(crm, parsed, match(parsed.visits)), parsed


MON, TUE, NEXT_MON = datetime(2026, 5, 18), datetime(2026, 5, 19), datetime(2026, 5, 25)


class LaterFileRepeatsEarlierWeeks(TmpCase):
    def files(self):
        a = weekly_file(self.path("Template for Weekly Sales Reporting 2026_CW21_Rep.xlsx"),
                        {"CW21": [(MON, "Acme Test Co.", "Send offer"), (TUE, "Beta Test GmbH", "")]})
        b = weekly_file(self.path("Template for Weekly Sales Reporting 2026_CW22_Rep.xlsx"),
                        {"CW21": [(MON, "Acme Test Co.", "Send offer"), (TUE, "Beta Test GmbH", "")],
                         "CW22": [(NEXT_MON, "Gamma Test Ltd.", "Call back")]})
        return a, b

    def test_next_weeks_file_only_adds_the_new_visits(self):
        crm = FakeCRM()
        a, b = self.files()
        first, _ = load(crm, a)
        self.assertEqual((first.activities_created, first.tasks_created), (2, 1))
        second, _ = load(crm, b)
        self.assertEqual((second.activities_created, second.activities_idempotent), (1, 2))
        self.assertEqual((second.tasks_created, second.tasks_idempotent), (1, 1))
        self.assertEqual(len(crm.activities), 2 + 1 + 1 + 1)  # 3 visits, 2 tasks
        again, _ = load(crm, b)
        self.assertEqual((again.activities_created, again.tasks_created), (0, 0))

    def test_same_visit_twice_on_one_day_stays_two_records(self):
        crm = FakeCRM()
        a = weekly_file(self.path("Template for Weekly Sales Reporting 2026_CW21_Rep.xlsx"),
                        {"CW21": [(MON, "Acme Test Co.", ""), (MON, "Acme Test Co.", "")]})
        report, _ = load(crm, a)
        self.assertEqual(report.activities_created, 2)
        self.assertEqual(len(crm.activities), 2)


class OldKeysAreRecognised(TmpCase):
    def seed_old_style(self, crm, parsed):
        """Writes what the previous version wrote: ids built from file, sheet and row."""
        result = match(parsed.visits)
        run = pipeline._Run(crm, parsed, result, pipeline.Options(), lambda *a, **k: None)
        run.ensure_vocabulary()
        run.import_companies()
        run.import_contacts()
        for v in parsed.visits:
            ck, kk = result.links[v.visit_key]
            payload = run.visit_payload(v, result.companies.get(ck), run.company_ids.get(ck), run.contact_ids.get(kk))
            crm.log_activity({**payload, "source_id": v.visit_key})
            for n in pipeline.task_notes(v):
                t = run.task_payload(v, n, result.companies.get(ck), run.company_ids.get(ck), None)
                crm.log_activity({**t, "source_id": f"{v.visit_key}#K{n.excel_row}"})
        for w in parsed.weeks:
            crm.log_activity({"source": "import", "source_system": "visitlist", "source_id": w.week_key,
                              "kind": "note", "subject": f"Weekly report CW{w.cw} {w.year} — {w.owner}",
                              "raw": {"visit_list_week": {"owner": w.owner, "year": w.year, "cw": w.cw}}})

    def test_reload_after_the_key_change_creates_nothing_new(self):
        a = weekly_file(self.path("Template for Weekly Sales Reporting 2026_CW21_Rep.xlsx"),
                        {"CW21": [(MON, "Acme Test Co.", "Send offer"), (TUE, "Beta Test GmbH", "")]})
        parsed = read_paths([a]).parsed
        crm = FakeCRM()
        self.seed_old_style(crm, parsed)
        before = dict(crm.activities)
        self.assertEqual(crm.count("task"), 1)
        report, _ = load(crm, a)
        self.assertEqual((report.activities_created, report.tasks_created, report.weeks_created), (0, 0, 0))
        self.assertEqual(crm.activities.keys(), before.keys())

    def test_a_later_file_repeating_old_style_records_creates_nothing_for_them(self):
        a = weekly_file(self.path("Template for Weekly Sales Reporting 2026_CW21_Rep.xlsx"),
                        {"CW21": [(MON, "Acme Test Co.", "Send offer"), (TUE, "Beta Test GmbH", "")]})
        b = weekly_file(self.path("Template for Weekly Sales Reporting 2026_CW22_Rep.xlsx"),
                        {"CW21": [(MON, "Acme Test Co.", "Send offer"), (TUE, "Beta Test GmbH", "")],
                         "CW22": [(NEXT_MON, "Gamma Test Ltd.", "")]})
        crm = FakeCRM()
        self.seed_old_style(crm, read_paths([a]).parsed)
        n = len(crm.activities)
        report, _ = load(crm, b)
        self.assertEqual((report.activities_created, report.tasks_created), (1, 0))
        self.assertEqual(len(crm.activities), n + 1)
