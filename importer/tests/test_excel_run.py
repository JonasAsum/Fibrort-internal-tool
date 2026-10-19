import os
import shutil
import unittest
from datetime import datetime

from excel.run import read_paths
from tests.excel_book import WEEKLY_HEADER, TmpCase, write_book


def week_file(path, cw, extra_customer=None):
    cells = dict(WEEKLY_HEADER)
    cells.update({"E4": "Owner：", "F4": "Jason Zhao"})
    cells.update({"B9": "Jason Zhao", "D9": datetime(2026, 5, 18), "F9": "Acme Test Co.", "P9": "Call"})
    if extra_customer:
        cells.update({"B10": "Jason Zhao", "D10": datetime(2026, 5, 19), "F10": extra_customer})
    return write_book(path, {f"CW{cw}": {"cells": cells}})


class ReadPaths(TmpCase):
    def test_corrupt_file_is_skipped_and_run_continues(self):
        good = week_file(self.path("Template for Weekly Sales Reporting 2026_CW21_Test Rep.xlsx"), 21)
        bad = self.path("Template for Weekly Sales Reporting 2026_CW22_Broken.xlsx")
        with open(bad, "wb") as f:
            f.write(b"not an xlsx")
        result = read_paths([good, bad])
        self.assertEqual([v.customer_raw for v in result.parsed.visits], ["Acme Test Co."])
        self.assertEqual([os.path.basename(s["file"]) for s in result.skipped_files],
                         [os.path.basename(bad)])

    def test_unchanged_file_is_not_parsed_again_and_a_changed_one_is(self):
        from unittest import mock
        path = week_file(self.path("Template for Weekly Sales Reporting 2026_CW21_Cache Rep.xlsx"), 21)
        with mock.patch("excel.run.load", wraps=__import__("excel.run").run.load) as spy:
            first = read_paths([path])
            second = read_paths([path])
            self.assertEqual(spy.call_count, 1)
            self.assertEqual([v.customer_raw for v in second.parsed.visits], ["Acme Test Co."])
            second.parsed.visits[0].customer_raw = "changed by a caller"  # the cache hands out copies
            self.assertEqual(read_paths([path]).parsed.visits[0].customer_raw, "Acme Test Co.")
            week_file(path, 21, extra_customer="Beta Test GmbH")
            self.assertEqual(len(read_paths([path]).parsed.visits), 2)
            self.assertEqual(spy.call_count, 2)
        self.assertEqual(len(first.parsed.visits), 1)

    def test_identical_bytes_are_read_once(self):
        first = week_file(self.path("a_CW21.xlsx"), 21)
        second = self.path("copy_CW21.xlsx")
        shutil.copy(first, second)
        result = read_paths([first, second])
        self.assertEqual(len(result.parsed.visits), 1)

    def test_newest_copy_wins_and_older_is_stale(self):
        old = week_file(self.path("Template_CW20_Test Rep.xlsx"), 20)
        new = week_file(self.path("Template_CW21_Test Rep.xlsx"), 21, extra_customer="Beta Test GmbH")
        result = read_paths([old, new])
        self.assertEqual(sorted(v.customer_raw for v in result.parsed.visits),
                         ["Acme Test Co.", "Beta Test GmbH"])
        stats = {os.path.basename(s.path): s for s in result.per_file_stats}
        self.assertTrue(stats["Template_CW20_Test Rep.xlsx"].stale_copy)
        self.assertFalse(stats["Template_CW21_Test Rep.xlsx"].stale_copy)


class RunFixes(TmpCase):
    def test_empty_customer_visit_is_kept_and_queued(self):
        cells = dict(WEEKLY_HEADER)
        cells.update({"E4": "Owner：", "F4": "Jason Zhao", "B9": "Jason Zhao",
                      "D9": datetime(2026, 5, 18), "P9": "Call"})
        path = write_book(self.path("Template for Weekly Sales Reporting 2026_CW21_Test Rep.xlsx"),
                          {"CW21": {"cells": cells}})
        result = read_paths([path])
        (v,) = result.parsed.visits
        self.assertEqual(v.customer_raw, "")
        self.assertIn("no_customer", {r.reason for r in result.review})

    def test_internal_rows_skip_company_review(self):
        cells = dict(WEEKLY_HEADER)
        cells.update({"E4": "Owner：", "F4": "Jason Zhao",
                      "B9": "Jason Zhao", "D9": datetime(2026, 5, 18), "F9": "Public Holiday",
                      "B10": "Jason Zhao", "D10": datetime(2026, 5, 19), "F10": "Team Building"})
        path = write_book(self.path("Template for Weekly Sales Reporting 2026_CW21_Test Rep.xlsx"),
                          {"CW21": {"cells": cells}})
        result = read_paths([path])
        self.assertEqual([v.extra["kind"] for v in result.parsed.visits], ["internal", "internal"])
        self.assertNotIn("company", {r.reason for r in result.review})

    def test_company_master_loaded_once_per_process(self):
        from excel import run
        run._master.cache_clear()
        first = run._master(run.IMPORTER_DIR)
        self.assertIs(run._master(run.IMPORTER_DIR), first)
        self.assertEqual(run._master.cache_info().misses, 1)


class SheetLimits(TmpCase):
    def test_sheet_beyond_the_row_cap_is_reported_not_silently_cut(self):
        from excel import sheet as sheet_mod
        self.assertEqual(sheet_mod.MAX_ROW, 20000)
        cells = dict(WEEKLY_HEADER)
        cells.update({"E4": "Owner：", "F4": "Jason Zhao"})
        for i in range(9, 40):
            cells.update({f"B{i}": "Jason Zhao", f"D{i}": datetime(2026, 5, 18), f"F{i}": f"Cust {i} Test Co."})
        path = write_book(self.path("Template for Weekly Sales Reporting 2026_CW21_Test Rep.xlsx"),
                          {"CW21": {"cells": cells}})
        old = sheet_mod.MAX_ROW
        sheet_mod.MAX_ROW = 20
        self.addCleanup(setattr, sheet_mod, "MAX_ROW", old)
        result = read_paths([path])
        (stat,) = result.per_file_stats
        self.assertTrue(any(name == "CW21" and "20" in why and "rows" in why for name, why in stat.sheet_errors),
                        stat.sheet_errors)


def one_visit_book(path, day, sheet="CW21", customer="Acme Test Co."):
    cells = dict(WEEKLY_HEADER)
    cells.update({"E4": "Owner：", "F4": "Jason Zhao", "B9": "Jason Zhao", "D9": day, "F9": customer, "P9": "Call"})
    return write_book(path, {sheet: {"cells": cells}})


def reasons(result):
    return {r for rs in result.review_by_visit.values() for r in rs}


class VisitYear(TmpCase):
    def at(self, path, iso):
        t = datetime.fromisoformat(iso).timestamp()
        os.utime(path, (t, t))

    def test_date_in_another_year_than_the_file_name_is_blocking(self):
        p = one_visit_book(self.path("Template for Weekly Sales Reporting 2026_CW21_Rep.xlsx"), datetime(2025, 5, 19))
        result = read_paths([p])
        self.assertIn("week_mismatch", reasons(result))
        self.assertEqual(result.auto_accepted, set())

    def test_matching_year_is_still_accepted(self):
        p = one_visit_book(self.path("Template for Weekly Sales Reporting 2026_CW21_Rep.xlsx"), datetime(2026, 5, 18))
        result = read_paths([p])
        self.assertNotIn("week_mismatch", reasons(result))
        self.assertEqual(len(result.auto_accepted), 1)

    def test_no_year_in_name_uses_sheet_cw_and_mtime_not_today(self):
        p = one_visit_book(self.path("Weekly CW21 Rep.xlsx"), "18/05")
        self.at(p, "2024-05-25")
        (v,) = read_paths([p]).parsed.visits
        self.assertEqual(v.date_iso, "2024-05-18")

    def test_no_year_in_name_and_date_of_another_year_than_mtime_is_blocking(self):
        p = one_visit_book(self.path("Weekly CW21 Rep.xlsx"), datetime(2026, 5, 18))
        self.at(p, "2025-05-25")
        self.assertIn("week_mismatch", reasons(read_paths([p])))

    def test_year_unknowable_is_blocking(self):
        p = one_visit_book(self.path("Weekly Rep.xlsx"), datetime(2026, 5, 18), sheet="Plan")
        result = read_paths([p])
        self.assertIn("no_year", reasons(result))
        self.assertEqual(result.auto_accepted, set())


def book_with(path, rows, sheet="CW21"):
    """rows: [(date, customer)] as visits of Jason Zhao on one sheet."""
    cells = dict(WEEKLY_HEADER)
    cells.update({"E4": "Owner：", "F4": "Jason Zhao"})
    for i, (day, customer) in enumerate(rows):
        cells.update({f"B{9 + i}": "Jason Zhao", f"D{9 + i}": day, f"F{9 + i}": customer, f"P{9 + i}": "Call"})
    return write_book(path, {sheet: {"cells": cells}})


MON = datetime(2026, 5, 18)


class Dedupe(TmpCase):
    def test_renamed_stale_copy_never_displaces_the_genuine_visit(self):
        real = book_with(self.path("Template for Weekly Sales Reporting 2026_CW21_Rep.xlsx"), [(MON, "Acme Test Co.")])
        stale = book_with(self.path("Template for Weekly Sales Reporting 2026_CW23_Rep.xlsx"),
                          [(MON, "Acme Test Co."), (datetime(2026, 5, 19), "Beta Test GmbH")], sheet="CW21")
        result = read_paths([real, stale])
        acme = next(v for v in result.parsed.visits if v.customer_raw == "Acme Test Co.")
        self.assertEqual(acme.workbook, os.path.basename(real))
        self.assertIn(acme.visit_key, result.auto_accepted)
        self.assertEqual([x["ignored"] for x in result.superseded], [os.path.basename(stale)])

    def test_renamed_copy_with_a_wrong_sheet_name_loses_too(self):
        real = book_with(self.path("Template for Weekly Sales Reporting 2026_CW21_Rep.xlsx"), [(MON, "Acme Test Co.")])
        stale = book_with(self.path("Template for Weekly Sales Reporting 2026_CW23_Rep.xlsx"), [(MON, "Acme Test Co.")],
                          sheet="CW23")
        result = read_paths([real, stale])
        (v,) = result.parsed.visits
        self.assertEqual(v.workbook, os.path.basename(real))
        self.assertIn(v.visit_key, result.auto_accepted)

    def test_winner_is_picked_among_visits_that_would_be_accepted(self):
        from types import SimpleNamespace as NS
        from excel.run import _winner
        good = NS(eligible=True, week_ok=False, cw=21, mtime=1.0, path="a", sheet="CW21")
        newer_blocked = NS(eligible=False, week_ok=True, cw=23, mtime=9.0, path="b", sheet="CW21")
        self.assertIs(_winner([newer_blocked, good]), good)
        only = [newer_blocked, NS(eligible=False, week_ok=False, cw=25, mtime=9.0, path="c", sheet="CW21")]
        self.assertIs(_winner(only), only[0])  # none acceptable: still one of them, the one matching its week

    def test_two_visits_to_one_customer_on_one_day_do_not_wipe_each_other(self):
        old = book_with(self.path("Template for Weekly Sales Reporting 2026_CW21_Rep.xlsx"),
                        [(MON, "Acme Test Co."), (MON, "Acme Test Co.")])
        new = book_with(self.path("Template for Weekly Sales Reporting 2026_CW22_Rep.xlsx"),
                        [(MON, "Acme Test Co."), (MON, "Acme Test Co."), (datetime(2026, 5, 25), "Beta Test GmbH")],
                        sheet="CW22")
        result = read_paths([old, new])
        got = sorted((v.date_iso, v.customer_raw, v.key_ordinal) for v in result.parsed.visits)
        self.assertEqual(got, [("2026-05-18", "Acme Test Co.", 0), ("2026-05-18", "Acme Test Co.", 1),
                               ("2026-05-25", "Beta Test GmbH", 0)])

    def test_same_rows_in_two_sheets_of_one_workbook_are_duplicates(self):
        cells = dict(WEEKLY_HEADER)
        cells.update({"E4": "Owner：", "F4": "Jason Zhao", "B9": "Jason Zhao", "D9": MON,
                      "F9": "Acme Test Co.", "P9": "Call"})
        p = write_book(self.path("Template for Weekly Sales Reporting 2026_Rep.xlsx"),
                       {"CW20": {"cells": cells}, "CW21": {"cells": cells}})
        result = read_paths([p])
        self.assertEqual([v.sheet for v in result.parsed.visits], ["CW21"])
        self.assertEqual(len(result.auto_accepted), 1)
        self.assertEqual([(x["ignored_sheet"], x["used_sheet"]) for x in result.superseded], [("CW20", "CW21")])


LUCAS_PLAN = {"C4": "Owner：", "D4": "Lucas Guo"}


def lucas_week(path, plans, customer):
    from tests.excel_book import LUCAS_HEADER
    cells = dict(LUCAS_HEADER)
    cells.update(LUCAS_PLAN)
    for r, (day, plan) in zip(range(6, 12), plans):
        cells.update({f"C{r}": day, f"D{r}": plan})
    cells.update({"C15": datetime(2026, 3, 2), "E15": customer, "L15": "Phone call"})
    return write_book(path, {"W01": {"cells": cells}})


class WeekNotes(TmpCase):
    def test_two_files_with_one_week_make_one_week_note(self):
        a = lucas_week(self.path("Weekly Sales Planning and Reporting 2026_CW10-Lucas.xlsx"),
                       [("Monday", "Alpha trip")], "Acme Test Co.")
        b = lucas_week(self.path("Weekly Sales Planning and Reporting 2026_CW10-Lucas2.xlsx"),
                       [("Monday", "Beta trip"), ("Tuesday", "Gamma")], "Beta Test GmbH")
        result = read_paths([a, b])
        (week,) = result.parsed.weeks
        plans = {d.weekday: d.plan for d in week.days}
        self.assertIn("Alpha trip", plans["Monday"])
        self.assertIn("Beta trip", plans["Monday"])
        self.assertEqual(plans["Tuesday"], "Gamma")


if __name__ == "__main__":
    unittest.main()


def test_a_guessed_month_keeps_a_visit_out_of_the_auto_accepted_set():
    from excel.run import AUTO_OK
    assert "date_month_guessed" not in AUTO_OK
