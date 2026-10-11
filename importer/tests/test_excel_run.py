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


if __name__ == "__main__":
    unittest.main()
