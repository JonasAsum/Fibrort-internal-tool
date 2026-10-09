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


if __name__ == "__main__":
    unittest.main()
