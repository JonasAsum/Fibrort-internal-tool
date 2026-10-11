import unittest
from datetime import date, datetime

from openpyxl.utils.datetime import to_excel

from excel.dates import parse_date, parse_due

REF = date(2026, 5, 11)  # Monday of CW20


class ParseDate(unittest.TestCase):
    def test_iso_like_forms(self):
        for text in ("2026/5/18", "2026-5-18", "2026.5.18"):
            with self.subTest(text=text):
                self.assertEqual(parse_date(text, REF), ("2026-05-18", ""))

    def test_day_month_names_take_year_from_ref(self):
        self.assertEqual(parse_date("13-May", REF), ("2026-05-13", ""))
        self.assertEqual(parse_date("26th,May", date(2026, 5, 20)), ("2026-05-26", ""))
        self.assertEqual(parse_date("13/May", REF), ("2026-05-13", ""))

    def test_month_names_with_year(self):
        self.assertEqual(parse_date("15th June 2026", REF), ("2026-06-15", ""))
        self.assertEqual(parse_date("June 15th 2026", REF), ("2026-06-15", ""))
        self.assertEqual(parse_date("18th    may 2026", REF), ("2026-05-18", ""))

    def test_ranges_take_first_day(self):
        self.assertEqual(parse_date("1st June--5th June2026", REF), ("2026-06-01", "range"))
        self.assertEqual(parse_date("20th  --21st  may 2026", REF), ("2026-05-20", "range"))

    def test_numeric_slash_resolved_by_ref(self):
        # 05/06 is 5 June or 6 May; a ref of 8 June picks 5 June, of 4 May picks 6 May.
        self.assertEqual(parse_date("05/06/2026", date(2026, 6, 8)), ("2026-06-05", ""))
        self.assertEqual(parse_date("05/06/2026", date(2026, 5, 4)), ("2026-05-06", ""))

    def test_unresolved_ambiguity_is_flagged(self):
        self.assertEqual(parse_date("05/06/2026", None), ("2026-06-05", "ambiguous"))
        self.assertEqual(parse_date("05/06/2026", date(2026, 1, 15))[1], "ambiguous")

    def test_unambiguous_slash_form(self):
        self.assertEqual(parse_date("13/05/2026", None), ("2026-05-13", ""))

    def test_serial_and_datetime(self):
        self.assertEqual(parse_date(to_excel(datetime(2026, 5, 18)), None), ("2026-05-18", ""))
        self.assertEqual(parse_date(datetime(2026, 5, 18, 9, 30), None), ("2026-05-18", ""))

    def test_formula_fallback_adds_days(self):
        lookup = lambda r, c: datetime(2026, 5, 18) if (r, c) == (9, 4) else None
        self.assertEqual(parse_date("=D9+1", REF, lookup=lookup), ("2026-05-19", "formula"))

    def test_formula_without_source_is_unresolved(self):
        self.assertEqual(parse_date("=D9+1", REF), (None, "formula"))

    def test_empty_and_unparsed(self):
        for empty in (None, "", "/", "-"):
            self.assertEqual(parse_date(empty, REF), (None, "empty"))
        self.assertEqual(parse_date("next week", REF), (None, "unparsed"))

    def test_year_hint_used_without_ref(self):
        self.assertEqual(parse_date("13-May", None, year=2025), ("2025-05-13", ""))


class ParseDue(unittest.TestCase):
    def test_decimal_month_day(self):
        self.assertEqual(parse_due("5.25", date(2026, 5, 20)), ("2026-05-25", "", "decimal"))
        self.assertEqual(parse_due(6.5, date(2026, 5, 20)), ("2026-06-05", "", "decimal"))
        self.assertEqual(parse_due("4.15", date(2026, 4, 10)), ("2026-04-15", "", "decimal"))

    def test_single_digit_day_picks_candidate_on_or_after_visit(self):
        # 5.1 is 1 May or 10 May; the visit is on 4 May so 10 May is the one ahead.
        self.assertEqual(parse_due("5.1", date(2026, 5, 4)), ("2026-05-10", "", "decimal"))

    def test_decimal_year_rolls_forward(self):
        self.assertEqual(parse_due("1.5", date(2026, 12, 20)), ("2027-01-05", "", "decimal"))

    def test_text_dates_use_visit_day_as_ref(self):
        self.assertEqual(parse_due("05/06/2026", date(2026, 6, 8)), ("2026-06-05", "", ""))

    def test_datetime_and_empty(self):
        self.assertEqual(parse_due(datetime(2026, 5, 25), date(2026, 5, 20)), ("2026-05-25", "", ""))
        self.assertEqual(parse_due(None, date(2026, 5, 20)), (None, "", ""))

    def test_phrase_kept(self):
        self.assertEqual(parse_due("next week", date(2026, 5, 20)), (None, "next week", "phrase"))


class OldYearAndPlaceholders(unittest.TestCase):
    def test_year_1900_keeps_day_and_month_and_takes_ref_year(self):
        # Excel stores "d/mmm" typed without a year as 1900.
        self.assertEqual(parse_date(datetime(1900, 5, 18), REF), ("2026-05-18", ""))
        self.assertEqual(parse_date(datetime(1900, 5, 18), None, year=2025), ("2025-05-18", ""))

    def test_year_1900_due_date_takes_visit_year(self):
        self.assertEqual(parse_due(datetime(1900, 5, 25), date(2026, 5, 20)), ("2026-05-25", "", ""))

    def test_placeholders_are_empty(self):
        for text in ("/", "-", "—", "not offer", "NONE", " n/a ", "无", "暂无", "没有"):
            with self.subTest(text=text):
                self.assertEqual(parse_date(text, REF), (None, "empty"))
                self.assertEqual(parse_due(text, date(2026, 5, 20)), (None, "", ""))


if __name__ == "__main__":
    unittest.main()


def test_absurd_number_in_a_date_cell_is_unparsed_not_a_crash():
    from excel.dates import parse_date
    assert parse_date(10 ** 30, None, 2026) == (None, "unparsed")
    assert parse_date(1e300, None, 2026) == (None, "unparsed")
