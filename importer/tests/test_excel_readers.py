import os
import unittest
from datetime import datetime

from excel.layouts import detect
from excel.readers import read_sheet
from excel.sheet import load
from tests.excel_book import LUCAS_HEADER, WEEKLY_HEADER, TmpCase, write_book
from visits.reader import parse_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MOCK = os.path.join(ROOT, "..", "mock_visits", "excel", "visit_list_Chen_Hao_2025.xlsx")
WEEKLY_NAME = "Template for Weekly Sales Reporting 2026_CW21_Test Rep.xlsx"
LUCAS_NAME = "Weekly Sales Planning and Reporting 2026_CW15-Lucas1.xlsx"
DETAIL_NAME = "2026_Template for Customer Visits in Detail-20260918.xlsx"


def weekly_cells(**extra):
    cells = dict(WEEKLY_HEADER)
    cells.update({"E4": "Owner：", "F4": "Jason Zhao"})
    cells.update(extra)
    return cells


class WeeklyLayout(TmpCase):
    def read(self, name, sheets):
        path = write_book(self.path(name), sheets)
        return [read_sheet(s, path) for s in load(path)]

    def test_merged_block_with_next_steps_and_spillover(self):
        cells = weekly_cells(
            B9="Jason Zhao", C9="CW21", D9=datetime(2026, 5, 18), E9="Shanghai",
            F9="Acme Test Co., Ltd., 1 Example Road", G9="B", H9="NEW", I9="Test Person",
            J9="13800138000", N9="Intro visit", N10="Second line",
            P9="1.Send brochure", P10="2.", P11="3.Book demo", Q9=5.25, R9="On going",
            B14="Jason Zhao", C14="CW21", D14="=D9+1", F14="Beta Test GmbH", P14="Call back",
        )
        sheets = {"CW21": {"cells": cells, "merges": ["F9:F12"]}}
        ((layout, visits, week, review),) = self.read(WEEKLY_NAME, sheets)
        self.assertEqual(layout, "weekly2026")
        self.assertIsNone(week)
        self.assertEqual(len(visits), 2)
        first, second = visits
        self.assertEqual((first.excel_row_start, first.excel_row_end), (9, 13))
        self.assertEqual(first.customer_raw, "Acme Test Co., Ltd.")
        self.assertEqual(first.owner, "Jason Zhao")
        self.assertEqual((first.date_iso, first.potential, first.customer_type_raw),
                         ("2026-05-18", "B", "NEW"))
        self.assertEqual((first.contact_raw, first.phone_raw), ("Test Person", "13800138000"))
        self.assertEqual(first.purpose, "Intro visit\nSecond line")
        self.assertEqual([n.text for n in first.notes], ["1.Send brochure", "3.Book demo"])
        self.assertEqual({n.deadline for n in first.notes}, {"2026-05-25"})
        self.assertEqual({n.highlight for n in first.notes}, {"green"})
        self.assertEqual(second.date_iso, "2026-05-19")
        self.assertEqual(second.customer_raw, "Beta Test GmbH")
        self.assertIn("due_decimal", {r.reason for r in review})

    def test_flat_rows_are_one_visit_each(self):
        cells = weekly_cells(
            B9="Mark Wang", D9=datetime(2026, 5, 18), F9="Flat One", N9="a",
            B10="Mark Wang", D10=datetime(2026, 5, 19), F10="Flat Two", N10="b",
            B11="Mark Wang", D11=datetime(2026, 5, 20), F11="Flat Three", N11="c",
        )
        ((_, visits, _, _),) = self.read(WEEKLY_NAME, {"CW21": {"cells": cells}})
        self.assertEqual([v.customer_raw for v in visits], ["Flat One", "Flat Two", "Flat Three"])
        self.assertEqual([v.excel_row_start for v in visits], [9, 10, 11])

    def test_empty_sheet_without_header_is_skipped(self):
        ((layout, visits, week, review),) = self.read(WEEKLY_NAME, {"Sheet1": {"cells": {"A1": "x"}}})
        self.assertIsNone(layout)
        self.assertEqual((visits, week, review), ([], None, []))


class WeeklyFixes(TmpCase):
    def read(self, cells):
        path = write_book(self.path(WEEKLY_NAME), {"CW21": {"cells": cells}})
        ((_, visits, _, review),) = [read_sheet(s, path) for s in load(path)]
        return visits, review

    def test_date_without_customer_is_still_a_visit(self):
        visits, review = self.read(weekly_cells(B9="Jason Zhao", D9=datetime(2026, 5, 18), P9="Call"))
        self.assertEqual([(v.customer_raw, v.date_iso) for v in visits], [("", "2026-05-18")])
        self.assertIn("no_customer", {r.reason for r in review})

    def test_placeholders_are_empty_in_text_fields(self):
        visits, review = self.read(weekly_cells(
            B9="Jason Zhao", D9=datetime(2026, 5, 18), F9="/", G9="-", H9="none", I9="暂无",
            J9="无", L9="n/a", E9="没有", N9="Not offer", R9="—"))
        (v,) = visits
        self.assertEqual((v.customer_raw, v.potential, v.customer_type_raw, v.contact_raw), ("", "", "", ""))
        self.assertEqual((v.phone_raw, v.email, v.city, v.purpose), ("", "", "", ""))
        self.assertEqual(v.extra["status"], "")
        self.assertIn("no_customer", {r.reason for r in review})

    def test_year_1900_date_takes_sheet_week_year(self):
        visits, review = self.read(weekly_cells(B9="Jason Zhao", D9=datetime(1900, 5, 18), F9="Acme Test Co."))
        self.assertEqual(visits[0].date_iso, "2026-05-18")
        self.assertNotIn("no_date", {r.reason for r in review})

    def test_two_phone_numbers_split(self):
        visits, _ = self.read(weekly_cells(B9="Jason Zhao", D9=datetime(2026, 5, 18), F9="Acme Test Co.",
                                           J9="13800138000 / 13900139000"))
        self.assertEqual(visits[0].phone_raw, "13800138000")
        self.assertEqual(visits[0].extra["phone_more"], "13900139000")

    def test_latin_alias_goes_to_name_en(self):
        visits, _ = self.read(weekly_cells(B9="Jason Zhao", D9=datetime(2026, 5, 18),
                                           F9="上海某某科技有限公司(Some Name)"))
        self.assertEqual(visits[0].customer_raw, "上海某某科技有限公司")
        self.assertEqual(visits[0].extra["name_en"], "Some Name")

    def test_internal_row_is_kept_with_kind(self):
        visits, review = self.read(weekly_cells(B9="Jason Zhao", D9=datetime(2026, 5, 18), F9="Public Holiday"))
        self.assertEqual(visits[0].customer_raw, "Public Holiday")
        self.assertEqual(visits[0].extra["kind"], "internal")


class LucasLayout(TmpCase):
    def test_block_key_plan_and_follow_up(self):
        cells = dict(LUCAS_HEADER)
        cells.update({
            "C4": "Owner：", "D4": "Lucas Guo", "H4": "Timing： ", "I4": "CW10",
            "C7": "Monday", "D7": "Visiting Test Co", "C9": "Wed", "D9": "Visit Beta",
            "C15": datetime(2026, 3, 2), "D15": "Shenzhen", "E15": "Acme Test Co.",
            "F15": "A", "G15": "exist", "H15": "Name Test", "I15": "0755-1234567",
            "J15": "Office addr", "K15": "a@test.example", "L15": "Phone call",
            "M15": "Send list", "N15": "Follow the offer", "O15": "next week",
        })
        path = write_book(self.path(LUCAS_NAME), {"W01": {"cells": cells}})
        ((layout, visits, week, review),) = [read_sheet(s, path) for s in load(path)]
        self.assertEqual(layout, "lucas")
        self.assertEqual(detect(load(path)[0], path), "lucas")
        (visit,) = visits
        self.assertEqual(visit.date_iso, "2026-03-02")
        self.assertEqual(visit.extra["follow_up"], "Follow the offer")
        self.assertEqual(visit.notes[0].deadline, "next week")
        self.assertEqual(week.monday_iso, "2026-03-02")
        self.assertEqual([(d.weekday, d.date_iso, d.plan) for d in week.days],
                         [("Monday", "2026-03-02", "Visiting Test Co"),
                          ("Wed", "2026-03-04", "Visit Beta")])


class DetailLayout(TmpCase):
    def test_one_meeting_per_sheet(self):
        cells = {
            "B2": "Template for Customer Visits", "B3": "Customer name:", "D3": "Test Expo Hall",
            "B5": "Participants", "D5": "Person One", "F5": "director",
            "D6": "Person Two", "F6": "manager",
            "B11": "ID", "C11": "Date", "D11": "Typ\n(E/I/A)", "E11": "from\n(Person)",
            "F11": "General topic", "G11": "Details", "H11": "WV", "I11": "Wer", "J11": "Status",
            "C12": "2026.9.15", "D12": "A", "E12": "Jason Yi", "F12": "visit to Test Co",
            "G12": "Line A", "G13": "Line B continues",
            "C14": "2026.9.16", "F14": "second topic", "G14": "Detail two",
            "H14": datetime(2026, 9, 30), "I14": "Jason Yi",
        }
        path = write_book(self.path(DETAIL_NAME), {"Tabelle1": {"cells": cells}})
        ((layout, visits, week, review),) = [read_sheet(s, path) for s in load(path)]
        self.assertEqual(layout, "detail")
        self.assertIsNone(week)
        (visit,) = visits
        self.assertEqual(visit.customer_raw, "Test Expo Hall")
        self.assertEqual(visit.date_iso, "2026-09-15")
        self.assertEqual(visit.purpose, "visit to Test Co; second topic")
        self.assertEqual((visit.excel_row_start, visit.excel_row_end), (12, 14))
        self.assertEqual(visit.extra["participants"], [
            {"name": "Person One", "role": "director"}, {"name": "Person Two", "role": "manager"}])
        self.assertEqual(visit.notes[0].text, "Line A\nLine B continues")
        green = [n for n in visit.notes if n.highlight == "green"]
        self.assertEqual([(n.deadline, n.excel_row) for n in green], [("2026-09-30", 14)])
        self.assertIn("participants", {r.reason for r in review})


class MockLayout(TmpCase):
    def test_matches_visits_reader_on_mock_workbook(self):
        if not os.path.exists(MOCK):
            self.skipTest("mock workbook not present")
        expected = parse_workbook(MOCK)
        visits, weeks = [], []
        for sheet in load(MOCK):
            layout, got, week, _ = read_sheet(sheet, MOCK)
            if layout is None:
                continue
            self.assertEqual(layout, "mock")
            visits += got
            if week:
                weeks.append(week)
        self.assertEqual(visits, expected.visits)
        self.assertEqual(weeks, expected.weeks)


if __name__ == "__main__":
    unittest.main()
