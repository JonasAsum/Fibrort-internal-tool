import unittest

from excel.headers import field_for, find_header
from excel.sheet import load
from tests.excel_book import LUCAS_HEADER, WEEKLY_HEADER, TmpCase, write_book


class FindHeader(TmpCase):
    def load_one(self, cells):
        path = write_book(self.path("w.xlsx"), {"CW21": {"cells": cells}})
        return load(path)[0]

    def test_header_at_b8(self):
        row, cols = find_header(self.load_one(WEEKLY_HEADER))
        self.assertEqual(row, 8)
        self.assertEqual(cols["date"], 4)
        self.assertEqual(cols["customer"], 6)
        self.assertEqual(cols["customer_type"], 8)
        self.assertEqual(cols["next_step"], 16)
        self.assertEqual(cols["due"], 17)
        self.assertEqual(cols["status"], 18)

    def test_header_at_c14_lucas_shape(self):
        row, cols = find_header(self.load_one(LUCAS_HEADER))
        self.assertEqual(row, 14)
        self.assertEqual(cols["date"], 3)
        self.assertEqual(cols["purpose"], 12)
        self.assertEqual(cols["follow_up"], 14)
        self.assertEqual(cols["due"], 15)

    def test_andrew_variant_plan_and_activity(self):
        cells = {
            "B8": "Plan", "C8": "Date", "E8": "Location/City", "F8": "Customer",
            "G8": "Potential", "H8": "Customer (new, exist, dealer)", "I8": "Contact Name",
            "N8": "activity", "P8": "Next Step / Action", "Q8": "Due Date", "R8": "Status",
        }
        row, cols = find_header(self.load_one(cells))
        self.assertEqual((row, cols["plan"], cols["date"]), (8, 2, 3))
        self.assertEqual(cols["purpose"], 14)
        self.assertEqual(cols["customer_type"], 8)
        self.assertEqual(cols["customer"], 6)

    def test_no_header_without_date_and_four_fields(self):
        self.assertIsNone(find_header(self.load_one({"A1": "Date", "B1": "Name"})))


class Synonyms(unittest.TestCase):
    def test_labels(self):
        cases = [
            ("Visiting date", "date"), ("Due Date", "due"), ("deadline ", "due"),
            ("Customer \n(full address if new)", "customer"),
            ("Customer (new, exist, dealer)", "customer_type"),
            ("Visiting Customer", None), ("Owner：", None), ("Phone No.", "phone"),
            ("Competiton Products", "competitor"), ("Purpose for Visiting ", "purpose"),
            ("Next Step / Action", "next_step"), ("Name", "contact"),
            ("Location/City", "city"), ("follow up", "follow_up"), ("WeChat", "wechat"),
            ("Status", "status"),
        ]
        for label, field in cases:
            with self.subTest(label=label):
                self.assertEqual(field_for(label), field)


if __name__ == "__main__":
    unittest.main()
