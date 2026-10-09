import unittest

from excel.split import phone_text, split_customer


class SplitCustomer(unittest.TestCase):
    def test_newline_separates_name_and_address(self):
        self.assertEqual(split_customer("Acme Test Co.\nNo. 1 Example Road"),
                         ("Acme Test Co.", "No. 1 Example Road", True))

    def test_legal_suffix_ends_the_name(self):
        self.assertEqual(split_customer("Acme Test 股份有限公司, Room 5, Building 2"),
                         ("Acme Test 股份有限公司", "Room 5, Building 2", True))

    def test_gmbh_is_a_suffix(self):
        self.assertEqual(split_customer("Example GmbH"), ("Example GmbH", "", True))

    def test_address_marker_splits_but_is_not_sure(self):
        name, address, sure = split_customer("Acme Test 上海市长宁区测试路1号")
        self.assertFalse(sure)
        self.assertIn("路", address)

    def test_plain_name_sure_only_when_short(self):
        self.assertEqual(split_customer("Plain Name"), ("Plain Name", "", True))
        self.assertFalse(split_customer("x" * 41)[2])


class PhoneText(unittest.TestCase):
    def test_numbers_lose_the_decimal_point(self):
        self.assertEqual(phone_text(75529900107.0), "75529900107")

    def test_text_kept_as_written(self):
        self.assertEqual(phone_text("0755-29900107"), "0755-29900107")
        self.assertEqual(phone_text(None), "")


if __name__ == "__main__":
    unittest.main()
