import unittest

from excel.split import latin_alias, looks_like_person_name, phone_numbers, phone_text, split_customer, visit_kind


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

    def test_long_multi_word_name_without_digit_or_marker_is_sure(self):
        name = "Qiqihar Heavy CNC Equipment Corporation Limited"
        self.assertEqual(split_customer(name), (name, "", True))
        self.assertFalse(split_customer("Some Name " + "x" * 80)[2])
        self.assertFalse(split_customer("Hafei Airbus Composite Materials Centre 2")[2])

    def test_fair_and_messe_are_exhibitions(self):
        self.assertEqual(visit_kind("The 29th Beijing Essen Welding & Cutting Fair"), "exhibition")
        self.assertEqual(visit_kind("EMO Messe Hannover"), "exhibition")


class PhoneText(unittest.TestCase):
    def test_numbers_lose_the_decimal_point(self):
        self.assertEqual(phone_text(75529900107.0), "75529900107")

    def test_text_kept_as_written(self):
        self.assertEqual(phone_text("0755-29900107"), "0755-29900107")
        self.assertEqual(phone_text(None), "")


class Placeholders(unittest.TestCase):
    def test_placeholder_customer_is_empty(self):
        for text in ("/", "-", "—", "Not Offer", "none", "N/A", "无", "暂无", "没有"):
            with self.subTest(text=text):
                self.assertEqual(split_customer(text), ("", "", True))

    def test_placeholder_phone_is_empty(self):
        self.assertEqual(phone_text("暂无"), "")
        self.assertEqual(phone_text(" - "), "")
        self.assertEqual(phone_numbers("/"), [])


class NameStaysWhole(unittest.TestCase):
    def test_short_latin_alias_is_not_split_off(self):
        self.assertEqual(split_customer("wendaxin 文达新"), ("wendaxin 文达新", "", True))

    def test_company_names_with_suffix_stay_whole(self):
        self.assertEqual(split_customer("CR0701256 四川成都某某科技有限公司"),
                         ("CR0701256 四川成都某某科技有限公司", "", True))
        self.assertEqual(split_customer("MINO 广州天河某某科技股份有限公司"),
                         ("MINO 广州天河某某科技股份有限公司", "", True))

    def test_marker_before_six_chars_does_not_split(self):
        self.assertEqual(split_customer("Ab 上海路"), ("Ab 上海路", "", True))

    def test_marker_after_six_chars_splits_unsure(self):
        name, address, sure = split_customer("Acme Test 上海市长宁区测试路1号")
        self.assertFalse(sure)
        self.assertTrue(address.startswith("市"))


class LatinAlias(unittest.TestCase):
    def test_chinese_name_with_latin_in_brackets(self):
        self.assertEqual(latin_alias("上海某某科技有限公司(Some Name)"), ("上海某某科技有限公司", "Some Name"))
        self.assertEqual(latin_alias("上海某某科技有限公司（Some Name）"), ("上海某某科技有限公司", "Some Name"))

    def test_open_bracket_without_close(self):
        self.assertEqual(latin_alias("上海某某科技有限公司（Some Name"), ("上海某某科技有限公司", "Some Name"))

    def test_other_cells_unchanged(self):
        self.assertEqual(latin_alias("wendaxin 文达新"), ("wendaxin 文达新", ""))
        self.assertEqual(latin_alias("Acme Test (Shanghai)"), ("Acme Test (Shanghai)", ""))


class PhoneNumbers(unittest.TestCase):
    def test_slash_comma_and_space_separate_numbers(self):
        self.assertEqual(phone_numbers("0755-29900107 / 13800138000"), ["0755-29900107", "13800138000"])
        self.assertEqual(phone_numbers("13800138000, 13900139000"), ["13800138000", "13900139000"])
        self.assertEqual(phone_numbers("13800138000;13900139000\n13700137000"),
                         ["13800138000", "13900139000", "13700137000"])

    def test_country_code_and_area_code_stay_with_number(self):
        self.assertEqual(phone_numbers("+86 13800138000"), ["+86 13800138000"])
        self.assertEqual(phone_numbers("0755 29900107"), ["0755 29900107"])

    def test_trailing_words_stay_with_the_number(self):
        self.assertEqual(phone_numbers("13800138000 (Mr. Li)"), ["13800138000 (Mr. Li)"])

    def test_glued_long_digit_run_is_two_numbers(self):
        self.assertEqual(phone_numbers("1380013800013900139000"), ["13800138000", "13900139000"])

    def test_phone_text_keeps_a_single_number_as_written(self):
        self.assertEqual(phone_text("13800138000"), "13800138000")
        self.assertEqual(phone_numbers(75529900107.0), ["75529900107"])


class Kind(unittest.TestCase):
    def test_internal_rows(self):
        for text in ("Public Holiday", "国庆假期", "休息", "Meeting Online", "Internal Meeting",
                     "Office", "Training", "培训", "Team Building"):
            with self.subTest(text=text):
                self.assertEqual(visit_kind(text), "internal")

    def test_a_to_do_written_in_the_customer_cell_is_internal(self):
        for text in ("Register a Tik Tok account.", "Get ready for the live broadcast.", "Attend weekly offline meetings.",
                     "Upload it on the web page. Related product information"):
            with self.subTest(text=text):
                self.assertEqual(visit_kind(text), "internal")
        self.assertEqual(visit_kind("Harbin Institute of Technology."), "")  # a name that ends with a full stop is a customer

    def test_exhibition_rows(self):
        for text in ("Exhibition", "展会", "CIIE Expo"):
            with self.subTest(text=text):
                self.assertEqual(visit_kind(text), "exhibition")

    def test_customer_rows(self):
        self.assertEqual(visit_kind("wendaxin 文达新"), "")
        self.assertEqual(visit_kind(""), "")
        # a sentence in the customer cell is no company, with or without a keyword
        for text in ("Visit the dealer in the south.", "Waiting for the samples to arrive"):
            with self.subTest(text=text):
                self.assertEqual(visit_kind(text), "internal")
        self.assertEqual(visit_kind("Institute of Technology of the North"), "")  # one grammar word: still a name
        self.assertEqual(visit_kind("Harbin Institute of Technology"), "")
        self.assertEqual(visit_kind("Gate to the East Co., Ltd"), "")  # a legal suffix keeps it a company
        self.assertEqual(visit_kind("on the way to xi'an 银川文达新"), "")  # holds a Chinese name: a person decides


if __name__ == "__main__":
    unittest.main()


class PersonName(unittest.TestCase):
    def test_what_cannot_be_a_person(self):
        for text in ("2026-05-25", "25/05", "Mr.", "Ms", "Dr. Prof.", "13851689143", "", None, " - "):
            with self.subTest(text=text):
                self.assertFalse(looks_like_person_name(text))

    def test_what_can_be_a_person(self):
        for text in ("Mr. Chen", "张三", "Geng Chen", "wang工", "A.Li"):
            with self.subTest(text=text):
                self.assertTrue(looks_like_person_name(text))
