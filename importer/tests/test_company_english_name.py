"""Names in the CRM follow the Excel: a Chinese name written in the file is stored as Chinese, an English
one as English. A company the CRM already holds under one language gets the other language's name from
the Excel written next to it ("legal name", which the CRM search reads). Nothing is translated."""
from unittest import mock

from excel.companies import Master
from excel.run import read_paths
from excel_margince import accepted
from tests.excel_book import TmpCase
from tests.fake_crm import FakeCRM
from tests.excel_book import WEEKLY_HEADER, write_book
from tests.test_visits_pipeline import MON
from visits import pipeline
from visits.match import match

CN, EN = "甲乙丙机电有限公司", "Alpha Beta Gamma Co., Ltd."
LISTED = {"code": "C1", "cn": CN, "en": EN}


def weekly_file(path, sheets):
    """One customer visit (type "exist", a contact met) per sheet; the rows without a type would be internal."""
    out = {}
    for name, rows in sheets.items():
        cells = dict(WEEKLY_HEADER)
        cells.update({"E4": "Owner：", "F4": "Jason Zhao"})
        for i, (day, customer, _step) in enumerate(rows):
            cells.update({f"B{9 + i}": "Jason Zhao", f"D{9 + i}": day, f"F{9 + i}": customer,
                          f"H{9 + i}": "exist", f"I{9 + i}": "Name Test", f"P{9 + i}": "Call"})
        out[name] = {"cells": cells}
    return write_book(path, out)


def master(*records):
    return Master.from_records([{"source": "master", "short": "", **r} for r in records])


class Names(TmpCase):
    def run_import(self, crm, customer, recs=(LISTED,), week="CW21"):
        name = f"Template for Weekly Sales Reporting 2026_{week}_Rep.xlsx"
        with mock.patch("excel.run._master", return_value=master(*recs)):
            path = weekly_file(self.path(name), {week: [(MON, customer, "")]})
            parsed = accepted(read_paths([path]))
        report = pipeline.run(crm, parsed, match(parsed.visits))
        return report, parsed

    @staticmethod
    def holding(crm, display_name, **extra):
        crm.companies["co_existing"] = {"id": None, "display_name": display_name, "lifecycle": None, **extra}
        return crm.companies["co_existing"]

    def test_english_in_the_excel_creates_the_company_under_the_english_name(self):
        crm = FakeCRM()
        self.run_import(crm, EN)
        (stored,) = crm.companies.values()
        self.assertEqual(stored["display_name"], EN)
        self.assertNotIn("legal_name", stored)

    def test_chinese_in_the_excel_creates_the_company_under_the_chinese_name(self):
        crm = FakeCRM()
        self.run_import(crm, CN)
        (stored,) = crm.companies.values()
        self.assertEqual(stored["display_name"], CN)
        self.assertNotIn("legal_name", stored)

    def test_english_in_the_excel_for_a_company_the_crm_holds_in_chinese_is_written_next_to_it(self):
        crm = FakeCRM()
        held = self.holding(crm, CN)
        self.run_import(crm, EN)
        self.assertEqual(len(crm.companies), 1)  # no second company
        self.assertEqual((held["display_name"], held["legal_name"]), (CN, EN))

    def test_the_english_name_is_the_one_written_in_the_excel_not_the_lists(self):
        crm = FakeCRM()
        held = self.holding(crm, CN)
        self.run_import(crm, "Alpha Beta Gamma Group")  # close enough to the list's name to match it
        self.assertEqual(held["legal_name"], "Alpha Beta Gamma Group")

    def test_chinese_in_the_excel_for_a_company_the_crm_holds_in_english_is_written_next_to_it(self):
        crm = FakeCRM()
        held = self.holding(crm, EN)
        self.run_import(crm, CN)
        self.assertEqual(len(crm.companies), 1)
        self.assertEqual((held["display_name"], held["legal_name"]), (EN, CN))

    def test_a_legal_name_already_there_is_never_overwritten(self):
        crm = FakeCRM()
        held = self.holding(crm, CN, legal_name="Registered Name GmbH")
        self.run_import(crm, EN)
        self.assertEqual(held["legal_name"], "Registered Name GmbH")

    def test_same_language_adds_nothing(self):
        crm = FakeCRM()
        held = self.holding(crm, EN)
        self.run_import(crm, EN)
        self.assertNotIn("legal_name", held)

    def test_unknown_company_is_created_as_written_and_adds_nothing(self):
        crm = FakeCRM()
        self.run_import(crm, "Zeta Totally New Ltd.")
        (stored,) = crm.companies.values()
        self.assertEqual(stored["display_name"], "Zeta Totally New Ltd.")
        self.assertNotIn("legal_name", stored)

    def test_visits_keep_the_same_identity_so_a_reimport_adds_nothing(self):
        crm = FakeCRM()
        first, _ = self.run_import(crm, EN)
        again, _ = self.run_import(crm, EN, week="CW22")  # the same visit in a later copy of the file
        self.assertEqual((first.activities_created, again.activities_created, again.activities_idempotent), (1, 0, 1))
        self.assertEqual(len(crm.companies), 1)

    def test_visit_source_ids_are_what_they_were_before_the_change(self):
        _, parsed = self.run_import(FakeCRM(), EN)
        (visit,) = parsed.visits
        self.assertIn("alphabetagamma", visit.source_id.replace("甲乙丙机电有限公司", "alphabetagamma"))
        self.assertEqual(visit.customer_raw, CN)  # the list's name, as the importer has always filed it
        self.assertEqual(visit.extra["customer_as_written"], EN)
