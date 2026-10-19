"""The two-button company question: "New company" or "Existing company" + pick one."""
import io
import json
import os
from datetime import datetime
from unittest import mock

from excel.companies import Master
from tests.excel_book import TmpCase
from tests.test_excel_run import one_visit_book
from webui import server

NAME = "Template for Weekly Sales Reporting 2026_CW21_Test Rep.xlsx"
ASKED = "Delta Trade Holdings Inc."  # invented: close to, but not the same as, the master's name


def _master():
    return Master.from_records([
        {"source": "master", "code": "C002", "cn": "丁戊己贸易股份有限公司", "en": "Delta Trading Inc.", "short": ""},
        {"source": "master", "code": "C003", "cn": "庚辛壬仪器有限公司", "en": "Gamma Instruments Ltd.", "short": ""},
    ])


class CompanyAnswer(TmpCase):
    def setUp(self):
        super().setUp()
        self.master = _master()
        patch = mock.patch("excel.run._master", return_value=self.master)
        patch.start()
        self.addCleanup(patch.stop)
        self._upload_dir = server.UPLOAD_DIR
        server.UPLOAD_DIR = self.path("uploads")
        self.addCleanup(setattr, server, "UPLOAD_DIR", self._upload_dir)
        self.addCleanup(server._forget)
        self.client = server.app.test_client()
        self.decisions = os.path.join(self.tmp, "company_decisions.json")  # TmpCase seeds it with its invented customers
        self.before = self.saved()

    def saved(self):
        with open(self.decisions, encoding="utf-8") as f:
            return json.load(f)

    def upload(self, customer=ASKED):
        book = one_visit_book(self.path("w.xlsx"), datetime(2026, 5, 18), customer=customer)
        with open(book, "rb") as f:
            data = {"files": (io.BytesIO(f.read()), NAME)}
        r = self.client.post("/api/upload", data=data, content_type="multipart/form-data")
        self.assertEqual(r.status_code, 200, r.get_json())
        return r.get_json()

    def answer(self, name, code):
        return self.client.post("/api/company_answer", json={"name": name, "code": code})

    def test_unsure_company_comes_back_as_a_question_with_candidates(self):
        d = self.upload()
        (q,) = d["company_questions"]
        self.assertEqual(q["name"], ASKED)
        self.assertEqual(q["visits"], 1)
        self.assertEqual(q["candidates"][0]["code"], "C002")
        self.assertEqual(q["candidates"][0]["en"], "Delta Trading Inc.")  # both names, so the list is readable
        self.assertEqual(d["unsure"], 1)

    def test_search_finds_a_company_by_any_of_its_names_or_its_code(self):
        self.upload()
        for text in ("delta", "丁戊己", "C002"):
            r = self.client.get("/api/company_search", query_string={"q": text}).get_json()
            self.assertEqual([x["code"] for x in r["results"]], ["C002"], text)  # no loosely related extras

    def test_existing_company_answer_releases_the_visit_and_is_remembered(self):
        self.upload()
        r = self.answer(ASKED, "C002")
        self.assertEqual(r.status_code, 200, r.get_json())
        d = r.get_json()
        self.assertEqual((d["unsure"], d["company_questions"]), (0, []))
        self.assertEqual(self.saved(), {**self.before, ASKED: "C002"})
        self.assertEqual(self.upload()["unsure"], 0)  # the next upload does not ask again

    def test_new_company_answer_releases_the_visit_and_is_remembered(self):
        self.upload()
        d = self.answer(ASKED, None).get_json()
        self.assertEqual((d["unsure"], d["company_questions"]), (0, []))
        self.assertEqual(self.saved(), {**self.before, ASKED: None})

    def test_unknown_code_is_refused_and_nothing_is_saved(self):
        self.upload()
        r = self.answer(ASKED, "C999")
        self.assertEqual(r.status_code, 400)
        self.assertEqual(self.saved(), self.before)

    def test_answer_for_a_name_nobody_asked_about_is_refused(self):
        self.upload()
        r = self.answer("Something Else Ltd", None)
        self.assertEqual(r.status_code, 400)
        self.assertEqual(self.saved(), self.before)

    def test_answer_without_an_upload_is_refused(self):
        self.assertEqual(self.answer(ASKED, None).status_code, 400)

    def test_close_candidate_is_offered_as_a_suggestion_only_from_85_percent(self):
        (q,) = self.upload()["company_questions"]
        top = q["candidates"][0]
        self.assertEqual(q["suggest"], top if top["score"] >= 0.85 else None)

    def test_several_answers_in_one_request_are_saved_and_read_once(self):
        self.upload()
        with mock.patch.object(server, "_reread", wraps=server._reread) as spy:
            r = self.client.post("/api/company_answers", json={"answers": [{"name": ASKED, "code": "C002"}]})
        self.assertEqual(r.status_code, 200, r.get_json())
        self.assertEqual(spy.call_count, 1)
        self.assertEqual(r.get_json()["company_questions"], [])
        self.assertEqual(self.saved(), {**self.before, ASKED: "C002"})

    def test_one_wrong_answer_in_a_batch_saves_none(self):
        self.upload()
        r = self.client.post("/api/company_answers", json={"answers": [
            {"name": ASKED, "code": "C002"}, {"name": "Never Asked Ltd", "code": None}]})
        self.assertEqual(r.status_code, 400)
        self.assertEqual(self.saved(), self.before)
        self.assertEqual(self.client.post("/api/company_answers", json={"answers": []}).status_code, 400)


LONG, SHORT = "Shenzhen Kestrel Intelligent Equipment Co., Ltd", "Shenzhen Kestrel Intelligent"


class PairAnswer(CompanyAnswer):
    """Two spellings of one new company in the same upload."""

    def upload_pair(self):
        from tests.excel_book import LUCAS_HEADER, write_book
        cells = dict(LUCAS_HEADER)
        cells.update({"C4": "Owner：", "D4": "Lucas Guo", "H4": "Timing： ", "I4": "CW10"})
        for i, name in enumerate((SHORT, LONG)):
            r = 15 + i
            cells.update({f"C{r}": datetime(2026, 3, 2 + i), f"D{r}": "Shenzhen", f"E{r}": name, f"F{r}": "A",
                          f"G{r}": "new", f"H{r}": "Name Test", f"K{r}": "a@test.example", f"L{r}": "Phone call"})
        book = write_book(self.path("lucas.xlsx"), {"W01": {"cells": cells}})
        with open(book, "rb") as f:
            data = {"files": (io.BytesIO(f.read()), "Weekly Sales Planning and Reporting 2026_CW10-Lucas.xlsx")}
        r = self.client.post("/api/upload", data=data, content_type="multipart/form-data")
        self.assertEqual(r.status_code, 200, r.get_json())
        return r.get_json()

    def pair_answer(self, same, a=LONG, b=SHORT):
        return self.client.post("/api/company_pair_answer", json={"a": a, "b": b, "same": same})

    def test_lookalike_names_come_back_as_one_pair_question(self):
        d = self.upload_pair()
        self.assertEqual(d["unsure"], 2)
        (q,) = d["company_questions"]
        self.assertEqual((q["kind"], q["a"], q["b"]), ("pair", LONG, SHORT))

    def test_same_company_answer_makes_one_company(self):
        self.upload_pair()
        d = self.pair_answer(True).get_json()
        self.assertEqual((d["unsure"], d["company_questions"], d["companies"]), (0, [], 1))
        self.assertEqual(self.upload_pair()["unsure"], 0)  # remembered

    def test_different_company_answer_keeps_two_companies(self):
        self.upload_pair()
        d = self.pair_answer(False).get_json()
        self.assertEqual((d["unsure"], d["company_questions"], d["companies"]), (0, [], 2))

    def test_answer_for_a_pair_nobody_asked_about_is_refused(self):
        self.upload_pair()
        self.assertEqual(self.pair_answer(True, a="Orion Harbor Logistics").status_code, 400)
