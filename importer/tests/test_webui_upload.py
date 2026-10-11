"""The web UI upload must read the real weekly-report layouts, not only the old
"Date in column A" visit list. Workbooks here are made up."""
import io
import os
from datetime import datetime

from tests.excel_book import LUCAS_HEADER, WEEKLY_HEADER, TmpCase, write_book
from webui import server

LUCAS_NAME = "Weekly Sales Planning and Reporting 2026_CW15-Lucas1.xlsx"
WEEKLY_NAME = "Template for Weekly Sales Reporting 2026_CW21_Test Rep.xlsx"


def lucas_book(path):
    cells = dict(LUCAS_HEADER)
    cells.update({
        "C4": "Owner：", "D4": "Lucas Guo", "H4": "Timing： ", "I4": "CW10",
        "C15": datetime(2026, 3, 2), "D15": "Shenzhen", "E15": "Acme Test Co.",
        "F15": "A", "G15": "exist", "H15": "Name Test", "K15": "a@test.example",
        "L15": "Phone call",
    })
    return write_book(path, {"W01": {"cells": cells}})


def weekly_book(path):
    cells = dict(WEEKLY_HEADER)
    cells.update({"E4": "Owner：", "F4": "Jason Zhao", "B9": "Jason Zhao",
                  "D9": datetime(2026, 5, 18), "F9": "Beta Test Co.", "P9": "Call"})
    return write_book(path, {"CW21": {"cells": cells}})


class Upload(TmpCase):
    def setUp(self):
        super().setUp()
        self._upload_dir = server.UPLOAD_DIR
        server.UPLOAD_DIR = self.path("uploads")
        self.addCleanup(setattr, server, "UPLOAD_DIR", self._upload_dir)
        self.addCleanup(server.CURRENT.update, files=[], parsed=None, result=None, run=None)
        self.client = server.app.test_client()

    def upload(self, path, name):
        with open(path, "rb") as f:
            data = {"files": (io.BytesIO(f.read()), name)}
        return self.client.post("/api/upload", data=data, content_type="multipart/form-data")

    def test_lucas_weekly_plan_uploads(self):
        # A customer the master list doesn't know waits in the review queue, so
        # nothing is accepted here; the upload itself must still succeed.
        r = self.upload(lucas_book(self.path("a.xlsx")), LUCAS_NAME)
        self.assertEqual(r.status_code, 200, r.get_json())

    def test_weekly_reporting_template_uploads(self):
        r = self.upload(weekly_book(self.path("b.xlsx")), WEEKLY_NAME)
        self.assertEqual(r.status_code, 200, r.get_json())
        self.assertEqual(len(server.CURRENT["parsed"].visits), 1)

    def test_lucas_copy_named_for_another_week_shows_its_visits_as_unsure(self):
        # Like "…CW13-Lucas2.xlsx" holding CW10's visits: found, shown, not dropped.
        name = "Weekly Sales Planning and Reporting 2026_CW13-Lucas2.xlsx"
        r = self.upload(lucas_book(self.path("c.xlsx")), name)
        self.assertEqual(r.status_code, 200, r.get_json())
        d = r.get_json()
        self.assertGreater(d["found"], 0)
        self.assertEqual(d["found"], d["visits"] + d["unsure"])
        self.assertTrue(any("week" in why for why in d["unsure_reasons"]))

    def test_lucas_layout_in_its_own_week_is_found(self):
        name = "Weekly Sales Planning and Reporting 2026_CW10-Lucas.xlsx"
        d = self.upload(lucas_book(self.path("d.xlsx")), name).get_json()
        self.assertGreater(d["found"], 0)

    def test_workbook_without_any_visit_list_is_refused_with_a_reason_and_keeps_nothing(self):
        from tests.excel_book import write_book
        p = write_book(self.path("customers.xlsx"), {"List": {"cells": {"A1": "Customer Code", "B1": "Name"}}})
        r = self.upload(p, "customers.xlsx")
        self.assertEqual(r.status_code, 422)
        self.assertIn("No visit list recognised", r.get_json()["error"])
        self.assertIn("customers.xlsx", r.get_json()["error"])
        self.assertFalse(self.client.get("/api/dataset").get_json()["loaded"])

    def test_same_visit_in_two_files_is_reported(self):
        from tests.e2e.test_reasons import NAME21, NAME22, weekly
        self.upload(weekly(self.path("a.xlsx")), NAME21)
        a, b = weekly(self.path("a.xlsx")), weekly(self.path("b.xlsx"), sheet="CW22")
        with open(a, "rb") as fa, open(b, "rb") as fb:
            data = {"files": [(io.BytesIO(fa.read()), NAME21), (io.BytesIO(fb.read()), NAME22)]}
        d = self.client.post("/api/upload", data=data, content_type="multipart/form-data").get_json()
        self.assertEqual(d["found"], 1)
        self.assertTrue(any(NAME22 == f["file"] and NAME21 in f["why"] for f in d["skipped_files"]), d["skipped_files"])


def ownerless_book(path, customer="Beta Test Co."):
    """A weekly report whose cells name no sales rep."""
    cells = dict(WEEKLY_HEADER)
    cells.update({"D9": datetime(2026, 5, 18), "F9": customer, "P9": "Call"})
    return write_book(path, {"CW21": {"cells": cells}})


class UploadOwner(Upload):
    def post(self, files, **form):
        data = {"files": [(io.BytesIO(open(p, "rb").read()), name) for p, name in files], **form}
        return self.client.post("/api/upload", data=data, content_type="multipart/form-data")

    def owners(self):
        return [v.owner for v in server.CURRENT["parsed"].visits]

    def test_rep_folder_of_the_browser_path_is_kept(self):
        r = self.post([(ownerless_book(self.path("a.xlsx")), "Devin Guo/" + WEEKLY_NAME)])
        self.assertEqual(r.status_code, 200, r.get_json())
        self.assertEqual(self.owners(), ["Devin Guo"])

    def test_rep_chosen_on_the_page_wins(self):
        r = self.post([(ownerless_book(self.path("a.xlsx")), WEEKLY_NAME)], rep="Mark Wang")
        self.assertEqual(self.owners(), ["Mark Wang"], r.get_json())

    def test_rep_in_the_file_name_is_found(self):
        name = "Weekly Sales Reporting 2026_CW21_Lucas Guo.xlsx"
        self.post([(ownerless_book(self.path("a.xlsx")), name)])
        self.assertEqual(self.owners(), ["Lucas Guo"])

    def test_unknown_rep_is_refused(self):
        r = self.post([(ownerless_book(self.path("a.xlsx")), WEEKLY_NAME)], rep="Nobody Known")
        self.assertEqual(r.status_code, 400)

    def test_reps_are_listed_for_the_page(self):
        self.assertIn("Jason Zhao", self.client.get("/api/reps").get_json()["reps"])

    def test_two_files_with_one_name_overwrite_nothing(self):
        a = ownerless_book(self.path("a.xlsx"), "Acme Test Co.")
        b = ownerless_book(self.path("b.xlsx"), "Beta Test Co.")
        d = self.post([(a, "Jason Zhao/" + WEEKLY_NAME), (b, "Devin Guo/" + WEEKLY_NAME)]).get_json()
        self.assertEqual(d["found"], 1)
        self.assertTrue(any("same file name" in f["why"] for f in d["skipped_files"]), d["skipped_files"])
        saved = [os.path.join(dp, n) for dp, _, ns in os.walk(server.UPLOAD_DIR) for n in ns]
        self.assertEqual(len(saved), 2)

    def test_a_failed_upload_leaves_no_earlier_data_to_import(self):
        self.upload(weekly_book(self.path("b.xlsx")), WEEKLY_NAME)
        self.assertIsNotNone(server.CURRENT["parsed"])
        r = self.post([(self.path("b.xlsx"), "notes.txt")])
        self.assertEqual(r.status_code, 400)
        self.assertIsNone(server.CURRENT["parsed"])
        self.assertFalse(self.client.get("/api/dataset").get_json()["loaded"])
