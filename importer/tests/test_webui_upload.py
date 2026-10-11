"""The web UI upload must read the real weekly-report layouts, not only the old
"Date in column A" visit list. Workbooks here are made up."""
import io
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
