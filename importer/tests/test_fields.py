import unittest

from common.fields import ensure_field


class FakeClient:
    def __init__(self, fields):
        self.fields = list(fields)
        self.created = []

    def custom_fields(self, object_):
        return list(self.fields)

    def create_custom_field(self, body):
        f = {**body, "column_name": "cf_" + body["label"].lower().replace(" ", "_").replace("(", "").replace(")", "")}
        self.fields.append(f)
        self.created.append(body["label"])
        return f


def picklist(label, options):
    return {"label": label, "type": "picklist", "options": options, "column_name": "cf_old"}


TYPES = ["Dealer", "End user", "Manufacturer"]


class EnsureField(unittest.TestCase):
    def test_creates_when_missing(self):
        c = FakeClient([])
        f = ensure_field(c, "company", "Customer type", "picklist", TYPES, "mail")
        self.assertEqual(c.created, ["Customer type"])
        self.assertEqual(f["label"], "Customer type")

    def test_reuses_field_that_accepts_every_option(self):
        c = FakeClient([picklist("customer type", TYPES + ["Other"])])
        f = ensure_field(c, "company", "Customer type", "picklist", TYPES, "mail")
        self.assertEqual(f["column_name"], "cf_old")
        self.assertEqual(c.created, [])

    def test_same_label_with_other_options_gets_its_own_field(self):
        # the Excel import made "Customer type" = new/existing/dealer
        c = FakeClient([picklist("Customer type", ["new", "existing", "dealer"])])
        f = ensure_field(c, "company", "Customer type", "picklist", TYPES, "mail")
        self.assertEqual(c.created, ["Customer type (mail)"])
        self.assertNotEqual(f["column_name"], "cf_old")

    def test_second_run_reuses_the_fallback_field(self):
        c = FakeClient([picklist("Customer type", ["new", "existing", "dealer"])])
        ensure_field(c, "company", "Customer type", "picklist", TYPES, "mail")
        ensure_field(c, "company", "Customer type", "picklist", TYPES, "mail")
        self.assertEqual(c.created, ["Customer type (mail)"])

    def test_same_label_with_other_type_gets_its_own_field(self):
        c = FakeClient([{"label": "Potential", "type": "text", "column_name": "cf_old"}])
        ensure_field(c, "company", "Potential", "picklist", ["A", "B", "C"], "visits")
        self.assertEqual(c.created, ["Potential (visits)"])

    def test_text_field_is_reused(self):
        c = FakeClient([{"label": "Sales reps", "type": "text", "column_name": "cf_old"}])
        f = ensure_field(c, "company", "Sales reps", "text", None, "visits")
        self.assertEqual(f["column_name"], "cf_old")


if __name__ == "__main__":
    unittest.main()
