import unittest

from visits.match import Contact
from visits.pipeline import contact_emails


class Emails(unittest.TestCase):
    def test_only_ascii_addresses_count(self):
        c = Contact(key="k", company_key="c", full_name="X", emails=["名字@test.example", "a.b@test.example", "x@例え.example"])
        self.assertEqual([e["email"] for e in contact_emails(c)], ["a.b@test.example"])
