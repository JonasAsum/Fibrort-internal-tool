"""The quality rules must fail when the readers are broken: inject faults into the readers and check that the
whole-population rules of each family notice. A rule that passes on broken output proves nothing."""
import collections

from excel import accounts, contacts
from tests.quality import test_quality_accounts as qa
from tests.quality import test_quality_contacts as qc


def rules(module, sample_dir):
    return collections.Counter(v["rule"] for v in module.collect_violations(module.World(sample_dir)))


def test_accounts_rules_see_numbers_that_are_not_placed(monkeypatch, sample_dir):
    monkeypatch.setattr(accounts, "parse_numbers", lambda text: ([], []))
    assert rules(qa, sample_dir)["number-not-placed"] > 100


def test_accounts_rules_see_phone_and_mobile_swapped(monkeypatch, sample_dir):
    table = list(accounts.ERP_TABLE)
    table = [("phone", ("=手机",)) if n == "phone" else ("mobile", ("=电话",)) if n == "mobile" else (n, p) for n, p in table]
    monkeypatch.setattr(accounts, "ERP_TABLE", table)
    assert rules(qa, sample_dir)["wrong-field"] > 100


def test_accounts_rules_see_two_cells_joined_in_one_field(monkeypatch, sample_dir):
    original = accounts._read_erp

    def joined(rd, h, cols):
        original(rd, h, cols)
        for row in rd.rows:
            row.address = f"{row.address} {row.cn}"
    monkeypatch.setattr(accounts, "_read_erp", joined)
    assert rules(qa, sample_dir)["mixed-cells"] > 100


def test_accounts_rules_see_a_value_lost_in_the_merge(monkeypatch, sample_dir):
    original = accounts._account

    def lossy(key, rows, info, items):
        a = original(key, rows, info, items)
        a.industry = ""
        return a
    monkeypatch.setattr(accounts, "_account", lossy)
    seen = rules(qa, sample_dir)
    assert seen["cell-lost"] > 100 and seen["merge-loses-value"] > 100


def test_contacts_rules_see_numbers_dropped(monkeypatch, sample_dir):
    original = contacts._reach
    monkeypatch.setattr(contacts, "_reach", lambda *a, **k: ([], original(*a, **k)[1]))
    seen = rules(qc, sample_dir)
    assert seen["number-not-placed"] > 10 and seen["number-unaccounted"] > 10


def test_contacts_rules_see_phone_and_email_swapped(monkeypatch, sample_dir):
    original = contacts._entry
    monkeypatch.setattr(contacts, "_entry", lambda ctx, sheet, r, name, phone, email, *rest, **kw:
                        original(ctx, sheet, r, name, email, phone, *rest, **kw))
    assert rules(qc, sample_dir)["email-not-placed"] > 5


def test_contacts_rules_see_a_label_dropped(monkeypatch, sample_dir):
    original = contacts._entry
    monkeypatch.setattr(contacts, "_entry", lambda *a, **kw: original(*a, **{**kw, "department": ""}))
    assert rules(qc, sample_dir)["cell-lost"] > 5
