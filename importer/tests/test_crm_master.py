"""The company master is read from the CRM; the customer list files only add short names and the USCC."""
from excel import crm_master, run as excel_run
from excel.companies import Master


class FakeClient:
    base_url = "http://crm.test"

    def __init__(self, companies, fields=None, fail=False):
        self.companies, self.fail = companies, fail
        self.fields = fields if fields is not None else [{"label": "ERP customer code", "column_name": "cf_code"},
                                                         {"label": "Unified social credit code", "column_name": "cf_uscc"}]

    def custom_fields(self, object_):
        return self.fields

    def paginate(self, path, params=None):
        if self.fail:
            raise RuntimeError("down")
        assert path == "/v1/companies"
        yield from self.companies


def fresh(monkeypatch, tmp_path, files=()):
    crm_master._CACHE.update(at=0.0, records=None)
    monkeypatch.setattr(Master, "read_files", staticmethod(lambda d: (list(files), [], {})))
    monkeypatch.setattr("excel.decisions.default_path", lambda d: tmp_path / "decisions.json")


CRM = [{"id": "i1", "display_name": "Shanghai Daiyi Technology Co., Ltd.", "legal_name": "上海戴屹科技有限公司", "cf_code": "CR0002076"},
       {"id": "i2", "display_name": "Hymson Laser Technology Group Co., Ltd", "legal_name": None, "cf_code": None,
        "cf_uscc": "91440300MA5DKJYC8U"}]


def test_a_crm_company_is_a_master_company_whatever_language_its_names_are_in(monkeypatch, tmp_path):
    fresh(monkeypatch, tmp_path)
    m = crm_master.load(lambda: FakeClient(CRM), ".")
    assert m.origin == "crm" and m.size == 2
    got = m.match("上海戴屹科技有限公司")
    assert got.code == "CR0002076" and got.method == "exact"
    assert m.match("Shanghai Daiyi Technology Co., Ltd.").code == "CR0002076"


def test_a_crm_company_without_an_erp_code_is_found_by_its_crm_id_and_carries_its_uscc(monkeypatch, tmp_path):
    fresh(monkeypatch, tmp_path)
    m = crm_master.load(lambda: FakeClient(CRM), ".")
    got = m.match("Hymson Laser Technology Group Co., Ltd")
    assert got.code == "crm:i2"
    assert m.get("crm:i2")["uscc"] == "91440300MA5DKJYC8U"


def test_the_lists_add_a_short_name_to_a_crm_company_and_keep_a_company_the_crm_does_not_hold_yet(monkeypatch, tmp_path):
    files = [{"source": "erp", "code": "CR0002076", "cn": "上海戴屹科技有限公司", "en": "", "short": "戴屹", "uscc": "", "legal": ""},
             {"source": "erp", "code": "CR9999999", "cn": "只在文件里的公司有限公司", "en": "", "short": "", "uscc": "", "legal": ""}]
    fresh(monkeypatch, tmp_path, files)
    m = crm_master.load(lambda: FakeClient(CRM), ".")
    assert m.get("CR0002076")["short"] == "戴屹" and m.get("CR9999999") is not None
    assert m.size == 3 and m.sources["lists_only"] == 1
    monkeypatch.setenv("CRM_MASTER_ONLY", "1")
    only = crm_master.load(lambda: FakeClient(CRM), ".")
    assert only.get("CR9999999") is None and only.size == 2


def test_a_list_company_the_crm_holds_under_the_same_name_without_a_code_is_one_company(monkeypatch, tmp_path):
    files = [{"source": "erp", "code": "CR0005555", "cn": "", "en": "Hymson Laser Technology Group Co., Ltd", "short": "", "uscc": "", "legal": ""}]
    fresh(monkeypatch, tmp_path, files)
    m = crm_master.load(lambda: FakeClient(CRM), ".")
    assert m.size == 2 and m.match("Hymson Laser Technology Group Co., Ltd").method == "exact"


def test_the_website_falls_back_to_the_files_and_says_why_when_the_crm_cannot_be_read(monkeypatch, tmp_path):
    from webui import server
    fresh(monkeypatch, tmp_path)
    monkeypatch.setattr(server, "_client", lambda: FakeClient([], fail=True))
    server._refresh_master()
    assert excel_run._CRM["master"] is None and "could not be read" in excel_run.master_note()
    monkeypatch.setattr(server, "_client", lambda: FakeClient([]))
    server._refresh_master()
    assert excel_run._CRM["master"] is None and "no companies" in excel_run.master_note()
    monkeypatch.setattr(server, "_client", lambda: FakeClient(CRM))
    server._refresh_master()
    assert excel_run._master(".").origin == "crm"
    excel_run._master.cache_clear()


def test_the_crm_is_logged_into_once_for_a_burst_of_reads(monkeypatch, tmp_path):
    fresh(monkeypatch, tmp_path)
    logins = []
    def make():
        logins.append(1)
        return FakeClient(CRM)
    crm_master.load(make, ".")
    crm_master.load(make, ".")
    assert len(logins) == 1
