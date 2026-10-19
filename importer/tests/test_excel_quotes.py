"""Quote log reader and Margince adapter, on invented workbooks, plus the real file when present."""
import datetime as dt
import glob
import os

import openpyxl
import pytest

from excel import quotes, quotes_margince
from excel.companies import Master
from excel.sheet import load
from tests.conftest import SAMPLE_DIR
from tests.fake_site_crm import SiteCRM

MASTER = Master.from_records([{"source": "master", "code": "KD00001", "cn": "宇宙星辰科技有限公司",
                               "en": "Cosmos Stellar Co., Ltd.", "short": ""}])
HEADER = ["Quote #", "Product Group", "Article #", "Product Type", "Company", "End customer", "Contact", "Clerk",
          "Responsible", "Representative", "Rep Code", "RFQ date", "Quote date", "Quote ID", "QTY",
          "estimated quote value", "est. manufacturing costs", "sales price", "DB II %", "Total sales volume",
          "Lead Time", "order date", "ERP order #", "Status Sales Chance", "Lost Index", "index", "R&D Hours",
          "Follow Up #1", "Follow Up #2", "Follow Up #3", "Follow Up #4", "Note#1", "Note#2", "Note#3", "Note#4",
          "Follow Up Remark"]
COSMOS = "Cosmos Stellar Co., Ltd."
D = dt.datetime(2026, 3, 5)


def row(no="QU260100001", group="Motion", ptype="EM.16", company=COSMOS, responsible="Yu, Richard", rev=1, qty=2,
        price=1000, total=2000, status="Open Sales Chance", remark=None, quote_date=D, **extra):
    cells = dict.fromkeys(HEADER)
    cells.update({"Quote #": no, "Product Group": group, "Product Type": ptype, "Company": company,
                  "Responsible": responsible, "Quote ID": rev, "QTY": qty, "sales price": price,
                  "Total sales volume": total, "Status Sales Chance": status, "Follow Up Remark": remark,
                  "Quote date": quote_date})
    cells.update(extra)
    return [cells[h] for h in HEADER]


def write(tmp_path, rows, name="Quote Log_test.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Quotelog"
    ws.append(HEADER)
    for r in rows:
        ws.append(r)
    path = str(tmp_path / name)
    wb.save(path)
    return path


def read(tmp_path, rows):
    return quotes.read_paths([write(tmp_path, rows)], master=MASTER)


def test_header_is_recognised(tmp_path):
    assert quotes.looks_like_quotes(load(write(tmp_path, [row()]))[0])


def test_revisions_share_one_deal(tmp_path):
    b = read(tmp_path, [row(rev=1), row(rev=2, price=1100, total=2200)])
    assert len(b.quotes) == 2 and list(b.deals()) == ["QU260100001"]
    assert [q.revision for q in b.deals()["QU260100001"]] == [1, 2]


def test_reused_and_missing_numbers_get_row_keys(tmp_path):
    b = read(tmp_path, [row(), row(), row(no=None), row(no="QU2605000xx")])
    keys = sorted(q.deal_key for q in b.quotes)
    assert keys == ["QU260100001", "row3", "row4", "row5"]
    assert {i.reason for i in b.review} >= {"quote_number_reused", "quote_number_unusable"}


def test_dates_money_and_errors(tmp_path):
    b = read(tmp_path, [row(quote_date="30/09/2026", rev="#N/A", price="44070 €", total=None, **{"order date": "Rehfuss"})])
    q = b.quotes[0]
    assert q.quote_date == "2026-09-30" and q.revision == 1 and q.price == 44070.0 and q.order_date is None
    assert q.amount == 88140.0 and any("order date" in i for i in q.issues)


@pytest.mark.parametrize("status,bucket", [("Won Order", "won"), ("Open Sales Chance", "open"),
                                           ("Cancelled without explanation", "lost"), ("Outbid", "lost"),
                                           ("Projekt on Hold", "open"), ("", "unknown"), ("0.4", "unknown")])
def test_status_bucket(status, bucket):
    assert quotes.status_bucket(status) == bucket


def test_company_waiting_for_an_answer_is_held_out(tmp_path):
    b = read(tmp_path, [row(), row(no="QU260100002", company="Zzyzx Unknown Works")])
    assert len(b.accepted()) == 2  # a company no master entry resembles is simply new
    next(o for o in b.orgs.values() if o.name == "Zzyzx Unknown Works").queued = True  # what a "same company?" question sets
    assert [q.deal_key for q in b.accepted()] == ["QU260100001"] and quotes.summarize(b)["quotes_queued"] == 1


def seats():
    return SiteCRM(seats={"Richard Yu": "seat-richard"})


def test_commit_writes_deals_offers_and_is_idempotent(tmp_path):
    b = read(tmp_path, [
        row(rev=1, remark="2026.3.10: called"), row(rev=2, price=1100, total=2200),
        row(no="QU260100002", status="Won Order"),
        row(no="QU260100003", status="Cancelled without explanation")])
    crm = seats()
    first = quotes_margince.commit(crm, b)
    assert (first["deals_created"], first["offers_created"], first["notes_created"]) == (3, 4, 1)
    by = {d["name"].split("(")[-1].rstrip(")"): d for d in crm.deals}
    won, lost, live = by["QU260100002"], by["QU260100003"], by["QU260100001"]
    assert won["status"] == "open" and "Won order" in won["description"]  # a win here would double the ERP orders
    assert lost["status"] == "lost" and lost["lost_reason"] == "Cancelled without explanation"
    assert live["owner_id"] == "seat-richard" and live["currency"] == "EUR" and live["amount_minor"] == 220000
    assert [o["status"] for o in crm.offers[live["id"]]] == ["draft", "sent"]  # only the live revision is sent
    assert crm.offers[lost["id"]][0]["status"] == "rejected"
    assert crm.offers[won["id"]][0]["status"] == "sent"
    assert crm.offers[live["id"]][1]["line_items"][0]["unit_price_minor"] == 110000
    crm.writes.clear()
    second = quotes_margince.commit(crm, b)
    assert not crm.writes and second["deals_created"] == second["offers_created"] == second["notes_created"] == 0
    assert second["deals_present"] == 3


def test_a_new_revision_adds_one_offer(tmp_path):
    crm = seats()
    quotes_margince.commit(crm, read(tmp_path, [row(rev=1)]))
    crm.writes.clear()
    c = quotes_margince.commit(crm, read(tmp_path, [row(rev=1), row(rev=2, price=900, total=1800)]))
    assert (c["deals_created"], c["offers_created"]) == (0, 1) and len(crm.offers[crm.deals[0]["id"]]) == 2


def test_offer_the_crm_will_not_send_stays_a_draft(tmp_path):
    crm = seats()
    crm.refuse_send = True
    c = quotes_margince.commit(crm, read(tmp_path, [row()]))
    assert c["offers_left_draft"] == 1 and crm.offers[crm.deals[0]["id"]][0]["status"] == "draft"


def test_plan_reads_only(tmp_path):
    crm = seats()
    b = read(tmp_path, [row(), row(no="QU260100002")])
    quotes_margince.commit(crm, read(tmp_path, [row()]))
    crm.writes.clear()
    p = quotes_margince.plan(crm, b)
    assert (p["deals_new"], p["deals_present"], p["offers_new"]) == (1, 1, 1) and not crm.writes


def test_real_quote_log_reads():
    found = glob.glob(os.path.join(SAMPLE_DIR + " cleaned", "10 *", "Quote Log*.xlsx")) or \
        glob.glob("/Users/jonas/Mac Documents/Shanghai Company/CRM Excel Lists 2026-10-11 cleaned/10 */Quote Log*.xlsx")
    if not found:
        pytest.skip("real quote log not present")
    b = quotes.read_paths(found, master=MASTER)
    assert len(b.quotes) == 353 and quotes.summarize(b)["deals"] <= 353
