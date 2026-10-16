"""Quality checks for the sales-order reader (folder 08), run on every row of every workbook.

The oracle (oracle_orders.py) is built from the cell ledger alone. Each rule below compares the reader with it
over the whole population: every file, every sheet, every row, every cell, every document total. A rule
listed in KNOWN_BUGS is a bug we know about and have not fixed: it is strict-xfail (the day it is fixed the
marker must be removed, and a surprise pass fails the run).

Declared drop rules (a cell or row that is not in a record must be named here, with its reason):
  DECLARED_DROPS = rows: blank, total row, column-totals row, repeated header, no document number
                   cells: above-header query / title rows, header row, placeholder-only text ("/", "-", "无")
                   files: kept-out kinds (invoice/packing list, stock list, formula workbook)
"""
import datetime as dt
import os
import re
from decimal import Decimal

import pytest

from excel import orders
from tests.quality import fam3_common as oc
from tests.quality import oracle_orders as oo
from tests.quality import snapshot, snapshot_orders

KNOWN_BUGS = {
    "column-beyond-cap": "excel/sheet.py reads 40 columns at most: the Material-property and Invoice-issued-value columns "
                         "(AO, AP) of the OIT exports are never read (the reader warns; not mine to change)",
}
ALL_RULES = sorted(set(KNOWN_BUGS) | {
    "file-kind", "kept-out-yields-nothing", "sheet-set", "column-meaning", "row-set", "field-value", "other-column-carried",
    "document-totals", "sheet-control-totals", "book-document-set", "book-totals", "book-quantity-and-lines",
    "book-currency", "book-cny", "book-status", "field-shape", "idempotent", "oracle-arithmetic",
    "cut-sheets-are-subsets", "one-customer-per-document",
})
PLACEHOLDERS = {"/", "-", "—", "not offer", "none", "n/a", "无", "暂无", "没有"}  # declared: a cell holding only one is empty
SENSITIVE = re.compile(r"bank|iban|swift|银行|账号|帐号|账户|纳税人|税号|\btax\s*(id|no|number)\b|\bvat\s*(id|no|number)\b", re.I)
STATUS_RULES = (  # declared status words -> key; first match wins
    ("cancelled", ("cancel", "取消", "作废", "已关闭")),
    ("in_progress", ("partially", "to receive", "部分", "待收货")),
    ("completed", ("completed", "已完成")),
    ("to_ship", ("to ship", "pending delivery", "待发货")),
    ("open", ("open", "开立")),
)


def _v(rule, file, sheet, row, detail):
    return {"rule": rule, "file": file, "sheet": sheet, "row": row, "detail": str(detail)}


def _clean(text):
    t = oc.norm(text)
    return "" if t.casefold() in PLACEHOLDERS else t


def _ledger_text(v):
    """A reader-side raw cell value in the ledger's text form."""
    if isinstance(v, dt.datetime):
        return v.isoformat(sep=" ") if v.time() != dt.time(0) else v.date().isoformat()
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _col(coord):
    return oc.ledger.col_number(re.match(r"[A-Z]+", coord)[0])


def _d(x):
    return None if x is None else Decimal(str(x))


def _same_cents(a, b):
    return a is not None and b is not None and oc.cent(Decimal(a)) == oc.cent(Decimal(b))


def _status_key(text):
    t = (text or "").lower()
    return next((k for k, words in STATUS_RULES if any(w in t for w in words)), "other")


@pytest.fixture(scope="session")
def facts(sample_dir):
    if not os.path.isdir(sample_dir):
        pytest.skip("sample folder not found")
    return oo.all_facts(sample_dir)


@pytest.fixture(scope="session")
def reads(facts):
    return {f.path: orders.read_orders(f.path) for f in facts}


@pytest.fixture(scope="session")
def book(facts):
    return orders.build_book([f.path for f in facts])


def _file_checks(facts, reads, out):
    for f in facts:
        r = reads[f.path]
        if r.kind != f.kind:
            out.append(_v("file-kind", f.file, "", 0, f"reader says {r.kind!r}, oracle says {f.kind!r} ({f.reason})"))
        elif f.kind != "orders" and not r.reason.endswith(f.reason):
            out.append(_v("file-kind", f.file, "", 0, f"reason {r.reason!r} does not say {f.reason!r}"))
        if set(r.headers_found) != set(f.sheets):
            out.append(_v("sheet-set", f.file, "", 0, f"reader reads sheets {sorted(r.headers_found)}, oracle {sorted(f.sheets)}"))
        for name, s in f.sheets.items():
            got = r.headers_found.get(name, {})
            for fld, (_, label) in s.columns.items():
                if oc.norm(got.get(fld) or "") != label:
                    out.append(_v("column-meaning", f.file, name, s.header_row,
                                  f"{fld}: reader uses column {got.get(fld)!r}, the header that means it is {label!r}"))
            for fld in got:
                if fld not in s.columns and fld != "currency":
                    out.append(_v("column-meaning", f.file, name, s.header_row, f"reader maps {fld} to {got[fld]!r}, oracle has none"))


def _row_checks(facts, reads, out):
    for f in facts:
        r = reads[f.path]
        by = {(ln.sheet, ln.row): ln for ln in r.lines}
        expected = set()
        for name, s in f.sheets.items():
            for ln in s.lines:
                expected.add((name, ln.row))
                got = by.get((name, ln.row))
                if got is None:
                    out.append(_v("row-set", f.file, name, ln.row, f"order line {ln.text['document_no']} not read"))
                    continue
                _line_checks(f, name, ln, got, out)
            skipped = {(x.sheet, x.row): x.reason for x in r.skipped if x.reason != "blank row"}
            for row, why in s.dropped:
                if (name, row) not in skipped:
                    out.append(_v("row-set", f.file, name, row, f"oracle: {why}; the reader read it as a line or lost it"))
            for key in skipped:
                if key[0] == name and key[1] not in {d[0] for d in s.dropped}:
                    out.append(_v("row-set", f.file, name, key[1], f"reader skipped a row ({skipped[key]}) the oracle takes as a line"))
        for key in by:
            if key not in expected:
                out.append(_v("row-set", f.file, key[0], key[1], "reader has a line the oracle does not"))


def _line_checks(f, sheet, ln, got, out):
    for fld in oo.FIELDS:
        raw = ln.text[fld]
        mine = got.__dict__[fld] if fld != "date" else got.date_iso
        if fld == "date":
            want = raw[:10] if raw else None
            if raw and not re.match(r"\d{4}-\d{2}-\d{2}", raw):
                out.append(_v("field-shape", f.file, sheet, ln.row, f"date cell {raw!r} is not an ISO date"))
            if mine != want:
                out.append(_v("field-value", f.file, sheet, ln.row, f"date: cell {raw!r} -> {mine!r}"))
        elif fld in oo.NUMERIC:
            want = ln.num[fld]
            if want is None and raw:
                out.append(_v("field-value", f.file, sheet, ln.row, f"{fld}: cell {raw!r} is not a plain number"))
            elif (_d(mine) if mine is not None else None) != want:
                out.append(_v("field-value", f.file, sheet, ln.row, f"{fld}: cell {raw!r} -> {mine!r}"))
        else:
            want = _clean(raw) or None
            if mine != want:
                out.append(_v("field-value", f.file, sheet, ln.row, f"{fld}: cell {raw!r} -> {mine!r}"))
    for label, text in ln.other.items():
        if _col(ln.other_coords[label]) > oo.READER_COLUMN_CAP:
            continue  # beyond the reader's column cap: the column-beyond-cap rule owns it
        if SENSITIVE.search(label):
            if label in got.extra:
                out.append(_v("other-column-carried", f.file, sheet, ln.row, f"sensitive column {label!r} was read"))
            continue
        if label not in got.extra or _ledger_text(got.extra[label]) != text:
            out.append(_v("other-column-carried", f.file, sheet, ln.row,
                          f"{ln.other_coords[label]} {label!r}: cell {text[:40]!r} -> {str(got.extra.get(label))[:40]!r}"))
    for label in got.extra:
        if label not in ln.other:
            out.append(_v("other-column-carried", f.file, sheet, ln.row, f"reader has extra column {label!r} without a cell"))


def _beyond_cap(facts, out):
    for f in facts:
        for name, s in f.sheets.items():
            for c in s.beyond_cap:
                out.append(_v("column-beyond-cap", f.file, name, c.row, f"{c.coord} {c.text[:30]!r}"))


def _document_checks(facts, reads, out):
    for f in facts:
        r = reads[f.path]
        want = {}
        for s in f.sheets.values():
            for ln in s.lines:
                t = want.setdefault(ln.text["document_no"], {"lines": 0, "quantity": Decimal(0), "net": Decimal(0),
                                                             "tax": Decimal(0), "gross": Decimal(0)})
                t["lines"] += 1
                for k, fld in (("quantity", "quantity"), ("net", "net_amount"), ("tax", "tax_amount"), ("gross", "gross_amount")):
                    t[k] += ln.num[fld] or Decimal(0)
        if set(r.documents) != set(want):
            out.append(_v("document-totals", f.file, "", 0, f"documents differ: {sorted(set(r.documents) ^ set(want))[:5]}"))
        for no, d in r.documents.items():
            w = want.get(no)
            if not w:
                continue
            got = {"lines": d.line_count, "quantity": _d(d.quantity), "net": _d(d.net_amount), "tax": _d(d.tax_amount),
                   "gross": _d(d.gross_amount)}
            for k in w:
                ok = got[k] == w[k] if k == "lines" else _same_cents(got[k], w[k])
                if not ok:
                    out.append(_v("document-totals", f.file, "", 0, f"{no} {k}: rows sum to {w[k]}, reader says {got[k]}"))
        # the file's own control totals against the rows (a property of the data, independent of any reader)
        for name, s in f.sheets.items():
            for t in s.totals_rows:
                for col, value in t["values"].items():
                    col_sum = sum((oc.dec(c.text) or Decimal(0) for ln in s.lines for c in [_cell_at(f, name, ln.row, col)] if c),
                                  Decimal(0))
                    if oc.cent(col_sum) != oc.cent(value):
                        out.append(_v("sheet-control-totals", f.file, name, t["row"],
                                      f"column {col}: the sheet states {value}, its rows sum to {col_sum}"))
                for fld, (col, label) in s.columns.items():
                    if fld in oo.NUMERIC and col in t["values"]:
                        mine = sum((_d(getattr(ln, {"quantity": "quantity", "net_amount": "net_amount", "tax_amount": "tax_amount",
                                                   "gross_amount": "gross_amount"}[fld])) or Decimal(0)
                                    for ln in r.lines if ln.sheet == name), Decimal(0))
                        if oc.cent(mine) != oc.cent(t["values"][col]):
                            out.append(_v("sheet-control-totals", f.file, name, t["row"],
                                          f"{fld} ({label}): the sheet states {t['values'][col]}, the reader's lines sum to {mine}"))


_CELLS = {}


def _cell_at(f, sheet, row, col):
    key = f.path
    if key not in _CELLS:
        _CELLS[key] = {(c.sheet, c.row, c.col): c for c in f.cells}
    return _CELLS[key].get((sheet, row, col))


def _stamp_key(f):
    return f.export_stamp or (f.newest_order_day + " 23:59:59" if f.newest_order_day else "")


def _blank(lines):
    return not any(_clean(ln.text["customer"]) for ln in lines) or all(
        ln.num["net_amount"] is None and ln.num["gross_amount"] is None for ln in lines)


def _book_checks(facts, book, out):
    order_files = [f for f in facts if f.kind == "orders"]
    holders = {}
    for f in order_files:
        per = {}
        for s in f.sheets.values():
            for ln in s.lines:
                per.setdefault((ln.entity, ln.text["document_no"]), []).append(ln)
        totals = oo.file_order_totals(f)
        for key, lines in per.items():
            holders.setdefault(key, []).append((f, oo.merged_lines(lines), totals[key]))
    expected, blank_docs = {}, []
    for key, found in holders.items():
        usable = [h for h in found if not _blank(h[1])]
        if not usable:
            blank_docs.append(key)
            continue
        newest = max(_stamp_key(h[0]) for h in usable)
        expected[key] = [h for h in usable if _stamp_key(h[0]) == newest]
    got = {}
    for o in book.orders:
        got.setdefault(o.document_no, []).append(o)
    exp_by_doc = {}
    for (ent, no), cand in expected.items():
        exp_by_doc.setdefault(no, []).append(((ent, no), cand))
    if len(book.orders) != len(expected):
        out.append(_v("book-document-set", "", "", 0, f"{len(book.orders)} orders in the book, {len(expected)} documents expected"))
    for no in set(exp_by_doc) - set(got):
        out.append(_v("book-document-set", "", "", 0, f"document {no} is in the files and not in the book"))
    for no in set(got) - set(exp_by_doc):
        out.append(_v("book-document-set", "", "", 0, f"order {no} is in the book and not in the files"))
    for no, pairs in exp_by_doc.items():
        mine = got.get(no, [])
        if len(mine) != len(pairs):
            continue
        used = set()
        for (key, cand) in pairs:
            ent = key[0]
            o = next((x for x in mine if (x.entity or "") == ent and id(x) not in used), None)
            if o is None:
                out.append(_v("book-document-set", "", "", 0, f"{no}: no order for selling company {ent!r}"))
                continue
            used.add(id(o))
            if o.file not in {h[0].display for h in cand}:
                out.append(_v("book-totals", o.file, o.sheet, o.row,
                              f"{no}: should come from the newest export {[h[0].display for h in cand]}"))
            ok = {"net": [], "tax": [], "gross": [], "quantity": [], "lines": []}
            for h in cand:
                t = h[2]
                ok["net"].append(_same_cents(o.net, t["net"]))
                ok["tax"].append(_same_cents(o.tax, t["tax"]))
                ok["gross"].append(_same_cents(o.gross, t["gross"]))
                ok["quantity"].append(sum(Decimal(str(i.quantity or 0)) for i in o.items) == t["quantity"])
                ok["lines"].append(len(o.items) == t["lines"])
            for k in ("net", "tax", "gross"):
                if not any(ok[k]):
                    out.append(_v("book-totals", o.file, o.sheet, o.row,
                                  f"{no} {k}: the book says {getattr(o, k)}, the rows of {[h[0].display for h in cand]} sum to "
                                  f"{[str(h[2][k]) for h in cand]}"))
            for k in ("quantity", "lines"):
                if not any(ok[k]):
                    out.append(_v("book-quantity-and-lines", o.file, o.sheet, o.row,
                                  f"{no} {k}: the book's items do not match the rows {[str(h[2][k]) for h in cand]}"))
            lines = cand[0][1]
            _currency(o, lines, out)
            if o.status_key != _status_key(o.status):
                out.append(_v("book-status", o.file, o.sheet, o.row, f"{no}: status {o.status!r} -> {o.status_key!r}"))
            customers = {c for c in (_clean(ln.text["customer"]) for ln in lines) if c}
            if len(customers) > 1:
                out.append(_v("one-customer-per-document", o.file, o.sheet, o.row, f"{no}: {sorted(customers)}"))
            if customers and o.customer not in customers:
                out.append(_v("one-customer-per-document", o.file, o.sheet, o.row, f"{no}: customer {o.customer!r} not on its lines"))


def _currency(o, lines, out):
    """The Chinese entity states local-currency totals: line columns (表体本币…), else the document total repeated."""
    entity = lines[0].entity
    local = "有限公司" in entity
    if not local:
        if o.cny_net is not None or o.cny_gross is not None or o.currency == "CNY":
            out.append(_v("book-cny", o.file, o.sheet, o.row, f"{o.document_no}: entity {entity!r} states no local currency but the order has CNY values"))
        return

    def total(line_attr, doc_attr):
        own = [getattr(ln, line_attr) for ln in lines]
        if own and all(v is not None for v in own):
            return sum(own, Decimal(0))
        docs = {getattr(ln, doc_attr) for ln in lines if getattr(ln, doc_attr) is not None}
        return docs.pop() if len(docs) == 1 else None

    want_net, want_gross = total("local_line_net", "local_doc_net"), total("local_line_gross", "local_doc_gross")
    for name, want, mine in (("cny_net", want_net, o.cny_net), ("cny_gross", want_gross, o.cny_gross)):
        if (want is None) != (mine is None) or (want is not None and not _same_cents(mine, want)):
            out.append(_v("book-cny", o.file, o.sheet, o.row, f"{o.document_no} {name}: columns give {want}, order has {mine}"))
    gross = sum((ln.num["gross_amount"] or Decimal(0) for ln in lines), Decimal(0))
    want_currency = "CNY" if want_gross is not None and abs(want_gross - gross) <= Decimal("0.05") else None
    if o.currency != want_currency:
        out.append(_v("book-currency", o.file, o.sheet, o.row,
                      f"{o.document_no}: CNY local total {want_gross}, transaction gross {gross}: currency should be {want_currency!r}, is {o.currency!r}"))


def _shape_checks(book, out):
    ids = [o.source_id for o in book.orders]
    for i in {x for x in ids if ids.count(x) > 1}:
        out.append(_v("field-shape", "", "", 0, f"order id {i} used twice"))
    for o in book.orders:
        if o.date_iso and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", o.date_iso):
            out.append(_v("field-shape", o.file, o.sheet, o.row, f"{o.document_no}: date {o.date_iso!r}"))
        if not o.customer:
            out.append(_v("field-shape", o.file, o.sheet, o.row, f"{o.document_no}: no customer"))
        for k in ("net", "tax", "gross"):
            v = getattr(o, k)
            if not isinstance(v, float) or v != v:
                out.append(_v("field-shape", o.file, o.sheet, o.row, f"{o.document_no}: {k} {v!r} is not a number"))
        if o.status_key not in ("open", "to_ship", "in_progress", "completed", "cancelled", "other"):
            out.append(_v("field-shape", o.file, o.sheet, o.row, f"{o.document_no}: status key {o.status_key!r}"))
        if o.planned_ship and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", o.planned_ship):
            out.append(_v("field-shape", o.file, o.sheet, o.row, f"{o.document_no}: planned ship {o.planned_ship!r}"))


def _oracle_self_checks(facts, out):
    """Properties of the DATA that make the oracle's own column choice trustworthy and show a damaged file."""
    for f in facts:
        for name, s in f.sheets.items():
            for ln in s.lines:
                n, t, g = ln.num["net_amount"], ln.num["tax_amount"], ln.num["gross_amount"]
                if None not in (n, t, g) and abs(n + t - g) > Decimal("0.011"):
                    out.append(_v("oracle-arithmetic", f.file, name, ln.row, f"net {n} + tax {t} != gross {g}"))
                if ln.bad_numbers:
                    out.append(_v("oracle-arithmetic", f.file, name, ln.row, f"not a plain number: {ln.bad_numbers}"))
        sheets = list(f.sheets.values())
        if len(sheets) > 1:
            total = max(sheets, key=lambda s: len(s.lines))
            have = {}
            for ln in total.lines:
                have[(ln.text["document_no"],) + oo.line_key(ln)] = have.get((ln.text["document_no"],) + oo.line_key(ln), 0) + 1
            for s in sheets:
                if s is total:
                    continue
                cnt = {}
                for ln in s.lines:
                    k = (ln.text["document_no"],) + oo.line_key(ln)
                    cnt[k] = cnt.get(k, 0) + 1
                for k, n in cnt.items():
                    if n > have.get(k, 0):
                        out.append(_v("cut-sheets-are-subsets", f.file, s.sheet, 0, f"line of {k[0]} is only on this sheet"))


def _idempotent(facts, out):
    paths = [f.path for f in facts]
    a, b = snapshot_orders.records_of(paths), snapshot_orders.records_of(paths)
    if a != b:
        diff = snapshot.diff(a, b)
        out.append(_v("idempotent", "", "", 0, snapshot.report(diff)))


@pytest.fixture(scope="session")
def violations(facts, reads, book):
    out = []
    _file_checks(facts, reads, out)
    _row_checks(facts, reads, out)
    _beyond_cap(facts, out)
    _document_checks(facts, reads, out)
    _book_checks(facts, book, out)
    _shape_checks(book, out)
    _oracle_self_checks(facts, out)
    for f in facts:  # a kept-out file must contribute no order
        if f.kind != "orders" and any(o.file == f.display for o in book.orders if f.display not in {g.display for g in facts if g.kind == "orders"}):
            out.append(_v("kept-out-yields-nothing", f.file, "", 0, "orders in the book come from a kept-out file"))
    _idempotent(facts, out)
    return out


def _param(rule):
    if rule in KNOWN_BUGS:
        return pytest.param(rule, marks=pytest.mark.xfail(strict=True, reason=KNOWN_BUGS[rule]))
    return rule


@pytest.mark.parametrize("rule", [_param(r) for r in ALL_RULES])
def test_rule_has_no_violations(rule, violations):
    bad = [v for v in violations if v["rule"] == rule]
    assert not bad, f"{len(bad)} violation(s) of '{rule}':\n" + "\n".join(
        f"  {v['file'][:40]} {v['sheet'][:20]} row {v['row']}: {v['detail'][:170]}" for v in bad[:6])


def test_every_file_and_row_is_checked(facts, violations):
    assert len(facts) == 9
    kinds = sorted(f.kind for f in facts)
    assert kinds == ["formula_workbook", "invoice_packing", "invoice_packing", "orders", "orders", "orders", "orders",
                     "orders", "stock_list"]
    lines = sum(len(s.lines) for f in facts for s in f.sheets.values())
    assert lines == 375 + 177 + 330 + 76 + 38 + 38  # formula workbook included: its copy is read but kept out


def test_kept_out_files_have_a_stated_reason(facts):
    for f in facts:
        if f.kind != "orders":
            assert f.reason, f.file
