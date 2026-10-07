"""Layer 1c: read the PDFs whose layout never changes (order confirmations, contracts, e-invoices)
and find the reference numbers that tie emails, orders and documents together.

Anything that does not match a known layout is still returned with its reference
numbers and a text excerpt, so the enrichment pass (layer 2) can read it.
"""
import re

REF_PATTERNS = {
    "sales_order": re.compile(r"\bSO\d{7}\b"),
    "quotation": re.compile(r"\bQU\d{9}\b"),
    "order_confirmation": re.compile(r"\bOC\d{8}\b"),
    "purchase_order": re.compile(r"\bPO\d{7}\b"),
    "fibro_order_no": re.compile(r"\b00\d{8}\b"),
    "supplier_contract": re.compile(r"\bOBR\d{7}\b"),
    "drawing": re.compile(r"\b[3-5]-\d{3}-\d{3}-\d{4}\b"),
    "table_model": re.compile(r"\b(?:ER|ES|RT)\.\d{2}\.\d{4}[.\d]*\d\b"),
    "vr_nc": re.compile(r"\bVR\.NC\.?\d{2}(?:[.\d]*\d)?\b"),
}

YEN = "[¥￥]"
MONEY = r"([\d,]+\.\d{2})"


def pdf_text(path: str, max_pages: int = 12) -> str:
    try:
        from pypdf import PdfReader
    except ImportError:  # pragma: no cover
        raise SystemExit("pypdf is not installed: run `.venv/bin/pip install -r requirements.txt` in importer/")
    try:
        reader = PdfReader(path)
        return "\n".join((page.extract_text() or "") for page in reader.pages[:max_pages])
    except Exception:
        return ""


def find_refs(*texts: str) -> dict[str, list[str]]:
    found: dict[str, set] = {}
    blob = "\n".join(t for t in texts if t)
    blob = re.sub(r"\bS0(\d{7})\b", r"SO\1", blob)  # the order-confirmation font prints SO as S0
    for kind, rx in REF_PATTERNS.items():
        hits = set(rx.findall(blob))
        if hits:
            found[kind] = sorted(hits)
    return {k: v for k, v in found.items()}


def _num(s: str) -> float:
    return float(s.replace(",", ""))


def _minor(s: str) -> int:
    return round(_num(s) * 100)


def parse_order_confirmation(text: str) -> dict | None:
    if "Order Confirmation" not in text and "Order NO" not in text:
        return None
    t = re.sub(r"\bS0(\d{7})\b", r"SO\1", text)
    doc = {"type": "order_confirmation", "currency": "CNY"}
    m = re.search(r"Order NO\.?\s*[:：]?\s*(SO\d{7})", t)
    if not m:
        return None
    doc["order_no"] = m.group(1)
    if m := re.search(r"Date\s*[:：]?\s*(\d{4}-\d{2}-\d{2})", t):
        doc["date"] = m.group(1)
    if m := re.search(r"Customer No\.?\s*[:：]?\s*(\S+)", t):
        doc["customer_no"] = m.group(1)
    if m := re.search(r"Contact Person\s*[:：]?\s*([^\n]+?)\s*(?:\n|Tel)", t):
        doc["our_contact"] = m.group(1).strip()
    for k, label in (("net_value_minor", "Net Value"), ("vat_minor", "VAT"), ("total_minor", "Total Value")):
        if m := re.search(rf"{label}\s*[:：]?\s*{MONEY}", t):
            doc[k] = _minor(m.group(1))
    if m := re.search(r"Terms of payment\s*[:：]?\s*([^\n]+)", t):
        doc["payment_terms"] = m.group(1).strip()
    if m := re.search(r"Terms of delivery\s*[:：]?\s*([A-Z]{3})\b", t):
        doc["incoterm"] = m.group(1)
    if m := re.search(r"Currency\s*[:：]?\s*([A-Z]{3})", t):
        doc["currency"] = m.group(1)
    lines = []
    flat = re.sub(r"\s+", " ", t)
    # item rows end with: qty unit-price net-price delivery-date
    for m in re.finditer(r"(\d+)\s+(.+?)\s+(\d+\.\d{2})\s+" + MONEY + r"\s+" + MONEY + r"\s+(\d{4}-\d{2}-\d{2})", flat):
        item, rest, qty, unit, net, delivery = m.groups()
        model = re.search(r"\b(?:ER|ES|RT|VR)\.[\w.]+", rest)
        part = re.search(r"\b\d-\d{3}-\d{3}-[\w]+", rest)
        lines.append({"item": int(item), "text": rest.strip(), "model": model.group(0).rstrip(".") if model else None,
                      "fibro_part": part.group(0) if part else None, "qty": float(qty),
                      "unit_price_minor": _minor(unit), "net_minor": _minor(net), "delivery_date": delivery})
    doc["lines"] = lines
    return doc


def parse_contract(text: str) -> dict | None:
    """FIBRO Shanghai sales contract (合同细则, number SO...)."""
    if "合同" not in text or "Contract" not in text:
        return None
    t = re.sub(r"\bS0(\d{7})\b", r"SO\1", text)
    m = re.search(r"Contract number\s*[:：]\s*(SO\d{7})", t) or re.search(r"合同号\s*[:：]\s*(SO\d{7})", t)
    if not m:
        return None
    doc = {"type": "sales_contract", "contract_no": m.group(1), "currency": "CNY"}
    if m := re.search(r"Date\s+(\d{8})", t):
        d = m.group(1)
        doc["date"] = f"{d[:4]}-{d[4:6]}-{d[6:]}"
    # the English sentence may wrap ('... six\nhundred seventy-five yuan (￥34,675.00)'); the Chinese one does not
    if m := (re.search(rf"本合同价款[^\n]*?{YEN}\s*{MONEY}", t)
             or re.search(rf"Contract value.{{0,250}}?{YEN}\s*{MONEY}", t, re.S)):
        doc["total_minor"] = _minor(m.group(1))
    if m := re.search(r"Party B\s*[:：]\s*([^\n]+)", t):
        doc["counterparty"] = m.group(1).strip()
    if m := re.search(r"交货时间[^\n]*?(\d+)\s*周", t):
        doc["delivery_weeks"] = int(m.group(1))
    if m := re.search(r"warranty period is\s*(\d)\s*years", t, re.I):
        doc["warranty_years"] = int(m.group(1))
    if m := re.search(r"within the (\d+) calendar days", t, re.I):
        doc["payment_days"] = int(m.group(1))
    if m := re.search(r"Net Total Amount\s*" + YEN + r"?\s*" + MONEY, t, re.I):
        doc["net_minor"] = _minor(m.group(1))
    flat = re.sub(r"\s+", " ", t)
    lines = []
    for m in re.finditer(r"(ER|ES|RT|VR)\.([\d.]+\d)\s.*?\s(\d+)\s*" + YEN + r"\s*" + MONEY + r"\s*" + YEN + r"?\s*" + MONEY, flat):
        lines.append({"model": f"{m.group(1)}.{m.group(2)}", "qty": float(m.group(3)),
                      "unit_price_minor": _minor(m.group(4)), "net_minor": _minor(m.group(5))})
    doc["lines"] = lines
    if m := re.search(r"(\d-\d{3}-\d{3}-\d{4})", t):
        doc["drawing"] = m.group(1)
    return doc


def parse_purchase_contract(text: str) -> dict | None:
    """Acme-style 买卖合同 (we are the buyer)."""
    if "买卖合同" not in text:
        return None
    doc = {"type": "purchase_contract", "currency": "CNY"}
    if m := re.search(r"合同编号\s*[:：]\s*([A-Z]+\d+)", text):
        doc["contract_no"] = m.group(1)
    if m := re.search(r"签订时间\s*[:：]\s*(\d{4})年(\d{2})月(\d{2})日", text):
        doc["date"] = "-".join(m.groups())
    if m := re.search(r"合计人民币金额[^\n]*?[¥￥]\s*([\d,]+(?:\.\d{2})?)", text):
        doc["total_minor"] = _minor(m.group(1) if "." in m.group(1) else m.group(1) + ".00")
    if m := re.search(r"交货期限\s*[:：]\s*(\d+)\s*天", text):
        doc["delivery_days"] = int(m.group(1))
    if m := re.search(r"提供(\d+)个月的保修", text):
        doc["warranty_months"] = int(m.group(1))
    if m := re.search(r"结算方式\s*[:：]\s*月结(\d+)天", text):
        doc["payment_days"] = int(m.group(1))
    if m := re.search(r"型号\s*(.+?)\s*单位", text, re.S):
        pass
    if m := re.search(r"\b((?:TR|NC|VR)[.\w-]+)", text):
        doc["model"] = m.group(1)
    return doc


def parse_invoice(text: str) -> dict | None:
    if "电子发票" not in text and "发票号码" not in text:
        return None
    doc = {"type": "invoice", "currency": "CNY"}
    if m := re.search(r"\b(\d{20})\b", text):
        doc["invoice_no"] = m.group(1)
    if m := re.search(r"(\d{4})年(\d{2})月(\d{2})日", text):
        doc["date"] = "-".join(m.groups())
    ids = re.findall(r"\b([0-9A-Z]{18})\b", text)
    names = re.findall(r"名称\s*[:：]\s*([^\n]+?)(?:\s+名称|\n|$)", text)
    if names:
        doc["buyer"] = names[0].strip()
        if len(names) > 1:
            doc["seller"] = names[1].strip()
    if len(ids) >= 2:
        doc["buyer_tax_id"], doc["seller_tax_id"] = ids[0], ids[1]
    if m := re.search(r"[（(]小写[）)]\s*[¥￥]\s*([\d,]+\.\d{2})", text) or re.search(r"[¥￥]\s*([\d,]+\.\d{2})\s*$", text,
                                                                            re.M):
        doc["total_minor"] = _minor(m.group(1))
    amounts = re.findall(r"[¥￥]\s*([\d,]+\.\d{2})", text)
    if len(amounts) >= 3:
        doc["net_minor"], doc["vat_minor"] = _minor(amounts[0]), _minor(amounts[1])
        doc["total_minor"] = _minor(amounts[-1])
    if m := re.search(r"合同编号\s*[:：]\s*([A-Z]+\d+)", text):
        doc["contract_ref"] = m.group(1)
    if m := re.search(r"规格型号\s*\S*\s*([A-Z][\w.\-]+)", text):
        doc["model"] = m.group(1)
    return doc


PARSERS = (parse_order_confirmation, parse_contract, parse_purchase_contract, parse_invoice)


def read_document(path: str) -> dict:
    text = pdf_text(path)
    out = {"text_chars": len(text), "readable": len(text.strip()) > 40, "excerpt": re.sub(r"[ \t]+", " ", text)[:3500],
           "refs": find_refs(text)}
    # amounts typed with Chinese full-width punctuation ('￥34，675.00', '１２．５') read like ASCII ones
    norm = text.translate(str.maketrans("０１２３４５６７８９．", "0123456789."))
    norm = re.sub(r"(?<=\d)[，､](?=\d{3}\b)", ",", norm)
    for parser in PARSERS:
        try:
            doc = parser(norm)
        except Exception as e:  # a layout change must degrade to "unparsed", not abort the run
            doc = None
            out["parse_error"] = repr(e)
        if doc:
            out["parsed"] = doc
            break
    return out
