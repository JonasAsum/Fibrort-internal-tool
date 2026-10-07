# Reading the mails: rules for both readings

You read one case file, `mail/staging/reading/<case>/case.md`, and write what a person understands from
those mails into `mail/staging/reading/<case>/pass_a.json` or `pass_b.json`, whichever you were told.

Two readers read every case independently. Code then checks every quote against the original mail, compares
the two readings, and Jonas reviews money, scan numbers and every disagreement. Only then does anything reach
the CRM, marked as read by AI.

**Do not open the other reader's file.** Your reading is only worth something if it is independent.

## The four rules

1. **Every fact carries a quote.** A quote is copied character for character from the cited mail (M<n>)
   or attachment (D<n>). That means the mail's own text, its quoted history, or its subject. Keep it short:
   the one sentence or table row that proves the fact (at most 300 characters). Never translate, fix or
   shorten inside a quote. A quote that code cannot find in that exact mail rejects the whole item.
2. **Numbers sit inside quotes.** An amount, quantity, order number or date you write must appear in one of
   the item's quotes, in any spelling (28,810.00 or 28810 or ￥28，810.00).
3. **Only names from the case file.** Use only the company names in "Companies you may name" and the people
   and addresses in "People in these mails". Never write a person's or company's name that isn't there.
4. **When unsure, leave it out.** Writing nothing is correct. A guess is wrong even when it turns out to be
   right. Use `"confidence": "low"` only when the quote supports the fact but someone could read it
   differently.

## What to write

### Deals: customer business only

A deal is something a **customer buys or asks to buy** from FIBRO: a sales order (SO…), a quotation (QU…),
or a clear inquiry or order without a number. **FIBRO buying from a supplier is never a deal** (see notes,
`purchase`).

- `key`: the SO or QU number as printed (`SO0000001`). For an inquiry without one, use `null` and set
  `anchor` to the mail where it starts. Code then names it `INQ-<company>-<date>`.
- If the deal is in "Deals code already made", write it **only** to add what code doesn't have (status,
  delivery date, lines from a scan). Use the same key.
- `status`:
  - `won`: the customer ordered. That means a PO from the customer, a signed contract, or FIBRO's order
    confirmation sent to the customer.
  - `lost`: the customer said no, bought elsewhere, or cancelled.
  - `open`: anything else.
  - The quote must show it.
- `amount_minor` (cents: 28,810.00 → 2881000) and `currency`: only if a mail or document states the deal
  total. Write `null` when there's no total; never add lines up yourself.
- `lines`: only as printed (model, text, qty, unit price in minor units).
- `expected_close`: the delivery or order date as stated, `YYYY-MM-DD`.
- `messages`: every mail about this deal (M aliases). These get linked to the deal in the CRM.

### Tasks: something still to do, by FIBRO

A task is a **request to FIBRO, or a promise FIBRO made**, that needed action. Examples: send a quote,
confirm a delivery date, return a signed contract.

- `due`: the date asked for or promised (`YYYY-MM-DD`). Work it out from the mail's date for words like
  "by Friday" or "next week". In that case the due quote is the sentence with those words.
- `done`: `true` only if a **later** mail shows it was done (the quote is from that later mail). Otherwise
  `false`.
- Not a task: newsletters, FYIs, things the customer must do, and anything only discussed.

### Notes: one per real case

A note records what a person would want to know later about a case, in 1–4 plain sentences. Examples: a
warranty case and its outcome, a service visit, an agreement or price list, an exhibition, a complaint.

- `type`: `case`, or `purchase` for FIBRO buying from a supplier. A purchase note says: our PO number, what,
  how many, the confirmed delivery date, delivered or invoiced, and anything open (which then also goes into
  `tasks`).
- `messages`: the mails it is about.
- `company`: only when the note is about another company in the list than the case's (a warranty case of
  Alpha's filed in the Acme folder). Tasks take the same field.
- Not a note: a thread that only moves a deal along (the deal covers it), or routine logistics with nothing
  to remember.
- Never in a note or task: personnel matters, such as someone's performance, pay, holidays or health, or
  quarrels between colleagues. Write the business decision ("Lisa handles sales, Amber payments"), never the
  argument behind it.

### Routing: mails filed under the wrong company, or under none

- `m`: the mail. `company`: the company it is really about, which must be in the list. `deal`: an SO/QU key
  if it is about one deal.
- The quote must show the tie: their name, their order number, or their person writing.
- In the `_unfiled-*` cases this is the main job: route every mail that is clearly about one customer or
  supplier. Internal FIBRO mails that are about no outside company stay unrouted.

### Contact titles

- Only when a mail states someone's job title and the case file lists none for them. The quote must contain
  the title exactly as you write it.

### Documents: what a scan says

- For each attachment marked SCAN that matters, open the file (`Read` it, it's a PDF) and write what is
  printed: `type` (order_confirmation, purchase_order, invoice, proforma, delivery_note, awb, service_report,
  agreement, price_list, other), `number`, `date`, `total_minor`, `currency`, and `summary` (one sentence).
- The quote is the line exactly as printed on the scan.
- Code can't check scan quotes, so Jonas checks every scan number. Be exact.

## Output: one JSON object

```json
{
  "case": "<case name from the heading>",
  "reader": "pass A",
  "deals": [{
    "key": "SO0000001", "anchor": null, "status": "won",
    "amount_minor": 3659400, "currency": "CNY", "expected_close": "2026-05-15",
    "lines": [{"model": "ER.11.0160.1.161.12.0.3.1", "text": "ER.11 dummy", "qty": 1, "unit_price_minor": 1829700}],
    "messages": ["M12", "M13"],
    "summary": "Acme ordered two ER.11 tables; delivery asked for in 4-6 weeks.",
    "evidence": {
      "status": [{"m": "M12", "q": "附件是订单，请确认"}],
      "amount": [{"m": "D7", "q": "合计 36,594.00"}],
      "expected_close": [{"m": "M12", "q": "交货期 2026-05-15"}],
      "lines": [{"m": "D7", "q": "ER.11.0160.1.161.12.0.3.1 1 18,297.00"}]
    },
    "confidence": "high"
  }],
  "tasks": [{
    "subject": "Confirm 4-6 week delivery for ER.11 (SO0000001)",
    "body": "Carrie asked whether both ER.11 tables can ship in 4-6 weeks.",
    "due": "2026-04-08", "done": true,
    "evidence": {
      "task": [{"m": "M14", "q": "Can we deliver in 4-6 weeks?"}],
      "due": [{"m": "M14", "q": "please confirm by tomorrow"}],
      "done": [{"m": "M15", "q": "4-6 weeks is OK"}]
    },
    "confidence": "high"
  }],
  "notes": [{
    "type": "case", "subject": "ER17 housing not accepted",
    "body": "Acme refused the ER17 housing from OC00000001 because ...",
    "messages": ["M3", "M4"],
    "evidence": {"note": [{"m": "M3", "q": "the housing is not acceptable"}]}
  }],
  "routing": [{"m": "M40", "company": "Acme", "deal": null, "evidence": [{"m": "M40", "q": "示例"}]}],
  "contacts": [{"email": "sales@acme.com", "title": "Sales", "evidence": [{"m": "M9", "q": "Mary 王芳 Sales"}]}],
  "documents": [{
    "d": "D7", "type": "order_confirmation", "number": "SO0000001", "date": "2026-04-03",
    "total_minor": 3659400, "currency": "CNY", "summary": "Order confirmation for two ER.11 tables.",
    "evidence": [{"m": "D7", "q": "合计 36,594.00"}]
  }]
}
```

Leave out any list you have nothing for, or write `[]`. Write valid JSON only, with no comments. When done, run
`.venv/bin/python -m mail.reading --company "<case>"` from `importer/` to see what the check says about your
reading.
