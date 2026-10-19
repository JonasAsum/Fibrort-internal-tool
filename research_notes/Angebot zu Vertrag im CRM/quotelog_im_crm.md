# Quote log in Excel: moving it into a CRM and reporting on it (B2B machine accessories, 4-7 users, CNY/EUR/USD)

Method note: findings come from web-search result summaries and vendor documentation (about 8 searches, no full-page fetches). Vendor docs describe their own product patterns; they are not independent comparisons. Items marked Inference are my reasoning, not sourced. Several requested topics (standard field list of a quote log, import of history, quote-to-order conversion report) had no direct sources and are listed under Gaps.

## 1. Deal with stage versus separate quote object: when is each right for a small team?

### Takeaway
Vendors split into two camps: the quote is a stage on the deal (lighter) or a separate record with its own lifecycle linked to the deal (needed for versions, approvals, quote-level reporting). No source compares them directly; for a team with few, high-value quotes, the evidence supports "deal carries the stage and the value, quote versions are optional child records or simply a version field".

### Cited Findings
- Bitrix24 treats a quote as a decision point inside the deal, built from the deal record, with its own statuses (e.g. awaiting approval, delivered to customer); it separates preparing from approving a quote as a control point — [Bitrix24](https://www.bitrix24.com/tools/crm/features/quote-management.php)
- HelloLeads lets quotes move through configurable stages such as Draft, Sent, Accepted — [HelloLeads](https://helloleads.io/learnmore-web/quote-management)
- Vtiger keeps a quote lifecycle separate from the deal (Draft, Reviewed, Delivered, Accepted, Rejected); the deal advances only after the customer accepts — [Vtiger](https://help.vtiger.com/article/117249802-What-are-Quotes)
- Vedain (a vendor blog) argues the lighter route: the rep moves the deal to a "Proposal Sent" stage and generates the quote from the same screen — [Vedain](https://vedain.com/blog/10-best-quotation-management-software-tools-in-2026)
- Zoho ties deal stages to a probability and an amount; these drive expected-income and pipeline-stage reports — [Zenatta (Zoho overview)](https://zenatta.com/deals-and-quotes-to-close-sales/)
- Salesforce forecast reports are driven by the Opportunity amount; with a quote object, syncing the primary quote to the opportunity is what keeps pipeline reports accurate (Conga CPQ documentation, not native Salesforce) — [Conga](https://documentation.conga.com/cpq-for-salesforce/current/cpq-for-users/finalizing-quotes/synchronize-quotes-with-opportunities/syncing-quotes)

### Inferences
- If the Excel log has one row per quote and the reports are about quotes (value, age, follow-up), the simplest faithful model is: one deal per enquiry, stage = quote lifecycle. Add a separate quote object only if you need generated PDFs, line items, approval, or several simultaneous offers for one enquiry.
- With 4-7 users and few, high-value quotes, a separate quote object adds maintenance (two statuses to keep in sync, risk that deal amount and quote amount diverge) without a reporting gain.
- If the CRM is Odoo-style (quotation is the sales-order document), the quotation is itself the quote record and the opportunity above it holds the stage; then use both but report on the opportunity.

### Gaps
- No independent source comparing deal-only versus quote-object setups for teams under 10 users.
- No official Odoo, Pipedrive or Zoho documentation was retrieved on this question.

## 2. Minimal fields and statuses to replace an Excel quote log without losing information

### Takeaway
No source publishes a canonical quote-log field list. The cited material supports the following building blocks: owner, amount with currency, stage/status with probability, close date, lost reason, linked company; the rest of the list below is my proposal.

### Cited Findings
- Deals carry amount and a probability tied to the current stage; Zoho uses these for expected income — [Zenatta](https://zenatta.com/deals-and-quotes-to-close-sales/)
- Lost reasons are a standard CRM concept; in Odoo they exist for opportunities, while standard Odoo allows cancelling a sales order without explanation, so third-party modules add a mandatory reason — [Odoo App Store: sale_lost_reason](https://apps.odoo.com/apps/modules/18.0/sale_lost_reason)
- Quote statuses used by vendors: Draft, Sent, Accepted (HelloLeads); Draft, Reviewed, Delivered, Accepted, Rejected (Vtiger) — see links in section 1
- Deal records link to a company and one or more people; companies are imported first, then people, then deals — [Pipeline CRM](https://help.pipelinecrm.com/article/how-to-import-your-csv-or-excel-file)
- HubSpot sums stage totals on "Amount in company currency", not the raw deal amount — [HubSpot community](https://community.hubspot.com/t/currencies-not-lining-up/143639)

### Inferences
Proposed minimal field set (map every Excel column; nothing silently dropped):
- Quote number (original Excel reference or number from the ERP) — natural unique key for import and duplicate detection
- Quote date (sent date), customer (linked company), contact, rep (owner), product / product group (picklist plus free text), value, currency, status/stage, expected decision date, next follow-up date and note, lost reason, order number / PO date once won, version/revision number, source Excel row or file name (audit trail), free-text comment.
Proposed stage set for few, high-value B2B quotes: 1 Enquiry / clarifying, 2 Quote in preparation, 3 Quote sent, 4 Negotiation / technical clarification, 5 Verbal commitment / awaiting PO, then Won (order received), Lost (with reason: price, competitor, no budget, project cancelled, no response), On hold. Define an exit criterion per stage (for example "Quote sent" = document delivered to customer with date logged; "Awaiting PO" = customer confirmed intent in writing) and set probabilities only once there are some closed deals to calibrate against.
- Keep "no response / expired" as a Lost reason rather than a stage, so open pipeline stays clean.

### Gaps
- No sourced standard field list or stage definitions specific to industrial machine accessories; the above is proposed, not benchmarked.

## 3. Reports for the managing director

### Takeaway
Sources agree on a core set: open pipeline by stage, stale deals by owner, deal age versus stage, win rate (two variants), stage conversion. Forecast and quote-to-order conversion follow from stage probability and won/lost data.

### Cited Findings
- Stale-deal report: open opportunities with last activity older than a threshold (often 14 days, some use 21+), grouped by owner and sorted by amount; described as the most effective pipeline-hygiene report — [Prospeo pipeline report template](https://prospeo.io/s/pipeline-report-template), [Outreach on pipeline ageing](https://www.outreach.ai/resources/blog/sales-pipeline-ageing)
- Deal age by stage: set an expected dwell time per stage and flag outliers (example: normal 12 days, flagged at 30); late-stage aging distorts forecasts; thresholds should be set per segment, not company-wide — [Outreach](https://www.outreach.ai/resources/blog/sales-pipeline-ageing)
- Two win-rate versions: pipeline win rate (open deals in denominator) and closed win rate (excludes open); a large gap points to stale pipeline, long cycles or poor qualification; rep-level via closed summary by owner — [Prospeo](https://prospeo.io/s/sales-pipeline-reporting)
- Stage conversion = deals entering next stage / deals entering current stage — [Prospeo](https://prospeo.io/s/sales-pipeline-reporting)
- Pipeline value by stage with close-date grouping — [VortexIQ HubSpot KPI doc](https://docs.vortexiq.ai/nerve-centre/kpi-cards/hubspot/pipeline-value-by-stage)
- Forecast in a multi-currency HubSpot account used company currency; a community workaround weights the amount with a stage-probability property (older, unofficial) — [HubSpot community](https://community.hubspot.com/t5/CRM/Custom-Deal-Revenue-Forecast-by-stages/m-p/433765)

### Inferences
Suggested set of 7 (replicates the Excel overview):
1. Open quotes by rep, with count, total value in one reporting currency, and age (days since quote date), sortable.
2. Quote value by stage (funnel / table), in reporting currency.
3. Quotes without follow-up: open deals with no next activity scheduled or last activity older than 14 days, by rep.
4. Win rate by rep and by product group, closed basis and by value as well as by count (with few high-value quotes, count-only rates are noisy).
5. Quote-to-order conversion: quotes sent in a period versus those that became orders, cohorted by quote month (cycles are long, so look at cohorts, not same-month ratios).
6. Weighted forecast by expected decision month: value x stage probability, plus an unweighted "commit" view for deals at awaiting-PO.
7. Lost analysis: lost value by reason and by competitor.
With few deals, a weekly list review (open quotes by age, oldest first) is likely more useful than percentage charts.

### Gaps
- No source defines a "quote-to-order conversion" report; it is derived from won/total sent. No source on the dashboards a managing director of a small industrial firm actually uses.

## 4. Revisions and re-quotes of the same enquiry; import of history; common pitfalls

### Takeaway
Rule from sources: revisions are versions of one quote on one deal; competing options are separate quotes on one deal; independent projects are separate deals. For import: companies first, then contacts, then deals, using a unique key to avoid duplicates.

### Cited Findings
Revisions
- PSO Hub: copying a quote under the existing number creates the next version; the old one is marked declined, the new starts as draft, both stay linked to the same deal — [PSO Hub](https://help.psohub.com/copy-a-quote)
- Maple CRM: keep a full version history and separate draft, sent and approved versions so reps know which one the prospect saw — [Maple CRM](https://maplecrm.com/blog/wp-json/wp/v2/posts/2698)
- Pipeliner: several distinct quotes may be linked to one Opportunity (different bundles); administrators can lock quotes once sent — [Pipeliner](https://help.pipelinersales.com/en/articles/6912065-quotes-creating-and-managing-quotes)
- Oracle CPQ: multiple versions per quote, but only one active quote per opportunity; accepted changes flow back to the opportunity — [Oracle](https://docs.oracle.com/en/cloud/saas/sales/fastg/create-and-update-a-quote-for-an-opportunity.html)
- An SAP support note describes version 2 not being linked to the same opportunity; check this after creating a version — [SAP](https://userapps.support.sap.com/sap/support/knowledge/en/2772844)
- Salesforce user report: several quotes under one opportunity, the synced (primary) quote sets the opportunity amount; on win the non-primary quotes are cancelled (user report, not official) — [Zoho community thread](https://help.zoho.com/portal/fr/community/topic/how-do-i-synchronize-a-quote-to-an-opportunity?page=7)
- Odoo third-party modules: cancelling a quotation can mark the lead lost and cancel related quotations; a lead with a confirmed sales order cannot be marked lost — [Odoo App Store](https://apps.odoo.com/apps/modules/19.0/md_lead_sale_cancel_reason)

Import and duplicates
- Import order companies, people, deals; one record type per import; company name must be unique in Pipeline CRM and a repeated company in one file imports only once — [Pipeline CRM](https://help.pipelinecrm.com/article/how-to-import-your-csv-or-excel-file), [Pipeline CRM limits](https://help.pipelinecrm.com/article/import-limits-and-what-you-can-t-import)
- Insycle matches companies by sales-ledger ID or name plus postcode — [Insycle](https://support.insycle.com/hc/en-us/articles/6586816077847)
- Freshworks: phone duplicate detection is exact only (with and without country code do not match); has an import history page with created/updated/skipped/error counts — [Freshworks](https://crmsupport.freshworks.com/support/solutions/articles/50000009162-how-to-import-contacts-from-a-csv-xlsx-file-)
- Pipeline CRM: imported activities take the import run date, not the original date, which matters for history migration — [Pipeline CRM](https://help.pipelinecrm.com/article/how-to-import-your-csv-or-excel-file)
- Pre-import cleanup: clear headers, no merged cells, standardised dates and phone numbers, no duplicate rows — [Vedain import guide](https://vedain.com/blog/how-to-import-your-contacts-from-excel-into-a-crm-without-losing-data)

Currency
- HubSpot: open deals use the latest rate set for the currency; closed deals keep the rate at close; pipeline totals sum Amount in company currency; multi-currency needs a paid tier — [HubSpot KB](https://knowledge.hubspot.com/account-management/add-and-edit-your-account-currencies)
- Pipedrive: report conversion via Open Exchange Rates, rates for statistical purposes only; a representative confirmed conversion of already-won deals keeps updating with exchange rates, so closed totals can drift; custom currencies behaved inconsistently in a 2024 thread (contradiction between help article and staff reply) — [Pipedrive help](https://support.pipedrive.com/article/how-is-currency-converted-in-pipedrive), [Pipedrive community](https://community.pipedrive.com/discussion/comment/11133)

### Inferences
- Model: one deal per enquiry; each Excel row with the same customer + product + same reference root (e.g. 2025-114, 2025-114-B) becomes a revision of one deal, not a new deal. Report on the latest version only (or the one marked active) to avoid double-counting value; keep older versions as history (child record, or a note/field with the superseded value and date). Lock the sent version.
- Duplicate quotes: use the Excel quote number as unique key in the import, and flag rows with same customer + product + value + date within a few days for manual review.
- Quotes without customer match: do not create ad-hoc companies. Import into a staging sheet, match on name normalised (case, legal suffix, Chinese/English name) plus city or tax ID, queue the unmatched for a human decision, then import. Keep the original customer text in a field either way.
- Import history by writing the original quote date into a custom "quote date" field and the original owner into the owner field; a created-date field set by import will be wrong in most CRMs, so do not use it for age reports. Log the original date in a note if the CRM cannot set it.
- Currency: store original amount and ISO currency per deal and never overwrite it; pick one reporting currency (for example EUR); freeze the rate for won/lost deals at the close date (HubSpot does this by default; Pipedrive apparently does not), use one rate table for open deals, and document that numbers are statistical. Check that the chosen CRM's multi-currency forecast works in the reporting currency before committing, since forecast tooling has been reported to ignore regional currencies.
- Few, high-value quotes make single wrong amounts or currencies material: add an import check that sums value per currency and compares it with the Excel totals.

### Gaps
- No source on migrating full quote history (activities, attachments) beyond date handling; no multi-currency import guidance (mapping ISO codes, rate handling) was found.
- No official documentation on Odoo quotation versions or Salesforce quote-to-order conversion was retrieved.
- Chinese-language CRM options (e.g. for CNY) and CRM-specific limits for a 4-7 user plan were not researched.
