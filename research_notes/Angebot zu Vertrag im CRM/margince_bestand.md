# Margince CRM: what it can already do for quote -> sales contract -> quote-log report (repository inventory)

Base path for all sources: `/Users/jonas/Mac Documents/Shanghai Company/CRM System/margince/` (abbreviated `M/`). Read-only review; nothing was run. Items marked INFERENCE are mine, not stated in the repo. "Contract" in Margince means a recorded agreement, not a generated sales-contract document (see Q2).

## Q1. Which entities exist and how are they linked? (quote/offer, line item, product, deal, contract, document, task, activity)

### Takeaway
Margince has a full, implemented quote object called **Offer** ("Angebot"), bound to a Deal, with line items, a product rate card, revisions, PDF render and a status lifecycle (draft, sent, accepted, rejected, expired, superseded). A separate **Contract** record exists (bound optionally to deal/company/project) but nothing links the accepted offer to a contract automatically. There is no workspace-wide list of offers (only per deal).

### Cited Findings
- Offer is a first-class object: `Offer` schema with `deal_id`, server-minted `offer_number`, `revision`, `status` enum `[draft, sent, accepted, rejected, expired, superseded]`, `currency`, `buyer_company_id`, `valid_until`, `intro_text`, `terms_text`, derived `net/tax/gross_minor`, `arr_minor`, `net_tcv_minor`, `fx_rate_to_base` (frozen at send), `template_id`, `pdf_asset_ref`, `accepted_at`, `line_items[]`, `source`, `captured_by`. — `M/backend/api/crm.yaml` (schema `Offer`, ~line 51499)
- Offer endpoints: `POST/GET /deals/{id}/offers`, `GET/PATCH/DELETE /offers/{id}`, `POST /offers/{id}/line-items`, `PATCH/DELETE /offers/{id}/line-items/{lineItemId}`, `/offers/{id}/send|accept|reject|regenerate|render`, `GET /offers/{id}/pdf`. — `M/backend/api/crm.yaml` lines ~22068-22440
- No workspace-wide offer list exists: the only list operation is `listDealOffers` under a deal; no `GET /offers`. — grep of `M/backend/api/crm.yaml` (no `/offers:` path, no `listOffers`)
- Offer number format is a simple installation-wide counter, `A-<n>`. — `M/backend/internal/modules/deals/offer.go` (`nextOfferNumber`, line ~205)
- Totals are server-computed in integer minor units, a client-supplied total is rejected 422; offer editable only in `draft`; `sent` is immutable, `regenerate` mints revision N+1 and marks the previous `superseded`. — `M/backend/api/crm.yaml` (comment block above `/deals/{id}/offers`, `regenerateOffer`), `M/docs/handbook/offers-and-products.md`
- Line item fields: position, `product_id`, description, unit, quantity, `unit_price_minor`, `billing_model` (one_time/recurring), `billing_interval_months` (1/3/6/12), `interval_count`, `discount_pct`, `tax_rate`. With a `product_id` the line SNAPSHOTS description/unit/price/tax; later product edits never touch the line. — `M/backend/api/crm.yaml` schema `OfferLineItemInput` (~line 51562)
- Product / price list: `Product` = "rate-card entry ... data an offer line snapshots from, never a configurator"; fields name, sku, description, unit, `unit_price_minor`, billing model/interval, currency, `default_tax_rate`, active. CRUD at `/products`, `/products/{id}`. UI at Settings -> Products and offers. No price tiers, volume discounts, bundles or multiple price lists in the schema. — `M/backend/api/crm.yaml` schema `Product` (~line 50774); `M/docs/handbook/offers-and-products.md`
- Deal is the parent: `/deals`, `/deals/{id}`, `/deals/{id}/advance` (stage move), `/deals/{id}/status`, `/deals/{id}/stakeholders`, `/deals/{id}/documents`. Pipelines/stages are configurable (`/pipelines`, `/stages`, stage types Open/Won/Lost, win probability, exit criteria `/stages/{id}/exit-criteria`). — `M/backend/api/crm.yaml` path list; `M/docs/handbook/the-pipeline.md`
- Accepting an offer (human-only) sets `accepted_at`, syncs `deal.amount_minor/currency` from the offer gross, emits `offer.accepted`. It does NOT move the deal stage and does NOT create a contract (stage movement and won-closing are separate user actions). — `M/backend/api/crm.yaml` `acceptOffer` (~line 22260); `M/backend/internal/modules/deals/offer_dealsync.go`; `M/docs/handbook/offers-and-products.md` ("The deal value and currency are updated")
- Contract entity: fields company_id (required counterparty), optional `deal_id`, `project_id`, `contract_number` (free text), title, `value_minor`, currency, `arr_minor`, `value_basis` (total | annualized_12m), starts/ends/renewal dates, `auto_renew`, `notice_period_days`, `payment_term_days`, `status` (draft, active, expired, cancelled, superseded; changed only via `changeContractStatus`), `signed_on`, cancellation dates, `superseded_by_id`, derived `under_contract`. Endpoints: `/contracts`, `/contracts/{id}`, `/contracts/{id}/status|cancellation|renewal`, `/companies/{id}/contracts`. — `M/backend/api/crm.yaml` schema `Contract` (~line 38061), paths ~25471-25650
- The Contract schema has no link to an offer (no `offer_id` field in the schema as read); the only upstream link is the optional `deal_id`. — `M/backend/api/crm.yaml` schema `Contract` (INFERENCE from absence in the read portion; verify with a full field list before relying on it)
- Contract creation is human-only and cookie-session-only: `createContract` has `security: [cookieAuth]`, `x-agent-access: human-only`, "the assistant reads contracts and never writes one". — `M/backend/api/crm.yaml` lines 25471-25483
- Documents: file attachments on company, contact, deal (`/attachments`, `/deals/{id}/documents`, `/companies/{id}/documents`) with categories Contract, Offer, Legal, Other; 25 MB default upload cap (1-100 MB configurable). Deal Rooms share deal files with external participants. — `M/docs/handbook/documents-and-files.md`, `M/docs/handbook/deal-rooms.md`, `M/backend/api/crm.yaml` `/deal-rooms*`
- Other entities present: tasks (`/tasks`), activities (`/activities`, emails, meetings), projects (`/projects`), leads (`/leads`), tags, lists, teams, commissions (`/commissions`), partners, invoices as a read-only mirror from accounting (`/companies/{id}/finance-summary`). Invoices are "never entered in Margince". — `M/backend/api/crm.yaml` paths; `M/docs/handbook/contracts-and-invoices.md`
- Link summary (from schemas): Company 1-n Deal; Deal 1-n Offer (revisions); Offer 1-n OfferLineItem, each optionally -> Product; Offer -> OfferTemplate (layout/locale); Company 1-n Contract; Contract -> Deal (optional), Project (optional).

### Inferences
- INFERENCE: Because the contract has `deal_id` and the offer has `deal_id`, a quote-to-contract link could be reconstructed through the deal, but the repo defines no automatic or explicit offer-to-contract relation.
- INFERENCE: A quote log needs a cross-deal list of offers; with only `GET /deals/{id}/offers`, a log would need either a new endpoint/report or reading offers deal by deal.

### Gaps
- I did not read the full `CreateContractRequest` / DB `contract` table definition field by field; the absence of an `offer_id` is based on the `Contract` response schema only.
- Exact column list of the `offer` table not read in the baseline migration (`M/backend/migrations/core/0001_baseline.up.sql`).

## Q2. Is there document / contract generation (templates, merge fields, PDF, docx) or only file attachment?

### Takeaway
Only the **offer** has generation: a built-in branded A4 PDF (Go library fpdf) with header/footer text per template, DE/EN labels. There is no sales-contract / order-confirmation generator, no docx, no merge-field engine, and no e-signature. The offer PDF uses the Latin-only Helvetica font, so Chinese text will not render (relevant for a China-based rep).

### Cited Findings
- Offer PDF renderer: A4 built with `github.com/go-pdf/fpdf`, prints title, buyer, line items, totals, the offer's own intro and terms, and the template's header/footer text. — `M/backend/internal/modules/deals/offer_pdf.go` (header comment, imports)
- Font limit: core Helvetica with cp1252 translation; the code states "a non-Latin locale (e.g. Cyrillic, CJK) would need an embedded TTF font instead, not attempted here since only DE/EN ship at launch." — `M/backend/internal/modules/deals/offer_pdf.go` lines ~321-326
- Template = `OfferTemplate`: name, locale (de-DE / en-US), `is_default` (max one default per locale), `layout` (loose JSON). Renderer honours only the layout keys `header` and `footer`; a logo reference "is not fetched or embedded"; "terms always come from the offer, never from the template". — `M/backend/api/crm.yaml` schema `OfferTemplate`; `M/backend/internal/modules/deals/offer_pdf.go` (`OfferTemplateLayoutTextKeys`); `M/docs/handbook/offers-and-products.md` ("Two honest limits")
- Known template gap: an offer whose template was deleted renders from an empty layout rather than the locale default; revision regeneration does not carry the template over. — `M/docs/handbook/offers-and-products.md`
- Render needs the object store: `POST /offers/{id}/render` answers 501 "not implemented" when no blobstore is wired; stored ref `pdf_asset_ref`; `GET /offers/{id}/pdf` downloads it. Render and download are human-only. — `M/backend/api/crm.yaml` lines ~22342-22440
- "Send does not email anything": `sendOffer` is only a status freeze (FX rate + buyer/issuer snapshot); "delivery to a counterparty is a separate capability that does not exist yet". The rep sends the PDF by hand. — `M/docs/handbook/offers-and-products.md`; `M/docs/reference/agent-tools.md` lines ~257-258; `M/backend/api/crm.yaml` `sendOffer` description
- Buyer cannot accept in the system: "the buyer cannot accept or reject it inside Margince"; the rep records the answer. Deal Rooms exist for sharing documents with external participants (comment threads) but are not an acceptance flow. — `M/docs/handbook/offers-and-products.md`; `M/docs/handbook/deal-rooms.md`
- E-signature: none. The contract schema says "never a signing ceremony — in-product e-signature stays removed (A94)"; `signed_on` is asserted by a human. DocuSign mails are only classified as infrastructure mail by the capture gate. — `M/backend/api/crm.yaml` `Contract.signed_on`; `M/docs/explanation/ingress-gate-and-auto-capture.md` lines ~181, 513
- Contract is a data record plus an uploaded signed file ("Add contract" ... "drop the signed file on Signed document"); contracts "are entered by hand, one by one". No template, no generated contract text. — `M/docs/handbook/contracts-and-invoices.md`
- No match for docx/merge-field/mail-merge generation of documents found in docs or API paths (path list of `crm.yaml` has no document-generation endpoints other than `/offers/{id}/render`). — `M/backend/api/crm.yaml` path inventory
- AI drafting exists for offers: AI task `offer_draft` ("Drafts offers and quotes from a deal's details", status shipped, interactive, ladder cheap_cloud/premium) and `regenerateOffer` can be AI-drafted from captured signal, returning `ai_generated`, `ai_disclosure`, `diff_from_previous`. — `M/backend/api/ai-tasks.yaml` (`offer_draft`, ~line 390); `M/backend/api/crm.yaml` `Offer` and `regenerateOffer`. Task `document_extract` reads an attached document (invoice, order confirmation) for deal facts a human may accept. — `M/backend/api/ai-tasks.yaml`
- No AI task or tool for drafting a contract / order confirmation was found (ai-tasks has `offer_draft` only). — `M/backend/api/ai-tasks.yaml` (grep: offer; contract drafting absent)

### Inferences
- INFERENCE: The existing offer render path (template + server-computed lines + PDF) is the closest reusable pattern for an order-confirmation PDF, but today it is offer-specific; a sales contract would need new code (or an external document tool fed by webhook/API).
- INFERENCE: For CN/EN/DE bilingual documents the PDF renderer would need an embedded TTF with CJK glyphs (code change in `offer_pdf.go`).

### Gaps
- Whether the frontend offer template form exposes more than header/footer: `M/frontend/src/screens/offertemplates.tsx` not read in detail (the gate test `offertemplatelayout_test.go` ties the form keys to the two renderer keys).
- Language set is DE/EN only ("DE/EN at launch"); whether `zh-CN` UI locale exists was not checked.

## Q3. Which reports or dashboards exist, and can the user define their own?

### Takeaway
Margince has a real reporting layer (saved, shareable, schedulable reports; targets; forecast; CSV export; a governed free-form analytics query). But the data it reports on is deals, leads and activities; **offers are not a reportable or filterable population**, so "open quotes by rep" or a quote log report is not available out of the box.

### Cited Findings
- Saved reports: `/analytics/reports` (create with name, audience private/team/workspace, selection), editions (snapshots), schedules (weekly/monthly, "not sent by email"), targets, evaluation, CSV export (`/analytics/editions/{id}/export.csv`, `/analytics/evaluate.csv`). Creating a report is human-only. — `M/backend/api/crm.yaml` (~lines 11074-11450); `M/docs/handbook/sales-reporting.md`
- User-selectable pieces are closed lists: metrics `bookings_won, closed_win_rate, open_pipeline, stage_age, qualified_pipeline_created, meetings_held, accepted_opportunities, forecast_landing`; blocks `bookings_trend, stage_distribution, owner_attainment, stage_age, sdr_outcomes, target_progress, forecast_support, pipeline_movement, metric_reading`; period this/last month, last week, this quarter, custom. — `M/backend/api/crm.yaml` schemas `ReportingMetricID`, `ReportingBlockKind`, `ReportingSelection`
- Win rate and value by stage exist (closed_win_rate, open_pipeline, stage_distribution). Forecast screens exist (`/forecast*`, `frontend/src/screens/analytics.forecast*.tsx`). — same schemas; `M/frontend/src/screens/`
- Free-form analytics: `POST /analytics/query` ("Answer a question nobody wrote a report for") with entity, group_by, measures, filters; `GET /analytics/schema` lists what the caller may ask. The entities available come from the hard-coded `prebuiltReports` map: `open-deals-per-company`, `stage-age`, `deals-by-stage` (entity deal), `leads-by-status` (lead), `activities-by-kind` (activity). No offer or contract entity. — `M/backend/internal/compose/reportspecs.go` (lines 63-259); `M/backend/internal/compose/analyticsqueryseam.go`
- Report/filter/list/export/saved-view object enums exclude offer and contract: filtered export `object` enum `[contact, company, deal, lead, project]`; saved lists/views record types `[deal, contact, company, lead, project]`; custom-field objects `[contact, company, deal, lead, project, contract]` (no offer). — `M/backend/api/crm.yaml` (export request ~11804; lines 19462, 49482-49515; `CreateCustomFieldRequest` ~46978)
- Saved views are per-user (columns, sort, filter), lists can be live (filter-driven) or shortlists. — `M/backend/api/crm.yaml` `/views`, `/lists`
- Offer-related numbers do exist per record: `arr_minor`, `net_tcv_minor` on each offer, and accepted offer feeds `deal.amount_minor`. So a deal-level "value by stage" reflects accepted offers, but not "open quotes" (sent, not yet accepted) as a population. — `M/backend/api/crm.yaml` `Offer`; `M/backend/internal/modules/deals/offer_dealsync.go`
- No dashboard builder / widget designer found; reporting is a fixed set of blocks plus the governed query. "Custom reports" screen and "your connected AI ... governed reporting tools" are the stated route for bespoke analysis. — `M/docs/handbook/sales-reporting.md` (line ~89)

### Inferences
- INFERENCE: A QuoteLog report would need either (a) a new `offer` analytics entity in `reportspecs.go` (core code change), or (b) an extension, or (c) building the log outside (webhook -> sheet/BI) and keeping only deal-level reporting in Margince.
- INFERENCE: Metric "accepted_opportunities" and "bookings_won" are deal-based; whether accepted offers count for them was not verified.

### Gaps
- Did not verify whether `prebuiltReports` can be extended by an extension unit (the extension capability list in Q4 has no report/analytics field, so probably not).
- `M/docs/handbook/analytics.md` ("six report sections") not read in full.

## Q4. Which extension points exist for building quote-to-contract without changing the core? (extensions, webhooks, automations, custom fields, MCP/REST, import)

### Takeaway
Extension points are real but each has a hard limit for this use case: webhooks and event subscriptions can **react** to `offer.accepted`, but the core's write door for extensions only offers `Activities().Create`, and contract creation over REST/MCP is human-only. So fully automatic contract creation inside Margince needs a core change (or a change to the human-only rule); an external generator fed by webhook is feasible.

### Cited Findings
- Philosophy: "You change it by changing the code. No config screens, no metadata engine ... a custom field or a workflow is a normal code change in your own copy" (README); runtime plugins explicitly rejected, extensions are compile-time Go units. — `M/README.md`; `M/docs/explanation/extensibility.md`
- Extension tier: `extensions/<name>/` own Go module; capabilities `Jurisdictions, Tools, Channels, Ingress, Subscriptions, Jobs, Secrets, Migrations`; own tables `ext_<name>_*`, own REST namespace `/v1/ext/<name>/`, own screen; shipped units are `extensions/de` (retention floors) and `extensions/openchannel` (reference unit). No extension for quotes/contracts/CPQ in the tree. — `M/docs/explanation/extensibility.md`; `M/extensions/` (dirs `de`, `openchannel`)
- Subscription can listen to event types (e.g. `offer.accepted`) but "a delivery has nobody behind it. The caller is the zero Caller, so `tx.Core()` refuses. Your own tables stay writable". — `M/docs/how-to/add-an-extension.md` (~lines 516-535)
- The core port for extensions is `type Core interface { Activities() ActivityRepo }` with only `Create(CreateActivityRequest)`. No contract/offer/deal write port. — `M/backend/pkg/extension/core.go`
- Outbound webhooks (implemented): subscriptions at `/webhook-subscriptions` (+ deliveries, replay, rotate-secret), Standard Webhooks signing, retries with backoff, dead-letter, owner-scope visibility gate; event catalog includes `offer.created|sent|accepted|rejected|superseded`, `deal.created|updated|stage_changed|...`, `contract.created|updated|status_changed|archived`. `offer.accepted` payload: `offer_id, deal_id, revision, gross_minor` (no line items; consumer must call REST for details). Outbound only. — `M/docs/explanation/outbound-webhooks.md`; `M/backend/api/public-events.yaml` (lines ~31-70, ~1104-1125)
- Automations: closed catalog. Triggers (8): record_created_updated, field_reaches_value, deal_enters_leaves_stage, no_activity_for_n_days, date_field_approaching, inbound_reply, task_overdue, list_membership_changed. Actions (7): create_task, notify, assign_owner, set_field, draft_email, request_approval, add_to_shortlist. "There is no rule builder, no expression language, and no user-defined trigger or action"; new members are a code change. No action creates a contract/offer or generates a document. Stage automation policies also exist (`/stage-automation/*`). (Doc text says "seven triggers and six actions"; code lists eight and seven, so the doc lags the code.) — `M/docs/explanation/automation.md`; `M/backend/internal/modules/automation/catalog_triggers.go`, `catalog_actions.go`
- Custom fields: runtime, typed real columns (`cf_<slug>`), types text, number, date, currency, picklist, multiselect (API), boolean, on objects contact, company, deal, lead, project, contract (API; the explanation doc still says five objects/six types). Not on offer. Admin-defined at `/custom-fields`. — `M/docs/explanation/custom-fields.md`; `M/backend/api/crm.yaml` `CreateCustomFieldRequest` (~line 46966)
- REST + MCP: contract-first OpenAPI (`M/backend/api/crm.yaml`, 592 paths); governed agent tools via `x-mcp-tool` (create/read/update/archive on offer, offer_template, product; search_records). Agents can draft offers (create, list, read, edit, add/remove lines, archive) but not send, accept, reject, regenerate or render. Each action needs a passport with a scope; autonomy tier green/yellow; append-only audit log. — `M/docs/reference/agent-tools.md`; `M/docs/handbook/offers-and-products.md` ("What an agent may do with offers"); `M/README.md`
- Agents cannot create contracts: `createContract` is human-only. Therefore neither an MCP agent nor an integration using a passport can create the contract record. — `M/backend/api/crm.yaml` lines 25471-25483
- MCP connector: served at `/mcp` only if `mcp.connector_enabled` and `--public-base-url`; `claude mcp add --transport http margince <base>/mcp`; OAuth consent where a human lends a passport. — `M/README.md`; `M/docs/how-to/connect-an-mcp-client.md`
- Import: `connector` enum is only `csv`; importable objects only `company, lead, contact`; upload via `/imports/sources`, dry-run, approve, undo, report. No Excel (.xlsx), no deals/offers/contracts import. (A Salesforce connector is "arriving with its own ticket".) A vCard import and LinkedIn import exist separately. — `M/backend/api/crm.yaml` `CreateImportRunRequest` (~line 31236), `ImportObject` enum
- Export: `/exports` filtered export (contact, company, deal, lead, project) to open formats, plus whole-workspace bundle `/exports/bundle`. Not offers. — `M/backend/api/crm.yaml`
- Approvals: yellow-tier agent actions stage for human approval (`/approvals`); approval kinds are entity-agnostic (update_record etc.). Roles/RBAC objects configurable (`/roles`). — `M/docs/handbook/approvals.md`; `M/docs/explanation/rbac-roles-and-teams.md`
- Frontend screens for all of the above are implemented: offers, offer templates, products, contracts, custom fields, automations, webhooks, reporting, import. — `M/frontend/src/screens/offers.tsx`, `offertemplates.tsx`, `products.tsx`, `contractform.tsx`, `customfields.tsx`, `automations.tsx`, `webhooks.tsx`, `reporting.*.tsx`, `import.tsx`

### Inferences
- INFERENCE (feasible path A, no core change): webhook `offer.accepted` -> external service (script, n8n or similar) fetches the offer via REST, generates the order-confirmation document in its own tooling, and a human uploads/attaches it (documents) and records the contract by hand. Limit: the REST call for contract creation is cookie-session-only.
- INFERENCE (path B, small core change): a core handler on `offer.accepted` creating a draft `contract` (and/or document) in the same transaction, plus an optional `offer_id` on contract; this contradicts the stated "assistant never writes a contract" rule only if done by an agent, and a system-attributed draft would need a deliberate policy decision.
- INFERENCE (path C, extension): an extension unit with own tables and a screen could hold quote log / confirmation data and subscribe to `offer.accepted`, but cannot write core records beyond activities, so it can log and notify, not create contracts in core.
- INFERENCE: The "quote log" currently kept in Excel would be a one-time CSV import only for companies/contacts/leads; QuoteLog rows (quotes) cannot be imported as offers by the importer.

### Gaps
- Did not read `docs/how-to/create-a-workflow.md` or `register-a-webhook.md`; the automation count mismatch (doc vs code) was only checked in code.
- Whether the REST API accepts any non-agent machine credential for human-only operations (service account) was not found; `/passports` mint agent credentials only (per README), `cookieAuth` marks human-only routes.
- Whether `request_approval`/stage automation can be wired to "offer accepted" was not verified (no offer trigger in the closed trigger list; `record_created_updated` on any entity might cover offers, unverified).

## Q5. What do the roadmap and changelog say about quotes, contracts, CPQ?

### Takeaway
No CPQ is planned in the files I read. The repo states offers are deliberately "never a configurator". Delivery of an offer to a buyer ("transport") is the one explicitly flagged missing capability. The changelog documents offer fixes and the 0.0.1 release (2026-09-10); I found no roadmap file mentioning quotes-to-contract.

### Cited Findings
- Version state: the CHANGELOG has `[Unreleased]` and `[0.0.1] - 2026-09-10`; product marked "coming autumn 2026". — `M/CHANGELOG.md` lines 15, 126; `M/README.md` line 3-4
- Product described as "a fast core for the 80% every sales team needs, plus a governed agent surface"; BUSL-1.1 license; "no Gradion-run SaaS in this repository"; runs on own infrastructure. — `M/README.md`
- Changelog items on offers: PDF no longer served to a reader who cannot open the buyer company; a record page no longer offers writes the server would refuse (New offer, stage move); passport scope refusal on `POST /offers/{id}/send` without `send` scope; `#/offer-templates` and `#/automations` standalone routes folded into Settings. — `M/CHANGELOG.md` lines ~119-125, ~734-736, ~1031, ~967
- Explicit missing capability: "delivery to a counterparty is a separate capability that does not exist yet" (sending an offer performs no transport). — `M/docs/reference/agent-tools.md` line ~258; `M/backend/api/crm.yaml` `sendOffer`
- No hits in README, docs, user-guide, AGENTS.md for CPQ, "configure-price-quote", e-signature roadmap or order-confirmation/sales-contract generation (grep of docs/, user-guide/, README.md, AGENTS.md, DESIGN.md). The only "quote" hits are unrelated (quoted mail text, estimates, quoting SQL) plus the customer-requests doc mentioning a mail task "Send the quote". — grep across `M/docs`, `M/user-guide`, `M/README.md`, `M/AGENTS.md`
- Open work is tracked in GitHub issues (`github.com/margince/margince/issues`), not in the repo, so any quote/contract roadmap there could not be checked. — `M/README.md`
- Repo process rules relevant to any custom build: contract-first (OpenAPI is the authority), every write commits domain row + audit_log + event_outbox in one transaction, `make check` as merge gate, an `AGENTS.md` for engineering rules, AI-disclosure rules in CONTRIBUTING. — `M/README.md`; `M/AGENTS.md`; `M/CONTRIBUTING.md`

### Inferences
- INFERENCE: Quote-to-contract is not on a visible roadmap; it would be a custom build on top of the existing offer/contract objects, consistent with the "you change it by changing the code" stance.

### Gaps
- GitHub issues / external roadmap not accessible from the repo; I did not browse them (read-only repo scope).
- `M/docs/README.md` index and `M/docs/evidence/` not searched for planned work beyond keyword greps.

## Summary table (existence check for the requested capabilities)

| Capability | Status | Evidence |
|---|---|---|
| Quote/offer object | EXISTS, implemented (API, backend, UI, tests) | `crm.yaml` Offer; `modules/deals/offer*.go`; `frontend/src/screens/offers.tsx` |
| Line items | EXISTS (qty, price, discount, tax, recurring billing) | `OfferLineItemInput` |
| Products / price list | EXISTS (flat rate card, snapshot into lines; no tiers/bundles/price books) | `Product`; `/products` |
| Deal stages / pipeline | EXISTS, configurable stages, Won/Lost | `/pipelines`, `/stages`, `the-pipeline.md` |
| Offer PDF from template | EXISTS, limited (header/footer text, DE/EN, Latin-only font, needs blobstore) | `offer_pdf.go` |
| Sales contract / order confirmation generation | MISSING | no endpoint, template or task |
| Docx / merge fields | MISSING | grep |
| E-signature | MISSING by decision ("stays removed") | `Contract.signed_on` text |
| Contract record | EXISTS, manual, human-only create, no offer link | `/contracts` |
| Auto-create contract on acceptance | MISSING | accept only syncs deal amount |
| Offer delivery to customer | MISSING ("does not exist yet") | `agent-tools.md` |
| Custom fields | EXISTS (not on offer) | `/custom-fields` |
| Saved reports / dashboards | EXISTS (fixed metrics/blocks, schedule, export, targets) | `/analytics/reports` |
| User-defined reports on offers | MISSING (entities limited to deal/lead/activity) | `reportspecs.go` |
| Workspace-wide offer list / quote log | MISSING | no `GET /offers` |
| Import CSV/Excel | CSV only; company, lead, contact only; no xlsx | `CreateImportRunRequest`, `ImportObject` |
| API | EXISTS (OpenAPI, 592 paths) | `crm.yaml` |
| Webhooks | EXISTS outbound, incl. `offer.*`, `contract.*` | `outbound-webhooks.md`, `public-events.yaml` |
| Automation engine | EXISTS but closed catalog; no offer/contract/document action | `automation.md` |
| Extension tier | EXISTS (compile-time Go units); core write port = activities only | `extensibility.md`, `pkg/extension/core.go` |
| MCP / AI agent layer | EXISTS; agents may draft offers, may not send/accept/render or create contracts; `offer_draft` AI task shipped; no contract-drafting task | `agent-tools.md`, `ai-tasks.yaml` |
| CPQ / roadmap for quotes-to-contract | none found in repo | greps, CHANGELOG |
