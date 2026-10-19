# Standard CRM systems: quote to contract / order confirmation (state 2026-10-18)

Method note: research capped at about 16 tool calls. Official pages were fetched for HubSpot (pricing, e-signature doc), Salesforce (Revenue Cloud pricing), Zoho (pricing, Sign doc, auto-convert solution page), Microsoft (pricing). Odoo documentation page and Pipedrive pricing page could not be fetched (empty / HTTP 403 / 404), so Odoo and Pipedrive rest on search snippets and are marked as such. Prices are list prices as shown on the pages on the fetch date (2026-10-18) and may differ by country/currency.

## 1. What is automated when a customer accepts, what a small team must set up, licence tier, limits (per system)

### Takeaway
All five systems can turn an accepted quote into an order or signed document, but the degree of automation differs sharply. Odoo does it natively and fully (sign, confirm, order, delivery/production). HubSpot, Zoho and Dynamics provide the objects and buttons; hands-free chaining needs workflows or custom code. Salesforce Revenue Cloud is the most capable but is priced and scoped for larger organisations. Pipedrive is the lightest and relies on Smart Docs or marketplace apps.

### Cited Findings

**HubSpot (now labelled "Revenue Hub" in the docs/pricing; formerly Commerce Hub)**
- E-signatures on quotes need a Revenue Hub Professional or Enterprise subscription and a Revenue Hub seat to configure; core-seat users can only be sender or countersigner. Legacy quotes need Sales Hub Starter/Pro/Enterprise. — [HubSpot KB](https://knowledge.hubspot.com/quotes/use-e-signatures-with-quotes)
- After signing: buyer and countersigners are notified, quote is "signed but not yet paid", signed PDF is sent to customer; if the contracts setting is on (and not in the Connected CPQ/Billing/Payments beta) a contract record is created automatically on acceptance; with online payments on, the buyer is led to payment. — [HubSpot KB](https://knowledge.hubspot.com/quotes/use-e-signatures-with-quotes)
- E-signature limits are pooled per account per month; a signature is counted once e-sign is turned on for a published quote even if never signed; resending an expired quote counts again; reset on the 1st of the month. — [HubSpot KB](https://knowledge.hubspot.com/quotes/use-e-signatures-with-quotes)
- List prices (USD, per seat/month, fetched 2026-10-18): Professional $85 annual commitment / $95 monthly; Enterprise $140; promo for new customers on annual billing $57 / $98. Billing/payments tools free for up to 2 users. — [HubSpot pricing](https://www.hubspot.com/pricing/commerce)
- Professional: quote templates, quote workflows, standard approvals (up to 10 approvers), signable attachments, product library up to 100 products, 25 e-signatures. Enterprise: advanced approvals (5 steps), quote rules (compatibility, pricing, quantity), product library up to 15 million, 50 e-signatures. More via paid limit increase. — [HubSpot pricing](https://www.hubspot.com/pricing/commerce)
- The e-signature doc does not mention workflow triggers; a third-party blog claims e-signature can trigger workflows — [Bizzyweb](https://bizzyweb.com/hubspot-commerce-hub) (unverified, third party). Third-party Hublead claims Sales Hub Starter suffices for e-sign; conflicts with the official doc, which is preferred.
- Setup for 4-7 users: product library, quote template, approval rule, e-sign, workflow(s). HubSpot partner implementation commonly quoted as $5,000-$20,000 (vendor-biased, see section 4).

**Salesforce (Sales Cloud + CPQ / Revenue Cloud)**
- CPQ went End of Sale in March 2025; existing customers keep support and can renew/add licences; EOL date disputed between consultancy sources (one says about 2029-2030, another says none announced) — [Everstage](https://www.everstage.com/cpq/salesforce-cpq-end-of-life), [Girikon](https://www.girikon.com/?p=25562%2F). Not verified on Salesforce's own lifecycle page.
- Revenue Cloud is a paid add-on to a Sales Cloud subscription, billed annually. Growth $150/user/month (quoting, configurator, order capture, subscriptions); Advanced $200/user/month (adds contracts and orders, invoicing, AI/analytics). Advanced needs an annual contract; no free trial. Fetched 2026-10-18. — [Salesforce pricing](https://www.salesforce.com/sales/revenue-lifecycle-management/pricing)
- Sales Cloud licence is on top (price not fetched, gap). Product reportedly renamed Agentforce Revenue Management — [search result summary of Salesforce regional page / third party](https://www.salesforce.com/in/sales/revenue-lifecycle-management/pricing/).
- Third-party: Revenue Cloud data model differs from CPQ, pricing logic should be redesigned, reimplementation 100k-500k USD per one negotiation-advisory firm (biased source, enterprise scale) — [Redress](https://redresscompliance.com/salesforce-revenue-cloud-migration-landing).
- Not covered by fetched sources: standard Sales Cloud Quotes templates, Contracts object, Flow automation on accept. Native Quote PDF templates and Flow exist but I did not verify them from documentation here (gap).

**Odoo (Sales app)**
- Online signature or online payment is enabled per quotation template (Sales > Settings > Quotations & Orders, then template "Confirmation" tab). Customer receives a portal link, signs (draw/auto/upload) and the quotation is confirmed and converted to a sales order, which can trigger delivery order, manufacturing order or project task depending on configuration; signed document stored on the order. Expired quotations cannot be signed. — [Cybrosys Odoo 18 guide](https://www.cybrosys.com/odoo/odoo-books/v18/sales/quotation-template/), [Odoo docs mirror](https://odoo-online.readthedocs.io/sv/latest/sales/online_customer_approval_on_quotation.html) (third party / old mirror; official docs page could not be fetched).
- Pricing (Odoo Online, fetched 2026-10-18): Standard about EUR 19.90/user/month annual (24.90 monthly), includes all apps incl. Sales and Sign; Custom about EUR 35.90 annual (44.90 monthly) adds Studio, multi-company, external API, Odoo.sh/on-premise hosting; One App Free plan exists. Not included: implementation services, custom code maintenance. — [Odoo pricing](https://www.odoo.com/pricing) (page layout hard to parse, confirm figures).
- Odoo is the only one here where quote, order, stock, manufacturing and invoice are in one database, so "handover to ERP" is not a separate step if Odoo is also the ERP.

**Zoho CRM (Quotes, Sales Orders, Blueprint, Zoho Sign)**
- Edition list prices per user/month (page shown in INR, fetched 2026-10-18): Standard 800, Professional 1,400, Enterprise 2,400, Ultimate 2,600; quotes/sales orders from Professional upward; Blueprint and Zoho Sign edition requirements not stated on that page. — [Zoho CRM pricing](https://www.zoho.com/crm/zohocrm-pricing.html)
- Quote to sales order: manual Convert (custom fields must be mapped first under Quote Conversion Mapping), Mass Convert, workflow-rule convert action for quotes, or API convert action. — [Zoho help search results](https://help.zoho.com/portal/en/kb/crm/automate-business-processes/workflow-management/articles/convert-records-using-workflow-rule), [API v8](https://www.zoho.com/crm/developer/docs/api/v8/inventory-convert.html)
- Zoho's own "auto-convert deal quote into sales order" recipe: workflow on Deals fires when a picklist changes, runs a Deluge function creating a sales order per related quote; needs Enterprise edition or above, sandbox testing, exact API names. — [Zoho solution page](https://www.zoho.com/crm/resources/solutions/auto-convert-a-deal-quote-into-a-sales-order.html)
- Zoho Sign: design an Inventory template (Settings > Customization > Templates > Inventory) with merge fields, then "Send with Zoho Sign" on a Quote or Sales Order; choose recipients, signing order, fields. Status sync back to CRM and edition requirements not covered on the page. — [Zoho Sign help](https://help.zoho.com/portal/en/kb/zoho-sign/integrations/zoho-apps/zoho-crm/articles/zoho-sign-inventory-templates)
- Zoho Sign third-party-reported prices: Standard $10, Professional $16, Enterprise $22 per user/month; credits add-on 0.10 USD each, API send reportedly 5 credits — [Signeasy blog](https://signeasy.com/blog/business/zoho-sign-pricing), [Zoho community](https://help.zoho.com/portal/ja/community/topic/announcing-zoho-sign-credits-a-revamp-of-our-add-on-pricing?page=9) (not official pricing page).
- Blueprint: official help is an index only in what I could fetch; details (edition, Zoho Sign action) unverified — [Zoho help](https://help.zoho.com/portal/en/kb/crm/automate-business-processes/blueprint/articles/blueprint).

**Lightweight option A: Dynamics 365 Sales**
- Create Order from an activated quote; products carry over; "Date Won" defaults to today; accepted quotes cannot be revised; order lines can be refreshed to current price list or price locked. In Sales Professional, an admin must enable order creation from quotes. Requires Professional, Enterprise or Premium licence plus sales role. — [Microsoft Learn search summary](https://learn.microsoft.com/en-au/dynamics365/sales/developer/convert-opportunity-quote-sales-order-invoice)
- List prices per user/month, billed yearly (fetched 2026-10-18): Professional $65, Enterprise $105, Premium $150. — [Microsoft pricing](https://www.microsoft.com/en-us/dynamics-365/products/sales/pricing)
- No native e-signature in the fetched material; usually via Power Automate plus a signing service (not verified here, gap).

**Lightweight option B: Pipedrive**
- Smart Docs (templates with deal/product fields, e-signature, tracking via PDF link) is free on Professional, Power, Enterprise plans; add-on for Essential/Advanced. — [Pipedrive blog](https://blog.pipedrive.com/features/smart-docs) (via search snippet; old 2021 add-on price $32.50/month likely outdated). Pipedrive pricing page returned 403, so current plan prices are a gap.
- No native sales order object; marketplace apps (PandaDoc, ExtraflowDoc) add quotes, e-sign, CPQ — [PandaDoc listing](https://www.pipedrive.com/en/marketplace/app/panda-doc-proposals-e-signatures-and-cpq/bde8d47250a2302c).

### Inferences
- For 5 users, 12-month licence cost roughly: HubSpot Pro about 5 x 85 x 12 = USD 5,100; Odoo Standard about 5 x 19.9 x 12 = EUR 1,200; Dynamics Professional about USD 3,900; Salesforce Sales Cloud + Revenue Cloud Advanced at least 5 x 200 x 12 = USD 12,000 before the Sales Cloud licence. Zoho Professional is the cheapest of the CRM-only options (in INR terms; convert before deciding).
- Salesforce Revenue Cloud is oversized for few high-value orders; CPQ's end-of-sale makes it a poor new start.

### Gaps
- Official Odoo 18/19 docs page content (fetch returned only a topic list); Odoo claims rest on guides.
- Pipedrive current plan prices; Zoho Blueprint and Zoho Sign edition limits; Salesforce standard quote templates/Contracts/Flow and Sales Cloud price; Dynamics e-signature options; HubSpot workflow trigger on quote acceptance (not confirmed in official doc).

## 2. Minimal pattern that removes most manual effort

### Takeaway
A quote built from a template (fields filled from deal, company, contact and line items), sent as a PDF or link with a signature or accept step, creating an order record and notifying the back office, covers most of the saving. The value is in not re-typing data and in having one signed document.

### Cited Findings
- Odoo shows the full minimal chain: template, online sign, auto-confirmation to sales order, downstream orders — [Cybrosys](https://www.cybrosys.com/odoo/odoo-books/v18/sales/quotation-template/).
- Dynamics: activate quote, Create Order, products carry over — [Microsoft Learn](https://learn.microsoft.com/en-au/dynamics365/sales/developer/convert-opportunity-quote-sales-order-invoice).
- Zoho: template with merge fields, Send with Zoho Sign, convert quote to sales order via workflow — sources in section 1.
- HubSpot: signed quote creates a contract record automatically when enabled — [HubSpot KB](https://knowledge.hubspot.com/quotes/use-e-signatures-with-quotes).

### Inferences
- Minimal pattern: (1) quote template with merge fields, (2) line items, (3) accept/sign or PDF confirmation, (4) one-click or workflow conversion to order confirmation, (5) email to back office. Approval step only above a value threshold.
- With few high-value orders, per-order manual conversion (one click) is acceptable; full automation of the accept event is optional.

### Gaps
- No independent time-saving measurements found.

## 3. Catalogue needed vs. template and fields only; what is NOT automated

### Takeaway
A product catalogue with prices is needed only for line-item totals, discounts, price lists, quote rules and approvals by margin. A document template with fields works without a catalogue, but then totals and line text are typed or pasted. None of the systems replaces ERP order entry, tax logic beyond basic rates, or pricing strategy.

### Cited Findings
- HubSpot product library limits (100 on Professional, 15 million on Enterprise); quote rules (compatibility, pricing, quantity) only on Enterprise — [HubSpot pricing](https://www.hubspot.com/pricing/commerce).
- Dynamics order lines depend on the price list; prices may be refreshed or locked on the order, i.e. price changes between quote and order are a real case to decide — [Microsoft Learn](https://learn.microsoft.com/en-au/dynamics365/sales/developer/convert-opportunity-quote-sales-order-invoice).
- Zoho quote to order conversion requires custom field mapping beforehand; auto-conversion needs custom Deluge code on Enterprise — sources in section 1.
- Salesforce: pricing logic should be redesigned, not copied, when moving to Revenue Cloud; reimplementation cost is significant — [Redress](https://redresscompliance.com/salesforce-revenue-cloud-migration-landing) (biased source).
- HubSpot is described as weak on CPQ and needing add-ons for manufacturing-type quoting — [Cloud House](https://cloudhousetechnologies.com/blog/crm-development-cost-manufacturing-companies-2026) (vendor source).
- Partial acceptance of quote lines is a documented limitation in a Dynamics Supply Chain context (workaround: split quotes) — [Microsoft Learn search result](https://learn.microsoft.com/sr-latn-rs/dynamics365/sales/developer/sample-process-quotes-sales-orders-invoices?view=op-9-1).

### Inferences
- Needs catalogue: totals, discounts, margin approvals, repeat items, reporting by product.
- Template and fields only: order confirmation text, payment and delivery terms, signatures, customer data, one summary price per deal. Fine for few, high-value, individually negotiated orders.
- Not automated by any of them out of the box (inferred from the absence of such features in the fetched docs, not from an explicit vendor statement): pricing rules for custom or machine-specific items without rules being built, customs/export tax treatment (VAT, reverse charge), entry of the order in a separate ERP (needs integration or manual entry), stock/lead-time checks outside Odoo, legal review of terms.

### Gaps
- No official statements on tax limits for quotes were found; no source on tax handling per system.

## 4. Typical effort and cost for a small manufacturer

### Takeaway
Directional vendor figures: USD 3,000-10,000 (Zoho/Pipedrive), 4,000-15,000 (Odoo CRM), 5,000-20,000 (HubSpot), 10,000-40,000+ (Salesforce) for standard SMB CRM projects; manufacturing-specific quote-to-order builds are quoted at USD 20,000-70,000. For 4-7 users with few orders, a template-and-fields approach in the existing CRM is far below these ranges.

### Cited Findings
- Ecosire (an Odoo implementer) ranges listed above; data migration USD 1,500-10,000; HubSpot onboarding often mandatory on Professional+ — [Ecosire](https://ecosire.com/fr/blog/crm-system-cost-guide-2026).
- Cloud House (custom build vendor): Zoho CRM Enterprise with Creator USD 20,000-40,000; HubSpot Sales Enterprise with integrations USD 45,000-70,000; recommends a paid two-week discovery sprint — [Cloud House](https://cloudhousetechnologies.com/blog/crm-development-cost-manufacturing-companies-2026).
- The Smarketers is a HubSpot Platinum partner (bias) — [The Smarketers](https://thesmarketers.com/?p=51448).

### Inferences
- All figures come from sellers of implementation services; treat as upper-range guides. For this client, a plausible low-cost path (own judgement, not sourced): templates plus conversion workflow configured in a few days of work.

### Gaps
- No independent day rates or case studies for machine-tool accessory manufacturers.
- No comparison of ongoing admin effort.
