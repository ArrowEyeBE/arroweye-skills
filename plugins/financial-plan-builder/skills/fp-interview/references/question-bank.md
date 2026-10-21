# Interview guide: Belgian BV financial plan (art. 5:4 WVV / CSA)

v0.5. The content of the interview; `../SKILL.md` says how to run it and how answers become `plan-data.json`.

## How to run the interview

- **Module 0 always comes first** and starts with the document language (NL / FR / EN).
  Conduct the conversation in whatever language the owner writes in; write the deliverable in the chosen document language.
- **Ask for documents first, questions second.** At the start of every module, ask whether the owner has a document that already answers it.
  If they do: read it, extract the fields, show them back as a short summary for confirmation, and ask only about what is missing.
  If they don't: ask the questions, one small group at a time (max 3–4 questions per turn).
- **Adapt to the business type (3.2).** The guide covers trade, production, hospitality, services, software, management companies and mixed businesses. Only ask the questions and templates that fit the selected types; skip the rest without mentioning them. Never assume a services business by default.
- **Skip empty modules with one question.** When the business type sells no goods and the owner has already said there are no staff, stock or investments (typical for a management company or a one-person consultancy), confirm it once ("So no stock, no staff and no investments in the plan period?") and skip modules 5, 6 and 8. Log the confirmation.
- **Never block on a number.** If the owner doesn't know, offer a sensible default or range (marked `default`), say what it is based on, and log it as an assumption to revisit.
- Every answer is written to `plan-data.json` (structured) and `interview-log.md` (verbatim + source), so the interview can be resumed.
- **Granularity:** the model is always monthly. The owner can give figures monthly, quarterly or as yearly targets.
  For yearly or quarterly input, ask how the year ramps up (linear / step at month X / seasonal) and spread it to months; show the result for confirmation.
- **Horizon:** from the incorporation date to the end of the 3rd financial year. The legal minimum is 2 years, but 3 is market practice.
- **Disclaimer, once, at the start (after 0.1, in the chosen language):** "I'll help you draft the financial plan and the calculation model. This is not tax, legal or accounting advice: you, the founders, remain responsible for the plan. The Belgian tax and payroll values I use were last checked on <`last_checked` from `belgian-rules.json`>, and each comes with its official source. Have the final plan reviewed by an accountant." If that date is older than `review_after_days` or from an earlier calendar year, add that the rules may be outdated and offer to run the `fp-rules-check` skill first.
- **Every module feeds a section of the plan.** The tag `→ §n` refers to the legal requirements of art. 5:4 WVV:
  §1 activity · §2 financing sources · §3 balance sheets · §4 P&L · §5 budget/cash flow · §6 assumptions · §7 external expert.

## Checkpoints (owner sign-off)

The owner approves the work at 4 checkpoints. At each one, show a **short summary**, not the full document, and ask for approval or changes.
Changes always go back into `plan-data.json`, and everything after it is regenerated. Never patch the Word or Excel by hand: the text, tables and chart must stay in sync.

| # | When | What the owner sees | Typical errors it catches |
|---|---|---|---|
| CP1 | End of the interview (modules 0–10) | Assumptions sheet: every driver per revenue stream, hires with their months, cost lines, investments, financing (including explicit "none"), every value marked `default` | Wrong rates or volumes, forgotten costs, financing that doesn't exist |
| CP2 | After the model is built | Key figures: revenue / gross margin / operating profit / net profit per year, cash curve, lowest cash month, headcount per year, plus the model's checks | "This doesn't look like my business", cash shortfalls |
| CP3 | After the text is drafted and reviewed by the fp-reviewer agent | The activity description, financing sources and assumptions narrative (numbers are placeholders filled from the model), plus the reviewer's questions | The text contradicting the numbers (wrong margin label, hires that aren't in the model) |
| CP4 | Final | Complete .docx + .xlsx, plus the reviewer's verdict on the final document (all 7 elements of art. 5:4) | Missing sections, signature page details |

---

## Module 0: Setup

| # | Question | Notes |
|---|---|---|
| 0.1 | In which language should the financial plan be written: **Nederlands, Français or English**? | Ask first. Hint: the notary's working language usually follows the region of the registered office. |
| 0.2 | Show the **preparation checklist** (below) in the chosen language. Ask: "Do you have most of this ready, or do you want to prepare first?" | **Start now:** continue with 0.3. **Prepare first:** generate the checklist as a .docx in the chosen language (so it can be shared with co-founders), save progress, and stop. The interview resumes at 0.3 next time. |
| 0.3 | Who am I talking to (founder / one of several founders / an adviser)? | |
| 0.4 | Please upload the documents you have from the checklist. | Sort each document to the module it belongs to (see the "Used in" column of the checklist). Items that are still missing are not blockers: they become defaults or open items, all listed at CP1. |
| 0.5 | Is a notary appointment planned? On what date? | Sets the deadline; the plan must be ready before the deed is signed. |

### Preparation checklist (step 0.2)

Keep it to one page. The checklist is derived from this question bank: when a module changes, update the matching item here.

**Documents to gather**

| Item | Used in |
|---|---|
| Draft statutes from the notary (company purpose, first financial year) | 1.4, 1.5 |
| Existing business plan, pitch deck or forecast, if any | 3, 4 |
| Price list or rate card | 4 |
| Signed contracts, letters of intent, reseller orders or concrete leads | 4 |
| Supplier price lists and a quote for the first stock order (trade, production, hospitality) | 5 |
| Lease agreement or rental offer for premises, if any | 3.7, 7 |
| Quotes for investments (fit-out, machines, kitchen, IT, vehicle, renovation) | 8 |
| Payroll cost simulation from a payroll provider, if you plan to hire | 6 |
| Loan offer from the bank or subsidy decision, if applicable | 9 |
| Management company: the (draft) management agreement with the client company | 4 |

**Decisions to make with your co-founders**

| Item | Used in |
|---|---|
| How much each founder contributes at incorporation, and in what form (cash / in kind) | 2.4, 9.1 |
| Share split | 2.5 |
| Will any founder lend money to the company? If not, say so explicitly. | 2.7, 9.2 |
| Founder pay: how much, from when, and who pays the director's social contributions. If a founder is already paid as a director elsewhere, a recent pay slip | 2.6 |
| Has each director been self-employed before (main or side activity)? For how many quarters in the last 10 years? | 2.9 |
| Does each founder keep another job, or work on this full-time? | 2.3 |

**Numbers to have a view on**

| Item | Used in |
|---|---|
| Your prices, and how much you expect to sell per month in year 1 (units, orders, guests, projects or billable days, whatever fits your business) | 4 |
| What it costs you to buy or make what you sell (purchase price, food cost %, materials) | 5 |
| How much stock you need to start with, if any | 5 |
| Opening hours or shifts, and how many people you need to run them | 6 |
| When you expect to hire, and whom | 6 |
| Your biggest fixed costs (rent, energy, software, vehicles) | 7 |
| Planned incorporation date and end of the first financial year | 1.3, 1.4 |

## Module 1: Company basics → cover, company data, §1

| # | Question | Minimum needed if no document |
|---|---|---|
| 1.1 | Company name | required |
| 1.2 | Registered office address (*zetel / siège*) | required |
| 1.3 | Planned date of incorporation / start of activity | required |
| 1.4 | End date of the first financial year (it may be longer or shorter than 12 months) | default: 31/12 of the next year if starting in Q4, otherwise 31/12 of the same year |
| 1.5 | Company purpose (*voorwerp / objet*) | **Document:** draft statutes → copy literally. Without statutes: draft a purpose clause from module 3 and mark it "to be aligned with the notary". |
| 1.6 | Is the company VAT-liable? Will VAT returns be monthly or quarterly? | default: VAT-liable, quarterly |
| 1.7 | NACE codes, if known | optional |

## Module 2: Founders → §1, §2

Per founder:

| # | Question | Notes |
|---|---|---|
| 2.1 | Full name, address, birth year | required (shown in the plan) |
| 2.2 | Role (director / shareholder only) | |
| 2.3 | Main occupation or side activity (*hoofdberoep / bijberoep*)? Any other income? | Affects social contributions and how much remuneration is realistic. |
| 2.4 | Contribution: cash (€), in kind (what, value), or none | Feeds equity in the opening balance. **In kind** (laptop, van, client portfolio, software…): explain that art. 5:7 WVV requires a special report by the founders and a report by a company auditor (*bedrijfsrevisor*), filed with the deed. Ask whether they still want it in kind (often a sale to the company or a cash contribution is simpler); if so, add the auditor's fee as a start-up cost. |
| 2.5 | Share split between founders | |
| 2.6 | Planned monthly remuneration from the company, from which month, and does it change over time? | **Document:** a recent pay slip if the founder is already paid as a director (common when setting up a management company). Split it into: gross pay → `618`; social contributions the company pays on the director's behalf → `618` (part of the remuneration; or the `dir_social` default); flat-rate expense or daily allowances → `61` (a reimbursement of costs, not remuneration); benefits in kind (phone, car) → their own line. Ask who will pay the social contributions. A pay change in a given month → a `schedule`. See tax note T1. |
| 2.7 | Will the founder lend money to the company? Amount, interest rate, repayment term | **Must appear in financing sources** (a frequent omission). |
| 2.8 | Relevant experience / why this team | Used in the activity description. |
| 2.9 | Per director: in the 40 quarters (10 years) before incorporation, were you self-employed (main or side activity, or as a helper) for more than 12 quarters? | No for every director → `company.contribution_exempt: true`: no company contribution during the first 3 years from the year of incorporation (the company requests it from its social insurance fund). Yes for one of them → `false`. Unsure → leave it out and add an open item. Rule and source: `belgian-rules.md`. |

## Module 3: Activity & business model → §1

| # | Question | Notes |
|---|---|---|
| 3.1 | Describe in a few sentences what the company does. | **Document:** business plan / pitch deck / website. |
| 3.2 | **What type of business is it?** Select all that apply: **trade** (buy and resell: shop, webshop, wholesale, import/distribution) · **production** (make your own products: manufacturing, food production, crafts) · **hospitality** (restaurant, café, bar, catering, accommodation) · **services** (agency, consultancy, trades/contractors, care, cleaning, …) · **software / digital products** (SaaS, apps, licences, online courses) · **management** (a management or holding company: the founder's director mandate and management services invoiced to one or a few related companies, often the founder's own operating company) · **other** (rental, platform/commission, …) | **Decides which questions apply in modules 4–8.** Many companies combine types (e.g. a bakery = production + trade; a bike shop = trade + repair service). |
| 3.3 | Which products or services do you sell? List them as separate offerings or product ranges. | Each one becomes a candidate revenue stream in module 4. |
| 3.4 | Who are the customers (consumers / businesses / government, size, sector, region, export)? | |
| 3.5 | **Through which channels do you sell?** Own shop · webshop · marketplaces (Bol, Amazon, Etsy, …) · resellers / wholesale · distributors · direct sales · markets / events | A channel can have its own price, margin, fees and payment terms. In that case it gets its own revenue stream. |
| 3.6 | How do you charge? Per product/unit · per order or basket · per guest/visit · per hour/day · per project · subscription · licence · commission · rental | Decides which revenue template module 4 uses. |
| 3.7 | Do you need premises (shop, workshop, production site, warehouse, kitchen, office)? Where? Rented or owned? From when? | Feeds rent, fit-out, energy and permits in modules 7–8. Revenue cannot start before the premises open. |
| 3.8 | Competitors, and why customers choose you | Skip for a management company with a captive client: there is no market in the usual sense, and `activity.market` then only describes the client relation (or is left out). |
| 3.9 | How do you see the business evolving over years 1, 2 and 3? (new products, new channels, a second location, export, bigger customers, stopping an activity) | Becomes the year-by-year assumptions narrative (module 11). |

## Module 4: Revenue → §4, §5, §6

**Document first:** an existing forecast, sales figures from a comparable shop/business, pipeline/CRM export, signed contracts, letters of intent, orders from resellers, price list.
Even with a document, confirm the drivers below so the assumptions can be quantified.

One block per revenue stream (product range × channel, or service line). Pick the matching template. Every template ends with: start month · ramp-up · growth per year.

**A. Product sales: shop / webshop / marketplace (B2C)**
- Orders or customers per month × average basket value. For a physical shop optionally: visitors per day × conversion % × average basket × opening days.
- Prices in B2C are usually quoted **incl. VAT**: ask which, and convert to excl. VAT.
- Returns % (important for webshops) · discounts and sales periods · marketplace or platform commission (→ module 5)

**B. Product sales: wholesale / resellers / distributors (B2B)**
- Number of resellers or B2B customers × order value × order frequency · price to reseller (or the reseller's margin on the retail price) · minimum order quantities · new resellers added per year

**C. Production / manufacturing**
- Per product: units sold per month × price (by channel: A or B above) · production capacity per month (machines, people, shifts) · production ramp-up after start

**D. Hospitality**
- Opening days per month · guests (covers) per day per service (breakfast/lunch/dinner) · average spend per guest, split food / drinks (different VAT rates and margins) · takeaway or delivery share
- Accommodation: rooms × occupancy % × average room rate
- Catering: events per month × guests × price per guest

**E. Time-based services**
- Day or hour rate (excl. VAT) · billable days or hours per month per person · number of people billing · utilisation ramp-up

**F. Projects / contracts**
- Average project value · projects per month or quarter · project length · invoicing schedule (deposit % / milestones / on delivery)

**G. Recurring / subscription** (SaaS, memberships, maintenance contracts, boxes)
- Monthly price · new customers per month · churn % per month · price indexation

**H. Training / events / tickets**
- Sessions per month · participants per session · price per participant

**I. Other** (rental, commissions, licences/royalties, advertising): free description + driver + amount

**J. Management fee** (management company invoicing a related company)
- Client company (or companies) · monthly fee excl. VAT · from which month · steps (a higher fee from month X) · indexation
- What the fee replaces: what the client pays the founder directly today (pay, allowances, contributions), so the fee can be justified from the founder's side: what it must cover (the director's pay and contributions, the company's own costs, a reserve). **Never justify it with the client's forecast or figures**: the plan must not commit a third party.
- Payment: at invoice, or on terms. A captive client paying at invoice is a legitimate 0-month term: write the reason in `payment_note`.
- Is there a (draft) management agreement? If not, add an open item.
- Model: unit "month" (or "agreement"), volume 1 per month, `price_excl_vat` = the first fee, a `price_schedule` for each step. Use unit "€" only when there is no fixed fee.

For every stream, also ask:
- **When does the customer pay?** At the till or online (0 days) · payout delay from a webshop/marketplace platform (e.g. weekly or bi-weekly) · deposit upfront · on invoice (B2B, default 30 days). A term that departs from the usual (B2B paid at invoice) gets its reason in `payment_note`, so the reviewer doesn't ask again.
- **Price changes on a known date** (a new rate card, a fee step): `price_schedule` instead of an amount-only stream.
- **Seasonality:** strong and weak months (e.g. December peak for retail, summer for terraces/tourism, quiet August for B2B)
- **VAT rate:** 21% standard; reduced rates apply to e.g. food (6%) and restaurant meals (12%, drinks 21%). Verify in `belgian-rules.md`.
- **Confidence:** contracted / likely / ambition. This feeds the narrative and the optional scenarios.

Sanity checks to raise with the owner:
- **Capacity:** is revenue ≤ what can physically be delivered? People × available days for services; machines and shifts for production; seats × table turns × opening days for hospitality; stock available for trade.
- Is year-on-year growth > 100% without matching hires, investments or extra stock?
- Is revenue booked before the activity can start (before incorporation, before the premises open, before the first stock arrives)?

## Module 5: Direct costs, purchasing & stock → §3, §4, §5

For trade, production and hospitality, this is usually the module with the most impact on cash. Spend time here.
**Document first:** supplier price lists, quotes for the first stock order, purchase conditions.

| # | Question | Notes |
|---|---|---|
| 5.1 | Per revenue stream: which costs rise directly with sales? Purchase price of goods · raw materials and ingredients · packaging · shipping and fulfilment · payment fees (card terminal, online payment provider) · marketplace/platform commission · sales commissions · royalties · subcontracted production or services | As € per unit, % of sales (e.g. food cost %), or € per project |
| 5.2 | **Opening stock:** what must be bought before you can sell? Amount and month. | Often the biggest cash need at launch for trade and production. |
| 5.3 | **Stock policy:** how much stock do you keep (weeks or months of sales)? Minimum order quantities? Delivery lead times? | Drives stock on the balance sheet and the cash tied up in it. |
| 5.4 | Supplier payment terms. Do suppliers require prepayment or a deposit (common for first orders and imports)? | default 30 days; first orders often prepaid |
| 5.5 | Do you buy from outside the EU? Transport, customs duties, currency risk | Add to the cost price. |
| 5.6 | Will you outsource part of the work (subcontracted production, subcontractors, logistics partner, freelancers)? When, how much? | |
| 5.7 | Expected waste, breakage or unsold stock (%) | Relevant for food, fashion, perishables |

Model note: purchases go out as cash when paid; stock sits on the balance sheet (NBB code 3) until it is sold. Only the cost of goods actually sold hits the P&L.

## Module 6: Personnel → §4, §5, §6

**Document first:** a payroll simulation from a social secretariat (SD Worx, Securex, Acerta, Liantis, Partena, …) gives the exact employer cost. This is the best possible input.

For shops, hospitality and production, start from the **opening hours or shifts**: how many people are needed per shift, and how many shifts per week? Then turn that into roles and FTEs.

Per planned hire:
- role · start month · FTE · gross monthly salary (or total employer cost if known)
- Contract type: employee (permanent or fixed-term) · part-time · student worker · flexi-job (e.g. hospitality, retail) · seasonal worker · temp-agency worker (interim). Each has a different cost; see `belgian-rules.md`.
- If only the gross salary is known: use the employer-cost formula in `belgian-rules.md` (gross × 13.92 months × (1 + employer social security contributions) + insurance + extra benefits), marked as an estimate.
- Extras: meal vouchers, company car, phone, laptop, work clothing, home-office allowance
- Is this the company's first hire? (Reduction on employer social security contributions for the first employee; see `belgian-rules.md`.)

Once there is staff, always add: workplace-accident insurance, payroll-service fee, occupational health service.

## Module 7: Operating costs → §4, §5

Walk through the checklist below, **filtered by the business types from 3.2** (only show lines that can apply). For each line the owner confirms, adjusts, or answers "not applicable".
Pre-fill with the proposals of `scripts/cost_defaults.py` (from `templates/cost-defaults.json`, per business type); see "Default cost lines" in `../SKILL.md`.

**All companies**
- One-off: notary · registration in the KBO/BCE · publication in the Belgian Official Gazette · bank account opening · logo/branding · website
- Monthly: accountant · software/IT/cloud · telecom/mobile · internet · marketing & advertising · transport / fuel / car lease · meals & representation · director expense allowance · bank fees
- Yearly: liability insurance (BA) · other insurance (fire, legal aid, cyber) · filing the annual accounts with the NBB/BNB · annual contribution of companies to the social insurance fund (*vennootschapsbijdrage*) · membership fees (UNIZO, Voka, sector federation, …) · domain/hosting
- Event-driven: workplace-accident insurance (once there are employees) · recruitment costs

**Companies with premises (shop, workshop, production site, warehouse, restaurant, office)**
- One-off: rental deposit (not a cost: cash tied up, booked as an asset) · permits (environmental, hospitality, terrace, sign) · opening or launch campaign
- Monthly: rent + service charges · energy (often a major cost for hospitality and production) · water · cleaning · waste collection · maintenance & repairs · security/alarm
- Yearly: fire and contents/stock insurance · property-related taxes passed on by the landlord

**Trade and e-commerce**
- Monthly: webshop platform and apps · card-terminal rental · till system (POS) · storage / logistics partner · packaging materials (if not already in module 5)

**Production and food**
- Monthly or yearly: machine maintenance · quality control / certification · product liability insurance · registration and annual contribution with the federal food safety agency (FAVV/AFSCA), where applicable

**Hospitality**
- Monthly or yearly: music licences (Unisono) · linen and laundry · small kitchen equipment and tableware replacement · FAVV/AFSCA contribution

**Home-based / office services**
- Monthly: home-office allowance for directors · coworking

**Management companies**
- The director's social contributions if the company pays them (`dir_social`, computed from the 618 lines) · the director's expense allowance (from the pay slip) · accountant and invoicing software · bank fees
- Usually not: marketing, IT, liability insurance (often covered by the client company): `cost_defaults.py` leaves them out when the only type is management

## Module 8: Investments → §3, §5

| # | Question | Minimum needed |
|---|---|---|
| 8.1 | What will you buy that lasts more than a year? IT hardware · furniture · fit-out of shop, restaurant or office · kitchen equipment · machines and production lines · moulds and tooling · till and POS hardware · vehicles, delivery van · renovation of premises · software or webshop development · takeover of an existing business (goodwill) · building or land | description |
| 8.2 | Amount excl. VAT and month of purchase | required |
| 8.3 | Useful life / depreciation period | defaults in `belgian-rules.md` (IT 3 yrs, vehicle 5 yrs, machines 5–10 yrs, fit-out and furniture 10 yrs, buildings 33 yrs, …) |
| 8.4 | Bought, leased, or contributed in kind by a founder? | Leasing → monthly cost; contribution in kind → equity, and the auditor's report of 2.4 applies |

**Document first:** quotes.

## Module 9: Financing → §2, §3, §5

| # | Question | Notes |
|---|---|---|
| 9.1 | Total equity contributed at incorporation (cash + in kind) and when it is paid in | links to 2.4 |
| 9.2 | Loans from founders: amount, interest, term, timing | links to 2.7 |
| 9.3 | Bank loans: amount, interest, term, grace period, guarantees (personal guarantee, pledge, mortgage) | **Document:** bank offer. Guarantees must be mentioned. |
| 9.4 | Subsidies or support (VLAIO, Wallonia/Brussels regional aid, Participatiefonds/PMV, winwinlening/prêt proxi) | amount, timing, conditions |
| 9.5 | Leasing or credit lines (*kaskrediet*) | |
| 9.6 | Planned dividends during the horizon: from the result of which financial year, how much (gross), paid in which month? | → `financing.dividends[]`. A dividend from a year's result is decided by the general meeting that approves the accounts (within 6 months after year end); paid during the year itself = an interim dividend (check that the articles allow it). The model tests the net assets after each distribution (art. 5:142 WVV) and shows the cash in the 12 months after payment (art. 5:143). Withholding tax and VVPRbis are the shareholder's matter: the plan uses the gross amount. |
| 9.7 | Do the draft articles make (part of) the contributions unavailable equity (*onbeschikbaar eigen vermogen*)? | → `financing.unavailable_equity` (amount); default 0. Used in the net-asset test for dividends. |
| 9.8 | Capital increases during the horizon | not modelled: open item |

For loans (2.7, 9.2–9.3), grants (9.4) and contributions in kind (2.4), record an explicit "none" in `financing.confirmed_none` when the answer is no: the document states "there are no loans" only on that basis. Credit lines and capital increases (9.5, 9.8) are not modelled: record them as open items and tell the owner. The financing section of the plan must be built only from this module.

**Hard rule: never add financing the owner did not state.** No buffer loans, no placeholder credit lines, nothing taken over from a template.
Sufficiency check (after the model is built):
- Does cash stay > 0 every month?
- Is the lowest cash balance thin? (Threshold to be decided, e.g. less than 1 month of fixed costs.)

If either check fails, go back to the owner with the size and the month of the shortfall and ask how to cover it (more equity, a loan, cost cuts, a later hire). Never "fix" this silently.

## Module 10: Tax settings → §4, §5

| # | Question | Default |
|---|---|---|
| 10.1 | Reduced corporate tax rate (20% on the first €100,000): (a) Will one or more other companies hold half or more of the shares? (b) Will the company hold shares in other companies (e.g. the operating company's shares in a management company)? (c) Does the owner want the reduced rate in the plan, or the standard rate as a cautious choice? | See T1. (a) yes → no reduced rate. (b) yes → the financial-company exclusion may apply: no reduced rate in the plan, open item for the accountant. Otherwise propose `reduced_rate: true`; the owner may still choose 25%. |
| 10.2 | Will the company make tax prepayments (*voorafbetalingen*)? | No: starters are exempt from the surcharge in their first 3 financial years (verify the conditions). |
| 10.3 | Dividend policy | see 9.6 |

T1: the reduced rate requires a small company, shareholders that are not for half or more other companies, no "financial company", and, **from the fifth taxable period after incorporation**, a minimum remuneration for at least one director (art. 215 WIB 92; amounts and source in `belgian-rules.md`). Within a 2–3 year plan the remuneration condition therefore does not apply yet. Explain this neutrally, mention that the pay condition will matter from year 5, and recommend confirming with an accountant. Holding shares in other companies, dividends from them and capital-gains taxes are outside the plan: see `management-company.md`.

## Module 11: Assumptions narrative → §6

Runs after checkpoint 2, when the numbers are approved. This module is not new questions. It is a **synthesis**: draft the year-by-year story (Y1 / Y2 / Y3) from modules 3–10 with the key drivers (rates, volumes, hires, months) and the model's figures as placeholders (`skills/fp-document/references/document-text.md`), then ask the owner:
- Is this how you see it?
- What would make this go faster or slower?
- What are the main risks, and what is your fallback? (e.g. shorter opening hours, a smaller first stock order, delaying a hire, fewer product ranges, side income for the founders)

Before drafting, read the writing rules in `document-text.md`: describe the activity at the level of the purpose clause (no workplace, tool names or task lists unless they drive a number), keep facts about other people and companies out of the text, and say "the only client in the figures" rather than promising exclusivity.

## Module 12: External expert & signing → §7, signature page

| # | Question | Notes |
|---|---|---|
| 12.1 | Will an accountant or external expert review or co-sign the plan? If so: name, firm, ITAA number, enterprise number. | If not, the document states that the founders drew up the plan without an external expert (art. 5:4, §2, 7°). Recommend a review by an accountant anyway. |
| 12.2 | Place and date of signing; which founders sign | |

The document always ends with a **disclaimer** (before the signature page), whether or not an expert is involved: the plan is an estimate based on the founders' information, the tax and payroll values were checked on <date> against the official sources listed, laws change, and the founders remain responsible.

---

## Output contract

When modules 0–10 are complete:
1. Run `validate_plan_data.py` and fix its errors.
2. Show the checkpoint 1 summary (see `../SKILL.md`) and get explicit confirmation.
3. Hand over `plan-data.json` to the model-building step (fp-model). Modules 11–12 follow after checkpoint 2.

## Glossary (NL / FR / EN)

| NL | FR | EN |
|---|---|---|
| financieel plan | plan financier | financial plan |
| aanvangsvermogen | capitaux propres de départ | starting equity |
| inbreng (in geld / in natura) | apport (en numéraire / en nature) | contribution (cash / in kind) |
| voorwerp | objet | company purpose |
| boekjaar | exercice comptable | financial year |
| balans | bilan | balance sheet |
| resultatenrekening | compte de résultats | profit & loss statement |
| kasstroom | flux de trésorerie | cash flow |
| begroting | budget | budget |
| hypotheses | hypothèses | assumptions |
| externe deskundige | expert externe | external expert |
| bedrijfsleider / bestuurder | dirigeant / administrateur | director |
| vennootschapsbelasting | impôt des sociétés | corporate tax |
| voorafbetalingen | versements anticipés | tax prepayments |
| bijberoep | activité complémentaire | self-employed as a side activity |
| voorraad | stock / inventaire | stock / inventory |
| huurwaarborg | garantie locative | rental deposit |
| inrichting | aménagement | fit-out |
| brutomarge | marge brute | gross margin |
| managementvennootschap | société de management | management company |
| managementvergoeding | rémunération de management | management fee |
| onbeschikbaar eigen vermogen | capitaux propres indisponibles | unavailable equity |
| tussentijds dividend | acompte sur dividende | interim dividend |
| vennootschapsbijdrage | cotisation des sociétés | annual company contribution |
