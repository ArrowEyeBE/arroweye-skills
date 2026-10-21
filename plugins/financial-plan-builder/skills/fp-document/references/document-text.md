# Document text: contract and writing rules

The Word document has two kinds of content:

- **Generated** by `build_document.py`: every table, the chart, the financing overview, the statements about what does *not* exist ("no loans are taken out"), the dividends section, the tax paragraph, the note on the status of the purpose clause, the disclaimer (including "the plan is kept by the notary and not filed") and the signature page. Never write these yourself.
- **Owner text**: the fields below in `plan-data.json`, drafted by Claude from the interview and approved by the owner at checkpoint 3. They are written **in the plan's language** (`meta.language`).

## Fields

| field | required | document section | content |
|---|---|---|---|
| `founders[]` | yes | 2 Founders | `name`, `role` (e.g. bestuurder / administrateur / director), `address`, `birth_year`, `side_activity` (true/false). Optional fields are left out of the table when nobody has them. |
| `company.address` | yes | 1 Company details | Registered office. |
| `company.enterprise_number` | no | 1 | Only if already known; otherwise "to be assigned" is printed. |
| `activity.purpose` | yes | 3.1 | The company's object, copied from the draft articles of association (ask the notary for the draft). Don't invent it. |
| `activity.purpose_status` | no | 3.1 (note) | Only when the clause is not yet the notary's text, e.g. "draft by the founder, to be aligned with the notary". Printed as a note under the clause; never write it inside the clause. |
| `activity.offer` | yes | 3.2 | What the company sells: products, services, ranges, opening hours, channels. Concrete. |
| `activity.business_model` | yes | 3.3 | How it earns money per revenue stream: who pays, for what, how prices are set, payment terms. |
| `activity.market` | no | 3.4 | Customers, location, competition, why customers will come. Leave it out when there is no market in the usual sense (a management company with one related client): the section is then not printed. |
| `activity.outlook` | no | 3.5 | Planned evolution after the plan's horizon. |
| `document.financing_note` | no | 4 | Extra explanation of the financing (e.g. why this split between equity and loan). The table and the "no other sources" statements are generated. |
| `assumptions.years[]` | yes, one per financial year | 8.7 | What happens in that year and why: ramp-up, hires, investments, price changes, with the key numbers as placeholders. The heading with the financial year and its period is generated: don't start the text with `{{fy.1}}` or a bold year label. |
| `assumptions.profitability` | yes | 8.8 | Why the business is profitable (or when it becomes so): margin, cost structure, break-even, with placeholders. |
| `assumptions.risks` | yes | 8.9 | The 2–4 main risks and what the founders will do if they happen. |
| `document.date` | yes for final | cover, signature | Date of the plan, `YYYY-MM-DD`. |
| `document.place` | yes for final | signature | Place of signing. |
| `document.signatories` | no | signature | Defaults to all founders' names. |
| `document.expert` | no | 9 External expert (art. 5:4, 7°) | `{name, firm, itaa, enterprise_number, address}`. Leave out when no expert assisted; the section then states that the founders drew up the plan without an external expert. |

Formatting inside a text field: a blank line starts a new paragraph; a block of lines starting with `- ` becomes a bullet list; `**bold**` is allowed.

## Numbers: always placeholders

**Every figure that comes out of the model must be written as a placeholder**, never typed. The generator fills it from the recalculated workbook, so the text cannot contradict the tables, even after the numbers change.

| placeholder | value |
|---|---|
| `{{revenue.1}}` … | P&L line per financial year (1 = first). Names: `revenue`, `grants`, `goods`, `services`, `gross_margin`, `staff_cost`, `depreciation`, `operating_result`, `financial_costs`, `profit_before_tax`, `corporate_tax`, `net_result` |
| `{{cash_end.2}}`, `{{equity.2}}`, `{{total_assets.2}}` | balance sheet at the end of a financial year |
| `{{lowest_cash_fy.1}}` | lowest month-end cash within a financial year |
| `{{gm_pct.1}}`, `{{op_pct.1}}`, `{{net_pct.1}}` | gross margin, operating result and net result as % of revenue |
| `{{fte.3}}` | staff in FTE at the end of a financial year |
| `{{lowest_cash}}`, `{{lowest_cash_month}}` | lowest cash over the whole plan, and its month |
| `{{stream.<id>.price}}` | price per unit of a revenue stream at the start (input) |
| `{{stream.<id>.price.2}}` | price in the last month of a financial year (follows a `price_schedule`) |
| `{{stream.<id>.volume.1}}`, `{{stream.<id>.revenue.1}}` | volume and revenue of one stream in a financial year |
| `{{equity_cash_total}}`, `{{inkind_total}}`, `{{loans_total}}`, `{{grants_total}}`, `{{financing_total}}`, `{{investments_total}}` | totals from `financing` / `investments` |
| `{{fy.2}}` | label and period of a financial year, e.g. "FY 2 (01/01/2028 – 31/12/2028)" |
| `{{start_date}}` | start date, written out |

Any other P&L or balance-sheet row key of the model (`a70`, `bs_cash`, …) also works as `{{a70.1}}`.
An unknown placeholder stops the final build.

**Placeholders carry their unit and sign.** Money comes out as "€ 5.000" (nl), "5 000 €" (fr), "€5,000" (en); percentages with "%"; a loss with a minus sign ("-€ 781"). So:
- never write "euro", "EUR", "€" or "%" next to a placeholder: "{{revenue.1}} euro" gives "€ 5.000 euro". The build stops on it.
- `{{name.1:abs}}` drops the minus sign, for "a loss of {{net_result.1:abs}}" (not "a loss of -€ 781").
- `{{name.1:num}}` gives the bare number (no € or % sign), for a table-like sentence or another unit.

Numbers that are **drivers given by the owner** and are not in the model as such (e.g. "26 opening days a month", "about 120 customers a day", "30 days payment term") may be written in the text. They must match the volumes and settings in `plan-data.json`; the reviewer checks this.

## Writing rules

- **Only what is in the model.** Every hire, investment, loan, price change or new activity mentioned in the text must exist in `plan-data.json`, in the same month or year. If the owner mentions something that isn't in the numbers, either add it to the numbers (back to checkpoint 1/2) or leave it out of the text.
- **Never mention financing that does not exist**, and never deny financing that does. The generated financing chapter is the reference.
- **Belgian margin terms.** *Brutomarge / marge brute* (code 9900) in the NBB schema is revenue minus goods (60) **and** services and other goods (61, which includes rent, director pay and most overheads). Don't call it the margin on goods. When the margin on goods matters (trade, hospitality, production), describe it as "cost of goods as a % of revenue" and use the stream's goods-cost percentage.
- **Plain, factual, third person** ("the company", "the founders"). No marketing superlatives, no promises ("guaranteed", "certain success"). Short paragraphs.
- **Assumptions are the core of the plan** (art. 5:4, 6°). For each financial year, explain the drivers behind revenue (volumes × prices, ramp-up, seasonality) and the main cost changes. A court or bankruptcy trustee reads this to judge whether the starting capital was reasonable.
- **Risks** are concrete and come with a fallback the founders can actually apply.
- **Describe the activity at the level of the purpose clause.** What the company does, for whom, how it earns money, not the founder's current task list, workplace, tool names or how work is shared with others, unless it drives a number. Such details change every year and read as commitments.
- **No facts or commitments of third parties.** Don't state another company's revenue, forecasts or plans, other people's shareholdings or projects, or holdings that may change. Justify a fee or a contract from the company's own side (what it replaces, what it must cover), never from the client's figures.
- **"In the figures", not promises.** When the plan has one client or one product, write "the only client in the figures of this plan", not "the only client"; other activities that are possible but not modelled are "not included in the figures".
- **Snapshot.** The plan describes the situation and expectations on its date; the disclaimer says it is kept by the notary and not filed (art. 5:4, § 1 WVV). Don't add sentences that try to anticipate every future change.
- **Language.** Write in the plan language with its usual terms: NL *bestuurder, brutomarge, bedrijfswinst, boekjaar*; FR *administrateur, marge brute, bénéfice d'exploitation, exercice*; EN *director, gross margin, operating profit, financial year*.
