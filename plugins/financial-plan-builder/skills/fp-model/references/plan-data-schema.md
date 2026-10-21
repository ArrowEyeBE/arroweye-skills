# plan-data.json: the single source of truth

The interview writes this file; `build_model.py` turns it into the Excel model; the Word document is generated from the recalculated model.
Never patch the Excel or Word output to fix a number. Change `plan-data.json` and regenerate.

Schema version: `0.1`. Dates are `"YYYY-MM-DD"` or `"YYYY-MM"` (a month is enough for anything except `company.start_date`). Amounts are in EUR **excluding VAT**. Rates are decimals (`0.21` = 21%).

## Top level

```json
{
  "meta": {...},
  "company": {...},
  "founders": [...],
  "activity": {...},
  "revenue_streams": [...],
  "cost_lines": [...],
  "staff": [...],
  "investments": [...],
  "financing": {"equity": [...], "loans": [...], "grants": [...], "dividends": [...]},
  "tax": {...},
  "rules": {...},
  "assumptions": {...},
  "document": {...},
  "defaults_used": [...],
  "open_items": [...]
}
```

`founders`, `activity`, `assumptions` and `document` hold the owner text of the Word document; they are not used by the model. Their fields, the number placeholders and the writing rules are in `skills/fp-document/references/document-text.md`. `open_items` lists what is still unknown (shown at the checkpoints).

## meta

| field | type | notes |
|---|---|---|
| `language` | `"nl"` \| `"fr"` \| `"en"` | Document **and** workbook language (interview question 0.1) |
| `schema_version` | string | `"0.1"` |
| `progress` | object | written by the interview and the `belgian-financial-plan` skill: `{"stage": ..., "modules_done": [0, 1, ...], "updated": "YYYY-MM-DD"}`. Stages: `interview`, `cp1_approved`, `cp2_approved`, `cp3_approved`, `final`. Not used by the model. |

## company

| field | required | notes |
|---|---|---|
| `name` | yes | without the legal form ("Peeters Consult", not "Peeters Consult BV"): the document adds "BV" / "SRL" |
| `legal_form` | yes | always `"BV"` for now |
| `address` | yes | registered office |
| `enterprise_number` | no | only when already assigned |
| `start_date` | yes | `YYYY-MM-DD`; first month of the model |
| `first_fy_end` | yes | end of the first financial year (may be longer or shorter than 12 months) |
| `fiscal_years` | no | number of financial years in the plan, default 3. The last one must end at least 2 years after `start_date` (art. 5:4 WVV; `ck_horizon`), so 2 is only enough when the plan starts on the first day of a financial year. |
| `vat_period_months` | no | 1 (monthly) or 3 (quarterly), default 3 |
| `business_types` | yes | any of `trade`, `production`, `hospitality`, `services`, `software`, `management`, `other`. `management`: a management or holding company invoicing related companies (template J) |
| `premises` | no | `"none"`, `"rented"` or `"owned"` (question 3.7); drives the default cost lines for premises |
| `food` | no | `true` for businesses selling food (FAVV/AFSCA); hospitality counts as food anyway |
| `vehicle` | no | `true` when a vehicle is used but not bought (lease); bought vehicles are recognised from `investments` |
| `contribution_exempt` | no | question 2.9: `true` when every director was self-employed for no more than 12 of the 40 quarters before incorporation (no company contribution for the first 3 years from the year of incorporation); `false` when not; absent = not asked yet |

## revenue_streams[]

One per product range × channel or per service line (interview module 4). Every template in the question bank is reduced to **volume per month × price per unit**:

| template | `unit` | volume = |
|---|---|---|
| A shop/webshop | `"customers"` or `"orders"` | customers or orders per month (price = average basket excl. VAT) |
| B wholesale | `"orders"` | orders per month (price = average order value) |
| C production | `"units"` | units sold per month |
| D hospitality | `"covers"` / `"room nights"` | guests per month (price = average spend excl. VAT) |
| E time-based | `"days"` / `"hours"` | billable days per month (all people together) |
| F projects | `"projects"` | projects invoiced per month (fraction allowed: 0.5 = one every 2 months) |
| G subscription | `"subscribers"` | active subscribers per month (after churn) |
| H training/events | `"participants"` | participants per month |
| J management fee | `"month"` | 1 per month (price = the monthly fee; steps in `price_schedule`) |
| when only an amount is known | `"€"` | revenue per month, price = 1 |

The interview converts yearly or quarterly answers into monthly volumes (ramp-up and seasonality included) and shows the result to the owner. The drivers behind the volumes (e.g. "120 customers/day × 26 days") go into `assumptions`, so the narrative can quote them.

| field | default | notes |
|---|---|---|
| `id`, `name`, `unit` | — | `id` is used in `defaults_used` |
| `price_excl_vat` | — | price per unit in month 1 |
| `price_indexation` | 0 | price increase per financial year |
| `price_schedule` | — | price steps on known dates: `{"2027-04": 3000}` = this price from that month on (until the next step). Monthly input prices in the model; not combined with `price_indexation` (put indexed prices in the schedule). Not for unit `"€"` |
| `vat_rate` | 0.21 | VAT on sales |
| `payment_lag_months` | 0 | months between invoice and payment (0 = at the till or at invoice) |
| `payment_note` | — | the owner's reason for an unusual term (e.g. "related client, pays at invoice"); shown to the reviewer |
| `goods_cost_pct` | 0 | cost of goods as % of revenue (goes through stock, NBB code 60) |
| `goods_cost_per_unit` | 0 | cost of goods per unit (alternative or addition to the %) |
| `other_var_cost_pct` | 0 | card/platform fees, commissions, subcontracting as % of revenue (code 61) |
| `other_var_vat_rate` | 0.21 | |
| `stock_cover_months` | 0 | stock kept, in months of next month's cost of goods |
| `supplier_lag_months` | 1 if goods costs, else 0 | |
| `purchase_vat_rate` | = `vat_rate` | |
| `volumes` | — | `{"2027-01": 2080, "2027-02": 2340, ...}`; missing months = 0 |

## cost_lines[]

| field | default | notes |
|---|---|---|
| `id`, `name` | — | |
| `category` | — | `61` services & other goods (incl. a director's expense allowance) · `618` director remuneration: gross pay and social contributions the company pays on the director's behalf (part of 61; used for the reduced-tax-rate check) · `640` other operating charges (small taxes, levies) · `65` financial charges (bank fees) |
| `amount` | — | per occurrence, excl. VAT |
| `interval_months` | 1 | 1 monthly · 3 quarterly · 12 yearly · 0 once |
| `start` | — | month of the first occurrence |
| `end` | none | last month (inclusive) |
| `vat_rate` | 0.21 | 0 for insurance, rent without VAT, director pay, bank fees, levies |
| `indexation` | 0 | increase per financial year |
| `schedule` | — | irregular amounts: `{"2028-01": 3000, ...}`; replaces `amount`/`interval_months` |
| `source` | — | where the number came from (quote, contract, default) |

## staff[]

| field | notes |
|---|---|
| `id`, `role`, `contract` | contract: employee type in plain words (bediende, arbeider, student, flexi, …) |
| `start`, `end` | months |
| `fte` | default 1. Always set it for part-time, students and flexi-jobs (hours per week / full-time hours): it drives the FTE figures even when `employer_cost_monthly` is given |
| `gross_monthly` | full-time gross salary; employer cost = gross × FTE × `salary_months_per_year`/12 × (1 + `employer_ssc_rate`) |
| `employer_cost_monthly` | total monthly employer cost from a payroll simulation; overrides the formula (not multiplied by FTE) |
| `extras_monthly` | meal vouchers, phone, … |

## investments[]

| field | default | notes |
|---|---|---|
| `id`, `name` | — | |
| `category` | 22 | `21` intangible · `22` tangible (22/27) · `28` financial fixed assets (rental deposit) |
| `amount`, `date` | — | |
| `life_years` | 3 | depreciation period; 0 for deposits |
| `vat_rate` | 0.21 | 0 for deposits and contributions in kind |
| `in_kind` | false | contributed by a founder in kind: no cash, adds to equity |

## financing

- `equity[]`: `{name, amount, date}` cash contributions (contributions in kind are investments with `in_kind: true`).
- `loans[]`: `{lender, type: "bank"|"founder"|"other", amount, date, rate, term_months, grace_months, guarantees}`. Annuity repayment after `grace_months` (interest-only); `term_months` includes the grace period. **Only what the owner stated: never add a loan to make cash work.** Several tranches = several loans.
- `grants[]`: `{name, amount, date}` booked as other operating income (74) when received.
- `dividends[]`: `{name?, fy, amount, date}`: gross dividend from the result of financial year `fy` (1 = first), paid in month `date`. Decided at the end of that financial year (shown as a debt, code 47/48, on the balance sheet after appropriation until paid), or at payment when paid during the year itself (interim dividend). Lowers equity and cash; `ck_dividend` tests the net assets after each decision (art. 5:142 WVV).
- `unavailable_equity`: amount of equity the articles make unavailable (default 0); used by `ck_dividend`.
- `confirmed_none`: which of `"loans"`, `"grants"`, `"in_kind"` the owner explicitly confirmed they don't have. Required for every one that is empty: the document states "there are no loans / grants / contributions in kind", so that must be the owner's answer, not a missing one. `validate_plan_data.py` and the final document build enforce this.

## tax

`{"reduced_rate": bool, "prepayments": bool, "payment_lag_months": 9}`. The model checks the remuneration condition itself (from taxable period 5) and falls back to the standard rate when it is not met. The other conditions (shareholders, financial company) are the owner's answer to question 10.1.

## rules (optional overrides)

Keys: `corporate_tax_standard`, `corporate_tax_reduced`, `corporate_tax_reduced_bracket`, `reduced_rate_min_director_pay`, `reduced_rate_pay_condition_from_period`, `employer_ssc_rate`, `salary_months_per_year`, `wage_indexation`, `depreciation_start_lag_months`, `cash_buffer_months`. Defaults, official sources and check dates: `belgian-rules.json` (explained in `belgian-rules.md`).

`rules_sources` (optional): where an overridden value comes from, shown in the Settings sheet instead of the registry source, e.g.
`{"reduced_rate_min_director_pay": {"source": "art. 215 WIB 92 (Fisconetplus)", "checked": "2027-01-10"}}`. Written by the `fp-rules-check` skill or when the owner's accountant supplies a value.

## defaults_used[]

Every value the owner did not give explicitly: item `id`s, `"<stream_id>.<field>"`, `"rules.<key>"`, `"tax.<key>"`, `"company.vat_period_months"`. These cells get a yellow background in Excel and are listed at checkpoint 1.

## Test-only fields

- `expected`: reference values per financial year for regression tests (`pl` / `bs` row keys from the model's `_map` sheet).
- `expect_status`: expected check statuses, e.g. `{"ck_cash": "error"}`.
