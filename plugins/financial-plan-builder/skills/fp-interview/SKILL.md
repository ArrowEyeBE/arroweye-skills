---
name: fp-interview
description: Interview the founders of a Belgian BV to collect everything the legal financial plan (art. 5:4 WVV) needs (company, founders, activity, revenue, purchasing and stock, staff, costs, investments, financing, tax), reading their documents first, and write it to plan-data.json. Ends with checkpoint 1. Use at the start of a financial plan, when resuming an unfinished interview, or when the owner wants to change the inputs of their plan.
---

# Financial plan: interview

> **Plugin folder:** `${CLAUDE_PLUGIN_ROOT}` below is this plugin's folder. If it is not filled in (e.g. in Cowork or claude.ai), use the folder that holds `.claude-plugin/plugin.json` with `"name": "financial-plan-builder"`; this file is `skills/<skill>/SKILL.md` inside it (or `agents/fp-reviewer.md`). Find it once, e.g. `find / -path '*financial-plan-builder/skills/fp-model/scripts/build_model.py' 2>/dev/null`, and use that path for every script.

Collects the inputs of the plan in a conversation and writes them to `plan-data.json`, the single source of truth for the model and the document. The content of the interview (modules, questions, checklists, sanity checks) is in `references/question-bank.md`; read it before starting. This file says how to run it and how answers become data.

## Files in the plan's working folder `<dir>`

| file | content |
|---|---|
| `plan-data.json` | the structured answers (`${CLAUDE_PLUGIN_ROOT}/skills/fp-model/references/plan-data-schema.md`) |
| `interview-log.md` | per module: the question, the owner's answer in their words, the source (document name or "owner"), and what was written to `plan-data.json` |
| `checklist.json` | only for the preparation checklist (0.2) |
| `documents/` | documents the owner provided (statutes, quotes, offers, payroll simulation) |

Write both files after **every module**, so the interview can be resumed at any time.

## Start or resume

- `plan-data.json` exists: read it and `interview-log.md`. `meta.progress.modules_done` lists the finished modules. Summarise in two lines what is known and continue with the first unfinished module. Don't ask again what is already answered.
- New plan: start with module 0. Question 0.1 (plan language) is always the first question; then the disclaimer (see the question bank); then the checklist (0.2). For "prepare first", fill the checklist in the plan language as JSON and run
  `python3 ${CLAUDE_PLUGIN_ROOT}/skills/fp-interview/scripts/write_checklist.py <dir>/checklist.json "<dir>/<Checklist name in the plan language>.docx"`, save progress and stop.

## How to ask

Follow "How to run the interview" in the question bank. In short:
- Documents first, then questions; at most 3–4 questions per turn; only the questions that fit the business types of 3.2.
- Talk in the owner's language; write data and log in a way the plan language can use.
- Never block on a number: propose a default with its basis, record it in `defaults_used`, and move on.
- Run the sanity checks of module 4 (capacity, growth without means, revenue before the activity can start) as soon as revenue and staff are known, and discuss what doesn't add up.

## From answers to plan-data.json

| module | writes |
|---|---|
| 0.1 | `meta.language`, `meta.schema_version: "0.1"` |
| 1 | `company.name, address, start_date, first_fy_end, vat_period_months`; 1.5 → `activity.purpose` (literally from the draft statutes) |
| 2 | `founders[]` (`name, role, address, birth_year, side_activity`); contributions → `financing.equity[]` (cash) or `investments[]` with `in_kind: true`; pay → `cost_lines` category `618` (a `schedule` when it changes; a pay slip split as in 2.6); loans → `financing.loans[]` with `type: "founder"`; 2.9 → `company.contribution_exempt` |
| 3 | `company.business_types`, `company.premises` (3.7), `company.food`; notes on offer, customers, channels, competition and evolution go to the log (the texts are drafted later, by fp-document) |
| 4 | `revenue_streams[]` (one per range × channel or service line): `unit, price_excl_vat, price_indexation` or `price_schedule, vat_rate, payment_lag_months` (+ `payment_note` when unusual), `volumes` |
| 5 | per stream: `goods_cost_pct` or `goods_cost_per_unit`, `other_var_cost_pct` (card and platform fees, commissions), `stock_cover_months`, `supplier_lag_months`, `purchase_vat_rate`; subcontracting that isn't per unit → `cost_lines` |
| 6 | `staff[]` (`role, contract, start, end, fte, gross_monthly` or `employer_cost_monthly`, `extras_monthly`); insurance, payroll service, occupational health → `cost_lines` |
| 7 | `cost_lines[]` with `source` (quote, contract, `default`); see "Default cost lines" below |
| 8 | `investments[]`; leasing → a monthly `cost_line` instead |
| 9 | `financing.equity, loans, grants, dividends, unavailable_equity`, and `financing.confirmed_none` |
| 10 | `tax.reduced_rate, prepayments` |
| any | unknowns → `open_items[]` (a sentence each); defaults → `defaults_used[]` |

Modules 11 (assumptions narrative) and 12 (expert, place and date) come after checkpoint 2, when the numbers are known; they feed fp-document.

### Conversion rules

- **Amounts exclude VAT.** B2C prices are often quoted incl. VAT: divide by (1 + rate) and say so.
- **Volumes are monthly.** For yearly or quarterly targets, ask how the year ramps up (linear, a step in month X, seasonal) and spread them; show the months back. Write the driver behind the volumes ("80 → 165 customers a day × 26 opening days") in the log: the narrative quotes it later and the reviewer checks it against the volumes.
- **Unit "€"** only when the owner can't give a driver; then `price_excl_vat: 1` and volumes are amounts. Always try a driver first (days × rate, projects × price): a plan with drivers is far stronger for art. 5:4, §2, 6°.
- **Opening stock**: the model builds stock from `stock_cover_months` (months of next month's cost of goods). Convert an opening-stock amount into months of cover and confirm the result with the owner.
- **Payment terms** in whole months: 30 days → 1; at the till → 0.
- **Staff**: a payroll simulation goes into `employer_cost_monthly` (total monthly cost; also set `fte`). Otherwise `gross_monthly` (full-time) and `fte`; the model applies the employer-cost formula. Always set `fte` for part-time, students and flexi-jobs.
- **Dates** within the plan: nothing before `start_date` (costs made before incorporation are paid by the founders, unless the company takes them over at its start: then put them in the first month).
- **Small-business VAT exemption** (*vrijstellingsregeling*) is not supported by the model. If the owner wants it, say so, put it in `open_items` and agree how to approximate it (all VAT rates 0 and costs incl. VAT).
- **Financing: only what the owner states.** For each of loans (2.7, 9.2–9.3), grants (9.4) and contributions in kind (2.4) the answer is either an entry or an explicit "none", recorded in `financing.confirmed_none`. The document prints "there are no loans" from that confirmation.
- **Dividends** (9.6) go into `financing.dividends[]` (`fy`, gross `amount`, payment `date`). Capital increases and credit lines (9.5, 9.8) are not modelled: record them in `open_items` and tell the owner.
- **Company name** without the legal form ("Peeters Consult", not "Peeters Consult BV"): the document adds it.
- **Purpose clause** drafted without the notary's draft articles (1.5): set `activity.purpose_status` (e.g. "draft by the founder, to be aligned with the notary"). The document prints it as a note under 3.1; don't add such a sentence inside the clause.
- **Management company** (type `management`): use template J, the module skips of "How to run the interview", and `references/management-company.md` for the questions the owner will ask that the plan cannot answer.

### Default cost lines (module 7)

Run `python3 ${CLAUDE_PLUGIN_ROOT}/skills/fp-interview/scripts/cost_defaults.py <dir>/plan-data.json` once modules 3, 6 and 8 are known. It lists the default lines that apply to this business (premises, staff, vehicles, food, contribution in kind, business type) and are not yet in the plan, in the plan language, with an indicative amount, a range and its basis (`templates/cost-defaults.json`).
- Show them grouped, a few at a time, with amount and range: the owner confirms, changes or declines each one. Lines marked `ask` have no usable default (rent, the company contribution exemption, a webshop platform): ask.
- Add accepted lines from their `entry`. Keep the id in `defaults_used` while the amount is the default; replace the `source` with the quote or contract when the owner has one.
- Declined lines: note them in the log, so the reviewer and the owner can see they were considered.

## Checkpoint 1

When modules 0–10 are done:

1. Validate: `python3 ${CLAUDE_PLUGIN_ROOT}/skills/fp-model/scripts/validate_plan_data.py <dir>/plan-data.json`. Fix every error (ask the owner where needed); keep the warnings for the summary.
2. Show the owner one compact summary, in the owner's language:
   - company, start date, financial years
   - per revenue stream: the driver line (volume × price, ramp-up, payment term, cost of goods)
   - staff with start months and cost; director pay per period
   - the largest cost lines, investments with month and amount
   - financing, including every confirmed "none"
   - every value marked `default`, the open items and the validator's warnings
3. Ask: "Is this right? Anything missing?" Process changes and repeat until the owner approves.
4. Set `meta.progress.stage` to `"cp1_approved"` and hand over to fp-model.

## Rules

- Never invent financing, revenue or a cost line the owner didn't give or accept as a default.
- Never assume a services business; the questions follow the business types of 3.2.
- Never put a guess in `plan-data.json` without marking it in `defaults_used` or `open_items`.
- Personal data of the founders (address, birth year) is only used for the plan; don't put it anywhere else.
