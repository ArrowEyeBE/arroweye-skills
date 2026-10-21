---
name: fp-model
description: Build the Excel calculation model of a Belgian BV financial plan from plan-data.json, recalculate it, run the built-in checks and present the key figures to the owner (checkpoint 2). Use after the interview is approved (checkpoint 1), whenever plan-data.json changes, or when the owner asks to see or change the numbers of their financial plan.
---

# Financial plan: Excel model

> **Plugin folder:** `${CLAUDE_PLUGIN_ROOT}` below is this plugin's folder. If it is not filled in (e.g. in Cowork or claude.ai), use the folder that holds `.claude-plugin/plugin.json` with `"name": "financial-plan-builder"`; this file is `skills/<skill>/SKILL.md` inside it (or `agents/fp-reviewer.md`). Find it once, e.g. `find / -path '*financial-plan-builder/skills/fp-model/scripts/build_model.py' 2>/dev/null`, and use that path for every script.

Turns `plan-data.json` (written by the interview, see `references/plan-data-schema.md`) into a fully formula-driven Excel workbook in the plan's language (NL/FR/EN). The workbook is the **only calculation engine**: the Word document later reads every number from it, so text, tables and charts cannot drift apart.

## Requirements

- Python 3 with `openpyxl`
- LibreOffice (for recalculation). If `recalc.py` reports it missing, ask the user to install it (macOS: `brew install --cask libreoffice`) and stop.

## Steps

`S` = `${CLAUDE_PLUGIN_ROOT}/skills/fp-model/scripts`. `<dir>` = the plan's working folder (where `plan-data.json` lives).

1. **Validate the input:** `python3 S/validate_plan_data.py <dir>/plan-data.json`. Any error: go back to the interview (fp-interview) and ask the owner; don't invent values.
2. **Build:** `python3 S/build_model.py <dir>/plan-data.json <dir>/model.xlsx`
3. **Recalculate:** `python3 S/recalc.py <dir>/model.xlsx`. This replaces the file with a recalculated version that shows values in every viewer.
   Status must be `success`. Formula errors are a bug in the builder, not something to patch by hand: report them and stop.
4. **Check:** `python3 S/check_model.py <dir>/model.xlsx` prints JSON with the checks, the key figures per financial year and the lowest cash month.
5. **Handle the checks** before showing anything to the owner:
   - `ck_balance` error: builder bug. Stop and report it.
   - `ck_cash` error (cash goes negative): **do not change the financing yourself.** Tell the owner how much money is missing at the lowest point and in which month, and what causes it (e.g. investment, stock purchase, slow-paying customers). Offer options: more equity, a loan (with amount, rate and term from the owner), postponing an investment or hire, lower costs. Update `plan-data.json` only with what the owner decides, then rebuild.
   - `ck_buffer` warning: mention it at the checkpoint as a risk; no change needed unless the owner wants one.
   - `ck_reduced` warning: the reduced tax rate is asked for, but director pay is too low in a year from the fifth taxable period on (the pay condition does not apply before; see `references/belgian-rules.md`). Explain the trade-off neutrally and let the owner choose.
   - `ck_equity` warning: negative equity at a year end. Mention the alarm-bell procedure; suggest a review by an accountant.
   - `ck_dividend` error: a dividend would bring equity below zero or below the unavailable equity (art. 5:142 WVV). Tell the owner which dividend and by how much; the owner lowers or postpones it. Also show the lowest cash in the 12 months after each payment (art. 5:143 is the board's judgement).
   - `ck_horizon` error: the financial years end before two years after incorporation (art. 5:4 WVV). Increase `fiscal_years` and rebuild; no owner decision needed.
   - `legal_rules.stale` true: the Belgian rules were last checked too long ago or in an earlier calendar year. Say so at the checkpoint and offer to run the `fp-rules-check` skill before the plan is finalised.
6. **Checkpoint 2.** Show the owner a short summary, not the workbook:
   - a table per financial year: revenue, gross margin (and %), operating result, net result, cash at year end, lowest cash, equity, staff in FTE
   - the lowest cash point (amount + month) in one sentence
   - check results that are not OK, in plain language
   - every value still marked as a default (yellow in Excel), from `defaults_used`
   - one line on the legal values: "Belgian tax and payroll rules last checked on <legal_rules.last_checked>", and the rules in `legal_rules.to_verify` in plain words (e.g. "the minimum director pay for the 20% rate is a draft law: confirm with your accountant"). Never present a tax value as certain.

   Ask: "Does this look like your business?" Changes go into `plan-data.json`, then repeat from step 2.

## Rules

- Never edit `model.xlsx` to fix a number; change `plan-data.json` and rebuild. (The owner may edit the Excel themselves later. That is their copy.)
- Never add financing, revenue or cost the owner did not state or confirm as a default.
- Amounts in `plan-data.json` exclude VAT; the model adds VAT itself.
- Every legal or tax value comes with its official source and check date (Settings sheet). In conversation and in the document, call them estimates to be confirmed by an accountant; the plan is not tax or legal advice.
- When describing results, use the P&L terms with their NBB codes (e.g. "bedrijfswinst (9901)") only in the Word document; in conversation, use plain words.

## Files

- `scripts/build_model.py`: builder (sheets: Read me, Key figures, Checks, Settings, Inputs, Revenue, Costs, Investments, Financing, Calculation, Cash flow, P&L, Balance sheet, Timeline, plus a hidden `_map` sheet)
- `scripts/i18n.py`: all labels and sheet names in NL/FR/EN
- `scripts/recalc.py`: LibreOffice recalculation + formula-error scan
- `scripts/validate_plan_data.py`: input validation (also run at checkpoint 1)
- `scripts/check_model.py`, `scripts/model_io.py`: read values back via the `_map` sheet
- `references/plan-data-schema.md`: input contract
- `references/belgian-rules.json`: the default value, official source, status and last-checked date of every legal and tax rule. The builder reads its defaults from here and shows the source and date in the Settings sheet.
- `references/belgian-rules.md`: explanation of those rules. To re-check them, use the `fp-rules-check` skill.
