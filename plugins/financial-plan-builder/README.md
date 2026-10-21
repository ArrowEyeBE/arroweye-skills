# Financial plan builder for a Belgian BV

A Claude Code plugin that helps founders draw up the **financial plan** the notary needs to incorporate a Belgian BV / SRL (art. 5:4 of the Code of Companies and Associations, WVV / CSA), in **Dutch, French or English**.

You talk; Claude asks. The plugin interviews you about your company, builds a formula-driven Excel model, and writes the Word document with every element the law requires:

1. a precise description of the activity
2. all sources of financing, with the guarantees given
3. the opening balance sheet and projected balance sheets
4. the projected profit and loss accounts
5. a budget of income and expenses for at least two years
6. the assumptions behind revenue and profitability
7. the external expert, if one assisted

It works for any type of business: shop, webshop, wholesale, production, restaurant or café, services, software, and management companies (a BV that invoices a management fee to your operating company).

> **Not advice.** The plan is the founders' responsibility. The plugin is not tax, legal or accounting advice. The Belgian tax and payroll values it uses carry their official source and the date they were last checked; laws change. Have the plan reviewed by an accountant before the notary appointment.

## Requirements

In **Claude Cowork** nothing needs to be installed: Python, the packages and LibreOffice are already there.

In **Claude Code** on your own computer:
- Python 3.9+ with `openpyxl`, `python-docx` and `matplotlib`
- [LibreOffice](https://www.libreoffice.org/download/) (recalculates the Excel model; macOS: `brew install --cask libreoffice`)

The plugin checks these at the start and tells you what is missing. It never installs anything without your yes:
- **Python packages:** it shows the command and offers to run it for you. If your Python refuses `pip install` ("externally-managed-environment", common with Homebrew and Linux), the packages go into the plugin's own environment in `~/.financial-plan-builder/venv`, and the scripts use it by themselves. To remove that environment, delete the folder.
- **LibreOffice and Python:** you install these yourself. In Claude Code you can run the command it shows by typing `! <command>`.

## Install

This plugin is published in the public `arroweye-skills` repository of ArrowEye Automations. Add that marketplace once, then install:

```
/plugin marketplace add ArrowEyeBE/arroweye-skills
/plugin install financial-plan-builder@arroweye-skills
```

No GitHub account is needed. The plugin lives in that repository as `plugins/financial-plan-builder`; a zip of it is attached to every release. To try a local checkout: `claude --plugin-dir /path/to/arroweye-skills/plugins/financial-plan-builder`.

## Use

Start a plan, or continue one, with the command in your language:

- `/financieel-plan [map]`
- `/plan-financier [dossier]`
- `/financial-plan [folder]`

The first question is the language of the plan. Have your documents ready if you can (draft statutes, price list, quotes, lease, bank offer, payroll simulation): the plugin reads them first and asks only what is missing. It can also give you a checklist to prepare with your co-founders.

You approve the work at four checkpoints:

| | You see | You check |
|---|---|---|
| 1 | All inputs: revenue drivers, staff, costs, investments, financing | "Is this my business? Is anything missing?" |
| 2 | Key figures per year, the lowest cash point | "Does this look right? Is the money enough?" |
| 3 | The texts of the plan, plus an independent reviewer's questions | "Is this how I'd explain it to the notary?" |
| 4 | The final Word document and Excel model | Signatures, the values still to verify |

Everything stays in one folder on your computer: `plan-data.json` (your answers), `interview-log.md`, the Excel model and the Word document. Keep the folder to update the plan later.

## What it will not do

- **Invent money.** If cash runs out, the plugin tells you how much and when, and you decide how to cover it.
- **Hide anything.** It only states "there are no loans" (or grants, or contributions in kind) after you confirmed it.
- Not (yet) supported: other legal forms than the BV, the small-business VAT exemption, VAT refunds (VAT credits are carried forward), credit lines and capital increases during the plan. The plugin tells you when one of these applies.
- Dividends are modelled as the gross amount leaving the company, with the net-asset test of art. 5:142 WVV; the shareholder's tax (withholding tax, VVPRbis) is for your accountant.

## Keeping the rules up to date

Tax rates, thresholds and contributions are listed with their official source in `skills/fp-model/references/belgian-rules.json`. The plugin warns when they were last checked more than six months ago or in an earlier year, and the `fp-rules-check` skill re-checks them against the official sources (FOD Financiën / SPF Finances, RSZ / ONSS, RSVZ / INASTI, Justel, NBB).

## For developers

- `skills/belgian-financial-plan/`: the flow and its checkpoints
- `skills/fp-interview/`: interview, question bank, default cost lines (`templates/cost-defaults.json`)
- `skills/fp-model/`: validation, Excel model, recalculation, checks, Belgian rules
- `skills/fp-document/`: Word document and the mechanical review
- `skills/fp-rules-check/`: re-checking the Belgian rules
- `agents/fp-reviewer.md`: independent review
- Tests: `python3 tests/run_tests.py` (fictional cases in `tests/cases/`: bakery, webshop, restaurant, management company)
- Try it locally: `claude --plugin-dir /path/to/financial-plan-builder`

## License

Copyright © 2026 ArrowEye Automations BV. All rights reserved; see `LICENSE`.
