---
name: fp-rules-check
description: Check whether the Belgian legal, tax and payroll values used by the financial-plan builder (corporate tax rates, minimum director pay for the reduced rate, employer contributions, VAT, company contribution, art. 5:4 WVV) are still current, against the official government sources. Use when the model reports the rules as stale or "to verify", before giving a plan to the notary, at the start of a new calendar year, or when the user asks whether the rules are up to date.
---

# Financial plan: check the Belgian rules

> **Plugin folder:** `${CLAUDE_PLUGIN_ROOT}` below is this plugin's folder. If it is not filled in (e.g. in Cowork or claude.ai), use the folder that holds `.claude-plugin/plugin.json` with `"name": "financial-plan-builder"`; this file is `skills/<skill>/SKILL.md` inside it (or `agents/fp-reviewer.md`). Find it once, e.g. `find / -path '*financial-plan-builder/skills/fp-model/scripts/build_model.py' 2>/dev/null`, and use that path for every script.

Laws, rates and thresholds change, sometimes mid-year or with retroactive effect. Every value the model uses is listed in `${CLAUDE_PLUGIN_ROOT}/skills/fp-model/references/belgian-rules.json`, with its official source, its status and the date it was last checked. This skill re-checks them.

## When to run it

- `check_model.py` reports `legal_rules.stale: true` (last check is older than `review_after_days` or from an earlier calendar year), or `legal_rules.to_verify` is not empty and the plan is about to be finalised (checkpoint 4).
- The owner or user asks: "are the rules still correct?", "is this up to date?".
- A plan starts in a calendar year after the `last_checked` date: indexed amounts usually change on 1 January.

## Two modes: ask which one applies

1. **For this plan** (default, for owners using the plugin): changes go into the plan's own `plan-data.json`. The plugin files are not touched, because a plugin update would overwrite them.
2. **For the plugin** (maintainer only, working in the plugin's source repository): update `belgian-rules.json` and `belgian-rules.md`, then run the tests.

## Steps

1. Read `belgian-rules.json`. List the rules with status `verify` first, then all others, with their current value and last-checked date.
2. For each rule, look up the current value **at the official source** in `sources` (`url` + `look_for`). Use web search to find the exact page when the URL is a portal. Treat secondary sources (accountancy firms, banks, news) only as pointers to the official text; say so when a value can only be confirmed by a secondary source.
3. Pay attention to the **date of effect**: a value can be voted but apply only from a later assessment year. The value the model needs is the one that applies to the plan's financial years. If different years need different values, report it. The model uses one value per rule, so pick the value for the first financial year and mention the difference as an assumption.
4. Report per rule in a table: rule · value in the model · value found · applies from · official source (link) · confidence (confirmed in the legal text / official explanatory page / only secondary sources).
5. **Never change a value without the user's explicit confirmation.** Propose the changes, then:
   - *For this plan:* write the new value to `plan-data.json` → `rules.<key>`, and the source to `rules_sources.<key>` = `{"source": "<official page or legal article>", "checked": "YYYY-MM-DD"}`. Rebuild the model (fp-model steps 2–4).
   - *For the plugin:* update `value`, `status`, `checked` (and `note` when useful) in `belgian-rules.json`, the matching line in `belgian-rules.md`, and `last_checked` when all rules were reviewed. Run `python3 tests/run_tests.py`; regression tolerances may need updating if a default changed, and that must be explained.
6. Items in `legal_items` (financial-plan content, alarm-bell procedure, VAT exemption, company contribution and its exemption, the self-employed contribution table, NBB schema) are not numbers in the model, but they shape the interview, the default cost lines and the document. The self-employed contribution table is published again every year: update its values and the source URL (it names the year). Check them too and report changes that would require changing the interview or the document text.

## Rules

- Official sources only for confirmation: Justel / Belgisch Staatsblad, FOD Financiën / Fisconetplus, RSZ (socialsecurity.be), RSVZ, FOD WASO, NBB, business.belgium.be.
- When the official source is unclear or the law is not yet final, keep the current value, set the status to `verify`, and say what is uncertain. Don't guess.
- This check does not replace an accountant. End the report with that reminder, especially for the reduced corporate tax rate conditions and payroll costs.
