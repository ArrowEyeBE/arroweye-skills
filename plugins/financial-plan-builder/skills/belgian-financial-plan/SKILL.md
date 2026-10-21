---
name: belgian-financial-plan
description: Make the legal financial plan (financieel plan / plan financier, art. 5:4 WVV) for founders incorporating a Belgian BV (besloten vennootschap / SRL), in Dutch, French or English, from interview to the Word document and Excel model for the notary. Use when someone wants to draw up, continue or update the financial plan for incorporating a Belgian company, whatever the type of business (shop, webshop, restaurant, production, services, software).
---

# Financial plan for a Belgian BV

> **Plugin folder:** `${CLAUDE_PLUGIN_ROOT}` below is this plugin's folder. If it is not filled in (e.g. in Cowork or claude.ai), use the folder that holds `.claude-plugin/plugin.json` with `"name": "financial-plan-builder"`; this file is `skills/<skill>/SKILL.md` inside it (or `agents/fp-reviewer.md`). Find it once, e.g. `find / -path '*financial-plan-builder/skills/fp-model/scripts/build_model.py' 2>/dev/null`, and use that path for every script.

Guides the founders from first question to a signed-ready financial plan: interview → `plan-data.json` → Excel model → Word document, with four checkpoints where the owner approves the work. Each step is a skill of this plugin; this skill runs them in order and keeps track of where the plan is.

## Start

1. **Check the environment:** `python3 ${CLAUDE_PLUGIN_ROOT}/skills/belgian-financial-plan/scripts/check_env.py`. Nothing is ever installed without the user's yes. When something is missing, tell the user in their language what it is and what it is for, then:
   - **Python packages** (`"claude_may_run": true`): show the `install` command (and the `note`, if any) and ask whether you may run it. Only after a yes, run it exactly as printed, then run `check_env.py` again. If it fails, show the error and stop; don't try other install methods without asking.
   - **LibreOffice or Python itself** (`"claude_may_run": false`): the user installs these (a large download, sometimes an administrator password). Show the command or link. In Claude Code they can type `! <command>` to run it in this session. Run `check_env.py` again when they say it is done.
   - **`python3` not found:** the script cannot run. Python 3.9 or newer comes from python.org; on a Mac, `xcode-select --install` also provides it.

   Continue only when the status is `ok`. All scripts are started with `python3`; when the packages live in the plugin's own environment (see the `note`), the scripts switch to it by themselves.
2. **Working folder:** use the folder the user named; otherwise propose a new folder `financieel-plan-<company>` (or `plan-financier-…` / `financial-plan-…` in the user's language) in the current directory and confirm it. Everything for this plan lives there.
3. **New or resume:** if the folder has a `plan-data.json`, read `meta.progress.stage` and continue from the matching step below. Tell the owner in one sentence where they left off.

## Steps

| stage | step | skill | ends with |
|---|---|---|---|
| (none) / `interview` | Interview, modules 0–10 | `fp-interview` | **Checkpoint 1**: the inputs → `cp1_approved` |
| `cp1_approved` | Build, recalculate and check the model | `fp-model` | **Checkpoint 2**: the key figures → `cp2_approved` |
| `cp2_approved` | Modules 11–12 of the interview (narrative, risks, expert, place and date); draft the owner text; build a draft document; run the **fp-reviewer** agent | `fp-document` steps 1–2 | **Checkpoint 3**: the texts and the reviewer's questions → `cp3_approved` |
| `cp3_approved` | Final document; rules check if needed; fp-reviewer on the final document | `fp-document` steps 4–5, `fp-rules-check` | **Checkpoint 4**: the files → `final` |

After each checkpoint, set `meta.progress.stage` and `meta.progress.updated` in `plan-data.json`.

**Before checkpoint 4**, if `check_model.py` reports `legal_rules.stale` or values `to_verify`, offer the `fp-rules-check` skill (mode "for this plan"). The owner may decline; then the document's disclaimer lists the values to verify.

**Going back.** A change to the inputs at any later stage goes into `plan-data.json`, and everything after it is rebuilt: model, checks, document, review. Show the owner what changed in the key figures. Text approved at checkpoint 3 stays, but the reviewer runs again, because a placeholder can now say something the text no longer supports.

## At every checkpoint

- Show a short summary in the owner's language, not the files. Ask for approval or changes; never assume approval.
- One checkpoint at a time. Never skip one, even when the owner is in a hurry: each one catches a different kind of error.
- Say plainly what is a default, an estimate or still open.

## Rules (all steps)

- `plan-data.json` is the only source of truth. Never patch the Excel or Word output; change the data and rebuild.
- Never invent financing. A cash shortfall is shown to the owner with its size and month, and resolved with them.
- The plan is the founders' responsibility: this is not tax, legal or accounting advice. Recommend a review by an accountant before the notary appointment, and say so once at the start and once at the end.
- The Belgian rules come from `skills/fp-model/references/belgian-rules.json`, each with an official source and check date; never state a tax or payroll value from memory.
- Not only for services: follow the business types the owner chose.

## Deliverables (in the working folder)

- `Financieel plan <Naam> BV.docx` / `Plan financier <Nom> SRL.docx` / `Financial plan <Name> BV.docx`: the plan for the notary, editable
- the calculation model next to it (`… - rekenmodel.xlsx` / `… - modèle de calcul.xlsx` / `… - calculation model.xlsx`)
- `plan-data.json` and `interview-log.md`: keep them to update the plan later
