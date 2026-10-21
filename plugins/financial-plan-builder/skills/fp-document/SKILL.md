---
name: fp-document
description: Write the Word document (.docx) of a Belgian BV financial plan (art. 5:4 WVV) in Dutch, French or English, from plan-data.json and the recalculated Excel model. Drafts the owner text, builds the document with all tables and charts taken from the model, and handles checkpoints 3 and 4. Use after the model is approved (checkpoint 2), or when the owner wants to change the text of their financial plan.
---

# Financial plan: Word document

> **Plugin folder:** `${CLAUDE_PLUGIN_ROOT}` below is this plugin's folder. If it is not filled in (e.g. in Cowork or claude.ai), use the folder that holds `.claude-plugin/plugin.json` with `"name": "financial-plan-builder"`; this file is `skills/<skill>/SKILL.md` inside it (or `agents/fp-reviewer.md`). Find it once, e.g. `find / -path '*financial-plan-builder/skills/fp-model/scripts/build_model.py' 2>/dev/null`, and use that path for every script.

Produces the financial plan the founders hand to the notary. All tables, the cash chart and every number inside the text are read from the recalculated model, so text and numbers can't contradict each other. The owner gets the .docx (editable) and the Excel model.

## Requirements

- Python 3 with `python-docx` and `matplotlib` (`pip install python-docx matplotlib`), plus the fp-model requirements.
- A recalculated model (`fp-model` steps 2–3) without check errors. With a check error, only a draft can be built.

## Steps

`S` = `${CLAUDE_PLUGIN_ROOT}/skills/fp-document/scripts`, `<dir>` = the plan's working folder.

1. **Draft the owner text.** First run modules 11–12 of the interview (`fp-interview/references/question-bank.md`: narrative, risks and fallbacks, expert, place and date). Then fill `founders`, `activity`, `assumptions` and `document` in `plan-data.json`, following `references/document-text.md`. Use the interview log and the owner's documents (draft articles of association for `activity.purpose`). Write every model figure as a placeholder (`{{revenue.1}}`), never as a typed number.
2. **Build a draft:** `python3 S/build_document.py <dir>/plan-data.json <dir>/model.xlsx <dir>/draft.docx --draft`
   The JSON output lists `gaps` (missing text, unknown placeholders). Fix them in `plan-data.json` and rebuild.
   Then launch the **fp-reviewer** agent with `plan-data.json`, the model and the draft. Resolve every "Must fix" finding before checkpoint 3: text fixes in the owner-text fields, number fixes back to fp-model. Keep the "Confirm with the owner" questions for checkpoint 3.
3. **Checkpoint 3.** Show the owner the texts as they will read (placeholders filled in): activity (3.1–3.5), the generated financing statements, the year-by-year assumptions, profitability and risks. Ask: "Is this how you would explain your business to the notary?" Changes go into `plan-data.json`. Changes to numbers go back to fp-model.
4. **Build the final document:** the same command without `--draft`, output named in the plan language:
   - NL: `Financieel plan <Naam> BV.docx`
   - FR: `Plan financier <Nom> SRL.docx`
   - EN: `Financial plan <Name> BV.docx`

   Also copy the model next to it as `… - rekenmodel.xlsx` / `… - modèle de calcul.xlsx` / `… - calculation model.xlsx`.
   The final build refuses when text is missing, a placeholder is unknown or a model check is in error. Never work around this; fix the cause.
   Run the fp-reviewer agent once more on the final document; its verdict must be READY (or READY AFTER CONFIRMATIONS with the owner's answers recorded).
5. **Checkpoint 4.** Give the owner both files and say:
   - Word asks to update fields when the document opens: answer yes, so the table of contents gets page numbers.
   - The values marked "to verify" (disclaimer section) must be confirmed; offer the `fp-rules-check` skill if the rules are stale.
   - Recommend a review by an accountant before the notary appointment. If an accountant assists, add them as `document.expert` and rebuild.
   - The founders sign the last page; the notary keeps the plan.

## Rules

- Never edit the generated .docx to change content; change `plan-data.json` and rebuild. (After delivery, the owner may edit their copy.)
- Never type a model figure into the text; use a placeholder. Numbers the owner gave as drivers (days open, customers per day) may be written, and must match `plan-data.json`.
- Never mention a hire, investment, loan or activity that is not in the numbers.
- The disclaimer and sources section is always included and is not editable through `plan-data.json`.

## Files

- `scripts/build_document.py`: document generator (python-docx, matplotlib chart)
- `scripts/doc_i18n.py`: all fixed texts in NL/FR/EN
- `scripts/review_plan.py`: mechanical review (facts, filled-in texts, findings), the input for the `fp-reviewer` agent
- `references/document-text.md`: owner-text fields, placeholders and writing rules
