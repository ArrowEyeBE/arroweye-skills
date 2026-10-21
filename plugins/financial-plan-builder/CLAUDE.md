# financial-plan-builder

A Claude plugin that produces the legal financial plan (art. 5:4 WVV) for founders incorporating a Belgian BV: interview → `plan-data.json` → Excel model → Word document, in NL, FR or EN.

## Layout

- `skills/belgian-financial-plan/`: orchestrator (environment check, stages, checkpoints); entry point for the commands `/financieel-plan`, `/plan-financier`, `/financial-plan` in `commands/`
- `skills/fp-interview/`: interview skill, guide (`references/question-bank.md`), notes for management companies (`references/management-company.md`) and checklist writer
- `skills/fp-model/`: input validation (`validate_plan_data.py`), Excel model builder, recalculation, checks, Belgian rules
- `skills/fp-rules-check/`: re-checks the Belgian rules against official sources (per plan or for the plugin)
- `skills/fp-document/`: Word document generator, owner-text contract (`references/document-text.md`)
- `agents/fp-reviewer.md`: independent, read-only review (legal completeness, text versus numbers, plausibility); its mechanical part is `skills/fp-document/scripts/review_plan.py`
- `templates/cost-defaults.json`: indicative default cost lines per business type (proposed by `skills/fp-interview/scripts/cost_defaults.py`, always confirmed by the owner)
- `.claude-plugin/plugin.json`: manifest; `README.md`: user documentation. Validate with `claude plugin validate .`; try with `claude --plugin-dir .`
- `tests/cases/*/plan-data.json`: public, fictional test cases (keep them diverse: bakery NL, webshop FR, restaurant EN, management company NL with dividends and a fee step). The `privacy` test fails if a name or address of the private case appears in a public file.
- `reference/`: **private and gitignored**. The founder's real plan and the regression case built from it. Never copy names, numbers or text from here into `skills/`, `templates/` or `tests/`.

## Principles

- The Excel workbook is the only calculation engine. The Word document reads its numbers from the recalculated workbook; numbers inside owner text are placeholders (`{{revenue.1}}`), never typed.
- `plan-data.json` is the single source of truth: fix inputs there and regenerate; never patch outputs.
- Never invent financing. Cash shortfalls are flagged and resolved with the owner.
- Not service-company-centric: every feature must work for trade, production and hospitality too.

## Development

- Requirements: Python 3 with `openpyxl`, `python-docx`, `matplotlib`; LibreOffice.
- Run all tests: `python3 tests/run_tests.py` (or `--lang nl`). Cases with owner text (`activity`) also get a final Word document, checked for all legal chapters and the model's figures. Every case is built in NL, FR and EN, recalculated with LibreOffice, and must have zero formula errors and a balanced balance sheet. The private regression case must match the accountant's plan within the documented tolerances.
- New label → add it to `skills/fp-model/scripts/i18n.py` in all three languages.
- Belgian rules: values, official sources and check dates live in `skills/fp-model/references/belgian-rules.json` (explained in `belgian-rules.md`). Re-check them with the `fp-rules-check` skill before every release and at the start of every calendar year; never hard-code a legal value in the builder.
- Legal completeness: `belgian-rules.md` maps each element of art. 5:4 §1–§3 WVV to a document chapter; the tests check that all seven §2 references appear and that the plan covers two years after incorporation (`ck_horizon`). Keep that table in sync when chapters change.
- Every output carries a disclaimer: Read me sheet + source/date per value in Settings (done); the Word document must end with a disclaimer section before the signature page.
