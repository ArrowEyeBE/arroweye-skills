# Belgian rules used by the model

**Last checked: 2026-10-15.** These values change often, sometimes by laws still being voted.

- The values the model uses, their **official source** and their check date live in `belgian-rules.json`; this file explains them. Keep both in sync.
- Re-check with the `fp-rules-check` skill before every public release, at the start of every calendar year, and whenever `check_model.py` reports the rules as stale. Lines marked ⚠ have status `verify` in the JSON.
- The workbook shows the official source and check date next to every value (Settings sheet) and carries a disclaimer (Read me sheet). The Word document must carry the same disclaimer.
- In the plan itself, always say that tax and social-security figures are estimates to be confirmed by an accountant.

## Official sources

| topic | source |
|---|---|
| Financial plan, founders' liability, alarm-bell procedure (WVV art. 5:4, 5:16, 5:153) | [Justel: Wetboek van vennootschappen en verenigingen](https://www.ejustice.just.fgov.be/eli/wet/2019/03/23/2019A40586/justel) |
| Corporate tax rates and reduced-rate conditions (WIB 92 art. 215) | [FOD Financiën – Fisconetplus](https://www.fisconetplus.be/) |
| VAT rates, small-business exemption (BTW-Wetboek art. 56bis) | [FOD Financiën – btw](https://financien.belgium.be/nl/btw), legal text on Fisconetplus |
| Employer social-security contributions | [RSZ – administratieve instructies](https://www.socialsecurity.be/employer/instructions/dmfa/nl/latest/instructions/socialsecuritycontributions/contributions.html) |
| Wage indexation | [FOD WASO – automatische loonindexering](https://werk.belgie.be/nl/themas/verloning/automatische-loonindexering) |
| Company contribution and its exemption | [RSVZ – bedrag](https://www.rsvz.be/nl/faq/hoeveel-vennootschapsbijdrage-moet-ik-betalen) · [RSVZ – vrijstelling](https://www.rsvz.be/nl/faq/kan-ik-vrijgesteld-worden-van-het-betalen-van-de-vennootschapsbijdrage) |
| Social contributions of self-employed directors | [FOD Sociale Zekerheid – sociale bijdragen 2026 (tabel)](https://socialsecurity.belgium.be/sites/default/files/content/docs/nl/sociaal-beleid-vorm-geven/zelfstandigen/bijdragennota_2026_tabel_nl.pdf); a new table is published every year |
| Distributions (net-asset and liquidity test, art. 5:142–5:143 WVV) | [Justel: Wetboek van vennootschappen en verenigingen](https://www.ejustice.just.fgov.be/eli/wet/2019/03/23/2019A40586/justel) |
| Balance-sheet and P&L schema | [NBB – models of annual accounts](https://www.nbb.be/en/central-balance-sheet-office/preparation-and-filing/what-file/models/models-companies-standard-specific-consolidated) |

The links under "Sources" at the end are secondary explanations used to spot recent changes; confirm against the official source.

## Legal frame: the financial plan (art. 5:4 WVV / CSA)

Before incorporation, the founders give the notary a financial plan that justifies the amount of the starting equity (*aanvangsvermogen*) in light of the planned activity over at least two years (§1). It contains at least (§2):
1. a precise description of the planned activity
2. an overview of all sources of financing at incorporation, with the guarantees given for them where applicable
3. an opening balance sheet (schema of art. 3:3, i.e. the NBB annual-accounts schema) and projected balance sheets after 12 and 24 months
4. a projected P&L after 12 and 24 months, in the same schema
5. a budget of expected income and expenses for at least 2 years after incorporation
6. a description of the assumptions behind the estimates of revenue and profitability
7. where applicable, the name of the external expert who assisted

§3 allows another periodicity for 3° and 4° (the model uses financial years) **provided the projections together cover at least two years after incorporation**. The model checks this (`ck_horizon`), and the document states it when the first financial year is not 12 months.

| element | where in the document |
|---|---|
| §1 justification of the starting equity | 4 Financing: sufficiency (lowest cash, equity) |
| 1° | 3 Activity |
| 2° | 4 Financing (overview, loan table with guarantees, generated "no other sources" statements) |
| 3° / 4° | 5 Balance sheets, 6 P&L (NBB codes), §3 note when needed |
| 5° | 7 Budget (monthly cash flow over the whole plan; appendix) |
| 6° | 8 Assumptions |
| 7° | 9 External expert (always present; states when no expert assisted) |

**Contributions in kind (art. 5:7):** the founders write a special report describing and valuing each contribution in kind, and a company auditor (*bedrijfsrevisor / réviseur d'entreprises*) reports on it; both are filed with the deed. This is outside the financial plan, but the auditor's fee is a start-up cost and the timing affects the notary appointment. Exceptions (listed securities, assets valued by an auditor in the last 6 months, assets from audited accounts) are rare at a start-up.

The plan is not published: "Dit stuk wordt niet neergelegd met de akte, maar door de notaris bewaard" (§1). The document says so in chapter 10, with the statement that the plan reflects the situation on its date. It matters if the company goes bankrupt within 3 years: the founders can be held liable if the starting equity was clearly insufficient for the planned activity.
Hence the model's hard rules: never add financing the owner didn't state, and always flag cash shortfalls.

Also relevant: if net assets risk becoming negative, the directors must follow the alarm-bell procedure (*alarmbelprocedure*, art. 5:153 WVV). The model warns when equity < 0.

## Corporate tax (*vennootschapsbelasting*)

| item | value | model key | status |
|---|---|---|---|
| Standard rate | 25% | `corporate_tax_standard` | stable |
| Reduced SME rate | 20% on the first €100,000 of taxable profit | `corporate_tax_reduced`, `corporate_tax_reduced_bracket` | stable |
| Minimum director remuneration for the reduced rate | **€51,000** gross/year for assessment year 2027 (€50,000, indexed yearly; was €45,000), or at least equal to the taxable income if that is lower | `reduced_rate_min_director_pay` | ⚠ Law of 15 July 2026 (Belgian Official Gazette 29/07/2026), from assessment year 2027; the indexed amount comes from a secondary source (Attentia): confirm the official figure. |
| When the remuneration condition starts | from the **fifth taxable period** after incorporation (art. 215, third paragraph, 4° WIB 92: "vanaf het vijfde belastbaar tijdperk vanaf hun oprichting"); the first financial year is the first taxable period, whatever its length | `reduced_rate_pay_condition_from_period` = 5 | stable (checked on the consolidated text) |
| Benefits in kind | at most 20% of the remuneration may consist of flat-rate benefits (company car, housing, …) | — (not modelled; check manually) | ⚠ part of the same reform |
| Other reduced-rate conditions | small company (art. 1:24 WVV); not when at least half of the shares are held by one or more other companies (art. 215, third paragraph, 2°); not when the shares the company holds have an investment value above 50% of its revalued paid-up capital plus taxed reserves, participations of at least 75% not counted (1°) | — (asked in interview 10.1, not calculated) | check with the accountant; relevant when a management company will hold the shares of the operating company |
| Prepayments (*voorafbetalingen*) | no surcharge for insufficient prepayments in the first 3 financial years of a newly founded SME | `tax.prepayments` | verify conditions |
| Tax payment without prepayments | assessment usually arrives 6–12 months after year end; model default 9 months | `tax.payment_lag_months` | rule of thumb |

The model applies the reduced rate when the owner asks for it (`tax.reduced_rate`), in every profitable financial year before taxable period 5, and from period 5 on only where the remuneration condition holds: the highest remuneration of a single director line (category 618) must be ≥ MIN(minimum, taxable base). Otherwise it falls back to 25% and raises a warning. A plan covers 2–3 years, so for a starter the remuneration condition normally never applies within the plan; the other conditions are the owner's answers to question 10.1.
Remuneration for this condition is the director's taxable remuneration: gross pay and benefits (including social contributions the company pays on the director's behalf, booked on 618). A flat-rate expense allowance is a reimbursement of the company's costs, not remuneration (category 61). Confirm the split with the accountant.
Not modelled: disallowed expenses (e.g. part of restaurant and car costs), notional interest, investment deduction.

## VAT (*btw / TVA*)

| item | value | status |
|---|---|---|
| Standard rate | 21% | stable |
| Reduced rates | 6% (e.g. food, books, some renovation) · 12% (e.g. restaurant meals excl. drinks) | check the specific product |
| No VAT | e.g. rent of commercial property (unless opted), insurance, bank fees, director remuneration, levies | |
| Filing | quarterly for small companies (most starters), otherwise monthly | `company.vat_period_months` |
| Small-business exemption (*vrijstellingsregeling*) | turnover ≤ €25,000/year; no VAT charged, input VAT not deductible | ⚠ **not supported by the model yet**: all companies are treated as VAT-liable |
| VAT credit | the model carries a credit forward (no refund). In reality, starters can request refunds, so the model is conservative on cash. | simplification |

## Personnel

| item | value | model key | status |
|---|---|---|---|
| Employer social-security contributions | ~25% for white-collar employees (*bedienden*), higher for blue-collar workers (*arbeiders*, ~30.6% plus holiday-pay contribution). Model default 27% to cover small sector contributions. | `employer_ssc_rate` | ⚠ estimate; reforms in 2026 (wage-moderation contribution, structural reduction changes). A payroll simulation always beats this. |
| Salary months per year | 13.92 = 12 + double holiday pay (~0.92) + end-of-year bonus (1, depends on joint committee) | `salary_months_per_year` | rule of thumb |
| Wage indexation | ~2%/year, timing depends on joint committee | `wage_indexation` | rule of thumb |
| First hire | reduction of employer contributions for the first employee | — | ⚠ amounts were changed in 2024–2026; include via `employer_cost_monthly` from a payroll simulation |
| Student workers, flexi-jobs | reduced contributions; enter as `employer_cost_monthly` | — | |
| Mandatory once there is staff | workplace-accident insurance, payroll provider fee, occupational health service | — | add as cost lines |

## Company costs

| item | value | status |
|---|---|---|
| Annual company contribution (*vennootschapsbijdrage*) to the social insurance fund | €409.30 (balance total ≤ €879,163.44) or €1,022.38 (above) for 2026 | ⚠ A reform approved by the Council of Ministers (2026) replaces the 2 amounts with 4 levels. |
| Exemption for starting companies | the first 3 years of the company's existence, counted from the year of incorporation, for every company except an NV, registered in the KBO, when the directors (and the majority of the active partners who are not directors) were self-employed for no more than 12 of the 40 quarters before incorporation (side activity and helper included). `company.contribution_exempt` (question 2.9); the company requests it from its social insurance fund. | stable (RSVZ page, checked 2026-10-15) |

## Self-employed directors (*sociaal statuut der zelfstandigen*)

A director of a BV is self-employed and owes quarterly social contributions personally. A management company often pays them on the director's behalf (they are then part of the remuneration, category 618). `cost_defaults.py` proposes the line `dir_social` from these 2026 values (`legal_items.self_employed_contributions`, ⚠ new table every year):

| item | 2026 value (without the fund's administration costs) |
|---|---|
| Rate, main occupation | 20.50% of the net taxable professional income up to €75,024.54; 14.16% between €75,024.54 and €110,562.42 |
| Minimum income / minimum contribution per quarter | €17,374.08 / €890.42 |
| Primostarter (first 4 quarters as a starter) | minimum income €8,972.07, minimum contribution €459.82 per quarter |
| Side activity | no contribution below €1,922.16 income; minimum €98.51 per quarter |
| Administration costs of the social insurance fund | about 3–4% on top, depending on the fund; the default uses 3.05% |

Contributions of the first years are provisional (based on an estimate) and regularised later on the actual income: the plan uses the estimate.
| Filing the annual accounts with the NBB | small yearly fee (micro schema) | verify current tariff |
| Notary and registration costs | one-off; ask for a quote | |

## Depreciation (defaults)

Linear, pro rata temporis from the month of purchase (`depreciation_start_lag_months` = 0).

| asset | years |
|---|---|
| IT hardware, software | 3 |
| Vehicles | 5 |
| Machines, production equipment | 5–10 |
| Shop / office fit-out, furniture | 10 |
| Buildings | 33 |
| Rental deposit (category 28) | not depreciated |

## Sources (secondary)

- [Attentia: minimum director remuneration from assessment year 2027](https://www.attentia.be/nl/nieuws/verhoging-vereiste-minimale-bedrijfsleidersbezoldigingen-vanaf-aanslagjaar-2027)
- [Fiscosearch: consolidated art. 215 WIB 92](https://fiscosearch.be/collecties/wetboek-inkomstenbelastingen-1992/artikel-215)
- [Jubel: increase of the minimum director remuneration](https://www.jubel.be/verhoging-van-de-minimale-bezoldiging-van-bedrijfsleiders/)
- [BNP Paribas Fortis: reduced corporate tax rate and director pay in 2026](https://www.bnpparibasfortis.be/nl/public/article/bezoldiging-bedrijfsleider-arizona)
- [VLAIO: exemption from the company contribution](https://www.vlaio.be/nl/subsidies-financiering/subsidiedatabank/vrijstelling-van-vennootschapsbijdrage)
- [Partena Professional: reform of the annual company contribution from 2026](https://www.partena-professional.be/nl/node/22547)
- [Securex: basic employer contributions](https://www.securex.be/nl/lex4you/werkgever/themas/verlonen/sociale-loonlasten-rsz/patronale-basisbijdragen)
- [Accountable: VAT exemption threshold of €25,000](https://www.accountable.eu/nl-be/help-center/wanneer-wordt-een-zelfstandige-btw-plichtig-in-belgie-jouw-complete-gids-over-de-drempel-van-e25-000/)
