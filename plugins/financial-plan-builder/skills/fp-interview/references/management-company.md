# Management companies: what the plan can and cannot say

A management company (*managementvennootschap / société de management*) is a BV through which a founder carries out a director mandate and management services for one or a few related companies, often their own operating company, and invoices a management fee. It is one of the most common reasons to incorporate a BV. Business type `management`, revenue template J in `question-bank.md`.

The owner will ask questions that go beyond the financial plan (pay structure, holding the operating company's shares, dividends, tax). The plugin is not tax advice: answer what the plan needs, point to the official source in `belgian-rules.md`, and put the rest on the list for the accountant (an open item or a separate note). Never state a Belgian rule from memory.

## In the plan

| topic | how |
|---|---|
| Revenue | template J: unit "month", 1 per month, the fee as price, a `price_schedule` for each step; VAT 21% unless the accountant says otherwise |
| Justifying the fee | from the founder's side: what it replaces (what the client company pays the founder today), what it must cover (the director's pay and contributions, the company's own costs, a reserve). Not from the client's revenue or forecast: the plan must not state commitments or figures of a third party |
| Payment term | at invoice (0) is fine for a captive client: write the reason in `payment_note` |
| Director pay | from a pay slip when there is one (question 2.6): gross pay and company-paid social contributions in 618, allowances in 61. A step in pay → `schedule` |
| Social contributions | the `dir_social` default when the company pays them (computed from the 618 lines and the yearly table) |
| Company contribution | question 2.9: a starting company is often exempt for its first 3 years |
| Reduced corporate tax rate | question 10.1. The pay condition starts in taxable period 5; the shareholder and financial-company exclusions can apply to a holding structure |
| Dividends | `financing.dividends[]` (question 9.6); the model tests the net assets (art. 5:142 WVV) and shows the cash after payment (art. 5:143) |
| Market (3.4) | no market in the usual sense: describe the client relation briefly, or leave `activity.market` out |
| Other clients | "the only client in the figures of this plan"; possible other clients or mandates are not in the figures |

## Outside the plan: for the accountant

Note these when the owner raises them; the plan stays neutral about them.
- **Holding the operating company's shares** in the management company: affects the reduced corporate tax rate (art. 215, third paragraph, 1° and 2° WIB 92), the taxation of dividends received and of later capital gains, and needs a contribution or a sale of shares. Decide with the accountant before the shares are moved.
- **Dividend tax for the shareholder** (withholding tax rates, VVPRbis, liquidation reserve) and the timing of a short or long first financial year: the plan models the gross dividend only.
- **The management agreement** between the two companies: fee, payment term, notice, exclusivity. Draft it before the first invoice.
- **Benefits in kind and allowances**: which costs the company can pay or reimburse, and how they are taxed.
- **Social contributions**: provisional contributions are regularised on the actual income; the fund (and the primostarter rules) decide the exact amounts.

## Writing the texts

Follow the writing rules of `skills/fp-document/references/document-text.md`; for a management company in particular:
- Describe the activity at the level of the purpose clause: management services and director mandates, not the founder's current task list, workplace or tools.
- Do not describe other people's plans (a co-founder's own management company) or shareholdings that may change.
- The plan is a snapshot kept by the notary and not filed with the deed (art. 5:4, § 1 WVV): the document says so in chapter 10, so the text does not need to anticipate every future change.
