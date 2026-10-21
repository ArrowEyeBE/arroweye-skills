#!/usr/bin/env python3
"""Validate plan-data.json before the model is built (end of the interview, checkpoint 1).

Usage: validate_plan_data.py <plan-data.json>

Prints JSON {status: ok|warnings|errors, errors: [...], warnings: [...]}; each item is {path, message}.
Errors: the model can't be built, or would be wrong (missing or impossible values, dates outside the plan,
financing that was never asked about). Warnings: things to show the owner at checkpoint 1.
Exit code 1 when there are errors.
"""
import calendar
import json
import re
import sys
from datetime import date, timedelta
from pathlib import Path

LANGS = {"nl", "fr", "en"}
BUSINESS_TYPES = {"trade", "production", "hospitality", "services", "software", "management", "other"}
GOODS_TYPES = {"trade", "production", "hospitality"}
VAT_RATES = {0, 0.06, 0.12, 0.21}
COST_CATEGORIES = {61, 618, 640, 65}
INVEST_CATEGORIES = {21, 22, 28}
INTERVALS = {0, 1, 3, 6, 12}
LOAN_TYPES = {"bank", "founder", "other"}
CONFIRMABLE = {"loans", "grants", "in_kind"}
MONTH = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
DAY = re.compile(r"^\d{4}-(0[1-9]|1[0-2])-\d{2}$")
LEGAL_FORM = re.compile(r"[\s,]+(BV|B\.V\.|SRL|S\.R\.L\.|BVBA|SPRL)\.?\s*$", re.I)
RULES = json.loads((Path(__file__).parent.parent / "references/belgian-rules.json").read_text())


class Validator:
    def __init__(self, d):
        self.d = d
        self.errors, self.warnings = [], []

    def err(self, path, msg):
        self.errors.append({"path": path, "message": msg})

    def warn(self, path, msg):
        self.warnings.append({"path": path, "message": msg})

    # ---- helpers
    def month(self, v, path, required=True):
        """'YYYY-MM' or 'YYYY-MM-DD' → (y, m), or None."""
        if v in (None, ""):
            if required:
                self.err(path, "missing")
            return None
        s = str(v)
        if not (MONTH.match(s) or DAY.match(s)):
            self.err(path, f"'{s}' is not a date (YYYY-MM or YYYY-MM-DD)")
            return None
        if DAY.match(s):
            try:
                date.fromisoformat(s)
            except ValueError:
                self.err(path, f"'{s}' is not a valid date")
                return None
        return int(s[:4]), int(s[5:7])

    def in_plan(self, ym, path, what="date"):
        if ym and self.start and self.end:
            if ym < self.start:
                self.err(path, f"{what} {ym[0]}-{ym[1]:02d} is before the start of the plan ({self.start[0]}-{self.start[1]:02d})")
            elif ym > self.end:
                self.warn(path, f"{what} {ym[0]}-{ym[1]:02d} is after the last financial year: it is ignored by the model")

    def number(self, x, key, path, required=True, lo=None, hi=None, integer=False):
        v = x.get(key)
        if v is None:
            if required:
                self.err(f"{path}.{key}", "missing")
            return None
        if isinstance(v, bool) or not isinstance(v, (int, float)):
            self.err(f"{path}.{key}", f"'{v}' is not a number")
            return None
        if integer and v != int(v):
            self.err(f"{path}.{key}", f"{v} must be a whole number")
        if lo is not None and v < lo:
            self.err(f"{path}.{key}", f"{v} is below {lo}")
        if hi is not None and v > hi:
            self.err(f"{path}.{key}", f"{v} is above {hi}")
        return v

    def vat(self, x, key, path):
        if key in x and x[key] not in VAT_RATES:
            self.err(f"{path}.{key}", f"{x[key]} is not a Belgian VAT rate (0, 0.06, 0.12, 0.21)")

    def ids(self, items, path):
        seen = set()
        for i, x in enumerate(items):
            if not x.get("id"):
                self.err(f"{path}[{i}].id", "missing")
            elif x["id"] in seen:
                self.err(f"{path}[{i}].id", f"duplicate id '{x['id']}'")
            seen.add(x.get("id"))

    # ---- sections
    def meta(self):
        m = self.d.get("meta") or {}
        if m.get("language") not in LANGS:
            self.err("meta.language", "must be nl, fr or en (interview question 0.1)")

    def company(self):
        c = self.d.get("company") or {}
        self.start = self.end = None
        for k in ("name", "address"):
            if not c.get(k):
                self.err(f"company.{k}", "missing")
        if c.get("name") and LEGAL_FORM.search(c["name"]):
            self.warn("company.name", "ends with the legal form; leave it out: the document adds 'BV' / 'SRL' itself")
        if c.get("legal_form", "BV") != "BV":
            self.err("company.legal_form", "only BV is supported")
        sd = c.get("start_date")
        if not sd or not DAY.match(str(sd)):
            self.err("company.start_date", "missing or not YYYY-MM-DD")
        else:
            self.start = self.month(sd, "company.start_date")
        fy1 = self.month(c.get("first_fy_end"), "company.first_fy_end")
        nfy = c.get("fiscal_years", 3)
        if not isinstance(nfy, int) or not 1 <= nfy <= 5:
            self.err("company.fiscal_years", "must be a whole number from 1 to 5")
            nfy = 3
        self.nfy = nfy
        if fy1 and self.start:
            months = (fy1[0] - self.start[0]) * 12 + fy1[1] - self.start[1] + 1
            if months < 1:
                self.err("company.first_fy_end", "is before the start date")
            elif months > 24:
                self.warn("company.first_fy_end", f"first financial year of {months} months: confirm with the notary")
            if DAY.match(str(c["first_fy_end"])) and int(str(c["first_fy_end"])[8:]) != calendar.monthrange(*fy1)[1]:
                self.warn("company.first_fy_end", "is not the last day of a month: the model treats it as the end of that month")
            self.end = (fy1[0] + nfy - 1, fy1[1])
            # same rule as the model's ck_horizon: last year end ≥ start + 2 years − 1 day
            s = date.fromisoformat(sd)
            two_years = date(s.year + 2, s.month, min(s.day, calendar.monthrange(s.year + 2, s.month)[1])) - timedelta(days=1)
            if date(self.end[0], self.end[1], calendar.monthrange(*self.end)[1]) < two_years:
                self.err("company.fiscal_years", "the plan must cover at least 2 years after incorporation (art. 5:4 WVV): add a financial year")
        if c.get("premises") not in (None, "none", "rented", "owned"):
            self.err("company.premises", "must be none, rented or owned")
        for k in ("food", "vehicle", "contribution_exempt"):
            if k in c and not isinstance(c[k], bool):
                self.err(f"company.{k}", "must be true or false")
        if c.get("vat_period_months", 3) not in (1, 3):
            self.err("company.vat_period_months", "must be 1 or 3")
        bt = c.get("business_types") or []
        if not bt:
            self.err("company.business_types", "missing (interview question 3.2)")
        for t in bt:
            if t not in BUSINESS_TYPES:
                self.err("company.business_types", f"unknown type '{t}'")
        self.types = set(bt)

    def founders(self):
        f = self.d.get("founders") or []
        if not f:
            self.err("founders", "at least one founder is needed (module 2)")
        for i, x in enumerate(f):
            if not x.get("name"):
                self.err(f"founders[{i}].name", "missing")
            if not x.get("address"):
                self.warn(f"founders[{i}].address", "missing: needed for the document")

    def streams(self):
        st = self.d.get("revenue_streams") or []
        if not st:
            self.err("revenue_streams", "no revenue stream (module 4)")
        self.ids(st, "revenue_streams")
        goods = False
        for i, s in enumerate(st):
            p = f"revenue_streams[{i}]"
            if not s.get("name"):
                self.err(f"{p}.name", "missing")
            if not s.get("unit"):
                self.err(f"{p}.unit", "missing")
            price = self.number(s, "price_excl_vat", p, lo=0)
            if price == 0:
                self.err(f"{p}.price_excl_vat", "price is 0")
            if s.get("unit") == "€" and price not in (None, 1):
                self.err(f"{p}.price_excl_vat", "a stream with unit '€' must have price 1 (volumes are amounts)")
            self.vat(s, "vat_rate", p)
            self.vat(s, "purchase_vat_rate", p)
            self.vat(s, "other_var_vat_rate", p)
            self.number(s, "payment_lag_months", p, required=False, lo=0, hi=12, integer=True)
            self.number(s, "supplier_lag_months", p, required=False, lo=0, hi=12, integer=True)
            idx = self.number(s, "price_indexation", p, required=False, lo=-0.5, hi=0.5)
            ps = s.get("price_schedule")
            if ps is not None:
                if not isinstance(ps, dict) or not ps:
                    self.err(f"{p}.price_schedule", "must be {\"YYYY-MM\": price, ...}")
                else:
                    if s.get("unit") == "€":
                        self.err(f"{p}.price_schedule", "a stream with unit '€' has no price: give it a real unit (e.g. month, agreement) and a price")
                    if idx:
                        self.err(f"{p}.price_indexation", "not applied with a price_schedule: put the indexed prices in the schedule")
                    for m, v in ps.items():
                        ym = self.month(m, f"{p}.price_schedule['{m}']")
                        if not isinstance(v, (int, float)) or isinstance(v, bool) or v <= 0:
                            self.err(f"{p}.price_schedule['{m}']", f"'{v}' must be a price > 0")
                        elif ym:
                            self.in_plan(ym, f"{p}.price_schedule['{m}']", "price step in")
            for k in ("goods_cost_pct", "other_var_cost_pct"):
                v = self.number(s, k, p, required=False, lo=0, hi=1)
                if v and v > 0.9:
                    self.warn(f"{p}.{k}", f"{v:.0%} of revenue: check")
            self.number(s, "goods_cost_per_unit", p, required=False, lo=0)
            self.number(s, "stock_cover_months", p, required=False, lo=0, hi=12)
            goods = goods or bool(s.get("goods_cost_pct") or s.get("goods_cost_per_unit"))
            vols = s.get("volumes")
            if not isinstance(vols, dict) or not vols:
                self.err(f"{p}.volumes", "missing: volumes per month are needed")
                continue
            for m, v in vols.items():
                ym = self.month(m, f"{p}.volumes['{m}']")
                if ym is None:
                    continue
                if not isinstance(v, (int, float)) or isinstance(v, bool) or v < 0:
                    self.err(f"{p}.volumes['{m}']", f"'{v}' must be a number ≥ 0")
                elif v:
                    self.in_plan(ym, f"{p}.volumes['{m}']", "volume in")
            if self.start and self.end and not any(v for m, v in vols.items() if MONTH.match(m) and self.start <= (int(m[:4]), int(m[5:])) <= self.end):
                self.err(f"{p}.volumes", "no volume within the plan's horizon")
        if self.types & GOODS_TYPES and st and not goods:
            self.warn("revenue_streams", "business type " + ", ".join(sorted(self.types & GOODS_TYPES)) + " but no cost of goods in any stream (module 5)")

    def costs(self):
        cl = self.d.get("cost_lines") or []
        self.ids(cl, "cost_lines")
        for i, c in enumerate(cl):
            p = f"cost_lines[{i}]"
            if not c.get("name"):
                self.err(f"{p}.name", "missing")
            if c.get("category") not in COST_CATEGORIES:
                self.err(f"{p}.category", f"'{c.get('category')}' must be one of 61, 618, 640, 65")
            self.vat(c, "vat_rate", p)
            self.number(c, "indexation", p, required=False, lo=-0.5, hi=0.5)
            if c.get("schedule"):
                for m, v in c["schedule"].items():
                    ym = self.month(m, f"{p}.schedule['{m}']")
                    if ym and v:
                        self.in_plan(ym, f"{p}.schedule['{m}']")
                continue
            self.number(c, "amount", p, lo=0)
            if c.get("interval_months", 1) not in INTERVALS:
                self.err(f"{p}.interval_months", "must be 0 (once), 1, 3, 6 or 12")
            s = self.month(c.get("start"), f"{p}.start")
            self.in_plan(s, f"{p}.start", "start")
            e = self.month(c.get("end"), f"{p}.end", required=False)
            if s and e and e < s:
                self.err(f"{p}.end", "is before the start")
        if not any(c.get("category") == 618 for c in cl):
            self.warn("cost_lines", "no director remuneration (618): confirm with the owner (question 2.6)")

    def staff(self):
        st = self.d.get("staff") or []
        self.ids(st, "staff")
        for i, s in enumerate(st):
            p = f"staff[{i}]"
            if not s.get("role"):
                self.err(f"{p}.role", "missing")
            sm = self.month(s.get("start"), f"{p}.start")
            self.in_plan(sm, f"{p}.start", "start")
            em = self.month(s.get("end"), f"{p}.end", required=False)
            if sm and em and em < sm:
                self.err(f"{p}.end", "is before the start")
            self.number(s, "fte", p, required=False, lo=0.01, hi=1)
            if not s.get("gross_monthly") and not s.get("employer_cost_monthly"):
                self.err(p, "needs gross_monthly or employer_cost_monthly")
            if "fte" not in s and re.search(r"stud|flexi|deeltijd|temps partiel|part", f"{s.get('contract', '')} {s.get('role', '')}", re.I):
                self.warn(f"{p}.fte", "not set for a part-time, student or flexi job: it counts as 1 FTE")
        names = " ".join(f"{c.get('id', '')} {c.get('name', '')}" for c in self.d.get("cost_lines") or [])
        for label, pat in (("workplace-accident insurance", r"accident|ongeval"), ("social secretariat", r"secr[ée]tariat|payroll"),
                           ("occupational health service", r"prevent|pr[ée]vention|arbeidsgenees|m[ée]decine du travail|occupational")):
            if st and not re.search(pat, names, re.I):
                self.warn("cost_lines", f"the company employs staff but has no {label} in the cost lines (module 6)")

    def investments(self):
        inv = self.d.get("investments") or []
        self.ids(inv, "investments")
        for i, x in enumerate(inv):
            p = f"investments[{i}]"
            if not x.get("name"):
                self.err(f"{p}.name", "missing")
            self.number(x, "amount", p, lo=0.01)
            ym = self.month(x.get("date"), f"{p}.date")
            self.in_plan(ym, f"{p}.date")
            cat = x.get("category", 22)
            if cat not in INVEST_CATEGORIES:
                self.err(f"{p}.category", "must be 21, 22 or 28")
            life = x.get("life_years", 3)
            if cat != 28 and (not isinstance(life, (int, float)) or life <= 0):
                self.err(f"{p}.life_years", "must be > 0 (only deposits, category 28, are not depreciated)")
            self.vat(x, "vat_rate", p)
            if x.get("in_kind"):
                if x.get("vat_rate", 0.21):
                    self.err(f"{p}.vat_rate", "a contribution in kind has no VAT: set 0")
                self.warn(p, "contribution in kind: art. 5:7 WVV requires a founders' report and a bedrijfsrevisor's report; "
                          "check the auditor's fee is in the cost lines (question 2.4)")

    def financing(self):
        fin = self.d.get("financing")
        if not isinstance(fin, dict):
            self.err("financing", "missing (module 9)")
            return
        eq = fin.get("equity") or []
        for i, e in enumerate(eq):
            p = f"financing.equity[{i}]"
            self.number(e, "amount", p, lo=0.01)
            self.in_plan(self.month(e.get("date"), f"{p}.date"), f"{p}.date")
        inkind = [x for x in self.d.get("investments") or [] if x.get("in_kind")]
        if not eq and not inkind:
            self.warn("financing.equity", "no contribution by the founders: a BV has no minimum capital, but the plan must justify the starting equity (art. 5:4)")
        for i, x in enumerate(fin.get("loans") or []):
            p = f"financing.loans[{i}]"
            if not x.get("lender"):
                self.err(f"{p}.lender", "missing")
            if x.get("type", "other") not in LOAN_TYPES:
                self.err(f"{p}.type", "must be bank, founder or other")
            self.number(x, "amount", p, lo=0.01)
            self.number(x, "rate", p, lo=0, hi=0.25)
            term = self.number(x, "term_months", p, lo=1, integer=True)
            grace = self.number(x, "grace_months", p, required=False, lo=0, integer=True) or 0
            if term and grace >= term:
                self.err(f"{p}.grace_months", "must be shorter than term_months (the term includes the grace period)")
            self.in_plan(self.month(x.get("date"), f"{p}.date"), f"{p}.date")
            if x.get("type") == "bank" and not x.get("guarantees"):
                self.warn(f"{p}.guarantees", "no guarantees recorded: the plan must mention them (art. 5:4, §2, 2°); write 'none' if there are none")
        for i, g in enumerate(fin.get("grants") or []):
            p = f"financing.grants[{i}]"
            self.number(g, "amount", p, lo=0.01)
            self.in_plan(self.month(g.get("date"), f"{p}.date"), f"{p}.date")
        self.number(fin, "unavailable_equity", "financing", required=False, lo=0)
        for i, x in enumerate(fin.get("dividends") or []):
            p = f"financing.dividends[{i}]"
            self.number(x, "amount", p, lo=0.01)
            fy = self.number(x, "fy", p, lo=1, hi=getattr(self, "nfy", 5), integer=True)
            ym = self.month(x.get("date"), f"{p}.date")
            self.in_plan(ym, f"{p}.date", "payment in")
            if ym and fy and self.end:
                fy_end = (self.end[0] - (self.nfy - int(fy)), self.end[1])
                prev_end = (fy_end[0] - 1, fy_end[1])
                if ym <= prev_end:
                    self.err(f"{p}.date", f"paid before financial year {fy} starts")
                elif ym <= fy_end:
                    self.warn(p, f"paid during financial year {fy} itself: an interim dividend; check that the articles of association allow it")
                elif (ym[0] - fy_end[0]) * 12 + ym[1] - fy_end[1] > 7:
                    self.warn(f"{p}.date", "paid more than 6 months after the end of the financial year it comes from: the general meeting approves the accounts within 6 months; check the date")
        # The document states "no loans / no grants / no contribution in kind": only when the owner confirmed it.
        conf = set(fin.get("confirmed_none") or [])
        for k in conf - CONFIRMABLE:
            self.err("financing.confirmed_none", f"unknown item '{k}' (use loans, grants, in_kind)")
        for k, what, empty, q in (("loans", "loans", not fin.get("loans"), "2.7 and 9.2–9.3"),
                                  ("grants", "grants", not fin.get("grants"), "9.4"),
                                  ("in_kind", "contributions in kind", not inkind, "2.4")):
            if empty and k not in conf:
                self.err("financing.confirmed_none", f"no {what} in the plan, but the owner did not confirm there are none (question {q})")
            if not empty and k in conf:
                self.err("financing.confirmed_none", f"'{k}' is confirmed as none, but the plan has {what}")

    def tax(self):
        t = self.d.get("tax") or {}
        for k in ("reduced_rate", "prepayments"):
            if k in t and not isinstance(t[k], bool):
                self.err(f"tax.{k}", "must be true or false")
        if "reduced_rate" not in t:
            self.warn("tax.reduced_rate", "not set (question 10.1): the model assumes false")
        start_period = RULES["rules"]["reduced_rate_pay_condition_from_period"]["value"]
        if t.get("reduced_rate") is False and getattr(self, "nfy", 3) < start_period:
            self.warn("tax.reduced_rate", f"false: for a starting company the director-pay condition only applies from taxable period {start_period} "
                      "(art. 215 WIB 92); confirm the owner wants the standard rate (question 10.1)")

    def open_items(self):
        for i, item in enumerate(self.d.get("open_items") or []):
            self.warn(f"open_items[{i}]", str(item))

    def run(self):
        self.meta()
        self.company()
        self.founders()
        self.streams()
        self.costs()
        self.staff()
        self.investments()
        self.financing()
        self.tax()
        self.open_items()
        status = "errors" if self.errors else "warnings" if self.warnings else "ok"
        return {"status": status, "errors": self.errors, "warnings": self.warnings}


def validate(data):
    return Validator(data).run()


def main():
    if any(a in ("-h", "--help") for a in sys.argv[1:]):
        print(__doc__)
        sys.exit(0)
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    try:
        data = json.loads(Path(sys.argv[1]).read_text())
    except json.JSONDecodeError as e:
        print(json.dumps({"status": "errors", "errors": [{"path": "", "message": f"not valid JSON: {e}"}], "warnings": []}))
        sys.exit(1)
    out = validate(data)
    print(json.dumps(out, indent=2, ensure_ascii=False))
    sys.exit(1 if out["errors"] else 0)


if __name__ == "__main__":
    main()
