#!/usr/bin/env python3
"""Summarise a recalculated model: checks + key figures per financial year (input for checkpoint 2).

Usage: check_model.py <recalculated.xlsx>
Exit code 1 if any check has status ERROR.
"""
import json
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import plugin_env  # noqa: E402

plugin_env.ensure()

from i18n import LABELS  # noqa: E402
from model_io import ModelValues  # noqa: E402

CHECKS = ("ck_balance", "ck_cash", "ck_buffer", "ck_reduced", "ck_equity", "ck_dividend", "ck_horizon")
FIGURES = {"revenue": "a70", "gross_margin": "a9900", "operating_result": "a9901", "profit_before_tax": "a9903",
           "corporate_tax": "a67", "net_result": "a9904", "cash_year_end": "bs_cash", "equity_year_end": "bs_eq",
           "lowest_cash_in_year": "kpi_min"}


def r2(v):
    return round(v, 2) if isinstance(v, (int, float)) else v


def legal_rules(meta, today=None):
    """How current the Belgian rules in the model are. Stale = older than the review period or from an earlier calendar year."""
    today = today or date.today()
    checked = date.fromisoformat(meta["rules_checked"])
    age = (today - checked).days
    return {"last_checked": meta["rules_checked"], "age_days": age,
            "stale": age > int(meta["rules_review_days"]) or today.year > checked.year,
            "to_verify": json.loads(meta["rules_verify"])}


def summarise(path):
    mv = ModelValues(path)
    lang = mv.meta["lang"]
    status = {LABELS[k][lang]: k[3:] for k in ("st_ok", "st_error", "st_warn")}
    checks = [{"check": k, "value": r2(mv.cell(k, "B")), "status": status.get(mv.cell(k, "C"), mv.cell(k, "C"))}
              for k in CHECKS]
    cash = mv.monthly("cash")
    months = mv.months()
    i = min(range(len(cash)), key=lambda k: cash[k])
    return {
        "financial_years": [f"{a} ({b})" for a, b in mv.fy_labels],
        "checks": checks,
        "key_figures": {name: [r2(v) for v in mv.annual(row)] for name, row in FIGURES.items()},
        "lowest_cash": {"amount": r2(cash[i]), "month": months[i].strftime("%Y-%m")},
        "balance_sheet_difference": [r2(v) for v in mv.annual("bs_diff", opening=True)],
        "legal_rules": legal_rules(mv.meta),
    }


def main():
    if any(a in ("-h", "--help") for a in sys.argv[1:]):
        print(__doc__)
        sys.exit(0)
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    out = summarise(sys.argv[1])
    print(json.dumps(out, indent=2, ensure_ascii=False))
    sys.exit(1 if any(c["status"] == "error" for c in out["checks"]) else 0)


if __name__ == "__main__":
    main()
