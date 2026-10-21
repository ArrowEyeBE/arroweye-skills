#!/usr/bin/env python3
"""Propose the default cost lines (interview module 7) that apply to a plan and are not in it yet.

Usage: cost_defaults.py <plan-data.json>

Reads templates/cost-defaults.json and prints JSON:
  applies    which conditions hold (premises, staff, vehicle, food, in_kind, type:…)
  present    default ids already covered by a line of the plan (same id or a matching name)
  proposals  [{id, kind, name, amount, range, ask, basis, entry}] in the plan language;
             `entry` is ready to add to cost_lines (or investments) once the owner agrees.
Nothing is added to the plan: the owner confirms, changes or declines every proposal.
"""
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
DEFAULTS = ROOT / "templates/cost-defaults.json"
RULES = ROOT / "skills/fp-model/references/belgian-rules.json"
GOODS_TYPES = {"trade", "production", "hospitality"}
VEHICLE = re.compile(r"(wagen|\bauto|voiture|camionnette|v[ée]hicule|vehicle|\bvan\b|\bcar\b|truck|scooter|bakfiets|v[ée]lo cargo|cargo bike)", re.I)

# Names that mean a default is already covered, in any language.
MATCH = {
    "notary": r"notari|notaire|notary",
    "kbo": r"\bkbo\b|\bbce\b|ondernemingsloket|guichet|crossroads",
    "accountant": r"boekhoud|accountant|comptab|expert-?comptable",
    "accounting_software": r"factur|invoic|boekhoudsoftware|logiciel compta",
    "bank": r"\bbank|banc",
    "liability_insurance": r"aansprakelijkheid|\bba\b|responsabilit|liability",
    "telecom": r"telefoon|gsm|internet|t[ée]l[ée]phone|phone|telecom|mobile",
    "it_software": r"\bit\b|software|hosting|cloud|logiciel|h[ée]bergement",
    "marketing": r"marketing|publicit|advertis|reclame",
    "nbb_filing": r"neerlegging|jaarrekening|nbb|bnb|d[ée]p[ôo]t des comptes|annual accounts",
    "company_contribution": r"vennootschapsbijdrage|cotisation des soci|company contribution",
    "rent": r"\bhuur(?!waarborg)|loyer|\brent\b",
    "rent_deposit": r"huurwaarborg|garantie locative|deposit",
    "energy": r"energie|[ée]nergie|energy|elektric|[ée]lectricit|gas\b|\bgaz",
    "water": r"\bwater\b|\beau\b",
    "fire_insurance": r"brand|incendie|fire",
    "cleaning": r"schoonma|poets|nettoyage|cleaning",
    "waste": r"afval|d[ée]chets|waste",
    "maintenance": r"onderhoud(?! (voertuig|machine))|herstel|entretien(?! (du v|des m))|maintenance(?! (vehicle|machine))|repair",
    "accident_insurance": r"arbeidsongeval|accidents? du travail|workplace accident",
    "social_secretariat": r"sociaal secretariaat|secr[ée]tariat social|payroll|loonverwerking",
    "prevention_service": r"preventie|arbeidsgeneesk|pr[ée]vention|m[ée]decine du travail|occupational health",
    "vehicle_fuel": r"brandstof|carburant|fuel|laden|recharge",
    "vehicle_insurance": r"verzekering (voertuig|wagen|auto)|assurance (du )?v[ée]hicule|vehicle insurance|autoverzekering",
    "vehicle_tax": r"verkeersbelasting|taxe de circulation|road tax",
    "vehicle_maintenance": r"onderhoud (voertuig|wagen|auto)|entretien du v[ée]hicule|vehicle maintenance",
    "payment_terminal": r"terminal|betaalterm",
    "pos": r"kassa|caisse|\bpos\b|till",
    "webshop_platform": r"webshop|e-?commerce|shopify|woocommerce",
    "favv": r"favv|afsca|food safety",
    "product_liability": r"productaansprakelijk|responsabilit[ée] produit|product liability",
    "machine_maintenance": r"onderhoud machine|entretien des machines|machine maintenance",
    "music_licence": r"unisono|muziek|musique|music",
    "linen": r"linnen|wasserij|linge|blanchisserie|linen|laundry",
    "small_equipment": r"klein materiaal|servies|petit mat[ée]riel|vaisselle|tableware",
    "professional_indemnity": r"beroepsaansprakelijk|responsabilit[ée] professionnelle|professional indemnity",
    "auditor_in_kind": r"revisor|r[ée]viseur|auditor",
    "dir_social": r"sociale bijdrage|cotisations? sociales?|social contribution",
    "dir_allowance": r"onkostenvergoeding|dagvergoeding|indemnit[ée]|allowance",
}


def add_months(ym, k):
    y, m = ym
    t = y * 12 + m - 1 + k
    return t // 12, t % 12 + 1


def ym(s):
    return int(str(s)[:4]), int(str(s)[5:7])


def lookup(path):
    x = json.loads(RULES.read_text())
    for k in path.split("."):
        x = x[k]
    return x


def conditions(d):
    c = d["company"]
    types = set(c.get("business_types", []))
    prem = c.get("premises")
    has_prem = prem in ("rented", "owned") or (prem is None and bool(types & GOODS_TYPES))
    vehicles = [i for i in d.get("investments", []) if VEHICLE.search(i.get("name", ""))]
    out = {"always": True, "premises": has_prem, "rented": has_prem and prem != "owned",
           "staff": bool(d.get("staff")), "vehicle": bool(vehicles) or bool(c.get("vehicle")),
           "food": bool(c.get("food")) or "hospitality" in types,
           "in_kind": any(i.get("in_kind") for i in d.get("investments", [])),
           "management_only": types == {"management"},
           "director_paid": any(x.get("category") == 618 and x.get("id") != "dir_social" for x in d.get("cost_lines", []))}
    for t in ("trade", "production", "hospitality", "services", "software", "management", "other"):
        out[f"type:{t}"] = t in types
    return out, vehicles


def start_month(code, d, vehicles):
    c = d["company"]
    if code == "first_staff" and d.get("staff"):
        return min(ym(s["start"]) for s in d["staff"])
    if code == "first_vehicle" and vehicles:
        return min(ym(v["date"]) for v in vehicles)
    if code == "fy1_filing":
        return add_months(ym(c["first_fy_end"]), 7)
    if code == "first_quarter_end":
        y, m = ym(c["start_date"])
        return y, m + (-m) % 3
    return ym(c["start_date"])


def plan_end(d):
    return add_months(ym(d["company"]["first_fy_end"]), 12 * (int(d["company"].get("fiscal_years", 3)) - 1))


def director_pay_yearly(d):
    """Yearly pay per 618 line over the first 12 months of the plan (the basis for the contribution estimate)."""
    start = ym(d["company"]["start_date"])
    months = [add_months(start, k) for k in range(12)]
    out = []
    for x in d.get("cost_lines", []):
        if x.get("category") != 618 or x.get("id") == "dir_social":
            continue
        if x.get("schedule"):
            vals = [x["schedule"].get(f"{y}-{m:02d}", 0) for y, m in months]
        else:
            s, n = ym(x["start"]), x.get("interval_months", 1)
            k0 = (s[0] - start[0]) * 12 + s[1] - start[1]
            vals = [x["amount"] if (k >= k0 and (k - k0) % n == 0 if n else k == k0) else 0 for k in range(12)]
        paid = [v for v in vals if v]
        out.append(sum(paid) / len(paid) * 12 if paid else 0)  # annualised from the months with pay
    return out


def self_employed(d):
    """Quarterly contribution per paid director: max(minimum, rate x yearly pay / 4), plus the fund's administration cost."""
    r = lookup("legal_items.self_employed_contributions")
    side = all(f.get("side_activity") for f in d.get("founders", [])) and bool(d.get("founders"))
    total = 0
    for pay in director_pay_yearly(d):
        if side:
            q = 0 if pay < r["side_activity_threshold"] else max(r["side_activity_min_quarterly"], r["rate"] * pay / 4)
        else:
            q = max(r["min_quarterly"], r["rate"] * min(pay, r["ceiling_first_bracket"]) / 4)
        total += q * (1 + r["admin_cost_typical"])
    return round(total, 2), side


def present(line, d):
    pat = re.compile(MATCH.get(line["id"], re.escape(line["id"])), re.I)
    items = d.get("investments", []) if line.get("kind") == "investment" else d.get("cost_lines", [])
    return any(x.get("id") == line["id"] or pat.search(x.get("name", "")) for x in items)


def propose(d):
    spec = json.loads(DEFAULTS.read_text())
    lang = d["meta"]["language"]
    cond, vehicles = conditions(d)
    types = set(d["company"].get("business_types", []))
    proposals, have = [], []
    for line in spec["lines"]:
        if not any(cond.get(w) for w in [line["when"]] + line.get("also", [])):
            continue
        if any(cond.get(w) for w in line.get("not_when", [])):
            continue
        if present(line, d):
            have.append(line["id"])
            continue
        amount = lookup(line["amount_from"]) if line.get("amount_from") else line.get("amount")
        basis = line["basis"]
        if line.get("compute") == "self_employed":
            amount, side = self_employed(d)
            basis += " Computed as a " + ("side activity" if side else "main occupation") + f" from {len(director_pay_yearly(d))} director pay line(s)."
        exempt_until = None
        if line.get("compute") == "company_contribution" and d["company"].get("contribution_exempt"):
            ex = lookup("legal_items.company_contribution_exemption")
            first_year = ym(d["company"]["start_date"])[0] + ex["years"]
            if (first_year, 1) > plan_end(d):
                have.append(f"{line['id']} (exempt until {first_year - 1}, after the plan)")
                continue
            exempt_until = first_year - 1
        for t, v in (line.get("by_type") or {}).items():
            if t in types:
                amount = max(amount or 0, v)
        n = len(d.get("staff", [])) if line.get("per") == "employee" else len(vehicles) or 1 if line.get("per") == "vehicle" else 1
        if amount is not None:
            amount = round(amount * n, 2)
        y, m = start_month(line.get("start", "plan"), d, vehicles)
        if exempt_until:
            y, m = exempt_until + 1, 1
        source = f"default (cost-defaults {spec['checked']})"
        if line.get("kind") == "investment":
            entry = {"id": line["id"], "name": line["name"][lang], "category": line["category"], "amount": amount,
                     "date": f"{y}-{m:02d}", "life_years": line.get("life_years", 0), "vat_rate": line["vat_rate"]}
        else:
            entry = {"id": line["id"], "name": line["name"][lang], "category": line["category"], "amount": amount,
                     "interval_months": line["interval_months"], "start": f"{y}-{m:02d}", "vat_rate": line["vat_rate"], "source": source}
        proposals.append({"id": line["id"], "kind": line.get("kind", "cost_line"), "name": line["name"][lang], "amount": amount,
                          "range": [round(r * n, 2) for r in line["range"]] if line.get("range") else None,
                          "per": line.get("per"), "count": n if line.get("per") else None,
                          "ask": bool(line.get("ask")) or amount is None
                          or (line.get("compute") == "company_contribution" and "contribution_exempt" not in d["company"]),
                          "basis": basis, "entry": entry})
    return {"checked": spec["checked"], "applies": [k for k, v in cond.items() if v], "present": have, "proposals": proposals}


def main():
    if any(a in ("-h", "--help") for a in sys.argv[1:]):
        print(__doc__)
        sys.exit(0)
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    print(json.dumps(propose(json.loads(Path(sys.argv[1]).read_text())), indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
