"""Score the code rules (rules/) against the model's earlier answers, without any API call.

Usage:
    python factsheets/rules_eval.py [--amc Magnum] [--month MAR] [--show 20]

For every fund already processed (output/<YEAR>/<MON>/<AMC>/fund_details.csv with an llm/<code>.json answer from
the model) it re-reads the fund's pages from the PDF text layer, runs the rules and compares field by field.
Per field it prints coverage (how often a rule returned a value) and agreement (how often that value matches the
model), then the disagreements so a new or changed rule can be checked before it is relied on. "Done by rules"
is the share of funds the agent would finish without calling the model.
"""
import argparse
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd

import fund_extractor as fx
import rules
from factsheet_parser import DATA, amc_tag, read_pdf

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output"


def norm_name(s):
    return " ".join(re.sub(r"[^a-z ]", " ", (s or "").lower()).split())


def same(field, rule_value, model_value):
    if field == "fund_managers":
        a = {(norm_name(m["name"]), m["managing_since"]) for m in rule_value}
        b = {(norm_name(m["name"]), m["managing_since"]) for m in model_value}
        return a == b
    if isinstance(rule_value, float) and model_value is not None:
        return abs(rule_value - model_value) <= max(0.01, 0.0005 * abs(model_value))
    return rule_value == model_value


def show(field, v):
    if field == "fund_managers":
        return "; ".join(f"{m['name']} {m['managing_since']}" for m in v or [])
    return str(v)


def cases(amc=None, month=None):
    for f in sorted(OUT.glob("*/*/*/fund_details.csv")):
        year, mon, house = f.parts[-4:-1]
        if (amc and amc.lower() not in house.lower()) or (month and month.upper() != mon) or f.stat().st_size <= 2:
            continue
        df = pd.read_csv(f, dtype={"scheme_code": str})
        texts = {}
        for code, g in df.groupby("scheme_code"):
            pdf = DATA / year / mon / g.factsheet_file.iloc[0]
            cache = f.parent / "llm" / f"{code}.json"
            pages = [int(p) for p in str(g.pdf_pages.iloc[0]).split(",") if p.strip().isdigit()]
            if not pdf.exists() or not cache.exists() or not pages:
                continue
            cached = json.loads(cache.read_text(encoding="utf-8"))
            if cached.get("source") == "rules":
                continue  # answered by the rules themselves: no model answer to score them against
            if pdf not in texts:
                texts[pdf] = read_pdf(pdf)[0]  # the text layer as the parser read it
            text = "\n\n".join(f"--- PDF page {p} ---\n{texts[pdf][p - 1]}" for p in pages)
            # the model's own answer, before rule values were merged in; the rules get the scheme master's name, as
            # the parser gives them
            yield f"{mon} {house} | {g.scheme_name.iloc[0]}", amc_tag(house), text, cached.get("model_result") or cached["result"]


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--amc", help="Only SIF folders whose name contains this text")
    ap.add_argument("--month", help="Only this month folder, e.g. MAR")
    ap.add_argument("--show", type=int, default=15, help="Disagreements to print per field (default 15)")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")

    n, done = 0, 0
    covered, agree, by_rule, wrong = Counter(), Counter(), defaultdict(Counter), defaultdict(list)
    for label, house, text, model in cases(args.amc, args.month):
        n += 1
        found = rules.run(rules.Page(fund_name=label.split(" | ", 1)[1], text=text, amc=house))
        for f in rules.FIELDS:
            if not found.has(f):
                continue
            covered[f] += 1
            ok = model.get(f) is None if f in found.not_printed else same(f, found.values[f], model.get(f))
            agree[f] += ok
            by_rule[found.source[f]]["ok" if ok else "wrong"] += 1
            if not ok:
                wrong[f].append(f"{label}\n      rules ({found.source[f]}): {show(f, found.values.get(f, 'not printed'))}\n      model: {show(f, model.get(f))}")
        done += fx.rules_complete(found)

    print(f"{n} funds with a model answer; done by rules alone: {done} ({done / max(n, 1):.0%})\n")
    print(f"{'field':<26}{'coverage':>10}{'agreement':>12}")
    for f in rules.FIELDS:
        print(f"{f:<26}{covered[f] / max(n, 1):>10.0%}{agree[f] / max(covered[f], 1):>12.0%}   ({agree[f]}/{covered[f]})")
    print("\nper rule:")
    for name, c in sorted(by_rule.items()):
        print(f"  {name:<36} ok {c['ok']:>4}   wrong {c['wrong']:>3}")
    for f in rules.FIELDS:
        if wrong[f]:
            print(f"\n--- {f}: {len(wrong[f])} disagreements" + (f" (first {args.show})" if len(wrong[f]) > args.show else ""))
            for w in wrong[f][:args.show]:
                print("  " + w)


if __name__ == "__main__":
    main()
