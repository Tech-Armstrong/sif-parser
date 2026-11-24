"""Merge the monthly SIF portfolio holdings and the factsheet managers and AUM into one JSON.

Usage:
    python merge_portfolio_factsheet.py [--year 2026] [--month SEP] [--amc Magnum --amc iSIF] [--holdings PATH]
                                        [--factsheets PATH] [--output PATH]

Ported from historical-mf (excel/merge_portfolio_factsheet.py). Inputs: output/holdings.json from sif_parser.py and
factsheets/output/factsheet_fund_managers.csv from factsheets/factsheet_parser.py.

--year and --month limit the output and write it beside the full file, so merged.json is kept: merged_2026.json,
merged_2026_september.json (or merged_september.json, that month of every year). fund_manager stays the fund's whole
list of tenures, as the tenures are worked out from every factsheet. --amc keeps only the SIFs whose fund house
contains that text (repeatable) and writes merged_selected.json unless --output is given.

    {
        "2026": {
            "Hybrid Long-Short": [
                {
                    "fund_name": "Magnum Hybrid Long-Short Fund",
                    "scheme_amfi_code": "SIF-13",
                    "fund_manager": [{"name": "Gaurav Mehta", "tenure": [["Oct-2025", "present"]]}],
                    "september": {
                        "AUM": 3980.49,
                        "portfolio": [{"instrument": "...", "isin": "...", "holdings": 2.8, ...}]
                    }
                }
            ]
        }
    }

Funds are joined on the SIF code (SIF-13), months on the month the data is as on: a factsheet in the SEP folder with
data as on 31 August (Platinum's) gives August's AUM. AUM is the month-end AUM in Rs crore, null where no factsheet
covers the month. A factsheet month with no disclosure file keeps its AUM with an empty portfolio, so nothing from
either source is dropped. fund_manager is the fund's whole list of managers under every year, each with a tenure list
of [in, out] dates (manager_tenures.py), so a manager who leaves and comes back has one entry with several stints, as
in historical-mf. Writes coverage.csv next to the output: what each fund-month has.
"""
import argparse
import calendar
import csv
import json
import os
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "factsheets"))

from manager_tenures import EXCLUDED_CATEGORIES, load_tenures  # noqa: E402

MONTHS = [m.lower() for m in calendar.month_name if m]
FIXED = ("fund_name", "scheme_amfi_code")


def load_aum(csv_path):
    """{(scheme_code, 'YYYY-MM'): (month-end AUM, fund name, category)}, one per fund-month."""
    aum = {}
    with open(csv_path, "r", encoding="utf-8", newline="") as file:
        for row in csv.DictReader(file):
            if row["category"] in EXCLUDED_CATEGORIES or not row["portfolio_month"]:
                continue
            end = float(row["aum_month_end_cr"]) if row["aum_month_end_cr"] else None
            aum[(row["scheme_code"], row["portfolio_month"])] = (end, row["scheme_name"], row["category"])
    return aum


def month_name(value):
    """9 / 09 / sep / SEP / September -> 'september'."""
    v = value.strip().lower()
    if v.isdigit() and 1 <= int(v) <= 12:
        return MONTHS[int(v) - 1]
    for name in MONTHS:
        if v in (name, name[:3]):
            return name
    raise argparse.ArgumentTypeError(f"invalid month: {value!r} (use 1-12, SEP or September)")


def select(holdings, aum, year=None, month=None):
    """holdings and AUM limited to one year and/or one month."""
    kept = {}
    for y, categories in holdings.items():
        if year and y != str(year):
            continue
        for category, funds in categories.items():
            for fund in funds:
                months = [m for m in MONTHS if m in fund and (not month or m == month)]
                if months:
                    entry = {**{k: fund[k] for k in FIXED}, **{m: fund[m] for m in months}}
                    kept.setdefault(y, {}).setdefault(category, []).append(entry)
    return kept, {
        (code, ym): value for (code, ym), value in aum.items()
        if (not year or ym[:4] == str(year)) and (not month or MONTHS[int(ym[5:]) - 1] == month)
    }


def select_amcs(holdings, aum, amcs):
    """holdings and AUM limited to the SIFs whose fund house in the scheme master contains one of amcs."""
    master = pd.read_excel(ROOT / "scheme_master_mapping" / "sif_scheme_master.xlsx", sheet_name="scheme_master", dtype=str)
    codes = {r.scheme_code for r in master.itertuples() if any(a.lower() in r.fund_house.lower() for a in amcs)}
    kept = {}
    for year, categories in holdings.items():
        for category, funds in categories.items():
            chosen = [f for f in funds if f.get("scheme_amfi_code") in codes]
            if chosen:
                kept.setdefault(year, {})[category] = chosen
    return kept, {key: value for key, value in aum.items() if key[0] in codes}


def managers(tenures):
    """
    [{"name", "startDate", "endDate"}, ...] -> [{"name": "...", "tenure": [(start, end), ...]}, ...]: one entry per
    manager, in the order they first appear, with every stint (a manager who leaves and comes back has several).
    As in historical-mf.
    """
    grouped = {}
    for t in tenures:
        grouped.setdefault(t["name"], []).append((t["startDate"], t["endDate"]))
    return [{"name": name, "tenure": stints} for name, stints in grouped.items()]


def merge(holdings, aum, tenures):
    """The merged JSON and one coverage row per fund-month."""
    merged, coverage, placed = {}, [], set()
    tenures = {code: managers(fund["tenures"]) for code, fund in tenures.items()}

    for year, categories in holdings.items():
        merged[year] = {}
        for category, funds in categories.items():
            merged[year][category] = []
            for fund in funds:
                code = str(fund.get("scheme_amfi_code") or "")
                if not code:
                    print(f"WARNING: {year} {fund['fund_name']} has no SIF code in the holdings: no AUM or managers joined")
                entry = {"fund_name": fund["fund_name"], "scheme_amfi_code": fund.get("scheme_amfi_code"),
                         "fund_manager": tenures.get(code, [])}
                for month in MONTHS:
                    if month not in fund:
                        continue
                    key = (code, f"{year}-{MONTHS.index(month) + 1:02d}")
                    placed.add(key)
                    entry[month] = {"AUM": aum.get(key, (None,))[0], "portfolio": fund[month]}
                    coverage.append((year, month, category, code, fund["fund_name"], len(fund[month]), key in aum))
                merged[year][category].append(entry)

    # factsheet months with no disclosure file
    for (code, ym), (end, name, category) in sorted(aum.items()):
        if (code, ym) in placed:
            continue
        year, month = ym[:4], MONTHS[int(ym[5:]) - 1]
        year_funds = merged.setdefault(year, {})
        entry = next((f for fs in year_funds.values() for f in fs if str(f.get("scheme_amfi_code")) == code), None)
        if entry is None:
            entry = {"fund_name": name, "scheme_amfi_code": code, "fund_manager": tenures.get(code, [])}
            year_funds.setdefault(category, []).append(entry)
        entry[month] = {"AUM": end, "portfolio": []}
        coverage.append((year, month, category, code, entry["fund_name"], 0, True))

    # calendar order within a fund, whichever source added the month (holdings.json lists them alphabetically)
    for categories in merged.values():
        for funds in categories.values():
            for i, fund in enumerate(funds):
                head = {k: fund[k] for k in ("fund_name", "scheme_amfi_code", "fund_manager")}
                funds[i] = {**head, **{m: fund[m] for m in MONTHS if m in fund}}

    return merged, coverage


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--holdings", default=str(ROOT / "output" / "holdings.json"))
    ap.add_argument("--factsheets", default=str(ROOT / "factsheets" / "output" / "factsheet_fund_managers.csv"))
    ap.add_argument("--year", type=int, help="Only this year, e.g. 2026")
    ap.add_argument("--month", type=month_name, help="Only this month, e.g. 9, SEP or September")
    ap.add_argument("--amc", action="append", help="Only SIFs whose fund house contains this text (repeatable)")
    ap.add_argument("--output", help="Output JSON (default: output/merged/merged[_<year>][_<month>][_selected].json)")
    args = ap.parse_args()

    suffix = "".join(f"_{v}" for v in (args.year, args.month, args.amc and "selected") if v)
    output = args.output or str(ROOT / "output" / "merged" / f"merged{suffix}.json")

    with open(args.holdings, "r", encoding="utf-8") as file:
        holdings = json.load(file)

    holdings, aum = select(holdings, load_aum(args.factsheets), args.year, args.month)
    if args.amc:
        holdings, aum = select_amcs(holdings, aum, args.amc)

    if not holdings and not aum:
        raise SystemExit(f"No holdings or factsheet data for{suffix.replace('_', ' ')}")

    merged, coverage = merge(holdings, aum, load_tenures(args.factsheets))

    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    with open(output + ".tmp", "w", encoding="utf-8") as file:
        json.dump(merged, file, indent=4, ensure_ascii=False, allow_nan=False)
    os.replace(output + ".tmp", output)

    coverage_file = os.path.join(os.path.dirname(os.path.abspath(output)), f"coverage{suffix}.csv")
    with open(coverage_file, "w", encoding="utf-8", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["year", "month", "category", "scheme_amfi_code", "fund_name", "holdings_rows", "has_factsheet"])
        writer.writerows(coverage)

    with_both = sum(1 for c in coverage if c[5] and c[6])
    print(f"{len(coverage)} fund-months -> {output}")
    print(f"  holdings and factsheet: {with_both}, holdings only: {sum(1 for c in coverage if c[5] and not c[6])}, "
          f"factsheet only: {sum(1 for c in coverage if not c[5] and c[6])}")


if __name__ == "__main__":
    main()
