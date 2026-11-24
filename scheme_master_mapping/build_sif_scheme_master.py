"""
Build sif_scheme_master.xlsx from AMFI's SIF scheme list.

AMFI gives SIFs their own code series (SIF-1, SIF-2, ...), kept apart from the mutual fund codes,
so they are not in scheme_master.xlsx or AMFI's NAVAll.txt. The list is read from a public mirror
of AMFI's NAV data; re-run this when new SIF strategies launch:

  python scheme_master_mapping/build_sif_scheme_master.py [--input amfi_nav_master_latest.csv]

Sheets:
  scheme_master   one row per strategy, same columns as scheme_master.xlsx; the code is the
                  Regular Growth plan (the MF master's convention), else Direct Growth
  all_plans       every plan / option with its code and ISIN
  former_names    names a strategy's disclosures used to carry, with its code and current name
                  (AMFI lists only the current name; kept here in FORMER_NAMES so a rebuild keeps them)
"""
from pathlib import Path
import argparse
import re
import sys

import pandas as pd

SOURCE_URL = "https://api.tigzig.com/mf/v1/download/amfi_nav_master_latest.csv"
OUTPUT_FILE = Path(__file__).resolve().with_name("sif_scheme_master.xlsx")

arg_parser = argparse.ArgumentParser(description="Build sif_scheme_master.xlsx from AMFI's SIF scheme list.")
arg_parser.add_argument("--input", default=SOURCE_URL, help="AMFI NAV master CSV: local path or URL")
args = arg_parser.parse_args()

# sif_parser parses sys.argv on import, so it must not see this script's options
sys.argv = sys.argv[:1]
sys.path.insert(0, str(OUTPUT_FILE.parent.parent))
from sif_parser import canonical_scheme_name  # noqa: E402

# spelling in AMFI's list -> name used in the disclosures
NAME_FIXES = {"qsif actice asset allocator": "qsif Active Asset Allocator"}
# AMFI's fund house -> brand: "iSIF SIF" -> "iSIF", "SUMMIT SIF" -> "Summit SIF"
FUND_HOUSE_FIXES = {"isif sif": "iSIF", "qsif sif": "qsif", "dynasif sif": "DynaSIF", "summit sif": "Summit SIF"}
# former name in a disclosure -> (current name, why we know it is the same strategy)
FORMER_NAMES = {
    "Altiva Equity Hybrid Long-Short Fund": (
        "Altiva Hybrid Long-Short Fund",
        "Title of the Oct 2025 statement; same sheet code AEHYLS and description as Nov 2025 onward",
    ),
}


def strategy_name(name):
    for wrong, right in NAME_FIXES.items():
        name = re.sub(wrong, right, name, flags=re.I)
    return canonical_scheme_name(name)


source = pd.read_csv(args.input, dtype=str, low_memory=False)
sif = source[source["source"] == "SIF"].copy()

sif["fund_house"] = sif["amc"].map(lambda a: FUND_HOUSE_FIXES.get(a.strip().lower(), a.strip()))
sif["strategy"] = sif["scheme_name"].map(strategy_name)
unnamed = sif[sif["strategy"].isna()]
if not unnamed.empty:
    sys.exit(f"Could not derive a strategy name for: {unnamed['scheme_name'].tolist()}")

sif["sort_code"] = sif["scheme_code"].astype(int)
sif["scheme_code"] = "SIF-" + sif["scheme_code"]
sif["category"] = sif["category_sub"].str.replace(r"\s+Fund$", "", regex=True).str.upper()
sif = sif.sort_values("sort_code")

rows = []
for strategy, plans in sif.groupby("strategy", sort=False):
    growth = plans[plans["scheme_option"] == "Growth"]
    regular = growth[growth["scheme_plan"] == "Regular"]
    chosen = (regular if not regular.empty else growth if not growth.empty else plans).iloc[0]
    if len(regular) > 1:
        print(f"NOTE: {strategy}: {len(regular)} Regular Growth codes {regular['scheme_code'].tolist()}, "
              f"using {chosen['scheme_code']}")
    rows.append(dict(scheme_code=chosen["scheme_code"], fund_house=chosen["fund_house"],
                     category=chosen["category"], scheme_name=strategy))

all_plans = sif.rename(columns={"scheme_name": "amfi_scheme_name", "strategy": "scheme_name",
                                "scheme_plan": "plan", "scheme_option": "option"})[
    ["scheme_code", "isin", "isin2", "scheme_name", "amfi_scheme_name", "fund_house",
     "category", "plan", "option", "first_date"]]

code_of = {row["scheme_name"]: row["scheme_code"] for row in rows}
unknown = [current for current, _ in FORMER_NAMES.values() if current not in code_of]
if unknown:
    sys.exit(f"FORMER_NAMES points at names not in AMFI's list: {unknown}")
former_names = pd.DataFrame(
    [dict(former_name=former, scheme_code=code_of[current], scheme_name=current, note=note)
     for former, (current, note) in FORMER_NAMES.items()],
    columns=["former_name", "scheme_code", "scheme_name", "note"],
)

with pd.ExcelWriter(OUTPUT_FILE) as writer:
    pd.DataFrame(rows).to_excel(writer, sheet_name="scheme_master", index=False)
    all_plans.to_excel(writer, sheet_name="all_plans", index=False)
    former_names.to_excel(writer, sheet_name="former_names", index=False)

print(f"Wrote {OUTPUT_FILE}: {len(rows)} strategies, {len(all_plans)} plans")
