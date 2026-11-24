"""
Load the SIF portfolio JSON ({year: {category: [funds]}}) into Neo4j.

Same graph shape as historical-mf/ingestion-pipeline/neo4j_parser.py, on the same instance, so every
name differs from the MF graph and the two never share a node, relationship, property or constraint:

  (:SIF_FUND {sifCode, sifFundName, sifCategory})
    -[:SIF_LATEST_MONTHLY_SNAPSHOT]->
  (:SIF_MONTHLY_SNAPSHOT {sifCode, sifYear, sifMonth, sifReportedDate, sifAUM})
    -[:SIF_PREVIOUS_MONTHLY_SNAPSHOT]-> (older snapshot)
    -[:SIF_HAS_ASSET_CLASS]->
  (:SIF_ASSET_CLASS {sifName, sifSnapshotYear, sifSnapshotMonth})
    -[:SIF_HAS_INSTRUMENT {sifHoldings, sifSubType, sifIncludedInNetAssets}]->
  (:SIF_INSTRUMENT {sifISIN, sifName, sifAssetClass, sifSubType, sifSector, sifIndustry, sifSymbol, sifMarketCap})

One difference from MF: a derivative's subType (derivative-hedging / derivative-directional-position /
derivatives) belongs to the fund-month, not the contract, so one contract can be hedging in one fund and
directional in another. sifSubType on SIF_HAS_INSTRUMENT is always the holding's value; the SIF_INSTRUMENT
node only carries sifSubType for non-derivatives, where it describes the instrument itself.

sifIncludedInNetAssets separates holdings that add up to the fund's net assets (true) from exposure the
AMC discloses outside them, mostly futures (false). Filter on it before summing sifHoldings.

From the factsheets (factsheets/factsheet_parser.py -> factsheets/output/factsheet_fund_managers.csv), as MF loads
its own:

  sifAUM on SIF_MONTHLY_SNAPSHOT     month-end AUM in Rs crore, for the month the factsheet's AUM is as on
  (:SIF_FUND_MANAGER {sifName})-[:SIF_MANAGES {sifStartDate, sifEndDate}]->(:SIF_FUND)

Tenures come from factsheets/manager_tenures.py: dates read 'Nov-2025' and are only what the factsheets print;
sifEndDate is 'present' while the fund's latest factsheet lists the manager. A fund's SIF_MANAGES relationships are
rebuilt on every load.
Managers are loaded only for the SIFs in --input (--all-managers: every SIF in the factsheet CSV).
"""
import argparse
import csv
import json
import os
import re
import sys
from collections import defaultdict
from datetime import datetime

from dotenv import load_dotenv
from neo4j import GraphDatabase, RoutingControl

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# this project's .env, else the MF project's (same instance); variables already set are left alone
load_dotenv(os.path.join(PROJECT_ROOT, ".env"))
load_dotenv(os.path.join(os.path.dirname(PROJECT_ROOT), "historical-mf", ".env"))

URI = os.getenv("NEO4J_URI")
USERNAME = os.getenv("NEO4J_USERNAME")
PASSWORD = os.getenv("NEO4J_PASSWORD")
DATABASE = os.getenv("NEO4J_DATABASE", "neo4j")

# one JSON per category (holdings.json in the parent folder repeats the same funds, so it is not read)
EXCEL_INPUT = os.path.join(PROJECT_ROOT, "output", "categories")
SNAPSHOT_INPUT = os.path.join(PROJECT_ROOT, "output", "scheme_snapshot.csv")
FUND_MANAGER_INPUT = os.path.join(PROJECT_ROOT, "factsheets", "output", "factsheet_fund_managers.csv")

sys.path.insert(0, os.path.join(PROJECT_ROOT, "factsheets"))

from manager_tenures import load_tenures  # noqa: E402

DERIVATIVE_SUB_TYPES = {"derivative-hedging", "derivative-directional-position", "derivatives"}
ASSET_CLASSES = {"Equity", "Debt", "REITs & InvITs", "Cash and Cash Equivalents", "MF and ETF", "AIF", "Commodities"}

arg_parser = argparse.ArgumentParser(description="Load SIF portfolio JSON into Neo4j.")
arg_parser.add_argument(
    "--input", nargs="+", default=[EXCEL_INPUT],
    help="JSON file(s) or folder(s) of JSON files (default: output/categories)",
)
arg_parser.add_argument(
    "--category", nargs="+",
    help='Only load these categories, e.g. --category "Hybrid Long-Short"',
)
arg_parser.add_argument(
    "--months", type=int,
    help="Only load the latest N months present in the input, e.g. --months 4 (default: all months)",
)
arg_parser.add_argument(
    "--fund-managers", default=FUND_MANAGER_INPUT,
    help="Factsheet manager / AUM CSV (default: factsheets/output/factsheet_fund_managers.csv)",
)
arg_parser.add_argument(
    "--managers-only", action="store_true",
    help="Only load SIF_FUND_MANAGER nodes and SIF_MANAGES relationships (for the SIFs in --input), not the portfolio snapshots",
)
arg_parser.add_argument(
    "--all-managers", action="store_true",
    help="Load managers for every SIF in the factsheet CSV, not only the SIFs in --input "
         "(creates SIF_FUND nodes with no portfolio for the others)",
)
arg_parser.add_argument(
    "--schema-only", action="store_true",
    help="Create the SIF constraints and exit, without uploading any data",
)
arg_parser.add_argument(
    "--dry-run", action="store_true",
    help="Print what would be loaded and exit without connecting to Neo4j",
)
args = arg_parser.parse_args()

if args.months is not None and args.months < 1:
    arg_parser.error("--months must be a positive number")


def collect_json_files(input_paths):
    found = []

    for input_path in input_paths:
        if os.path.isdir(input_path):
            found.extend(
                os.path.join(input_path, file_name)
                for file_name in os.listdir(input_path)
                if file_name.lower().endswith(".json")
            )
        elif os.path.isfile(input_path):
            found.append(input_path)
        else:
            print("Input not found:", input_path)

    if not found:
        print("No JSON files found in:", ", ".join(input_paths))

    return found


def load_date_mismatches(csv_path):
    """(scheme name, YYYY-MM) of statements filed under the wrong month (date_check in scheme_snapshot.csv)."""
    if not os.path.isfile(csv_path):
        print("WARNING: scheme_snapshot.csv not found, not checking statement dates:", csv_path)
        return set()

    with open(csv_path, encoding="utf-8", newline="") as file:
        return {
            (row["scheme_name"], row["portfolio_date"][:7])
            for row in csv.DictReader(file)
            if row.get("date_check", "OK") != "OK"
        }


def load_factsheet_aum(csv_path):
    """{(sif code, YYYY-MM): month-end AUM in Rs crore} from the factsheet CSV, keyed on the month the AUM is as on."""
    aum = {}

    with open(csv_path, encoding="utf-8", newline="") as file:
        for row in csv.DictReader(file):
            if row["aum_month_end_cr"] and row["portfolio_month"]:
                aum.setdefault((row["scheme_code"], row["portfolio_month"]), float(row["aum_month_end_cr"]))

    return aum


wanted_categories = (
    {name.strip().lower() for name in args.category}
    if args.category else None
)

fund_managers, factsheet_aum = [], {}

if os.path.isfile(args.fund_managers):
    fund_managers = [
        {"sifCode": code, **fund}
        for code, fund in load_tenures(args.fund_managers, wanted_categories).items()
    ]
    factsheet_aum = load_factsheet_aum(args.fund_managers)
    print(f"Fund managers: {sum(len(f['tenures']) for f in fund_managers)} tenures across {len(fund_managers)} funds; "
          f"factsheet AUM for {len(factsheet_aum)} fund-months")
else:
    print("Factsheet CSV not found, loading no SIF_MANAGES and no sifAUM:", args.fund_managers)

# --managers-only still reads --input, but only for the funds whose managers to load (no snapshots)
json_files = sorted(collect_json_files(args.input))

if not json_files and not (args.managers_only and args.all_managers and fund_managers):
    print("Nothing to upload.")
    sys.exit()

date_mismatches = load_date_mismatches(SNAPSHOT_INPUT)

monthly_data = defaultdict(list)
# fund name, category and date per fund-month, apart from the holdings: a factsheet-only month in merged.json (AUM and
# managers, no portfolio: WSIF JUN-SEP 2026, Infinity) has no holding to read them from
snapshot_meta = {}
input_codes = set()  # every SIF in --input (after --category), for the manager filter below

for json_file in json_files:
    print("Loading:", os.path.basename(json_file))

    with open(json_file, "r", encoding="utf-8") as file:
        file_data = json.load(file)

    for year, categories in file_data.items():
        for category_name, funds in categories.items():
            if wanted_categories and category_name.strip().lower() not in wanted_categories:
                continue

            for fund in funds:
                fund_name = fund.get("fund_name")
                sif_code = fund.get("scheme_amfi_code")

                if not sif_code:
                    print(f"WARNING: Missing SIF code for fund: {fund_name}")
                    continue

                input_codes.add(str(sif_code))

                for month, month_records in fund.items():
                    if month in ["fund_name", "scheme_amfi_code", "fund_manager"]:
                        continue

                    # output/merged/merged.json (merge_portfolio_factsheet.py) wraps each month as {"AUM", "portfolio"};
                    # holdings.json and the category files have the bare list
                    month_aum = None
                    if isinstance(month_records, dict):
                        month_aum = month_records.get("AUM")
                        month_records = month_records.get("portfolio")

                    if not isinstance(month_records, list):
                        continue

                    try:
                        month_number = datetime.strptime(month, "%B").month
                    except ValueError:
                        print(f"WARNING: Unknown month '{month}' for fund '{fund_name}'")
                        continue

                    date_object = datetime(int(year), month_number, 1)
                    month_key = date_object.strftime("%Y-%m")

                    if month_aum is not None:
                        factsheet_aum.setdefault((str(sif_code), month_key), month_aum)

                    # a statement for another month would be a false snapshot (Apex MAR 2026 holds the May file)
                    if (fund_name, month_key) in date_mismatches:
                        print(f"WARNING: skipping {fund_name} {month_key}: statement is for a different month")
                        continue

                    snapshot_meta.setdefault((str(sif_code), month_key), (fund_name, category_name, date_object))
                    monthly_data[(str(sif_code), month_key)].extend(
                        {
                            "fund_name": fund_name,
                            "category": category_name,
                            "date": date_object,
                            "item": item,
                        }
                        for item in month_records
                    )

# managers only for the SIFs being loaded: the factsheet CSV covers every SIF processed so far, and loading
# them all would create SIF_FUND nodes (with managers, no portfolio) for SIFs left out of --input
if not args.all_managers:
    skipped = sorted({f["sifCode"] for f in fund_managers} - input_codes)
    fund_managers = [f for f in fund_managers if f["sifCode"] in input_codes]
    if skipped:
        print(f"Managers not loaded for {len(skipped)} SIF(s) not in --input (use --all-managers to load them):",
              ", ".join(skipped))
    if not fund_managers and args.managers_only:
        print("No managers to upload for the SIFs in --input.")
        sys.exit()

if args.managers_only:
    monthly_data.clear()

if args.months:
    # keep only the latest N "YYYY-MM" months found across all funds
    latest_months = sorted({month_key for _, month_key in monthly_data})[-args.months:]
    print(f"Uploading latest {len(latest_months)} month(s):", ", ".join(latest_months))

    monthly_data = {key: entries for key, entries in monthly_data.items() if key[1] in latest_months}

# per fund oldest month first: each new snapshot takes over SIF_LATEST_MONTHLY_SNAPSHOT
sorted_months = sorted(monthly_data)


def build_instruments(entries):
    instruments = {}

    for entry in entries:
        item = entry["item"]
        asset_class = str(item.get("asset_class") or "").strip()

        if asset_class not in ASSET_CLASSES:
            print("WARNING: unrecognized asset_class, skipping record:", asset_class, item.get("instrument"))
            continue

        key = item.get("isin") or item.get("instrument")

        if not key:
            print("  WARNING: dropping instrument with no ISIN/name:", item)
            continue

        holdings = item.get("holdings") or 0

        if key in instruments:
            instruments[key]["holdings"] += holdings
            continue

        sub_type = item.get("subType")
        instruments[key] = {
            "key": key,
            "name": item.get("instrument"),
            "assetClass": asset_class,
            "subType": sub_type,
            "instrumentSubType": None if sub_type in DERIVATIVE_SUB_TYPES else sub_type,
            "sector": item.get("sector"),
            "industry": item.get("industry"),
            "symbol": item.get("symbol"),
            "marketCap": item.get("market_cap"),
            "holdings": holdings,
            "includedInNetAssets": item.get("included_in_net_assets"),
        }

    return list(instruments.values())


print()
print(f"Fund/month snapshots found: {len(sorted_months)}")

if args.dry_run:
    totals = defaultdict(int)

    for sif_code, month_key in sorted_months:
        entries = monthly_data[(sif_code, month_key)]
        instruments = build_instruments(entries)
        for instrument in instruments:
            totals[instrument["subType"] if instrument["subType"] in DERIVATIVE_SUB_TYPES else "(other)"] += 1
        aum = factsheet_aum.get((sif_code, month_key))
        print(" ", sif_code, "->", month_key, "->", snapshot_meta[(sif_code, month_key)][0], "->", len(instruments), "instruments",
              "| AUM", f"{aum} Cr" if aum is not None else "not in the factsheets")

    print("Holdings by subType:", dict(totals))
    print()
    for fund in fund_managers:
        print(fund["sifCode"], fund["fundName"])
        for tenure in fund["tenures"]:
            print("   ", tenure["name"], "|", tenure["startDate"], "->", tenure["endDate"])
    sys.exit()

if not (URI and USERNAME and PASSWORD):
    sys.exit("NEO4J_URI / NEO4J_USERNAME / NEO4J_PASSWORD are not set (.env here or in historical-mf)")

driver = GraphDatabase.driver(URI, auth=(USERNAME, PASSWORD))
driver.verify_connectivity()

print("Connected to Neo4j:", re.sub(r"^.*://", "", URI), "| database:", DATABASE)

constraints = [
    "CREATE CONSTRAINT sif_fund_code_unique IF NOT EXISTS FOR (f:SIF_FUND) REQUIRE f.sifCode IS UNIQUE",
    "CREATE CONSTRAINT sif_snapshot_fund_year_month_unique IF NOT EXISTS FOR (s:SIF_MONTHLY_SNAPSHOT) REQUIRE (s.sifCode, s.sifYear, s.sifMonth) IS UNIQUE",
    "CREATE CONSTRAINT sif_instrument_isin_unique IF NOT EXISTS FOR (i:SIF_INSTRUMENT) REQUIRE i.sifISIN IS UNIQUE",
    "CREATE CONSTRAINT sif_fund_manager_name_unique IF NOT EXISTS FOR (m:SIF_FUND_MANAGER) REQUIRE m.sifName IS UNIQUE",
]

for statement in constraints:
    driver.execute_query(statement, database_=DATABASE)

print("Constraints and indexes verified.")

if args.schema_only:
    driver.close()
    print("Schema only: no data uploaded. See ingestion-pipeline/sif_schema_preview.cypher to draw the full schema.")
    sys.exit()

query = """

MERGE (fund:SIF_FUND {
    sifCode: $sifCode
})
SET fund.sifFundName = $fundName,
    fund.sifCategory = $category

MERGE (snapshot:SIF_MONTHLY_SNAPSHOT {
    sifCode: $sifCode,
    sifYear: $year,
    sifMonth: $month
})
SET snapshot.sifReportedDate = date($reportedDate),
    snapshot.sifAUM = $AUM

WITH fund, snapshot

OPTIONAL MATCH (fund)-[latestRel:SIF_LATEST_MONTHLY_SNAPSHOT]->(currentLatest)

FOREACH (_ IN CASE
    WHEN currentLatest IS NOT NULL
         AND currentLatest <> snapshot
         AND snapshot.sifReportedDate > currentLatest.sifReportedDate
    THEN [1]
    ELSE []
END |
    MERGE (snapshot)-[:SIF_PREVIOUS_MONTHLY_SNAPSHOT]->(currentLatest)
    DELETE latestRel
    MERGE (fund)-[:SIF_LATEST_MONTHLY_SNAPSHOT]->(snapshot)
)

FOREACH (_ IN CASE
    WHEN currentLatest IS NULL
    THEN [1]
    ELSE []
END |
    MERGE (fund)-[:SIF_LATEST_MONTHLY_SNAPSHOT]->(snapshot)
)

WITH snapshot

// a re-load replaces the snapshot's holdings, so a corrected file leaves no stale rows behind
// (SIF_ASSET_CLASS nodes belong to one snapshot; SIF_INSTRUMENT nodes are shared and stay)
OPTIONAL MATCH (snapshot)-[:SIF_HAS_ASSET_CLASS]->(stale:SIF_ASSET_CLASS)
DETACH DELETE stale

WITH DISTINCT snapshot, $instruments AS instruments

UNWIND instruments AS item

// one SIF_ASSET_CLASS node per snapshot and asset class name
MERGE (snapshot)-[:SIF_HAS_ASSET_CLASS]->(assetClass:SIF_ASSET_CLASS {
    sifName: item.assetClass
})
SET assetClass.sifSnapshotYear = snapshot.sifYear,
    assetClass.sifSnapshotMonth = snapshot.sifMonth

// every asset class shares one SIF_INSTRUMENT node keyed by ISIN (synthetic for derivatives, CASHANDEQ for cash)
MERGE (instrument:SIF_INSTRUMENT {
    sifISIN: item.key
})
SET instrument.sifName = item.name,
    instrument.sifAssetClass = item.assetClass,
    instrument.sifSubType = item.instrumentSubType,
    instrument.sifSector = item.sector,
    instrument.sifIndustry = item.industry,
    instrument.sifSymbol = item.symbol,
    instrument.sifMarketCap = item.marketCap

MERGE (assetClass)-[h:SIF_HAS_INSTRUMENT]->(instrument)
SET h.sifHoldings = item.holdings,
    h.sifSubType = item.subType,
    h.sifIncludedInNetAssets = item.includedInNetAssets

RETURN count(*) AS processed

"""

for sif_code, month_key in sorted_months:
    entries = monthly_data[(sif_code, month_key)]
    fund_name, category_name, report_date = snapshot_meta[(sif_code, month_key)]
    instruments = build_instruments(entries)

    driver.execute_query(
        query,
        sifCode=sif_code,
        fundName=fund_name,
        category=category_name,
        year=report_date.year,
        month=report_date.strftime("%B"),
        reportedDate=report_date.strftime("%Y-%m-%d"),
        AUM=factsheet_aum.get((sif_code, month_key)),
        instruments=instruments,
        database_=DATABASE,
        routing_=RoutingControl.WRITE,
    )

    print("Snapshot processed:", sif_code, month_key, fund_name, "| instruments:", len(instruments),
          "" if instruments else "(factsheet only: AUM, no portfolio)")

# a fund's SIF_MANAGES relationships are rebuilt on each load: a tenure's sifEndDate moves as new factsheets arrive
manager_query = """

UNWIND $funds AS row

MERGE (fund:SIF_FUND {
    sifCode: row.sifCode
})
ON CREATE SET fund.sifFundName = row.fundName,
    fund.sifCategory = row.category

WITH fund, row

OPTIONAL MATCH (:SIF_FUND_MANAGER)-[old:SIF_MANAGES]->(fund)
DELETE old

WITH DISTINCT fund, row

UNWIND row.tenures AS tenure

MERGE (manager:SIF_FUND_MANAGER {
    sifName: tenure.name
})

CREATE (manager)-[:SIF_MANAGES {
    sifStartDate: tenure.startDate,
    sifEndDate: tenure.endDate
}]->(fund)

"""

if fund_managers:
    driver.execute_query(
        manager_query,
        funds=fund_managers,
        database_=DATABASE,
        routing_=RoutingControl.WRITE,
    )
    print("Fund managers processed:", sum(len(f["tenures"]) for f in fund_managers), "tenures")

driver.close()

print()
print("-" * 60)
print("GRAPH CREATED SUCCESSFULLY!")
print("-" * 60)
