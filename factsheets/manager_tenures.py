"""Fund manager tenures from output/factsheet_fund_managers.csv, shared by the merged JSON and the Neo4j loader."""
import csv
import os
from collections import defaultdict
from datetime import datetime

# factsheet categories left out of the tenures (MF left out "Large and Mid Cap"); none for SIFs
EXCLUDED_CATEGORIES = set()

EXITS_FILE = "factsheet_manager_exits.csv"


def month_year(value):
    """'2024-11-04' or '2024-11' -> 'Nov-2024'."""
    return datetime.strptime(value[:7], "%Y-%m").strftime("%b-%Y")


def load_exits(csv_path):
    """{(scheme_code, manager): [printed exit dates, oldest first]} from the exits file beside csv_path, if there is one."""
    exits = defaultdict(list)
    path = os.path.join(os.path.dirname(os.path.abspath(csv_path)), EXITS_FILE)

    if os.path.isfile(path):
        with open(path, "r", encoding="utf-8", newline="") as file:
            for row in csv.DictReader(file):
                exits[(row["scheme_code"], row["fund_manager"])].append(row["exit_date"])

    return {key: sorted(dates) for key, dates in exits.items()}


def load_tenures(csv_path, wanted_categories=None):
    """
    {scheme_code: {"fundName", "category", "tenures": [{"name", "startDate", "endDate"}]}}.

    A tenure is an unbroken run of the fund's factsheets listing the manager, so someone who leaves and comes back
    has two. startDate is the date printed during the tenure ('Nov-2024', the earliest if the print changed). When no
    factsheet in the tenure prints one, or the printed date is after the first factsheet listing them, it is the
    first month of the tenure: the first factsheet listing the manager (its portfolio month, the month the data is
    as on), so a manager added in a later factsheet starts that month.
    endDate of a manager's latest tenure is the exit date printed on the fund's pages (manager_exits.py) for a
    manager the fund's latest factsheet no longer lists, and 'present' otherwise - also when no exit date is printed.
    An earlier tenure (they came back after it) ends at the exit date printed for it, else the last month a
    factsheet listed them in it.
    """
    funds = {}
    exits = load_exits(csv_path)

    with open(csv_path, "r", encoding="utf-8", newline="") as file:
        for row in csv.DictReader(file):
            if row["category"] in EXCLUDED_CATEGORIES:
                continue

            if wanted_categories and row["category"].strip().lower() not in wanted_categories:
                continue

            fund = funds.setdefault(
                row["scheme_code"],
                {"name": row["scheme_name"], "category": row["category"], "months": set(), "managers": defaultdict(dict)},
            )
            fund["months"].add(row["portfolio_month"])

            if row["fund_manager"]:
                fund["managers"][row["fund_manager"]][row["portfolio_month"]] = row["managing_since"]

    result = {}

    for code, fund in funds.items():
        months = sorted(fund["months"])
        tenures = []

        for name, listed in fund["managers"].items():
            runs, run = [], []

            for index, month in enumerate(months):
                if month in listed:
                    run.append(month)

                if run and (month not in listed or index == len(months) - 1):
                    runs.append(run)
                    run = []

            for number, run in enumerate(runs):
                since = min((listed[m] for m in run if listed[m]), default=None) or run[0]
                # a printed date after the first factsheet listing them is not when they started (DynaSIF JUL
                # 2026 prints 'August 7, 2026' for Mayur Patel, listed since JUN): the first listed month is
                if since[:7] > run[0]:
                    since = run[0]
                left = next((d for d in exits.get((code, name), []) if d[:7] >= run[-1]), None)
                if number < len(runs) - 1:
                    # a stint they came back after: the printed exit date, else the last month listed in it
                    end = month_year(left if left and left[:7] < runs[number + 1][0] else run[-1])
                else:
                    # the latest stint: 'present' unless an exit date is printed for a manager no longer listed
                    end = month_year(left) if left and run[-1] != months[-1] else "present"
                tenures.append({"name": name, "startDate": month_year(since), "endDate": end})

        result[code] = {"fundName": fund["name"], "category": fund["category"], "tenures": tenures}

    return result
