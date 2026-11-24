"""Exit dates of fund managers as printed on the fund's factsheet pages, for manager_tenures.py.

Usage:
    python factsheets/manager_exits.py

Reads output/factsheet_fund_managers.csv for each fund's managers and pages, and writes
output/factsheet_manager_exits.csv. It reads the PDFs' text layer only, so a run costs nothing. Run it after the
factsheet parser. A manager with no printed exit gets no row, and their tenure's endDate is 'present'
(manager_tenures.py).
"""
import csv
import re
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from manager_tenures import EXCLUDED_CATEGORIES  # noqa: E402
from factsheet_parser import DATA, read_pdf  # noqa: E402  (the text layer as the parser read it)
from rules.common import DATE, parse_date  # noqa: E402

INPUT = ROOT / "output" / "factsheet_fund_managers.csv"
OUTPUT = ROOT / "output" / "factsheet_manager_exits.csv"


def flatten(text):
    """One line, without the full stops in 'Mr.' and 'w.e.f.' that would end a sentence for the patterns below."""
    text = re.sub(r"\b(?:Mr|Ms|Mrs|Dr)\.\s*", "", " ".join(text.split()))
    return re.sub(r"\bw\.?\s?e\.?\s?f\b\.?", "wef", text, flags=re.I)


def name_rx(name):
    """The first and last words of the name, up to two words apart: 'Shibani Kurian' / 'Shibani Sircar Kurian'."""
    words = [re.escape(w) for w in name.split()]
    if len(words) == 1:
        return rf"\b{words[0]}\b"
    return rf"\b{words[0]}(?:\W+\w+){{0,2}}?\W+{words[-1]}\b"


def exit_patterns(n):
    """Every wording seen on a fund's pages that dates a manager's exit; the date is group 'd'."""
    return [
        # 'Ihab Dalwai have ceased to be fund manager of the scheme wef June 15, 2026' (ICICI, Invesco, Baroda, Mirae)
        rf"{n}[^.]{{0,70}}?\b(?:ceased?|ceases)\b[^.]{{0,90}}?(?P<d>{DATE})",
        # 'Milind Agrawal was co managing the Equity portion of Fund till 30th Apr 2026' (SBI)
        rf"{n}[^.]{{0,40}}?was\s+co[\s-]*managing[^.]{{0,70}}?till\s+(?P<d>{DATE})",
        # 'Jitendra Sriram was appointed as Fund Manager wef July 1, 2026 in place of Sanjay Chawla' (Baroda)
        rf"(?:wef|effect\s+from)\s+(?P<d>{DATE})[^.]{{0,40}}?\bin\s+place\s+of\b[^.]{{0,50}}?{n}",
        # 'With effect from Jan 15, 2026, Kedar Karnik replaced Pranavi Kulkarni & Dhawal Dalal' (Edelweiss)
        rf"(?:wef|effect\s+from)\s+(?P<d>{DATE}),?[^.]{{0,40}}?\breplaced\b[^.]{{0,50}}?{n}",
        # 'Mehul Dalmia has replaced Raj Koradia and has been appointed ..., effective August 3, 2026' (Edelweiss)
        rf"\breplaced\s+[^.]{{0,40}}?{n}[^.]{{0,80}}?\beffective\s+(?P<d>{DATE})",
        # 'Anish Tawakley is managing this fund wef May 16, 2026. Prior to that the same was managed by Abhishek Singh' (DSP)
        rf"(?:wef|since)\s+(?P<d>{DATE})\.?\s*Prior\s+to\s+that[^.]{{0,40}}?managed\s+by\s+[^.]{{0,30}}?{n}",
        # 'Kunal Jain^^ (Managing fund since 30-May-25 to 01-Apr-26)' (Canara)
        rf"{n}\W{{0,4}}Managing\s+fund\s+since\s+{DATE}\s+to\s+(?P<d>{DATE})",
    ]


def printed_exits(name, text, year):
    """[(YYYY-MM[-DD], as printed)] for this manager on these pages."""
    n = name_rx(name)
    found = []
    for pattern in exit_patterns(n):
        for m in re.finditer(pattern, text, re.I):
            date = parse_date(m.group("d"))
            if date:
                found.append((date, m.group(0)))
    # Nippon's line is cut off in the PDF, with no year: 'Ashutosh Bhargava (Managing Since August 2020 till till March 10'
    for m in re.finditer(rf"{n}\s*\(Managing\s+Since\s+[A-Za-z]+\s+\d{{4}}\s+(?:to\s+)?till\s+(?:till\s+)?(?P<m>[A-Za-z]+)\s+\d", text, re.I):
        date = parse_date(f"{m.group('m')} 1, {year}")
        if date:
            found.append((date[:7], m.group(0)))
    return found


def main():
    with open(INPUT, "r", encoding="utf-8", newline="") as file:
        rows = [r for r in csv.DictReader(file) if r["category"] not in EXCLUDED_CATEGORIES and r["fund_manager"] and r["pdf_pages"]]

    # a note repeats month after month, so each fund's pages are searched for every manager the fund has ever listed
    managers, fund_months = defaultdict(set), {}
    for r in rows:
        managers[r["scheme_code"]].add(r["fund_manager"])
        fund_months[(r["year"], r["factsheet_month"], r["amc"], r["factsheet_file"], r["pdf_pages"], r["scheme_code"])] = r["scheme_name"]

    by_pdf = defaultdict(list)
    for key in fund_months:
        by_pdf[key[:4]].append(key)

    exits = {}
    for (year, month, amc, file), keys in by_pdf.items():
        pdf = DATA / year / month / file
        if not pdf.exists():
            continue
        texts, _ = read_pdf(pdf)
        for key in keys:
            text = flatten(" ".join(texts[int(q) - 1] for q in key[4].split(",")))
            for name in managers[key[5]]:
                for date, printed in printed_exits(name, text, year):
                    exits.setdefault((key[5], name, date), (fund_months[key], month, key[4], printed))

    with open(OUTPUT, "w", encoding="utf-8", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(["scheme_code", "scheme_name", "fund_manager", "exit_date", "factsheet_month", "pdf_pages", "as_printed"])
        for (code, name, date), (scheme, month, pages, printed) in sorted(exits.items()):
            writer.writerow([code, scheme, name, date, month, pages, printed])

    print(f"{len(exits)} printed exits -> {OUTPUT}")


if __name__ == "__main__":
    main()
