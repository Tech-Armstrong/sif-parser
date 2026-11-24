"""
Add SEBI industry classification, as published on nseindia.com, to the AMFI
market-cap lists in this folder.

Input : nse_industry.csv  (symbol, macro_economic_sector, sector, industry,
        basic_industry, series, isin, fetched_on, source) - pulled from NSE's
        own get-quote API (www.nseindia.com/api/NextApi/apiClient/GetQuoteApi,
        functionName=getSymbolData -> equityResponse[0].secInfo).
Output: every <YEAR>/market_cap_*.xlsx gets four columns after the existing ones
        (header row 2): NSE Macro Sector | NSE Sector | NSE Industry | NSE Basic Industry

Matching: NSE Symbol first; rows whose symbol NSE no longer recognises (renamed,
e.g. ZOMATO -> ETERNAL) are filled by ISIN from any list where that ISIN matched.
BSE-only companies stay blank (NSE has no page for them).

    python excel\\market_cap_mapping\\add_nse_industry.py
"""
from pathlib import Path
import csv
import glob
import openpyxl

HERE = Path(__file__).resolve().parent
SRC = HERE / "nse_industry.csv"
NEW_COLS = [("NSE Macro Sector", "macro_economic_sector"), ("NSE Sector", "sector"),
            ("NSE Industry", "industry"), ("NSE Basic Industry", "basic_industry")]
HEADER_ROW = 2


def s(v):
    return str(v).strip() if v is not None else ""


def main():
    by_sym = {}
    with open(SRC, encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if r["industry"]:
                by_sym[r["symbol"].upper()] = r
    files = sorted(glob.glob(str(HERE / "20*" / "market_cap_*.xlsx")))
    files = [f for f in files if not Path(f).name.startswith("~$")]

    books, by_isin = {}, {}
    for f in files:                                   # pass 1: symbol match, learn ISIN -> industry
        wb = openpyxl.load_workbook(f)
        ws = wb.worksheets[0]
        hdr = {s(c.value): c.column for c in ws[HEADER_ROW] if c.value is not None}
        for name, _ in NEW_COLS:                      # create or reuse the 4 columns
            if name not in hdr:
                col = ws.max_column + 1
                ws.cell(HEADER_ROW, col, name)._style = ws.cell(HEADER_ROW, hdr["ISIN"])._style
                hdr[name] = col
        books[f] = (wb, ws, hdr)
        for row in range(HEADER_ROW + 1, ws.max_row + 1):
            rec = by_sym.get(s(ws.cell(row, hdr["NSE Symbol"]).value).upper())
            if rec:
                for col in ("ISIN", "New ISIN"):
                    if col in hdr and s(ws.cell(row, hdr[col]).value):
                        by_isin.setdefault(s(ws.cell(row, hdr[col]).value).upper(), rec)

    for f, (wb, ws, hdr) in books.items():           # pass 2: write
        n = filled = by_isin_n = 0
        for row in range(HEADER_ROW + 1, ws.max_row + 1):
            isin = s(ws.cell(row, hdr["ISIN"]).value).upper()
            if not isin:
                continue
            n += 1
            rec = by_sym.get(s(ws.cell(row, hdr["NSE Symbol"]).value).upper())
            if rec is None:
                rec = by_isin.get(isin)
                if rec is None and "New ISIN" in hdr:
                    rec = by_isin.get(s(ws.cell(row, hdr["New ISIN"]).value).upper())
                by_isin_n += rec is not None
            for name, key in NEW_COLS:
                ws.cell(row, hdr[name], rec[key] if rec else None)
            filled += rec is not None
        wb.save(f)
        print(f"{Path(f).name}: {filled}/{n} rows with NSE industry ({by_isin_n} via ISIN)")


if __name__ == "__main__":
    main()
