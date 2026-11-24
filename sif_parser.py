"""SIF portfolio parser — usage: python sif_parser.py "<root folder containing <YEAR>/<MON>/<SIF>/files>" [out_dir]

Maps SIF (Specialised Investment Fund) monthly portfolio disclosures (Excel only; PDFs are ignored)
-> standardized schema (holdings, scheme_snapshot, amc_reported_totals).
Adapted from the mutual-fund parser (historical-mf/excel/standardized_format/mf_portfolio_parser.py):
SIFs are run by MF houses (Altiva = Edelweiss, iSIF = ICICI, Magnum = SBI, RedHex = HSBC, Summit = Invesco, ...)
and print the same portfolio layouts, so the row parsing / classification is shared."""
from pathlib import Path
import argparse
import hashlib
import json
import pandas as pd, glob, re, warnings, io, sys, os, calendar

warnings.filterwarnings("ignore")
PROJECT_ROOT = Path(__file__).resolve().parent

_arg_parser = argparse.ArgumentParser(
    description="Parse SIF monthly disclosures into the standardized format."
)
_arg_parser.add_argument(
    "root", nargs="?", default=str(PROJECT_ROOT / "data" / "monthly_portfolio"),
    help="Disclosures root laid out as <YEAR>/<MON>/<SIF>/*.xls*",
)
_arg_parser.add_argument(
    "out_dir", nargs="?", default=str(PROJECT_ROOT / "output"),
    help="Output directory",
)
_arg_parser.add_argument(
    "--category",
    help='Only parse funds in this category, e.g. "Hybrid Long-Short" (see get_fund_category)',
)
_arg_parser.add_argument("--year", help="Only parse this year folder, e.g. 2026")
_arg_parser.add_argument("--month", help="Only parse this month folder, e.g. APR")
_arg_parser.add_argument("--amc", action="append", help="Only parse SIF folders whose name contains this text (repeatable)")
_arg_parser.add_argument(
    "--split-json", metavar="HOLDINGS_JSON",
    help="Skip parsing; split an existing holdings.json into categories/<Category>.json",
)
# imported (factsheets/instrument_match.py uses the AMFI lookups below): defaults, not the importer's command line
ARGS = _arg_parser.parse_args() if __name__ == "__main__" else _arg_parser.parse_args([])

# Portfolio workbooks built from a factsheet by factsheets/factsheet_portfolio.py, for SIFs that publish no
# disclosure: "<Scheme> (from factsheet).xlsx". Parsed after the real disclosures, so a real file filed for the
# same scheme-month wins the duplicate check.
FROM_FACTSHEET_SUFFIX = " (from factsheet).xlsx"

ROOT = ARGS.root
OUT_DIR = ARGS.out_dir
MARKET_CAP_ROOT = PROJECT_ROOT / "market_cap_mapping"
sys.path.insert(0, str(PROJECT_ROOT / "industry_sector_mapping"))
from industry_sector import add_sector_columns  # SEBI industry -> sector
# built by scheme_master_mapping/build_sif_scheme_master.py; SIFs are not in the MF scheme_master.xlsx
SIF_SCHEME_MASTER_FILE = PROJECT_ROOT / "scheme_master_mapping" / "sif_scheme_master.xlsx"


# NOTE: portfolio_date is taken from the month folder name (JAN/FEB); extend MONTH_END for other months/years.
# ISO 6166: 2-letter country code + 9 alphanumeric + 1 check digit.
# Any country (IN, US, FR, GB, ...) so overseas holdings keep their ISIN.
_ISIN_SHAPE = re.compile(r'^[A-Z]{2}[A-Z0-9]{9}[0-9]$')


def isin_check_digit_ok(code):
    # ISO 6166 check digit: letters -> 10..35, then Luhn over the digit string
    digits = "".join(str(int(ch, 36)) for ch in code[:-1])
    total = 0
    for k, d in enumerate(reversed(digits)):
        n = int(d)
        if k % 2 == 0:
            n *= 2
            if n > 9:
                n -= 9
        total += n
    return (10 - total % 10) % 10 == int(code[-1])


class _IsinPattern:
    """
    Drop-in for the old compiled regex (ISIN.match(x)).
    IN... codes: shape only (as before). Other countries: shape + valid check
    digit, so AMC pseudo-codes like Motilal's "TREP01012026" are not taken
    for an overseas ISIN.
    """
    def match(self, value):
        value = str(value)
        if not _ISIN_SHAPE.match(value):
            return None
        if value.startswith("IN") or isin_check_digit_ok(value):
            return _ISIN_SHAPE.match(value)
        return None


ISIN = _IsinPattern()


def is_domestic_isin(isin):
    return bool(isin) and isin.startswith("IN")
MONTH_NUM = {
    "JAN": 1,
    "FEB": 2,
    "MAR": 3,
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DEC": 12,
}
def sif_brand(folder):
    """SIF folder -> brand name: "RedHex" / "RedHex SIF" -> "RedHex SIF", "iSIF" -> "iSIF"."""
    folder = norm(folder)
    return folder if re.search(r"sif$", folder, re.I) else f"{folder} SIF"

def norm(s): return " ".join(str(s).split()).strip()
def tonum(v):
    if isinstance(v, bool): return None
    if isinstance(v, (int, float)) and not pd.isna(v): return float(v)
    if isinstance(v, str):
        t = v.strip().replace(",", "")
        m = re.match(r'^\(?(-?\d+\.?\d*)\)?$', t)
        if m: return -abs(float(m.group(1))) if t.startswith("(") else float(m.group(1))
    return None

def clean_symbol(value):
    if pd.isna(value):
        return None

    value = str(value).strip()

    if not value or value.lower() in ("nan", "none", "null"):
        return None

    # Treat placeholder dashes ("-", "--", "–", "—") as blank
    if not value.strip("-–— "):
        return None

    return value


# "Limited", "LIMITED", "Ltd.", "LTD." and the misspellings seen in AMC files
# ("Limted", "Limtied", "LIMITE") -> one suffix
LTD_RX = re.compile(r"\b(?:lim(?:ited|ted|tied|ite)|ltd)\b\.*", re.I)
# names cut off at a column width: "APOLLO HOSPITALS ENTERPRISE LT" (last word only, so "LT Foods" is safe)
LTD_TRUNCATED_RX = re.compile(r"^(\S+\s+\S+\s+.*?)\blt\.?$", re.I)


def standardize_ltd(name):
    """Hindalco Industries Limited -> Hindalco Industries Ltd (LTD when the name is all caps)."""
    if not isinstance(name, str) or not name:
        return name
    ltd = "LTD" if name.upper() == name else "Ltd"
    name = LTD_TRUNCATED_RX.sub(lambda m: m.group(1) + ltd, name)
    return LTD_RX.sub(ltd, name)

SIF_SCHEME_MASTER_CACHE = None

# every SIF strategy title ends in "... Long-Short Fund" (Equity / Equity Ex-Top 100 / Hybrid / Active Asset Allocator ...)
SIF_TITLE_RX = re.compile(r"(.*?\blong\s*-?\s*short\s+(?:fund|sif)\b)", re.I)  # "Sapphire Equity Long-Short SIF"
# brand spellings that title-casing would break
SIF_WORD_CASE = {"isif": "iSIF", "dynasif": "DynaSIF", "qsif": "qsif", "wsif": "WSIF", "msif": "MSIF", "sif": "SIF", "redhex": "RedHex"}


def canonical_scheme_name(title):
    """
    One spelling per SIF strategy across months / AMCs:
      "SIFAHLS APEX HYBRID LONG-SHORT FUND APEX HYBRID ..." -> "Apex Hybrid Long-Short Fund"
      "SCHEME NAME : Magnum Equity Ex-Top 100 Long Short Fund" -> "Magnum Equity Ex-Top 100 Long-Short Fund"
      "Prism Hybrid Long - Short Fund" -> "Prism Hybrid Long-Short Fund"
    """
    m = SIF_TITLE_RX.search(norm(title))
    if not m:
        return None
    name = m.group(1)
    # "PORTFOLIO STATEMENT OF <scheme>" (Altiva), "SCHEME NAME : <scheme>" (Magnum)
    name = re.sub(r"^.*?\b(portfolio (statement )?of|scheme name\s*:)\s*", "", name, flags=re.I)
    # leading scheme short code: "SIFAHLS APEX ...", "EQLSSIF Summit ...", "1 DIVLSF DIVINITI ...";
    # all-caps brands that also end in SIF ("WSIF Equity ...", "MSIF ...", "DYNASIF ...") stay
    name = re.sub(r"^(\d+\s+)?(?![A-Z]SIF\s|DYNASIF\s)([A-Z0-9]*\d[A-Z0-9]*|[A-Z]*SIF[A-Z]+|[A-Z]+SIF)\s+(?=\w)", "", name)
    words = [SIF_WORD_CASE.get(w.lower(), w[:1].upper() + w[1:].lower()) for w in name.split()]
    name = " ".join(words)
    name = re.sub(r"long\s*-?\s*short", "Long-Short", name, flags=re.I)
    name = re.sub(r"ex\s*-?\s*top\s*-?\s*(\d+)", r"Ex-Top \1", name, flags=re.I)
    return name


def load_sif_scheme_master():
    """Canonical SIF scheme name (lower case) -> its scheme_master row in sif_scheme_master.xlsx."""
    global SIF_SCHEME_MASTER_CACHE

    if SIF_SCHEME_MASTER_CACHE is None:
        df = pd.read_excel(SIF_SCHEME_MASTER_FILE, sheet_name="scheme_master", dtype=str)
        SIF_SCHEME_MASTER_CACHE = {norm(r["scheme_name"]).lower(): r for r in df.to_dict("records")}
        print(f"Loaded {len(SIF_SCHEME_MASTER_CACHE)} SIF schemes from {SIF_SCHEME_MASTER_FILE.name}.")

    return SIF_SCHEME_MASTER_CACHE


def get_scheme_master_row(scheme_name):
    if not scheme_name:
        return None

    return load_sif_scheme_master().get(norm(scheme_name).lower())


SIF_FORMER_NAMES_CACHE = None


def current_scheme_name(scheme_name):
    """
    A strategy's current name for a name it used to carry, from the former_names sheet of sif_scheme_master.xlsx
    ("Altiva Equity Hybrid Long-Short Fund", its title in October 2025, -> "Altiva Hybrid Long-Short Fund"), so all its
    months are one fund under one SIF code. Any other name comes back as it is.
    """
    global SIF_FORMER_NAMES_CACHE

    if SIF_FORMER_NAMES_CACHE is None:
        with pd.ExcelFile(SIF_SCHEME_MASTER_FILE) as book:
            df = book.parse("former_names", dtype=str) if "former_names" in book.sheet_names else None
        SIF_FORMER_NAMES_CACHE = {} if df is None else {
            norm(r["former_name"]).lower(): norm(r["scheme_name"]) for r in df.to_dict("records")
        }

    return SIF_FORMER_NAMES_CACHE.get(norm(scheme_name).lower(), scheme_name)


def load_market_cap_mapping(mapping_file):
    """
    Load AMFI six-month market-cap mapping.

    The AMFI Excel has:
        Row 1 -> title
        Row 2 -> column headers
        Row 3 onwards -> data
    """

    if not mapping_file.exists():
        raise FileNotFoundError(
            f"Market cap mapping file not found: {mapping_file}"
        )

    df = pd.read_excel(mapping_file, header=1)

    required_columns = [
        "ISIN",
        "NSE Symbol",
        "BSE Symbol",
        "Categorization as per SEBI Circular dated Oct 6, 2017"
    ]
    
    missing = [
        col for col in required_columns
        if col not in df.columns
    ]

    if missing:
        raise ValueError(
            f"Invalid market cap mapping file: {mapping_file}\n"
            f"Missing columns: {missing}"
        )

    mapping = {}
    new_isin_aliases = {}

    for _, row in df.iterrows():
        isin = norm(row["ISIN"]).upper()

        if not isin:
            continue

        if not ISIN.match(isin):
            continue

        nse_symbol = clean_symbol(row["NSE Symbol"])
        bse_symbol = clean_symbol(row["BSE Symbol"])

        symbol = nse_symbol or bse_symbol or None

        category = norm(
            row["Categorization as per SEBI Circular dated Oct 6, 2017"]
        )

        if category not in ("Large Cap", "Mid Cap", "Small Cap"):
            category = None

        mapping[isin] = {
            "symbol": symbol,
            "market_cap_category": category,
            "isin": isin,
            "company": norm(row["Company name"]) if "Company name" in df.columns else "",
            # SEBI industry as published on nseindia.com (added by market_cap_mapping/add_nse_industry.py)
            "industry": norm(row["NSE Industry"]) or None if "NSE Industry" in df.columns and not pd.isna(row["NSE Industry"]) else None,
        }

        # ISIN changed (split / listing): "New ISIN" points at the same company,
        # so holdings that already carry the new ISIN resolve to this row too
        if "New ISIN" in df.columns:
            new_isin = norm(row["New ISIN"]).upper() if not pd.isna(row["New ISIN"]) else ""
            if new_isin and ISIN.match(new_isin):
                new_isin_aliases.setdefault(new_isin, dict(mapping[isin], isin=new_isin))

    for new_isin, data in new_isin_aliases.items():
        mapping.setdefault(new_isin, data)

    print(
        f"  -> Loaded market cap mapping: "
        f"{mapping_file.name} | {len(mapping)} ISINs"
    )

    return mapping

def get_market_cap_file(year, month_num):
    """
    Select the correct six-month AMFI market-cap file.

    Jan-Jun  -> June mapping
    Jul-Dec  -> December mapping
    """

    year_dir = MARKET_CAP_ROOT / str(year)

    if month_num <= 6:
        mapping_file = (
            year_dir /
            f"market_cap_June_{year}.xlsx"
        )
    else:
        mapping_file = (
            year_dir /
            f"market_cap_December_{year}.xlsx"
        )

    if not mapping_file.exists():
        # The AMFI list for the current half-year is only published after it
        # ends (e.g. Dec 2026 list in Jan 2027). Until then use the latest
        # earlier list instead of skipping the disclosure.
        candidates = []
        for f in MARKET_CAP_ROOT.glob("*/market_cap_*_*.xlsx"):
            m = re.match(r"market_cap_(June|December)_(\d{4})\.xlsx$", f.name)
            if m:
                candidates.append(((int(m.group(2)), 6 if m.group(1) == "June" else 12), f))
        earlier = [c for c in candidates if c[0] < (year, 6 if month_num <= 6 else 12)]
        if not earlier:
            raise FileNotFoundError(
                f"\nMarket cap mapping required but not found:\n"
                f"  Year: {year}\n"
                f"  Month: {month_num}\n"
                f"  Expected file: {mapping_file}\n"
            )
        fallback = max(earlier)[1]
        print(f"  NOTE: {mapping_file.name} not available yet - using latest list {fallback.name}")
        return fallback

    return mapping_file

MARKET_CAP_CACHE = {}

def get_market_cap_mapping(year, month_num):
    """
    Return the appropriate market-cap mapping.

    Mapping is loaded only once per year/half-year
    and then reused from memory.
    """

    period = "H1" if month_num <= 6 else "H2"
    cache_key = (year, period)

    if cache_key in MARKET_CAP_CACHE:
        return MARKET_CAP_CACHE[cache_key]

    mapping_file = get_market_cap_file(
        year,
        month_num
    )

    mapping = load_market_cap_mapping(
        mapping_file
    )

    MARKET_CAP_CACHE[cache_key] = mapping

    return mapping

ISIN_ALIASES = None

def get_isin_aliases():
    """
    old <-> new ISIN pairs from market_cap_mapping/isin_changes_*.csv,
    grouped so every ISIN of a company points at all its other ISINs.
    """
    global ISIN_ALIASES
    if ISIN_ALIASES is None:
        groups = {}
        for f in sorted(MARKET_CAP_ROOT.glob("isin_changes_*.csv")):
            for _, row in pd.read_csv(f, dtype=str).iterrows():
                old, new = norm(row.get("old") or "").upper(), norm(row.get("new") or "").upper()
                if not (old and new and ISIN.match(old) and ISIN.match(new)):
                    continue
                merged = groups.get(old, {old}) | groups.get(new, {new})
                for i in merged:
                    groups[i] = merged
        ISIN_ALIASES = {i: [a for a in g if a != i] for i, g in groups.items()}
    return ISIN_ALIASES


def available_market_cap_periods():
    """(year, half) of every AMFI file present, half = 0 for June, 1 for December."""
    periods = []
    for f in MARKET_CAP_ROOT.glob("*/market_cap_*.xlsx"):
        m = re.fullmatch(r"market_cap_(June|December)_(\d{4})\.xlsx", f.name)
        if m:
            periods.append((int(m.group(2)), 0 if m.group(1) == "June" else 1))
    return periods


def find_in_other_periods(isin, year, month_num):
    """
    Look the ISIN (or its old/new alias) up in the other half-year AMFI files,
    nearest period first (later before earlier on a tie).
    """
    here = year * 2 + (0 if month_num <= 6 else 1)
    periods = sorted(
        (p for p in available_market_cap_periods() if p[0] * 2 + p[1] != here),
        key=lambda p: (abs(p[0] * 2 + p[1] - here), -(p[0] * 2 + p[1])),
    )
    candidates = [isin] + get_isin_aliases().get(isin, [])
    for y, half in periods:
        mapping = get_market_cap_mapping(y, 6 if half == 0 else 12)
        for c in candidates:
            if c in mapping:
                return dict(mapping[c], isin=isin, period=f"{'June' if half == 0 else 'December'} {y}")
    return None


def get_market_cap_data(isin, year, month_num):
    """
    Return NSE symbol and market-cap category
    for an ISIN using the appropriate six-month mapping.
    """

    # AMFI's list only covers Indian ISINs; overseas holdings have no SEBI market-cap bucket
    if not is_domestic_isin(isin):
        return {
            "symbol": None,
            "market_cap_category": None
        }

    mapping = get_market_cap_mapping(
        year,
        month_num
    )

    data = mapping.get(isin)

    # ISIN changed (split / face-value change): the period's AMFI list may carry
    # the other ISIN of the same company (e.g. Shriram Finance INE721A01013 <-> INE721A01047)
    if data is None:
        for alias in get_isin_aliases().get(isin, ()):
            if alias in mapping:
                data = dict(mapping[alias], isin=isin)
                break

    # Not in this period's list (IPO after the cut-off, e.g. Vishal Mega Mart in
    # H1 2024, or AMC files re-issued with today's ISIN): use the nearest period
    if data is None:
        data = find_in_other_periods(isin, year, month_num)
        if data is not None:
            print(
                f"  INFO: Market cap from nearest period {data['period']} | "
                f"Year={year} | Month={month_num} | ISIN={isin}"
            )

    # Listed after the latest half-year list: historical-mf's AMFI list, which adds the newer listings
    if data is None:
        data = get_supplementary().get(isin)
        if data is not None:
            print(f"  INFO: Market cap from {data['period']} | Year={year} | Month={month_num} | ISIN={isin}")

    # Listed after the latest AMFI list: provisional values, only until a list covers the ISIN
    if data is None:
        data = get_isin_override(isin)

    if data is None:
        print(
            f"  WARNING: Market cap mapping missing | "
            f"Year={year} | Month={month_num} | ISIN={isin}"
        )

        return {
            "symbol": None,
            "market_cap_category": None
        }

    return data

SUPPLEMENTARY_DIR = MARKET_CAP_ROOT / "supplementary"
SUPPLEMENTARY = None


def get_supplementary():
    """
    ISIN -> AMFI entry from market_cap_mapping/supplementary/*.xlsx: historical-mf's AMFI list
    (api/input/AMFI_data_JAN_JUL_2026.xlsx, columns Company Name / ISIN Code / BSE Symbol / NSE Symbol /
    Market Cap), the June 2026 list plus the listings after it (Indo-MIM, Molbio Diagnostics, ...). Used only
    for ISINs no half-year list covers; its Industry / Sector columns are another taxonomy and are not used.
    """
    global SUPPLEMENTARY
    if SUPPLEMENTARY is None:
        SUPPLEMENTARY = {}
        for f in sorted(SUPPLEMENTARY_DIR.glob("*.xlsx"), reverse=True):  # later file names first
            df = pd.read_excel(f, dtype=str)
            if not {"Company Name", "ISIN Code", "Market Cap"} <= set(df.columns):
                print(f"  WARNING: {f.name} is not an AMFI list (Company Name / ISIN Code / Market Cap), skipped")
                continue
            for _, row in df.iterrows():
                isin = norm(row["ISIN Code"] if pd.notna(row["ISIN Code"]) else "").upper()
                if not ISIN.match(isin) or isin in SUPPLEMENTARY:
                    continue
                cat = norm(row["Market Cap"] if pd.notna(row["Market Cap"]) else "")
                SUPPLEMENTARY[isin] = {
                    "isin": isin,
                    "symbol": clean_symbol(row.get("NSE Symbol")) or clean_symbol(row.get("BSE Symbol")) or None,
                    "market_cap_category": cat if cat in ("Large Cap", "Mid Cap", "Small Cap") else None,
                    "company": norm(row["Company Name"]) if pd.notna(row["Company Name"]) else "",
                    "industry": None,
                    "period": f.stem,
                }
    return SUPPLEMENTARY


PARENT_INDEX_CACHE = {}

# words that describe the special instrument rather than the company
PARENT_NOISE_RX = re.compile(
    r"\([^)]*\)|\b\d+(\.\d+)?\s*%|\b\d{1,2}[-./]\w{2,4}[-./]\d{2,4}\b|\b\d{2}[a-z]{3}\d{2}\b|"
    r"\bpartly\s*paid\s*(up)?(\s*shares?)?\b|\bright(s)?\s*(entitlements?|issue|shares?)?\b|\brights\w*|"
    r"\b(non\s*conv(ertible)?|rede(emable)?|ncd|ncrps|pref(erence)?|shares?|unlisted|dvr|class\s*a|covered\s*call|fv\s*\d+|md)\b",
    re.I,
)
COMPANY_STOPWORDS = {"ltd", "limited", "the", "and", "company", "co", "corp", "corporation", "india", "inc"}


def company_key(name):
    """Normalise a company / instrument name for parent-company matching."""
    n = str(name or "").lower().replace("&", " and ").replace("\u2019", "").replace("'", "")
    n = PARENT_NOISE_RX.sub(" ", n)
    n = re.split(r"\s+-\s+", n)[0]
    words = [w for w in re.findall(r"[a-z0-9]+", n) if w not in COMPANY_STOPWORDS]
    return " ".join(words)


def get_parent_index(year, month_num):
    """
    (by_company_code, by_name) indexes over the AMFI list's ordinary equity
    ISINs (INE<code>01...), cached per half-year like the market-cap mapping.
    """
    period = (year, "H1" if month_num <= 6 else "H2")
    if period not in PARENT_INDEX_CACHE:
        by_code, by_name = {}, {}
        for isin, data in get_market_cap_mapping(year, month_num).items():
            if not (isin.startswith("INE") and isin[7:9] == "01"):
                continue
            by_code.setdefault(isin[3:7], data)
            key = company_key(data.get("company"))
            if key:
                by_name.setdefault(key, data)
        PARENT_INDEX_CACHE[period] = (by_code, by_name)
    return PARENT_INDEX_CACHE[period]


MANUAL_OVERRIDES_FILE = MARKET_CAP_ROOT / "manual_overrides.csv"
MANUAL_OVERRIDES = None


ISIN_OVERRIDES = None


def get_isin_override(isin):
    """manual_overrides.csv rows with an ISIN and a symbol or market cap (recent listings no AMFI list covers yet)."""
    global ISIN_OVERRIDES
    if ISIN_OVERRIDES is None:
        ISIN_OVERRIDES = {}
        if MANUAL_OVERRIDES_FILE.exists():
            for _, row in pd.read_csv(MANUAL_OVERRIDES_FILE, dtype=str).fillna("").iterrows():
                code = norm(row["isin"]).upper()
                if code and (norm(row["symbol"]) or norm(row["market_cap_category"])):
                    ISIN_OVERRIDES[code] = {
                        "isin": code,
                        "symbol": norm(row["symbol"]) or None,
                        "market_cap_category": norm(row["market_cap_category"]) or None,
                    }
    return ISIN_OVERRIDES.get(isin)


def find_manual_override(name, year, month_num):
    """
    Hand-maintained entries for equity printed without an ISIN that no AMFI
    list covers yet (e.g. ALLCARGO GLOBAL LTD. after the Allcargo demerger,
    awaiting listing). Matched on company_key(name); a blank symbol / market
    cap is filled from the AMFI lists when the override gives an ISIN.
    """
    global MANUAL_OVERRIDES
    if MANUAL_OVERRIDES is None:
        MANUAL_OVERRIDES = {}
        if MANUAL_OVERRIDES_FILE.exists():
            for _, row in pd.read_csv(MANUAL_OVERRIDES_FILE, dtype=str).fillna("").iterrows():
                key = company_key(row["company_name"])
                if key:
                    MANUAL_OVERRIDES[key] = {
                        "isin": norm(row["isin"]).upper() or None,
                        "symbol": norm(row["symbol"]) or None,
                        "market_cap_category": norm(row["market_cap_category"]) or None,
                    }
    entry = MANUAL_OVERRIDES.get(company_key(name))
    if entry is None:
        return None
    data = dict(entry)
    if data["isin"] and not (data["symbol"] and data["market_cap_category"]):
        amfi = get_market_cap_data(data["isin"], year, month_num)
        data["symbol"] = data["symbol"] or amfi["symbol"]
        data["market_cap_category"] = data["market_cap_category"] or amfi["market_cap_category"]
        data["industry"] = amfi.get("industry")
    return data


def find_parent_company(isin, name, year, month_num):
    """
    Listed parent of a partly-paid / preference / rights / DVR / covered-call
    line: first by the ISIN's company code (IN9397D... -> INE397D01...),
    then by the instrument name. Returns the AMFI entry or None.
    """
    by_code, by_name = get_parent_index(year, month_num)
    if isin and len(isin) >= 7 and isin[3:7] in by_code:
        return by_code[isin[3:7]]
    key = company_key(name)
    if not key:
        return None
    if key in by_name:
        return by_name[key]
    # "tvs motor" vs "tvs motor company": accept a unique whole-word prefix match
    hits = [d for k, d in by_name.items() if k.startswith(key + " ") or key.startswith(k + " ")]
    return hits[0] if len(hits) == 1 else None


# derivative_exposure_pct: Canara "Outstanding derivative exposure as % to net assets Long / (Short)" (per-stock hedge,
# no separate futures lines); listed first so "exposure" / "% to net assets" do not claim the column
HEADER_MAP = [("derivative_exposure_pct", r'derivative exposure'), ("isin", r'\bisin'), ("coupon", r'coupon'), ("instrument_name", r'name|instrument'), ("industry_rating", r'industry|rating'),
              ("quantity", r'quantity'), ("market_value", r'market|mkt val|exposure'), ("pct_nav", r'% ?to|% of net|to nav|net assets|% to aum'),
              ("ytc", r'ytc|yield to call'), ("yield", r'yield|ytm'), ("maturity_date", r'maturity'), ("put_call", r'put/call'),
              ("market_cap", r'market capitali'), ("derivative_pct", r'^derivative( % to nav)?$'), ("unhedged_pct", r'^unhedged( % to nav)?$'), ("notes", r'notes'),
              ("serial_no", r'sr\.? ?no|sl no')]

def map_header(row):
    cols = {}
    for c, v in row.items():
        if pd.isna(v): continue
        h = norm(v).lower()
        if "risk-o-meter" in h or h in ("sector / rating", "percent", "sector/rating"): continue
        for fld, rx in HEADER_MAP:
            if fld in cols: continue
            if fld == "instrument_name" and "isin" in h: continue
            if fld == "market_value" and "capitali" in h: continue
            if fld == "industry_rating" and "name" in h: continue
            if re.search(rx, h): cols[fld] = c; break
    return cols

def kotak_fallback_name(txts, cells, cols):
    """Name when the header's name column is empty. Kotak puts a coupon tag ("FRB", "ZCB") in its own column
    left of the name and prints "Total" / "Grand Total" in the Industry / Rating column."""
    txts = [v for v in txts if not re.match(r'^\s*(FRB|FRN|ZCB|VRR|STRIPS)\s*$', v, re.I)]
    name = txts[0] if txts else ""
    ir = cells.get(cols.get("industry_rating")) if "industry_rating" in cols else None
    if not str(name).strip() and isinstance(ir, str) and re.match(r'^\s*(grand total|sub ?total|total)\s*$', ir, re.I):
        name = ir
    return name


END_RX = re.compile(r'^(grand total|net assets?:?$|total net assets|total : (?!others).*)', re.I)
TOTAL_RX = re.compile(r'(^|\b)(sub ?total|total)\b', re.I)
# rows that carry a value but are really section headers (ICICI style)
# ICICI Flexi Cap Jan 2024: "Preference Shares" heading carries the block total (listed + unlisted pref shares), so the
# look-ahead aggregate check misses it and it would become a holding with no ISIN
SPECIAL_EQUITY_HEADING = r'(preference shares?|partly paid( up)? shares?|rights( entitlements?)?)'
VALUED_SECTION = re.compile(r'^(equity & equity related instruments.*|foreign securities.*|privately placed/unlisted|[b-e]\) listed/awaiting listing.*|listed / awaiting listing on stock exchanges|money market instruments|treasury bills|others|debt instruments|unlisted|certificate of deposits|commercial papers|term deposits|units of .*|' + SPECIAL_EQUITY_HEADING + r')$', re.I)
SECTION_LIKE = re.compile(r'^(\(?[a-fA-F]\)\s*(?!repo\b)|equity|foreign securities|debt|money market|others?$|derivatives|reit|mutual fund units|exchange traded|cash & cash|other current|treps$|treps /|treps - tri|triparty repo/|cblo/|reverse repo /|short term deposits|treasury bill|units of|term deposits|deposits|certificate|commercial|privately|securiti|unlisted|listed|index / stock|' + SPECIAL_EQUITY_HEADING + r'$)', re.I)
SYMBOLS = re.compile(r'(\*\*|\*|£|@|#|\^|\$|~|&$)\s*$')

def split_symbols(name):
    syms = []
    while True:
        m = SYMBOLS.search(name)
        if not m or len(name) < 3: break
        syms.insert(0, m.group(1)); name = name[:m.start()].rstrip()
    # "A**" = awaiting listing (Altiva, Apex, Platinum, Magnum): the "A" belongs to the marker, not the name
    if "**" in syms and re.search(r"\sA$", name):
        name = name[:-1].rstrip()
        syms.insert(0, "A")
    return name, "".join(syms)

RATING_RX = re.compile(r'\b(CRISIL|ICRA|CARE|IND|FITCH|BWR|ACUITE|SOV|SOVEREIGN)\b|^(AAA|AA|A1\+)', re.I)

# ISIN suffixes for special equity instruments (same as the API pipeline)
SPECIAL_ISIN_SUFFIX = {
    "Preference Shares": "-PS",
    "Partly Paid Shares": "-PP",
    "Rights Entitlements": "-R",
    "Listed Equity - DVR": "-DVR",
}
COVERED_CALL_SUFFIX = "-CC"
# special equity whose own ISIN (when printed) is used instead of parent ISIN + suffix
OWN_ISIN_SUB_TYPES = ("Partly Paid Shares", "Preference Shares", "Rights Entitlements")


# Units of an Alternative Investment Fund: their own asset class (see classify)
AIF_ASSET_CLASS = "AIF"
AIF_SUB_TYPE = "AIF Units"
AIF_HEADING = re.compile(r'alternative investment|\(aif\)|\baif units?\b', re.I)
AIF_NAME = re.compile(r'corp(orate|\.)? debt (market|mkt)\.? dev|\bcdmdf\b', re.I)
AIF_ISINS = {"INF0RQ622028"}  # CDMDF Class A2

# Money market instruments are classified under "Debt" (Neo4j schema has no Money Market asset class)
MONEY_MARKET_SUB_TYPES = {"Treasury Bills", "Commercial Paper", "Certificate of Deposit"}
# Futures/options are classified under "Equity" with subType "derivative-..." (Neo4j schema has no Derivatives asset class)
# Commodities (Multi Asset funds) are their own asset class, split by sub type:
#   Physical   - Edelweiss "Others > a) Gold / Silver", Tata "A) COMMODITIES PHYSICAL"
#   Derivative - exchange traded commodity futures / options: Axis "(b) Commodity Futures (ETCD)",
#                Edelweiss "(b) Exchange Traded Commodity Derivatives", ICICI "Exchange Traded Commodity Derivatives >
#                A) LISTED ON COMMODITY EXCHANGES (Quantity in Lots)", Nippon "Commodity Options" and
#                "Details of Commodity Future / Index Future", Tata "B) LISTED ON COMMODITY EXCHANGES (Quantity in Lots)"
COMMODITY_PHYSICAL, COMMODITY_DERIVATIVE = "Physical", "Derivative"
# the commodity "Derivative" sub type shares the synthetic ISIN / Long-Short handling of equity F&O
# SIF hybrid funds also hedge rates with interest rate swap legs (Debt, synthetic ISIN INR-IRS-XXXXXX)
IRS_SUB_TYPE = "interest-rate-swaps-derivative"
DERIVATIVE_SUB_TYPES = {"Index Futures", "Stock Futures", "Index Options", "Stock Options", "Currency Futures", COMMODITY_DERIVATIVE, IRS_SUB_TYPE}
DERIVATIVE_CODES = {"Index Futures": "IF", "Index Options": "IO", "Stock Futures": "SF", "Stock Options": "SO", "Currency Futures": "CF", COMMODITY_DERIVATIVE: "CO", IRS_SUB_TYPE: "IRS"}
# commodity contract names: ICICI "Gold (1 KG-1000 GMS) Commodity April 2024 Future", Axis "Silver March 2026 Commodity Future",
# Tata "GOLD (1 KG-1000 GMS) COMMODITYFEB2024CFUT", Nippon "FUTCOM_SILVER_03/05/2024", Edelweiss "GOLDMINI-05Feb2026-MCX"
COMMODITY_DERIVATIVE_NAME = re.compile(r'commodity\s*[a-z]+\s*\d{4}\s*(future|cfut)|commodity future|^futcom_|^(gold|silver)\w*-\d', re.I)


def commodity_name(name):
    """Underlying commodity: "GOLD MINI (100 GRAMS) COMMODITY" -> GOLD, "FUTCOM_CRUDEOIL_19/03/2024" -> CRUDEOIL."""
    m = re.match(r"[A-Za-z]+", re.sub(r"^(physical\s+|futcom_)", "", str(name or "").strip(), flags=re.I))
    return m.group(0).upper() if m else "COMMODITY"


def is_futures(df):
    """Futures rows (stock / index / currency / commodity, not options) for the grand-total exclusion checks."""
    st, nm = df.asset_sub_type.fillna(""), df.instrument_name.fillna("")
    return st.str.contains("Futures") | ((st == COMMODITY_DERIVATIVE) & ~nm.str.contains(r'option|\b(?:call|put)\b', case=False))


def derivative_isin(sub_type, name):
    """
    Synthetic ISIN for a derivative contract:
      index F&O -> NIFTY-IF-XXXXXX / NIFTY-IO-XXXXXX
      stock F&O -> <first letter of each word>-SF-XXXXXX / -SO-XXXXXX
    XXXXXX is a 6-char alphanumeric hash of the contract name, so the same
    contract gets the same ISIN every month and on every re-run.
    """
    code = DERIVATIVE_CODES[sub_type]
    name = re.sub(r"\s+", " ", str(name or "")).strip()
    pair = re.search(r"(USD|EUR|GBP|JPY)(INR)", name.upper())
    prefix = "NIFTY" if code in ("IF", "IO") else "INR" if code == "IRS" else (pair.group(0) if code == "CF" and pair else (
        commodity_name(name) if code == "CO" else "".join(w[0] for w in re.findall(r"[A-Za-z0-9]+", re.sub(r"-[A-Za-z]{3}\d{4}$", "", name))).upper()))
    n = int(hashlib.sha1(f"{code}|{name.upper()}".encode("utf-8")).hexdigest(), 16)
    chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    suffix = "".join(chars[(n >> (6 * k)) % 36] for k in range(6))
    return f"{prefix}-{code}-{suffix}"

def isin_security_type(isin):
    """
    Security-type code of a corporate ISIN (characters 8-9 of INE<co>NN<serial><chk>):
    01 equity, 07/08/09 debentures/bonds, 14 commercial paper, 15 securitised debt (PTC), 16 certificate of deposit, ...
    Once an issuer exhausts a series NSDL swaps the leading digit for a letter and keeps the
    last one: Axis Bank CD INE238AD6579 -> "D6" = 16, Kotak Mahindra Prime NCD INE916DA7RW2 -> "A7" = 07.
    """
    if not (isin and len(isin) == 12 and isin.startswith("INE")):
        return ""
    code = isin[7:9]
    if re.fullmatch(r'[A-Z]\d', code):
        return {"6": "16", "5": "15", "4": "14", "7": "07", "8": "08", "9": "09"}.get(code[1], code)
    return code


def classify(sec, name, isin, ind):
    s, n, i, il = sec.lower(), name.lower(), isin or "", (ind or "").lower()
    # iSIF Feb 2026 prints "Eternal Ltd (Put Option)" with the stock's own ISIN: a contract, not a second holding of it
    if re.search(r'\((?:put|call)\s+option\)', n): return "Equity", "Index Options" if re.search(r'nifty|sensex|\bindex\b', n) else "Stock Options"
    st9 = isin_security_type(i)
    if re.search(r'net receivable|net current (asset|liabilit)|current assets|net payable|cash / net|^accrued interest$', n): return "Cash and Cash Equivalents", "Net Receivables / (Payables)"
    if "margin" in n and not i: return "Cash and Cash Equivalents", "Margin Money (Derivatives)"
    if re.search(r'cash and bank|cash & bank', n): return "Cash and Cash Equivalents", "Cash & Bank Balance"
    if re.search(r'^cash & cash equivalents?:?$', n): return "Cash and Cash Equivalents", "Cash & Equivalents (Aggregate)"
    # TREPS lines by name first: Edelweiss prints the TREPS block straight after the Derivatives block (section lineage still says derivatives)
    if re.search(r'treps|tri-?party|reverse repo|cblo|^repo$|^[a-z]\) repo$', n) or (n.startswith("clearing corporation of india") and "std" not in n): return "Cash and Cash Equivalents", "TREPS / Reverse Repo"
    # Physical commodities held by multi-asset funds (Edelweiss "Others > a) Gold": row "Gold", AMC code IDIA00500001;
    # Tata "A) COMMODITIES PHYSICAL": "SILVER (30 KG) COMMODITY")
    if re.fullmatch(r'(physical\s+)?(gold|silver)(\s+bars?.*)?', n) or re.search(r'commodit\w*\s+physical|physical\s+commodit', s):
        return "Commodities", COMMODITY_PHYSICAL
    # Exchange-traded commodity derivatives (MCX futures / options)
    if not (i and i.startswith("IN")) and (re.search(r'commodit', s) or COMMODITY_DERIVATIVE_NAME.search(n)):
        return "Commodities", COMMODITY_DERIVATIVE
    # SBI writes options as "NIFTY28-Oct-2025CE24700" under an "Index Options" heading
    # a real Indian ISIN (equity, REIT/InvIT units, bonds) is never an F&O contract, even under a "Derivatives" heading
    is_security = bool(i) and i.startswith("IN")
    if not is_security and (re.search(r'\b(put|call)\b', n) or re.search(r'\d(ce|pe)\d{3,6}$', n)) and ("option" in il or "option" in s or "deriv" in s): return "Equity", "Index Options" if re.search(r'nifty|sensex|bank nifty', n) else "Stock Options"
    # Altiva (Edelweiss) lists option legs by the underlying's name only, under "(B)Index / Stock Option"
    if not is_security and re.search(r'\boptions?$', s): return "Equity", "Index Options" if re.search(r'nifty|sensex|\bindex\b', n) else "Stock Options"
    # Interest rate swap legs ("Interest Rate Swap - 30/07/2029 Fixed" / "... Floating", Arudha): inside net assets
    if not is_security and "interest rate swap" in n: return "Debt", IRS_SUB_TYPE
    # Tata prints options as "NIFTYNSEJUN2026OPT P" inside the listed-equity block, with no derivatives heading
    if not is_security and re.search(r'\d{4}\s*opt\s*[pc]$', n): return "Equity", "Index Options" if re.search(r'nifty|sensex', n) else "Stock Options"
    # ICICI lists written calls as "<Company> (Covered call)" inside the listed-equity block, without an ISIN
    if "covered call" in n and not i: return "Equity", "Index Options" if re.search(r'nifty|sensex', n) else "Stock Options"
    if not is_security and (re.search(r'future', n + " " + s + " " + il) or "deriv" in s or re.search(r'_\(?\d\d/\d\d/\d{4}\)?$', n) or re.search(r'\dfut$', n)):
        return "Equity", "Index Futures" if re.search(r'nifty|sensex|index', n + il) else "Stock Futures"
    if re.search(r'treps|tri-?party|reverse repo|^a\) repo|cblo|collateralized borrowing', n + " || " + s) or (n.startswith("clearing corporation of india") and "std" not in n):
        return "Cash and Cash Equivalents", "TREPS / Reverse Repo"
    # Overseas stocks, ADRs, ETFs and funds (US..., FR..., LU..., IE... ISINs) -> International Equity
    if i and not i.startswith("IN"): return "Equity", "Foreign Equity"
    if re.search(r'std - margin|short term deposit', n + s): return "Cash and Cash Equivalents", "Short Term Deposits"
    if re.search(r'term deposit|deposits with|fixed deposit', n + s): return "Cash and Cash Equivalents", "Term Deposits"
    # REIT / InvIT units before the MF check: ABSL (Jan / Feb 2024) prints them under "Mutual Fund Units > InvIT / ReIT"
    # (India Grid Trust INE219X23014, Embassy Office Parks REIT INE041025011); bonds issued by the trusts stay Debt
    if not i.startswith("INF") and st9 not in ("07", "08", "09", "15"):
        if st9 == "25" or re.search(r'\breit\b|real estate investment trust', n): return "REITs & InvITs", "REIT Units"
        if st9 == "23" or re.search(r'infrastructure investment trust|\binvit\b', n): return "REITs & InvITs", "InvIT Units"
    # AIF units are their own asset class: the only AIF held is the Corporate Debt Market Development Fund (CDMDF, INF0RQ622028).
    # ICICI Conservative Hybrid prints it under "Units of an Alternative Investment Fund (AIF)"; Sundaram / PPFAS /
    # HSBC / Baroda BNP list it with no AIF heading, so match the name and ISIN too
    if i.startswith("INF") or AIF_HEADING.search(s):
        if AIF_HEADING.search(s) or AIF_NAME.search(n) or i in AIF_ISINS: return AIF_ASSET_CLASS, AIF_SUB_TYPE
    if i.startswith("INF") or re.search(r'mutual fund units|exchange traded|\betf\b', n + s):
        if re.search(r'\betf\b|exchange traded', n): return "MF and ETF", "Exchange Traded Funds"
        return "MF and ETF", "Mutual Fund Units"
    # bonds / debentures (incl. CCDs) keep their debt ISIN type even inside a REIT/InvIT block (ICICI Multi Asset Aug 2026: Motherson CCD)
    if st9 in ("07", "08", "09"): return "Debt", "Corporate Bonds / Debentures"
    # 15 = PTCs issued by securitisation trusts (Sansar Trust INE0QTL15013, India Universal Trust INE1CBK15037)
    if st9 == "15" or re.search(r'securiti[sz]ed debt', s): return "Debt", "Securitised Debt"
    if st9 == "25" or re.search(r'\breit\b|real estate investment trust|real estate trust', n + s): return "REITs & InvITs", "REIT Units"
    if st9 == "23" or "infrastructure investment trust" in n or "invit" in n: return "REITs & InvITs", "InvIT Units"
    # Trust a "Preference Shares" heading: Invesco (Aug 2025) prints the pref share without an ISIN, named only by
    # the company, and PGIM uses a placeholder ISIN (IN25H25DUM01) that would otherwise match the SDL pattern below
    if st9 == "04" or "preference" in n or "ncrps" in n or "preference share" in s: return "Equity", "Preference Shares"
    if re.match(r'IN002\d{3}[A-Z]', i) or re.search(r't-bill|tbill|treasury bill', n): return "Debt", "Treasury Bills"
    if st9 == "14" or "commercial paper" in n or "commercial paper" in s: return "Debt", "Commercial Paper"
    # st9 already maps alphanumeric CD codes (Axis Bank INE238AD6579 -> 16); the heading still catches CDs without an ISIN
    if st9 == "16" or "certificate of deposit" in n or "certificate of deposit" in s: return "Debt", "Certificate of Deposit"
    if re.match(r'IN00\d', i) or "government of india" in n: return "Debt", "Central Government Securities"
    if re.match(r'IN[1-3]\d', i) or "state government" in n or " sdl" in n: return "Debt", "State Development Loans"
    if st9 in ("07", "08", "09") or ("debt" in s and i): return "Debt", "Corporate Bonds / Debentures"
    # bonds printed without an ISIN, under a bonds heading (workbooks built from a factsheet: "GMR Airports Ltd (13/08/2028)"
    # with only a rating, which would otherwise read as listed equity below)
    if not i and re.search(r'corporate bonds?|debentures?|\bncds?\b', s): return "Debt", "Corporate Bonds / Debentures"
    if "partly paid" in n or "(pp)" in n: return "Equity", "Partly Paid Shares"
    if "dvr" in n: return "Equity", "Listed Equity - DVR"
    if i == "IN9155A01020": return "Equity", "Listed Equity - DVR"
    if i.startswith("IN9"): return "Equity", "Partly Paid Shares"
    if "rights" in n or st9 == "20": return "Equity", "Rights Entitlements"
    if "warrant" in n: return "Equity", "Warrants"
    if "overseas" in s or "foreign" in s: return "Equity", "Foreign Equity"
    if re.search(r'unlisted|privately placed', s): return "Equity", "Unlisted Equity"
    if "dvr" in n: return "Equity", "Listed Equity - DVR"
    if i.startswith("IN") or il: return "Equity", "Listed Equity"
    # No asset class: the row is not an instrument (heading / aggregate the
    # generic subtotal check missed). parse_file drops it and logs it for review.
    return None, None

UNCLASSIFIED_ROWS = []

SUM_COLUMNS = ("quantity", "market_value_lakhs", "market_value_inr", "pct_to_nav", "margin_lakhs")


def merge_duplicate_isins(H):
    """
    One row per ISIN per file and record type: locked-in + free shares of the
    same company, hedged + unhedged blocks (Tata), several strikes of the same
    option name, etc. are summed so Neo4j's unique-ISIN constraint holds.
    The kept row is the first one; its name is the shortest variant
    ("Premier Energies Limited" rather than "... -Locked IN").
    """
    if H.empty or "isin" not in H:
        return H
    key = H["isin"].notna() & (H["isin"].astype(str) != "")
    # included / excluded rows stay apart so the NAV reconciliation still holds (hedge legs the AMC leaves out)
    gkey = ["record_type", "isin", "included_in_net_assets"]
    dup = key & H.duplicated(gkey, keep=False)
    if not dup.any():
        return H
    rows = []
    for _, g in H[dup].groupby(gkey, sort=False):
        first = g.iloc[0].copy()
        for c in SUM_COLUMNS:
            if c in g and g[c].notna().any():
                first[c] = g[c].sum(skipna=True)
        first["instrument_name"] = min(g.instrument_name.astype(str), key=len)
        first["pct_below_threshold"] = None if first.get("pct_to_nav") is not None and abs(first["pct_to_nav"] or 0) >= 0.01 and g.pct_below_threshold.isna().any() else first.get("pct_below_threshold")
        fs = first.get("footnote_symbols")
        first["footnote_symbols"] = (fs if isinstance(fs, str) else "") + f" [merged {len(g)} rows: {','.join(str(r) for r in g.source_row)}]"
        rows.append(first)
    merged = pd.DataFrame(rows)
    out = pd.concat([H[~dup], merged]).sort_values(["record_type", "source_row"], kind="stable")
    return out.reset_index(drop=True)

def find_aggregate_rows(df, hdr, cols):
    """
    Find rows without an ISIN whose market value is the sum of neighbouring
    instrument rows, i.e. aggregates printed like holdings:

      header    -> value == sum of the rows directly below it, up to the next
                   blank / total row (ICICI "Equity shares",
                   "Foreign Securities/Overseas ETFs")
      subtotal  -> section-like label whose value == sum of the rows directly
                   above it (ABSL Nov 2024 "TREPS / Reverse Repo" carrying
                   the Margin + Cash subtotal)

    Returns {row_index: "header" | "subtotal"}.
    """
    rows = {}
    for i in range(hdr + 1, len(df)):
        row = df.iloc[i]
        cells = {c: v for c, v in row.items() if pd.notna(v) and str(v).strip() != ""}
        if not cells:
            rows[i] = None
            continue
        name = cells.get(cols.get("instrument_name")) if "instrument_name" in cols else None
        if name is None or tonum(name) is not None:
            txts = [v for c, v in sorted(cells.items()) if isinstance(v, str) and not ISIN.match(v.strip()) and c != cols.get("industry_rating")]
            name = kotak_fallback_name(txts, cells, cols)
        name = norm(name)
        isin_cell = norm(cells.get(cols.get("isin"), "")) if "isin" in cols else ""
        rows[i] = dict(name=name, mv=tonum(cells.get(cols.get("market_value"))) if "market_value" in cols else None,
                       isin=bool(ISIN.match(isin_cell)))
        if END_RX.match(name):
            break

    def is_stop(r):
        return r is None or r["mv"] is None or TOTAL_RX.search(r["name"]) or END_RX.match(r["name"])

    found = {}
    idx = sorted(rows)
    for i in idx:
        r = rows[i]
        if r is None or r["isin"] or r["mv"] is None or abs(r["mv"]) < 0.01 or not r["name"]:
            continue
        if TOTAL_RX.search(r["name"]) or END_RX.match(r["name"]):
            continue
        # look ahead: heading that carries the block total
        kids, j = [], i + 1
        while j in rows and not is_stop(rows[j]):
            kids.append(rows[j]["mv"]); j += 1
        if kids:
            tol = 0.01 if len(kids) == 1 else max(0.5, 1e-5 * abs(r["mv"]))
            if abs(sum(kids) - r["mv"]) <= tol:
                found[i] = "header"
                continue
        # look behind: unlabelled subtotal printed on a section-like label
        if SECTION_LIKE.match(r["name"]):
            prev, k = [], i - 1
            while k in rows and rows[k] is not None and rows[k]["mv"] is not None and not END_RX.match(rows[k]["name"]):
                prev.append(rows[k]["mv"]); k -= 1
            if len(prev) >= 2 and abs(sum(prev) - r["mv"]) <= max(0.5, 1e-5 * abs(r["mv"])):
                found[i] = "subtotal"
    return found

# derivative tables headed by the AMC's own hedging label, usually below the grand total
HEDGING_TABLE_RX = r'^(\(?[a-e][.)]\s*)?(other than )?hedging positions? through (stock |index )?(futures|options)'


def get_excel_engine(f):
    ext = os.path.splitext(f)[1].lower()

    if ext == ".xlsb":
        return "pyxlsb"

    with open(f, "rb") as fh:
        signature = fh.read(8)

    if signature == b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1':
        return "xlrd"

    if signature[:4] == b'PK\x03\x04':
        return "openpyxl"

    raise ValueError(
        f"Unknown Excel format: {f} | signature={signature.hex(' ')}"
    )


def find_header_row(df):
    """Row index of the portfolio table header (ISIN + name / quantity / market value), or None."""
    row_texts = [" ".join(str(v) for v in row if pd.notna(v)).lower() for _, row in df.iterrows()]
    hdr = next((i for i, t in enumerate(row_texts) if "isin" in t and ("instrument" in t or "name" in t)), None)
    if hdr is None:
        # Header with a blank name cell (e.g. HSBC Focused SEP 2024): ISIN + quantity/market value
        hdr = next((i for i, t in enumerate(row_texts) if re.search(r'\bisin\b', t) and re.search(r'quantity|market', t)), None)
    return hdr


# workbook sheets that never hold a scheme portfolio
NON_PORTFOLIO_SHEET = re.compile(r'^(index|derivative|disclaimer|notes?\b)|performance', re.I)


def parse_file(f):
    """
    Parse every portfolio sheet of one workbook. SIF workbooks may open on an "Index"
    sheet (Diviniti, Platinum), carry a "Derivative" annexure (iSIF) or a "Notes" sheet
    (RedHex), or hold one sheet per scheme. Returns [(holdings, snapshot, totals), ...].
    """
    eng = get_excel_engine(f)

    print(f"  -> Engine: {eng}")

    xf = pd.ExcelFile(f, engine=eng)
    sheets = {}
    for sh in xf.sheet_names:
        if NON_PORTFOLIO_SHEET.search(sh.strip()):
            continue
        d = xf.parse(sh, header=None)
        hdr = find_header_row(d)
        if hdr is not None:
            sheets[sh] = (d, hdr)

    if not sheets:
        raise ValueError(f"No portfolio sheet (ISIN header) in {xf.sheet_names}")

    return [parse_sheet(f, xf, sh, d, hdr, set(sheets)) for sh, (d, hdr) in sheets.items()]


def scheme_title(toptxt, sheet):
    """Scheme name from the text above the table header, else from the sheet name."""
    # cells may hold several lines: "HSBC Mutual Fund - SIF\nRedHex Hybrid Long-Short Fund\n..."
    lines = [norm(l) for t in toptxt for l in str(t).split("\n") if norm(l)]
    for t in lines:
        # skip "Summit SIF (offered by Invesco Mutual Fund)", "DIVINITI SIF Offered By ITI Mutual Fund"
        if re.search(r'mutual fund', t, re.I) and not SIF_TITLE_RX.search(t):
            continue
        name = canonical_scheme_name(re.sub(r'^' + re.escape(sheet.strip()) + r'\s+', '', t))
        if name:
            return name
    # Titanium May 2026 names the sheet after the scheme ("Titanium Equity Long Short Fund")
    return canonical_scheme_name(sheet)


def parse_sheet(f, xf, sheet, df, hdr, portfolio_sheets):
    parts = re.split(r"[\\/]", os.path.normpath(f))

    fname = parts[-1]
    amc_folder = parts[-2]
    month = parts[-3]
    year = int(parts[-4])
    amc_name = sif_brand(amc_folder)

    month_num = MONTH_NUM[month.upper()]
    last_day = calendar.monthrange(year, month_num)[1]
    portfolio_date = f"{year}-{month_num:02d}-{last_day:02d}"

    row_texts = [" ".join(str(v) for v in row if pd.notna(v)).lower() for _, row in df.iterrows()]
    cols = map_header(df.iloc[hdr])
    # a workbook built from a factsheet carries no total rows; the heuristic would take a row equal to the sum of
    # the rows after it for one (DynaSIF SEP 2026: two -0.03% option legs in a row)
    aggregate_rows = {} if fname.endswith(FROM_FACTSHEET_SUFFIX) else find_aggregate_rows(df, hdr, cols)
    top = df.iloc[:hdr]
    toptxt = [v for _, r in top.iterrows() for v in r if pd.notna(v) and isinstance(v, str)]
    # scheme name / code / date
    scheme = scheme_title(toptxt, sheet)
    if not scheme:
        raise ValueError(f"Scheme title not found on sheet {sheet!r}")
    scheme = current_scheme_name(scheme)
    master_row = get_scheme_master_row(scheme)
    code = master_row["scheme_code"] if master_row else None
    if master_row:
        amc_name = master_row["fund_house"]
    else:
        print(f"  WARNING: scheme not in sif_scheme_master.xlsx: {scheme}")
    toptxt = [norm(v) for v in toptxt]
    # "Derivative" annexure sheet (iSIF / HDFC style)
    deriv_sheet = next((sh for sh in xf.sheet_names if sh.lower().startswith("derivative")), None)
    dtxt = next((t for t in toptxt if re.search(r'as on|month ended|as of', t, re.I)), None)
    dval = next((v for _, r in top.iterrows() for v in r if hasattr(v, "year")), None)
    reported_date = str(dval.date()) if dval is not None else (dtxt or "")
    holdings, totals = [], []
    # Tata: hedge legs sit in the equity block under the stock's own ISIN, flagged "NAME^", with the
    # legend "^ Hedging positions through futures as on ..." (other AMCs use ^ for YTC / below-0.01% notes)
    caret_hedge = any(re.match(r'^\^\s*hedging positions through (commodity )?futures', t) for t in row_texts)
    # per-stock derivative column rows (Canara / HDFC): options when the HDFC derivative sheet only lists options
    # (Flexi Cap Jan/Feb 2024 covered calls), futures otherwise
    stock_derivative_rows, stock_derivative_type = False, "Stock Futures"
    if deriv_sheet:
        dtxt = " ".join(str(v) for v in xf.parse(deriv_sheet, header=None).values.ravel() if isinstance(v, str)).lower()
        if "option price" in dtxt and "futures price" not in dtxt:
            stock_derivative_type = "Stock Options"
    blank_nca = None   # "Net Receivables/(Payables)" printed with no value (Edelweiss Multi Asset Dec 2025)
    sec_path, grand = [], (None, None, None, None)
    last_i = hdr
    for i in range(hdr + 1, len(df)):
        row = df.iloc[i]
        cells = {c: v for c, v in row.items() if pd.notna(v) and str(v).strip() != ""}
        if not cells: continue
        g = lambda k: cells.get(cols.get(k)) if k in cols else None
        name = g("instrument_name")
        if name is None or tonum(name) is not None:
            txts = [v for c, v in sorted(cells.items()) if isinstance(v, str) and not ISIN.match(v.strip()) and c != cols.get("industry_rating") and not re.match(r'^[A-Z0-9_]{5,12}$', v.strip())]
            name = kotak_fallback_name(txts, cells, cols)
        name = norm(name)
        mv, pct_raw = tonum(g("market_value")), g("pct_nav")
        isin_cell = norm(g("isin") or "")
        isin = isin_cell if ISIN.match(isin_cell) else ""
        if END_RX.match(name): grand = (mv, tonum(pct_raw), name, i + 1); last_i = i; break
        if TOTAL_RX.search(name) and not isin and not re.search(r'expense', name, re.I):
            totals.append(dict(section_path=" > ".join(sec_path), total_label=name, market_value_lakhs=mv, pct_raw=pct_raw, source_row=i + 1)); continue
        if i in aggregate_rows and not VALUED_SECTION.match(name):
            # Heading that carries its block total (ICICI "Equity shares", "Foreign Securities/Overseas ETFs")
            # or an unlabelled subtotal: keep it as a reported total, not an instrument.
            kind = aggregate_rows[i]
            totals.append(dict(section_path=" > ".join(sec_path + [name]) if kind == "header" else " > ".join(sec_path),
                               total_label=f"(aggregate row: {name})", market_value_lakhs=mv, pct_raw=pct_raw, source_row=i + 1))
            if kind == "header":
                sec_path = sec_path[:1] + [name]
            continue
        if mv is None and tonum(pct_raw) is None and not isin and re.search(r'net receivable|net current asset', name, re.I):
            blank_nca = (i, name, list(sec_path))
            continue
        nil = str(g("market_value") or "").strip().upper() == "NIL" or str(pct_raw or "").strip().upper() == "NIL"
        valued_section = bool(VALUED_SECTION.match(name)) and not isin
        is_section = (not isin) and (valued_section or ((mv is None) and tonum(pct_raw) is None) or nil) and bool(SECTION_LIKE.match(name) or (mv is None and tonum(pct_raw) is None))
        if is_section:
            if valued_section and mv is not None:
                totals.append(dict(section_path=name, total_label="(value shown on section header)", market_value_lakhs=mv, pct_raw=pct_raw, source_row=i + 1))
            top_rx = r'^([A-E]\)\s(?!listed|repo)|equity|debt instruments|money market|others?$|derivatives|reit instruments|cash & cash|other current|units of|mutual fund units$|exchange traded funds$|d\) mutual fund units|e\) others|b\) debt|c\) money)'
            if re.match(r'^(equity & equity|equity and equity|a\) equity|debt instruments?$|investment in mutual fund|b\) debt|money market|c\) money|others?$|d\) mutual|e\) others|derivatives|reit instruments|cash & cash|other current|short term deposits|units of|units issued by reits|reits? & invits?|term deposits|deposits \(|equity & equity related instruments|mutual fund units$|real estate investment trusts?$|infrastructure investment trusts?$)', name, re.I):
                sec_path = [name] + (["(nil)"] if nil else [])
                if nil: sec_path = [name]
            elif sec_path and re.match(r'^derivatives', sec_path[0], re.I) and re.match(r'^treps', name, re.I):
                sec_path = [name]  # Edelweiss: TREPS block follows Derivatives with no top-level heading
            else:
                sec_path = sec_path[:1] + [name]
            continue
        if mv is None and tonum(pct_raw) is None and not isin: continue
        if not name and not isin: continue
        ind_rating = norm(g("industry_rating")) if g("industry_rating") is not None else ""
        clean, syms = split_symbols(name)
        clean = re.sub(r'^EQ - ', '', clean)
        ac, st = classify(" > ".join(sec_path), clean, isin, ind_rating)
        if ac is None:
            print(f"  WARNING: dropped unclassifiable row {i + 1}: {name!r} (mv={mv})")
            UNCLASSIFIED_ROWS.append(dict(source_file=f"{year}/{month}/{amc_folder}/{fname}", source_row=i + 1,
                                          section_path=" > ".join(sec_path), row_label=name,
                                          market_value_lakhs=mv, pct_raw=pct_raw))
            continue
        # Tata "^" hedge leg: a futures contract on the stock, not a second holding of the same ISIN
        # (otherwise merge_duplicate_isins nets it into the long position)
        if caret_hedge and "^" in syms and ac == "Equity" and st not in DERIVATIVE_SUB_TYPES:
            st = "Index Futures" if re.search(r'nifty|sensex|\bindex\b', clean, re.I) else "Stock Futures"
            isin = ""
        # the stock's ISIN on an option line (see classify): the leg gets its own synthetic ISIN
        if st in OPTION_SUB_TYPES and re.search(r'\((?:put|call)\s+option\)', clean, re.I):
            isin = ""
        # Kotak "Futures" block: "<Company>-FEB2024" rows carry the underlying stock's ISIN
        if ac == "Equity" and st not in DERIVATIVE_SUB_TYPES and sec_path and re.match(r'^futures$', sec_path[-1], re.I) \
                and re.search(r'-[A-Z]{3}\d{4}$', clean, re.I):
            st = "Index Futures" if re.search(r'nifty|sensex|\bindex\b', clean, re.I) else "Stock Futures"
            isin = ""

# Keep the original ISIN for market-cap/symbol lookup
        lookup_isin = isin
        output_isin = isin
        is_covered_call = "covered call" in clean.lower()

        market_cap_data = (
            get_market_cap_data(lookup_isin, year, month_num)
            if (ac == "Equity" and st not in DERIVATIVE_SUB_TYPES and st not in SPECIAL_ISIN_SUFFIX) or ac == "REITs & InvITs"
            else {
                "symbol": None,
                "market_cap_category": None
            }
        )

# Listed equity printed without an ISIN (e.g. "KWALITY WALLS INDIA LTD" before its
# ISIN was allotted): match the company name against the AMFI list for the period
        if not isin and ac == "Equity" and st == "Listed Equity":
            by_name = find_manual_override(clean, year, month_num) or find_parent_company(None, clean, year, month_num)
            if by_name:
                market_cap_data = by_name
                output_isin = by_name["isin"]

# Partly paid / preference / rights / DVR / covered calls: take ISIN, symbol and
# market cap from the listed parent company and suffix the ISIN to keep it unique
        if st in SPECIAL_ISIN_SUFFIX or is_covered_call:
            suffix = COVERED_CALL_SUFFIX if is_covered_call else SPECIAL_ISIN_SUFFIX[st]
            parent = find_parent_company(isin, clean, year, month_num)
            # partly paid / preference / rights usually carry their own ISIN (IN9..., INE...04..., INE...20...):
            # keep it; the parent ISIN + suffix is only the fallback when the AMC prints no ISIN, a placeholder
            # (PGIM IN25H25DUM01) or the ordinary share's ISIN
            own_isin = (st in OWN_ISIN_SUB_TYPES and not is_covered_call
                        and isin and isin_check_digit_ok(isin) and not (parent and parent["isin"] == isin))
            if own_isin:
                if parent:
                    market_cap_data = parent
            elif parent:
                market_cap_data = parent
                output_isin = f"{parent['isin']}{suffix}"
            else:
                print(f"  WARNING: parent company not found for {st}: {clean!r} (ISIN={isin})")
                if isin:
                    output_isin = f"{isin}{suffix}"
        # physical commodities: one ISIN per commodity (AMC codes like IDIA00500001 differ by AMC)
        if ac == "Commodities" and st == COMMODITY_PHYSICAL:
            output_isin = commodity_name(clean)
        if not isin and ac == "Cash and Cash Equivalents" and not re.search(r'cash|treps|repo|money market|others|current|margin|deposit|cblo|trepS', " ".join(sec_path), re.I):
            sec_path = [name]
        is_rating = bool(RATING_RX.search(ind_rating)) or ac == "Debt"
        qty = tonum(g("quantity"))
        direction = isin_cell if isin_cell in ("Long", "Short") else ("Short" if st in DERIVATIVE_SUB_TYPES and ((qty or 0) < 0 or (mv or 0) < 0) else ("Long" if st in DERIVATIVE_SUB_TYPES else None))
        holdings.append(dict(
            source_file=f"{year}/{month}/{amc_folder}/{fname}", source_sheet=sheet, source_row=i + 1, record_type="HOLDING",
            amc_name=amc_name, scheme_name=scheme, amc_scheme_code=code, portfolio_date=portfolio_date,
            amc_section_l1=sec_path[0] if sec_path else None, amc_section_l2=sec_path[1] if len(sec_path) > 1 else None,
            instrument_name_raw=name,
            instrument_name=clean,
            isin=output_isin or None,
            symbol=market_cap_data["symbol"],
            amc_security_code=(norm(cells.get(0)) if cols.get("serial_no") != 0 and cols.get("instrument_name") != 0 and isinstance(cells.get(0), str) and cells.get(0) not in ("|",) else (str(int(cells[1])) if amc_folder.startswith(("Baroda", "Union")) and tonum(cells.get(1)) else None)),
            asset_class=ac, asset_sub_type=st,
            listing_status=("Unlisted" if re.search(r'unlisted|privately', " ".join(sec_path), re.I) else ("Listed / Awaiting Listing" if ac in ("Equity", "REITs & InvITs", "Debt", "MF and ETF", AIF_ASSET_CLASS) and st not in MONEY_MARKET_SUB_TYPES | DERIVATIVE_SUB_TYPES else None)),
            industry=None if is_rating else (ind_rating or None), credit_rating=ind_rating if is_rating and ind_rating else None,
            market_cap_category=market_cap_data["market_cap_category"],
            nse_industry=market_cap_data.get("industry"),
            quantity=qty, market_value_lakhs=mv, market_value_inr=round(mv * 1e5, 2) if mv is not None else None,
            pct_to_nav_raw=pct_raw, pct_to_nav=None, pct_below_threshold=None,
            coupon_pct=tonum(g("coupon")), maturity_date=str(g("maturity_date").date()) if hasattr(g("maturity_date"), "date") else None,
            yield_raw=tonum(g("yield")), ytc_raw=tonum(g("ytc")), put_call=g("put_call"),
            derivative_position=direction, option_type=option_type_of(str(v) for v in cells.values()) if st in OPTION_SUB_TYPES else None, derivative_pct_hdfc=tonum(g("derivative_pct")), unhedged_pct_hdfc=tonum(g("unhedged_pct")),
            is_placed_as_margin="margin" in (name + " ".join(sec_path)).lower() and ac == "Debt",
            included_in_net_assets=True, footnote_symbols=(syms or "") + (" " + norm(g("notes")) if g("notes") else ""),
        ))
        # DISABLED for now: the per-stock derivative column below books a derivative line outside net assets.
        # # Per-stock derivative column on the equity row, booked as one derivative line named after the stock
        # # (outside net assets; market value is derived from net assets after the loop):
        # #   Canara "Outstanding derivative exposure as % to net assets Long / (Short)" - signed, no futures lines printed
        # #   HDFC "Derivative" / "Derivative % to NAV" - unsigned; "Unhedged" is % to NAV - hedge for a short and
        # #   % to NAV + hedge for a long. Used instead of the dated contracts on the "Derivative..." sheet.
        # dexp = tonum(g("derivative_exposure_pct"))
        # if dexp is None and tonum(g("derivative_pct")):
            # dexp = abs(tonum(g("derivative_pct")))
            # unhedged, held = tonum(g("unhedged_pct")), tonum(pct_raw)
            # # rounding can hide a tiny hedge: treat it as a short, the usual case
            # if not (unhedged is not None and held is not None and unhedged - held >= 0.005):
                # dexp = -dexp
        # if dexp and ac == "Equity" and st not in DERIVATIVE_SUB_TYPES:
            # stock_derivative_rows = True
            # holdings.append(dict(
                # source_file=f"{year}/{month}/{amc_folder}/{fname}", source_sheet=sheet, source_row=i + 1, record_type="DERIVATIVE_DISCLOSURE",
                # amc_name=amc_name, scheme_name=scheme, amc_scheme_code=code, portfolio_date=portfolio_date,
                # amc_section_l1="Outstanding derivative exposure", amc_section_l2=sec_path[1] if len(sec_path) > 1 else None,
                # instrument_name_raw=name, instrument_name=clean, isin=None,
                # asset_class="Equity", asset_sub_type=stock_derivative_type, industry=None if is_rating else (ind_rating or None),
                # pct_to_nav_raw=dexp, derivative_position="Short" if dexp < 0 else "Long",
                # included_in_net_assets=False, is_placed_as_margin=False,
                # footnote_symbols="from the derivative exposure column; market value derived from net assets",
            # ))
    src = f"{year}/{month}/{amc_folder}/{fname}"

    def add_derivative(sh, row, section, nm, sub_type, qty, mv, pct_raw, ls, industry=None, sub=None, **extra):
        # shorts are booked negative (HSBC prints the value unsigned next to a negative quantity)
        ls = ls or ("Short" if (qty or 0) < 0 or (mv or 0) < 0 else "Long")
        if ls == "Short":
            mv = -abs(mv) if mv is not None else None
            pct_raw = -abs(tonum(pct_raw)) if tonum(pct_raw) is not None else pct_raw
        holdings.append(dict(source_file=src, source_sheet=sh, source_row=row + 1, record_type="DERIVATIVE_DISCLOSURE",
            amc_name=amc_name, scheme_name=scheme, amc_scheme_code=code, portfolio_date=portfolio_date,
            amc_section_l1=section, amc_section_l2=sub, instrument_name_raw=nm, instrument_name=nm, isin=None,
            asset_class="Commodities" if sub_type == COMMODITY_DERIVATIVE else "Equity", asset_sub_type=sub_type,
            industry=industry or None, quantity=qty, market_value_lakhs=mv, market_value_inr=round(mv * 1e5, 2) if mv is not None else None,
            pct_to_nav_raw=pct_raw, derivative_position=ls, included_in_net_assets=False, is_placed_as_margin=False, **extra))

    def has_derivatives():
        return any(h.get("asset_sub_type") in DERIVATIVE_SUB_TYPES for h in holdings)

    def futures_sub_type(nm, hint=""):
        if re.search(r'futcur|usdinr|eurinr|gbpinr|jpyinr', nm, re.I): return "Currency Futures"
        return "Index Futures" if re.search(r'nifty|index|sensex', nm + " " + hint, re.I) else "Stock Futures"

    def long_short(cells):
        return next((norm(v).title() for v in cells.values() if isinstance(v, str) and norm(v).lower() in ("long", "short")), None)

    def scan_derivative_blocks(d, start, sh):
        """Derivative tables printed below the grand total (not part of net assets)."""
        for j in range(start, len(d)):
            t = " ".join(str(v) for v in d.iloc[j] if pd.notna(v)).lower()
            # PPFAS prints a plain "Derivatives" table (stock + currency futures) below the grand total
            # Nippon Multi Asset: "Details of Commodity Future / Index Future" (FUTCOM_SILVER_03/05/2024)
            # UTI (Jan/Feb 2024): a repeated column header, then a "FUTURES" block (stock futures with the underlying's ISIN)
            # HSBC (Jan/Feb 2024): "Disclosure in Derivatives | Quantity | Market Value | % To net assets" on the NOTES sheet
            if re.search(r'disclosure in derivatives|details of (stock|commodity) future|^derivatives$|^futures$', t.strip()):
                prev = " ".join(str(v) for v in d.iloc[j - 1] if pd.notna(v)).lower()
                # Prism leaves a blank line between the "Derivatives" heading and its table
                first_row = next((x for x in range(j + 1, min(j + 4, len(d))) if d.iloc[x].notna().any()), j + 1)
                dh = map_header(d.iloc[j] if "quantity" in t else d.iloc[j - 1] if t.strip() == "futures" and "quantity" in prev else d.iloc[first_row])
                if "instrument_name" not in dh or "disclosure" in t: dh["instrument_name"] = d.iloc[j].first_valid_index()
                k, sub = first_row, None
                while k < len(d):
                    r = d.iloc[k]; cells = {c: v for c, v in r.items() if pd.notna(v) and str(v).strip() != ""}
                    if not cells: break
                    nm = norm(cells.get(dh.get("instrument_name"), "")); mv = tonum(cells.get(dh.get("market_value")))
                    # RedHex (HSBC) "Notes": the SEBI summary ("For the period ended ...", contract counts) follows the list
                    if re.match(r'^for the (period|half)', nm, re.I) or tonum(nm) is not None: break
                    # ICICI sub-headings inside the block: "Exchange Traded Commodity Derivatives" > "A) LISTED ON COMMODITY EXCHANGES ..."
                    if mv is None and nm and len(cells) == 1:
                        sub = nm
                    if mv is not None and nm and not TOTAL_RX.search(nm):
                        is_commodity = bool(re.search(r'commodit', sub or "", re.I) or COMMODITY_DERIVATIVE_NAME.search(nm))
                        add_derivative(sh, k, norm(d.iloc[j].dropna().iloc[0]), nm,
                                       COMMODITY_DERIVATIVE if is_commodity else futures_sub_type(nm, str(cells.get(dh.get("industry_rating")))),
                                       tonum(cells.get(dh.get("quantity"))), mv, cells.get(dh.get("pct_nav")), long_short(cells),
                                       industry=norm(cells.get(dh.get("industry_rating"), "")), sub=sub)
                    k += 1

    def scan_hedging_tables(d, start, sh):
        """
        SEBI derivative disclosure: "A. Hedging Positions through Futures as on ..." / "B. Other than Hedging
        Positions through Futures ...", header "Underlying | Long / Short | ... | Margin | [Quantity | Market Value
        | % to Net Assets]". Only read when the portfolio itself lists no derivatives (Mirae), since most AMCs
        repeat their inline futures here.
        Mirae Feb 2024 prints only the margin per stock: the stated "Total exposure due to futures ... as a %age
        of net assets : -16.12 %" is then split across the stocks pro rata to margin (flagged as estimated).
        """
        j = start
        while j < len(d):
            t = " ".join(str(v) for v in d.iloc[j] if pd.notna(v)).lower().strip()
            # Apex: "(a) Hedging Position through Futures", "(d) Other than Hedging Positions through Options"
            if not re.match(HEDGING_TABLE_RX, t):
                j += 1; continue
            section = norm(d.iloc[j].dropna().iloc[0])
            hi = None
            for h in range(j + 1, min(j + 4, len(d))):
                row_texts_h = [v.strip().lower() for v in d.iloc[h] if isinstance(v, str)]
                # an empty section sits right above the next one: that header is not this section's
                if row_texts_h and re.match(HEDGING_TABLE_RX, row_texts_h[0]):
                    break
                if any(v.startswith("underlying") for v in row_texts_h):
                    hi = h
                    break
            if hi is None:
                j += 1; continue
            col = {}
            for c, v in d.iloc[hi].items():
                h = norm(v).lower() if isinstance(v, str) else ""
                for fld, rx in (("name", r'^underlying'), ("quantity", r'quantity'), ("market_value", r'market|fair value'),
                                ("pct_nav", r'% ?to net|net assets'), ("margin", r'margin'), ("current_price", r'current price')):
                    if fld not in col and re.search(rx, h): col[fld] = c; break
            rows, k = [], hi + 1
            while k < len(d):
                cells = {c: v for c, v in d.iloc[k].items() if pd.notna(v) and str(v).strip() != ""}
                nm = norm(cells.get(col.get("name"), ""))
                if not cells or re.match(r'^(total|for the|nil$|note)', nm, re.I) or re.match(r'^\(?[a-e][.)]\s', nm, re.I):
                    break
                if nm:
                    rows.append((k, nm, cells))
                k += 1
            is_option_table = "option" in t
            hedge_position = "Other" if t.find("other than") >= 0 else "Hedging"

            def table_sub_type(nm):
                if is_option_table:
                    return "Index Options" if re.search(r'nifty|sensex|\bindex\b', nm, re.I) else "Stock Options"
                return futures_sub_type(nm)

            if "market_value" in col or "pct_nav" in col or is_option_table or (re.match(r'^\(', t) and "margin" in col):
                # Apex Apr 2026 prints no % per leg, only "Total percentage of existing assets ... is 1.65%":
                # a one-leg table takes it; legs of larger tables are booked with no value
                stated = None
                if "market_value" not in col and "pct_nav" not in col and len(rows) == 1:
                    tail = " ".join(" ".join(str(v) for v in d.iloc[x] if pd.notna(v)) for x in range(k, min(k + 2, len(d))))
                    m = re.search(r'total percentage of existing assets.*?is\s+(-?[\d.]+)\s*%', tail, re.I)
                    if m:
                        stated = float(m.group(1)) / (100 if pct_is_fraction else 1)
                for r_i, nm, cells in rows:
                    mv = tonum(cells.get(col.get("market_value")))
                    pct = stated if stated is not None else cells.get(col.get("pct_nav"))
                    if mv is None and tonum(pct) is None and ("pct_nav" in col or "market_value" in col): continue
                    add_derivative(sh, r_i, section, nm, table_sub_type(nm), tonum(cells.get(col.get("quantity"))), mv,
                                   pct, long_short(cells), amc_hedge_position=hedge_position,
                                   margin_lakhs=tonum(cells.get(col.get("margin"))), current_price=tonum(cells.get(col.get("current_price"))))
            elif "margin" in col and rows:
                tail = " ".join(" ".join(str(v) for v in d.iloc[x] if pd.notna(v)) for x in range(k, min(k + 8, len(d))))
                m = re.search(r'total exposure due to futures.*?net assets\s*:?\s*(-?[\d.]+)\s*%', tail, re.I)
                margins = [tonum(cells.get(col["margin"])) or 0 for _, _, cells in rows]
                if m and sum(margins) > 0:
                    total = abs(float(m.group(1)))
                    for (r_i, nm, cells), mg in zip(rows, margins):
                        # stated total is in percent; store in the main sheet's % scale
                        pct = total * mg / sum(margins) / (100 if pct_is_fraction else 1)
                        add_derivative(sh, r_i, section, nm, futures_sub_type(nm), None, None, round(pct, 8), long_short(cells),
                                       margin_lakhs=mg, current_price=tonum(cells.get(col.get("current_price"))),
                                       footnote_symbols="estimated: stated total futures exposure split pro rata to margin (no quantity / value in source)")
            j = max(k, j + 1)

    # DISABLED for now: derivative disclosure tables below the grand total, on other sheets and on the
    # "Derivative" sheet are outside net assets and push holdings away from 100. Re-enable by uncommenting.
    # scan_derivative_blocks(df, last_i + 1, sheet)
    # # derivative sheets are read below; other extra sheets may hold the disclosure tables (HSBC / RedHex "Notes").
    # # Other schemes' portfolio sheets, index / performance / disclaimer sheets are not this scheme's tables.
    # extra = {sh: xf.parse(sh, header=None) for sh in xf.sheet_names
             # if sh not in portfolio_sheets and (re.match(r'^notes?\b', sh.strip(), re.I) or not NON_PORTFOLIO_SHEET.search(sh.strip()))}
    # for sh, d in extra.items():
        # scan_derivative_blocks(d, 0, sh)
    # main_pcts = [tonum(h["pct_to_nav_raw"]) for h in holdings if h["record_type"] == "HOLDING" and tonum(h["pct_to_nav_raw"]) is not None]
    # pct_is_fraction = (grand[1] is not None and abs(grand[1] - 1) < 0.02) or (grand[1] is None and sum(main_pcts) < 5)
    # if not has_derivatives():
        # scan_hedging_tables(df, last_i + 1, sheet)
        # for sh, d in extra.items():
            # if not has_derivatives():
                # scan_hedging_tables(d, 0, sh)
    # if deriv_sheet:
        # # HDFC "Derivative<scheme>" sheet, one or more tables. Columns are mapped from each "Underlying" header:
        # # older files lead with a "Scheme Name" column, newer ones start at "Underlying | Industry | ...".
        # # Futures tables carry a market value; options tables (HDFC Flexi Cap covered calls, Jan/Feb 2024) only a
        # # quantity and the current option price, so their value is quantity x price, booked only when the main
        # # sheet lists no derivatives. ICICI's "Derivative" sheet has prices and margin but no market value: its
        # # contracts are already in the main sheet.
        # sh = deriv_sheet
        # d = xf.parse(sh, header=None)
        # had_derivatives = has_derivatives()
        # col, section = None, None
        # for k in range(len(d)):
            # r = d.iloc[k]
            # texts = [norm(v).lower() for v in r.values if isinstance(v, str)]
            # first = texts[0] if texts else ""
            # if any(t.startswith("underlying") for t in texts):
                # col = {}
                # for c, v in r.items():
                    # h = norm(v).lower() if isinstance(v, str) else ""
                    # for fld, rx in (("name", r'^underlying'), ("industry", r'^industry'), ("quantity", r'long\s*/\s*\(?short'),
                                    # ("entry_price", r'price when purchased'), ("current_price", r'^current'),
                                    # ("margin", r'margin'), ("market_value", r'market value')):
                        # if fld not in col and re.search(rx, h): col[fld] = c; break
                # is_option = any("option price" in t for t in texts)
                # if "name" not in col or not ("market_value" in col or (is_option and "current_price" in col and not had_derivatives)):
                    # col = None
                # continue
            # if re.match(r'^([a-e]\.\s|total|for the|scheme name)', first):
                # col = None
                # if re.match(r'^[a-e]\.\s', first): section = norm(next(v for v in r.values if isinstance(v, str)))
                # continue
            # if col is None or r.dropna().empty: continue
            # nm = norm(r[col["name"]]) if pd.notna(r[col["name"]]) else ""
            # qty = tonum(r[col["quantity"]]) if "quantity" in col else None
            # if not nm or qty is None: continue
            # if "market_value" in col:
                # mvv = tonum(r[col["market_value"]])
                # sub_type = "Index Futures" if "nifty" in nm.lower() else "Stock Futures"
            # else:
                # price = tonum(r[col["current_price"]])
                # mvv = round(qty * price / 1e5, 6) if price is not None else None
                # sub_type = "Index Options" if re.search(r'nifty|sensex', nm, re.I) else "Stock Options"
            # # stock contracts are already booked from the equity table's Derivative column; keep index contracts
            # if mvv is None or (stock_derivative_rows and sub_type.startswith("Stock")): continue
            # if stock_derivative_rows:
                # nm = re.sub(r'\s*\d{2}-\d{2}-\d{4}.*$', '', nm) or nm  # "Nifty29-02-2024" -> "Nifty", like the stock lines
            # add_derivative(sh, k, section or "Hedging Positions through Futures", nm, sub_type, qty, mvv, None, None,
                           # industry=norm(r[col["industry"]]) if "industry" in col and pd.notna(r[col["industry"]]) else None,
                           # entry_price=tonum(r[col["entry_price"]]) if "entry_price" in col else None,
                           # current_price=tonum(r[col["current_price"]]) if "current_price" in col else None,
                           # margin_lakhs=tonum(r[col["margin"]]) if "margin" in col else None)
    H = pd.DataFrame(holdings)
    # --- value unit: Rs. lakhs is the norm; Titanium (Tata) JAN / MAY 2026 print rupees ("MKT VAL(Rs.)" or a bare
    # "MKT VAL" with net assets of ~2e9). No SIF comes near 1e7 lakhs (Rs 1 lakh crore) of net assets.
    unit = norm(df.iloc[hdr][cols["market_value"]]).lower() if "market_value" in cols else ""
    to_lakhs = (1e-5 if re.search(r'\brs\.?\)', unit) and not re.search(r'lakh|lac|crore|cr\b', unit)
                else 100 if re.search(r'crore|\bcr\b', unit)
                else 1e-5 if grand[0] is not None and abs(grand[0]) > 1e7 and not re.search(r'lakh|lac', unit)
                else 1)
    if to_lakhs != 1:
        print(f"  -> Market values in {'rupees' if to_lakhs < 1 else 'crores'} ({unit!r}): converted to lakhs")
        for c in ("market_value_lakhs", "margin_lakhs"):
            if c in H:
                H[c] = H[c] * to_lakhs
        H["market_value_inr"] = (H["market_value_lakhs"] * 1e5).round(2)
        grand = (grand[0] * to_lakhs if grand[0] is not None else None,) + grand[1:]
        for t in totals:
            if t["market_value_lakhs"] is not None:
                t["market_value_lakhs"] *= to_lakhs
    # derivative tables keep the name as printed (iSIF "HDFC Bank Ltd $$"): every final name loses its footnote marks
    H["instrument_name"] = H["instrument_name"].map(lambda n: standardize_ltd(split_symbols(n)[0] if isinstance(n, str) else n))
    deriv =H.asset_sub_type.isin(DERIVATIVE_SUB_TYPES) & H["isin"].isna()
    if deriv.any():
        # the same name under "Hedging" and "Other than Hedging" (Apex, Platinum) is two positions, not one to merge
        other = (H["amc_hedge_position"] == "Other") if "amc_hedge_position" in H else pd.Series(False, index=H.index)
        H.loc[deriv, "isin"] = [derivative_isin(st, f"{nm} (other than hedging)" if o else nm)
                                for st, nm, o in zip(H.loc[deriv, "asset_sub_type"], H.loc[deriv, "instrument_name"], other[deriv])]
    main = H[H.record_type == "HOLDING"]
    # --- percent scale: fraction (0-1) vs percent (0-100)
    pnum = main.pct_to_nav_raw.map(tonum)
    gp = grand[1]
    frac = (gp is not None and abs(gp - 1) < 0.02) or (gp is None and pnum.sum(skipna=True) < 5)
    scale = 100 if frac else 1
    def pct_std(v):
        n = tonum(v)
        return round(n * scale, 6) if n is not None else None
    H["pct_to_nav"] = H.pct_to_nav_raw.map(pct_std)
    # "<0.01% of NAV" markers: 0.00%, # / * (various AMCs), @ (HDFC), ^ (ICICI)
    H["pct_below_threshold"] = H.pct_to_nav_raw.map(lambda v: "<0.01%" if isinstance(v, str) and re.search(r'^\$?0\.00%?$|^[#*@^]$', v.strip()) else None)
    # store them as 0.01 (-0.01 for short positions) so totals never come out null
    below = H.pct_below_threshold.notna()
    H.loc[below, "pct_to_nav"] = [(-0.01 if (mv or 0) < 0 else 0.01) for mv in H.loc[below, "market_value_lakhs"]]
    # % to NAV left blank by the AMC: derive from market value / net assets. Covers the Edelweiss Multi Asset
    # gold line and HDFC's "Derivative..." sheet, which has no % column (the main sheet's "Derivative"
    # column is this same ratio, rounded, per underlying stock)
    if grand[0]:
        blank = H.pct_to_nav.isna() & H.market_value_lakhs.notna()
        H.loc[blank, "pct_to_nav"] = (H.loc[blank, "market_value_lakhs"] / grand[0] * 100).round(6)
        # and the reverse for derivative lines printed with a % only (Canara exposure column, Mirae Feb 2024 estimate)
        no_mv = H.market_value_lakhs.isna() & H.pct_to_nav.notna() & (H.record_type != "HOLDING")
        H.loc[no_mv, "market_value_lakhs"] = (H.loc[no_mv, "pct_to_nav"] * grand[0] / 100).round(6)
        H.loc[no_mv, "market_value_inr"] = (H.loc[no_mv, "market_value_lakhs"] * 1e5).round(2)
    ycols = [c for c in ("yield_raw", "ytc_raw") if c in H]
    for c in ycols:
        H[c.replace("_raw", "_pct")] = H[c].map(lambda v: None if v is None or pd.isna(v) else (round(v * 100, 4) if abs(v) < 1 else v))
    # --- futures that the AMC excludes from its grand total (Baroda BNP, DSP)
    mvsum = main.market_value_lakhs.sum(skipna=True)
    if grand[0] is not None:
        futs = main[is_futures(main)]
        short_futs = futs[futs.market_value_lakhs < 0]
        if len(futs) and abs(mvsum - grand[0]) > 1 and abs((mvsum - futs.market_value_lakhs.sum()) - grand[0]) < 1:
            H.loc[futs.index, "included_in_net_assets"] = False
        # Tata Multi Asset: long commodity futures are in the totals, the "^" short hedge legs (stock and commodity) are not
        elif len(short_futs) and abs(mvsum - grand[0]) > 1 and abs((mvsum - short_futs.market_value_lakhs.sum()) - grand[0]) < 1:
            H.loc[short_futs.index, "included_in_net_assets"] = False
        # Tata Balanced Advantage May 2026: futures and a Nifty put printed inside the equity block, all left out of the totals
        else:
            derivs = main[main.asset_sub_type.isin(DERIVATIVE_SUB_TYPES)]
            if len(derivs) and abs(mvsum - grand[0]) > 1 and abs((mvsum - derivs.market_value_lakhs.sum()) - grand[0]) < 1:
                H.loc[derivs.index, "included_in_net_assets"] = False
        # Short equity lines the AMC lists but leaves out of its totals
        # (Tata Large Cap Jan 2025: "INDUSIND BANK LTD^" qty -550000, value -5489)
        shorts = main[(main.asset_class == "Equity") & ~main.asset_sub_type.isin(DERIVATIVE_SUB_TYPES) & (main.market_value_lakhs < 0)]
        if len(shorts) and abs(mvsum - grand[0]) > 1 and abs((mvsum - shorts.market_value_lakhs.sum()) - grand[0]) < 1:
            H.loc[shorts.index, "included_in_net_assets"] = False
    # AMC printed the Net Receivables label but left the value blank: book the balancing figure
    # (grand total - instruments), excluding futures when that is what makes the residual smaller
    if blank_nca is not None and grand[0] is not None:
        inc_now = H[(H.record_type == "HOLDING") & (H.included_in_net_assets)]
        res_a = grand[0] - inc_now.market_value_lakhs.sum(skipna=True)
        fut_idx = inc_now[is_futures(inc_now)].index
        res_b = res_a + H.loc[fut_idx, "market_value_lakhs"].sum(skipna=True)
        if len(fut_idx) and abs(res_b) < abs(res_a):
            H.loc[fut_idx, "included_in_net_assets"] = False
            res_a = res_b
        if abs(res_a) >= 0.01:
            r0, nm0, sp0 = blank_nca
            nca = {c: None for c in H.columns}
            nca.update(source_file=f"{year}/{month}/{amc_folder}/{fname}", source_sheet=sheet, source_row=r0 + 1, record_type="HOLDING",
                       amc_name=amc_name, scheme_name=scheme, amc_scheme_code=code, portfolio_date=portfolio_date,
                       amc_section_l1=sp0[0] if sp0 else None, amc_section_l2=sp0[1] if len(sp0) > 1 else None,
                       instrument_name_raw=nm0, instrument_name=nm0, asset_class="Cash and Cash Equivalents", asset_sub_type="Net Receivables / (Payables)",
                       market_value_lakhs=round(res_a, 6), market_value_inr=round(res_a * 1e5, 2), pct_to_nav=round(res_a / grand[0] * 100, 6),
                       included_in_net_assets=True, is_placed_as_margin=False,
                       footnote_symbols="value blank in source; derived as grand total minus listed instruments")
            H = pd.concat([H, pd.DataFrame([nca])], ignore_index=True)
    # DISABLED derivatives: keep only what the AMC counts in its grand total, so holdings sum to 100
    # (in-table futures it leaves out, flagged above, are dropped here)
    H = H[H.included_in_net_assets.astype(bool)].reset_index(drop=True)
    H = merge_duplicate_isins(H)
    inc = H[(H.record_type == "HOLDING") & (H.included_in_net_assets)]
    snap = dict(source_file=f"{year}/{month}/{amc_folder}/{fname}", source_sheet=sheet, amc_name=amc_name, scheme_name=scheme, amc_scheme_code=code,
                portfolio_date=portfolio_date, reported_date_text=reported_date, header_row=hdr + 1, pct_scale_in_source="fraction (0-1)" if frac else "percent (0-100)",
                net_assets_lakhs=grand[0], grand_total_label=grand[2], grand_total_row=grand[3], holdings_count=len(inc),
                sum_market_value_lakhs=round(inc.market_value_lakhs.sum(skipna=True), 2), sum_pct_to_nav=round(inc.pct_to_nav.sum(skipna=True), 4),
                derivative_disclosure_rows=int((H.record_type != "HOLDING").sum() + (~H.included_in_net_assets & (H.record_type == "HOLDING")).sum()))
    snap["mv_reconciliation_diff_lakhs"] = round(snap["sum_market_value_lakhs"] - grand[0], 4) if grand[0] is not None else None
    snap["recon_status"] = "OK" if snap["mv_reconciliation_diff_lakhs"] is not None and abs(snap["mv_reconciliation_diff_lakhs"]) < 1 else ("NO MARKET VALUE IN SOURCE" if grand[0] is None else "CHECK")
    # File filed under the wrong month folder (Apex MAR 2026 folder holds the "as on May 31, 2026" statement)
    reported_month = dval.month if dval is not None and pd.notna(dval) else next(
        (n for n, name in MONTH_NAME.items() if re.search(r'\b' + name[:3], reported_date, re.I)), None)
    snap["date_check"] = "OK" if reported_month in (None, month_num) else f"MISMATCH: statement is for {MONTH_NAME[reported_month].title()}"
    if snap["date_check"] != "OK":
        print(f"  WARNING: {snap['date_check']} but filed under {month} ({reported_date!r})")
    T = pd.DataFrame(totals); T["source_file"] = snap["source_file"]
    return H, snap, T

MONTH_NAME = {
    1: "january",
    2: "february",
    3: "march",
    4: "april",
    5: "may",
    6: "june",
    7: "july",
    8: "august",
    9: "september",
    10: "october",
    11: "november",
    12: "december",
}

def get_fund_category(scheme_name, source_file):
    """
    Derive the SEBI SIF investment-strategy category from the scheme.

    Examples:
        iSIF Equity Ex-Top 100 Long-Short Fund      -> Equity Ex-Top 100 Long-Short
        Altiva Hybrid Long-Short Fund               -> Hybrid Long-Short
        Dynasif Active Asset Allocator Long-Short   -> Active Asset Allocator Long-Short
        Arudha Equity Long-Short Fund               -> Equity Long-Short
    """

    text = f"{scheme_name or ''} {source_file or ''}".lower()
    # file names use _ / - / – as separators: "Arudha-Equity-Long-Short-Fund-31-March-2026"
    text = re.sub(r"[_\-–]+", " ", text)

    # most specific first: "Equity Ex-Top 100" before "Equity", "Sectoral Debt" before "Debt"
    category_patterns = [
        (r"\bex\s*top\s*100\b", "Equity Ex-Top 100 Long-Short"),
        (r"\bsector\s+rotation\b", "Sector Rotation Long-Short"),
        (r"\bsectoral\s+debt\b", "Sectoral Debt Long-Short"),
        (r"\bdebt\s+long\s*short\b", "Debt Long-Short"),
        (r"\bactive\s+asset\s+allocator\b", "Active Asset Allocator Long-Short"),
        (r"\bhybrid\b", "Hybrid Long-Short"),
        (r"\bequity\b", "Equity Long-Short"),
    ]

    for pattern, category in category_patterns:
        if re.search(pattern, text):
            return category

    return "Others"


def clean_json_value(value):
    """
    Convert pandas/numpy values into JSON-safe Python values.
    """

    if pd.isna(value):
        return None

    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass

    return value


# Asset classes reported as a single total row per fund-month in the JSON,
# keyed by a synthetic ISIN (these rows mostly have no real ISIN).
# Debt rows all carry a real ISIN, so they are listed one per instrument like equity
AGGREGATE_ASSET_CLASSES = {
    "Cash and Cash Equivalents": "CASHANDEQ",
}


def collapse_aggregate_holdings(holdings):
    """
    Replace every Cash and Cash Equivalents row in one fund-month
    with a single row holding the summed % to NAV.
    """

    kept = []
    totals = {}
    seen_isin = {}

    for holding in holdings:
        asset_class = holding.get("asset_class")

        if asset_class not in AGGREGATE_ASSET_CLASSES:
            # Neo4j has a unique-ISIN constraint per snapshot: fold any repeat
            # ISIN in the fund-month (e.g. the same scheme filed twice) into one row
            isin = holding.get("isin")
            if isin and isin in seen_isin:
                prev = seen_isin[isin]
                if holding.get("holdings") is not None:
                    prev["holdings"] = round((prev.get("holdings") or 0) + holding["holdings"], 6)
                continue
            if isin:
                seen_isin[isin] = holding
            kept.append(holding)
            continue

        pct = holding.get("holdings")
        total = totals.setdefault(asset_class, None)

        if pct is not None:
            totals[asset_class] = (total or 0) + pct

    for asset_class, total in totals.items():
        kept.append({
            "instrument": asset_class,
            "isin": AGGREGATE_ASSET_CLASSES[asset_class],
            "asset_class": asset_class,
            "industry": "Others",
            "sector": "Others",
            "holdings": round(total, 6) if total is not None else None,
            "subType": None,
            "symbol": None,
            "market_cap": None,
            "included_in_net_assets": True,
        })

    return kept


def apply_nse_industry_fallback(H):
    """
    After the AMC industry -> SEBI sector mapping: rows whose AMC label gave no
    sector (blank, Miscellaneous, Others, Not Applicable, multi-sector buckets like
    "Units of Infrastructure Investment Trusts") take the industry published on
    nseindia.com, carried in the market-cap list ("NSE Industry"; equity, REITs and
    InvITs only). industry_source records where the industry came from: AMC | NSE.
    """
    if "nse_industry" not in H.columns:
        H["nse_industry"] = None
    amc = H["industry"].fillna("").astype(str).str.strip()
    nse = H["nse_industry"].fillna("").astype(str).str.strip()
    use_nse = H["sector"].isna() & nse.ne("")
    H["industry_source"] = amc.ne("").map({True: "AMC", False: None})
    H.loc[use_nse, "industry"] = nse[use_nse]
    H.loc[use_nse, "industry_source"] = "NSE"
    if use_nse.any():
        H = add_sector_columns(H)
    print(f"  -> Industry from NSE (AMC label gave no sector): {int(use_nse.sum())} rows")
    return H


# contract words / expiry / strike in a derivative leg's name:
# "CALL Coal India Ltd 26/05/2026 450", "Maruti Suzuki India Ltd June 2026 Future", "HDFC Bank Ltd.(Put Option) $$"
DERIVATIVE_NAME_NOISE_RX = re.compile(
    r"\b(call|put|futures?|options?|fut|opt|ce|pe|jan(uary)?|feb(ruary)?|mar(ch)?|apr(il)?|may|june?|july?|aug(ust)?|"
    r"sep(t|tember)?|oct(ober)?|nov(ember)?|dec(ember)?)\b|\d+(\.\d+)?", re.I)


# derivative legs named by an old / short / misspelt company name (key -> key of the listed name)
UNDERLYING_ALIASES = {
    "zomato": "eternal",                                            # Arudha "Zomato Ltd 26/05/2026"
    "dmart": "avenue supermarts",                                   # Arudha "Dmart 29/09/2026"
    "cromptongreaves consumer electricals": "crompton greaves consumer electricals",  # Platinum
    "adani ports sez": "adani ports special economic zone",         # Arudha "Adani Ports & SEZ Ltd"
    "kotak bank": "kotak mahindra bank",                            # Arudha "Kotak Bank Ltd"
    "life insurance corpora": "life insurance of",                  # Prism "Life Insurance Corpora Fut"
}


def underlying_key(name):
    key = company_key(DERIVATIVE_NAME_NOISE_RX.sub(" ", str(name or "")))
    return UNDERLYING_ALIASES.get(key, key)


def apply_underlying_industry(H):
    """
    Stock futures / options legs (the short book of a long-short SIF) that the AMC prints
    without an industry (Altiva, Arudha, Platinum, RedHex, Dyna) take the industry of the
    underlying stock: first a Listed Equity row with the same name in the same file, then
    in any file, then the AMFI list's NSE industry. industry_source = Underlying.
    """
    stock = H["asset_sub_type"].isin(["Stock Futures", "Stock Options"]) & H["sector"].isna()
    if not stock.any():
        return H
    equity = H[(H["asset_sub_type"] == "Listed Equity") & H["sector"].notna()]
    keys = equity["instrument_name"].map(underlying_key)
    by_file = {(f, k): ind for f, k, ind in zip(equity["source_file"], keys, equity["industry"]) if k}
    by_name = {k: ind for k, ind in zip(keys, equity["industry"]) if k}
    by_symbol = {s: ind for s, ind in zip(equity["symbol"], equity["industry"]) if isinstance(s, str) and s}
    for i in H.index[stock]:
        key = underlying_key(H.at[i, "instrument_name"])
        industry = by_file.get((H.at[i, "source_file"], key)) or by_name.get(key)
        # NSE contract codes: "OPTSTKAXISBANK30-Mar-2026CE", "FUTSTKBAJAJ-AUTO30-Mar-2026"
        contract = re.match(r"^(?:OPT|FUT)STK(.+?)\d{2}-[A-Za-z]{3}-\d{4}", str(H.at[i, "instrument_name"]))
        if industry is None and contract:
            industry = by_symbol.get(contract.group(1).upper())
        if industry is None and key:
            year, month_num = int(str(H.at[i, "portfolio_date"])[:4]), int(str(H.at[i, "portfolio_date"])[5:7])
            parent = find_parent_company(None, key, year, month_num)
            industry = parent.get("industry") if parent else None
        if industry:
            H.at[i, "industry"] = industry
            H.at[i, "industry_source"] = "Underlying"
    H = add_sector_columns(H)
    print(f"  -> Industry from the underlying stock (derivative legs): {int(H['industry_source'].eq('Underlying').sum())} rows")
    return H


# JSON subType of a derivative leg
DERIVATIVE_HEDGING = "derivative-hedging"
DERIVATIVE_DIRECTIONAL = "derivative-directional-position"
DERIVATIVE_MIXED = "derivatives"
OPTION_SUB_TYPES = ("Stock Options", "Index Options")
# hedge bigger than the stock held by no more than this still counts as covered (lot sizes round up)
HEDGE_SIZE_TOLERANCE = 1.05


def option_type_of(texts):
    """Call / Put from a row's cells: "(Put Option)", Titanium "Call Option", "OPTSTKBDL30-Mar-2026CE", Tata "...OPT P"."""
    for t in texts:
        if re.search(r'\bcall\b|\d\s*ce\d*$|\bopt\s*c$', t, re.I):
            return "Call"
        if re.search(r'\bput\b|\d\s*pe\d*$|\bopt\s*p$', t, re.I):
            return "Put"
    return None


def match_held_key(key, held_keys):
    """
    A held stock's key for a shortened or re-spaced derivative name, when exactly one fits:
    Prism "Life Insurance Corpora Fut", "Tata Motors Pass Vehicles Fut", "Ultra Tech Cement Ltd Fut".
    """
    words = key.split()
    compact = key.replace(" ", "")
    hits = []
    for held in held_keys:
        held_words = held.split()
        if held.replace(" ", "") == compact or (
            1 < len(words) <= len(held_words) and all(h.startswith(w) for w, h in zip(words, held_words))
        ):
            hits.append(held)
    return hits[0] if len(hits) == 1 else None


def apply_derivative_type(H):
    """
    subType of every derivative leg, per fund-month:
      derivative-hedging              offsets a stock the fund holds, at no more than the size held
      derivatives                     offsets a stock the fund holds, but is larger: part hedged, part naked
      derivative-directional-position no matching stock (naked), a leg adding to the stock (long future,
                                      long call, short put), or an index / commodity / currency contract
    Interest rate swaps keep their own subType. Legs printed with no quantity (Apex) cannot be sized against
    the holding, so they take the AMC's own table heading: "Hedging Position" -> derivative-hedging,
    "Other than Hedging" -> derivative-directional-position.

    Offsetting legs against a long stock: short future, short call, long put. Altiva prints most options
    with no call / put ("SHARE OPTIONS"): a short one is taken as a covered call. Size is compared in shares
    (an option's value is only its premium), else in value for futures. A collar (long put + short call)
    covers the same shares once, so hedge size = short futures + max(short calls, long puts).
    """
    H["derivative_type"] = None
    legs = H["asset_sub_type"].isin(DERIVATIVE_SUB_TYPES) & (H["asset_sub_type"] != IRS_SUB_TYPE)
    H.loc[legs, "derivative_type"] = DERIVATIVE_DIRECTIONAL

    # column only exists when some file has such a table
    hedge_position = H.get("amc_hedge_position", pd.Series(index=H.index, dtype=object))
    by_heading = legs & H["quantity"].isna() & hedge_position.notna()
    H.loc[by_heading & hedge_position.eq("Hedging"), "derivative_type"] = DERIVATIVE_HEDGING

    for source_file, rows in H.groupby("source_file", sort=False):
        equity = rows[(rows["asset_sub_type"] == "Listed Equity") & (rows["quantity"].fillna(1) > 0)]
        held = {r.isin: (r.quantity, r.market_value_lakhs) for r in equity.itertuples() if isinstance(r.isin, str)}
        by_key = {underlying_key(r.instrument_name): r.isin for r in equity.itertuples() if isinstance(r.isin, str)}
        by_symbol = {r.symbol: r.isin for r in equity.itertuples() if isinstance(r.symbol, str) and isinstance(r.isin, str)}
        year, month_num = int(str(rows["portfolio_date"].iloc[0])[:4]), int(str(rows["portfolio_date"].iloc[0])[5:7])

        offsetting = {}
        for r in rows[rows["asset_sub_type"].isin(["Stock Futures", "Stock Options"])].itertuples():
            if by_heading.get(r.Index, False):
                continue
            key = underlying_key(r.instrument_name)
            contract = re.match(r"^(?:OPT|FUT)STK(.+?)\d{2}-[A-Za-z]{3}-\d{4}", str(r.instrument_name))
            isin = by_key.get(key) or (by_symbol.get(contract.group(1).upper()) if contract else None)
            if isin is None and key:
                isin = by_key.get(match_held_key(key, by_key))
            if isin is None and key:
                parent = find_parent_company(None, key, year, month_num)
                isin = parent["isin"] if parent and parent["isin"] in held else None
            if isin is None:
                continue
            # the quantity's sign is the side: Titanium prints a long put with a negative value
            qty = r.quantity if pd.notna(r.quantity) and r.quantity != 0 else None
            short = qty < 0 if qty is not None else (r.derivative_position == "Short" or (r.market_value_lakhs or 0) < 0)
            otype = r.option_type if isinstance(r.option_type, str) else option_type_of([str(r.instrument_name)])
            if r.asset_sub_type == "Stock Futures":
                bucket = "future" if short else None
            elif short:
                bucket = None if otype == "Put" else "call"
            else:
                bucket = "put" if otype == "Put" else None
            if bucket:
                offsetting.setdefault(isin, []).append((r.Index, bucket, qty, r.market_value_lakhs))

        for isin, offs in offsetting.items():
            held_qty, held_mv = held[isin]
            if pd.notna(held_qty) and all(o[2] is not None for o in offs):
                size = {b: sum(abs(o[2]) for o in offs if o[1] == b) for b in ("future", "call", "put")}
                hedge, base = size["future"] + max(size["call"], size["put"]), held_qty
            else:
                hedge, base = sum(abs(o[3] or 0) for o in offs if o[1] == "future"), held_mv
            label = DERIVATIVE_HEDGING if hedge <= (base or 0) * HEDGE_SIZE_TOLERANCE else DERIVATIVE_MIXED
            H.loc[[o[0] for o in offs], "derivative_type"] = label

    print(f"  -> Derivative subType: {H['derivative_type'].value_counts().to_dict()}")
    return H


def apply_others_industry(H):
    """
    Last step after the AMC label mapping and the NSE fallback (market-cap file):
    rows of every asset class still without a sector get "Others" as industry and
    sector (industry_source = Others).
    """
    others = H["sector"].isna()
    H.loc[others, ["industry", "sector"]] = "Others"
    H.loc[others, "industry_source"] = "Others"
    print(f"  -> Industry / sector set to Others (no AMC or NSE match): {int(others.sum())} rows")
    return H


def save_holdings_json(df, output_file):
    """
    Save holdings in the exact API schema:

    {
        "YEAR": {
            "CATEGORY": [
                {
                    "fund_name": "...",
                    "scheme_amfi_code": "...",
                    "january": [
                        {
                            "instrument": "...",
                            "isin": "...",
                            "asset_class": "...",
                            "industry": "...",
                            "sector": "...",
                            "holdings": ...,
                            "subType": "...",
                            "symbol": "...",
                            "market_cap": "..."
                        }
                    ]
                }
            ]
        }
    }
    """

    result = {}

    # a hedge leg the AMC leaves out of net assets must not collide with the
    # held position of the same ISIN in Neo4j: keep only the included row
    if "included_in_net_assets" in df:
        inc_flag = df.included_in_net_assets.astype(str).str.lower().isin(["true", "1"])
        clash = df["isin"].notna() & df.duplicated(["source_file", "isin"], keep=False)
        df = df[~(clash & ~inc_flag)]

    for _, row in df.iterrows():

        # ---------------------------------------------------------
        # YEAR
        # ---------------------------------------------------------
        source_file = clean_json_value(row.get("source_file"))

        if source_file:
            parts = re.split(r"[\\/]", str(source_file))

            # source_file is expected to look like:
            # APR/Aditya Birla Capital Mutual Fund/file.xlsx
            #
            # The year comes from the parser ROOT, so use portfolio_date
            # as the reliable year source.
            portfolio_date = clean_json_value(
                row.get("portfolio_date")
            )

            if portfolio_date:
                year = str(portfolio_date)[:4]
            else:
                year = str(int(os.path.basename(
                    os.path.normpath(ROOT)
                )))

        else:
            portfolio_date = clean_json_value(
                row.get("portfolio_date")
            )

            year = (
                str(portfolio_date)[:4]
                if portfolio_date
                else str(int(os.path.basename(
                    os.path.normpath(ROOT)
                )))
            )

        # ---------------------------------------------------------
        # FUND INFORMATION
        # ---------------------------------------------------------
        fund_name = clean_json_value(
            row.get("scheme_name")
        )

        scheme_amfi_code = clean_json_value(
            row.get("amc_scheme_code")
        )

        # ---------------------------------------------------------
        # CATEGORY
        # ---------------------------------------------------------
        category = get_fund_category(
            fund_name,
            source_file
        )

        # ---------------------------------------------------------
        # MONTH
        # ---------------------------------------------------------
        if portfolio_date:
            month_num = int(str(portfolio_date)[5:7])
        else:
            month_num = None

        if month_num is None:
            continue

        month_name = MONTH_NAME[month_num]

        # ---------------------------------------------------------
        # CREATE YEAR
        # ---------------------------------------------------------
        if year not in result:
            result[year] = {}

        # ---------------------------------------------------------
        # CREATE CATEGORY
        # ---------------------------------------------------------
        if category not in result[year]:
            result[year][category] = []

        # ---------------------------------------------------------
        # FIND EXISTING FUND
        # ---------------------------------------------------------
        fund = None

        for existing_fund in result[year][category]:
            if (
                existing_fund["fund_name"] == fund_name
                and existing_fund["scheme_amfi_code"] == scheme_amfi_code
            ):
                fund = existing_fund
                break

        # ---------------------------------------------------------
        # CREATE FUND
        # ---------------------------------------------------------
        if fund is None:
            fund = {
                "fund_name": fund_name,
                "scheme_amfi_code": scheme_amfi_code,
            }

            result[year][category].append(fund)

        # ---------------------------------------------------------
        # CREATE MONTH
        # ---------------------------------------------------------
        if month_name not in fund:
            fund[month_name] = []

        # ---------------------------------------------------------
        # HOLDING
        # ---------------------------------------------------------
        holding = {
            "instrument": clean_json_value(
                row.get("instrument_name")
            ),

            "isin": clean_json_value(
                row.get("isin")
            ),

            "asset_class": clean_json_value(
                row.get("asset_class")
            ),

            "industry": clean_json_value(
                row.get("industry")
            ),

            "sector": clean_json_value(
                row.get("sector")
            ),

            "holdings": clean_json_value(
                row.get("pct_to_nav")
            ),

            "subType": None,

            "symbol": clean_json_value(
                row.get("symbol")
            ),

            "market_cap": clean_json_value(
                row.get("market_cap_category")
            ),

            # False: exposure the AMC leaves out of net assets (most futures), so holdings can
            # exceed or fall short of 100 unless these rows are separated
            "included_in_net_assets": clean_json_value(
                row.get("included_in_net_assets")
            ),
        }

        # ---------------------------------------------------------
        # SUB TYPE
        # ---------------------------------------------------------
        asset_sub_type = clean_json_value(
            row.get("asset_sub_type")
        )

        # asset_sub_type of the row's asset class, e.g. Listed Equity, Preference Shares (Equity),
        # Corporate Bonds / Debentures, interest-rate-swaps-derivative (Debt), Physical (Commodities).
        # Derivative legs take derivative-hedging / derivative-directional-position / derivatives
        # (see apply_derivative_type).
        holding["subType"] = clean_json_value(row.get("derivative_type")) or asset_sub_type

        # ---------------------------------------------------------
        # ADD HOLDING
        # ---------------------------------------------------------
        fund[month_name].append(holding)

    # -------------------------------------------------------------
    # COLLAPSE CASH INTO ONE AGGREGATE ROW PER FUND-MONTH
    # -------------------------------------------------------------
    for categories in result.values():
        for funds in categories.values():
            for fund in funds:
                for month_name, holdings in fund.items():
                    if not isinstance(holdings, list):
                        continue
                    fund[month_name] = collapse_aggregate_holdings(holdings)

    # -------------------------------------------------------------
    # WRITE JSON
    # -------------------------------------------------------------
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(
            result,
            f,
            indent=4,
            ensure_ascii=False,
            allow_nan=False
        )

    save_category_json_files(
        result,
        os.path.join(os.path.dirname(output_file), "categories")
    )


def save_category_json_files(result, category_dir):
    """
    Split {year: {category: [funds]}} into one file per category,
    <category_dir>/<Category>.json, in the same schema, so each category
    can be loaded into Neo4j on its own.
    """

    by_category = {}

    for year, categories in result.items():
        for category, funds in categories.items():
            by_category.setdefault(category, {})[year] = {category: funds}

    os.makedirs(category_dir, exist_ok=True)

    # A full run rewrites every category: drop files of categories that no longer occur
    # (e.g. an old Others.json once every fund maps to a strategy). A --category run keeps the rest.
    if not ARGS.category:
        written = {re.sub(r"[^A-Za-z0-9]+", "_", c).strip("_") + ".json" for c in by_category}
        for old in glob.glob(os.path.join(category_dir, "*.json")):
            if os.path.basename(old) not in written:
                os.remove(old)
                print(f"Removed stale {old}")

    for category, data in sorted(by_category.items()):
        file_name = re.sub(r"[^A-Za-z0-9]+", "_", category).strip("_") + ".json"
        path = os.path.join(category_dir, file_name)

        with open(path, "w", encoding="utf-8") as f:
            json.dump(
                data,
                f,
                indent=4,
                ensure_ascii=False,
                allow_nan=False
            )

        fund_count = sum(len(v[category]) for v in data.values())
        print(f"Wrote {path} ({fund_count} fund-years)")


if __name__ == "__main__":
    if ARGS.split_json:
        with open(ARGS.split_json, "r", encoding="utf-8") as f:
            existing = json.load(f)

        save_category_json_files(
            existing,
            os.path.join(os.path.dirname(os.path.abspath(ARGS.split_json)), "categories")
        )
        sys.exit(0)

    Hs, Ss, Ts = [], [], []

    files = sorted(
        glob.glob(
            os.path.join(
                ROOT,
                glob.escape(ARGS.year) if ARGS.year else "*",           # YEAR
                glob.escape(ARGS.month.upper()) if ARGS.month else "*",  # MONTH
                "*",      # AMC
                "*.xls*"  # FILE
            )
        )
    )

    # Skip Excel lock files ("~$<name>.xlsx") left behind while a workbook is open
    files = [f for f in files if not os.path.basename(f).startswith("~$")]
    # workbooks built from a factsheet go after the real disclosures (the first reconciled one is kept)
    files.sort(key=lambda f: f.endswith(FROM_FACTSHEET_SUFFIX))

    if ARGS.amc:
        files = [
            f for f in files
            if any(a.lower() in os.path.basename(os.path.dirname(f)).lower() for a in ARGS.amc)
        ]

    # Cheap pre-filter on the file path; rows are re-checked against the
    # scheme name after parsing.
    if ARGS.category:
        files = [
            f for f in files
            if get_fund_category(None, os.path.relpath(f, ROOT)) == ARGS.category
        ]

    print(f"\nFound {len(files)} Excel files to process.\n")

    
    SKIPPED_FILES = []
    SEEN, DUPLICATES = {}, []

    for f in files:
        print(f)

        try:
            for H, s, T in parse_file(f):
                key = (s["scheme_name"], s["portfolio_date"])
                # Same scheme-month filed twice (AUG iSIF "... Long Short Fund" + "... Long-Short Fund",
                # MAR Arudha re-issue, Summit "EQLSSIF (2)" sheet): keep one, prefer a reconciled one
                if key in SEEN:
                    k = SEEN[key]
                    if Ss[k]["recon_status"] == "OK" or s["recon_status"] != "OK":
                        print(f"  DUPLICATE: {key[0]} {key[1]} already parsed from {Ss[k]['source_file']} - skipped")
                        DUPLICATES.append(dict(kept=Ss[k]["source_file"], kept_sheet=Ss[k]["source_sheet"],
                                               dropped=s["source_file"], dropped_sheet=s["source_sheet"]))
                        continue
                    print(f"  DUPLICATE: {key[0]} {key[1]} replaces {Ss[k]['source_file']} (that one did not reconcile)")
                    DUPLICATES.append(dict(kept=s["source_file"], kept_sheet=s["source_sheet"],
                                           dropped=Ss[k]["source_file"], dropped_sheet=Ss[k]["source_sheet"]))
                    Hs[k], Ss[k], Ts[k] = H, s, T
                    continue
                SEEN[key] = len(Ss)
                Hs.append(H)
                Ss.append(s)
                Ts.append(T)

        except Exception as e:
            print(f"\nSKIPPED FILE: {f}")
            print(f"ERROR: {type(e).__name__}: {e}\n")

            SKIPPED_FILES.append({
                "source_file": f,
                "error_type": type(e).__name__,
                "error": str(e),
            })

            continue

    # Only combine successfully parsed files
    if not Hs:
        print("No files were parsed successfully.")
        sys.exit(1)

    H = pd.concat(Hs, ignore_index=True)
    H = add_sector_columns(H)
    H = apply_nse_industry_fallback(H)
    H = apply_underlying_industry(H)
    H = apply_others_industry(H)
    H = apply_derivative_type(H)
    S = pd.DataFrame(Ss)
    T = pd.concat(Ts, ignore_index=True)
        
    if ARGS.category:
        H = H[[
            get_fund_category(scheme, source) == ARGS.category
            for scheme, source in zip(H["scheme_name"], H["source_file"])
        ]].reset_index(drop=True)

    # A category run writes holdings_<Category>.json etc. so it never
    # overwrites the full-run outputs.
    suffix = (
        "_" + re.sub(r"[^A-Za-z0-9]+", "_", ARGS.category).strip("_")
        if ARGS.category else ""
    )

    os.makedirs(OUT_DIR, exist_ok=True)
    OUTPUT_HOLDING_COLUMNS = [
    "source_file",
    "source_sheet",
    "source_row",
    "record_type",
    "amc_name",
    "scheme_name",
    "amc_scheme_code",
    "portfolio_date",
    "amc_section_l1",
    "amc_section_l2",
    "instrument_name",
    "isin",
    "symbol",
    "amc_security_code",
    "asset_class",
    "asset_sub_type",
    "derivative_type",
    "listing_status",
    "industry",
    "sector",
    "industry_source",
    "market_cap_category",
    "pct_to_nav",
    "pct_below_threshold",
    "included_in_net_assets",
    "is_placed_as_margin",
]
    # H.to_json(os.path.join(OUT_DIR, "holdings.json"), index=False,indent=4,orient="records")
    H[OUTPUT_HOLDING_COLUMNS].to_csv(
        os.path.join(OUT_DIR, f"holdings{suffix}.csv"),
        index=False
    )
    save_holdings_json(
    H[OUTPUT_HOLDING_COLUMNS],
    os.path.join(OUT_DIR, f"holdings{suffix}.json")
)

    S.to_csv(
        os.path.join(OUT_DIR, f"scheme_snapshot{suffix}.csv"),
        index=False
    )

    T.to_csv(
        os.path.join(OUT_DIR, f"amc_reported_totals{suffix}.csv"),
        index=False
    )
    pd.DataFrame(UNCLASSIFIED_ROWS, columns=["source_file", "source_row", "section_path", "row_label", "market_value_lakhs", "pct_raw"]).to_csv(
        os.path.join(OUT_DIR, f"unclassified_rows{suffix}.csv"),
        index=False
    )
    print(f"\nDropped unclassifiable rows: {len(UNCLASSIFIED_ROWS)} (see unclassified_rows{suffix}.csv)")
    if DUPLICATES:
        print(f"Duplicate scheme-months skipped: {len(DUPLICATES)}")
        for d in DUPLICATES:
            print(f"  kept {d['kept']} [{d['kept_sheet']}] | dropped {d['dropped']} [{d['dropped_sheet']}]")
    if SKIPPED_FILES:
        print(f"Skipped files: {len(SKIPPED_FILES)}")
        for sk in SKIPPED_FILES:
            print(f"  {sk['source_file']}: {sk['error_type']}: {sk['error']}")
    pd.set_option("display.width", 250); pd.set_option("display.max_rows", 300); pd.set_option("display.max_colwidth", 45)
    print(S[["source_file", "scheme_name", "amc_scheme_code", "reported_date_text", "pct_scale_in_source", "holdings_count", "net_assets_lakhs", "mv_reconciliation_diff_lakhs", "sum_pct_to_nav", "recon_status", "date_check", "derivative_disclosure_rows"]].to_string())
    print(H.groupby(["record_type", "asset_class", "asset_sub_type"]).size())
