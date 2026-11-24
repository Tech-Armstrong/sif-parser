"""Company name -> ISIN, symbol and market cap from the AMFI market-cap lists, for holdings printed without an ISIN.

SIF factsheets print their holdings by name only ('The Indian Hotels Company Limited', 'Adani Port and Special
Economic Zone Ltd.', 'FSN E-COMMERCE VENTURES Ltd.'). The AMFI six-monthly lists in market_cap_mapping/ (the same
files historical-mf uses) carry each listed company's name, ISIN, NSE / BSE symbol, SEBI market-cap bucket and NSE
industry. Both sides of the match go through one standardisation, sif_parser.company_key plus a few factsheet
spellings (see match_key), and a name is tried, first hit wins:

    alias      factsheets/factsheet_instrument_alias.csv (printed_name -> isin), hand-made fixes
    exact      the standardised name in the period's AMFI list (Jan-Jun: June list, Jul-Dec: December list, else
               the latest list before it, as sif_parser does)
    other      the standardised name in another period's list, nearest first (listed after the period's cut-off),
               then in market_cap_mapping/supplementary/ (historical-mf's AMFI list with the newer listings)
    symbol     the name is an NSE symbol ('Optstkdrreddy' prints the contract, 'DRREDDY' the stock)
    prefix     a unique whole-word prefix ('tvs motor' / 'tvs motor company')
    fuzzy      one close name (difflib ratio >= 0.92, clear of the runner-up): accepted but flagged for review

A name none of these settle stays unmatched and is reported; add it to the alias file.
"""
import difflib
import re
import sys
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent))
import sif_parser as sp  # noqa: E402  (AMFI loaders and company_key; importing it does not run the parser)

ALIAS_FILE = ROOT / "factsheet_instrument_alias.csv"
FUZZY_CUTOFF = 0.92
FUZZY_MARGIN = 0.04

# factsheet spellings company_key leaves apart: the AMFI list's spelling on the right
WORD_FIXES = {"pvt": "private", "intl": "international", "mfg": "manufacturing", "inds": "industries",
              "engg": "engineering", "fin": "finance", "svcs": "services", "tech": "technologies",
              "elec": "electricals"}  # WSIF 'Crompton Greaves Consumer Elec Ltd'


def match_key(name):
    """
    sif_parser.company_key plus: 'Ltd.' cut short by a wrap, word fixes ('Engg' -> 'engineering'), plurals
    ('Adani Port' / 'Adani Ports'), and words split by a hyphen ('E-Commerce' / 'Ecommerce').
    """
    n = re.sub(r"\bltd\b\.?|\blimited\b", " ", str(name or ""), flags=re.I)
    # a one- or two-letter prefix is part of the word ('E-Commerce' / 'Ecommerce'); other hyphens split words
    # ('Cube Highways Trust-InvIT Fund', 'Indo-MIM')
    n = re.sub(r"\b([a-z]{1,2})-(?=[a-z])", r"\1", n, flags=re.I)
    n = n.replace("-", " ")
    key = sp.company_key(n)
    words = [WORD_FIXES.get(w, w) for w in key.split()]
    words = [w[:-1] if len(w) > 3 and w.endswith("s") and not w.endswith("ss") else w for w in words]
    # initials written apart ('J.K. Cement' -> 'j k cement') are one word ('JK Cement')
    out, in_run = [], False
    for w in words:
        single = len(w) == 1 and w.isalpha()
        if single and in_run:
            out[-1] += w
        else:
            out.append(w)
        in_run = single
    return " ".join(out)


@dataclass
class Match:
    printed_name: str
    isin: str = None
    company: str = None
    symbol: str = None
    market_cap: str = None
    industry: str = None
    method: str = "unmatched"
    score: float = None
    candidates: str = None  # the runner-up names, when fuzzy or ambiguous

    @property
    def matched(self):
        return bool(self.isin)

    @property
    def needs_review(self):
        return self.method in ("fuzzy", "unmatched")


def periods():
    """Every AMFI list present as (year, month_num), month 6 (June list) or 12 (December list)."""
    return sorted((y, 6 if half == 0 else 12) for y, half in sp.available_market_cap_periods())


def period_of(year, month_num):
    """The list sif_parser uses for a portfolio month (the latest earlier one while the period's is not out yet)."""
    f = sp.get_market_cap_file(year, month_num)
    m = re.match(r"market_cap_(June|December)_(\d{4})", f.name)
    return int(m.group(2)), 6 if m.group(1) == "June" else 12


@lru_cache(maxsize=None)
def name_index(year, month_num):
    """match_key -> [AMFI entries] over one list; ordinary shares (INE<co>01...) first, so they win a tie."""
    out = {}
    for isin, d in sorted(sp.get_market_cap_mapping(year, month_num).items(),
                          key=lambda kv: not (kv[0].startswith("INE") and kv[0][7:9] == "01")):
        if d.get("company"):
            out.setdefault(match_key(d["company"]), []).append(d)
    return out


@lru_cache(maxsize=None)
def supplementary_index():
    """match_key -> [entries] over sif_parser's supplementary AMFI list (listings after the latest half-year list)."""
    out = {}
    for d in sp.get_supplementary().values():
        if d.get("company"):
            out.setdefault(match_key(d["company"]), []).append(d)
    return out


@lru_cache(maxsize=None)
def symbol_index(year, month_num):
    out = {}
    for d in sp.get_market_cap_mapping(year, month_num).values():
        if d.get("symbol"):
            out.setdefault(d["symbol"].upper(), d)
    return out


@lru_cache(maxsize=None)
def aliases():
    """{(fund_house or '', compact printed name): isin} from factsheet_instrument_alias.csv."""
    if not ALIAS_FILE.exists():
        return {}
    df = pd.read_csv(ALIAS_FILE, dtype=str).fillna("")
    return {(r.fund_house.strip(), compact(r.printed_name)): r.isin.strip().upper() for r in df.itertuples() if r.isin.strip()}


def compact(s):
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def entry(isin, year, month_num):
    """AMFI entry for an ISIN (any period, ISIN aliases, manual overrides), as sif_parser looks it up."""
    d = sp.get_market_cap_data(isin, year, month_num)
    return dict(d, isin=isin)


def as_match(printed, d, method, score=None, candidates=None):
    return Match(printed_name=printed, isin=d["isin"], company=d.get("company") or None, symbol=d.get("symbol"),
                 market_cap=d.get("market_cap_category"), industry=d.get("industry"), method=method,
                 score=score, candidates=candidates)


def match(name, year, month_num, fund_house=""):
    """The AMFI company a printed holding name is, or an unmatched Match."""
    printed = " ".join(str(name or "").split())
    if not printed:
        return Match(printed_name=printed)
    al = aliases()
    isin = al.get((fund_house, compact(printed))) or al.get(("", compact(printed)))
    if isin:
        return as_match(printed, entry(isin, year, month_num), "alias")

    key = match_key(printed)
    if not key:
        return Match(printed_name=printed)
    here = period_of(year, month_num)
    order = [here] + sorted((p for p in periods() if p != here),
                            key=lambda p: (abs(p[0] * 12 + p[1] - here[0] * 12 - here[1]), -(p[0] * 12 + p[1])))

    for i, p in enumerate(order):
        hits = name_index(*p).get(key)
        if hits:
            return as_match(printed, hits[0], "exact" if i == 0 else f"exact ({'June' if p[1] == 6 else 'December'} {p[0]} list)")
    hits = supplementary_index().get(key)
    if hits:
        return as_match(printed, hits[0], f"exact ({hits[0]['period']})")

    sym = re.sub(r"[^A-Z0-9&-]", "", printed.upper())
    if sym and " " not in printed.strip():
        d = symbol_index(*here).get(sym)
        if d:
            return as_match(printed, d, "symbol")

    for p in order:
        idx = name_index(*p)
        hits = [k for k in idx if k.startswith(key + " ") or key.startswith(k + " ")]
        if len(hits) == 1:
            return as_match(printed, idx[hits[0]][0], "prefix")
        if len(hits) > 1:
            break  # ambiguous in the nearest list: do not look further back

    idx = name_index(*here)
    scored = sorted(((difflib.SequenceMatcher(None, key, k).ratio(), k) for k in
                     difflib.get_close_matches(key, list(idx), n=3, cutoff=FUZZY_CUTOFF - 0.1)), reverse=True)
    if scored:
        best, k = scored[0]
        runner = scored[1][0] if len(scored) > 1 else 0
        others = "; ".join(idx[k2][0]["company"] for _, k2 in scored[1:]) or None
        if best >= FUZZY_CUTOFF and best - runner >= FUZZY_MARGIN:
            return as_match(printed, idx[k][0], "fuzzy", round(best, 3), others)
        return Match(printed_name=printed, method="unmatched", score=round(best, 3),
                     candidates="; ".join(idx[k2][0]["company"] for _, k2 in scored))
    return Match(printed_name=printed)


def standard_name(m: Match):
    """The name a matched holding is written under: the AMFI company name, 'Ltd' standardised as sif_parser does."""
    return sp.standardize_ltd(m.company) if m.company else m.printed_name
