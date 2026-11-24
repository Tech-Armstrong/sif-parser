"""
Industry -> Sector mapping based on the SEBI/AMFI standardized industry
classification framework (v.July-2023).

Files in this folder:
  sebi_industry_classification.csv   macro_economic_sector, sector, industry, basic_industry
                                     (the official 4-level table: 12 / 22 / 59 / 197)
  industry_alias.csv                 raw industry label as AMCs print it -> SEBI industry.
                                     Blank industry = deliberately unmapped (non-equity,
                                     index, MF units, mixed buckets ...).

Resolution order for a raw label:
  1. industry_alias.csv            (normalized key)
  2. SEBI industry name            (normalized key)
  3. SEBI basic-industry name      (rolled up to its industry)
  else None.

Normalization ignores case, extra spaces, commas, trailing '#', spacing around
'-' and 'and' vs '&', so "Cement and Cement Products", "BANKS" and
"It - Software" resolve without an alias row.

CLI: fill sector columns into an existing holdings CSV and report unmapped labels
  python excel\\industry_sector_mapping\\industry_sector.py excel\\standardized_format\\output\\holdings.csv
"""
from pathlib import Path
import csv
import re
import sys

HERE = Path(__file__).resolve().parent
SEBI_FILE = HERE / "sebi_industry_classification.csv"
ALIAS_FILE = HERE / "industry_alias.csv"

_UNMAPPED = object()


def industry_key(label):
    if label is None:
        return ""
    s = str(label).strip().lower()
    if s in ("", "nan", "none"):
        return ""
    s = s.rstrip("#").strip()
    s = s.replace(",", " ")
    s = re.sub(r"\band\b", "&", s)
    s = re.sub(r"\s*-\s*", "-", s)
    s = re.sub(r"\s*&\s*", " & ", s)
    return " ".join(s.split())


class IndustrySectorMapper:
    def __init__(self, sebi_file=SEBI_FILE, alias_file=ALIAS_FILE):
        self.industries = {}   # industry -> (sector, macro)
        by_key, basic = {}, {}
        with open(sebi_file, encoding="utf-8") as f:
            for r in csv.DictReader(f):
                self.industries[r["industry"]] = (r["sector"], r["macro_economic_sector"])
                by_key[industry_key(r["industry"])] = r["industry"]
                basic.setdefault(industry_key(r["basic_industry"]), r["industry"])
        self.aliases = {}
        if Path(alias_file).exists():
            with open(alias_file, encoding="utf-8") as f:
                for r in csv.DictReader(f):
                    ind = (r.get("industry") or "").strip()
                    if ind and ind not in self.industries:
                        raise ValueError(f"{alias_file}: '{ind}' is not a SEBI industry")
                    self.aliases[industry_key(r["raw_industry"])] = ind or _UNMAPPED
        self._by_key, self._basic = by_key, basic
        self._cache = {}

    def industry(self, raw):
        """Raw AMC label -> canonical SEBI industry (or None)."""
        k = industry_key(raw)
        if not k:
            return None
        if k not in self._cache:
            hit = self.aliases.get(k) or self._by_key.get(k) or self._basic.get(k)
            self._cache[k] = None if hit is _UNMAPPED else hit
        return self._cache[k]

    def sector(self, raw):
        ind = self.industry(raw)
        return self.industries[ind][0] if ind else None

    def macro_sector(self, raw):
        ind = self.industry(raw)
        return self.industries[ind][1] if ind else None

    def is_known(self, raw):
        """True if the label is either mapped or deliberately left unmapped."""
        k = industry_key(raw)
        return not k or k in self.aliases or k in self._by_key or k in self._basic


_DEFAULT = None


def get_mapper():
    global _DEFAULT
    if _DEFAULT is None:
        _DEFAULT = IndustrySectorMapper()
    return _DEFAULT


def get_sector(raw):
    return get_mapper().sector(raw)


def add_sector_columns(df, industry_col="industry"):
    """Add sebi_industry, sector, macro_economic_sector columns to a DataFrame."""
    m = get_mapper()
    src = df[industry_col]
    df["sebi_industry"] = src.map(m.industry)
    df["sector"] = src.map(m.sector)
    df["macro_economic_sector"] = src.map(m.macro_sector)
    return df


if __name__ == "__main__":
    import pandas as pd

    if len(sys.argv) < 2:
        sys.exit("usage: industry_sector.py <holdings.csv> [out.csv]")
    src = Path(sys.argv[1])
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else src.with_name(src.stem + "_with_sector.csv")
    df = pd.read_csv(src, low_memory=False)
    add_sector_columns(df)
    df.to_csv(out, index=False)
    m = get_mapper()
    unknown = df.loc[~df["industry"].map(m.is_known), "industry"].value_counts()
    eq = df["asset_class"].eq("Equity") & df["industry"].notna()
    print(f"wrote {out}")
    print(f"equity rows with industry: {eq.sum()}, sector filled: {df.loc[eq, 'sector'].notna().sum()}")
    if len(unknown):
        print("\nNEW labels not in industry_alias.csv / SEBI table (add them):")
        print(unknown.to_string())
