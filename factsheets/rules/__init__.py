"""Code extractors: the answer when the fund agent cannot give one, and a cross-check on its answer when it can.

Ported from historical-mf (excel/factsheets/rules). The generic and AMC rules are kept as they are: a SIF is run
by a mutual fund house and often prints its parent's factsheet layout, and the parser passes the parent's name
along with the SIF's ('Magnum SIF SBI'), so amc="SBI" rules run on Magnum SIF too. SIF-only layouts are in sif.py.

Each rule is a plain function that gets a Page and returns some of these fields, or nothing:

    aum_month_end_cr, aum_monthly_average_cr   float, Rs crore
    aum_as_on_date                             'YYYY-MM-DD'
    fund_managers                              [{"name", "managing_since", "managing_since_as_printed", "role"}]

A rule that can tell a field is genuinely absent (ABSL: 'AUM as on last day' prints no date) returns NOT_PRINTED
for it: the field stays empty, and the model is not asked for something the page does not have. ASK_MODEL is
the opposite: the value is there but no rule should guess it (Mirae prints three funds side by side), so the
generic rules are skipped for that field and the model reads it.

Register it with @rule(<fields it can return>), optionally amc="Axis" to run it only on factsheets whose file
name contains that text. For each field the first rule that returns a value wins; AMC-specific rules run before
generic ones, then sif.py's before MF's, otherwise rules run in the order they are defined. A rule that raises is skipped with a warning,
so one bad pattern cannot stop a run.

Adding a case: write the function in the module for its field (aum.py, managers.py) or in a new module added to
the import at the bottom, then run `python excel/factsheets/rules_eval.py` to see how it scores against the
model's earlier answers before relying on it.

    @rule("aum_month_end_cr", amc="HDFC")
    def hdfc_aum(page):
        m = re.search(r"Fund Size\\s*([\\d,.]+)\\s*Cr", page.text)
        return m and {"aum_month_end_cr": number(m.group(1))}
"""
import re
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple

FIELDS = ("aum_month_end_cr", "aum_monthly_average_cr", "aum_as_on_date", "fund_managers")
NOT_PRINTED = object()
ASK_MODEL = object()  # the page prints it, but not in a way a rule can read (Mirae's side-by-side columns)


@dataclass
class Page:
    """What a rule can read for one fund: its PDF text layer (all its pages so far) and docling markdown."""
    fund_name: str
    text: str
    markdown: str = ""
    amc: str = ""


@dataclass
class Rule:
    fn: Callable[[Page], Optional[dict]]
    fields: Tuple[str, ...]
    amc: Optional[str] = None


RULES: List[Rule] = []


def rule(*fields, amc=None):
    unknown = set(fields) - set(FIELDS)
    if unknown:
        raise ValueError(f"unknown field(s) {sorted(unknown)}; choose from {FIELDS}")

    def register(fn):
        RULES.append(Rule(fn, fields, amc))
        return fn
    return register


@dataclass
class Found:
    values: Dict[str, object] = field(default_factory=dict)
    source: Dict[str, str] = field(default_factory=dict)  # field -> rule that produced it
    not_printed: set = field(default_factory=set)          # fields a rule found to be absent from the page
    for_model: set = field(default_factory=set)            # fields left to the model (ASK_MODEL)

    def has(self, f):
        return f in self.values or f in self.not_printed

    def settled(self, f):
        return self.has(f) or f in self.for_model


_SPELLINGS = [("largecap", "large cap"), ("midcap", "mid cap"), ("smallcap", "small cap"), ("flexicap", "flexi cap"),
              ("cycles", "cycle"), ("&", " and ")]
_BRAND_WORDS = {"fund", "the", "mutual", "absl", "sun", "life", "prudential", "pru", "capital", "india", "mahindra", "sif"}


def compact(s):
    """Lower-case letters only, common spellings unified: 'Large & Midcap' and 'LARGE AND MID CAP' compare equal."""
    s = s.lower()
    for a, b in _SPELLINGS:
        s = s.replace(a, b)
    return re.sub(r"[^a-z]", "", s)


def fund_core(fund_name, amc=""):
    """The part of a fund's name that tells it apart: 'Kotak Conservative Hybrid Fund' -> 'conservativehybrid'."""
    brand = set(re.sub(r"[^a-z ]", " ", amc.lower()).split()) | _BRAND_WORDS
    words = re.sub(r"\([^)]*\)", " ", fund_name.lower().replace("&", " and ")).split()
    return compact(" ".join(w for w in words if re.sub(r"[^a-z]", "", w) not in brand))


# earlier names a factsheet may still print, by distinguishing name, e.g. {"conservativehybrid": ("debthybrid",)}
# (Kotak MF's 'KOTAK DEBT HYBRID FUND'); none seen in the SIF factsheets so far
FORMER_NAMES = {}


def fund_cores(fund_name, amc=""):
    core = fund_core(fund_name, amc)
    return {core, *FORMER_NAMES.get(core, ())} - {""}


def names_fund(page: Page) -> bool:
    """
    The pages carry the fund's own name. Rules read whatever page they are given, so on a neighbouring fund's
    page (Kotak's index sent 'Conservative Hybrid Fund' to the International REIT FoF page) they would answer for
    the wrong fund; they answer nothing instead, and the agent looks for the fund's real page.
    The distinguishing words alone are not enough: 'smallcap' is in every market-cap table (Kotak SEP Small Cap
    read the Pioneer Fund page) and 'focused' in prose ('companies focused on ...', Invesco MAY Technology Fund),
    so the full name or a title line has to be there - or, for a SIF, the distinguishing name followed by 'Fund':
    Apex's details page prints only 'Name Hybrid Long Short Fund', and 'hybridlongshortfund' names one strategy.
    """
    cores = fund_cores(page.fund_name, page.amc)
    if not cores:
        return True
    text = compact(page.text)
    return (compact(page.fund_name) in text or any(len(c) >= 10 and c + "fund" in text for c in cores)
            or bool(title_line(page.text, page.fund_name, page.amc)))


def title_line(text, fund_name, amc=""):
    """
    A line that is the fund's name and nothing else - 'BALANCED ADVANTAGE FUND', 'PGIM India Balanced Advantage
    Fund', 'KOTAK DEBT HYBRID FUND' - as a fund's own page prints at its top; a mention inside a sentence or a list
    of schemes ('Scheme Name PGIM India Global Select ... PGIM India Balanced Advantage Fund') is not one.
    """
    cores = fund_cores(fund_name, amc)
    for line in text.split("\n"):
        line = " ".join(re.sub(r"\(.*?\)|\(.*$", " ", line).split()).strip(" .:-")
        if line and len(line) <= 80 and re.search(r"\bfund\b|\bfof\b|\bsif\b", line, re.I) and fund_core(line, amc) in cores:
            return line
    return None


def printed_name(page: Page):
    """The fund's name as the page prints it ('HDFC Flexi Cap Fund'), else None."""
    return title_line(page.text, page.fund_name, page.amc)


def clean_text(text):
    """
    Text-layer quirks that break patterns: U+FFFE is how some layers print a hyphen at a line break ('01-Mar\\ufffe2025',
    Tata); fonts with ligatures print 'Shetty' as 'She\\u01a9y' (Motilal).
    """
    for a, b in (("\r\n", "\n"), ("￾", "-"), ("­", "-"), ("Ʃ", "tt"), ("ﬀ", "ff"), ("ﬁ", "fi"),
                 ("ﬂ", "fl"), ("ﬃ", "ffi")):
        text = text.replace(a, b)
    return text


def run(page: Page) -> Found:
    found = Found()
    page.text = clean_text(page.text)
    if page.fund_name and page.amc and not names_fund(page):
        return found
    # stable: AMC-specific first, then the SIF layouts (sif.py imports aum / managers, so theirs register first),
    # then definition order
    for r in sorted(RULES, key=lambda r: (r.amc is None, not r.fn.__module__.endswith(".sif"))):
        if r.amc and r.amc.lower() not in page.amc.lower():
            continue
        if all(found.settled(f) for f in r.fields):
            continue
        try:
            out = r.fn(page) or {}
        except Exception as e:
            print(f"  ! rule {r.fn.__name__} failed on {page.fund_name}: {type(e).__name__}: {e}")
            continue
        for f, v in out.items():
            if f not in r.fields or found.settled(f) or v is None or v == [] or v == "":
                continue
            if v is ASK_MODEL:
                found.for_model.add(f)
                continue
            if v is NOT_PRINTED:
                found.not_printed.add(f)
            else:
                found.values[f] = v
            found.source[f] = r.fn.__name__
    return found


from . import sif, aum, managers  # noqa: E402,F401  (registers the rules; sif's first, so its layouts win)
