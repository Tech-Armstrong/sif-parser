"""Fund manager / AUM extraction from SIF monthly factsheets (PDF) with docling and a LangGraph extraction agent.

Usage:
    python factsheets/factsheet_parser.py [--year 2026] [--month JUN] [--amc Magnum]
                                          [--category "Hybrid Long-Short"] [--force | --force-llm]

Ported from historical-mf (excel/factsheets/factsheet_parser.py); the fund agent, its checks, the code rules and the
caching are the same. API_KEY, endpoint and deployment_name in the repo-root .env point the agent at a GPT-5
deployment on Azure OpenAI; without them the run carries on with the code rules alone (see fund_extractor.py).

Factsheets are read from data/factsheets/<YEAR>/<MON>/<Name> Factsheet.pdf, one per SIF and month (the portfolio
disclosures sif_parser.py reads are in data/monthly_portfolio/<YEAR>/<MON>/<SIF>/). <Name> is the SIF's fund house in
scheme_master_mapping/sif_scheme_master.xlsx, or the start of it ('Dyna Factsheet.pdf' -> DynaSIF); see factsheet_house(). A SIF factsheet runs to 2-26
pages, often without an index, so where MF reads the index this finds each fund's pages itself:

  Pass 1 (pages)  copies of one factsheet saved under two names (say 'Magnum Factsheet.pdf' and
                  'Magnum SIF Factsheet.pdf') are read once - the same text layer is the same factsheet.
                  A SIF that publishes one PDF per strategy has several (Altiva's are merged into one per month). For every scheme master strategy of
                  the fund house, every page of every PDF is scored from the text layer (milliseconds): it prints the
                  strategy's name (with or without the brand: Apex's details page says 'Name Hybrid Long Short
                  Fund'), as a title line, with fund-manager and AUM labels, and few other strategies' titles. The best
                  page is the strategy's; the pages after it that still name it come along. A strategy no page names
                  is listed as missing (not launched yet, or printed under a name that needs a row in
                  factsheet_scheme_alias.csv) -> output/<YEAR>/<MON>/<SIF>/index.md
  Pass 2 (funds)  as in MF: docling converts every strategy's pages in one call per PDF, and a tool-calling
                  LangGraph agent per strategy (fund_extractor.py, run in parallel) reads them, decides itself which
                  further pages to read or search, and submits the managers, their dates, the AUM and its date; our
                  code rules (rules/) answer only when it cannot (pass2_graph).
                  -> output/<YEAR>/<MON>/<SIF>/funds.md, fund_details.csv and
                     output/factsheet_fund_managers.csv (every factsheet, one row per fund manager)

Docling output is cached per PDF page under output/<YEAR>/<MON>/<SIF>/pages/<pdf>/ and the model's answers per fund
under .../llm/ (keyed on model, prompt version and page content), so a re-run costs nothing for unchanged pages;
--force redoes both, --force-llm only the LLM step.

The factsheet month is the folder name; the data date (as_on_date, portfolio_month) is read from the pages.

A page whose text layer is unreadable (Magnum DEC 2025 prints in a font whose text layer is every letter shifted by
29: '0DJQXP' for 'Magnum') gets the decoded words added for page finding and the rules, and goes to the model as a
picture as well.

A person printed under several spellings gets one name through factsheet_manager_alias.csv.
"""
import argparse
import hashlib
import operator
import os
import re
import sys
from collections import Counter
from itertools import takewhile
from pathlib import Path
from typing import Annotated, TypedDict

# docling's models are already downloaded: without this every run spends ~3-6 min checking the Hugging Face hub
# before the first page (set HF_HUB_OFFLINE=0 once to fetch models on a new machine)
os.environ.setdefault("HF_HUB_OFFLINE", "1")

import pandas as pd
import pypdfium2 as pdfium
from langgraph.graph import END, START, StateGraph
from langgraph.types import Send

import fund_extractor as fx
import rules
from fund_extractor import Pages
from rules.common import DATE, NAME, NUM, TITLE, number, parse_date

ROOT = Path(__file__).resolve().parent
PROJECT = ROOT.parent
DATA = PROJECT / "data" / "factsheets"                 # factsheet PDFs: <YEAR>/<MON>/<Name> Factsheet.pdf
PORTFOLIO_DATA = PROJECT / "data" / "monthly_portfolio"  # portfolio disclosures: <YEAR>/<MON>/<SIF>/*.xls*
OUT = ROOT / "output"
MASTER = PROJECT / "scheme_master_mapping" / "sif_scheme_master.xlsx"
ALIAS = ROOT / "factsheet_scheme_alias.csv"
MANAGER_ALIAS = ROOT / "factsheet_manager_alias.csv"

MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

# sif_scheme_master category -> category name used by sif_parser.py (get_fund_category)
CATEGORIES = {
    "EQUITY LONG-SHORT": "Equity Long-Short", "HYBRID LONG-SHORT": "Hybrid Long-Short",
    "EQUITY EX-TOP 100 LONG-SHORT": "Equity Ex-Top 100 Long-Short",
    "ACTIVE ASSET ALLOCATOR LONG-SHORT": "Active Asset Allocator Long-Short",
    "SECTOR ROTATION LONG-SHORT": "Sector Rotation Long-Short",
}

# the mutual fund house behind each SIF: its name is passed to the rules with the SIF's, so a rule written for the
# house's MF layout (amc="SBI") runs on its SIF too (Magnum prints SBI's 'Fund Manager Details' table)
PARENT_AMC = {
    "Altiva SIF": "Edelweiss", "Apex SIF": "Aditya Birla", "Arudha SIF": "Bandhan", "Diviniti SIF": "ITI",
    "iSIF": "ICICI Prudential", "Magnum SIF": "SBI", "Prism SIF": "Jio BlackRock", "RedHex SIF": "HSBC",
    "Summit SIF": "Invesco",
}

MAX_FUND_PAGES = 3
LLM_WORKERS = 4  # fund agents run in parallel (the model's rate limiter still caps requests per minute)
DOCLING_THREADS = 8  # ~30% faster per page than docling's default 4, same output (16 is no faster)


def amc_tag(house):
    """The AMC text the rules and name matching get: 'Magnum SIF' -> 'Magnum SIF SBI'."""
    return f"{house} {PARENT_AMC[house]}" if house in PARENT_AMC else house


def slug(s):
    return re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")


# ---------------------------------------------------------------- docling
_CONVERTER = None


def converter():
    global _CONVERTER
    if _CONVERTER is None:
        from docling.datamodel.accelerator_options import AcceleratorOptions
        from docling.datamodel.base_models import InputFormat
        from docling.datamodel.pipeline_options import PdfPipelineOptions, TableFormerMode
        from docling.document_converter import DocumentConverter, PdfFormatOption

        opts = PdfPipelineOptions()
        opts.accelerator_options = AcceleratorOptions(num_threads=DOCLING_THREADS)
        opts.do_ocr = False  # factsheets carry a text layer
        opts.do_table_structure = True  # Magnum's 'Fund Manager Details' and Apex's 'Scheme Details' tables
        opts.table_structure_options.mode = TableFormerMode.FAST
        _CONVERTER = DocumentConverter(format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=opts)})
    return _CONVERTER


def docling_pages(pdf, page_nos, cache_dir, force=False, source=None):
    """
    Markdown for the given 1-based PDF pages, docling-converted in one call on a small PDF holding only those
    pages. traverse_pictures keeps text docling files under a picture (an illustrated sidebar). Cached as
    pages/<pdf>/page_NNN.md. source identifies the PDF's content: a different factsheet saved under the same file
    name (a month's PDF replaced by the right one) empties the cache instead of reusing the old file's pages.
    """
    cache_dir.mkdir(parents=True, exist_ok=True)
    marker = cache_dir / "_source.txt"
    if source:
        if marker.exists() and marker.read_text(encoding="utf-8") != source:
            for old in cache_dir.glob("page_*.md"):
                old.unlink()
        marker.write_text(source, encoding="utf-8")
    path = lambda p: cache_dir / f"page_{p:03d}.md"
    todo = sorted({p for p in page_nos if force or not path(p).exists()})
    if todo:
        src = pdfium.PdfDocument(str(pdf))
        mini = pdfium.PdfDocument.new()
        mini.import_pages(src, [p - 1 for p in todo])
        tmp = cache_dir / "_pages.pdf"
        mini.save(str(tmp))
        mini.close()
        src.close()
        try:
            doc = converter().convert(str(tmp)).document
            for k, p in enumerate(todo, start=1):
                md = doc.export_to_markdown(page_no=k, traverse_pictures=True)
                md = re.sub(r"[.…·]{4,}", " … ", md)  # dot leaders
                path(p).write_text(md, encoding="utf-8")
        finally:
            tmp.unlink(missing_ok=True)
    return {p: path(p).read_text(encoding="utf-8") for p in page_nos}


# ---------------------------------------------------------------- text layer
SHIFT = 29  # '0DJQXP' + 29 = 'Magnum'
SHIFTED_RUN = re.compile(r"[\x21-\x5d]{4,}")
SHIFTED_WORDS = re.compile(r"Fund|Manag|Details|Since|Experience|Allotment|Benchmark|Investment|Objective|Plans|Load", re.I)


def repair(text):
    """
    (text, unreadable): a page printed in a font whose text layer is every letter shifted by 29 gets the decoded
    words appended ('0DJQXP+\\EULG/RQJ6KRUW)XQG' -> 'MagnumHybridLongShortFund'), which is enough for the page
    finder, the name checks and most rules; its figures are often in a normal font and stay as they are. Only runs
    that decode to a capitalised word are taken: a real upper-case word ('NIFTY') decodes to lower case.
    """
    words = [w for w in ("".join(chr(ord(c) + SHIFT) for c in m.group(0)) for m in SHIFTED_RUN.finditer(text))
             if re.search(r"[A-Z][a-z]", w)]
    if len(SHIFTED_WORDS.findall(" ".join(words))) < 3:
        return text, False
    return text + "\n\n[unreadable text layer, decoded words:]\n" + "\n".join(words), True


def read_pdf(pdf):
    """(page texts, unreadable pages) of one PDF."""
    doc = pdfium.PdfDocument(str(pdf))
    try:
        raw = [doc[i].get_textpage().get_text_range().replace("\r\n", "\n") for i in range(len(doc))]
    finally:
        doc.close()
    texts, unreadable = [], set()
    for i, t in enumerate(raw):
        t, bad = repair(t)
        texts.append(t)
        if bad:
            unreadable.add(i + 1)
    return texts, unreadable


FACTSHEET_FILE = re.compile(r"^\s*(?P<name>.+?)\s+factsheet\s*$", re.I)


def factsheet_house(pdf, houses):
    """
    The fund house a factsheet belongs to, from its file name '<Name> Factsheet.pdf': the house named <Name>, else the
    one house whose name starts with it ('Dyna Factsheet.pdf' -> DynaSIF, 'Magnum Factsheet.pdf' -> Magnum SIF).
    None when the file is not named that way, or <Name> fits no house or several.
    """
    m = FACTSHEET_FILE.match(Path(pdf).stem)
    if not m:
        return None
    name = " ".join(m["name"].split()).lower()
    exact = [h for h in houses if h.lower() == name]
    if exact:
        return exact[0]
    hits = [h for h in houses if h.lower().startswith(name)]
    return hits[0] if len(hits) == 1 else None


def factsheet_folders(houses, year=None, month=None, amc=None, quiet=False):
    """
    [(year, MON, fund house, [pdf, ...])] for the factsheets in DATA/<YEAR>/<MON>/, in date order. A PDF whose name
    maps to no fund house is reported (unless quiet) and skipped. amc: only houses whose name contains one of these.
    """
    out = []
    for ydir in sorted(p for p in DATA.iterdir() if p.is_dir() and re.fullmatch(r"\d{4}", p.name)):
        if year and ydir.name != str(year):
            continue
        for mdir in sorted((p for p in ydir.iterdir() if p.is_dir() and p.name.upper() in MONTHS),
                           key=lambda p: MONTHS.index(p.name.upper())):
            if month and mdir.name.upper() != month.upper():
                continue
            by_house = {}
            for pdf in sorted(mdir.glob("*.pdf")):
                house = factsheet_house(pdf, houses)
                if house is None:
                    if not quiet:
                        print(f"{ydir.name} {mdir.name} {pdf.name}: ! no fund house of the SIF scheme master matches the "
                              "name (expected '<SIF> Factsheet.pdf'), skipped")
                    continue
                by_house.setdefault(house, []).append(pdf)
            for house in sorted(by_house, key=str.lower):
                if amc and not any(a.lower() in house.lower() for a in amc):
                    continue
                out.append((int(ydir.name), mdir.name.upper(), house, by_house[house]))
    return out


def distinct_pdfs(pdfs):
    """
    [(pdf, texts, unreadable pages, [same factsheet under other names])] for one SIF month's PDFs: copies of one
    factsheet (the same text layer) are read once, under the shortest file name.
    """
    seen = {}
    for pdf in sorted(pdfs, key=lambda p: (len(p.name), p.name.lower())):
        texts, unreadable = read_pdf(pdf)
        key = hashlib.sha1("\f".join(texts).encode("utf-8")).hexdigest()
        if key in seen:
            seen[key][3].append(pdf.name)
        else:
            seen[key] = (pdf, texts, unreadable, [])
    return list(seen.values())


INDEX_HEADING = re.compile(r"^\s*(?:i\s*n\s*d\s*e\s*x|contents|table of contents)\s*$", re.I | re.M)


def index_pages(texts):
    """
    Pages within the first 5 headed INDEX / CONTENTS (iSIF, Platinum): they name every strategy. A page printing an
    AUM figure is a fund's own page, not an index: DynaSIF's 'Benchmark' / 'Index : BSE 500 TRI' put a bare 'Index'
    line on the Active Asset Allocator's page 3, which was then skipped (its managers and AUM lost for MAR-AUG 2026).
    """
    return [i + 1 for i, t in enumerate(texts[:5]) if INDEX_HEADING.search(t) and not has_aum(t)]


# ---------------------------------------------------------------- pass 1: the fund's pages
MANAGER_LABEL = re.compile(r"fund\s*managers?\b", re.I)
PERSON = re.compile(rf"\b{TITLE}[A-Z]|managing|\bsince\b", re.I)
AUM_FIGURE = re.compile(r"(?:\bA?AUM\b|assets\s+under\s+management|fund\s+size)[^\n]{0,80}?\d[\d,]*\.\d|"
                        r"(?:\bA?AUM\b|fund\s+size)[^\d]{0,40}\n[^\d\n]{0,20}\d[\d,]*\.\d", re.I)


def has_managers(text):
    return bool(MANAGER_LABEL.search(text) and PERSON.search(text))


def has_aum(text):
    return bool(AUM_FIGURE.search(text))


def fund_titles(text, amc):
    """How many different funds the page prints as a title line: 1-2 on a fund's own page, 3+ on a summary page."""
    out = set()
    for line in text.replace("\r\n", "\n").split("\n"):
        line = " ".join(re.sub(r"\(.*?\)|\(.*$", " ", line).split()).strip(" .:-")
        if line and len(line) <= 80 and re.search(r"\b(?:fund|fof|sif)$", line, re.I):
            out.add(rules.fund_core(line, amc))
    return len(out - {""})


def page_score(text, name, amc, lone):
    """
    How much a page looks like this fund's own page; None when it does not name the fund (unless the PDF covers
    only this fund: lone). Name +3/+4, title line +3, fund managers +5, an AUM figure +5, other funds' titles -2 each.
    """
    page = rules.Page(fund_name=name, text=text, amc=amc)
    named = rules.names_fund(page)
    if not named and not lone:
        return None
    s = (4 if rules.compact(name) in rules.compact(text) else 3) if named else 0
    if rules.title_line(text, name, amc):
        s += 3
    if has_managers(text):
        s += 5
    if has_aum(text):
        s += 5
    return s - 2 * max(0, fund_titles(text, amc) - 1)


def locate(target, names, pdfs, amc):
    """
    The fund's pages: (pdf entry, [PDF pages], score) or None. The best-scoring page of all PDFs, then the next pages
    of that PDF that still name the fund; when the best page lacks the fund-manager or the AUM block, the best page
    carrying it is added (Apex prints the strategy's name on one page and its details on another).
    """
    best = None
    for entry in pdfs:
        pdf, texts, _, _, skip = entry
        named_somewhere = {n for n in names for i, t in enumerate(texts) if i + 1 not in skip
                           and rules.names_fund(rules.Page(fund_name=n, text=t, amc=amc))}
        if not named_somewhere:
            continue
        # a PDF naming no other strategy of this SIF is this fund's alone: its unnamed pages may be the fund's too
        lone = not any(rules.names_fund(rules.Page(fund_name=o, text=t, amc=amc)) for o in target["others"]
                       for i, t in enumerate(texts) if i + 1 not in skip)
        for i, t in enumerate(texts):
            if i + 1 in skip:
                continue
            for n in names:
                s = page_score(t, n, amc, lone)
                if s is not None and (best is None or s > best[2]):
                    best = (entry, i + 1, s, n, lone)
    if best is None:
        return None
    entry, p, s, name, lone = best
    texts, skip = entry[1], entry[4]
    named = lambda q: rules.names_fund(rules.Page(fund_name=name, text=texts[q - 1], amc=amc))
    pages = [p] + list(takewhile(lambda q: q <= len(texts) and q not in skip and named(q), range(p + 1, p + MAX_FUND_PAGES)))
    for need in (has_managers, has_aum):
        if not any(need(texts[q - 1]) for q in pages):
            extra = [(page_score(texts[q - 1], name, amc, lone), q) for q in range(1, len(texts) + 1)
                     if q not in skip and q not in pages and need(texts[q - 1]) and (lone or named(q))]
            extra = [(sc, q) for sc, q in extra if sc is not None]
            if extra:
                pages.append(max(extra, key=lambda e: (e[0], -e[1]))[1])
    return entry, pages, s, name


def relocate(d, texts, amc, skip_pages):
    """
    The fund's own page when the one first tried turns out to be another fund's: a page with the fund's name as a
    title line and fund-manager and AUM labels; of several, the one printing the fewest fund titles, then the nearest
    the page first tried. None when there is no such page. Text layer only (milliseconds).
    """
    tried = set(d.get("tried", [])) | set(skip_pages)
    near = (d.get("tried") or [d["pdf_page"]])[0]
    found = []
    for i, t in enumerate(texts):
        if i + 1 in tried or not re.search(r"fund\s*manager", t, re.I) \
                or not re.search(r"\bA?AUM\b|assets\s+under\s+management|fund\s+size", t, re.I):
            continue
        t = t.replace("\r\n", "\n")
        if rules.title_line(t, d["index_name"], amc) or rules.title_line(t, d["scheme_name"], amc):
            found.append((fund_titles(t, amc), abs(i + 1 - near), i + 1))
    return min(found)[2] if found else None


# ---------------------------------------------------------------- lookups across the factsheet (from MF)
MANAGED_SINCE = re.compile(rf"(?P<names>(?:\s*(?:,|&|\band\b)?\s*{TITLE}{NAME})+)\s+(?:has|have)\s+"
                           rf"been\s+managing\s+the\s+(?:fund|scheme)\s+since\s+(?P<since>{DATE})")
SCHEME_INCEPTION = re.compile(rf"Scheme\s+Inception\s+date\s+is\s+({DATE})")


def managed_since(segment):
    """{manager: date as printed} from 'Mr. A has been managing the fund since <date> and Ms. B & Mr. C have been ...'."""
    dates = {}
    for m in MANAGED_SINCE.finditer(segment):
        for n in re.split(rf"\s*(?:,|&|\band\b)?\s*{TITLE}", m.group("names")):
            if n.strip(" ,&"):
                dates[" ".join(n.strip(" ,&").split())] = m.group("since")
    return dates


def manager_dates_finder(texts, amc):
    """
    lookup(fund_name, fund_text) -> {manager: date as printed} from fund-manager pages that date each manager apart
    from the fund's page ('Scheme Inception date is 27/07/2005. Ms. X has been managing the fund since 09/05/2019.'),
    matched on the fund page's 'Allotment Date', else on the fund's name in a '<fund> - Growth' heading. Kept from MF:
    no SIF factsheet seen so far prints them so, and then it finds nothing.
    """
    table = None

    def build():
        by_inception, by_name = {}, {}
        for t in texts:
            if not re.search(r"been\s+managing\s+the\s+(?:fund|scheme)\s+since", t):
                continue
            starts = list(SCHEME_INCEPTION.finditer(t))
            for s, nxt in zip(starts, starts[1:] + [None]):
                dates = by_inception.setdefault(parse_date(s.group(1)), {})
                for name, since in managed_since(t[s.end():nxt.start() if nxt else len(t)]).items():
                    dates[name] = since if parse_date(dates.get(name, since)) == parse_date(since) else None
            heads = list(re.finditer(r"^(?P<h>[A-Z][^\n]{3,200}?)\s*-\s*Growth\b", t, re.M))
            for h, nxt in zip(heads, heads[1:] + [None]):
                dates = managed_since(t[h.end():nxt.start() if nxt else len(t)])
                core = rules.fund_core(re.sub(r"\(.*?\)", " ", h.group("h")), amc)
                if dates and core:
                    by_name.setdefault(core, dates)
        return by_inception, by_name

    def lookup(fund_name, fund_text=""):
        nonlocal table
        if table is None:
            table = build()
        by_inception, by_name = table
        allot = re.search(rf"Allotment\s+Date\s*:?\s*({DATE})", fund_text, re.I)
        dates = by_inception.get(parse_date(allot.group(1)), {}) if allot else {}
        named = next((by_name[c] for c in rules.fund_cores(fund_name, amc) if c in by_name), {})
        dates = {n: d for n, d in dates.items() if d}
        ends = {(n.split()[0], n.split()[-1]) for n in dates}
        return {**{n: d for n, d in named.items() if (n.split()[0], n.split()[-1]) not in ends}, **dates}
    return lookup


# ---------------------------------------------------------------- pass 2: the fund agents (from MF)
class Pass2(TypedDict, total=False):
    results: Annotated[dict, operator.or_]          # fund key -> FundDetails (or None); merged across branches
    fs: fx.FactsheetDate


def pass2_graph(pages: Pages, details, out_dir, factsheet_month, force, force_llm, amc):
    """
    LangGraph for pass 2 of one PDF:

        START -> first_pages -> Send(extract_fund) x N in parallel -> END
              \\-> factsheet_date -> END   (in parallel)

    first_pages makes one docling call for every fund's pages (docling's model load is per call, so batching is far
    cheaper than per fund) and renders pages printed as pictures; both run here, not in the parallel branches,
    because pdfium is not thread-safe. Each extract_fund branch is fund_extractor's agent for one fund: it reads
    further pages, searches the factsheet or moves to the fund's real page itself. A fund answered not-found is moved
    once to the page relocate() finds and read again, by the agent or, without the model, by the rules.
    """
    pdf, texts = pages.pdf, pages.texts
    dates_lookup = manager_dates_finder(texts, amc)
    cache_dir = out_dir / "pages" / slug(pdf.stem)

    def first_pages(state: Pass2):
        source = hashlib.sha1("\f".join(texts).encode("utf-8")).hexdigest()
        mds = docling_pages(pdf, sorted({q for d in details.values() for q in d["pages"]}), cache_dir, force, source)
        for d in details.values():
            d["md"] = "\n\n".join(f"<!-- PDF page {q} -->\n\n{mds[q]}" for q in d["pages"])
            d["images"] = [pages.png(q) for q in d["pages"] if pages.is_image(q)]
        return {}

    def fan_out(state: Pass2):
        return [Send("extract_fund", {"key": k}) for k in details] or END

    def run(k, d, cache_name=None):
        raw = "\n\n".join(f"--- PDF page {q} ---\n{texts[q - 1]}" for q in d["pages"])
        try:
            return fx.extract(d["index_name"], d["scheme_name"], d["category"], d["md"], raw, pages, d["pages"],
                              factsheet_month=factsheet_month, images=d.get("images"),
                              cache_file=out_dir / "llm" / f"{cache_name or k}.json", force=force_llm, amc=amc,
                              manager_dates=dates_lookup)
        except Exception as e:  # one bad answer should not drop the rest of the factsheet
            if fx.auth_error(e):
                raise
            print(f"  ! {d['scheme_name']}: {type(e).__name__}: {str(e)[:200]}")
            return None

    def extract_fund(arg):
        k = arg["key"]
        d = details[k]
        res = run(k, d)
        if res is not None and not res.fund_found:
            d["tried"] = list(d["pages"])
            p = relocate(d, texts, amc, pages.skip)
            if p:
                print(f"  {d['scheme_name']}: PDF page {d['pdf_page']} is another fund's, reading page {p}", flush=True)
                more = takewhile(lambda q: q <= len(texts) and q not in pages.skip and pages.names(q, d["index_name"]),
                                 range(p + 1, p + MAX_FUND_PAGES))
                d.update(pdf_page=p, pages=[p, *more], md="", images=[])
                res = run(k, d, cache_name=f"{k}_page{p}")
        return {"results": {k: res}}

    def factsheet_date(state: Pass2):
        # data date read from the opening pages: used for funds whose page prints no AUM / NAV date; a fund's own
        # printed date wins but is flagged when its month differs
        return {"fs": pages.factsheet_date()}

    g = StateGraph(Pass2)
    g.add_node("first_pages", first_pages)
    g.add_node("extract_fund", extract_fund)
    g.add_node("factsheet_date", factsheet_date)
    g.add_edge(START, "first_pages")
    g.add_edge(START, "factsheet_date")
    g.add_conditional_edges("first_pages", fan_out, ["extract_fund", END])
    g.add_edge("extract_fund", END)
    g.add_edge("factsheet_date", END)
    return g.compile()


# ---------------------------------------------------------------- driver
def aum_check(r):
    """Flag for review: month-end and monthly-average AUM rarely differ by more than ~30%."""
    if r is None or r.aum_month_end_cr is None:
        return "missing"
    if r.aum_monthly_average_cr:
        ratio = r.aum_month_end_cr / r.aum_monthly_average_cr
        if not 0.7 <= ratio <= 1.43:
            return f"end/avg ratio {ratio:.2f}"
    return "ok"


def process_house(house, pdf_files, year, mon, master, alias, manager_alias, categories, force, force_llm):
    amc = amc_tag(house)
    out_dir = OUT / str(year) / mon / house
    out_dir.mkdir(parents=True, exist_ok=True)

    # ---- pass 1: each strategy's pages, over the distinct PDFs of the SIF month
    pdfs = [(pdf, texts, bad, dupes, set(index_pages(texts))) for pdf, texts, bad, dupes in distinct_pdfs(pdf_files)]
    expected = master[(master.fund_house == house) & master.category.map(CATEGORIES).isin(categories)]
    every = master[master.fund_house == house].scheme_name.tolist()
    targets, missing = {}, []
    for r in expected.itertuples():
        code = str(r.scheme_code)
        names = [r.scheme_name, *alias[(alias.fund_house == house) & (alias.scheme_code.astype(str) == code)].factsheet_name]
        t = dict(scheme_code=code, scheme_name=r.scheme_name, category=CATEGORIES[r.category],
                 others=[n for n in every if n != r.scheme_name])
        hit = locate(t, names, pdfs, amc)
        if not hit:
            missing.append((r.scheme_name, "no page names it"))
            continue
        entry, pages, score, name = hit
        if not any(has_managers(entry[1][q - 1]) or has_aum(entry[1][q - 1]) for q in pages):
            # Apex JUL names its two coming strategies on the cover only
            missing.append((r.scheme_name, f"`{entry[0].name}` names it on PDF page {pages[0]}, but no page prints its "
                                           "fund managers or AUM"))
            continue
        targets[code] = dict(t, pdf=entry[0], pdf_page=pages[0], pages=pages, score=score, index_name=name)
    write_index_md(out_dir / "index.md", house, year, mon, pdfs, targets, missing)
    if not pdfs:
        print("  ! no PDF")
    elif not targets:
        print(f"  ! no strategy of this SIF found: {'; '.join(f'{m}: {why}' for m, why in missing) or 'none in the scheme master'}")

    # ---- pass 2: per PDF, the per-fund extraction agents (graph above)
    rows, month = [], f"{mon.title()} {year}"
    for pdf, texts, bad, _, skip in pdfs:
        details = {k: dict(t, md="", result=None) for k, t in targets.items() if t["pdf"] == pdf}
        if not details:
            continue
        pages = Pages(pdf, texts, skip=skip, date_cache=out_dir / "llm" / f"_factsheet_date_{slug(pdf.stem)}.json",
                      force=force_llm, pictured=bad)
        out = pass2_graph(pages, details, out_dir, month, force, force_llm, amc).invoke(
            {"results": {}}, {"max_concurrency": LLM_WORKERS})
        for k, res in out["results"].items():
            details[k]["result"] = res
        fs = out["fs"]
        if fx.date_out_of_window(fs.data_as_on_date, month):
            fs = fx.FactsheetDate()
        page_dates = [d["result"].aum_as_on_date for d in details.values() if d["result"] and d["result"].aum_as_on_date]
        as_on = fs.data_as_on_date or (Counter(page_dates).most_common(1)[0][0] if page_dates else None)

        for d in sorted(details.values(), key=lambda d: d["pdf_page"]):
            r = d["result"]
            fund_as_on = (r and r.aum_as_on_date) or as_on
            source = "page" if r and r.aum_as_on_date else "factsheet"
            if source == "page" and fs.data_as_on_date and fund_as_on[:7] != fs.data_as_on_date[:7]:
                source = f"page; factsheet says {fs.data_as_on_date}"
            base = dict(year=year, factsheet_month=mon, as_on_date=fund_as_on, as_on_source=source,
                        portfolio_month=fund_as_on[:7] if fund_as_on else None, amc=house, fund_house=house,
                        factsheet_file=pdf.name, category=d["category"], scheme_code=d["scheme_code"],
                        scheme_name=d["scheme_name"], factsheet_fund_name=(r and r.fund_name_on_page) or d["index_name"],
                        pdf_pages=",".join(map(str, (r and r.fund_found and r.source_pages) or d["pages"])),
                        aum_month_end_cr=r and r.aum_month_end_cr, aum_monthly_avg_cr=r and r.aum_monthly_average_cr,
                        aum_check=aum_check(r), notes=r and r.notes)
            managers = [dict(fund_manager=manager_alias.get((house, m.name), m.name), managing_since=m.managing_since,
                             managing_since_raw=m.managing_since_as_printed, manager_role=m.role)
                        for m in (r.fund_managers if r else [])]
            for mgr in managers or [dict(fund_manager=None, managing_since=None, managing_since_raw=None, manager_role=None)]:
                rows.append(dict(base, **mgr))
            (out_dir / "funds").mkdir(exist_ok=True)
            (out_dir / "funds" / f"{d['scheme_code']}_{slug(d['scheme_name'])}.md").write_text(
                f"# {d['scheme_name']}\n\nscheme_code {d['scheme_code']} | {d['category']} | {pdf.name} PDF pages {d['pages']}\n\n"
                + d["md"], encoding="utf-8")
    write_funds_md(out_dir / "funds.md", house, year, mon, rows)
    pd.DataFrame(rows).to_csv(out_dir / "fund_details.csv", index=False)
    return rows


def md_table(headers, rows):
    esc = lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).replace("|", "/")
    return "\n".join(["| " + " | ".join(headers) + " |", "|" + "---|" * len(headers)]
                     + ["| " + " | ".join(esc(v) for v in r) + " |" for r in rows])


def write_index_md(path, house, year, mon, pdfs, targets, missing):
    parts = [f"# {house} - {mon} {year} factsheet: where each strategy was found", ""]
    parts += [f"- `{pdf.name}`: {len(texts)} pages"
              + (f"; index pages {sorted(skip)}" if skip else "")
              + (f"; unreadable text layer on pages {sorted(bad)} (sent to the model as pictures)" if bad else "")
              + (f"; same factsheet as {', '.join(f'`{n}`' for n in dupes)} (read once)" if dupes else "")
              for pdf, texts, bad, dupes, skip in pdfs] or ["- no factsheet PDF for this SIF month"]
    rows = [(t["category"], t["scheme_code"], t["scheme_name"], t["pdf"].name, ", ".join(map(str, t["pages"])), t["score"])
            for t in sorted(targets.values(), key=lambda t: (t["pdf"].name, t["pdf_page"]))]
    parts += ["", "## Strategies found", "",
              md_table(["Category", "Scheme code", "Scheme master name", "PDF", "PDF pages", "Page score"], rows)]
    if missing:
        parts += ["", "## In the scheme master but not found in this month's factsheet", "",
                  "Not launched yet, or printed under another name (add it to factsheet_scheme_alias.csv).", ""]
        parts += [f"- {m}: {why}" for m, why in missing]
    path.write_text("\n".join(parts) + "\n", encoding="utf-8")


def write_funds_md(path, house, year, mon, rows):
    parts = [f"# {house} - {mon} {year} factsheet: fund managers and AUM", ""]
    for code in dict.fromkeys(r["scheme_code"] for r in rows):
        rs = [r for r in rows if r["scheme_code"] == code]
        r0 = rs[0]
        parts += [f"## {r0['scheme_name']}", "",
                  f"- Category: {r0['category']}",
                  f"- Scheme code: {code}",
                  f"- Name in factsheet: {r0['factsheet_fund_name']}",
                  f"- Factsheet: {r0['factsheet_file']}, PDF pages {r0['pdf_pages']}",
                  f"- AUM as on {r0['as_on_date'] or '?'}: month end {r0['aum_month_end_cr'] if r0['aum_month_end_cr'] is not None else 'NOT FOUND'} Cr,"
                  f" monthly average {r0['aum_monthly_avg_cr'] if r0['aum_monthly_avg_cr'] is not None else 'NOT FOUND'} Cr"
                  f" (check: {r0['aum_check']})"]
        if r0["notes"]:
            parts.append(f"- Notes: {r0['notes']}")
        parts += ["", md_table(["Fund manager", "Role", "Managing since", "As printed"],
                               [(r["fund_manager"] or "NOT FOUND", r["manager_role"], r["managing_since"], r["managing_since_raw"]) for r in rs]), ""]
    if not rows:
        parts.append("No strategy of this SIF found in the factsheet (see index.md).")
    path.write_text("\n".join(parts) + "\n", encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--year", help="Only this year folder, e.g. 2026")
    ap.add_argument("--month", help="Only this month folder, e.g. JUN")
    ap.add_argument("--amc", action="append", help="Only SIF folders whose name contains this text, e.g. Magnum (repeatable)")
    ap.add_argument("--category", action="append", help=f"Only these categories (repeatable). Default: {', '.join(CATEGORIES.values())}")
    ap.add_argument("--force", action="store_true", help="Re-run docling and the LLM even when cached")
    ap.add_argument("--force-llm", action="store_true", help="Re-ask the LLM even when an answer is cached (docling cache kept)")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")

    problem = fx.missing_credentials()
    if problem:
        fx.disable_model(problem)  # the code rules still run; funds they cannot finish are flagged in notes

    master = pd.read_excel(MASTER, sheet_name="scheme_master", dtype={"scheme_code": str})
    alias = pd.read_csv(ALIAS, dtype={"scheme_code": str})
    manager_alias = {(r.fund_house, r.printed_name): r.manager_name for r in pd.read_csv(MANAGER_ALIAS).itertuples()}
    categories = set(args.category or CATEGORIES.values())
    unknown = categories - set(CATEGORIES.values())
    if unknown:
        ap.error(f"unknown category {sorted(unknown)}; choose from {sorted(CATEGORIES.values())}")
    houses = set(master.fund_house)

    folders = factsheet_folders(houses, args.year, args.month, args.amc)

    for year, mon, house, pdf_files in folders:
        print(f"{year} {mon} {house}", flush=True)
        try:
            rows = process_house(house, pdf_files, year, mon, master, alias, manager_alias, categories, args.force, args.force or args.force_llm)
        except Exception as e:  # one bad folder should not stop the batch
            if fx.auth_error(e):
                raise SystemExit(f"Azure OpenAI rejected the request ({type(e).__name__}): check API_KEY, endpoint "
                                 "and deployment_name in the repo-root .env")
            print(f"  ! FAILED: {type(e).__name__}: {e}")
            continue
        funds = {r["scheme_code"] for r in rows}
        no_mgr = {r["scheme_code"] for r in rows if not r["fund_manager"]}
        no_aum = {r["scheme_code"] for r in rows if r["aum_month_end_cr"] is None}
        odd = {r["scheme_code"] for r in rows if r["aum_check"] not in ("ok", "missing")}
        print(f"  {len(funds)} funds, {len(rows)} manager rows; missing manager: {len(no_mgr)}, "
              f"missing AUM: {len(no_aum)}, AUM to review: {len(odd)}", flush=True)

    # combined file over every factsheet processed so far (not just this run); an output whose SIF month no longer has
    # a PDF is left out, so a month a SIF's factsheet is not downloaded for is simply absent
    have = {(str(y), m, h) for y, m, h, _ in factsheet_folders(houses, quiet=True)}
    parts = [pd.read_csv(f, dtype={"scheme_code": str}) for f in sorted(OUT.glob("*/*/*/fund_details.csv"))
             if f.stat().st_size > 2  # a folder that yielded no funds writes an empty file
             and (f.parts[-4], f.parts[-3], f.parts[-2]) in have]
    parts = [p for p in parts if not p.empty]
    if parts:
        allrows = pd.concat(parts, ignore_index=True)
        allrows.to_csv(OUT / "factsheet_fund_managers.csv", index=False)
        print(f"\n{len(allrows)} rows -> {OUT / 'factsheet_fund_managers.csv'}")


if __name__ == "__main__":
    main()
