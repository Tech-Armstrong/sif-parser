"""Portfolio workbooks built from the factsheet, for SIFs that publish no portfolio disclosure.

Usage:
    python factsheets/factsheet_portfolio.py [--year 2026] [--month SEP] [--amc Dyna] [--force | --force-llm]

Some SIFs (FROM_FACTSHEET below) publish no monthly portfolio disclosure but print every holding in their
factsheet. For each strategy of such a SIF, from data/factsheets/<YEAR>/<MON>/<Name> Factsheet.pdf:

  1. pages     the factsheet pages that name the strategy, and the page after each (tables run on)
  2. read      docling + the model read the portfolio table (portfolio_extractor.py); submit checks the rows
               against the printed subtotals and the grand total
  3. match     each equity / REIT / InvIT name, and each stock derivative's underlying, is matched to its AMFI
               company (instrument_match.py): ISIN, symbol, market cap. Matched rows are written under the AMFI name.
  4. write     "<Scheme> (from factsheet).xlsx" to data/monthly_portfolio/<YEAR>/<MON>/<SIF>/, laid out like a
               disclosure (ISIN, name,
               industry, market value, % to net assets, section headings), which sif_parser.py then parses like any
               other SIF's. Market value = % to net assets x the month-end AUM printed with the portfolio; the
               rounding gap to the printed grand total (at most MAX_ROUNDING % points) is booked to net receivables.

No workbook is written when the factsheet prints only top holdings, the rows do not add up to the grand total, or
no month-end AUM is printed; any workbook from an earlier run is then removed. A folder that already holds a real
disclosure (any other .xls* in its data/monthly_portfolio folder) is skipped: sif_parser prefers the real file anyway.

Review files, in factsheets/output/<YEAR>/<MON>/<SIF>/:
  portfolio.md           per strategy: totals, checks, every row with its printed and matched name
  portfolio_matches.csv  one row per holding (printed name, ISIN, how it was matched)
and factsheets/output/factsheet_portfolio_matches.csv over every folder. Names matched 'fuzzy' and 'unmatched' need a
look: fix them in factsheets/factsheet_instrument_alias.csv (printed_name -> isin) and re-run (no model call).
"""
import argparse
import hashlib
import re
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font

import factsheet_parser as fp
import fund_extractor as fx
import instrument_match as im
import portfolio_extractor as px
import rules
from fund_extractor import Pages

# SIFs whose factsheet prints the full portfolio and that publish no disclosure (checked by hand, Nov 2026).
# WSIF prints it in MAY 2026 only (top 10 from JUN). Not Infinity SIF: its factsheet prints the debt half of the
# portfolio as category weights only ('Debentures and Bonds 14.16%') and refers to Kotak's website for the rest.
FROM_FACTSHEET = {"Arthaya SIF", "DynaSIF", "Sapphire SIF", "WSIF"}
SUFFIX = im.sp.FROM_FACTSHEET_SUFFIX
MAX_PAGES = 6
# rounding gap booked to net receivables: rows printed to 2 decimals are each off by up to 0.005, so a gap beyond
# ~0.004 per row (capped) is a misread or missing row, not rounding: no workbook
ROUNDING_PER_ROW, ROUNDING_BASE, MAX_ROUNDING = 0.004, 0.02, 0.4
WORKERS = 4

MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
               "November", "December"]
NCA_NAME = re.compile(r"net receivable|net current (asset|liabilit)|current assets|net payable|cash / net", re.I)
TBILL_NAME = re.compile(r"t-?bill|treasury bill", re.I)
TOTAL_WORDS = re.compile(r"\b(sub\s*-?\s*total|grand\s+total|total)\b", re.I)
# sub-headings sif_parser.classify reads, for rows the factsheet prints without an ISIN
DERIV_BLOCKS = [("Stock Future", "Stock Futures"), ("Index Future", "Index Futures"),
                ("Stock Option", "Stock Options"), ("Index Option", "Index Options")]
DEBT_KINDS = {"Government Security": "Government Securities", "Treasury Bill": "Treasury Bills",
              "Commercial Paper": "Commercial Paper", "Certificate of Deposit": "Certificate of Deposit",
              "Corporate Debt": "Corporate Bonds / Debentures"}
DEBT_SUBS = list(dict.fromkeys(DEBT_KINDS.values()))
OTHER_KINDS = {"Commodity (Physical)": "Commodities Physical", "Commodity Future": "Commodity Derivatives",
               "Other": "Other Instruments"}
OTHER_SUBS = list(dict.fromkeys(OTHER_KINDS.values()))


def strategy_pages(names, others, texts, skip, amc):
    """
    Pages naming the strategy (any of its names; spacing, case and hyphens ignored: 'Long - Short'), and the page
    after each - a table runs on - unless that page names another strategy of the SIF (DynaSIF prints each
    strategy on its own two pages: the page after one strategy's is the next strategy's). At most MAX_PAGES.
    """
    names_page = lambda ns, t: any(rules.compact(n) in rules.compact(t) for n in ns) or any(
        rules.names_fund(rules.Page(fund_name=n, text=t, amc=amc)) for n in ns)
    ok = [i + 1 for i in range(len(texts)) if i + 1 not in skip]
    named = [p for p in ok if names_page(names, texts[p - 1])]
    nxt = [p + 1 for p in named if p + 1 in ok and p + 1 not in named and not names_page(others, texts[p])]
    return sorted(set(named) | set(nxt))[:MAX_PAGES]


CASH_LINE = re.compile(r"margin|cash|treps|tri-?party|repo|current asset|receivable|payable", re.I)


def fix_known_slips(d: px.StrategyPortfolio):
    """
    (portfolio, notes) with two slips the model makes on these layouts corrected in code:
    - a negative Equity row is a short derivative leg (cash equity cannot be short): Sapphire SEP prints shorts
      only in its 'Outstanding derivative exposure Long / (Short)' column ('Britannia Industries Ltd (0.43)')
    - a cash-type line listed only as a printed total (Sapphire AUG 'Margin on Derivatives 0.56'), when adding it
      closes the gap to the grand total, is a holding
    """
    notes, holdings = [], []
    for h in d.holdings:
        if h.kind == "Equity" and h.pct_to_net_assets < 0:
            h = h.model_copy(update={"kind": "Stock Future", "underlying": h.underlying or h.name})
            notes.append(f"'{h.name}' {h.pct_to_net_assets}: negative equity read as a short stock future")
        holdings.append(h)
    d = d.model_copy(update={"holdings": holdings})
    blocks = {h.block.strip().lower() for h in holdings}
    names = {" ".join(h.name.lower().split()) for h in holdings}
    for t in d.printed_totals:
        label = " ".join(t.label.lower().split())
        if t.is_grand_total or not CASH_LINE.search(t.label) or label in names or (t.block or "").strip().lower() in blocks:
            continue
        _, gap = px.basis(d)
        if abs(gap - t.pct) <= 0.05 + 0.002 * len(holdings):
            kind = "TREPS / Repo" if re.search(r"treps|tri-?party|repo", t.label, re.I) else "Cash / Net Current Assets"
            add = px.Holding(block=t.block or t.label, kind=kind, name=t.label, pct_to_net_assets=t.pct)
            d = d.model_copy(update={"holdings": d.holdings + [add]})
            notes.append(f"'{t.label}' {t.pct} was listed only as a total: added as a holding")
    return d, notes


def drop_total_rows(holdings):
    """
    (holdings, notes) without rows the model returned for a printed total: a row named like a total whose % is the
    sum of the other rows of its block, or of all other rows (Arthaya SEP 'Equity Shares (Sub Total) 73.12',
    'Grand Total 100'). A one-line cash block named '... Sub Total' (DynaSIF TREPS) has no other rows: kept.
    """
    keep, notes = [], []
    for h in holdings:
        if px.TOTAL_NAME.search(h.name):
            rest = [o for o in holdings if o is not h]
            block = [o for o in rest if o.block.strip().lower() == h.block.strip().lower()]
            sums = [sum(o.pct_to_net_assets for o in grp) for grp in (
                block, [o for o in block if o.kind not in px.FUTURES | px.OPTIONS], rest,
                [o for o in rest if o.kind not in px.FUTURES], [o for o in rest if o.kind not in px.FUTURES | px.OPTIONS])
                if grp]
            if any(abs(s - h.pct_to_net_assets) <= 0.05 + 0.003 * len(block or rest) for s in sums):
                notes.append(f"dropped '{h.name}' {h.pct_to_net_assets}: a printed total, not a holding")
                continue
        keep.append(h)
    return keep, notes


def month_end(year, mon):
    m = fp.MONTHS.index(mon) + 1
    return pd.Timestamp(year, m, 1) + pd.offsets.MonthEnd(0)


# ---------------------------------------------------------------- rows of the workbook
def contract(h: px.Holding, underlying):
    """A derivative's name as sif_parser reads it: '<Company> October 2026 Future', '<Index> 22800 Put October 2026 Option'."""
    exp = ""
    if h.expiry and re.match(r"\d{4}-\d{2}", h.expiry):
        exp = f"{MONTH_NAMES[int(h.expiry[5:7]) - 1]} {h.expiry[:4]}"
    if h.kind in px.FUTURES:
        return " ".join(filter(None, [underlying, exp, "Future"]))
    strike = f"{h.strike:g}" if h.strike is not None else ""
    return " ".join(filter(None, [underlying, strike, h.option_type or "", exp, "Option"]))


def refine_other(h: px.Holding):
    """
    A row the model left as kind 'Other' whose block or name says what it is (it labels the same rows correctly in
    other months): DynaSIF 'Exchange Traded Funds' > 'ICICI Prudential Gold ETF', 'Gold' > 'Goldten' (physical gold
    in MCX delivery units, hedged with FUTCOM_GOLDTEN futures), 'FUTCOM_GOLD_04/12/2026'.
    """
    if h.kind != "Other":
        return h
    text = f"{h.block} {h.name}"
    if re.match(r"futcom_", h.name, re.I) or re.search(r"commodity\s+(future|derivative)", h.block, re.I):
        kind = "Commodity Future"
    elif re.search(r"\betf\b|exchange traded|mutual fund|\bbees\b", text, re.I):
        kind = "Mutual Fund"
    elif re.fullmatch(r"(physical\s+)?(gold|silver)\w*(\s.*)?", h.name.strip(), re.I) or re.fullmatch(r"gold|silver|precious metals?|commodit\w*", h.block.strip(), re.I):
        kind = "Commodity (Physical)"
    else:
        return h
    return h.model_copy(update={"kind": kind})


def build_rows(d: px.StrategyPortfolio, house, year, month_num):
    """
    (sections, matches): sections = [(heading, [row dict])] in disclosure order; matches = one dict per holding for
    the review files. Equity / REIT rows and stock derivatives' underlyings are matched to the AMFI list.
    """
    aum_lakhs = d.aum_month_end_cr * 100
    blocks = {k: [] for k in ("equity", "reit", "mf", "cash")}
    deriv = {sub: [] for _, sub in DERIV_BLOCKS}
    debt = {sub: [] for sub in DEBT_SUBS}
    other = {sub: [] for sub in OTHER_SUBS}
    matches = []

    def row(name, isin, industry, pct):
        return dict(name=name, isin=isin or "", industry=industry or "", pct=round(pct, 4), mv=round(pct * aum_lakhs / 100, 4))

    for h in d.holdings:
        # a block printed only as its total ('TREPS / Reverse Repo Sub Total'): sif_parser drops rows named '... Total'
        h = h.model_copy(update={"name": TOTAL_WORDS.sub(" ", h.name).strip(" -:") or h.block})
        h = refine_other(h)
        rec = dict(block=h.block, kind=h.kind, printed_name=h.name, pct_to_net_assets=h.pct_to_net_assets,
                   industry_printed=h.industry_or_sector)
        if h.kind in ("Equity", "REIT / InvIT"):
            m = im.match(h.name, year, month_num, house)
            name = im.standard_name(m)
            r = row(name, m.isin, "" if m.matched else h.industry_or_sector, h.pct_to_net_assets)
            blocks["equity" if h.kind == "Equity" else "reit"].append(r)
            rec.update(written_name=name, isin=m.isin, company=m.company, symbol=m.symbol, market_cap=m.market_cap,
                       match=m.method, score=m.score, candidates=m.candidates)
        elif h.kind in px.FUTURES | px.OPTIONS:
            under = h.underlying or h.name
            m = None
            if h.kind.startswith("Stock"):
                m = im.match(under, year, month_num, house)
                under = im.standard_name(m)
            name = contract(h, under)
            sub = dict(DERIV_BLOCKS)[h.kind]
            deriv[sub].append(row(name, "", h.industry_or_sector if not (m and m.matched) else "", h.pct_to_net_assets))
            rec.update(written_name=name, isin=m.isin if m else None, company=m.company if m else None,
                       symbol=m.symbol if m else None, match=(m.method if m else "index"), score=m.score if m else None,
                       candidates=m.candidates if m else None)
        else:
            name = h.name
            if h.kind == "Treasury Bill" and not TBILL_NAME.search(name):
                name = f"{name} (Treasury Bill)"
            if h.kind == "Government Security" and not re.search(r"government of india|\bgoi\b|\bg-?sec", name, re.I):
                name = f"{name} (Government of India)"
            if h.kind == "TREPS / Repo" and not re.search(r"treps|tri-?party|repo|cblo|clearing corporation", name, re.I):
                name = f"TREPS - {name}"
            if h.kind == "Cash / Net Current Assets" and not NCA_NAME.search(name) and not re.search(r"cash", name, re.I):
                name = f"Net Receivables / (Payables) - {name}"
            rated = h.rating or h.industry_or_sector
            r = row(name, "", rated, h.pct_to_net_assets)
            if h.kind in DEBT_KINDS:
                debt[DEBT_KINDS[h.kind]].append(r)
            elif h.kind in OTHER_KINDS:
                other[OTHER_KINDS[h.kind]].append(r)
            else:
                blocks["mf" if h.kind == "Mutual Fund" else "cash"].append(r)
            rec.update(written_name=name, match="n/a")
        matches.append(rec)

    nested = lambda subs: [(sub, rows) for sub, rows in subs.items() if rows]
    sections = [("Equity & Equity Related", blocks["equity"]), ("Units issued by REITs & InvITs", blocks["reit"]),
                ("Debt Instruments", nested(debt)), ("Mutual Fund Units", blocks["mf"]),
                ("Derivatives", nested(deriv)), ("Cash & Cash Equivalents", blocks["cash"]), ("Others", nested(other))]
    return [(h, rows) for h, rows in sections if rows], matches


def ensure_classified(sections):
    """
    Run every row through sif_parser.classify, as the parser will, and fix the ones it cannot place - a row it
    cannot classify is dropped in parsing and the workbook no longer adds up (Sapphire 'Call,cash and other current
    asset', an unmatched equity printed without an industry). Returns notes for the review file.
    """
    notes = []
    for heading, rows in sections:
        for sub, rs in (rows if rows and isinstance(rows[0], tuple) else [(None, rows)]):
            sec = " > ".join(filter(None, [heading, sub]))
            for r in rs:
                if im.sp.classify(sec, im.sp.split_symbols(r["name"])[0], r["isin"], r["industry"])[0]:
                    continue
                before = r["name"]
                if heading == "Cash & Cash Equivalents":
                    r["name"] = f"Net Receivables / (Payables) - {r['name']}"
                elif heading in ("Equity & Equity Related", "Units issued by REITs & InvITs"):
                    r["industry"] = r["industry"] or "Unmatched (no ISIN)"
                ok = im.sp.classify(sec, im.sp.split_symbols(r["name"])[0], r["isin"], r["industry"])[0]
                notes.append(f"'{before}' written as '{r['name']}' / '{r['industry']}' so sif_parser can classify it" if ok
                             else f"'{before}' {r['pct']}: sif_parser cannot classify it and will drop it")
    return notes


def book_rounding(sections, d: px.StrategyPortfolio):
    """
    Book the gap between the printed grand total and the rows it counts to net receivables. Returns
    (basis, gap) - basis says whether the grand total counts the futures - or None when the gap is too big.
    """
    which, gap = px.basis(d)
    if abs(gap) > min(MAX_ROUNDING, ROUNDING_BASE + ROUNDING_PER_ROW * len(d.holdings)):
        return None
    if abs(gap) >= 0.0001:
        aum_lakhs = d.aum_month_end_cr * 100
        cash = next((rows for h, rows in sections if h == "Cash & Cash Equivalents"), None)
        nca = next((r for r in cash or [] if NCA_NAME.search(r["name"])), None)
        if nca is None:
            nca = dict(name="Net Receivables / (Payables)", isin="", industry="", pct=0.0, mv=0.0)
            if cash is None:
                sections.append(("Cash & Cash Equivalents", [nca]))
            else:
                cash.append(nca)
        nca["pct"] = round(nca["pct"] + gap, 4)
        nca["mv"] = round(nca["pct"] * aum_lakhs / 100, 4)
    return which, gap


def write_workbook(path, scheme_name, as_on, sections, d: px.StrategyPortfolio, pdf_name, pages):
    wb = Workbook()
    ws = wb.active
    ws.title = "Portfolio"  # sif_parser strips a sheet name printed before the title: never the scheme name
    bold = Font(bold=True)
    ws.append([scheme_name])
    ws["A1"].font = bold
    ws.append([f"Portfolio as on {MONTH_NAMES[as_on.month - 1]} {as_on.day}, {as_on.year}"])
    ws.append([f"Built from the factsheet ({pdf_name}, PDF pages {', '.join(map(str, pages))}) by "
               "factsheets/factsheet_portfolio.py: the SIF publishes no portfolio disclosure. Market value = % to net "
               "assets x month-end AUM; names matched to the AMFI market-cap list."])
    ws.append([])
    ws.append(["Name of the Instrument", "ISIN", "Industry / Rating", "Market Value (Rs. in Lakhs)", "% to Net Assets"])
    for c in ws[5]:
        c.font = bold
    for heading, rows in sections:
        ws.append([heading])
        ws.cell(ws.max_row, 1).font = bold
        if rows and isinstance(rows[0], tuple):  # (sub-heading, rows): derivatives, debt, others
            for sub, sub_rows in rows:
                ws.append([sub])
                for r in sub_rows:
                    ws.append([r["name"], r["isin"], r["industry"], r["mv"], r["pct"]])
        else:
            for r in rows:
                ws.append([r["name"], r["isin"], r["industry"], r["mv"], r["pct"]])
    ws.append(["Grand Total", "", "", round(d.aum_month_end_cr * 100, 4), px.grand_total(d)])
    ws.cell(ws.max_row, 1).font = bold
    ws.column_dimensions["A"].width = 60
    for col, w in (("B", 15), ("C", 30), ("D", 22), ("E", 16)):
        ws.column_dimensions[col].width = w
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# ---------------------------------------------------------------- one SIF month
def process_house(house, pdf_files, year, mon, master, alias, force, force_llm, retry=False):
    xlsx_dir = fp.PORTFOLIO_DATA / str(year) / mon / house  # where sif_parser.py reads disclosures
    amc = fp.amc_tag(house)
    out_dir = fp.OUT / str(year) / mon / house
    out_dir.mkdir(parents=True, exist_ok=True)
    month_num = fp.MONTHS.index(mon) + 1
    month = f"{mon.title()} {year}"

    # no index-page skip: these factsheets have none, and index_pages() takes DynaSIF's 'Benchmark / Index : BSE 500
    # TRI' line for an INDEX heading, which hid the Active Asset Allocator's first page
    pdfs = [(pdf, texts, bad, set()) for pdf, texts, bad, _ in fp.distinct_pdfs(pdf_files)]
    strategies = master[master.fund_house == house]
    jobs = []
    names_of = {str(r.scheme_code): [r.scheme_name, *alias[(alias.fund_house == house)
                                                            & (alias.scheme_code.astype(str) == str(r.scheme_code))].factsheet_name]
                for r in strategies.itertuples()}
    for r in strategies.itertuples():
        code = str(r.scheme_code)
        names = names_of[code]
        others = [n for c, ns in names_of.items() if c != code for n in ns]
        for pdf, texts, bad, skip in pdfs:
            pages = strategy_pages(names, others, texts, skip, amc)
            if pages:
                jobs.append(dict(code=code, scheme_name=r.scheme_name, category=fp.CATEGORIES.get(r.category, r.category),
                                 pdf=pdf, texts=texts, bad=bad, skip=skip, pages=pages))
                break

    # docling (not thread-safe, model load per call): every PDF's pages in one call, before the agents run
    for pdf, texts, bad, skip in pdfs:
        mine = [j for j in jobs if j["pdf"] == pdf]
        if not mine:
            continue
        source = hashlib.sha1("\f".join(texts).encode("utf-8")).hexdigest()
        mds = fp.docling_pages(pdf, sorted({q for j in mine for q in j["pages"]}), out_dir / "pages" / fp.slug(pdf.stem),
                               force, source)
        pages_obj = Pages(pdf, texts, skip=skip, pictured=bad)
        for j in mine:
            j["md"] = "\n\n".join(f"<!-- PDF page {q} -->\n\n{mds[q]}" for q in j["pages"])
            j["raw"] = "\n\n".join(f"--- PDF page {q} ---\n{texts[q - 1]}" for q in j["pages"])
            j["images"] = [pages_obj.png(q) for q in j["pages"] if pages_obj.is_image(q)]
            j["pages_obj"] = pages_obj

    def run(j):
        try:
            return px.extract(j["scheme_name"], j["category"], j["md"], j["raw"], j["pages_obj"], j["pages"], month,
                              images=j["images"], cache_file=out_dir / "llm" / f"portfolio_{j['code']}.json", force=force_llm,
                              retry_unresolved=retry)
        except Exception as e:
            if fx.auth_error(e):
                raise
            return None, None, f"{type(e).__name__}: {str(e)[:200]}"

    with ThreadPoolExecutor(WORKERS) as ex:
        results = list(ex.map(run, jobs))

    report, all_matches, written = [f"# {house} - {mon} {year}: portfolios read from the factsheet", ""], [], []
    found_codes = {j["code"] for j in jobs}
    for r in strategies.itertuples():
        if str(r.scheme_code) not in found_codes:
            report += [f"## {r.scheme_name}", "", "- No page of the factsheet names it (not launched yet?): no workbook.", ""]
            remove_stale(xlsx_dir, r.scheme_name)
    # one table read for two strategies (DynaSIF JUN-AUG before the page fix): neither is trusted
    rowset = lambda d: tuple((h.name, h.pct_to_net_assets) for h in d.holdings) if d is not None and d.holdings else None
    sets = [rowset(d) for d, _, _ in results]
    for j, (d, issues, error), rs in zip(jobs, results, sets):
        target = xlsx_dir / f"{j['scheme_name']}{SUFFIX}"
        head = [f"## {j['scheme_name']}", "", f"- Scheme code: {j['code']}; {j['pdf'].name}, pages read {j['pages']}"]
        why_not, cleaned = None, []
        if d is not None and d.holdings:
            kept, cleaned = drop_total_rows(d.holdings)
            d, fixed = fix_known_slips(d.model_copy(update={"holdings": kept}))
            cleaned += fixed
        if d is None:
            why_not = f"model gave no answer ({error})"
        elif rs is not None and sets.count(rs) > 1:
            why_not = "the same holdings were read for another strategy of this SIF (pages overlap): check the pages"
        elif not d.fund_found:
            why_not = "the factsheet prints no portfolio for it"
        elif not d.full_portfolio:
            why_not = "the factsheet prints only top holdings / an allocation, not the full portfolio"
        elif not d.aum_month_end_cr:
            why_not = "no month-end AUM printed with the portfolio (needed for market values)"
        sections, matches, rounding = [], [], None
        if why_not is None:
            sections, matches = build_rows(d, house, year, month_num)
            cleaned += ensure_classified(sections)
            rounding = book_rounding(sections, d)
            if rounding is None:
                a, nf, nd = px.sums(d.holdings)
                why_not = (f"rows add up to {a} (all) / {nf} (no futures) / {nd} (no derivatives), grand total "
                           f"{px.grand_total(d)}: further apart than rounding explains")
        if why_not:
            remove_stale(xlsx_dir, j["scheme_name"])
            report += head + [f"- **No workbook**: {why_not}."] + ([f"- Unresolved: {'; '.join(issues)}"] if issues else []) + [""]
            continue
        as_on = pd.Timestamp(d.portfolio_as_on_date) if d.portfolio_as_on_date and fx.DATE_OK.fullmatch(d.portfolio_as_on_date) \
            and len(d.portfolio_as_on_date) == 10 else month_end(year, mon)
        xlsx_dir.mkdir(parents=True, exist_ok=True)
        write_workbook(target, j["scheme_name"], as_on, sections, d, j["pdf"].name, d.source_pages or j["pages"])
        written.append(target.name)
        which, gap = rounding
        eq = [m for m in matches if m["kind"] in ("Equity", "REIT / InvIT")]
        review = [m for m in matches if m.get("match") in ("fuzzy", "unmatched")]
        report += head + [
            f"- Portfolio as on {as_on.date()}; month-end AUM Rs {d.aum_month_end_cr} cr; {len(d.holdings)} rows",
            f"- Grand total {px.grand_total(d)} counts {which} rows; rounding gap {gap:+.4f} booked to net receivables",
            f"- Equity / REIT / InvIT names matched: {sum(1 for m in eq if m['isin'])} of {len(eq)}"
            + (f"; to review: {', '.join(m['printed_name'] + ' (' + m['match'] + ')' for m in review)}" if review else ""),
            f"- Workbook: `{target.name}`",
        ] + ([f"- Cleaned: {'; '.join(cleaned)}"] if cleaned else []) \
          + ([f"- Unresolved checks: {'; '.join(issues)}"] if issues else []) + ([f"- Model notes: {d.notes}"] if d.notes else [])
        report += ["", fp.md_table(["Block", "Kind", "Printed name", "Written as", "ISIN", "Match", "%"],
                                   [(m["block"], m["kind"], m["printed_name"], m.get("written_name"), m.get("isin"),
                                     m.get("match"), m["pct_to_net_assets"]) for m in matches]), ""]
        for m in matches:
            all_matches.append(dict(year=year, month=mon, fund_house=house, scheme_code=j["code"], scheme_name=j["scheme_name"], **m))
    (out_dir / "portfolio.md").write_text("\n".join(report) + "\n", encoding="utf-8")
    pd.DataFrame(all_matches).to_csv(out_dir / "portfolio_matches.csv", index=False)
    return written, all_matches


def remove_stale(folder, scheme_name):
    f = folder / f"{scheme_name}{SUFFIX}"
    if f.exists():
        f.unlink()
        print(f"  removed earlier {f.name}")


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--year")
    ap.add_argument("--month")
    ap.add_argument("--amc", action="append", help="Only SIF folders whose name contains this text (repeatable)")
    ap.add_argument("--force", action="store_true", help="Re-run docling and the model even when cached")
    ap.add_argument("--force-llm", action="store_true", help="Re-ask the model even when an answer is cached")
    ap.add_argument("--retry-unresolved", action="store_true",
                    help="Re-ask the model only for strategies whose cached answer still fails a check")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")

    problem = fx.missing_credentials()
    if problem:
        raise SystemExit(f"{problem}: the portfolio tables are read by the model only.")
    master = pd.read_excel(fp.MASTER, sheet_name="scheme_master", dtype={"scheme_code": str})
    alias = pd.read_csv(fp.ALIAS, dtype={"scheme_code": str})

    folders = []
    for year, mon, house, pdf_files in fp.factsheet_folders(set(master.fund_house), args.year, args.month, args.amc):
        if house not in FROM_FACTSHEET:
            continue
        xdir = fp.PORTFOLIO_DATA / str(year) / mon / house
        real = [f.name for f in xdir.glob("*.xls*") if not f.name.endswith(SUFFIX) and not f.name.startswith("~$")]
        if real:
            print(f"{year} {mon} {house}: has a portfolio disclosure ({real[0]}), skipped")
            continue
        folders.append((year, mon, house, pdf_files))

    for year, mon, house, pdf_files in folders:
        print(f"{year} {mon} {house}", flush=True)
        try:
            written, matches = process_house(house, pdf_files, year, mon, master, alias, args.force, args.force or args.force_llm,
                                             args.retry_unresolved)
        except Exception as e:
            if fx.auth_error(e):
                raise SystemExit(f"Azure OpenAI rejected the request ({type(e).__name__}): check the repo-root .env")
            print(f"  ! FAILED: {type(e).__name__}: {e}")
            continue
        review = sum(1 for m in matches if m.get("match") in ("fuzzy", "unmatched"))
        print(f"  {len(written)} workbook(s) written; {len(matches)} rows, {review} name(s) to review", flush=True)

    parts = [pd.read_csv(f, dtype={"scheme_code": str}) for f in sorted(fp.OUT.glob("*/*/*/portfolio_matches.csv"))
             if f.stat().st_size > 2]
    parts = [p for p in parts if not p.empty]
    if parts:
        allrows = pd.concat(parts, ignore_index=True)
        allrows.to_csv(fp.OUT / "factsheet_portfolio_matches.csv", index=False)
        print(f"\n{len(allrows)} rows -> {fp.OUT / 'factsheet_portfolio_matches.csv'}")


if __name__ == "__main__":
    main()
