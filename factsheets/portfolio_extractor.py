"""A strategy's full portfolio (every holding and its % to net assets) read off its factsheet pages by the model.

For SIFs that publish no portfolio disclosure but print the whole portfolio in the factsheet (factsheet_portfolio.py
lists them). Same machinery as fund_extractor.py: the shared GPT-5 model, the tool-calling agent loop (run_agent)
and the factsheet's page tools; what differs is the answer (StrategyPortfolio below) and the checks submit runs:

  - each holding's % is printed on the source pages (a misread or invented figure)
  - the rows add up to every subtotal the pages print, block by block (a dropped or doubled row, a wrong sign)
  - the rows add up to the grand total (100%), with or without the futures - AMCs do both
  - the data date is in, or just before, the factsheet month

The agent fixes the answer or confirms it with a reason in notes; the third submit is taken whatever the checks say
and its remaining problems are stored. Answers are cached per strategy, keyed on the model, PORTFOLIO_PROMPT_VERSION
and the pages' content, so a re-run makes no model call for unchanged pages.
"""
import json
import re
from typing import List, Literal, Optional

from langchain_core.messages import HumanMessage, SystemMessage
from langchain_core.tools import tool
from pydantic import BaseModel, Field

import fund_extractor as fx
from fund_extractor import Pages

PORTFOLIO_PROMPT_VERSION = "sif-portfolio-3"  # bump when the prompt, schema or checks change
MAX_STEPS = 10
MAX_SUBMITS = 3
SUM_TOLERANCE = 0.3  # % points: rows printed to 2 decimals drift a little from the printed total

Kind = Literal["Equity", "REIT / InvIT", "Stock Future", "Index Future", "Stock Option", "Index Option",
               "Treasury Bill", "Government Security", "Corporate Debt", "Commercial Paper", "Certificate of Deposit",
               "Mutual Fund", "Commodity Future", "Commodity (Physical)", "TREPS / Repo", "Cash / Net Current Assets",
               "Other"]
FUTURES = {"Stock Future", "Index Future"}
OPTIONS = {"Stock Option", "Index Option"}
TOTAL_NAME = re.compile(r"\b(sub\s*-?\s*total|grand\s+total|total)\b", re.I)


class Holding(BaseModel):
    block: str = Field(description="The heading the row is printed under, as printed (e.g. 'Equity & Equity Related', "
                                   "'INVIT', 'Index / Stock Futures', 'Treasury Bills'); the same text the block's Sub Total belongs to")
    kind: Kind
    name: str = Field(description="Instrument as printed, wrapped lines joined, footnote marks (*, #, ^, $) dropped. "
                                  "For a derivative the whole contract as printed, e.g. 'NTPC Limited October 2026 Future'")
    underlying: Optional[str] = Field(default=None, description="Derivatives only: the company or index, without expiry, strike or contract words ('NTPC Limited', 'NIFTY', 'Nifty Midcap')")
    expiry: Optional[str] = Field(default=None, description="Derivatives only: expiry as YYYY-MM-DD, or YYYY-MM when only month and year are printed")
    strike: Optional[float] = Field(default=None, description="Options only: strike price")
    option_type: Optional[Literal["Call", "Put"]] = None
    industry_or_sector: Optional[str] = Field(default=None, description="Industry / sector printed for the row (or its industry heading); null if none")
    rating: Optional[str] = Field(default=None, description="Credit rating printed for a debt row; null otherwise")
    pct_to_net_assets: float = Field(description="% to net assets as printed, in percent (3.87 for 3.87%), with its sign: short / sold positions negative")


class PrintedTotal(BaseModel):
    block: Optional[str] = Field(default=None, description="The block heading this total closes, exactly as in Holding.block; null for the grand total")
    label: str = Field(description="As printed: 'Sub Total', 'Total', 'Equity Shares', 'Grand Total', ...")
    pct: float
    is_grand_total: bool = False


class StrategyPortfolio(BaseModel):
    fund_found: bool = Field(description="False if the factsheet prints no portfolio for this strategy")
    full_portfolio: bool = Field(description="True when every holding is printed; False when only top holdings / a sector or asset split is printed")
    portfolio_as_on_date: Optional[str] = Field(default=None, description="YYYY-MM-DD date the portfolio is printed as on")
    aum_month_end_cr: Optional[float] = Field(default=None, description="The strategy's month-end AUM / net assets in Rs crore, as printed")
    holdings: List[Holding] = Field(default_factory=list)
    printed_totals: List[PrintedTotal] = Field(default_factory=list, description="Every Sub Total / Total / Grand Total line of the portfolio table")
    source_pages: List[int] = Field(default_factory=list)
    notes: Optional[str] = None


class PortfolioSubmission(StrategyPortfolio):
    confirm: bool = Field(default=False, description="Only after submit reported problems: true to keep the answer because the pages really print it so; say why in notes")


SYSTEM_PROMPT = """You read the portfolio table of ONE investment strategy (the fund) from an Indian Specialized Investment Fund (SIF) monthly factsheet and return EVERY holding with its % to net assets.

You get the pages that name the fund in two forms:
- DOCLING MARKDOWN: layout-aware tables, but a sidebar can get merged into a table.
- PDF TEXT LAYER: the literal characters. Long names wrap onto the next line ('ZF Commercial Vehicle Control' then
  'Systems India Limited ...'), and page headers or footers can land in the middle of the table.
Use both and cross-check. Read the page image when you are given one.

Tools: read_page(page) for a page you were not given (the portfolio often continues on the next page),
search_factsheet(query) to find pages, submit(answer) for the final answer. submit checks the answer and replies
with problems to fix, or 'accepted'.

What to return
- holdings: one row per instrument line of the portfolio table, in printed order. Include equity, REIT / InvIT units,
  futures, options, treasury bills, bonds, mutual fund units, TREPS / repo, cash and net current assets / net
  receivables (payables) - every line that carries a % and is an instrument or a cash line.
- Leave out: Sub Total / Total / Grand Total lines (they go in printed_totals), industry or sector heading rows that
  only total the stocks under them, and the separate summary panels (market-cap split, sector allocation, asset
  allocation, 'portfolio exposure', top-10 lists) - unless the fund prints NO full table, then set full_portfolio=false.
- A table with separate columns per stock for equity %, derivative / hedge % and net % (e.g. 'HDFC Bank Ltd. 5.70% -3.76%
  1.93%'): return the equity column as an Equity row (skip it when 0.00%) and the derivative column as a Stock Future
  row of that company (underlying = the company, skip when 0.00%); never the net column. The block for these rows is
  the table's heading for that column (e.g. 'Equity Shares' and 'Derivatives').
- Keep each % exactly as printed, in percent, with its sign. A short / sold future or option is negative.
- name: the instrument as printed with wrapped lines joined. For derivatives also fill underlying, expiry, and for
  options strike and option_type.
- kind: Index Future / Index Option when the underlying is an index (NIFTY, Nifty Midcap, Bank Nifty, Sensex, ...).
  Corporate Debt for bonds / debentures / NCDs; Commercial Paper and Certificate of Deposit for those; Commodity
  Future for exchange-traded gold / silver / crude / gas contracts; Commodity (Physical) for physical gold / silver.
- Never return a total as a holding: not 'Sub Total' / 'Total' / 'Grand Total' lines, and not a heading printed with
  its block's total ('Equity Shares 73.12%', 'Total Equity Holdings 98.35%', 'BANKS 28.07% -20.66% 7.41%').
- Cash lines ('TREPS / Reverse Repo', 'Call, cash and other current asset', 'Net current assets', 'Cash & cash
  equivalents') are holdings, returned exactly once each - also when the cash block is printed as a single line
  with a total-style label and no instrument under it: then return one row named after the block (without
  'Sub Total' / 'Total') AND list its total in printed_totals.
- Each instrument line once: when the same table is printed again (a summary, a continuation repeating rows, a
  second column layout), do not add its rows a second time.
- block: the heading the row sits under, as printed, so the rows of a block add up to that block's Sub Total.
- printed_totals: every Sub Total / Total line with the block it closes, and the Grand Total (is_grand_total=true).
- aum_month_end_cr: the fund's month-end AUM / net assets in Rs crore as printed on its pages.
- portfolio_as_on_date: the date the portfolio is printed as on (YYYY-MM-DD).
If the factsheet prints only top holdings for this fund, return full_portfolio=false with what is printed.
If it prints no portfolio for this fund, return fund_found=false.
Another strategy of the same SIF may be on the same or next page: read only this fund's table.
Return null for anything not printed. Never infer a value from outside the factsheet."""


def figures(text):
    """Text layer with Indian digit grouping and % signs dropped, for 'is this number printed' checks."""
    return re.sub(r"(?<=\d),(?=\d)", "", text).replace("%", " ")


def printed(v, text_figures):
    """The value (either sign) is printed: 3.87 as '3.87', '3.870', '(3.87)', '-3.87'; 0.003 as '0.003' or '0.00'."""
    a = abs(v)
    forms = {f"{a:.2f}", f"{a:.3f}", f"{a:.4f}", f"{a:.1f}", str(a), f"{a:g}"}
    return any(re.search(rf"(?<![\d.]){re.escape(s)}(?!\d)", text_figures) for s in forms)


def sums(holdings):
    """(all rows, without futures, without futures and options)."""
    total = lambda rows: round(sum(h.pct_to_net_assets for h in rows), 4)
    return (total(holdings), total([h for h in holdings if h.kind not in FUTURES]),
            total([h for h in holdings if h.kind not in FUTURES | OPTIONS]))


def grand_total(d: StrategyPortfolio):
    g = next((t.pct for t in d.printed_totals if t.is_grand_total), None)
    return 100.0 if g is None else g


def basis(d: StrategyPortfolio):
    """
    Which rows the grand total counts, and the gap: ('all' | 'no futures' | 'no derivatives', grand - that sum).
    Arthaya's 100% leaves the futures out (their notional), DynaSIF's counts them.
    """
    g = grand_total(d)
    named = zip(("all", "no futures", "no derivatives"), sums(d.holdings))
    return min(((name, round(g - s, 4)) for name, s in named), key=lambda x: abs(x[1]))


def find_issues(d: StrategyPortfolio, pages: Pages, month: str) -> List[str]:
    if not d.fund_found:
        return []
    out = []
    if not d.holdings:
        out.append("holdings is empty although fund_found is true.")
        return out
    used = [p for p in d.source_pages if pages.valid(p)]
    if not used:
        out.append("source_pages is empty or not pages of this factsheet: list the PDF pages the table is on.")
    if not d.full_portfolio:
        return out  # only top holdings: nothing to reconcile, the driver writes no portfolio
    pictured = any(pages.is_image(p) for p in used)
    text = figures("\n".join(pages.text(p) for p in used or range(1, len(pages) + 1)))
    if not pictured:
        missing = [h for h in d.holdings if not printed(h.pct_to_net_assets, text)]
        if missing:
            out.append(f"{len(missing)} % figure(s) are not printed on source_pages {used}: "
                       + "; ".join(f"{h.name[:50]} {h.pct_to_net_assets}" for h in missing[:10])
                       + ". Give each % exactly as printed (and add the page it is on to source_pages).")
    for h in d.holdings:
        if TOTAL_NAME.search(h.name):
            out.append(f"Holding '{h.name}' {h.pct_to_net_assets} is named like a total. A total goes in printed_totals, "
                       "not holdings; only a one-line cash block is also a holding, named without 'Total'.")
    blockless = []
    for t in d.printed_totals:
        if t.is_grand_total or not t.block:
            continue
        rows = [h for h in d.holdings if h.block.strip().lower() == t.block.strip().lower()]
        if not rows:
            blockless.append(t)
            continue
        # a block total may count all its rows, only the cash / equity legs, or only the derivative legs
        # (Arthaya's industry headings total the equity column; Infinity prints a cash and a futures total)
        legs = [rows, [h for h in rows if h.kind not in FUTURES | OPTIONS], [h for h in rows if h.kind in FUTURES | OPTIONS]]
        tol = max(0.06, 0.006 * len(rows))
        if not any(r and abs(sum(h.pct_to_net_assets for h in r) - t.pct) <= tol for r in legs):
            s = round(sum(h.pct_to_net_assets for h in rows), 4)
            out.append(f"Block '{t.block}': the {len(rows)} rows add up to {s}, the printed '{t.label}' is {t.pct}. "
                       "A row is missing, doubled, has the wrong sign or belongs to another block.")
    which, gap = basis(d)
    if abs(gap) > SUM_TOLERANCE:
        a, nf, nd = sums(d.holdings)
        hint = ""
        near = [t for t in d.printed_totals if not t.is_grand_total and abs(abs(gap) - abs(t.pct)) <= 0.1]
        if near:
            t = near[0]
            hint = (f" The gap {gap:+.2f} is the printed '{t.label}' ({t.pct}): that block "
                    + ("is missing from holdings (a one-line cash block is a holding too)." if gap > 0 else
                       "is counted twice, or its total was returned as a holding."))
        out.append(f"The rows add up to {a} (all), {nf} (without futures) and {nd} (without futures and options); "
                   f"the grand total is {grand_total(d)}." + (hint or " A holding, a cash / net current assets line or "
                   "a whole block is missing or doubled, a total was returned as a holding, or a sign is wrong."))
    for t in blockless:
        out.append(f"Printed total '{t.label}' {t.pct} names block '{t.block}', which no holding has. If it totals other "
                   "blocks, set its block to null; if it is a one-line cash block, add that line as a holding.")
    seen = set()
    for h in d.holdings:
        k = (h.block.lower(), " ".join(h.name.lower().split()), h.pct_to_net_assets)
        if k in seen:
            out.append(f"'{h.name}' {h.pct_to_net_assets} appears twice in block '{h.block}': keep it once unless printed twice.")
        seen.add(k)
    for h in d.holdings:
        if h.kind in FUTURES | OPTIONS and not h.underlying:
            out.append(f"Derivative '{h.name}' has no underlying.")
    if fx.date_out_of_window(d.portfolio_as_on_date, month):
        out.append(f"portfolio_as_on_date {d.portfolio_as_on_date} is not in or just before the factsheet month ({month}).")
    if d.aum_month_end_cr is None:
        out.append("aum_month_end_cr is null: give the fund's month-end AUM / net assets in Rs crore as printed on its pages.")
    return out


def score(d: StrategyPortfolio, issues):
    """Lower is better: the gap to the grand total first, then the number of unresolved checks."""
    if not d.fund_found or not d.full_portfolio or not d.holdings:
        return (float("inf"), len(issues or []))
    return (round(abs(basis(d)[1]), 2), len(issues or []))


def cache_key(scheme_name, markdown, text, month, images):
    return fx.cache_key(f"portfolio|{PORTFOLIO_PROMPT_VERSION}|{scheme_name}", markdown, text, month, images)


def extract(scheme_name, category, markdown, text, pages: Pages, first_pages, factsheet_month, images=None,
            cache_file=None, force=False, retry_unresolved=False):
    """
    (StrategyPortfolio or None, unresolved issues, error) for one strategy. Cached in cache_file (JSON) on the
    model, prompt version and pages; an answer the model could not finish is retried on the next run, and with
    retry_unresolved one whose checks still reported problems is asked again (the model's reading varies run to run).
    """
    key = cache_key(scheme_name, markdown, text, factsheet_month, images)
    previous = None  # the cached answer a retry has to beat
    if cache_file and cache_file.exists() and not force:
        cached = json.loads(cache_file.read_text(encoding="utf-8"))
        if cached.get("key") == key and not cached.get("error"):
            previous = (StrategyPortfolio.model_validate(cached["result"]), cached.get("unresolved") or [])
            if not (retry_unresolved and cached.get("unresolved")):
                return previous[0], previous[1], None
    if fx._model_off:
        return None, None, fx._model_off

    read, images_out = set(first_pages), []
    box = {"tries": 0, "final": None, "best": None, "warned": False}

    @tool
    def read_page(page: int) -> str:
        """Read one PDF page of the factsheet (text layer; a page printed as a picture is also sent as an image)."""
        t, png = pages.read(page)
        if png:
            images_out.append(png)
        if pages.valid(page):
            read.add(page)
        return t

    @tool
    def search_factsheet(query: str) -> str:
        """Find the factsheet's pages that print every word of the query, with a snippet of each (best first)."""
        return pages.search(query, near=first_pages[0] if first_pages else None)

    @tool(args_schema=PortfolioSubmission)
    def submit(**answer) -> str:
        """Submit the strategy's portfolio. It is checked; the reply says 'accepted' or lists problems to fix."""
        s = PortfolioSubmission(**answer)
        d = StrategyPortfolio(**s.model_dump(exclude={"confirm"}))
        issues = find_issues(d, pages, factsheet_month)
        box["tries"] += 1
        if box["best"] is None or len(issues) <= len(box["best"][1]):
            box["best"] = (d, issues)
        if not issues or (s.confirm and box["warned"]) or box["tries"] >= MAX_SUBMITS:
            box["final"] = (d, issues)
            return "accepted"
        box["warned"] = True
        return ("Not accepted yet. Problems:\n" + "\n".join(f"- {i}" for i in issues)
                + "\nFix them (read more pages if needed) and submit again, or submit with confirm=true and say in "
                  "notes why the pages really print it so.")

    user = (f"Fund: {scheme_name}\nCategory: {category}\nFactsheet month (folder): {factsheet_month}\n"
            f"These are PDF pages {', '.join(map(str, first_pages))} of {len(pages)}.\n\n"
            f"<docling_markdown>\n{markdown}\n</docling_markdown>\n\n<pdf_text_layer>\n{text}\n</pdf_text_layer>")
    content = [{"type": "text", "text": user}] + [fx.png_part(p) for p in images or []]
    state = fx.run_agent([SystemMessage(SYSTEM_PROMPT), HumanMessage(content)], [read_page, search_factsheet, submit],
                         finished=lambda: box["final"] is not None, max_steps=MAX_STEPS, images=images_out)
    answer, issues = box["final"] or box["best"] or (None, None)
    error = state.get("error")
    if answer is None and not error:
        error = f"agent made {state.get('steps', 0)} model calls without submitting"
    # a retry keeps the earlier answer unless the new one reconciles at least as well (the model's reading varies)
    if previous is not None and (answer is None or error or score(answer, issues) > score(*previous)):
        return previous[0], previous[1], None
    if cache_file and answer is not None:
        cache_file.parent.mkdir(parents=True, exist_ok=True)
        cache_file.write_text(json.dumps({"key": key, "model": fx.MODEL, "result": answer.model_dump(),
                                          "unresolved": issues, "error": error,
                                          "tool_calls": fx.trace(state.get("messages", []))},
                                         indent=2, ensure_ascii=False), encoding="utf-8")
    return answer, issues or [], error
