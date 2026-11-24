"""Fund manager / AUM extraction from factsheet pages: a tool-calling LangGraph agent, code rules as the fallback.

Ported from historical-mf (excel/factsheets/fund_extractor.py) for SIF factsheets. What differs: the prompts speak
of SIF investment strategies, and a page whose text layer is unreadable (Magnum DEC 2025 prints its labels in a
font whose text layer is shifted: '0DJQXP' for 'Magnum') goes to the model as a picture as well as text.

factsheet_parser.py finds each fund's first page(s) and converts them with docling; this module turns them into a
validated FundDetails object:

    fund agent:  START -> think -> act -> think -> ... -> END         (run_agent), for every fund
    rules                                                            only when the agent gives no answer

`rules` runs our own extractors (the rules/ package, scored against the model's answers by rules_eval.py). The
agent does not see their values; they answer for a fund only when the model is off (no credentials), fails or
never submits, and are otherwise stored beside the agent's answer for comparison.

The fund agent decides for itself where to look and what to trust. Its tools:
    read_page(page)          any PDF page of the factsheet (a page printed as a picture comes as an image)
    search_factsheet(query)  pages printing the query's words - the fund's real page when the one it was given is
                             another fund's, a manager's start date printed on a separate page (Kotak), an AUM
    factsheet_date()         the data month-end of the whole factsheet (ABSL prints 'AUM as on last day')
    submit(answer)           its final answer. Plain-code checks (no tokens) run here and come back as feedback:
                             an AUM date outside the factsheet month, an AUM label next to a figure it left null,
                             month-end vs average AUM far apart, a manager the page says has left, source pages that
                             do not name the fund. The agent fixes the answer or confirms it with a reason in notes.

The agent's answer stands, including the fields it leaves null, except where reconcile() finds that a rule or the
page itself settles a manager or a date.

The model is a GPT-5 deployment on Azure OpenAI, called through LangChain (langchain-openai) on Azure's v1 API.
Its settings are API_KEY, endpoint and deployment_name in the repo-root .env (see .env.example). Without them the
run carries on with the rules alone. FACTSHEET_LLM_RPM caps requests per minute across all threads (default 60;
keep it under the deployment's RPM limit); FACTSHEET_LLM_REASONING_EFFORT sets GPT-5's reasoning effort (default low).
"""
import base64
import hashlib
import io
import json
import os
import re
import threading
from typing import Callable, List, Optional, TypedDict
from urllib.parse import urlsplit

import openai
import pypdfium2 as pdfium
from dotenv import load_dotenv
from langchain_core.exceptions import ModelAuthenticationError, ModelPermissionDeniedError, ModelRateLimitError
from langchain_core.messages import HumanMessage, SystemMessage, ToolMessage
from langchain_core.rate_limiters import InMemoryRateLimiter
from langchain_core.tools import tool
from langchain_openai import ChatOpenAI
from langgraph.graph import END, START, StateGraph
from pydantic import BaseModel, Field

import rules
from rules import managers
from rules.common import DATE, parse_date

load_dotenv()  # searches upward, so the repo-root .env is found

MODEL = (os.getenv("deployment_name") or "").strip()  # the Azure deployment (GPT-5); also part of the cache key
API_KEY = (os.getenv("API_KEY") or "").strip()
ENDPOINT = (os.getenv("endpoint") or "").strip()
RPM = float(os.getenv("FACTSHEET_LLM_RPM", "60"))  # requests per minute; keep under the deployment's limit
REASONING_EFFORT = os.getenv("FACTSHEET_LLM_REASONING_EFFORT", "low")  # minimal / low / medium / high
RATE_LIMITED = (ModelRateLimitError, openai.RateLimitError)
# bump when a prompt, a tool or the schema changes, so cached answers are redone (not bumped when the rules hint
# was dropped and the renamed-fund lines and submit's printed-figure / page-name checks were added: the earlier
# answers were replayed against the new checks and only the ones they flag were redone)
PROMPT_VERSION = "sif-1"
MAX_STEPS = 8         # model calls one fund agent may make
MAX_SUBMITS = 3       # the third submit is accepted whatever the checks say (their problems are kept)
IMAGE_PAGE_CHARS = 300  # a page with less text than this is a picture (Baroda JAN Multi Asset): read it as an image
READ_CHARS = 6000       # text a read_page tool call returns; fund pages are 2-5k characters
FUND_LABELS = re.compile(r"fund\s*manager|\bA?AUM\b|assets\s+under\s+management|fund\s+size", re.I)

PDFIUM_LOCK = threading.Lock()


class Pages:
    """
    One factsheet as the agent's tools see it: every PDF page's text layer, page pictures on demand, page search.
    Shared by fund agents running in parallel, so rendering (pdfium is not thread-safe) and the factsheet-level date
    (one model call per factsheet) are behind locks.
    """

    def __init__(self, pdf, texts, skip=(), date_cache=None, force=False, pictured=()):
        self.pdf, self.texts, self.skip = pdf, [t.replace("\r\n", "\n") for t in texts], set(skip)
        self.pictured = set(pictured)  # pages whose text layer was unreadable: the model also gets their picture
        self._compact = [rules.compact(t) for t in self.texts]
        self._date_cache, self._force = date_cache, force
        self._date, self._date_lock = None, threading.Lock()

    def __len__(self):
        return len(self.texts)

    def text(self, p):
        return self.texts[p - 1]

    def valid(self, p):
        return isinstance(p, int) and 1 <= p <= len(self.texts)

    def is_image(self, p):
        return p in self.pictured or len(self.texts[p - 1].strip()) < IMAGE_PAGE_CHARS

    def png(self, p, scale=2.0):
        """PNG bytes of one PDF page (~144 dpi)."""
        with PDFIUM_LOCK:
            doc = pdfium.PdfDocument(str(self.pdf))
            try:
                img = doc[p - 1].render(scale=scale).to_pil()
            finally:
                doc.close()
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    def names(self, p, name):
        """The page prints this name (spacing, case, '&'/'and', 'Midcap'/'Mid Cap' ignored)."""
        c = rules.compact(name or "")
        return bool(c) and c in self._compact[p - 1]

    def search(self, query, limit=8, near=None):
        """
        Pages printing every word of the query, as '<page>: [fund page] ...snippet...' lines; the best first
        (most hits, then nearest `near`). Index pages are left out.
        """
        words = [w for w in (rules.compact(w) for w in query.split()) if w]
        if not words:
            return "empty query"
        hits = []
        for i, c in enumerate(self._compact):
            if i + 1 in self.skip or not all(w in c for w in words):
                continue
            hits.append((-sum(c.count(w) for w in words), abs(i + 1 - near) if near else 0, i + 1))
        if not hits:
            return f"no page prints all of: {query}"
        out = [f"{len(hits)} page(s) print '{query}'" + (f"; first {limit}" if len(hits) > limit else "") + ":"]
        for _, _, p in sorted(hits)[:limit]:
            t = self.texts[p - 1]
            m = re.search(r"\s*".join(re.escape(ch) for ch in query.split()[0]), t, re.I)
            at = m.start() if m else 0
            snippet = " ".join(t[max(0, at - 120):at + 220].split())
            out.append(f"- page {p}{' [has fund manager / AUM labels]' if FUND_LABELS.search(t) else ''}: ...{snippet}...")
        return "\n".join(out)

    def read(self, p):
        """(text for the model, png or None): the page's text layer, and its picture when it has almost no text."""
        if not self.valid(p):
            return f"no page {p}: the factsheet has pages 1-{len(self.texts)}", None
        t = self.texts[p - 1]
        if self.is_image(p):
            how = "its text layer was unreadable and is repaired below" if p in self.pictured else "printed as a picture"
            return f"--- PDF page {p} --- ({how}; attached as an image)\n{t}", self.png(p)
        more = f"\n[... {len(t) - READ_CHARS} more characters cut]" if len(t) > READ_CHARS else ""
        return f"--- PDF page {p} ---\n{t[:READ_CHARS]}{more}", None

    def factsheet_date(self):
        """The data month-end of the whole factsheet (the module's factsheet_date), asked once and shared."""
        with self._date_lock:
            if self._date is None:
                opening = "\n\n".join(f"--- PDF page {i + 1} ---\n{self.texts[i][:2500]}" for i in range(min(10, len(self.texts))))
                self._date = factsheet_date(opening, cache_file=self._date_cache, force=self._force)
            return self._date


class FundManager(BaseModel):
    name: str = Field(description="Person's name only: no Mr./Ms./Dr., footnote marks (*, ^, #, $) or role")
    managing_since: Optional[str] = Field(default=None, description="YYYY-MM-DD, or YYYY-MM when only month and year are printed; null if not printed")
    managing_since_as_printed: Optional[str] = Field(default=None, description="The date exactly as printed, e.g. '01-Jul-26', 'November 2023'")
    role: Optional[str] = Field(default=None, description="Role if printed, e.g. 'Equity portion', 'Debt', 'Overseas investments'; else null")


class FundDetails(BaseModel):
    fund_found: bool = Field(description="False if the factsheet has no page carrying this fund's own details")
    fund_name_on_page: Optional[str] = None
    fund_managers: List[FundManager]
    aum_month_end_cr: Optional[float] = Field(default=None, description="Month-end AUM in Rs crore")
    aum_monthly_average_cr: Optional[float] = Field(default=None, description="Monthly average AUM (AAUM) in Rs crore")
    aum_as_on_date: Optional[str] = Field(default=None, description="YYYY-MM-DD date printed for the AUM (else for the NAV / portfolio); null if none printed")
    source_pages: List[int] = Field(default_factory=list, description="PDF page numbers the managers and AUM were read from")
    notes: Optional[str] = Field(default=None, description="Anything ambiguous, e.g. a manager appointed after the AUM date; else null")


class Submission(FundDetails):
    confirm: bool = Field(default=False, description="Only after submit reported problems: true to keep the flagged values because the pages really print them so; say why in notes")


SYSTEM_PROMPT = """You read pages from Indian Specialized Investment Fund (SIF) monthly factsheets and pull out, for ONE named investment strategy (called the fund below), its fund managers and its AUM.

A SIF is run by a mutual fund house (Altiva SIF by Edelweiss, Magnum SIF by SBI, iSIF by ICICI Prudential, ...) and
its factsheet calls each fund an "Investment Strategy". The same people often manage the house's mutual fund
schemes, and the factsheet lists those schemes and their performance: they are not this fund.

You get the fund's first pages in two forms:
- DOCLING MARKDOWN: layout-aware, keeps tables, but sometimes merges a sidebar into a table so that labels and values end up in different cells or rows.
- PDF TEXT LAYER: the literal characters on the page; reading order can jump between columns, but values sit next to their labels more often.
Use both and cross-check them. When they disagree, prefer the reading where every label gets a sensible value.
Some pages are printed as a picture with almost no text layer, or their text layer is unreadable; then you also get
the page as an IMAGE - read it, and trust it over garbled text.

How you work
You have tools and decide yourself where to look.
- read_page(page): any PDF page's text layer. A fund's details often continue on the next page.
- search_factsheet(query): pages printing all the query's words, with a snippet. Use it when the pages you were
  given are another fund's (search the fund's distinguishing name), when a manager's start date is printed
  elsewhere (search 'managing the fund since' or the manager's name), or to find where the AUM is printed.
- factsheet_date(): the month-end the whole factsheet's data is as on, read from its opening pages. Use it for
  aum_as_on_date only when the fund's own pages print no AUM / NAV date, and say so in notes.
- submit(answer): your final answer, with source_pages = the PDF pages you read it from. It is checked; if it
  reports problems, fix them (read or search more if needed), or submit again with confirm=true and explain in
  notes why the pages really print it so.
If the factsheet has no page of this fund's own, submit fund_found=false with no managers and no AUM.
A factsheet may print the fund's name without its brand ('Hybrid Long Short Fund' for Apex Hybrid Long-Short Fund)
or with 'Long Short' / 'ExTop 100' spelled differently; a page of the same SIF whose printed category matches the
given Category is this fund's page. One factsheet may cover several strategies of the SIF: read only this one's.
Most funds need no tool besides submit; read or search only for what the given pages lack.

Fund managers
- List every person the page names as a fund manager of THIS fund, including co-managers and dedicated managers for overseas, debt or equity portions.
- A name may wrap onto the next line ('Ms. Shibani Sircar' then 'Kurian'): give the full name; a lone surname is
  not a separate manager.
- Pages often mention other schemes the same person manages (performance tables, "manages N schemes", "Performance
  of other schemes managed by the fund managers") - do not list managers of other schemes.
- An assistant fund manager named for this fund (also in a footnote: 'Mr. X has been appointed as the Assistant Fund
  Manager in <this fund> with effect from <date>') is listed, with role 'Assistant Fund Manager' and that date.
- managing_since is the date that person started managing THIS fund. Return YYYY-MM-DD, or YYYY-MM when only month and year are printed. A two-digit year means 20YY ("01-Jul-26" is 2026-07-01).
- A footnote may say a manager was appointed with effect from a date after the AUM date; still list them as printed and mention it in notes.
- A manager the page says has ceased to manage the fund (or managed it 'till' a date) on or before the AUM date is not listed.

AUM
- Month-end AUM may be labelled AUM, Total AUM, Month end AUM, Closing AUM, AUM as on <date>, or Assets Under Management.
- Monthly average AUM may be labelled AAUM, Monthly Average AUM, Monthly AVG AUM or Average AUM.
- Return both in Rs crore. Convert if the page states lakhs or another unit. The rupee sign may come through as "v", "`" or "Rs".
- Do not confuse AUM with NAV, expense ratio, portfolio turnover or holding values.
- aum_as_on_date is the date PRINTED on the page for the AUM; if the AUM line has none ("AUM as on last day"),
  the date printed for the NAV / portfolio ("NAV as on January 30, 2026"). Copy the day exactly as printed
  (30 stays 30). Return YYYY-MM-DD. Never derive it from the factsheet folder month, and never use other dates
  (risk-free rate / MIBOR, allotment, inception, manager dates, performance tables).
- If managers and their managing-since dates cannot be paired one to one (e.g. three managers, two dates), return
  managing_since null for the ones you cannot pair and say so in notes.

Return null for anything not printed in the factsheet. Never infer a value from outside it."""


def missing_credentials():
    """Message when the Azure OpenAI settings are incomplete (checked before any docling work starts)."""
    missing = [k for k, v in (("API_KEY", API_KEY), ("endpoint", ENDPOINT), ("deployment_name", MODEL)) if not v]
    if not missing:
        return None
    return f"Azure OpenAI not configured: {', '.join(missing)} missing in the repo-root .env (see .env.example)"


def base_url():
    """
    Azure's v1 API root for the resource in `endpoint`, which may be just the resource
    (https://<name>.openai.azure.com/) or a full target URI copied from the portal (.../openai/deployments/...).
    """
    u = urlsplit(ENDPOINT if "://" in ENDPOINT else "https://" + ENDPOINT)
    return f"{u.scheme}://{u.netloc}/openai/v1/"


def auth_error(e):
    """A rejected key: no retry fixes it, so the run stops instead of failing every fund in turn."""
    return (isinstance(e, (ModelAuthenticationError, ModelPermissionDeniedError))
            or 401 in (getattr(e, "status_code", None), getattr(e, "code", None))
            or 403 in (getattr(e, "status_code", None), getattr(e, "code", None))
            or "API key" in str(e))


_model = None
_model_off = None  # why the model is off for this run (not configured); funds then keep the rules' answer
_lock = threading.Lock()


def model():
    """
    One shared chat model; the lock stops parallel graph branches each building one. Its rate limiter spaces
    requests to RPM per minute across all of them; the SDK retries 429 / 5xx with back-off (honouring Azure's
    retry-after) and the timeout keeps a stalled model from hanging the run. GPT-5 takes no temperature.
    """
    global _model
    with _lock:
        if _model is None:
            _model = ChatOpenAI(
                model=MODEL, api_key=API_KEY, base_url=base_url(), reasoning_effort=REASONING_EFFORT,
                timeout=180, max_retries=4,
                rate_limiter=InMemoryRateLimiter(requests_per_second=RPM / 60, check_every_n_seconds=0.1, max_bucket_size=1))
    return _model


def disable_model(reason):
    global _model_off
    with _lock:
        if _model_off is None:
            print(f"  ! model off for the rest of this run, using code rules only: {reason}", flush=True)
            _model_off = reason


def cache_key(fund_name, markdown, text, factsheet_month, images):
    h = hashlib.sha256()
    for part in (MODEL, PROMPT_VERSION, fund_name, markdown, text, factsheet_month or "", *(images or [])):
        h.update(part if isinstance(part, bytes) else part.encode("utf-8"))
        h.update(b"\0")
    return h.hexdigest()


def read_cache(cache_file, key, force):
    """
    The stored entry for unchanged pages, else None. Only answers the model took part in are reused: a rules-only
    answer is redone (free, and a changed rule then takes effect), as is one the model could not finish (retried).
    """
    if cache_file and cache_file.exists() and not force:
        cached = json.loads(cache_file.read_text(encoding="utf-8"))
        if cached.get("key") == key and cached.get("source", "model") != "rules" and not cached.get("model_pending"):
            return cached
    return None


def read_model_answer(cache_file, key, force):
    """
    (the agent's own stored answer, its unresolved issues) for unchanged pages, else (None, None). Only the
    agent's answer is reused - before rule values were merged in - so the rules and lookups run again on every
    run for free.
    """
    cached = read_cache(cache_file, key, force)
    if cached is None:
        return None, None
    return FundDetails.model_validate(cached.get("model_result") or cached["result"]), cached.get("unresolved")


def write_cache(cache_file, key, result, **extra):
    if cache_file:
        cache_file.parent.mkdir(parents=True, exist_ok=True)
        model_used = {} if extra.get("source") == "rules" else {"model": MODEL}  # only when the model took part
        cache_file.write_text(json.dumps({"key": key, **model_used, "result": result.model_dump(), **extra},
                                         indent=2, ensure_ascii=False), encoding="utf-8")


# ---------------------------------------------------------------- tool-calling agent loop
class AgentState(TypedDict, total=False):
    messages: list
    steps: int    # model calls so far
    error: str    # why the model could not be asked


def png_part(png):
    return {"type": "image", "base64": base64.b64encode(png).decode("ascii"), "mime_type": "image/png"}


def run_agent(messages, tools, finished: Callable[[], bool], max_steps, images: list = None) -> AgentState:
    """
    The LangGraph loop the agent runs:    START -> think --(tool calls)--> act --(not finished)--> think
                                                       \\--(model unavailable)--> END      \\--> END

    think: the model picks its next tool call(s). tool_choice='any' makes it always call one, so it ends by calling
           its finishing tool (submit), never by chatting.
    act:   runs them and returns the results; a tool can append page pictures to `images`, which go to the model as
           one message after the tool results. Ends once finished() says so or after max_steps model calls.
    """
    by_name = {t.name: t for t in tools}
    llm = model().bind_tools(tools, tool_choice="any")

    def think(s: AgentState):
        if _model_off:
            return {"error": _model_off}
        try:
            ai = llm.invoke(s["messages"])
        except RATE_LIMITED as e:  # Azure limits are per minute: this fund keeps the rules' answer and is re-asked next run
            return {"error": f"rate limited: {str(e)[:200]}"}
        except Exception as e:
            if auth_error(e):
                raise
            return {"error": f"{type(e).__name__}: {str(e)[:200]}"}
        return {"messages": s["messages"] + [ai], "steps": s.get("steps", 0) + 1}

    def act(s: AgentState):
        out = []
        for call in s["messages"][-1].tool_calls:
            t = by_name.get(call["name"])
            try:
                result = t.invoke(call["args"]) if t else f"unknown tool {call['name']}; use one of {sorted(by_name)}"
            except Exception as e:  # a bad argument is the model's to fix, not a crash
                result = f"tool error: {type(e).__name__}: {str(e)[:300]}"
            out.append(ToolMessage(str(result), tool_call_id=call["id"], name=call["name"]))
        if images:
            out.append(HumanMessage([{"type": "text", "text": "Pages printed as pictures, from your read_page calls:"}]
                                    + [png_part(p) for p in images]))
            images.clear()
        return {"messages": s["messages"] + out}

    def after_think(s: AgentState):
        return END if s.get("error") or not s["messages"][-1].tool_calls else "act"

    def after_act(s: AgentState):
        return END if finished() or s.get("steps", 0) >= max_steps else "think"

    g = StateGraph(AgentState)
    g.add_node("think", think)
    g.add_node("act", act)
    g.add_edge(START, "think")
    g.add_conditional_edges("think", after_think, ["act", END])
    g.add_conditional_edges("act", after_act, ["think", END])
    return g.compile().invoke({"messages": messages, "steps": 0}, {"recursion_limit": 2 * max_steps + 4})


def trace(messages):
    """The agent's tool calls, for the cache file: what it looked at and why it ended where it did."""
    out = []
    for m in messages:
        for c in getattr(m, "tool_calls", None) or []:
            args = {k: v for k, v in c["args"].items() if k in ("page", "query", "confirm", "fund_found", "source_pages")}
            out.append(f"{c['name']}({json.dumps(args, ensure_ascii=False)})")
    return out


# ---------------------------------------------------------------- checks
DATE_OK = re.compile(r"\d{4}-\d{2}(?:-\d{2})?")
MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]


def date_out_of_window(date, factsheet_month):
    """
    An AUM date the factsheet cannot be reporting: a factsheet carries the data of its own month or the one or two
    before (ABSL's JAN factsheet is as on 31-Dec), never a year-old date from a performance table (Motilal JAN:
    'NAV as on 30-January-2025' was read as the AUM date).
    """
    m = re.fullmatch(r"(\w{3})\w*\s+(\d{4})", (factsheet_month or "").strip())
    if not date or not m or m.group(1).lower() not in MONTHS or not DATE_OK.fullmatch(date):
        return False
    folder = int(m.group(2)) * 12 + MONTHS.index(m.group(1).lower())
    seen = int(date[:4]) * 12 + int(date[5:7]) - 1
    return not folder - 2 <= seen <= folder


AUM_LABEL = re.compile(r"(?:\bA?AUM\b|Assets\s+Under\s+Management)[^\n]{0,80}?\d[\d,]*\.\d", re.I)
MANAGER_LABEL = re.compile(r"fund\s*manager", re.I)
SCALARS = ("aum_month_end_cr", "aum_monthly_average_cr", "aum_as_on_date")


def find_issues(d: FundDetails, text: str, month: str = None) -> List[str]:
    """Problems with an answer that a second look at the pages can fix (run before clean() drops bad dates)."""
    if not d.fund_found:
        return []
    out = []
    if date_out_of_window(d.aum_as_on_date, month):
        out.append(f"aum_as_on_date {d.aum_as_on_date} is not in or just before the factsheet month ({month}): it is "
                   "probably a performance-table, inception or other date. Use the date printed for the AUM itself.")
    if d.aum_month_end_cr is None and AUM_LABEL.search(text):
        out.append("aum_month_end_cr is null, but the text layer prints an AUM label with a figure next to it. "
                   "If that figure is this fund's month-end AUM, return it in Rs crore.")
    if not d.fund_managers and MANAGER_LABEL.search(text):
        out.append("fund_managers is empty, but the pages print a 'Fund Manager' label. List this fund's managers.")
    if d.aum_month_end_cr and d.aum_monthly_average_cr:
        ratio = d.aum_month_end_cr / d.aum_monthly_average_cr
        if not 0.7 <= ratio <= 1.43:
            out.append(f"Month-end AUM {d.aum_month_end_cr} and monthly average AUM {d.aum_monthly_average_cr} differ by a "
                       f"factor of {ratio:.2f}; they rarely differ by more than 30%. Re-check both labels and units (crore vs "
                       "lakh) and that neither is a NAV, holding value or another scheme's figure. Keep them if they are as printed.")
    # a figure the pages do not print: Kotak JUL Contra 'AUM: `5,449.89 crs' came back as 54.4989
    figures = re.sub(r"(?<=\d),(?=\d)", "", text)  # '1,33,372.11' -> '133372.11'
    for f in ("aum_month_end_cr", "aum_monthly_average_cr"):
        v = getattr(d, f)
        if v is not None and len(text.strip()) > 200 and not any(
                re.search(rf"(?<![\d.]){re.escape(s)}(?!\d)", figures)
                for s in {str(v), f"{v:.3f}", f"{v:.2f}", f"{v:.1f}", f"{v:.0f}"}):  # Mirae prints '2,169.200'
            out.append(f"{f} {v} is not printed on the source pages. Give the figure as printed (in Rs crore); if you "
                       "converted it from lakhs or another unit, confirm and say so in notes.")
    for m in d.fund_managers:
        if m.managing_since and not DATE_OK.fullmatch(m.managing_since):
            out.append(f"managing_since '{m.managing_since}' for {m.name} is not YYYY-MM-DD or YYYY-MM.")
    if d.aum_as_on_date and not DATE_OK.fullmatch(d.aum_as_on_date):
        out.append(f"aum_as_on_date '{d.aum_as_on_date}' is not YYYY-MM-DD.")
    return out


def rules_complete(found: rules.Found) -> bool:
    """The rules found everything a fund needs, or that the page does not print it (scored by rules_eval.py)."""
    listed = found.values.get("fund_managers") or []
    return (found.has("aum_month_end_cr") and found.has("aum_as_on_date")
            and bool(listed) and all(m.get("managing_since") or m.get("since_not_printed") for m in listed))


def from_rules(values) -> FundDetails:
    d = FundDetails(fund_found=True, fund_managers=[FundManager(**m) for m in values.get("fund_managers") or []])
    for f in SCALARS:
        setattr(d, f, values.get(f))
    return d


def _name_rx(name):
    return r"\s+".join(re.escape(w) for w in name.split())


def ceased(name, text):
    """The sentence saying this manager has left the fund, and the date in it: (sentence, 'YYYY-MM-DD' or None)."""
    s = re.search(rf"{_name_rx(name)}[^.]{{0,40}}?(?:has\s+)?(?:ceased\s+to\s+be|was\s+co[\s-]*managing[^.]{{0,60}}?till)"
                  rf"(?P<rest>[^.]{{0,160}})", text, re.I)
    if not s:
        return None
    return " ".join(s.group(0).split())[:160], parse_date((re.search(DATE, s.group("rest")) or [None])[0])


def drop_ceased(d: FundDetails, text: str, factsheet_month=None) -> List[str]:
    """
    Managers the page itself says have left: 'Amit Ganatra has ceased to be Fund Manager with effect from close of
    business hours of January 27, 2026' (Invesco), 'Mr Milind Agrawal was co managing the Equity portion of Fund
    till 30th Apr 2026' (SBI). Dropped when that date is on or before the AUM date (no date printed: dropped too).
    Used when the rules answer alone; the agent gets the same finding from submit and decides.
    """
    notes, keep = [], []
    for m in d.fund_managers:
        c = ceased(m.name, text)
        if c:
            sentence, left = c
            cutoff = d.aum_as_on_date or (None if date_out_of_window(left, factsheet_month) else left)
            if not left or not cutoff or left[:len(cutoff)] <= cutoff:
                notes.append(f"{m.name} left the fund ({sentence[:120]}): not listed")
                continue
        keep.append(m)
    d.fund_managers = keep
    return notes


def fill_dates(d: FundDetails, dates: dict) -> List[str]:
    """Start dates printed elsewhere in the factsheet for managers whose date is still missing."""
    notes = []
    # by full name, else by first and last name: 'Shibani Kurian' on one page is 'Shibani Sircar Kurian' on another
    by_name = {rules.compact(n): v for n, v in (dates or {}).items()}
    by_ends = {}
    for n, v in (dates or {}).items():
        w = n.split()
        by_ends.setdefault((rules.compact(w[0]), rules.compact(w[-1])), set()).add(v)
    for m in d.fund_managers:
        w = m.name.split()
        ends = by_ends.get((rules.compact(w[0]), rules.compact(w[-1]))) if w else None
        printed = by_name.get(rules.compact(m.name)) or (next(iter(ends)) if ends and len(ends) == 1 else None)
        if not m.managing_since and printed and parse_date(printed):
            m.managing_since, m.managing_since_as_printed = parse_date(printed), printed
            notes.append(f"managing_since of {m.name} from the factsheet's fund-manager pages")
    return notes


def normal_text(text):
    """Text layer without the quirks that break name and date patterns: rules.clean_text, and '4\\nth April 2008'."""
    return re.sub(r"(\d)\s*\n\s*(st|nd|rd|th)\b", r"\1\2", rules.clean_text(text))


def same_person(a, b):
    """'Shibani Kurian' / 'Shibani Sircar Kurian', 'Bharath S' / 'S Bharath'; case and punctuation ignored."""
    x, y = [rules.compact(w) for w in a.split()], [rules.compact(w) for w in b.split()]
    return bool(x and y) and (sorted(x) == sorted(y) or (x[0], x[-1]) == (y[0], y[-1]))


def reconcile(d: FundDetails, found: rules.Found, text, first_text, amc, footnotes) -> List[str]:
    """
    Where the agent's answer differs from what a rule or the page settles, on the fund's own pages (cases seen):
    - dropped: a manager the rule's complete list lacks and the fund's first page does not print (Edelweiss AUG Large
      Cap: a footnote on the next page, about another scheme, named four people)
    - added: a manager the rule's complete list has (Tata's assistant fund manager), or a footnote names (ICICI
      'overseas investment is managed by', HDFC 'Overseas Investments: Mr. X (since ...)'), unless they have left
    - dates: the date printed on the manager's own line wins (Edelweiss Amit Vora: '23-Jun-24' came back as 2023), then
      the rule's for a date left null, then the inception date for 'since inception'; a day the model left out of a
      date it quoted in full is restored
    """
    notes = []
    ruled = found.values.get("fund_managers") or []
    extra = managers.footnote_managers(footnotes)
    first = rules.compact(first_text)
    keep = []
    for m in d.fund_managers:
        if ruled and not any(same_person(m.name, r["name"]) for r in ruled + extra) and rules.compact(m.name) not in first:
            notes.append(f"{m.name} dropped: not in the fund's manager block")
        else:
            keep.append(m)
    d.fund_managers = keep
    for r in ruled + extra:
        if not any(same_person(m.name, r["name"]) for m in d.fund_managers) and not ceased(r["name"], footnotes):
            d.fund_managers.append(FundManager(**r))
            notes.append(f"{r['name']} added from the fund's pages")
    page = rules.Page(fund_name="", text=text, amc=amc)
    for m in d.fund_managers:
        old = m.managing_since
        beside = managers.date_beside_name(text, m.name)
        rule = next((r for r in ruled + extra if same_person(m.name, r["name"]) and r.get("managing_since")), None)
        if beside and parse_date(beside) and not (old or "").startswith(parse_date(beside)):
            m.managing_since, m.managing_since_as_printed = parse_date(beside), beside
        elif not old and rule:
            m.managing_since, m.managing_since_as_printed = rule["managing_since"], rule["managing_since_as_printed"]
        elif not old and managers.since_inception(text, m.name) and managers.inception(page):
            m.managing_since, m.managing_since_as_printed = parse_date(managers.inception(page)), "Since Inception"
        elif old and len(old) == 7 and m.managing_since_as_printed and (parse_date(m.managing_since_as_printed) or "").startswith(old):
            m.managing_since = parse_date(m.managing_since_as_printed)
        if old and m.managing_since != old:
            notes.append(f"managing_since of {m.name} {old} -> {m.managing_since} as printed")
    return notes


def clean(d: FundDetails) -> FundDetails:
    """Drop malformed dates and duplicate names rather than pass them downstream."""
    seen, kept = set(), []
    for m in d.fund_managers:
        m.name = re.sub(r"^[IVX]{1,4}\.\s*", "", m.name)  # Sundaram OCT numbers its managers: 'I. Ashwin Jain'
        key = " ".join(m.name.lower().split())
        if not key or key in seen:
            continue
        seen.add(key)
        if m.managing_since and not DATE_OK.fullmatch(m.managing_since):
            m.managing_since = None
        kept.append(m)
    d.fund_managers = kept
    if d.aum_as_on_date and not DATE_OK.fullmatch(d.aum_as_on_date):
        d.aum_as_on_date = None
    return d


def blank(d: FundDetails):
    """Another fund's figures must not be reported as this fund's."""
    d.fund_managers, d.aum_month_end_cr, d.aum_monthly_average_cr, d.aum_as_on_date = [], None, None, None


# ---------------------------------------------------------------- fund agent
def fund_agent(fund_name, amc, pages: Pages, first_pages, month, messages):
    """
    Runs the agent for one fund. Returns (answer, unresolved issues, tool-call trace, error); answer is None when
    the model could not be asked or never submitted.
    """
    read, images = set(first_pages), []
    box = {"tries": 0, "final": None, "best": None, "warned": False}

    def review(d: FundDetails) -> List[str]:
        if not d.fund_found:
            return []
        used = [p for p in d.source_pages if pages.valid(p)] or list(first_pages)
        text = "\n\n".join(pages.text(p) for p in used)
        out = find_issues(d, text, month)
        pictured = any(pages.is_image(p) for p in used)
        if pictured:  # a figure read from a picture is not in any text layer
            out = [i for i in out if "is not printed on the source pages" not in i]
        bad = [p for p in d.source_pages if not pages.valid(p)]
        if bad:
            out.append(f"source_pages {bad} are not pages of this factsheet (1-{len(pages)}).")
        if not d.source_pages:
            out.append("source_pages is empty: list the PDF pages you read the managers and AUM from.")
        elif not pictured and not rules.names_fund(rules.Page(fund_name=fund_name, text=text, amc=amc)):
            out.append(f"None of source_pages {used} prints this fund's name ({fund_name}): they may be another fund's. "
                       "Search for the fund's own page, or submit fund_found=false if it has none.")
        else:
            # Baroda JAN Multi Asset: its own page is a picture (32); the agent took the AUM from page 31, the
            # Aggressive Hybrid Fund's, and page 64 (a list of managers' funds) passed the name check for both
            other = [p for p in used if len(pages.text(p).strip()) > 200
                     and not rules.names_fund(rules.Page(fund_name=fund_name, text=pages.text(p), amc=amc))]
            if other:
                out.append(f"source_pages {other} do not print this fund's name ({fund_name}). If the AUM or managers "
                           "came from them, make sure they are this fund's and not a neighbouring fund's (a page printed "
                           "as a picture may be the fund's own); otherwise drop them from source_pages.")
        seen = "\n\n".join(pages.text(p) for p in sorted(read | set(used)))
        for m in d.fund_managers:
            c = ceased(m.name, seen)
            if c:
                out.append(f"The pages say '{c[0]}'. Leave {m.name} out if they left on or before the AUM date.")
        return out

    @tool
    def read_page(page: int) -> str:
        """Read one PDF page of the factsheet (text layer; a page printed as a picture is also sent as an image)."""
        text, png = pages.read(page)
        if png:
            images.append(png)
        if pages.valid(page):
            read.add(page)
        return text

    @tool
    def search_factsheet(query: str) -> str:
        """Find the factsheet's pages that print every word of the query, with a snippet of each (best first)."""
        return pages.search(query, near=first_pages[0] if first_pages else None)

    @tool
    def factsheet_date() -> str:
        """The month-end the whole factsheet's data is stated as on, read from its opening pages, with the evidence."""
        fs = pages.factsheet_date()
        return f"{fs.data_as_on_date} (from: {fs.evidence})" if fs.data_as_on_date else "no data date printed on the opening pages"

    @tool(args_schema=Submission)
    def submit(**answer) -> str:
        """Submit the final answer for this fund. It is checked; the reply says 'accepted' or lists problems to fix."""
        s = Submission(**answer)
        d = FundDetails(**s.model_dump(exclude={"confirm"}))
        issues = review(d)
        box["tries"] += 1
        if box["best"] is None or len(issues) <= len(box["best"][1]):
            box["best"] = (d, issues)
        # confirm counts only once submit has reported problems: sent up front it would skip the checks
        if not issues or (s.confirm and box["warned"]) or box["tries"] >= MAX_SUBMITS:
            box["final"] = (d, issues)
            return "accepted"
        box["warned"] = True
        return ("Not accepted yet. Problems:\n" + "\n".join(f"- {i}" for i in issues)
                + "\nFix them (read or search more if needed) and submit again, or submit with confirm=true and say "
                  "in notes why the pages really print it so.")

    state = run_agent(messages, [read_page, search_factsheet, factsheet_date, submit],
                      finished=lambda: box["final"] is not None, max_steps=MAX_STEPS, images=images)
    answer, issues = box["final"] or box["best"] or (None, None)
    error = state.get("error")
    if answer is None and not error:
        error = f"agent made {state.get('steps', 0)} model calls without submitting an answer"
    return answer, issues, trace(state.get("messages", [])), error


def extract(fund_name, scheme_name, category, markdown, text, pages: Pages, first_pages, factsheet_month=None,
            images=None, cache_file=None, force=False, amc="", manager_dates=None, average_aum=None):
    """
    FundDetails for one fund, starting from its first page(s) (markdown and text of `first_pages`; images are PNG
    bytes of those printed as pictures). `pages` is the whole factsheet, for the agent's tools. Every answer is
    written to cache_file (JSON) with its source - "agent", or "rules" when the agent could not answer - keyed on
    model, prompt version and first pages, so a re-run makes no model call for unchanged pages. The rules' values
    are stored beside the agent's answer (disagreements) for rules_eval.py.

    manager_dates(fund_name, text) -> {manager name: date as printed}: dates printed elsewhere in the factsheet (Kotak's
    'About our fund managers' pages), filled in for managers still without a date (free; the agent may already
    have found them with search_factsheet). average_aum(fund_name): the monthly average AUM printed elsewhere (Mirae).

    The stored answer is then reconciled with the rules and the pages (reconcile); that runs on every run, for free.
    """
    key = cache_key(fund_name, markdown, text, factsheet_month, images)
    cached, cached_issues = read_model_answer(cache_file, key, force)

    page = rules.Page(fund_name=fund_name, text=text, markdown=markdown, amc=amc)
    found, dropped = rules.run(page), []
    date = found.values.get("aum_as_on_date")
    if date_out_of_window(date, factsheet_month):
        del found.values["aum_as_on_date"]  # a wrong date is worse than none: the agent is asked for it
        dropped.append(f"rule {found.source.pop('aum_as_on_date')} read AUM date {date}, outside {factsheet_month}")
    on_page = rules.names_fund(page)

    # every fund goes to the agent; the rules answer only when it cannot (no credentials, an API error, no submit)
    answer, issues, calls, error = cached, cached_issues, None, None
    if answer is None:
        if _model_off:
            error = _model_off
        else:
            user = (f"Fund: {fund_name}\nScheme master name: {scheme_name}\nCategory: {category}\n"
                    f"Factsheet month (folder): {factsheet_month or 'unknown'}\n"
                    f"These are PDF pages {', '.join(map(str, first_pages))} of {len(pages)}"
                    f"{'' if on_page else ' - they do not seem to print this fund name'}.\n\n"
                    f"<docling_markdown>\n{markdown}\n</docling_markdown>\n\n"
                    f"<pdf_text_layer>\n{text}\n</pdf_text_layer>")
            content = [{"type": "text", "text": user}] + [png_part(png) for png in images or []]
            # the system prompt is identical for every fund and comes first, so provider prompt caching can reuse it
            answer, issues, calls, error = fund_agent(fund_name, amc, pages, list(first_pages), factsheet_month,
                                                      [SystemMessage(SYSTEM_PROMPT), HumanMessage(content)])

    if answer is not None:
        result = clean(answer.model_copy(deep=True))
        notes = [result.notes]
        if issues:
            notes.append("kept despite: " + " | ".join(i[:100] for i in issues))
    else:
        # rules only: the guard says whether these are the fund's pages; the parser looks elsewhere when not
        result = clean(from_rules(found.values))
        result.fund_found = on_page
        result.source_pages = list(first_pages)
        notes = dropped + drop_ceased(result, text, factsheet_month)
        if error:
            notes.append(f"code rules only, model not asked: {error[:120]}")
    if not result.fund_found:
        blank(result)
    elif not result.fund_name_on_page:
        result.fund_name_on_page = rules.printed_name(page) or fund_name
    if result.fund_found:
        first_text = pages.text(first_pages[0]) if first_pages else text
        # a footnote can sit on the fund's next page even when a wrong printed index page left it out of the first
        # pages (ICICI JAN Balanced Advantage: index says 72, the fund is on 57 and its footnote on 58)
        after = first_pages[-1] + 1 if first_pages else 0
        footnotes = text + ("\n\n" + pages.text(after) if pages.valid(after) and pages.names(after, fund_name) else "")
        notes += reconcile(result, found, normal_text(text), normal_text(first_text), amc, normal_text(footnotes))
    if manager_dates and result.fund_found:
        notes += fill_dates(result, manager_dates(fund_name, text))
    if average_aum and result.fund_found and result.aum_month_end_cr and result.aum_monthly_average_cr is None:
        result.aum_monthly_average_cr = average_aum(fund_name)
    result.notes = "; ".join(filter(None, notes)) or None

    if answer is not None:
        # error here: the model dropped out after a submit the checks refused. Its best answer stands for this run
        # and is asked again on the next, rather than being stored as final
        pending = {"model_pending": error} if error else {}
        write_cache(cache_file, key, result, source="agent", unresolved=issues, rules=found.source, **pending,
                    model_result=answer.model_dump(), tool_calls=calls if calls is not None else read_trace(cache_file),
                    disagreements={f: [found.values[f], getattr(answer, f)] for f in SCALARS
                                   if f in found.values and getattr(answer, f) is not None and found.values[f] != getattr(answer, f)})
    else:
        write_cache(cache_file, key, result, source="rules", rules=found.source, not_printed=sorted(found.not_printed),
                    model_pending=error)  # why the model could not answer; retried on the next run
    return result


def read_trace(cache_file):
    """The tool calls stored with a cached answer, kept when the answer is reused."""
    try:
        return json.loads(cache_file.read_text(encoding="utf-8")).get("tool_calls")
    except Exception:
        return None


# ---------------------------------------------------------------- factsheet date
class FactsheetDate(BaseModel):
    data_as_on_date: Optional[str] = Field(default=None, description="YYYY-MM-DD month-end the factsheet's portfolio / AUM data is as on")
    evidence: Optional[str] = Field(default=None, description="The phrase on the page the date was taken from")


FACTSHEET_DATE_PROMPT = """You get the opening pages of an Indian Specialized Investment Fund (SIF) monthly factsheet (cover, index, top holdings, market review).
Return the month-end date the factsheet's portfolio and AUM data is stated as on, taken from a date PRINTED on these
pages, e.g. "As on 31-December-25", "as on January 30, 2026", "MONTHLY FACTSHEET - 31ST JANUARY, 2026".
A factsheet is often published in the month after its data, so an explicit "as on" / portfolio / holdings date wins
over a bare month name on the cover. Ignore risk-free rate / MIBOR, NFO, tax-year and other event dates.
Return YYYY-MM-DD and quote the phrase you used; return null if no such date is printed - do not guess."""


def factsheet_date(opening_text, cache_file=None, force=False):
    """
    Data month-end of a whole factsheet, for funds whose page prints no AUM / NAV date (ABSL: 'AUM as on last
    day'). The folder month is deliberately not passed: given it, the model echoed it back for ABSL JAN, whose
    data is as on 31-December-25. Without the model it is unknown and each fund's own printed date is used.
    """
    key = cache_key("factsheet date", FACTSHEET_DATE_PROMPT, opening_text, None, None)
    cached = read_cache(cache_file, key, force)
    if cached is not None:
        return FactsheetDate.model_validate(cached["result"])
    if _model_off:
        return FactsheetDate()
    try:
        result = model().with_structured_output(FactsheetDate).invoke(
            [SystemMessage(FACTSHEET_DATE_PROMPT), HumanMessage(f"<pages>\n{opening_text}\n</pages>")]) or FactsheetDate()
    except Exception as e:
        if auth_error(e):
            raise
        print(f"  ! factsheet date: {type(e).__name__}: {str(e)[:200]}")
        return FactsheetDate()
    if result.data_as_on_date and not DATE_OK.fullmatch(result.data_as_on_date):
        result.data_as_on_date = None
    write_cache(cache_file, key, result)
    return result
