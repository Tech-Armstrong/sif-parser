"""Rules for SIF factsheet layouts that MF's rules (aum.py, managers.py) do not read.

As in managers.py, a manager rule returns the whole list or nothing: each pattern matches the complete manager
block of one layout. SIF factsheets often print one 'Managing Since' date for every manager of a strategy launched
with all of them (RedHex, Altiva); a rule then gives that date to each.
"""
import re

from . import ASK_MODEL, rule
from .aum import CUR, UNIT
from .common import DATE, MARKS, NAME, NAME_WRAP, NUM, TITLE, manager, number, parse_date
from .managers import _consecutive, _covered


# ---------------------------------------------------------------- fund managers
@rule("fund_managers")
def title_name_paren_since(page):
    """
    Diviniti 'FUND MANAGER' / 'Mr. Alok Ranjan (Since 11-July-2026)'; Apex 'Mr. Lovelish Solanki (Since March 25,
    2026)' / 'Mr. Mohit Sharma (Since March 25, 2026)' / ..., one manager per line, the lines one after another.
    """
    rx = rf"^[ \t]*{TITLE}(?P<name>{NAME}){MARKS}[ \t]*\(\s*Since\s+(?P<since>{DATE})\s*\)[ \t]*$"
    rows = _consecutive(list(re.finditer(rx, page.text, re.M)), page.text)
    return rows and {"fund_managers": [manager(m.group("name"), m.group("since")) for m in rows]}


@rule("fund_managers")
def managing_this_investment_strategy(page):
    """
    iSIF 'Fund Managers :' / 'Mittul Kalawadia (Managing this Investment Strategy \\nsince June, 2026 & Overall 20 years
    of experience)', one entry per manager up to 'Inception Date'. Every entry of the block must be read.
    """
    head = re.search(r"Fund\s+Managers?\s*:", page.text)
    if not head:
        return None
    block = re.split(r"Inception\s+Date|Plans\s*:|Benchmark\s*:", page.text[head.end():], maxsplit=1)[0]
    rows = _covered(rf"(?:{TITLE})?(?P<name>{NAME}){MARKS}\s*\(\s*Managing\s+this\s+Investment\s+Strategy\s+since\s+"
                    rf"(?P<since>{DATE})\s*&\s*Overall[^)]*\)", " ".join(block.split()))
    return rows and {"fund_managers": [manager(m.group("name"), m.group("since")) for m in rows]}


@rule("fund_managers")
def name_experience_in_managing_since(page):
    """
    Summit (Invesco's layout, other order): 'Fund Manager & Experience' / 'Hiten Jain' / 'Total Experience' /
    'Experience in managing this fund:' / 'Since July 22, 2026' / '17 Years' (AUG: the first name on the heading's line).
    """
    head = re.search(r"Fund\s+Manager\s*&\s*Experience\b", page.text)
    if not head:
        return None
    block = re.split(r"\n\s*(?:AAuM|AuM|Date\s+of\s+Allotment|NAV\b)", page.text[head.end():], maxsplit=1, flags=re.I)[0]
    rows = _covered(rf"(?:{TITLE})?(?P<name>{NAME})\s*(?:\((?P<role>[^)]*)\))?\s*Total\s+Experience\s*(?:\d+\s*Years?\s*)?"
                    rf"Experience\s+in\s+managing\s+this\s+fund\s*:?\s*Since\s+(?P<since>{DATE})\s*(?:\d+\s*Years?)?",
                    " ".join(block.split()))
    return rows and {"fund_managers": [manager(m.group("name"), m.group("since"), m.group("role")) for m in rows]}


@rule("fund_managers")
def portion_lines_names_only(page):
    """
    Arudha 'Fund Manager:' / 'Equity Portion: Mr. Kapil Kankonkar, Mr. Nilesh Saha' / 'Debt Portion: Mr. Brijesh Shah,
    Mr. Debraj Lahiri.' - no dates, except one who joined later: 'Mr. Rishi Sharma (w.e.f. \\nMarch 27, 2026)'.
    """
    head = re.search(r"Fund\s+Managers?\s*:[ \t]*\n", page.text)
    if not head:
        return None
    lines = re.findall(r"^[ \t]*([A-Za-z ]+?)\s+Portion\s*:\s*(.+?)(?=^[ \t]*[A-Za-z ]+?\s+Portion\s*:|\n\s*\*|\Z)",
                       page.text[head.end():head.end() + 600], re.M | re.S)
    if not lines:
        return None
    found = []
    for role, names in lines:
        for part in re.split(rf"\s*,\s*(?={TITLE})|\s*&\s*(?={TITLE})|\s+and\s+(?={TITLE})", " ".join(names.split()).rstrip(".")):
            e = re.fullmatch(rf"{TITLE}(?P<name>{NAME}){MARKS}\s*(?:\(\s*w\.e\.f\.?\s*(?P<since>{DATE})\s*\))?\.?", part.strip())
            if not e:
                return None
            found.append(manager(e.group("name"), e.group("since"), role))
    return {"fund_managers": found}


@rule("fund_managers")
def sections_with_managing_since(page):
    """
    Altiva Ex-Top 100 'Fund Manager(s)' / 'Equity:' / 'Mr. Nikhil Gada (Managing since 8\\nth June 2026) | Experience:
    14 Years' / 'Mr. Ashish Sood1' / 'Overseas:' / 'Mr. Amit Vora (Managing since ...)'. A name with a footnote number
    and no date is the one the footnote dates: '1 Mr. Ashish Sood has been appointed as the Assistant Fund Manager in
    Altiva Equity Ex-Top 100 Long-Short Fund with effect from 03rd August 2026.'
    """
    head = re.search(r"Fund\s+Manager\(s\)\s*\n", page.text)
    if not head:
        return None
    block = re.split(r"\n\s*(?:Subscription|Inception|Expense|Plan\s*&|Exit\s+load|Min\b)", page.text[head.end():], maxsplit=1, flags=re.I)[0]
    block = re.sub(r"(\d)\s*\n\s*(st|nd|rd|th)\b", r"\1\2", block)
    found, role = [], None
    for line in block.split("\n"):
        s = line.strip()
        sec = re.fullmatch(r"([A-Za-z ]+):", s)
        if sec:
            role = sec.group(1)
            continue
        if not s:
            continue
        m = re.fullmatch(rf"{TITLE}(?P<name>{NAME})\s*\(\s*Managing\s+since\s+(?P<since>{DATE})\s*\).*", s)
        if m:
            found.append(manager(m.group("name"), m.group("since"), role))
            continue
        m = re.fullmatch(rf"{TITLE}(?P<name>{NAME}?)(?P<note>\d)", s)
        if not m:
            return None
        note = re.search(rf"^\s*{m.group('note')}\s*{TITLE}{re.escape(m.group('name'))}\s+has\s+been\s+appointed\s+as\s+the\s+"
                         rf"(?P<role>Assistant\s+Fund\s+Manager|Fund\s+Manager)[^.\n]*?with\s+effect\s+from\s+(?P<since>{DATE})",
                         _flat_lines(page.text), re.M)
        if not note:
            return None
        found.append(manager(m.group("name"), note.group("since"), " ".join(note.group("role").split())))
    return found and {"fund_managers": found}


def _flat_lines(text):
    """'8\\nth June 2026' (an ordinal wrapped onto the next line, Altiva) put back together."""
    return re.sub(r"(\d)\s*\n\s*(st|nd|rd|th)\b", r"\1\2", text)


@rule("fund_managers")
def name_experience_since_columns(page):
    """
    Altiva Hybrid's 'Fund Manager' table, read column by column per section: 'Name Experience Managing Since' /
    'Equity' / 'Mr. Bhavesh Jain' / 'Mr. Bharat Lahoti' / '17 Years' / '18 Years' / '20th October 2025' / 'Debt' / ...
    / 'Overseas Mr. Amit Vora 16 Years 20th October 2025'. Per section the dates pair one to one with the names, or one
    date (a merged cell) is every name's.
    """
    head = re.search(r"^Name\s+Experience\s+Managing\s+Since\s*$", page.text, re.M)
    if not head:
        return None
    block = re.split(r"\n\s*(?:Subscription|Plan\s*&|Exit\s+load|Inception)", page.text[head.end():], maxsplit=1, flags=re.I)[0]
    # older months put the section and its first name on one line ('Equity Mr. Bhavesh Jain') and the last
    # experience and the first date on another ('18 Years 20th October 2025')
    sections, cur = [], None
    for line in block.split("\n"):
        s = line.strip()
        if not s:
            continue
        inline = re.fullmatch(rf"(?P<role>[A-Z][a-z]+)\s+{TITLE}(?P<name>{NAME})\s+\d+\s*Years?\s+(?P<since>{DATE})", s)
        named = re.fullmatch(rf"(?:(?P<role>[A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\s+)?{TITLE}(?P<name>{NAME}){MARKS}", s)
        dated = re.fullmatch(rf"(?:\d+\s*Years?\s*)?(?P<since>{DATE})?", s)
        if inline:
            sections.append([inline.group("role"), [inline.group("name")], [inline.group("since")]])
            cur = None
        elif named:
            if named.group("role"):
                cur = [named.group("role"), [], []]
                sections.append(cur)
            if cur is None:
                return None
            cur[1].append(named.group("name"))
        elif re.fullmatch(r"[A-Z][A-Za-z ]{2,30}", s):
            cur = [s, [], []]
            sections.append(cur)
        elif dated and cur is not None:
            if dated.group("since"):
                cur[2].append(dated.group("since"))
        else:
            return None
    found = []
    for role, names, dates in sections:
        if not names or not (len(dates) == len(names) or len(dates) == 1):
            return None
        found += [manager(n, dates[k] if len(dates) > 1 else dates[0], role) for k, n in enumerate(names)]
    return found and {"fund_managers": found}


@rule("fund_managers")
def bulleted_names_one_date(page):
    """
    RedHex (HSBC's box): 'Name of Fund Managers' / '• Mr. Shriram Ramanathan (Fixed Income, REITs, \\nand InvITs)' / ...
    / 'Total Experience : 25 Years / 33 Years / ...' / 'Managing Since : June 19, 2026' - one date for every manager
    (the strategy's launch). MF's bulleted_names_slash_dates reads the box when each manager has a date.
    """
    from .managers import _bulleted_names
    rest, names = _bulleted_names(page.text)
    since = rest and re.search(rf"Managing\s+since\s*:?\s*({DATE})\s*\n", rest, re.I)
    if not names or not since or len(names) == 1:
        return None
    if re.match(rf"\s*/\s*{DATE}", rest[since.end(1):]):
        return None  # several dates: MF's rule pairs them
    return {"fund_managers": [manager(n, since.group(1), r) for n, r in names]}


@rule("fund_managers")
def name_role_line_managing_this_fund_since(page):
    """
    Prism: 'Mr. Manish Kalani' / 'Fund Manager' / 'Total Experience 19 Years' / 'Managing this fund since' / 'July 2026',
    once per manager ('SIF CIO and Fund Manager' as the role line). Every entry of the block must be read.
    """
    rx = (rf"^[ \t]*{TITLE}(?P<name>{NAME}){MARKS}[ \t]*\n\s*(?P<role>[^\n]*Fund\s+Manager[^\n]*)\n\s*Total\s+Experience[^\n]*\n"
          rf"\s*Managing\s+this\s+fund\s+since\s*\n?\s*(?P<since>{DATE})")
    rows = list(re.finditer(rx, page.text, re.M | re.I))
    if not rows or len(re.findall(r"Managing\s+this\s+fund\s+since", page.text, re.I)) != len(rows):
        return None
    out = []
    for m in rows:
        role = " ".join(m.group("role").split())
        out.append(manager(m.group("name"), m.group("since"), None if role.lower() == "fund manager" else role))
    return {"fund_managers": out}


# ---------------------------------------------------------------- AUM
@rule("aum_monthly_average_cr")
def two_monthly_averages(page):
    """
    Prism prints 'Monthly Average AUM ₹ 76.48 Cr' in its FUND SIZE box every month since July (left over from its first
    month) and the month's own 'Monthly Average AUM ₹ 369.21 Cr' further down: which one is current is not in the text
    layer, so the model reads it.
    """
    figures = {f for f in re.findall(rf"Monthly\s+Average\s+AUM\s*\n?\s*{CUR}\s*({NUM})\s*Cr", page.text, re.I)}
    if len(figures) > 1:
        return {"aum_monthly_average_cr": ASK_MODEL}


@rule("aum_month_end_cr", "aum_monthly_average_cr", "aum_as_on_date")
def values_above_aum_as_on_label(page):
    """
    Apex prints the figures above the label that gives their unit and date: 'Month End AUM 353.87' / 'Monthly Average
    AUM 275.32' / 'AUM as on Sept 30, 2026 (₹ Cr)'.
    """
    m = re.search(rf"^Month\s+End\s+AUM\s+({NUM})\s*\n\s*Monthly\s+Average\s+AUM\s+({NUM})\s*\n\s*AUM\s+as\s+on\s+({DATE})\s*"
                  rf"\(\s*{CUR}\s*Cr\.?\s*\)", page.text, re.M | re.I)
    return m and {"aum_month_end_cr": number(m.group(1)), "aum_monthly_average_cr": number(m.group(2)),
                  "aum_as_on_date": parse_date(m.group(3))}


@rule("aum_month_end_cr", "aum_monthly_average_cr")
def month_end_and_monthly_avg_one_line(page):
    """Altiva: 'AUM (in INR Cr) Month end: Rs. 1,548 Cr | Monthly Avg: Rs. 1,508 Cr'."""
    m = re.search(rf"AUM[^\n]{{0,20}}?Month\s+end\s*:\s*{CUR}\s*({NUM})\s*Cr\b[^\n|]*\|\s*Monthly\s+Avg\.?\s*:\s*{CUR}\s*({NUM})\s*Cr\b",
                  page.text, re.I)
    return m and {"aum_month_end_cr": number(m.group(1)), "aum_monthly_average_cr": number(m.group(2))}


@rule("aum_month_end_cr", "aum_monthly_average_cr", "aum_as_on_date")
def aum_paren_as_on_then_figure(page):
    """
    RedHex: 'AUM (as on 30.09.26) ₹ 1504.00 Cr.' / 'AAUM (for the 30.09.2026) ₹ 1289.04 Cr.', or with the figure on the
    next line: 'AAUM (as on 30.06.2026)' / '` 561.39 crores'. The first-page 'AUM' / '(as on 30.09.2026)' / 'Rs. 1,504
    crores' reads the same.
    """
    out = {}
    for label, field in (("AAUM", "aum_monthly_average_cr"), ("AUM", "aum_month_end_cr")):
        # case-sensitive label (AUM must not match the 'AuM' of prose), any case for the unit ('crores')
        m = re.search(rf"(?<![A-Za-z]){label}\s*\(\s*(?i:as\s+on|for\s+the)\s+({DATE})\s*\)\s*{CUR}\s*({NUM})\s*(?i:{UNIT})", page.text)
        if m and (m.group("unit") or "").lower().startswith("cr"):
            out[field] = number(m.group(2))
            if field == "aum_month_end_cr":
                out["aum_as_on_date"] = parse_date(m.group(1))
    return out if "aum_month_end_cr" in out else None
