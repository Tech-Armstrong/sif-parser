"""Fund manager rules: every manager of THIS fund with the date they started managing it.

A rule returns the whole list or nothing: a partial list would hide a co-manager, so each pattern matches the
complete manager block of one layout.
"""
import re

from . import rule
from .common import DATE, MARKS, NAME, NAME_WRAP, TITLE, manager

ROLES = r"Equity|Debt|Fixed\s+Income|Overseas|Foreign|Arbitrage|Commodit(?:y|ies)|Gold|Silver|International"


@rule("fund_managers")
def name_is_managing_since(page):
    """Axis: 'Shreyash Devalkar is managing the scheme since 23rd November 2016 and he manages ...'."""
    found = [manager(m.group(1), m.group(2)) for m in re.finditer(
        rf"({NAME})\s+is\s+managing\s+the\s+(?:scheme|fund)\s+since\s+({DATE})", page.text)]
    return {"fund_managers": found}


@rule("fund_managers")
def names_line_then_since_each(page):
    """
    Axis tile: 'Mr. Devang Shah Mr. Sachin Jain Mr. Jayesh Sundar' on one line (optionally '(for Foreign
    Securities)' under the last), then per manager in the same order 'Work experience: 20 years.He has been
    managing this fund since 7th June 2016'. Names and dates must pair one to one.
    """
    first = re.search(r"Work\s+experience\s*:", page.text)
    if not first:
        return None
    lines = page.text[:first.start()].rstrip("\n").split("\n")
    role = None
    if lines and re.fullmatch(r"\s*\(([^)]*)\)\s*", lines[-1]):
        role = re.fullmatch(r"\s*\(([^)]*)\)\s*", lines.pop()).group(1)
    names = [n for n in re.split(rf"\b{TITLE}", lines[-1] if lines else "") if n.strip()]
    dates = re.findall(rf"has\s+been\s+managing\s+this\s+fund\s+since\s+({DATE})", page.text[first.start():])
    if not names or len(names) != len(dates) or not re.match(rf"\s*{TITLE}", lines[-1]):
        return None
    if role:
        role = re.sub(r"^for\s+", "", role.strip(), flags=re.I)
    return {"fund_managers": [manager(n, d, role if i == len(names) - 1 else None) for i, (n, d) in enumerate(zip(names, dates))]}


@rule("fund_managers")
def title_name_managing_fund_since(page):
    """
    Canara: 'Mr. Avnish Jain (Managing fund since 7-Oct-13 & Overall experience of 31 years) (For Debt Portfolio )',
    also 'Ms. Silky Jain (Assistant Fund Manager) (Managing fund since ...' and '*Mr. Amit Kadam (Dedicated Fund
    Manager for Overseas investments)\\nManaging fund since ...'. A tenure printed as 'since 30-May-25 to 01-Apr-26'
    has ended, so that person is not a manager of the fund any more.
    """
    found = []
    for m in re.finditer(rf"{TITLE}(?P<name>{NAME}){MARKS}\s*(?:\((?P<pre>[^)]*)\)\s*)?\(?\s*Managing\s+(?:the\s+)?fund\s+since\s+"
                         rf"(?P<since>{DATE})(?P<until>\s+to\s+{DATE})?[^)]*\)\s*(?:\(\s*For\s+(?P<post>[^)]+?)\s*\))?", page.text):
        if m.group("until"):
            continue
        role = m.group("post") and re.sub(r"\s*Portfolio$", "", m.group("post").strip(), flags=re.I)
        found.append(manager(m.group("name"), m.group("since"), role or m.group("pre")))
    return {"fund_managers": found}


@rule("fund_managers")
def fund_manager_dash_since(page):
    """ABSL: 'Fund Manager - Mr. Mohit Sharma' then 'Managing the Fund Since: November 01, 2024', once per manager."""
    found = [manager(m.group(1), m.group(2)) for m in re.finditer(
        rf"Fund\s+Manager\s*[-:]\s*(?:{TITLE})?({NAME})\s*\n\s*Managing\s+the\s+Fund\s+Since\s*:?\s*({DATE})", page.text, re.I)]
    return {"fund_managers": found}


def _bulleted_names(text):
    """
    (rest of the text, [(name, role)]) after 'Name of Fund Managers', one '• Mr. Neelotpal Sahai (Equity)' bullet
    per manager; a role may wrap ('(Overseas \\nInvestments)'). Without bullets every line must start with a title.
    """
    head = re.search(r"Name\s+of\s+(?:the\s+)?Fund\s+Managers?\s*:?[ \t]*\n", text, re.I)
    if not head:
        return None, []
    rest, lines, bullets = text[head.end():], [], None
    for line in rest.split("\n"):
        s = line.strip()
        if lines and lines[-1].count("(") > lines[-1].count(")"):
            lines[-1] += " " + s
            continue
        mark = re.match(r"(?:[•·]\s*(?:[IVX]{1,4}\.)?|[IVX]{1,4}\.)\s*", s)  # '• Name', Sundaram OCT '• I. Name' / 'II.Name'
        bullet = bool(mark)
        bullets = bullet if bullets is None else bullets
        if not s or bullet != bullets or (not bullets and not re.match(TITLE, s)):
            break  # 'Total \\nExperience' after the last bullet is not a name
        lines.append(s[mark.end():] if mark else s)
    names = []
    for s in lines:
        m = re.fullmatch(rf"(?:{TITLE})?(?P<name>{NAME}){MARKS}\s*(?:\((?P<role>[^)]*)\))?\s*{MARKS}", s)
        if not m:
            return None, []
        names.append((m.group("name"), m.group("role")))
    return rest, names


@rule("fund_managers")
def bulleted_names_slash_dates(page):
    """
    The 'Fund Manager Details' box of ABSL (from SEP 2026), Edelweiss, HSBC, Tata and Nippon: 'Name of Fund
    Managers', one bullet per manager (role in brackets), 'Total Experience 22 / 18', then 'Managing since
    November 03,2023 / November 21,2022' - dates in the same order, wrapped anywhere ('April \\n01,2017').
    Names and dates must pair one to one.
    """
    rest, names = _bulleted_names(page.text)
    since = rest and re.search(r"Managing\s+since\s*:?\s*", rest, re.I)
    if not names or not since:
        return None
    block = re.split(r"\n\s*(?:Minimum|Load\s+Structure|Exit\s+Load|Expense|Total\s+Experience|Benchmark)", rest[since.end():],
                     maxsplit=1)[0]
    dates = re.findall(DATE, block)
    if not dates:  # HSBC now and then prints the dates line above its label: 'Nov 26, 2022 / Oct 01, 2025 / \nAug 26, 2026 Managing Since'
        exp = re.search(r"Total\s+Experience[^\n]*\n", rest[:since.start()], re.I)
        dates = re.findall(DATE, rest[exp.end():since.start()]) if exp else []
    if len(names) != len(dates):
        return None
    return {"fund_managers": [manager(n, d, r) for (n, r), d in zip(names, dates)]}


@rule("fund_managers", amc="Sundaram")
def sundaram_dates_by_first_name(page):
    """
    Sundaram's box dates only some managers, by first name: '• Madanagopal Ramu' / '• Shalav Saket (Overseas)' /
    'Managing Since: 16 Mar 2026 (Madan)', or '01 Jan 2022 (Bharath); 04\\nOct 2024 (Clyton)'. A manager with no
    date has none printed.
    """
    rest, names = _bulleted_names(page.text)
    since = rest and re.search(r"Managing\s+since\s*:?\s*", rest, re.I)
    if not names or not since:
        return None
    block = re.split(r"\n\s*(?:Minimum|Total\s+Experience)", rest[since.end():], maxsplit=1)[0]
    dated = re.findall(rf"({DATE})\s*\(([A-Za-z]+)\)", block)
    if not dated or len(dated) != len(re.findall(DATE, block)):
        return None
    out = {n: [n, None, r] for n, r in names}
    for d, who in dated:
        hits = [n for n in out if n.lower().startswith(who.lower())]
        if len(hits) != 1:
            return None
        out[hits[0]][1] = d
    return {"fund_managers": [manager(*v) for v in out.values()]}


@rule("fund_managers")
def value_line_above_label(page):
    """
    ABSL (older layout), each value on the line ABOVE its label: ' Mr. Vishal Gajwani' / 'Fund Manager', then
    ' November 01, 2024' / 'Managing Fund Since'. Several managers share the lines ('Mr. A & Mr. B'); names and
    dates must pair one to one.
    """
    who = re.search(r"^([^\n]*)\n\s*Fund\s+Managers?\s*$", page.text, re.M)
    since = re.search(r"^([^\n]*)\n\s*Managing\s+Fund\s+Since\s*$", page.text, re.M)
    if not who or not since or not re.match(rf"\s*{TITLE}", who.group(1)):
        return None
    names = [n for n in re.split(rf"\s*(?:&|,|\band\b)?\s*\b{TITLE}", who.group(1)) if n.strip(" &,")]
    dates = re.findall(DATE, since.group(1))
    if not names or len(names) != len(dates):
        return None
    return {"fund_managers": [manager(n.strip(" &,"), d) for n, d in zip(names, dates)]}


@rule("fund_managers")
def name_experience_managing_since(page):
    """
    DSP 'FUND MANAGER' block, one entry per manager: name (role in brackets on the same or the next line), then
    'Total work experience of 18 years. Managing the Scheme since May 2021.' wrapped anywhere - even 'Managing
    t–his Scheme since May \\n2022.' The block is cut at each 'Total work experience'; every entry needs a date.
    """
    head = re.search(r"^FUND\s+MANAGERS?\s*$", page.text, re.M)
    if not head:
        return None
    block = page.text[head.end():]
    stop = re.search(r"^NAV\s+AS\s+ON\b", block, re.M | re.I)
    block = block[:stop.start()] if stop else block[:2000]
    found, start = [], 0
    for exp in re.finditer(r"Total\s+work\s+experience", block):
        chunk = block[start:exp.start()].strip(" .\t\n")  # also the full stop after the previous entry's date
        since = re.search(rf"since\s+({DATE})", block[exp.end():])
        if not chunk or not since:
            return None
        role = re.search(r"\(([^)]+)\)", chunk)
        name = re.sub(r"\([^)]*\)", " ", chunk).split("\n")[0]
        if not re.fullmatch(rf"\s*(?:{TITLE})?{NAME}\s*", name):
            return None
        found.append(manager(name, since.group(1), role and role.group(1)))
        start = exp.end() + since.end()
    return {"fund_managers": found}


@rule("fund_managers")
def manager_table_rows(page):
    """
    One row per manager under a 'Fund Manager' heading, date and experience in either order:
    Edelweiss 'Mr. Bharat Lahoti 18 years 01-Oct-21'; Baroda 'Equity Mr. Sanjay Chawla 14-Nov-18 33 years'.
    """
    head = re.search(r"Fund\s+Managers?\b", page.text, re.I)
    if not head:
        return None
    text = page.text[head.start():]
    # 'Fixed IncomeMr. Vikram Pamnani' (no space), 'Mr. Rohan Korde^ 01-May-26' (footnote mark),
    # 'Mr. Trideep \nBhattacharya 25 years 01-Oct-21' and 'Mr. Ankeet\nPandya\n01-Jan-25 11 years' (wrapped),
    # 'Mr. Kedar Karnik9 19 years' (footnote number glued to the name)
    years = r"\d+(?:\.\d+)?\s*[Yy]ears?"
    rows = list(re.finditer(rf"^[ \t]*(?:({ROLES})[ \t]*)?{TITLE}({NAME_WRAP}){MARKS}\d?\s+(?:{years}\s+)?({DATE})"
                            rf"(?:[ \t]+{years})?[ \t]*$", text, re.M))
    if not rows:
        return None
    # every titled line from the table's heading to the line after its last row must be a row we read, else a
    # co-manager is missing and the model is asked instead
    heads = list(re.finditer(r"Fund\s+Managers?\b", text[:rows[0].start()], re.I))
    after = text.find("\n", rows[-1].end() + 1)
    block = text[heads[-1].start() if heads else rows[0].start():after if after > 0 else len(text)]
    if len(re.findall(rf"^[ \t]*(?:(?:{ROLES})[ \t]*)?{TITLE}", block, re.M)) != len(rows):
        return None
    return {"fund_managers": [manager(m.group(2), m.group(3), m.group(1)) for m in rows]}


def inception(page):
    """The scheme's inception / allotment date, for managers printed as managing it 'since inception'."""
    # 'Date of Inception / Allotment May 24, 2013', 'Date of allotment: 10 June 2011', 'A llotm ent D a te\n13-May-2013'
    # (Motilal's spaced-out text layer) - a label, never the word 'inception' of 'Since inception. Mr. X: <date>'
    # not 'within 15 days from the date of allotment' in an exit load (Union APR took the market-cap date after it);
    # 'Date of Inception/ \nAllotment: 9th July 2010': Mirae SEP
    m = re.search(rf"(?<!the )(?<!the\n)(?<!the \n)(?:Date\s+of\s+Inception\s*/\s*Allotment|Date\s+of\s+(?:Inception|Allotment)|(?:Inception|A\s*llotm\s*ent)\s*D\s*a\s*te)"
                  rf"[^\n\d]{{0,40}}\n?\s*({DATE})", page.text, re.I)
    if not m and "whiteoak" in page.amc.lower():
        # WhiteOak prints its labels apart from the values; the inception date is the line with the ordinal
        # moved in front: 'nd 02 August 2022'
        m = re.search(rf"^\s*(?:st|nd|rd|th)\s+({DATE})\s*$", page.text, re.M)
    return m and m.group(1)


def _since(page, printed):
    """(date, as printed) for 'Since Inception' / 'its inception' or a printed date."""
    if re.search(r"inception", printed, re.I):
        return inception(page), "Since Inception"
    return printed, printed


def _titled_lines(block):
    return len(re.findall(rf"^[ \t]*{TITLE}", block, re.M))


@rule("fund_managers")
def name_line_then_managing_this_since(page):
    """
    Motilal / WhiteOak: one entry per manager, the name line (role in brackets) then 'Managing this fund since
    05-Aug-2025' or 'Managing this Scheme from 11th August 2022 / from its inception'; Motilal's foreign-securities
    manager reads 'He has been appointed as the Fund Manager for Foreign securities w.e.f. November 18, 2025.'
    Section lines 'For Equity Component' give the role. Every titled line in the block must be an entry read.
    """
    rx = (rf"^[ \t]*{TITLE}(?P<name>{NAME}){MARKS}[ \t]*(?:\((?P<role>[^)\n]*)\)?)?[^\n]*\n\s*"
          rf"(?:Managing\s+this\s+(?:fund|scheme)\s+(?:since|from)\s+(?P<since>{DATE}|its\s+inception)"
          rf"|He\s+has\s+been\s+appointed\s+as\s+the\s+Fund\s+Manager\s+for\s+(?P<role2>[A-Za-z ]+?)\s+w\.e\.f\.?\s*(?P<since2>{DATE}))")
    rows = list(re.finditer(rx, page.text, re.M | re.I))
    # every entry of the block read: counted by its date line, as a name can be printed letter-spaced
    # ('M r. A nk it A ga rw a l', Motilal FEB) and then is not seen as a name at all
    block = page.text[rows[0].start():] if rows else ""  # to the end: a letter-spaced entry may come last
    if not rows or len(re.findall(r"Managing\s+this\s+(?:fund|scheme)|appointed\s+as\s+the\s+Fund\s+Manager", block, re.I)) != len(rows):
        return None
    sections =[(s.start(), s.group(1)) for s in re.finditer(r"^For\s+([^\n]+?)\s*$", page.text, re.M)]
    found = []
    for m in rows:
        since, printed = _since(page, m.group("since") or m.group("since2"))
        if since is None:
            return None  # 'its inception' with no inception date on the page: the model reads it
        # Motilal's 'For Equity Component' line above a group of managers (only within this block)
        section = [s for pos, s in sections if rows[0].start() - 200 < pos < m.start()]
        found.append(manager(m.group("name"), since, m.group("role") or m.group("role2") or (section[-1] if section else None), printed))
    return {"fund_managers": found}


@rule("fund_managers")
def name_role_experience_managing_since(page):
    """
    PPFAS: 'Mr. Rajeev Thakkar: Chief Investment Officer - Equity \\nand Director' / 'Total Experience: 32 Years' /
    'Managing Since: Since Inception', once per manager. Every titled line in the block must be an entry read.
    """
    rx = (rf"^[ \t]*{TITLE}(?P<name>{NAME})\s*[:\-–]\s*(?P<role>[^\n]*(?:\n(?!\s*Total)[^\n]*)?)\n\s*Total\s+Experience[^\n]*\n"
          rf"\s*Managing\s+Since\s*:\s*(?P<since>Since\s+Inception|{DATE})")
    rows = list(re.finditer(rx, page.text, re.M | re.I))
    if not rows or _titled_lines(page.text[rows[0].start():rows[-1].end()]) != len(rows):
        return None
    found = []
    for m in rows:
        since, printed = _since(page, m.group("since"))
        if since is None:
            return None
        found.append(manager(m.group("name"), since, m.group("role"), printed))
    return {"fund_managers": found}


@rule("fund_managers")
def wef_date_then_name(page):
    """
    PGIM: 'Fund Manager: (w.e.f. April 15, 2024) Mr. Vivek Sharma (Equity Portion) (Over 15 years ...); (w.e.f.
    July 16, 2022) Mr. \\nPuneet Pal (Debt Portion) (...).' - the block runs to the sentence's full stop.
    """
    head = re.search(r"Fund\s+Managers?\s*:\s*(?=\(w\.e\.f)", page.text, re.I)
    if not head:
        return None
    block = re.split(r"\)\s*\.\s*\n", page.text[head.end():], maxsplit=1)[0]
    rows = list(re.finditer(rf"\(w\.e\.f\.?\s*(?P<since>{DATE})\)\s*{TITLE}(?P<name>{NAME_WRAP}){MARKS}\s*"
                            rf"(?:\((?P<role>[^)]*Portion)\))?", block, re.I))
    if not rows or len(re.findall(r"w\.e\.f", block, re.I)) != len(rows):
        return None
    return {"fund_managers": [manager(m.group("name"), m.group("since"), m.group("role")) for m in rows]}


@rule("fund_managers")
def managing_since_name_colon_date(page):
    """
    Union: 'Name of Fund Managers: Mr. Sanjay \\nBembalkar & Mr Vinod Malviya (For Arbitrage \\nPortion)' and
    'Managing Since: Mr. Sanjay Bembalkar: \\nJanuary 25, 2023. Mr. Vinod Malviya: Since \\ninception.', the two
    in either order and 'Total Experience' around them. Each named manager needs a date.
    """
    end = r"(?=Managing\s+Since|T\s*o\s*t\s*a\s*l\s+E|Name\s+of\s+Fund|The\s+risk|Additional|SIP\s+Amount|Minimum|$)"
    names = re.search(rf"Name\s+of\s+Fund\s+Managers?\s*:(.*?){end}", page.text, re.I | re.S)
    dates = re.search(rf"Managing\s+Since\s*:(.*?){end}", page.text, re.I | re.S)
    if not names or not dates:
        return None
    rows = list(re.finditer(rf"{TITLE}(?P<name>{NAME_WRAP})\s*:\s*(?P<since>Since\s+inception|{DATE})", dates.group(1), re.I))
    if not rows or len(re.findall(rf"\b{TITLE}", names.group(1))) != len(rows):
        return None
    found = []
    for r in rows:
        since, printed = _since(page, r.group("since"))
        if since is None:
            return None
        who = r"\s+".join(map(re.escape, r.group("name").split()))
        role = re.search(rf"{who}\s*\(\s*(?:For\s+)?([^)]*)\)", names.group(1), re.I)
        found.append(manager(r.group("name"), since, role and role.group(1), printed))
    return {"fund_managers": found}


def _consecutive(rows, text):
    """rows (regex matches in order) when nothing but blank space or a separator lies between them, else None."""
    for a, b in zip(rows, rows[1:]):
        if re.sub(r"[\s,;&]|\band\b", "", text[a.end():b.start()]):
            return None
    return rows or None


@rule("fund_managers")
def name_total_experience_managing_since(page):
    """
    HSBC (until JUL 2026), per manager: 'Neelotpal Sahai (Equity)' / 'Total Experience 31 Years' / 'Managing Since
    May 27, 2013'. The entries follow one another with nothing in between.
    """
    rx = (rf"^[ \t]*(?:{TITLE})?(?P<name>{NAME}){MARKS}[ \t]*(?:\((?P<role>[^)\n]*)\))?[ \t]*\n\s*Total\s+Experience[^\n]*\n"
          rf"\s*Managing\s+Since\s*:?\s*(?P<since>{DATE})")
    rows = _consecutive(list(re.finditer(rx, page.text, re.M)), page.text)
    return rows and {"fund_managers": [manager(m.group("name"), m.group("since"), m.group("role")) for m in rows]}


@rule("fund_managers")
def name_paren_managing_since(page):
    """
    Nippon (until AUG 2026): 'Rishit Parikh (Assistant Fund Manager) (Managing Since Aug 2024)' / 'Total Experience
    of more than 17 years'; Tata (until JUL 2026): 'Abhinav Sharma (Managing Since 05-Apr-23 and overall\\nexperience
    of 19 years), Hasmukh Vishariya (Managing Since\\n01-Mar-25 and ...)', with an 'ASSISTANT FUND MANAGER' heading
    over the assistant. Every 'Managing Since' under the 'Fund Manager' heading must be an entry read.
    """
    head = re.search(r"^\s*FUND\s+MANAGER(?:\(S\)|S)?\s*$", page.text, re.M | re.I)
    if not head:
        return None
    block = re.split(r"\n\s*(?:AMFI|BENCHMARK|Benchmark|NAV\b)", page.text[head.end():], maxsplit=1)[0]
    # the 'ASSISTANT FUND MANAGER' heading becomes a mark a name cannot run into (NAME_WRAP would take
    # 'FUND MANAGER Kapil Malhotra')
    block = re.sub(r"^\s*ASSISTANT\s+FUND\s+MANAGERS?\s*$", "§", block, flags=re.M)
    rx = (rf"(?:{TITLE})?(?P<name>{NAME_WRAP}){MARKS}\s*(?:\((?P<role>[^)]*)\)\s*)?\(\s*Managing\s+Since\s+(?P<since>{DATE})"
          rf"(?:\s*\)|\s+and\s+overall\s+experience[^)]*\))")
    rows = list(re.finditer(rx, block))
    if not rows or len(re.findall(r"Managing\s+Since", block)) != len(rows):
        return None
    found = []
    for m in rows:
        assistant = "§" in block[:m.start()]
        found.append(manager(m.group("name"), m.group("since"), m.group("role") or (assistant and "Assistant Fund Manager")))
    return {"fund_managers": found}


@rule("fund_managers", amc="Sundaram")
def sundaram_names_only(page):
    """Sundaram (until AUG 2026) prints names only: 'Fund Managers Rohit Seksaria, Shalav Saket\\n(Overseas)'."""
    m = re.search(r"^Fund\s+Managers?\s+(.+?)\n\s*Month\s+End\s+AUM", page.text, re.M | re.S)
    if not m:
        return None
    found = []
    for part in re.split(r"\s*,\s*|\s+&\s+|\s+and\s+", " ".join(m.group(1).split())):
        e = re.fullmatch(rf"(?:{TITLE})?(?P<name>{NAME})\s*(?:\((?P<role>[^)]*)\))?", part)
        if not e:
            return None
        found.append(manager(e.group("name"), None, e.group("role")))
    return {"fund_managers": found}


@rule("fund_managers", amc="ICICI")
def icici_currently_managed_by(page):
    """
    ICICI performance note: 'The scheme is currently managed by Sankaran Naren, Antariksha Banerjee ... and Masoomi
    Jhurmarvala.' then per manager 'Mr. Sankaran Naren has been managing this fund since Feb 2012.' Every listed
    manager needs a date.
    """
    m = re.search(r"currently\s+managed\s+by\s+(.*?)\s*\.(?:\s|$)", page.text, re.S)
    if not m:
        return None
    names = [n.strip() for n in re.split(r",|&|\band\b", " ".join(m.group(1).split())) if n.strip()]
    found = []
    for n in names:
        who = r"\s+".join(map(re.escape, n.split()))
        d = re.search(rf"{TITLE}{who}\s+has\s+been\s+managing\s+this\s+fund\s+since\s+({DATE})", page.text)
        if not d:
            return None
        found.append(manager(n, d.group(1)))
    return {"fund_managers": found}


def _covered(rx, block):
    """Matches of rx over the joined block, or None when any word of the block is left unread (a manager missed)."""
    rows = list(re.finditer(rx, block))
    left = re.sub(rx, " ", block)
    return rows if rows and not re.search(r"[A-Za-z]{2}", left) else None


@rule("fund_managers", amc="HDFC")
def hdfc_name_since_exp_table(page):
    """
    HDFC 'FUND MANAGER' table 'Name Since Total Exp', cells wrapped onto several lines: 'Gopal Agrawal' /
    '(Equity Portfolio)' / 'July 29,' / '2022' / 'Over 22' / 'years', or 'Amit Ganatra February' / '01, 2026'.
    """
    head = re.search(r"^Name\s+Since\s+Total\s+Exp\s*$", page.text, re.M)
    if not head:
        return None
    block = re.split(r"\n(?:DATE\s+OF\s+ALLOTMENT|NAV\b|INVESTMENT\s+OBJECTIVE|[A-Z][A-Z &/]{8,}\n)", page.text[head.end():], maxsplit=1)[0]
    rows = _covered(rf"(?P<name>{NAME}){MARKS}\s*(?:\((?P<role>[^)]*)\)\s*)?(?P<since>{DATE})\s*Over\s+\d+\s*years?",
                    " ".join(block.split()))
    return rows and {"fund_managers": [manager(m.group("name"), m.group("since"), m.group("role")) for m in rows]}


@rule("fund_managers", amc="SBI")
def sbi_name_experience_since_table(page):
    """
    SBI 'Name of Fund Managers | Total Experience | Managing Since' table, cells wrapped: 'Mr. Dinesh \\nBalachandran'
    / '- Equity' / '24 years Oct 2019', 'Ms. Mansi Sajeja \\n(Debt Portion) 19 years June 2021', 'Mr. Saurabh
    Pant# 18 years May - \\n2026'. The table ends at a footnote ('#') or the next section.
    """
    head = re.search(r"Managing\s*\n?\s*Since[ \t]*\n", page.text)
    if not head or not re.search(r"Fund\s*\n?\s*Managers?", page.text[max(0, head.start() - 120):head.start()]):
        return None
    block = re.split(r"\n\s*(?:#|Investment\s+Objective|TER\b|Lump\s*sum|Fund\s+Details|Minimum|\d\.\s)", page.text[head.end():],
                     maxsplit=1)[0]
    rows = _covered(rf"{TITLE}(?P<name>{NAME}){MARKS}\s*(?:-\s*(?P<r1>[A-Za-z ]+?)\s*|\((?P<r2>[^)]*)\)\s*)?\d+\s*years?\s+(?P<since>{DATE})",
                    " ".join(block.split()))
    return rows and {"fund_managers": [manager(m.group("name"), m.group("since"), m.group("r1") or m.group("r2")) for m in rows]}


@rule("fund_managers", amc="SBI")
def sbi_names_then_first_name_dates(page):
    """
    SBI (until JUL 2026): '• Fund Manager: Mr. Dinesh Balachandran (Equity \\nPortfolio) Ms. Mansi Sajeja (Debt
    Portion) & ...' / 'Managing Since:' / 'Mr. Dinesh May-2018', 'Mr. Bhavin Vithlani (w.e.f. April 2024)' or 'Mr.
    Dinesh Balachandran - Oct 2019' / 'Total Experience:' - dates matched by first name. Every manager named needs
    exactly one date.
    """
    m = re.search(r"Fund\s+Managers?\s*:\s*(?P<names>.*?)\n\s*Managing\s+Since\s*:\s*(?P<dates>.*?)\n\s*Total\s+Experience",
                  page.text, re.S)
    if not m:
        return None
    names = []
    for part in re.split(r"\s*(?:,|&|\band\b)?\s*(?=\b(?:Mrs|Mr|Ms|Dr|Shri)\b\.?\s)", " ".join(m.group("names").split())):
        if not part.strip():
            continue
        e = re.fullmatch(rf"{TITLE}(?P<name>{NAME}){MARKS}\s*(?:\((?P<role>[^)]*)\))?\s*{MARKS}", part.strip())
        if not e:
            return None
        names.append((e.group("name"), e.group("role")))
    dated = re.findall(rf"{TITLE}([A-Z][A-Za-z.]*)(?:\s+[A-Z][A-Za-z.]*)*?\s*{MARKS}\s*(?:-\s*|\(\s*w\.e\.f\.?\s*)?({DATE})",
                       m.group("dates"))
    if len(dated) != len(names):
        return None
    found = []
    for name, role in names:
        hit = [d for first, d in dated if name.lower().startswith(first.lower().rstrip("."))]
        if len(hit) != 1:
            return None
        found.append(manager(name, hit[0], role))
    return {"fund_managers": found}


@rule("fund_managers", amc="Invesco")
def invesco_experience_since(page):
    """
    Invesco 'Fund Manager & Experience': per manager 'Amey Sathe (Equity Investments)' / 'Total Experience 18 Years'
    / 'Experience in managing this fund:' / 'Since November 04, 2025'.
    """
    head = re.search(r"Fund\s+Manager\s*&\s*Experience\s*\n", page.text)
    if not head:
        return None
    block = re.split(r"\n\s*(?:Asset\s+Allocation|Performance\s+Attributes|Market\s+Capitali[sz]ation)", page.text[head.end():], maxsplit=1)[0]
    rows = _covered(rf"(?P<name>{NAME})\s*(?:\((?P<role>[^)]*)\))?\s*Total\s+Experience\s+\d+\s*Years?\s*Experience\s+in\s+managing\s+"
                    rf"this\s+fund\s*:?\s*Since\s+(?P<since>{DATE})", " ".join(block.split()))
    return rows and {"fund_managers": [manager(m.group("name"), m.group("since"), m.group("role")) for m in rows]}


@rule("fund_managers", amc="Kotak")
def kotak_names_no_dates(page):
    """
    Kotak's fund page names its managers with no date ('Fund Manager*: Mr. Abhishek Bisen &\\nMr. Dhananjay
    Tikariha'); the dates are on the 'About our fund managers' pages, not read here. Only a manager who joined
    recently carries one: 'Mr. Deepak Agrawal\\n(effective October\\n24, 2025)'.
    """
    m = re.search(r"Fund\s+Manager\*?\s*:\s*(.*?)\n\s*AAUM\s*:", page.text, re.S)
    if not m:
        return None
    found = []
    for part in re.split(rf"\s*(?:,|&|\band\b)\s*(?={TITLE})", " ".join(m.group(1).split())):
        e = re.fullmatch(rf"{TITLE}(?P<name>{NAME}){MARKS}\s*(?:\(\s*effective\s+(?P<since>{DATE})\s*\))?", part.strip())
        if not e:
            return None
        found.append(manager(e.group("name"), e.group("since")))
    return {"fund_managers": found}


def footnote_managers(text):
    """
    Managers only a footnote names, so no block rule sees them: ICICI 'overseas investment is managed by Ms. Sharmila
    D'Silva', HDFC 'Dedicated Fund Manager for Overseas Investments: Mr. Dhruv Muchhal (since June 22, 2023)'.
    """
    found = []
    for m in re.finditer(rf"overseas\s+investment\s+is\s+managed\s+by\s+{TITLE}(?P<name>{NAME_WRAP})", text, re.I):
        # NAME takes a full stop inside a name, so the sentence's end comes along: 'D'Silva. Refer page'
        found.append(manager(re.split(r"\.\s", m.group("name"))[0], None, "Overseas"))
    # HDFC AUG also announces the next one: '... (since June 22, 2023) (Total Experience: ...). Mr. Gopal Agrawal w.e.f. September 01, 2026'
    for block in re.finditer(r"Overseas\s+Investments\s*:(.{0,260})", text, re.I | re.S):
        for m in re.finditer(rf"{TITLE}(?P<name>{NAME_WRAP})\s*\(?\s*(?:since|w\.e\.f\.?)\s+(?P<since>{DATE})", block.group(1), re.I):
            found.append(manager(m.group("name"), m.group("since"), "Overseas"))
    return found


def date_beside_name(text, name):
    """The date on a table row's manager line: 'Mr. Amit Vora 16 years 23-Jun-24'."""
    who = r"\s+".join(map(re.escape, name.split()))
    m = re.search(rf"{who}{MARKS}[ \t]+\d+(?:\.\d+)?\s*[Yy]ears?[ \t]+({DATE})", text)
    return m and m.group(1)


def since_inception(text, name):
    """The page says this manager has run the fund from the start: '(since inception)', 'from its inception', 'Managing Since: Inception'."""
    who = r"\s+".join(map(re.escape, name.split()))
    m = re.search(rf"{who}(?P<gap>.{{0,110}}?)(?:since\s*:?|from\s+its)\s+inception", text, re.I | re.S)
    return bool(m and not re.search(rf"\b{TITLE}|\d{{4}}", m.group("gap")))
