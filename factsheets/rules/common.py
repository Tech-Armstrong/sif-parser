"""Shared patterns and parsers for the rules: dates, numbers and person names as factsheets print them."""
import re
from datetime import datetime

MON = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?"
ORD = r"(?:st|nd|rd|th)?"
# 14-Nov-18, 23rd November 2016, November 01, 2024, 31st March, 2026, June 2022, 30.04.2026, 30.09.26 (HSBC)
DATE = (rf"(?:\d{{1,2}}{ORD}[\s\-/.]*{MON}[\s\-/.,]*\d{{4}}(?!\d)|\d{{1,2}}{ORD}[\s\-/.]*{MON}[\s\-/.,]*\d{{2}}(?!\d)"
        rf"|{MON}\s*\d{{1,2}}{ORD},?\s*\d{{4}}|{MON}[\s\-,']*\d{{4}}|{MON}[\-']\d{{2}}(?!\d)|\d{{1,2}}[./-]\d{{1,2}}[./-]\d{{4}}"
        rf"|\d{{1,2}}[./]\d{{1,2}}[./]\d{{2}}(?![\d.]))")
NUM = r"\d{1,3}(?:,\d{2,3})+(?:\.\d+)?|\d+(?:\.\d+)?"
NAME = r"[A-Z][A-Za-z'.]*(?:[ \t]+[A-Z][A-Za-z'.]*){0,3}"  # 1-4 capitalised words on one line
NAME_WRAP = r"[A-Z][A-Za-z'.]*(?:(?:[ \t]*\n[ \t]*|[ \t]+)[A-Z][A-Za-z'.]*){0,3}"  # the same, wrapped onto the next line
MARKS = r"[*^#$@†]*"  # footnote marks printed after a name
TITLE = r"(?:Mrs|Mr|Ms|Dr|Shri)\b\.?\s*"  # \b: 'Shri' must not eat the start of 'Shridatta'

_FORMATS_DAY = ["%d %b %Y", "%d %b %y", "%b %d %Y", "%d %m %Y", "%d %m %y"]
_FORMATS_MONTH = ["%b %Y", "%b %y"]


def parse_date(s):
    """'14-Nov-18' -> '2018-11-14', 'June 2022' -> '2022-06'; None if it does not read as a date."""
    if not s:
        return None
    t = re.sub(r"(\d)(st|nd|rd|th)\b", r"\1", s.strip(), flags=re.I)
    t = re.sub(r"[\s\-/.,']+", " ", t).strip()
    # any spelling of a month to its 3 letters: 'Sept', 'Septembr' (ICICI), 'September'
    t = re.sub(r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*", r"\1", t, flags=re.I)
    for fmts, out in ((_FORMATS_DAY, "%Y-%m-%d"), (_FORMATS_MONTH, "%Y-%m")):
        for f in fmts:
            try:
                return datetime.strptime(t, f).strftime(out)
            except ValueError:
                pass
    return None


def number(s):
    return float(s.replace(",", ""))


def clean_name(s):
    s = re.sub(r"[*^#$@†]+", " ", s)
    s = re.sub(rf"^\s*{TITLE}", "", s.strip())
    return " ".join(s.split()).strip(" .,-")


def manager(name, since=None, role=None, printed=None):
    """
    One manager. since=None means the page prints no date for this person (Sundaram dates only its lead manager,
    Kotak none): the agent then does not ask the model for it. printed: the text shown instead of the date
    ('Since Inception', with since the inception date).
    """
    printed = printed or since
    return {"name": clean_name(name), "managing_since": parse_date(since),
            "managing_since_as_printed": " ".join(printed.split()) if printed else None,
            "role": " ".join(role.split()).strip(" ()-–") or None if role else None,
            "since_not_printed": since is None}
