"""AUM rules: month-end and monthly-average AUM in Rs crore, and the date printed for them."""
import re

from . import ASK_MODEL, NOT_PRINTED, rule
from .common import DATE, NUM, number, parse_date

UNIT = r"(?P<unit>Cr\b\.?|Crs\b\.?|Crores?\b|crore\b|Lakhs?\b|Lacs?\b)"  # 'crs': Kotak
CUR = r"(?:[₹`v]|Rs\.?|INR)?"
# between a label and its figure: footnote marks, '(₹) :', 'As on April 30, 2026 :', 'for the month of\nSeptember,
# 2026:' (Invesco), '(30-Sep-2026)' (Motilal) - but no other figure; the forms with a date are tried first so the
# day of the date is not taken for the figure
GAP = (rf"(?:[^\d]{{0,30}}?(?:as\s+on|for\s+the\s+month\s+of|\()\s*{DATE}\s*\)?[^\d\n]{{0,15}}?"
       rf"|[^\d\n]{{0,30}}?\n?[^\d\n]{{0,10}}?)")

AVG_LABEL = (r"Monthly\s*(?:Average|Avg\.?|AVG)\s*(?:Assets\s+Under\s+Management\s*\(AAUM\)|AUM)"
             r"|Monthly\s*AAUM|\bAAUM\b|Average\s+AUM|\bAvg\.?\s+AUM")  # 'Avg. AUM INR 3,763 Cr.': Sundaram
END_LABEL = (r"Month\s*End\s*(?:Assets\s+Under\s+Management\s*\(AUM\)|AUM)|Closing\s+AUM|Total\s+AUM"
             r"|(?<![A-Za-z])AUM(?![A-Za-z])")


def _figures(text, label):
    """(value in crore, match) for every '<label> ... <figure> <unit>' in text."""
    # the unit may sit in brackets after the figure: '`1,845.11 (` cr)' (Motilal)
    for m in re.finditer(rf"(?:{label}){GAP}{CUR}\s*(?P<num>{NUM})\s*(?:\(\s*{CUR}\s*)?{UNIT}?", text, re.I):
        if re.match(r",?\s*\d{4}\b", text[m.end("num"):]) and not m.group("unit"):
            continue  # 'March 31, 2026': the day of a date
        unit = m.group("unit")
        if re.match(rf"[ \t]*{CUR}[ \t]*(?:{NUM})[ \t]+{CUR}[ \t]*(?:{NUM})", text[m.end():]):
            continue  # one figure per fund across a snapshot table ('Month End AUM ` 84,011.96 ` 9,792.70 ` 16,330.36 ...')
        if unit is None and re.match(r"\s*\)", text[m.end("num"):]):
            continue  # a broken date in brackets: 'Latest AUM (31-August-20 26)' (Motilal AUG)
        if unit is None and re.search(r"\b(?:top|of)\s*$", text[m.start():m.start("num")], re.I):
            continue  # a label column: 'AAUM (in ` Cr)' / '% of top 5 holdings' (Diviniti)
        if unit is None:  # 'AUM as on March 31, 2026 (in ₹Crore)' then 'Month End AUM 5,103.03'; '(Rs. in Cr.) 5162.17'
            before = text[max(0, m.start() - 200):m.start("num")]
            if not re.search(r"crore|\bcrs?\b", before, re.I):  # '(Rs in Crs) 246.02': Diviniti
                continue
            unit = "crore"
        value = number(m.group("num"))
        yield (value / 100 if unit.lower().startswith("la") else value), m


@rule("aum_monthly_average_cr", "aum_month_end_cr", "aum_as_on_date")
def axis_style(page):
    """Axis: 'MONTHLYAVERAGE 30,376.30Cr. AS ON 31st March, 2026 28,661.27Cr.' (the AUM label is a section header)."""
    m = re.search(rf"MONTHLY\s*AVERAGE\s*({NUM})\s*Cr\.?\s*AS\s+ON\s*({DATE})\s*({NUM})\s*Cr", page.text, re.I)
    return m and {"aum_monthly_average_cr": number(m.group(1)), "aum_as_on_date": parse_date(m.group(2)),
                  "aum_month_end_cr": number(m.group(3))}


@rule("aum_month_end_cr", "aum_monthly_average_cr", "aum_as_on_date", amc="Aditya Birla")
def absl_fund_details(page):
    """
    ABSL 'Fund Details' table (from SEP 2026): 'AUM as on August 31, 2026 29,710.82' / 'AAUM as on August 31,
    2026 30,020.10'. No unit is printed; ABSL reports AUM in Rs crore (its earlier layout said '(in ₹Crore)').
    """
    end = re.search(rf"^AUM\s+as\s+on\s+({DATE})\s+({NUM})\s*$", page.text, re.M | re.I)
    avg = re.search(rf"^AAUM\s+as\s+on\s+({DATE})\s+({NUM})\s*$", page.text, re.M | re.I)
    if not end:
        return None
    return {"aum_month_end_cr": number(end.group(2)), "aum_as_on_date": parse_date(end.group(1)),
            "aum_monthly_average_cr": avg and number(avg.group(2))}


@rule("aum_month_end_cr", "aum_monthly_average_cr", "aum_as_on_date", amc="HDFC")
def hdfc_labels_then_figures(page):
    """
    HDFC: both labels first, then both figures - 'ASSETS UNDER MANAGEMENT €' / 'As on September 30, 2026' /
    'Average for Month of September, 2026' / '₹102,533.96Cr.' / '₹104,989.70Cr.' (month-end, then average).
    """
    m = re.search(rf"ASSETS\s+UNDER\s+MANAGEMENT[^\n]*\n\s*As\s+on\s+({DATE})\s*\n\s*Average\s+for\s+Month\s+of\s+{DATE}"
                  rf"\s*\n\s*{CUR}\s*({NUM})\s*Cr\.?\s*\n\s*{CUR}\s*({NUM})\s*Cr", page.text, re.I)
    return m and {"aum_as_on_date": parse_date(m.group(1)), "aum_month_end_cr": number(m.group(2)),
                  "aum_monthly_average_cr": number(m.group(3))}


@rule("aum_month_end_cr", "aum_monthly_average_cr", "aum_as_on_date", amc="Parag Parikh")
def ppfas_labels_then_figures(page):
    """
    PPFAS: 'AUM as on September 30, 2026:' / 'AAUM as on September 30, 2026:' with no figure on the line; the two
    figures come elsewhere on the page, together and in the same order: '` 1,45,070.50 Crores' / '` 1,46,528.06 Crores'.
    """
    end = re.search(rf"^AUM\s+as\s+on\s+({DATE})\s*:\s*$", page.text, re.M | re.I)
    avg = re.search(rf"^AAUM\s+as\s+on\s+{DATE}\s*:\s*$", page.text, re.M | re.I)
    pairs = re.findall(rf"^\s*{CUR}\s*({NUM})\s*Crores?\s*\n\s*{CUR}\s*({NUM})\s*Crores?\s*$", page.text, re.M | re.I)
    if not end or not avg or len(pairs) != 1:
        return None
    return {"aum_as_on_date": parse_date(end.group(1)), "aum_month_end_cr": number(pairs[0][0]),
            "aum_monthly_average_cr": number(pairs[0][1])}


@rule("aum_as_on_date", amc="Sundaram")
def sundaram_no_date(page):
    """
    Sundaram prints 'AUM (Cr) 4,155.35 Cr' (until AUG 2026 'Month End AUM INR 3,872 Cr.') with no date: the
    factsheet's own date is used for it.
    """
    if re.search(r"^(?:AUM\s*\(Cr\)|Month\s+End\s+AUM\s+INR)", page.text, re.M | re.I):
        return {"aum_as_on_date": NOT_PRINTED}


@rule("aum_month_end_cr", "aum_monthly_average_cr", amc="Nippon")
def nippon_fund_size(page):
    """Nippon (until AUG 2026): 'Fund Size' / 'Monthly Average: ₹ 48,143.23 Cr' / 'Month End: ₹ 48,531.15 Cr'."""
    m = re.search(rf"Fund\s+Size\s*\n\s*Monthly\s+Average\s*:\s*{CUR}\s*({NUM})\s*Cr\b[^\n]*\n\s*Month\s+End\s*:\s*{CUR}\s*({NUM})\s*Cr\b",
                  page.text, re.I)
    return m and {"aum_monthly_average_cr": number(m.group(1)), "aum_month_end_cr": number(m.group(2))}


@rule("aum_month_end_cr", "aum_monthly_average_cr", "aum_as_on_date", amc="Tata")
def tata_fund_size(page):
    """
    Tata (until JUL 2026): 'FUND SIZE' / 'Rs. 2783.80 (Rs. in Cr.)' / 'MONTHLY AVERAGE AUM' / 'Rs. 2758.68 (Rs. in
    Cr.)'; the only date is the portfolio's 'As on 31st July 2026 PORTFOLIO', the same month end.
    """
    fig = rf"\s*\n\s*Rs\.?\s*({NUM})\s*\(Rs\.?\s*in\s*Cr\.?\)"
    end = re.search(rf"^FUND\s+SIZE{fig}", page.text, re.M)
    avg = re.search(rf"^MONTHLY\s+AVERAGE\s+AUM{fig}", page.text, re.M)
    day = re.search(rf"^As\s+on\s+({DATE})\s+PORTFOLIO", page.text, re.M | re.I)
    return end and {"aum_month_end_cr": number(end.group(1)), "aum_monthly_average_cr": avg and number(avg.group(1)),
                    "aum_as_on_date": day and parse_date(day.group(1))}


@rule("aum_month_end_cr", "aum_monthly_average_cr", "aum_as_on_date", amc="Union")
def union_split_labels(page):
    """
    Union (until JUL 2026) prints the labels apart from the figures, each figure with its own sub-label:
    'As on 31st July 2026* : ` 1,467.28 crore' and 'Average for July 2026 : ` 2,384.37 crore'.
    """
    end = re.search(rf"^As\s+on\s+({DATE})\*?\s*:\s*{CUR}\s*({NUM})\s*crore", page.text, re.M | re.I)
    avg = re.search(rf"^Average\s+for\s+{DATE}\s*:\s*{CUR}\s*({NUM})\s*crore", page.text, re.M | re.I)
    return end and {"aum_as_on_date": parse_date(end.group(1)), "aum_month_end_cr": number(end.group(2)),
                    "aum_monthly_average_cr": avg and number(avg.group(1))}


@rule("aum_month_end_cr", "aum_monthly_average_cr", amc="Mirae")
def mirae_columns(page):
    """
    Mirae prints three funds side by side ('Monthly Average AUM (₹ Cr.) 38,692.946 45,665.455 4,787.594'); which
    column is this fund's is not in the text layer, so the generic rules must not take the first figure.
    """
    if re.search(rf"AUM\s*\([^)]*\)\s*{NUM}\s+{NUM}", page.text, re.I):
        return {"aum_month_end_cr": ASK_MODEL, "aum_monthly_average_cr": ASK_MODEL}


@rule("aum_monthly_average_cr")
def monthly_average(page):
    for value, m in _figures(page.text, AVG_LABEL):
        if re.search(r"quarter[^\n]{0,30}$", page.text[max(0, m.start() - 40):m.start()], re.I):
            continue  # PGIM 'Quarterly (Apr-June 2026) Average AUM' comes before the monthly one
        return {"aum_monthly_average_cr": value}


@rule("aum_month_end_cr")
def month_end(page):
    for value, m in _figures(page.text, END_LABEL):
        if re.search(r"(average|avg)\W*$", page.text[max(0, m.start() - 20):m.start()], re.I):
            continue  # the 'AUM' of 'Monthly Average AUM' (Baroda's 'Monthly AUM' is the month-end figure)
        return {"aum_month_end_cr": value}


@rule("aum_as_on_date")
def aum_date(page):
    """
    The date printed for the AUM ('AUM as on ...', Motilal's 'Latest AUM (31-Jan-2026)'), else for the NAV ('NAV AS
    ON MARCH 30, 2026'), else Edelweiss's 'Data as on'. The agent drops a date outside the factsheet's month.
    """
    # (?<!A): Union prints 'AAUM as on Sep. 2026' (a month) above 'AUM as on 30th Sep. 2026'
    for label in (r"(?:(?<![A-Za-z])AUM|Assets\s+Under\s+Management)[#*\s]*(?:\([^)]{0,20}\))?\s*:?\s*as\s+on",
                  r"(?<![A-Za-z])AUM\s*\(", r"NAV\s*:?\s*\(?\s*as\s+on", r"Data\s+as\s+on"):
        m = re.search(rf"(?:{label})\s*:?\s*({DATE})", page.text, re.I)
        if m:
            return {"aum_as_on_date": parse_date(m.group(1))}
    if re.search(r"AUM\s+as\s+on\s+last\s+day", page.text, re.I):  # ABSL: no date printed for the AUM
        return {"aum_as_on_date": NOT_PRINTED}
