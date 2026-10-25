"""Registry of the official NIKE, Inc. filings used as primary sources (Form 10-K and Form 10-Q).

Each filing is identified by its SEC accession number. The SEC EDGAR URLs are the canonical
citations. The documents and XBRL packages are downloaded from NIKE's investor-relations
mirror of the same EDGAR filings (www.sec.gov requires a contact e-mail in the User-Agent
for automated access).

Scope: fiscal 2019 to fiscal 2021 (1 June 2018 – 31 May 2021), by year and by quarter.
The reports for fiscal 2022 and fiscal 2023 are read only to see whether NIKE later
presented the fiscal 2019–2021 figures differently.
"""

CIK = "320187"
COMPANY = "NIKE, Inc."
FISCAL_YEAR_END = "May 31"

SCOPE_YEARS = [2019, 2020, 2021]
QUARTERS = [(fy, q) for fy in SCOPE_YEARS for q in (1, 2, 3, 4)]

FILINGS = [
    # id, form, fiscal year, quarter, accession, primary document, XBRL instance (filings before inline XBRL), filed,
    # period end, role ("scope": source of the figures; "check": read only for later presentations of them)
    ("10-K FY2019", "10-K", 2019, None, "0000320187-19-000051", "nke-531201910k.htm", "nke-20190531.xml", "2019-07-23", "2019-05-31", "scope"),
    ("10-Q Q1 FY2019", "10-Q", 2019, 1, "0000320187-18-000213", "nke-08312018x10q.htm", "nke-20180831.xml", "2018-10-05", "2018-08-31", "scope"),
    ("10-Q Q2 FY2019", "10-Q", 2019, 2, "0000320187-19-000007", "nke-11302018x10q.htm", "nke-20181130.xml", "2019-01-08", "2018-11-30", "scope"),
    ("10-Q Q3 FY2019", "10-Q", 2019, 3, "0000320187-19-000030", "nke-02282019x10q.htm", "nke-20190228.xml", "2019-04-04", "2019-02-28", "scope"),
    ("10-K FY2020", "10-K", 2020, None, "0000320187-20-000047", "nke-531202010k.htm", None, "2020-07-24", "2020-05-31", "scope"),
    ("10-Q Q1 FY2020", "10-Q", 2020, 1, "0000320187-19-000071", "nke-831201910q.htm", None, "2019-10-04", "2019-08-31", "scope"),
    ("10-Q Q2 FY2020", "10-Q", 2020, 2, "0000320187-20-000004", "nke-1130201910q.htm", None, "2020-01-07", "2019-11-30", "scope"),
    ("10-Q Q3 FY2020", "10-Q", 2020, 3, "0000320187-20-000018", "nke-229202010q.htm", None, "2020-04-07", "2020-02-29", "scope"),
    ("10-K FY2021", "10-K", 2021, None, "0000320187-21-000028", "nke-20210531.htm", None, "2021-07-20", "2021-05-31", "scope"),
    ("10-Q Q1 FY2021", "10-Q", 2021, 1, "0000320187-20-000061", "nke-20200831.htm", None, "2020-10-08", "2020-08-31", "scope"),
    ("10-Q Q2 FY2021", "10-Q", 2021, 2, "0000320187-21-000003", "nke-20201130.htm", None, "2021-01-05", "2020-11-30", "scope"),
    ("10-Q Q3 FY2021", "10-Q", 2021, 3, "0000320187-21-000012", "nke-20210228.htm", None, "2021-04-02", "2021-02-28", "scope"),
    ("10-Q Q1 FY2022", "10-Q", 2022, 1, "0000320187-21-000045", "nke-20210831.htm", None, "2021-10-05", "2021-08-31", "check"),
    ("10-Q Q2 FY2022", "10-Q", 2022, 2, "0000320187-22-000004", "nke-20211130.htm", None, "2022-01-06", "2021-11-30", "check"),
    ("10-Q Q3 FY2022", "10-Q", 2022, 3, "0000320187-22-000017", "nke-20220228.htm", None, "2022-04-05", "2022-02-28", "check"),
    ("10-K FY2022", "10-K", 2022, None, "0000320187-22-000038", "nke-20220531.htm", None, "2022-07-21", "2022-05-31", "check"),
    ("10-K FY2023", "10-K", 2023, None, "0000320187-23-000039", "nke-20230531.htm", None, "2023-07-20", "2023-05-31", "check"),
]
_FIELDS = ("id", "form", "fy", "quarter", "accession", "document", "instance", "filed", "period_end", "role")
QUARTER_NAMES = {1: "first", 2: "second", 3: "third"}


def all_filings():
    return [dict(zip(_FIELDS, row)) for row in FILINGS]


def filing(fid):
    for row in FILINGS:
        if row[0] == fid:
            return dict(zip(_FIELDS, row))
    raise KeyError(f"No filing registered as {fid}")


def annual_filing(fy):
    return filing(f"10-K FY{fy}")


def quarterly_filing(fy, q):
    return filing(f"10-Q Q{q} FY{fy}")


def source_filing(period):
    """The filing whose own figures a period's values come from: FY2020 -> 10-K FY2020, FY2020Q2 -> 10-Q Q2 FY2020.
    The fourth quarter has no 10-Q; its values are the year (10-K) minus the first nine months (third-quarter 10-Q)."""
    fy, q = parse_period(period)
    if q is None or q == 4:
        return annual_filing(fy)
    return quarterly_filing(fy, q)


def parse_period(period):
    """'FY2020' -> (2020, None); 'FY2020Q3' -> (2020, 3)."""
    return (int(period[2:6]), int(period[7]) if len(period) > 6 and period[6] == "Q" else None)


def sec_document_url(f):
    return f"https://www.sec.gov/Archives/edgar/data/{CIK}/{f['accession'].replace('-', '')}/{f['document']}"


def sec_index_url(f):
    return f"https://www.sec.gov/Archives/edgar/data/{CIK}/{f['accession'].replace('-', '')}/{f['accession']}-index.htm"


def ir_package_url(f):
    a = f["accession"]
    return f"https://investors.nike.com/sec-filings/all-sec-filings/content/{a}/{a}-xbrl.zip"


def ir_document_url(f):
    return f"https://investors.nike.com/sec-filings/all-sec-filings/content/{f['accession']}/{f['document']}"


def document_title(f):
    if f["form"] == "10-K":
        return f"NIKE, Inc. Form 10-K for the fiscal year ended May 31, {f['fy']}"
    d = f["period_end"]
    month = {"08": "August", "11": "November", "02": "February"}[d[5:7]]
    return f"NIKE, Inc. Form 10-Q for the quarterly period ended {month} {int(d[8:])}, {d[:4]}"
