"""Registry of the official NIKE, Inc. Form 10-K filings used as primary sources.

Each filing is identified by its SEC accession number. The SEC EDGAR URLs are the
canonical citations. The XBRL packages are downloaded from NIKE's investor-relations
mirror of the same EDGAR filings (www.sec.gov requires a contact e-mail in the
User-Agent for automated access); values are cross-checked against the SEC's own
XBRL "company facts" API, which is keyed by the same accession numbers.
"""

CIK = "320187"
COMPANY = "NIKE, Inc."
FISCAL_YEAR_END = "May 31"

FILINGS = [
    # fy, accession, primary iXBRL document, filing date, period end
    (2020, "0000320187-20-000047", "nke-531202010k.htm", "2020-07-24", "2020-05-31"),
    (2021, "0000320187-21-000028", "nke-20210531.htm", "2021-07-20", "2021-05-31"),
    (2022, "0000320187-22-000038", "nke-20220531.htm", "2022-07-21", "2022-05-31"),
    (2023, "0000320187-23-000039", "nke-20230531.htm", "2023-07-20", "2023-05-31"),
    (2024, "0000320187-24-000044", "nke-20240531.htm", "2024-07-25", "2024-05-31"),
    (2025, "0000320187-25-000047", "nke-20250531.htm", "2025-07-17", "2025-05-31"),
    (2026, "0000320187-26-000088", "nke-20260531.htm", "2026-07-15", "2026-05-31"),
]

FIRST_FY = 2020
LAST_FY = 2026


def filing(fy):
    for row in FILINGS:
        if row[0] == fy:
            return dict(zip(("fy", "accession", "document", "filed", "period_end"), row))
    raise KeyError(f"No 10-K registered for FY{fy}")


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
    return f"NIKE, Inc. Form 10-K for the fiscal year ended May 31, {f['fy']}"
