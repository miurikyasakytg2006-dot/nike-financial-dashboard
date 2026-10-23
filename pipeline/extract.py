"""Download NIKE 10-K XBRL packages and extract every catalogued metric with provenance.

Usage:  python -m pipeline.extract

Outputs (committed, human-readable):
  data/extracted/facts.csv            one row per value found in a filing (current year and
                                      comparatives), with page, section and evidence text
  data/extracted/sec_api_check.csv    consolidated values compared with the SEC's own
                                      XBRL "company facts" API for the same accession

Downloaded filings are cached in data/cache/ (not committed).
"""
import csv
import datetime
import io
import json
import re
import sys
import urllib.request
import zipfile
from decimal import Decimal
from pathlib import Path

from .ixbrl import Document
from .metrics import REPORTED
from .sources import CIK, FILINGS, filing, ir_package_url

ROOT = Path(__file__).resolve().parent.parent
CACHE = ROOT / "data" / "cache"
OUT = ROOT / "data" / "extracted"
USER_AGENT = "NIKE-FY2020-2026 academic dataset (university coursework; non-commercial)"
MILLION = Decimal(10) ** 6


def _get(url):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=120) as r:
        return r.read()


def filing_document(f):
    """Path to the cached primary iXBRL document, downloading the package if needed."""
    target = CACHE / f["accession"] / f["document"]
    if not target.exists():
        data = _get(ir_package_url(f))
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            if f["document"] not in z.namelist():
                raise RuntimeError(f"{f['document']} not found in XBRL package {f['accession']}")
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(z.read(f["document"]))
    return target


def companyfacts():
    target = CACHE / "companyfacts.json"
    if not target.exists():
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(_get(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{int(CIK):010d}.json"))
    return json.loads(target.read_text())


# ---------------------------------------------------------------- XBRL facts
def _fiscal_year(fact, period):
    if period == "instant":
        d = fact["instant"]
        return int(d[:4]) if d and d[5:] == "05-31" else None
    s, e = fact["start"], fact["end"]
    if s and e and s[5:] == "06-01" and e[5:] == "05-31" and int(e[:4]) - int(s[:4]) == 1:
        return int(e[:4])
    return None


def extract_xbrl(doc, f):
    rows, problems = [], []
    for m in REPORTED:
        spec = m.get("xbrl")
        if not spec:
            continue
        by_fy = {}
        for fact in doc.facts:
            if fact["concept"] not in spec["concepts"] or fact["dims"] != spec["dims"]:
                continue
            fy = _fiscal_year(fact, spec["period"])
            if fy is not None and fact["value"] is not None:
                by_fy.setdefault(fy, []).append(fact)
        for fy, cands in sorted(by_fy.items()):
            values = {c["value"] for c in cands}
            if len(values) > 1:
                problems.append(f"FY{f['fy']} 10-K: {m['id']} FY{fy} tagged with different values {sorted(values)}")
            preferred = [c for c in cands if re.search(spec["section"], c["section"])]
            c = (preferred or cands)[0]
            rows.append(dict(
                filing_fy=f["fy"], accession=f["accession"], metric_id=m["id"], period_fy=fy,
                value=c["value"] * spec["sign"] / MILLION, method="xbrl", concept=c["concept"],
                dims=json.dumps(c["dims"], sort_keys=True) if c["dims"] else "",
                page=c["page"], section=c["section"], evidence=c["displayed"]))
    return rows, problems


# ---------------------------------------------------------------- MD&A tables
def _num(cell):
    s = cell.replace(",", "").replace("$", "").replace("%", "").strip()
    if s in ("", "—", "–", "-"):
        return None
    neg = s.startswith("(")
    v = Decimal(s.strip("()"))
    return -v if neg else v


def _label(cell):
    return re.sub(r"(\(\d+\),?)+:?$", "", cell).strip().upper()


REVENUE_ROWS = {
    "SALES TO WHOLESALE CUSTOMERS": ("nike_brand_wholesale_revenue", "nike_brand_wholesale_pct_change_reported", "nike_brand_wholesale_pct_change_cn"),
    "SALES THROUGH NIKE DIRECT": ("nike_direct_revenue", "nike_direct_pct_change_reported", "nike_direct_pct_change_cn"),
    "TOTAL NIKE BRAND REVENUES": ("nike_brand_revenue", "nike_brand_pct_change_reported", "nike_brand_pct_change_cn"),
    "TOTAL NIKE, INC. REVENUES": ("revenues", "revenues_pct_change_reported", "revenues_pct_change_cn"),
    # Value-only rows, used to cross-check the tagged note values.
    "GLOBAL BRAND DIVISIONS": ("global_brand_divisions_revenue", None, None),
    "CONVERSE": ("converse_revenue", None, None),
    "CORPORATE": ("corporate_revenue", None, None),
}


def extract_mdna_tables(doc, f):
    rows, problems = [], []
    fy = f["fy"]
    table = next((t for t in doc.tables
                  if {"SALES THROUGH NIKE DIRECT", "TOTAL NIKE, INC. REVENUES"} <= {_label(r[0]) for r in t["rows"]}), None)
    if table is None:
        return rows, [f"FY{fy} 10-K: MD&A revenue table not found"]
    header = " ".join(table["rows"][0])
    if f"FISCAL {fy}" not in header or f"FISCAL {fy - 1}" not in header:
        problems.append(f"FY{fy} 10-K: unexpected MD&A revenue table header: {header}")
    seen = {}
    for r in table["rows"]:
        lab = _label(r[0])
        if lab not in REVENUE_ROWS or len(r) != 8:
            continue
        vals = [_num(c) for c in r[1:]]
        if lab in seen:
            if seen[lab] != vals:
                problems.append(f"FY{fy} 10-K: MD&A row {lab} appears twice with different values")
            continue
        seen[lab] = vals
        value_id, pct_id, cn_id = REVENUE_ROWS[lab]
        common = dict(filing_fy=fy, accession=f["accession"], method="mdna_table", concept="", dims="",
                      page=table["page"], section=table["section"], evidence=" | ".join(r))
        rows.append(dict(common, metric_id=value_id, period_fy=fy, value=vals[0]))
        rows.append(dict(common, metric_id=value_id, period_fy=fy - 1, value=vals[1]))
        if pct_id:
            rows.append(dict(common, metric_id=pct_id, period_fy=fy, value=vals[2] / 100))
            rows.append(dict(common, metric_id=cn_id, period_fy=fy, value=vals[3] / 100))
    missing = set(REVENUE_ROWS) - set(seen)
    if missing:
        problems.append(f"FY{fy} 10-K: MD&A revenue rows not found: {sorted(missing)}")

    ebit = next(((t, r) for t in doc.tables for r in t["rows"]
                 if re.match(r"^TOTAL NIKE, INC\. (EARNINGS BEFORE INTEREST AND TAXES|EBIT)$", _label(r[0]))), None)
    if ebit is None:
        problems.append(f"FY{fy} 10-K: MD&A Total NIKE, Inc. EBIT row not found")
    else:
        t, r = ebit
        vals = [_num(c) for c in r[1:]]
        common = dict(filing_fy=fy, accession=f["accession"], metric_id="ebit_reported", method="mdna_table",
                      concept="", dims="", page=t["page"], section=t["section"], evidence=" | ".join(r))
        rows.append(dict(common, period_fy=fy, value=vals[0]))
        rows.append(dict(common, period_fy=fy - 1, value=vals[1]))
    return rows, problems


# ---------------------------------------------------------------- MD&A narrative
SHARE_RE = re.compile(r"(?:represented|to be) approximately (\d+)% of (?:our )?total NIKE Brand revenues for fiscal (\d{4})"
                      r"(?:,? compared to (\d+)% for fiscal (\d{4}))?")
DIGITAL_RE = re.compile(r"(?:NIKE Brand Digital|Digital commerce) sales were \$(\d+\.\d) billion (?:for|in) fiscal (\d{4}) "
                        r"compared to \$(\d+\.\d) billion (?:for|in) fiscal (\d{4})")
DIGITAL_FROM_TO_RE = re.compile(r"NIKE Brand Digital sales of \d+% from \$(\d+\.\d) billion in fiscal (\d{4}) "
                                r"to \$(\d+\.\d) billion in fiscal (\d{4})")


def _quote(text, m):
    start = max(text.rfind("•", 0, m.start()), text.rfind(". ", 0, m.start()) + 1, 0)
    end = text.find(".", m.end())
    return text[start:end + 1 if end >= 0 else len(text)].strip(" •")


def extract_narrative(doc, f):
    rows, problems = [], []
    fy = f["fy"]

    def add(metric_id, period, value, line, m):
        rows.append(dict(filing_fy=fy, accession=f["accession"], metric_id=metric_id, period_fy=period, value=value,
                         method="mdna_text", concept="", dims="", page=line["page"], section=line["section"],
                         evidence=_quote(line["text"], m)))

    share_done, digital_done = set(), False
    for line in doc.lines:
        for m in SHARE_RE.finditer(line["text"]):
            if "NIKE Direct" not in line["text"][:m.end()]:
                continue
            pairs = [(int(m.group(2)), m.group(1))]
            if m.group(3):
                pairs.append((int(m.group(4)), m.group(3)))
            for period, pct in pairs:
                if period not in share_done and period in (fy, fy - 1):
                    share_done.add(period)
                    add("nike_direct_share_stated", period, Decimal(pct) / 100, line, m)
        if digital_done:
            continue
        m = DIGITAL_RE.search(line["text"])
        if m and int(m.group(2)) == fy:
            add("nike_brand_digital_sales", fy, Decimal(m.group(1)), line, m)
            add("nike_brand_digital_sales", int(m.group(4)), Decimal(m.group(3)), line, m)
            digital_done = True
            continue
        m = DIGITAL_FROM_TO_RE.search(line["text"])
        if m and int(m.group(4)) == fy:
            add("nike_brand_digital_sales", fy, Decimal(m.group(3)), line, m)
            add("nike_brand_digital_sales", int(m.group(2)), Decimal(m.group(1)), line, m)
            digital_done = True
    if not digital_done:
        problems.append(f"FY{fy} 10-K: NIKE Brand Digital sales sentence not found")
    if fy not in share_done:
        problems.append(f"FY{fy} 10-K: no statement of NIKE Direct share for fiscal {fy} (recorded as not disclosed)")
    return rows, problems


# ---------------------------------------------------------------- SEC API cross-check
def sec_api_check(rows):
    cf = companyfacts()["facts"]["us-gaap"]
    out = []
    for r in rows:
        if r["method"] != "xbrl" or r["dims"]:
            continue
        spec = next(m for m in REPORTED if m["id"] == r["metric_id"])["xbrl"]
        end = f"{r['period_fy']}-05-31"
        start = f"{r['period_fy'] - 1}-06-01"
        matches = [x for x in cf.get(r["concept"].split(":")[1], {}).get("units", {}).get("USD", [])
                   if x["accn"] == r["accession"] and x["end"] == end
                   and (spec["period"] == "instant" or x.get("start") == start)]
        sec_values = sorted({Decimal(str(x["val"])) * spec["sign"] / MILLION for x in matches})
        out.append(dict(filing_fy=r["filing_fy"], accession=r["accession"], metric_id=r["metric_id"],
                        period_fy=r["period_fy"], extracted_value=r["value"],
                        sec_api_value=";".join(str(v) for v in sec_values),
                        result="match" if sec_values == [r["value"]] else ("not in API" if not sec_values else "MISMATCH")))
    return out


def _fmt(v):
    if isinstance(v, Decimal):
        v = v.normalize()
        return format(v, "f")
    return "" if v is None else v


def write_csv(path, rows, fields):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow({k: _fmt(r.get(k)) for k in fields})


def main():
    all_rows, problems = [], []
    for fy, *_ in FILINGS:
        f = filing(fy)
        doc = Document(filing_document(f))
        for extractor in (extract_xbrl, extract_mdna_tables, extract_narrative):
            rows, probs = extractor(doc, f)
            all_rows += rows
            problems += probs
        print(f"FY{fy}: {f['accession']} parsed, {len(doc.facts)} tagged facts")
    fields = ["filing_fy", "accession", "metric_id", "period_fy", "value", "method", "concept", "dims",
              "page", "section", "evidence"]
    write_csv(OUT / "facts.csv", all_rows, fields)
    (OUT / "meta.json").write_text(json.dumps({"extracted_on": datetime.date.today().isoformat(),
                                               "filings": [a for _, a, *_ in FILINGS]}, indent=2) + "\n")
    check = sec_api_check(all_rows)
    write_csv(OUT / "sec_api_check.csv", check, list(check[0].keys()))
    mism = [c for c in check if c["result"] != "match"]
    print(f"{len(all_rows)} extracted values -> data/extracted/facts.csv")
    print(f"SEC API cross-check: {len(check) - len(mism)}/{len(check)} match")
    for c in mism:
        print("  ", c)
    for p in problems:
        print("NOTE:", p)
    blocking = [p for p in problems if "recorded as not disclosed" not in p] + [c for c in mism if c["result"] == "MISMATCH"]
    return 1 if blocking else 0


if __name__ == "__main__":
    sys.exit(main())
