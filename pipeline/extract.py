"""Extract every catalogued metric, with provenance, from NIKE's 10-K and 10-Q filings and fourth-quarter earnings releases.

Usage:  python -m pipeline.extract

Output (committed, human-readable):
  data/extracted/facts.csv   one row per value found in a filing (the period it reports and the comparative
                             periods printed beside it), with page, section and the text shown in the document

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

from .ixbrl import Document, displayed
from .metrics import REGIONS, REPORTED
from .quotes import QUOTES
from .sources import all_filings, ir_document_url, ir_package_url

ROOT = Path(__file__).resolve().parent.parent
CACHE = ROOT / "data" / "cache"
OUT = ROOT / "data" / "extracted"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) academic dataset (university coursework; non-commercial)"
MILLION = Decimal(10) ** 6
FIELDS = ["filing", "accession", "metric_id", "period", "value", "method", "concept", "dims", "page", "section", "evidence"]


def _get(url):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=180) as r:
        return r.read()


def filing_paths(f):
    """Cached primary document (and XBRL instance for older filings), downloaded if needed."""
    folder = CACHE / f["accession"]
    doc = folder / f["document"]
    inst = folder / f["instance"] if f["instance"] else None
    if f["form"] == "8-K":     # an earnings release: one HTML exhibit, no XBRL package
        if not doc.exists():
            folder.mkdir(parents=True, exist_ok=True)
            doc.write_bytes(_get(ir_document_url(f)))
        return doc, None
    if not doc.exists() or (inst and not inst.exists()):
        folder.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(io.BytesIO(_get(ir_package_url(f)))) as z:
            z.extractall(folder)
        if not doc.exists():   # filings before inline XBRL: the document is not in the XBRL package
            doc.write_bytes(_get(ir_document_url(f)))
    return doc, inst


# ---------------------------------------------------------------- periods
QUARTER_END = {"08": 1, "11": 2, "02": 3, "05": 4}


def _fy(date):
    return int(date[:4]) + (1 if int(date[5:7]) >= 6 else 0)


def period_of(fact, kind):
    """Period id of a fact: FY2020 (year or 31 May balance), FY2020Q1..Q4 (three months or quarter-end balance),
    FY2020H1 (six months), FY2020M9 (nine months); None for anything else."""
    if kind == "instant":
        d = fact["instant"]
        if not d or d[5:7] not in QUARTER_END or d[8:] not in ("28", "29", "30", "31"):
            return None
        q = QUARTER_END[d[5:7]]
        return f"FY{_fy(d)}" if q == 4 else f"FY{_fy(d)}Q{q}"
    s, e = fact["start"], fact["end"]
    if not s or not e or s[8:] != "01" or e[5:7] not in QUARTER_END:
        return None
    months = (int(e[:4]) - int(s[:4])) * 12 + int(e[5:7]) - int(s[5:7]) + 1
    fy = _fy(e)
    if months == 3:
        return f"FY{fy}Q{QUARTER_END[e[5:7]]}"
    if s[5:7] == "06":
        return {12: f"FY{fy}", 6: f"FY{fy}H1", 9: f"FY{fy}M9"}.get(months)
    return None


def _local(qname):
    return qname.split(":")[-1]


def _dims(dims):
    return {_local(k): _local(v) for k, v in dims.items()}


# ---------------------------------------------------------------- XBRL facts
def extract_xbrl(doc, f):
    rows, problems = [], []
    for m in REPORTED:
        spec = m.get("xbrl")
        if not spec:
            continue
        concepts = set(spec["concepts"])
        by_period = {}
        for fact in doc.facts:
            if _local(fact["concept"]) not in concepts or _dims(fact["dims"]) not in spec["dims"]:
                continue
            p = period_of(fact, spec["period"])
            if p is not None and fact["value"] is not None:
                by_period.setdefault(p, []).append(fact)
        for p, cands in sorted(by_period.items()):
            values = {c["value"] for c in cands}
            if len(values) > 1:
                problems.append(f"{f['id']}: {m['id']} {p} tagged with different values {sorted(values)}")
            preferred = [c for c in cands if re.search(spec["section"], c["section"], re.IGNORECASE)]
            c = (preferred or cands)[0]
            value = c["value"] * spec["sign"]
            page, section, evidence = c["page"], c["section"], c["displayed"]
            if not doc.inline:
                # Filings before inline XBRL: find the same amount on the printed row the metric belongs to.
                shown = displayed(value)
                hit = (doc.find_row(spec["section"], spec["row"], shown)
                       or (value < 0 and doc.find_row(spec["section"], spec["row"], shown[:-1]))
                       or (value == 0 and doc.find_row(spec["section"], spec["row"], "—")))
                if hit:
                    t, _ = hit
                    page, section, evidence = t["page"], t["section"], shown
                else:
                    problems.append(f"{f['id']}: {m['id']} {p} = {shown} not found on a printed '{spec['row']}' row")
            rows.append(dict(filing=f["id"], accession=f["accession"], metric_id=m["id"], period=p,
                             value=value / MILLION, method="xbrl", concept=c["concept"],
                             dims=json.dumps(c["dims"], sort_keys=True) if c["dims"] else "",
                             page=page, section=section, evidence=evidence))
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


def _columns(f, n_cells, with_cn=True):
    """Period of each value cell of an MD&A table row, by report type: ('value'|'pct'|'pct_cn', period) per cell."""
    fy, q = f["fy"], f["quarter"]
    if f["form"] == "10-K":
        cur, prior, prior2 = f"FY{fy}", f"FY{fy - 1}", f"FY{fy - 2}"
        # The last two columns give the previous year's change again (against the year before it).
        cols = ([("value", cur), ("value", prior), ("pct", cur), ("pct_cn", cur), ("value", prior2), ("pct", prior), ("pct_cn", prior)]
                if with_cn else [("value", cur), ("value", prior), ("pct", cur), ("value", prior2), (None, None)])
    else:
        cur, prior = f"FY{fy}Q{q}", f"FY{fy - 1}Q{q}"
        three = [("value", cur), ("value", prior), ("pct", cur)] + ([("pct_cn", cur)] if with_cn else [])
        cols = three
        if q in (2, 3):
            ytd = "H1" if q == 2 else "M9"
            cols = three + [("value", f"FY{fy}{ytd}"), ("value", f"FY{fy - 1}{ytd}"), ("pct", f"FY{fy}{ytd}")] + \
                ([("pct_cn", f"FY{fy}{ytd}")] if with_cn else [])
    return cols if len(cols) == n_cells else None


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
SEGMENT_ROWS = {   # % change columns only; the amounts are read from the notes and the revenue table already
    "TOTAL NIKE BRAND": (None, "nike_brand_pct_change_reported", "nike_brand_pct_change_cn"),
    "TOTAL NIKE, INC. REVENUES": (None, "revenues_pct_change_reported", "revenues_pct_change_cn"),
}
MDNA = re.compile(r"^ITEM [27]\.", re.IGNORECASE)


def _row_values(f, table, r, ids, with_cn=True):
    """Facts for one MD&A table row: ids = (value metric, % change metric, currency-neutral % change metric)."""
    cols = _columns(f, len(r) - 1, with_cn)
    if cols is None:
        return None
    out = []
    common = dict(filing=f["id"], accession=f["accession"], method="mdna_table", concept="", dims="",
                  page=table["page"], section=table["section"], evidence=" | ".join(r))
    for (what, period), cell in zip(cols, r[1:]):
        mid = {"value": ids[0], "pct": ids[1] if len(ids) > 1 else None, "pct_cn": ids[2] if len(ids) > 2 else None}.get(what)
        if not mid:
            continue
        v = _num(cell)
        if v is None:
            if what != "value":     # a dash in a % column: no meaningful change
                continue
            v = Decimal(0)          # a dash in an amount column: zero
        out.append(dict(common, metric_id=mid, period=period, value=v / 100 if what != "value" else v))
    return out


def extract_mdna_tables(doc, f):
    rows, problems = [], []
    mdna_tables = [t for t in doc.tables if MDNA.search(t["section"])]
    table = next((t for t in mdna_tables
                  if {"SALES THROUGH NIKE DIRECT", "TOTAL NIKE, INC. REVENUES"} <= {_label(r[0]) for r in t["rows"]}), None)
    if table is None:
        return rows, [f"{f['id']}: MD&A revenue table not found"]
    header = " ".join(" ".join(r) for r in table["rows"][:2]).upper()
    end_year = f["period_end"][:4]
    if (f["form"] == "10-K" and f"FISCAL {f['fy']}" not in header) or end_year not in header:
        problems.append(f"{f['id']}: unexpected MD&A revenue table header: {header}")
    seen = {}
    for r in table["rows"]:
        lab = _label(r[0])
        if lab == "TOTAL NIKE BRAND":
            lab = "TOTAL NIKE BRAND REVENUES"
        if lab not in REVENUE_ROWS:
            continue
        got = _row_values(f, table, r, REVENUE_ROWS[lab])
        if got is None:
            problems.append(f"{f['id']}: MD&A revenue row {lab} has {len(r) - 1} cells")
            continue
        vals = [g["value"] for g in got]
        if lab in seen:
            if seen[lab] != vals:
                problems.append(f"{f['id']}: MD&A row {lab} appears twice with different values")
            continue
        seen[lab] = vals
        rows += got
    missing = set(REVENUE_ROWS) - set(seen)
    if missing:
        problems.append(f"{f['id']}: MD&A revenue rows not found: {sorted(missing)}")

    # The revenues-by-segment table of the Operating Segments discussion prints the NIKE Brand and NIKE, Inc. % changes
    # a second time, on another page: a second place to read them.
    seg = next((t for t in mdna_tables
                if {"GREATER CHINA", "TOTAL NIKE BRAND", "TOTAL NIKE, INC. REVENUES"} <= {_label(r[0]) for r in t["rows"]}
                and "SALES THROUGH NIKE DIRECT" not in {_label(r[0]) for r in t["rows"]}), None)
    if seg is None:
        problems.append(f"{f['id']}: MD&A revenues-by-segment table not found")
    else:
        for r in seg["rows"]:
            ids = SEGMENT_ROWS.get(_label(r[0]))
            if ids is None:
                continue
            got = _row_values(f, seg, r, ids)
            if got is None:
                problems.append(f"{f['id']}: MD&A segment row {r[0]} has {len(r) - 1} cells")
                continue
            rows += [dict(g, method="mdna_segment_table") for g in got]

    ebit = next(((t, r) for t in mdna_tables for r in t["rows"]
                 if re.match(r"^TOTAL NIKE, INC\. (EARNINGS BEFORE INTEREST AND TAXES|EBIT)$", _label(r[0]))), None)
    if ebit is None:
        problems.append(f"{f['id']}: MD&A Total NIKE, Inc. EBIT row not found")
    else:
        got = _row_values(f, ebit[0], ebit[1], ("ebit_reported",), with_cn=False)
        if got is None:
            problems.append(f"{f['id']}: MD&A EBIT row has {len(ebit[1]) - 1} cells")
        else:
            rows += got

    # One table per geographic segment, right after the segment's name.
    names = {en.upper(): rid for rid, _, _, en, _ in REGIONS}
    found = set()
    for t in mdna_tables:
        rid = names.get(t["before"][-1].upper()) if t["before"] else None
        labs = [_label(r[0]) for r in t["rows"]]
        if rid is None or "SALES THROUGH NIKE DIRECT" not in labs or rid in found:
            continue
        found.add(rid)
        ids = {"SALES TO WHOLESALE CUSTOMERS": (f"{rid}_wholesale_revenue",), "SALES THROUGH NIKE DIRECT": (f"{rid}_direct_revenue",),
               "TOTAL REVENUES": (f"{rid}_revenue",), "EARNINGS BEFORE INTEREST AND TAXES": (f"{rid}_ebit",)}
        squeezed = {k.replace(" ", ""): k for k in ids}   # one report prints "EARNINGS BEFOREINTEREST AND TAXES"
        done = set()
        for r in t["rows"]:
            lab = squeezed.get(_label(r[0]).replace(" ", ""))
            if lab is None or lab in done:
                continue
            done.add(lab)
            got = _row_values(f, t, r, ids[lab], with_cn=lab != "EARNINGS BEFORE INTEREST AND TAXES")
            if got is None:
                problems.append(f"{f['id']}: MD&A {rid} row {lab} has {len(r) - 1} cells")
            else:
                rows += got
        if len(done) != 4:
            problems.append(f"{f['id']}: MD&A table for {rid} lacks rows {sorted(set(ids) - done)}")
    if len(found) != 4:
        problems.append(f"{f['id']}: MD&A regional tables found for {sorted(found)} only")
    return rows, problems


def extract_quarterly_table(doc, f):
    """'Selected Quarterly Financial Data (Unaudited)' of the 10-K (fiscal 2019 and 2020 reports only)."""
    rows = []
    table = next((t for t in doc.tables if "SELECTED QUARTERLY" in t["section"].upper() and len(t["rows"]) > 4), None)
    if table is None:
        return rows, []
    fy = f["fy"]
    periods = [p for q in (1, 2, 3, 4) for p in (f"FY{fy}Q{q}", f"FY{fy - 1}Q{q}")]
    wanted = {"REVENUES": "revenues", "GROSS PROFIT": "gross_profit", "NET INCOME (LOSS)": "net_income", "NET INCOME": "net_income"}
    problems = []
    for r in table["rows"]:
        mid = wanted.get(_label(r[0]))
        if not mid:
            continue
        if len(r) != 9:
            problems.append(f"{f['id']}: quarterly data row {r[0]} has {len(r) - 1} cells")
            continue
        for period, cell in zip(periods, r[1:]):
            rows.append(dict(filing=f["id"], accession=f["accession"], metric_id=mid, period=period, value=_num(cell),
                             method="quarterly_table", concept="", dims="", page=table["page"], section=table["section"],
                             evidence=" | ".join(r)))
    return rows, problems


# ---------------------------------------------------------------- MD&A narrative
PART = {"first quarter": "Q1", "second quarter": "Q2", "third quarter": "Q3", "first six months": "H1",
        "first nine months": "M9"}
AMOUNT = re.compile(r"\$([\d,]+(?:\.\d)?) (million|billion)")
PCT_AMOUNT = re.compile(r"(\d+)%")


def _periods(text, fy_default=None):
    """Periods named in a clause: 'the second quarter and first six months of fiscal 2019' -> [FY2019Q2, FY2019H1];
    'fiscal 2019' -> [FY2019]; 'both the first quarter of fiscal 2019 and fiscal 2018' -> [FY2019Q1, FY2018Q1]."""
    m = re.search(r"both the (first|second|third) quarter of fiscal (\d{4}) and fiscal (\d{4})", text)
    if m:
        q = PART[f"{m.group(1)} quarter"]
        return [f"FY{m.group(2)}{q}", f"FY{m.group(3)}{q}"]
    parts = re.findall(r"(first quarter|second quarter|third quarter|first six months|first nine months)", text)
    years = re.findall(r"fiscal (\d{4})", text)
    if parts and years:
        return [f"FY{years[-1]}{PART[p]}" for p in parts]
    if years:
        return [f"FY{y}" for y in years]
    return []


def _amount_b(num, unit):
    v = Decimal(num.replace(",", ""))
    return v / 1000 if unit == "million" else v


def _sentence(text, start):
    s = text.rfind(". ", 0, start)
    s = 0 if s < 0 else s + 2
    e = text.find(".", start)
    while e != -1 and e + 1 < len(text) and text[e + 1].isdigit():   # decimal point, not the end of the sentence
        e = text.find(".", e + 1)
    return text[s:e + 1 if e >= 0 else len(text)].strip()


def parse_pairs(sentence, pattern):
    """(period, value) pairs from a sentence 'X were A [and B] for <periods>, [respectively,] compared to C [and D]
    for <periods | the respective prior year periods>'."""
    if " compared to " in sentence:
        left, right = sentence.split(" compared to ", 1)
    else:
        left, right = sentence, ""
    cur_periods = _periods(left)
    cur_vals = pattern(left)
    out = list(zip(cur_periods, cur_vals))
    if len(cur_periods) == 2 and len(cur_vals) == 2 and "both" in left:
        out = [(cur_periods[0], cur_vals[0]), (cur_periods[1], cur_vals[1])]
    elif len(cur_periods) == 2 and len(cur_vals) == 1 and "both" in left:
        out = [(cur_periods[0], cur_vals[0]), (cur_periods[1], cur_vals[0])]
    if right:
        prior_vals = pattern(right)
        prior_periods = _periods(right)
        if not prior_periods and "prior year periods" in right:
            prior_periods = [p.replace(p[2:6], str(int(p[2:6]) - 1), 1) for p in cur_periods]
        out += list(zip(prior_periods, prior_vals))
    return out


def extract_narrative(doc, f):
    rows, problems = [], []
    own = {f"FY{f['fy']}"} if f["form"] == "10-K" else {f"FY{f['fy']}Q{f['quarter']}"}

    def add(metric_id, period, value, line, sentence):
        rows.append(dict(filing=f["id"], accession=f["accession"], metric_id=metric_id, period=period, value=value,
                         method="mdna_text", concept="", dims="", page=line["page"], section=line["section"],
                         evidence=sentence))

    done = {"share": set(), "digital": set(), "direct_cn": set()}
    for line in doc.lines:
        if not MDNA.search(line["section"]):
            continue
        text = line["text"]
        # The NIKE Brand discussion states the NIKE Direct currency-neutral change of the report's own period in a
        # sentence, ahead of the regional sections (which use the same words for their own NIKE Direct sales).
        for m in re.finditer(r"On a currency-neutral basis, NIKE Direct revenues (increased|decreased|grew|declined|rose|fell) "
                             r"(\d+)%(?! and)", text):
            s = _sentence(text, m.start())
            if done["direct_cn"] or re.search(r"six months|nine months", s):
                continue
            (period,) = own
            done["direct_cn"].add(period)
            add("nike_direct_pct_change_cn", period, _pct(m.group(1), m.group(2)), line, s)
        for m in re.finditer(r"NIKE Direct revenues represented approximately", text):
            s = _sentence(text, m.start())
            for period, pct in parse_pairs(s, lambda x: [Decimal(p) / 100 for p in PCT_AMOUNT.findall(x)]):
                if period not in done["share"]:
                    done["share"].add(period)
                    add("nike_direct_share_stated", period, pct, line, s)
        for m in re.finditer(r"(?:NIKE Brand Digital|Digital commerce|Digital) sales were \$", text):
            s = _sentence(text, m.start())
            for period, v in parse_pairs(s, lambda x: [_amount_b(n, u) for n, u in AMOUNT.findall(x)]):
                if period not in done["digital"]:
                    done["digital"].add(period)
                    add("nike_brand_digital_sales", period, v, line, s)
    for kind in ("share", "digital", "direct_cn"):
        if not own & done[kind] and f["role"] == "scope":
            problems.append(f"{f['id']}: no {kind} sentence for {sorted(own)}")
    return rows, problems


INVENTORY_NOTE = re.compile(r"Inventory balances of \$([\d,]+) million and \$([\d,]+) million (?:at|as of) ")


def extract_inventory_note(doc, f):
    """The inventories note repeats the balance sheet amounts in a sentence: a second place to read them."""
    end = datetime.date.fromisoformat(f["period_end"])
    when = f"{end:%B} {end.day}, {end.year}"
    own = f"FY{f['fy']}" if f["form"] == "10-K" else f"FY{f['fy']}Q{f['quarter']}"
    for line in doc.lines:
        m = INVENTORY_NOTE.search(line["text"])
        if not m or not re.match(r"NOTE \d+ ?[—–-] ?INVENTORIES", line["section"], re.IGNORECASE):
            continue
        s = _sentence(line["text"], m.start())
        if f["form"] == "10-Q" and when not in s or f["form"] == "10-K" and f"May 31, {f['fy']}" not in s:
            return [], [f"{f['id']}: inventories note does not name {when}: {s}"]
        common = dict(filing=f["id"], accession=f["accession"], metric_id="inventories", method="note_text", concept="",
                      dims="", page=line["page"], section=line["section"], evidence=s)
        return [dict(common, period=own, value=Decimal(m.group(1).replace(",", ""))),
                dict(common, period=f"FY{f['fy'] - 1}", value=Decimal(m.group(2).replace(",", "")))], []
    # The 10-Ks for fiscal 2019 and 2020 word this note differently; their balances are confirmed by the next report.
    return [], [f"{f['id']}: inventories note sentence not found"] if f["form"] == "10-Q" and f["role"] == "scope" else []


def extract_quotes(doc, f):
    """The passages of pipeline.quotes that come from this report, verbatim, with page and section."""
    rows, problems = [], []
    for qt in QUOTES:
        if qt["filing"] != f["id"]:
            continue
        line = next((ln for ln in doc.lines if qt["start"] in ln["text"]), None)
        if line is None:
            problems.append(f"{f['id']}: quote '{qt['start']}' not found")
            continue
        text, pos, parts = line["text"], line["text"].index(qt["start"]), []
        for _ in range(qt["sentences"]):
            s = _sentence(text, pos)
            parts.append(s)
            pos = text.find(s, pos) + len(s) + 1
        rows.append(dict(id=qt["id"], filing=f["id"], page=line["page"], section=line["section"], text=" ".join(parts)))
    return rows, problems


# ---------------------------------------------------------------- fourth-quarter earnings releases (Form 8-K)
# The release of each fiscal year gives the fourth quarter's three-month figures (income statement, revenues by
# segment, EBIT) beside the same quarter of the year before, and the quarter's % change, reported and excluding
# currency changes. Its tables have no page numbers; the section is the table title.
RELEASE_INCOME = {
    "REVENUES": "revenues", "COST OF SALES": "cost_of_sales", "GROSS PROFIT": "gross_profit",
    "DEMAND CREATION EXPENSE": "demand_creation_expense", "OPERATING OVERHEAD EXPENSE": "operating_overhead_expense",
    "TOTAL SELLING AND ADMINISTRATIVE EXPENSE": "total_selling_admin_expense",
    "INTEREST EXPENSE (INCOME), NET": "interest_expense_net", "OTHER (INCOME) EXPENSE, NET": "other_expense_net",
    "INCOME BEFORE INCOME TAXES": "income_before_taxes", "INCOME TAX EXPENSE": "income_tax_expense", "NET INCOME": "net_income",
}
RELEASE_SEGMENTS = {   # divisional revenues: (value metric, % change metric, % change excluding currency metric)
    "GLOBAL BRAND DIVISIONS": ("global_brand_divisions_revenue", None, None),
    "TOTAL NIKE BRAND": ("nike_brand_revenue", "nike_brand_pct_change_reported", "nike_brand_pct_change_cn"),
    "CONVERSE": ("converse_revenue", None, None),
    "CORPORATE": ("corporate_revenue", None, None),
    "TOTAL NIKE, INC. REVENUES": ("revenues", "revenues_pct_change_reported", "revenues_pct_change_cn"),
}
RELEASE_REGIONS = {en.upper(): rid for rid, _, _, en, _ in REGIONS}
NUMBER = re.compile(r"\(?-?[\d,]+(?:\.\d+)?\)?|—|–")


def _release_label(cell):
    """Row label without footnote digits and without the '(loss)' / '(benefit)' that NIKE adds in a loss quarter."""
    s = re.sub(r"\d+$", "", cell).strip().upper()
    return re.sub(r"\s*\((LOSS|BENEFIT)\)\s*", " ", s).strip()


def _release_tables(path):
    """{title: [(label, figures, printed row)]} for the release tables, the title being the first line naming one."""
    from lxml import html as lxml_html
    doc = lxml_html.fromstring(path.read_bytes().decode("utf-8", "replace"))
    titles = ("CONSOLIDATED STATEMENTS OF INCOME", "DIVISIONAL REVENUES", "EARNINGS BEFORE INTEREST AND TAXES")
    tables = {}
    for t in doc.iter("table"):
        text = " ".join(t.text_content().split()).upper()
        title = next((x for x in titles if x in text and "THREE MONTHS ENDED" in text), None)
        if title is None or title in tables:
            continue
        rows = []
        for tr in t.iter("tr"):
            cells = [c for c in (" ".join(td.text_content().split()) for td in tr.xpath("./td|./th")) if c and c != "$"]
            if not cells or NUMBER.fullmatch(cells[0]):
                continue
            rows.append((cells[0], NUMBER.findall(" ".join(cells[1:])), " | ".join(cells)))
        tables[title] = rows
    narrative = " ".join(doc.text_content().split())
    return tables, narrative


def _pct(word, n):
    return Decimal(n) / 100 * (-1 if word.lower() in ("decreased", "fell", "declined", "down", "were down") else 1)


def extract_release(path, f):
    """Facts from a fourth-quarter earnings release: three-month figures of the quarter and of the same quarter a year
    earlier, the quarter's % changes, and the narrative sentences that state them again."""
    fy = f["fy"]
    cur, prior = f"FY{fy}Q4", f"FY{fy - 1}Q4"
    tables, text = _release_tables(path)
    rows, problems = [], []
    common = dict(filing=f["id"], accession=f["accession"], concept="", dims="", page=None)

    def add(mid, period, value, section, evidence, method="q4_release"):
        rows.append(dict(common, metric_id=mid, period=period, value=value, method=method, section=section, evidence=evidence))

    income = tables.get("CONSOLIDATED STATEMENTS OF INCOME", [])
    for label, nums, printed in income:
        mid = RELEASE_INCOME.get(_release_label(label))
        if mid and len(nums) >= 2:
            add(mid, cur, _num(nums[0]), "CONSOLIDATED STATEMENTS OF INCOME", printed)
            add(mid, prior, _num(nums[1]), "CONSOLIDATED STATEMENTS OF INCOME", printed)
            if mid == "revenues" and len(nums) >= 3 and _num(nums[2]) is not None:   # the quarter's % change, again
                add("revenues_pct_change_reported", cur, _num(nums[2]) / 100, "CONSOLIDATED STATEMENTS OF INCOME", printed,
                    "q4_release_repeat")
    region, done = None, False
    for label, nums, printed in tables.get("DIVISIONAL REVENUES", []):
        key = _release_label(label)
        if done:
            # Below the NIKE, Inc. total the table repeats NIKE Brand by product type, ending with the same total.
            if key == "TOTAL NIKE BRAND REVENUES" and len(nums) >= 4:
                add("nike_brand_revenue", cur, _num(nums[0]), "DIVISIONAL REVENUES", printed, "q4_release_repeat")
                add("nike_brand_pct_change_reported", cur, _num(nums[2]) / 100, "DIVISIONAL REVENUES", printed, "q4_release_repeat")
                add("nike_brand_pct_change_cn", cur, _num(nums[3]) / 100, "DIVISIONAL REVENUES", printed, "q4_release_repeat")
                break
            continue
        if key in RELEASE_REGIONS and not nums:
            region = RELEASE_REGIONS[key]
            continue
        if key == "TOTAL" and region and len(nums) >= 2:
            add(f"{region}_revenue", cur, _num(nums[0]), "DIVISIONAL REVENUES", printed)
            add(f"{region}_revenue", prior, _num(nums[1]), "DIVISIONAL REVENUES", printed)
            region = None
            continue
        ids = RELEASE_SEGMENTS.get(key)
        if ids and len(nums) >= 4:
            add(ids[0], cur, _num(nums[0]), "DIVISIONAL REVENUES", printed)
            add(ids[0], prior, _num(nums[1]), "DIVISIONAL REVENUES", printed)
            for mid, cell in ((ids[1], nums[2]), (ids[2], nums[3])):
                if mid and _num(cell) is not None:
                    add(mid, cur, _num(cell) / 100, "DIVISIONAL REVENUES", printed)
            done = key == "TOTAL NIKE, INC. REVENUES"
    for label, nums, printed in tables.get("EARNINGS BEFORE INTEREST AND TAXES", []):
        key = _release_label(label)
        if key.startswith("TOTAL NIKE, INC.") and "EARNINGS BEFORE INTEREST AND TAXES" in key and len(nums) >= 2:
            add("ebit_reported", cur, _num(nums[0]), "EARNINGS BEFORE INTEREST AND TAXES", printed)
            add("ebit_reported", prior, _num(nums[1]), "EARNINGS BEFORE INTEREST AND TAXES", printed)
    # Narrative: the fourth-quarter review, up to the review of the full year.
    start = text.find("Fourth Quarter Income Statement Review")
    end = min([i for i in (text.find(f"Fiscal {fy} Income Statement Review", start), text.find("Full Year Income Statement Review", start))
               if i > start] or [len(text)])
    quarter = text[start:end] if start >= 0 else ""
    section = "Fourth Quarter Income Statement Review"
    rest = r"(?:[^.]|\.(?=\d))*\."     # the rest of the sentence; a decimal point ($4.5 billion) does not end it
    m = re.search(r"Revenues for NIKE, Inc\. (increased|decreased|rose|fell) (\d+) percent to \$[\d.]+ (?:billion|million), (up|down) "
                  r"(\d+) percent on a currency-neutral basis" + rest, quarter)
    if m:
        add("revenues_pct_change_cn", cur, _pct(m.group(3), m.group(4)), section, m.group(0), "q4_release_text")
    m = re.search(r"Revenues for the NIKE Brand were \$[\d.]+ (?:billion|million), (up|down) (\d+) percent[^.]*?currency-neutral basis" + rest,
                  quarter)
    if m:
        add("nike_brand_pct_change_cn", cur, _pct(m.group(1), m.group(2)), section, m.group(0), "q4_release_text")
    m = re.search(r"NIKE Direct fourth quarter (?:sales|revenues) (increased|decreased|rose|fell) (\d+) percent" + rest, text)
    if m:
        add("nike_direct_pct_change_reported", cur, _pct(m.group(1), m.group(2)), "Highlights", m.group(0), "q4_release_text")
    for title in ("CONSOLIDATED STATEMENTS OF INCOME", "DIVISIONAL REVENUES", "EARNINGS BEFORE INTEREST AND TAXES"):
        if not tables.get(title):
            problems.append(f"{f['id']}: release table {title} not found")
    if not quarter:
        problems.append(f"{f['id']}: fourth-quarter review not found")
    return rows, problems


# ---------------------------------------------------------------- output
def _fmt(v):
    if isinstance(v, Decimal):
        return format(v.normalize(), "f")
    return "" if v is None else v


def write_csv(path, rows, fields):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow({k: _fmt(r.get(k)) for k in fields})


def main():
    all_rows, quotes, problems = [], [], []
    for f in all_filings():
        if f["form"] == "8-K":
            path, _ = filing_paths(f)
            rows, probs = extract_release(path, f)
            all_rows += rows
            problems += probs
            print(f"{f['id']}: {f['accession']} parsed, {len(rows)} values from the earnings release")
            continue
        doc = Document(*filing_paths(f))
        for extractor in (extract_xbrl, extract_mdna_tables, extract_narrative, extract_quarterly_table, extract_inventory_note):
            rows, probs = extractor(doc, f)
            all_rows += rows
            problems += probs
        rows, probs = extract_quotes(doc, f)
        quotes += rows
        problems += probs
        print(f"{f['id']}: {f['accession']} parsed, {len(doc.facts)} XBRL facts")
    write_csv(OUT / "facts.csv", all_rows, FIELDS)
    order = [q["id"] for q in QUOTES]
    write_csv(OUT / "quotes.csv", sorted(quotes, key=lambda r: order.index(r["id"])), ["id", "filing", "page", "section", "text"])
    (OUT / "meta.json").write_text(json.dumps({"extracted_on": datetime.date.today().isoformat(),
                                               "filings": [f["accession"] for f in all_filings()]}, indent=2) + "\n")
    print(f"{len(all_rows)} extracted values -> data/extracted/facts.csv")
    for p in problems:
        print("NOTE:", p)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
