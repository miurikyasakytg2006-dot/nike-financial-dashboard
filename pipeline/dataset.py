"""Assemble observations from extracted facts and compute KPIs, checks and comparisons.

Selection rule (documented in the workbook):
  * A fiscal year's value is the one in that year's Form 10-K; a quarter's value (first to third quarter) is the one
    in that quarter's Form 10-Q (the three-month column).
  * The fourth quarter has no 10-Q. Its flow values are the year (10-K) minus the first nine months (third-quarter
    10-Q); its balances are the 31 May balances of the 10-K.
  * The "prior" value is the same period one year earlier as presented in the same report (its comparative
    column). Growth is always computed on that same-report basis, as NIKE does. The exception is a first- to
    third-quarter inventory balance: the 10-Q compares it with the previous 31 May, so its prior is the balance in the
    previous year's 10-Q for the same quarter.
  * Later reports that present a period again are compared with the original; any difference is reported as a
    restatement note, never silently substituted.
"""
import csv
import json
import re
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from .metrics import (A, CHECK_CONFIRMS, CHECKS, COMPARISONS, KIND_Q4, KIND_RELEASE, KPIS, Q, REPORTED, STATUS_CONFLICT,
                      STATUS_MATCHED, STATUS_MISSING, STATUS_PENDING, USD_B, applies)
from .sources import QUARTERS, SCOPE_YEARS, all_filings, annual_filing, filing, quarterly_filing, release_filing

ROOT = Path(__file__).resolve().parent.parent
FACTS_CSV = ROOT / "data" / "extracted" / "facts.csv"
PAGE_CHECK_CSV = ROOT / "data" / "extracted" / "page_check.csv"
PERIODS_A = [f"FY{y}" for y in SCOPE_YEARS]
PERIODS_Q = [f"FY{fy}Q{q}" for fy, q in QUARTERS]
PERIODS = {A: PERIODS_A, Q: PERIODS_Q}
IDENT = re.compile(r"[a-z][a-z0-9_]*")
FILED = {f["id"]: f["filed"] for f in all_filings()}
YTD = {2: "H1", 3: "M9"}
MINUS = "−"
# What a later report shows for a figure it prints again (column "Ở báo cáo năm sau" of the workbook).
REPEATED = "Báo cáo năm sau in lại đúng số này: "
REPEATED_Q4 = "Báo cáo năm sau in lại đúng số cả năm và số 9 tháng: "
NOT_REPEATED = "Báo cáo năm sau không in lại số này."
CHANGED = "CÓ THAY ĐỔI: "


def method_of(m):
    if m.get("xbrl"):
        return "xbrl"
    return "mdna_text" if m["mdna"] in ("digital", "share") else "mdna_table"


def is_instant(m):
    return m.get("xbrl", {}).get("period") == "instant"


def freq_of(period):
    return Q if "Q" in period else A


def period_label(period):
    """FY2020 -> 'FY2020'; FY2020Q3 -> 'Q3 FY2020'."""
    return f"Q{period[7]} {period[:6]}" if freq_of(period) == Q else period


def shift(period, years=-1):
    return f"FY{int(period[2:6]) + years}{period[6:]}"


def load_facts(path=FACTS_CSV):
    facts = {}
    with open(path, encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            r["value"] = Decimal(r["value"]) if r["value"] != "" else None
            r["page"] = int(r["page"]) if r["page"] else None
            facts[(r["filing"], r["metric_id"], r["period"], r["method"])] = r
    return facts


def load_page_check(path=PAGE_CHECK_CSV):
    """Results of pipeline.recheck_pages: {(filing id, metric id, period): row}."""
    if not path.exists():
        return {}
    with open(path, encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        if "period" not in (reader.fieldnames or []):   # readings made for an earlier layout of the dataset
            return {}
        return {(r["filing"], r["metric_id"], r["period"]): r for r in reader}


def vi_num(v, decimals=0, signed=False):
    """Vietnamese number format: '.' groups thousands, ',' marks decimals, '−' for negatives (27.758; 12,4; −0,38).

    Rounds half away from zero, as JavaScript (Intl), Excel and DAX do, so every version shows the same digits.
    """
    q = Decimal(str(v)).quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
    s = f"{abs(q):,.{decimals}f}".replace(",", "§").replace(".", ",").replace("§", ".")
    if q < 0:
        return MINUS + s
    return ("+" + s) if signed and q > 0 else s


def _decimals(v):
    return max(0, -Decimal(v).normalize().as_tuple().exponent)


def fmt_value(v, unit, signed=False):
    if v is None:
        return "—"
    if unit == "%":
        return vi_num(v * 100, 0, signed) + "%"
    if unit == USD_B:
        return vi_num(v, max(1, _decimals(v)), signed) + " tỷ USD"
    return vi_num(v, 0, signed) + " triệu USD"


def printed(v, unit):
    """An amount the way NIKE prints it (US style): 37,403 and (790) in millions; $5.5 billion or $2,781 million."""
    if unit == USD_B:
        return f"${v} billion" if _decimals(v) <= 1 else f"${v * 1000:,.0f} million"
    return f"({abs(v):,.0f})" if v < 0 else f"{v:,.0f}"


def same_value(a, b, unit):
    """Equal, allowing for a figure given once in millions and once rounded to 0.1 billion (Digital sales)."""
    if a == b:
        return True
    if unit != USD_B:
        return False
    coarse = min(_decimals(a), _decimals(b))
    q = Decimal(1).scaleb(-coarse)
    return a.quantize(q, rounding=ROUND_HALF_UP) == b.quantize(q, rounding=ROUND_HALF_UP)


def short_section(section):
    """Printed section heading, with the long Item 2 / Item 7 titles shortened to their usual name."""
    if re.match(r"^ITEM [27]\.", section, re.IGNORECASE):
        return section[:6].upper().rstrip(". ") + " – MD&A"
    if re.match(r"^ITEM 1\. FINANCIAL STATEMENTS", section, re.IGNORECASE):
        return "ITEM 1 – FINANCIAL STATEMENTS"
    return section


class Facts:
    """Extracted facts with lookups by source filing and across later filings."""

    def __init__(self, facts):
        self.facts = facts
        self.by_metric = {}
        for (fid, mid, period, method), r in facts.items():
            self.by_metric.setdefault((mid, period), []).append(r)

    def get(self, fid, m, period, methods=None):
        for method in methods or [method_of(m)]:
            r = self.facts.get((fid, m["id"], period, method))
            if r is not None and r["value"] is not None:
                return r
        return None

    def later(self, m, period, after_fid):
        """Other presentations of the same period in reports filed after `after_fid`, one per report."""
        out = {}
        for r in self.by_metric.get((m["id"], period), []):
            if FILED[r["filing"]] > FILED[after_fid] and r["method"] in (method_of(m), "mdna_table", "xbrl"):
                out.setdefault(r["filing"], r)
        return sorted(out.values(), key=lambda r: FILED[r["filing"]])


def _prior_methods(m):
    # Reports filed before NIKE broke revenue down by channel in its notes give the comparative only in the MD&A.
    return [method_of(m), "mdna_table"] if method_of(m) == "xbrl" else [method_of(m)]


# ---------------------------------------------------------------- observations
def _blank(key, period, m, src):
    return dict(key=key, period=period, freq=freq_of(period), fy=int(period[2:6]),
                quarter=int(period[7]) if freq_of(period) == Q else None, metric=m, value=None, prior=None, begin=None,
                source=src["id"], source2="", accession=src["accession"], filed=src["filed"], page=None, page2=None,
                section="", evidence="", concept="", dims="", method=method_of(m), approx=False, rounded="", confirms=[],
                conflicts=[], cross_check="", status=None, restatement="", verified_by="", verified_on="", team_note="")


def _fill(o, r):
    o.update(value=r["value"], page=r["page"], section=short_section(r["section"]), evidence=r["evidence"],
             concept=r["concept"], dims=r["dims"])
    # An XBRL tag holds the digits only; the filing prints negative amounts in parentheses outside the tag.
    if r["method"] == "xbrl" and r["value"] < 0 and o["evidence"][:1] != "(":
        o["evidence"] = f"({o['evidence']})"


def _restatement(fx, o, m, period, fid):
    later = fx.later(m, period, fid)
    if not later:
        return NOT_REPEATED
    changed = [r for r in later if not same_value(r["value"], o["value"], m["unit"])]
    if not changed:
        return REPEATED + ", ".join(dict.fromkeys(r["filing"] for r in later)) + "."
    return CHANGED + "; ".join(f"{r['filing']} in lại là {fmt_value(r['value'], m['unit'])}" for r in changed) + "."


def observe(fx, m, period):
    freq = freq_of(period)
    fy = int(period[2:6])
    q = int(period[7]) if freq == Q else None
    key = f"{period}|{m['id']}"
    if freq == A or q < 4 or is_instant(m):
        src = annual_filing(fy) if freq == A or q == 4 else quarterly_filing(fy, q)
        own = f"FY{fy}" if (freq == A or q == 4) else period
        o = _blank(key, period, m, src)
        r = fx.get(src["id"], m, own)
        if r is None:
            o["status"] = STATUS_MISSING
            o["cross_check"] = f"{src['id']} không công bố chỉ tiêu này."
            return o
        _fill(o, r)
        if m["unit"] != "%":
            if is_instant(m):
                if freq == A or q == 4:
                    p = fx.get(src["id"], m, f"FY{fy - 1}")
                else:
                    p = fx.get(quarterly_filing(fy - 1, q)["id"], m, shift(period)) if fy - 1 in SCOPE_YEARS else None
            else:
                p = fx.get(src["id"], m, shift(own), _prior_methods(m))
            o["prior"] = p["value"] if p else None
        o["restatement"] = _restatement(fx, o, m, own, src["id"])
        return o
    ann, q3 = annual_filing(fy), quarterly_filing(fy, 3)
    if m["unit"] == "%":
        # A fourth-quarter percentage cannot be derived; NIKE states some in its fourth-quarter earnings release.
        rel = release_filing(fy)
        r = fx.get(rel["id"], m, period, ["q4_release", "q4_release_text"]) if rel else None
        o = _blank(key, period, m, rel or ann)
        if r is None:
            o["status"] = STATUS_MISSING
            o["cross_check"] = "NIKE không nêu tỷ lệ này cho riêng quý 4: không có trong 10-K, 10-Q và thông cáo kết quả quý 4."
            return o
        _fill(o, r)
        o.update(method="q4_release", kind=KIND_RELEASE, restatement=NOT_REPEATED)
        return o
    # Fourth quarter of a flow metric: year minus nine months.
    o = _blank(key, period, m, ann)
    o.update(method="q4_derived", source2=q3["id"], kind=KIND_Q4)
    y, n = fx.get(ann["id"], m, f"FY{fy}"), fx.get(q3["id"], m, f"FY{fy}M9")
    if y is None or n is None:
        o["status"] = STATUS_MISSING
        o["cross_check"] = "Thiếu số cả năm hoặc số 9 tháng để tính quý 4."
        return o
    unit = m["unit"]
    # Digital sales are given rounded to 0.1 billion, except the fiscal 2019 quarters (in millions).
    rounded = [what for what, r in (("số cả năm", y), ("số 9 tháng", n)) if unit == USD_B and _decimals(r["value"]) <= 1]
    o.update(value=y["value"] - n["value"], page=y["page"], page2=n["page"], section=short_section(y["section"]),
             approx=bool(rounded), concept=y["concept"], dims=y["dims"],
             rounded="hai số đều đã làm tròn" if len(rounded) == 2 else f"{rounded[0]} đã làm tròn" if rounded else "")
    o["evidence"] = f"{ann['id']} (cả năm): {printed(y['value'], unit)}; {q3['id']} (9 tháng): {printed(n['value'], unit)}"
    yp, np_ = fx.get(ann["id"], m, f"FY{fy - 1}", _prior_methods(m)), fx.get(q3["id"], m, f"FY{fy - 1}M9", _prior_methods(m))
    o["prior"] = yp["value"] - np_["value"] if yp and np_ else None
    later_y, later_n = fx.later(m, f"FY{fy}", ann["id"]), fx.later(m, f"FY{fy}M9", q3["id"])
    changed = [r for r in later_y if not same_value(r["value"], y["value"], unit)] + \
              [r for r in later_n if not same_value(r["value"], n["value"], unit)]
    if changed:
        o["restatement"] = "CÓ THAY ĐỔI ở số cả năm hoặc số 9 tháng: " + "; ".join(
            f"{r['filing']} in lại là {fmt_value(r['value'], unit)}" for r in changed) + "."
    elif later_y or later_n:
        o["restatement"] = (REPEATED_Q4 + ", ".join(dict.fromkeys(r["filing"] for r in later_y + later_n)) + ".")
    else:
        o["restatement"] = NOT_REPEATED
    return o


def build_observations(facts):
    """One observation per (reported metric, period), before cross-check statuses."""
    fx = facts if isinstance(facts, Facts) else Facts(facts)
    obs = []
    for m in REPORTED:
        for freq in (A, Q):
            if applies(m, freq):
                obs += [observe(fx, m, p) for p in PERIODS[freq]]
    by_key = {o["key"]: o for o in obs}
    for o in obs:   # opening balance of inventories: previous quarter's closing balance
        if o["metric"]["id"] != "inventories" or o["value"] is None:
            continue
        if o["freq"] == A or o["quarter"] == 1:   # the 10-K, or the first-quarter 10-Q, prints the previous 31 May balance
            r = fx.get(o["source"], o["metric"], f"FY{o['fy'] - 1}")
            o["begin"] = r["value"] if r else None
        else:
            prev = by_key.get(f"FY{o['fy']}Q{o['quarter'] - 1}|inventories")
            o["begin"] = prev["value"] if prev else None
    cross_check(fx, obs)
    return obs


# ---------------------------------------------------------------- independent confirmations
def _mdna_counterpart(fx, o):
    """The MD&A table printing the same amount as an XBRL note value, in the same report."""
    m = o["metric"]
    own = o["period"] if o["freq"] == Q else f"FY{o['fy']}"
    r = fx.facts.get((o["source"], m["id"], own, "mdna_table"))
    return r


def cross_check(fx, obs):
    by_key = {o["key"]: o for o in obs}
    for o in obs:
        m, unit = o["metric"], o["metric"]["unit"]
        if o["value"] is None:
            continue
        if o["method"] == "xbrl":
            md = _mdna_counterpart(fx, o)
            if md is not None:
                if md["value"] == o["value"]:
                    what = "bảng EBIT khu vực" if m["id"].endswith("_ebit") else "bảng doanh thu"
                    o["confirms"].append(f"khớp {what} trong MD&A (tr. {md['page']})")
                else:
                    o["conflicts"].append(f"MD&A (tr. {md['page']}) ghi {fmt_value(md['value'], unit)}")
            note = fx.facts.get((o["source"], m["id"], o["period"] if o["freq"] == Q else f"FY{o['fy']}", "note_text"))
            if note is not None:
                if note["value"] == o["value"]:
                    o["confirms"].append(f"khớp câu trong thuyết minh hàng tồn kho (tr. {note['page']})")
                else:
                    o["conflicts"].append(f"thuyết minh hàng tồn kho (tr. {note['page']}) ghi {fmt_value(note['value'], unit)}")
            # Quarterly table of the 10-K (fiscal 2019 and 2020): revenues, gross profit, net income by quarter.
            for r in fx.by_metric.get((m["id"], o["period"]), []):
                if r["method"] == "quarterly_table":
                    if r["value"] == o["value"]:
                        o["confirms"].append(f"khớp bảng số liệu quý trong {r['filing']} (tr. {r['page']})")
                    else:
                        o["conflicts"].append(f"bảng số liệu quý của {r['filing']} ghi {fmt_value(r['value'], unit)}")
                    break
        if o["method"] == "q4_derived":
            for r in fx.by_metric.get((m["id"], o["period"]), []):
                if r["method"] == "quarterly_table":
                    if r["value"] == o["value"]:
                        o["confirms"].append(f"khớp bảng số liệu quý trong {r['filing']} (tr. {r['page']})")
                    else:
                        o["conflicts"].append(f"bảng số liệu quý của {r['filing']} ghi {fmt_value(r['value'], unit)}")
                    break
            # NIKE's fourth-quarter earnings release prints the quarter's three-month figure (and the next year's
            # release prints it again as the comparative).
            released = [r for r in fx.by_metric.get((m["id"], o["period"]), []) if r["method"] == "q4_release"]
            same = sorted({r["filing"] for r in released if r["value"] == o["value"]}, key=lambda fid: FILED[fid])
            if same:
                o["confirms"].append("khớp số quý 4 NIKE công bố trong thông cáo kết quả quý 4 (" + ", ".join(same) + ")")
            for r in released:
                if r["value"] != o["value"]:
                    o["conflicts"].append(f"thông cáo {r['filing']} ghi {fmt_value(r['value'], unit)}")
        if o["method"] == "q4_release":
            # The same percentage printed a second time in the release: in a sentence, or in a repeated table row.
            for r in fx.by_metric.get((m["id"], o["period"]), []):
                if r["filing"] != o["source"] or r["method"] not in ("q4_release_text", "q4_release_repeat") or \
                        r["evidence"] == o["evidence"]:
                    continue
                where = (f"câu trong phần “{r['section']}” của cùng thông cáo" if r["method"] == "q4_release_text"
                         else f"dòng in lại trong bảng “{r['section']}” của cùng thông cáo")
                if r["value"] == o["value"]:
                    o["confirms"].append(f"khớp {where}")
                else:
                    o["conflicts"].append(f"{where} ghi {fmt_value(r['value'], unit)}")
        if o["method"] == "mdna_table" and unit == "%":
            # The same % change printed a second time in the same report: the revenues-by-segment table (NIKE Brand,
            # NIKE, Inc.) or the NIKE Direct sentence of the NIKE Brand discussion.
            own = o["period"] if o["freq"] == Q else f"FY{o['fy']}"
            for method, where in (("mdna_segment_table", "bảng doanh thu theo khu vực trong MD&A"), ("mdna_text", "câu trong MD&A")):
                r = fx.facts.get((o["source"], m["id"], own, method))
                if r is None:
                    continue
                if r["value"] == o["value"]:
                    o["confirms"].append(f"khớp {where} (tr. {r['page']})")
                else:
                    o["conflicts"].append(f"{where} (tr. {r['page']}) ghi {fmt_value(r['value'], unit)}")
        # A quarter's figures add up to the year-to-date figure printed in the next 10-Q of the same year.
        if o["freq"] == Q and o["quarter"] in (2, 3) and not is_instant(m) and unit != "%":
            parts = [by_key.get(f"FY{o['fy']}Q{k}|{m['id']}") for k in range(1, o["quarter"] + 1)]
            ytd = fx.get(o["source"], m, f"FY{o['fy']}{YTD[o['quarter']]}")
            if ytd and all(p and p["value"] is not None for p in parts):
                total = sum(p["value"] for p in parts)
                months = "6" if o["quarter"] == 2 else "9"
                if abs(total - ytd["value"]) <= rounding_tolerance(unit, o["quarter"]):
                    o["confirms"].append(f"cộng các quý từ Q1 bằng số {months} tháng trong {o['source']} (tr. {ytd['page']})")
                else:
                    o["conflicts"].append(f"cộng các quý ra {fmt_value(total, unit)}, khác số {months} tháng "
                                          f"{fmt_value(ytd['value'], unit)}")
        if o["restatement"].startswith((REPEATED, REPEATED_Q4)):
            o["confirms"].append("báo cáo năm sau in lại đúng số này")


def finalize_statuses(obs):
    """Status and cross-check text from the confirmations found so far (call after the KPI comparisons)."""
    for o in obs:
        if o["value"] is None:
            continue
        if o["conflicts"]:
            o["status"] = STATUS_CONFLICT
            o["cross_check"] = "MÂU THUẪN: " + "; ".join(o["conflicts"]) + "."
            continue
        lead = ""
        if o["method"] == "q4_derived":
            lead = (f"Quý 4 ước tính (không có 10-Q cho quý 4): cả năm trừ 9 tháng đầu, {o['rounded']}. " if o["approx"]
                    else "Quý 4 tự tính (không có 10-Q cho quý 4): cả năm trừ 9 tháng đầu. ")
        if o["confirms"]:
            o["status"] = STATUS_MATCHED
            text = "; ".join(dict.fromkeys(o["confirms"]))
            o["cross_check"] = lead + text[:1].upper() + text[1:] + "."
        else:
            o["status"] = STATUS_PENDING
            o["cross_check"] = lead + "Chưa có nguồn thứ hai để đối chiếu."
    return obs


# ---------------------------------------------------------------- KPIs
def env_for(obs, period):
    env = {"days_in_period": Decimal(365) if freq_of(period) == A else Decimal("91.25")}
    for o in obs:
        if o["period"] == period:
            env[o["metric"]["id"]] = o["value"]
            env[o["metric"]["id"] + "__prior"] = o["prior"]
            if o["metric"]["id"] == "inventories":
                env["inventories__begin"] = o["begin"]
    return env


def evaluate(expr, env):
    names = IDENT.findall(expr)
    if any(env.get(n) is None for n in names):
        return None
    return eval(expr, {"__builtins__": {}}, {n: Decimal(env[n]) for n in names})  # catalogue expressions only


def compute(obs):
    """KPIs, checks and comparisons for every period: {period: {id: Decimal|None}}."""
    out = {}
    for freq in (A, Q):
        for period in PERIODS[freq]:
            env = env_for(obs, period)
            for k in KPIS + CHECKS + COMPARISONS:
                env[k["id"]] = evaluate(k["expr"], env) if applies(k, freq) else None
            out[period] = env
    return out


def apply_cross_check_statuses(obs, results):
    """Confirm MD&A values that an independent calculation from the financial statements reproduces."""
    confirms = {c["confirms"]: c for c in COMPARISONS}
    confirms.update({metric: next(c for c in CHECKS if c["id"] == cid) for cid, metric in CHECK_CONFIRMS.items()})
    for o in obs:
        c = confirms.get(o["metric"]["id"])
        if c is None or o["value"] is None:
            continue
        diff = results[o["period"]][c["id"]]
        if diff is None:
            continue
        pts = o["metric"]["unit"] == "%"
        shown = vi_num(diff * 100, 2, signed=True) + " điểm %" if pts else vi_num(diff, 0, signed=True) + " triệu USD"
        tol = vi_num(Decimal(str(c["tol"])) * (100 if pts else 1), 0) + (" điểm %" if pts else " triệu USD")
        if abs(diff) <= Decimal(str(c["tol"])):
            o["confirms"].append(f"khớp số tính lại từ báo cáo tài chính (chênh {shown}, ngưỡng ±{tol})")
        else:
            o["conflicts"].append(f"chênh {shown} so với giá trị tự tính")
    return obs


def apply_page_checks(obs, checks):
    """Confirm MD&A-only values that a second reading of the rendered report page matched.

    The check must have read the same value that is in the dataset now; if the extraction changed since, the value
    stays unconfirmed until recheck_pages is run again.
    """
    for o in obs:
        own = o["period"] if o["freq"] == Q else f"FY{o['fy']}"
        c = checks.get((o["source"], o["metric"]["id"], own))
        if c is None or o["value"] is None or Decimal(c["extracted_value"]) != o["value"]:
            continue
        if c["result"] == "match":
            o["confirms"].append(f"đã đọc lại ngay trên trang {o['source']} (tr. {c['page']}): khớp")
        else:
            o["conflicts"].append(f"đọc lại trên trang {o['source']} (tr. {c['page']}) được {c['read_value'] or 'không có số'}")
    return obs


def assemble(facts=None):
    """Observations with statuses, and the KPI / check results: the whole dataset in one call."""
    facts = facts or load_facts()
    obs = build_observations(facts)
    results = compute(obs)
    apply_cross_check_statuses(obs, results)
    apply_page_checks(obs, load_page_check())
    finalize_statuses(obs)
    return obs, results


def dims_label(dims):
    if not dims:
        return ""
    d = json.loads(dims)
    return "; ".join(f"{k.split(':')[1]}={v.split(':')[1]}" for k, v in sorted(d.items()))


def ytd_checks(facts):
    """Every quarter-sum check: (metric, fiscal year, months, sum of quarters, year-to-date value, source, page)."""
    fx = facts if isinstance(facts, Facts) else Facts(facts)
    out = []
    for m in REPORTED:
        if not applies(m, Q) or is_instant(m) or m["unit"] == "%":
            continue
        for fy in SCOPE_YEARS:
            qs = [fx.get(quarterly_filing(fy, k)["id"], m, f"FY{fy}Q{k}") for k in (1, 2, 3)]
            for k, ytd_id in ((2, "H1"), (3, "M9")):
                ytd = fx.get(quarterly_filing(fy, k)["id"], m, f"FY{fy}{ytd_id}")
                if ytd is None or any(x is None for x in qs[:k]):
                    continue
                total = sum(x["value"] for x in qs[:k])
                out.append(dict(metric=m, fy=fy, months=3 * k, total=total, ytd=ytd["value"],
                                filing=quarterly_filing(fy, k)["id"], page=ytd["page"],
                                ok=abs(total - ytd["value"]) <= rounding_tolerance(m["unit"], k)))
    return out


def rounding_tolerance(unit, quarters):
    """Largest gap between a sum of rounded quarters and the rounded year-to-date figure: 1 million for amounts in
    millions; for Digital sales (rounded to 0.1 billion from fiscal 2020) half a unit per rounded figure."""
    return Decimal("0.05") * (quarters + 1) if unit == USD_B else Decimal(1)


def source_of(o):
    return filing(o["source"])
