"""Assemble observations from extracted facts and compute KPIs, checks and comparisons.

Selection rule (documented in the workbook):
  * The value for fiscal year t is the one reported in the Form 10-K for fiscal year t.
  * The "prior" value is fiscal year t-1 as presented in that same 10-K (its comparative
    column). Growth is always computed on that same-filing basis, as NIKE does.
  * Later 10-Ks that present fiscal year t again are compared with the original; any
    difference is reported as a restatement note, never silently substituted.
"""
import csv
import json
import re
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from .metrics import (CHECK_CONFIRMS, CHECKS, COMPARISONS, KPIS, REPORTED, STATUS_CONFLICT, STATUS_MATCHED,
                      STATUS_MISSING, STATUS_PENDING, USD_B)
from .sources import FIRST_FY, LAST_FY, filing

ROOT = Path(__file__).resolve().parent.parent
FACTS_CSV = ROOT / "data" / "extracted" / "facts.csv"
SEC_CHECK_CSV = ROOT / "data" / "extracted" / "sec_api_check.csv"
PAGE_CHECK_CSV = ROOT / "data" / "extracted" / "page_check.csv"
YEARS = list(range(FIRST_FY, LAST_FY + 1))
IDENT = re.compile(r"[a-z][a-z0-9_]*")


def method_of(m):
    if m.get("xbrl"):
        return "xbrl"
    return "mdna_text" if m["mdna"] in ("digital", "share") else "mdna_table"


def load_facts(path=FACTS_CSV):
    facts = {}
    with open(path, encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            r["value"] = Decimal(r["value"]) if r["value"] != "" else None
            for k in ("filing_fy", "period_fy"):
                r[k] = int(r[k])
            r["page"] = int(r["page"]) if r["page"] else None
            facts[(r["filing_fy"], r["metric_id"], r["period_fy"], r["method"])] = r
    return facts


def load_sec_check(path=SEC_CHECK_CSV):
    with open(path, encoding="utf-8") as fh:
        return {(int(r["filing_fy"]), r["metric_id"], int(r["period_fy"])): r["result"] for r in csv.DictReader(fh)}


def load_page_check(path=PAGE_CHECK_CSV):
    """Results of pipeline.recheck_pages: {(fiscal year, metric id): row}."""
    if not path.exists():
        return {}
    with open(path, encoding="utf-8") as fh:
        return {(int(r["filing_fy"]), r["metric_id"]): r for r in csv.DictReader(fh)}


MINUS = "−"


def vi_num(v, decimals=0, signed=False):
    """Vietnamese number format: '.' groups thousands, ',' marks decimals, '−' for negatives (27.758; 12,4; −0,38).

    Rounds half away from zero, as JavaScript (Intl), Excel and DAX do, so every version shows the same digits.
    """
    q = Decimal(str(v)).quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
    s = f"{abs(q):,.{decimals}f}".replace(",", "§").replace(".", ",").replace("§", ".")
    if q < 0:
        return MINUS + s
    return ("+" + s) if signed and q > 0 else s


def _fmt_num(v, unit, signed=False):
    if v is None:
        return "—"
    if unit == "%":
        return vi_num(v * 100, 0, signed) + "%"
    if unit == USD_B:
        return vi_num(v, 1, signed) + " tỷ USD"
    return vi_num(v, 0, signed) + " triệu USD"


def restatement_note(facts, m, fy, value):
    method = method_of(m)
    later = [(t, facts[(t, m["id"], fy, method)]["value"]) for t in range(fy + 1, LAST_FY + 1)
             if (t, m["id"], fy, method) in facts]
    if m["unit"] == "%" and m["id"] != "nike_direct_share_stated":
        return "Không áp dụng: tỷ lệ % chỉ lấy từ 10-K của chính năm đó."
    if not later:
        return "Chưa có 10-K năm sau để so sánh." if fy == LAST_FY else "10-K năm sau không trình bày lại số này."
    changed = [(t, v) for t, v in later if v != value]
    if not changed:
        return "Không thay đổi khi trình bày lại trong 10-K " + ", ".join(f"FY{t}" for t, _ in later) + "."
    parts = [f"10-K FY{t} trình bày lại là {_fmt_num(v, m['unit'])} (chênh {_fmt_num(v - value, m['unit'], signed=True)})"
             for t, v in changed]
    return "ĐIỀU CHỈNH HỒI TỐ: " + "; ".join(parts) + "."


def short_section(section):
    """Printed section heading, with the long Item 7 title shortened to its usual name."""
    if section.startswith("ITEM 7"):
        return "ITEM 7 – MD&A"
    return section


def build_observations(facts, sec_check):
    """One observation per (reported metric, fiscal year), before cross-check statuses."""
    obs = []
    for m in REPORTED:
        method = method_of(m)
        for fy in YEARS:
            f = filing(fy)
            cur = facts.get((fy, m["id"], fy, method))
            # Percentages are not used with a prior-year base, so no comparative is carried for them.
            prior = facts.get((fy, m["id"], fy - 1, method)) if m["unit"] != "%" else None
            o = dict(key=f"FY{fy}|{m['id']}", fy=fy, metric=m, value=cur["value"] if cur else None,
                     prior=prior["value"] if prior else None, accession=f["accession"], filed=f["filed"],
                     page=cur["page"] if cur else None, section=short_section(cur["section"]) if cur else "",
                     evidence=cur["evidence"] if cur else "", concept=cur["concept"] if cur else "",
                     dims=cur["dims"] if cur else "", method=method, cross_check="", status=None,
                     restatement="", verified_by="", verified_on="", team_note="")
            # An XBRL tag holds the digits only; the filing prints negative amounts in parentheses outside the tag.
            if method == "xbrl" and cur and cur["value"] is not None and cur["value"] < 0 and o["evidence"][:1] != "(":
                o["evidence"] = f"({o['evidence']})"
            if cur is None:
                o["status"] = STATUS_MISSING
                o["cross_check"] = f"10-K FY{fy} không công bố chỉ tiêu này."
            elif method == "xbrl":
                o["restatement"] = restatement_note(facts, m, fy, cur["value"])
                if not o["dims"]:
                    sec = sec_check.get((fy, m["id"], fy))
                    o["cross_check"] = "Khớp API XBRL của SEC (cùng số accession)." if sec == "match" else f"API SEC: {sec}"
                    o["status"] = STATUS_MATCHED if sec == "match" else STATUS_PENDING
                else:
                    md = facts.get((fy, m["id"], fy, "mdna_table"))
                    same = md is not None and md["value"] == cur["value"]
                    o["cross_check"] = (f"Khớp bảng doanh thu trong MD&A (trang {md['page']})." if same
                                        else "Không có/không khớp bảng MD&A để đối chiếu.")
                    o["status"] = STATUS_MATCHED if same else STATUS_PENDING
            else:
                o["restatement"] = restatement_note(facts, m, fy, cur["value"])
                o["status"] = STATUS_PENDING
                o["cross_check"] = "Số liệu trong phần MD&A, không gắn thẻ XBRL nên không đối chiếu tự động được; kiểm tra trực tiếp trên 10-K."
            obs.append(o)
    return obs


def env_for(obs, fy):
    env = {}
    for o in obs:
        if o["fy"] == fy:
            env[o["metric"]["id"]] = o["value"]
            env[o["metric"]["id"] + "__prior"] = o["prior"]
    return env


def evaluate(expr, env):
    names = IDENT.findall(expr)
    if any(env.get(n) is None for n in names):
        return None
    return eval(expr, {"__builtins__": {}}, {n: Decimal(env[n]) for n in names})  # catalogue expressions only


def compute(obs):
    """KPIs, checks and comparisons for every fiscal year: {fy: {id: Decimal|None}}."""
    out = {}
    for fy in YEARS:
        env = env_for(obs, fy)
        for k in KPIS + CHECKS + COMPARISONS:
            env[k["id"]] = evaluate(k["expr"], env)
        out[fy] = env
    return out


def apply_cross_check_statuses(obs, results):
    """Promote MD&A values to 'matched' when an independent calculation confirms them."""
    confirms = {c["confirms"]: c for c in COMPARISONS}
    confirms.update({metric: next(c for c in CHECKS if c["id"] == cid) for cid, metric in CHECK_CONFIRMS.items()})
    for o in obs:
        c = confirms.get(o["metric"]["id"])
        if c is None or o["value"] is None:
            continue
        diff = results[o["fy"]][c["id"]]
        ok = diff is not None and abs(diff) <= Decimal(str(c["tol"]))
        unit = "điểm %" if o["metric"]["unit"] == "%" else "triệu USD"
        shown = (vi_num(diff * 100, 2, signed=True) if unit == "điểm %" else vi_num(diff, 0, signed=True)) + " " + unit
        tol = vi_num(Decimal(str(c["tol"])) * (100 if unit == "điểm %" else 1), 0)
        if ok:
            o["status"] = STATUS_MATCHED
            o["cross_check"] = f"Khớp với giá trị tự tính từ BCTC (chênh {shown}, ngưỡng ±{tol} {unit})."
        else:
            o["status"] = STATUS_PENDING
            o["cross_check"] = f"CHÊNH LỆCH với giá trị tự tính: {shown}."
    return obs


def apply_page_checks(obs, checks):
    """Confirm MD&A-only values that a second reading of the rendered 10-K page matched.

    The check must have read the same value that is in the dataset now; if the extraction
    changed since, the value stays pending until recheck_pages is run again.
    """
    for o in obs:
        c = checks.get((o["fy"], o["metric"]["id"]))
        if c is None or o["status"] != STATUS_PENDING or o["value"] is None:
            continue
        if Decimal(c["extracted_value"]) != o["value"]:
            continue
        if c["result"] == "match":
            o["status"] = STATUS_MATCHED
            o["cross_check"] = (f"Đọc lại trực tiếp trên trang 10-K (tr. {c['page']}) bằng cách thứ hai: khớp. "
                                f"Ảnh đối chiếu: handover/bang-chung-doi-chieu/{c['screenshot']}.")
        else:
            o["status"] = STATUS_CONFLICT
            o["cross_check"] = f"CHÊNH LỆCH khi đọc lại trên trang 10-K (tr. {c['page']}): đọc được {c['read_value'] or 'không có số'}."
    return obs


def dims_label(dims):
    if not dims:
        return ""
    d = json.loads(dims)
    return "; ".join(f"{k.split(':')[1]}={v.split(':')[1]}" for k, v in sorted(d.items()))
