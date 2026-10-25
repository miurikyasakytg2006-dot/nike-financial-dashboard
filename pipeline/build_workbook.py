"""Build the Excel handover workbook from the extracted facts.

Usage:  python -m pipeline.build_workbook

Writes handover/NIKE_Financial_Dataset_FY2019-2021.xlsx. If the workbook already exists,
the human-verification columns (status "Đã xác minh (người)", verifier, date, team note)
are carried over by observation key; a verified value that no longer matches the
source extraction is flagged as conflicting evidence instead of silently kept.
Every KPI and check on sheets 05 and 06 is a live Excel formula; the Python result of the
same expression is stored as the cached value so previews show numbers before recalculation.
"""
import csv
import datetime
import json
import re
import subprocess
import sys
from decimal import Decimal
from pathlib import Path

import xlsxwriter
from openpyxl import load_workbook
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from . import dataset
from .metrics import (A, CHECKS, COMPARISONS, KIND_GAAP, KIND_GAAP_A, KIND_GAAP_Q, KIND_Q4, KIND_RELEASE, KPIS, KPIS_BY_ID, PCT, Q,
                      REGIONS, REPORTED, REPORTED_BY_ID, STATUS_CONFLICT, STATUS_HUMAN, STATUS_MATCHED, STATUS_MISSING,
                      STATUS_PENDING, STATUSES, USD_B, USD_M, applies, location)
from .quotes import QUOTES_BY_ID
from .sources import (QUARTERS, SCOPE_YEARS, all_filings, document_title, filing, ir_document_url, sec_document_url)

ROOT = Path(__file__).resolve().parent.parent
WORKBOOK = ROOT / "handover" / "NIKE_Financial_Dataset_FY2019-2021.xlsx"
META = ROOT / "data" / "extracted" / "meta.json"
QUOTES_CSV = ROOT / "data" / "extracted" / "quotes.csv"
DASHBOARD_URL = "https://miurikyasakytg2006-dot.github.io/nike-financial-dashboard/"

SHEET_OVERVIEW = "01 Tổng quan"
SHEET_A = "02 Số liệu năm"
SHEET_Q = "03 Số liệu quý"
SHEET_R = "04 Khu vực"
SHEET_KPI_A = "05 KPI năm"
SHEET_KPI_Q = "06 KPI quý"
SHEET_SRC = "07 Nguồn và định nghĩa"
OBS_SHEETS = (SHEET_A, SHEET_Q, SHEET_R)
HEADER_ROW = 3          # zero-based: tables start on Excel row 4
LAST_DATA_ROW = 700     # bounded lookup range used by the KPI formulas
NO_VALUE = "—"          # a KPI cell without its inputs, the same dash as an empty text cell on sheets 02–04

# Observation table columns: (header, field, maximum width; actual width follows content). The figure, its source and
# its page come first: they are also what a printed copy shows (columns A–H).
OBS_COLUMNS = [
    ("Kỳ", "period_label", 11),
    ("Chỉ tiêu", "vi", 60),
    ("Giá trị", "value", 12),
    ("Cùng kỳ năm trước", "prior", 14),
    ("Đơn vị", "unit", 10),
    ("Tài liệu nguồn (mở trên SEC)", "doc", 16),
    ("Trang", "page", 10),
    ("Trạng thái dữ liệu", "status", 24),
    ("Số liệu của", "scope", 14),
    ("Kênh", "channel", 22),
    ("Loại số liệu", "kind", 44),
    ("Mục trong báo cáo", "section", 40),
    ("Vị trí chi tiết", "location", 60),
    ("Số hoặc câu in trong báo cáo (số kiểu Mỹ: 39,117 = 39.117; “|” ngăn các ô của một dòng)", "evidence", 50),
    ("Đối chiếu", "cross_check", 54),
    ("Ở báo cáo năm sau", "restatement", 50),
    ("Người xác minh", "verified_by", 16),
    ("Ngày xác minh", "verified_on", 14),
    ("Ghi chú của nhóm", "team_note", 30),
    ("Dòng trong báo cáo (tiếng Anh)", "en", 60),
    ("Thẻ XBRL", "xbrl", 60),
    ("Kỳ báo cáo", "dates", 26),
    ("Ngày nộp SEC", "filed", 13),
    ("Số hồ sơ SEC", "accession", 24),
    ("Mã chỉ tiêu", "metric_id", 40),
    ("Mã quan sát", "key", 40),
]
FIELD_COL = {f: i for i, (_, f, _) in enumerate(OBS_COLUMNS)}
HEADER_OF = {f: h for h, f, _ in OBS_COLUMNS}
FILL_IN_FIELDS = {"verified_by", "verified_on", "team_note"}
PROSE_FIELDS = {"location", "evidence", "cross_check", "restatement", "xbrl"}
QUARTER_MONTHS = {1: ("06-01", "08-31"), 2: ("09-01", "11-30"), 3: ("12-01", None), 4: ("03-01", "05-31")}


class Styles:
    """Cached XlsxWriter formats; Arial throughout, restrained colours."""
    BASE = {"font_name": "Arial", "font_size": 10, "valign": "top"}
    INPUT_BLUE = "#0000FF"
    LINK_GREEN = "#008000"

    def __init__(self, wb):
        self.wb, self.cache = wb, {}
        wb.formats[0].set_font_name("Arial")
        wb.formats[0].set_font_size(10)

    def __call__(self, **props):
        key = tuple(sorted(props.items()))
        if key not in self.cache:
            self.cache[key] = self.wb.add_format({**self.BASE, **props})
        return self.cache[key]

    def num(self, unit, color=None, **extra):
        fmt = {USD_M: "#,##0;(#,##0);0", USD_B: "0.0##", PCT: "0.0%;−0.0%;0.0%", "lần": "0.00", "ngày": "0"}.get(unit, "General")
        props = {"num_format": fmt, **extra}
        if color:
            props["font_color"] = color
        return self(**props)


def _d(v):
    return float(v) if isinstance(v, Decimal) else v


def keep(text):
    """Explanatory text of sheets 01 and 07: a report name with its period, a quarter or month with its number, a figure
    with its unit stay on one line (Excel wraps after any hyphen and at any space). Figures and quotations elsewhere
    keep their plain characters."""
    nbh, nbsp = "‑", " "
    text = re.sub(r"\b(10|8)-([KQ])\b", "\\1" + nbh + "\\2", text)
    text = text.replace("non-GAAP", "non" + nbh + "GAAP").replace("COVID-19", "COVID" + nbh + "19")
    text = re.sub(r"(10" + nbh + "[KQ]|8" + nbh + "K) (?=FY\\d{4}|Q[1-4]|quý)", "\\1" + nbsp, text)
    text = re.sub(r"\b([Qq]uý|[Tt]háng|Q[1-4]) (\d|FY\d{4})", "\\1" + nbsp + "\\2", text)
    text = re.sub(r"([Tt]ỷ lệ) %", "\\1" + nbsp + "%", text)
    text = re.sub(r"(\d) (tỷ|triệu) USD", "\\1" + nbsp + "\\2" + nbsp + "USD", text)
    return re.sub(r"(\d) điểm %", "\\1" + nbsp + "điểm" + nbsp + "%", text)


def _vn_date(iso):
    y, mth, d = iso.split("-")
    return f"{d}/{mth}/{y}"


def period_dates(o):
    """Calendar dates a period covers: '01/06/2019 – 31/08/2019', or the balance date for a balance."""
    fy, q = o["fy"], o["quarter"]
    if q is None:
        start, end = f"{fy - 1}-06-01", f"{fy}-05-31"
    else:
        year = fy - 1 if q in (1, 2) else fy
        start_year = fy - 1 if q in (1, 2, 3) else fy
        start = f"{start_year}-{QUARTER_MONTHS[q][0]}"
        end = (f"{year}-{QUARTER_MONTHS[q][1]}" if q != 3 else
               f"{fy}-02-{29 if fy % 4 == 0 else 28}")
    if dataset.is_instant(o["metric"]):
        return f"Số dư tại {_vn_date(end)}"
    return f"{_vn_date(start)} – {_vn_date(end)}"


def text_width(s):
    """Display width in Excel column units, roughly: capitals and m, w, %, & are about a quarter wider than digits."""
    return sum(1.25 if ch.isupper() or ch in "mw%&" else 1 for ch in s)


def date_text(v):
    """A verification date as typed by the team; Excel may have stored it as a date."""
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime("%d/%m/%Y")
    return "" if v is None else str(v)


def sheet_of(m, freq):
    if m["sheet"] == "regions":
        return SHEET_R
    return SHEET_A if freq == A else SHEET_Q


def obs_key(period, metric_id):
    return f"{dataset.period_label(period)}|{metric_id}"


# ------------------------------------------------------------------ carry-over of human verification
def previous_verifications(path):
    if not path.exists():
        return {}
    wb = load_workbook(path, read_only=True)
    out = {}
    for name in OBS_SHEETS:
        if name not in wb.sheetnames:
            continue
        rows = list(wb[name].iter_rows(min_row=HEADER_ROW + 1, values_only=True))
        if not rows:
            continue
        idx = {h: i for i, h in enumerate(rows[0]) if h}
        for r in rows[1:]:
            key = r[idx["Mã quan sát"]] if "Mã quan sát" in idx else None
            if not key:
                continue
            out[key] = dict(status=r[idx["Trạng thái dữ liệu"]], value=r[idx["Giá trị"]],
                            verified_by=r[idx["Người xác minh"]] or "", verified_on=date_text(r[idx["Ngày xác minh"]]),
                            team_note=r[idx["Ghi chú của nhóm"]] or "")
    return out


def merge_verifications(obs, previous):
    for o in obs:
        p = previous.get(obs_key(o["period"], o["metric"]["id"]))
        if not p:
            continue
        o["verified_by"], o["verified_on"], o["team_note"] = p["verified_by"], p["verified_on"], p["team_note"]
        if p["status"] == STATUS_HUMAN:
            same = (p["value"] is None and o["value"] is None) or (
                p["value"] is not None and o["value"] is not None and abs(Decimal(str(p["value"])) - o["value"]) < Decimal("0.0001"))
            if same:
                o["status"] = STATUS_HUMAN
            else:
                o["status"] = STATUS_CONFLICT
                o["cross_check"] = (f"MÂU THUẪN: giá trị đã được xác minh là {p['value']} nhưng lần trích xuất mới cho "
                                    f"{_d(o['value'])}. Cần kiểm tra lại. " + o["cross_check"])


# ------------------------------------------------------------------ sheets 02 / 03 / 04
def row_values(o):
    m = o["metric"]
    q4 = o["method"] == "q4_derived"
    release = o["method"] == "q4_release"
    form = "10-K" if o["source"].startswith("10-K") else "10-Q"
    if q4:
        where = (f"Tự tính: cả năm ({location(m, '10-K')}) trừ 9 tháng ({o['source2']}, cùng dòng"
                 + (f", trang {o['page2']})" if o["page2"] else ")"))
    elif release:
        where = f"Thông cáo kết quả quý 4 (Form 8-K, Exhibit 99.1), {o['section']}"
    else:
        where = location(m, form)
    region = {r[0]: r[2] for r in REGIONS}.get(m["region"], "")
    untagged = "Không gắn thẻ XBRL (thông cáo kết quả quý 4)" if release else "Không gắn thẻ XBRL (MD&A)"
    # Financial statements: audited in the 10-K, unaudited in a 10-Q ("Unaudited Condensed Consolidated ...").
    kind = m["kind"] if m["kind"] != KIND_GAAP else KIND_GAAP_A if form == "10-K" else KIND_GAAP_Q
    return dict(period_label=dataset.period_label(o["period"]), vi=m["vi"], unit=m["unit"], scope=m["scope"],
                channel=(f"{region} – {m['channel']}" if region else m["channel"]) or "—",
                kind=KIND_Q4 if q4 else KIND_RELEASE if release else kind, status=o["status"], page=o["page"],
                section=o["section"] or "—", location=where, evidence=o["evidence"] or "—", cross_check=o["cross_check"],
                restatement=o["restatement"] or "—", verified_by=o["verified_by"], verified_on=o["verified_on"],
                team_note=o["team_note"], en=m["en"],
                xbrl=(o["concept"] + (f" [{dataset.dims_label(o['dims'])}]" if o["dims"] else "")) or untagged,
                dates=period_dates(o), filed=_vn_date(o["filed"]), accession=o["accession"],
                metric_id=m["id"], key=obs_key(o["period"], m["id"]))


def write_observation_sheet(wb, st, name, title, subtitle, obs):
    ws = wb.add_worksheet(name)
    ws.hide_gridlines(2)
    printed = FIELD_COL["status"]       # last column of the printed copy
    ws.merge_range(0, 0, 0, printed, title, st(bold=True, font_size=14))
    ws.merge_range(1, 0, 1, printed, subtitle, st(font_color="#444444", text_wrap=True))
    ws.set_row(1, 30)
    head = st(bold=True, bg_color="#E7EAEE", text_wrap=True, valign="vcenter", bottom=1)
    head_fill = st(bold=True, bg_color="#FFF2CC", text_wrap=True, valign="vcenter", bottom=1)
    rows = [row_values(o) for o in obs]
    for c, (h, field, cap) in enumerate(OBS_COLUMNS):
        # Width follows the longest entry (capped for long prose, which wraps); headers wrap to at most three lines.
        if field in ("value", "prior", "page", "doc"):
            width = cap
        else:
            longest = max((text_width(str(rv[field] or "")) for rv in rows), default=0)
            width = max(min(longest + 2, cap), max(len(w) for w in h.split()) + 2, len(h) / 3 + 2, 8)
        # Every figure is already checked, so the columns for a member's own check stay hidden until someone needs them.
        ws.set_column(c, c, width, None, {"hidden": field in FILL_IN_FIELDS})
        ws.write(HEADER_ROW, c, h, head_fill if field in FILL_IN_FIELDS else head)
    ws.set_row(HEADER_ROW, 44)
    text = st()
    # Long prose wraps so the end of it stays visible (the figure in a quoted line, "khớp" in the cross-check);
    # rows keep no fixed height, so Excel sizes them to the text.
    prose = st(text_wrap=True)
    link = st(font_color="#1F4E9A", underline=1, text_wrap=True)
    for i, (o, vals) in enumerate(zip(obs, rows)):
        r = HEADER_ROW + 1 + i
        m = o["metric"]
        f = filing(o["source"])
        italic = o["method"] == "q4_derived"   # values the dataset calculates, not figures NIKE printed
        for field, c in FIELD_COL.items():
            if field in ("value", "prior"):
                v = o[field]
                # NIKE reports these percentages rounded to whole numbers, so show them that way.
                fmt = (st(num_format="0%;−0%;0%", font_color=Styles.INPUT_BLUE) if m["unit"] == PCT
                       else st.num(m["unit"], Styles.INPUT_BLUE, italic=italic) if italic
                       else st.num(m["unit"], Styles.INPUT_BLUE))
                if v is None:
                    ws.write_blank(r, c, None, fmt)
                else:
                    ws.write_number(r, c, _d(v), fmt)
            elif field == "doc":
                label = o["source"] + (f" và {o['source2']}" if o["source2"] else "")
                # A cell holds one link: a fourth-quarter row opens the 10-K; the 10-Q is linked on sheet 07.
                tip = f"{document_title(f)} – SEC EDGAR" + (f". {o['source2']}: link ở sheet 07." if o["source2"] else "")
                ws.write_url(r, c, sec_document_url(f), link, label, tip=tip)
            elif field == "page":
                if vals["page"] is None:
                    ws.write_blank(r, c, None, text)
                elif o["page2"]:
                    ws.write_string(r, c, f"{vals['page']} (10-K)\n{o['page2']} (10-Q)", st(align="center", text_wrap=True))
                else:
                    ws.write_number(r, c, vals["page"], st(align="center"))
            else:
                ws.write_string(r, c, str(vals[field] if vals[field] is not None else ""),
                                prose if field in PROSE_FIELDS else text)
    last = HEADER_ROW + len(obs)
    ws.autofilter(HEADER_ROW, 0, last, len(OBS_COLUMNS) - 1)
    ws.freeze_panes(HEADER_ROW + 1, 2)
    sc = FIELD_COL["status"]
    ws.data_validation(HEADER_ROW + 1, sc, LAST_DATA_ROW, sc, {
        "validate": "list", "source": list(STATUSES),
        "error_title": "Trạng thái không hợp lệ", "error_message": "Chọn một trạng thái trong danh sách."})
    rng = f"{xl_col_to_name(sc)}{HEADER_ROW + 2}:{xl_col_to_name(sc)}{LAST_DATA_ROW + 1}"
    for value, fmt in ((STATUS_PENDING, st(bg_color="#FFF4CE")), (STATUS_MISSING, st(font_color="#6B6B6B", italic=True)),
                       (STATUS_HUMAN, st(bg_color="#E2F0D9")), (STATUS_CONFLICT, st(bg_color="#F8D7DA", bold=True))):
        ws.conditional_format(rng, {"type": "cell", "criteria": "==", "value": f'"{value}"', "format": fmt})
    # On paper: the figure, its source, page and status (the explanatory columns would shrink the type to nothing).
    ws.set_landscape()
    ws.set_paper(9)
    ws.print_area(0, 0, last, printed)
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(HEADER_ROW)
    ws.set_footer("&C&8Trang &P/&N")
    return ws


# ------------------------------------------------------------------ sheets 05 / 06
def input_label(tok):
    """Name of a formula input: 'Doanh thu NIKE Direct – cùng kỳ năm trước'; the opening balance of 'Hàng tồn kho cuối
    kỳ' is 'Hàng tồn kho đầu kỳ'. None for a token that is not a metric or KPI."""
    if tok == "days_in_period":
        return "Số ngày của kỳ"
    base, suffix = tok, ""
    for s, label in (("__prior", " – cùng kỳ năm trước"), ("__begin", " – đầu kỳ")):
        if tok.endswith(s):
            base, suffix = tok[:-len(s)], label
    name = REPORTED_BY_ID[base]["vi"] if base in REPORTED_BY_ID else KPIS_BY_ID[base]["vi"] if base in KPIS_BY_ID else None
    if name is None:
        return None
    if suffix == " – đầu kỳ" and "cuối kỳ" in name:
        return name.replace("cuối kỳ", "đầu kỳ")
    return name + suffix


def readable(expr):
    """Turn a catalogue expression into a Vietnamese description of the formula, with the signs used in the notes."""
    text = dataset.IDENT.sub(lambda mt: f"[{input_label(mt.group(0)) or mt.group(0)}]", expr)
    text = text.replace(" - ", " − ").replace(" * ", " × ").replace("1000", "1.000")
    # A trailing "− 1" or "/ 2" stays on the line of what it applies to.
    return re.sub(r" ([−/]) (\d+)$", " \\1 \\2", text)


def write_kpi_sheet(wb, st, name, freq, obs, results):
    ws = wb.add_worksheet(name)
    ws.hide_gridlines(2)
    periods = dataset.PERIODS[freq]
    labels = [dataset.period_label(p) for p in periods]
    first_col = 3
    note_col = first_col + len(periods)
    ws.set_column(2, 2, 10)
    ws.set_column(first_col, note_col - 1, 11 if freq == A else 10.5)
    ws.set_column(note_col, note_col, 58)
    ws.set_column(note_col + 1, note_col + 1, 62)
    what = "năm FY2019–FY2021" if freq == A else "12 quý, Q1 FY2019 – Q4 FY2021"
    # Titles start in column B: a printed copy runs from the labels to the last period (the codes, formulas and notes
    # are for the screen).
    ws.write(0, 1, f"Chỉ số tính toán (KPI) và kiểm tra đối chiếu – {what}", st(bold=True, font_size=14))
    ws.write(1, 1, "Mọi ô số trong sheet này là công thức Excel. Số màu xanh lá lấy từ sheet 02–04 theo kỳ và mã chỉ tiêu, "
                   f"số màu đen được tính ngay trong sheet này. “{NO_VALUE}” là thiếu số đầu vào, không phải bằng 0.",
             st(font_color="#444444"))
    head = st(bold=True, bg_color="#E7EAEE", text_wrap=True, valign="vcenter", bottom=1)
    head_r = st(bold=True, bg_color="#E7EAEE", align="right", valign="vcenter", bottom=1)
    head_gap = st(bold=True, bg_color="#E7EAEE", text_wrap=True, valign="vcenter", bottom=1, indent=1, align="left")
    for c, h in enumerate(["Mã", "Chỉ tiêu", "Đơn vị"] + labels + ["Công thức", "Ghi chú và cách đọc"]):
        ws.write(HEADER_ROW, c, h, head if c < first_col else head_gap if c >= note_col else head_r)
    ws.set_row(HEADER_ROW, 22)
    ws.freeze_panes(HEADER_ROW + 1, 3)
    section_fmt = st(bold=True, bg_color="#F3F4F6", top=1)
    prose = st(text_wrap=True, indent=1, align="left")   # set in from the last period's figures
    row = HEADER_ROW + 1
    refs = {p: {} for p in periods}
    written_labels, ids = [], []
    by_key = {(o["period"], o["metric"]["id"]): o for o in obs}

    def section(title):
        nonlocal row
        row += 1
        for c in range(note_col + 2):
            ws.write_blank(row, c, None, section_fmt)
        ws.write(row, 1, title, section_fmt)
        row += 1

    def lookup(sheet, col, key_text):
        val_col = xl_col_to_name(FIELD_COL[col])
        key_col = xl_col_to_name(FIELD_COL["key"])
        rng_v = f"'{sheet}'!${val_col}${HEADER_ROW + 2}:${val_col}${LAST_DATA_ROW + 1}"
        rng_k = f"'{sheet}'!${key_col}${HEADER_ROW + 2}:${key_col}${LAST_DATA_ROW + 1}"
        look = f"INDEX({rng_v},MATCH({key_text},{rng_k},0))"
        return f'=IF({look}="","{NO_VALUE}",{look})'

    items = [k for k in KPIS + CHECKS + COMPARISONS if applies(k, freq)]
    used = []
    for k in items:
        for tok in dataset.IDENT.findall(k["expr"]):
            base = tok.replace("__prior", "").replace("__begin", "")
            if (base in REPORTED_BY_ID and applies(REPORTED_BY_ID[base], freq) or tok == "days_in_period") and tok not in used:
                used.append(tok)
    order = [m["id"] for m in REPORTED]
    used.sort(key=lambda t: (t == "days_in_period", order.index(t.replace("__prior", "").replace("__begin", ""))
                             if t != "days_in_period" else 0, t.endswith("__begin"), t.endswith("__prior")))
    section("A. Số liệu đầu vào (lấy từ sheet 02–04)")
    for tok in used:
        ws.write(row, 0, tok, st(font_color="#6B6B6B"))
        ids.append(tok)
        if tok == "days_in_period":
            label, unit = "Số ngày của kỳ", "ngày"
            ws.write(row, 1, label, st())
            ws.write(row, 2, unit, st())
            for j, p in enumerate(periods):
                v = 365 if freq == A else 91.25
                ws.write_number(row, first_col + j, v, st(num_format="0.00" if freq == Q else "0", font_color=Styles.INPUT_BLUE))
                refs[p][tok] = xl_rowcol_to_cell(row, first_col + j)
            ws.write(row, note_col, "Năm: 365 ngày. Quý: 365/4 = 91,25 ngày, để các quý so sánh được với nhau.",
                     st(font_color="#6B6B6B", indent=1, align="left"))
            written_labels.append(label)
            row += 1
            continue
        base = tok.replace("__prior", "").replace("__begin", "")
        m = REPORTED_BY_ID[base]
        sheet = sheet_of(m, freq)
        label = input_label(tok)
        written_labels.append(label)
        ws.write(row, 1, label, st())
        ws.write(row, 2, m["unit"], st())
        fmt = (st(num_format="0%;−0%;0%", font_color=Styles.LINK_GREEN) if m["unit"] == PCT
               else st.num(m["unit"], Styles.LINK_GREEN))
        for j, p in enumerate(periods):
            col = first_col + j
            hdr = xl_rowcol_to_cell(HEADER_ROW, col, row_abs=True)
            o = by_key[(p, base)]
            if tok.endswith("__begin"):
                fy, q = int(p[2:6]), int(p[7]) if freq == Q else None
                if freq == A:
                    formula, cached = lookup(sheet, "prior", f'{hdr}&"|{base}"'), o["begin"]
                elif q == 1 and fy - 1 in SCOPE_YEARS:
                    formula, cached = lookup(SHEET_A, "value", f'"FY{fy - 1}|{base}"'), o["begin"]
                elif q == 1:
                    formula, cached = lookup(SHEET_A, "prior", f'"FY{fy}|{base}"'), o["begin"]
                else:
                    formula, cached = "=" + refs[periods[j - 1]][base], o["begin"]
                note = ("Năm: tồn kho ngày 31/5 năm trước (in trong cùng 10-K). Quý: tồn kho cuối quý liền trước; quý 1 lấy tồn kho "
                        "ngày 31/5 ở sheet 02.")
            else:
                prior = tok.endswith("__prior")
                formula, cached = lookup(sheet, "prior" if prior else "value", f'{hdr}&"|{base}"'), o["prior" if prior else "value"]
                note = f"Tra cứu sheet “{sheet}” theo khóa Kỳ|{base}, cột “{HEADER_OF['prior' if prior else 'value']}”."
            # A missing source value shows a dash (text), never a blank a previewer could render as 0.
            ws.write_formula(row, col, formula, fmt, NO_VALUE if cached is None else _d(cached))
            refs[p][tok] = xl_rowcol_to_cell(row, col)
        ws.write(row, note_col, note, st(font_color="#6B6B6B", indent=1, align="left"))
        row += 1

    def write_expr_rows(group, title, kind):
        nonlocal row
        section(title)
        for k in group:
            unit = k.get("unit", USD_M if kind == "check" else PCT)
            ws.write(row, 0, k["id"], st(font_color="#6B6B6B"))
            ids.append(k["id"])
            written_labels.append(k["vi"])
            ws.write(row, 1, k["vi"], st())
            ws.write(row, 2, {"check": USD_M, "cmp": "điểm %"}.get(kind, k.get("unit_label", unit)), st())
            toks = list(dict.fromkeys(dataset.IDENT.findall(k["expr"])))
            # Comparisons (differences between two percentages) and contributions are shown in percentage points (×100).
            scale = 100 if kind == "cmp" else k.get("scale", 1)
            for j, p in enumerate(periods):
                col = first_col + j
                cells = [refs[p][t] for t in toks]
                body = dataset.IDENT.sub(lambda mt: refs[p][mt.group(0)], k["expr"]).replace(" ", "")
                if scale != 1:
                    body = f"({body})*{scale}"
                formula = f'=IF(COUNT({",".join(cells)})<{len(cells)},"{NO_VALUE}",{body})'
                v = results[p][k["id"]]
                if kind == "cmp":
                    fmt = st(num_format="+0.00;−0.00;0.00", align="right")
                elif k.get("num_format"):
                    fmt = st(num_format=k["num_format"], align="right")
                else:
                    fmt = st.num(unit, None, align="right")
                ws.write_formula(row, col, formula, fmt, NO_VALUE if v is None else _d(v) * scale)
                refs[p][k["id"]] = xl_rowcol_to_cell(row, col)
            first = xl_rowcol_to_cell(row, first_col)
            last = xl_rowcol_to_cell(row, first_col + len(periods) - 1)
            if kind in ("check", "cmp"):
                tol = k["tol"]
                lim = f"{tol * scale:g}"
                tol_txt = f"±{tol} triệu USD" if kind == "check" else f"±{lim} điểm %"
                ws.write(row, note_col, f"{readable(k['expr'])}   (ngưỡng {tol_txt})", prose)
                ok = all(results[p][k["id"]] is None or abs(results[p][k["id"]]) <= Decimal(str(tol)) for p in periods)
                ws.write_formula(row, note_col + 1,
                                 f'=IF(AND(MAX({first}:{last})<={lim},MIN({first}:{last})>=-{lim}),"Đạt","Kiểm tra lại")',
                                 st(bold=True, indent=1, align="left"), "Đạt" if ok else "Kiểm tra lại")
            else:
                ws.write(row, note_col, readable(k["expr"]), prose)
                ws.write(row, note_col + 1, k["note"], prose)
            row += 1

    write_expr_rows([k for k in items if k in KPIS and k["group"] == "channels"], "B. Kênh phân phối (số của NIKE Brand)", "kpi")
    write_expr_rows([k for k in items if k in KPIS and k["group"] == "financial"], "C. Kết quả tài chính (số của NIKE, Inc.)", "kpi")
    write_expr_rows([k for k in items if k in KPIS and k["group"] == "regions"], "D. Bốn khu vực của NIKE Brand", "kpi")
    write_expr_rows([k for k in items if k in CHECKS], "E. Kiểm tra các phép cộng trừ (đúng thì chênh lệch bằng 0, cột cuối là kết luận)", "check")
    write_expr_rows([k for k in items if k in COMPARISONS],
                    "F. So với tỷ lệ NIKE tự công bố (các ô là phần chênh lệch"
                    + (f"; “{NO_VALUE}” là kỳ NIKE không nêu tỷ lệ)" if freq == Q else ")"), "cmp")
    ws.set_column(0, 0, min(max(len(i) for i in ids) + 2, 40))
    ws.set_column(1, 1, min(max(len(t) for t in written_labels) + 2, 66))
    ws.conditional_format(HEADER_ROW + 1, note_col + 1, row, note_col + 1,
                          {"type": "cell", "criteria": "==", "value": '"Kiểm tra lại"', "format": st(font_color="#C00000", bold=True)})
    ws.set_landscape()
    ws.set_paper(9)
    ws.print_area(0, 1, row, note_col - 1)
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(HEADER_ROW)
    ws.set_footer("&C&8Trang &P/&N")
    return ws


# ------------------------------------------------------------------ sheet 07
GLOSSARY = [
    ("Năm tài chính (FY)", "Fiscal year", "Năm tài chính của NIKE kết thúc ngày 31/5, ví dụ FY2020 là từ 01/06/2019 đến 31/05/2020.", ""),
    ("Quý (Q1–Q4)", "Fiscal quarter", "Q1: tháng 6–8, Q2: tháng 9–11, Q3: tháng 12–2, Q4: tháng 3–5. Ví dụ Q4 FY2020 là 01/03/2020 – 31/05/2020.",
     "Quý 1–3 có báo cáo 10-Q riêng; quý 4 không có."),
    ("Quý 4 (tự tính)", "Fourth quarter",
     "Không có 10-Q cho quý 4, nên số quý 4 = số cả năm trong 10-K trừ số 9 tháng đầu trong 10-Q quý 3.",
     "Ghi “Nhóm tự tính” ở cột “Loại số liệu”, chữ nghiêng. Số nào có trong thông cáo kết quả quý 4 đều khớp; thông cáo "
     "không có NIKE Direct, bán buôn và Digital."),
    ("Thông cáo kết quả quý 4", "Earnings release (Form 8-K, Exhibit 99.1)",
     "Thông cáo NIKE công bố khoảng một tháng trước 10-K, có số riêng của quý 4: doanh thu theo khu vực, báo cáo kết quả "
     "kinh doanh và % thay đổi của quý.", "Không kiểm toán. Dùng cho tỷ lệ % của quý 4 và để đối chiếu số quý 4 tự tính."),
    ("Lũy kế 6 tháng, 9 tháng", "Six / nine months ended", "Số cộng dồn từ đầu năm tài chính, in trong 10-Q quý 2 và quý 3.",
     "Dùng để kiểm tra: cộng các quý phải bằng số lũy kế."),
    ("Cùng kỳ năm trước", "Prior-year period",
     "Mỗi báo cáo in kèm số của cùng kỳ năm trước để so sánh, ví dụ 10-K FY2021 in cả số FY2020. Riêng tồn kho cuối quý "
     "1–3: 10-Q chỉ in kèm số dư ngày 31/5, nên số cùng kỳ lấy từ 10-Q cùng quý năm trước.",
     "Tăng trưởng được tính với số này, giống cách NIKE tự tính."),
    ("In lại ở báo cáo năm sau", "Prior periods presented again",
     "Báo cáo năm sau in lại số cũ để so sánh. Nếu NIKE đổi cách phân loại, số in lại có thể khác số gốc.",
     "Cột “Ở báo cáo năm sau” ghi kết quả so sánh. FY2019–FY2021: không số nào khác."),
    ("MD&A", "Management's Discussion and Analysis",
     "Phần thảo luận của ban điều hành trong báo cáo: giải thích kết quả, có các bảng doanh thu và tỷ lệ tăng trưởng.",
     "Không được kiểm toán."),
    ("Thuyết minh", "Notes to the financial statements",
     "Phần ghi chú chi tiết đi kèm báo cáo tài chính, ví dụ doanh thu theo kênh và theo khu vực (“Revenues”, “Operating Segments”).",
     "Năm: đã kiểm toán. Quý: không kiểm toán (10-Q ghi “Unaudited”)."),
    ("NIKE, Inc.", "NIKE, Inc. (consolidated)", "Toàn tập đoàn: NIKE Brand + Converse + Corporate.", "Phạm vi các chỉ tiêu tài chính."),
    ("NIKE Brand", "NIKE Brand", "Thương hiệu NIKE và Jordan; đến hết quý 2 FY2020 gồm cả Hurley (tính vào Bắc Mỹ). Không gồm Converse.",
     "Phạm vi các chỉ tiêu kênh và khu vực."),
    ("NIKE Direct", "NIKE Direct / Sales through Direct to Consumer",
     "Doanh thu NIKE Brand bán thẳng cho người tiêu dùng, qua cửa hàng của NIKE và Digital (ứng dụng, website của NIKE).",
     "NIKE không công bố lợi nhuận riêng của kênh này."),
    ("NIKE Brand Digital", "Digital commerce sales (Q2–Q3 FY2021: Digital sales)",
     "Phần doanh số NIKE Direct bán qua ứng dụng và website của NIKE.",
     "Chỉ có trong MD&A. Số cả năm và các quý từ FY2020 đã làm tròn đến 0,1 tỷ USD."),
    ("NIKE Direct ngoài Digital", "—", "NIKE Direct trừ doanh số Digital, gần với doanh thu cửa hàng do NIKE sở hữu.",
     "Ước tính, sai số theo độ làm tròn của số Digital."),
    ("Bán buôn (Wholesale)", "Sales to Wholesale Customers",
     "Doanh thu NIKE Brand bán sản phẩm cho nhà bán lẻ và nhà phân phối.", "Không gồm Converse."),
    ("Global Brand Divisions", "Global Brand Divisions",
     "Doanh thu: cấp phép thương hiệu NIKE và khoản nhỏ khác, không thuộc khu vực nào. Chi phí: quảng cáo và vận hành NIKE "
     "Brand quản lý tập trung, gồm cả vận hành Digital toàn cầu và công nghệ.",
     "Doanh thu dưới 0,2% doanh thu NIKE Brand. Chi phí không chia về khu vực."),
    ("Bốn khu vực", "North America, EMEA, Greater China, APLA",
     "Bắc Mỹ; châu Âu, Trung Đông và châu Phi (EMEA); Đại Trung Hoa; châu Á – Thái Bình Dương và Mỹ Latinh (APLA).",
     "Bốn khu vực cộng Global Brand Divisions bằng doanh thu NIKE Brand."),
    ("Converse", "Converse", "Mảng kinh doanh riêng của NIKE, Inc.: bán trực tiếp, bán buôn và cấp phép thương hiệu Converse.",
     "Không tính vào NIKE Direct và bán buôn của NIKE Brand."),
    ("Corporate", "Corporate",
     "Doanh thu: chủ yếu là lãi/lỗ phòng ngừa rủi ro tỷ giá liên quan đến doanh thu. Chi phí: chủ yếu là chi phí hành chính "
     "chung không chia cho bộ phận nào (phòng ban quản lý tập trung, trụ sở, bảo hiểm, phúc lợi).",
     "Doanh thu có thể âm. Chỉ dùng khi cộng đối chiếu ra doanh thu NIKE, Inc."),
    ("Biên lợi nhuận gộp", "Gross margin", "Lợi nhuận gộp / doanh thu.", "Chỉ có ở phạm vi NIKE, Inc.; NIKE không công bố biên gộp theo kênh."),
    ("Chi phí bán hàng và quản lý (SG&A)", "Total selling and administrative expense",
     "Chi phí tạo nhu cầu (Demand creation: quảng cáo, tài trợ, sự kiện thương hiệu) cộng chi phí vận hành chung (Operating "
     "overhead: lương, thuê mặt bằng, chi phí hành chính khác).",
     "NIKE không công bố khoản này theo kênh."),
    ("EBIT", "Earnings Before Interest and Taxes",
     "Lợi nhuận sau thuế + chi phí (thu nhập) lãi thuần + chi phí thuế thu nhập, theo định nghĩa của NIKE.",
     "Chỉ tiêu non-GAAP. EBIT khu vực chưa trừ chi phí Global Brand Divisions và Corporate (NIKE không chia về khu vực)."),
    ("Tăng trưởng theo USD", "Reported growth (% change)",
     "Thay đổi so với cùng kỳ năm trước, tính bằng USD theo số báo cáo, chưa loại trừ ảnh hưởng tỷ giá.",
     "Tính ở sheet 05–06 từ số kỳ này và số cùng kỳ năm trước (thường in ngay trong báo cáo đó)."),
    ("Tăng trưởng loại trừ tỷ giá", "Currency-neutral growth (% change excluding currency changes)",
     "Tăng trưởng NIKE tính lại theo tỷ giá của cùng kỳ năm trước, để bỏ phần chênh lệch do quy đổi ngoại tệ.",
     "Chỉ tiêu non-GAAP do NIKE tính, lấy nguyên văn từ MD&A."),
    ("Điểm %", "Percentage points", "Hiệu của hai tỷ lệ %, ví dụ từ 30% lên 35% là +5 điểm %.",
     "Dùng cho thay đổi của tỷ trọng, của biên lợi nhuận và ở mục F sheet 05–06."),
    ("Số hồ sơ SEC", "Accession number", "Mã SEC cấp cho mỗi lần nộp báo cáo, dùng để tìm đúng tài liệu trên EDGAR.",
     "Cột “Số hồ sơ SEC” ở sheet 02–04 và mục A dưới đây."),
    ("GAAP / Non-GAAP", "U.S. GAAP / non-GAAP financial measure",
     "GAAP: số theo chuẩn mực kế toán Mỹ, nằm trong báo cáo tài chính. Non-GAAP: chỉ tiêu công ty tự định nghĩa, nằm ngoài "
     "báo cáo tài chính.", "Xem cột “Loại số liệu” ở sheet 02–04."),
    ("XBRL", "eXtensible Business Reporting Language",
     "Nhãn dữ liệu gắn vào các con số trong báo cáo nộp SEC để máy đọc được.",
     "Báo cáo FY2019 dùng tệp XBRL riêng; từ quý 1 FY2020, nhãn nằm ngay trong báo cáo."),
]

GAPS = [
    ("Lợi nhuận và chi phí theo kênh", "Không công bố",
     "NIKE không công bố lợi nhuận, biên gộp hay chi phí riêng của NIKE Direct và bán buôn.",
     "Không ước tính. Các chỉ số lợi nhuận chỉ tính cho NIKE, Inc. và cho từng khu vực."),
    ("Quý 4", "Tự tính",
     "Không có báo cáo 10-Q cho quý 4.",
     "Số quý 4 = cả năm trừ 9 tháng đầu. Số nào có trong thông cáo kết quả quý 4 đều khớp; thông cáo không có NIKE Direct, "
     "bán buôn và Digital. Tỷ lệ % quý 4 lấy từ thông cáo; tỷ lệ nào NIKE không nêu (phần lớn tỷ lệ của NIKE Direct, bán "
     "buôn và tỷ trọng) thì để trống."),
    ("Doanh số Digital", "Độ chính xác thấp",
     "Chỉ có trong MD&A. Quý 1–3 FY2019 tính đến triệu USD; số cả năm và các quý từ FY2020 làm tròn đến 0,1 tỷ USD.",
     "Số quý 4 và tỷ trọng Digital trong NIKE Direct chỉ là xấp xỉ. Không tính tăng trưởng Digital: tính từ số đã làm tròn "
     "sẽ lệch với tỷ lệ NIKE nêu (ví dụ FY2020: 44,7% so với 47%)."),
    ("Tăng trưởng loại trừ tỷ giá", "Không tự tính lại được",
     "Muốn tính lại cần doanh thu theo từng loại tiền tệ, mà NIKE không công bố.",
     "Lấy từ bảng doanh thu trong MD&A; tỷ lệ của NIKE, Inc., NIKE Brand và NIKE Direct còn khớp với chỗ NIKE in lại trong "
     "cùng báo cáo. Cả 36 tỷ lệ quý đã đọc lại trên trang 10-Q."),
    ("Số năm trước của FY2019 theo kênh và khu vực", "Lấy từ MD&A",
     "Thuyết minh doanh thu theo kênh bắt đầu từ FY2019, nên phần XBRL không có số FY2018.",
     "Lấy số năm trước từ bảng doanh thu trong MD&A của chính báo cáo đó."),
    ("Tồn kho cuối quý của FY2018", "Không có trong bộ báo cáo",
     "Cần báo cáo 10-Q của FY2018, ngoài phạm vi đề tài.",
     "Thay đổi tồn kho so với cùng kỳ của Q1–Q3 FY2019 để trống."),
]


def load_quotes(path=QUOTES_CSV):
    with open(path, encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def write_sources_sheet(wb, st, obs, results, extracted_on):
    ws = wb.add_worksheet(SHEET_SRC)
    ws.hide_gridlines(2)
    for c, w in enumerate([24, 34, 60, 42, 14, 24, 22, 28]):
        ws.set_column(c, c, w)
    ws.write(0, 0, "Nguồn và định nghĩa", st(bold=True, font_size=14))
    head = st(bold=True, bg_color="#E7EAEE", text_wrap=True, valign="vcenter", bottom=1)
    head_gap = st(bold=True, bg_color="#E7EAEE", text_wrap=True, valign="vcenter", bottom=1, indent=1, align="left")
    title = st(bold=True, font_size=11)
    wrap = st(text_wrap=True)
    # Every column after the first is set in a little, so neighbouring texts never run together.
    wrap_gap = st(text_wrap=True, indent=1, align="left")
    link = st(font_color="#1F4E9A", underline=1, text_wrap=True, indent=1, align="left")
    row = 2

    def table(name, headers, rows, link_cols=(), verbatim=()):
        nonlocal row
        ws.write(row, 0, name, title)
        row += 1
        for c, h in enumerate(headers):
            ws.write(row, c, h, head_gap if c else head)
        row += 1
        for r in rows:
            for c, v in enumerate(r):
                if c in link_cols and isinstance(v, tuple):
                    ws.write_url(row, c, v[0], link, v[1], tip=v[0])
                elif isinstance(v, str):
                    ws.write_string(row, c, v if c in verbatim else keep(v), wrap_gap if c else wrap)
                else:
                    ws.write(row, c, v, wrap_gap if c else wrap)
            row += 1
        row += 1

    src_rows = []
    kinds = {"10-K": "Báo cáo năm", "10-Q": "Báo cáo quý", "8-K": "Thông cáo kết quả quý 4"}
    uses = {"scope": "Nguồn số liệu", "q4release": "Tỷ lệ % của quý 4 và đối chiếu số quý 4 tự tính",
            "check": "Chỉ để kiểm tra NIKE có in lại số cũ khác đi không"}
    for f in all_filings():
        # Short link texts: the full addresses broke mid-word; the address shows on hover, the accession identifies the file.
        src_rows.append([f["id"], document_title(f), (sec_document_url(f), "Mở trên SEC EDGAR"),
                         (ir_document_url(f), "Mở bản sao"), _vn_date(f["filed"]), f["accession"],
                         kinds[f["form"]], uses[f["role"]]])
    table("A. Báo cáo nguồn (Form 10-K, 10-Q và thông cáo kết quả quý 4 nộp SEC)",
          ["Báo cáo", "Tài liệu", "Liên kết SEC EDGAR", "Bản sao trên NIKE Investor Relations", "Ngày nộp SEC", "Số hồ sơ SEC",
           "Loại", "Dùng để"], src_rows, link_cols=(2, 3))
    table("B. Thuật ngữ", ["Thuật ngữ", "Tiếng Anh / tên trong báo cáo", "Định nghĩa", "Ghi chú"], [list(g) for g in GLOSSARY])
    table("C. Số liệu còn thiếu và giới hạn", ["Vấn đề", "Tình trạng", "Mô tả", "Cách xử lý"], [list(g) for g in GAPS])
    changed = [o for o in obs if o["restatement"].startswith("CÓ THAY ĐỔI")]
    repeated = [o for o in obs if o["restatement"].startswith((dataset.REPEATED, dataset.REPEATED_Q4))]
    table("D. Số cũ được in lại ở báo cáo năm sau", ["Kiểm tra", "Kết quả", "Chi tiết", "Ảnh hưởng"], [
        ["So với 10-K FY2020–FY2023 và 10-Q FY2020–FY2022", f"{len(changed)} số thay đổi",
         f"{len(repeated)} số được in lại ở ít nhất một báo cáo sau, tất cả giống số gốc (số Digital chỉ khác cách làm tròn, "
         "ví dụ 1.747 triệu USD và 1,7 tỷ USD).", "Không cần điều chỉnh." if not changed else f"Xem cột “{HEADER_OF['restatement']}”."]])
    counts = {s: sum(1 for o in obs if o["status"] == s) for s in STATUSES}
    meaning = {
        STATUS_MATCHED: "Đã kiểm tra lại và khớp: so với chỗ khác NIKE in cùng số (bảng trong MD&A, bảng số liệu quý của 10-K, "
                        "thông cáo quý 4, báo cáo năm sau), với số lũy kế, với số tính lại từ báo cáo tài chính, hoặc đọc lại "
                        "trên trang báo cáo.",
        STATUS_PENDING: "Chưa đối chiếu với nguồn khác.",
        STATUS_HUMAN: "Đã có thành viên tự kiểm tra trên báo cáo, có ghi tên và ngày.",
        STATUS_MISSING: "NIKE không nêu số này cho kỳ đó (không có trong 10-K, 10-Q và thông cáo kết quả quý 4). Để trống, "
                        "không điền 0.",
        STATUS_CONFLICT: "Các nguồn cho số khác nhau, chưa dùng số này.",
    }
    # Only the statuses that occur are listed.
    table("E. Trạng thái dữ liệu", ["Trạng thái", "Tiếng Anh", "Ý nghĩa", "Số lượng"],
          [[s, en, meaning[s], counts[s]] for s, en in STATUSES.items() if counts[s]])
    periods = dataset.PERIODS_A + dataset.PERIODS_Q
    checks_ok = all(results[p][c["id"]] is None or abs(results[p][c["id"]]) <= Decimal(str(c["tol"]))
                    for p in periods for c in CHECKS + COMPARISONS)
    ytd = dataset.ytd_checks(dataset.load_facts())
    pages = list(dataset.load_page_check().values())
    page_ok = sum(r["result"] == "match" for r in pages)
    page_on = max((r.get("checked_on") or extracted_on for r in pages), default=extracted_on)

    def tally(confirm, conflict, only=lambda o: True):
        """'matched/compared khớp' over the observations with a confirmation starting with `confirm`, or a conflict
        starting with `conflict` (either may be a tuple of prefixes)."""
        ok = sum(1 for o in obs if only(o) and any(c.startswith(confirm) for c in o["confirms"]))
        bad = sum(1 for o in obs if only(o) and any(c.startswith(conflict) for c in o["conflicts"]))
        return f"{ok}/{ok + bad} khớp"

    again = ("bảng doanh thu theo khu vực trong MD&A", "câu trong MD&A")
    table("F. Kết quả đối chiếu", ["Kiểm tra", "Kết quả", "Chi tiết", "Ngày"], [
        ["Thuyết minh so với bảng MD&A",
         tally(("khớp bảng doanh thu trong MD&A", "khớp bảng EBIT khu vực trong MD&A"), "MD&A (tr.", lambda o: o["method"] == "xbrl"),
         "Doanh thu theo kênh và theo khu vực, EBIT khu vực ở thuyết minh so với các bảng trong MD&A của chính báo cáo đó.",
         _vn_date(extracted_on)],
        ["Tỷ lệ % so với chỗ NIKE in lại trong cùng báo cáo", tally(tuple("khớp " + w for w in again), again),
         "Tỷ lệ tăng của NIKE, Inc. và NIKE Brand in lại ở bảng doanh thu theo khu vực; tỷ lệ loại trừ tỷ giá của NIKE Direct "
         "in lại trong một câu ở MD&A.", _vn_date(extracted_on)],
        ["Cộng các quý so với số lũy kế", f"{sum(c['ok'] for c in ytd)}/{len(ytd)} khớp",
         "Quý 1 + quý 2 so với số 6 tháng (10-Q quý 2), quý 1 + quý 2 + quý 3 so với số 9 tháng (10-Q quý 3). Chênh tối đa 1 "
         "triệu USD do làm tròn; với số Digital, chênh không quá mức làm tròn của NIKE.", _vn_date(extracted_on)],
        ["Số quý so với bảng số liệu quý của 10-K", tally("khớp bảng số liệu quý trong", "bảng số liệu quý của"),
         "10-K FY2019 và FY2020 có bảng doanh thu, lợi nhuận gộp và lợi nhuận sau thuế từng quý, dùng để kiểm tra cả số quý 4 "
         "tự tính. 10-K FY2021 không còn bảng này.", _vn_date(extracted_on)],
        ["Số quý 4 tự tính so với thông cáo kết quả quý 4",
         tally("khớp số quý 4 NIKE công bố trong thông cáo", "thông cáo 8-K", lambda o: o["method"] == "q4_derived"),
         "Thông cáo kết quả quý 4 (8-K) in số riêng của quý 4: báo cáo kết quả kinh doanh, doanh thu theo khu vực, EBIT. Mọi số "
         "quý 4 tự tính có trong thông cáo đều khớp; thông cáo không có NIKE Direct, bán buôn và Digital.", _vn_date(extracted_on)],
        ["Kiểm tra cộng trừ và so với tỷ lệ NIKE công bố", "Đạt" if checks_ok else "Có lỗi",
         "Xem mục E và F ở sheet 05 và 06.", _vn_date(extracted_on)],
        ["Đọc lại trên trang 10-Q", f"{page_ok}/{len(pages)} khớp",
         f"{len(pages)} tỷ lệ tăng trưởng loại trừ tỷ giá của quý 1–3 được đọc lại trên đúng trang 10-Q (tỷ lệ của bán buôn "
         "chỉ in một lần trong mỗi 10-Q). Ảnh chụp ở trang “Đối chiếu 36 số” trên dashboard web.", _vn_date(page_on)],
    ])
    quotes = load_quotes()
    table("G. NIKE tự giải thích trong 10-K (trích nguyên văn)", ["Nội dung", "Báo cáo, trang", "Trích dẫn (tiếng Anh)", "Ý chính"],
          [[QUOTES_BY_ID[q["id"]]["topic"], f"{q['filing']}, tr. {q['page']}", q["text"], QUOTES_BY_ID[q["id"]]["vi"]]
           for q in quotes], verbatim=(2,))
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.set_footer("&C&8Trang &P/&N")
    return ws


# ------------------------------------------------------------------ sheet 01
def write_overview_sheet(wb, st, obs, extracted_on):
    ws = wb.add_worksheet(SHEET_OVERVIEW)
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 30)
    ws.set_column(1, 1, 110)
    ws.write(0, 0, "Bộ dữ liệu tài chính NIKE, Inc. FY2019–FY2021", st(bold=True, font_size=16))
    ws.write(1, 0, "Số liệu theo năm và theo quý cho phần phân tích chiến lược bán trực tiếp (NIKE Direct) của NIKE.",
             st(font_color="#444444"))
    label = st(bold=True, text_wrap=True)
    wrap = st(text_wrap=True)
    h = st(bold=True, font_size=11, bottom=1)
    n = {s: sum(1 for o in obs if o["status"] == s) for s in STATUSES}
    total = len(obs)
    scope = [f for f in all_filings() if f["role"] == "scope"]
    blocks = [
        ("Thông tin chung", [
            ("Nội dung", "Doanh thu theo kênh (NIKE Direct, bán buôn) và theo khu vực, kết quả tài chính của NIKE, Inc., "
                         "FY2019–FY2021, theo năm và theo 12 quý."),
            ("Năm tài chính và quý", "Năm tài chính của NIKE kết thúc ngày 31/5, ví dụ FY2020 là từ 01/06/2019 đến 31/05/2020. "
                                     "Q1: tháng 6–8, Q2: tháng 9–11, Q3: tháng 12–2, Q4: tháng 3–5."),
            ("Đơn vị", "Triệu USD, trừ khi cột “Đơn vị” ghi khác (tỷ USD với doanh số Digital, % với các tỷ lệ)."),
            ("Nguồn", f"{sum(f['form'] == '10-K' for f in scope)} báo cáo năm Form 10-K và "
                      f"{sum(f['form'] == '10-Q' for f in scope)} báo cáo quý Form 10-Q mà NIKE, Inc. nộp cho SEC, cùng 3 "
                      "thông cáo kết quả quý 4 (Form 8-K) cho tỷ lệ % của quý 4. 5 báo cáo của các năm sau (đến 10-K FY2023) "
                      "chỉ dùng để kiểm tra NIKE có in lại số cũ khác đi không."),
            ("Quý 4", "Không có 10-Q cho quý 4, nên số quý 4 = số cả năm (10-K) trừ số 9 tháng đầu (in trong 10-Q quý 3). Số nào "
                      "có trong thông cáo kết quả quý 4 của NIKE (Form 8-K) đều khớp; thông cáo không có NIKE Direct, bán buôn và "
                      "Digital. Tỷ lệ % quý 4 lấy từ thông cáo đó. Riêng hàng tồn kho cuối quý 4 là số ngày 31/5 in trong 10-K."),
            ("Cùng kỳ năm trước", "Số của kỳ nào lấy từ báo cáo của chính kỳ đó. Báo cáo nào cũng in kèm số cùng kỳ năm trước "
                                  "để so sánh; cột “Cùng kỳ năm trước” ghi số này và tăng trưởng được tính từ nó. Riêng tồn kho "
                                  "cuối quý 1–3 lấy số cùng kỳ từ 10-Q năm trước, vì 10-Q chỉ in kèm số dư ngày 31/5."),
            ("Tình trạng", f"Lấy số ngày {_vn_date(extracted_on)}. {total} số liệu: " + ", ".join(
                f"{n[s]} {lab}" for s, lab in ((STATUS_MATCHED, "số đã được kiểm tra lại và khớp"),
                                               (STATUS_HUMAN, "số đã được người xác minh"), (STATUS_PENDING, "số chờ kiểm tra"),
                                               (STATUS_MISSING, "ô NIKE không công bố (tỷ lệ % của quý 4)"),
                                               (STATUS_CONFLICT, "số có mâu thuẫn")) if n[s]) + "."),
        ]),
        ("Cách dùng", [
            ("1. Đọc số liệu", "Sheet 02 (theo năm), 03 (theo quý) và 04 (bốn khu vực), mỗi dòng là một chỉ tiêu của một kỳ. "
                               "Dùng nút lọc ở dòng tiêu đề để chọn chỉ tiêu hoặc kỳ."),
            ("2. Xem nguồn", "Bấm link ở cột “Tài liệu nguồn (mở trên SEC)” để mở báo cáo, rồi tìm theo số trang ở cột “Trang” "
                             "(số in ở cuối trang)."),
            ("3. Trích vào bài", "Ghi nguồn là báo cáo của kỳ đó kèm số trang. Số quý 4 ghi là tự tính từ 10-K và 10-Q quý 3; "
                                 "tỷ lệ % quý 4 ghi nguồn là thông cáo kết quả quý 4 (8-K)."),
            ("4. Chỉ số", "Sheet 05 (năm) và 06 (quý) tính các tỷ lệ bằng công thức Excel; cột “Công thức” ghi cách tính, mục "
                          "E–F là các phép kiểm tra."),
            ("5. Tra cứu", "Sheet 07: danh sách báo cáo, thuật ngữ, số liệu còn thiếu, kết quả đối chiếu và các câu NIKE tự giải "
                           "thích trong 10-K."),
        ]),
        ("Quy ước", [
            ("Màu chữ", "Xanh dương là số nhập từ báo cáo, xanh lá là số lấy từ sheet khác, đen là số do công thức tính. Số quý 4 "
                        "tự tính để chữ nghiêng."),
            ("Số âm", "Để trong ngoặc đơn, ví dụ (790), như cách NIKE trình bày; tỷ lệ % dùng dấu trừ."),
            ("Ô trống", f"Số NIKE không công bố để trống, không điền 0. Ở sheet 05–06, “{NO_VALUE}” là không tính được vì thiếu số."),
            ("Tỷ lệ %", "Lưu dạng thập phân (0,14 hiển thị 14%) để dùng được trong công thức."),
            ("Cột đang ẩn", "Sheet 02–04 có ba cột ẩn: “Người xác minh”, “Ngày xác minh”, “Ghi chú của nhóm”, để thành viên ghi "
                            "lại khi tự kiểm tra một số trên báo cáo. Mọi số đều đã được kiểm tra tự động (cột “Đối chiếu”)."),
            ("Dấu phân cách", "Số được lưu dạng số, nên dấu hàng nghìn và dấu thập phân theo cài đặt vùng của máy: máy đặt tiếng "
                              "Việt hiện 39.117 và 31,6%; máy đặt tiếng Anh hiện 39,117 và 31.6%."),
        ]),
    ]
    row = 3
    for title, items in blocks:
        ws.write(row, 0, title, h)
        ws.write_blank(row, 1, None, h)
        ws.set_row(row, 20)     # room under the letters for the rule
        row += 1
        for lab, txt in items:
            ws.write(row, 0, lab, label)
            ws.write_string(row, 1, keep(txt), wrap)
            row += 1
        row += 1
    ws.write(row, 0, "Liên kết dashboard", label)
    ws.write_url(row, 1, DASHBOARD_URL, st(font_color="#1F4E9A", underline=1), DASHBOARD_URL)
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.set_footer("&C&8Trang &P/&N")
    ws.activate()
    return ws


# ------------------------------------------------------------------ main
def git_user_name():
    """The builder's name as set in git (the repository's own setting first), for the files' author field."""
    try:
        out = subprocess.run(["git", "config", "user.name"], cwd=ROOT, capture_output=True, encoding="utf-8")
    except OSError:
        return ""
    return out.stdout.strip()


def build(path=WORKBOOK, active_sheet=None):
    """Write the workbook; `active_sheet` (QA renders only) selects the sheet shown on open."""
    obs, results = dataset.assemble()
    merge_verifications(obs, previous_verifications(path))
    extracted_on = json.loads(META.read_text())["extracted_on"]

    path.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(path))
    wb.set_properties({"title": "NIKE Financial Dataset FY2019–FY2021",
                       "subject": "Dữ liệu tài chính NIKE, Inc. từ Form 10-K, 10-Q và thông cáo kết quả quý 4 (Form 8-K)",
                       "author": git_user_name(),
                       "comments": "Nguồn: NIKE, Inc. Form 10-K, 10-Q và thông cáo kết quả quý 4 (Form 8-K), FY2019–FY2021 "
                                   "(SEC EDGAR)."})
    st = Styles(wb)
    write_overview_sheet(wb, st, obs, extracted_on)
    write_observation_sheet(wb, st, SHEET_A, "Số liệu theo năm – FY2019 đến FY2021",
                            "Kết quả hợp nhất của NIKE, Inc. và doanh thu theo kênh của NIKE Brand. Mỗi dòng là một chỉ tiêu "
                            "của một năm tài chính.",
                            [o for o in obs if o["freq"] == A and o["metric"]["sheet"] != "regions"])
    write_observation_sheet(wb, st, SHEET_Q, "Số liệu theo quý – Q1 FY2019 đến Q4 FY2021",
                            "Cùng các chỉ tiêu theo 12 quý. Quý 1–3 lấy từ 10-Q. Không có 10-Q cho quý 4, nên quý 4 = cả "
                            "năm trừ 9 tháng đầu (chữ nghiêng); tỷ lệ % quý 4 lấy từ thông cáo kết quả quý 4 (8-K).",
                            [o for o in obs if o["freq"] == Q and o["metric"]["sheet"] != "regions"])
    write_observation_sheet(wb, st, SHEET_R, "Bốn khu vực của NIKE Brand",
                            "Doanh thu theo kênh của Bắc Mỹ, EMEA (châu Âu, Trung Đông và châu Phi), Đại Trung Hoa và APLA "
                            "(châu Á – Thái Bình Dương và Mỹ Latinh) theo năm và theo quý; EBIT khu vực theo năm.",
                            [o for o in obs if o["metric"]["sheet"] == "regions"])
    write_kpi_sheet(wb, st, SHEET_KPI_A, A, obs, results)
    write_kpi_sheet(wb, st, SHEET_KPI_Q, Q, obs, results)
    write_sources_sheet(wb, st, obs, results, extracted_on)
    if active_sheet:
        wb.get_worksheet_by_name(active_sheet).activate()
    wb.close()
    return obs, results


def main():
    obs, _ = build()
    print(f"Workbook written: {WORKBOOK.relative_to(ROOT)} ({len(obs)} observations)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
