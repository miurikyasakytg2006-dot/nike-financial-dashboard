"""Build the Excel handover workbook from the extracted facts.

Usage:  python -m pipeline.build_workbook

Writes handover/NIKE_Financial_Dataset_FY2020-2026.xlsx. If the workbook already exists,
the human-verification columns (status "Đã xác minh (người)", verifier, date, team note)
are carried over by observation key; a verified value that no longer matches the
source extraction is flagged as conflicting evidence instead of silently kept.
Every KPI and check on sheet 04 is a live Excel formula; the Python result of the same
expression is stored as the cached value so previews show numbers before recalculation.
"""
import datetime
import json
import subprocess
import sys
from decimal import Decimal
from pathlib import Path

import xlsxwriter
from openpyxl import load_workbook
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from . import dataset
from .metrics import (CHECKS, COMPARISONS, KPIS, KPIS_BY_ID, REPORTED, REPORTED_BY_ID, STATUS_CONFLICT,
                      STATUS_HUMAN, STATUS_MATCHED, STATUS_MISSING, STATUS_PENDING, STATUSES, USD_B, USD_M, PCT)
from .sources import FILINGS, document_title, filing, ir_document_url, sec_document_url, sec_index_url

ROOT = Path(__file__).resolve().parent.parent
WORKBOOK = ROOT / "handover" / "NIKE_Financial_Dataset_FY2020-2026.xlsx"
META = ROOT / "data" / "extracted" / "meta.json"
DASHBOARD_URL = "https://miurikyasakytg2006-dot.github.io/nike-financial-dashboard/"

SHEET_OVERVIEW = "01 Tổng quan"
SHEET_FIN = "02 Dữ liệu tài chính"
SHEET_CH = "03 Kênh phân phối"
SHEET_KPI = "04 KPI & Tính toán"
SHEET_SRC = "05 Nguồn & Định nghĩa"
HEADER_ROW = 3          # zero-based: tables start on Excel row 4
LAST_DATA_ROW = 400     # bounded lookup range used by sheet 04 formulas

# Observation table columns: (header, field, maximum width; actual width follows content)
OBS_COLUMNS = [
    ("Năm tài chính", "fy_label", 10),
    ("Chỉ tiêu", "vi", 64),
    ("Giá trị", "value", 12),
    ("Giá trị năm trước (cùng 10-K)", "prior", 14),
    ("Đơn vị", "unit", 10),
    ("Phạm vi báo cáo", "scope", 14),
    ("Kênh", "channel", 24),
    ("Loại số liệu", "kind", 46),
    ("Trạng thái dữ liệu", "status", 24),
    ("Tài liệu nguồn (mở trên SEC)", "doc", 14),
    ("Trang", "page", 7),
    ("Mục trong 10-K", "section", 44),
    ("Vị trí chi tiết", "location", 60),
    ("Số hoặc câu hiển thị trong tài liệu", "evidence", 50),
    ("Đối chiếu", "cross_check", 52),
    ("Trình bày lại ở 10-K sau", "restatement", 52),
    ("Người xác minh", "verified_by", 16),
    ("Ngày xác minh", "verified_on", 14),
    ("Ghi chú của nhóm", "team_note", 30),
    ("Dòng trong 10-K (tiếng Anh)", "en", 64),
    ("Thẻ XBRL", "xbrl", 60),
    ("Kỳ báo cáo", "period", 26),
    ("Ngày nộp SEC", "filed", 13),
    ("Số accession", "accession", 24),
    ("Mã chỉ tiêu", "metric_id", 42),
    ("Mã quan sát", "key", 50),
]
FIELD_COL = {f: i for i, (_, f, _) in enumerate(OBS_COLUMNS)}
FILL_IN_FIELDS = {"verified_by", "verified_on", "team_note"}
PROSE_FIELDS = {"location", "evidence", "cross_check", "restatement", "xbrl"}


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
        fmt = {USD_M: "#,##0;(#,##0);0", USD_B: "0.0", PCT: "0.0%;−0.0%;0.0%", "lần": "0.00", "ngày": "0.0"}.get(unit, "General")
        props = {"num_format": fmt, **extra}
        if color:
            props["font_color"] = color
        return self(**props)


def _d(v):
    return float(v) if isinstance(v, Decimal) else v


def period_label(m, fy):
    if m.get("xbrl", {}).get("period") == "instant":
        return f"Số dư tại 31/05/{fy}"
    return f"01/06/{fy - 1} – 31/05/{fy}"


def _vn_date(iso):
    y, mth, d = iso.split("-")
    return f"{d}/{mth}/{y}"


def text_width(s):
    """Display width in Excel column units, roughly: capitals and m, w, %, & are about a quarter wider than digits."""
    return sum(1.25 if ch.isupper() or ch in "mw%&" else 1 for ch in s)


def date_text(v):
    """A verification date as typed by the team; Excel may have stored it as a date."""
    if isinstance(v, (datetime.date, datetime.datetime)):
        return v.strftime("%d/%m/%Y")
    return "" if v is None else str(v)


# ------------------------------------------------------------------ carry-over of human verification
def previous_verifications(path):
    if not path.exists():
        return {}
    wb = load_workbook(path, read_only=True)
    out = {}
    for name in (SHEET_FIN, SHEET_CH):
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
        p = previous.get(o["key"])
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


# ------------------------------------------------------------------ sheets 02 / 03
def write_observation_sheet(wb, st, name, title, subtitle, obs):
    ws = wb.add_worksheet(name)
    ws.hide_gridlines(2)
    ws.write(0, 0, title, st(bold=True, font_size=14))
    ws.write(1, 0, subtitle, st(font_color="#444444"))
    head = st(bold=True, bg_color="#E7EAEE", text_wrap=True, valign="vcenter", bottom=1)
    head_fill = st(bold=True, bg_color="#FFF2CC", text_wrap=True, valign="vcenter", bottom=1)
    rows = []
    for o in obs:
        m = o["metric"]
        rows.append(dict(fy_label=f"FY{o['fy']}", vi=m["vi"], unit=m["unit"], scope=m["scope"], channel=m["channel"] or "—",
                         kind=m["kind"], status=o["status"], page=o["page"], section=o["section"] or "—",
                         location=m["location"], evidence=o["evidence"] or "—", cross_check=o["cross_check"],
                         restatement=o["restatement"] or "—", verified_by=o["verified_by"], verified_on=o["verified_on"],
                         team_note=o["team_note"], en=m["en"],
                         xbrl=(o["concept"] + (f" [{dataset.dims_label(o['dims'])}]" if o["dims"] else "")) or "Không gắn thẻ XBRL (MD&A)",
                         period=period_label(m, o["fy"]), filed=_vn_date(o["filed"]), accession=o["accession"],
                         metric_id=m["id"], key=o["key"]))
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
    link = st(font_color="#1F4E9A", underline=1)
    for i, (o, vals) in enumerate(zip(obs, rows)):
        r = HEADER_ROW + 1 + i
        m = o["metric"]
        f = filing(o["fy"])
        for field, c in FIELD_COL.items():
            if field in ("value", "prior"):
                v = o[field]
                # NIKE reports these percentages rounded to whole numbers, so show them that way.
                fmt = (st(num_format="0%;−0%;0%", font_color=Styles.INPUT_BLUE) if m["unit"] == PCT
                       else st.num(m["unit"], Styles.INPUT_BLUE))
                if v is None:
                    ws.write_blank(r, c, None, fmt)
                else:
                    ws.write_number(r, c, _d(v), fmt)
            elif field == "doc":
                ws.write_url(r, c, sec_document_url(f), link, f"10-K FY{o['fy']}",
                             tip=f"{document_title(f)} – SEC EDGAR")
            elif field == "page":
                if vals["page"] is None:
                    ws.write_blank(r, c, None, text)
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
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(HEADER_ROW)
    return ws


# ------------------------------------------------------------------ sheet 04
def readable(expr):
    """Turn a catalogue expression into a Vietnamese description of the formula."""
    def name(tok):
        base, prior = (tok[:-7], True) if tok.endswith("__prior") else (tok, False)
        if base in REPORTED_BY_ID:
            label = REPORTED_BY_ID[base]["vi"]
        elif base in KPIS_BY_ID:
            label = KPIS_BY_ID[base]["vi"]
        else:
            return tok
        return f"[{label}{' – năm trước' if prior else ''}]"
    return dataset.IDENT.sub(lambda mt: name(mt.group(0)), expr)


def write_kpi_sheet(wb, st, obs, results, sheet_cols):
    ws = wb.add_worksheet(SHEET_KPI)
    ws.hide_gridlines(2)
    years = dataset.YEARS
    first_col = 3
    ids = [m["id"] + "__prior" for m in REPORTED] + [k["id"] for k in KPIS + CHECKS + COMPARISONS]
    ws.set_column(0, 0, min(max(len(i) for i in ids) + 2, 44))
    written_labels = []
    ws.set_column(2, 2, 10)
    ws.set_column(first_col, first_col + len(years) - 1, 11)
    ws.set_column(first_col + len(years), first_col + len(years), 58)
    ws.set_column(first_col + len(years) + 1, first_col + len(years) + 1, 62)
    note_col = first_col + len(years)
    ws.write(0, 0, "Chỉ số tính toán (KPI) và kiểm tra đối chiếu", st(bold=True, font_size=14))
    ws.write(1, 0, "Mọi ô số trong sheet này là công thức Excel. Số màu xanh lá lấy từ sheet 02/03 theo năm và mã chỉ tiêu, "
                   "số màu đen được tính ngay trong sheet này. “–” là thiếu số đầu vào, không phải bằng 0.", st(font_color="#444444"))
    head = st(bold=True, bg_color="#E7EAEE", text_wrap=True, valign="vcenter", bottom=1)
    for c, h in enumerate(["Mã", "Chỉ tiêu", "Đơn vị"] + [f"FY{y}" for y in years] + ["Công thức", "Ghi chú và cách đọc"]):
        ws.write(HEADER_ROW, c, h, head if c < first_col or c >= note_col else st(bold=True, bg_color="#E7EAEE", align="right", valign="vcenter", bottom=1))
    ws.set_row(HEADER_ROW, 22)
    ws.freeze_panes(HEADER_ROW + 1, 3)
    section_fmt = st(bold=True, bg_color="#F3F4F6", top=1)
    row = HEADER_ROW + 1
    refs = {y: {} for y in years}

    def section(title):
        nonlocal row
        row += 1
        for c in range(note_col + 2):
            ws.write_blank(row, c, None, section_fmt)
        ws.write(row, 1, title, section_fmt)
        row += 1

    # A. inputs
    used = []
    for k in KPIS + CHECKS + COMPARISONS:
        for tok in dataset.IDENT.findall(k["expr"]):
            base = tok[:-7] if tok.endswith("__prior") else tok
            if base in REPORTED_BY_ID and tok not in used:
                used.append(tok)
    used.sort(key=lambda t: ([m["id"] for m in REPORTED].index(t.replace("__prior", "")), t.endswith("__prior")))
    section("A. Số liệu đầu vào (lấy từ sheet 02 và 03)")
    by_key = {(o["fy"], o["metric"]["id"]): o for o in obs}
    for tok in used:
        base, prior = (tok[:-7], True) if tok.endswith("__prior") else (tok, False)
        m = REPORTED_BY_ID[base]
        sheet = SHEET_FIN if m["sheet"] == "financial" else SHEET_CH
        val_col = xl_col_to_name(sheet_cols[sheet][("prior" if prior else "value")])
        key_col = xl_col_to_name(sheet_cols[sheet]["key"])
        ws.write(row, 0, tok, st(font_color="#6B6B6B"))
        label_text = m["vi"] + (" – năm trước (cùng 10-K)" if prior else "")
        written_labels.append(label_text)
        ws.write(row, 1, label_text, st())
        ws.write(row, 2, m["unit"], st())
        rng_v = f"'{sheet}'!${val_col}${HEADER_ROW + 2}:${val_col}${LAST_DATA_ROW + 1}"
        rng_k = f"'{sheet}'!${key_col}${HEADER_ROW + 2}:${key_col}${LAST_DATA_ROW + 1}"
        for j, y in enumerate(years):
            col = first_col + j
            hdr = xl_rowcol_to_cell(HEADER_ROW, col, row_abs=True)
            look = f'INDEX({rng_v},MATCH({hdr}&"|{base}",{rng_k},0))'
            o = by_key[(y, base)]
            cached = o["prior"] if prior else o["value"]
            fmt = st(num_format="0%;−0%;0%", font_color=Styles.LINK_GREEN) if m["unit"] == PCT else st.num(m["unit"], Styles.LINK_GREEN)
            # A missing source value shows "–" (text), never a blank a previewer could render as 0.
            ws.write_formula(row, col, f'=IF({look}="","–",{look})', fmt, "–" if cached is None else _d(cached))
            refs[y][tok] = xl_rowcol_to_cell(row, col)
        ws.write(row, note_col, f"Tra cứu sheet “{sheet}” theo khóa FY|{base}, cột “{'Giá trị năm trước (cùng 10-K)' if prior else 'Giá trị'}”.",
                 st(font_color="#6B6B6B"))
        row += 1

    def write_expr_rows(items, title, kind):
        nonlocal row
        section(title)
        for k in items:
            unit = k.get("unit", USD_M if kind == "check" else PCT)
            ws.write(row, 0, k["id"], st(font_color="#6B6B6B"))
            written_labels.append(k["vi"])
            ws.write(row, 1, k["vi"], st())
            ws.write(row, 2, {"check": USD_M, "cmp": "điểm %"}.get(kind, unit), st())
            toks = list(dict.fromkeys(dataset.IDENT.findall(k["expr"])))
            # Comparisons are differences between two percentages; show them in percentage points (×100).
            scale = 100 if kind == "cmp" else 1
            for j, y in enumerate(years):
                col = first_col + j
                cells = [refs[y][t] for t in toks]
                body = dataset.IDENT.sub(lambda mt: refs[y][mt.group(0)], k["expr"]).replace(" ", "")
                if scale != 1:
                    body = f"({body})*{scale}"
                formula = f'=IF(COUNT({",".join(cells)})<{len(cells)},"–",{body})'
                v = results[y][k["id"]]
                if kind == "cmp":
                    fmt = st(num_format="+0.00;−0.00;0.00", align="right")
                elif k.get("num_format"):
                    fmt = st(num_format=k["num_format"], align="right")
                else:
                    fmt = st.num(unit, None, align="right")
                ws.write_formula(row, col, formula, fmt, "–" if v is None else _d(v) * scale)
                refs[y][k["id"]] = xl_rowcol_to_cell(row, col)
            first = xl_rowcol_to_cell(row, first_col)
            last = xl_rowcol_to_cell(row, first_col + len(years) - 1)
            if kind in ("check", "cmp"):
                tol = k["tol"]
                lim = f"{tol * scale:g}"
                tol_txt = f"±{tol} triệu USD" if kind == "check" else f"±{lim} điểm %"
                ws.write(row, note_col, f"{readable(k['expr'])}   (ngưỡng {tol_txt})", st(text_wrap=True))
                ok = all(results[y][k["id"]] is None or abs(results[y][k["id"]]) <= Decimal(str(tol)) for y in years)
                ws.write_formula(row, note_col + 1,
                                 f'=IF(AND(MAX({first}:{last})<={lim},MIN({first}:{last})>=-{lim}),"Đạt","Kiểm tra lại")',
                                 st(bold=True), "Đạt" if ok else "Kiểm tra lại")
            else:
                ws.write(row, note_col, readable(k["expr"]), st(text_wrap=True))
                ws.write(row, note_col + 1, k["note"], st(text_wrap=True))
            row += 1

    write_expr_rows([k for k in KPIS if k["group"] == "channels"], "B. Chỉ số kênh phân phối – phạm vi NIKE Brand", "kpi")
    write_expr_rows([k for k in KPIS if k["group"] == "financial"], "C. Chỉ số kết quả tài chính – phạm vi NIKE, Inc.", "kpi")
    write_expr_rows(CHECKS, "D. Kiểm tra các phép cộng trừ (đúng thì chênh lệch bằng 0, cột cuối là kết luận)", "check")
    write_expr_rows(COMPARISONS, "E. So với tỷ lệ NIKE tự công bố (các ô là phần chênh lệch, cột cuối là kết luận)", "cmp")
    ws.set_column(1, 1, min(max(len(t) for t in written_labels) + 2, 66))
    ws.conditional_format(HEADER_ROW + 1, note_col + 1, row, note_col + 1,
                          {"type": "cell", "criteria": "==", "value": '"Kiểm tra lại"', "format": st(font_color="#C00000", bold=True)})
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    return ws


# ------------------------------------------------------------------ sheet 05
GLOSSARY = [
    ("Năm tài chính (FY)", "Fiscal year", "Năm tài chính của NIKE kết thúc ngày 31/5, ví dụ FY2026 là từ 01/06/2025 đến 31/05/2026.", ""),
    ("NIKE, Inc.", "NIKE, Inc. (consolidated)", "Toàn tập đoàn: NIKE Brand + Converse + Corporate.", "Phạm vi của sheet 02."),
    ("NIKE Brand", "NIKE Brand", "Thương hiệu NIKE, gồm cả Jordan Brand, không gồm Converse.", "Phạm vi của sheet 03."),
    ("NIKE Direct", "NIKE Direct / Sales through Direct to Consumer",
     "Doanh thu NIKE Brand bán thẳng cho người tiêu dùng, qua cửa hàng của NIKE và kênh online (NIKE Brand Digital).",
     "NIKE không công bố lợi nhuận riêng của kênh này."),
    ("NIKE Brand Digital", "NIKE Brand Digital sales (FY2020–21: Digital commerce sales)",
     "Phần doanh số NIKE Direct bán qua ứng dụng và website của NIKE.", "Chỉ có trong MD&A, làm tròn đến 0,1 tỷ USD."),
    ("Bán buôn (Wholesale)", "Sales to Wholesale Customers",
     "Doanh thu NIKE Brand bán cho nhà bán lẻ, nhà phân phối và bên nhận cấp phép.", "Không gồm Converse."),
    ("Global Brand Divisions", "Global Brand Divisions",
     "Doanh thu cấp phép thương hiệu và các khoản khác không thuộc khu vực địa lý nào.", "Dưới 0,3% doanh thu NIKE Brand."),
    ("Converse", "Converse", "Mảng kinh doanh riêng của NIKE, Inc., có kênh bán trực tiếp và bán buôn riêng.",
     "Không tính vào NIKE Direct và bán buôn ở sheet 03."),
    ("Corporate", "Corporate revenues", "Chủ yếu là lãi/lỗ phòng ngừa rủi ro tỷ giá liên quan đến doanh thu.",
     "Có thể âm. Chỉ dùng khi cộng đối chiếu ra doanh thu NIKE, Inc."),
    ("Biên lợi nhuận gộp", "Gross margin", "Lợi nhuận gộp / Doanh thu.", "Chỉ có ở phạm vi NIKE, Inc.; NIKE không công bố biên gộp theo kênh."),
    ("Chi phí bán hàng và quản lý (SG&A)", "Total selling and administrative expense",
     "Chi phí tạo nhu cầu (Demand creation) cộng chi phí vận hành chung (Operating overhead).",
     "Chi phí chung của cả công ty, không chia được theo kênh."),
    ("EBIT", "Earnings Before Interest and Taxes",
     "Lợi nhuận sau thuế + Chi phí (thu nhập) lãi thuần + Chi phí thuế thu nhập, theo định nghĩa của NIKE.", "Non-GAAP, phạm vi NIKE, Inc."),
    ("Lợi nhuận hoạt động", "Operating income",
     "NIKE không có dòng “Operating income” trên báo cáo kết quả kinh doanh.", "Dùng EBIT thay thế."),
    ("Tăng trưởng theo báo cáo", "Reported growth (% change)", "Thay đổi so với năm trước, tính bằng USD theo số báo cáo.",
     "Tính ở sheet 04 từ số năm nay và số năm trước trong cùng 10-K."),
    ("Tăng trưởng loại trừ tỷ giá", "Currency-neutral growth (% change excluding currency changes)",
     "Tăng trưởng tính theo tỷ giá của năm trước, không tính phần thay đổi do tỷ giá.", "Non-GAAP do NIKE tính, lấy nguyên văn từ MD&A."),
    ("GAAP / Non-GAAP", "U.S. GAAP / non-GAAP financial measure",
     "GAAP: số theo chuẩn mực kế toán Mỹ, nằm trong báo cáo tài chính đã kiểm toán. Non-GAAP: chỉ tiêu công ty tự định nghĩa, "
     "nằm ngoài báo cáo tài chính.", "Xem cột “Loại số liệu” ở sheet 02–03."),
    ("Số năm trước (cùng 10-K)", "Comparative prior-year amount",
     "Số của năm trước in kèm trong 10-K năm hiện tại để so sánh.", "Dùng làm mẫu số khi tính tăng trưởng."),
    ("Trình bày lại / phân loại lại", "Restatement / reclassification",
     "Khi 10-K năm sau trình bày lại số của năm trước theo cách phân loại mới.",
     "Giữ số gốc, chênh lệch ghi ở cột “Trình bày lại ở 10-K sau”."),
    ("XBRL", "eXtensible Business Reporting Language",
     "Nhãn dữ liệu gắn vào các con số trong báo cáo tài chính của 10-K để máy đọc được.",
     "Cột “Thẻ XBRL” ghi thẻ của từng số; số trong MD&A không có thẻ."),
]

GAPS = [
    ("Lợi nhuận và chi phí theo kênh", "Không công bố",
     "NIKE không công bố lợi nhuận, biên gộp hay chi phí riêng của NIKE Direct và bán buôn.",
     "Không ước tính. Các chỉ số lợi nhuận chỉ tính cho NIKE, Inc."),
    ("Tỷ trọng NIKE Direct FY2026 do NIKE nêu", "Thiếu",
     "10-K FY2026 không có câu “approximately …% of total NIKE Brand revenues”.",
     "Để trống. Tỷ trọng tự tính ở sheet 04 vẫn đủ 7 năm."),
    ("Doanh số Digital", "Độ chính xác thấp",
     "Chỉ có trong MD&A, làm tròn đến 0,1 tỷ USD. Số FY2023 được trình bày lại từ 12,6 xuống 12,4 tỷ USD.",
     "Giữ số gốc, có ghi chú. Tỷ trọng Digital trong NIKE Direct chỉ là xấp xỉ."),
    ("Tăng trưởng loại trừ tỷ giá", "Không tự tính lại được",
     "Muốn tính lại cần doanh thu theo từng loại tiền tệ, mà NIKE không công bố.",
     "Lấy nguyên văn từ bảng doanh thu trong MD&A, đã đọc lại trên trang 10-K."),
]

RESTATEMENT_LOG = [
    ("10-K FY2024", "NIKE Direct: cửa hàng và Digital",
     "“Within NIKE Direct revenues, there were certain reclassifications made between NIKE-owned retail stores and NIKE Brand Digital in the prior period to conform to current period presentation.”",
     "Doanh số Digital FY2023: 12,6 tỷ USD (10-K FY2023) được trình bày lại thành 12,4 tỷ USD. Tổng NIKE Direct không đổi."),
    ("10-K FY2022, FY2025, FY2026", "Các lần phân loại lại khác",
     "FY2022: doanh thu chia theo nhóm người tiêu dùng Nam, Nữ, Trẻ em (hàng unisex xếp vào Nam) và báo cáo riêng Jordan Brand. "
     "FY2025: bỏ chỉ tiêu “wholesale equivalent revenues”. FY2026: phân loại lại một số khoản năm trước (Note 1).",
     "Không ảnh hưởng tới các chỉ tiêu ở đây."),
]


def write_sources_sheet(wb, st, obs, results, extracted_on):
    ws = wb.add_worksheet(SHEET_SRC)
    ws.hide_gridlines(2)
    for c, w in enumerate([26, 30, 70, 50, 14, 22, 26, 40]):
        ws.set_column(c, c, w)
    ws.write(0, 0, "Nguồn và định nghĩa", st(bold=True, font_size=14))
    head = st(bold=True, bg_color="#E7EAEE", text_wrap=True, valign="vcenter", bottom=1)
    title = st(bold=True, font_size=11)
    wrap = st(text_wrap=True)
    link = st(font_color="#1F4E9A", underline=1, text_wrap=True)
    row = 2

    def table(name, headers, rows, link_cols=()):
        nonlocal row
        ws.write(row, 0, name, title)
        row += 1
        for c, h in enumerate(headers):
            ws.write(row, c, h, head)
        row += 1
        for r in rows:
            for c, v in enumerate(r):
                if c in link_cols and isinstance(v, tuple):
                    ws.write_url(row, c, v[0], link, v[1])
                else:
                    ws.write(row, c, v, wrap)
            row += 1
        row += 1

    src_rows = []
    for fy, *_ in FILINGS:
        f = filing(fy)
        src_rows.append([f"FY{fy} ({_vn_date(f'{fy - 1}-06-01')} – {_vn_date(f['period_end'])})", document_title(f),
                         (sec_document_url(f), sec_document_url(f)), (ir_document_url(f), "Bản sao trên NIKE Investor Relations"),
                         _vn_date(f["filed"]), f["accession"]])
    table("A. Báo cáo nguồn (Form 10-K nộp SEC)",
          ["Năm tài chính và kỳ báo cáo", "Tài liệu", "Liên kết SEC EDGAR", "Bản sao trên NIKE Investor Relations",
           "Ngày nộp SEC", "Số accession"], src_rows, link_cols=(2, 3))
    table("B. Thuật ngữ", ["Thuật ngữ", "Tiếng Anh / tên trong 10-K", "Định nghĩa", "Ghi chú"],
          [list(g) for g in GLOSSARY])
    table("C. Số liệu còn thiếu và giới hạn", ["Vấn đề", "Tình trạng", "Mô tả", "Cách xử lý"],
          [list(g) for g in GAPS])
    table("D. Các lần NIKE trình bày lại hoặc phân loại lại số liệu", ["Tài liệu", "Nội dung", "Trích dẫn từ 10-K", "Ảnh hưởng"],
          [list(g) for g in RESTATEMENT_LOG])
    counts = {s: sum(1 for o in obs if o["status"] == s) for s in STATUSES}
    meaning = {
        STATUS_MATCHED: "Đã đối chiếu với ít nhất một nguồn khác: dữ liệu XBRL của SEC, bảng trong MD&A, số tính lại từ báo cáo "
                        "tài chính hoặc trang 10-K.",
        STATUS_PENDING: "Chưa đối chiếu với nguồn khác.",
        STATUS_HUMAN: "Đã có thành viên tự kiểm tra trên 10-K, có ghi tên và ngày.",
        STATUS_MISSING: "NIKE không công bố số này trong 10-K năm đó. Để trống, không điền 0.",
        STATUS_CONFLICT: "Các nguồn cho số khác nhau, chưa dùng số này.",
    }
    # Only the statuses that occur are listed.
    table("E. Trạng thái dữ liệu", ["Trạng thái", "Tiếng Anh", "Ý nghĩa", "Số lượng"],
          [[s, en, meaning[s], counts[s]] for s, en in STATUSES.items() if counts[s]])
    checks_ok = all(all(results[y][c["id"]] is None or abs(results[y][c["id"]]) <= Decimal(str(c["tol"])) for y in dataset.YEARS)
                    for c in CHECKS + COMPARISONS)
    sec_rows = list(dataset.load_sec_check().values())
    pages = list(dataset.load_page_check().values())
    page_ok = sum(r["result"] == "match" for r in pages)
    page_on = max((r.get("checked_on") or extracted_on for r in pages), default=extracted_on)
    table("F. Kết quả đối chiếu", ["Kiểm tra", "Kết quả", "Chi tiết", "Ngày"], [
        ["Đối chiếu với dữ liệu XBRL của SEC", f"{sec_rows.count('match')}/{len(sec_rows)} khớp",
         "Số hợp nhất của 7 báo cáo, kể cả số năm trước in kèm, so với dữ liệu SEC của đúng báo cáo đó.", _vn_date(extracted_on)],
        ["Thuyết minh so với bảng MD&A", "Khớp toàn bộ",
         "Doanh thu theo kênh ở thuyết minh “Revenues” so với bảng doanh thu trong MD&A của cùng 10-K.", _vn_date(extracted_on)],
        ["Kiểm tra cộng trừ và so với tỷ lệ NIKE công bố", "Đạt" if checks_ok else "Có lỗi",
         "Xem mục D và E ở sheet 04.", _vn_date(extracted_on)],
        ["Đọc lại trên trang 10-K", f"{page_ok}/{len(pages)} khớp",
         "35 số chỉ có trong MD&A (28 tỷ lệ tăng trưởng loại trừ tỷ giá, 7 số Digital) được đọc lại trên đúng trang 10-K. "
         "Ảnh chụp ở trang “Đối chiếu 35 số” trên dashboard web.", _vn_date(page_on)],
    ])
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    return ws


# ------------------------------------------------------------------ sheet 01
def write_overview_sheet(wb, st, obs, extracted_on):
    ws = wb.add_worksheet(SHEET_OVERVIEW)
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 30)
    ws.set_column(1, 1, 110)
    ws.write(0, 0, "Bộ dữ liệu tài chính NIKE, Inc. FY2020–FY2026", st(bold=True, font_size=16))
    ws.write(1, 0, "Số liệu cho đề tài “Đổi mới và tái cân bằng chiến lược phân phối tại Nike: Từ Direct-to-Consumer (D2C) "
                   "đến bán buôn và những biến động tài chính giai đoạn FY2020–FY2026”.", st(font_color="#444444"))
    label = st(bold=True, text_wrap=True)
    wrap = st(text_wrap=True)
    h = st(bold=True, font_size=11, bottom=1)
    n = {s: sum(1 for o in obs if o["status"] == s) for s in STATUSES}
    total = len(obs)
    blocks = [
        ("Thông tin chung", [
            ("Nội dung", "Doanh thu theo kênh (NIKE Direct, bán buôn) và kết quả tài chính của NIKE, Inc., FY2020–FY2026."),
            ("Năm tài chính", "Năm tài chính của NIKE kết thúc ngày 31/5, ví dụ FY2026 là từ 01/06/2025 đến 31/05/2026."),
            ("Đơn vị", "Triệu USD, trừ khi cột “Đơn vị” ghi khác (tỷ USD với doanh số Digital, % với các tỷ lệ)."),
            ("Nguồn", "7 báo cáo thường niên Form 10-K mà NIKE, Inc. nộp cho SEC. Số trên báo cáo tài chính và thuyết minh lấy từ "
                      "phần Item 8; EBIT, các tỷ lệ % NIKE công bố và doanh số Digital lấy từ phần MD&A (Item 7)."),
            ("Số năm trước", "Số của năm nào lấy từ 10-K của chính năm đó. Cột “Giá trị năm trước (cùng 10-K)” là số năm trước in kèm "
                             "trong cùng báo cáo, dùng để tính tăng trưởng. Nếu báo cáo năm sau trình bày lại, chênh lệch ghi ở cột "
                             "“Trình bày lại ở 10-K sau”."),
            ("Tình trạng", f"Lấy số ngày {_vn_date(extracted_on)}. {total} số liệu: " + ", ".join(
                f"{n[s]} {label}" for s, label in ((STATUS_MATCHED, "khớp nguồn chính thức"), (STATUS_HUMAN, "đã được người xác minh"),
                                                   (STATUS_PENDING, "chờ kiểm tra"), (STATUS_MISSING, "số NIKE không công bố"),
                                                   (STATUS_CONFLICT, "mâu thuẫn")) if n[s]) + "."),
        ]),
        ("Cách dùng", [
            ("1. Đọc số liệu", "Sheet 02 (hợp nhất NIKE, Inc.) và 03 (kênh phân phối NIKE Brand), mỗi dòng là một chỉ tiêu của một năm. "
                               "Dùng nút lọc ở dòng tiêu đề để chọn chỉ tiêu hoặc năm."),
            ("2. Xem nguồn", "Bấm link ở cột “Tài liệu nguồn (mở trên SEC)” để mở báo cáo 10-K, rồi tìm theo số trang ở cột “Trang” "
                             "(số in ở chân trang, dạng “2026 FORM 10-K 83”)."),
            ("3. Trích vào bài", "Ghi nguồn là 10-K của năm đó kèm số trang."),
            ("4. Chỉ số", "Sheet 04 tính các tỷ lệ bằng công thức Excel; cột “Công thức” ghi cách tính, mục D–E là các phép kiểm tra."),
            ("5. Tra cứu", "Sheet 05: danh sách 10-K, thuật ngữ, số liệu còn thiếu và các lần NIKE trình bày lại số liệu."),
        ]),
        ("Quy ước", [
            ("Màu chữ", "Xanh dương là số nhập từ báo cáo, xanh lá là số lấy từ sheet khác, đen là số do công thức tính."),
            ("Số âm", "Để trong ngoặc đơn, ví dụ (50), như cách NIKE trình bày; tỷ lệ % dùng dấu trừ."),
            ("Ô trống", "Số NIKE không công bố để trống, không điền 0. Ở sheet 04, “–” là không tính được vì thiếu số."),
            ("Tỷ lệ %", "Lưu dạng thập phân (0,14 hiển thị 14%) để dùng được trong công thức."),
        ]),
    ]
    row = 3
    for title, items in blocks:
        ws.write(row, 0, title, h)
        ws.write_blank(row, 1, None, h)
        row += 1
        for lab, txt in items:
            ws.write(row, 0, lab, label)
            ws.write(row, 1, txt, wrap)
            row += 1
        row += 1
    ws.write(row, 0, "Liên kết dashboard", label)
    ws.write_url(row, 1, DASHBOARD_URL, st(font_color="#1F4E9A", underline=1), DASHBOARD_URL)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
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
    facts = dataset.load_facts()
    obs = dataset.build_observations(facts, dataset.load_sec_check())
    results = dataset.compute(obs)
    dataset.apply_cross_check_statuses(obs, results)
    dataset.apply_page_checks(obs, dataset.load_page_check())
    merge_verifications(obs, previous_verifications(path))
    extracted_on = json.loads(META.read_text())["extracted_on"]

    path.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(path))
    wb.set_properties({"title": "NIKE Financial Dataset FY2020–FY2026", "subject": "Dữ liệu tài chính NIKE, Inc. từ Form 10-K",
                       "author": git_user_name(), "comments": "Nguồn: NIKE, Inc. Form 10-K FY2020–FY2026 (SEC EDGAR)."})
    st = Styles(wb)
    write_overview_sheet(wb, st, obs, extracted_on)
    fin = [o for o in obs if o["metric"]["sheet"] == "financial"]
    ch = [o for o in obs if o["metric"]["sheet"] == "channels"]
    write_observation_sheet(wb, st, SHEET_FIN, "Dữ liệu tài chính hợp nhất – NIKE, Inc.",
                            "Báo cáo kết quả kinh doanh, hàng tồn kho, EBIT và tăng trưởng doanh thu NIKE công bố. Mỗi dòng là một chỉ tiêu của một năm tài chính.", fin)
    write_observation_sheet(wb, st, SHEET_CH, "Kênh phân phối – NIKE Brand",
                            "Doanh thu NIKE Direct, bán buôn, Global Brand Divisions, doanh số Digital và tăng trưởng theo kênh NIKE công bố. "
                            "Converse và Corporate chỉ để đối chiếu.", ch)
    sheet_cols = {s: FIELD_COL for s in (SHEET_FIN, SHEET_CH)}
    write_kpi_sheet(wb, st, obs, results, sheet_cols)
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
