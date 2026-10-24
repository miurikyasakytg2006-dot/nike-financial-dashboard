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
    ("Loại số liệu", "kind", 30),
    ("Trạng thái dữ liệu", "status", 24),
    ("Tài liệu nguồn (mở trên SEC)", "doc", 14),
    ("Trang", "page", 7),
    ("Mục trong 10-K", "section", 40),
    ("Vị trí chi tiết", "location", 60),
    ("Số hoặc câu hiển thị trong tài liệu", "evidence", 50),
    ("Đối chiếu tự động", "cross_check", 52),
    ("Trình bày lại ở 10-K sau", "restatement", 52),
    ("Người xác minh", "verified_by", 16),
    ("Ngày xác minh", "verified_on", 14),
    ("Ghi chú của nhóm", "team_note", 30),
    ("Dòng trong 10-K (tiếng Anh)", "en", 60),
    ("Thẻ XBRL", "xbrl", 60),
    ("Kỳ báo cáo", "period", 26),
    ("Ngày nộp SEC", "filed", 13),
    ("Số accession", "accession", 24),
    ("Mã chỉ tiêu", "metric_id", 42),
    ("Mã quan sát", "key", 50),
]
FIELD_COL = {f: i for i, (_, f, _) in enumerate(OBS_COLUMNS)}
FILL_IN_FIELDS = {"verified_by", "verified_on", "team_note"}


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
        # Width follows the longest entry (capped for long prose); headers wrap to at most three lines.
        if field in ("value", "prior", "page", "doc"):
            width = cap
        else:
            longest = max((len(str(rv[field] or "")) for rv in rows), default=0)
            width = max(min(longest + 2, cap), max(len(w) for w in h.split()) + 2, len(h) / 3 + 2, 8)
        ws.set_column(c, c, width)
        ws.write(HEADER_ROW, c, h, head_fill if field in FILL_IN_FIELDS else head)
    ws.set_row(HEADER_ROW, 44)
    text = st()
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
                ws.write_string(r, c, str(vals[field] if vals[field] is not None else ""), text)
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
    ws.write(1, 0, "Mọi ô số trong sheet này là công thức Excel. Chữ xanh lá: lấy từ sheet 02/03 theo năm và mã chỉ tiêu; "
                   "chữ đen: tính trong sheet này. Sửa số liệu ở sheet 02/03 thì sheet này tự cập nhật. "
                   "“–” nghĩa là thiếu số liệu đầu vào, không phải bằng 0.", st(font_color="#444444"))
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
    write_expr_rows(CHECKS, "D. Kiểm tra đối chiếu số học (kỳ vọng bằng 0; cột cuối là kết quả)", "check")
    write_expr_rows(COMPARISONS, "E. Đối chiếu với tỷ lệ NIKE tự công bố (chênh lệch; cột cuối là kết quả)", "cmp")
    ws.set_column(1, 1, min(max(len(t) for t in written_labels) + 2, 66))
    ws.conditional_format(HEADER_ROW + 1, note_col + 1, row, note_col + 1,
                          {"type": "cell", "criteria": "==", "value": '"Kiểm tra lại"', "format": st(font_color="#C00000", bold=True)})
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    return ws


# ------------------------------------------------------------------ sheet 05
GLOSSARY = [
    ("Năm tài chính (FY)", "Fiscal year", "Năm tài chính của NIKE kết thúc ngày 31/5. FY2026 là giai đoạn 01/06/2025 – 31/05/2026.",
     "Không quy đổi sang năm dương lịch. Khi so sánh với công ty khác cần lưu ý lệch kỳ."),
    ("NIKE, Inc.", "NIKE, Inc. (consolidated)", "Toàn bộ tập đoàn: NIKE Brand + Converse + Corporate.",
     "Các chỉ tiêu ở sheet 02 (doanh thu, giá vốn, SG&A, EBIT, tồn kho) thuộc phạm vi này."),
    ("NIKE Brand", "NIKE Brand", "Thương hiệu NIKE (gồm cả Jordan Brand), chia theo 4 khu vực địa lý và Global Brand Divisions. Không gồm Converse.",
     "Phạm vi của các chỉ tiêu kênh phân phối ở sheet 03."),
    ("NIKE Direct", "NIKE Direct / Sales through Direct to Consumer",
     "Doanh thu NIKE Brand bán trực tiếp cho người tiêu dùng qua cửa hàng do NIKE sở hữu và nền tảng số (NIKE Brand Digital).",
     "Là doanh thu, không phải lợi nhuận. NIKE không công bố lợi nhuận riêng của kênh này."),
    ("NIKE Brand Digital", "NIKE Brand Digital sales (FY2020–21: Digital commerce sales)",
     "Phần doanh số NIKE Direct qua ứng dụng và website của NIKE.",
     "Là một phần của NIKE Direct. Chỉ công bố trong MD&A, làm tròn 0,1 tỷ USD."),
    ("Bán buôn (Wholesale)", "Sales to Wholesale Customers",
     "Doanh thu NIKE Brand bán cho nhà bán lẻ, nhà phân phối độc lập, bên nhận cấp phép.", "Không gồm bán buôn của Converse."),
    ("Global Brand Divisions", "Global Brand Divisions", "Doanh thu cấp phép thương hiệu NIKE và doanh thu khác không thuộc khu vực địa lý.",
     "Rất nhỏ (dưới 0,3% doanh thu NIKE Brand); không thuộc Direct hay Wholesale."),
    ("Converse", "Converse", "Phân khúc kinh doanh riêng của NIKE, Inc., có kênh trực tiếp và bán buôn riêng.",
     "Không nằm trong số liệu NIKE Direct/Wholesale của bộ dữ liệu."),
    ("Corporate", "Corporate revenues", "Chủ yếu là lãi/lỗ phòng ngừa rủi ro tỷ giá liên quan đến doanh thu, quản lý tập trung.",
     "Có thể âm. Chỉ dùng để đối chiếu NIKE Brand + Converse + Corporate = NIKE, Inc."),
    ("Biên lợi nhuận gộp", "Gross margin", "Lợi nhuận gộp / Doanh thu.", "Phạm vi NIKE, Inc. NIKE không công bố biên gộp theo kênh."),
    ("Chi phí bán hàng và quản lý (SG&A)", "Total selling and administrative expense",
     "Tổng của chi phí tạo nhu cầu (Demand creation) và chi phí vận hành chung (Operating overhead).",
     "Chi phí hợp nhất. Không thể phân bổ tin cậy cho NIKE Direct hay Wholesale từ số liệu công bố."),
    ("EBIT", "Earnings Before Interest and Taxes",
     "Lợi nhuận sau thuế + Chi phí (thu nhập) lãi thuần + Chi phí thuế thu nhập (định nghĩa của NIKE).",
     "Non-GAAP ở cấp NIKE, Inc. Đây là thước đo hiệu quả hoạt động NIKE dùng."),
    ("Lợi nhuận hoạt động", "Operating income",
     "NIKE không trình bày dòng “Operating income” trên báo cáo kết quả kinh doanh.",
     "Bộ dữ liệu không tự tạo chỉ tiêu này; dùng EBIT và ghi rõ là non-GAAP. Các trang tổng hợp bên thứ ba thường tự tính, không phải số NIKE công bố."),
    ("Tăng trưởng theo báo cáo", "Reported growth (% change)", "Thay đổi so với năm trước tính bằng USD theo báo cáo.",
     "Sheet 04 tự tính từ số năm nay và số năm trước trong cùng 10-K."),
    ("Tăng trưởng loại trừ tỷ giá", "Currency-neutral growth (% change excluding currency changes)",
     "Tăng trưởng tính lại theo tỷ giá của năm trước để loại ảnh hưởng chuyển đổi ngoại tệ.",
     "Non-GAAP do NIKE tính; không tự tính lại được từ BCTC. Lấy nguyên văn từ MD&A."),
    ("GAAP / Non-GAAP", "U.S. GAAP / non-GAAP financial measure",
     "GAAP: số liệu theo chuẩn mực kế toán Mỹ trong BCTC đã kiểm toán. Non-GAAP: chỉ tiêu do công ty tự định nghĩa, trình bày ngoài BCTC.",
     "Cột “Loại số liệu” ở sheet 02–03 ghi rõ từng trường hợp."),
    ("Số năm trước (cùng 10-K)", "Comparative prior-year amount",
     "Số của năm trước được trình bày làm cột so sánh trong 10-K của năm hiện tại.", "Dùng làm mẫu số khi tính tăng trưởng."),
    ("Trình bày lại / phân loại lại", "Restatement / reclassification",
     "Khi 10-K năm sau trình bày lại số của năm trước theo cách phân loại mới.",
     "Bộ dữ liệu giữ số gốc và ghi chênh lệch ở cột “Trình bày lại ở 10-K sau”."),
    ("XBRL", "eXtensible Business Reporting Language",
     "Thẻ dữ liệu máy đọc được gắn trực tiếp vào số liệu trong 10-K nộp SEC.",
     "Cột “Thẻ XBRL” cho biết số liệu được lấy từ thẻ nào; dùng để kiểm tra tự động."),
]

GAPS = [
    ("Lợi nhuận và chi phí theo kênh", "Không công bố",
     "NIKE không công bố lợi nhuận, biên gộp hay chi phí riêng của NIKE Direct và Wholesale; SG&A là chi phí hợp nhất.",
     "Không tính và không ước lượng. Mọi phân tích lợi nhuận chỉ ở phạm vi NIKE, Inc."),
    ("Tỷ trọng NIKE Direct FY2026 do NIKE nêu", "Thiếu",
     "10-K FY2026 không có câu “approximately …% of total NIKE Brand revenues”.",
     "Để trống, trạng thái “Thiếu / không công bố”. Tỷ trọng tự tính ở sheet 04 vẫn có đủ 7 năm."),
    ("Doanh số Digital", "Độ chính xác thấp",
     "Chỉ công bố trong MD&A, làm tròn 0,1 tỷ USD; FY2023 được trình bày lại từ 12,6 xuống 12,4 tỷ USD do phân loại lại giữa cửa hàng và Digital.",
     "Giữ số gốc, ghi chú trình bày lại; tỷ trọng Digital/Direct ghi rõ là xấp xỉ."),
    ("Tăng trưởng loại trừ tỷ giá", "Không tự tính lại được",
     "Do NIKE tính bằng tỷ giá nội bộ; không có dữ liệu để tính lại.",
     "Lấy nguyên văn từ bảng doanh thu trong MD&A, rồi đọc lại trên trang 10-K để đối chiếu (ảnh chụp trong handover/bang-chung-doi-chieu/)."),
]

RESTATEMENT_LOG = [
    ("10-K FY2024", "NIKE Direct: cửa hàng và Digital",
     "“Within NIKE Direct revenues, there were certain reclassifications made between NIKE-owned retail stores and NIKE Brand Digital in the prior period to conform to current period presentation.”",
     "Doanh số Digital FY2023: 12,6 tỷ USD (10-K FY2023) được trình bày lại thành 12,4 tỷ USD. Tổng NIKE Direct không đổi."),
    ("10-K FY2022, FY2025, FY2026", "Các lần phân loại lại khác",
     "FY2022: doanh thu chia theo nhóm người tiêu dùng Nam, Nữ, Trẻ em (hàng unisex xếp vào Nam) và báo cáo riêng Jordan Brand. "
     "FY2025: bỏ chỉ tiêu “wholesale equivalent revenues”. FY2026: phân loại lại một số khoản năm trước (Note 1).",
     "Không làm thay đổi chỉ tiêu nào trong bộ dữ liệu."),
]


def write_sources_sheet(wb, st, obs, results, extracted_on):
    ws = wb.add_worksheet(SHEET_SRC)
    ws.hide_gridlines(2)
    for c, w in enumerate([26, 30, 70, 50, 14, 22, 26, 40]):
        ws.set_column(c, c, w)
    ws.write(0, 0, "Nguồn, định nghĩa và giới hạn của dữ liệu", st(bold=True, font_size=14))
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
    table("A. Tài liệu nguồn chính thức (Form 10-K nộp SEC)",
          ["Năm tài chính và kỳ báo cáo", "Tài liệu", "Liên kết SEC EDGAR", "Bản sao trên NIKE Investor Relations",
           "Ngày nộp SEC", "Số accession"], src_rows, link_cols=(2, 3))
    table("B. Thuật ngữ và định nghĩa", ["Thuật ngữ", "Tiếng Anh / tên trong 10-K", "Định nghĩa", "Lưu ý phạm vi và cách dùng"],
          [list(g) for g in GLOSSARY])
    table("C. Khoảng trống dữ liệu và giới hạn", ["Vấn đề", "Tình trạng", "Mô tả", "Cách xử lý trong bộ dữ liệu"],
          [list(g) for g in GAPS])
    table("D. Các lần trình bày lại / phân loại lại NIKE công bố", ["Tài liệu", "Nội dung", "Trích dẫn từ 10-K", "Ảnh hưởng tới bộ dữ liệu"],
          [list(g) for g in RESTATEMENT_LOG])
    counts = {s: sum(1 for o in obs if o["status"] == s) for s in STATUSES}
    meaning = {
        STATUS_MATCHED: ("Số liệu lấy tự động từ 10-K và được kiểm tra chéo bằng ít nhất một cách thứ hai: dữ liệu XBRL do SEC công bố "
                         "(xác nhận đọc đúng thẻ và đúng kỳ), bảng trong MD&A, phép tính lại từ BCTC, hoặc đọc lại trực tiếp trên "
                         "trang 10-K (có ảnh chụp đối chiếu).",),
        STATUS_PENDING: ("Số liệu chính thức nhưng chưa được kiểm tra chéo bằng cách thứ hai. Cần người mở 10-K đối chiếu.",),
        STATUS_HUMAN: ("Thành viên nhóm đã tự mở 10-K, tìm đúng trang và xác nhận số liệu; ghi tên và ngày.",),
        STATUS_MISSING: ("NIKE không công bố số liệu này trong 10-K tương ứng. Ô giá trị để trống, không thay bằng 0.",),
        STATUS_CONFLICT: ("Hai nguồn chính thức cho kết quả khác nhau, lần đọc lại trên trang 10-K ra số khác, hoặc số đã xác minh khác số "
                          "trích xuất mới. Không dùng cho đến khi giải quyết.",),
    }
    table("E. Trạng thái dữ liệu", ["Trạng thái", "Tiếng Anh", "Ý nghĩa", "Số quan sát hiện tại"],
          [[s, en, meaning[s][0], counts[s]] for s, en in STATUSES.items()])
    checks_ok = all(all(results[y][c["id"]] is None or abs(results[y][c["id"]]) <= Decimal(str(c["tol"])) for y in dataset.YEARS)
                    for c in CHECKS + COMPARISONS)
    sec_rows = list(dataset.load_sec_check().values())
    pages = list(dataset.load_page_check().values())
    page_ok = sum(r["result"] == "match" for r in pages)
    page_on = max((r.get("checked_on") or extracted_on for r in pages), default=extracted_on)
    table("F. Kiểm tra tự động (lần chạy gần nhất)", ["Kiểm tra", "Kết quả", "Chi tiết", "Ngày kiểm tra"], [
        ["Đối chiếu với API XBRL của SEC", f"{sec_rows.count('match')}/{len(sec_rows)} khớp",
         "Mọi giá trị hợp nhất (không chia theo kênh) của 7 báo cáo, kể cả cột so sánh, so với dữ liệu SEC theo cùng số accession.", _vn_date(extracted_on)],
        ["Thuyết minh XBRL so với bảng MD&A", "Khớp toàn bộ",
         "Doanh thu NIKE Direct, bán buôn, Global Brand Divisions, NIKE Brand, Converse, Corporate và NIKE, Inc. ở hai vị trí khác nhau trong cùng 10-K.", _vn_date(extracted_on)],
        ["Đối chiếu số học và tỷ lệ NIKE công bố", "Đạt" if checks_ok else "Có lỗi",
         "Xem mục D và E ở sheet 04.", _vn_date(extracted_on)],
        ["Đọc lại trên trang 10-K", f"{page_ok}/{len(pages)} khớp",
         "Các số chỉ có trong MD&A: mở từng 10-K, tìm trang theo chân trang; 28 tỷ lệ tăng trưởng loại trừ tỷ giá đọc lại theo vị trí "
         "cột “% CHANGE EXCLUDING CURRENCY CHANGES” đầu tiên, 7 số Digital đọc lại từ câu văn trên đúng trang. Ảnh chụp và bảng so sánh: "
         "handover/bang-chung-doi-chieu/doi-chieu-35-so.html (cũng có trên dashboard web).", _vn_date(page_on)],
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
    ws.write(1, 0, "Cơ sở số liệu cho đề tài “Đổi mới và tái cân bằng chiến lược phân phối tại Nike: Từ Direct-to-Consumer (D2C) "
                   "đến bán buôn và những biến động tài chính giai đoạn FY2020–FY2026”.", st(font_color="#444444"))
    label = st(bold=True, text_wrap=True)
    wrap = st(text_wrap=True)
    h = st(bold=True, font_size=11, bottom=1)
    n = {s: sum(1 for o in obs if o["status"] == s) for s in STATUSES}
    total = len(obs)
    blocks = [
        ("Thông tin chung", [
            ("Mục đích", "Tập hợp số liệu chính thức về doanh thu theo kênh phân phối (NIKE Direct, bán buôn) và kết quả tài chính hợp nhất "
                         "của NIKE, Inc. để nhóm dùng trong báo cáo, thuyết trình và dashboard."),
            ("Kỳ nghiên cứu", "FY2020–FY2026. Năm tài chính của NIKE kết thúc ngày 31/5, ví dụ FY2026 = 01/06/2025 – 31/05/2026. "
                              "FY không trùng với năm dương lịch."),
            ("Đơn vị", "Triệu USD, trừ khi cột “Đơn vị” ghi khác (tỷ USD cho doanh số Digital; % cho tỷ lệ). Không quy đổi sang VND."),
            ("Nguồn", "7 báo cáo thường niên Form 10-K của NIKE, Inc. nộp Ủy ban Chứng khoán Mỹ (SEC), FY2020–FY2026. "
                      "Không dùng trang tổng hợp tài chính của bên thứ ba."),
            ("Thứ tự ưu tiên nguồn", "(1) Báo cáo tài chính đã kiểm toán và thuyết minh (Item 8), đọc từ dữ liệu XBRL gắn trong chính 10-K. "
                                     "(2) Phần MD&A (Item 7) của cùng 10-K cho số liệu non-GAAP và số do ban điều hành nêu: EBIT, tăng trưởng loại trừ tỷ giá, "
                                     "doanh số Digital. (3) API XBRL của SEC chỉ dùng để đối chiếu."),
            ("Quy tắc chọn số", "Số của năm FY t lấy từ 10-K FY t. “Giá trị năm trước” là cột so sánh trong cùng 10-K đó và được dùng để tính "
                                "tăng trưởng (cùng cơ sở NIKE dùng). Nếu 10-K năm sau trình bày lại, chênh lệch được ghi ở cột "
                                "“Trình bày lại ở 10-K sau”, không thay thế số gốc."),
            ("Tình trạng dữ liệu", f"Trích xuất tự động ngày {_vn_date(extracted_on)}. {total} quan sát: " + ", ".join(
                f"{n[s]} {label}" for s, label in ((STATUS_MATCHED, "khớp nguồn chính thức"), (STATUS_HUMAN, "đã được người xác minh"),
                                                   (STATUS_PENDING, "chờ người kiểm tra"), (STATUS_MISSING, "không được công bố"),
                                                   (STATUS_CONFLICT, "mâu thuẫn")) if n[s]) + "."),
        ]),
        ("Cách sử dụng", [
            ("1. Đọc số liệu", "Sheet 02 (hợp nhất NIKE, Inc.) và 03 (kênh phân phối NIKE Brand). Mỗi dòng là một số liệu của một năm. "
                               "Dùng nút lọc ở dòng tiêu đề để chọn chỉ tiêu hoặc năm."),
            ("2. Kiểm tra nguồn", "Bấm liên kết ở cột “Tài liệu nguồn” để mở 10-K trên SEC EDGAR. Tìm số trang ở cột “Trang” (số in ở chân trang, "
                                  "dạng “2026 FORM 10-K 83”) rồi đối chiếu với cột “Số hoặc câu hiển thị trong tài liệu”."),
            ("3. Ghi nhận xác minh", "Sau khi tự đối chiếu, chọn “Đã xác minh (người)” ở cột “Trạng thái dữ liệu” và điền tên, ngày ở các cột nền vàng."),
            ("4. Xem chỉ số", "Sheet 04 tính các tỷ lệ bằng công thức Excel từ sheet 02–03. Cột “Công thức” diễn giải từng chỉ số; mục D–E là kiểm tra đối chiếu."),
            ("5. Tra cứu", "Sheet 05: danh mục 10-K, thuật ngữ, khoảng trống dữ liệu và các lần NIKE trình bày lại số liệu."),
        ]),
        ("Quy ước trình bày", [
            ("Chữ xanh dương", "Số liệu nhập từ tài liệu nguồn (không phải công thức)."),
            ("Chữ xanh lá", "Công thức lấy số liệu từ sheet khác."),
            ("Chữ đen", "Công thức tính toán."),
            ("Số âm", "Trong ngoặc đơn, ví dụ (50), theo cách trình bày của NIKE. Tỷ lệ % dùng dấu trừ."),
            ("Ô trống", "Kèm trạng thái “Thiếu / không công bố”: NIKE không công bố số liệu đó. Không thay bằng 0. Ở sheet 04, “–” là không tính được do thiếu đầu vào."),
            ("Tỷ lệ %", "Lưu dạng thập phân (0,14 hiển thị 14%) để dùng được trong công thức."),
            ("Cột tiêu đề nền vàng", "Dành cho nhóm điền: người xác minh, ngày xác minh, ghi chú."),
        ]),
        ("Lưu ý khi dùng cho nghiên cứu", [
            ("Không suy ra nhân quả", "Các chỉ số cho thấy số liệu thay đổi cùng giai đoạn; không chứng minh thay đổi chiến lược kênh là nguyên nhân."),
            ("Không có lợi nhuận theo kênh", "NIKE không công bố lợi nhuận hay chi phí riêng của NIKE Direct và bán buôn; SG&A là chi phí hợp nhất."),
            ("Không trộn phạm vi", "Tỷ trọng kênh dùng mẫu số NIKE Brand; biên lợi nhuận dùng doanh thu NIKE, Inc."),
            ("Trích dẫn", "Khi dùng số liệu, trích dẫn 10-K gốc (năm, trang) thay vì trích dẫn file này."),
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
                       "comments": "Nguồn: NIKE, Inc. Form 10-K FY2020–FY2026 (SEC EDGAR)."})
    st = Styles(wb)
    write_overview_sheet(wb, st, obs, extracted_on)
    fin = [o for o in obs if o["metric"]["sheet"] == "financial"]
    ch = [o for o in obs if o["metric"]["sheet"] == "channels"]
    write_observation_sheet(wb, st, SHEET_FIN, "Dữ liệu tài chính hợp nhất – NIKE, Inc.",
                            "Báo cáo kết quả kinh doanh, hàng tồn kho, EBIT và tăng trưởng doanh thu NIKE công bố. Mỗi dòng: một chỉ tiêu × một năm tài chính.", fin)
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
