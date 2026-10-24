"""Build the Excel version of the dashboard, for viewing in Excel for the web (OneDrive).

Usage:  python -m pipeline.build_excel_dashboard

Reads site/data/nike-financials.json (the validated data layer of the web dashboard) and
the static texts of site/index.html, and writes handover/NIKE_Dashboard_FY2020-2026.xlsx
with the same two views, charts, summaries, tables and sources. Every chart and summary
reads the year table on the last sheet; a slicer on that table replaces the web page's
year filter, and a second slicer replaces the growth-basis switch. XlsxWriter cannot
create slicers, so on Windows the script then runs pipeline/excel_slicers.ps1, which adds
them with the installed Excel and saves the workbook (this also caches formula results).
"""
import json
import math
import re
import subprocess
import sys
from pathlib import Path

import xlsxwriter
from lxml import html

from .build_workbook import DASHBOARD_URL
from .dataset import vi_num

ROOT = Path(__file__).resolve().parent.parent
SITE_DATA = ROOT / "site" / "data" / "nike-financials.json"
INDEX_HTML = ROOT / "site" / "index.html"
APP_JS = ROOT / "site" / "assets" / "app.js"
EVIDENCE_PAGE = "bang-chung-doi-chieu/doi-chieu-35-so.html"
OUT = ROOT / "handover" / "NIKE_Dashboard_FY2020-2026.xlsx"
SLICER_SCRIPT = Path(__file__).resolve().parent / "excel_slicers.ps1"

SHEET_CH = "Chiến lược phân phối"
SHEET_PERF = "Kết quả tài chính"
SHEET_SRC = "Nguồn và phương pháp"
SHEET_DATA = "Dữ liệu biểu đồ"
STATUS_PENDING = "Chờ người kiểm tra"

# Same palette as site/assets/app.js
C = dict(direct="#2a78d6", wholesale="#eb6834", gbd="#a3a9b2", direct_light="#86b6ef", digital="#1c5cab",
         total="#475467", overhead="#6250d6", demand="#1baf7a", ink="#15191e", ink2="#3b4350", muted="#5b6472",
         grid="#e6e9ee", axis="#c3c9d1", rule="#d9dde3", band="#f4f5f7", link="#1f4e9a")
FONT, SERIF = "Arial", "Cambria"
MINUS = "−"

# Page grid: column A is a margin, B..M are twelve 100-pixel columns (1,200 px of content).
GRID_W = 13.57
COL_PX = 100
ROW_PX = 20
# Printing: a fixed scale, because Excel ignores manual page breaks under fit-to-page scaling. At 79% the grid
# (about 936 pt with the margin columns) fits A4 and Letter landscape; the breaks keep chart rows and table groups whole.
PRINT_SCALE = 79
PAGE_ROWS_PT = (595.3 - 2 * 0.4 * 72) * 100 / PRINT_SCALE   # row height that fits on an A4-landscape page

def app_texts():
    """Chart source lines and the growth subtitle, read from app.js so the versions cannot drift apart."""
    js = APP_JS.read_text(encoding="utf-8")
    sources = dict(re.findall(r"sourceLine\('(\w+)', '([^']*)'", js))
    m = re.search(r"sourceLine\('growth', cn \? '([^']*)'\s*:\s*'([^']*)'", js)
    sources["growth_cn"], sources["growth_reported"] = m.group(1), m.group(2)
    m = re.search(r"getElementById\('growth-sub'\)\.textContent = cn\s*\?\s*'([^']*)'\s*:\s*'([^']*)'", js)
    return sources, {"cn": m.group(1), "reported": m.group(2)}


# ------------------------------------------------------------------ number formatting (vi-VN, as in app.js)
def _vi(v, d):
    return vi_num(abs(v), d)


def num(v, d=0):
    return "—" if v is None else (MINUS if v < 0 else "") + _vi(v, d)


def amount(v):
    return "—" if v is None else (f"({_vi(v, 0)})" if v < 0 else _vi(v, 0))


def pct(v, d=1):
    return "—" if v is None else num(v * 100, d) + "%"


def signed_pct(v, d=1):
    if v is None:
        return "—"
    return ("+" if v > 0 else MINUS if v < 0 else "") + _vi(v * 100, d) + "%"


def bn(millions, d=1):
    return "—" if millions is None else num(millions / 1000, d)


# Table rows, as TABLES in app.js: (kind, id or group title, label, unit, formatter, total row)
TABLES = {
    SHEET_CH: [
        ("g", "Doanh thu NIKE Brand"),
        ("obs", "nike_direct_revenue", "NIKE Direct", "triệu USD", amount, False),
        ("obs", "nike_brand_wholesale_revenue", "Bán buôn (Wholesale)", "triệu USD", amount, False),
        ("obs", "global_brand_divisions_revenue", "Global Brand Divisions", "triệu USD", amount, False),
        ("obs", "nike_brand_revenue", "Tổng doanh thu NIKE Brand", "triệu USD", amount, True),
        ("g", "Cơ cấu theo kênh"),
        ("kpi", "direct_share", "Tỷ trọng NIKE Direct", "%", pct, False),
        ("kpi", "wholesale_share", "Tỷ trọng bán buôn", "%", pct, False),
        ("obs", "nike_direct_share_stated", "Tỷ trọng NIKE Direct do NIKE nêu (xấp xỉ)", "%", lambda v: pct(v, 0), False),
        ("g", "Tăng trưởng so với năm trước"),
        ("kpi", "direct_growth", "NIKE Direct, theo báo cáo", "%", signed_pct, False),
        ("obs", "nike_direct_pct_change_cn", "NIKE Direct, loại trừ tỷ giá (non-GAAP)", "%", lambda v: signed_pct(v, 0), False),
        ("kpi", "wholesale_growth", "Bán buôn, theo báo cáo", "%", signed_pct, False),
        ("obs", "nike_brand_wholesale_pct_change_cn", "Bán buôn, loại trừ tỷ giá (non-GAAP)", "%", lambda v: signed_pct(v, 0), False),
        ("kpi", "nike_brand_growth", "NIKE Brand, theo báo cáo", "%", signed_pct, False),
        ("g", "NIKE Brand Digital (một phần của NIKE Direct)"),
        ("obs", "nike_brand_digital_sales", "Doanh số Digital", "tỷ USD", lambda v: num(v, 1), False),
        ("kpi", "digital_share_of_direct", "Digital trong NIKE Direct (xấp xỉ)", "%", lambda v: pct(v, 0), False),
    ],
    SHEET_PERF: [
        ("g", "Báo cáo kết quả kinh doanh"),
        ("obs", "revenues", "Doanh thu", "triệu USD", amount, True),
        ("obs", "cost_of_sales", "Giá vốn hàng bán", "triệu USD", amount, False),
        ("obs", "gross_profit", "Lợi nhuận gộp", "triệu USD", amount, True),
        ("obs", "demand_creation_expense", "Chi phí tạo nhu cầu", "triệu USD", amount, False),
        ("obs", "operating_overhead_expense", "Chi phí vận hành chung", "triệu USD", amount, False),
        ("obs", "total_selling_admin_expense", "Tổng chi phí bán hàng và quản lý", "triệu USD", amount, True),
        ("obs", "interest_expense_net", "Chi phí (thu nhập) lãi thuần", "triệu USD", amount, False),
        ("obs", "other_expense_net", "Chi phí (thu nhập) khác thuần", "triệu USD", amount, False),
        ("obs", "income_before_taxes", "Lợi nhuận trước thuế", "triệu USD", amount, True),
        ("obs", "income_tax_expense", "Chi phí thuế thu nhập", "triệu USD", amount, False),
        ("obs", "net_income", "Lợi nhuận sau thuế", "triệu USD", amount, True),
        ("g", "Tỷ lệ trên doanh thu và tăng trưởng"),
        ("kpi", "revenue_growth", "Tăng trưởng doanh thu, theo báo cáo", "%", signed_pct, False),
        ("obs", "revenues_pct_change_cn", "Tăng trưởng doanh thu, loại trừ tỷ giá (non-GAAP)", "%", lambda v: signed_pct(v, 0), False),
        ("kpi", "gross_margin", "Biên lợi nhuận gộp", "%", pct, False),
        ("kpi", "sga_ratio", "Chi phí bán hàng và quản lý / doanh thu", "%", pct, False),
        ("kpi", "demand_creation_ratio", "Trong đó: tạo nhu cầu / doanh thu", "%", pct, False),
        ("kpi", "operating_overhead_ratio", "Trong đó: vận hành chung / doanh thu", "%", pct, False),
        ("kpi", "net_income_margin", "Biên lợi nhuận sau thuế", "%", pct, False),
        ("g", "EBIT (non-GAAP, định nghĩa của NIKE)"),
        ("obs", "ebit_reported", "EBIT do NIKE công bố", "triệu USD", amount, False),
        ("kpi", "ebit_margin", "Biên EBIT", "%", pct, False),
        ("g", "Hàng tồn kho"),
        ("obs", "inventories", "Hàng tồn kho tại 31/5", "triệu USD", amount, False),
        ("kpi", "inventory_growth", "Thay đổi so với đầu kỳ", "%", signed_pct, False),
        ("kpi", "inventory_turnover", "Vòng quay hàng tồn kho", "lần", lambda v: num(v, 2), False),
        ("kpi", "days_inventory", "Số ngày tồn kho bình quân", "ngày", lambda v: num(v, 1), False),
    ],
}


# ------------------------------------------------------------------ texts from the web page
def page_texts():
    root = html.parse(str(INDEX_HTML)).getroot()

    def t(xpath):
        return " ".join(root.xpath(xpath)[0].text_content().split())

    def notes(labelledby):
        aside = root.xpath(f"//aside[@aria-labelledby='{labelledby}']")[0]
        pairs = zip(aside.xpath(".//dt"), aside.xpath(".//dd"))
        return t(f"//h3[@id='{labelledby}']"), [(" ".join(a.text_content().split()), " ".join(b.text_content().split())) for a, b in pairs]

    def fig(fid):
        f = root.xpath(f"//figure[@id='{fid}']")[0]
        sub = f.xpath(".//p[@class='figure-sub']")
        return " ".join(f.xpath(".//h3")[0].text_content().split()), " ".join(sub[0].text_content().split()) if sub else ""

    return dict(
        title=t("//header[@class='masthead']//h1"), subtitle=t("//p[@class='masthead-sub']"),
        q_ch=t("//section[@id='view-channels']//h2"), q_perf=t("//section[@id='view-performance']//h2"),
        summary_label=t("//section[@id='view-channels']//p[@class='summary-label']"),
        notes_ch=notes("notes-channels"), notes_perf=notes("notes-performance"),
        figs={k: fig(v) for k, v in dict(share="fig-share", channelRevenue="fig-channel-revenue", growth="fig-growth",
                                         digital="fig-digital", revenue="fig-revenue", grossMargin="fig-gross-margin",
                                         sga="fig-sga", ebit="fig-ebit", inventory="fig-inventory", days="fig-days").items()},
        table_ch=t("//h3[@id='table-channels-title']"), table_perf=t("//h3[@id='table-performance-title']"),
        method=[" ".join(li.text_content().split()) for li in root.xpath("//section[@aria-labelledby='footer-method']//li")],
        fine=t("//p[@class='fine']"),
    )


# ------------------------------------------------------------------ workbook
class Book:
    def __init__(self, path, data):
        self.d = data
        self.years = data["years"]
        # Arial 10 rows are 17 px by default (XlsxWriter assumes Calibri's 20 px and would then skip writing the
        # 20 px rows the layout is designed in, leaving Excel to draw them at 17 px).
        self.wb = xlsxwriter.Workbook(str(path), {"use_future_functions": True, "default_row_height": 17})
        self.wb.set_properties({"title": "NIKE FY2020–FY2026: phân phối và kết quả tài chính",
                                "comments": "Phiên bản Excel của dashboard. Nguồn: NIKE, Inc. Form 10-K FY2020–FY2026 (SEC EDGAR)."})
        self.wb.formats[0].set_font_name(FONT)
        self.wb.formats[0].set_font_size(10)
        self.cache = {}
        self.texts = page_texts()
        self.sources, self.growth_sub = app_texts()

    def f(self, **p):
        base = {"font_name": FONT, "font_size": 10, "font_color": C["ink2"], "valign": "top"}
        key = tuple(sorted(p.items()))
        if key not in self.cache:
            self.cache[key] = self.wb.add_format({**base, **p})
        return self.cache[key]

    # data access
    def ob(self, mid, y):
        return self.d["observations"][mid].get(str(y)) or {}

    def val(self, mid, y):
        return self.ob(mid, y).get("value")

    def kpi(self, kid, y):
        return self.d["kpis"][kid]["values"][str(y)]

    def pending_count(self):
        return self.d["meta"]["status_counts"].get(STATUS_PENDING, 0)


# Columns of the year table on the data sheet: (header, value for a year)
def year_columns(b):
    return [
        ("Năm tài chính", lambda y: f"FY{y}"),
        ("Năm", lambda y: y),
        ("Tỷ trọng NIKE Direct", lambda y: b.kpi("direct_share", y)),
        ("Tỷ trọng bán buôn", lambda y: b.kpi("wholesale_share", y)),
        ("Tỷ trọng Global Brand Divisions", lambda y: b.kpi("gbd_share", y)),
        ("NIKE Direct (tỷ USD)", lambda y: b.val("nike_direct_revenue", y) / 1000),
        ("Bán buôn (tỷ USD)", lambda y: b.val("nike_brand_wholesale_revenue", y) / 1000),
        ("Tăng trưởng NIKE Direct theo báo cáo", lambda y: b.kpi("direct_growth", y)),
        ("Tăng trưởng bán buôn theo báo cáo", lambda y: b.kpi("wholesale_growth", y)),
        ("Tăng trưởng NIKE Direct loại trừ tỷ giá", lambda y: b.val("nike_direct_pct_change_cn", y)),
        ("Tăng trưởng bán buôn loại trừ tỷ giá", lambda y: b.val("nike_brand_wholesale_pct_change_cn", y)),
        ("NIKE Direct (biểu đồ tăng trưởng)", None),   # formula: follows the growth-basis slicer
        ("Bán buôn (biểu đồ tăng trưởng)", None),
        ("NIKE Brand Digital (tỷ USD)", lambda y: b.val("nike_brand_digital_sales", y)),
        ("Doanh thu NIKE, Inc. (tỷ USD)", lambda y: b.val("revenues", y) / 1000),
        ("Biên lợi nhuận gộp", lambda y: b.kpi("gross_margin", y)),
        ("Vận hành chung / doanh thu", lambda y: b.kpi("operating_overhead_ratio", y)),
        ("Tạo nhu cầu / doanh thu", lambda y: b.kpi("demand_creation_ratio", y)),
        ("Tổng SG&A / doanh thu", lambda y: b.kpi("sga_ratio", y)),
        ("Biên EBIT", lambda y: b.kpi("ebit_margin", y)),
        ("Hàng tồn kho (tỷ USD)", lambda y: b.val("inventories", y) / 1000),
        ("Số ngày tồn kho bình quân", lambda y: b.kpi("days_inventory", y)),
        # Vietnamese-formatted text used by the summaries (same rounding as the web page)
        ("Tỷ trọng NIKE Direct (chữ)", lambda y: pct(b.kpi("direct_share", y))),
        ("Tỷ trọng bán buôn (chữ)", lambda y: pct(b.kpi("wholesale_share", y))),
        ("NIKE Direct (chữ)", lambda y: bn(b.val("nike_direct_revenue", y))),
        ("Bán buôn (chữ)", lambda y: bn(b.val("nike_brand_wholesale_revenue", y))),
        ("Doanh thu (chữ)", lambda y: bn(b.val("revenues", y))),
        ("Biên lợi nhuận gộp (chữ)", lambda y: pct(b.kpi("gross_margin", y))),
        ("Biên EBIT (chữ)", lambda y: pct(b.kpi("ebit_margin", y))),
        ("Hàng tồn kho (chữ)", lambda y: bn(b.val("inventories", y))),
    ]


class DataSheet:
    """Year table (filtered by the slicer), basis table, summary helpers and table texts."""
    TOP = 3            # header row of the year table (0-based)
    BASIS_TOP = 13     # header row of the basis table; must not share rows with the year table
    HELP_TOP = 18
    TEXT_TOP = 44

    def __init__(self, b, ws):
        self.b = b
        self.ws = ws
        self.cols = year_columns(b)
        self.col = {h: i for i, (h, _) in enumerate(self.cols)}
        self.first = self.TOP + 1
        self.last = self.TOP + len(b.years)

    def ref(self, header, absolute=True):
        c = self.col[header]
        return xlsxwriter.utility.xl_range_abs(self.first, c, self.last, c)

    def q(self, rng):
        return f"'{SHEET_DATA}'!{rng}"

    def series(self, header):
        c = self.col[header]
        return [SHEET_DATA, self.first, c, self.last, c]
    def write(self):
        b, ws = self.b, self.ws
        ws.write(0, 0, "Dữ liệu cho biểu đồ và tóm tắt", b.f(bold=True, font_size=14, font_color=C["ink"]))
        ws.write(1, 0, "Lấy từ site/data/nike-financials.json (cùng dữ liệu với dashboard web). Slicer “Năm tài chính” lọc bảng tblNam; "
                       "biểu đồ và tóm tắt chỉ dùng các năm đang hiện. Không sửa sheet này.", b.f(font_color=C["muted"]))
        pct_cols = {h for h, _ in self.cols if h.startswith(("Tỷ trọng", "Tăng trưởng", "Biên", "Vận hành", "Tạo nhu cầu", "Tổng SG&A"))
                    or h.endswith("(biểu đồ tăng trưởng)")}
        for i, (h, _) in enumerate(self.cols):
            ws.set_column(i, i, 16 if i > 1 else 13)
        ws.add_table(self.TOP, 0, self.last, len(self.cols) - 1, {
            "name": "tblNam", "style": "Table Style Light 1", "autofilter": True,
            "columns": [{"header": h} for h, _ in self.cols]})
        basis = xlsxwriter.utility.xl_rowcol_to_cell(self.HELP_TOP + 1, 2, row_abs=True, col_abs=True)
        for r, y in enumerate(b.years, start=self.first):
            for c, (h, fn) in enumerate(self.cols):
                if fn is None:
                    which = "NIKE Direct" if h.startswith("NIKE Direct") else "bán buôn"
                    rep = xlsxwriter.utility.xl_rowcol_to_cell(r, self.col[f"Tăng trưởng {which} theo báo cáo"])
                    cn = xlsxwriter.utility.xl_rowcol_to_cell(r, self.col[f"Tăng trưởng {which} loại trừ tỷ giá"])
                    ws.write_formula(r, c, f'=IF({basis}="cn",{cn},{rep})', b.f(num_format="0.0%"))
                else:
                    v = fn(y)
                    fmt = b.f(num_format="0.0%") if h in pct_cols else b.f(num_format="0.000") if isinstance(v, float) else b.f()
                    ws.write(r, c, v, fmt)

        # Growth basis: a two-row table filtered by its own slicer.
        bt = self.BASIS_TOP
        ws.add_table(bt, 0, bt + 2, 0, {"name": "tblCoSo", "style": "Table Style Light 1",
                                         "columns": [{"header": "Cơ sở tính tăng trưởng"}]})
        ws.write(bt + 1, 0, "Theo báo cáo")
        ws.write(bt + 2, 0, "Loại trừ tỷ giá")
        ws.set_column(0, 0, 24)

        # Summary helpers (rows below both tables, so slicer filtering never hides them).
        yr = self.ref("Năm")
        h = self.HELP_TOP
        ws.write(h, 0, "Biến phụ cho tóm tắt", b.f(bold=True))
        helpers = [
            ("Cơ sở tăng trưởng", f'=IF(AND(SUBTOTAL(103,A{bt + 3})=1,SUBTOTAL(103,A{bt + 2})=0),"cn","reported")'),
            ("Năm đầu", f"=AGGREGATE(15,5,{yr},1)"),
            ("Năm cuối", f"=AGGREGATE(14,5,{yr},1)"),
            ("Số năm", f"=SUBTOTAL(103,{yr})"),
            ("Năm tỷ trọng Direct cao nhất",
             f"=INDEX({yr},MATCH(AGGREGATE(14,5,{self.ref('Tỷ trọng NIKE Direct')},1),{self.ref('Tỷ trọng NIKE Direct')},0))"),
            ("Năm tỷ trọng Direct thấp nhất",
             f"=INDEX({yr},MATCH(AGGREGATE(15,5,{self.ref('Tỷ trọng NIKE Direct')},1),{self.ref('Tỷ trọng NIKE Direct')},0))"),
            ("Năm biên EBIT cao nhất", f"=INDEX({yr},MATCH(AGGREGATE(14,5,{self.ref('Biên EBIT')},1),{self.ref('Biên EBIT')},0))"),
        ]
        self.h = {}
        for i, (label, formula) in enumerate(helpers):
            r = h + 1 + i
            ws.write(r, 1, label)
            ws.write_formula(r, 2, formula)
            self.h[label] = xlsxwriter.utility.xl_rowcol_to_cell(r, 2, row_abs=True, col_abs=True)
        assert self.h["Cơ sở tăng trưởng"] == basis
        row = h + 1 + len(helpers)
        self.summary_cells = {}
        for name, formula in (("channels", self.channels_summary()), ("performance", self.performance_summary()),
                              ("growth_sub", f'=IF({self.qh("Cơ sở tăng trưởng")}="cn","{b.growth_sub["cn"]}","{b.growth_sub["reported"]}")'),
                              ("growth_source", f'=IF({self.qh("Cơ sở tăng trưởng")}="cn","{b.sources["growth_cn"]}","{b.sources["growth_reported"]}")'),
                              ("reports", f'=" Báo cáo: 10-K FY"&{self.qh("Năm đầu")}&IF({self.qh("Năm đầu")}={self.qh("Năm cuối")},"","–FY"&{self.qh("Năm cuối")})'
                                          f'&" (danh sách ở sheet “{SHEET_SRC}”)."')):
            ws.write(row, 1, name)
            ws.write_formula(row, 2, formula)
            self.summary_cells[name] = self.q(xlsxwriter.utility.xl_rowcol_to_cell(row, 2, row_abs=True, col_abs=True))
            row += 1
        self.write_table_texts()

    def qh(self, label):
        return self.q(self.h[label])

    def text_at(self, header, year_cell):
        """Text of a year-table column for the year held in `year_cell`."""
        return f"INDEX({self.ref(header)},{year_cell}-{self.b.years[0] - 1})"

    def channels_summary(self):
        a, b_, hi, lo, n = self.h["Năm đầu"], self.h["Năm cuối"], self.h["Năm tỷ trọng Direct cao nhất"], self.h["Năm tỷ trọng Direct thấp nhất"], self.h["Số năm"]
        T = lambda col, y: self.text_at(col, y)
        one = (f'"Trong FY"&{a}&", NIKE Direct chiếm "&{T("Tỷ trọng NIKE Direct (chữ)", a)}&" doanh thu NIKE Brand ("&'
               f'{T("NIKE Direct (chữ)", a)}&" tỷ USD); bán buôn chiếm "&{T("Tỷ trọng bán buôn (chữ)", a)}&" ("&{T("Bán buôn (chữ)", a)}&" tỷ USD)."')
        many = (f'"Tỷ trọng NIKE Direct trong doanh thu NIKE Brand là "&{T("Tỷ trọng NIKE Direct (chữ)", a)}&" ở FY"&{a}&" và "&'
                f'{T("Tỷ trọng NIKE Direct (chữ)", b_)}&" ở FY"&{b_}&IF({n}>2,"; cao nhất "&{T("Tỷ trọng NIKE Direct (chữ)", hi)}&" (FY"&{hi}&"), thấp nhất "&'
                f'{T("Tỷ trọng NIKE Direct (chữ)", lo)}&" (FY"&{lo}&")","")&". Doanh thu NIKE Direct từ "&{T("NIKE Direct (chữ)", a)}&" đến "&'
                f'{T("NIKE Direct (chữ)", b_)}&" tỷ USD; bán buôn từ "&{T("Bán buôn (chữ)", a)}&" đến "&{T("Bán buôn (chữ)", b_)}&" tỷ USD."')
        return f"=IF({a}={b_},{one},{many})"

    def performance_summary(self):
        a, b_, ehi, n = self.h["Năm đầu"], self.h["Năm cuối"], self.h["Năm biên EBIT cao nhất"], self.h["Số năm"]
        T = lambda col, y: self.text_at(col, y)
        one = (f'"Trong FY"&{a}&", doanh thu NIKE, Inc. là "&{T("Doanh thu (chữ)", a)}&" tỷ USD, biên lợi nhuận gộp "&{T("Biên lợi nhuận gộp (chữ)", a)}&'
               f'", biên EBIT (non-GAAP) "&{T("Biên EBIT (chữ)", a)}&", hàng tồn kho cuối kỳ "&{T("Hàng tồn kho (chữ)", a)}&" tỷ USD."')
        many = (f'"Doanh thu NIKE, Inc. từ "&{T("Doanh thu (chữ)", a)}&" (FY"&{a}&") đến "&{T("Doanh thu (chữ)", b_)}&" tỷ USD (FY"&{b_}&'
                f'"). Biên lợi nhuận gộp "&{T("Biên lợi nhuận gộp (chữ)", a)}&" (FY"&{a}&") và "&{T("Biên lợi nhuận gộp (chữ)", b_)}&" (FY"&{b_}&'
                f'"); biên EBIT (non-GAAP) "&{T("Biên EBIT (chữ)", a)}&" (FY"&{a}&") và "&{T("Biên EBIT (chữ)", b_)}&" (FY"&{b_}&")"&'
                f'IF({n}>2,", cao nhất "&{T("Biên EBIT (chữ)", ehi)}&" (FY"&{ehi}&")","")&'
                f'". Hàng tồn kho cuối kỳ từ "&{T("Hàng tồn kho (chữ)", a)}&" đến "&{T("Hàng tồn kho (chữ)", b_)}&" tỷ USD."')
        return f"=IF({a}={b_},{one},{many})"

    def write_table_texts(self):
        """Two rows per dashboard table row: the figure for each year (with the 10-K links), then its page reference.

        Figures and page references live in separate rows so every figure in a year column ends at the same
        place, under the year header (a page reference inside the cell would push reported figures left).
        """
        b, ws = self.b, self.ws
        ny = len(b.years)
        r = self.TEXT_TOP
        ws.write(r, 0, "Văn bản cho bảng số liệu (định dạng như dashboard web)", b.f(bold=True))
        ws.write_row(r + 1, 0, ["Bảng", "Mã"] + [f"FY{y}" for y in b.years] + [f"Liên kết FY{y}" for y in b.years])
        r += 2
        self.text_rows = {}
        for sheet, rows in TABLES.items():
            for spec in rows:
                if spec[0] == "g":
                    continue
                kind, mid, _label, _unit, fmt, _total = spec
                texts, pages, links = [], [], []
                for y in b.years:
                    if kind == "kpi":
                        v = b.kpi(mid, y)
                        texts.append("—" if v is None else fmt(v))
                        pages.append("")
                        links.append("")
                        continue
                    o = b.ob(mid, y)
                    if o.get("value") is None:
                        texts.append("Không công bố")
                        pages.append("")
                        links.append("")
                        continue
                    mark = "† " if o.get("status") == STATUS_PENDING else ""
                    page = o.get("page")
                    texts.append(fmt(o["value"]))
                    pages.append(mark + (f"tr. {page}" if page else ""))
                    links.append(b.d["filings"][str(y)]["url"] if page else "")
                ws.write_row(r, 0, [sheet, mid] + texts + links)
                ws.write_row(r + 1, 0, [sheet, f"{mid} (trang)"] + pages)
                self.text_rows[(sheet, mid)] = (xlsxwriter.utility.xl_range_abs(r, 2, r, 1 + ny),
                                                xlsxwriter.utility.xl_range_abs(r, 2 + ny, r, 1 + 2 * ny),
                                                xlsxwriter.utility.xl_range_abs(r + 1, 2, r + 1, 1 + ny))
                r += 2


# ------------------------------------------------------------------ dashboard sheets
class Dashboard:
    YEAR_SLICER = (600, 54)    # width, height in points (B..I, rows 8-10)
    BASIS_SLICER = (450, 54)

    def __init__(self, b, data_sheet, ws, name, view):
        self.b, self.data, self.ws, self.name, self.view = b, data_sheet, ws, name, view
        ws.hide_gridlines(2)
        ws.set_column(0, 0, 2.71)
        ws.set_column(1, 12, GRID_W)
        ws.set_column(13, 13, 2.71)
        ws.set_default_row(15)
        ws.set_landscape()
        ws.set_paper(9)
        ws.set_print_scale(PRINT_SCALE)
        ws.set_margins(0.3, 0.3, 0.4, 0.4)

    def paginate(self, starts, end):
        """Manual page breaks between blocks (chart rows, table groups) so that none is split across pages."""
        size = lambda r: self.ws.row_sizes.get(r, [self.ws.default_row_height])[0] * 0.75   # XlsxWriter keeps pixels
        breaks, used = [], 0
        for a, z in zip(starts, starts[1:] + [end]):
            h = sum(size(r) for r in range(a, z))
            if used and used + h > PAGE_ROWS_PT:
                breaks.append(a)
                used = 0
            used += h
        self.ws.set_h_pagebreaks(breaks)

    # helpers
    def text(self, r, c1, c2, value, fmt, height=None):
        if height:
            self.ws.set_row(r, height)
        if c1 == c2:
            self.ws.write(r, c1, value, fmt)
        else:
            self.ws.merge_range(r, c1, r, c2, value, fmt)

    def formula(self, r1, c1, r2, c2, formula, fmt):
        if (r1, c1) == (r2, c2):
            self.ws.write_formula(r1, c1, formula, fmt)
        else:
            self.ws.merge_range(r1, c1, r2, c2, "", fmt)
            self.ws.write_formula(r1, c1, formula, fmt)

    def header(self):
        b, ws, t = self.b, self.ws, self.b.texts
        ws.set_row(0, 9)
        self.text(1, 1, 8, t["title"], b.f(font_name=SERIF, font_size=22, bold=True, font_color=C["ink"]), 36)
        ws.merge_range(2, 1, 3, 8, t["subtitle"], b.f(font_size=11, font_color=C["ink2"], text_wrap=True))
        link = b.f(font_color=C["link"], underline=1, align="right")
        ws.write_url(1, 12, DASHBOARD_URL + b.d["meta"]["workbook"], link, "Tải bộ dữ liệu Excel")
        ws.write_url(2, 12, f"internal:'{SHEET_SRC}'!A1", link, "Nguồn và phương pháp")
        ws.write_url(3, 12, DASHBOARD_URL, link, "Mở dashboard web")
        ws.set_row(4, 9)
        # Tabs
        band = b.f(bg_color=C["band"], top=1, bottom=1, top_color=C["rule"], bottom_color=C["rule"], valign="vcenter")
        ws.set_row(5, 27)
        for c in range(1, 13):
            ws.write_blank(5, c, None, band)
        active = b.f(bold=True, font_size=11, font_color=C["ink"], bg_color=C["band"], bottom=5, bottom_color=C["ink"],
                     top=1, top_color=C["rule"], valign="vcenter", align="center")
        idle = b.f(font_size=11, font_color=C["muted"], bg_color=C["band"], top=1, bottom=1, top_color=C["rule"],
                   bottom_color=C["rule"], valign="vcenter", align="center")
        for i, (sheet, label) in enumerate(((SHEET_CH, "Chiến lược phân phối"), (SHEET_PERF, "Kết quả tài chính"))):
            c1 = 1 + i * 2
            if sheet == self.name:
                ws.merge_range(5, c1, 5, c1 + 1, label, active)
            else:
                ws.merge_range(5, c1, 5, c1 + 1, "", idle)
                ws.write_url(5, c1, f"internal:'{sheet}'!A1", idle, label)
        ws.set_row(6, 9)
        # Rows 8-10: year slicer (added by excel_slicers.ps1) and how to use it
        self.year_slicer_cell = "B8"
        for r in range(7, 10):
            ws.set_row(r, 18)
        ws.merge_range(7, 9, 9, 12, "Bấm một năm để lọc, giữ Ctrl để chọn nhiều năm. Nút xóa lọc ở góc slicer đưa về đủ 7 năm.",
                       b.f(font_size=9, font_color=C["muted"], text_wrap=True, valign="vcenter"))
        ws.set_row(10, 12)
        # Question and summary
        q = t["q_ch"] if self.view == "channels" else t["q_perf"]
        ws.merge_range(11, 1, 15, 6, q, b.f(font_name=SERIF, font_size=17, bold=True, font_color=C["ink"], text_wrap=True))
        self.text(11, 8, 12, t["summary_label"], b.f(font_size=9, font_color=C["muted"]))
        self.formula(12, 8, 15, 12, f"={self.data.summary_cells[self.view]}", b.f(font_size=11.5, font_color=C["ink"], text_wrap=True))
        for r in range(12, 16):
            ws.set_row(r, 21)
        for c in range(1, 13):
            ws.write_blank(16, c, None, b.f(bottom=1, bottom_color=C["rule"]))
        ws.set_row(16, 9)
        ws.set_row(17, 15)
        return 18

    def figure_head(self, r, c1, c2, key, sub_formula=None):
        b = self.b
        title, sub = b.texts["figs"][key]
        self.text(r, c1, c2, title, b.f(bold=True, font_size=12, font_color=C["ink"]), 19)
        if sub_formula:
            self.formula(r + 1, c1, r + 1, c2, sub_formula, b.f(font_size=9.5, font_color=C["muted"], text_wrap=True))
        else:
            self.text(r + 1, c1, c2, sub, b.f(font_size=9.5, font_color=C["muted"], text_wrap=True))
        self.ws.set_row(r + 1, 27)

    def source(self, r, c1, c2, key=None, formula=None):
        src = formula or f'"{self.b.sources[key]}"'
        self.formula(r, c1, r, c2, f"={src}&{self.data.summary_cells['reports']}",
                     self.b.f(font_size=8.5, font_color=C["muted"], text_wrap=True))
        self.ws.set_row(r, 26 if c2 - c1 > 3 else 38)   # narrow (three-column) figures need a third line

    def notes(self, r1, r2, c1, c2, title, items):
        b, ws = self.b, self.ws
        bold = b.f(bold=True, font_color=C["ink"])
        plain = b.f(font_color=C["ink2"])
        cell = b.f(text_wrap=True)
        ws.merge_range(r1, c1, r2, c2, "", cell)
        parts = []
        for i, (dt, dd) in enumerate(items):
            parts += [bold, ("\n" if i else "") + dt + "\n", plain, dd + "\n"]
        ws.write_rich_string(r1, c1, b.f(bold=True, font_size=12, font_color=C["ink"]), title + "\n", *parts, cell)

    def table(self, r):
        b, ws, d = self.b, self.ws, self.data
        view_rows = TABLES[self.name]
        title = b.texts["table_ch"] if self.view == "channels" else b.texts["table_perf"]
        self.text(r, 1, 12, title, b.f(bold=True, font_size=13, font_color=C["ink"]), 22)
        key = ("Số NIKE công bố có số trang 10-K ở dòng nhỏ ngay dưới (bấm vào số trang để mở báo cáo trên SEC). "
               "Số in nghiêng là số nhóm tự tính.")
        if self.view == "performance":
            key = "Triệu USD trừ khi ghi khác. Số âm trong ngoặc đơn. " + key
        if b.pending_count():
            key += " † chưa được kiểm tra."
        self.text(r + 1, 1, 12, key, b.f(font_size=9, font_color=C["muted"]), 18)
        h = r + 2
        head = b.f(bold=True, font_color=C["ink"], bottom=1, bottom_color=C["ink2"], valign="bottom")
        ws.merge_range(h, 1, h, 3, "Chỉ tiêu", head)
        ws.write(h, 4, "Đơn vị", head)
        yr = d.ref("Năm")
        for j in range(len(b.years)):
            ws.write_formula(h, 5 + j, f'=IFERROR(AGGREGATE(15,5,{d.q(yr)},{j + 1}),"")',
                             b.f(bold=True, font_color=C["ink"], bottom=1, bottom_color=C["ink2"], align="right",
                                 num_format='"FY"0', valign="bottom"))
        ws.write_blank(h, 12, None, head)
        ws.set_row(h, 20)
        r = h + 1
        first_year = b.years[0] - 1
        self.table_starts = [h - 2]    # print blocks: title to the end of the first group, then each further group
        for spec in view_rows:
            if spec[0] == "g":
                if r > h + 1:
                    self.table_starts.append(r)
                g = b.f(bold=True, font_color=C["ink"], bg_color=C["band"], valign="vcenter")
                ws.merge_range(r, 1, r, 12, spec[1], g)
                ws.set_row(r, 20)
                r += 1
                continue
            kind, mid, label, unit, _fmt, total = spec
            txt, links, pages = d.text_rows[(self.name, mid)]
            reported = kind == "obs"     # reported figures get a second, small row with their 10-K page
            rule = {"bottom": 1, "bottom_color": C["grid"]}
            top = {"top": 1, "top_color": C["ink2"]} if total else {}
            border = top if reported else {**top, **rule}
            weight = {"bold": True} if total else {}
            ws.merge_range(r, 1, r, 3, label, b.f(font_color=C["ink"], valign="vcenter", **border, **weight))
            ws.write(r, 4, unit, b.f(font_color=C["muted"], valign="vcenter", **border))
            cell = b.f(italic=not reported, align="right", font_color=C["ink"] if reported else C["ink2"], valign="vcenter",
                       **border, **weight)
            for j in range(len(b.years)):
                hdr = xlsxwriter.utility.xl_rowcol_to_cell(h, 5 + j, row_abs=True)
                ws.write_formula(r, 5 + j, f'=IF({hdr}="","",INDEX({d.q(txt)},{hdr}-{first_year}))', cell)
            ws.write_blank(r, 12, None, b.f(**border))
            ws.set_row(r, 20)
            r += 1
            if not reported:
                continue
            small = b.f(font_size=8, font_color=C["muted"], valign="top", **rule)
            ws.merge_range(r, 1, r, 3, "", small)
            ws.write_blank(r, 4, None, small)
            ref = b.f(font_size=8, font_color=C["muted"], underline=1, align="right", valign="top", **rule)
            for j in range(len(b.years)):
                hdr = xlsxwriter.utility.xl_rowcol_to_cell(h, 5 + j, row_abs=True)
                pick_p = f"INDEX({d.q(pages)},{hdr}-{first_year})"
                pick_l = f"INDEX({d.q(links)},{hdr}-{first_year})"
                # &"" keeps an empty page cell empty (INDEX of a blank cell returns 0).
                ws.write_formula(r, 5 + j, f'=IF({hdr}="","",IF({pick_l}="",{pick_p}&"",HYPERLINK({pick_l},{pick_p})))', ref)
            ws.write_blank(r, 12, None, small)
            ws.set_row(r, 13)
            r += 1
        return r


# ------------------------------------------------------------------ charts
def axis_font(size=9):
    return {"name": FONT, "size": size, "color": C["muted"]}


def base_chart(b, kind, subtype=None, legend=True):
    ch = b.wb.add_chart({"type": kind, **({"subtype": subtype} if subtype else {})})
    ch.set_title({"none": True})
    ch.set_chartarea({"border": {"none": True}, "fill": {"color": "#FFFFFF"}})
    ch.set_plotarea({"border": {"none": True}})
    ch.set_legend({"position": "top", "font": {"name": FONT, "size": 9, "color": C["ink2"]}} if legend else {"none": True})
    return ch


def value_axis(fmt, **extra):
    return {"num_format": fmt, "num_font": axis_font(), "line": {"none": True}, "major_tick_mark": "none",
            "major_gridlines": {"visible": True, "line": {"color": C["grid"], "width": 0.75}}, **extra}


def category_axis(**extra):
    return {"num_font": axis_font(), "line": {"color": C["axis"]}, "major_tick_mark": "none", **extra}


def labels(fmt, color=None, position=None, size=8, bold=False):
    # Numeric labels, not labels taken from text cells: Excel attaches per-point custom labels to the
    # visible points in order, so with years filtered out they would land on the wrong columns.
    # The decimal mark therefore follows the viewer's Excel language (34,8% or 34.8%).
    lab = {"value": True, "num_format": fmt, "font": {"name": FONT, "size": size, "color": color or C["ink2"], "bold": bold}}
    if position:
        lab["position"] = position
    return lab


def column(name, values, color, **extra):
    return {"name": name, "values": values, "fill": {"color": color}, "border": {"none": True}, **extra}


def line(name, values, color, **extra):
    return {"name": name, "values": values, "line": {"color": color, "width": 2},
            "marker": {"type": "circle", "size": 7, "fill": {"color": color}, "border": {"color": "#FFFFFF", "width": 1.25}}, **extra}


def build_charts(b, d):
    cats = d.series("Năm tài chính")
    s = d.series
    charts = {}

    ch = base_chart(b, "bar", "percent_stacked")
    ch.add_series(column("NIKE Direct", s("Tỷ trọng NIKE Direct"), C["direct"], categories=cats, gap=45,
                         border={"color": "#FFFFFF", "width": 1},
                         data_labels=labels("0.0%", "#FFFFFF", "center", 9, True)))
    ch.add_series(column("Bán buôn (Wholesale)", s("Tỷ trọng bán buôn"), C["wholesale"], categories=cats,
                         border={"color": "#FFFFFF", "width": 1},   # dark labels: white on orange is only 3.2:1
                         data_labels=labels("0.0%", C["ink"], "center", 9, True)))
    ch.add_series(column("Global Brand Divisions (dưới 0,3%)", s("Tỷ trọng Global Brand Divisions"), C["gbd"], categories=cats))
    ch.set_x_axis(value_axis("0%", min=0, max=1, major_unit=0.25))
    ch.set_y_axis(category_axis(reverse=True, crossing="max", num_font=axis_font(10)))
    charts["share"] = ch

    ch = base_chart(b, "column")
    ch.add_series(column("NIKE Direct", s("NIKE Direct (tỷ USD)"), C["direct"], categories=cats, gap=70,
                         data_labels=labels("0.0", position="outside_end")))
    ch.add_series(column("Bán buôn (Wholesale)", s("Bán buôn (tỷ USD)"), C["wholesale"], categories=cats,
                         data_labels=labels("0.0", position="outside_end")))
    ch.set_y_axis(value_axis("0", min=0))
    ch.set_x_axis(category_axis())
    charts["channelRevenue"] = ch

    ch = base_chart(b, "column")
    ch.add_series(column("NIKE Direct", s("NIKE Direct (biểu đồ tăng trưởng)"), C["direct"], categories=cats, gap=70))
    ch.add_series(column("Bán buôn (Wholesale)", s("Bán buôn (biểu đồ tăng trưởng)"), C["wholesale"], categories=cats))
    ch.set_y_axis(value_axis(f"0%;{MINUS}0%;0%"))
    ch.set_x_axis(category_axis(label_position="low", line={"color": C["ink"], "width": 1.25}))
    charts["growth"] = ch

    ch = base_chart(b, "column")
    ch.add_series(column("NIKE Direct (tổng)", s("NIKE Direct (tỷ USD)"), C["direct_light"], categories=cats, gap=110, overlap=100,
                         data_labels=labels("0.0", position="outside_end")))
    ch.add_series(column("Trong đó: NIKE Brand Digital", s("NIKE Brand Digital (tỷ USD)"), C["digital"], categories=cats))
    ch.set_y_axis(value_axis("0", min=0))
    ch.set_x_axis(category_axis())
    charts["digital"] = ch

    ch = base_chart(b, "column", legend=False)
    ch.add_series(column("Doanh thu", s("Doanh thu NIKE, Inc. (tỷ USD)"), C["total"], categories=cats, gap=60,
                         data_labels=labels("0.0", position="outside_end")))
    ch.set_y_axis(value_axis("0", min=0))
    ch.set_x_axis(category_axis())
    charts["revenue"] = ch

    ch = base_chart(b, "line", legend=False)
    ch.add_series(line("Biên lợi nhuận gộp", s("Biên lợi nhuận gộp"), C["total"], categories=cats,
                       data_labels=labels("0.0%", position="above")))
    ch.set_y_axis(value_axis("0%", min=0.40, max=0.48, major_unit=0.02))
    ch.set_x_axis(category_axis())
    charts["grossMargin"] = ch

    ch = base_chart(b, "column", "stacked")
    ch.add_series(column("Vận hành chung", s("Vận hành chung / doanh thu"), C["overhead"], categories=cats, gap=60,
                         border={"color": "#FFFFFF", "width": 0.75}))
    ch.add_series(column("Tạo nhu cầu (marketing)", s("Tạo nhu cầu / doanh thu"), C["demand"], categories=cats,
                         border={"color": "#FFFFFF", "width": 0.75}))
    total = b.wb.add_chart({"type": "line"})
    total.add_series({"name": "Tổng SG&A / doanh thu", "values": s("Tổng SG&A / doanh thu"), "categories": cats,
                      "line": {"none": True}, "marker": {"type": "none"},
                      "data_labels": labels("0.0%", position="above")})
    ch.combine(total)
    ch.set_y_axis(value_axis("0%", min=0))
    ch.set_x_axis(category_axis())
    ch.set_legend({"position": "top", "font": {"name": FONT, "size": 9, "color": C["ink2"]}, "delete_series": [2]})
    charts["sga"] = ch

    ch = base_chart(b, "line", legend=False)
    ch.add_series(line("Biên EBIT", s("Biên EBIT"), C["total"], categories=cats,
                       data_labels=labels("0.0%", position="above")))
    ch.set_y_axis(value_axis("0%", min=0))
    ch.set_x_axis(category_axis())
    charts["ebit"] = ch

    ch = base_chart(b, "column", legend=False)
    ch.add_series(column("Hàng tồn kho", s("Hàng tồn kho (tỷ USD)"), C["total"], categories=cats, gap=60,
                         data_labels=labels("0.0", position="outside_end")))
    ch.set_y_axis(value_axis("0", min=0))
    ch.set_x_axis(category_axis())
    charts["inventory"] = ch

    ch = base_chart(b, "line", legend=False)
    ch.add_series(line("Số ngày tồn kho", s("Số ngày tồn kho bình quân"), C["total"], categories=cats,
                       data_labels=labels("0", position="above")))
    ch.set_y_axis(value_axis("0", min=90, max=120, major_unit=10))
    ch.set_x_axis(category_axis())
    charts["days"] = ch
    return charts


def place(ws, r, c, chart, width, height):
    chart.set_size({"width": width, "height": height})
    ws.insert_chart(r, c, chart, {"x_offset": 0, "y_offset": 0, "object_position": 1})


# ------------------------------------------------------------------ assembly
def build_channels(b, d, ws, charts):
    p = Dashboard(b, d, ws, SHEET_CH, "channels")
    r = p.header()
    starts = [0]
    p.figure_head(r, 1, 12, "share")
    place(ws, r + 2, 1, charts["share"], 1200, 340)
    p.source(r + 19, 1, 12, "share")
    r += 21
    starts.append(r)
    # Row: channel revenue | growth (the basis slicer sits above the growth chart: 3 rows of 24 px)
    p.figure_head(r, 1, 6, "channelRevenue")
    p.figure_head(r, 7, 12, "growth", sub_formula=f"={d.summary_cells['growth_sub']}")
    p.growth_slicer_cell = xlsxwriter.utility.xl_rowcol_to_cell(r + 2, 7)
    for k in (2, 3, 4):
        ws.set_row(r + k, 18)
    place(ws, r + 2, 1, charts["channelRevenue"], 590, 352)
    place(ws, r + 5, 7, charts["growth"], 590, 280)
    p.source(r + 19, 1, 6, "channelRevenue")
    p.source(r + 19, 7, 12, formula=d.summary_cells["growth_source"])
    r += 21
    starts.append(r)
    # Row: digital | reading notes
    p.figure_head(r, 1, 6, "digital")
    place(ws, r + 2, 1, charts["digital"], 590, 300)
    p.source(r + 17, 1, 6, "digital")
    title, items = b.texts["notes_ch"]
    p.notes(r, r + 17, 7, 12, title, items)
    r += 19
    end = p.table(r)
    p.paginate(starts + p.table_starts, end)
    return p


def build_performance(b, d, ws, charts):
    p = Dashboard(b, d, ws, SHEET_PERF, "performance")
    r = p.header()
    starts = [0]
    for n, row_keys in enumerate((("revenue", "grossMargin", "sga"), ("ebit", "inventory", "days"))):
        if n:
            starts.append(r)
        for i, key in enumerate(row_keys):
            c1 = 1 + 4 * i
            p.figure_head(r, c1, c1 + 3, key)
            place(ws, r + 2, c1, charts[key], 390, 260)
            p.source(r + 15, c1, c1 + 3, key)
        r += 17
    starts.append(r)
    title, items = b.texts["notes_perf"]
    p.text(r, 1, 12, title, b.f(bold=True, font_size=12, font_color=C["ink"]), 20)
    for i, (dt, dd) in enumerate(items):
        c1 = 1 + 4 * i
        ws.merge_range(r + 1, c1, r + 4, c1 + 3, "", b.f(text_wrap=True))
        ws.write_rich_string(r + 1, c1, b.f(bold=True, font_color=C["ink"]), dt + "\n", b.f(font_color=C["ink2"]), dd, b.f(text_wrap=True))
    # Height follows the longest note: its title line plus about 56 characters a line (12.75 pt) across four columns.
    lines = max(1 + math.ceil(len(dd) / 56) for _dt, dd in items)
    for k in range(1, 5):
        ws.set_row(r + k, max(20, math.ceil(lines * 12.75 / 4) + 1))
    r += 6
    end = p.table(r)
    p.paginate(starts + p.table_starts, end)
    return p


def build_sources(b, ws):
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2.71)
    ws.set_column(1, 1, 62)
    ws.set_column(2, 2, 4)
    ws.set_column(3, 3, 62)
    t = b.texts
    ws.write(1, 1, "Nguồn và phương pháp", b.f(font_name=SERIF, font_size=20, bold=True, font_color=C["ink"]))
    ws.set_row(1, 32)
    h = b.f(bold=True, font_size=11, font_color=C["ink"], bottom=1, bottom_color=C["rule"])
    ws.write(3, 1, "Báo cáo Form 10-K (SEC EDGAR)", h)
    link = b.f(font_color=C["link"], underline=1)
    for i, y in enumerate(b.years):
        f = b.d["filings"][str(y)]
        dd, mm, yy = f["filed"][8:10], f["filed"][5:7], f["filed"][:4]
        ws.write_url(4 + i, 1, f["url"], link, f"Form 10-K FY{y}, nộp ngày {dd}/{mm}/{yy}, số {f['accession']}")
    ws.write(3, 3, "Phương pháp", h)
    for i, item in enumerate(t["method"]):
        ws.write(4 + i, 3, "–  " + item, b.f(text_wrap=True))
    # Row height follows the longest bullet on the row (about 80 characters per line in a 62-wide column).
    for i in range(max(len(b.years), len(t["method"]))):
        lines = math.ceil(len(t["method"][i]) / 80) if i < len(t["method"]) else 1
        ws.set_row(4 + i, max(30, 15 * lines + 3))
    r = 4 + max(len(b.years), len(t["method"])) + 1
    ws.write(r, 1, "Tình trạng dữ liệu", h)
    c = b.d["meta"]["status_counts"]
    total = sum(c.values())
    parts = [(c.get("Khớp nguồn chính thức", 0), " khớp nguồn chính thức"), (c.get("Đã xác minh (người)", 0), " đã được nhóm kiểm tra trên 10-K"),
             (c.get(STATUS_PENDING, 0), " chưa được kiểm tra (†)"), (c.get("Thiếu / không công bố", 0), " không được NIKE công bố")]
    extracted = b.d["meta"]["extracted_on"]
    status = (f"{total} số liệu báo cáo: " + ", ".join(f"{n}{label}" for n, label in parts if n) +
              f". Dữ liệu trích xuất ngày {extracted[8:10]}/{extracted[5:7]}/{extracted[:4]}.")
    ws.merge_range(r + 1, 1, r + 2, 1, status, b.f(text_wrap=True))
    ws.write(r + 3, 1, t["fine"], b.f(font_size=9, font_color=C["muted"]))
    ws.write(r, 3, "Liên kết", h)
    ws.write_url(r + 1, 3, DASHBOARD_URL, link, "Dashboard web")
    ws.write_url(r + 2, 3, DASHBOARD_URL + b.d["meta"]["workbook"], link, "Bộ dữ liệu Excel (đầy đủ nguồn từng số)")
    ws.write_url(r + 3, 3, DASHBOARD_URL + EVIDENCE_PAGE, link, "Đối chiếu 35 số trên trang 10-K (ảnh chụp)")
    ws.set_landscape()
    ws.fit_to_pages(1, 0)


def build(path=OUT, add_slicers=True):
    data = json.loads(SITE_DATA.read_text(encoding="utf-8"))
    b = Book(path, data)
    # Tab order: two dashboards, sources, data. excel_slicers.ps1 finds the data sheet as the last tab.
    ws_ch, ws_perf, ws_src, ws_data = (b.wb.add_worksheet(n) for n in (SHEET_CH, SHEET_PERF, SHEET_SRC, SHEET_DATA))
    d = DataSheet(b, ws_data)
    d.write()
    charts = build_charts(b, d)
    ch = build_channels(b, d, ws_ch, charts)
    perf = build_performance(b, d, ws_perf, charts)
    build_sources(b, ws_src)
    ws_data.hide()   # sorting or editing the year table would break the slicer-driven summaries
    ws_ch.activate()
    b.wb.close()
    yw, yh = Dashboard.YEAR_SLICER
    bw, bh = Dashboard.BASIS_SLICER
    anchors = dict(year=[f"1:{ch.year_slicer_cell}:{yw}:{yh}", f"2:{perf.year_slicer_cell}:{yw}:{yh}"],
                   basis=f"1:{ch.growth_slicer_cell}:{bw}:{bh}")
    if add_slicers:
        run_slicer_script(path, anchors)
    return anchors


def run_slicer_script(path, anchors):
    if sys.platform != "win32":
        print("Slicers not added: needs Windows with Excel (pipeline/excel_slicers.ps1).")
        return
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(SLICER_SCRIPT), "-Path", str(path),
           "-YearAnchors", ",".join(anchors["year"]), "-BasisAnchor", anchors["basis"]]
    subprocess.run(cmd, check=True)


def main():
    build()
    print(f"Excel dashboard written: {OUT.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
