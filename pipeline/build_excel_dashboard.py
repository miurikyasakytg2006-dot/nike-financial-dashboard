"""Build the Excel version of the dashboard, for viewing in Excel or Excel for the web (OneDrive).

Usage:  python -m pipeline.build_excel_dashboard

Reads site/data/nike-financials.json (the validated data layer of the web dashboard) and the static texts of
site/index.html and site/assets/app.js, and writes handover/NIKE_Dashboard_FY2019-2021.xlsx with the same sections,
key findings, charts, tables, NIKE's own explanations and sources. The key findings and notes are the strings the
pipeline wrote into the JSON, so the three dashboard versions say the same thing. Charts read a hidden data sheet.
On Windows the workbook is then opened and saved once by the installed Excel (pipeline/excel_save.ps1), which stores
the chart caches that Excel for the web shows.
"""
import json
import math
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import xlsxwriter
from lxml import html
from xlsxwriter.utility import xl_rowcol_to_cell

from .build_workbook import DASHBOARD_URL, git_user_name
from .dataset import vi_num

ROOT = Path(__file__).resolve().parent.parent
SITE_DATA = ROOT / "site" / "data" / "nike-financials.json"
INDEX_HTML = ROOT / "site" / "index.html"
APP_JS = ROOT / "site" / "assets" / "app.js"
OUT = ROOT / "handover" / "NIKE_Dashboard_FY2019-2021.xlsx"
SAVE_SCRIPT = Path(__file__).resolve().parent / "excel_save.ps1"

SHEET_OVERVIEW = "Tổng quan"
SHEET_Q = "Theo quý"
SHEET_PROFIT = "Khu vực và lợi nhuận"
SHEET_TABLES = "Bảng số liệu"
SHEET_SRC = "Nguồn và giải thích"
SHEET_DATA = "Dữ liệu biểu đồ"
SHEETS = [SHEET_OVERVIEW, SHEET_Q, SHEET_PROFIT, SHEET_TABLES, SHEET_SRC]
STATUS_PENDING = "Chờ người kiểm tra"
COVID = "FY2020Q4"

# Same palette as site/assets/app.js
C = dict(direct="#2a78d6", wholesale="#eb6834", gbd="#a3a9b2", direct_light="#86b6ef", digital="#1c5cab",
         total="#475467", margin="#6250d6", up="#1baf7a", down="#c8453b", past="#98a2b3", ink="#15191e", ink2="#3b4350",
         muted="#5b6472", faint="#6f7784", grid="#e6e9ee", axis="#c3c9d1", rule="#d9dde3", band="#f4f5f7", link="#1f4e9a")
FONT, SERIF = "Arial", "Cambria"
MINUS = "−"
NBSP, NB_HYPHEN = "\u00a0", "\u2011"   # escapes, so that no editor turns them into a space and a hyphen

# Page grid: column A is a margin, B..M are twelve 100-pixel columns (1,200 px of content).
GRID_W = 13.57
# Printing: a fixed scale, because Excel ignores manual page breaks under fit-to-page scaling. At 79% the grid
# (about 936 pt with the margin columns) fits A4 and Letter landscape; the breaks keep chart rows whole.
PRINT_SCALE = 79
PAGE_ROWS_PT = (595.3 - 2 * 0.4 * 72) * 100 / PRINT_SCALE   # row height that fits on an A4-landscape page


def glue(text):
    """Keep together what reads badly when a wrapped cell splits it. Excel ends a line after any hyphen, so "10-K",
    "10-Q" and "non-GAAP" take a non-breaking hyphen (Arial has none; Excel draws a matching one from a fallback font).
    A report stays with its period ("10-K FY2020", "10-Q Q3 FY2020", "Q4 FY2020"), a figure with its unit, and the two
    ends of a period range ("FY2019 → FY2021")."""
    text = (text.replace("10-K", "10" + NB_HYPHEN + "K").replace("10-Q", "10" + NB_HYPHEN + "Q")
            .replace("non-GAAP", "non" + NB_HYPHEN + "GAAP"))
    text = re.sub(r"(10" + NB_HYPHEN + "[KQ]) (?=Q[1-4]|FY)", "\\1" + NBSP, text)
    text = re.sub(r"\b(Q[1-4]) (FY\d{4})", "\\1" + NBSP + "\\2", text)
    text = re.sub(r"(FY\d{4}) → (FY\d{4})", "\\1" + NBSP + "→" + NBSP + "\\2", text)
    text = re.sub(r"(\d) (tỷ|triệu) USD", "\\1" + NBSP + "\\2" + NBSP + "USD", text)
    text = re.sub(r"(\d) ngày", "\\1" + NBSP + "ngày", text)
    return re.sub(r"(\d) điểm %", "\\1" + NBSP + "điểm" + NBSP + "%", text)


# ------------------------------------------------------------------ number formatting (vi-VN, as in app.js)
def num(v, d=0):
    return "—" if v is None else (MINUS if v < 0 else "") + vi_num(abs(v), d)


def signed(v, d=0):
    if v is None:
        return "—"
    s = vi_num(abs(v), d)
    return ("" if set(s) <= set("0,.") else "+" if v > 0 else MINUS) + s


def amount(v):
    return "—" if v is None else (f"({vi_num(abs(v), 0)})" if v < 0 else vi_num(v, 0))


def pct(v, d=1):
    return "—" if v is None else num(v * 100, d) + "%"


def signed_pct(v, d=1):
    return "—" if v is None else signed(v * 100, d) + "%"


def pts(v, d=1):
    return "—" if v is None else signed(v * 100, d) + " điểm %"


def bn(millions, d=1):
    return "—" if millions is None else num(millions / 1000, d)


def stated(v):
    """Digital sales as NIKE stated them: in millions for fiscal 2019 quarters (0,77), to 0.1 billion afterwards."""
    if v is None:
        return "—"
    s = f"{round(v, 3):.3f}".rstrip("0")
    return num(v, max(1, len(s.split(".")[1])))


def bn_delta(v):
    return signed(v, 2 if abs(v) < 0.1 else 1)


# Table rows, as in app.js: (kind, id or group title, label, unit, formatter, total row)
CHANNEL_ROWS = [
    ("g", "Doanh thu NIKE Brand theo kênh"),
    ("obs", "nike_direct_revenue", "NIKE Direct", "triệu USD", amount, False),
    ("obs", "nike_brand_wholesale_revenue", "Bán buôn (Wholesale)", "triệu USD", amount, False),
    ("obs", "global_brand_divisions_revenue", "Global Brand Divisions", "triệu USD", amount, False),
    ("obs", "nike_brand_revenue", "Tổng doanh thu NIKE Brand", "triệu USD", amount, True),
    ("g", "Cơ cấu và tăng trưởng"),
    ("kpi", "direct_share", "Tỷ trọng NIKE Direct", "%", pct, False),
    ("obs", "nike_direct_share_stated", "Tỷ trọng NIKE Direct do NIKE nêu (xấp xỉ)", "%", lambda v: pct(v, 0), False),
    ("kpi", "direct_growth", "Tăng trưởng NIKE Direct, theo báo cáo", "%", signed_pct, False),
    ("obs", "nike_direct_pct_change_cn", "Tăng trưởng NIKE Direct, loại trừ tỷ giá (non-GAAP)", "%", lambda v: signed_pct(v, 0), False),
    ("kpi", "wholesale_growth", "Tăng trưởng bán buôn, theo báo cáo", "%", signed_pct, False),
    ("obs", "nike_brand_wholesale_pct_change_cn", "Tăng trưởng bán buôn, loại trừ tỷ giá (non-GAAP)", "%", lambda v: signed_pct(v, 0), False),
    ("kpi", "nike_brand_growth", "Tăng trưởng NIKE Brand, theo báo cáo", "%", signed_pct, False),
    ("kpi", "direct_contribution", "Trong đó: đóng góp của NIKE Direct", "điểm %", lambda v: signed(v * 100, 1), False),
    ("kpi", "wholesale_contribution", "Trong đó: đóng góp của bán buôn", "điểm %", lambda v: signed(v * 100, 1), False),
    ("g", "NIKE Brand Digital (một phần của NIKE Direct)"),
    ("obs", "nike_brand_digital_sales", "Doanh số Digital", "tỷ USD", stated, False),
    ("kpi", "digital_share_of_direct", "Digital trong NIKE Direct (xấp xỉ)", "%", lambda v: pct(v, 0), False),
    ("kpi", "direct_ex_digital", "NIKE Direct ngoài Digital (ước tính)", "triệu USD", amount, False),
]
FINANCE_ROWS = [
    ("g", "Báo cáo kết quả kinh doanh (NIKE, Inc.)"),
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
    ("kpi", "demand_creation_ratio", "Chi phí tạo nhu cầu / doanh thu", "%", pct, False),
    ("kpi", "operating_overhead_ratio", "Chi phí vận hành chung / doanh thu", "%", pct, False),
    ("kpi", "net_income_margin", "Biên lợi nhuận sau thuế", "%", pct, False),
    ("g", "EBIT (non-GAAP, định nghĩa của NIKE)"),
    ("obs", "ebit_reported", "EBIT do NIKE công bố", "triệu USD", amount, False),
    ("kpi", "ebit_margin", "Biên EBIT", "%", pct, False),
    ("g", "Hàng tồn kho"),
    ("obs", "inventories", "Hàng tồn kho cuối kỳ", "triệu USD", amount, False),
    ("kpi", "inventory_growth", "So với cùng thời điểm năm trước", "%", signed_pct, False),
    ("kpi", "days_inventory", "Số ngày tồn kho bình quân", "ngày", lambda v: num(v, 0), False),
]


def region_rows(regions):
    rows = []
    for r in regions:
        rid = r["id"]
        rows += [("g", r["vi"] if r["vi"] == r["short"] else f"{r['vi']} ({r['short']})"),
                 ("obs", f"{rid}_revenue", "Doanh thu", "triệu USD", amount, True),
                 ("obs", f"{rid}_direct_revenue", "NIKE Direct", "triệu USD", amount, False),
                 ("obs", f"{rid}_wholesale_revenue", "Bán buôn", "triệu USD", amount, False),
                 ("kpi", f"{rid}_direct_share", "Tỷ trọng NIKE Direct", "%", pct, False),
                 ("kpi", f"{rid}_growth", "Tăng trưởng doanh thu, theo báo cáo", "%", signed_pct, False),
                 ("obs", f"{rid}_ebit", "EBIT", "triệu USD", amount, False),
                 ("kpi", f"{rid}_ebit_margin", "Biên EBIT", "%", pct, False)]
    return rows


# ------------------------------------------------------------------ texts from the web page
def for_excel(text):
    """Web wording that does not hold in the workbook: the grey band of the quarterly charts, and where the list of
    reports is."""
    text = re.sub(r" ?Vùng tô xám là quý 4 FY2020[^.]*\.", "", text)
    return text.replace("(danh sách ở cuối trang)", f"(danh sách ở sheet “{SHEET_SRC}”)")


def page_texts():
    root = html.parse(str(INDEX_HTML)).getroot()

    def text(el):
        return glue(" ".join(el.text_content().split()))

    def t(xpath):
        return text(root.xpath(xpath)[0])

    def fig(fid):
        f = root.xpath(f"//figure[@id='{fid}']")[0]
        sub = f.xpath(".//p[@class='figure-sub']")
        return text(f.xpath(".//h3")[0]), for_excel(text(sub[0])) if sub else ""

    section = lambda sid: (t(f"//section[@id='{sid}']//p[@class='block-kicker']"), t(f"//section[@id='{sid}']//h2"))
    notes = root.xpath("//aside[@aria-labelledby='notes-title']")[0]
    return dict(
        title=t("//header[@class='masthead']//h1"), subtitle=t("//p[@class='masthead-sub']"),
        findings_title=t("//section[@id='phat-hien']//h2"), findings_note=t("//section[@id='phat-hien']//p[@class='block-note']"),
        sections={sid: section(sid) for sid in ("kenh", "quy", "digital", "khu-vuc", "loi-nhuan", "nike-giai-thich")},
        quotes_note=t("//section[@id='nike-giai-thich']//p[@class='block-note']"),
        figs={k: fig(v) for k, v in dict(share="fig-share", change="fig-change", shareQ="fig-share-q", channelQ="fig-channel-q",
                                         digital="fig-digital", digitalQ="fig-digital-q", regions="fig-regions",
                                         regionsEbit="fig-regions-ebit", bridge="fig-ebit-bridge", marginsQ="fig-margins-q",
                                         inventoryQ="fig-inventory-q", daysQ="fig-days-q").items()},
        notes_title=text(notes.xpath(".//h3")[0]),
        notes=[(text(a), text(b)) for a, b in zip(notes.xpath(".//dt"), notes.xpath(".//dd"))],
        method=[text(li) for li in root.xpath("//section[@aria-labelledby='footer-method']//li")],
        fine=t("//p[@class='fine']"),
    )


def app_texts():
    """Chart source lines and the two period-dependent subtitles, read from app.js so the versions cannot drift."""
    js = APP_JS.read_text(encoding="utf-8")
    sources = {}
    for cid, body in re.findall(r"sourceLine\('(\w+)', (.+?)\);\n", js, flags=re.S):
        m = re.match(r"(?:quarterSources\()?'([^']*)'", body)
        text = m.group(1)
        if body.startswith("quarterSources("):
            tail = re.search(r"function quarterSources\(text\) \{\s*return text \+ '([^']*)'", js).group(1)
            text += tail
        sources[cid] = glue(for_excel(text))
    subs = {}
    for cid in ("change-sub", "bridge-sub"):
        m = re.search(rf"getElementById\('{cid}'\)\.textContent = '([^']*)' \+ a \+ '([^']*)' \+ b \+\s*'([^']*)'", js)
        subs[cid] = lambda a, b, m=m: glue(m.group(1) + a + m.group(2) + b + m.group(3))
    return sources, subs


# ------------------------------------------------------------------ workbook
class Book:
    def __init__(self, path, data):
        self.d = data
        # Arial 10 rows are 17 px by default (XlsxWriter assumes Calibri's 20 px and would then skip writing the
        # 20 px rows the layout is designed in, leaving Excel to draw them at 17 px).
        self.wb = xlsxwriter.Workbook(str(path), {"default_row_height": 17})
        self.wb.set_properties({"title": "NIKE FY2019–FY2021: bán trực tiếp và kết quả tài chính",
                                "comments": "Phiên bản Excel của dashboard. Nguồn: NIKE, Inc. Form 10-K và 10-Q FY2019–FY2021 (SEC EDGAR)."})
        self.wb.formats[0].set_font_name(FONT)
        self.wb.formats[0].set_font_size(10)
        self.cache = {}
        self.texts = page_texts()
        self.sources, self.subs = app_texts()
        self.A, self.Q = data["periods"]["A"], data["periods"]["Q"]

    def f(self, **p):
        base = {"font_name": FONT, "font_size": 10, "font_color": C["ink2"], "valign": "top"}
        key = tuple(sorted(p.items()))
        if key not in self.cache:
            self.cache[key] = self.wb.add_format({**base, **p})
        return self.cache[key]

    def ob(self, mid, p):
        return self.d["observations"][mid].get(p) or {}

    def val(self, mid, p):
        return self.ob(mid, p).get("value")

    def kpi(self, kid, p):
        return self.d["kpis"][kid]["values"][p]

    def label(self, p):
        return self.d["labels"][p]


class DataSheet:
    """Hidden sheet with every series the charts plot, and the text of each custom data label."""

    def __init__(self, b, ws):
        self.b, self.ws = b, ws
        self.row = 2
        self.ranges = {}
        ws.write(0, 0, "Dữ liệu cho biểu đồ", b.f(bold=True, font_size=14, font_color=C["ink"]))
        ws.write(1, 0, "Lấy từ site/data/nike-financials.json (cùng dữ liệu với dashboard web). Không sửa sheet này.",
                 b.f(font_color=C["muted"]))
        ws.set_column(0, 1, 14)
        ws.set_column(2, 30, 16)

    def block(self, name, cats, columns):
        """cats: list of category values (or of (year, quarter) pairs for a two-level axis);
        columns: {header: list of values}. Returns references by header (and "cats")."""
        ws, r0 = self.ws, self.row
        ws.write(r0, 0, name, self.b.f(bold=True))
        two = isinstance(cats[0], tuple)
        ncat = 2 if two else 1
        heads = (["Năm", "Quý"] if two else ["Kỳ"]) + list(columns)
        ws.write_row(r0 + 1, 0, heads, self.b.f(bold=True))
        for i, cat in enumerate(cats):
            r = r0 + 2 + i
            ws.write_row(r, 0, list(cat) if two else [cat])
            for j, values in enumerate(columns.values()):
                v = values[i]
                if v is None:
                    ws.write_blank(r, ncat + j, None)
                else:
                    ws.write(r, ncat + j, v)
        first, last = r0 + 2, r0 + 1 + len(cats)
        refs = {"cats": [SHEET_DATA, first, 0, last, ncat - 1]}
        for j, h in enumerate(columns):
            c = ncat + j
            refs[h] = [SHEET_DATA, first, c, last, c]
            refs[h + "@cells"] = [f"='{SHEET_DATA}'!{xl_rowcol_to_cell(first + i, c, row_abs=True, col_abs=True)}"
                                  for i in range(len(cats))]
        self.row = last + 3
        self.ranges[name] = refs
        return refs


# ------------------------------------------------------------------ charts
def axis_font(size=9, color=None):
    return {"name": FONT, "size": size, "color": color or C["muted"]}


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


def text_labels(cells, keep, color=None, position=None, size=8, bold=False):
    """Data labels taken from text cells (Vietnamese number format, as on the web), shown only at `keep` indexes."""
    lab = {"value": True, "font": {"name": FONT, "size": size, "color": color or C["ink2"], "bold": bold},
           "custom": [{"value": cell} if i in keep else {"delete": True} for i, cell in enumerate(cells)]}
    if position:
        lab["position"] = position
    return lab


def line(name, values, color, cats, **extra):
    return {"name": name, "values": values, "categories": cats, "line": {"color": color, "width": 2},
            "marker": {"type": "circle", "size": 7, "fill": {"color": color}, "border": {"color": "#FFFFFF", "width": 1.25}},
            **extra}


def bars(name, values, color, cats, **extra):
    return {"name": name, "values": values, "categories": cats, "fill": {"color": color}, "border": {"none": True}, **extra}


def build_charts(b, d):
    A, Q = b.A, b.Q
    k, v = b.kpi, b.val
    covid = Q.index(COVID)
    qcats = [(p[:6] if p.endswith("Q1") else "", "Q" + p[7]) for p in Q]
    first_last = lambda n: {0, n - 1}
    charts = {}

    # 1. channel mix (annual 100% bars)
    r = d.block("Cơ cấu kênh", A, {
        "NIKE Direct": [k("direct_share", p) for p in A], "Bán buôn": [k("wholesale_share", p) for p in A],
        "GBD": [k("gbd_share", p) for p in A],
        "nhãn Direct": [pct(k("direct_share", p)) for p in A], "nhãn bán buôn": [pct(k("wholesale_share", p)) for p in A]})
    ch = base_chart(b, "bar", "percent_stacked")
    ch.add_series(bars("NIKE Direct", r["NIKE Direct"], C["direct"], r["cats"], gap=60, border={"color": "#FFFFFF", "width": 1},
                       data_labels=text_labels(r["nhãn Direct@cells"], set(range(3)), "#FFFFFF", "center", 9, True)))
    # Dark labels on orange: white on the wholesale orange is only 3.2:1.
    ch.add_series(bars("Bán buôn (Wholesale)", r["Bán buôn"], C["wholesale"], r["cats"], border={"color": "#FFFFFF", "width": 1},
                       data_labels=text_labels(r["nhãn bán buôn@cells"], set(range(3)), C["ink"], "center", 9, True)))
    ch.add_series(bars("Global Brand Divisions (dưới 0,2%)", r["GBD"], C["gbd"], r["cats"]))
    ch.set_x_axis(value_axis("0%", min=0, max=1, major_unit=0.25))
    ch.set_y_axis(category_axis(reverse=True, crossing="max", num_font=axis_font(10)))
    charts["share"] = ch

    # 1. change by channel, FY2019 -> FY2021
    a, z = A[0], A[-1]
    items = [("NIKE Direct", "nike_direct_revenue", C["direct"]), ("Bán buôn", "nike_brand_wholesale_revenue", C["wholesale"]),
             ("Global Brand Divisions", "global_brand_divisions_revenue", C["gbd"]), ("Tổng NIKE Brand", "nike_brand_revenue", C["total"])]
    deltas = [(v(mid, z) - v(mid, a)) / 1000 for _, mid, _ in items]
    r = d.block("Thay đổi theo kênh", [n for n, _, _ in items], {"Thay đổi (tỷ USD)": deltas,
                                                                  "nhãn": [bn_delta(x) for x in deltas]})
    ch = base_chart(b, "bar", legend=False)
    ch.add_series(bars("Thay đổi", r["Thay đổi (tỷ USD)"], C["total"], r["cats"], gap=60,
                       points=[{"fill": {"color": c}, "border": {"none": True}} for _, _, c in items],
                       data_labels=text_labels(r["nhãn@cells"], set(range(4)), C["ink2"], "outside_end", 9)))
    ch.set_x_axis(value_axis("0"))
    ch.set_y_axis(category_axis(reverse=True, crossing="max", num_font=axis_font(10, C["ink2"]), label_position="low"))
    charts["change"] = ch

    # 2. quarters: direct share
    share = [k("direct_share", p) for p in Q]
    hi = share.index(max(share))
    r = d.block("Tỷ trọng NIKE Direct theo quý", qcats, {"Tỷ trọng": share, "nhãn": [pct(x) for x in share]})
    ch = base_chart(b, "line", legend=False)
    ch.add_series(line("Tỷ trọng NIKE Direct", r["Tỷ trọng"], C["direct"], r["cats"],
                       data_labels=text_labels(r["nhãn@cells"], {0, hi, len(Q) - 1}, position="above")))
    ch.set_y_axis(value_axis("0%", min=0.25, max=0.5, major_unit=0.05))
    ch.set_x_axis(category_axis())
    charts["shareQ"] = ch

    # 2. quarters: revenue by channel
    dq = [v("nike_direct_revenue", p) / 1000 for p in Q]
    wq = [v("nike_brand_wholesale_revenue", p) / 1000 for p in Q]
    r = d.block("Doanh thu theo kênh, từng quý", qcats, {"NIKE Direct": dq, "Bán buôn": wq,
                                                         "nhãn Direct": [num(x, 1) for x in dq], "nhãn bán buôn": [num(x, 1) for x in wq]})
    ch = base_chart(b, "line")
    keep = {0, covid, len(Q) - 1}
    ch.add_series(line("NIKE Direct", r["NIKE Direct"], C["direct"], r["cats"],
                       data_labels=text_labels(r["nhãn Direct@cells"], keep, position="below")))
    ch.add_series(line("Bán buôn (Wholesale)", r["Bán buôn"], C["wholesale"], r["cats"],
                       data_labels=text_labels(r["nhãn bán buôn@cells"], keep, position="above")))
    ch.set_y_axis(value_axis("0", min=0))
    ch.set_x_axis(category_axis())
    charts["channelQ"] = ch

    # 3. digital inside NIKE Direct (annual)
    dig = [v("nike_brand_digital_sales", p) for p in A]
    rest = [k("direct_ex_digital", p) / 1000 for p in A]
    r = d.block("Digital và phần còn lại", A, {"Digital": dig, "Phần còn lại": rest,
                                               "nhãn Digital": [num(x, 1) for x in dig], "nhãn còn lại": [num(x, 1) for x in rest]})
    ch = base_chart(b, "bar", "stacked")
    ch.add_series(bars("NIKE Brand Digital", r["Digital"], C["digital"], r["cats"], gap=60, border={"color": "#FFFFFF", "width": 1},
                       data_labels=text_labels(r["nhãn Digital@cells"], set(range(3)), "#FFFFFF", "center", 9, True)))
    ch.add_series(bars("Phần còn lại, chủ yếu là cửa hàng (ước tính)", r["Phần còn lại"], C["direct_light"], r["cats"],
                       border={"color": "#FFFFFF", "width": 1},
                       data_labels=text_labels(r["nhãn còn lại@cells"], set(range(3)), C["ink"], "center", 9, True)))
    ch.set_x_axis(value_axis("0", min=0))
    ch.set_y_axis(category_axis(reverse=True, crossing="max", num_font=axis_font(10)))
    charts["digital"] = ch

    # 3. digital share of NIKE Direct (quarters)
    ds = [k("digital_share_of_direct", p) for p in Q]
    hi = ds.index(max(ds))
    r = d.block("Digital / NIKE Direct, từng quý", qcats, {"Tỷ trọng Digital": ds, "nhãn": ["≈" + pct(x, 0) for x in ds]})
    ch = base_chart(b, "line", legend=False)
    q4_points = [{"fill": {"color": "#FFFFFF"}, "border": {"color": C["digital"], "width": 1.5}} if p.endswith("Q4") else None
                 for p in Q]
    ch.add_series(line("Digital / NIKE Direct", r["Tỷ trọng Digital"], C["digital"], r["cats"], points=q4_points,
                       data_labels=text_labels(r["nhãn@cells"], {0, hi, len(Q) - 1}, position="above")))
    ch.set_y_axis(value_axis("0%", min=0, max=0.8, major_unit=0.2))
    ch.set_x_axis(category_axis())
    charts["digitalQ"] = ch

    # 4. regions: FY2019 and FY2021 side by side (Excel has no dumbbell chart)
    regions = b.d["regions"]
    rnames = ["Toàn NIKE Brand"] + [x["short"] for x in regions]
    s19 = [k("direct_share", a)] + [k(f"{x['id']}_direct_share", a) for x in regions]
    s21 = [k("direct_share", z)] + [k(f"{x['id']}_direct_share", z) for x in regions]
    r = d.block("Tỷ trọng NIKE Direct theo khu vực", rnames, {a: s19, z: s21, "nhãn " + a: [pct(x) for x in s19],
                                                              "nhãn " + z: [pct(x) for x in s21]})
    ch = base_chart(b, "bar")
    ch.add_series(bars(a, r[a], C["past"], r["cats"], gap=70, overlap=-10,
                       data_labels=text_labels(r["nhãn " + a + "@cells"], set(range(5)), C["ink2"], "outside_end", 8)))
    ch.add_series(bars(z, r[z], C["direct"], r["cats"],
                       data_labels=text_labels(r["nhãn " + z + "@cells"], set(range(5)), C["ink"], "outside_end", 8, True)))
    ch.set_x_axis(value_axis("0%", min=0, max=0.6, major_unit=0.1))
    ch.set_y_axis(category_axis(reverse=True, num_font=axis_font(10, C["ink2"])))
    charts["regions"] = ch

    e19 = [k(f"{x['id']}_ebit_margin", a) for x in regions]
    e21 = [k(f"{x['id']}_ebit_margin", z) for x in regions]
    r = d.block("Biên EBIT theo khu vực", [x["short"] for x in regions], {a: e19, z: e21, "nhãn " + a: [pct(x) for x in e19],
                                                                          "nhãn " + z: [pct(x) for x in e21]})
    ch = base_chart(b, "bar")
    ch.add_series(bars(a, r[a], C["past"], r["cats"], gap=70, overlap=-10,
                       data_labels=text_labels(r["nhãn " + a + "@cells"], set(range(4)), C["ink2"], "outside_end", 8)))
    ch.add_series(bars(z, r[z], C["direct"], r["cats"],
                       data_labels=text_labels(r["nhãn " + z + "@cells"], set(range(4)), C["ink"], "outside_end", 8, True)))
    ch.set_x_axis(value_axis("0%", min=0, max=0.5, major_unit=0.1))
    ch.set_y_axis(category_axis(reverse=True, num_font=axis_font(10, C["ink2"])))
    charts["regionsEbit"] = ch

    # 5. EBIT margin bridge FY2019 -> FY2021 (stacked columns over an invisible base)
    steps = [("Biên gộp", k("gross_margin", z) - k("gross_margin", a)),
             ("Tạo nhu cầu", -(k("demand_creation_ratio", z) - k("demand_creation_ratio", a))),
             ("Vận hành chung", -(k("operating_overhead_ratio", z) - k("operating_overhead_ratio", a)))]
    start, end = k("ebit_margin", a), k("ebit_margin", z)
    steps.append(("Khác", end - start - sum(s for _, s in steps)))
    cats, base, height, colors, texts, running = [a], [0], [start], [C["total"]], [pct(start)], start
    for name, s in steps:
        cats.append(name)
        base.append(running if s >= 0 else running + s)
        height.append(abs(s))
        colors.append(C["up"] if s >= 0 else C["down"])
        texts.append(signed(s * 100, 1))
        running += s
    cats.append(z)
    base.append(0)
    height.append(end)
    colors.append(C["total"])
    texts.append(pct(end))
    r = d.block("Biên EBIT theo từng khoản", cats, {"Nền": base, "Cột": height, "nhãn": texts})
    ch = base_chart(b, "column", "stacked", legend=False)
    ch.add_series({"name": "Nền", "values": r["Nền"], "categories": r["cats"], "fill": {"none": True}, "border": {"none": True},
                   "gap": 80})
    # Labels inside the bar ends: white on the two dark totals, dark ink on the steps.
    labels = text_labels(r["nhãn@cells"], set(range(len(cats))), C["ink2"], "inside_end", 9)
    for i, lab in enumerate(labels["custom"]):
        if i in (0, len(cats) - 1):
            lab["font"] = {"name": FONT, "size": 9, "color": "#FFFFFF", "bold": True}
    ch.add_series(bars("Biên EBIT", r["Cột"], C["total"], r["cats"],
                       points=[{"fill": {"color": c}, "border": {"none": True}} for c in colors], data_labels=labels))
    ch.set_y_axis(value_axis("0%", min=0))
    ch.set_x_axis(category_axis())
    charts["bridge"] = ch

    # 5. margins by quarter
    gm = [k("gross_margin", p) for p in Q]
    em = [k("ebit_margin", p) for p in Q]
    r = d.block("Biên gộp và biên EBIT, từng quý", qcats, {"Biên gộp": gm, "Biên EBIT": em,
                                                           "nhãn gộp": [pct(x) for x in gm], "nhãn EBIT": [pct(x) for x in em]})
    ch = base_chart(b, "line")
    ch.add_series(line("Biên lợi nhuận gộp", r["Biên gộp"], C["total"], r["cats"],
                       data_labels=text_labels(r["nhãn gộp@cells"], keep, position="above")))
    ch.add_series(line("Biên EBIT (non-GAAP)", r["Biên EBIT"], C["margin"], r["cats"],
                       data_labels=text_labels(r["nhãn EBIT@cells"], keep, position="below")))
    ch.set_y_axis(value_axis(f"0%;{MINUS}0%;0%", min=-0.2, max=0.6, major_unit=0.2))
    ch.set_x_axis(category_axis(label_position="low"))
    charts["marginsQ"] = ch

    # 5. inventories and days of inventory by quarter
    inv = [v("inventories", p) / 1000 for p in Q]
    hi = inv.index(max(inv))
    r = d.block("Tồn kho cuối quý", qcats, {"Tồn kho (tỷ USD)": inv, "nhãn": [num(x, 1) for x in inv]})
    ch = base_chart(b, "column", legend=False)
    ch.add_series(bars("Hàng tồn kho", r["Tồn kho (tỷ USD)"], C["total"], r["cats"], gap=70,
                       data_labels=text_labels(r["nhãn@cells"], {0, hi, len(Q) - 1}, position="outside_end")))
    ch.set_y_axis(value_axis("0", min=0))
    ch.set_x_axis(category_axis())
    charts["inventoryQ"] = ch

    days = [k("days_inventory", p) for p in Q]
    hi = days.index(max(days))
    r = d.block("Số ngày tồn kho, từng quý", qcats, {"Số ngày": days, "nhãn": [num(x, 0) for x in days]})
    ch = base_chart(b, "line", legend=False)
    ch.add_series(line("Số ngày tồn kho", r["Số ngày"], C["total"], r["cats"],
                       data_labels=text_labels(r["nhãn@cells"], {0, hi, len(Q) - 1}, position="above")))
    ch.set_y_axis(value_axis("0", min=60, max=180, major_unit=20))
    ch.set_x_axis(category_axis())
    charts["daysQ"] = ch
    return charts


def place(ws, r, c, chart, width, height):
    chart.set_size({"width": width, "height": height})
    ws.insert_chart(r, c, chart, {"x_offset": 0, "y_offset": 0, "object_position": 1})


# ------------------------------------------------------------------ dashboard sheets
class Page:
    def __init__(self, b, ws, name):
        self.b, self.ws, self.name = b, ws, name
        ws.hide_gridlines(2)
        ws.set_column(0, 0, 2.71)
        ws.set_column(1, 12, GRID_W)
        ws.set_column(13, 13, 2.71)
        ws.set_default_row(15)
        ws.set_landscape()
        ws.set_paper(9)
        ws.set_print_scale(PRINT_SCALE)
        ws.set_margins(0.3, 0.3, 0.4, 0.4)
        self.starts = [0]

    def paginate(self, end):
        """Manual page breaks between blocks (chart rows, findings) so that none is split across pages."""
        size = lambda r: self.ws.row_sizes.get(r, [self.ws.default_row_height])[0] * 0.75   # XlsxWriter keeps pixels
        breaks, used = [], 0
        for a, z in zip(self.starts, self.starts[1:] + [end]):
            h = sum(size(r) for r in range(a, z))
            if used and used + h > PAGE_ROWS_PT:
                breaks.append(a)
                used = 0
            used += h
        self.ws.set_h_pagebreaks(breaks)

    def text(self, r, c1, c2, value, fmt, height=None):
        if height:
            self.ws.set_row(r, height)
        if c1 == c2:
            self.ws.write(r, c1, value, fmt)
        else:
            self.ws.merge_range(r, c1, r, c2, value, fmt)

    def header(self):
        b, ws, t = self.b, self.ws, self.b.texts
        ws.set_row(0, 9)
        ws.merge_range(1, 1, 2, 9, t["title"], b.f(font_name=SERIF, font_size=20, bold=True, font_color=C["ink"], text_wrap=True,
                                                    valign="bottom"))
        ws.set_row(1, 27)
        ws.set_row(2, 27)
        ws.merge_range(3, 1, 4, 9, t["subtitle"], b.f(font_size=10.5, font_color=C["ink2"], text_wrap=True))
        ws.set_row(3, 17)
        ws.set_row(4, 17)
        link = b.f(font_color=C["link"], underline=1, align="right")
        ws.write_url(2, 12, DASHBOARD_URL + b.d["meta"]["workbook"], link, "Tải bộ dữ liệu Excel")
        ws.write_url(3, 12, f"internal:'{SHEET_SRC}'!A1", link, "Nguồn và phương pháp")
        ws.write_url(4, 12, DASHBOARD_URL, link, "Mở dashboard web")
        ws.set_row(5, 9)
        # Sheet links in one band, like the section links of the web page.
        band = b.f(bg_color=C["band"], top=1, bottom=1, top_color=C["rule"], bottom_color=C["rule"], valign="vcenter")
        ws.set_row(6, 27)
        for c in range(1, 13):
            ws.write_blank(6, c, None, band)
        active = b.f(bold=True, font_size=10.5, font_color=C["ink"], bg_color=C["band"], bottom=5, bottom_color=C["ink"],
                     top=1, top_color=C["rule"], valign="vcenter", align="center")
        idle = b.f(font_size=10.5, font_color=C["muted"], bg_color=C["band"], top=1, bottom=1, top_color=C["rule"],
                   bottom_color=C["rule"], valign="vcenter", align="center")
        spans = [2, 2, 3, 2, 3]
        c1 = 1
        for sheet, span in zip(SHEETS, spans):
            if sheet == self.name:
                ws.merge_range(6, c1, 6, c1 + span - 1, sheet, active)
            else:
                ws.merge_range(6, c1, 6, c1 + span - 1, "", idle)
                ws.write_url(6, c1, f"internal:'{sheet}'!A1", idle, sheet)
            c1 += span
        ws.set_row(7, 12)
        return 8

    def section(self, r, sid, note=None):
        b, ws = self.b, self.ws
        kicker, question = b.texts["sections"][sid]
        for c in range(1, 13):
            ws.write_blank(r, c, None, b.f(top=1, top_color=C["rule"]))
        ws.set_row(r, 9)
        self.text(r + 1, 1, 12, kicker, b.f(bold=True, font_size=9.5, font_color=C["muted"]), 16)
        self.text(r + 2, 1, 12, question, b.f(font_name=SERIF, font_size=16, bold=True, font_color=C["ink"]), 26)
        r += 3
        if note:
            ws.merge_range(r, 1, r + 1, 9, glue(note), b.f(font_size=11, font_color=C["ink2"], text_wrap=True))
            ws.set_row(r, 17)
            ws.set_row(r + 1, 17)
            r += 2
        ws.set_row(r, 8)
        return r + 1

    def figure(self, r, c1, c2, key, chart, height, sub=None, source=None, legend_row=False):
        b, ws = self.b, self.ws
        title, fig_sub = b.texts["figs"][key]
        self.text(r, c1, c2, title, b.f(bold=True, font_size=11.5, font_color=C["ink"]), 19)
        ws.merge_range(r + 1, c1, r + 1, c2, sub or fig_sub, b.f(font_size=9, font_color=C["muted"], text_wrap=True))
        ws.set_row(r + 1, 26)
        rows = math.ceil(height / 20)
        for k in range(rows):
            ws.set_row(r + 2 + k, 15)
        place(ws, r + 2, c1, chart, 100 * (c2 - c1 + 1) - 10, height)
        src = r + 2 + math.ceil(height / 20)
        ws.merge_range(src, c1, src, c2, source or b.sources[key], b.f(font_size=8, font_color=C["muted"], text_wrap=True))
        ws.set_row(src, 34 if c2 - c1 < 6 else 24)
        return src + 2


def write_findings(p, r):
    b, ws = p.b, p.ws
    p.text(r, 1, 12, b.texts["findings_title"], b.f(font_name=SERIF, font_size=16, bold=True, font_color=C["ink"]), 26)
    p.text(r + 1, 1, 12, b.texts["findings_note"], b.f(font_size=10, font_color=C["ink2"]), 18)
    r += 2
    findings = b.d["findings"]
    for row in range(2):
        p.starts.append(r)
        for col in range(3):
            f = findings[row * 3 + col]
            c1 = 1 + col * 4
            c2 = c1 + 3
            ws.merge_range(r, c1, r, c2 - 0, glue(f["label"]),
                           b.f(bold=True, font_size=9.5, font_color=C["muted"], top=1, top_color=C["rule"], valign="bottom"))
            ws.merge_range(r + 1, c1, r + 1, c2, glue(f["value"]), b.f(bold=True, font_size=22, font_color=C["ink"], valign="vcenter"))
            ws.merge_range(r + 2, c1, r + 5, c2, glue(f["text"]), b.f(font_size=10, font_color=C["ink2"], text_wrap=True))
            ws.merge_range(r + 6, c1, r + 6, c2, glue(f["source"]), b.f(font_size=8, font_color=C["muted"], text_wrap=True))
        ws.set_row(r, 22)
        ws.set_row(r + 1, 34)
        for k in range(2, 6):
            ws.set_row(r + k, 17)
        ws.set_row(r + 6, 24)
        ws.set_row(r + 7, 10)
        r += 8
    return r


def build_overview(b, ws, charts):
    p = Page(b, ws, SHEET_OVERVIEW)
    r = p.header()
    r = write_findings(p, r)
    p.starts.append(r)
    r = p.section(r, "kenh")
    a, z = b.A[0], b.A[-1]
    r1 = p.figure(r, 1, 6, "share", charts["share"], 250,
                  source=b.sources["share"] + glue(" Báo cáo: " + ", ".join(f"10-K {y}" for y in b.A) + "."))
    p.figure(r, 7, 12, "change", charts["change"], 250, sub=b.subs["change-sub"](a, z),
             source=b.sources["change"] + glue(f" Báo cáo: 10-K {a}, 10-K {z}."))
    p.paginate(r1)
    ws.activate()
    return p


def build_quarters(b, ws, charts):
    p = Page(b, ws, SHEET_Q)
    r = p.header()
    r0 = p.section(r, "quy", b.d["notes"]["quy"])
    p.starts.append(r0)
    r = p.figure(r0, 1, 6, "shareQ", charts["shareQ"], 280)
    p.figure(r0, 7, 12, "channelQ", charts["channelQ"], 280)
    p.starts.append(r)
    r = p.section(r, "digital")
    r2 = p.figure(r, 1, 6, "digital", charts["digital"], 250, source=b.sources["digital"] + glue(
        " Báo cáo: " + ", ".join(f"10-K {y}" for y in b.A) + "."))
    p.figure(r, 7, 12, "digitalQ", charts["digitalQ"], 250)
    p.paginate(r2)
    return p


def build_profit(b, ws, charts):
    p = Page(b, ws, SHEET_PROFIT)
    r = p.header()
    a, z = b.A[0], b.A[-1]
    r = p.section(r, "khu-vuc")
    p.starts.append(r)
    two = glue(f" Báo cáo: 10-K {a}, 10-K {z}.")
    r_next = p.figure(r, 1, 6, "regions", charts["regions"], 260, source=b.sources["regions"] + two)
    p.figure(r, 7, 12, "regionsEbit", charts["regionsEbit"], 260, source=b.sources["regionsEbit"] + two)
    r = r_next
    p.starts.append(r)
    r = p.section(r, "loi-nhuan")
    r_next = p.figure(r, 1, 6, "bridge", charts["bridge"], 280, sub=b.subs["bridge-sub"](a, z), source=b.sources["bridge"] + two)
    p.figure(r, 7, 12, "marginsQ", charts["marginsQ"], 280)
    r = r_next
    p.starts.append(r)
    r_next = p.figure(r, 1, 6, "inventoryQ", charts["inventoryQ"], 230)
    p.figure(r, 7, 12, "daysQ", charts["daysQ"], 230)
    r = r_next
    ws = p.ws
    ws.merge_range(r, 1, r + 1, 9, glue(b.d["notes"]["ton-kho"]), b.f(font_size=11, font_color=C["ink2"], text_wrap=True))
    ws.set_row(r, 17)
    ws.set_row(r + 1, 17)
    r += 3
    p.starts.append(r)
    p.text(r, 1, 12, b.texts["notes_title"], b.f(bold=True, font_size=12, font_color=C["ink"]), 20)
    items = b.texts["notes"]
    for i, (dt, dd) in enumerate(items):
        row, col = divmod(i, 3)
        c1 = 1 + 4 * col
        top = r + 1 + row * 6
        ws.merge_range(top, c1, top + 5, c1 + 3, "", b.f(text_wrap=True))
        ws.write_rich_string(top, c1, b.f(bold=True, font_color=C["ink"]), dt + "\n", b.f(font_color=C["ink2"]), dd,
                             b.f(text_wrap=True))
    for k in range(1, 13):
        ws.set_row(r + k, 17)
    end = r + 13
    p.paginate(end)
    return p


def write_table(p, r, title, periods, rows, two_level=False, breaks=None, break_before=()):
    """One table: a figure row per metric and, under each reported figure, a small row with its page (a link to the
    report), so every figure in a column ends at the same place. A printed page ends before each group named in
    `break_before` (its row is added to `breaks`), and the column headers are repeated there, so a figure is never
    separated from its page, nor a page from its headers."""
    b, ws = p.b, p.ws
    n = len(periods)
    p.text(r, 1, 1 + 2 + n, glue(title), b.f(bold=True, font_size=12, font_color=C["ink"]), 22)
    r += 1
    head = b.f(bold=True, font_color=C["ink"], bottom=1, bottom_color=C["ink2"], valign="bottom")
    head_r = b.f(bold=True, font_color=C["ink"], bottom=1, bottom_color=C["ink2"], valign="bottom", align="right")

    def header_rows(r):
        if two_level:
            for i, fy in enumerate(b.A):
                ws.merge_range(r, 3 + 4 * i, r, 6 + 4 * i, fy, b.f(bold=True, font_color=C["ink2"], align="center",
                                                                  bottom=1, bottom_color=C["rule"]))
            r += 1
        ws.write(r, 1, "Chỉ tiêu", head)
        ws.write(r, 2, "Đơn vị", head)
        for j, p_ in enumerate(periods):
            ws.write(r, 3 + j, ("Q" + p_[7]) if two_level else p_, head_r)
        ws.set_row(r, 20)
        return r + 1

    r = header_rows(r)
    for spec in rows:
        if spec[0] == "g":
            if spec[1] in break_before:
                breaks.append(r)
                r = header_rows(r)
            ws.merge_range(r, 1, r, 2 + n, glue(spec[1]), b.f(bold=True, font_color=C["ink"], bg_color=C["band"], valign="vcenter"))
            ws.set_row(r, 20)
            r += 1
            continue
        kind, mid, label, unit, fmt, total = spec
        reported = kind == "obs"
        rule = {"bottom": 1, "bottom_color": C["grid"]}
        top = {"top": 1, "top_color": C["ink2"]} if total else {}
        border = top if reported else {**top, **rule}
        weight = {"bold": True} if total else {}
        ws.write(r, 1, glue(label), b.f(font_color=C["ink"], valign="vcenter", text_wrap=True, **border, **weight))
        ws.write(r, 2, unit, b.f(font_color=C["muted"], valign="vcenter", **border))
        pages = []
        for j, per in enumerate(periods):
            italic, text, page, link = True, "—", "", ""
            if kind == "kpi":
                v = b.kpi(mid, per)
                text = "—" if v is None else fmt(v)
            else:
                o = b.ob(mid, per)
                if o.get("value") is None:
                    text = "Không công bố"
                elif o.get("method") == "q4_derived":
                    text = ("≈" + num(o["value"], 1)) if o.get("approx") else fmt(o["value"])
                else:
                    italic, text = False, fmt(o["value"])
                    mark = "† " if o.get("status") == STATUS_PENDING else ""
                    page = mark + (f"tr. {o['page']}" if o.get("page") else "")
                    link = b.d["filings"][o["source"]]["url"] if o.get("page") else ""
            missing = text == "Không công bố"
            cell = b.f(italic=italic and not missing, align="right", valign="vcenter", **border, **weight,
                       font_color=C["muted"] if missing else C["ink"] if not italic else C["ink2"],
                       font_size=8.5 if missing else 10)
            ws.write(r, 3 + j, text, cell)
            pages.append((page, link))
        ws.set_row(r, 30 if len(label) > 34 else 20)
        r += 1
        if not reported:
            continue
        small = b.f(font_size=8, font_color=C["muted"], valign="top", **rule)
        ws.write_blank(r, 1, None, small)
        ws.write_blank(r, 2, None, small)
        ref = b.f(font_size=8, font_color=C["faint"], align="right", valign="top", **rule)
        for j, (page, link) in enumerate(pages):
            if link:
                ws.write_url(r, 3 + j, link, ref, page)
            else:
                ws.write(r, 3 + j, page, ref)
        ws.set_row(r, 13)
        r += 1
    return r + 1


def build_tables(b, ws):
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2.71)
    ws.set_column(1, 1, 44)
    ws.set_column(2, 2, 10)
    ws.set_column(3, 14, 10.5)
    ws.set_landscape()
    ws.set_paper(9)
    # A fixed scale (Excel ignores manual page breaks under fit-to-page): the twelve quarter columns fit A4 landscape.
    ws.set_print_scale(66)
    ws.set_margins(0.3, 0.3, 0.4, 0.4)
    ws.write(1, 1, SHEET_TABLES, b.f(font_name=SERIF, font_size=18, bold=True, font_color=C["ink"]))
    ws.set_row(1, 28)
    ws.write(2, 1, glue("Số NIKE công bố có số trang ở dòng nhỏ ngay dưới (bấm để mở báo cáo trên SEC). Số in nghiêng là số "
                        "nhóm tự tính, gồm cả số quý 4 (cả năm trừ 9 tháng). ≈ là ước tính từ số đã làm tròn. Số âm để trong "
                        "ngoặc đơn."), b.f(font_size=9, font_color=C["muted"]))
    p = Page.__new__(Page)
    p.b, p.ws = b, ws
    # Each table starts on a new printed page and takes two: channels, then the income statement and ratios; two regions
    # per page.
    breaks, finance = [], {FINANCE_ROWS[0][1]}
    regions = region_rows(b.d["regions"])
    r = write_table(p, 4, "Theo năm, FY2019–FY2021", b.A, CHANNEL_ROWS + FINANCE_ROWS, breaks=breaks, break_before=finance)
    breaks.append(r + 1)
    r = write_table(p, r + 1, "Theo quý, Q1 FY2019 – Q4 FY2021", b.Q, CHANNEL_ROWS + FINANCE_ROWS, two_level=True,
                    breaks=breaks, break_before=finance)
    breaks.append(r + 1)
    write_table(p, r + 1, "Bốn khu vực của NIKE Brand, FY2019–FY2021", b.A, regions, breaks=breaks,
                break_before={[spec[1] for spec in regions if spec[0] == "g"][2]})
    ws.set_h_pagebreaks(breaks)
    ws.freeze_panes(0, 3)


def build_sources(b, ws):
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 2.71)
    ws.set_column(1, 1, 62)
    ws.set_column(2, 2, 4)
    ws.set_column(3, 3, 62)
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    t = b.texts
    ws.write(1, 1, SHEET_SRC, b.f(font_name=SERIF, font_size=20, bold=True, font_color=C["ink"]))
    ws.set_row(1, 32)
    h = b.f(bold=True, font_size=11, font_color=C["ink"], bottom=1, bottom_color=C["rule"])
    kicker, question = t["sections"]["nike-giai-thich"]
    ws.write(3, 1, question, h)
    ws.merge_range(4, 1, 4, 3, t["quotes_note"], b.f(font_size=9.5, font_color=C["muted"]))
    r = 5
    link = b.f(font_color=C["link"], underline=1)
    quotes = b.d["quotes"]
    for i in range(0, len(quotes), 2):
        height = 0
        for q, col in zip(quotes[i:i + 2], (1, 3)):
            ws.write_rich_string(r, col, b.f(bold=True, font_size=9, font_color=C["muted"]), q["topic"] + "\n",
                                 b.f(italic=True, font_color=C["ink"]), "“" + q["text"] + "”\n",
                                 b.f(font_color=C["ink2"]), glue(q["vi"]) + "\n",
                                 b.f(font_size=8.5, font_color=C["muted"]), glue(f"NIKE, {q['filing']}, tr. {q['page']}"),
                                 b.f(text_wrap=True))
            # About 78 characters of the English quote and 88 of the Vietnamese line fit a 62-wide column.
            lines = math.ceil(len(q["text"]) / 78) + math.ceil(len(q["vi"]) / 88) + 2
            height = max(height, 13.5 * lines + 8)
        ws.set_row(r, height)
        r += 1
    r += 1
    ws.write(r, 1, glue("Báo cáo nộp SEC (Form 10-K và 10-Q)"), h)
    groups = [("Báo cáo năm (10-K)", lambda f: f["form"] == "10-K" and f["role"] == "scope"),
              ("Báo cáo quý (10-Q)", lambda f: f["form"] == "10-Q" and f["role"] == "scope"),
              ("Chỉ dùng để xem số có bị trình bày lại không", lambda f: f["role"] == "check")]
    rr = r + 1
    for title, keep in groups:
        ws.write(rr, 1, glue(title), b.f(bold=True, font_size=9.5, font_color=C["ink2"]))
        rr += 1
        for fid, f in b.d["filings"].items():
            if keep(f):
                dd, mm, yy = f["filed"][8:10], f["filed"][5:7], f["filed"][:4]
                ws.write_url(rr, 1, f["url"], link, glue(f"{fid}, nộp ngày {dd}/{mm}/{yy}, số {f['accession']}"))
                rr += 1
    ws.write(r, 3, "Phương pháp", h)
    for i, item in enumerate(t["method"]):
        ws.write(r + 1 + i, 3, "–  " + item, b.f(text_wrap=True))
        ws.set_row(r + 1 + i, max(30, 15 * math.ceil(len(item) / 70) + 3))
    r = max(rr, r + 1 + len(t["method"])) + 1
    ws.write(r, 1, "Tình trạng dữ liệu", h)
    c = b.d["meta"]["status_counts"]
    total = sum(c.values())
    missing = [p for m in b.d["observations"].values() for p, o in m.items() if o["value"] is None]
    all_q4 = missing and all(p.endswith("Q4") for p in missing)
    parts = [(c.get("Khớp nguồn chính thức", 0), " khớp nguồn chính thức"), (c.get("Đã xác minh (người)", 0), " đã được nhóm kiểm tra trên báo cáo"),
             (c.get(STATUS_PENDING, 0), " chưa được kiểm tra (†)"),
             (c.get("Thiếu / không công bố", 0), " là các tỷ lệ % NIKE không nêu cho riêng quý 4" if all_q4 else " NIKE không công bố")]
    extracted = b.d["meta"]["extracted_on"]
    status = (f"{total} số liệu báo cáo: " + ", ".join(f"{n}{label}" for n, label in parts if n) +
              f". Dữ liệu trích xuất ngày {extracted[8:10]}/{extracted[5:7]}/{extracted[:4]}.")
    ws.merge_range(r + 1, 1, r + 2, 1, glue(status), b.f(text_wrap=True))
    ws.write(r + 3, 1, t["fine"], b.f(font_size=9, font_color=C["muted"]))
    ws.write(r, 3, "Liên kết", h)
    ws.write_url(r + 1, 3, DASHBOARD_URL, link, "Dashboard web")
    ws.write_url(r + 2, 3, DASHBOARD_URL + b.d["meta"]["workbook"], link, "Bộ dữ liệu Excel (đầy đủ nguồn từng số)")
    pages = b.d["meta"]["evidence"]
    n = re.search(r"(\d+)-so", pages).group(1)
    ws.write_url(r + 3, 3, DASHBOARD_URL + pages, link, glue(f"Đối chiếu {n} số trên trang 10-Q (ảnh chụp)"))


def build(path=OUT, save_with_excel=True):
    data = json.loads(SITE_DATA.read_text(encoding="utf-8"))
    b = Book(path, data)
    sheets = {n: b.wb.add_worksheet(n) for n in SHEETS + [SHEET_DATA]}
    d = DataSheet(b, sheets[SHEET_DATA])
    charts = build_charts(b, d)
    build_overview(b, sheets[SHEET_OVERVIEW], charts)
    build_quarters(b, sheets[SHEET_Q], charts)
    build_profit(b, sheets[SHEET_PROFIT], charts)
    build_tables(b, sheets[SHEET_TABLES])
    build_sources(b, sheets[SHEET_SRC])
    sheets[SHEET_DATA].hide()
    b.wb.close()
    if save_with_excel:
        run_save_script(path)
    set_author(path, git_user_name())
    return b


def set_author(path, name):
    """Author and 'last modified by' in the file properties. Excel stamps its own user name there when the save
    script saves the workbook; the builder's name belongs there instead."""
    if not name:
        return
    tmp = Path(path).with_suffix(".tmp.xlsx")
    value = escape(name)
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                xml = data.decode("utf-8")
                for tag in ("dc:creator", "cp:lastModifiedBy"):
                    xml, n = re.subn(rf"<{tag}\s*/>|<{tag}>.*?</{tag}>", f"<{tag}>{value}</{tag}>", xml, flags=re.S)
                    if not n:
                        xml = xml.replace("</cp:coreProperties>", f"<{tag}>{value}</{tag}></cp:coreProperties>")
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(path)


def run_save_script(path):
    if sys.platform != "win32":
        print("Not saved with Excel: needs Windows with Excel (pipeline/excel_save.ps1).")
        return
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(SAVE_SCRIPT), "-Path", str(path)]
    subprocess.run(cmd, check=True)


def main():
    build()
    print(f"Excel dashboard written: {OUT.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
