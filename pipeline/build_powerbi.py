"""Build the Power BI version of the dashboard as a Power BI Project (PBIP).

Usage:  python -m pipeline.build_powerbi [output folder]      (default: powerbi/)

Reads site/data/nike-financials.json (written by build_site_data from the validated
workbook) and the wording of site/index.html and site/assets/app.js, and writes:

  NIKE_Dashboard.pbip                 open this in Power BI Desktop
  NIKE_Dashboard.SemanticModel/       TMDL; the data is embedded in Power Query (no file paths)
  NIKE_Dashboard.Report/              PBIR pages and visuals

Power BI Desktop turns the project into the .pbix handed to the team: open the .pbip,
Refresh, then File > Save as > Power BI file. Figures come from the JSON unchanged; the
only derived values are USD million to billion and "rest of NIKE Direct" (NIKE Direct
minus Digital), which draws Digital inside NIKE Direct as a stacked column.
"""
import glob
import json
import re
import shutil
import subprocess
import sys
import uuid
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from lxml import html as lxml_html

ROOT = Path(__file__).resolve().parent.parent
SITE = ROOT / "site"
DATA = SITE / "data" / "nike-financials.json"
OUT = ROOT / "powerbi"
NAME = "NIKE_Dashboard"
THEME = "Fluent2-CY26SU10"
THEME_SOURCES = [
    "C:/Program Files/WindowsApps/Microsoft.MicrosoftPowerBIDesktop_*/bin/WebView2Resources/minerva/sharedresources/BaseThemes",
    "C:/Program Files/Microsoft Power BI Desktop/bin/WebView2Resources/minerva/sharedresources/BaseThemes",
]
SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/"
VC_SCHEMA = SCHEMA + "item/report/definition/visualContainer/2.4.0/schema.json"
PAGE_SCHEMA = SCHEMA + "item/report/definition/page/2.0.0/schema.json"
STATUS_PENDING = "Chờ người kiểm tra"
MINUS = "\u2212"

# Palette and type, as in site/assets/app.js and styles.css
C = dict(direct="#2a78d6", wholesale="#eb6834", gbd="#a3a9b2", directLight="#86b6ef", digital="#1c5cab",
         total="#475467", overhead="#6250d6", demand="#1baf7a",
         ink="#15191e", ink2="#3b4350", muted="#5b6472", faint="#6f7784", grid="#e6e9ee", axis="#c3c9d1", rule="#dde1e6")
SANS, SERIF = "Segoe UI", "Cambria"
PAGE_W, MARGIN = 1280, 40

# Table rows mirror TABLES in site/assets/app.js: (kind, id, label, unit, format).
TABLES = {
    "kenh": [
        ("g", "Doanh thu NIKE Brand"),
        ("obs", "nike_direct_revenue", "NIKE Direct", "triệu USD", "amount"),
        ("obs", "nike_brand_wholesale_revenue", "Bán buôn (Wholesale)", "triệu USD", "amount"),
        ("obs", "global_brand_divisions_revenue", "Global Brand Divisions", "triệu USD", "amount"),
        ("obs", "nike_brand_revenue", "Tổng doanh thu NIKE Brand", "triệu USD", "amount"),
        ("g", "Cơ cấu theo kênh"),
        ("kpi", "direct_share", "Tỷ trọng NIKE Direct", "%", "pct1"),
        ("kpi", "wholesale_share", "Tỷ trọng bán buôn", "%", "pct1"),
        ("obs", "nike_direct_share_stated", "Tỷ trọng NIKE Direct do NIKE nêu (xấp xỉ)", "%", "pct0"),
        ("g", "Tăng trưởng so với năm trước"),
        ("kpi", "direct_growth", "NIKE Direct, theo báo cáo", "%", "spct1"),
        ("obs", "nike_direct_pct_change_cn", "NIKE Direct, loại trừ tỷ giá (non-GAAP)", "%", "spct0"),
        ("kpi", "wholesale_growth", "Bán buôn, theo báo cáo", "%", "spct1"),
        ("obs", "nike_brand_wholesale_pct_change_cn", "Bán buôn, loại trừ tỷ giá (non-GAAP)", "%", "spct0"),
        ("kpi", "nike_brand_growth", "NIKE Brand, theo báo cáo", "%", "spct1"),
        ("g", "NIKE Brand Digital (một phần của NIKE Direct)"),
        ("obs", "nike_brand_digital_sales", "Doanh số Digital", "tỷ USD", "num1"),
        ("kpi", "digital_share_of_direct", "Digital trong NIKE Direct (xấp xỉ)", "%", "pct0"),
    ],
    "ketqua": [
        ("g", "Báo cáo kết quả kinh doanh"),
        ("obs", "revenues", "Doanh thu", "triệu USD", "amount"),
        ("obs", "cost_of_sales", "Giá vốn hàng bán", "triệu USD", "amount"),
        ("obs", "gross_profit", "Lợi nhuận gộp", "triệu USD", "amount"),
        ("obs", "demand_creation_expense", "Chi phí tạo nhu cầu", "triệu USD", "amount"),
        ("obs", "operating_overhead_expense", "Chi phí vận hành chung", "triệu USD", "amount"),
        ("obs", "total_selling_admin_expense", "Tổng chi phí bán hàng và quản lý", "triệu USD", "amount"),
        ("obs", "interest_expense_net", "Chi phí (thu nhập) lãi thuần", "triệu USD", "amount"),
        ("obs", "other_expense_net", "Chi phí (thu nhập) khác thuần", "triệu USD", "amount"),
        ("obs", "income_before_taxes", "Lợi nhuận trước thuế", "triệu USD", "amount"),
        ("obs", "income_tax_expense", "Chi phí thuế thu nhập", "triệu USD", "amount"),
        ("obs", "net_income", "Lợi nhuận sau thuế", "triệu USD", "amount"),
        ("g", "Tỷ lệ trên doanh thu và tăng trưởng"),
        ("kpi", "revenue_growth", "Tăng trưởng doanh thu, theo báo cáo", "%", "spct1"),
        ("obs", "revenues_pct_change_cn", "Tăng trưởng doanh thu, loại trừ tỷ giá (non-GAAP)", "%", "spct0"),
        ("kpi", "gross_margin", "Biên lợi nhuận gộp", "%", "pct1"),
        ("kpi", "sga_ratio", "Chi phí bán hàng và quản lý / doanh thu", "%", "pct1"),
        ("kpi", "demand_creation_ratio", "Trong đó: tạo nhu cầu / doanh thu", "%", "pct1"),
        ("kpi", "operating_overhead_ratio", "Trong đó: vận hành chung / doanh thu", "%", "pct1"),
        ("kpi", "net_income_margin", "Biên lợi nhuận sau thuế", "%", "pct1"),
        ("g", "EBIT (non-GAAP, định nghĩa của NIKE)"),
        ("obs", "ebit_reported", "EBIT do NIKE công bố", "triệu USD", "amount"),
        ("kpi", "ebit_margin", "Biên EBIT", "%", "pct1"),
        ("g", "Hàng tồn kho"),
        ("obs", "inventories", "Hàng tồn kho tại 31/5", "triệu USD", "amount"),
        ("kpi", "inventory_growth", "Thay đổi so với đầu kỳ", "%", "spct1"),
        ("kpi", "inventory_turnover", "Vòng quay hàng tồn kho", "lần", "num2"),
        ("kpi", "days_inventory", "Số ngày tồn kho bình quân", "ngày", "num1"),
    ],
}
TABLE_NAMES = {"kenh": "BangKenh", "ketqua": "BangKetQua"}


# ------------------------------------------------------------------ vi-VN formatting (as app.js)
def nf(v, d):
    q = Decimal(abs(v)).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)
    return f"{q:,.{d}f}".replace(",", "§").replace(".", ",").replace("§", ".")


def num(v, d=0):
    return "—" if v is None else (MINUS if v < 0 else "") + nf(v, d)


FORMATS = {
    "amount": lambda v: "(" + nf(v, 0) + ")" if v < 0 else nf(v, 0),
    "pct1": lambda v: num(v * 100, 1) + "%",
    "pct0": lambda v: num(v * 100, 0) + "%",
    "spct1": lambda v: ("+" if v > 0 else MINUS if v < 0 else "") + nf(v * 100, 1) + "%",
    "spct0": lambda v: ("+" if v > 0 else MINUS if v < 0 else "") + nf(v * 100, 0) + "%",
    "num1": lambda v: num(v, 1),
    "num2": lambda v: num(v, 2),
}


def vn_date(iso):
    y, m, d = iso.split("-")
    return f"{d}/{m}/{y}"


def table_cells(row, y, obs, kpis, filings):
    """Figure, page reference and 10-K link of one table cell, as renderTable() in app.js shows them."""
    kind, mid, _, _, fmt = row
    if kind == "kpi":
        v = kpis[mid]["values"][str(y)]
        return ("—" if v is None else FORMATS[fmt](v)), "", None
    o = obs[mid].get(str(y))
    if not o or o["value"] is None:
        return "Không công bố", "", None
    mark = "†" if o["status"] == STATUS_PENDING else ""
    if not o.get("page"):
        return FORMATS[fmt](o["value"]), mark, None
    return FORMATS[fmt](o["value"]), " ".join(filter(None, [mark, f"tr. {o['page']}"])), filings[str(y)]["url"]


# ------------------------------------------------------------------ wording from the web dashboard
class Wording:
    def __init__(self):
        doc = lxml_html.fromstring((SITE / "index.html").read_text(encoding="utf-8"))
        js = (SITE / "assets" / "app.js").read_text(encoding="utf-8")

        def one(xp):
            els = doc.xpath(xp)
            if not els:
                raise SystemExit(f"index.html: nothing matches {xp}")
            return els[0]

        def text(xp):
            return " ".join(one(xp).text_content().split())

        self.title = text("//h1")
        self.subtitle = text("//p[@class='masthead-sub']")
        self.tab = {v: text(f"//button[@id='tab-{v}']") for v in ("channels", "performance")}
        self.question = {v: text(f"//section[@id='view-{v}']//h2") for v in ("channels", "performance")}
        self.summary_label = text("//p[@class='summary-label']")
        self.fig = {}
        for f in doc.xpath("//figure[@id]"):
            sub = f.xpath(".//p[contains(@class,'figure-sub')]")
            self.fig[f.get("id")] = (" ".join(f.xpath(".//h3")[0].text_content().split()),
                                     " ".join(sub[0].text_content().split()) if sub else "")
        self.notes = {}
        for v in ("channels", "performance"):
            aside = one(f"//section[@id='view-{v}']//aside")
            self.notes[v] = (" ".join(aside.xpath(".//h3")[0].text_content().split()),
                             [(" ".join(dt.text_content().split()), " ".join(dd.text_content().split()))
                              for dt, dd in zip(aside.xpath(".//dt"), aside.xpath(".//dd"))])
        self.table_title, self.table_key, self.pending_key = {}, {}, ""
        for v in ("channels", "performance"):
            self.table_title[v] = text(f"//h3[@id='table-{v}-title']")
            key = one(f"//section[@id='view-{v}']//p[@class='table-key']")
            pending = key.xpath(".//span[@class='pending-key']")
            if pending:
                self.pending_key = " ".join(pending[0].text_content().split())
                pending[0].drop_tree()
            self.table_key[v] = " ".join(key.text_content().split())
        self.footer_title = text("//footer//h2")
        self.footer_heads = [" ".join(h.text_content().split()) for h in doc.xpath("//footer//h3")]
        self.method = [" ".join(li.text_content().split()) for li in doc.xpath("//footer//ul[@class='plain-list']/li")]
        self.fine = text("//footer//p[@class='fine']")

        self.sources = dict(re.findall(r"sourceLine\('(\w+)', '([^']*)'", js))
        m = re.search(r"sourceLine\('growth', cn \? '([^']*)'\s*:\s*'([^']*)'", js)
        if not m:
            raise SystemExit("app.js: growth source line not found")
        self.sources["growth_cn"], self.sources["growth"] = m.group(1), m.group(2)
        self.legends = {k: re.findall(r"name: '([^']*)', label: '([^']*)'", body)
                        for k, body in re.findall(r"(\w+): function \(\) \{\s*return \[(.*?)\];", js, re.S)}
        self.legends["growth"] = self.legends["channelRevenue"]
        m = re.search(r"getElementById\('growth-sub'\)\.textContent = cn\s*\?\s*'([^']*)'\s*:\s*'([^']*)'", js)
        if not m:
            raise SystemExit("app.js: growth subtitle not found")
        self.growth_sub_cn, self.growth_sub = m.group(1), m.group(2)
        for k in ("share", "channelRevenue", "digital", "revenue", "grossMargin", "sga", "ebit", "inventory", "days"):
            if k not in self.sources:
                raise SystemExit(f"app.js: source line for {k} not found")

    def legend(self, chart, name):
        return dict(self.legends[chart])[name]


# ------------------------------------------------------------------ semantic model (TMDL)
def tmdl_name(n):
    return n if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", n) else "'" + n.replace("'", "''") + "'"


def m_value(v):
    if v is None:
        return "null"
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, (int, float)):
        return repr(v)
    return '"' + str(v).replace('"', '""') + '"'


def m_source(columns, rows):
    head = ", ".join(f'#"{c["name"]}" = {c["m"]}' for c in columns)
    body = ",\n".join("        {" + ", ".join(m_value(r[c["name"]]) for c in columns) + "}" for r in rows)
    return f"let\n    Source = #table(type table [{head}], {{\n{body}\n    }})\nin\n    Source"


def tmdl_table(name, columns, rows, measures=(), hidden=False):
    out = [f"table {tmdl_name(name)}"] + (["\tisHidden"] if hidden else []) + [""]
    for m in measures:
        expr = m["expr"].strip("\n").split("\n")
        if len(expr) == 1:
            out.append(f"\tmeasure {tmdl_name(m['name'])} = {expr[0].strip()}")
        else:
            out.append(f"\tmeasure {tmdl_name(m['name'])} =")
            out += ["\t\t\t" + line for line in expr]
        if m.get("format"):
            out.append(f"\t\tformatString: {m['format']}")
        if m.get("folder"):
            out.append(f"\t\tdisplayFolder: {m['folder']}")
        out.append("")
    for c in columns:
        out += [f"\tcolumn {tmdl_name(c['name'])}", f"\t\tdataType: {c['type']}"]
        if c.get("format"):
            out.append(f"\t\tformatString: {c['format']}")
        if c.get("hidden"):
            out.append("\t\tisHidden")
        out += ["\t\tsummarizeBy: none", f"\t\tsourceColumn: {c['name']}"]
        if c.get("sort_by"):
            out.append(f"\t\tsortByColumn: {tmdl_name(c['sort_by'])}")
        if c.get("category"):
            out.append(f"\t\tdataCategory: {c['category']}")
        out.append("")
    out += [f"\tpartition {tmdl_name(name)} = m", "\t\tmode: import", "\t\tsource ="]
    out += ["\t\t\t\t" + line for line in m_source(columns, rows).split("\n")]
    return "\n".join(out) + "\n"


def col(name, kind, **kw):
    m_types = {"int64": "Int64.Type", "double": "number", "string": "text"}
    return dict(name=name, type=kind, m=m_types[kind], **kw)


def dax_summary_channels():
    return """
VAR a = MIN ( Nam[FY] )
VAR b = MAX ( Nam[FY] )
VAR shA = CALCULATE ( SUM ( Nam[direct_share] ), Nam[FY] = a )
VAR shB = CALCULATE ( SUM ( Nam[direct_share] ), Nam[FY] = b )
VAR hi = MINX ( TOPN ( 1, VALUES ( Nam[FY] ), CALCULATE ( SUM ( Nam[direct_share] ) ), DESC ), Nam[FY] )
VAR lo = MINX ( TOPN ( 1, VALUES ( Nam[FY] ), CALCULATE ( SUM ( Nam[direct_share] ) ), ASC ), Nam[FY] )
VAR shHi = CALCULATE ( SUM ( Nam[direct_share] ), Nam[FY] = hi )
VAR shLo = CALCULATE ( SUM ( Nam[direct_share] ), Nam[FY] = lo )
VAR wsA = CALCULATE ( SUM ( Nam[wholesale_share] ), Nam[FY] = a )
VAR dA = CALCULATE ( SUM ( Nam[nike_direct_revenue] ), Nam[FY] = a ) / 1000
VAR dB = CALCULATE ( SUM ( Nam[nike_direct_revenue] ), Nam[FY] = b ) / 1000
VAR wA = CALCULATE ( SUM ( Nam[nike_brand_wholesale_revenue] ), Nam[FY] = a ) / 1000
VAR wB = CALCULATE ( SUM ( Nam[nike_brand_wholesale_revenue] ), Nam[FY] = b ) / 1000
RETURN
    IF (
        a = b,
        "Trong FY" & a & ", NIKE Direct chiếm " & FORMAT ( shA, "0.0%", "vi-VN" ) & " doanh thu NIKE Brand ("
            & FORMAT ( dA, "#,0.0", "vi-VN" ) & " tỷ USD), còn bán buôn chiếm " & FORMAT ( wsA, "0.0%", "vi-VN" ) & " ("
            & FORMAT ( wA, "#,0.0", "vi-VN" ) & " tỷ USD).",
        "Tỷ trọng NIKE Direct trong doanh thu NIKE Brand là " & FORMAT ( shA, "0.0%", "vi-VN" ) & " ở FY" & a & " và "
            & FORMAT ( shB, "0.0%", "vi-VN" ) & " ở FY" & b
            & IF ( hi <> a && hi <> b, ", cao nhất " & FORMAT ( shHi, "0.0%", "vi-VN" ) & " (FY" & hi & ")" )
            & IF ( lo <> a && lo <> b, IF ( hi <> a && hi <> b, " và thấp nhất ", ", thấp nhất " )
                & FORMAT ( shLo, "0.0%", "vi-VN" ) & " (FY" & lo & ")" )
            & ". Doanh thu NIKE Direct đi từ " & FORMAT ( dA, "#,0.0", "vi-VN" ) & " đến " & FORMAT ( dB, "#,0.0", "vi-VN" )
            & " tỷ USD, còn bán buôn từ " & FORMAT ( wA, "#,0.0", "vi-VN" ) & " đến " & FORMAT ( wB, "#,0.0", "vi-VN" ) & " tỷ USD."
    )
"""


def dax_summary_performance():
    return """
VAR a = MIN ( Nam[FY] )
VAR b = MAX ( Nam[FY] )
VAR rA = CALCULATE ( SUM ( Nam[revenues] ), Nam[FY] = a ) / 1000
VAR rB = CALCULATE ( SUM ( Nam[revenues] ), Nam[FY] = b ) / 1000
VAR gA = CALCULATE ( SUM ( Nam[gross_margin] ), Nam[FY] = a )
VAR gB = CALCULATE ( SUM ( Nam[gross_margin] ), Nam[FY] = b )
VAR eA = CALCULATE ( SUM ( Nam[ebit_margin] ), Nam[FY] = a )
VAR eB = CALCULATE ( SUM ( Nam[ebit_margin] ), Nam[FY] = b )
VAR yHi = MINX ( TOPN ( 1, VALUES ( Nam[FY] ), CALCULATE ( SUM ( Nam[ebit_margin] ) ), DESC ), Nam[FY] )
VAR eHi = CALCULATE ( SUM ( Nam[ebit_margin] ), Nam[FY] = yHi )
VAR iA = CALCULATE ( SUM ( Nam[inventories] ), Nam[FY] = a ) / 1000
VAR iB = CALCULATE ( SUM ( Nam[inventories] ), Nam[FY] = b ) / 1000
RETURN
    IF (
        a = b,
        "Trong FY" & a & ", doanh thu NIKE, Inc. là " & FORMAT ( rA, "#,0.0", "vi-VN" ) & " tỷ USD, biên lợi nhuận gộp "
            & FORMAT ( gA, "0.0%", "vi-VN" ) & ", biên EBIT (non-GAAP) " & FORMAT ( eA, "0.0%", "vi-VN" )
            & ", hàng tồn kho cuối kỳ " & FORMAT ( iA, "#,0.0", "vi-VN" ) & " tỷ USD.",
        "Doanh thu NIKE, Inc. đi từ " & FORMAT ( rA, "#,0.0", "vi-VN" ) & " (FY" & a & ") đến " & FORMAT ( rB, "#,0.0", "vi-VN" )
            & " tỷ USD (FY" & b & "). Biên lợi nhuận gộp là " & FORMAT ( gA, "0.0%", "vi-VN" ) & " (FY" & a & ") và "
            & FORMAT ( gB, "0.0%", "vi-VN" ) & " (FY" & b & "), còn biên EBIT (non-GAAP) là " & FORMAT ( eA, "0.0%", "vi-VN" ) & " (FY" & a & ") và "
            & FORMAT ( eB, "0.0%", "vi-VN" ) & " (FY" & b & ")"
            & IF ( yHi <> a && yHi <> b, ", cao nhất " & FORMAT ( eHi, "0.0%", "vi-VN" ) & " (FY" & yHi & ")" )
            & ". Hàng tồn kho cuối kỳ đi từ " & FORMAT ( iA, "#,0.0", "vi-VN" ) & " đến " & FORMAT ( iB, "#,0.0", "vi-VN" ) & " tỷ USD."
    )
"""


SIGNED = "+0.0%;−0.0%;0.0%"
SIGNED0 = "+0%;−0%;0%"
CN = 'IF ( [Cơ sở đang chọn] = "Loại trừ tỷ giá", SUM ( Nam[{cn}] ), SUM ( Nam[{rep}] ) )'
MEASURES = [
    # Distribution strategy (NIKE Brand)
    dict(name="Tỷ trọng NIKE Direct", expr="SUM ( Nam[direct_share] )", format="0.0%"),
    dict(name="Tỷ trọng bán buôn", expr="SUM ( Nam[wholesale_share] )", format="0.0%"),
    dict(name="Tỷ trọng Global Brand Divisions", expr="SUM ( Nam[gbd_share] )", format="0.00%"),
    dict(name="NIKE Direct (tỷ USD)", expr="DIVIDE ( SUM ( Nam[nike_direct_revenue] ), 1000 )", format="#,0.0"),
    dict(name="Bán buôn (tỷ USD)", expr="DIVIDE ( SUM ( Nam[nike_brand_wholesale_revenue] ), 1000 )", format="#,0.0"),
    dict(name="NIKE Brand (tỷ USD)", expr="DIVIDE ( SUM ( Nam[nike_brand_revenue] ), 1000 )", format="#,0.0"),
    dict(name="Cơ sở đang chọn", expr='SELECTEDVALUE ( CoSo[Cơ sở], "Theo báo cáo" )'),
    dict(name="Tăng trưởng NIKE Direct", expr=CN.format(cn="nike_direct_pct_change_cn", rep="direct_growth"), format=SIGNED),
    dict(name="Tăng trưởng bán buôn", expr=CN.format(cn="nike_brand_wholesale_pct_change_cn", rep="wholesale_growth"), format=SIGNED),
    dict(name="NIKE Direct theo báo cáo", expr="SUM ( Nam[direct_growth] )", format=SIGNED),
    dict(name="NIKE Direct loại trừ tỷ giá", expr="SUM ( Nam[nike_direct_pct_change_cn] )", format=SIGNED0),
    dict(name="Bán buôn theo báo cáo", expr="SUM ( Nam[wholesale_growth] )", format=SIGNED),
    dict(name="Bán buôn loại trừ tỷ giá", expr="SUM ( Nam[nike_brand_wholesale_pct_change_cn] )", format=SIGNED0),
    dict(name="Digital (tỷ USD)", expr="SUM ( Nam[nike_brand_digital_sales] )", format="#,0.0"),
    dict(name="Phần còn lại của NIKE Direct (tỷ USD)", expr="[NIKE Direct (tỷ USD)] - [Digital (tỷ USD)]", format="#,0.0"),
    dict(name="Digital trong NIKE Direct", expr="SUM ( Nam[digital_share_of_direct] )", format="0%"),
    # Financial performance (NIKE, Inc.)
    dict(name="Doanh thu (tỷ USD)", expr="DIVIDE ( SUM ( Nam[revenues] ), 1000 )", format="#,0.0"),
    dict(name="Tăng trưởng doanh thu", expr="SUM ( Nam[revenue_growth] )", format=SIGNED),
    dict(name="Tăng trưởng doanh thu loại trừ tỷ giá", expr="SUM ( Nam[revenues_pct_change_cn] )", format=SIGNED0),
    dict(name="Biên lợi nhuận gộp", expr="SUM ( Nam[gross_margin] )", format="0.0%"),
    dict(name="Lợi nhuận gộp (tỷ USD)", expr="DIVIDE ( SUM ( Nam[gross_profit] ), 1000 )", format="#,0.0"),
    dict(name="SG&A / doanh thu", expr="SUM ( Nam[sga_ratio] )", format="0.0%"),
    dict(name="Vận hành chung / doanh thu", expr="SUM ( Nam[operating_overhead_ratio] )", format="0.0%"),
    dict(name="Tạo nhu cầu / doanh thu", expr="SUM ( Nam[demand_creation_ratio] )", format="0.0%"),
    dict(name="SG&A (tỷ USD)", expr="DIVIDE ( SUM ( Nam[total_selling_admin_expense] ), 1000 )", format="#,0.0"),
    dict(name="Biên EBIT", expr="SUM ( Nam[ebit_margin] )", format="0.0%"),
    dict(name="EBIT do NIKE công bố (tỷ USD)", expr="DIVIDE ( SUM ( Nam[ebit_reported] ), 1000 )", format="#,0.00"),
    dict(name="Biên lợi nhuận sau thuế", expr="SUM ( Nam[net_income_margin] )", format="0.0%"),
    dict(name="Hàng tồn kho (tỷ USD)", expr="DIVIDE ( SUM ( Nam[inventories] ), 1000 )", format="#,0.0"),
    dict(name="Thay đổi hàng tồn kho", expr="SUM ( Nam[inventory_growth] )", format=SIGNED),
    dict(name="Vòng quay hàng tồn kho", expr="SUM ( Nam[inventory_turnover] )", format="0.00"),
    dict(name="Số ngày tồn kho", expr="SUM ( Nam[days_inventory] )", format="0.0"),
    dict(name="Tồn kho bình quân (tỷ USD)", expr="DIVIDE ( SUM ( Nam[average_inventory] ), 1000 )", format="#,0.00"),
    dict(name="Giá vốn hàng bán (tỷ USD)", expr="DIVIDE ( SUM ( Nam[cost_of_sales] ), 1000 )", format="#,0.0"),
    # Text
    dict(name="Tóm tắt kênh", expr=dax_summary_channels()),
    dict(name="Tóm tắt kết quả", expr=dax_summary_performance()),
]


def semantic_model(d, w, out):
    years = d["years"]
    obs, kpis, filings = d["observations"], d["kpis"], d["filings"]
    nam_cols = [col("FY", "int64", format=r"\F\Y0"), col("Năm", "string", sort_by="FY"), col("Năm (ngắn)", "string", sort_by="FY"),
                col("Báo cáo", "string", sort_by="FY"), col("Kỳ báo cáo", "string"), col("Ngày nộp SEC", "string"),
                col("Số accession", "string"), col("Liên kết SEC EDGAR", "string", category="WebUrl")]
    nam_cols += [col(k, "double", hidden=True) for k in list(obs) + list(kpis)]
    nam_rows = []
    for y in years:
        f = filings[str(y)]
        r = {"FY": y, "Năm": f"FY{y}", "Năm (ngắn)": f"FY{str(y)[2:]}", "Báo cáo": f"Form 10-K FY{y}", "Kỳ báo cáo": f"01/06/{y - 1} – 31/05/{y}",
             "Ngày nộp SEC": vn_date(f["filed"]), "Số accession": f["accession"], "Liên kết SEC EDGAR": f["url"]}
        r.update({k: (obs[k].get(str(y)) or {}).get("value") for k in obs})
        r.update({k: kpis[k]["values"][str(y)] for k in kpis})
        nam_rows.append(r)

    tables = {"Nam": tmdl_table("Nam", nam_cols, nam_rows, MEASURES)}
    for view, spec in TABLES.items():
        name = TABLE_NAMES[view]
        rows = []
        for order, row in enumerate(spec, 1):
            if row[0] == "g":
                label = row[1].upper()
            else:
                # Percentages carry their own sign; other units go into the row label (one row level only).
                label = row[2] if row[3] == "%" else f"{row[2]}, {row[3]}"
            for i, y in enumerate(years):
                figure, page, link = ("", "", None) if row[0] == "g" else table_cells(row, y, obs, kpis, filings)
                # Each year is two matrix columns: the figure under "FY2020", then its page reference under a
                # blank header (non-breaking spaces, a different count per year so the headers stay distinct).
                cell = {"Thứ tự": order, "Chỉ tiêu": label, "FY": y}
                rows.append({**cell, "Cột": f"FY{y}", "Thứ tự cột": 2 * i, "Loại": "Số", "Hiển thị": figure, "Liên kết": None})
                rows.append({**cell, "Cột": " " * (i + 1), "Thứ tự cột": 2 * i + 1, "Loại": "Trang", "Hiển thị": page,
                             "Liên kết": link})
        cols = [col("Thứ tự", "int64", hidden=True), col("Chỉ tiêu", "string", sort_by="Thứ tự"),
                col("FY", "int64", hidden=True, format="0"), col("Cột", "string", sort_by="Thứ tự cột"),
                col("Thứ tự cột", "int64", hidden=True), col("Loại", "string", hidden=True), col("Hiển thị", "string"),
                col("Liên kết", "string", category="WebUrl")]
        # Blank (not "") for years outside the slicer range, so the matrix drops those columns.
        measures = [dict(name=f"Giá trị {name}", expr=f'IF ( NOT ISEMPTY ( {name} ), COALESCE ( SELECTEDVALUE ( {name}[Hiển thị] ), "" ) )'),
                    dict(name=f"Liên kết {name}", expr=f"SELECTEDVALUE ( {name}[Liên kết] )"),
                    dict(name=f"Màu chữ {name}", expr=f'IF ( SELECTEDVALUE ( {name}[Loại] ) = "Trang", "{C["faint"]}", "{C["ink"]}" )')]
        tables[name] = tmdl_table(name, cols, rows, measures)
    tables["CoSo"] = tmdl_table("CoSo", [col("Cơ sở", "string", sort_by="Thứ tự"), col("Thứ tự", "int64", hidden=True)],
                                [{"Cơ sở": "Theo báo cáo", "Thứ tự": 1}, {"Cơ sở": "Loại trừ tỷ giá", "Thứ tự": 2}])

    sm = out / f"{NAME}.SemanticModel"
    (sm / "definition" / "tables").mkdir(parents=True, exist_ok=True)
    write_json(sm / "definition.pbism", {"$schema": SCHEMA + "item/semanticModel/definitionProperties/1.0.0/schema.json",
                                         "version": "4.2", "settings": {}})
    write_text(sm / "definition" / "database.tmdl", "database\n\tcompatibilityLevel: 1600\n")
    order = json.dumps(list(tables), ensure_ascii=False, separators=(",", ":"))
    write_text(sm / "definition" / "model.tmdl",
               "model Model\n\tculture: vi-VN\n\tdefaultPowerBIDataSourceVersion: powerBI_V3\n\tsourceQueryCulture: vi-VN\n"
               "\tdataAccessOptions\n\t\tlegacyRedirects\n\t\treturnErrorValuesAsNull\n\n"
               f"annotation PBI_QueryOrder = {order}\n\nannotation __PBI_TimeIntelligenceEnabled = 0\n\n"
               + "".join(f"ref table {tmdl_name(t)}\n" for t in tables))
    rel = "".join(f"relationship {uuid.uuid5(uuid.NAMESPACE_URL, 'nike/' + t)}\n\tfromColumn: {t}.FY\n\ttoColumn: Nam.FY\n\n"
                  for t in TABLE_NAMES.values())
    write_text(sm / "definition" / "relationships.tmdl", rel)
    for t, body in tables.items():
        write_text(sm / "definition" / "tables" / f"{t}.tmdl", body)


# ------------------------------------------------------------------ report (PBIR)
def lit(v):
    if isinstance(v, bool):
        s = "true" if v else "false"
    elif isinstance(v, int):
        s = f"{v}L"
    elif isinstance(v, float):
        s = f"{v:g}D"
    else:
        s = "'" + str(v).replace("'", "''") + "'"
    return {"expr": {"Literal": {"Value": s}}}


def fill(c):
    return {"solid": {"color": lit(c)}}


def props(**kw):
    return [{"properties": {k: (v if isinstance(v, dict) else lit(v)) for k, v in kw.items()}}]


def column(entity, prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def measure(prop, entity="Nam"):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": entity}}, "Property": prop}}


def ref(field):
    f = field.get("Column") or field.get("Measure")
    return f"{f['Expression']['SourceRef']['Entity']}.{f['Property']}"


def proj(field, display=None):
    f = field.get("Column") or field.get("Measure")
    p = {"field": field, "queryRef": ref(field), "nativeQueryRef": f["Property"]}
    if display:
        p["displayName"] = display
    if "Column" in field:
        p["active"] = True
    return p


YEAR_AXIS = column("Nam", "Năm")
SHORT_YEAR_AXIS = column("Nam", "Năm (ngắn)")  # "FY20": every year keeps its label on the narrow charts, as in app.js


def container_objects(title=None, subtitle=None, subtitle_measure=None):
    o = {"border": props(show=False), "dropShadow": props(show=False), "background": props(show=False),
         "padding": props(top=0.0, bottom=0.0, left=0.0, right=0.0)}
    if title:
        o["title"] = props(show=True, text=title, fontFamily=SANS, fontSize=13.0, bold=True, fontColor=fill(C["ink"]),
                           alignment="left", titleWrap=True)
    else:
        o["title"] = props(show=False)
    if subtitle or subtitle_measure:
        text = {"expr": measure(subtitle_measure)} if subtitle_measure else lit(subtitle)
        o["subTitle"] = props(show=True, text=text, fontFamily=SANS, fontSize=10.0, fontColor=fill(C["muted"]),
                              alignment="left", titleWrap=True)
    return o


class Page:
    def __init__(self, name, display, height):
        self.name, self.display, self.height, self.visuals, self.z = name, display, height, [], 0

    def add(self, name, x, y, w, h, visual):
        self.z += 1
        self.visuals.append({"$schema": VC_SCHEMA, "name": name,
                             "position": {"x": x, "y": y, "z": self.z * 1000, "height": h, "width": w, "tabOrder": self.z * 1000},
                             "visual": visual})

    def text(self, name, x, y, w, h, paragraphs):
        self.add(name, x, y, w, h, {"visualType": "textbox", "objects": {"general": [{"properties": {"paragraphs": paragraphs}}]},
                                    "visualContainerObjects": container_objects(), "drillFilterOtherVisuals": True})


def run(text, size=11, color=C["ink2"], bold=False, family=SANS, url=None):
    style = {"fontFamily": family, "fontSize": f"{size}pt", "color": color}
    if bold:
        style["fontWeight"] = "bold"
    r = {"value": text, "textStyle": style}
    if url:
        r["url"] = url
    return r


def para(*runs):
    return {"textRuns": list(runs), "horizontalTextAlignment": "left"}


def cartesian(kind, series, title, subtitle=None, subtitle_measure=None, tooltips=(), labels=True, value_start=None,
              totals=False, legend=True, axis=YEAR_AXIS):
    """series: [(measure, display name, colour)]"""
    objects = {
        "categoryAxis": props(showAxisTitle=False, fontFamily=SANS, fontSize=10.0, labelColor=fill(C["muted"])),
        "valueAxis": props(showAxisTitle=False, fontFamily=SANS, fontSize=10.0, labelColor=fill(C["muted"]),
                           gridlineShow=True, gridlineColor=fill(C["grid"]),
                           **({"start": value_start} if value_start is not None else {})),
        "legend": props(show=legend, position="Top", showTitle=False, fontFamily=SANS, fontSize=10.0, labelColor=fill(C["ink2"])),
        "labels": props(show=labels, labelDisplayUnits=1.0, fontFamily=SANS, fontSize=9.0, color=fill(C["ink2"])),
        "dataPoint": [{"properties": {"fill": fill(series[0][2])}}]
                     + [{"properties": {"fill": fill(c)}, "selector": {"metadata": ref(measure(m))}} for m, _, c in series],
    }
    if kind == "lineChart":
        objects["lineStyles"] = props(lineChartType="linear", showMarker=True, strokeWidth=2)
    if kind == "hundredPercentStackedBarChart":
        # Thicker bars than the theme's 50% gap, so the share labels fit inside them.
        objects["categoryAxis"][0]["properties"]["innerPadding"] = lit(25.0)
        objects["labels"][0]["properties"].update(color=fill("#FFFFFF"), labelPrecision=lit(1))
        # Dark ink on the orange wholesale bar would read at about 5.3:1 (white is 3.2:1), but Power BI Desktop
        # 2.159 draws labels inside 100% stacked bars in white whatever colour is set; the rule is kept because
        # it is harmless and takes effect if a later build honours it.
        objects["labels"].append({"properties": {"color": fill(C["ink"])},
                                  "selector": {"metadata": ref(measure("Tỷ trọng bán buôn"))}})
    if totals:
        objects["totals"] = props(show=True, fontFamily=SANS, fontSize=9.0, color=fill(C["ink2"]))
    query = {"Category": {"projections": [proj(axis)]},
             "Y": {"projections": [proj(measure(m), d) for m, d, _ in series]}}
    if tooltips:
        query["Tooltips"] = {"projections": [proj(measure(m), d) for m, d in tooltips]}
    sort = {"sort": [{"field": axis, "direction": "Ascending"}], "isDefaultSort": False}
    return {"visualType": kind, "query": {"queryState": query, "sortDefinition": sort}, "objects": objects,
            "visualContainerObjects": container_objects(title, subtitle, subtitle_measure), "drillFilterOtherVisuals": True}


def year_slicer():
    return {"visualType": "slicer",
            "query": {"queryState": {"Values": {"projections": [proj(column("Nam", "FY"), "Năm tài chính")]}}},
            "objects": {"data": props(mode="Between"), "header": props(show=False),
                        "numericInputStyle": props(fontFamily=SANS, fontSize=11.0, fontColor=fill(C["ink"])),
                        "slider": props(show=True, color=fill(C["ink2"]))},
            "visualContainerObjects": container_objects("Năm tài chính"),
            "syncGroup": {"groupName": "NamTaiChinh", "fieldChanges": True, "filterChanges": True},
            "drillFilterOtherVisuals": True}


def basis_slicer():
    return {"visualType": "slicer",
            "query": {"queryState": {"Values": {"projections": [proj(column("CoSo", "Cơ sở"), "Cơ sở tính tăng trưởng")]}}},
            "objects": {"data": props(mode="Basic"), "general": props(orientation=1.0),
                        "selection": props(singleSelect=True, strictSingleSelect=True, selectAllCheckboxEnabled=False),
                        "header": props(show=False),
                        "items": props(fontFamily=SANS, fontSize=10.0, fontColor=fill(C["ink"]))},
            "visualContainerObjects": container_objects(), "drillFilterOtherVisuals": True}


def card(m):
    return {"visualType": "card", "query": {"queryState": {"Values": {"projections": [proj(measure(m))]}}},
            "objects": {"labels": props(fontFamily=SANS, fontSize=12.0, color=fill(C["ink"])),
                        "categoryLabels": props(show=False), "wordWrap": props(show=True)},
            "visualContainerObjects": container_objects(), "drillFilterOtherVisuals": True}


def matrix(table):
    value = measure(f"Giá trị {table}", table)
    every_cell = {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": ref(value)}
    # Columns alternate figure / page reference per year, all right-aligned: every figure ends under its year
    # header and the page references sit in their own column after it (no value-name subheader row).
    return {"visualType": "pivotTable",
            "query": {"queryState": {"Rows": {"projections": [proj(column(table, "Chỉ tiêu"))]},
                                     "Columns": {"projections": [proj(column(table, "Cột"))]},
                                     "Values": {"projections": [proj(value, "Giá trị")]}}},
            "objects": {"subTotals": props(rowSubtotals=False, columnSubtotals=False),
                        "rowHeaders": props(fontFamily=SANS, fontSize=10.0, fontColor=fill(C["ink"]), wordWrap=True),
                        # Year headers right-aligned over their figures; the corner title ("Chỉ tiêu") left-aligned
                        # over the row labels.
                        "columnHeaders": props(fontFamily=SANS, fontSize=10.0, fontColor=fill(C["ink"]), bold=True,
                                               alignment="Right", titleAlignment="Left", wordWrap=True),
                        "values": [{"properties": {"fontFamily": lit(SANS), "fontSize": lit(10.0), "fontColor": fill(C["ink"])}},
                                   # page references: muted, and Ctrl + click opens that year's 10-K
                                   {"properties": {"fontColor": {"solid": {"color": {"expr": measure(f"Màu chữ {table}", table)}}},
                                                   "webURL": {"expr": measure(f"Liên kết {table}", table)}},
                                    "selector": every_cell}],
                        "columnFormatting": [{"properties": {"alignment": lit("Right")}, "selector": {"metadata": ref(value)}}],
                        "grid": props(gridVertical=False, gridHorizontal=True, gridHorizontalColor=fill(C["grid"]),
                                      rowPadding=3.0)},
            "visualContainerObjects": container_objects(), "drillFilterOtherVisuals": True}


def header(page, w, view):
    page.text("tieu_de", MARGIN, 16, 820, 50, [para(run(w.title, 20, C["ink"], True, SERIF))])
    page.text("mo_ta", MARGIN, 66, 820, 50, [para(run(w.subtitle, 11, C["ink2"]))])
    page.add("loc_nam", 900, 16, 340, 110, year_slicer())
    # Question and summary each take half the row (580 units, the channel page's chart columns): at 460 units the
    # performance summary needs five lines and the card cuts the fifth off with "…".
    page.text("cau_hoi", MARGIN, 140, 580, 110, [para(run(w.question[view], 17, C["ink"], True, SERIF))])
    page.text("tom_tat_nhan", 660, 140, 580, 26, [para(run(w.summary_label, 10, C["muted"]))])
    page.add("tom_tat", 660, 166, 580, 110, card("Tóm tắt kênh" if view == "channels" else "Tóm tắt kết quả"))


def source(page, name, x, y, w, text, h=44):
    page.text(name, x, y, w, h, [para(run(text, 9, C["muted"]))])


def notes(page, name, x, y, w, h, title, items):
    paragraphs = [para(run(title, 13, C["ink"], True))]
    for dt, dd in items:
        paragraphs += [para(run(" ", 5, C["ink2"])), para(run(dt, 11, C["ink"], True)), para(run(dd, 10.5, C["ink2"]))]
    page.text(name, x, y, w, h, paragraphs)


PAGE_REF_KEY = "có số trang 10-K ở cột nhỏ bên phải (Ctrl + bấm vào số trang để mở báo cáo trên SEC)"


def table_key(text):
    """The web table key reworded for Power BI: page references sit in their own column and there are no italics."""
    key = re.sub(r"có kèm số trang 10-K(?: \(bấm để mở báo cáo trên SEC\))?", PAGE_REF_KEY, text)
    key = key.replace(" Số in nghiêng là số nhóm tự tính.", " Số nhóm tự tính không có số trang.")
    if PAGE_REF_KEY not in key or "in nghiêng" in key:
        raise SystemExit(f"index.html: table key wording changed; update table_key() in build_powerbi.py: {text}")
    return key


def table_section(page, w, view, table, y, rows):
    page.text("bang_tieu_de", MARGIN, y, 1200, 32, [para(run(w.table_title[view], 13, C["ink"], True))])
    key = table_key(w.table_key[view])
    if w.pending:
        key += " " + w.pending_key
    page.text("bang_chu_thich", MARGIN, y + 32, 1200, 28, [para(run(key, 9.5, C["muted"]))])
    h = 40 + 25 * rows
    page.add("bang", MARGIN, y + 62, 1200, h, matrix(table))
    return y + 62 + h


def page_channels(w):
    L, R, CW = MARGIN, 660, 580
    p = Page("kenh", w.tab["channels"], 0)
    header(p, w, "channels")
    t, s = w.fig["fig-share"]
    lg = lambda n: w.legend("share", n)
    p.add("co_cau", L, 300, 1200, 340, cartesian(
        "hundredPercentStackedBarChart",
        [("Tỷ trọng NIKE Direct", lg("NIKE Direct"), C["direct"]), ("Tỷ trọng bán buôn", lg("Bán buôn"), C["wholesale"]),
         ("Tỷ trọng Global Brand Divisions", lg("Global Brand Divisions"), C["gbd"])],
        t, s, tooltips=[("NIKE Direct (tỷ USD)", "NIKE Direct, tỷ USD"), ("Bán buôn (tỷ USD)", "Bán buôn, tỷ USD"),
                        ("NIKE Brand (tỷ USD)", "Doanh thu NIKE Brand (mẫu số), tỷ USD")]))
    source(p, "nguon_co_cau", L, 644, 1200, w.sources["share"])

    t, s = w.fig["fig-channel-revenue"]
    lg = lambda n: w.legend("channelRevenue", n)
    p.add("doanh_thu_kenh", L, 710, CW, 340, cartesian(
        "clusteredColumnChart", [("NIKE Direct (tỷ USD)", lg("NIKE Direct"), C["direct"]),
                                 ("Bán buôn (tỷ USD)", lg("Bán buôn"), C["wholesale"])],
        t, s, tooltips=[("NIKE Brand (tỷ USD)", "Tổng NIKE Brand, tỷ USD")]))
    source(p, "nguon_doanh_thu_kenh", L, 1054, CW, w.sources["channelRevenue"])

    t, _ = w.fig["fig-growth"]
    p.add("tang_truong", R, 710, CW, 340, cartesian(
        "clusteredColumnChart", [("Tăng trưởng NIKE Direct", lg("NIKE Direct"), C["direct"]),
                                 ("Tăng trưởng bán buôn", lg("Bán buôn"), C["wholesale"])],
        t, subtitle_measure="Phụ đề tăng trưởng", labels=False,
        tooltips=[("NIKE Direct theo báo cáo", "NIKE Direct, theo báo cáo"), ("NIKE Direct loại trừ tỷ giá", "NIKE Direct, loại trừ tỷ giá"),
                  ("Bán buôn theo báo cáo", "Bán buôn, theo báo cáo"), ("Bán buôn loại trừ tỷ giá", "Bán buôn, loại trừ tỷ giá")]))
    # In the chart's title row, right of the title: the slicer's header icons (shown while it is hovered or selected)
    # then float in the gap above the chart instead of covering the subtitle, at any zoom.
    p.add("co_so_tang_truong", R + 330, 710, CW - 330, 24, basis_slicer())
    source(p, "nguon_tang_truong", R, 1054, CW,
           "Theo báo cáo: " + w.sources["growth"].removeprefix("Nguồn: ") + " Loại trừ tỷ giá: "
           + w.sources["growth_cn"].removeprefix("Nguồn: "), h=58)

    t, s = w.fig["fig-digital"]
    p.add("digital", L, 1130, CW, 340, cartesian(
        "columnChart", [("Digital (tỷ USD)", w.legend("digital", "Digital"), C["digital"]),
                        ("Phần còn lại của NIKE Direct (tỷ USD)", "Phần còn lại của NIKE Direct (nhóm tự tính, xấp xỉ)", C["directLight"])],
        t, s + " Chiều cao cả cột bằng doanh thu NIKE Direct.", labels=False, totals=True,
        tooltips=[("NIKE Direct (tỷ USD)", "NIKE Direct (tổng), tỷ USD"), ("Digital trong NIKE Direct", "Digital trong NIKE Direct (≈)")]))
    source(p, "nguon_digital", L, 1474, CW, w.sources["digital"], h=58)
    notes(p, "ghi_chu", R, 1130, CW, 420, *w.notes["channels"])

    end = table_section(p, w, "channels", "BangKenh", 1570, len(TABLES["kenh"]))
    p.height = end + 40
    return p


def page_performance(w):
    p = Page("ketqua", w.tab["performance"], 0)
    header(p, w, "performance")
    cw, xs = 386, [MARGIN, MARGIN + 407, MARGIN + 814]
    chart = lambda *a, **k: cartesian(*a, axis=SHORT_YEAR_AXIS, **k)

    def fig(name, fid, src, x, y, visual):
        t, s = w.fig[fid]
        visual["visualContainerObjects"] = container_objects(t, s)
        p.add(name, x, y, cw, 300, visual)
        source(p, "nguon_" + name, x, y + 304, cw, w.sources[src], h=60)

    y1, y2 = 300, 690
    fig("doanh_thu", "fig-revenue", "revenue", xs[0], y1, chart(
        "clusteredColumnChart", [("Doanh thu (tỷ USD)", "Doanh thu", C["total"])], None, legend=False,
        tooltips=[("Tăng trưởng doanh thu", "So với năm trước, theo báo cáo"),
                  ("Tăng trưởng doanh thu loại trừ tỷ giá", "Loại trừ tỷ giá (NIKE, non-GAAP)")]))
    fig("bien_gop", "fig-gross-margin", "grossMargin", xs[1], y1, chart(
        "lineChart", [("Biên lợi nhuận gộp", "Biên lợi nhuận gộp", C["total"])], None, legend=False,
        value_start=lit(nice_min(w.data, "gross_margin", 0.02)),
        tooltips=[("Lợi nhuận gộp (tỷ USD)", "Lợi nhuận gộp, tỷ USD"), ("Doanh thu (tỷ USD)", "Doanh thu, tỷ USD")]))
    fig("sga", "fig-sga", "sga", xs[2], y1, chart(
        "columnChart", [("Vận hành chung / doanh thu", w.legend("sga", "Vận hành chung"), C["overhead"]),
                        ("Tạo nhu cầu / doanh thu", w.legend("sga", "Tạo nhu cầu"), C["demand"])],
        None, labels=False, totals=True,
        tooltips=[("SG&A / doanh thu", "Tổng SG&A / doanh thu"), ("SG&A (tỷ USD)", "Tổng SG&A, tỷ USD")]))
    fig("ebit", "fig-ebit", "ebit", xs[0], y2, chart(
        "lineChart", [("Biên EBIT", "Biên EBIT (non-GAAP)", C["total"])], None, legend=False, value_start=lit(0.0),
        tooltips=[("EBIT do NIKE công bố (tỷ USD)", "EBIT do NIKE công bố, tỷ USD"),
                  ("Biên lợi nhuận sau thuế", "Biên lợi nhuận sau thuế (GAAP)")]))
    fig("ton_kho", "fig-inventory", "inventory", xs[1], y2, chart(
        "clusteredColumnChart", [("Hàng tồn kho (tỷ USD)", "Hàng tồn kho", C["total"])], None, legend=False,
        tooltips=[("Thay đổi hàng tồn kho", "So với 31/5 năm trước"), ("Vòng quay hàng tồn kho", "Vòng quay hàng tồn kho (lần)")]))
    fig("so_ngay_ton_kho", "fig-days", "days", xs[2], y2, chart(
        "lineChart", [("Số ngày tồn kho", "Số ngày tồn kho bình quân", C["total"])], None, legend=False,
        value_start=lit(nice_min(w.data, "days_inventory", 10)),
        tooltips=[("Tồn kho bình quân (tỷ USD)", "Tồn kho bình quân đầu và cuối kỳ, tỷ USD"),
                  ("Giá vốn hàng bán (tỷ USD)", "Giá vốn hàng bán, tỷ USD")]))
    title, items = w.notes["performance"]
    notes(p, "ghi_chu", MARGIN, 1080, 1200, 230, title, items)
    end = table_section(p, w, "performance", "BangKetQua", 1330, len(TABLES["ketqua"]))
    p.height = end + 40
    return p


def page_sources(w, d):
    p = Page("nguon", w.footer_title, 760)
    p.text("tieu_de", MARGIN, 16, 1200, 50, [para(run(w.footer_title, 20, C["ink"], True, SERIF))])
    p.text("bao_cao_tieu_de", MARGIN, 80, 1200, 30, [para(run(w.footer_heads[0], 13, C["ink"], True))])
    fields = [column("Nam", c) for c in ("Báo cáo", "Ngày nộp SEC", "Số accession", "Liên kết SEC EDGAR")]
    p.add("bao_cao", MARGIN, 112, 1200, 250, {
        "visualType": "tableEx",
        "query": {"queryState": {"Values": {"projections": [proj(f) for f in fields]}},
                  "sortDefinition": {"sort": [{"field": fields[0], "direction": "Ascending"}], "isDefaultSort": False}},
        "objects": {"values": props(fontFamily=SANS, fontSize=10.5, urlIcon=False),
                    "columnHeaders": props(fontFamily=SANS, fontSize=10.5, bold=True),
                    "total": props(totals=False)},
        "visualContainerObjects": container_objects(), "drillFilterOtherVisuals": True})
    p.text("phuong_phap", MARGIN, 390, 740, 250,
           [para(run(w.footer_heads[1], 13, C["ink"], True))] + [para(run("– " + li, 10.5, C["ink2"])) for li in w.method])
    p.text("tinh_trang", 820, 390, 420, 250, [
        para(run(w.footer_heads[2], 13, C["ink"], True)), para(run(status_sentence(d), 10.5, C["ink2"])),
        para(run(w.fine, 9.5, C["muted"]))])
    p.text("ban_web", MARGIN, 660, 1200, 60, [
        para(run("Bản web của dashboard: ", 10.5, C["ink2"]),
             run(DASHBOARD_URL, 10.5, C["direct"], url=DASHBOARD_URL))])
    return p


DASHBOARD_URL = "https://miurikyasakytg2006-dot.github.io/nike-financial-dashboard/"


def status_sentence(d):
    """Same wording as renderFooter() in app.js: only statuses that occur are listed."""
    c = d["meta"]["status_counts"]
    parts = [(c.get("Khớp nguồn chính thức", 0), " khớp nguồn chính thức"),
             (c.get("Đã xác minh (người)", 0), " đã được nhóm kiểm tra trên 10-K"),
             (c.get(STATUS_PENDING, 0), " chưa được kiểm tra (†)"),
             (c.get("Thiếu / không công bố", 0), " không được NIKE công bố")]
    shown = ", ".join(f"{n}{t}" for n, t in parts if n)
    return f"{sum(c.values())} số liệu báo cáo: {shown}. Dữ liệu trích xuất ngày {vn_date(d['meta']['extracted_on'])}."


def nice_min(d, kpi, step):
    vals = [v for v in d["kpis"][kpi]["values"].values() if v is not None]
    m = min(vals)
    return float(Decimal(str((m - step / 2) // step * step)).quantize(Decimal("0.0001")))


def report(d, w, out):
    rp = out / f"{NAME}.Report"
    write_json(rp / "definition.pbir", {"$schema": SCHEMA + "item/report/definitionProperties/2.0.0/schema.json",
                                        "version": "4.0", "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}})
    theme_rel = Path("StaticResources") / "SharedResources" / "BaseThemes" / f"{THEME}.json"
    theme = rp / theme_rel
    if not theme.exists():
        found = [str(OUT / f"{NAME}.Report" / theme_rel)] + [p for pattern in THEME_SOURCES for p in glob.glob(f"{pattern}/{THEME}.json")]
        found += [str(Path(store_install_dir()) / THEME_SOURCES[0].split("*/", 1)[1] / f"{THEME}.json")] if store_install_dir() else []
        found = [p for p in found if Path(p).exists()]
        if not found:
            raise SystemExit(f"Base theme {THEME}.json not found; install Power BI Desktop or copy it to {theme}")
        theme.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(found[0], theme)
    df = rp / "definition"
    write_json(df / "version.json", {"$schema": SCHEMA + "item/report/definition/versionMetadata/1.0.0/schema.json", "version": "2.0.0"})
    write_json(df / "report.json", {
        "$schema": SCHEMA + "item/report/definition/report/3.3.0/schema.json",
        "themeCollection": {"baseTheme": {"name": THEME, "reportVersionAtImport": {"visual": "2.4.0", "report": "3.0.0", "page": "2.0.0"},
                                          "type": "SharedResources"}},
        "objects": {"outspacePane": props(expanded=False, visible=False)},
        "resourcePackages": [{"name": "SharedResources", "type": "SharedResources",
                              "items": [{"name": THEME, "path": f"BaseThemes/{THEME}.json", "type": "BaseTheme"}]}],
        "settings": {"locale": "vi-VN", "useStylableVisualContainerHeader": True, "exportDataMode": "AllowSummarized",
                     "defaultDrillFilterOtherVisuals": True, "allowChangeFilterTypes": True}})
    pages = [page_channels(w), page_performance(w), page_sources(w, d)]
    write_json(df / "pages" / "pages.json", {"$schema": SCHEMA + "item/report/definition/pagesMetadata/1.0.0/schema.json",
                                             "pageOrder": [p.name for p in pages], "activePageName": pages[0].name})
    for p in pages:
        pd = df / "pages" / p.name
        write_json(pd / "page.json", {
            "$schema": PAGE_SCHEMA, "name": p.name, "displayName": p.display, "displayOption": "FitToWidth",
            "height": p.height, "width": PAGE_W,
            "objects": {"background": props(color=fill("#FFFFFF"), transparency=0.0),
                        "outspace": props(color=fill("#FFFFFF"), transparency=0.0)}})
        for v in p.visuals:
            write_json(pd / "visuals" / v["name"] / "visual.json", v)


# ------------------------------------------------------------------ files
def store_install_dir():
    """Install folder of the Microsoft Store build of Power BI Desktop (WindowsApps cannot be listed)."""
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-Command",
                            "(Get-AppxPackage -Name Microsoft.MicrosoftPowerBIDesktop).InstallLocation"],
                           capture_output=True, text=True, timeout=60)
        return r.stdout.strip() or None
    except (OSError, subprocess.SubprocessError):
        return None


def write_text(path, text):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def write_json(path, obj):
    write_text(path, json.dumps(obj, ensure_ascii=False, indent=2) + "\n")


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else OUT
    d = json.loads(DATA.read_text(encoding="utf-8"))
    w = Wording()
    w.data = d
    w.pending = d["meta"]["status_counts"].get(STATUS_PENDING, 0) > 0
    MEASURES.append(dict(name="Phụ đề tăng trưởng",
                         expr=f'IF ( [Cơ sở đang chọn] = "Loại trừ tỷ giá", "{w.growth_sub_cn}", "{w.growth_sub}" )'))
    stale = [p for p in (out / f"{NAME}.Report" / "definition" / "pages").glob("*/visuals/*")] if out.exists() else []
    semantic_model(d, w, out)
    report(d, w, out)
    written = {v.parent for v in (out / f"{NAME}.Report" / "definition" / "pages").glob("*/visuals/*/visual.json")}
    write_json(out / f"{NAME}.pbip", {"$schema": SCHEMA + "pbip/pbipProperties/1.0.0/schema.json", "version": "1.0",
                                      "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
                                      "settings": {"enableAutoRecovery": True}})
    write_text(out / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n*.pbix\n")
    print(f"Power BI project written: {out / (NAME + '.pbip')}")
    leftovers = sorted(set(stale) - written)
    if leftovers:
        print("NOTE: visuals from an earlier build are still on disk (remove them):", *leftovers, sep="\n  ")
    return 0


if __name__ == "__main__":
    sys.exit(main())
