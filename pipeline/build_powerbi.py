"""Build the Power BI version of the dashboard as a Power BI Project (PBIP).

Usage:  python -m pipeline.build_powerbi [output folder]      (default: powerbi/)

Reads site/data/nike-financials.json (written by build_site_data from the validated
workbook) and the wording of site/index.html and site/assets/app.js, and writes:

  NIKE_Dashboard.pbip                 open this in Power BI Desktop
  NIKE_Dashboard.SemanticModel/       TMDL; the data is embedded in Power Query (no file paths)
  NIKE_Dashboard.Report/              PBIR pages and visuals

Power BI Desktop turns the project into the .pbix handed to the team: open the .pbip,
Refresh, then File > Save as > Power BI file. Figures come from the JSON unchanged; the
key findings and notes are the strings the pipeline wrote there, so the web, Excel and
Power BI versions say the same thing.
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
COVID = "FY2020Q4"
DASHBOARD_URL = "https://miurikyasakytg2006-dot.github.io/nike-financial-dashboard/"
SRC_PAGE = "Nguồn và giải thích"

# Palette and type, as in site/assets/app.js and styles.css
C = dict(direct="#2a78d6", wholesale="#eb6834", gbd="#a3a9b2", directLight="#86b6ef", digital="#1c5cab",
         total="#475467", margin="#6250d6", up="#1baf7a", down="#c8453b", past="#98a2b3",
         ink="#15191e", ink2="#3b4350", muted="#5b6472", faint="#6f7784", grid="#e6e9ee", axis="#c3c9d1", rule="#dde1e6")
SANS, SERIF = "Segoe UI", "Cambria"
PAGE_W, MARGIN = 1280, 40
L, R, CW = MARGIN, 660, 580          # the two chart columns
NBSP, NB_HYPHEN, WJ = "\u00a0", "\u2011", "\u2060"   # escapes, so that no editor turns them into plain characters


def glue(text):
    """Keep together what reads badly when a text box or a table label splits it (the same rule as in
    build_excel_dashboard.py): "10-K", "10-Q" and "non-GAAP" take a non-breaking hyphen, a report stays with its
    period, a figure with its unit, and the two ends of a period range."""
    text = (text.replace("10-K", "10" + NB_HYPHEN + "K").replace("10-Q", "10" + NB_HYPHEN + "Q")
            .replace("8-K", "8" + NB_HYPHEN + "K").replace("non-GAAP", "non" + NB_HYPHEN + "GAAP")
            .replace("COVID-19", "COVID" + NB_HYPHEN + "19"))
    text = re.sub(r"(10" + NB_HYPHEN + "[KQ]) (?=Q[1-4]|FY)", "\\1" + NBSP, text)
    text = re.sub(r"\b(Q[1-4]) (FY\d{4})", "\\1" + NBSP + "\\2", text)
    text = re.sub(r"(FY\d{4}) → (FY\d{4})", "\\1" + NBSP + "→" + NBSP + "\\2", text)
    # A word joiner on each side of the en dash keeps a period range such as "FY2019–FY2021" on one line.
    text = re.sub(r"(FY\d{4})–(FY\d{4})", "\\1" + WJ + "–" + WJ + "\\2", text)
    text = re.sub(r"(\d) (tỷ|triệu) USD", "\\1" + NBSP + "\\2" + NBSP + "USD", text)
    text = re.sub(r"(\d) ngày", "\\1" + NBSP + "ngày", text)
    return re.sub(r"(\d) điểm %", "\\1" + NBSP + "điểm" + NBSP + "%", text)


def for_pbi(text):
    """Web wording that does not hold in the report: the grey band of the quarterly charts, the hollow fourth-quarter
    points, the totals printed at the end of the web's Digital bars, and where the list of reports is."""
    text = re.sub(r" ?Vùng tô xám là quý 4 FY2020[^.]*\.", "", text)
    text = text.replace("Điểm rỗng (quý 4) là ước tính.", "Số quý 4 là ước tính.")
    text = text.replace(" Nhãn cuối thanh là tổng NIKE Direct.", "")
    return text.replace("(danh sách ở cuối trang)", f"(danh sách ở trang “{SRC_PAGE}”)")


def region_title(r):
    return r["vi"] if r["vi"] == r["short"] else f"{r['vi']} ({r['short']})"


# ------------------------------------------------------------------ vi-VN formatting (as app.js)
def nf(v, d):
    q = Decimal(abs(v)).quantize(Decimal(1).scaleb(-d), rounding=ROUND_HALF_UP)
    return f"{q:,.{d}f}".replace(",", "§").replace(".", ",").replace("§", ".")


def num(v, d=0):
    return "—" if v is None else (MINUS if v < 0 and nf(v, d).strip("0,.") else "") + nf(v, d)


def signed(v, d=0):
    s = nf(v, d)
    return ("" if not s.strip("0,.") else "+" if v > 0 else MINUS) + s


def stated(v):
    s = f"{round(v, 3):.3f}".rstrip("0")
    return num(v, max(1, len(s.split(".")[1])))


FORMATS = {
    "amount": lambda v: "(" + nf(v, 0) + ")" if v < 0 else nf(v, 0),
    "pct1": lambda v: num(v * 100, 1) + "%",
    "pct0": lambda v: num(v * 100, 0) + "%",
    "spct1": lambda v: signed(v * 100, 1) + "%",
    "spct0": lambda v: signed(v * 100, 0) + "%",
    "pts1": lambda v: signed(v * 100, 1),
    "stated": stated,
    "num0": lambda v: num(v, 0),
}

# Table rows mirror the row lists in site/assets/app.js: (kind, id, label, unit, format).
CHANNEL_ROWS = [
    ("g", "Doanh thu NIKE Brand theo kênh"),
    ("obs", "nike_direct_revenue", "NIKE Direct", "triệu USD", "amount"),
    ("obs", "nike_brand_wholesale_revenue", "Bán buôn (Wholesale)", "triệu USD", "amount"),
    ("obs", "global_brand_divisions_revenue", "Global Brand Divisions", "triệu USD", "amount"),
    ("obs", "nike_brand_revenue", "Tổng doanh thu NIKE Brand", "triệu USD", "amount"),
    ("g", "Cơ cấu và tăng trưởng"),
    ("kpi", "direct_share", "Tỷ trọng NIKE Direct", "%", "pct1"),
    ("obs", "nike_direct_share_stated", "Tỷ trọng NIKE Direct do NIKE nêu (xấp xỉ)", "%", "pct0"),
    ("kpi", "direct_growth", "Tăng trưởng NIKE Direct, theo USD", "%", "spct1"),
    ("obs", "nike_direct_pct_change_cn", "Tăng trưởng NIKE Direct, loại trừ tỷ giá (non-GAAP)", "%", "spct0"),
    ("kpi", "wholesale_growth", "Tăng trưởng bán buôn, theo USD", "%", "spct1"),
    ("obs", "nike_brand_wholesale_pct_change_cn", "Tăng trưởng bán buôn, loại trừ tỷ giá (non-GAAP)", "%", "spct0"),
    ("kpi", "nike_brand_growth", "Tăng trưởng NIKE Brand, theo USD", "%", "spct1"),
    ("kpi", "direct_contribution", "Trong đó: đóng góp của NIKE Direct", "điểm %", "pts1"),
    ("kpi", "wholesale_contribution", "Trong đó: đóng góp của bán buôn", "điểm %", "pts1"),
    ("g", "NIKE Brand Digital (một phần của NIKE Direct)"),
    ("obs", "nike_brand_digital_sales", "Doanh số Digital", "tỷ USD", "stated"),
    ("kpi", "digital_share_of_direct", "Digital trong NIKE Direct (xấp xỉ)", "%", "pct0"),
    ("kpi", "direct_ex_digital", "NIKE Direct ngoài Digital (ước tính)", "triệu USD", "amount"),
]
FINANCE_ROWS = [
    ("g", "Báo cáo kết quả kinh doanh (NIKE, Inc.)"),
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
    ("kpi", "revenue_growth", "Tăng trưởng doanh thu, theo USD", "%", "spct1"),
    ("obs", "revenues_pct_change_cn", "Tăng trưởng doanh thu, loại trừ tỷ giá (non-GAAP)", "%", "spct0"),
    ("kpi", "gross_margin", "Biên lợi nhuận gộp", "%", "pct1"),
    ("kpi", "demand_creation_ratio", "Chi phí tạo nhu cầu / doanh thu", "%", "pct1"),
    ("kpi", "operating_overhead_ratio", "Chi phí vận hành chung / doanh thu", "%", "pct1"),
    ("kpi", "net_income_margin", "Biên lợi nhuận sau thuế", "%", "pct1"),
    ("g", "EBIT (non-GAAP, định nghĩa của NIKE)"),
    ("obs", "ebit_reported", "EBIT do NIKE công bố", "triệu USD", "amount"),
    ("kpi", "ebit_margin", "Biên EBIT", "%", "pct1"),
    ("g", "Hàng tồn kho"),
    ("obs", "inventories", "Hàng tồn kho cuối kỳ", "triệu USD", "amount"),
    ("kpi", "inventory_growth", "So với cùng thời điểm năm trước", "%", "spct1"),
    ("kpi", "days_inventory", "Số ngày tồn kho bình quân", "ngày", "num0"),
]


def region_rows(regions):
    rows = []
    for r in regions:
        rid = r["id"]
        rows += [("g", region_title(r)),
                 ("obs", f"{rid}_revenue", "Doanh thu", "triệu USD", "amount"),
                 ("obs", f"{rid}_direct_revenue", "NIKE Direct", "triệu USD", "amount"),
                 ("obs", f"{rid}_wholesale_revenue", "Bán buôn", "triệu USD", "amount"),
                 ("kpi", f"{rid}_direct_share", "Tỷ trọng NIKE Direct", "%", "pct1"),
                 ("kpi", f"{rid}_growth", "Tăng trưởng doanh thu, theo USD", "%", "spct1"),
                 ("obs", f"{rid}_ebit", "EBIT", "triệu USD", "amount"),
                 ("kpi", f"{rid}_ebit_margin", "Biên EBIT", "%", "pct1")]
    return rows


def table_cell(row, p, d):
    """Figure, page reference and report link of one table cell, as renderTable() in app.js shows them."""
    kind, mid, _, _, fmt = row
    if kind == "kpi":
        v = d["kpis"][mid]["values"][p]
        return ("—" if v is None else FORMATS[fmt](v)), "", None
    o = d["observations"][mid].get(p)
    if not o or o["value"] is None:
        return "Không công bố", "", None
    if o["method"] == "q4_derived":
        return ("≈" + num(o["value"], 1)) if o["approx"] else FORMATS[fmt](o["value"]), "", None
    mark = "†" if o["status"] == STATUS_PENDING else ""
    url = d["filings"][o["source"]]["url"]
    if d["filings"][o["source"]]["form"] == "8-K":   # an earnings release has no page numbers
        return FORMATS[fmt](o["value"]), " ".join(filter(None, [mark, "8-K"])), url
    if not o.get("page"):
        return FORMATS[fmt](o["value"]), mark, None
    return FORMATS[fmt](o["value"]), " ".join(filter(None, [mark, f"tr. {o['page']}"])), url


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
        self.findings_title = text("//section[@id='phat-hien']//h2")
        self.findings_note = text("//section[@id='phat-hien']//p[@class='block-note']")
        # The terms line as (bold, text) runs: the names are bold, as on the web page.
        terms = one("//p[@id='terms']")
        flat = lambda s: re.sub(r"\s+", " ", s or "")
        runs = [(False, flat(terms.text))] + [r for el in terms for r in ((el.tag == "b", flat(el.text_content())), (False, flat(el.tail)))]
        self.terms = [(bold, t) for bold, t in runs if t]
        self.sections = {s: (text(f"//section[@id='{s}']//p[@class='block-kicker']"), text(f"//section[@id='{s}']//h2"))
                         for s in ("kenh", "quy", "digital", "khu-vuc", "loi-nhuan", "nike-giai-thich")}
        self.quotes_note = text("//section[@id='nike-giai-thich']//p[@class='block-note']")
        self.fig = {}
        for f in doc.xpath("//figure[@id]"):
            sub = f.xpath(".//p[contains(@class,'figure-sub')]")
            self.fig[f.get("id")] = (" ".join(f.xpath(".//h3")[0].text_content().split()),
                                     for_pbi(" ".join(sub[0].text_content().split())) if sub else "")
        aside = one("//aside[@aria-labelledby='notes-title']")
        self.notes = (" ".join(aside.xpath(".//h3")[0].text_content().split()),
                      [(" ".join(dt.text_content().split()), " ".join(dd.text_content().split()))
                       for dt, dd in zip(aside.xpath(".//dt"), aside.xpath(".//dd"))])
        self.footer_title = text("//footer//h2")
        self.footer_heads = [" ".join(h.text_content().split()) for h in doc.xpath("//footer//h3")]
        self.method = [" ".join(li.text_content().split()) for li in doc.xpath("//footer//ul[@class='plain-list']/li")]
        self.fine = text("//footer//p[@class='fine']")
        self.sources = {}
        tail = re.search(r"function quarterSources\(text\) \{\s*return text \+ '([^']*)'", js).group(1)
        for cid, body in re.findall(r"sourceLine\('(\w+)', (.+?)\);\n", js, flags=re.S):
            m = re.match(r"(quarterSources\()?'([^']*)'", body)
            self.sources[cid] = for_pbi(m.group(2) + (tail if m.group(1) else ""))
        self.subs = {}
        for cid in ("change-sub", "bridge-sub"):
            m = re.search(rf"getElementById\('{cid}'\)\.textContent = '([^']*)' \+ a \+ '([^']*)' \+ b \+\s*'([^']*)'", js)
            if not m:
                raise SystemExit(f"app.js: subtitle {cid} not found")
            self.subs[cid] = m.groups()
        for k in ("share", "change", "shareQ", "channelQ", "digital", "digitalQ", "regions", "regionsEbit", "bridge", "marginsQ",
                  "inventoryQ", "daysQ"):
            if k not in self.sources:
                raise SystemExit(f"app.js: source line for {k} not found")

    def sub(self, cid, a, b):
        p = self.subs[cid]
        return p[0] + a + p[1] + b + p[2]


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


SIGNED = "+0.0%;−0.0%;0.0%"
SIGNED0 = "+0%;−0%;0%"


QSUFFIX = " (quý)"   # measure names are unique across the whole model; visuals show their own display names


def measure_name(m, entity):
    return m + QSUFFIX if entity == "Quy" else m


def measures_for(table):
    s = lambda c: f"SUM ( {table}[{c}] )"
    bn = lambda c: f"DIVIDE ( SUM ( {table}[{c}] ), 1000 )"
    return [dict(m, name=measure_name(m["name"], table)) for m in [
        dict(name="Tỷ trọng NIKE Direct", expr=s("direct_share"), format="0.0%"),
        dict(name="Tỷ trọng bán buôn", expr=s("wholesale_share"), format="0.0%"),
        dict(name="Tỷ trọng Global Brand Divisions", expr=s("gbd_share"), format="0.00%"),
        dict(name="NIKE Direct (tỷ USD)", expr=bn("nike_direct_revenue"), format="#,0.0"),
        dict(name="Bán buôn (tỷ USD)", expr=bn("nike_brand_wholesale_revenue"), format="#,0.0"),
        dict(name="NIKE Brand (tỷ USD)", expr=bn("nike_brand_revenue"), format="#,0.0"),
        dict(name="Tăng trưởng NIKE Direct", expr=s("direct_growth"), format=SIGNED),
        dict(name="Tăng trưởng bán buôn", expr=s("wholesale_growth"), format=SIGNED),
        dict(name="Digital (tỷ USD)", expr=s("nike_brand_digital_sales"), format="#,0.0"),
        dict(name="NIKE Direct ngoài Digital (tỷ USD)", expr=bn("direct_ex_digital"), format="#,0.0"),
        dict(name="Digital trong NIKE Direct", expr=s("digital_share_of_direct"), format="0%"),
        dict(name="Doanh thu (tỷ USD)", expr=bn("revenues"), format="#,0.0"),
        dict(name="Biên lợi nhuận gộp", expr=s("gross_margin"), format="0.0%"),
        dict(name="Biên EBIT", expr=s("ebit_margin"), format="0.0%"),
        dict(name="Hàng tồn kho (tỷ USD)", expr=bn("inventories"), format="#,0.0"),
        dict(name="So với cùng thời điểm năm trước", expr=s("inventory_growth"), format=SIGNED),
        dict(name="Số ngày tồn kho", expr=s("days_inventory"), format="0"),
        dict(name="Giá vốn hàng bán (tỷ USD)", expr=bn("cost_of_sales"), format="#,0.00"),
    ]]


def semantic_model(d, w, out):
    obs, kpis = d["observations"], d["kpis"]
    A, Q = d["periods"]["A"], d["periods"]["Q"]
    values = list(obs) + list(kpis)

    def period_rows(periods):
        rows = []
        for i, p in enumerate(periods):
            r = {"Kỳ": d["labels"][p], "Thứ tự": i, "Năm": p[:6], "Quý": ("Q" + p[7]) if "Q" in p else p,
                 "Thời gian": d["dates"][p]}
            r.update({k: (obs[k].get(p) or {}).get("value") for k in obs})
            r.update({k: kpis[k]["values"][p] for k in kpis})
            rows.append(r)
        return rows

    base = [col("Kỳ", "string", sort_by="Thứ tự"), col("Thứ tự", "int64", hidden=True), col("Năm", "string", sort_by="Thứ tự"),
            col("Quý", "string", sort_by="Thứ tự"), col("Thời gian", "string")]
    cols = base + [col(k, "double", hidden=True) for k in values]
    tables = {"Nam": tmdl_table("Nam", cols, period_rows(A), measures_for("Nam")),
              "Quy": tmdl_table("Quy", cols, period_rows(Q), measures_for("Quy"))}

    # Change of NIKE Brand revenue by channel, FY2019 -> FY2021 (one measure per colour).
    a, z = A[0], A[-1]
    items = [("NIKE Direct", "nike_direct_revenue", "Direct"), ("Bán buôn", "nike_brand_wholesale_revenue", "Bán buôn"),
             ("Global Brand Divisions", "global_brand_divisions_revenue", "GBD"), ("Toàn NIKE Brand", "nike_brand_revenue", "Tổng")]
    # The figure goes into the row name ("NIKE Direct  +4,6"): labels inside stacked bars would be unreadable or dropped.
    delta = lambda m: (obs[m][z]["value"] - obs[m][a]["value"]) / 1000
    ch_rows = [{"Kênh": f"{n}  {signed(delta(m), 2 if abs(delta(m)) < 0.1 else 1)}", "Thứ tự": i, "Nhóm": g, "Thay đổi": delta(m)}
               for i, (n, m, g) in enumerate(items)]
    ch_measures = [dict(name=f"Thay đổi – {g}", expr=f'CALCULATE ( SUM ( ThayDoi[Thay đổi] ), ThayDoi[Nhóm] = "{g}" )',
                        format='+#,0.00;−#,0.00;0.00') for _, _, g in items]
    tables["ThayDoi"] = tmdl_table("ThayDoi", [col("Kênh", "string", sort_by="Thứ tự"), col("Thứ tự", "int64", hidden=True),
                                               col("Nhóm", "string", hidden=True), col("Thay đổi", "double", hidden=True)],
                                   ch_rows, ch_measures)

    # EBIT margin bridge, FY2019 -> FY2021, in percentage points.
    k = lambda kid, p: kpis[kid]["values"][p]
    steps = [("Biên gộp", k("gross_margin", z) - k("gross_margin", a)),
             ("Tạo nhu cầu", -(k("demand_creation_ratio", z) - k("demand_creation_ratio", a))),
             ("Vận hành chung", -(k("operating_overhead_ratio", z) - k("operating_overhead_ratio", a)))]
    start, end = k("ebit_margin", a), k("ebit_margin", z)
    steps.append(("Khác", end - start - sum(s for _, s in steps)))
    # As for the change chart, each step's figure is part of its name.
    pct_ = lambda v: num(v * 100, 1) + "%"
    br = [{"Bước": f"{a} {pct_(start)}", "Thứ tự": 0, "Nền": 0.0, "Tổng": start * 100, "Tăng": None, "Giảm": None}]
    running = start
    for i, (n, s) in enumerate(steps, 1):
        br.append({"Bước": f"{n} {signed(s * 100, 1)}", "Thứ tự": i, "Nền": (running if s >= 0 else running + s) * 100,
                   "Tổng": None, "Tăng": s * 100 if s >= 0 else None, "Giảm": -s * 100 if s < 0 else None})
        running += s
    br.append({"Bước": f"{z} {pct_(end)}", "Thứ tự": len(steps) + 1, "Nền": 0.0, "Tổng": end * 100, "Tăng": None, "Giảm": None})
    br_measures = [dict(name="Cầu nối – nền", expr="SUM ( CauNoi[Nền] )", format="0.0"),
                   dict(name="Cầu nối – biên EBIT", expr="SUM ( CauNoi[Tổng] )", format='0.0"%"'),
                   dict(name="Cầu nối – làm biên tăng", expr="SUM ( CauNoi[Tăng] )", format='"+"0.0'),
                   dict(name="Cầu nối – làm biên giảm", expr="SUM ( CauNoi[Giảm] )", format='"−"0.0')]
    tables["CauNoi"] = tmdl_table("CauNoi", [col("Bước", "string", sort_by="Thứ tự"), col("Thứ tự", "int64", hidden=True)]
                                  + [col(c, "double", hidden=True) for c in ("Nền", "Tổng", "Tăng", "Giảm")], br, br_measures)

    # Regions: FY2019 and FY2021 side by side.
    regs = [("Toàn NIKE Brand", "direct_share", None)] + [(r["short"], f"{r['id']}_direct_share", f"{r['id']}_ebit_margin")
                                                          for r in d["regions"]]
    kv_rows = [{"Khu vực": n, "Thứ tự": i, "Direct " + a: k(sh, a), "Direct " + z: k(sh, z),
                "EBIT " + a: k(eb, a) if eb else None, "EBIT " + z: k(eb, z) if eb else None}
               for i, (n, sh, eb) in enumerate(regs)]
    kv_cols = [col("Khu vực", "string", sort_by="Thứ tự"), col("Thứ tự", "int64", hidden=True)] + [
        col(c, "double", hidden=True) for c in ("Direct " + a, "Direct " + z, "EBIT " + a, "EBIT " + z)]
    kv_measures = [dict(name=f"Tỷ trọng Direct {p}", expr=f"SUM ( KhuVuc[Direct {p}] )", format="0.0%") for p in (a, z)] + [
        dict(name=f"Biên EBIT khu vực {p}", expr=f"SUM ( KhuVuc[EBIT {p}] )", format="0.0%") for p in (a, z)]
    tables["KhuVuc"] = tmdl_table("KhuVuc", kv_cols, kv_rows, kv_measures)

    # Tables: annual and regional with a page column per year; quarterly without (twelve columns already).
    def cells(name, periods, rows, with_pages, labels):
        out_rows = []
        for order, row in enumerate(rows, 1):
            label = glue(row[1]).upper() if row[0] == "g" else glue(row[2] if row[3] == "%" else f"{row[2]}, {row[3]}")
            for i, p in enumerate(periods):
                figure, page, link = ("", "", None) if row[0] == "g" else table_cell(row, p, d)
                cell = {"Thứ tự": order, "Chỉ tiêu": label, "Kỳ": p}
                step = 2 if with_pages else 1
                out_rows.append({**cell, "Cột": labels(p), "Thứ tự cột": step * i, "Loại": "Số", "Hiển thị": figure, "Liên kết": None})
                if with_pages:
                    # Each period is two matrix columns: the figure, then its page reference under a blank header
                    # (non-breaking spaces, a different count per period so the headers stay distinct).
                    out_rows.append({**cell, "Cột": NBSP * (i + 1), "Thứ tự cột": 2 * i + 1, "Loại": "Trang", "Hiển thị": page,
                                     "Liên kết": link})
        cols_ = [col("Thứ tự", "int64", hidden=True), col("Chỉ tiêu", "string", sort_by="Thứ tự"), col("Kỳ", "string", hidden=True),
                 col("Cột", "string", sort_by="Thứ tự cột"), col("Thứ tự cột", "int64", hidden=True), col("Loại", "string", hidden=True),
                 col("Hiển thị", "string"), col("Liên kết", "string", category="WebUrl")]
        ms = [dict(name=f"Giá trị {name}", expr=f'COALESCE ( SELECTEDVALUE ( {name}[Hiển thị] ), "" )'),
              dict(name=f"Liên kết {name}", expr=f"SELECTEDVALUE ( {name}[Liên kết] )"),
              dict(name=f"Màu chữ {name}", expr=f'IF ( SELECTEDVALUE ( {name}[Loại] ) = "Trang", "{C["faint"]}", "{C["ink"]}" )')]
        return tmdl_table(name, cols_, out_rows, ms)

    tables["BangNam"] = cells("BangNam", A, CHANNEL_ROWS + FINANCE_ROWS, True, lambda p: p)
    tables["BangQuy"] = cells("BangQuy", Q, CHANNEL_ROWS + FINANCE_ROWS, False, lambda p: d["labels"][p])
    tables["BangKhuVuc"] = cells("BangKhuVuc", A, region_rows(d["regions"]), True, lambda p: p)

    # Reports list for the sources page.
    groups = {("10-K", "scope"): (1, "Báo cáo năm (10-K)"), ("10-Q", "scope"): (2, "Báo cáo quý (10-Q)"),
              ("8-K", "q4release"): (3, "Thông cáo kết quả quý 4 (8-K)")}
    bc_rows = []
    for i, (fid, f) in enumerate(d["filings"].items()):
        g = groups.get((f["form"], f["role"]), (4, "Báo cáo các năm sau (chỉ để kiểm tra NIKE có in lại số cũ khác đi không)"))
        bc_rows.append({"Báo cáo": glue(fid), "Thứ tự": g[0] * 100 + i, "Nhóm": g[1],
                        "Ngày nộp SEC": vn_date(f["filed"]), "Số hồ sơ SEC": f["accession"], "Liên kết SEC EDGAR": f["url"]})
    tables["BaoCao"] = tmdl_table("BaoCao", [col("Báo cáo", "string", sort_by="Thứ tự"), col("Thứ tự", "int64", hidden=True),
                                             col("Nhóm", "string"), col("Ngày nộp SEC", "string"), col("Số hồ sơ SEC", "string"),
                                             col("Liên kết SEC EDGAR", "string", category="WebUrl")], bc_rows)

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
    write_text(sm / "definition" / "relationships.tmdl", "")
    for t, body in tables.items():
        write_text(sm / "definition" / "tables" / f"{t}.tmdl", body)
    return tables


def vn_date(iso):
    y, m, d = iso.split("-")
    return f"{d}/{m}/{y}"


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


def measure(prop, entity):
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


def container_objects(title=None, subtitle=None):
    o = {"border": props(show=False), "dropShadow": props(show=False), "background": props(show=False),
         "padding": props(top=0.0, bottom=0.0, left=0.0, right=0.0)}
    if title:
        o["title"] = props(show=True, text=glue(title), fontFamily=SANS, fontSize=13.0, bold=True, fontColor=fill(C["ink"]),
                           alignment="left", titleWrap=True)
    else:
        o["title"] = props(show=False)
    if subtitle:
        o["subTitle"] = props(show=True, text=glue(subtitle), fontFamily=SANS, fontSize=10.0, fontColor=fill(C["muted"]),
                              alignment="left", titleWrap=True)
    return o


class Page:
    def __init__(self, name, display):
        self.name, self.display, self.height, self.visuals, self.z = name, display, 720, [], 0

    def add(self, name, x, y, w, h, visual):
        self.z += 1
        self.visuals.append({"$schema": VC_SCHEMA, "name": name,
                             "position": {"x": x, "y": y, "z": self.z * 1000, "height": h, "width": w, "tabOrder": self.z * 1000},
                             "visual": visual})
        self.height = max(self.height, y + h + 40)

    def text(self, name, x, y, w, h, paragraphs):
        self.add(name, x, y, w, h, {"visualType": "textbox", "objects": {"general": [{"properties": {"paragraphs": paragraphs}}]},
                                    "visualContainerObjects": container_objects(), "drillFilterOtherVisuals": True})


def run(text, size=11, color=C["ink2"], bold=False, family=SANS, url=None, italic=False):
    style = {"fontFamily": family, "fontSize": f"{size}pt", "color": color}
    if bold:
        style["fontWeight"] = "bold"
    if italic:
        style["fontStyle"] = "italic"
    r = {"value": glue(text), "textStyle": style}
    if url:
        r["url"] = url
    return r


def para(*runs):
    return {"textRuns": list(runs), "horizontalTextAlignment": "left"}


def cartesian(kind, entity, axis, series, title, subtitle, tooltips=(), labels=True, value_start=None, value_end=None,
              legend=True, label_density=None, transparent=(), label_off=(), label_colors=()):
    """series: [(measure, display name, colour)]; label_colors: [(measure, label colour)]"""
    mref = lambda m: measure(measure_name(m, entity), entity)
    labels_props = props(show=labels, labelDisplayUnits=1.0, fontFamily=SANS, fontSize=9.0, color=fill(C["ink2"]))[0]["properties"]
    if label_density is not None:
        labels_props["labelDensity"] = lit(float(label_density))
    objects = {
        "categoryAxis": props(showAxisTitle=False, fontFamily=SANS, fontSize=10.0, labelColor=fill(C["muted"]), concatenateLabels=False),
        "valueAxis": props(showAxisTitle=False, fontFamily=SANS, fontSize=10.0, labelColor=fill(C["muted"]),
                           gridlineShow=True, gridlineColor=fill(C["grid"]),
                           **({"start": lit(float(value_start))} if value_start is not None else {}),
                           **({"end": lit(float(value_end))} if value_end is not None else {})),
        "legend": props(show=legend, position="Top", showTitle=False, fontFamily=SANS, fontSize=10.0, labelColor=fill(C["ink2"])),
        "labels": [{"properties": labels_props}] + [{"properties": {"show": lit(False)}, "selector": {"metadata": ref(mref(m))}}
                                                    for m in label_off]
                  + [{"properties": {"color": fill(c)}, "selector": {"metadata": ref(mref(m))}} for m, c in label_colors],
        "dataPoint": [{"properties": {"fill": fill(series[0][2])}}]
                     + [{"properties": {"fill": fill(c), **({"fillTransparency": lit(100.0)} if m in transparent else {})},
                         "selector": {"metadata": ref(mref(m))}} for m, _, c in series],
    }
    if kind == "lineChart":
        objects["lineStyles"] = props(lineChartType="linear", showMarker=True, strokeWidth=2)
    if kind == "hundredPercentStackedBarChart":
        objects["categoryAxis"][0]["properties"]["innerPadding"] = lit(25.0)
        objects["labels"][0]["properties"].update(color=fill("#FFFFFF"), labelPrecision=lit(1))
    query = {"Category": {"projections": [proj(axis)]},
             "Y": {"projections": [proj(mref(m), d) for m, d, _ in series]}}
    if tooltips:
        query["Tooltips"] = {"projections": [proj(mref(m), d) for m, d in tooltips]}
    sort = {"sort": [{"field": axis, "direction": "Ascending"}], "isDefaultSort": False}
    return {"visualType": kind, "query": {"queryState": query, "sortDefinition": sort}, "objects": objects,
            "visualContainerObjects": container_objects(title, subtitle), "drillFilterOtherVisuals": True}


def matrix(table, wide=False):
    value = measure(f"Giá trị {table}", table)
    every_cell = {"data": [{"dataViewWildcard": {"matchingOption": 1}}], "metadata": ref(value)}
    return {"visualType": "pivotTable",
            "query": {"queryState": {"Rows": {"projections": [proj(column(table, "Chỉ tiêu"))]},
                                     "Columns": {"projections": [proj(column(table, "Cột"))]},
                                     "Values": {"projections": [proj(value, "Giá trị")]}}},
            "objects": {"subTotals": props(rowSubtotals=False, columnSubtotals=False),
                        "rowHeaders": props(fontFamily=SANS, fontSize=9.5 if wide else 10.0, fontColor=fill(C["ink"]), wordWrap=True),
                        "columnHeaders": props(fontFamily=SANS, fontSize=9.5 if wide else 10.0, fontColor=fill(C["ink"]), bold=True,
                                               alignment="Right", titleAlignment="Left", wordWrap=True),
                        "values": [{"properties": {"fontFamily": lit(SANS), "fontSize": lit(9.5 if wide else 10.0),
                                                   "fontColor": fill(C["ink"])}},
                                   # page references: muted, and Ctrl + click opens the report
                                   {"properties": {"fontColor": {"solid": {"color": {"expr": measure(f"Màu chữ {table}", table)}}},
                                                   "webURL": {"expr": measure(f"Liên kết {table}", table)}},
                                    "selector": every_cell}],
                        "columnFormatting": [{"properties": {"alignment": lit("Right")}, "selector": {"metadata": ref(value)}}],
                        "grid": props(gridVertical=False, gridHorizontal=True, gridHorizontalColor=fill(C["grid"]),
                                      rowPadding=3.0)},
            "visualContainerObjects": container_objects(), "drillFilterOtherVisuals": True}


# ------------------------------------------------------------------ pages
def header(page, w):
    page.text("tieu_de", MARGIN, 16, 900, 70, [para(run(w.title, 20, C["ink"], True, SERIF))])
    page.text("mo_ta", MARGIN, 88, 900, 60, [para(run(w.subtitle, 10.5, C["ink2"]))])
    page.text("lien_ket", 960, 20, 280, 70, [
        para(run("Mở dashboard web", 10.5, C["direct"], url=DASHBOARD_URL)),
        para(run("Tải bộ dữ liệu Excel", 10.5, C["direct"], url=DASHBOARD_URL + "downloads/NIKE_Financial_Dataset_FY2019-2021.xlsx"))])
    return 150


def section(page, w, key, sid, y, note=None):
    kicker, question = w.sections[sid]
    page.text(f"{key}_tieu_de", MARGIN, y, 1200, 64, [para(run(kicker, 10, C["muted"], True)),
                                                      para(run(question, 16, C["ink"], True, SERIF))])
    y += 66
    if note:
        page.text(f"{key}_ghi_chu", MARGIN, y, 900, 46, [para(run(note, 11, C["ink2"]))])
        y += 50
    return y + 6


def source(page, name, x, y, w_, text, h=46):
    page.text(name, x, y, w_, h, [para(run(text, 9, C["muted"]))])


def page_overview(w, d):
    p = Page("tong_quan", "Tổng quan")
    y = header(p, w)
    p.text("phat_hien_tieu_de", MARGIN, y, 1200, 100, [
        para(run(w.findings_title, 16, C["ink"], True, SERIF)), para(run(w.findings_note, 10.5, C["ink2"])),
        para(*[run(t, 9.5, C["ink"] if bold else C["ink2"], bold) for bold, t in w.terms])])
    y += 108
    for i, f in enumerate(d["findings"]):
        row, c = divmod(i, 3)
        x = MARGIN + c * 410
        top = y + row * 230
        p.text(f"phat_hien_{i + 1}", x, top, 380, 220, [
            para(run(f["label"], 10, C["muted"], True)), para(run(f["value"], 24, C["ink"], True)),
            para(run(f["text"], 10.5, C["ink2"])), para(run(" ", 4)), para(run(f["source"], 8.5, C["muted"]))])
    y += 470
    y = section(p, w, "kenh", "kenh", y)
    a, z = d["periods"]["A"][0], d["periods"]["A"][-1]
    t, s = w.fig["fig-share"]
    p.add("co_cau", L, y, CW, 300, cartesian(
        "hundredPercentStackedBarChart", "Nam", column("Nam", "Kỳ"),
        [("Tỷ trọng NIKE Direct", "NIKE Direct", C["direct"]), ("Tỷ trọng bán buôn", "Bán buôn (Wholesale)", C["wholesale"]),
         ("Tỷ trọng Global Brand Divisions", "Global Brand Divisions (dưới 0,2%)", C["gbd"])],
        t, s, tooltips=[("NIKE Direct (tỷ USD)", "NIKE Direct, tỷ USD"), ("Bán buôn (tỷ USD)", "Bán buôn, tỷ USD"),
                        ("NIKE Brand (tỷ USD)", "Doanh thu NIKE Brand (mẫu số), tỷ USD")]))
    source(p, "nguon_co_cau", L, y + 304, CW, w.sources["share"] + " Báo cáo: " + ", ".join(f"10-K {x}" for x in d["periods"]["A"]) + ".")
    t, _ = w.fig["fig-change"]
    groups = [("Thay đổi – Direct", "NIKE Direct", C["direct"]), ("Thay đổi – Bán buôn", "Bán buôn", C["wholesale"]),
              ("Thay đổi – GBD", "Global Brand Divisions", C["gbd"]), ("Thay đổi – Tổng", "Toàn NIKE Brand", C["total"])]
    # Stacked, not clustered: each row has one non-blank measure (its colour), so its bar sits centred on the row.
    v = cartesian("barChart", "ThayDoi", column("ThayDoi", "Kênh"), groups, t, w.sub("change-sub", a, z), legend=False,
                  labels=False)
    v["objects"]["categoryAxis"][0]["properties"]["innerPadding"] = lit(30.0)
    p.add("thay_doi", R, y, CW, 300, v)
    source(p, "nguon_thay_doi", R, y + 304, CW, w.sources["change"] + f" Báo cáo: 10-K {a}, 10-K {z}.")
    return p


def page_quarters(w, d):
    p = Page("theo_quy", "Theo quý")
    y = header(p, w)
    y = section(p, w, "quy", "quy", y, d["notes"]["quy"])
    axis = column("Quy", "Kỳ")
    t, s = w.fig["fig-share-q"]
    p.add("ty_trong_quy", L, y, CW, 320, cartesian(
        "lineChart", "Quy", axis, [("Tỷ trọng NIKE Direct", "Tỷ trọng NIKE Direct", C["direct"])], t, s, legend=False,
        value_start=0.25, value_end=0.5, label_density=0,
        tooltips=[("NIKE Direct (tỷ USD)", "NIKE Direct, tỷ USD"), ("NIKE Brand (tỷ USD)", "NIKE Brand, tỷ USD")]))
    source(p, "nguon_ty_trong_quy", L, y + 324, CW, w.sources["shareQ"])
    t, s = w.fig["fig-channel-q"]
    p.add("kenh_quy", R, y, CW, 320, cartesian(
        "lineChart", "Quy", axis, [("NIKE Direct (tỷ USD)", "NIKE Direct", C["direct"]), ("Bán buôn (tỷ USD)", "Bán buôn (Wholesale)", C["wholesale"])],
        # Two lines with a label on every point would collide; the values are in the tooltips and on page 4.
        t, s, value_start=0.0, labels=False,
        tooltips=[("Tăng trưởng NIKE Direct", "NIKE Direct so với cùng kỳ"), ("Tăng trưởng bán buôn", "Bán buôn so với cùng kỳ")]))
    source(p, "nguon_kenh_quy", R, y + 324, CW, w.sources["channelQ"])
    y += 390
    y = section(p, w, "digital", "digital", y)
    t, s = w.fig["fig-digital"]
    v = cartesian("barChart", "Nam", column("Nam", "Kỳ"),
                  [("Digital (tỷ USD)", "NIKE Brand Digital", C["digital"]),
                   ("NIKE Direct ngoài Digital (tỷ USD)", "NIKE Direct ngoài Digital (ước tính)", C["directLight"])],
                  t, s, tooltips=[("NIKE Direct (tỷ USD)", "NIKE Direct (tổng), tỷ USD"), ("Digital trong NIKE Direct", "Digital trong NIKE Direct (≈)")],
                  label_colors=[("Digital (tỷ USD)", "#FFFFFF"), ("NIKE Direct ngoài Digital (tỷ USD)", C["ink"])])
    v["objects"]["categoryAxis"][0]["properties"]["innerPadding"] = lit(30.0)
    p.add("digital", L, y, CW, 300, v)
    source(p, "nguon_digital", L, y + 304, CW, w.sources["digital"] + " Báo cáo: " + ", ".join(f"10-K {x}" for x in d["periods"]["A"]) + ".")
    t, s = w.fig["fig-digital-q"]
    p.add("digital_quy", R, y, CW, 300, cartesian(
        "lineChart", "Quy", axis, [("Digital trong NIKE Direct", "Digital / NIKE Direct", C["digital"])], t, s, legend=False,
        value_start=0.0, value_end=0.8, label_density=0,
        tooltips=[("Digital (tỷ USD)", "Digital, tỷ USD (quý 4: ước tính)"), ("NIKE Direct (tỷ USD)", "NIKE Direct, tỷ USD")]))
    source(p, "nguon_digital_quy", R, y + 304, CW, w.sources["digitalQ"])
    return p


def page_profit(w, d):
    p = Page("khu_vuc_loi_nhuan", "Khu vực và lợi nhuận")
    y = header(p, w)
    a, z = d["periods"]["A"][0], d["periods"]["A"][-1]
    y = section(p, w, "khu_vuc", "khu-vuc", y)
    axis = column("KhuVuc", "Khu vực")
    t, s = w.fig["fig-regions"]
    p.add("khu_vuc", L, y, CW, 340, cartesian(
        "clusteredBarChart", "KhuVuc", axis, [(f"Tỷ trọng Direct {a}", a, C["past"]), (f"Tỷ trọng Direct {z}", z, C["direct"])], t, s))
    source(p, "nguon_khu_vuc", L, y + 344, CW, w.sources["regions"] + f" Báo cáo: 10-K {a}, 10-K {z}.")
    t, s = w.fig["fig-regions-ebit"]
    v = cartesian("clusteredBarChart", "KhuVuc", axis, [(f"Biên EBIT khu vực {a}", a, C["past"]), (f"Biên EBIT khu vực {z}", z, C["direct"])], t, s)
    p.add("khu_vuc_ebit", R, y, CW, 340, v)
    # The NIKE Brand row has no EBIT of its own; a visual-level filter (on the container) keeps the four regions.
    p.visuals[-1]["filterConfig"] = {"filters": [{
        "name": "chi_khu_vuc", "field": column("KhuVuc", "Khu vực"), "type": "Categorical",
        "filter": {"Version": 2, "From": [{"Name": "k", "Entity": "KhuVuc", "Type": 0}],
                   "Where": [{"Condition": {"Not": {"Expression": {"In": {
                       "Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "k"}}, "Property": "Khu vực"}}],
                       "Values": [[{"Literal": {"Value": "'Toàn NIKE Brand'"}}]]}}}}}]},
        "howCreated": "User"}]}
    source(p, "nguon_khu_vuc_ebit", R, y + 344, CW, w.sources["regionsEbit"] + f" Báo cáo: 10-K {a}, 10-K {z}.")
    y += 410
    y = section(p, w, "loi_nhuan", "loi-nhuan", y)
    t, _ = w.fig["fig-ebit-bridge"]
    # Figures are in the step names and the colours are explained in the subtitle, so neither labels nor a legend
    # (which would list the invisible base series) are drawn.
    v = cartesian("columnChart", "CauNoi", column("CauNoi", "Bước"),
                  [("Cầu nối – nền", "Nền", "#FFFFFF"), ("Cầu nối – biên EBIT", "Biên EBIT", C["total"]),
                   ("Cầu nối – làm biên tăng", "Làm biên EBIT tăng", C["up"]), ("Cầu nối – làm biên giảm", "Làm biên EBIT giảm", C["down"])],
                  t, w.sub("bridge-sub", a, z) + " Cột xám là biên EBIT, xanh lá làm biên tăng, đỏ làm biên giảm.",
                  transparent=("Cầu nối – nền",), labels=False, legend=False, value_start=0.0)
    v["objects"]["categoryAxis"][0]["properties"]["innerPadding"] = lit(30.0)
    p.add("cau_noi", L, y, CW, 320, v)
    source(p, "nguon_cau_noi", L, y + 324, CW, w.sources["bridge"] + f" Báo cáo: 10-K {a}, 10-K {z}.")
    t, s = w.fig["fig-margins-q"]
    p.add("bien_quy", R, y, CW, 320, cartesian(
        "lineChart", "Quy", column("Quy", "Kỳ"), [("Biên lợi nhuận gộp", "Biên lợi nhuận gộp", C["total"]),
                                                 ("Biên EBIT", "Biên EBIT (non-GAAP)", C["margin"])], t, s, labels=False,
        value_start=-0.2, value_end=0.6, tooltips=[("Doanh thu (tỷ USD)", "Doanh thu NIKE, Inc., tỷ USD")]))
    source(p, "nguon_bien_quy", R, y + 324, CW, w.sources["marginsQ"])
    y += 380
    t, s = w.fig["fig-inventory-q"]
    p.add("ton_kho", L, y, CW, 280, cartesian(
        "clusteredColumnChart", "Quy", column("Quy", "Kỳ"), [("Hàng tồn kho (tỷ USD)", "Hàng tồn kho", C["total"])], t, s, legend=False,
        label_density=40, tooltips=[("So với cùng thời điểm năm trước", "So với cùng thời điểm năm trước")]))
    source(p, "nguon_ton_kho", L, y + 284, CW, w.sources["inventoryQ"])
    t, s = w.fig["fig-days-q"]
    p.add("so_ngay", R, y, CW, 280, cartesian(
        "lineChart", "Quy", column("Quy", "Kỳ"), [("Số ngày tồn kho", "Số ngày tồn kho", C["total"])], t, s, legend=False,
        value_start=60.0, value_end=180.0, label_density=0,
        tooltips=[("Hàng tồn kho (tỷ USD)", "Hàng tồn kho cuối quý, tỷ USD"), ("Giá vốn hàng bán (tỷ USD)", "Giá vốn hàng bán, tỷ USD")]))
    source(p, "nguon_so_ngay", R, y + 284, CW, w.sources["daysQ"])
    y += 340
    p.text("ton_kho_ghi_chu", MARGIN, y, 1200, 46, [para(run(d["notes"]["ton-kho"], 11, C["ink2"]))])
    y += 60
    title, items = w.notes
    p.text("cach_doc_tieu_de", MARGIN, y, 1200, 28, [para(run(title, 13, C["ink"], True))])
    for i, (dt, dd) in enumerate(items):
        row, c = divmod(i, 3)
        p.text(f"cach_doc_{i + 1}", MARGIN + c * 410, y + 32 + row * 130, 380, 124,
               [para(run(dt, 10.5, C["ink"], True)), para(run(dd, 10, C["ink2"]))])
    return p


def page_tables(w, d):
    p = Page("bang_so_lieu", "Bảng số liệu")
    p.text("tieu_de", MARGIN, 16, 1200, 40, [para(run("Bảng số liệu", 20, C["ink"], True, SERIF))])
    # Power BI tables cannot set italics per cell, so the key says in words which figures the team calculated.
    p.text("chu_thich", MARGIN, 58, 1200, 64, [para(run(text, 9.5, C["muted"])) for text in (
        "Số NIKE công bố có số trang ở cột nhỏ bên phải (Ctrl + bấm vào số trang để mở báo cáo trên SEC); ô không có số trang "
        "là số nhóm tự tính. Bảng theo quý không có cột số trang (xem bản web hoặc bộ dữ liệu Excel).",
        "Quý 4: không có 10-Q, nên nhóm lấy cả năm trừ 9 tháng; ≈ là ước tính từ số NIKE đã làm tròn. Tỷ lệ % quý 4 lấy từ "
        "thông cáo kết quả quý 4 (ghi “8-K”).",
        "Số tiền âm để trong ngoặc đơn, ví dụ (790); tỷ lệ âm có dấu trừ.")])
    n = len(CHANNEL_ROWS + FINANCE_ROWS)
    y = 130
    for name, title, rows, wide in (("BangNam", "Theo năm, FY2019–FY2021", n, False), ("BangQuy", "Theo quý, Q1 FY2019 – Q4 FY2021", n, True),
                                    ("BangKhuVuc", "Bốn khu vực của NIKE Brand, FY2019–FY2021", len(region_rows(d["regions"])), False)):
        p.text(f"{name}_tieu_de", MARGIN, y, 1200, 30, [para(run(title, 13, C["ink"], True))])
        h = 40 + 25 * rows
        p.add(name, MARGIN, y + 32, 1200, h, matrix(name, wide))
        y += 32 + h + 30
    return p


def page_sources(w, d):
    p = Page("nguon", SRC_PAGE)
    p.text("tieu_de", MARGIN, 16, 1200, 40, [para(run(SRC_PAGE, 20, C["ink"], True, SERIF))])
    kicker, question = w.sections["nike-giai-thich"]
    p.text("trich_tieu_de", MARGIN, 64, 1200, 50, [para(run(question, 13, C["ink"], True)), para(run(w.quotes_note, 9.5, C["muted"]))])
    y = 120
    for i, q in enumerate(d["quotes"]):
        row, c = divmod(i, 2)
        p.text(f"trich_{i + 1}", MARGIN + c * 620, y + row * 225, 580, 215, [
            para(run(q["topic"], 9.5, C["muted"], True)), para(run("“" + q["text"] + "”", 10.5, C["ink"], italic=True)),
            para(run(q["vi"], 10, C["ink2"])),
            para(run(f"NIKE, {q['filing']}, tr. {q['page']} · ", 8.5, C["muted"]), run("mở trên SEC", 8.5, C["direct"], url=q["url"]))])
    y += ((len(d["quotes"]) + 1) // 2) * 225 + 10
    p.text("bao_cao_tieu_de", MARGIN, y, 1200, 30, [para(run(w.footer_heads[0], 13, C["ink"], True))])
    fields = [column("BaoCao", c) for c in ("Nhóm", "Báo cáo", "Ngày nộp SEC", "Số hồ sơ SEC", "Liên kết SEC EDGAR")]
    p.add("bao_cao", MARGIN, y + 32, 1200, 560, {
        "visualType": "tableEx",
        "query": {"queryState": {"Values": {"projections": [proj(f) for f in fields]}},
                  "sortDefinition": {"sort": [{"field": fields[1], "direction": "Ascending"}], "isDefaultSort": False}},
        "objects": {"values": props(fontFamily=SANS, fontSize=10.0, urlIcon=False),
                    "columnHeaders": props(fontFamily=SANS, fontSize=10.0, bold=True),
                    "total": props(totals=False)},
        "visualContainerObjects": container_objects(), "drillFilterOtherVisuals": True})
    y += 620
    p.text("phuong_phap", MARGIN, y, 740, 440,
           [para(run(w.footer_heads[1], 13, C["ink"], True))] + [para(run("– " + li, 10.5, C["ink2"])) for li in w.method])
    p.text("tinh_trang", 820, y, 420, 300, [
        para(run(w.footer_heads[2], 13, C["ink"], True)), para(run(status_sentence(d), 10.5, C["ink2"])),
        para(run(w.fine, 9.5, C["muted"])), para(run(" ", 6)),
        para(run("Bản web của dashboard: ", 10.5, C["ink2"]), run(DASHBOARD_URL, 10.5, C["direct"], url=DASHBOARD_URL))])
    return p


def status_sentence(d):
    """Same wording as renderFooter() in app.js: only statuses that occur are listed."""
    c = d["meta"]["status_counts"]
    missing = [p for m in d["observations"].values() for p, o in m.items() if o["value"] is None]
    all_q4 = bool(missing) and all(p.endswith("Q4") for p in missing)
    parts = [(c.get("Khớp nguồn chính thức", 0), " số đã được kiểm tra lại và khớp"),
             (c.get("Đã xác minh (người)", 0), " số đã được nhóm kiểm tra trên báo cáo"),
             (c.get(STATUS_PENDING, 0), " số chưa được kiểm tra (†)"),
             (c.get("Thiếu / không công bố", 0), " ô là tỷ lệ % NIKE không nêu cho riêng quý 4" if all_q4 else " ô NIKE không công bố")]
    shown = ", ".join(f"{n}{t}" for n, t in parts if n)
    return f"{sum(c.values())} số liệu: {shown}. Dữ liệu lấy ngày {vn_date(d['meta']['extracted_on'])}."


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
    pages = [page_overview(w, d), page_quarters(w, d), page_profit(w, d), page_tables(w, d), page_sources(w, d)]
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
    return pages


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
    stale = sorted({p.parent for p in (out / f"{NAME}.Report" / "definition" / "pages").glob("*/page.json")}) if out.exists() else []
    stale_tables = sorted((out / f"{NAME}.SemanticModel" / "definition" / "tables").glob("*.tmdl")) if out.exists() else []
    tables = semantic_model(d, w, out)
    pages = report(d, w, out)
    write_json(out / f"{NAME}.pbip", {"$schema": SCHEMA + "pbip/pbipProperties/1.0.0/schema.json", "version": "1.0",
                                      "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
                                      "settings": {"enableAutoRecovery": True}})
    write_text(out / ".gitignore", "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n*.pbix\n")
    print(f"Power BI project written: {out / (NAME + '.pbip')}")
    current_pages = {out / f"{NAME}.Report" / "definition" / "pages" / p.name for p in pages}
    current_visuals = {out / f"{NAME}.Report" / "definition" / "pages" / p.name / "visuals" / v["name"] for p in pages for v in p.visuals}
    leftovers = [str(s) for s in stale if s not in current_pages]
    leftovers += [str(v) for p in current_pages for v in (p / "visuals").glob("*") if v not in current_visuals]
    leftovers += [str(t) for t in stale_tables if t.stem not in tables]
    if leftovers:
        print("NOTE: files from an earlier build are still on disk (Power BI would load them; remove them before opening):",
              *leftovers, sep="\n  ")
    return 0


if __name__ == "__main__":
    sys.exit(main())
