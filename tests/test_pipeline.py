"""Data-integrity tests: extraction, workbook, validation, KPI formulas and the three dashboard versions."""
import json
import posixpath
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from decimal import Decimal

import pytest
from openpyxl import load_workbook

from pipeline import dataset
from pipeline.build_site_data import read_observations, validate
from pipeline.build_workbook import FIELD_COL, HEADER_ROW, SHEET_CH, SHEET_FIN, WORKBOOK
from pipeline.metrics import REPORTED, STATUS_CONFLICT, STATUS_HUMAN, STATUS_MATCHED

# Representative values re-read by hand from the filings (early and late years, both scopes).
GOLDEN = {
    (2020, "revenues"): 37403, (2020, "nike_direct_revenue"): 12382, (2020, "nike_brand_wholesale_revenue"): 23156,
    (2020, "inventories"): 7367, (2021, "ebit_reported"): 6923, (2022, "nike_direct_revenue"): 18726,
    (2023, "nike_brand_wholesale_revenue"): 27397, (2024, "inventories"): 7519, (2025, "demand_creation_expense"): 4689,
    (2026, "revenues"): 46398, (2026, "nike_direct_revenue"): 17720, (2026, "nike_brand_wholesale_revenue"): 27453,
    (2026, "nike_brand_revenue"): 45222, (2026, "net_income"): 3108, (2026, "interest_expense_net"): -50,
    (2023, "nike_brand_digital_sales"): Decimal("12.6"),
}


@pytest.fixture(scope="module")
def facts():
    return dataset.load_facts()


@pytest.fixture(scope="module")
def observations(facts):
    obs = dataset.build_observations(facts, dataset.load_sec_check())
    results = dataset.compute(obs)
    dataset.apply_cross_check_statuses(obs, results)
    dataset.apply_page_checks(obs, dataset.load_page_check())
    return obs, results


def _mdna_only(o):
    return o["metric"]["id"].endswith("_pct_change_cn") or o["metric"]["id"] == "nike_brand_digital_sales"


def test_every_metric_and_year_present_once(observations):
    obs, _ = observations
    keys = [o["key"] for o in obs]
    assert len(keys) == len(set(keys)) == len(REPORTED) * len(dataset.YEARS)
    missing = [o["key"] for o in obs if o["value"] is None]
    assert missing == ["FY2026|nike_direct_share_stated"]


@pytest.mark.parametrize("key,value", sorted(GOLDEN.items()))
def test_golden_values(observations, key, value):
    obs, _ = observations
    o = next(o for o in obs if (o["fy"], o["metric"]["id"]) == key)
    assert o["value"] == value


def test_sec_api_cross_check_all_match():
    results = dataset.load_sec_check()
    assert results and set(results.values()) == {"match"}


def test_only_disclosed_restatement_is_digital_fy2023(observations):
    obs, _ = observations
    restated = [o["key"] for o in obs if o["restatement"].startswith("CÓ THAY ĐỔI")]
    assert restated == ["FY2023|nike_brand_digital_sales"]


def test_reconciliations_and_comparisons_within_tolerance(observations):
    from pipeline.metrics import CHECKS, COMPARISONS
    _, results = observations
    for c in CHECKS + COMPARISONS:
        for y in dataset.YEARS:
            v = results[y][c["id"]]
            if v is not None:
                assert abs(v) <= Decimal(str(c["tol"])), (c["id"], y, v)


def test_kpi_formulas_against_hand_calculation(observations):
    _, r = observations
    assert r[2020]["gross_margin"] == Decimal(16241) / Decimal(37403)
    assert r[2026]["direct_share"] == Decimal(17720) / Decimal(45222)
    assert r[2021]["direct_growth"] == Decimal(16370) / Decimal(12382) - 1
    assert r[2026]["ebit"] == 3108 + 792 - 50
    assert r[2020]["average_inventory"] == (Decimal(7367) + Decimal(5622)) / 2
    assert r[2024]["days_inventory"] == 365 / (Decimal(28475) / ((Decimal(7519) + Decimal(8454)) / 2))


def test_missing_values_never_become_zero(observations):
    _, r = observations
    assert r[2026]["cmp_direct_share"] is None
    assert dataset.evaluate("a / b", {"a": Decimal(1), "b": None}) is None


def test_mdna_only_values_are_matched_only_through_the_page_check(observations):
    obs, _ = observations
    checks = dataset.load_page_check()
    for o in obs:
        if _mdna_only(o) and o["status"] == STATUS_MATCHED:
            assert checks[(o["fy"], o["metric"]["id"])]["result"] == "match", o["key"]
            assert o["cross_check"].startswith("Đã đọc lại ngay trên trang 10-K"), o["key"]


def test_page_check_covers_every_mdna_only_value_with_evidence(observations):
    obs, _ = observations
    checks = dataset.load_page_check()
    mdna = [o for o in obs if _mdna_only(o)]
    assert len(mdna) == 35
    for o in mdna:
        c = checks[(o["fy"], o["metric"]["id"])]
        assert c["result"] == "match" and Decimal(c["extracted_value"]) == o["value"], o["key"]
        assert (dataset.ROOT / "handover" / "bang-chung-doi-chieu" / c["screenshot"]).exists(), c["screenshot"]


# ---------------------------------------------------------------- workbook validation
def test_workbook_passes_validation(facts):
    errors, _ = validate(read_observations(WORKBOOK), facts)
    assert errors == []


def _mutated_copy(tmp_path, sheet, key, field, new_value):
    path = tmp_path / "wb.xlsx"
    shutil.copyfile(WORKBOOK, path)
    wb = load_workbook(path)
    ws = wb[sheet]
    for row in ws.iter_rows(min_row=HEADER_ROW + 2):
        if row[FIELD_COL["key"]].value == key:
            row[FIELD_COL[field]].value = new_value
            break
    else:
        raise AssertionError(key)
    wb.save(path)
    return path


@pytest.mark.parametrize("sheet,key,field,value,expected", [
    (SHEET_FIN, "FY2024|revenues", "value", 51000, "differs from the 10-K extraction"),
    (SHEET_CH, "FY2022|nike_direct_revenue", "value", None, "must be marked missing"),
    (SHEET_CH, "FY2021|nike_direct_pct_change_cn", "value", 30, "whole percent"),
    (SHEET_FIN, "FY2020|cost_of_sales", "unit", "tỷ USD", "unit"),
    (SHEET_CH, "FY2025|nike_brand_revenue", "scope", "NIKE, Inc.", "reporting scope"),
    (SHEET_FIN, "FY2023|gross_profit", "status", STATUS_CONFLICT, "conflicting evidence"),
    (SHEET_CH, "FY2024|nike_direct_revenue", "status", STATUS_HUMAN, "verifier's name"),
    (SHEET_FIN, "FY2022|net_income", "page", None, "incomplete provenance"),
    (SHEET_FIN, "FY2022|net_income", "page", 99, "page 99 differs"),
    (SHEET_FIN, "FY2021|inventories", "prior", 6000, "comparative"),
    (SHEET_CH, "FY2026|nike_direct_revenue", "key", "FY2026|revenues", "Duplicate"),
])
def test_validator_rejects_bad_data(tmp_path, facts, sheet, key, field, value, expected):
    errors, _ = validate(read_observations(_mutated_copy(tmp_path, sheet, key, field, value)), facts)
    assert any(expected in e for e in errors), errors


def test_workbook_formulas_recalculate_independently():
    from pipeline.verify_workbook import verify
    n, problems = verify(WORKBOOK)
    assert n > 400 and problems == []


def _check_excel_dashboard(path, site_data):
    from pipeline.build_excel_dashboard import SHEET_CH, SHEET_DATA, SHEET_PERF, SHEET_SRC
    wb = load_workbook(path)
    assert wb.sheetnames == [SHEET_CH, SHEET_PERF, SHEET_SRC, SHEET_DATA]
    ws = wb[SHEET_DATA]
    assert ws.sheet_state == "hidden"
    rows = list(ws[ws.tables["tblNam"].ref])
    header = [c.value for c in rows[0]]
    years = site_data["years"]
    assert [r[header.index("Năm")].value for r in rows[1:]] == years
    for col, kid in (("Tỷ trọng NIKE Direct", "direct_share"), ("Biên EBIT", "ebit_margin"), ("Số ngày tồn kho bình quân", "days_inventory")):
        # XlsxWriter stores 16 significant digits, so compare with a relative tolerance.
        want = [site_data["kpis"][kid]["values"][str(y)] for y in years]
        assert [r[header.index(col)].value for r in rows[1:]] == pytest.approx(want, rel=1e-12), col
    for col, mid in (("NIKE Brand Digital (tỷ USD)", "nike_brand_digital_sales"),
                     ("Tăng trưởng NIKE Direct loại trừ tỷ giá", "nike_direct_pct_change_cn")):
        for r, y in zip(rows[1:], years):
            assert _same(r[header.index(col)].value, site_data["observations"][mid][str(y)]["value"]), (col, y)
    # Table texts: figures and their 10-K pages are in separate rows, so every figure ends under its year header.
    start = next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 1).value).startswith("Văn bản cho bảng"))
    n = 0
    for row in ws.iter_rows(min_row=start + 2, max_col=2 + len(years), values_only=True):
        if row[1] is None:
            break
        cells = [v or "" for v in row[2:]]
        if row[1].endswith(" (trang)"):
            assert all(re.fullmatch(r"(† )?(tr\. \d+)?", v) for v in cells), row
        else:
            assert not any("tr." in v for v in cells), row
        n += 1
    assert n > 40, n
    _check_print_layout(path)


XL = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "x": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
      "p": "http://schemas.openxmlformats.org/package/2006/relationships"}
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def _check_print_layout(path):
    """Both dashboards print at a fixed scale, with manual page breaks that never cut a chart or slicer in two."""
    def part(base, target):     # relationship targets are absolute or relative to the part's folder
        return target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
    with zipfile.ZipFile(path) as z:
        rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        for sheet in ET.fromstring(z.read("xl/workbook.xml")).find("m:sheets", XL)[:2]:
            name = part("xl", rels[sheet.get(REL_ID)])
            ws = ET.fromstring(z.read(name))
            assert ws.find("m:pageSetup", XL).get("scale") == "79", name
            breaks = [int(b.get("id")) for b in ws.iterfind("m:rowBreaks/m:brk", XL)]
            assert breaks, name
            srels = ET.fromstring(z.read(name.replace("worksheets/", "worksheets/_rels/") + ".rels"))
            drawing = next(part("xl/worksheets", r.get("Target")) for r in srels.iterfind("p:Relationship", XL)
                           if r.get("Type").endswith("/drawing"))
            anchors = list(ET.fromstring(z.read(drawing)).iter(f"{{{XL['x']}}}twoCellAnchor"))
            assert len(anchors) >= 4, drawing
            for a in anchors:
                top = int(a.find("x:from/x:row", XL).text)
                bottom = int(a.find("x:to/x:row", XL).text) - (a.find("x:to/x:rowOff", XL).text == "0")
                assert not [b for b in breaks if top < b <= bottom], (name, top, bottom, breaks)


def test_excel_dashboard_build_matches_site_data(tmp_path, site_data):
    from pipeline.build_excel_dashboard import build
    path = tmp_path / "dashboard.xlsx"
    build(path, add_slicers=False)
    _check_excel_dashboard(path, site_data)


@pytest.mark.filterwarnings("ignore:Unknown extension")   # openpyxl does not read slicer definitions
def test_delivered_excel_dashboard_is_current(site_data):
    """The handover copy must have been rebuilt from the current data and carry its two slicer caches."""
    from pipeline.build_excel_dashboard import OUT
    _check_excel_dashboard(OUT, site_data)
    with zipfile.ZipFile(OUT) as z:
        assert len([n for n in z.namelist() if n.startswith("xl/slicerCaches/")]) == 2
    # Excel recalculated this copy when the slicers were added, so the tables' displayed values are stored:
    # figures carry no page reference, and page cells hold only "tr. N" (a blank page must not show as 0).
    wb = load_workbook(OUT, data_only=True)
    for name in wb.sheetnames[:2]:
        ws = wb[name]
        head = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 2).value == "Chỉ tiêu")
        page_rows = 0
        for r in range(head + 1, ws.max_row + 1):
            cells = [ws.cell(r, c).value for c in range(6, 13)]
            assert all(v is None or isinstance(v, str) for v in cells), (name, r, cells)
            if ws.cell(r, 2).value is None and any(cells):
                assert all(re.fullmatch(r"(† )?(tr\. \d+)?", v or "") for v in cells), (name, r, cells)
                page_rows += 1
            else:
                assert not any("tr." in (v or "") for v in cells), (name, r, cells)
        assert page_rows > 5, (name, page_rows)


def _tmdl_rows(path):
    """Rows of the #table literal embedded in a TMDL partition, as dicts keyed by column name."""
    text = path.read_text(encoding="utf-8")
    names = re.findall(r'#"((?:[^"]|"")*)" = [\w.]+', text)
    token = re.compile(r'"((?:[^"]|"")*)"|(null)|(-?[\d.]+(?:[eE][-+]?\d+)?)')
    rows = []
    for line in re.findall(r"^\s*\{(.*)\},?\s*$", text, re.M):
        vals = [m.group(1).replace('""', '"') if m.group(1) is not None else None if m.group(2) else float(m.group(3))
                for m in token.finditer(line)]
        rows.append(dict(zip(names, vals)))
    return rows


def _same(got, want):
    return got is None if want is None else got == pytest.approx(want, rel=1e-12)


def test_powerbi_model_matches_site_data(site_data):
    from pipeline.build_powerbi import NAME
    tables = dataset.ROOT / "powerbi" / f"{NAME}.SemanticModel" / "definition" / "tables"
    rows = {int(r["FY"]): r for r in _tmdl_rows(tables / "Nam.tmdl")}
    assert sorted(rows) == site_data["years"]
    for y in site_data["years"]:
        for mid, by_year in site_data["observations"].items():
            assert _same(rows[y][mid], by_year[str(y)]["value"]), (mid, y)
        for kid, k in site_data["kpis"].items():
            assert _same(rows[y][kid], k["values"][str(y)]), (kid, y)
    # Table texts: each year is two matrix columns, the figure under "FY…" and then its 10-K page, so every
    # figure ends under its year header. Page references link to that year's 10-K.
    for name in ("BangKenh", "BangKetQua"):
        cells = _tmdl_rows(tables / f"{name}.tmdl")
        figures, pages = [c for c in cells if c["Loại"] == "Số"], [c for c in cells if c["Loại"] == "Trang"]
        assert len(figures) == len(pages) > 0 and len(figures) + len(pages) == len(cells), name
        for c in figures:
            assert c["Cột"] == f"FY{int(c['FY'])}" and c["Thứ tự cột"] % 2 == 0, c
            assert "tr." not in c["Hiển thị"] and c["Liên kết"] is None, c
        for c in pages:
            assert re.fullmatch(r"(†|† tr\. \d+|tr\. \d+)?", c["Hiển thị"]), c
            url = site_data["filings"][str(int(c["FY"]))]["url"]
            assert c["Liên kết"] == (url if "tr." in c["Hiển thị"] else None), c


def test_delivered_pbix_opens_on_all_years():
    """The .pbix must have three pages and no saved year filter (it once shipped filtered to FY2024–2026)."""
    pbix = dataset.ROOT / "handover" / "NIKE_Dashboard_FY2020-2026.pbix"
    with zipfile.ZipFile(pbix) as z:
        pages = [n for n in z.namelist() if re.fullmatch(r"Report/definition/pages/[^/]+/page\.json", n)]
        slicers = [n for n in z.namelist() if re.search(r"/visuals/loc_nam/visual\.json$", n)]
        assert len(pages) == 3 and len(slicers) == 2
        for n in slicers:
            assert '"filter"' not in z.read(n).decode("utf-8"), n


def test_site_json_matches_workbook(observations, site_data):
    obs, results = observations
    for o in obs:
        got = site_data["observations"][o["metric"]["id"]][str(o["fy"])]["value"]
        assert got == (None if o["value"] is None else float(o["value"])), o["key"]
    for kid, k in site_data["kpis"].items():
        for y in dataset.YEARS:
            want = results[y][kid]
            assert k["values"][str(y)] == (None if want is None else float(want)), (kid, y)
