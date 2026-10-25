"""Key findings and section summaries, written once from the computed figures.

The web dashboard, the Excel dashboard and the Power BI report all show these exact strings, so the three versions
cannot say different things. Each finding states figures only: what changed, by how much, and when. None of them
says why; NIKE's own explanations are quoted separately, from the 10-K (pipeline.quotes).
"""
from decimal import Decimal

from .dataset import vi_num

B = Decimal(1000)


def _pct(v, d=1):
    return vi_num(v * 100, d) + "%"


def _spct(v, d=1):
    return vi_num(v * 100, d, signed=True) + "%"


def _pts(v, d=1):
    return vi_num(v * 100, d, signed=True) + " điểm %"


def _bn(v, d=1):
    return vi_num(v / B, d)


def findings(obs, results):
    """Six findings: {id, value, label, text, source, target} — `target` is the dashboard section that shows it."""
    val = {(o["period"], o["metric"]["id"]): o["value"] for o in obs}
    k = lambda kid, p: results[p][kid]
    v = lambda mid, p: val[(p, mid)]
    share19, share21 = k("direct_share", "FY2019"), k("direct_share", "FY2021")
    d_direct = v("nike_direct_revenue", "FY2021") - v("nike_direct_revenue", "FY2019")
    d_whs = v("nike_brand_wholesale_revenue", "FY2021") - v("nike_brand_wholesale_revenue", "FY2019")
    d_brand = v("nike_brand_revenue", "FY2021") - v("nike_brand_revenue", "FY2019")
    dig19, dig21 = v("nike_brand_digital_sales", "FY2019"), v("nike_brand_digital_sales", "FY2021")
    ex19, ex21 = k("direct_ex_digital", "FY2019"), k("direct_ex_digital", "FY2021")
    q = "FY2020Q4"
    hi = max((p for p in results if "Q" in p), key=lambda p: results[p]["direct_share"])
    regions21 = {r: k(f"{r}_direct_share", "FY2021") for r in ("na", "emea", "gc", "apla")}
    rise = {r: regions21[r] - k(f"{r}_direct_share", "FY2019") for r in regions21}
    top, most = max(regions21, key=regions21.get), max(rise, key=rise.get)
    up19, up21 = k(f"{most}_direct_share", "FY2019"), regions21[most]
    names = {"na": "Bắc Mỹ", "emea": "EMEA", "gc": "Đại Trung Hoa", "apla": "APLA"}
    more = "nhiều hơn" if dig21 - dig19 > d_direct / B else "ít hơn"
    return [
        dict(id="share", target="kenh", value=_pct(share21), label="Tỷ trọng NIKE Direct, FY2021",
             text=f"NIKE Direct chiếm {_pct(share21)} doanh thu NIKE Brand ở FY2021, tăng từ {_pct(share19)} ở FY2019 "
                  f"({_pts(share21 - share19)}).",
             source="Thuyết minh “Revenues”, 10-K FY2019 và FY2021. Tỷ trọng tự tính."),
        dict(id="growth", target="kenh", value=f"+{_bn(d_direct)} tỷ USD", label="NIKE Direct tăng thêm, FY2019 → FY2021",
             text=f"Từ FY2019 đến FY2021, doanh thu NIKE Brand tăng {_bn(d_brand)} tỷ USD, trong đó NIKE Direct tăng "
                  f"{_bn(d_direct)} tỷ USD và bán buôn tăng {_bn(d_whs)} tỷ USD.",
             source="Thuyết minh “Revenues”, 10-K FY2019 và FY2021."),
        dict(id="digital", target="digital", value=f"{vi_num(dig21, 1)} tỷ USD", label="Doanh số Digital, FY2021",
             text=f"Doanh số Digital tăng từ {vi_num(dig19, 1)} lên {vi_num(dig21, 1)} tỷ USD (+{vi_num(dig21 - dig19, 1)} tỷ USD), "
                  f"{more} mức tăng của cả NIKE Direct. Phần NIKE Direct ngoài Digital, chủ yếu là cửa hàng, "
                  f"{'giảm' if ex21 < ex19 else 'tăng'} từ khoảng {_bn(ex19)} {'xuống' if ex21 < ex19 else 'lên'} {_bn(ex21)} "
                  f"tỷ USD (ước tính).",
             source="MD&A, 10-K FY2019 và FY2021 (số Digital đã làm tròn)."),
        dict(id="covid", target="quy", value=_spct(k("wholesale_growth", q)), label="Bán buôn, quý 4 FY2020 so với cùng kỳ",
             text=f"Quý 4 FY2020 (tháng 3–5/2020), bán buôn giảm {_pct(-k('wholesale_growth', q))} so với cùng kỳ, NIKE Direct "
                  f"giảm {_pct(-k('direct_growth', q))}. Tỷ trọng NIKE Direct lên {_pct(k('direct_share', q))}"
                  + (", cao nhất trong 12 quý." if hi == q else "."),
             source="Tự tính: 10-K FY2020 trừ 10-Q quý 3 FY2020 (9 tháng)."),
        dict(id="profit", target="loi-nhuan", value=_pct(k("ebit_margin", "FY2021")), label="Biên EBIT, FY2021",
             text=f"Biên EBIT FY2021 là {_pct(k('ebit_margin', 'FY2021'))}, so với {_pct(k('ebit_margin', 'FY2019'))} ở FY2019 "
                  f"và {_pct(k('ebit_margin', 'FY2020'))} ở FY2020. Biên gộp gần như giữ nguyên ({_pct(k('gross_margin', 'FY2019'))} "
                  f"→ {_pct(k('gross_margin', 'FY2021'))}); chi phí tạo nhu cầu giảm từ "
                  f"{_pct(k('demand_creation_ratio', 'FY2019'))} xuống {_pct(k('demand_creation_ratio', 'FY2021'))} doanh thu.",
             source="Báo cáo kết quả kinh doanh và bảng EBIT trong MD&A, 10-K FY2019–FY2021. Biên tự tính."),
        dict(id="regions", target="khu-vuc", value=_pts(up21 - up19), label=f"{names[most]}: tỷ trọng NIKE Direct",
             text=f"{names[most]} tăng tỷ trọng NIKE Direct nhiều nhất trong bốn khu vực: {_pct(up19)} (FY2019) → "
                  f"{_pct(up21)} (FY2021). {names[top]} có tỷ trọng cao nhất ở FY2021, {_pct(regions21[top])}.",
             source="Thuyết minh “Revenues”, 10-K FY2019 và FY2021. Tỷ trọng tự tính."),
    ]


def section_notes(obs, results):
    """One plain line under some section headings."""
    k = lambda kid, p: results[p][kid]
    val = {(o["period"], o["metric"]["id"]): o["value"] for o in obs}
    inv19, inv20 = val[("FY2019", "inventories")], val[("FY2020", "inventories")]
    fy21 = [f"FY2021Q{i}" for i in (1, 2, 3, 4)]
    direct_up = all(k("direct_growth", p) > 0 for p in fy21)
    whs_down = [p for p in fy21 if k("wholesale_growth", p) < 0]
    after = ("Cả bốn quý FY2021, NIKE Direct đều tăng so với cùng kỳ" if direct_up else "")
    if after and whs_down == fy21[:3]:
        after += ", còn bán buôn giảm ở ba quý đầu"
    return {
        "quy": (f"Doanh thu NIKE, Inc. quý 4 FY2020 giảm {_pct(-k('revenue_growth', 'FY2020Q4'))} so với cùng kỳ."
                + (f" {after}." if after else "")),
        "ton-kho": (f"Tồn kho ngày 31/5/2020 là {_bn(inv20)} tỷ USD, tăng {_pct(inv20 / inv19 - 1)} so với một năm trước; "
                    f"số ngày tồn kho quý 4 FY2020 là {vi_num(k('days_inventory', 'FY2020Q4'), 0)} ngày, cùng kỳ năm trước "
                    f"{vi_num(k('days_inventory', 'FY2019Q4'), 0)} ngày."),
    }
