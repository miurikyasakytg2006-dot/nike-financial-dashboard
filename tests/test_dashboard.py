"""Browser tests for the static dashboard (Playwright + Chrome)."""
import re

ECHARTS_SERIES = """id => {
  const c = echarts.getInstanceByDom(document.querySelector('.chart[data-chart="' + id + '"]'));
  const o = c.getOption();
  return {x: o.xAxis[0].data || null, y: o.yAxis[0].data || null,
          series: o.series.map(s => ({name: s.name, data: s.data.map(d => d && typeof d === 'object' ? d.value : d)}))};
}"""


def vi_number(text):
    """Parse a vi-VN formatted figure from a table cell: '12.382', '(50)', '−4,4%', '+8%'."""
    t = text.split("tr.")[0].replace("†", "").replace("%", "").strip()
    neg = t.startswith("(") or t.startswith("−") or t.startswith("-")
    t = t.strip("()+−-").replace(".", "").replace(",", ".")
    v = float(t)
    return -v if neg else v


def table(page, view):
    return page.eval_on_selector(f"#table-{view}", """t => Array.from(t.querySelectorAll('tbody tr')).filter(r => !r.classList.contains('group'))
        .map(r => Array.from(r.children).map(c => c.textContent.trim()))""")


def test_initial_render_has_no_errors(page):
    page.open()
    assert page.console_errors == []
    for chart in ("share", "channelRevenue", "growth", "digital"):
        assert page.locator(f'.chart[data-chart="{chart}"] canvas').count() == 1
        assert "Số liệu chi tiết" in page.get_attribute(f'.chart[data-chart="{chart}"]', "aria-label")
    assert page.is_hidden("#view-performance")
    assert page.locator("#table-channels thead th").count() == 2 + 7


def test_table_values_match_data(page, site_data):
    page.open()
    years = site_data["years"]
    rows = {r[0]: r for r in table(page, "channels")}
    direct = rows["NIKE Direct"]
    for i, y in enumerate(years):
        assert vi_number(direct[2 + i]) == site_data["observations"]["nike_direct_revenue"][str(y)]["value"]
    share = rows["Tỷ trọng NIKE Direct"]
    for i, y in enumerate(years):
        assert round(vi_number(share[2 + i]), 1) == round(site_data["kpis"]["direct_share"]["values"][str(y)] * 100, 1)
    page.open("?view=performance")
    rows = {r[0]: r for r in table(page, "performance")}
    for label, metric in (("Doanh thu", "revenues"), ("Chi phí (thu nhập) lãi thuần", "interest_expense_net"),
                          ("Hàng tồn kho tại 31/5", "inventories"), ("EBIT do NIKE công bố", "ebit_reported")):
        for i, y in enumerate(years):
            assert vi_number(rows[label][2 + i]) == site_data["observations"][metric][str(y)]["value"], (label, y)


def test_chart_data_match_data(page, site_data):
    page.open()
    years = site_data["years"]
    rev = page.evaluate(ECHARTS_SERIES, "channelRevenue")
    assert [s["name"] for s in rev["series"]] == ["NIKE Direct", "Bán buôn"]
    assert rev["series"][0]["data"] == [site_data["observations"]["nike_direct_revenue"][str(y)]["value"] / 1000 for y in years]
    share = page.evaluate(ECHARTS_SERIES, "share")
    assert share["series"][0]["data"] == [site_data["kpis"]["direct_share"]["values"][str(y)] for y in years]


def test_year_filter_updates_everything_and_url(page, site_data):
    page.open()
    page.select_option("#year-from", "2022")
    page.wait_for_function("location.search.includes('from=2022')")
    assert [h for h in page.locator("#table-channels thead th").all_inner_texts()][2:] == [f"FY{y}" for y in range(2022, 2027)]
    for chart in ("share", "channelRevenue", "growth", "digital"):
        data = page.evaluate(ECHARTS_SERIES, chart)
        assert len(data["series"][0]["data"]) == 5, chart
    summary = page.inner_text("#summary-channels")
    assert "FY2022" in summary and "FY2020" not in summary
    assert page.locator('#year-to option[value="2021"]').is_disabled()
    page.wait_for_function("document.getElementById('live').textContent.includes('FY2022')")


def test_deep_link_and_reset(page):
    page.open("?view=performance&from=2021&to=2024&basis=cn")
    assert page.is_visible("#view-performance")
    assert page.get_attribute("#tab-performance", "aria-selected") == "true"
    assert page.input_value("#year-from") == "2021" and page.input_value("#year-to") == "2024"
    page.click("#reset")
    assert page.input_value("#year-from") == "2020" and page.input_value("#year-to") == "2026"
    page.wait_for_function("!location.search.includes('from')")
    assert "view=performance" in page.url


def test_tabs_keyboard(page):
    page.open()
    page.focus("#tab-channels")
    page.keyboard.press("ArrowRight")
    assert page.is_visible("#view-performance")
    assert page.evaluate("document.activeElement.id") == "tab-performance"
    for chart in ("revenue", "grossMargin", "sga", "ebit", "inventory", "days"):
        assert page.locator(f'.chart[data-chart="{chart}"] canvas').count() == 1
    page.keyboard.press("Home")
    assert page.is_visible("#view-channels")
    outline = page.evaluate("getComputedStyle(document.activeElement).outlineStyle")
    assert outline != "none"


def test_growth_basis_toggle(page, site_data):
    page.open()
    years = site_data["years"]
    reported = page.evaluate(ECHARTS_SERIES, "growth")
    assert reported["series"][0]["data"] == [site_data["kpis"]["direct_growth"]["values"][str(y)] for y in years]
    page.check('#growth-basis input[value="cn"]', force=True)
    page.wait_for_function("location.search.includes('basis=cn')")
    cn = page.evaluate(ECHARTS_SERIES, "growth")
    assert cn["series"][0]["data"] == [site_data["observations"]["nike_direct_pct_change_cn"][str(y)]["value"] for y in years]
    assert "non-GAAP" in page.inner_text("#growth-sub")


def test_series_toggle_and_last_series_guard(page):
    page.open()
    buttons = page.locator('[data-legend="channelRevenue"] button')
    buttons.nth(1).click()
    assert page.locator('[data-legend="channelRevenue"] button').nth(1).get_attribute("aria-pressed") == "false"
    assert [s["name"] for s in page.evaluate(ECHARTS_SERIES, "channelRevenue")["series"]] == ["NIKE Direct"]
    page.locator('[data-legend="channelRevenue"] button').nth(0).click()
    assert [s["name"] for s in page.evaluate(ECHARTS_SERIES, "channelRevenue")["series"]] == ["NIKE Direct"]
    page.wait_for_function("document.getElementById('live').textContent.includes('ít nhất một')")


def test_tooltip_shows_units_and_source(page):
    page.open()
    page.locator('.chart[data-chart="channelRevenue"]').scroll_into_view_if_needed()
    box = page.locator('.chart[data-chart="channelRevenue"]').bounding_box()
    page.mouse.move(box["x"] + box["width"] * 0.55, box["y"] + box["height"] * 0.6)
    page.wait_for_timeout(400)
    tip = page.evaluate("""() => Array.from(document.querySelectorAll('.chart[data-chart="channelRevenue"] div'))
        .map(d => d.innerText).filter(t => t.includes('tỷ USD')).join(' ')""")
    assert "tỷ USD" in tip and "Nguồn: 10-K FY" in tip and "FY20" in tip


def test_source_links_and_missing_state(page, site_data):
    page.open()
    hrefs = page.eval_on_selector_all("#table-channels a.page-ref", "as => as.map(a => a.href)")
    allowed = {f["url"] for f in site_data["filings"].values()}
    assert hrefs and set(hrefs) <= allowed
    assert all(h.startswith("https://www.sec.gov/Archives/edgar/data/320187/") for h in hrefs)
    rows = {r[0]: r for r in table(page, "channels")}
    assert rows["Tỷ trọng NIKE Direct do NIKE nêu (xấp xỉ)"][-1].startswith("Không công bố")
    assert page.locator("#table-channels .mark-slot", has_text="†").count() > 0
    assert page.locator("#filing-list a").count() == 7


def test_csv_and_png_exports(page):
    page.open()
    with page.expect_download() as d:
        page.click("#csv-channels")
    text = open(d.value.path(), encoding="utf-8-sig").read()
    assert text.startswith("Chỉ tiêu,Mã,Loại") and "nike_direct_revenue" in text and "21519" in text
    with page.expect_download() as d:
        page.click('.export-png[data-chart="share"]')
    data = open(d.value.path(), "rb").read()
    assert data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) > 20000
    assert re.match(r"NIKE_share_FY2020-FY2026\.png", d.value.suggested_filename)


def test_workbook_download_available(page, server):
    page.open()
    href = page.get_attribute("#download-workbook", "href")
    resp = page.request.get(server + href)
    assert resp.ok and resp.body()[:2] == b"PK"


def test_accessible_names_and_headings(page):
    page.open()
    unnamed = page.evaluate("""() => Array.from(document.querySelectorAll('button, select, a, input'))
        .filter(e => e.offsetParent !== null || e.type === 'radio')
        .filter(e => !(e.getAttribute('aria-label') || e.textContent.trim() || (e.labels && e.labels.length)))
        .map(e => e.outerHTML.slice(0, 80))""")
    assert unnamed == []
    levels = page.eval_on_selector_all("h1, h2, h3", "hs => hs.filter(h => h.offsetParent !== null).map(h => +h.tagName[1])")
    assert levels[0] == 1 and all(b - a <= 1 for a, b in zip(levels, levels[1:]))
    assert page.locator("#table-channels thead th[scope=col]").count() == 9


def test_mobile_layout_has_no_page_overflow(browser, server):
    ctx = browser.new_context(viewport={"width": 390, "height": 844})
    pg = ctx.new_page()
    for q in ("", "?view=performance"):
        pg.goto(server + q)
        pg.wait_for_selector('html[data-ready="true"]')
        assert pg.evaluate("document.documentElement.scrollWidth <= document.documentElement.clientWidth")
    ctx.close()


def test_reduced_motion_disables_chart_animation(browser, server):
    ctx = browser.new_context(reduced_motion="reduce")
    pg = ctx.new_page()
    pg.goto(server)
    pg.wait_for_selector('html[data-ready="true"]')
    animated = pg.evaluate("""() => Array.from(document.querySelectorAll('#view-channels .chart'))
        .map(el => echarts.getInstanceByDom(el).getOption().animation)""")
    assert animated and not any(animated)
    ctx.close()
