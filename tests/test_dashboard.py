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
    # † marks and their key appear only while some figures are still pending review.
    pending = site_data["meta"]["status_counts"]["Chờ người kiểm tra"]
    assert (page.locator("#table-channels .mark-slot", has_text="†").count() > 0) == (pending > 0)
    assert page.locator("#view-channels .pending-key").is_hidden() == (pending == 0)
    assert page.locator("#filing-list a").count() == 7


def test_xlsx_and_png_exports(page, site_data, tmp_path):
    from openpyxl import load_workbook
    page.open()
    with page.expect_download() as d:
        page.click("#xlsx-channels")
    assert d.value.suggested_filename == "NIKE_kenh-phan-phoi_FY2020-FY2026.xlsx"
    saved = tmp_path / d.value.suggested_filename
    d.value.save_as(saved)
    ws = load_workbook(saved).active
    rows = [[c.value for c in r] for r in ws.iter_rows()]
    assert rows[0][:4] == ["Chỉ tiêu", "Đơn vị", "Loại", "FY2020"]
    direct = next(r for r in rows if r[0] == "NIKE Direct")
    assert direct[3:10] == [site_data["observations"]["nike_direct_revenue"][str(y)]["value"] for y in site_data["years"]]
    stated = next(r for r in rows if r[0] == "Tỷ trọng NIKE Direct do NIKE nêu (xấp xỉ)")
    assert stated[9] == "Không công bố"
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


HEADER_VS_FIGURES = """(view) => {
  const right = n => { const r = document.createRange(); r.selectNodeContents(n); return r.getBoundingClientRect().right; };
  const heads = Array.from(document.querySelectorAll(`#table-${view} thead th`)).slice(2);
  return Array.from(document.querySelectorAll(`#table-${view} tbody tr`)).filter(r => !r.classList.contains('group'))
    .flatMap(r => Array.from(r.querySelectorAll('td.num')).map((td, i) =>
      [r.querySelector('th').textContent + ' / ' + heads[i].textContent, Math.abs(right(heads[i]) - right(td.firstChild))]));
}"""


def assert_headers_align(pg, view):
    """Every cell (reported, calculated or 'Không công bố') ends exactly under its year header."""
    offsets = pg.evaluate(HEADER_VS_FIGURES, view)
    off = [(cell, round(d, 1)) for cell, d in offsets if d > 1]
    assert offsets and not off, (view, off)


def test_tables_fit_and_headers_align_on_desktop(page):
    for q in ("", "?view=performance"):
        page.open(q)
        view = "performance" if q else "channels"
        width, client = page.evaluate(f"() => {{ const s = document.querySelector('#view-{view} .table-scroll'); "
                                      f"return [s.scrollWidth, s.clientWidth]; }}")
        assert width <= client, (view, width, client)
        assert_headers_align(page, view)


def test_print_draws_both_views(page):
    page.open()
    page.evaluate("() => window.dispatchEvent(new Event('beforeprint'))")
    page.emulate_media(media="print")
    for chart in ("revenue", "grossMargin", "sga", "ebit", "inventory", "days"):
        size = page.evaluate(f"() => {{ const c = document.querySelector('.chart[data-chart=\"{chart}\"] canvas'); "
                             f"return c ? c.width * c.height : 0; }}")
        assert size > 0, chart
    assert page.locator("#table-performance tbody tr").count() > 20
    for view in ("channels", "performance"):
        assert_headers_align(page, view)
    page.emulate_media(media="screen")
    page.evaluate("() => window.dispatchEvent(new Event('afterprint'))")
    assert page.is_hidden("#view-performance")


PINNED = """(view) => {
  const box = document.querySelector(`#view-${view} .table-scroll`), b = box.getBoundingClientRect();
  const heads = Array.from(document.querySelectorAll(`#table-${view} thead th`)), edge = heads[0].getBoundingClientRect().right;
  const label = document.querySelector(`#table-${view} tbody tr:not(.group) th`).getBoundingClientRect();
  const clipped = Array.from(box.querySelectorAll('.group-label')).map(g => [g.textContent, g.getBoundingClientRect()])
    .filter(([, r]) => r.left < b.left - 0.5 || r.right > b.right + 0.5).map(([t]) => t);
  // A column cut by the pinned column's edge would show the tail of an earlier year beside the labels.
  const partial = heads.slice(1).map(h => [h.textContent, h.getBoundingClientRect()])
    .filter(([, r]) => r.left < edge - 0.5 && r.right > edge + 0.5).map(([t]) => t);
  return {scrolled: box.scrollLeft > 0 && box.classList.contains('scrolled'),
          header: Math.round(heads[0].getBoundingClientRect().left - label.left), clipped, partial};
}"""


def test_narrow_layouts(browser, server):
    for width in (390, 1024):
        ctx = browser.new_context(viewport={"width": width, "height": 844})
        pg = ctx.new_page()
        for q in ("", "?view=performance"):
            pg.goto(server + q)
            pg.wait_for_selector('html[data-ready="true"]')
            assert pg.evaluate("document.documentElement.scrollWidth <= document.documentElement.clientWidth"), width
            view = "performance" if q else "channels"
            assert_headers_align(pg, view)
            # The table opens on the latest years; the label column, its header and the group names stay in view,
            # and the first year in view starts at the label column's edge.
            pinned = pg.evaluate(PINNED, view)
            assert pinned == {"scrolled": True, "header": 0, "clipped": [], "partial": []}, (width, view, pinned)
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
