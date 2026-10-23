/* NIKE FY2020–FY2026 dashboard: plain JavaScript + Apache ECharts.
   All numbers come from data/nike-financials.json, which the Python pipeline builds
   from the validated Excel workbook. Nothing here computes new financial values except
   unit conversion (USD million to USD billion) and formatting. */
(function () {
  'use strict';

  var C = {
    direct: '#2a78d6', wholesale: '#eb6834', gbd: '#a3a9b2', directLight: '#86b6ef', digital: '#1c5cab',
    total: '#475467', overhead: '#6250d6', demand: '#1baf7a',
    ink: '#15191e', ink2: '#3b4350', muted: '#5b6472', grid: '#e6e9ee', axis: '#c3c9d1'
  };
  var FONT = '"Source Sans 3", system-ui, -apple-system, "Segoe UI", Arial, sans-serif';
  var REDUCED = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  var MINUS = '−';
  var DEFAULTS = { view: 'channels', basis: 'reported' };
  var STATUS_PENDING = 'Chờ người kiểm tra';

  var D = null;
  var state = { view: 'channels', from: null, to: null, basis: 'reported', hidden: {} };
  var charts = {};

  // ------------------------------------------------------------------ formatting
  var nfCache = {};
  function nf(d) {
    if (!nfCache[d]) nfCache[d] = new Intl.NumberFormat('vi-VN', { minimumFractionDigits: d, maximumFractionDigits: d });
    return nfCache[d];
  }
  function num(v, d) { if (v == null) return '—'; var s = nf(d || 0).format(Math.abs(v)); return v < 0 ? MINUS + s : s; }
  function amount(v) { if (v == null) return '—'; var s = nf(0).format(Math.abs(v)); return v < 0 ? '(' + s + ')' : s; }
  function pct(v, d) { return v == null ? '—' : num(v * 100, d == null ? 1 : d) + '%'; }
  function signedPct(v, d) {
    if (v == null) return '—';
    var s = nf(d == null ? 1 : d).format(Math.abs(v * 100));
    return (v > 0 ? '+' : v < 0 ? MINUS : '') + s + '%';
  }
  var dateFmt = new Intl.DateTimeFormat('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric', timeZone: 'UTC' });
  function isoDate(iso) { return dateFmt.format(new Date(iso + 'T00:00:00Z')); }
  function bn(millions, d) { return millions == null ? '—' : num(millions / 1000, d == null ? 1 : d); }
  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; });
  }

  // ------------------------------------------------------------------ data access
  function years() { return D.years.filter(function (y) { return y >= state.from && y <= state.to; }); }
  function ob(id, y) { return (D.observations[id] || {})[y] || null; }
  function val(id, y) { var o = ob(id, y); return o ? o.value : null; }
  function kpi(id, y) { return D.kpis[id].values[y]; }
  function period(y) { return '01/06/' + (y - 1) + ' – 31/05/' + y; }
  function page(id, y) { var o = ob(id, y); return o && o.page ? o.page : null; }
  function srcNote(id, y) { var p = page(id, y); return 'Nguồn: 10-K FY' + y + (p ? ', tr. ' + p : '') + '.'; }

  // ------------------------------------------------------------------ URL state
  function readURL() {
    var p = new URLSearchParams(location.search);
    state.view = p.get('view') === 'performance' ? 'performance' : 'channels';
    var f = Number(p.get('from')), t = Number(p.get('to'));
    state.from = D.years.indexOf(f) >= 0 ? f : D.years[0];
    state.to = D.years.indexOf(t) >= 0 ? t : D.years[D.years.length - 1];
    if (state.from > state.to) { var x = state.from; state.from = state.to; state.to = x; }
    state.basis = p.get('basis') === 'cn' ? 'cn' : 'reported';
  }
  function writeURL() {
    var p = new URLSearchParams();
    if (state.view !== DEFAULTS.view) p.set('view', state.view);
    if (state.from !== D.years[0]) p.set('from', state.from);
    if (state.to !== D.years[D.years.length - 1]) p.set('to', state.to);
    if (state.basis !== DEFAULTS.basis) p.set('basis', state.basis);
    var q = p.toString();
    history.replaceState(null, '', (q ? '?' + q : location.pathname) + location.hash);
  }

  // ------------------------------------------------------------------ chart scaffolding
  function base() {
    return {
      animation: !REDUCED, animationDuration: 250, animationDurationUpdate: 250,
      textStyle: { fontFamily: FONT, color: C.ink2, fontSize: 13 },
      grid: { left: 4, right: 12, top: 22, bottom: 4, containLabel: true },
      tooltip: {
        trigger: 'axis', confine: true, backgroundColor: '#fff', borderColor: C.axis, borderWidth: 1, padding: [8, 10],
        textStyle: { color: C.ink, fontFamily: FONT, fontSize: 13 },
        axisPointer: { type: 'shadow', shadowStyle: { color: 'rgba(21,25,30,0.05)' } },
        extraCssText: 'box-shadow:0 2px 8px rgba(21,25,30,.12);border-radius:4px;max-width:300px;white-space:normal;'
      }
    };
  }
  function catAxis(ys, chartId) {
    // Narrow charts use the compact "FY20" form so every year keeps its label.
    var el = chartId && document.querySelector('.chart[data-chart="' + chartId + '"]');
    var compact = el && el.clientWidth / Math.max(ys.length, 1) < 62;
    return {
      type: 'category', data: ys.map(function (y) { return compact ? 'FY' + String(y).slice(2) : 'FY' + y; }), axisTick: { show: false },
      axisLine: { onZero: false, lineStyle: { color: C.axis } },
      axisLabel: { color: C.muted, fontSize: 12.5, fontFamily: FONT, hideOverlap: true }
    };
  }
  function valAxis(formatter, extra) {
    var a = {
      type: 'value', axisLabel: { color: C.muted, fontSize: 12.5, fontFamily: FONT, formatter: formatter },
      splitLine: { lineStyle: { color: C.grid, width: 1, type: 'solid' } }, axisLine: { show: false }, axisTick: { show: false }
    };
    return Object.assign(a, extra || {});
  }
  function tip(title, rows, foot) {
    var h = '<div style="font-weight:600;margin-bottom:4px">' + esc(title) + '</div>';
    rows.forEach(function (r) {
      h += '<div style="display:flex;align-items:baseline;gap:8px;line-height:1.55">' +
        '<span style="display:inline-block;width:12px;height:3px;border-radius:2px;background:' + r.color + ';flex:none;transform:translateY(-3px)"></span>' +
        '<span style="font-weight:600;font-variant-numeric:tabular-nums;white-space:nowrap">' + esc(r.value) + '</span>' +
        '<span style="color:' + C.muted + '">' + esc(r.label) + '</span></div>';
    });
    if (foot) h += '<div style="margin-top:6px;color:' + C.muted + ';font-size:12px;line-height:1.4">' + esc(foot) + '</div>';
    return h;
  }
  function endLabel(ys, fmt, color) {
    // Selective direct labels: first and last visible year only.
    return {
      show: true, position: 'top', distance: 4, color: color || C.ink2, fontSize: 12.5, fontFamily: FONT,
      formatter: function (p) { return (p.dataIndex === 0 || p.dataIndex === ys.length - 1) && p.value != null ? fmt(p.value) : ''; }
    };
  }
  function lineSeries(name, data, color, ys, fmt) {
    // End labels sit on the side away from the line's neighbouring segment, so they never cross it.
    var last = data.length - 1;
    var items = data.map(function (v, i) {
      if (v == null || (i !== 0 && i !== last) || data.length === 1) return v;
      var neighbour = data[i === 0 ? 1 : last - 1];
      return { value: v, label: { position: neighbour != null && neighbour > v ? 'bottom' : 'top' } };
    });
    return {
      name: name, type: 'line', data: items, symbol: 'circle', symbolSize: 8, showSymbol: true, connectNulls: false,
      lineStyle: { width: 2, color: color, cap: 'round', join: 'round' },
      itemStyle: { color: color, borderColor: '#fff', borderWidth: 2 },
      label: endLabel(ys, fmt), emphasis: { scale: 1.25 }
    };
  }
  function columnSeries(name, data, color, extra) {
    return Object.assign({
      name: name, type: 'bar', data: data, barMaxWidth: 24,
      itemStyle: { color: color, borderRadius: [4, 4, 0, 0] }, emphasis: { itemStyle: { opacity: 0.85 } }
    }, extra || {});
  }
  function hidden(id) { return state.hidden[id] || (state.hidden[id] = {}); }
  function visible(id, name) { return !hidden(id)[name]; }

  function setChart(id, option, description) {
    var el = document.querySelector('.chart[data-chart="' + id + '"]');
    if (!charts[id]) {
      charts[id] = echarts.init(el, null, { renderer: 'canvas' });
      if (window.ResizeObserver) new ResizeObserver(function () { charts[id].resize(); }).observe(el);
    }
    charts[id].setOption(option, { notMerge: true });
    el.setAttribute('aria-label', description + ' Số liệu chi tiết có trong bảng bên dưới.');
  }

  function legend(id, items, toggle) {
    var el = document.querySelector('[data-legend="' + id + '"]');
    if (!el) return;
    el.replaceChildren();
    if (toggle) el.setAttribute('aria-label', 'Bật hoặc tắt chuỗi số liệu');
    items.forEach(function (it) {
      var node = document.createElement(toggle ? 'button' : 'span');
      node.className = 'legend-item';
      if (toggle) {
        node.type = 'button';
        node.setAttribute('aria-pressed', String(visible(id, it.name)));
        node.addEventListener('click', function () { toggleSeries(id, it.name); });
      }
      var sw = document.createElement('span');
      sw.className = 'swatch' + (it.line ? ' swatch-line' : '');
      sw.style.background = it.color;
      sw.setAttribute('aria-hidden', 'true');
      var tx = document.createElement('span');
      tx.textContent = it.label;
      node.append(sw, tx);
      el.append(node);
    });
  }
  function toggleSeries(id, name) {
    var h = hidden(id);
    var items = LEGENDS[id]();
    var shown = items.filter(function (it) { return !h[it.name]; });
    if (!h[name] && shown.length === 1) { announce('Cần giữ ít nhất một chuỗi số liệu trên biểu đồ.'); return; }
    h[name] = !h[name];
    RENDER[id]();
    var btn = Array.prototype.find.call(document.querySelectorAll('[data-legend="' + id + '"] button'),
      function (b) { return b.textContent === items.filter(function (it) { return it.name === name; })[0].label; });
    if (btn) btn.focus();
    announce((h[name] ? 'Đã ẩn ' : 'Đã hiện ') + name + '.');
  }

  function sourceLine(id, text, ys) {
    var el = document.querySelector('[data-source="' + id + '"]');
    el.replaceChildren(document.createTextNode(text + ' Báo cáo: '));
    ys.forEach(function (y, i) {
      var a = document.createElement('a');
      a.href = D.filings[y].url; a.target = '_blank'; a.rel = 'noopener';
      a.textContent = '10-K FY' + y;
      a.setAttribute('aria-label', 'Mở báo cáo 10-K FY' + y + ' trên SEC EDGAR (thẻ mới)');
      el.append(a, document.createTextNode(i < ys.length - 1 ? ', ' : '.'));
    });
  }

  // ------------------------------------------------------------------ view 1 charts
  var LEGENDS = {
    share: function () {
      return [{ name: 'NIKE Direct', label: 'NIKE Direct', color: C.direct },
        { name: 'Bán buôn', label: 'Bán buôn (Wholesale)', color: C.wholesale },
        { name: 'Global Brand Divisions', label: 'Global Brand Divisions (dưới 0,3%)', color: C.gbd }];
    },
    channelRevenue: function () {
      return [{ name: 'NIKE Direct', label: 'NIKE Direct', color: C.direct },
        { name: 'Bán buôn', label: 'Bán buôn (Wholesale)', color: C.wholesale }];
    },
    growth: function () { return LEGENDS.channelRevenue(); },
    digital: function () {
      return [{ name: 'NIKE Direct', label: 'NIKE Direct (tổng)', color: C.directLight },
        { name: 'Digital', label: 'Trong đó: NIKE Brand Digital', color: C.digital }];
    },
    sga: function () {
      return [{ name: 'Vận hành chung', label: 'Vận hành chung', color: C.overhead },
        { name: 'Tạo nhu cầu', label: 'Tạo nhu cầu (marketing)', color: C.demand }];
    }
  };

  function renderShare() {
    var ys = years();
    var d = ys.map(function (y) { return kpi('direct_share', y); });
    var w = ys.map(function (y) { return kpi('wholesale_share', y); });
    var g = ys.map(function (y) { return kpi('gbd_share', y); });
    var o = base();
    o.grid = { left: 4, right: 4, top: 6, bottom: 4, containLabel: true };
    o.tooltip.formatter = function (ps) {
      var i = ps[0].dataIndex, y = ys[i];
      return tip('FY' + y + ' (' + period(y) + ')', [
        { color: C.direct, value: pct(d[i]), label: 'NIKE Direct, ' + bn(val('nike_direct_revenue', y)) + ' tỷ USD' },
        { color: C.wholesale, value: pct(w[i]), label: 'Bán buôn, ' + bn(val('nike_brand_wholesale_revenue', y)) + ' tỷ USD' },
        { color: C.gbd, value: pct(g[i], 2), label: 'Global Brand Divisions' }
      ], 'Mẫu số: doanh thu NIKE Brand ' + bn(val('nike_brand_revenue', y)) + ' tỷ USD. ' + srcNote('nike_direct_revenue', y));
    };
    o.xAxis = valAxis(function (v) { return num(v * 100, 0) + '%'; }, { min: 0, max: 1, interval: 0.25 });
    var label = function (side) {
      return {
        type: 'category', inverse: true, position: side, data: ys.map(function (y) { return 'FY' + y; }),
        axisTick: { show: false }, axisLine: { show: false },
        axisLabel: side === 'left'
          ? { formatter: function (v, i) { return '{fy|' + v + '}{v|' + pct(d[i]) + '}'; },
              rich: { fy: { color: C.muted, fontSize: 13, fontFamily: FONT, width: 56 }, v: { color: C.ink, fontSize: 14, fontWeight: 600, fontFamily: FONT, width: 48, align: 'right' } } }
          : { formatter: function (v, i) { return '{v|' + pct(w[i]) + '}'; },
              rich: { v: { color: C.ink, fontSize: 14, fontWeight: 600, fontFamily: FONT } } }
      };
    };
    o.yAxis = [label('left'), label('right')];
    var bar = function (name, data, color) {
      return { name: name, type: 'bar', stack: 'share', yAxisIndex: 0, data: data, barWidth: 22,
        itemStyle: { color: color, borderColor: '#fff', borderWidth: 1 }, emphasis: { itemStyle: { opacity: 0.85 } } };
    };
    o.series = [bar('NIKE Direct', d, C.direct), bar('Bán buôn', w, C.wholesale), bar('Global Brand Divisions', g, C.gbd)];
    legend('share', LEGENDS.share(), false);
    var a = ys[0], b = ys[ys.length - 1];
    setChart('share', o, 'Biểu đồ thanh 100% theo năm: tỷ trọng NIKE Direct trong doanh thu NIKE Brand là ' + pct(kpi('direct_share', a)) +
      ' ở FY' + a + ' và ' + pct(kpi('direct_share', b)) + ' ở FY' + b + '; phần còn lại chủ yếu là bán buôn.');
    sourceLine('share', 'Nguồn: thuyết minh “Revenues” trong Form 10-K (bảng phân tách doanh thu, cột Total NIKE Brand). Tỷ trọng là tính toán.', ys);
  }

  function renderChannelRevenue() {
    var ys = years();
    var o = base();
    o.xAxis = catAxis(ys, 'channelRevenue');
    o.yAxis = valAxis(function (v) { return num(v, 0); }, { min: 0 });
    o.tooltip.formatter = function (ps) {
      var y = ys[ps[0].dataIndex];
      return tip('FY' + y + ' (' + period(y) + ')', [
        { color: C.direct, value: bn(val('nike_direct_revenue', y)) + ' tỷ USD', label: 'NIKE Direct' },
        { color: C.wholesale, value: bn(val('nike_brand_wholesale_revenue', y)) + ' tỷ USD', label: 'Bán buôn' }
      ], 'Tổng NIKE Brand ' + bn(val('nike_brand_revenue', y)) + ' tỷ USD. ' + srcNote('nike_direct_revenue', y));
    };
    var fmt = function (v) { return num(v, 1); };
    o.series = [];
    if (visible('channelRevenue', 'NIKE Direct'))
      o.series.push(columnSeries('NIKE Direct', ys.map(function (y) { return val('nike_direct_revenue', y) / 1000; }), C.direct, { label: endLabel(ys, fmt) }));
    if (visible('channelRevenue', 'Bán buôn'))
      o.series.push(columnSeries('Bán buôn', ys.map(function (y) { return val('nike_brand_wholesale_revenue', y) / 1000; }), C.wholesale, { label: endLabel(ys, fmt) }));
    legend('channelRevenue', LEGENDS.channelRevenue(), true);
    var a = ys[0], b = ys[ys.length - 1];
    setChart('channelRevenue', o, 'Biểu đồ cột: doanh thu NIKE Direct từ ' + bn(val('nike_direct_revenue', a)) + ' (FY' + a + ') đến ' +
      bn(val('nike_direct_revenue', b)) + ' tỷ USD (FY' + b + '); bán buôn từ ' + bn(val('nike_brand_wholesale_revenue', a)) + ' đến ' +
      bn(val('nike_brand_wholesale_revenue', b)) + ' tỷ USD.');
    sourceLine('channelRevenue', 'Nguồn: thuyết minh “Revenues” trong Form 10-K.', ys);
  }

  function growthValues(channel, y) {
    if (state.basis === 'cn') return val(channel === 'direct' ? 'nike_direct_pct_change_cn' : 'nike_brand_wholesale_pct_change_cn', y);
    return kpi(channel === 'direct' ? 'direct_growth' : 'wholesale_growth', y);
  }
  function renderGrowth() {
    var ys = years();
    var cn = state.basis === 'cn';
    document.getElementById('growth-sub').textContent = cn
      ? '% so với năm trước loại trừ tỷ giá, NIKE Brand. Số NIKE công bố (non-GAAP).'
      : '% so với năm trước theo báo cáo, NIKE Brand. Tính toán từ 10-K.';
    var o = base();
    o.xAxis = catAxis(ys, 'growth');
    o.yAxis = valAxis(function (v) { return num(v * 100, 0) + '%'; });
    o.tooltip.formatter = function (ps) {
      var y = ys[ps[0].dataIndex];
      return tip('FY' + y + ' so với FY' + (y - 1), [
        { color: C.direct, value: signedPct(kpi('direct_growth', y)), label: 'NIKE Direct, theo báo cáo' },
        { color: C.direct, value: signedPct(val('nike_direct_pct_change_cn', y), 0), label: 'NIKE Direct, loại trừ tỷ giá' },
        { color: C.wholesale, value: signedPct(kpi('wholesale_growth', y)), label: 'Bán buôn, theo báo cáo' },
        { color: C.wholesale, value: signedPct(val('nike_brand_wholesale_pct_change_cn', y), 0), label: 'Bán buôn, loại trừ tỷ giá' }
      ], 'Đang vẽ: ' + (cn ? 'loại trừ tỷ giá (NIKE công bố trong MD&A).' : 'theo báo cáo (tính toán).') + ' ' + srcNote('nike_direct_pct_change_cn', y));
    };
    var bars = function (name, key, color) {
      return columnSeries(name, ys.map(function (y) {
        var v = growthValues(key, y);
        return { value: v, itemStyle: { borderRadius: v != null && v < 0 ? [0, 0, 4, 4] : [4, 4, 0, 0] } };
      }), color, { barGap: '15%' });
    };
    o.series = [];
    if (visible('growth', 'NIKE Direct')) o.series.push(bars('NIKE Direct', 'direct', C.direct));
    if (visible('growth', 'Bán buôn')) o.series.push(bars('Bán buôn', 'wholesale', C.wholesale));
    if (o.series.length) o.series[0].markLine = { silent: true, symbol: 'none', data: [{ yAxis: 0 }], label: { show: false }, lineStyle: { color: C.ink2, width: 1, type: 'solid' } };
    legend('growth', LEGENDS.growth(), true);
    document.querySelectorAll('#growth-basis input').forEach(function (r) { r.checked = r.value === state.basis; });
    var b = ys[ys.length - 1];
    setChart('growth', o, 'Biểu đồ cột phân kỳ: tăng trưởng ' + (cn ? 'loại trừ tỷ giá' : 'theo báo cáo') + ' của NIKE Direct và bán buôn; FY' + b +
      ': NIKE Direct ' + signedPct(growthValues('direct', b), cn ? 0 : 1) + ', bán buôn ' + signedPct(growthValues('wholesale', b), cn ? 0 : 1) + '.');
    sourceLine('growth', cn ? 'Nguồn: bảng doanh thu trong MD&A (Item 7), cột “% change excluding currency changes”; NIKE làm tròn đến số nguyên, chưa có đối chiếu độc lập.'
      : 'Nguồn: thuyết minh “Revenues” trong Form 10-K. Tăng trưởng tính từ số năm nay và số năm trước trong cùng 10-K; chênh lệch với % NIKE công bố không quá ±0,4 điểm %.', ys);
  }

  function renderDigital() {
    var ys = years();
    var o = base();
    o.xAxis = catAxis(ys, 'digital');
    o.yAxis = valAxis(function (v) { return num(v, 0); }, { min: 0 });
    o.tooltip.formatter = function (ps) {
      var y = ys[ps[0].dataIndex];
      var dig = ob('nike_brand_digital_sales', y);
      var restated = dig && dig.restatement && dig.restatement.indexOf('ĐIỀU CHỈNH') === 0;  // note written by the pipeline
      return tip('FY' + y + ' (' + period(y) + ')', [
        { color: C.directLight, value: bn(val('nike_direct_revenue', y)) + ' tỷ USD', label: 'NIKE Direct (tổng)' },
        { color: C.digital, value: num(val('nike_brand_digital_sales', y), 1) + ' tỷ USD', label: 'NIKE Brand Digital' },
        { color: C.digital, value: '≈ ' + pct(kpi('digital_share_of_direct', y), 0), label: 'Digital trong NIKE Direct' }
      ], (restated ? 'Số gốc theo 10-K FY' + y + '. ' + dig.restatement.replace('ĐIỀU CHỈNH HỒI TỐ: ', '') + ' ' : '') +
        'Digital chỉ được công bố làm tròn 0,1 tỷ USD trong MD&A. ' + srcNote('nike_brand_digital_sales', y));
    };
    o.series = [
      columnSeries('NIKE Direct', ys.map(function (y) { return val('nike_direct_revenue', y) / 1000; }), C.directLight,
        { barWidth: 22, label: endLabel(ys, function (v) { return num(v, 1); }) }),
      columnSeries('Digital', ys.map(function (y) { return val('nike_brand_digital_sales', y); }), C.digital,
        { barWidth: 22, barGap: '-100%', itemStyle: { color: C.digital, borderRadius: [0, 0, 0, 0] }, z: 3 })
    ];
    legend('digital', LEGENDS.digital(), false);
    var a = ys[0], b = ys[ys.length - 1];
    setChart('digital', o, 'Biểu đồ cột lồng nhau: doanh số NIKE Brand Digital từ ' + num(val('nike_brand_digital_sales', a), 1) + ' (FY' + a + ') đến ' +
      num(val('nike_brand_digital_sales', b), 1) + ' tỷ USD (FY' + b + '), là một phần của doanh thu NIKE Direct.');
    sourceLine('digital', 'Nguồn: MD&A (Item 7), đoạn về NIKE Direct; NIKE Direct từ thuyết minh “Revenues”. Số FY2023 là số gốc, được trình bày lại ở 10-K FY2024.', ys);
  }

  // ------------------------------------------------------------------ view 2 charts
  function renderRevenue() {
    var ys = years();
    var o = base();
    o.xAxis = catAxis(ys, 'revenue');
    o.yAxis = valAxis(function (v) { return num(v, 0); }, { min: 0 });
    o.tooltip.formatter = function (ps) {
      var y = ys[ps[0].dataIndex];
      return tip('FY' + y + ' (' + period(y) + ')', [
        { color: C.total, value: bn(val('revenues', y)) + ' tỷ USD', label: 'Doanh thu NIKE, Inc.' },
        { color: C.total, value: signedPct(kpi('revenue_growth', y)), label: 'so với năm trước, theo báo cáo' },
        { color: C.total, value: signedPct(val('revenues_pct_change_cn', y), 0), label: 'loại trừ tỷ giá (NIKE, non-GAAP)' }
      ], srcNote('revenues', y));
    };
    o.series = [columnSeries('Doanh thu', ys.map(function (y) { return val('revenues', y) / 1000; }), C.total,
      { label: endLabel(ys, function (v) { return num(v, 1); }) })];
    var a = ys[0], b = ys[ys.length - 1];
    setChart('revenue', o, 'Biểu đồ cột: doanh thu NIKE, Inc. từ ' + bn(val('revenues', a)) + ' (FY' + a + ') đến ' + bn(val('revenues', b)) + ' tỷ USD (FY' + b + ').');
    sourceLine('revenue', 'Nguồn: Consolidated Statements of Income.', ys);
  }

  function niceMin(values, step) {
    var m = Math.min.apply(null, values.filter(function (v) { return v != null; }));
    return Math.floor((m - step / 2) / step) * step;
  }

  function renderGrossMargin() {
    var ys = years();
    var data = ys.map(function (y) { return kpi('gross_margin', y); });
    var o = base();
    o.xAxis = catAxis(ys, 'grossMargin');
    o.xAxis.boundaryGap = true;
    o.yAxis = valAxis(function (v) { return num(v * 100, 0) + '%'; }, { min: niceMin(data, 0.02), scale: true });
    o.tooltip.axisPointer = { type: 'line', lineStyle: { color: C.axis } };
    o.tooltip.formatter = function (ps) {
      var y = ys[ps[0].dataIndex];
      return tip('FY' + y + ' (' + period(y) + ')', [
        { color: C.total, value: pct(kpi('gross_margin', y)), label: 'Biên lợi nhuận gộp' },
        { color: C.total, value: bn(val('gross_profit', y)) + ' tỷ USD', label: 'Lợi nhuận gộp' },
        { color: C.total, value: bn(val('revenues', y)) + ' tỷ USD', label: 'Doanh thu' }
      ], 'Biên gộp là tính toán. ' + srcNote('gross_profit', y));
    };
    o.series = [lineSeries('Biên lợi nhuận gộp', data, C.total, ys, function (v) { return pct(v); })];
    var a = ys[0], b = ys[ys.length - 1];
    setChart('grossMargin', o, 'Biểu đồ đường: biên lợi nhuận gộp từ ' + pct(kpi('gross_margin', a)) + ' (FY' + a + ') đến ' + pct(kpi('gross_margin', b)) + ' (FY' + b + ').');
    sourceLine('grossMargin', 'Nguồn: Consolidated Statements of Income; biên gộp là tính toán.', ys);
  }

  function renderSga() {
    var ys = years();
    var o = base();
    o.xAxis = catAxis(ys, 'sga');
    o.yAxis = valAxis(function (v) { return num(v * 100, 0) + '%'; }, { min: 0 });
    o.tooltip.formatter = function (ps) {
      var y = ys[ps[0].dataIndex];
      return tip('FY' + y + ' (' + period(y) + ')', [
        { color: C.ink2, value: pct(kpi('sga_ratio', y)), label: 'Tổng SG&A / doanh thu, ' + bn(val('total_selling_admin_expense', y)) + ' tỷ USD' },
        { color: C.overhead, value: pct(kpi('operating_overhead_ratio', y)), label: 'Vận hành chung' },
        { color: C.demand, value: pct(kpi('demand_creation_ratio', y)), label: 'Tạo nhu cầu' }
      ], 'Tỷ lệ là tính toán. ' + srcNote('total_selling_admin_expense', y));
    };
    var both = visible('sga', 'Vận hành chung') && visible('sga', 'Tạo nhu cầu');
    o.series = [];
    if (visible('sga', 'Vận hành chung'))
      o.series.push(columnSeries('Vận hành chung', ys.map(function (y) { return kpi('operating_overhead_ratio', y); }), C.overhead,
        { stack: 'sga', itemStyle: { color: C.overhead, borderColor: '#fff', borderWidth: 1, borderRadius: both ? 0 : [4, 4, 0, 0] } }));
    if (visible('sga', 'Tạo nhu cầu'))
      o.series.push(columnSeries('Tạo nhu cầu', ys.map(function (y) { return kpi('demand_creation_ratio', y); }), C.demand,
        { stack: 'sga', itemStyle: { color: C.demand, borderColor: '#fff', borderWidth: 1, borderRadius: [4, 4, 0, 0] },
          label: both ? {
            show: true, position: 'top', distance: 4, color: C.ink2, fontSize: 12.5, fontFamily: FONT,
            formatter: function (p) { return (p.dataIndex === 0 || p.dataIndex === ys.length - 1) ? pct(kpi('sga_ratio', ys[p.dataIndex])) : ''; }
          } : { show: false } }));
    legend('sga', LEGENDS.sga(), true);
    var a = ys[0], b = ys[ys.length - 1];
    setChart('sga', o, 'Biểu đồ cột chồng: tổng chi phí bán hàng và quản lý bằng ' + pct(kpi('sga_ratio', a)) + ' doanh thu ở FY' + a + ' và ' + pct(kpi('sga_ratio', b)) + ' ở FY' + b + '.');
    sourceLine('sga', 'Nguồn: Consolidated Statements of Income; tỷ lệ là tính toán. Nhãn trên cột là tổng SG&A / doanh thu.', ys);
  }

  function renderEbit() {
    var ys = years();
    var data = ys.map(function (y) { return kpi('ebit_margin', y); });
    var o = base();
    o.xAxis = catAxis(ys, 'ebit');
    o.xAxis.boundaryGap = true;
    o.yAxis = valAxis(function (v) { return num(v * 100, 0) + '%'; }, { min: 0 });
    o.tooltip.axisPointer = { type: 'line', lineStyle: { color: C.axis } };
    o.tooltip.formatter = function (ps) {
      var y = ys[ps[0].dataIndex];
      return tip('FY' + y + ' (' + period(y) + ')', [
        { color: C.total, value: pct(kpi('ebit_margin', y)), label: 'Biên EBIT (non-GAAP)' },
        { color: C.total, value: bn(val('ebit_reported', y), 2) + ' tỷ USD', label: 'EBIT do NIKE công bố' },
        { color: C.total, value: pct(kpi('net_income_margin', y)), label: 'Biên lợi nhuận sau thuế (GAAP)' }
      ], 'EBIT = lợi nhuận sau thuế + chi phí lãi thuần + thuế. ' + srcNote('ebit_reported', y));
    };
    o.series = [lineSeries('Biên EBIT', data, C.total, ys, function (v) { return pct(v); })];
    var a = ys[0], b = ys[ys.length - 1];
    setChart('ebit', o, 'Biểu đồ đường: biên EBIT từ ' + pct(kpi('ebit_margin', a)) + ' (FY' + a + ') đến ' + pct(kpi('ebit_margin', b)) + ' (FY' + b + ').');
    sourceLine('ebit', 'Nguồn: bảng EBIT trong MD&A (Item 7) và Consolidated Statements of Income; biên là tính toán.', ys);
  }

  function renderInventory() {
    var ys = years();
    var o = base();
    o.xAxis = catAxis(ys, 'inventory');
    o.yAxis = valAxis(function (v) { return num(v, 0); }, { min: 0 });
    o.tooltip.formatter = function (ps) {
      var y = ys[ps[0].dataIndex];
      return tip('Tại 31/05/' + y + ' (FY' + y + ')', [
        { color: C.total, value: bn(val('inventories', y), 2) + ' tỷ USD', label: 'Hàng tồn kho' },
        { color: C.total, value: signedPct(kpi('inventory_growth', y)), label: 'so với 31/05/' + (y - 1) },
        { color: C.total, value: num(kpi('inventory_turnover', y), 2) + ' lần', label: 'Vòng quay hàng tồn kho' }
      ], srcNote('inventories', y));
    };
    o.series = [columnSeries('Hàng tồn kho', ys.map(function (y) { return val('inventories', y) / 1000; }), C.total,
      { label: endLabel(ys, function (v) { return num(v, 1); }) })];
    var a = ys[0], b = ys[ys.length - 1];
    setChart('inventory', o, 'Biểu đồ cột: hàng tồn kho cuối kỳ từ ' + bn(val('inventories', a)) + ' (FY' + a + ') đến ' + bn(val('inventories', b)) + ' tỷ USD (FY' + b + ').');
    sourceLine('inventory', 'Nguồn: Consolidated Balance Sheets.', ys);
  }

  function renderDays() {
    var ys = years();
    var data = ys.map(function (y) { return kpi('days_inventory', y); });
    var o = base();
    o.xAxis = catAxis(ys, 'days');
    o.xAxis.boundaryGap = true;
    o.yAxis = valAxis(function (v) { return num(v, 0); }, { min: niceMin(data, 10), scale: true });
    o.tooltip.axisPointer = { type: 'line', lineStyle: { color: C.axis } };
    o.tooltip.formatter = function (ps) {
      var y = ys[ps[0].dataIndex];
      return tip('FY' + y + ' (' + period(y) + ')', [
        { color: C.total, value: num(kpi('days_inventory', y), 1) + ' ngày', label: 'Số ngày tồn kho bình quân' },
        { color: C.total, value: bn(kpi('average_inventory', y), 2) + ' tỷ USD', label: 'Tồn kho bình quân đầu và cuối kỳ' },
        { color: C.total, value: bn(val('cost_of_sales', y)) + ' tỷ USD', label: 'Giá vốn hàng bán' }
      ], 'Tính toán: 365 × tồn kho bình quân / giá vốn. ' + srcNote('inventories', y));
    };
    o.series = [lineSeries('Số ngày tồn kho', data, C.total, ys, function (v) { return num(v, 0); })];
    var a = ys[0], b = ys[ys.length - 1];
    setChart('days', o, 'Biểu đồ đường: số ngày tồn kho bình quân từ ' + num(kpi('days_inventory', a), 0) + ' (FY' + a + ') đến ' + num(kpi('days_inventory', b), 0) + ' ngày (FY' + b + ').');
    sourceLine('days', 'Nguồn: Consolidated Balance Sheets và Statements of Income; chỉ số là tính toán.', ys);
  }

  var RENDER = {
    share: renderShare, channelRevenue: renderChannelRevenue, growth: renderGrowth, digital: renderDigital,
    revenue: renderRevenue, grossMargin: renderGrossMargin, sga: renderSga, ebit: renderEbit, inventory: renderInventory, days: renderDays
  };
  var VIEW_CHARTS = {
    channels: ['share', 'channelRevenue', 'growth', 'digital'],
    performance: ['revenue', 'grossMargin', 'sga', 'ebit', 'inventory', 'days']
  };

  // ------------------------------------------------------------------ summaries
  function summary(el, parts) {
    el.replaceChildren();
    parts.forEach(function (p) {
      if (typeof p === 'string') el.append(document.createTextNode(p));
      else { var s = document.createElement('strong'); s.textContent = p.b; el.append(s); }
    });
  }
  function extreme(ys, f, sign) {
    var best = ys[0];
    ys.forEach(function (y) { if (sign * f(y) > sign * f(best)) best = y; });
    return best;
  }
  function renderSummaries() {
    var ys = years(), a = ys[0], b = ys[ys.length - 1];
    var share = function (y) { return kpi('direct_share', y); };
    var parts;
    if (a === b) {
      parts = ['Trong FY' + a + ', NIKE Direct chiếm ', { b: pct(share(a)) }, ' doanh thu NIKE Brand (', { b: bn(val('nike_direct_revenue', a)) + ' tỷ USD' },
        '); bán buôn chiếm ', { b: pct(kpi('wholesale_share', a)) }, ' (', { b: bn(val('nike_brand_wholesale_revenue', a)) + ' tỷ USD' }, ').'];
    } else {
      var hi = extreme(ys, share, 1), lo = extreme(ys, share, -1);
      parts = ['Tỷ trọng NIKE Direct trong doanh thu NIKE Brand là ', { b: pct(share(a)) }, ' ở FY' + a + ' và ', { b: pct(share(b)) }, ' ở FY' + b];
      if (ys.length > 2) parts = parts.concat(['; cao nhất ', { b: pct(share(hi)) }, ' (FY' + hi + '), thấp nhất ', { b: pct(share(lo)) }, ' (FY' + lo + ')']);
      parts = parts.concat(['. Doanh thu NIKE Direct từ ', { b: bn(val('nike_direct_revenue', a)) }, ' đến ', { b: bn(val('nike_direct_revenue', b)) + ' tỷ USD' },
        '; bán buôn từ ', { b: bn(val('nike_brand_wholesale_revenue', a)) }, ' đến ', { b: bn(val('nike_brand_wholesale_revenue', b)) + ' tỷ USD' }, '.']);
    }
    summary(document.getElementById('summary-channels'), parts);

    var em = function (y) { return kpi('ebit_margin', y); };
    if (a === b) {
      parts = ['Trong FY' + a + ', doanh thu NIKE, Inc. là ', { b: bn(val('revenues', a)) + ' tỷ USD' }, ', biên lợi nhuận gộp ', { b: pct(kpi('gross_margin', a)) },
        ', biên EBIT (non-GAAP) ', { b: pct(em(a)) }, ', hàng tồn kho cuối kỳ ', { b: bn(val('inventories', a)) + ' tỷ USD' }, '.'];
    } else {
      var ehi = extreme(ys, em, 1);
      parts = ['Doanh thu NIKE, Inc. từ ', { b: bn(val('revenues', a)) }, ' (FY' + a + ') đến ', { b: bn(val('revenues', b)) + ' tỷ USD' }, ' (FY' + b + '). Biên lợi nhuận gộp ',
        { b: pct(kpi('gross_margin', a)) }, ' và ', { b: pct(kpi('gross_margin', b)) }, '; biên EBIT (non-GAAP) ', { b: pct(em(a)) }, ' và ', { b: pct(em(b)) }];
      if (ys.length > 2) parts = parts.concat([', cao nhất ', { b: pct(em(ehi)) }, ' (FY' + ehi + ')']);
      parts = parts.concat(['. Hàng tồn kho cuối kỳ từ ', { b: bn(val('inventories', a)) }, ' đến ', { b: bn(val('inventories', b)) + ' tỷ USD' }, '.']);
    }
    summary(document.getElementById('summary-performance'), parts);
  }

  // ------------------------------------------------------------------ tables
  // Row spec: {g: group title} | {obs: id, unit, fmt, total} | {kpi: id, unit, fmt}
  var TABLES = {
    channels: [
      { g: 'Doanh thu NIKE Brand' },
      { obs: 'nike_direct_revenue', label: 'NIKE Direct', unit: 'triệu USD', fmt: amount },
      { obs: 'nike_brand_wholesale_revenue', label: 'Bán buôn (Wholesale)', unit: 'triệu USD', fmt: amount },
      { obs: 'global_brand_divisions_revenue', label: 'Global Brand Divisions', unit: 'triệu USD', fmt: amount },
      { obs: 'nike_brand_revenue', label: 'Tổng doanh thu NIKE Brand', unit: 'triệu USD', fmt: amount, total: true },
      { g: 'Cơ cấu theo kênh' },
      { kpi: 'direct_share', label: 'Tỷ trọng NIKE Direct', unit: '%', fmt: function (v) { return pct(v); } },
      { kpi: 'wholesale_share', label: 'Tỷ trọng bán buôn', unit: '%', fmt: function (v) { return pct(v); } },
      { obs: 'nike_direct_share_stated', label: 'Tỷ trọng NIKE Direct do NIKE nêu (xấp xỉ)', unit: '%', fmt: function (v) { return pct(v, 0); } },
      { g: 'Tăng trưởng so với năm trước' },
      { kpi: 'direct_growth', label: 'NIKE Direct, theo báo cáo', unit: '%', fmt: function (v) { return signedPct(v); } },
      { obs: 'nike_direct_pct_change_cn', label: 'NIKE Direct, loại trừ tỷ giá (non-GAAP)', unit: '%', fmt: function (v) { return signedPct(v, 0); } },
      { kpi: 'wholesale_growth', label: 'Bán buôn, theo báo cáo', unit: '%', fmt: function (v) { return signedPct(v); } },
      { obs: 'nike_brand_wholesale_pct_change_cn', label: 'Bán buôn, loại trừ tỷ giá (non-GAAP)', unit: '%', fmt: function (v) { return signedPct(v, 0); } },
      { kpi: 'nike_brand_growth', label: 'NIKE Brand, theo báo cáo', unit: '%', fmt: function (v) { return signedPct(v); } },
      { g: 'NIKE Brand Digital (một phần của NIKE Direct)' },
      { obs: 'nike_brand_digital_sales', label: 'Doanh số Digital', unit: 'tỷ USD', fmt: function (v) { return num(v, 1); } },
      { kpi: 'digital_share_of_direct', label: 'Digital trong NIKE Direct (xấp xỉ)', unit: '%', fmt: function (v) { return pct(v, 0); } }
    ],
    performance: [
      { g: 'Báo cáo kết quả kinh doanh' },
      { obs: 'revenues', label: 'Doanh thu', unit: 'triệu USD', fmt: amount, total: true },
      { obs: 'cost_of_sales', label: 'Giá vốn hàng bán', unit: 'triệu USD', fmt: amount },
      { obs: 'gross_profit', label: 'Lợi nhuận gộp', unit: 'triệu USD', fmt: amount, total: true },
      { obs: 'demand_creation_expense', label: 'Chi phí tạo nhu cầu', unit: 'triệu USD', fmt: amount },
      { obs: 'operating_overhead_expense', label: 'Chi phí vận hành chung', unit: 'triệu USD', fmt: amount },
      { obs: 'total_selling_admin_expense', label: 'Tổng chi phí bán hàng và quản lý', unit: 'triệu USD', fmt: amount, total: true },
      { obs: 'interest_expense_net', label: 'Chi phí (thu nhập) lãi thuần', unit: 'triệu USD', fmt: amount },
      { obs: 'other_expense_net', label: 'Chi phí (thu nhập) khác thuần', unit: 'triệu USD', fmt: amount },
      { obs: 'income_before_taxes', label: 'Lợi nhuận trước thuế', unit: 'triệu USD', fmt: amount, total: true },
      { obs: 'income_tax_expense', label: 'Chi phí thuế thu nhập', unit: 'triệu USD', fmt: amount },
      { obs: 'net_income', label: 'Lợi nhuận sau thuế', unit: 'triệu USD', fmt: amount, total: true },
      { g: 'Tỷ lệ trên doanh thu và tăng trưởng' },
      { kpi: 'revenue_growth', label: 'Tăng trưởng doanh thu, theo báo cáo', unit: '%', fmt: function (v) { return signedPct(v); } },
      { obs: 'revenues_pct_change_cn', label: 'Tăng trưởng doanh thu, loại trừ tỷ giá (non-GAAP)', unit: '%', fmt: function (v) { return signedPct(v, 0); } },
      { kpi: 'gross_margin', label: 'Biên lợi nhuận gộp', unit: '%', fmt: function (v) { return pct(v); } },
      { kpi: 'sga_ratio', label: 'Chi phí bán hàng và quản lý / doanh thu', unit: '%', fmt: function (v) { return pct(v); } },
      { kpi: 'demand_creation_ratio', label: 'Trong đó: tạo nhu cầu / doanh thu', unit: '%', fmt: function (v) { return pct(v); } },
      { kpi: 'operating_overhead_ratio', label: 'Trong đó: vận hành chung / doanh thu', unit: '%', fmt: function (v) { return pct(v); } },
      { kpi: 'net_income_margin', label: 'Biên lợi nhuận sau thuế', unit: '%', fmt: function (v) { return pct(v); } },
      { g: 'EBIT (non-GAAP, định nghĩa của NIKE)' },
      { obs: 'ebit_reported', label: 'EBIT do NIKE công bố', unit: 'triệu USD', fmt: amount },
      { kpi: 'ebit_margin', label: 'Biên EBIT', unit: '%', fmt: function (v) { return pct(v); } },
      { g: 'Hàng tồn kho' },
      { obs: 'inventories', label: 'Hàng tồn kho tại 31/5', unit: 'triệu USD', fmt: amount },
      { kpi: 'inventory_growth', label: 'Thay đổi so với đầu kỳ', unit: '%', fmt: function (v) { return signedPct(v); } },
      { kpi: 'inventory_turnover', label: 'Vòng quay hàng tồn kho', unit: 'lần', fmt: function (v) { return num(v, 2); } },
      { kpi: 'days_inventory', label: 'Số ngày tồn kho bình quân', unit: 'ngày', fmt: function (v) { return num(v, 1); } }
    ]
  };

  function cell(tag, text, cls) {
    var c = document.createElement(tag);
    if (text != null) c.textContent = text;
    if (cls) c.className = cls;
    return c;
  }
  function renderTable(view) {
    var ys = years();
    var table = document.getElementById('table-' + view);
    table.replaceChildren();
    var thead = document.createElement('thead'), hr = document.createElement('tr');
    hr.append(cell('th', 'Chỉ tiêu'), cell('th', 'Đơn vị'));
    ys.forEach(function (y) {
      var th = cell('th', 'FY' + y); th.scope = 'col'; th.title = period(y); hr.append(th);
    });
    hr.children[0].scope = 'col'; hr.children[1].scope = 'col';
    thead.append(hr);
    var tbody = document.createElement('tbody');
    TABLES[view].forEach(function (row) {
      var tr = document.createElement('tr');
      if (row.g) {
        tr.className = 'group';
        var th = cell('th', row.g); th.colSpan = ys.length + 2; th.scope = 'colgroup';
        tr.append(th); tbody.append(tr); return;
      }
      if (row.total) tr.className = 'total';
      var rh = cell('th', row.label); rh.scope = 'row';
      tr.append(rh, cell('td', row.unit, 'unit'));
      ys.forEach(function (y) {
        // Every numeric cell has the same two trailing slots (marker, page reference) so that
        // reported and calculated figures share one right edge.
        var td, mark = cell('span', null, 'mark-slot'), ref = cell('span', null, 'ref-slot');
        if (row.kpi) {
          var v = kpi(row.kpi, y);
          td = cell('td', v == null ? '—' : row.fmt(v), 'num calc');
          td.title = 'Tính toán: ' + D.kpis[row.kpi].vi + '. ' + D.kpis[row.kpi].note;
        } else {
          var o = ob(row.obs, y);
          if (!o || o.value == null) {
            td = cell('td', 'Không công bố', 'num missing');
            td.title = (o && o.cross_check) || 'NIKE không công bố số liệu này trong 10-K.';
          } else {
            td = cell('td', row.fmt(o.value), 'num');
            td.title = D.metrics[row.obs].en + '. ' + o.status + '. ' + (o.restatement || '');
            if (o.status === STATUS_PENDING) {
              mark.textContent = '†';
              mark.setAttribute('aria-label', 'chờ người kiểm tra');
            }
            if (o.page) {
              var a = document.createElement('a');
              a.className = 'page-ref'; a.href = D.filings[y].url; a.target = '_blank'; a.rel = 'noopener';
              a.textContent = 'tr. ' + o.page;
              a.setAttribute('aria-label', 'Mở 10-K FY' + y + ' trên SEC, trang ' + o.page);
              ref.append(a);
            }
          }
        }
        if (!mark.textContent) mark.setAttribute('aria-hidden', 'true');
        if (!ref.childNodes.length) ref.setAttribute('aria-hidden', 'true');
        td.append(mark, ref);
        tr.append(td);
      });
      tbody.append(tr);
    });
    table.append(thead, tbody);
  }

  function csvFor(view) {
    var ys = years();
    var rows = [['Chỉ tiêu', 'Mã', 'Loại', 'Đơn vị (giá trị thô)'].concat(ys.map(function (y) { return 'FY' + y; })).concat(['Nguồn'])];
    TABLES[view].forEach(function (row) {
      if (row.g) return;
      var id = row.kpi || row.obs;
      var unit = row.unit === '%' ? 'tỷ lệ thập phân' : row.unit;
      var vals = ys.map(function (y) {
        var v = row.kpi ? kpi(id, y) : val(id, y);
        return v == null ? '' : String(Math.round(v * 1e6) / 1e6);
      });
      var source = row.kpi ? 'Tính toán: ' + D.kpis[id].vi : D.metrics[id].location;
      rows.push([row.label, id, row.kpi ? 'Tính toán' : 'NIKE công bố', unit].concat(vals).concat([source]));
    });
    rows.push([]);
    rows.push(['Nguồn: NIKE, Inc. Form 10-K FY' + ys[0] + '–FY' + ys[ys.length - 1] + ' (SEC EDGAR). Trích xuất ' + D.meta.extracted_on + '.']);
    return '﻿' + rows.map(function (r) {
      return r.map(function (c) { var s = String(c == null ? '' : c); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }).join(',');
    }).join('\r\n');
  }
  function download(blob, name) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.append(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }
  function rangeTag() { return 'FY' + state.from + '-FY' + state.to; }

  // PNG export: the chart image with its title, unit line and source baked in, ready for slides.
  function exportPng(id) {
    var chart = charts[id];
    if (!chart) return;
    var fig = document.querySelector('.chart[data-chart="' + id + '"]').closest('figure');
    var title = fig.querySelector('h3').textContent;
    var sub = fig.querySelector('.figure-sub').textContent;
    var legendItems = Array.prototype.map.call(fig.querySelectorAll('.legend .legend-item'), function (n) {
      return n.getAttribute('aria-pressed') === 'false' ? null : { text: n.textContent, color: n.querySelector('.swatch').style.background };
    }).filter(Boolean);
    var src = fig.querySelector('.source').textContent;
    var scale = 2;
    var img = new Image();
    img.onload = function () {
      var pad = 28 * scale, width = img.width + pad * 2;
      var canvas = document.createElement('canvas');
      var ctx = canvas.getContext('2d');
      var lines = function (text, font, maxW) {
        ctx.font = font;
        var words = text.split(' '), out = [], line = '';
        words.forEach(function (w) { var t = line ? line + ' ' + w : w; if (ctx.measureText(t).width > maxW && line) { out.push(line); line = w; } else line = t; });
        if (line) out.push(line);
        return out;
      };
      var fTitle = '600 ' + 20 * scale + 'px ' + FONT, fSub = 14 * scale + 'px ' + FONT, fSrc = 12 * scale + 'px ' + FONT;
      var tL = lines(title, fTitle, img.width), sL = lines(sub, fSub, img.width);
      // Legend: colour swatch + label per series, wrapped to the image width.
      ctx.font = fSub;
      var legendRows = [], rowItems = [], rowW = 0;
      legendItems.forEach(function (it) {
        var w = 18 * scale + ctx.measureText(it.text).width + 20 * scale;
        if (rowItems.length && rowW + w > img.width) { legendRows.push(rowItems); rowItems = []; rowW = 0; }
        rowItems.push(it); rowW += w;
      });
      if (rowItems.length) legendRows.push(rowItems);
      var srcL = lines(src + ' Phân tích học thuật độc lập.', fSrc, img.width);
      var h = pad + tL.length * 26 * scale + sL.length * 20 * scale + legendRows.length * 22 * scale + 8 * scale + img.height + 10 * scale + srcL.length * 17 * scale + pad;
      canvas.width = width; canvas.height = h;
      ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, width, h);
      ctx.textBaseline = 'top';
      var yy = pad;
      ctx.fillStyle = C.ink; ctx.font = fTitle; tL.forEach(function (l) { ctx.fillText(l, pad, yy); yy += 26 * scale; });
      ctx.fillStyle = C.muted; ctx.font = fSub; sL.forEach(function (l) { ctx.fillText(l, pad, yy); yy += 20 * scale; });
      legendRows.forEach(function (row) {
        var x = pad;
        row.forEach(function (it) {
          ctx.fillStyle = it.color; ctx.fillRect(x, yy + 4 * scale, 12 * scale, 12 * scale);
          ctx.fillStyle = C.ink2; ctx.fillText(it.text, x + 18 * scale, yy);
          x += 18 * scale + ctx.measureText(it.text).width + 20 * scale;
        });
        yy += 22 * scale;
      });
      yy += 8 * scale;
      ctx.drawImage(img, pad, yy); yy += img.height + 10 * scale;
      ctx.fillStyle = C.muted; ctx.font = fSrc; srcL.forEach(function (l) { ctx.fillText(l, pad, yy); yy += 17 * scale; });
      canvas.toBlob(function (blob) { download(blob, 'NIKE_' + id + '_' + rangeTag() + '.png'); announce('Đã tạo ảnh PNG: ' + title + '.'); }, 'image/png');
    };
    img.src = chart.getDataURL({ type: 'png', pixelRatio: scale, backgroundColor: '#ffffff' });
  }

  // ------------------------------------------------------------------ controls
  var liveTimer = null;
  function announce(msg) {
    var el = document.getElementById('live');
    el.textContent = '';
    clearTimeout(liveTimer);
    liveTimer = setTimeout(function () { el.textContent = msg; }, 50);
  }
  function fillSelects() {
    ['year-from', 'year-to'].forEach(function (id) {
      var s = document.getElementById(id);
      s.replaceChildren();
      D.years.forEach(function (y) {
        var o = document.createElement('option'); o.value = y; o.textContent = 'FY' + y;
        o.disabled = id === 'year-from' ? y > state.to : y < state.from;
        s.append(o);
      });
      s.value = id === 'year-from' ? state.from : state.to;
    });
  }
  function showView(view, focusTab) {
    state.view = view;
    ['channels', 'performance'].forEach(function (v) {
      var tab = document.getElementById('tab-' + v), panel = document.getElementById('view-' + v);
      var on = v === view;
      tab.setAttribute('aria-selected', String(on));
      tab.tabIndex = on ? 0 : -1;
      panel.hidden = !on;
      if (on && focusTab) tab.focus();
    });
  }
  function renderAll() {
    fillSelects();
    renderSummaries();
    VIEW_CHARTS[state.view].forEach(function (id) { RENDER[id](); charts[id].resize(); });
    renderTable(state.view);
    writeURL();
  }
  function bind() {
    document.querySelectorAll('[role="tab"]').forEach(function (tab) {
      tab.addEventListener('click', function () { showView(tab.dataset.view); renderAll(); });
      tab.addEventListener('keydown', function (e) {
        if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft' && e.key !== 'Home' && e.key !== 'End') return;
        e.preventDefault();
        var next = state.view === 'channels' ? 'performance' : 'channels';
        if (e.key === 'Home') next = 'channels';
        if (e.key === 'End') next = 'performance';
        showView(next, true); renderAll();
      });
    });
    document.getElementById('year-from').addEventListener('change', function (e) {
      state.from = Number(e.target.value); renderAll(); announce('Đang hiển thị FY' + state.from + '–FY' + state.to + '.');
    });
    document.getElementById('year-to').addEventListener('change', function (e) {
      state.to = Number(e.target.value); renderAll(); announce('Đang hiển thị FY' + state.from + '–FY' + state.to + '.');
    });
    document.getElementById('reset').addEventListener('click', function () {
      state.from = D.years[0]; state.to = D.years[D.years.length - 1]; state.basis = 'reported'; state.hidden = {};
      renderAll(); announce('Đã đặt lại: FY' + state.from + '–FY' + state.to + ', tăng trưởng theo báo cáo, hiện tất cả chuỗi số liệu.');
    });
    document.querySelectorAll('#growth-basis input').forEach(function (r) {
      r.addEventListener('change', function () {
        state.basis = r.value; renderGrowth(); writeURL();
        announce(r.value === 'cn' ? 'Tăng trưởng loại trừ tỷ giá (non-GAAP).' : 'Tăng trưởng theo báo cáo.');
      });
    });
    var resizeTimer = null;
    window.addEventListener('resize', function () {
      // Axis label density depends on chart width, so redraw (not just rescale) after resizing.
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(function () { VIEW_CHARTS[state.view].forEach(function (id) { RENDER[id](); }); }, 200);
    });
    document.querySelectorAll('.export-png').forEach(function (b) {
      b.addEventListener('click', function () { exportPng(b.dataset.chart); });
    });
    ['channels', 'performance'].forEach(function (v) {
      document.getElementById('csv-' + v).addEventListener('click', function () {
        download(new Blob([csvFor(v)], { type: 'text/csv;charset=utf-8' }), 'NIKE_' + (v === 'channels' ? 'kenh-phan-phoi' : 'ket-qua-tai-chinh') + '_' + rangeTag() + '.csv');
        announce('Đã tạo file CSV.');
      });
    });
  }

  function renderFooter() {
    var list = document.getElementById('filing-list');
    D.years.forEach(function (y) {
      var f = D.filings[y];
      var li = document.createElement('li');
      var a = document.createElement('a');
      a.href = f.url; a.target = '_blank'; a.rel = 'noopener';
      a.textContent = 'Form 10-K FY' + y;
      var filed = isoDate(f.filed);
      var acc = document.createElement('span'); acc.className = 'nowrap'; acc.textContent = f.accession;
      li.append(a, document.createTextNode(', nộp ngày ' + filed + ', số '), acc);
      list.append(li);
    });
    var c = D.meta.status_counts;
    var total = Object.keys(c).reduce(function (s, k) { return s + c[k]; }, 0);
    document.getElementById('status-summary').textContent =
      total + ' số liệu báo cáo: ' + c['Khớp nguồn chính thức'] + ' khớp nguồn chính thức qua kiểm tra độc lập, ' +
      c['Chờ người kiểm tra'] + ' chờ người kiểm tra (†), ' + c['Thiếu / không công bố'] + ' không được công bố, ' +
      c['Đã xác minh (người)'] + ' đã được thành viên nhóm xác minh. Dữ liệu trích xuất ngày ' +
      isoDate(D.meta.extracted_on) + '.';
  }

  function fail(msg) {
    var el = document.getElementById('load-error');
    el.textContent = msg;
    el.hidden = false;
  }

  function start() {
    if (!window.echarts) { fail('Không tải được thư viện biểu đồ (assets/vendor/echarts-5.6.0.min.js). Kiểm tra lại thư mục site đã được triển khai đầy đủ.'); return; }
    fetch('data/nike-financials.json', { cache: 'no-cache' })
      .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
      .then(function (data) {
        D = data;
        readURL();
        return document.fonts ? document.fonts.ready : null;
      })
      .then(function () {
        bind();
        renderFooter();
        showView(state.view);
        renderAll();
        document.documentElement.dataset.ready = 'true';
      })
      .catch(function (e) {
        fail('Không tải được dữ liệu (data/nike-financials.json): ' + e.message +
          '. Nếu bạn mở file index.html trực tiếp trên máy, hãy chạy máy chủ web cục bộ: python3 -m http.server --directory site');
      });
  }

  document.addEventListener('DOMContentLoaded', start);
})();
