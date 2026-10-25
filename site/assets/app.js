/* NIKE FY2019–FY2021 dashboard: plain JavaScript + Apache ECharts.
   All numbers come from data/nike-financials.json, which the Python pipeline builds from the validated Excel
   workbook; the key findings and section notes are written there too, so the Excel and Power BI versions show the
   same words. Nothing here computes new financial values except differences between two reported figures, unit
   conversion (USD million to USD billion) and formatting. */
(function () {
  'use strict';

  var C = {
    direct: '#2a78d6', wholesale: '#eb6834', gbd: '#a3a9b2', directLight: '#86b6ef', digital: '#1c5cab',
    total: '#475467', margin: '#6250d6', up: '#1baf7a', down: '#c8453b', past: '#98a2b3',
    ink: '#15191e', ink2: '#3b4350', muted: '#5b6472', grid: '#e6e9ee', axis: '#c3c9d1', band: 'rgba(21,25,30,0.055)'
  };
  var FONT = '"Source Sans 3", system-ui, -apple-system, "Segoe UI", Arial, sans-serif';
  var REDUCED = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  var MINUS = '−';
  var COVID = 'FY2020Q4';
  var STATUS_PENDING = 'Chờ người kiểm tra';

  var D = null;
  var state = { change: 'FY2019-FY2021', bridge: 'FY2019-FY2021', table: 'A' };
  var charts = {};

  // ------------------------------------------------------------------ formatting
  var nfCache = {};
  function nf(d) {
    if (!nfCache[d]) nfCache[d] = new Intl.NumberFormat('vi-VN', { minimumFractionDigits: d, maximumFractionDigits: d });
    return nfCache[d];
  }
  function num(v, d) { if (v == null) return '—'; var s = nf(d || 0).format(Math.abs(v)); return v < 0 ? MINUS + s : s; }
  function signed(v, d) {
    if (v == null) return '—';
    var s = nf(d || 0).format(Math.abs(v));
    return (Math.round(Math.abs(v) * Math.pow(10, d || 0)) === 0 ? '' : v > 0 ? '+' : MINUS) + s;
  }
  function amount(v) { if (v == null) return '—'; var s = nf(0).format(Math.abs(v)); return v < 0 ? '(' + s + ')' : s; }
  function pct(v, d) { return v == null ? '—' : num(v * 100, d == null ? 1 : d) + '%'; }
  function signedPct(v, d) { return v == null ? '—' : signed(v * 100, d == null ? 1 : d) + '%'; }
  function pts(v, d) { return v == null ? '—' : signed(v * 100, d == null ? 1 : d) + ' điểm %'; }
  function bn(millions, d) { return millions == null ? '—' : num(millions / 1000, d == null ? 1 : d); }
  function decimals(v) { var s = String(v); return s.indexOf('.') < 0 ? 0 : s.length - s.indexOf('.') - 1; }
  // Digital sales as NIKE stated them: in millions for fiscal 2019 quarters (0,77), to 0.1 billion afterwards (1,1).
  function stated(v) { return v == null ? '—' : num(v, Math.max(1, decimals(Math.round(v * 1000) / 1000))); }
  var dateFmt = new Intl.DateTimeFormat('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric', timeZone: 'UTC' });
  function isoDate(iso) { return dateFmt.format(new Date(iso + 'T00:00:00Z')); }
  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; });
  }

  // ------------------------------------------------------------------ data access
  function A() { return D.periods.A; }
  function Q() { return D.periods.Q; }
  function ob(id, p) { return (D.observations[id] || {})[p] || null; }
  function val(id, p) { var o = ob(id, p); return o ? o.value : null; }
  function kpi(id, p) { return D.kpis[id].values[p]; }
  function label(p) { return D.labels[p]; }
  function quarter(p) { return +p.slice(7); }
  function fyOf(p) { return p.slice(0, 6); }
  function srcNote(id, p) {
    var o = ob(id, p);
    if (!o) return '';
    if (o.method === 'q4_derived')
      return 'Quý 4 tự tính: cả năm ở ' + o.source + (o.page ? ' (tr. ' + o.page + ')' : '') + ' trừ 9 tháng đầu ở ' + o.source2 +
        (o.page2 ? ' (tr. ' + o.page2 + ')' : '') + '.';
    return 'Nguồn: ' + o.source + (o.page ? ', tr. ' + o.page : '') + '.';
  }
  function title(p) { return label(p) + ' (' + D.dates[p] + ')'; }

  // ------------------------------------------------------------------ chart scaffolding
  function base() {
    return {
      animation: !REDUCED, animationDuration: 250, animationDurationUpdate: 250,
      textStyle: { fontFamily: FONT, color: C.ink2, fontSize: 13 },
      grid: { left: 4, right: 16, top: 26, bottom: 4, containLabel: true },
      tooltip: {
        trigger: 'axis', confine: true, backgroundColor: '#fff', borderColor: C.axis, borderWidth: 1, padding: [8, 10],
        textStyle: { color: C.ink, fontFamily: FONT, fontSize: 13 },
        axisPointer: { type: 'shadow', shadowStyle: { color: 'rgba(21,25,30,0.05)' } },
        extraCssText: 'box-shadow:0 2px 8px rgba(21,25,30,.12);border-radius:4px;max-width:320px;white-space:normal;'
      }
    };
  }
  function axisText(extra) { return Object.assign({ color: C.muted, fontSize: 12.5, fontFamily: FONT }, extra || {}); }
  function quarterAxis() {
    // "Q1" over its fiscal year once per year; the other quarters show the quarter only.
    return {
      type: 'category', data: Q(), axisTick: { show: false }, axisLine: { onZero: false, lineStyle: { color: C.axis } },
      axisLabel: axisText({ interval: 0, lineHeight: 15,
        formatter: function (p) { return quarter(p) === 1 ? 'Q1\n' + fyOf(p) : 'Q' + quarter(p); } })
    };
  }
  function valAxis(formatter, extra) {
    return Object.assign({
      type: 'value', axisLabel: axisText({ formatter: formatter }),
      splitLine: { lineStyle: { color: C.grid, width: 1, type: 'solid' } }, axisLine: { show: false }, axisTick: { show: false }
    }, extra || {});
  }
  function tip(head, rows, foot) {
    var h = '<div style="font-weight:600;margin-bottom:4px">' + esc(head) + '</div>';
    rows.forEach(function (r) {
      h += '<div style="display:flex;align-items:baseline;gap:8px;line-height:1.55">' +
        '<span style="display:inline-block;width:12px;height:3px;border-radius:2px;background:' + r.color + ';flex:none;transform:translateY(-3px)"></span>' +
        '<span style="font-weight:600;font-variant-numeric:tabular-nums;white-space:nowrap">' + esc(r.value) + '</span>' +
        '<span style="color:' + C.muted + '">' + esc(r.label) + '</span></div>';
    });
    if (foot) h += '<div style="margin-top:6px;color:' + C.muted + ';font-size:12px;line-height:1.4">' + esc(foot) + '</div>';
    return h;
  }
  function labelStyle(extra) {
    return Object.assign({ show: true, color: C.ink2, fontSize: 12.5, fontFamily: FONT }, extra || {});
  }
  // Selective direct labels: only the indexes given (first, last, the extreme, the quarter the story is about).
  function lineSeries(name, data, color, fmt, labelled, extra) {
    var items = data.map(function (v, i) {
      if (v == null || labelled.indexOf(i) < 0) return { value: v, label: { show: false } };
      // A label sits on the side away from the line's neighbouring segments, so it never crosses the line.
      var prev = data[i - 1], next = data[i + 1];
      var below = (prev == null || prev > v) && (next == null || next > v);
      return { value: v, label: { show: true, position: below ? 'bottom' : 'top', formatter: function () { return fmt(v); } } };
    });
    return Object.assign({
      name: name, type: 'line', data: items, symbol: 'circle', symbolSize: 8, showSymbol: true, connectNulls: false,
      lineStyle: { width: 2, color: color, cap: 'round', join: 'round' },
      itemStyle: { color: color, borderColor: '#fff', borderWidth: 2 },
      label: labelStyle({ distance: 5 }), emphasis: { scale: 1.25 }
    }, extra || {});
  }
  // Grey band over the whole category of the fourth quarter of FY2020 (a markArea from that quarter to itself would
  // have no width on a category axis). `y` is any value inside the axis range; the band always spans the full height.
  function covidBand(y) {
    return { type: 'custom', silent: true, z: 0, tooltip: { show: false }, data: [[COVID, y]], encode: { x: 0, y: 1 },
      renderItem: function (params, api) {
        var x = api.coord([api.value(0), api.value(1)])[0], w = api.size([1, 0])[0], r = params.coordSys;
        return { type: 'rect', shape: { x: x - w / 2, y: r.y, width: w, height: r.height }, style: { fill: C.band } };
      } };
  }
  function extremes(data, which) {
    var idx = 0;
    data.forEach(function (v, i) { if (v != null && (which > 0 ? v > data[idx] : v < data[idx])) idx = i; });
    return idx;
  }
  function niceMin(values, step) {
    var m = Math.min.apply(null, values.filter(function (v) { return v != null; }));
    return Math.floor((m - step / 2) / step) * step;
  }
  function niceMax(values, step) {
    var m = Math.max.apply(null, values.filter(function (v) { return v != null; }));
    return Math.ceil((m + step / 2) / step) * step;
  }

  function setChart(id, option, description) {
    var el = document.querySelector('.chart[data-chart="' + id + '"]');
    // At least two canvas pixels per CSS pixel, so a printed chart is as sharp as the text around it.
    if (!charts[id]) charts[id] = echarts.init(el, null, { renderer: 'canvas', devicePixelRatio: Math.max(2, window.devicePixelRatio || 1) });
    charts[id].setOption(option, { notMerge: true });
    el.setAttribute('aria-label', description + ' Số liệu chi tiết có trong bảng ở cuối trang.');
  }
  function legend(id, items) {
    var el = document.querySelector('[data-legend="' + id + '"]');
    if (!el) return;
    el.replaceChildren();
    items.forEach(function (it) {
      var node = document.createElement('span');
      node.className = 'legend-item';
      var sw = document.createElement('span');
      sw.className = 'swatch' + (it.shape ? ' swatch-' + it.shape : '');
      sw.style.background = it.color;
      if (it.hollow) { sw.style.background = '#fff'; sw.style.boxShadow = 'inset 0 0 0 2px ' + it.color; }
      sw.setAttribute('aria-hidden', 'true');
      var tx = document.createElement('span');
      tx.textContent = it.label;
      node.append(sw, tx);
      el.append(node);
    });
  }
  function sourceLine(id, text, filingIds) {
    var el = document.querySelector('[data-source="' + id + '"]');
    el.replaceChildren(document.createTextNode(text + (filingIds && filingIds.length ? ' Báo cáo: ' : '')));
    (filingIds || []).forEach(function (fid, i) {
      var a = document.createElement('a');
      a.href = D.filings[fid].url; a.target = '_blank'; a.rel = 'noopener';
      a.textContent = fid; a.className = 'nowrap';
      a.setAttribute('aria-label', 'Mở ' + fid + ' trên SEC EDGAR (thẻ mới)');
      el.append(a, document.createTextNode(i < filingIds.length - 1 ? ', ' : '.'));
    });
  }
  function annualFilings() { return A().map(function (p) { return '10-K ' + p; }); }
  function quarterSources(text) {
    return text + ' Báo cáo: 10-Q quý 1–3 và 10-K của FY2019–FY2021 (danh sách ở cuối trang).';
  }

  // ------------------------------------------------------------------ 1. channels
  function renderShare() {
    var ys = A();
    var d = ys.map(function (p) { return kpi('direct_share', p); });
    var w = ys.map(function (p) { return kpi('wholesale_share', p); });
    var g = ys.map(function (p) { return kpi('gbd_share', p); });
    var o = base();
    o.grid = { left: 4, right: 4, top: 6, bottom: 4, containLabel: true };
    o.tooltip.formatter = function (ps) {
      var p = ys[ps[0].dataIndex];
      return tip(title(p), [
        { color: C.direct, value: pct(kpi('direct_share', p)), label: 'NIKE Direct, ' + bn(val('nike_direct_revenue', p)) + ' tỷ USD' },
        { color: C.wholesale, value: pct(kpi('wholesale_share', p)), label: 'Bán buôn, ' + bn(val('nike_brand_wholesale_revenue', p)) + ' tỷ USD' },
        { color: C.gbd, value: pct(kpi('gbd_share', p), 2), label: 'Global Brand Divisions' }
      ], 'Mẫu số: doanh thu NIKE Brand ' + bn(val('nike_brand_revenue', p)) + ' tỷ USD. ' + srcNote('nike_direct_revenue', p));
    };
    o.xAxis = valAxis(function (v) { return num(v * 100, 0) + '%'; }, { min: 0, max: 1, interval: 0.25 });
    var side = function (where) {
      return {
        type: 'category', inverse: true, position: where, data: ys, axisTick: { show: false }, axisLine: { show: false },
        axisLabel: where === 'left'
          ? { formatter: function (v, i) { return '{fy|' + v + '}{v|' + pct(d[i]) + '}'; },
              rich: { fy: { color: C.muted, fontSize: 13, fontFamily: FONT, width: 56 }, v: { color: C.ink, fontSize: 14, fontWeight: 600, fontFamily: FONT, width: 48, align: 'right' } } }
          : { formatter: function (v, i) { return '{v|' + pct(w[i]) + '}'; },
              rich: { v: { color: C.ink, fontSize: 14, fontWeight: 600, fontFamily: FONT } } }
      };
    };
    o.yAxis = [side('left'), side('right')];
    var bar = function (name, data, color) {
      return { name: name, type: 'bar', stack: 'share', data: data, barWidth: 24,
        itemStyle: { color: color, borderColor: '#fff', borderWidth: 1 }, emphasis: { itemStyle: { opacity: 0.85 } } };
    };
    o.series = [bar('NIKE Direct', d, C.direct), bar('Bán buôn', w, C.wholesale), bar('Global Brand Divisions', g, C.gbd)];
    legend('share', [{ label: 'NIKE Direct', color: C.direct }, { label: 'Bán buôn (Wholesale)', color: C.wholesale },
      { label: 'Global Brand Divisions (dưới 0,2%)', color: C.gbd }]);
    setChart('share', o, 'Biểu đồ thanh 100% theo năm: tỷ trọng NIKE Direct trong doanh thu NIKE Brand là ' + d.map(function (v, i) {
      return pct(v) + ' ở ' + ys[i]; }).join(', ') + '; phần còn lại chủ yếu là bán buôn.');
    sourceLine('share', 'Nguồn: thuyết minh “Revenues” (bảng phân tách doanh thu, cột Total NIKE Brand). Tỷ trọng do nhóm tự tính.', annualFilings());
  }

  var CHANGE_ITEMS = [
    { name: 'NIKE Direct', id: 'nike_direct_revenue', color: C.direct },
    { name: 'Bán buôn', id: 'nike_brand_wholesale_revenue', color: C.wholesale },
    { name: 'Global Brand Divisions', id: 'global_brand_divisions_revenue', color: C.gbd },
    { name: 'Toàn NIKE Brand', id: 'nike_brand_revenue', color: C.total }
  ];
  function bnDelta(v) { return signed(v, Math.abs(v) < 0.1 ? 2 : 1); }
  function renderChange() {
    var pr = state.change.split('-'), a = pr[0], b = pr[1];
    var vals = CHANGE_ITEMS.map(function (it) { return (val(it.id, b) - val(it.id, a)) / 1000; });
    document.getElementById('change-sub').textContent = 'Tỷ USD, thay đổi từ ' + a + ' đến ' + b + '. Số âm là giảm.';
    var o = base();
    o.grid = { left: 4, right: 56, top: 6, bottom: 4, containLabel: true };
    o.tooltip.trigger = 'item';
    o.tooltip.formatter = function (pnt) {
      var it = CHANGE_ITEMS[pnt.dataIndex], va = val(it.id, a), vb = val(it.id, b);
      return tip(it.name + ', ' + a + ' → ' + b, [
        { color: it.color, value: bnDelta((vb - va) / 1000) + ' tỷ USD', label: 'thay đổi' },
        { color: it.color, value: bn(va, 2) + ' → ' + bn(vb, 2) + ' tỷ USD', label: '' },
        { color: it.color, value: signedPct(vb / va - 1), label: 'so với ' + a }
      ], 'Nguồn: thuyết minh “Revenues”, 10-K ' + a + ' và 10-K ' + b + '.');
    };
    // Every bar carries its figure, so the value axis shows no labels; it starts just short of the lowest bar (a
    // decline of 0,02 should not stretch the axis to −1).
    var lo = Math.min.apply(null, vals.concat(0)), hi = Math.max.apply(null, vals.concat(0));
    o.xAxis = valAxis(null, { min: lo < 0 ? lo - (hi - lo) * 0.04 : 0, max: hi, axisLabel: { show: false }, splitLine: { show: false } });
    o.yAxis = { type: 'category', inverse: true, data: CHANGE_ITEMS.map(function (it) { return it.name; }), axisTick: { show: false },
      axisLine: { show: false }, axisLabel: axisText({ fontSize: 13.5, color: C.ink2 }) };
    o.series = [{
      type: 'bar', barWidth: 24,
      // A negative figure is labelled on the right of the zero line: left of its bar it would run into the channel
      // names on a narrow screen.
      data: vals.map(function (v, i) {
        return { value: v, itemStyle: { color: CHANGE_ITEMS[i].color, borderRadius: v < 0 ? [4, 0, 0, 4] : [0, 4, 4, 0] },
          label: labelStyle({ position: 'right', distance: 6, fontWeight: i === 3 ? 600 : 400,
            formatter: function () { return bnDelta(v); } }) };
      }),
      markLine: { silent: true, symbol: 'none', label: { show: false }, data: [{ xAxis: 0 }], lineStyle: { color: C.ink2, width: 1, type: 'solid' } }
    }];
    setChart('change', o, 'Biểu đồ thanh: doanh thu NIKE Brand thay đổi từ ' + a + ' đến ' + b + ': ' + CHANGE_ITEMS.map(function (it, i) {
      return it.name + ' ' + bnDelta(vals[i]) + ' tỷ USD'; }).join(', ') + '.');
    sourceLine('change', 'Nguồn: thuyết minh “Revenues”. Thay đổi = số năm sau trừ số năm trước, mỗi năm lấy từ 10-K của năm đó.',
      ['10-K ' + a, '10-K ' + b]);
  }

  // ------------------------------------------------------------------ 2. quarters
  function renderShareQ() {
    var ps = Q(), data = ps.map(function (p) { return kpi('direct_share', p); });
    var o = base();
    o.tooltip.axisPointer = { type: 'line', z: 1, lineStyle: { color: C.axis } };
    o.tooltip.formatter = function (pp) {
      var p = ps[pp[0].dataIndex], said = val('nike_direct_share_stated', p);
      return tip(title(p), [
        { color: C.direct, value: pct(kpi('direct_share', p)), label: 'NIKE Direct / NIKE Brand' },
        { color: C.direct, value: bn(val('nike_direct_revenue', p), 2) + ' tỷ USD', label: 'NIKE Direct' },
        { color: C.total, value: bn(val('nike_brand_revenue', p), 2) + ' tỷ USD', label: 'NIKE Brand' }
      ], (said != null ? 'NIKE nêu: khoảng ' + pct(said, 0) + '. ' : '') + srcNote('nike_direct_revenue', p));
    };
    o.xAxis = quarterAxis();
    o.yAxis = valAxis(function (v) { return num(v * 100, 0) + '%'; }, { min: niceMin(data, 0.05), max: niceMax(data, 0.05) });
    var s = lineSeries('Tỷ trọng NIKE Direct', data, C.direct, function (v) { return pct(v); },
      [0, data.length - 1, extremes(data, 1)]);
    o.series = [s, covidBand(data[ps.indexOf(COVID)])];
    var hi = extremes(data, 1);
    setChart('shareQ', o, 'Biểu đồ đường 12 quý: tỷ trọng NIKE Direct từ ' + pct(data[0]) + ' (' + label(ps[0]) + ') đến ' +
      pct(data[data.length - 1]) + ' (' + label(ps[ps.length - 1]) + '), cao nhất ' + pct(data[hi]) + ' ở ' + label(ps[hi]) + '.');
    sourceLine('shareQ', quarterSources('Nguồn: thuyết minh “Revenues”. Tỷ trọng do nhóm tự tính; quý 4 = cả năm trừ 9 tháng đầu.'));
  }

  function renderChannelQ() {
    var ps = Q();
    var d = ps.map(function (p) { return val('nike_direct_revenue', p) / 1000; });
    var w = ps.map(function (p) { return val('nike_brand_wholesale_revenue', p) / 1000; });
    var c = ps.indexOf(COVID), last = ps.length - 1;
    var o = base();
    o.tooltip.axisPointer = { type: 'line', z: 1, lineStyle: { color: C.axis } };
    o.tooltip.formatter = function (pp) {
      var p = ps[pp[0].dataIndex];
      return tip(title(p), [
        { color: C.direct, value: bn(val('nike_direct_revenue', p), 2) + ' tỷ USD', label: 'NIKE Direct, ' + signedPct(kpi('direct_growth', p)) + ' so với cùng kỳ' },
        { color: C.wholesale, value: bn(val('nike_brand_wholesale_revenue', p), 2) + ' tỷ USD', label: 'Bán buôn, ' + signedPct(kpi('wholesale_growth', p)) + ' so với cùng kỳ' }
      ], srcNote('nike_direct_revenue', p));
    };
    o.xAxis = quarterAxis();
    o.yAxis = valAxis(function (v) { return num(v, 0); }, { min: 0 });
    var fmt = function (v) { return num(v, 1); };
    var sw = lineSeries('Bán buôn', w, C.wholesale, fmt, [0, c, last]);
    // In Q4 FY2020 the two lines come within a label's height of each other: this label goes to the left of its point,
    // slightly raised, clear of both lines.
    Object.assign(sw.data[c].label, { position: 'left', offset: [0, -6] });
    o.series = [lineSeries('NIKE Direct', d, C.direct, fmt, [0, c, last]), sw, covidBand(d[c])];
    legend('channelQ', [{ label: 'NIKE Direct', color: C.direct, shape: 'line' }, { label: 'Bán buôn (Wholesale)', color: C.wholesale, shape: 'line' }]);
    setChart('channelQ', o, 'Biểu đồ đường 12 quý: doanh thu bán buôn ' + num(w[c], 1) + ' tỷ USD và NIKE Direct ' + num(d[c], 1) +
      ' tỷ USD ở ' + label(COVID) + '; quý cuối ' + num(w[last], 1) + ' và ' + num(d[last], 1) + ' tỷ USD.');
    sourceLine('channelQ', quarterSources('Nguồn: thuyết minh “Revenues”.'));
  }

  // ------------------------------------------------------------------ 3. digital
  function renderDigital() {
    var ys = A();
    var dig = ys.map(function (p) { return val('nike_brand_digital_sales', p); });
    var rest = ys.map(function (p) { return kpi('direct_ex_digital', p) / 1000; });
    var o = base();
    o.grid = { left: 4, right: 48, top: 6, bottom: 4, containLabel: true };
    o.tooltip.formatter = function (pp) {
      var p = ys[pp[0].dataIndex];
      return tip(title(p), [
        { color: C.digital, value: stated(val('nike_brand_digital_sales', p)) + ' tỷ USD', label: 'Digital, ≈ ' + pct(kpi('digital_share_of_direct', p), 0) + ' NIKE Direct' },
        { color: C.directLight, value: '≈ ' + bn(kpi('direct_ex_digital', p)) + ' tỷ USD', label: 'NIKE Direct ngoài Digital (ước tính)' },
        { color: C.total, value: bn(val('nike_direct_revenue', p), 2) + ' tỷ USD', label: 'NIKE Direct' }
      ], 'Digital: ' + srcNote('nike_brand_digital_sales', p).replace('Nguồn: ', '') + ' NIKE Direct: thuyết minh “Revenues”.');
    };
    o.xAxis = valAxis(function (v) { return num(v, 0); }, { min: 0 });
    o.yAxis = { type: 'category', inverse: true, data: ys, axisTick: { show: false }, axisLine: { show: false },
      axisLabel: axisText({ fontSize: 13 }) };
    var inside = function (color) {
      return labelStyle({ position: 'inside', color: color, fontWeight: 600, formatter: function (pnt) { return num(pnt.value, 1); } });
    };
    o.series = [
      { name: 'Digital', type: 'bar', stack: 'direct', barWidth: 24, data: dig, label: inside('#fff'),
        itemStyle: { color: C.digital, borderColor: '#fff', borderWidth: 1 } },
      { name: 'Ngoài Digital', type: 'bar', stack: 'direct', barWidth: 24, data: rest,
        label: inside(C.ink), itemStyle: { color: C.directLight, borderColor: '#fff', borderWidth: 1, borderRadius: [0, 4, 4, 0] } },
      // Total at the end of each bar.
      { name: 'NIKE Direct', type: 'bar', stack: 'direct', barWidth: 24, data: ys.map(function () { return 0; }), silent: true,
        itemStyle: { color: 'transparent' }, tooltip: { show: false },
        label: labelStyle({ position: 'right', distance: 8, fontWeight: 600,
          formatter: function (pnt) { return num(val('nike_direct_revenue', ys[pnt.dataIndex]) / 1000, 1); } }) }
    ];
    legend('digital', [{ label: 'NIKE Brand Digital', color: C.digital }, { label: 'NIKE Direct ngoài Digital (ước tính)', color: C.directLight }]);
    setChart('digital', o, 'Biểu đồ thanh chồng: doanh số Digital ' + dig.map(function (v, i) { return stated(v) + ' tỷ USD ở ' + ys[i]; }).join(', ') +
      '; NIKE Direct ngoài Digital khoảng ' + rest.map(function (v) { return num(v, 1); }).join(', ') + ' tỷ USD.');
    sourceLine('digital', 'Nguồn: phần thảo luận của ban điều hành (MD&A), đoạn về NIKE Direct (số Digital do NIKE làm tròn), và thuyết minh “Revenues”. Nhãn cuối thanh là tổng NIKE Direct.',
      annualFilings());
  }

  function q4Uncertainty(p) {
    // Q4 Digital = year (rounded to 0.1 billion) minus nine months (in millions for fiscal 2019, else rounded to 0.1).
    var err = fyOf(p) === 'FY2019' ? 0.05 : 0.1;
    return err * 1000 / val('nike_direct_revenue', p);
  }
  function renderDigitalQ() {
    var ps = Q(), data = ps.map(function (p) { return kpi('digital_share_of_direct', p); });
    var o = base();
    o.tooltip.axisPointer = { type: 'line', z: 1, lineStyle: { color: C.axis } };
    o.tooltip.formatter = function (pp) {
      var p = ps[pp[0].dataIndex], q4 = quarter(p) === 4;
      return tip(title(p), [
        { color: C.digital, value: '≈ ' + pct(kpi('digital_share_of_direct', p), 0), label: 'Digital / NIKE Direct' },
        { color: C.digital, value: (q4 ? '≈ ' + num(val('nike_brand_digital_sales', p), 1) : stated(val('nike_brand_digital_sales', p))) + ' tỷ USD',
          label: 'Digital' + (q4 ? ' (ước tính)' : '') },
        { color: C.total, value: bn(val('nike_direct_revenue', p), 2) + ' tỷ USD', label: 'NIKE Direct' }
      ], (q4 ? 'Ước tính: Digital cả năm trừ 9 tháng đầu, ' + (fyOf(p) === 'FY2019' ? 'trong đó số cả năm đã làm tròn' : 'cả hai số đều đã làm tròn') +
        '; sai số khoảng ±' + num(q4Uncertainty(p) * 100, 0) + ' điểm %. ' : '') +
        srcNote('nike_brand_digital_sales', p));
    };
    o.xAxis = quarterAxis();
    o.yAxis = valAxis(function (v) { return num(v * 100, 0) + '%'; }, { min: 0, max: niceMax(data, 0.1) });
    var s = lineSeries('Digital / NIKE Direct', data, C.digital, function (v) { return '≈' + pct(v, 0); },
      [0, data.length - 1, extremes(data, 1)]);
    s.data.forEach(function (it, i) {
      if (quarter(ps[i]) === 4) it.itemStyle = { color: '#fff', borderColor: C.digital, borderWidth: 2 };
    });
    o.series = [s, covidBand(data[ps.indexOf(COVID)])];
    var hi = extremes(data, 1);
    setChart('digitalQ', o, 'Biểu đồ đường 12 quý: tỷ trọng Digital trong NIKE Direct khoảng ' + pct(data[0], 0) + ' ở ' + label(ps[0]) +
      ', cao nhất khoảng ' + pct(data[hi], 0) + ' ở ' + label(ps[hi]) + ' (ước tính), ' + pct(data[data.length - 1], 0) + ' ở ' + label(ps[ps.length - 1]) + '.');
    sourceLine('digitalQ', quarterSources('Nguồn: phần thảo luận của ban điều hành (MD&A) ở từng báo cáo (số Digital) và thuyết minh “Revenues”.'));
  }

  // ------------------------------------------------------------------ 4. regions
  function dumbbell(id, rows, fmt, step, legendLabels) {
    // rows: [{name, a, b, bold}]; a = FY2019 (grey), b = FY2021 (blue), joined by a thin line.
    var all = rows.reduce(function (s, r) { return s.concat([r.a, r.b]); }, []);
    var names = rows.map(function (r) { return r.name; });
    var o = base();
    o.grid = { left: 4, right: 20, top: 6, bottom: 4, containLabel: true };
    o.tooltip.trigger = 'item';
    o.xAxis = valAxis(function (v) { return num(v * 100, 0) + '%'; }, { min: Math.max(0, niceMin(all, step)), max: niceMax(all, step) });
    o.yAxis = { type: 'category', inverse: true, data: names, axisTick: { show: false }, axisLine: { show: false },
      axisLabel: axisText({ fontSize: 13.5, color: C.ink2, formatter: function (v, i) { return rows[i].bold ? '{b|' + v + '}' : v; },
        rich: { b: { fontWeight: 600, color: C.ink, fontSize: 13.5, fontFamily: FONT } } }) };
    var dot = function (key, color, other) {
      return { name: key === 'a' ? legendLabels[0] : legendLabels[1], type: 'scatter', symbolSize: 12, z: 3,
        itemStyle: { color: color, borderColor: '#fff', borderWidth: 2 },
        data: rows.map(function (r, i) {
          var left = key === 'a' ? r.a <= r.b : r.b < r.a;
          return { value: [r[key], i], label: labelStyle({ position: left ? 'left' : 'right', distance: 8,
            fontWeight: key === 'b' ? 600 : 400, formatter: function () { return fmt(r[key]); } }) };
        }) };
    };
    o.tooltip.formatter = function (pnt) {
      var r = rows[pnt.value[1]];
      return tip(r.name, [
        { color: C.past, value: fmt(r.a), label: legendLabels[0] },
        { color: C.direct, value: fmt(r.b), label: legendLabels[1] },
        { color: C.ink2, value: pts(r.b - r.a), label: 'thay đổi' }
      ], r.note || '');
    };
    o.series = [
      { type: 'custom', silent: true, z: 1, tooltip: { show: false },
        renderItem: function (params, api) {
          var s = api.coord([api.value(0), api.value(2)]), e = api.coord([api.value(1), api.value(2)]);
          return { type: 'line', shape: { x1: s[0], y1: s[1], x2: e[0], y2: e[1] }, style: { stroke: C.axis, lineWidth: 2 } };
        },
        data: rows.map(function (r, i) { return [r.a, r.b, i]; }), encode: { x: [0, 1], y: 2 } },
      dot('a', C.past), dot('b', C.direct)
    ];
    legend(id, [{ label: legendLabels[0], color: C.past, shape: 'dot' }, { label: legendLabels[1], color: C.direct, shape: 'dot' }]);
    return o;
  }
  // After drawing: widen the value axis to the next multiple of 10% until the label beside each outer dot fits inside the
  // plot (on a narrow screen it would otherwise run into the region names or past the right edge). Ends on multiples of
  // 10% keep the axis labels regular.
  function fitDumbbell(id, o, rows) {
    var chart = charts[id], room = 54;
    var px = function (v) { return chart.convertToPixel({ xAxisIndex: 0 }, v); };
    for (var k = 0; k < 4; k++) {
      var left = rows.some(function (r) { return px(Math.min(r.a, r.b)) - px(o.xAxis.min) < room; }) && o.xAxis.min > 0;
      var right = rows.some(function (r) { return px(o.xAxis.max) - px(Math.max(r.a, r.b)) < room; });
      if (!left && !right) return;
      if (left) o.xAxis.min = Math.max(0, (Math.ceil(o.xAxis.min * 10 - 1e-6) - 1) / 10);
      if (right) o.xAxis.max = (Math.floor(o.xAxis.max * 10 + 1e-6) + 1) / 10;
      chart.setOption(o, { notMerge: true });
    }
  }
  function regionRows(kid, total) {
    var rows = D.regions.map(function (r) {
      return { name: r.short, a: kpi(r.id + '_' + kid, 'FY2019'), b: kpi(r.id + '_' + kid, 'FY2021'), note: r.vi + ' (' + r.en + ').' };
    });
    if (total) rows.unshift({ name: total.name, a: kpi(total.kpi, 'FY2019'), b: kpi(total.kpi, 'FY2021'), bold: true, note: total.note });
    return rows;
  }
  function renderRegions() {
    var rows = regionRows('direct_share', { name: 'Toàn NIKE Brand', kpi: 'direct_share', note: 'Bốn khu vực cộng Global Brand Divisions.' });
    var o = dumbbell('regions', rows, function (v) { return pct(v); }, 0.05, ['FY2019', 'FY2021']);
    setChart('regions', o, 'Biểu đồ điểm: tỷ trọng NIKE Direct FY2019 và FY2021: ' + rows.map(function (r) {
      return r.name + ' ' + pct(r.a) + ' và ' + pct(r.b); }).join('; ') + '.');
    fitDumbbell('regions', o, rows);
    sourceLine('regions', 'Nguồn: thuyết minh “Revenues” (cột từng khu vực). Tỷ trọng do nhóm tự tính.', ['10-K FY2019', '10-K FY2021']);
  }
  function renderRegionsEbit() {
    var rows = regionRows('ebit_margin');
    var o = dumbbell('regionsEbit', rows, function (v) { return pct(v); }, 0.05, ['FY2019', 'FY2021']);
    setChart('regionsEbit', o, 'Biểu đồ điểm: biên EBIT khu vực FY2019 và FY2021: ' + rows.map(function (r) {
      return r.name + ' ' + pct(r.a) + ' và ' + pct(r.b); }).join('; ') + '.');
    fitDumbbell('regionsEbit', o, rows);
    sourceLine('regionsEbit', 'Nguồn: thuyết minh về các bộ phận kinh doanh (“Operating Segments”). Biên do nhóm tự tính: EBIT khu vực / doanh thu khu vực.',
      ['10-K FY2019', '10-K FY2021']);
  }

  // ------------------------------------------------------------------ 5. costs and profit
  function renderBridge() {
    var pr = state.bridge.split('-'), a = pr[0], b = pr[1];
    var start = kpi('ebit_margin', a), end = kpi('ebit_margin', b);
    var steps = [
      { name: 'Biên gộp', d: kpi('gross_margin', b) - kpi('gross_margin', a), from: kpi('gross_margin', a), to: kpi('gross_margin', b), what: 'biên lợi nhuận gộp' },
      { name: 'Tạo nhu cầu', d: -(kpi('demand_creation_ratio', b) - kpi('demand_creation_ratio', a)), from: kpi('demand_creation_ratio', a), to: kpi('demand_creation_ratio', b), what: 'chi phí tạo nhu cầu / doanh thu' },
      { name: 'Vận hành chung', d: -(kpi('operating_overhead_ratio', b) - kpi('operating_overhead_ratio', a)), from: kpi('operating_overhead_ratio', a), to: kpi('operating_overhead_ratio', b), what: 'chi phí vận hành chung / doanh thu' }
    ];
    // The last step closes the bridge exactly: other (income) expense, plus rounding in the statement totals.
    var sum = steps.reduce(function (s, x) { return s + x.d; }, 0);
    steps.push({ name: 'Khác', d: end - start - sum, from: kpi('other_expense_ratio', a), to: kpi('other_expense_ratio', b), what: 'chi phí (thu nhập) khác thuần / doanh thu' });
    var cats = [a].concat(steps.map(function (s) { return s.name; }), [b]);
    var baseData = [], barData = [], running = start;
    baseData.push(0); barData.push({ value: start, itemStyle: { color: C.total, borderRadius: [4, 4, 0, 0] } });
    steps.forEach(function (s) {
      var up = s.d >= 0;
      baseData.push(up ? running : running + s.d);
      barData.push({ value: Math.abs(s.d), itemStyle: { color: up ? C.up : C.down, borderRadius: up ? [4, 4, 0, 0] : [0, 0, 4, 4] } });
      running += s.d;
    });
    baseData.push(0); barData.push({ value: end, itemStyle: { color: C.total, borderRadius: [4, 4, 0, 0] } });
    document.getElementById('bridge-sub').textContent = 'NIKE, Inc. Cột đầu và cuối: biên EBIT năm ' + a + ' và ' + b +
      ' (% doanh thu). Cột giữa: mức thay đổi do từng khoản (điểm %); khoản chi phí nào giảm so với doanh thu thì làm biên EBIT tăng.';
    var o = base();
    o.tooltip.formatter = function (pp) {
      var i = pp[0].dataIndex;
      if (i === 0 || i === cats.length - 1) {
        var p = i === 0 ? a : b;
        return tip('Biên EBIT ' + p, [{ color: C.total, value: pct(kpi('ebit_margin', p)), label: 'EBIT / doanh thu' },
          { color: C.total, value: bn(val('ebit_reported', p), 2) + ' tỷ USD', label: 'EBIT do NIKE công bố' }], srcNote('ebit_reported', p));
      }
      var s = steps[i - 1];
      return tip(s.name, [{ color: s.d >= 0 ? C.up : C.down, value: pts(s.d), label: 'tác động lên biên EBIT' },
        { color: C.ink2, value: pct(s.from) + ' → ' + pct(s.to), label: s.what }],
        s.name === 'Khác' ? 'Gồm cả chênh lệch làm tròn của các dòng tổng trên báo cáo.' : 'Tự tính từ báo cáo kết quả kinh doanh.');
    };
    // On a narrow chart the step names go on two lines of about the same length, so neighbouring names never overlap.
    var el = document.querySelector('.chart[data-chart="bridge"]');
    var narrow = el.clientWidth / cats.length < 100;
    var twoLines = { 'Biên gộp': 'Biên\ngộp', 'Tạo nhu cầu': 'Tạo\nnhu cầu', 'Vận hành chung': 'Vận hành\nchung' };
    o.xAxis = { type: 'category', data: cats, axisTick: { show: false }, axisLine: { lineStyle: { color: C.axis } },
      axisLabel: axisText({ interval: 0, fontSize: 12.5, lineHeight: 15,
        formatter: function (v) { return narrow ? twoLines[v] || v : v; } }) };
    var tops = baseData.map(function (b, i) { return b + barData[i].value; });
    o.yAxis = valAxis(function (v) { return num(v * 100, 0) + '%'; }, { min: 0, max: niceMax(tops, 0.02) });
    o.series = [
      { type: 'bar', stack: 'w', data: baseData, itemStyle: { color: 'transparent' }, silent: true, tooltip: { show: false }, barWidth: 24 },
      { type: 'bar', stack: 'w', data: barData, barWidth: 24,
        label: labelStyle({ position: 'top', distance: 4, formatter: function (pnt) {
          var i = pnt.dataIndex;
          if (i === 0) return pct(start);
          if (i === cats.length - 1) return pct(end);
          return signed(steps[i - 1].d * 100, 1);
        } }) }
    ];
    // The legend lists only the colours this period shows.
    var keys = [{ label: 'Biên EBIT', color: C.total }];
    if (steps.some(function (s) { return s.d >= 0; })) keys.push({ label: 'Làm biên EBIT tăng', color: C.up });
    if (steps.some(function (s) { return s.d < 0; })) keys.push({ label: 'Làm biên EBIT giảm', color: C.down });
    legend('bridge', keys);
    setChart('bridge', o, 'Biểu đồ thác nước: biên EBIT từ ' + pct(start) + ' (' + a + ') đến ' + pct(end) + ' (' + b + '); ' +
      steps.map(function (s) { return s.name + ' ' + pts(s.d); }).join(', ') + '.');
    sourceLine('bridge', 'Nguồn: báo cáo kết quả kinh doanh và bảng EBIT ở phần thảo luận của ban điều hành (MD&A). Các tỷ lệ do nhóm tự tính.', ['10-K ' + a, '10-K ' + b]);
  }

  function renderMarginsQ() {
    var ps = Q();
    var gm = ps.map(function (p) { return kpi('gross_margin', p); });
    var em = ps.map(function (p) { return kpi('ebit_margin', p); });
    var c = ps.indexOf(COVID), last = ps.length - 1;
    var o = base();
    o.tooltip.axisPointer = { type: 'line', z: 1, lineStyle: { color: C.axis } };
    o.tooltip.formatter = function (pp) {
      var p = ps[pp[0].dataIndex];
      return tip(title(p), [
        { color: C.total, value: pct(kpi('gross_margin', p)), label: 'Biên lợi nhuận gộp' },
        { color: C.margin, value: pct(kpi('ebit_margin', p)), label: 'Biên EBIT (non-GAAP)' },
        { color: C.ink2, value: bn(val('revenues', p), 2) + ' tỷ USD', label: 'Doanh thu NIKE, Inc.' }
      ], srcNote('gross_profit', p));
    };
    o.xAxis = quarterAxis();
    o.yAxis = valAxis(function (v) { return num(v * 100, 0) + '%'; }, { min: niceMin(em, 0.1), max: niceMax(gm, 0.1) });
    var fmt = function (v) { return pct(v); };
    var sg = lineSeries('Biên lợi nhuận gộp', gm, C.total, fmt, [0, c, last]);
    var se = lineSeries('Biên EBIT', em, C.margin, fmt, [0, c, last]);
    se.markLine = { silent: true, symbol: 'none', label: { show: false }, data: [{ yAxis: 0 }], lineStyle: { color: C.ink2, width: 1, type: 'solid' } };
    o.series = [sg, se, covidBand(gm[c])];
    legend('marginsQ', [{ label: 'Biên lợi nhuận gộp', color: C.total, shape: 'line' }, { label: 'Biên EBIT (non-GAAP)', color: C.margin, shape: 'line' }]);
    setChart('marginsQ', o, 'Biểu đồ đường 12 quý: biên gộp ' + pct(gm[c]) + ' và biên EBIT ' + pct(em[c]) + ' ở ' + label(COVID) +
      '; quý cuối ' + pct(gm[last]) + ' và ' + pct(em[last]) + '.');
    sourceLine('marginsQ', quarterSources('Nguồn: báo cáo kết quả kinh doanh. Biên do nhóm tự tính; quý 4 = cả năm trừ 9 tháng đầu.'));
  }

  function renderInventoryQ() {
    var ps = Q(), data = ps.map(function (p) { return val('inventories', p) / 1000; });
    var hi = extremes(data, 1), last = ps.length - 1;
    var o = base();
    o.tooltip.formatter = function (pp) {
      var p = ps[pp[0].dataIndex];
      return tip(label(p) + ', số dư cuối quý', [
        { color: C.total, value: bn(val('inventories', p), 2) + ' tỷ USD', label: 'Hàng tồn kho' },
        { color: C.total, value: signedPct(kpi('inventory_growth', p)), label: 'so với cùng thời điểm năm trước' }
      ], srcNote('inventories', p));
    };
    o.xAxis = quarterAxis();
    o.yAxis = valAxis(function (v) { return num(v, 0); }, { min: 0 });
    o.series = [{
      type: 'bar', barMaxWidth: 24, itemStyle: { color: C.total, borderRadius: [4, 4, 0, 0] },
      data: data.map(function (v, i) {
        return { value: v, label: labelStyle({ show: [0, hi, last].indexOf(i) >= 0, position: 'top', distance: 4,
          formatter: function () { return num(v, 1); } }) };
      })
    }, covidBand(data[ps.indexOf(COVID)])];
    setChart('inventoryQ', o, 'Biểu đồ cột 12 quý: hàng tồn kho cuối quý từ ' + num(data[0], 1) + ' tỷ USD (' + label(ps[0]) +
      '), cao nhất ' + num(data[hi], 1) + ' tỷ USD (' + label(ps[hi]) + '), cuối kỳ ' + num(data[last], 1) + ' tỷ USD.');
    sourceLine('inventoryQ', quarterSources('Nguồn: bảng cân đối kế toán. Tồn kho cuối quý 4 là tồn kho ngày 31/5 in trong báo cáo năm (10-K).'));
  }

  function renderDaysQ() {
    var ps = Q(), data = ps.map(function (p) { return kpi('days_inventory', p); });
    var hi = extremes(data, 1), last = ps.length - 1;
    var o = base();
    o.tooltip.axisPointer = { type: 'line', z: 1, lineStyle: { color: C.axis } };
    o.tooltip.formatter = function (pp) {
      var p = ps[pp[0].dataIndex];
      return tip(title(p), [
        { color: C.total, value: num(kpi('days_inventory', p), 0) + ' ngày', label: 'Số ngày tồn kho' },
        { color: C.total, value: bn(kpi('average_inventory', p), 2) + ' tỷ USD', label: 'Tồn kho bình quân đầu và cuối quý' },
        { color: C.total, value: bn(val('cost_of_sales', p), 2) + ' tỷ USD', label: 'Giá vốn hàng bán của quý' }
      ], 'Cách tính: 91,25 × tồn kho bình quân / giá vốn của quý. ' + srcNote('cost_of_sales', p));
    };
    o.xAxis = quarterAxis();
    o.yAxis = valAxis(function (v) { return num(v, 0); }, { min: niceMin(data, 20), max: niceMax(data, 20) });
    var s = lineSeries('Số ngày tồn kho', data, C.total, function (v) { return num(v, 0); }, [0, hi, last]);
    o.series = [s, covidBand(data[ps.indexOf(COVID)])];
    setChart('daysQ', o, 'Biểu đồ đường 12 quý: số ngày tồn kho từ ' + num(data[0], 0) + ' (' + label(ps[0]) + '), cao nhất ' +
      num(data[hi], 0) + ' ngày (' + label(ps[hi]) + '), cuối kỳ ' + num(data[last], 0) + ' ngày.');
    sourceLine('daysQ', quarterSources('Nguồn: bảng cân đối kế toán và báo cáo kết quả kinh doanh. Chỉ số do nhóm tự tính.'));
  }

  var RENDER = {
    share: renderShare, change: renderChange, shareQ: renderShareQ, channelQ: renderChannelQ, digital: renderDigital,
    digitalQ: renderDigitalQ, regions: renderRegions, regionsEbit: renderRegionsEbit, bridge: renderBridge,
    marginsQ: renderMarginsQ, inventoryQ: renderInventoryQ, daysQ: renderDaysQ
  };
  function renderCharts() { Object.keys(RENDER).forEach(function (id) { RENDER[id](); }); }

  // ------------------------------------------------------------------ findings, notes, quotes
  function cell(tag, text, cls) {
    var c = document.createElement(tag);
    if (text != null) c.textContent = text;
    if (cls) c.className = cls;
    return c;
  }
  var SECTION_NAMES = { kenh: 'Kênh phân phối', quy: 'Theo quý', digital: 'Digital và cửa hàng', 'khu-vuc': 'Khu vực', 'loi-nhuan': 'Chi phí và lợi nhuận' };
  function renderFindings() {
    var box = document.getElementById('findings');
    box.replaceChildren();
    D.findings.forEach(function (f) {
      var art = cell('article', null, 'finding');
      var foot = cell('p', f.source + ' ', 'finding-foot');
      var a = cell('a', 'Xem biểu đồ: ' + SECTION_NAMES[f.target]);
      a.href = '#' + f.target;
      foot.append(a);
      art.append(cell('h3', f.label), cell('p', f.value, 'finding-value'), cell('p', f.text, 'finding-text'), foot);
      box.append(art);
    });
    document.getElementById('note-quy').textContent = D.notes.quy;
    document.getElementById('note-ton-kho').textContent = D.notes['ton-kho'];
  }
  function renderQuotes() {
    var box = document.getElementById('quotes');
    box.replaceChildren();
    D.quotes.forEach(function (q) {
      var fig = cell('figure', null, 'quote');
      var bq = cell('blockquote', q.text);
      bq.lang = 'en';
      var cap = cell('figcaption', 'NIKE, ' + q.filing + ', tr. ' + q.page + ' · ');
      var a = cell('a', 'mở trên SEC');
      a.href = q.url; a.target = '_blank'; a.rel = 'noopener';
      a.setAttribute('aria-label', 'Mở ' + q.filing + ' trên SEC EDGAR (thẻ mới), trang ' + q.page);
      cap.append(a);
      fig.append(cell('p', q.topic, 'quote-topic'), bq, cell('p', q.vi, 'quote-vi'), cap);
      box.append(fig);
    });
  }

  // ------------------------------------------------------------------ tables
  function fmtPct1(v) { return pct(v); }
  function fmtPct0(v) { return pct(v, 0); }
  function fmtGrowth(v) { return signedPct(v); }
  function fmtGrowth0(v) { return signedPct(v, 0); }
  function fmtPts(v) { return signed(v * 100, 1); }
  // Row spec: {g: group title} | {obs: id, label, unit, fmt, total} | {kpi: id, label, unit, fmt}
  var CHANNEL_ROWS = [
    { g: 'Doanh thu NIKE Brand theo kênh' },
    { obs: 'nike_direct_revenue', label: 'NIKE Direct', unit: 'triệu USD', fmt: amount },
    { obs: 'nike_brand_wholesale_revenue', label: 'Bán buôn (Wholesale)', unit: 'triệu USD', fmt: amount },
    { obs: 'global_brand_divisions_revenue', label: 'Global Brand Divisions', unit: 'triệu USD', fmt: amount },
    { obs: 'nike_brand_revenue', label: 'Tổng doanh thu NIKE Brand', unit: 'triệu USD', fmt: amount, total: true },
    { g: 'Cơ cấu và tăng trưởng' },
    { kpi: 'direct_share', label: 'Tỷ trọng NIKE Direct', unit: '%', fmt: fmtPct1 },
    { obs: 'nike_direct_share_stated', label: 'Tỷ trọng NIKE Direct do NIKE nêu (xấp xỉ)', unit: '%', fmt: fmtPct0 },
    { kpi: 'direct_growth', label: 'Tăng trưởng NIKE Direct, theo USD', unit: '%', fmt: fmtGrowth },
    { obs: 'nike_direct_pct_change_cn', label: 'Tăng trưởng NIKE Direct, loại trừ tỷ giá (non-GAAP)', unit: '%', fmt: fmtGrowth0 },
    { kpi: 'wholesale_growth', label: 'Tăng trưởng bán buôn, theo USD', unit: '%', fmt: fmtGrowth },
    { obs: 'nike_brand_wholesale_pct_change_cn', label: 'Tăng trưởng bán buôn, loại trừ tỷ giá (non-GAAP)', unit: '%', fmt: fmtGrowth0 },
    { kpi: 'nike_brand_growth', label: 'Tăng trưởng NIKE Brand, theo USD', unit: '%', fmt: fmtGrowth },
    { kpi: 'direct_contribution', label: 'Trong đó: đóng góp của NIKE Direct', unit: 'điểm %', fmt: fmtPts },
    { kpi: 'wholesale_contribution', label: 'Trong đó: đóng góp của bán buôn', unit: 'điểm %', fmt: fmtPts },
    { g: 'NIKE Brand Digital (một phần của NIKE Direct)' },
    { obs: 'nike_brand_digital_sales', label: 'Doanh số Digital', unit: 'tỷ USD', fmt: stated },
    { kpi: 'digital_share_of_direct', label: 'Digital trong NIKE Direct (xấp xỉ)', unit: '%', fmt: fmtPct0 },
    { kpi: 'direct_ex_digital', label: 'NIKE Direct ngoài Digital (ước tính)', unit: 'triệu USD', fmt: amount }
  ];
  var FINANCE_ROWS = [
    { g: 'Báo cáo kết quả kinh doanh (NIKE, Inc.)' },
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
    { kpi: 'revenue_growth', label: 'Tăng trưởng doanh thu, theo USD', unit: '%', fmt: fmtGrowth },
    { obs: 'revenues_pct_change_cn', label: 'Tăng trưởng doanh thu, loại trừ tỷ giá (non-GAAP)', unit: '%', fmt: fmtGrowth0 },
    { kpi: 'gross_margin', label: 'Biên lợi nhuận gộp', unit: '%', fmt: fmtPct1 },
    { kpi: 'demand_creation_ratio', label: 'Chi phí tạo nhu cầu / doanh thu', unit: '%', fmt: fmtPct1 },
    { kpi: 'operating_overhead_ratio', label: 'Chi phí vận hành chung / doanh thu', unit: '%', fmt: fmtPct1 },
    { kpi: 'net_income_margin', label: 'Biên lợi nhuận sau thuế', unit: '%', fmt: fmtPct1 },
    { g: 'EBIT (non-GAAP, định nghĩa của NIKE)' },
    { obs: 'ebit_reported', label: 'EBIT do NIKE công bố', unit: 'triệu USD', fmt: amount },
    { kpi: 'ebit_margin', label: 'Biên EBIT', unit: '%', fmt: fmtPct1 },
    { g: 'Hàng tồn kho' },
    { obs: 'inventories', label: 'Hàng tồn kho cuối kỳ', unit: 'triệu USD', fmt: amount },
    { kpi: 'inventory_growth', label: 'So với cùng thời điểm năm trước', unit: '%', fmt: fmtGrowth },
    { kpi: 'days_inventory', label: 'Số ngày tồn kho bình quân', unit: 'ngày', fmt: function (v) { return num(v, 0); } }
  ];
  function regionTableRows() {
    var rows = [];
    D.regions.forEach(function (r) {
      rows.push({ g: r.vi === r.short ? r.vi : r.vi + ' (' + r.short + ')' });
      rows.push({ obs: r.id + '_revenue', label: 'Doanh thu', unit: 'triệu USD', fmt: amount, total: true });
      rows.push({ obs: r.id + '_direct_revenue', label: 'NIKE Direct', unit: 'triệu USD', fmt: amount });
      rows.push({ obs: r.id + '_wholesale_revenue', label: 'Bán buôn', unit: 'triệu USD', fmt: amount });
      rows.push({ kpi: r.id + '_direct_share', label: 'Tỷ trọng NIKE Direct', unit: '%', fmt: fmtPct1 });
      rows.push({ kpi: r.id + '_growth', label: 'Tăng trưởng doanh thu, theo USD', unit: '%', fmt: fmtGrowth });
      rows.push({ obs: r.id + '_ebit', label: 'EBIT', unit: 'triệu USD', fmt: amount });
      rows.push({ kpi: r.id + '_ebit_margin', label: 'Biên EBIT', unit: '%', fmt: fmtPct1 });
    });
    return rows;
  }
  function tableSpec(kind) {
    if (kind === 'R') return { periods: A(), rows: regionTableRows(), name: 'Khu vực', file: 'khu-vuc' };
    return { periods: kind === 'A' ? A() : Q(), rows: CHANNEL_ROWS.concat(FINANCE_ROWS),
      name: kind === 'A' ? 'Theo năm' : 'Theo quý', file: kind === 'A' ? 'theo-nam' : 'theo-quy' };
  }
  var NEGATIVE = 'Số tiền âm để trong ngoặc đơn, ví dụ (790); tỷ lệ âm có dấu trừ.';
  var TABLE_KEYS = {
    A: 'Số NIKE công bố có kèm số trang của 10-K (bấm để mở báo cáo trên SEC). <i>Số in nghiêng</i> là số nhóm tự tính. ' + NEGATIVE,
    Q: 'Quý 1–3: số trong báo cáo quý (10-Q), có số trang. Quý 4: không có 10-Q, nên nhóm lấy cả năm trừ 9 tháng.<br>' +
      '<i>Số in nghiêng</i> là số nhóm tự tính; ≈ là ước tính từ số NIKE đã làm tròn. Tỷ lệ % quý 4 lấy từ thông cáo kết quả quý 4 (ghi “8-K”).<br>' +
      NEGATIVE,
    R: 'Số NIKE công bố có kèm số trang của 10-K. <i>Số in nghiêng</i> là số nhóm tự tính. EBIT khu vực chưa trừ chi phí Global Brand Divisions và Corporate. ' +
      NEGATIVE
  };
  function buildTable(kind, table) {
    var spec = tableSpec(kind), ps = spec.periods;
    table.replaceChildren();
    var thead = document.createElement('thead');
    if (kind === 'Q') {
      var yr = document.createElement('tr');
      yr.className = 'years';
      yr.append(cell('th'), cell('th', null, 'col-unit'));
      A().forEach(function (y) {
        var th = cell('th', null, 'year'); th.colSpan = 4; th.scope = 'colgroup';
        th.append(cell('span', y, 'year-label'));
        yr.append(th);
      });
      thead.append(yr);
    }
    var hr = document.createElement('tr');
    var h0 = cell('th', 'Chỉ tiêu'), h1 = cell('th', 'Đơn vị', 'col-unit');
    h0.scope = 'col'; h1.scope = 'col';
    hr.append(h0, h1);
    ps.forEach(function (p) {
      var th = cell('th', kind === 'Q' ? 'Q' + quarter(p) : p, 'period');
      th.scope = 'col'; th.title = D.dates[p];
      if (kind === 'Q') th.setAttribute('aria-label', label(p));
      hr.append(th);
    });
    thead.append(hr);
    var tbody = document.createElement('tbody');
    spec.rows.forEach(function (row) {
      var tr = document.createElement('tr');
      if (row.g) {
        tr.className = 'group';
        var th = cell('th', null); th.colSpan = ps.length + 2; th.scope = 'colgroup';
        th.append(cell('span', row.g, 'group-label'));
        tr.append(th); tbody.append(tr); return;
      }
      if (row.total) tr.className = 'total';
      var rh = cell('th', row.label); rh.scope = 'row';
      rh.append(cell('span', ' (' + row.unit + ')', 'unit-inline'));
      tr.append(rh, cell('td', row.unit, 'unit col-unit'));
      ps.forEach(function (p) {
        // Every numeric cell has the same two trailing slots (marker, page reference), so reported and calculated
        // figures share one right edge.
        var td, mark = cell('span', null, 'mark-slot'), ref = cell('span', null, 'ref-slot');
        var q1 = kind === 'Q' && quarter(p) === 1 ? ' q1' : '';
        if (row.kpi) {
          var v = kpi(row.kpi, p);
          td = cell('td', v == null ? '—' : row.fmt(v), 'num calc' + q1);
          td.title = 'Nhóm tự tính: ' + D.kpis[row.kpi].vi + '. ' + D.kpis[row.kpi].note;
        } else {
          var o = ob(row.obs, p);
          if (!o || o.value == null) {
            td = cell('td', null, 'num missing' + q1);
            // Two short lines, so this column is no wider than the columns of figures.
            var miss = cell('span', null, 'missing-text');
            miss.append(cell('span', 'Không'), document.createTextNode(' '), cell('span', 'công bố'));
            td.append(miss);
            td.title = (o && o.cross_check) || 'NIKE không công bố số liệu này.';
          } else if (o.method === 'q4_derived') {
            // An estimate from rounded figures keeps one decimal only.
            td = cell('td', o.approx ? '≈' + num(o.value, 1) : row.fmt(o.value), 'num calc' + q1);
            td.title = o.cross_check + ' Số in trong báo cáo (kiểu Mỹ): ' + o.evidence + '.';
          } else {
            td = cell('td', row.fmt(o.value), 'num' + q1);
            td.title = D.metrics[row.obs].en + '. ' + o.status + '. ' + (o.method === 'q4_release' ? o.cross_check : o.restatement || '');
            if (o.status === STATUS_PENDING) { mark.textContent = '†'; mark.setAttribute('aria-label', 'chờ người kiểm tra'); }
            // A page reference links to the report; a fourth-quarter earnings release has no page numbers.
            var release = D.filings[o.source].form === '8-K';
            if (o.page || release) {
              var a = document.createElement('a');
              a.className = 'page-ref'; a.href = D.filings[o.source].url; a.target = '_blank'; a.rel = 'noopener';
              a.textContent = release ? '8-K' : 'tr. ' + o.page;
              a.setAttribute('aria-label', release ? 'Mở thông cáo kết quả quý 4 (' + o.source + ') trên SEC' :
                'Mở ' + o.source + ' trên SEC, trang ' + o.page);
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
  function renderTable() {
    var table = document.getElementById('data-table');
    buildTable(state.table, table);
    document.getElementById('table-key').innerHTML = TABLE_KEYS[state.table];
    fitTable(table);
  }
  // The page shows one table at a time; a printed copy has all three.
  function renderPrintTables() {
    var box = document.getElementById('print-tables');
    box.replaceChildren();
    ['A', 'Q', 'R'].forEach(function (kind) {
      var table = cell('table', null, 'data-table' + (kind === 'Q' ? ' print-q' : ''));
      buildTable(kind, table);
      var key = cell('p', null, 'table-key');
      key.innerHTML = TABLE_KEYS[kind];
      box.append(cell('h3', 'Bảng số liệu – ' + tableSpec(kind).name.toLowerCase()), key, table);
    });
  }
  // Where the table scrolls sideways (narrow screens, the quarterly table) it opens on the latest periods, beside the
  // pinned label column. The label column takes up the spare width, so the first period in view starts right at its
  // edge instead of showing the tail of the period before it.
  function fitTable(table) {
    var box = table.parentNode, rows = table.tHead.rows, cols = Array.from(rows[rows.length - 1].cells);
    var width = function (c) { return c.getBoundingClientRect().width; };
    box.style.removeProperty('--label-width');
    if (box.scrollWidth > box.clientWidth) {
      var room = box.clientWidth - width(cols[0]), used = 0, i = cols.length;
      while (i > 1 && used + width(cols[i - 1]) <= room + 0.5) used += width(cols[--i]);
      if (used > 0) box.style.setProperty('--label-width', (width(cols[0]) + room - used) + 'px');
    }
    box.style.setProperty('--pin', width(cols[0]) + 'px');
    box.scrollLeft = box.scrollWidth;
    markScrolled(box);
  }
  // While columns slide under the pinned label column, its edge fades them out (styles.css).
  function markScrolled(box) { box.classList.toggle('scrolled', box.scrollLeft > 0); }

  // Table download as .xlsx: unlike a CSV, it opens correctly whatever the viewer's regional settings
  // (list separator and decimal mark differ between Vietnamese and English Windows). Minimal writer:
  // one sheet, inline strings, numbers stored as numbers, ZIP without compression.
  var UNIT_STYLE = { 'triệu USD': 2, '%': 3, 'tỷ USD': 4, 'ngày': 5, 'điểm %': 6 };
  var XLSX_STYLES = '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="5"><numFmt numFmtId="164" formatCode="#,##0;(#,##0);0"/><numFmt numFmtId="165" formatCode="0.0%"/>' +
    '<numFmt numFmtId="166" formatCode="0.0##"/><numFmt numFmtId="167" formatCode="0"/><numFmt numFmtId="168" formatCode="0.0"/></numFmts>' +
    '<fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="11"/><name val="Arial"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="7">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    [164, 165, 166, 167, 168].map(function (f) { return '<xf numFmtId="' + f + '" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'; }).join('') +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  function tableRows(kind) {
    var spec = tableSpec(kind), ps = spec.periods, bold = function (t) { return { t: t, bold: true }; };
    var rows = [[bold('Chỉ tiêu'), bold('Đơn vị'), bold('Loại')].concat(ps.map(function (p) { return bold(label(p)); }), [bold('Nguồn'), bold('Mã')])];
    spec.rows.forEach(function (row) {
      if (row.g) { rows.push([bold(row.g)]); return; }
      var id = row.kpi || row.obs;
      var vals = ps.map(function (p) {
        var v = row.kpi ? kpi(id, p) : val(id, p);
        if (v == null) return row.kpi ? '—' : 'Không công bố';
        // Contributions are shown in percentage points, like the dashboard table.
        if (row.unit === 'điểm %') v = v * 100;
        return { v: Math.round(v * 1e9) / 1e9, s: UNIT_STYLE[row.unit] };
      });
      var source = row.kpi ? 'Nhóm tự tính: ' + D.kpis[id].vi : (kind === 'Q' ? D.metrics[id].location_10q : D.metrics[id].location_10k);
      rows.push([row.label, row.unit, row.kpi ? 'Nhóm tự tính' : 'NIKE công bố'].concat(vals, [source, id]));
    });
    rows.push([]);
    if (kind === 'Q') rows.push(['Không có 10-Q cho quý 4, nên quý 4 = số cả năm (10-K) trừ số 9 tháng đầu (10-Q quý 3); tỷ lệ % quý 4 lấy từ thông cáo kết quả quý 4 (8-K).']);
    rows.push(['Nguồn: NIKE, Inc. Form 10-K, 10-Q và thông cáo kết quả quý 4 (Form 8-K), FY2019–FY2021 (SEC EDGAR). Dữ liệu lấy ngày ' +
      isoDate(D.meta.extracted_on) + '. Dashboard: ' + location.href.split(/[?#]/)[0]]);
    return rows;
  }

  var CRC_TABLE = (function () {
    var t = [];
    for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  function zipHeader(size, fields) {  // fields: [offset, value, byteLength]
    var v = new DataView(new ArrayBuffer(size));
    fields.forEach(function (f) { if (f[2] === 4) v.setUint32(f[0], f[1], true); else v.setUint16(f[0], f[1], true); });
    return v.buffer;
  }
  function zipBlob(files, type) {
    var enc = new TextEncoder(), parts = [], dir = [], offset = 0;
    files.forEach(function (f) {
      var name = enc.encode(f.name), data = enc.encode(f.text), crc = crc32(data), n = data.length;
      // Stored (no compression), UTF-8 names (flag 0x0800), date 1980-01-01.
      parts.push(zipHeader(30, [[0, 0x04034b50, 4], [4, 20, 2], [6, 0x0800, 2], [12, 0x21, 2], [14, crc, 4], [18, n, 4], [22, n, 4],
        [26, name.length, 2]]), name, data);
      dir.push(zipHeader(46, [[0, 0x02014b50, 4], [4, 20, 2], [6, 20, 2], [8, 0x0800, 2], [14, 0x21, 2], [16, crc, 4], [20, n, 4], [24, n, 4],
        [28, name.length, 2], [42, offset, 4]]), name);
      offset += 30 + name.length + n;
    });
    var size = dir.reduce(function (s, b) { return s + b.byteLength; }, 0);
    parts = parts.concat(dir, [zipHeader(22, [[0, 0x06054b50, 4], [8, files.length, 2], [10, files.length, 2], [12, size, 4], [16, offset, 4]])]);
    return new Blob(parts, { type: type });
  }
  function colName(j) {
    var s = '';
    for (j += 1; j > 0; j = Math.floor((j - 1) / 26)) s = String.fromCharCode(65 + (j - 1) % 26) + s;
    return s;
  }
  function xlsxBlob(sheetName, rows, widths) {
    var x = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
    var body = rows.map(function (r, i) {
      return '<row r="' + (i + 1) + '">' + r.map(function (c, j) {
        var ref = colName(j) + (i + 1);
        if (c == null || c === '') return '';
        if (typeof c === 'object' && c.v != null) return '<c r="' + ref + '" s="' + c.s + '"><v>' + c.v + '</v></c>';
        var text = typeof c === 'object' ? c.t : c, style = typeof c === 'object' && c.bold ? ' s="1"' : '';
        return '<c r="' + ref + '" t="inlineStr"' + style + '><is><t xml:space="preserve">' + x(text) + '</t></is></c>';
      }).join('') + '</row>';
    }).join('');
    var cols = widths.map(function (w, j) { return '<col min="' + (j + 1) + '" max="' + (j + 1) + '" width="' + w + '" customWidth="1"/>'; }).join('');
    var head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n', ns = 'http://schemas.openxmlformats.org/';
    var rel = ns + 'officeDocument/2006/relationships', pkg = ns + 'package/2006/', sml = ns + 'spreadsheetml/2006/main';
    var ct = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';
    return zipBlob([
      { name: '[Content_Types].xml', text: head + '<Types xmlns="' + pkg + 'content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="' + ct + 'sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="' + ct + 'worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="' + ct + 'styles+xml"/></Types>' },
      { name: '_rels/.rels', text: head + '<Relationships xmlns="' + pkg + 'relationships"><Relationship Id="rId1" Type="' + rel + '/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: 'xl/workbook.xml', text: head + '<workbook xmlns="' + sml + '" xmlns:r="' + rel + '"><sheets><sheet name="' + x(sheetName) + '" sheetId="1" r:id="rId1"/></sheets></workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', text: head + '<Relationships xmlns="' + pkg + 'relationships"><Relationship Id="rId1" Type="' + rel + '/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="' + rel + '/styles" Target="styles.xml"/></Relationships>' },
      { name: 'xl/styles.xml', text: head + XLSX_STYLES },
      { name: 'xl/worksheets/sheet1.xml', text: head + '<worksheet xmlns="' + sml + '"><cols>' + cols + '</cols><sheetData>' + body + '</sheetData></worksheet>' }
    ], ct + 'sheet');
  }
  function download(blob, name) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.append(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }

  // PNG export: the chart image with its title, unit line, legend and source baked in, ready for slides.
  function exportPng(id) {
    var chart = charts[id];
    if (!chart) return;
    var fig = document.querySelector('.chart[data-chart="' + id + '"]').closest('figure');
    var heading = fig.querySelector('h3').textContent;
    var sub = fig.querySelector('.figure-sub').textContent;
    var legendItems = Array.prototype.map.call(fig.querySelectorAll('.legend .legend-item'), function (n) {
      var sw = n.querySelector('.swatch');
      return { text: n.textContent, color: sw.style.boxShadow ? '#ffffff' : sw.style.background, ring: sw.style.boxShadow };
    });
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
      var tL = lines(heading, fTitle, img.width), sL = lines(sub, fSub, img.width);
      ctx.font = fSub;
      var legendRows = [], rowItems = [], rowW = 0;
      legendItems.forEach(function (it) {
        var w = 18 * scale + ctx.measureText(it.text).width + 20 * scale;
        if (rowItems.length && rowW + w > img.width) { legendRows.push(rowItems); rowItems = []; rowW = 0; }
        rowItems.push(it); rowW += w;
      });
      if (rowItems.length) legendRows.push(rowItems);
      var srcL = lines(src, fSrc, img.width);
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
      var tag = id === 'change' ? '_' + state.change : id === 'bridge' ? '_' + state.bridge : '';
      canvas.toBlob(function (blob) { download(blob, 'NIKE_' + id + tag + '.png'); announce('Đã tạo ảnh PNG: ' + heading + '.'); }, 'image/png');
    };
    img.src = chart.getDataURL({ type: 'png', pixelRatio: scale, backgroundColor: '#ffffff' });
  }

  // ------------------------------------------------------------------ footer
  function renderFooter() {
    var table = document.getElementById('filing-list');
    table.querySelectorAll('tbody').forEach(function (b) { b.remove(); });
    var groups = [['Báo cáo năm (10-K)', function (f) { return f.form === '10-K' && f.role === 'scope'; }],
      ['Báo cáo quý (10-Q)', function (f) { return f.form === '10-Q' && f.role === 'scope'; }],
      ['Thông cáo kết quả quý 4 (8-K)', function (f) { return f.role === 'q4release'; }],
      ['Báo cáo các năm sau (chỉ để kiểm tra NIKE có in lại số cũ khác đi không)', function (f) { return f.role === 'check'; }]];
    groups.forEach(function (g) {
      var tbody = document.createElement('tbody'), gr = document.createElement('tr'), th = cell('th', g[0]);
      th.colSpan = 3; th.scope = 'colgroup';
      gr.append(th); tbody.append(gr);
      Object.keys(D.filings).filter(function (k) { return g[1](D.filings[k]); }).forEach(function (k) {
        var f = D.filings[k], a = document.createElement('a');
        a.href = f.url; a.target = '_blank'; a.rel = 'noopener'; a.textContent = k;
        a.setAttribute('aria-label', 'Mở ' + f.title + ' trên SEC EDGAR (thẻ mới)');
        var tr = document.createElement('tr'), link = cell('td');
        link.append(a);
        tr.append(link, cell('td', isoDate(f.filed)), cell('td', f.accession));
        tbody.append(tr);
      });
      table.append(tbody);
    });
    var c = D.meta.status_counts;
    var total = Object.keys(c).reduce(function (s, k) { return s + c[k]; }, 0);
    // Missing figures are described by what they are when they are all of one kind (the fourth-quarter ratios).
    var missing = [];
    Object.keys(D.observations).forEach(function (id) {
      Object.keys(D.observations[id]).forEach(function (p) { if (D.observations[id][p].value == null) missing.push(p); });
    });
    var allQ4 = missing.length && missing.every(function (p) { return /Q4$/.test(p); });
    // Only statuses that occur are listed.
    var parts = [[c['Khớp nguồn chính thức'], ' số đã được kiểm tra lại và khớp'],
      [c['Đã xác minh (người)'], ' số đã được nhóm kiểm tra trên báo cáo'], [c[STATUS_PENDING], ' số chưa được kiểm tra (†)'],
      [c['Thiếu / không công bố'], allQ4 ? ' ô là tỷ lệ % NIKE không nêu cho riêng quý 4' : ' ô NIKE không công bố']]
      .filter(function (p) { return p[0]; }).map(function (p) { return p[0] + p[1]; });
    document.getElementById('status-summary').textContent =
      total + ' số liệu: ' + parts.join(', ') + '. Dữ liệu lấy ngày ' + isoDate(D.meta.extracted_on) + '.';
    document.getElementById('evidence-link').href = D.meta.evidence;
  }

  // ------------------------------------------------------------------ controls
  var liveTimer = null;
  function announce(msg) {
    var el = document.getElementById('live');
    el.textContent = '';
    clearTimeout(liveTimer);
    liveTimer = setTimeout(function () { el.textContent = msg; }, 50);
  }
  // Kept on one line because they read badly when split: a report name with its period ("Form 10-Q", "10-Q quý 3
  // FY2020", "10-Q Q3 FY2020", "8-K"), a quarter or months with their numbers ("quý 4 FY2020", "quý 1–3", "tháng
  // 3–5/2020"), "non-GAAP" and "COVID-19" (browsers may break right after a hyphen), a period range, a figure and its
  // unit (USD, điểm %, ngày), "tỷ lệ %", "(xấp xỉ)", "(ước tính)" and a short name in quotation marks.
  var KEEP = new RegExp('((?:Form )?10-[KQ](?: quý [1-4])?(?: Q[1-4])?(?: FY\\d{4})?|(?:Form )?8-K(?: Q4 FY\\d{4})?|' +
    'quý [1-4](?:–[1-4])?(?: FY\\d{4})?|Q[1-4] FY\\d{4}|tháng \\d{1,2}(?:–\\d{1,2})?(?:/\\d{4})?|non-GAAP|COVID-19|' +
    'FY\\d{4}(?:–| → )FY\\d{4}|[±+−]?\\d[\\d.,]*(?:–\\d[\\d.,]*)? (?:tỷ|triệu) USD|(?:[+−]?\\d[\\d.,]* )?điểm %|tỷ lệ %|' +
    '\\d[\\d.,]* ngày|\\((?:xấp xỉ|ước tính)\\)|“[^”]{1,24}”)');
  function keepTogether(root) {
    var walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT), found = [], n;
    while ((n = walker.nextNode())) {
      if (KEEP.test(n.nodeValue) && !n.parentNode.closest('.nowrap, script, style, option, .chart, .visually-hidden, .finding-value'))
        found.push(n);
    }
    found.forEach(function (node) {
      var frag = document.createDocumentFragment();
      node.nodeValue.split(KEEP).forEach(function (s, i) {
        if (s) frag.append(i % 2 ? cell('span', s, 'nowrap') : document.createTextNode(s));
      });
      node.parentNode.replaceChild(frag, node);
    });
  }
  function watchSections() {
    if (!window.IntersectionObserver) return;
    var links = Array.from(document.querySelectorAll('.section-nav a'));
    var seen = {};
    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (e) { seen[e.target.id] = e.isIntersecting; });
      var current = links.map(function (a) { return a.getAttribute('href').slice(1); }).filter(function (id) { return seen[id]; })[0];
      links.forEach(function (a) {
        if (a.getAttribute('href') === '#' + current) a.setAttribute('aria-current', 'true'); else a.removeAttribute('aria-current');
      });
    }, { rootMargin: '-56px 0px -55% 0px' });
    links.forEach(function (a) { var s = document.getElementById(a.getAttribute('href').slice(1)); if (s) io.observe(s); });
  }
  function bind() {
    document.querySelectorAll('[data-toggle] input').forEach(function (r) {
      r.addEventListener('change', function () {
        var which = r.closest('[data-toggle]').dataset.toggle;
        state[which] = r.value;
        if (which === 'table') { renderTable(); announce('Đang hiển thị bảng ' + tableSpec(r.value).name.toLowerCase() + '.'); }
        else { RENDER[which](); announce('Đang so sánh ' + r.value.replace('-', ' với ') + '.'); }
        keepTogether(document.body);
      });
    });
    document.querySelector('.table-scroll').addEventListener('scroll', function (e) { markScrolled(e.target); }, { passive: true });
    var resizeTimer = null;
    window.addEventListener('resize', function () {
      // Label density depends on chart width, so redraw (not just rescale) after resizing; refit the table too.
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(function () {
        Object.keys(charts).forEach(function (id) { charts[id].resize(); });
        renderCharts();
        keepTogether(document.body);
        fitTable(document.getElementById('data-table'));
      }, 200);
    });
    document.querySelectorAll('.export-png').forEach(function (b) {
      b.addEventListener('click', function () { exportPng(b.dataset.chart); });
    });
    document.getElementById('xlsx-table').addEventListener('click', function () {
      var spec = tableSpec(state.table);
      var widths = [46, 11, 14].concat(spec.periods.map(function () { return 11; }), [70, 30]);
      download(xlsxBlob(spec.name, tableRows(state.table), widths), 'NIKE_bang-' + spec.file + '_FY2019-FY2021.xlsx');
      announce('Đã tạo file Excel.');
    });
    // Charts are drawn at screen width; printing resizes them to the page first.
    window.addEventListener('beforeprint', function () { Object.keys(charts).forEach(function (id) { charts[id].resize(); }); });
    window.addEventListener('afterprint', function () { Object.keys(charts).forEach(function (id) { charts[id].resize(); }); });
  }

  function fail(msg) {
    var el = document.getElementById('load-error');
    el.textContent = msg;
    el.hidden = false;
  }

  function start() {
    if (!window.echarts) { fail('Không tải được thư viện biểu đồ (assets/vendor/echarts-5.6.0.min.js). Hãy kiểm tra xem thư mục site đã được đưa lên đầy đủ chưa.'); return; }
    fetch('data/nike-financials.json', { cache: 'no-cache' })
      .then(function (r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
      .then(function (data) {
        D = data;
        return document.fonts ? document.fonts.ready : null;
      })
      .then(function () {
        bind();
        renderFooter();
        renderFindings();
        renderQuotes();
        renderCharts();
        renderTable();
        renderPrintTables();
        keepTogether(document.body);
        watchSections();
        document.documentElement.dataset.ready = 'true';
      })
      .catch(function (e) {
        fail('Không tải được dữ liệu (data/nike-financials.json): ' + e.message +
          '. Nếu bạn mở file index.html trực tiếp trên máy, hãy chạy máy chủ web cục bộ: python3 -m http.server --directory site');
      });
  }

  document.addEventListener('DOMContentLoaded', start);
})();
