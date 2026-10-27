/* Sheets charts: our own canvas renderer for column, bar, line, area, pie, doughnut and scatter charts.
   A chart reads a range live (first row = series names, first column = categories when they are text). */
(function () {
  "use strict";
  var SH = window.SH, C = SH.core, F = SH.fmt;
  var PALETTE = ["#2f7fd8", "#e8743b", "#19a979", "#945ecf", "#ed4a7b", "#13a4b4", "#bf8b00", "#6c8893", "#d24d57", "#5a9a3a"];
  var TYPES = [["column", "Column"], ["bar", "Bar"], ["line", "Line"], ["area", "Area"], ["pie", "Pie"], ["doughnut", "Doughnut"], ["scatter", "Scatter"]];

  function isNum(v) { return typeof v === "number" && isFinite(v); }
  // read the chart's range into categories + series
  function data(wb, ch) {
    var g = C.parseRange(ch.range);
    if (!g) return { error: "The data range is not valid." };
    var sh = g.sheet ? wb.sheetByName(g.sheet) : wb.sheet();
    if (!sh) return { error: "The sheet in the data range was not found." };
    var r1 = Math.min(g.r1, Math.max(sh.maxR, sh.smaxR || 0)), c1 = Math.min(g.c1, Math.max(sh.maxC, sh.smaxC || 0));
    if (r1 < g.r0 || c1 < g.c0) return { error: "The data range is empty." };
    if ((r1 - g.r0 + 1) * (c1 - g.c0 + 1) > 200000) return { error: "That range is too large to chart (200,000 cells max)." };
    var V = [], nf = null;
    for (var r = g.r0; r <= r1; r++) {
      if (sh.hidden && sh.hidden.has(r)) continue;
      var row = [];
      for (var c = g.c0; c <= c1; c++) { var v = wb.value(sh, r, c); row.push(v instanceof C.Err ? null : v); var s = wb.style(sh, r, c); if (!nf && isNum(v) && s && s.nf) nf = s.nf; }
      V.push(row);
    }
    var W = V[0].length, header = V.length > 1 && V[0].some(function (v, j) { return (j > 0 || W === 1) && v !== null && !isNum(v); });
    var body = header ? V.slice(1) : V;
    var catCol = W > 1 && (ch.type === "scatter" ? true : body.some(function (row) { return row[0] !== null && !isNum(row[0]); }) || F.isDateFmt(wb.style(sh, g.r0 + (header ? 1 : 0), g.c0) && wb.style(sh, g.r0 + (header ? 1 : 0), g.c0).nf));
    var dateCats = catCol && F.isDateFmt((wb.style(sh, g.r0 + (header ? 1 : 0), g.c0) || {}).nf);
    var cats = body.map(function (row, i) {
      if (!catCol) return String(i + 1);
      var v = row[0];
      if (dateCats && isNum(v)) return F.format(v, (wb.style(sh, g.r0 + (header ? 1 : 0), g.c0) || {}).nf).t;
      return v === null ? "" : F.format(v).t;
    });
    var series = [];
    for (var j = catCol ? 1 : 0; j < W; j++) {
      series.push({ name: header && V[0][j] !== null ? F.format(V[0][j]).t : "Series " + (series.length + 1), vals: body.map(function (row) { return isNum(row[j]) ? row[j] : null; }), color: PALETTE[series.length % PALETTE.length] });
    }
    if (ch.type === "scatter") return { xs: body.map(function (row) { return isNum(row[0]) ? row[0] : null; }), series: series, nf: nf };
    if (!series.length) return { error: "Add at least one column of numbers to the range." };
    return { cats: cats, series: series, nf: nf };
  }
  function nice(lo, hi, ticks) {
    if (lo === hi) { if (lo === 0) hi = 1; else if (lo > 0) lo = 0; else hi = 0; }
    var span = hi - lo, raw = span / Math.max(1, ticks), mag = Math.pow(10, Math.floor(Math.log10(raw))), norm = raw / mag;
    var step = (norm < 1.5 ? 1 : norm < 3 ? 2 : norm < 7 ? 5 : 10) * mag;
    return { min: Math.floor(lo / step) * step, max: Math.ceil(hi / step) * step, step: step };
  }
  function tickText(v, nf, step) {
    if (nf && !F.isDateFmt(nf)) { var t = F.format(v, nf.replace(/\.0+/, step < 1 ? ".00" : "")).t; if (t.length < 14) return t; }
    var a = Math.abs(v);
    if (a >= 1e9) return C.r15(v / 1e9) + "B";
    if (a >= 1e6) return C.r15(v / 1e6) + "M";
    if (a >= 1e4) return C.r15(v / 1e3) + "K";
    return F.format(C.r15(v), step < 1 ? "#,##0.0#" : "#,##0").t;
  }
  function ellipsis(g, t, max) {
    if (g.measureText(t).width <= max) return t;
    while (t.length > 1 && g.measureText(t + "…").width > max) t = t.slice(0, -1);
    return t + "…";
  }

  // theme: { bg, fg, muted, grid }
  function draw(cv, wb, ch, theme, wCss, hCss) {
    var dpr = window.devicePixelRatio || 1, W = wCss, H = hCss;
    cv.width = Math.round(W * dpr); cv.height = Math.round(H * dpr);
    var g = cv.getContext("2d");
    g.setTransform(dpr, 0, 0, dpr, 0, 0);
    g.fillStyle = theme.bg; g.fillRect(0, 0, W, H);
    var font = "12px " + theme.font, d = data(wb, ch), top = 12;
    g.textBaseline = "middle";
    if (ch.title) { g.fillStyle = theme.fg; g.font = "600 14px " + theme.font; g.textAlign = "center"; g.fillText(ellipsis(g, ch.title, W - 20), W / 2, 18); top = 34; }
    if (d.error) { g.fillStyle = theme.muted; g.font = font; g.textAlign = "center"; g.fillText(d.error, W / 2, H / 2); return d; }
    var pie = ch.type === "pie" || ch.type === "doughnut";
    // legend
    var items = pie ? d.cats.map(function (c, i) { return { name: c || "(blank)", color: PALETTE[i % PALETTE.length] }; }) : d.series;
    var legendH = 0;
    g.font = font;
    if (ch.legend !== false && (items.length > 1 || pie)) {
      var x = 0, lines = 1, maxW = W - 24;
      items.forEach(function (it) { var w = Math.min(160, g.measureText(it.name).width) + 26; if (x + w > maxW) { lines++; x = 0; } x += w; });
      lines = Math.min(lines, 3); legendH = lines * 18 + 8;
      var lx = 12, ly = H - legendH + 6, line = 0;
      items.forEach(function (it) {
        var w = Math.min(160, g.measureText(it.name).width) + 26;
        if (lx + w > W - 12) { line++; lx = 12; ly += 18; }
        if (line >= 3) return;
        g.fillStyle = it.color; g.fillRect(lx, ly + 4, 10, 10);
        g.fillStyle = theme.fg; g.textAlign = "left"; g.fillText(ellipsis(g, it.name, 150), lx + 14, ly + 9);
        lx += w;
      });
    }
    var plot = { x: 12, y: top, w: W - 24, h: H - top - legendH - 8 };
    if (pie) return drawPie(g, d, ch, theme, plot), d;
    if (ch.type === "scatter") return drawScatter(g, d, theme, plot, font), d;
    drawAxes(g, d, ch, theme, plot, font);
    return d;
  }
  function range(series) {
    var lo = Infinity, hi = -Infinity;
    series.forEach(function (s) { s.vals.forEach(function (v) { if (v !== null) { if (v < lo) lo = v; if (v > hi) hi = v; } }); });
    if (lo === Infinity) { lo = 0; hi = 1; }
    return [Math.min(0, lo), Math.max(0, hi)];
  }
  function drawAxes(g, d, ch, theme, P, font) {
    var horiz = ch.type === "bar", rg = range(d.series), sc = nice(rg[0], rg[1], horiz ? Math.max(2, Math.floor(P.w / 90)) : Math.max(2, Math.floor(P.h / 45)));
    g.font = font;
    var labW = 0, k, tv;
    for (tv = sc.min; tv <= sc.max + sc.step / 2; tv += sc.step) labW = Math.max(labW, g.measureText(tickText(tv, d.nf, sc.step)).width);
    var catW = 0; d.cats.forEach(function (c) { catW = Math.max(catW, g.measureText(c).width); });
    var left = horiz ? Math.min(P.w * 0.35, catW + 10) : labW + 10, bottom = horiz ? 22 : 24;
    var A = { x: P.x + left, y: P.y + 4, w: P.w - left - 4, h: P.h - bottom - 4 };
    if (A.w < 20 || A.h < 20) return;
    function vpos(v) { var t = (v - sc.min) / (sc.max - sc.min); return horiz ? A.x + t * A.w : A.y + A.h - t * A.h; }
    // grid + value labels
    g.strokeStyle = theme.grid; g.lineWidth = 1; g.fillStyle = theme.muted;
    for (tv = sc.min, k = 0; tv <= sc.max + sc.step / 2 && k < 60; tv += sc.step, k++) {
      var p = Math.round(vpos(tv)) + 0.5;
      g.beginPath();
      if (horiz) { g.moveTo(p, A.y); g.lineTo(p, A.y + A.h); g.textAlign = "center"; g.fillText(tickText(tv, d.nf, sc.step), p, A.y + A.h + 12); }
      else { g.moveTo(A.x, p); g.lineTo(A.x + A.w, p); g.textAlign = "right"; g.fillText(tickText(tv, d.nf, sc.step), A.x - 6, p); }
      g.stroke();
    }
    var n = d.cats.length, band = (horiz ? A.h : A.w) / Math.max(1, n), zero = vpos(Math.max(sc.min, Math.min(0, sc.max)));
    // category labels (skip some if they would overlap)
    var every = Math.max(1, Math.ceil((horiz ? 16 : catW + 10) / Math.max(1, band)));
    g.fillStyle = theme.muted;
    d.cats.forEach(function (c, i) {
      if (i % every) return;
      var mid = (horiz ? A.y : A.x) + band * (i + 0.5);
      if (horiz) { g.textAlign = "right"; g.fillText(ellipsis(g, c, left - 8), A.x - 6, mid); }
      else { g.textAlign = "center"; g.fillText(ellipsis(g, c, Math.max(band * every - 4, 30)), mid, A.y + A.h + 13); }
    });
    var S = d.series.length;
    if (ch.type === "column" || ch.type === "bar") {
      var inner = band * 0.78, bw = inner / S;
      d.series.forEach(function (s, si) {
        g.fillStyle = s.color;
        s.vals.forEach(function (v, i) {
          if (v === null) return;
          var o = (horiz ? A.y : A.x) + band * i + (band - inner) / 2 + bw * si, p = vpos(v);
          if (horiz) g.fillRect(Math.min(p, zero), o + 0.5, Math.max(1, Math.abs(p - zero)), Math.max(1, bw - 1));
          else g.fillRect(o + 0.5, Math.min(p, zero), Math.max(1, bw - 1), Math.max(1, Math.abs(p - zero)));
        });
      });
    } else {
      d.series.forEach(function (s) {
        var pts = s.vals.map(function (v, i) { return v === null ? null : [A.x + band * (i + 0.5), vpos(v)]; });
        if (ch.type === "area") {
          g.globalAlpha = 0.22; g.fillStyle = s.color; g.beginPath(); var started = false, lastX = 0, firstX = 0;
          pts.forEach(function (p) { if (!p) return; if (!started) { g.moveTo(p[0], zero); g.lineTo(p[0], p[1]); firstX = p[0]; started = true; } else g.lineTo(p[0], p[1]); lastX = p[0]; });
          if (started) { g.lineTo(lastX, zero); g.lineTo(firstX, zero); g.closePath(); g.fill(); }
          g.globalAlpha = 1;
        }
        g.strokeStyle = s.color; g.lineWidth = 2.2; g.lineJoin = "round"; g.beginPath();
        var pen = false;
        pts.forEach(function (p) { if (!p) { pen = false; return; } if (!pen) { g.moveTo(p[0], p[1]); pen = true; } else g.lineTo(p[0], p[1]); });
        g.stroke();
        if (n <= 40) { g.fillStyle = s.color; pts.forEach(function (p) { if (p) { g.beginPath(); g.arc(p[0], p[1], 2.8, 0, Math.PI * 2); g.fill(); } }); }
      });
    }
    g.strokeStyle = theme.axis; g.lineWidth = 1; g.beginPath();
    if (horiz) { g.moveTo(Math.round(zero) + 0.5, A.y); g.lineTo(Math.round(zero) + 0.5, A.y + A.h); } else { g.moveTo(A.x, Math.round(zero) + 0.5); g.lineTo(A.x + A.w, Math.round(zero) + 0.5); }
    g.stroke();
  }
  function drawScatter(g, d, theme, P, font) {
    var xs = d.xs.filter(function (v) { return v !== null; });
    if (!xs.length || !d.series.length) { g.fillStyle = theme.muted; g.textAlign = "center"; g.fillText("Scatter needs numbers in the first column (X) and at least one more column (Y).", P.x + P.w / 2, P.y + P.h / 2); return; }
    var rx = nice(Math.min.apply(null, xs), Math.max.apply(null, xs), Math.max(2, Math.floor(P.w / 90))), ry = range(d.series);
    var all = []; d.series.forEach(function (s) { s.vals.forEach(function (v) { if (v !== null) all.push(v); }); });
    ry = nice(Math.min.apply(null, all.length ? all : [0]), Math.max.apply(null, all.length ? all : [1]), Math.max(2, Math.floor(P.h / 45)));
    g.font = font;
    var lw = 0; for (var t = ry.min; t <= ry.max + ry.step / 2; t += ry.step) lw = Math.max(lw, g.measureText(tickText(t, d.nf, ry.step)).width);
    var A = { x: P.x + lw + 10, y: P.y + 4, w: P.w - lw - 14, h: P.h - 28 };
    function X(v) { return A.x + (v - rx.min) / (rx.max - rx.min) * A.w; }
    function Y(v) { return A.y + A.h - (v - ry.min) / (ry.max - ry.min) * A.h; }
    g.strokeStyle = theme.grid; g.fillStyle = theme.muted; g.lineWidth = 1;
    var k = 0;
    for (t = ry.min; t <= ry.max + ry.step / 2 && k++ < 60; t += ry.step) { var p = Math.round(Y(t)) + 0.5; g.beginPath(); g.moveTo(A.x, p); g.lineTo(A.x + A.w, p); g.stroke(); g.textAlign = "right"; g.fillText(tickText(t, d.nf, ry.step), A.x - 6, p); }
    k = 0;
    for (t = rx.min; t <= rx.max + rx.step / 2 && k++ < 60; t += rx.step) { var q = Math.round(X(t)) + 0.5; g.beginPath(); g.moveTo(q, A.y); g.lineTo(q, A.y + A.h); g.stroke(); g.textAlign = "center"; g.fillText(tickText(t, null, rx.step), q, A.y + A.h + 13); }
    d.series.forEach(function (s) {
      g.fillStyle = s.color;
      s.vals.forEach(function (v, i) { var x = d.xs[i]; if (v === null || x === null) return; g.beginPath(); g.arc(X(x), Y(v), 3.6, 0, Math.PI * 2); g.fill(); });
    });
  }
  function drawPie(g, d, ch, theme, P) {
    var vals = d.series[0].vals.map(function (v) { return v !== null && v > 0 ? v : 0; }), tot = vals.reduce(function (a, b) { return a + b; }, 0);
    var cx = P.x + P.w / 2, cy = P.y + P.h / 2, R = Math.max(10, Math.min(P.w, P.h) / 2 - 6), a = -Math.PI / 2;
    if (!tot) { g.fillStyle = theme.muted; g.textAlign = "center"; g.fillText("No positive values to show.", cx, cy); return; }
    vals.forEach(function (v, i) {
      if (!v) return;
      var sweep = v / tot * Math.PI * 2;
      g.fillStyle = PALETTE[i % PALETTE.length]; g.beginPath(); g.moveTo(cx, cy); g.arc(cx, cy, R, a, a + sweep); g.closePath(); g.fill();
      g.strokeStyle = theme.bg; g.lineWidth = 1.5; g.stroke();
      if (sweep > 0.28) {
        var mid = a + sweep / 2, rr = ch.type === "doughnut" ? R * 0.78 : R * 0.64;
        g.fillStyle = "#ffffff"; g.font = "600 12px " + theme.font; g.textAlign = "center";
        g.fillText(Math.round(v / tot * 100) + "%", cx + Math.cos(mid) * rr, cy + Math.sin(mid) * rr);
      }
      a += sweep;
    });
    if (ch.type === "doughnut") { g.fillStyle = theme.bg; g.beginPath(); g.arc(cx, cy, R * 0.55, 0, Math.PI * 2); g.fill(); }
  }
  SH.chart = { data: data, draw: draw, TYPES: TYPES, PALETTE: PALETTE, nice: nice };
})();
