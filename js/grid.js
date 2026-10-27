/* Sheets grid: a canvas that draws only the visible cells (fast at 100k rows), frozen panes, merges,
   right-to-left sheets, selection, fill handle, resizing, in-cell editor with formula autocomplete,
   click-to-insert references, clipboard (TSV in and out) and floating charts. */
(function () {
  "use strict";
  var SH = window.SH, C = SH.core, F = SH.fmt, EN = SH.engine, KW = C.KW, MAXR = C.MAXR, MAXC = C.MAXC;
  var HW = 48, HH = 24, DW = 100, DH = 24, FPX = 13;
  var THEMES = {
    paper: { bg: "#ffffff", grid: "#e1e4e8", text: "#1d2228", hdrBg: "#f4f5f7", hdrLine: "#d5d9de", hdrText: "#5f6875", hdrSel: "#d3ecea", hdrSelText: "#0a6763", sel: "#0e9f9a", selFill: "rgba(14,159,154,0.10)", frozen: "#aab2bc", muted: "#8a929e", axis: "#9aa3ae", chartGrid: "#eceef1" },
    dark: { bg: "#15181c", grid: "#2b3138", text: "#e6e9ee", hdrBg: "#1c2025", hdrLine: "#2f353c", hdrText: "#9aa3ae", hdrSel: "#1d3a39", hdrSelText: "#5ee6e0", sel: "#3dd6d0", selFill: "rgba(61,214,208,0.14)", frozen: "#56606b", muted: "#7d8693", axis: "#59626d", chartGrid: "#262b31" }
  };
  var REF_COLORS = ["#2f7fd8", "#d9480f", "#2b8a3e", "#9c36b5", "#c2255c", "#0b7285", "#e67700"];
  var FONT = "Inter, ui-sans-serif, system-ui, -apple-system, 'Segoe UI', Roboto, 'Noto Sans Arabic', Tahoma, sans-serif";

  function lum(hex) { var n = parseInt(hex.slice(1), 16), r = n >> 16 & 255, g = n >> 8 & 255, b = n & 255; function f(x) { x /= 255; return x <= 0.03928 ? x / 12.92 : Math.pow((x + 0.055) / 1.055, 2.4); } return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b); }
  function contrast(a, b) { var x = lum(a), y = lum(b); return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05); }
  function mix(a, b, t) { var p = parseInt(a.slice(1), 16), q = parseInt(b.slice(1), 16); function ch(s) { return Math.round(((p >> s) & 255) * (1 - t) + ((q >> s) & 255) * t); } return "#" + ((1 << 24) + (ch(16) << 16) + (ch(8) << 8) + ch(0)).toString(16).slice(1); }

  function Grid(host, wb, cb) {
    var self = this;
    this.host = host; this.cb = cb || {}; this.sel = { r: 0, c: 0, r0: 0, c0: 0, r1: 0, c1: 0 }; this.edit = null; this.clip = null;
    this.dark = false; this.widths = new Map(); this.raf = 0; this.axKey = ""; this.chartBoxes = new Map(); this.refs = [];
    this.scroller = S3.h("div", { "class": "sh-scroll", tabindex: "0", role: "grid", "aria-label": "Spreadsheet grid. Arrow keys move, type to edit, F2 edits the cell." });
    this.sticky = S3.h("div", { "class": "sh-sticky" });
    this.canvas = S3.h("canvas", { "class": "sh-canvas", "aria-hidden": "true" });
    this.overlay = S3.h("div", { "class": "sh-overlay" });
    this.chartLayer = S3.h("div", { "class": "sh-charts" });
    this.editor = S3.h("textarea", { "class": "sh-editor", spellcheck: "false", autocomplete: "off", dir: "auto", "aria-label": "Cell editor", rows: "1" });
    this.ac = S3.h("div", { "class": "sh-ac", role: "listbox", hidden: true });
    this.tip = S3.h("div", { "class": "sh-tip", hidden: true });
    this.live = S3.h("div", { "class": "sr", "aria-live": "polite" });
    this.spacer = S3.h("div", { "class": "sh-spacer" });
    this.overlay.appendChild(this.chartLayer); this.overlay.appendChild(this.editor); this.overlay.appendChild(this.ac); this.overlay.appendChild(this.tip);
    this.sticky.appendChild(this.canvas); this.sticky.appendChild(this.overlay);
    this.scroller.appendChild(this.sticky); this.scroller.appendChild(this.spacer);
    host.appendChild(this.scroller); host.appendChild(this.live);
    this.ctx = this.canvas.getContext("2d");
    this.setWorkbook(wb);
    this.scroller.addEventListener("scroll", function () { self.onScroll(); });
    if (window.ResizeObserver) new ResizeObserver(function () { self.draw(); }).observe(this.scroller); else window.addEventListener("resize", function () { self.draw(); });
    this.wire();
  }
  var G = Grid.prototype;
  G.setWorkbook = function (wb) {
    var self = this;
    this.wb = wb; this.axKey = ""; this.clip = null; this.edit = null;
    this.hideEditor();
    if (!wb.__grid) { wb.__grid = true; wb.on(function () { self.axKey = ""; self.cfCache = null; self.draw(); }); }
    this.select(0, 0, 0, 0, 0, 0);
    this.scroller.scrollTop = 0; this.scroller.scrollLeft = 0;
    this.draw();
  };
  G.sheet = function () { return this.wb.sheet(); };
  G.theme = function () { return this.dark ? THEMES.dark : THEMES.paper; };

  // ---------- geometry ----------
  G.axes = function () {
    var sh = this.sheet(), key = sh.uid + ":" + sh.nrows + ":" + sh.ncols + ":" + (sh.hiddenVer || 0) + ":" + this.sizeVer;
    if (key === this.axKey && this.rowPos) return;
    this.axKey = key;
    var nr = sh.nrows, nc = sh.ncols, rp = new Float64Array(nr + 1), cp = new Float64Array(nc + 1), i;
    var hasH = sh.rowH.size > 0, hid = sh.hidden;
    for (i = 0; i < nr; i++) rp[i + 1] = rp[i] + (hid.size && hid.has(i) ? 0 : hasH ? (sh.rowH.get(i) || DH) : DH);
    for (i = 0; i < nc; i++) cp[i + 1] = cp[i] + (sh.colW.get(i) || DW);
    this.rowPos = rp; this.colPos = cp;
  };
  G.sizeVer = 0;
  function upper(arr, n, pos) { var lo = 0, hi = n; while (lo < hi) { var m = (lo + hi) >> 1; if (arr[m] > pos) hi = m; else lo = m + 1; } return Math.max(0, Math.min(n - 1, lo - 1)); }
  G.rowAt = function (pos) { return upper(this.rowPos, this.sheet().nrows + 1, pos); };
  G.colAt = function (pos) { return upper(this.colPos, this.sheet().ncols + 1, pos); };
  G.view = function () {
    var sh = this.sheet(), fr = Math.min(sh.freeze.r, sh.nrows - 1), fc = Math.min(sh.freeze.c, sh.ncols - 1);
    return { W: this.scroller.clientWidth, H: this.scroller.clientHeight, sx: Math.abs(this.scroller.scrollLeft), sy: this.scroller.scrollTop, fr: fr, fc: fc, FW: this.colPos[fc], FH: this.rowPos[fr], rtl: sh.rtl };
  };
  // logical x of column c (LTR space, from the viewport's leading edge)
  G.cx = function (c, v) { return HW + this.colPos[c] - (c >= v.fc ? v.sx : 0); };
  G.ry = function (r, v) { return HH + this.rowPos[r] - (r >= v.fr ? v.sy : 0); };
  G.cw = function (c) { return this.colPos[c + 1] - this.colPos[c]; };
  G.rh = function (r) { return this.rowPos[r + 1] - this.rowPos[r]; };
  // pointer -> cell / header (lx: logical x)
  G.hit = function (clientX, clientY) {
    this.axes();
    var rect = this.scroller.getBoundingClientRect(), v = this.view();
    var lx = v.rtl ? rect.left + v.W - clientX : clientX - rect.left, y = clientY - rect.top;
    var o = { lx: lx, y: y, inside: lx >= 0 && y >= 0 && lx <= v.W && y <= v.H };
    o.r = y < HH ? -1 : y < HH + v.FH ? this.rowAt(y - HH) : this.rowAt(y - HH + v.sy);
    o.c = lx < HW ? -1 : lx < HW + v.FW ? this.colAt(lx - HW) : this.colAt(lx - HW + v.sx);
    if (y < HH && lx < HW) o.zone = "corner"; else if (y < HH) o.zone = "colhdr"; else if (lx < HW) o.zone = "rowhdr"; else o.zone = "cell";
    if (o.zone === "colhdr") { var c = o.c, x1 = this.cx(c, v) + this.cw(c); if (Math.abs(lx - x1) <= 4) o.colEdge = c; else if (Math.abs(lx - this.cx(c, v)) <= 3 && c > 0) o.colEdge = c - 1; }
    if (o.zone === "rowhdr") { var r = o.r, y1 = this.ry(r, v) + this.rh(r); if (Math.abs(y - y1) <= 3) o.rowEdge = r; else if (Math.abs(y - this.ry(r, v)) <= 2 && r > 0) o.rowEdge = r - 1; }
    if (o.zone === "cell") {
      var s = this.sel, e = this.selRect(s.r1, s.c1, v);
      if (e && Math.abs(lx - (e.x + e.w)) <= 5 && Math.abs(y - (e.y + e.h)) <= 5 && !this.edit) o.fillHandle = true;
      var f = this.sheet().filter;
      if (f && o.r === f.r0 && o.c >= f.c0 && o.c <= f.c1) { var xr = this.cx(o.c, v) + this.cw(o.c); if (lx >= xr - 20) o.filterBtn = o.c; }
    }
    return o;
  };
  G.selRect = function (r, c, v) {
    var m = this.wb.mergeAt(this.sheet(), r, c), r1 = r, c1 = c;
    if (m) { r = m.r0; c = m.c0; r1 = m.r1; c1 = m.c1; }
    return { x: this.cx(c, v), y: this.ry(r, v), w: this.cx(c1, v) + this.cw(c1) - this.cx(c, v), h: this.ry(r1, v) + this.rh(r1) - this.ry(r, v) };
  };
  // physical rect (relative to the scroller) of a cell, for popovers and the editor
  G.cellBox = function (r, c) {
    this.axes();
    var v = this.view(), b = this.selRect(r, c, v);
    return { x: v.rtl ? v.W - b.x - b.w : b.x, y: b.y, w: b.w, h: b.h, v: v };
  };

  // ---------- drawing ----------
  G.draw = function () { var self = this; if (this.raf) return; this.raf = requestAnimationFrame(function () { self.raf = 0; self.drawNow(); }); };
  G.onScroll = function () {
    var sh = this.sheet(), sy = this.scroller.scrollTop, maxY = this.scroller.scrollHeight - this.scroller.clientHeight;
    if (sy > maxY - 400 && sh.nrows < MAXR) { sh.nrows = Math.min(MAXR, sh.nrows + 1000); this.axKey = ""; }
    var sx = Math.abs(this.scroller.scrollLeft), maxX = this.scroller.scrollWidth - this.scroller.clientWidth;
    if (sx > maxX - 300 && sh.ncols < MAXC) { sh.ncols = Math.min(MAXC, sh.ncols + 10); this.axKey = ""; }
    if (this.edit) this.placeEditor();
    this.drawNow();
  };
  G.drawNow = function () {
    if (this.raf) { cancelAnimationFrame(this.raf); this.raf = 0; }
    this.axes();
    var sh = this.sheet(), v = this.view(), W = v.W, H = v.H, dpr = window.devicePixelRatio || 1, cv = this.canvas, g = this.ctx, T = this.theme();
    if (!W || !H) return;
    this.scroller.setAttribute("dir", v.rtl ? "rtl" : "ltr");
    this.sticky.classList.toggle("rtl", !!v.rtl);
    var tw = HW + this.colPos[sh.ncols] + 80, th = HH + this.rowPos[sh.nrows] + 80;
    this.spacer.style.width = tw + "px"; this.spacer.style.height = th + "px";
    if (cv.width !== Math.round(W * dpr) || cv.height !== Math.round(H * dpr)) { cv.width = Math.round(W * dpr); cv.height = Math.round(H * dpr); }
    cv.style.width = W + "px"; cv.style.height = H + "px";
    this.overlay.style.width = W + "px"; this.overlay.style.height = H + "px";
    this.dpr = dpr; this.v = v;
    this.geo(); g.fillStyle = T.bg; g.fillRect(0, 0, W, H);
    this.cfPrep(sh);
    var X0 = HW + v.FW, Y0 = HH + v.FH;
    var panes = [{ ox: v.sx, oy: v.sy, x0: X0, y0: Y0, x1: W, y1: H, fR: false, fC: false }];
    if (v.fr) panes.push({ ox: v.sx, oy: 0, x0: X0, y0: HH, x1: W, y1: Y0, fR: true, fC: false });
    if (v.fc) panes.push({ ox: 0, oy: v.sy, x0: HW, y0: Y0, x1: X0, y1: H, fR: false, fC: true });
    if (v.fr && v.fc) panes.push({ ox: 0, oy: 0, x0: HW, y0: HH, x1: X0, y1: Y0, fR: true, fC: true });
    var self = this;
    panes.forEach(function (p) { if (p.x1 > p.x0 && p.y1 > p.y0) self.drawPane(p, v, T); });
    panes.forEach(function (p) { if (p.x1 > p.x0 && p.y1 > p.y0) self.drawSel(p, v, T); });
    this.drawHeaders(v, T, panes);
    this.geo(); g.strokeStyle = T.frozen; g.lineWidth = 1.5;
    if (v.fr) { g.beginPath(); g.moveTo(0, Y0 - 0.5); g.lineTo(W, Y0 - 0.5); g.stroke(); }
    if (v.fc) { g.beginPath(); g.moveTo(X0 - 0.5, 0); g.lineTo(X0 - 0.5, H); g.stroke(); }
    this.layoutCharts(v);
    if (this.cb.onDraw) this.cb.onDraw();
  };
  G.geo = function () { var d = this.dpr || 1, v = this.v; if (v && v.rtl) this.ctx.setTransform(-d, 0, 0, d, v.W * d, 0); else this.ctx.setTransform(d, 0, 0, d, 0, 0); };
  G.txt = function () { var d = this.dpr || 1; this.ctx.setTransform(d, 0, 0, d, 0, 0); };
  G.measure = function (font, t) {
    var k = font + "\u0001" + t, w = this.widths.get(k);
    if (w === undefined) { this.ctx.font = font; w = this.ctx.measureText(t).width; if (this.widths.size > 8000) this.widths.clear(); this.widths.set(k, w); }
    return w;
  };
  var RTL_CH = /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/;
  function rtlFirst(t) { var m = /[A-Za-z\u00C0-\u024F\u0370-\u04FF\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/.exec(t); return !!m && RTL_CH.test(m[0]); }
  function fontOf(s) { return (s && s.i ? "italic " : "") + (s && s.b ? "700 " : "400 ") + ((s && s.fs) || FPX) + "px " + FONT; }
  G.cellLook = function (sh, r, c, cell) {
    var v = this.wb.eng.value(sh, r, c), s = cell && cell.s, fm = F.format(v, s && s.nf), cf = this.cfAt(sh, r, c, v);
    var al = s && s.ha;
    if (!al) al = typeof v === "number" ? "right" : typeof v === "boolean" || v instanceof C.Err ? "center" : "left";
    return { r: r, c: c, v: v, s: s, t: fm.t, color: (cf && cf.fc) || fm.color || (s && s.fc) || null, bg: (cf && cf.bg) || (s && s.bg) || null, bold: cf && cf.b, al: al, bar: cf && cf.bar };
  };
  G.drawPane = function (p, v, T) {
    var g = this.ctx, sh = this.sheet(), self = this;
    var c0 = this.colAt(p.x0 - HW + p.ox), c1 = this.colAt(p.x1 - HW + p.ox), r0 = this.rowAt(p.y0 - HH + p.oy), r1 = this.rowAt(p.y1 - HH + p.oy);
    if (p.fC) c1 = Math.min(c1, v.fc - 1); else c0 = Math.max(c0, v.fc);
    if (p.fR) r1 = Math.min(r1, v.fr - 1); else r0 = Math.max(r0, v.fr);
    function X(c) { return HW + self.colPos[c] - p.ox; }
    function Y(r) { return HH + self.rowPos[r] - p.oy; }
    this.geo(); g.save(); g.beginPath(); g.rect(p.x0, p.y0, p.x1 - p.x0, p.y1 - p.y0); g.clip();
    // grid lines
    if (sh.grid !== false) {
      g.strokeStyle = T.grid; g.lineWidth = 1; g.beginPath();
      for (var c = c0; c <= c1 + 1 && c <= sh.ncols; c++) { var x = Math.round(X(c)) - 0.5; g.moveTo(x, p.y0); g.lineTo(x, p.y1); }
      for (var r = r0; r <= r1 + 1 && r <= sh.nrows; r++) { if (r < sh.nrows && r > 0 && self.rh(r - 1) === 0) continue; var y = Math.round(Y(r)) - 0.5; g.moveTo(p.x0, y); g.lineTo(p.x1, y); }
      g.stroke();
    }
    var texts = [], covered = new Set(), merges = sh.merges.filter(function (m) { return m.r1 >= r0 && m.r0 <= r1 && m.c1 >= c0 && m.c0 <= c1; });
    merges.forEach(function (m) {
      for (var a = m.r0; a <= m.r1; a++) for (var b = m.c0; b <= m.c1; b++) covered.add(a * KW + b);
      var cell = sh.cells.get(m.r0 * KW + m.c0), L = self.cellLook(sh, m.r0, m.c0, cell), x = X(m.c0), y = Y(m.r0), w = X(m.c1) + self.cw(m.c1) - x, h = Y(m.r1) + self.rh(m.r1) - y;
      g.fillStyle = L.bg || T.bg; g.fillRect(x, y, w - 1, h - 1);
      if (cell && cell.s && cell.s.bd) self.borders(cell.s.bd, x, y, w, h);
      if (L.t) texts.push(self.textItem(L, cell, x, y, w, h, x, w));
    });
    var empty = function (rr, cc) { return !covered.has(rr * KW + cc) && self.wb.eng.value(sh, rr, cc) === null; };
    for (r = r0; r <= r1; r++) {
      var rh = this.rh(r); if (!rh) continue;
      var y0 = Y(r);
      for (c = c0; c <= c1; c++) {
        var k = r * KW + c;
        if (covered.has(k)) continue;
        var cell = sh.cells.get(k);
        if (!cell && !(this.wb.eng.spill.size && this.wb.eng.spillOf(sh, r, c)) && !this.cfRules.length) continue;
        var L = this.cellLook(sh, r, c, cell), x0 = X(c), w = this.cw(c);
        if (L.bg) { g.fillStyle = L.bg; g.fillRect(x0, y0, w - 1, rh - 1); }
        if (L.bar) { g.fillStyle = L.bar.color; g.globalAlpha = 0.32; g.fillRect(x0 + 2, y0 + 3, Math.max(0, (w - 5) * L.bar.f), rh - 7); g.globalAlpha = 1; }
        if (cell && cell.s && cell.s.bd) this.borders(cell.s.bd, x0, y0, w, rh);
        if (!L.t) continue;
        // let left/right aligned text spill over empty neighbours
        var clipX = x0, clipW = w;
        if (!(L.s && L.s.wr) && typeof L.v !== "number") {
          var fw = this.measure(fontOf(L.s), L.t) + 6;
          if (fw > w) {
            if (L.al === "left") { var cc = c + 1, ext = w; while (ext < fw && cc < sh.ncols && cc <= c1 + 8 && empty(r, cc)) { ext += this.cw(cc); cc++; } clipW = ext; }
            else if (L.al === "right") { var cl = c - 1, ex2 = w; while (ex2 < fw && cl >= 0 && cl >= c0 - 8 && empty(r, cl)) { ex2 += this.cw(cl); clipX -= this.cw(cl); cl--; } clipW = ex2; }
          }
        }
        texts.push(this.textItem(L, cell, x0, y0, w, rh, clipX, clipW));
      }
    }
    g.restore();
    // text, drawn un-mirrored
    this.txt(); g.save(); g.beginPath();
    var px0 = v.rtl ? v.W - p.x1 : p.x0; g.rect(px0, p.y0, p.x1 - p.x0, p.y1 - p.y0); g.clip();
    g.textBaseline = "alphabetic";
    texts.forEach(function (t) { self.drawText(t, v, T); });
    g.restore();
    // filter buttons
    var f = sh.filter;
    if (f && f.r0 >= r0 && f.r0 <= r1) {
      this.geo(); g.save(); g.beginPath(); g.rect(p.x0, p.y0, p.x1 - p.x0, p.y1 - p.y0); g.clip();
      for (c = Math.max(c0, f.c0); c <= Math.min(c1, f.c1); c++) {
        var bx = X(c) + this.cw(c) - 19, by = Y(f.r0) + (this.rh(f.r0) - 16) / 2, on = !!f.cols[c];
        g.fillStyle = on ? T.sel : T.hdrBg; g.strokeStyle = on ? T.sel : T.hdrLine; g.lineWidth = 1;
        g.fillRect(bx, by, 16, 16); g.strokeRect(bx + 0.5, by + 0.5, 15, 15);
        g.fillStyle = on ? "#ffffff" : T.hdrText; g.beginPath();
        if (on) { g.moveTo(bx + 4, by + 5); g.lineTo(bx + 12, by + 5); g.lineTo(bx + 9, by + 9); g.lineTo(bx + 9, by + 12); g.lineTo(bx + 7, by + 12); g.lineTo(bx + 7, by + 9); }
        else { g.moveTo(bx + 4.5, by + 6.5); g.lineTo(bx + 11.5, by + 6.5); g.lineTo(bx + 8, by + 10.5); }
        g.closePath(); g.fill();
      }
      g.restore();
    }
  };
  G.textItem = function (L, cell, x, y, w, h, clipX, clipW) {
    var s = L.s, f = this.sheet().filter, padR = f && L.r === f.r0 && L.c >= f.c0 && L.c <= f.c1 ? 20 : 0, font = (s && s.i ? "italic " : "") + (s && s.b || L.bold ? "700 " : "400 ") + ((s && s.fs) || FPX) + "px " + FONT;
    return { t: L.t, font: font, color: L.color, bg: L.bg, al: L.al, auto: !(s && s.ha), va: (s && s.va) || "bottom", wr: s && s.wr, u: s && s.u, st: s && s.st, num: typeof L.v === "number", x: x, y: y, w: w - padR, h: h, cx: clipX, cw: padR ? Math.min(clipW, w - padR) : clipW, fs: (s && s.fs) || FPX, rtl: RTL_CH.test(L.t) && rtlFirst(L.t) };
  };
  G.drawText = function (t, v, T) {
    var g = this.ctx, color = t.color || T.text;
    if (this.dark) { var bgc = t.bg || T.bg; if (contrast(color, bgc) < 3) color = lum(bgc) > 0.4 ? "#1d2228" : "#e6e9ee"; }
    else if (t.bg && !t.color && lum(t.bg) < 0.2) color = "#ffffff";
    g.fillStyle = color; g.font = t.font;
    var al = t.al, pad = 4, text = t.t, lines = [text], lh = Math.round(t.fs * 1.3);
    if (t.wr) lines = this.wrap(t.font, text, t.w - pad * 2);
    else if (text.indexOf("\n") >= 0) lines = text.split("\n");
    if (t.num && !t.wr && this.measure(t.font, text) > t.w - 3) { var n = Math.max(1, Math.floor((t.w - 4) / Math.max(1, this.measure(t.font, "#")))); lines = ["#".repeat(n)]; al = "center"; }
    var g0 = v.rtl ? v.W - t.x - t.w - (t.w < t.cw ? 0 : 0) : t.x, pal = v.rtl && t.auto ? (al === "left" ? "right" : al === "right" ? "left" : al) : al;
    g.direction = t.rtl ? "rtl" : "ltr";
    var cx0 = v.rtl ? v.W - t.cx - t.cw : t.cx;
    g.save(); g.beginPath(); g.rect(cx0, t.y, t.cw - 1, t.h - 1); g.clip();
    var total = lines.length * lh, top = t.va === "top" ? t.y + 2 : t.va === "middle" ? t.y + (t.h - total) / 2 : t.y + t.h - total - 3;
    g.textAlign = pal;
    var ax = pal === "left" ? g0 + pad : pal === "right" ? g0 + t.w - pad : g0 + t.w / 2;
    for (var i = 0; i < lines.length; i++) {
      var by = top + lh * (i + 1) - Math.round((lh - t.fs) / 2) - 2;
      g.fillText(lines[i], ax, by);
      if (t.u || t.st) {
        var tw = this.measure(t.font, lines[i]), lx = pal === "left" ? ax : pal === "right" ? ax - tw : ax - tw / 2;
        g.fillRect(lx, t.u ? by + 2 : by - Math.round(t.fs * 0.32), tw, 1);
      }
    }
    g.restore();
  };
  G.wrap = function (font, text, max) {
    var out = [], self = this;
    String(text).split("\n").forEach(function (para) {
      var words = para.split(/(\s+)/), line = "";
      words.forEach(function (w) {
        var tryL = line + w;
        if (line && self.measure(font, tryL.trimEnd()) > max) { out.push(line.trimEnd()); line = w.trimStart(); }
        else line = tryL;
        while (self.measure(font, line) > max && line.length > 1) {
          var k = line.length; while (k > 1 && self.measure(font, line.slice(0, k)) > max) k--;
          out.push(line.slice(0, k)); line = line.slice(k);
        }
      });
      out.push(line);
    });
    return out.slice(0, 50);
  };
  G.borders = function (bd, x, y, w, h) {
    var g = this.ctx;
    function line(side, x1, y1, x2, y2) {
      var b = bd[side]; if (!b) return;
      g.strokeStyle = b.c; g.lineWidth = b.s === "thick" ? 3 : b.s === "medium" ? 2 : 1; g.setLineDash(b.s === "dashed" ? [4, 3] : []);
      g.beginPath(); g.moveTo(x1, y1); g.lineTo(x2, y2); g.stroke();
    }
    var X1 = Math.round(x) - 0.5, Y1 = Math.round(y) - 0.5, X2 = Math.round(x + w) - 0.5, Y2 = Math.round(y + h) - 0.5;
    line("t", X1, Y1, X2, Y1); line("b", X1, Y2, X2, Y2); line("l", X1, Y1, X1, Y2); line("r", X2, Y1, X2, Y2);
    g.setLineDash([]);
  };
  G.drawSel = function (p, v, T) {
    var g = this.ctx, s = this.sel, self = this, sh = this.sheet();
    function rectOf(g0) {
      var r1 = Math.min(g0.r1, sh.nrows - 1), c1 = Math.min(g0.c1, sh.ncols - 1);
      var x = HW + self.colPos[g0.c0] - (g0.c0 >= v.fc ? p.ox : 0), y = HH + self.rowPos[g0.r0] - (g0.r0 >= v.fr ? p.oy : 0);
      var x2 = HW + self.colPos[c1 + 1] - (c1 >= v.fc ? p.ox : 0), y2 = HH + self.rowPos[r1 + 1] - (r1 >= v.fr ? p.oy : 0);
      return { x: x, y: y, w: x2 - x, h: y2 - y };
    }
    this.geo(); g.save(); g.beginPath(); g.rect(p.x0, p.y0, p.x1 - p.x0, p.y1 - p.y0); g.clip();
    // references of the formula being typed
    this.refs.forEach(function (rf) {
      var q = rectOf(rf); g.strokeStyle = rf.color; g.lineWidth = 2; g.setLineDash([]); g.strokeRect(q.x + 0.5, q.y + 0.5, q.w - 2, q.h - 2);
      g.globalAlpha = 0.08; g.fillStyle = rf.color; g.fillRect(q.x, q.y, q.w - 1, q.h - 1); g.globalAlpha = 1;
    });
    var R = rectOf(s), single = s.r0 === s.r1 && s.c0 === s.c1 || this.isMergeSel();
    if (!single) { g.fillStyle = T.selFill; g.fillRect(R.x, R.y, R.w - 1, R.h - 1); }
    g.strokeStyle = T.sel; g.lineWidth = 2; g.strokeRect(R.x - 0.5, R.y - 0.5, R.w, R.h);
    var act = this.selRect(s.r, s.c, { fc: v.fc, fr: v.fr, sx: p.ox, sy: p.oy });
    if (!single) { g.lineWidth = 1; g.strokeRect(act.x + 0.5, act.y + 0.5, act.w - 2, act.h - 2); }
    if (!this.edit) { g.fillStyle = T.sel; g.fillRect(R.x + R.w - 4, R.y + R.h - 4, 6, 6); g.fillStyle = T.bg; g.fillRect(R.x + R.w - 5, R.y + R.h - 5, 1, 7); g.fillRect(R.x + R.w - 5, R.y + R.h - 5, 7, 1); }
    if (this.fillTo) { var Fq = rectOf(this.fillTo); g.strokeStyle = T.sel; g.lineWidth = 1; g.setLineDash([4, 3]); g.strokeRect(Fq.x + 0.5, Fq.y + 0.5, Fq.w - 1, Fq.h - 1); g.setLineDash([]); }
    if (this.clip && this.clip.uid === sh.uid) { var Q = rectOf(this.clip.g); g.strokeStyle = T.sel; g.lineWidth = 1.5; g.setLineDash([5, 4]); g.lineDashOffset = 0; g.strokeRect(Q.x + 1, Q.y + 1, Q.w - 3, Q.h - 3); g.setLineDash([]); }
    g.restore();
  };
  G.isMergeSel = function () { var s = this.sel, m = this.wb.mergeAt(this.sheet(), s.r0, s.c0); return m && m.r0 === s.r0 && m.c0 === s.c0 && m.r1 === s.r1 && m.c1 === s.c1; };
  G.drawHeaders = function (v, T) {
    var g = this.ctx, sh = this.sheet(), s = this.sel, self = this, W = v.W, H = v.H;
    this.geo();
    g.fillStyle = T.hdrBg; g.fillRect(0, 0, W, HH); g.fillRect(0, 0, HW, H);
    var fullCols = s.r0 === 0 && s.r1 >= MAXR - 1, fullRows = s.c0 === 0 && s.c1 >= MAXC - 1;
    var cols = [], rows = [];
    function colsIn(x0, x1, ox, lim0, lim1) { var a = Math.max(lim0, self.colAt(x0 - HW + ox)), b = Math.min(lim1, self.colAt(x1 - HW + ox)); for (var c = a; c <= b; c++) cols.push({ c: c, x: HW + self.colPos[c] - ox }); }
    function rowsIn(y0, y1, oy, lim0, lim1) { var a = Math.max(lim0, self.rowAt(y0 - HH + oy)), b = Math.min(lim1, self.rowAt(y1 - HH + oy)); for (var r = a; r <= b; r++) if (self.rh(r)) rows.push({ r: r, y: HH + self.rowPos[r] - oy }); }
    if (v.fc) colsIn(HW, HW + v.FW, 0, 0, v.fc - 1);
    colsIn(HW + v.FW, W, v.sx, v.fc, sh.ncols - 1);
    if (v.fr) rowsIn(HH, HH + v.FH, 0, 0, v.fr - 1);
    rowsIn(HH + v.FH, H, v.sy, v.fr, sh.nrows - 1);
    g.save(); g.beginPath(); g.rect(HW, 0, W - HW, HH); g.clip();
    cols.forEach(function (o) { var w = self.cw(o.c); if (o.c >= s.c0 && o.c <= s.c1) { g.fillStyle = fullCols ? T.sel : T.hdrSel; g.fillRect(o.x, 0, w, HH); } g.fillStyle = T.hdrLine; g.fillRect(o.x + w - 1, 4, 1, HH - 8); });
    g.restore();
    g.save(); g.beginPath(); g.rect(0, HH, HW, H - HH); g.clip();
    rows.forEach(function (o) { var h = self.rh(o.r); if (o.r >= s.r0 && o.r <= s.r1) { g.fillStyle = fullRows ? T.sel : T.hdrSel; g.fillRect(0, o.y, HW, h); } g.fillStyle = T.hdrLine; g.fillRect(4, o.y + h - 1, HW - 8, 1); });
    g.restore();
    g.fillStyle = T.hdrLine; g.fillRect(0, HH - 1, W, 1); g.fillRect(HW - 1, 0, 1, H);
    // corner triangle
    g.fillStyle = T.hdrLine; g.beginPath(); g.moveTo(HW - 5, HH - 5); g.lineTo(HW - 5, HH - 15); g.lineTo(HW - 15, HH - 5); g.closePath(); g.fill();
    this.txt(); g.textBaseline = "middle"; g.textAlign = "center"; g.font = "500 11.5px " + FONT;
    g.save(); g.beginPath(); g.rect(v.rtl ? 0 : HW, 0, W - HW, HH); g.clip();
    cols.forEach(function (o) { var w = self.cw(o.c), sel = o.c >= s.c0 && o.c <= s.c1; g.fillStyle = sel ? (fullCols ? "#ffffff" : T.hdrSelText) : T.hdrText; var px = v.rtl ? W - o.x - w : o.x; g.fillText(C.colName(o.c), px + w / 2, HH / 2 + 0.5); });
    g.restore();
    g.save(); g.beginPath(); g.rect(v.rtl ? W - HW : 0, HH, HW, H - HH); g.clip();
    rows.forEach(function (o) { var h = self.rh(o.r), sel = o.r >= s.r0 && o.r <= s.r1; g.fillStyle = sel ? (fullRows ? "#ffffff" : T.hdrSelText) : T.hdrText; g.fillText(String(o.r + 1), v.rtl ? W - HW / 2 : HW / 2, o.y + h / 2 + 0.5); });
    g.restore();
  };

  // ---------- conditional formatting ----------
  G.cfPrep = function (sh) {
    this.cfRules = sh.cf || [];
    if (!this.cfCache || this.cfCache.ver !== this.wb.eng.ver || this.cfCache.uid !== sh.uid) this.cfCache = { ver: this.wb.eng.ver, uid: sh.uid, stats: new Map(), fx: new Map() };
  };
  function numOf(t) { var p = F.parseInput(String(t)); return typeof p.v === "number" ? p.v : null; }
  G.cfStats = function (sh, rule) {
    var st = this.cfCache.stats.get(rule.id);
    if (st) return st;
    var wb = this.wb, nums = [], counts = new Map(), g = wb.clampArea(sh, rule.range), lim = 0;
    for (var r = g.r0; r <= g.r1 && lim < 200000; r++) for (var c = g.c0; c <= g.c1; c++, lim++) {
      var v = wb.eng.value(sh, r, c);
      if (typeof v === "number") nums.push(v);
      if (rule.type === "dup" && v !== null && v !== "") { var k = typeof v === "string" ? v.toLowerCase() : String(v); counts.set(k, (counts.get(k) || 0) + 1); }
    }
    nums.sort(function (a, b) { return a - b; });
    st = { min: nums.length ? nums[0] : 0, max: nums.length ? nums[nums.length - 1] : 0, nums: nums, counts: counts };
    if (rule.type === "top") { var n = Math.max(1, parseInt(rule.a, 10) || 10); st.cut = nums.length ? nums[Math.max(0, nums.length - n)] : Infinity; }
    if (rule.type === "scale" && (rule.colors || []).length === 3) st.mid = nums.length ? nums[Math.floor(nums.length / 2)] : 0;
    this.cfCache.stats.set(rule.id, st);
    return st;
  };
  G.cfAt = function (sh, r, c, v) {
    var rules = this.cfRules, out = null;
    for (var i = 0; i < rules.length; i++) {
      var R = rules[i], g = R.range;
      if (r < g.r0 || r > g.r1 || c < g.c0 || c > g.c1) continue;
      var hit = false, a = numOf(R.a), b = numOf(R.b), x;
      switch (R.type) {
        case "gt": hit = typeof v === "number" && a !== null && v > a; break;
        case "lt": hit = typeof v === "number" && a !== null && v < a; break;
        case "ge": hit = typeof v === "number" && a !== null && v >= a; break;
        case "le": hit = typeof v === "number" && a !== null && v <= a; break;
        case "eq": hit = a !== null ? v === a : typeof v === "string" && v.toLowerCase() === String(R.a).toLowerCase(); break;
        case "ne": hit = v !== null && (a !== null ? v !== a : !(typeof v === "string" && v.toLowerCase() === String(R.a).toLowerCase())); break;
        case "between": hit = typeof v === "number" && a !== null && b !== null && v >= Math.min(a, b) && v <= Math.max(a, b); break;
        case "contains": hit = v !== null && R.a !== "" && F.format(v).t.toLowerCase().indexOf(String(R.a).toLowerCase()) >= 0; break;
        case "blank": hit = v === null || v === ""; break;
        case "notblank": hit = !(v === null || v === ""); break;
        case "dup": if (v !== null && v !== "") { var st = this.cfStats(sh, R); hit = st.counts.get(typeof v === "string" ? v.toLowerCase() : String(v)) > 1; } break;
        case "top": hit = typeof v === "number" && v >= this.cfStats(sh, R).cut; break;
        case "formula": {
          var key = R.id + ":" + r + ":" + c;
          x = this.cfCache.fx.get(key);
          if (x === undefined) { var f = String(R.a).replace(/^=/, ""); x = this.wb.eng.evalFormula(EN.moveFormula(f, r - g.r0, c - g.c0), sh, r, c); x = x === true || (typeof x === "number" && x !== 0); this.cfCache.fx.set(key, x); }
          hit = x; break;
        }
        case "scale": if (typeof v === "number") {
          var s2 = this.cfStats(sh, R), cols = R.colors && R.colors.length >= 2 ? R.colors : ["#f8696b", "#63be7b"], t;
          if (cols.length === 3) { var mid = s2.mid; t = v <= mid ? (mid === s2.min ? 0 : (v - s2.min) / (mid - s2.min)) / 2 : 0.5 + (s2.max === mid ? 0 : (v - mid) / (s2.max - mid)) / 2; }
          else t = s2.max === s2.min ? 0.5 : (v - s2.min) / (s2.max - s2.min);
          t = Math.max(0, Math.min(1, t));
          var col = cols.length === 3 ? (t < 0.5 ? mix(cols[0], cols[1], t * 2) : mix(cols[1], cols[2], (t - 0.5) * 2)) : mix(cols[0], cols[1], t);
          out = out || {}; if (!out.bg) out.bg = col;
        } continue;
        case "bar": if (typeof v === "number") { var s3 = this.cfStats(sh, R), lo = Math.min(0, s3.min), span = s3.max - lo || 1; out = out || {}; out.bar = { f: Math.max(0, (v - lo) / span), color: (R.colors && R.colors[0]) || "#2f7fd8" }; } continue;
      }
      if (hit) { out = out || {}; var sty = R.style || {}; if (sty.bg && !out.bg) out.bg = sty.bg; if (sty.fc && !out.fc) out.fc = sty.fc; if (sty.b) out.b = true; }
    }
    return out;
  };

  // ---------- selection & navigation ----------
  G.expand = function (g) {
    var sh = this.sheet(), changed = true, guard = 0;
    while (changed && guard++ < 20) {
      changed = false;
      sh.merges.forEach(function (m) {
        if (C.overlaps(m, g) && (m.r0 < g.r0 || m.c0 < g.c0 || m.r1 > g.r1 || m.c1 > g.c1)) { g.r0 = Math.min(g.r0, m.r0); g.c0 = Math.min(g.c0, m.c0); g.r1 = Math.max(g.r1, m.r1); g.c1 = Math.max(g.c1, m.c1); changed = true; }
      });
    }
    return g;
  };
  G.select = function (r0, c0, r1, c1, ar, ac, noScroll) {
    var sh = this.sheet(), g = this.expand({ r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) });
    if (ar === undefined) { ar = r0; ac = c0; }
    var m = this.wb.mergeAt(sh, ar, ac); if (m) { ar = m.r0; ac = m.c0; }
    this.sel = { r: ar, c: ac, r0: g.r0, c0: g.c0, r1: Math.min(g.r1, MAXR - 1), c1: Math.min(g.c1, MAXC - 1) };
    if (!noScroll) this.scrollTo(this.anchor ? this.anchor.r : ar, this.anchor ? this.anchor.c : ac);
    this.draw();
    var txt = C.addr(ar, ac) + ": " + (this.wb.text(sh, ar, ac) || "empty");
    this.live.textContent = txt;
    if (this.cb.onSel) this.cb.onSel(this.sel);
    if (!this.edit) this.parkEditor();
  };
  G.selectAll = function () { this.select(0, 0, MAXR - 1, MAXC - 1, this.sel.r, this.sel.c, true); };
  G.scrollTo = function (r, c) {
    this.axes();
    var v = this.view(), sc = this.scroller;
    if (r >= v.fr && r < this.sheet().nrows) {
      var top = this.rowPos[r] - v.FH, bot = this.rowPos[r + 1] - v.FH, vh = v.H - HH - v.FH;
      if (top < v.sy) sc.scrollTop = top; else if (bot > v.sy + vh) sc.scrollTop = bot - vh;
    }
    if (c >= v.fc && c < this.sheet().ncols) {
      var left = this.colPos[c] - v.FW, right = this.colPos[c + 1] - v.FW, vw = v.W - HW - v.FW, cur = v.sx;
      var nx = left < cur ? left : right > cur + vw ? Math.min(left, right - vw) : cur;
      if (nx !== cur) sc.scrollLeft = v.rtl ? -nx : nx;
    }
  };
  G.visible = function (r) { return this.rh(r) > 0; };
  G.step = function (r, c, dr, dc) {
    var sh = this.sheet(), m = this.wb.mergeAt(sh, r, c);
    if (m) { if (dr > 0) r = m.r1; if (dc > 0) c = m.c1; if (dr < 0) r = m.r0; if (dc < 0) c = m.c0; }
    r += dr; c += dc;
    while (dr && r > 0 && r < sh.nrows - 1 && !this.visible(r)) r += dr;
    return { r: C.clamp(r, 0, MAXR - 1), c: C.clamp(c, 0, MAXC - 1) };
  };
  G.jump = function (r, c, dr, dc) {          // Ctrl+arrow: to the edge of the data
    var sh = this.sheet(), eng = this.wb.eng, maxR = Math.max(sh.nrows - 1, sh.maxR), maxC = Math.max(sh.ncols - 1, sh.maxC);
    function has(rr, cc) { return eng.value(sh, rr, cc) !== null; }
    function inb(rr, cc) { return rr >= 0 && cc >= 0 && rr <= maxR && cc <= maxC; }
    var nr = r + dr, nc = c + dc;
    if (!inb(nr, nc)) return { r: r, c: c };
    if (has(r, c) && has(nr, nc)) { while (inb(nr + dr, nc + dc) && has(nr + dr, nc + dc)) { nr += dr; nc += dc; } return { r: nr, c: nc }; }
    while (inb(nr, nc) && !has(nr, nc) && (dr ? Math.abs(nr - r) < 1100000 : true)) { if (!inb(nr + dr, nc + dc)) return { r: nr, c: nc }; nr += dr; nc += dc; if (dr > 0 && nr > sh.maxR && nr > r) return { r: Math.min(maxR, Math.max(sh.maxR, r)), c: nc }; if (dc > 0 && nc > sh.maxC && nc > c) return { r: nr, c: Math.min(maxC, Math.max(sh.maxC, c)) }; }
    return { r: nr, c: nc };
  };
  G.move = function (dr, dc, extend, ctrl) {
    var s = this.sel, base = extend ? (this.anchor2 || { r: s.r === s.r0 ? s.r1 : s.r0, c: s.c === s.c0 ? s.c1 : s.c0 }) : { r: s.r, c: s.c };
    var to = ctrl ? this.jump(base.r, base.c, dr, dc) : this.step(base.r, base.c, dr, dc);
    if (extend) { this.anchor2 = to; this.select(s.r, s.c, to.r, to.c, s.r, s.c, true); this.scrollTo(to.r, to.c); }
    else { this.anchor2 = null; this.select(to.r, to.c, to.r, to.c); }
  };
  // Enter / Tab inside a selection cycle through it
  G.advance = function (dr, dc) {
    var s = this.sel;
    if (s.r0 === s.r1 && s.c0 === s.c1 || this.isMergeSel()) { var to = this.step(s.r, s.c, dr, dc); this.select(to.r, to.c, to.r, to.c); return; }
    var r = s.r, c = s.c;
    if (dr) { r += dr; if (r > s.r1) { r = s.r0; c = c + 1 > s.c1 ? s.c0 : c + 1; } else if (r < s.r0) { r = s.r1; c = c - 1 < s.c0 ? s.c1 : c - 1; } }
    else { c += dc; if (c > s.c1) { c = s.c0; r = r + 1 > s.r1 ? s.r0 : r + 1; } else if (c < s.c0) { c = s.c1; r = r - 1 < s.r0 ? s.r1 : r - 1; } }
    this.sel.r = r; this.sel.c = c; this.scrollTo(r, c); this.draw(); if (this.cb.onSel) this.cb.onSel(this.sel); this.parkEditor();
  };

  // ---------- editor ----------
  G.parkEditor = function () {
    var e = this.editor;
    e.classList.add("parked"); e.value = "";
    var b = this.cellBox(this.sel.r, this.sel.c);
    e.style.left = Math.max(0, b.x) + "px"; e.style.top = Math.max(0, b.y) + "px";
  };
  G.hideEditor = function () { this.editor.classList.add("parked"); this.editor.value = ""; this.ac.hidden = true; this.tip.hidden = true; this.refs = []; };
  G.beginEdit = function (mode, text, fromBar) {
    var s = this.sel, sh = this.sheet();
    if (this.edit) return;
    var orig = this.wb.inputText(sh, s.r, s.c);
    this.edit = { mode: mode, r: s.r, c: s.c, uid: sh.uid, orig: orig, ins: null };
    var e = this.editor;
    e.classList.remove("parked");
    if (text !== undefined) e.value = text; else e.value = orig;
    var st = this.wb.style(sh, s.r, s.c) || {};
    e.style.font = fontOf(st); e.style.textAlign = st.ha || "start"; e.style.direction = sh.rtl ? "rtl" : "";
    this.placeEditor();
    if (!fromBar) { e.focus({ preventScroll: true }); var n = e.value.length; e.setSelectionRange(n, n); }
    this.onInput();
    this.draw();
  };
  G.placeEditor = function () {
    if (!this.edit) return;
    var e = this.editor, b = this.cellBox(this.edit.r, this.edit.c), v = b.v;
    var maxW = Math.max(120, v.W - b.x - 8);
    e.style.left = b.x + "px"; e.style.top = b.y + "px"; e.style.minWidth = b.w + "px"; e.style.minHeight = b.h + "px"; e.style.maxWidth = maxW + "px";
    e.style.height = "0px";
    var meas = Math.min(this.measure(e.style.font || fontOf(null), e.value.split("\n").reduce(function (a, l) { return l.length > a.length ? l : a; }, "")) + 24, maxW);
    e.style.width = Math.max(b.w, meas) + "px";
    e.style.height = Math.max(b.h, e.scrollHeight) + "px";
  };
  G.commit = function (dr, dc, all) {
    if (!this.edit) return;
    var ed = this.edit, sh = this.wb.sheetByUid(ed.uid), text = this.editor.value, self = this, s = this.sel;
    this.edit = null; this.hideEditor();
    if (sh && (text !== ed.orig || all)) {
      try {
        this.wb.run("Edit", function () {
          self.wb.setInput(sh, ed.r, ed.c, text);
          if (all) {
            var x = self.wb.cell(sh, ed.r, ed.c), a = self.wb.clampArea(sh, s);
            for (var r = a.r0; r <= a.r1; r++) for (var c = a.c0; c <= a.c1; c++) {
              if (r === ed.r && c === ed.c) continue;
              if (x && x.f !== undefined) self.wb.setInput(sh, r, c, "=" + EN.moveFormula(x.f, r - ed.r, c - ed.c)); else self.wb.setInput(sh, r, c, text);
            }
          }
        });
      } catch (e) { S3.toast(e.message || "Could not save that cell.", "error"); }
    }
    if (this.cb.onCommit) this.cb.onCommit();
    if (all) { this.draw(); this.parkEditor(); this.focus(); return; }
    if (dr || dc) this.advance(dr, dc); else { this.select(s.r0, s.c0, s.r1, s.c1, s.r, s.c, true); }
    this.focus();
  };
  G.cancel = function () { if (!this.edit) return; this.edit = null; this.hideEditor(); if (this.cb.onCommit) this.cb.onCommit(); this.parkEditor(); this.draw(); this.focus(); };
  G.focus = function () { if (this.touch) this.scroller.focus({ preventScroll: true }); else this.editor.focus({ preventScroll: true }); };
  // can a click insert a reference at the caret? (after = ( , + - * / ^ & < > :)
  G.pointable = function () {
    if (!this.edit) return false;
    var e = this.editor, t = e.value;
    if (t[0] !== "=") return false;
    var i = e.selectionStart, ins = this.edit.ins;
    if (ins && i === ins.end) return true;
    var before = t.slice(0, i).replace(/\s+$/, "");
    return /[=(,;+\-*/^&<>:]$/.test(before);
  };
  G.insertRef = function (g) {
    var e = this.editor, t = e.value, ins = this.edit.ins, sh = this.sheet(), ed = this.edit;
    var txt = (sh.uid !== ed.uid ? C.quoteSheet(sh.name) + "!" : "") + (g.r0 === g.r1 && g.c0 === g.c1 ? C.addr(g.r0, g.c0) : C.addr(g.r0, g.c0) + ":" + C.addr(g.r1, g.c1));
    var a = ins && e.selectionStart === ins.end ? ins.start : e.selectionStart, b = ins && e.selectionStart === ins.end ? ins.end : e.selectionEnd;
    e.value = t.slice(0, a) + txt + t.slice(b);
    this.edit.ins = { start: a, end: a + txt.length };
    e.setSelectionRange(a + txt.length, a + txt.length);
    this.onInput();
  };
  G.onInput = function () {
    if (!this.edit) return;
    this.placeEditor();
    var t = this.editor.value, self = this;
    this.refs = [];
    if (t[0] === "=") {
      var toks = EN.tokenize(t.slice(1), true), i = 0, sh = this.sheet(), ed = this.wb.sheetByUid(this.edit.uid);
      toks.forEach(function (tk) {
        if (tk.t !== "ref") return;
        var target = tk.sheet ? self.wb.sheetByName(tk.sheet) : ed;
        if (target === sh) self.refs.push({ r0: tk.r0, c0: tk.c0, r1: Math.min(tk.r1, sh.nrows - 1), c1: Math.min(tk.c1, sh.ncols - 1), color: REF_COLORS[i % REF_COLORS.length] });
        i++;
      });
    }
    this.autocomplete();
    if (this.cb.onEdit) this.cb.onEdit(t);
    this.draw();
  };
  G.autocomplete = function () {
    var e = this.editor, t = e.value, i = e.selectionStart, ac = this.ac, tip = this.tip, self = this;
    ac.hidden = true; tip.hidden = true; this.acItems = null;
    if (!this.edit || t[0] !== "=") return;
    var before = t.slice(0, i), m = /(^|[=(,;+\-*/^&<>:\s])([A-Za-z][A-Za-z0-9.]*)$/.exec(before.slice(1).length ? before : "");
    var b = this.cellBox(this.edit.r, this.edit.c), top = b.y + Math.max(b.h, parseFloat(e.style.height) || b.h) + 2;
    if (m && !/^[A-Za-z]{1,3}\d+$/.test(m[2])) {
      var pre = m[2].toUpperCase(), list = SH.fnNames().filter(function (n) { return n.indexOf(pre) === 0; }).sort(function (x, y) { return rankFn(x, pre) - rankFn(y, pre) || (x < y ? -1 : 1); }).slice(0, 8);
      if (list.length && !(list.length === 1 && list[0] === pre && /\($/.test(t.slice(i, i + 1)))) {
        ac.textContent = "";
        this.acItems = { list: list, idx: 0, start: i - m[2].length, end: i };
        list.forEach(function (n, k) {
          var doc = (SH.FNDOC[n] || "").split("|"), row = S3.h("div", { "class": "sh-acrow" + (k ? "" : " on"), role: "option", "data-i": String(k) }, [S3.h("b", { text: n }), S3.h("span", { text: doc[1] || "" })]);
          row.addEventListener("mousedown", function (ev) { ev.preventDefault(); self.acItems.idx = k; self.acPick(); });
          ac.appendChild(row);
        });
        ac.style.left = Math.max(0, Math.min(b.x, b.v.W - 300)) + "px"; ac.style.top = top + "px"; ac.hidden = false;
        return;
      }
    }
    // signature of the function the caret is inside
    var depth = 0;
    for (var j = before.length - 1; j >= 0; j--) {
      var ch = before[j];
      if (ch === ")") depth++;
      else if (ch === "(") { if (depth === 0) { var nm = /([A-Za-z][A-Za-z0-9.]*)$/.exec(before.slice(0, j)); if (nm && SH.FNDOC[nm[1].toUpperCase()]) { var d = SH.FNDOC[nm[1].toUpperCase()].split("|"); tip.textContent = d[0] + " — " + d[1]; tip.style.left = Math.max(0, Math.min(b.x, b.v.W - 320)) + "px"; tip.style.top = top + "px"; tip.hidden = false; } break; } depth--; }
    }
  };
  var POPULAR = "SUM AVERAGE COUNT IF IFERROR VLOOKUP XLOOKUP COUNTIF SUMIF COUNTA MAX MIN ROUND INDEX MATCH SUMIFS COUNTIFS AND OR TODAY NOW DATE LEFT RIGHT MID LEN TEXT CONCAT TEXTJOIN TRIM UPPER LOWER PROPER SUBSTITUTE FIND SEARCH AVERAGEIF AVERAGEIFS ROUNDUP ROUNDDOWN ABS INT MOD SQRT POWER PMT FV PV UNIQUE SORT FILTER SEQUENCE YEAR MONTH DAY EDATE EOMONTH NETWORKDAYS DATEDIF IFS SWITCH SUMPRODUCT MEDIAN".split(" ");
  function rankFn(n, pre) { if (n === pre) return -1; var i = POPULAR.indexOf(n); return i < 0 ? 1000 : i; }
  G.acMove = function (d) { var a = this.acItems; if (!a) return; a.idx = (a.idx + d + a.list.length) % a.list.length; Array.prototype.forEach.call(this.ac.children, function (el, k) { el.classList.toggle("on", k === a.idx); }); };
  G.acPick = function () {
    var a = this.acItems, e = this.editor; if (!a) return;
    var name = a.list[a.idx], t = e.value;
    e.value = t.slice(0, a.start) + name + "(" + t.slice(a.end).replace(/^\(/, "");
    var p = a.start + name.length + 1; e.setSelectionRange(p, p); this.ac.hidden = true; this.acItems = null; this.onInput();
  };

  // ---------- clipboard ----------
  function tsvCell(t) { return /["\t\r\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t; }
  G.copy = function (e, cut) {
    var sh = this.sheet(), s = this.sel, a = this.wb.clampArea(sh, s), lines = [], html = ["<table>"];
    if ((a.r1 - a.r0 + 1) * (a.c1 - a.c0 + 1) > 2e6) { S3.toast("That selection is too large to copy.", "error"); return; }
    for (var r = a.r0; r <= a.r1; r++) {
      if (sh.hidden.has(r)) continue;
      var row = [], hr = ["<tr>"];
      for (var c = a.c0; c <= a.c1; c++) { var t = this.wb.text(sh, r, c); row.push(tsvCell(t)); hr.push("<td>" + S3.esc(t) + "</td>"); }
      lines.push(row.join("\t")); hr.push("</tr>"); html.push(hr.join(""));
    }
    html.push("</table>");
    var tsv = lines.join("\r\n");
    this.clip = { uid: sh.uid, g: a, block: this.wb.copyBlock(sh, a), tsv: tsv, cut: !!cut };
    if (e && e.clipboardData) { e.clipboardData.setData("text/plain", tsv); e.clipboardData.setData("text/html", html.join("")); e.preventDefault(); }
    this.draw();
    return tsv;
  };
  function norm(t) { return String(t || "").replace(/\r\n/g, "\n").replace(/\n$/, ""); }
  G.paste = function (text) {
    var sh = this.sheet(), s = this.sel, self = this, out;
    try {
      if (this.clip && norm(text) === norm(this.clip.tsv)) {
        var clip = this.clip;
        this.wb.run(clip.cut ? "Move" : "Paste", function () {
          out = self.wb.pasteBlock(sh, s.r0, s.c0, clip.block, "all", s);
          if (clip.cut) {
            var src = self.wb.sheetByUid(clip.uid);
            if (src) self.wb.forCells(src, clip.g, function (r, c) { if (!(src === sh && C.inRange(out, r, c))) self.wb.set(src, r, c, null); });
            self.clip = null;
          }
        });
      } else {
        var grid = SH.io.parseDelimited(String(text).replace(/\r?\n$/, ""), "\t", { rows: 200000, cols: 2000, text: 32767 }).rows;
        if (!grid.length) return;
        out = this.wb.pasteText(sh, s.r0, s.c0, grid);
      }
    } catch (err) { S3.toast(err.message || "Paste failed", "error"); return; }
    if (out) this.select(out.r0, out.c0, out.r1, out.c1, out.r0, out.c0);
  };

  // ---------- charts (floating boxes over the grid) ----------
  G.layoutCharts = function (v) {
    var sh = this.sheet(), self = this, seen = new Set(), T = this.theme();
    this.chartLayer.style.left = (v.rtl ? 0 : HW) + "px"; this.chartLayer.style.top = HH + "px";
    this.chartLayer.style.width = (v.W - HW) + "px"; this.chartLayer.style.height = (v.H - HH) + "px";
    sh.charts.forEach(function (ch) {
      seen.add(ch.id);
      var box = self.chartBoxes.get(ch.id);
      if (!box) box = self.makeChartBox(ch);
      box.ch = ch;
      var x = ch.x - v.sx, y = ch.y - v.sy;
      box.el.style.top = y + "px";
      if (v.rtl) { box.el.style.right = x + "px"; box.el.style.left = "auto"; } else { box.el.style.left = x + "px"; box.el.style.right = "auto"; }
      box.el.style.width = ch.w + "px"; box.el.style.height = ch.h + "px";
      var key = self.wb.eng.ver + ":" + JSON.stringify(ch) + ":" + self.dark + ":" + (sh.hiddenVer || 0);
      if (box.key !== key) { box.key = key; SH.chart.draw(box.cv, self.wb, ch, { bg: T.bg, fg: T.text, muted: T.muted, grid: T.chartGrid, axis: T.axis, font: FONT }, ch.w - 2, ch.h - 2); }
    });
    this.chartBoxes.forEach(function (b, id) { if (!seen.has(id) || b.uid !== sh.uid) { b.el.remove(); self.chartBoxes.delete(id); } });
  };
  G.makeChartBox = function (ch) {
    var self = this, el = S3.h("div", { "class": "sh-chart", tabindex: "0", role: "img", "aria-label": "Chart: " + (ch.title || ch.type) }), cv = S3.h("canvas"), bar = S3.h("div", { "class": "sh-chartbar" });
    var edit = S3.h("button", { "class": "sh-cbtn", type: "button", title: "Edit chart", "aria-label": "Edit chart" }, [S3.h("svg", { "class": "i" }, [])]);
    edit.innerHTML = '<svg class="i"><use href="#s-pen"/></svg>';
    var png = S3.h("button", { "class": "sh-cbtn", type: "button", title: "Download PNG", "aria-label": "Download chart as PNG", html: '<svg class="i"><use href="#i-download"/></svg>' });
    var del = S3.h("button", { "class": "sh-cbtn", type: "button", title: "Delete chart", "aria-label": "Delete chart", html: '<svg class="i"><use href="#i-trash"/></svg>' });
    var grip = S3.h("div", { "class": "sh-cgrip", title: "Resize" });
    bar.appendChild(edit); bar.appendChild(png); bar.appendChild(del);
    el.appendChild(cv); el.appendChild(bar); el.appendChild(grip);
    this.chartLayer.appendChild(el);
    var box = { el: el, cv: cv, ch: ch, uid: this.sheet().uid, key: "" };
    this.chartBoxes.set(ch.id, box);
    edit.addEventListener("click", function (e) { e.stopPropagation(); if (self.cb.onChartEdit) self.cb.onChartEdit(box.ch); });
    del.addEventListener("click", function (e) { e.stopPropagation(); if (self.cb.onChartDelete) self.cb.onChartDelete(box.ch); });
    png.addEventListener("click", function (e) {
      e.stopPropagation();
      cv.toBlob(function (b) { if (b) { S3.download(b, (box.ch.title || "chart").replace(/[^\w؀-ۿ -]+/g, "").trim().slice(0, 60) + ".png", "image/png"); S3.done(); } });
    });
    function drag(e, resize) {
      if (e.button !== 0) return;
      e.preventDefault(); e.stopPropagation();
      var sx = e.clientX, sy = e.clientY, o = { x: box.ch.x, y: box.ch.y, w: box.ch.w, h: box.ch.h }, rtl = self.sheet().rtl, moved = false;
      el.setPointerCapture(e.pointerId);
      function mv(ev) {
        var dx = (ev.clientX - sx) * (rtl ? -1 : 1), dy = ev.clientY - sy;
        if (Math.abs(dx) + Math.abs(dy) > 2) moved = true;
        if (resize) { box.ch.w = Math.max(200, o.w + dx); box.ch.h = Math.max(140, o.h + dy); } else { box.ch.x = Math.max(0, o.x + dx); box.ch.y = Math.max(0, o.y + dy); }
        self.drawNow();
      }
      function up() {
        el.removeEventListener("pointermove", mv); el.removeEventListener("pointerup", up); el.removeEventListener("pointercancel", up);
        if (!moved) return;
        var n = { x: box.ch.x, y: box.ch.y, w: box.ch.w, h: box.ch.h };
        box.ch.x = o.x; box.ch.y = o.y; box.ch.w = o.w; box.ch.h = o.h;
        var sh = self.sheet(), id = box.ch.id;
        self.wb.setMeta(sh, resize ? "Resize chart" : "Move chart", function () { sh.charts = sh.charts.map(function (c) { return c.id === id ? Object.assign({}, c, n) : c; }); });
      }
      el.addEventListener("pointermove", mv); el.addEventListener("pointerup", up); el.addEventListener("pointercancel", up);
    }
    el.addEventListener("pointerdown", function (e) { if (e.target.closest(".sh-cbtn")) return; drag(e, e.target === grip); });
    el.addEventListener("dblclick", function () { if (self.cb.onChartEdit) self.cb.onChartEdit(box.ch); });
    el.addEventListener("keydown", function (e) { if (e.key === "Delete" && self.cb.onChartDelete) { e.preventDefault(); self.cb.onChartDelete(box.ch); } });
    return box;
  };

  // ---------- input wiring ----------
  G.wire = function () {
    var self = this, sc = this.scroller, ed = this.editor, drag = null, autoT = 0, lastPt = null;
    sc.addEventListener("pointerdown", function (e) {
      if (e.target.closest(".sh-chart") || e.target === ed || e.target.closest(".sh-ac")) return;
      self.touch = e.pointerType === "touch";
      var h = self.hit(e.clientX, e.clientY);
      if (!h.inside) return;
      if (e.button === 2) {
        if (h.zone === "cell" && !C.inRange(self.sel, h.r, h.c)) self.select(h.r, h.c, h.r, h.c, h.r, h.c, true);
        if (h.zone === "colhdr" && !(h.c >= self.sel.c0 && h.c <= self.sel.c1 && self.sel.r0 === 0 && self.sel.r1 >= MAXR - 1)) self.select(0, h.c, MAXR - 1, h.c, 0, h.c, true);
        if (h.zone === "rowhdr" && !(h.r >= self.sel.r0 && h.r <= self.sel.r1 && self.sel.c0 === 0 && self.sel.c1 >= MAXC - 1)) self.select(h.r, 0, h.r, MAXC - 1, h.r, 0, true);
        return;
      }
      if (e.button !== 0) return;
      if (self.touch && h.zone === "cell" && !self.edit) { drag = { kind: "tap", h: h, x: e.clientX, y: e.clientY }; return; }
      e.preventDefault();
      if (h.colEdge !== undefined) { drag = { kind: "colsize", c: h.colEdge, x: e.clientX, w0: self.cw(h.colEdge), orig: self.sheet().colW.get(h.colEdge) }; }
      else if (h.rowEdge !== undefined) { drag = { kind: "rowsize", r: h.rowEdge, y: e.clientY, h0: self.rh(h.rowEdge), orig: self.sheet().rowH.get(h.rowEdge) }; }
      else if (h.zone === "corner") { if (self.edit) self.commit(0, 0); self.selectAll(); }
      else if (h.zone === "colhdr") { if (self.edit) self.commit(0, 0); var c0 = e.shiftKey ? self.sel.c : h.c; self.select(0, c0, MAXR - 1, h.c, 0, c0, true); drag = { kind: "cols", c0: c0 }; }
      else if (h.zone === "rowhdr") { if (self.edit) self.commit(0, 0); var r0 = e.shiftKey ? self.sel.r : h.r; self.select(r0, 0, h.r, MAXC - 1, r0, 0, true); drag = { kind: "rows", r0: r0 }; }
      else if (h.filterBtn !== undefined && !self.edit) { if (self.cb.onFilter) setTimeout(function () { self.cb.onFilter(h.filterBtn); }, 0); return; }
      else if (h.fillHandle) { drag = { kind: "fill", src: Object.assign({}, self.sel) }; }
      else if (self.edit && self.pointable()) { drag = { kind: "ref", r0: h.r, c0: h.c }; self.insertRef({ r0: h.r, c0: h.c, r1: h.r, c1: h.c }); }
      else {
        if (self.edit) self.commit(0, 0);
        if (e.shiftKey) self.select(self.sel.r, self.sel.c, h.r, h.c, self.sel.r, self.sel.c, true);
        else self.select(h.r, h.c, h.r, h.c, h.r, h.c, true);
        drag = { kind: "cells", r0: e.shiftKey ? self.sel.r : h.r, c0: e.shiftKey ? self.sel.c : h.c };
      }
      try { sc.setPointerCapture(e.pointerId); } catch (x) { /* ignore */ }
      if (!self.edit) self.focus();
    });
    function track(e) {
      if (!drag) {
        if (e.pointerType === "mouse") { var h0 = self.hit(e.clientX, e.clientY); sc.style.cursor = h0.colEdge !== undefined ? "col-resize" : h0.rowEdge !== undefined ? "row-resize" : h0.fillHandle ? "crosshair" : h0.filterBtn !== undefined ? "pointer" : h0.zone === "cell" ? "cell" : "default"; }
        return;
      }
      lastPt = { x: e.clientX, y: e.clientY };
      var h = self.hit(e.clientX, e.clientY), sh = self.sheet();
      if (drag.kind === "tap") { if (Math.abs(e.clientX - drag.x) + Math.abs(e.clientY - drag.y) > 8) drag = null; return; }
      if (drag.kind === "colsize") { var dx = (e.clientX - drag.x) * (sh.rtl ? -1 : 1); sh.colW.set(drag.c, Math.max(16, drag.w0 + dx)); self.axKey = ""; self.drawNow(); return; }
      if (drag.kind === "rowsize") { sh.rowH.set(drag.r, Math.max(10, drag.h0 + e.clientY - drag.y)); self.axKey = ""; self.drawNow(); return; }
      if (h.r < 0) h.r = 0; if (h.c < 0) h.c = 0;
      if (drag.kind === "cells") self.select(drag.r0, drag.c0, h.r, h.c, drag.r0, drag.c0, true);
      else if (drag.kind === "cols") self.select(0, drag.c0, MAXR - 1, h.c, 0, drag.c0, true);
      else if (drag.kind === "rows") self.select(drag.r0, 0, h.r, MAXC - 1, drag.r0, 0, true);
      else if (drag.kind === "ref") { self.insertRef({ r0: Math.min(drag.r0, h.r), c0: Math.min(drag.c0, h.c), r1: Math.max(drag.r0, h.r), c1: Math.max(drag.c0, h.c) }); }
      else if (drag.kind === "fill") {
        var s = drag.src, t = Object.assign({}, s);
        var down = h.r > s.r1 ? h.r - s.r1 : 0, up = h.r < s.r0 ? s.r0 - h.r : 0, right = h.c > s.c1 ? h.c - s.c1 : 0, left = h.c < s.c0 ? s.c0 - h.c : 0;
        if (Math.max(down, up) >= Math.max(right, left)) { if (down) t.r1 = h.r; else if (up) t.r0 = h.r; }
        else { if (right) t.c1 = h.c; else if (left) t.c0 = h.c; }
        self.fillTo = (t.r0 !== s.r0 || t.r1 !== s.r1 || t.c0 !== s.c0 || t.c1 !== s.c1) ? t : null;
        self.draw();
      }
      // auto-scroll near the edges while dragging
      if (!autoT && (drag.kind === "cells" || drag.kind === "fill" || drag.kind === "ref" || drag.kind === "rows" || drag.kind === "cols")) {
        autoT = setInterval(function () {
          if (!drag || !lastPt) { clearInterval(autoT); autoT = 0; return; }
          var rect = sc.getBoundingClientRect(), dy = lastPt.y > rect.bottom - 10 ? 24 : lastPt.y < rect.top + HH + 4 ? -24 : 0, dxp = lastPt.x > rect.right - 10 ? 40 : lastPt.x < rect.left + HW + 4 ? -40 : 0;
          if (!dy && !dxp) return;
          if (dy && drag.kind !== "cols") sc.scrollTop += dy;
          if (dxp && drag.kind !== "rows") sc.scrollLeft += self.sheet().rtl ? -dxp : dxp;
          track({ clientX: lastPt.x, clientY: lastPt.y, pointerType: "mouse" });
        }, 50);
      }
    }
    sc.addEventListener("pointermove", track);
    function end(e) {
      if (autoT) { clearInterval(autoT); autoT = 0; }
      if (!drag) return;
      var d = drag, sh = self.sheet(); drag = null;
      if (d.kind === "tap") { var h = d.h; if (h.filterBtn !== undefined) { if (self.cb.onFilter) self.cb.onFilter(h.filterBtn); return; } if (h.fillHandle) return; self.select(h.r, h.c, h.r, h.c, h.r, h.c, true); return; }
      if (d.kind === "colsize" || d.kind === "rowsize") {
        var m = d.kind === "colsize" ? sh.colW : sh.rowH, idx = d.kind === "colsize" ? d.c : d.r, now = m.get(idx);
        if (d.orig === undefined) m.delete(idx); else m.set(idx, d.orig);
        var list = self.sel.r0 === 0 && self.sel.r1 >= MAXR - 1 && d.kind === "colsize" && idx >= self.sel.c0 && idx <= self.sel.c1 ? range(self.sel.c0, self.sel.c1) : self.sel.c0 === 0 && self.sel.c1 >= MAXC - 1 && d.kind === "rowsize" && idx >= self.sel.r0 && idx <= self.sel.r1 ? range(self.sel.r0, self.sel.r1) : [idx];
        if (now !== d.orig) self.wb.resize(sh, d.kind === "colsize" ? "c" : "r", list, now);
        self.axKey = ""; self.draw(); return;
      }
      if (d.kind === "fill" && self.fillTo) {
        var to = self.fillTo; self.fillTo = null;
        try { self.wb.fill(sh, d.src, to); } catch (err) { S3.toast(err.message, "error"); }
        self.select(to.r0, to.c0, to.r1, to.c1, d.src.r, d.src.c, true);
      }
      self.fillTo = null; self.draw();
      if (self.edit) self.editor.focus({ preventScroll: true });
    }
    sc.addEventListener("pointerup", end); sc.addEventListener("pointercancel", end);
    sc.addEventListener("dblclick", function (e) {
      if (e.target.closest(".sh-chart")) return;
      var h = self.hit(e.clientX, e.clientY);
      if (h.colEdge !== undefined) { self.autofit(h.colEdge); return; }
      if (h.rowEdge !== undefined) { self.wb.resize(self.sheet(), "r", [h.rowEdge], null); return; }
      if (h.zone === "cell" && !self.edit && h.filterBtn === undefined) { self.select(h.r, h.c, h.r, h.c); self.touch = false; self.beginEdit("edit"); }
    });
    sc.addEventListener("contextmenu", function (e) { if (e.target.closest(".sh-chart") || e.target === ed) return; e.preventDefault(); var h = self.hit(e.clientX, e.clientY); if (self.cb.onContext) self.cb.onContext(e.clientX, e.clientY, h); });
    var onKey = function (e) { self.key(e); };
    ed.addEventListener("keydown", onKey);
    sc.addEventListener("keydown", function (e) { if (e.target === sc) onKey(e); });
    ed.addEventListener("input", function () { if (!self.edit) { var t = ed.value; if (t) { self.beginEdit("enter", t); } return; } self.onInput(); });
    ed.addEventListener("compositionstart", function () { if (!self.edit) self.beginEdit("enter", ""); });
    ed.addEventListener("click", function () { if (self.edit) { self.edit.mode = "edit"; self.autocomplete(); } });
    ed.addEventListener("keyup", function (e) { if (self.edit && /^Arrow|Home|End/.test(e.key)) self.autocomplete(); });
    document.addEventListener("copy", function (e) { if (self.ownsFocus() && !self.edit) self.copy(e, false); });
    document.addEventListener("cut", function (e) { if (self.ownsFocus() && !self.edit) self.copy(e, true); });
    document.addEventListener("paste", function (e) {
      if (!self.ownsFocus() || self.edit) return;
      var t = e.clipboardData && e.clipboardData.getData("text/plain");
      e.preventDefault();
      if (t) self.paste(t);
    });
  };
  function range(a, b) { var o = []; for (var i = a; i <= b; i++) o.push(i); return o; }
  G.ownsFocus = function () { var a = document.activeElement; return a === this.editor || a === this.scroller; };
  G.autofit = function (c) {
    var sh = this.sheet(), max = 30, self = this, n = 0, cols = this.sel.r0 === 0 && this.sel.r1 >= MAXR - 1 && c >= this.sel.c0 && c <= this.sel.c1 ? range(this.sel.c0, this.sel.c1) : [c];
    var widths = new Map();
    cols.forEach(function (cc) {
      max = 30; n = 0;
      for (var r = 0; r <= Math.min(Math.max(sh.maxR, sh.smaxR || 0), 20000); r++) {
        var x = sh.cells.get(r * KW + cc); if (!x && !self.wb.eng.spill.size) continue;
        var t = self.wb.text(sh, r, cc); if (!t) continue;
        max = Math.max(max, self.measure(fontOf(x && x.s), t.split("\n")[0]) + 14); if (++n > 5000) break;
      }
      widths.set(cc, Math.min(600, Math.ceil(max)));
    });
    this.wb.run("Autofit", function () { self.wb.meta(sh); widths.forEach(function (w, cc) { sh.colW.set(cc, w); }); });
  };
  G.key = function (e) {
    var k = e.key, ctrl = e.ctrlKey || e.metaKey, sh = this.sheet(), s = this.sel;
    if (e.isComposing) return;
    if (this.edit) {
      if (this.acItems && !this.ac.hidden) {
        if (k === "ArrowDown") { e.preventDefault(); this.acMove(1); return; }
        if (k === "ArrowUp") { e.preventDefault(); this.acMove(-1); return; }
        if (k === "Tab" || (k === "Enter" && !e.altKey)) { e.preventDefault(); this.acPick(); return; }
        if (k === "Escape") { e.preventDefault(); this.ac.hidden = true; this.acItems = null; return; }
      }
      if (k === "Enter" && e.altKey) { e.preventDefault(); var ta = this.editor, p = ta.selectionStart; ta.value = ta.value.slice(0, p) + "\n" + ta.value.slice(ta.selectionEnd); ta.setSelectionRange(p + 1, p + 1); this.onInput(); return; }
      if (k === "Enter") { e.preventDefault(); if (ctrl) this.commit(0, 0, true); else this.commit(e.shiftKey ? -1 : 1, 0); return; }
      if (k === "Tab") { e.preventDefault(); this.commit(0, e.shiftKey ? -1 : 1); return; }
      if (k === "Escape") { e.preventDefault(); this.cancel(); return; }
      if (k === "F2") { e.preventDefault(); this.edit.mode = this.edit.mode === "edit" ? "enter" : "edit"; return; }
      if (this.edit.mode === "enter" && /^Arrow/.test(k) && !this.pointable()) {
        e.preventDefault();
        var d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
        if (sh.rtl) d = [d[0], -d[1]];
        this.commit(0, 0); this.move(d[0], d[1], false, false); return;
      }
      return;
    }
    var dirs = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (dirs[k]) { e.preventDefault(); var dd = dirs[k]; if (sh.rtl) dd = [dd[0], -dd[1]]; this.move(dd[0], dd[1], e.shiftKey, ctrl); return; }
    if (k === "Enter") { e.preventDefault(); this.advance(e.shiftKey ? -1 : 1, 0); return; }
    if (k === "Tab") { e.preventDefault(); this.advance(0, e.shiftKey ? -1 : 1); return; }
    if (k === "Home") { e.preventDefault(); if (ctrl) this.select(sh.freeze.r, sh.freeze.c, sh.freeze.r, sh.freeze.c); else this.select(s.r, 0, s.r, 0); return; }
    if (k === "End" && ctrl) { e.preventDefault(); var er = Math.max(0, sh.maxR), ec = Math.max(0, sh.maxC); this.select(er, ec, er, ec); return; }
    if (k === "PageDown" || k === "PageUp") {
      e.preventDefault(); var rows = Math.max(1, Math.floor((this.scroller.clientHeight - HH) / DH) - 1) * (k === "PageDown" ? 1 : -1);
      var nr = C.clamp(s.r + rows, 0, MAXR - 1); this.scroller.scrollTop += rows * DH; this.select(nr, s.c, nr, s.c); return;
    }
    if (k === "F2") { e.preventDefault(); this.beginEdit("edit"); return; }
    if (k === "Delete" || k === "Backspace") { e.preventDefault(); this.wb.clear(sh, s, "contents"); return; }
    if (k === "Escape") { if (this.clip) { this.clip = null; this.draw(); } return; }
    if (this.cb.onKey && this.cb.onKey(e)) { e.preventDefault(); return; }
  };

  SH.Grid = Grid; SH.GRID = { HW: HW, HH: HH, DW: DW, DH: DH, THEMES: THEMES, FONT: FONT };
})();
