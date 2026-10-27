/* Sheets model: workbook / sheets / cells (immutable cell objects in a sparse Map), undo/redo,
   insert/delete rows and columns (formulas everywhere are rewritten), sort, filter, merge, fill series,
   copy/paste blocks, find/replace and (de)serialisation. Pure logic, no DOM. */
(function () {
  "use strict";
  var SH = window.SH, C = SH.core, F = SH.fmt, EN = SH.engine, KW = C.KW, MAXR = C.MAXR, MAXC = C.MAXC;
  var DEF_ROWS = 1000, DEF_COLS = 26, MAX_HIST = 200;

  function Sheet(uid, name) {
    this.uid = uid; this.name = name; this.cells = new Map(); this.colW = new Map(); this.rowH = new Map();
    this.merges = []; this.freeze = { r: 0, c: 0 }; this.rtl = false; this.cf = []; this.charts = []; this.filter = null;
    this.hidden = new Set(); this.maxR = -1; this.maxC = -1; this.nrows = DEF_ROWS; this.ncols = DEF_COLS; this.grid = true;
    this.smaxR = -1; this.smaxC = -1;
  }
  function metaOf(sh) {
    return C.clone({ name: sh.name, colW: Array.from(sh.colW), rowH: Array.from(sh.rowH), merges: sh.merges, freeze: sh.freeze, rtl: sh.rtl,
      cf: sh.cf, charts: sh.charts, filter: sh.filter, nrows: sh.nrows, ncols: sh.ncols, grid: sh.grid });
  }
  function applyMeta(sh, m) {
    m = C.clone(m);
    sh.name = m.name; sh.colW = new Map(m.colW); sh.rowH = new Map(m.rowH); sh.merges = m.merges; sh.freeze = m.freeze; sh.rtl = m.rtl;
    sh.cf = m.cf; sh.charts = m.charts; sh.filter = m.filter; sh.nrows = m.nrows; sh.ncols = m.ncols; sh.grid = m.grid;
  }
  function extents(sh) {
    var mr = -1, mc = -1;
    sh.cells.forEach(function (cell, k) { var r = Math.floor(k / KW), c = k % KW; if (r > mr) mr = r; if (c > mc) mc = c; });
    sh.maxR = mr; sh.maxC = mc;
    sh.nrows = Math.max(sh.nrows, Math.min(MAXR, mr + 101)); sh.ncols = Math.max(sh.ncols, Math.min(MAXC, mc + 6));
  }

  function Workbook(name) {
    this.id = C.uid("wb"); this.name = name || "Untitled spreadsheet"; this.sheets = []; this.active = 0; this.seq = 1;
    this.byUid = new Map(); this.eng = new EN.Engine(this); this.undoStack = []; this.redoStack = []; this.tx = null; this.depth = 0;
    this.listeners = []; this.created = Date.now();
  }
  var W = Workbook.prototype;
  W.on = function (fn) { this.listeners.push(fn); };
  W.emit = function (info) { this.listeners.forEach(function (f) { try { f(info); } catch (e) { /* listener errors must not break edits */ } }); };
  W.sheet = function () { return this.sheets[this.active]; };
  W.sheetByName = function (n) { n = String(n).toLowerCase(); for (var i = 0; i < this.sheets.length; i++) if (this.sheets[i].name.toLowerCase() === n) return this.sheets[i]; return null; };
  W.sheetByUid = function (u) { return this.byUid.get(u) || null; };
  W.uniqueName = function (base) { var n = base, i = 1; while (this.sheetByName(n)) n = base + " " + (++i); return n; };
  W.addSheetRaw = function (name, at) {
    var sh = new Sheet(this.seq++, name);
    this.byUid.set(sh.uid, sh);
    if (at === undefined) this.sheets.push(sh); else this.sheets.splice(at, 0, sh);
    return sh;
  };
  W.cell = function (sh, r, c) { return sh.cells.get(r * KW + c); };
  W.value = function (sh, r, c) { return this.eng.value(sh, r, c); };
  W.style = function (sh, r, c) { var x = sh.cells.get(r * KW + c); return x && x.s; };
  W.text = function (sh, r, c) { var x = sh.cells.get(r * KW + c); return F.format(this.eng.value(sh, r, c), x && x.s && x.s.nf).t; };

  // ---------- transactions & history ----------
  W.run = function (label, fn) {
    var top = !this.tx;
    if (top) this.tx = { label: label, cells: new Map(), metas: new Map(), changed: [] };
    this.depth++;
    try { fn(); } finally { this.depth--; }
    if (top) { var tx = this.tx; this.tx = null; this.commit(tx); }
  };
  W.set = function (sh, r, c, cell) {
    var lk = r * KW + c, gk = sh.uid * EN.SW + lk, tx = this.tx;
    if (!tx) throw new Error("set() outside run()");
    if (!tx.cells.has(gk)) tx.cells.set(gk, { uid: sh.uid, r: r, c: c, before: sh.cells.get(lk) });
    if (cell && (cell.v !== undefined || cell.f !== undefined || cell.s)) sh.cells.set(lk, cell); else sh.cells.delete(lk);
    if (cell) { if (r > sh.maxR) sh.maxR = r; if (c > sh.maxC) sh.maxC = c; if (r >= sh.nrows - 1) sh.nrows = Math.min(MAXR, r + 101); if (c >= sh.ncols - 1) sh.ncols = Math.min(MAXC, c + 6); }
    tx.changed.push({ sh: sh, r: r, c: c });
  };
  W.meta = function (sh) { if (this.tx && !this.tx.metas.has(sh.uid)) this.tx.metas.set(sh.uid, metaOf(sh)); };
  W.commit = function (tx) {
    var self = this, entry = { label: tx.label, cells: [], metas: [] };
    tx.cells.forEach(function (p) { var sh = self.byUid.get(p.uid); entry.cells.push({ uid: p.uid, r: p.r, c: p.c, before: p.before, after: sh.cells.get(p.r * KW + p.c) }); });
    tx.metas.forEach(function (m, uid) { var sh = self.byUid.get(uid); if (sh) entry.metas.push({ uid: uid, before: m, after: metaOf(sh) }); });
    if (!entry.cells.length && !entry.metas.length) return;
    this.push(entry);
    if (tx.changed.length) this.eng.update(dedupe(tx.changed));
    this.refreshFilters();
    this.emit({ cells: tx.changed, label: tx.label });
  };
  function dedupe(list) { var seen = new Set(), out = []; list.forEach(function (x) { var k = x.sh.uid * EN.SW + x.r * KW + x.c; if (!seen.has(k)) { seen.add(k); out.push(x); } }); return out; }
  W.push = function (entry) {
    this.undoStack.push(entry); this.redoStack = [];
    var total = 0;
    for (var i = this.undoStack.length - 1; i >= 0; i--) { var e = this.undoStack[i]; total += e.size || e.cells.length || 1; if (total > 3e6 || this.undoStack.length - i > MAX_HIST) { this.undoStack.splice(0, i + 1); break; } }
  };
  // Structural change (rows/cols/sheets/sort…): whole-workbook snapshot (cell objects are shared, never mutated)
  W.snap = function () {
    return { active: this.active, seq: this.seq, sheets: this.sheets.map(function (sh) { return { uid: sh.uid, cells: new Map(sh.cells), meta: metaOf(sh) }; }) };
  };
  W.restore = function (s) {
    var self = this, old = this.byUid;
    this.byUid = new Map(); this.sheets = s.sheets.map(function (x) {
      var sh = old.get(x.uid) || new Sheet(x.uid, x.meta.name);
      sh.cells = new Map(x.cells); applyMeta(sh, x.meta); extents(sh); self.byUid.set(sh.uid, sh); return sh;
    });
    this.active = Math.min(s.active, this.sheets.length - 1); this.seq = Math.max(this.seq, s.seq);
    this.eng.rebuild(); this.refreshFilters();
  };
  W.runStruct = function (label, fn) {
    if (this.tx) throw new Error("runStruct inside run");
    var before = this.snap();
    fn();
    this.sheets.forEach(extents);
    var after = this.snap(), size = 0;
    after.sheets.forEach(function (x) { size += x.cells.size; });
    this.push({ label: label, struct: true, before: before, after: after, size: size * 2 + 1 });
    this.eng.rebuild(); this.refreshFilters();
    this.emit({ struct: true, label: label });
  };
  W.canUndo = function () { return this.undoStack.length > 0; };
  W.canRedo = function () { return this.redoStack.length > 0; };
  W.undo = function () { var e = this.undoStack.pop(); if (!e) return null; this.apply(e, true); this.redoStack.push(e); return e.label; };
  W.redo = function () { var e = this.redoStack.pop(); if (!e) return null; this.apply(e, false); this.undoStack.push(e); return e.label; };
  W.apply = function (e, back) {
    var self = this;
    if (e.struct) { this.restore(back ? e.before : e.after); this.emit({ struct: true, undo: true }); return; }
    var changed = [];
    e.cells.forEach(function (p) {
      var sh = self.byUid.get(p.uid); if (!sh) return;
      var v = back ? p.before : p.after, lk = p.r * KW + p.c;
      if (v) { sh.cells.set(lk, v); if (p.r > sh.maxR) sh.maxR = p.r; if (p.c > sh.maxC) sh.maxC = p.c; } else sh.cells.delete(lk);
      changed.push({ sh: sh, r: p.r, c: p.c });
    });
    e.metas.forEach(function (m) { var sh = self.byUid.get(m.uid); if (sh) applyMeta(sh, back ? m.before : m.after); });
    if (changed.length) this.eng.update(changed);
    this.refreshFilters();
    this.emit({ cells: changed, undo: true, meta: e.metas.length > 0 });
  };

  // ---------- cell helpers ----------
  function mk(v, f, s, q) { var o = {}; if (f !== undefined) o.f = f; else if (v !== undefined && v !== null) o.v = v; if (s) o.s = s; if (q) o.q = 1; return o; }
  function withStyle(cell, s) { var o = mk(cell && cell.v, cell && cell.f, s, cell && cell.q); return o; }
  function mergeStyle(s, patch) {
    var o = {}, k;
    if (s) for (k in s) o[k] = s[k];
    for (k in patch) { if (patch[k] === null || patch[k] === undefined || patch[k] === false) delete o[k]; else o[k] = patch[k]; }
    for (k in o) return o;
    return undefined;
  }
  // put typed text into a cell (keeps its formatting; adopts a format the text implies, e.g. "12%")
  W.setInput = function (sh, r, c, text) {
    var old = this.cell(sh, r, c), s = old && old.s, p = F.parseInput(text);
    if (p.f !== undefined) { this.set(sh, r, c, mk(undefined, EN.normalize(p.f), s)); return; }
    if (p.nf && (!s || !s.nf || s.nf === "General") && typeof p.v === "number") s = mergeStyle(s, { nf: p.nf });
    if (s && s.nf === "@" && p.v !== null && typeof p.v !== "string") p = { v: String(text), text: true };
    this.set(sh, r, c, mk(p.v, undefined, s, p.text && typeof p.v === "string" && looksSpecial(p.v)));
  };
  function looksSpecial(t) { var p = F.parseInput(t); return p.f !== undefined || typeof p.v !== "string"; }
  // the text shown in the formula bar / editor for a cell
  W.inputText = function (sh, r, c) {
    var x = this.cell(sh, r, c);
    if (!x) return "";
    if (x.f !== undefined) return "=" + x.f;
    var v = x.v, nf = x.s && x.s.nf;
    if (v === undefined || v === null) return "";
    if (typeof v === "string") return x.q ? "'" + v : v;
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (v instanceof C.Err) return v.e;
    if (nf && F.isDateFmt(nf)) {
      var d = F.ymd(v), t = F.hms(v), ds = F.isTimeOnly(nf) ? "" : d.y + "-" + pad2(d.m) + "-" + pad2(d.d);
      var ts = (t.h || t.mi || t.s || F.isTimeOnly(nf)) ? t.h + ":" + pad2(t.mi) + (t.s ? ":" + pad2(t.s) : "") : "";
      return (ds + (ds && ts ? " " : "") + ts) || "0";
    }
    if (nf && /%/.test(nf.replace(/"[^"]*"/g, ""))) return String(C.r15(v * 100)) + "%";
    return String(C.r15(v));
  };
  function pad2(n) { return n < 10 ? "0" + n : String(n); }
  // clamp a selection to the area worth touching (whole-column selections stop at the data)
  W.clampArea = function (sh, g) {
    var r1 = Math.min(g.r1, Math.max(sh.maxR, Math.min(sh.nrows - 1, 999))), c1 = Math.min(g.c1, Math.max(sh.maxC, Math.min(sh.ncols - 1, 25)));
    return { r0: g.r0, c0: g.c0, r1: Math.max(g.r0, r1), c1: Math.max(g.c0, c1) };
  };
  W.forCells = function (sh, g, fn) {   // existing cells only, row-major
    var a = this.clampArea(sh, g), area = (a.r1 - a.r0 + 1) * (a.c1 - a.c0 + 1);
    if (area > sh.cells.size * 4) {
      var list = [];
      sh.cells.forEach(function (cell, k) { var r = Math.floor(k / KW), c = k % KW; if (r >= a.r0 && r <= a.r1 && c >= a.c0 && c <= a.c1) list.push(k); });
      list.sort(function (x, y) { return x - y; }).forEach(function (k) { fn(Math.floor(k / KW), k % KW, sh.cells.get(k)); });
      return;
    }
    for (var r = a.r0; r <= a.r1; r++) for (var c = a.c0; c <= a.c1; c++) { var x = sh.cells.get(r * KW + c); if (x) fn(r, c, x); }
  };
  W.styleRange = function (sh, g, patch) {
    var self = this, a = this.clampArea(sh, g);
    if ((a.r1 - a.r0 + 1) * (a.c1 - a.c0 + 1) > 600000) throw new Error("That selection is too large to format at once.");
    this.run("Format", function () {
      for (var r = a.r0; r <= a.r1; r++) for (var c = a.c0; c <= a.c1; c++) {
        var x = self.cell(sh, r, c), s = mergeStyle(x && x.s, typeof patch === "function" ? patch(x && x.s, r, c) : patch);
        if (!x && !s) continue;
        self.set(sh, r, c, withStyle(x, s));
      }
    });
  };
  W.borders = function (sh, g, mode, color) {
    var a = this.clampArea(sh, g), col = color || "#000000", line = { s: "thin", c: col };
    this.styleRange(sh, a, function (s, r, c) {
      var bd = s && s.bd ? C.clone(s.bd) : {};
      if (mode === "none") return { bd: null };
      if (mode === "all") bd = { t: line, r: line, b: line, l: line };
      if (mode === "outer" || mode === "top") if (r === a.r0) bd.t = line;
      if (mode === "outer" || mode === "bottom") if (r === a.r1) bd.b = line;
      if (mode === "outer" || mode === "left") if (c === a.c0) bd.l = line;
      if (mode === "outer" || mode === "right") if (c === a.c1) bd.r = line;
      if (mode === "thickbottom" && r === a.r1) bd.b = { s: "medium", c: col };
      for (var k in bd) return { bd: bd };
      return { bd: null };
    });
  };
  W.clear = function (sh, g, what) {      // what: "contents" | "formats" | "all"
    var self = this;
    this.run(what === "formats" ? "Clear formats" : "Clear", function () {
      self.forCells(sh, g, function (r, c, x) {
        if (what === "all") self.set(sh, r, c, null);
        else if (what === "formats") self.set(sh, r, c, x.s ? mk(x.v, x.f, undefined, x.q) : x);
        else if (x.v !== undefined || x.f !== undefined) self.set(sh, r, c, x.s ? { s: x.s } : null);
      });
    });
  };

  // ---------- sheets ----------
  W.newSheet = function (name, at) {
    var self = this, sh;
    this.runStruct("Add sheet", function () { sh = self.addSheetRaw(self.uniqueName(name || "Sheet" + (self.sheets.length + 1)), at); self.active = self.sheets.indexOf(sh); });
    return sh;
  };
  W.renameSheet = function (sh, name) {
    name = String(name).trim().replace(/[\\/?*[\]:]/g, "").slice(0, 31);
    if (!name || (this.sheetByName(name) && this.sheetByName(name) !== sh)) return false;
    var old = sh.name, self = this;
    if (old === name) return true;
    this.runStruct("Rename sheet", function () {
      self.sheets.forEach(function (s) { rewriteAll(s, function (f) { return EN.renameSheetIn(f, old, name); }); });
      self.sheets.forEach(function (s) { s.charts.forEach(function (ch) { ch.range = EN.renameSheetIn(ch.range, old, name); }); });
      sh.name = name;
    });
    return true;
  };
  W.deleteSheet = function (sh) {
    if (this.sheets.length < 2) return false;
    var self = this, i = this.sheets.indexOf(sh);
    this.runStruct("Delete sheet", function () {
      self.sheets.splice(i, 1); self.byUid.delete(sh.uid);
      self.sheets.forEach(function (s) { rewriteAll(s, function (f) { return EN.dropSheetIn(f, sh.name); }); });
      self.active = Math.max(0, Math.min(self.active, self.sheets.length - 1));
    });
    return true;
  };
  W.duplicateSheet = function (sh) {
    var self = this, copy;
    this.runStruct("Duplicate sheet", function () {
      copy = self.addSheetRaw(self.uniqueName(sh.name + " copy"), self.sheets.indexOf(sh) + 1);
      copy.cells = new Map(sh.cells); applyMeta(copy, metaOf(sh)); copy.name = self.uniqueName(sh.name + " copy");
      copy.charts.forEach(function (ch) { ch.id = C.uid("ch"); ch.range = retarget(ch.range, sh.name, copy.name); });
      self.active = self.sheets.indexOf(copy);
    });
    return copy;
  };
  function retarget(range, from, to) { return EN.renameSheetIn(range, from, to); }
  W.moveSheet = function (sh, delta) {
    var self = this, i = this.sheets.indexOf(sh), j = i + delta;
    if (j < 0 || j >= this.sheets.length) return;
    this.runStruct("Move sheet", function () { self.sheets.splice(i, 1); self.sheets.splice(j, 0, sh); self.active = j; });
  };
  function rewriteAll(sh, fn) {
    sh.cells.forEach(function (x, k) { if (x.f === undefined) return; var g = fn(x.f); if (g !== x.f) sh.cells.set(k, mk(undefined, g, x.s)); });
  }

  // ---------- rows & columns ----------
  function shiftMap(m, at, n) {
    var out = new Map();
    m.forEach(function (v, k) { if (n < 0 && k >= at && k < at - n) return; var p = k >= at ? k + n : k; if (p >= 0) out.set(p, v); });
    return out;
  }
  function shiftRange(g, axis, at, n) {
    var p0 = axis === "r" ? g.r0 : g.c0, p1 = axis === "r" ? g.r1 : g.c1;
    if (n > 0) { if (p0 >= at) p0 += n; if (p1 >= at) p1 += n; }
    else { var d0 = at, d1 = at - n - 1; if (p0 >= d0 && p1 <= d1) return null; p0 = p0 < d0 ? p0 : p0 > d1 ? p0 + n : d0; p1 = p1 < d0 ? p1 : p1 > d1 ? p1 + n : d0 - 1; }
    var o = { r0: g.r0, c0: g.c0, r1: g.r1, c1: g.c1 };
    if (axis === "r") { o.r0 = p0; o.r1 = p1; } else { o.c0 = p0; o.c1 = p1; }
    return o;
  }
  W.insertDelete = function (sh, axis, at, n) {   // n > 0 insert, n < 0 delete
    var self = this, max = axis === "r" ? MAXR : MAXC;
    this.runStruct((n > 0 ? "Insert " : "Delete ") + (axis === "r" ? "rows" : "columns"), function () {
      var cells = new Map();
      sh.cells.forEach(function (x, k) {
        var r = Math.floor(k / KW), c = k % KW, p = axis === "r" ? r : c;
        if (n < 0 && p >= at && p < at - n) return;
        if (p >= at) p += n;
        if (p < 0 || p >= max) return;
        cells.set(axis === "r" ? p * KW + c : r * KW + p, x);
      });
      sh.cells = cells;
      self.sheets.forEach(function (s) { rewriteAll(s, function (f) { return EN.shiftFormula(f, s.name, sh.name, axis, at, n); }); });
      if (axis === "r") sh.rowH = shiftMap(sh.rowH, at, n); else sh.colW = shiftMap(sh.colW, at, n);
      sh.merges = sh.merges.map(function (g) { return shiftRange(g, axis, at, n); }).filter(function (g) { return g && (g.r1 > g.r0 || g.c1 > g.c0); });
      sh.cf.forEach(function (rule) { rule.range = shiftRange(rule.range, axis, at, n) || rule.range; });
      sh.cf = sh.cf.filter(function (rule) { return rule.range; });
      if (sh.filter) { var fg = shiftRange(sh.filter, axis, at, n); if (fg) { var cols = {}; for (var k in sh.filter.cols) { var cc = +k; if (axis === "c") { if (n < 0 && cc >= at && cc < at - n) continue; if (cc >= at) cc += n; } cols[cc] = sh.filter.cols[k]; } fg.cols = cols; } sh.filter = fg; }
      self.sheets.forEach(function (s) { s.charts.forEach(function (ch) { ch.range = EN.shiftFormula(ch.range, s.name, sh.name, axis, at, n); }); });
      if (axis === "r") sh.nrows = Math.min(MAXR, Math.max(1, sh.nrows + n)); else sh.ncols = Math.min(MAXC, Math.max(1, sh.ncols + n));
    });
  };
  W.resize = function (sh, axis, list, px) {
    this.run("Resize", function () {
      this.meta(sh);
      var m = axis === "r" ? sh.rowH : sh.colW;
      list.forEach(function (i) { if (px === null) m.delete(i); else m.set(i, Math.max(axis === "r" ? 8 : 12, Math.min(2000, Math.round(px)))); });
    }.bind(this));
  };
  W.setMeta = function (sh, label, fn) { var self = this; this.run(label, function () { self.meta(sh); fn(sh); }); };

  // ---------- merge ----------
  W.mergeAt = function (sh, r, c) { for (var i = 0; i < sh.merges.length; i++) { var g = sh.merges[i]; if (r >= g.r0 && r <= g.r1 && c >= g.c0 && c <= g.c1) return g; } return null; };
  W.merge = function (sh, g) {
    var self = this;
    if (g.r0 === g.r1 && g.c0 === g.c1) return;
    this.run("Merge", function () {
      self.meta(sh);
      sh.merges = sh.merges.filter(function (m) { return !C.overlaps(m, g); });
      sh.merges.push({ r0: g.r0, c0: g.c0, r1: g.r1, c1: g.c1 });
      self.forCells(sh, g, function (r, c, x) { if ((r !== g.r0 || c !== g.c0) && (x.v !== undefined || x.f !== undefined)) self.set(sh, r, c, x.s ? { s: x.s } : null); });
    });
  };
  W.unmerge = function (sh, g) { var self = this; this.run("Unmerge", function () { self.meta(sh); sh.merges = sh.merges.filter(function (m) { return !C.overlaps(m, g); }); }); };

  // ---------- copy / paste ----------
  W.copyBlock = function (sh, g) {
    var a = this.clampArea(sh, g), rows = [];
    for (var r = a.r0; r <= a.r1; r++) { var row = []; for (var c = a.c0; c <= a.c1; c++) row.push(sh.cells.get(r * KW + c) || null); rows.push(row); }
    var merges = sh.merges.filter(function (m) { return m.r0 >= a.r0 && m.r1 <= a.r1 && m.c0 >= a.c0 && m.c1 <= a.c1; }).map(function (m) { return { r0: m.r0 - a.r0, c0: m.c0 - a.c0, r1: m.r1 - a.r0, c1: m.c1 - a.c0 }; });
    return { r0: a.r0, c0: a.c0, rows: rows, merges: merges, uid: sh.uid };
  };
  // paste a block at (r, c); fills a larger selection by tiling. mode: "all" | "values" | "formats"
  W.pasteBlock = function (sh, r, c, block, mode, target) {
    var self = this, h = block.rows.length, w = block.rows[0].length, th = h, tw = w;
    if (target && (target.r1 - target.r0 + 1) % h === 0 && (target.c1 - target.c0 + 1) % w === 0) { th = target.r1 - target.r0 + 1; tw = target.c1 - target.c0 + 1; }
    if (r + th > MAXR || c + tw > MAXC) throw new Error("The paste does not fit on the sheet.");
    this.run("Paste", function () {
      for (var i = 0; i < th; i++) for (var j = 0; j < tw; j++) {
        var src = block.rows[i % h][j % w], old = self.cell(sh, r + i, c + j), rr = r + i, cc = c + j;
        var dr = rr - (block.r0 + i % h), dc = cc - (block.c0 + j % w);
        if (mode === "formats") { self.set(sh, rr, cc, withStyle(old, src && src.s)); continue; }
        if (mode === "values") {
          var bsh = self.byUid.get(block.uid) || sh, v = src ? self.eng.value(bsh, block.r0 + i % h, block.c0 + j % w) : null;
          if (src && src.f === undefined) v = src.v;
          self.set(sh, rr, cc, mk(v instanceof C.Err ? v : v, undefined, old && old.s)); continue;
        }
        if (!src) { self.set(sh, rr, cc, null); continue; }
        self.set(sh, rr, cc, src.f !== undefined ? mk(undefined, EN.moveFormula(src.f, dr, dc), src.s) : src);
      }
      if (mode !== "values" && block.merges.length) {
        self.meta(sh);
        var area = { r0: r, c0: c, r1: r + th - 1, c1: c + tw - 1 };
        sh.merges = sh.merges.filter(function (m) { return !C.overlaps(m, area); });
        block.merges.forEach(function (m) { sh.merges.push({ r0: m.r0 + r, c0: m.c0 + c, r1: m.r1 + r, c1: m.c1 + c }); });
      }
    });
    return { r0: r, c0: c, r1: r + th - 1, c1: c + tw - 1 };
  };
  // text grid (from the clipboard / another spreadsheet app) -> cells
  W.pasteText = function (sh, r, c, grid) {
    var self = this;
    if (!grid.length) return null;
    var w = Math.max.apply(null, grid.map(function (x) { return x.length; }));
    if (r + grid.length > MAXR || c + w > MAXC) throw new Error("The paste does not fit on the sheet.");
    this.run("Paste", function () { grid.forEach(function (row, i) { row.forEach(function (t, j) { self.setInput(sh, r + i, c + j, t); }); }); });
    return { r0: r, c0: c, r1: r + grid.length - 1, c1: c + w - 1 };
  };

  // ---------- fill (drag the handle) ----------
  var SERIES = [F.MONTHS, F.MONTHS.map(function (m) { return m.slice(0, 3); }), F.DAYS, F.DAYS.map(function (d) { return d.slice(0, 3); })];
  function seriesOf(t) {
    if (typeof t !== "string") return null;
    for (var i = 0; i < SERIES.length; i++) { var j = SERIES[i].map(function (x) { return x.toLowerCase(); }).indexOf(t.toLowerCase()); if (j >= 0) return { list: SERIES[i], idx: j, upper: t === t.toUpperCase(), lower: t === t.toLowerCase() }; }
    return null;
  }
  W.fill = function (sh, src, dst) {
    var self = this;
    var vertical = !(dst.r0 === src.r0 && dst.r1 === src.r1);
    var lines = vertical ? range(src.c0, src.c1) : range(src.r0, src.r1);
    var before = vertical ? dst.r0 < src.r0 : dst.c0 < src.c0;
    var n = vertical ? src.r1 - src.r0 + 1 : src.c1 - src.c0 + 1;
    var targets = vertical ? (before ? range(src.r0 - 1, dst.r0, -1) : range(src.r1 + 1, dst.r1)) : (before ? range(src.c0 - 1, dst.c0, -1) : range(src.c1 + 1, dst.c1));
    this.run("Fill", function () {
      lines.forEach(function (line) {
        var srcPos = vertical ? range(src.r0, src.r1) : range(src.c0, src.c1);
        if (before) srcPos.reverse();
        var cells = srcPos.map(function (p) { return vertical ? self.cell(sh, p, line) : self.cell(sh, line, p); });
        var vals = cells.map(function (x) { return x && x.f === undefined ? x.v : undefined; });
        var isNum = vals.every(function (v) { return typeof v === "number"; });
        var nf0 = cells[0] && cells[0].s && cells[0].s.nf, step = null;
        if (isNum && n >= 2) step = (vals[n - 1] - vals[0]) / (n - 1);
        else if (isNum && n === 1 && F.isDateFmt(nf0)) step = 1;
        var textNum = vals.map(function (v) { var m = typeof v === "string" && /^(.*?)(\d+)(\D*)$/.exec(v); return m ? { pre: m[1], n: +m[2], w: m[2].length, post: m[3] } : null; });
        var tStep = null;
        if (textNum.every(function (t) { return t && t.pre === textNum[0].pre && t.post === textNum[0].post; })) tStep = n >= 2 ? (textNum[n - 1].n - textNum[0].n) / (n - 1) : 1;
        var ser = vals.map(seriesOf), sStep = null;
        if (ser.every(function (s) { return s && s.list === ser[0].list; })) sStep = n >= 2 ? ser[1].idx - ser[0].idx : 1;
        targets.forEach(function (p, t) {
          var k = t % n, srcCell = cells[k], cell, sgn = before ? -1 : 1;
          if (step !== null && isNum) cell = mk(C.r15(vals[n - 1] + step * (t + 1)), undefined, srcCell && srcCell.s);
          else if (tStep !== null && Number.isInteger(tStep)) { var tn = textNum[n - 1], num = tn.n + tStep * (t + 1); var ns = String(Math.abs(num)); while (ns.length < tn.w) ns = "0" + ns; cell = mk(tn.pre + (num < 0 ? "-" : "") + ns + tn.post, undefined, srcCell && srcCell.s); }
          else if (sStep !== null) { var s0 = ser[n - 1], L = s0.list.length, ix = (((s0.idx + sStep * (t + 1)) % L) + L) % L, word = s0.list[ix]; cell = mk(s0.upper ? word.toUpperCase() : s0.lower ? word.toLowerCase() : word, undefined, srcCell && srcCell.s); }
          else if (srcCell && srcCell.f !== undefined) {
            var off = (Math.floor(t / n) + 1) * n * sgn;
            cell = mk(undefined, EN.moveFormula(srcCell.f, vertical ? off : 0, vertical ? 0 : off), srcCell.s);
          }
          else cell = srcCell || null;
          if (vertical) self.set(sh, p, line, cell); else self.set(sh, line, p, cell);
        });
      });
    });
  };
  function range(a, b, step) { var out = []; step = step || 1; if (step > 0) for (var i = a; i <= b; i++) out.push(i); else for (var j = a; j >= b; j--) out.push(j); return out; }

  // ---------- sort ----------
  function sortKey(v) { return v === null || v === undefined || v === "" ? 4 : v instanceof C.Err ? 3 : typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2; }
  W.sortRange = function (sh, g, keys, header) {
    var self = this, a = this.clampArea(sh, g), r0 = a.r0 + (header ? 1 : 0);
    if (r0 >= a.r1) return;
    var rows = [];
    for (var r = r0; r <= a.r1; r++) {
      var row = { r: r, cells: [], keys: keys.map(function (k) { return self.eng.value(sh, r, k.c); }) };
      for (var c = a.c0; c <= a.c1; c++) row.cells.push(sh.cells.get(r * KW + c));
      rows.push(row);
    }
    rows.sort(function (x, y) {
      for (var i = 0; i < keys.length; i++) {
        var p = x.keys[i], q = y.keys[i], kp = sortKey(p), kq = sortKey(q), d;
        if (kp === 4 || kq === 4) d = kp === kq ? 0 : kp === 4 ? 1 : -1;          // blanks always last
        else d = (kp !== kq ? kp - kq : kp === 0 ? p - q : kp === 1 ? p.localeCompare(q, undefined, { sensitivity: "base", numeric: true }) : 0) * (keys[i].asc ? 1 : -1);
        if (d) return d;
      }
      return x.r - y.r;
    });
    this.run("Sort", function () {
      rows.forEach(function (row, i) {
        var rr = r0 + i;
        row.cells.forEach(function (x, j) {
          var cell = x && x.f !== undefined ? mk(undefined, EN.moveFormula(x.f, rr - row.r, 0), x.s) : x || null;
          var cur = sh.cells.get(rr * KW + a.c0 + j) || null;
          if (cur !== cell) self.set(sh, rr, a.c0 + j, cell);
        });
      });
    });
  };
  // the block of non-empty cells around (r, c), like the "current region"
  W.region = function (sh, r, c) {
    var self = this, g = { r0: r, c0: c, r1: r, c1: c }, grew = true, guard = 0;
    function filled(rr, cc) { return rr >= 0 && cc >= 0 && self.eng.value(sh, rr, cc) !== null; }
    while (grew && guard++ < 2000) {
      grew = false;
      var i;
      for (i = g.c0; i <= g.c1 && !grew; i++) { if (filled(g.r0 - 1, i) || filled(g.r0 - 1, i - 1) || filled(g.r0 - 1, i + 1)) { g.r0--; grew = true; } }
      for (i = g.c0; i <= g.c1 && !grew; i++) { if (filled(g.r1 + 1, i) || filled(g.r1 + 1, i - 1) || filled(g.r1 + 1, i + 1)) { g.r1++; grew = true; } }
      for (i = g.r0; i <= g.r1 && !grew; i++) { if (filled(i, g.c0 - 1)) { g.c0--; grew = true; } }
      for (i = g.r0; i <= g.r1 && !grew; i++) { if (filled(i, g.c1 + 1)) { g.c1++; grew = true; } }
      if (g.r1 > sh.maxR + 1 || g.c1 > sh.maxC + 1) break;
    }
    return g;
  };

  // ---------- filter ----------
  W.setFilter = function (sh, g) {
    var self = this;
    this.run(g ? "Filter" : "Remove filter", function () { self.meta(sh); sh.filter = g ? { r0: g.r0, c0: g.c0, r1: g.r1, c1: g.c1, cols: {} } : null; });
  };
  W.setFilterCol = function (sh, c, crit) {
    var self = this;
    this.run("Filter", function () { self.meta(sh); if (crit) sh.filter.cols[c] = crit; else delete sh.filter.cols[c]; });
  };
  W.refreshFilters = function () { var self = this; this.sheets.forEach(function (sh) { self.computeHidden(sh); }); };
  W.computeHidden = function (sh) {
    var f = sh.filter, hidden = new Set(), self = this;
    if (f) {
      var tests = Object.keys(f.cols).map(function (k) {
        var cr = f.cols[k], c = +k;
        if (cr.hide) { var set = new Set(cr.hide); return function (r) { return !set.has(self.text(sh, r, c)); }; }
        if (cr.op) { var pred = SH.crit(cr.op + (cr.val === undefined ? "" : cr.val)); if (cr.op === "contains") { var t = String(cr.val || "").toLowerCase(); pred = function (v) { return F.format(v).t.toLowerCase().indexOf(t) >= 0; }; } return function (r) { return pred(self.eng.value(sh, r, c)); }; }
        return function () { return true; };
      });
      if (tests.length) for (var r = f.r0 + 1; r <= f.r1; r++) { for (var i = 0; i < tests.length; i++) if (!tests[i](r)) { hidden.add(r); break; } }
    }
    var changed = hidden.size !== sh.hidden.size;
    if (!changed) hidden.forEach(function (r) { if (!sh.hidden.has(r)) changed = true; });
    sh.hidden = hidden;
    if (changed) sh.hiddenVer = (sh.hiddenVer || 0) + 1;
    return changed;
  };

  // ---------- find & replace ----------
  W.find = function (q, o) {
    var self = this, out = [], sheets = o.all ? this.sheets : [this.sheet()];
    if (!q) return out;
    var rx = o.regex ? safeRx(q, o.caseSens) : new RegExp((o.whole ? "^" : "") + q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + (o.whole ? "$" : ""), o.caseSens ? "" : "i");
    if (!rx) return out;
    sheets.forEach(function (sh) {
      var keys = Array.from(sh.cells.keys()).sort(function (a, b) { return a - b; });
      keys.forEach(function (k) {
        var r = Math.floor(k / KW), c = k % KW, x = sh.cells.get(k), hay = o.formulas && x.f !== undefined ? "=" + x.f : self.text(sh, r, c);
        if (hay && rx.test(hay)) out.push({ sh: sh, r: r, c: c });
      });
    });
    return out;
  };
  function safeRx(q, cs) { try { return new RegExp(q, cs ? "" : "i"); } catch (e) { return null; } }
  W.replaceAll = function (q, rep, o, only) {
    var self = this, hits = only || this.find(q, Object.assign({}, o, { formulas: true })), count = 0;
    var rx = new RegExp((o.whole ? "^" : "") + q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + (o.whole ? "$" : ""), o.caseSens ? "g" : "gi");
    this.run("Replace", function () {
      hits.forEach(function (h) {
        var x = self.cell(h.sh, h.r, h.c);
        if (!x) return;
        if (x.f !== undefined) { var f2 = ("=" + x.f).replace(rx, rep); if (f2 !== "=" + x.f) { self.setInput(h.sh, h.r, h.c, f2); count++; } return; }
        if (typeof x.v !== "string" && x.v !== undefined) { var t = self.inputText(h.sh, h.r, h.c), t2 = t.replace(rx, rep); if (t2 !== t) { self.setInput(h.sh, h.r, h.c, t2); count++; } return; }
        if (typeof x.v === "string") { var s2 = x.v.replace(rx, rep); if (s2 !== x.v) { self.set(h.sh, h.r, h.c, mk(s2, undefined, x.s)); count++; } }
      });
    });
    return count;
  };

  // ---------- serialisation (autosave) ----------
  W.toJSON = function () {
    return {
      v: 1, id: this.id, name: this.name, active: this.active, created: this.created,
      sheets: this.sheets.map(function (sh) {
        var cells = [];
        sh.cells.forEach(function (x, k) { var o = {}; if (x.f !== undefined) o.f = x.f; else if (x.v !== undefined) o.v = x.v instanceof C.Err ? { e: x.v.e } : x.v; if (x.s) o.s = x.s; if (x.q) o.q = 1; cells.push([Math.floor(k / KW), k % KW, o]); });
        var m = metaOf(sh); m.cells = cells; return m;
      })
    };
  };
  var STYLE_KEYS = { b: "boolean", i: "boolean", u: "boolean", st: "boolean", fs: "number", fc: "string", bg: "string", ha: "string", va: "string", wr: "boolean", nf: "string", bd: "object", ff: "string" };
  function cleanStyle(s) {
    if (!s || typeof s !== "object") return undefined;
    var o = {}, any = false;
    for (var k in STYLE_KEYS) if (typeof s[k] === STYLE_KEYS[k] && s[k] !== null) {
      var v = s[k];
      if (k === "fc" || k === "bg") { if (!/^#[0-9a-fA-F]{6}$/.test(v)) continue; }
      if (k === "nf") v = v.slice(0, 200);
      if (k === "fs") v = Math.max(6, Math.min(72, v));
      if (k === "bd") { var bd = {}; ["t", "r", "b", "l"].forEach(function (sd) { var x = v[sd]; if (x && typeof x === "object" && /^#[0-9a-fA-F]{6}$/.test(x.c)) bd[sd] = { s: x.s === "medium" || x.s === "thick" || x.s === "dashed" ? x.s : "thin", c: x.c }; }); v = bd; }
      o[k] = v; any = true;
    }
    return any ? o : undefined;
  }
  function cleanRange(g) { if (!g || typeof g !== "object") return null; var o = { r0: g.r0 | 0, c0: g.c0 | 0, r1: g.r1 | 0, c1: g.c1 | 0 }; if (o.r0 < 0 || o.c0 < 0 || o.r1 < o.r0 || o.c1 < o.c0 || o.r1 >= MAXR || o.c1 >= MAXC) return null; return o; }
  // never trusts the input: unknown fields dropped, sizes capped
  function fromJSON(o, limits) {
    limits = limits || { cells: 2e6 };
    if (!o || typeof o !== "object" || !Array.isArray(o.sheets) || !o.sheets.length) throw new Error("Not a workbook");
    var wb = new Workbook(String(o.name || "Untitled spreadsheet").slice(0, 120)), total = 0;
    if (typeof o.id === "string" && /^[\w-]{4,60}$/.test(o.id)) wb.id = o.id;
    if (typeof o.created === "number") wb.created = o.created;
    o.sheets.slice(0, 200).forEach(function (s) {
      var sh = wb.addSheetRaw(wb.uniqueName(String(s.name || "Sheet").slice(0, 31) || "Sheet"));
      (Array.isArray(s.cells) ? s.cells : []).forEach(function (e) {
        if (!Array.isArray(e) || total >= limits.cells) return;
        var r = e[0] | 0, c = e[1] | 0, x = e[2] || {};
        if (r < 0 || c < 0 || r >= MAXR || c >= MAXC) return;
        var cell = {};
        if (typeof x.f === "string") cell.f = x.f.slice(0, 8192);
        else if (x.v && typeof x.v === "object" && typeof x.v.e === "string") cell.v = C.errOf(x.v.e);
        else if (typeof x.v === "number" || typeof x.v === "string" || typeof x.v === "boolean") cell.v = typeof x.v === "string" ? x.v.slice(0, 32767) : x.v;
        var st = cleanStyle(x.s); if (st) cell.s = st;
        if (x.q) cell.q = 1;
        if (cell.f === undefined && cell.v === undefined && !cell.s) return;
        sh.cells.set(r * KW + c, cell); total++;
      });
      (Array.isArray(s.colW) ? s.colW : []).forEach(function (p) { if (Array.isArray(p) && p[0] >= 0 && p[0] < MAXC && p[1] > 0) sh.colW.set(p[0] | 0, Math.min(2000, +p[1])); });
      (Array.isArray(s.rowH) ? s.rowH : []).forEach(function (p) { if (Array.isArray(p) && p[0] >= 0 && p[0] < MAXR && p[1] > 0) sh.rowH.set(p[0] | 0, Math.min(2000, +p[1])); });
      sh.merges = (Array.isArray(s.merges) ? s.merges : []).map(cleanRange).filter(Boolean).slice(0, 20000);
      sh.freeze = { r: Math.max(0, Math.min(200, (s.freeze && s.freeze.r) | 0)), c: Math.max(0, Math.min(50, (s.freeze && s.freeze.c) | 0)) };
      sh.rtl = !!s.rtl; sh.grid = s.grid !== false;
      sh.cf = (Array.isArray(s.cf) ? s.cf : []).slice(0, 100).map(cleanRule).filter(Boolean);
      sh.charts = (Array.isArray(s.charts) ? s.charts : []).slice(0, 50).map(cleanChart).filter(Boolean);
      if (s.filter && cleanRange(s.filter)) { sh.filter = cleanRange(s.filter); sh.filter.cols = {}; var fc = s.filter.cols || {}; Object.keys(fc).slice(0, 200).forEach(function (k) { var x = fc[k]; if (!x) return; if (Array.isArray(x.hide)) sh.filter.cols[k | 0] = { hide: x.hide.slice(0, 5000).map(String) }; else if (typeof x.op === "string") sh.filter.cols[k | 0] = { op: x.op.slice(0, 10), val: String(x.val == null ? "" : x.val).slice(0, 200) }; }); }
      sh.nrows = Math.max(DEF_ROWS, Math.min(MAXR, s.nrows | 0)); sh.ncols = Math.max(DEF_COLS, Math.min(MAXC, s.ncols | 0));
      extents(sh);
    });
    wb.active = Math.max(0, Math.min(wb.sheets.length - 1, o.active | 0));
    wb.eng.rebuild(); wb.refreshFilters();
    return wb;
  }
  var CF_TYPES = { gt: 1, lt: 1, ge: 1, le: 1, eq: 1, ne: 1, between: 1, contains: 1, blank: 1, notblank: 1, dup: 1, scale: 1, bar: 1, top: 1, formula: 1 };
  function cleanRule(x) {
    if (!x || !CF_TYPES[x.type]) return null;
    var g = cleanRange(x.range); if (!g) return null;
    var o = { id: typeof x.id === "string" ? x.id.slice(0, 40) : C.uid("cf"), type: x.type, range: g, a: x.a == null ? "" : String(x.a).slice(0, 500), b: x.b == null ? "" : String(x.b).slice(0, 500), style: cleanStyle(x.style) || {} };
    if (Array.isArray(x.colors)) o.colors = x.colors.filter(function (c) { return /^#[0-9a-fA-F]{6}$/.test(c); }).slice(0, 3);
    return o;
  }
  var CH_TYPES = { bar: 1, column: 1, line: 1, area: 1, pie: 1, doughnut: 1, scatter: 1 };
  function cleanChart(x) {
    if (!x || !CH_TYPES[x.type] || typeof x.range !== "string") return null;
    return { id: typeof x.id === "string" ? x.id.slice(0, 40) : C.uid("ch"), type: x.type, range: x.range.slice(0, 300), title: String(x.title || "").slice(0, 200),
      x: Math.max(0, +x.x || 0), y: Math.max(0, +x.y || 0), w: Math.max(160, Math.min(2400, +x.w || 480)), h: Math.max(120, Math.min(1600, +x.h || 300)), legend: x.legend !== false };
  }
  function blank(name) { var wb = new Workbook(name); wb.addSheetRaw("Sheet1"); return wb; }

  SH.model = { Workbook: Workbook, Sheet: Sheet, blank: blank, fromJSON: fromJSON, mk: mk, mergeStyle: mergeStyle, withStyle: withStyle,
    cleanStyle: cleanStyle, cleanRule: cleanRule, cleanChart: cleanChart, extents: extents, DEF_ROWS: DEF_ROWS, DEF_COLS: DEF_COLS };
})();
