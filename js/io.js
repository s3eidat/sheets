/* Sheets I/O: .xlsx in/out (ExcelJS, files treated as untrusted: zip directory checked before unpacking,
   sizes capped, formulas imported as text and run only by our engine, macros/links ignored), CSV/TSV
   (formula-injection guard on export), IndexedDB autosave and print/PDF of a range. */
(function () {
  "use strict";
  var SH = window.SH, C = SH.core, F = SH.fmt, EN = SH.engine, M = SH.model, KW = C.KW;
  var LIMITS = { bytes: 40 * 1024 * 1024, unzipped: 400 * 1024 * 1024, entries: 5000, rows: 200000, cols: 2000, cells: 2000000, formula: 8192, text: 32767 };
  var EXCEL_ERRORS = { "#DIV/0!": 1, "#N/A": 1, "#NAME?": 1, "#NULL!": 1, "#NUM!": 1, "#REF!": 1, "#VALUE!": 1 };

  // ---------- zip directory check (before any decompression) ----------
  function zipCheck(buf) {
    if (ArrayBuffer.isView(buf)) buf = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    var u8 = new Uint8Array(buf), dv = new DataView(buf), n = u8.length, eocd = -1;
    if (n < 22 || n > LIMITS.bytes) throw new Error(n > LIMITS.bytes ? "That file is larger than 40 MB." : "That file is not a valid .xlsx.");
    if (u8[0] !== 0x50 || u8[1] !== 0x4b) throw new Error("That file is not a valid .xlsx (it is not a zip package).");
    for (var i = n - 22; i >= Math.max(0, n - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error("That .xlsx is damaged (no zip directory).");
    var count = dv.getUint16(eocd + 10, true), cdSize = dv.getUint32(eocd + 12, true), cdOff = dv.getUint32(eocd + 16, true);
    if (count > LIMITS.entries) throw new Error("That file has too many parts to be a normal spreadsheet.");
    if (cdOff + cdSize > n || count === 0xffff || cdOff === 0xffffffff) throw new Error("That .xlsx uses an unsupported or damaged zip layout.");
    var p = cdOff, total = 0, names = [];
    for (var k = 0; k < count; k++) {
      if (p + 46 > n || dv.getUint32(p, true) !== 0x02014b50) throw new Error("That .xlsx is damaged (bad zip directory).");
      var csize = dv.getUint32(p + 20, true), usize = dv.getUint32(p + 24, true), nl = dv.getUint16(p + 28, true), xl = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
      if (usize === 0xffffffff || csize === 0xffffffff) throw new Error("That file is too large to open here.");
      var name = ""; for (var j = 0; j < nl && p + 46 + j < n; j++) name += String.fromCharCode(u8[p + 46 + j]);
      names.push(name);
      total += usize;
      if (usize > 50 * 1024 * 1024 && usize / Math.max(1, csize) > 200) throw new Error("That file looks like a zip bomb (a part expands " + Math.round(usize / Math.max(1, csize)) + "×). Not opened.");
      p += 46 + nl + xl + cl;
    }
    if (total > LIMITS.unzipped) throw new Error("That file expands to more than 400 MB. Not opened.");
    if (names.indexOf("xl/workbook.xml") < 0) throw new Error("That file is not an Excel workbook (.xlsx).");
    return { macros: names.some(function (x) { return /vbaProject\.bin$/i.test(x); }), links: names.some(function (x) { return /^xl\/externalLinks\//i.test(x); }), total: total };
  }

  // ---------- styles between ExcelJS and ours ----------
  function argb(c) {
    if (!c || typeof c.argb !== "string") return null;
    var h = c.argb.length === 8 ? c.argb.slice(2) : c.argb;
    return /^[0-9a-fA-F]{6}$/.test(h) ? "#" + h.toLowerCase() : null;
  }
  var BORDER_IN = { thin: "thin", hair: "thin", dotted: "dashed", dashed: "dashed", dashDot: "dashed", medium: "medium", mediumDashed: "medium", thick: "thick", double: "thick" };
  function styleIn(cell, cache) {
    var s = {}, f = cell.font, fill = cell.fill, al = cell.alignment, bd = cell.border, nf = cell.numFmt;
    if (f) { if (f.bold) s.b = true; if (f.italic) s.i = true; if (f.underline) s.u = true; if (f.strike) s.st = true; if (f.size && f.size !== 11 && f.size !== 10) s.fs = Math.max(6, Math.min(72, Math.round(f.size * 1.2))); var fc = argb(f.color); if (fc && fc !== "#000000") s.fc = fc; }
    if (fill && fill.type === "pattern" && fill.pattern === "solid") { var bg = argb(fill.fgColor); if (bg && bg !== "#ffffff") s.bg = bg; }
    if (al) {
      if (al.horizontal === "left" || al.horizontal === "center" || al.horizontal === "right") s.ha = al.horizontal;
      if (al.horizontal === "centerContinuous") s.ha = "center";
      if (al.vertical === "top" || al.vertical === "bottom") s.va = al.vertical; if (al.vertical === "middle") s.va = "middle";
      if (al.wrapText) s.wr = true;
    }
    if (bd) { var o = {}; ["top", "right", "bottom", "left"].forEach(function (k) { var x = bd[k]; if (x && BORDER_IN[x.style]) o[k[0]] = { s: BORDER_IN[x.style], c: argb(x.color) || "#000000" }; }); for (var q in o) { s.bd = o; break; } }
    if (typeof nf === "string" && nf && nf !== "General") s.nf = nf.slice(0, 200);
    var clean = M.cleanStyle(s);
    if (!clean) return undefined;
    var key = JSON.stringify(clean), hit = cache.get(key);
    if (hit) return hit;
    cache.set(key, clean);
    return clean;
  }
  var BORDER_OUT = { thin: "thin", medium: "medium", thick: "thick", dashed: "dashed" };
  function styleOut(wc, s) {
    if (!s) return;
    var font = {};
    if (s.b) font.bold = true; if (s.i) font.italic = true; if (s.u) font.underline = true; if (s.st) font.strike = true;
    if (s.fs) font.size = Math.round(s.fs / 1.2 * 2) / 2;
    if (s.fc) font.color = { argb: "FF" + s.fc.slice(1).toUpperCase() };
    for (var k in font) { wc.font = font; break; }
    if (s.bg) wc.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + s.bg.slice(1).toUpperCase() } };
    if (s.ha || s.va || s.wr) { var al = {}; if (s.ha) al.horizontal = s.ha; if (s.va) al.vertical = s.va === "middle" ? "middle" : s.va; if (s.wr) al.wrapText = true; wc.alignment = al; }
    if (s.bd) { var b = {}; [["t", "top"], ["r", "right"], ["b", "bottom"], ["l", "left"]].forEach(function (p) { var x = s.bd[p[0]]; if (x) b[p[1]] = { style: BORDER_OUT[x.s] || "thin", color: { argb: "FF" + x.c.slice(1).toUpperCase() } }; }); wc.border = b; }
    if (s.nf) wc.numFmt = s.nf;
  }

  // ---------- .xlsx import ----------
  function textOf(v) {
    if (v && Array.isArray(v.richText)) return v.richText.map(function (t) { return String(t.text || ""); }).join("");
    if (v && typeof v.text === "string") return v.text;
    return null;
  }
  async function readXlsx(buf, name) {
    if (typeof ExcelJS === "undefined") throw new Error("The .xlsx reader did not load.");
    if (ArrayBuffer.isView(buf)) buf = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    var info = zipCheck(buf), notes = [];
    if (info.macros) notes.push("macros were ignored");
    if (info.links) notes.push("links to other files were ignored");
    var x = new ExcelJS.Workbook();
    await x.xlsx.load(buf);
    var wb = new M.Workbook(String(name || "Imported").replace(/\.xlsx?$/i, "").slice(0, 120)), cache = new Map(), total = 0, cut = false, meta = null;
    x.worksheets.forEach(function (ws) {
      if (ws.name === "_s3meta") { meta = readMeta(ws); return; }
      if (wb.sheets.length >= 200) return;
      var sh = wb.addSheetRaw(wb.uniqueName(String(ws.name || "Sheet").replace(/[\\/?*[\]:]/g, "").slice(0, 31) || "Sheet"));
      ws.eachRow({ includeEmpty: false }, function (row, rn) {
        if (rn > LIMITS.rows) { cut = true; return; }
        if (row.height && row.height > 0) sh.rowH.set(rn - 1, Math.max(8, Math.min(2000, Math.round(row.height * 4 / 3))));
        row.eachCell({ includeEmpty: false }, function (cell, cn) {
          if (cn > LIMITS.cols || total >= LIMITS.cells) { cut = true; return; }
          if (cell.type === ExcelJS.ValueType.Merge) return;
          var v = cell.value, out = {}, st = styleIn(cell, cache);
          var formula = null;
          try { formula = cell.formula; } catch (e) { formula = null; }
          if (!formula && v && typeof v === "object" && typeof v.formula === "string") formula = v.formula;
          if (formula && typeof formula === "string") {
            out.f = EN.fromFile(formula.replace(/^=/, "")).slice(0, LIMITS.formula);
          } else {
            if (v && typeof v === "object" && !(v instanceof Date) && "result" in v) v = v.result;
            if (v instanceof Date) { out.v = F.fromDate(v); if (!st || !st.nf) st = M.mergeStyle(st, { nf: "yyyy-mm-dd" }); }
            else if (typeof v === "number" && isFinite(v)) out.v = v;
            else if (typeof v === "boolean") out.v = v;
            else if (typeof v === "string") out.v = v.slice(0, LIMITS.text);
            else if (v && typeof v === "object" && typeof v.error === "string") out.v = C.errOf(v.error);
            else { var t = textOf(v); if (t !== null) out.v = t.slice(0, LIMITS.text); }
          }
          if (st) out.s = st;
          if (out.f === undefined && out.v === undefined && !out.s) return;
          sh.cells.set((rn - 1) * KW + (cn - 1), out); total++;
        });
      });
      var cc = Math.min(ws.columnCount || 0, LIMITS.cols);
      for (var c = 1; c <= cc; c++) { var col = ws.getColumn(c); if (col && col.width) sh.colW.set(c - 1, Math.max(12, Math.min(2000, Math.round(col.width * 7 + 5)))); }
      var merges = (ws.model && ws.model.merges) || [];
      merges.slice(0, 20000).forEach(function (m) { var g = C.parseRange(m); if (g) sh.merges.push({ r0: g.r0, c0: g.c0, r1: g.r1, c1: g.c1 }); });
      var view = ws.views && ws.views[0];
      if (view) {
        if (view.state === "frozen") sh.freeze = { r: Math.max(0, Math.min(200, view.ySplit | 0)), c: Math.max(0, Math.min(50, view.xSplit | 0)) };
        if (view.rightToLeft) sh.rtl = true;
        if (view.showGridLines === false) sh.grid = false;
      }
      M.extents(sh);
    });
    if (!wb.sheets.length) wb.addSheetRaw("Sheet1");
    if (meta) applyMeta(wb, meta);
    wb.eng.rebuild(); wb.refreshFilters();
    if (cut) notes.push("very large sheets were cut to " + LIMITS.rows.toLocaleString() + " rows × " + LIMITS.cols + " columns");
    return { wb: wb, notes: notes };
  }
  function readMeta(ws) {
    var s = "";
    for (var r = 1; r <= 80; r++) { var v = ws.getCell(r, 1).value; if (typeof v !== "string") break; s += v; if (s.length > 2e6) return null; }
    try { var o = JSON.parse(s); return o && o.v === 1 ? o : null; } catch (e) { return null; }
  }
  function applyMeta(wb, meta) {
    if (!meta.sheets || typeof meta.sheets !== "object") return;
    wb.sheets.forEach(function (sh) {
      var m = meta.sheets[sh.name];
      if (!m) return;
      if (Array.isArray(m.cf)) sh.cf = m.cf.slice(0, 100).map(M.cleanRule).filter(Boolean);
      if (Array.isArray(m.charts)) sh.charts = m.charts.slice(0, 50).map(M.cleanChart).filter(Boolean);
    });
  }

  // ---------- .xlsx export ----------
  function resultOut(v) {
    if (v instanceof C.Err) return EXCEL_ERRORS[v.e] ? { error: v.e } : undefined;
    if (v === null || v === undefined) return undefined;
    return v;
  }
  var CF_OUT = { gt: "greaterThan", lt: "lessThan", ge: "greaterThanOrEqual", le: "lessThanOrEqual", eq: "equal", ne: "notEqual", between: "between" };
  function cfFormula(t) { var n = F.parseInput(String(t)); return typeof n.v === "number" ? String(n.v) : '"' + String(t).replace(/"/g, '""') + '"'; }
  async function writeXlsx(wb) {
    if (typeof ExcelJS === "undefined") throw new Error("The .xlsx writer did not load.");
    var x = new ExcelJS.Workbook();
    x.creator = "S3eidat Sheets"; x.created = new Date(); x.modified = new Date();
    x.calcProperties = { fullCalcOnLoad: true };
    var meta = { v: 1, sheets: {} }, needMeta = false;
    wb.sheets.forEach(function (sh) {
      var fr = sh.freeze || { r: 0, c: 0 };
      var ws = x.addWorksheet(sh.name, { views: [fr.r || fr.c ? { state: "frozen", xSplit: fr.c, ySplit: fr.r, rightToLeft: !!sh.rtl, showGridLines: sh.grid !== false } : { rightToLeft: !!sh.rtl, showGridLines: sh.grid !== false }] });
      ws.properties.defaultColWidth = 13.57;
      sh.colW.forEach(function (px, c) { ws.getColumn(c + 1).width = Math.max(1, (px - 5) / 7); });
      sh.rowH.forEach(function (px, r) { ws.getRow(r + 1).height = Math.round(px * 0.75 * 4) / 4; });
      var keys = Array.from(sh.cells.keys()).sort(function (a, b) { return a - b; });
      keys.forEach(function (k) {
        var r = Math.floor(k / KW), c = k % KW, cell = sh.cells.get(k), wc = ws.getCell(r + 1, c + 1);
        if (cell.f !== undefined) {
          var res = resultOut(wb.value(sh, r, c));
          wc.value = res === undefined ? { formula: EN.toFile(cell.f) } : { formula: EN.toFile(cell.f), result: res };
        } else if (cell.v !== undefined) {
          var v = cell.v;
          wc.value = v instanceof C.Err ? (EXCEL_ERRORS[v.e] ? { error: v.e } : v.e) : v;
        }
        styleOut(wc, cell.s);
      });
      sh.merges.forEach(function (g) { try { ws.mergeCells(g.r0 + 1, g.c0 + 1, g.r1 + 1, g.c1 + 1); } catch (e) { /* overlapping merges are skipped */ } });
      if (sh.filter) ws.autoFilter = C.rangeStr(sh.filter);
      sh.cf.forEach(function (rule) {
        var st = {}, fill = rule.style && rule.style.bg, fc = rule.style && rule.style.fc;
        if (fill) st.fill = { type: "pattern", pattern: "solid", bgColor: { argb: "FF" + fill.slice(1).toUpperCase() } };
        if (fc || (rule.style && rule.style.b)) st.font = {}; if (fc) st.font.color = { argb: "FF" + fc.slice(1).toUpperCase() }; if (rule.style && rule.style.b) st.font.bold = true;
        var spec = null;
        if (CF_OUT[rule.type]) spec = { type: "cellIs", operator: CF_OUT[rule.type], formulae: rule.type === "between" ? [cfFormula(rule.a), cfFormula(rule.b)] : [cfFormula(rule.a)], style: st };
        else if (rule.type === "contains") spec = { type: "containsText", operator: "containsText", text: String(rule.a), style: st };
        else if (rule.type === "scale" && rule.colors && rule.colors.length >= 2) spec = { type: "colorScale", cfvo: rule.colors.length === 3 ? [{ type: "min" }, { type: "percentile", value: 50 }, { type: "max" }] : [{ type: "min" }, { type: "max" }], color: rule.colors.map(function (c) { return { argb: "FF" + c.slice(1).toUpperCase() }; }) };
        if (spec) { try { ws.addConditionalFormatting({ ref: C.rangeStr(rule.range), rules: [spec] }); } catch (e) { /* best effort */ } }
      });
      if (sh.cf.length || sh.charts.length) { meta.sheets[sh.name] = { cf: sh.cf, charts: sh.charts }; needMeta = true; }
    });
    if (needMeta) {
      // charts and our rules ride along in a hidden sheet so they come back when the file is reopened here
      var ms = x.addWorksheet("_s3meta", { state: "veryHidden" }), s = JSON.stringify(meta);
      for (var i = 0, r = 1; i < s.length && r <= 80; i += 30000, r++) ms.getCell(r, 1).value = s.slice(i, i + 30000);
    }
    var buf = await x.xlsx.writeBuffer();
    return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  // ---------- CSV / TSV ----------
  function sniff(text) {
    var line = text.slice(0, 5000).split(/\r?\n/).slice(0, 5).join("\n"), best = ",", bc = -1;
    [",", ";", "\t", "|"].forEach(function (d) { var n = line.split(d).length; if (n > bc) { bc = n; best = d; } });
    return best;
  }
  function parseDelimited(text, delim, limits) {
    limits = limits || LIMITS;
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
    if (!delim || delim === "auto") delim = sniff(text);
    var rows = [], row = [], cell = "", i = 0, n = text.length, q = false, cut = false;
    while (i < n) {
      var ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i += 2; continue; } q = false; i++; continue; }
        cell += ch; i++; continue;
      }
      if (ch === '"' && cell === "") { q = true; i++; continue; }
      if (ch === delim) { row.push(cell.slice(0, limits.text)); cell = ""; i++; if (row.length > limits.cols) { cut = true; } continue; }
      if (ch === "\r" || ch === "\n") {
        row.push(cell.slice(0, limits.text)); cell = "";
        if (row.length > limits.cols) { row.length = limits.cols; cut = true; }
        rows.push(row); row = [];
        if (rows.length >= limits.rows) { cut = true; break; }
        if (ch === "\r" && text[i + 1] === "\n") i++;
        i++; continue;
      }
      cell += ch; i++;
    }
    if (!cut && (cell !== "" || row.length)) { row.push(cell.slice(0, limits.text)); if (row.length > limits.cols) { row.length = limits.cols; cut = true; } rows.push(row); }
    return { rows: rows, delim: delim, cut: cut };
  }
  // CSV -> new workbook. Cells that look like formulas stay text: a downloaded CSV should not run anything.
  function readCsv(text, name, delim) {
    if (text.length > LIMITS.bytes) throw new Error("That file is larger than 40 MB.");
    var p = parseDelimited(text, delim), wb = new M.Workbook(String(name || "Imported").replace(/\.(csv|tsv|txt)$/i, "").slice(0, 120)), sh = wb.addSheetRaw("Sheet1"), styles = new Map();
    p.rows.forEach(function (row, r) {
      row.forEach(function (t, c) {
        if (t === "") return;
        var x = F.parseInput(t), cell;
        if (x.f !== undefined || x.text) cell = { v: t, q: x.f !== undefined ? 1 : undefined };
        else if (x.v === null) return;
        else { cell = { v: x.v }; if (x.nf) { var s = styles.get(x.nf); if (!s) styles.set(x.nf, s = { nf: x.nf }); cell.s = s; } }
        if (cell.q === undefined) delete cell.q;
        sh.cells.set(r * KW + c, cell);
      });
    });
    M.extents(sh);
    wb.eng.rebuild();
    return { wb: wb, notes: p.cut ? ["the file was cut to " + LIMITS.rows.toLocaleString() + " rows × " + LIMITS.cols + " columns"] : [], delim: p.delim };
  }
  var NUM = /^-?(0|[1-9]\d*)(\.\d+)?([eE][+-]?\d+)?$/;
  function csvField(s, delim, guard) {
    if (guard && /^[=+\-@\t\r]/.test(s) && !NUM.test(s)) s = "'" + s;        // stops other spreadsheet apps running it as a formula
    return /["\r\n]|^\s|\s$/.test(s) || s.indexOf(delim) >= 0 ? '"' + s.replace(/"/g, '""') + '"' : s;
  }
  function plain(wb, sh, r, c) {
    var v = wb.value(sh, r, c), s = wb.style(sh, r, c);
    if (v === null) return "";
    if (v instanceof C.Err) return v.e;
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (typeof v === "number") {
      if (s && s.nf && F.isDateFmt(s.nf)) { var hasT = /[hs]/i.test(s.nf.replace(/"[^"]*"/g, "")); return F.format(v, F.isTimeOnly(s.nf) ? "h:mm:ss" : hasT ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd").t; }
      return F.numText(v);
    }
    return String(v);
  }
  function writeCsv(wb, sh, g, o) {
    o = o || {};
    var delim = o.delim || ",", guard = o.guard !== false, lines = [];
    g = g || { r0: 0, c0: 0, r1: Math.max(sh.maxR, sh.smaxR || -1), c1: Math.max(sh.maxC, sh.smaxC || -1) };
    for (var r = g.r0; r <= g.r1; r++) {
      var row = [];
      for (var c = g.c0; c <= g.c1; c++) row.push(csvField(o.shown ? wb.text(sh, r, c) : plain(wb, sh, r, c), delim, guard));
      lines.push(row.join(delim));
    }
    return (o.bom === false ? "" : "﻿") + lines.join("\r\n") + (lines.length ? "\r\n" : "");
  }

  // ---------- autosave (IndexedDB, this browser only) ----------
  var dbp = null;
  function db() {
    if (dbp) return dbp;
    dbp = new Promise(function (res, rej) {
      if (!window.indexedDB) return rej(new Error("No IndexedDB"));
      var rq = indexedDB.open("s3-sheets", 1);
      rq.onupgradeneeded = function () { var d = rq.result; if (!d.objectStoreNames.contains("books")) d.createObjectStore("books", { keyPath: "id" }); };
      rq.onsuccess = function () { res(rq.result); };
      rq.onerror = function () { rej(rq.error); };
      rq.onblocked = function () { rej(new Error("Storage is blocked")); };
    });
    dbp.catch(function () { dbp = null; });
    return dbp;
  }
  function tx(mode, fn) {
    return db().then(function (d) {
      return new Promise(function (res, rej) {
        var t = d.transaction("books", mode), s = t.objectStore("books"), out = fn(s);
        t.oncomplete = function () { res(out && out.result !== undefined ? out.result : out); };
        t.onerror = function () { rej(t.error); }; t.onabort = function () { rej(t.error || new Error("Storage aborted")); };
      });
    });
  }
  var store = {
    put: function (wb) { var j = wb.toJSON(); return tx("readwrite", function (s) { s.put({ id: wb.id, name: wb.name, updated: Date.now(), sheets: wb.sheets.length, data: j }); }); },
    get: function (id) { return tx("readonly", function (s) { return s.get(id); }); },
    del: function (id) { return tx("readwrite", function (s) { s.delete(id); }); },
    list: function () {
      return tx("readonly", function (s) { return s.getAll(); }).then(function (all) {
        return (all || []).map(function (r) { return { id: r.id, name: r.name, updated: r.updated, sheets: r.sheets }; }).sort(function (a, b) { return b.updated - a.updated; });
      });
    }
  };

  // ---------- print / PDF ----------
  function cssColor(c) { return /^#[0-9a-fA-F]{6}$/.test(c || "") ? c : null; }
  function printHtml(wb, sh, g, charts) {
    var root = document.createElement("div"), tbl = document.createElement("table"), covered = new Set();
    root.className = "sh-print"; tbl.className = "sh-ptable";
    if (sh.rtl) root.setAttribute("dir", "rtl");
    var cg = document.createElement("colgroup");
    for (var c = g.c0; c <= g.c1; c++) { var col = document.createElement("col"); col.style.width = (sh.colW.get(c) || 100) + "px"; cg.appendChild(col); }
    tbl.appendChild(cg);
    for (var r = g.r0; r <= g.r1; r++) {
      if (sh.hidden.has(r)) continue;
      var tr = document.createElement("tr");
      tr.style.height = (sh.rowH.get(r) || 24) + "px";
      for (c = g.c0; c <= g.c1; c++) {
        if (covered.has(r * KW + c)) continue;
        var td = document.createElement("td"), m = wb.mergeAt(sh, r, c), s = wb.style(sh, r, c) || {}, v = wb.value(sh, r, c), fm = F.format(v, s.nf);
        if (m && m.r0 === r && m.c0 === c) {
          td.rowSpan = Math.min(m.r1, g.r1) - r + 1; td.colSpan = Math.min(m.c1, g.c1) - c + 1;
          for (var a = r; a <= Math.min(m.r1, g.r1); a++) for (var b = c; b <= Math.min(m.c1, g.c1); b++) if (a !== r || b !== c) covered.add(a * KW + b);
        }
        td.textContent = fm.t;
        var st = td.style;
        if (s.b) st.fontWeight = "700"; if (s.i) st.fontStyle = "italic"; if (s.u || s.st) st.textDecoration = (s.u ? "underline " : "") + (s.st ? "line-through" : "");
        if (s.fs) st.fontSize = s.fs + "px";
        var fc = cssColor(fm.color) || cssColor(s.fc), bg = cssColor(s.bg);
        if (fc) st.color = fc; if (bg) st.background = bg;
        st.textAlign = s.ha || (typeof v === "number" ? "right" : typeof v === "boolean" || v instanceof C.Err ? "center" : "left");
        st.verticalAlign = s.va === "top" ? "top" : s.va === "middle" ? "middle" : "bottom";
        st.whiteSpace = s.wr ? "pre-wrap" : "pre";
        if (s.bd) [["t", "Top"], ["r", "Right"], ["b", "Bottom"], ["l", "Left"]].forEach(function (p) { var x = s.bd[p[0]]; if (x && cssColor(x.c)) st["border" + p[1]] = (x.s === "thick" ? "3px" : x.s === "medium" ? "2px" : "1px") + (x.s === "dashed" ? " dashed " : " solid ") + x.c; });
        tr.appendChild(td);
      }
      tbl.appendChild(tr);
    }
    root.appendChild(tbl);
    (charts || []).forEach(function (url) { var img = document.createElement("img"); img.src = url; img.className = "sh-pchart"; img.alt = "Chart"; root.appendChild(img); });
    return root;
  }

  SH.io = { zipCheck: zipCheck, readXlsx: readXlsx, writeXlsx: writeXlsx, parseDelimited: parseDelimited, readCsv: readCsv, writeCsv: writeCsv, csvField: csvField, store: store, printHtml: printHtml, LIMITS: LIMITS };
})();
