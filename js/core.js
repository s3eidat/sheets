/* Sheets core: addresses, ranges, error values and small helpers. Pure logic, no DOM. */
(function () {
  "use strict";
  var SH = window.SH = window.SH || {};

  var MAXR = 1048576, MAXC = 16384;          // the same grid size as the common .xlsx limit
  var KW = 16384;                             // local key = r * KW + c

  // ---------- error values ----------
  function Err(code) { this.e = code; }
  Err.prototype.toString = function () { return this.e; };
  var E = {};
  ["#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "#CIRC!", "#SPILL!", "#CALC!", "#ERROR!"].forEach(function (c) { E[c] = new Err(c); });
  var ERR_HELP = {
    "#DIV/0!": "Division by zero (or by an empty cell).",
    "#VALUE!": "A value has the wrong type, e.g. text where a number is needed.",
    "#REF!": "The formula points at a cell or sheet that no longer exists.",
    "#NAME?": "Unknown function or name. Check the spelling.",
    "#N/A": "Not found / not available (e.g. a lookup with no match).",
    "#NUM!": "The number is invalid or too large for this calculation.",
    "#CIRC!": "Circular reference: the formula depends on its own result.",
    "#SPILL!": "The result needs more cells, but some are not empty.",
    "#CALC!": "The calculation returned nothing (e.g. FILTER matched no rows).",
    "#ERROR!": "The formula could not be read. Check brackets, quotes and commas."
  };
  function isErr(v) { return v instanceof Err; }
  function errOf(code) { return E[code] || new Err(code); }

  // A 2-D array value (results of ranges in expressions, UNIQUE, SORT…).
  function Mat(rows) { this.rows = rows; }
  Mat.prototype.h = function () { return this.rows.length; };
  Mat.prototype.w = function () { return this.rows.length ? this.rows[0].length : 0; };

  // ---------- A1 helpers ----------
  function colName(c) {
    var s = ""; c = c + 1;
    while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }
  function colIndex(s) {
    var n = 0; s = s.toUpperCase();
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1;
  }
  function addr(r, c, absR, absC) { return (absC ? "$" : "") + colName(c) + (absR ? "$" : "") + (r + 1); }
  function parseAddr(s) {
    var m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(String(s).trim());
    if (!m) return null;
    var c = colIndex(m[1]), r = +m[2] - 1;
    if (c < 0 || c >= MAXC || r < 0 || r >= MAXR) return null;
    return { r: r, c: c };
  }
  function needsQuote(name) { return !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) || /^[A-Za-z]{1,3}\d+$/.test(name) || /^(TRUE|FALSE)$/i.test(name); }
  function quoteSheet(name) { return needsQuote(name) ? "'" + name.replace(/'/g, "''") + "'" : name; }
  // "Sheet 1!A1:B5", "A1", "B:B", "3:5" -> {sheet, r0, c0, r1, c1} (sheet may be null)
  function parseRange(s) {
    s = String(s || "").trim(); var sheet = null, m;
    if ((m = /^'((?:[^']|'')+)'!(.*)$/.exec(s))) { sheet = m[1].replace(/''/g, "'"); s = m[2]; }
    else if ((m = /^([^!:'\s][^!']*)!(.*)$/.exec(s))) { sheet = m[1]; s = m[2]; }
    var parts = s.split(":"), a, b;
    if (parts.length === 1) { a = parseAddr(parts[0]); return a ? { sheet: sheet, r0: a.r, c0: a.c, r1: a.r, c1: a.c } : null; }
    if (parts.length !== 2) return null;
    if ((a = parseAddr(parts[0])) && (b = parseAddr(parts[1]))) return norm({ sheet: sheet, r0: a.r, c0: a.c, r1: b.r, c1: b.c });
    if ((m = /^\$?([A-Za-z]{1,3})$/.exec(parts[0])) && (b = /^\$?([A-Za-z]{1,3})$/.exec(parts[1])))
      return norm({ sheet: sheet, r0: 0, c0: colIndex(m[1]), r1: MAXR - 1, c1: colIndex(b[1]) });
    if ((m = /^\$?(\d+)$/.exec(parts[0])) && (b = /^\$?(\d+)$/.exec(parts[1])))
      return norm({ sheet: sheet, r0: +m[1] - 1, c0: 0, r1: +b[1] - 1, c1: MAXC - 1 });
    return null;
  }
  function norm(g) {
    var o = { sheet: g.sheet, r0: Math.min(g.r0, g.r1), c0: Math.min(g.c0, g.c1), r1: Math.max(g.r0, g.r1), c1: Math.max(g.c0, g.c1) };
    if (o.r0 < 0 || o.c0 < 0 || o.r1 >= MAXR || o.c1 >= MAXC) return null;
    return o;
  }
  function rangeStr(g, withSheet) {
    var p = withSheet && g.sheet ? quoteSheet(g.sheet) + "!" : "";
    if (g.r0 === 0 && g.r1 === MAXR - 1) return p + colName(g.c0) + ":" + colName(g.c1);
    if (g.c0 === 0 && g.c1 === MAXC - 1) return p + (g.r0 + 1) + ":" + (g.r1 + 1);
    var a = addr(g.r0, g.c0), b = addr(g.r1, g.c1);
    return p + (a === b ? a : a + ":" + b);
  }
  function inRange(g, r, c) { return r >= g.r0 && r <= g.r1 && c >= g.c0 && c <= g.c1; }
  function overlaps(a, b) { return a.r0 <= b.r1 && a.r1 >= b.r0 && a.c0 <= b.c1 && a.c1 >= b.c0; }

  // ---------- misc ----------
  function clamp(v, a, b) { return v < a ? a : v > b ? b : v; }
  function uid(p) { return (p || "x") + Date.now().toString(36) + Math.floor(Math.random() * 1e9).toString(36); }
  function clone(o) { return o == null ? o : JSON.parse(JSON.stringify(o)); }
  // Round to 15 significant digits, the way spreadsheet apps show and compare results.
  function r15(x) { return typeof x === "number" && isFinite(x) && x !== 0 ? +x.toPrecision(15) : x; }

  SH.core = {
    MAXR: MAXR, MAXC: MAXC, KW: KW, Err: Err, E: E, ERR_HELP: ERR_HELP, isErr: isErr, errOf: errOf, Mat: Mat,
    colName: colName, colIndex: colIndex, addr: addr, parseAddr: parseAddr, parseRange: parseRange, rangeStr: rangeStr,
    quoteSheet: quoteSheet, needsQuote: needsQuote, norm: norm, inRange: inRange, overlaps: overlaps,
    clamp: clamp, uid: uid, clone: clone, r15: r15
  };
})();
