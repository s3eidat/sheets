/* Sheets formula engine — written from scratch, no eval.
   tokenize -> parse (precedence: comparison, then &, then + -, then * /, then ^, then %, then unary minus) -> AST (cached per text)
   -> evaluate against a dependency graph. Edits mark only the transitive dependents dirty; dirty cells are
   evaluated in sheet order, pulling dirty precedents on demand (deep chains are deferred, not recursed),
   and a cell met again while it is being computed is a circular reference (#CIRC!).
   Arrays (UNIQUE, SORT, FILTER, A1:A3*2…) spill into the cells below/right, or show #SPILL!. */
(function () {
  "use strict";
  var SH = window.SH, C = SH.core, F = SH.fmt, E = C.E, Err = C.Err, Mat = C.Mat;
  var MAXR = C.MAXR, MAXC = C.MAXC, KW = C.KW, SW = 34359738368;       // global key = uid * 2^35 + r * 16384 + c
  var DIRTY = 1, BUSY = 2, DEPTH = 120;
  var ERRS = ["#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!", "#CIRC!", "#SPILL!", "#CALC!", "#ERROR!"];
  var VOLATILE = { NOW: 1, TODAY: 1, RAND: 1, RANDBETWEEN: 1, OFFSET: 1, INDIRECT: 1 };

  // ---------- tokenizer ----------
  var RX_FN = /^([A-Za-z_][A-Za-z0-9_.]*)\s*\(/;
  var RX_CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?(?![A-Za-z0-9_(.!])/;
  var RX_COLS = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(!])/;
  var RX_ROWS = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![A-Za-z0-9_(.!])/;
  var RX_NUM = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/;
  var RX_SHEETQ = /^'((?:[^']|'')+)'!/;
  var RX_SHEET = /^([A-Za-z_À-￿][A-Za-z0-9_.À-￿]*)!/;
  var RX_ID = /^[A-Za-z_À-￿][A-Za-z0-9_.À-￿]*/;

  function okRC(r, c) { return r >= 0 && r < MAXR && c >= 0 && c < MAXC; }
  function matchRef(body) {
    var m;
    if ((m = RX_CELL.exec(body))) {
      var c0 = C.colIndex(m[2]), r0 = +m[4] - 1;
      if (!okRC(r0, c0)) return null;
      if (!m[6]) return { kind: "cell", r0: r0, c0: c0, r1: r0, c1: c0, a: [!!m[3], !!m[1], !!m[3], !!m[1]], len: m[0].length };
      var c1 = C.colIndex(m[6]), r1 = +m[8] - 1;
      if (!okRC(r1, c1)) return null;
      return { kind: "area", r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1), a: [!!m[3], !!m[1], !!m[7], !!m[5]], len: m[0].length };
    }
    if ((m = RX_COLS.exec(body))) {
      var a = C.colIndex(m[2]), b = C.colIndex(m[4]);
      if (a >= MAXC || b >= MAXC) return null;
      return { kind: "cols", r0: 0, c0: Math.min(a, b), r1: MAXR - 1, c1: Math.max(a, b), a: [false, !!m[1], false, !!m[3]], len: m[0].length };
    }
    if ((m = RX_ROWS.exec(body))) {
      var p = +m[2] - 1, q = +m[4] - 1;
      if (p < 0 || q < 0 || p >= MAXR || q >= MAXR) return null;
      return { kind: "rows", r0: Math.min(p, q), c0: 0, r1: Math.max(p, q), c1: MAXC - 1, a: [!!m[1], false, !!m[3], false], len: m[0].length };
    }
    return null;
  }
  // lenient: stop at the first bad character instead of throwing (used for highlighting while typing)
  function tokenize(src, lenient) {
    var toks = [], i = 0, n = src.length, m;
    while (i < n) {
      var ch = src[i];
      if (ch === " " || ch === "\n" || ch === "\r" || ch === "\t" || ch === " ") { i++; continue; }
      var st = i, rest = src.slice(i);
      if (ch === '"') {
        var s = "", j = i + 1, closed = false;
        while (j < n) { if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } closed = true; break; } s += src[j++]; }
        if (!closed) { if (lenient) { toks.push({ t: "str", v: s, s: st, e: n }); break; } throw E["#ERROR!"]; }
        toks.push({ t: "str", v: s, s: st, e: j + 1 }); i = j + 1; continue;
      }
      if (ch === "#") {
        var up = rest.toUpperCase(), hit = null;
        for (var x = 0; x < ERRS.length; x++) if (up.indexOf(ERRS[x]) === 0) { hit = ERRS[x]; break; }
        if (!hit) { if (lenient) break; throw E["#ERROR!"]; }
        toks.push({ t: "err", v: hit, s: st, e: i + hit.length }); i += hit.length; continue;
      }
      if ((m = RX_FN.exec(rest)) && !/!$/.test(m[1])) { toks.push({ t: "fn", v: m[1], s: st, e: st + m[1].length }); i += m[1].length; continue; }
      var sheet = null, sp = 0;
      if (ch === "'") { m = RX_SHEETQ.exec(rest); if (!m) { if (lenient) break; throw E["#ERROR!"]; } sheet = m[1].replace(/''/g, "'"); sp = m[0].length; }
      else if ((m = RX_SHEET.exec(rest))) { sheet = m[1]; sp = m[0].length; }
      var ref = matchRef(sp ? rest.slice(sp) : rest);
      if (ref) { ref.t = "ref"; ref.sheet = sheet; ref.s = st; ref.ss = st + sp; ref.e = st + sp + ref.len; toks.push(ref); i = ref.e; continue; }
      if (sp) { if (lenient) break; throw E["#ERROR!"]; }
      if ((m = RX_NUM.exec(rest))) { toks.push({ t: "num", v: parseFloat(m[0]), s: st, e: st + m[0].length }); i += m[0].length; continue; }
      if ((m = RX_ID.exec(rest))) {
        var u = m[0].toUpperCase();
        toks.push(u === "TRUE" || u === "FALSE" ? { t: "bool", v: u === "TRUE", s: st, e: st + m[0].length } : { t: "name", v: m[0], s: st, e: st + m[0].length });
        i += m[0].length; continue;
      }
      var two = src.substr(i, 2);
      if (two === "<=" || two === ">=" || two === "<>") { toks.push({ t: "op", v: two, s: st, e: i + 2 }); i += 2; continue; }
      if ("+-*/^&=<>%(),;{}".indexOf(ch) >= 0) { toks.push({ t: "op", v: ch, s: st, e: i + 1 }); i++; continue; }
      if (lenient) break;
      throw E["#ERROR!"];
    }
    return toks;
  }

  // ---------- parser ----------
  var CMP = { "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1 };
  function parse(src) {
    var toks = tokenize(src), k = 0;
    function isOp(v) { var t = toks[k]; return t && t.t === "op" && t.v === v; }
    function expect(v) { if (!isOp(v)) throw E["#ERROR!"]; k++; }
    function cmp() { var a = cat(); while (toks[k] && toks[k].t === "op" && CMP[toks[k].v]) { var op = toks[k++].v; a = { t: "bin", op: op, a: a, b: cat() }; } return a; }
    function cat() { var a = add(); while (isOp("&")) { k++; a = { t: "bin", op: "&", a: a, b: add() }; } return a; }
    function add() { var a = mul(); while (isOp("+") || isOp("-")) { var op = toks[k++].v; a = { t: "bin", op: op, a: a, b: mul() }; } return a; }
    function mul() { var a = pow(); while (isOp("*") || isOp("/")) { var op = toks[k++].v; a = { t: "bin", op: op, a: a, b: pow() }; } return a; }
    function pow() { var a = post(); while (isOp("^")) { k++; a = { t: "bin", op: "^", a: a, b: post() }; } return a; }
    function post() { var a = un(); while (isOp("%")) { k++; a = { t: "pct", a: a }; } return a; }
    function un() {
      if (isOp("-")) { k++; return { t: "neg", a: un() }; }
      if (isOp("+")) { k++; return un(); }
      return prim();
    }
    function constant() {
      var t = toks[k], neg = false;
      if (isOp("-")) { neg = true; k++; t = toks[k]; }
      if (!t) throw E["#ERROR!"];
      k++;
      if (t.t === "num") return neg ? -t.v : t.v;
      if (neg) throw E["#ERROR!"];
      if (t.t === "str" || t.t === "bool") return t.v;
      if (t.t === "err") return C.errOf(t.v);
      throw E["#ERROR!"];
    }
    function prim() {
      var t = toks[k];
      if (!t) throw E["#ERROR!"];
      if (t.t === "num" || t.t === "str" || t.t === "bool") { k++; return { t: t.t, v: t.v }; }
      if (t.t === "err") { k++; return { t: "err", v: t.v }; }
      if (t.t === "ref") { k++; return { t: "ref", sheet: t.sheet, kind: t.kind, r0: t.r0, c0: t.c0, r1: t.r1, c1: t.c1 }; }
      if (t.t === "name") { k++; return { t: "name", v: t.v }; }
      if (t.t === "fn") {
        k++; expect("(");
        var args = [];
        if (isOp(")")) { k++; return { t: "fn", name: fname(t.v), args: args }; }
        for (;;) {
          if (isOp(",") || isOp(";") || isOp(")")) args.push({ t: "empty" }); else args.push(cmp());
          if (isOp(",") || isOp(";")) { k++; continue; }
          expect(")"); break;
        }
        return { t: "fn", name: fname(t.v), args: args };
      }
      if (isOp("(")) { k++; var e = cmp(); expect(")"); return e; }
      if (isOp("{")) {
        k++; var rows = [[]];
        for (;;) {
          rows[rows.length - 1].push(constant());
          if (isOp(",")) { k++; continue; }
          if (isOp(";")) { k++; rows.push([]); continue; }
          expect("}"); break;
        }
        var w = rows[0].length;
        if (rows.some(function (r) { return r.length !== w; })) throw E["#ERROR!"];
        return { t: "arr", v: new Mat(rows) };
      }
      throw E["#ERROR!"];
    }
    if (!toks.length) throw E["#ERROR!"];
    var ast = cmp();
    if (k < toks.length) throw E["#ERROR!"];
    return ast;
  }
  function fname(n) { return n.toUpperCase().replace(/^(_XLFN\.|_XLWS\.)+/, ""); }
  var astCache = new Map();
  function parseCached(src) {
    var a = astCache.get(src);
    if (a !== undefined) return a;
    try { a = parse(src); } catch (e) { a = e instanceof Err ? e : E["#ERROR!"]; }
    if (astCache.size > 20000) astCache.clear();
    astCache.set(src, a);
    return a;
  }

  // ---------- formula text rewriting (copy/fill, insert/delete rows, rename sheet, file import/export) ----------
  function rewrite(f, fn) {
    var toks;
    try { toks = tokenize(f); } catch (e) { return f; }
    var out = "", pos = 0;
    toks.forEach(function (t) {
      var rep = fn(t);
      if (rep != null) { out += f.slice(pos, t.s) + rep; pos = t.e; }
    });
    return out + f.slice(pos);
  }
  function cellTxt(r, c, ar, ac) { return (ac ? "$" : "") + C.colName(c) + (ar ? "$" : "") + (r + 1); }
  function refBody(kind, r0, c0, r1, c1, a) {
    if (kind === "cell") return cellTxt(r0, c0, a[0], a[1]);
    if (kind === "area") return cellTxt(r0, c0, a[0], a[1]) + ":" + cellTxt(r1, c1, a[2], a[3]);
    if (kind === "cols") return (a[1] ? "$" : "") + C.colName(c0) + ":" + (a[3] ? "$" : "") + C.colName(c1);
    return (a[0] ? "$" : "") + (r0 + 1) + ":" + (a[2] ? "$" : "") + (r1 + 1);
  }
  function prefix(f, t) { return f.slice(t.s, t.ss); }
  // relative parts move by (dr, dc); a ref pushed off the grid becomes #REF!
  function moveFormula(f, dr, dc) {
    if (!dr && !dc) return f;
    return rewrite(f, function (t) {
      if (t.t !== "ref") return null;
      var a = t.a, r0 = t.r0, c0 = t.c0, r1 = t.r1, c1 = t.c1;
      if (t.kind !== "cols") { if (!a[0]) r0 += dr; if (!a[2]) r1 += dr; }
      if (t.kind !== "rows") { if (!a[1]) c0 += dc; if (!a[3]) c1 += dc; }
      if (!okRC(r0, c0) || !okRC(r1, c1)) return "#REF!";
      return prefix(f, t) + refBody(t.kind, r0, c0, r1, c1, a);
    });
  }
  function sameName(a, b) { return String(a).toLowerCase() === String(b).toLowerCase(); }
  // insert (n > 0) or delete (n < 0) rows/cols at `at` on sheet `target`
  function shiftFormula(f, host, target, axis, at, n) {
    return rewrite(f, function (t) {
      if (t.t !== "ref" || !sameName(t.sheet || host, target)) return null;
      if ((axis === "r" && t.kind === "cols") || (axis === "c" && t.kind === "rows")) return null;
      var p0 = axis === "r" ? t.r0 : t.c0, p1 = axis === "r" ? t.r1 : t.c1, max = axis === "r" ? MAXR : MAXC;
      if (n > 0) { if (p0 >= at) p0 += n; if (p1 >= at) p1 += n; if (p1 >= max) { if (p0 >= max) return "#REF!"; p1 = max - 1; } }
      else {
        var d0 = at, d1 = at - n - 1;
        if (p0 >= d0 && p1 <= d1) return "#REF!";
        p0 = p0 < d0 ? p0 : p0 > d1 ? p0 + n : d0;
        p1 = p1 < d0 ? p1 : p1 > d1 ? p1 + n : d0 - 1;
      }
      var r0 = t.r0, c0 = t.c0, r1 = t.r1, c1 = t.c1;
      if (axis === "r") { r0 = p0; r1 = p1; } else { c0 = p0; c1 = p1; }
      if (t.kind === "cell" && (r0 !== r1 || c0 !== c1)) return "#REF!";
      return prefix(f, t) + refBody(t.kind, r0, c0, r1, c1, t.a);
    });
  }
  function renameSheetIn(f, oldName, newName) {
    return rewrite(f, function (t) { return t.t === "ref" && t.sheet && sameName(t.sheet, oldName) ? C.quoteSheet(newName) + "!" + f.slice(t.ss, t.e) : null; });
  }
  function dropSheetIn(f, name) { return rewrite(f, function (t) { return t.t === "ref" && t.sheet && sameName(t.sheet, name) ? "#REF!" : null; }); }
  // tidy what was typed: upper-case names and refs, close open brackets
  function normalize(f) {
    var toks;
    try { toks = tokenize(f, true); } catch (e) { return f; }
    var depth = 0;
    toks.forEach(function (t) { if (t.t === "op" && t.v === "(") depth++; if (t.t === "op" && t.v === ")") depth--; });
    var g = rewrite(f, function (t) {
      if (t.t === "fn") return t.v.toUpperCase();
      if (t.t === "ref") return prefix(f, t) + refBody(t.kind, t.r0, t.c0, t.r1, t.c1, t.a);
      if (t.t === "bool") return t.v ? "TRUE" : "FALSE";
      return null;
    });
    if (depth > 0 && depth < 20) g += ")".repeat(depth);
    return g;
  }
  var NEWFN = "XLOOKUP XMATCH IFS SWITCH CONCAT TEXTJOIN MAXIFS MINIFS IFNA UNIQUE SORT FILTER SEQUENCE STDEV.S STDEV.P VAR.S VAR.P MODE.SNGL RANK.EQ CEILING.MATH FLOOR.MATH DAYS TEXTBEFORE TEXTAFTER XOR UNICHAR UNICODE".split(" ");
  function toFile(f) { return rewrite(f, function (t) { return t.t === "fn" && NEWFN.indexOf(fname(t.v)) >= 0 ? "_xlfn." + fname(t.v) : null; }); }
  function fromFile(f) { return rewrite(f, function (t) { return t.t === "fn" && /^_xl/i.test(t.v) ? fname(t.v) : null; }); }

  // ---------- values & coercion ----------
  function Rng(cx, sh, r0, c0, r1, c1) { this.cx = cx; this.sh = sh; this.r0 = r0; this.c0 = c0; this.r1 = r1; this.c1 = c1; }
  Rng.prototype.single = function () { return this.r0 === this.r1 && this.c0 === this.c1; };
  Rng.prototype.get = function (i, j) { return this.cx.eng.get(this.sh, this.r0 + i, this.c0 + j, this.cx); };
  Rng.prototype.H = function () { return this.r1 - this.r0 + 1; };          // full size
  Rng.prototype.W = function () { return this.c1 - this.c0 + 1; };
  Rng.prototype.h = function () { return Math.max(1, Math.min(this.r1, usedR(this.sh)) - this.r0 + 1); };   // clamped to data
  Rng.prototype.w = function () { return Math.max(1, Math.min(this.c1, usedC(this.sh)) - this.c0 + 1); };
  function usedR(sh) { return Math.max(sh.maxR, sh.smaxR || -1); }
  function usedC(sh) { return Math.max(sh.maxC, sh.smaxC || -1); }

  function toMat(x) {
    if (x instanceof Mat) return x;
    if (x instanceof Rng) {
      var h = x.h(), w = x.w(), rows = new Array(h);
      for (var i = 0; i < h; i++) { var row = rows[i] = new Array(w); for (var j = 0; j < w; j++) row[j] = x.get(i, j); }
      return new Mat(rows);
    }
    return new Mat([[x === undefined ? null : x]]);
  }
  function deref(x) {
    if (x instanceof Rng) return x.single() ? x.get(0, 0) : toMat(x);
    if (x instanceof Mat && x.h() === 1 && x.w() === 1) return x.rows[0][0];
    return x;
  }
  // a single value from an argument; a multi-cell range in a scalar slot uses the host's row/column (implicit intersection)
  function S(x) {
    if (x instanceof Rng) {
      if (x.single()) return x.get(0, 0);
      var cx = x.cx;
      if (x.c0 === x.c1 && cx.r >= x.r0 && cx.r <= x.r1 && cx.sh === x.sh) return x.get(cx.r - x.r0, 0);
      if (x.r0 === x.r1 && cx.c >= x.c0 && cx.c <= x.c1 && cx.sh === x.sh) return x.get(0, cx.c - x.c0);
      throw E["#VALUE!"];
    }
    if (x instanceof Mat) return x.rows[0][0];
    return x;
  }
  function num(v) {
    if (typeof v === "number") return v;
    if (v === null || v === undefined) return 0;
    if (typeof v === "boolean") return v ? 1 : 0;
    if (v instanceof Err) throw v;
    if (typeof v === "string") {
      if (v.trim() === "") throw E["#VALUE!"];
      var p = F.parseInput(v);
      if (typeof p.v === "number" && !p.text) return p.v;
    }
    throw E["#VALUE!"];
  }
  function str(v) {
    if (typeof v === "string") return v;
    if (v === null || v === undefined) return "";
    if (typeof v === "number") return F.numText(v);
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (v instanceof Err) throw v;
    return String(v);
  }
  function bool(v) {
    if (typeof v === "boolean") return v;
    if (typeof v === "number") return v !== 0;
    if (v === null || v === undefined) return false;
    if (v instanceof Err) throw v;
    if (typeof v === "string") { var u = v.toUpperCase(); if (u === "TRUE") return true; if (u === "FALSE") return false; }
    throw E["#VALUE!"];
  }
  function rank(v) { return typeof v === "number" ? 0 : typeof v === "string" ? 1 : typeof v === "boolean" ? 2 : 3; }
  // spreadsheet ordering: numbers < text < logicals; text ignores case; blank acts as 0 / "" / FALSE
  function cmp(a, b) {
    if (a instanceof Err) throw a; if (b instanceof Err) throw b;
    if (a === null || a === undefined) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
    if (b === null || b === undefined) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
    var ra = rank(a), rb = rank(b);
    if (ra !== rb) return ra < rb ? -1 : 1;
    if (ra === 0) { if (C.r15(a) === C.r15(b)) return 0; return a < b ? -1 : 1; }
    if (ra === 1) { var x = a.toLowerCase(), y = b.toLowerCase(); return x === y ? 0 : x < y ? -1 : 1; }
    return a === b ? 0 : a ? 1 : -1;
  }
  function chk(x) { if (typeof x === "number" && !isFinite(x)) throw E["#NUM!"]; return x; }
  function scalarOp(op, a, b) {
    switch (op) {
      case "+": return chk(num(a) + num(b));
      case "-": return chk(num(a) - num(b));
      case "*": return chk(num(a) * num(b));
      case "/": { var x = num(a), y = num(b); if (y === 0) throw E["#DIV/0!"]; return chk(x / y); }
      case "^": { var p = num(a), q = num(b); if (p === 0 && q <= 0) throw p === 0 && q === 0 ? E["#NUM!"] : E["#DIV/0!"]; var r = Math.pow(p, q); if (isNaN(r)) throw E["#NUM!"]; return chk(r); }
      case "&": return str(a) + str(b);
      case "=": return cmp(a, b) === 0;
      case "<>": return cmp(a, b) !== 0;
      case "<": return cmp(a, b) < 0;
      case ">": return cmp(a, b) > 0;
      case "<=": return cmp(a, b) <= 0;
      case ">=": return cmp(a, b) >= 0;
    }
    throw E["#VALUE!"];
  }
  function safe(fn) { try { return fn(); } catch (e) { if (e instanceof Err) return e; throw e; } }
  function at(m, i, j) {
    var H = m.rows.length, W = H ? m.rows[0].length : 0;
    if (H === 1) i = 0; if (W === 1) j = 0;
    if (i >= H || j >= W) return E["#N/A"];
    return m.rows[i][j];
  }
  // element-wise over arrays with broadcasting (a 1x1 or a single row/column stretches)
  function lift(fn, list) {
    var H = 1, W = 1, any = false;
    list.forEach(function (x) { if (x instanceof Mat) { any = true; H = Math.max(H, x.h()); W = Math.max(W, x.w()); } });
    if (!any) return fn.apply(null, list);
    var rows = new Array(H);
    for (var i = 0; i < H; i++) {
      var row = rows[i] = new Array(W);
      for (var j = 0; j < W; j++) {
        var args = list.map(function (x) { return x instanceof Mat ? at(x, i, j) : x; });
        row[j] = safe(function () { return fn.apply(null, args); });
      }
    }
    return new Mat(rows);
  }

  // ---------- evaluator ----------
  function ev(node, cx) {
    switch (node.t) {
      case "num": case "str": case "bool": return node.v;
      case "err": throw C.errOf(node.v);
      case "empty": return null;
      case "arr": return node.v;
      case "name": throw E["#NAME?"];
      case "ref": {
        var sh = node.sheet ? cx.eng.wb.sheetByName(node.sheet) : cx.sh;
        if (!sh) throw E["#REF!"];
        return new Rng(cx, sh, node.r0, node.c0, node.r1, node.c1);
      }
      case "neg": return lift(function (v) { return -num(v); }, [deref(ev(node.a, cx))]);
      case "pct": return lift(function (v) { return num(v) / 100; }, [deref(ev(node.a, cx))]);
      case "bin": {
        var a = deref(ev(node.a, cx)), b = deref(ev(node.b, cx)), op = node.op;
        if (a instanceof Mat || b instanceof Mat) return lift(function (x, y) { return scalarOp(op, x, y); }, [a, b]);
        return scalarOp(op, a, b);
      }
      case "fn": return call(node, cx);
    }
    throw E["#VALUE!"];
  }
  function call(node, cx) {
    var d = SH.FN[node.name];
    if (!d) throw E["#NAME?"];
    var n = node.args.length;
    if (n < d.min || n > d.max) throw E["#N/A"];
    if (d.lazy) return d.f(node.args, cx);
    var args = new Array(n);
    for (var i = 0; i < n; i++) args[i] = node.args[i].t === "empty" ? null : ev(node.args[i], cx);
    if (d.lift) {
      var dl = args.map(function (x) { return x instanceof Rng ? deref(x) : x instanceof Mat && x.h() === 1 && x.w() === 1 ? x.rows[0][0] : x; });
      if (dl.some(function (x) { return x instanceof Mat; })) return lift(function () { return d.f(Array.prototype.slice.call(arguments), cx); }, dl);
      return d.f(dl, cx);
    }
    return d.f(args, cx);
  }

  // ---------- engine: values, dependency graph, recalculation ----------
  function Defer(k) { this.k = k; }
  function Engine(wb) { this.wb = wb; this.reset(); }
  Engine.prototype.reset = function () {
    this.vals = new Map(); this.st = new Map(); this.deps = new Map(); this.rev = new Map();
    this.colIdx = new Map(); this.wide = new Set(); this.vol = new Set(); this.spill = new Map(); this.anchors = new Map();
    this.changedSpill = []; this.ver = (this.ver || 0) + 1;
  };
  Engine.prototype.gk = function (sh, r, c) { return sh.uid * SW + r * KW + c; };
  function split(k) { var uid = Math.floor(k / SW), rem = k - uid * SW; return { uid: uid, r: Math.floor(rem / KW), c: rem % KW }; }

  Engine.prototype.register = function (k, sh, ast) {
    var self = this, dep = { cells: [], ranges: [] };
    (function walk(n) {
      if (!n || typeof n !== "object") return;
      if (n.t === "ref") {
        var s = n.sheet ? self.wb.sheetByName(n.sheet) : sh;
        if (!s) return;
        if (n.kind === "cell") dep.cells.push(self.gk(s, n.r0, n.c0));
        else dep.ranges.push([s.uid, n.r0, n.c0, n.r1, n.c1]);
        return;
      }
      if (n.t === "fn") { if (VOLATILE[n.name]) dep.vol = true; n.args.forEach(walk); return; }
      if (n.a) walk(n.a); if (n.b) walk(n.b);
    })(ast);
    this.deps.set(k, dep);
    dep.cells.forEach(function (p) { var s = self.rev.get(p); if (!s) self.rev.set(p, s = new Set()); s.add(k); });
    dep.ranges.forEach(function (g) {
      if (g[4] - g[2] > 64) { self.wide.add(k); return; }
      for (var c = g[2]; c <= g[4]; c++) { var key = g[0] * KW + c, b = self.colIdx.get(key); if (!b) self.colIdx.set(key, b = new Set()); b.add(k); }
    });
    if (dep.vol) this.vol.add(k);
  };
  Engine.prototype.unregister = function (k) {
    var dep = this.deps.get(k), self = this;
    if (!dep) return;
    dep.cells.forEach(function (p) { var s = self.rev.get(p); if (s) { s.delete(k); if (!s.size) self.rev.delete(p); } });
    dep.ranges.forEach(function (g) {
      if (g[4] - g[2] > 64) { self.wide.delete(k); return; }
      for (var c = g[2]; c <= g[4]; c++) { var key = g[0] * KW + c, b = self.colIdx.get(key); if (b) { b.delete(k); if (!b.size) self.colIdx.delete(key); } }
    });
    this.vol.delete(k); this.deps.delete(k);
  };
  Engine.prototype.inRanges = function (fk, uid, r, c) {
    var d = this.deps.get(fk);
    if (!d) return false;
    for (var i = 0; i < d.ranges.length; i++) { var g = d.ranges[i]; if (g[0] === uid && r >= g[1] && r <= g[3] && c >= g[2] && c <= g[4]) return true; }
    return false;
  };
  // every formula that (transitively) depends on the seeds
  Engine.prototype.collect = function (seeds) {
    var self = this, seen = new Set(), dirty = new Set(), stack = seeds.slice();
    this.vol.forEach(function (k) { stack.push(k); });
    while (stack.length) {
      var k = stack.pop();
      if (seen.has(k)) continue;
      seen.add(k);
      if (this.deps.has(k)) dirty.add(k);
      var d = this.rev.get(k);
      if (d) d.forEach(function (x) { if (!seen.has(x)) stack.push(x); });
      var p = split(k), b = this.colIdx.get(p.uid * KW + p.c);
      if (b) b.forEach(function (x) { if (!seen.has(x) && self.inRanges(x, p.uid, p.r, p.c)) stack.push(x); });
      if (this.wide.size) this.wide.forEach(function (x) { if (!seen.has(x) && self.inRanges(x, p.uid, p.r, p.c)) stack.push(x); });
      var an = this.anchors.get(k);
      if (an && an.ok) {
        var sh = this.wb.sheetByUid(an.uid);
        for (var r = an.r0; r <= an.r1; r++) for (var c = an.c0; c <= an.c1; c++) if (r !== an.r0 || c !== an.c0) stack.push(this.gk(sh, r, c));
      }
    }
    return dirty;
  };
  Engine.prototype.rebuild = function () {
    var self = this, all = [];
    this.reset();
    this.wb.sheets.forEach(function (sh) {
      sh.smaxR = -1; sh.smaxC = -1;
      sh.cells.forEach(function (cell, lk) {
        if (cell.f === undefined) return;
        var k = sh.uid * SW + lk;
        self.register(k, sh, parseCached(cell.f));
        all.push(k);
      });
    });
    this.recalc(all);
  };
  // cells whose content changed: [{sh, r, c}]
  Engine.prototype.update = function (list) {
    var self = this, seeds = [];
    list.forEach(function (x) {
      var k = self.gk(x.sh, x.r, x.c), cell = x.sh.cells.get(x.r * KW + x.c);
      self.unregister(k);
      if (cell && cell.f !== undefined) self.register(k, x.sh, parseCached(cell.f));
      else { self.clearSpill(k); self.vals.delete(k); self.st.delete(k); }
      seeds.push(k);
      if (self.anchors.size) self.anchors.forEach(function (an, ak) { if (an.uid === x.sh.uid && C.inRange(an, x.r, x.c) && ak !== k) seeds.push(ak); });
    });
    this.recalc(seeds);
  };
  Engine.prototype.recalc = function (seeds, round) {
    var self = this, dirty = this.collect(seeds), order = Array.from(dirty).sort(function (a, b) { return a - b; });
    order.forEach(function (k) { self.st.set(k, DIRTY); });
    var stack = order.reverse(), guard = 0, limit = order.length * 8 + 2000;
    this.changedSpill = [];
    while (stack.length) {
      var k = stack.pop();
      if (this.st.get(k) !== DIRTY) continue;
      try { this.evalCell(k, 0); }
      catch (e) { if (e instanceof Defer) { stack.push(k, e.k); if (++guard > limit) break; } else throw e; }
    }
    this.st.forEach(function (v, k) { self.vals.set(k, E["#CIRC!"]); });
    this.st.clear();
    this.ver++;
    var ch = this.changedSpill;
    if (ch.length && (round || 0) < 4) this.recalc(ch, (round || 0) + 1);
  };
  Engine.prototype.evalCell = function (k, depth) {
    var p = split(k), sh = this.wb.sheetByUid(p.uid), cell = sh && sh.cells.get(p.r * KW + p.c), res;
    if (!cell || cell.f === undefined) { this.st.delete(k); return; }
    this.st.set(k, BUSY);
    try { res = this.compute(cell.f, sh, p.r, p.c, depth); }
    catch (e) {
      if (e instanceof Defer) { this.st.set(k, DIRTY); throw e; }
      res = e instanceof Err ? e : E["#VALUE!"];
    }
    this.st.delete(k);
    this.store(k, sh, p.r, p.c, res);
  };
  Engine.prototype.compute = function (f, sh, r, c, depth) {
    var ast = parseCached(f);
    if (ast instanceof Err) return ast;
    var cx = { eng: this, sh: sh, r: r, c: c, depth: depth || 0 };
    var x = ev(ast, cx);
    if (x instanceof Rng) x = x.single() ? x.get(0, 0) : toMat(x);
    if (x instanceof Mat) { if (!x.h() || !x.w()) return E["#CALC!"]; if (x.h() === 1 && x.w() === 1) x = x.rows[0][0]; }
    if (x === null || x === undefined) return 0;
    return typeof x === "number" ? C.r15(x) : x;
  };
  Engine.prototype.get = function (sh, r, c, cx) {
    var k = sh.uid * SW + r * KW + c, s = this.st.get(k);
    if (s === DIRTY) { if (cx && cx.depth > DEPTH) throw new Defer(k); this.evalCell(k, (cx ? cx.depth : 0) + 1); }
    else if (s === BUSY) throw E["#CIRC!"];
    var cell = sh.cells.get(r * KW + c);
    if (cell) {
      if (cell.f !== undefined) { var v = this.vals.get(k); return v === undefined ? null : v; }
      if (cell.v !== undefined) return cell.v;
    }
    if (this.spill.size) {
      var a = this.spill.get(k);
      if (a !== undefined) {
        var sa = this.st.get(a);
        if (sa === DIRTY) { if (cx && cx.depth > DEPTH) throw new Defer(a); this.evalCell(a, (cx ? cx.depth : 0) + 1); }
        else if (sa === BUSY) throw E["#CIRC!"];
        var w = this.vals.get(k);
        return w === undefined ? null : w;
      }
    }
    return null;
  };
  // value for display (no evaluation)
  Engine.prototype.value = function (sh, r, c) {
    var cell = sh.cells.get(r * KW + c);
    if (cell) {
      if (cell.f !== undefined) { var v = this.vals.get(sh.uid * SW + r * KW + c); return v === undefined ? null : v; }
      if (cell.v !== undefined) return cell.v;
    }
    if (this.spill.size) { var k = sh.uid * SW + r * KW + c; if (this.spill.has(k)) { var w = this.vals.get(k); return w === undefined ? null : w; } }
    return null;
  };
  Engine.prototype.spillOf = function (sh, r, c) { var a = this.spill.get(sh.uid * SW + r * KW + c); return a === undefined ? null : split(a); };
  Engine.prototype.clearSpill = function (k) {
    var an = this.anchors.get(k);
    if (!an) return;
    if (an.ok) {
      var sh = this.wb.sheetByUid(an.uid);
      for (var r = an.r0; r <= an.r1; r++) for (var c = an.c0; c <= an.c1; c++) {
        if (r === an.r0 && c === an.c0) continue;
        var k2 = this.gk(sh, r, c);
        if (this.spill.get(k2) === k) { this.spill.delete(k2); this.vals.delete(k2); this.changedSpill.push(k2); }
      }
    }
    this.anchors.delete(k);
  };
  function same(a, b) { return a === b || (a instanceof Err && b instanceof Err && a.e === b.e); }
  Engine.prototype.store = function (k, sh, r, c, res) {
    if (!(res instanceof Mat)) { this.clearSpill(k); this.vals.set(k, res); return; }
    var h = res.h(), w = res.w(), rect = { uid: sh.uid, r0: r, c0: c, r1: r + h - 1, c1: c + w - 1, ok: false };
    var old = this.anchors.get(k), oldVals = new Map(), self = this;
    if (old && old.ok) for (var r2 = old.r0; r2 <= old.r1; r2++) for (var c2 = old.c0; c2 <= old.c1; c2++) { var kk = this.gk(sh, r2, c2); if (kk !== k) oldVals.set(kk, this.vals.get(kk)); }
    var ch = this.changedSpill; this.changedSpill = [];
    this.clearSpill(k);
    this.changedSpill = ch;
    var blocked = rect.r1 >= MAXR || rect.c1 >= MAXC || h * w > 2e6;
    for (var i = 0; i < h && !blocked; i++) for (var j = 0; j < w; j++) {
      if (!i && !j) continue;
      var cl = sh.cells.get((r + i) * KW + c + j);
      if (cl && (cl.f !== undefined || (cl.v !== undefined && cl.v !== null))) { blocked = true; break; }
      var other = this.spill.get(this.gk(sh, r + i, c + j));
      if (other !== undefined && other !== k) { blocked = true; break; }
    }
    if (blocked) {
      this.anchors.set(k, rect); this.vals.set(k, E["#SPILL!"]);
      oldVals.forEach(function (v, kk) { self.changedSpill.push(kk); });
      return;
    }
    rect.ok = true; this.anchors.set(k, rect);
    for (var a = 0; a < h; a++) for (var b = 0; b < w; b++) {
      var v = res.rows[a][b]; if (v === null || v === undefined) v = 0; else if (typeof v === "number") v = C.r15(v);
      if (!a && !b) { this.vals.set(k, v); continue; }
      var k2 = this.gk(sh, r + a, c + b);
      this.spill.set(k2, k); this.vals.set(k2, v);
      if (!oldVals.has(k2) || !same(oldVals.get(k2), v)) this.changedSpill.push(k2);
      oldVals.delete(k2);
    }
    oldVals.forEach(function (v, kk) { self.changedSpill.push(kk); });
    if (rect.r1 > (sh.smaxR || -1)) sh.smaxR = rect.r1;
    if (rect.c1 > (sh.smaxC || -1)) sh.smaxC = rect.c1;
  };
  // evaluate a formula once without storing it (self-test, conditional-format rules)
  Engine.prototype.evalFormula = function (f, sh, r, c) {
    try { return this.compute(f, sh, r || 0, c || 0, 0); } catch (e) { return e instanceof Err ? e : E["#VALUE!"]; }
  };

  SH.engine = {
    Engine: Engine, tokenize: tokenize, parse: parse, parseCached: parseCached, rewrite: rewrite, moveFormula: moveFormula,
    shiftFormula: shiftFormula, renameSheetIn: renameSheetIn, dropSheetIn: dropSheetIn, normalize: normalize,
    toFile: toFile, fromFile: fromFile, refBody: refBody, SW: SW
  };
  SH.calc = {
    Rng: Rng, toMat: toMat, deref: deref, S: S, num: num, str: str, bool: bool, cmp: cmp, chk: chk, lift: lift,
    safe: safe, ev: ev, scalarOp: scalarOp, usedR: usedR, usedC: usedC
  };
})();
