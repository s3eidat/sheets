/* Sheets function library: 130+ everyday spreadsheet functions with the usual semantics
   (ranges skip text, direct arguments coerce, errors propagate, criteria like ">5" and "a*"). */
(function () {
  "use strict";
  var SH = window.SH, C = SH.core, F = SH.fmt, K = SH.calc, E = C.E, Err = C.Err, Mat = C.Mat, Rng = K.Rng;
  var num = K.num, str = K.str, bool = K.bool, S = K.S, cmp = K.cmp, chk = K.chk, toMat = K.toMat, deref = K.deref, ev = K.ev;
  var FN = SH.FN = {}, DOC = SH.FNDOC = {};

  function def(names, min, max, f, opt, doc) {
    names.split(" ").forEach(function (n) {
      var d = { min: min, max: max, f: f };
      if (opt) for (var k in opt) d[k] = opt[k];
      FN[n] = d;
      if (doc) DOC[n] = doc;
    });
  }
  var L = { lift: true }, LZ = { lazy: true }, VOL = { vol: true };
  function n(a) { return num(S(a)); }
  function s(a) { return str(S(a)); }
  function b(a) { return bool(S(a)); }
  function opt(a, d) { return a === undefined || a === null ? d : a; }
  function on(a, d) { return a === undefined || a === null ? d : n(a); }
  function int(a) { return Math.trunc(n(a)); }

  // iterate every value of an argument; fn(value, direct)
  function each(x, fn) {
    if (x instanceof Rng) { var h = x.h(), w = x.w(); for (var i = 0; i < h; i++) for (var j = 0; j < w; j++) fn(x.get(i, j), false); }
    else if (x instanceof Mat) x.rows.forEach(function (r) { r.forEach(function (v) { fn(v, false); }); });
    else fn(x, true);
  }
  // numbers per the usual rules: ranges/arrays keep only numbers; direct args coerce (text "5", TRUE)
  function nums(args) {
    var out = [];
    args.forEach(function (a) {
      if (a === undefined) return;
      each(a, function (v, direct) {
        if (v instanceof Err) throw v;
        if (typeof v === "number") out.push(v);
        else if (direct && v !== null) out.push(num(v));
      });
    });
    return out;
  }
  function sum(a) { var t = 0; for (var i = 0; i < a.length; i++) t += a[i]; return t; }
  function dims(x) { if (x instanceof Rng) return [x.h(), x.w()]; if (x instanceof Mat) return [x.h(), x.w()]; return [1, 1]; }
  function getter(x) {
    if (x instanceof Rng) return function (i, j) { return x.get(i, j); };
    if (x instanceof Mat) return function (i, j) { var r = x.rows[i]; return r && j < r.length ? r[j] : null; };
    return function () { return x; };
  }

  // ---------- criteria ("5", ">=10", "<>", "a*", "~*") ----------
  function wild(p) {
    var rx = "", i = 0;
    for (; i < p.length; i++) {
      var ch = p[i];
      if (ch === "~" && i + 1 < p.length) { rx += p[++i].replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); continue; }
      rx += ch === "*" ? "[\\s\\S]*" : ch === "?" ? "[\\s\\S]" : ch.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
    return new RegExp("^" + rx + "$", "i");
  }
  function hasWild(p) { return /[*?~]/.test(p); }
  function crit(c) {
    if (c instanceof Err) throw c;
    if (typeof c === "number") return function (v) { return typeof v === "number" ? C.r15(v) === C.r15(c) : typeof v === "string" && v.trim() !== "" && numOrNull(v) === c; };
    if (typeof c === "boolean") return function (v) { return v === c; };
    if (c === null || c === undefined) return function (v) { return v === null || v === ""; };
    var t = String(c), m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(t), op = m[1] || "", rest = m[2];
    var rn = rest.trim() === "" ? null : numOrNull(rest);
    if (op === "" || op === "=" || op === "<>") {
      var eq;
      if (rest === "") eq = function (v) { return v === null || v === ""; };
      else if (rn !== null) eq = function (v) { return (typeof v === "number" && C.r15(v) === C.r15(rn)) || (typeof v === "string" && v.toLowerCase() === rest.toLowerCase()); };
      else if (/^(true|false)$/i.test(rest)) { var bv = /^true$/i.test(rest); eq = function (v) { return v === bv; }; }
      else if (hasWild(rest)) { var rx = wild(rest); eq = function (v) { return typeof v === "string" && rx.test(v); }; }
      else { var lo = rest.toLowerCase(); eq = function (v) { return typeof v === "string" && v.toLowerCase() === lo; }; }
      if (op === "<>") { if (rest === "") return function (v) { return !(v === null || v === ""); }; return function (v) { return !eq(v); }; }
      return eq;
    }
    if (rn !== null) return function (v) { if (typeof v !== "number") return false; var d = C.r15(v) === C.r15(rn) ? 0 : v < rn ? -1 : 1; return op === "<" ? d < 0 : op === ">" ? d > 0 : op === "<=" ? d <= 0 : d >= 0; };
    var low = rest.toLowerCase();
    return function (v) { if (typeof v !== "string") return false; var x = v.toLowerCase(), d = x === low ? 0 : x < low ? -1 : 1; return op === "<" ? d < 0 : op === ">" ? d > 0 : op === "<=" ? d <= 0 : d >= 0; };
  }
  function numOrNull(t) { var p = F.parseInput(String(t)); return typeof p.v === "number" && !p.f && !p.text ? p.v : null; }
  // shared driver for *IFS functions: pairs [range, criterion]...; fn(i, j) for every matching cell
  function ifs(pairs, fn) {
    if (!pairs.length) throw E["#VALUE!"];
    var first = pairs[0][0], D = dims(first), H = first instanceof Rng ? first.H() : D[0], W = first instanceof Rng ? first.W() : D[1];
    var preds = pairs.map(function (p) {
      var g = p[0];
      if (!(g instanceof Rng) && !(g instanceof Mat)) g = new Mat([[g]]);
      var gh = g instanceof Rng ? g.H() : g.h(), gw = g instanceof Rng ? g.W() : g.w();
      if (gh !== H || gw !== W) throw E["#VALUE!"];
      return { get: getter(g), test: crit(deref(p[1]) instanceof Mat ? deref(p[1]).rows[0][0] : deref(p[1])) };
    });
    var h = D[0], w = D[1];
    for (var i = 0; i < h; i++) for (var j = 0; j < w; j++) {
      var ok = true;
      for (var q = 0; q < preds.length && ok; q++) { var v = preds[q].get(i, j); if (v instanceof Err) { ok = false; break; } ok = preds[q].test(v); }
      if (ok) fn(i, j);
    }
    return { h: h, w: w, H: H, W: W, blankMatch: preds.every(function (p) { return p.test(null); }) };
  }
  function pairsOf(args, from) { var p = []; for (var i = from; i < args.length; i += 2) p.push([args[i], args[i + 1]]); if (args.length - from & 1) throw E["#N/A"]; return p; }
  function sameShape(a, ref) { if (a === undefined || a === null) return ref; if (a instanceof Rng && ref instanceof Rng) return new Rng(a.cx, a.sh, a.r0, a.c0, a.r0 + ref.H() - 1, a.c0 + ref.W() - 1); return a; }

  // ---------- math & aggregates ----------
  def("SUM", 1, 255, function (a) { return sum(nums(a)); }, null, "SUM(number1, [number2], …)|Adds numbers and ranges.");
  def("AVERAGE", 1, 255, function (a) { var x = nums(a); if (!x.length) throw E["#DIV/0!"]; return sum(x) / x.length; }, null, "AVERAGE(number1, [number2], …)|Arithmetic mean.");
  def("MIN", 1, 255, function (a) { var x = nums(a); return x.length ? Math.min.apply(null, x) : 0; }, null, "MIN(number1, …)|Smallest number.");
  def("MAX", 1, 255, function (a) { var x = nums(a); return x.length ? Math.max.apply(null, x) : 0; }, null, "MAX(number1, …)|Largest number.");
  def("PRODUCT", 1, 255, function (a) { var x = nums(a); return x.length ? x.reduce(function (p, v) { return p * v; }, 1) : 0; }, null, "PRODUCT(number1, …)|Multiplies numbers.");
  def("COUNT", 1, 255, function (a) {
    var c = 0;
    a.forEach(function (x) { each(x, function (v, direct) { if (typeof v === "number") c++; else if (direct && (typeof v === "boolean" || (typeof v === "string" && numOrNull(v) !== null))) c++; }); });
    return c;
  }, null, "COUNT(value1, …)|Counts numbers.");
  def("COUNTA", 1, 255, function (a) { var c = 0; a.forEach(function (x) { if (x === undefined) return; each(x, function (v) { if (v !== null) c++; }); }); return c; }, null, "COUNTA(value1, …)|Counts non-empty cells.");
  def("COUNTBLANK", 1, 1, function (a) {
    var x = a[0], filled = 0, total;
    if (x instanceof Rng) { total = x.H() * x.W(); each(x, function (v) { if (v !== null && v !== "") filled++; }); return total - filled; }
    total = dims(x)[0] * dims(x)[1]; each(x, function (v) { if (v !== null && v !== "") filled++; }); return total - filled;
  }, null, "COUNTBLANK(range)|Counts empty cells.");
  def("COUNTIF", 2, 2, function (a) { var c = 0, r = ifs([[a[0], a[1]]], function () { c++; }); if (r.blankMatch) c += r.H * r.W - r.h * r.w; return c; }, null, "COUNTIF(range, criterion)|Counts cells that match, e.g. \">5\" or \"a*\".");
  def("COUNTIFS", 2, 254, function (a) { var c = 0, r = ifs(pairsOf(a, 0), function () { c++; }); if (r.blankMatch) c += r.H * r.W - r.h * r.w; return c; }, null, "COUNTIFS(range1, criterion1, …)|Counts rows matching all criteria.");
  def("SUMIF", 2, 3, function (a) {
    var g = getter(sameShape(a[2], a[0])), t = 0;
    ifs([[a[0], a[1]]], function (i, j) { var v = g(i, j); if (v instanceof Err) throw v; if (typeof v === "number") t += v; });
    return t;
  }, null, "SUMIF(range, criterion, [sum_range])|Adds the cells that match.");
  def("SUMIFS", 3, 255, function (a) {
    var g = getter(a[0]), t = 0;
    ifs(pairsOf(a, 1), function (i, j) { var v = g(i, j); if (v instanceof Err) throw v; if (typeof v === "number") t += v; });
    return t;
  }, null, "SUMIFS(sum_range, range1, criterion1, …)|Adds cells where all criteria match.");
  def("AVERAGEIF", 2, 3, function (a) {
    var g = getter(sameShape(a[2], a[0])), t = 0, c = 0;
    ifs([[a[0], a[1]]], function (i, j) { var v = g(i, j); if (v instanceof Err) throw v; if (typeof v === "number") { t += v; c++; } });
    if (!c) throw E["#DIV/0!"]; return t / c;
  }, null, "AVERAGEIF(range, criterion, [average_range])|Average of matching cells.");
  def("AVERAGEIFS", 3, 255, function (a) {
    var g = getter(a[0]), t = 0, c = 0;
    ifs(pairsOf(a, 1), function (i, j) { var v = g(i, j); if (v instanceof Err) throw v; if (typeof v === "number") { t += v; c++; } });
    if (!c) throw E["#DIV/0!"]; return t / c;
  }, null, "AVERAGEIFS(average_range, range1, criterion1, …)|Average where all criteria match.");
  function mxifs(isMax) {
    return function (a) {
      var g = getter(a[0]), best = null;
      ifs(pairsOf(a, 1), function (i, j) { var v = g(i, j); if (typeof v === "number" && (best === null || (isMax ? v > best : v < best))) best = v; });
      return best === null ? 0 : best;
    };
  }
  def("MAXIFS", 3, 255, mxifs(true), null, "MAXIFS(max_range, range1, criterion1, …)|Largest value where criteria match.");
  def("MINIFS", 3, 255, mxifs(false), null, "MINIFS(min_range, range1, criterion1, …)|Smallest value where criteria match.");
  def("SUMPRODUCT", 1, 255, function (a) {
    var ms = a.map(function (x) { return toMat(deref(x) instanceof Mat ? deref(x) : x instanceof Rng ? x : new Mat([[x]])); });
    var h = ms[0].h(), w = ms[0].w(), t = 0;
    ms.forEach(function (m) { if (m.h() !== h || m.w() !== w) throw E["#VALUE!"]; });
    for (var i = 0; i < h; i++) for (var j = 0; j < w; j++) {
      var p = 1;
      for (var q = 0; q < ms.length; q++) { var v = ms[q].rows[i][j]; if (v instanceof Err) throw v; p *= typeof v === "number" ? v : 0; }
      t += p;
    }
    return t;
  }, null, "SUMPRODUCT(array1, [array2], …)|Multiplies matching items and adds the products.");
  def("SUBTOTAL", 2, 255, function (a, cx) {
    var code = int(a[0]), fnc = code % 100, skipHidden = true;
    var map = { 1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN", 6: "PRODUCT", 7: "STDEV", 8: "STDEVP", 9: "SUM", 10: "VAR", 11: "VARP" };
    if (!map[fnc]) throw E["#VALUE!"];
    var rest = a.slice(1).map(function (x) {
      if (!(x instanceof Rng) || !skipHidden || !x.sh.hidden || !x.sh.hidden.size) return x;
      var rows = [];
      for (var i = 0; i < x.h(); i++) { if (x.sh.hidden.has(x.r0 + i)) continue; var row = []; for (var j = 0; j < x.w(); j++) row.push(x.get(i, j)); rows.push(row); }
      return rows.length ? new Mat(rows) : new Mat([[null]]);
    });
    return FN[map[fnc]].f(rest, cx);
  }, null, "SUBTOTAL(function_num, range1, …)|9 = SUM, 1 = AVERAGE, 2 = COUNT…; skips rows hidden by a filter.");
  function sorted(a) { return nums(a).sort(function (x, y) { return x - y; }); }
  def("MEDIAN", 1, 255, function (a) { var x = sorted(a); if (!x.length) throw E["#NUM!"]; var m = x.length >> 1; return x.length % 2 ? x[m] : (x[m - 1] + x[m]) / 2; }, null, "MEDIAN(number1, …)|Middle value.");
  def("MODE MODE.SNGL", 1, 255, function (a) {
    var x = nums(a), cnt = new Map(), best = null, bc = 1;
    x.forEach(function (v) { var c = (cnt.get(v) || 0) + 1; cnt.set(v, c); if (c > bc) { bc = c; best = v; } });
    if (best === null) throw E["#N/A"]; return best;
  }, null, "MODE(number1, …)|Most frequent value.");
  function variance(x, pop) { var k = x.length; if (k < (pop ? 1 : 2)) throw E["#DIV/0!"]; var m = sum(x) / k, t = 0; x.forEach(function (v) { t += (v - m) * (v - m); }); return t / (pop ? k : k - 1); }
  def("STDEV STDEV.S", 1, 255, function (a) { return Math.sqrt(variance(nums(a), false)); }, null, "STDEV.S(number1, …)|Sample standard deviation.");
  def("STDEVP STDEV.P", 1, 255, function (a) { return Math.sqrt(variance(nums(a), true)); }, null, "STDEV.P(number1, …)|Population standard deviation.");
  def("VAR VAR.S", 1, 255, function (a) { return variance(nums(a), false); }, null, "VAR.S(number1, …)|Sample variance.");
  def("VARP VAR.P", 1, 255, function (a) { return variance(nums(a), true); }, null, "VAR.P(number1, …)|Population variance.");
  def("LARGE", 2, 2, function (a) { var x = sorted([a[0]]), k = int(a[1]); if (k < 1 || k > x.length) throw E["#NUM!"]; return x[x.length - k]; }, null, "LARGE(array, k)|k-th largest value.");
  def("SMALL", 2, 2, function (a) { var x = sorted([a[0]]), k = int(a[1]); if (k < 1 || k > x.length) throw E["#NUM!"]; return x[k - 1]; }, null, "SMALL(array, k)|k-th smallest value.");
  def("RANK RANK.EQ", 2, 3, function (a) {
    var v = n(a[0]), x = nums([a[1]]), asc = on(a[2], 0) !== 0, r = 1, found = false;
    x.forEach(function (y) { if (y === v) found = true; if (asc ? y < v : y > v) r++; });
    if (!found) throw E["#N/A"]; return r;
  }, null, "RANK.EQ(number, ref, [order])|Rank of a number in a list (0 = largest first).");

  function roundTo(x, d, mode) {
    var m = Math.pow(10, d), y = +(Math.abs(x) * m).toPrecision(15), z;
    z = mode === 0 ? Math.round(y) : mode > 0 ? Math.ceil(y) : Math.floor(y);
    return C.r15((x < 0 ? -z : z) / m);
  }
  def("ROUND", 1, 2, function (a) { return roundTo(n(a[0]), int(opt(a[1], 0)), 0); }, L, "ROUND(number, [digits])|Rounds half away from zero.");
  def("ROUNDUP", 1, 2, function (a) { return roundTo(n(a[0]), int(opt(a[1], 0)), 1); }, L, "ROUNDUP(number, [digits])|Rounds away from zero.");
  def("ROUNDDOWN", 1, 2, function (a) { return roundTo(n(a[0]), int(opt(a[1], 0)), -1); }, L, "ROUNDDOWN(number, [digits])|Rounds toward zero.");
  def("TRUNC", 1, 2, function (a) { return roundTo(n(a[0]), int(opt(a[1], 0)), -1); }, L, "TRUNC(number, [digits])|Cuts off decimals.");
  def("INT", 1, 1, function (a) { return Math.floor(n(a[0])); }, L, "INT(number)|Rounds down to an integer.");
  def("MROUND", 2, 2, function (a) { var x = n(a[0]), m = n(a[1]); if (m === 0) return 0; if (x * m < 0) throw E["#NUM!"]; return C.r15(Math.round(+(x / m).toPrecision(15)) * m); }, L, "MROUND(number, multiple)|Rounds to the nearest multiple.");
  function ceilFloor(up) {
    return function (a) {
      var x = n(a[0]), sig = a[1] === undefined ? (x < 0 && !up ? 1 : 1) : n(a[1]);
      if (sig === 0) return 0;
      if (x > 0 && sig < 0) throw E["#NUM!"];
      var q = +(x / sig).toPrecision(15);
      return C.r15((up ? Math.ceil(q) : Math.floor(q)) * sig);
    };
  }
  def("CEILING", 1, 2, ceilFloor(true), L, "CEILING(number, [significance])|Rounds up to a multiple.");
  def("FLOOR", 1, 2, ceilFloor(false), L, "FLOOR(number, [significance])|Rounds down to a multiple.");
  def("CEILING.MATH", 1, 3, function (a) { var x = n(a[0]), sg = Math.abs(on(a[1], 1)) || 1, md = on(a[2], 0); var q = +(x / sg).toPrecision(15); return C.r15((x < 0 && md ? Math.floor(q) : Math.ceil(q)) * sg); }, L, "CEILING.MATH(number, [significance], [mode])|Rounds up to a multiple.");
  def("FLOOR.MATH", 1, 3, function (a) { var x = n(a[0]), sg = Math.abs(on(a[1], 1)) || 1, md = on(a[2], 0); var q = +(x / sg).toPrecision(15); return C.r15((x < 0 && md ? Math.ceil(q) : Math.floor(q)) * sg); }, L, "FLOOR.MATH(number, [significance], [mode])|Rounds down to a multiple.");
  def("ABS", 1, 1, function (a) { return Math.abs(n(a[0])); }, L, "ABS(number)|Absolute value.");
  def("SIGN", 1, 1, function (a) { return Math.sign(n(a[0])); }, L, "SIGN(number)|1, 0 or -1.");
  def("MOD", 2, 2, function (a) { var x = n(a[0]), y = n(a[1]); if (y === 0) throw E["#DIV/0!"]; return C.r15(x - y * Math.floor(x / y)); }, L, "MOD(number, divisor)|Remainder (sign of the divisor).");
  def("QUOTIENT", 2, 2, function (a) { var y = n(a[1]); if (y === 0) throw E["#DIV/0!"]; return Math.trunc(n(a[0]) / y); }, L, "QUOTIENT(numerator, denominator)|Integer part of a division.");
  def("POWER", 2, 2, function (a) { return K.scalarOp("^", n(a[0]), n(a[1])); }, L, "POWER(number, power)|number ^ power.");
  def("SQRT", 1, 1, function (a) { var x = n(a[0]); if (x < 0) throw E["#NUM!"]; return Math.sqrt(x); }, L, "SQRT(number)|Square root.");
  def("EXP", 1, 1, function (a) { return chk(Math.exp(n(a[0]))); }, L, "EXP(number)|e raised to a power.");
  def("LN", 1, 1, function (a) { var x = n(a[0]); if (x <= 0) throw E["#NUM!"]; return Math.log(x); }, L, "LN(number)|Natural logarithm.");
  def("LOG", 1, 2, function (a) { var x = n(a[0]), bs = on(a[1], 10); if (x <= 0 || bs <= 0 || bs === 1) throw E["#NUM!"]; return C.r15(Math.log(x) / Math.log(bs)); }, L, "LOG(number, [base])|Logarithm (base 10 by default).");
  def("LOG10", 1, 1, function (a) { var x = n(a[0]); if (x <= 0) throw E["#NUM!"]; return Math.log10(x); }, L, "LOG10(number)|Base-10 logarithm.");
  def("PI", 0, 0, function () { return Math.PI; }, null, "PI()|3.14159265358979");
  def("FACT", 1, 1, function (a) { var x = Math.floor(n(a[0])), r = 1; if (x < 0 || x > 170) throw E["#NUM!"]; for (var i = 2; i <= x; i++) r *= i; return r; }, L, "FACT(number)|Factorial.");
  def("SIN", 1, 1, function (a) { return Math.sin(n(a[0])); }, L, "SIN(angle)|Sine (radians).");
  def("COS", 1, 1, function (a) { return Math.cos(n(a[0])); }, L, "COS(angle)|Cosine (radians).");
  def("TAN", 1, 1, function (a) { return Math.tan(n(a[0])); }, L, "TAN(angle)|Tangent (radians).");
  def("RADIANS", 1, 1, function (a) { return n(a[0]) * Math.PI / 180; }, L, "RADIANS(degrees)|Degrees to radians.");
  def("DEGREES", 1, 1, function (a) { return n(a[0]) * 180 / Math.PI; }, L, "DEGREES(radians)|Radians to degrees.");
  def("RAND", 0, 0, function () { return Math.random(); }, VOL, "RAND()|Random number 0–1 (changes on every edit).");
  def("RANDBETWEEN", 2, 2, function (a) { var lo = Math.ceil(n(a[0])), hi = Math.floor(n(a[1])); if (lo > hi) throw E["#NUM!"]; return lo + Math.floor(Math.random() * (hi - lo + 1)); }, VOL, "RANDBETWEEN(bottom, top)|Random whole number.");
  def("SUMSQ", 1, 255, function (a) { return sum(nums(a).map(function (v) { return v * v; })); }, null, "SUMSQ(number1, …)|Sum of squares.");

  // ---------- logical ----------
  function evd(node, cx) { return deref(ev(node, cx)); }
  def("IF", 1, 3, function (a, cx) {
    var c = evd(a[0], cx);
    if (c instanceof Mat) {
      var t = a[1] ? evd(a[1], cx) : true, f = a[2] ? evd(a[2], cx) : false;
      return K.lift(function (cc, tt, ff) { return bool(cc) ? (tt === null ? 0 : tt) : (ff === null ? 0 : ff); }, [c, t, f]);
    }
    if (bool(c)) return a.length > 1 ? (a[1].t === "empty" ? 0 : ev(a[1], cx)) : true;
    return a.length > 2 ? (a[2].t === "empty" ? 0 : ev(a[2], cx)) : false;
  }, LZ, "IF(test, value_if_true, [value_if_false])|Chooses a value by a condition.");
  def("IFS", 2, 254, function (a, cx) { for (var i = 0; i + 1 < a.length; i += 2) if (bool(S(evd(a[i], cx)))) return ev(a[i + 1], cx); throw E["#N/A"]; }, LZ, "IFS(test1, value1, [test2, value2], …)|First value whose test is TRUE.");
  function logic(a, init, step) {
    var acc = init, seen = false;
    a.forEach(function (x) { if (x === undefined) return; each(x, function (v, direct) {
      if (v instanceof Err) throw v;
      if (typeof v === "boolean" || typeof v === "number") { acc = step(acc, bool(v)); seen = true; }
      else if (direct && typeof v === "string") { acc = step(acc, bool(v)); seen = true; }
    }); });
    if (!seen) throw E["#VALUE!"];
    return acc;
  }
  def("AND", 1, 255, function (a) { return logic(a, true, function (p, v) { return p && v; }); }, null, "AND(logical1, …)|TRUE if all are TRUE.");
  def("OR", 1, 255, function (a) { return logic(a, false, function (p, v) { return p || v; }); }, null, "OR(logical1, …)|TRUE if any is TRUE.");
  def("XOR", 1, 255, function (a) { return logic(a, false, function (p, v) { return p !== v; }); }, null, "XOR(logical1, …)|TRUE if an odd number are TRUE.");
  def("NOT", 1, 1, function (a) { return !b(a[0]); }, L, "NOT(logical)|Reverses TRUE/FALSE.");
  def("TRUE", 0, 0, function () { return true; }, null, "TRUE()|The value TRUE.");
  def("FALSE", 0, 0, function () { return false; }, null, "FALSE()|The value FALSE.");
  function tryEval(node, cx) { try { return evd(node, cx); } catch (e) { if (e instanceof Err) return e; throw e; } }
  def("IFERROR", 2, 2, function (a, cx) {
    var v = tryEval(a[0], cx);
    if (v instanceof Mat) { var alt = null; return K.lift(function (x) { if (x instanceof Err) { if (alt === null) alt = tryEval(a[1], cx); if (alt instanceof Err) throw alt; return alt; } return x; }, [v]); }
    if (v instanceof Err) return a[1].t === "empty" ? 0 : ev(a[1], cx);
    return v === null ? 0 : v;
  }, LZ, "IFERROR(value, value_if_error)|Replaces any error.");
  def("IFNA", 2, 2, function (a, cx) { var v = tryEval(a[0], cx); if (v instanceof Err) { if (v.e === "#N/A") return ev(a[1], cx); throw v; } return v === null ? 0 : v; }, LZ, "IFNA(value, value_if_na)|Replaces #N/A.");
  def("SWITCH", 3, 254, function (a, cx) {
    var x = S(evd(a[0], cx)), i = 1;
    for (; i + 1 < a.length; i += 2) if (cmp(x, S(evd(a[i], cx))) === 0) return ev(a[i + 1], cx);
    if (i < a.length) return ev(a[i], cx);
    throw E["#N/A"];
  }, LZ, "SWITCH(expression, value1, result1, …, [default])|Matches a value against a list.");
  def("CHOOSE", 2, 255, function (a, cx) { var i = Math.trunc(num(S(evd(a[0], cx)))); if (i < 1 || i >= a.length) throw E["#VALUE!"]; return ev(a[i], cx); }, LZ, "CHOOSE(index, value1, value2, …)|Picks a value by position.");

  // ---------- information ----------
  function info(test) { return function (a, cx) { var v = tryEval(a[0], cx); if (v instanceof Mat) return K.lift(function (x) { return test(x); }, [v]); return test(v); }; }
  def("ISBLANK", 1, 1, info(function (v) { return v === null; }), LZ, "ISBLANK(value)|TRUE if the cell is empty.");
  def("ISNUMBER", 1, 1, info(function (v) { return typeof v === "number"; }), LZ, "ISNUMBER(value)|TRUE for numbers.");
  def("ISTEXT", 1, 1, info(function (v) { return typeof v === "string"; }), LZ, "ISTEXT(value)|TRUE for text.");
  def("ISNONTEXT", 1, 1, info(function (v) { return typeof v !== "string"; }), LZ, "ISNONTEXT(value)|TRUE unless text.");
  def("ISLOGICAL", 1, 1, info(function (v) { return typeof v === "boolean"; }), LZ, "ISLOGICAL(value)|TRUE for TRUE/FALSE.");
  def("ISERROR", 1, 1, info(function (v) { return v instanceof Err; }), LZ, "ISERROR(value)|TRUE for any error.");
  def("ISERR", 1, 1, info(function (v) { return v instanceof Err && v.e !== "#N/A"; }), LZ, "ISERR(value)|TRUE for errors other than #N/A.");
  def("ISNA", 1, 1, info(function (v) { return v instanceof Err && v.e === "#N/A"; }), LZ, "ISNA(value)|TRUE for #N/A.");
  def("ISEVEN", 1, 1, function (a) { return Math.floor(Math.abs(n(a[0]))) % 2 === 0; }, L, "ISEVEN(number)|TRUE if even.");
  def("ISODD", 1, 1, function (a) { return Math.floor(Math.abs(n(a[0]))) % 2 === 1; }, L, "ISODD(number)|TRUE if odd.");
  def("NA", 0, 0, function () { throw E["#N/A"]; }, null, "NA()|The #N/A error.");
  def("N", 1, 1, function (a) { var v = S(a[0]); return typeof v === "number" ? v : v === true ? 1 : 0; }, null, "N(value)|Value as a number (text → 0).");

  // ---------- lookup & reference ----------
  function lookEq(v, x) {
    if (typeof x === "string" && typeof v === "string") return hasWild(x) ? wild(x).test(v) : v.toLowerCase() === x.toLowerCase();
    if (typeof x === "number" && typeof v === "number") return C.r15(v) === C.r15(x);
    return v === x;
  }
  function sameKind(v, x) { return (typeof v === "number" && typeof x === "number") || (typeof v === "string" && typeof x === "string") || (typeof v === "boolean" && typeof x === "boolean"); }
  // position in a 1-D list; mode 0 exact, 1 largest <= x (sorted asc), -1 smallest >= x (sorted desc)
  function findIn(len, get, x, mode) {
    var best = -1, i;
    if (mode === 0) { for (i = 0; i < len; i++) if (lookEq(get(i), x)) return i; return -1; }
    for (i = 0; i < len; i++) {
      var v = get(i);
      if (v === null || !sameKind(v, x)) continue;
      var c = cmp(v, x);
      if (mode > 0) { if (c <= 0) best = i; else break; }
      else { if (c >= 0) best = i; else break; }
    }
    return best;
  }
  def("VLOOKUP", 3, 4, function (a) {
    var x = S(a[0]), t = a[1], col = int(a[2]), approx = a[3] === undefined ? true : b(a[3]);
    if (x instanceof Err) throw x;
    var D = dims(t), g = getter(t);
    if (col < 1) throw E["#VALUE!"];
    if (col > (t instanceof Rng ? t.W() : D[1])) throw E["#REF!"];
    var i = findIn(D[0], function (k) { return g(k, 0); }, x, approx ? 1 : 0);
    if (i < 0) throw E["#N/A"];
    var v = g(i, col - 1); return v === null ? 0 : v;
  }, null, "VLOOKUP(value, table, column, [approximate])|Looks down the first column; FALSE = exact match.");
  def("HLOOKUP", 3, 4, function (a) {
    var x = S(a[0]), t = a[1], row = int(a[2]), approx = a[3] === undefined ? true : b(a[3]);
    if (x instanceof Err) throw x;
    var D = dims(t), g = getter(t);
    if (row < 1) throw E["#VALUE!"];
    if (row > (t instanceof Rng ? t.H() : D[0])) throw E["#REF!"];
    var j = findIn(D[1], function (k) { return g(0, k); }, x, approx ? 1 : 0);
    if (j < 0) throw E["#N/A"];
    var v = g(row - 1, j); return v === null ? 0 : v;
  }, null, "HLOOKUP(value, table, row, [approximate])|Looks across the first row.");
  function vec(x) {
    var D = dims(x), g = getter(x);
    if (D[0] > 1 && D[1] > 1) throw E["#N/A"];
    return D[0] >= D[1] || D[1] === 1 ? { len: D[0], get: function (i) { return g(i, 0); }, col: true } : { len: D[1], get: function (j) { return g(0, j); }, col: false };
  }
  def("MATCH", 2, 3, function (a) {
    var x = S(a[0]), v = vec(a[1]), mode = a[2] === undefined ? 1 : Math.sign(n(a[2]));
    if (x instanceof Err) throw x;
    var i = findIn(v.len, v.get, x, mode);
    if (i < 0) throw E["#N/A"]; return i + 1;
  }, null, "MATCH(value, range, [type])|Position of a value; 0 = exact, 1 = largest ≤, -1 = smallest ≥.");
  function xfind(x, v, mode, search) {
    var order = [], i;
    for (i = 0; i < v.len; i++) order.push(i);
    if (search < 0) order.reverse();
    if (mode === 0 || mode === 2) { for (i = 0; i < order.length; i++) { var y = v.get(order[i]); if (mode === 2 ? lookEq(y, x) : (typeof x === "string" && typeof y === "string" ? y.toLowerCase() === x.toLowerCase() : lookEq(y, x))) return order[i]; } return -1; }
    var best = -1, bv = null;
    for (i = 0; i < order.length; i++) {
      var w = v.get(order[i]);
      if (w === null || !sameKind(w, x)) continue;
      var c = cmp(w, x);
      if (c === 0) return order[i];
      if (mode < 0 && c < 0 && (bv === null || cmp(w, bv) > 0)) { best = order[i]; bv = w; }
      if (mode > 0 && c > 0 && (bv === null || cmp(w, bv) < 0)) { best = order[i]; bv = w; }
    }
    return best;
  }
  def("XLOOKUP", 3, 6, function (a) {
    var x = S(a[0]), v = vec(a[1]), ret = a[2], mode = int(opt(a[4], 0)), search = int(opt(a[5], 1));
    if (x instanceof Err) throw x;
    var i = xfind(x, v, mode, search);
    if (i < 0) { if (a[3] !== undefined && a[3] !== null) return deref(a[3]); throw E["#N/A"]; }
    var D = dims(ret), g = getter(ret), row = [], k;
    if (v.col) { if (D[1] === 1) return nz(g(i, 0)); for (k = 0; k < D[1]; k++) row.push(nz(g(i, k))); return new Mat([row]); }
    if (D[0] === 1) return nz(g(0, i));
    for (k = 0; k < D[0]; k++) row.push([nz(g(k, i))]); return new Mat(row);
  }, null, "XLOOKUP(value, lookup_range, return_range, [if_not_found], [match_mode], [search_mode])|Modern lookup in any direction.");
  function nz(v) { return v === null ? 0 : v; }
  def("XMATCH", 2, 4, function (a) {
    var x = S(a[0]), i = xfind(x, vec(a[1]), int(opt(a[2], 0)), int(opt(a[3], 1)));
    if (i < 0) throw E["#N/A"]; return i + 1;
  }, null, "XMATCH(value, range, [match_mode], [search_mode])|Position of a value.");
  def("LOOKUP", 2, 3, function (a) {
    var x = S(a[0]), v = vec(a[1]), i = findIn(v.len, v.get, x, 1);
    if (i < 0) throw E["#N/A"];
    var r = a[2] === undefined ? v : vec(a[2]);
    return nz(r.get(i));
  }, null, "LOOKUP(value, lookup_vector, [result_vector])|Approximate lookup in a sorted list.");
  def("INDEX", 2, 3, function (a) {
    var t = a[0], D = t instanceof Rng ? [t.H(), t.W()] : dims(t), g = getter(t), r = a[1] === undefined || a[1] === null ? 0 : int(a[1]), c = a[2] === undefined || a[2] === null ? 0 : int(a[2]);
    if (a[2] === undefined && D[0] === 1 && D[1] > 1) { c = r; r = 1; }
    if (a[2] === undefined && D[1] === 1) c = 1;
    if (r < 0 || c < 0 || r > D[0] || c > D[1]) throw E["#REF!"];
    if (r === 0 && c === 0) return t;
    if (r === 0) { var col = []; for (var i = 0; i < Math.min(D[0], dims(t)[0]); i++) col.push([nz(g(i, c - 1))]); return new Mat(col); }
    if (c === 0) { var row = []; for (var j = 0; j < Math.min(D[1], dims(t)[1]); j++) row.push(nz(g(r - 1, j))); return new Mat([row]); }
    return nz(g(r - 1, c - 1));
  }, null, "INDEX(range, row, [column])|Value at a position in a range.");
  def("ROW", 0, 1, function (a, cx) { var x = a[0]; if (x === undefined) return cx.r + 1; if (!(x instanceof Rng)) throw E["#VALUE!"]; return x.r0 + 1; }, null, "ROW([ref])|Row number.");
  def("COLUMN", 0, 1, function (a, cx) { var x = a[0]; if (x === undefined) return cx.c + 1; if (!(x instanceof Rng)) throw E["#VALUE!"]; return x.c0 + 1; }, null, "COLUMN([ref])|Column number.");
  def("ROWS", 1, 1, function (a) { var x = a[0]; return x instanceof Rng ? x.H() : dims(x)[0]; }, null, "ROWS(range)|Number of rows.");
  def("COLUMNS", 1, 1, function (a) { var x = a[0]; return x instanceof Rng ? x.W() : dims(x)[1]; }, null, "COLUMNS(range)|Number of columns.");
  def("OFFSET", 3, 5, function (a) {
    var x = a[0]; if (!(x instanceof Rng)) throw E["#VALUE!"];
    var r0 = x.r0 + int(a[1]), c0 = x.c0 + int(a[2]), h = a[3] == null ? x.H() : int(a[3]), w = a[4] == null ? x.W() : int(a[4]);
    if (h < 1 || w < 1 || r0 < 0 || c0 < 0 || r0 + h > C.MAXR || c0 + w > C.MAXC) throw E["#REF!"];
    return new Rng(x.cx, x.sh, r0, c0, r0 + h - 1, c0 + w - 1);
  }, VOL, "OFFSET(ref, rows, cols, [height], [width])|A range shifted from a starting cell.");
  def("INDIRECT", 1, 2, function (a, cx) {
    var g = C.parseRange(s(a[0])); if (!g) throw E["#REF!"];
    var sh = g.sheet ? cx.eng.wb.sheetByName(g.sheet) : cx.sh; if (!sh) throw E["#REF!"];
    return new Rng(cx, sh, g.r0, g.c0, g.r1, g.c1);
  }, VOL, "INDIRECT(text)|The range named by a text like \"B2\" or \"Sheet2!A1:A9\".");
  def("TRANSPOSE", 1, 1, function (a) { var m = toMat(deref(a[0]) instanceof Mat ? deref(a[0]) : a[0]), rows = []; for (var j = 0; j < m.w(); j++) { rows.push([]); for (var i = 0; i < m.h(); i++) rows[j].push(nz(m.rows[i][j])); } return new Mat(rows); }, null, "TRANSPOSE(range)|Swaps rows and columns.");
  function matOf(x) { var d = x instanceof Rng ? toMat(x) : x instanceof Mat ? x : new Mat([[x]]); return d; }
  function keyOf(v) { return v === null ? "" : typeof v === "string" ? "s" + v.toLowerCase() : typeof v + String(v instanceof Err ? v.e : v); }
  def("UNIQUE", 1, 3, function (a) {
    var m = matOf(a[0]), byCol = a[1] != null && b(a[1]), once = a[2] != null && b(a[2]);
    var items = byCol ? m.rows[0].map(function (_, j) { return m.rows.map(function (r) { return r[j]; }); }) : m.rows;
    var cnt = new Map(), order = [];
    items.forEach(function (it) { var k = it.map(keyOf).join("\u0001"); if (!cnt.has(k)) { cnt.set(k, 0); order.push([k, it]); } cnt.set(k, cnt.get(k) + 1); });
    var out = order.filter(function (p) { return !once || cnt.get(p[0]) === 1; }).map(function (p) { return p[1].map(nz); });
    if (!out.length) throw E["#CALC!"];
    if (byCol) { var t = []; for (var i = 0; i < out[0].length; i++) t.push(out.map(function (c) { return c[i]; })); return new Mat(t); }
    return new Mat(out);
  }, null, "UNIQUE(range, [by_column], [exactly_once])|Distinct rows of a range (spills).");
  def("SORT", 1, 4, function (a) {
    var m = matOf(a[0]), idx = int(opt(a[1], 1)) - 1, ord = n(opt(a[2], 1)) < 0 ? -1 : 1, byCol = a[3] != null && b(a[3]);
    var rows = byCol ? m.rows[0].map(function (_, j) { return m.rows.map(function (r) { return r[j]; }); }) : m.rows.slice();
    if (idx < 0 || idx >= rows[0].length) throw E["#VALUE!"];
    rows = rows.map(function (r, i) { return { r: r, i: i }; });
    rows.sort(function (x, y) { var p = x.r[idx], q = y.r[idx], c; if (p === null && q === null) c = 0; else if (p === null) return 1; else if (q === null) return -1; else c = safeCmp(p, q) * ord; return c || x.i - y.i; });
    rows = rows.map(function (o) { return o.r.map(nz); });
    if (byCol) { var t = []; for (var i = 0; i < rows[0].length; i++) t.push(rows.map(function (c) { return c[i]; })); return new Mat(t); }
    return new Mat(rows);
  }, null, "SORT(range, [sort_index], [order], [by_column])|Sorted copy of a range (spills).");
  function safeCmp(p, q) { if (p instanceof Err || q instanceof Err) return p instanceof Err ? (q instanceof Err ? 0 : 1) : -1; return cmp(p, q); }
  def("FILTER", 2, 3, function (a) {
    var m = matOf(a[0]), inc = matOf(deref(a[1]) instanceof Mat ? deref(a[1]) : a[1]), out = [];
    if (inc.w() === 1 && inc.h() === m.h()) { m.rows.forEach(function (r, i) { var v = inc.rows[i][0]; if (v instanceof Err) throw v; if (bool(v)) out.push(r.map(nz)); }); }
    else if (inc.h() === 1 && inc.w() === m.w()) {
      var keep = inc.rows[0].map(function (v) { if (v instanceof Err) throw v; return bool(v); });
      if (keep.some(Boolean)) out = m.rows.map(function (r) { return r.filter(function (_, j) { return keep[j]; }).map(nz); });
    } else throw E["#VALUE!"];
    if (!out.length) { if (a[2] !== undefined) return deref(a[2]); throw E["#CALC!"]; }
    return new Mat(out);
  }, null, "FILTER(range, include, [if_empty])|Rows where include is TRUE, e.g. FILTER(A2:C9, C2:C9>100).");
  def("SEQUENCE", 1, 4, function (a) {
    var h = int(a[0]), w = int(opt(a[1], 1)), st = n(opt(a[2], 1)), step = n(opt(a[3], 1)), rows = [];
    if (h < 1 || w < 1 || h * w > 1e6) throw E["#VALUE!"];
    for (var i = 0; i < h; i++) { var r = []; for (var j = 0; j < w; j++) r.push(C.r15(st + (i * w + j) * step)); rows.push(r); }
    return new Mat(rows);
  }, null, "SEQUENCE(rows, [columns], [start], [step])|A list of numbers (spills).");

  // ---------- text ----------
  function txtArgs(a) { var out = []; a.forEach(function (x) { if (x === undefined) return; each(x, function (v) { if (v instanceof Err) throw v; out.push(str(v)); }); }); return out; }
  def("CONCAT", 1, 255, function (a) { return txtArgs(a).join(""); }, null, "CONCAT(text1, …)|Joins texts and ranges.");
  def("CONCATENATE", 1, 255, function (a) { return a.map(function (x) { return s(x); }).join(""); }, null, "CONCATENATE(text1, …)|Joins texts.");
  def("TEXTJOIN", 3, 255, function (a) {
    var d = s(a[0]), skip = b(a[1]), out = [];
    a.slice(2).forEach(function (x) { each(x, function (v) { if (v instanceof Err) throw v; var t = str(v); if (!(skip && t === "")) out.push(t); }); });
    return out.join(d);
  }, null, "TEXTJOIN(delimiter, ignore_empty, text1, …)|Joins texts with a separator.");
  function ln(a, dflt) { var x = a === undefined ? dflt : Math.trunc(n(a)); if (x < 0) throw E["#VALUE!"]; return x; }
  def("LEFT", 1, 2, function (a) { var t = s(a[0]); return Array.from(t).slice(0, ln(a[1], 1)).join(""); }, L, "LEFT(text, [count])|First characters.");
  def("RIGHT", 1, 2, function (a) { var t = Array.from(s(a[0])), k = ln(a[1], 1); return k ? t.slice(-k).join("") : ""; }, L, "RIGHT(text, [count])|Last characters.");
  def("MID", 3, 3, function (a) { var t = Array.from(s(a[0])), st = Math.trunc(n(a[1])), k = ln(a[2], 0); if (st < 1) throw E["#VALUE!"]; return t.slice(st - 1, st - 1 + k).join(""); }, L, "MID(text, start, count)|Characters from the middle.");
  def("LEN", 1, 1, function (a) { return Array.from(s(a[0])).length; }, L, "LEN(text)|Number of characters.");
  def("TRIM", 1, 1, function (a) { return s(a[0]).replace(/ +/g, " ").replace(/^ | $/g, ""); }, L, "TRIM(text)|Removes extra spaces.");
  def("UPPER", 1, 1, function (a) { return s(a[0]).toUpperCase(); }, L, "UPPER(text)|UPPER CASE.");
  def("LOWER", 1, 1, function (a) { return s(a[0]).toLowerCase(); }, L, "LOWER(text)|lower case.");
  def("PROPER", 1, 1, function (a) { return s(a[0]).toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, function (m0, p, c) { return p + c.toUpperCase(); }); }, L, "PROPER(text)|Capitalises Each Word.");
  def("SUBSTITUTE", 3, 4, function (a) {
    var t = s(a[0]), o = s(a[1]), nw = s(a[2]);
    if (!o) return t;
    if (a[3] === undefined) return t.split(o).join(nw);
    var k = Math.trunc(n(a[3])), pos = -1; if (k < 1) throw E["#VALUE!"];
    for (var i = 0; i < k; i++) { pos = t.indexOf(o, pos + 1); if (pos < 0) return t; }
    return t.slice(0, pos) + nw + t.slice(pos + o.length);
  }, L, "SUBSTITUTE(text, old, new, [instance])|Replaces text by matching.");
  def("REPLACE", 4, 4, function (a) { var t = Array.from(s(a[0])), st = Math.trunc(n(a[1])), k = ln(a[2], 0); if (st < 1) throw E["#VALUE!"]; return t.slice(0, st - 1).join("") + s(a[3]) + t.slice(st - 1 + k).join(""); }, L, "REPLACE(text, start, count, new)|Replaces text by position.");
  def("FIND", 2, 3, function (a) { var f = s(a[0]), t = s(a[1]), st = a[2] === undefined ? 1 : Math.trunc(n(a[2])); if (st < 1 || st > t.length + 1) throw E["#VALUE!"]; var i = t.indexOf(f, st - 1); if (i < 0) throw E["#VALUE!"]; return i + 1; }, L, "FIND(find, within, [start])|Position of text (case-sensitive).");
  def("SEARCH", 2, 3, function (a) {
    var f = s(a[0]), t = s(a[1]), st = a[2] === undefined ? 1 : Math.trunc(n(a[2]));
    if (st < 1 || st > t.length + 1) throw E["#VALUE!"];
    if (!f) return st;
    var rx = new RegExp(wild(f).source.slice(1, -1), "i"), m = rx.exec(t.slice(st - 1));
    if (!m) throw E["#VALUE!"]; return m.index + st;
  }, L, "SEARCH(find, within, [start])|Position of text (any case, * and ? wildcards).");
  def("TEXT", 2, 2, function (a) { var v = S(a[0]), f = s(a[1]); if (v instanceof Err) throw v; if (typeof v === "string") { var p = numOrNull(v); if (p === null) return F.format(v, f).t; v = p; } if (typeof v === "boolean") return v ? "TRUE" : "FALSE"; return F.format(v === null ? 0 : v, f).t; }, L, "TEXT(value, format)|Formats a number as text, e.g. TEXT(A1, \"0.00%\").");
  def("VALUE", 1, 1, function (a) { var v = S(a[0]); if (typeof v === "number") return v; if (v === null) return 0; var p = numOrNull(str(v)); if (p === null) throw E["#VALUE!"]; return p; }, L, "VALUE(text)|Text to number.");
  def("REPT", 2, 2, function (a) { var k = Math.trunc(n(a[1])); if (k < 0 || k * s(a[0]).length > 32767) throw E["#VALUE!"]; return s(a[0]).repeat(k); }, L, "REPT(text, times)|Repeats text.");
  def("EXACT", 2, 2, function (a) { return s(a[0]) === s(a[1]); }, L, "EXACT(text1, text2)|TRUE if identical (case-sensitive).");
  def("CHAR", 1, 1, function (a) { var k = Math.trunc(n(a[0])); if (k < 1 || k > 255) throw E["#VALUE!"]; return String.fromCharCode(k); }, L, "CHAR(number)|Character from a code.");
  def("CODE", 1, 1, function (a) { var t = s(a[0]); if (!t) throw E["#VALUE!"]; return t.charCodeAt(0); }, L, "CODE(text)|Code of the first character.");
  def("UNICHAR", 1, 1, function (a) { var k = Math.trunc(n(a[0])); if (k < 1 || k > 0x10ffff) throw E["#VALUE!"]; return String.fromCodePoint(k); }, L, "UNICHAR(number)|Unicode character.");
  def("UNICODE", 1, 1, function (a) { var t = s(a[0]); if (!t) throw E["#VALUE!"]; return t.codePointAt(0); }, L, "UNICODE(text)|Unicode code point.");
  def("CLEAN", 1, 1, function (a) { return s(a[0]).replace(/[\x00-\x1f]/g, ""); }, L, "CLEAN(text)|Removes non-printing characters.");
  function beforeAfter(after) {
    return function (a) {
      var t = s(a[0]), d = s(a[1]), k = a[2] === undefined ? 1 : Math.trunc(n(a[2]));
      if (k === 0) throw E["#VALUE!"];
      var pos = -1, i;
      if (k > 0) { for (i = 0; i < k; i++) { pos = t.indexOf(d, pos < 0 ? 0 : pos + Math.max(d.length, 1)); if (pos < 0) throw E["#N/A"]; } }
      else { pos = t.length + 1; for (i = 0; i < -k; i++) { pos = t.lastIndexOf(d, pos - 1 - (i ? 0 : 0)); if (pos < 0) throw E["#N/A"]; } }
      return after ? t.slice(pos + d.length) : t.slice(0, pos);
    };
  }
  def("TEXTBEFORE", 2, 3, beforeAfter(false), L, "TEXTBEFORE(text, delimiter, [instance])|Text before a delimiter.");
  def("TEXTAFTER", 2, 3, beforeAfter(true), L, "TEXTAFTER(text, delimiter, [instance])|Text after a delimiter.");
  def("FIXED", 1, 3, function (a) { var d = a[1] === undefined ? 2 : Math.trunc(n(a[1])), nc = a[2] !== undefined && b(a[2]); return F.format(n(a[0]), (nc ? "0" : "#,##0") + (d > 0 ? "." + "0".repeat(d) : "")).t; }, L, "FIXED(number, [decimals], [no_commas])|Number as text with fixed decimals.");

  // ---------- date & time (serials: 1 = 1900-01-01, 60 = the 1900-02-29 that never was) ----------
  function dn(a) { var v = n(a); if (v < 0) throw E["#NUM!"]; return v; }
  def("TODAY", 0, 0, function () { return F.today(); }, VOL, "TODAY()|Today's date (serial).");
  def("NOW", 0, 0, function () { return F.now(); }, VOL, "NOW()|Date and time now.");
  def("DATE", 3, 3, function (a) {
    var y = Math.trunc(n(a[0])), m = Math.trunc(n(a[1])), d = Math.trunc(n(a[2]));
    if (y >= 0 && y < 1900) y += 1900;
    if (y < 0 || y > 9999) throw E["#NUM!"];
    var v = F.serial(y, m, d);
    if (v < 0 || v > 2958465) throw E["#NUM!"]; return v;
  }, L, "DATE(year, month, day)|Date serial; months and days may overflow.");
  def("TIME", 3, 3, function (a) { var t = (Math.trunc(n(a[0])) * 3600 + Math.trunc(n(a[1])) * 60 + Math.trunc(n(a[2]))) / 86400; if (t < 0) throw E["#NUM!"]; return t - Math.floor(t); }, L, "TIME(hour, minute, second)|Time as a fraction of a day.");
  def("YEAR", 1, 1, function (a) { return F.ymd(dn(a[0])).y; }, L, "YEAR(date)|Year of a date.");
  def("MONTH", 1, 1, function (a) { return F.ymd(dn(a[0])).m; }, L, "MONTH(date)|Month 1–12.");
  def("DAY", 1, 1, function (a) { return F.ymd(dn(a[0])).d; }, L, "DAY(date)|Day of the month.");
  def("HOUR", 1, 1, function (a) { return F.hms(dn(a[0])).h; }, L, "HOUR(time)|Hour 0–23.");
  def("MINUTE", 1, 1, function (a) { return F.hms(dn(a[0])).mi; }, L, "MINUTE(time)|Minute 0–59.");
  def("SECOND", 1, 1, function (a) { return F.hms(dn(a[0])).s; }, L, "SECOND(time)|Second 0–59.");
  def("WEEKDAY", 1, 2, function (a) {
    var w = F.weekday(dn(a[0])), t = a[1] === undefined ? 1 : Math.trunc(n(a[1]));
    if (t === 1 || t === 17) return w + 1;
    if (t === 2 || t === 11) return (w + 6) % 7 + 1;
    if (t === 3) return (w + 6) % 7;
    if (t >= 12 && t <= 16) return (w - (t - 10) + 7) % 7 + 1;
    throw E["#NUM!"];
  }, L, "WEEKDAY(date, [type])|Day of the week (1 = Sunday by default; type 2: 1 = Monday).");
  def("WEEKNUM", 1, 2, function (a) {
    var v = Math.floor(dn(a[0])), t = a[1] === undefined ? 1 : Math.trunc(n(a[1])), y = F.ymd(v).y, j1 = F.serial(y, 1, 1);
    var startDay = t === 2 || t === 11 ? 1 : t >= 12 && t <= 17 ? (t - 10) % 7 : 0;
    var off = (F.weekday(j1) - startDay + 7) % 7;
    return Math.floor((v - j1 + off) / 7) + 1;
  }, L, "WEEKNUM(date, [type])|Week of the year.");
  function addMonths(v, m) { var d = F.ymd(v), mm = d.m - 1 + m, y = d.y + Math.floor(mm / 12); mm = ((mm % 12) + 12) % 12 + 1; return { y: y, m: mm, d: d.d }; }
  def("EDATE", 2, 2, function (a) { var t = addMonths(Math.floor(dn(a[0])), Math.trunc(n(a[1]))); var v = F.serial(t.y, t.m, Math.min(t.d, F.daysIn(t.y, t.m))); if (v < 0) throw E["#NUM!"]; return v; }, L, "EDATE(start, months)|Same day some months later.");
  def("EOMONTH", 2, 2, function (a) { var t = addMonths(Math.floor(dn(a[0])), Math.trunc(n(a[1]))); var v = F.serial(t.y, t.m, F.daysIn(t.y, t.m)); if (v < 0) throw E["#NUM!"]; return v; }, L, "EOMONTH(start, months)|Last day of the month.");
  def("DAYS", 2, 2, function (a) { return Math.floor(n(a[0])) - Math.floor(n(a[1])); }, L, "DAYS(end, start)|Days between two dates.");
  def("DATEVALUE", 1, 1, function (a) { var p = F.parseDate(s(a[0]).trim()); if (!p) throw E["#VALUE!"]; return Math.floor(p.v); }, L, "DATEVALUE(text)|Date text to a serial.");
  function holidays(x) { var set = new Set(); if (x !== undefined && x !== null) each(x, function (v) { if (typeof v === "number") set.add(Math.floor(v)); }); return set; }
  def("NETWORKDAYS", 2, 3, function (a) {
    var s0 = Math.floor(dn(a[0])), e0 = Math.floor(dn(a[1])), hol = holidays(a[2]), sign = 1;
    if (s0 > e0) { var t = s0; s0 = e0; e0 = t; sign = -1; }
    if (e0 - s0 > 400000) throw E["#NUM!"];
    var c = 0;
    for (var d = s0; d <= e0; d++) { var w = F.weekday(d); if (w !== 0 && w !== 6 && !hol.has(d)) c++; }
    return c * sign;
  }, null, "NETWORKDAYS(start, end, [holidays])|Working days (Mon–Fri) between two dates, inclusive.");
  def("WORKDAY", 2, 3, function (a) {
    var d = Math.floor(dn(a[0])), k = Math.trunc(n(a[1])), hol = holidays(a[2]), step = k < 0 ? -1 : 1, guard = 0;
    while (k !== 0) { d += step; var w = F.weekday(d); if (w !== 0 && w !== 6 && !hol.has(d)) k -= step; if (++guard > 1e6 || d < 0) throw E["#NUM!"]; }
    return d;
  }, null, "WORKDAY(start, days, [holidays])|Date after some working days.");
  def("DATEDIF", 3, 3, function (a) {
    var s0 = Math.floor(dn(a[0])), e0 = Math.floor(dn(a[1])), u = s(a[2]).toUpperCase();
    if (s0 > e0) throw E["#NUM!"];
    var A = F.ymd(s0), B = F.ymd(e0);
    var months = (B.y - A.y) * 12 + B.m - A.m - (B.d < A.d ? 1 : 0);
    switch (u) {
      case "Y": return Math.floor(months / 12);
      case "M": return months;
      case "D": return e0 - s0;
      case "YM": return months % 12;
      case "MD": { if (B.d >= A.d) return B.d - A.d; var pm = B.m === 1 ? 12 : B.m - 1, py = B.m === 1 ? B.y - 1 : B.y; return F.daysIn(py, pm) - A.d + B.d; }
      case "YD": {
        var y = B.y, st = F.serial(y, A.m, Math.min(A.d, F.daysIn(y, A.m)));
        if (st > e0) { y--; st = F.serial(y, A.m, Math.min(A.d, F.daysIn(y, A.m))); }
        return e0 - st;
      }
    }
    throw E["#NUM!"];
  }, null, "DATEDIF(start, end, unit)|Difference in \"Y\", \"M\", \"D\", \"YM\", \"MD\" or \"YD\".");
  def("YEARFRAC", 2, 3, function (a) {
    var s0 = Math.floor(dn(a[0])), e0 = Math.floor(dn(a[1])), basis = a[2] === undefined ? 0 : Math.trunc(n(a[2]));
    if (s0 > e0) { var t = s0; s0 = e0; e0 = t; }
    if (basis === 3) return (e0 - s0) / 365;
    if (basis === 2) return (e0 - s0) / 360;
    var A = F.ymd(s0), B = F.ymd(e0), d1 = A.d, d2 = B.d;
    if (d1 === 31) d1 = 30; if (d2 === 31 && d1 === 30) d2 = 30;
    return ((B.y - A.y) * 360 + (B.m - A.m) * 30 + d2 - d1) / 360;
  }, null, "YEARFRAC(start, end, [basis])|Fraction of a year between dates.");

  // ---------- finance ----------
  function pmt(r, nper, pv, fv, type) {
    if (nper === 0) throw E["#NUM!"];
    if (r === 0) return -(pv + fv) / nper;
    var p = Math.pow(1 + r, nper);
    return -(r * (fv + pv * p)) / ((1 + r * type) * (p - 1));
  }
  function fvf(r, nper, p, pv, type) { if (r === 0) return -(pv + p * nper); var q = Math.pow(1 + r, nper); return -(pv * q + p * (1 + r * type) * (q - 1) / r); }
  def("PMT", 3, 5, function (a) { return chk(pmt(n(a[0]), n(a[1]), n(a[2]), on(a[3], 0), on(a[4], 0) ? 1 : 0)); }, L, "PMT(rate, periods, present_value, [future_value], [type])|Payment per period of a loan.");
  def("FV", 3, 5, function (a) { return chk(fvf(n(a[0]), n(a[1]), n(a[2]), on(a[3], 0), on(a[4], 0) ? 1 : 0)); }, L, "FV(rate, periods, payment, [present_value], [type])|Future value of an investment.");
  def("PV", 3, 5, function (a) {
    var r = n(a[0]), nper = n(a[1]), p = n(a[2]), fv = on(a[3], 0), type = on(a[4], 0) ? 1 : 0;
    if (r === 0) return -(fv + p * nper);
    var q = Math.pow(1 + r, nper); return chk(-(fv + p * (1 + r * type) * (q - 1) / r) / q);
  }, L, "PV(rate, periods, payment, [future_value], [type])|Present value.");
  def("NPER", 3, 5, function (a) {
    var r = n(a[0]), p = n(a[1]), pv = n(a[2]), fv = on(a[3], 0), type = on(a[4], 0) ? 1 : 0;
    if (r === 0) { if (p === 0) throw E["#NUM!"]; return -(pv + fv) / p; }
    var x = (p * (1 + r * type) - fv * r) / (p * (1 + r * type) + pv * r);
    if (x <= 0) throw E["#NUM!"]; return chk(Math.log(x) / Math.log(1 + r));
  }, L, "NPER(rate, payment, present_value, [future_value], [type])|Number of periods.");
  def("RATE", 3, 6, function (a) {
    var nper = n(a[0]), p = n(a[1]), pv = n(a[2]), fv = on(a[3], 0), type = on(a[4], 0) ? 1 : 0, r = on(a[5], 0.1);
    for (var i = 0; i < 100; i++) {
      var f = pv * Math.pow(1 + r, nper) + p * (1 + r * type) * (Math.pow(1 + r, nper) - 1) / (r || 1e-12) + fv;
      var h = 1e-7, r2 = r + h, f2 = pv * Math.pow(1 + r2, nper) + p * (1 + r2 * type) * (Math.pow(1 + r2, nper) - 1) / r2 + fv;
      var d = (f2 - f) / h; if (!d) break;
      var nr = r - f / d; if (!isFinite(nr)) break;
      if (Math.abs(nr - r) < 1e-10) return nr;
      r = nr;
    }
    throw E["#NUM!"];
  }, L, "RATE(periods, payment, present_value, [future_value], [type], [guess])|Interest rate per period.");
  def("NPV", 2, 255, function (a) { var r = n(a[0]), x = nums(a.slice(1)), t = 0; x.forEach(function (v, i) { t += v / Math.pow(1 + r, i + 1); }); return chk(t); }, null, "NPV(rate, value1, …)|Net present value of future cash flows.");
  def("IRR", 1, 2, function (a) {
    var x = nums([a[0]]), r = on(a[1], 0.1);
    for (var i = 0; i < 100; i++) {
      var f = 0, d = 0;
      x.forEach(function (v, k) { f += v / Math.pow(1 + r, k); d -= k * v / Math.pow(1 + r, k + 1); });
      if (!d) break;
      var nr = r - f / d; if (!isFinite(nr)) break;
      if (Math.abs(nr - r) < 1e-10) return nr;
      r = nr;
    }
    throw E["#NUM!"];
  }, null, "IRR(values, [guess])|Internal rate of return.");
  function ipmt(r, per, nper, pv, fv, type) {
    var p = pmt(r, nper, pv, fv, type), i;
    if (per === 1) i = type ? 0 : -pv * r;
    else i = type ? (fvf(r, per - 2, p, pv, 1) - p) * r : fvf(r, per - 1, p, pv, 0) * r;
    return i;
  }
  def("IPMT", 4, 6, function (a) { var per = n(a[1]), nper = n(a[2]); if (per < 1 || per > nper) throw E["#NUM!"]; return chk(ipmt(n(a[0]), per, nper, n(a[3]), on(a[4], 0), on(a[5], 0) ? 1 : 0)); }, L, "IPMT(rate, period, periods, present_value, [fv], [type])|Interest part of a payment.");
  def("PPMT", 4, 6, function (a) { var r = n(a[0]), per = n(a[1]), nper = n(a[2]), pv = n(a[3]), fv = on(a[4], 0), type = on(a[5], 0) ? 1 : 0; if (per < 1 || per > nper) throw E["#NUM!"]; return chk(pmt(r, nper, pv, fv, type) - ipmt(r, per, nper, pv, fv, type)); }, L, "PPMT(rate, period, periods, present_value, [fv], [type])|Principal part of a payment.");

  SH.fnNames = function () { return Object.keys(FN).sort(); };
  SH.crit = crit;
})();
