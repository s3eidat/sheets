/* Sheets formats: date serials (1900 system, with the historic 1900 leap-year quirk), number format
   codes ("#,##0.00", "0%", "yyyy-mm-dd", "[Red]…;…" sections) and parsing what people type. Pure logic. */
(function () {
  "use strict";
  var SH = window.SH, C = SH.core;
  var MS = 86400000, EPOCH = Date.UTC(1899, 11, 30);
  var MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

  // ---------- date serials ----------
  function utc(y, m0, d) { var t = new Date(Date.UTC(2000, m0, d)); t.setUTCFullYear(y); return t.getTime(); }
  // serial of y-m-d; month and day may overflow (DATE(2024,14,1) = 2025-02-01). Serial 60 is the fictional 1900-02-29.
  function serial(y, m, d) {
    y += Math.floor((m - 1) / 12); m = (((m - 1) % 12) + 12) % 12 + 1;
    var diff = Math.round((utc(y, m - 1, 1) - EPOCH) / MS);
    return (diff < 61 ? diff - 1 : diff) + d - 1;
  }
  function ymd(s) {
    s = Math.floor(s);
    if (s === 60) return { y: 1900, m: 2, d: 29 };
    if (s === 0) return { y: 1900, m: 1, d: 0 };
    var t = new Date(EPOCH + (s < 60 ? s + 1 : s) * MS);
    return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
  }
  function weekday(s) { return ((Math.floor(s) + 6) % 7 + 7) % 7; }                // 0 = Sunday
  function hms(s) {
    var secs = Math.round((s - Math.floor(s)) * 86400);
    if (secs >= 86400) secs = 86399;
    return { h: Math.floor(secs / 3600), mi: Math.floor(secs / 60) % 60, s: secs % 60 };
  }
  function fromDate(dt) { return dt.getTime() / MS + 25569; }                         // JS Date (UTC) -> serial
  function today() { var n = new Date(); return serial(n.getFullYear(), n.getMonth() + 1, n.getDate()); }
  function now() { var n = new Date(); return today() + (n.getHours() * 3600 + n.getMinutes() * 60 + n.getSeconds()) / 86400; }

  // ---------- general number text ----------
  function expText(v, digits) {
    var s = v.toExponential(digits), p = s.split("e"), m = p[0];
    if (m.indexOf(".") >= 0) m = m.replace(/\.?0+$/, "");
    var e = +p[1];
    return m + "E" + (e < 0 ? "-" : "+") + (Math.abs(e) < 10 ? "0" : "") + Math.abs(e);
  }
  function general(v) {
    if (!isFinite(v)) return "#NUM!";
    var a = Math.abs(v);
    if (a === 0) return "0";
    if (a >= 1e11 || a < 1e-5) return expText(v, 5);
    if (Number.isInteger(v)) return String(v);
    return String(+v.toPrecision(10));
  }
  // Number -> text inside formulas (concatenation, TEXTJOIN…): up to 15 significant digits.
  function numText(v) {
    if (!isFinite(v)) return "#NUM!";
    var s = String(C.r15(v));
    return s.indexOf("e") >= 0 ? s.replace("e+", "E+").replace("e-", "E-") : s;
  }

  // ---------- format codes ----------
  var COLORS = { black: "#000000", blue: "#0000ff", cyan: "#00aaaa", green: "#008000", magenta: "#cc00cc", red: "#d00000", white: "#ffffff", yellow: "#c8a000" };
  var cache = new Map();

  function splitSections(code) {
    var out = [], cur = "", q = false;
    for (var i = 0; i < code.length; i++) {
      var ch = code[i];
      if (ch === '"') q = !q;
      if (!q && ch === "\\" && i + 1 < code.length) { cur += ch + code[++i]; continue; }
      if (!q && ch === ";") { out.push(cur); cur = ""; continue; }
      cur += ch;
    }
    out.push(cur);
    return out;
  }
  function tokenize(sec) {
    var t = [], i = 0, m, s = sec;
    while (i < s.length) {
      var ch = s[i], rest = s.slice(i);
      if (ch === '"') { var j = s.indexOf('"', i + 1); if (j < 0) j = s.length; t.push({ k: "lit", v: s.slice(i + 1, j) }); i = j + 1; continue; }
      if (ch === "\\") { t.push({ k: "lit", v: s[i + 1] || "" }); i += 2; continue; }
      if (ch === "_") { t.push({ k: "lit", v: " " }); i += 2; continue; }
      if (ch === "*") { i += 2; continue; }
      if (ch === "[") {
        var e = s.indexOf("]", i); if (e < 0) e = s.length;
        var inner = s.slice(i + 1, e), lo = inner.toLowerCase();
        if (COLORS[lo]) t.push({ k: "color", v: COLORS[lo] });
        else if (/^(h+|m+|s+)$/i.test(inner)) t.push({ k: "d", v: inner.toLowerCase(), el: true });
        else if (inner[0] === "$") { var cur = inner.slice(1).split("-")[0]; if (cur) t.push({ k: "lit", v: cur }); }
        i = e + 1; continue;
      }
      if ((m = /^general/i.exec(rest))) { t.push({ k: "gen" }); i += 7; continue; }
      if ((m = /^(am\/pm|a\/p)/i.exec(rest))) { t.push({ k: "ampm", v: m[0].length > 3 ? "AMPM" : "AP" }); i += m[0].length; continue; }
      if ((m = /^(y+|m+|d+|h+|s+)/i.exec(rest))) { t.push({ k: "d", v: m[0].toLowerCase() }); i += m[0].length; continue; }
      if ((m = /^[eE][+-]/.exec(rest))) { t.push({ k: "E", v: m[0][1] }); i += 2; continue; }
      if ("0#?".indexOf(ch) >= 0) { t.push({ k: "ph", v: ch }); i++; continue; }
      if (ch === ".") { t.push({ k: "dot" }); i++; continue; }
      if (ch === ",") { t.push({ k: "comma" }); i++; continue; }
      if (ch === "%") { t.push({ k: "pct" }); i++; continue; }
      if (ch === "@") { t.push({ k: "at" }); i++; continue; }
      t.push({ k: "lit", v: ch }); i++;
    }
    return t;
  }
  function compileSection(src) {
    var toks = tokenize(src), sec = { toks: toks, color: null, kind: "text" };
    toks.forEach(function (t) { if (t.k === "color") sec.color = t.v; });
    if (toks.some(function (t) { return t.k === "d"; })) { sec.kind = "date"; sec.ampm = toks.some(function (t) { return t.k === "ampm"; }); return sec; }
    if (toks.some(function (t) { return t.k === "gen"; }) && !toks.some(function (t) { return t.k === "ph"; })) { sec.kind = "gen"; sec.pre = lits(toks, 0, idx(toks, "gen")); sec.suf = lits(toks, idx(toks, "gen") + 1, toks.length); return sec; }
    if (!toks.some(function (t) { return t.k === "ph"; })) { sec.kind = toks.some(function (t) { return t.k === "at"; }) ? "text" : "lit"; return sec; }
    sec.kind = "num";
    var first = -1, last = -1;
    toks.forEach(function (t, i) { if (t.k === "ph" || t.k === "dot" || t.k === "E" || (t.k === "comma" && first >= 0)) { if (first < 0) first = i; last = i; } });
    // trailing commas right after the last placeholder scale by 1000
    var scale = 0, j = last;
    while (j >= first && toks[j].k === "comma") { scale++; j--; }
    last = j;
    var body = toks.slice(first, last + 1), dot = -1, ePos = -1;
    body.forEach(function (t, i) { if (t.k === "dot" && dot < 0) dot = i; if (t.k === "E" && ePos < 0) ePos = i; });
    var intEnd = dot >= 0 ? dot : ePos >= 0 ? ePos : body.length;
    var intPart = body.slice(0, intEnd), decPart = dot >= 0 ? body.slice(dot + 1, ePos >= 0 ? ePos : body.length) : [];
    sec.minInt = intPart.filter(function (t) { return t.k === "ph" && t.v === "0"; }).length;
    sec.group = intPart.some(function (t) { return t.k === "comma"; });
    var dp = decPart.filter(function (t) { return t.k === "ph"; });
    sec.decMax = dp.length;
    sec.decMin = 0; for (var k = 0; k < dp.length && dp[k].v === "0"; k++) sec.decMin++;
    sec.sci = ePos >= 0 ? { sign: body[ePos].v, digits: body.slice(ePos + 1).filter(function (t) { return t.k === "ph"; }).length || 1 } : null;
    sec.scale = scale;
    sec.pct = toks.filter(function (t) { return t.k === "pct"; }).length;
    sec.pre = lits(toks, 0, first);
    sec.suf = lits(toks, last + 1 + scale, toks.length);
    return sec;
  }
  function idx(toks, k) { for (var i = 0; i < toks.length; i++) if (toks[i].k === k) return i; return -1; }
  function lits(toks, a, b) {
    var s = "";
    for (var i = a; i < b; i++) { var t = toks[i]; if (t.k === "lit") s += t.v; else if (t.k === "pct") s += "%"; else if (t.k === "comma") s += ","; }
    return s;
  }
  function compile(code) {
    var c = cache.get(code);
    if (c) return c;
    c = splitSections(code).map(compileSection);
    if (cache.size > 500) cache.clear();
    cache.set(code, c);
    return c;
  }
  function fixedHalfUp(x, d) {
    d = Math.min(d, 20);
    var m = Math.pow(10, d), y = Math.round(+(x * m).toPrecision(15));
    var s = (y / m).toFixed(d);
    return /e/.test(s) ? x.toFixed(d) : s;
  }
  function group(intStr) { return intStr.replace(/\B(?=(\d{3})+(?!\d))/g, ","); }
  function fmtNum(sec, v, autoMinus) {
    var neg = v < 0, x = Math.abs(v) * Math.pow(100, sec.pct) / Math.pow(1000, sec.scale), body;
    if (sec.sci) {
      var exp = x === 0 ? 0 : Math.floor(Math.log10(x)), mant = x === 0 ? 0 : x / Math.pow(10, exp);
      var ms = fixedHalfUp(mant, sec.decMax);
      if (+ms >= 10) { exp++; ms = fixedHalfUp(mant / 10, sec.decMax); }
      var ee = String(Math.abs(exp)); while (ee.length < sec.sci.digits) ee = "0" + ee;
      body = trimDec(ms, sec) + "E" + (exp < 0 ? "-" : sec.sci.sign === "+" ? "+" : "") + ee;
    } else {
      body = trimDec(fixedHalfUp(x, sec.decMax), sec);
    }
    var zero = !/[1-9]/.test(body);
    return (neg && autoMinus && !zero ? "-" : "") + sec.pre + body + sec.suf;
  }
  function trimDec(s, sec) {
    var p = s.split("."), ip = p[0], dp = p[1] || "";
    while (dp.length > sec.decMin && dp[dp.length - 1] === "0") dp = dp.slice(0, -1);
    if (ip === "0" && sec.minInt === 0) ip = "";
    while (ip.length < sec.minInt) ip = "0" + ip;
    if (sec.group) ip = group(ip);
    return ip + (dp.length ? "." + dp : "");
  }
  function pad(n, w) { n = String(n); while (n.length < w) n = "0" + n; return n; }
  function fmtDate(sec, v) {
    if (v < 0 || v >= 2958466) return "#####";
    var D = ymd(v), T = hms(v), toks = sec.toks, out = "";
    if (Math.round((v - Math.floor(v)) * 86400) >= 86400) { D = ymd(v + 1); T = { h: 0, mi: 0, s: 0 }; }
    for (var i = 0; i < toks.length; i++) {
      var t = toks[i];
      if (t.k === "lit") { out += t.v; continue; }
      if (t.k === "comma") { out += ","; continue; }
      if (t.k === "dot") { out += "."; continue; }
      if (t.k === "ph") { if (t.v === "0") out += "0"; continue; }
      if (t.k === "ampm") { out += t.v === "AMPM" ? (T.h < 12 ? "AM" : "PM") : (T.h < 12 ? "A" : "P"); continue; }
      if (t.k !== "d") continue;
      var f = t.v, c = f[0];
      if (t.el) { var tot = v * 24; out += c === "h" ? Math.floor(tot) : c === "m" ? Math.floor(tot * 60) : Math.round(tot * 3600); continue; }
      if (c === "m" && f.length <= 2 && isMinute(toks, i)) c = "M";
      if (c === "y") out += f.length <= 2 ? pad(D.y % 100, 2) : D.y;
      else if (c === "m") out += f.length === 1 ? D.m : f.length === 2 ? pad(D.m, 2) : f.length === 3 ? MONTHS[D.m - 1].slice(0, 3) : f.length === 5 ? MONTHS[D.m - 1][0] : MONTHS[D.m - 1];
      else if (c === "d") out += f.length === 1 ? D.d : f.length === 2 ? pad(D.d, 2) : f.length === 3 ? DAYS[weekday(v)].slice(0, 3) : DAYS[weekday(v)];
      else if (c === "h") { var h = sec.ampm ? (T.h % 12 || 12) : T.h; out += f.length === 1 ? h : pad(h, 2); }
      else if (c === "M") out += f.length === 1 ? T.mi : pad(T.mi, 2);
      else if (c === "s") out += f.length === 1 ? T.s : pad(T.s, 2);
    }
    return out;
  }
  function isMinute(toks, i) {
    for (var a = i - 1; a >= 0; a--) { var t = toks[a]; if (t.k === "d") return t.v[0] === "h"; }
    for (var b = i + 1; b < toks.length; b++) { var u = toks[b]; if (u.k === "d") return u.v[0] === "s"; }
    return false;
  }
  // value + format code -> { t: text, color }
  function format(v, code) {
    if (v === null || v === undefined) return { t: "" };
    if (v instanceof C.Err) return { t: v.e };
    if (typeof v === "boolean") return { t: v ? "TRUE" : "FALSE" };
    if (!code || code === "General") return { t: typeof v === "number" ? general(v) : String(v) };
    var secs = compile(code), sec;
    if (typeof v === "string") {
      sec = secs[3] || (secs.length === 1 && secs[0].kind === "text" ? secs[0] : null);
      if (!sec) return { t: v };
      return { t: sec.toks.map(function (t) { return t.k === "at" ? v : t.k === "lit" ? t.v : ""; }).join(""), color: sec.color };
    }
    if (!isFinite(v)) return { t: "#NUM!" };
    var auto = true;
    if (v < 0 && secs.length >= 2 && secs[1].toks.length) { sec = secs[1]; auto = false; }
    else if (v === 0 && secs.length >= 3 && secs[2].toks.length) sec = secs[2];
    else sec = secs[0];
    var t;
    if (sec.kind === "date") t = fmtDate(sec, v);
    else if (sec.kind === "num") t = fmtNum(sec, auto ? v : Math.abs(v), auto);
    else if (sec.kind === "gen") t = sec.pre + general(auto ? v : Math.abs(v)) + sec.suf;
    else if (sec.kind === "lit") t = sec.toks.map(function (x) { return x.k === "lit" ? x.v : ""; }).join("");
    else t = general(v);
    return { t: t, color: sec.color };
  }
  function isDateFmt(code) {
    if (!code || code === "General") return false;
    var s = compile(code)[0];
    return s.kind === "date";
  }
  function isTimeOnly(code) { return isDateFmt(code) && !/[yd]/i.test(code.replace(/"[^"]*"/g, "")) ; }
  function decimals(code) {
    if (!code || code === "General") return null;
    var s = compile(code)[0];
    return s.kind === "num" ? s.decMax : null;
  }
  // Add or remove one decimal place from a number format (toolbar buttons).
  function bumpDecimals(code, delta, sample) {
    if (!code || code === "General") {
      var d = 0;
      if (typeof sample === "number" && !Number.isInteger(sample)) { var g = general(sample), i = g.indexOf("."); d = i < 0 ? 0 : g.length - i - 1; }
      d = Math.max(0, d + delta);
      return d ? "0." + "0".repeat(d) : "0";
    }
    return splitSections(code).map(function (sec) {
      if (delta > 0) {
        if (/0\.0*|#\.#*/.test(sec)) return sec.replace(/(\.[0#]*)/, "$10");
        return sec.replace(/([0#])(?![0#,])/, "$1.0");
      }
      return sec.replace(/\.([0#]*)[0#]/, function (m0, keep) { return keep.length ? "." + keep : ""; });
    }).join(";");
  }

  // ---------- parsing what people type ----------
  var CUR = { "$": "$", "€": "€", "£": "£", "jod": "JOD", "jd": "JD", "sar": "SAR", "aed": "AED", "usd": "USD", "eur": "EUR", "egp": "EGP", "kwd": "KWD" };
  var NUMRX = /^([+-]?)(\d{1,3}(?:,\d{3})+|\d*)(\.\d+)?(?:[eE]([+-]?\d+))?$/;
  var settings = { dateOrder: "DMY" };
  function num(s) {
    var m = NUMRX.exec(s);
    if (!m || (!m[2] && !m[3])) return null;
    var v = parseFloat((m[1] || "") + m[2].replace(/,/g, "") + (m[3] || "") + (m[4] ? "e" + m[4] : ""));
    return isFinite(v) ? { v: v, grouped: m[2].indexOf(",") >= 0, dec: m[3] ? m[3].length - 1 : 0 } : null;
  }
  function monthIdx(s) {
    s = s.toLowerCase();
    for (var i = 0; i < 12; i++) if (MONTHS[i].toLowerCase() === s || MONTHS[i].slice(0, 3).toLowerCase() === s || (s.length >= 3 && MONTHS[i].toLowerCase().indexOf(s) === 0)) return i + 1;
    return 0;
  }
  function timeFrac(h, mi, s, ap) {
    h = +h; mi = +mi; s = +(s || 0);
    if (ap) { ap = ap.toLowerCase(); if (h === 12) h = 0; if (ap[0] === "p") h += 12; }
    if (h > 23 || mi > 59 || s > 59) return null;
    return (h * 3600 + mi * 60 + s) / 86400;
  }
  var TIME = "(?:[ T]+(\\d{1,2}):(\\d{2})(?::(\\d{2}))?\\s*([ap]m?)?)?";
  function parseDate(t) {
    var m, y, mo, d, tm = null, rx;
    if ((m = new RegExp("^(\\d{4})[-/.](\\d{1,2})[-/.](\\d{1,2})" + TIME + "$", "i").exec(t))) { y = +m[1]; mo = +m[2]; d = +m[3]; rx = m.slice(4); }
    else if ((m = new RegExp("^(\\d{1,2})[-/.](\\d{1,2})[-/.](\\d{2,4})" + TIME + "$", "i").exec(t))) {
      var a = +m[1], b = +m[2]; y = +m[3]; if (m[3].length === 2) y += y < 30 ? 2000 : 1900;
      if (a > 12 || (b <= 12 && settings.dateOrder === "DMY")) { d = a; mo = b; } else { mo = a; d = b; }
      if (mo > 12) { var sw = mo; mo = d; d = sw; }
      rx = m.slice(4);
    }
    else if ((m = new RegExp("^(\\d{1,2})[ -]([A-Za-z]{3,9})\\.?[ -,]+(\\d{4})" + TIME + "$", "i").exec(t)) && monthIdx(m[2])) { d = +m[1]; mo = monthIdx(m[2]); y = +m[3]; rx = m.slice(4); }
    else if ((m = new RegExp("^([A-Za-z]{3,9})\\.? (\\d{1,2}),? (\\d{4})" + TIME + "$", "i").exec(t)) && monthIdx(m[1])) { mo = monthIdx(m[1]); d = +m[2]; y = +m[3]; rx = m.slice(4); }
    else return null;
    if (mo < 1 || mo > 12 || d < 1 || d > 31 || y < 1900 || y > 9999) return null;
    if (d > daysIn(y, mo)) return null;
    var v = serial(y, mo, d), fmt = /^\d{4}/.test(t) ? "yyyy-mm-dd" : /[A-Za-z]/.test(t) ? "d mmm yyyy" : settings.dateOrder === "DMY" ? "dd/mm/yyyy" : "mm/dd/yyyy";
    if (rx && rx[0] !== undefined) { tm = timeFrac(rx[0], rx[1], rx[2], rx[3]); if (tm === null) return null; v += tm; fmt += " h:mm" + (rx[3] ? " AM/PM" : ""); }
    return { v: v, nf: fmt };
  }
  function daysIn(y, m) { return y === 1900 && m === 2 ? 29 : Math.round((utc(y, m, 1) - utc(y, m - 1, 1)) / MS); }
  // text -> { f } (formula) | { v, nf?, text? }
  function parseInput(raw) {
    var s = raw == null ? "" : String(raw);
    if (s === "") return { v: null };
    if (s[0] === "'") return { v: s.slice(1), text: true };
    if (s[0] === "=" && s.length > 1) return { f: s.slice(1) };
    var t = s.trim();
    if (/^true$/i.test(t)) return { v: true };
    if (/^false$/i.test(t)) return { v: false };
    var n = num(t), m;
    if (n) return { v: n.v, nf: n.grouped ? (n.dec ? "#,##0." + "0".repeat(n.dec) : "#,##0") : undefined };
    if ((m = /^\((.+)\)$/.exec(t)) && (n = num(m[1])) && n.v >= 0) return { v: -n.v };
    if ((m = /^(.+?)\s*%$/.exec(t)) && (n = num(m[1]))) return { v: C.r15(n.v / 100), nf: n.dec ? "0." + "0".repeat(n.dec) + "%" : "0%" };
    if ((m = /^([+-]?)\s*(\$|€|£|[A-Za-z]{2,3})\s*([\d.,]+)$/.exec(t)) && CUR[m[2].toLowerCase()] && (n = num(m[3]))) return cur(CUR[m[2].toLowerCase()], m[1] === "-" ? -n.v : n.v, true);
    if ((m = /^([+-]?[\d.,]+)\s*(\$|€|£|[A-Za-z]{2,3})$/.exec(t)) && CUR[m[2].toLowerCase()] && (n = num(m[1]))) return cur(CUR[m[2].toLowerCase()], n.v, false);
    var d = parseDate(t);
    if (d) return d;
    if ((m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([ap]m?)?$/i.exec(t))) {
      var tf = timeFrac(m[1], m[2], m[3], m[4]);
      if (tf !== null) return { v: tf, nf: m[4] ? "h:mm AM/PM" : m[3] ? "h:mm:ss" : "h:mm" };
    }
    return { v: s };
  }
  function cur(sym, v, before) {
    var three = sym === "JOD" || sym === "JD" || sym === "KWD", dec = three ? "0.000" : "0.00";
    var sy = sym.length > 1 ? '"' + sym + ' "' : sym;
    return { v: v, nf: before || sym.length === 1 ? sy + "#,##" + dec : "#,##" + dec + ' "' + sym + '"' };
  }
  // Number formats offered in the toolbar.
  var PRESETS = [
    { id: "General", label: "Automatic" },
    { id: "#,##0.00", label: "Number 1,234.56" },
    { id: "#,##0", label: "Integer 1,235" },
    { id: "0.00%", label: "Percent 12.34%" },
    { id: "0%", label: "Percent 12%" },
    { id: "$#,##0.00", label: "Dollar $1,234.56" },
    { id: '"JOD "#,##0.000', label: "Dinar JOD 1,234.560" },
    { id: "€#,##0.00", label: "Euro €1,234.56" },
    { id: "£#,##0.00", label: "Pound £1,234.56" },
    { id: '"SAR "#,##0.00', label: "Riyal SAR 1,234.56" },
    { id: '"AED "#,##0.00', label: "Dirham AED 1,234.56" },
    { id: "#,##0.00;[Red]-#,##0.00", label: "Negatives in red" },
    { id: "0.00E+00", label: "Scientific 1.23E+03" },
    { id: "yyyy-mm-dd", label: "Date 2026-10-27" },
    { id: "dd/mm/yyyy", label: "Date 27/10/2026" },
    { id: "mm/dd/yyyy", label: "Date 10/27/2026" },
    { id: "d mmm yyyy", label: "Date 27 Oct 2026" },
    { id: "dddd, mmmm d, yyyy", label: "Long date" },
    { id: "h:mm", label: "Time 14:30" },
    { id: "h:mm AM/PM", label: "Time 2:30 PM" },
    { id: "yyyy-mm-dd h:mm", label: "Date + time" },
    { id: "@", label: "Plain text" }
  ];

  SH.fmt = {
    serial: serial, ymd: ymd, weekday: weekday, hms: hms, fromDate: fromDate, today: today, now: now, daysIn: daysIn,
    general: general, numText: numText, format: format, compile: compile, isDateFmt: isDateFmt, isTimeOnly: isTimeOnly,
    decimals: decimals, bumpDecimals: bumpDecimals, parseInput: parseInput, parseDate: parseDate, PRESETS: PRESETS,
    MONTHS: MONTHS, DAYS: DAYS, settings: settings, fixedHalfUp: fixedHalfUp
  };
})();
