/* Sheets UI: menus, toolbar, formula bar, name box, sheet tabs, status bar, dialogs (find/replace, sort,
   filter, conditional formatting, charts, number format, workbooks, help), files and autosave. */
(function () {
  "use strict";
  var SH = window.SH, C = SH.core, F = SH.fmt, EN = SH.engine, M = SH.model, IO = SH.io, $ = S3.$, h = S3.h;
  var MAXR = C.MAXR, MAXC = C.MAXC;
  var app = SH.app = { wb: null, grid: null };

  // ---------- small DOM helpers ----------
  function closeAll() { S3.$$(".sh-pop").forEach(function (p) { p.remove(); }); S3.$$("[aria-expanded=true]").forEach(function (b) { b.setAttribute("aria-expanded", "false"); }); }
  function popAt(x, y, el, anchor) {
    closeAll();
    el.classList.add("sh-pop");
    document.body.appendChild(el);
    var w = el.offsetWidth, hh = el.offsetHeight, vw = document.documentElement.clientWidth, vh = window.innerHeight;
    var left = Math.max(6, Math.min(x, vw - w - 6)), top = y + hh > vh - 6 ? Math.max(6, y - hh - (anchor ? anchor.offsetHeight + 4 : 0)) : y;
    el.style.left = left + window.scrollX + "px"; el.style.top = top + window.scrollY + "px";
    if (anchor) anchor.setAttribute("aria-expanded", "true");
    var first = el.querySelector("button:not([disabled]),input,select"); if (first) first.focus({ preventScroll: true });
    return el;
  }
  function menu(items, x, y, anchor) {
    var el = h("div", { role: "menu", "class": "sh-menu" });
    items.forEach(function (it) {
      if (it === "-") { el.appendChild(h("div", { "class": "sh-sep", role: "separator" })); return; }
      if (it.hidden) return;
      var b = h("button", { type: "button", role: "menuitem", "class": "sh-mi" + (it.on ? " on" : ""), disabled: it.disabled ? true : false }, [
        h("span", { "class": "sh-mck", text: it.on ? "✓" : "" }), h("span", { "class": "sh-ml", text: it.label }), h("span", { "class": "sh-mk", text: it.kbd || "" })]);
      b.addEventListener("click", function () { closeAll(); app.grid.focus(); it.act(); });
      el.appendChild(b);
    });
    el.addEventListener("keydown", function (e) {
      var bs = S3.$$("button:not([disabled])", el), i = bs.indexOf(document.activeElement);
      if (e.key === "ArrowDown") { e.preventDefault(); bs[(i + 1) % bs.length].focus(); }
      if (e.key === "ArrowUp") { e.preventDefault(); bs[(i - 1 + bs.length) % bs.length].focus(); }
      if (e.key === "Escape") { e.preventDefault(); closeAll(); (anchor || app.grid).focus(); }
    });
    return popAt(x, y, el, anchor);
  }
  function menuFor(btn, itemsFn) {
    btn.setAttribute("aria-haspopup", "menu"); btn.setAttribute("aria-expanded", "false");
    btn.addEventListener("click", function (e) {
      e.stopPropagation();
      if (btn.getAttribute("aria-expanded") === "true") { closeAll(); return; }
      var r = btn.getBoundingClientRect(); menu(itemsFn(), r.left, r.bottom + 4, btn);
    });
  }
  document.addEventListener("pointerdown", function (e) { if (!e.target.closest(".sh-pop") && !e.target.closest("[aria-expanded=true]")) closeAll(); });
  function dialog(title, body, actions, opts) {
    opts = opts || {};
    var dlg = h("dialog", { "class": "sh-dlg" + (opts.wide ? " wide" : "") + (opts.side ? " side" : ""), "aria-label": title });
    var head = h("div", { "class": "row between sh-dhead" }, [h("h2", { text: title }), h("button", { "class": "btn ghost icon small", type: "button", "aria-label": "Close", html: '<svg class="i"><use href="#i-x"/></svg>' })]);
    var foot = h("div", { "class": "row sh-dfoot" });
    (actions || []).forEach(function (a) { var b = h("button", { "class": "btn small" + (a.primary ? " primary" : "") + (a.ghost ? " ghost" : ""), type: "button", text: a.label }); b.addEventListener("click", function () { if (a.act() !== false && !a.keep) close(); }); foot.appendChild(b); if (a.ref) a.ref(b); });
    dlg.appendChild(head); dlg.appendChild(body); if (actions && actions.length) dlg.appendChild(foot);
    document.body.appendChild(dlg);
    function close() { try { dlg.close(); } catch (e) { /* closed */ } dlg.remove(); app.grid.focus(); if (opts.onClose) opts.onClose(); }
    head.lastChild.addEventListener("click", close);
    dlg.addEventListener("close", function () { if (dlg.isConnected) { dlg.remove(); if (opts.onClose) opts.onClose(); } });
    dlg.addEventListener("keydown", function (e) { if (e.key === "Enter" && opts.enter && e.target.tagName !== "TEXTAREA" && e.target.tagName !== "BUTTON") { e.preventDefault(); if (opts.enter() !== false) close(); } });
    if (opts.side && dlg.show) dlg.show(); else if (dlg.showModal) dlg.showModal(); else dlg.setAttribute("open", "");
    var f = dlg.querySelector("input,select,textarea"); if (f) { f.focus(); if (f.select && f.type === "text") f.select(); }
    return { el: dlg, close: close };
  }
  function field(label, input) { return h("label", { "class": "field" }, [h("span", { text: label }), input]); }
  function sel(options, value) { var s = h("select", { "class": "select" }); options.forEach(function (o) { var op = h("option", { value: o[0], text: o[1] }); if (o[0] === value) op.selected = true; s.appendChild(op); }); return s; }
  function inp(value, attrs) { var i = h("input", Object.assign({ "class": "input", type: "text", value: value == null ? "" : String(value), dir: "auto", autocomplete: "off" }, attrs || {})); return i; }
  function colorInp(v) { return h("input", { type: "color", "class": "sh-color", value: v || "#000000" }); }
  function guard(fn) { return function () { try { return fn.apply(null, arguments); } catch (e) { S3.toast(e.message || "Something went wrong.", "error"); } }; }

  // ---------- selection helpers ----------
  function sheet() { return app.wb.sheet(); }
  function S() { return app.grid.sel; }
  function selRange() { var s = S(); return { r0: s.r0, c0: s.c0, r1: s.r1, c1: s.c1 }; }
  function dataRange() {
    var s = S(), sh = sheet();
    if (s.r0 === s.r1 && s.c0 === s.c1) return app.wb.region(sh, s.r, s.c);
    return app.wb.clampArea(sh, s);
  }
  function activeStyle() { var s = S(); return app.wb.style(sheet(), s.r, s.c) || {}; }
  function fmt(patch) { guard(function () { app.wb.styleRange(sheet(), selRange(), patch); })(); }
  function toggle(key) { var on = !activeStyle()[key], p = {}; p[key] = on ? true : null; fmt(p); }

  // ---------- actions ----------
  var A = app.actions = {
    undo: function () { var l = app.wb.undo(); if (!l) S3.toast("Nothing to undo"); afterChange(); },
    redo: function () { var l = app.wb.redo(); if (!l) S3.toast("Nothing to redo"); afterChange(); },
    bold: function () { toggle("b"); }, italic: function () { toggle("i"); }, underline: function () { toggle("u"); }, strike: function () { toggle("st"); },
    wrap: function () { toggle("wr"); },
    align: function (a) { fmt({ ha: activeStyle().ha === a ? null : a }); },
    valign: function (a) { fmt({ va: a === "bottom" ? null : a }); },
    numfmt: function (code) { fmt({ nf: code === "General" ? null : code }); },
    decimals: function (d) { var s = S(), v = app.wb.value(sheet(), s.r, s.c); fmt({ nf: F.bumpDecimals(activeStyle().nf, d, v) }); },
    merge: guard(function () {
      var s = S(), sh = sheet(), g = app.wb.clampArea(sh, s);
      if (app.grid.isMergeSel() || sh.merges.some(function (m) { return C.overlaps(m, g) && (m.r0 !== m.r1 || m.c0 !== m.c1); }) && g.r0 === g.r1 && g.c0 === g.c1) { app.wb.unmerge(sh, g); return; }
      if (g.r0 === g.r1 && g.c0 === g.c1) { S3.toast("Select two or more cells to merge."); return; }
      var lost = 0; app.wb.forCells(sh, g, function (r, c, x) { if ((r !== g.r0 || c !== g.c0) && (x.v !== undefined || x.f !== undefined)) lost++; });
      if (lost && !window.confirm("Merging keeps only the top-left value. " + lost + " other value(s) will be removed. Merge anyway?")) return;
      app.wb.merge(sh, g); app.grid.select(g.r0, g.c0, g.r1, g.c1, g.r0, g.c0, true);
    }),
    insertRows: guard(function (below) { var s = S(), n = s.r1 - s.r0 + 1; if (n > 5000) n = 1; app.wb.insertDelete(sheet(), "r", below ? s.r1 + 1 : s.r0, n); }),
    insertCols: guard(function (right) { var s = S(), n = s.c1 - s.c0 + 1; if (n > 500) n = 1; app.wb.insertDelete(sheet(), "c", right ? s.c1 + 1 : s.c0, n); }),
    deleteRows: guard(function () { var s = S(), sh = sheet(), r1 = Math.min(s.r1, Math.max(sh.maxR, s.r0)); app.wb.insertDelete(sh, "r", s.r0, -(r1 - s.r0 + 1)); app.grid.select(s.r0, s.c, s.r0, s.c); }),
    deleteCols: guard(function () { var s = S(), sh = sheet(), c1 = Math.min(s.c1, Math.max(sh.maxC, s.c0)); app.wb.insertDelete(sh, "c", s.c0, -(c1 - s.c0 + 1)); app.grid.select(s.r, s.c0, s.r, s.c0); }),
    clear: function (what) { app.wb.clear(sheet(), selRange(), what); },
    sort: guard(function (asc) {
      var sh = sheet(), s = S(), g = dataRange();
      var header = g.r1 > g.r0 && looksHeader(sh, g);
      app.wb.sortRange(sh, g, [{ c: C.clamp(s.c, g.c0, g.c1), asc: asc }], header);
      S3.toast("Sorted " + (g.r1 - g.r0 + (header ? 0 : 1)) + " rows" + (header ? " (kept the header row)" : ""));
    }),
    filter: guard(function () {
      var sh = sheet();
      if (sh.filter) { app.wb.setFilter(sh, null); S3.toast("Filter removed"); return; }
      var g = dataRange();
      if (g.r1 <= g.r0) { S3.toast("Select a table with a header row first.", "error"); return; }
      app.wb.setFilter(sh, g); S3.toast("Filter on: use the buttons in the header row");
    }),
    freeze: function (r, c) { var sh = sheet(); app.wb.setMeta(sh, "Freeze", function () { sh.freeze = { r: Math.max(0, Math.min(200, r)), c: Math.max(0, Math.min(50, c)) }; }); },
    rtl: function () { var sh = sheet(); app.wb.setMeta(sh, "Direction", function () { sh.rtl = !sh.rtl; }); S3.toast(sh.rtl ? "Right-to-left sheet" : "Left-to-right sheet"); },
    gridlines: function () { var sh = sheet(); app.wb.setMeta(sh, "Gridlines", function () { sh.grid = sh.grid === false; }); },
    darkGrid: function () { app.grid.dark = !app.grid.dark; S3.store("sheets.darkGrid", app.grid.dark); app.grid.draw(); },
    autosum: guard(function (fn) {
      fn = fn || "SUM";
      var sh = sheet(), s = S(), wb = app.wb;
      if (s.r0 === s.r1 && s.c0 === s.c1) {
        var r = s.r - 1, c = s.c;
        while (r >= 0 && typeof wb.value(sh, r, c) === "number") r--;
        var text;
        if (r < s.r - 1) text = "=" + fn + "(" + C.addr(r + 1, c) + ":" + C.addr(s.r - 1, c) + ")";
        else { var cc = s.c - 1; while (cc >= 0 && typeof wb.value(sh, s.r, cc) === "number") cc--; text = cc < s.c - 1 ? "=" + fn + "(" + C.addr(s.r, cc + 1) + ":" + C.addr(s.r, s.c - 1) + ")" : "=" + fn + "()"; }
        app.grid.beginEdit("edit", text);
        if (/\(\)$/.test(text)) { var p = text.length - 1; app.grid.editor.setSelectionRange(p, p); }
        return;
      }
      var g = wb.clampArea(sh, s);
      wb.run("AutoSum", function () { for (var c2 = g.c0; c2 <= g.c1; c2++) wb.setInput(sh, g.r1 + 1, c2, "=" + fn + "(" + C.addr(g.r0, c2) + ":" + C.addr(g.r1, c2) + ")"); });
    }),
    fillDown: guard(function () { var s = S(), sh = sheet(); if (s.r1 === s.r0) return; app.wb.pasteBlock(sh, s.r0 + 1, s.c0, app.wb.copyBlock(sh, { r0: s.r0, c0: s.c0, r1: s.r0, c1: s.c1 }), "all", { r0: s.r0 + 1, c0: s.c0, r1: app.wb.clampArea(sh, s).r1, c1: s.c1 }); }),
    fillRight: guard(function () { var s = S(), sh = sheet(); if (s.c1 === s.c0) return; app.wb.pasteBlock(sh, s.r0, s.c0 + 1, app.wb.copyBlock(sh, { r0: s.r0, c0: s.c0, r1: s.r1, c1: s.c0 }), "all", { r0: s.r0, c0: s.c0 + 1, r1: s.r1, c1: app.wb.clampArea(sh, s).c1 }); }),
    today: function () { var s = S(), sh = sheet(); app.wb.run("Date", function () { var d = F.ymd(F.today()); app.wb.setInput(sh, s.r, s.c, d.y + "-" + pad(d.m) + "-" + pad(d.d)); }); },
    pasteValues: guard(function () { var cl = app.grid.clip, s = S(); if (!cl) { S3.toast("Copy cells in this sheet first (Ctrl+C)."); return; } app.wb.pasteBlock(sheet(), s.r0, s.c0, cl.block, "values", s); }),
    pasteFormats: guard(function () { var cl = app.grid.clip, s = S(); if (!cl) { S3.toast("Copy cells in this sheet first (Ctrl+C)."); return; } app.wb.pasteBlock(sheet(), s.r0, s.c0, cl.block, "formats", s); }),
    copy: function () { var t = app.grid.copy(null, false); if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(t).then(function () { S3.toast("Copied"); }, function () { S3.toast("Press Ctrl+C to copy"); }); },
    cut: function () { var t = app.grid.copy(null, true); if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(t).catch(function () { /* ignore */ }); },
    paste: function () {
      if (navigator.clipboard && navigator.clipboard.readText) navigator.clipboard.readText().then(function (t) { app.grid.paste(t); }, function () { S3.toast("Press Ctrl+V to paste"); });
      else S3.toast("Press Ctrl+V to paste");
    }
  };
  function pad(n) { return n < 10 ? "0" + n : String(n); }
  function looksHeader(sh, g) {
    var top = 0, below = 0;
    for (var c = g.c0; c <= g.c1; c++) { var a = app.wb.value(sh, g.r0, c), b = app.wb.value(sh, g.r0 + 1, c); if (typeof a === "string") top++; if (typeof b === "number") below++; }
    return top > 0 && (top === g.c1 - g.c0 + 1 || below > 0);
  }
  function afterChange() { app.grid.axKey = ""; app.grid.select(S().r0, S().c0, S().r1, S().c1, S().r, S().c, true); renderTabs(); }

  // ---------- sheets ----------
  function switchSheet(i) {
    if (i === app.wb.active || i < 0 || i >= app.wb.sheets.length) return;
    if (app.grid.edit) app.grid.commit(0, 0);
    app.sels[sheet().uid] = Object.assign({}, S(), { sx: Math.abs(app.grid.scroller.scrollLeft), sy: app.grid.scroller.scrollTop });
    app.wb.active = i; app.grid.axKey = "";
    var s = app.sels[sheet().uid] || { r: 0, c: 0, r0: 0, c0: 0, r1: 0, c1: 0, sx: 0, sy: 0 };
    app.grid.drawNow();
    app.grid.scroller.scrollTop = s.sy; app.grid.scroller.scrollLeft = sheet().rtl ? -s.sx : s.sx;
    app.grid.select(s.r0, s.c0, s.r1, s.c1, s.r, s.c, true);
    renderTabs(); queueSave();
  }
  function renderTabs() {
    var box = $("#tabs"), wb = app.wb;
    box.textContent = "";
    wb.sheets.forEach(function (sh, i) {
      var b = h("button", { "class": "sh-tab" + (i === wb.active ? " on" : ""), type: "button", role: "tab", "aria-selected": String(i === wb.active), dir: "auto", text: sh.name, title: "Double-click to rename, right-click for more" });
      b.addEventListener("click", function () { if (i === wb.active) tabMenu(b, sh); else switchSheet(i); });
      b.addEventListener("dblclick", function () { renameTab(b, sh); });
      b.addEventListener("contextmenu", function (e) { e.preventDefault(); if (i !== wb.active) switchSheet(i); tabMenu(b, sheet()); });
      box.appendChild(b);
    });
    var on = box.querySelector(".on"); if (on && on.scrollIntoView) on.scrollIntoView({ block: "nearest", inline: "nearest" });
    $("#btnUndo").disabled = !wb.canUndo(); $("#btnRedo").disabled = !wb.canRedo();
  }
  function tabMenu(b, sh) {
    var r = b.getBoundingClientRect(), wb = app.wb, i = wb.sheets.indexOf(sh);
    menu([
      { label: "Rename", act: function () { renameTab(S3.$$(".sh-tab")[i], sh); } },
      { label: "Duplicate", act: function () { wb.duplicateSheet(sh); afterSheetChange(); } },
      { label: "Move left", disabled: i === 0, act: function () { wb.moveSheet(sh, -1); afterSheetChange(); } },
      { label: "Move right", disabled: i === wb.sheets.length - 1, act: function () { wb.moveSheet(sh, 1); afterSheetChange(); } },
      { label: sh.rtl ? "Left-to-right" : "Right-to-left", act: A.rtl },
      "-",
      { label: "Delete sheet", disabled: wb.sheets.length < 2, act: function () { if (window.confirm("Delete the sheet “" + sh.name + "”? You can undo this.")) { wb.deleteSheet(sh); afterSheetChange(); } } }
    ], r.left, r.top - 4, b);
  }
  function renameTab(b, sh) {
    if (!b) return;
    var i = inp(sh.name, { "class": "sh-tabin", maxlength: "31", "aria-label": "Sheet name" }), done = false;
    b.replaceWith(i); i.focus(); i.select();
    function fin(ok) {
      if (done) return; done = true;
      if (ok && i.value.trim() && i.value.trim() !== sh.name && !app.wb.renameSheet(sh, i.value)) S3.toast("That name is empty or already used.", "error");
      renderTabs(); app.grid.focus();
    }
    i.addEventListener("keydown", function (e) { if (e.key === "Enter") fin(true); if (e.key === "Escape") fin(false); });
    i.addEventListener("blur", function () { fin(true); });
  }
  function afterSheetChange() { app.grid.axKey = ""; app.grid.select(0, 0, 0, 0); renderTabs(); }

  // ---------- formula bar, name box, status ----------
  function syncBar() {
    if (!app.grid) return;
    var s = S(), sh = sheet(), fx = $("#fx"), nb = $("#namebox");
    if (!app.grid.edit && document.activeElement !== fx) fx.value = app.wb.inputText(sh, s.r, s.c);
    if (document.activeElement !== nb) nb.value = s.r0 === s.r1 && s.c0 === s.c1 || app.grid.isMergeSel() ? C.addr(s.r, s.c) : C.rangeStr(s);
    var st = activeStyle();
    [["#tbBold", st.b], ["#tbItalic", st.i], ["#tbUnder", st.u], ["#tbWrap", st.wr], ["#tbLeft", st.ha === "left"], ["#tbCenter", st.ha === "center"], ["#tbRight", st.ha === "right"],
      ["#tbMerge", !!app.wb.mergeAt(sh, s.r, s.c)], ["#tbFilter", !!sh.filter], ["#tbRtl", !!sh.rtl]].forEach(function (p) { $(p[0]).setAttribute("aria-pressed", String(!!p[1])); });
    var nf = st.nf || "General", nfs = $("#tbFmt");
    nfs.value = F.PRESETS.some(function (p) { return p.id === nf; }) ? nf : "custom";
    $("#tbSize").value = String(st.fs || 13);
    $("#tbFc").style.setProperty("--sw", st.fc || "currentColor"); $("#tbBg").style.setProperty("--sw", st.bg || "transparent");
    var v = app.wb.value(sh, s.r, s.c), err = $("#cellErr");
    if (v instanceof C.Err && C.ERR_HELP[v.e]) { err.textContent = v.e + " " + C.ERR_HELP[v.e]; err.hidden = false; } else err.hidden = true;
    statusLater();
  }
  var statusLater = S3.debounce(function () {
    var s = S(), sh = sheet(), wb = app.wb, sum = 0, n = 0, cnt = 0, mn = Infinity, mx = -Infinity, out = $("#status");
    if (s.r0 === s.r1 && s.c0 === s.c1) { out.textContent = ""; return; }
    var a = wb.clampArea(sh, s), cells = (a.r1 - a.r0 + 1) * (a.c1 - a.c0 + 1);
    if (cells > 2e6) { out.textContent = ""; return; }
    for (var r = a.r0; r <= a.r1; r++) { if (sh.hidden.has(r)) continue; for (var c = a.c0; c <= a.c1; c++) { var v = wb.value(sh, r, c); if (v === null) continue; cnt++; if (typeof v === "number") { sum += v; n++; if (v < mn) mn = v; if (v > mx) mx = v; } } }
    var nf = activeStyle().nf, f = function (x) { return F.format(C.r15(x), nf && !F.isDateFmt(nf) ? nf : "#,##0.##").t; };
    out.textContent = n ? "Sum " + f(sum) + "  ·  Average " + f(sum / n) + "  ·  Min " + f(mn) + "  ·  Max " + f(mx) + "  ·  Count " + cnt : cnt ? "Count " + cnt : "";
  }, 60);
  function wireBar() {
    var fx = $("#fx"), nb = $("#namebox"), g = app.grid;
    fx.addEventListener("focus", function () { if (!g.edit) g.beginEdit("edit", fx.value, true); });
    fx.addEventListener("input", function () { if (!g.edit) g.beginEdit("edit", fx.value, true); g.editor.value = fx.value; g.edit.mode = "edit"; g.onInput(); });
    fx.addEventListener("keydown", function (e) {
      if (e.key === "Enter" && !e.altKey) { e.preventDefault(); g.commit(e.shiftKey ? -1 : 1, 0); }
      else if (e.key === "Tab") { e.preventDefault(); g.commit(0, e.shiftKey ? -1 : 1); }
      else if (e.key === "Escape") { e.preventDefault(); g.cancel(); syncBar(); }
    });
    nb.addEventListener("keydown", function (e) {
      if (e.key === "Escape") { syncBar(); g.focus(); return; }
      if (e.key !== "Enter") return;
      e.preventDefault();
      var r = C.parseRange(nb.value.trim());
      if (!r) { S3.toast("Type a cell like B12 or a range like A1:C9.", "error"); return; }
      if (r.sheet) { var sh = app.wb.sheetByName(r.sheet); if (!sh) { S3.toast("No sheet named “" + r.sheet + "”.", "error"); return; } switchSheet(app.wb.sheets.indexOf(sh)); }
      g.select(r.r0, r.c0, r.r1, r.c1, r.r0, r.c0); g.focus();
    });
  }

  // ---------- files ----------
  function openFiles(files) {
    var f = files[0];
    if (!f) return;
    if (f.size > IO.LIMITS.bytes) { S3.toast("That file is larger than 40 MB.", "error"); return; }
    var name = f.name || "file", isX = /\.xlsx$/i.test(name) || /spreadsheetml/.test(f.type), isOld = /\.xls$/i.test(name);
    if (isOld) { S3.toast("Old .xls files are not supported. Save it as .xlsx or .csv first.", "error"); return; }
    setChip("Opening " + name + "…", "");
    var p = isX ? S3.readFile(f).then(function (buf) { return IO.readXlsx(buf, name); }) : S3.readFile(f, "text").then(function (t) { return IO.readCsv(t, name, /\.tsv$/i.test(name) ? "\t" : "auto"); });
    p.then(function (res) {
      load(res.wb);
      S3.toast("Opened " + name + (res.notes.length ? " (" + res.notes.join("; ") + ")" : ""));
      S3.done();
    }).catch(function (e) { setChip("Could not open the file", "warn"); S3.toast(e && e.message ? e.message : "Could not read that file.", "error"); });
  }
  function baseName() { return (app.wb.name || "spreadsheet").replace(/[\\/:*?"<>|]+/g, "").trim().slice(0, 80) || "spreadsheet"; }
  var saveX = guard(function () {
    if (app.grid.edit) app.grid.commit(0, 0);
    setChip("Preparing .xlsx…", "");
    IO.writeXlsx(app.wb).then(function (blob) { S3.download(blob, baseName() + ".xlsx", blob.type); S3.done(); setChip("Downloaded .xlsx", "ok"); })
      .catch(function (e) { S3.toast("Could not create the .xlsx: " + (e && e.message || e), "error"); setChip("Export failed", "warn"); });
  });
  function saveCsv(delim) {
    var guardOn = S3.store("sheets.csvGuard") !== false, sh = sheet();
    var text = IO.writeCsv(app.wb, sh, null, { delim: delim, guard: guardOn });
    S3.download(new Blob([text], { type: delim === "\t" ? "text/tab-separated-values" : "text/csv" }), baseName() + (app.wb.sheets.length > 1 ? " - " + sh.name : "") + (delim === "\t" ? ".tsv" : ".csv"), "text/csv");
    S3.done();
  }
  function printNow(useSel) {
    var sh = sheet(), s = S(), g = useSel && !(s.r0 === s.r1 && s.c0 === s.c1) ? app.wb.clampArea(sh, s) : { r0: 0, c0: 0, r1: Math.max(0, sh.maxR, sh.smaxR || 0), c1: Math.max(0, sh.maxC, sh.smaxC || 0) };
    if ((g.r1 - g.r0 + 1) * (g.c1 - g.c0 + 1) > 300000) { S3.toast("That is too much to print at once. Select a smaller range.", "error"); return; }
    var charts = useSel ? [] : Array.from(app.grid.chartBoxes.values()).map(function (b) { try { return b.cv.toDataURL("image/png"); } catch (e) { return null; } }).filter(Boolean);
    var area = $("#printArea"); area.textContent = "";
    area.appendChild(h("h1", { "class": "sh-ptitle", dir: "auto", text: app.wb.name + (app.wb.sheets.length > 1 ? " — " + sh.name : "") }));
    area.appendChild(IO.printHtml(app.wb, sh, g, charts));
    document.body.classList.add("sh-printing");
    var done = function () { document.body.classList.remove("sh-printing"); area.textContent = ""; window.removeEventListener("afterprint", done); };
    window.addEventListener("afterprint", done);
    setTimeout(function () { window.print(); setTimeout(done, 1000); S3.done(); }, 50);
  }

  // ---------- autosave ----------
  var saveTimer = 0, storeOk = true;
  function setChip(t, kind) { var c = $("#saveState"); c.textContent = t; c.className = "chip sh-save" + (kind ? " " + kind : ""); }
  function queueSave() {
    if (!storeOk) return;
    clearTimeout(saveTimer); setChip("Saving…", "");
    saveTimer = setTimeout(function () {
      var wb = app.wb;
      IO.store.put(wb).then(function () { S3.store("sheets.last", wb.id); setChip("Saved on this device", "ok"); })
        .catch(function () { storeOk = false; setChip("Not saved in this browser — download a copy", "warn"); });
    }, 350);
  }
  function load(wb) {
    if (app.grid && app.grid.edit) app.grid.cancel();
    app.wb = wb; app.sels = {};
    wb.on(function () { queueSave(); renderTabs(); syncBar(); updateEmpty(); });
    $("#docName").value = wb.name;
    if (app.grid) app.grid.setWorkbook(wb);
    renderTabs(); syncBar(); updateEmpty(); queueSave();
  }
  app.load = load;
  function updateEmpty() {
    var wb = app.wb, empty = wb.sheets.length === 1 && !wb.sheet().cells.size && !wb.sheet().charts.length;
    $("#emptyCard").hidden = !empty;
  }
  function workbooks() {
    IO.store.list().then(function (list) {
      var body = h("div", { "class": "stack sh-books" });
      if (!list.length) body.appendChild(h("p", { "class": "muted", text: "Nothing saved yet. Workbooks are saved automatically in this browser as you work." }));
      list.forEach(function (it) {
        var open = h("button", { "class": "btn small", type: "button", text: it.id === app.wb.id ? "Open (current)" : "Open" }), del = h("button", { "class": "btn ghost small", type: "button", "aria-label": "Delete " + it.name, html: '<svg class="i"><use href="#i-trash"/></svg>' });
        var row = h("div", { "class": "row sh-book" }, [h("div", { "class": "grow" }, [h("b", { dir: "auto", text: it.name }), h("div", { "class": "faint", text: it.sheets + " sheet" + (it.sheets === 1 ? "" : "s") + " · " + new Date(it.updated).toLocaleString() })]), open, del]);
        open.addEventListener("click", function () { IO.store.get(it.id).then(function (rec) { try { load(M.fromJSON(rec.data)); d.close(); } catch (e) { S3.toast("That saved workbook could not be read.", "error"); } }); });
        del.addEventListener("click", function () {
          if (!window.confirm("Delete “" + it.name + "” from this browser? This cannot be undone.")) return;
          IO.store.del(it.id).then(function () { row.remove(); if (it.id === app.wb.id) { load(M.blank()); } });
        });
        body.appendChild(row);
      });
      body.appendChild(h("p", { "class": "faint", text: "Saved only in this browser on this device. Download .xlsx to keep a copy or move it elsewhere." }));
      var d = dialog("My workbooks", body, [{ label: "New workbook", primary: true, act: function () { load(M.blank()); } }]);
    }).catch(function () { S3.toast("Saved workbooks are not available in this browser (private mode?).", "error"); });
  }

  // ---------- dialogs ----------
  function findDialog(replace) {
    var q = inp(app.lastFind || "", { placeholder: "Find", "aria-label": "Find" }), rep = inp("", { placeholder: "Replace with", "aria-label": "Replace with" });
    var cs = h("input", { type: "checkbox" }), whole = h("input", { type: "checkbox" }), all = h("input", { type: "checkbox" }), formulas = h("input", { type: "checkbox" });
    var info = h("p", { "class": "faint sh-findinfo", "aria-live": "polite", text: " " }), hits = [], at = -1;
    var body = h("div", { "class": "stack" }, [field("Find", q), replace ? field("Replace with", rep) : null,
      h("div", { "class": "row" }, [h("label", { "class": "check" }, [cs, " Match case"]), h("label", { "class": "check" }, [whole, " Entire cell"]), h("label", { "class": "check" }, [all, " All sheets"]), h("label", { "class": "check" }, [formulas, " Look in formulas"])]), info]);
    function opts() { return { caseSens: cs.checked, whole: whole.checked, all: all.checked, formulas: formulas.checked }; }
    function search() { app.lastFind = q.value; hits = app.wb.find(q.value, opts()); at = -1; info.textContent = q.value ? hits.length + " match" + (hits.length === 1 ? "" : "es") : " "; }
    function next() {
      if (!hits.length || at < 0 && !q.value) search();
      if (!hits.length) { info.textContent = q.value ? "No matches" : " "; return; }
      at = (at + 1) % hits.length; var hh = hits[at];
      if (hh.sh !== sheet()) switchSheet(app.wb.sheets.indexOf(hh.sh));
      app.grid.select(hh.r, hh.c, hh.r, hh.c); info.textContent = (at + 1) + " of " + hits.length;
    }
    [q, cs, whole, all, formulas].forEach(function (el) { el.addEventListener("input", search); el.addEventListener("change", search); });
    var acts = [{ label: "Find next", primary: !replace, keep: true, act: next }];
    if (replace) {
      acts.push({ label: "Replace", keep: true, act: function () { var cur = hits[at]; if (!cur) { next(); return; } app.wb.replaceAll(q.value, rep.value, opts(), [cur]); search(); next(); } });
      acts.push({ label: "Replace all", primary: true, keep: true, act: function () { if (!q.value) return; var n = app.wb.replaceAll(q.value, rep.value, opts()); info.textContent = "Replaced " + n + " cell" + (n === 1 ? "" : "s"); search(); } });
    }
    if (q.value) search();
    dialog(replace ? "Find and replace" : "Find", body, acts, { side: true, enter: function () { next(); return false; } });
  }
  function sortDialog() {
    var sh = sheet(), g = dataRange(), head = h("input", { type: "checkbox" }), s = S();
    head.checked = looksHeader(sh, g);
    function cols() { var out = []; for (var c = g.c0; c <= g.c1; c++) { var t = head.checked ? app.wb.text(sh, g.r0, c) : ""; out.push([String(c), "Column " + C.colName(c) + (t ? " — " + t : "")]); } return out; }
    var k1 = sel(cols(), String(C.clamp(s.c, g.c0, g.c1))), o1 = sel([["1", "A → Z, smallest first"], ["0", "Z → A, largest first"]], "1");
    var k2 = sel([["", "(none)"]].concat(cols()), ""), o2 = sel([["1", "A → Z"], ["0", "Z → A"]], "1");
    head.addEventListener("change", function () { [k1, k2].forEach(function (sl, i) { var v = sl.value; sl.textContent = ""; (i ? [["", "(none)"]] : []).concat(cols()).forEach(function (o) { sl.appendChild(h("option", { value: o[0], text: o[1] })); }); sl.value = v; }); });
    var body = h("div", { "class": "stack" }, [h("p", { "class": "faint", text: "Range " + C.rangeStr(g) + " (" + (g.r1 - g.r0 + 1) + " rows)" }), h("label", { "class": "check" }, [head, " The first row is a header"]),
      h("div", { "class": "grid-2" }, [field("Sort by", k1), field("Order", o1), field("Then by", k2), field("Order", o2)])]);
    var run = guard(function () { var keys = [{ c: +k1.value, asc: o1.value === "1" }]; if (k2.value !== "") keys.push({ c: +k2.value, asc: o2.value === "1" }); app.wb.sortRange(sh, g, keys, head.checked); S3.toast("Sorted"); });
    dialog("Sort range", body, [{ label: "Cancel", ghost: true, act: function () { } }, { label: "Sort", primary: true, act: run }], { enter: run });
  }
  function filterPop(c) {
    var sh = sheet(), f = sh.filter, wb = app.wb;
    if (!f) return;
    var cur = f.cols[c] || {}, counts = new Map(), order = [];
    for (var r = f.r0 + 1; r <= f.r1 && order.length < 5000; r++) { var t = wb.text(sh, r, c); if (!counts.has(t)) { counts.set(t, 0); order.push(t); } counts.set(t, counts.get(t) + 1); }
    order.sort(function (a, b) { if (a === "") return 1; if (b === "") return -1; var na = parseFloat(a.replace(/,/g, "")), nb = parseFloat(b.replace(/,/g, "")); return !isNaN(na) && !isNaN(nb) ? na - nb : a.localeCompare(b, undefined, { numeric: true }); });
    var hidden = new Set(cur.hide || []), search = inp("", { placeholder: "Search values", "aria-label": "Search values" }), list = h("div", { "class": "sh-flist" });
    var op = sel([["", "Any value"], ["contains", "Text contains"], [">", "Greater than"], ["<", "Less than"], ["=", "Equal to"], ["<>", "Not equal to"]], cur.op || ""), val = inp(cur.val || "", { placeholder: "Value", "aria-label": "Condition value" });
    function paint() {
      list.textContent = "";
      var q = search.value.toLowerCase();
      order.filter(function (t) { return !q || t.toLowerCase().indexOf(q) >= 0; }).slice(0, 1000).forEach(function (t) {
        var cb = h("input", { type: "checkbox" }); cb.checked = !hidden.has(t);
        cb.addEventListener("change", function () { if (cb.checked) hidden.delete(t); else hidden.add(t); });
        list.appendChild(h("label", { "class": "check sh-fitem" }, [cb, h("span", { dir: "auto", text: t === "" ? "(blanks)" : t }), h("span", { "class": "faint", text: String(counts.get(t)) })]));
      });
    }
    search.addEventListener("input", paint); paint();
    function bt(label, fn, cls) { var b = h("button", { "class": "btn small " + (cls || ""), type: "button", text: label }); b.addEventListener("click", fn); return b; }
    var colRange = { r0: f.r0, c0: f.c0, r1: f.r1, c1: f.c1 };
    var el = h("div", { "class": "sh-filter", role: "dialog", "aria-label": "Filter column " + C.colName(c) }, [
      h("div", { "class": "row" }, [bt("Sort A → Z", function () { closeAll(); wb.sortRange(sh, colRange, [{ c: c, asc: true }], true); }), bt("Sort Z → A", function () { closeAll(); wb.sortRange(sh, colRange, [{ c: c, asc: false }], true); })]),
      h("div", { "class": "row" }, [op, val]), search,
      h("div", { "class": "row" }, [bt("Select all", function () { hidden.clear(); paint(); }, "ghost"), bt("Clear", function () { order.forEach(function (t) { hidden.add(t); }); paint(); }, "ghost")]),
      list,
      h("div", { "class": "row between" }, [bt("Remove filter", function () { closeAll(); wb.setFilter(sh, null); }, "ghost"), h("div", { "class": "row" }, [bt("Cancel", function () { closeAll(); }, "ghost"), bt("Apply", function () {
        closeAll();
        var crit = op.value ? { op: op.value, val: val.value } : hidden.size ? { hide: Array.from(hidden) } : null;
        wb.setFilterCol(sh, c, crit);
        S3.toast(sh.hidden.size ? sh.hidden.size + " row" + (sh.hidden.size === 1 ? "" : "s") + " hidden" : "Showing all rows");
      }, "primary")])])
    ]);
    var b = app.grid.cellBox(f.r0, c), rect = app.grid.scroller.getBoundingClientRect();
    popAt(rect.left + b.x + b.w - 260, rect.top + b.y + b.h + 2, el);
  }
  var CF_LABELS = [["gt", "Greater than"], ["lt", "Less than"], ["ge", "Greater than or equal"], ["le", "Less than or equal"], ["between", "Between"], ["eq", "Equal to"], ["ne", "Not equal to"],
    ["contains", "Text contains"], ["blank", "Is empty"], ["notblank", "Is not empty"], ["dup", "Duplicate values"], ["top", "Top N values"], ["formula", "Custom formula is TRUE"], ["scale", "Colour scale"], ["bar", "Data bars"]];
  function cfDialog() {
    var sh = sheet(), body = h("div", { "class": "stack" }), listEl = h("div", { "class": "stack sh-cflist" });
    var g = app.wb.clampArea(sh, S());
    var rangeI = inp(C.rangeStr(g), { "aria-label": "Apply to range" }), type = sel(CF_LABELS, "gt"), a = inp("", { placeholder: "Value" }), b = inp("", { placeholder: "and" });
    var bg = colorInp("#fde2e1"), fc = colorInp("#b3261e"), bold = h("input", { type: "checkbox" }), c1 = colorInp("#f8696b"), c2 = colorInp("#ffeb84"), c3 = colorInp("#63be7b"), three = h("input", { type: "checkbox" });
    three.checked = true;
    var styleRow = h("div", { "class": "row" }, [field("Fill", bg), field("Text", fc), h("label", { "class": "check" }, [bold, " Bold"])]);
    var scaleRow = h("div", { "class": "row" }, [field("Low", c1), field("Middle", c2), field("High", c3), h("label", { "class": "check" }, [three, " 3 colours"])]);
    var editing = null;
    function vis() {
      var t = type.value;
      a.parentNode.hidden = /^(blank|notblank|dup|scale|bar)$/.test(t);
      b.parentNode.hidden = t !== "between";
      a.placeholder = t === "formula" ? "=$B2>100" : t === "top" ? "10" : "Value";
      styleRow.hidden = t === "scale" || t === "bar"; scaleRow.hidden = t !== "scale" && t !== "bar";
      c2.parentNode.hidden = t === "bar" || !three.checked; c3.parentNode.hidden = t === "bar"; three.parentNode.hidden = t === "bar"; c1.parentNode.firstChild.textContent = t === "bar" ? "Bar colour" : "Low";
    }
    type.addEventListener("change", vis); three.addEventListener("change", vis);
    function paint() {
      listEl.textContent = "";
      if (!sh.cf.length) listEl.appendChild(h("p", { "class": "faint", text: "No rules on this sheet yet." }));
      sh.cf.forEach(function (rule) {
        var lab = (CF_LABELS.filter(function (x) { return x[0] === rule.type; })[0] || ["", rule.type])[1] + (rule.a !== "" && !/^(blank|notblank|dup|scale|bar)$/.test(rule.type) ? " " + rule.a + (rule.type === "between" ? " and " + rule.b : "") : "");
        var sw = h("span", { "class": "sh-cfsw" }); sw.style.background = rule.type === "scale" ? "linear-gradient(90deg," + (rule.colors || []).join(",") + ")" : rule.type === "bar" ? (rule.colors || ["#2f7fd8"])[0] : (rule.style.bg || "transparent"); sw.style.color = rule.style.fc || ""; sw.textContent = "Ab";
        var ed = h("button", { "class": "btn ghost small", type: "button", text: "Edit" }), del = h("button", { "class": "btn ghost small", type: "button", "aria-label": "Delete rule", html: '<svg class="i"><use href="#i-trash"/></svg>' });
        ed.addEventListener("click", function () { editing = rule.id; rangeI.value = C.rangeStr(rule.range); type.value = rule.type; a.value = rule.a; b.value = rule.b; if (rule.style.bg) bg.value = rule.style.bg; if (rule.style.fc) fc.value = rule.style.fc; bold.checked = !!rule.style.b; if (rule.colors) { c1.value = rule.colors[0]; if (rule.type === "scale") { three.checked = rule.colors.length === 3; if (rule.colors.length === 3) { c2.value = rule.colors[1]; c3.value = rule.colors[2]; } else c3.value = rule.colors[1]; } } vis(); addBtn.textContent = "Save rule"; });
        del.addEventListener("click", function () { app.wb.setMeta(sh, "Remove rule", function () { sh.cf = sh.cf.filter(function (x) { return x.id !== rule.id; }); }); paint(); });
        listEl.appendChild(h("div", { "class": "row sh-cfrow" }, [sw, h("div", { "class": "grow" }, [h("b", { text: lab }), h("div", { "class": "faint", text: C.rangeStr(rule.range) })]), ed, del]));
      });
    }
    var addBtn = null;
    function save() {
      var rg = C.parseRange(rangeI.value);
      if (!rg) { S3.toast("Type a range like B2:B20.", "error"); return false; }
      var t = type.value, rule = { id: editing || C.uid("cf"), type: t, range: { r0: rg.r0, c0: rg.c0, r1: Math.min(rg.r1, MAXR - 1), c1: Math.min(rg.c1, MAXC - 1) }, a: a.value, b: b.value, style: { bg: bg.value, fc: fc.value, b: bold.checked || undefined } };
      if (t === "scale") rule.colors = three.checked ? [c1.value, c2.value, c3.value] : [c1.value, c3.value];
      if (t === "bar") rule.colors = [c1.value];
      if (t === "formula" && !/^=/.test(a.value)) { S3.toast("A custom rule starts with =, e.g. =$C2>100", "error"); return false; }
      var clean = M.cleanRule(rule);
      if (!clean) return false;
      app.wb.setMeta(sh, "Conditional format", function () { var i = sh.cf.map(function (x) { return x.id; }).indexOf(clean.id); if (i >= 0) sh.cf[i] = clean; else sh.cf.push(clean); });
      editing = null; addBtn.textContent = "Add rule"; paint();
      return false;
    }
    body.appendChild(listEl);
    body.appendChild(h("h3", { "class": "sh-h3", text: "Rule" }));
    body.appendChild(h("div", { "class": "grid-2" }, [field("Apply to range", rangeI), field("Format cells if", type)]));
    body.appendChild(h("div", { "class": "grid-2" }, [field("Value", a), field("and", b)]));
    body.appendChild(styleRow); body.appendChild(scaleRow);
    vis(); paint();
    dialog("Conditional formatting", body, [{ label: "Done", ghost: true, act: function () { } }, { label: "Add rule", primary: true, keep: true, act: save, ref: function (bt) { addBtn = bt; } }], { wide: true });
  }
  function chartDialog(existing) {
    var sh = sheet(), g = dataRange(), ch = existing ? Object.assign({}, existing) : { id: C.uid("ch"), type: "column", range: C.quoteSheet(sh.name) + "!" + C.rangeStr(g), title: "", legend: true };
    var types = h("div", { "class": "seg sh-ctypes", role: "group", "aria-label": "Chart type" });
    SH.chart.TYPES.forEach(function (t) { var b = h("button", { type: "button", "aria-pressed": String(t[0] === ch.type), "data-v": t[0], text: t[1] }); b.addEventListener("click", function () { ch.type = t[0]; S3.$$("button", types).forEach(function (x) { x.setAttribute("aria-pressed", String(x === b)); }); prev(); }); types.appendChild(b); });
    var rangeI = inp(ch.range, { "aria-label": "Data range" }), title = inp(ch.title, { placeholder: "Chart title (optional)" }), leg = h("input", { type: "checkbox" });
    leg.checked = ch.legend !== false;
    var cv = h("canvas", { "class": "sh-cprev" }), note = h("p", { "class": "faint", text: "First row = series names, first column = labels (when it is text)." });
    function prev() {
      ch.range = rangeI.value; ch.title = title.value; ch.legend = leg.checked;
      var T = SH.GRID.THEMES[app.grid.dark ? "dark" : "paper"];
      var d = SH.chart.draw(cv, app.wb, ch, { bg: T.bg, fg: T.text, muted: T.muted, grid: T.chartGrid, axis: T.axis, font: SH.GRID.FONT }, Math.min(520, window.innerWidth - 80), 260);
      note.textContent = d && d.error ? d.error : "First row = series names, first column = labels (when it is text).";
    }
    [rangeI, title].forEach(function (el) { el.addEventListener("input", S3.debounce(prev, 150)); }); leg.addEventListener("change", prev);
    var body = h("div", { "class": "stack" }, [types, h("div", { "class": "grid-2" }, [field("Data range", rangeI), field("Title", title)]), h("label", { "class": "check" }, [leg, " Show legend"]), cv, note]);
    function save() {
      var d = SH.chart.data(app.wb, ch);
      if (d.error) { S3.toast(d.error, "error"); return false; }
      var c = M.cleanChart(Object.assign({}, ch, existing ? {} : placeFor(g)));
      app.wb.setMeta(sh, existing ? "Edit chart" : "Insert chart", function () { if (existing) sh.charts = sh.charts.map(function (x) { return x.id === c.id ? c : x; }); else sh.charts = sh.charts.concat([c]); });
      updateEmpty(); S3.done();
    }
    dialog(existing ? "Edit chart" : "Insert chart", body, [{ label: "Cancel", ghost: true, act: function () { } }, { label: existing ? "Save" : "Insert", primary: true, act: save }], { wide: true, enter: save });
    setTimeout(prev, 0);
  }
  function placeFor(g) { var gr = app.grid; gr.axes(); var c = Math.min(g.c1 + 1, sheet().ncols - 1); return { x: gr.colPos[c] + 16, y: gr.rowPos[Math.min(g.r0, sheet().nrows - 1)], w: 480, h: 300 }; }
  function fmtDialog() {
    var s = S(), v = app.wb.value(sheet(), s.r, s.c), code = inp(activeStyle().nf || "General", { "aria-label": "Format code", "class": "input mono" }), out = h("div", { "class": "out sh-fprev", "aria-live": "polite" });
    var list = h("div", { "class": "sh-fpresets" });
    F.PRESETS.forEach(function (p) { var b = h("button", { "class": "chip", type: "button", text: p.label }); b.addEventListener("click", function () { code.value = p.id; prev(); }); list.appendChild(b); });
    function prev() { var sample = typeof v === "number" ? v : 1234.5678, f = F.format(sample, code.value || "General"); out.textContent = f.t || " "; out.style.color = f.color || ""; }
    code.addEventListener("input", prev); prev();
    var body = h("div", { "class": "stack" }, [field("Format code", code), h("div", { "class": "faint", text: "0 = digit, # = optional digit, , = thousands, % = percent, \"text\" = literal, ; separates positive;negative;zero, yyyy mm dd h:mm for dates." }), out, list]);
    dialog("Number format", body, [{ label: "Cancel", ghost: true, act: function () { } }, { label: "Apply", primary: true, act: function () { A.numfmt(code.value.trim() || "General"); } }], { wide: true, enter: function () { A.numfmt(code.value.trim() || "General"); } });
  }
  function helpDialog(tab) {
    var q = inp("", { placeholder: "Search functions", "aria-label": "Search functions" }), list = h("div", { "class": "sh-fnlist" });
    function paint() {
      list.textContent = "";
      var t = q.value.toUpperCase();
      SH.fnNames().filter(function (n) { return !t || n.indexOf(t) >= 0 || (SH.FNDOC[n] || "").toUpperCase().indexOf(t) >= 0; }).forEach(function (n) {
        var d = (SH.FNDOC[n] || n + "()|").split("|"), b = h("button", { "class": "sh-fn", type: "button", title: "Insert into the active cell" }, [h("b", { "class": "mono", text: d[0] }), h("span", { text: d[1] })]);
        b.addEventListener("click", function () { dlg.close(); app.grid.beginEdit("edit", "=" + n + "("); });
        list.appendChild(b);
      });
    }
    q.addEventListener("input", paint); paint();
    var keys = [["Ctrl+C / X / V", "Copy, cut, paste (works with other spreadsheet apps)"], ["Ctrl+Z / Ctrl+Y", "Undo / redo"], ["F2 or double-click", "Edit a cell"], ["Enter / Tab", "Confirm and move down / right"], ["Alt+Enter", "New line inside a cell"],
      ["Ctrl+Enter", "Put the entry into every selected cell"], ["Esc", "Cancel editing"], ["Shift+arrows", "Extend the selection"], ["Ctrl+arrows", "Jump to the edge of the data"], ["Ctrl+A", "Select everything"], ["Ctrl+B / I / U", "Bold / italic / underline"],
      ["Ctrl+D / Ctrl+R", "Fill down / right"], ["Ctrl+;", "Today's date"], ["Alt+=", "AutoSum"], ["Ctrl+F / Ctrl+H", "Find / replace"], ["Ctrl+S", "Download .xlsx"], ["Ctrl+O", "Open a file"], ["Ctrl+P", "Print / PDF"],
      ["Typing =", "Start a formula; click cells to insert references, Tab picks a suggested function"]];
    var kt = h("table", { "class": "tbl" }); keys.forEach(function (k) { kt.appendChild(h("tr", null, [h("td", null, [h("kbd", { text: k[0] })]), h("td", { text: k[1] })])); });
    var pane1 = h("div", { "class": "stack" }, [q, list]), pane2 = h("div", { "class": "tbl-wrap" }, [kt]);
    var seg = h("div", { "class": "seg", role: "group", "aria-label": "Help" }), b1 = h("button", { type: "button", "aria-pressed": "true", text: "Functions (" + SH.fnNames().length + ")" }), b2 = h("button", { type: "button", "aria-pressed": "false", text: "Keyboard" });
    function show(i) { b1.setAttribute("aria-pressed", String(!i)); b2.setAttribute("aria-pressed", String(!!i)); pane1.hidden = !!i; pane2.hidden = !i; }
    b1.addEventListener("click", function () { show(0); }); b2.addEventListener("click", function () { show(1); });
    seg.appendChild(b1); seg.appendChild(b2);
    var dlg = dialog("Help", h("div", { "class": "stack" }, [seg, pane1, pane2]), null, { wide: true });
    show(tab === "keys" ? 1 : 0);
  }
  function settingsDialog() {
    var order = sel([["DMY", "Day/Month/Year (27/10/2026)"], ["MDY", "Month/Day/Year (10/27/2026)"]], F.settings.dateOrder), guardCb = h("input", { type: "checkbox" });
    guardCb.checked = S3.store("sheets.csvGuard") !== false;
    var body = h("div", { "class": "stack" }, [field("When I type a date like 03/04/2026", order), h("label", { "class": "check" }, [guardCb, " Protect CSV downloads: prefix cells that start with = + - @ so other apps don't run them as formulas"])]);
    dialog("Settings", body, [{ label: "Save", primary: true, act: function () { F.settings.dateOrder = order.value; S3.store("sheets.dateOrder", order.value); S3.store("sheets.csvGuard", guardCb.checked); } }]);
  }

  // ---------- menus ----------
  function fileMenu() {
    return [
      { label: "New workbook", act: function () { load(M.blank()); } },
      { label: "Open .xlsx / .csv…", kbd: "Ctrl+O", act: function () { $("#fileIn").click(); } },
      { label: "My workbooks…", act: workbooks },
      "-",
      { label: "Download .xlsx", kbd: "Ctrl+S", act: saveX },
      { label: "Download .csv (this sheet)", act: function () { saveCsv(","); } },
      { label: "Download .tsv (this sheet)", act: function () { saveCsv("\t"); } },
      "-",
      { label: "Print / PDF — whole sheet", kbd: "Ctrl+P", act: function () { printNow(false); } },
      { label: "Print / PDF — selection", act: function () { printNow(true); } },
      "-",
      { label: "Open the sample budget", act: function () { load(SH.sample()); } },
      { label: "Settings…", act: settingsDialog }
    ];
  }
  function editMenu() {
    return [
      { label: "Undo", kbd: "Ctrl+Z", disabled: !app.wb.canUndo(), act: A.undo }, { label: "Redo", kbd: "Ctrl+Y", disabled: !app.wb.canRedo(), act: A.redo }, "-",
      { label: "Cut", kbd: "Ctrl+X", act: A.cut }, { label: "Copy", kbd: "Ctrl+C", act: A.copy }, { label: "Paste", kbd: "Ctrl+V", act: A.paste },
      { label: "Paste values only", act: A.pasteValues }, { label: "Paste formatting only", act: A.pasteFormats }, "-",
      { label: "Fill down", kbd: "Ctrl+D", act: A.fillDown }, { label: "Fill right", kbd: "Ctrl+R", act: A.fillRight }, "-",
      { label: "Find…", kbd: "Ctrl+F", act: function () { findDialog(false); } }, { label: "Find and replace…", kbd: "Ctrl+H", act: function () { findDialog(true); } }, "-",
      { label: "Clear contents", kbd: "Del", act: function () { A.clear("contents"); } }, { label: "Clear formatting", act: function () { A.clear("formats"); } }, { label: "Clear all", act: function () { A.clear("all"); } },
      { label: "Select all", kbd: "Ctrl+A", act: function () { app.grid.selectAll(); } }
    ];
  }
  function insertMenu() {
    return [
      { label: "Row above", act: function () { A.insertRows(false); } }, { label: "Row below", act: function () { A.insertRows(true); } },
      { label: "Column left", act: function () { A.insertCols(false); } }, { label: "Column right", act: function () { A.insertCols(true); } }, "-",
      { label: "Delete rows", act: A.deleteRows }, { label: "Delete columns", act: A.deleteCols }, "-",
      { label: "Sheet", act: function () { app.wb.newSheet(); afterSheetChange(); } }, { label: "Chart…", act: function () { chartDialog(); } },
      { label: "Function…", act: function () { helpDialog("fn"); } }, { label: "Today's date", kbd: "Ctrl+;", act: A.today }
    ];
  }
  function formatMenu() {
    var st = activeStyle();
    return [
      { label: "Bold", kbd: "Ctrl+B", on: st.b, act: A.bold }, { label: "Italic", kbd: "Ctrl+I", on: st.i, act: A.italic }, { label: "Underline", kbd: "Ctrl+U", on: st.u, act: A.underline }, { label: "Strikethrough", on: st.st, act: A.strike }, "-",
      { label: "Number format…", act: fmtDialog }, { label: "Conditional formatting…", act: cfDialog }, "-",
      { label: "Wrap text", on: st.wr, act: A.wrap }, { label: "Merge / unmerge", act: A.merge }, { label: "Fit column width", act: function () { var s = S(); for (var c = s.c0; c <= Math.min(s.c1, s.c0 + 50); c++) app.grid.autofit(c); } }, "-",
      { label: "Clear formatting", act: function () { A.clear("formats"); } }
    ];
  }
  function dataMenu() {
    var sh = sheet();
    return [
      { label: "Sort A → Z", act: function () { A.sort(true); } }, { label: "Sort Z → A", act: function () { A.sort(false); } }, { label: "Sort range…", act: sortDialog }, "-",
      { label: sh.filter ? "Remove filter" : "Create a filter", act: A.filter },
      { label: "Show all rows (clear filter values)", disabled: !sh.filter, act: function () { app.wb.setMeta(sh, "Clear filter", function () { sh.filter.cols = {}; }); } }, "-",
      { label: "Remove duplicate rows", act: removeDupes }
    ];
  }
  function removeDupes() {
    var sh = sheet(), g = dataRange(), wb = app.wb, header = looksHeader(sh, g), seen = new Set(), keep = [], drop = 0;
    for (var r = g.r0 + (header ? 1 : 0); r <= g.r1; r++) { var k = []; for (var c = g.c0; c <= g.c1; c++) k.push(wb.text(sh, r, c).toLowerCase()); var key = k.join("\u0001"); if (seen.has(key)) { drop++; continue; } seen.add(key); keep.push(r); }
    if (!drop) { S3.toast("No duplicate rows in " + C.rangeStr(g)); return; }
    if (!window.confirm("Remove " + drop + " duplicate row(s) from " + C.rangeStr(g) + "?")) return;
    var block = keep.map(function (r) { var row = []; for (var c = g.c0; c <= g.c1; c++) row.push(wb.cell(sh, r, c) || null); return row; });
    wb.run("Remove duplicates", function () {
      var start = g.r0 + (header ? 1 : 0);
      for (var r = start; r <= g.r1; r++) for (var c = g.c0; c <= g.c1; c++) { var x = block[r - start] ? block[r - start][c - g.c0] : null; wb.set(sh, r, c, x && x.f !== undefined ? M.mk(undefined, EN.moveFormula(x.f, r - keep[r - start], 0), x.s) : x); }
    });
    S3.toast("Removed " + drop + " duplicate row(s)");
  }
  function viewMenu() {
    var sh = sheet(), s = S();
    return [
      { label: "No frozen rows or columns", on: !sh.freeze.r && !sh.freeze.c, act: function () { A.freeze(0, 0); } },
      { label: "Freeze top row", on: sh.freeze.r === 1 && !sh.freeze.c, act: function () { A.freeze(1, sh.freeze.c); } },
      { label: "Freeze first column", on: sh.freeze.c === 1, act: function () { A.freeze(sh.freeze.r, 1); } },
      { label: "Freeze rows above " + (s.r + 1) + " and columns left of " + C.colName(s.c), act: function () { A.freeze(s.r, s.c); } }, "-",
      { label: "Gridlines", on: sh.grid !== false, act: A.gridlines },
      { label: "Dark grid", on: app.grid.dark, act: A.darkGrid },
      { label: "Right-to-left sheet", on: !!sh.rtl, act: A.rtl }
    ];
  }

  // ---------- toolbar ----------
  function wireToolbar() {
    var g = app.grid;
    function on(id, fn) { $(id).addEventListener("click", function () { fn(); g.focus(); }); }
    on("#btnUndo", A.undo); on("#btnRedo", A.redo);
    on("#tbBold", A.bold); on("#tbItalic", A.italic); on("#tbUnder", A.underline); on("#tbWrap", A.wrap); on("#tbMerge", A.merge);
    on("#tbLeft", function () { A.align("left"); }); on("#tbCenter", function () { A.align("center"); }); on("#tbRight", function () { A.align("right"); });
    on("#tbDecUp", function () { A.decimals(1); }); on("#tbDecDown", function () { A.decimals(-1); });
    on("#tbCur", function () { A.numfmt(S3.store("sheets.cur") || '"JOD "#,##0.000'); }); on("#tbPct", function () { A.numfmt("0%"); });
    on("#tbSortA", function () { A.sort(true); }); on("#tbSortZ", function () { A.sort(false); }); on("#tbFilter", A.filter);
    on("#tbChart", function () { chartDialog(); }); on("#tbCF", cfDialog); on("#tbFind", function () { findDialog(true); }); on("#tbRtl", A.rtl);
    on("#btnSave", saveX);
    var nf = $("#tbFmt");
    F.PRESETS.forEach(function (p) { nf.appendChild(h("option", { value: p.id, text: p.label })); });
    nf.appendChild(h("option", { value: "custom", text: "Custom…" }));
    nf.addEventListener("change", function () { if (nf.value === "custom") fmtDialog(); else { A.numfmt(nf.value); if (/JOD|\$|€|£|SAR|AED/.test(nf.value)) S3.store("sheets.cur", nf.value); } g.focus(); });
    var size = $("#tbSize");
    [9, 10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 32, 40, 48].forEach(function (n) { size.appendChild(h("option", { value: String(n), text: String(n) })); });
    size.addEventListener("change", function () { fmt({ fs: +size.value === 13 ? null : +size.value }); g.focus(); });
    // colour pickers open their own popover instead of a menu
    ["#tbFc", "#tbBg"].forEach(function (id, i) {
      var b = $(id);
      b.addEventListener("click", function (e) {
        e.stopPropagation();
        if (b.getAttribute("aria-expanded") === "true") { closeAll(); return; }
        var key = i ? "bg" : "fc", pal = ["#000000", "#434343", "#666666", "#999999", "#cccccc", "#ffffff", "#c0392b", "#e67e22", "#f1c40f", "#27ae60", "#16a085", "#2980b9", "#8e44ad", "#d81b60", "#fde2e1", "#fff2cc", "#e2f0d9", "#ddebf7", "#ede7f6", "#e0f7fa"];
        var el = h("div", { "class": "sh-palette", role: "dialog", "aria-label": i ? "Fill colour" : "Text colour" }), gridEl = h("div", { "class": "sh-swatches" });
        pal.forEach(function (c) { var sw = h("button", { type: "button", "class": "sh-sw", title: c, "aria-label": c }); sw.style.background = c; sw.addEventListener("click", function () { closeAll(); var p = {}; p[key] = c; fmt(p); g.focus(); }); gridEl.appendChild(sw); });
        var custom = colorInp(activeStyle()[key] || (i ? "#fff2cc" : "#c0392b"));
        custom.addEventListener("change", function () { closeAll(); var p = {}; p[key] = custom.value; fmt(p); g.focus(); });
        var reset = h("button", { "class": "btn ghost small", type: "button", text: i ? "No fill" : "Automatic" });
        reset.addEventListener("click", function () { closeAll(); var p = {}; p[key] = null; fmt(p); g.focus(); });
        el.appendChild(gridEl); el.appendChild(h("div", { "class": "row between" }, [reset, h("label", { "class": "row sh-custom" }, [h("span", { text: "Custom" }), custom])]));
        var r = b.getBoundingClientRect(); popAt(r.left, r.bottom + 4, el, b);
      });
    });
    menuFor($("#tbBorders"), function () {
      function bd(mode) { return function () { app.wb.borders(sheet(), selRange(), mode); }; }
      return [{ label: "All borders", act: bd("all") }, { label: "Outside border", act: bd("outer") }, { label: "Bottom border", act: bd("bottom") }, { label: "Thick bottom border", act: bd("thickbottom") },
        { label: "Top border", act: bd("top") }, { label: "Left border", act: bd("left") }, { label: "Right border", act: bd("right") }, "-", { label: "No borders", act: bd("none") }];
    });
    menuFor($("#tbValign"), function () { var va = activeStyle().va || "bottom"; return [{ label: "Top", on: va === "top", act: function () { A.valign("top"); } }, { label: "Middle", on: va === "middle", act: function () { A.valign("middle"); } }, { label: "Bottom", on: va === "bottom", act: function () { A.valign("bottom"); } }]; });
    menuFor($("#tbSum"), function () { return ["SUM", "AVERAGE", "COUNT", "MAX", "MIN"].map(function (f) { return { label: f === "SUM" ? "Sum" : f[0] + f.slice(1).toLowerCase(), kbd: f === "SUM" ? "Alt+=" : "", act: function () { A.autosum(f); } }; }); });
    menuFor($("#tbFreeze"), viewMenu);
    menuFor($("#tbRows"), insertMenu);
    menuFor($("#mFile"), fileMenu); menuFor($("#mEdit"), editMenu); menuFor($("#mInsert"), insertMenu); menuFor($("#mFormat"), formatMenu); menuFor($("#mData"), dataMenu); menuFor($("#mView"), viewMenu);
    $("#mHelp").addEventListener("click", function () { helpDialog("fn"); });
    $("#btnAddSheet").addEventListener("click", function () { app.wb.newSheet(); afterSheetChange(); });
    $("#btnSheets").addEventListener("click", function (e) { e.stopPropagation(); var r = e.currentTarget.getBoundingClientRect(); menu(app.wb.sheets.map(function (sh, i) { return { label: sh.name, on: i === app.wb.active, act: function () { switchSheet(i); } }; }), r.left, r.top - 4, e.currentTarget); });
    var name = $("#docName");
    name.addEventListener("change", function () { app.wb.name = name.value.trim().slice(0, 120) || "Untitled spreadsheet"; queueSave(); });
    name.addEventListener("keydown", function (e) { if (e.key === "Enter") { name.blur(); g.focus(); } });
    var fin = $("#fileIn");
    fin.addEventListener("change", function () { openFiles(Array.from(fin.files || [])); fin.value = ""; });
    $("#emptyOpen").addEventListener("click", function () { fin.click(); });
    $("#emptySample").addEventListener("click", function () { load(SH.sample()); });
    var appEl = $("#app");
    appEl.addEventListener("dragover", function (e) { if (e.dataTransfer && Array.prototype.indexOf.call(e.dataTransfer.types || [], "Files") >= 0) { e.preventDefault(); appEl.classList.add("drag"); } });
    appEl.addEventListener("dragleave", function (e) { if (e.target === appEl || !appEl.contains(e.relatedTarget)) appEl.classList.remove("drag"); });
    appEl.addEventListener("drop", function (e) { appEl.classList.remove("drag"); if (e.dataTransfer && e.dataTransfer.files.length) { e.preventDefault(); openFiles(Array.from(e.dataTransfer.files)); } });
  }
  function contextMenu(x, y, hit) {
    var items = [{ label: "Cut", kbd: "Ctrl+X", act: A.cut }, { label: "Copy", kbd: "Ctrl+C", act: A.copy }, { label: "Paste", kbd: "Ctrl+V", act: A.paste }, { label: "Paste values only", act: A.pasteValues }, "-"];
    if (hit.zone === "colhdr") items = items.concat([{ label: "Insert column left", act: function () { A.insertCols(false); } }, { label: "Insert column right", act: function () { A.insertCols(true); } }, { label: "Delete column(s)", act: A.deleteCols }, { label: "Fit width", act: function () { app.grid.autofit(hit.c); } }, { label: "Sort A → Z", act: function () { A.sort(true); } }, { label: "Sort Z → A", act: function () { A.sort(false); } }]);
    else if (hit.zone === "rowhdr") items = items.concat([{ label: "Insert row above", act: function () { A.insertRows(false); } }, { label: "Insert row below", act: function () { A.insertRows(true); } }, { label: "Delete row(s)", act: A.deleteRows }]);
    else items = items.concat([{ label: "Insert row above", act: function () { A.insertRows(false); } }, { label: "Insert column left", act: function () { A.insertCols(false); } }, { label: "Delete row(s)", act: A.deleteRows }, { label: "Delete column(s)", act: A.deleteCols }, "-",
      { label: "Clear contents", kbd: "Del", act: function () { A.clear("contents"); } }, { label: "Sort A → Z", act: function () { A.sort(true); } }, { label: "Conditional formatting…", act: cfDialog }, { label: "Insert chart…", act: function () { chartDialog(); } }]);
    menu(items, x, y);
  }
  function shortcuts(e) {
    var ctrl = e.ctrlKey || e.metaKey, k = e.key.toLowerCase();
    if (e.altKey && (k === "=" || e.key === "=")) { A.autosum("SUM"); return true; }
    if (!ctrl) return false;
    if (k === "z" && !e.shiftKey) { A.undo(); return true; }
    if (k === "y" || (k === "z" && e.shiftKey)) { A.redo(); return true; }
    if (k === "b") { A.bold(); return true; } if (k === "i") { A.italic(); return true; } if (k === "u") { A.underline(); return true; }
    if (k === "a") { app.grid.selectAll(); return true; }
    if (k === "f") { findDialog(false); return true; } if (k === "h") { findDialog(true); return true; }
    if (k === "s") { saveX(); return true; } if (k === "o") { $("#fileIn").click(); return true; } if (k === "p") { printNow(!(S().r0 === S().r1 && S().c0 === S().c1)); return true; }
    if (k === "d") { A.fillDown(); return true; } if (k === "r") { A.fillRight(); return true; }
    if (k === ";") { A.today(); return true; }
    return false;
  }
  // global keys (when focus is outside the grid, e.g. on a toolbar button)
  document.addEventListener("keydown", function (e) {
    if (!app.grid || app.grid.ownsFocus() || document.querySelector("dialog[open]")) return;
    var t = e.target; if (t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.tagName === "SELECT")) return;
    var ctrl = e.ctrlKey || e.metaKey;
    if (ctrl && /^[zysop]$/i.test(e.key) && shortcuts(e)) e.preventDefault();
  });

  // ---------- boot ----------
  app.boot = function () {
    F.settings.dateOrder = S3.store("sheets.dateOrder") || (/^en-US/i.test(navigator.language || "") ? "MDY" : "DMY");
    var start = M.blank();
    app.wb = start;
    app.grid = new SH.Grid($("#gridwrap"), start, {
      onSel: syncBar, onCommit: syncBar,
      onEdit: function (t) { var fx = $("#fx"); if (document.activeElement !== fx) fx.value = t; },
      onContext: contextMenu, onFilter: filterPop, onKey: shortcuts,
      onChartEdit: function (ch) { chartDialog(ch); },
      onChartDelete: function (ch) { var sh = sheet(); app.wb.setMeta(sh, "Delete chart", function () { sh.charts = sh.charts.filter(function (x) { return x.id !== ch.id; }); }); app.grid.focus(); }
    });
    app.grid.dark = S3.store("sheets.darkGrid") === true;
    wireBar(); wireToolbar();
    load(start);
    var last = S3.store("sheets.last");
    if (last) IO.store.get(last).then(function (rec) { if (rec && rec.data && app.wb === start && !start.sheet().cells.size) { try { load(M.fromJSON(rec.data)); setChip("Restored from this device", "ok"); } catch (e) { /* start fresh */ } } }).catch(function () { /* no storage */ });
    app.grid.focus();
  };
  app.dialogs = { find: findDialog, sort: sortDialog, cf: cfDialog, chart: chartDialog, fmt: fmtDialog, help: helpDialog, books: workbooks };
  app.openFiles = openFiles; app.saveCsv = saveCsv; app.saveX = saveX; app.switchSheet = switchSheet;
})();
