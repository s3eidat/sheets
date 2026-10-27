/* Sheets — sample workbook, self-test and boot. The app itself lives in js/. */
(function () {
  "use strict";
  var SH = window.SH, C = SH.core, M = SH.model, IO = SH.io, F = SH.fmt;

  // ---------- sample: a household budget with formulas, cross-sheet SUMIF, formats, rules and a chart ----------
  SH.sample = function () {
    var wb = M.blank("Household budget"), B = wb.sheets[0], self = wb;
    B.name = "Budget";
    var T = wb.addSheetRaw("Transactions");
    var JOD = '"JOD "#,##0.000', hdr = { b: true, bg: "#e7f3f2", fc: "#0a5653", bd: { b: { s: "medium", c: "#0e9f9a" } } };
    function put(sh, a, text, s) { var p = C.parseAddr(a); self.setInput(sh, p.r, p.c, text); if (s) { var x = self.cell(sh, p.r, p.c) || {}; self.set(sh, p.r, p.c, M.withStyle(x, M.mergeStyle(x.s, s))); } }
    var tx = [["2026-10-01", "Rent", "Rent", 450], ["2026-10-02", "Supermarket", "Groceries", 96.5], ["2026-10-04", "Electricity bill", "Utilities", 54.2], ["2026-10-05", "Fuel", "Transport", 40],
      ["2026-10-07", "School books", "School", 150], ["2026-10-09", "Pharmacy", "Health", 22.5], ["2026-10-10", "Supermarket", "Groceries", 120.25], ["2026-10-11", "Restaurant", "Eating out", 38.6],
      ["2026-10-13", "Water bill", "Utilities", 18.7], ["2026-10-15", "Taxi", "Transport", 12.2], ["2026-10-16", "Savings transfer", "Savings", 300], ["2026-10-18", "Bakery and vegetables", "Groceries", 45],
      ["2026-10-20", "Fuel", "Transport", 42], ["2026-10-22", "Café", "Eating out", 21.4], ["2026-10-24", "Supermarket", "Groceries", 85], ["2026-10-25", "Internet", "Utilities", 15.5],
      ["2026-10-27", "Pizza night", "Eating out", 52.3], ["2026-10-28", "Bus card", "Transport", 40]];
    var cats = [["Rent", 450], ["Groceries", 320], ["Utilities", 95], ["Transport", 120], ["School", 150], ["Health", 60], ["Eating out", 80], ["Savings", 300]];
    wb.run("Sample", function () {
      put(B, "A1", "Household budget — October 2026", { b: true, fs: 20, fc: "#0a5653" });
      put(B, "A2", "All amounts in Jordanian dinars. Actuals add up the Transactions sheet.", { i: true, fc: "#6b7480" });
      ["Category", "Budget", "Actual", "Difference", "% used"].forEach(function (t, i) { put(B, C.colName(i) + "3", t, M.mergeStyle(hdr, { ha: i ? "right" : null })); });
      cats.forEach(function (c, i) {
        var r = i + 4;
        put(B, "A" + r, c[0]);
        put(B, "B" + r, String(c[1]), { nf: JOD });
        put(B, "C" + r, "=SUMIF(Transactions!$C$2:$C$40,A" + r + ",Transactions!$D$2:$D$40)", { nf: JOD });
        put(B, "D" + r, "=B" + r + "-C" + r, { nf: '#,##0.000;[Red]-#,##0.000' });
        put(B, "E" + r, "=IFERROR(C" + r + "/B" + r + ",0)", { nf: "0%" });
      });
      var tot = { b: true, bd: { t: { s: "thin", c: "#0e9f9a" } } };
      put(B, "A12", "Total", tot); put(B, "B12", "=SUM(B4:B11)", M.mergeStyle(tot, { nf: JOD })); put(B, "C12", "=SUM(C4:C11)", M.mergeStyle(tot, { nf: JOD }));
      put(B, "D12", "=B12-C12", M.mergeStyle(tot, { nf: '#,##0.000;[Red]-#,##0.000' })); put(B, "E12", "=C12/B12", M.mergeStyle(tot, { nf: "0%" }));
      put(B, "A14", "Monthly income", { b: true }); put(B, "B14", "1850", { nf: JOD });
      put(B, "A15", "Spent so far"); put(B, "B15", "=C12", { nf: JOD });
      put(B, "A16", "Left over", { b: true }); put(B, "B16", "=B14-B15", { nf: JOD, b: true, fc: "#1e7d4f" });
      put(B, "A17", "Status"); put(B, "B17", '=IF(B16>=0,"On track ✓","Over budget")', { b: true });
      put(B, "A18", "Biggest overspend"); put(B, "B18", '=INDEX(A4:A11,MATCH(MIN(D4:D11),D4:D11,0))');
      put(B, "A19", "Days left in month"); put(B, "B19", "=DATEDIF(DATE(2026,10,27),EOMONTH(DATE(2026,10,27),0),\"D\")");
      ["Date", "Description", "Category", "Amount"].forEach(function (t, i) { put(T, C.colName(i) + "1", t, M.mergeStyle(hdr, { ha: i === 3 ? "right" : null })); });
      tx.forEach(function (t, i) { var r = i + 2; put(T, "A" + r, t[0], { nf: "d mmm yyyy" }); put(T, "B" + r, t[1]); put(T, "C" + r, t[2]); put(T, "D" + r, String(t[3]), { nf: JOD }); });
      put(T, "C21", "Total", { b: true }); put(T, "D21", "=SUBTOTAL(9,D2:D19)", { b: true, nf: JOD });
    });
    B.colW.set(0, 170); [1, 2, 3].forEach(function (c) { B.colW.set(c, 118); }); B.colW.set(4, 80);
    B.rowH.set(0, 36); B.freeze = { r: 3, c: 0 };
    B.merges.push({ r0: 0, c0: 0, r1: 0, c1: 4 });
    B.cf = [M.cleanRule({ id: "cf1", type: "lt", range: { r0: 3, c0: 3, r1: 10, c1: 3 }, a: "0", style: { bg: "#fde2e1", fc: "#b3261e", b: true } }),
      M.cleanRule({ id: "cf2", type: "bar", range: { r0: 3, c0: 4, r1: 10, c1: 4 }, colors: ["#0e9f9a"] })];
    B.charts = [M.cleanChart({ id: "ch1", type: "column", range: "Budget!A3:C11", title: "Budget vs actual (JOD)", x: 618, y: 30, w: 470, h: 300 })];
    T.colW.set(0, 110); T.colW.set(1, 190); T.colW.set(2, 120); T.colW.set(3, 120); T.freeze = { r: 1, c: 0 };
    T.filter = { r0: 0, c0: 0, r1: 18, c1: 3, cols: {} };
    wb.undoStack = []; wb.redoStack = [];
    wb.eng.rebuild(); wb.refreshFilters();
    return wb;
  };

  window.S3_DEMO = async function () {
    SH.app.load(SH.sample());
    SH.app.grid.select(5, 2, 5, 2);
    await new Promise(function (r) { setTimeout(r, 200); });
  };

  // ---------- self-test ----------
  function near(a, b, e) { return typeof a === "number" && Math.abs(a - b) <= (e || 1e-9); }
  window.S3_SELFTEST = async function () {
    var fails = SH.tests.engine(), t0;
    function ok(c, m) { if (!c) fails.push(m); }

    // sample workbook: cross-sheet SUMIF + totals
    var s = SH.sample(), sb = s.sheetByName("Budget");
    ok(near(s.value(sb, 11, 2), 1604.15, 1e-6), "sample total actual " + s.value(sb, 11, 2));
    ok(s.value(sb, 16, 1) === "On track ✓" && s.value(sb, 17, 1) === "Eating out", "sample status " + s.value(sb, 16, 1) + " / " + s.value(sb, 17, 1));

    // CSV: formula-injection guard on export, formulas stay text on import, caps
    var wb = M.blank("csv"), sh = wb.sheets[0];
    wb.run("t", function () {
      wb.set(sh, 0, 0, { v: "=HYPERLINK(\"x.invalid\",\"y\")", q: 1 }); wb.set(sh, 0, 1, { v: "+1+2" }); wb.set(sh, 0, 2, { v: -5 }); wb.set(sh, 0, 3, { v: "@SUM(A1)" });
      wb.set(sh, 0, 4, { v: "a,\"b\"" }); wb.set(sh, 0, 5, { v: 45292, s: { nf: "yyyy-mm-dd" } }); wb.set(sh, 0, 6, { f: "1/3" }); wb.set(sh, 0, 7, { v: "عمر" });
    });
    var csv = IO.writeCsv(wb, sh, null, { guard: true, bom: false });
    var want = "\"'=HYPERLINK(\"\"x.invalid\"\",\"\"y\"\")\",'+1+2,-5,'@SUM(A1),\"a,\"\"b\"\"\",2024-01-01,0.333333333333333,عمر\r\n";
    ok(csv === want, "csv guard: " + csv);
    var noGuard = IO.writeCsv(wb, sh, null, { guard: false, bom: false });
    ok(noGuard.indexOf("+1+2") === 0 || noGuard.indexOf("\"=HYPERLINK") === 0, "csv without guard keeps text");
    var rc = IO.readCsv("name,amount\n\"=WEBSERVICE(\"\"evil.invalid\"\")\",5\n\"multi\nline\",\"1,234.5\"\n\"unterminated", "t.csv");
    var rs = rc.wb.sheets[0];
    ok(rs.cells.get(1 * C.KW + 0).v === "=WEBSERVICE(\"evil.invalid\")" && rs.cells.get(1 * C.KW + 0).f === undefined, "csv import keeps formulas as text");
    ok(rc.wb.value(rs, 2, 0) === "multi\nline" && rc.wb.value(rs, 2, 1) === 1234.5 && rc.wb.value(rs, 3, 0) === "unterminated", "csv quoted fields / unterminated quote");
    var wide = IO.parseDelimited("a\tb\tc\n1\t2\t3", "\t", { rows: 1, cols: 2, text: 10 });
    ok(wide.rows.length === 1 && wide.cut, "csv caps rows");
    ok(IO.csvField("-12.5", ",", true) === "-12.5" && IO.csvField("-x", ",", true) === "'-x", "negative numbers are not guarded");

    // .xlsx: zip directory guard against bombs and non-xlsx files
    function throws(fn, rx) { try { fn(); return false; } catch (e) { return rx.test(e.message); } }
    ok(throws(function () { IO.zipCheck(new TextEncoder().encode("hello, not a zip at all").buffer); }, /not a valid/), "non-zip rejected");
    var z = new Uint8Array(await S3.zip([{ name: "xl/workbook.xml", data: "<x/>" }, { name: "xl/worksheets/sheet1.xml", data: "<y/>" }]).arrayBuffer());
    ok(!!IO.zipCheck(z.buffer.slice(0)), "small zip passes");
    var bomb = z.slice(0), dv = new DataView(bomb.buffer);
    for (var i = 0; i < bomb.length - 4; i++) if (dv.getUint32(i, true) === 0x02014b50) { dv.setUint32(i + 24, 900 * 1024 * 1024, true); break; }
    ok(throws(function () { IO.zipCheck(bomb.buffer); }, /zip bomb|expands/), "zip bomb rejected");
    var mac = new Uint8Array(await S3.zip([{ name: "xl/workbook.xml", data: "<x/>" }, { name: "xl/vbaProject.bin", data: "MZ" }]).arrayBuffer());
    ok(IO.zipCheck(mac.buffer).macros === true, "macro part detected (and ignored)");
    ok(throws(function () { IO.zipCheck(new Uint8Array(emptyZip()).buffer); }, /not an Excel/), "zip without a workbook rejected");

    // .xlsx round trip: values, formulas, formats, merges, freeze, RTL, widths, sheets, charts + rules
    var x = M.blank("Round trip"), a = x.sheets[0], b2;
    a.name = "Data";
    b2 = x.addSheetRaw("عربي");
    x.run("t", function () {
      x.setInput(a, 0, 0, "Item"); x.setInput(a, 0, 1, "Qty"); x.setInput(a, 0, 2, "Price");
      x.setInput(a, 1, 0, "Tea"); x.setInput(a, 1, 1, "3"); x.setInput(a, 1, 2, "1.250");
      x.setInput(a, 2, 0, "Coffee"); x.setInput(a, 2, 1, "2"); x.setInput(a, 2, 2, "4.5");
      x.setInput(a, 3, 0, "Total"); x.setInput(a, 3, 2, "=SUMPRODUCT(B2:B3,C2:C3)");
      x.setInput(a, 4, 0, "=XLOOKUP(\"Coffee\",A2:A3,C2:C3)"); x.setInput(a, 5, 0, "2024-03-15"); x.setInput(a, 6, 0, "TRUE"); x.setInput(a, 7, 0, "=1/0");
      x.set(a, 0, 0, M.withStyle(x.cell(a, 0, 0), { b: true, i: true, bg: "#fff2cc", fc: "#c0392b", ha: "center", fs: 16, bd: { b: { s: "thin", c: "#000000" } } }));
      x.set(a, 3, 2, M.withStyle(x.cell(a, 3, 2), { nf: '"JOD "#,##0.000', wr: true, va: "top" }));
      x.setInput(b2, 0, 0, "مرحبا"); x.setInput(b2, 1, 0, "=Data!C4*2");
    });
    a.merges.push({ r0: 9, c0: 0, r1: 9, c1: 2 }); a.freeze = { r: 1, c: 1 }; a.colW.set(0, 150); a.rowH.set(0, 40); b2.rtl = true;
    a.cf = [M.cleanRule({ id: "r1", type: "gt", range: { r0: 1, c0: 1, r1: 2, c1: 1 }, a: "2", style: { bg: "#e2f0d9" } })];
    a.charts = [M.cleanChart({ id: "c1", type: "pie", range: "Data!A1:B3", title: "Qty", x: 300, y: 10, w: 300, h: 200 })];
    x.eng.rebuild();
    var blob = await IO.writeXlsx(x), back = (await IO.readXlsx(await blob.arrayBuffer(), "rt.xlsx")).wb, A2 = back.sheetByName("Data"), B3 = back.sheetByName("عربي");
    ok(A2 && B3 && back.sheets.length === 2, "xlsx sheets " + back.sheets.map(function (q) { return q.name; }).join(","));
    if (A2 && B3) {
      ok(back.cell(A2, 3, 2).f === "SUMPRODUCT(B2:B3,C2:C3)" && near(back.value(A2, 3, 2), 12.75), "xlsx formula " + (back.cell(A2, 3, 2) || {}).f + " = " + back.value(A2, 3, 2));
      ok(back.cell(A2, 4, 0).f === "XLOOKUP(\"Coffee\",A2:A3,C2:C3)" && back.value(A2, 4, 0) === 4.5, "xlsx _xlfn round trip " + (back.cell(A2, 4, 0) || {}).f);
      ok(back.value(A2, 1, 0) === "Tea" && back.value(A2, 2, 2) === 4.5 && back.value(A2, 6, 0) === true, "xlsx values");
      ok(back.value(A2, 5, 0) === 45366 && F.isDateFmt((back.style(A2, 5, 0) || {}).nf), "xlsx date " + back.value(A2, 5, 0));
      ok(back.value(A2, 7, 0) instanceof C.Err && back.value(A2, 7, 0).e === "#DIV/0!", "xlsx error value");
      var st = back.style(A2, 0, 0) || {};
      ok(st.b && st.i && st.bg === "#fff2cc" && st.fc === "#c0392b" && st.ha === "center" && st.fs === 16 && st.bd && st.bd.b && st.bd.b.c === "#000000", "xlsx styles " + JSON.stringify(st));
      var st2 = back.style(A2, 3, 2) || {};
      ok(st2.nf === '"JOD "#,##0.000' && st2.wr && st2.va === "top", "xlsx number format " + JSON.stringify(st2));
      ok(A2.merges.length === 1 && A2.merges[0].c1 === 2 && A2.freeze.r === 1 && A2.freeze.c === 1, "xlsx merge/freeze " + JSON.stringify(A2.merges) + JSON.stringify(A2.freeze));
      ok(Math.abs(A2.colW.get(0) - 150) <= 2 && Math.abs(A2.rowH.get(0) - 40) <= 2, "xlsx sizes " + A2.colW.get(0) + "/" + A2.rowH.get(0));
      ok(B3.rtl && back.value(B3, 0, 0) === "مرحبا" && near(back.value(B3, 1, 0), 25.5), "xlsx RTL + cross-sheet " + back.value(B3, 1, 0));
      ok(A2.charts.length === 1 && A2.charts[0].type === "pie" && A2.cf.length === 1 && A2.cf[0].type === "gt", "xlsx charts/rules via hidden meta sheet");
      ok(!back.sheetByName("_s3meta"), "meta sheet stays hidden");
    }
    // a hostile .xlsx: web/DDE-style formulas are only text to our engine
    var hx = new ExcelJS.Workbook(), hs = hx.addWorksheet("S");
    hs.getCell("A1").value = { formula: "WEBSERVICE(\"example.invalid/steal\")" };
    hs.getCell("A2").value = { formula: "cmd|' /C calc'!A0" };
    hs.getCell("A3").value = { formula: "HYPERLINK(\"example.invalid\",\"click\")" };
    hs.getCell("A4").value = { richText: [{ text: "<img src=x onerror=alert(1)>" }] };
    var hostile = (await IO.readXlsx(await hx.xlsx.writeBuffer(), "h.xlsx")).wb, H1 = hostile.sheets[0];
    ok(hostile.value(H1, 0, 0) instanceof C.Err && hostile.value(H1, 0, 0).e === "#NAME?", "WEBSERVICE is unknown (#NAME?)");
    ok(hostile.value(H1, 1, 0) instanceof C.Err, "DDE-style formula is an error, never run");
    ok(hostile.value(H1, 2, 0) instanceof C.Err, "HYPERLINK not executed");
    ok(hostile.value(H1, 3, 0) === "<img src=x onerror=alert(1)>", "rich text kept as plain text");

    // performance: 10,000 rows with formulas, recalculation and scrolling within budget
    var big = M.blank("perf"), P = big.sheets[0];
    t0 = performance.now();
    big.run("fill", function () {
      for (var r = 0; r < 10000; r++) {
        big.set(P, r, 0, { v: "Row " + (r + 1) }); big.set(P, r, 1, { v: r % 97 }); big.set(P, r, 2, { v: (r * 7) % 13 + 0.5 });
        big.set(P, r, 3, { f: "B" + (r + 1) + "*C" + (r + 1) }); big.set(P, r, 4, { f: r ? "E" + r + "+D" + (r + 1) : "D1" });
      }
      big.set(P, 0, 6, { f: "SUM(D:D)" }); big.set(P, 1, 6, { f: "COUNTIF(B:B,\">50\")" });
    });
    var tCalc = performance.now() - t0;
    ok(near(big.value(P, 9999, 4), big.value(P, 0, 6), 1e-6), "10k running total = SUM(D:D)");
    t0 = performance.now();
    big.run("edit", function () { big.setInput(P, 0, 1, "1000"); });
    var tEdit = performance.now() - t0;
    ok(near(big.value(P, 9999, 4), big.value(P, 0, 6), 1e-6), "10k recalc after edit");
    var box = S3.h("div", { "class": "sh-gridwrap sh-perf" });
    document.body.appendChild(box);
    var g = new SH.Grid(box, big, {}), times = [];
    g.drawNow();
    for (var k = 0; k < 40; k++) { var st0 = performance.now(); g.scroller.scrollTop = k * 5900; g.drawNow(); times.push(performance.now() - st0); }
    times.sort(function (p, q) { return p - q; });
    var median = times[Math.floor(times.length / 2)], p90 = times[Math.floor(times.length * 0.9)];
    box.remove();
    ok(median < 16 && p90 < 40, "10k-row scroll frame median " + median.toFixed(1) + " ms, p90 " + p90.toFixed(1) + " ms (budget 16/40)");
    ok(tCalc < 4000 && tEdit < 1000, "10k build " + tCalc.toFixed(0) + " ms, edit+recalc " + tEdit.toFixed(0) + " ms");
    SH.perf = { build: Math.round(tCalc), edit: Math.round(tEdit), frameMedian: +median.toFixed(2), frameP90: +p90.toFixed(2) };
    return fails.length ? fails.slice(0, 10).join(" | ") : true;
  };
  function emptyZip() { return new Uint8Array([0x50, 0x4b, 0x05, 0x06, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]).buffer; }

  SH.app.boot();
})();
