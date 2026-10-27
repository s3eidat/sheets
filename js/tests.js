/* Sheets self-test, part 1: the formula engine, formats and model against known answers. Pure logic. */
(function () {
  "use strict";
  var SH = window.SH;

  function engineTests() {
    var C = SH.core, F = SH.fmt, EN = SH.engine, M = SH.model, fails = [];
    function ok(c, m) { if (!c) fails.push(m); }
    function show(v) { return v instanceof C.Err ? v.e : JSON.stringify(v); }
    var wb = M.blank("t"), T = wb.sheets[0];
    wb.renameSheet(T, "T");
    var S2 = wb.newSheet("My Sheet");
    function put(sh, a, text) { var p = C.parseAddr(a); wb.run("t", function () { wb.setInput(sh, p.r, p.c, text); }); }
    function val(sh, a) { var p = C.parseAddr(a); return wb.value(sh, p.r, p.c); }
    [1, 2, 3, 4, 5].forEach(function (v, i) { put(T, "A" + (i + 1), String(v)); });
    ["apple", "banana", "cherry", "apple", "date"].forEach(function (v, i) { put(T, "B" + (i + 1), v); });
    [10, 20, 30, 40, 50].forEach(function (v, i) { put(T, "C" + (i + 1), String(v)); });
    function f(formula, expect, tol) {
      var got = wb.eng.evalFormula(formula, T, 20, 20);
      var pass = expect instanceof C.Err ? got instanceof C.Err && got.e === expect.e
        : typeof expect === "number" ? typeof got === "number" && Math.abs(got - expect) <= (tol || 1e-9) : got === expect;
      ok(pass, "=" + formula + " → " + show(got) + " (want " + show(expect) + ")");
    }
    var E = C.E;
    // precedence & operators
    f("1+2*3", 7); f("(1+2)*3", 9); f("-2^2", 4); f("2^3^2", 64); f("2*3%", 0.06); f("1&2+3", "15"); f("1+2=3", true);
    f("10/4", 2.5); f("5-3-1", 1); f("\"a\"&\"b\"", "ab"); f("1/0", E["#DIV/0!"]); f("\"x\"+1", E["#VALUE!"]); f("FOO(1)", E["#NAME?"]);
    f("SUM(1,2", E["#ERROR!"]); f("\"a\"=\"A\"", true); f("0.1+0.2=0.3", true); f("\"3\"*2", 6); f("TRUE+1", 2); f("INDEX({1,2;3,4}*2,2,1)", 6);
    f("SUM({1,2;3,4}*2)", 20); f("3>2>1", true); f("\"b\">\"a\"", true); f("1<\"a\"", true);
    // aggregates & criteria
    f("SUM(A1:A5)", 15); f("AVERAGE(A1:A5)", 3); f("COUNT(A1:B5)", 5); f("COUNTA(A1:B5)", 10); f("COUNTBLANK(A1:A10)", 5);
    f("COUNTIF(B1:B5,\"apple\")", 2); f("COUNTIF(A1:A5,\">2\")", 3); f("COUNTIF(B1:B5,\"*an*\")", 1); f("COUNTIFS(B1:B5,\"apple\",A1:A5,\">1\")", 1);
    f("SUMIF(B1:B5,\"apple\",C1:C5)", 50); f("SUMIF(A1:A5,\">=4\")", 9); f("SUMIFS(C1:C5,A1:A5,\">=2\",A1:A5,\"<=4\")", 90);
    f("AVERAGEIF(B1:B5,\"a*\",C1:C5)", 25); f("AVERAGEIFS(C1:C5,A1:A5,\">3\")", 45); f("MAXIFS(C1:C5,B1:B5,\"apple\")", 40); f("MINIFS(C1:C5,A1:A5,\">1\")", 20);
    f("MIN(A1:A5)", 1); f("MAX(A1:A5,7)", 7); f("PRODUCT(A1:A5)", 120); f("SUMPRODUCT(A1:A5,C1:C5)", 550); f("MEDIAN(1,3,2,4)", 2.5);
    f("ROUND(STDEV(A1:A5),6)", 1.581139); f("LARGE(A1:A5,2)", 4); f("SMALL(A1:A5,2)", 2); f("RANK(4,A1:A5)", 2); f("MODE(1,2,2,3)", 2);
    f("SUM(A1:A5,\"2\",TRUE)", 18); f("SUBTOTAL(9,C1:C5)", 150);
    // rounding & math
    f("ROUND(2.5,0)", 3); f("ROUND(-2.5,0)", -3); f("ROUND(1.005,2)", 1.01); f("ROUNDUP(3.14159,2)", 3.15); f("ROUNDDOWN(-3.789,1)", -3.7);
    f("ROUND(1234.567,-2)", 1200); f("MROUND(10,3)", 9); f("CEILING(4.3,0.5)", 4.5); f("FLOOR(4.7,2)", 4); f("TRUNC(-4.9)", -4);
    f("ABS(-3)", 3); f("POWER(2,10)", 1024); f("SQRT(16)", 4); f("SQRT(-1)", E["#NUM!"]); f("MOD(-3,2)", 1); f("MOD(3,-2)", -1); f("MOD(5,0)", E["#DIV/0!"]);
    f("INT(-2.5)", -3); f("RANDBETWEEN(1,1)", 1); f("AND(RAND()>=0,RAND()<1)", true); f("LOG(8,2)", 3); f("LN(EXP(2))", 2); f("FACT(5)", 120); f("PI()", Math.PI);
    // logic & info
    f("IF(A1>0,\"pos\",\"neg\")", "pos"); f("IF(FALSE,1)", false); f("IFS(A1>3,\"big\",A1>0,\"small\")", "small"); f("IFS(A1>9,1)", E["#N/A"]);
    f("AND(TRUE,1)", true); f("OR(FALSE,0)", false); f("XOR(TRUE,TRUE)", false); f("NOT(FALSE)", true); f("IFERROR(1/0,\"x\")", "x"); f("IFERROR(5,\"x\")", 5);
    f("IFNA(NA(),1)", 1); f("ISNA(NA())", true); f("ISERROR(1/0)", true); f("ISBLANK(Z99)", true); f("ISNUMBER(A1)", true); f("ISTEXT(B1)", true);
    f("CHOOSE(2,\"a\",\"b\")", "b"); f("SWITCH(2,1,\"x\",2,\"y\",\"z\")", "y"); f("SWITCH(9,1,\"x\",\"z\")", "z");
    // lookup
    f("VLOOKUP(\"cherry\",B1:C5,2,FALSE)", 30); f("VLOOKUP(\"CHERRY\",B1:C5,2,0)", 30); f("VLOOKUP(3.5,A1:C5,3,TRUE)", 30); f("VLOOKUP(3.5,A1:C5,3)", 30);
    f("VLOOKUP(0,A1:C5,3,TRUE)", E["#N/A"]); f("VLOOKUP(\"zzz\",B1:C5,2,FALSE)", E["#N/A"]); f("VLOOKUP(9,A1:C5,3)", 50); f("VLOOKUP(\"ch*\",B1:C5,2,FALSE)", 30);
    f("VLOOKUP(1,A1:C5,4,FALSE)", E["#REF!"]); f("HLOOKUP(\"b\",{\"a\",\"b\";1,2},2,FALSE)", 2);
    f("XLOOKUP(\"date\",B1:B5,C1:C5)", 50); f("XLOOKUP(\"x\",B1:B5,C1:C5,\"none\")", "none"); f("XLOOKUP(3.5,A1:A5,C1:C5,,-1)", 30); f("XLOOKUP(3.5,A1:A5,C1:C5,,1)", 40);
    f("XLOOKUP(\"apple\",B1:B5,C1:C5,,0,-1)", 40); f("INDEX(C1:C5,4)", 40); f("INDEX(A1:C5,2,3)", 20); f("INDEX(A1:C1,3)", 10); f("MATCH(\"cherry\",B1:B5,0)", 3); f("MATCH(3.7,A1:A5,1)", 3);
    f("MATCH(\"x\",B1:B5,0)", E["#N/A"]); f("XMATCH(4,A1:A5)", 4); f("ROWS(A1:C5)", 5); f("COLUMNS(A1:C5)", 3); f("ROW(C4)", 4); f("COLUMN(C4)", 3);
    f("SUM(OFFSET(A1,1,0,2,1))", 5); f("INDIRECT(\"C2\")", 20); f("LOOKUP(4.5,A1:A5,C1:C5)", 40);
    // text
    f("CONCAT(\"a\",1,TRUE)", "a1TRUE"); f("CONCAT(A1:A3)", "123"); f("TEXTJOIN(\", \",TRUE,\"a\",\"\",\"b\")", "a, b"); f("LEFT(\"Hello\",2)", "He"); f("RIGHT(\"Hello\",3)", "llo");
    f("MID(\"Hello\",2,3)", "ell"); f("LEN(\"Hello\")", 5); f("TRIM(\"  a   b \")", "a b"); f("UPPER(\"abc\")", "ABC"); f("LOWER(\"ABC\")", "abc");
    f("PROPER(\"hello wORLD\")", "Hello World"); f("SUBSTITUTE(\"a-b-c\",\"-\",\"+\")", "a+b+c"); f("SUBSTITUTE(\"a-b-c\",\"-\",\"+\",2)", "a-b+c");
    f("FIND(\"l\",\"Hello\")", 3); f("SEARCH(\"L*O\",\"Hello\")", 3); f("FIND(\"z\",\"Hello\")", E["#VALUE!"]); f("REPLACE(\"abcdef\",2,3,\"X\")", "aXef");
    f("TEXT(0.256,\"0.0%\")", "25.6%"); f("TEXT(1234.5,\"#,##0.00\")", "1,234.50"); f("TEXT(DATE(2024,3,5),\"yyyy-mm-dd\")", "2024-03-05"); f("TEXT(DATE(2024,3,5),\"dddd\")", "Tuesday");
    f("VALUE(\"1,234.5\")", 1234.5); f("VALUE(\"abc\")", E["#VALUE!"]); f("REPT(\"ab\",3)", "ababab"); f("EXACT(\"a\",\"A\")", false); f("TEXTBEFORE(\"a-b-c\",\"-\")", "a"); f("TEXTAFTER(\"a-b-c\",\"-\",2)", "c");
    f("1&\"\"", "1"); f("(0.1+0.2)&\"\"", "0.3"); f("INDEX(LEFT(B1:B2,1),2)", "b");
    // dates (1900 system incl. the fictional 1900-02-29 = serial 60)
    f("DATE(1900,1,1)", 1); f("DATE(1900,2,28)", 59); f("DATE(1900,2,29)", 60); f("DATE(1900,3,1)", 61); f("DATE(2000,1,1)", 36526); f("DATE(2024,1,1)", 45292);
    f("DATE(2024,14,1)", 45689); f("DATE(2024,3,0)", 45351); f("YEAR(45292)", 2024); f("MONTH(60)", 2); f("DAY(60)", 29); f("MONTH(61)", 3); f("DAY(59)", 28);
    f("WEEKDAY(DATE(2024,1,1))", 2); f("WEEKDAY(1)", 1); f("WEEKDAY(DATE(2024,1,1),2)", 1); f("EDATE(DATE(2024,1,31),1)", 45351); f("EOMONTH(DATE(2024,1,15),1)", 45351);
    f("EOMONTH(DATE(2023,12,15),-1)", 45260); f("NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))", 23); f("NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),DATE(2024,1,2))", 22);
    f("WORKDAY(DATE(2024,1,5),1)", 45299); f("DATEDIF(DATE(2020,1,15),DATE(2024,3,10),\"Y\")", 4); f("DATEDIF(DATE(2020,1,15),DATE(2024,3,10),\"M\")", 49);
    f("DATEDIF(DATE(2020,1,15),DATE(2024,3,10),\"D\")=DATE(2024,3,10)-DATE(2020,1,15)", true); f("DATEDIF(DATE(2020,1,15),DATE(2024,3,10),\"MD\")", 24);
    f("DATEDIF(DATE(2020,1,15),DATE(2024,3,10),\"YM\")", 1); f("DATEDIF(DATE(2024,3,1),DATE(2020,1,1),\"Y\")", E["#NUM!"]); f("TODAY()=INT(NOW())", true);
    f("DAYS(DATE(2024,3,1),DATE(2024,2,1))", 29); f("HOUR(0.75)", 18); f("MINUTE(TIME(10,35,0))", 35); f("DATEVALUE(\"2024-03-15\")", 45366); f("\"2024-01-02\"-\"2024-01-01\"", 1);
    // finance
    f("ROUND(PMT(5%/12,360,200000),2)", -1073.64); f("ROUND(FV(0.06/12,10,-200,-500,1),2)", 2581.4); f("ROUND(PV(0.08/12,12*20,500,,0),2)", -59777.15);
    f("ROUND(NPER(0.12/12,-100,-1000,10000,1),4)", 59.6739); f("ROUND(NPV(0.1,-10000,3000,4200,6800),2)", 1188.44); f("ROUND(RATE(48,-200,8000),6)", 0.007701);
    f("ROUND(IRR({-70000,12000,15000,18000,21000,26000}),4)", 0.0866); f("ROUND(IPMT(0.1/12,1,36,8000),2)", -66.67); f("ROUND(PPMT(0.1/12,1,24,2000),2)", -75.62);
    // arrays
    f("ROWS(UNIQUE(B1:B5))", 4); f("INDEX(SORT(A1:A5,1,-1),1)", 5); f("ROWS(FILTER(A1:A5,A1:A5>2))", 3); f("SUM(FILTER(C1:C5,B1:B5=\"apple\"))", 50);
    f("FILTER(A1:A5,A1:A5>9)", E["#CALC!"]); f("FILTER(A1:A5,A1:A5>9,\"-\")", "-"); f("SUM(SEQUENCE(4))", 10); f("SUM(TRANSPOSE(A1:A3))", 6); f("SUM(A1:A5*C1:C5)", 550);
    f("SUM(IF(A1:A5>2,C1:C5,0))", 120);

    // references, recalculation, $refs, cross-sheet
    put(T, "E1", "=A1+A2"); ok(val(T, "E1") === 3, "E1=A1+A2");
    put(T, "A1", "10"); ok(val(T, "E1") === 12, "recalc after edit: " + show(val(T, "E1")));
    put(S2, "A1", "=T!E1*2"); ok(val(S2, "A1") === 24, "cross-sheet " + show(val(S2, "A1")));
    put(T, "E2", "='My Sheet'!A1+1"); ok(val(T, "E2") === 25, "quoted sheet ref " + show(val(T, "E2")));
    put(T, "A2", "3"); ok(val(T, "E2") === 27, "transitive recalc across sheets " + show(val(T, "E2")));
    ok(EN.moveFormula("A1+$A$1+A$1+$A1", 1, 1) === "B2+$A$1+B$1+$A2", "move $refs: " + EN.moveFormula("A1+$A$1+A$1+$A1", 1, 1));
    ok(EN.moveFormula("SUM(A1:B2)*'My Sheet'!C3", 2, 0) === "SUM(A3:B4)*'My Sheet'!C5", "move range + sheet ref");
    ok(EN.moveFormula("A1", -1, 0) === "#REF!", "move off the grid → #REF!");
    ok(EN.normalize("sum(a1:b2") === "SUM(A1:B2)", "normalize: " + EN.normalize("sum(a1:b2"));
    ok(EN.toFile("XLOOKUP(1,A:A,B:B)+SUM(1)") === "_xlfn.XLOOKUP(1,A:A,B:B)+SUM(1)" && EN.fromFile("_xlfn.XLOOKUP(1,A:A,B:B)") === "XLOOKUP(1,A:A,B:B)", "_xlfn prefix");
    // long chains: downward and upward (deferred evaluation must not overflow the stack)
    wb.run("chain", function () { for (var i = 0; i < 3000; i++) wb.setInput(T, 100 + i, 6, i ? "=G" + (100 + i) + "+1" : "1"); });
    ok(val(T, "G3100") === 3000, "downward chain " + show(val(T, "G3100")));
    wb.run("chain", function () { for (var i = 2999; i >= 0; i--) wb.setInput(T, 100 + i, 7, i === 2999 ? "1" : "=H" + (102 + i) + "+1"); });
    ok(val(T, "H101") === 3000, "upward chain " + show(val(T, "H101")));
    put(T, "G101", "5"); ok(val(T, "G3100") === 3004, "chain recalc " + show(val(T, "G3100")));
    // cycles
    put(T, "J1", "=J2"); put(T, "J2", "=J1"); ok(show(val(T, "J1")) === "#CIRC!" && show(val(T, "J2")) === "#CIRC!", "two-cell cycle");
    put(T, "J3", "=J3+1"); ok(show(val(T, "J3")) === "#CIRC!", "self reference");
    put(T, "J4", "=J1*2"); ok(show(val(T, "J4")) === "#CIRC!", "error flows from a cycle");
    put(T, "J2", "5"); ok(val(T, "J1") === 5 && val(T, "J4") === 10, "cycle broken → recalculates " + show(val(T, "J1")));
    // spills
    put(T, "L1", "=SEQUENCE(3)"); ok(val(T, "L3") === 3, "spill " + show(val(T, "L3")));
    put(T, "M1", "=SUM(L1:L3)"); ok(val(T, "M1") === 6, "sum of a spill " + show(val(T, "M1")));
    put(T, "L2", "x"); ok(show(val(T, "L1")) === "#SPILL!" && show(val(T, "M1")) === "#SPILL!", "blocked spill " + show(val(T, "L1")) + " " + show(val(T, "M1")));
    put(T, "L2", ""); ok(val(T, "L2") === 2 && val(T, "M1") === 6, "unblocked spill " + show(val(T, "M1")));
    put(T, "L1", "=SEQUENCE(4)"); ok(val(T, "M1") === 6 && val(T, "L4") === 4, "spill grows");
    put(T, "N1", "=L4*10"); put(T, "L1", "=SEQUENCE(2)"); ok(val(T, "N1") === 0, "dependent of a shrinking spill " + show(val(T, "N1")));

    // formats
    function fm(v, code, want) { var t = F.format(v, code).t; ok(t === want, "format " + v + " " + code + " → " + t + " (want " + want + ")"); }
    fm(1234.5, "#,##0.00", "1,234.50"); fm(-1234.5, "#,##0.00", "-1,234.50"); fm(0.125, "0.0%", "12.5%"); fm(5.25, '"JOD "#,##0.000', "JOD 5.250");
    fm(-5, "$#,##0.00", "-$5.00"); fm(45292, "dd/mm/yyyy", "01/01/2024"); fm(45292.5, "yyyy-mm-dd h:mm", "2024-01-01 12:00"); fm(60, "yyyy-mm-dd", "1900-02-29");
    fm(0.75, "h:mm AM/PM", "6:00 PM"); fm(1 / 3, "General", "0.3333333333"); fm(1e12, "General", "1E+12"); fm(12345.678, "0.00E+00", "1.23E+04");
    fm(1500000, "#,##0,,\"M\"", "2M"); fm(0, "0.00;-0.00;\"zero\"", "zero"); fm(45366, "d mmm yyyy", "15 Mar 2024"); fm(3.5, "0", "4"); fm(0.5, "#.00", ".50");
    ok(F.format(-2, "0;[Red]-0").color === "#d00000", "red negative colour");
    ok(F.bumpDecimals("0.00", 1) === "0.000" && F.bumpDecimals("0.00", -1) === "0.0" && F.bumpDecimals("#,##0", 1) === "#,##0.0", "decimal buttons");
    function pin(text, v, nf) { var p = F.parseInput(text); ok(p.v === v && (nf === undefined || p.nf === nf), "parse " + text + " → " + JSON.stringify(p)); }
    pin("12%", 0.12, "0%"); pin("1,234", 1234, "#,##0"); pin("$5.5", 5.5); pin("2024-03-15", 45366, "yyyy-mm-dd"); pin("'=1", "=1"); pin("TRUE", true);
    pin("JOD 5.250", 5.25, '"JOD "#,##0.000'); pin("(12)", -12); pin("hello", "hello"); pin("1e3", 1000); pin("14:30", 14.5 / 24, "h:mm");
    F.settings.dateOrder = "DMY"; pin("15/03/2024", 45366); pin("03/04/2024", 45385); F.settings.dateOrder = "MDY"; pin("03/04/2024", 45355); F.settings.dateOrder = "DMY";

    // structural edits rewrite formulas
    var wb2 = M.blank("s"), A = wb2.sheets[0];
    function put2(a, t) { var p = C.parseAddr(a); wb2.run("t", function () { wb2.setInput(A, p.r, p.c, t); }); }
    function v2(a) { var p = C.parseAddr(a); return wb2.value(A, p.r, p.c); }
    function f2(a) { var p = C.parseAddr(a), x = wb2.cell(A, p.r, p.c); return x && x.f; }
    put2("A1", "1"); put2("A2", "2"); put2("A3", "=SUM(A1:A2)"); put2("B1", "=A2*2");
    wb2.insertDelete(A, "r", 1, 1); ok(f2("A4") === "SUM(A1:A3)" && v2("A4") === 3 && f2("B1") === "A3*2", "insert row: " + f2("A4") + " " + f2("B1"));
    wb2.undo(); ok(f2("A3") === "SUM(A1:A2)" && v2("A3") === 3, "undo insert");
    wb2.insertDelete(A, "r", 1, -1); ok(f2("A2") === "SUM(A1:A1)" && v2("A2") === 1 && show(v2("B1")) === "#REF!", "delete row: " + f2("A2") + " " + f2("B1"));
    wb2.undo();
    wb2.insertDelete(A, "c", 0, 1); ok(f2("C1") === "B2*2" && f2("B3") === "SUM(B1:B2)", "insert column " + f2("C1"));
    wb2.undo(); ok(f2("B1") === "A2*2", "undo insert column");
    put2("A1", "9"); ok(v2("A3") === 11, "set"); wb2.undo(); ok(v2("A3") === 3, "undo set " + v2("A3")); wb2.redo(); ok(v2("A3") === 11, "redo set");
    // sort & fill
    var wb3 = M.blank("o"), B = wb3.sheets[0];
    function put3(a, t) { var p = C.parseAddr(a); wb3.run("t", function () { wb3.setInput(B, p.r, p.c, t); }); }
    function v3(a) { var p = C.parseAddr(a); return wb3.value(B, p.r, p.c); }
    put3("A1", "3"); put3("A2", "1"); put3("A3", "2"); put3("B1", "=A1*10"); put3("B2", "=A2*10"); put3("B3", "=A3*10"); put3("A4", "");
    wb3.sortRange(B, { r0: 0, c0: 0, r1: 2, c1: 1 }, [{ c: 0, asc: true }], false);
    ok(v3("A1") === 1 && v3("A3") === 3 && v3("B1") === 10 && v3("B3") === 30, "sort moves rows with formulas: " + [v3("A1"), v3("B1"), v3("B3")].join(","));
    put3("D1", "1"); put3("D2", "2"); wb3.fill(B, { r0: 0, c0: 3, r1: 1, c1: 3 }, { r0: 0, c0: 3, r1: 4, c1: 3 }); ok(v3("D5") === 5, "fill linear series " + v3("D5"));
    put3("E1", "Item 1"); wb3.fill(B, { r0: 0, c0: 4, r1: 0, c1: 4 }, { r0: 0, c0: 4, r1: 2, c1: 4 }); ok(v3("E3") === "Item 3", "fill text+number " + v3("E3"));
    put3("F1", "Jan"); wb3.fill(B, { r0: 0, c0: 5, r1: 0, c1: 5 }, { r0: 0, c0: 5, r1: 0, c1: 7 }); ok(v3("H1") === "Mar", "fill months across " + v3("H1"));
    put3("G2", "=D2*2"); wb3.fill(B, { r0: 1, c0: 6, r1: 1, c1: 6 }, { r0: 1, c0: 6, r1: 3, c1: 6 }); ok(v3("G4") === 8, "fill formula " + v3("G4"));
    put3("I1", "7"); wb3.fill(B, { r0: 0, c0: 8, r1: 0, c1: 8 }, { r0: 0, c0: 8, r1: 2, c1: 8 }); ok(v3("I3") === 7, "fill single number copies");
    // filter + SUBTOTAL ignores hidden rows
    var wb4 = M.blank("f"), Q = wb4.sheets[0];
    wb4.run("t", function () { ["Item", "a", "b", "a"].forEach(function (t, i) { wb4.setInput(Q, i, 0, t); }); ["N", "1", "2", "3"].forEach(function (t, i) { wb4.setInput(Q, i, 1, t); }); wb4.setInput(Q, 5, 1, "=SUBTOTAL(9,B2:B4)"); });
    wb4.setFilter(Q, { r0: 0, c0: 0, r1: 3, c1: 1 }); wb4.setFilterCol(Q, 0, { hide: ["b"] });
    ok(Q.hidden.has(2) && !Q.hidden.has(1) && Q.hidden.size === 1, "filter hides rows");
    // copy/paste moves relative refs
    var blk = wb3.copyBlock(B, { r0: 0, c0: 1, r1: 0, c1: 1 });
    wb3.pasteBlock(B, 5, 3, blk, "all"); ok(wb3.cell(B, 5, 3).f === "C6*10", "paste shifts refs " + wb3.cell(B, 5, 3).f);
    // serialisation round trip
    var back = M.fromJSON(JSON.parse(JSON.stringify(wb.toJSON())));
    ok(back.value(back.sheetByName("T"), 0, 4) === 13 && back.value(back.sheetByName("My Sheet"), 0, 0) === 26, "JSON round trip " + back.value(back.sheetByName("T"), 0, 4));
    return fails;
  }
  SH.tests = { engine: engineTests };
})();
