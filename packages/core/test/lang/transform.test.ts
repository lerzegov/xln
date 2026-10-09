import { describe, expect, it } from "vitest";
import {
  FormulaError,
  builtinCollision,
  compile,
  compileWithDiagnostics,
  decompile,
  decompileWithDiagnostics,
} from "../../src/index.js";

const codes = (r: { diagnostics: { code: string }[] }) => r.diagnostics.map((d) => d.code);

describe("decompile: stored → display", () => {
  it("strips _xlfn., _xlfn._xlws., _xlpm., _xleta.", () => {
    expect(decompile("_xlfn.LAMBDA(_xlpm.b,_xlpm.g, _xlpm.b*(1+_xlpm.g))")).toBe("LAMBDA(b,g, b*(1+g))");
    expect(decompile("_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A1:A9,B1:B9>0))")).toBe("SORT(FILTER(A1:A9,B1:B9>0))");
    expect(decompile("_xlfn.GROUPBY(NameList[Kind], NameList[Name], _xleta.COUNTA, , 0)")).toBe("GROUPBY(NameList[Kind], NameList[Name], COUNTA, , 0)");
    expect(decompile("_xlfn.LET(_xlpm.f, _xlfn.LAMBDA(_xlpm.t, _xlpm.t*2), _xlpm.f(3))")).toBe("LET(f, LAMBDA(t, t*2), f(3))");
  });

  it("ANCHORARRAY(x) → x#, keeping complex arguments as calls", () => {
    expect(decompile("_xlfn.ANCHORARRAY('S1'!$B$1)")).toBe("'S1'!$B$1#");
    expect(decompile("ROWS(_xlfn.ANCHORARRAY(E1))*Rate")).toBe("ROWS(E1#)*Rate");
    expect(decompile("SUM(_xlfn.ANCHORARRAY(Revenue))")).toBe("SUM(Revenue#)");
    expect(decompile("_xlfn.ANCHORARRAY(INDEX(A:A,1))")).toBe("ANCHORARRAY(INDEX(A:A,1))");
  });

  it("SINGLE(x) → @x where that means the same", () => {
    expect(decompile("_xlfn.SINGLE(A1:A10)")).toBe("@A1:A10");
    expect(decompile("_xlfn.SINGLE(A1:A10)*2")).toBe("@A1:A10*2");
    expect(decompile("_xlfn.SINGLE(1+2)")).toBe("SINGLE(1+2)");
    // `@x:B5` would read as @(x:B5): keep the call.
    expect(decompile("_xlfn.SINGLE(A1):B5")).toBe("SINGLE(A1):B5");
  });

  it("[#This Row] → @ in structured references", () => {
    expect(decompile("Tbl[[#This Row],[Col]]")).toBe("Tbl[@Col]");
    expect(decompile("Tbl[[#This Row],[Col A]]")).toBe("Tbl[@[Col A]]");
    expect(decompile("Tbl[[#This Row],[A]:[B]]")).toBe("Tbl[@[A]:[B]]");
    expect(decompile("Tbl[#This Row]")).toBe("Tbl[@]");
    expect(decompile("Tbl[[#Headers],[Col]]")).toBe("Tbl[[#Headers],[Col]]");
  });

  it("un-qualifies the home sheet of a sheet-scoped name", () => {
    const ctx = { homeSheet: "BS", localNames: ["FixedAssetRoll_base"] };
    expect(decompile("_xlfn.ANCHORARRAY('BS'!$C$6)", ctx)).toBe("$C$6#");
    expect(decompile("_xlfn.CHOOSEROWS('BS'!FixedAssetRoll_base, 1)", ctx)).toBe("CHOOSEROWS(FixedAssetRoll_base, 1)");
    // Other sheets, and names not scoped to the home sheet, stay qualified, quoted as Excel
    // writes them (spec §14 issue 7): column-like names bare, cell-like or spaced ones quoted.
    expect(decompile("IS!Sales+'BS'!Other+'Check'!$B$36", ctx)).toBe("IS!Sales+BS!Other+Check!$B$36");
    expect(decompile("S1!A1+'SCF recursive'!B2+'S1:S3'!C3+[1]Ext!D4", ctx)).toBe("'S1'!A1+'SCF recursive'!B2+'S1:S3'!C3+[1]Ext!D4");
    expect(decompile("'bs'!A1", ctx)).toBe("A1"); // sheet names are case-insensitive
  });

  it("keeps the home sheet on a deleted reference: a bare #REF! is the error constant (FEEDBACK 2026-10-08)", () => {
    const ctx = { homeSheet: "IS" };
    // Each stored form, decompiled and compiled again, is what Excel stored.
    const cases: [string, string][] = [
      ["_xlfn.ANCHORARRAY(IS!#REF!)", "IS!#REF!#"],
      ["_xlfn.SINGLE(IS!#REF!)", "@IS!#REF!"],
      ["SUM(IS!#REF!)", "SUM(IS!#REF!)"],
      ["IS!#REF!+1", "IS!#REF!+1"],
      ["_xlfn.LET(_xlpm.x,IS!#REF!,_xlpm.x+1)", "LET(x,IS!#REF!,x+1)"],
      ["_xlfn.XLOOKUP(1,IS!#REF!,IS!$A$1:$A$3)", "XLOOKUP(1,IS!#REF!,$A$1:$A$3)"],
      ["_xlfn.TAKE(_xlfn.ANCHORARRAY(IS!#REF!),1)", "TAKE(IS!#REF!#,1)"],
      ["_xlfn.ANCHORARRAY(IS!#REF!)+_xlfn.ANCHORARRAY(IS!$B$2)", "IS!#REF!#+$B$2#"],
      ["_xlfn.ANCHORARRAY('My Sheet'!#REF!)", "'My Sheet'!#REF!#"],
      ["#REF!", "#REF!"],
    ];
    for (const [stored, display] of cases) {
      expect(decompile(stored, ctx)).toBe(display);
      expect(compile(display, ctx)).toBe(stored);
    }
  });

  it("keeps the author's spacing and line breaks", () => {
    expect(decompile("_xlfn.LAMBDA(_xlpm.x,\r\n  _xlpm.x*2)")).toBe("LAMBDA(x,\r\n  x*2)");
    expect(decompile("_xlfn.LET( _xlpm.base, Assumptions[Revenue Y0],  _xlpm.base * 2 )")).toBe("LET( base, Assumptions[Revenue Y0],  base * 2 )");
  });

  it("warns about F6: a modern function stored bare, and _xludf. poisoning", () => {
    const bare = decompileWithDiagnostics("SEQUENCE(1,3)");
    expect(bare.text).toBe("SEQUENCE(1,3)");
    expect(codes(bare)).toEqual(["bare-prefix"]);
    const poisoned = decompileWithDiagnostics("_xludf.SEQUENCE(1,3)");
    expect(poisoned.text).toBe("_xludf.SEQUENCE(1,3)");
    expect(codes(poisoned)).toEqual(["poisoned"]);
    expect(codes(decompileWithDiagnostics("_xlfn.FUTUREFN(1)"))).toEqual(["unknown-function"]);
    expect(decompile("_xlfn.FUTUREFN(1)")).toBe("_xlfn.FUTUREFN(1)");
  });

  it("is idempotent on display text", () => {
    for (const d of ["LAMBDA(x, x+1)", "SORT(A1#)", "Tbl[@Col]", "@A1:A3", "LET(Rate,5,Rate*2)"]) {
      expect(decompile(d)).toBe(d);
    }
  });
});

describe("compile: display → stored", () => {
  it("prefixes from the catalogue", () => {
    expect(compile("SEQUENCE(1,5)")).toBe("_xlfn.SEQUENCE(1,5)");
    expect(compile("SORT(FILTER(A1:A9,B1:B9>0))")).toBe("_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A1:A9,B1:B9>0))");
    expect(compile("SUM(A1:A3)+IFERROR(1/0,0)")).toBe("SUM(A1:A3)+IFERROR(1/0,0)");
    expect(compile("concat(\"a\",\"b\")")).toBe('_xlfn.CONCAT("a","b")'); // Excel upper-cases function names
    expect(compile("STDEV.S(A1:A9)")).toBe("_xlfn.STDEV.S(A1:A9)");
  });

  it("_xlpm. at every declaration and use of a LAMBDA parameter or LET binding", () => {
    expect(compile("LAMBDA(b, g, b*(1+g))")).toBe("_xlfn.LAMBDA(_xlpm.b, _xlpm.g, _xlpm.b*(1+_xlpm.g))");
    expect(compile("LET(f, LAMBDA(t, t*2), f(3))")).toBe("_xlfn.LET(_xlpm.f, _xlfn.LAMBDA(_xlpm.t, _xlpm.t*2), _xlpm.f(3))");
    // An optional parameter is `_xlop.y` in the list, without brackets; its uses stay `_xlpm.y` (F9).
    expect(compile("LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))")).toBe("_xlfn.LAMBDA(_xlpm.x,_xlop.y,IF(_xlfn.ISOMITTED(_xlpm.y),_xlpm.x,_xlpm.x+_xlpm.y))");
    expect(compile("LAMBDA(x,x)(1)")).toBe("_xlfn.LAMBDA(_xlpm.x,_xlpm.x)(1)");
  });

  it("LET shadowing a defined name (f7_base C3)", () => {
    expect(compile("LET(Rate,5,Rate*2)", { names: ["Rate"] })).toBe("_xlfn.LET(_xlpm.Rate,5,_xlpm.Rate*2)");
    // Outside the LET, Rate is the defined name again.
    expect(compile("LET(Rate,5,Rate*2)+Rate", { names: ["Rate"] })).toBe("_xlfn.LET(_xlpm.Rate,5,_xlpm.Rate*2)+Rate");
    // A binding's own value is evaluated before the binding exists.
    expect(compile("LET(Rate,Rate*2,Rate)", { names: ["Rate"] })).toBe("_xlfn.LET(_xlpm.Rate,Rate*2,_xlpm.Rate)");
    // Case-insensitive, as Excel names are.
    expect(compile("LAMBDA(x, X+1)")).toBe("_xlfn.LAMBDA(_xlpm.x, _xlpm.X+1)");
  });

  it("scopes: an inner LAMBDA parameter does not leak", () => {
    expect(compile("LET(a, LAMBDA(z, z), a(1)+z)", { names: ["z"] })).toBe("_xlfn.LET(_xlpm.a, _xlfn.LAMBDA(_xlpm.z, _xlpm.z), _xlpm.a(1)+z)");
  });

  it("x# → ANCHORARRAY(x), @x → SINGLE(x), [@Col] → [[#This Row],[Col]]", () => {
    expect(compile("ROWS(E1#)")).toBe("ROWS(_xlfn.ANCHORARRAY(E1))");
    expect(compile("SUM('S1'!$B$1#)")).toBe("SUM(_xlfn.ANCHORARRAY('S1'!$B$1))");
    expect(compile("@A1:A10")).toBe("_xlfn.SINGLE(A1:A10)");
    expect(compile("Tbl[@Col]*2")).toBe("Tbl[[#This Row],[Col]]*2");
    expect(compile("Tbl[@[Col A]]")).toBe("Tbl[[#This Row],[Col A]]");
    expect(compile("[@Qty]*[@Price]")).toBe("[[#This Row],[Qty]]*[[#This Row],[Price]]");
  });

  it("eta-reduced function arguments → _xleta.", () => {
    expect(compile("GROUPBY(T[K], T[N], COUNTA, , 0)")).toBe("_xlfn.GROUPBY(T[K], T[N], _xleta.COUNTA, , 0)");
    expect(compile("BYROW(A1:C3, SUM)")).toBe("_xlfn.BYROW(A1:C3, _xleta.SUM)");
    // Not when the argument is a defined name, nor outside functions that take a LAMBDA.
    expect(compile("BYROW(A1:C3, Sum)", { names: ["Sum"] })).toBe("_xlfn.BYROW(A1:C3, Sum)");
    expect(compile("SUM(Rate)", { names: ["Rate"] })).toBe("SUM(Rate)");
  });

  it("re-qualifies the home sheet", () => {
    const ctx = { homeSheet: "BS", localNames: ["FixedAssetRoll_base"] };
    expect(compile("$C$6#", ctx)).toBe("_xlfn.ANCHORARRAY(BS!$C$6)");
    expect(compile("CHOOSEROWS(FixedAssetRoll_base, 1)", ctx)).toBe("_xlfn.CHOOSEROWS(BS!FixedAssetRoll_base, 1)");
    expect(compile("A1", { homeSheet: "S1" })).toBe("'S1'!A1");
    expect(compile("'IS'!A1+Model!B2", {})).toBe("IS!A1+Model!B2");
    expect(compile("IS!Sales+Other", { ...ctx, names: ["Other"] })).toBe("IS!Sales+Other");
    expect(compile("'SCF recursive'!A1+B2", { homeSheet: "SCF recursive" })).toBe("'SCF recursive'!A1+'SCF recursive'!B2");
  });

  it("is idempotent on stored text", () => {
    for (const s of [
      "_xlfn.LAMBDA(_xlpm.x, _xlpm.x+1)",
      "_xlfn.ANCHORARRAY('S1'!$B$1)",
      "_xlfn._xlws.SORT(A1:A3)",
      "_xlfn.LET(_xlpm.Rate, 5, _xlpm.Rate*2)",
      "_xlfn.GROUPBY(T[K], T[N], _xleta.COUNTA, , 0)",
      "Tbl[[#This Row],[Col]]",
      "_xlfn.SINGLE(A1:A3)",
    ]) {
      expect(compile(s)).toBe(s);
    }
  });

  it("writes CR LF on request, leaving strings alone", () => {
    expect(compile('LAMBDA(x,\n  x&"a\nb")', { crlf: true })).toBe('_xlfn.LAMBDA(_xlpm.x,\r\n  _xlpm.x&"a\nb")');
  });
});

describe("a cell called like a function (it holds a LAMBDA)", () => {
  it("is kept verbatim both ways: no prefix, no unknown function", () => {
    for (const s of ["C2(D2,E2)", "$C$2(1, 2)", "Sheet1!C2(1)", "'S 1'!$C$2(1)", "C2(1)(2)", "LOG10(C2(1))"]) {
      const c = compileWithDiagnostics(s);
      expect(c.text, s).toBe(s);
      expect(codes(c), s).toEqual([]);
      expect(decompile(c.text), s).toBe(s);
    }
  });
  it("in a sheet's name, the reference takes the sheet like any other; LOG10 stays the built-in", () => {
    const ctx = { homeSheet: "S 1" };
    expect(compile("C2(D2,E2)", ctx)).toBe("'S 1'!C2('S 1'!D2,'S 1'!E2)");
    expect(decompile("'S 1'!C2('S 1'!D2,'S 1'!E2)", ctx)).toBe("C2(D2,E2)");
    expect(compile("LOG10(100)")).toBe("LOG10(100)");
    expect(compile("ATAN2(1,2)+DAYS360(A1,B1)+T(A1)")).toBe("ATAN2(1,2)+DAYS360(A1,B1)+T(A1)");
  });
});

describe("compile: the traps", () => {
  it("F6: refuses an unknown function instead of emitting it bare", () => {
    const r = compileWithDiagnostics("SEQUENSE(3)");
    expect(r.text).toBe("");
    expect(r.diagnostics[0]).toMatchObject({ severity: "error", code: "unknown-function", start: 0, end: 8 });
    expect(() => compile("SEQUENSE(3)")).toThrow(FormulaError);
    expect(() => compile("SEQUENSE(3)")).toThrow(/line 1, column 1: unknown function 'SEQUENSE'/);
  });

  it("F6: SEQUENCE without a prefix gets one", () => {
    expect(compile("SEQUENCE(1,3)")).toBe("_xlfn.SEQUENCE(1,3)");
    expect(compile("COLUMNS(SEQUENCE(1,3))")).toBe("COLUMNS(_xlfn.SEQUENCE(1,3))");
  });

  it("calls to defined names are fine when the names are known", () => {
    expect(compile("P_Add1(41)", { names: ["P_Add1"] })).toBe("P_Add1(41)");
    expect(compile("Mod.Fn(2)", { names: ["Mod.Fn"] })).toBe("Mod.Fn(2)");
    expect(compile("Growλ(100, 0.1)", { names: ["Growλ"] })).toBe("Growλ(100, 0.1)");
    expect(() => compile("Mod.Fn(2)")).toThrow(/unknown function 'Mod.Fn'/);
    expect(compile("Mod.Fn(2)", { allowUnknownFunctions: true })).toBe("Mod.Fn(2)");
  });

  it("T12: Fact vs FACT, the built-in wins and we say so", () => {
    expect(builtinCollision("Fact")?.name).toBe("FACT");
    expect(builtinCollision("Growλ")).toBeUndefined();
    const r = compileWithDiagnostics("Fact(5)", { names: ["Fact"] });
    expect(r.text).toBe("FACT(5)");
    expect(codes(r)).toEqual(["builtin-collision"]);
    expect(r.diagnostics[0]!.message).toMatch(/calls the built-in FACT/);
    // The definition of the name itself is fine.
    expect(compile("LAMBDA(n, 1)")).toBe("_xlfn.LAMBDA(_xlpm.n, 1)");
  });

  it("warns on arity outside the catalogue's range", () => {
    expect(codes(compileWithDiagnostics("SEQUENCE()"))).toEqual(["arity"]);
    expect(codes(compileWithDiagnostics("IF(1,2,3,4)"))).toEqual(["arity"]);
  });

  it("keeps _xludf. and unknown _xlfn. names, with a warning", () => {
    const r = compileWithDiagnostics("_xludf.SEQUENCE(1,3)");
    expect(r.text).toBe("_xludf.SEQUENCE(1,3)");
    expect(codes(r)).toEqual(["poisoned"]);
    expect(codes(compileWithDiagnostics("_xlfn.FUTUREFN(1)"))).toEqual(["unknown-function"]);
  });

  it("corrects a wrong stored prefix", () => {
    const r = compileWithDiagnostics("_xlfn.FILTER(A1:A3,B1:B3)");
    expect(r.text).toBe("_xlfn._xlws.FILTER(A1:A3,B1:B3)");
    expect(codes(r)).toEqual(["wrong-prefix"]);
  });
});

describe("round trip on the probe definitions", () => {
  const cases: [string, string, object?][] = [
    ["_xlfn.LAMBDA(_xlpm.n, 1)", "LAMBDA(n, 1)"],
    ["_xlfn.LAMBDA(_xlpm.b,_xlpm.g, _xlpm.b*(1+_xlpm.g))", "LAMBDA(b,g, b*(1+g))"],
    ["_xlfn.LAMBDA(_xlpm.x, _xlpm.x*10)", "LAMBDA(x, x*10)"],
    ["_xlfn.ANCHORARRAY('S1'!$B$1)", "'S1'!$B$1#"],
    ["_xlfn.ANCHORARRAY('S1'!$B$1)", "$B$1#", { homeSheet: "S1" }],
    ["_xlfn.LET(_xlpm.Rate, 5, _xlpm.Rate*2)", "LET(Rate, 5, Rate*2)", { names: ["Rate"] }],
    ["_xlfn.LAMBDA(_xlpm.x, _xlpm.x*Rate)", "LAMBDA(x, x*Rate)", { names: ["Rate"] }],
  ];
  it.each(cases)("%s ⇄ %s", (stored, display, ctx = {}) => {
    expect(decompile(stored, ctx)).toBe(display);
    expect(compile(display, ctx)).toBe(stored);
  });
});
