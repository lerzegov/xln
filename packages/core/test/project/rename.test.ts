// Stretch G at the formula level: the token rewrite of a renamed name, on the traps of
// probe F7 (strings, LET variables, case, scope, shared prefixes) and the ones it left out.
import { describe, expect, it } from "vitest";
import { RenameContext, renameInFormula, type NameRename, type ScopedName } from "../../src/index.js";

// f7_base.xlsx: Rate, RateX, Rate2, Fn, Loc (workbook), S2!Loc, Spl.
const NAMES: ScopedName[] = [
  { name: "Rate", scope: undefined },
  { name: "RateX", scope: undefined },
  { name: "Rate2", scope: undefined },
  { name: "Fn", scope: undefined },
  { name: "Loc", scope: undefined },
  { name: "Loc", scope: "S2" },
  { name: "Spl", scope: undefined },
];

const ctx = (renames: NameRename[], names = NAMES) => new RenameContext(names, renames);
const RATE = ctx([{ scope: undefined, from: "Rate", to: "Growth" }]);
const LOC2 = ctx([{ scope: "S2", from: "Loc", to: "Spot" }]);
const rn = (text: string, home: string | undefined, c = RATE) => renameInFormula(text, home, c).text;

describe("renameInFormula: F7's traps", () => {
  it("renames the token and nothing that only looks like it", () => {
    expect(rn("A2*Rate", "S1")).toBe("A2*Growth");
    expect(rn("RateX+Rate", "S1")).toBe("RateX+Growth");
    expect(rn('"Rate is "&Rate', "S1")).toBe('"Rate is "&Growth');
    expect(rn("rate*3", "S1")).toBe("Growth*3");
    expect(rn("_xlfn.ROWS(E1#)*Rate", "S1")).toBe("_xlfn.ROWS(E1#)*Growth");
    expect(rn("Rate2+Fn(1)", "S1")).toBe("Rate2+Fn(1)");
  });

  it("leaves LET and LAMBDA variables of the same spelling, stored and display form", () => {
    expect(rn("_xlfn.LET(_xlpm.Rate,5,_xlpm.Rate*2)", "S1")).toBe("_xlfn.LET(_xlpm.Rate,5,_xlpm.Rate*2)");
    expect(rn("LET(Rate, 5, Rate*2)", "S1")).toBe("LET(Rate, 5, Rate*2)");
    expect(rn("LET(x, Rate, LET(Rate, 2, Rate*x))", "S1")).toBe("LET(x, Growth, LET(Rate, 2, Rate*x))");
    expect(rn("_xlfn.LAMBDA(_xlpm.x, _xlpm.x*Rate)", undefined)).toBe("_xlfn.LAMBDA(_xlpm.x, _xlpm.x*Growth)");
  });

  it("leaves strings, INDIRECT's text and structured references alone", () => {
    expect(rn('INDIRECT("Rate")', "S1")).toBe('INDIRECT("Rate")');
    expect(rn("SUM(Tbl[Rate])*Rate", "S1")).toBe("SUM(Tbl[Rate])*Growth");
    expect(rn("Tbl[[#This Row],[Rate]]", "S1")).toBe("Tbl[[#This Row],[Rate]]");
  });

  it("a sheet-local name shadows the workbook one on its sheet only", () => {
    expect(rn("S2!Loc+Loc", "S1", LOC2)).toBe("S2!Spot+Loc");
    expect(rn("Loc", "S2", LOC2)).toBe("Spot");
    expect(rn("Rate+Loc", "S2", LOC2)).toBe("Rate+Spot");
    expect(rn("'S2'!Loc*2", "S1", LOC2)).toBe("'S2'!Spot*2");
    expect(rn("S1!A1+S2!Loc", "S2", LOC2)).toBe("S1!A1+S2!Spot");
    // The workbook Loc, renamed, is not the one S2's formulas read bare.
    const wbLoc = ctx([{ scope: undefined, from: "Loc", to: "Where" }]);
    expect(rn("Loc", "S2", wbLoc)).toBe("Loc");
    expect(rn("Loc", "S1", wbLoc)).toBe("Where");
    // S1!Loc: S1 has no local Loc, so it reads the workbook's.
    expect(rn("S1!Loc", "S2", wbLoc)).toBe("S1!Where");
  });

  it("spill references: Name# stored as ANCHORARRAY(Name), and in display form", () => {
    const spl = ctx([{ scope: undefined, from: "Spl", to: "Series" }]);
    expect(rn("SUM(_xlfn.ANCHORARRAY(Spl))", "S1", spl)).toBe("SUM(_xlfn.ANCHORARRAY(Series))");
    expect(rn("SUM(Spl#)", "S1", spl)).toBe("SUM(Series#)");
  });

  it("LAMBDA names called like functions, and this workbook's [0]!Name", () => {
    const fn = ctx([{ scope: undefined, from: "Fn", to: "Fx" }]);
    expect(rn("Fn(10)", "S1", fn)).toBe("Fx(10)");
    expect(rn("Rate2+Fn(1)", "S1", fn)).toBe("Rate2+Fx(1)");
    expect(rn("[0]!Rate", undefined)).toBe("[0]!Growth");
    // Another workbook's Rate is not ours.
    expect(rn("[1]!Rate+[Other.xlsx]S1!Rate", "S1")).toBe("[1]!Rate+[Other.xlsx]S1!Rate");
  });

  it("keeps layout, line breaks and everything else byte for byte", () => {
    expect(rn("LET(\r\n  a, Rate,\r\n  a * 2\r\n)", "S1")).toBe("LET(\r\n  a, Growth,\r\n  a * 2\r\n)");
    const r = renameInFormula("Rate + RATE + rate", "S1", RATE);
    expect(r).toMatchObject({ text: "Growth + Growth + Growth", count: 3 });
  });

  it("several renames at once, read on the names before", () => {
    const swap = ctx([
      { scope: undefined, from: "Rate", to: "RateX" },
      { scope: undefined, from: "RateX", to: "Rate" },
    ]);
    expect(rn("Rate+RateX*2", "S1", swap)).toBe("RateX+Rate*2");
  });
});

describe("renameInFormula: refusals", () => {
  it("a new name a LET or LAMBDA variable would capture", () => {
    const toX = ctx([{ scope: undefined, from: "Rate", to: "x" }]);
    const r = renameInFormula("LAMBDA(x, x*Rate)", undefined, toX);
    expect(r.captured).toEqual(["x"]);
    expect(r.text).toBe("LAMBDA(x, x*Rate)");
    // Stored form: `_xlpm.x` is the variable, but the display form would read `x*x`.
    expect(renameInFormula("_xlfn.LAMBDA(_xlpm.x, _xlpm.x*Rate)", undefined, toX).captured).toEqual(["x"]);
  });

  it("a new name a sheet's local name would capture", () => {
    // Renaming the workbook Rate to Loc: on S2, Loc is S2's own.
    const toLoc = ctx([{ scope: undefined, from: "Rate", to: "Loc" }], NAMES.filter((n) => n.name !== "Loc" || n.scope === "S2"));
    expect(renameInFormula("Rate*2", "S2", toLoc).captured).toEqual(["Loc"]);
    expect(rn("Rate*2", "S1", toLoc)).toBe("Loc*2");
  });

  it("an existing reader of the new spelling that the renamed name would capture", () => {
    // S2!Loc → Spot while a workbook name Spot exists: S2's bare Spot would change meaning.
    const names = [...NAMES, { name: "Spot", scope: undefined }];
    const c = ctx([{ scope: "S2", from: "Loc", to: "Spot" }], names);
    expect(renameInFormula("Spot+1", "S2", c).captured).toEqual(["Spot"]);
    expect(renameInFormula("Spot+1", "S1", c).count).toBe(0);
    expect(renameInFormula("Spot+1", "S1", c).captured).toBeUndefined();
    // An unknown name of the new spelling would start to read the renamed name.
    expect(renameInFormula("Growth*2", "S1", RATE).captured).toEqual(["Growth"]);
  });

  it("a formula that does not parse but may hold the name", () => {
    expect(renameInFormula("Rate+(", "S1", RATE)).toMatchObject({ unparsed: true, count: 0, text: "Rate+(" });
    expect(renameInFormula("Other+(", "S1", RATE).unparsed).toBeUndefined();
  });
});
