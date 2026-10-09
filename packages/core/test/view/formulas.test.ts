import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { readWorkbook, type CellFormula, type Sheet, type WorkbookSnapshot } from "../../src/file/index.js";
import { renderFormulaView, sheetFormulaView, shortValue, UNCALCULATED } from "../../src/view/index.js";
import { applyChangeSet, audit } from "../../src/index.js";

const RESULTS = join(import.meta.dirname, "..", "..", "..", "..", "probes", "results");
const f7 = readWorkbook(new Uint8Array(readFileSync(join(RESULTS, "f7_base.xlsx"))));

describe("sheetFormulaView on f7_base.xlsx", () => {
  const s1 = sheetFormulaView(f7, "S1");
  const at = (cell: string) => s1.find((l) => l.cell === cell)!;

  it("lists every formula once, row by row, left to right", () => {
    expect(s1.map((l) => l.cell)).toEqual([
      "A1", "C1", "E1", "B2", "C2", "B3", "C3", "B4", "C4", "B5", "C5",
      "B6", "C6", "B7", "C7", "B8", "C8", "B9", "C9", "B10", "C10", "B11",
    ]);
  });

  it("gives shared children their own text and their master", () => {
    expect(at("B2")).toMatchObject({ kind: "shared", formula: "A2*Rate", extent: "B2:B11", groupSize: 10, master: "B2" });
    expect(at("B3")).toMatchObject({ kind: "shared", formula: "A3*Rate", stored: "A3*Rate", master: "B2" });
    expect(at("B11").formula).toBe("A11*Rate");
  });

  it("shows a spill anchor with its saved extent, in display form", () => {
    expect(at("E1")).toMatchObject({ kind: "dynamic-array", extent: "E1:E3", rows: 3, cols: 1, formula: "SEQUENCE(3)*Rate", valueText: "0.1" });
    expect(at("C9").formula).toBe("ROWS(E1#)*Rate");
    expect(at("C9").refs).toEqual([{ sheet: "S1", address: "E1", refKind: "cell", start: 5, end: 7 }]);
    expect(at("C7")).toMatchObject({ formula: 'INDIRECT("Rate")', valueText: "#REF!", names: [] });
    expect(at("C2").valueText).toBe('"Rate is 0.1"');
    expect(at("B4").valueText).toBe("0.3");
  });

  it("resolves names with sheet scope and LET shadowing, with their spans", () => {
    const names = (cell: string) => at(cell).names.map((n) => [n.key, at(cell).formula.slice(n.start, n.end)]);
    expect(names("B3")).toEqual([["Rate", "Rate"]]);
    expect(names("C3")).toEqual([]); // LET(Rate, 5, Rate*2): the variable, not the name
    expect(names("C5")).toEqual([["S2!Loc", "Loc"], ["Loc", "Loc"]]);
    expect(names("C10")).toEqual([["Rate2", "Rate2"], ["Fn", "Fn"]]);
    expect(names("C2")).toEqual([["Rate", "Rate"]]); // not the "Rate" inside the string
    const s2 = sheetFormulaView(f7, "S2");
    expect(s2.find((l) => l.cell === "A2")!.names.map((n) => n.key)).toEqual(["Rate", "S2!Loc"]);
    expect(s2.find((l) => l.cell === "A4")).toMatchObject({ formula: "Loc*ROW()", master: "A3" });
  });

  it("throws for a sheet the workbook does not have", () => {
    expect(() => sheetFormulaView(f7, "Nope")).toThrow(/no sheet/);
  });
});

function formula(cell: string, kind: CellFormula["kind"], text: string | undefined, extra: Partial<CellFormula> = {}): CellFormula {
  return { cell, kind, text, attributes: {}, value: { type: "n", raw: "1", value: 1 }, ...extra };
}

function snapshot(formulas: CellFormula[], spills: Sheet["spills"] = []): WorkbookSnapshot {
  const sheet: Sheet = {
    name: "Calc",
    sheetId: 1,
    position: 0,
    state: "visible",
    kind: "worksheet",
    relId: "rId1",
    part: "xl/worksheets/sheet1.xml",
    formulas,
    sharedFormulas: [],
    spills,
    conditionalFormats: [],
    dataValidations: [],
    tables: [],
  };
  return { sheets: [sheet], definedNames: [], tables: [], charts: [], otherParts: [], parts: [], foreignModuleStores: [], workbookPart: "xl/workbook.xml", warnings: [] };
}

describe("sheetFormulaView on arrays, spills and data tables", () => {
  const wb = snapshot(
    [
      formula("C1", "dynamic-array", "_xlfn.SEQUENCE(3)", { range: "C1:C3" }),
      // A stale <f> inside C1's spill: Excel shows the cell as part of C1#.
      formula("C2", "normal", "1+1"),
      formula("A4", "array", "{1,2;3,4}", { range: "A4:B5" }),
      formula("B5", "normal", "2"),
      formula("E4", "data-table", undefined, { range: "E4:E6", attributes: { dt2D: "0", dtr: "0", r1: "A1" } }),
      formula("F4", "data-table", undefined, { range: "F4:G6", attributes: { dt2D: "1", r1: "A1", r2: "B1" } }),
      formula("A2", "dynamic-array", "Missing+1", { range: "A2", value: { type: "e", raw: "#NAME?", value: "#NAME?" } }),
    ],
    [{ anchor: "C1", extent: "C1:C3" }],
  );
  const view = sheetFormulaView(wb, "calc");

  it("skips cells inside another cell's spill, array or data table", () => {
    expect(view.map((l) => l.cell)).toEqual(["C1", "A2", "A4", "E4", "F4"]);
  });

  it("shows data tables as TABLE(row input, column input)", () => {
    expect(view.find((l) => l.cell === "E4")).toMatchObject({ kind: "data-table", formula: "TABLE(, A1)", rows: 3, cols: 1, stored: "" });
    expect(view.find((l) => l.cell === "F4")!.formula).toBe("TABLE(A1, B1)");
  });

  it("keeps an unknown name as a use without a key", () => {
    expect(view.find((l) => l.cell === "A2")).toMatchObject({ rows: 1, cols: 1, valueText: "#NAME?", names: [{ id: "Missing", key: undefined, start: 0, end: 7 }] });
  });
});

describe("shortValue", () => {
  it("keeps values short and readable", () => {
    expect(shortValue({ type: "n", raw: "", value: 0.30000000000000004 })).toBe("0.3");
    expect(shortValue({ type: "n", raw: "", value: 1234567.891 })).toBe("1234567.891");
    expect(shortValue({ type: "n", raw: "", value: 1.5e20 })).toBe("1.5e+20");
    expect(shortValue({ type: "b", raw: "", value: true })).toBe("TRUE");
    expect(shortValue({ type: "str", raw: "", value: 'say "hi"' })).toBe('"say ""hi"""');
    expect(shortValue({ type: "n", raw: undefined, value: undefined })).toBeUndefined();
  });
});

describe("renderFormulaView", () => {
  it("aligns cells, kinds, formulas and values, and places names in the document", () => {
    const lines = sheetFormulaView(f7, "S1");
    const r = renderFormulaView(lines, { sheet: "S1", workbook: "f7_base.xlsx" });
    const text = r.text.split("\n");
    expect(text[0]).toBe("// Sheet S1: 22 formulas in order of appearance (row by row, left to right). Read-only view of f7_base.xlsx.");
    expect(text).toContain("Spl  E1#   (3×1)        = SEQUENCE(3)*Rate      → 0.1 …");
    expect(text).toContain("     B3    shared ← B2  = A3*Rate               → 0.2");
    expect(r.entries).toHaveLength(22);
    for (const e of r.entries) {
      const l = lines[e.index]!;
      expect(r.text.slice(e.address.start, e.address.end)).toBe(l.cell);
      expect(e.names.map((s) => r.text.slice(s.start, s.end))).toEqual(l.names.map((n) => l.formula.slice(n.start, n.end)));
      expect(e.refs.map((s) => r.text.slice(s.start, s.end))).toEqual(l.refs.map((n) => l.formula.slice(n.start, n.end)));
      expect(text[e.line]!.slice(5).startsWith(l.cell)).toBe(true);
      expect(e.lhs.map((s) => r.text.slice(s.start, s.end))).toEqual(l.lhs.map((n) => n.display));
    }
  });

  it("gives a long formula a pretty-printed block and still places its names", () => {
    const long = "_xlfn.LET(_xlpm.a, Rate * 1000000, _xlpm.b, Rate2 + Rate + RateX + Loc + Fn(1) + Fn(2) + Fn(3), _xlpm.a + _xlpm.b + SUM(A1:A20))";
    const wb = snapshot([formula("B2", "normal", long), formula("B3", "normal", "B2")]);
    wb.definedNames = ["Rate", "Rate2", "RateX", "Loc", "Fn"].map((name, index) => ({
      name, scope: { kind: "workbook" }, hidden: false, comment: undefined, definition: "1", attributes: {}, index, isXlPrefixed: false, isBuiltIn: false,
    }));
    const lines = sheetFormulaView(wb, "Calc");
    const r = renderFormulaView(lines, { sheet: "Calc", formulaWidth: 40, width: 60 });
    const e = r.entries[0]!;
    expect(e.lastLine).toBeGreaterThan(e.line + 2);
    const text = r.text.split("\n");
    expect(text[e.line]).toBe("B2  = LET(");
    // The value column is narrow here (the other formula is short): the value follows the block.
    expect(text[e.lastLine]).toBe(" ".repeat(10) + "→ 1");
    expect(e.names.map((s) => r.text.slice(s.start, s.end))).toEqual(["Rate", "Rate2", "Rate", "RateX", "Loc", "Fn", "Fn", "Fn"]);
    expect(e.refs.map((s) => r.text.slice(s.start, s.end))).toEqual(["A1:A20"]);
    expect(r.entries[1]!.line).toBe(e.lastLine + 2); // a blank line after the block
  });

  it("says when a sheet has no formulas", () => {
    expect(renderFormulaView([], { sheet: "Empty" }).text).toContain("// No formulas on this sheet.");
  });
});

describe("left-hand side: the defined names on a line's location", () => {
  type Def = WorkbookSnapshot["definedNames"][number];
  const def = (name: string, definition: string, extra: Partial<Def> = {}): Def => ({
    name, scope: { kind: "workbook" }, hidden: false, comment: undefined, definition, attributes: {}, index: 0,
    isXlPrefixed: name.toLowerCase().startsWith("_xl"), isBuiltIn: name.startsWith("_xlnm."), ...extra,
  });
  const wb = snapshot(
    [
      formula("C1", "dynamic-array", "_xlfn.SEQUENCE(3)", { range: "C1:C3" }),
      formula("B2", "normal", "1"),
      formula("D2", "shared-master", "B2+1", { si: 0 }),
      formula("B3", "normal", "B2*2"),
      formula("D3", "shared-child", undefined, { si: 0 }),
      formula("A4", "array", "{1,2;3,4}", { range: "A4:B5" }),
      formula("E4", "data-table", undefined, { range: "E4:E6", attributes: { dt2D: "0", dtr: "0", r1: "A1" } }),
    ],
    [{ anchor: "C1", extent: "C1:C3" }],
  );
  wb.sheets[0]!.sharedFormulas = [{ si: 0, master: "D2", range: "D2:D3", text: "B2+1", cells: ["D2", "D3"] }];
  wb.sheets.push({ ...wb.sheets[0]!, name: "Other", sheetId: 2, position: 1, formulas: [], spills: [], sharedFormulas: [] });
  const other = { kind: "sheet" as const, name: "Other", position: 1 };
  const calc = { kind: "sheet" as const, name: "Calc", position: 0 };
  wb.definedNames = [
    def("Price", "'Calc'!$B$2"),
    def("Total", "Calc!$B$3"),
    def("Qty", "Calc!B3", { scope: calc }),
    def("Series", "Calc!$C$1:$C$3"),
    def("SeriesS", "_xlfn.ANCHORARRAY(Calc!$C$1)"),
    def("Block", "Calc!$B$2:$B$3"),
    def("Group", "Calc!$D$2:$D$3"),
    def("Elsewhere", "Other!$B$2"),
    def("Mine", "Calc!$B$2", { scope: other }),
    def("hid", "Calc!$A$4", { hidden: true }),
    def("ArrayAll", "(Calc!$A$4:$B$5)"),
    def("AVeryLongNameForTheDataTableThatOverflows", "Calc!$E$4:$E$6"),
    def("Child", "Calc!$D$3"),
    def("Formula", "Calc!$B$2+1"),
    def("_xlnm.Print_Area", "Calc!$B$2", { scope: calc }),
  ];
  const view = sheetFormulaView(wb, "Calc");
  const lhs = (cell: string) => view.find((l) => l.cell === cell)!.lhs.map((n) => `${n.display}:${n.target}`);

  it("names a cell, a spill by x# or by its extent, an array by its range, sorted", () => {
    expect(lhs("B2")).toEqual(["Other!Mine:cell", "Price:cell"]);
    expect(lhs("B3")).toEqual(["Qty:cell", "Total:cell"]);
    expect(lhs("C1")).toEqual(["Series:extent", "SeriesS:spill"]);
    expect(lhs("A4")).toEqual(["ArrayAll:extent", "hid:cell"]);
    expect(lhs("E4")).toEqual(["AVeryLongNameForTheDataTableThatOverflows:extent"]);
    expect(lhs("D3")).toEqual(["Child:cell"]);
  });

  it("does not list a larger block, a shared group, another sheet's cell, a formula or a built-in", () => {
    expect(lhs("D2")).toEqual([]);
    for (const n of ["Block", "Group", "Elsewhere", "Formula", "_xlnm.Print_Area"]) expect(view.some((l) => l.lhs.some((x) => x.name === n)), n).toBe(false);
    expect(sheetFormulaView(wb, "Other")).toEqual([]);
  });

  it("keys names as the project does, and marks hidden ones", () => {
    expect(view.find((l) => l.cell === "B3")!.lhs.map((n) => n.key)).toEqual(["Calc!Qty", "Total"]);
    expect(view.find((l) => l.cell === "B2")!.lhs[0]).toMatchObject({ key: "Other!Mine", scope: "Other", name: "Mine" });
    expect(view.find((l) => l.cell === "A4")!.lhs.find((n) => n.name === "hid")!.hidden).toBe(true);
  });

  it("renders the names as a first column, a long list on a line above, with spans", () => {
    const r = renderFormulaView(view, { sheet: "Calc" });
    const text = r.text.split("\n");
    expect(text).toContain("Series, SeriesS    C1#  (3×1)        = SEQUENCE(3)  → 1 …");
    expect(text).toContain("Other!Mine, Price  B2                = 1            → 1");
    const e4 = r.entries.find((e) => e.cell === "E4")!;
    expect(text[e4.line]).toBe("AVeryLongNameForTheDataTableThatOverflows");
    expect(text[e4.line + 1]!.startsWith(" ".repeat(19) + "E4")).toBe(true);
    for (const e of r.entries) {
      const l = view[e.index]!;
      expect(r.text.slice(e.address.start, e.address.end)).toBe(l.cell);
      expect(e.lhs.map((s) => r.text.slice(s.start, s.end))).toEqual(l.lhs.map((n) => n.display));
      expect(e.start).toBe(e.lhs[0]?.start ?? e.address.start - 19);
    }
    const plain = renderFormulaView(view, { sheet: "Calc", nameColumn: false });
    expect(plain.text.split("\n")).toContain("B2                = 1            → 1");
    expect(plain.entries.every((e) => e.lhs.length === 0)).toBe(true);
  });

  it("finds P_Spill and Spl on the probe fixtures", () => {
    const probe = readWorkbook(new Uint8Array(readFileSync(join(RESULTS, "probe_mac.xlsx"))));
    expect(sheetFormulaView(probe, "S1").find((l) => l.cell === "B1")!.lhs).toEqual([
      { key: "P_Spill", name: "P_Spill", scope: undefined, display: "P_Spill", hidden: false, target: "spill" },
    ]);
    expect(sheetFormulaView(f7, "S1").filter((l) => l.lhs.length > 0).map((l) => l.cell)).toEqual(["E1"]);
    expect(sheetFormulaView(f7, "S2").every((l) => l.lhs.length === 0)).toBe(true);
  });
});

describe("a formula a build wrote: not calculated since (FEEDBACK 2026-10-08)", () => {
  // The build drops the cell's <v> and sets fullCalcOnLoad="1"; Excel drops the flag when
  // it saves (probe F06). Until then the formula has no value, and a spill no known extent.
  const bytes = new Uint8Array(readFileSync(join(RESULTS, "f7_base.xlsx")));
  const built = readWorkbook(applyChangeSet(bytes, [{ op: "set-cell-formula", sheet: "S1", range: "E1", stored: "_xlfn.SEQUENCE(4)*Rate", display: "SEQUENCE(4)*Rate" }]));
  const lines = sheetFormulaView(built, "S1");
  const at = (cell: string) => lines.find((l) => l.cell === cell)!;

  it("is marked uncalculated, with no value; the other formulas keep theirs", () => {
    expect(f7.fullCalcOnLoad).toBeUndefined();
    expect(built.fullCalcOnLoad).toBe(true);
    expect(at("E1")).toMatchObject({ kind: "dynamic-array", formula: "SEQUENCE(4)*Rate", uncalculated: true, valueText: undefined });
    expect(at("E1").value).toBeUndefined();
    expect(at("B4")).toMatchObject({ valueText: "0.3" });
    expect(at("B4").uncalculated).toBeUndefined();
    expect(sheetFormulaView(f7, "S1").some((l) => l.uncalculated)).toBe(false);
  });

  it("reads so in the text view, a spill without a size", () => {
    const text = renderFormulaView(lines, { sheet: "S1" }).text;
    const e1 = text.split("\n").find((l) => l.includes("E1#"))!;
    expect(e1).toContain("(spill)");
    expect(e1).toContain(`→ ${UNCALCULATED}`);
  });

  it("the spill census counts it apart instead of dropping it", () => {
    const census = audit(built, { workbook: "built.xlsx" }).spills;
    expect(census.uncalculated).toBe(1);
    expect(census.spills.some((s) => s.anchor === "E1")).toBe(false);
    expect(audit(f7, { workbook: "f7.xlsx" }).spills.uncalculated).toBe(0);
  });
});
