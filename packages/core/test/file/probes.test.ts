// The Excel-saved probe workbooks (Mac and Windows) in probes/results.
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { formulaTextAt, readWorkbook, type WorkbookSnapshot } from "../../src/index.js";

const DIR = join(import.meta.dirname, "..", "..", "..", "..", "probes", "results");
const load = (f: string): WorkbookSnapshot => readWorkbook(new Uint8Array(readFileSync(join(DIR, f))));
const byName = (wb: WorkbookSnapshot, name: string, scope?: number) =>
  wb.definedNames.find(
    (n) => n.name === name && (scope === undefined ? n.scope.kind === "workbook" : n.scope.kind === "sheet" && n.scope.position === scope),
  );

describe("every probe workbook parses", () => {
  const files = readdirSync(DIR).filter((f) => f.endsWith(".xlsx"));
  it("finds the probe files", () => expect(files.length).toBeGreaterThanOrEqual(8));
  for (const f of files) {
    it(f, () => {
      const wb = load(f);
      expect(wb.workbookPart).toBe("xl/workbook.xml");
      // F8 builds its own sheets (N, D, Slot, Sh; P for the formula-free base); F10 and F11
      // (probes/kits) have Sheet1 and Labels; the LAMBDA
      // workbooks (lambda_optional_mac, F9) are new workbooks with Excel's default sheet.
      if (!["f8_", "f9_", "f10_", "f11_", "lambda_"].some((p) => f.startsWith(p))) expect(wb.sheets.map((s) => s.name)).toEqual(["S1", "S2"]);
      expect(wb.sheets.length).toBeGreaterThan(0);
      expect(wb.sheets.every((s) => s.part !== undefined && s.kind === "worksheet")).toBe(true);
      expect(wb.definedNames.length).toBeGreaterThan(0);
      expect(wb.warnings).toEqual([]);
    });
  }
});

describe("probe_win.xlsx", () => {
  const wb = load("probe_win.xlsx");

  it("reads sheets in order with position, id, state and part", () => {
    expect(wb.sheets.map(({ name, sheetId, position, state, part }) => ({ name, sheetId, position, state, part }))).toEqual([
      { name: "S1", sheetId: 1, position: 0, state: "visible", part: "xl/worksheets/sheet1.xml" },
      { name: "S2", sheetId: 2, position: 1, state: "visible", part: "xl/worksheets/sheet2.xml" },
    ]);
  });

  it("reads every defined name in file order", () => {
    expect(wb.definedNames.map((n) => n.name)).toEqual([
      "Fact", "Growλ", "Mod.Fn", "P_Add1", "P_Loc", "P_Local", "P_Local2", "P_Long3990", "P_Multi", "P_Spill",
    ]);
    expect(wb.definedNames.map((n) => n.index)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
    expect(wb.definedNames.some((n) => n.isXlPrefixed || n.hidden)).toBe(false);
  });

  it("reads the comment", () => {
    expect(byName(wb, "P_Add1")?.comment).toBe("probe comment");
    expect(byName(wb, "P_Add1")?.attributes).toEqual({});
    expect(byName(wb, "Fact")?.comment).toBeUndefined();
  });

  it("scopes localSheetId by position", () => {
    expect(byName(wb, "P_Local", 1)?.scope).toEqual({ kind: "sheet", position: 1, name: "S2" });
    expect(byName(wb, "P_Local", 1)?.definition).toBe("5");
  });

  it("keeps CR LF inside a definition", () => {
    expect(byName(wb, "P_Multi")?.definition).toBe("_xlfn.LAMBDA(_xlpm.x,\r\n  _xlpm.x*2)");
  });

  it("keeps stored prefixes and ANCHORARRAY", () => {
    expect(byName(wb, "P_Spill")?.definition).toBe("_xlfn.ANCHORARRAY('S1'!$B$1)");
    expect(byName(wb, "Growλ")?.definition).toBe("_xlfn.LAMBDA(_xlpm.b,_xlpm.g, _xlpm.b*(1+_xlpm.g))");
    expect(byName(wb, "P_Long3990")?.definition.length).toBeGreaterThan(3990);
  });

  it("reads cell formulas and cached values", () => {
    const s1 = wb.sheets[0]!;
    const at = (c: string) => s1.formulas.find((f) => f.cell === c);
    expect(at("A1")).toMatchObject({ kind: "dynamic-array", text: "P_Add1(41)", range: "A1", cm: 1, value: { type: "n", value: 42 } });
    expect(at("A4")).toMatchObject({ kind: "normal", text: "P_K2", value: { type: "e", value: "#NAME?" } });
    expect(at("A10")).toMatchObject({ text: "Growλ(100, 0.1)" });
    expect(s1.spills).toContainEqual({ anchor: "B1", extent: "B1:F1" });
    expect(s1.formulas.map((f) => f.cell)).toEqual(["A1", "B1", "C2", "A3", "A4", "A5", "A6", "A7", "A9", "A10"]);
  });
});

describe("f7_base.xlsx", () => {
  const wb = load("f7_base.xlsx");
  const [s1, s2] = wb.sheets as [WorkbookSnapshot["sheets"][0], WorkbookSnapshot["sheets"][0]];

  it("has 13 shared-formula cells in 2 groups", () => {
    const shared = wb.sheets.flatMap((s) => s.formulas.filter((f) => f.kind.startsWith("shared")));
    expect(shared).toHaveLength(13);
    expect(s1.sharedFormulas).toEqual([
      { si: 0, master: "B2", range: "B2:B11", text: "A2*Rate", cells: ["B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "B10", "B11"] },
    ]);
    expect(s2.sharedFormulas).toEqual([{ si: 0, master: "A3", range: "A3:A5", text: "Loc*ROW()", cells: ["A3", "A4", "A5"] }]);
    const b5 = s1.formulas.find((f) => f.cell === "B5")!;
    expect(b5).toMatchObject({ kind: "shared-child", si: 0, master: "B2", text: undefined, value: { value: 0.4 } });
    // The hook for W2: a shifter receives the master text and the offset.
    expect(formulaTextAt(s1, b5, (t, r, c) => `${t}@${r},${c}`)).toBe("A2*Rate@3,0");
  });

  it("finds the spill anchor and its saved extent", () => {
    expect(s1.spills).toContainEqual({ anchor: "E1", extent: "E1:E3" });
    expect(s1.formulas.find((f) => f.cell === "E1")).toMatchObject({ kind: "dynamic-array", text: "_xlfn.SEQUENCE(3)*Rate" });
  });

  it("keeps other <f> attributes raw", () => {
    expect(s1.formulas.find((f) => f.cell === "C7")).toMatchObject({
      kind: "dynamic-array",
      attributes: { aca: "1", ca: "1" },
      value: { type: "e", value: "#REF!" },
    });
  });

  it("decodes entities in formula text and reads str values", () => {
    expect(s1.formulas.find((f) => f.cell === "C2")).toMatchObject({
      text: '"Rate is "&Rate',
      value: { type: "str", value: "Rate is 0.1" },
    });
  });

  it("reads one CF and one DV formula", () => {
    expect(s1.conditionalFormats).toEqual([{ sqref: "A2:A11", type: "expression", priority: 1, formulas: ["A2>Rate*50"], ext: false }]);
    expect(s1.dataValidations).toEqual([{ sqref: "G1", type: "decimal", formula1: "Rate+Loc", formula2: undefined, ext: false }]);
    expect(s2.conditionalFormats).toEqual([]);
  });

  it("separates a workbook name from a sheet name of the same name", () => {
    expect(byName(wb, "Loc")?.definition).toBe("3");
    expect(byName(wb, "Loc", 1)?.definition).toBe("7");
  });

  it("counts all <f> elements", () => {
    expect(wb.sheets.reduce((n, s) => n + s.formulas.length, 0)).toBe(28);
  });
});

describe("probe_patched.xlsx and probe_resaved.xlsx", () => {
  it("reads names written by the file patch and lists the custom XML part", () => {
    for (const f of ["probe_patched.xlsx", "probe_resaved.xlsx"]) {
      const wb = load(f);
      expect(byName(wb, "Z_Add2")?.comment).toBe("doc written by file patch");
      expect(byName(wb, "Z_Loc", 1)?.scope).toEqual({ kind: "sheet", position: 1, name: "S2" });
      expect(wb.otherParts).toEqual([{ kind: "customXml", path: "customXml/item1.xml" }]);
    }
    expect(byName(load("probe_resaved.xlsx"), "Z_Bare")?.definition).toBe("_xludf.SEQUENCE(1,3)");
  });
});
