// pullProject on the probe workbooks saved by Excel (probes/results).
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { definitionHash, LOCK_FILE, MANIFEST_FILE, parseModule, pullProject } from "../../src/index.js";
import { checkProject } from "./check.js";

const RESULTS = join(import.meta.dirname, "..", "..", "..", "..", "probes", "results");
const load = (f: string) => new Uint8Array(readFileSync(join(RESULTS, f)));
const fixtures = readdirSync(RESULTS).filter((f) => f.endsWith(".xlsx")).sort();

describe("pullProject on every probe workbook", () => {
  for (const f of fixtures) {
    it(f, () => {
      const r = pullProject(load(f), f);
      const c = checkProject(r);
      expect(c.failures).toEqual([]);
      expect(c.entries).toBe(r.report.names);
      // Deterministic: a second pull writes the same bytes.
      expect(pullProject(load(f), f).files).toEqual(r.files);
      // Files are LF only and parse as JSON where they should.
      for (const [p, t] of Object.entries(r.files)) {
        expect(t.includes("\r"), p).toBe(false);
        if (p.endsWith(".json")) JSON.parse(t);
      }
    });
  }
});

describe("probe_win.xlsx", () => {
  const r = pullProject(load("probe_win.xlsx"), "some/folder/probe_win.xlsx");

  it("reports names by kind, scope and module", () => {
    expect(r.report.workbook).toBe("probe_win.xlsx");
    expect(r.report.names).toBe(10);
    expect(r.report.byKind).toMatchObject({ lambda: 7, constant: 2, spill: 1 });
    expect(r.report.byScope).toEqual({ workbook: 8, sheet: 2, perSheet: { S2: 2 } });
    expect(r.report.modules.map((m) => [m.module, m.names])).toEqual([
      ["Mod", 1],
      ["P", 6],
      ["_unmanaged", 2],
    ]);
    expect(Object.keys(r.files)).toEqual(["names/Mod.xln", "names/P.xln", "names/_unmanaged.xln", "names/sheets/S1.xln", "names/sheets/S2.xln", MANIFEST_FILE, LOCK_FILE]);
    expect(r.report.cells).toEqual({ named: 1, slots: 0, unnamed: 10, blocks: 0, blockCells: 0 });
  });

  it("writes AFE syntax with doc comments, per-name @sheet and decompiled, LF-only definitions", () => {
    const p = r.files["names/P.xln"]!;
    expect(p).toContain("/** probe comment */\nP_Add1 = LAMBDA(x, x+1);");
    expect(p).toContain("P_Multi = LAMBDA(x,\n  x*2);");
    // Sheet-scoped members of a module stay in it, each with @sheet(S2) above it (no blocks, 2026-10-07).
    expect(p).toContain("  x*2);\n\n@sheet('S2')\nP_Local = 5;\n\n@sheet('S2')\nP_Local2 = 6;\n");
    expect(p).not.toContain("@scope");
    // A module's name on a formula cell is a cell statement: it lives in its sheet's file,
    // with @workbook above it and a bare address, among the sheet's cells in order.
    expect(p).not.toContain("P_Spill");
    expect(r.files["names/sheets/S1.xln"]).toContain("\n\n@A1 = P_Add1(41);\n@workbook\nP_Spill @B1# = SEQUENCE(1,5);\n@C2 = COLUMNS(P_Spill);\n");
    expect(r.files["names/sheets/S1.xln"]).not.toContain("@scope");
    expect(p.match(/@sheet\('S2'\)/g)).toHaveLength(2);
    // The 3,990-character definition is laid out over several lines.
    const long = parseModule(p).entries.find((e) => e.name === "P_Long3990")!;
    expect(long.formula.split("\n").length).toBeGreaterThan(1);
    expect(r.files["names/_unmanaged.xln"]).toContain("Growλ = LAMBDA(b,g, b*(1+g));");
  });

  it("locks every name with a whitespace-insensitive hash", () => {
    const lock = JSON.parse(r.files[LOCK_FILE]!);
    expect(lock.format).toBe("xln.lock/4");
    expect(Object.keys(lock.names)).toHaveLength(10);
    expect(lock.names["P_Multi"].definition).toBe(definitionHash("_xlfn.LAMBDA(_xlpm.x, _xlpm.x*2)"));
    expect(lock.names["P_Add1"].comment).toMatch(/^h:[0-9a-f]{16}$/);
    expect(lock.names["S2!P_Local"]).toMatchObject({ comment: null, hidden: false });
  });
});

describe("f7_base.xlsx manifest", () => {
  const r = pullProject(load("f7_base.xlsx"), "f7_base.xlsx");
  const m = JSON.parse(r.files[MANIFEST_FILE]!);

  it("writes a sheet's own names to names/sheets/<Sheet>.xln", () => {
    expect(r.report.sheetFiles.map((f) => f.file)).toContain("names/sheets/S2.xln");
    expect(m.names["S2!Loc"].file).toBe("names/sheets/S2.xln");
    expect(r.files["names/sheets/S2.xln"]).toMatch(/^(\/\/.*\n)+\n/);
    expect(r.files["names/sheets/S2.xln"]).not.toMatch(/^@(scope|workbook)/m);
  });

  it("lists sheets, the spill map and the 1×1 dynamic-array cells separately", () => {
    expect(m.sheets.map((s: { name: string; state: string }) => [s.name, s.state])).toEqual([
      ["S1", "visible"],
      ["S2", "visible"],
    ]);
    expect(m.spills).toEqual({ S1: [{ anchor: "E1", extent: "E1:E3", rows: 3, cols: 1 }] });
    expect(m.dynamicArrayCells).toEqual({ S1: ["C4", "C7", "C10"] });
    expect(m.names["Spl"].spill).toEqual({ sheet: "S1", anchor: "E1", extent: "E1:E3", spilling: true });
  });

  it("indexes the cells using each name, with sheet scope and LET shadowing", () => {
    // C3 is LET(Rate, 5, Rate*2): its Rate is the LET variable. C2 has "Rate" in a string.
    // C7 is INDIRECT("Rate"): text, not a reference. B3:B11 are shared-formula children.
    expect(m.names["Rate"].usedBy.cells.S1).toEqual(["A1", "C1", "E1", "B2:C2", "B3:B7", "B8:C9", "B10:B11"]);
    expect(m.names["Rate"].usedBy.names).toEqual(["Fn", "Rate2"]);
    expect(m.names["Rate"].usedBy.conditionalFormats).toEqual({ S1: ["A2:A11"] });
    expect(m.names["Rate"].usedBy.dataValidations).toEqual({ S1: ["G1"] });
    // S1!C5 is 'S2'!Loc+Loc: one use of each. On S2, a bare Loc is S2's own.
    expect(m.names["Loc"].usedBy.cells).toEqual({ S1: ["C5"] });
    expect(m.names["S2!Loc"].usedBy.cells).toEqual({ S1: ["C5"], S2: ["A1:B1", "A2:A5"] });
    expect(m.names["Fn"]).toMatchObject({ kind: "lambda", arity: { required: 1, optional: 0 }, uses: ["Rate"] });
  });
});
