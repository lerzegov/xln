// M3b-1, cell statements end to end: pull writes them, the plan compares source,
// lockfile and workbook per statement (E6), the backend writes the cells, read-back
// passes, and the next build has nothing to do. Edits made "in Excel" are simulated by
// applying a change set to the workbook (as a live edit would leave it) or by editing
// the lockfile (as if the last pull had seen another state).
import { describe, expect, it } from "vitest";
import { applyChangeSet, decompile, LOCK_FILE, parseModule, readWorkbook, type BuildResult, type Change } from "../../src/index.js";
import { build, edit, fixture, pulled, why } from "./helpers.js";
import { buildWorkbook } from "../../src/build/build.js";

const F7 = fixture("f7_base.xlsx");
const F8 = fixture("f8_base.xlsx");
const S1 = "names/sheets/S1.xln";
const SLOT = "names/sheets/Slot.xln";

function built(bytes: Uint8Array, files: Record<string, string>, opts = {}): BuildResult {
  const r = buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files }, { embed: false, provenance: false, ...opts });
  expect(r.status, why(r)).toBe("built");
  expect(r.readBack!.ok, r.readBack!.problems.join("\n")).toBe(true);
  const again = build(r.bytes!, { ...files, ...r.files });
  expect(again.status, why(again)).toBe("up-to-date");
  return r;
}

function formulaAt(bytes: Uint8Array, sheet: string, cell: string): string | undefined {
  const s = readWorkbook(bytes).sheets.find((x) => x.name === sheet)!;
  return s.formulas.find((f) => f.cell === cell)?.text;
}

/** The workbook after an edit made in Excel (the same patch a live edit leaves behind). */
function inExcel(bytes: Uint8Array, ...changes: Change[]): Uint8Array {
  return applyChangeSet(bytes, changes);
}

function lockOf(files: Record<string, string>): { format: string; cells: Record<string, { sheet: string; range: string; name?: string; formula: string | null }> } {
  return JSON.parse(files[LOCK_FILE]!);
}

describe("parsing cell statements", () => {
  it("reads named cells, slots, unnamed cells and blocks, with their addresses", () => {
    const text = "@scope(BS)\n/** doc */\nRev @C6 = A1*2;\nSlot @C7 = ;\n@C5 = Model!Years;\n@B40:G40 = SUM(B30:B39);\n@workbook\nG @'SCF x'!$D$9 = 1;\n@hidden\nH = 2;\n";
    const m = parseModule(text);
    expect(m.diagnostics).toEqual([]);
    expect(m.entries.map((e) => [e.name, e.scope, e.cell?.sheet, e.cell?.range, e.formula, e.doc])).toEqual([
      ["Rev", "BS", undefined, "C6", "A1*2", "doc"],
      ["Slot", "BS", undefined, "C7", "", undefined],
      ["G", undefined, "SCF x", "D9", "1", undefined],
      ["H", undefined, undefined, undefined, "2", undefined],
    ]);
    expect(m.cells.map((e) => [e.name, e.cell!.range, e.formula])).toEqual([
      ["Rev", "C6", "A1*2"],
      ["Slot", "C7", ""],
      ["", "C5", "Model!Years"],
      ["", "B40:G40", "SUM(B30:B39)"],
      ["G", "D9", "1"],
    ]);
    const g = m.cells[4]!.cell!;
    expect(text.slice(g.rangeStart, g.end)).toBe("$D$9");
    expect(m.entries[3]!.hidden).toBe(true);
  });

  it("reports what is not a statement", () => {
    expect(parseModule("X @ = 1;").diagnostics.map((d) => d.message)).toEqual(["'X': expected a cell address after '@' (C6, B4:G4, Sheet!C6)"]);
    expect(parseModule("X = ;").diagnostics.map((d) => d.message)).toEqual(["'X': empty definition"]);
    expect(parseModule("/** d */\n@C5 = 1;").diagnostics.map((d) => d.message)).toEqual(["an unnamed cell statement takes no doc comment or annotation"]);
  });
});

describe("pull → edit → build, per cell statement", () => {
  it("a pulled project builds as up to date, every statement in sync", () => {
    const files = pulled(F7);
    const r = build(F7, files);
    expect(r.status, why(r)).toBe("up-to-date");
    expect(r.plan.unchangedCells).toBe(parseModule(files[S1]!).cells.length + parseModule(files["names/sheets/S2.xln"]!).cells.length);
  });

  it("one edited statement: exactly one set-cell-formula, written and read back", () => {
    const files = pulled(F7);
    edit(files, S1, "@C8 = Rate*3;", "@C8 = Rate*4 + LET(x, 1, x);");
    const r = built(F7, files);
    expect(r.plan.changeSet.changes).toEqual([
      { op: "set-cell-formula", sheet: "S1", range: "C8", stored: "Rate*4 + _xlfn.LET(_xlpm.x, 1, _xlpm.x)", display: "Rate*4 + LET(x, 1, x)", previous: "Rate*3" },
    ]);
    expect(formulaAt(r.bytes!, "S1", "C8")).toBe("Rate*4 + _xlfn.LET(_xlpm.x, 1, _xlpm.x)");
    expect(lockOf(r.files!).cells["S1!C8"]!.formula).not.toBe(lockOf(files).cells["S1!C8"]!.formula);
  });

  it("a block is one statement: one change over its range, each cell filled", () => {
    const files = pulled(F7);
    edit(files, S1, "@B2:B11 = A2*Rate;", "@B2:B11 = A2*Rate*2;");
    const r = built(F7, files);
    expect(r.plan.changeSet.changes).toMatchObject([{ op: "set-cell-formula", sheet: "S1", range: "B2:B11", stored: "A2*Rate*2" }]);
    expect(decompile(readWorkbook(r.bytes!).sheets[0]!.formulas.find((f) => f.cell === "B2")!.text!)).toBe("A2*Rate*2");
  });

  it("a named cell on its spill (C6#): only the formula changes", () => {
    const files = pulled(F8);
    edit(files, "names/sheets/N.xln", "@B2 = A1*10;", "@B2 = A1*11;");
    edit(files, "names/sheets/D.xln", "Spill @C6# = SEQUENCE(3);", "Spill @C6# = SEQUENCE(5);");
    const r = built(F8, files);
    expect(r.plan.changeSet.changes.map((c) => c.op)).toEqual(["set-cell-formula", "set-cell-formula"]);
  });

  it("filling a slot with '#': the cell gets its formula and the name becomes 'S'!$B$1#", () => {
    const files = pulled(F8);
    expect(files[SLOT]).toContain("@workbook\nRevenue @B1 = ;\n@workbook\nCosts @B2 = ;");
    edit(files, SLOT, "Revenue @B1 = ;", "Revenue @B1# = SEQUENCE(1, 3) * 10;");
    const r = built(F8, files);
    expect(r.plan.changeSet.changes).toEqual([
      { op: "set-name", name: "Revenue", scope: null, stored: "_xlfn.ANCHORARRAY(Slot!$B$1)", display: "Slot!$B$1#", comment: null, hidden: false, fields: ["definition"] },
      { op: "set-cell-formula", sheet: "Slot", range: "B1", stored: "_xlfn.SEQUENCE(1, 3) * 10", display: "SEQUENCE(1, 3) * 10", name: "Revenue" },
    ]);
    expect(readWorkbook(r.bytes!).definedNames.find((d) => d.name === "Revenue")!.definition).toBe("_xlfn.ANCHORARRAY(Slot!$B$1)");
  });

  it("= ; clears a named cell", () => {
    const files = pulled(F7);
    edit(files, S1, "Spl @E1# = SEQUENCE(3)*Rate;", "Spl @E1# = ;");
    const r = build(F7, files);
    expect(r.plan.changeSet.changes).toEqual([{ op: "clear-cell-formula", sheet: "S1", range: "E1", name: "Spl", previous: "_xlfn.SEQUENCE(3)*Rate" }]);
  });

  it("a removed named statement refuses the build: clear with = ; or delete the name in Excel", () => {
    const files = pulled(F7);
    edit(files, S1, "Spl @E1# = SEQUENCE(3)*Rate;\n", "");
    const r = build(F7, files);
    expect(r.status).toBe("refused");
    expect(r.plan.problems.map((p) => p.code)).toContain("statement-removed");
    expect(r.plan.changeSet.changes.filter((c) => c.op === "delete-name")).toEqual([]);
  });

  it("an edited address refuses the build, with the pulled one as the fix", () => {
    const files = pulled(F7);
    edit(files, S1, "Spl @E1# =", "Spl @E4# =");
    edit(files, S1, "@C8 = Rate*3;", "@C12 = Rate*3;");
    const r = build(F7, files);
    expect(r.status).toBe("refused");
    // The checker's errors, as the editor shows them (M3d).
    const p = r.plan.problems.filter((x) => x.code === "address");
    expect(p.map((x) => x.message)).toEqual([
      "Spl: the address is set in Excel and read-only; the last pull had @E1",
      "no cell statement at S1!C12 in the last pull: addresses are set in Excel (write the formula in Excel, then pull)",
    ]);
    expect(files[S1]!.slice(p[0]!.fix!.start, p[0]!.fix!.end)).toBe("E4");
    expect(p[0]!.fix!.text).toBe("E1");
  });

  it("a cell statement in a module file without its sheet is refused", () => {
    const files = pulled(F7);
    files["names/_unmanaged.xln"] += "\n@C8 = Rate*5;\n";
    expect(build(F7, files).plan.problems.map((x) => x.message)).toContain("a cell statement outside a sheet file names its sheet: @Sheet!C8 (or write it in names/sheets/<Sheet>.xln as @C8)");
  });
});

describe("E6: the cell-level three-way check", () => {
  const excelEdit: Change = { op: "set-cell-formula", sheet: "S1", range: "C8", stored: "Rate*30", display: "Rate*30" };

  it("changed in Excel only: Excel's formula is kept and reported", () => {
    const files = pulled(F7);
    const r = build(inExcel(F7, excelEdit), files);
    expect(r.status, why(r)).toBe("up-to-date");
    expect(r.plan.excelChanges.map((x) => x.message)).toEqual(["S1!C8: the formula was changed in Excel since the last pull; the build keeps Excel's (pull again to update the source)"]);
  });

  it("changed in both, differently: a conflict, never merged, with both sides", () => {
    const files = pulled(F7);
    edit(files, S1, "@C8 = Rate*3;", "@C8 = Rate*4;");
    const r = build(inExcel(F7, excelEdit), files);
    expect(r.status).toBe("refused");
    expect(r.plan.conflicts).toMatchObject([
      { kind: "both-changed", key: "S1!C8", excel: { name: "", display: "Rate*30", cell: { sheet: "S1", range: "C8" } }, source: { display: "Rate*4", file: S1 } },
    ]);
  });

  it("changed in both, the same way: no conflict, nothing to do", () => {
    const files = pulled(F7);
    edit(files, S1, "@C8 = Rate*3;", "@C8 = Rate * 30;");
    expect(build(inExcel(F7, excelEdit), files).status).toBe("up-to-date");
  });

  it("a slot filled in Excel and in the source: a conflict", () => {
    const files = pulled(F8);
    edit(files, SLOT, "Costs @B2 = ;", "Costs @B2 = 5;");
    const r = build(inExcel(F8, { op: "set-cell-formula", sheet: "Slot", range: "B2", stored: "7", display: "7" }), files);
    expect(r.plan.conflicts.map((c) => [c.kind, c.key])).toEqual([["created-in-both", "Costs"]]);
  });

  it("a named cell moved in Excel: the statement follows its name", () => {
    // The last pull saw Spl at E2 (rows were inserted above it since): source and lockfile say E2.
    const files = pulled(F7);
    const lock = lockOf(files);
    lock.cells["Spl"]!.range = "E2";
    files[LOCK_FILE] = JSON.stringify(lock);
    edit(files, S1, "Spl @E1# = SEQUENCE(3)*Rate;", "Spl @E2# = SEQUENCE(4)*Rate;");
    const r = build(F7, files);
    expect(r.plan.problems.filter((p) => p.severity === "error")).toEqual([]);
    expect(r.plan.changeSet.changes).toMatchObject([{ op: "set-cell-formula", sheet: "S1", range: "E1", name: "Spl" }]);
    expect(r.plan.excelChanges.map((x) => x.kind)).toEqual(["moved"]);
  });

  it("an unnamed cell moved in Excel: one statement emptied, one cell added", () => {
    // The last pull saw C10's formula at C12; the source says so too.
    const files = pulled(F7);
    const lock = lockOf(files);
    const e = lock.cells["S1!C10"]!;
    delete lock.cells["S1!C10"];
    lock.cells["S1!C12"] = { ...e, range: "C12" };
    files[LOCK_FILE] = JSON.stringify(lock);
    edit(files, S1, "@C10 = ", "@C12 = ");
    const r = build(F7, files);
    expect(r.plan.changeSet.changes).toEqual([]);
    expect(r.plan.excelChanges.map((x) => [x.kind, x.message])).toEqual([
      ["changed", "S1!C12: the formula was changed in Excel since the last pull (the cell is now empty); the build keeps Excel's (pull again to update the source)"],
      ["created", "S1: formula cells added in Excel since the last pull (C10); the build leaves them (pull again to bring them into the source)"],
    ]);
  });

  it("a format-1 lockfile (no cells) still reads: the workbook is the base", () => {
    const files = pulled(F7);
    const lock = lockOf(files) as unknown as Record<string, unknown>;
    delete lock["cells"];
    lock["format"] = "xln.lock/1";
    files[LOCK_FILE] = JSON.stringify(lock);
    expect(build(F7, files).status).toBe("up-to-date");
    edit(files, S1, "@C8 = Rate*3;", "@C8 = Rate*4;");
    const r = built(F7, files);
    expect(r.plan.changeSet.changes).toHaveLength(1);
    expect(lockOf(r.files!).format).toBe("xln.lock/4");
  });
});

describe("M3d: a scope changes only where the source changes it", () => {
  it("builds workbook names on a sheet's cells as they are; removing @workbook is one rescope-name, the source untouched", () => {
    const files = pulled(F8);
    const asIs = build(F8, files);
    expect(asIs.status, why(asIs)).toBe("up-to-date");
    edit(files, SLOT, "@workbook\nRevenue @B1 = ;", "Revenue @B1 = ;");
    const r = build(F8, files);
    expect(r.status, why(r)).toBe("built");
    expect(r.plan.changeSet.changes).toEqual([{ op: "rescope-name", name: "Revenue", from: null, to: "Slot" }]);
    expect(Object.keys(r.files!).sort()).toEqual(["workbook.manifest.json", LOCK_FILE]);
    expect(readWorkbook(r.bytes!).definedNames.find((d) => d.name === "Revenue")!.scope).toMatchObject({ kind: "sheet", name: "Slot" });
    expect(build(r.bytes!, { ...files, ...r.files }).status).toBe("up-to-date");
  });

  it('an old config with "names": {"scope": "excel"} changes nothing: the same edit is the same rescope-name (the setting is gone, 2026-10-06)', () => {
    const files = pulled(F8);
    files["xln.config.json"] = JSON.stringify({ names: { scope: "excel" } });
    edit(files, SLOT, "@workbook\nRevenue @B1 = ;", "Revenue @B1 = ;");
    const r = build(F8, files);
    expect(r.status, why(r)).toBe("built");
    expect(r.plan.changeSet.changes).toEqual([{ op: "rescope-name", name: "Revenue", from: null, to: "Slot" }]);
  });

  it("adding @workbook to a local name on a cell is one rescope-name the other way", () => {
    const files = pulled(F8);
    edit(files, SLOT, "@workbook\nRevenue @B1 = ;", "Revenue @B1 = ;");
    const r = build(F8, files);
    const back = { ...files, ...r.files };
    edit(back, SLOT, "Revenue @B1 = ;", "@workbook\nRevenue @B1 = ;");
    const r2 = build(r.bytes!, back);
    expect(r2.status, why(r2)).toBe("built");
    expect(r2.plan.changeSet.changes).toEqual([{ op: "rescope-name", name: "Revenue", from: "Slot", to: null }]);
    expect(readWorkbook(r2.bytes!).definedNames.find((d) => d.name === "Revenue")!.scope).toMatchObject({ kind: "workbook" });
  });
});
