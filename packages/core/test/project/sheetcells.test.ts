// Sheet-cell names in sheet files (feedback 2026-10-07): a workbook name no module owns whose
// definition is fixed to one sheet's cells is written in that sheet's file with `@workbook`;
// `_unmanaged.xln` keeps the others. Removing `@workbook` makes it local (rescope-name). And
// the pull's note on names Create from Selection took from a computed value or a corner.
import { describe, expect, it } from "vitest";
import {
  checkFile,
  labelName,
  LOCK_FILE,
  parseLockfile,
  pullProject,
  readWorkbook,
  sheetCellsOf,
  SourceModel,
  unbuiltEdits,
  type CheckContext,
  type PullResult,
} from "../../src/index.js";
import { build, edit, fixture, why, withWorkbookXml, workbookXml } from "../build/helpers.js";

const F7 = fixture("f7_base.xlsx");
const S1 = "names/sheets/S1.xln";
const S2 = "names/sheets/S2.xln";
const U = "names/_unmanaged.xln";

/** f7_base with more defined names (as Excel stores them) before its own. */
function withNames(bytes: Uint8Array, names: string): Uint8Array {
  const xml = workbookXml(bytes);
  return withWorkbookXml(bytes, xml.replace("<definedNames>", `<definedNames>${names}`));
}

// S1: A2:A11 values, B2:B11 a filled formula, C2 = "Rate is "&Rate, E1 spills E1:E3. S2: formulas in A1:A5, B1.
const EXTRA = [
  '<definedName name="InSpill">S1!$E$2</definedName>', // a cell inside a spill
  '<definedName name="Blk">S1!$B$2:$B$4</definedName>', // a range of formula cells
  '<definedName name="Val">S1!$A$2</definedName>', // a cell holding a value
  '<definedName name="EmptyRng">S2!$H$20:$H$22</definedName>', // empty cells
  '<definedName name="Rel">S1!A2:A3</definedName>', // relative
  '<definedName name="Both">S1!$A$1,S2!$A$1</definedName>', // a union over two sheets
  '<definedName name="Thru">S1:S2!$A$1</definedName>', // 3-D
  '<definedName name="WholeCol">S1!$A:$A</definedName>', // a whole column
  '<definedName name="Calc">SUM(S1!$A$2:$A$4)</definedName>', // a formula
  '<definedName name="MOD.Cell">S1!$A$3</definedName>', // a module's
  '<definedName name="LocRng" localSheetId="1">S1!$A$2:$A$3</definedName>', // local to S2, on S1's cells
].join("");

const BOOK = withNames(F7, EXTRA);

function pull(bytes: Uint8Array): PullResult {
  return pullProject(bytes, "book.xlsx");
}

function model(files: Record<string, string>): SourceModel {
  const m = new SourceModel();
  for (const [p, t] of Object.entries(files)) if (p.startsWith("names/") && p.endsWith(".xln")) m.setFile(p, t);
  return m;
}

function ctxOf(r: PullResult, files: Record<string, string>): CheckContext {
  return { sheets: r.snapshot.sheets.map((s) => s.name), tables: r.snapshot.tables.map((t) => t.displayName), lock: parseLockfile(files[LOCK_FILE]!) };
}

function codes(files: Record<string, string>, path: string, ctx: CheckContext): [string, string, string][] {
  const m = model(files);
  return checkFile(m, path, ctx).map((p) => [p.severity, p.code ?? "", m.files.get(path)!.text.slice(p.start, p.end)]);
}

describe("sheetCellsOf: a definition fixed to one sheet's cells", () => {
  const sheets = ["S1", "S 2"];
  it("one absolute cell, area or spill, written with its sheet", () => {
    expect(sheetCellsOf("S1!$B$12", sheets)).toBe("S1");
    expect(sheetCellsOf("s1!$B$12:$G$12", sheets)).toBe("S1");
    expect(sheetCellsOf("_xlfn.ANCHORARRAY(S1!$A$12)", sheets)).toBe("S1");
    expect(sheetCellsOf("'S 2'!$A$1", sheets)).toBe("S 2");
    expect(sheetCellsOf("S1!$B$12", undefined)).toBe("S1");
  });
  it("anything else: relative, mixed, unions, 3-D, whole rows or columns, formulas, constants, an unknown sheet, no sheet", () => {
    for (const d of ["S1!B12", "S1!$B12", "S1!B$12:$C$13", "S1!$A$1,S1!$B$1", "S1:S2!$A$1", "S1!$A:$A", "S1!$1:$1", "SUM(S1!$A$1:$A$3)", "42", "Other!$A$1", "$A$1", "[1]S1!$A$1", "S1!$A$1:S1!$B$2"]) {
      expect(sheetCellsOf(d, sheets), d).toBeUndefined();
    }
  });
});

describe("pull: names on one sheet's cells go to the sheet's file", () => {
  const r = pull(BOOK);

  it("cell inside a spill, range, value cell, empty range: in the sheet's file with @workbook, after the sheet's own names", () => {
    expect(r.files[S1]).toContain("\n\n// Other names on S1.\n\n@workbook\nBlk = 'S1'!$B$2:$B$4;\n@workbook\nInSpill = 'S1'!$E$2;\n@workbook\nVal = 'S1'!$A$2;\n");
    expect(r.files[S2]).toContain("\n\n// Other names on S2.\n\nLoc = 7;\nLocRng = 'S1'!$A$2:$A$3;\n@workbook\nEmptyRng = 'S2'!$H$20:$H$22;\n");
  });

  it("_unmanaged.xln keeps the names not tied to one sheet's cells; a module keeps its names", () => {
    const u = r.files[U]!;
    for (const n of ["Both", "Thru", "Rel", "WholeCol", "Calc", "Rate", "Fn"]) expect(u).toMatch(new RegExp(`^${n} = `, "m"));
    for (const n of ["Blk", "InSpill", "Val", "EmptyRng", "MOD.Cell"]) expect(u).not.toMatch(new RegExp(`^${n} = `, "m"));
    expect(r.files["names/MOD.xln"]).toContain("\nMOD.Cell = 'S1'!$A$3;\n");
    expect(u.split("\n").slice(0, 3)).toEqual([
      "// Workbook-scoped names that no module owns and that are not tied to one sheet's cells,",
      "// pulled by xln from book.xlsx. A name on a sheet's cells is in that sheet's file",
      "// (names/sheets/), a name with a module prefix (ANA.GROW) in its module's file.",
    ]);
  });

  it("the report and the manifest place them there; scopes stay Excel's", () => {
    expect(r.report.unmanaged).toBe(10);
    expect(r.report.sheetFiles.map((s) => [s.sheet, s.names])).toEqual([
      ["S1", 4],
      ["S2", 3],
    ]);
    expect(r.names.find((n) => n.name === "Blk")!.scope).toBeUndefined();
    const m = JSON.parse(r.files["workbook.manifest.json"]!);
    const fileOf = (key: string) => (m.names as Record<string, { file: string }>)[key]?.file;
    expect(fileOf("Blk")).toBe(S1);
    expect(fileOf("EmptyRng")).toBe(S2);
    expect(fileOf("Calc")).toBe(U);
  });

  it("a name on empty cells of a sheet: in the sheet's file", () => {
    const only = pull(withNames(F7, '<definedName name="Far">S2!$K$40:$K$41</definedName>'));
    expect(only.files[S2]).toContain("@workbook\nFar = 'S2'!$K$40:$K$41;\n");
  });

  it("the pulled project has no findings on these names (no file-placement, the hint where it applies)", () => {
    const files = { ...r.files };
    const ctx = ctxOf(r, files);
    const all = [S1, S2, U, "names/MOD.xln"].flatMap((p) => codes(files, p, ctx).map((c) => [p, ...c]));
    expect(all.filter((c) => c[2] !== "workbook-on-cell" || !["Blk", "InSpill", "Val", "EmptyRng", "Spl"].includes(c[3]!))).toEqual([]);
  });
});

describe("removing @workbook makes the name local: the build moves it (rescope-name)", () => {
  const r = pull(BOOK);

  it("a workbook range on S1's cells, read nowhere: the hint, its fix, then one rescope-name; the next pull writes it local", () => {
    const files = { ...r.files };
    const found = checkFile(model(files), S1, ctxOf(r, files)).filter((p) => p.code === "workbook-on-cell");
    const blk = found.find((p) => files[S1]!.slice(p.start, p.end) === "Blk")!;
    expect(blk.severity).toBe("hint");
    expect(blk.message).toBe("workbook name on cells of S1, read only on S1: remove @workbook to make it local to S1 (the build then moves it)");
    expect(found.find((p) => files[S1]!.slice(p.start, p.end) === "InSpill")!.message).toMatch(/^workbook name on a cell of S1/);
    const fix = blk.fix!;
    expect(fix.title).toBe("Remove @workbook: make Blk local to S1");
    files[S1] = files[S1]!.slice(0, fix.start) + fix.text + files[S1]!.slice(fix.end);
    // The fix writes it as a pull writes a local name: without its own sheet.
    expect(files[S1]).toContain("\n\n// Other names on S1.\n\nBlk = $B$2:$B$4;\n@workbook\nInSpill");
    expect(codes(files, S1, ctxOf(r, files)).filter((c) => c[2] === "Blk")).toEqual([["info", "rescope", "Blk"]]);

    const b = build(BOOK, files);
    expect(b.status, why(b)).toBe("built");
    expect(b.plan.changeSet.changes).toEqual([{ op: "rescope-name", name: "Blk", from: null, to: "S1" }]);
    const d = readWorkbook(b.bytes!).definedNames.find((n) => n.name === "Blk")!;
    expect([d.scope.kind === "sheet" ? d.scope.name : "workbook", d.definition]).toEqual(["S1", "S1!$B$2:$B$4"]);

    // Pulled again it is local, written as the fix wrote it: the pull rewrites nothing.
    const again = pull(b.bytes!);
    expect(again.files[S1]).toBe(files[S1]);
    expect(build(b.bytes!, again.files).status).toBe("up-to-date");
  });

  it("removing @workbook by hand, the sheet left in the reference: the same rescope-name, read back as equal", () => {
    const files = { ...r.files };
    edit(files, S1, "@workbook\nVal = 'S1'!$A$2;", "Val = 'S1'!$A$2;");
    const b = build(BOOK, files);
    expect(b.status, why(b)).toBe("built");
    expect(b.plan.changeSet.changes).toEqual([{ op: "rescope-name", name: "Val", from: null, to: "S1" }]);
    expect(pull(b.bytes!).files[S1]).toContain("\nVal = $A$2;\n");
  });

  it("read by another sheet unqualified, it needs workbook scope: no hint", () => {
    const files = { ...r.files };
    edit(files, S2, "@A1 = Loc;", "@A1 = Loc + SUM(Blk);");
    expect(codes(files, S1, ctxOf(r, files)).filter((c) => c[2] === "Blk")).toEqual([]);
  });

  it("a workbook name on S1's cells left in _unmanaged.xln: a placement warning whose fix moves it to S1's file with @workbook", () => {
    const files = { ...r.files };
    edit(files, S1, "@workbook\nVal = 'S1'!$A$2;\n", "");
    files[U] += "\nVal = 'S1'!$A$2;\n";
    const found = checkFile(model(files), U, ctxOf(r, files)).filter((p) => p.code === "file-placement");
    expect(found.map((p) => p.message)).toEqual(["Val is on cells of S1: a pull puts it in names/sheets/S1.xln (with @workbook above it); move it there"]);
    const fix = found[0]!.fixes![0]!;
    expect(fix.title).toBe(`Move it to ${S1}`);
    const there = fix.elsewhere![0]!;
    expect(there.path).toBe(S1);
    const moved = files[S1]!.slice(0, there.start) + there.text + files[S1]!.slice(there.end);
    expect(moved.endsWith("\n@workbook\nVal = 'S1'!$A$2;\n")).toBe(true);
  });
});

describe("pull → pull on a project made by the old placement (workbook names on cells in _unmanaged.xln)", () => {
  it("the guard finds no edit: moving the names is the pull's, not the author's", () => {
    const r = pull(BOOK);
    const old = { ...r.files };
    // The old placement: the @workbook names of the "Other names" sections in _unmanaged.xln, without @workbook.
    const moved: string[] = [];
    for (const p of [S1, S2]) {
      old[p] = old[p]!.replace(/@workbook\n(\w+ = [^;]+;)\n/g, (all, stmt: string, offset: number, text: string) => {
        if (text.lastIndexOf("// Other names on", offset) < 0) return all;
        moved.push(stmt);
        return "";
      });
    }
    expect(moved.sort()).toEqual(["Blk = 'S1'!$B$2:$B$4;", "EmptyRng = 'S2'!$H$20:$H$22;", "InSpill = 'S1'!$E$2;", "Val = 'S1'!$A$2;"]);
    old[U] += moved.join("\n") + "\n";
    expect(unbuiltEdits({ workbook: BOOK, fileName: "book.xlsx", files: old })).toEqual([]);
    expect(build(BOOK, old).status).toBe("up-to-date");
    // The pull writes the new placement.
    expect(pull(BOOK).files[S1]).toContain("@workbook\nBlk = 'S1'!$B$2:$B$4;");
  });
});

describe("the value-label note: names Create from Selection took from a computed cell's value or a corner", () => {
  it("labelName: Excel's spelling of a label", () => {
    expect(labelName("Rate is 0.1")).toBe("Rate_is_0.1");
    expect(labelName("  italian10 ")).toBe("italian10");
    expect(labelName(2025)).toBeUndefined(); // a number gives no name (F11)
    expect(labelName("1st year")).toBe("_1st_year");
    expect(labelName(0.5)).toBeUndefined();
    expect(labelName(true)).toBeUndefined();
    expect(labelName("")).toBeUndefined();
  });

  it("after the value of a formula cell just left of a row, or just above a column: a note", () => {
    // C2 = "Rate is "&Rate shows "Rate is 0.1"; D2 and C3:C… are next to it.
    const r = pull(withNames(F7, '<definedName name="Rate_is_0.1">S1!$D$2:$F$2</definedName><definedName name="rate_IS_0.1" localSheetId="1">S1!$C$3:$C$4</definedName>'));
    expect(r.report.valueLabels).toEqual([
      { key: "Rate_is_0.1", kind: "value", cell: "'S1'!C2", range: "'S1'!D2:F2" },
      { key: "S2!rate_IS_0.1", kind: "value", cell: "'S1'!C2", range: "'S1'!C3:C4" },
    ]);
    expect(r.report.notes).toContain("Rate_is_0.1: named after the current value of 'S1'!C2, a formula result; the name stays as is when that value changes");
  });

  it("no note: a value typed in (A2 = 1), a cell not next to the label, a label that differs, a column named from its left", () => {
    const r = pull(
      withNames(
        F7,
        [
          '<definedName name="_1">S1!$B$2:$D$2</definedName>', // A2 holds 1, typed
          '<definedName name="Rate_is_0.1">S1!$E$2:$F$2</definedName>', // not next to C2
          '<definedName name="Rate_is">S1!$D$2:$F$2</definedName>', // another text
          '<definedName name="Rate_is_0.1x">S1!$C$3:$C$4</definedName>',
        ].join(""),
      ),
    );
    expect(r.report.valueLabels).toEqual([]);
  });

  it("the corner of a two-way Create from Selection: a note when another name covers one of the block's rows or columns", () => {
    // C2's text names D3:E4, the block below and right of it; Col = D3:D4 is one of its columns.
    const corner = '<definedName name="Rate_is_0.1">S1!$D$3:$E$4</definedName>';
    const alone = pull(withNames(F7, corner));
    expect(alone.report.valueLabels).toEqual([]);
    const r = pull(withNames(F7, corner + '<definedName name="Col">S1!$D$3:$D$4</definedName>'));
    expect(r.report.valueLabels).toEqual([{ key: "Rate_is_0.1", kind: "corner", cell: "'S1'!C2", range: "'S1'!D3:E4" }]);
    expect(r.report.notes).toContain("Rate_is_0.1: named after 'S1'!C2, the corner of a Create from Selection with both Top row and Left column; it covers the whole block 'S1'!D3:E4");
  });

  it("the probe workbooks: no note", () => {
    for (const f of ["f7_base.xlsx", "f8_dynamic_arrays.xlsx"]) {
      let bytes: Uint8Array;
      try {
        bytes = fixture(f);
      } catch {
        continue;
      }
      expect(pull(bytes).report.valueLabels, f).toEqual([]);
    }
  });
});
