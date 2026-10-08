// The `#` of a named cell statement (author's decision, 2026-10-05): `Name @C3# = f;` puts
// the name on the cell's spill, `Name @C3 = f;` on the cell alone. Pull writes `#` exactly
// when the name has it; adding or removing it is one `set-name`, checked three-way; a
// statement without `#` on a formula that spills gets a warning with a quick fix; nothing
// moves a name to `#` on its own; a project pulled before the `#` was written still builds
// as up to date, and a pull writes the `#` in.
import { describe, expect, it } from "vitest";
import { applyChangeSet, formatCellAddress, formatEntry, LOCK_FILE, parseModule, pullProject, readWorkbook, replaceZipEntries, scanCellTarget, unbuiltEdits, utf8, type BuildResult, type Change } from "../../src/index.js";
import { Package } from "../../src/file/package.js";
import { buildWorkbook } from "../../src/build/build.js";
import { build, edit, fixture, pulled, why } from "./helpers.js";

const F7 = fixture("f7_base.xlsx");
const F8 = fixture("f8_base.xlsx");
const S1 = "names/sheets/S1.xln";
const D = "names/sheets/D.xln";
const SLOT = "names/sheets/Slot.xln";

function built(bytes: Uint8Array, files: Record<string, string>): BuildResult {
  const r = buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files }, { embed: false, provenance: false });
  expect(r.status, why(r)).toBe("built");
  expect(r.readBack!.ok, r.readBack!.problems.join("\n")).toBe(true);
  const again = build(r.bytes!, { ...files, ...r.files });
  expect(again.status, why(again)).toBe("up-to-date");
  return r;
}

function definition(bytes: Uint8Array, name: string): string {
  return readWorkbook(bytes).definedNames.find((d) => d.name === name)!.definition;
}

/**
 * f8_base with Slot!B1 holding SEQUENCE(1,3,10) spilled over B1:D1 and the name Revenue
 * left on B1: the author's case (a formula written before the build moved names, or typed
 * in Excel). Excel's save records the extent; the backend writes the anchor alone.
 */
function spilledOnCell(): Uint8Array {
  const typed = applyChangeSet(F8, [{ op: "set-cell-formula", sheet: "Slot", range: "B1", stored: "_xlfn.SEQUENCE(1,3,10)", display: "SEQUENCE(1,3,10)", name: "Revenue" }]);
  const part = readWorkbook(typed).sheets.find((x) => x.name === "Slot")!.part!;
  const xml = new Package(typed).text(part)!;
  expect(xml).toContain('<f t="array" ref="B1">');
  return replaceZipEntries(typed, new Map([[part, utf8(xml.replace('<f t="array" ref="B1">', '<f t="array" ref="B1:D1">'))]]));
}

const setName = (name: string, stored: string, comment: string | null = null): Change => ({ op: "set-name", name, scope: null, stored, display: stored, comment, hidden: false, fields: ["definition"] });

describe("the # in a cell address: parser and formatter", () => {
  it("reads C3#, Sheet!C3# and 'S 1'!C3# on named statements", () => {
    const text = "@scope(IS)\nSales @C3# = SEQUENCE(1,5);\nCOGS @C4 = Sales*0.6;\nSlot @C9# = ;\n@workbook\nA @IS!$D$3# = 1;\nB @'S 1'!C3# = 2;\n";
    const m = parseModule(text);
    expect(m.diagnostics).toEqual([]);
    expect(m.cells.map((e) => [e.name, e.cell!.sheet, e.cell!.range, e.cell!.spill, e.formula])).toEqual([
      ["Sales", undefined, "C3", true, "SEQUENCE(1,5)"],
      ["COGS", undefined, "C4", false, "Sales*0.6"],
      ["Slot", undefined, "C9", true, ""],
      ["A", "IS", "D3", true, "1"],
      ["B", "S 1", "C3", true, "2"],
    ]);
    const a = m.cells[3]!.cell!;
    expect(text.slice(a.rangeStart, a.end)).toBe("$D$3#");
    expect(text.slice(a.start, a.end)).toBe("@IS!$D$3#");
    const t = scanCellTarget("@C3#=1", 0)!;
    expect([t.range, t.spill, t.end]).toEqual(["C3", true, 4]);
  });

  it("refuses # on an unnamed statement and on a range", () => {
    expect(parseModule("@C5# = 1;").diagnostics.map((d) => d.message)).toEqual(["@C5#: only a named cell statement takes '#' (it says what the name covers)"]);
    expect(parseModule("X @B4:G4# = 1;").diagnostics.map((d) => d.message)).toEqual(["'X': '#' follows one cell (the spill's anchor), not a range"]);
  });

  it("writes # back", () => {
    expect(formatCellAddress("C3", undefined, true)).toBe("C3#");
    expect(formatCellAddress("C3", "S 1", true)).toBe("'S 1'!C3#");
    expect(formatCellAddress("C3", "IS")).toBe("IS!C3");
    expect(formatEntry({ name: "Sales", cell: "C3#", formula: "SEQUENCE(1,5)" })).toBe("Sales @C3# = SEQUENCE(1,5);");
    expect(formatEntry({ name: "Slot", cell: "IS!C9#", formula: "" })).toBe("Slot @IS!C9# = ;");
  });
});

describe("pull writes # exactly when the name has it", () => {
  it("names on a spill get #, names on a cell and slots do not", () => {
    const files = pulled(F8);
    expect(files[D]).toContain("Spill @C6# = SEQUENCE(3);");
    expect(files[D]).toContain("Spill2 @F6# = SEQUENCE(4);");
    expect(files[SLOT]).toContain("Revenue @B1 = ;");
    expect(pulled(F7)[S1]).toContain("Spl @E1# = SEQUENCE(3)*Rate;");
  });

  it("a name on the first cell of a spill is written without #, with a warning and a quick fix", () => {
    const wb = spilledOnCell();
    const files = pulled(wb);
    expect(files[SLOT]).toContain("Revenue @B1 = SEQUENCE(1,3,10);");
    const r = build(wb, files);
    expect(r.status, why(r)).toBe("up-to-date");
    const p = r.plan.problems.filter((x) => x.code === "spill-uncovered");
    expect(p.map((x) => [x.severity, x.message])).toEqual([["warning", "Revenue @B1: the formula spills over B1:D1, but Revenue covers only B1 (write @B1# to name the spill)"]]);
    // The fix inserts the # right after the address.
    const fix = p[0]!.fix!;
    const text = files[SLOT]!;
    expect(text.slice(fix.start - "B1".length, fix.end)).toBe("B1");
    expect(fix.text).toBe("#");
  });
});

describe("adding or removing the # is one set-name", () => {
  it("adding #: the name goes on the spill, nothing else changes", () => {
    const wb = spilledOnCell();
    const files = pulled(wb);
    const fix = build(wb, files).plan.problems.find((x) => x.code === "spill-uncovered")!.fix!;
    files[SLOT] = files[SLOT]!.slice(0, fix.start) + fix.text + files[SLOT]!.slice(fix.end);
    expect(files[SLOT]).toContain("Revenue @B1# = SEQUENCE(1,3,10);");
    const r = built(wb, files);
    expect(r.plan.changeSet.changes).toEqual([
      { op: "set-name", name: "Revenue", scope: null, stored: "_xlfn.ANCHORARRAY(Slot!$B$1)", display: "Slot!$B$1#", comment: null, hidden: false, fields: ["definition"] },
    ]);
    expect(r.plan.problems.filter((x) => x.code.startsWith("spill"))).toEqual([]);
    expect(definition(r.bytes!, "Revenue")).toBe("_xlfn.ANCHORARRAY(Slot!$B$1)");
  });

  it("removing #: the name covers the cell alone (and the build warns that the formula spills)", () => {
    const files = pulled(F8);
    edit(files, D, "Spill @C6# =", "Spill @C6 =");
    const r = built(F8, files);
    expect(r.plan.changeSet.changes).toMatchObject([{ op: "set-name", name: "Spill", scope: null, stored: "D!$C$6", fields: ["definition"] }]);
    expect(r.plan.problems.map((x) => x.code)).toEqual(["spill-uncovered"]);
    expect(definition(r.bytes!, "Spill")).toBe("D!$C$6");
  });

  it("filling a slot without #: the cell gets its formula, the name stays on the cell", () => {
    const files = pulled(F8);
    edit(files, SLOT, "Revenue @B1 = ;", "Revenue @B1 = SEQUENCE(1, 3) * 10;");
    const r = built(F8, files);
    expect(r.plan.changeSet.changes).toEqual([
      { op: "set-cell-formula", sheet: "Slot", range: "B1", stored: "_xlfn.SEQUENCE(1, 3) * 10", display: "SEQUENCE(1, 3) * 10", name: "Revenue" },
    ]);
    expect(definition(r.bytes!, "Revenue")).toBe(definition(F8, "Revenue"));
  });

  it("# on a slot left empty: a warning, the name is not written until the cell gets a formula", () => {
    const files = pulled(F8);
    edit(files, SLOT, "Revenue @B1 = ;", "Revenue @B1# = ;");
    const r = build(F8, files);
    expect(r.status, why(r)).toBe("up-to-date");
    expect(r.plan.problems.map((x) => [x.severity, x.code])).toEqual([["warning", "spill-empty"]]);
    expect(r.plan.sourceEdits.spill).toEqual(new Set(["revenue"]));
  });
});

describe("the # three-way: lockfile, workbook, source", () => {
  const onSpill = setName("Revenue", "_xlfn.ANCHORARRAY(Slot!$B$1)");

  it("changed in Excel only: kept and reported", () => {
    const wb = spilledOnCell();
    const files = pulled(wb);
    const r = build(applyChangeSet(wb, [onSpill]), files);
    expect(r.status, why(r)).toBe("up-to-date");
    expect(r.plan.excelChanges.map((x) => [x.kind, x.key])).toEqual([["changed", "Revenue"]]);
  });

  it("added in Excel and in the source: no conflict, nothing to do", () => {
    const wb = spilledOnCell();
    const files = pulled(wb);
    edit(files, SLOT, "Revenue @B1 =", "Revenue @B1# =");
    const r = build(applyChangeSet(wb, [onSpill]), files);
    expect(r.status, why(r)).toBe("up-to-date");
    expect(r.plan.conflicts).toEqual([]);
  });

  it("added in the source, the name's comment changed in Excel: a conflict", () => {
    const wb = spilledOnCell();
    const files = pulled(wb);
    edit(files, SLOT, "Revenue @B1 =", "Revenue @B1# =");
    const r = build(applyChangeSet(wb, [{ ...setName("Revenue", "'Slot'!$B$1", "from Excel"), fields: ["comment"] }]), files);
    expect(r.status).toBe("refused");
    expect(r.plan.conflicts.map((c) => [c.kind, c.key])).toEqual([["both-changed", "Revenue"]]);
  });

  it("removed in the source, the name redefined as a range in Excel: a conflict", () => {
    const files = pulled(F8);
    edit(files, D, "Spill @C6# =", "Spill @C6 =");
    const r = build(applyChangeSet(F8, [setName("Spill", "'D'!$C$6:$C$8")]), files);
    expect(r.status).toBe("refused");
    expect(r.plan.conflicts.map((c) => [c.kind, c.key])).toEqual([["both-changed", "Spill"]]);
  });
});

describe("pull and the # (every pull is fresh)", () => {
  it("Excel put the name on the spill: nothing for the guard; the pull writes the #", () => {
    const wb = spilledOnCell();
    const files = pulled(wb);
    const xl = applyChangeSet(wb, [setName("Revenue", "_xlfn.ANCHORARRAY(Slot!$B$1)")]);
    expect(unbuiltEdits({ workbook: xl, fileName: "book.xlsx", files })).toEqual([]);
    const p = pullProject(xl, "book.xlsx");
    expect(p.files[SLOT]).toBe(files[SLOT]!.replace("Revenue @B1 =", "Revenue @B1# ="));
    expect(build(xl, { ...files, ...p.files }).status).toBe("up-to-date");
  });

  it("a # added in the source and not built yet is an edit the guard lists", () => {
    const wb = spilledOnCell();
    const files = pulled(wb);
    edit(files, SLOT, "Revenue @B1 =", "Revenue @B1# =");
    expect(unbuiltEdits({ workbook: wb, fileName: "book.xlsx", files }).map((e) => [e.key, e.file, e.what])).toEqual([["Revenue", SLOT, "update Revenue (definition)"]]);
  });
});

describe("a project pulled before the # was written (lockfile format 3)", () => {
  /** What the pull of earlier versions wrote: no #, lockfile format 3. */
  function oldProject(bytes: Uint8Array): Record<string, string> {
    const files = pulled(bytes);
    for (const p of Object.keys(files)) if (p.startsWith("names/")) files[p] = files[p]!.split("# =").join(" =");
    files[LOCK_FILE] = files[LOCK_FILE]!.replace('"xln.lock/4"', '"xln.lock/3"');
    return files;
  }

  it("builds as up to date: a missing # on a name on the spill (lockfile and workbook) reads as #", () => {
    const files = oldProject(F8);
    expect(files[D]).toContain("Spill @C6 = SEQUENCE(3);");
    const r = build(F8, files);
    expect(r.status, why(r)).toBe("up-to-date");
    expect(r.plan.implicitSpill).toEqual(["Spill", "Spill2"]);
    expect(r.plan.problems.filter((x) => x.code === "spill-uncovered")).toEqual([]);
  });

  it("a build keeps format 3 while the source relies on it; adding # to a plain name still works", () => {
    const files = oldProject(F8);
    edit(files, "names/sheets/N.xln", "@B2 = A1*10;", "@B2 = A1*11;");
    edit(files, SLOT, "Revenue @B1 = ;", "Revenue @B1# = SEQUENCE(1,3);");
    const r = built(F8, files);
    expect(r.plan.changeSet.changes.map((c) => c.op)).toEqual(["set-name", "set-cell-formula", "set-cell-formula"]);
    expect(JSON.parse(r.files![LOCK_FILE]!).format).toBe("xln.lock/3");
  });

  it("a project without the # is no edit for the pull's guard; the pull writes the # in and the lockfile in format 4", () => {
    const files = oldProject(F8);
    expect(unbuiltEdits({ workbook: F8, fileName: "book.xlsx", files })).toEqual([]);
    const p = pullProject(F8, "book.xlsx");
    expect(p.files[D]).toContain("Spill @C6# = SEQUENCE(3);");
    expect(p.files[D]).toContain("Spill2 @F6# = SEQUENCE(4);");
    expect(JSON.parse(p.files[LOCK_FILE]!).format).toBe("xln.lock/4");
    const r = build(F8, { ...files, ...p.files });
    expect(r.status, why(r)).toBe("up-to-date");
    expect(r.plan.implicitSpill).toEqual([]);
  });

  it("a format-4 lockfile reads a missing # as the cell alone", () => {
    const files = pulled(F8);
    edit(files, D, "Spill @C6# =", "Spill @C6 =");
    expect(build(F8, files).plan.changeSet.changes.map((c) => c.op)).toEqual(["set-name"]);
  });
});
