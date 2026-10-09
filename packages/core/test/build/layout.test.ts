// Cells keep the workbook's layout (author's decision, 2026-10-09): pull lays a long cell
// formula out on several lines in the source, but a build writes it on one line when the
// cell's formula was on one line in the workbook (and for a new formula, a slot). A formula
// laid out by hand in Excel keeps the source's layout. Defined names are not concerned.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { applyChangeSet, buildReportLines, describeChange, equalModuloWhitespace, hasLayoutBreaks, ONE_LINE_AS_WORKBOOK, ONE_LINE_NEW, readWorkbook, unbuiltEdits, type BuildResult, type SetCellFormula } from "../../src/index.js";
import { buildWorkbook } from "../../src/build/build.js";
import { build, edit, fixture, pulled, why } from "./helpers.js";

const F8 = fixture("f8_base.xlsx");
const CORPUS = process.env["XLN_CORPUS"];
const LBO = CORPUS ? join(CORPUS, "lbo-ep03r", "dist", "lbo-ep03r.xlsx") : undefined;
const N = "names/sheets/N.xln";
const SLOT = "names/sheets/Slot.xln";

const MULTI = 'LET(\n    x, A1*10,\n    s, "a\nb",\n    x + LEN(s)\n)';
const FLAT = 'LET(x, A1*10, s, "a\nb", x + LEN(s))';

function built(bytes: Uint8Array, files: Record<string, string>): BuildResult {
  const r = buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files }, { embed: false, provenance: false });
  expect(r.status, why(r)).toBe("built");
  expect(r.readBack!.ok, r.readBack!.problems.join("\n")).toBe(true);
  return r;
}

function cellText(bytes: Uint8Array, sheet: string, cell: string): string | undefined {
  return readWorkbook(bytes).sheets.find((s) => s.name === sheet)!.formulas.find((f) => f.cell === cell)?.text;
}

/** No edit is left: build → plan, pull → build and a second pull see nothing to do. */
function settled(after: Uint8Array, files: Record<string, string>): void {
  expect(build(after, files).status).toBe("up-to-date");
  expect(unbuiltEdits({ workbook: after, fileName: "book.xlsx", files })).toEqual([]);
  const again = pulled(after);
  expect(build(after, again).status).toBe("up-to-date");
  expect(pulled(after)).toEqual(again);
}

describe("cells keep the workbook's layout", () => {
  it("a one-line formula edited into several lines is written on one line; strings keep their line breaks", () => {
    const files = pulled(F8);
    edit(files, N, "@B2 = A1*10;", `@B2 = ${MULTI};`);
    const r = built(F8, files);
    const ch = r.plan.changeSet.changes.find((c): c is SetCellFormula => c.op === "set-cell-formula" && c.range === "B2")!;
    expect(ch.layout).toBe(ONE_LINE_AS_WORKBOOK);
    expect(ch.display).toBe(FLAT);
    expect(ch.stored).toBe('_xlfn.LET(_xlpm.x, A1*10, _xlpm.s, "a\nb", _xlpm.x + LEN(_xlpm.s))');
    expect(hasLayoutBreaks(cellText(r.bytes!, "N", "B2")!)).toBe(false);
    expect(cellText(r.bytes!, "N", "B2")).toContain('"a\nb"');
    expect(describeChange(ch)).toBe("set formula of N!B2 [on one line, as in the workbook]");
    expect(buildReportLines(r).join("\n")).toContain("[on one line, as in the workbook]");
    // The source keeps its layout; nothing is left to build or pull.
    expect(files[N]).toContain(MULTI);
    settled(r.bytes!, { ...files, ...r.files });
  });

  it("the one-line text compiles to what the source compiles to, modulo whitespace", () => {
    const files = pulled(F8);
    edit(files, N, "@B1 = SUM(A1:A3);", "@B1 = XLOOKUP(\n    2,\n    A1:A3,\n    A1:A3\n) * 7;");
    const r = built(F8, files);
    const ch = r.plan.changeSet.changes.find((c): c is SetCellFormula => c.op === "set-cell-formula" && c.range === "B1")!;
    expect(ch.stored).toBe("_xlfn.XLOOKUP(2, A1:A3, A1:A3) * 7");
    expect(equalModuloWhitespace(ch.display, "XLOOKUP(\n    2,\n    A1:A3,\n    A1:A3\n) * 7")).toBe(true);
  });

  it("a formula laid out on several lines in Excel keeps the source's layout", () => {
    // As if typed in Excel with Alt+Enter: line breaks stored as CR LF.
    const handLaid = applyChangeSet(F8, [{ op: "set-cell-formula", sheet: "N", range: "B2", stored: "IF(A1>0,\r\n  A1*10,\r\n  0)", display: "IF(A1>0,\n  A1*10,\n  0)" }]);
    const files = pulled(handLaid);
    expect(files[N]).toContain("@B2 = IF(A1>0,\n  A1*10,\n  0);");
    edit(files, N, "A1*10,", "A1*20,");
    const r = built(handLaid, files);
    const ch = r.plan.changeSet.changes.find((c): c is SetCellFormula => c.op === "set-cell-formula" && c.range === "B2")!;
    expect(ch.layout).toBeUndefined();
    expect(ch.display).toBe("IF(A1>0,\n  A1*20,\n  0)");
    expect(hasLayoutBreaks(cellText(r.bytes!, "N", "B2")!)).toBe(true);
    expect(describeChange(ch)).toBe("set formula of N!B2");
    settled(r.bytes!, { ...files, ...r.files });
  });

  it("a slot gets its new formula on one line", () => {
    const files = pulled(F8);
    edit(files, SLOT, "Revenue @B1 = ;", "Revenue @B1 = SUM(\n    N!A1:A3\n);");
    const r = built(F8, files);
    const ch = r.plan.changeSet.changes.find((c): c is SetCellFormula => c.op === "set-cell-formula" && c.name === "Revenue")!;
    expect(ch.layout).toBe(ONE_LINE_NEW);
    expect(ch.previous).toBeUndefined();
    expect(cellText(r.bytes!, "Slot", "B1")).toBe("SUM(N!A1:A3)");
    expect(describeChange(ch)).toBe("fill Slot!B1 (Revenue) [on one line]");
    settled(r.bytes!, { ...files, ...r.files });
  });

  it("a source already on one line has no layout note", () => {
    const files = pulled(F8);
    edit(files, N, "@B2 = A1*10;", "@B2 = A1*11;");
    const r = built(F8, files);
    const ch = r.plan.changeSet.changes.find((c): c is SetCellFormula => c.op === "set-cell-formula")!;
    expect(ch.layout).toBeUndefined();
    expect(ch.display).toBe("A1*11");
  });

  it.skipIf(!LBO)("lbo-ep03r (XLN_CORPUS): the MCP trial's FN.LAG edit writes the FIXPOINT cells on one line", () => {
    const bytes = new Uint8Array(readFileSync(LBO!));
    const files = pulled(bytes, "lbo-ep03r.xlsx");
    const lag = /FN\.SEEDROW\(0, FN\.PREV\(([A-Za-z_]+)\)\)/g;
    for (const p of ["names/sheets/SCF recursive.xln", "names/sheets/SCF.xln"]) files[p] = files[p]!.replace(lag, "FN.LAG($1, 0)");
    edit(files, "names/FN.xln", "FN.MAXDEV = ", "FN.LAG = LAMBDA(row, seed, FN.SEEDROW(seed, FN.PREV(row)) );\nFN.MAXDEV = ");
    const r = buildWorkbook({ workbook: bytes, fileName: "lbo-ep03r.xlsx", files }, { embed: false, provenance: false });
    expect(r.status, why(r)).toBe("built");
    expect(r.readBack!.ok, r.readBack!.problems.join("\n")).toBe(true);
    const cells = r.plan.changeSet.changes.filter((c): c is SetCellFormula => c.op === "set-cell-formula");
    expect(cells).toHaveLength(15);
    expect(cells.filter((c) => c.layout !== undefined).map((c) => [c.range, c.layout])).toEqual([
      ["C11", ONE_LINE_AS_WORKBOOK],
      ["C47", ONE_LINE_AS_WORKBOOK],
      ["C83", ONE_LINE_AS_WORKBOOK],
    ]);
    for (const c of ["C11", "C47", "C83"]) expect(hasLayoutBreaks(cellText(r.bytes!, "SCF recursive", c)!)).toBe(false);
    const after = { ...files, ...r.files };
    expect(buildWorkbook({ workbook: r.bytes!, fileName: "lbo-ep03r.xlsx", files: after }, { embed: false, provenance: false }).status).toBe("up-to-date");
    const again = pulled(r.bytes!, "lbo-ep03r.xlsx");
    expect(buildWorkbook({ workbook: r.bytes!, fileName: "lbo-ep03r.xlsx", files: again }, { embed: false, provenance: false }).status).toBe("up-to-date");
    expect(pulled(r.bytes!, "lbo-ep03r.xlsx")).toEqual(again);
  });

  it("defined names keep the source's layout", () => {
    const files = pulled(F8);
    files["names/_unmanaged.xln"] = (files["names/_unmanaged.xln"] ?? "") + "\nTwice = LAMBDA(x,\n    x * 2\n);\n";
    const r = built(F8, files);
    const set = r.plan.changeSet.changes.find((c) => c.op === "set-name" && c.name === "Twice")!;
    expect(set.op === "set-name" && hasLayoutBreaks(set.stored)).toBe(true);
    expect(readWorkbook(r.bytes!).definedNames.find((d) => d.name === "Twice")!.definition).toContain("\r\n");
  });
});
