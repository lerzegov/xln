// Trim references and references to other workbooks, as probe F10 measured them (Excel for
// Mac, probes/results/f10_trim_extref_mac.xlsx): `A1.:.A10` is stored
// `_xlfn._TRO_ALL(A1:A10)` (`:.` _TRO_TRAILING, `.:` _TRO_LEADING), in names and cells
// alike; another workbook is `[n]`, the n-th <externalReference>, whose part names the file.
// xln shows `[n]` by the file's name and compiles it back while the link exists.
import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { checkFile, compileWithDiagnostics, decompile, LOCK_FILE, parseLockfile, pullProject, readWorkbook, SourceModel, type CheckContext, type WorkbookLink } from "../../src/index.js";
import { build, edit, fixture, pulled, why } from "./helpers.js";

const F10 = fixture("f10_trim_extref_mac.xlsx");
const LINKS: WorkbookLink[] = [{ index: 1, book: "Other.xlsx" }];

// What Excel stored → what xln shows (the kit's typed text, with Excel's spelling of the link).
const NAMES: [name: string, stored: string, shown: string][] = [
  ["TrimAll", "_xlfn._TRO_ALL(Sheet1!$A$1:$A$10)", "Sheet1!$A$1.:.$A$10"],
  ["TrimEnd", "_xlfn._TRO_TRAILING(Sheet1!$A$1:$A$10)", "Sheet1!$A$1:.$A$10"],
  ["TrimStart", "_xlfn._TRO_LEADING(Sheet1!$A$1:$A$10)", "Sheet1!$A$1.:$A$10"],
  ["TrimFn", "_xlfn.TRIMRANGE(Sheet1!$A$1:$A$10)", "TRIMRANGE(Sheet1!$A$1:$A$10)"],
  ["Ext", "[1]Sheet1!$A$1", "[Other.xlsx]Sheet1!$A$1"],
  ["ExtName", "[1]!OtherVal", "Other.xlsx!OtherVal"],
];
const CELLS: [cell: string, stored: string, shown: string][] = [
  ["C1", "ROWS(_xlfn._TRO_ALL(A1:A10))", "ROWS(A1.:.A10)"],
  ["C2", "ROWS(_xlfn._TRO_TRAILING(A1:A10))", "ROWS(A1:.A10)"],
  ["C3", "ROWS(_xlfn._TRO_LEADING(A1:A10))", "ROWS(A1.:A10)"],
  ["C4", "ROWS(_xlfn.TRIMRANGE(A1:A10))", "ROWS(TRIMRANGE(A1:A10))"],
  ["C5", "SUM(_xlfn._TRO_TRAILING(A:A))", "SUM(A:.A)"],
  ["C6", "[1]Sheet1!$A$1", "[Other.xlsx]Sheet1!$A$1"],
  ["C7", "[1]!OtherVal", "Other.xlsx!OtherVal"],
];

describe("the F10 workbook's link", () => {
  it("reads <externalReferences> and the link part's target", () => {
    // Excel stored the absolute path on the author's Mac: only its end is checked here.
    expect(readWorkbook(F10).externalLinks).toEqual([
      { index: 1, book: "Other.xlsx", target: expect.stringMatching(/^\/.*\/probes\/kits\/f10\/Other\.xlsx$/), part: "xl/externalLinks/externalLink1.xml" },
    ]);
  });
});

describe("round trip on what Excel stored (F10)", () => {
  const wb = readWorkbook(F10);
  it.each(NAMES)("name %s: %s ↔ %s", (name, stored, shown) => {
    expect(wb.definedNames.find((d) => d.name === name)?.definition).toBe(stored);
    expect(decompile(stored, { links: wb.externalLinks })).toBe(shown);
    const c = compileWithDiagnostics(shown, { links: wb.externalLinks, names: wb.definedNames.map((d) => d.name) });
    expect(c.diagnostics).toEqual([]);
    expect(c.text).toBe(stored);
  });
  it.each(CELLS)("cell %s: %s ↔ %s", (cell, stored, shown) => {
    expect(wb.sheets[0]!.formulas.find((f) => f.cell === cell)?.text).toBe(stored);
    expect(decompile(stored, { links: wb.externalLinks })).toBe(shown);
    expect(compileWithDiagnostics(shown, { links: wb.externalLinks }).text).toBe(stored);
  });

  it("the pull writes the trim references and the file's name", () => {
    const files = pulled(F10);
    for (const [name, , shown] of NAMES) expect(files["names/_unmanaged.xln"]).toContain(`${name} = ${shown};`);
    for (const [cell, , shown] of CELLS) expect(files["names/sheets/Sheet1.xln"]).toContain(`@${cell} = ${shown};`);
    expect(JSON.parse(files["workbook.manifest.json"]!).externalLinks).toEqual([{ index: 1, book: "Other.xlsx" }]);
  });
});

describe("trim references in other places", () => {
  const rt = (shown: string, stored: string, home?: string) => {
    const ctx = home === undefined ? {} : { homeSheet: home };
    expect(compileWithDiagnostics(shown, ctx).text).toBe(stored);
    expect(decompile(stored, ctx)).toBe(shown);
  };
  it("whole rows and columns, a home sheet, inside @ and functions", () => {
    rt("SUM(1:.1)", "SUM(_xlfn._TRO_TRAILING(1:1))");
    rt("A:.A", "_xlfn._TRO_TRAILING(S!A:A)", "S");
    rt("'My Sheet'!B2.:.C9", "_xlfn._TRO_ALL('My Sheet'!B2:C9)");
    rt("@A1.:.A3", "_xlfn.SINGLE(_xlfn._TRO_ALL(A1:A3))");
    rt("[Other.xlsx]Sheet1!A1.:.A3", "_xlfn._TRO_ALL([Other.xlsx]Sheet1!A1:A3)"); // no links known: as written
  });
  it("a trim function Excel would not show as an operator stays a call", () => {
    for (const s of ["_xlfn._TRO_ALL(A1)", "_xlfn._TRO_ALL(A1:INDEX(B:B,3))", "B1:_xlfn._TRO_ALL(A1:A10)"]) {
      expect(compileWithDiagnostics(decompile(s), {}).text).toBe(s);
    }
  });
  it("a trim operator between operands that are not one range: the same function around the range", () => {
    expect(compileWithDiagnostics("A1:.INDEX(B:B,3)", {}).text).toBe("_xlfn._TRO_TRAILING(A1:INDEX(B:B,3))");
  });
});

describe("references to other workbooks", () => {
  const c = (shown: string, links: readonly WorkbookLink[] | null = LINKS) => compileWithDiagnostics(shown, links === null ? {} : { links });
  it("by the file's name, quoted or not, with or without a folder, case ignored", () => {
    expect(c("'[Other.xlsx]Sheet1'!$A$1").text).toBe("[1]Sheet1!$A$1");
    expect(c("'C:\\dir\\[Other.xlsx]Sheet1'!A1").text).toBe("[1]Sheet1!A1");
    expect(c("other.XLSX!OtherVal").text).toBe("[1]!OtherVal");
    expect(c("Other.xlsx!Fn(2)").text).toBe("[1]!Fn(2)");
    expect(c("'[Other.xlsx]My Sheet'!A1").text).toBe("'[1]My Sheet'!A1");
    expect(c("[1]Sheet1!A1").text).toBe("[1]Sheet1!A1");
  });
  it("shown as Excel's formula bar shows them: quoted only when the file or the sheet needs it", () => {
    const links: WorkbookLink[] = [{ index: 1, book: "Other.xlsx" }, { index: 2, book: "My Book.xlsx" }];
    expect(decompile("'[1]My Sheet'!A1", { links })).toBe("'[Other.xlsx]My Sheet'!A1");
    expect(decompile("[2]Sheet1!A1+[2]!Rate", { links })).toBe("'[My Book.xlsx]Sheet1'!A1+'My Book.xlsx'!Rate");
    expect(c("'[My Book.xlsx]Sheet1'!A1+'My Book.xlsx'!Rate", links).text).toBe("[2]Sheet1!A1+[2]!Rate");
  });
  it("two links to files of the same name keep their numbers", () => {
    const links: WorkbookLink[] = [{ index: 1, book: "Other.xlsx" }, { index: 2, book: "other.xlsx" }];
    expect(decompile("[2]Sheet1!A1", { links })).toBe("[2]Sheet1!A1");
  });
  it("an error when the workbook has no link to the file: Excel writes the link, not xln", () => {
    for (const [f, book] of [["[Missing.xlsx]Sheet1!A1", "Missing.xlsx"], ["Missing.xlsx!Rate", "Missing.xlsx"], ["[3]Sheet1!A1", "[3]"]] as const) {
      const r = c(f);
      expect(r.text).toBe("");
      expect(r.diagnostics.map((d) => [d.severity, d.code])).toEqual([["error", "external-link"]]);
      expect(r.diagnostics[0]!.message).toBe(`this workbook has no link to ${book} (its links: [1] Other.xlsx): Excel writes a link when a formula first names the other workbook, xln does not; type the reference once in Excel, save, then pull`);
    }
    expect(c("[Other.xlsx]Sheet1!A1", []).diagnostics[0]!.message).toContain("(it has none)");
  });
  it("without links (no workbook known) kept as written", () => {
    expect(c("[Missing.xlsx]Sheet1!A1", null)).toEqual({ text: "[Missing.xlsx]Sheet1!A1", diagnostics: [] });
    expect(c("[0]!ChartName").text).toBe("[0]!ChartName");
  });
});

describe("the checker", () => {
  const problems = (text: string, links: readonly WorkbookLink[] | undefined) => {
    const m = new SourceModel();
    m.setFile("names/_unmanaged.xln", text);
    const ctx: CheckContext = { sheets: ["Sheet1"], ...(links ? { links } : {}) };
    return checkFile(m, "names/_unmanaged.xln", ctx).map((p) => [p.severity, p.code, text.slice(p.start, p.end)]);
  };
  it("says nothing on trim references and linked workbooks", () => {
    expect(problems("A = Sheet1!$A$1.:.$A$10;\nB = SUM(Sheet1!A:.A);\nC = [Other.xlsx]Sheet1!$A$1;\nD = Other.xlsx!OtherVal;\n", LINKS)).toEqual([]);
  });
  it("an error on a workbook the file has no link to", () => {
    expect(problems("E = [Missing.xlsx]Sheet1!$A$1 + 1;\n", LINKS)).toEqual([["error", "external-link", "[Missing.xlsx]Sheet1!$A$1"]]);
  });
  it("a definition as the workbook has it since the last pull stays quiet", () => {
    const r = pullProject(F10, "f10.xlsx");
    const ctx: CheckContext = { sheets: ["Sheet1"], lock: parseLockfile(r.files[LOCK_FILE]!), links: r.snapshot.externalLinks };
    const m = new SourceModel();
    for (const [p, t] of Object.entries(r.files)) m.setFile(p, t);
    expect(checkFile(m, "names/_unmanaged.xln", ctx).filter((p) => p.severity !== "hint")).toEqual([]);
    expect(checkFile(m, "names/sheets/Sheet1.xln", ctx).filter((p) => p.severity !== "hint")).toEqual([]);
  });
});

describe("building the F10 workbook", () => {
  const linkParts = (bytes: Uint8Array) => {
    const zip = unzipSync(bytes);
    return Object.fromEntries(Object.keys(zip).filter((p) => p.includes("externalLink")).map((p) => [p, strFromU8(zip[p]!)]));
  };

  it("writes trim references and linked references in Excel's stored form; the link parts stay byte for byte", () => {
    const files = pulled(F10);
    edit(files, "names/_unmanaged.xln", "TrimAll = Sheet1!$A$1.:.$A$10;", "TrimAll = Sheet1!$A$1.:.$A$11;\nExtTwo = '[Other.xlsx]Sheet1'!$A$2 + Other.xlsx!OtherVal;");
    edit(files, "names/sheets/Sheet1.xln", "@C6 = [Other.xlsx]Sheet1!$A$1;", "@C6 = [Other.xlsx]Sheet1!$A$1 + ROWS(A1:.A3);");
    const r = build(F10, files);
    expect(r.status, why(r)).toBe("built");
    expect(r.readBack!.ok, r.readBack!.problems.join("\n")).toBe(true);
    const wb = readWorkbook(r.bytes!);
    expect(wb.definedNames.find((d) => d.name === "TrimAll")!.definition).toBe("_xlfn._TRO_ALL(Sheet1!$A$1:$A$11)");
    expect(wb.definedNames.find((d) => d.name === "ExtTwo")!.definition).toBe("[1]Sheet1!$A$2 + [1]!OtherVal");
    expect(wb.sheets[0]!.formulas.find((f) => f.cell === "C6")!.text).toBe("[1]Sheet1!$A$1 + ROWS(_xlfn._TRO_TRAILING(A1:A3))");
    expect(linkParts(r.bytes!)).toEqual(linkParts(F10));
    const before = unzipSync(F10);
    const after = unzipSync(r.bytes!);
    for (const p of Object.keys(before).filter((p) => p.includes("externalLink"))) expect(after[p]).toEqual(before[p]);
    // Nothing left to do on the built file.
    expect(build(r.bytes!, { ...files, ...r.files }).status).toBe("up-to-date");
  });

  it("refuses a reference to a workbook the file has no link to", () => {
    const files = pulled(F10);
    edit(files, "names/_unmanaged.xln", "TrimAll = Sheet1!$A$1.:.$A$10;", "TrimAll = [Missing.xlsx]Sheet1!$A$1;");
    const r = build(F10, files);
    expect(r.status).not.toBe("built");
    expect(why(r)).toContain("this workbook has no link to Missing.xlsx");
  });
});
