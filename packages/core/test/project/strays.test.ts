// M3e: files that have no place in a project's names/ (VS Code cannot refuse a file made in
// the Explorer). The checker flags them and the build refuses them, with the same message.
import { describe, expect, it } from "vitest";
import { buildWorkbook, checkFile, pullProject, SourceModel, strayFile } from "../../src/index.js";
import { fixture } from "../build/helpers.js";

const F7 = fixture("f7_base.xlsx");
const SHEETS = ["S1", "S2"];

describe("strayFile", () => {
  it("what a pull writes has its place", () => {
    for (const p of ["names/_unmanaged.xln", "names/FN.xln", "names/FN~2.xln", "names/sheets/S1.xln", "names/sheets/s2.xln", "names/Größe_1.xln", "names/_x.xln"]) {
      expect(strayFile(p, SHEETS), p).toBeUndefined();
    }
    // Every file of a pulled project.
    for (const p of Object.keys(pullProject(F7, "book.xlsx").files)) expect(strayFile(p, SHEETS), p).toBeUndefined();
  });

  it("outside names/, hidden files, and an unknown workbook: not judged", () => {
    expect(strayFile("xln.config.json", SHEETS)).toBeUndefined();
    expect(strayFile("notes.txt", SHEETS)).toBeUndefined();
    expect(strayFile("names/.DS_Store", SHEETS)).toBeUndefined();
    expect(strayFile("names/sheets/Foo.xln")).toBeUndefined();
  });

  it("a sheet file for a sheet the workbook lacks", () => {
    expect(strayFile("names/sheets/Foo.xln", SHEETS)).toBe("there is no sheet Foo in the workbook: sheets are created in Excel, then pulled (a sheet file is named after an existing sheet)");
    // `IS~2.xln`, as a pull names a clash, reads as the sheet `IS~2`: the file's `@sheet(IS)` then names it.
    const m = new SourceModel();
    m.setFile("names/sheets/S1~2.xln", "@sheet(S1)\n\nX = 1;\n");
    expect(checkFile(m, "names/sheets/S1~2.xln", { sheets: SHEETS }).filter((p) => p.severity === "error")).toEqual([]);
    m.setFile("names/sheets/S1~2.xln", "X = 1;\n");
    expect(checkFile(m, "names/sheets/S1~2.xln", { sheets: SHEETS }).map((p) => p.code)).toContain("stray-file");
  });

  it("anything not .xln, folders a pull does not make, a module named not after a prefix", () => {
    expect(strayFile("names/notes.txt", SHEETS)).toMatch(/^names\/notes\.txt: only \.xln files belong in names\//);
    expect(strayFile("names/sheets/S1.xln.bak", SHEETS)).toMatch(/only \.xln files/);
    expect(strayFile("names/sheets/old/S1.xln", SHEETS)).toMatch(/no folders/);
    expect(strayFile("names/lib/FN.xln", SHEETS)).toMatch(/a module lives directly in names\//);
    for (const p of ["names/my module.xln", "names/FIN.old.xln", "names/1FIN.xln", "names/a-b.xln"]) {
      expect(strayFile(p, SHEETS), p).toMatch(/a module file is named after its prefix/);
    }
  });
});

describe("the checker and the build say the same", () => {
  const pulled = pullProject(F7, "book.xlsx").files;

  it("checkFile flags the file on its first line; the build refuses it", () => {
    const files: Record<string, string> = { ...pulled, "names/my module.xln": "// a module\nM.X = 1;\n", "names/sheets/Foo.xln": "X @A1 = 1;\n", "names/notes.txt": "" };
    const m = new SourceModel();
    for (const [p, t] of Object.entries(files)) if (p.startsWith("names/") && p.endsWith(".xln")) m.setFile(p, t);
    const mod = checkFile(m, "names/my module.xln", { sheets: SHEETS }).filter((p) => p.code === "stray-file");
    expect(mod.map((p) => [p.severity, m.files.get("names/my module.xln")!.text.slice(p.start, p.end)])).toEqual([["error", "// a module"]]);
    const foo = checkFile(m, "names/sheets/Foo.xln", { sheets: SHEETS });
    expect(foo.filter((p) => p.code === "stray-file").map((p) => p.message)).toEqual([strayFile("names/sheets/Foo.xln", SHEETS)]);
    // Said once: not again as the old unknown-sheet error.
    expect(foo.filter((p) => p.code === "unknown-sheet" && p.start === 0)).toEqual([]);

    const b = buildWorkbook({ workbook: F7, fileName: "book.xlsx", files }, { provenance: false, embed: false });
    expect(b.status).toBe("refused");
    const stray = b.plan.problems.filter((p) => p.code === "stray-file");
    expect(stray.map((p) => [p.file, p.line, p.message])).toEqual([
      ["names/my module.xln", 1, strayFile("names/my module.xln", SHEETS)],
      ["names/sheets/Foo.xln", 1, strayFile("names/sheets/Foo.xln", SHEETS)],
      ["names/notes.txt", 1, strayFile("names/notes.txt", SHEETS)],
    ]);
  });

  it("a hidden file in names/ does not stop a build", () => {
    const b = buildWorkbook({ workbook: F7, fileName: "book.xlsx", files: { ...pulled, "names/.DS_Store": "" } }, { provenance: false, embed: false });
    expect(b.status).toBe("up-to-date");
  });
});
