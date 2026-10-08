// `@renamed('Cash Flow'!Old)`: a quoted sheet in the annotation's argument (bug of
// 2026-10-07: "'@renamed(' is not closed"). The parser keeps the argument as written and
// `renamedFrom` reads the sheet out of it; a rename that moves a name to or from a sheet
// whose name needs quotes builds and the build consumes the note.
import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { parseSourceFile, pullProject, readWorkbook, renamedFrom } from "../../src/index.js";
import { build, edit, why } from "./helpers.js";

const M = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
const T = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const rels = (entries: [string, string, string][]) =>
  `<Relationships xmlns="${PR}">${entries.map(([id, type, target]) => `<Relationship Id="${id}" Type="${T}/${type}" Target="${target}"/>`).join("")}</Relationships>`;

const BOOK = zipSync(
  Object.fromEntries(
    Object.entries({
      "[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
      "_rels/.rels": rels([["rId1", "officeDocument", "xl/workbook.xml"]]),
      "xl/workbook.xml": `<workbook ${M} ${R}><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Cash Flow" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="CF_Rate" localSheetId="1">0.1</definedName><definedName name="Wb_Rate">0.2</definedName></definedNames><calcPr calcId="191029"/></workbook>`,
      "xl/_rels/workbook.xml.rels": rels([
        ["rId1", "worksheet", "worksheets/sheet1.xml"],
        ["rId2", "worksheet", "worksheets/sheet2.xml"],
      ]),
      "xl/worksheets/sheet1.xml": `<worksheet ${M}><sheetData/></worksheet>`,
      "xl/worksheets/sheet2.xml": `<worksheet ${M}><sheetData/></worksheet>`,
    }).map(([k, v]) => [k, strToU8(v)]),
  ),
);
const U = "names/_unmanaged.xln";
const CF = "names/sheets/Cash Flow.xln";

describe("@renamed with a quoted sheet", () => {
  it("parses, `''` included", () => {
    const anns = (text: string) => parseSourceFile(U, text).entries[0]!.annotations;
    const p = parseSourceFile(U, "@renamed('Cash Flow'!Old)\nNew = 1;\n");
    expect(p.diagnostics).toEqual([]);
    expect(renamedFrom(p.entries[0]!.annotations)).toEqual({ name: "Old", scope: "Cash Flow" });
    expect(renamedFrom(anns("@renamed( 'Bob''s plan'!Old )\nNew = 1;\n"))).toEqual({ name: "Old", scope: "Bob's plan" });
    expect(renamedFrom(anns("@renamed('Cash Flow'!Old) @hidden\nNew = 1;\n"))).toEqual({ name: "Old", scope: "Cash Flow" });
    expect(parseSourceFile(U, "@renamed('Cash Flow'!Old\nNew = 1;\n").diagnostics[0]!.message).toBe("'@renamed(' is not closed");
  });

  it("a rename out of 'Cash Flow' to the workbook, and one into it, build and consume their notes", () => {
    const files = { ...pullProject(BOOK, "book.xlsx").files };
    edit(files, CF, "CF_Rate = 0.1;", "@renamed(!Wb_Rate)\nLocal_rate = 0.2;");
    edit(files, U, "Wb_Rate = 0.2;", "@renamed('Cash Flow'!CF_Rate)\nGlobal_rate = 0.1;");
    const r = build(BOOK, files);
    expect(r.status, why(r)).toBe("built");
    const names = readWorkbook(r.bytes!).definedNames.map((d) => `${d.scope.kind === "sheet" ? `${d.scope.name}!` : ""}${d.name}=${d.definition}`).sort();
    expect(names).toEqual(["Cash Flow!Local_rate=0.2", "Global_rate=0.1"]);
    expect(r.renamedConsumed!.map((e) => e.annotation).sort()).toEqual(["@renamed(!Wb_Rate)", "@renamed('Cash Flow'!CF_Rate)"]);
    for (const t of Object.values(r.sourceFiles!)) expect(t).not.toContain("@renamed");
  });
});
