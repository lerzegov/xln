// M3b-2, the cell backend: every rule of probe F8's recipe on the F8 workbooks, compared
// with what F8's patcher wrote (f8_p1.xlsx, f8_q1.xlsx: files Excel opened without repair
// and with the oracle's values), plus read-back (E3) of each build.
import { unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import {
  applyChangeSetWithReport,
  decompile,
  ensureDynamicArrayCm,
  findDynamicArrayCm,
  newMetadataXml,
  patchSheetXml,
  rawZipRecords,
  readBack,
  readSheetCells,
  readWorkbook,
  replaceZipEntries,
  rewriteZip,
  utf8,
  type Change,
  type CellSnapshot,
} from "../../src/index.js";
import { decodeText, Package } from "../../src/file/package.js";
import { fixture } from "./helpers.js";

const BASE = fixture("f8_base.xlsx");
const P1 = fixture("f8_p1.xlsx");
const PLAIN = fixture("f8_plain_base.xlsx");
const Q1 = fixture("f8_q1.xlsx");
const SHEET = { N: "xl/worksheets/sheet1.xml", D: "xl/worksheets/sheet2.xml", Slot: "xl/worksheets/sheet3.xml", Sh: "xl/worksheets/sheet4.xml" } as const;

function set(sheet: string, range: string, stored: string, display = decompile(stored)): Change {
  return { op: "set-cell-formula", sheet, range, stored, display };
}
function clear(sheet: string, range: string): Change {
  return { op: "clear-cell-formula", sheet, range };
}

function part(bytes: Uint8Array, name: string): string {
  return decodeText(new Package(bytes).raw(name)!);
}

function cells(bytes: Uint8Array, sheetPart: string): Map<string, CellSnapshot> {
  return readSheetCells(part(bytes, sheetPart));
}

/** The `<c>` element of a cell as stored. */
function cellXml(bytes: Uint8Array, sheetPart: string, ref: string): string | undefined {
  const xml = part(bytes, sheetPart);
  const at = xml.indexOf(`<c r="${ref}"`);
  if (at < 0) return undefined;
  const selfEnd = xml.indexOf("/>", at);
  const close = xml.indexOf("</c>", at);
  const open = xml.indexOf(">", at);
  return selfEnd >= 0 && selfEnd < open ? xml.slice(at, selfEnd + 2) : xml.slice(at, close + 4);
}

/** Applies, reads back (E3) and returns the bytes and the patch report. */
function apply(bytes: Uint8Array, changes: Change[]) {
  const r = applyChangeSetWithReport(bytes, changes);
  const rb = readBack(bytes, r.bytes, changes, []);
  expect(rb.problems).toEqual([]);
  return { ...r, readBack: rb };
}

// The F8 edits xln makes the recipe's way (p1 made some of them deliberately otherwise: the traps).
const F8_EDITS: Change[] = [
  set("N", "B1", "SUM(A1:A3)*100"),
  set("N", "B2", "_xlfn.LET(_xlpm.x,A1*10,_xlpm.x+1)"),
  clear("N", "B3"),
  set("N", "B5", "SUM(A1:A3*2)"),
  set("N", "B6", "1+2"),
  set("N", "C1", "A1+100"),
  set("N", "C2", 'UPPER("ab")&A1'),
  set("N", "C3", "_xlfn.XLOOKUP(2,A1:A3,A1:A3)*7"),
  set("N", "C4", "A1:A3*3"),
  set("D", "C6", "_xlfn.SEQUENCE(5)"),
  set("D", "F6", "_xlfn.SEQUENCE(2)"),
  set("D", "N6", "42"),
  set("D", "P6", "_xlfn.SEQUENCE(3)"),
  set("Slot", "B1", "_xlfn.SEQUENCE(1,3)*10"),
  set("Slot", "B2", "SUM(_xlfn.ANCHORARRAY(B1))/2"),
  set("Slot", "B6", "Costs*0.25"),
  set("Slot", "B9", "Tax+1"),
  set("Slot", "F1", "SUM(Revenue)"),
  set("Sh", "B3", "A3*100"),
  set("Sh", "D1", "A1+2000"),
];

describe("F8 recipe on f8_base.xlsx", () => {
  const r = apply(BASE, F8_EDITS);
  const out = r.bytes;

  it("writes every formula in dynamic-array form: cm of the XLDAPR record, t=array, ref = the cell, no value", () => {
    expect(r.report.metadata).toBe("kept");
    expect(r.report.cm).toBe(1);
    expect(cellXml(out, SHEET.N, "B1")).toBe('<c r="B1" cm="1"><f t="array" ref="B1">SUM(A1:A3)*100</f></c>');
    expect(cellXml(out, SHEET.N, "B2")).toBe('<c r="B2" cm="1"><f t="array" ref="B2">_xlfn.LET(_xlpm.x,A1*10,_xlpm.x+1)</f></c>');
    // As F8 wrote it into p1 and Excel opened it with the oracle's values.
    for (const ref of ["B5", "B6", "C4"]) expect(cellXml(out, SHEET.N, ref)).toBe(cellXml(P1, SHEET.N, ref));
    const kinds = readWorkbook(out).sheets[0]!.formulas.filter((f) => ["B1", "B2", "B5", "B6", "C1", "C2", "C3", "C4"].includes(f.cell));
    expect(kinds.map((f) => f.kind)).toEqual(Array(8).fill("dynamic-array"));
  });

  it("turns value and shared-string cells into formulas: <v> and t go", () => {
    expect(cellXml(BASE, SHEET.N, "C2")).toBe('<c r="C2" t="s"><v>0</v></c>');
    expect(cellXml(out, SHEET.N, "C1")).toBe('<c r="C1" cm="1"><f t="array" ref="C1">A1+100</f></c>');
    expect(cellXml(out, SHEET.N, "C2")).toBe('<c r="C2" cm="1"><f t="array" ref="C2">UPPER("ab")&amp;A1</f></c>');
  });

  it("clears a formula: <f>, <v>, t, cm go, the element stays (as in p1)", () => {
    expect(cellXml(out, SHEET.N, "B3")).toBe('<c r="B3"/>');
    expect(r.report.sheets["N"]!.cleared).toEqual(["B3"]);
  });

  it("spill anchors: ref = the anchor, the old spill's cells removed (bigger, smaller, spill → scalar), as in p1", () => {
    for (const ref of ["C6", "F6", "P6"]) expect(cellXml(out, SHEET.D, ref)).toBe(cellXml(P1, SHEET.D, ref));
    expect(cellXml(out, SHEET.D, "N6")).toBe('<c r="N6" cm="1"><f t="array" ref="N6">42</f></c>');
    expect(r.report.sheets["D"]!.ghosts).toEqual(["C7", "F7", "N7", "C8", "F8", "N8", "F9"]);
    for (const g of r.report.sheets["D"]!.ghosts) {
      expect(cellXml(out, SHEET.D, g)).toBeUndefined();
      expect(cellXml(P1, SHEET.D, g)).toBeUndefined();
    }
    // Spills whose anchor did not change keep theirs.
    expect(cellXml(out, SHEET.D, "H7")).toBe('<c r="H7"><v>2</v></c>');
  });

  it("slots: <c> inserted in column order, a missing <row> in row order without spans", () => {
    expect(r.report.sheets["Slot"]!.inserted).toEqual(["B1", "F1", "B2", "B6", "B9"]);
    expect(r.report.sheets["Slot"]!.rowsInserted).toEqual([9]);
    const xml = part(out, SHEET.Slot);
    expect(xml).toContain('<c r="A1" t="s"><v>2</v></c><c r="B1" cm="1"><f t="array" ref="B1">_xlfn.SEQUENCE(1,3)*10</f></c><c r="F1" cm="1"><f t="array" ref="F1">SUM(Revenue)</f></c></row>');
    expect(xml).toContain('</row><row r="9"><c r="B9" cm="1"><f t="array" ref="B9">Tax+1</f></c></row></sheetData>');
    // The same places as F8's patcher chose; only the form differs (p1 wrote these plain).
    const strip = (s: string) => s.split('<c r="').map((c) => c.slice(0, c.indexOf('"'))).join(",");
    expect(strip(xml)).toBe(strip(part(P1, SHEET.Slot)));
  });

  it("shared formulas: an edited child gets its own formula, the master stays; an edited master un-shares its group", () => {
    expect(cellXml(out, SHEET.Sh, "B1")).toBe('<c r="B1"><f t="shared" ref="B1:B5" si="0">A1*2</f><v>2</v></c>');
    expect(cellXml(out, SHEET.Sh, "B3")).toBe('<c r="B3" cm="1"><f t="array" ref="B3">A3*100</f></c>');
    expect(cellXml(out, SHEET.Sh, "B4")).toBe('<c r="B4"><f t="shared" si="0"/><v>8</v></c>');
    expect(cellXml(out, SHEET.Sh, "D1")).toBe('<c r="D1" cm="1"><f t="array" ref="D1">A1+2000</f></c>');
    // The other members: their own text, translated, value kept (exactly as p1).
    for (const ref of ["D2", "D3", "D4", "D5"]) expect(cellXml(out, SHEET.Sh, ref)).toBe(cellXml(P1, SHEET.Sh, ref));
    expect(r.report.sheets["Sh"]!.unshared).toEqual(["D2", "D3", "D4", "D5"]);
  });

  it("drops calcChain.xml with its relationship and content type; sets fullCalcOnLoad", () => {
    expect(r.report.calcChainDropped).toBe("xl/calcChain.xml");
    const names = Object.keys(unzipSync(out));
    expect(names).not.toContain("xl/calcChain.xml");
    expect(part(out, "xl/_rels/workbook.xml.rels")).not.toContain("calcChain");
    expect(part(out, "[Content_Types].xml")).not.toContain("calcChain");
    expect(part(out, "xl/_rels/workbook.xml.rels")).toBe(part(BASE, "xl/_rels/workbook.xml.rels").replace('<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>', ""));
    expect(part(out, "xl/workbook.xml")).toContain('fullCalcOnLoad="1"');
  });

  it("copies every other part byte for byte; the sheets change only inside <sheetData>", () => {
    const a = rawZipRecords(BASE);
    const b = rawZipRecords(out);
    const changed = [...a.keys()].filter((k) => !b.has(k) || !a.get(k)!.every((x, i) => x === b.get(k)![i]) || a.get(k)!.length !== b.get(k)!.length);
    expect(changed.sort()).toEqual(["[Content_Types].xml", "xl/_rels/workbook.xml.rels", "xl/calcChain.xml", "xl/workbook.xml", ...Object.values(SHEET)].sort());
    expect(r.readBack.cellsChecked).toBeGreaterThan(20);
  });

  it("matches F8's patched cells, cell by cell, except where p1 deliberately did otherwise", () => {
    // p1's traps and alternatives: plain formulas (B4, B7, plain B1… on N; N6), kept refs and
    // ghosts (H6, J6, L6), promotion of C2 and the shrunk ref of E1 on Sh.
    const skipped = new Set(["N!B4", "N!B7", "D!H6", "D!J6", "D!L6", "D!N6", "Sh!C1", "Sh!C2", "Sh!C3", "Sh!C4", "Sh!C5", "Sh!E1", "Sh!E5"]);
    for (const [name, sp] of Object.entries(SHEET)) {
      const ours = cells(out, sp);
      const theirs = cells(P1, sp);
      for (const [ref, t] of theirs) {
        if (skipped.has(`${name}!${ref}`)) continue;
        const o = ours.get(ref);
        expect(o, `${name}!${ref}`).toBeDefined();
        expect(o!.f?.text, `${name}!${ref}`).toBe(t.f?.text);
        expect(o!.v, `${name}!${ref}`).toBe(t.v);
        expect(o!.attrs["s"], `${name}!${ref}`).toBe(t.attrs["s"]);
        if (t.f) expect(o!.f!.attrs["t"] === "array" || o!.f!.attrs["t"] === undefined || o!.f!.attrs["t"] === "shared").toBe(true);
      }
      for (const ref of ours.keys()) if (![...skipped].some((s) => s.startsWith(name + "!") && s.endsWith("!" + ref))) expect(theirs.has(ref), `${name}!${ref}`).toBe(true);
    }
  });
});

describe("ranges and shared groups", () => {
  it("writes a range statement as the top-left formula shifted into each cell", () => {
    const r = apply(BASE, [set("Sh", "B2:C3", "A2*$A$1+B$1")]);
    expect(cellXml(r.bytes, SHEET.Sh, "B2")).toBe('<c r="B2" cm="1"><f t="array" ref="B2">A2*$A$1+B$1</f></c>');
    expect(cellXml(r.bytes, SHEET.Sh, "C2")).toBe('<c r="C2" cm="1"><f t="array" ref="C2">B2*$A$1+C$1</f></c>');
    expect(cellXml(r.bytes, SHEET.Sh, "B3")).toBe('<c r="B3" cm="1"><f t="array" ref="B3">A3*$A$1+B$1</f></c>');
    expect(cellXml(r.bytes, SHEET.Sh, "C3")).toBe('<c r="C3" cm="1"><f t="array" ref="C3">B3*$A$1+C$1</f></c>');
    // Only children were edited: the masters and the other children are untouched.
    expect(cellXml(r.bytes, SHEET.Sh, "B1")).toBe(cellXml(BASE, SHEET.Sh, "B1"));
    expect(cellXml(r.bytes, SHEET.Sh, "C4")).toBe(cellXml(BASE, SHEET.Sh, "C4"));
    expect(r.report.sheets["Sh"]!.unshared).toEqual([]);
  });

  it("a range over a master un-shares the members outside it; those inside get the range's formula", () => {
    const r = apply(BASE, [set("Sh", "E1:E2", "A1*40")]);
    expect(cellXml(r.bytes, SHEET.Sh, "E1")).toBe('<c r="E1" cm="1"><f t="array" ref="E1">A1*40</f></c>');
    expect(cellXml(r.bytes, SHEET.Sh, "E2")).toBe('<c r="E2" cm="1"><f t="array" ref="E2">A2*40</f></c>');
    expect(cellXml(r.bytes, SHEET.Sh, "E3")).toBe('<c r="E3"><f>A3*4</f><v>12</v></c>');
    expect(cellXml(r.bytes, SHEET.Sh, "E5")).toBe('<c r="E5"><f>A5*4</f><v>20</v></c>');
    expect(r.report.sheets["Sh"]!.unshared).toEqual(["E3", "E4", "E5"]);
  });

  it("a whole group rewritten leaves no shared formula behind", () => {
    const r = apply(BASE, [set("Sh", "C1:C5", "A1+7")]);
    expect(readWorkbook(r.bytes).sheets[3]!.sharedFormulas.map((g) => g.si)).toEqual([0, 2, 3]);
    expect(r.report.sheets["Sh"]!.unshared).toEqual([]);
  });

  it("clearing a master un-shares the rest of its group", () => {
    const r = apply(BASE, [clear("Sh", "B1")]);
    expect(cellXml(r.bytes, SHEET.Sh, "B1")).toBe('<c r="B1"/>');
    expect(cellXml(r.bytes, SHEET.Sh, "B2")).toBe('<c r="B2"><f>A2*2</f><v>4</v></c>');
  });

  it("the later of two changes to a cell wins; sets come after clears", () => {
    const r = apply(BASE, [set("N", "B1:B2", "A1"), clear("N", "B2")]);
    // orderChanges: clear first, then set.
    expect(cellXml(r.bytes, SHEET.N, "B2")).toBe('<c r="B2" cm="1"><f t="array" ref="B2">A2</f></c>');
  });
});

describe("attributes, styles and edge cases (synthetic sheets)", () => {
  const sheet = (data: string) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData>${data}</sheetData><pageMargins left="0.7"/></worksheet>`;

  it("keeps s and other attributes, drops t, cm and vm, keeps extLst children", () => {
    const xml = sheet('<row r="1"><c r="A1" s="3" t="e" vm="2" ph="1"><v>#N/A</v><extLst><ext uri="x"/></extLst></c></row>');
    const { xml: out } = patchSheetXml(xml, [{ op: "set", range: "A1", stored: "1+1" }], 4);
    expect(out).toContain('<c r="A1" s="3" cm="4" ph="1"><f t="array" ref="A1">1+1</f><extLst><ext uri="x"/></extLst></c>');
  });

  it("an old spill's styled cells keep their style, empty; plain ones go", () => {
    const xml = sheet('<row r="1"><c r="A1" s="1" cm="1"><f t="array" ref="A1:A3">SEQUENCE(3)</f><v>1</v></c></row><row r="2"><c r="A2" s="1"><v>2</v></c></row><row r="3"><c r="A3"><v>3</v></c></row>');
    const { xml: out, report } = patchSheetXml(xml, [{ op: "set", range: "A1", stored: "_xlfn.SEQUENCE(2)" }], 1);
    expect(out).toContain('<row r="2"><c r="A2" s="1"/></row><row r="3"></row>');
    expect(report.ghosts).toEqual(["A2", "A3"]);
  });

  it("inserts into an empty <sheetData/> and a self-closing <row/>", () => {
    const xml = sheet("").replace("<sheetData></sheetData>", "<sheetData/>");
    expect(patchSheetXml(xml, [{ op: "set", range: "B2", stored: "1" }], 1).xml).toContain('<sheetData><row r="2"><c r="B2" cm="1"><f t="array" ref="B2">1</f></c></row></sheetData>');
    const xml2 = sheet('<row r="2" ht="20" customHeight="1"/>');
    expect(patchSheetXml(xml2, [{ op: "set", range: "C2", stored: "1" }], 1).xml).toContain('<row r="2" ht="20" customHeight="1"><c r="C2" cm="1"><f t="array" ref="C2">1</f></c></row>');
  });

  it("escapes formula text and keeps a prefixed namespace", () => {
    const xml = '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>';
    const out = patchSheetXml(xml, [{ op: "set", range: "A1:B1", stored: 'IF(A2<1,"<&>",A2)' }], 1).xml;
    expect(out).toBe(
      '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row r="1"><x:c r="A1" cm="1"><x:f t="array" ref="A1">IF(A2&lt;1,"&lt;&amp;&gt;",A2)</x:f></x:c><x:c r="B1" cm="1"><x:f t="array" ref="B1">IF(B2&lt;1,"&lt;&amp;&gt;",B2)</x:f></x:c></x:row></x:sheetData></x:worksheet>',
    );
  });

  it("refuses part of a legacy array and a data table, as Excel does", () => {
    const xml = sheet('<row r="1"><c r="A1"><f t="array" ref="A1:A2">B1:B2</f><v>1</v></c><c r="C1"><f t="dataTable" ref="C1:C2" dt2D="0" dtr="0" r1="E1"/><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>');
    expect(() => patchSheetXml(xml, [{ op: "set", range: "A2", stored: "1" }], 1)).toThrow(/legacy array/);
    expect(() => patchSheetXml(xml, [{ op: "set", range: "C2", stored: "1" }], 1)).toThrow(/data table/);
    // The legacy array's anchor may change: it becomes a dynamic array, its old area emptied.
    const r = patchSheetXml(xml, [{ op: "set", range: "A1", stored: "B1:B3" }], 1);
    expect(r.report.ghosts).toEqual(["A2"]);
  });

  it("clearing an empty or absent cell, or one holding a typed value, changes nothing", () => {
    const xml = sheet('<row r="1"><c r="A1" s="2"/><c r="B1" t="s"><v>0</v></c><c r="C1"><v>5</v></c></row>');
    expect(patchSheetXml(xml, [{ op: "clear", range: "A1:D3" }], 1).xml).toBe(xml);
  });

  it("read-back accepts a clear over value cells and keeps them", () => {
    apply(BASE, [clear("N", "A1:C4")]);
  });
});

describe("metadata.xml", () => {
  it("is created with its relationship and content type when absent (F8 q1)", () => {
    const changes = [set("P", "B1", "_xlfn.SEQUENCE(3)*A1"), set("P", "C1", "SUM(_xlfn.ANCHORARRAY(B1))")];
    const r = apply(PLAIN, changes);
    expect(r.report.metadata).toBe("created");
    // The same record F8 created in q1, which Excel kept as it was.
    expect(part(r.bytes, "xl/metadata.xml")).toBe(part(Q1, "xl/metadata.xml"));
    expect(part(r.bytes, "xl/_rels/workbook.xml.rels")).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>');
    expect(part(r.bytes, "[Content_Types].xml")).toContain('<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>');
    expect(cellXml(r.bytes, "xl/worksheets/sheet1.xml", "B1")).toBe(cellXml(Q1, "xl/worksheets/sheet1.xml", "B1"));
    expect(readWorkbook(r.bytes).sheets[0]!.formulas.map((f) => f.kind)).toEqual(["dynamic-array", "dynamic-array"]);
    expect(Object.keys(unzipSync(r.bytes)).at(-1)).toBe("xl/metadata.xml");
  });

  it("finds the XLDAPR record when it is not the first; adds it to a part holding other types", () => {
    const rich =
      '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes>' +
      '<futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" i="0"/></ext></extLst></bk></futureMetadata>' +
      '<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>';
    expect(findDynamicArrayCm(rich)).toBeUndefined();
    const e = ensureDynamicArrayCm(rich);
    expect(e.cm).toBe(1);
    expect(e.xml).toContain('<metadataTypes count="2"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/><metadataType name="XLDAPR"');
    expect(e.xml).toContain('</futureMetadata><futureMetadata name="XLDAPR" count="1"><bk>');
    expect(e.xml).toContain('<cellMetadata count="1"><bk><rc t="2" v="0"/></bk></cellMetadata><valueMetadata');
    expect(findDynamicArrayCm(e.xml)).toBe(1);
    expect(ensureDynamicArrayCm(e.xml).changed).toBe(false);
    // A second block after a first that is something else.
    const two = newMetadataXml().replace('<cellMetadata count="1"><bk><rc t="1" v="0"/></bk>', '<cellMetadata count="2"><bk><rc t="9" v="0"/></bk><bk><rc t="1" v="0"/></bk>');
    expect(findDynamicArrayCm(two)).toBe(2);
  });
});

describe("zip rewriting", () => {
  it("adds and removes entries; untouched records stay byte for byte; the result unzips", () => {
    const out = rewriteZip(BASE, { remove: ["xl/calcChain.xml"], add: new Map([["xl/extra.xml", utf8("<a/>")]]) });
    const files = unzipSync(out);
    expect(Object.keys(files)).toEqual([...Object.keys(unzipSync(BASE)).filter((n) => n !== "xl/calcChain.xml"), "xl/extra.xml"]);
    expect(new TextDecoder().decode(files["xl/extra.xml"])).toBe("<a/>");
    const a = rawZipRecords(BASE);
    for (const [k, v] of rawZipRecords(out)) if (k !== "xl/extra.xml") expect(Buffer.from(v).equals(Buffer.from(a.get(k)!))).toBe(true);
    expect(() => rewriteZip(BASE, { remove: ["nope"] })).toThrow(/no entry/);
    expect(() => rewriteZip(BASE, { add: new Map([["xl/workbook.xml", utf8("")]]) })).toThrow(/exists/);
  });
});

describe("read-back catches a wrong build", () => {
  it("flags an untouched cell that changed, a cell left with its value, a calcChain kept", () => {
    const changes = [set("N", "B1", "SUM(A1:A3)*100")];
    const good = applyChangeSetWithReport(BASE, changes).bytes;
    const xml = part(good, SHEET.N);
    const bad1 = replaceZipEntries(good, new Map([[SHEET.N, utf8(xml.replace('<c r="B2"><f>A1*10</f><v>10</v>', '<c r="B2"><f>A1*11</f><v>10</v>'))]]));
    expect(readBack(BASE, bad1, changes, []).problems).toContain("N!B2 changed although no change touched it");
    const bad2 = replaceZipEntries(good, new Map([[SHEET.N, utf8(xml.replace("SUM(A1:A3)*100</f>", "SUM(A1:A3)*100</f><v>6</v>"))]]));
    expect(readBack(BASE, bad2, changes, []).problems).toContain("N!B1: keeps a cached value");
    const bad3 = replaceZipEntries(good, new Map([[SHEET.N, utf8(xml.replace("SUM(A1:A3)*100</f>", "SUM(A1:A3)*10</f>"))]]));
    expect(readBack(BASE, bad3, changes, []).problems.join()).toContain("the built file says SUM(A1:A3)*10");
    // Name-only builds still keep calcChain: the cell rules apply only to cell builds.
    const kept = rewriteZip(BASE, { replace: new Map([[SHEET.N, utf8(xml)], ["xl/workbook.xml", utf8(part(good, "xl/workbook.xml"))]]) });
    expect(readBack(BASE, kept, changes, []).problems.join("\n")).toMatch(/calculation chain is still referenced/);
  });

  it("refuses a stored form Excel does not write, although it decompiles to the source (M3d)", () => {
    // The pre-M3d spelling of an optional parameter: our decompiler reads it as `[y]`.
    const wrong = "_xlfn.LAMBDA(_xlpm.x,[_xlpm.y],_xlpm.x+IF(_xlfn.ISOMITTED(_xlpm.y),0,_xlpm.y))";
    const cell = [set("N", "B1", wrong + "(1)")];
    const rb = readBack(BASE, applyChangeSetWithReport(BASE, cell).bytes, cell, []);
    expect(rb.ok).toBe(false);
    expect(rb.problems).toEqual([
      "N!B1: not Excel's stored form: '[_xlpm.y]': an optional LAMBDA parameter is stored as '_xlop.y', without brackets (Excel drops the name otherwise)",
    ]);
    const name: Change[] = [{ op: "set-name", name: "Opt", scope: null, stored: wrong, display: decompile(wrong), comment: null, hidden: false, fields: ["created"] }];
    const rn = readBack(BASE, applyChangeSetWithReport(BASE, name).bytes, name, []);
    expect(rn.problems).toEqual([
      "Opt: not Excel's stored form: '[_xlpm.y]': an optional LAMBDA parameter is stored as '_xlop.y', without brackets (Excel drops the name otherwise)",
    ]);
    // Excel's own form passes.
    const right: Change[] = [{ ...name[0]!, stored: wrong.replace("[_xlpm.y]", "_xlop.y") } as Change];
    expect(readBack(BASE, applyChangeSetWithReport(BASE, right).bytes, right, []).problems).toEqual([]);
    const spill = [set("N", "B1", "SUM(A1#)", "SUM(A1#)")];
    expect(readBack(BASE, applyChangeSetWithReport(BASE, spill).bytes, spill, []).problems).toEqual([
      "N!B1: not Excel's stored form: the spill operator '#' is display form; Excel stores _xlfn.ANCHORARRAY(…)",
    ]);
  });
});
