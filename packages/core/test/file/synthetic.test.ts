// Hand-built packages for what the probe files do not contain: Tables, x14 extensions,
// absolute relationship targets, prefixed SpreadsheetML, shared strings, sheet states.
import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { readWorkbook, XlsxError } from "../../src/index.js";

const M = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
const T = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function rels(entries: [string, string, string][]): string {
  return `<Relationships xmlns="${PR}">${entries
    .map(([id, type, target]) => `<Relationship Id="${id}" Type="${T}/${type}" Target="${target}"/>`)
    .join("")}</Relationships>`;
}

function zip(parts: Record<string, string>): Uint8Array {
  return zipSync(Object.fromEntries(Object.entries(parts).map(([k, v]) => [k, strToU8(v)])));
}

const ROOT_RELS = rels([["rId1", "officeDocument", "xl/workbook.xml"]]);

describe("synthetic workbook", () => {
  const bytes = zip({
    "_rels/.rels": ROOT_RELS,
    "xl/workbook.xml": `<workbook ${M} ${R}><sheets>
      <sheet name="Data" sheetId="5" r:id="rId1"/>
      <sheet name="Hidden one" sheetId="2" state="hidden" r:id="rId2"/>
      <sheet name="Chart1" sheetId="3" state="veryHidden" r:id="rId3"/>
    </sheets><definedNames>
      <definedName name="_xlnm.Print_Area" localSheetId="0" hidden="1">Data!$A$1:$C$3</definedName>
      <definedName name="Bad" localSheetId="7">1</definedName>
      <definedName name="Fn" function="1" vbProcedure="0" description="d">1&amp;2</definedName>
    </definedNames></workbook>`,
    "xl/_rels/workbook.xml.rels": rels([
      ["rId1", "worksheet", "/xl/worksheets/sheet1.xml"],
      ["rId2", "worksheet", "worksheets/sheet2.xml"],
      ["rId3", "chartsheet", "chartsheets/sheet1.xml"],
      ["rId4", "sharedStrings", "sharedStrings.xml"],
    ]),
    "xl/sharedStrings.xml": `<sst ${M}><si><t>plain</t></si><si><r><t>ri</t></r><r><t xml:space="preserve">ch </t></r><rPh><t>x</t></rPh></si></sst>`,
    "xl/worksheets/sheet1.xml": `<worksheet ${M} ${R} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
      <sheetData>
        <row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><f>A1&amp;"!"</f><v>1</v></c></row>
        <row><c t="b"><f>TRUE</f><v>1</v></c><c t="inlineStr"><f>"x"</f><is><t>x</t></is></c></row>
        <row r="4"><c r="C4"><f t="array" ref="C4:D5">{1,2;3,4}</f><v>1</v></c><c r="E4"><f t="dataTable" ref="E4:E6" dt2D="0" dtr="0" r1="A1"/></c></row>
      </sheetData>
      <conditionalFormatting sqref="A1"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="formula" val="Lo"/><cfvo type="max"/></colorScale></cfRule></conditionalFormatting>
      <tableParts count="1"><tablePart r:id="rId9"/></tableParts>
      <extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="expression" priority="3" id="{1}"><xm:f>'Hidden one'!A1&gt;Lim</xm:f></x14:cfRule><xm:sqref>B1:B9</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>
      <ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations count="1"><x14:dataValidation type="list"><x14:formula1><xm:f>'Hidden one'!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>C1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>
    </worksheet>`,
    "xl/worksheets/_rels/sheet1.xml.rels": rels([["rId9", "table", "../tables/table1.xml"]]),
    "xl/tables/table1.xml": `<table ${M} id="3" name="Table3" displayName="tblSales" ref="A10:C13" totalsRowCount="1"><autoFilter ref="A10:C12"/><tableColumns count="3">
      <tableColumn id="1" name="Qty"/><tableColumn id="2" name="Price"/>
      <tableColumn id="3" name="Total" totalsRowFunction="custom"><calculatedColumnFormula>tblSales[[#This Row],[Qty]]*tblSales[[#This Row],[Price]]</calculatedColumnFormula><totalsRowFormula>SUBTOTAL(109,[Total])</totalsRowFormula></tableColumn>
    </tableColumns></table>`,
    // SpreadsheetML with an explicit prefix, as some non-Excel writers emit it.
    "xl/worksheets/sheet2.xml": `<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row r="2"><x:c r="B2"><x:f>Data!B1</x:f><x:v>3</x:v></x:c></x:row></x:sheetData></x:worksheet>`,
    "xl/chartsheets/sheet1.xml": `<chartsheet ${M}/>`,
    "xl/charts/chart1.xml": "<c/>",
    "xl/externalLinks/externalLink1.xml": "<e/>",
    "xl/pivotTables/pivotTable1.xml": "<p/>",
    "xl/pivotCache/pivotCacheDefinition1.xml": "<p/>",
    "xl/pivotCache/pivotCacheRecords1.xml": "<p/>",
    "xl/media/image1.png": "not decompressed",
  });
  const wb = readWorkbook(bytes);
  const [data, hidden, chart] = wb.sheets as [(typeof wb.sheets)[0], (typeof wb.sheets)[0], (typeof wb.sheets)[0]];

  it("reads sheet states, kinds and absolute and relative targets", () => {
    expect(wb.sheets.map(({ name, sheetId, state, kind, part }) => ({ name, sheetId, state, kind, part }))).toEqual([
      { name: "Data", sheetId: 5, state: "visible", kind: "worksheet", part: "xl/worksheets/sheet1.xml" },
      { name: "Hidden one", sheetId: 2, state: "hidden", kind: "worksheet", part: "xl/worksheets/sheet2.xml" },
      { name: "Chart1", sheetId: 3, state: "veryHidden", kind: "chartsheet", part: "xl/chartsheets/sheet1.xml" },
    ]);
    expect(chart.formulas).toEqual([]);
  });

  it("flags built-in names, invalid scopes and keeps other attributes", () => {
    const [pa, bad, fn] = wb.definedNames;
    expect(pa).toMatchObject({ hidden: true, isBuiltIn: true, isXlPrefixed: true, scope: { kind: "sheet", position: 0, name: "Data" } });
    expect(bad).toMatchObject({ scopeInvalid: true, scope: { kind: "sheet", position: 7, name: "" } });
    expect(fn).toMatchObject({ definition: "1&2", attributes: { function: "1", vbProcedure: "0", description: "d" }, isBuiltIn: false });
    expect(wb.warnings).toEqual(['defined name "Bad": localSheetId 7 is not a sheet position']);
  });

  it("reads cached values of every type, and cells without r", () => {
    const at = (c: string) => data.formulas.find((f) => f.cell === c)?.value;
    expect(at("B1")).toEqual({ type: "s", raw: "1", value: "rich " });
    expect(at("A2")).toEqual({ type: "b", raw: "1", value: true });
    expect(at("B2")).toEqual({ type: "inlineStr", raw: "x", value: "x" });
    expect(data.formulas.find((f) => f.cell === "C4")).toMatchObject({ kind: "array", range: "C4:D5", text: "{1,2;3,4}" });
    expect(data.formulas.find((f) => f.cell === "E4")).toMatchObject({
      kind: "data-table",
      text: "",
      attributes: { dt2D: "0", dtr: "0", r1: "A1" },
      value: { raw: undefined, value: undefined },
    });
    expect(data.spills).toEqual([]); // a legacy array is not a spill
  });

  it("reads conditional formatting and validation from the main part and x14 extensions", () => {
    expect(data.conditionalFormats).toEqual([
      { sqref: "A1", type: "colorScale", priority: 2, formulas: ["Lo"], ext: false },
      { sqref: "B1:B9", type: "expression", priority: 3, formulas: ["'Hidden one'!A1>Lim"], ext: true },
    ]);
    expect(data.dataValidations).toEqual([
      { sqref: "C1", type: "list", formula1: "'Hidden one'!$A$1:$A$3", formula2: undefined, ext: true },
    ]);
  });

  it("reads Tables through the sheet relationships", () => {
    expect(wb.tables).toEqual([
      {
        id: 3,
        name: "Table3",
        displayName: "tblSales",
        sheet: { position: 0, name: "Data" },
        part: "xl/tables/table1.xml",
        ref: "A10:C13",
        headerRowCount: 1,
        totalsRowCount: 1,
        columns: [
          { id: 1, name: "Qty" },
          { id: 2, name: "Price" },
          {
            id: 3,
            name: "Total",
            calculatedColumnFormula: "tblSales[[#This Row],[Qty]]*tblSales[[#This Row],[Price]]",
            totalsRowFormula: "SUBTOTAL(109,[Total])",
            totalsRowFunction: "custom",
          },
        ],
      },
    ]);
    expect(data.tables).toEqual(["tblSales"]);
  });

  it("reads prefixed SpreadsheetML", () => {
    expect(hidden.formulas).toMatchObject([{ cell: "B2", text: "Data!B1", value: { type: "n", value: 3 } }]);
  });

  it("lists the parts a rename must look at", () => {
    expect(wb.otherParts).toEqual([
      { kind: "chartsheet", path: "xl/chartsheets/sheet1.xml" },
      { kind: "chart", path: "xl/charts/chart1.xml" },
      { kind: "externalLink", path: "xl/externalLinks/externalLink1.xml" },
      { kind: "pivotTable", path: "xl/pivotTables/pivotTable1.xml" },
      { kind: "pivotCache", path: "xl/pivotCache/pivotCacheDefinition1.xml" },
    ]);
    expect(wb.parts).toContain("xl/media/image1.png");
  });
});

describe("errors", () => {
  it("rejects bytes that are not a zip", () => {
    expect(() => readWorkbook(strToU8("hello"))).toThrow(XlsxError);
  });
  it("rejects a zip without a workbook", () => {
    expect(() => readWorkbook(zip({ "a.txt": "x" }))).toThrow(XlsxError);
  });
  it("falls back to xl/workbook.xml without root relationships", () => {
    const wb = readWorkbook(zip({ "xl/workbook.xml": `<workbook ${M}><sheets/></workbook>` }));
    expect(wb.sheets).toEqual([]);
  });
  it("warns about a shared child without master", () => {
    const wb = readWorkbook(
      zip({
        "_rels/.rels": ROOT_RELS,
        "xl/workbook.xml": `<workbook ${M} ${R}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
        "xl/_rels/workbook.xml.rels": rels([["rId1", "worksheet", "worksheets/s.xml"]]),
        "xl/worksheets/s.xml": `<worksheet ${M}><sheetData><row r="1"><c r="A1"><f t="shared" si="4"/></c></row></sheetData></worksheet>`,
      }),
    );
    expect(wb.sheets[0]!.formulas[0]).toMatchObject({ kind: "shared-child", si: 4 });
    expect(wb.sheets[0]!.formulas[0]!.master).toBeUndefined();
    expect(wb.warnings).toEqual(["xl/worksheets/s.xml: shared formula si=4 in A1 has no master"]);
  });
});
