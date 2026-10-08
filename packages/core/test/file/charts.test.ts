// Chart parts: the `<c:f>` formulas of series, titles and labels, and the sheet whose
// drawing shows each chart (a worksheet or a chartsheet). C10 reads them before it calls
// a name unused.
import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { readWorkbook } from "../../src/index.js";

const M = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
const T = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function rels(entries: [string, string, string][]): string {
  return `<Relationships xmlns="${PR}">${entries
    .map(([id, type, target]) => `<Relationship Id="${id}" Type="${type.includes("/") ? type : `${T}/${type}`}" Target="${target}"/>`)
    .join("")}</Relationships>`;
}

const zip = (parts: Record<string, string>): Uint8Array => zipSync(Object.fromEntries(Object.entries(parts).map(([k, v]) => [k, strToU8(v)])));

const SERIES = `<c:chartSpace ${C}><c:chart><c:title><c:tx><c:strRef><c:f>'Data'!$A$1</c:f></c:strRef></c:tx></c:title><c:plotArea><c:barChart>
  <c:ser><c:idx val="0"/><c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>
  <c:cat><c:strRef><c:f>'Data'!$A$2:$A$5</c:f><c:strCache><c:ptCount val="0"/></c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:f>[0]!Sales</c:f><c:numCache><c:ptCount val="0"/></c:numCache></c:numRef></c:val></c:ser>
</c:barChart></c:plotArea></c:chart></c:chartSpace>`;

describe("charts", () => {
  const bytes = zip({
    "_rels/.rels": rels([["rId1", "officeDocument", "xl/workbook.xml"]]),
    "xl/workbook.xml": `<workbook ${M} ${R}><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Chart1" sheetId="2" r:id="rId2"/></sheets>
      <definedNames><definedName name="Sales">Data!$B$2:$B$5</definedName></definedNames></workbook>`,
    "xl/_rels/workbook.xml.rels": rels([
      ["rId1", "worksheet", "worksheets/sheet1.xml"],
      ["rId2", "chartsheet", "chartsheets/sheet1.xml"],
    ]),
    "xl/worksheets/sheet1.xml": `<worksheet ${M} ${R}><sheetData/><drawing r:id="rId1"/></worksheet>`,
    "xl/worksheets/_rels/sheet1.xml.rels": rels([["rId1", "drawing", "../drawings/drawing1.xml"]]),
    "xl/drawings/drawing1.xml": "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>",
    "xl/drawings/_rels/drawing1.xml.rels": rels([["rId1", "chart", "../charts/chart1.xml"]]),
    "xl/charts/chart1.xml": SERIES,
    "xl/chartsheets/sheet1.xml": `<chartsheet ${M} ${R}><drawing r:id="rId1"/></chartsheet>`,
    "xl/chartsheets/_rels/sheet1.xml.rels": rels([["rId1", "drawing", "../drawings/drawing2.xml"]]),
    "xl/drawings/drawing2.xml": "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>",
    "xl/drawings/_rels/drawing2.xml.rels": rels([["rId1", "http://schemas.microsoft.com/office/2014/relationships/chartEx", "../charts/chartEx1.xml"]]),
    "xl/charts/chartEx1.xml": `<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>'Data'!Local</cx:f></cx:numDim></cx:data></cx:chartData></cx:chartSpace>`,
    "xl/charts/chart9.xml": `<c:chartSpace ${C}><c:chart/></c:chartSpace>`,
  });
  const wb = readWorkbook(bytes);

  it("reads every <c:f> with the element it belongs to", () => {
    expect(wb.charts[0]).toEqual({
      part: "xl/charts/chart1.xml",
      kind: "chart",
      sheet: { position: 0, name: "Data" },
      drawing: "xl/drawings/drawing1.xml",
      formulas: [
        { text: "'Data'!$A$1", element: "tx" },
        { text: "Data!$B$1", element: "tx" },
        { text: "'Data'!$A$2:$A$5", element: "cat" },
        { text: "[0]!Sales", element: "val" },
      ],
    });
  });

  it("finds a chartsheet's chart and an Office 2016 chart", () => {
    expect(wb.charts[1]).toEqual({
      part: "xl/charts/chartEx1.xml",
      kind: "chartEx",
      sheet: { position: 1, name: "Chart1" },
      drawing: "xl/drawings/drawing2.xml",
      formulas: [{ text: "'Data'!Local", element: "numDim" }],
    });
  });

  it("lists a chart no drawing places, without a sheet", () => {
    expect(wb.charts[2]).toEqual({ part: "xl/charts/chart9.xml", kind: "chart", formulas: [] });
    expect(wb.warnings).toEqual([]);
  });
});
