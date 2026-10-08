#!/usr/bin/env node
// make-traps.mjs: the seeded-trap workbook of the brief (§7), made WITHOUT Excel.
//
//     node probes/fixtures/make-traps.mjs        (from the repository root)
//
// It patches a copy of an Excel-saved probe file (probes/results/f7_base.xlsx, Excel 16.115
// on macOS) at file level, as probes/filelevel/probe_f7_rename.py does, and writes
// probes/fixtures/traps.xlsx. The content of every part it does not name is kept byte for byte.
// The edits are exact-text insertions at places that occur once in the base file: each
// is checked, so a different base fails loudly instead of producing a broken package.
//
// One trap per audit check (the name or cell, and the check that must catch it):
//
//   BareSeq      SEQUENCE(1,3) stored without _xlfn.                     C2 bare prefix
//   Poisoned     _xludf.SEQUENCE(1,3), as Excel re-saves the above       C2 poisoned
//   Fact         a LAMBDA named like FACT: every Fact(…) calls FACT      C3 (and C10: it can never be called)
//   Broken       'S1'!#REF!+1                                            C4 #REF!
//   ReadsLocal   OnlySecond*2, where OnlySecond exists only on S2        C5
//   S1!A18       Fn(1,2): Fn takes one argument                          C6
//   S1!A20       Opt(1,2,3): Opt takes a and an optional [b]             C6 (S1!A19 Opt(1) is fine)
//   Long         1+1+… over 8,192 characters                             C7
//   Margin_high  Sales_high-Cost_high*1.1 where its family has no *1.1   C11 (Volume_*, Sales_*, Cost_*, Margin_*)
//   CycA, CycB   read each other                                         C12
//   Tax          LAMBDA(x, x*0.27)                                       C13
//   S1!A13       SUM(E1:E3), a fixed reference to the spill E1#          C9 in a cell
//   SpillFixed   'S1'!$E$1:$E$3, the spill's saved extent as a range     C9 in a name (decision 3)
//   Unused       42, read by nothing                                     C10
//   ChartOnly    read only by the chart on S1 (series values [0]!ChartOnly): not unused
//   RelRef       'S1'!A2 without $                                       C4 relative reference (decision 2)
//
// Opening the result in Excel is not needed for the tests. If the author opens it, Excel
// will recalculate on load (fullCalcOnLoad) and may refuse `Long` (probe T15).

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const HERE = dirname(fileURLToPath(import.meta.url));
export const BASE = join(HERE, "..", "results", "f7_base.xlsx");
export const OUT = join(HERE, "traps.xlsx");

const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** Replaces the one occurrence of `find` in `text`; throws if it is not there exactly once. */
function once(text, find, replace, where) {
  const at = text.indexOf(find);
  if (at < 0 || text.indexOf(find, at + 1) >= 0) throw new Error(`${where}: expected exactly one ${JSON.stringify(find.slice(0, 60))}`);
  return text.slice(0, at) + replace + text.slice(at + find.length);
}

// [name, stored definition, localSheetId?]
export const TRAP_NAMES = [
  ["BareSeq", "SEQUENCE(1,3)"],
  ["Broken", "'S1'!#REF!+1"],
  ["ChartOnly", "'S1'!$A$2:$A$11"],
  ["Cost_base", "Sales_base*0.6"],
  ["Cost_high", "Sales_high*0.6"],
  ["CycA", "CycB+1"],
  ["CycB", "CycA+1"],
  ["Fact", "_xlfn.LAMBDA(_xlpm.n, _xlpm.n*2)"],
  ["Long", "1" + "+1".repeat(4150)],
  ["Margin_base", "Sales_base-Cost_base"],
  ["Margin_high", "Sales_high-Cost_high*1.1"],
  ["OnlySecond", "5", 1],
  ["Opt", "_xlfn.LAMBDA(_xlpm.a,_xlop.b, _xlpm.a+IF(_xlfn.ISOMITTED(_xlpm.b),0,_xlpm.b))"],
  ["Poisoned", "_xludf.SEQUENCE(1,3)"],
  ["Price", "2"],
  ["ReadsLocal", "OnlySecond*2"],
  ["RelRef", "'S1'!A2"],
  ["Sales_base", "Volume_base*Price"],
  ["Sales_high", "Volume_high*Price"],
  ["SpillFixed", "'S1'!$E$1:$E$3"],
  ["Tax", "_xlfn.LAMBDA(_xlpm.x, _xlpm.x*0.27)"],
  ["Unused", "42"],
  ["Volume_base", "100"],
  ["Volume_high", "120"],
];

// [cell on S1, stored formula, cached value, value type]
export const TRAP_CELLS = [
  ["A13", "SUM(E1:E3)", "0.6"],
  ["A14", "SUM(BareSeq)", "#NAME?", "e"],
  ["A15", "SUM(Poisoned)", "#NAME?", "e"],
  ["A16", "Broken*2", "#REF!", "e"],
  ["A17", "ReadsLocal", "#NAME?", "e"],
  ["A18", "Fn(1,2)", "#VALUE!", "e"],
  ["A19", "Opt(1)", "1"],
  ["A20", "Opt(1,2,3)", "#VALUE!", "e"],
  ["A21", "Long", "4151"],
  ["A22", "Margin_base+Margin_high", "-6.4"],
  ["A23", "CycA", "0"],
  ["A24", "Tax(100)", "27"],
  ["A25", "SUM(SpillFixed)", "0.6"],
  ["A26", "RelRef", "2"],
];

const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";

const CHART = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}"><c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:numRef><c:f>'S1'!$A$2:$A$11</c:f></c:numRef></c:cat><c:val><c:numRef><c:f>[0]!ChartOnly</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;

const DRAWING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>13</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>16</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;

const rels = (id, type, target) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG_REL}"><Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/></Relationships>`;

/** The traps workbook from the bytes of f7_base.xlsx. */
export function makeTraps(base) {
  const files = unzipSync(base);
  const text = (p) => strFromU8(files[p]);
  const out = { ...files };

  let wb = text("xl/workbook.xml");
  const open = wb.indexOf("<definedNames>");
  const close = wb.indexOf("</definedNames>");
  if (open < 0 || close < 0) throw new Error("workbook.xml: no <definedNames>");
  const existing = wb.slice(open + "<definedNames>".length, close);
  const added = TRAP_NAMES.map(([n, def, lsid]) => ({ n, xml: `<definedName name="${n}"${lsid !== undefined ? ` localSheetId="${lsid}"` : ""}>${esc(def)}</definedName>` }));
  // Excel writes names sorted; keep the base's and slot the new ones in by name.
  const parts = [];
  for (let k = existing.indexOf("<definedName "); k >= 0; ) {
    const end = existing.indexOf("</definedName>", k) + "</definedName>".length;
    const xml = existing.slice(k, end);
    parts.push({ n: xml.slice(xml.indexOf('name="') + 6, xml.indexOf('"', xml.indexOf('name="') + 6)), xml });
    k = existing.indexOf("<definedName ", end);
  }
  const all = [...parts, ...added].sort((a, b) => a.n.toLowerCase().localeCompare(b.n.toLowerCase()) || a.xml.localeCompare(b.xml));
  wb = wb.slice(0, open) + "<definedNames>" + all.map((p) => p.xml).join("") + wb.slice(close);
  wb = once(wb, '<calcPr calcId="181029"/>', '<calcPr calcId="181029" fullCalcOnLoad="1"/>', "workbook.xml");
  out["xl/workbook.xml"] = strToU8(wb);

  let s1 = text("xl/worksheets/sheet1.xml");
  s1 = once(s1, '<dimension ref="A1:E11"/>', '<dimension ref="A1:E26"/>', "sheet1.xml");
  const rows = TRAP_CELLS.map(([cell, f, v, t]) => {
    const r = cell.slice(1);
    return `<row r="${r}" spans="1:1"><c r="${cell}"${t ? ` t="${t}"` : ""}><f>${esc(f)}</f><v>${esc(v)}</v></c></row>`;
  }).join("");
  s1 = once(s1, "</sheetData>", rows + "</sheetData>", "sheet1.xml");
  s1 = once(s1, '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>', '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/>', "sheet1.xml");
  out["xl/worksheets/sheet1.xml"] = strToU8(s1);
  out["xl/worksheets/_rels/sheet1.xml.rels"] = strToU8(rels("rId1", "drawing", "../drawings/drawing1.xml"));
  out["xl/drawings/drawing1.xml"] = strToU8(DRAWING);
  out["xl/drawings/_rels/drawing1.xml.rels"] = strToU8(rels("rId1", "chart", "../charts/chart1.xml"));
  out["xl/charts/chart1.xml"] = strToU8(CHART);

  let s2 = text("xl/worksheets/sheet2.xml");
  s2 = once(s2, '<dimension ref="A1:B5"/>', '<dimension ref="A1:C5"/>', "sheet2.xml");
  s2 = once(s2, "<v>7.2</v></c></row>", '<v>7.2</v></c><c r="C1"><f>OnlySecond</f><v>5</v></c></row>', "sheet2.xml");
  out["xl/worksheets/sheet2.xml"] = strToU8(s2);

  let ct = text("[Content_Types].xml");
  ct = once(
    ct,
    "</Types>",
    '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
      '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
    "[Content_Types].xml",
  );
  out["[Content_Types].xml"] = strToU8(ct);

  // Fixed timestamps: the same base gives the same entries.
  const mtime = new Date(Date.UTC(2026, 9, 4, 12, 0, 0));
  return zipSync(Object.fromEntries(Object.entries(out).map(([k, v]) => [k, [v, { mtime }]])), { level: 6 });
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const bytes = makeTraps(new Uint8Array(readFileSync(BASE)));
  writeFileSync(OUT, bytes);
  console.log(`wrote ${OUT} (${bytes.length} bytes, ${TRAP_NAMES.length} names and ${TRAP_CELLS.length} cells seeded)`);
}
