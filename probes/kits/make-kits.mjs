#!/usr/bin/env node
// make-kits.mjs: the workbooks of probes F10 and F11 (author's steps: probes/kits/F10-F11.md).
//
//     node probes/kits/make-kits.mjs        (from the repository root)
//
// Writes, WITHOUT Excel, three small packages with the minimum parts Excel needs:
//   kits/f10/f10_kit.xlsx    Sheet1!A3:A6 = 10, 20, 30, 40, the rest of A1:A10 empty, so the
//                            three trim ranges over A1:A10 give 4, 6 and 8 rows
//   kits/f10/Other.xlsx      Sheet1!A1 = 42, A2 = 7, the name OtherVal = Sheet1!$A$2
//   kits/f11/f11_kit.xlsx    sheet Labels: the labels of f11/labels.json in column A, the row
//                            number in column B (the value each name should cover)
// The F11 analysis test (packages/core/test/file/probe-f11.test.ts) reads the same list.

import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strToU8, zipSync } from "fflate";

const HERE = dirname(fileURLToPath(import.meta.url));
// The labels and the names assumed for them, shared with the analysis test.
const F11_LABELS = JSON.parse(readFileSync(join(HERE, "f11", "labels.json"), "utf8"));
const labelText = (l) => (l.padTo ? l.label.padEnd(l.padTo, "x") : l.label);
const esc = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** A cell: a number (`style` 1 is a short date), or an inline string kept with its spaces. */
function cell(ref, v, style) {
  const s = style ? ` s="${style}"` : "";
  if (typeof v === "number") return `<c r="${ref}"${s}><v>${v}</v></c>`;
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
}

/** rows: [rowNumber, [ [col, value, style?], … ]] */
function workbook(sheet, rows, names = []) {
  const sheetData = rows.map(([r, cells]) => `<row r="${r}">${cells.map(([c, v, st]) => cell(`${c}${r}`, v, st)).join("")}</row>`).join("");
  const defined = names.length ? `<definedNames>${names.map(([n, d]) => `<definedName name="${esc(n)}">${esc(d)}</definedName>`).join("")}</definedNames>` : "";
  const files = {
    "[Content_Types].xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    "_rels/.rels":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${MAIN}" xmlns:r="${REL}">` +
      `<sheets><sheet name="${esc(sheet)}" sheetId="1" r:id="rId1"/></sheets>${defined}<calcPr calcId="181029" fullCalcOnLoad="1"/></workbook>`,
    "xl/_rels/workbook.xml.rels":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`,
    "xl/styles.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${MAIN}">` +
      `<fonts count="1"><font><sz val="12"/><name val="Calibri"/></font></fonts>` +
      `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
      `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>` +
      `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
    "xl/worksheets/sheet1.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${MAIN}" xmlns:r="${REL}">` +
      `<cols><col min="1" max="1" width="28" customWidth="1"/></cols><sheetData>${sheetData}</sheetData></worksheet>`,
  };
  return zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)])));
}

function write(rel, bytes) {
  const path = join(HERE, rel);
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, bytes);
  console.log(`wrote probes/kits/${rel}`);
}

write("f10/f10_kit.xlsx", workbook("Sheet1", [3, 4, 5, 6].map((r) => [r, [["A", (r - 2) * 10]]])));
write("f10/Other.xlsx", workbook("Sheet1", [[1, [["A", 42]]], [2, [["A", 7]]]], [["OtherVal", "Sheet1!$A$2"]]));
write(
  "f11/f11_kit.xlsx",
  workbook(
    "Labels",
    F11_LABELS.map((l) => [l.row, [...(l.label === "" ? [] : [["A", labelText(l), l.date ? 1 : undefined]]), ["B", l.row]]]),
  ),
);
