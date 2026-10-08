#!/usr/bin/env node
// make-afe.mjs: SYNTHETIC workbooks shaped like those Microsoft's Advanced Formula
// Environment (AFE, Excel Labs) saves, made WITHOUT Excel and WITHOUT AFE.
//
//     node probes/fixtures/make-afe.mjs        (from the repository root)
//
// No AFE-saved workbook from a clean Excel save was at hand (2026-10-07). These follow
// the format measured in AFE's own JavaScript bundle and in a third-party AFE-shaped file
// (probes/README.md § AFE-saved workbooks). They are for xln's tests only: do not treat
// them as evidence of what AFE or Excel does. Real AFE-saved files go in probes/results/afe/.
//
// Base: probes/results/f9_lambda_mac.xlsx (Excel 16.115, macOS), whose names include
// ANA.CUBE, ANA.GROW (an "ANA" module), FACT and MAKEADDER (AFE's unprefixed Workbook module).
//
//   afe-synthetic-v11.xlsx  AFE 1.1+: customXml/item1.xml <AFEJSONBlob> (base64 of the
//                           UTF-16LE JSON store, schema afeprojects/0.2) with its itemProps,
//                           relationships and content type, and the add-in's web extension
//                           parts whose setting points at the item's ID. The modules agree
//                           with the workbook's names.
//   afe-synthetic-v10.xlsx  AFE 1.0: a very hidden sheet AFE_hidden_codesheet_49ddb8b8 (its
//                           cell layout is NOT reproduced: one placeholder cell) and the
//                           locale sheet e00eb4de3c8a421cba9b8f4cb8546ec, very hidden, as
//                           older AFE builds left it. Added after Sheet1, so no localSheetId moves.
//
// Every part the script does not name is copied as it was (the zip is rewritten).

import { readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

const HERE = dirname(fileURLToPath(import.meta.url));
export const BASE = join(HERE, "..", "results", "f9_lambda_mac.xlsx");
export const OUT_V11 = join(HERE, "afe-synthetic-v11.xlsx");
export const OUT_V10 = join(HERE, "afe-synthetic-v10.xlsx");

export const AFE_NS = "http://schemas.advancedformulaenvironment.officeapps.live.com/afejsonblob/1.0";
export const AFE_SCHEMA = "http://schemas.advancedformulaenvironment.officeapps.live.com/afeprojects/0.2";
export const ITEM_ID = "{0AFE5E11-0000-4000-8000-5E1F5E1F5E1F}";
const SETTINGS_KEY = "projectV0_1-56c6e055-265e-4713-816e-a646dbb708de";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";

/** The modules of the synthetic store, as AFE would hold them (LF line breaks). */
export const MODULES = [
  {
    path: "/projects/Workbook",
    text:
      "// --- Workbook module ---\n// A file of name definitions of the form:\n//    name = definition;\n\n" +
      "/** Factorial, recursive. */\nFACT = LAMBDA(n, IF(n <= 1, 1, n * FACT(n - 1)));\n\n" +
      "MAKEADDER = LAMBDA(n,\n    LAMBDA(x, x + n)\n);\n",
  },
  {
    path: "/projects/ANA",
    text:
      "// Analytic helpers\n\nCUBE = LAMBDA(x, x ^ 3);\n\n" +
      "/** Compound growth over [periods] (default 1). */\nGROW = LAMBDA(value, rate, [periods],\n    value * (1 + rate) ^ IF(ISOMITTED(periods), 1, periods)\n);\n",
  },
];
export const EXPORTED = ["FACT", "MAKEADDER", "ANA.CUBE", "ANA.GROW"];
const LOCALE = { listSeparator: ",", rowSeparator: ";", columnSeparator: ",", decimalSeparator: ".", statementSeparator: ";", localeName: "en-us" };

/** The text of AFE's custom XML item for a store, as AFE's ExcelJSONStore writes it (no XML declaration). */
export function afeBlobXml(store) {
  const json = JSON.stringify(store);
  const bytes = new Uint8Array(json.length * 2);
  for (let i = 0; i < json.length; i++) {
    const c = json.charCodeAt(i);
    bytes[2 * i] = c & 0xff;
    bytes[2 * i + 1] = c >> 8;
  }
  return `<AFEJSONBlob xmlns="${AFE_NS}">${Buffer.from(bytes).toString("base64")}</AFEJSONBlob>`;
}

export function syntheticStore(modules = MODULES, exported = EXPORTED) {
  return { schema: AFE_SCHEMA, files: modules, projectNames: exported, locale: LOCALE };
}

function once(text, find, replace, where) {
  const at = text.indexOf(find);
  if (at < 0 || text.indexOf(find, at + 1) >= 0) throw new Error(`${where}: expected exactly one ${JSON.stringify(find.slice(0, 60))}`);
  return text.slice(0, at) + replace + text.slice(at + find.length);
}

const rels = (items) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="${PKG_REL}">${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join("")}</Relationships>`;

function pack(out) {
  const mtime = new Date(Date.UTC(2026, 9, 7, 12, 0, 0));
  return zipSync(Object.fromEntries(Object.entries(out).map(([k, v]) => [k, [v, { mtime }]])), { level: 6 });
}

/**
 * Adds AFE 1.1's custom XML part (and, unless `webextension` is false, the add-in's parts
 * that point at it) to a workbook's bytes. `xml` defaults to the synthetic store.
 */
export function addAfePart(base, { xml = afeBlobXml(syntheticStore()), itemId = ITEM_ID, webextension = true } = {}) {
  const files = unzipSync(base);
  const text = (p) => strFromU8(files[p]);
  const out = { ...files };
  let n = 1;
  while (out[`customXml/item${n}.xml`]) n++;
  out[`customXml/item${n}.xml`] = strToU8(xml);
  out[`customXml/itemProps${n}.xml`] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\r\n<ds:datastoreItem ds:itemID="${itemId}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="${AFE_NS}"/></ds:schemaRefs></ds:datastoreItem>`,
  );
  out[`customXml/_rels/item${n}.xml.rels`] = strToU8(rels([["rId1", `${REL}/customXmlProps`, `itemProps${n}.xml`]]));
  let wbRels = text("xl/_rels/workbook.xml.rels");
  wbRels = once(wbRels, "</Relationships>", `<Relationship Id="rIdAfe${n}" Type="${REL}/customXml" Target="../customXml/item${n}.xml"/></Relationships>`, "workbook rels");
  out["xl/_rels/workbook.xml.rels"] = strToU8(wbRels);
  let ct = text("[Content_Types].xml");
  ct = once(ct, "</Types>", `<Override PartName="/customXml/itemProps${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>`, "content types");
  if (webextension) {
    out["xl/webextensions/webextension1.xml"] = strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{5C7C8FEF-A7C0-4F0F-9C7E-38A6C7E37EC8}"><we:reference id="wa200003696" version="1.2.0.0" store="en-US" storeType="OMEX"/><we:alternateReferences><we:reference id="wa200003696" version="1.2.0.0" store="" storeType="OMEX"/></we:alternateReferences><we:properties><we:property name="${SETTINGS_KEY}" value="{&quot;kind&quot;:&quot;AFEJSONBlobNode&quot;,&quot;id&quot;:&quot;${itemId}&quot;}"/></we:properties><we:bindings/><we:snapshot xmlns:r="${REL}"/></we:webextension>`,
    );
    out["xl/webextensions/taskpanes.xml"] = strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"><wetp:taskpane dockstate="right" visibility="0" width="350" row="1"><wetp:webextensionref xmlns:r="${REL}" r:id="rId1"/></wetp:taskpane></wetp:taskpanes>`,
    );
    out["xl/webextensions/_rels/taskpanes.xml.rels"] = strToU8(rels([["rId1", "http://schemas.microsoft.com/office/2011/relationships/webextension", "webextension1.xml"]]));
    let root = text("_rels/.rels");
    root = once(root, "</Relationships>", `<Relationship Id="rIdAfeTp" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="xl/webextensions/taskpanes.xml"/></Relationships>`, "package rels");
    out["_rels/.rels"] = strToU8(root);
    ct = once(
      ct,
      "</Types>",
      `<Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/><Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/></Types>`,
      "content types",
    );
  }
  out["[Content_Types].xml"] = strToU8(ct);
  return pack(out);
}

/** Adds very hidden sheets with the given names after the last sheet. */
export function addHiddenSheets(base, names) {
  const files = unzipSync(base);
  const text = (p) => strFromU8(files[p]);
  const out = { ...files };
  let wb = text("xl/workbook.xml");
  let wbRels = text("xl/_rels/workbook.xml.rels");
  let ct = text("[Content_Types].xml");
  names.forEach((name, i) => {
    let n = 1;
    while (out[`xl/worksheets/sheet${n}.xml`]) n++;
    const cell = i === 0 ? `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>synthetic: AFE 1.0 code sheet, layout not reproduced</t></is></c></row></sheetData>` : "<sheetData/>";
    out[`xl/worksheets/sheet${n}.xml`] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><dimension ref="A1"/>${cell}</worksheet>`);
    wbRels = once(wbRels, "</Relationships>", `<Relationship Id="rIdAfeS${n}" Type="${REL}/worksheet" Target="worksheets/sheet${n}.xml"/></Relationships>`, "workbook rels");
    wb = once(wb, "</sheets>", `<sheet name="${name}" sheetId="${100 + n}" state="veryHidden" r:id="rIdAfeS${n}"/></sheets>`, "workbook");
    ct = once(ct, "</Types>", `<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`, "content types");
  });
  out["xl/workbook.xml"] = strToU8(wb);
  out["xl/_rels/workbook.xml.rels"] = strToU8(wbRels);
  out["[Content_Types].xml"] = strToU8(ct);
  return pack(out);
}

export function makeV11(base) {
  return addAfePart(base);
}

export function makeV10(base) {
  return addHiddenSheets(base, ["AFE_hidden_codesheet_49ddb8b8", "e00eb4de3c8a421cba9b8f4cb8546ec"]);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const base = new Uint8Array(readFileSync(BASE));
  writeFileSync(OUT_V11, makeV11(base));
  writeFileSync(OUT_V10, makeV10(base));
  console.log(`wrote ${OUT_V11} and ${OUT_V10} (synthetic)`);
}
