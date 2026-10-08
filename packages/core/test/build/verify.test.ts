// E4: cached values compared cell by cell.
import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { strToU8, zipSync } from "fflate";
import { readCellValues, replaceZipEntries, utf8, valuesOrigin, valuesWarning, valueText, verifyValues } from "../../src/index.js";
import { decodeText, Package } from "../../src/file/package.js";
import { build, edit, fixture, pulled } from "./helpers.js";

const F7 = fixture("f7_base.xlsx");

describe("verify", () => {
  it("reads every cell value, spill cells included", () => {
    const s1 = readCellValues(F7).find((s) => s.sheet === "S1")!;
    expect(s1.cells.get("E2")).toBe(0.2); // a spill cell beyond its anchor (SEQUENCE(3)*Rate)
    expect(typeof s1.cells.get("A1")).toBe("number");
  });

  it("a build of names changes no cached value", () => {
    const files = pulled(F7);
    edit(files, "names/_unmanaged.xln", "RateX = 0.5;", "/** doc */\nRateX = 0.5;");
    const r = build(F7, files);
    const v = verifyValues(F7, r.bytes!);
    expect(v.changed).toEqual([]);
    expect(v.cells).toBeGreaterThan(20);
  });

  it("F7: Excel's recalculation of the patched rename changed no value", () => {
    const v = verifyValues(F7, fixture("f7_patched_resaved.xlsx"));
    expect(v.changed.map((c) => `${c.sheet}!${c.cell}: ${valueText(c.before)} → ${valueText(c.after)}`)).toEqual([]);
  });

  it("reports a changed cell", () => {
    const pkg = new Package(F7);
    const xml = decodeText(pkg.raw("xl/worksheets/sheet1.xml")!);
    const at = xml.indexOf("<v>", xml.indexOf('r="A1"'));
    const end = xml.indexOf("</v>", at);
    const changed = replaceZipEntries(F7, new Map([["xl/worksheets/sheet1.xml", utf8(xml.slice(0, at) + "<v>999" + xml.slice(end))]]));
    const v = verifyValues(F7, changed);
    expect(v.changed).toMatchObject([{ sheet: "S1", cell: "A1", after: 999 }]);
    expect(verifyValues(F7, changed, { tolerance: 1e9 }).changed).toEqual([]);
  });
});

describe("values Excel calculated (valuesOrigin)", () => {
  const RESULTS = join(import.meta.dirname, "..", "..", "..", "..", "probes", "results");

  it("every Excel-saved fixture is excel; the files xln wrote are partial or none", () => {
    const got = Object.fromEntries(
      readdirSync(RESULTS)
        .filter((f) => f.endsWith(".xlsx"))
        .sort()
        .map((f) => [f, valuesOrigin(new Uint8Array(readFileSync(join(RESULTS, f)))).origin]),
    );
    // f8_p1, f8_p2, f8_q1: cell formulas rewritten at file level, never opened since.
    const notExcel = Object.entries(got).filter(([, o]) => o !== "excel");
    expect(notExcel).toEqual([["f8_p1.xlsx", "partial"], ["f8_p2.xlsx", "partial"], ["f8_q1.xlsx", "none"]]);
    expect(valuesOrigin(fixture("f8_p1.xlsx"))).toEqual({ origin: "partial", formulaCells: 49, withoutValue: 26, fullCalcOnLoad: true });
  });

  it("a placeholder 0 on every formula cell is no value only with fullCalcOnLoad (XlsxWriter's)", () => {
    // As XlsxWriter writes them: 0 on a number formula, "" on a text formula.
    const M = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    const T = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const rel = (id: string, type: string, target: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="${id}" Type="${T}/${type}" Target="${target}"/></Relationships>`;
    const book = (calcPr: string) =>
      zipSync({
        "_rels/.rels": strToU8(rel("rId1", "officeDocument", "xl/workbook.xml")),
        "xl/workbook.xml": strToU8(`<workbook ${M} xmlns:r="${T}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>${calcPr}</workbook>`),
        "xl/_rels/workbook.xml.rels": strToU8(rel("rId1", "worksheet", "worksheets/sheet1.xml")),
        "xl/worksheets/sheet1.xml": strToU8(`<worksheet ${M}><sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1-2</f><v>0</v></c><c r="C1" t="str"><f>""&amp;""</f><v></v></c></row></sheetData></worksheet>`),
      });
    expect(valuesOrigin(book('<calcPr calcId="124519" fullCalcOnLoad="1"/>'))).toEqual({ origin: "none", formulaCells: 2, withoutValue: 1, fullCalcOnLoad: true });
    // Without fullCalcOnLoad the 0 and the "" are Excel's.
    expect(valuesOrigin(book('<calcPr calcId="191029"/>')).origin).toBe("excel");
    expect(valuesWarning("original.xlsx", valuesOrigin(fixture("f8_p1.xlsx")), "before")).toBe(
      "original.xlsx was written after Excel last saved it: 26 of 49 formula cells have no value Excel calculated, and compare as empty; open and save it in Excel first",
    );
  });
});

describe.skipIf(!process.env["XLN_CORPUS"])("values Excel calculated on the corpus (XLN_CORPUS)", () => {
  it("the Python-written workbooks have none; the ones saved by Excel have Excel's", () => {
    const root = process.env["XLN_CORPUS"]!;
    const got: Record<string, string> = {};
    for (const d of readdirSync(root, { withFileTypes: true })) {
      const dist = join(root, d.name, "dist");
      if (!d.isDirectory() || !existsSync(dist)) continue;
      for (const f of readdirSync(dist)) if (f.endsWith(".xlsx") && !f.startsWith("~$")) got[f] = valuesOrigin(new Uint8Array(readFileSync(join(dist, f)))).origin;
    }
    expect(Object.keys(got).length).toBeGreaterThan(0);
    // XlsxWriter's signature (AppVersion 12.0000, calcId 124519, fullCalcOnLoad, 0 or "" on every formula cell).
    for (const f of ["excel-layers-demo.xlsx", "lbo-ep03.xlsx", "lbo-ep03r.xlsx", "lbo-ep03r-circ.xlsx"]) if (f in got) expect(got[f], f).toBe("none");
    for (const f of ["excel-layers-demo-luca.xlsx", "excel-layers-demo-v3.xlsx", "lbo-ep02.xlsx"]) if (f in got) expect(got[f], f).toBe("excel");
  });
});
