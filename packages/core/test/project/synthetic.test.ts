// A hand-built workbook for what the probe files lack: hidden names, built-ins, helper
// names, multi-line comments, CR LF definitions, Tables, a sheet name needing quotes,
// underscore modules with a sheet-scoped member, a sheet name that is no valid file name,
// and a definition that does not parse.
import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { LOCK_FILE, MANIFEST_FILE, parseModule, parseSourceFile, pullProject } from "../../src/index.js";
import { checkProject } from "./check.js";

const M = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const PR = "http://schemas.openxmlformats.org/package/2006/relationships";
const T = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function rels(entries: [string, string, string][]): string {
  return `<Relationships xmlns="${PR}">${entries.map(([id, type, target]) => `<Relationship Id="${id}" Type="${T}/${type}" Target="${target}"/>`).join("")}</Relationships>`;
}

const bytes = zipSync(
  Object.fromEntries(
    Object.entries({
      "_rels/.rels": rels([["rId1", "officeDocument", "xl/workbook.xml"]]),
      "xl/workbook.xml": `<workbook ${M} ${R}><sheets>
        <sheet name="Data" sheetId="1" r:id="rId1"/>
        <sheet name="Cash Flow" sheetId="2" state="hidden" r:id="rId2"/>
        <sheet name="x&lt;y&gt;|z." sheetId="3" r:id="rId3"/>
      </sheets><definedNames>
        <definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$C$3</definedName>
        <definedName name="_xlfn.SINGLE" hidden="1">#NAME?</definedName>
        <definedName name="Secret" hidden="1">42</definedName>
        <definedName name="Doc" comment="Line one&#13;&#10;  line two with */ inside">_xlfn.LET(_xlpm.a,&#13;&#10;  1,&#13;&#10;  _xlpm.a+&quot;x&#13;&#10;y&quot;)</definedName>
        <definedName name="IN_Rate">0.05</definedName>
        <definedName name="IN_Tax">0.3</definedName>
        <definedName name="IN_Years">tblIn[Year]</definedName>
        <definedName name="Rev" localSheetId="1">'Cash Flow'!$B$2:$F$2</definedName>
        <definedName name="Rev">Data!$B$2</definedName>
        <definedName name="Total" localSheetId="1">SUM('Cash Flow'!Rev)</definedName>
        <definedName name="Broken">SUM(1,</definedName>
        <definedName name="IN_Local" localSheetId="0">Data!$A$1</definedName>
        <definedName name="Odd" localSheetId="2">1</definedName>
        <definedName name="Poisoned">_xludf.XLOOKUP(1,Data!A1:A3,Data!B1:B3)</definedName>
      </definedNames></workbook>`,
      "xl/_rels/workbook.xml.rels": rels([
        ["rId1", "worksheet", "worksheets/sheet1.xml"],
        ["rId2", "worksheet", "worksheets/sheet2.xml"],
        ["rId3", "worksheet", "worksheets/sheet3.xml"],
      ]),
      "xl/worksheets/sheet1.xml": `<worksheet ${M} ${R}><sheetData>
        <row r="1"><c r="A1"><f>Rev*IN_Rate</f><v>1</v></c><c r="B1"><f>'Cash Flow'!Total</f><v>1</v></c></row>
      </sheetData><tableParts count="1"><tablePart r:id="rId9"/></tableParts></worksheet>`,
      "xl/worksheets/_rels/sheet1.xml.rels": rels([["rId9", "table", "../tables/table1.xml"]]),
      "xl/tables/table1.xml": `<table ${M} id="1" name="tblIn" displayName="tblIn" ref="A10:B12"><tableColumns count="2">
        <tableColumn id="1" name="Year"/><tableColumn id="2" name="Taxed"><calculatedColumnFormula>tblIn[[#This Row],[Year]]*IN_Tax</calculatedColumnFormula></tableColumn>
      </tableColumns></table>`,
      "xl/worksheets/sheet3.xml": `<worksheet ${M}><sheetData/></worksheet>`,
      "xl/worksheets/sheet2.xml": `<worksheet ${M}><sheetData><row r="1"><c r="A1"><f>Rev+Total</f><v>1</v></c></row></sheetData></worksheet>`,
    }).map(([k, v]) => [k, strToU8(v)]),
  ),
);

describe("synthetic workbook", () => {
  const r = pullProject(bytes, "C:\\models\\synthetic.xlsx");
  const m = JSON.parse(r.files[MANIFEST_FILE]!);

  it("writes every regular name once, loss-free", () => {
    const c = checkProject(r);
    expect(c.failures).toEqual([]);
    expect(r.report.workbook).toBe("synthetic.xlsx");
    expect(r.report.names).toBe(12);
    expect(r.report.builtIns).toEqual(["Data!_xlnm.Print_Area"]);
    expect(r.report.helpersIgnored).toEqual(["_xlfn.SINGLE"]);
    expect(r.report.hidden).toBe(1);
  });

  it("classifies", () => {
    const kind = Object.fromEntries(r.names.map((n) => [n.scope ? `${n.scope}!${n.name}` : n.name, n.classification.kind]));
    expect(kind).toEqual({
      Broken: "unparsed",
      Doc: "formula",
      IN_Rate: "constant",
      IN_Tax: "constant",
      "Data!IN_Local": "range",
      "x<y>|z.!Odd": "constant",
      IN_Years: "table",
      Poisoned: "formula",
      Rev: "range",
      Secret: "constant",
      "Cash Flow!Rev": "range",
      "Cash Flow!Total": "formula",
    });
    expect(r.report.warnings.some((w) => w.startsWith("Broken: does not parse"))).toBe(true);
    expect(r.report.warnings.some((w) => w.startsWith("Poisoned:") && w.includes("user-defined"))).toBe(true);
  });

  it("groups IN_ into a module, workbook names into _unmanaged, sheet names into one file per sheet", () => {
    // IN_Local (on Data!A1, a formula cell) and Rev (on Data!B2, an empty cell) are cell
    // statements: they live in Data's file, not in their module or _unmanaged.xln.
    expect(r.report.modules.map((x) => [x.module, x.names, x.workbookScoped, x.sheetScoped])).toEqual([
      ["IN", 3, 3, 0],
      ["_unmanaged", 4, 4, 0],
    ]);
    const none = { named: 0, slots: 0, unnamed: 0, blocks: 0, blockCells: 0 };
    expect(r.report.sheetFiles).toEqual([
      { sheet: "Data", file: "names/sheets/Data.xln", names: 2, cells: { ...none, named: 1, slots: 1, unnamed: 1 } },
      { sheet: "Cash Flow", file: "names/sheets/Cash Flow.xln", names: 2, cells: { ...none, unnamed: 1 } },
      { sheet: "x<y>|z.", file: "names/sheets/x%3Cy%3E%7Cz%2E.xln", names: 1, cells: none },
    ]);
    expect(r.report.unmanaged).toBe(4);
    expect(Object.keys(r.files).filter((f) => f.endsWith(".xln"))).toEqual([
      "names/IN.xln",
      "names/_unmanaged.xln",
      "names/sheets/Data.xln",
      "names/sheets/Cash Flow.xln",
      "names/sheets/x%3Cy%3E%7Cz%2E.xln",
    ]);
    expect(r.files["names/IN.xln"]).toBe("// module: IN, pulled by xln from synthetic.xlsx.\n\nIN_Rate = 0.05;\nIN_Tax = 0.3;\nIN_Years = tblIn[Year];\n");
    // Cells in sheet order; the workbook-scoped slot with @workbook above it, its address bare (M3d).
    expect(r.files["names/sheets/Data.xln"]).toContain("\n\nIN_Local @A1 = Rev*IN_Rate;\n@B1 = 'Cash Flow'!Total;\n@workbook\nRev @B2 = ;\n");
    // Sheet files: no blocks; the file's name gives the sheet.
    const cf = r.files["names/sheets/Cash Flow.xln"]!;
    expect(cf.split("\n").find((l) => l !== "" && !l.startsWith("//"))).toBe("@A1 = Rev+Total;");
    expect(cf).toContain("\n\n@A1 = Rev+Total;\n\n// Other names on Cash Flow.\n\nRev = $B$2:$F$2;\nTotal = SUM(Rev);\n");
    expect(cf.match(/^@(scope|workbook)/gm)).toBeNull();
    expect(r.files["names/sheets/x%3Cy%3E%7Cz%2E.xln"]).toContain("\n\nOdd = 1;\n");
    expect(parseSourceFile("names/sheets/x%3Cy%3E%7Cz%2E.xln", r.files["names/sheets/x%3Cy%3E%7Cz%2E.xln"]!).entries[0]!.scope).toBe("x<y>|z.");
    expect(parseSourceFile("names/sheets/Cash Flow.xln", cf).entries.map((e) => [e.name, e.scope])).toEqual([
      ["Rev", "Cash Flow"],
      ["Total", "Cash Flow"],
    ]);
    const u = r.files["names/_unmanaged.xln"]!;
    expect(u).not.toContain("@scope");
    expect(u).toContain("@hidden\nSecret = 42;");
    expect(u).toContain("Broken = SUM(1,;");
    // CR LF in layout becomes LF; the CR LF inside the string is the string's value and stays.
    expect(u).toContain('/**\n * Line one\n *   line two with *\\/ inside\n */\nDoc = LET(a,\n  1,\n  a+"x\r\ny");');
    const doc = parseModule(u).entries.find((e) => e.name === "Doc")!;
    expect(doc.doc).toBe("Line one\n  line two with */ inside");
  });

  it("indexes usage across sheet scopes and Table columns", () => {
    expect(m.names["Rev"].usedBy).toEqual({ cells: { Data: ["A1"] } });
    expect(m.names["Cash Flow!Rev"].usedBy).toEqual({ names: ["Cash Flow!Total"], cells: { "Cash Flow": ["A1"] } });
    expect(m.names["Cash Flow!Total"].usedBy.cells).toEqual({ Data: ["B1"], "Cash Flow": ["A1"] });
    expect(m.names["IN_Tax"].usedBy).toEqual({ tableColumns: ["tblIn[Taxed]"] });
    expect(m.tables).toEqual([
      { name: "tblIn", sheet: "Data", ref: "A10:B12", headerRowCount: 1, totalsRowCount: 0, columns: ["Year", "Taxed"], calculatedColumns: { Taxed: "tblIn[[#This Row],[Year]]*IN_Tax" } },
    ]);
    expect(m.sheets.map((s: { name: string; state: string }) => s.state)).toEqual(["visible", "hidden", "visible"]);
    expect(m.names["Cash Flow!Total"].file).toBe("names/sheets/Cash Flow.xln");
    expect(m.names["Data!IN_Local"]).toMatchObject({ module: "IN", file: "names/sheets/Data.xln", cell: { sheet: "Data", range: "A1" } });
    expect(m.unnamedCells).toEqual({ Data: ["B1"], "Cash Flow": ["A1"] });
    expect(m.builtInNames).toEqual([{ name: "_xlnm.Print_Area", scope: "Data", definition: "Data!$A$1:$C$3" }]);
  });

  it("locks hidden flags and comments", () => {
    const lock = JSON.parse(r.files[LOCK_FILE]!);
    expect(lock.names["Secret"].hidden).toBe(true);
    expect(lock.names["Doc"].comment).not.toBeNull();
    expect(Object.keys(lock.names)).toHaveLength(12);
  });
});

describe("sheet files that clash on a case- or normalisation-insensitive file system", () => {
  // Excel itself refuses "BS" next to "bs", but not "Café" in NFC next to "Café" in NFD,
  // which APFS takes for the same file. The reader polices neither.
  const sheets = ["BS", "bs", "Café", "Café"];
  const wb = zipSync(
    Object.fromEntries(
      Object.entries({
        "_rels/.rels": rels([["rId1", "officeDocument", "xl/workbook.xml"]]),
        "xl/workbook.xml": `<workbook ${M} ${R}><sheets>${sheets.map((s, k) => `<sheet name="${s}" sheetId="${k + 1}" r:id="rId${k + 1}"/>`).join("")}</sheets>
          <definedNames>${sheets.map((_, k) => `<definedName name="X" localSheetId="${k}">${k}</definedName>`).join("")}</definedNames></workbook>`,
        "xl/_rels/workbook.xml.rels": rels(sheets.map((_, k): [string, string, string] => [`rId${k + 1}`, "worksheet", `worksheets/sheet${k + 1}.xml`])),
        ...Object.fromEntries(sheets.map((_, k) => [`xl/worksheets/sheet${k + 1}.xml`, `<worksheet ${M}><sheetData/></worksheet>`])),
      }).map(([k, v]) => [k, strToU8(v)]),
    ),
  );
  const r = pullProject(wb, "clash.xlsx");

  it("gives each sheet its own file; a file name that cannot say its sheet carries @sheet(Name)", () => {
    expect(r.report.sheetFiles.map((f) => f.file)).toEqual([
      "names/sheets/BS.xln",
      "names/sheets/bs~2.xln",
      "names/sheets/Café.xln",
      "names/sheets/Café~2.xln",
    ]);
    expect(checkProject(r).failures).toEqual([]);
    for (const [k, f] of r.report.sheetFiles.entries()) {
      expect(parseSourceFile(f.file, r.files[f.file]!).entries).toMatchObject([{ name: "X", scope: sheets[k], formula: String(k) }]);
      expect(r.files[f.file]!.includes("@sheet(")).toBe(f.file.includes("~"));
    }
  });
});
