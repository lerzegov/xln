// AFE coexistence: the module store of Microsoft's Advanced Formula Environment (Excel
// Labs) in a workbook is recognised, reported (pull, check C14, build) and never changed.
//
// Fixtures: probes/fixtures/afe-synthetic-v1{0,1}.xlsx are SYNTHETIC (make-afe.mjs: no
// Excel, no AFE), shaped by the format measured in AFE's bundle. Real AFE-saved files go in
// probes/results/afe/*.xlsx; the last block runs on each when there are any.
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import {
  AFE_BLOB_NS,
  afeStatus,
  audit,
  buildWorkbook,
  parseAfeBlob,
  pullProject,
  rawZipRecords,
  readWorkbook,
  type BuildResult,
  type WorkbookSnapshot,
} from "../../src/index.js";
// @ts-ignore: a plain ES module, run by node to make the fixtures
import { addAfePart, afeBlobXml, BASE, makeV10, makeV11, OUT_V10, OUT_V11, syntheticStore, MODULES, EXPORTED, ITEM_ID } from "../../../../probes/fixtures/make-afe.mjs";
import { edit } from "../build/helpers.js";

const REAL = join(import.meta.dirname, "..", "..", "..", "..", "probes", "results", "afe");
const v11 = new Uint8Array(readFileSync(OUT_V11 as string));
const v10 = new Uint8Array(readFileSync(OUT_V10 as string));
const base = new Uint8Array(readFileSync(BASE as string));

function build(bytes: Uint8Array, files: Record<string, string>, embed = false): BuildResult {
  return buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files }, { embed, provenance: false });
}

/** The zip records (header and compressed data) of the parts AFE owns. */
function afeRecords(bytes: Uint8Array): Map<string, Uint8Array> {
  const all = rawZipRecords(bytes);
  return new Map([...all].filter(([n]) => n.startsWith("customXml/") || n.startsWith("xl/webextensions/") || n === "xl/worksheets/sheet2.xml" || n === "xl/worksheets/sheet3.xml"));
}

function sameRecords(a: Map<string, Uint8Array>, b: Map<string, Uint8Array>, only?: (n: string) => boolean): string[] {
  const out: string[] = [];
  for (const [n, r] of a) {
    if (only && !only(n)) continue;
    const s = b.get(n);
    if (!s || s.length !== r.length || !s.every((x, i) => x === r[i])) out.push(n);
  }
  return out;
}

const withModules = (modules: { path: string; text: string }[], exported = EXPORTED as string[], locale?: Record<string, string>) => {
  const store = syntheticStore(modules, exported);
  if (locale) store.locale = locale;
  return addAfePart(base, { xml: afeBlobXml(store) }) as Uint8Array;
};

describe("synthetic fixtures", () => {
  it("are what make-afe.mjs makes from f9_lambda_mac.xlsx", () => {
    for (const [made, committed] of [
      [makeV11(base), v11],
      [makeV10(base), v10],
    ] as [Uint8Array, Uint8Array][]) {
      const a = unzipSync(made);
      const b = unzipSync(committed);
      expect(Object.keys(a).sort()).toEqual(Object.keys(b).sort());
      for (const k of Object.keys(b)) expect(Buffer.from(a[k]!).equals(Buffer.from(b[k]!)), k).toBe(true);
    }
  });
});

describe("reading AFE's store", () => {
  it("finds AFE 1.1's custom XML part by its namespace, with modules, exported names, locale and the add-in's link", () => {
    const wb = readWorkbook(v11);
    expect(wb.warnings).toEqual([]);
    expect(wb.foreignModuleStores).toHaveLength(1);
    const s = wb.foreignModuleStores[0]!;
    expect(s).toMatchObject({ tool: "afe", kind: "custom-xml", part: "customXml/item1.xml", namespace: AFE_BLOB_NS, schema: "http://schemas.advancedformulaenvironment.officeapps.live.com/afeprojects/0.2", itemId: ITEM_ID, linked: true });
    expect(s.modules!.map((m) => [m.name, m.path])).toEqual([
      ["Workbook", "/projects/Workbook"],
      ["ANA", "/projects/ANA"],
    ]);
    expect(s.modules![1]!.text).toBe((MODULES as { text: string }[])[1]!.text);
    expect(s.exportedNames).toEqual(EXPORTED);
    expect(s.locale).toEqual({ listSeparator: ",", decimalSeparator: ".", localeName: "en-us" });
    expect(s.unreadable).toBeUndefined();
  });

  it("finds AFE 1.0's code sheet and the locale sheet by their names", () => {
    const wb = readWorkbook(v10);
    expect(wb.foreignModuleStores.map((s) => [s.kind, s.sheet, s.state])).toEqual([
      ["code-sheet", "AFE_hidden_codesheet_49ddb8b8", "veryHidden"],
      ["locale-sheet", "e00eb4de3c8a421cba9b8f4cb8546ec", "veryHidden"],
    ]);
    expect(wb.foreignModuleStores[0]!.unreadable).toContain("AFE 1.0");
  });

  it("finds nothing in a workbook without AFE, nor in xln's own embedded part", () => {
    expect(readWorkbook(base).foreignModuleStores).toEqual([]);
    const files = { ...pullProject(base, "book.xlsx").files };
    const r = build(base, files, true);
    expect(r.status).toBe("built");
    expect(readWorkbook(r.bytes!).foreignModuleStores).toEqual([]);
  });

  it("says why a store cannot be read, and still recognises it as AFE's", () => {
    expect(parseAfeBlob(`<AFEJSONBlob xmlns="${AFE_BLOB_NS}">not base64!</AFEJSONBlob>`).unreadable).toContain("not base64");
    expect(parseAfeBlob(`<AFEJSONBlob xmlns="${AFE_BLOB_NS}">QUJD</AFEJSONBlob>`).unreadable).toContain("not the JSON");
    const other = afeBlobXml({ schema: "http://schemas.advancedformulaenvironment.officeapps.live.com/afeprojects/9.9", files: [] });
    expect(parseAfeBlob(other).unreadable).toContain("afeprojects/9.9");
    const later = `<AFEStore xmlns="http://schemas.advancedformulaenvironment.officeapps.live.com/afestore/2.0">x</AFEStore>`;
    const wb = readWorkbook(addAfePart(base, { xml: later, webextension: false }));
    expect(wb.foreignModuleStores).toHaveLength(1);
    expect(wb.foreignModuleStores[0]!.unreadable).toContain("later AFE format");
    expect(wb.foreignModuleStores[0]!.linked).toBe(false);
    const c14 = audit(wb, { only: ["C14"] }).findings.map((f) => f.rule);
    expect(c14).toEqual(["C14.afe-unreadable"]);
  });
});

describe("check C14", () => {
  const c14 = (wb: WorkbookSnapshot) => audit(wb, { only: ["C14"] }).findings;

  it("reports the store and its modules as info, and no drift when AFE's text and the names agree", () => {
    const f = c14(readWorkbook(v11));
    expect(f.map((x) => `${x.rule} ${x.severity} ${x.where.kind} ${x.where.ref}`)).toEqual(["C14.afe-store info part customXml/item1.xml"]);
    expect(f[0]!.message).toContain("4 names as module text here, modules Workbook (2), ANA (2)");
    expect(afeStatus(readWorkbook(v11))[0]!.entries.map((e) => `${e.name} ${e.state}`)).toEqual(["FACT same", "MAKEADDER same", "ANA.CUBE same", "ANA.GROW same"]);
  });

  it("notes (info) a name whose AFE text differs, on the name; lists AFE's names the workbook lacks", () => {
    const wb = readWorkbook(withModules([{ path: "/projects/ANA", text: "CUBE = LAMBDA(x, x ^ 4);\nSQUARE = LAMBDA(x, x ^ 2);\nGROW = LAMBDA(value, rate, [periods], value * (1 + rate) ^ IF(ISOMITTED(periods), 1, periods));" }]));
    const f = c14(wb);
    expect(f.map((x) => `${x.rule} ${x.where.kind === "name" ? x.where.key : x.where.ref}`)).toEqual(["C14.afe-absent customXml/item1.xml", "C14.afe-store customXml/item1.xml", "C14.afe-drift ANA.CUBE"]);
    const drift = f.find((x) => x.rule === "C14.afe-drift")!;
    expect(drift.severity).toBe("info");
    expect(drift.message).toBe("AFE's module ANA has LAMBDA(x, x ^ 4); the workbook has LAMBDA(x, x ^ 3)");
    expect(f[0]!.message).toBe("AFE's module ANA defines 1 name the workbook does not have: ANA.SQUARE");
  });

  it("compares modulo layout, case and number spelling, with siblings named without their module", () => {
    const wb = readWorkbook(withModules([{ path: "/projects/ANA", text: "cube = lambda(X, X^3.0);\nGROW = LAMBDA(value,rate,[periods],\n  value*(1+rate)^if(isomitted(periods),1,periods));\nTWICE = LAMBDA(x, CUBE(x) * 2);" }]));
    // The workbook's own ANA.TWICE calls ANA.CUBE by its full name.
    const twice = { ...wb.definedNames.find((d) => d.name === "ANA.CUBE")!, name: "ANA.TWICE", definition: "_xlfn.LAMBDA(_xlpm.x, ANA.CUBE(_xlpm.x) * 2)" };
    const st = afeStatus({ ...wb, definedNames: [...wb.definedNames, twice] })[0]!;
    expect(st.entries.map((e) => `${e.name} ${e.state}`)).toEqual(["ANA.CUBE same", "ANA.GROW same", "ANA.TWICE same"]);
  });

  it("does not compare text written with other separators", () => {
    const wb = readWorkbook(withModules([{ path: "/projects/ANA", text: "CUBE = LAMBDA(x; x ^ 3);" }], ["ANA.CUBE"], { listSeparator: ";", decimalSeparator: ",", localeName: "it-it" }));
    const f = c14(wb);
    expect(f.map((x) => x.rule)).toEqual(["C14.afe-not-compared", "C14.afe-store"]);
    expect(afeStatus(wb)[0]!.entries[0]!.state).toBe("not-compared");
  });

  it("notes AFE 1.0's code sheet and the locale sheet", () => {
    expect(c14(readWorkbook(v10)).map((x) => `${x.rule} ${x.where.sheet}`)).toEqual([
      "C14.afe-code-sheet AFE_hidden_codesheet_49ddb8b8",
      "C14.afe-locale-sheet e00eb4de3c8a421cba9b8f4cb8546ec",
    ]);
  });

  it("can be turned off like any check", () => {
    expect(audit(readWorkbook(v11), { rules: { C14: "off" } }).findings.filter((f) => f.check === "C14")).toEqual([]);
  });
});

describe("pull", () => {
  it("reports AFE's store in the report and a note, and reads the names from the Name Manager only", () => {
    const p = pullProject(v11, "book.xlsx");
    expect(p.report.foreignModules).toEqual([
      { tool: "afe", kind: "custom-xml", part: "customXml/item1.xml", modules: [{ name: "Workbook", names: 2 }, { name: "ANA", names: 2 }], differs: [], absent: [] },
    ]);
    expect(p.report.notes.some((n) => n.includes("Advanced Formula Environment") && n.includes("Workbook (2), ANA (2)"))).toBe(true);
    // The same project as from the workbook without AFE's part (but the manifest, which lists parts).
    const plain = pullProject(base, "book.xlsx").files;
    for (const [k, v] of Object.entries(plain)) if (k.startsWith("names/")) expect(p.files[k], k).toBe(v);
  });
});

describe("build", () => {
  it("keeps AFE's parts byte for byte, and warns naming the names AFE's modules also define", () => {
    const files = { ...pullProject(v11, "book.xlsx").files };
    edit(files, "names/ANA.xln", "LAMBDA(x, x ^ 3)", "LAMBDA(x, x * x * x)");
    const r = build(v11, files);
    expect(r.status).toBe("built");
    const w = r.plan.problems.filter((p) => p.code === "afe-modules");
    expect(w).toHaveLength(1);
    expect(w[0]!.severity).toBe("warning");
    expect(w[0]!.message).toContain("this build changes 1 name that AFE's modules also define (ANA.CUBE)");
    expect(w[0]!.message).toContain("xln leaves AFE's copy exactly as it is");
    expect(sameRecords(afeRecords(v11), rawZipRecords(r.bytes!))).toEqual([]);
    // And the check after the build says where they now differ.
    expect(audit(readWorkbook(r.bytes!), { only: ["C14"] }).findings.filter((f) => f.rule === "C14.afe-drift").map((f) => f.where.key)).toEqual(["ANA.CUBE"]);
  });

  it("keeps them with the embedded source too (xln's part beside AFE's)", () => {
    const files = { ...pullProject(v11, "book.xlsx").files };
    edit(files, "names/ANA.xln", "LAMBDA(x, x ^ 3)", "LAMBDA(x, x * x * x)");
    const r = build(v11, files, true);
    expect(r.status).toBe("built");
    const after = rawZipRecords(r.bytes!);
    expect(sameRecords(afeRecords(v11), after)).toEqual([]);
    expect([...after.keys()].filter((n) => /^customXml\/item\d+\.xml$/.test(n)).sort()).toEqual(["customXml/item1.xml", "customXml/item2.xml"]);
    expect(readWorkbook(r.bytes!).foreignModuleStores.map((s) => s.part)).toEqual(["customXml/item1.xml"]);
  });

  it("says nothing when the build changes no name AFE defines", () => {
    const files = { ...pullProject(v11, "book.xlsx").files };
    edit(files, "names/_unmanaged.xln", "MAP(Sheet1!$A$1:$A$3, ABS)", "MAP(Sheet1!$A$1:$A$3, SIGN)");
    const r = build(v11, files);
    expect(r.status).toBe("built");
    expect(r.plan.problems.filter((p) => p.code === "afe-modules")).toEqual([]);
  });

  it("warns without names when AFE's store cannot be read (AFE 1.0's code sheet), and keeps the sheets", () => {
    const files = { ...pullProject(v10, "book.xlsx").files };
    edit(files, "names/ANA.xln", "LAMBDA(x, x ^ 3)", "LAMBDA(x, x * x * x)");
    const r = build(v10, files);
    expect(r.status).toBe("built");
    const w = r.plan.problems.find((p) => p.code === "afe-modules");
    expect(w?.message).toContain("xln cannot read it");
    expect(sameRecords(afeRecords(v10), rawZipRecords(r.bytes!), (n) => n.startsWith("xl/worksheets/sheet2") || n.startsWith("xl/worksheets/sheet3"))).toEqual([]);
  });
});

// ---- real AFE-saved workbooks, when the author has dropped some in ------------------------
const real = existsSync(REAL) ? readdirSync(REAL).filter((f) => f.endsWith(".xlsx") && !f.startsWith("~$")) : [];

describe.skipIf(real.length === 0)("AFE-saved workbooks in probes/results/afe", () => {
  for (const f of real) {
    describe(f, () => {
      const bytes = new Uint8Array(readFileSync(join(REAL, f)));
      const wb = readWorkbook(bytes);

      it("has a readable AFE store whose modules parse and whose names agree with the workbook", () => {
        const stores = wb.foreignModuleStores.filter((s) => s.kind === "custom-xml");
        expect(stores.length, "no AFE custom XML part found").toBeGreaterThan(0);
        for (const s of stores) {
          expect(s.unreadable).toBeUndefined();
          expect(s.namespace).toBe(AFE_BLOB_NS);
          expect(s.linked).toBe(true);
        }
        // Right after AFE saved them, its text and the names must agree.
        const off = afeStatus(wb).flatMap((s) => s.entries.filter((e) => e.state === "differs" || e.state === "absent").map((e) => `${e.name} ${e.state}: ${e.afe} | ${e.workbook ?? ""}`));
        expect(off).toEqual([]);
      });

      it("keeps AFE's parts byte for byte through a build", () => {
        const files = { ...pullProject(bytes, f).files };
        const r = buildWorkbook({ workbook: bytes, fileName: f, files }, { force: true, provenance: false });
        expect(r.status).toBe("built");
        const before = rawZipRecords(bytes);
        const after = rawZipRecords(r.bytes!);
        expect(sameRecords(before, after, (n) => n.startsWith("customXml/") || n.startsWith("xl/webextensions/"))).toEqual([]);
      });
    });
  }
});
