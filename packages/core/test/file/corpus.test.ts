// The real workbooks outside the repo: XLN_CORPUS=<folder>, workbooks at */dist/*.xlsx.
// Counts from readWorkbook are checked against a plain string count over the raw XML,
// which shares no code with the parser.
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { readWorkbook } from "../../src/index.js";

const CORPUS = process.env["XLN_CORPUS"];

function workbooks(root: string): string[] {
  const out: string[] = [];
  for (const d of readdirSync(root, { withFileTypes: true })) {
    if (!d.isDirectory()) continue;
    const dist = join(root, d.name, "dist");
    if (!existsSync(dist)) continue;
    for (const f of readdirSync(dist)) if (f.endsWith(".xlsx") && !f.startsWith("~$")) out.push(join(dist, f));
  }
  return out.sort();
}

function count(hay: string, needle: string): number {
  let n = 0;
  for (let i = hay.indexOf(needle); i >= 0; i = hay.indexOf(needle, i + needle.length)) n++;
  return n;
}

function rawCounts(bytes: Uint8Array): { names: number; formulas: number } {
  const files = unzipSync(bytes);
  const wb = strFromU8(files["xl/workbook.xml"]!);
  let formulas = 0;
  for (const [name, data] of Object.entries(files)) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/.test(name)) continue;
    const s = strFromU8(data);
    formulas += count(s, "<f>") + count(s, "<f ") + count(s, "<f/>");
  }
  return { names: count(wb, "<definedName "), formulas };
}

describe.skipIf(!CORPUS)("corpus (XLN_CORPUS)", () => {
  const files = CORPUS ? workbooks(CORPUS) : [];
  it("has workbooks", () => expect(files.length).toBeGreaterThan(0));
  for (const path of files) {
    it(path.slice(CORPUS!.length + 1), () => {
      const bytes = new Uint8Array(readFileSync(path));
      const t0 = performance.now();
      const wb = readWorkbook(bytes);
      const ms = performance.now() - t0;
      const raw = rawCounts(bytes);
      const formulas = wb.sheets.reduce((n, s) => n + s.formulas.length, 0);
      const spills = wb.sheets.reduce((n, s) => n + s.spills.length, 0);
      const shared = wb.sheets.reduce((n, s) => n + s.sharedFormulas.length, 0);
      console.log(
        `${path.split("/").pop()}: ${bytes.length} bytes, ${wb.sheets.length} sheets, ` +
          `${wb.definedNames.length} names (raw ${raw.names}), ${formulas} formulas (raw ${raw.formulas}), ` +
          `${shared} shared groups, ${spills} spills, ${wb.tables.length} tables, ` +
          `${wb.warnings.length} warnings, ${ms.toFixed(1)} ms`,
      );
      expect(wb.definedNames.length).toBe(raw.names);
      expect(formulas).toBe(raw.formulas);
      expect(wb.warnings).toEqual([]);
    });
  }
});
