// The cell backend on the real workbooks (XLN_CORPUS), in memory only: every formula cell
// rewritten with its own formula (dynamic-array form, shared groups un-shared), a spill
// anchor cleared, a slot filled; each build read back (E3), and the formulas mean what
// they meant.
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { applyChangeSetWithReport, decompile, formulaTextAt, parseCell, readBack, readWorkbook, shiftFormula, type Change } from "../../src/index.js";

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

describe.skipIf(!CORPUS)("cell backend on the corpus (XLN_CORPUS)", () => {
  const files = CORPUS ? workbooks(CORPUS) : [];
  for (const path of files) {
    it(path.slice(CORPUS!.length + 1), () => {
      const bytes = new Uint8Array(readFileSync(path));
      const wb = readWorkbook(bytes);
      const names = wb.definedNames.map((d) => d.name);
      const changes: Change[] = [];
      const want = new Map<string, string>();
      for (const s of wb.sheets) {
        for (const f of s.formulas) {
          if (f.kind === "data-table" || f.kind === "array") continue;
          const text = formulaTextAt(s, f, shiftFormula);
          if (text === undefined || text.trim() === "") continue;
          changes.push({ op: "set-cell-formula", sheet: s.name, range: f.cell, stored: text, display: decompile(text, { names }) });
          want.set(`${s.name}!${f.cell}`, text);
        }
      }
      if (changes.length === 0) return;
      const r = applyChangeSetWithReport(bytes, changes);
      const rb = readBack(bytes, r.bytes, changes, []);
      expect(rb.problems.slice(0, 20)).toEqual([]);
      const after = readWorkbook(r.bytes);
      for (const s of after.sheets) {
        for (const f of s.formulas) {
          const w = want.get(`${s.name}!${f.cell}`);
          if (w === undefined) continue;
          expect(f.kind, `${s.name}!${f.cell}`).toBe("dynamic-array");
          expect(f.text, `${s.name}!${f.cell}`).toBe(w);
          want.delete(`${s.name}!${f.cell}`);
        }
        expect(s.sharedFormulas).toEqual([]);
      }
      expect([...want.keys()].slice(0, 10)).toEqual([]);

      // A spill anchor cleared and a slot filled below the first sheet's last row.
      const spillSheet = wb.sheets.find((s) => s.spills.some((x) => x.extent.includes(":")));
      const sheet = wb.sheets.find((s) => s.kind === "worksheet")!;
      const lastRow = Math.max(0, ...sheet.formulas.map((f) => parseCell(f.cell)!.row));
      const more: Change[] = [{ op: "set-cell-formula", sheet: sheet.name, range: `A${lastRow + 3}:B${lastRow + 3}`, stored: "ROW()*2", display: "ROW()*2" }];
      if (spillSheet) more.push({ op: "clear-cell-formula", sheet: spillSheet.name, range: spillSheet.spills.find((x) => x.extent.includes(":"))!.anchor });
      const r2 = applyChangeSetWithReport(bytes, more);
      expect(readBack(bytes, r2.bytes, more, []).problems).toEqual([]);
    });
  }
});
