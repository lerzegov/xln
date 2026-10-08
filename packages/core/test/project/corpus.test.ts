// pullProject on the real workbooks: XLN_CORPUS=<folder>, workbooks at */dist/*.xlsx.
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { pullProject } from "../../src/index.js";
import { checkProject } from "./check.js";

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

describe.skipIf(!CORPUS)("pull on the corpus (XLN_CORPUS)", () => {
  const files = CORPUS ? workbooks(CORPUS) : [];
  it("has workbooks", () => expect(files.length).toBeGreaterThan(0));
  for (const path of files) {
    it(path.slice(CORPUS!.length + 1), () => {
      const bytes = new Uint8Array(readFileSync(path));
      const t0 = performance.now();
      const r = pullProject(bytes, path);
      const ms = performance.now() - t0;
      const c = checkProject(r);
      const kinds = Object.entries(r.report.byKind)
        .filter(([, n]) => n > 0)
        .map(([k, n]) => `${k} ${n}`)
        .join(", ");
      console.log(
        `${r.report.workbook}: ${r.report.names} names (${r.report.byScope.workbook} workbook, ${r.report.byScope.sheet} sheet) · ${kinds} · ` +
          `modules ${r.report.modules.map((m) => `${m.module} ${m.names}`).join(", ")} · ` +
          `sheet files ${r.report.sheetFiles.map((f) => `${f.sheet} ${f.names}`).join(", ") || "none"} · unmanaged ${r.report.unmanaged} · ` +
          `built-ins ${r.report.builtIns.length} · repaired ${c.repaired.length} · warnings ${r.report.warnings.length} · ` +
          `unindexed formulas ${r.report.unparsedFormulas} · ${ms.toFixed(0)} ms`,
      );
      for (const w of r.report.warnings.slice(0, 5)) console.log(`  warning: ${w}`);
      expect(c.failures).toEqual([]);
      expect(c.entries).toBe(r.report.names);
      expect(r.report.byKind.unparsed).toBe(0);
      // Deterministic: a second pull writes the same bytes.
      expect(pullProject(bytes, path).files).toEqual(r.files);
    });
  }
});
