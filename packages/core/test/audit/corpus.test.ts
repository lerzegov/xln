// The audit on real workbooks: the probe fixtures always, the corpus with
// XLN_CORPUS=<folder> (workbooks at */dist/*.xlsx). On each: no crash, the same report
// twice, under 300 ms; the findings are printed so a run shows what the audit says.
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { audit, cellValueMap, readWorkbook, type AuditReport, type WorkbookSnapshot } from "../../src/index.js";

const CORPUS = process.env["XLN_CORPUS"];
const RESULTS = join(import.meta.dirname, "..", "..", "..", "..", "probes", "results");

function corpusWorkbooks(root: string): string[] {
  const out: string[] = [];
  for (const d of readdirSync(root, { withFileTypes: true })) {
    if (!d.isDirectory()) continue;
    const dist = join(root, d.name, "dist");
    if (!existsSync(dist)) continue;
    for (const f of readdirSync(dist)) if (f.endsWith(".xlsx") && !f.startsWith("~$")) out.push(join(dist, f));
  }
  return out.sort();
}

function check(wb: WorkbookSnapshot, file: string): { r: AuditReport; ms: number } {
  const r = audit(wb, { workbook: file });
  let ms = Infinity;
  for (let k = 0; k < 3; k++) {
    const t0 = performance.now();
    const again = audit(wb, { workbook: file });
    ms = Math.min(ms, performance.now() - t0);
    expect(JSON.stringify(again), `${file}: stable`).toBe(JSON.stringify(r));
  }
  return { r, ms };
}

function report(file: string, r: AuditReport, ms: number): void {
  const lines = [`${file}: ${r.counts.error} errors, ${r.counts.warning} warnings, ${r.counts.info} info; ${r.census.total} names, tiers ${r.census.tiers.total} (${r.census.tiers.percent}%); ${r.spills.spills.length} spills; ${ms.toFixed(0)} ms`];
  for (const f of r.findings) lines.push(`  ${f.rule} ${f.where.kind} ${f.where.sheet ? f.where.sheet + "!" : ""}${f.where.name ?? f.where.ref}: ${f.message.slice(0, 140)}`);
  console.log(lines.join("\n"));
}

describe("audit on the probe workbooks", () => {
  for (const f of readdirSync(RESULTS).filter((x) => x.endsWith(".xlsx"))) {
    it(f, () => {
      const { r } = check(readWorkbook(new Uint8Array(readFileSync(join(RESULTS, f)))), f);
      // probe_win/probe_mac carry Fact (T12) and Z_Bare / _xludf. (F6) in some files.
      expect(r.findings.every((x) => x.message.length > 0)).toBe(true);
      if (f === "probe_win.xlsx") expect(r.findings.some((x) => x.rule === "C3.lambda" && x.where.name === "Fact")).toBe(true);
    });
  }
});

describe.skipIf(!CORPUS)("audit on the corpus", () => {
  const files = CORPUS ? corpusWorkbooks(CORPUS) : [];
  it("finds the corpus", () => expect(files.length).toBeGreaterThan(0));
  for (const file of files) {
    it(file.slice(CORPUS!.length + 1), () => {
      const wb = readWorkbook(new Uint8Array(readFileSync(file)));
      const { r, ms } = check(wb, file.split("/").pop()!);
      report(file.slice(CORPUS!.length + 1), r, ms);
      expect(ms).toBeLessThan(300);
      expect(r.census.total).toBeGreaterThan(0);
      // The author's LBO models: with their check harness declared (as xln.config.json does),
      // nothing is unused and the sentinels (FN.DEV's 1E+99, FN.FIXPOINT's 1E+300) pass C13.
      if (/(^|\/)lbo-[^/]*\.xlsx$/.test(file)) {
        const h = audit(wb, { harness: ["Check!*", "CHK.*", "Model!Fix*"] });
        expect(h.findings.filter((f) => f.check === "C10" && f.severity === "warning").map((f) => f.where.key)).toEqual([]);
        expect(h.byCheck.C13).toEqual({ error: 0, warning: 0, info: 0 });
        expect(h.census.tiers.excluded).toBeGreaterThan(0);
      }
      // C15 with the cells' values (decided 2026-10-07): the corpus's names are written by
      // its generator, never renamed after their labels, so nothing drifts. lbo-ep02's Peers
      // sheet writes each name in the column of labels and names its header row PR.Names
      // under "Excel name": a header row of text, which C15 leaves out.
      const c15 = audit(wb, { only: ["C15"], values: cellValueMap(new Uint8Array(readFileSync(file))) });
      expect(c15.findings.map((f) => `${f.where.key}: ${f.message}`)).toEqual([]);
    });
  }
});
