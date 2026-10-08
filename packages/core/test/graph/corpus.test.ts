// The dependency graph on real workbooks: the probe fixtures always, the corpus with
// XLN_CORPUS=<folder> (workbooks at */dist/*.xlsx). Prints per workbook the nodes,
// edges, circular references, unresolvable references, C9 findings, unused names and the
// time to build, and checks the order: every formula once, every edge outside a cycle
// pointing backwards; the calculation-order views list exactly the lines of the views in
// order of appearance.
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { basename, join } from "node:path";
import { describe, expect, it } from "vitest";
import { buildGraph, readWorkbook, sheetCalcView, sheetFormulaView, workbookFormulaView, workbookNameIndex, type DependencyGraph, type WorkbookSnapshot } from "../../src/index.js";

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

const cpuMs = (u: NodeJS.CpuUsage) => (u.user + u.system) / 1000;

function check(wb: WorkbookSnapshot, file: string): { g: DependencyGraph; cpuMs: number } {
  // Best of three warm builds, as for the formula view (the first includes compiling the code).
  let g = buildGraph(wb);
  let best = Infinity;
  for (let k = 0; k < 3; k++) {
    const c0 = process.cpuUsage();
    g = buildGraph(wb);
    best = Math.min(best, cpuMs(process.cpuUsage(c0)));
  }

  const order = g.order();
  expect(order).toHaveLength(g.nodes.length);
  expect(new Set(order.map((n) => n.id)).size).toBe(g.nodes.length);
  const at = new Map(order.map((n, k) => [n.id, k]));
  for (const n of g.nodes) {
    for (const p of g.precedents(n)) {
      if (n.cycle !== undefined && n.cycle === p.cycle) continue;
      if (p === n) continue;
      // LAMBDA recursion: members of one component, not a circular reference.
      if (n.kind === "name" && p.kind === "name" && n.name!.lambda && p.name!.lambda && g.recursions.some((r) => r.members.includes(n) && r.members.includes(p))) continue;
      expect(at.get(p.id)!, `${file}: ${n.label} before ${p.label}`).toBeLessThan(at.get(n.id)!);
      expect(p.level, `${file}: level of ${n.label} vs ${p.label}`).toBeLessThanOrEqual(n.level);
    }
  }

  // The views in calculation order hold the same lines as in order of appearance.
  const index = workbookNameIndex(wb);
  const cache = new Map<string, unknown>();
  const all = workbookFormulaView(wb, index, { graph: g });
  let total = 0;
  for (const s of wb.sheets) {
    const appearance = sheetFormulaView(wb, s.name, index, { cache }).map((l) => l.cell);
    const calc = sheetCalcView(wb, s.name, index, { graph: g }).map((l) => l.cell);
    expect(calc.length, `${file}!${s.name}`).toBe(appearance.length);
    expect([...calc].sort(), `${file}!${s.name}`).toEqual([...appearance].sort());
    total += appearance.length;
  }
  expect(all).toHaveLength(total);
  expect(g.nodes.filter((n) => n.kind === "formula")).toHaveLength(total);
  return { g, cpuMs: best };
}

function report(file: string, g: DependencyGraph, ms: number): void {
  const s = g.stats();
  const lines = [
    `${file}: ${s.formulas} formula nodes · ${s.inputs} inputs · ${s.names} names · ${s.edges} edges · ` +
      `${s.cycles} cycles · ${s.recursions} LAMBDA recursions · flagged: ${s.dynamic} dynamic, ${s.external} external, ${s.broken} broken · ` +
      `C9 ${g.spillRefs.length} · build ${ms.toFixed(1)} ms CPU`,
  ];
  for (const c of g.cycles) lines.push(`  cycle ${c.id} (${c.members.length}): ${c.members.slice(0, 8).map((m) => m.label).join(", ")}${c.members.length > 8 ? ", …" : ""}`);
  for (const n of g.flagged()) for (const f of n.flags) lines.push(`  ${f.kind} ${n.label}: ${f.reason} (${f.text.slice(0, 50)})`);
  for (const f of g.spillRefs) lines.push(`  C9 ${f.node.label}: ${f.ref} → ${f.use} (${f.fit})`);
  const u = g.unusedNames();
  if (u.unused.length) lines.push(`  unused names: ${u.unused.map((n) => n.label).join(", ")}`);
  if (u.onlyByUnused.length) lines.push(`  used only by unused names: ${u.onlyByUnused.map((n) => n.label).join(", ")}`);
  const nc = g.nameCycles();
  if (nc.length) lines.push(`  name cycles: ${nc.map((c) => (c.recursive ? "recursive " : "") + c.members.map((m) => m.label).join(" → ")).join("; ")}`);
  console.log(lines.join("\n"));
}

describe("dependency graph on the probe fixtures", () => {
  for (const f of readdirSync(RESULTS).filter((n) => n.endsWith(".xlsx"))) {
    it(f, () => {
      const wb = readWorkbook(new Uint8Array(readFileSync(join(RESULTS, f))));
      const { g, cpuMs } = check(wb, f);
      report(f, g, cpuMs);
    });
  }

  it("f7_base: INDIRECT is flagged, not guessed", () => {
    const wb = readWorkbook(new Uint8Array(readFileSync(join(RESULTS, "f7_base.xlsx"))));
    const g = buildGraph(wb);
    expect(g.node("S1!C7")!.flags).toMatchObject([{ kind: "dynamic", text: 'INDIRECT("Rate")' }]);
    expect(g.cycles).toEqual([]);
  });
});

describe.skipIf(!CORPUS)("dependency graph on the corpus (XLN_CORPUS)", () => {
  const files = CORPUS ? corpusWorkbooks(CORPUS) : [];
  it("has workbooks", () => expect(files.length).toBeGreaterThan(0));
  for (const path of files) {
    it(path.slice(CORPUS!.length + 1), () => {
      const wb = readWorkbook(new Uint8Array(readFileSync(path)));
      const { g, cpuMs } = check(wb, basename(path));
      report(basename(path), g, cpuMs);
      expect(cpuMs).toBeLessThan(200);
      const name = basename(path);
      if (name === "lbo-ep03r-circ.xlsx") {
        // One interest ↔ cash loop per scenario: IS interest lines and 'SCF recursive', with their names.
        expect(g.cycles.length).toBeGreaterThanOrEqual(1);
        for (const c of g.cycles) expect(c.members.some((m) => m.sheet === "SCF recursive" && m.kind === "formula")).toBe(true);
      }
      if (name === "lbo-ep03r.xlsx") expect(g.cycles).toEqual([]);
    });
  }
});
