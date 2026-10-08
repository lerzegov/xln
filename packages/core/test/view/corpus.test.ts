// The formula view on every sheet of real workbooks: the probe fixtures always, the corpus
// with XLN_CORPUS=<folder> (workbooks at */dist/*.xlsx).
//
// Checked independently of the view's own logic: every <f> cell appears exactly once
// unless it lies inside another cell's array, spill or data table; every displayed
// formula compiles back to the text stored at that cell (a shared child's moved text)
// modulo whitespace; every name and reference span lands on its text in the rendered
// document; rendering a whole workbook takes under 100 ms. Left-hand sides: every name
// classified as a single cell or a spill whose cell has a line of its own is placed on
// that line, exactly once in the whole workbook; the others are counted and reported.
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { basename, join } from "node:path";
import { describe, expect, it } from "vitest";
import {
  classify,
  compile,
  definitionTarget,
  equalModuloWhitespace,
  extentSize,
  formatCell,
  formulaTextAt,
  parseCell,
  readWorkbook,
  renderFormulaView,
  sheetFormulaView,
  shiftFormula,
  workbookNameIndex,
  type Sheet,
  type WorkbookSnapshot,
} from "../../src/index.js";

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

/** The <f> cells Excel shows on their own: not inside another anchor's saved extent. */
function expectedCells(sheet: Sheet): string[] {
  const areas: { anchor: string; r1: number; c1: number; r2: number; c2: number }[] = [];
  for (const f of sheet.formulas) {
    if (f.kind !== "array" && f.kind !== "dynamic-array" && f.kind !== "data-table") continue;
    const range = f.range ?? sheet.spills.find((s) => s.anchor === f.cell)?.extent;
    if (!range) continue;
    const [a, b = a] = range.split(":");
    const p = parseCell(a!)!;
    const q = parseCell(b!)!;
    areas.push({ anchor: f.cell, r1: p.row, c1: p.col, r2: q.row, c2: q.col });
  }
  return sheet.formulas
    .filter((f) => {
      const c = parseCell(f.cell)!;
      return !areas.some((x) => x.anchor !== f.cell && c.row >= x.r1 && c.row <= x.r2 && c.col >= x.c1 && c.col <= x.c2);
    })
    .map((f) => f.cell);
}

interface Stats {
  sheets: number;
  formulas: number;
  lines: number;
  /** Lines with at least one name on the left, and names placed on lines. */
  named: number;
  placed: number;
  /** Single-cell or spill names on a formula cell: expected on a line. */
  expected: number;
  /** Single-cell or spill names on a cell without a line: value cells, cells inside another spill. */
  onValue: number;
  inside: string[];
  skipped: number;
  spills: number;
  shared: number;
  names: number;
  unresolved: number;
  ms: number;
  coldMs: number;
  cpuMs: number;
  coldCpuMs: number;
}

function checkWorkbook(wb: WorkbookSnapshot, file: string): Stats {
  const st: Stats = { sheets: 0, formulas: 0, lines: 0, named: 0, placed: 0, expected: 0, onValue: 0, inside: [], skipped: 0, spills: 0, shared: 0, names: 0, unresolved: 0, ms: 0, coldMs: 0, cpuMs: 0, coldCpuMs: 0 };
  const allNames = wb.definedNames.filter((d) => !d.isXlPrefixed).map((d) => d.name);
  const build = () => {
    const index = workbookNameIndex(wb);
    const cache = new Map<string, unknown>();
    return wb.sheets.map((s) => {
      const lines = sheetFormulaView(wb, s.name, index, { cache });
      return { sheet: s, lines, rendered: renderFormulaView(lines, { sheet: s.name, workbook: file }) };
    });
  };
  // The first run includes compiling the code; the best of three later runs is what a
  // reader waits for once the extension is warm, with the noise of the test files that
  // Vitest runs alongside filtered out.
  // Vitest runs each test file in a process of its own, so the process's CPU time
  // is not inflated by the other files running alongside; wall time is.
  const cpuMs = (u: NodeJS.CpuUsage) => (u.user + u.system) / 1000;
  let t0 = performance.now();
  let c0 = process.cpuUsage();
  let views = build();
  st.coldMs = performance.now() - t0;
  st.coldCpuMs = cpuMs(process.cpuUsage(c0));
  st.ms = Infinity;
  st.cpuMs = Infinity;
  for (let k = 0; k < 3; k++) {
    t0 = performance.now();
    c0 = process.cpuUsage();
    views = build();
    st.ms = Math.min(st.ms, performance.now() - t0);
    st.cpuMs = Math.min(st.cpuMs, cpuMs(process.cpuUsage(c0)));
  }

  for (const { sheet, lines, rendered } of views) {
    st.sheets++;
    st.formulas += sheet.formulas.length;
    st.lines += lines.length;
    const want = expectedCells(sheet);
    st.skipped += sheet.formulas.length - want.length;
    const got = lines.map((l) => l.cell);
    expect(new Set(got).size, `${sheet.name}: a cell listed twice`).toBe(got.length);
    expect([...got].sort(), `${sheet.name}: cells`).toEqual([...want].sort());
    for (let k = 1; k < lines.length; k++) {
      const a = lines[k - 1]!;
      const b = lines[k]!;
      expect(a.row < b.row || (a.row === b.row && a.col < b.col), `${sheet.name}: order at ${b.cell}`).toBe(true);
    }
    for (const l of lines) {
      if (l.kind === "dynamic-array") {
        const size = extentSize(l.extent!);
        if (size && size.rows * size.cols > 1) st.spills++;
      }
      if (l.kind === "shared") st.shared++;
      st.names += l.names.length;
      if (l.lhs.length > 0) st.named++;
      st.placed += l.lhs.length;
      st.unresolved += l.names.filter((n) => n.key === undefined).length;
      if (l.kind === "data-table" || l.stored.trim() === "") continue;
      const f = sheet.formulas.find((x) => x.cell === l.cell)!;
      expect(l.stored, `${sheet.name}!${l.cell}: stored text`).toBe(formulaTextAt(sheet, f, shiftFormula));
      expect(l.error, `${sheet.name}!${l.cell}: ${l.stored}`).toBeUndefined();
      const back = compile(l.formula, { names: allNames, links: wb.externalLinks, allowUnknownFunctions: true });
      expect(equalModuloWhitespace(back, l.stored), `${sheet.name}!${l.cell}: ${l.formula} → ${back} ≠ ${l.stored}`).toBe(true);
      for (const n of l.names) expect(l.formula.slice(n.start, n.end)).toBe(n.id);
    }
    expect(rendered.entries).toHaveLength(lines.length);
    for (const e of rendered.entries) {
      const l = lines[e.index]!;
      expect(rendered.text.slice(e.address.start, e.address.end)).toBe(l.cell);
      expect(e.names.map((s) => rendered.text.slice(s.start, s.end))).toEqual(l.names.map((n) => n.id));
      expect(e.refs.map((s) => rendered.text.slice(s.start, s.end))).toEqual(l.refs.map((r) => l.formula.slice(r.start, r.end)));
      expect(e.lhs.map((s) => rendered.text.slice(s.start, s.end))).toEqual(l.lhs.map((n) => n.display));
    }
  }

  // Every single-cell or spill name on a cell with a line of its own sits on that line, once.
  const placedAt = new Map<string, string[]>();
  for (const { sheet, lines } of views) for (const l of lines) for (const n of l.lhs) placedAt.set(n.key, [...(placedAt.get(n.key) ?? []), `${sheet.name}!${l.cell}`]);
  for (const d of wb.definedNames) {
    if (d.isXlPrefixed || d.scopeInvalid) continue;
    const c = classify(d.definition);
    const t = definitionTarget(d.definition);
    const scope = d.scope.kind === "sheet" ? d.scope.name : undefined;
    const key = scope === undefined ? d.name : `${scope}!${d.name}`;
    let at: { sheet: string | undefined; cell: string } | undefined;
    if (c.kind === "spill" && c.anchor) at = { sheet: c.anchor.sheet ?? scope, cell: c.anchor.cell };
    else if (c.kind === "range" && t && !t.spill && t.r1 === t.r2 && t.c1 === t.c2) at = { sheet: t.sheet ?? scope, cell: formatCell({ row: t.r1, col: t.c1 }) };
    if (!at || at.sheet === undefined) continue;
    const view = views.find((v) => v.sheet.name.toLowerCase() === at!.sheet!.toLowerCase());
    if (!view) continue;
    const line = view.lines.find((l) => l.cell === at!.cell);
    if (!line) {
      if (view.sheet.formulas.some((f) => f.cell === at!.cell)) st.inside.push(`${key} (${at.sheet}!${at.cell})`);
      else st.onValue++;
      expect(placedAt.get(key), `${key} placed although its cell has no line`).toBeUndefined();
      continue;
    }
    st.expected++;
    expect(placedAt.get(key), `${key} on ${at.sheet}!${at.cell}`).toEqual([`${view.sheet.name}!${at.cell}`]);
  }
  return st;
}

function report(file: string, s: Stats): void {
  console.log(
    `${file}: ${s.sheets} sheets · ${s.formulas} <f> cells → ${s.lines} lines (${s.skipped} inside another spill/array) · ` +
      `${s.spills} spills · ${s.shared} shared · ${s.names} name uses (${s.unresolved} unresolved) · ` +
      `left-hand side: ${s.named} lines named, ${s.placed} names placed (${s.expected} single-cell/spill names on formula cells, all placed; ${s.onValue} on value cells${s.inside.length ? `; inside another spill: ${s.inside.join(", ")}` : ""}) · view+render ${s.cpuMs.toFixed(1)} ms CPU, ${s.ms.toFixed(1)} ms wall (first run ${s.coldCpuMs.toFixed(1)} / ${s.coldMs.toFixed(1)} ms)`,
  );
}

describe("formula view on the probe fixtures", () => {
  for (const f of readdirSync(RESULTS).filter((n) => n.endsWith(".xlsx"))) {
    it(f, () => {
      const wb = readWorkbook(new Uint8Array(readFileSync(join(RESULTS, f))));
      report(f, checkWorkbook(wb, f));
    });
  }
});

describe.skipIf(!CORPUS)("formula view on the corpus (XLN_CORPUS)", () => {
  const files = CORPUS ? corpusWorkbooks(CORPUS) : [];
  it("has workbooks", () => expect(files.length).toBeGreaterThan(0));
  for (const path of files) {
    it(path.slice(CORPUS!.length + 1), () => {
      const wb = readWorkbook(new Uint8Array(readFileSync(path)));
      const s = checkWorkbook(wb, basename(path));
      report(basename(path), s);
      expect(s.cpuMs).toBeLessThan(100);
    });
  }
});
