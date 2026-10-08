// The editor's project model on the real workbooks: XLN_CORPUS=<folder>, workbooks at
// */dist/*.xlsx. Pulls in memory (nothing is written), loads the result like the
// extension does, and checks it agrees with the manifest.
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { formulaToSource, MANIFEST_FILE, pullProject, readWorkbook } from "@xln/core";
import { describe, expect, it } from "vitest";
import { completions, liveProblems } from "../src/model/editor.js";
import { parseManifest } from "../src/model/manifest.js";
import { isNamesFile, Project } from "../src/model/project.js";
import { hoverMarkdown, renderUsages } from "../src/model/usages.js";

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

describe.skipIf(!CORPUS)("editor model on the corpus (XLN_CORPUS)", () => {
  for (const path of CORPUS ? workbooks(CORPUS) : []) {
    it(path.slice(CORPUS!.length + 1), () => {
      const bytes = new Uint8Array(readFileSync(path));
      const pulled = pullProject(bytes, path);
      const t0 = performance.now();
      const p = new Project("mem:/c", parseManifest(pulled.files[MANIFEST_FILE]!));
      for (const [f, text] of Object.entries(pulled.files)) if (isNamesFile(f)) p.setFile(f, text);
      for (const d of p.defs) p.analysis(d);
      for (const f of p.files.keys()) expect(p.problems(f).filter((x) => x.code !== "workbook-on-cell").map((x) => `${x.severity} ${x.code} ${f}: ${x.message}`)).toEqual([]);
      const ms = performance.now() - t0;

      expect(p.defs.length).toBe(pulled.report.names);
      for (const d of p.defs) {
        const text = d.file.text;
        expect(text.slice(d.entry.offset, d.entry.offset + d.name.length)).toBe(d.name);
        // A pulled file has no comments inside formulas: the formula is one run of text.
        expect(text.slice(formulaToSource(d.entry, 0), formulaToSource(d.entry, d.entry.formula.length))).toBe(d.entry.formula);
        // Live resolution on display text agrees with the manifest's on stored text. A cell
        // statement's text is its cell's formula, which the manifest indexes as a cell.
        if (!d.entry.cell) {
          const live = p.uses(d).map((u) => u.def.key.toLowerCase()).sort();
          const manifest = (p.manifestName(d.key)?.uses ?? []).map((k) => k.toLowerCase()).sort();
          expect(live, d.key).toEqual(manifest);
        }
        const usedBy = p
          .usedBy(d.key)
          .filter((u) => !u.def.entry.cell)
          .map((u) => u.def.key.toLowerCase())
          .sort();
        expect(usedBy, d.key).toEqual((p.manifestName(d.key)?.usedBy?.names ?? []).map((k) => k.toLowerCase()).sort());
      }
      const wb = readWorkbook(bytes);
      for (const d of p.defs) {
        expect(hoverMarkdown(p, d, wb)).toContain(d.name);
        const doc = renderUsages(d.key, p.manifestName(d.key), p.manifest, wb);
        const lines = doc.text.split("\n");
        // Every listed cell shows its formula: the workbook agrees with the manifest.
        for (const i of doc.items.filter((x) => x.kind === "cell")) expect(lines[i.line], d.key).toContain(": =");
      }
      console.log(`${pulled.report.workbook}: ${p.defs.length} names, ${p.files.size} files, model load + analysis + problems ${ms.toFixed(0)} ms`);

      // Checks as you type (M3c): a pulled project has no errors; warnings are listed.
      const t1 = performance.now();
      const live = [...p.files.keys()].flatMap((f) => liveProblems(p, f));
      const liveMs = performance.now() - t1;
      const shown = live.filter((x) => x.severity !== "hint");
      for (const x of shown) console.log(`  live ${x.severity} ${x.code} ${x.path}:${p.files.get(x.path)!.lines.position(x.start).line + 1} ${x.message}`);
      console.log(`  live checks: ${live.length} (${live.filter((x) => x.severity === "hint").length} hints) in ${liveMs.toFixed(0)} ms`);
      expect(live.filter((x) => x.severity === "error").map((x) => `${x.path} ${x.message}`)).toEqual([]);
    });
  }

  it.skipIf(!CORPUS || !existsSync(join(CORPUS ?? "", "lbo-ep03r/dist/lbo-ep03r.xlsx")))("completion in lbo-ep03r BS.xln lists IS!Sales_base", () => {
    const path = join(CORPUS!, "lbo-ep03r/dist/lbo-ep03r.xlsx");
    const pulled = pullProject(new Uint8Array(readFileSync(path)), path);
    const p = new Project("mem:/c", parseManifest(pulled.files[MANIFEST_FILE]!));
    for (const [f, text] of Object.entries(pulled.files)) if (isNamesFile(f)) p.setFile(f, text);
    const bs = "names/sheets/BS.xln";
    const text = p.files.get(bs)!.text;
    // In `NetIncome_base @C22# = IS!NetIncome_base;`, with the cursor after `= `.
    const at = text.indexOf("NetIncome_base @C22# = ") + "NetIncome_base @C22# = ".length;
    const items = completions(p, bs, at)!;
    const labels = items.map((i) => i.label);
    expect(labels).toContain("IS!Sales_base");
    expect(labels).toContain("NetIncome_base"); // BS's own, bare
    expect(labels).toContain("FN.");
    expect(labels).toContain("'SCF recursive'!ExcessCash_base");
    expect(items.find((i) => i.label === "IS!Sales_base")!.filterText).toBe("Sales_base");
    const fn = completions(p, bs, at + "IS!NetIncome_base".length)!;
    expect(fn).toBeDefined();
    // After `FN.`: the module's members.
    const q = new Project("mem:/c", p.manifest);
    for (const [f, sf] of p.files) q.setFile(f, f === bs ? text.slice(0, at) + "FN." + text.slice(at) : sf.text);
    const members = completions(q, bs, at + 3)!.map((i) => i.label);
    expect(members).toContain("FN.SEEDROW");
    expect(members.every((m) => m.startsWith("FN."))).toBe(true);
  });
});
