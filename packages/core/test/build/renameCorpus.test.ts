// M5 on the real workbooks (XLN_CORPUS), in memory only: in each, the defined name that the
// most cell formulas read is renamed with `renameInProject`, then built. Every formula of
// the built workbook must read as before except the renamed token (checked token by token
// with the tokenizer, not with the rename's own resolver), keep its kind, extent and
// metadata; a second build has nothing to do, the pull guard sees nothing unbuilt, and a
// fresh pull gives back the renamed source without its @renamed lines.
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { invalidName, pullProject, readWorkbook, renameInProject, tokenize, unbuiltEdits, type SourceRename, type WorkbookSnapshot } from "../../src/index.js";
import { SourceModel } from "../../src/check/model.js";
import { build, why } from "./helpers.js";

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

/** Every formula by place: definitions, cells (kind, extent, metadata, text), CF, DV. */
function formulas(wb: WorkbookSnapshot): Map<string, { text: string; meta: string }> {
  const out = new Map<string, { text: string; meta: string }>();
  for (const d of wb.definedNames) out.set(`name ${d.scope.kind === "sheet" ? d.scope.name + "!" : ""}${d.name.toLowerCase()}`, { text: d.definition, meta: `${d.hidden} ${d.comment ?? ""}` });
  for (const s of wb.sheets) {
    for (const f of s.formulas) out.set(`cell ${s.name}!${f.cell}`, { text: f.text ?? "", meta: `${f.kind} ${f.range ?? ""} ${f.si ?? ""} ${f.cm ?? ""} ${JSON.stringify(f.attributes)} ${JSON.stringify(f.value)}` });
    s.conditionalFormats.forEach((c, i) => out.set(`cf ${s.name}#${i}`, { text: c.formulas.join("\u0000"), meta: c.sqref }));
    s.dataValidations.forEach((v, i) => out.set(`dv ${s.name}#${i}`, { text: `${v.formula1 ?? ""}\u0000${v.formula2 ?? ""}`, meta: v.sqref }));
  }
  return out;
}

/** Whether `after` is `before` with some name tokens spelled `from` replaced by `to`, and nothing else. */
function onlyRenamed(before: string, after: string, from: string, to: string): boolean {
  if (before === after) return true;
  const a = tokenize(before);
  const b = tokenize(after);
  if (a.length !== b.length) return false;
  return a.every((t, i) => {
    const u = b[i]!;
    if (t.text === u.text) return true;
    if (t.kind !== "name" || u.kind !== "name" || t.value?.toLowerCase() !== from.toLowerCase() || u.value !== to) return false;
    // The qualifier, if any, is kept as written.
    return t.text.slice(0, t.text.length - t.value.length) === u.text.slice(0, u.text.length - u.value.length);
  });
}

/** Names read by the most cell statements of the project, most read first. */
function mostRead(files: Record<string, string>): string[] {
  const model = new SourceModel();
  for (const [p, t] of Object.entries(files)) if (p.startsWith("names/") && p.endsWith(".xln")) model.setFile(p, t);
  const count = new Map<string, number>();
  for (const d of model.cellDefs.concat(model.defs.filter((x) => x.entry.cell))) {
    for (const o of model.analysis(d).occurrences) if (o.kind === "use" && o.key !== undefined) count.set(o.key, (count.get(o.key) ?? 0) + 1);
  }
  return [...count].sort((a, b) => b[1] - a[1]).map(([k]) => k);
}

function withoutRenamed(files: Record<string, string>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [p, t] of Object.entries(files)) if (p.endsWith(".xln")) out[p] = t.split("\n").filter((l) => !l.trim().startsWith("@renamed(")).join("\n");
  return out;
}

describe.skipIf(!CORPUS)("rename across cells on the corpus (XLN_CORPUS)", () => {
  for (const path of CORPUS ? workbooks(CORPUS) : []) {
    it(path.slice(CORPUS!.length + 1), () => {
      const bytes = new Uint8Array(readFileSync(path));
      const files = { ...pullProject(bytes, "book.xlsx").files };
      // The most read name a rename can take (one the workbook reads only where the build reaches).
      let r: SourceRename | undefined;
      let result: ReturnType<typeof build> | undefined;
      const tried: string[] = [];
      for (const key of mostRead(files).slice(0, 8)) {
        const id = key.slice(key.lastIndexOf("!") + 1);
        const to = invalidName(`${id}_rn`) ? `rn_${id.replace(/[^\p{L}\p{N}_]/gu, "_")}` : `${id}_rn`;
        const x = renameInProject(files, key, to);
        if (typeof x === "string") {
          tried.push(`${key}: ${x}`);
          continue;
        }
        const b = build(bytes, { ...files, ...x.files });
        if (b.status !== "built") {
          tried.push(`${key}: ${why(b).split("\n")[0]}`);
          continue;
        }
        r = x;
        result = b;
        break;
      }
      expect(r, tried.join("\n")).toBeDefined();
      const rn = r!;
      const built = result!;
      expect(built.readBack!.problems).toEqual([]);
      const rename = built.plan.changeSet.changes.find((c) => c.op === "rename-name")!;
      expect(rename.op === "rename-name" && rename.references !== undefined && rename.references.cells > 0, JSON.stringify(rename)).toBe(true);
      // Only the rename: no cell set, nothing turned into a dynamic array.
      expect(built.plan.changeSet.changes.filter((c) => c.op !== "rename-name" && c.op !== "set-name").map((c) => c.op)).toEqual([]);

      const from = rn.from.slice(rn.from.lastIndexOf("!") + 1);
      const to = rn.to.slice(rn.to.lastIndexOf("!") + 1);
      const a = formulas(readWorkbook(bytes));
      const b = formulas(readWorkbook(built.bytes!));
      let changed = 0;
      const bad: string[] = [];
      const scope = rn.from.slice(0, rn.from.lastIndexOf("!") + 1);
      const oldKey = `name ${scope}${from.toLowerCase()}`;
      for (const [k, x] of a) {
        // The renamed name itself is found under its new key.
        const y = b.get(k === oldKey ? `name ${scope}${to.toLowerCase()}` : k);
        if (!y) {
          bad.push(`${k}: gone`);
          continue;
        }
        if (x.meta !== y.meta) bad.push(`${k}: ${x.meta} → ${y.meta}`);
        if (!onlyRenamed(x.text, y.text, from, to)) bad.push(`${k}: ${x.text} → ${y.text}`);
        if (x.text !== y.text) changed++;
      }
      expect(bad.slice(0, 10)).toEqual([]);
      expect(changed).toBeGreaterThan(0);

      const after = { ...files, ...rn.files, ...built.files };
      expect(build(built.bytes!, after).status).toBe("up-to-date");
      expect(unbuiltEdits({ workbook: built.bytes!, fileName: "book.xlsx", files: after })).toEqual([]);
      // A pull lays out a definition over several lines once it is longer than its width
      // (100): a longer name can tip one over. Layout aside, the files are the same.
      const repulled = withoutRenamed(pullProject(built.bytes!, "book.xlsx").files);
      const want = withoutRenamed(after);
      const squeeze = (t: string | undefined) => (t ?? "").split(/\s+/).join("");
      expect(Object.keys(repulled).sort()).toEqual(Object.keys(want).sort());
      expect(Object.keys(want).filter((p) => squeeze(want[p]) !== squeeze(repulled[p]))).toEqual([]);
    }, 120_000);
  }
});
