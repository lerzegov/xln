// M4 on the author's real library and workbooks: XLN_CORPUS=<excel-models>, the library at
// _shared/lib (read only: nothing here writes), workbooks at */dist/*.xlsx. Skipped
// without XLN_CORPUS.
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { compileWithDiagnostics, libraryStatus, pullProject, readLibrary, readWorkbook, renderLibStatus, projectCopies, workbookCopies, commentLength, COMMENT_MAX } from "../../src/index.js";

const CORPUS = process.env["XLN_CORPUS"];
const LIB = CORPUS ? join(CORPUS, "_shared", "lib") : "";

function libFiles(): Record<string, string> {
  const files: Record<string, string> = {};
  for (const f of readdirSync(LIB)) if (f.endsWith(".lambda")) files[f] = readFileSync(join(LIB, f), "utf8");
  return files;
}

describe.skipIf(!CORPUS || !existsSync(LIB))("the real library (XLN_CORPUS/_shared/lib)", () => {
  // The 17 FN.* functions of M4; the author adds to the library (ANA.*, EASY.* published
  // from a test project on 2026-10-07), so the others are only counted.
  it("reads all files (the 17 FN.* at least) with every field, and each definition compiles", () => {
    const files = libFiles();
    expect(Object.keys(files).filter((f) => f.startsWith("FN.")).length).toBe(17);
    const lib = readLibrary(files);
    for (const p of lib.problems) console.log(`${p.severity} ${p.path}:${p.line ?? ""} ${p.message}`);
    expect(lib.problems.filter((p) => p.severity === "error")).toEqual([]);
    expect(lib.functions.length).toBe(Object.keys(files).length);
    const names = lib.functions.map((f) => f.name);
    for (const fn of lib.functions) {
      expect(fn.path).toBe(`${fn.name}.lambda`);
      const keys = fn.fields.map((f) => f.key);
      expect(keys.slice(0, 3)).toEqual(["name", "summary", "params"]);
      for (const k of keys) expect(["name", "summary", "params", "impromptu", "example"]).toContain(k);
      expect(fn.summary).not.toBe("");
      expect(fn.definition.startsWith("LAMBDA(")).toBe(true);
      expect(fn.rationale).not.toBe("");
      expect(commentLength(fn.doc)).toBeLessThanOrEqual(COMMENT_MAX);
      const c = compileWithDiagnostics(fn.definition, { names });
      expect(c.diagnostics.filter((d) => d.severity === "error")).toEqual([]);
    }
    const grow = lib.get("fn.grow")!;
    expect(grow.params).toEqual(["seed", "growth", "periods"]);
    expect(grow.doc).toBe("Growth chain: seed the first period, compound it thereafter.\n@param seed\n@param growth\n@param periods");
    expect(grow.fields.find((f) => f.key === "impromptu")!.value).toBe("x[year:first] = <seed>\nx[year:rest]  = x[year:prev] * (1 + <g>)");
    expect(grow.fields.find((f) => f.key === "example")!.value).toBe("FN.GROW(10000, 0.05, 5)\n->  {10000, 10500, 11025, 11576.25, 12155.0625}");
    expect(grow.rationale.split("\n")[0]).toBe("Five of the six line items in SalesCOGS are exactly this shape, which is why it");
    // A two-line summary is one sentence.
    expect(lib.get("FN.AVGPART")!.summary).toBe("Time-weighted average of ONE part (above or below zero) of a balance moving linearly over the period, given both parts' end values.");
    // A summary too long for the Name Manager is cut at a word; the @param lines stay.
    const fix = lib.get("FN.FIXPOINT")!;
    console.log(fix.doc);
    expect(fix.docShortened).toBe(true);
    expect(fix.doc.split("\n").slice(1)).toEqual(["@param step", "@param seed", "@param tol", "@param cap"]);
    expect(fix.doc.split("\n")[0]!.endsWith("…")).toBe(true);
    console.log(
      lib.functions
        .map((f) => `${f.name}: ${f.params.join(", ")} · doc ${commentLength(f.doc)} chars${f.docShortened ? " (shortened)" : ""} · fields ${f.fields.map((x) => x.key).join(",")}`)
        .join("\n"),
    );
  });
});

describe.skipIf(!CORPUS || !existsSync(LIB))("lib status of the LBO workbooks against _shared/lib (M4 exit)", () => {
  const lib = CORPUS && existsSync(LIB) ? readLibrary(libFiles()) : undefined;
  for (const [dir, file] of [
    ["lbo-ep02", "lbo-ep02.xlsx"],
    ["lbo-ep03", "lbo-ep03.xlsx"],
    ["lbo-ep03r", "lbo-ep03r.xlsx"],
  ] as const) {
    const path = CORPUS ? join(CORPUS, dir, "dist", file) : "";
    it.skipIf(!path || !existsSync(path))(file, () => {
      const bytes = new Uint8Array(readFileSync(path));
      const wb = readWorkbook(bytes);
      const r = libraryStatus(lib!, workbookCopies(wb), { target: file, kind: "workbook", library: "_shared/lib" });
      console.log(renderLibStatus(r).text);
      // Every library function has a state; built by the Python build, nothing carries a tag.
      expect(r.items.filter((i) => i.state !== "local-only").length).toBe(lib!.functions.length);
      expect(r.counts.outdated + r.counts.modified).toBe(0);
      // The IN.* input readers are not library candidates: counted, not listed.
      expect(r.counts["local-only"]).toBe(0);
      expect(Object.keys(r.otherModules)).toEqual(["IN"]);
      // The same answer from the pulled project (its source compiled again).
      const pulled = pullProject(bytes, file);
      const p = libraryStatus(lib!, projectCopies(pulled.files, wb), { target: file, kind: "project", library: "_shared/lib" });
      expect(p.items.map((i) => [i.name, i.state])).toEqual(r.items.map((i) => [i.name, i.state]));
    });
  }
});
