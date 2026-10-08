// D5/D6 on the real workbooks (XLN_CORPUS), in memory only: build with the source embedded
// and module names tagged, read back; the part carries the project byte for byte, and a
// pull gives it back from the workbook; a second build is up to date; no tag in any pulled file.
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { buildLockfile, buildWorkbook, LOCK_FILE, lockfileJson, moduleOfPath, parseLockfile, parseModule, pullProject, readEmbeddedSource, readWorkbook, splitProvenance, stringifyJson, unbuiltEdits, type BuildResult } from "../../src/index.js";
import { Package } from "../../src/file/package.js";

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

function why(r: BuildResult): string {
  return [...r.plan.problems.filter((p) => p.severity === "error").map((p) => p.message), ...r.plan.conflicts.map((c) => c.message), ...(r.readBack?.problems ?? []), r.error ?? ""].join("\n");
}

const source = (files: Record<string, string>) => Object.fromEntries(Object.entries(files).filter(([p]) => (p.startsWith("names/") && p.endsWith(".xln")) || p === LOCK_FILE));

describe.skipIf(!CORPUS)("embedded source and provenance on the corpus (XLN_CORPUS)", () => {
  const files = CORPUS ? workbooks(CORPUS) : [];
  for (const path of files) {
    it(path.slice(CORPUS!.length + 1), () => {
      const bytes = new Uint8Array(readFileSync(path));
      const project = { ...pullProject(bytes, "book.xlsx").files };
      expect(unbuiltEdits({ workbook: bytes, fileName: "book.xlsx", files: project })).toEqual([]);
      const r = buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files: project }, { embed: true });
      expect(r.status, why(r)).toBe("built");
      expect(r.readBack!.ok, why(r)).toBe(true);
      const built = { ...project, ...r.files };
      expect(new Package(r.bytes!).names.filter((n) => /^customXml\/item\d+\.xml$/i.test(n))).toHaveLength(1);
      // Every name in a module file is tagged; none elsewhere.
      const tagged = readWorkbook(r.bytes!).definedNames.filter((d) => d.comment && splitProvenance(d.comment).tag);
      const inModules = Object.entries(project)
        .filter(([p]) => moduleOfPath(p) !== undefined)
        .reduce((n, [, t]) => n + parseModule(t).entries.filter((e) => e.cell === undefined).length, 0);
      expect(tagged.length).toBe(inModules);
      const back = pullProject(r.bytes!, "book.xlsx");
      expect(source(readEmbeddedSource(r.bytes!)!.files)).toEqual(source(built));
      expect(source(back.files)).toEqual(source(built));
      for (const [p, t] of Object.entries(back.files)) if (p.endsWith(".xln")) expect(t.includes("[xln "), p).toBe(false);
      expect(back.provenance.every((x) => x.state === "unchanged")).toBe(true);
      const again = buildWorkbook({ workbook: r.bytes!, fileName: "book.xlsx", files: back.files }, { embed: true });
      expect(again.status, why(again)).toBe("up-to-date");
    }, 60_000);
  }
});

// Lockfile format 3 on the corpus: a project an earlier xln left with a format-2
// lockfile (`sha256:` hashes) reads without drift or unbuilt edits, and the next build writes format 4.
describe.skipIf(!CORPUS)("format-2 lockfiles on the corpus (XLN_CORPUS)", () => {
  const files = CORPUS ? workbooks(CORPUS) : [];
  const v2Lock = (bytes: Uint8Array) => {
    const p = pullProject(bytes, "book.xlsx");
    return stringifyJson(lockfileJson(buildLockfile("book.xlsx", p.names, p.statements, "v2"))) + "\n";
  };
  for (const path of files) {
    it(path.slice(CORPUS!.length + 1), () => {
      const bytes = new Uint8Array(readFileSync(path));
      const project = { ...pullProject(bytes, "book.xlsx").files, [LOCK_FILE]: v2Lock(bytes) };
      const plain = buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files: project }, { embed: false, provenance: false });
      expect(plain.status, why(plain)).toBe("up-to-date");
      expect(plain.plan.excelChanges).toEqual([]);

      const r = buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files: project }, { embed: true });
      expect(r.status, why(r)).toBe("built");
      expect(parseLockfile(r.files![LOCK_FILE]!).format).toBe("xln.lock/4");
      expect(r.files![LOCK_FILE]).not.toContain("sha256:");
      // The pull's guard: a format-2 lockfile of the same state holds no source edit.
      expect(unbuiltEdits({ workbook: bytes, fileName: "book.xlsx", files: project })).toEqual([]);
      // And after the build, a pull writes format 4 and agrees with the built project.
      const back = pullProject(r.bytes!, "book.xlsx");
      expect(parseLockfile(back.files[LOCK_FILE]!).format).toBe("xln.lock/4");
      const again = buildWorkbook({ workbook: r.bytes!, fileName: "book.xlsx", files: back.files }, { embed: true });
      expect(again.status, why(again)).toBe("up-to-date");
    }, 120_000);
  }
});
