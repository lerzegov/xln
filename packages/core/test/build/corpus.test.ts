// Round-trip identity and an edit of every kind on the real workbooks (XLN_CORPUS).
// The corpus is only read: every build here stays in memory.
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { applyChangeSet, pullProject, readBack, readWorkbook, verifyValues } from "../../src/index.js";
import { build, definedNamesXml, rewriteAll, sameNames, why } from "./helpers.js";


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

describe.skipIf(!CORPUS)("build on the corpus (XLN_CORPUS)", () => {
  const files = CORPUS ? workbooks(CORPUS) : [];
  it("has workbooks", () => expect(files.length).toBeGreaterThan(0));
  for (const path of files) {
    it(path.slice(CORPUS!.length + 1), () => {
      const bytes = new Uint8Array(readFileSync(path));
      const project = () => ({ ...pullProject(bytes, "book.xlsx").files });

      // No edits: nothing to do; forced, <definedNames> byte-identical.
      const none = build(bytes, project());
      expect(none.status, why(none)).toBe("up-to-date");
      const forced = build(bytes, project(), true);
      expect(forced.status, why(forced)).toBe("built");
      expect(definedNamesXml(forced.bytes!) === definedNamesXml(bytes), "<definedNames> differs").toBe(true);
      expect(verifyValues(bytes, forced.bytes!).changed).toEqual([]);

      // Every name rewritten from source: semantically equal, reads back.
      const { changes, inSync } = rewriteAll(bytes);
      const all = applyChangeSet(bytes, changes);
      expect(sameNames(readWorkbook(bytes).definedNames, readWorkbook(all).definedNames)).toEqual([]);
      expect(readBack(bytes, all, changes, inSync).problems).toEqual([]);

      // One edit of each kind on the first workbook-scoped name nothing uses... or on new names.
      const p = project();
      const target = Object.keys(p).find((k) => k.startsWith("names/") && k.endsWith(".xln"))!;
      p[target] += "\n/** added by the corpus test\n * second line */\nXlnTestAdded = LET(a, 1,\n    b, a + 1,\n    b * 2\n);\n";
      const edited = build(bytes, p);
      expect(edited.status, why(edited)).toBe("built");
      expect(edited.plan.changeSet.changes).toHaveLength(1);
      expect(verifyValues(bytes, edited.bytes!).changed).toEqual([]);
      const again = build(edited.bytes!, { ...p, ...edited.files });
      expect(again.status, why(again)).toBe("up-to-date");
      // …then renamed, then deleted.
      const p2 = { ...p, ...edited.files };
      p2[target] = p2[target]!.replace("XlnTestAdded =", "@renamed(XlnTestAdded)\nXlnTestRenamed =");
      const renamed = build(edited.bytes!, p2);
      expect(renamed.status, why(renamed)).toBe("built");
      const p3 = { ...p2, ...renamed.files };
      p3[target] = p3[target]!.replace(/\/\*\* added by the corpus test[\s\S]*$/, "");
      const deleted = build(renamed.bytes!, p3);
      expect(deleted.status, why(deleted)).toBe("built");
      expect(definedNamesXml(deleted.bytes!) === definedNamesXml(bytes), "after add, rename, delete: <definedNames> differs").toBe(true);

      // M3b: one cell statement edited → exactly one set-cell-formula, written, read back.
      const pc = pullProject(bytes, "book.xlsx");
      const c = pc.report.cells;
      console.log(`${path.slice(CORPUS!.length + 1)}: cell statements ${c.named} named, ${c.unnamed} unnamed (${c.blocks} blocks over ${c.blockCells} cells), ${c.slots} slots`);
      const st = pc.statements.find((s) => s.kind === "named" && !s.spills && s.display !== "" && !s.display.includes("\n"));
      if (st) {
        const file = Object.keys(pc.files).find((f) => pc.files[f]!.includes(`${st.name} @`))!;
        const pe = { ...pc.files };
        const line = pe[file]!.split("\n").find((l) => l.startsWith(`${st.name} @`))!;
        pe[file] = pe[file]!.replace(line, line.replace(/;$/, " + 0;"));
        const cell = build(bytes, pe);
        expect(cell.status, why(cell)).toBe("built");
        expect(cell.plan.changeSet.changes.filter((x) => x.op === "set-cell-formula")).toHaveLength(1);
        expect(cell.readBack!.ok, cell.readBack!.problems.join("\n")).toBe(true);
        expect(build(cell.bytes!, { ...pe, ...cell.files }).status).toBe("up-to-date");
      }
    });
  }
});
