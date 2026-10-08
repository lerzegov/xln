// Round-trip identity (M3a acceptance): pull, then build with no edits.
//   - nothing to do: `up-to-date`, and a forced build leaves <definedNames> byte-identical,
//     changes only `fullCalcOnLoad`, and copies every other zip entry byte for byte;
//   - rewriting every name from its source (compile of the pulled display text) gives
//     <definedNames> semantically equal to the original.
import { readdirSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { applyChangeSet, pullProject, rawZipRecords, readBack, readWorkbook } from "../../src/index.js";
import { build, definedNamesXml, fixture, RESULTS, rewriteAll, sameNames, why, workbookXml } from "./helpers.js";

const NAMES = [...readdirSync(RESULTS).filter((f) => f.endsWith(".xlsx")).sort(), "traps.xlsx"];

/** Names stored without the prefix Excel needs (probe F6, seeded traps): a build repairs them. */
const REPAIRS: Record<string, string[]> = { "probe_patched.xlsx": ["Z_Bare"], "traps.xlsx": ["BareSeq"] };

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  return a.length === b.length && a.every((x, i) => x === b[i]);
}

describe("round trip on the probe workbooks", () => {
  for (const name of NAMES) {
    describe(name, () => {
      const bytes = fixture(name);
      const files = () => pullProject(bytes, "book.xlsx").files;
      const repairs = REPAIRS[name] ?? [];

      it("no edits: nothing to change but the F6 repairs", () => {
        const r = build(bytes, { ...files() });
        expect(r.plan.conflicts, why(r)).toEqual([]);
        expect(r.plan.problems.filter((p) => p.severity === "error"), why(r)).toEqual([]);
        if (repairs.length > 0) {
          expect(r.status, why(r)).toBe("built");
          expect(r.plan.changeSet.changes.map((c) => c.op === "set-name" && c.repair && c.name)).toEqual(repairs);
          // Once repaired, the next build has nothing to do.
          expect(build(r.bytes!, { ...files(), ...r.files }).status).toBe("up-to-date");
        } else {
          expect(r.status).toBe("up-to-date");
          expect(r.plan.changeSet.changes).toEqual([]);
        }
      });

      it("a forced build: <definedNames> byte-identical, other entries byte for byte, fullCalcOnLoad set", () => {
        // A workbook with a repair to make is compared once repaired.
        const base = repairs.length > 0 ? build(bytes, { ...files() }).bytes! : bytes;
        const r = build(base, { ...pullProject(base, "book.xlsx").files }, true);
        expect(r.status, why(r)).toBe("built");
        const out = r.bytes!;
        expect(definedNamesXml(out) === definedNamesXml(base), "<definedNames> differs").toBe(true);
        const before = rawZipRecords(base);
        const after = rawZipRecords(out);
        expect([...after.keys()]).toEqual([...before.keys()]);
        for (const [k, rec] of before) if (k !== "xl/workbook.xml") expect(bytesEqual(rec, after.get(k)!), k).toBe(true);
        expect(workbookXml(out)).toMatch(/<calcPr [^>]*fullCalcOnLoad="1"/);
        // Only the calcPr tag differs in the workbook part.
        const strip = (s: string) => s.replace(/<calcPr[^>]*>/, "");
        expect(strip(workbookXml(out)) === strip(workbookXml(base)), "workbook part differs outside calcPr").toBe(true);
        expect(r.readBack!.ok, r.readBack!.problems.join("\n")).toBe(true);
      });

      it("every name rewritten from source: semantically equal, reads back", () => {
        const { changes, inSync } = rewriteAll(bytes);
        const out = applyChangeSet(bytes, changes);
        expect(sameNames(readWorkbook(bytes).definedNames, readWorkbook(out).definedNames, repairs)).toEqual([]);
        const rb = readBack(bytes, out, changes, inSync);
        expect(rb.problems).toEqual([]);
      });
    });
  }
});
