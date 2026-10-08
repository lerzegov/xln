// Round trip compile(decompile(x)) == x (modulo whitespace) and idempotence, on every
// definition and cell formula of the probe workbooks and, when XLN_CORPUS is set, the corpus.
import { readdirSync, existsSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { lookupFunction } from "../../src/index.js";
import { roundTrip } from "./roundtrip.js";
import { readWorkbookStrings } from "./xlsx-strings.js";

const RESULTS = join(import.meta.dirname, "..", "..", "..", "..", "probes", "results");
const fixtures = readdirSync(RESULTS).filter((f) => f.endsWith(".xlsx")).map((f) => join(RESULTS, f));

function corpusFiles(root: string): string[] {
  return readdirSync(root, { withFileTypes: true })
    .filter((d) => d.isDirectory())
    .flatMap((d) => {
      const dist = join(root, d.name, "dist");
      return existsSync(dist) ? readdirSync(dist).filter((f) => f.endsWith(".xlsx") && !f.startsWith("~$")).map((f) => join(dist, f)) : [];
    });
}

describe("round trip on the probe workbooks", () => {
  it.each(fixtures.map((f) => [f.split("/").pop()!, f]))("%s", (_label, path) => {
    const r = roundTrip(readWorkbookStrings(path), _label);
    expect(r.failures).toEqual([]);
    expect(r.quotingOnly).toEqual([]);
    expect(r.definitions + r.formulas).toBeGreaterThan(0);
    for (const fn of r.calledFunctions) expect(lookupFunction(fn), `${fn} is called but not in the catalogue`).toBeDefined();
  });

  it("only probe_patched's Z_Bare (SEQUENCE without _xlfn., probe F6) needs repair", () => {
    const repaired = fixtures.flatMap((f) => roundTrip(readWorkbookStrings(f), f.split("/").pop()!).repaired);
    expect(repaired).toEqual(["probe_patched.xlsx name Z_Bare: SEQUENCE(1,3)  →  _xlfn.SEQUENCE(1,3)"]);
  });
});

const corpus = process.env.XLN_CORPUS;
describe.skipIf(!corpus)("round trip on the corpus (XLN_CORPUS)", () => {
  it("every definition and cell formula round-trips", () => {
    const files = corpusFiles(corpus!);
    expect(files.length).toBeGreaterThan(0);
    let defs = 0;
    let fmls = 0;
    const failures: string[] = [];
    const repaired: string[] = [];
    const missing = new Set<string>();
    // Signature help for what the corpus calls should come from a Microsoft page.
    const unconfirmed = new Set<string>();
    const quoting: string[] = [];
    for (const f of files) {
      const r = roundTrip(readWorkbookStrings(f), f.slice(corpus!.length + 1));
      defs += r.definitions;
      fmls += r.formulas;
      failures.push(...r.failures);
      repaired.push(...r.repaired);
      quoting.push(...r.quotingOnly);
      for (const fn of r.calledFunctions) {
        const info = lookupFunction(fn);
        if (!info) missing.add(fn);
        else if (!info.internal && (!info.params || info.paramsUnsure)) unconfirmed.add(fn);
      }
    }
    console.log(
      `corpus: ${files.length} workbooks, ${defs} definitions, ${fmls} cell formulas; ` +
        `${failures.length} failures, ${repaired.length} repaired (F6), ${quoting.length} equal only modulo sheet quotes, functions missing from the catalogue: ${[...missing].join(", ") || "none"}`,
    );
    for (const x of [...failures, ...repaired, ...quoting.slice(0, 5)].slice(0, 30)) console.log(x);
    expect(failures).toEqual([]);
    expect([...missing]).toEqual([]);
    expect([...unconfirmed]).toEqual([]);
  });
});
