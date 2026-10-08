// Excel accepts definitions up to ~8,190 characters (probe T15): a chain like x+1+1+…
// is ~4,000 operators deep and must not overflow the stack anywhere.
import { describe, expect, it } from "vitest";
import { compile, decompile, equalModuloWhitespace, parse, prettyPrint, walk } from "../../src/index.js";

const stored = "_xlfn.LAMBDA(_xlpm.x,_xlpm.x" + "+1".repeat(4090) + ")";
const display = "LAMBDA(x,x" + "+1".repeat(4090) + ")";

describe("deep operator chains", () => {
  it("parse and walk", () => {
    let count = 0;
    walk(parse(stored).body, () => count++);
    expect(count).toBeGreaterThan(8000);
  });
  it("decompile", () => expect(decompile(stored)).toBe(display));
  it("compile", () => expect(compile(display)).toBe(stored));
  it("pretty-print", () => expect(equalModuloWhitespace(prettyPrint(display), display)).toBe(true));
});
