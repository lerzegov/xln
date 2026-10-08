// LAMBDA stored forms as Excel saves them (probe F9, probes/results/f9_lambda_mac.xlsx and
// lambda_optional_mac.xlsx): compile(decompile(x)) gives back x exactly, character for
// character, for every name and cell formula, and the display forms read as typed.
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { compile, decompile, formulaCursor, nameUses, parse, prettyPrint } from "../../src/index.js";
import { readWorkbookStrings } from "./xlsx-strings.js";

const RESULTS = join(import.meta.dirname, "..", "..", "..", "..", "probes", "results");
const F9 = readWorkbookStrings(join(RESULTS, "f9_lambda_mac.xlsx"));
const OPT = readWorkbookStrings(join(RESULTS, "lambda_optional_mac.xlsx"));

describe("F9: Excel-saved LAMBDA forms", () => {
  for (const wb of [F9, OPT]) {
    const names = wb.names.map((n) => n.name);
    for (const n of wb.names) {
      it(`name ${n.name} round-trips exactly`, () => {
        expect(compile(decompile(n.text, { names }), { names })).toBe(n.text);
      });
    }
    for (const c of wb.formulas) {
      it(`cell ${c.cell} round-trips exactly`, () => {
        expect(compile(decompile(c.text, { names }), { names })).toBe(c.text);
      });
    }
  }

  it("covers the forms the probe set out to measure", () => {
    expect(F9.names.length).toBe(13);
    expect(F9.formulas.length).toBe(27);
  });

  const def = (name: string) => F9.names.find((n) => n.name === name)!.text;

  it("optional parameters: `_xlop.` in the list, `_xlpm.` in the body, `[p]` on display", () => {
    expect(def("ALLOPT")).toBe("_xlfn.LAMBDA(_xlop.x,_xlop.y, IF(_xlfn.ISOMITTED(_xlpm.x), 1, _xlpm.x) * IF(_xlfn.ISOMITTED(_xlpm.y), 2, _xlpm.y))");
    expect(decompile(def("ALLOPT"))).toBe("LAMBDA([x],[y], IF(ISOMITTED(x), 1, x) * IF(ISOMITTED(y), 2, y))");
    expect(decompile(def("TWOOPT"))).toBe("LAMBDA(a,[b],[d], a + IF(ISOMITTED(b), 0, b) + IF(ISOMITTED(d), 0, d))");
    // Nested: an optional parameter of an inner LAMBDA, inside a LET, inside a LAMBDA.
    expect(decompile(def("OPTNEST"))).toBe("LAMBDA(x,[k], LET(m, IF(ISOMITTED(k), 2, k), LAMBDA(y,[z], x * m + y + IF(ISOMITTED(z), 0, z))))");
    const f = parse(def("TWOOPT")).body;
    expect(f.kind === "lambda" && f.params.map((p) => p.optional)).toEqual([false, true, true]);
  });

  it("the module sample's ANA.GROW compiles to Excel's text", () => {
    const typed = "LAMBDA(value, rate, [periods],\n    value * (1 + rate) ^ IF(ISOMITTED(periods), 1, periods)\n)";
    expect(compile(typed)).toBe("_xlfn.LAMBDA(_xlpm.value, _xlpm.rate, _xlop.periods,\n    _xlpm.value * (1 + _xlpm.rate) ^ IF(_xlfn.ISOMITTED(_xlpm.periods), 1, _xlpm.periods)\n)");
    expect(OPT.names[0]!.text).toBe("_xlfn.LAMBDA(_xlpm.value,_xlpm.rate,_xlop.periods, _xlpm.value * (1 + _xlpm.rate) ^ IF(_xlfn.ISOMITTED(_xlpm.periods), 1, _xlpm.periods))");
  });

  it("eta-reduced functions, recursion, immediate calls", () => {
    expect(decompile(def("ETAMAP"))).toBe("MAP(Sheet1!$A$1:$A$3, ABS)");
    expect(compile("BYROW(Sheet1!$A$1:$A$3, SUM)")).toBe(def("ETABYROW"));
    expect(def("FACT")).toBe("_xlfn.LAMBDA(_xlpm.n, IF(_xlpm.n <= 1, 1, _xlpm.n * FACT(_xlpm.n - 1)))");
    expect(compile("LAMBDA(n, IF(n <= 1, 1, n * FACT(n - 1)))", { names: ["FACT"] })).toBe(def("FACT"));
    expect(decompile(def("IMMED"))).toBe("LAMBDA(x, x + 1)(2)");
  });

  it("`_xlop.` parameters are parameters for scoping: no name uses, cursor locals", () => {
    expect(nameUses(parse(def("TWOOPT")).body)).toEqual([]);
    expect(nameUses(parse(def("OPTNEST")).body)).toEqual([]);
    const stored = def("ALLOPT");
    const c = formulaCursor(stored, stored.indexOf("1, _xlpm.x"));
    expect(c.locals.map((l) => [l.name, l.kind])).toEqual(expect.arrayContaining([["x", "lambda"], ["y", "lambda"]]));
  });

  it("prettyPrint keeps the stored and the display spelling of optional parameters", () => {
    expect(prettyPrint(def("ALLOPT"), { width: 30 })).toContain("_xlfn.LAMBDA(_xlop.x, _xlop.y,");
    expect(prettyPrint(decompile(def("ALLOPT")), { width: 30 })).toContain("LAMBDA([x], [y],");
  });
});
