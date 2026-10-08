// The strict stored-form check: silent on everything Excel itself saved, precise on
// deliberately wrong stored text.
import { existsSync, readdirSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { checkStoredForm, compile } from "../../src/index.js";
import { readWorkbookStrings } from "./xlsx-strings.js";

const ROOT = join(import.meta.dirname, "..", "..", "..", "..", "probes");
const fixtures = [
  ...readdirSync(join(ROOT, "results")).filter((f) => f.endsWith(".xlsx")).map((f) => join(ROOT, "results", f)),
  join(ROOT, "fixtures", "traps.xlsx"),
];

function sweep(files: string[]): { checked: number; problems: string[] } {
  let checked = 0;
  const problems: string[] = [];
  for (const f of files) {
    const wb = readWorkbookStrings(f);
    const names = wb.names.map((x) => x.name);
    const all: [string, string][] = [...wb.names.map((x): [string, string] => [`name ${x.name}`, x.text]), ...wb.formulas.map((c): [string, string] => [`${c.part}!${c.cell}`, c.text])];
    for (const [where, text] of all) {
      checked++;
      for (const d of checkStoredForm(text, { names })) problems.push(`${f.split("/").pop()} ${where}: ${d.code}`);
    }
  }
  return { checked, problems };
}

const codes = (text: string, names: string[] = []) => checkStoredForm(text, { names }).map((d) => d.code);
const messages = (text: string) => checkStoredForm(text).map((d) => d.message);

describe("checkStoredForm on files Excel saved", () => {
  it("the probe workbooks: only the deliberate bare SEQUENCE (probe F6) is caught", () => {
    const r = sweep(fixtures);
    expect(r.checked).toBeGreaterThan(300);
    expect(r.problems.sort()).toEqual(["probe_patched.xlsx name Z_Bare: wrong-prefix", "traps.xlsx name BareSeq: wrong-prefix"]);
  });

  const corpus = process.env.XLN_CORPUS;
  it.skipIf(!corpus)("the corpus (XLN_CORPUS): nothing", () => {
    const files = readdirSync(corpus!, { withFileTypes: true })
      .filter((d) => d.isDirectory())
      .flatMap((d) => {
        const dist = join(corpus!, d.name, "dist");
        return existsSync(dist) ? readdirSync(dist).filter((f) => f.endsWith(".xlsx") && !f.startsWith("~$")).map((f) => join(dist, f)) : [];
      });
    const r = sweep(files);
    expect(r.checked).toBeGreaterThan(1000);
    expect(r.problems).toEqual([]);
  });
});

describe("checkStoredForm on a cell called like a function", () => {
  it("is Excel's stored form (MyLambda!F2 of the corpus): not an unknown function", () => {
    expect(checkStoredForm("C2(D2,E2)")).toEqual([]);
    expect(checkStoredForm("'S 1'!$C$2(1)+LOG10(100)")).toEqual([]);
    expect(checkStoredForm("C2(_xlpm.x)").map((d) => d.code)).toEqual(["unbound-parameter"]);
  });
});

describe("checkStoredForm on wrong stored text", () => {
  it("brackets in a LAMBDA parameter list (the M3d bug)", () => {
    const bad = "_xlfn.LAMBDA(_xlpm.value,_xlpm.rate,[_xlpm.periods], _xlpm.value*(1+_xlpm.rate)^IF(_xlfn.ISOMITTED(_xlpm.periods),1,_xlpm.periods))";
    expect(codes(bad)).toEqual(["bracketed-parameter"]);
    expect(messages(bad)[0]).toBe("'[_xlpm.periods]': an optional LAMBDA parameter is stored as '_xlop.periods', without brackets (Excel drops the name otherwise)");
    const d = checkStoredForm(bad)[0]!;
    expect(bad.slice(d.start, d.end)).toBe("[_xlpm.periods]");
    expect(codes("_xlfn.LAMBDA([x], 1)")).toEqual(["bracketed-parameter"]);
  });

  it("parameters and variables: unprefixed, unbound, `_xlop.` in a body", () => {
    expect(codes("_xlfn.LAMBDA(x, x+1)")).toEqual(["unprefixed-parameter", "unprefixed-parameter"]);
    expect(codes("_xlfn.LET(x, 1, _xlpm.x)")).toEqual(["unprefixed-parameter"]);
    expect(codes("_xlfn.LET(_xlpm.x, 1, _xlpm.y)")).toEqual(["unbound-parameter"]);
    expect(codes("_xlpm.x+1")).toEqual(["unbound-parameter"]);
    // A LET binding cannot see itself.
    expect(codes("_xlfn.LET(_xlpm.x, _xlpm.x+1, _xlpm.x)")).toEqual(["unbound-parameter"]);
    expect(codes("_xlfn.LAMBDA(_xlop.p, _xlop.p)")).toEqual(["optional-use"]);
    expect(codes("_xlfn.LET(_xlop.p, 1, _xlpm.p)")).toEqual(["unprefixed-parameter"]);
    expect(codes("_xlfn.LAMBDA(_xlpm.f, _xlpm.g(1))")).toEqual(["unbound-parameter"]);
    // Bound through nesting: LET inside LAMBDA inside LET.
    expect(codes("_xlfn.LET(_xlpm.a, 1, _xlfn.LAMBDA(_xlpm.x,_xlop.y, _xlfn.LET(_xlpm.z, _xlpm.a+_xlpm.x, _xlpm.z*_xlpm.y))(2))")).toEqual([]);
  });

  it("display-only syntax", () => {
    expect(codes("SUM(A1#)")).toEqual(["display-syntax"]);
    expect(codes("@A1:A3")).toEqual(["display-syntax"]);
    expect(codes("Table1[@Col]")).toEqual(["display-syntax"]);
    expect(codes("SUM(_xlfn.ANCHORARRAY(A1))+_xlfn.SINGLE(A1:A3)+Table1[[#This Row],[Col]]")).toEqual([]);
  });

  it("functions: unknown, missing or wrong prefix, defined names called", () => {
    expect(codes("SEQUENCE(3)")).toEqual(["wrong-prefix"]);
    expect(messages("SEQUENCE(3)")).toEqual(["the function 'SEQUENCE' must be stored as '_xlfn.SEQUENCE'"]);
    expect(codes("_xlfn.SUM(1)")).toEqual(["wrong-prefix"]);
    expect(codes("LAMBDA(_xlpm.x, 1)")).toEqual(["wrong-prefix"]);
    expect(codes("NOSUCHFN(1)")).toEqual(["unknown-function"]);
    expect(codes("_xlfn.NOSUCHFN(1)")).toEqual(["unknown-function"]);
    expect(codes("_xlfn.MAP(A1:A3, _xleta.NOSUCH)")).toEqual(["unknown-function"]);
    expect(codes("ANA.GROW(1, 2)", ["ANA.GROW"])).toEqual([]);
    expect(codes("_xludf.SEQUENCE(3)")).toEqual([]);
    expect(codes("[1]!Ext(1)")).toEqual([]);
  });

  it("whatever the compiler emits passes", () => {
    for (const display of [
      "LAMBDA(value, rate, [periods], value * (1 + rate) ^ IF(ISOMITTED(periods), 1, periods))",
      "LET(f, LAMBDA(t, [u], t * 3), f(2))",
      "MAP(A1:A3, ABS) + SUM(B1#) + @C1:C3",
      "LAMBDA(n, IF(n <= 1, 1, n * FACT(n - 1)))",
    ]) {
      expect(checkStoredForm(compile(display, { names: ["FACT"] }), { names: ["FACT"] })).toEqual([]);
    }
  });
});
