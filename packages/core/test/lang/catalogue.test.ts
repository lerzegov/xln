import { describe, expect, it } from "vitest";
import { arityProblem, catalogue, lookupFunction, MAX_ARGS, paramIndex, parseCatalogue, parseParams, parseXlm, xlmCollision } from "../../src/index.js";
import { XLM_COMMANDS, XLM_FUNCTIONS } from "../../src/lang/xlm-data.js";

// The sets in xlformula.py, whose prefixes were confirmed in saved workbooks.
const XLFN = [
  "ANCHORARRAY", "ARRAYTOTEXT", "BYCOL", "BYROW", "CHOOSECOLS", "CHOOSEROWS",
  "CONCAT", "DROP", "EXPAND", "FORMULATEXT", "GROUPBY", "HSTACK", "IFS", "ISOMITTED",
  "LAMBDA", "LET", "MAKEARRAY", "MAP", "MAXIFS", "MINIFS", "PERCENTOF",
  "PIVOTBY", "RANDARRAY", "REDUCE", "SCAN", "SEQUENCE", "SINGLE", "SORTBY",
  "SWITCH", "TAKE", "TEXTAFTER", "TEXTBEFORE", "TEXTJOIN", "TEXTSPLIT",
  "TOCOL", "TOROW", "TRIMRANGE", "UNIQUE", "VALUETOTEXT", "VSTACK",
  "WRAPCOLS", "WRAPROWS", "XLOOKUP", "XMATCH",
];
const XLWS = ["FILTER", "SORT"];
const LEGACY = [
  "ABS", "AND", "AVERAGE", "COLUMN", "COLUMNS", "COUNT", "COUNTA", "COUNTIF",
  "EDATE", "EOMONTH", "ERROR.TYPE", "EXACT", "EXP", "IF", "IFERROR", "INDEX", "ISERROR", "ISNA", "ISNUMBER", "LN", "MATCH",
  "MAX", "MIN", "N", "NA", "NOT", "OFFSET", "OR", "POWER", "ROUND", "ROW",
  "ROWS", "SUBSTITUTE", "SUM", "SUMPRODUCT", "T", "TEXT", "TRANSPOSE", "YEAR",
];
const ETA = ["AVERAGE", "COUNT", "COUNTA", "MAX", "MIN", "PERCENTOF", "SUM"];

describe("function catalogue", () => {
  it("parses, and has the size we expect", () => {
    expect(catalogue().size).toBeGreaterThan(450);
  });

  it("agrees with xlformula.py on every prefix it knows", () => {
    for (const n of XLFN) expect(lookupFunction(n)?.prefix, n).toBe("_xlfn.");
    for (const n of XLWS) expect(lookupFunction(n)?.prefix, n).toBe("_xlfn._xlws.");
    for (const n of LEGACY) expect(lookupFunction(n)?.prefix, n).toBe("");
    for (const n of ETA) expect(lookupFunction(n), n).toBeDefined();
  });

  it("looks up case-insensitively", () => {
    expect(lookupFunction("sequence")?.name).toBe("SEQUENCE");
    expect(lookupFunction("Stdev.S")).toMatchObject({ prefix: "_xlfn.", since: "2010" });
    expect(lookupFunction("NoSuchFn")).toBeUndefined();
  });

  it("lists the entries whose prefix is not confirmed", () => {
    const unsure = [...catalogue().values()].filter((f) => f.prefixUnsure).map((f) => f.name);
    expect(unsure).toEqual([
      "ECMA.CEILING", "ISO.CEILING", "DBCS", "COPILOT", "DETECTLANGUAGE", "FIELDVALUE", "IMAGE", "PY",
      "REGEXEXTRACT", "REGEXREPLACE", "REGEXTEST", "STOCKHISTORY", "TRANSLATE",
    ]);
  });

  it("marks the functions Excel writes for operators as internal (@, #, the trim references of probe F10)", () => {
    const internal = [...catalogue().values()].filter((f) => f.internal).map((f) => f.name);
    expect(internal.sort()).toEqual(["ANCHORARRAY", "SINGLE", "_TRO_ALL", "_TRO_LEADING", "_TRO_TRAILING"]);
  });

  it("rejects malformed lines", () => {
    expect(() => parseCatalogue("SUM - 1")).toThrow(/expected 5 or 6 columns/);
    expect(() => parseCatalogue("Sum - 1 * 2007")).toThrow(/upper case/);
    expect(() => parseCatalogue("SUM x 1 * 2007")).toThrow(/prefix/);
    expect(() => parseCatalogue("SUM - 2 1 2007")).toThrow(/arity/);
    expect(() => parseCatalogue("SUM - 1 * 1999")).toThrow(/version/);
    expect(() => parseCatalogue("SUM - 1 * 2007 maybe")).toThrow(/flag/);
    expect(() => parseCatalogue("SUM - 1 * 2007\nSUM - 1 * 2007")).toThrow(/duplicate/);
  });
});

describe("parameter names (catalogue-params.ts)", () => {
  const all = [...catalogue().values()];

  it("covers every function a user can type", () => {
    expect(all.filter((f) => !f.internal && !f.params).map((f) => f.name)).toEqual([]);
    expect(all.filter((f) => f.internal && f.params).map((f) => f.name)).toEqual([]);
  });

  it("lists the entries not confirmed by a Microsoft page", () => {
    expect(all.filter((f) => f.paramsUnsure).map((f) => f.name)).toEqual([
      "LEFTB", "LENB", "MIDB", "RIGHTB", "ECMA.CEILING", "COPILOT",
    ]);
  });

  it("agrees with the arity, except where Microsoft's brackets knowingly differ", () => {
    const differ = all.filter((f) => f.params && arityProblem(f.params, f)).map((f) => f.name);
    expect(differ).toEqual(["PROB", "REDUCE", "SCAN"]);
    for (const f of all) {
      if (!f.params || differ.includes(f.name)) continue;
      const args = f.params.filter((p) => p !== "...");
      expect(args.filter((p) => !p.startsWith("[")).length, f.name).toBe(f.minArgs);
      if (f.params.includes("...")) expect(f.maxArgs, f.name).toBe(MAX_ARGS);
      else expect(args.length, f.name).toBe(f.maxArgs);
    }
  });

  it("has the shapes signature help relies on", () => {
    expect(lookupFunction("XLOOKUP")?.params).toEqual([
      "lookup_value", "lookup_array", "return_array", "[if_not_found]", "[match_mode]", "[search_mode]",
    ]);
    expect(lookupFunction("SUM")?.params).toEqual(["number1", "[number2]", "..."]);
    expect(lookupFunction("PI")?.params).toEqual([]);
    expect(lookupFunction("LINEST")?.params?.[0]).toBe("known_y's");
    expect(lookupFunction("LAMBDA")?.params).toEqual(["[parameter1]", "[parameter2]", "...", "calculation"]);
  });

  it("maps arguments to parameters, repeating groups", () => {
    const sumifs = lookupFunction("SUMIFS")!.params!;
    const at = (params: string[], n: number, count?: number) => params[paramIndex(params, n, count)];
    expect([0, 1, 2, 3, 4, 5, 6].map((k) => at(sumifs, k))).toEqual([
      "sum_range", "criteria_range1", "criteria1", "[criteria_range2]", "[criteria2]", "[criteria_range2]", "[criteria2]",
    ]);
    const sum = lookupFunction("SUM")!.params!;
    expect(at(sum, 7)).toBe("[number2]");
    const ifs = lookupFunction("IFS")!.params!;
    expect([6, 7].map((k) => at(ifs, k))).toEqual(["[logical_test3]", "[value_if_true3]"]);
    const lambda = lookupFunction("LAMBDA")!.params!;
    expect([0, 1, 2, 3].map((k) => at(lambda, k, 4))).toEqual(["[parameter1]", "[parameter2]", "[parameter2]", "calculation"]);
    expect(at(lambda, 0, 1)).toBe("calculation");
    expect(at(lookupFunction("MAP")!.params!, 2, 3)).toBe("lambda");
    expect(paramIndex(lookupFunction("ABS")!.params!, 1)).toBe(-1);
  });

  it("rejects malformed lines and unknown names", () => {
    const fns = parseCatalogue("SUM - 1 * 2007\nABS - 1 1 2007\nSINGLE xlfn 1 1 2021 internal");
    expect(parseParams("SUM(number1, [number2], ...)\nABS(number)", fns).get("SUM")?.params).toEqual(["number1", "[number2]", "..."]);
    expect(() => parseParams("NOSUCH(x)", fns)).toThrow(/not in the catalogue/);
    expect(() => parseParams("SINGLE(x)", fns)).toThrow(/internal/);
    expect(() => parseParams("ABS number", fns)).toThrow(/NAME\(params\)/);
    expect(() => parseParams("ABS(number) maybe", fns)).toThrow(/flag/);
    expect(() => parseParams("ABS(num ber)", fns)).toThrow(/bad parameter/);
    expect(() => parseParams("ABS(number)\nABS(number)", fns)).toThrow(/duplicate/);
    expect(() => parseParams("ABS(number, [x])", fns)).toThrow(/MAX is 1/);
    expect(() => parseParams("ABS([number])", fns)).toThrow(/MIN is 1/);
    expect(() => parseParams("SUM(number1, [number2])", fns)).toThrow(/MAX is \*/);
    expect(() => parseParams("SUM(number1, ..., [x], ...)", fns)).toThrow(/more than one/);
    expect(() => parseParams("ABS(number)  !arity", fns)).toThrow(/arity agrees/);
    expect(parseParams("ABS([number])  !arity ?", fns).get("ABS")).toMatchObject({ unsure: true, arityDiffers: true });
  });
});

describe("Excel 4.0 macro names (xlm-data.ts)", () => {
  it("parses, and holds no worksheet function (those are the built-in rule's)", () => {
    const functions = parseXlm(XLM_FUNCTIONS);
    const commands = parseXlm(XLM_COMMANDS);
    expect(functions.size).toBeGreaterThan(100);
    expect(commands.size).toBeGreaterThan(350);
    for (const n of ["GROUP", "GET.CELL", "EVALUATE", "FILES"]) expect(functions.has(n), n).toBe(true);
    for (const n of ["OPEN", "SAVE", "SAVE.AS", "COPY", "TABLE", "SCALE", "A1.R1C1"]) expect(commands.has(n), n).toBe(true);
    expect([...functions, ...commands].filter((n) => lookupFunction(n))).toEqual([]);
    expect(() => parseXlm("0x00F5 Group")).toThrow(/bad name/);
    expect(() => parseXlm("GROUP")).toThrow(/expected/);
  });

  it("matches macro functions case-insensitively, never a command equivalent or a worksheet function", () => {
    expect(xlmCollision("Group")).toBe("GROUP");
    expect(xlmCollision("get.cell")).toBe("GET.CELL");
    expect(xlmCollision("Evaluate")).toBe("EVALUATE");
    expect(xlmCollision("Files")).toBe("FILES");
    for (const n of ["Open", "Save", "Save.As", "Save.As?", "Copy", "Table", "Scale"]) expect(xlmCollision(n), n).toBeUndefined();
    expect(xlmCollision("Sort")).toBeUndefined();
    expect(xlmCollision("Info")).toBeUndefined();
    expect(xlmCollision("Groups")).toBeUndefined();
  });

  it("still matches a function that is also a command (Ftab wins)", () => {
    const functions = parseXlm(XLM_FUNCTIONS);
    // None as transcribed; this holds the rule for any name added to both tables.
    const both = [...parseXlm(XLM_COMMANDS)].filter((n) => functions.has(n));
    for (const n of both) expect(xlmCollision(n), n).toBe(n);
  });
});
