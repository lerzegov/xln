import { createHash } from "node:crypto";
import { describe, expect, it } from "vitest";
import {
  classify,
  compressCells,
  definitionHash,
  extentSize,
  fileSystemKey,
  formatEntry,
  formatScope,
  formulaToSource,
  moduleFileName,
  nameUses,
  parse,
  parseModule,
  proposeModules,
  sha256,
  sheetFileName,
  sheetFromFileName,
  sourceToFormula,
  stringifyJson,
} from "../../src/index.js";

describe("sha256", () => {
  it("matches the standard test vectors", () => {
    expect(sha256("")).toBe("e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855");
    expect(sha256("abc")).toBe("ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad");
    expect(sha256("abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq")).toBe(
      "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1",
    );
    // UTF-8 and padding edge cases, against Node's implementation
    for (const s of ["é", "Growλ", "😀 x", "a".repeat(55), "a".repeat(56), "a".repeat(64), "a".repeat(1000)]) {
      expect(sha256(s), s).toBe(createHash("sha256").update(s, "utf8").digest("hex"));
    }
  });

  it("hashes definitions modulo whitespace and sheet quotes", () => {
    expect(definitionHash("_xlfn.LAMBDA(_xlpm.b,_xlpm.g, _xlpm.b*(1+_xlpm.g))")).toBe(
      definitionHash("_xlfn.LAMBDA(_xlpm.b, _xlpm.g,\r\n  _xlpm.b * (1 + _xlpm.g))"),
    );
    expect(definitionHash("'BS'!$A$1")).toBe(definitionHash("BS!$A$1"));
    expect(definitionHash("A1 B1")).not.toBe(definitionHash("A1B1"));
    expect(definitionHash("1+2")).not.toBe(definitionHash("1+3"));
  });
});

describe("classify", () => {
  const k = (s: string) => classify(s);
  it("constants", () => {
    for (const s of ["5", "-5", "0.1", '"text"', "TRUE", "#N/A", "{1,2;3,4}", "50%", "(3)"]) expect(k(s).kind, s).toBe("constant");
  });
  it("ranges", () => {
    for (const s of ["'S1'!$B$1", "Sheet1!$A$1:$B$4", "$A:$A", "('S'!A1,'S'!B2)", "S1:S3!A1", "'S'!#REF!"]) expect(k(s).kind, s).toBe("range");
  });
  it("spills, with the anchor", () => {
    expect(k("_xlfn.ANCHORARRAY('S1'!$B$1)")).toEqual({ kind: "spill", anchor: { sheet: "S1", cell: "B1" } });
    expect(k("Model!$C$11#")).toEqual({ kind: "spill", anchor: { sheet: "Model", cell: "C11" } });
    expect(k("$C$11#")).toEqual({ kind: "spill", anchor: { sheet: undefined, cell: "C11" } });
    expect(k("Other#").kind).toBe("spill");
  });
  it("Table references", () => {
    expect(k("tblRates[base]")).toEqual({ kind: "table", table: "tblRates" });
    expect(k("tblRates[[#Headers],[base]]").kind).toBe("table");
  });
  it("LAMBDA with arity and optional parameters", () => {
    expect(k("_xlfn.LAMBDA(_xlpm.b,_xlpm.g, _xlpm.b*(1+_xlpm.g))")).toEqual({
      kind: "lambda",
      arity: { required: 2, optional: 0 },
      params: ["b", "g"],
    });
    expect(k("LAMBDA(x, [y], x + IF(ISOMITTED(y), 0, y))")).toMatchObject({ arity: { required: 1, optional: 1 }, params: ["x", "[y]"] });
    expect(k("LAMBDA(42)")).toMatchObject({ kind: "lambda", arity: { required: 0, optional: 0 } });
  });
  it("formulas, including aliases and invoked LAMBDAs", () => {
    for (const s of ["Rate*2", "OtherName", "IN.SET(\"periods\")", "LAMBDA(x, x)(1)", "_xlfn.VSTACK(a, b)", "OFFSET(A1,0,0,3)", "A1+1"]) {
      expect(k(s).kind, s).toBe("formula");
    }
  });
  it("unparsed", () => {
    expect(k("SUM(1,")).toMatchObject({ kind: "unparsed" });
  });
});

describe("nameUses", () => {
  const uses = (s: string) => nameUses(parse(s).body).map((u) => (u.sheet ? `${u.sheet}!${u.id}` : u.id));
  it("finds names and LAMBDA names called as functions", () => {
    expect(uses("Rate*2 + FN.PICK(k, tbl[a], tbl[b]) + 'S2'!Loc")).toEqual(["Rate", "FN.PICK", "k", "S2!Loc"]);
  });
  it("respects LET and LAMBDA shadowing, in display and stored form", () => {
    expect(uses("LET(Rate, 5, Rate*2)")).toEqual([]);
    expect(uses("_xlfn.LET(_xlpm.Rate, 5, _xlpm.Rate*2)")).toEqual([]);
    expect(uses("LET(Rate, Rate*2, Rate)")).toEqual(["Rate"]);
    expect(uses("LAMBDA(f, f(1) + g(1))")).toEqual(["g"]);
  });
  it("ignores strings, built-ins, external and 3-D references", () => {
    expect(uses('"Rate is " & SUM(x) & [1]Book!Name & S1:S3!Name')).toEqual(["x"]);
  });
});

describe("proposeModules", () => {
  it("groups by dot prefix and by clear upper-case underscore prefix", () => {
    const m = proposeModules(["FN.PICK", "fn.Other", "IN_Rate", "IN_Tax", "IN_Years", "BS_base", "BS_payout", "BS_deleverage", "Q_A", "Q_B", "Plain", "Mod.Sub.f"]);
    expect(m.get("FN.PICK")).toBe("FN");
    expect(m.get("fn.Other")).toBe("FN");
    expect(m.get("IN_Rate")).toBe("IN");
    expect(m.get("BS_base")).toBeUndefined();
    expect(m.get("Q_A")).toBeUndefined(); // only two names
    expect(m.get("Plain")).toBeUndefined();
    expect(m.get("Mod.Sub.f")).toBe("Mod");
  });
  it("does not depend on input order", () => {
    const a = ["b.X", "B.y", "A_One", "A_Two", "A_Three"];
    expect([...proposeModules(a)].sort()).toEqual([...proposeModules([...a].reverse())].sort());
  });
  it("makes safe file names", () => {
    expect(moduleFileName("FN")).toBe("FN.xln");
    expect(moduleFileName(undefined)).toBe("_unmanaged.xln");
    expect(moduleFileName("A?\\b")).toBe("A__b.xln");
  });
});

describe("sheet file names", () => {
  it("keep spaces and letters, escape what file systems refuse", () => {
    expect(sheetFileName("BS")).toBe("BS.xln");
    expect(sheetFileName("SCF recursive")).toBe("SCF recursive.xln");
    expect(sheetFileName("Crescità Δ")).toBe("Crescità Δ.xln");
    expect(sheetFileName('a/b\\c:d*e?f"g<h>i|j')).toBe("a%2Fb%5Cc%3Ad%2Ae%3Ff%22g%3Ch%3Ei%7Cj.xln");
    expect(sheetFileName("100%")).toBe("100%25.xln");
    expect(sheetFileName("tab\there")).toBe("tab%09here.xln");
    expect(sheetFileName(".hidden")).toBe("%2Ehidden.xln");
    expect(sheetFileName("end. ")).toBe("end.%20.xln");
    expect(sheetFileName("dot.")).toBe("dot%2E.xln");
    expect(sheetFileName("CON")).toBe("%43ON.xln");
    expect(sheetFileName("lpt1")).toBe("%6Cpt1.xln");
    expect(sheetFileName("CONTROL")).toBe("CONTROL.xln");
  });
  it("are reversible", () => {
    for (const s of ["BS", "SCF recursive", 'a/b\\c:d*e?f"g<h>i|j', "100%", "%25", ".x.", " x ", "CON", "nul", "Crescità", "∑|∏", "It's", "\u{1F600}<"]) {
      expect(sheetFromFileName(sheetFileName(s)), s).toBe(s);
    }
    expect(sheetFromFileName("x%3cy.xln")).toBe("x<y"); // hex in either case
    expect(sheetFromFileName("bad%zz.xln")).toBeUndefined();
    expect(sheetFromFileName("CON.xln")).toBeUndefined(); // not something sheetFileName writes
    expect(sheetFromFileName("%41.xln")).toBeUndefined(); // needless escape
    expect(sheetFromFileName("BS.txt")).toBeUndefined();
  });
  it("compare as macOS and Windows file systems do", () => {
    expect(fileSystemKey("BS.xln")).toBe(fileSystemKey("bs.xln"));
    expect(fileSystemKey("Caf\u00e9.xln")).toBe(fileSystemKey("Cafe\u0301.xln"));
    expect(fileSystemKey("A.xln")).not.toBe(fileSystemKey("B.xln"));
  });
});

describe("parseModule", () => {
  it("applies @scope to every name after it, until the next @scope or @workbook", () => {
    const text = [
      "A = 1;",
      "@scope(BS)",
      "",
      "/** doc */",
      "B = 2;",
      "@hidden",
      "C = 3;",
      "@scope('SCF recursive')",
      "D = 4;",
      "@workbook",
      "E = 5;",
      "@scope(It's)   // unquoted works too when nothing needs quoting",
      "F = 6;",
    ].join("\n");
    const { entries, diagnostics } = parseModule(text);
    expect(diagnostics).toEqual([]);
    expect(entries.map((e) => [e.name, e.scope, e.hidden])).toEqual([
      ["A", undefined, false],
      ["B", "BS", false],
      ["C", "BS", true],
      ["D", "SCF recursive", false],
      ["E", undefined, false],
      ["F", "It's", false],
    ]);
    expect(entries[1]!.doc).toBe("doc");
    // @hidden is per name and stays among the annotations; the scope blocks do not.
    expect(entries[2]!.annotations.map((a) => a.name)).toEqual(["hidden"]);
    expect(entries[1]!.annotations).toEqual([]);
  });

  it("reports a bare @scope and an argument to @workbook", () => {
    const { entries, diagnostics } = parseModule("@scope()\nA = 1;\n@workbook(BS)\nB = 2;");
    expect(diagnostics.map((d) => [d.line, d.message])).toEqual([
      [1, "@scope needs a sheet name"],
      [3, "@workbook takes no argument"],
    ]);
    expect(entries.map((e) => e.scope)).toEqual([undefined, undefined]);
  });

  it("quotes the sheet in @scope by Excel's rule, as formulas do (spec §14 issue 7)", () => {
    expect(formatScope("BS")).toBe("@scope(BS)");
    expect(formatScope("A1")).toBe("@scope('A1')");
    expect(formatScope("x<y>|z.")).toBe("@scope('x<y>|z.')");
    expect(formatScope("SCF recursive")).toBe("@scope('SCF recursive')");
    expect(formatScope("a(b)")).toBe("@scope('a(b)')");
    expect(formatScope("It's")).toBe("@scope('It''s')");
    expect(formatScope(undefined)).toBe("@workbook");
  });

  it("reads definitions, doc comments, annotations and types", () => {
    const text = [
      "// module: Fin",
      "/** Period axis. */",
      "Timeline = LAMBDA(start, periods, SEQUENCE(1, periods, start, 1));",
      "",
      "@scope('Cash Flow')   // sheet-scoped",
      "@hidden",
      "Revenue = 'Cash Flow'!$C$10#;",
      "/* not a doc */",
      "Literals = IF(TRUE, \"a ; b // c\", {1,2;3,4});",
      "Multi = LET(",
      "    x, 1,   // first",
      "    /* inline */ x + 1",
      ");",
    ].join("\n");
    const { entries, diagnostics } = parseModule(text);
    expect(diagnostics).toEqual([]);
    expect(entries.map((e) => e.name)).toEqual(["Timeline", "Revenue", "Literals", "Multi"]);
    expect(entries[0]).toMatchObject({ doc: "Period axis.", line: 3, hidden: false });
    expect(entries[1]).toMatchObject({ scope: "Cash Flow", hidden: true, formula: "'Cash Flow'!$C$10#" });
    expect(entries[2]!.doc).toBeUndefined();
    expect(entries[2]!.formula).toBe('IF(TRUE, "a ; b // c", {1,2;3,4})');
    expect(entries[3]!.formula).toBe("LET(\n    x, 1,\n      x + 1\n)");
  });

  it("rejects a type declaration, on the type (not xln v1: planned with the dimension layer)", () => {
    const { entries, diagnostics } = parseModule("A = 1;\nPeriods : scalar = Model!$B$3;\nB = 2;");
    expect(diagnostics).toEqual([{ line: 2, col: 11, message: "type declarations are not part of xln v1 (planned with the dimension layer)" }]);
    expect(entries.map((e) => e.name)).toEqual(["A", "Periods", "B"]);
  });

  it("module files: @sheet(Sheet) above a name makes it local to Sheet, per name; old blocks still read", () => {
    const pm = parseModule("A = 1;\n/** Doc. */\n@sheet('Cash Flow')\n@hidden\nB = 2;\nC = 3;\n@scope(S1)\nD = 4;\n@sheet(S2)\nE = 5;\n@workbook\nF = 6;");
    expect(pm.diagnostics).toEqual([]);
    expect(pm.entries.map((e) => [e.name, e.scope ?? null])).toEqual([["A", null], ["B", "Cash Flow"], ["C", null], ["D", "S1"], ["E", "S2"], ["F", null]]);
    expect(pm.entries[1]).toMatchObject({ doc: "Doc.", hidden: true });
  });

  it("keeps structured references and quoted sheets intact", () => {
    const { entries } = parseModule("A = T[[#This Row],[Q1 '[x']]] + 'It''s; here'!A1;");
    expect(entries[0]!.formula).toBe("T[[#This Row],[Q1 '[x']]] + 'It''s; here'!A1");
  });

  it("reports problems with line and column", () => {
    const { entries, diagnostics } = parseModule("A = 1;\nB 2;\nC = 3");
    expect(entries.map((e) => e.name)).toEqual(["A", "C"]);
    expect(diagnostics.map((d) => d.line)).toEqual([2, 3]);
  });

  it("round-trips doc comments exactly", () => {
    for (const doc of ["plain", " padded ", "two\nlines", "\nleading blank", "has */ inside", "  indented\n  * star", "x\n\ny"]) {
      const text = formatEntry({ name: "N", doc, formula: "1" });
      const { entries } = parseModule(text);
      expect(entries[0]!.doc, JSON.stringify(doc)).toBe(doc);
    }
  });

  it("round-trips scope quoting", () => {
    for (const scope of ["S2", "Cash Flow", "It's", "Model", "a(b)", "x<y>|z.", "A1", "@x", "a;b", "Δ"]) {
      const { entries } = parseModule(formatEntry({ name: "N", scope, formula: "1" }));
      expect(entries[0]!.scope).toBe(scope);
    }
  });
});

describe("helpers", () => {
  it("compresses cell lists into rectangles", () => {
    expect(compressCells(["C2", "B2", "B3", "C3", "E1", "A10", "B10"])).toEqual(["E1", "B2:C3", "A10:B10"]);
    expect(compressCells(["B3", "B4", "B5", "C5"])).toEqual(["B3:B4", "B5:C5"]);
  });
  it("measures extents", () => {
    expect(extentSize("C4")).toEqual({ rows: 1, cols: 1 });
    expect(extentSize("C11:G11")).toEqual({ rows: 1, cols: 5 });
  });
  it("writes compact, stable JSON", () => {
    expect(stringifyJson({ a: [1, 2], b: { c: "x" }, d: undefined, e: [{ f: [] }] })).toBe(
      '{\n  "a": [1, 2],\n  "b": { "c": "x" },\n  "e": [\n    { "f": [] }\n  ]\n}',
    );
  });
});

describe("parseModule positions (for the editor)", () => {
  it("records where entries, names, scope blocks and formulas are", () => {
    const text = ["A = 1;", "@scope('S 1')", "/** doc */", "@hidden", "B  =  x + Y#;", "@workbook", "C = 2"].join("\n");
    const m = parseModule(text);
    const [a, b, c] = m.entries;
    expect(text.slice(a!.start, a!.end)).toBe("A = 1;");
    expect(text.slice(b!.start, b!.end)).toBe("/** doc */\n@hidden\nB  =  x + Y#;");
    expect(text.slice(b!.offset, b!.offset + 1)).toBe("B");
    expect(c!.end).toBe(text.length);
    const fs = formulaToSource(b!, 0);
    expect(text.slice(fs, fs + b!.formula.length)).toBe("x + Y#");
    expect(m.scopes.map((d) => [d.scope, d.line, text.slice(d.offset, d.end)])).toEqual([
      ["S 1", 2, "@scope('S 1')"],
      [undefined, 6, "@workbook"],
    ]);
  });

  it("maps formula offsets around removed comments, both ways", () => {
    const text = "M = LET(  // why\n  x, 1, /* inline */ x + Rate\n);";
    const e = parseModule(text).entries[0]!;
    expect(e.formula).toBe("LET(\n  x, 1,   x + Rate\n)");
    for (const word of ["LET", "x, 1", "x + Rate", ")"]) {
      const i = e.formula.indexOf(word);
      const src = formulaToSource(e, i);
      expect(text.slice(src, src + word.length)).toBe(word);
      expect(sourceToFormula(e, src)).toBe(i);
    }
    expect(sourceToFormula(e, text.indexOf("why"))).toBeUndefined();
    expect(sourceToFormula(e, text.indexOf("inline"))).toBeUndefined();
    expect(sourceToFormula(e, 0)).toBeUndefined();
  });
});

describe("nameUses spans", () => {
  it("cover the identifier only", () => {
    const src = "'S 1'!Loc + Spl# + FN.X(1)";
    expect(nameUses(parse(src).body).map((u) => src.slice(u.span.start, u.span.end))).toEqual(["Loc", "Spl", "FN.X"]);
  });
});
