import { describe, expect, it } from "vitest";
import { quoteSheet, significant, tokenize, type Token } from "../../src/index.js";

/** `kind:text` for each significant token. */
function toks(src: string): string[] {
  return significant(tokenize(src)).map((t) => `${t.kind}:${t.text}`);
}
function one(src: string): Token {
  const ts = significant(tokenize(src));
  expect(ts).toHaveLength(1);
  return ts[0]!;
}

describe("tokenizer", () => {
  it("covers the source exactly, spans included", () => {
    const src = "_xlfn.LET(_xlpm.a, 'S 1'!$A$1:$B$2,\r\n  _xlpm.a * 2)";
    const all = tokenize(src);
    expect(all.map((t) => t.text).join("")).toBe(src);
    for (const t of all) expect(src.slice(t.start, t.end)).toBe(t.text);
  });

  it("strings with doubled quotes", () => {
    const t = one('"say ""hi"""');
    expect(t.kind).toBe("string");
    expect(t.value).toBe('say "hi"');
    expect(toks('"a"&"b"')).toEqual(['string:"a"', "op:&", 'string:"b"']);
  });

  it("quoted sheet names with an escaped quote", () => {
    const t = one("'It''s'!A1");
    expect(t).toMatchObject({ kind: "ref", value: "A1", refKind: "cell" });
    expect(t.qual).toMatchObject({ sheet: "It's", quoted: true, raw: "'It''s'!" });
  });

  it("unquoted, 3-D and external qualifiers", () => {
    expect(one("Sheet1!$B$3").qual).toMatchObject({ sheet: "Sheet1", quoted: false });
    expect(one("S1:S3!A1").qual).toMatchObject({ sheet: "S1", sheet2: "S3" });
    expect(one("'S 1:S 3'!A1").qual).toMatchObject({ sheet: "S 1", sheet2: "S 3" });
    expect(one("[1]Sheet1!A1").qual).toMatchObject({ book: "[1]", sheet: "Sheet1" });
    const ext = one("[0]!Rate");
    expect(ext).toMatchObject({ kind: "name", value: "Rate" });
    expect(ext.qual).toMatchObject({ book: "[0]" });
    expect(ext.qual?.sheet).toBeUndefined();
    expect(one("'[2]My Sheet'!B2").qual).toMatchObject({ book: "[2]", sheet: "My Sheet" });
    expect(one("'C:\\dir\\[Book.xlsx]Data'!A1").qual).toMatchObject({ book: "C:\\dir\\[Book.xlsx]", sheet: "Data" });
    expect(one("Assumpt!MarketShare_base")).toMatchObject({ kind: "name", value: "MarketShare_base" });
    expect(one("Sheet1!#REF!")).toMatchObject({ kind: "ref", refKind: "error", value: "#REF!" });
  });

  it("structured references", () => {
    expect(one("Tbl[[#This Row],[Col]]")).toMatchObject({ kind: "structref", value: "Tbl", inner: "[#This Row],[Col]" });
    expect(one("Assumptions[Revenue Y0]")).toMatchObject({ kind: "structref", value: "Assumptions", inner: "Revenue Y0" });
    expect(one("[@Col]")).toMatchObject({ kind: "structref", value: "", inner: "@Col" });
    expect(one("T[[#Headers],[A]:[B]]")).toMatchObject({ inner: "[#Headers],[A]:[B]" });
    // `'` escapes a bracket inside a column name.
    expect(one("T[Col'[x']]")).toMatchObject({ inner: "Col'[x']" });
    expect(one("Tbl[]")).toMatchObject({ kind: "structref", inner: "" });
  });

  it("error literals vs the spill operator", () => {
    for (const e of ["#N/A", "#DIV/0!", "#NAME?", "#REF!", "#VALUE!", "#NUM!", "#NULL!", "#SPILL!", "#CALC!", "#GETTING_DATA"]) {
      expect(one(e)).toMatchObject({ kind: "error", text: e });
    }
    expect(toks("A1#")).toEqual(["ref:A1", "op:#"]);
    expect(toks("'S1'!$B$1#")).toEqual(["ref:'S1'!$B$1", "op:#"]);
    expect(toks("Revenue#*2")).toEqual(["name:Revenue", "op:#", "op:*", "number:2"]);
  });

  it("implicit intersection @", () => {
    expect(toks("@A1:A10")).toEqual(["op:@", "ref:A1:A10"]);
  });

  it("trim-range dot references", () => {
    expect(one("A1:.B100")).toMatchObject({ kind: "ref", refKind: "area", value: "A1:.B100" });
    expect(one("A1.:B100")).toMatchObject({ kind: "ref", refKind: "area", value: "A1.:B100" });
    expect(one("A1.:.B100")).toMatchObject({ kind: "ref", refKind: "area", value: "A1.:.B100" });
    expect(one("A:.A")).toMatchObject({ kind: "ref", refKind: "cols" });
    expect(toks("A1.:INDEX(B:B,3)")).toEqual(["ref:A1", "op:.:", "name:INDEX", "(:(", "ref:B:B", ",:,", "number:3", "):)"]);
  });

  it("array constants", () => {
    expect(toks("{1,2;3,4}")).toEqual(["{:{", "number:1", ",:,", "number:2", ";:;", "number:3", ",:,", "number:4", "}:}"]);
    expect(toks('{"a",TRUE;#N/A,-1}')).toEqual(['{:{', 'string:"a"', ",:,", "bool:TRUE", ";:;", "error:#N/A", ",:,", "op:-", "number:1", "}:}"]);
  });

  it("numbers", () => {
    for (const n of ["1", "1.5", ".5", "1E5", "1e-3", "2.5E+10", "10"]) expect(one(n)).toMatchObject({ kind: "number", text: n });
    expect(toks("10%")).toEqual(["number:10", "op:%"]);
    expect(one("2A").kind).toBe("invalid");
  });

  it("cells, areas, whole rows and columns", () => {
    expect(one("$A$1")).toMatchObject({ refKind: "cell" });
    expect(one("XFD1048576")).toMatchObject({ refKind: "cell" });
    expect(one("XFE1")).toMatchObject({ kind: "name" }); // beyond the last column
    expect(one("A1:B2")).toMatchObject({ refKind: "area" });
    expect(one("A:A")).toMatchObject({ refKind: "cols" });
    expect(one("$A:$C")).toMatchObject({ refKind: "cols" });
    expect(one("1:1")).toMatchObject({ refKind: "rows" });
    expect(one("$2:$5")).toMatchObject({ refKind: "rows" });
    expect(one("'S1'!$B$1:$C$3")).toMatchObject({ refKind: "area", value: "$B$1:$C$3" });
  });

  it("function names that look like cells", () => {
    expect(toks("LOG10(100)")).toEqual(["name:LOG10", "(:(", "number:100", "):)"]);
    expect(toks("ATAN2(1,1)")[0]).toBe("name:ATAN2");
    expect(toks("log10(1)")[0]).toBe("name:log10");
    expect(toks("DAYS360(A1,B1)")[0]).toBe("name:DAYS360");
    expect(toks("T(A1)")[0]).toBe("name:T");
  });

  it("a cell called like a function (it holds a LAMBDA) is a reference", () => {
    expect(toks("C2(D2,E2)")).toEqual(["ref:C2", "(:(", "ref:D2", ",:,", "ref:E2", "):)"]);
    expect(toks("$C$2(1)")[0]).toBe("ref:$C$2");
    expect(toks("C$2(1)")[0]).toBe("ref:C$2");
    expect(toks("Sheet1!C2(1)")[0]).toBe("ref:Sheet1!C2");
    expect(toks("'S 1'!$C$2(1)")[0]).toBe("ref:'S 1'!$C$2");
    // A qualifier or `$` makes even a function's spelling a cell: LOG10 is a column and a row.
    expect(toks("Sheet1!LOG10(1)")[0]).toBe("ref:Sheet1!LOG10");
    expect(toks("$LOG$10(1)")[0]).toBe("ref:$LOG$10");
    // Not cells: names and built-ins.
    expect(toks("XFE1(1)")[0]).toBe("name:XFE1");
    expect(toks("Fn.C2(1)")[0]).toBe("name:Fn.C2");
  });

  it("names: dotted, prefixed, non-ASCII", () => {
    expect(one("Mod.Fn")).toMatchObject({ kind: "name", value: "Mod.Fn" });
    expect(one("Growλ")).toMatchObject({ kind: "name", value: "Growλ" });
    expect(one("_xlfn._xlws.SORT")).toMatchObject({ kind: "name" });
    expect(one("_xlpm.x")).toMatchObject({ kind: "name" });
    expect(one("\\path")).toMatchObject({ kind: "name" });
    expect(one("TRUE")).toMatchObject({ kind: "bool" });
    expect(toks("TRUE()")[0]).toBe("name:TRUE");
  });

  it("space as intersection vs layout", () => {
    expect(tokenize("A1:B5 B2:C3").map((t) => t.kind)).toEqual(["ref", "isect", "ref", "eof"]);
    expect(tokenize("Rows Cols").map((t) => t.kind)).toEqual(["name", "isect", "name", "eof"]);
    expect(tokenize("(A1) (B1)").map((t) => t.kind)).toEqual(["(", "ref", ")", "isect", "(", "ref", ")", "eof"]);
    expect(tokenize("A1 + B1").map((t) => t.kind)).toEqual(["ref", "ws", "op", "ws", "ref", "eof"]);
    expect(tokenize("SUM( A1 )").map((t) => t.kind)).toEqual(["name", "(", "ws", "ref", "ws", ")", "eof"]);
  });

  it("CR LF and LF inside formulas are layout", () => {
    expect(tokenize("LAMBDA(x,\r\n  x*2)").filter((t) => t.kind === "ws").map((t) => t.text)).toEqual(["\r\n  "]);
    expect(tokenize("LAMBDA(x,\n  x*2)").some((t) => t.kind === "isect")).toBe(false);
  });

  it("reports bad input as invalid tokens with a position", () => {
    const bad = tokenize('1+"open');
    expect(bad[2]).toMatchObject({ kind: "invalid", start: 2, message: "unterminated string" });
    expect(tokenize("'Sheet").find((t) => t.kind === "invalid")?.message).toBe("unterminated quoted sheet name");
  });
});

describe("quoteSheet", () => {
  it("quotes only when needed, unless asked", () => {
    expect(quoteSheet("Model")).toBe("Model");
    expect(quoteSheet("SCF recursive")).toBe("'SCF recursive'");
    expect(quoteSheet("It's")).toBe("'It''s'");
    expect(quoteSheet("S1")).toBe("'S1'"); // looks like a cell
    expect(quoteSheet("R1C1")).toBe("'R1C1'");
    expect(quoteSheet("0 Tables")).toBe("'0 Tables'");
    expect(quoteSheet("Model", true)).toBe("'Model'");
  });
});
