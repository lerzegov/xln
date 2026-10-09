import { describe, expect, it } from "vitest";
import { compile, decompile, equalModuloWhitespace, hasLayoutBreaks, oneLine, prettyPrint } from "../../src/index.js";

describe("equalModuloWhitespace", () => {
  it("ignores layout, as Excel re-spaces (T14)", () => {
    expect(equalModuloWhitespace("LAMBDA(b, g, b*(1+g))", "LAMBDA(b,g, b*(1+g))")).toBe(true);
    expect(equalModuloWhitespace("LAMBDA(x,\r\n  x*2)", "LAMBDA(x, x*2)")).toBe(true);
    expect(equalModuloWhitespace("T[[#This Row], [A]]", "T[[#This Row],[A]]")).toBe(true);
  });
  it("does not ignore what matters", () => {
    expect(equalModuloWhitespace('"a b"', '"ab"')).toBe(false);
    expect(equalModuloWhitespace("A1:B5 B2:C3", "A1:B5,B2:C3")).toBe(false);
    expect(equalModuloWhitespace("A1+1", "A1+2")).toBe(false);
    expect(equalModuloWhitespace("T[Col A]", "T[ColA]")).toBe(false);
  });
  it("treats optional sheet quotes as equal", () => {
    expect(equalModuloWhitespace("BS!X", "'BS'!X")).toBe(true);
    expect(equalModuloWhitespace("BS!X", "IS!X")).toBe(false);
  });
});

describe("prettyPrint", () => {
  it("leaves short formulas on one line, normalising spaces", () => {
    expect(prettyPrint("LAMBDA(x,x+1)")).toBe("LAMBDA(x, x + 1)");
    expect(prettyPrint("=SUM(A1:A3)*2")).toBe("=SUM(A1:A3) * 2");
    expect(prettyPrint("{1,2;3,4}")).toBe("{1,2;3,4}");
  });

  it("puts a long LET one binding per line", () => {
    const src =
      "LET(base, Assumptions[Revenue Y0], g, Assumptions[Growth], n, Periods, base * (1 + g) ^ SEQUENCE(1, n, 0, 1))";
    expect(prettyPrint(src)).toBe(
      [
        "LET(",
        "    base, Assumptions[Revenue Y0],",
        "    g, Assumptions[Growth],",
        "    n, Periods,",
        "    base * (1 + g)^SEQUENCE(1, n, 0, 1)",
        ")",
      ].join("\n"),
    );
  });

  it("puts a long LAMBDA body on its own lines, nesting LETs", () => {
    const src =
      "LAMBDA(start, periods, LET(s, SEQUENCE(1, periods, start, 1), labels, TEXT(s, \"0\"), VSTACK(s, labels)))";
    expect(prettyPrint(src, { width: 60 })).toBe(
      [
        "LAMBDA(start, periods,",
        "    LET(",
        "        s, SEQUENCE(1, periods, start, 1),",
        '        labels, TEXT(s, "0"),',
        "        VSTACK(s, labels)",
        "    )",
        ")",
      ].join("\n"),
    );
  });

  it("breaks long argument lists", () => {
    const src = "IF(AND(Revenue > Threshold, Margin > MinimumMargin), Revenue * Margin, NA())";
    expect(prettyPrint(src, { width: 60 })).toBe(
      ["IF(", "    AND(Revenue > Threshold, Margin > MinimumMargin),", "    Revenue * Margin,", "    NA()", ")"].join("\n"),
    );
  });

  it("packs short arguments into lines when asked", () => {
    const src = "MAX(C6, C23, C28, C33, C38, C43, C48, C53, C58, C63, ABS(C68))";
    const packed = prettyPrint(src, { width: 30, pack: true });
    expect(packed).toBe(["MAX(", "    C6, C23, C28, C33, C38,", "    C43, C48, C53, C58, C63,", "    ABS(C68)", ")"].join("\n"));
    expect(equalModuloWhitespace(packed, src)).toBe(true);
    // A long argument keeps the one-per-line layout.
    const long = "IF(AND(Revenue > Threshold, Margin > MinimumMargin), 1, 2)";
    expect(prettyPrint(long, { width: 40, pack: true })).toBe(prettyPrint(long, { width: 40 }));
  });

  it("keeps intersection spaces, unions and the stored form intact", () => {
    for (const src of [
      "SUM(A1:B5 B2:C3)",
      "SUM((A1,B1))",
      "LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))(1)",
      "_xlfn.LET(_xlpm.a, 'S 1'!$A$1#, -_xlpm.a%^2&\"x\")",
      "@A1:A10",
      "Tbl[@[Col A]]*2",
    ]) {
      const pretty = prettyPrint(src, { width: 20 });
      expect(equalModuloWhitespace(pretty, src), pretty).toBe(true);
    }
  });

  it("pretty display text compiles to the same stored text", () => {
    const stored =
      "_xlfn.LET(_xlpm.base, Assumptions[Revenue Y0], _xlpm.g, Assumptions[Growth], _xlpm.base * (1 + _xlpm.g)^_xlfn.SEQUENCE(1, Periods, 0, 1))";
    const pretty = prettyPrint(decompile(stored), { width: 40 });
    expect(pretty.split("\n").length).toBeGreaterThan(3);
    expect(equalModuloWhitespace(compile(pretty, { names: ["Periods"] }), stored)).toBe(true);
  });
});

describe("oneLine (cells keep the workbook's layout)", () => {
  it("joins pull's layout the way the formula bar reads it", () => {
    expect(oneLine("LET(\n    x, A1*10,\n    y, 2,\n    x + y\n)")).toBe("LET(x, A1*10, y, 2, x + y)");
    expect(oneLine("LAMBDA(a, b,\n    a + b\n)(1, 2)")).toBe("LAMBDA(a, b, a + b)(1, 2)");
    expect(oneLine("IF(\r\n  A1 > 0,\r\n  1,\r\n  2\r\n)")).toBe("IF(A1 > 0, 1, 2)");
    expect(oneLine("\n  SUM(A1:A3)\n")).toBe("SUM(A1:A3)");
    expect(oneLine("{1,\n 2;\n 3,4}")).toBe("{1, 2;3,4}");
    expect(oneLine("A1 +\n    B1")).toBe("A1 + B1");
    expect(oneLine("A1\n+ B1")).toBe("A1 + B1");
  });
  it("keeps strings, spaces on one line, and the intersection operator", () => {
    expect(oneLine('CONCAT(\n    "a\nb",\n    "  c  "\n)')).toBe('CONCAT("a\nb", "  c  ")');
    expect(oneLine("SUM(A1:B5\n    B2:C3)")).toBe("SUM(A1:B5 B2:C3)");
    expect(oneLine("SUM(  A1  ,\n B1)")).toBe("SUM(  A1  , B1)");
    expect(oneLine('"x\ny"')).toBe('"x\ny"');
    expect(hasLayoutBreaks('"x\ny"')).toBe(false);
    expect(hasLayoutBreaks("LET(x,\n1, x)")).toBe(true);
  });
  it("equals the input modulo whitespace, and commutes with compile", () => {
    const forms = [
      'LET(\n    x, SEQUENCE(3),\n    s, "a\nb",\n    HSTACK(x, LEN(s))\n)',
      "LAMBDA(n, [k],\n    IF(ISOMITTED(k), n, n*k)\n)(2)",
      "XLOOKUP(\n    A1,\n    B1:B9,\n    C1:C9\n) + FILTER(D1#,\n    D1# > 0)",
    ];
    for (const src of forms) {
      const flat = oneLine(src);
      expect(hasLayoutBreaks(flat)).toBe(false);
      expect(equalModuloWhitespace(flat, src)).toBe(true);
      expect(compile(flat)).toBe(oneLine(compile(src)));
      expect(equalModuloWhitespace(compile(flat), compile(src))).toBe(true);
    }
    const long = "LET(" + Array.from({ length: 8 }, (_, i) => `val_${i}, SUM(A${i + 1}:A${i + 9})`).join(", ") + ", val_0 + val_7)";
    const pretty = prettyPrint(long, { width: 60 });
    expect(pretty).toContain("\n");
    expect(oneLine(pretty)).toBe(prettyPrint(long, { width: 100000 }));
  });
});
