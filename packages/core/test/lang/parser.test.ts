import { describe, expect, it } from "vitest";
import { FormulaError, formatDiagnostic, parse, tryParse, type Expr } from "../../src/index.js";

/** Structure as an S-expression, to check precedence and node kinds. */
function sx(e: Expr, src: string): string {
  const leaf = (n: Expr) => src.slice(n.span.start, n.span.end);
  switch (e.kind) {
    case "binary":
      return `(${e.op === " " ? "isect" : e.op} ${sx(e.left, src)} ${sx(e.right, src)})`;
    case "unary":
      return `(${e.op}u ${sx(e.operand, src)})`;
    case "postfix":
      return `(${e.op}p ${sx(e.operand, src)})`;
    case "paren":
      return `[${sx(e.expr, src)}]`;
    case "call":
      return `(call ${e.fn.text} ${e.args.map((a) => sx(a, src)).join(" ")})`.replace(/ \)$/, ")");
    case "invoke":
      return `(invoke ${sx(e.callee, src)} ${e.args.map((a) => sx(a, src)).join(" ")})`;
    case "lambda":
      return `(lambda [${e.params.map((p) => (p.optional ? `?${p.name.text}` : p.name.text)).join(" ")}] ${sx(e.body, src)})`;
    case "let":
      return `(let [${e.bindings.map((b) => `${b.name.text}=${sx(b.value, src)}`).join(" ")}] ${sx(e.body, src)})`;
    case "array":
      return `{${e.rows.map((r) => r.map((x) => sx(x, src)).join(",")).join(";")}}`;
    case "missing":
      return "_";
    default:
      return leaf(e);
  }
}
const p = (src: string) => sx(parse(src).body, src);

function errorOf(src: string): { message: string; start: number } {
  try {
    parse(src);
  } catch (e) {
    expect(e).toBeInstanceOf(FormulaError);
    const d = (e as FormulaError).diagnostics[0]!;
    return { message: d.message, start: d.start };
  }
  throw new Error(`no error for ${src}`);
}

describe("parser: precedence", () => {
  it("arithmetic and comparison", () => {
    expect(p("1+2*3")).toBe("(+ 1 (* 2 3))");
    expect(p("1-2-3")).toBe("(- (- 1 2) 3)");
    expect(p("2^3^2")).toBe("(^ (^ 2 3) 2)");
    expect(p("a&b=c")).toBe("(= (& a b) c)");
    expect(p("1+2<>3*4")).toBe("(<> (+ 1 2) (* 3 4))");
  });
  it("negation binds tighter than ^, percent tighter than ^", () => {
    expect(p("-2^2")).toBe("(^ (-u 2) 2)");
    expect(p("2^50%")).toBe("(^ 2 (%p 50))");
    expect(p("-A1:A2")).toBe("(-u A1:A2)");
    expect(p("-A1:INDEX(B:B,1)")).toBe("(-u (: A1 (call INDEX B:B 1)))");
  });
  it("reference operators: range > intersection > union", () => {
    expect(p("(A1:B2 B1:C3,D4)")).toBe("[(, (isect A1:B2 B1:C3) D4)]");
    expect(p("A1:INDEX(B:B,2)")).toBe("(: A1 (call INDEX B:B 2))");
    expect(p("A1.:INDEX(B:B,2)")).toBe("(.: A1 (call INDEX B:B 2))");
  });
  it("spill and @", () => {
    expect(p("A1#*2")).toBe("(* (#p A1) 2)");
    expect(p("@A1:A10")).toBe("(@u A1:A10)");
    expect(p("@Start:Finish")).toBe("(@u (: Start Finish))");
    expect(p("x:y")).toBe("x:y"); // columns X:Y, as in Excel
    expect(p("ROWS(Revenue#)")).toBe("(call ROWS (#p Revenue))");
  });
});

describe("parser: commas", () => {
  it("separate arguments inside a call, unite inside parentheses", () => {
    expect(p("SUM(A1,B1)")).toBe("(call SUM A1 B1)");
    expect(p("SUM((A1,B1))")).toBe("(call SUM [(, A1 B1)])");
    expect(p("SUM((A1,B1),C1)")).toBe("(call SUM [(, A1 B1)] C1)");
  });
  it("unite at the top level (a name over several areas)", () => {
    expect(p("'S1'!$A$1:$B$2,'S1'!$D$1")).toBe("(, 'S1'!$A$1:$B$2 'S1'!$D$1)");
  });
  it("keeps empty arguments", () => {
    expect(p("IF(A1,,1)")).toBe("(call IF A1 _ 1)");
    expect(p("_xlfn.GROUPBY(T[K], T[N], _xleta.COUNTA, , 0)")).toBe("(call _xlfn.GROUPBY T[K] T[N] _xleta.COUNTA _ 0)");
    expect(p("NOW()")).toBe("(call NOW)");
  });
});

describe("parser: LET, LAMBDA and invocation", () => {
  it("binders become nodes", () => {
    expect(p("LET(a,1,b,a+1,a*b)")).toBe("(let [a=1 b=(+ a 1)] (* a b))");
    expect(p("_xlfn.LAMBDA(_xlpm.x, _xlpm.x+1)")).toBe("(lambda [_xlpm.x] (+ _xlpm.x 1))");
    expect(p("LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))")).toBe("(lambda [x ?y] (call IF (call ISOMITTED y) x (+ x y)))");
    expect(p("LAMBDA(42)")).toBe("(lambda [] 42)");
  });
  it("immediate invocation", () => {
    expect(p("LAMBDA(x,x)(1)")).toBe("(invoke (lambda [x] x) 1)");
    expect(p("LAMBDA(x,LAMBDA(y,x+y))(1)(2)")).toBe("(invoke (invoke (lambda [x] (lambda [y] (+ x y))) 1) 2)");
    expect(p("_xlpm.curve(_xlpm.t)")).toBe("(call _xlpm.curve _xlpm.t)");
  });
  it("a cell called like a function: an invocation of the reference", () => {
    expect(p("C2(D2,E2)")).toBe("(invoke C2 D2 E2)");
    expect(p("$C$2(1)+1")).toBe("(+ (invoke $C$2 1) 1)");
    expect(p("Sheet1!C2()")).toBe("(invoke Sheet1!C2 )");
    expect(p("'S 1'!$C$2(1,)")).toBe("(invoke 'S 1'!$C$2 1 _)");
    expect(p("C2(1)(2)")).toBe("(invoke (invoke C2 1) 2)");
    expect(p("-C2(1)%")).toBe("(%p (-u (invoke C2 1)))");
    // Built-ins spelled like cells stay calls; a space before `(` is still intersection.
    expect(p("LOG10(100)")).toBe("(call LOG10 100)");
    expect(p("ATAN2(1,2)")).toBe("(call ATAN2 1 2)");
    expect(p("DAYS360(A1,B1)")).toBe("(call DAYS360 A1 B1)");
    expect(p("T(A1)")).toBe("(call T A1)");
    expect(p("C2 (D2)")).toBe("(isect C2 [D2])");
    expect(p("A1:C2(1)")).toBe("(: A1 (invoke C2 1))");
  });
  it("rejects malformed binders with positions", () => {
    expect(errorOf("LET(a,1)").message).toMatch(/LET needs pairs/);
    expect(errorOf("LET(1,2,3)")).toMatchObject({ start: 4, message: expect.stringMatching(/must be a plain name/) });
    expect(errorOf("LET(x1,2,x1)").message).toMatch(/looks like a cell reference/);
    expect(errorOf("LAMBDA(x,x,x+1)").message).toMatch(/declared twice/);
    expect(errorOf("LAMBDA(x,)").message).toMatch(/no body/);
  });
});

describe("parser: literals", () => {
  it("array constants", () => {
    expect(p("{1,2;3,4}")).toBe("{1,2;3,4}");
    expect(p('{-1,"a";TRUE,#N/A}')).toBe('{(-u 1),"a";TRUE,#N/A}');
    expect(errorOf("{1,2;3}").message).toMatch(/rows differ in length/);
    expect(errorOf("{1,A1}").message).toMatch(/array constants may only contain/);
  });
  it("structured references and qualified names", () => {
    expect(p("SUM(Tbl[[#This Row],[A]:[B]])")).toBe("(call SUM Tbl[[#This Row],[A]:[B]])");
    expect(p("[0]!Rate*2")).toBe("(* [0]!Rate 2)");
    expect(p("'It''s'!A1+[1]Sheet1!B2")).toBe("(+ 'It''s'!A1 [1]Sheet1!B2)");
  });
  it("leading = and multi-line text", () => {
    const f = parse("=LAMBDA(x,\r\n  x*2)");
    expect(f.equals).toBe(true);
    expect(sx(f.body, f.src)).toBe("(lambda [x] (* x 2))");
  });
});

describe("parser: errors are readable and positioned (C1)", () => {
  it("unbalanced parentheses", () => {
    expect(errorOf("SUM(1,2")).toMatchObject({ start: 3, message: "this '(' is never closed" });
    expect(errorOf("SUM(1))")).toMatchObject({ start: 6, message: "unmatched ')'" });
    expect(errorOf("(1+2")).toMatchObject({ start: 4 });
  });
  it("missing operands and stray tokens", () => {
    expect(errorOf("1+")).toMatchObject({ start: 2, message: "the formula ends where an operand was expected" });
    expect(errorOf("1 2").message).toMatch(/unexpected '2'/);
    expect(errorOf("").message).toBe("empty formula");
    expect(errorOf('"abc').message).toBe("unterminated string");
  });
  it("formats with line, column and a caret", () => {
    const src = "LET(\n  a, 1,\n  a +)";
    const r = tryParse(src);
    expect(r.formula).toBeUndefined();
    expect(formatDiagnostic(src, r.diagnostics[0]!)).toBe(
      "line 3, column 6: expected an operand before ')'\n    a +)\n       ^",
    );
  });
});

describe("parser: spans", () => {
  it("cover the source of each node", () => {
    const src = "_xlfn.LET(_xlpm.a, 1,  _xlpm.a + Sheet1!B2)";
    const f = parse(src);
    expect(f.body.kind).toBe("let");
    if (f.body.kind !== "let") return;
    expect(src.slice(f.body.bindings[0]!.name.span.start, f.body.bindings[0]!.name.span.end)).toBe("_xlpm.a");
    expect(src.slice(f.body.body.span.start, f.body.body.span.end)).toBe("_xlpm.a + Sheet1!B2");
    expect(f.body.span).toEqual({ start: 0, end: src.length });
  });
});
