// Hand-written recursive-descent parser for Excel formulas (FEASIBILITY §8.5).
//
// Precedence, loosest first, as Microsoft documents it:
//   comparison  =  <>  <  >  <=  >=
//   &
//   +  -
//   *  /
//   ^
//   %            (postfix)
//   -  +  @      (prefix)
//   ,            union, only inside parentheses or at the top level
//   ␠            intersection
//   :            range (and the trim forms :.  .:  .:.)
//   #            spill (postfix); call and invocation bind here too, and so does a
//                cell called like a function (`C2(x, y)`, an invocation of the cell)

import type {
  ArrayLit,
  Expr,
  Formula,
  Ident,
  Lambda,
  Let,
} from "./ast.js";
import { FormulaError, type Diagnostic } from "./errors.js";
import { tokenize, type Token } from "./tokens.js";

const COMPARISON = new Set(["=", "<>", "<", ">", "<=", ">="]);
const RANGE_OPS = new Set([":", ":.", ".:", ".:."]);

/** Strips `_xlfn.`, `_xlfn._xlws.`, `_xlpm.`, `_xlop.`, `_xleta.`, `_xludf.` from an identifier. */
export function stripPrefix(id: string): { prefix: string; base: string } {
  for (const p of ["_xlfn._xlws.", "_xlfn.", "_xlpm.", "_xlop.", "_xleta.", "_xludf."]) {
    if (id.length > p.length && id.slice(0, p.length).toLowerCase() === p) {
      return { prefix: p, base: id.slice(p.length) };
    }
  }
  return { prefix: "", base: id };
}

class Parser {
  private readonly toks: Token[];
  private pos = 0;
  /** Whether a comma at the current nesting is the union operator (true) or an argument separator. */
  private readonly unionOk: boolean[] = [true];

  constructor(private readonly src: string) {
    this.toks = tokenize(src).filter((t) => t.kind !== "ws");
  }

  private peek(k = 0): Token {
    return this.toks[Math.min(this.pos + k, this.toks.length - 1)]!;
  }
  private next(): Token {
    const t = this.peek();
    if (this.pos < this.toks.length - 1) this.pos++;
    return t;
  }
  private isOp(t: Token, ...ops: string[]): boolean {
    return t.kind === "op" && ops.includes(t.text);
  }

  fail(at: { start: number; end: number }, message: string): never {
    const d: Diagnostic = { severity: "error", code: "syntax", message, start: at.start, end: at.end };
    throw new FormulaError(this.src, [d]);
  }

  private describe(t: Token): string {
    if (t.kind === "eof") return "the end of the formula";
    if (t.kind === "invalid") return t.message ?? `'${t.text}'`;
    return `'${t.text}'`;
  }

  private expect(kind: Token["kind"], what: string): Token {
    const t = this.peek();
    if (t.kind !== kind) this.fail(t, `expected ${what}, found ${this.describe(t)}`);
    return this.next();
  }

  parseFormula(): Formula {
    let equals = false;
    const first = this.peek();
    if (this.isOp(first, "=")) {
      equals = true;
      this.next();
    }
    if (this.peek().kind === "eof") this.fail(this.peek(), "empty formula");
    const body = this.parseExpr();
    const t = this.peek();
    if (t.kind !== "eof") {
      if (t.kind === ")") this.fail(t, "unmatched ')'");
      if (t.kind === "isect") this.fail(t, "unexpected space");
      this.fail(t, `unexpected ${this.describe(t)} after a complete expression`);
    }
    return { kind: "formula", span: { start: 0, end: this.src.length }, equals, body, src: this.src };
  }

  private parseExpr(): Expr {
    return this.parseComparison();
  }

  private binaryLevel(ops: (t: Token) => boolean, sub: () => Expr): Expr {
    let left = sub();
    while (ops(this.peek())) {
      const op = this.next();
      const right = sub();
      left = { kind: "binary", op: op.text, left, right, span: { start: left.span.start, end: right.span.end } };
    }
    return left;
  }

  private parseComparison(): Expr {
    return this.binaryLevel((t) => t.kind === "op" && COMPARISON.has(t.text), () => this.parseConcat());
  }
  private parseConcat(): Expr {
    return this.binaryLevel((t) => this.isOp(t, "&"), () => this.parseAdditive());
  }
  private parseAdditive(): Expr {
    return this.binaryLevel((t) => this.isOp(t, "+", "-"), () => this.parseMultiplicative());
  }
  private parseMultiplicative(): Expr {
    return this.binaryLevel((t) => this.isOp(t, "*", "/"), () => this.parsePower());
  }
  private parsePower(): Expr {
    return this.binaryLevel((t) => this.isOp(t, "^"), () => this.parsePercent());
  }
  private parsePercent(): Expr {
    let e = this.parseUnary();
    while (this.isOp(this.peek(), "%")) {
      const t = this.next();
      e = { kind: "postfix", op: "%", operand: e, span: { start: e.span.start, end: t.end } };
    }
    return e;
  }
  private parseUnary(): Expr {
    const t = this.peek();
    if (this.isOp(t, "-", "+", "@")) {
      this.next();
      const operand = this.parseUnary();
      return { kind: "unary", op: t.text, operand, span: { start: t.start, end: operand.span.end } };
    }
    return this.parseUnion();
  }
  private parseUnion(): Expr {
    return this.binaryLevel(
      (t) => t.kind === "," && this.unionOk[this.unionOk.length - 1] === true,
      () => this.parseIntersection(),
    );
  }
  private parseIntersection(): Expr {
    let left = this.parseRange();
    while (this.peek().kind === "isect") {
      this.next();
      const right = this.parseRange();
      left = { kind: "binary", op: " ", left, right, span: { start: left.span.start, end: right.span.end } };
    }
    return left;
  }
  private parseRange(): Expr {
    return this.binaryLevel((t) => t.kind === "op" && RANGE_OPS.has(t.text), () => this.parsePostfix());
  }

  /** Primary followed by `#` spills and immediate invocations `(…)(…)`, `C2(…)`. */
  private parsePostfix(): Expr {
    let e = this.parsePrimary();
    for (;;) {
      const t = this.peek();
      if (this.isOp(t, "#")) {
        this.next();
        e = { kind: "postfix", op: "#", operand: e, span: { start: e.span.start, end: t.end } };
        continue;
      }
      if (
        t.kind === "(" &&
        t.start === e.span.end &&
        (e.kind === "call" || e.kind === "lambda" || e.kind === "let" || e.kind === "invoke" || e.kind === "paren" || (e.kind === "ref" && e.refKind === "cell"))
      ) {
        // A cell called like a function, `C2(x, y)`: the cell holds a LAMBDA (see the tokenizer).
        const { args, end } = this.parseArgs();
        e = { kind: "invoke", callee: e, args, span: { start: e.span.start, end } };
        continue;
      }
      return e;
    }
  }

  private parsePrimary(): Expr {
    const t = this.peek();
    const span = { start: t.start, end: t.end };
    switch (t.kind) {
      case "number":
        this.next();
        return { kind: "number", text: t.text, span };
      case "string":
        this.next();
        return { kind: "string", value: t.value ?? "", span };
      case "bool":
        this.next();
        return { kind: "bool", value: t.value === "TRUE", span };
      case "error":
        this.next();
        return { kind: "error", text: t.text, span };
      case "ref":
        this.next();
        return { kind: "ref", qual: t.qual, address: t.value ?? t.text, refKind: t.refKind ?? "cell", span };
      case "structref":
        this.next();
        return { kind: "structref", table: t.value ?? "", inner: t.inner ?? "", span };
      case "name": {
        this.next();
        const nx = this.peek();
        if (nx.kind === "(" && nx.start === t.end) return this.parseCall(t);
        return { kind: "name", qual: t.qual, id: t.value ?? t.text, span };
      }
      case "(": {
        this.next();
        this.unionOk.push(true);
        const expr = this.parseExpr();
        this.unionOk.pop();
        const close = this.peek();
        if (close.kind !== ")") this.fail(close, `expected ')' to close the '(' at column ${t.start + 1}, found ${this.describe(close)}`);
        this.next();
        return { kind: "paren", expr, span: { start: t.start, end: close.end } };
      }
      case "{":
        return this.parseArray();
      case "invalid":
        this.fail(t, t.message ?? `unexpected '${t.text}'`);
        break;
      case "eof":
        this.fail(t, "the formula ends where an operand was expected");
        break;
      case ")":
        this.fail(t, "expected an operand before ')'");
        break;
      case ",":
        this.fail(t, "expected an operand before ','");
        break;
      default:
        break;
    }
    return this.fail(t, `expected an operand, found ${this.describe(t)}`);
  }

  /** `(a, , b)`: arguments, with empty ones as `missing`. Commas here separate, never unite. */
  private parseArgs(): { args: Expr[]; end: number } {
    const open = this.expect("(", "'('");
    const args: Expr[] = [];
    this.unionOk.push(false);
    if (this.peek().kind === ")") {
      this.unionOk.pop();
      return { args, end: this.next().end };
    }
    for (;;) {
      const t = this.peek();
      if (t.kind === "," || t.kind === ")") {
        args.push({ kind: "missing", span: { start: t.start, end: t.start } });
      } else {
        args.push(this.parseExpr());
      }
      const sep = this.peek();
      if (sep.kind === ",") {
        this.next();
        continue;
      }
      if (sep.kind === ")") {
        this.next();
        this.unionOk.pop();
        return { args, end: sep.end };
      }
      if (sep.kind === "eof") this.fail({ start: open.start, end: open.end }, "this '(' is never closed");
      this.fail(sep, `expected ',' or ')' in the argument list, found ${this.describe(sep)}`);
    }
  }

  private parseCall(nameTok: Token): Expr {
    const fn: Ident = { kind: "ident", text: nameTok.value ?? nameTok.text, qual: nameTok.qual, span: { start: nameTok.start, end: nameTok.end } };
    const { args, end } = this.parseArgs();
    const span = { start: nameTok.start, end };
    const base = stripPrefix(fn.text);
    const upper = base.base.toUpperCase();
    if (!fn.qual && (base.prefix === "" || base.prefix === "_xlfn.")) {
      if (upper === "LAMBDA") return this.toLambda(fn, args, span);
      if (upper === "LET") return this.toLet(fn, args, span);
    }
    return { kind: "call", fn, args, span };
  }

  /** A LET/LAMBDA variable at its declaration: a bare, unqualified name. */
  private binder(arg: Expr, what: string): Ident {
    if (arg.kind === "name" && !arg.qual) return { kind: "ident", text: arg.id, span: arg.span };
    if (arg.kind === "ref") this.fail(arg.span, `${what} '${this.src.slice(arg.span.start, arg.span.end)}' looks like a cell reference; choose another name`);
    if (arg.kind === "missing") this.fail(arg.span, `${what} is missing`);
    return this.fail(arg.span, `${what} must be a plain name`);
  }

  private toLambda(fn: Ident, args: Expr[], span: { start: number; end: number }): Lambda {
    if (args.length === 0) this.fail(span, "LAMBDA needs at least a body");
    const body = args[args.length - 1]!;
    if (body.kind === "missing") this.fail({ start: span.end - 1, end: span.end }, "LAMBDA has no body (its last argument is empty)");
    const params = args.slice(0, -1).map((a, k) => {
      // Optional parameter: `[x]` in display form.
      if (a.kind === "structref" && a.table === "" && isPlainIdent(a.inner)) {
        const name: Ident = { kind: "ident", text: a.inner, span: { start: a.span.start + 1, end: a.span.end - 1 } };
        return { name, optional: true, span: a.span };
      }
      // `_xlop.x` in stored form: Excel's spelling of `[x]` (probe F9).
      const name = this.binder(a, `LAMBDA parameter ${k + 1}`);
      return { name, optional: stripPrefix(name.text).prefix === "_xlop.", span: name.span };
    });
    const seen = new Set<string>();
    for (const p of params) {
      const key = stripPrefix(p.name.text).base.toLowerCase();
      if (seen.has(key)) this.fail(p.name.span, `LAMBDA parameter '${p.name.text}' is declared twice`);
      seen.add(key);
    }
    return { kind: "lambda", fn, params, body, span };
  }

  private toLet(fn: Ident, args: Expr[], span: { start: number; end: number }): Let {
    if (args.length < 3 || args.length % 2 === 0) {
      this.fail(span, `LET needs pairs of name and value followed by a result; it has ${args.length} argument${args.length === 1 ? "" : "s"}`);
    }
    const bindings: Let["bindings"] = [];
    for (let k = 0; k + 1 < args.length; k += 2) {
      const value = args[k + 1]!;
      if (value.kind === "missing") this.fail(value.span, `LET value for argument ${k + 2} is missing`);
      bindings.push({ name: this.binder(args[k]!, `LET name (argument ${k + 1})`), value });
    }
    const body = args[args.length - 1]!;
    if (body.kind === "missing") this.fail({ start: span.end - 1, end: span.end }, "LET has no result (its last argument is empty)");
    return { kind: "let", fn, bindings, body, span };
  }

  private parseArray(): ArrayLit {
    const open = this.next();
    const rows: Expr[][] = [[]];
    for (;;) {
      rows[rows.length - 1]!.push(this.parseArrayElement());
      const t = this.peek();
      if (t.kind === ",") {
        this.next();
        continue;
      }
      if (t.kind === ";") {
        this.next();
        rows.push([]);
        continue;
      }
      if (t.kind === "}") {
        this.next();
        const width = rows[0]!.length;
        for (const r of rows) {
          if (r.length !== width) {
            const at = r[0]!.span;
            this.fail({ start: at.start, end: r[r.length - 1]!.span.end }, `array constant rows differ in length: ${r.length} instead of ${width}`);
          }
        }
        return { kind: "array", rows, span: { start: open.start, end: t.end } };
      }
      if (t.kind === "eof") this.fail(open, "this '{' is never closed");
      this.fail(t, `expected ',' ';' or '}' in an array constant, found ${this.describe(t)}`);
    }
  }

  private parseArrayElement(): Expr {
    const t = this.peek();
    if (this.isOp(t, "-", "+")) {
      this.next();
      const n = this.peek();
      if (n.kind !== "number") this.fail(n, "array constants may only contain numbers, text, TRUE/FALSE and errors");
      this.next();
      const operand: Expr = { kind: "number", text: n.text, span: { start: n.start, end: n.end } };
      return { kind: "unary", op: t.text, operand, span: { start: t.start, end: n.end } };
    }
    if (t.kind === "number" || t.kind === "string" || t.kind === "bool" || t.kind === "error") return this.parsePrimary();
    return this.fail(t, `array constants may only contain numbers, text, TRUE/FALSE and errors; found ${this.describe(t)}`);
  }
}

function isPlainIdent(s: string): boolean {
  if (s.length === 0) return false;
  const toks = tokenize(s);
  return toks.length === 2 && toks[0]!.kind === "name" && !toks[0]!.qual && toks[0]!.text === s;
}

/** Parse a formula (stored or display form, with or without a leading `=`).
 *  Throws `FormulaError` with a positioned diagnostic on bad syntax. */
export function parse(src: string): Formula {
  return new Parser(src).parseFormula();
}

/** Like `parse`, but returns diagnostics instead of throwing. */
export function tryParse(src: string): { formula?: Formula; diagnostics: Diagnostic[] } {
  try {
    return { formula: parse(src), diagnostics: [] };
  } catch (e) {
    if (e instanceof FormulaError) return { diagnostics: e.diagnostics };
    throw e;
  }
}

