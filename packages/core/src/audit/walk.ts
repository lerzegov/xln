// Questions the audit asks of one formula's syntax tree: the calls it makes (to defined
// names and to built-ins, with their arguments), how deep its functions nest, and the
// numbers written inside LAMBDA bodies.

import { children, leftSpine, type Expr, type Lambda, type Ref } from "../lang/ast.js";
import { lookupFunction, type FunctionInfo } from "../lang/catalogue.js";
import { stripPrefix } from "../lang/parser.js";
import type { Qualifier, Span } from "../lang/tokens.js";

export function unwrap(e: Expr): Expr {
  while (e.kind === "paren") e = e.expr;
  return e;
}

export type CallUse =
  | { kind: "name"; id: string; sheet: string | undefined; args: Expr[]; span: Span }
  | { kind: "builtin"; info: FunctionInfo; id: string; args: Expr[]; span: Span }
  /** A cell called like a function, `C2(…)`, `'S 1'!$C$2(…)`: the LAMBDA the cell holds. `span` is the reference's. */
  | { kind: "cell"; ref: Ref; sheet: string | undefined; args: Expr[]; span: Span };

function external(q: Qualifier | undefined): boolean {
  return q !== undefined && (q.book !== undefined || q.sheet === undefined || q.sheet2 !== undefined);
}

/**
 * Every call in `body`, in source order: to a defined name (`FN.PICK(…)`, `'S'!Fn(…)`), to
 * a built-in function, or to a cell (`C2(…)`: the LAMBDA the cell holds). LET/LAMBDA variables shadow names, as in `nameUses`; a call
 * spelled like a built-in calls the built-in (probe T12). Calls through `_xludf.` and to
 * other workbooks are not listed.
 */
export function callUses(body: Expr): CallUse[] {
  const out: CallUse[] = [];
  const visit = (node: Expr, scope: ReadonlySet<string>): void => {
    switch (node.kind) {
      case "call": {
        const fn = node.fn;
        const { prefix, base } = stripPrefix(fn.text);
        const shadowed = fn.qual === undefined && scope.has(base.toLowerCase());
        if (!external(fn.qual) && !shadowed && prefix !== "_xlpm." && prefix !== "_xludf.") {
          const info = fn.qual === undefined ? lookupFunction(base) : undefined;
          if (info) out.push({ kind: "builtin", info, id: base, args: node.args, span: fn.span });
          else if (prefix === "") out.push({ kind: "name", id: fn.text, sheet: fn.qual?.sheet, args: node.args, span: { start: fn.span.end - fn.text.length, end: fn.span.end } });
        }
        for (const a of node.args) visit(a, scope);
        return;
      }
      case "invoke": {
        const callee = unwrap(node.callee);
        if (callee.kind === "ref" && callee.refKind === "cell" && !external(callee.qual)) {
          out.push({ kind: "cell", ref: callee, sheet: callee.qual?.sheet, args: node.args, span: callee.span });
        }
        visit(node.callee, scope);
        for (const a of node.args) visit(a, scope);
        return;
      }
      case "lambda": {
        const inner = new Set(scope);
        for (const p of node.params) inner.add(stripPrefix(p.name.text).base.toLowerCase());
        visit(node.body, inner);
        return;
      }
      case "let": {
        const inner = new Set(scope);
        for (const b of node.bindings) {
          visit(b.value, new Set(inner));
          inner.add(stripPrefix(b.name.text).base.toLowerCase());
        }
        visit(node.body, inner);
        return;
      }
      case "binary": {
        const { first, links } = leftSpine(node);
        visit(first, scope);
        for (const b of links) visit(b.right, scope);
        return;
      }
      default:
        for (const c of children(node)) visit(c, scope);
    }
  };
  visit(body, new Set());
  return out;
}

/** Functions Excel writes for operators (`#`, `@`): not a nesting level the author wrote. */
const OPERATOR_FUNCTIONS = new Set(["ANCHORARRAY", "SINGLE"]);

/** How deep function calls (and LET, LAMBDA, invocations) nest. */
export function nestingDepth(body: Expr): number {
  let max = 0;
  const visit = (node: Expr, depth: number): void => {
    let d = depth;
    if (node.kind === "call") {
      if (!OPERATOR_FUNCTIONS.has(stripPrefix(node.fn.text).base.toUpperCase())) d++;
    } else if (node.kind === "lambda" || node.kind === "let" || node.kind === "invoke") d++;
    if (d > max) max = d;
    if (node.kind === "binary") {
      const { first, links } = leftSpine(node);
      visit(first, d);
      for (const b of links) visit(b.right, d);
      return;
    }
    for (const c of children(node)) visit(c, d);
  };
  visit(body, 0);
  return max;
}

/** The outermost LAMBDAs in `body` (a definition `LAMBDA(…)`, or LAMBDAs inside a formula). */
export function outerLambdas(body: Expr): Lambda[] {
  const out: Lambda[] = [];
  const visit = (node: Expr): void => {
    if (node.kind === "lambda") {
      out.push(node);
      return;
    }
    if (node.kind === "binary") {
      const { first, links } = leftSpine(node);
      visit(first);
      for (const b of links) visit(b.right);
      return;
    }
    for (const c of children(node)) visit(c);
  };
  visit(body);
  return out;
}

export interface NumberLiteral {
  value: number;
  /** As written, with a leading `-` when negated. */
  text: string;
  span: Span;
}

/** Number literals in `node`, a negation (`-1`) and a percent (`5%`) folded into the literal. */
export function numbersIn(node: Expr, src: string): NumberLiteral[] {
  const out: NumberLiteral[] = [];
  const visit = (n: Expr): void => {
    if (n.kind === "number") {
      out.push({ value: Number(n.text), text: n.text, span: n.span });
      return;
    }
    if (n.kind === "unary" && n.op === "-" && unwrap(n.operand).kind === "number") {
      const lit = unwrap(n.operand) as Extract<Expr, { kind: "number" }>;
      out.push({ value: -Number(lit.text), text: src.slice(n.span.start, n.span.end), span: n.span });
      return;
    }
    if (n.kind === "postfix" && n.op === "%" && unwrap(n.operand).kind === "number") {
      const lit = unwrap(n.operand) as Extract<Expr, { kind: "number" }>;
      out.push({ value: Number(lit.text) / 100, text: src.slice(n.span.start, n.span.end), span: n.span });
      return;
    }
    if (n.kind === "binary") {
      const { first, links } = leftSpine(n);
      visit(first);
      for (const b of links) visit(b.right);
      return;
    }
    for (const c of children(n)) visit(c);
  };
  visit(node);
  return out;
}

/** Code-point order: the same on every machine, whatever its locale. */
export function cmp(a: string, b: string): number {
  return a < b ? -1 : a > b ? 1 : 0;
}
