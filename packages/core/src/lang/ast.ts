// AST for Excel formulas. Every node carries the span of source text it was parsed from,
// so transformations can rewrite a node and keep the author's spacing everywhere else.

import type { Qualifier, Span } from "./tokens.js";

interface Base {
  span: Span;
}

export interface NumberLit extends Base {
  kind: "number";
  text: string;
}
export interface StringLit extends Base {
  kind: "string";
  /** Decoded value (`""` already turned into `"`). */
  value: string;
}
export interface BoolLit extends Base {
  kind: "bool";
  value: boolean;
}
export interface ErrorLit extends Base {
  kind: "error";
  text: string;
}
/** `{1,2;3,4}`: rows of constant elements. */
export interface ArrayLit extends Base {
  kind: "array";
  rows: Expr[][];
}
/** A cell, area, whole columns/rows, or a qualified `#REF!`. */
export interface Ref extends Base {
  kind: "ref";
  qual?: Qualifier;
  /** `$A$1`, `A1:B2`, `A1:.B100`, `A:A`, `1:3`, `#REF!` */
  address: string;
  refKind: "cell" | "area" | "cols" | "rows" | "error";
}
/** A defined name or a LET/LAMBDA variable, as written (may carry `_xlpm.` or `_xleta.`; `_xlop.` only at a LAMBDA's parameter). */
export interface Name extends Base {
  kind: "name";
  qual?: Qualifier;
  id: string;
}
/** A structured reference. `table` is "" inside a Table (`[@Col]`). */
export interface StructRef extends Base {
  kind: "structref";
  table: string;
  /** Text between the outer brackets, as written. */
  inner: string;
}
/** A function name or binder at a declaration site. Leaf; `text` as written. */
export interface Ident extends Base {
  kind: "ident";
  text: string;
  qual?: Qualifier;
}
export interface Missing extends Base {
  kind: "missing";
}
export interface Paren extends Base {
  kind: "paren";
  expr: Expr;
}
export interface Unary extends Base {
  kind: "unary";
  /** `-`, `+`, or `@` (implicit intersection). */
  op: string;
  operand: Expr;
}
export interface Postfix extends Base {
  kind: "postfix";
  /** `%`, or `#` (spill range). */
  op: string;
  operand: Expr;
}
export interface Binary extends Base {
  kind: "binary";
  /** Arithmetic, comparison, `&`, range `:` (and trim `:.` `.:` `.:.`), ` ` intersection, `,` union. */
  op: string;
  left: Expr;
  right: Expr;
}
/** `NAME(args)`. `fn.text` may carry `_xlfn.`, `_xlfn._xlws.`, `_xlpm.`, `_xludf.`. */
export interface Call extends Base {
  kind: "call";
  fn: Ident;
  args: Expr[];
}
/** A LAMBDA parameter. Optional ones are written `[y]` in display form and `_xlop.y` in
 *  stored form (measured, probe F9); `span` covers the parameter as written, brackets included. */
export interface LambdaParam {
  name: Ident;
  optional: boolean;
  span: Span;
}
/** `LAMBDA(x, [y], body)`. */
export interface Lambda extends Base {
  kind: "lambda";
  fn: Ident;
  params: LambdaParam[];
  body: Expr;
}
/** `LET(a, 1, b, 2, body)`. */
export interface Let extends Base {
  kind: "let";
  fn: Ident;
  bindings: { name: Ident; value: Expr }[];
  body: Expr;
}
/** Calling the result of an expression: `LAMBDA(x, x+1)(2)`. */
export interface Invoke extends Base {
  kind: "invoke";
  callee: Expr;
  args: Expr[];
}

export type Expr =
  | NumberLit
  | StringLit
  | BoolLit
  | ErrorLit
  | ArrayLit
  | Ref
  | Name
  | StructRef
  | Ident
  | Missing
  | Paren
  | Unary
  | Postfix
  | Binary
  | Call
  | Lambda
  | Let
  | Invoke;

/** A whole formula: optional leading `=`, then one expression. */
export interface Formula {
  kind: "formula";
  span: Span;
  equals: boolean;
  body: Expr;
  /** The source text the spans index into. */
  src: string;
}

/** Child nodes in source order. Missing arguments are skipped. */
export function children(node: Expr): Expr[] {
  switch (node.kind) {
    case "array":
      return node.rows.flat();
    case "paren":
      return [node.expr];
    case "unary":
    case "postfix":
      return [node.operand];
    case "binary":
      return [node.left, node.right];
    case "call":
      return [node.fn, ...node.args.filter((a) => a.kind !== "missing")];
    case "lambda":
      return [node.fn, ...node.params.map((p) => p.name), node.body];
    case "let":
      return [node.fn, ...node.bindings.flatMap((b) => [b.name, b.value]), node.body];
    case "invoke":
      return [node.callee, ...node.args.filter((a) => a.kind !== "missing")];
    default:
      return [];
  }
}

/** Visit every node, parents first. */
export function walk(node: Expr, visit: (n: Expr) => void): void {
  visit(node);
  if (node.kind === "binary") {
    const { first, links } = leftSpine(node);
    for (const b of links.slice(0, -1).reverse()) visit(b);
    walk(first, visit);
    for (const b of links) walk(b.right, visit);
    return;
  }
  for (const c of children(node)) walk(c, visit);
}

/**
 * A rewrite: return the new text for a node, or undefined to keep it and recurse.
 * `inner(child)` renders a child with the same rewrite applied.
 */
export type Rewrite = (node: Expr, inner: (n: Expr) => string) => string | undefined;

/** Render a node back to text, applying `rw`; untouched text (spacing, line breaks) is kept. */
export function render(src: string, node: Expr, rw: Rewrite): string {
  const inner = (n: Expr): string => render(src, n, rw);
  const r = rw(node, inner);
  if (r !== undefined) return r;
  if (node.kind === "binary") return renderChain(src, node, rw, inner);
  let out = "";
  let pos = node.span.start;
  for (const c of children(node)) {
    out += src.slice(pos, c.span.start) + inner(c);
    pos = c.span.end;
  }
  return out + src.slice(pos, node.span.end);
}

/**
 * Left-deep operator chains (`x+1+1+…`) can be thousands of levels deep in a definition
 * of 8,000 characters; walk their left spine in a loop instead of recursing.
 */
function renderChain(src: string, node: Binary, rw: Rewrite, inner: (n: Expr) => string): string {
  const spine: Binary[] = [node];
  let cur: Expr = node.left;
  let head: string | undefined;
  while (cur.kind === "binary") {
    head = rw(cur, inner);
    if (head !== undefined) break;
    spine.push(cur);
    cur = cur.left;
  }
  let out = head ?? inner(cur);
  for (let k = spine.length - 1; k >= 0; k--) {
    const b = spine[k]!;
    out += src.slice(b.left.span.end, b.right.span.start) + inner(b.right);
  }
  return out;
}

/** The operands of a left-deep chain and the binary nodes joining them, outermost last. */
export function leftSpine(node: Binary): { first: Expr; links: Binary[] } {
  const links: Binary[] = [];
  let cur: Expr = node;
  while (cur.kind === "binary") {
    links.push(cur);
    cur = cur.left;
  }
  return { first: cur, links: links.reverse() };
}

/** The source text of a node. */
export function textOf(src: string, node: { span: Span }): string {
  return src.slice(node.span.start, node.span.end);
}
