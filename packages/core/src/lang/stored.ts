// A strict check of Excel's stored grammar, run on every definition and cell formula a build
// writes. The compiler is meant to produce only stored forms Excel accepts; this check is the
// safety net for the class of bug where it does not and our own decompiler reads the wrong
// form back as if it were right (M3d: `[_xlpm.p]` read back fine, Excel dropped the name).
// It reads the stored text on its own terms and never by round-tripping it through
// decompile.

import { leftSpine, type Expr } from "./ast.js";
import { lookupFunction } from "./catalogue.js";
import { FormulaError, type Diagnostic } from "./errors.js";
import { parse, stripPrefix } from "./parser.js";
import type { Span } from "./tokens.js";

export interface StoredCheckOptions {
  /** Every defined name in the workbook: a call to one of them is not an unknown function. */
  names?: Iterable<string>;
}

function problem(code: string, span: Span, message: string): Diagnostic {
  return { severity: "error", code, message, start: span.start, end: span.end };
}

/**
 * The ways `stored` departs from what Excel writes: brackets in a LAMBDA parameter list,
 * `_xlpm.`/`_xlop.` outside an enclosing LAMBDA or LET that binds them, a parameter used
 * without `_xlpm.`, display-only syntax (`x#`, prefix `@`, `[@Col]`), and functions missing
 * from the catalogue or stored without the prefix it gives. Empty when the text is clean.
 */
export function checkStoredForm(stored: string, opts: StoredCheckOptions = {}): Diagnostic[] {
  let f;
  try {
    f = parse(stored);
  } catch (e) {
    if (e instanceof FormulaError) return e.diagnostics;
    throw e;
  }
  const src = stored;
  const names = new Set<string>();
  for (const n of opts.names ?? []) names.add(n.toLowerCase());
  const out: Diagnostic[] = [];

  /** A function as stored: its prefix must be the catalogue's. */
  const fnPrefix = (fn: { text: string; span: Span }, what: string): void => {
    const { prefix, base } = stripPrefix(fn.text);
    const info = lookupFunction(base);
    if (!info) {
      out.push(problem("unknown-function", fn.span, `${what} '${fn.text}': '${base}' is not in the function catalogue`));
      return;
    }
    if (prefix !== info.prefix && !info.prefixUnsure) {
      out.push(problem("wrong-prefix", fn.span, `${what} '${fn.text}' must be stored as '${info.prefix}${info.name}'`));
    }
  };

  /** A `_xlpm.` reference must name a variable in scope; a bare one must not. */
  const variable = (id: string, span: Span, scope: ReadonlySet<string>): boolean => {
    const { prefix, base } = stripPrefix(id);
    const key = base.toLowerCase();
    if (prefix === "_xlop.") {
      out.push(problem("optional-use", span, `'${id}': _xlop. marks an optional parameter in a LAMBDA's parameter list only; a use is stored as '_xlpm.${base}'`));
      return true;
    }
    if (prefix === "_xlpm.") {
      if (!scope.has(key)) out.push(problem("unbound-parameter", span, `'${id}' is not bound by an enclosing LAMBDA or LET`));
      return true;
    }
    if (prefix === "" && scope.has(key)) {
      out.push(problem("unprefixed-parameter", span, `'${id}' is a LAMBDA or LET variable here and must be stored as '_xlpm.${base}' (Excel would read a defined name)`));
      return true;
    }
    return false;
  };

  const visit = (node: Expr, scope: ReadonlySet<string>): void => {
    switch (node.kind) {
      case "name": {
        if (node.qual) return;
        if (variable(node.id, node.span, scope)) return;
        const { prefix, base } = stripPrefix(node.id);
        // `_xleta.SUM`: a function passed by name (probe F9); it has no prefix of its own.
        if (prefix === "_xleta." && !lookupFunction(base)) {
          out.push(problem("unknown-function", node.span, `'${node.id}': '${base}' is passed as a function but is not in the function catalogue`));
        }
        return;
      }
      case "structref":
        if (node.inner.trimStart().startsWith("@")) {
          out.push(problem("display-syntax", node.span, `'${src.slice(node.span.start, node.span.end)}' is display form; Excel stores [#This Row],[…]`));
        }
        return;
      case "postfix":
        if (node.op === "#") out.push(problem("display-syntax", node.span, `the spill operator '#' is display form; Excel stores _xlfn.ANCHORARRAY(…)`));
        visit(node.operand, scope);
        return;
      case "unary":
        if (node.op === "@") out.push(problem("display-syntax", node.span, `'@' is display form; Excel stores _xlfn.SINGLE(…)`));
        visit(node.operand, scope);
        return;
      case "lambda": {
        fnPrefix(node.fn, "the function");
        const inner = new Set(scope);
        for (const p of node.params) {
          const { prefix, base } = stripPrefix(p.name.text);
          const written = src.slice(p.span.start, p.span.end);
          if (p.span.start !== p.name.span.start) {
            out.push(problem("bracketed-parameter", p.span, `'${written}': an optional LAMBDA parameter is stored as '_xlop.${base}', without brackets (Excel drops the name otherwise)`));
          } else if (prefix !== "_xlpm." && prefix !== "_xlop.") {
            out.push(problem("unprefixed-parameter", p.span, `LAMBDA parameter '${written}' must be stored as '_xlpm.${base}'`));
          }
          inner.add(base.toLowerCase());
        }
        visit(node.body, inner);
        return;
      }
      case "let": {
        fnPrefix(node.fn, "the function");
        const inner = new Set(scope);
        for (const b of node.bindings) {
          const { prefix, base } = stripPrefix(b.name.text);
          if (prefix !== "_xlpm.") out.push(problem("unprefixed-parameter", b.name.span, `LET name '${b.name.text}' must be stored as '_xlpm.${base}'`));
          // A binding's own value cannot see the binding.
          visit(b.value, new Set(inner));
          inner.add(base.toLowerCase());
        }
        visit(node.body, inner);
        return;
      }
      case "call": {
        const fn = node.fn;
        if (!fn.qual && !variable(fn.text, fn.span, scope)) {
          const { prefix, base } = stripPrefix(fn.text);
          // `_xludf.` is Excel's own spelling of a call it did not recognise (probe F6).
          if (prefix !== "_xludf." && !(prefix === "" && names.has(base.toLowerCase()) && !lookupFunction(base))) {
            fnPrefix(fn, "the function");
          }
        }
        for (const a of node.args) visit(a, scope);
        return;
      }
      case "binary": {
        const { first, links } = leftSpine(node);
        visit(first, scope);
        for (const b of links) visit(b.right, scope);
        return;
      }
      case "paren":
        visit(node.expr, scope);
        return;
      case "invoke":
        visit(node.callee, scope);
        for (const a of node.args) visit(a, scope);
        return;
      case "array":
        for (const r of node.rows) for (const e of r) visit(e, scope);
        return;
      default:
        return;
    }
  };
  visit(f.body, new Set());
  return out;
}
