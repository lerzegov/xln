// Which defined names a formula reads (A6, B3). Walks the AST, so text inside strings,
// structured references and sheet names never counts, and LET/LAMBDA variables shadow
// names of the same spelling (in display form they carry no `_xlpm.` to tell them apart).

import { leftSpine, type Expr } from "../lang/ast.js";
import { lookupFunction } from "../lang/catalogue.js";
import { stripPrefix } from "../lang/parser.js";
import type { Qualifier, Span } from "../lang/tokens.js";

export interface NameUse {
  /** The identifier as written, without qualifier. */
  id: string;
  /** Sheet named in front of `!`, if any. */
  sheet: string | undefined;
  /** Where the identifier is written, qualifier and `#` excluded. */
  span: Span;
}

function external(q: Qualifier | undefined): boolean {
  return q !== undefined && (q.book !== undefined || q.sheet === undefined || q.sheet2 !== undefined);
}

/** The identifier at the end of a node's span: a qualifier (`'S 1'!`) comes before it. */
function idSpan(span: Span, id: string): Span {
  return { start: span.end - id.length, end: span.end };
}

/** Every reference to a possible defined name in `body`, in source order (duplicates kept). */
export function nameUses(body: Expr): NameUse[] {
  const out: NameUse[] = [];
  const visit = (node: Expr, scope: ReadonlySet<string>): void => {
    switch (node.kind) {
      case "name": {
        if (external(node.qual)) return;
        const { prefix, base } = stripPrefix(node.id);
        if (prefix === "_xlpm." || prefix === "_xleta.") return;
        if (prefix !== "") return; // `_xlfn.` / `_xludf.` spellings are functions, not names
        if (!node.qual && scope.has(base.toLowerCase())) return;
        out.push({ id: node.id, sheet: node.qual?.sheet, span: idSpan(node.span, node.id) });
        return;
      }
      case "call": {
        const fn = node.fn;
        const { prefix, base } = stripPrefix(fn.text);
        // A defined name holding a LAMBDA is called like a function: `FN.PICK(…)`. A call
        // spelled like a built-in calls the built-in even if a name has that spelling (T12).
        const shadowed = fn.qual === undefined && (scope.has(base.toLowerCase()) || lookupFunction(base) !== undefined);
        if (prefix === "" && !external(fn.qual) && !shadowed) {
          out.push({ id: fn.text, sheet: fn.qual?.sheet, span: idSpan(fn.span, fn.text) });
        }
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
      case "array":
        for (const r of node.rows) for (const e of r) visit(e, scope);
        return;
      case "paren":
        visit(node.expr, scope);
        return;
      case "unary":
      case "postfix":
        visit(node.operand, scope);
        return;
      case "invoke":
        visit(node.callee, scope);
        for (const a of node.args) visit(a, scope);
        return;
      default:
        return;
    }
  };
  visit(body, new Set());
  return out;
}

/**
 * Resolves names the way Excel does: `Sheet!X` is X scoped to Sheet, falling back to the
 * workbook's X; a bare X is the local X of the sheet the formula lives on (a cell's sheet,
 * or a sheet-scoped name's own sheet), falling back to the workbook's X.
 */
export class NameResolver {
  /** lower-case `sheet!name` (sheet "" for workbook scope) → key */
  private readonly byKey = new Map<string, string>();

  constructor(names: Iterable<{ name: string; scope: string | undefined; key: string }>) {
    for (const n of names) this.byKey.set(`${(n.scope ?? "").toLowerCase()}!${n.name.toLowerCase()}`, n.key);
  }

  resolve(use: NameUse, homeSheet: string | undefined): string | undefined {
    const id = use.id.toLowerCase();
    const sheet = use.sheet ?? homeSheet;
    if (sheet !== undefined) {
      const local = this.byKey.get(`${sheet.toLowerCase()}!${id}`);
      if (local !== undefined) return local;
    }
    return this.byKey.get(`!${id}`);
  }
}
