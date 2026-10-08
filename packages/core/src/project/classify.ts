// A3: classify a defined name by the shape of its definition's syntax tree.
// The decision is syntactic on purpose: a name over a dynamic-array cell that did not
// spill (saved 1×1 with `cm`) is a range, and a `x#` name is a spill whatever extent the
// file last saved; the manifest carries the extents.

import type { Expr, Lambda } from "../lang/ast.js";
import { lookupFunction } from "../lang/catalogue.js";
import { stripPrefix, tryParse } from "../lang/parser.js";
import { parseCell } from "../file/cellref.js";
import type { Classification } from "./types.js";

function unwrap(e: Expr): Expr {
  while (e.kind === "paren") e = e.expr;
  return e;
}

function isConstant(e: Expr): boolean {
  e = unwrap(e);
  switch (e.kind) {
    case "number":
    case "string":
    case "bool":
    case "error":
      return true;
    case "array":
      return e.rows.every((r) => r.every(isConstant));
    case "unary":
      return (e.op === "-" || e.op === "+") && isConstant(e.operand);
    case "postfix":
      return e.op === "%" && isConstant(e.operand);
    default:
      return false;
  }
}

const RANGE_OPS = new Set([":", " ", ","]);

/** A fixed reference: cells, areas, unions and intersections of them. */
function isRange(e: Expr): boolean {
  e = unwrap(e);
  if (e.kind === "ref") return true;
  if (e.kind === "binary" && RANGE_OPS.has(e.op)) return isRange(e.left) && isRange(e.right);
  return false;
}

function anchorOf(e: Expr): Classification["anchor"] {
  e = unwrap(e);
  if (e.kind === "ref" && e.refKind === "cell") return { sheet: e.qual?.sheet, cell: e.address.split("$").join("") };
  return undefined;
}

/** The operand of `x#` or `ANCHORARRAY(x)`, or undefined. */
function spillOperand(e: Expr): Expr | undefined {
  e = unwrap(e);
  if (e.kind === "postfix" && e.op === "#") return e.operand;
  if (e.kind === "call" && !e.fn.qual && e.args.length === 1) {
    const { base } = stripPrefix(e.fn.text);
    if (base.toUpperCase() === "ANCHORARRAY") return e.args[0];
  }
  return undefined;
}

/**
 * The cells a definition denotes when it is exactly one reference: a cell or an area
 * (`'BS'!$C$6`, `'BS'!$C$6:$G$6`, with or without `$`), or the spill of a cell
 * (`'BS'!$C$6#`, stored `ANCHORARRAY('BS'!$C$6)`). Undefined for anything else (unions,
 * whole rows or columns, trim ranges, 3-D or external references, formulas).
 */
export function definitionTarget(definition: string): DefinitionTarget | undefined {
  const { formula } = tryParse(definition);
  if (!formula) return undefined;
  let e = unwrap(formula.body);
  const spilled = spillOperand(e);
  if (spilled) e = unwrap(spilled);
  if (e.kind !== "ref" || (e.refKind !== "cell" && e.refKind !== "area")) return undefined;
  const q = e.qual;
  if (q && (q.book !== undefined || q.sheet === undefined || q.sheet2 !== undefined)) return undefined;
  if (spilled && e.refKind !== "cell") return undefined;
  const [a, b = a, extra] = e.address.split("$").join("").split(":");
  if (extra !== undefined || a!.includes(".") || b!.includes(".")) return undefined;
  const p = parseCell(a!);
  const r = parseCell(b!);
  if (!p || !r) return undefined;
  return {
    sheet: q?.sheet,
    spill: spilled !== undefined,
    absolute: e.address.split(":").every(isAbsoluteCell),
    r1: Math.min(p.row, r.row),
    c1: Math.min(p.col, r.col),
    r2: Math.max(p.row, r.row),
    c2: Math.max(p.col, r.col),
  };
}

/** `$B$12`: `$`, the column, `$`, the row. */
function isAbsoluteCell(a: string): boolean {
  const parts = a.split("$");
  return parts.length === 3 && parts[0] === "" && parts[1] !== "" && parts[2] !== "";
}

/**
 * The sheet whose cells a definition is fixed to (sheet-cell names, 2026-10-07): one
 * absolute cell, area or spill of one sheet, written with its sheet (`Mortgage!$B$12`,
 * `Mortgage!$B$12:$G$12`, `Mortgage!$A$12#`). `sheets` are the workbook's worksheets; the
 * sheet comes back spelled as there (as written when `sheets` is unknown). Undefined for
 * anything else: relative references, unions, whole rows or columns, 3-D or external
 * references, formulas, constants.
 */
export function sheetCellsOf(definition: string, sheets: Iterable<string> | undefined): string | undefined {
  const t = definitionTarget(definition);
  if (!t || !t.absolute || t.sheet === undefined) return undefined;
  if (sheets === undefined) return t.sheet;
  const want = t.sheet.toLowerCase();
  for (const s of sheets) if (s.toLowerCase() === want) return s;
  return undefined;
}

export interface DefinitionTarget {
  /** The sheet written in front of `!`; undefined when the reference has none. */
  sheet: string | undefined;
  /** `x#`: the spill of the cell `r1`, `c1`. */
  spill: boolean;
  /** Every row and column written with `$` (`$B$12:$G$12`): the cells stay put wherever it is read. */
  absolute: boolean;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** Classify a definition, stored or display form. */
export function classify(definition: string): Classification {
  const { formula, diagnostics } = tryParse(definition);
  if (!formula) return { kind: "unparsed", error: diagnostics[0]?.message ?? "does not parse" };
  const body = unwrap(formula.body);
  if (body.kind === "lambda") return lambdaClass(body);
  if (isConstant(body)) return { kind: "constant" };
  const spilled = spillOperand(body);
  if (spilled) {
    const anchor = anchorOf(spilled);
    return anchor ? { kind: "spill", anchor } : { kind: "spill" };
  }
  if (isRange(body)) return { kind: "range" };
  if (body.kind === "structref" && body.table !== "") return { kind: "table", table: body.table };
  return { kind: "formula" };
}

function lambdaClass(l: Lambda): Classification & { kind: "lambda" } {
  const optional = l.params.filter((p) => p.optional).length;
  return {
    kind: "lambda",
    arity: { required: l.params.length - optional, optional },
    params: l.params.map((p) => {
      const n = stripPrefix(p.name.text).base;
      return p.optional ? `[${n}]` : n;
    }),
  };
}

/**
 * What a cell gives when it is called like a function (`C2(x, y)`, or `Fn(x, y)` with `Fn`
 * a name on C2): Excel calls the LAMBDA the cell holds, and gives #VALUE! or #CALC! when it
 * holds anything else.
 */
export type CellCallee =
  /** The cell's formula is a LAMBDA (possibly as a LET's result). */
  | (Classification & { kind: "lambda" })
  /** Certainly not a LAMBDA: a value or nothing, or a formula whose result is a number, text, … */
  | { kind: "not-lambda"; holds: "value" | "formula" }
  /** A formula that may give a LAMBDA (a name, a reference, IF, INDEX, a LAMBDA's call…). */
  | { kind: "unknown" };

/**
 * Built-ins that can give back one of their arguments or a reference to a cell, and so a
 * cell's LAMBDA (`IF(x, C2, C3)`, `INDEX(C2:C3, k)`): calling their result may be fine.
 */
const PASS_THROUGH = new Set(["IF", "IFS", "IFERROR", "IFNA", "CHOOSE", "SWITCH", "INDEX", "INDIRECT", "OFFSET", "XLOOKUP", "LOOKUP", "VLOOKUP", "HLOOKUP", "SINGLE", "REDUCE"]);
const VALUE_OPS = new Set(["+", "-", "*", "/", "^", "&", "=", "<>", "<", ">", "<=", ">="]);

/** `formula`: the cell's formula (stored or display form); undefined or "" when it has none. */
export function cellCallee(formula: string | undefined): CellCallee {
  if (formula === undefined || formula.trim() === "") return { kind: "not-lambda", holds: "value" };
  const { formula: f } = tryParse(formula);
  if (!f) return { kind: "unknown" };
  const notLambda = { kind: "not-lambda", holds: "formula" } as const;
  const result = (e: Expr): CellCallee => {
    e = unwrap(e);
    if (e.kind === "lambda") return lambdaClass(e);
    if (e.kind === "let") return result(e.body);
    if (isConstant(e)) return notLambda;
    if (e.kind === "binary" && VALUE_OPS.has(e.op)) return notLambda;
    if (e.kind === "unary" && e.op !== "@") return notLambda;
    if (e.kind === "postfix" && e.op === "%") return notLambda;
    if (e.kind === "call" && !e.fn.qual) {
      const { prefix, base } = stripPrefix(e.fn.text);
      const info = prefix === "" || prefix.startsWith("_xlfn.") ? lookupFunction(base) : undefined;
      if (info && !PASS_THROUGH.has(info.name)) return notLambda;
    }
    return { kind: "unknown" };
  };
  return result(f.body);
}
