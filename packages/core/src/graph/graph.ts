// The cell dependency graph of a workbook (B6 b; the audit's C9, C10, C12 read it too).
//
// Nodes are what a modeller thinks in: a formula block (a dynamic array with the extent
// it spilled to when saved, a legacy array, a data table, or one cell of a normal or
// shared formula: the lines of the formula view), the input cells formulas read (one
// node per distinct referenced rectangle, minus nothing: see `inputNode`), and the
// defined names. An edge A → B reads "A depends on B". Everything is resolved on the
// syntax tree: cells, areas, whole rows and columns, `x#`, defined names (with sheet
// scope and LET/LAMBDA shadowing), structured references, 3-D and cross-sheet
// references. What cannot be resolved without evaluating (`INDIRECT`, `OFFSET` with
// computed arguments, other workbooks, `#REF!`) is recorded on the node as a flag, never
// guessed.
//
// The saved file is the oracle (FEASIBILITY §11.3): spill extents are those saved, so a
// spill that grew since is not seen grown.

import type { Sheet, Table, WorkbookSnapshot } from "../file/types.js";
import { parseCell } from "../file/cellref.js";
import { walk, type Expr, type Ref } from "../lang/ast.js";
import { lookupFunction } from "../lang/catalogue.js";
import { stripPrefix, tryParse } from "../lang/parser.js";

import { nameUses } from "../project/refs.js";
import { nameKey } from "../project/types.js";
import { cachedBody, sheetFormulaView, workbookNameIndex, type FormulaViewLine, type NameIndex } from "../view/formulas.js";
import { addressLabel, sheetLabel } from "../view/render.js";
import { area, intersect, isAbsolute, MAX_COL, MAX_ROW, rectOf, rectText, RectIndex, sameRect, type Rect } from "./rect.js";
import { parseStructInner, structRect } from "./structref.js";
import { DependencyGraph, type GraphFlag, type GraphFlagCode, type GraphNode, type SpillRefFinding } from "./model.js";

export interface GraphOptions {
  /** Formula parses shared with `sheetFormulaView` (per workbook). */
  cache?: Map<string, unknown>;
}

interface SheetEntry {
  sheet: Sheet;
  index: RectIndex<GraphNode>;
}

interface Ctx {
  node: GraphNode;
  /** Sheet that unqualified references and names belong to (undefined: a workbook name). */
  home: string | undefined;
  /** The formula's cell, for `[@Col]`. */
  row?: number;
  col?: number;
  src: string;
}

function unwrap(e: Expr): Expr {
  while (e.kind === "paren") e = e.expr;
  return e;
}

/** An integer literal (`3`, `-1`, `(2)`), or undefined. */
function literalInt(e: Expr | undefined): number | undefined {
  if (!e) return undefined;
  e = unwrap(e);
  if (e.kind === "number") {
    const v = Number(e.text);
    return Number.isInteger(v) ? v : undefined;
  }
  if (e.kind === "unary" && (e.op === "-" || e.op === "+")) {
    const v = literalInt(e.operand);
    return v === undefined ? undefined : e.op === "-" ? -v : v;
  }
  return undefined;
}

/**
 * A built-in function passed as a value (`GROUPBY(…, COUNTA)`, stored `_xleta.COUNTA`):
 * in display form it reads like a name. Only when no defined name has that spelling.
 */
function etaFunction(id: string): boolean {
  return lookupFunction(stripPrefix(id).base) !== undefined;
}

function cellRect(row: number, col: number): Rect {
  return { r1: row, c1: col, r2: row, c2: col };
}

/** Builds the dependency graph of `wb`. `names` resolves the names formulas read (default: the file's own). */
export function buildGraph(wb: WorkbookSnapshot, names: NameIndex = workbookNameIndex(wb), opts: GraphOptions = {}): DependencyGraph {
  const cache = opts.cache ?? new Map<string, unknown>();
  const nodes: GraphNode[] = [];
  const prec: Set<number>[] = [];
  const byKey = new Map<string, GraphNode>();
  const spillRefs: SpillRefFinding[] = [];

  const add = (n: Omit<GraphNode, "id" | "flags" | "level" | "cycle">): GraphNode => {
    const node: GraphNode = { ...n, id: nodes.length, flags: [], level: 0, cycle: undefined };
    nodes.push(node);
    prec.push(new Set());
    byKey.set(node.key, node);
    return node;
  };
  const edge = (from: GraphNode, to: GraphNode): void => {
    prec[from.id]!.add(to.id);
  };
  const flag = (node: GraphNode, kind: GraphFlag["kind"], code: GraphFlagCode, reason: string, text: string): void => {
    if (!node.flags.some((f) => f.kind === kind && f.reason === reason && f.text === text)) node.flags.push({ kind, code, reason, text });
  };

  // ---- formula nodes, one per line of each sheet's formula view ----------------------
  const sheets = new Map<string, SheetEntry>();
  const lines: { node: GraphNode; line: FormulaViewLine; sheet: Sheet }[] = [];
  for (const s of wb.sheets) {
    const own: GraphNode[] = [];
    if (s.formulas.length > 0) {
      for (const l of sheetFormulaView(wb, s.name, names, { cache })) {
        const rect = l.kind !== "shared" && l.rows !== undefined && l.cols !== undefined ? { r1: l.row, c1: l.col, r2: l.row + l.rows - 1, c2: l.col + l.cols - 1 } : cellRect(l.row, l.col);
        const node = add({ kind: "formula", key: `${s.name}!${l.cell}`, label: `${sheetLabel(s.name)}!${addressLabel(l)}`, sheet: s.name, position: s.position, rect, line: l, name: undefined });
        own.push(node);
        lines.push({ node, line: l, sheet: s });
      }
    }
    sheets.set(s.name.toLowerCase(), { sheet: s, index: new RectIndex(own) });
  }
  const sheetOf = (name: string): SheetEntry | undefined => sheets.get(name.toLowerCase());

  // ---- name nodes ---------------------------------------------------------------------
  const defs: { node: GraphNode; definition: string }[] = [];
  for (const d of wb.definedNames) {
    if (d.isXlPrefixed || d.scopeInvalid) continue;
    const scope = d.scope.kind === "sheet" ? d.scope.name : undefined;
    const key = nameKey({ name: d.name, scope });
    if (byKey.has(`name:${key}`)) continue;
    const parsed = tryParse(d.definition).formula;
    const lambda = parsed !== undefined && unwrap(parsed.body).kind === "lambda";
    const node = add({
      kind: "name",
      key: `name:${key}`,
      label: scope === undefined ? d.name : `${sheetLabel(scope)}!${d.name}`,
      sheet: scope,
      position: d.scope.kind === "sheet" ? d.scope.position : undefined,
      rect: undefined,
      line: undefined,
      name: { key, name: d.name, scope, definition: d.definition, lambda, hidden: d.hidden, defined: true },
    });
    defs.push({ node, definition: d.definition });
  }
  /** The node of a resolved name key; a key the file does not define (a project's own name) gets a node of its own. */
  const nameNode = (key: string): GraphNode => {
    const hit = byKey.get(`name:${key}`);
    if (hit) return hit;
    const bang = key.lastIndexOf("!");
    const scope = bang < 0 ? undefined : key.slice(0, bang);
    const name = bang < 0 ? key : key.slice(bang + 1);
    return add({
      kind: "name",
      key: `name:${key}`,
      label: scope === undefined ? name : `${sheetLabel(scope)}!${name}`,
      sheet: scope,
      position: undefined,
      rect: undefined,
      line: undefined,
      name: { key, name, scope, definition: undefined, lambda: false, hidden: false, defined: false },
    });
  };

  // ---- resolution -----------------------------------------------------------------------
  const tableNamed = new Map(wb.tables.map((t) => [t.displayName.toLowerCase(), t]));
  /** Input cells: one node per distinct rectangle a formula reads that formulas do not fill. */
  const inputNode = (sheet: Sheet, rect: Rect, label: string | undefined): GraphNode => {
    const key = `in:${sheet.name}!${rectText(rect)}`;
    const hit = byKey.get(key);
    if (hit) {
      if (label !== undefined && hit.label.includes("!")) hit.label = label;
      return hit;
    }
    return add({ kind: "input", key, label: label ?? `${sheetLabel(sheet.name)}!${rectText(rect)}`, sheet: sheet.name, position: sheet.position, rect, line: undefined, name: undefined });
  };

  const addRect = (ctx: Ctx, sheetName: string, rect: Rect, text: string, opts: { c9?: boolean; label?: string } = {}): void => {
    const sh = sheetOf(sheetName);
    if (!sh) {
      flag(ctx.node, "broken", "no-sheet", `no sheet '${sheetName}' in the workbook`, text);
      return;
    }
    let covered = 0;
    for (const hit of sh.index.query(rect)) {
      edge(ctx.node, hit);
      covered += area(intersect(hit.rect!, rect)!);
      const l = hit.line!;
      if (opts.c9 && hit !== ctx.node && l.kind === "dynamic-array" && area(hit.rect!) > 1 && !sameRect(rect, cellRect(l.row, l.col))) {
        const local = ctx.home !== undefined && ctx.home.toLowerCase() === sh.sheet.name.toLowerCase();
        const s = hit.rect!;
        const whole = `${local ? "" : sheetLabel(sh.sheet.name) + "!"}${l.cell}#`;
        const fit = sameRect(rect, s) ? "exact" : sameRect(intersect(rect, s)!, rect) ? "part" : "beyond";
        let use = whole;
        if (fit === "part" && area(rect) === 1) {
          // One cell of the spill: its position in it.
          const at = s.r1 === s.r2 ? [rect.c1 - s.c1 + 1] : s.c1 === s.c2 ? [rect.r1 - s.r1 + 1] : [rect.r1 - s.r1 + 1, rect.c1 - s.c1 + 1];
          use = `INDEX(${whole}, ${at.join(", ")})`;
        }
        spillRefs.push({ node: ctx.node, ref: text, sheet: sh.sheet.name, range: rectText(rect), spill: hit, use, fit });
      }
    }
    if (covered < area(rect)) edge(ctx.node, inputNode(sh.sheet, rect, opts.label));
  };

  const sheetsOf = (ctx: Ctx, r: Ref, text: string): string[] | undefined => {
    const q = r.qual;
    if (q?.sheet2 !== undefined && q.sheet !== undefined) {
      const a = sheetOf(q.sheet);
      const b = sheetOf(q.sheet2);
      if (!a || !b) {
        flag(ctx.node, "broken", "no-sheet", `3-D reference over a sheet the workbook does not have`, text);
        return undefined;
      }
      const lo = Math.min(a.sheet.position, b.sheet.position);
      const hi = Math.max(a.sheet.position, b.sheet.position);
      return wb.sheets.filter((s) => s.position >= lo && s.position <= hi).map((s) => s.name);
    }
    const sheet = q?.sheet ?? ctx.home;
    if (sheet === undefined) {
      flag(ctx.node, "dynamic", "unqualified-in-name", "reference without a sheet in a workbook-scoped name (it follows the active sheet)", text);
      return undefined;
    }
    return [sheet];
  };

  const textOf = (ctx: Ctx, e: Expr): string => ctx.src.slice(e.span.start, e.span.end);

  /**
   * A name no defined name resolves: a Table's own name (`SUM(tblSales)` reads its data
   * rows; Tables and defined names share one namespace), a built-in function passed as a
   * value, or an unknown name.
   */
  const unresolvedName = (ctx: Ctx, id: string): void => {
    const table = tableNamed.get(stripPrefix(id).base.toLowerCase());
    if (table) {
      const r = structRect(table, { areas: [], columns: [] }, undefined);
      if ("error" in r) flag(ctx.node, "broken", "table", r.error, id);
      else addRect(ctx, table.sheet.name, r.rect, id, { label: table.displayName });
    } else if (!etaFunction(id)) flag(ctx.node, "broken", "unknown-name", `unknown name '${id}' (#NAME?)`, id);
  };

  const ref = (ctx: Ctx, r: Ref, override?: Rect): void => {
    const text = textOf(ctx, r);
    if (r.refKind === "error") {
      flag(ctx.node, "broken", "ref-deleted", "#REF!: the reference was deleted", text);
      return;
    }
    if (r.qual?.book !== undefined) {
      flag(ctx.node, "external", "external", "reference to another workbook", text);
      return;
    }
    if (ctx.node.kind === "name" && !isAbsolute(r.address, r.refKind)) {
      // Stored relative to A1, it moves with the cell that reads the name.
      flag(ctx.node, "dynamic", "relative-in-name", "relative reference in a name: it points elsewhere from each cell that reads it", text);
      return;
    }
    const rect = override ?? rectOf(r.address, r.refKind);
    if (!rect) {
      flag(ctx.node, "broken", "unreadable", "unreadable reference", text);
      return;
    }
    for (const s of sheetsOf(ctx, r, text) ?? []) addRect(ctx, s, rect, text, { c9: override === undefined && (r.refKind === "cell" || r.refKind === "area") });
  };

  /** `x#` (stored `ANCHORARRAY(x)`) on a cell: the dynamic array anchored there. */
  const spill = (ctx: Ctx, operand: Expr, text: string): boolean => {
    const r = unwrap(operand);
    if (r.kind !== "ref") return false;
    if (r.qual?.book !== undefined) {
      flag(ctx.node, "external", "external", "reference to another workbook", text);
      return true;
    }
    if (ctx.node.kind === "name" && r.refKind === "cell" && !isAbsolute(r.address, r.refKind)) {
      flag(ctx.node, "dynamic", "relative-in-name", "relative reference in a name: it points elsewhere from each cell that reads it", text);
      return true;
    }
    const at = r.refKind === "cell" ? parseCell(r.address.split("$").join("")) : undefined;
    const sheet = r.qual?.sheet ?? ctx.home;
    if (!at || sheet === undefined) {
      flag(ctx.node, "broken", "unreadable", "spill reference on something other than one cell", text);
      return true;
    }
    const sh = sheetOf(sheet);
    const anchor = sh?.index.at(at.row, at.col);
    if (!sh || !anchor || anchor.rect!.r1 !== at.row || anchor.rect!.c1 !== at.col) {
      flag(ctx.node, "broken", "no-anchor", `${r.address}# : no formula anchored at that cell (Excel gives #REF!)`, text);
      return true;
    }
    edge(ctx.node, anchor);
    return true;
  };

  const tableAt = (sheet: string | undefined, row: number | undefined, col: number | undefined): Table | undefined => {
    if (sheet === undefined || row === undefined || col === undefined) return undefined;
    return wb.tables.find((t) => {
      if (t.sheet.name.toLowerCase() !== sheet.toLowerCase()) return false;
      const r = rectOf(t.ref, t.ref.includes(":") ? "area" : "cell");
      return r !== undefined && row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
    });
  };

  const resolve = (ctx: Ctx, body: Expr): void => {
    const skip = new Set<Expr>();
    walk(body, (n) => {
      if (skip.has(n)) return;
      switch (n.kind) {
        case "ref":
          ref(ctx, n);
          return;
        case "error":
          if (n.text.toUpperCase() === "#REF!") flag(ctx.node, "broken", "ref-deleted", "#REF!: the reference was deleted", n.text);
          return;
        case "name":
          if (n.qual?.book !== undefined) flag(ctx.node, "external", "external", "name in another workbook", textOf(ctx, n));
          return;
        case "postfix":
          if (n.op === "#" && spill(ctx, n.operand, textOf(ctx, n))) skip.add(unwrap(n.operand));
          return;
        case "binary": {
          // `A1:B9` written as two references (`A1 : B9`, `S!A1:S!B9`): the box they span.
          const l = unwrap(n.left);
          const r = unwrap(n.right);
          if ((n.op === ":" || n.op === ":." || n.op === ".:" || n.op === ".:.") && l.kind === "ref" && r.kind === "ref" && l.refKind === "cell" && r.refKind === "cell" && !l.qual?.book && (r.qual === undefined || r.qual.sheet === l.qual?.sheet)) {
            const a = rectOf(l.address, "cell");
            const b = rectOf(r.address, "cell");
            if (a && b) {
              ref(ctx, l, { r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1), r2: Math.max(a.r2, b.r2), c2: Math.max(a.c2, b.c2) });
              skip.add(l);
              skip.add(r);
            }
          }
          return;
        }
        case "call": {
          const fn = n.fn;
          if (fn.qual?.book !== undefined) {
            flag(ctx.node, "external", "external", "function or name in another workbook", textOf(ctx, n));
            return;
          }
          if (fn.qual) return;
          const base = stripPrefix(fn.text).base.toUpperCase();
          if (base === "INDIRECT") flag(ctx.node, "dynamic", "indirect", "INDIRECT: the reference is computed", textOf(ctx, n));
          else if (base === "ANCHORARRAY" && n.args.length === 1) {
            if (spill(ctx, n.args[0]!, textOf(ctx, n))) skip.add(unwrap(n.args[0]!));
          } else if (base === "OFFSET") {
            const target = n.args[0] && unwrap(n.args[0]);
            const [dr, dc, h, w] = [1, 2, 3, 4].map((k) => {
              const a = n.args[k];
              return a === undefined || a.kind === "missing" ? null : literalInt(a);
            });
            const fixed = dr !== undefined && dc !== undefined && h !== undefined && w !== undefined;
            const base = target?.kind === "ref" && (target.refKind === "cell" || target.refKind === "area") ? rectOf(target.address, target.refKind) : undefined;
            if (target?.kind === "ref" && base && fixed) {
              const r1 = base.r1 + (dr ?? 0);
              const c1 = base.c1 + (dc ?? 0);
              const r2 = r1 + (h ?? base.r2 - base.r1 + 1) - 1;
              const c2 = c1 + (w ?? base.c2 - base.c1 + 1) - 1;
              if (r1 < 1 || c1 < 1 || r2 > MAX_ROW || c2 > MAX_COL || r2 < r1 || c2 < c1) flag(ctx.node, "broken", "ref-deleted", "OFFSET moves off the sheet (#REF!)", textOf(ctx, n));
              else ref(ctx, target, { r1, c1, r2, c2 });
              skip.add(target);
            } else {
              flag(ctx.node, "dynamic", "offset", "OFFSET with computed arguments", textOf(ctx, n));
              // Its first argument is a position, not a value it reads.
              if (target?.kind === "ref") skip.add(target);
            }
          }
          return;
        }
        case "structref": {
          const text = textOf(ctx, n);
          const table = n.table === "" ? tableAt(ctx.home, ctx.row, ctx.col) : tableNamed.get(n.table.toLowerCase());
          if (!table) {
            flag(ctx.node, "broken", "table", n.table === "" ? "table reference outside a Table" : `no Table '${n.table}' in the workbook`, text);
            return;
          }
          const spec = parseStructInner(n.inner);
          if (!spec) {
            flag(ctx.node, "broken", "table", "unreadable table reference", text);
            return;
          }
          const r = structRect(table, spec, ctx.row);
          if ("error" in r) {
            flag(ctx.node, "broken", "table", r.error, text);
            return;
          }
          // A this-row reference reads one cell: named by its address, not by the column.
          const label = spec.areas.includes("#this row") ? undefined : r.label;
          addRect(ctx, table.sheet.name, r.rect, text, label !== undefined ? { label } : {});
          return;
        }
        default:
          return;
      }
    });
  };

  for (const { node, line, sheet } of lines) {
    const ctx: Ctx = { node, home: sheet.name, row: line.row, col: line.col, src: line.formula };
    if (line.error !== undefined) flag(node, "broken", "unparsable", `does not parse: ${line.error}`, line.stored);
    if (line.kind === "data-table") {
      // A data table recalculates the formulas in the row above it and the column to its
      // left (which also hold the values it substitutes) with its input cells replaced:
      // it reads all three.
      const f = sheet.formulas.find((x) => x.cell === line.cell);
      for (const a of [f?.attributes["r1"], f?.attributes["r2"]]) {
        const at = a && parseCell(a.split("$").join(""));
        if (at) addRect(ctx, sheet.name, cellRect(at.row, at.col), a);
      }
      const t = node.rect!;
      const above: Rect = { r1: t.r1 - 1, c1: Math.max(1, t.c1 - 1), r2: t.r1 - 1, c2: t.c2 };
      const left: Rect = { r1: t.r1, c1: t.c1 - 1, r2: t.r2, c2: t.c1 - 1 };
      if (above.r1 >= 1) addRect(ctx, sheet.name, above, rectText(above));
      if (left.c1 >= 1) addRect(ctx, sheet.name, left, rectText(left));
      continue;
    }
    for (const u of line.names) {
      if (u.key !== undefined) edge(node, nameNode(u.key));
      else unresolvedName(ctx, u.id);
    }
    const body = line.stored.trim() === "" ? undefined : cachedBody(cache, line.stored);
    if (body) resolve(ctx, body);
  }
  for (const { node, definition } of defs) {
    const parsed = tryParse(definition).formula;
    if (!parsed) {
      flag(node, "broken", "unparsable", "the definition does not parse", definition);
      continue;
    }
    const ctx: Ctx = { node, home: node.name!.scope, src: definition };
    for (const u of nameUses(parsed.body)) {
      const key = names.resolve(u, ctx.home);
      if (key !== undefined) edge(node, nameNode(key));
      else unresolvedName(ctx, u.id);
    }
    resolve(ctx, parsed.body);
  }

  // ---- names read outside the cells: conditional formats, validations, Table columns, built-ins
  const usedElsewhere = new Set<number>();
  const seed = (text: string | undefined, home: string | undefined): void => {
    if (!text) return;
    const parsed = tryParse(text).formula;
    if (!parsed) return;
    for (const u of nameUses(parsed.body)) {
      const key = names.resolve(u, home);
      const n = key !== undefined ? byKey.get(`name:${key}`) : undefined;
      if (n) usedElsewhere.add(n.id);
    }
  };
  for (const s of wb.sheets) {
    for (const cf of s.conditionalFormats) for (const f of cf.formulas) seed(f, s.name);
    for (const dv of s.dataValidations) {
      seed(dv.formula1, s.name);
      seed(dv.formula2, s.name);
    }
  }
  for (const t of wb.tables) for (const c of t.columns) {
    seed(c.calculatedColumnFormula, t.sheet.name);
    seed(c.totalsRowFormula, t.sheet.name);
  }
  for (const d of wb.definedNames) if (d.isBuiltIn) seed(d.definition, d.scope.kind === "sheet" ? d.scope.name : undefined);

  return new DependencyGraph(
    nodes,
    prec.map((s) => [...s].sort((a, b) => a - b)),
    spillRefs,
    usedElsewhere,
  );
}
