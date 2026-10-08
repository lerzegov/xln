// B6 (b): the formula view in calculation order, from the dependency graph. The same
// lines as in order of appearance, each with its level (the longest chain of formulas
// from the inputs), the circular reference it belongs to, and what it reads directly
// (defined names first). For one sheet, or for the whole workbook across sheets, since
// dependencies cross sheets.

import type { WorkbookSnapshot } from "../file/types.js";
import { buildGraph } from "../graph/graph.js";
import type { DependencyGraph, GraphNode } from "../graph/model.js";
import { workbookNameIndex, type FormulaViewLine, type NameIndex } from "./formulas.js";

export type FormulaOrder = "appearance" | "calculation";

export interface CalcViewOptions {
  /** Formula parses shared across calls (per workbook). */
  cache?: Map<string, unknown>;
  /** A graph already built for the workbook (with the same names). */
  graph?: DependencyGraph;
}

/** How `node` reads from a formula on `sheet`: bare on the same sheet (and for workbook names), `Sheet!…` otherwise. */
export function relativeLabel(node: GraphNode, sheet: string): string {
  const same = node.sheet !== undefined && node.sheet.toLowerCase() === sheet.toLowerCase();
  if (node.kind === "name") return node.name!.scope === undefined || same ? node.name!.name : node.label;
  if (!same || !node.label.includes("!")) return node.label;
  return node.label.slice(node.label.lastIndexOf("!") + 1);
}

/** What `node` reads directly, as seen from its sheet: defined names first (as written), then cells. */
export function dependsOn(graph: DependencyGraph, node: GraphNode): string[] {
  const sheet = node.sheet ?? "";
  const ps = graph.precedents(node);
  const names = ps.filter((p) => p.kind === "name");
  const order = new Map<string, number>();
  node.line?.names.forEach((u, k) => {
    if (u.key !== undefined && !order.has(u.key)) order.set(u.key, k);
  });
  names.sort((a, b) => (order.get(a.name!.key) ?? 1e9) - (order.get(b.name!.key) ?? 1e9) || a.id - b.id);
  const cells = ps.filter((p) => p.kind !== "name");
  return [...names, ...cells].map((p) => relativeLabel(p, sheet));
}

function annotate(graph: DependencyGraph, nodes: GraphNode[]): FormulaViewLine[] {
  return nodes.map((n) => {
    const line: FormulaViewLine = { ...n.line!, level: n.level, dependsOn: dependsOn(graph, n) };
    if (n.cycle !== undefined) line.cycle = n.cycle;
    return line;
  });
}

/**
 * The formulas of `sheetName` in calculation order: each after every formula it depends
 * on (on any sheet, through names too), the members of a circular reference together;
 * otherwise in order of appearance, so a sheet laid out top-down keeps its order.
 * Throws if there is no such sheet.
 */
export function sheetCalcView(wb: WorkbookSnapshot, sheetName: string, names: NameIndex = workbookNameIndex(wb), opts: CalcViewOptions = {}): FormulaViewLine[] {
  const sheet = wb.sheets.find((s) => s.name === sheetName) ?? wb.sheets.find((s) => s.name.toLowerCase() === sheetName.toLowerCase());
  if (!sheet) throw new Error(`no sheet '${sheetName}' in the workbook`);
  const graph = opts.graph ?? buildGraph(wb, names, opts.cache ? { cache: opts.cache } : {});
  return annotate(graph, graph.order({ sheet: sheet.name }).filter((n) => n.kind === "formula" && n.sheet === sheet.name));
}

/** Every formula of the workbook in calculation order; among formulas ready together, sheet order, then row and column. */
export function workbookFormulaView(wb: WorkbookSnapshot, names: NameIndex = workbookNameIndex(wb), opts: CalcViewOptions = {}): FormulaViewLine[] {
  const graph = opts.graph ?? buildGraph(wb, names, opts.cache ? { cache: opts.cache } : {});
  return annotate(graph, graph.order().filter((n) => n.kind === "formula"));
}

/** The header text for an order. */
export function orderText(order: FormulaOrder): string {
  return order === "calculation" ? "calculation order (each formula after what it reads)" : "order of appearance (row by row, left to right)";
}

