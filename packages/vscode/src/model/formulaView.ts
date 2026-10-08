// The formula view of one sheet (B6 a), or of a sheet or the whole workbook in
// calculation order (B6 b), as the editor needs it: the core's lines and rendered text,
// plus lookups from a document offset to the name, reference or entry under it, and
// from a cell to its entry; in calculation order, what a line reads and what reads it.
// Plain TypeScript without the vscode API.
//
// Names resolve against the project's live definitions when the workbook has been
// pulled (so a click lands on the entry in its `.xln` file), otherwise against the
// workbook's own names (still shown and hovered, but there is nothing to go to).

import {
  buildGraph,
  formatCell,
  NameResolver,
  orderText,
  parseCell,
  relativeLabel,
  renderFormulaView,
  sheetCalcView,
  sheetFormulaView,
  workbookFormulaView,
  workbookNameIndex,
  type DependencyGraph,
  type FormulaOrder,
  type FormulaViewLine,
  type GraphNode,
  type NameIndex,
  type RenderedEntry,
  type WorkbookSnapshot,
} from "@xln/core";
import type { Project } from "./project.js";

/** The sheet part of a workbook view's URI: `[` cannot occur in a sheet name. */
export const WORKBOOK_VIEW = "[workbook]";

export interface FormulaViewDoc {
  workbook: string;
  /** The sheet, or `WORKBOOK_VIEW` for every sheet in calculation order. */
  sheet: string;
  order: FormulaOrder;
  /** In calculation order: the graph the order comes from. */
  graph?: DependencyGraph;
  text: string;
  lines: FormulaViewLine[];
  entries: RenderedEntry[];
  /** Whether names resolve against a pulled project (they can be followed). */
  linked: boolean;
}

/** The project's names as an index, keyed like `Project.lookup`. */
export function projectNameIndex(project: Project): NameIndex {
  return new NameResolver(project.defs.map((d) => ({ name: d.name, scope: d.scope, key: d.key })));
}

/** The dependency graph of a snapshot with a project's names (or the file's). */
export function graphFor(wb: WorkbookSnapshot, project: Project | undefined, cache?: Map<string, unknown>): DependencyGraph {
  return buildGraph(wb, project ? projectNameIndex(project) : workbookNameIndex(wb), cache ? { cache } : {});
}

/**
 * Builds the view of `sheet` (or of the whole workbook, `WORKBOOK_VIEW`, in calculation
 * order); undefined when the workbook has no such sheet. `cache` is shared across the
 * sheets of one workbook snapshot (formula parses); `graph` across views in calculation
 * order (built here when missing).
 */
export function buildFormulaView(
  wb: WorkbookSnapshot,
  sheet: string,
  workbook: string,
  project: Project | undefined,
  cache?: Map<string, unknown>,
  order: FormulaOrder = "appearance",
  graph?: DependencyGraph,
): FormulaViewDoc | undefined {
  const all = sheet === WORKBOOK_VIEW;
  const s = all ? undefined : (wb.sheets.find((x) => x.name === sheet) ?? wb.sheets.find((x) => x.name.toLowerCase() === sheet.toLowerCase()));
  if (!s && !all) return undefined;
  if (all) order = "calculation";
  const index = project ? projectNameIndex(project) : workbookNameIndex(wb);
  const opts = cache ? { cache } : {};
  let lines: FormulaViewLine[];
  if (order === "calculation") {
    graph ??= buildGraph(wb, index, opts);
    lines = s ? sheetCalcView(wb, s.name, index, { graph }) : workbookFormulaView(wb, index, { graph });
  } else lines = sheetFormulaView(wb, s!.name, index, opts);
  const r = renderFormulaView(lines, {
    sheet: s?.name ?? "",
    workbook,
    ...(order === "calculation" ? { order: orderText(order), levels: true, sheets: all } : {}),
  });
  let text = r.text;
  if (!project) {
    // Appended, so the offsets of the entries stay as rendered.
    text += `\n// Names are not linked: pull ${workbook} (xln: Pull workbook) to follow them to their definitions.\n`;
  }
  const doc: FormulaViewDoc = { workbook, sheet: s?.name ?? WORKBOOK_VIEW, order, text, lines, entries: r.entries, linked: project !== undefined };
  if (order === "calculation") doc.graph = graph!;
  return doc;
}

function within(s: { start: number; end: number }, offset: number): boolean {
  return s.start <= offset && offset <= s.end;
}

/** The entry whose lines contain `offset`. */
export function entryAt(doc: FormulaViewDoc, offset: number): RenderedEntry | undefined {
  // Entries are in document order: binary search on their start (names, or the address).
  let lo = 0;
  let hi = doc.entries.length - 1;
  let found: RenderedEntry | undefined;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const e = doc.entries[mid]!;
    if (e.start <= offset) {
      found = e;
      lo = mid + 1;
    } else hi = mid - 1;
  }
  return found;
}

/** The name use under `offset`: its line, index and resolved key. */
export function nameAt(doc: FormulaViewDoc, offset: number): { entry: RenderedEntry; line: FormulaViewLine; k: number } | undefined {
  const entry = entryAt(doc, offset);
  if (!entry) return undefined;
  const k = entry.names.findIndex((s) => within(s, offset));
  return k < 0 ? undefined : { entry, line: doc.lines[entry.index]!, k };
}

/** The defined name on the left of an entry (its left-hand side) under `offset`. */
export function lhsAt(doc: FormulaViewDoc, offset: number): { entry: RenderedEntry; line: FormulaViewLine; k: number } | undefined {
  const entry = entryAt(doc, offset);
  if (!entry) return undefined;
  const k = entry.lhs.findIndex((s) => within(s, offset));
  return k < 0 ? undefined : { entry, line: doc.lines[entry.index]!, k };
}

/** Outline label of an entry: `Name — C6#`, or the address alone when the cell has no name; `Sheet!C6#` with `sheet`. */
export function entryLabel(l: FormulaViewLine, sheet = false): string {
  const address = (sheet ? `${l.sheet}!` : "") + (l.kind === "dynamic-array" ? `${l.cell}#` : l.cell);
  return l.lhs.length > 0 ? `${l.lhs.map((n) => n.display).join(", ")} — ${address}` : address;
}

/** The cell reference under `offset`. */
export function refAt(doc: FormulaViewDoc, offset: number): { entry: RenderedEntry; line: FormulaViewLine; k: number } | undefined {
  const entry = entryAt(doc, offset);
  if (!entry) return undefined;
  const k = entry.refs.findIndex((s) => within(s, offset));
  return k < 0 ? undefined : { entry, line: doc.lines[entry.index]!, k };
}

/** Whether `offset` is on an entry's cell address. */
export function addressAt(doc: FormulaViewDoc, offset: number): RenderedEntry | undefined {
  const e = entryAt(doc, offset);
  return e && within({ start: e.address.start, end: e.address.end + 1 }, offset) ? e : undefined;
}

/**
 * The entry showing `cell` (`B3`, `$B$3`, or the top-left of `B3:C9`): the cell's own,
 * or the array, spill or data table whose saved extent holds it. In the workbook view,
 * `sheet` says on which sheet.
 */
export function entryForCell(doc: FormulaViewDoc, cell: string, sheet?: string): RenderedEntry | undefined {
  const a = parseCell(cell.split(":")[0]!.split("$").join(""));
  if (!a) return undefined;
  const on = (l: FormulaViewLine) => sheet === undefined || l.sheet.toLowerCase() === sheet.toLowerCase();
  for (const e of doc.entries) {
    const l = doc.lines[e.index]!;
    if (l.row === a.row && l.col === a.col && on(l)) return e;
  }
  for (const e of doc.entries) {
    const l = doc.lines[e.index]!;
    if (l.kind === "shared" || l.rows === undefined || l.cols === undefined || !on(l)) continue;
    if (a.row >= l.row && a.row < l.row + l.rows && a.col >= l.col && a.col < l.col + l.cols) return e;
  }
  return undefined;
}

/** Hover text for an entry's address: kind, extent, stored text, value. */
export function entryHover(doc: FormulaViewDoc, e: RenderedEntry): string {
  const l = doc.lines[e.index]!;
  const kind: Record<FormulaViewLine["kind"], string> = {
    normal: "formula",
    shared: "shared formula",
    array: "legacy array formula (Ctrl+Shift+Enter)",
    "dynamic-array": "dynamic array",
    "data-table": "data table",
  };
  const out = [`**${l.sheet}!${l.cell}** · ${kind[l.kind]}`];
  if (l.level !== undefined) out.push(`level ${l.level}${l.cycle !== undefined ? ` · part of circular reference ↻${l.cycle}` : ""}`);
  if (l.kind === "dynamic-array" && l.extent) out.push(`spilled to \`${l.extent}\` when saved (${l.rows}×${l.cols})`);
  else if (l.kind === "shared") out.push(`group \`${l.extent ?? "?"}\` (${l.groupSize ?? "?"} cells), text stored in ${l.master ?? "?"}`);
  else if (l.extent && l.extent !== l.cell) out.push(`over \`${l.extent}\``);
  if (l.stored) out.push("stored as:\n```\n" + l.stored + "\n```");
  if (l.valueText !== undefined) out.push(`saved value: \`${l.valueText}\``);
  if (l.error) out.push(`does not parse: ${l.error}`);
  return out.join("\n\n");
}

/** A node of the graph to show in a hover: its label from the line's sheet, and the cell to open, if any. */
export interface HoverNode {
  label: string;
  kind: GraphNode["kind"];
  /** Sheet and top-left cell of a formula block (only formulas have a line to open). */
  sheet?: string;
  cell?: string;
  /** The defined name it is reached through, when not directly. */
  via?: string;
}

/**
 * What an entry reads and what reads it: from the graph of a calculation-order view, or
 * `graph`. Names come first, as they are written; then the cells, directly or through
 * names (`via`): a model that names every block reads almost everything through a name,
 * and the hover should still lead to the line behind it.
 */
export function entryLinks(doc: FormulaViewDoc, e: RenderedEntry, graph = doc.graph): { precedents: HoverNode[]; dependents: HoverNode[] } | undefined {
  const g = graph;
  const l = doc.lines[e.index]!;
  const node = g?.formulaAt(l.sheet, l.cell);
  if (!g || !node) return undefined;
  const show = (n: GraphNode, via?: GraphNode): HoverNode => {
    const h: HoverNode = { label: relativeLabel(n, l.sheet), kind: n.kind };
    if (n.kind === "formula" && n.rect && n.sheet !== undefined) {
      h.sheet = n.sheet;
      h.cell = formatCell({ row: n.rect.r1, col: n.rect.c1 });
    }
    if (via) h.via = relativeLabel(via, l.sheet);
    return h;
  };
  const links = (step: (n: GraphNode) => GraphNode[]): HoverNode[] => {
    const direct = step(node);
    const names = direct.filter((n) => n.kind === "name").sort((a, b) => a.id - b.id);
    const out = names.map((n) => show(n));
    const seen = new Set<number>([node.id]);
    const cells: [GraphNode, GraphNode | undefined][] = [];
    for (const n of direct) if (n.kind !== "name" && !seen.has(n.id)) {
      seen.add(n.id);
      cells.push([n, undefined]);
    }
    // Through chains of names (`Total` reads `Spill` reads C6#), breadth first.
    const visited = new Set<number>(names.map((n) => n.id));
    const todo = names.map((n) => [n, n] as const);
    while (todo.length > 0) {
      const [at, first] = todo.shift()!;
      for (const m of step(at)) {
        if (m.kind === "name") {
          if (!visited.has(m.id)) {
            visited.add(m.id);
            todo.push([m, first]);
          }
        } else if (!seen.has(m.id)) {
          seen.add(m.id);
          cells.push([m, first]);
        }
      }
    }
    cells.sort((a, b) => a[0].id - b[0].id);
    return [...out, ...cells.map(([n, via]) => show(n, via))];
  };
  return { precedents: links((n) => g.precedents(n)), dependents: links((n) => g.dependents(n)) };
}

