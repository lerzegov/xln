// The dependency graph as built by `buildGraph`: nodes, edges both ways, strongly
// connected components (circular references), levels, the calculation order, and the
// questions the audit asks of it (C9 spill references, C10 unused names, C12 name cycles).

import { parseCell } from "../file/cellref.js";
import type { FormulaViewLine } from "../view/formulas.js";
import type { Rect } from "./rect.js";

export type GraphNodeKind = "formula" | "input" | "name";

/**
 * What a flag is about, stable for tools (the audit maps them to its checks):
 * `ref-deleted` #REF!; `unknown-name`; `table` an unknown Table or column, or a table
 * reference outside a Table; `no-sheet`; `no-anchor` `x#` on a cell where no dynamic array
 * is anchored; `unreadable`; `unparsable`; `relative-in-name`; `unqualified-in-name` a
 * reference without a sheet in a workbook-scoped name; `indirect`; `offset`; `external`.
 */
export type GraphFlagCode =
  | "ref-deleted"
  | "unknown-name"
  | "table"
  | "no-sheet"
  | "no-anchor"
  | "unreadable"
  | "unparsable"
  | "relative-in-name"
  | "unqualified-in-name"
  | "indirect"
  | "offset"
  | "external";

/** Why a node's dependencies are incomplete or wrong: they cannot be read from the file alone. */
export interface GraphFlag {
  /** `dynamic`: computed reference (INDIRECT, OFFSET); `external`: another workbook; `broken`: #REF!, unknown name or Table, unparsable. */
  kind: "dynamic" | "external" | "broken";
  code: GraphFlagCode;
  reason: string;
  /** The text in the formula or definition it is about. */
  text: string;
}

export interface GraphName {
  /** `Sheet!Name` or `Name`, as `nameKey` and `NameIndex.resolve` give it. */
  key: string;
  name: string;
  /** Sheet of a sheet-scoped name. */
  scope: string | undefined;
  /** The definition as stored; undefined for a name the file does not define. */
  definition: string | undefined;
  lambda: boolean;
  hidden: boolean;
  /** Whether the workbook defines it (a project may resolve names the file lacks). */
  defined: boolean;
}

export interface GraphNode {
  /** Index in `DependencyGraph.nodes`. */
  id: number;
  kind: GraphNodeKind;
  /** Unique: `Sheet!C6` (formula), `in:Sheet!A1:B3` (input), `name:Sheet!Name` (name). */
  key: string;
  /** For people: `BS!C6#`, `tblAssumpt[base]`, `Assumpt!C36`, `BS!AccountsReceivable_base`. */
  label: string;
  /** The sheet of a formula or input; a name's scope. */
  sheet: string | undefined;
  /** Sheet position (0-based), for ordering. */
  position: number | undefined;
  /** The cells a formula block or input covers. */
  rect: Rect | undefined;
  /** The formula view line of a formula node. */
  line: FormulaViewLine | undefined;
  name: GraphName | undefined;
  flags: GraphFlag[];
  /**
   * Longest chain of formulas from the inputs: 0 for inputs (and names that read nothing),
   * 1 for a formula that reads only inputs; a name has its target's level. Members of a
   * cycle share one.
   */
  level: number;
  /** The circular reference (1-based number into `cycles`) this node is part of. */
  cycle: number | undefined;
}

export interface GraphCycle {
  /** 1-based. */
  id: number;
  /** In sheet order, then row and column; names and inputs last. */
  members: GraphNode[];
}

/** C9: a fixed reference to cells of a spill (`C10:G10`) where `C10#` would follow the spill. */
export interface SpillRefFinding {
  /** The formula or defined name holding the reference. */
  node: GraphNode;
  /** The reference as written. */
  ref: string;
  sheet: string;
  /** The cells it covers (`C10:G10`). */
  range: string;
  /** The dynamic array whose saved extent it overlaps. */
  spill: GraphNode;
  /**
   * What to write instead, qualified when on another sheet than the reader's: `C10#`; for
   * one cell inside the spill, `INDEX(C10#, 3)` (`INDEX(C10#, 2, 3)` for a 2-D spill).
   */
  use: string;
  /**
   * `exact`: the reference covers exactly the saved extent; `part`: it lies inside it;
   * `beyond`: it reaches cells outside the spill too (`B5:F6` over the spill `B6:F6`).
   */
  fit: "exact" | "part" | "beyond";
}

/** C10: names nothing reads, and names read only by such names. */
export interface UnusedNames {
  unused: GraphNode[];
  onlyByUnused: GraphNode[];
}

type Key = [number, number, number, number];

function lessKey(a: Key, b: Key): boolean {
  for (let k = 0; k < 4; k++) if (a[k] !== b[k]) return a[k]! < b[k]!;
  return false;
}

/** A binary min-heap of components by key. */
class Heap {
  private readonly items: { key: Key; comp: number }[] = [];
  get size(): number {
    return this.items.length;
  }
  push(key: Key, comp: number): void {
    const a = this.items;
    a.push({ key, comp });
    let i = a.length - 1;
    while (i > 0) {
      const p = (i - 1) >> 1;
      if (!lessKey(a[i]!.key, a[p]!.key)) break;
      [a[i], a[p]] = [a[p]!, a[i]!];
      i = p;
    }
  }
  pop(): number {
    const a = this.items;
    const top = a[0]!.comp;
    const last = a.pop()!;
    if (a.length > 0) {
      a[0] = last;
      let i = 0;
      for (;;) {
        const l = 2 * i + 1;
        const r = l + 1;
        let m = i;
        if (l < a.length && lessKey(a[l]!.key, a[m]!.key)) m = l;
        if (r < a.length && lessKey(a[r]!.key, a[m]!.key)) m = r;
        if (m === i) break;
        [a[i], a[m]] = [a[m]!, a[i]!];
        i = m;
      }
    }
    return top;
  }
}

/**
 * Strongly connected components (Tarjan, iterative: chains of formulas run thousands
 * deep). With edges pointing at precedents, a component comes out after every component
 * it depends on, so the output is already a calculation order.
 */
export function tarjan(n: number, adj: readonly (readonly number[])[]): { comp: Int32Array; comps: number[][] } {
  const index = new Int32Array(n).fill(-1);
  const low = new Int32Array(n);
  const onStack = new Uint8Array(n);
  const comp = new Int32Array(n).fill(-1);
  const comps: number[][] = [];
  const stack: number[] = [];
  let counter = 0;
  const callNode: number[] = [];
  const callEdge: number[] = [];
  for (let s = 0; s < n; s++) {
    if (index[s] !== -1) continue;
    callNode.push(s);
    callEdge.push(0);
    index[s] = low[s] = counter++;
    stack.push(s);
    onStack[s] = 1;
    while (callNode.length > 0) {
      const v = callNode[callNode.length - 1]!;
      const k = callEdge[callEdge.length - 1]!;
      const out = adj[v]!;
      if (k < out.length) {
        callEdge[callEdge.length - 1] = k + 1;
        const w = out[k]!;
        if (index[w] === -1) {
          index[w] = low[w] = counter++;
          stack.push(w);
          onStack[w] = 1;
          callNode.push(w);
          callEdge.push(0);
        } else if (onStack[w]) low[v] = Math.min(low[v]!, index[w]!);
        continue;
      }
      callNode.pop();
      callEdge.pop();
      if (callNode.length > 0) {
        const u = callNode[callNode.length - 1]!;
        low[u] = Math.min(low[u]!, low[v]!);
      }
      if (low[v] === index[v]) {
        const members: number[] = [];
        let w: number;
        do {
          w = stack.pop()!;
          onStack[w] = 0;
          comp[w] = comps.length;
          members.push(w);
        } while (w !== v);
        comps.push(members.sort((a, b) => a - b));
      }
    }
  }
  return { comp, comps };
}

const KIND_RANK: Record<GraphNodeKind, number> = { formula: 0, name: 1, input: 2 };

export class DependencyGraph {
  /** Number of edges (distinct "depends on" pairs). */
  readonly edges: number;
  /** Circular references: components of two or more nodes, or a node reading itself, other than LAMBDA recursion. */
  readonly cycles: GraphCycle[];
  /** LAMBDA names calling themselves or each other: recursion, which Excel allows. */
  readonly recursions: GraphCycle[];
  private readonly dependentIds: number[][];
  private readonly byKey: Map<string, GraphNode>;
  private readonly comp: Int32Array;
  /** Components in a calculation order (precedents first). */
  private readonly comps: number[][];
  private readonly compPrec: number[][];
  private readonly orders = new Map<string, GraphNode[]>();

  constructor(
    readonly nodes: GraphNode[],
    private readonly precedentIds: number[][],
    /** C9 findings: fixed references into a spill, in formulas and in definitions. */
    readonly spillRefs: SpillRefFinding[],
    /** Names read by conditional formats, data validations, Table column formulas or built-in names. */
    private readonly usedElsewhere: ReadonlySet<number>,
  ) {
    this.byKey = new Map(nodes.map((n) => [n.key, n]));
    this.dependentIds = nodes.map(() => []);
    let edges = 0;
    precedentIds.forEach((ps, a) => {
      edges += ps.length;
      for (const b of ps) this.dependentIds[b]!.push(a);
    });
    this.edges = edges;

    const { comp, comps } = tarjan(nodes.length, precedentIds);
    this.comp = comp;
    this.comps = comps.map((c) => [...c].sort((a, b) => this.compareNodes(nodes[a]!, nodes[b]!)));
    this.compPrec = comps.map((members, c) => {
      const s = new Set<number>();
      for (const m of members) for (const p of precedentIds[m]!) if (comp[p] !== c) s.add(comp[p]!);
      return [...s];
    });

    // Levels, in Tarjan's order (precedents first).
    const level = new Int32Array(comps.length);
    comps.forEach((members, c) => {
      let max = 0;
      for (const p of this.compPrec[c]!) max = Math.max(max, level[p]!);
      level[c] = max + (members.some((m) => nodes[m]!.kind === "formula") ? 1 : 0);
      for (const m of members) nodes[m]!.level = level[c]!;
    });

    this.cycles = [];
    this.recursions = [];
    for (const members of this.comps) {
      const self = members.length === 1 && precedentIds[members[0]!]!.includes(members[0]!);
      if (members.length < 2 && !self) continue;
      const ms = members.map((m) => nodes[m]!);
      if (ms.every((n) => n.kind === "name" && n.name!.lambda)) {
        this.recursions.push({ id: this.recursions.length + 1, members: ms });
        continue;
      }
      const id = this.cycles.length + 1;
      for (const n of ms) n.cycle = id;
      this.cycles.push({ id, members: ms });
    }
    // Number cycles in sheet order of their first member.
    this.cycles.sort((a, b) => this.compareNodes(a.members[0]!, b.members[0]!));
    this.cycles.forEach((c, k) => {
      c.id = k + 1;
      for (const n of c.members) n.cycle = c.id;
    });
  }

  private compareNodes(a: GraphNode, b: GraphNode): number {
    return (
      KIND_RANK[a.kind] - KIND_RANK[b.kind] ||
      (a.position ?? -1) - (b.position ?? -1) ||
      (a.rect?.r1 ?? 0) - (b.rect?.r1 ?? 0) ||
      (a.rect?.c1 ?? 0) - (b.rect?.c1 ?? 0) ||
      a.id - b.id
    );
  }

  private id(n: GraphNode | number): number {
    return typeof n === "number" ? n : n.id;
  }

  /** The node with this key (`Sheet!C6`, `in:Sheet!A1:B3`, `name:Sheet!Name`). */
  node(key: string): GraphNode | undefined {
    return this.byKey.get(key);
  }

  /** The node of a defined name, by its key (`Sheet!Name` or `Name`). */
  nameNode(key: string): GraphNode | undefined {
    return this.byKey.get(`name:${key}`);
  }

  /** The formula block holding a cell: its own, or the array, spill or data table over it. */
  formulaAt(sheet: string, cell: string): GraphNode | undefined {
    const at = cell.split("$").join("");
    const direct = this.byKey.get(`${sheet}!${at}`);
    if (direct?.kind === "formula") return direct;
    const a = parseCell(at);
    if (!a) return undefined;
    const { row, col } = a;
    const low = sheet.toLowerCase();
    return this.nodes.find((n) => n.kind === "formula" && n.sheet!.toLowerCase() === low && n.rect!.r1 <= row && row <= n.rect!.r2 && n.rect!.c1 <= col && col <= n.rect!.c2);
  }

  /** What `n` reads directly. */
  precedents(n: GraphNode | number): GraphNode[] {
    return this.precedentIds[this.id(n)]!.map((i) => this.nodes[i]!);
  }

  /** What reads `n` directly. */
  dependents(n: GraphNode | number): GraphNode[] {
    return this.dependentIds[this.id(n)]!.map((i) => this.nodes[i]!);
  }

  private reach(start: number, adj: number[][]): Set<number> {
    const seen = new Set<number>([start]);
    const todo = [start];
    while (todo.length > 0) for (const p of adj[todo.pop()!]!) if (!seen.has(p)) {
      seen.add(p);
      todo.push(p);
    }
    seen.delete(start);
    return seen;
  }

  /** Everything `n` depends on, directly or not. */
  allPrecedents(n: GraphNode | number): GraphNode[] {
    return [...this.reach(this.id(n), this.precedentIds)].sort((a, b) => a - b).map((i) => this.nodes[i]!);
  }

  /** Everything that depends on `n`, directly or not. */
  allDependents(n: GraphNode | number): GraphNode[] {
    return [...this.reach(this.id(n), this.dependentIds)].sort((a, b) => a - b).map((i) => this.nodes[i]!);
  }

  /** The inputs `n` depends on, directly or not: input cells, and names that read nothing (constants). */
  inputsOf(n: GraphNode | number): GraphNode[] {
    return this.allPrecedents(n).filter((x) => x.kind === "input" || (x.kind === "name" && this.precedentIds[x.id]!.length === 0 && x.name!.defined));
  }

  /**
   * Nodes in calculation order: every node after what it depends on, a cycle's members
   * together. Among nodes that are ready, formulas come in sheet order, then row, then
   * column, so a sheet already laid out top-down keeps its order; names and inputs come as
   * soon as they are ready. With `sheet`, the order is the best one for that sheet's
   * formulas (other sheets' formulas come as early as they can, so they never hold back
   * a formula of the sheet that could come earlier).
   */
  order(opts: { sheet?: string } = {}): GraphNode[] {
    const cacheKey = opts.sheet === undefined ? "" : "!" + opts.sheet.toLowerCase();
    const hit = this.orders.get(cacheKey);
    if (hit) return hit;
    const low = opts.sheet?.toLowerCase();
    const key = (c: number): Key => {
      let best: Key = [-1, -1, -1, c];
      for (const m of this.comps[c]!) {
        const n = this.nodes[m]!;
        if (n.kind !== "formula") continue;
        if (low !== undefined && n.sheet!.toLowerCase() !== low) continue;
        const k: Key = [low !== undefined ? 0 : n.position!, n.rect!.r1, n.rect!.c1, c];
        if (best[0] === -1 || lessKey(k, best)) best = k;
      }
      return best;
    };
    const waiting = this.compPrec.map((p) => p.length);
    const compDeps: number[][] = this.comps.map(() => []);
    this.compPrec.forEach((ps, c) => {
      for (const p of ps) compDeps[p]!.push(c);
    });
    const heap = new Heap();
    waiting.forEach((w, c) => {
      if (w === 0) heap.push(key(c), c);
    });
    const out: GraphNode[] = [];
    while (heap.size > 0) {
      const c = heap.pop();
      for (const m of this.comps[c]!) out.push(this.nodes[m]!);
      for (const d of compDeps[c]!) if (--waiting[d]! === 0) heap.push(key(d), d);
    }
    this.orders.set(cacheKey, out);
    return out;
  }

  /** Formula nodes with a flag (dynamic, external or broken references). */
  flagged(): GraphNode[] {
    return this.nodes.filter((n) => n.flags.length > 0);
  }

  /**
   * C10. A name is used when a formula reads it (directly, or through other names), or a
   * conditional format, data validation, Table column formula or built-in name does.
   * `unused`: nothing reads it; `onlyByUnused`: read only by unused names. Names in charts
   * are not seen (the snapshot does not read chart parts).
   */
  unusedNames(): UnusedNames {
    const live = new Set<number>();
    const todo: number[] = [];
    for (const n of this.nodes) {
      if (n.kind === "formula" || this.usedElsewhere.has(n.id)) {
        live.add(n.id);
        todo.push(n.id);
      }
    }
    while (todo.length > 0) for (const p of this.precedentIds[todo.pop()!]!) if (!live.has(p)) {
      live.add(p);
      todo.push(p);
    }
    const unused: GraphNode[] = [];
    const onlyByUnused: GraphNode[] = [];
    for (const n of this.nodes) {
      if (n.kind !== "name" || !n.name!.defined || live.has(n.id)) continue;
      const users = this.dependentIds[n.id]!.filter((d) => d !== n.id);
      (users.length === 0 ? unused : onlyByUnused).push(n);
    }
    return { unused, onlyByUnused };
  }

  /**
   * C12: cycles among defined names through their definitions alone (name → name edges).
   * `recursive` cycles are LAMBDAs calling themselves or each other, which Excel allows.
   */
  nameCycles(): (GraphCycle & { recursive: boolean })[] {
    const names = this.nodes.filter((n) => n.kind === "name");
    const local = new Map(names.map((n, k) => [n.id, k]));
    const adj = names.map((n) => this.precedentIds[n.id]!.filter((p) => local.has(p)).map((p) => local.get(p)!));
    const { comps } = tarjan(names.length, adj);
    const out: (GraphCycle & { recursive: boolean })[] = [];
    for (const c of comps) {
      if (c.length < 2 && !adj[c[0]!]!.includes(c[0]!)) continue;
      const members = c.map((k) => names[k]!).sort((a, b) => a.label.localeCompare(b.label));
      out.push({ id: out.length + 1, members, recursive: members.every((m) => m.name!.lambda) });
    }
    return out;
  }

  /** Counts for a summary. */
  stats(): { formulas: number; inputs: number; names: number; edges: number; cycles: number; recursions: number; dynamic: number; external: number; broken: number } {
    const count = (k: GraphNodeKind) => this.nodes.filter((n) => n.kind === k).length;
    const flagged = (k: GraphFlag["kind"]) => this.nodes.filter((n) => n.flags.some((f) => f.kind === k)).length;
    return {
      formulas: count("formula"),
      inputs: count("input"),
      names: count("name"),
      edges: this.edges,
      cycles: this.cycles.length,
      recursions: this.recursions.length,
      dynamic: flagged("dynamic"),
      external: flagged("external"),
      broken: flagged("broken"),
    };
  }
}
