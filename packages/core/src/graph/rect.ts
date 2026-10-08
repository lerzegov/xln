// Rectangles of cells and a per-sheet index over the formula blocks, so a reference to
// an area finds the blocks it overlaps without visiting its cells: `A:A` is one query,
// not a million.

import { columnName, formatCell, parseCell } from "../file/cellref.js";

export const MAX_ROW = 1048576;
export const MAX_COL = 16384;

/** 1-based, inclusive. */
export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export function area(r: Rect): number {
  return (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);
}

export function intersect(a: Rect, b: Rect): Rect | undefined {
  const r1 = Math.max(a.r1, b.r1);
  const r2 = Math.min(a.r2, b.r2);
  const c1 = Math.max(a.c1, b.c1);
  const c2 = Math.min(a.c2, b.c2);
  return r1 <= r2 && c1 <= c2 ? { r1, c1, r2, c2 } : undefined;
}

export function sameRect(a: Rect, b: Rect): boolean {
  return a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;
}

export function contains(a: Rect, row: number, col: number): boolean {
  return row >= a.r1 && row <= a.r2 && col >= a.c1 && col <= a.c2;
}

/** `A1`, `A1:B2` (also `$` and the trim forms `A1:.B2`), whole columns `A:C`, whole rows `1:3`. */
export function rectOf(address: string, kind: "cell" | "area" | "cols" | "rows" | "error"): Rect | undefined {
  const clean = address.split("$").join("");
  const parts = clean.split(":").map((p) => p.split(".").join(""));
  const [a, b = a] = parts;
  if (a === undefined || b === undefined || parts.length > 2) return undefined;
  if (kind === "cell" || kind === "area") {
    const p = parseCell(a);
    const q = parseCell(b);
    if (!p || !q) return undefined;
    return { r1: Math.min(p.row, q.row), c1: Math.min(p.col, q.col), r2: Math.max(p.row, q.row), c2: Math.max(p.col, q.col) };
  }
  if (kind === "rows") {
    const x = Number(a);
    const y = Number(b);
    if (!Number.isInteger(x) || !Number.isInteger(y)) return undefined;
    return { r1: Math.min(x, y), c1: 1, r2: Math.max(x, y), c2: MAX_COL };
  }
  if (kind === "cols") {
    const p = parseCell(a + "1");
    const q = parseCell(b + "1");
    if (!p || !q) return undefined;
    return { r1: 1, c1: Math.min(p.col, q.col), r2: MAX_ROW, c2: Math.max(p.col, q.col) };
  }
  return undefined;
}

/** Whether every row and column in the address is fixed with `$` (`$B$3`, `$B$3:$D$9`, `$C:$C`, `$2:$4`). */
export function isAbsolute(address: string, kind: "cell" | "area" | "cols" | "rows" | "error"): boolean {
  const parts = address.split(":").map((p) => p.split(".").join(""));
  return parts.every((p) => {
    if (!p.startsWith("$")) return false;
    if (kind === "cols" || kind === "rows") return true;
    // `$B$3`: a second `$` between the column letters and the row digits.
    return p.indexOf("$", 1) > 1;
  });
}

/** `B3`, `B3:D9`, `C:C`, `2:4`. */
export function rectText(r: Rect): string {
  if (r.r1 === 1 && r.r2 === MAX_ROW) {
    return `${columnName(r.c1)}:${columnName(r.c2)}`;
  }
  if (r.c1 === 1 && r.c2 === MAX_COL) return `${r.r1}:${r.r2}`;
  const a = formatCell({ row: r.r1, col: r.c1 });
  return r.r1 === r.r2 && r.c1 === r.c2 ? a : `${a}:${formatCell({ row: r.r2, col: r.c2 })}`;
}

/**
 * Rectangles of one sheet, sorted by first row. A query scans the entries whose first
 * row lies between its first row minus the tallest entry's height and its last row: for
 * the short blocks of a model that is a binary search and a few comparisons.
 */
export class RectIndex<T extends { rect: Rect | undefined }> {
  private readonly items: T[];
  private readonly starts: number[];
  private readonly tallest: number;

  constructor(items: readonly T[]) {
    this.items = [...items].sort((a, b) => a.rect!.r1 - b.rect!.r1 || a.rect!.c1 - b.rect!.c1);
    this.starts = this.items.map((x) => x.rect!.r1);
    let tallest = 1;
    for (const x of this.items) tallest = Math.max(tallest, x.rect!.r2 - x.rect!.r1 + 1);
    this.tallest = tallest;
  }

  /** Entries overlapping `q`, in row-then-column order. */
  query(q: Rect): T[] {
    const out: T[] = [];
    let lo = 0;
    let hi = this.starts.length;
    const from = q.r1 - this.tallest + 1;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.starts[mid]! < from) lo = mid + 1;
      else hi = mid;
    }
    for (let k = lo; k < this.items.length && this.starts[k]! <= q.r2; k++) {
      const r = this.items[k]!.rect!;
      if (r.r2 >= q.r1 && r.c1 <= q.c2 && r.c2 >= q.c1) out.push(this.items[k]!);
    }
    return out;
  }

  /** The entry holding the cell, if any. */
  at(row: number, col: number): T | undefined {
    return this.query({ r1: row, c1: col, r2: row, c2: col })[0];
  }
}
