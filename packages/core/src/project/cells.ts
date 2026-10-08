// Cell lists in the manifest: a name used by a whole block of formulas would otherwise
// list every cell. Cells are merged into row runs, then runs with the same columns on
// consecutive rows into rectangles. The result covers exactly the input cells.

import { formatCell, parseCell } from "../file/cellref.js";

export function compressCells(cells: Iterable<string>): string[] {
  const byRow = new Map<number, Set<number>>();
  const odd: string[] = [];
  for (const c of cells) {
    const a = parseCell(c);
    if (!a) {
      odd.push(c);
      continue;
    }
    let s = byRow.get(a.row);
    if (!s) byRow.set(a.row, (s = new Set()));
    s.add(a.col);
  }
  // row runs: "c1:c2" → rows having exactly that run
  const runs = new Map<string, number[]>();
  for (const row of [...byRow.keys()].sort((x, y) => x - y)) {
    const cols = [...byRow.get(row)!].sort((x, y) => x - y);
    for (let k = 0; k < cols.length; ) {
      let e = k;
      while (e + 1 < cols.length && cols[e + 1] === cols[e]! + 1) e++;
      const key = `${cols[k]}:${cols[e]}`;
      let rows = runs.get(key);
      if (!rows) runs.set(key, (rows = []));
      rows.push(row);
      k = e + 1;
    }
  }
  const rects: { r1: number; r2: number; c1: number; c2: number }[] = [];
  for (const [key, rows] of runs) {
    const [c1, c2] = key.split(":").map(Number) as [number, number];
    for (let k = 0; k < rows.length; ) {
      let e = k;
      while (e + 1 < rows.length && rows[e + 1] === rows[e]! + 1) e++;
      rects.push({ r1: rows[k]!, r2: rows[e]!, c1, c2 });
      k = e + 1;
    }
  }
  rects.sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1);
  const out = rects.map(({ r1, r2, c1, c2 }) => {
    const tl = formatCell({ row: r1, col: c1 });
    return r1 === r2 && c1 === c2 ? tl : `${tl}:${formatCell({ row: r2, col: c2 })}`;
  });
  return out.concat([...new Set(odd)].sort());
}

/** Rows and columns of an `A1:B2` extent (or a single cell); undefined if unreadable. */
export function extentSize(extent: string): { rows: number; cols: number } | undefined {
  const [a, b = a] = extent.split(":");
  const p = parseCell(a!);
  const q = parseCell(b!);
  if (!p || !q) return undefined;
  return { rows: Math.abs(q.row - p.row) + 1, cols: Math.abs(q.col - p.col) + 1 };
}
