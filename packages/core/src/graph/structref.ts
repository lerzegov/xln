// Structured references (`Tbl[Col]`, `Tbl[[#Headers],[A]:[B]]`, `[@Col]`) resolved to the
// rectangle of the Table they denote. The text between the outer brackets is read with
// Excel's bracket nesting and `'` escapes, in stored form (`[#This Row],[Col]`) and in
// display form (`@Col`, `@[Col]`).

import { parseCell } from "../file/cellref.js";
import type { Table } from "../file/types.js";
import type { Rect } from "./rect.js";

export type StructArea = "#all" | "#data" | "#headers" | "#totals" | "#this row";

export interface StructSpec {
  /** Special items, lower-case; empty means the data rows. */
  areas: StructArea[];
  /** Columns named, as written (escapes removed): one, or the two ends of `[A]:[B]`. */
  columns: string[];
}

const AREAS = new Set<string>(["#all", "#data", "#headers", "#totals", "#this row"]);

function unescape(s: string): string {
  let out = "";
  for (let k = 0; k < s.length; k++) {
    if (s[k] === "'" && k + 1 < s.length) k++;
    out += s[k];
  }
  return out;
}

/** The items of a structured reference's inner text; undefined when it cannot be read. */
export function parseStructInner(inner: string): StructSpec | undefined {
  let s = inner.trim();
  const spec: StructSpec = { areas: [], columns: [] };
  if (s.startsWith("@")) {
    spec.areas.push("#this row");
    s = s.slice(1).trim();
  }
  const item = (text: string): boolean => {
    const t = unescape(text).trim();
    if (t === "") return true;
    const low = t.toLowerCase();
    if (t.startsWith("#")) {
      if (!AREAS.has(low)) return false;
      spec.areas.push(low as StructArea);
    } else spec.columns.push(t);
    return true;
  };
  if (!s.includes("[")) return item(s) ? spec : undefined;
  // `[a],[b]:[c]`: groups in brackets, separated by `,` or `:` (and spaces).
  let k = 0;
  while (k < s.length) {
    const c = s[k]!;
    if (c === " " || c === "," || c === ":") {
      k++;
      continue;
    }
    if (c !== "[") return undefined;
    let q = k + 1;
    let depth = 1;
    while (q < s.length) {
      const d = s[q]!;
      if (d === "'") {
        q += 2;
        continue;
      }
      if (d === "[") depth++;
      else if (d === "]" && --depth === 0) break;
      q++;
    }
    if (q >= s.length) return undefined;
    if (!item(s.slice(k + 1, q))) return undefined;
    k = q + 1;
  }
  return spec;
}

export type StructResult = { rect: Rect; label: string } | { error: string };

/**
 * The rectangle `spec` denotes in `table`. `row` is the row of the formula, for `#This Row`
 * (undefined when there is none: a defined name).
 */
export function structRect(table: Table, spec: StructSpec, row: number | undefined): StructResult {
  const [a, b = a] = table.ref.split(":");
  const p = parseCell(a!);
  const q = parseCell(b!);
  if (!p || !q) return { error: `Table ${table.displayName} has no readable range (${table.ref})` };
  const top = Math.min(p.row, q.row);
  const bottom = Math.max(p.row, q.row);
  const left = Math.min(p.col, q.col);
  const dataTop = top + table.headerRowCount;
  const dataBottom = bottom - table.totalsRowCount;

  let c1 = left;
  let c2 = Math.max(p.col, q.col);
  if (spec.columns.length > 0) {
    const idx = spec.columns.map((name) => table.columns.findIndex((c) => c.name.toLowerCase() === name.toLowerCase()));
    const missing = idx.findIndex((i) => i < 0);
    if (missing >= 0) return { error: `Table ${table.displayName} has no column '${spec.columns[missing]}'` };
    c1 = left + Math.min(...idx);
    c2 = left + Math.max(...idx);
  }

  let r1 = Infinity;
  let r2 = -Infinity;
  const span = (x: number, y: number): void => {
    r1 = Math.min(r1, x);
    r2 = Math.max(r2, y);
  };
  const areas = spec.areas.length > 0 ? spec.areas : ["#data"];
  for (const area of areas) {
    switch (area) {
      case "#all":
        span(top, bottom);
        break;
      case "#data":
        span(dataTop, dataBottom);
        break;
      case "#headers":
        if (table.headerRowCount === 0) return { error: `Table ${table.displayName} has no header row` };
        span(top, dataTop - 1);
        break;
      case "#totals":
        if (table.totalsRowCount === 0) return { error: `Table ${table.displayName} has no totals row` };
        span(dataBottom + 1, bottom);
        break;
      case "#this row":
        if (row === undefined) return { error: `this-row reference outside a cell` };
        if (row < dataTop || row > dataBottom) return { error: `this-row reference from row ${row}, outside the data of ${table.displayName}` };
        span(row, row);
        break;
    }
  }
  return { rect: { r1, c1, r2, c2 }, label: structLabel(table.displayName, spec) };
}

const AREA_TEXT: Record<StructArea, string> = { "#all": "#All", "#data": "#Data", "#headers": "#Headers", "#totals": "#Totals", "#this row": "#This Row" };

/** `Tbl[Col]`, `Tbl[#All]`, `Tbl[[#Headers],[A]:[B]]`: the reference written the way Excel shows it. */
export function structLabel(table: string, spec: StructSpec): string {
  const items = spec.areas.filter((a) => a !== "#data" || spec.areas.length > 1).map((a) => `[${AREA_TEXT[a]}]`);
  const [a, b] = spec.columns;
  if (a !== undefined) items.push(b !== undefined && b.toLowerCase() !== a.toLowerCase() ? `[${a}]:[${b}]` : `[${a}]`);
  if (items.length === 0) return `${table}[#Data]`;
  if (items.length === 1) return `${table}${items[0]}`;
  return `${table}[${items.join(",")}]`;
}
