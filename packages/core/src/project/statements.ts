// M3b (D7, D8): the cell statements of a sheet, in sheet order (row by row, left to right,
// as `xln formulas` lists them). Excel owns the layout; the source owns the formulas of
// the cells Excel has:
//
//   named    a formula cell with a defined name on it (`Name @C6 = …;`): the name's
//            definition is exactly that cell, `$C$6` or its spill `$C$6#`
//   slot     a defined name on an empty cell (`Name @C6 = ;`)
//   unnamed  any other formula cell (`@C5 = …;`); a block holding one formula filled
//            across (equal after shifting relative references, `shiftFormula`) is one
//            statement over its range (`@B40:G40 = …;`), written as the top-left formula
//
// A spill anchor is one cell; the cells it spilled into (and those of a legacy array or a
// data table) are not statements. A name counts only when it is scoped to the sheet or to
// the workbook; when several sit on one cell, one is the statement (sheet scope first, then
// the spill form, then by name) and the others stay ordinary names.

import { formatCell, parseCell } from "../file/cellref.js";
import { formulaTextAt } from "../file/shared.js";
import type { CellFormula, DefinedName, FormulaKind, Sheet, WorkbookSnapshot } from "../file/types.js";
import { shiftFormula } from "../lang/shift.js";
import { decompile, type WorkbookLink } from "../lang/transform.js";
import { definitionTarget } from "./classify.js";
import { layoutToLf } from "./layout.js";
import { normalizeDefinition } from "./lockfile.js";
import { compareNames } from "./modules.js";
import { nameKey } from "./types.js";

export type CellStatementKind = "named" | "slot" | "unnamed";

export interface CellStatement {
  sheet: string;
  /** `C6`, or `B40:G40` for a block. */
  range: string;
  /** Top-left cell, 1-based. */
  row: number;
  col: number;
  rows: number;
  cols: number;
  kind: CellStatementKind;
  /** For `named` and `slot`. */
  name?: string;
  /** The name's sheet scope; undefined for workbook scope. */
  scope?: string;
  /** `Sheet!Name` or `Name`. */
  key?: string;
  /** The name is defined as the anchor's spill (`$C$6#`). */
  onSpill?: boolean;
  /** The top-left cell's formula as stored; undefined for a slot. */
  stored?: string;
  /** Display form (LF line breaks); "" for a slot. */
  display: string;
  /** The anchor spilled beyond itself when the file was saved. */
  spills?: boolean;
  /** Why `display` is the stored text: it did not parse. */
  error?: string;
}

/** Addresses of the cells holding a stored value, per sheet name (`readCellValues`). */
export type ValueCells = ReadonlyMap<string, ReadonlySet<string>>;

/** A formula cell as a statement sees it: the cell Excel shows a formula for. */
export interface FormulaCell {
  row: number;
  col: number;
  kind: FormulaKind;
  /** Its own formula (a shared child's moved from the master); undefined when unreadable. */
  stored: string | undefined;
  /** Size of an array, dynamic array or data table (1×1 otherwise). */
  rows: number;
  cols: number;
}

export interface SheetFormulaCells {
  /** Address → formula cell, for the cells not inside another cell's array or spill. */
  cells: Map<string, FormulaCell>;
  /** Inside another cell's array, spill or data table (a ghost cell). */
  covered(row: number, col: number): boolean;
}

interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

function rectOf(range: string | undefined): Rect | undefined {
  if (!range) return undefined;
  const [a, b = a] = range.split(":");
  const p = parseCell(a!);
  const q = parseCell(b!);
  if (!p || !q) return undefined;
  return { r1: Math.min(p.row, q.row), c1: Math.min(p.col, q.col), r2: Math.max(p.row, q.row), c2: Math.max(p.col, q.col) };
}

/** `C6`, or `B40:G40`. */
export function rangeText(r: Rect): string {
  const tl = formatCell({ row: r.r1, col: r.c1 });
  return r.r1 === r.r2 && r.c1 === r.c2 ? tl : `${tl}:${formatCell({ row: r.r2, col: r.c2 })}`;
}

/** A range as `{r1, c1, r2, c2}`, `$` and case ignored; undefined when it is not one. */
export function parseCellRange(range: string): Rect | undefined {
  return rectOf(range.split("$").join(""));
}

/** The formula cells of a sheet, each with its own text, ghost cells set apart. */
export function sheetFormulaCells(sheet: Sheet): SheetFormulaCells {
  const covers: { anchor: string; r: Rect }[] = [];
  const all: { f: CellFormula; row: number; col: number }[] = [];
  for (const f of sheet.formulas) {
    const a = parseCell(f.cell);
    if (!a) continue;
    all.push({ f, row: a.row, col: a.col });
    if (f.kind === "array" || f.kind === "dynamic-array" || f.kind === "data-table") {
      const r = rectOf(f.range ?? sheet.spills.find((s) => s.anchor === f.cell)?.extent);
      if (r && (r.r2 > r.r1 || r.c2 > r.c1)) covers.push({ anchor: f.cell, r });
    }
  }
  const inside = (cell: string | undefined, row: number, col: number): boolean =>
    covers.some(({ anchor, r }) => anchor !== cell && row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2);
  const cells = new Map<string, FormulaCell>();
  for (const { f, row, col } of all) {
    if (inside(f.cell, row, col)) continue;
    let rows = 1;
    let cols = 1;
    if (f.kind === "array" || f.kind === "dynamic-array" || f.kind === "data-table") {
      const r = rectOf(f.range ?? sheet.spills.find((s) => s.anchor === f.cell)?.extent);
      if (r) {
        rows = r.r2 - r.r1 + 1;
        cols = r.c2 - r.c1 + 1;
      }
    }
    const stored = f.kind === "data-table" ? undefined : formulaTextAt(sheet, f, shiftFormula);
    cells.set(formatCell({ row, col }), { row, col, kind: f.kind, stored, rows, cols });
  }
  return { cells, covered: (row, col) => inside(undefined, row, col) };
}

/** Comparable form of a formula: its tokens without layout (as the lockfile hashes it). */
function norm(text: string): string {
  return normalizeDefinition(text);
}

/** Whether `text` at a cell `dr` rows down and `dc` columns right of `base` is `base`'s formula filled there. */
export function sameFilled(base: string, dr: number, dc: number, text: string): boolean {
  return norm(shiftFormula(base, dr, dc)) === norm(text);
}

/** A defined name that may be a cell statement: one cell of a sheet, scoped to it or to the workbook. */
interface Candidate {
  d: DefinedName;
  name: string;
  scope: string | undefined;
  sheet: string;
  row: number;
  col: number;
  spill: boolean;
}

function candidates(wb: WorkbookSnapshot): Map<string, Candidate[]> {
  const out = new Map<string, Candidate[]>();
  const sheetByLower = new Map(wb.sheets.map((s) => [s.name.toLowerCase(), s.name]));
  for (const d of wb.definedNames) {
    if (d.isXlPrefixed || d.isBuiltIn || d.scopeInvalid) continue;
    const t = definitionTarget(d.definition);
    if (!t || t.r1 !== t.r2 || t.c1 !== t.c2) continue;
    const scope = d.scope.kind === "sheet" ? d.scope.name : undefined;
    const sheet = sheetByLower.get((t.sheet ?? scope ?? "").toLowerCase());
    if (sheet === undefined) continue;
    if (scope !== undefined && scope.toLowerCase() !== sheet.toLowerCase()) continue;
    const k = `${sheet.toLowerCase()}!${t.r1},${t.c1}`;
    const l = out.get(k) ?? [];
    l.push({ d, name: d.name, scope, sheet, row: t.r1, col: t.c1, spill: t.spill });
    out.set(k, l);
  }
  for (const l of out.values()) {
    l.sort((a, b) => (a.scope === undefined ? 1 : 0) - (b.scope === undefined ? 1 : 0) || (a.spill ? 0 : 1) - (b.spill ? 0 : 1) || compareNames(a.name, b.name));
  }
  return out;
}

function displayOf(stored: string, cache: Map<string, { display: string; error?: string }>, links: readonly WorkbookLink[]): { display: string; error?: string } {
  let p = cache.get(stored);
  if (p) return p;
  try {
    p = { display: layoutToLf(decompile(stored, { links })) };
  } catch (e) {
    p = { display: layoutToLf(stored), error: e instanceof Error ? e.message.split("\n")[0]! : String(e) };
  }
  cache.set(stored, p);
  return p;
}

export interface CellStatementOptions {
  /** Cells with a stored value: a name on such a cell is an input, not a slot. Without it, every cell without a formula counts as empty. */
  values?: ValueCells;
  /** Shared across sheets to decompile each distinct formula once. */
  cache?: Map<string, { display: string; error?: string }>;
}

/** The cell statements of every worksheet, by sheet name, each list in sheet order. */
export function cellStatements(wb: WorkbookSnapshot, opts: CellStatementOptions = {}): Map<string, CellStatement[]> {
  const cand = candidates(wb);
  const cache = opts.cache ?? new Map();
  const out = new Map<string, CellStatement[]>();
  for (const sheet of wb.sheets) {
    if (sheet.kind !== "worksheet") continue;
    out.set(sheet.name, sheetStatements(sheet, cand, opts.values?.get(sheet.name), cache, wb.externalLinks ?? []));
  }
  return out;
}

function sheetStatements(sheet: Sheet, cand: Map<string, Candidate[]>, values: ReadonlySet<string> | undefined, cache: Map<string, { display: string; error?: string }>, links: readonly WorkbookLink[]): CellStatement[] {
  const fc = sheetFormulaCells(sheet);
  const here = sheet.name.toLowerCase();
  const stmts: CellStatement[] = [];
  const named = (c: Candidate): Pick<CellStatement, "name" | "scope" | "key" | "onSpill"> => {
    const s: Pick<CellStatement, "name" | "scope" | "key" | "onSpill"> = { name: c.name, key: nameKey(c) };
    if (c.scope !== undefined) s.scope = c.scope;
    if (c.spill) s.onSpill = true;
    return s;
  };
  const withFormula = (s: CellStatement, f: FormulaCell): CellStatement => {
    const p = displayOf(f.stored!, cache, links);
    s.stored = f.stored!;
    s.display = p.display;
    if (p.error !== undefined) s.error = p.error;
    return s;
  };

  // Unnamed cells that may join a block: one formula of their own, not a spill or an array.
  const free = new Map<string, FormulaCell>();
  for (const [addr, f] of fc.cells) {
    if (f.stored === undefined || f.stored.trim() === "") continue;
    const c = cand.get(`${here}!${f.row},${f.col}`)?.[0];
    if (c) {
      stmts.push(withFormula({ sheet: sheet.name, range: addr, row: f.row, col: f.col, rows: 1, cols: 1, kind: "named", ...named(c), display: "" }, f));
      if (f.rows * f.cols > 1) stmts[stmts.length - 1]!.spills = true;
      continue;
    }
    if (f.kind === "data-table") continue;
    if (f.rows * f.cols > 1) {
      stmts.push(withFormula({ sheet: sheet.name, range: addr, row: f.row, col: f.col, rows: 1, cols: 1, kind: "unnamed", spills: true, display: "" }, f));
      continue;
    }
    free.set(`${f.row},${f.col}`, f);
  }

  // Slots: names on empty cells (no formula, no value, not a ghost of a spill).
  for (const [k, l] of cand) {
    if (!k.startsWith(here + "!")) continue;
    const c = l[0]!;
    const addr = formatCell({ row: c.row, col: c.col });
    const f = fc.cells.get(addr);
    if (f && f.stored !== undefined && f.stored.trim() !== "") continue;
    if (f && f.kind === "data-table") continue;
    if (fc.covered(c.row, c.col) || values?.has(addr)) continue;
    stmts.push({ sheet: sheet.name, range: addr, row: c.row, col: c.col, rows: 1, cols: 1, kind: "slot", ...named(c), display: "" });
  }

  // Blocks: row runs of one formula filled right, then runs over the same columns on
  // consecutive rows filled down, into rectangles.
  const order = [...free.values()].sort((a, b) => a.row - b.row || a.col - b.col);
  const used = new Set<string>();
  const runs: { row: number; c1: number; c2: number; f: FormulaCell }[] = [];
  for (const f of order) {
    const k = `${f.row},${f.col}`;
    if (used.has(k)) continue;
    used.add(k);
    let c2 = f.col;
    for (;;) {
      const n = free.get(`${f.row},${c2 + 1}`);
      if (!n || used.has(`${f.row},${c2 + 1}`) || !sameFilled(f.stored!, 0, c2 + 1 - f.col, n.stored!)) break;
      c2++;
      used.add(`${f.row},${c2}`);
    }
    runs.push({ row: f.row, c1: f.col, c2, f });
  }
  const open = new Map<string, { r1: number; r2: number; c1: number; c2: number; f: FormulaCell }>();
  const rects: { r1: number; r2: number; c1: number; c2: number; f: FormulaCell }[] = [];
  for (const run of runs) {
    const key = `${run.c1}:${run.c2}`;
    const o = open.get(key);
    if (o && o.r2 === run.row - 1 && sameFilled(o.f.stored!, run.row - o.r1, 0, run.f.stored!)) {
      o.r2 = run.row;
      continue;
    }
    const r = { r1: run.row, r2: run.row, c1: run.c1, c2: run.c2, f: run.f };
    open.set(key, r);
    rects.push(r);
  }
  for (const r of rects) {
    stmts.push(withFormula({ sheet: sheet.name, range: rangeText(r), row: r.r1, col: r.c1, rows: r.r2 - r.r1 + 1, cols: r.c2 - r.c1 + 1, kind: "unnamed", display: "" }, r.f));
  }
  return stmts.sort((a, b) => a.row - b.row || a.col - b.col);
}
