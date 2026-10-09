// B6 (a): the formulas of one sheet in order of appearance, row by row and left to right,
// each in display form with its saved value and the defined names it reads.
//
// One line per formula the reader would see in Excel's formula bar: a dynamic array once
// at its anchor, with the extent it spilled to when the file was saved; a legacy array
// or data table once at its top-left cell; every cell of a shared-formula group with its
// own text (the master's moved by `shiftFormula`). A formula cell lying inside another
// cell's array, spill or data table is not listed: Excel shows that cell as part of the
// other one.

import { formulaTextAt } from "../file/shared.js";
import { formatCell, parseCell } from "../file/cellref.js";
import type { CachedValue, CellFormula, Sheet, WorkbookSnapshot } from "../file/types.js";
import { walk, type Expr } from "../lang/ast.js";
import { tryParse } from "../lang/parser.js";
import { shiftFormula } from "../lang/shift.js";
import { decompile, type WorkbookLink } from "../lang/transform.js";
import { quoteSheet } from "../lang/tokens.js";
import { definitionTarget, type DefinitionTarget } from "../project/classify.js";
import { layoutToLf } from "../project/pull.js";
import { NameResolver, nameUses, type NameUse } from "../project/refs.js";
import { nameKey } from "../project/types.js";

export type FormulaViewKind = "normal" | "shared" | "array" | "dynamic-array" | "data-table";

/** Resolves a name use to a key (`Sheet!Name` or `Name`); `NameResolver` is one. */
export interface NameIndex {
  resolve(use: NameUse, homeSheet: string | undefined): string | undefined;
}

/** A defined name read by the formula, where it is written in `formula`. */
export interface FormulaViewName {
  /** As written, without qualifier. */
  id: string;
  /** Sheet written in front of `!`, if any. */
  sheet: string | undefined;
  /** The resolved key (`Sheet!Name` or `Name`); undefined when no such name exists. */
  key: string | undefined;
  start: number;
  end: number;
}

/** A cell reference in `formula` (qualifier included in the span), for navigation. */
export interface FormulaViewRef {
  /** The sheet it points at: the qualifier's, or the formula's own sheet. */
  sheet: string;
  /** As written, e.g. `$B$3`, `A1:C4`, `1:3`. */
  address: string;
  refKind: "cell" | "area" | "cols" | "rows" | "error";
  start: number;
  end: number;
}

/**
 * A defined name whose definition is exactly this line's location: the cell itself
 * (`$C$17`), the anchor's spill (`$C$6#`), or the saved extent of its array or spill as a
 * fixed range (`$C$6:$G$6`). Shown as the line's left-hand side.
 */
export interface FormulaViewLhs {
  /** The name's key (`Sheet!Name` or `Name`), as `Project.lookup` takes it. */
  key: string;
  name: string;
  /** Sheet of a sheet-scoped name; undefined for workbook scope. */
  scope: string | undefined;
  /** As the view writes it: bare for this sheet's and workbook names, `Sheet!Name` for another sheet's. */
  display: string;
  hidden: boolean;
  /** What the definition denotes: the cell, the spill (`x#`), or the saved extent. */
  target: "cell" | "spill" | "extent";
}

export interface FormulaViewLine {
  /** The sheet the formula is on. */
  sheet: string;
  /** Address of the cell (the anchor for an array, spill or data table). */
  cell: string;
  row: number;
  col: number;
  kind: FormulaViewKind;
  /** Saved extent of an array, dynamic array or data table (`C6:C11`); for a shared formula, its group's. */
  extent?: string;
  /** Size of `extent` for arrays, dynamic arrays and data tables. */
  rows?: number;
  cols?: number;
  /** For a shared formula: the cell that stores the text. */
  master?: string;
  /** For a shared formula: the number of cells in its group. */
  groupSize?: number;
  /** The formula as stored at this cell (a shared child's text already moved); "" for a data table. */
  stored: string;
  /** Display form, without `=`: no `_xlfn.`/`_xlpm.`, `x#` for `ANCHORARRAY(x)`, LF line breaks. */
  formula: string;
  /** The value saved with the file (for a spill: its anchor's); absent when `uncalculated`. */
  value?: CachedValue;
  /** `value` formatted short; undefined when none was saved. */
  valueText: string | undefined;
  /**
   * The formula has no value because Excel has not calculated it since a tool wrote it (an
   * xln build): see `uncalculated`. Its extent, for a dynamic array, is the anchor alone.
   */
  uncalculated?: true;
  /** Defined names read by the formula, in source order (duplicates kept). */
  names: FormulaViewName[];
  /** Cell references in the formula, in source order. */
  refs: FormulaViewRef[];
  /** Defined names on this line's location, sorted by `display` (empty when none). */
  lhs: FormulaViewLhs[];
  /** Why `formula` is shown as stored: it did not parse. */
  error?: string;
  /** Calculation order only: the longest chain of formulas from the inputs to this one (1 = reads inputs only). */
  level?: number;
  /** Calculation order only: the number (1-based) of the circular reference this formula is part of. */
  cycle?: number;
  /** Calculation order only: what the formula reads directly, defined names first (`DependencyGraph` labels). */
  dependsOn?: string[];
}

/** The defined names of a workbook as a `NameIndex` (keys `Sheet!Name` / `Name`). */
export function workbookNameIndex(wb: WorkbookSnapshot): NameResolver {
  return new NameResolver(
    wb.definedNames
      .filter((d) => !d.isXlPrefixed && !d.scopeInvalid)
      .map((d) => {
        const scope = d.scope.kind === "sheet" ? d.scope.name : undefined;
        return { name: d.name, scope, key: nameKey({ name: d.name, scope }) };
      }),
  );
}

interface Placed {
  lhs: Omit<FormulaViewLhs, "display" | "target">;
  t: DefinitionTarget;
  /** The sheet the definition refers to. */
  sheet: string;
}

// Parsed once per set of defined names: the views of every sheet share it.
const placedCache = new WeakMap<readonly unknown[], Placed[]>();

/** The defined names whose definition is one reference (cell, area or spill), with its sheet. */
function placedNames(wb: WorkbookSnapshot): Placed[] {
  let out = placedCache.get(wb.definedNames);
  if (out) return out;
  out = [];
  for (const d of wb.definedNames) {
    if (d.isXlPrefixed || d.isBuiltIn || d.scopeInvalid) continue;
    const t = definitionTarget(d.definition);
    if (!t) continue;
    const scope = d.scope.kind === "sheet" ? d.scope.name : undefined;
    // A sheet-scoped name written without a sheet refers to its own sheet; a workbook name
    // without one has no fixed sheet (Excel always stores the sheet, so this is rare).
    const sheet = t.sheet ?? scope;
    if (sheet === undefined) continue;
    out.push({ lhs: { key: nameKey({ name: d.name, scope }), name: d.name, scope, hidden: d.hidden }, t, sheet });
  }
  placedCache.set(wb.definedNames, out);
  return out;
}

function sameRect(a: Rect, b: Rect): boolean {
  return a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;
}

/** The names placed on `line`: its cell, its spill, or its saved extent. */
function lhsOf(line: FormulaViewLine, onSheet: readonly Placed[], sheet: string): FormulaViewLhs[] {
  const cell: Rect = { r1: line.row, c1: line.col, r2: line.row, c2: line.col };
  // A shared formula's extent is its group's: a name over the group is a block, not this cell.
  const extent = line.kind !== "shared" ? rect(line.extent) : undefined;
  const out: FormulaViewLhs[] = [];
  for (const p of onSheet) {
    const t = p.t;
    let target: FormulaViewLhs["target"] | undefined;
    if (t.spill) target = t.r1 === line.row && t.c1 === line.col ? "spill" : undefined;
    else if (sameRect(t, cell)) target = "cell";
    else if (extent && sameRect(t, extent)) target = "extent";
    if (!target) continue;
    const local = p.lhs.scope === undefined || p.lhs.scope.toLowerCase() === sheet.toLowerCase();
    out.push({ ...p.lhs, display: local ? p.lhs.name : `${quoteSheet(p.lhs.scope!)}!${p.lhs.name}`, target });
  }
  return out.sort((a, b) => {
    const x = a.display.toLowerCase();
    const y = b.display.toLowerCase();
    return x < y ? -1 : x > y ? 1 : a.display < b.display ? -1 : a.display > b.display ? 1 : 0;
  });
}

interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

function rect(range: string | undefined): Rect | undefined {
  if (!range) return undefined;
  const [a, b = a] = range.split(":");
  const p = parseCell(a!);
  const q = parseCell(b!);
  if (!p || !q) return undefined;
  return { r1: Math.min(p.row, q.row), c1: Math.min(p.col, q.col), r2: Math.max(p.row, q.row), c2: Math.max(p.col, q.col) };
}

/** Shortest readable form of a saved value: 15 significant digits at most, as Excel shows. */
export function shortValue(v: CachedValue | undefined): string | undefined {
  if (!v || v.value === undefined) return undefined;
  if (typeof v.value === "boolean") return v.value ? "TRUE" : "FALSE";
  if (typeof v.value === "number") {
    const x = v.value;
    const a = Math.abs(x);
    if (a !== 0 && (a >= 1e15 || a < 1e-9)) return x.toExponential(6).replace(/\.?0+e/, "e");
    return String(Number(x.toPrecision(a >= 1e6 ? 15 : 10)));
  }
  if (v.type === "e") return v.value;
  const s = v.value;
  const cut = s.length > 40 ? s.slice(0, 39) + "…" : s;
  return '"' + cut.split('"').join('""').split("\n").join("\\n") + '"';
}

/** Excel's `TABLE(row input, column input)` for a data table cell. */
function dataTableText(f: CellFormula): string {
  const a = f.attributes;
  const r1 = a.r1 ?? "";
  const r2 = a.r2 ?? "";
  if (a.dt2D === "1" || a.dt2D === "true") return `TABLE(${r1}, ${r2})`;
  return a.dtr === "1" || a.dtr === "true" ? `TABLE(${r1}, )` : `TABLE(, ${r1})`;
}

/** Display form, its AST and its parse error, once per distinct stored text. */
type Parsed = { formula: string; body: Expr | undefined; error: string | undefined };

function displayOf(stored: string, cache: Map<string, Parsed>, links: readonly WorkbookLink[] = []): Parsed {
  let p = cache.get(stored);
  if (p) return p;
  let formula = layoutToLf(stored);
  let error: string | undefined;
  try {
    formula = layoutToLf(decompile(stored, { links }));
  } catch (e) {
    error = e instanceof Error ? e.message.split("\n")[0] : String(e);
  }
  const body = error === undefined ? tryParse(formula).formula?.body : undefined;
  p = { formula, body, error };
  cache.set(stored, p);
  return p;
}

/** The syntax tree of a formula `sheetFormulaView` parsed into `cache` (undefined if it did not parse). */
export function cachedBody(cache: Map<string, unknown>, stored: string): Expr | undefined {
  return (cache.get(stored) as Parsed | undefined)?.body;
}

function external(sheetQual: { book?: string; sheet?: string; sheet2?: string } | undefined): boolean {
  return sheetQual !== undefined && (sheetQual.book !== undefined || sheetQual.sheet === undefined || sheetQual.sheet2 !== undefined);
}

function refsOf(body: Expr, home: string): FormulaViewRef[] {
  const out: FormulaViewRef[] = [];
  walk(body, (n) => {
    if (n.kind !== "ref" || external(n.qual)) return;
    out.push({ sheet: n.qual?.sheet ?? home, address: n.address, refKind: n.refKind, start: n.span.start, end: n.span.end });
  });
  return out.sort((a, b) => a.start - b.start);
}

export interface FormulaViewOptions {
  /** Shared across calls to parse each distinct formula once (per workbook). */
  cache?: Map<string, unknown>;
}

/**
 * The formulas of `sheetName` in order of appearance. `names` resolves the defined names
 * each formula reads, with the sheet as home (`workbookNameIndex(wb)` for the file's own
 * names, or an index over a project's live definitions). Throws if there is no such sheet.
 */
export function sheetFormulaView(wb: WorkbookSnapshot, sheetName: string, names: NameIndex = workbookNameIndex(wb), opts: FormulaViewOptions = {}): FormulaViewLine[] {
  const sheet = findSheet(wb, sheetName);
  if (!sheet) throw new Error(`no sheet '${sheetName}' in the workbook`);
  const cache = (opts.cache ?? new Map()) as Map<string, Parsed>;

  const groups = new Map(sheet.sharedFormulas.map((g) => [g.si, g]));
  const cells: { f: CellFormula; row: number; col: number }[] = [];
  const covers: { anchor: string; r: Rect }[] = [];
  for (const f of sheet.formulas) {
    const a = parseCell(f.cell);
    if (!a) continue;
    cells.push({ f, row: a.row, col: a.col });
    if (f.kind === "array" || f.kind === "dynamic-array" || f.kind === "data-table") {
      const r = rect(f.range ?? sheet.spills.find((s) => s.anchor === f.cell)?.extent);
      if (r && (r.r2 > r.r1 || r.c2 > r.c1)) covers.push({ anchor: f.cell, r });
    }
  }
  const inside = (c: { f: CellFormula; row: number; col: number }): boolean =>
    covers.some(({ anchor, r }) => anchor !== c.f.cell && c.row >= r.r1 && c.row <= r.r2 && c.col >= r.c1 && c.col <= r.c2);
  cells.sort((x, y) => x.row - y.row || x.col - y.col);
  const here = sheet.name.toLowerCase();
  const onSheet = placedNames(wb).filter((p) => p.sheet.toLowerCase() === here);

  const out: FormulaViewLine[] = [];
  for (const c of cells) {
    if (inside(c)) continue;
    const f = c.f;
    const line: FormulaViewLine = {
      sheet: sheet.name,
      cell: formatCell({ row: c.row, col: c.col }),
      row: c.row,
      col: c.col,
      kind: f.kind === "shared-master" || f.kind === "shared-child" ? "shared" : f.kind,
      stored: "",
      formula: "",
      valueText: shortValue(f.value),
      names: [],
      refs: [],
      lhs: [],
    };
    if (uncalculated(wb, f)) line.uncalculated = true;
    else line.value = f.value;
    if (f.kind === "array" || f.kind === "dynamic-array" || f.kind === "data-table") {
      const extent = f.range ?? sheet.spills.find((s) => s.anchor === f.cell)?.extent ?? f.cell;
      const r = rect(extent);
      line.extent = extent;
      if (r) {
        line.rows = r.r2 - r.r1 + 1;
        line.cols = r.c2 - r.c1 + 1;
      }
    }
    if (line.kind === "shared" && f.si !== undefined) {
      const g = groups.get(f.si);
      line.master = g?.master ?? f.master;
      if (g?.range) line.extent = g.range;
      if (g) line.groupSize = g.cells.length;
    }
    if (onSheet.length > 0) line.lhs = lhsOf(line, onSheet, sheet.name);
    if (f.kind === "data-table") {
      line.formula = dataTableText(f);
      out.push(line);
      continue;
    }
    const stored = formulaTextAt(sheet, f, shiftFormula);
    if (stored === undefined) {
      line.error = `shared formula ${f.si ?? "?"} has no master in the file`;
      out.push(line);
      continue;
    }
    line.stored = stored;
    if (stored.trim() === "") {
      out.push(line);
      continue;
    }
    const p = displayOf(stored, cache, wb.externalLinks ?? []);
    line.formula = p.formula;
    if (p.error !== undefined) line.error = p.error;
    if (p.body) {
      line.names = nameUses(p.body).map((u) => ({ id: u.id, sheet: u.sheet, key: names.resolve(u, sheet.name), start: u.span.start, end: u.span.end }));
      line.refs = refsOf(p.body, sheet.name);
    }
    out.push(line);
  }
  return out;
}

/**
 * A formula Excel has not calculated since a tool wrote it. The rule: the cell has a formula
 * but no saved value at all (no `<v>`, no inline string: Excel saves one for every formula
 * cell, `<v></v>` for an empty text), and the workbook still carries `fullCalcOnLoad="1"`,
 * which the build sets and Excel drops when it saves (probe F06). Both together mean the
 * file was written after Excel last saved it, and the value is missing, not empty. A data
 * table's cells are Excel's, never written by a build.
 */
function uncalculated(wb: WorkbookSnapshot, f: CellFormula): boolean {
  return wb.fullCalcOnLoad === true && f.kind !== "data-table" && f.value.raw === undefined;
}

function findSheet(wb: WorkbookSnapshot, name: string): Sheet | undefined {
  return wb.sheets.find((s) => s.name === name) ?? wb.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase());
}
