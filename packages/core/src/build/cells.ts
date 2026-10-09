// E6, the cell-level three-way check (M3b): for each cell statement of the source, its
// formula in the source, in the lockfile (what the last pull or build saw) and in the
// workbook now.
//
//   changed in source only   → set-cell-formula (or clear-cell-formula for `= ;`)
//   changed in Excel only    → kept, reported (pull again)
//   changed in both          → a conflict, never merged (unless both sides agree)
//
// Identity: a named cell (and a slot) follows its name, so a cell moved in Excel (rows
// inserted above it) is written where the name now points; an unnamed cell is its
// address, so a moved unnamed cell shows as one changed (or emptied) and one added.
// The address in the source is read-only: it must be the lockfile's. Its `#` is not
// (author's decision, 2026-10-05): it says what the name of a named statement covers, the
// spill (`'S'!$C$6#`) or the cell alone, and adding or removing it is a change of the name,
// three-way like the formula (lockfile, workbook, source). Nothing moves a name to `#` on
// its own: a statement without `#` on a formula that spills gets a warning and a quick fix.
// A project pulled before the `#` was written (lockfile format 3 or older) reads a missing
// `#` as the spill when the name is on it in the lockfile and in the workbook.
//
// Formulas are compared as the lockfile hashes them (tokens without layout). A source
// text that is exactly the workbook's formula decompiled counts as no edit even when its
// compiled form differs (a stored formula lacking a prefix): cells are never repaired.

import { columnName, formatCell } from "../file/cellref.js";
import type { DefinedName, WorkbookSnapshot } from "../file/types.js";
import { equalModuloWhitespace, hasLayoutBreaks, oneLine } from "../lang/format.js";
import { quoteSheet } from "../lang/tokens.js";
import { compileWithDiagnostics, decompile, type WorkbookLink } from "../lang/transform.js";
import { definitionTarget } from "../project/classify.js";
import { layoutToLf } from "../project/layout.js";
import { definitionHash, definitionHashLike, definitionHashV2, explicitSpill, isV2Hash, type LockCell, type Lockfile } from "../project/lockfile.js";
import { formatCellAddress } from "../project/module.js";
import { cellStatements, parseCellRange, rangeText, sameFilled, sheetFormulaCells, type SheetFormulaCells, type ValueCells } from "../project/statements.js";
import { ONE_LINE_AS_WORKBOOK, ONE_LINE_NEW, type Change, type Scope, type SetCellFormula } from "./changes.js";
import type { Conflict, ExcelChange, NameState } from "./plan.js";
import type { SourceCell, SourceProblem } from "./source.js";

/** A cell statement the built workbook must match (read-back of cells, M3b-2). */
export interface CellState {
  sheet: string;
  /** Where the statement is now (a named cell moved in Excel: its new place). */
  range: string;
  /** The name's key, for a named cell or a slot. */
  name?: string;
  /** The source's formula, display form; "" for an empty cell. */
  display: string;
  /** The top-left formula as stored (compiled, or the workbook's when unchanged); null for an empty cell. */
  stored: string | null;
}

/** What the cell planner needs to know about the name of a named statement. */
export interface NamedCell {
  /** The name's key in the source after the build (`BS!Revenue`, `Revenue`). */
  key: string;
  name: string;
  scope: Scope;
  /** The name in the workbook now, through its identity; undefined when Excel deleted it. */
  current: DefinedName | undefined;
  /** Its key in the lockfile (where it was at the last pull), if it was there. */
  lockKey: string | undefined;
}

export interface CellPlanInput {
  wb: WorkbookSnapshot;
  lock: Lockfile;
  values: ValueCells | undefined;
  cells: readonly SourceCell[];
  /** The name of a named statement; undefined when the names part refused it (unknown sheet, duplicate). */
  named: (c: SourceCell) => NamedCell | undefined;
  /** Every defined name (compile context). */
  names: readonly string[];
  /**
   * Stretch G: the workbook's formula on `sheet` with the build's renamed names rewritten
   * (undefined when no rename reaches it). A statement whose only edit is that rewrite is
   * left to the rename: its cells keep their formulas' form, only the token changes.
   */
  renamed?: (stored: string, sheet: string) => string | undefined;
}

export interface CellPlan {
  changes: Change[];
  conflicts: Conflict[];
  excelChanges: ExcelChange[];
  problems: SourceProblem[];
  /** The `#` of each named statement whose name is in the workbook: name key (lower case) → what it means for the name. */
  spill: Map<string, NameSpill>;
  /** Keys of the named statements without `#` read as on the spill (a project pulled before the `#` was written). */
  implicitSpill: string[];
  inSync: CellState[];
  unchanged: number;
  /** Keys (lower case, the statement's and its lock key) of statements whose only source edit is a rename's token: rewritten by the rename (stretch G), not set. */
  renamedOnly: Set<string>;
  /** Lockfile cell key → the format-3 hash of a format-1/2 entry the source still matches (see `upgradeLockfile`). */
  lockUpgrades: Map<string, string>;
  /** Lock keys (lower case) of the statements whose source differs from the lockfile: edited, new, removed. */
  sourceEdits: Set<string>;
}

/** A named statement's `#` against its name (lockfile, workbook). */
export interface NameSpill {
  /** The definition the source asks for: the workbook's when its `#` agrees, else the current cell or its spill, stored. */
  stored: string;
  /** The source added or removed the `#` since the lockfile. */
  edited: boolean;
  /** `edited`, and the build writes it (not while the cell stays empty). */
  srcChanged: boolean;
  /** Excel put the name on the spill or took it off (or made it something else) since the lockfile. */
  xlChanged: boolean;
}

type Rect = { r1: number; c1: number; r2: number; c2: number };

/** `'S'!$C$6`, stored. */
function cellRef(sheet: string, row: number, col: number): string {
  return `${quoteSheet(sheet)}!$${columnName(col)}$${row}`;
}

function spillRef(ref: string): string {
  return `_xlfn.ANCHORARRAY(${ref})`;
}

/** The workbook's side of a statement. */
type XlState =
  | { kind: "formula"; stored: string }
  | { kind: "empty" }
  /** The cell holds a value (typed in Excel). */
  | { kind: "value" }
  /** Inside another cell's spill or array. */
  | { kind: "ghost" }
  /** A block whose cells no longer hold one formula filled across. */
  | { kind: "mixed"; stored: string }
  /** The name no longer names one cell. */
  | { kind: "not-a-cell"; definition: string };

/** The workbook's side as a lock hash, computed the way `like` (the lockfile's hash) was. */
function xlHash(s: XlState, like: string | null): string | null {
  switch (s.kind) {
    case "formula":
      return definitionHashLike(like, s.stored);
    case "empty":
      return null;
    default:
      return `!${s.kind}`;
  }
}

function lower(s: string): string {
  return s.toLowerCase();
}

function displayOf(stored: string, links: readonly WorkbookLink[]): string {
  try {
    return layoutToLf(decompile(stored, { links }));
  } catch {
    return layoutToLf(stored);
  }
}

function xlDisplay(s: XlState, links: readonly WorkbookLink[]): string {
  switch (s.kind) {
    case "formula":
    case "mixed":
      return displayOf(s.stored, links);
    case "empty":
      return "";
    case "value":
      return "// holds a value typed in Excel";
    case "ghost":
      return "// inside another cell's spill";
    case "not-a-cell":
      return `// the name is now ${s.definition}`;
  }
}

export function planCells(input: CellPlanInput): CellPlan {
  const links = input.wb.externalLinks ?? [];
  const { wb, lock, values } = input;
  const out: CellPlan = { changes: [], conflicts: [], excelChanges: [], problems: [], spill: new Map(), implicitSpill: [], inSync: [], unchanged: 0, renamedOnly: new Set(), lockUpgrades: new Map(), sourceEdits: new Set() };
  const explicit = explicitSpill(lock);
  const v2 = lock.cells !== undefined;
  const lockCells = new Map<string, { key: string; e: LockCell }>();
  for (const [key, e] of Object.entries(lock.cells ?? {})) lockCells.set(lower(key), { key, e });
  const sheetByLower = new Map(wb.sheets.map((s) => [lower(s.name), s.name]));
  const fcCache = new Map<string, SheetFormulaCells>();
  const fcOf = (sheet: string): SheetFormulaCells => {
    let fc = fcCache.get(sheet);
    if (!fc) fcCache.set(sheet, (fc = sheetFormulaCells(wb.sheets.find((s) => s.name === sheet)!)));
    return fc;
  };

  const cellState = (sheet: string, r: Rect): XlState => {
    const fc = fcOf(sheet);
    const tl = fc.cells.get(formatCell({ row: r.r1, col: r.c1 }));
    const one = (row: number, col: number, f = fc.cells.get(formatCell({ row, col }))): XlState => {
      if (f && f.stored !== undefined && f.stored.trim() !== "") return { kind: "formula", stored: f.stored };
      if (f?.kind === "data-table" || fc.covered(row, col)) return { kind: "ghost" };
      return values?.get(sheet)?.has(formatCell({ row, col })) ? { kind: "value" } : { kind: "empty" };
    };
    const first = one(r.r1, r.c1, tl);
    if (r.r1 === r.r2 && r.c1 === r.c2) return first;
    if (first.kind !== "formula" || (tl && tl.rows * tl.cols > 1)) return first.kind === "formula" ? { kind: "mixed", stored: first.stored } : first;
    for (let row = r.r1; row <= r.r2; row++) {
      for (let col = r.c1; col <= r.c2; col++) {
        if (row === r.r1 && col === r.c1) continue;
        const f = fc.cells.get(formatCell({ row, col }));
        if (!f || f.stored === undefined || f.rows * f.cols > 1 || !sameFilled(first.stored, row - r.r1, col - r.c1, f.stored)) return { kind: "mixed", stored: first.stored };
      }
    }
    return first;
  };

  const seen = new Map<string, SourceCell>();
  const sourced = new Set<string>();
  for (const c of input.cells) {
    const where = { file: c.file, line: c.line };
    const addr = c.address;
    const shown = addr.written === false ? undefined : addr.sheet;
    const label = c.name === undefined ? `@${formatCellAddress(addr.range, shown)}` : `${c.name} @${formatCellAddress(addr.range, shown, addr.spill)}`;
    const problem = (code: string, message: string, extra: Partial<SourceProblem> = {}): void => {
      out.problems.push({ severity: "error", code, message: `${label}: ${message}`, ...where, ...extra });
    };

    // The sheet: the address's, else the @scope block's.
    let sheet: string | undefined;
    if (addr.sheet !== undefined) {
      sheet = sheetByLower.get(lower(addr.sheet));
      if (sheet === undefined) {
        problem("unknown-sheet", `the workbook has no sheet '${addr.sheet}'`);
        continue;
      }
      if (c.scope !== undefined && lower(c.scope) !== lower(sheet)) {
        problem("address", `a name local to ${c.scope} has its cell on ${c.scope}: the address cannot name sheet ${addr.sheet}`);
        continue;
      }
    } else if (c.scope !== undefined) {
      sheet = sheetByLower.get(lower(c.scope));
      if (sheet === undefined) {
        // The names part reports the unknown @scope for named statements.
        if (c.name === undefined) problem("unknown-sheet", `@scope(${c.scope}): the workbook has no such sheet`);
        continue;
      }
    } else {
      problem("address", `a cell statement outside a sheet file names its sheet: @Sheet!${addr.range}`);
      continue;
    }
    const rect = parseCellRange(addr.range);
    if (!rect) {
      problem("address", `'${addr.range}' is not a cell or a range`);
      continue;
    }
    const range = rangeText(rect);
    if (c.name !== undefined && (rect.r1 !== rect.r2 || rect.c1 !== rect.c2)) {
      problem("address", "a named cell statement is one cell");
      continue;
    }

    const info = c.name !== undefined ? input.named(c) : undefined;
    if (c.name !== undefined && !info) continue;
    const id = info ? (info.lockKey ?? info.key) : `${sheet}!${range}`;
    const dup = seen.get(lower(info ? info.key : id));
    if (dup) {
      problem("duplicate", `${info ? info.key : `${sheet}!${range}`} has another statement at ${dup.file}:${dup.line}`);
      continue;
    }
    seen.set(lower(info ? info.key : id), c);

    // Where the statement is in the workbook now: a named cell follows its name.
    let curSheet = sheet;
    let curRect: Rect = rect;
    /** Whether the name is on the spill in the workbook; undefined when it is not one cell. */
    let xlSpill: boolean | undefined;
    let xl: XlState | undefined;
    if (info) {
      const d = info.current;
      if (!d) {
        // Deleted in Excel: the names part reports it; a formula edited in the source conflicts.
        const lockE = lockCells.get(lower(id))?.e;
        if (lockE) sourced.add(lower(lockCells.get(lower(id))!.key));
        let h: string | null = null;
        try {
          h = c.formula === "" ? null : definitionHashLike(lockE?.formula ?? null, compileWithDiagnostics(c.formula, { names: input.names, links, allowUnknownFunctions: true }).text);
        } catch {
          h = definitionHashLike(lockE?.formula ?? null, c.formula);
        }
        if (lockE && h !== lockE.formula) {
          out.sourceEdits.add(lower(info.key));
          out.conflicts.push({
            kind: "deleted-in-excel",
            key: info.key,
            message: `${info.key} was deleted in Excel since the last pull, and its cell's formula changed in the source`,
            source: { name: c.name!, scope: info.scope, display: c.formula, comment: null, hidden: false, cell: { sheet, range }, file: c.file, line: c.line },
          });
        }
        continue;
      }
      const t = definitionTarget(d.definition);
      const tSheet = t ? sheetByLower.get(lower(t.sheet ?? (d.scope.kind === "sheet" ? d.scope.name : ""))) : undefined;
      if (!t || tSheet === undefined || t.r1 !== t.r2 || t.c1 !== t.c2 || (d.scope.kind === "sheet" && lower(d.scope.name) !== lower(tSheet))) {
        xl = { kind: "not-a-cell", definition: displayOf(d.definition, links) };
      } else {
        curSheet = tSheet;
        curRect = { r1: t.r1, c1: t.c1, r2: t.r1, c2: t.c1 };
        xlSpill = t.spill;
      }
    }
    xl ??= cellState(curSheet, curRect);

    let lockE = lockCells.get(lower(id))?.e;
    if (lockE) sourced.add(lower(lockCells.get(lower(id))!.key));
    if (!lockE) {
      if (v2) {
        out.sourceEdits.add(lower(info ? info.key : `${sheet}!${range}`));
        problem(
          "address",
          info
            ? `${info.key} was not a cell statement at the last pull: cells are named in Excel (then pull)`
            : `no cell statement at ${sheet}!${range} at the last pull: addresses are set in Excel (then pull)`,
        );
        continue;
      }
      // A format-1 lockfile has no cells: the workbook is the base (a two-way check).
      if (xl.kind === "not-a-cell") {
        problem("address", `${info!.key} does not name one cell in the workbook`);
        continue;
      }
      lockE = { sheet: curSheet, range: rangeText(curRect), formula: xlHash(xl, null) };
    }

    // The address is read-only.
    if (lower(lockE.sheet) !== lower(sheet) || lockE.range !== range) {
      const was = lower(lockE.sheet) === lower(sheet) ? lockE.range : formatCellAddress(lockE.range, lockE.sheet);
      problem("address", `the address is set in Excel and read-only; the last pull had @${was}`, lower(lockE.sheet) === lower(sheet) ? { fix: { start: addr.rangeStart, end: addr.rangeEnd, text: lockE.range } } : {});
      continue;
    }

    // The `#`: what the name covers, three-way.
    let covers = addr.spill;
    if (info?.current) {
      const d = info.current;
      const lockDef = info.lockKey !== undefined ? lock.names[info.lockKey]?.definition : undefined;
      // Where the name was at the lockfile: its definition then, recognised by its hash.
      let lockSpill: boolean | undefined;
      if (lockDef !== undefined) {
        if (definitionHashLike(lockDef, d.definition) === lockDef) lockSpill = xlSpill;
        else {
          const r = parseCellRange(lockE.range);
          const ref = r ? cellRef(lockE.sheet, r.r1, r.c1) : undefined;
          if (ref !== undefined && definitionHashLike(lockDef, ref) === lockDef) lockSpill = false;
          else if (ref !== undefined && definitionHashLike(lockDef, spillRef(ref)) === lockDef) lockSpill = true;
        }
      }
      if (!covers && !explicit && lockSpill === true && xlSpill === true) {
        covers = true;
        out.implicitSpill.push(info.key);
      }
      const base = lockSpill ?? xlSpill;
      const edited = base !== undefined && covers !== base;
      if (edited) {
        out.sourceEdits.add(lower(info.key));
        out.sourceEdits.add(lower(lockCells.get(lower(id))?.key ?? id));
      }
      const ref = cellRef(curSheet, curRect.r1, curRect.c1);
      const ns: NameSpill = { stored: xlSpill === covers ? d.definition : covers ? spillRef(ref) : ref, edited, srcChanged: edited, xlChanged: xlSpill !== base };
      if (edited && covers && c.formula === "") {
        // Not measured: a name on `C6#` of a cell left empty. Written when the cell gets a
        // formula; the checker's `spill-empty` warning says so (check.ts).
        ns.srcChanged = false;
        ns.stored = d.definition;
      }
      if (edited && xlSpill === undefined) {
        out.conflicts.push({
          kind: "both-changed",
          key: info.key,
          message: `${info.key}: its '#' was changed in the source and the name was redefined in Excel since the last pull (${displayOf(d.definition, links)}): resolve by hand`,
          source: { name: c.name!, scope: info.scope, display: c.formula, comment: null, hidden: false, cell: { sheet, range }, file: c.file, line: c.line },
        });
        continue;
      }
      out.spill.set(lower(info.key), ns);
    }
    // A name without `#` on a formula that spills is the checker's `spill-uncovered` warning
    // (check.ts): the editor, `xln check` and the build report it alike (§14 issue 8).

    // The source's formula.
    let srcStored: string | null = null;
    let srcErrors: string[] = [];
    if (c.formula !== "") {
      try {
        const r = compileWithDiagnostics(c.formula, { names: input.names, links });
        srcErrors = r.diagnostics.filter((g) => g.severity === "error").map((g) => g.message);
        srcStored = srcErrors.length ? compileWithDiagnostics(c.formula, { names: input.names, links, allowUnknownFunctions: true }).text : r.text;
      } catch (e) {
        srcErrors = [(e as Error).message.split("\n")[0]!];
        srcStored = c.formula;
      }
    }
    let lockHash = lockE.formula;
    if (isV2Hash(lockHash) && srcStored !== null && definitionHashV2(srcStored) === lockHash) {
      // The source still has what a format-1/2 lockfile locked: compare in format 3, so
      // Excel's spelling of a number (`1E-14` saved as `0.00000000000001`) is no change.
      lockHash = definitionHash(srcStored);
      const lk = lockCells.get(lower(id));
      if (lk) out.lockUpgrades.set(lk.key, lockHash);
    }
    const srcHash = srcStored === null ? null : definitionHashLike(lockHash, srcStored);
    const xh = xlHash(xl, lockHash);
    let srcChanged = srcHash !== lockHash;
    const xlChanged = xh !== lockHash;
    if (srcChanged && !xlChanged && xl.kind === "formula" && c.formula !== "" && equalModuloWhitespace(displayOf(xl.stored, links), c.formula)) srcChanged = false;
    // The only edit is a rename's token (stretch G): the rename rewrites the cells.
    const viaRename = srcChanged && !xlChanged && xl.kind === "formula" && srcStored !== null ? input.renamed?.(xl.stored, curSheet) : undefined;
    const renamedOnly = viaRename !== undefined && (definitionHashLike(lockHash, viaRename) === srcHash || equalModuloWhitespace(displayOf(viaRename, links), c.formula));

    const key = info ? info.key : `${curSheet}!${range}`;
    if (srcChanged) {
      out.sourceEdits.add(lower(key));
      out.sourceEdits.add(lower(lockCells.get(lower(id))?.key ?? id));
    }
    const curRange = rangeText(curRect);
    if (info && xl.kind !== "not-a-cell" && (lower(curSheet) !== lower(lockE.sheet) || curRange !== lockE.range)) {
      out.excelChanges.push({
        kind: "moved",
        key,
        message: `${info.key} moved in Excel from ${lockE.sheet}!${lockE.range} to ${curSheet}!${curRange} since the last pull; its statement follows the name (pull again to update the address)`,
      });
    }
    const state: CellState = { sheet: curSheet, range: curRange, display: c.formula, stored: srcStored };
    if (info) state.name = info.key;
    const sideState = (display: string, scope: Scope): NameState => ({ name: c.name ?? "", scope, display, comment: null, hidden: false, cell: { sheet: curSheet, range: curRange } });

    if (!srcChanged) {
      if (xlChanged) {
        out.excelChanges.push({ kind: "changed", key, message: `${key}: the formula was changed in Excel since the last pull${xl.kind === "empty" ? " (the cell is now empty)" : ""}; the build keeps Excel's (pull again to update the source)` });
      } else {
        out.unchanged++;
        out.inSync.push({ ...state, stored: xl.kind === "formula" ? xl.stored : null });
      }
      continue;
    }
    if (xlChanged) {
      if (xh === srcHash) {
        out.unchanged++;
        out.inSync.push(state);
        continue;
      }
      out.conflicts.push({
        kind: lockHash === null && xh !== null && xl.kind === "formula" ? "created-in-both" : "both-changed",
        key,
        message: `${key}: the formula was changed in Excel since the last pull and in the source: resolve by hand (pull into a copy to see Excel's version)`,
        excel: sideState(xlDisplay(xl, links), info?.scope ?? curSheet),
        source: { ...sideState(c.formula, info?.scope ?? curSheet), file: c.file, line: c.line },
      });
      continue;
    }
    if (srcErrors.length) {
      for (const m of srcErrors) problem("syntax", m);
      continue;
    }
    if (renamedOnly) {
      for (const k of [key, lockCells.get(lower(id))?.key ?? id]) out.renamedOnly.add(lower(k));
      out.inSync.push({ ...state, stored: viaRename! });
      continue;
    }
    if (srcStored === null) {
      // `= ;`: clear the cells (a slot already empty has nothing to do).
      if (xl.kind === "formula" || xl.kind === "mixed") {
        const ch: Change = { op: "clear-cell-formula", sheet: curSheet, range: curRange, previous: xl.stored };
        if (info) ch.name = info.key;
        out.changes.push(ch);
      }
    } else {
      // Cells keep the workbook's layout (author's decision, 2026-10-09): a formula Excel
      // had on one line is written on one line, though pull laid it out on several in the
      // source; one laid out by hand in Excel keeps the source's layout. A new formula (a
      // slot) is written on one line, as Excel's own cells mostly are.
      const prev = xl.kind === "formula" || xl.kind === "mixed" ? xl.stored : undefined;
      const flat = (prev === undefined || !hasLayoutBreaks(prev)) && hasLayoutBreaks(srcStored);
      const stored = flat ? oneLine(srcStored) : srcStored;
      const ch: SetCellFormula = { op: "set-cell-formula", sheet: curSheet, range: curRange, stored, display: flat ? oneLine(c.formula) : c.formula };
      if (info) ch.name = info.key;
      if (prev !== undefined) ch.previous = prev;
      if (flat) ch.layout = prev === undefined ? ONE_LINE_NEW : ONE_LINE_AS_WORKBOOK;
      out.changes.push(ch);
      state.stored = stored;
    }
    out.inSync.push(state);
  }

  if (v2) {
    // Statements the source no longer has: the cells stay as they are (named ones: the
    // names part deletes the name).
    const missing = new Map<string, string[]>();
    for (const [k, { e }] of lockCells) {
      if (!sourced.has(k)) out.sourceEdits.add(k);
      if (sourced.has(k) || e.name !== undefined) continue;
      const l = missing.get(e.sheet) ?? [];
      l.push(e.range);
      missing.set(e.sheet, l);
    }
    for (const [sheet, ranges] of missing) {
      out.problems.push({
        severity: "warning",
        code: "cell-missing",
        message: `${sheet}: ${ranges.length} cell statement${ranges.length === 1 ? "" : "s"} of the last pull ${ranges.length === 1 ? "is" : "are"} not in the source (${ranges.slice(0, 10).join(", ")}${ranges.length > 10 ? ", …" : ""}); the build leaves ${ranges.length === 1 ? "that cell" : "those cells"} as ${ranges.length === 1 ? "it is" : "they are"} (write @C5 = ; to clear one)`,
      });
    }
    // Formula cells Excel added since the last pull.
    const covered = new Map<string, Set<string>>();
    for (const { e } of lockCells.values()) {
      const r = parseCellRange(e.range);
      if (!r) continue;
      const s = covered.get(lower(e.sheet)) ?? new Set<string>();
      for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) s.add(`${row},${col}`);
      covered.set(lower(e.sheet), s);
    }
    const now = cellStatements(wb, values ? { values } : {});
    for (const [sheet, list] of now) {
      const s = covered.get(lower(sheet));
      const added: string[] = [];
      for (const st of list) {
        if (st.kind !== "unnamed") continue;
        let all = true;
        for (let row = st.row; row < st.row + st.rows && all; row++) for (let col = st.col; col < st.col + st.cols && all; col++) if (!s?.has(`${row},${col}`)) all = false;
        if (!all) added.push(st.range);
      }
      if (added.length) {
        out.excelChanges.push({
          kind: "created",
          key: sheet,
          message: `${sheet}: formula cells added in Excel since the last pull (${added.slice(0, 10).join(", ")}${added.length > 10 ? ", …" : ""}); the build leaves them (pull again to bring them into the source)`,
        });
      }
    }
  }
  return out;
}
