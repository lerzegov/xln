// The audit's result: findings of checks C1–C13 (PROJECT-BRIEF §4 C), the name census
// (C8) and the spill census (C9). Plain data, so the CLI prints it as JSON for agents and
// the editor turns it into diagnostics.

import type { CellValue } from "../build/verify.js";
import type { DependencyGraph } from "../graph/model.js";
import type { NameKind } from "../project/types.js";
import type { NameIndex } from "../view/formulas.js";

export const CHECK_IDS = ["C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "C10", "C11", "C12", "C13", "C14", "C15"] as const;
export type CheckId = (typeof CHECK_IDS)[number];

export type AuditSeverity = "error" | "warning" | "info";

/** Where a finding is: a defined name, a cell (or a shared formula's group), a format, a validation, a Table column, a chart, a package part (C14: another tool's store). */
export interface FindingWhere {
  kind: "name" | "cell" | "cf" | "dv" | "table" | "chart" | "part";
  /** The sheet: a cell's, a format's or validation's, a Table's or chart's; a sheet-scoped name's scope. */
  sheet?: string;
  /** A defined name (`kind: "name"`), or a Table column `tblSales[Total]` (`kind: "table"`). */
  name?: string;
  /** A defined name's key: `Sheet!Name` or `Name`. */
  key?: string;
  /** A cell (the anchor or the master of a shared formula), the `sqref` of a format or validation, a chart part, a package part. */
  ref?: string;
  /** The shared formula's group, or a dynamic array's saved extent. */
  range?: string;
  /** What the finding points at inside the formula as stored, and the text there. */
  span?: { start: number; end: number };
  text?: string;
}

export interface Finding {
  check: CheckId;
  /** The rule, `C2.bare-prefix`: what `AuditOptions.rules` and the rule table (`RULES`) are keyed by. */
  rule: string;
  severity: AuditSeverity;
  where: FindingWhere;
  message: string;
  hint?: string;
  /** The values the message was made from, for tools. */
  data?: Record<string, unknown>;
}

/** C8: one family of names that differ only by a coordinate tag (`Sales_base`, `Sales_payout`). */
export interface NameFamily {
  /** Sheet of the members (a family never mixes scopes). */
  scope: string | undefined;
  stem: string;
  position: "suffix" | "prefix";
  /** Member keys, in the order of `tags`. */
  members: string[];
  tags: string[];
}

export interface CensusOptions {
  /** Names left out of the tiers (harness, solver settings): glob patterns on keys (`Check!*`, `FIX.*`), case-insensitive. */
  exclude?: string[];
  /** Coordinate tags to use instead of detecting them (`["base", "payout"]`): `Sales_base`, `base_Sales`. Kept even when every family carrying one differs. */
  tags?: string[];
  /** A tag is a coordinate when at least this many families carry it (default 2). */
  minFamilies?: number;
}

export interface NameCensus {
  /** Defined names, Excel's own (`_xl…`) left out. */
  total: number;
  hidden: number;
  byKind: Record<NameKind, number>;
  byScope: { workbook: number; sheets: { sheet: string; names: number }[] };
  /** `_xlnm.` names (print areas, filters): not counted above. */
  builtIns: number;
  /** Coordinate tags found, with how many families carry each. */
  coordinates: { tag: string; position: "suffix" | "prefix"; families: number }[];
  families: NameFamily[];
  /**
   * Names standing in for dimensions (FEASIBILITY §11.2, the excel-models SUMMARY §3.2
   * tiers made generic): T1 coordinate addressing, T2 line-item identity, T3 axis
   * duplication, T4 the axis itself. `total` = T1 + T2 + T3 + T4, of `of` names.
   */
  tiers: { T1: number; T2: number; T3: number; T4: number; total: number; of: number; percent: number; T2spellings: number; library: number; excluded: number };
  /** Names in T1 (accessors), T4 (axis), the library (other LAMBDAs) and left out by `exclude`. */
  tierNames: { T1: string[]; T4: string[]; library: string[]; excluded: string[] };
  /** Short names defined on more than one sheet: the sheet stands in for a dataset coordinate. */
  sheetDuplicates: { name: string; sheets: string[] }[];
  /** How to read the tiers. */
  explanation: string;
}

/** C9: one dynamic array and the names over it. */
export interface SpillEntry {
  sheet: string;
  anchor: string;
  /** Saved extent (`C6:G6`). */
  extent: string;
  rows: number;
  cols: number;
  /** Display form. */
  formula: string;
  /**
   * Names defined over it: `spill` (`'S'!$C$6#`, follows the spill), `extent` (the saved
   * extent as a fixed range), `part` (a fixed range inside it), `anchor` (the anchor cell alone).
   */
  names: { key: string; how: "spill" | "extent" | "part" | "anchor" }[];
}

export interface SpillCensus {
  /** Dynamic arrays that spilled to more than one cell when saved, in sheet, row, column order. */
  spills: SpillEntry[];
  /** Dynamic-array formulas saved as one cell (Excel writes even 1×1 results that way). */
  singleCell: number;
  /** Dynamic-array formulas a build wrote and Excel has not calculated since (`FormulaViewLine.uncalculated`): their spill is not known, so they are in neither count above. */
  uncalculated: number;
  /** Per sheet: spills, named through `x#`, named only through a fixed range, unnamed. */
  bySheet: { sheet: string; spills: number; spillNamed: number; fixedNamed: number; unnamed: number }[];
}

export interface AuditCounts {
  error: number;
  warning: number;
  info: number;
}

export interface AuditReport {
  format: "xln-audit/1";
  workbook: string | undefined;
  counts: AuditCounts;
  byCheck: Record<CheckId, AuditCounts>;
  /** Ordered by check, then place (names, cells, formats, validations, Tables, charts), then rule. */
  findings: Finding[];
  census: NameCensus;
  spills: SpillCensus;
  /** Which checks ran. */
  checks: CheckId[];
}

export interface AuditOptions {
  /** File name, for the report. */
  workbook?: string;
  /** Run only these checks (C8 and C9 census sections are always computed). */
  only?: readonly CheckId[];
  /** Drop findings below this severity. */
  minSeverity?: AuditSeverity;
  /** Severity per rule id (`C10.unused`) or per check (`C13`), or `off`. */
  rules?: Record<string, AuditSeverity | "off">;
  /** C7. Defaults: definition and formula length warn at 7,500 and error at 8,192 characters; nesting warn above 48, error above 64. */
  limits?: { lengthWarn?: number; lengthError?: number; nestingWarn?: number; nestingError?: number };
  /**
   * C13: literal numbers allowed in LAMBDA bodies (default `DEFAULT_CONSTANTS`), and the
   * magnitude from which a number is a sentinel (an "infinity" such as 1E+99), allowed too
   * (default `DEFAULT_SENTINEL_ABOVE`, 1E+90).
   */
  constants?: { allow?: readonly number[]; sentinelAbove?: number };
  /**
   * The check harness: names that exist to be read by a person or a solver, not by formulas
   * (`Check!*`, `CHK.*`, `Model!Fix*`). Glob patterns on keys (`Sheet!Name`, `Module.Name`),
   * case-insensitive. C10 does not report them (nor the names only they read), and the C8
   * census leaves them out of the tiers as it does `census.exclude`.
   */
  harness?: readonly string[];
  census?: CensusOptions;
  /**
   * C15: the cells' cached values, sheet name → address → value (`readCellValues`). The
   * snapshot holds formulas only, and a label is a typed value; without them C15 says nothing.
   */
  values?: ReadonlyMap<string, ReadonlyMap<string, CellValue>>;
  /** A graph already built for this snapshot, and the name index it used. */
  graph?: DependencyGraph;
  names?: NameIndex;
}
