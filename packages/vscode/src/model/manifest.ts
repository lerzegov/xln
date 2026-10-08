// The parts of workbook.manifest.json (core/src/project/manifest.ts, format xln.manifest/1)
// that the extension reads. Everything optional is checked before use: the file is
// written by `xln pull`, but a user may have edited or truncated it.
import { extentSize } from "@xln/core";

export interface ManifestSpill {
  sheet: string | null;
  anchor: string;
  extent: string | null;
  spilling: boolean | null;
}

export interface ManifestUsedBy {
  names?: string[];
  /** Sheet → cells and rectangles (`B3:B7`). */
  cells?: Record<string, string[]>;
  /** Sheet → `sqref` of each conditional format. */
  conditionalFormats?: Record<string, string[]>;
  /** Sheet → `sqref` of each data validation. */
  dataValidations?: Record<string, string[]>;
  /** `Table[Column]`. */
  tableColumns?: string[];
}

export interface ManifestName {
  name: string;
  scope: string | null;
  kind: string;
  arity?: { required: number; optional: number };
  params?: string[];
  table?: string;
  spill?: ManifestSpill;
  error?: string;
  hidden?: boolean;
  module: string | null;
  /** For a named cell or a slot (M3b): the cell its statement is written on. */
  cell?: { sheet: string; range: string };
  file: string;
  uses?: string[];
  usedBy?: ManifestUsedBy;
}

export interface ManifestTable {
  name: string;
  sheet: string;
  ref: string;
  columns: string[];
  calculatedColumns?: Record<string, string>;
}

export interface Manifest {
  format: string;
  workbook: string;
  sheets: { name: string; position: number; state?: string; kind?: string }[];
  tables: ManifestTable[];
  names: Record<string, ManifestName>;
  /** Sheet → the ranges of its unnamed cell statements (M3b); absent in older manifests. */
  unnamedCells?: Record<string, string[]>;
  /** Sheet → the dynamic arrays that spilled beyond their anchor when the workbook was saved. */
  spills?: Record<string, { anchor: string; extent: string }[]>;
  /** Links to other workbooks (`[n]` in stored formulas); absent when the workbook has none. */
  externalLinks?: { index: number; book: string }[];
}

/** Parses the manifest text; throws with a readable message when it is not one. */
export function parseManifest(text: string): Manifest {
  const m = JSON.parse(text) as Partial<Manifest>;
  if (typeof m !== "object" || m === null || typeof m.format !== "string" || !m.format.startsWith("xln.manifest/")) {
    throw new Error("not an xln manifest (no \"format\": \"xln.manifest/…\")");
  }
  return {
    format: m.format,
    workbook: typeof m.workbook === "string" ? m.workbook : "",
    sheets: Array.isArray(m.sheets) ? m.sheets : [],
    tables: Array.isArray(m.tables) ? m.tables : [],
    names: typeof m.names === "object" && m.names !== null ? m.names : {},
    ...(typeof m.unnamedCells === "object" && m.unnamedCells !== null ? { unnamedCells: m.unnamedCells } : {}),
    ...(typeof m.spills === "object" && m.spills !== null ? { spills: m.spills } : {}),
    ...(Array.isArray(m.externalLinks) ? { externalLinks: m.externalLinks.filter((l) => typeof l?.index === "number" && typeof l?.book === "string") } : {}),
  };
}

/** Number of cells in `A1` or `A1:C3`; 1 when it cannot tell. */
export function cellCount(range: string): number {
  const size = extentSize(range);
  return size ? size.rows * size.cols : 1;
}

/** Counts of the workbook-side usages of a name. */
export function usageCounts(u: ManifestUsedBy | undefined): { cells: number; conditionalFormats: number; dataValidations: number; tableColumns: number } {
  const sum = (r: Record<string, string[]> | undefined, f: (s: string) => number) =>
    Object.values(r ?? {}).reduce((a, list) => a + list.reduce((b, x) => b + f(x), 0), 0);
  return {
    cells: sum(u?.cells, cellCount),
    conditionalFormats: sum(u?.conditionalFormats, () => 1),
    dataValidations: sum(u?.dataValidations, () => 1),
    tableColumns: u?.tableColumns?.length ?? 0,
  };
}
