// E4, `xln verify`: compares the cached values of two copies of a workbook cell by cell:
// the file before a build, and the same file after the author opened the built copy in
// Excel (which recalculates it, F4) and saved it. A build of names that should not change
// numbers must show no changed cell.
//
// Every cell with a value counts, not only formula cells: a spill's values beyond its
// anchor are cells without a formula of their own.

import { Package, relTypeIs } from "../file/package.js";
import { formatCell, parseCell } from "../file/cellref.js";
import { childElements, ownText, parseXml, readElement, XmlReader, firstChild } from "../file/xml.js";
import { richText } from "../file/worksheet.js";
import { workbookPartOf } from "./apply.js";

export type CellValue = number | string | boolean | { error: string };

export interface SheetValues {
  sheet: string;
  /** Address → value; cells without a stored value are absent. */
  cells: Map<string, CellValue>;
}

/** Every cached cell value of every worksheet, in sheet order. */
export function readCellValues(bytes: Uint8Array): SheetValues[] {
  return readValues(bytes).sheets;
}

/**
 * Whether a file's cached values are Excel's. Excel writes a value for every formula cell
 * and never writes `fullCalcOnLoad` (every Excel-saved file in probes/results and the
 * corpus). A file written by another tool after Excel's last save carries
 * `fullCalcOnLoad="1"` (xln's builds after a cell change, XlsxWriter) and no value, or a
 * placeholder 0 or "" (XlsxWriter), on the formula cells it did not calculate.
 */
export interface ValuesOrigin {
  /** `excel`: Excel calculated them; `none`: no formula cell has a value Excel calculated (the file was never saved by Excel); `partial`: some formula cells have none (written by a tool after Excel's last save). */
  origin: "excel" | "none" | "partial";
  formulaCells: number;
  /** Formula cells without a value (absent or empty). */
  withoutValue: number;
  /** `<calcPr fullCalcOnLoad="1">`: written by something other than Excel since its last save. */
  fullCalcOnLoad: boolean;
}

interface ValueStats {
  formulaCells: number;
  withoutValue: number;
  /** Formula cells whose value could be a writer's placeholder: none, empty, or the number 0. */
  placeholder: number;
}

function readValues(bytes: Uint8Array): { sheets: SheetValues[]; origin: ValuesOrigin } {
  const pkg = new Package(bytes);
  const wbPart = workbookPartOf(pkg);
  const wb = parseXml(pkg.text(wbPart) ?? "");
  const rels = pkg.rels(wbPart);
  const ssRel = rels.find((r) => relTypeIs(r.type, "sharedStrings") && !r.external);
  const ssText = ssRel ? pkg.text(ssRel.target) : undefined;
  const shared = ssText ? childElements(parseXml(ssText), "si").map(richText) : [];
  const relById = new Map(rels.map((r) => [r.id, r]));
  const out: SheetValues[] = [];
  const stats: ValueStats = { formulaCells: 0, withoutValue: 0, placeholder: 0 };
  const sheetsEl = firstChild(wb, "sheets");
  for (const s of sheetsEl ? childElements(sheetsEl, "sheet") : []) {
    const name = s.attrs["name"] ?? "";
    let relId: string | undefined;
    for (const [k, v] of Object.entries(s.attrs)) if (k.endsWith(":id")) relId = v;
    const rel = relId ? relById.get(relId) : undefined;
    const text = rel && !rel.external && relTypeIs(rel.type, "worksheet") ? pkg.text(rel.target) : undefined;
    out.push({ sheet: name, cells: text === undefined ? new Map() : sheetValues(text, shared, stats) });
  }
  const full = firstChild(wb, "calcPr")?.attrs["fullCalcOnLoad"];
  const fullCalcOnLoad = full === "1" || full === "true";
  const { formulaCells, withoutValue, placeholder } = stats;
  // Without fullCalcOnLoad a 0 is Excel's; with it, a file of zeros is XlsxWriter's.
  const none = formulaCells > 0 && (withoutValue === formulaCells || (fullCalcOnLoad && placeholder === formulaCells));
  const origin = none ? "none" : fullCalcOnLoad && withoutValue > 0 ? "partial" : "excel";
  return { sheets: out, origin: { origin, formulaCells, withoutValue, fullCalcOnLoad } };
}

/** Where a file's cached values come from (see ValuesOrigin). */
export function valuesOrigin(bytes: Uint8Array): ValuesOrigin {
  return readValues(bytes).origin;
}

/**
 * The warning for one side of a verify whose values are not (all) Excel's; undefined when
 * they are. `file` is the file's name as the user knows it.
 */
export function valuesWarning(file: string, v: ValuesOrigin, side: "before" | "after"): string | undefined {
  if (v.origin === "excel") return undefined;
  const fix = side === "before" ? "open and save it in Excel first" : "open it in Excel and save it, then verify again";
  if (v.origin === "none") return `${file} has no values Excel calculated (it was never saved by Excel): ${fix}, or the comparison is empty`;
  return `${file} was written after Excel last saved it: ${v.withoutValue} of ${v.formulaCells} formula cells have no value Excel calculated, and compare as empty; ${fix}`;
}

/** `readCellValues` as sheet name → address → value, as the audit (C15) and the pull's notes take it. */
export function cellValueMap(bytes: Uint8Array): Map<string, ReadonlyMap<string, CellValue>> {
  return new Map(readCellValues(bytes).map((s) => [s.sheet, s.cells]));
}

function sheetValues(xml: string, shared: readonly string[], stats: ValueStats): Map<string, CellValue> {
  const cells = new Map<string, CellValue>();
  const r = new XmlReader(xml);
  let row = 0;
  let col = 0;
  for (let t = r.next(); t; t = r.next()) {
    if (t.type !== "open") continue;
    if (t.local === "row") {
      row = t.attrs["r"] !== undefined ? Number(t.attrs["r"]) : row + 1;
      col = 0;
      continue;
    }
    if (t.local !== "c") continue;
    const c = readElement(r, t);
    const a = c.attrs["r"] !== undefined ? parseCell(c.attrs["r"]) : { row, col: col + 1 };
    if (!a) continue;
    col = a.col;
    const addr = formatCell(a);
    const type = c.attrs["t"] ?? "n";
    if (firstChild(c, "f")) {
      stats.formulaCells++;
      const fv = firstChild(c, "v");
      const raw = fv ? ownText(fv) : "";
      if (raw === "") stats.withoutValue++;
      if (raw === "" || (type === "n" && Number(raw) === 0)) stats.placeholder++;
    }
    if (type === "inlineStr") {
      const is = firstChild(c, "is");
      if (is) cells.set(addr, richText(is));
      continue;
    }
    const v = firstChild(c, "v");
    if (!v) continue;
    const raw = ownText(v);
    if (type === "s") cells.set(addr, shared[Number(raw)] ?? `#string ${raw}`);
    else if (type === "b") cells.set(addr, raw === "1" || raw === "true");
    else if (type === "e") cells.set(addr, { error: raw });
    else if (type === "str" || type === "d") cells.set(addr, raw);
    else if (raw !== "") cells.set(addr, Number(raw));
  }
  return cells;
}

export interface ValueChange {
  sheet: string;
  cell: string;
  /** Undefined: no value on that side. */
  before: CellValue | undefined;
  after: CellValue | undefined;
}

export interface VerifyReport {
  sheets: number;
  cells: number;
  changed: ValueChange[];
  /** Sheets present on one side only. */
  sheetsAdded: string[];
  sheetsRemoved: string[];
  /** Where each side's values come from: a side whose origin is not `excel` makes the comparison empty or partial. */
  beforeValues: ValuesOrigin;
  afterValues: ValuesOrigin;
}

export interface VerifyOptions {
  /** Numbers within this relative difference count as equal. Default 0: exact. */
  tolerance?: number;
}

function same(a: CellValue | undefined, b: CellValue | undefined, tol: number): boolean {
  if (a === undefined || b === undefined) return a === b;
  if (typeof a === "number" && typeof b === "number") {
    if (a === b) return true;
    return tol > 0 && Math.abs(a - b) <= tol * Math.max(Math.abs(a), Math.abs(b));
  }
  if (typeof a === "object" && typeof b === "object") return a.error === b.error;
  return a === b;
}

/** Compares the cached values of `before` and `after`, cell by cell. */
export function verifyValues(before: Uint8Array, after: Uint8Array, opts: VerifyOptions = {}): VerifyReport {
  const tol = opts.tolerance ?? 0;
  const ra = readValues(before);
  const rb = readValues(after);
  const a = ra.sheets;
  const b = rb.sheets;
  const bByName = new Map(b.map((s) => [s.sheet, s]));
  const changed: ValueChange[] = [];
  let cells = 0;
  for (const sa of a) {
    const sb = bByName.get(sa.sheet);
    if (!sb) continue;
    const addrs = new Set([...sa.cells.keys(), ...sb.cells.keys()]);
    cells += addrs.size;
    for (const cell of [...addrs].sort(byAddress)) {
      const x = sa.cells.get(cell);
      const y = sb.cells.get(cell);
      if (!same(x, y, tol)) changed.push({ sheet: sa.sheet, cell, before: x, after: y });
    }
  }
  const aNames = new Set(a.map((s) => s.sheet));
  return {
    sheets: a.filter((s) => bByName.has(s.sheet)).length,
    cells,
    changed,
    sheetsAdded: b.filter((s) => !aNames.has(s.sheet)).map((s) => s.sheet),
    sheetsRemoved: a.filter((s) => !bByName.has(s.sheet)).map((s) => s.sheet),
    beforeValues: ra.origin,
    afterValues: rb.origin,
  };
}

function byAddress(x: string, y: string): number {
  const a = parseCell(x);
  const b = parseCell(y);
  if (!a || !b) return x < y ? -1 : x > y ? 1 : 0;
  return a.row - b.row || a.col - b.col;
}

/** A value as text for reports: `12.5`, `"text"`, `TRUE`, `#REF!`, `(empty)`. */
export function valueText(v: CellValue | undefined): string {
  if (v === undefined) return "(empty)";
  if (typeof v === "object") return v.error;
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "string") return JSON.stringify(v);
  return String(v);
}
