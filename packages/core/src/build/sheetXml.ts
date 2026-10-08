// Patching a worksheet part for cell-formula changes (M3b, probe F8's recipe). The part is
// read with the XML tokenizer and new markup is spliced in at the tokens' offsets: only the
// `<c>` elements a change reaches are rewritten, a missing `<row>` or `<c>` is inserted for
// a slot, and every other byte stays as Excel wrote it. No regular expressions: probe F7's
// first run corrupted a sheet with one over `<f>` elements.
//
// What the recipe asks of each cell (probes/README.md § F8):
//   - a written formula is in dynamic-array form, `cm="N"` + `<f t="array" ref="<cell>">`:
//     a plain `<f>` has legacy implicit-intersection semantics, and `t="array"` without
//     `cm` is a legacy CSE array;
//   - the cell's `<v>`, `<is>` and `t` go (Excel recalculates: `fullCalcOnLoad`); `s` and
//     the other attributes stay;
//   - a spill anchor's old spill area is emptied, or the stale values block the new spill
//     (`#SPILL!`);
//   - a shared-formula master that changes un-shares its group: every other member gets
//     its own text, translated from the master's.

import { columnName, parseCell, type CellAddress } from "../file/cellref.js";
import { XmlError, XmlReader, type XmlOpen } from "../file/xml.js";
import { shiftFormula } from "../lang/shift.js";
import { escapeXmlAttr, escapeXmlText } from "./workbookXml.js";

/** A child element of a `<c>`. */
export interface CellChild {
  local: string;
  start: number;
  end: number;
  open: XmlOpen;
  /** Text content (for `<f>` and `<v>`), entities decoded. */
  text: string;
}

export interface CellElement {
  /** `B3`, from `r` or from the cell's position. */
  ref: string;
  row: number;
  col: number;
  start: number;
  end: number;
  open: XmlOpen;
  children: CellChild[];
}

export interface RowElement {
  r: number;
  start: number;
  end: number;
  open: XmlOpen;
  /** Where the content ends (the close tag's start); `end` when self-closing. */
  innerEnd: number;
  cells: CellElement[];
}

export interface SheetXmlLayout {
  /** Qualified-name prefix of `<sheetData>` (`""` or `x:`). */
  prefix: string;
  sheetData: { start: number; end: number; innerEnd: number; open: XmlOpen } | undefined;
  rows: RowElement[];
}

function prefixOf(name: string): string {
  const k = name.indexOf(":");
  return k < 0 ? "" : name.slice(0, k + 1);
}

/** Rows and cells of a worksheet part with their offsets. Cells without `r` get the address
 *  of their position, as Excel reads them. */
export function scanSheetXml(xml: string): SheetXmlLayout {
  const r = new XmlReader(xml);
  const out: SheetXmlLayout = { prefix: "", sheetData: undefined, rows: [] };
  let depth = 0;
  let row: RowElement | undefined;
  let cell: CellElement | undefined;
  let child: CellChild | undefined;
  let rowNo = 0;
  let colNo = 0;
  let inSheetData = false;
  for (let t = r.next(); t; t = r.next()) {
    if (t.type === "open") {
      depth++;
      if (depth === 2 && t.local === "sheetData") {
        if (out.sheetData) throw new XmlError("more than one <sheetData>", t.start);
        out.prefix = prefixOf(t.name);
        out.sheetData = { start: t.start, end: t.end, innerEnd: t.end, open: t };
        inSheetData = true;
      } else if (depth === 3 && inSheetData) {
        if (t.local !== "row") continue;
        const rr = t.attrs["r"];
        rowNo = rr !== undefined ? Number(rr) : rowNo + 1;
        if (!Number.isInteger(rowNo) || rowNo < 1) throw new XmlError(`row number "${rr}" is not valid`, t.start);
        colNo = 0;
        row = { r: rowNo, start: t.start, end: t.end, open: t, innerEnd: t.end, cells: [] };
      } else if (depth === 4 && row && t.local === "c") {
        const rr = t.attrs["r"];
        let a: CellAddress | undefined = rr !== undefined ? parseCell(rr) : { row: rowNo, col: colNo + 1 };
        if (!a || a.row !== rowNo) throw new XmlError(`cell address "${rr}" is not valid in row ${rowNo}`, t.start);
        colNo = a.col;
        cell = { ref: columnName(a.col) + a.row, row: a.row, col: a.col, start: t.start, end: t.end, open: t, children: [] };
      } else if (depth === 5 && cell) {
        child = { local: t.local, start: t.start, end: t.end, open: t, text: "" };
      }
    } else if (t.type === "text") {
      if (child && depth === 5) child.text += t.text;
    } else {
      if (depth === 5 && cell && child) {
        child.end = t.end;
        cell.children.push(child);
        child = undefined;
      } else if (depth === 4 && row && cell) {
        cell.end = t.end;
        row.cells.push(cell);
        cell = undefined;
      } else if (depth === 3 && row) {
        row.innerEnd = t.start;
        row.end = t.end;
        out.rows.push(row);
        row = undefined;
      } else if (depth === 2 && inSheetData) {
        out.sheetData!.innerEnd = t.start;
        out.sheetData!.end = t.end;
        inSheetData = false;
      }
      depth--;
    }
  }
  return out;
}

// ---------------------------------------------------------------------------------------
// Ranges

export interface Rect {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** `B40:G40`, `C6` (with or without `$`) as a rectangle; undefined if it is not a range of cells. */
export function parseRange(range: string): Rect | undefined {
  const k = range.indexOf(":");
  const a = parseCell(k < 0 ? range : range.slice(0, k));
  const b = k < 0 ? a : parseCell(range.slice(k + 1));
  if (!a || !b) return undefined;
  return { r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) };
}

export function rectSize(r: Rect): number {
  return (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);
}

/** Cells of a rectangle, row by row. */
export function* rectCells(r: Rect): Generator<CellAddress> {
  for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) yield { row, col };
}

function inRect(r: Rect, row: number, col: number): boolean {
  return row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
}

// ---------------------------------------------------------------------------------------
// The patch

/** A cell change on one sheet, as the change set gives it. */
export type SheetCellOp =
  | { op: "set"; range: string; /** The top-left cell's stored formula. */ stored: string }
  | { op: "clear"; range: string };

/** What the patch did, for reports and tests. */
export interface SheetPatchReport {
  /** Cells given a formula, row by row. */
  set: string[];
  /** Cells whose formula was cleared. */
  cleared: string[];
  /** `<c>` elements inserted for slots. */
  inserted: string[];
  /** `<row>` elements inserted. */
  rowsInserted: number[];
  /** Cells of an old spill (or legacy array) emptied: removed, or kept for their style. */
  ghosts: string[];
  /** Shared-group members that got their own text because their master changed. */
  unshared: string[];
}

export class SheetPatchError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SheetPatchError";
  }
}

type Action =
  | { kind: "set"; text: string }
  | { kind: "clear" }
  | { kind: "ghost" }
  | { kind: "unshare"; text: string };

/** Attributes describing a cell's value, gone with the value: the value type, the cell-metadata
 *  index (rewritten for a formula) and the value-metadata index (rich values). */
const VALUE_ATTRS = new Set(["t", "cm", "vm"]);
/** Children replaced or removed with the formula and its cached value. */
const VALUE_CHILDREN = new Set(["f", "v", "is"]);

function fOf(c: CellElement): CellChild | undefined {
  return c.children.find((x) => x.local === "f");
}

function openTag(open: XmlOpen, attrs: [string, string][], selfClosing: boolean): string {
  let s = `<${open.name}`;
  for (const [k, v] of attrs) s += ` ${k}="${escapeXmlAttr(v)}"`;
  return s + (selfClosing ? "/>" : ">");
}

/** The cell's attributes without the value's, with `cm` placed where Excel puts it (after `r`, `s`). */
function cellAttrs(c: CellElement, cm: number | undefined): [string, string][] {
  const out: [string, string][] = [];
  let placed = cm === undefined;
  for (const [k, v] of Object.entries(c.open.attrs)) {
    if (!placed && k !== "r" && k !== "s") {
      out.push(["cm", String(cm)]);
      placed = true;
    }
    if (!VALUE_ATTRS.has(k)) out.push([k, v]);
  }
  if (!placed) out.push(["cm", String(cm)]);
  return out;
}

function formulaXml(prefix: string, ref: string, text: string): string {
  return `<${prefix}f t="array" ref="${ref}">${escapeXmlText(text)}</${prefix}f>`;
}

/** A new `<c>` for a slot. */
function newCellXml(prefix: string, ref: string, text: string, cm: number): string {
  return `<${prefix}c r="${ref}" cm="${cm}">${formulaXml(prefix, ref, text)}</${prefix}c>`;
}

function rewriteCell(xml: string, c: CellElement, a: Action, cm: number): string {
  const p = prefixOf(c.open.name);
  const rest = c.children.filter((x) => !VALUE_CHILDREN.has(x.local)).map((x) => xml.slice(x.start, x.end));
  switch (a.kind) {
    case "set":
      return openTag(c.open, cellAttrs(c, cm), false) + formulaXml(p, c.ref, a.text) + rest.join("") + `</${c.open.name}>`;
    case "clear":
    case "ghost": {
      const attrs = cellAttrs(c, undefined);
      // An old spill's cell with nothing but its address goes, as Excel drops it; one with a
      // style keeps it, so the spill area keeps its formatting.
      if (a.kind === "ghost" && rest.length === 0 && attrs.every(([k]) => k === "r")) return "";
      return rest.length === 0 ? openTag(c.open, attrs, true) : openTag(c.open, attrs, false) + rest.join("") + `</${c.open.name}>`;
    }
    case "unshare": {
      // Only the <f> changes: the formula means what it meant, so its cached value stays valid.
      const f = fOf(c)!;
      const attrs = Object.entries(f.open.attrs).filter(([k]) => k !== "t" && k !== "ref" && k !== "si");
      const fx = openTag(f.open, attrs, false) + escapeXmlText(a.text) + `</${f.open.name}>`;
      return xml.slice(c.start, f.start) + fx + xml.slice(f.end, c.end);
    }
  }
}

/**
 * Applies cell changes to a worksheet part. `cm` is the 1-based `cellMetadata` record of
 * the dynamic-array type. Changes are applied in order: a later change to the same cell
 * wins. Throws `SheetPatchError` for a change Excel itself would refuse (part of a legacy
 * array, a data table).
 */
export function patchSheetXml(xml: string, ops: readonly SheetCellOp[], cm: number): { xml: string; report: SheetPatchReport } {
  const layout = scanSheetXml(xml);
  if (!layout.sheetData) throw new SheetPatchError("the worksheet has no <sheetData>");
  const report: SheetPatchReport = { set: [], cleared: [], inserted: [], rowsInserted: [], ghosts: [], unshared: [] };
  const cells = new Map<string, CellElement>();
  for (const row of layout.rows) for (const c of row.cells) cells.set(c.ref, c);

  // What each cell becomes; a later op on a cell replaces an earlier one.
  const targets = new Map<string, Action>();
  const order: CellAddress[] = [];
  for (const op of ops) {
    const rect = parseRange(op.range);
    if (!rect) throw new SheetPatchError(`'${op.range}' is not a cell or a range of cells`);
    for (const a of rectCells(rect)) {
      const ref = columnName(a.col) + a.row;
      if (!targets.has(ref)) order.push(a);
      targets.set(ref, op.op === "set" ? { kind: "set", text: shiftFormula(op.stored, a.row - rect.r1, a.col - rect.c1) } : { kind: "clear" });
    }
  }

  // Legacy arrays and data tables cover cells that are not their own: Excel refuses to
  // change part of one, and so does xln.
  const arrays: { anchor: CellElement; rect: Rect; dynamic: boolean }[] = [];
  for (const c of cells.values()) {
    const f = fOf(c);
    if (!f) continue;
    const t = f.open.attrs["t"];
    if (t === "dataTable") {
      const rect = parseRange(f.open.attrs["ref"] ?? c.ref) ?? { r1: c.row, c1: c.col, r2: c.row, c2: c.col };
      for (const a of rectCells(rect)) if (targets.has(columnName(a.col) + a.row)) throw new SheetPatchError(`${columnName(a.col) + a.row} is part of a data table at ${c.ref}: change it in Excel`);
    }
    if (t === "array") {
      const rect = parseRange(f.open.attrs["ref"] ?? c.ref);
      if (rect) arrays.push({ anchor: c, rect, dynamic: c.open.attrs["cm"] !== undefined });
    }
  }
  for (const arr of arrays) {
    if (arr.dynamic || rectSize(arr.rect) === 1) continue;
    for (const ref of targets.keys()) {
      const a = parseCell(ref)!;
      if (ref !== arr.anchor.ref && inRect(arr.rect, a.row, a.col) && !targets.has(arr.anchor.ref)) {
        throw new SheetPatchError(`${ref} is part of the legacy array formula at ${arr.anchor.ref} (${arr.anchor.children.find((x) => x.local === "f")!.open.attrs["ref"]}): change the array's top-left cell`);
      }
    }
  }

  const actions = new Map<string, Action>(targets);
  // An array formula whose anchor changes leaves its old area: the values there would block
  // the new spill. Cells there that hold a formula of their own are not the array's.
  for (const arr of arrays) {
    if (!targets.has(arr.anchor.ref) || rectSize(arr.rect) === 1) continue;
    for (const row of layout.rows) {
      if (row.r < arr.rect.r1 || row.r > arr.rect.r2) continue;
      for (const c of row.cells) {
        if (c === arr.anchor || !inRect(arr.rect, c.row, c.col) || actions.has(c.ref) || fOf(c)) continue;
        actions.set(c.ref, { kind: "ghost" });
        report.ghosts.push(c.ref);
      }
    }
  }

  // Shared groups whose master changes: the other members get the master's text moved to them.
  const masters = new Map<string, { cell: CellElement; text: string }>();
  const members = new Map<string, CellElement[]>();
  for (const c of cells.values()) {
    const f = fOf(c);
    const si = f?.open.attrs["si"];
    if (f?.open.attrs["t"] !== "shared" || si === undefined) continue;
    if (f.open.attrs["ref"] !== undefined) masters.set(si, { cell: c, text: f.text });
    else {
      let m = members.get(si);
      if (!m) members.set(si, (m = []));
      m.push(c);
    }
  }
  for (const c of targets.keys()) {
    const cell = cells.get(c);
    const f = cell && fOf(cell);
    if (!cell || f?.open.attrs["t"] !== "shared") continue;
    const si = f.open.attrs["si"];
    if (si === undefined || !masters.has(si)) throw new SheetPatchError(`${c} belongs to shared formula si=${si ?? "?"}, which has no master`);
  }
  for (const [si, m] of masters) {
    if (!targets.has(m.cell.ref)) continue;
    for (const c of members.get(si) ?? []) {
      if (actions.has(c.ref)) continue;
      actions.set(c.ref, { kind: "unshare", text: shiftFormula(m.text, c.row - m.cell.row, c.col - m.cell.col) });
      report.unshared.push(c.ref);
    }
  }

  // Splices: rewritten cells, and insertions for slots.
  const edits: { start: number; end: number; text: string }[] = [];
  for (const [ref, a] of actions) {
    const c = cells.get(ref);
    if (!c) continue;
    // Clearing removes formulas only: a value the author typed is not a formula to clear.
    if (a.kind === "clear" && !fOf(c)) continue;
    edits.push({ start: c.start, end: c.end, text: rewriteCell(xml, c, a, cm) });
  }
  const sorted = [...order].sort((a, b) => a.row - b.row || a.col - b.col);
  for (const a of sorted) {
    const ref = columnName(a.col) + a.row;
    const act = targets.get(ref)!;
    if (act.kind === "set") report.set.push(ref);
    else report.cleared.push(ref);
  }
  const missing = sorted.filter((a) => {
    const ref = columnName(a.col) + a.row;
    return !cells.has(ref) && targets.get(ref)!.kind === "set";
  });
  const p = layout.prefix;
  const rowsByR = new Map(layout.rows.map((r) => [r.r, r]));
  const sd = layout.sheetData;
  const newRows: string[] = [];
  let k = 0;
  while (k < missing.length) {
    const r = missing[k]!.row;
    const inRow: CellAddress[] = [];
    while (k < missing.length && missing[k]!.row === r) inRow.push(missing[k++]!);
    const xmlOf = (a: CellAddress) => {
      const ref = columnName(a.col) + a.row;
      report.inserted.push(ref);
      return newCellXml(p, ref, (targets.get(ref) as { text: string }).text, cm);
    };
    const row = rowsByR.get(r);
    if (row?.open.selfClosing) {
      edits.push({ start: row.start, end: row.end, text: openTag(row.open, Object.entries(row.open.attrs), false) + inRow.map(xmlOf).join("") + `</${row.open.name}>` });
    } else if (row) {
      for (const a of inRow) {
        const next = row.cells.find((c) => c.col > a.col);
        const at = next ? next.start : row.innerEnd;
        edits.push({ start: at, end: at, text: xmlOf(a) });
      }
    } else {
      // `spans` is an optimisation hint; Excel accepts a row without it (F8).
      const text = `<${p}row r="${r}">${inRow.map(xmlOf).join("")}</${p}row>`;
      report.rowsInserted.push(r);
      if (sd.open.selfClosing) newRows.push(text);
      else {
        const next = layout.rows.find((x) => x.r > r);
        const at = next ? next.start : sd.innerEnd;
        edits.push({ start: at, end: at, text });
      }
    }
  }
  if (newRows.length > 0) edits.push({ start: sd.start, end: sd.end, text: `<${sd.open.name}>${newRows.join("")}</${sd.open.name}>` });
  // Insertions (empty spans) go before a rewrite that starts at the same offset.
  edits.sort((a, b) => a.start - b.start || a.end - b.end);
  let out = "";
  let at = 0;
  for (const e of edits) {
    if (e.start < at) throw new SheetPatchError("overlapping edits (internal error)");
    out += xml.slice(at, e.start) + e.text;
    at = e.end;
  }
  out += xml.slice(at);
  report.ghosts.sort(byRef);
  report.unshared.sort(byRef);
  return { xml: out, report };
}

function byRef(x: string, y: string): number {
  const a = parseCell(x)!;
  const b = parseCell(y)!;
  return a.row - b.row || a.col - b.col;
}

// ---------------------------------------------------------------------------------------
// Reading every cell (for the read-back)

export interface CellSnapshot {
  /** The cell's attributes, `r` included. */
  attrs: Record<string, string>;
  f?: { attrs: Record<string, string>; text: string };
  v?: string;
  /** The `<is>` markup, as stored. */
  is?: string;
  /** Other children's markup (`extLst`), as stored. */
  other: string;
}

/** Every `<c>` of a worksheet part, by address. */
export function readSheetCells(xml: string): Map<string, CellSnapshot> {
  const out = new Map<string, CellSnapshot>();
  for (const row of scanSheetXml(xml).rows) {
    for (const c of row.cells) {
      const s: CellSnapshot = { attrs: c.open.attrs, other: "" };
      for (const x of c.children) {
        if (x.local === "f") s.f = { attrs: x.open.attrs, text: x.text };
        else if (x.local === "v") s.v = x.text;
        else if (x.local === "is") s.is = xml.slice(x.start, x.end);
        else s.other += xml.slice(x.start, x.end);
      }
      out.set(c.ref, s);
    }
  }
  return out;
}

/** The part with `<sheetData>` cut out: what a cell build must leave byte-identical. */
export function outsideSheetData(xml: string): string {
  const l = scanSheetXml(xml);
  return l.sheetData ? xml.slice(0, l.sheetData.start) + xml.slice(l.sheetData.end) : xml;
}
