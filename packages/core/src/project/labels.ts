// Names that *Create from Selection* took from a cell's current value rather than from a
// label (feedback 2026-10-07, is-model's Mortgage sheet): A12 = XLOOKUP(…) shows
// "italian10", and Left column named B12:G12 `italian10`. When the lookup changes, A12
// shows another text and the name stays `italian10`. With both Top row and Left column,
// Excel also names the whole block after the corner cell (`Mortgage` = B12:G12). A pull
// says so (an info note, nothing changes). Conservative on purpose: the label cell must be
// the one Create from Selection reads (just left of a row, just above a column, or the
// corner), its text must give the name exactly, and a value label counts only when the
// cell is computed (a formula, or a cell a spill or array fills).
//
// And names whose label no longer gives them (`labelDrift`, check C15, decided 2026-10-07):
// the author named B12:G12 from the label "Gross income", then renamed the name
// `Gross_ind_income` (F2, `xln rename`). The build renames the name; the label cell still
// reads "Gross income", and xln never writes cell values, so it can only point it out.

import { formatCell, parseCell } from "../file/cellref.js";
import type { DefinedName, WorkbookSnapshot } from "../file/types.js";
import type { CellValue } from "../build/verify.js";
import { quoteSheet as quote } from "../lang/tokens.js";
import { definitionTarget } from "./classify.js";
import { sheetFormulaCells } from "./statements.js";

export interface ValueLabel {
  /** `Sheet!Name` or `Name`. */
  key: string;
  /** `value`: named after a computed cell's current value; `corner`: after the corner of a two-way Create from Selection. */
  kind: "value" | "corner";
  /** The label cell, `Mortgage!A12`. */
  cell: string;
  /** The cells the name covers, `Mortgage!B12:G12`. */
  range: string;
  note: string;
}

// Create from Selection's conversion, as probe F11 measured it (Excel for Mac, 2026-10-07,
// probes/results/f11_create_from_selection_mac.xlsx): the label is trimmed; each character
// a name cannot hold becomes one `_` (a CR LF line break is one character); characters a
// name cannot hold at the end are dropped (`Margin %` → `Margin`); `_` goes in front when
// the first character cannot start a name (`2024 sales`, `€uro`: `€` may follow but not
// start); a label that reads as an A1 reference or as TRUE/FALSE gets `_` behind (`Q1_`,
// `Tax2024_`, `R_`, `C_`, `rc_`, `True_`), one that reads as an R1C1 reference only gets it
// in front (`_R1C1`); a number gives no name, a date the text it shows (`31/01/2024` →
// `_31_01_2024`); a name over 255 characters is not made.
const NAME_START = /[\p{L}_\\]/u;
const NAME_CHAR = /[\p{L}\p{N}\p{M}_\\.?€]/u;
const MAX_NAME = 255;

interface Conversion {
  name: string;
  /** `_` added in front, behind. */
  front: boolean;
  back: boolean;
}

function convert(text: string): Conversion | undefined {
  const s = text.split("\r\n").join("\n").trim();
  const chars = [...s];
  // Characters a name cannot hold at the end are dropped; the measured case is `Margin %`.
  let end = chars.length;
  while (end > 0 && !NAME_CHAR.test(chars[end - 1]!)) end--;
  if (end === 0) return undefined;
  let body = "";
  for (let i = 0; i < end; i++) body += NAME_CHAR.test(chars[i]!) ? chars[i] : "_";
  let name = body;
  let front = false;
  let back = false;
  if (!NAME_START.test([...body][0]!)) {
    name = `_${body}`;
    front = true;
  } else if (isA1Cell(body) || /^(r|c|rc|true|false)$/i.test(body)) {
    name = `${body}_`;
    back = true;
  } else if (isR1C1(body)) {
    name = `_${body}`;
    front = true;
  }
  return name.length > MAX_NAME ? undefined : { name, front, back };
}

/** `Q1`, `TAX2024`: inside the sheet's bounds (XFD1048576). */
function isA1Cell(s: string): boolean {
  const a = parseCell(s);
  return a !== undefined && !s.includes("$") && a.col <= 16384 && a.row <= 1048576;
}

/** `R1C1`, `R2C`, `RC3`, `R4`, `C5` (an A1 cell is tested before). */
function isR1C1(s: string): boolean {
  const up = s.toUpperCase();
  let i = 0;
  const digits = (): void => {
    while (i < up.length && up[i]! >= "0" && up[i]! <= "9") i++;
  };
  if (up[i] === "R") {
    i++;
    digits();
    if (i === up.length) return true;
  }
  if (up[i] !== "C") return false;
  i++;
  digits();
  return i === up.length;
}

/**
 * The name Create from Selection makes from a cell (F11, above). Undefined for a cell that
 * gives none: empty, an error, a logical, a number, a text of only characters a name cannot
 * hold, or over 255 characters. A date gives the text it shows, which depends on the cell's
 * format and the reader's locale: pass it as `shown` when known (`labelNames` tries the
 * common spellings without it).
 */
export function labelName(v: CellValue | undefined, shown?: string): string | undefined {
  if (typeof v === "string") return convert(v)?.name;
  if (typeof v === "number" && shown !== undefined) return convert(shown)?.name;
  return undefined;
}

/**
 * The names Create from Selection may have given a cell, the likeliest first: `labelName`,
 * and for a whole number in the range of dates, the name of each common short-date spelling
 * of it (day, month and year in either order, `/`, `-` or `.`). Not measured, so accepted
 * too: characters a name cannot hold at the start of a label dropped instead of made `_`.
 */
export function labelNames(v: CellValue | undefined, shown?: string): string[] {
  const out = new Set<string>();
  const add = (s: string | undefined) => {
    if (s !== undefined) out.add(s);
  };
  add(labelName(v, shown));
  if (typeof v === "string") {
    const chars = [...v.trim()];
    let i = 0;
    while (i < chars.length && !NAME_CHAR.test(chars[i]!)) i++;
    if (i > 0) add(convert(chars.slice(i).join(""))?.name);
  } else if (typeof v === "number" && shown === undefined) {
    for (const d of dateSpellings(v)) add(convert(d)?.name);
  }
  return [...out];
}

/** Common short-date spellings of a 1900-system date serial (none for a fraction or out of range). */
function dateSpellings(serial: number): string[] {
  if (!Number.isInteger(serial) || serial < 1 || serial > 2958465) return [];
  // Serial 60 is Excel's 29 February 1900, a day that did not exist; later serials are one ahead.
  const days = serial > 60 ? serial - 1 : serial;
  const t = new Date(Date.UTC(1899, 11, 31) + days * 86400000);
  const y = String(t.getUTCFullYear());
  const m = t.getUTCMonth() + 1;
  const d = t.getUTCDate();
  const pad = (n: number) => String(n).padStart(2, "0");
  const out: string[] = [];
  for (const sep of ["/", "-", "."]) {
    for (const [mm, dd] of [[pad(m), pad(d)], [String(m), String(d)]] as const) {
      out.push(`${dd}${sep}${mm}${sep}${y}`, `${mm}${sep}${dd}${sep}${y}`, `${y}${sep}${mm}${sep}${dd}`);
    }
  }
  return out;
}

/** The parts of a label's conversion, for `labelReplacement`: whether `_` was added in front or behind. */
export function labelConversion(text: string): Conversion | undefined {
  return convert(text);
}

export type Target = { d: DefinedName; key: string; sheet: string; r1: number; c1: number; r2: number; c2: number };

/** Each name fixed to one sheet's cells (with `spills`, a spill by its saved extent), and whether a cell is computed. */
export function labelContext(wb: WorkbookSnapshot, names: readonly DefinedName[], spills: boolean) {
  const sheets = new Map(wb.sheets.filter((s) => s.kind === "worksheet").map((s) => [s.name.toLowerCase(), s]));
  const computed = new Map<string, (row: number, col: number) => boolean>();
  const isComputed = (sheet: string, row: number, col: number): boolean => {
    let f = computed.get(sheet);
    if (!f) {
      const fc = sheetFormulaCells(sheets.get(sheet.toLowerCase())!);
      f = (r, c) => {
        const cell = fc.cells.get(formatCell({ row: r, col: c }));
        return (cell !== undefined && cell.stored !== undefined && cell.stored.trim() !== "") || fc.covered(r, c);
      };
      computed.set(sheet, f);
    }
    return f(row, col);
  };
  const targets: Target[] = [];
  for (const d of names) {
    const t = definitionTarget(d.definition);
    const scope = d.scope.kind === "sheet" ? d.scope.name : undefined;
    if (!t || (t.spill && !spills) || !t.absolute) continue;
    const sheet = sheets.get((t.sheet ?? scope ?? "").toLowerCase());
    if (sheet === undefined) continue;
    let { r2, c2 } = t;
    if (t.spill) {
      const end = sheet.spills.find((s) => s.anchor === formatCell({ row: t.r1, col: t.c1 }))?.extent.split(":")[1];
      const e = end !== undefined ? parseCell(end) : undefined;
      if (e) ({ row: r2, col: c2 } = e);
    }
    targets.push({ d, key: scope === undefined ? d.name : `${scope}!${d.name}`, sheet: sheet.name, r1: t.r1, c1: t.c1, r2, c2 });
  }
  return { targets, isComputed };
}

/** Names that look named after a computed value or a two-way corner, one note each. */
export function valueLabels(wb: WorkbookSnapshot, names: readonly DefinedName[], values: ReadonlyMap<string, ReadonlyMap<string, CellValue>>): ValueLabel[] {
  // Each name's cells, for the corner's evidence: another name on one of its rows or columns.
  const { targets, isComputed } = labelContext(wb, names, false);

  const out: ValueLabel[] = [];
  for (const t of targets) {
    const cells = values.get(t.sheet);
    const at = (row: number, col: number) => (row >= 1 && col >= 1 ? cells?.get(formatCell({ row, col })) : undefined);
    const same = (row: number, col: number) => labelName(at(row, col))?.toLowerCase() === t.d.name.toLowerCase();
    const range = `${quote(t.sheet)}!${formatCell({ row: t.r1, col: t.c1 })}${t.r1 === t.r2 && t.c1 === t.c2 ? "" : `:${formatCell({ row: t.r2, col: t.c2 })}`}`;
    const label = (row: number, col: number) => `${quote(t.sheet)}!${formatCell({ row, col })}`;
    // Left column (a row) or Top row (a column): the label is the cell just before it.
    const before: [number, number][] = [];
    if (t.r1 === t.r2) before.push([t.r1, t.c1 - 1]);
    if (t.c1 === t.c2) before.push([t.r1 - 1, t.c1]);
    const hit = before.find(([r, c]) => same(r, c) && isComputed(t.sheet, r, c));
    if (hit) {
      const cell = label(hit[0], hit[1]);
      out.push({ key: t.key, kind: "value", cell, range, note: `${t.key}: named after the current value of ${cell}, a formula result; the name stays as is when that value changes` });
      continue;
    }
    // The corner: the cell above and left of the block, and another name on one of the
    // block's full rows or columns (what Create from Selection made beside it).
    if (same(t.r1 - 1, t.c1 - 1)) {
      const sibling = targets.some(
        (o) => o !== t && o.sheet === t.sheet && ((o.r1 === o.r2 && o.r1 >= t.r1 && o.r1 <= t.r2 && o.c1 === t.c1 && o.c2 === t.c2) || (o.c1 === o.c2 && o.c1 >= t.c1 && o.c1 <= t.c2 && o.r1 === t.r1 && o.r2 === t.r2)),
      );
      if (sibling) {
        const cell = label(t.r1 - 1, t.c1 - 1);
        out.push({ key: t.key, kind: "corner", cell, range, note: `${t.key}: named after ${cell}, the corner of a Create from Selection with both Top row and Left column; it covers the whole block ${range}` });
      }
    }
  }
  return out;
}

export interface LabelDrift {
  /** `Sheet!Name` or `Name`. */
  key: string;
  name: string;
  /** The label cell, `'IS'!A12`. */
  cell: string;
  /** Its text. */
  label: string;
  /** The name the label gives (`labelNames`' first spelling). */
  converted: string;
  /** The cells the name covers, `'IS'!B12:G12`. */
  range: string;
  /** `row`: the label is left of the name's first cell (Left column); `column`: above it (Top row). */
  along: "row" | "column";
  /** Names from the same column (or row) of labels, over the same span, that match their labels. */
  matching: string[];
  /** The label's name when no name has that spelling: probably the name before a rename. */
  renamedFrom?: string;
}

/**
 * Names whose label cell no longer gives them (C15). Conservative: the evidence that Create
 * from Selection made a name is other names in the same column of labels (Left column) or
 * row of labels (Top row), over the same span, whose labels do give them; a block whose
 * names never matched their labels (named by hand) gives nothing. A name is reported when
 * its label is text typed in the cell (a computed label is the pull's note), no label of it
 * gives its name (left, above, the corner), no other name on the same cells matches the
 * label, its cells are not a row or column of text (a header), and its block has at least
 * as many names that match as names that do not.
 */
export function labelDrift(wb: WorkbookSnapshot, names: readonly DefinedName[], values: ReadonlyMap<string, ReadonlyMap<string, CellValue>>): LabelDrift[] {
  const { targets, isComputed } = labelContext(wb, names, true);
  const lower = (s: string) => s.toLowerCase();
  const valueAt = (sheet: string, row: number, col: number): CellValue | undefined => (row >= 1 && col >= 1 ? values.get(sheet)?.get(formatCell({ row, col })) : undefined);
  const gives = (v: CellValue | undefined, name: string) => labelNames(v).some((n) => lower(n) === lower(name));
  const spellings = new Set(names.map((d) => lower(d.name)));
  const cellsKey = (t: Target) => `${lower(t.sheet)}\u0000${t.r1}:${t.c1}:${t.r2}:${t.c2}`;
  const sameCells = new Map<string, Target[]>();
  for (const t of targets) sameCells.set(cellsKey(t), [...(sameCells.get(cellsKey(t)) ?? []), t]);
  const textHeader = (t: Target): boolean => {
    if (t.r1 === t.r2 && t.c1 === t.c2) return false;
    let text = 0;
    for (let r = t.r1; r <= t.r2; r++) {
      for (let c = t.c1; c <= t.c2; c++) {
        const v = valueAt(t.sheet, r, c);
        if (typeof v === "string" && v.trim() !== "") text++;
        else if (v !== undefined && v !== "") return false;
      }
    }
    return text > 0;
  };

  // Per name and direction, its label cell when it holds text typed there, grouped by block:
  // the sheet, the column (or row) of labels, and the span of the names.
  type Entry = { t: Target; along: "row" | "column"; row: number; col: number; label: string; matches: boolean };
  const blocks = new Map<string, Entry[]>();
  const explained = new Set<Target>();
  for (const t of targets) {
    const before: { along: "row" | "column"; row: number; col: number }[] = [];
    if (t.r1 === t.r2) before.push({ along: "row", row: t.r1, col: t.c1 - 1 });
    if (t.c1 === t.c2) before.push({ along: "column", row: t.r1 - 1, col: t.c1 });
    // A block of several rows and columns is a corner's name (the pull's note), not a label's.
    if (before.length === 0 || gives(valueAt(t.sheet, t.r1 - 1, t.c1 - 1), t.d.name)) explained.add(t);
    for (const b of before) {
      const v = valueAt(t.sheet, b.row, b.col);
      const matches = gives(v, t.d.name);
      if (matches) explained.add(t);
      if (typeof v !== "string" || labelName(v) === undefined || isComputed(t.sheet, b.row, b.col)) continue;
      const block = b.along === "row" ? `${lower(t.sheet)}\u0000row\u0000${b.col}\u0000${t.c1}:${t.c2}` : `${lower(t.sheet)}\u0000column\u0000${b.row}\u0000${t.r1}:${t.r2}`;
      blocks.set(block, [...(blocks.get(block) ?? []), { t, along: b.along, row: b.row, col: b.col, label: v, matches }]);
    }
  }

  const out: LabelDrift[] = [];
  const seen = new Set<Target>();
  for (const entries of blocks.values()) {
    const matching = entries.filter((e) => e.matches);
    // Another name on the same cells that the label gives (an alias beside it) is not drift;
    // nor a row (or column) of text, a table's header: the cell before it heads the column of
    // labels ("Excel name" before the peers' names, lbo-ep02's Peers sheet).
    const drifted = entries.filter((e) => !explained.has(e.t) && !(sameCells.get(cellsKey(e.t)) ?? []).some((o) => gives(e.label, o.d.name)) && !textHeader(e.t));
    if (matching.length === 0 || drifted.length === 0 || matching.length < drifted.length) continue;
    for (const e of drifted) {
      if (seen.has(e.t)) continue;
      seen.add(e.t);
      const spelled = labelNames(e.label);
      const t = e.t;
      const d: LabelDrift = {
        key: t.key,
        name: t.d.name,
        cell: `${quote(t.sheet)}!${formatCell({ row: e.row, col: e.col })}`,
        label: e.label,
        converted: spelled[0]!,
        range: `${quote(t.sheet)}!${formatCell({ row: t.r1, col: t.c1 })}${t.r1 === t.r2 && t.c1 === t.c2 ? "" : `:${formatCell({ row: t.r2, col: t.c2 })}`}`,
        along: e.along,
        matching: matching.map((m) => m.t.d.name),
      };
      if (!spelled.some((n) => spellings.has(lower(n)))) d.renamedFrom = spelled[0]!;
      out.push(d);
    }
  }
  return out;
}
