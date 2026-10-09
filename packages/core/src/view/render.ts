// The formula view as a plain-text document: one entry per formula, in columns that stay
// put down the page (defined names on the cell, cell, kind, formula, saved value). The
// name column reads like the left-hand side of an equation: what the block is, then
// where it sits. Names too long for the column go on a line of their own above the entry. A formula that does not fit its
// column, or that Excel stored on several lines, gets a block of its own: its first line
// stays in the formula column, the rest is indented under it, pretty-printed (LET
// bindings one per line, long argument lists one argument per line) when it was a
// single long line. An empty row in the sheet becomes an empty line, so the sheet's own
// sections show.
//
// Positions of names and references come back as document offsets, for an editor to
// link them.

import { prettyPrint } from "../lang/format.js";
import { quoteSheet } from "../lang/tokens.js";
import type { FormulaViewLine } from "./formulas.js";

export interface RenderOptions {
  /** The sheet, for the header. */
  sheet: string;
  /** The workbook's file name, for the header. */
  workbook?: string;
  /** Longest formula kept on the entry's line (default 64); longer ones get a block. */
  formulaWidth?: number;
  /** Line width that block formulas are pretty-printed to (default 110). */
  width?: number;
  /** What the document's header says about the order (default: order of appearance). */
  order?: string;
  /** Show the defined names on each cell as a first column (default true). */
  nameColumn?: boolean;
  /** Widest name column (default 32); longer name lists go on a line above their entry. */
  nameWidth?: number;
  /** Lines from several sheets: each address is written with its sheet (`BS!C6#`), and the header names the workbook. */
  sheets?: boolean;
  /**
   * Calculation order: a first column with each line's `level` (`↻n` after it for a member
   * of circular reference n), and a comment line before each circular reference.
   */
  levels?: boolean;
}

/** A span in the rendered document (offsets into `text`). */
export interface RenderedSpan {
  start: number;
  end: number;
}

export interface RenderedEntry {
  /** Index into the input lines. */
  index: number;
  cell: string;
  /** 0-based document line of the entry's first line (its names, when they sit above it), and of its last. */
  line: number;
  lastLine: number;
  /** Document offset of the entry's first character: its first name, or its address. */
  start: number;
  /** `lhs[k]` of the input line (the defined names on the cell), placed in the document; empty when not shown. */
  lhs: RenderedSpan[];
  /** The cell address (after the name column). */
  address: RenderedSpan;
  /** Where the formula text starts and ends in the document. */
  formula: RenderedSpan;
  /** `names[k]` and `refs[k]` of the input line, placed in the document. */
  names: RenderedSpan[];
  refs: RenderedSpan[];
}

export interface RenderedView {
  text: string;
  entries: RenderedEntry[];
}

/** Kind column: the saved size of a spill `(r×c)` or legacy array `{r×c}`, shared groups, data tables. */
/** Shown in place of the value of a formula Excel has not calculated since a build wrote it. */
export const UNCALCULATED = "(not calculated since the build)";

export function kindTag(l: FormulaViewLine): string {
  const size = l.rows !== undefined && l.cols !== undefined ? `${l.rows}×${l.cols}` : "";
  switch (l.kind) {
    case "dynamic-array":
      // Not calculated since the build: the saved extent is the anchor alone, not a size.
      return size && !l.uncalculated ? `(${size})` : "(spill)";
    case "array":
      return size ? `{${size}}` : "{array}";
    case "data-table":
      return size ? `table ${size}` : "table";
    case "shared":
      return l.master === undefined || l.master === l.cell ? `shared ×${l.groupSize ?? "?"}` : `shared ← ${l.master}`;
    default:
      return "";
  }
}

/** Cell column: a dynamic array's anchor reads `C6#`, the reference to its whole spill. */
export function addressLabel(l: FormulaViewLine): string {
  return l.kind === "dynamic-array" ? l.cell + "#" : l.cell;
}

/** A sheet as written in front of `!` in labels: quoted only when Excel would need it to read it back. */
export function sheetLabel(sheet: string): string {
  return quoteSheet(sheet);
}

/** Level column: `5`, or `5 ↻2` for a member of circular reference 2. */
export function levelTag(l: FormulaViewLine): string {
  if (l.level === undefined) return "";
  return l.cycle !== undefined ? `${l.level} ↻${l.cycle}` : String(l.level);
}

function nonWsIndex(s: string): number[] {
  const out: number[] = [];
  for (let i = 0; i < s.length; i++) {
    const c = s[i]!;
    if (c !== " " && c !== "\t" && c !== "\n" && c !== "\r" && c !== " ") out.push(i);
  }
  return out;
}

/**
 * The text shown for a formula and a map from offsets in `l.formula` to offsets in it.
 * Pretty-printing only moves layout whitespace, so each non-blank character keeps its rank.
 */
function layout(l: FormulaViewLine, fits: boolean, width: number): { text: string; map: (i: number) => number } {
  const id = { text: l.formula, map: (i: number) => i };
  if (fits || l.formula.includes("\n") || l.error !== undefined || l.stored === "") return id;
  let pretty: string;
  try {
    pretty = prettyPrint(l.formula, { width, pack: true });
  } catch {
    return id;
  }
  const a = nonWsIndex(l.formula);
  const b = nonWsIndex(pretty);
  if (a.length !== b.length || a.some((i, k) => l.formula[i] !== pretty[b[k]!])) return id;
  const rank = new Map<number, number>();
  a.forEach((i, k) => rank.set(i, k));
  return {
    text: pretty,
    map: (i) => {
      const k = rank.get(i);
      return k === undefined ? i : b[k]!;
    },
  };
}

/** Lines of the document's header comment. */
const HEADER = 4;

function pad(s: string, n: number): string {
  return s.length >= n ? s : s + " ".repeat(n - s.length);
}

function plural(n: number, one: string): string {
  return `${n} ${one}${n === 1 ? "" : "s"}`;
}

/** Renders the view of one sheet. */
export function renderFormulaView(lines: readonly FormulaViewLine[], opts: RenderOptions): RenderedView {
  const fw = opts.formulaWidth ?? 64;
  const width = opts.width ?? 110;
  const order = opts.order ?? "order of appearance (row by row, left to right)";
  const out: string[] = [];
  let offset = 0;
  const push = (s: string): void => {
    out.push(s);
    offset += s.length + 1;
  };

  const of = opts.workbook ? ` of ${opts.workbook}` : "";
  const levels = opts.levels ?? false;
  if (opts.sheets) {
    const n = new Set(lines.map((l) => l.sheet)).size;
    push(`// Workbook${opts.workbook ? " " + opts.workbook : ""}: ${plural(lines.length, "formula")} on ${plural(n, "sheet")} in ${order}. Read-only view.`);
  } else push(`// Sheet ${opts.sheet}: ${plural(lines.length, "formula")} in ${order}. Read-only view${of}.`);
  if (levels) {
    push("// level · names · cell · kind · formula · saved value. Level: the longest chain of formulas from the inputs");
    push("// (1: reads inputs only); ↻n: part of circular reference n. Names: those defined as the cell, its spill (C6#)");
    push("// or its saved extent. C6# (6×1): dynamic array and its saved size; {r×c} legacy array; shared: shared formula.");
  } else {
    push("// names · cell · kind · formula · saved value. Names: those defined as the cell, its spill (C6#) or its saved");
    push("// extent. C6# (6×1): dynamic array and the size it spilled to when saved; {r×c} legacy array;");
    push("// shared ×n / shared ← B2: shared formula and its master; table: data table.");
  }
  if (lines.length === 0) {
    push("");
    push("// No formulas on this sheet.");
  }

  const showNames = opts.nameColumn ?? true;
  const lhs = lines.map((l) => (showNames ? (l.lhs ?? []).map((n) => n.display).join(", ") : ""));
  const cap = opts.nameWidth ?? 32;
  const nw = Math.max(0, ...lhs.filter((s) => s.length <= cap).map((s) => s.length));
  const nameCol = nw > 0 ? nw + 2 : 0;
  const prefixes = lines.map((l) => (opts.sheets ? sheetLabel(l.sheet) + "!" : ""));
  const labels = lines.map((l, k) => prefixes[k] + addressLabel(l));
  const levelTags = lines.map((l) => (levels ? levelTag(l) : ""));
  const lw = levels ? Math.max(0, ...levelTags.map((s) => s.length)) + 2 : 0;
  const tags = lines.map(kindTag);
  const aw = Math.max(0, ...labels.map((s) => s.length));
  const tw = Math.max(0, ...tags.map((s) => s.length));
  const formulaCol = lw + nameCol + aw + 2 + (tw > 0 ? tw + 2 : 0);
  const inline = lines.map((l) => !l.formula.includes("\n") && l.formula.length <= fw);
  const widest = Math.max(0, ...lines.filter((_, k) => inline[k]).map((l) => l.formula.length));
  const valueCol = formulaCol + 2 + Math.min(fw, widest) + 2;
  const indent = " ".repeat(formulaCol + 2);

  const entries: RenderedEntry[] = [];
  let prevRow: number | undefined;
  let prevSheet: string | undefined;
  let prevCycle: number | undefined;
  let prevBlock = false;
  lines.forEach((l, index) => {
    const shown = layout(l, inline[index]!, Math.max(40, width - indent.length));
    const parts = shown.text.split("\n");
    // A formula on several lines stands apart; one that is merely long stays in the list.
    const block = parts.length > 1;
    // A gap in the rows (or, in calculation order, a step back or another sheet) starts a section.
    const jump = prevRow !== undefined && (l.row - prevRow > 1 || l.row < prevRow || l.sheet !== prevSheet);
    const cycleStart = levels && l.cycle !== undefined && l.cycle !== prevCycle;
    const blank = out[out.length - 1] === "";
    if (!blank && (out.length === HEADER || jump || block || prevBlock || cycleStart)) push("");
    if (cycleStart) {
      const members = lines.filter((x) => x.cycle === l.cycle).length;
      push(`// ↻${l.cycle}: circular reference, ${plural(members, "formula")} here that depend on each other (Excel iterates them or reports it)`);
    }
    prevRow = l.row;
    prevSheet = l.sheet;
    prevCycle = levels ? l.cycle : undefined;
    prevBlock = block;

    const names = lhs[index]!;
    const above = names.length > nw;
    const firstLine = out.length;
    const entryStart = offset;
    if (above) push(names);
    const level = lw > 0 ? pad(levelTags[index]!, lw - 2) + "  " : "";
    const head = level + (nameCol > 0 ? pad(above ? "" : names, nw) + "  " : "") + pad(labels[index]!, aw) + "  " + (tw > 0 ? pad(tags[index]!, tw) + "  " : "") + "= ";
    const value = l.uncalculated ? `→ ${UNCALCULATED}` : l.valueText !== undefined ? "→ " + l.valueText + (l.rows !== undefined && l.rows * (l.cols ?? 1) > 1 ? " …" : "") : "";
    const note = l.error !== undefined ? `  // does not parse: ${l.error}` : "";

    const starts: number[] = [];
    const startOffset = offset;
    parts.forEach((p, k) => {
      let s = (k === 0 ? head : indent) + p;
      starts.push(offset + (k === 0 ? head.length : indent.length));
      if (k === 0 && value && (parts.length === 1 || s.length < valueCol - 1)) s = (s.length < valueCol - 1 ? pad(s, valueCol) : s + "  ") + value;
      if (k === parts.length - 1) s += note;
      push(k === 0 ? s.trimEnd() || s : s);
    });
    if (value && parts.length > 1 && (head + parts[0]).length >= valueCol - 1) push(" ".repeat(valueCol) + value);

    // Formula offset → document offset, through the lines it was split into.
    const lineStart: number[] = [];
    let acc = 0;
    for (const p of parts) {
      lineStart.push(acc);
      acc += p.length + 1;
    }
    const place = (i: number): number => {
      const j = shown.map(i);
      let k = lineStart.length - 1;
      while (k > 0 && lineStart[k]! > j) k--;
      return starts[k]! + (j - lineStart[k]!);
    };
    const span = (s: { start: number; end: number }): RenderedSpan => ({ start: place(s.start), end: s.end > s.start ? place(s.end - 1) + 1 : place(s.start) });
    const lastPart = parts[parts.length - 1]!;
    // The names start the entry, on a line of their own or in the name column.
    const lhsSpans: RenderedSpan[] = [];
    if (names.length > 0) {
      let at = above ? entryStart : entryStart + lw;
      for (const n of l.lhs) {
        lhsSpans.push({ start: at, end: at + n.display.length });
        at += n.display.length + 2;
      }
    }
    const addressStart = startOffset + lw + nameCol + prefixes[index]!.length;
    entries.push({
      index,
      cell: l.cell,
      line: firstLine,
      lastLine: out.length - 1,
      start: entryStart,
      lhs: lhsSpans,
      address: { start: addressStart, end: addressStart + l.cell.length },
      formula: { start: starts[0]!, end: starts[starts.length - 1]! + lastPart.length },
      names: l.names.map(span),
      refs: l.refs.map(span),
    });
  });
  return { text: out.join("\n") + "\n", entries };
}
