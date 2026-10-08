// The label notice (decided 2026-10-07): after a rename, the text cells that still read the
// old name, and how to fix them in Excel with Find & Replace. xln never writes cell values
// (D7: Excel owns labels), so it says exactly what to type instead. The author's models
// keep per row a readable text ("Gross income"), the name's text ("Gross_income", what
// Create from Selection reads) and the cells; a rename leaves one or more texts stale.
//
// Conservative and cheap: only the sheet of the name's cells, only its rows and columns
// (the label left of a row, above a column, a description further left), only text typed
// in a cell (a formula's result changes with the formula, Find & Replace would not reach
// it). A cell counts when its whole text gives the old name by Create from Selection's
// conversion (`labelNames`); one whose words resemble the old name is listed to check by
// eye, never put in the replace.
//
// Excel's Replace dialog (support.microsoft.com, "Find or replace text and numbers on a
// worksheet" and "Keyboard shortcuts in Excel", read 2026-10-07): Ctrl+H on Windows,
// Ctrl+H or Cmd+Shift+H on Mac, or Home > Find & Select > Replace; Within Sheet or
// Workbook; on the Replace tab Look in offers Formulas only; Match entire cell contents.
// With Match entire cell contents a formula cell never matches (its content starts with `=`).

import { formatCell, parseCell } from "../file/cellref.js";
import type { WorkbookSnapshot } from "../file/types.js";
import type { CellValue } from "../build/verify.js";
import type { Change, RenameName } from "../build/changes.js";
import { quoteSheet } from "../lang/tokens.js";
import { labelContext, labelConversion, labelNames } from "./labels.js";

export interface LabelRename {
  /** The name before the rename. */
  from: string;
  to: string;
  /** The name's scope before the rename: a sheet's name, or null for the workbook. */
  scope: string | null;
}

export interface LabelCell {
  /** `IS!B5`, `'Cash Flow'!B5`. */
  cell: string;
  text: string;
}

export interface LabelReplace {
  /** The text to find: a stale cell's whole text. */
  find: string;
  /** The same text for the new name, in the label's style (spaces kept). */
  replace: string;
  cells: string[];
  /** Other cells of the sheet with this very text (outside the name's rows and columns): Replace All would change them too. */
  alsoOnSheet: string[];
}

export interface LabelNotice {
  from: string;
  to: string;
  /** The sheet of the name's cells. */
  sheet: string;
  /** Cells whose whole text gives the old name. */
  stale: LabelCell[];
  /** One Find & Replace pair per distinct text of `stale`. */
  pairs: LabelReplace[];
  /** Cells whose words resemble the old name, to check by eye: never in a pair. */
  maybe: LabelCell[];
}

const lower = (s: string) => s.toLowerCase();

/** The words of a text or a name, lower case: `Gross_ind. income` → gross, ind, income. */
function words(s: string): string[] {
  const out: string[] = [];
  let w = "";
  for (const c of s.toLowerCase()) {
    if (/[\p{L}\p{N}]/u.test(c)) w += c;
    else if (w) {
      out.push(w);
      w = "";
    }
  }
  if (w) out.push(w);
  return out;
}

function containsRun(hay: readonly string[], needle: readonly string[]): boolean {
  for (let i = 0; i + needle.length <= hay.length; i++) if (needle.every((w, k) => hay[i + k] === w)) return true;
  return false;
}

/** A text gives the name by Create from Selection's conversion (any of its readings, case ignored). */
export function labelGives(text: string, name: string): boolean {
  return labelNames(text).some((n) => lower(n) === lower(name));
}

/**
 * The text a label should read for the renamed name, in the label's own style: "Gross
 * income" (from Gross_income) becomes "Gross ind income" for Gross_ind_income; a label
 * written as the name itself becomes the new name. Whatever it returns gives `to` by the
 * conversion; when no style can be kept, it is `to` itself.
 */
export function labelReplacement(text: string, from: string, to: string): string {
  const t = text.trim();
  if (lower(t) === lower(from)) return to;
  const conv = labelConversion(t);
  let candidate: string | undefined;
  let aligned = false;
  if (conv !== undefined && lower(conv.name) === lower(from)) {
    aligned = true;
    // The label's characters line up with the name's, past a `_` Excel put in front (a
    // leading digit) and short of one it put behind (`Q1_`) or characters it dropped at the end.
    const label = [...t];
    const name = [...from];
    const offset = conv.front ? 1 : 0;
    const body = name.length - (conv.back ? 1 : 0);
    // Which character the label has where the name has `_`: one kind only, or no style to keep.
    const seps = new Set<string>();
    for (let i = offset; i < body; i++) if (name[i] === "_") seps.add(label[i - offset]!);
    const sep = seps.size === 1 ? [...seps][0]! : seps.size === 0 ? "_" : undefined;
    if (sep !== undefined) {
      let core = to;
      if (conv.front && core.startsWith("_")) core = core.slice(1);
      if (conv.back && core.endsWith("_")) core = core.slice(0, -1);
      // Try the label's style without the added `_`, then with the name as it is.
      for (const c of [core.split("_").join(sep), to.split("_").join(sep)]) {
        if (c !== "" && labelGives(c, to)) {
          candidate = c;
          break;
        }
      }
    }
  }
  // A run of spaces Excel made one `_` (the text does not line up with the name).
  if (!aligned && !t.includes("_") && t.includes(" ")) candidate = to.split("_").join(" ").trim();
  return candidate !== undefined && candidate !== "" && labelGives(candidate, to) ? candidate : to;
}

/**
 * The label notice for one rename, on the workbook before the build (its labels are what
 * the build leaves) and its cells' values. Undefined when the name is not on one sheet's
 * cells (a LAMBDA, a constant) or no text cell has anything to say.
 */
export function labelNotice(wb: WorkbookSnapshot, values: ReadonlyMap<string, ReadonlyMap<string, CellValue>>, rename: LabelRename): LabelNotice | undefined {
  const d = wb.definedNames.find(
    (n) => lower(n.name) === lower(rename.from) && (rename.scope === null ? n.scope.kind === "workbook" : n.scope.kind === "sheet" && lower(n.scope.name) === lower(rename.scope)),
  );
  if (!d) return undefined;
  const { targets, isComputed } = labelContext(wb, [d], true);
  const t = targets[0];
  if (!t) return undefined;
  const cells = [...values.entries()].find(([s]) => lower(s) === lower(t.sheet))?.[1];
  if (!cells) return undefined;
  const at = (address: string) => `${quoteSheet(t.sheet)}!${address}`;
  const oldWords = words(rename.from);

  const stale: (LabelCell & { row: number; col: number })[] = [];
  const maybe: (LabelCell & { row: number; col: number })[] = [];
  const typed: { address: string; text: string }[] = [];
  for (const [address, v] of cells) {
    if (typeof v !== "string" || v.trim() === "") continue;
    const p = parseCell(address);
    if (!p || isComputed(t.sheet, p.row, p.col)) continue;
    typed.push({ address, text: v });
    const inRows = p.row >= t.r1 && p.row <= t.r2;
    const inCols = p.col >= t.c1 && p.col <= t.c2;
    // The name's own cells are its data, not its labels; the corner is the two-way label.
    const corner = p.row === t.r1 - 1 && p.col === t.c1 - 1;
    if ((inRows && inCols) || !(inRows || inCols || corner)) continue;
    const c = { cell: at(formatCell(p)), text: v, row: p.row, col: p.col };
    if (labelGives(v, rename.from)) stale.push(c);
    else if (oldWords.length > 0) {
      const w = words(v);
      if (w.join(" ") === oldWords.join(" ") || (oldWords.length >= 2 && containsRun(w, oldWords))) maybe.push(c);
    }
  }
  if (stale.length === 0 && maybe.length === 0) return undefined;
  const order = (a: { row: number; col: number }, b: { row: number; col: number }) => a.row - b.row || a.col - b.col;
  stale.sort(order);
  maybe.sort(order);

  const pairs: LabelReplace[] = [];
  const staleCells = new Set(stale.map((s) => s.cell));
  for (const s of stale) {
    const p = pairs.find((x) => lower(x.find) === lower(s.text));
    if (p) p.cells.push(s.cell);
    else pairs.push({ find: s.text, replace: labelReplacement(s.text, rename.from, rename.to), cells: [s.cell], alsoOnSheet: [] });
  }
  // Excel matches the whole content, case ignored (Match case off), whitespace as is.
  for (const p of pairs) {
    p.alsoOnSheet = typed
      .filter((c) => lower(c.text) === lower(p.find) && !staleCells.has(at(c.address)))
      .map((c) => ({ c, p: parseCell(c.address)! }))
      .sort((a, b) => order({ row: a.p.row, col: a.p.col }, { row: b.p.row, col: b.p.col }))
      .map((x) => at(x.c.address));
  }
  const strip = ({ cell, text }: LabelCell): LabelCell => ({ cell, text });
  return { from: rename.from, to: rename.to, sheet: t.sheet, stale: stale.map(strip), pairs, maybe: maybe.map(strip) };
}

/** How to open Excel's Replace dialog, per platform (Microsoft's shortcut list). */
export const EXCEL_REPLACE_KEYS = "Ctrl+H on Windows; Ctrl+H or ⌘⇧H on Mac; or Home → Find & Select → Replace";

const list = (cells: readonly string[], max = 8) => cells.slice(0, max).join(", ") + (cells.length > max ? `, … (${cells.length - max} more)` : "");

/**
 * The notice as text, the one form every surface shows (CLI, MCP, the editor's output):
 * the stale cells, the Find & Replace pairs with the dialog's settings, the cells to check by eye.
 */
export function labelNoticeLines(n: LabelNotice): string[] {
  const out: string[] = [];
  const head = `${n.from} → ${n.to}`;
  if (n.stale.length) {
    out.push(`${head}: ${n.stale.length === 1 ? "1 label still reads" : `${n.stale.length} labels still read`} the old name: ${list(n.stale.map((s) => s.cell))}`);
    out.push(`In Excel, on sheet ${n.sheet}: Find & Replace (${EXCEL_REPLACE_KEYS})`);
    const many = n.pairs.length > 1;
    n.pairs.forEach((p, i) => {
      const lead = many ? `${i + 1}. ` : "";
      const pad = " ".repeat(lead.length);
      out.push(`  ${lead}Find what:     ${p.find}`);
      out.push(`  ${pad}Replace with:  ${p.replace}`);
    });
    out.push(`  Within: Sheet · Look in: Formulas · Match entire cell contents ✓ · then Replace All${many ? " (for each pair)" : ""}`);
    for (const p of n.pairs) {
      if (p.alsoOnSheet.length) out.push(`  Replace All also changes ${list(p.alsoOnSheet)} ("${p.find}", away from the name's rows and columns): to leave ${p.alsoOnSheet.length === 1 ? "it" : "them"}, use Find Next and Replace instead`);
    }
  } else out.push(`${head}: no label reads the old name exactly`);
  if (n.maybe.length) out.push(`Check by eye (not in the replace): ${n.maybe.slice(0, 8).map((m) => `${m.cell} "${m.text}"`).join(", ")}${n.maybe.length > 8 ? `, … (${n.maybe.length - 8} more)` : ""}`);
  return out;
}

/** The one-line form (C15's hint): `in Excel, Find & Replace "Gross income" with "Gross ind income" (Ctrl+H; Match entire cell contents)`. */
export function labelReplaceShort(find: string, replace: string): string {
  return `in Excel, Find & Replace "${find}" with "${replace}" (Ctrl+H, ⌘⇧H on Mac; Match entire cell contents)`;
}

/**
 * The notices of a build's renames, on the workbook before the build (the build leaves its
 * labels as they are). Only renames with something to say; values unreadable, none.
 */
export function renameLabelNotices(wb: WorkbookSnapshot, values: () => ReadonlyMap<string, ReadonlyMap<string, CellValue>> | undefined, changes: readonly Change[]): LabelNotice[] {
  const renames = changes.filter((c): c is RenameName => c.op === "rename-name");
  if (renames.length === 0) return [];
  const v = values();
  if (!v) return [];
  const out: LabelNotice[] = [];
  for (const r of renames) {
    const n = labelNotice(wb, v, { from: r.from, to: r.to, scope: r.scope });
    if (n) out.push(n);
  }
  return out;
}
