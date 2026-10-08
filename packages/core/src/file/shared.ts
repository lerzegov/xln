// Text of a shared-formula child. The file stores the text once, on the master; each
// child means the same formula with its relative references moved by the child's
// offset from the master. Moving references needs the formula tokenizer (lang/), so
// this layer takes the shifter as a parameter instead of depending on it.
import { parseCell } from "./cellref.js";
import type { CellFormula, Sheet } from "./types.js";

/**
 * Rewrites `masterText` for a cell `rows` down and `cols` right of the master:
 * relative references move, absolute ones (`$A$1`) and names do not.
 */
export type ReferenceShifter = (masterText: string, rows: number, cols: number) => string;

/**
 * The formula text that applies at `formula.cell`: its own text for anything but a shared
 * child; for a child, the master's text passed through `shift`.
 *
 * The shifter is `shiftFormula` from `lang/` (this layer does not depend on the formula
 * layer). `(t) => t` gives the master's text unshifted, which is enough to tell which
 * names a child reads: names do not move in a fill.
 */
export function formulaTextAt(sheet: Sheet, formula: CellFormula, shift: ReferenceShifter): string | undefined {
  if (formula.kind !== "shared-child") return formula.text;
  if (formula.text !== undefined) return formula.text;
  const group = sheet.sharedFormulas.find((g) => g.si === formula.si);
  if (!group?.master || group.text === undefined) return undefined;
  const m = parseCell(group.master);
  const c = parseCell(formula.cell);
  if (!m || !c) return undefined;
  return shift(group.text, c.row - m.row, c.col - m.col);
}
