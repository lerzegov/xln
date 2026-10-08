// Moving a formula to another cell, the way Excel fills a shared formula: relative parts
// of A1 references move by the offset, `$`-anchored parts stay. Done on the AST (with a
// token fallback for text that does not parse), so strings, structured references, names
// and sheet names are never touched; only `ref` nodes are rewritten.
//
// Shared formulas store their text once, on the master (`<f t="shared" ref=… si=…>`); a
// child means that text moved by its offset from the master. `shiftFormula` is the
// `ReferenceShifter` that `file/formulaTextAt` takes.

import { render, type Expr } from "./ast.js";
import { parse } from "./parser.js";
import { tokenize, type Qualifier } from "./tokens.js";

const MAX_COL = 16384; // XFD
const MAX_ROW = 1048576;

function colNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function colLetters(col: number): string {
  let s = "";
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function isLetter(c: string | undefined): boolean {
  return c !== undefined && ((c >= "A" && c <= "Z") || (c >= "a" && c <= "z"));
}
function isDigit(c: string | undefined): boolean {
  return c !== undefined && c >= "0" && c <= "9";
}

/**
 * Moves one end of a reference: a cell (`$A1`), a column (`$C`) or a row (`7`). Returns
 * undefined when the result falls off the sheet. Case and `$` are kept as written.
 */
function shiftEnd(end: string, dRow: number, dCol: number): string | undefined {
  let i = 0;
  let out = "";
  // column part
  const colAbs = end[i] === "$" && isLetter(end[i + 1]);
  if (colAbs) i++;
  const c0 = i;
  while (isLetter(end[i])) i++;
  if (i > c0) {
    const letters = end.slice(c0, i);
    if (colAbs || dCol === 0) out += (colAbs ? "$" : "") + letters;
    else {
      const col = colNumber(letters) + dCol;
      if (col < 1 || col > MAX_COL) return undefined;
      out += colLetters(col);
    }
  } else if (colAbs) return end; // not a reference end we know: leave it
  // row part
  if (i === end.length) return out;
  const rowAbs = end[i] === "$";
  if (rowAbs) i++;
  const r0 = i;
  while (isDigit(end[i])) i++;
  if (i === r0 || i !== end.length) return end;
  const digits = end.slice(r0, i);
  if (rowAbs || dRow === 0) return out + (rowAbs ? "$" : "") + digits;
  const row = Number(digits) + dRow;
  if (row < 1 || row > MAX_ROW) return undefined;
  return out + String(row);
}

/** The range operator inside an area address (`:`, `:.`, `.:`, `.:.`) and where it sits. */
function splitArea(address: string): { a: string; op: string; b: string } | undefined {
  const k = address.indexOf(":");
  if (k < 0) return undefined;
  const start = address[k - 1] === "." ? k - 1 : k;
  const end = address[k + 1] === "." ? k + 2 : k + 1;
  return { a: address.slice(0, start), op: address.slice(start, end), b: address.slice(end) };
}

/**
 * Moves an A1 address (no qualifier): a cell, an area, whole columns or whole rows.
 * Returns undefined when any end falls off the sheet (the reference becomes `#REF!`).
 */
export function shiftAddress(address: string, dRow: number, dCol: number): string | undefined {
  if (address.startsWith("#")) return address; // `Sheet!#REF!`
  const parts = splitArea(address);
  if (!parts) return shiftEnd(address, dRow, dCol);
  const a = shiftEnd(parts.a, dRow, dCol);
  const b = shiftEnd(parts.b, dRow, dCol);
  return a === undefined || b === undefined ? undefined : a + parts.op + b;
}

function refText(qual: Qualifier | undefined, address: string | undefined): string {
  return (qual?.raw ?? "") + (address ?? "#REF!");
}

/**
 * The formula Excel means at a cell `dRow` rows down and `dCol` columns right of where
 * `text` is written: relative parts of references move, `$` parts do not. A reference
 * moved off the sheet becomes `#REF!` (keeping its sheet: `S2!#REF!`), and a spill of it
 * (`A1#`) becomes `#REF!` too. Names, strings, structured references, function names and
 * LET/LAMBDA variables are left alone; references on other sheets and in other workbooks
 * move like any other (Excel fills them the same way). Layout is kept.
 */
export function shiftFormula(text: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) return text;
  let f;
  try {
    f = parse(text);
  } catch {
    return shiftTokens(text, dRow, dCol);
  }
  const body = render(f.src, f.body, (node: Expr) => {
    if (node.kind === "ref") return refText(node.qual, shiftAddress(node.address, dRow, dCol));
    if (node.kind === "postfix" && node.op === "#" && node.operand.kind === "ref") {
      const r = node.operand;
      if (shiftAddress(r.address, dRow, dCol) === undefined) return refText(r.qual, undefined);
    }
    return undefined;
  });
  return text.slice(0, f.body.span.start) + body + text.slice(f.body.span.end);
}

/** Same as `shiftFormula` for text the parser rejects: the tokenizer still finds the references. */
function shiftTokens(text: string, dRow: number, dCol: number): string {
  let out = "";
  for (const t of tokenize(text)) {
    if (t.kind === "ref" && t.value !== undefined) out += refText(t.qual, shiftAddress(t.value, dRow, dCol));
    else out += t.text;
  }
  return out;
}
