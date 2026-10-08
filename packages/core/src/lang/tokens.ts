// Tokenizer for Excel formulas, in stored form (`_xlfn.SEQUENCE(…)`) and in display form
// (`SEQUENCE(…)`). One scanner for both: prefixes are just part of an identifier.
//
// Every token keeps its source span, and whitespace is kept as a token too, so callers
// can splice edits into the original text and keep the author's layout.

import { lookupFunction } from "./catalogue.js";

export interface Span {
  start: number;
  end: number;
}

/** A sheet or workbook qualifier in front of `!`: `Sheet1!`, `'It''s'!`, `[1]Sheet!`, `[0]!`, `S1:S3!`. */
export interface Qualifier {
  /** Text of the whole qualifier as written, including the `!`. */
  raw: string;
  /** External book part as written, with its brackets (`[1]`, `[Book.xlsx]`) or a path before it. */
  book?: string;
  /** First sheet, unquoted and un-escaped. Absent for `[0]!Name`. */
  sheet?: string;
  /** Last sheet of a 3-D reference (`S1:S3!A1`). */
  sheet2?: string;
  /** Whether the qualifier was written in single quotes. */
  quoted: boolean;
}

export type TokenKind =
  | "ws" // whitespace that is only layout
  | "isect" // whitespace that is the intersection operator
  | "number"
  | "string"
  | "bool"
  | "error"
  | "ref" // cell, area, whole columns or rows, possibly qualified; or a qualified #REF!
  | "name" // identifier: defined name, function name, LET/LAMBDA variable, possibly qualified
  | "structref" // structured reference: `Tbl[Col]`, `[@Col]`, `Tbl[[#Headers],[A]:[B]]`
  | "op" // + - * / ^ & = <> < > <= >= % : :. .: .:. # @
  | "("
  | ")"
  | "{"
  | "}"
  | ","
  | ";"
  | "invalid"
  | "eof";

export interface Token extends Span {
  kind: TokenKind;
  text: string;
  /** For `ref`, `name`, `error`: the qualifier, when present. */
  qual?: Qualifier;
  /** For `ref`: the address without qualifier. For `name`: the identifier without qualifier.
   *  For `structref`: the table name ("" when omitted). For `string`: the decoded value. */
  value?: string;
  /** For `structref`: the text between the outer brackets. */
  inner?: string;
  /** For `ref`: what kind of address. */
  refKind?: "cell" | "area" | "cols" | "rows" | "error";
  /** For `invalid`: why. */
  message?: string;
}

export const ERROR_LITERALS = [
  "#NULL!",
  "#DIV/0!",
  "#VALUE!",
  "#REF!",
  "#NAME?",
  "#NUM!",
  "#N/A",
  "#GETTING_DATA",
  "#SPILL!",
  "#CALC!",
  "#FIELD!",
  "#BLOCKED!",
  "#CONNECT!",
  "#UNKNOWN!",
  "#BUSY!",
  "#PYTHON!",
  "#EXTERNAL!",
] as const;

const MAX_COL = 16384; // XFD
const MAX_ROW = 1048576;

const LETTER = /\p{L}/u;
const ALNUM = /[\p{L}\p{N}\p{M}]/u;

function isDigit(c: string | undefined): boolean {
  return c !== undefined && c >= "0" && c <= "9";
}
function isAsciiLetter(c: string | undefined): boolean {
  return c !== undefined && ((c >= "A" && c <= "Z") || (c >= "a" && c <= "z"));
}
function isWs(c: string | undefined): boolean {
  return c === " " || c === "\t" || c === "\r" || c === "\n" || c === "\u00a0";
}
/** Characters that may start an identifier (a defined name, a function, a sheet). */
function isWordStart(c: string | undefined): boolean {
  return c !== undefined && (c === "_" || c === "\\" || c === "$" || LETTER.test(c));
}
/** Characters that may continue an identifier, plus `$` for absolute references. */
function isWordChar(c: string | undefined): boolean {
  return (
    c !== undefined &&
    (c === "_" || c === "\\" || c === "." || c === "?" || c === "$" || ALNUM.test(c))
  );
}

function colNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/** `$A$1`, `B7`, `XFD1048576`: the parts of a single-cell A1 reference, or null. */
export function parseCell(s: string): { col: string; row: string } | null {
  let i = 0;
  if (s[i] === "$") i++;
  const c0 = i;
  while (isAsciiLetter(s[i])) i++;
  const col = s.slice(c0, i);
  if (col.length < 1 || col.length > 3 || colNumber(col) > MAX_COL) return null;
  if (s[i] === "$") i++;
  const r0 = i;
  while (isDigit(s[i])) i++;
  const row = s.slice(r0, i);
  if (i !== s.length || row.length === 0 || row[0] === "0" || Number(row) > MAX_ROW) return null;
  return { col, row };
}

function isColumn(s: string): boolean {
  const t = s.startsWith("$") ? s.slice(1) : s;
  if (t.length < 1 || t.length > 3) return false;
  for (const ch of t) if (!isAsciiLetter(ch)) return false;
  return colNumber(t) <= MAX_COL;
}

function isRow(s: string): boolean {
  const t = s.startsWith("$") ? s.slice(1) : s;
  if (t.length === 0 || t[0] === "0") return false;
  for (const ch of t) if (!isDigit(ch)) return false;
  return Number(t) <= MAX_ROW;
}

/**
 * Whether a sheet name must be quoted in front of `!`: Excel's rule, the one rule xln
 * writes everywhere (formulas stored and shown, `@sheet(…)`, addresses; spec §14 issue 7).
 * Measured on every Excel-saved workbook of the probes and the corpus (2026-10-07): Excel
 * writes `BS!`, `IS!`, `SCF!`, `FCF!` bare in the cell formulas it saves (1,000+ references,
 * none quoted) and `'S1'!`, `'S2'!` (cell-like), `'SCF recursive'!`, `'0 Tables'!` quoted.
 * So a name that looks like a column needs no quotes; one that looks like a cell does, and
 * so (not measured, as Excel's documentation has it) does an R1C1-like one (`R`, `C`,
 * `R1C1`, `RC2`) and `TRUE`/`FALSE`. Anything but letters, digits, `_` and `.`, or a first
 * character that is not a letter or `_`, needs quotes.
 */
export function sheetNeedsQuotes(sheet: string): boolean {
  if (sheet.length === 0) return true;
  if (!(sheet[0] === "_" || LETTER.test(sheet[0]!))) return true;
  for (const ch of sheet) if (!(ch === "_" || ch === "." || ALNUM.test(ch))) return true;
  if (parseCell(sheet)) return true;
  if (looksR1C1(sheet)) return true;
  const up = sheet.toUpperCase();
  return up === "TRUE" || up === "FALSE";
}

function looksR1C1(s: string): boolean {
  const u = s.toUpperCase();
  let i = 0;
  const num = () => {
    while (isDigit(u[i])) i++;
  };
  if (u[i] === "R") {
    i++;
    num();
    if (i === u.length) return true;
  }
  if (u[i] === "C") {
    i++;
    num();
    return i === u.length;
  }
  return false;
}

/** Writes a sheet name as it must appear before `!`. `always` quotes even when not needed. */
export function quoteSheet(sheet: string, always = false): string {
  return always || sheetNeedsQuotes(sheet) ? "'" + sheet.split("'").join("''") + "'" : sheet;
}

const OPS3 = [".:."];
const OPS2 = ["<>", "<=", ">=", ":.", ".:"];
const OPS1 = "+-*/^&=<>%:#@";

/** Tokenize a formula. The result always ends with an `eof` token. Never throws: bad input
 *  becomes `invalid` tokens, which the parser reports with their position. */
export function tokenize(src: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  const n = src.length;

  const push = (t: Omit<Token, "text"> & { text?: string }): void => {
    out.push({ ...t, text: t.text ?? src.slice(t.start, t.end) } as Token);
  };

  /** Reads a run of identifier characters starting at p. */
  const readWord = (p: number): number => {
    let q = p;
    while (q < n && isWordChar(src[q])) q++;
    return q;
  };

  /** The range operator at p (`:`, `:.`, `.:`, `.:.`), as its length, or 0. */
  const rangeOpAt = (p: number): number => {
    if (src.startsWith(".:.", p)) return 3;
    if (src.startsWith(":.", p) || src.startsWith(".:", p)) return 2;
    return src[p] === ":" ? 1 : 0;
  };

  /**
   * Reads what follows a qualifier, or a bare reference/name, starting at p.
   * Returns the token (without pushing) or null if nothing referable is there.
   */
  const readTarget = (start: number, p: number, qual: Qualifier | undefined): Token | null => {
    // Qualified error: `Sheet1!#REF!`
    if (src[p] === "#") {
      for (const e of ERROR_LITERALS) {
        if (src.slice(p, p + e.length).toUpperCase() === e) {
          return mk("ref", start, p + e.length, { qual, value: e, refKind: "error" });
        }
      }
      return null;
    }
    // Row range `1:3`, `$1:$1`, `1:.5`
    if (isDigit(src[p]) || (src[p] === "$" && isDigit(src[p + 1]))) {
      let q = p + (src[p] === "$" ? 1 : 0);
      while (isDigit(src[q])) q++;
      const a = src.slice(p, q);
      const op = rangeOpAt(q);
      if (isRow(a) && op) {
        let r = q + op;
        const r0 = r;
        if (src[r] === "$") r++;
        while (isDigit(src[r])) r++;
        if (isRow(src.slice(r0, r)) && !isWordChar(src[r])) {
          return mk("ref", start, r, { qual, value: src.slice(p, r), refKind: "rows" });
        }
      }
      return null;
    }
    if (!isWordStart(src[p])) return null;
    let q = readWord(p);
    let word = src.slice(p, q);
    // A word directly followed by `(` is a function or a name called, except a cell called
    // like a function (`C2(x, y)`: the cell holds a LAMBDA). Excel resolves a cell-shaped
    // word to the built-in function when the catalogue has it (`LOG10(`, `log10(`); with
    // `$` or a qualifier it can only be the cell (`$C$2(`, `Sheet1!C2(`, `Sheet1!LOG10(`).
    if (src[q] === "(") {
      if (parseCell(word) && (qual !== undefined || word.includes("$") || !lookupFunction(word))) {
        return mk("ref", start, q, { qual, value: word, refKind: "cell" });
      }
      if (word.startsWith("$")) return null;
      return mk("name", start, q, { qual, value: word });
    }
    if (src[q] === "[") {
      if (word.startsWith("$")) return null;
      return mk("name", start, q, { qual, value: word });
    }
    // `A1.:B5`: the dot belongs to the trim operator, not to the word.
    if (word.endsWith(".") && src[q] === ":") {
      const w = word.slice(0, -1);
      if (parseCell(w) || isColumn(w)) {
        q--;
        word = w;
      }
    }
    if (parseCell(word)) {
      // Merge `A1:B2` (and trim forms) into one area token, so a qualifier covers the area.
      const op = rangeOpAt(q);
      if (op) {
        const r0 = q + op;
        let r = readWord(r0);
        let w2 = src.slice(r0, r);
        if (w2.endsWith(".") && src[r] === ":") {
          w2 = w2.slice(0, -1);
          r--;
        }
        if (parseCell(w2) && src[r] !== "(" && src[r] !== "!" && src[r] !== "[") {
          return mk("ref", start, r, { qual, value: src.slice(p, r), refKind: "area" });
        }
      }
      return mk("ref", start, q, { qual, value: word, refKind: "cell" });
    }
    if (isColumn(word)) {
      const op = rangeOpAt(q);
      if (op) {
        const r0 = q + op;
        let r = readWord(r0);
        let w2 = src.slice(r0, r);
        if (w2.endsWith(".") && src[r] === ":") {
          w2 = w2.slice(0, -1);
          r--;
        }
        if (isColumn(w2) && src[r] !== "(" && src[r] !== "!" && src[r] !== "[") {
          return mk("ref", start, r, { qual, value: src.slice(p, r), refKind: "cols" });
        }
      }
    }
    if (word.startsWith("$")) return null;
    const up = word.toUpperCase();
    if (!qual && (up === "TRUE" || up === "FALSE")) {
      return mk("bool", start, q, { value: up });
    }
    return mk("name", start, q, { qual, value: word });
  };

  function mk(kind: TokenKind, start: number, end: number, extra: Partial<Token>): Token {
    return { kind, start, end, text: src.slice(start, end), ...extra };
  }

  /** Reads `[...]` with Excel's bracket nesting and `'` escapes; returns the index after `]`. */
  const readBrackets = (p: number): number => {
    let depth = 0;
    let q = p;
    while (q < n) {
      const c = src[q];
      if (c === "'" && depth > 0) {
        q += 2;
        continue;
      }
      if (c === "[") depth++;
      else if (c === "]") {
        depth--;
        if (depth === 0) return q + 1;
      }
      q++;
    }
    return -1;
  };

  /** After a qualifier ending at p: read the target, or emit an invalid token. */
  const finishQualified = (start: number, p: number, qual: Qualifier): void => {
    const t = readTarget(start, p, qual);
    if (t) {
      out.push(t);
      i = t.end;
    } else {
      push({ kind: "invalid", start, end: p, message: "expected a reference or a name after '!'" });
      i = p;
    }
  };

  /** Parses the inside of a quoted qualifier: `[Book]Sheet`, `Sheet1:Sheet3`, `C:\dir\[Book.xlsx]Sheet`. */
  const splitQuoted = (body: string): Pick<Qualifier, "book" | "sheet" | "sheet2"> => {
    let book: string | undefined;
    let rest = body;
    const close = body.lastIndexOf("]");
    if (close >= 0 && body.includes("[")) {
      book = body.slice(0, close + 1);
      rest = body.slice(close + 1);
    }
    const colon = rest.indexOf(":");
    // A colon inside a quoted name is a 3-D range only when there is no book path before it.
    if (colon > 0 && book === undefined && !rest.includes("\\")) {
      return { sheet: rest.slice(0, colon), sheet2: rest.slice(colon + 1) };
    }
    return book !== undefined ? { book, sheet: rest || undefined } : { sheet: rest };
  };

  while (i < n) {
    const c = src[i]!;
    const start = i;

    if (isWs(c)) {
      while (i < n && isWs(src[i])) i++;
      push({ kind: "ws", start, end: i });
      continue;
    }

    if (c === '"') {
      let q = i + 1;
      let val = "";
      let closed = false;
      while (q < n) {
        if (src[q] === '"') {
          if (src[q + 1] === '"') {
            val += '"';
            q += 2;
            continue;
          }
          closed = true;
          q++;
          break;
        }
        val += src[q];
        q++;
      }
      if (!closed) {
        push({ kind: "invalid", start, end: n, message: "unterminated string" });
        i = n;
        continue;
      }
      push({ kind: "string", start, end: q, value: val });
      i = q;
      continue;
    }

    if (c === "'") {
      let q = i + 1;
      let body = "";
      let closed = false;
      while (q < n) {
        if (src[q] === "'") {
          if (src[q + 1] === "'") {
            body += "'";
            q += 2;
            continue;
          }
          closed = true;
          q++;
          break;
        }
        body += src[q];
        q++;
      }
      if (!closed || src[q] !== "!") {
        push({
          kind: "invalid",
          start,
          end: closed ? q : n,
          message: closed ? "a quoted sheet name must be followed by '!'" : "unterminated quoted sheet name",
        });
        i = closed ? q : n;
        continue;
      }
      q++;
      const qual: Qualifier = { raw: src.slice(start, q), quoted: true, ...splitQuoted(body) };
      finishQualified(start, q, qual);
      continue;
    }

    if (c === "[") {
      const close = readBrackets(i);
      if (close < 0) {
        push({ kind: "invalid", start, end: n, message: "unclosed '['" });
        i = n;
        continue;
      }
      const inner = src.slice(i + 1, close - 1);
      // External workbook: `[1]!Name`, `[1]Sheet1!A1`, `[Book.xlsx]Sheet1!A1`.
      if (src[close] === "!") {
        const qual: Qualifier = { raw: src.slice(start, close + 1), book: src.slice(start, close), quoted: false };
        finishQualified(start, close + 1, qual);
        continue;
      }
      if (isWordStart(src[close]) && !inner.includes("[")) {
        let q = readWord(close);
        let sheet2: string | undefined;
        const sheet = src.slice(close, q);
        if (src[q] === ":" && isWordStart(src[q + 1])) {
          const q2 = readWord(q + 1);
          if (src[q2] === "!") {
            sheet2 = src.slice(q + 1, q2);
            q = q2;
          }
        }
        if (src[q] === "!") {
          const qual: Qualifier = {
            raw: src.slice(start, q + 1),
            book: src.slice(start, close),
            sheet,
            ...(sheet2 !== undefined ? { sheet2 } : {}),
            quoted: false,
          };
          finishQualified(start, q + 1, qual);
          continue;
        }
      }
      // Structured reference without a table name (inside a Table): `[@Col]`, `[[#This Row],[A]]`.
      push({ kind: "structref", start, end: close, value: "", inner });
      i = close;
      continue;
    }

    if (c === "#") {
      let matched = false;
      for (const e of ERROR_LITERALS) {
        if (src.slice(i, i + e.length).toUpperCase() === e) {
          push({ kind: "error", start, end: i + e.length, value: e });
          i += e.length;
          matched = true;
          break;
        }
      }
      if (matched) continue;
      push({ kind: "op", start, end: i + 1 });
      i++;
      continue;
    }

    // Numbers, and whole-row references `1:1`.
    if (isDigit(c) || (c === "." && isDigit(src[i + 1])) || (c === "$" && isDigit(src[i + 1]))) {
      const rows = readTarget(start, i, undefined);
      if (rows && rows.refKind === "rows") {
        out.push(rows);
        i = rows.end;
        continue;
      }
      if (c === "$") {
        push({ kind: "invalid", start, end: i + 1, message: "unexpected '$'" });
        i++;
        continue;
      }
      let q = i;
      while (isDigit(src[q])) q++;
      if (src[q] === "." && !src.startsWith(".:", q)) {
        q++;
        while (isDigit(src[q])) q++;
      }
      if ((src[q] === "E" || src[q] === "e") && (isDigit(src[q + 1]) || ((src[q + 1] === "+" || src[q + 1] === "-") && isDigit(src[q + 2])))) {
        q += 2;
        while (isDigit(src[q])) q++;
      }
      if (isWordChar(src[q]) && src[q] !== ".") {
        // `2A`, `1E`: not a number and not a name.
        const q2 = readWord(q);
        push({ kind: "invalid", start, end: q2, message: `'${src.slice(start, q2)}' is not a number, a reference or a name` });
        i = q2;
        continue;
      }
      push({ kind: "number", start, end: q });
      i = q;
      continue;
    }

    if (isWordStart(c)) {
      const q = readWord(i);
      // Sheet qualifier: `Sheet1!`, or 3-D `Sheet1:Sheet3!`.
      if (src[q] === "!" && !src.slice(i, q).startsWith("$")) {
        const qual: Qualifier = { raw: src.slice(i, q + 1), sheet: src.slice(i, q), quoted: false };
        finishQualified(start, q + 1, qual);
        continue;
      }
      if (src[q] === ":" && isWordStart(src[q + 1])) {
        const q2 = readWord(q + 1);
        if (src[q2] === "!") {
          const qual: Qualifier = {
            raw: src.slice(i, q2 + 1),
            sheet: src.slice(i, q),
            sheet2: src.slice(q + 1, q2),
            quoted: false,
          };
          finishQualified(start, q2 + 1, qual);
          continue;
        }
      }
      const t = readTarget(start, i, undefined);
      if (!t) {
        push({ kind: "invalid", start, end: q, message: `'${src.slice(start, q)}' is not a reference or a name` });
        i = q;
        continue;
      }
      if (t.kind === "name" && src[t.end] === "[") {
        const close = readBrackets(t.end);
        if (close < 0) {
          push({ kind: "invalid", start, end: n, message: "unclosed '[' in structured reference" });
          i = n;
          continue;
        }
        push({ kind: "structref", start, end: close, value: t.value, inner: src.slice(t.end + 1, close - 1) });
        i = close;
        continue;
      }
      out.push(t);
      i = t.end;
      continue;
    }

    const three = src.slice(i, i + 3);
    if (OPS3.includes(three)) {
      push({ kind: "op", start, end: i + 3 });
      i += 3;
      continue;
    }
    const two = src.slice(i, i + 2);
    if (OPS2.includes(two)) {
      push({ kind: "op", start, end: i + 2 });
      i += 2;
      continue;
    }
    if (OPS1.includes(c)) {
      push({ kind: "op", start, end: i + 1 });
      i++;
      continue;
    }
    if (c === "(" || c === ")" || c === "{" || c === "}" || c === "," || c === ";") {
      push({ kind: c, start, end: i + 1 });
      i++;
      continue;
    }
    push({ kind: "invalid", start, end: i + 1, message: `unexpected character '${c}'` });
    i++;
  }

  markIntersections(out);
  out.push({ kind: "eof", start: n, end: n, text: "" });
  return out;
}

const ENDS_OPERAND = new Set<TokenKind>(["ref", "name", "structref", ")"]);
const STARTS_OPERAND = new Set<TokenKind>(["ref", "name", "structref", "("]);

/** Whitespace between the end of one operand and the start of another is Excel's
 *  intersection operator (`A1:B5 B2:C3`); everywhere else it is layout. */
function markIntersections(toks: Token[]): void {
  for (let k = 1; k + 1 < toks.length; k++) {
    const t = toks[k]!;
    if (t.kind !== "ws") continue;
    const prev = toks[k - 1]!;
    const next = toks[k + 1]!;
    const prevEnds = ENDS_OPERAND.has(prev.kind) || (prev.kind === "op" && prev.text === "#");
    if (prevEnds && STARTS_OPERAND.has(next.kind)) t.kind = "isect";
  }
}

/** The significant tokens: no layout whitespace, no eof. */
export function significant(toks: Token[]): Token[] {
  return toks.filter((t) => t.kind !== "ws" && t.kind !== "eof");
}
