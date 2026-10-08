// What surrounds a cursor in a formula being typed (M3c: completion and signature help).
// The text before the cursor is usually not a whole formula (`=IF(Sales > `), so this
// reads tokens, not a syntax tree: the identifier being typed and its qualifier, whether
// the cursor sits inside a string or a structured reference, the calls left open, and the
// LET and LAMBDA variables in scope.

import { stripPrefix } from "./parser.js";
import { significant, tokenize, type Span, type Token } from "./tokens.js";

/** A call (or bracket) open at the cursor. */
export interface CursorFrame {
  /** The function as written (`SUM`, `FN.PICK`, `_xlfn.LET`); undefined for `(…)`, `{…}` and `f(…)(…)`. */
  fn: string | undefined;
  /** The sheet written before the function (`'S'!Fn(`). */
  sheet: string | undefined;
  /** Where the function name is written. */
  fnSpan: Span | undefined;
  /** `(` of a call or group, `{` of an array constant. */
  bracket: "(" | "{";
  /** The arguments before the one at the cursor, as spans of the formula text. */
  args: Span[];
  /** Where the argument at the cursor starts. */
  argStart: number;
}

/** A LET or LAMBDA variable in scope at the cursor. */
export interface CursorLocal {
  /** As written, without `_xlpm.`. */
  name: string;
  kind: "let" | "lambda";
  /** Where it is declared. */
  span: Span;
  /** For a LET variable: its value (a LAMBDA gives signature help). */
  value?: Span;
}

export interface FormulaCursor {
  /**
   * Nothing to complete: inside a string, a structured reference's brackets or a quoted
   * sheet name, or in the middle of a number or an absolute address (`$A`).
   */
  inert: boolean;
  /** The identifier ending at the cursor ("" when none). */
  word: Span & { text: string };
  /** The sheet written before it (`IS!Sa`, `'SCF recursive'!`), un-quoted. */
  sheet: string | undefined;
  /** Where the qualifier starts, when there is one. */
  qualifierStart: number | undefined;
  /** Calls and brackets open at the cursor, outermost first. */
  frames: CursorFrame[];
  /** LET and LAMBDA variables in scope, innermost first. */
  locals: CursorLocal[];
}

const WORD = /[\p{L}\p{N}\p{M}_.\\?]/u;
const WORD_START = /[\p{L}_\\]/u;

function isWordChar(c: string | undefined): boolean {
  return c !== undefined && WORD.test(c);
}

/** The variable a single-token argument declares (`x`, `_xlpm.x`, LAMBDA's `[y]`), or undefined. */
function declared(src: string, arg: Span): { name: string; span: Span } | undefined {
  const toks = significant(tokenize(src.slice(arg.start, arg.end))).filter((t) => t.kind !== "eof");
  if (toks.length !== 1) return undefined;
  const t = toks[0]!;
  if (t.kind === "name" && !t.qual) {
    return { name: stripPrefix(t.value ?? t.text).base, span: { start: arg.start + t.start, end: arg.start + t.end } };
  }
  // `[y]`: an optional LAMBDA parameter reads as a structured reference without a table.
  if (t.kind === "structref" && t.value === "" && t.inner !== undefined) {
    const inner = t.inner.trim();
    if (inner !== "" && [...inner].every((c) => isWordChar(c))) {
      const at = arg.start + t.start + t.text.indexOf(inner);
      return { name: stripPrefix(inner).base, span: { start: at, end: at + inner.length } };
    }
  }
  return undefined;
}

function upperBase(fn: string | undefined): string {
  return fn === undefined ? "" : stripPrefix(fn).base.toUpperCase();
}

/** The cursor context at `offset` of a formula (display form, without its `=`). */
export function formulaCursor(formula: string, offset: number): FormulaCursor {
  const at = Math.max(0, Math.min(offset, formula.length));
  let ws = at;
  while (ws > 0 && isWordChar(formula[ws - 1])) ws--;
  const text = formula.slice(ws, at);
  let inert = false;
  // `1.5E`, `A1.`: a number or an address, not a name; `$A`: an absolute address.
  if (text !== "" && !WORD_START.test(text[0]!)) inert = true;
  if (formula[ws - 1] === "$") inert = true;

  let sheet: string | undefined;
  let qualifierStart: number | undefined;
  if (formula[ws - 1] === "!") {
    let q = ws - 1;
    if (formula[q - 1] === "'") {
      // `'It''s'!`: back to the opening quote, doubled quotes inside.
      let k = q - 2;
      let body = "";
      while (k >= 0) {
        if (formula[k] === "'") {
          if (formula[k - 1] === "'") {
            body = "'" + body;
            k -= 2;
            continue;
          }
          break;
        }
        body = formula[k] + body;
        k--;
      }
      if (k >= 0) {
        sheet = body;
        qualifierStart = k;
      }
    } else {
      let k = q;
      while (k > 0 && isWordChar(formula[k - 1])) k--;
      if (k < q) {
        sheet = formula.slice(k, q);
        qualifierStart = k;
      }
    }
  }

  const before = formula.slice(0, qualifierStart ?? ws);
  const all = tokenize(before);
  const last = all.filter((t) => t.kind !== "eof").at(-1);
  // An unterminated string, structured reference or quoted sheet runs to the cursor.
  if (last && last.kind === "invalid" && last.end === before.length && sheet === undefined) inert = true;

  const toks = significant(all).filter((t) => t.kind !== "eof");
  const frames: CursorFrame[] = [];
  for (let k = 0; k < toks.length; k++) {
    const t: Token = toks[k]!;
    if (t.kind === "(" || t.kind === "{") {
      const prev = toks[k - 1];
      const isCall = t.kind === "(" && prev !== undefined && prev.kind === "name" && prev.end === t.start;
      frames.push({
        fn: isCall ? (prev.value ?? prev.text) : undefined,
        sheet: isCall ? prev.qual?.sheet : undefined,
        fnSpan: isCall ? { start: prev.end - (prev.value ?? prev.text).length, end: prev.end } : undefined,
        bracket: t.kind,
        args: [],
        argStart: t.end,
      });
    } else if (t.kind === ")" || t.kind === "}") {
      frames.pop();
    } else if (t.kind === ",") {
      const f = frames.at(-1);
      if (f) {
        f.args.push({ start: f.argStart, end: t.start });
        f.argStart = t.end;
      }
    }
  }

  const locals: CursorLocal[] = [];
  for (let k = frames.length - 1; k >= 0; k--) {
    const f = frames[k]!;
    if (f.bracket !== "(" || f.sheet !== undefined) continue;
    const fn = upperBase(f.fn);
    const found: CursorLocal[] = [];
    if (fn === "LET") {
      // A binding's name is in scope once its value is complete: `LET(a, 1, b, a + |`.
      for (let j = 0; j + 1 < f.args.length; j += 2) {
        const d = declared(before, f.args[j]!);
        if (d) found.push({ name: d.name, kind: "let", span: d.span, value: f.args[j + 1]! });
      }
    } else if (fn === "LAMBDA") {
      for (const a of f.args) {
        const d = declared(before, a);
        if (d) found.push({ name: d.name, kind: "lambda", span: d.span });
      }
    }
    locals.push(...found.reverse());
  }
  return { inert, word: { start: ws, end: at, text }, sheet, qualifierStart, frames, locals };
}

/** The innermost call open at the cursor (a group or an array constant inside it is looked through), with the argument the cursor is in. */
export function activeCall(c: FormulaCursor): { frame: CursorFrame; index: number } | undefined {
  for (let k = c.frames.length - 1; k >= 0; k--) {
    const f = c.frames[k]!;
    if (f.fn !== undefined) return { frame: f, index: f.args.length };
  }
  return undefined;
}
