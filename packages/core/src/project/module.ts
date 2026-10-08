// `.xln` module files: AFE module syntax plus annotations (FEASIBILITY §3.1).
//
//   // line comment                 /* block comment */
//   /** doc comment: the Name Manager comment of the next name */
//   @hidden                         annotations apply to the next definition only
//   @from(lib #353921)              the library version this copy came from (written by
//                                   Insert, Take and Publish; the build carries it in the
//                                   provenance tag, pull writes it back)
//   name = formula;                 the formula ends at the first `;` outside strings,
//                                   quoted sheet names, brackets and array braces
//
// `name : type = …` (an AFE-style type declaration) is an error: types are not part of
// xln v1; they come with the dimension layer (author's decision, 2026-10-07).
//
// Scope is per name in every file (M3d for sheet files, 2026-10-06; module files
// 2026-10-07):
//
//   names/sheets/<Sheet>.xln        every name is local to the file's sheet, unless the
//                                   line above it says `@workbook` (an annotation like
//                                   `@hidden`: the next statement only; both may stack)
//   names/<Module>.xln              workbook scope, unless the line above the name says
//                                   `@sheet(Sheet)`: local to Sheet (the next statement only)
//
// Files written before then have blocks: `@scope(Sheet)` opens a block of names local to
// Sheet, until the next `@scope(…)` or `@workbook` (back to workbook scope). They are still
// read: a sheet file with a `@scope(…)` line in the block form (`form: "sheet-blocks"`,
// `convertSheetBlocks`), a module file with blocks as before (`convertModuleBlocks`
// rewrites it per name). In a sheet file `@sheet(Name)` on a line of its own names the
// file's sheet (a pull's `IS~2.xln`). The sheet in `@scope(…)` and `@sheet(…)` is quoted
// ('…', `''` for a quote) only when this syntax needs it: blank, spaces, parentheses or
// quotes. Excel's own quoting rule does not apply here.
// Comments inside a formula are layout: the parser drops them. In a doc comment `*/` is
// written `*\/`.
//
// Cell statements (M3b, D7/D8): the formula of an existing cell, written where the sheet
// has it.
//
//   Name @C6 = formula;             a named cell: the name is on C6 (`'S'!$C$6`)
//   Name @C6# = formula;            the name is on C6's spill (`'S'!$C$6#`, stored
//                                   `_xlfn.ANCHORARRAY('S'!$C$6)`)
//   Name @C6 = ;                    a slot: a name on an empty cell, waiting for a formula
//   @C5 = formula;                  an unnamed formula cell
//   @B40:G40 = formula;             a block holding one formula filled across; the
//                                   formula is the top-left cell's
//   Name @'S 1'!C6 = …;             the address names its sheet: in a module file, and in
//                                   an old sheet file's @workbook block
//
// The address is set in Excel and read-only, but for its `#` (author's decision,
// 2026-10-05): a named statement's address says what the name covers, and adding or
// removing the `#` changes the name. Unnamed statements and blocks take no `#`. In a sheet
// file the sheet is the file's and the address is bare; elsewhere it is the `@scope(…)`
// block's unless written. A written sheet is quoted ('…') only when it holds blanks,
// quotes, `!`, `=`, `;`, `@`, parentheses or `/`. An empty right-hand side is allowed
// only in a cell statement.

import { sheetNeedsQuotes } from "../lang/tokens.js";
import { sheetFromFileName } from "./modules.js";
import { FROM, formatLibBase, parseLibBase } from "./provenance.js";

const SCOPE = "scope";
const WORKBOOK = "workbook";
const SHEET = "sheet";

/** Where a cell statement's address is written. */
export interface CellTarget {
  /** The sheet written in the address (`@BS!C6`), un-quoted; undefined when not written. */
  sheet?: string;
  /** The address without `$`, upper case: `C6` or `B40:G40`. */
  range: string;
  /** Offset of the `@`. */
  start: number;
  /** Offset of the range (after `Sheet!`, if any). */
  rangeStart: number;
  /** `@C6#`: the name covers the cell's spill. The `#` is the last character of the address. */
  spill: boolean;
  /** Offset just after the address (after its `#`, if any). */
  end: number;
}

export interface Annotation {
  name: string;
  /** Text between the parentheses (a quoted sheet name is un-quoted); undefined without them. */
  arg?: string;
  line: number;
  /** Offsets in the text: the `@`, and just after the annotation (after its `)`, if any). */
  offset: number;
  end: number;
}

export interface ModuleEntry {
  /** "" for an unnamed cell statement (`@C5 = …;`), which appears only in `ParsedModule.cells`. */
  name: string;
  /** For a cell statement: its address. */
  cell?: CellTarget;
  /** The sheet the name is local to (a sheet file's, its `@sheet(…)`, or its old `@scope(…)` block's); undefined for workbook scope. */
  scope?: string;
  /** `@workbook` above it in a sheet file: a workbook-scoped name. */
  workbook?: boolean;
  /** For a cell statement: the sheet of its cell (the address's, else the sheet file's or the block's). */
  cellSheet?: string;
  /** From `@hidden`. */
  hidden: boolean;
  /** From `@from(lib #353921)`: the library base, 6 hex digits (absent when malformed: the checker says why). */
  from?: string;
  doc?: string;
  /**
   * Doc comments before `doc` for the same statement (offsets of `/**` and just after `*\/`):
   * only the last one is the Name Manager comment, so these are lost.
   */
  extraDocs?: { start: number; end: number }[];
  /** Annotations written directly before the name (a sheet file's `@workbook` among them); `@scope`/`@workbook` blocks are not. */
  annotations: Annotation[];
  /** The formula in display form, comments removed, trimmed. */
  formula: string;
  /** 1-based line of the name. */
  line: number;
  /** Offset of the name in the module text. */
  offset: number;
  /** Offset where the entry begins: its doc comment or first annotation, else the name. */
  start: number;
  /** Offset just after the closing `;` (or the end of the text when it is missing). */
  end: number;
  /**
   * Where `formula` came from: pairs `[i, src]` (flattened), each meaning that `formula[i]`
   * onwards sits at `src` in the module text, up to the next pair. Comments removed from
   * the formula break the text into such runs. Read it with `formulaToSource`.
   */
  formulaMap: number[];
}

/** An `@scope(…)` or `@workbook` line: the start of a block. */
export interface ScopeDirective {
  /** The sheet; undefined for `@workbook`. */
  scope: string | undefined;
  line: number;
  offset: number;
  /** Offset just after the directive. */
  end: number;
}

export interface ModuleDiagnostic {
  line: number;
  col: number;
  message: string;
}

export interface ParsedModule {
  /** Named definitions, named cell statements included. */
  entries: ModuleEntry[];
  /** Every cell statement, named and unnamed, in text order (the named ones are also in `entries`). */
  cells: ModuleEntry[];
  diagnostics: ModuleDiagnostic[];
  /** The `@scope(…)`/`@workbook` block directives in text order (none in a sheet file of the new form). */
  scopes: ScopeDirective[];
  /**
   * `module`: a module file (or a text parsed without a path); `sheet`: a sheet file, its
   * names local to its sheet unless annotated `@workbook`; `sheet-blocks`: a sheet file
   * written before M3d, with `@scope`/`@workbook` blocks.
   */
  form: "module" | "sheet" | "sheet-blocks";
  /** For a sheet file: its sheet (from the file name). */
  sheet?: string;
  /** A sheet file's first `@sheet(…)` line (in a module file `@sheet(…)` annotates a name). */
  sheetDirective?: { sheet: string; offset: number; end: number };
}

export interface ParseOptions {
  /** The sheet of a sheet file (`names/sheets/<Sheet>.xln`): its names are local to it unless annotated `@workbook`. */
  sheet?: string;
  /** A file in `names/sheets/` whose name gives no sheet: a sheet file only with `@sheet(Name)`. */
  sheetFile?: boolean;
}

/** The sheet of a sheet file's project path (`names/sheets/IS.xln` → `IS`); undefined for any other path. */
export function sheetOfPath(path: string): string | undefined {
  const dir = "names/sheets/";
  if (!path.startsWith(dir)) return undefined;
  const file = path.slice(dir.length);
  if (file.includes("/") || !file.endsWith(".xln")) return undefined;
  return sheetFromFileName(file);
}

/**
 * Parses a project file: a sheet file (known by its path; `@sheet(Name)` in it names its
 * sheet when the file name cannot, as when a pull added `~2`) or a module file.
 */
export function parseSourceFile(path: string, text: string): ParsedModule {
  const dir = "names/sheets/";
  if (!path.startsWith(dir) || path.slice(dir.length).includes("/")) return parseModule(text);
  const sheet = sheetOfPath(path);
  return parseModule(text, sheet === undefined ? { sheetFile: true } : { sheet, sheetFile: true });
}

/** The module-text offset of `formula[index]` (`index` may equal the formula's length). */
export function formulaToSource(entry: Pick<ModuleEntry, "formulaMap">, index: number): number {
  const m = entry.formulaMap;
  let k = 0;
  while (k + 2 < m.length && m[k + 2]! <= index) k += 2;
  return m[k + 1]! + (index - m[k]!);
}

/**
 * The formula index at module-text offset `offset`, or undefined when the offset is not
 * inside the formula text (before it, after it, or inside a comment).
 */
export function sourceToFormula(entry: Pick<ModuleEntry, "formulaMap" | "formula">, offset: number): number | undefined {
  const m = entry.formulaMap;
  for (let k = 0; k < m.length; k += 2) {
    const at = m[k]!;
    const next = k + 2 < m.length ? m[k + 2]! : entry.formula.length;
    const src = m[k + 1]!;
    if (offset >= src && offset - src <= next - at) return at + (offset - src);
  }
  return undefined;
}

function isWs(c: string | undefined): boolean {
  return c === " " || c === "\t" || c === "\r" || c === "\n";
}

function isNameChar(c: string | undefined): boolean {
  return c !== undefined && !isWs(c) && c !== ":" && c !== "=" && c !== ";" && c !== "(" && c !== ")" && c !== "/" && c !== "@";
}

function isAsciiLetter(c: string | undefined): boolean {
  return c !== undefined && ((c >= "A" && c <= "Z") || (c >= "a" && c <= "z"));
}

function isDigitChar(c: string | undefined): boolean {
  return c !== undefined && c >= "0" && c <= "9";
}

/** Characters that end an unquoted sheet in a cell address, or force quotes when writing one. */
const ADDRESS_SHEET_STOP = new Set([" ", "\t", "\r", "\n", "!", "'", "=", ";", "@", "(", ")", "/", ":"]);

/** `C6` or `$C$6` at `p`: its end and its text without `$`, upper case. */
function scanA1(text: string, p: number): { end: number; cell: string } | undefined {
  let k = p;
  if (text[k] === "$") k++;
  const c0 = k;
  while (isAsciiLetter(text[k])) k++;
  if (k - c0 < 1 || k - c0 > 3) return undefined;
  const col = text.slice(c0, k).toUpperCase();
  if (text[k] === "$") k++;
  const r0 = k;
  while (isDigitChar(text[k])) k++;
  if (k === r0 || k - r0 > 7 || text[r0] === "0") return undefined;
  if (text[k] !== undefined && /[\p{L}\p{N}_.?\\$]/u.test(text[k]!)) return undefined;
  return { end: k, cell: col + text.slice(r0, k) };
}

/**
 * The address of a cell statement at `at` (an `@`): `@C6`, `@B40:G40`, `@BS!C6`,
 * `@'SCF recursive'!C6`. Undefined when the text there is not one (an annotation).
 */
export function scanCellTarget(text: string, at: number): CellTarget | undefined {
  if (text[at] !== "@") return undefined;
  let k = at + 1;
  let sheet: string | undefined;
  if (text[k] === "'") {
    let s = "";
    let j = k + 1;
    for (; j < text.length; j++) {
      if (text[j] === "'") {
        if (text[j + 1] === "'") {
          s += "'";
          j++;
        } else break;
      } else if (text[j] === "\n") return undefined;
      else s += text[j];
    }
    if (text[j] !== "'" || text[j + 1] !== "!" || s === "") return undefined;
    sheet = s;
    k = j + 2;
  } else {
    let e = k;
    while (e < text.length && !ADDRESS_SHEET_STOP.has(text[e]!)) e++;
    if (text[e] === "!" && e > k) {
      sheet = text.slice(k, e);
      k = e + 1;
    }
  }
  const rangeStart = k;
  const a = scanA1(text, k);
  if (!a) return undefined;
  let range = a.cell;
  k = a.end;
  if (text[k] === ":") {
    const b = scanA1(text, k + 1);
    if (!b) return undefined;
    range += ":" + b.cell;
    k = b.end;
  }
  const spill = text[k] === "#";
  if (spill) k++;
  const t: CellTarget = { range, start: at, rangeStart, spill, end: k };
  if (sheet !== undefined) t.sheet = sheet;
  return t;
}

/** The address of a cell statement as written after `@`: `C6`, or `Sheet!C6` with the sheet quoted when needed; `C6#` for a name on the spill. */
export function formatCellAddress(range: string, sheet?: string, spill = false): string {
  const r = spill ? range + "#" : range;
  if (sheet === undefined) return r;
  return `${sheetNeedsQuotes(sheet) ? "'" + sheet.split("'").join("''") + "'" : sheet}!${r}`;
}

function unescapeDoc(s: string): string {
  return s.split("*\\/").join("*/");
}

/** The text of a doc comment given what lies between `/**` and `*\/`. */
export function docText(inner: string): string {
  if (!inner.includes("\n")) {
    let s = inner.replace(/\r$/, "");
    if (s.startsWith(" ")) s = s.slice(1);
    if (s.endsWith(" ")) s = s.slice(0, -1);
    return unescapeDoc(s);
  }
  const lines = inner.split("\n").map((l) => (l.endsWith("\r") ? l.slice(0, -1) : l));
  if (lines.length > 0 && lines[0]!.trim() === "") lines.shift();
  if (lines.length > 0 && lines[lines.length - 1]!.trim() === "") lines.pop();
  return unescapeDoc(
    lines
      .map((l) => {
        let k = 0;
        while (l[k] === " " || l[k] === "\t") k++;
        if (l[k] !== "*") return l.slice(k);
        k++;
        if (l[k] === " ") k++;
        return l.slice(k);
      })
      .join("\n"),
  );
}

/** Writes a doc comment for `text` (no trailing newline). */
export function formatDoc(text: string, indent = ""): string {
  const t = text.split("*/").join("*\\/");
  if (!t.includes("\n")) return `${indent}/** ${t} */`;
  return [`${indent}/**`, ...t.split("\n").map((l) => `${indent} *${l === "" ? "" : " " + l}`), `${indent} */`].join("\n");
}

/**
 * Parses an `.xln` file. Never throws; problems come back as diagnostics. With `sheet`
 * (a sheet file, `parseSourceFile`) its names are local to that sheet unless annotated
 * `@workbook`; a sheet file with `@scope(…)` lines is read in the old block form.
 */
export function parseModule(text: string, opts: ParseOptions = {}): ParsedModule {
  if (opts.sheet === undefined && !opts.sheetFile) {
    const { fileSheet: _, ...module } = parseText(text, undefined, true);
    return { ...module, form: "module" };
  }
  const { fileSheet, ...blocks } = parseText(text, undefined, false);
  const sheet = fileSheet ?? opts.sheet;
  if (sheet === undefined) {
    // A file in names/sheets/ whose name gives no sheet: read as a module's.
    const { fileSheet: _, ...module } = parseText(text, undefined, true);
    return { ...module, form: "module" };
  }
  if (blocks.scopes.some((d) => d.scope !== undefined)) return { ...blocks, form: "sheet-blocks", sheet };
  const { fileSheet: _, ...annotated } = parseText(text, sheet, false);
  return { ...annotated, form: "sheet", sheet };
}

/**
 * `sheet`: a sheet file of the new form, where `@workbook` is an annotation. `module`: a
 * module file, where `@sheet(…)` is one. Otherwise (a sheet file's first reading) `@sheet`
 * names the file's sheet, and both blocks are read.
 */
function parseText(text: string, sheet: string | undefined, module: boolean): Omit<ParsedModule, "form" | "sheet"> & { fileSheet?: string } {
  let sheetDirective: ParsedModule["sheetDirective"];
  const annotate = sheet !== undefined;
  const entries: ModuleEntry[] = [];
  const cells: ModuleEntry[] = [];
  const diagnostics: ModuleDiagnostic[] = [];
  const scopes: ScopeDirective[] = [];
  let fileSheet: string | undefined;
  const n = text.length;
  let i = 0;
  let line = 1;
  let lineStart = 0;
  let doc: string | undefined;
  let docSpan: { start: number; end: number } | undefined;
  let extraDocs: { start: number; end: number }[] = [];
  let annotations: Annotation[] = [];
  let scope: string | undefined;
  /** Where the doc comment or annotations of the next entry begin. */
  let pendingStart: number | undefined;

  const advance = (to: number): void => {
    for (; i < to; i++) {
      if (text[i] === "\n") {
        line++;
        lineStart = i + 1;
      }
    }
  };
  const fail = (message: string): void => {
    diagnostics.push({ line, col: i - lineStart + 1, message });
  };
  const skipLine = (): void => {
    const e = text.indexOf("\n", i);
    advance(e < 0 ? n : e);
  };
  /** Skips a `/* … *\/` comment starting at i; returns its inner text, or undefined if unclosed. */
  const blockComment = (open: number): string | undefined => {
    const e = text.indexOf("*/", i + open);
    if (e < 0) {
      fail("comment is not closed");
      advance(n);
      return undefined;
    }
    const inner = text.slice(i + open, e);
    advance(e + 2);
    return inner;
  };

  while (i < n) {
    const c = text[i];
    if (isWs(c)) {
      advance(i + 1);
      continue;
    }
    if (c === "/" && text[i + 1] === "/") {
      skipLine();
      continue;
    }
    if (c === "/" && text[i + 1] === "*") {
      const isDoc = text[i + 2] === "*" && text[i + 3] !== "/";
      const at = i;
      const inner = blockComment(isDoc ? 3 : 2);
      if (isDoc && inner !== undefined) {
        if (docSpan !== undefined) extraDocs.push(docSpan);
        doc = docText(inner);
        docSpan = { start: at, end: i };
        pendingStart ??= at;
      }
      continue;
    }
    if (c === "@") {
      const target = scanCellTarget(text, i);
      if (target) {
        let k = target.end;
        while (isWs(text[k])) k++;
        if (text[k] === "=") {
          // `@C5 = formula;`: an unnamed cell. It has no name to carry a comment or annotations.
          if (annotations.some((a) => a.name === WORKBOOK)) fail(`@workbook scopes a name, and @${target.range} has none: remove the @workbook line above it`);
          else if (doc !== undefined || annotations.length > 0) fail("an unnamed cell statement takes no doc comment or annotation");
          if (target.spill) {
            advance(target.end - 1);
            fail(`@${target.range}#: only a named cell statement takes '#' (it says what the name covers)`);
          }
          const stmtLine = line;
          advance(k);
          statement("", stmtLine, target.start, target);
          continue;
        }
      }
      const startLine = line;
      const at = i;
      let k = i + 1;
      while (k < n && /[\p{L}\p{N}_]/u.test(text[k]!)) k++;
      const name = text.slice(i + 1, k);
      advance(k);
      if (name === "") {
        fail("'@' must be followed by an annotation name");
        skipLine();
        continue;
      }
      let arg: string | undefined;
      let k2 = i;
      while (text[k2] === " " || text[k2] === "\t") k2++;
      if (text[k2] === "(") {
        advance(k2 + 1);
        while (text[i] === " " || text[i] === "\t") advance(i + 1);
        if (text[i] === "'") {
          let s = "";
          let j = i + 1;
          for (; j < n; j++) {
            if (text[j] === "'") {
              if (text[j + 1] === "'") {
                s += "'";
                j++;
              } else break;
            } else s += text[j];
          }
          const quoted = i;
          advance(Math.min(j + 1, n));
          arg = s;
          // `@renamed('Cash Flow'!Old)`: a quoted sheet, then the name. The argument stays
          // as written (quotes and `''` included); `renamedFrom` reads the sheet out of it.
          if (text[i] === "!") {
            let e = i + 1;
            while (e < n && text[e] !== ")" && text[e] !== "\n") e++;
            if (text[e] !== ")") {
              fail(`'@${name}(' is not closed`);
              skipLine();
              continue;
            }
            arg = text.slice(quoted, e).trim();
            advance(e);
          }
          while (text[i] === " " || text[i] === "\t") advance(i + 1);
          if (text[i] !== ")") {
            fail(`'@${name}(' is not closed`);
            skipLine();
            continue;
          }
          advance(i + 1);
        } else {
          const e = text.indexOf(")", i);
          const nl = text.indexOf("\n", i);
          if (e < 0 || (nl >= 0 && nl < e)) {
            fail(`'@${name}(' is not closed`);
            skipLine();
            continue;
          }
          arg = text.slice(i, e).trim();
          advance(e + 1);
        }
      }
      if (name === SHEET && module) {
        // A module file: the next statement is local to this sheet.
        if (arg === undefined || arg === "") diagnostics.push({ line: startLine, col: 1, message: "@sheet needs the sheet's name: @sheet(Name) on the line above a name local to that sheet" });
        else {
          annotations.push({ name, arg, line: startLine, offset: at, end: i });
          pendingStart ??= at;
        }
      } else if (name === SHEET) {
        // `@sheet(Name)`: a sheet file whose name cannot say its sheet (a pull's `~2`).
        if (arg === undefined || arg === "") diagnostics.push({ line: startLine, col: 1, message: "@sheet needs the sheet's name: @sheet(Name)" });
        else {
          fileSheet ??= arg;
          sheetDirective ??= { sheet: arg, offset: at, end: i };
        }
      } else if (annotate && name === WORKBOOK) {
        // A sheet file: the next statement is workbook-scoped.
        if (arg !== undefined) diagnostics.push({ line: startLine, col: 1, message: "@workbook takes no argument: write @workbook on the line above the name" });
        annotations.push({ name, line: startLine, offset: at, end: i });
        pendingStart ??= at;
      } else if (name === SCOPE) {
        if (arg === undefined || arg === "") diagnostics.push({ line: startLine, col: 1, message: "@scope needs a sheet name" });
        else {
          scope = arg;
          scopes.push({ scope, line: startLine, offset: at, end: i });
        }
      } else if (name === WORKBOOK) {
        if (arg !== undefined) diagnostics.push({ line: startLine, col: 1, message: "@workbook takes no argument" });
        scope = undefined;
        scopes.push({ scope: undefined, line: startLine, offset: at, end: i });
      } else {
        annotations.push(arg === undefined ? { name, line: startLine, offset: at, end: i } : { name, arg, line: startLine, offset: at, end: i });
        pendingStart ??= at;
      }
      continue;
    }
    if (!isNameChar(c)) {
      fail(`unexpected '${c}': expected a name, a comment or an annotation`);
      skipLine();
      continue;
    }
    // name = formula ;
    const nameLine = line;
    const nameOffset = i;
    let k = i;
    while (k < n && isNameChar(text[k])) k++;
    const name = text.slice(i, k);
    advance(k);
    while (isWs(text[i])) advance(i + 1);
    let cell: CellTarget | undefined;
    if (text[i] === "@") {
      cell = scanCellTarget(text, i);
      if (!cell) {
        fail(`'${name}': expected a cell address after '@' (C6, B4:G4, Sheet!C6)`);
        skipLine();
        doc = undefined;
        docSpan = undefined;
        extraDocs = [];
        annotations = [];
        pendingStart = undefined;
        continue;
      }
      if (cell.spill && cell.range.includes(":")) {
        advance(cell.end - 1);
        fail(`'${name}': '#' follows one cell (the spill's anchor), not a range`);
      }
      advance(cell.end);
      while (isWs(text[i])) advance(i + 1);
    } else if (text[i] === ":") {
      // `Name : type = …`: AFE's type declaration. Not xln v1 (decided 2026-10-07); an error
      // on the type, and the statement is read on so the rest of the file still checks.
      const e = text.indexOf("=", i);
      let t = i + 1;
      while (t < n && (text[t] === " " || text[t] === "\t")) t++;
      advance(t);
      fail("type declarations are not part of xln v1 (planned with the dimension layer)");
      if (e < 0) {
        advance(n);
        break;
      }
      advance(e);
    }
    if (text[i] !== "=") {
      fail(`'${name}': expected '='`);
      skipLine();
      doc = undefined;
      docSpan = undefined;
      extraDocs = [];
      annotations = [];
      pendingStart = undefined;
      continue;
    }
    statement(name, nameLine, nameOffset, cell);
  }
  if (doc !== undefined || annotations.length > 0) {
    diagnostics.push({ line, col: i - lineStart + 1, message: "doc comment or annotation without a definition after it" });
  }
  return { entries, cells, diagnostics, scopes, ...(fileSheet !== undefined ? { fileSheet } : {}), ...(sheetDirective !== undefined ? { sheetDirective } : {}) };

  /** Reads `= formula;` (i is at the `=`) and records the entry. */
  function statement(name: string, nameLine: number, nameOffset: number, cell: CellTarget | undefined): void {
    advance(i + 1);
    const { formula, closed, map } = scanFormula();
    const label = name === "" && cell ? `@${cell.range}` : name;
    if (!closed) fail(`'${label}': definition is not ended by ';'`);
    const lead = formula.length - formula.trimStart().length;
    const trimmed = formula.trim();
    const entry: ModuleEntry = {
      name,
      hidden: false,
      annotations,
      formula: trimmed,
      line: nameLine,
      offset: nameOffset,
      start: pendingStart ?? nameOffset,
      end: i,
      formulaMap: shiftMap(map, lead),
    };
    if (doc !== undefined) entry.doc = doc;
    if (extraDocs.length > 0) entry.extraDocs = extraDocs;
    for (const a of annotations) {
      if (a.name === "hidden") entry.hidden = true;
      else if (a.name === FROM && entry.from === undefined) {
        const b = parseLibBase(a.arg);
        if ("hash" in b) entry.from = b.hash;
      }
    }
    const local = module ? annotations.find((a) => a.name === SHEET)?.arg : undefined;
    if (annotate) {
      const global = annotations.some((a) => a.name === WORKBOOK);
      if (global) entry.workbook = true;
      else entry.scope = sheet;
    } else if (local !== undefined) entry.scope = local;
    else if (scope !== undefined) entry.scope = scope;
    if (cell) {
      entry.cell = cell;
      const on = cell.sheet ?? (annotate ? sheet : entry.scope);
      if (on !== undefined) entry.cellSheet = on;
    }
    // `Name @C6 = ;` is a slot (and `@C5 = ;` clears a cell); a name needs a definition.
    if (entry.formula === "" && !cell) fail(`'${name}': empty definition`);
    if (name !== "") entries.push(entry);
    if (cell) cells.push(entry);
    doc = undefined;
    docSpan = undefined;
    extraDocs = [];
    annotations = [];
    pendingStart = undefined;
  }

  function scanFormula(): { formula: string; closed: boolean; map: number[] } {
    let out = "";
    let brace = 0;
    let bracket = 0;
    const map: number[] = [0, i];
    /** Starts a new run when the text skipped a comment, so `out` no longer follows `text`. */
    const resync = (): void => {
      const lastAt = map[map.length - 2]!;
      const lastSrc = map[map.length - 1]!;
      if (lastSrc + (out.length - lastAt) === i) return;
      if (lastAt === out.length) map.splice(map.length - 2, 2);
      map.push(out.length, i);
    };
    while (i < n) {
      resync();
      const c = text[i]!;
      if (bracket > 0) {
        if (c === "'" && i + 1 < n) {
          out += c + text[i + 1];
          advance(i + 2);
          continue;
        }
        if (c === "[") bracket++;
        else if (c === "]") bracket--;
        out += c;
        advance(i + 1);
        continue;
      }
      if (c === '"' || c === "'") {
        let j = i + 1;
        for (; j < n; j++) {
          if (text[j] === c) {
            if (text[j + 1] === c) j++;
            else break;
          }
        }
        const e = Math.min(j + 1, n);
        out += text.slice(i, e);
        advance(e);
        continue;
      }
      if (c === "/" && text[i + 1] === "/") {
        out = out.replace(/[ \t]+$/, "");
        skipLine();
        continue;
      }
      if (c === "/" && text[i + 1] === "*") {
        blockComment(2);
        out += " ";
        continue;
      }
      if (c === "[") bracket++;
      else if (c === "{") brace++;
      else if (c === "}") brace--;
      else if (c === ";" && brace <= 0) {
        return { formula: out, closed: true, map: finish() };
      }
      out += c;
      advance(i + 1);
    }
    return { formula: out, closed: false, map: finish() };

    function finish(): number[] {
      resync();
      if (text[i] === ";") advance(i + 1);
      return map;
    }
  }
}

/** Drops `lead` characters from the front of a formula map. */
function shiftMap(map: number[], lead: number): number[] {
  const out: number[] = [];
  for (let k = 0; k < map.length; k += 2) {
    const at = map[k]! - lead;
    const next = k + 2 < map.length ? map[k + 2]! - lead : Infinity;
    if (next <= 0) continue;
    if (at < 0) out.push(0, map[k + 1]! - at);
    else out.push(at, map[k + 1]!);
  }
  if (out.length === 0) out.push(0, (map[map.length - 1] ?? 0) + lead);
  return out;
}

export interface WritableEntry {
  name: string;
  /** In a module file: the sheet the name is local to, written `@sheet(Sheet)` above it. */
  scope?: string | undefined;
  /** In a sheet file: `@workbook` above the name (a workbook-scoped name). */
  workbook?: boolean;
  hidden?: boolean;
  /** The library base (`@from(lib #…)`), 6 hex digits. */
  from?: string | undefined;
  doc?: string | undefined;
  /** Display form; may span lines (LF). "" for a slot or a cleared cell. */
  formula: string;
  /** A cell statement's address as written after `@` (`formatCellAddress`, `#` included); `name` "" for an unnamed cell. */
  cell?: string;
}

/**
 * Whether `@sheet(…)`, `@scope(…)` or an address quotes this sheet: Excel's rule, as in
 * formulas (spec §14 issue 7; `sheetNeedsQuotes`). Reading accepts quotes either way.
 */
export function scopeNeedsQuotes(sheet: string): boolean {
  return sheetNeedsQuotes(sheet);
}

/** `@scope(Sheet)` for a sheet, `@workbook` for workbook scope (undefined): the old block lines. */
export function formatScope(sheet: string | undefined): string {
  if (sheet === undefined) return "@" + WORKBOOK;
  return `@${SCOPE}(${scopeArg(sheet)})`;
}

/** `@sheet(Sheet)`: a module name local to Sheet (and, alone on a line, a sheet file's sheet). */
export function formatSheetAnnotation(sheet: string): string {
  return `@${SHEET}(${scopeArg(sheet)})`;
}

function scopeArg(sheet: string): string {
  return scopeNeedsQuotes(sheet) ? "'" + sheet.split("'").join("''") + "'" : sheet;
}

/**
 * Renders one definition with its doc comment and annotations: doc comment, then
 * `@workbook` (sheet file) or `@sheet(…)` (module file), `@hidden`, `@from`.
 */
export function formatEntry(e: WritableEntry): string {
  const lines: string[] = [];
  if (e.doc !== undefined && e.doc !== "") lines.push(formatDoc(e.doc));
  if (e.scope !== undefined) lines.push(formatSheetAnnotation(e.scope));
  if (e.workbook) lines.push("@" + WORKBOOK);
  if (e.hidden) lines.push("@hidden");
  if (e.from !== undefined) lines.push(formatLibBase(e.from));
  const lhs = e.cell === undefined ? e.name : e.name === "" ? `@${e.cell}` : `${e.name} @${e.cell}`;
  lines.push(e.formula === "" ? `${lhs} = ;` : `${lhs} = ${e.formula};`);
  return lines.join("\n");
}

/** Whether a line holds only blanks. */
function blankText(s: string): boolean {
  return s.trim() === "";
}

/**
 * An old sheet file (`@scope(Sheet)`/`@workbook` blocks, `form: "sheet-blocks"`) in the
 * new form: the block lines go, each name of a `@workbook` block gets `@workbook` on the
 * line above it, and an address naming the file's own sheet loses its sheet (`@IS!C2` →
 * `@C2`). Everything else (comments, doc comments, layout) stays. Returns the reason
 * instead when the file cannot be converted: a block of another sheet's names, or a cell
 * of another sheet (they belong in that sheet's file).
 */
export function convertSheetBlocks(text: string, sheet: string): { text: string } | { reason: string } {
  const pm = parseModule(text, { sheet });
  if (pm.form !== "sheet-blocks") return { text };
  if (pm.diagnostics.length > 0) return { reason: `line ${pm.diagnostics[0]!.line}: ${pm.diagnostics[0]!.message}` };
  const low = (s: string) => s.toLowerCase();
  const other = pm.scopes.find((d) => d.scope !== undefined && low(d.scope) !== low(sheet));
  if (other) return { reason: `line ${other.line}: @scope(${other.scope}) holds names local to another sheet; move them to that sheet's file first` };
  const edits: { start: number; end: number; text: string }[] = [];
  const all = [...pm.entries, ...pm.cells.filter((c) => c.name === "")];
  for (const e of all) {
    if (e.cell?.sheet !== undefined) {
      if (low(e.cell.sheet) !== low(sheet)) return { reason: `line ${e.line}: ${e.name || "@" + e.cell.range} is a cell of ${e.cell.sheet}; move it to that sheet's file first` };
      edits.push({ start: e.cell.start + 1, end: e.cell.rangeStart, text: "" });
    }
    if (e.name !== "" && e.scope === undefined) {
      const ls = text.lastIndexOf("\n", e.start - 1) + 1;
      const indent = text.slice(ls, e.start);
      edits.push({ start: e.start, end: e.start, text: `@${WORKBOOK}\n${blankText(indent) ? indent : ""}` });
    }
  }
  return applyEdits(text, [...edits, ...directiveRemovals(text, pm.scopes)]);
}

/**
 * A module file with `@scope(Sheet)`/`@workbook` blocks (written before 2026-10-07) in the
 * per-name form: each name of a `@scope(Sheet)` block gets `@sheet(Sheet)` on the line above
 * it, an unnamed cell statement there names its sheet in its address (`@Sheet!C6`), and the
 * block lines go. Comments, doc comments and layout stay. A file without blocks comes back
 * as it is.
 */
export function convertModuleBlocks(text: string): { text: string } | { reason: string } {
  const pm = parseModule(text);
  if (pm.scopes.length === 0) return { text };
  if (pm.diagnostics.length > 0) return { reason: `line ${pm.diagnostics[0]!.line}: ${pm.diagnostics[0]!.message}` };
  const edits: { start: number; end: number; text: string }[] = [];
  for (const e of [...pm.entries, ...pm.cells.filter((c) => c.name === "")]) {
    if (e.scope === undefined || e.annotations.some((a) => a.name === SHEET)) continue;
    if (e.name === "") {
      if (e.cell && e.cell.sheet === undefined) edits.push({ start: e.cell.start + 1, end: e.cell.start + 1, text: formatCellAddress("", e.scope) });
      continue;
    }
    const ls = text.lastIndexOf("\n", e.offset - 1) + 1;
    const lead = text.slice(ls, e.offset);
    if (blankText(lead)) edits.push({ start: ls, end: ls, text: `${lead}${formatSheetAnnotation(e.scope)}\n` });
    else edits.push({ start: e.offset, end: e.offset, text: `${formatSheetAnnotation(e.scope)} ` });
  }
  return applyEdits(text, [...edits, ...directiveRemovals(text, pm.scopes)]);
}

function applyEdits(text: string, edits: { start: number; end: number; text: string }[]): { text: string } {
  edits.sort((a, b) => b.start - a.start || b.end - a.end);
  let out = text;
  for (const e of edits) out = out.slice(0, e.start) + e.text + out.slice(e.end);
  return { text: out };
}

/** Edits removing block directives: each one's line, and one blank line after it when one is left above. */
function directiveRemovals(text: string, scopes: readonly ScopeDirective[]): { start: number; end: number; text: string }[] {
  const edits: { start: number; end: number; text: string }[] = [];
  for (const d of scopes) {
    let s = text.lastIndexOf("\n", d.offset - 1) + 1;
    let t = text.indexOf("\n", d.end);
    t = t < 0 ? text.length : t + 1;
    if (!blankText(text.slice(s, d.offset)) || !blankText(text.slice(d.end, t))) {
      edits.push({ start: d.offset, end: d.end, text: "" });
      continue;
    }
    const before = text.slice(0, s);
    if ((s === 0 || before.endsWith("\n\n")) && text[t] === "\n") t++;
    else if (before.endsWith("\n\n") && t >= text.length) s--;
    edits.push({ start: s, end: t, text: "" });
  }
  return edits;
}
