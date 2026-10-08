// decompile: stored form (as in xl/workbook.xml and <f>) → the form Excel displays.
// compile:   display form → stored form.
//
// Both decide on the AST and then splice replacements into the original text, so the
// author's spacing and line breaks survive in both directions.

import { children, leftSpine, render, type Expr, type Formula, type Ref } from "./ast.js";
import { lookupFunction, type FunctionInfo } from "./catalogue.js";
import { FormulaError, type Diagnostic } from "./errors.js";
import { toCrLf } from "./format.js";
import { parse, stripPrefix } from "./parser.js";
import { quoteSheet, sheetNeedsQuotes, type Qualifier, type Span } from "./tokens.js";

export interface FormulaContext {
  /**
   * The sheet a defined name is scoped to. References to it are shown unqualified by
   * `decompile` and qualified again by `compile` (Excel qualifies them on entry, probe T06).
   * Leave it out for cell formulas and workbook-scoped names.
   */
  homeSheet?: string;
  /** Names scoped to `homeSheet`: `'Home'!X` decompiles to `X` and compiles back. */
  localNames?: Iterable<string>;
  /** Every defined name in the workbook (any scope). A call to one of them is not an unknown function. */
  names?: Iterable<string>;
  /**
   * The workbook's links to other workbooks (`<externalReferences>`, probe F10): `[1]` in
   * a stored formula is the first. With it, `decompile` shows `[1]Sheet1!A1` as
   * `[Other.xlsx]Sheet1!A1` and `[1]!Name` as `Other.xlsx!Name`, and `compile` turns them
   * back (a folder before the file's name is ignored); a workbook it has no link to is an
   * error, since only Excel writes a link's part. Without it, references to other
   * workbooks are kept as written.
   */
  links?: readonly WorkbookLink[];
}

/** One link to another workbook: its number in stored formulas and the file's name. */
export interface WorkbookLink {
  /** 1-based: `[1]`. */
  index: number;
  /** The file name, `Other.xlsx` (no folder). */
  book: string;
}

export interface CompileOptions extends FormulaContext {
  /** Emit calls to functions that are neither in the catalogue nor defined names. Default false:
   *  such a call is an error, because Excel would poison it as `_xludf.` (probe F6). */
  allowUnknownFunctions?: boolean;
  /** Write line breaks as CR LF, as Excel stores them. Default: keep as written. */
  crlf?: boolean;
}

export interface TransformResult {
  text: string;
  diagnostics: Diagnostic[];
}

/** Functions whose function-valued arguments may be a bare function name (an eta-reduced
 *  LAMBDA, stored as `_xleta.SUM`). */
const TAKES_LAMBDA = new Set(["BYROW", "BYCOL", "MAP", "REDUCE", "SCAN", "MAKEARRAY", "GROUPBY", "PIVOTBY"]);

/** Operators that bind tighter than prefix `@`: an `@x` operand there would change meaning. */
const TIGHTER_THAN_AT = new Set([":", ":.", ".:", ".:.", " ", ","]);

type Replace = (inner: (n: Expr) => string) => string;

function lowerSet(xs: Iterable<string> | undefined): Set<string> {
  const s = new Set<string>();
  if (xs) for (const x of xs) s.add(x.toLowerCase());
  return s;
}

function sameSheet(q: Qualifier | undefined, home: string | undefined): boolean {
  return (
    q !== undefined &&
    home !== undefined &&
    q.book === undefined &&
    q.sheet2 === undefined &&
    q.sheet !== undefined &&
    q.sheet.toLowerCase() === home.toLowerCase()
  );
}

function renderFormula(f: Formula, edits: Map<Expr, Replace>): string {
  const text = render(f.src, f.body, (node, inner) => edits.get(node)?.(inner));
  return f.src.slice(0, f.body.span.start) + text + f.src.slice(f.body.span.end);
}

function diag(severity: Diagnostic["severity"], code: string, span: Span, message: string): Diagnostic {
  return { severity, code, message, start: span.start, end: span.end };
}

// ---------------------------------------------------------------------------------------
// Structured references: `[#This Row],` (stored) ↔ `@` (displayed)
// ---------------------------------------------------------------------------------------

const THIS_ROW = "[#this row]";

function isSimpleColumn(s: string): boolean {
  if (s.length === 0) return false;
  for (const ch of s) if (!(ch === "_" || /[\p{L}\p{N}]/u.test(ch))) return false;
  return true;
}

/** `[#This Row],[Col]` → `@Col`; `#This Row` → `@`; otherwise undefined. */
export function structInnerToDisplay(inner: string): string | undefined {
  if (inner.trim().toLowerCase() === "#this row") return "@";
  if (inner.slice(0, THIS_ROW.length).toLowerCase() !== THIS_ROW) return undefined;
  let rest = inner.slice(THIS_ROW.length).trimStart();
  if (!rest.startsWith(",")) return undefined;
  rest = rest.slice(1).trimStart();
  if (!rest.startsWith("[")) return undefined;
  // `[Col]` alone, with a plain name, is shown without its brackets.
  if (rest.endsWith("]") && rest.indexOf("]") === rest.length - 1) {
    const col = rest.slice(1, -1);
    if (isSimpleColumn(col)) return "@" + col;
  }
  return "@" + rest;
}

/** `@Col` → `[#This Row],[Col]`; `@` → `#This Row`; otherwise undefined. */
export function structInnerToStored(inner: string): string | undefined {
  const t = inner.trimStart();
  if (!t.startsWith("@")) return undefined;
  const rest = t.slice(1).trim();
  if (rest === "") return "#This Row";
  if (rest.startsWith("[")) return "[#This Row]," + rest;
  return "[#This Row],[" + rest + "]";
}

// ---------------------------------------------------------------------------------------
// Trim references: `A1.:.A10` (displayed) ↔ `_xlfn._TRO_ALL(A1:A10)` (stored), probe F10
// ---------------------------------------------------------------------------------------

/** The hidden function Excel stores for each trim operator (measured, names and cells alike). */
const TRIM_FUNCTION: Record<string, string> = { ".:.": "_TRO_ALL", ":.": "_TRO_TRAILING", ".:": "_TRO_LEADING" };
const TRIM_OPERATOR: Record<string, string> = { _TRO_ALL: ".:.", _TRO_TRAILING: ":.", _TRO_LEADING: ".:" };
const RANGE_OPS = new Set([":", ":.", ".:", ".:."]);

/**
 * A range address's trim operator and the plain range: `A1:.A10` → `:.`, `A1:A10`.
 * Undefined for a plain range or a single cell. An address holds a `.` only next to its `:`.
 */
function splitTrim(address: string): { op: string; plain: string } | undefined {
  const k = address.indexOf(":");
  if (k < 0) return undefined;
  const lead = address[k - 1] === ".";
  const trail = address[k + 1] === ".";
  if (!lead && !trail) return undefined;
  return { op: (lead ? "." : "") + ":" + (trail ? "." : ""), plain: address.slice(0, lead ? k - 1 : k) + ":" + address.slice(trail ? k + 2 : k + 1) };
}

/** A plain range (area, columns, rows) a trim function's argument can be shown as, with the operator in it. */
function trimmable(e: Expr | undefined): e is Ref {
  return e !== undefined && e.kind === "ref" && e.refKind !== "cell" && e.refKind !== "error" && e.address.includes(":") && splitTrim(e.address) === undefined;
}

// ---------------------------------------------------------------------------------------
// Other workbooks: `[1]Sheet1!A1` (stored) ↔ `'[Other.xlsx]Sheet1'!A1` (displayed), probe F10
// ---------------------------------------------------------------------------------------

/** The text between a book part's last brackets: `1` of `[1]`, `Other.xlsx` of `C:\dir\[Other.xlsx]`. */
function bookInner(book: string): string | undefined {
  const close = book.lastIndexOf("]");
  const open = book.lastIndexOf("[", close);
  return open >= 0 && close > open ? book.slice(open + 1, close) : undefined;
}

function linkNumber(s: string): number | undefined {
  if (s === "" || s.length > 6) return undefined;
  for (const c of s) if (c < "0" || c > "9") return undefined;
  return Number(s);
}

/** A file name Excel quotes in a reference: anything beyond letters, digits, `_` and `.`. */
function bookNeedsQuotes(book: string): boolean {
  for (const ch of book) if (!(ch === "_" || ch === "." || /[\p{L}\p{N}\p{M}]/u.test(ch))) return true;
  return book === "";
}

/**
 * The display qualifier of a stored `[n]` one, as Excel's formula bar shows it while the
 * other workbook is open: `[Other.xlsx]Sheet1!`, `'[My Book.xlsx]Sheet 1'!` (quoted when
 * the file or a sheet needs it), `Other.xlsx!` for a workbook-level name. Undefined when
 * not a link of `links`, or when two links share the file's name (`[n]` then stays: the
 * name alone would not tell them apart on the way back).
 */
function linkDisplay(q: Qualifier | undefined, links: readonly WorkbookLink[] | undefined): string | undefined {
  if (!q || q.book === undefined || !links) return undefined;
  const n = linkNumber(bookInner(q.book) ?? "");
  const link = n !== undefined && n > 0 ? links.find((l) => l.index === n) : undefined;
  if (!link || link.book === "" || links.filter((l) => l.book.toLowerCase() === link.book.toLowerCase()).length > 1) return undefined;
  if (q.sheet === undefined) return (bookNeedsQuotes(link.book) ? `'${link.book.split("'").join("''")}'` : link.book) + "!";
  const sheets = q.sheet2 !== undefined ? [q.sheet, q.sheet2] : [q.sheet];
  const body = `[${link.book}]${sheets.join(":")}`;
  return bookNeedsQuotes(link.book) || sheets.some((s) => sheetNeedsQuotes(s)) ? `'${body.split("'").join("''")}'!` : `${body}!`;
}

/** `Other.xlsx`, `Plan.xlsm`: an extension of Excel's own files. */
export function looksLikeWorkbookFile(s: string): boolean {
  const dot = s.lastIndexOf(".");
  return dot > 0 && ["xlsx", "xlsm", "xlsb", "xls", "xltx", "xltm", "xlam"].includes(s.slice(dot + 1).toLowerCase());
}

/** The error for another workbook the file has no link to: only Excel writes a link's part. */
function noLink(book: string, links: readonly WorkbookLink[]): string {
  const have = links.length ? `its links: ${links.map((l) => `[${l.index}] ${l.book}`).join(", ")}` : "it has none";
  return `this workbook has no link to ${book} (${have}): Excel writes a link when a formula first names the other workbook, xln does not; type the reference once in Excel, save, then pull`;
}

/**
 * The stored qualifier of a reference into another workbook: `[n]Sheet1!`, `'[n]My Sheet'!`,
 * `[n]!`. `keep`: as written (no links known, `[0]!` the workbook itself, not another one).
 */
function linkStored(q: Qualifier, links: readonly WorkbookLink[] | undefined, nameOnly: boolean): { text: string } | { error: string } | "keep" {
  if (!links) return "keep";
  let n: number | undefined;
  let file: string | undefined;
  if (q.book !== undefined) {
    const inner = bookInner(q.book) ?? q.book;
    n = linkNumber(inner);
    if (n === 0) return "keep";
    if (n === undefined) file = inner;
  } else if (nameOnly && q.sheet !== undefined && q.sheet2 === undefined) {
    // `Other.xlsx!Name`: a workbook-level name of a linked file (a sheet named like a linked file reads as the file).
    const link = links.find((l) => l.book.toLowerCase() === q.sheet!.toLowerCase());
    if (link) return { text: `[${link.index}]!` };
    if (looksLikeWorkbookFile(q.sheet)) return { error: noLink(q.sheet, links) };
    return "keep";
  } else return "keep";
  if (n !== undefined && !links.some((l) => l.index === n)) return { error: noLink(`[${n}]`, links) };
  if (file !== undefined) {
    const named = file.toLowerCase();
    const link = links.find((l) => l.book.toLowerCase() === named);
    if (!link) return { error: noLink(file, links) };
    n = link.index;
  }
  if (q.sheet === undefined) return { text: `[${n}]!` };
  const sheets = q.sheet2 !== undefined ? [q.sheet, q.sheet2] : [q.sheet];
  const body = `[${n}]${sheets.join(":")}`;
  return { text: sheets.some((s) => sheetNeedsQuotes(s)) ? `'${body.split("'").join("''")}'!` : `${body}!` };
}

// ---------------------------------------------------------------------------------------
// decompile
// ---------------------------------------------------------------------------------------

/** Kinds that can carry a postfix `#` in display form. */
function spillable(e: Expr): boolean {
  return e.kind === "ref" || e.kind === "name";
}

/** Kinds that can follow a prefix `@` without parentheses and mean the same. */
function atable(e: Expr): boolean {
  switch (e.kind) {
    case "ref":
    case "name":
    case "structref":
    case "call":
    case "lambda":
    case "let":
    case "invoke":
    case "paren":
      return true;
    case "postfix":
      return e.op === "#";
    case "binary":
      return TIGHTER_THAN_AT.has(e.op);
    default:
      return false;
  }
}

/** Stored form → display form, with warnings (F6 prefix health, `_xludf.` poisoning). */
export function decompileWithDiagnostics(stored: string, ctx: FormulaContext = {}): TransformResult {
  const f = parse(stored);
  const edits = new Map<Expr, Replace>();
  /** A qualified reference's or name's qualifier as Excel shows it (another workbook by its file's name). */
  const shownQualifier = (q: Qualifier | undefined): string | undefined => linkDisplay(q, ctx.links) ?? excelQualifier(q);
  /** A qualified reference or name with its qualifier as Excel writes it. */
  const requote = (node: Expr, rest: string): void => {
    const q = shownQualifier((node as { qual?: Qualifier }).qual);
    if (q !== undefined) edits.set(node, () => q + rest);
  };
  const diagnostics: Diagnostic[] = [];
  const local = lowerSet(ctx.localNames);
  const home = ctx.homeSheet;

  const visit = (node: Expr, parent: Expr | undefined): void => {
    switch (node.kind) {
      case "ref":
        if (sameSheet(node.qual, home)) edits.set(node, () => node.address);
        else requote(node, node.address);
        break;
      case "name": {
        const { prefix, base } = stripPrefix(node.id);
        if (node.qual && !(sameSheet(node.qual, home) && local.has(node.id.toLowerCase()))) requote(node, node.id);
        if (prefix === "_xlpm." || prefix === "_xlop.") edits.set(node, () => base);
        else if (prefix === "_xleta.") {
          if (lookupFunction(base)) edits.set(node, () => base);
          else diagnostics.push(diag("warning", "unknown-function", node.span, `'${base}' is passed as a function but is not in the catalogue`));
        } else if (sameSheet(node.qual, home) && local.has(node.id.toLowerCase())) {
          edits.set(node, () => node.id);
        }
        break;
      }
      case "structref": {
        const d = structInnerToDisplay(node.inner);
        if (d !== undefined) edits.set(node, () => `${node.table}[${d}]`);
        break;
      }
      case "lambda":
        edits.set(node.fn, () => stripPrefix(node.fn.text).base);
        for (const p of node.params) {
          const { prefix, base } = stripPrefix(p.name.text);
          // `_xlop.y` is shown `[y]`; a bracketed `[y]` keeps its brackets around the name.
          edits.set(p.name, () => (prefix === "_xlop." ? `[${base}]` : base));
        }
        break;
      case "let":
        edits.set(node.fn, () => stripPrefix(node.fn.text).base);
        for (const b of node.bindings) edits.set(b.name, () => stripPrefix(b.name.text).base);
        break;
      case "call": {
        const fn = node.fn;
        if (fn.qual) {
          // `[1]!MyFn(…)`: a function in another workbook.
          const q = shownQualifier(fn.qual);
          if (q !== undefined) edits.set(fn, () => q + fn.text);
          break;
        }
        const { prefix, base } = stripPrefix(fn.text);
        const info = lookupFunction(base);
        const upper = base.toUpperCase();
        if (prefix === "_xlfn." || prefix === "_xlfn._xlws.") {
          if (!info) {
            diagnostics.push(diag("warning", "unknown-function", fn.span, `'${fn.text}': '${base}' is not in the catalogue; kept with its prefix`));
            break;
          }
          if (info.prefix !== prefix) {
            diagnostics.push(diag("warning", "wrong-prefix", fn.span, `'${fn.text}' is stored with ${prefix} but the catalogue says ${info.prefix || "no prefix"}`));
          }
          const arg = node.args[0];
          if (upper === "ANCHORARRAY" && node.args.length === 1 && arg && spillable(arg)) {
            edits.set(node, (inner) => inner(arg) + "#");
            break;
          }
          // `_xlfn._TRO_ALL(A1:A10)` is shown `A1.:.A10`, when its argument is a plain range and
          // no range operator joins it to its neighbour (`B1:_TRO_ALL(A1:A10)` would group otherwise).
          const trimOp = TRIM_OPERATOR[upper];
          const parentRange = parent?.kind === "binary" && RANGE_OPS.has(parent.op);
          if (trimOp !== undefined && node.args.length === 1 && trimmable(arg) && !parentRange) {
            const k = arg.address.indexOf(":");
            const address = arg.address.slice(0, k) + trimOp + arg.address.slice(k + 1);
            const q = sameSheet(arg.qual, home) ? "" : (shownQualifier(arg.qual) ?? arg.qual?.raw ?? "");
            edits.set(node, () => q + address);
            break;
          }
          const parentBinds = parent?.kind === "binary" && TIGHTER_THAN_AT.has(parent.op);
          const parentSpills = parent?.kind === "postfix" && parent.op === "#";
          if (upper === "SINGLE" && node.args.length === 1 && arg && atable(arg) && !parentBinds && !parentSpills) {
            edits.set(node, (inner) => "@" + inner(arg));
            break;
          }
          edits.set(fn, () => base);
        } else if (prefix === "_xlpm." || prefix === "_xlop.") {
          edits.set(fn, () => base);
        } else if (prefix === "_xludf.") {
          diagnostics.push(
            diag("warning", "poisoned", fn.span, `'${fn.text}': Excel did not recognise '${base}' when it loaded this formula and stored it as a user-defined function; it evaluates to #NAME? until rewritten`),
          );
        } else if (info && info.prefix !== "") {
          diagnostics.push(
            diag("warning", "bare-prefix", fn.span, `'${base}' is stored without its ${info.prefix} prefix: Excel shows #NAME? and re-saves it as _xludf.${base} (probe F6)`),
          );
        }
        break;
      }
      case "binary": {
        const { first, links } = leftSpine(node);
        visit(first, links[0]);
        for (const b of links) visit(b.right, b);
        return;
      }
      default:
        break;
    }
    for (const c of children(node)) visit(c, node);
  };
  visit(f.body, undefined);
  return { text: renderFormula(f, edits), diagnostics };
}

/** Stored form → display form. Throws `FormulaError` if the text does not parse. */
export function decompile(stored: string, ctx: FormulaContext = {}): string {
  return decompileWithDiagnostics(stored, ctx).text;
}

// ---------------------------------------------------------------------------------------
// compile
// ---------------------------------------------------------------------------------------


/**
 * A sheet qualifier as Excel writes it (`sheetNeedsQuotes`): `'BS'!` becomes `BS!`, a bare
 * `S1!` would become `'S1'!`. Undefined when it is already so, or names another workbook
 * (kept as written). Comparisons ignore optional quotes; the text follows Excel's rule.
 */
export function excelQualifier(q: Qualifier | undefined): string | undefined {
  if (!q || q.book !== undefined || q.sheet === undefined) return undefined;
  const sheets = q.sheet2 !== undefined ? [q.sheet, q.sheet2] : [q.sheet];
  const quote = sheets.some((x) => sheetNeedsQuotes(x));
  const body = sheets.map((x) => (quote ? x.split("'").join("''") : x)).join(":");
  const text = quote ? `'${body}'!` : `${body}!`;
  return text === q.raw ? undefined : text;
}

function storedName(info: FunctionInfo): string {
  return info.prefix + info.name;
}

/** Display form → stored form, with all diagnostics. `text` is empty when there are errors. */
export function compileWithDiagnostics(display: string, opts: CompileOptions = {}): TransformResult {
  const f = parse(display);
  const edits = new Map<Expr, Replace>();
  const diagnostics: Diagnostic[] = [];
  /**
   * A qualifier as Excel stores it: another workbook as its link's `[n]` (an error when it
   * has none), a sheet quoted by Excel's rule. Undefined when the text stays as written.
   */
  const storedQualifier = (q: Qualifier | undefined, span: Span, nameOnly: boolean): string | undefined => {
    if (!q) return undefined;
    const l = linkStored(q, opts.links, nameOnly);
    if (l === "keep") return excelQualifier(q);
    if ("error" in l) {
      diagnostics.push(diag("error", "external-link", span, l.error));
      return undefined;
    }
    return l.text === q.raw ? undefined : l.text;
  };
  /** A qualified reference or name with its qualifier as Excel writes it. */
  const requote = (node: Expr, rest: string): void => {
    const q = storedQualifier((node as { qual?: Qualifier }).qual, node.span, node.kind === "name");
    if (q !== undefined) edits.set(node, () => q + rest);
  };
  const names = lowerSet(opts.names);
  const local = lowerSet(opts.localNames);
  const home = opts.homeSheet;
  const homeQ = home !== undefined ? quoteSheet(home) + "!" : "";

  /** `scope` holds the lower-case LET/LAMBDA variables visible at this node. */
  const visit = (node: Expr, scope: ReadonlySet<string>, argOf: FunctionInfo | undefined): void => {
    const recurse = (n: Expr, s: ReadonlySet<string> = scope, a?: FunctionInfo): void => visit(n, s, a);
    switch (node.kind) {
      case "ref": {
        // `A1.:.A10` is stored `_xlfn._TRO_ALL(A1:A10)`, its qualifier inside (probe F10).
        const trim = node.refKind === "error" ? undefined : splitTrim(node.address);
        if (!trim) {
          if (!node.qual && home !== undefined) edits.set(node, () => homeQ + node.address);
          else requote(node, node.address);
          return;
        }
        const q = !node.qual ? homeQ : (storedQualifier(node.qual, node.span, false) ?? node.qual.raw);
        edits.set(node, () => `_xlfn.${TRIM_FUNCTION[trim.op]}(${q}${trim.plain})`);
        return;
      }
      case "name": {
        const { prefix, base } = stripPrefix(node.id);
        const key = base.toLowerCase();
        if (node.qual) {
          requote(node, node.id);
          return;
        }
        // A use of a parameter is `_xlpm.` even when the parameter is optional (`_xlop.`).
        if ((prefix === "" || prefix === "_xlpm." || prefix === "_xlop.") && scope.has(key)) {
          edits.set(node, () => "_xlpm." + base);
        } else if (prefix === "_xlpm." || prefix === "_xlop.") {
          diagnostics.push(diag("warning", "stray-parameter", node.span, `'${node.id}' is not a parameter of any enclosing LAMBDA or LET`));
        } else if (prefix === "_xleta.") {
          const info = lookupFunction(base);
          if (info) edits.set(node, () => "_xleta." + info.name);
          else diagnostics.push(diag("error", "unknown-function", node.span, `'${base}' is passed as a function but is not in the catalogue`));
        } else if (prefix === "") {
          const info = lookupFunction(base);
          if (argOf && TAKES_LAMBDA.has(argOf.name) && info && !names.has(key) && !info.internal) {
            edits.set(node, () => "_xleta." + info.name);
          } else if (home !== undefined && local.has(key)) {
            edits.set(node, () => homeQ + node.id);
          }
        }
        return;
      }
      case "structref": {
        const s = structInnerToStored(node.inner);
        if (s !== undefined) edits.set(node, () => `${node.table}[${s}]`);
        return;
      }
      case "postfix":
        if (node.op === "#") {
          edits.set(node, (inner) => "_xlfn.ANCHORARRAY(" + inner(node.operand) + ")");
        }
        recurse(node.operand);
        return;
      case "unary":
        if (node.op === "@") {
          edits.set(node, (inner) => "_xlfn.SINGLE(" + inner(node.operand) + ")");
        }
        recurse(node.operand);
        return;
      case "lambda": {
        edits.set(node.fn, () => "_xlfn.LAMBDA");
        const inner = new Set(scope);
        for (const p of node.params) inner.add(stripPrefix(p.name.text).base.toLowerCase());
        // An optional parameter `[y]` is stored `_xlop.y`, without brackets; its uses in the
        // body stay `_xlpm.y` (probe F9). `[_xlpm.y]` makes Excel drop the whole name.
        // The brackets lie outside the parameter's Ident, so the LAMBDA is rendered here.
        edits.set(node, (render) => {
          let out = render(node.fn);
          let pos = node.fn.span.end;
          for (const p of node.params) {
            const base = stripPrefix(p.name.text).base;
            out += f.src.slice(pos, p.span.start) + (p.optional ? "_xlop." : "_xlpm.") + base;
            pos = p.span.end;
          }
          return out + f.src.slice(pos, node.body.span.start) + render(node.body) + f.src.slice(node.body.span.end, node.span.end);
        });
        recurse(node.body, inner);
        return;
      }
      case "let": {
        edits.set(node.fn, () => "_xlfn.LET");
        const inner = new Set(scope);
        for (const b of node.bindings) {
          const base = stripPrefix(b.name.text).base;
          edits.set(b.name, () => "_xlpm." + base);
          // A binding's own value cannot see the binding: `LET(Rate, Rate*2, …)` reads the outer Rate.
          recurse(b.value, new Set(inner));
          inner.add(base.toLowerCase());
        }
        recurse(node.body, inner);
        return;
      }
      case "call": {
        const fn = node.fn;
        const { prefix, base } = stripPrefix(fn.text);
        const key = base.toLowerCase();
        let info: FunctionInfo | undefined;
        if (fn.qual) {
          // `[1]!MyFn(…)`: a function in another workbook; nothing to prefix.
          const q = storedQualifier(fn.qual, fn.span, true);
          if (q !== undefined) edits.set(fn, () => q + fn.text);
        } else if ((prefix === "" || prefix === "_xlpm." || prefix === "_xlop.") && scope.has(key)) {
          edits.set(fn, () => "_xlpm." + base);
        } else if (prefix === "_xludf.") {
          diagnostics.push(diag("warning", "poisoned", fn.span, `'${fn.text}' is a poisoned call (probe F6); it stays #NAME? until rewritten without '_xludf.'`));
        } else if (prefix === "_xlpm." || prefix === "_xlop.") {
          diagnostics.push(diag("warning", "stray-parameter", fn.span, `'${fn.text}' is not a parameter of any enclosing LAMBDA or LET`));
        } else if ((info = lookupFunction(base)) !== undefined) {
          const stored = storedName(info);
          if (prefix !== "" && prefix !== info.prefix) {
            diagnostics.push(diag("warning", "wrong-prefix", fn.span, `'${fn.text}' corrected to '${stored}' from the catalogue`));
          }
          if (stored !== fn.text) edits.set(fn, () => stored);
          if (info.prefixUnsure) {
            diagnostics.push(diag("warning", "prefix-unsure", fn.span, `the stored prefix of ${info.name} is not confirmed; check it once in a file saved by Excel`));
          }
          if (names.has(key)) {
            diagnostics.push(
              diag("warning", "builtin-collision", fn.span, `'${base}(…)' calls the built-in ${info.name}, not the defined name '${base}' (probe T12); rename the name`),
            );
          }
          const n = node.args.length;
          if (n < info.minArgs || n > info.maxArgs) {
            const want = info.minArgs === info.maxArgs ? `${info.minArgs}` : info.maxArgs >= 255 ? `at least ${info.minArgs}` : `${info.minArgs} to ${info.maxArgs}`;
            diagnostics.push(diag("warning", "arity", fn.span, `${info.name} takes ${want} argument${want === "1" ? "" : "s"}; it is given ${n}`));
          }
        } else if (prefix !== "") {
          diagnostics.push(diag("warning", "unknown-function", fn.span, `'${fn.text}': '${base}' is not in the catalogue; kept as written`));
        } else if (!names.has(key) && !opts.allowUnknownFunctions) {
          diagnostics.push(
            diag("error", "unknown-function", fn.span, `unknown function '${base}': it is neither in the function catalogue nor a defined name, and Excel would store it as _xludf.${base} (#NAME?)`),
          );
        }
        for (const a of node.args) recurse(a, scope, info);
        return;
      }
      case "binary": {
        const { first, links } = leftSpine(node);
        recurse(first);
        for (const b of links) {
          // A trim operator between operands that are not one range token (`A1:.INDEX(B:B,9)`,
          // `A1 .:. A10`): the same function around the plain range, spacing kept.
          const fnName = TRIM_FUNCTION[b.op];
          if (fnName !== undefined) {
            edits.set(b, (inner) => {
              const between = f.src.slice(b.left.span.end, b.right.span.start);
              const k = between.indexOf(b.op);
              return `_xlfn.${fnName}(${inner(b.left)}${between.slice(0, k)}:${between.slice(k + b.op.length)}${inner(b.right)})`;
            });
          }
          recurse(b.right);
        }
        return;
      }
      default:
        for (const c of children(node)) recurse(c);
    }
  };
  visit(f.body, new Set(), undefined);

  if (diagnostics.some((d) => d.severity === "error")) return { text: "", diagnostics };
  let text = renderFormula(f, edits);
  if (opts.crlf) text = toCrLf(text);
  return { text, diagnostics };
}

/** Display form → stored form. Throws `FormulaError` on bad syntax or an unknown function. */
export function compile(display: string, opts: CompileOptions = {}): string {
  const r = compileWithDiagnostics(display, opts);
  if (r.diagnostics.some((d) => d.severity === "error")) throw new FormulaError(display, r.diagnostics);
  return r.text;
}
