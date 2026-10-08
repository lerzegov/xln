// The source checker (M3d): one set of checks for a project's `.xln` files, run by the
// editor as you type and by the build before it plans. The build refuses exactly the
// errors found here (and, besides, what only the workbook can tell: conflicts with Excel's
// edits, a deletion or rename that would break a reference, see plan.ts).
//
// What it checks, with the problem on the exact text and a message that says what to
// write instead:
//
//   statements      the file's syntax, annotations, a doc comment over 255 characters
//                   (a formula over 8,192 characters is the build's: it counts the stored form);
//                   a LAMBDA's `@param` names against its parameters (a renamed parameter
//                   leaves a stale `@param` behind, and the Name Manager tooltip with it)
//   addresses       cell or range, `#` only on a named single cell, a bare address in a
//                   sheet file (the sheet is the file's), the address the last pull saw
//                   (read-only, but for its `#`)
//   names           defined twice; a new name Excel would refuse
//   scope           a workbook name on a sheet's cell read only there is a hint (remove
//                   `@workbook` to make it local); a scope changed in the source is a
//                   re-scope the build applies (info). Excel and the source both set scope
//                   (one is edited at a time; each sync carries the change, 2026-10-06)
//   formulas        every formula parsed and compiled with the real compiler; names nothing
//                   defines, another sheet's local name read bare, a sheet the workbook
//                   lacks, unknown functions, argument counts
//
// The context comes from the project: the workbook's sheets and Tables (the manifest in
// the editor, the workbook in the build), the lockfile, and `xln.config.json`.

import { walk, type Expr } from "../lang/ast.js";
import { stripPrefix } from "../lang/parser.js";
import { docParamSpans } from "../project/doc.js";
import { builtinCollision, catalogue, lookupFunction, xlmCollision } from "../lang/catalogue.js";
import { parseCell, quoteSheet } from "../lang/tokens.js";
import { compile, compileWithDiagnostics, decompile, looksLikeWorkbookFile, type WorkbookLink } from "../lang/transform.js";
import { callUses } from "../audit/walk.js";
import { columnName, formatCell, parseCell as parseCellAddress } from "../file/cellref.js";
import { cellCallee, classify, definitionTarget, sheetCellsOf, type CellCallee } from "../project/classify.js";
import { definitionHashLike, explicitSpill, splitNameKey, type LockCell, type Lockfile } from "../project/lockfile.js";
import { convertModuleBlocks, convertSheetBlocks, formulaToSource, sheetOfPath, type Annotation } from "../project/module.js";
import { moduleFileName, proposeModules, sheetFileName, SHEETS_DIR, UNMANAGED } from "../project/modules.js";
import { buildProvenanceTag, commentLength, COMMENT_MAX, docTagOverflow, FROM, moduleOfPath, parseLibBase } from "../project/provenance.js";
import type { NameKind } from "../project/types.js";
import { defKey, type Loc, type NameDef, type SourceFile, type SourceModel } from "./model.js";

export interface CheckContext {
  /** The workbook's sheets, in order; undefined when unknown (no sheet is checked). */
  sheets?: readonly string[];
  /** The workbook's Table names. */
  tables?: readonly string[];
  /** The last pull or build: the addresses and scopes it saw. */
  lock?: Lockfile;
  /**
   * Per sheet, the formulas that spilled beyond their cell when the workbook was saved
   * (`anchor` `C3`, `extent` `C3:G3`): the manifest's spill map in the editor, the
   * workbook's in the build and `xln check`. Without it no spill is checked.
   */
  spills?: Readonly<Record<string, readonly { anchor: string; extent: string }[]>>;
  /**
   * The workbook's links to other workbooks (the manifest's `externalLinks` in the editor,
   * the workbook's in the build and `xln check`): a formula naming a workbook it has no link
   * to is an error. Without it, references to other workbooks are not checked.
   */
  links?: readonly WorkbookLink[];
}

/** The spill map of a workbook, as `CheckContext.spills` takes it: spills over more than one cell. */
export function spillMap(sheets: readonly { name: string; spills: readonly { anchor: string; extent: string }[] }[]): Record<string, { anchor: string; extent: string }[]> {
  const out: Record<string, { anchor: string; extent: string }[]> = {};
  for (const s of sheets) {
    const real = s.spills.filter((x) => x.extent.includes(":") && x.extent.split(":")[0] !== x.extent.split(":")[1]);
    if (real.length) out[s.name] = real.map((x) => ({ anchor: x.anchor, extent: x.extent }));
  }
  return out;
}

/** A quick fix: replace `start`…`end` of the file with `text`. */
export interface Fix {
  title: string;
  start: number;
  end: number;
  text: string;
  /**
   * Edits of other project files made with this one (a statement moved to the file a pull
   * puts it in). `create`: the file does not exist yet, and `text` is all of it.
   */
  elsewhere?: { path: string; start: number; end: number; text: string; create?: boolean }[];
}

export interface Problem {
  path: string;
  start: number;
  end: number;
  message: string;
  /** `hint` is shown as dots under the text, not in the Problems panel (a spelling to match). */
  severity: "error" | "warning" | "info" | "hint";
  /** What the problem is, for tools and to match the audit's findings (`C5.other-sheet`). */
  code?: string;
  /** A quick fix (an edited address restored, a `#` inserted). */
  fix?: Fix;
  /** More quick fixes, the first one preferred (qualify, did you mean). */
  fixes?: Fix[];
  /** The name concerned (`Sheet!Name` or `Name`), for a problem of a name. */
  key?: string;
  /** A finding in a formula as the workbook has it since the last pull: the build leaves it (and does not list it). */
  inWorkbook?: true;
  /** Text that says nothing any more (a spent `@renamed`): the editor shows it faded. */
  unnecessary?: true;
}

const lower = (s: string) => s.toLowerCase();

const NAME_START = /[\p{L}_\\]/u;
const NAME_CHAR = /[\p{L}\p{N}\p{M}_\\.?]/u;

/**
 * Why Excel would refuse `name` as a defined name, or why xln does; undefined when it is
 * legal. A name spelled like a built-in is refused only when `definition` is a LAMBDA (or
 * not given): probe T12 found the clash only in calls (`Fact(5)` calls FACT), and a value
 * such as `Rate` is read as the name (probe F7's workbook has one); calling it is C3's.
 */
export function invalidName(name: string, definition?: string): string | undefined {
  if (name.length === 0) return "a name cannot be empty";
  if (name.length > 255) return "a name is at most 255 characters";
  if (!NAME_START.test(name[0]!)) return "a name starts with a letter, '_' or '\\'";
  for (const c of name) if (!NAME_CHAR.test(c)) return `'${c}' is not allowed in a name`;
  const up = name.toUpperCase();
  if (up === "R" || up === "C") return "'R' and 'C' are reserved (R1C1 references)";
  if (parseCell(name)) return "a name cannot look like a cell reference";
  if (/^R[0-9]*C[0-9]*$/i.test(name) || /^R[0-9]+$/i.test(name) || /^C[0-9]+$/i.test(name)) return "a name cannot look like an R1C1 reference";
  if (up.startsWith("_XL")) return "names starting with _xl are Excel's own";
  if (builtinCollision(name) && (definition === undefined || classify(definition).kind === "lambda")) return `'${name}' collides with the built-in function ${up}: a call to it calls the built-in (probe T12)`;
  return undefined;
}

/**
 * The warning for a LAMBDA named like an Excel 4.0 macro function that is not a worksheet
 * function (`Group`, `Get.Cell`, `Evaluate`); undefined otherwise. Only LAMBDAs, as for
 * the built-in rule: the clash is in calls. Not the command equivalents (`Open`, `Save`):
 * see xlmCollision. A warning, not a refusal: AFE issue #10 reports it, xln has not
 * measured it.
 */
export function xlmNameWarning(name: string, definition: string): string | undefined {
  if (xlmCollision(name) === undefined || classify(definition).kind !== "lambda") return undefined;
  return `${name} is also an Excel 4.0 macro function: Excel may call that instead or refuse the name (AFE #10, not measured)`;
}

/** `@renamed(Old)` → the old name and scope (`null`: workbook scope; undefined: the same scope). */
export function renamedFrom(annotations: readonly Annotation[]): { name: string; scope: string | null | undefined } | undefined {
  const a = annotations.find((x) => x.name === "renamed");
  const arg = (a?.arg ?? "").trim();
  if (!a || arg === "") return undefined;
  const bang = arg.lastIndexOf("!");
  if (bang < 0) return { name: arg, scope: undefined };
  let sheet = arg.slice(0, bang).trim();
  if (sheet.length >= 2 && sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).split("''").join("'");
  return { name: arg.slice(bang + 1).trim(), scope: sheet === "" ? null : sheet };
}

/**
 * Whether the rename a name's `@renamed(…)` records (`r`, from `renamedFrom`) is built: the
 * lockfile (the last pull or build) has the name in its scope and not the old one. The
 * build ignores such a note (plan.ts), the checker says it can go (`renamed-built`), and a
 * writing build removes it (consumedRenamedEdits). A change of spelling only
 * (`@renamed(rate)` on `Rate`) is never built: it is not a rename. `has` answers whether
 * the lockfile has a name in a scope (undefined: workbook scope), ignoring case.
 */
export function renameBuilt(name: string, scope: string | undefined, r: { name: string; scope: string | null | undefined }, has: (name: string, scope: string | undefined) => boolean): boolean {
  const oldScope = r.scope === undefined ? scope : (r.scope ?? undefined);
  if (lower(r.name) === lower(name) && lower(oldScope ?? "") === lower(scope ?? "")) return false;
  return !has(r.name, oldScope) && has(name, scope);
}

/**
 * What removing annotation `a` takes out of `text`: its whole line (line break included)
 * when it stands alone on it, as `xln rename`, Rename Symbol and most hands write it;
 * otherwise the annotation and the blanks after it (`@renamed(Old) @hidden` keeps
 * `@hidden`), or, at the end of a line, the blanks before it.
 */
export function annotationRemoval(text: string, a: { offset: number; end: number }): { start: number; end: number } {
  const lineStart = text.lastIndexOf("\n", a.offset - 1) + 1;
  const nl = text.indexOf("\n", a.end);
  const lineEnd = nl < 0 ? text.length : nl;
  const blank = (s: string) => s.trim() === "";
  if (blank(text.slice(lineStart, a.offset)) && blank(text.slice(a.end, lineEnd))) return { start: lineStart, end: nl < 0 ? text.length : nl + 1 };
  const isBlank = (c: string | undefined) => c === " " || c === "\t";
  let end = a.end;
  while (isBlank(text[end])) end++;
  if (end < lineEnd && text[end] !== "\r") return { start: a.offset, end };
  let start = a.offset;
  while (start > lineStart && isBlank(text[start - 1])) start--;
  return { start, end: a.end };
}

/** The source name carrying `@renamed(…)` of the lockfile key `was` (`Rate`, `S2!Loc`), if any. */
export function renamedTo(model: SourceModel, was: string): NameDef | undefined {
  const w = lower(was);
  return model.defs.find((d) => {
    const r = renamedFrom(d.entry.annotations);
    if (!r) return false;
    const scope = r.scope === undefined ? d.scope : (r.scope ?? undefined);
    return lower(defKey(r.name, scope)) === w;
  });
}

/** `Sheet!` as a formula writes it: Excel's quoting rule (`sheetNeedsQuotes`). */
export function sheetPrefix(sheet: string): string {
  return `${quoteSheet(sheet)}!`;
}

/** Damerau–Levenshtein distance at most 1, ignoring case. */
export function oneEditAway(a: string, b: string): boolean {
  const x = lower(a);
  const y = lower(b);
  if (x === y) return false;
  if (Math.abs(x.length - y.length) > 1) return false;
  let i = 0;
  while (i < x.length && i < y.length && x[i] === y[i]) i++;
  if (x.length === y.length) {
    if (x.slice(i + 1) === y.slice(i + 1)) return true;
    return x[i] === y[i + 1] && x[i + 1] === y[i] && x.slice(i + 2) === y.slice(i + 2);
  }
  return x.length < y.length ? x.slice(i) === y.slice(i + 1) : x.slice(i + 1) === y.slice(i);
}

function plural(n: number, word: string): string {
  return `${n} ${word}${n === 1 ? "" : "s"}`;
}

function wantText(min: number, max: number): string {
  if (max >= 255) return `at least ${plural(min, "argument")}`;
  if (min === max) return plural(min, "argument");
  return `${min} to ${max} arguments`;
}

export const KIND_LABEL: Record<NameKind, string> = {
  constant: "constant",
  range: "range",
  spill: "spill",
  table: "Table reference",
  formula: "formula",
  lambda: "LAMBDA",
  unparsed: "does not parse",
};

function span(at: { start: number; end: number }): { start: number; end: number } {
  return { start: at.start, end: at.end };
}

function didYouMean(at: { start: number; end: number }, names: string[]): Fix[] {
  const uniq = [...new Set(names)].slice(0, 3);
  return uniq.map((n) => ({ title: `Did you mean ${n}?`, ...span(at), text: n }));
}

/** The codes of the checks as you type the audit also makes (`xln check`): the editor shows one of each. */
export const LIVE_CODES: ReadonlySet<string> = new Set(["C4.unknown-name", "C5.other-sheet", "C6.lambda-arity", "C6.builtin-arity", "C6.not-a-function", "C6.not-a-lambda"]);

/** The lockfile, indexed for the checks. */
interface LockIndex {
  /** Lower-case key → the key as written. */
  names: Map<string, string>;
  /** Lower-case name → keys (any scope). */
  byName: Map<string, string[]>;
  /** Lower-case name key → its cell statement. */
  namedCells: Map<string, LockCell>;
  /** `sheet!range`, lower case: the unnamed statements. */
  unnamed: Set<string>;
  /** Whether the lockfile records cells at all (format 2 and later). */
  hasCells: boolean;
}

function indexLock(lock: Lockfile): LockIndex {
  const ix: LockIndex = { names: new Map(), byName: new Map(), namedCells: new Map(), unnamed: new Set(), hasCells: lock.cells !== undefined };
  for (const key of Object.keys(lock.names)) {
    ix.names.set(lower(key), key);
    const n = lower(splitNameKey(key).name);
    const l = ix.byName.get(n) ?? [];
    l.push(key);
    ix.byName.set(n, l);
  }
  for (const c of Object.values(lock.cells ?? {})) {
    if (c.name !== undefined) ix.namedCells.set(lower(c.name), c);
    else ix.unnamed.add(lower(`${c.sheet}!${c.range}`));
  }
  return ix;
}

const indexes = new WeakMap<Lockfile, LockIndex>();
function lockIndex(lock: Lockfile): LockIndex {
  let ix = indexes.get(lock);
  if (!ix) indexes.set(lock, (ix = indexLock(lock)));
  return ix;
}

const MODULE_START = /[\p{L}_]/u;
const MODULE_CHAR = /[\p{L}\p{N}\p{M}_]/u;

/**
 * Why a file at `path` (project-relative) has no place in a project's `names/`, or
 * undefined when it has one (or is not below `names/`). VS Code cannot refuse a file
 * created in the Explorer, so the editor flags it and the build refuses it (M3e):
 *
 * - only `.xln` files: anything else is never read, and the author may think it is;
 * - `names/sheets/<Sheet>.xln` for a sheet the workbook has (sheets are made in Excel and
 *   come with a pull), no folders below `sheets/`;
 * - modules directly in `names/`, named after their prefix (`FIN.xln`; `~2` as a pull adds
 *   on a clash; `_unmanaged.xln`). Pulled prefixes keep only letters, digits and `_` (a
 *   pull writes `_` for the characters a file name cannot hold).
 *
 * Hidden files (`.DS_Store`) are the system's and not judged. `sheets`: the workbook's,
 * when known.
 */
export function strayFile(path: string, sheets?: readonly string[]): string | undefined {
  if (!path.startsWith("names/")) return undefined;
  const parts = path.split("/");
  const file = parts[parts.length - 1]!;
  if (file.startsWith(".")) return undefined;
  if (!file.toLowerCase().endsWith(".xln")) {
    return `${path}: only .xln files belong in names/ (the build reads nothing else); move it out of the project`;
  }
  if (parts[1] === "sheets" && parts.length > 2) {
    if (parts.length > 3) return `${path}: names/sheets/ holds one file per sheet, no folders; move it to names/sheets/<Sheet>.xln`;
    const sheet = sheetOfPath(path);
    // A name a pull would not write (`IS~2.xln`) is the `sheet-file` warning: `@sheet(…)` may name it.
    if (sheet !== undefined && sheets && !sheets.some((s) => lower(s) === lower(sheet))) {
      return `there is no sheet ${sheet} in the workbook: sheets are created in Excel, then pulled (a sheet file is named after an existing sheet)`;
    }
    return undefined;
  }
  if (parts.length > 2) return `${path}: a module lives directly in names/ (names/<Prefix>.xln), and sheet files in names/sheets/; the build reads no other folder`;
  let base = file.slice(0, -4);
  const tilde = base.lastIndexOf("~");
  if (tilde > 0 && [...base.slice(tilde + 1)].every((c) => c >= "0" && c <= "9") && tilde < base.length - 1) base = base.slice(0, tilde);
  if (base === "_unmanaged") return undefined;
  const chars = [...base];
  if (!MODULE_START.test(chars[0]!) || !chars.every((c) => MODULE_CHAR.test(c))) {
    return `${file}: a module file is named after its prefix (names/FIN.xln for FIN.NPV, …): letters, digits and _, starting with a letter or _; rename it (xln: New module checks the prefix)`;
  }
  return undefined;
}

/** The checker's problem for a file `strayFile` rejects: on the file's first line. */
function strayProblem(path: string, text: string, sheets?: readonly string[]): Problem | undefined {
  const message = strayFile(path, sheets);
  if (message === undefined) return undefined;
  const e = text.indexOf("\n");
  return { path, start: 0, end: e < 0 ? text.length : e, severity: "error", code: "stray-file", message };
}

/** Problems of every names file of the model, files in path order. */
export function checkProject(model: SourceModel, ctx: CheckContext = {}): Problem[] {
  const out: Problem[] = [];
  for (const path of [...model.files.keys()].sort()) if (path.startsWith("names/") && path.endsWith(".xln")) out.push(...checkFile(model, path, ctx));
  return out;
}

/** The problems of one file. */
export function checkFile(model: SourceModel, path: string, ctx: CheckContext = {}): Problem[] {
  const file = model.files.get(path);
  if (!file) return [];
  const out: Problem[] = [];
  const pm = file.parsed;
  const lock = ctx.lock ? lockIndex(ctx.lock) : undefined;
  const sheetList = ctx.sheets;
  const sheetSet = sheetList ? new Set(sheetList.map(lower)) : undefined;

  // ---- the file ---------------------------------------------------------------------
  for (const d of pm.diagnostics) {
    const start = file.lines.offset({ line: d.line - 1, character: d.col - 1 });
    out.push({ path, start, end: Math.min(start + 1, file.text.length), message: d.message, severity: "error", code: "syntax" });
  }
  const firstLineEnd = (() => {
    const e = file.text.indexOf("\n");
    return e < 0 ? file.text.length : e;
  })();
  if (path.startsWith("names/sheets/") && pm.sheet === undefined) {
    out.push({ path, start: 0, end: firstLineEnd, severity: "warning", code: "sheet-file", message: `${path}: this file name is not one a pull writes for a sheet, so its names are read as a module's (workbook scope unless @sheet(Sheet) is above them); name it after its sheet (names/sheets/<Sheet>.xln)` });
  }
  // `@sheet(…)` in the file names its sheet (a pull's `IS~2.xln`): then the name is no claim.
  const stray = pm.sheet !== undefined && pm.sheet !== sheetOfPath(path) ? undefined : strayProblem(path, file.text, sheetList);
  if (stray) out.push(stray);
  else if (pm.sheet !== undefined && sheetSet && !sheetSet.has(lower(pm.sheet))) {
    out.push({ path, start: 0, end: firstLineEnd, severity: "error", code: "unknown-sheet", message: `the workbook has no sheet '${pm.sheet}' (a sheet file's name gives its sheet): rename the file after an existing sheet, or move its statements` });
  }
  if (pm.form === "sheet-blocks") {
    const first = pm.scopes[0]!;
    const conv = convertSheetBlocks(file.text, pm.sheet!);
    const p: Problem = {
      path,
      start: first.offset,
      end: first.end,
      severity: "info",
      code: "old-blocks",
      message: `this sheet file has @scope/@workbook blocks (written before M3d). Every name in a sheet file is local to ${pm.sheet} unless @workbook is on the line above it${"reason" in conv ? `; it cannot be converted as it is (${conv.reason})` : ": convert the file"}`,
    };
    if ("text" in conv) p.fix = { title: "Convert to per-name @workbook", start: 0, end: file.text.length, text: conv.text };
    out.push(p);
  }
  // Module files: per-name `@sheet(…)` since 2026-10-07. Old blocks are read as they are,
  // with the conversion offered; a `@workbook` line in a file without blocks says nothing.
  if (pm.form === "module" && !path.startsWith(`names/${SHEETS_DIR}/`)) {
    const first = pm.scopes.find((d) => d.scope !== undefined);
    if (first) {
      const conv = convertModuleBlocks(file.text);
      const p: Problem = {
        path,
        start: first.offset,
        end: first.end,
        severity: "info",
        code: "old-blocks",
        message: `this module file has @scope/@workbook blocks (written before 2026-10-07). A module name is workbook-scoped unless @sheet(Sheet) is on the line above it${"reason" in conv ? `; it cannot be converted as it is (${conv.reason})` : ": convert the file"}`,
      };
      if ("text" in conv) p.fix = { title: "Convert to per-name @sheet", start: 0, end: file.text.length, text: conv.text };
      out.push(p);
    } else {
      for (const d of pm.scopes) {
        const ls = file.text.lastIndexOf("\n", d.offset - 1) + 1;
        let le = file.text.indexOf("\n", d.end);
        le = le < 0 ? file.text.length : le;
        const alone = file.text.slice(ls, d.offset).trim() === "" && file.text.slice(d.end, le).trim() === "";
        out.push({
          path,
          start: d.offset,
          end: d.end,
          severity: "hint",
          code: "redundant-workbook",
          message: "@workbook in a module file says nothing: a module name is workbook-scoped unless @sheet(Sheet) is above it",
          fix: alone ? { title: "Remove @workbook", start: ls, end: le < file.text.length ? le + 1 : le, text: "" } : { title: "Remove @workbook", start: d.offset, end: d.end, text: "" },
        });
      }
    }
  }
  // Directives in a file kind where they mean nothing (§14 issue 4): said, not ignored silently.
  if (path === `names/${UNMANAGED}.xln`) {
    for (const d of pm.scopes) {
      if (d.scope === undefined) continue;
      out.push({ path, start: d.offset, end: d.end, severity: "warning", code: "directive", message: `@scope(${d.scope}) in _unmanaged.xln, which holds the workbook-scoped names no module owns: a pull puts names local to ${d.scope} in names/sheets/${sheetFileName(d.scope)} (or in their module's file); move them there` });
    }
  }
  if (pm.form !== "sheet" && sheetSet) {
    for (const s of pm.scopes) {
      if (s.scope !== undefined && !sheetSet.has(lower(s.scope))) out.push({ path, start: s.offset, end: s.end, severity: "error", code: "unknown-sheet", message: `@scope(${s.scope}): the workbook has no sheet '${s.scope}'` });
    }
  }

  // ---- each statement -----------------------------------------------------------------
  const scope = model.scopeIndex();
  const tables = new Set((ctx.tables ?? []).map(lower));
  const allSheets = new Set([...(sheetSet ?? []), ...scope.locals.keys()]);
  const allNames = model.defs.map((d) => d.name);
  const seenCells = new Map<string, NameDef>();
  const modules = proposeModules(model.defs.map((d) => d.name));
  for (const def of [...model.defs, ...model.cellDefs]) {
    if (def.file !== file) continue;
    const e = def.entry;
    const label = def.name !== "" ? def.name : `@${e.cell!.range}`;
    const a = model.analysis(def);
    const home = model.homeOf(def);

    // Annotations.
    let froms = 0;
    for (const an of e.annotations) {
      if (an.name === "hidden" || (an.name === "workbook" && pm.form === "sheet")) continue;
      const at = { start: an.offset, end: an.end };
      if (an.name === "sheet" && pm.form === "module") {
        // A module name local to a sheet: the sheet must exist.
        if (sheetSet && an.arg !== undefined && !sheetSet.has(lower(an.arg))) out.push({ path, ...at, severity: "error", code: "unknown-sheet", message: `@sheet(${an.arg}): the workbook has no sheet '${an.arg}'` });
        continue;
      }
      if (an.name === "renamed") {
        if ((an.arg ?? "").trim() === "") out.push({ path, ...at, severity: "error", code: "annotation", message: "@renamed needs the previous name: @renamed(OldName)" });
        else if (lock && def.name !== "" && an === e.annotations.find((x) => x.name === "renamed")) {
          // A note whose rename is in the workbook already (built by an older build, a browser
          // build, another tool, or written by hand): spent. A hint, as `redundant-workbook`.
          const r = renamedFrom(e.annotations)!;
          if (renameBuilt(def.name, def.scope, r, (n, s) => lock.names.has(lower(defKey(n, s))))) {
            const written = file.text.slice(an.offset, an.end);
            out.push({ path, ...at, severity: "hint", code: "renamed-built", key: def.key, unnecessary: true, message: `${written}: the rename is built; this line can go`, fix: { title: `Remove ${written}`, ...annotationRemoval(file.text, an), text: "" } });
          }
        }
        continue;
      }
      if (an.name === FROM) {
        // The library base: written by Insert, Take and Publish; the build carries it in the
        // provenance tag, which only names of a module file get.
        const b = parseLibBase(an.arg);
        if ("error" in b) out.push({ path, ...at, severity: "error", code: "annotation", message: b.error });
        else if (++froms > 1) out.push({ path, ...at, severity: "error", code: "annotation", message: "@from is given twice: a copy has one library base; keep one" });
        else if (def.name === "" || moduleOfPath(path) === undefined)
          out.push({ path, ...at, severity: "warning", code: "annotation", message: "@from records the library base of a workbook name in a module file (names/FN.xln); here no provenance tag carries it into the workbook, so a pull drops it" });
        continue;
      }
      const known = pm.form === "sheet" ? "@hidden, @workbook, @renamed(Old), @from(lib #…); and on a line of its own @sheet(Name)" : "@hidden, @sheet(Sheet), @renamed(Old), @from(lib #…)";
      out.push({ path, ...at, severity: "warning", code: "annotation", message: `@${an.name} is not an annotation xln knows (${known}); it is ignored` });
    }
    // Doc comments and annotations may come in any order; but one statement has one doc
    // comment (the Name Manager's), and the parser keeps the last.
    for (const x of e.extraDocs ?? []) {
      let end = x.end;
      while (file.text[end] === " " || file.text[end] === "\t") end++;
      if (file.text[end] === "\n") end++;
      out.push({ path, start: x.start, end: x.end, severity: "warning", code: "doc-twice", ...(def.key ? { key: def.key } : {}), message: `${label}: two doc comments: only the last is kept (it is the Name Manager comment); merge them, or remove this one`, fix: { title: "Remove this doc comment", start: x.start, end, text: "" } });
    }
    if (def.name !== "" && e.doc !== undefined && commentLength(e.doc) > COMMENT_MAX) {
      const close = file.text.indexOf("*/", e.start);
      out.push({
        path,
        start: e.start,
        end: close >= 0 && close < e.offset ? close + 2 : e.offset,
        severity: "error",
        code: "comment-length",
        key: def.key,
        message: `${label}: the doc comment has ${commentLength(e.doc)} characters (a line break counts 2); the Name Manager takes ${COMMENT_MAX}: shorten it`,
      });
    }
    // Room for the provenance tag the build adds to every name of a module file: without it
    // the build drops the tag, and with it a `@from` base the workbook would carry.
    const tagOver = def.name !== "" ? docTagOverflow(e.doc, buildProvenanceTag(path, file.text, e.from)) : undefined;
    if (tagOver) {
      const close = file.text.indexOf("*/", e.start);
      out.push({ path, start: e.start, end: close >= 0 && close < e.offset ? close + 2 : e.offset, severity: "warning", code: "provenance", key: def.key, message: `${label}: ${tagOver.message}` });
    }

    if (def.name !== "" && e.doc !== undefined && a.body) for (const p of docParamProblems(file, def, a.body, label)) out.push(p);

    // The cell statement's address.
    if (e.cell) {
      for (const p of addressProblems(def, file, ctx, lock, sheetSet)) out.push(p);
      const sp = def.name !== "" ? spillProblem(def, ctx) : undefined;
      if (sp) out.push(sp);
      if (def.name === "" && e.cellSheet !== undefined) {
        const k = lower(`${e.cellSheet}!${e.cell.range}`);
        const first = seenCells.get(k);
        if (first) out.push({ path, start: e.cell.start, end: e.cell.end, severity: "error", code: "duplicate", message: `@${e.cell.range}: another statement writes this cell (${first.file.path}:${first.entry.line}): keep one` });
        else seenCells.set(k, def);
      }
    }

    // The name.
    if (def.name !== "") {
      const first = model.lookup(def.key);
      if (first && first !== def) {
        out.push({ ...model.nameLoc(def), severity: "error", code: "duplicate", key: def.key, message: `${def.key} is also defined at ${first.file.path}:${first.entry.line}: keep one (a name can be defined once per scope)` });
      } else {
        // A workbook name no module owns, fixed to one sheet's cells: a pull writes it in that sheet's file.
        const cellsOf = def.scope === undefined && !e.cell && modules.get(def.name) === undefined ? sheetCellsOf(e.formula, ctx.sheets) : undefined;
        for (const p of nameProblems(model, def, file, lock, cellsOf)) out.push(p);
        const p = placementProblem(model, def, file, modules, lock, cellsOf);
        if (p) out.push(p);
      }
    }

    // The formula. Its findings are errors where the source changed it since the last pull;
    // a formula as the workbook has it is the workbook's (the build leaves it, the audit
    // reports it), so there they are warnings.
    const found: Problem[] = [];
    /** Findings on a name the last pull had, which the source renamed, moved or deleted: the reader is the source's to fix. */
    const broken = new Map<Problem, string>();
    for (const er of a.errors) {
      const sp = er.span.end > er.span.start ? er.span : { start: er.span.start, end: er.span.start + 1 };
      const at = model.loc(def, { start: Math.min(sp.start, e.formula.length), end: Math.min(sp.end, e.formula.length) });
      if (at.end <= at.start) at.end = at.start + 1;
      found.push({ ...at, message: `${label}: ${er.message}`, severity: "error", code: "syntax" });
    }
    let compiled: ReturnType<typeof compileWithDiagnostics> | undefined;
    if (a.body) {
      liveChecks(model, def, a, home, ctx, scope, tables, allSheets, found, broken);
      // Sheets read by cell references (`Ratios!B4`): the workbook must have them.
      if (sheetSet) {
        walk(a.body, (n) => {
          if (n.kind !== "ref" || n.qual?.sheet === undefined || n.qual.book !== undefined) return;
          for (const s of [n.qual.sheet, n.qual.sheet2]) {
            if (s === undefined || sheetSet.has(lower(s))) continue;
            const at = model.loc(def, n.span);
            const near = (sheetList ?? []).filter((x) => oneEditAway(x, s));
            found.push({ ...at, severity: "error", code: "unknown-sheet", message: `${s}!${n.address}: the workbook has no sheet '${s}'${near.length ? ` (did you mean ${near[0]}?)` : ""}` });
          }
        });
      }
      // The compiler's own errors (unknown functions, `_xleta.`), where the checks above said nothing there.
      const links = ctx.links !== undefined ? { links: ctx.links } : {};
      const opts = e.cell ? { names: allNames, ...links } : { names: allNames, ...links, ...(def.scope !== undefined ? { homeSheet: def.scope, localNames: (scope.locals.get(lower(def.scope)) ?? []).map((d) => d.name) } : {}) };
      try {
        compiled = compileWithDiagnostics(e.formula, opts);
        if (compiled.diagnostics.some((g) => g.severity === "error")) compiled = { ...compiled, text: compileWithDiagnostics(e.formula, { ...opts, allowUnknownFunctions: true }).text };
      } catch {
        compiled = undefined;
      }
      for (const g of compiled?.diagnostics ?? []) {
        // A prefixed call the catalogue does not know (`_xlfn.FOO`) is refused here, before
        // any writing: Excel would show #NAME? (and the read-back's stored-form check refuse it).
        if (g.severity !== "error" && g.code !== "unknown-function") continue;
        const at = model.loc(def, { start: g.start, end: Math.max(g.end, g.start + 1) });
        if (found.some((p) => p.severity === "error" && p.start < at.end && at.start < p.end)) continue;
        const message = g.severity === "error" ? g.message : `${g.message.replace(/; kept as written$/, "")}: Excel would show #NAME?; write a function the catalogue knows (without a prefix: the build adds it)`;
        found.push({ ...at, severity: "error", code: g.code, message: `${label}: ${message}` });
      }
    }
    const asInWorkbook = found.some((p) => p.severity === "error") && ctx.lock !== undefined && unchangedSinceLock(def, compiled?.text || e.formula, ctx.lock);
    for (const p of found) {
      const was = broken.get(p);
      // A formula unchanged since the last pull is the workbook's, and its findings are
      // warnings, unless what it reads is gone because of the source: a rename, re-scope or
      // deletion breaks it, and the build refuses on it (§14 issue 10). The formula's text is
      // the same; what it means is not. The code is the build's for the same refusal.
      const to = was !== undefined ? renamedTo(model, was) : undefined;
      if (was !== undefined && asInWorkbook && to !== undefined) {
        out.push({ ...p, code: "in-use", message: `${p.message}: the source renames ${was} to ${to.name}, so write ${to.name} here (xln rename, or Rename Symbol in the editor, renames every reader); the build refuses until then` });
      } else if (was !== undefined && asInWorkbook) out.push({ ...p, code: "in-use", message: `${p.message}: the last pull had ${was}, which the source renames, moves or deletes, so this formula breaks and the build refuses (change the formula, or keep the name)` });
      else if (asInWorkbook && p.severity === "error") out.push({ ...p, severity: "warning", inWorkbook: true, message: `${p.message} (as in the workbook since the last pull: the build leaves it; fix it here to have it written)` });
      else out.push(p);
    }
  }
  // In text order (a stable sort: problems at one place keep the order they were found in).
  return out.sort((a, b) => a.start - b.start);
}

/**
 * A LAMBDA's doc comment against its parameters (author's idea, 2026-10-07: FN.SPREAD's
 * parameter renamed to `periodi` kept `@param periods`, and nothing said so). An `@param`
 * naming no parameter is a warning, with a fix renaming it to the parameter at its
 * position (when that one is not documented already) and one removing it; parameters left
 * out while others are documented are a hint. `@param [p]` and `@param p` both document
 * `[p]`.
 */
function docParamProblems(file: SourceFile, def: NameDef, body: Expr, label: string): Problem[] {
  let e: Expr = body;
  while (e.kind === "paren") e = e.expr;
  if (e.kind !== "lambda") return [];
  const params = e.params.map((p) => ({ name: stripPrefix(p.name.text).base, optional: p.optional }));
  const text = file.text;
  const entry = def.entry;
  const open = text.lastIndexOf("/**", entry.offset);
  if (open < entry.start) return [];
  const close = text.indexOf("*/", open + 3);
  if (close < 0 || close > entry.offset) return [];
  const tags = docParamSpans(text, open + 3, close);
  if (tags.length === 0) return [];
  const path = file.path;
  const known = new Set(params.map((p) => lower(p.name)));
  const documented = new Set(tags.map((t) => lower(t.name)).filter((n) => known.has(n)));
  const shown = params.map((p) => (p.optional ? `[${p.name}]` : p.name)).join(", ");
  const out: Problem[] = [];
  const renamedTo = new Set<string>();
  tags.forEach((t, k) => {
    if (known.has(lower(t.name))) return;
    const fixes: Fix[] = [];
    const at = params[k];
    if (at && !documented.has(lower(at.name)) && !renamedTo.has(lower(at.name))) {
      renamedTo.add(lower(at.name));
      fixes.push({ title: `Rename @param ${t.name} to ${at.name}`, start: t.nameStart, end: t.nameEnd, text: at.name });
    }
    fixes.push({ title: `Remove @param ${t.name}`, start: t.removeStart, end: t.removeEnd, text: "" });
    out.push({
      path,
      start: t.nameStart,
      end: t.nameEnd,
      severity: "warning",
      code: "doc-param",
      key: def.key,
      message: `${label}: @param ${t.name} names no parameter of the LAMBDA (${params.length ? `its parameters: ${shown}` : "it takes none"}); the doc comment is the Name Manager's tooltip: rename or remove it`,
      fixes,
    });
  });
  const missing = params.filter((p) => !documented.has(lower(p.name)) && !renamedTo.has(lower(p.name)));
  if (missing.length) {
    const names = missing.map((p) => p.name);
    out.push({
      path,
      start: open,
      end: close + 2,
      severity: "hint",
      code: "doc-param-missing",
      key: def.key,
      message: `${label}: the doc comment documents some parameters but not ${names.join(", ")}: add @param ${names.length === 1 ? names[0] : "lines for them"}`,
    });
  }
  return out;
}

/** Whether a statement's formula is the one the lockfile has (the last pull or build): not an edit. */
function unchangedSinceLock(def: NameDef, stored: string, lock: Lockfile): boolean {
  const e = def.entry;
  let like: string | null | undefined;
  if (e.cell) {
    const cells = Object.values(lock.cells ?? {});
    const c =
      def.name !== ""
        ? cells.find((x) => x.name !== undefined && lower(x.name) === lower(def.key))
        : cells.find((x) => x.name === undefined && lower(x.sheet) === lower(e.cellSheet ?? "") && x.range === e.cell!.range);
    like = c?.formula;
  } else {
    const k = Object.keys(lock.names).find((x) => lower(x) === lower(def.key));
    like = k === undefined ? undefined : lock.names[k]!.definition;
  }
  if (like === undefined || like === null) return false;
  return definitionHashLike(like, stored) === like;
}


/** Removes the `@workbook` annotation of an entry: its line, or the word when the line holds more. */
function removeWorkbookFix(file: SourceFile, def: NameDef, title: string): Fix | undefined {
  const text = file.text;
  const k = text.indexOf("@workbook", def.entry.start);
  if (k < 0 || k > def.entry.offset) return undefined;
  const ls = text.lastIndexOf("\n", k - 1) + 1;
  let le = text.indexOf("\n", k);
  le = le < 0 ? text.length : le;
  const rest = text.slice(ls, k) + text.slice(k + "@workbook".length, le);
  if (rest.trim() === "") return { title, start: ls, end: le < text.length ? le + 1 : le, text: "" };
  let e = k + "@workbook".length;
  while (text[e] === " " || text[e] === "\t") e++;
  return { title, start: k, end: e, text: "" };
}

/**
 * *Remove @workbook* on a workbook name fixed to the file's sheet's cells
 * (`amount = Mortgage!$B$12;`): the reference loses its own sheet too (`amount = $B$12;`),
 * as a pull writes a local name, so the pull after the build rewrites nothing. Only for a
 * formula written in one piece (no comment inside); otherwise the plain fix.
 */
function makeLocalFix(file: SourceFile, def: NameDef, sheet: string, title: string): Fix | undefined {
  const base = removeWorkbookFix(file, def, title);
  const e = def.entry;
  if (!base || e.formulaMap.length !== 2) return base;
  const fs = formulaToSource(e, 0);
  const fe = fs + e.formula.length;
  if (fs < base.end || file.text.slice(fs, fe) !== e.formula) return base;
  let local: string;
  try {
    local = decompile(compile(e.formula), { homeSheet: sheet });
  } catch {
    return base;
  }
  if (local === e.formula) return base;
  return { title, start: base.start, end: fe, text: base.text + file.text.slice(base.end, fs) + local };
}

/**
 * What a named cell statement's `#` says against the workbook's spill (§6.3), one rule for
 * the editor, `xln check` and the build (§14 issue 8):
 * - no `#` on a formula that spilled beyond its cell when the workbook was saved: the name
 *   covers the first cell only. A warning, with the fix that adds the `#`. Not in a project
 *   pulled before the `#` was written (lockfile format 3 or older) for a name already on
 *   its spill: the build reads the missing `#` there as `#`.
 * - `#` on a cell left empty (`Name @C6# = ;`) that the name is not on yet: not written
 *   until the cell gets a formula (an empty cell's spill is not measured).
 */
function spillProblem(def: NameDef, ctx: CheckContext): Problem | undefined {
  const e = def.entry;
  const c = e.cell!;
  const sheet = e.cellSheet;
  if (sheet === undefined || c.range.includes(":")) return undefined;
  const at = parseCellAddress(c.range);
  if (!at) return undefined;
  const ref = `'${sheet.split("'").join("''")}'!$${columnName(at.col)}$${at.row}`;
  const lockKey = ctx.lock ? Object.keys(ctx.lock.names).find((k) => lower(k) === lower(def.key)) : undefined;
  const lockDef = lockKey !== undefined ? ctx.lock!.names[lockKey]!.definition : undefined;
  const lockOnSpill = lockDef !== undefined && definitionHashLike(lockDef, `_xlfn.ANCHORARRAY(${ref})`) === lockDef;
  const base = { path: def.file.path, start: c.start, end: c.end, severity: "warning" as const, key: def.key };
  if (c.spill) {
    if (e.formula !== "" || lockOnSpill) return undefined;
    return { ...base, code: "spill-empty", message: `${def.name} @${c.range}#: '#' on a cell left empty is not written yet; ${def.key} goes on ${c.range}# when the cell gets a formula` };
  }
  if (e.formula === "") return undefined;
  const sheetKey = Object.keys(ctx.spills ?? {}).find((s) => lower(s) === lower(sheet));
  const sp = sheetKey !== undefined ? ctx.spills![sheetKey]!.find((x) => x.anchor === c.range) : undefined;
  if (!sp || !sp.extent.includes(":")) return undefined;
  if (ctx.lock && !explicitSpill(ctx.lock) && lockOnSpill) return undefined;
  return {
    ...base,
    code: "spill-uncovered",
    message: `${def.name} @${c.range}: the formula spills over ${sp.extent}, but ${def.name} covers only ${c.range} (write @${c.range}# to name the spill)`,
    fix: { title: `Name the whole spill: @${c.range}#`, start: c.end, end: c.end, text: "#" },
  };
}

/** A cell statement's address: its form, its sheet, and what the last pull saw. */
function addressProblems(def: NameDef, file: SourceFile, ctx: CheckContext, lock: LockIndex | undefined, sheetSet: Set<string> | undefined): Problem[] {
  const out: Problem[] = [];
  const pm = file.parsed;
  const c = def.entry.cell!;
  const path = file.path;
  const at = { path, start: c.start, end: c.end };
  const bare = c.spill ? `@${c.range}#` : `@${c.range}`;
  if (pm.form === "sheet") {
    if (c.sheet !== undefined) {
      const sheetEnd = c.rangeStart;
      if (lower(c.sheet) === lower(pm.sheet!)) {
        out.push({ ...at, severity: "warning", code: "address", message: `the sheet is the file's (${pm.sheet}): write ${bare}`, fix: { title: `Write ${bare}`, start: c.start + 1, end: sheetEnd, text: "" } });
      } else {
        out.push({ ...at, severity: "error", code: "address", message: `${c.sheet}!${c.range} is a cell of ${c.sheet}, and this file holds ${pm.sheet}'s statements: write it in names/sheets/ for ${c.sheet} (as ${bare})` });
        return out;
      }
    }
  } else {
    const sheet = def.entry.cellSheet;
    if (sheet === undefined) {
      out.push({ ...at, severity: "error", code: "address", message: `a cell statement outside a sheet file names its sheet: @Sheet!${c.range} (or write it in names/sheets/<Sheet>.xln as @${c.range})` });
      return out;
    }
    if (c.sheet !== undefined && def.scope !== undefined && lower(c.sheet) !== lower(def.scope)) {
      out.push({ ...at, severity: "error", code: "address", message: `a name local to ${def.scope} has its cell on ${def.scope}: the address cannot name sheet ${c.sheet}` });
      return out;
    }
  }
  const sheet = def.entry.cellSheet!;
  if (sheetSet && !sheetSet.has(lower(sheet))) {
    if (pm.form !== "sheet" || c.sheet !== undefined) out.push({ ...at, severity: "error", code: "unknown-sheet", message: `the workbook has no sheet '${sheet}'` });
    return out;
  }
  if (def.name !== "" && c.range.includes(":")) {
    out.push({ ...at, severity: "error", code: "address", message: `${def.name}: a named cell statement is one cell (@C6, or @C6# for its spill); a range is an unnamed block (@B4:G4 = …;)` });
    return out;
  }
  if (!lock || !lock.hasCells) return out;
  const sameSheet = (s: string) => lower(s) === lower(sheet);
  if (def.name !== "") {
    let lc = lock.namedCells.get(lower(def.key));
    const r = renamedFrom(def.entry.annotations);
    if (!lc && r) lc = lock.namedCells.get(lower(defKey(r.name, r.scope === undefined ? def.scope : (r.scope ?? undefined))));
    if (!lc) {
      // A scope changed in the source: the same name in another scope.
      const others = (lock.byName.get(lower(def.name)) ?? []).map((k) => lock.namedCells.get(lower(k))).filter((x): x is LockCell => x !== undefined);
      if (others.length === 1) lc = others[0];
    }
    if (!lc) {
      const known = lock.names.has(lower(def.key));
      out.push({ ...at, severity: "error", code: "address", message: known ? `${def.key} was not a cell statement at the last pull: cells are named in Excel (then pull)` : `${def.key} is not in the last pull: a cell is named in Excel (Name Manager or the Name Box), then pulled` });
      return out;
    }
    if (sameSheet(lc.sheet) && lc.range === c.range) return out;
    const was = sameSheet(lc.sheet) ? lc.range : `${lc.sheet}!${lc.range}`;
    const p: Problem = { ...at, severity: "error", code: "address", message: `${def.name}: the address is set in Excel and read-only; the last pull had @${was}${sameSheet(lc.sheet) ? "" : " (move the name in Excel, then pull)"}` };
    if (sameSheet(lc.sheet)) p.fix = { title: `Restore the address @${lc.range}`, start: c.rangeStart, end: c.spill ? c.end - 1 : c.end, text: lc.range };
    out.push(p);
    return out;
  }
  if (!lock.unnamed.has(lower(`${sheet}!${c.range}`)) && ![...lock.namedCells.values()].some((x) => sameSheet(x.sheet) && x.range === c.range)) {
    out.push({ ...at, severity: "error", code: "address", message: `no cell statement at ${sheet}!${c.range} in the last pull: addresses are set in Excel (write the formula in Excel, then pull)` });
  }
  return out;
}

/**
 * A name in a file a pull would not write it to (§14 issues 2 and 6): the build takes the
 * file as it is, the next pull moves the name, and the build after that changes its
 * provenance tag. The structure is fixed (§2.1): a name with a module prefix goes to its
 * module's file, a workbook name without one to `_unmanaged.xln`, a sheet's local name
 * without one to that sheet's file. A warning at the name, with a fix moving it where the
 * pull would (a workbook-scoped name only), and, in a module file, one adding the prefix.
 */
function placementProblem(model: SourceModel, def: NameDef, file: SourceFile, modules: ReadonlyMap<string, string | undefined>, lock: LockIndex | undefined, cellsOf: string | undefined): Problem | undefined {
  const pm = file.parsed;
  const e = def.entry;
  if (e.cell || pm.form === "sheet-blocks") return undefined;
  const path = file.path;
  const inSheets = path.startsWith(`names/${SHEETS_DIR}/`);
  if (inSheets && pm.form !== "sheet") return undefined; // the `sheet-file` warning says it
  const fileModule = inSheets ? undefined : moduleOfPath(path);
  const unmanaged = !inSheets && fileModule === undefined;
  if (unmanaged && path !== `names/${UNMANAGED}.xln`) return undefined;
  const module = modules.get(def.name);
  // A workbook name on one sheet's cells goes to that sheet's file (2026-10-07).
  const home = def.scope ?? cellsOf;
  const target = module !== undefined ? `names/${moduleFileName(module)}` : home === undefined ? `names/${UNMANAGED}.xln` : `names/${SHEETS_DIR}/${sheetFileName(home)}`;
  const here =
    module !== undefined
      ? fileModule !== undefined && lower(moduleFileName(module)) === lower(moduleFileName(fileModule))
      : home === undefined
        ? unmanaged
        : inSheets && pm.sheet !== undefined && lower(pm.sheet) === lower(home);
  if (here) return undefined;
  // A sheet's local name in an old `@scope` block of `_unmanaged.xln`: the `@scope` line has the warning.
  const perName = e.annotations.some((a) => a.name === "sheet");
  if (unmanaged && def.scope !== undefined && !perName) return undefined;
  const label = def.scope !== undefined && !inSheets ? `${def.name} (local to ${def.scope})` : def.name;
  let message: string;
  if (module === undefined && def.scope === undefined && cellsOf !== undefined) {
    message = `${def.name} is on cells of ${cellsOf}: a pull puts it in ${target} (with @workbook above it); move it there`;
  } else if (fileModule !== undefined && module === undefined) {
    message = `${label} has no ${fileModule}. prefix: a pull will put it in ${target.slice("names/".length)}; rename it ${fileModule}.${def.name} or move it there`;
  } else if (fileModule !== undefined) {
    message = `${label} has the prefix of module ${module}, not ${fileModule}: a pull will put it in ${target}; move it there`;
  } else if (unmanaged && module === undefined) {
    message = `${label} has no module prefix: a pull puts a name local to ${def.scope} that no module owns in ${target}; move it there`;
  } else if (inSheets && def.scope === undefined) {
    message = `${def.name}: a pull puts workbook names that are not on one sheet's cells in _unmanaged.xln (or their module): move it to ${target}, or remove @workbook to make it local to ${pm.sheet}`;
  } else {
    message = `${label} has the prefix of module ${module}: a pull will put it in ${target}${def.scope !== undefined ? ` (with @sheet(${def.scope}) above it)` : ""}; move it there`;
  }
  const p: Problem = { ...model.nameLoc(def), severity: "warning", code: "file-placement", key: def.key, message };
  const fixes: Fix[] = [];
  if (def.scope === undefined) {
    const move = moveFix(model, def, file, target);
    if (move) fixes.push(move);
  }
  if (fileModule !== undefined && module === undefined) {
    // The name the pull keeps here. A name the last pull had is renamed (`@renamed`), not created anew.
    const to = `${fileModule}.${def.name}`;
    const known = lock?.names.has(lower(def.key)) === true;
    const at = model.nameLoc(def);
    fixes.push(known ? { title: `Rename it ${to}`, start: e.start, end: at.end, text: `@renamed(${def.name})\n${file.text.slice(e.start, at.start)}${to}` } : { title: `Rename it ${to}`, start: at.start, end: at.end, text: to });
  }
  if (fixes.length) p.fixes = fixes;
  return p;
}

/**
 * Moves a workbook-scoped name's statement (doc comment and annotations with it) to the end
 * of `target`, a module file, `_unmanaged.xln` or a sheet file. `@workbook` says "workbook
 * scope" only in a sheet file: it stays behind when the target is not one, and comes along
 * (or is added) when it is. Undefined when the statement shares its lines with another, or
 * the target is not a file of the form it should be.
 */
function moveFix(model: SourceModel, def: NameDef, file: SourceFile, target: string): Fix | undefined {
  const text = file.text;
  const e = def.entry;
  // The statement's whole lines, and one blank line after it when one is left above.
  const ls = text.lastIndexOf("\n", e.start - 1) + 1;
  if (text.slice(ls, e.start).trim() !== "") return undefined;
  let le = text.indexOf("\n", e.end);
  if (le >= 0 && text.slice(e.end, le).trim() !== "") return undefined;
  le = le < 0 ? text.length : le + 1;
  let from = ls;
  if (text[le] === "\n" && (ls === 0 || text.slice(0, ls).endsWith("\n\n"))) le++;
  else if (le >= text.length && text.slice(0, ls).endsWith("\n\n")) from--;
  let stmt = text.slice(e.start, e.end);
  const toSheet = target.startsWith(`names/${SHEETS_DIR}/`);
  for (const an of [...e.annotations].reverse()) {
    if (an.name !== "workbook" || toSheet) continue;
    const a = an.offset - e.start;
    let b = an.end - e.start;
    while (stmt[b] === " " || stmt[b] === "\t") b++;
    if (stmt[b] === "\n") b++;
    stmt = stmt.slice(0, a) + stmt.slice(b);
  }
  // In a sheet file `@workbook` goes on the line above the name, after the doc comment.
  if (toSheet && !e.annotations.some((an) => an.name === "workbook")) {
    const k = Math.min(e.offset, ...e.annotations.map((an) => an.offset)) - e.start;
    stmt = stmt.slice(0, k) + "@workbook\n" + stmt.slice(k);
  }
  const t = model.files.get(target);
  if (!t) {
    const head = toSheet
      ? `// Sheet ${sheetOfPath(target) ?? ""}.\n`
      : target === `names/${UNMANAGED}.xln`
        ? "// Workbook-scoped names that no module owns.\n"
        : `// module: ${target.slice("names/".length, -".xln".length)}\n`;
    return { title: `Move it to ${target}`, start: from, end: le, text: "", elsewhere: [{ path: target, start: 0, end: 0, text: `${head}\n${stmt}\n`, create: true }] };
  }
  if (t.parsed.form !== (toSheet ? "sheet" : "module")) return undefined;
  const last = t.parsed.scopes[t.parsed.scopes.length - 1];
  const reopen = !toSheet && last !== undefined && last.scope !== undefined ? "@workbook\n\n" : "";
  const end = t.text.length;
  const sep = t.text === "" || t.text.endsWith("\n\n") ? "" : t.text.endsWith("\n") ? "\n" : "\n\n";
  return { title: `Move it to ${target}`, start: from, end: le, text: "", elsewhere: [{ path: target, start: end, end, text: `${sep}${reopen}${stmt}\n` }] };
}

/** A definition that is one cell (or its spill), not an area. */
function isOneCell(formula: string): boolean {
  const t = definitionTarget(formula);
  return t !== undefined && t.r1 === t.r2 && t.c1 === t.c2;
}

/** A name: legal, and its scope against the last pull. */
function nameProblems(model: SourceModel, def: NameDef, file: SourceFile, lock: LockIndex | undefined, cellsOf?: string): Problem[] {
  const out: Problem[] = [];
  const pm = file.parsed;
  const e = def.entry;
  const at = model.nameLoc(def);
  const r = renamedFrom(e.annotations);
  const inLock = lock?.names.has(lower(def.key)) === true;
  // A name the last pull did not have: Excel must accept it.
  if (lock && !inLock && !r && (lock.byName.get(lower(def.name)) ?? []).length === 0) {
    const bad = invalidName(def.name, def.entry.formula);
    if (bad) out.push({ ...at, severity: "error", code: "invalid-name", key: def.key, message: `${def.name}: ${bad}` });
  }
  // A name spelled like a macro function, new since the last pull (a renamed one too).
  if (!lock || (lock.byName.get(lower(def.name)) ?? []).length === 0) {
    const w = xlmNameWarning(def.name, e.formula);
    if (w) out.push({ ...at, severity: "warning", code: "xlm-name", key: def.key, message: w });
  }
  // A workbook name on a sheet's cell (Create from Selection makes them), read only on its
  // own sheet: it could be local. A hint, not a warning: the author may want it workbook-
  // wide. Read from another sheet, or by a name of another scope, it needs workbook scope
  // (or the other sheet would need `Sheet!Name`): no finding (feedback 2026-10-07). The same
  // for a workbook name whose definition is fixed to the sheet's cells (a cell inside a
  // spill, a range), which a pull writes in the sheet's file too.
  const home = pm.sheet === undefined ? undefined : lower(pm.sheet);
  const onCells = e.cell !== undefined || (cellsOf !== undefined && lower(cellsOf) === home);
  if (onCells && def.scope === undefined && pm.form === "sheet" && model.readers(def.key).every((r) => r === def || (model.homeOf(r) !== undefined && lower(model.homeOf(r)!) === home))) {
    const p: Problem = {
      ...at,
      severity: "hint",
      code: "workbook-on-cell",
      message: `workbook name on ${e.cell !== undefined || isOneCell(e.formula) ? "a cell" : "cells"} of ${pm.sheet}, read only on ${pm.sheet}: remove @workbook to make it local to ${pm.sheet} (the build then moves it)`,
    };
    const title = `Remove @workbook: make ${def.name} local to ${pm.sheet}`;
    const fix = e.cell === undefined ? makeLocalFix(file, def, pm.sheet!, title) : removeWorkbookFix(file, def, title);
    if (fix) p.fix = fix;
    out.push(p);
  }
  // A scope that differs from the last pull's: the same name, in one other scope, gone from the source.
  if (lock && !inLock && !r) {
    const others = (lock.byName.get(lower(def.name)) ?? []).filter((k) => model.lookup(k) === undefined);
    if (others.length === 1) {
      const was = splitNameKey(others[0]!).scope;
      const from = was === undefined ? "workbook scope" : `sheet ${was}`;
      const to = def.scope === undefined ? "workbook scope" : `sheet ${def.scope}`;
      out.push({ ...at, severity: "info", code: "rescope", message: `${def.name}: the build moves it from ${from} to ${to} (a scope change, rescope-name)` });
    }
  }
  return out;
}

/**
 * The checks as you type (M3c) on one definition's formula: names nothing defines,
 * another sheet's local name read without its sheet, functions the catalogue does not
 * know, argument counts, a name spelled in another case. Positions are on the token.
 */
function liveChecks(
  model: SourceModel,
  def: NameDef,
  a: ReturnType<SourceModel["analysis"]>,
  home: string | undefined,
  ctx: CheckContext,
  scope: ReturnType<SourceModel["scopeIndex"]>,
  tables: ReadonlySet<string>,
  sheets: ReadonlySet<string>,
  out: Problem[],
  broken: Map<Problem, string> = new Map(),
): void {
  if (!a.body) return;
  const lockNames = ctx.lock ? lockIndex(ctx.lock).names : undefined;
  /** The name a use read at the last pull (Excel's resolution: the home sheet's, else the workbook's). */
  const wasRead = (id: string, sheet: string | undefined): string | undefined => {
    if (!lockNames) return undefined;
    const local = sheet ?? home;
    return (local !== undefined ? lockNames.get(lower(`${local}!${id}`)) : undefined) ?? lockNames.get(lower(id));
  };
  const unresolved = (p: Problem, id: string, sheet: string | undefined): void => {
    out.push(p);
    const was = wasRead(id, sheet);
    if (was !== undefined) broken.set(p, was);
  };
  const binders = a.occurrences.filter((o) => o.kind === "binder").map((o) => o.id);
  for (const o of a.occurrences) {
    if (o.kind !== "use") continue;
    const at: Loc = model.loc(def, o.span);
    if (o.key !== undefined) {
      const target = model.lookup(o.key);
      if (target && o.id !== target.name && lower(o.id) === lower(target.name)) {
        out.push({
          ...at,
          severity: "hint",
          code: "spelling",
          message: `${o.id} is defined as ${target.name}: Excel reads it case-insensitively and shows it as ${target.name}`,
          fixes: [{ title: `Match the name's spelling: ${target.name}`, ...span(at), text: target.name }],
        });
      }
      continue;
    }
    if (!o.call && (lookupFunction(o.id) !== undefined || tables.has(lower(o.id)))) continue; // `BYROW(x, SUM)`, `SUM(tblSales)`
    if (o.sheet !== undefined) {
      // `Other.xlsx!Rate`: a name in another workbook, not a sheet (the compiler checks the link).
      if (!sheets.has(lower(o.sheet)) && (looksLikeWorkbookFile(o.sheet) || (ctx.links ?? []).some((l) => lower(l.book) === lower(o.sheet!)))) continue;
      const own = scope.locals.get(lower(o.sheet)) ?? [];
      const near = [...own, ...scope.workbook].filter((d) => oneEditAway(d.name, o.id)).map((d) => d.name);
      const noSheet = !sheets.has(lower(o.sheet));
      unresolved({
        ...at,
        severity: "error",
        code: noSheet ? "unknown-sheet" : "C4.unknown-name",
        message: noSheet ? `${o.sheet}!${o.id}: the workbook has no sheet ${o.sheet}` : `${o.sheet}!${o.id}: no name ${o.id} on ${o.sheet} or in the workbook: #NAME? in Excel`,
        ...(noSheet ? {} : { fixes: didYouMean(at, near) }),
      }, o.id, o.sheet);
      continue;
    }
    const elsewhere = model.sheetsWith(o.id, home, ctx.sheets);
    if (elsewhere.length > 0) {
      const name = scope.locals.get(lower(elsewhere[0]!))!.find((d) => lower(d.name) === lower(o.id))!.name;
      unresolved({
        ...at,
        severity: "error",
        code: "C5.other-sheet",
        message: `${o.id} is local to ${elsewhere.join(", ")}, not ${home === undefined ? "a workbook name" : `to ${home}`}: unqualified it is #NAME? in Excel; write ${sheetPrefix(elsewhere[0]!)}${name}`,
        fixes: elsewhere.map((s) => ({ title: `Qualify: ${sheetPrefix(s)}${name}`, ...span(at), text: sheetPrefix(s) + name })),
      }, o.id, undefined);
      continue;
    }
    const inScopeDefs = [...(home !== undefined ? (scope.locals.get(lower(home)) ?? []) : []), ...scope.workbook];
    const inScope = inScopeDefs.map((d) => d.name);
    if (o.call) {
      const lambdas = inScopeDefs.filter((d) => model.analysis(d).classification.kind === "lambda").map((d) => d.name);
      const builtins = [...catalogue().values()].filter((f) => !f.internal).map((f) => f.name);
      const near = [...lambdas, ...builtins].filter((n) => oneEditAway(n, o.id));
      unresolved({
        ...at,
        severity: "error",
        code: "unknown-function",
        message: `${o.id}(…): no built-in function and no LAMBDA of that name; Excel would store it as _xludf.${o.id} (#NAME?)${near.length ? `: did you mean ${near[0]}?` : ""}`,
        fixes: didYouMean(at, near),
      }, o.id, undefined);
      continue;
    }
    const near = [...binders, ...inScope].filter((n) => oneEditAway(n, o.id));
    unresolved({
      ...at,
      severity: "error",
      code: "C4.unknown-name",
      message: `${o.id} is not a name in scope, a LET/LAMBDA variable or a function: #NAME? in Excel${near.length ? ` (did you mean ${near[0]}?)` : ""}`,
      fixes: didYouMean(at, near),
    }, o.id, undefined);
  }
  /** A cell called (`C2(x, y)`, or a name on C2 called): the LAMBDA it holds, by its cell statement. */
  const cellCall = (at: Loc, fn: string, sheet: string, cell: string, n: number): void => {
    const callee = calledCell(model, ctx, sheet, cell);
    if (callee === undefined) return;
    if (callee.kind === "lambda" && callee.arity) {
      const { required, optional } = callee.arity;
      if (n < required || n > required + optional) {
        out.push({ ...at, severity: "error", code: "C6.lambda-arity", message: `${fn}(${(callee.params ?? []).join(", ")}) takes ${wantText(required, required + optional)}; it is given ${n}` });
      }
    } else if (callee.kind === "not-lambda") {
      const holds = callee.holds === "value" ? "holds a value or nothing" : "has a formula that gives a value";
      out.push({ ...at, severity: "warning", code: "C6.not-a-lambda", message: `${fn} is called with ${plural(n, "argument")}, but ${sheetPrefix(sheet)}${cell} ${holds}, not a LAMBDA: Excel gives #VALUE! or #CALC!` });
    }
  };
  for (const c of callUses(a.body)) {
    const n = c.args.length;
    const at = model.loc(def, c.kind === "builtin" ? { start: c.span.end - c.id.length, end: c.span.end } : c.span);
    if (c.kind === "builtin") {
      if (n < c.info.minArgs || n > c.info.maxArgs) {
        out.push({ ...at, severity: "warning", code: "C6.builtin-arity", message: `${c.info.name} takes ${wantText(c.info.minArgs, c.info.maxArgs)}; it is given ${n}` });
      }
      continue;
    }
    if (c.kind === "cell") {
      const sheet = c.sheet ?? home;
      const cell = parseCellAddress(c.ref.address);
      if (sheet !== undefined && cell) cellCall(at, def.entry.formula.slice(c.span.start, c.span.end), sheet, formatCell(cell), n);
      continue;
    }
    const key = model.resolveName(c.id, c.sheet, home);
    const target = key === undefined ? undefined : model.lookup(key);
    if (!target) continue;
    const cls = model.analysis(target).classification;
    if (cls.kind === "lambda" && cls.arity) {
      const { required, optional } = cls.arity;
      if (n < required || n > required + optional) {
        out.push({ ...at, severity: "error", code: "C6.lambda-arity", message: `${c.id}(${(cls.params ?? []).join(", ")}) takes ${wantText(required, required + optional)}; it is given ${n}` });
      }
    } else if (cls.kind === "range" || cls.kind === "spill") {
      // A name on a cell called: the LAMBDA the cell holds (`Fn @C2 = LAMBDA(…)`, `Fn(x, y)`).
      const t = nameCell(target);
      if (t) cellCall(at, c.id, t.sheet, t.cell, n);
      else out.push({ ...at, severity: "warning", code: "C6.not-a-lambda", message: `${c.id} is called with ${plural(n, "argument")}, but it is a ${cls.kind === "spill" ? "spill range" : "range of several cells"}, not a LAMBDA: Excel gives #VALUE! or #CALC!` });
    } else if (cls.kind === "constant" || cls.kind === "table") {
      out.push({ ...at, severity: "error", code: "C6.not-a-function", message: `${c.id} is a ${KIND_LABEL[cls.kind]}, not a LAMBDA, but is called with ${plural(n, "argument")}` });
    }
  }
}

/** The one cell a name of kind range stands for: its cell statement's, or its definition's (`=S!$C$2`). */
function nameCell(def: NameDef): { sheet: string; cell: string } | undefined {
  const c = def.entry.cell;
  if (c) {
    const cell = parseCellAddress(c.range);
    return !c.spill && cell && def.entry.cellSheet !== undefined ? { sheet: def.entry.cellSheet, cell: formatCell(cell) } : undefined;
  }
  const t = definitionTarget(def.entry.formula);
  const sheet = t?.sheet ?? def.scope;
  if (!t || t.spill || t.r1 !== t.r2 || t.c1 !== t.c2 || sheet === undefined) return undefined;
  return { sheet, cell: formatCell({ row: t.r1, col: t.c1 }) };
}

/**
 * What a called cell holds, from the cell statements of the source: the statement covering
 * the cell (a slot: a value), or, when the last pull recorded the sheet's cells and none
 * covers it, a value or nothing. Undefined when the source cannot tell (no workbook).
 */
function calledCell(model: SourceModel, ctx: CheckContext, sheet: string, cell: string): CellCallee | undefined {
  const p = parseCellAddress(cell)!;
  const covers = (d: NameDef): boolean => {
    const c = d.entry.cell;
    if (!c || d.entry.cellSheet === undefined || lower(d.entry.cellSheet) !== lower(sheet)) return false;
    const [a, b = a] = c.range.split(":");
    const x = parseCellAddress(a!);
    const y = parseCellAddress(b!);
    return x !== undefined && y !== undefined && Math.min(x.row, y.row) <= p.row && p.row <= Math.max(x.row, y.row) && Math.min(x.col, y.col) <= p.col && p.col <= Math.max(x.col, y.col);
  };
  const d = model.defs.find(covers) ?? model.cellDefs.find(covers);
  if (d) return cellCallee(d.entry.formula);
  const known = ctx.lock?.cells !== undefined && (ctx.sheets ?? []).some((s) => lower(s) === lower(sheet));
  return known ? { kind: "not-lambda", holds: "value" } : undefined;
}
