// Writing help in `.xln` files (M3c), without the vscode API: completion, signature help,
// and the checks as you type (the core checker's, M3d). All read the project model's live sources, so they
// follow unsaved edits in every file of the project.
//
// Names are offered and checked the way Excel resolves them (`NameResolver`): in a
// formula of sheet S, S's local names bare, other sheets' local names as `Sheet!Name`,
// workbook names bare; LET and LAMBDA variables shadow them; a call spelled like a
// built-in calls the built-in (probe T12).

import {
  activeCall,
  catalogue,
  checkFile,
  formulaCursor,
  formulaToSource,
  KIND_LABEL,
  sheetPrefix,
  oneEditAway,
  lookupFunction,
  moduleFileOf,
  paramIndex,
  paramDoc,
  parseDocComment,
  significant,
  sourceToFormula,
  stripPrefix,
  tokenize,
  tryParse,
  type FormulaCursor,
  type FunctionInfo,
  type Library,
  type NameKind,
  type ScopeIndex,
} from "@xln/core";
import { libraryCompletions, libraryItemDoc, libraryModules, librarySignature } from "./library.js";
import { symbolDetail, type NameDef, type Problem, type Project } from "./project.js";

const lower = (s: string) => s.toLowerCase();

// ---- where the cursor is --------------------------------------------------------------

/** A cursor inside the formula of a definition or cell statement. */
export interface FormulaSite {
  def: NameDef;
  /** Index into `def.entry.formula`. */
  index: number;
  /** The sheet whose local names the formula reads bare. */
  home: string | undefined;
  cursor: FormulaCursor;
}

function isBlank(text: string, from: number, to: number): boolean {
  for (let k = from; k < to; k++) if (!" \t\r\n".includes(text[k]!)) return false;
  return true;
}

/**
 * The formula index at a file offset, blanks around the formula included (`X = |;`, an
 * empty slot). Undefined on the left-hand side, in a cell address, in a comment, after `;`.
 */
export function formulaIndex(def: NameDef, offset: number): number | undefined {
  const e = def.entry;
  const i = sourceToFormula(e, offset);
  if (i !== undefined) return i;
  const text = def.file.text;
  const s = formulaToSource(e, 0);
  if (offset < s && isBlank(text, offset, s)) {
    let k = offset - 1;
    while (k >= 0 && " \t\r\n".includes(text[k]!)) k--;
    const lhsEnd = e.cell ? e.cell.end : e.offset + e.name.length;
    return text[k] === "=" && k >= lhsEnd ? 0 : undefined;
  }
  const end = formulaToSource(e, e.formula.length);
  const close = text[e.end - 1] === ";" ? e.end - 1 : e.end;
  if (offset > end && offset <= close && isBlank(text, end, offset)) return e.formula.length;
  return undefined;
}

export function formulaSite(project: Project, path: string, offset: number): FormulaSite | undefined {
  const def = project.defAt(path, offset);
  if (!def) return undefined;
  const index = formulaIndex(def, offset);
  if (index === undefined) return undefined;
  return { def, index, home: project.homeOf(def), cursor: formulaCursor(def.entry.formula, index) };
}

// ---- names in scope -------------------------------------------------------------------

/** `Sheet!` as a formula writes it (core `sheetPrefix`). */
export { sheetPrefix, oneEditAway };

function scopeOf(project: Project): ScopeIndex {
  return project.scopeIndex();
}

// ---- completion -----------------------------------------------------------------------

export type CompletionKind = "variable" | "name" | "module" | "function";

export interface Completion {
  label: string;
  kind: CompletionKind;
  /** For a name: what it is. */
  nameKind?: NameKind;
  insertText: string;
  /** What the typed word is matched against (a qualified name matches on its name). */
  filterText: string;
  detail: string;
  /** Markdown. */
  documentation?: string;
  sortText: string;
  /** The file range the item replaces (the word being typed), as offsets. */
  start: number;
  end: number;
  /** Complete again after inserting it (a module prefix `FN.`). */
  retrigger?: boolean;
  /** M4: a library function the project lacks; accepting it adds `adds` (it and what it calls) to the project. */
  library?: { name: string; adds: string[] };
}

/** Sort ranks: variables in scope, then the home sheet's names, then other names, then functions. */
const RANK = { variable: "0", local: "1", other: "2", function: "3" } as const;

function versionText(since: string): string {
  return since === "2007" ? "Excel 2007 or earlier" : since === "365" ? "Excel 365" : `Excel ${since}`;
}

/** `SUM(number1, [number2], ...)`, or `SUM(1..255 args)` while the catalogue has no parameter names. */
export function functionSignature(info: FunctionInfo): string {
  if (info.params) return `${info.name}(${info.params.join(", ")})`;
  const n = info.minArgs === info.maxArgs ? `${info.minArgs}` : `${info.minArgs}..${info.maxArgs}`;
  return `${info.name}(${n} arg${info.maxArgs === 1 ? "" : "s"})`;
}

function nameDoc(d: NameDef): string {
  const parts: string[] = [];
  if (d.entry.doc) {
    const doc = parseDocComment(d.entry.doc);
    if (doc.summary) parts.push(doc.summary);
    if (doc.params.length) parts.push(doc.params.map((p) => `*@param* \`${p.name}\` — ${p.text}`).join("  \n"));
  }
  const f = d.entry.formula.length > 300 ? d.entry.formula.slice(0, 299) + "…" : d.entry.formula;
  if (f !== "") parts.push("```xln\n" + f + "\n```");
  return parts.join("\n\n");
}

function nameItem(project: Project, d: NameDef, label: string, rank: string, start: number, end: number, filterText = label): Completion {
  const c = project.analysis(d).classification;
  const where = d.scope === undefined ? "workbook" : `local to ${d.scope}`;
  const what = c.kind === "lambda" ? symbolDetail(c, d.entry.formula) : KIND_LABEL[c.kind];
  return {
    label,
    kind: "name",
    nameKind: c.kind,
    insertText: label,
    filterText,
    detail: `${what} · ${where}`,
    documentation: nameDoc(d),
    sortText: `${rank}_${lower(label)}`,
    start,
    end,
  };
}

function functionItem(info: FunctionInfo, start: number, end: number): Completion {
  const doc: string[] = [];
  if (info.params) doc.push(info.params.map((p) => `\`${p}\``).join(", "));
  doc.push(`${versionText(info.since)}${info.prefix ? ` · stored ${info.prefix}${info.name}` : ""}`);
  return {
    label: info.name,
    kind: "function",
    insertText: info.name,
    filterText: info.name,
    detail: `${functionSignature(info)} · ${versionText(info.since)}`,
    documentation: doc.join("\n\n"),
    sortText: `${RANK.function}_${lower(info.name)}`,
    start,
    end,
  };
}

function libraryItems(project: Project, library: Library, prefix: string, start: number, end: number): Completion[] {
  return libraryCompletions(project, library, prefix).map((c) => ({
    label: c.fn.name,
    kind: "function",
    insertText: c.fn.name,
    filterText: c.fn.name,
    detail: `from library · ${librarySignature(c.fn)}`,
    documentation: libraryItemDoc(c, moduleFileOf(c.fn.name)),
    // After the project's own names, before the built-in functions.
    sortText: `${RANK.other}~${lower(c.fn.name)}`,
    start,
    end,
    library: { name: c.fn.name, adds: c.adds },
  }));
}

/**
 * The completion list at a file offset; undefined where nothing completes. With a
 * `library`, its functions the project lacks are offered too (marked "from library").
 */
export function completions(project: Project, path: string, offset: number, library?: Library): Completion[] | undefined {
  const site = formulaSite(project, path, offset);
  if (!site || site.cursor.inert) return undefined;
  const { def, cursor, home } = site;
  const word = cursor.word;
  const start = formulaToSource(def.entry, word.start);
  const end = offset;
  const scope = scopeOf(project);
  const homeLocals = home !== undefined ? (scope.locals.get(lower(home)) ?? []) : [];
  const out: Completion[] = [];

  if (cursor.sheet !== undefined) {
    // `IS!Sa|`: that sheet's local names.
    for (const d of scope.locals.get(lower(cursor.sheet)) ?? []) out.push(nameItem(project, d, d.name, RANK.local, start, end));
    return out;
  }

  const dot = word.text.lastIndexOf(".");
  if (dot >= 0) {
    // `FN.|`: the members of the module; functions with a dot (`NORM.`) too.
    const prefix = lower(word.text.slice(0, dot + 1));
    const homeKeys = new Set(homeLocals.map((d) => lower(d.name)));
    for (const d of homeLocals) if (lower(d.name).startsWith(prefix)) out.push(nameItem(project, d, d.name, RANK.local, start, end));
    for (const d of scope.workbook) if (lower(d.name).startsWith(prefix) && !homeKeys.has(lower(d.name))) out.push(nameItem(project, d, d.name, RANK.other, start, end));
    for (const info of catalogue().values()) if (!info.internal && lower(info.name).startsWith(prefix)) out.push(functionItem(info, start, end));
    if (library) out.push(...libraryItems(project, library, word.text.slice(0, dot + 1), start, end));
    return out;
  }

  const seen = new Set<string>();
  for (const v of cursor.locals) {
    if (seen.has(lower(v.name))) continue;
    seen.add(lower(v.name));
    out.push({
      label: v.name,
      kind: "variable",
      insertText: v.name,
      filterText: v.name,
      detail: v.kind === "let" ? "LET variable" : "LAMBDA parameter",
      sortText: `${RANK.variable}_${lower(v.name)}`,
      start,
      end,
    });
  }
  const modules = new Map<string, number>();
  const bare = (d: NameDef, rank: string): void => {
    const n = lower(d.name);
    if (seen.has(n)) return;
    seen.add(n);
    const dotAt = d.name.indexOf(".");
    if (dotAt > 0) {
      const m = d.name.slice(0, dotAt + 1);
      modules.set(m, (modules.get(m) ?? 0) + 1);
      return;
    }
    out.push(nameItem(project, d, d.name, rank, start, end));
  };
  for (const d of homeLocals) bare(d, RANK.local);
  for (const d of scope.workbook) bare(d, RANK.other);
  for (const [m, n] of modules) {
    out.push({
      label: m,
      kind: "module",
      insertText: m,
      filterText: m,
      detail: `module · ${n} name${n === 1 ? "" : "s"}`,
      sortText: `${RANK.other}_${lower(m)}`,
      start,
      end,
      retrigger: true,
    });
  }
  // A module only the library has: its prefix leads to the library's functions.
  if (library) {
    const known = new Set([...modules.keys()].map(lower));
    for (const m of libraryModules(project, library)) {
      if (known.has(lower(m))) continue;
      out.push({ label: m, kind: "module", insertText: m, filterText: m, detail: "module · from library", sortText: `${RANK.other}_${lower(m)}`, start, end, retrigger: true });
    }
  }
  for (const [sheet, list] of scope.locals) {
    if (home !== undefined && sheet === lower(home)) continue;
    for (const d of list) {
      const q = sheetPrefix(d.scope!) + d.name;
      out.push(nameItem(project, d, q, RANK.other, start, end, d.name));
    }
  }
  for (const info of catalogue().values()) if (!info.internal) out.push(functionItem(info, start, end));
  return out;
}

// ---- signature help -------------------------------------------------------------------

export interface SignatureParam {
  label: string;
  documentation?: string;
}

export interface Signature {
  label: string;
  documentation?: string;
  params: SignatureParam[];
  /** The parameter the cursor is in (a repeating tail maps onto its last parameter). */
  active: number;
}

/**
 * Which parameter argument `index` fills: the catalogue's rule (`paramIndex`: `...`
 * repeats its group, so SUMIFS's sixth argument is `[criteria_range2]`); otherwise, past
 * the end of a repeating list, the last real one.
 */
function activeParam(params: string[], index: number): number {
  const k = paramIndex(params, index);
  if (k >= 0) return k;
  if (index < params.length) {
    const p = params[index]!;
    return p === "..." && index > 0 ? index - 1 : index;
  }
  const last = params[params.length - 1];
  if (last === undefined) return 0;
  if (last === "...") return Math.max(0, params.length - 2);
  if (last.endsWith("...")) return params.length - 1;
  return index;
}

function lambdaParams(text: string): string[] | undefined {
  const body = tryParse(text).formula?.body;
  let e = body;
  while (e && e.kind === "paren") e = e.expr;
  if (!e || e.kind !== "lambda") return undefined;
  return e.params.map((p) => {
    const n = stripPrefix(p.name.text).base;
    return p.optional ? `[${n}]` : n;
  });
}

/** Signature help at a file offset: the call the cursor is in, and the argument. */
export function signatureHelp(project: Project, path: string, offset: number): Signature | undefined {
  const site = formulaSite(project, path, offset);
  if (!site || site.cursor.inert) return undefined;
  const call = activeCall(site.cursor);
  if (!call) return undefined;
  const fn = call.frame.fn!;
  const sheet = call.frame.sheet;
  const { prefix, base } = stripPrefix(fn);
  const make = (name: string, params: string[], doc: string | undefined, paramText: (p: string) => string | undefined): Signature => ({
    label: `${name}(${params.join(", ")})`,
    ...(doc ? { documentation: doc } : {}),
    params: params.map((p) => {
      const d = paramText(p);
      return d ? { label: p, documentation: d } : { label: p };
    }),
    active: activeParam(params, call.index),
  });

  if (sheet === undefined) {
    // A LET variable holding a LAMBDA: `LET(sq, LAMBDA(x, x*x), sq(|`.
    const v = site.cursor.locals.find((l) => lower(l.name) === lower(base));
    if (v) {
      const params = v.value && lambdaParams(site.def.entry.formula.slice(v.value.start, v.value.end));
      return params ? make(v.name, params, undefined, () => undefined) : undefined;
    }
    const info = prefix === "" || prefix.startsWith("_xlfn.") ? lookupFunction(base) : undefined;
    if (info) {
      if (!info.params) return { label: functionSignature(info), documentation: versionText(info.since), params: [], active: 0 };
      return make(info.name, info.params, versionText(info.since), () => undefined);
    }
  }
  if (prefix !== "") return undefined;
  const key = project.resolveName(fn, sheet, site.home);
  const d = key === undefined ? undefined : project.lookup(key);
  if (!d) return undefined;
  const c = project.analysis(d).classification;
  if (c.kind !== "lambda") return undefined;
  const doc = parseDocComment(d.entry.doc ?? "");
  return make(fn, c.params ?? [], doc.summary || undefined, (p) => paramDoc(doc, p));
}

// ---- checks as you type ---------------------------------------------------------------

/** The codes of the checks as you type (M3c), now part of the core checker (M3d). */
const LIVE = new Set(["C4.unknown-name", "C5.other-sheet", "C6.lambda-arity", "C6.builtin-arity", "C6.not-a-function", "C6.not-a-lambda", "unknown-function", "unknown-sheet", "spelling", "comment-length"]);

/**
 * The checks as you type of one file (M3c) on the source text: names nothing defines,
 * another sheet's local name read without its sheet, unknown functions and sheets,
 * argument counts, a name spelled in another case, doc comments over the Name Manager's
 * 255 characters. They are the core checker's (`checkFile`), which `Project.problems`
 * publishes whole; this is the part of it the audit may also report.
 */
export function liveProblems(project: Project, path: string): Problem[] {
  return checkFile(project, path, project.checkContext()).filter((p) => p.code !== undefined && LIVE.has(p.code));
}

// ---- xln: New module ------------------------------------------------------------------

function isPrefixChar(c: string, first: boolean): boolean {
  return c === "_" || (c >= "A" && c <= "Z") || (c >= "a" && c <= "z") || (!first && c >= "0" && c <= "9");
}

/** Why `prefix` cannot name a new module of the project, or undefined when it can. */
export function modulePrefixProblem(project: Project, prefix: string): string | undefined {
  if (prefix === "") return "type a prefix, e.g. FIN";
  if (![...prefix].every((c, k) => isPrefixChar(c, k === 0))) return "a prefix is letters, digits and _ (no dot), starting with a letter or _";
  // `AB1.X`, `TRUE.X`: a formula would read the start of such a name as something else.
  const t = significant(tokenize(prefix));
  if (t.length !== 1 || t[0]!.kind !== "name") return `${prefix} reads as a cell or a value in a formula; pick another prefix`;
  if (lower(prefix) === "_unmanaged") return "_unmanaged.xln holds the names no module owns";
  const file = lower(`names/${prefix}.xln`);
  if ([...project.files.keys()].some((p) => lower(p) === file)) return `names/${prefix}.xln exists already`;
  const taken = project.defs.find((d) => lower(d.name).startsWith(lower(prefix) + "."));
  if (taken) return `${taken.name} (${taken.file.path}) already uses the prefix ${prefix}.`;
  return undefined;
}

/** The text of a new module file: a header and a sample LAMBDA with its doc comment. */
export function newModuleText(prefix: string): string {
  return [
    `// module: ${prefix}. Workbook names that start with ${prefix}. (LAMBDAs, inputs, ...).`,
    `// Nothing here is in the workbook until xln build writes it; the build tags each name's`,
    `// Name Manager comment with its module ([xln ${prefix} #hash]). A line \`// @version 1.0\``,
    `// here puts a version in the tags.`,
    ``,
    `/**`,
    ` * Grows a value by a rate, compounded over a number of periods.`,
    ` * @param value the starting value`,
    ` * @param rate the growth rate per period (0.05 for 5%)`,
    ` * @param [periods] how many periods; 1 when omitted`,
    ` */`,
    `${prefix}.GROW = LAMBDA(value, rate, [periods],`,
    `    value * (1 + rate) ^ IF(ISOMITTED(periods), 1, periods)`,
    `);`,
    ``,
  ].join("\n");
}
