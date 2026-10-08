// What moves between the library and a project, as text edits the caller shows and the
// user confirms (M4 principle: nothing moves without an explicit action, and nothing is
// built by these: the user builds).
//
// - Insert: a library function (and the library functions it calls that the project
//   lacks) goes into the module file of its prefix, `names/FN.xln`, created if missing,
//   in name order among the file's workbook names.
// - Update ("Take the library's version"): an entry's definition and doc comment are
//   replaced by the library's.
// - Insert and Take record the library version the copy now is, `@from(lib #…)` (the
//   library base, decided 2026-10-07); Publish sets it to the version it wrote
//   (`libraryBaseEdit`); so does *Record library base* on a copy identical to the library
//   that records none. Nothing else writes it. Each of them also hands the caller the
//   base's text to keep in `library-bases/` (bases.ts), for the three-way diff.
// - Publish: a project entry becomes a `.lambda` file: a new one with a generated header,
//   or the existing one with its header fields and rationale kept and only the definition
//   (and the summary and params, when they changed) replaced.

import { equalModuloWhitespace } from "../lang/format.js";
import { parseModule, formatEntry, formatDoc, formulaToSource, type ModuleEntry } from "../project/module.js";
import { compareNames, moduleFileName, UNMANAGED } from "../project/modules.js";
import { parseDocComment } from "../project/doc.js";
import { NAMES_DIR } from "../project/pull.js";
import { collapseSpace, lambdaFileName, lambdaParameters, parseLambdaFile, type LibraryFunction } from "./lambdaFile.js";
import { definitionBase, libraryClosure, libraryFunctionHash, type Library } from "./library.js";
import { FROM, buildProvenanceTag, docTagOverflow, formatLibBase } from "../project/provenance.js";
import { libraryBase, libraryFunctionBase, type LibraryBase } from "./bases.js";
import { libStateLabel, projectLibraryStatus } from "./status.js";

export interface TextEdit {
  start: number;
  end: number;
  text: string;
}

/** Edits to one project file; `create`: the file does not exist yet (the edits apply to ""). */
export interface FileEdit {
  path: string;
  create: boolean;
  edits: TextEdit[];
}

/** Applies non-overlapping edits to a text. */
export function applyEdits(text: string, edits: readonly TextEdit[]): string {
  let out = text;
  for (const e of [...edits].sort((a, b) => b.start - a.start || b.end - a.end)) out = out.slice(0, e.start) + e.text + out.slice(e.end);
  return out;
}

/** The module file a name belongs in: `FN.GROW` → `names/FN.xln`; no prefix → `names/_unmanaged.xln`. */
export function moduleFileOf(name: string): string {
  const dot = name.indexOf(".");
  return `${NAMES_DIR}/${dot > 0 ? moduleFileName(name.slice(0, dot)) : `${UNMANAGED}.xln`}`;
}

function entryText(fn: LibraryFunction): string {
  return formatEntry({ name: fn.name, doc: fn.doc, from: libraryFunctionHash(fn), formula: fn.definition });
}

/** The header of a module file created for a library function. */
export function libraryModuleHeader(module: string): string {
  return `// module: ${module}. Functions from the library are added here by xln (Insert library function); xln build writes them.\n`;
}

/** Where a new workbook name goes in a module file: its sort position among the names before the first `@scope` block. */
function insertionPoint(text: string, name: string): { at: number; before: boolean } {
  const pm = parseModule(text);
  const firstBlock = pm.scopes.find((s) => s.scope !== undefined)?.offset ?? Infinity;
  const entries = pm.entries.filter((e) => e.scope === undefined && e.start < firstBlock);
  const next = entries.find((e) => compareNames(e.name, name) > 0);
  if (next) return { at: next.start, before: true };
  const last = entries[entries.length - 1];
  if (last) return { at: last.end, before: false };
  return { at: Number.isFinite(firstBlock) ? firstBlock : text.length, before: Number.isFinite(firstBlock) };
}

/**
 * The edits that add library function `name` to a project (files: path → text), with the
 * library functions it calls that the project does not define. Undefined when the library
 * has no such function. `has`: whether the project defines a name (default: a
 * workbook-scoped name in one of `files`).
 */
export function libraryInsertion(library: Library, files: Readonly<Record<string, string>>, name: string, has?: (name: string) => boolean): { names: string[]; edits: FileEdit[]; bases: LibraryBase[] } | undefined {
  if (!library.get(name)) return undefined;
  let defined: Set<string> | undefined;
  const hasName =
    has ??
    ((n: string) => {
      if (!defined) {
        defined = new Set();
        for (const [p, t] of Object.entries(files)) {
          if (!p.startsWith(NAMES_DIR + "/") || !p.endsWith(".xln") || p.startsWith(`${NAMES_DIR}/sheets/`)) continue;
          for (const e of parseModule(t).entries) if (e.scope === undefined) defined.add(e.name.toLowerCase());
        }
      }
      return defined.has(n.toLowerCase());
    });
  const fns = libraryClosure(library, name, hasName);
  const byFile = new Map<string, LibraryFunction[]>();
  for (const fn of fns) {
    const path = moduleFileOf(fn.name);
    const list = byFile.get(path) ?? [];
    list.push(fn);
    byFile.set(path, list);
  }
  const edits: FileEdit[] = [];
  for (const [path, list] of byFile) {
    const existing = Object.keys(files).find((p) => p.toLowerCase() === path.toLowerCase());
    if (existing === undefined) {
      const module = path.slice(NAMES_DIR.length + 1, -".xln".length);
      const body = [...list].sort((a, b) => compareNames(a.name, b.name)).map(entryText);
      edits.push({ path, create: true, edits: [{ start: 0, end: 0, text: libraryModuleHeader(module) + "\n" + body.join("\n\n") + "\n" }] });
      continue;
    }
    // Each function at its sort position in the file as it is; several at one point go in name order.
    const text = files[existing]!;
    const points = new Map<string, { at: number; before: boolean; texts: string[] }>();
    for (const fn of [...list].sort((a, b) => compareNames(a.name, b.name))) {
      const p = insertionPoint(text, fn.name);
      const k = `${p.at}:${p.before}`;
      const pt = points.get(k) ?? { ...p, texts: [] };
      pt.texts.push(entryText(fn));
      points.set(k, pt);
    }
    const merged: TextEdit[] = [];
    for (const p of points.values()) {
      if (p.before) merged.push({ start: p.at, end: p.at, text: p.texts.map((t) => t + "\n\n").join("") });
      else {
        const lead = p.at === 0 ? "" : text.slice(0, p.at).endsWith("\n") ? "\n" : "\n\n";
        const trail = p.at >= text.length ? "\n" : "";
        merged.push({ start: p.at, end: p.at, text: lead + p.texts.join("\n\n") + trail });
      }
    }
    edits.push({ path: existing, create: false, edits: merged });
  }
  return { names: fns.map((f) => f.name), edits, bases: fns.map(libraryFunctionBase) };
}

/** The offsets of an entry's doc comment (`/** … *\/`), if it has one. */
function docRange(text: string, e: ModuleEntry): { start: number; end: number } | undefined {
  if (e.doc === undefined) return undefined;
  const start = text.indexOf("/**", e.start);
  if (start < 0 || start > e.offset) return undefined;
  const close = text.indexOf("*/", start + 3);
  if (close < 0 || close > e.offset) return undefined;
  return { start, end: close + 2 };
}

/** The offsets of an entry's formula in the file (from its first character to its last). */
export function formulaRange(e: ModuleEntry): { start: number; end: number } {
  return { start: formulaToSource(e, 0), end: e.formula.length === 0 ? formulaToSource(e, 0) : formulaToSource(e, e.formula.length - 1) + 1 };
}

/**
 * The edits that give a module file's entry the library's definition and doc comment
 * ("Take the library's version"), and records the library version it now is:
 * `@from(lib #…)`, written or updated. Other annotations and the name stay as written.
 */
export function libraryReplacement(text: string, entry: ModuleEntry, fn: LibraryFunction): TextEdit[] {
  const edits: TextEdit[] = [];
  const base = libraryFunctionHash(fn);
  const doc = docRange(text, entry);
  const hasFrom = entry.annotations.some((a) => a.name === FROM);
  if (doc) edits.push({ start: doc.start, end: doc.end, text: formatDoc(fn.doc) });
  // A doc comment and a missing `@from` inserted at one place go in one edit, in that order.
  else edits.push({ start: entry.start, end: entry.start, text: formatDoc(fn.doc) + "\n" + (hasFrom ? "" : formatLibBase(base) + "\n") });
  if (doc || hasFrom) {
    const b = libraryBaseEdit(text, entry, base);
    if (b) edits.push(b);
  }
  const f = formulaRange(entry);
  edits.push({ start: f.start, end: f.end, text: fn.definition });
  return edits;
}

/**
 * The edit that records `base` as an entry's library base: its `@from(lib #…)` rewritten,
 * or one added on the line above the name. Undefined when it already says so.
 */
export function libraryBaseEdit(text: string, entry: ModuleEntry, base: string): TextEdit | undefined {
  const want = formatLibBase(base);
  const a = entry.annotations.find((x) => x.name === FROM);
  if (a) return text.slice(a.offset, a.end) === want ? undefined : { start: a.offset, end: a.end, text: want };
  const lineStart = text.lastIndexOf("\n", entry.offset - 1) + 1;
  const indent = text.slice(lineStart, entry.offset);
  // The name starts its line (the usual layout): the annotation gets a line of its own above it.
  if (indent.trim() === "") return { start: entry.offset, end: entry.offset, text: `${want}\n${indent}` };
  return { start: entry.offset, end: entry.offset, text: `${want} ` };
}

/** The library version a project definition would be once published (`libraryHash` of its stored form). */
export function publishedBase(formula: string, name: string): string {
  return definitionBase(formula, name);
}

/** The entry of `name` in a module file's text (workbook scope). */
export function findEntry(text: string, name: string): ModuleEntry | undefined {
  return parseModule(text).entries.find((e) => e.scope === undefined && e.name.toLowerCase() === name.toLowerCase());
}

// ---- publish ----------------------------------------------------------------------------

/** What a project entry gives the library. */
export interface PublishSource {
  name: string;
  /** The doc comment (summary and `@param` lines). */
  doc: string | undefined;
  /** Display form. */
  formula: string;
}

export interface PublishResult {
  /** The `.lambda` file name (`FN.GROW.lambda`), or the existing file's path. */
  path: string;
  created: boolean;
  /** The whole new file text. */
  text: string;
  /** What changed, for the confirmation: `definition`, `summary`, `params`, `param docs`. */
  changed: string[];
  /** Why it cannot be published (not a LAMBDA). */
  error?: string;
  /** The library version the file holds afterwards: the copy's base once published (its text kept in `library-bases/`). */
  base?: LibraryBase;
}

const VALUE_COL = 13;
const WRAP = 88;

/** The first sentence of a text: up to a `.`, `!` or `?` followed by a blank or the end. */
export function firstSentence(s: string): { first: string; rest: string } {
  const t = collapseSpace(s);
  for (let i = 0; i < t.length; i++) {
    const c = t[i];
    if ((c === "." || c === "!" || c === "?") && (i + 1 === t.length || t[i + 1] === " ")) return { first: t.slice(0, i + 1), rest: t.slice(i + 2) };
  }
  return { first: t, rest: "" };
}

/** Whether a doc summary says what the library's does (the same words, or the library's shortened with `…`). */
export function sameSummary(doc: string, library: string): boolean {
  const a = collapseSpace(doc);
  const b = collapseSpace(library);
  if (a === b) return true;
  return a.endsWith("…") && b.startsWith(a.slice(0, -1).trimEnd());
}

function wrap(text: string, width: number): string[] {
  const out: string[] = [];
  let cur = "";
  for (const w of text.split(" ").filter((x) => x !== "")) {
    if (cur !== "" && cur.length + 1 + w.length > width) {
      out.push(cur);
      cur = w;
    } else cur = cur === "" ? w : `${cur} ${w}`;
  }
  if (cur !== "" || out.length === 0) out.push(cur);
  return out;
}

function fieldLines(key: string, value: string, col = VALUE_COL): string[] {
  const pad = Math.max(col - 2, key.length + 1);
  const lines = wrap(value, Math.max(20, WRAP - col));
  return lines.map((l, i) => (i === 0 ? `# ${key.padEnd(pad)}${l}` : `#${" ".repeat(pad + 1)}${l}`).trimEnd());
}

/** `@param` lines; an optional parameter keeps its brackets, as the `params` field writes it. */
function paramDocLines(params: { name: string; text: string }[], lambdaParams: readonly string[]): string[] {
  const optional = new Set(lambdaParams.filter((p) => p.startsWith("[")).map((p) => p.slice(1, -1).trim().toLowerCase()));
  return params.filter((p) => p.text !== "").map((p) => `# @param ${optional.has(p.name.toLowerCase()) ? `[${p.name}]` : p.name} ${p.text}`);
}

/**
 * The `.lambda` file for a project entry: `existing` is the library's file for that name
 * (path and text) when there is one.
 */
export function publishLambda(src: PublishSource, existing?: { path: string; text: string }): PublishResult {
  const r = publishText(src, existing);
  if (r.error) return r;
  const fn = parseLambdaFile(r.path, r.text).fn;
  return { ...r, base: libraryBase(src.name, fn?.definition ?? src.formula, r.path) };
}

function publishText(src: PublishSource, existing?: { path: string; text: string }): PublishResult {
  const params = lambdaParameters(src.formula);
  const path = existing?.path ?? lambdaFileName(src.name);
  if (!params) return { path, created: !existing, text: existing?.text ?? "", changed: [], error: `${src.name} is not a LAMBDA: only LAMBDAs go to the library` };
  const doc = parseDocComment(src.doc ?? "");
  const definition = src.formula.split("\n").map((l) => l.trimEnd()).join("\n");

  if (!existing) {
    const { first, rest } = firstSentence(doc.summary);
    const lines = [...fieldLines("name", src.name), ...fieldLines("summary", first), ...fieldLines("params", params.join(", "))];
    const rationale = [...(rest !== "" ? wrap(rest, WRAP - 2).map((l) => `# ${l}`) : []), ...paramDocLines(doc.params, params)];
    if (rationale.length) lines.push("#", ...rationale);
    return { path, created: true, text: [...lines, "", definition, ""].join("\n"), changed: ["new file"] };
  }

  const parsed = parseLambdaFile(existing.path, existing.text);
  const fn = parsed.fn;
  const crlf = existing.text.includes("\r\n");
  const lines = existing.text.split("\n").map((l) => (l.endsWith("\r") ? l.slice(0, -1) : l));
  if (!fn) {
    // An unreadable file is replaced whole: say so.
    const fresh = publishText(src);
    return { ...fresh, path, created: false, changed: ["the existing file could not be read; replaced whole"] };
  }
  const changed: string[] = [];
  // Line replacements, bottom-up so line numbers hold.
  const replace: { from: number; to: number; with: string[] }[] = [];
  const defOld = fn.definition;
  if (!equalModuloWhitespace(defOld, definition)) {
    changed.push("definition");
    replace.push({ from: fn.definitionLine - 1, to: lines.length, with: [definition, ""] });
  }
  const colOf = (line: number): number => {
    const l = lines[line - 1] ?? "";
    let i = l.indexOf("#") + 1;
    while (l[i] === " ") i++;
    while (i < l.length && l[i] !== " " && l[i] !== "\t") i++;
    while (l[i] === " " || l[i] === "\t") i++;
    return i;
  };
  const summaryField = fn.fields.find((f) => f.key === "summary");
  if (doc.summary !== "" && !sameSummary(doc.summary, fn.summary)) {
    changed.push("summary");
    if (summaryField) replace.push({ from: summaryField.line - 1, to: summaryField.endLine, with: fieldLines("summary", collapseSpace(doc.summary), colOf(summaryField.line)) });
    else {
      const nf = fn.fields.find((f) => f.key === "name")!;
      replace.push({ from: nf.endLine, to: nf.endLine, with: fieldLines("summary", collapseSpace(doc.summary), colOf(nf.line)) });
    }
  }
  const paramsField = fn.fields.find((f) => f.key === "params");
  const sameParams = params.length === fn.params.length && params.every((p, i) => p.toLowerCase() === fn.params[i]!.toLowerCase());
  if (!sameParams || !paramsField) {
    changed.push("params");
    if (paramsField) replace.push({ from: paramsField.line - 1, to: paramsField.endLine, with: fieldLines("params", params.join(", "), colOf(paramsField.line)) });
    else {
      const after = summaryField ?? fn.fields.find((f) => f.key === "name")!;
      replace.push({ from: after.endLine, to: after.endLine, with: fieldLines("params", params.join(", "), colOf(after.line)) });
    }
  }
  // Parameter descriptions live in the rationale as `@param` lines.
  const described = doc.params.filter((p) => p.text !== "");
  const libDocs = fn.paramDocs;
  const sameDocs = described.length === Object.keys(libDocs).length && described.every((p) => libDocs[p.name.toLowerCase()] === p.text);
  if (!sameDocs && described.length) {
    changed.push("param docs");
    // Old `@param` lines (and their indented continuations) of the comment block go; the new ones close it.
    const end = fn.definitionLine - 1;
    let blockEnd = end;
    while (blockEnd > 0 && lines[blockEnd - 1]!.trim() === "") blockEnd--;
    for (let k = 0; k < blockEnd; k++) {
      const body = lines[k]!.trimStart().startsWith("#") ? lines[k]!.trimStart().slice(1).trim() : "";
      if (!body.startsWith("@param ")) continue;
      let j = k + 1;
      while (j < blockEnd && /^#\s{2,}\S/.test(lines[j]!.trimStart()) && !lines[j]!.trimStart().slice(1).trim().startsWith("@")) j++;
      replace.push({ from: k, to: j, with: [] });
      k = j - 1;
    }
    replace.push({ from: blockEnd, to: blockEnd, with: [...(fn.rationale === "" ? ["#"] : []), ...paramDocLines(described, params)] });
  }
  if (changed.length === 0) return { path, created: false, text: existing.text, changed };
  replace.sort((a, b) => b.from - a.from || b.to - a.to);
  for (const r of replace) lines.splice(r.from, r.to - r.from, ...r.with);
  let text = lines.join("\n");
  if (!text.endsWith("\n")) text += "\n";
  if (crlf) text = text.split("\n").join("\r\n");
  return { path, created: false, text, changed };
}

// ---- room for the provenance tag --------------------------------------------------------

/** A name whose doc comment leaves no room for the tag the build adds. */
export interface DocTagWarning {
  name: string;
  /** The names file, project-relative. */
  path: string;
  message: string;
}

/**
 * Insert, Take, Publish and *Record library base* write `@from`, which the build carries in
 * the provenance tag, but a doc comment too long for the tag gets none (feedback 2026-10-07:
 * Publish succeeded, and only the next build said the base was not carried). This says it
 * before they write, on `files` as they would leave the project: for each of `names`, the
 * checker's warning (`docTagOverflow`) on the doc comment the build would tag (the
 * project's, not the library file's). The caller shows it and writes anyway: the author
 * decides.
 */
export function docTagWarnings(files: Readonly<Record<string, string>>, names: readonly string[]): DocTagWarning[] {
  const wanted = new Set(names.map((n) => n.toLowerCase()));
  const out: DocTagWarning[] = [];
  for (const path of Object.keys(files).sort()) {
    if (!path.startsWith(NAMES_DIR + "/") || !path.endsWith(".xln")) continue;
    const text = files[path]!;
    for (const e of parseModule(text).entries) {
      if (e.name === "" || !wanted.has(e.name.toLowerCase())) continue;
      const over = docTagOverflow(e.doc, buildProvenanceTag(path, text, e.from));
      if (over) out.push({ name: e.name, path, message: `${e.name}: ${over.message}` });
    }
  }
  return out;
}

// ---- record library base ----------------------------------------------------------------

/** One `@from(lib #…)` *Record library base* writes. */
export interface BaseRecording {
  name: string;
  /** The names file, project-relative. */
  path: string;
  /** The library version recorded. */
  hash: string;
  edit: TextEdit;
  /** The base's text, for `library-bases/`. */
  base: LibraryBase;
}

/**
 * *Record library base* (author's idea, 2026-10-07): a copy identical to the library that
 * records no base (inserted or published before `@from` existed) reads "differs" as soon as
 * either side changes. This records the library's version as its base, explicitly, never on
 * its own. `names`: the entries asked for (undefined: every identical entry without a base);
 * one that is not identical, or has a base already, is skipped with the reason.
 */
export function libraryBaseRecordings(
  library: Library,
  files: Readonly<Record<string, string>>,
  names?: readonly string[],
): { recorded: BaseRecording[]; skipped: { name: string; reason: string }[] } {
  const report = projectLibraryStatus(library, files, undefined, { target: "", library: "" });
  const byName = new Map(report.items.map((i) => [i.name.toLowerCase(), i]));
  const wanted = names ?? report.items.filter((i) => i.noBase).map((i) => i.name);
  const recorded: BaseRecording[] = [];
  const skipped: { name: string; reason: string }[] = [];
  for (const name of wanted) {
    const item = byName.get(name.toLowerCase());
    const fn = library.get(name);
    if (!fn) {
      skipped.push({ name, reason: "not in the library" });
      continue;
    }
    if (!item || item.state === "missing" || item.file === undefined) {
      skipped.push({ name: fn.name, reason: "not a workbook name of the project" });
      continue;
    }
    if (item.state !== "identical") {
      skipped.push({ name: fn.name, reason: `${libStateLabel(item.state)}, not identical to the library: only an identical copy gets the library's version as its base this way (Take or Publish records one)` });
      continue;
    }
    if (!item.noBase) {
      skipped.push({ name: fn.name, reason: `already records a base, @from(lib #${item.base})` });
      continue;
    }
    const text = files[item.file] ?? "";
    const entry = findEntry(text, fn.name);
    const base = libraryFunctionBase(fn);
    const edit = entry ? libraryBaseEdit(text, entry, base.hash) : undefined;
    if (!edit) {
      skipped.push({ name: fn.name, reason: `not found in ${item.file}` });
      continue;
    }
    recorded.push({ name: fn.name, path: item.file, hash: base.hash, edit, base });
  }
  return { recorded, skipped };
}
