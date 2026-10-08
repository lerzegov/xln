// Rename in the source (M5, completing D4): the edits `xln rename` and the editor's Rename
// Symbol (F2) make to a project's `.xln` files. A rename in the source is otherwise
// indistinguishable from a deletion plus a creation, and nothing renames on its own (no
// hidden transformations), so it is explicit: the name's statement gets the new name and
// `@renamed(Old)` above it, the record the next build reads; every reader in the
// project's files (definitions, cell statements) gets the new name token, found on the
// parsed formulas (strings, LET/LAMBDA variables of that spelling, other scopes' names
// untouched). The build then renames the name and rewrites the same token in the
// workbook's formulas (stretch G), so source, lockfile and workbook agree again.
//
// `@renamed` records where the name was at the last pull or build (the lockfile): a name
// renamed twice before a build keeps its first `@renamed`; renamed back, it loses it; a
// name created in the source and not built yet gets none.

import { defKey, SourceModel, type NameDef } from "../check/model.js";
import { annotationRemoval, invalidName, renameBuilt, renamedFrom, xlmNameWarning } from "../check/check.js";

export { annotationRemoval };
import { parseLockfile, type Lockfile } from "../project/lockfile.js";
import { LOCK_FILE } from "../project/pull.js";
import { parseSourceFile } from "../project/module.js";
import type { Change } from "./changes.js";
import { isSourcePath, RENAMED } from "./source.js";

export interface SourceEdit {
  /** Project-relative path. */
  path: string;
  /** Offsets in the file's text before the edit. */
  start: number;
  end: number;
  text: string;
  /** What the edit is, for a preview. */
  label: "name" | "reference" | "annotation";
}

export interface SourceRename {
  /** The name's key before (`Rate`, `S2!Loc`) and after. */
  from: string;
  to: string;
  /** Every edit, files in path order, offsets ascending. */
  edits: SourceEdit[];
  /** The new text of each file the rename changes. */
  files: Record<string, string>;
  /** Readers rewritten (formula tokens), and in how many statements. */
  references: number;
  readers: number;
  /** `added`: `@renamed(Old)` written; `kept`: an earlier one still says where the name was; `removed`: renamed back; `none`: not built yet, nothing to record. */
  annotation: "added" | "kept" | "removed" | "none";
  /** What the new name risks without being refused (a LAMBDA named like an Excel 4.0 macro function). */
  warnings: string[];
}

export interface SourceRenameOptions {
  /** The lockfile (the last pull or build): where the name is in the workbook. Without it the name is taken to be there as the source has it. */
  lock?: Lockfile;
  /** Table names of the workbook: a name may not take one. */
  tables?: readonly string[];
}

const lower = (s: string) => s.toLowerCase();

/** `S2!Loc`, `'Cash Flow'!Total`, `Rate` → sheet and name. */
function splitKey(key: string): { sheet: string | undefined; name: string } {
  const bang = key.lastIndexOf("!");
  if (bang < 0) return { sheet: undefined, name: key.trim() };
  let sheet = key.slice(0, bang).trim();
  if (sheet.length >= 2 && sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).split("''").join("'");
  return { sheet: sheet === "" ? undefined : sheet, name: key.slice(bang + 1).trim() };
}

/** The definition `name` means: `Sheet!Name`, else the workbook's, else the only local one. */
export function findRenameTarget(model: SourceModel, name: string): NameDef | string {
  const { sheet, name: id } = splitKey(name);
  if (sheet !== undefined) return model.lookup(defKey(id, sheet)) ?? `no name ${id} local to ${sheet} in the project`;
  const wb = model.lookup(id);
  if (wb) return wb;
  const locals = model.defs.filter((d) => d.scope !== undefined && lower(d.name) === lower(id) && model.lookup(d.key) === d);
  if (locals.length === 1) return locals[0]!;
  if (locals.length > 1) return `${id} is local to ${locals.map((d) => d.scope).join(", ")}: say which (${locals[0]!.key})`;
  return `no name ${id} in the project`;
}

/** Every reader of `key` with the key each `use` reads, statement by statement (defs then cell statements, in file order). */
function readings(model: SourceModel): { def: NameDef; keys: (string | undefined)[]; locals: number }[] {
  return [...model.defs, ...model.cellDefs].map((def) => {
    const occ = model.analysis(def).occurrences;
    return { def, keys: occ.filter((o) => o.kind === "use").map((o) => (o.kind === "use" && o.key !== undefined ? lower(o.key) : undefined)), locals: occ.filter((o) => o.kind === "local").length };
  });
}

/**
 * The edits that rename `name` (`Rate`, `S2!Loc`) to `to` in a project's source. Returns a
 * message instead when the rename cannot be made: unknown name, an invalid or taken new
 * name, or a reader the new name would make read something else (a LET or LAMBDA
 * variable, a sheet's local name, an unknown name of that spelling).
 */
export function renameInSource(model: SourceModel, name: string, to: string, opts: SourceRenameOptions = {}): SourceRename | string {
  const def = findRenameTarget(model, name);
  if (typeof def === "string") return def;
  if (def.name === "") return `${name} is not a name`;
  to = to.trim();
  if (to === def.name) return `${def.key} is already called ${to}`;
  const newKey = defKey(to, def.scope);
  const taken = model.lookup(newKey);
  if (taken && taken !== def) return `${newKey} exists already (${taken.file.path}:${taken.entry.line})`;
  const table = (opts.tables ?? []).find((t) => lower(t) === lower(to));
  if (table) return `${to} is the name of a Table (${table}): a defined name cannot take it`;

  const edits: SourceEdit[] = [];
  const nameAt = model.nameLoc(def);
  edits.push({ ...nameAt, text: to, label: "name" });
  const readers = new Set<NameDef>();
  for (const d of [...model.defs, ...model.cellDefs]) {
    for (const o of model.analysis(d).occurrences) {
      if (o.kind !== "use" || o.key === undefined || lower(o.key) !== lower(def.key)) continue;
      edits.push({ ...model.loc(d, o.span), text: to, label: "reference" });
      readers.add(d);
    }
  }

  // `@renamed`: where the name is in the workbook, from the lockfile.
  const locked = opts.lock ? new Set(Object.keys(opts.lock.names).map(lower)) : undefined;
  const inWorkbook = (key: string) => locked === undefined || locked.has(lower(key));
  const ann = def.entry.annotations.find((a) => a.name === "renamed");
  const r = renamedFrom(def.entry.annotations);
  const fromAnn = r ? defKey(r.name, r.scope === undefined ? def.scope : (r.scope ?? undefined)) : undefined;
  let base: { key: string; name: string; scope: string | undefined } | undefined;
  if (fromAnn !== undefined && inWorkbook(fromAnn) && !(locked && locked.has(lower(def.key)))) base = { key: fromAnn, ...splitKey(fromAnn), scope: r!.scope === undefined ? def.scope : (r!.scope ?? undefined) };
  else if (inWorkbook(def.key)) base = { key: def.key, name: def.name, scope: def.scope };
  let annotation: SourceRename["annotation"] = "none";
  const file = def.file.text;
  const back = base !== undefined && lower(base.name) === lower(to) && lower(base.scope ?? "") === lower(def.scope ?? "");
  // A new name must be one Excel takes and xln writes; the name the workbook has is fine to go back to.
  const bad = back ? undefined : invalidName(to, def.entry.formula);
  if (bad) return `${to}: ${bad}`;
  if (base !== undefined) {
    if (back) {
      if (ann) {
        edits.push({ path: def.file.path, ...annotationRemoval(file, ann), text: "", label: "annotation" });
        annotation = "removed";
      }
    } else if (ann && base.key === fromAnn) annotation = "kept";
    else {
      const text = `@renamed(${def.name})`;
      if (ann) edits.push({ path: def.file.path, start: ann.offset, end: ann.end, text, label: "annotation" });
      else {
        const lineStart = file.lastIndexOf("\n", nameAt.start - 1) + 1;
        const indent = /^[ \t]*/.exec(file.slice(lineStart, nameAt.start))![0];
        edits.push({ path: def.file.path, start: lineStart, end: lineStart, text: `${indent}${text}\n`, label: "annotation" });
      }
      annotation = "added";
    }
  }

  edits.sort((a, b) => (a.path < b.path ? -1 : a.path > b.path ? 1 : a.start - b.start || a.end - b.end));
  const sources: Record<string, string> = {};
  for (const [path, f] of model.files) sources[path] = f.text;
  const files = applySourceEdits(sources, edits);
  if (typeof files === "string") return files;

  // Every reader must read after the rename what it read before.
  const after = new SourceModel();
  for (const [path, f] of model.files) after.setFile(path, files[path] ?? f.text);
  const was = readings(model);
  const now = readings(after);
  const oldKey = lower(def.key);
  const captured: string[] = [];
  if (was.length !== now.length) return "internal: the rename changed the statements";
  was.forEach((w, i) => {
    const n = now[i]!;
    const want = w.keys.map((k) => (k === oldKey ? lower(newKey) : k));
    const same = n.locals === w.locals && n.keys.length === want.length && n.keys.every((k, j) => k === want[j]);
    if (!same) captured.push(`${w.def.file.path}:${w.def.entry.line} (${w.def.name || "@" + (w.def.entry.cell?.range ?? "")})`);
  });
  if (captured.length) return `renaming ${def.key} to ${to} would change what ${captured.length === 1 ? "this formula reads" : "these formulas read"} (a LET or LAMBDA variable, a sheet's local name, or an unknown name spelled ${to}): ${captured.slice(0, 10).join(", ")}${captured.length > 10 ? ", …" : ""}. Choose another name`;

  const xlm = back ? undefined : xlmNameWarning(to, def.entry.formula);
  return { from: def.key, to: newKey, edits, files, references: edits.filter((e) => e.label === "reference").length, readers: readers.size, annotation, warnings: xlm ? [xlm] : [] };
}

/** `renameInSource` on a project's files (path → text): its `names/**` and lockfile. */
export function renameInProject(projectFiles: Readonly<Record<string, string>>, name: string, to: string, opts: Omit<SourceRenameOptions, "lock"> = {}): SourceRename | string {
  const model = new SourceModel();
  for (const [path, text] of Object.entries(projectFiles)) if (isSourcePath(path)) model.setFile(path, text);
  let lock: Lockfile | undefined;
  const lockText = projectFiles[LOCK_FILE];
  if (lockText !== undefined) {
    try {
      lock = parseLockfile(lockText);
    } catch (e) {
      return `the lockfile cannot be read: ${(e as Error).message}`;
    }
  }
  return renameInSource(model, name, to, { ...opts, ...(lock ? { lock } : {}) });
}

/**
 * The text of each file `edits` change (path → new text); the others are left out. Edits
 * of a file must not overlap; a message says so when they do.
 */
export function applySourceEdits(files: Readonly<Record<string, string>>, edits: readonly SourceEdit[]): Record<string, string> | string {
  const byPath = new Map<string, SourceEdit[]>();
  for (const e of edits) byPath.set(e.path, [...(byPath.get(e.path) ?? []), e]);
  const out: Record<string, string> = {};
  for (const [path, list] of byPath) {
    const text = files[path];
    if (text === undefined) return `internal: no file ${path}`;
    list.sort((a, b) => a.start - b.start || a.end - b.end);
    let res = "";
    let at = 0;
    for (const e of list) {
      if (e.start < at) return `internal: overlapping edits in ${path}`;
      res += text.slice(at, e.start) + e.text;
      at = e.end;
    }
    out[path] = res + text.slice(at);
  }
  return out;
}

/** One spent `@renamed(…)` a build removes, and the edit that removes it from the source. */
export interface RenamedConsumed extends SourceEdit {
  /** The annotation as written: `@renamed(Old)`. */
  annotation: string;
  /** The name's key now (`Pace`, `S2!Spot`). */
  key: string;
  /** 1-based line of the annotation, before the edit. */
  line: number;
  /** The rename was in the workbook before this build (an older build, a browser build's copy, another tool, a note written by hand). */
  earlier?: true;
}

/**
 * The edits that remove the spent `@renamed(…)` annotations once a build has written the
 * workbook (author's decisions, 2026-10-07). `@renamed` is a note of a pending change: it
 * tells the build that a name is the old one renamed, not a deletion plus a creation.
 * Once the rename is in the workbook the note is spent, and the next pull would drop it
 * anyway, so the build removes it itself. This is the one exception to "a build never
 * changes the source": it removes exactly the annotations whose rename (`rename-name`) or
 * scope change (`rescope-name`) is in `changes`, and, given the lockfile the build read,
 * those whose rename was built already (`renameBuilt`: the build ignores them; left by an
 * older build, a browser build, another tool or a hand), nothing else. A pending rename
 * and a change of spelling only (`@renamed(rate)` on `Rate`) stay. Files in path order,
 * offsets ascending.
 */
export function consumedRenamedEdits(files: Readonly<Record<string, string>>, changes: readonly Change[], lock?: Lockfile): RenamedConsumed[] {
  const eq = (a: string | null | undefined, b: string | null | undefined) => lower(a ?? "") === lower(b ?? "");
  const renames = changes.filter((c) => c.op === "rename-name");
  const rescopes = changes.filter((c) => c.op === "rescope-name");
  const locked = lock ? new Set(Object.keys(lock.names).map(lower)) : undefined;
  if (renames.length === 0 && rescopes.length === 0 && !locked) return [];
  const out: RenamedConsumed[] = [];
  for (const path of Object.keys(files).filter(isSourcePath).sort()) {
    const text = files[path]!;
    for (const e of parseSourceFile(path, text).entries) {
      const a = e.annotations.find((x) => x.name === RENAMED);
      const r = renamedFrom(e.annotations);
      if (!a || !r || e.name === "") continue;
      // `@renamed(Old)`: the name's own scope; `@renamed(Sheet!Old)`, `@renamed(!Old)` (workbook scope): the one written.
      const old = r.name;
      const oldScope = r.scope === undefined ? e.scope : (r.scope ?? undefined);
      const applied = !eq(old, e.name)
        ? renames.some((c) => eq(c.scope, oldScope) && eq(c.from, old) && eq(c.to, e.name))
        : !eq(oldScope, e.scope) && rescopes.some((c) => eq(c.name, e.name) && eq(c.from, oldScope) && eq(c.to, e.scope));
      const earlier = !applied && locked !== undefined && renameBuilt(e.name, e.scope, r, (n, sc) => locked.has(lower(defKey(n, sc))));
      if (!applied && !earlier) continue;
      out.push({ path, ...annotationRemoval(text, a), text: "", label: "annotation", annotation: text.slice(a.offset, a.end), key: defKey(e.name, e.scope), line: a.line, ...(earlier ? { earlier: true as const } : {}) });
    }
  }
  return out;
}
