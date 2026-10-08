// `xln lib status | publish | take | base` (M4): the library folder is read here (node:fs),
// everything else is @xln/core. The library is only written by `lib publish` without
// `--dry-run`, one file, the one named. Publish, take and base also keep the base's text in
// the project's `library-bases/<hash>.json` (for the three-way diff); nothing else writes
// there, and nothing prunes it.

import { existsSync, mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { basename, dirname, extname, isAbsolute, join, resolve } from "node:path";
import {
  CONFIG_FILE,
  diffText,
  LOCK_FILE,
  MANIFEST_FILE,
  NAMES_DIR,
  lineDiff,
  libStatusJson,
  libraryStatus,
  projectLibraryStatus,
  parseConfig,
  parseModule,
  publishLambda,
  readLibrary,
  readWorkbook,
  renderLibStatus,
  workbookCopies,
  withKnownVersions,
  applyEdits,
  backupName,
  findEntry,
  libraryBaseEdit,
  libraryFunctionHash,
  libraryReplacement,
  libStateLabel,
  publishedBase,
  type LibState,
  type ModuleEntry,
  type WorkbookSnapshot,
  isLambdaFile,
  BASES_DIR,
  baseFiles,
  isBasePath,
  libraryBaseRecordings,
  docTagWarnings,
  type DocTagWarning,
  libraryFunctionBase,
  readBases,
  withBases,
  type BaseRecording,
  type LibraryBase,
  type Library,
  type LibStatusReport,
  type PublishResult,
} from "@xln/core";

export interface LibIo {
  out: (s: string) => void;
  err: (s: string) => void;
}

/** `~/x` → the home folder's x. */
export function expandHome(p: string): string {
  if (p === "~") return homedir();
  if (p.startsWith("~/") || p.startsWith("~\\")) return join(homedir(), p.slice(2));
  return p;
}

/** The library folder: `--lib` (relative to the current folder), else the project's `xln.config.json` (relative to the project). */
export function libraryDir(project: string | undefined, flag: string | undefined): string | { error: string } {
  if (flag !== undefined) return resolve(expandHome(flag));
  const config = project !== undefined ? join(project, CONFIG_FILE) : undefined;
  if (config === undefined || !existsSync(config)) return { error: `no library: pass --lib <dir>, or set "library" in ${project !== undefined ? join(project, CONFIG_FILE) : CONFIG_FILE}` };
  const { config: c, problems } = parseConfig(readFileSync(config, "utf8"));
  if (c.library === undefined) return { error: `no library: pass --lib <dir>, or set "library" in ${config}${problems.length ? ` (${problems.join("; ")})` : ""}` };
  const p = expandHome(c.library);
  return isAbsolute(p) ? p : resolve(project!, p);
}

/** The `.lambda` files of a folder, by file name. */
export function readLibraryDir(dir: string): Library {
  const files: Record<string, string> = {};
  for (const f of readdirSync(dir)) if (isLambdaFile(f)) files[f] = readFileSync(join(dir, f), "utf8");
  return readLibrary(files);
}

/** A project folder's names files, lockfile, manifest and kept library bases (path → text). */
export function readProjectFiles(dir: string): Record<string, string> {
  const files: Record<string, string> = {};
  const walk = (d: string, rel: string, depth: number): void => {
    if (depth > 6 || !existsSync(d)) return;
    for (const e of readdirSync(d, { withFileTypes: true })) {
      const r = `${rel}/${e.name}`;
      if (e.isDirectory()) walk(join(d, e.name), r, depth + 1);
      else if (e.name.endsWith(".xln")) files[r] = readFileSync(join(d, e.name), "utf8");
    }
  };
  walk(join(dir, NAMES_DIR), NAMES_DIR, 0);
  for (const f of [LOCK_FILE, MANIFEST_FILE]) if (existsSync(join(dir, f))) files[f] = readFileSync(join(dir, f), "utf8");
  const bases = join(dir, BASES_DIR);
  if (existsSync(bases)) {
    for (const f of readdirSync(bases)) {
      const r = `${BASES_DIR}/${f}`;
      if (isBasePath(r)) files[r] = readFileSync(join(bases, f), "utf8");
    }
  }
  return files;
}

/**
 * Keeps the bases' texts in the project's `library-bases/` (a file per version, the same
 * whoever writes it). Returns the paths written, project-relative: those new or different.
 */
export function writeBases(project: string, bases: Iterable<LibraryBase>): string[] {
  const written: string[] = [];
  for (const [rel, text] of Object.entries(baseFiles(bases))) {
    const path = join(project, ...rel.split("/"));
    if (existsSync(path) && readFileSync(path, "utf8") === text) continue;
    mkdirSync(dirname(path), { recursive: true });
    writeFileSync(path, text, "utf8");
    written.push(rel);
  }
  return written;
}

/** The workbook a project was pulled from (its manifest's, beside the folder), if it is there. */
export function projectWorkbook(dir: string, files: Record<string, string>): string | undefined {
  const m = files[MANIFEST_FILE];
  if (m === undefined) return undefined;
  try {
    const name = (JSON.parse(m) as { workbook?: unknown }).workbook;
    if (typeof name !== "string" || name === "") return undefined;
    const p = join(dirname(resolve(dir)), name);
    return existsSync(p) ? p : undefined;
  } catch {
    return undefined;
  }
}

function isWorkbookPath(p: string): boolean {
  const e = extname(p).toLowerCase();
  return e === ".xlsx" || e === ".xlsm";
}

/** `lbo.xlsx` → `lbo.xln` beside it. */
function projectOf(workbook: string): string {
  return join(dirname(workbook), basename(workbook, extname(workbook)) + ".xln");
}

export interface LibStatusCommand {
  target: string;
  lib?: string;
  json: boolean;
  diffs: boolean;
}

export function runLibStatus(cmd: LibStatusCommand): { report: LibStatusReport } | { error: string } {
  const target = resolve(cmd.target);
  if (!existsSync(target)) return { error: `no such file or folder: ${cmd.target}` };
  const workbookMode = isWorkbookPath(target) && statSync(target).isFile();
  const project = workbookMode ? projectOf(target) : target;
  const dir = libraryDir(existsSync(project) ? project : undefined, cmd.lib);
  if (typeof dir !== "string") return { error: dir.error };
  if (!existsSync(dir)) return { error: `no such library folder: ${dir}` };
  const library = readLibraryDir(dir);
  if (workbookMode) {
    const wb = readWorkbook(new Uint8Array(readFileSync(target)));
    const copies = workbookCopies(wb);
    const backup = backupOf(target);
    if (backup) withKnownVersions(copies, backup);
    // The project beside the workbook may keep the bases' texts.
    if (existsSync(join(projectOf(target), BASES_DIR))) withBases(copies, readBases(readProjectFiles(projectOf(target))));
    return { report: libraryStatus(library, copies, { target: basename(target), kind: "workbook", library: dir }) };
  }
  const files = readProjectFiles(target);
  if (!Object.keys(files).some((f) => f.endsWith(".xln"))) return { error: `${cmd.target} is neither a workbook nor a project folder (no ${NAMES_DIR}/*.xln)` };
  const wbPath = projectWorkbook(target, files);
  // The bases are the source's (`@from`); with the workbook, what of the source is not
  // built yet shows, and the workbook (and the build's backup of it) may hold the base's
  // text for a three-way diff.
  const wb = wbPath ? readWorkbook(new Uint8Array(readFileSync(wbPath))) : undefined;
  const backup = wbPath ? backupOf(wbPath) : undefined;
  return { report: projectLibraryStatus(library, files, wb, { target: basename(target), library: dir, ...(wbPath ? { workbook: basename(wbPath) } : {}), ...(backup ? { backup } : {}) }) };
}

/** The build's backup of a workbook (`model.backup.xlsx` beside it), when there is one that reads. */
function backupOf(workbook: string): WorkbookSnapshot | undefined {
  const p = join(dirname(workbook), backupName(basename(workbook)));
  if (!existsSync(p)) return undefined;
  try {
    return readWorkbook(new Uint8Array(readFileSync(p)));
  } catch {
    return undefined;
  }
}

export interface LibPublishCommand {
  project: string;
  name: string;
  lib?: string;
  dryRun: boolean;
  json: boolean;
}

export interface LibPublishOutcome {
  /** The library file written (or that would be). */
  file: string;
  result: PublishResult;
  /** The file's text before (empty for a new file). */
  before: string;
  written: boolean;
  /** The library base the project's entry now records (`@from(lib #…)`), and whether its names file was (or would be) changed for it. */
  base: { hash: string; file: string; changed: boolean };
  /** The base's text kept in `library-bases/` (written, or that would be). */
  kept: string[];
  /** The project's doc comment leaves no room for the provenance tag carrying the base (`docTagWarnings`). */
  warnings: DocTagWarning[];
}

export function runLibPublish(cmd: LibPublishCommand): LibPublishOutcome | { error: string } {
  const project = resolve(cmd.project);
  if (!existsSync(project) || !statSync(project).isDirectory()) return { error: `no such project folder: ${cmd.project}` };
  const dir = libraryDir(project, cmd.lib);
  if (typeof dir !== "string") return { error: dir.error };
  if (!existsSync(dir)) return { error: `no such library folder: ${dir}` };
  const files = readProjectFiles(project);
  let found: { doc: string | undefined; formula: string; name: string; file: string; entry: ModuleEntry } | undefined;
  for (const [path, text] of Object.entries(files)) {
    if (!path.endsWith(".xln")) continue;
    for (const e of parseModule(text).entries) {
      if (e.scope === undefined && !e.cell && e.name.toLowerCase() === cmd.name.toLowerCase()) found = { doc: e.doc, formula: e.formula, name: e.name, file: path, entry: e };
    }
  }
  if (!found) return { error: `${cmd.name} is not a workbook-scoped name of ${cmd.project}` };
  const library = readLibraryDir(dir);
  const fn = library.get(found.name);
  const before = fn ? readFileSync(join(dir, fn.path), "utf8") : "";
  const result = publishLambda({ name: found.name, doc: found.doc, formula: found.formula }, fn ? { path: fn.path, text: before } : undefined);
  if (result.error) return { error: result.error };
  const file = join(dir, result.path);
  const write = !cmd.dryRun && result.changed.length > 0;
  if (write) writeFileSync(file, result.text, "utf8");
  // The copy and the library are now equal: the entry records the published version as its base.
  const hash = publishedBase(found.formula, found.name);
  const text = files[found.file]!;
  const edit = libraryBaseEdit(text, found.entry, hash);
  const after = edit ? applyEdits(text, [edit]) : text;
  if (edit && !cmd.dryRun) writeFileSync(join(project, ...found.file.split("/")), after, "utf8");
  const kept = result.base ? (cmd.dryRun ? Object.keys(baseFiles([result.base])) : writeBases(project, [result.base])) : [];
  // The tag goes on the workbook comment, the project's doc comment (not the library file's summary).
  const warnings = docTagWarnings({ ...files, [found.file]: after }, [found.name]);
  return { file, result, before, written: write, base: { hash, file: found.file, changed: edit !== undefined }, kept, warnings };
}

export interface LibTakeCommand {
  project: string;
  name: string;
  lib?: string;
  dryRun: boolean;
  discard: boolean;
  json: boolean;
}

export interface LibTakeOutcome {
  /** The names file changed (or that would be), project-relative. */
  file: string;
  /** The copy's state before. */
  state: LibState;
  before: string;
  after: string;
  /** The library version recorded as the base. */
  base: string;
  written: boolean;
  /** The base's text kept in `library-bases/` (written, or that would be). */
  kept: string[];
  /** Refused: the copy holds an edit of its own (modified, both changed) or may (differs), and --discard was not given. */
  refused?: string;
  /** The doc comment taken leaves no room for the provenance tag carrying the base (`docTagWarnings`). */
  warnings: DocTagWarning[];
}

/**
 * "Take the library's version" from the command line: the entry's definition and doc
 * comment become the library's and `@from(lib #…)` records that version. A copy with an
 * edit of its own (or one that cannot be told) is only replaced with `--discard`.
 */
export function runLibTake(cmd: LibTakeCommand): LibTakeOutcome | { error: string } {
  const project = resolve(cmd.project);
  if (!existsSync(project) || !statSync(project).isDirectory()) return { error: `no such project folder: ${cmd.project}` };
  const dir = libraryDir(project, cmd.lib);
  if (typeof dir !== "string") return { error: dir.error };
  if (!existsSync(dir)) return { error: `no such library folder: ${dir}` };
  const library = readLibraryDir(dir);
  const fn = library.get(cmd.name);
  if (!fn) return { error: `${cmd.name} is not in the library (${dir})` };
  const files = readProjectFiles(project);
  let found: { file: string; entry: ModuleEntry } | undefined;
  for (const [path, text] of Object.entries(files)) {
    if (!path.endsWith(".xln")) continue;
    const e = findEntry(text, fn.name);
    if (e && !e.cell) found = { file: path, entry: e };
  }
  if (!found) return { error: `${cmd.name} is not a workbook-scoped name of ${cmd.project}` };
  const item = projectLibraryStatus(library, files, undefined, { target: basename(project), library: dir }).items.find((i) => i.name.toLowerCase() === fn.name.toLowerCase());
  const state = item?.state ?? "differs";
  const before = files[found.file]!;
  const after = applyEdits(before, libraryReplacement(before, found.entry, fn));
  const base = libraryFunctionHash(fn);
  const out: LibTakeOutcome = { file: found.file, state, before, after, base, written: false, kept: [], warnings: docTagWarnings({ ...files, [found.file]: after }, [fn.name]) };
  const own = state === "modified" || state === "both-changed" ? `${fn.name} was edited here since its library base (${libStateLabel(state)})` : state === "differs" ? `${fn.name} records no library base, so an edit of its own cannot be ruled out (differs)` : undefined;
  if (own && !cmd.discard) {
    out.refused = `${own}: taking the library's version discards that edit. Run again with --discard to take it anyway, or publish the edit (xln lib publish)`;
    return out;
  }
  if (!cmd.dryRun && after !== before) {
    writeFileSync(join(project, ...found.file.split("/")), after, "utf8");
    out.written = true;
  }
  const kept = libraryFunctionBase(fn);
  out.kept = cmd.dryRun ? Object.keys(baseFiles([kept])) : writeBases(project, [kept]);
  return out;
}

export interface LibBaseCommand {
  project: string;
  /** The entry; undefined with `all`. */
  name?: string;
  all: boolean;
  lib?: string;
  dryRun: boolean;
  json: boolean;
}

export interface LibBaseOutcome {
  recorded: BaseRecording[];
  skipped: { name: string; reason: string }[];
  /** Names files changed (or that would be), project-relative. */
  files: string[];
  /** `library-bases/` files written (or that would be). */
  kept: string[];
  written: boolean;
  /** Doc comments that leave no room for the provenance tag carrying the base (`docTagWarnings`). */
  warnings: DocTagWarning[];
}

/**
 * *Record library base* from the command line: `@from(lib #…)` on an entry identical to the
 * library that records no base (`--all`: every such entry), with the base's text kept.
 * Never done by anything else (status only says it can be).
 */
export function runLibBase(cmd: LibBaseCommand): LibBaseOutcome | { error: string } {
  const project = resolve(cmd.project);
  if (!existsSync(project) || !statSync(project).isDirectory()) return { error: `no such project folder: ${cmd.project}` };
  const dir = libraryDir(project, cmd.lib);
  if (typeof dir !== "string") return { error: dir.error };
  if (!existsSync(dir)) return { error: `no such library folder: ${dir}` };
  const library = readLibraryDir(dir);
  const files = readProjectFiles(project);
  const r = libraryBaseRecordings(library, files, cmd.all ? undefined : [cmd.name!]);
  const byFile = new Map<string, BaseRecording[]>();
  for (const x of r.recorded) byFile.set(x.path, [...(byFile.get(x.path) ?? []), x]);
  const after: Record<string, string> = { ...files };
  for (const [path, list] of byFile) after[path] = applyEdits(files[path]!, list.map((x) => x.edit));
  if (!cmd.dryRun) for (const path of byFile.keys()) writeFileSync(join(project, ...path.split("/")), after[path]!, "utf8");
  const bases = r.recorded.map((x) => x.base);
  const kept = cmd.dryRun ? Object.keys(baseFiles(bases)) : writeBases(project, bases);
  const warnings = docTagWarnings(after, r.recorded.map((x) => x.name));
  return { ...r, files: [...byFile.keys()], kept, written: !cmd.dryRun && r.recorded.length > 0, warnings };
}

export function libBaseText(cmd: LibBaseCommand, o: LibBaseOutcome): string {
  const what = cmd.all ? "--all" : cmd.name!;
  const lines: string[] = [];
  if (o.recorded.length === 0) lines.push(`xln lib base ${what}: nothing to record${cmd.all ? " (no function is identical to the library without a base)" : ""}`);
  else {
    lines.push(`xln lib base ${what}${cmd.dryRun ? " --dry-run: would record" : ": recorded"} the library base of ${o.recorded.length} function${o.recorded.length === 1 ? "" : "s"}${cmd.dryRun ? "" : "; build to carry it into the workbook"}`);
    for (const x of o.recorded) lines.push(`  ${x.name}  @from(lib #${x.hash})  ${x.path}  (text kept in ${BASES_DIR}/${x.hash}.json)`);
  }
  for (const s of o.skipped) lines.push(`  skipped ${s.name}: ${s.reason}`);
  return lines.join("\n") + "\n" + warningText(o.warnings);
}

/** The doc-comment-too-long-for-the-tag warnings, one line each (empty when none). */
export function warningText(warnings: readonly DocTagWarning[]): string {
  return warnings.map((w) => `warning: ${w.path}: ${w.message}\n`).join("");
}

export function libTakeText(cmd: LibTakeCommand, o: LibTakeOutcome): string {
  if (o.refused) return `xln lib take ${cmd.name}: refused: ${o.refused}\n${fileDiff(o.before, o.after)}\n`;
  if (o.after === o.before) return `xln lib take ${cmd.name}: ${o.file} already has the library's version; nothing to write\n`;
  const head = cmd.dryRun ? `xln lib take ${cmd.name} --dry-run: would update ${o.file} (${libStateLabel(o.state)} → the library's version, @from(lib #${o.base}))` : `xln lib take ${cmd.name}: updated ${o.file} (${libStateLabel(o.state)} → the library's version, @from(lib #${o.base})); build to write it into the workbook`;
  return `${head}\n${fileDiff(o.before, o.after)}\n${warningText(o.warnings)}`;
}

function split(t: string): string[] {
  const lines = t.split("\n").map((l) => (l.endsWith("\r") ? l.slice(0, -1) : l));
  if (lines.length && lines[lines.length - 1] === "") lines.pop();
  return lines;
}

/** The file's diff, `-` before, `+` after, with a line of context around each change. */
export function fileDiff(before: string, after: string): string {
  const d = lineDiff(split(before), split(after));
  const changed = (i: number) => d[i] !== undefined && d[i]!.op !== " ";
  const keep = d.map((_, i) => changed(i) || changed(i - 1) || changed(i + 1));
  const out: string[] = [];
  let gap = false;
  d.forEach((l, i) => {
    if (keep[i]) {
      if (gap) out.push("  …");
      gap = false;
      out.push(diffText([l]));
    } else gap = true;
  });
  if (gap) out.push("  …");
  return out.join("\n");
}

export function libPublishText(cmd: LibPublishCommand, o: LibPublishOutcome): string {
  const r = o.result;
  const base = o.base.changed ? `${cmd.dryRun ? "would record" : "recorded"} @from(lib #${o.base.hash}) on ${cmd.name} in ${o.base.file} (its library base; text kept in ${BASES_DIR}/${o.base.hash}.json)\n` : "";
  if (r.changed.length === 0) return `xln lib publish ${cmd.name}: ${o.file} already has this definition; nothing to write\n${base}${warningText(o.warnings)}`;
  const head = cmd.dryRun
    ? `xln lib publish ${cmd.name} --dry-run: would ${r.created ? "create" : "update"} ${o.file} (${r.changed.join(", ")})`
    : `xln lib publish ${cmd.name}: ${r.created ? "created" : "updated"} ${o.file} (${r.changed.join(", ")})`;
  return `${head}\n${fileDiff(o.before, r.text)}\n${base}${warningText(o.warnings)}`;
}

const LIB_USAGE = `usage: xln lib status <workbook.xlsx | project-folder> [--lib <dir>] [--json] [--no-diff]
       xln lib publish <project-folder> <Name> [--lib <dir>] [--dry-run] [--json]
       xln lib take <project-folder> <Name> [--lib <dir>] [--dry-run] [--discard] [--json]
       xln lib base <project-folder> <Name | --all> [--lib <dir>] [--dry-run] [--json]
`;

export function libUsage(): string {
  return LIB_USAGE;
}

export function mainLib(args: string[], io: LibIo): number {
  const sub = args[0];
  const rest = args.slice(1);
  const positional: string[] = [];
  let lib: string | undefined;
  let json = false;
  let dryRun = false;
  let diffs = true;
  let discard = false;
  let all = false;
  for (let i = 0; i < rest.length; i++) {
    const a = rest[i]!;
    if (a === "--lib") {
      lib = rest[++i];
      if (lib === undefined) {
        io.err(`xln: --lib needs a folder\n${LIB_USAGE}`);
        return 2;
      }
    } else if (a === "--json") json = true;
    else if (a === "--dry-run" && (sub === "publish" || sub === "take" || sub === "base")) dryRun = true;
    else if (a === "--all" && sub === "base") all = true;
    else if (a === "--discard" && sub === "take") discard = true;
    else if (a === "--no-diff" && sub === "status") diffs = false;
    else if (a.startsWith("--")) {
      io.err(`xln: unknown option ${a}\n${LIB_USAGE}`);
      return 2;
    } else positional.push(a);
  }
  const fail = (error: string): number => {
    if (json) io.out(JSON.stringify({ ok: false, error }, null, 2) + "\n");
    else io.err(`xln: ${error}\n`);
    return 2;
  };
  if (sub === "status" && positional.length === 1) {
    const r = runLibStatus({ target: positional[0]!, ...(lib !== undefined ? { lib } : {}), json, diffs });
    if ("error" in r) return fail(r.error);
    if (json) io.out(JSON.stringify({ ok: true, ...(libStatusJson(r.report) as object) }, null, 2) + "\n");
    else {
      // The project's settings that could not be used: a library key in the wrong place is why a library goes unread.
      const config = join(isWorkbookPath(positional[0]!) ? projectOf(resolve(positional[0]!)) : resolve(positional[0]!), CONFIG_FILE);
      if (existsSync(config)) for (const p of parseConfig(readFileSync(config, "utf8")).problems) io.out(`note: ${CONFIG_FILE}: ${p}\n`);
      io.out(renderLibStatus(r.report, { diffs }).text);
    }
    return 0;
  }
  if (sub === "publish" && positional.length === 2) {
    const cmd: LibPublishCommand = { project: positional[0]!, name: positional[1]!, dryRun, json, ...(lib !== undefined ? { lib } : {}) };
    const o = runLibPublish(cmd);
    if ("error" in o) return fail(o.error);
    if (json) {
      io.out(
        JSON.stringify(
          { ok: true, file: o.file, created: o.result.created, changed: o.result.changed, written: o.written, dryRun, base: o.base, kept: o.kept, warnings: o.warnings.map((w) => w.message), diff: fileDiff(o.before, o.result.text).split("\n") },
          null,
          2,
        ) + "\n",
      );
    } else io.out(libPublishText(cmd, o));
    return 0;
  }
  if (sub === "take" && positional.length === 2) {
    const cmd: LibTakeCommand = { project: positional[0]!, name: positional[1]!, dryRun, discard, json, ...(lib !== undefined ? { lib } : {}) };
    const o = runLibTake(cmd);
    if ("error" in o) return fail(o.error);
    if (json) {
      const body = { ok: o.refused === undefined, file: o.file, state: o.state, base: o.base, written: o.written, kept: o.kept, dryRun, warnings: o.warnings.map((w) => w.message), ...(o.refused ? { refused: o.refused } : {}), diff: fileDiff(o.before, o.after).split("\n") };
      io.out(JSON.stringify(body, null, 2) + "\n");
    } else if (o.refused) io.err(libTakeText(cmd, o));
    else io.out(libTakeText(cmd, o));
    return o.refused ? 1 : 0;
  }
  if (sub === "base" && (all ? positional.length === 1 : positional.length === 2)) {
    const cmd: LibBaseCommand = { project: positional[0]!, all, dryRun, json, ...(all ? {} : { name: positional[1]! }), ...(lib !== undefined ? { lib } : {}) };
    const o = runLibBase(cmd);
    if ("error" in o) return fail(o.error);
    // The one named and not recordable: say why, exit 1 (as take refuses).
    const refused = !all && o.recorded.length === 0;
    if (json) {
      const body = { ok: !refused, dryRun, written: o.written, recorded: o.recorded.map((x) => ({ name: x.name, file: x.path, base: x.hash })), skipped: o.skipped, files: o.files, kept: o.kept, warnings: o.warnings.map((w) => w.message) };
      io.out(JSON.stringify(body, null, 2) + "\n");
    } else if (refused) io.err(libBaseText(cmd, o));
    else io.out(libBaseText(cmd, o));
    return refused ? 1 : 0;
  }
  io.err(LIB_USAGE);
  return 2;
}
