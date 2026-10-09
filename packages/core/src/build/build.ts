// `xln build` without the file system: workbook bytes, project files and lockfile in;
// the plan, the new workbook bytes (read back, E3) and the new lockfile and manifest out.
// The CLI and the extension do the writing, the lock-file guard (E1) and the backup (E5).

import { readWorkbook } from "../file/workbook.js";
import type { DefinedName, WorkbookSnapshot } from "../file/types.js";
import { afeNamesTouched, afeStoreLabel } from "../audit/afe.js";
import type { Change } from "./changes.js";
import type { SourceProblem } from "./source.js";
import { cellKey, commentHash, definitionHash, LOCK_FORMAT, LOCK_FORMAT_IMPLICIT_SPILL, lockCell, lockfileText, parseLockfile, splitNameKey, type LockCell, type LockEntry, type Lockfile } from "../project/lockfile.js";
import { parseCellRange, type CellStatement } from "../project/statements.js";
import { readCellValues } from "./verify.js";
import { compareNames } from "../project/modules.js";
import { LOCK_FILE, MANIFEST_FILE, pullProject } from "../project/pull.js";
import { applyChangeSet, ApplyError } from "./apply.js";
import { describeChanges, scopedKey, type Scope, type SetEmbeddedSource } from "./changes.js";
import { CONFIG_FILE } from "../audit/config.js";
import { stripProvenance } from "../project/provenance.js";
import { readEmbeddedSource } from "./embedXml.js";
import { isSourcePath } from "./source.js";
import { addProvenance, checkCommentLengths } from "./tags.js";
import { planBuild, upgradeLockfile, type BuildPlan } from "./plan.js";
import { readBack, type ReadBackReport } from "./readback.js";
import { applySourceEdits, consumedRenamedEdits, type RenamedConsumed } from "./renameSource.js";

export interface BuildInput {
  /** The workbook as it is now. */
  workbook: Uint8Array;
  fileName: string;
  /** Project files, path → text (at least `names/**\/*.xln`). */
  files: Readonly<Record<string, string>>;
  /** Text of `xln.lock.json`; defaults to `files[LOCK_FILE]`. */
  lockText?: string;
}

export interface BuildOptions {
  /** Write the workbook even when there is nothing to change (it then only gains `fullCalcOnLoad`). */
  force?: boolean;
  /** Plan only: stop before writing (status `planned`). */
  dryRun?: boolean;
  /**
   * D5: embed the project source in the workbook (default false: opt-in since 2026-10-05,
   * `xln build --embed` or `"build": {"embed": true}` in xln.config.json). Without it an
   * existing part is left as it is: not refreshed, not removed.
   */
  embed?: boolean;
  /** D6: tag the comments of module names with their provenance (default true). */
  provenance?: boolean;
}

export type BuildStatus =
  /** New bytes ready to be written. */
  | "built"
  /** Source and workbook agree: nothing to write. */
  | "up-to-date"
  /** Errors in the source or conflicts with Excel: nothing written. */
  | "refused"
  /** The built bytes did not read back as intended: nothing may be written. */
  | "read-back-failed"
  /** A dry run: the plan has changes, nothing was applied. */
  | "planned";

export interface BuildResult {
  status: BuildStatus;
  plan: BuildPlan;
  /** For `built` (and for `read-back-failed`, to inspect). */
  bytes?: Uint8Array;
  readBack?: ReadBackReport;
  /** For `built`: the project files to write once the workbook is written. */
  files?: Record<string, string>;
  /**
   * For `built`: the `@renamed(…)` annotations of the renames this build made (and of
   * those an earlier build made, `earlier`), and the edits that remove them (offsets in
   * the source the build read). Written, with
   * `sourceFiles`, only where the lockfile is: when the workbook the project belongs to
   * is the one written.
   */
  renamedConsumed?: RenamedConsumed[];
  /** With `renamedConsumed`: the new text of each names file it changes. */
  sourceFiles?: Record<string, string>;
  /** Why the change set could not be applied, for `read-back-failed`. */
  error?: string;
}

/** `lbo.xlsx` → `lbo.xln.xlsx`: what a browser build writes beside the workbook (E1). */
export function builtCopyName(fileName: string): string {
  return withSuffix(fileName, ".xln");
}

/** `lbo.xlsx` → `lbo.backup.xlsx`: the previous file, kept beside the workbook (E5). */
export function backupName(fileName: string): string {
  return withSuffix(fileName, ".backup");
}

function withSuffix(fileName: string, suffix: string): string {
  const slash = Math.max(fileName.lastIndexOf("/"), fileName.lastIndexOf("\\"));
  const dot = fileName.lastIndexOf(".");
  return dot > slash + 1 ? fileName.slice(0, dot) + suffix + fileName.slice(dot) : fileName + suffix;
}

function baseName(path: string): string {
  const i = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
  return i < 0 ? path : path.slice(i + 1);
}

export function buildWorkbook(input: BuildInput, opts: BuildOptions = {}): BuildResult {
  const lockText = input.lockText ?? input.files[LOCK_FILE];
  if (lockText === undefined) throw new Error(`the project has no ${LOCK_FILE}: pull the workbook first`);
  const lock = parseLockfile(lockText);
  const wb = readWorkbook(input.workbook);
  const fileName = baseName(input.fileName);
  const plan = planBuild({ workbook: wb, fileName, files: input.files, lock, values: valueCells(input.workbook), bytes: input.workbook });
  if (opts.provenance !== false) addProvenance(plan, wb, input.files);
  checkCommentLengths(plan);
  const afe = afeBuildWarning(wb, plan.changeSet.changes);
  if (afe) plan.problems.push(afe);
  if (plan.conflicts.length > 0 || plan.problems.some((p) => p.severity === "error")) return { status: "refused", plan };
  const embed = opts.embed === true;
  if (plan.changeSet.changes.length === 0 && !opts.force && !embed) return { status: "up-to-date", plan };
  if (opts.dryRun && !embed) return { status: plan.changeSet.changes.length ? "planned" : "up-to-date", plan };

  let bytes: Uint8Array;
  try {
    bytes = applyChangeSet(input.workbook, plan.changeSet);
  } catch (e) {
    if (e instanceof ApplyError) return { status: "read-back-failed", plan, error: e.message };
    throw e;
  }

  // The built workbook as a pull sees it: its lockfile and manifest.
  const pulled = pullProject(bytes, fileName);
  const files: Record<string, string> = {
    [LOCK_FILE]: lockfileText(nextLockfile(upgradeLockfile(lock, plan), plan, pulled.snapshot.definedNames, pulled.snapshot.sheets.map((s) => s.name), pulled.statements)),
    [MANIFEST_FILE]: pulled.files[MANIFEST_FILE]!,
  };
  // The source files are the author's: a build never rewrites them (M3d), with one
  // exception: the `@renamed(…)` notes of the renames this build makes are spent once it
  // has made them, and go (2026-10-07; consumedRenamedEdits). The caller writes them with
  // the lockfile, only when it writes the workbook the project belongs to.
  // A note whose rename an earlier build made goes too, but only with a build that writes
  // something of its own: one with nothing to write leaves the source (the checker's
  // `renamed-built` hint and its quick fix cover that).
  const renamedConsumed = consumedRenamedEdits(input.files, plan.changeSet.changes, plan.changeSet.changes.length > 0 || opts.force ? lock : undefined);
  const sourceFiles = applySourceEdits(input.files, renamedConsumed);
  if (typeof sourceFiles === "string") throw new Error(sourceFiles);

  // D5: the project as it stands after this build goes into the workbook, unless the part already says exactly that.
  if (embed) {
    const payload = embeddedFiles({ ...input.files, ...sourceFiles, ...files });
    const had = readEmbeddedSource(input.workbook);
    if (!had || had.damaged.length > 0 || !sameFiles(had.files, payload)) {
      const op: SetEmbeddedSource = { op: "set-embedded-source", files: payload };
      plan.changeSet.changes.push(op);
      bytes = applyChangeSet(bytes, [op]);
    }
  }
  if (plan.changeSet.changes.length === 0 && !opts.force) return { status: "up-to-date", plan };
  if (opts.dryRun) return { status: "planned", plan };

  const rb = readBack(input.workbook, bytes, plan.changeSet.changes, plan.inSync);
  if (!rb.ok) return { status: "read-back-failed", plan, bytes, readBack: rb };
  const out: BuildResult = { status: "built", plan, bytes, readBack: rb, files };
  if (renamedConsumed.length) {
    out.renamedConsumed = renamedConsumed;
    out.sourceFiles = sourceFiles;
  }
  return out;
}

/** One line per source file the build changed: `removed @renamed(Rate) from names/_unmanaged.xln: the rename is built`. */
export function renamedConsumedLines(r: BuildResult): string[] {
  const byPath = new Map<string, string[]>();
  for (const e of r.renamedConsumed ?? []) byPath.set(e.path, [...(byPath.get(e.path) ?? []), e.annotation]);
  return [...byPath].map(([path, anns]) => renamedConsumedLine(path, anns));
}

/** `removed @renamed(Rate) from names/_unmanaged.xln: the rename is built`. */
export function renamedConsumedLine(path: string, annotations: readonly string[]): string {
  return `removed ${annotations.join(", ")} from ${path}: the rename${annotations.length === 1 ? " is" : "s are"} built`;
}

/**
 * The warning for a build into a workbook that carries AFE's copy of the names
 * (file/afe.ts): the copy keeps the text of before the build, and xln does not touch it
 * (no hidden transformations). Said only when the build changes names, and naming the
 * ones AFE's modules define when its store is readable. Undefined when there is nothing to say.
 */
export function afeBuildWarning(wb: WorkbookSnapshot, changes: readonly Change[]): SourceProblem | undefined {
  const stores = wb.foreignModuleStores.filter((s) => s.kind !== "locale-sheet");
  if (stores.length === 0) return undefined;
  const touched: string[] = [];
  for (const c of changes) {
    if (c.op === "set-name" && (c.fields.includes("created") || c.fields.includes("definition"))) touched.push(c.name);
    else if (c.op === "delete-name" || c.op === "rescope-name") touched.push(c.name);
    else if (c.op === "rename-name") touched.push(c.from, c.to);
  }
  if (touched.length === 0) return undefined;
  const where = stores.map(afeStoreLabel).join(" and ");
  const hit = afeNamesTouched(wb, touched);
  // A store xln can read and that defines none of these names has nothing to disagree about.
  if (hit.length === 0 && stores.every((s) => s.modules !== undefined)) return undefined;
  const list = hit.slice(0, 12).join(", ") + (hit.length > 12 ? `, … (${hit.length - 12} more)` : "");
  const what = hit.length
    ? `this build changes ${hit.length} name${hit.length === 1 ? "" : "s"} that AFE's modules also define (${list}), and AFE's copy keeps the text of before the build`
    : `xln cannot read it, so it may hold the text of before the build for names this build changes`;
  return {
    severity: "warning",
    code: "afe-modules",
    message:
      `the workbook carries modules of Microsoft's Advanced Formula Environment (Excel Labs) in ${where}; ${what}. ` +
      "xln leaves AFE's copy exactly as it is. When AFE opens it shows its own text, and saving its modules from AFE may write that text back over this build: " +
      "before saving modules in AFE, bring them in line with the Name Manager (edit them in AFE), or do not save AFE's modules. xln check lists where they differ (C14)",
  };
}

/**
 * What a build embeds (D5), an archive copy for a workbook handed on (pull never reads it,
 * 2026-10-06): every names file, the lockfile the build leaves, and the
 * audit config. Not the manifest: it is derived from the workbook and pull rewrites it.
 * In path order, so the same project embeds the same bytes.
 */
export function embeddedFiles(files: Readonly<Record<string, string>>): Record<string, string> {
  const out: Record<string, string> = {};
  const keep = (p: string) => isSourcePath(p) || p === LOCK_FILE || p === CONFIG_FILE;
  for (const p of Object.keys(files).filter(keep).sort()) out[p] = files[p]!;
  return out;
}

function sameFiles(a: Readonly<Record<string, string>>, b: Readonly<Record<string, string>>): boolean {
  const ka = Object.keys(a).sort();
  const kb = Object.keys(b).sort();
  return ka.length === kb.length && ka.every((k, i) => k === kb[i] && a[k] === b[k]);
}

/** Addresses of the cells holding a value, per sheet; undefined when they cannot be read. */
export function valueCells(bytes: Uint8Array): Map<string, Set<string>> | undefined {
  try {
    return new Map(readCellValues(bytes).map((s) => [s.sheet, new Set(s.cells.keys())]));
  } catch {
    return undefined;
  }
}

/**
 * The lockfile after a build: the names the build put in sync take their state from the
 * built file; a name Excel changed (or deleted) and the build kept keeps its old entry, so
 * the next build still sees the change; a name Excel created stays out of it. Format 4,
 * or 3 while the source still has statements without `#` read as on the spill (a project
 * pulled before the `#` was written: format 4 would read them as the cell alone); an old
 * entry kept from a format-1 or -2 lockfile keeps its `sha256:` hash, which every
 * comparison recognises (lockfile.ts).
 */
function nextLockfile(old: Lockfile, plan: BuildPlan, built: readonly DefinedName[], sheets: readonly string[], builtStatements: readonly CellStatement[]): Lockfile {
  const lk = (name: string, scope: Scope | undefined) => `${(scope ?? "").toLowerCase()}!${name.toLowerCase()}`;
  const entries = new Map<string, { key: string; name: string; scope: Scope; entry: LockEntry }>();
  const gone = new Set<string>();
  for (const c of plan.changeSet.changes) {
    if (c.op === "delete-name") gone.add(lk(c.name, c.scope));
    else if (c.op === "rename-name") gone.add(lk(c.from, c.scope));
    else if (c.op === "rescope-name") gone.add(lk(c.name, c.from));
  }
  for (const [key, entry] of Object.entries(old.names)) {
    const { name, scope } = splitNameKey(key);
    if (!gone.has(lk(name, scope))) entries.set(lk(name, scope), { key, name, scope: scope ?? null, entry });
  }
  const byKey = new Map(built.map((d) => [lk(d.name, d.scope.kind === "sheet" ? d.scope.name : null), d]));
  for (const s of plan.inSync) {
    const d = byKey.get(lk(s.name, s.scope));
    if (!d) continue;
    const scope = d.scope.kind === "sheet" ? d.scope.name : null;
    entries.set(lk(d.name, scope), {
      key: scope === null ? d.name : `${scope}!${d.name}`,
      name: d.name,
      scope,
      entry: { definition: definitionHash(d.definition), comment: commentHash(stripProvenance(d.comment === undefined ? undefined : d.comment.split("\r\n").join("\n"))), hidden: d.hidden },
    });
  }
  const pos = (scope: Scope) => (scope === null ? -1 : sheets.findIndex((s) => s.toLowerCase() === scope.toLowerCase()));
  const sorted = [...entries.values()].sort((a, b) => pos(a.scope) - pos(b.scope) || compareNames(a.name, b.name));
  const names: Record<string, LockEntry> = {};
  for (const e of sorted) names[e.key] = e.entry;
  const out: Lockfile = { format: plan.implicitSpill.length > 0 ? LOCK_FORMAT_IMPLICIT_SPILL : LOCK_FORMAT, workbook: old.workbook, names };

  // Cells (E6), the same way: the statements in sync take the built file's state (found by
  // the name, or by the range for an unnamed one); the others keep their old entry. A
  // format-1 lockfile gains `cells` once the source has cell statements.
  if (old.cells === undefined && plan.cellsInSync.length === 0) return out;
  const cells = new Map<string, { key: string; e: LockCell }>();
  const goneKeys = new Set<string>();
  for (const c of plan.changeSet.changes) {
    if (c.op === "delete-name") goneKeys.add(scopedKey(c.name, c.scope).toLowerCase());
    else if (c.op === "rename-name") goneKeys.add(scopedKey(c.from, c.scope).toLowerCase());
    else if (c.op === "rescope-name") goneKeys.add(scopedKey(c.name, c.from).toLowerCase());
  }
  for (const [key, e] of Object.entries(old.cells ?? {})) if (!goneKeys.has(key.toLowerCase())) cells.set(key.toLowerCase(), { key, e });
  const builtByKey = new Map(builtStatements.map((s) => [cellKey(s).toLowerCase(), s]));
  for (const s of plan.cellsInSync) {
    const key = cellKey({ key: s.name, sheet: s.sheet, range: s.range });
    const b = builtByKey.get(key.toLowerCase());
    const e: LockCell = b ? lockCell(b) : { sheet: s.sheet, range: s.range, formula: s.stored === null ? null : definitionHash(s.stored) };
    if (s.name !== undefined) e.name = s.name;
    cells.set(key.toLowerCase(), { key, e });
  }
  const sheetPos = (sheet: string) => sheets.findIndex((s) => s.toLowerCase() === sheet.toLowerCase());
  const sortedCells = [...cells.values()].sort((a, b) => {
    const ra = parseCellRange(a.e.range);
    const rb = parseCellRange(b.e.range);
    return sheetPos(a.e.sheet) - sheetPos(b.e.sheet) || (ra?.r1 ?? 0) - (rb?.r1 ?? 0) || (ra?.c1 ?? 0) - (rb?.c1 ?? 0);
  });
  out.cells = {};
  for (const c of sortedCells) out.cells[c.key] = c.e;
  return out;
}

/** Where a build problem is, for a report line: ` names/X.xln:12`, or "". */
function problemWhere(p: { file?: string; line?: number }): string {
  return p.file ? ` ${p.file}${p.line ? `:${p.line}` : ""}` : "";
}

/**
 * The body of a build report (the CLI's and the editor's output), indented. A refused
 * build starts with why (errors, then conflicts) before the plan it did not write, so
 * the reason is the first thing read (feedback 2026-10-07: a refusal reason printed last,
 * after the plan, went unnoticed). Then warnings, then notes.
 */
export function buildReportLines(r: BuildResult): string[] {
  const out: string[] = [];
  const errors = r.plan.problems.filter((p) => p.severity === "error");
  const warnings = r.plan.problems.filter((p) => p.severity !== "error");
  const conflicts: string[] = [];
  for (const c of r.plan.conflicts) {
    conflicts.push(`  CONFLICT ${c.kind} ${c.key}: ${c.message}`);
    if (c.excel) conflicts.push(`    in Excel:  ${c.excel.display.split("\n").join(" ")}`);
    if (c.source) conflicts.push(`    in source: ${c.source.display.split("\n").join(" ")}  (${c.source.file}:${c.source.line})`);
  }
  const refused = r.status === "refused";
  if (refused) {
    for (const p of errors) out.push(`  error${problemWhere(p)}: ${p.message}`);
    out.push(...conflicts);
    if (r.plan.changeSet.changes.length) out.push("  the plan, not written:");
  }
  for (const l of describeChanges(r.plan.changeSet.changes)) out.push(`  ${refused ? "  " : ""}${l}`);
  if (!refused) {
    out.push(...conflicts);
    for (const p of errors) out.push(`  error${problemWhere(p)}: ${p.message}`);
  }
  for (const p of warnings) out.push(`  ${p.severity}${problemWhere(p)}: ${p.message}`);
  for (const x of r.plan.excelChanges) out.push(`  note: ${x.message}`);
  for (const p of r.readBack?.ok === false ? r.readBack.problems : []) out.push(`  read-back: ${p}`);
  if (r.error) out.push(`  ${r.error}`);
  return out;
}

/** Why a build was refused, one line each: its errors (with file:line), then its conflicts. */
export function refusalReasons(r: BuildResult): string[] {
  return [
    ...r.plan.problems.filter((p) => p.severity === "error").map((p) => `${p.file ? `${p.file}${p.line ? `:${p.line}` : ""}: ` : ""}${p.message}`),
    ...r.plan.conflicts.map((c) => `${c.key}: ${c.message}`),
  ];
}
