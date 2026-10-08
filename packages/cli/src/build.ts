// `xln build` and `xln verify`: the file system around the core's buildWorkbook and
// verifyValues. The safety layer that needs the disk lives here: the lock-file guard
// (E1), the backup (E5), and the re-read of what was written (E3, after the in-memory
// read-back the core already did).

import { copyFileSync, existsSync, readdirSync, readFileSync, renameSync, statSync, writeFileSync } from "node:fs";
import { basename, dirname, join, relative, resolve, sep } from "node:path";
import {
  backupName,
  buildWorkbook,
  CONFIG_FILE,
  embedSetting,
  buildReportLines,
  parseConfig,
  isLocked,
  lockFileName,
  readBack,
  renamedConsumedLines,
  cellValueMap,
  labelNoticeLines,
  readWorkbook,
  renameLabelNotices,
  valueText,
  valuesWarning,
  verifyValues,
  type BuildResult,
  type Change,
  type VerifyReport,
} from "@xln/core";

export interface BuildCommand {
  workbook: string;
  /** Project folder; default `<workbook>.xln` beside it. */
  project?: string;
  /** Write the built workbook here instead of over the original (the original and the lockfile stay as they are). */
  out?: string;
  dryRun: boolean;
  force: boolean;
  /** D5: embed the project source; undefined follows the project's `build.embed` (default no). */
  embed?: boolean;
  /** D6: provenance tags on module names' comments (default true). */
  tags?: boolean;
  json: boolean;
  /** E7: close the workbook in desktop Excel first (without saving) and open it again after. */
  reopen?: boolean;
  /** With `reopen`: close it even when Excel has unsaved changes (they are lost). */
  discard?: boolean;
}

export type BuildExit =
  /** Built and written, or nothing to do, or a dry run. */
  | 0
  /** Refused: errors in the source, or conflicts with edits made in Excel. */
  | 1
  /** Bad usage, or the workbook or project cannot be read. */
  | 2
  /** Excel has the target open (`~$` owner file): nothing written. */
  | 3
  /** The built file did not read back as intended: nothing written, or the original restored. */
  | 4;

export interface BuildOutcome {
  exit: BuildExit;
  result?: BuildResult;
  /** The file written, if any. */
  written?: string;
  backup?: string;
  /** Project files rewritten (lockfile, manifest). */
  projectFiles: string[];
  /** The `@renamed(…)` annotations the build removed from the source, one line per file. */
  renamedRemoved?: string[];
  /** After a build that renamed names: the label cells that still read an old name, and the Find & Replace that fixes them in Excel (one block of lines per rename). */
  labelNotices?: string[][];
  message: string;
  /** What in the project's xln.config.json could not be used (a key unknown or misplaced). */
  configNotes?: string[];
}

/** A project folder's files as the build reads them: `.xln` and `.json` text, other files below names/ listed empty. */
export function readProjectTree(dir: string): Record<string, string> {
  const files: Record<string, string> = {};
  const walk = (d: string, depth: number) => {
    if (depth > 6) return;
    for (const e of readdirSync(d, { withFileTypes: true })) {
      const p = join(d, e.name);
      const rel = relative(dir, p).split(sep).join("/");
      if (e.isDirectory()) walk(p, depth + 1);
      else if (e.name.endsWith(".xln") || e.name.endsWith(".json")) files[rel] = readFileSync(p, "utf8");
      // Any other file below names/ is listed, unread: the checker refuses it (M3e).
      else if (rel.startsWith("names/")) files[rel] = "";
    }
  };
  walk(dir, 0);
  return files;
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  return a.length === b.length && a.every((x, i) => x === b[i]);
}

export function runBuild(cmd: BuildCommand, defaultProject: (wb: string) => string): BuildOutcome {
  const o = buildOnce(cmd, defaultProject);
  const config = join(resolve(cmd.project ?? defaultProject(resolve(cmd.workbook))), CONFIG_FILE);
  if (existsSync(config)) {
    const notes = parseConfig(readFileSync(config, "utf8")).problems;
    if (notes.length) o.configNotes = notes;
  }
  return o;
}

function buildOnce(cmd: BuildCommand, defaultProject: (wb: string) => string): BuildOutcome {
  const workbook = resolve(cmd.workbook);
  const project = resolve(cmd.project ?? defaultProject(workbook));
  if (!existsSync(project) || !statSync(project).isDirectory()) return { exit: 2, projectFiles: [], message: `no project folder ${project}: pull the workbook first` };
  const original = new Uint8Array(readFileSync(workbook));
  const files = readProjectTree(project);
  let result: BuildResult;
  const input = { workbook: original, fileName: basename(workbook), files };
  const opts = { force: cmd.force, embed: embedSetting(cmd.embed, files[CONFIG_FILE]), provenance: cmd.tags !== false };
  try {
    result = buildWorkbook(input, opts);
    // A dry run shows the plan even when the file backend cannot apply it yet (cell
    // formulas before M3b-2): the change set is what it reports.
    if (cmd.dryRun && result.status === "read-back-failed" && result.bytes === undefined) {
      const why = result.error;
      result = buildWorkbook(input, { ...opts, dryRun: true });
      return { exit: 0, result, projectFiles: [], message: `dry run: nothing written (the file backend cannot apply this change set: ${why})` };
    }
  } catch (e) {
    return { exit: 2, projectFiles: [], message: (e as Error).message };
  }
  if (result.status === "refused") return { exit: 1, result, projectFiles: [], message: "build refused: nothing written" };
  if (result.status === "read-back-failed") return { exit: 4, result, projectFiles: [], message: "the built workbook did not read back as intended: nothing written" };
  if (result.status === "up-to-date") return { exit: 0, result, projectFiles: [], message: "up to date: the workbook already matches the source" };
  if (cmd.dryRun) return { exit: 0, result, projectFiles: [], message: "dry run: nothing written" };

  const target = resolve(cmd.out ?? workbook);
  // E1: Excel keeps `~$<file>` beside an open workbook; re-listed right before writing.
  if (isLocked(target, readdirSync(dirname(target)))) {
    return { exit: 3, result, projectFiles: [], message: `Excel has ${basename(target)} open (${lockFileName(target)} exists): nothing written. Close it in Excel and build again, or run xln build --reopen to have xln close it and open it again after the build (it stops when Excel has unsaved changes, unless you add --discard).` };
  }
  // E5: the previous file stays beside the new one.
  let backup: string | undefined;
  if (existsSync(target)) {
    backup = join(dirname(target), backupName(basename(target)));
    copyFileSync(target, backup);
  }
  const bytes = result.bytes!;
  const temp = target + ".xln-tmp";
  writeFileSync(temp, bytes);
  renameSync(temp, target);
  // E3 on disk: what is in the file now must be what was built, and read back as built.
  const onDisk = new Uint8Array(readFileSync(target));
  const rb = readBack(original, onDisk, result.plan.changeSet.changes, result.plan.inSync);
  if (!sameBytes(onDisk, bytes) || !rb.ok) {
    if (backup) copyFileSync(backup, target);
    return { exit: 4, result, projectFiles: [], message: `the written file did not read back as built; ${backup ? "the original was restored" : "remove it"}: ${rb.problems.join("; ")}` };
  }
  const projectFiles: string[] = [];
  // The lockfile follows the workbook the project belongs to; a copy elsewhere leaves it alone.
  if (target === workbook) {
    for (const [rel, text] of Object.entries(result.files ?? {})) {
      writeFileSync(join(project, ...rel.split("/")), text, "utf8");
      projectFiles.push(rel);
    }
    // The `@renamed(…)` notes of the renames just built, or built earlier, are spent: the one source edit a build makes.
    for (const [rel, text] of Object.entries(result.sourceFiles ?? {})) writeFileSync(join(project, ...rel.split("/")), text, "utf8");
  }
  const renamed = target === workbook ? renamedConsumedLines(result) : [];
  const notices = labelNoticesOf(original, result.plan.changeSet.changes);
  const n = result.plan.changeSet.changes.length;
  return { exit: 0, result, written: target, ...(backup ? { backup } : {}), projectFiles, ...(renamed.length ? { renamedRemoved: renamed } : {}), ...(notices.length ? { labelNotices: notices } : {}), message: `built ${basename(target)}: ${n} change${n === 1 ? "" : "s"}` };
}

/** The label notices of the build's renames, read on the workbook before the build (its labels are the built file's too). */
export function labelNoticesOf(original: Uint8Array, changes: readonly Change[]): string[][] {
  if (!changes.some((c) => c.op === "rename-name")) return [];
  try {
    const values = () => {
      try {
        return cellValueMap(original);
      } catch {
        return undefined;
      }
    };
    return renameLabelNotices(readWorkbook(original), values, changes).map(labelNoticeLines);
  } catch {
    // The notice is advice: a workbook it cannot read costs the notice, not the build.
    return [];
  }
}

export function buildText(cmd: BuildCommand, o: BuildOutcome): string {
  const out: string[] = [`xln build ${basename(cmd.workbook)}: ${o.message}`];
  const r = o.result;
  if (r) {
    // A refusal's reasons first, then the plan it did not write, warnings, notes.
    out.push(...buildReportLines(r));
    if (r.status === "built" || r.status === "up-to-date" || r.status === "planned") {
      out.push(`  ${r.plan.unchanged} name${r.plan.unchanged === 1 ? "" : "s"} and ${r.plan.unchangedCells} cell statement${r.plan.unchangedCells === 1 ? "" : "s"} unchanged`);
    }
  }
  for (const n of o.configNotes ?? []) out.push(`  note: ${CONFIG_FILE}: ${n}`);
  if (o.backup) out.push(`  previous file kept as ${basename(o.backup)}`);
  if (o.projectFiles.length) out.push(`  updated ${o.projectFiles.join(", ")}`);
  for (const l of o.renamedRemoved ?? []) out.push(`  ${l}`);
  for (const block of o.labelNotices ?? []) for (const l of block) out.push(`  ${l}`);
  if (o.written) out.push(`  open it in Excel, save, then run: xln verify ${basename(o.written)}${o.backup ? "" : " --before <copy before the build>"}`);
  return out.join("\n") + "\n";
}

export function buildJson(o: BuildOutcome): unknown {
  const r = o.result;
  return {
    ok: o.exit === 0,
    exit: o.exit,
    status: r?.status ?? "error",
    message: o.message,
    written: o.written ?? null,
    backup: o.backup ?? null,
    projectFiles: o.projectFiles,
    renamedRemoved: o.renamedRemoved ?? [],
    labelNotices: o.labelNotices ?? [],
    changeSet: r?.plan.changeSet ?? null,
    conflicts: r?.plan.conflicts ?? [],
    problems: r?.plan.problems ?? [],
    excelChanges: r?.plan.excelChanges ?? [],
    unchanged: r?.plan.unchanged ?? 0,
    unchangedCells: r?.plan.unchangedCells ?? 0,
    readBack: r?.readBack ?? null,
    configNotes: o.configNotes ?? [],
  };
}

export interface VerifyCommand {
  workbook: string;
  /** The copy before the build; default the backup `<workbook>.backup.xlsx`. */
  before?: string;
  tolerance?: number;
  json: boolean;
}

export function runVerify(cmd: VerifyCommand): { report: VerifyReport; before: string; warnings: string[] } {
  const after = resolve(cmd.workbook);
  const before = resolve(cmd.before ?? join(dirname(after), backupName(basename(after))));
  if (!existsSync(before)) throw new Error(`no copy before the build at ${before}: pass --before <file>`);
  const report = verifyValues(new Uint8Array(readFileSync(before)), new Uint8Array(readFileSync(after)), cmd.tolerance !== undefined ? { tolerance: cmd.tolerance } : {});
  // A side whose values Excel did not calculate makes the comparison empty: say so first.
  const warnings = [valuesWarning(basename(before), report.beforeValues, "before"), valuesWarning(basename(after), report.afterValues, "after")].filter((w): w is string => w !== undefined);
  return { report, before, warnings };
}

export function verifyText(cmd: VerifyCommand, before: string, r: VerifyReport, warnings: readonly string[] = []): string {
  const out = warnings.map((w) => `warning: ${w}`);
  out.push(`xln verify ${basename(cmd.workbook)} against ${basename(before)}: ${r.cells} cells on ${r.sheets} sheets, ${r.changed.length} changed`);
  for (const c of r.changed.slice(0, 200)) out.push(`  ${c.sheet}!${c.cell}: ${valueText(c.before)} → ${valueText(c.after)}`);
  if (r.changed.length > 200) out.push(`  … ${r.changed.length - 200} more (--json lists all)`);
  if (r.sheetsAdded.length) out.push(`  sheets only after: ${r.sheetsAdded.join(", ")}`);
  if (r.sheetsRemoved.length) out.push(`  sheets only before: ${r.sheetsRemoved.join(", ")}`);
  return out.join("\n") + "\n";
}
