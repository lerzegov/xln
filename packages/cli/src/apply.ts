// `xln apply`: applies a change set (JSON, the format `xln build --json` prints) to a
// workbook with the file backend, without a project: for hand-built change sets, the
// Excel checks of the backend, and agents. The same safety as `xln build`: the lock-file
// guard (E1), read-back before and after writing (E3), the backup (E5). The project's
// lockfile is not touched: pull again afterwards if the workbook has a project.

import { copyFileSync, existsSync, readdirSync, readFileSync, renameSync, writeFileSync } from "node:fs";
import { basename, dirname, join, resolve } from "node:path";
import {
  applyChangeSetWithReport,
  ApplyError,
  backupName,
  CHANGESET_FORMAT,
  describeChange,
  isLocked,
  lockFileName,
  readBack,
  type ApplyReport,
  type Change,
  type NameState,
  type ReadBackReport,
} from "@xln/core";

export interface ApplyCommand {
  workbook: string;
  changes: string;
  out?: string;
  dryRun: boolean;
  json: boolean;
}

export interface ApplyOutcome {
  /** 0 written (or dry run), 2 bad input, 3 Excel has the file open, 4 does not apply or read back. */
  exit: 0 | 2 | 3 | 4;
  message: string;
  changes: Change[];
  report?: ApplyReport;
  readBack?: ReadBackReport;
  written?: string;
  backup?: string;
}

function readChanges(path: string): Change[] {
  const json = JSON.parse(readFileSync(path, "utf8")) as unknown;
  const changes = Array.isArray(json) ? json : (json as { format?: string; changes?: unknown; changeSet?: { changes?: unknown } }).changeSet?.changes ?? (json as { changes?: unknown }).changes;
  if (!Array.isArray(changes)) throw new Error(`${path}: expected a change set (${CHANGESET_FORMAT}) or a list of changes`);
  return changes as Change[];
}

/** The names a change set sets, as the read-back checks them. */
function namesSet(changes: readonly Change[]): NameState[] {
  return changes.flatMap((c) => (c.op === "set-name" ? [{ name: c.name, scope: c.scope, display: c.display, comment: c.comment, hidden: c.hidden }] : []));
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  return a.length === b.length && a.every((x, i) => x === b[i]);
}

export function runApply(cmd: ApplyCommand): ApplyOutcome {
  const workbook = resolve(cmd.workbook);
  let changes: Change[];
  try {
    changes = readChanges(resolve(cmd.changes));
  } catch (e) {
    return { exit: 2, message: (e as Error).message, changes: [] };
  }
  const original = new Uint8Array(readFileSync(workbook));
  let bytes: Uint8Array;
  let report: ApplyReport;
  try {
    ({ bytes, report } = applyChangeSetWithReport(original, changes));
  } catch (e) {
    if (e instanceof ApplyError) return { exit: 4, message: `the change set does not apply: ${e.message}. Nothing written.`, changes };
    throw e;
  }
  const inSync = namesSet(changes);
  const rb = readBack(original, bytes, changes, inSync);
  if (!rb.ok) return { exit: 4, message: "the result did not read back as intended: nothing written", changes, report, readBack: rb };
  if (cmd.dryRun) return { exit: 0, message: "dry run: nothing written", changes, report, readBack: rb };

  const target = resolve(cmd.out ?? workbook);
  if (isLocked(target, readdirSync(dirname(target)))) {
    return { exit: 3, message: `Excel has ${basename(target)} open (${lockFileName(target)} exists): close it in Excel, or use --reopen. Nothing written.`, changes, report, readBack: rb };
  }
  let backup: string | undefined;
  if (existsSync(target)) {
    backup = join(dirname(target), backupName(basename(target)));
    copyFileSync(target, backup);
  }
  const temp = target + ".xln-tmp";
  writeFileSync(temp, bytes);
  renameSync(temp, target);
  const onDisk = new Uint8Array(readFileSync(target));
  const rb2 = readBack(original, onDisk, changes, inSync);
  if (!sameBytes(onDisk, bytes) || !rb2.ok) {
    if (backup) copyFileSync(backup, target);
    return { exit: 4, message: `the written file did not read back as built; ${backup ? "the original was restored" : "remove it"}`, changes, report, readBack: rb2 };
  }
  const n = changes.length;
  return { exit: 0, message: `applied ${n} change${n === 1 ? "" : "s"} to ${basename(target)}`, changes, report, readBack: rb2, written: target, ...(backup ? { backup } : {}) };
}

export function applyText(cmd: ApplyCommand, o: ApplyOutcome): string {
  const out = [`xln apply ${basename(cmd.workbook)}: ${o.message}`];
  for (const c of o.changes) out.push(`  ${describeChange(c)}`);
  if (o.report) {
    for (const [sheet, s] of Object.entries(o.report.sheets)) {
      const bits = [
        s.inserted.length ? `${s.inserted.length} cell${s.inserted.length === 1 ? "" : "s"} inserted` : "",
        s.ghosts.length ? `old spill emptied: ${s.ghosts.join(" ")}` : "",
        s.unshared.length ? `un-shared: ${s.unshared.join(" ")}` : "",
      ].filter((x) => x);
      if (bits.length) out.push(`  ${sheet}: ${bits.join("; ")}`);
    }
    if (o.report.metadata === "created" || o.report.metadata === "extended") out.push(`  xl/metadata.xml: dynamic-array record ${o.report.metadata === "created" ? "created" : "added"}`);
    if (o.report.calcChainDropped) out.push(`  ${o.report.calcChainDropped} dropped (Excel rebuilds it)`);
  }
  if (o.readBack && !o.readBack.ok) for (const p of o.readBack.problems) out.push(`  read-back: ${p}`);
  if (o.backup) out.push(`  previous file kept as ${basename(o.backup)}`);
  return out.join("\n") + "\n";
}

export function applyJson(o: ApplyOutcome): unknown {
  return { ok: o.exit === 0, exit: o.exit, message: o.message, changes: o.changes, report: o.report ?? null, readBack: o.readBack ?? null, written: o.written ?? null, backup: o.backup ?? null };
}
