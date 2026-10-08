// What `xln: Build workbook` needs besides the core's buildWorkbook: where the built file
// goes, how the summary reads, and the two sides of a conflict as text for a diff. Plain
// TypeScript, so Vitest can test it.

import { backupName, buildReportLines, builtCopyName, formatCellAddress, refusalReasons, formatEntry, type BuildResult, type Conflict, type NameState } from "@xln/core";

export interface BuildTarget {
  /** File name written beside the workbook. */
  file: string;
  /** Overwrites the workbook (desktop) or writes a new file next to it (browser, E1). */
  overwrite: boolean;
  /** Backup of the previous file (desktop only). */
  backup: string | undefined;
  /** The lockfile follows only when the workbook itself is rewritten. */
  updateLock: boolean;
}

/**
 * Desktop overwrites the workbook, after the lock-file guard, keeping a backup. The browser
 * cannot see Excel's `~$` owner file (measured in vscode.dev), so it never overwrites: it writes
 * `<name>.xln.xlsx` beside the workbook and leaves the original and the lockfile alone.
 */
export function buildTarget(workbook: string, web: boolean): BuildTarget {
  return web
    ? { file: builtCopyName(workbook), overwrite: false, backup: undefined, updateLock: false }
    : { file: workbook, overwrite: true, backup: backupName(workbook), updateLock: true };
}

/** The lines the output channel shows for a build. */
export function formatBuildSummary(workbook: string, r: BuildResult): string[] {
  const lines: string[] = [];
  const n = r.plan.changeSet.changes.length;
  const head: Record<BuildResult["status"], string> = {
    built: `${n} change${n === 1 ? "" : "s"}`,
    "up-to-date": "up to date: the workbook already matches the source",
    refused: `refused: ${r.plan.conflicts.length} conflict(s), ${r.plan.problems.filter((p) => p.severity === "error").length} error(s); nothing written`,
    "read-back-failed": "the built workbook did not read back as intended; nothing written",
    planned: `${n} change${n === 1 ? "" : "s"} planned; nothing written (dry run)`,
  };
  lines.push(`xln build ${workbook}: ${head[r.status]}`);
  // The CLI's report: a refusal's reasons first.
  lines.push(...buildReportLines(r));
  return lines;
}

/** The text of the modal a refused build shows: its reasons, the first few, then how many more. */
export function refusalMessage(workbook: string, r: BuildResult, max = 5): string {
  const reasons = refusalReasons(r);
  const shown = reasons.slice(0, max).map((s) => `• ${s}`);
  if (reasons.length > max) shown.push(`… and ${reasons.length - max} more`);
  return `Build of ${workbook} refused: nothing was written.\n\n${shown.join("\n")}`;
}

/** A build problem placed in a project file, for the Problems panel. */
export interface BuildDiagnostic {
  /** Project-relative path (`names/sheets/Ratios.xln`). */
  file: string;
  /** 0-based line. */
  line: number;
  message: string;
  code: string;
}

/**
 * The blocking errors of a build placed in their files (source `xln build`, apart from the
 * live checks): what refused the build, at its file:line. Problems without a place (a
 * conflict, a reader in a format) are in the modal and the output only.
 */
export function buildDiagnostics(r: BuildResult): BuildDiagnostic[] {
  const out: BuildDiagnostic[] = [];
  for (const p of r.plan.problems) if (p.severity === "error" && p.file !== undefined) out.push({ file: p.file, line: Math.max(0, (p.line ?? 1) - 1), message: p.message, code: p.code });
  for (const c of r.plan.conflicts) if (c.source) out.push({ file: c.source.file, line: Math.max(0, c.source.line - 1), message: c.message, code: `conflict.${c.kind}` });
  return out;
}

function entryText(s: NameState | undefined, absent: string): string {
  if (!s) return `// ${absent}\n`;
  // Module-file form: `@sheet(…)` above a local name. A cell statement (M3b) keeps its
  // address; a workbook-scoped or unnamed one names its sheet.
  const cell = s.cell && (s.scope === null || s.name === "" ? formatCellAddress(s.cell.range, s.cell.sheet) : s.cell.range);
  return `${formatEntry({ name: s.name, scope: s.name === "" ? undefined : (s.scope ?? undefined), hidden: s.hidden, doc: s.comment ?? undefined, formula: s.display, ...(cell ? { cell } : {}) })}\n`;
}

/** Both sides of a conflict as `.xln` text: Excel's version left, the source's right. */
export function conflictTexts(c: Conflict): { excel: string; source: string } {
  return {
    excel: entryText(c.excel, `${c.key} is not in the workbook (deleted in Excel)`),
    source: entryText(c.source, `${c.key} is not in the source (deleted there)`),
  };
}
