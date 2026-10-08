// The detail lines the xln output channel shows for the actions that have no CLI formatter
// to share (Rename Symbol, the audit's summary, the formula view). Plain TypeScript, so
// Vitest can test it; the header line is the caller's (log.ts `Activity`).

import { renderAuditReport, type AuditReport, type SourceRename } from "@xln/core";

/** What a Rename Symbol edits: files and counts, and what happened to `@renamed`. */
export function renameLogLines(r: SourceRename, project?: string): string[] {
  const byFile = new Map<string, number>();
  for (const e of r.edits) byFile.set(e.path, (byFile.get(e.path) ?? 0) + 1);
  const files = [...byFile].map(([p, n]) => `${p} (${n})`).join(", ");
  const lines = [
    `${project ? `${project}: ` : ""}${r.edits.length} edit(s) in ${byFile.size} file(s): ${files}`,
    `formulas reading ${r.from}: ${r.references} reference(s) in ${r.readers} statement(s)`,
  ];
  const annotation: Record<SourceRename["annotation"], string> = {
    added: `@renamed(${r.from}) added: the build renames the name in the workbook and rewrites it in the cells`,
    kept: "@renamed kept: an earlier one still says where the name is in the workbook",
    removed: "@renamed removed: renamed back to the workbook's name",
    none: "no @renamed: the name is not in the workbook yet",
  };
  lines.push(annotation[r.annotation]);
  lines.push("source edits only (unsaved): nothing reaches the workbook until a build");
  return lines;
}

/**
 * The audit's counts as the CLI's `xln check` prints them: its first line (header) and
 * the checks that found something; the census and the findings stay in the report.
 */
export function auditLogLines(r: AuditReport): string[] {
  const text = renderAuditReport(r, { maxPerCheck: 0 }).text;
  const top = text.split("\n\n")[0]!.split("\n");
  return [top[0]!, ...top.slice(1).filter((l) => !l.endsWith(" none") && !l.endsWith(" census below"))];
}
