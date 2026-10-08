// What `xln: Pull workbook` needs besides the core's `pullProject`: where the project goes
// and how the summary reads. The CLI makes the same choices (packages/cli/src/main.ts).

import { formatUnbuiltEdit, LOCK_FILE, MANIFEST_FILE, NAMES_DIR, type PullReport, type UnbuiltEdit } from "@xln/core";

const WORKBOOK_EXTENSIONS = [".xlsx", ".xlsm"];

export function isWorkbookName(name: string): boolean {
  const l = name.toLowerCase();
  return !name.startsWith("~$") && WORKBOOK_EXTENSIONS.some((e) => l.endsWith(e));
}

/** `lbo.xlsx` → `lbo.xln`, the project folder beside the workbook. */
export function projectFolderName(workbook: string): string {
  const dot = workbook.lastIndexOf(".");
  return (dot > 0 ? workbook.slice(0, dot) : workbook) + ".xln";
}

/** A folder is a project when it holds the lockfile, the manifest and `names/`. */
export function isProjectListing(entries: readonly { name: string; isDirectory: boolean }[]): boolean {
  const has = (n: string, dir: boolean) => entries.some((e) => e.name === n && e.isDirectory === dir);
  return has(LOCK_FILE, false) && has(MANIFEST_FILE, false) && has(NAMES_DIR, true);
}

/** The notification after a pull. */
export function pullNotification(workbook: string, names: number, discarded = 0, foreign: PullReport["foreignModules"] = [], labels: PullReport["valueLabels"] = []): string {
  return `xln: pulled ${names} names from ${workbook} into ${projectFolderName(workbook)}.${discarded ? ` ${discarded} source edit(s) not built were replaced.` : ""}${labelSentence(labels)}${afeSentence(foreign)}`;
}

/** One sentence on names taken from a cell's current value or a corner label; "" when none. */
export function labelSentence(labels: PullReport["valueLabels"] = []): string {
  if (labels.length === 0) return "";
  const shown = labels.slice(0, 3).map((l) => l.key).join(", ") + (labels.length > 3 ? ", …" : "");
  return ` ${labels.length === 1 ? "One name looks" : `${labels.length} names look`} named after a cell's current value or a corner label (${shown}): see the xln output.`;
}

/** One sentence on AFE's copy of the names, for a notification; "" when the workbook has none. */
export function afeSentence(foreign: PullReport["foreignModules"] = []): string {
  const stores = foreign.filter((f) => f.kind !== "locale-sheet");
  if (stores.length === 0) return "";
  const differs = stores.reduce((n, s) => n + (s.differs?.length ?? 0), 0);
  const modules = stores.flatMap((s) => s.modules?.map((m) => m.name) ?? []);
  return ` It also carries Advanced Formula Environment (AFE) modules${modules.length ? ` (${modules.join(", ")})` : ""}, left as they are${differs ? `; ${differs} name(s) differ from AFE's text (see Problems, C14)` : ""}.`;
}

/**
 * The question before a pull that would replace source edits not built yet: the modal's
 * message and its detail (one edit per line, the first 15).
 */
export function unbuiltQuestion(workbook: string, edits: readonly UnbuiltEdit[], rewritten: readonly string[] = []): { message: string; detail: string } {
  const shown = edits.slice(0, 15).map(formatUnbuiltEdit);
  if (edits.length > 15) shown.push(`… and ${edits.length - 15} more (see the xln output)`);
  if (rewritten.length) shown.push("", ...rewrittenGroup(rewritten));
  return {
    message: `${projectFolderName(workbook)} has ${edits.length} source edit(s) not built yet. A pull writes the project as ${workbook} has it now and would replace them.`,
    detail: shown.join("\n"),
  };
}

/** The modal's group of files rewritten for layout or comments alone (the first 10). */
function rewrittenGroup(rewritten: readonly string[]): string[] {
  const shown = rewritten.slice(0, 10);
  if (rewritten.length > 10) shown.push(`… and ${rewritten.length - 10} more (see the xln output)`);
  return ["Layout or comments only: will be rewritten", ...shown];
}

/**
 * The question before a pull whose only difference from the project is layout or
 * comments (`//` lines, blank lines, line breaks): nothing unbuilt, but the text goes.
 */
export function rewrittenQuestion(workbook: string, rewritten: readonly string[]): { message: string; detail: string } {
  return {
    message: `${projectFolderName(workbook)} has no edits that are not built, but ${rewritten.length} file(s) differ in layout or comments from what the pull writes. A pull writes them as ${workbook} has them, without those comments and that layout.`,
    detail: rewrittenGroup(rewritten).join("\n"),
  };
}

/** The summary lines the output channel shows after a pull. */
export function formatPullSummary(r: PullReport, out: string, written: readonly string[], notices: readonly string[], discarded: readonly UnbuiltEdit[] = []): string[] {
  const kinds = Object.entries(r.byKind)
    .filter(([, n]) => n > 0)
    .map(([k, n]) => `${k} ${n}`)
    .join(" · ");
  const perSheet = Object.entries(r.byScope.perSheet)
    .map(([s, n]) => `${s} ${n}`)
    .join(", ");
  const lines = [
    `xln pull ${r.workbook} → ${out}`,
    `  ${r.names} names: ${r.byScope.workbook} workbook-scoped, ${r.byScope.sheet} sheet-scoped${perSheet ? ` (${perSheet})` : ""}`,
    `  by kind: ${kinds || "none"}`,
    `  modules: ${r.modules.map((m) => `${m.module} (${m.names})`).join(", ") || "none"}`,
  ];
  if (r.sheetFiles.length) lines.push(`  sheet files: ${r.sheetFiles.map((s) => `${s.sheet} (${s.names})`).join(", ")}`);
  if (r.builtIns.length) lines.push(`  built-in names (manifest only): ${r.builtIns.join(", ")}`);
  if (r.unparsedFormulas) lines.push(`  formulas the manifest could not index: ${r.unparsedFormulas}`);
  if (discarded.length) {
    lines.push(`  replaced ${discarded.length} source edit(s) not built (Discard and pull):`);
    for (const e of discarded) lines.push(`    ${formatUnbuiltEdit(e)}`);
  }
  lines.push(written.length ? `  wrote ${written.length} files: ${written.join(", ")}` : "  wrote nothing: the project already says this");
  for (const w of r.warnings) lines.push(`  warning: ${w}`);
  for (const n of r.notes ?? []) lines.push(`  note: ${n}`);
  for (const n of notices) lines.push(`  note: ${n}`);
  return lines;
}
