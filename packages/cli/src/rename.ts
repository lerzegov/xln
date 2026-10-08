// `xln rename <project> <Old> <New>` (M5): renames a defined name in the project's source,
// explicitly (a rename in the source is otherwise a deletion plus a creation, and nothing
// renames on its own). It writes the new name with `@renamed(Old)` above it and the new
// token in every reader in the project's `.xln` files (core `renameInProject`). With the
// workbook beside the project it first plans the next build on the renamed source: a
// rename that build would refuse (a chart or a Table column reads the name, the new name
// would be captured somewhere) is refused here, nothing written, with the build's reasons.
// The workbook is only read; the next `xln build` renames the name and rewrites its token
// in the cells.

import { existsSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { basename, extname, join, resolve } from "node:path";
import { buildWorkbook, describeChange, MANIFEST_FILE, renameInProject, type BuildResult, type SourceRename } from "@xln/core";
import { projectWorkbook, readProjectFiles } from "./lib.js";
import { labelNoticesOf } from "./build.js";

export interface RenameCommand {
  /** The project folder, or the workbook (its `<name>.xln` beside it). */
  project: string;
  name: string;
  to: string;
  dryRun: boolean;
  json: boolean;
}

export interface RenameOutcome {
  /** 0 renamed (or dry run); 1 refused; 2 bad usage or unreadable. */
  exit: 0 | 1 | 2;
  project: string;
  error?: string;
  rename?: SourceRename;
  /** Project-relative paths written. */
  written: string[];
  /** The workbook the plan was made against, if beside the project. */
  workbook?: string;
  /** What the next build does with the rename (`rename Rate → Pace, rewriting it in …`). */
  build?: string[];
  /** The label notice for after the build: the cells that will still read the old name, and the Find & Replace for Excel. */
  labelNotice?: string[];
  /** The build's errors this rename brings: why it was refused. */
  refused?: string[];
}

function isWorkbookPath(p: string): boolean {
  const e = extname(p).toLowerCase();
  return e === ".xlsx" || e === ".xlsm";
}

/** The build's errors and conflicts, as lines. */
function errorsOf(r: BuildResult): string[] {
  return [
    ...r.plan.problems.filter((p) => p.severity === "error").map((p) => `${p.file ? `${p.file}${p.line ? `:${p.line}` : ""}: ` : ""}${p.message}`),
    ...r.plan.conflicts.map((c) => `${c.key}: ${c.message}`),
  ];
}

export function runRename(cmd: RenameCommand): RenameOutcome {
  const given = resolve(cmd.project);
  const project = isWorkbookPath(given) ? join(given.slice(0, given.length - extname(given).length) + ".xln") : given;
  const out: RenameOutcome = { exit: 2, project, written: [] };
  if (!existsSync(project) || !statSync(project).isDirectory()) return { ...out, error: `no such project folder: ${project}` };
  const files = readProjectFiles(project);
  let tables: string[] = [];
  try {
    const m = files[MANIFEST_FILE];
    if (m !== undefined) tables = ((JSON.parse(m) as { tables?: { name?: string }[] }).tables ?? []).map((t) => t.name ?? "").filter(Boolean);
  } catch {
    // An unreadable manifest only costs the Table-name check.
  }
  const r = renameInProject(files, cmd.name, cmd.to, { tables });
  if (typeof r === "string") return { ...out, exit: 1, error: r };
  out.rename = r;

  // The next build, planned on the renamed source: what it will do, and whether it would refuse.
  const wbPath = projectWorkbook(project, files);
  if (wbPath) {
    out.workbook = wbPath;
    const bytes = new Uint8Array(readFileSync(wbPath));
    const plan = (f: Record<string, string>) => buildWorkbook({ workbook: bytes, fileName: wbPath, files: f }, { dryRun: true });
    const before = new Set(errorsOf(plan(files)));
    const after = plan({ ...files, ...r.files });
    const renameChange = after.plan.changeSet.changes.find((c) => c.op === "rename-name" && c.to.toLowerCase() === cmd.to.trim().toLowerCase());
    out.build = renameChange ? [describeChange(renameChange)] : [];
    if (renameChange) {
      const notice = labelNoticesOf(bytes, [renameChange])[0];
      if (notice) out.labelNotice = notice;
    }
    const brought = errorsOf(after).filter((e) => !before.has(e));
    if (brought.length) {
      out.refused = brought;
      out.exit = 1;
      return out;
    }
  }
  if (!cmd.dryRun) {
    for (const [rel, text] of Object.entries(r.files)) {
      writeFileSync(join(project, ...rel.split("/")), text, "utf8");
      out.written.push(rel);
    }
  }
  out.exit = 0;
  return out;
}

export function renameText(cmd: RenameCommand, o: RenameOutcome): string {
  const lines: string[] = [];
  const head = `xln rename ${cmd.name} → ${cmd.to} in ${basename(o.project)}`;
  if (o.error) return `${head}: ${o.error}\n`;
  const r = o.rename!;
  if (o.refused) {
    lines.push(`${head}: refused, nothing written: the next build would refuse it:`);
    for (const e of o.refused.slice(0, 30)) lines.push(`  ${e}`);
    if (o.refused.length > 30) lines.push(`  … ${o.refused.length - 30} more`);
    return lines.join("\n") + "\n";
  }
  lines.push(`${head}${cmd.dryRun ? " (dry run: nothing written)" : ""}`);
  for (const w of r.warnings) lines.push(`  warning: ${w}`);
  const per = new Map<string, { name: boolean; refs: number; ann: boolean }>();
  for (const e of r.edits) {
    const p = per.get(e.path) ?? { name: false, refs: 0, ann: false };
    if (e.label === "name") p.name = true;
    else if (e.label === "reference") p.refs++;
    else p.ann = true;
    per.set(e.path, p);
  }
  for (const [path, p] of per) {
    const what = [p.name ? `the name${p.ann ? (r.annotation === "removed" ? ", @renamed removed (renamed back)" : r.annotation === "added" ? `, @renamed(${r.from.slice(r.from.lastIndexOf("!") + 1)})` : "") : ""}` : "", p.refs ? `${p.refs} reference${p.refs === 1 ? "" : "s"}` : ""].filter(Boolean).join(", ");
    lines.push(`  ${path}: ${what}`);
  }
  if (r.annotation === "kept") lines.push(`  @renamed kept: the workbook still has the name it records`);
  if (r.annotation === "none") lines.push(`  not built yet: no @renamed needed`);
  if (o.workbook) {
    if (o.build?.length) lines.push(`  the next build will ${o.build.join("; ")}`);
    else lines.push(`  the next build has no rename to make (the name is not in the workbook yet, or renamed back)`);
  } else lines.push("  no workbook beside the project: the build will check the workbook's other readers (charts, Table columns, pivots)");
  if (!cmd.dryRun) lines.push(o.written.length ? `  wrote ${o.written.join(", ")}` : "  nothing to write");
  lines.push(`  then: xln build ${o.workbook ? basename(o.workbook) : "<workbook>"}`);
  if (o.labelNotice) {
    lines.push("  after the build, in Excel (xln never writes cell values):");
    for (const l of o.labelNotice) lines.push(`    ${l}`);
  }
  return lines.join("\n") + "\n";
}

export function renameJson(o: RenameOutcome): unknown {
  return {
    ok: o.exit === 0,
    exit: o.exit,
    project: o.project,
    ...(o.error ? { error: o.error } : {}),
    ...(o.rename ? { from: o.rename.from, to: o.rename.to, annotation: o.rename.annotation, warnings: o.rename.warnings, references: o.rename.references, readers: o.rename.readers, edits: o.rename.edits } : {}),
    written: o.written,
    ...(o.workbook ? { workbook: o.workbook, build: o.build ?? [], labelNotice: o.labelNotice ?? [] } : {}),
    ...(o.refused ? { refused: o.refused } : {}),
  };
}

export const RENAME_USAGE = `       xln rename <project | workbook.xlsx> <Old> <New> [--dry-run] [--json]`;
