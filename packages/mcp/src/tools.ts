// The tools, as plain functions: arguments in, a text summary and structured data out.
// Each one calls the same function the CLI command runs (no shelling out), so an agent
// gets the CLI's verdicts and the CLI's wording. A refusal or an unreadable file is a
// ToolFailure; findings and refused plans are results, since acting on them is the point.

import { createHash } from "node:crypto";
import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { basename, dirname, extname } from "node:path";
import {
  check,
  defaultOut,
  pull,
  refusedText,
  summary as pullSummary,
  runBuild,
  buildJson,
  buildText,
  readProjectTree,
  runLibStatus,
  type BuildOutcome,
} from "@xln/cli";
import {
  CHECK_IDS,
  classify,
  isLocked,
  libStatusJson,
  placeText,
  pullProject,
  readSourceProject,
  renderLibStatus,
  scopeNeedsQuotes,
  type AuditReport,
  type CheckId,
  type Finding,
  type SourceFinding,
  type SourceName,
} from "@xln/core";
import type { Roots } from "./paths.js";

export interface ToolOutput {
  /** A few lines an agent (or a person reading the transcript) can act on. */
  text: string;
  data: Record<string, unknown>;
}

/** A tool call that did not do its job: the text is the CLI's message. */
export class ToolFailure extends Error {
  constructor(
    message: string,
    readonly data: Record<string, unknown> = {},
  ) {
    super(message);
  }
}

function plural(n: number, word: string): string {
  return `${n} ${word}${n === 1 ? "" : "s"}`;
}

function errorText(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}

/** A path the agent gave, inside the roots and existing. */
function existing(roots: Roots, path: string): string {
  const abs = roots.resolve(path);
  if (!existsSync(abs)) throw new ToolFailure(`no such file: ${path}`);
  return abs;
}

function isWorkbook(p: string): boolean {
  const e = extname(p).toLowerCase();
  return (e === ".xlsx" || e === ".xlsm") && statSync(p).isFile();
}

// ---- xln_check ---------------------------------------------------------------------------

export interface CheckArgs {
  path: string;
  only?: string[];
  severity?: "error" | "warning" | "info" | "hint";
  /** "full": the whole census and spill census too (large on big workbooks). */
  detail?: "summary" | "full";
  maxFindings?: number;
}

const SEVERITY_ORDER: Record<string, number> = { error: 0, warning: 1, info: 2, hint: 3 };

function findingLine(f: Finding): string {
  const p = placeText(f.where);
  return `${f.severity} ${f.check} ${p.prefix}${p.label}${p.suffix}: ${f.message}${f.hint ? ` (→ ${f.hint})` : ""}`;
}

function sourceLine(f: SourceFinding): string {
  return `${f.severity} ${f.file}:${f.line}:${f.column} ${f.code}: ${f.message}${f.fix ? ` [quick fix: ${f.fix.title}]` : ""}`;
}

export function checkTool(roots: Roots, args: CheckArgs): ToolOutput {
  const abs = existing(roots, args.path);
  const only = args.only?.map((x) => x.trim().toUpperCase());
  const bad = only?.filter((x) => !(CHECK_IDS as readonly string[]).includes(x)) ?? [];
  if (bad.length || only?.length === 0) throw new ToolFailure(`--only takes checks C1 to C13 (got ${(args.only ?? []).join(",")})`);
  let r: ReturnType<typeof check>;
  try {
    r = check({ workbook: abs, json: true, ...(only ? { only: only as CheckId[] } : {}), ...(args.severity ? { severity: args.severity } : {}) });
  } catch (e) {
    throw new ToolFailure(`cannot check ${args.path}: ${errorText(e)}`);
  }
  const max = args.maxFindings ?? 200;
  const errors = r.counts.error + (r.source?.counts.error ?? 0);
  const warnings = r.counts.warning + (r.source?.counts.warning ?? 0);
  const verdict = errors > 0 ? "errors" : warnings > 0 ? "warnings" : "clean";
  const byCheck = Object.fromEntries(Object.entries(r.byCheck).filter(([, c]) => c.error + c.warning + c.info > 0));
  const data: Record<string, unknown> = {
    ok: errors === 0,
    verdict,
    workbook: r.workbook,
    counts: r.counts,
    byCheck,
    checks: r.checks,
    findings: r.findings.slice(0, max),
    findingsTotal: r.findings.length,
  };
  if (r.source) data["source"] = { project: r.source.project, counts: r.source.counts, findings: r.source.findings.slice(0, max), findingsTotal: r.source.findings.length };
  if (r.config) data["config"] = r.config;
  if (args.detail === "full") {
    data["census"] = r.census;
    data["spills"] = r.spills;
  } else data["census"] = censusSummary(r);

  const lines = [`xln check ${r.workbook ?? args.path}: ${verdict === "clean" ? "no errors or warnings" : `${plural(errors, "error")}, ${plural(warnings, "warning")}`} (exit code ${errors > 0 ? 1 : 0})`];
  lines.push(`  workbook: ${r.counts.error} error, ${r.counts.warning} warning, ${r.counts.info} info${Object.keys(byCheck).length ? ` (${Object.entries(byCheck).map(([c, n]) => `${c} ${n.error + n.warning + n.info}`).join(", ")})` : ""}`);
  if (r.source) {
    const c = r.source.counts;
    lines.push(`  source ${r.source.project}: ${c.error} error, ${c.warning} warning, ${c.info} info, ${c.hint} hint (the editor's checks as you type)`);
  }
  const shown = 40;
  const all = [
    ...r.findings.map((f) => ({ rank: SEVERITY_ORDER[f.severity]!, line: findingLine(f) })),
    ...(r.source?.findings ?? []).map((f) => ({ rank: SEVERITY_ORDER[f.severity]!, line: sourceLine(f) })),
  ].sort((a, b) => a.rank - b.rank);
  for (const f of all.slice(0, shown)) lines.push(`  ${f.line}`);
  if (all.length > shown) lines.push(`  … ${all.length - shown} more in the structured result`);
  if (r.findings.length > max || (r.source?.findings.length ?? 0) > max) lines.push(`  findings cut at ${max} per list (maxFindings)`);
  const n = r.census;
  lines.push(`  census: ${plural(n.total, "name")} (${n.byScope.workbook} workbook-scoped); ${plural(r.spills.spills.length, "spill")}`);
  return { text: lines.join("\n"), data };
}

function censusSummary(r: AuditReport): Record<string, unknown> {
  const n = r.census;
  const s = r.spills;
  return {
    total: n.total,
    hidden: n.hidden,
    byKind: n.byKind,
    byScope: n.byScope,
    tiers: n.tiers,
    coordinates: n.coordinates,
    spills: { spilled: s.spills.length, singleCell: s.singleCell, bySheet: s.bySheet },
    note: 'detail: "full" adds the name families, tier names and every spill with the names over it',
  };
}

// ---- xln_names ---------------------------------------------------------------------------

export interface NamesArgs {
  path: string;
  query?: string;
  names?: string[];
  kind?: string;
  scope?: string;
  limit?: number;
  offset?: number;
}

export interface NameRow {
  name: string;
  /** `Sheet!Name` or `Name`. */
  key: string;
  /** The sheet for a sheet-scoped name; null for workbook scope. */
  scope: string | null;
  kind: string;
  /** As Excel displays it; for a named cell, the cell's formula ("" for a slot). */
  definition: string;
  doc: string | null;
  hidden: boolean;
  /** For a named cell or slot: where the name sits, `Sheet!C6` (`#` for the spill). */
  cell: string | null;
  params?: string[];
  libBase?: string;
  /** Where the name is (or a pull writes it) in the project, `names/FN.xln:12`. */
  file: string;
}

function row(n: SourceName): NameRow {
  const key = n.scope === undefined ? n.name : `${n.scope}!${n.name}`;
  let kind: string;
  let params: string[] | undefined;
  if (n.cell) kind = n.formula === "" ? "slot" : "cell";
  else {
    const c = classify(n.formula);
    kind = c.kind;
    params = c.params;
  }
  const cell = n.cell ? `${n.cell.sheet !== undefined ? `${quoteSheet(n.cell.sheet)}!` : ""}${n.cell.range}${n.cell.spill ? "#" : ""}` : null;
  return {
    name: n.name,
    key,
    scope: n.scope ?? null,
    kind,
    definition: n.formula,
    doc: n.doc ?? null,
    hidden: n.hidden,
    cell,
    ...(params ? { params } : {}),
    ...(n.libBase ? { libBase: n.libBase } : {}),
    file: `${n.file}:${n.line}`,
  };
}

function quoteSheet(s: string): string {
  return scopeNeedsQuotes(s) ? `'${s.split("'").join("''")}'` : s;
}

export function namesTool(roots: Roots, args: NamesArgs): ToolOutput {
  const abs = existing(roots, args.path);
  let files: Record<string, string>;
  let source: "workbook" | "project";
  try {
    if (statSync(abs).isDirectory()) {
      files = readProjectTree(abs);
      source = "project";
      if (!Object.keys(files).some((f) => f.startsWith("names/") && f.endsWith(".xln"))) throw new Error(`${args.path} is neither a workbook nor a project folder (no names/*.xln)`);
    } else {
      // The workbook as a pull would write it, nothing written: the same view as a project.
      files = pullProject(new Uint8Array(readFileSync(abs)), abs).files;
      source = "workbook";
    }
  } catch (e) {
    throw new ToolFailure(`cannot read ${args.path}: ${errorText(e)}`);
  }
  const project = readSourceProject(files);
  let rows = project.names.map(row);
  const total = rows.length;
  if (args.names?.length) {
    const want = new Set(args.names.map((n) => n.toLowerCase()));
    rows = rows.filter((r) => want.has(r.key.toLowerCase()) || want.has(r.name.toLowerCase()));
  }
  if (args.kind) rows = rows.filter((r) => r.kind === args.kind);
  if (args.scope !== undefined) {
    const s = args.scope.toLowerCase();
    rows = rows.filter((r) => (s === "workbook" ? r.scope === null : r.scope?.toLowerCase() === s));
  }
  if (args.query) {
    const q = args.query.toLowerCase();
    rows = rows.filter((r) => r.key.toLowerCase().includes(q) || r.definition.toLowerCase().includes(q) || (r.doc ?? "").toLowerCase().includes(q));
  }
  const matched = rows.length;
  const offset = args.offset ?? 0;
  const limit = args.limit ?? 100;
  const page = rows.slice(offset, offset + limit);
  const data: Record<string, unknown> = { source, path: abs, total, matched, offset, returned: page.length, names: page };
  if (source === "project" && project.problems.length) data["problems"] = project.problems.map((p) => ({ severity: p.severity, code: p.code, message: p.message, file: p.file, line: p.line }));
  const lines = [`${matched} of ${plural(total, "name")} in ${roots.show(abs)} (${source})${page.length < matched ? `, showing ${offset + 1}–${offset + page.length}` : ""}`];
  for (const r of page.slice(0, 30)) {
    const def = r.definition.split("\n").map((l) => l.trim()).join(" ");
    lines.push(`  ${r.key}${r.cell ? ` @${r.cell}` : ""} [${r.kind}] = ${def.length > 100 ? def.slice(0, 97) + "…" : def}`);
  }
  if (page.length > 30) lines.push(`  … ${page.length - 30} more in the structured result`);
  return { text: lines.join("\n"), data };
}

// ---- xln_pull ----------------------------------------------------------------------------

export interface PullArgs {
  workbook: string;
  out?: string;
  discard?: boolean;
}

export function pullTool(roots: Roots, args: PullArgs): ToolOutput {
  const wb = existing(roots, args.workbook);
  if (!isWorkbook(wb)) throw new ToolFailure(`${args.workbook} is not a workbook (.xlsx or .xlsm)`);
  const out = args.out !== undefined ? roots.resolve(args.out) : defaultOut(wb);
  roots.check(out);
  let o: ReturnType<typeof pull>;
  try {
    o = pull({ workbook: wb, out, json: true, discard: args.discard === true });
  } catch (e) {
    throw new ToolFailure(`cannot pull ${args.workbook}: ${errorText(e)}`);
  }
  if (o.refused) {
    // Same wording as the CLI, but the way out an agent has is the tool's argument.
    const text = refusedText(o, wb).replace("pull with --discard", "pull with discard: true (only when the user agrees to lose them)").trimEnd();
    throw new ToolFailure(text, { refused: true, out: o.out, unbuilt: o.unbuilt, rewritten: o.rewritten, notices: o.notices });
  }
  const data = { ok: true, out: o.out, files: o.written, unchanged: o.unchanged ?? [], notices: o.notices, report: o.report, discarded: o.unbuilt, rewritten: o.rewritten, provenance: o.provenance ?? [] };
  return { text: pullSummary(o).trimEnd(), data };
}

// ---- xln_build_plan and xln_build ---------------------------------------------------------

export interface BuildArgs {
  workbook: string;
  project?: string;
  out?: string;
}

function buildPaths(roots: Roots, args: BuildArgs): { wb: string; project: string; out: string | undefined } {
  const wb = existing(roots, args.workbook);
  if (!isWorkbook(wb)) throw new ToolFailure(`${args.workbook} is not a workbook (.xlsx or .xlsm)`);
  const project = args.project !== undefined ? roots.resolve(args.project) : defaultOut(wb);
  roots.check(project);
  const out = args.out !== undefined ? roots.resolve(args.out) : undefined;
  return { wb, project, out };
}

/** A short fingerprint of a change set: xln_build with `planId` refuses when the plan moved since. */
function planId(o: BuildOutcome): string {
  return createHash("sha256").update(JSON.stringify(o.result?.plan.changeSet.changes ?? [])).digest("hex").slice(0, 12);
}

function excelHasOpen(file: string): boolean {
  return existsSync(dirname(file)) && isLocked(file, readdirSync(dirname(file)));
}

function runBuildSafe(cmd: Parameters<typeof runBuild>[0]): BuildOutcome {
  try {
    return runBuild(cmd, defaultOut);
  } catch (e) {
    return { exit: 2, projectFiles: [], message: errorText(e) };
  }
}

export function buildPlanTool(roots: Roots, args: Omit<BuildArgs, "out">): ToolOutput {
  const { wb, project } = buildPaths(roots, args);
  const o = runBuildSafe({ workbook: wb, project, dryRun: true, force: false, json: true });
  const text = buildText({ workbook: wb, dryRun: true, force: false, json: false }, o).trimEnd();
  if (o.exit === 2) throw new ToolFailure(text, buildJson(o) as Record<string, unknown>);
  const json = buildJson(o) as Record<string, unknown>;
  const status = o.result?.status;
  const willWrite = o.exit === 0 && status !== "up-to-date" && (o.result?.plan.changeSet.changes.length ?? 0) > 0;
  const open = excelHasOpen(wb);
  const id = planId(o);
  const data = { ...json, planId: id, willWrite, excelOpen: open };
  const tail = [
    open ? `note: Excel has ${basename(wb)} open (~$ file): a build would be refused until it is closed` : "",
    willWrite ? `planId ${id}: show this plan to the user; xln_build with confirm: true and this planId writes exactly it` : status === "refused" ? "fix what is listed (or ask the user how), then plan again" : "",
  ].filter(Boolean);
  return { text: [text, ...tail].join("\n"), data };
}

export interface BuildWriteArgs extends BuildArgs {
  confirm: boolean;
  planId?: string;
}

export function buildTool(roots: Roots, args: BuildWriteArgs): ToolOutput {
  if (args.confirm !== true) throw new ToolFailure("xln_build writes the workbook: pass confirm: true once the user has approved the plan from xln_build_plan");
  const { wb, project, out } = buildPaths(roots, args);
  if (args.planId !== undefined) {
    const plan = runBuildSafe({ workbook: wb, project, dryRun: true, force: false, json: true });
    const now = planId(plan);
    if (plan.exit === 0 && now !== args.planId) {
      throw new ToolFailure(`xln build ${basename(wb)}: refused: the plan changed since planId ${args.planId} (now ${now}); nothing written. Run xln_build_plan again and show the user the new plan.`, { planId: now });
    }
  }
  const o = runBuildSafe({ workbook: wb, project, ...(out ? { out } : {}), dryRun: false, force: false, json: true });
  const text = buildText({ workbook: wb, dryRun: false, force: false, json: false }, o).trimEnd();
  const json = buildJson(o) as Record<string, unknown>;
  if (o.exit !== 0) throw new ToolFailure(text, json);
  return { text, data: json };
}

// ---- xln_lib_status ----------------------------------------------------------------------

export interface LibStatusArgs {
  path: string;
  lib?: string;
  diffs?: boolean;
}

export function libStatusTool(roots: Roots, args: LibStatusArgs): ToolOutput {
  const abs = existing(roots, args.path);
  const lib = args.lib !== undefined ? roots.resolve(args.lib) : undefined;
  const diffs = args.diffs ?? true;
  let r: ReturnType<typeof runLibStatus>;
  try {
    r = runLibStatus({ target: abs, ...(lib !== undefined ? { lib } : {}), json: true, diffs });
  } catch (e) {
    throw new ToolFailure(`cannot compare ${args.path} with the library: ${errorText(e)}`);
  }
  if ("error" in r) throw new ToolFailure(r.error);
  const json = libStatusJson(r.report) as Record<string, unknown>;
  return { text: renderLibStatus(r.report, { diffs }).text.trimEnd(), data: { ok: true, ...json } };
}
