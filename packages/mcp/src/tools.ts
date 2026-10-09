// The tools, as plain functions: arguments in, a text summary and structured data out.
// Each one calls the same function the CLI command runs (no shelling out), so an agent
// gets the CLI's verdicts and the CLI's wording. A refusal or an unreadable file is a
// ToolFailure; findings and refused plans are results, since acting on them is the point.

import { createHash } from "node:crypto";
import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { basename, dirname, extname, join } from "node:path";
import {
  check,
  defaultOut,
  formulas,
  graphSummary,
  graphText,
  pull,
  refusedText,
  summary as pullSummary,
  runBuild,
  runVerify,
  verifyText,
  buildJson,
  buildText,
  readProjectTree,
  runLibStatus,
  runLibPublish,
  runLibTake,
  runLibBase,
  libPublishText,
  libTakeText,
  libBaseText,
  libPublishJson,
  libTakeJson,
  libBaseJson,
  libBaseRefused,
  libraryDir,
  projectWorkbook,
  readProjectFiles,
  runRename,
  renameText,
  renameJson,
  type BuildOutcome,
  type FormulasOutcome,
  type GraphSummary,
  type RenameOutcome,
} from "@xln/cli";
import {
  backupName,
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
  type FormulaViewLine,
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
    spills: { spilled: s.spills.length, singleCell: s.singleCell, ...(s.uncalculated ? { uncalculated: s.uncalculated } : {}), bySheet: s.bySheet },
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

// ---- xln_formulas ------------------------------------------------------------------------

export interface FormulasArgs {
  workbook: string;
  sheet?: string;
  order?: "appearance" | "calculation";
  /** All sheets as one list in calculation order (`xln formulas --workbook`). */
  workbookWide?: boolean;
  query?: string;
  names?: string[];
  /** "full": each line as `xln formulas --json` gives it (spans, stored text, value object). */
  detail?: "compact" | "full";
  limit?: number;
  offset?: number;
}

/** A line of the formula view, short: what an agent reads to follow the model. */
export interface FormulaRow {
  sheet: string;
  /** The cell, `C6#` for a dynamic array's anchor. */
  cell: string;
  kind: string;
  /** Saved extent of an array, spill or data table, or a shared formula's group, when more than the cell. */
  extent?: string;
  /** Names defined as this cell, its spill or its extent (the line's left-hand side). */
  defines: string[];
  formula: string;
  /** The value saved with the file, formatted short. */
  value: string | null;
  /** Defined names the formula reads (keys, each once). */
  reads: string[];
  /** Cell references as written, each once (qualified when on another sheet). */
  refs: string[];
  level?: number;
  cycle?: number;
  dependsOn?: string[];
  error?: string;
}

function formulaRow(l: FormulaViewLine): FormulaRow {
  const reads = [...new Set(l.names.map((n) => n.key ?? (n.sheet !== undefined ? `${n.sheet}!${n.id}` : n.id)))];
  const refs = [...new Set(l.refs.map((r) => (r.sheet !== l.sheet ? `${quoteSheet(r.sheet)}!${r.address}` : r.address)))];
  return {
    sheet: l.sheet,
    cell: l.cell + (l.kind === "dynamic-array" ? "#" : ""),
    kind: l.kind,
    ...(l.extent !== undefined && l.extent !== l.cell ? { extent: l.extent } : {}),
    defines: l.lhs.map((x) => x.key),
    formula: l.formula,
    value: l.valueText ?? null,
    reads,
    refs,
    ...(l.level !== undefined ? { level: l.level } : {}),
    ...(l.cycle !== undefined ? { cycle: l.cycle } : {}),
    ...(l.dependsOn?.length ? { dependsOn: l.dependsOn } : {}),
    ...(l.error ? { error: l.error } : {}),
  };
}

export function formulasTool(roots: Roots, args: FormulasArgs): ToolOutput {
  const wb = existing(roots, args.workbook);
  if (!isWorkbook(wb)) throw new ToolFailure(`${args.workbook} is not a workbook (.xlsx or .xlsm)`);
  // The CLI's argument checks, named by the tool's arguments.
  if (args.workbookWide && args.sheet !== undefined) throw new ToolFailure("workbookWide (--workbook) and sheet exclude each other");
  if (args.workbookWide && args.order === "appearance") throw new ToolFailure("workbookWide (--workbook) lists formulas in calculation order only");
  let o: FormulasOutcome;
  try {
    o = formulas({ workbook: wb, json: true, ...(args.sheet !== undefined ? { sheet: args.sheet } : {}), ...(args.order ? { order: args.order } : {}), ...(args.workbookWide ? { all: true } : {}) });
  } catch (e) {
    throw new ToolFailure(`cannot read formulas of ${args.workbook}: ${errorText(e)}`);
  }
  let lines = o.sheets.flatMap((s) => s.lines);
  const total = lines.length;
  const bySheet: Record<string, number> = {};
  for (const l of lines) bySheet[l.sheet] = (bySheet[l.sheet] ?? 0) + 1;
  if (args.names?.length) {
    const want = new Set(args.names.map((n) => n.toLowerCase()));
    const hit = (key: string | undefined, name: string) => (key !== undefined && want.has(key.toLowerCase())) || want.has(name.toLowerCase());
    lines = lines.filter((l) => l.lhs.some((x) => hit(x.key, x.name)) || l.names.some((n) => hit(n.key, n.id)));
  }
  if (args.query) {
    const q = args.query.toLowerCase();
    lines = lines.filter((l) => `${l.sheet}!${l.cell}`.toLowerCase().includes(q) || l.formula.toLowerCase().includes(q) || l.lhs.some((x) => x.key.toLowerCase().includes(q)));
  }
  const matched = lines.length;
  const offset = args.offset ?? 0;
  const limit = args.limit ?? 50;
  const page = lines.slice(offset, offset + limit);
  const rows = page.map(formulaRow);
  const data: Record<string, unknown> = {
    ok: true,
    workbook: o.workbook,
    order: o.order,
    ...(args.workbookWide ? { workbookWide: true } : {}),
    sheets: bySheet,
    total,
    matched,
    offset,
    returned: page.length,
    lines: args.detail === "full" ? page : rows,
  };
  const how = args.workbookWide ? "the whole workbook in calculation order" : `${o.order === "calculation" ? "calculation order" : "order of appearance"}, sheet by sheet`;
  const out = [`xln formulas ${o.workbook}: ${plural(total, "formula")} on ${plural(Object.keys(bySheet).length, "sheet")} (${how})`];
  if (matched !== total || page.length < matched) {
    const more = offset + page.length < matched ? ` (offset ${offset + page.length} for more)` : "";
    out.push(`  ${matched} matched${page.length < matched ? `, showing ${page.length === 0 ? "none" : `${offset + 1}–${offset + page.length}`}${more}` : ""}`);
  }
  out.push(`  per sheet: ${Object.entries(bySheet).map(([s, n]) => `${s} ${n}`).join(", ") || "none"}`);
  for (const r of rows) {
    const def = r.formula.split("\n").map((l) => l.trim()).join(" ");
    const lvl = r.level !== undefined ? `${r.level}${r.cycle !== undefined ? ` ↻${r.cycle}` : ""}  ` : "";
    const lhs = r.defines.length ? `${r.defines.join(", ")}  ` : "";
    out.push(`  ${lvl}${lhs}${quoteSheet(r.sheet)}!${r.cell}${r.extent ? ` (${r.extent})` : ""} = ${def.length > 160 ? def.slice(0, 157) + "…" : def}${r.value !== null ? `  → ${r.value}` : ""}`);
  }
  return { text: out.join("\n"), data };
}

// ---- xln_graph ---------------------------------------------------------------------------

export interface GraphArgs {
  workbook: string;
  maxItems?: number;
}

const GRAPH_LISTS = ["cycles", "recursions", "flagged", "spillRefs", "unusedNames", "usedOnlyByUnusedNames", "nameCycles"] as const;

export function graphTool(roots: Roots, args: GraphArgs): ToolOutput {
  const wb = existing(roots, args.workbook);
  if (!isWorkbook(wb)) throw new ToolFailure(`${args.workbook} is not a workbook (.xlsx or .xlsm)`);
  let s: GraphSummary;
  try {
    s = graphSummary({ workbook: wb, json: true }).summary;
  } catch (e) {
    throw new ToolFailure(`cannot read ${args.workbook}: ${errorText(e)}`);
  }
  const max = args.maxItems ?? 200;
  const totals: Record<string, number> = {};
  const cut: Record<string, unknown> = {};
  for (const k of GRAPH_LISTS) {
    totals[k] = s[k].length;
    cut[k] = s[k].slice(0, max);
  }
  const over = GRAPH_LISTS.some((k) => s[k].length > max);
  return { text: graphText(s).trimEnd() + (over ? `\n  lists cut at ${max} items in the structured result (maxItems); totals has the full counts` : ""), data: { ok: true, ...s, ...cut, totals } };
}

// ---- xln_verify --------------------------------------------------------------------------

export interface VerifyArgs {
  workbook: string;
  before?: string;
  tolerance?: number;
  maxChanges?: number;
}

export function verifyTool(roots: Roots, args: VerifyArgs): ToolOutput {
  const wb = existing(roots, args.workbook);
  if (!isWorkbook(wb)) throw new ToolFailure(`${args.workbook} is not a workbook (.xlsx or .xlsm)`);
  const before = args.before !== undefined ? existing(roots, args.before) : join(dirname(wb), backupName(basename(wb)));
  roots.check(before);
  const cmd = { workbook: wb, before, json: true, ...(args.tolerance !== undefined ? { tolerance: args.tolerance } : {}) };
  let r: ReturnType<typeof runVerify>;
  try {
    r = runVerify(cmd);
  } catch (e) {
    // The CLI's message names its flag; the tool's argument is `before`.
    throw new ToolFailure(`cannot verify ${args.workbook}: ${errorText(e).replace("pass --before <file>", "pass before: <file>")}`);
  }
  const { report, warnings } = r;
  const same = report.changed.length === 0 && report.sheetsAdded.length === 0 && report.sheetsRemoved.length === 0;
  const max = args.maxChanges ?? 200;
  const data = { ok: report.changed.length === 0, verdict: same ? "same" : "changed", before: r.before, warnings, ...report, changed: report.changed.slice(0, max), changedTotal: report.changed.length };
  return { text: verifyText(cmd, r.before, report, warnings).trimEnd().replace("(--json lists all)", "(changedTotal; the structured result lists up to maxChanges)"), data };
}

// ---- xln_rename --------------------------------------------------------------------------

export interface RenameArgs {
  path: string;
  name: string;
  to: string;
  dryRun?: boolean;
}

function lineOf(text: string, offset: number): number {
  let n = 1;
  for (let i = 0; i < offset && i < text.length; i++) if (text.charCodeAt(i) === 10) n++;
  return n;
}

export function renameTool(roots: Roots, args: RenameArgs): ToolOutput {
  const given = existing(roots, args.path);
  // A workbook names the project beside it; the workbook itself is only read (to plan the next build).
  const project = isWorkbook(given) ? defaultOut(given) : given;
  roots.check(project);
  const files = existsSync(project) && statSync(project).isDirectory() ? readProjectFiles(project) : {};
  const beside = projectWorkbook(project, files);
  if (beside !== undefined) roots.check(beside);
  const dryRun = args.dryRun === true;
  const cmd = { project, name: args.name, to: args.to, dryRun, json: true };
  let o: RenameOutcome;
  try {
    o = runRename(cmd);
  } catch (e) {
    o = { exit: 2, project, written: [], error: errorText(e) };
  }
  const text = renameText(cmd, o).trimEnd();
  const json = renameJson(o) as Record<string, unknown>;
  if (o.exit !== 0) throw new ToolFailure(text, json);
  // Each edit with its line and the text it replaces, so the list reads without the offsets.
  const edits = (o.rename?.edits ?? []).map((e) => ({ ...e, line: lineOf(files[e.path] ?? "", e.start), was: (files[e.path] ?? "").slice(e.start, e.end) }));
  const tail = dryRun ? "dry run: nothing written; call again without dryRun to write these edits" : "the workbook is unchanged until xln_build_plan and xln_build (with the user's approval)";
  return { text: `${text}\n${tail}`, data: { ...json, dryRun, edits } };
}

// ---- xln_lib_publish, xln_lib_take, xln_lib_base -----------------------------------------

/** A project folder the agent named, inside the roots. */
function projectFolder(roots: Roots, path: string): string {
  const abs = existing(roots, path);
  if (!statSync(abs).isDirectory()) throw new ToolFailure(`${path} is not a project folder (pull the workbook first: xln_pull)`);
  return abs;
}

export interface LibPublishArgs {
  project: string;
  name: string;
  lib?: string;
  dryRun?: boolean;
  confirm?: boolean;
}

export function libPublishTool(roots: Roots, args: LibPublishArgs): ToolOutput {
  const project = projectFolder(roots, args.project);
  const given = args.lib !== undefined ? roots.resolve(args.lib) : undefined;
  const dryRun = args.dryRun === true;
  // The library is shared with other workbooks: it is written only under a root, and only with consent.
  const dir = libraryDir(project, given);
  if (typeof dir !== "string") throw new ToolFailure(dir.error);
  try {
    roots.check(dir);
  } catch {
    throw new ToolFailure(`the library folder ${dir} is outside the allowed roots: xln_lib_publish writes into it, so it must lie under a --root (reading a library elsewhere stays allowed)`, { library: dir });
  }
  if (!dryRun && args.confirm !== true) throw new ToolFailure("xln_lib_publish writes into the shared library: call it with dryRun: true, show the user the diff, then pass confirm: true once they approve");
  const cmd = { project, name: args.name, lib: dir, dryRun, json: true };
  let o: ReturnType<typeof runLibPublish>;
  try {
    o = runLibPublish(cmd);
  } catch (e) {
    throw new ToolFailure(`cannot publish ${args.name}: ${errorText(e)}`);
  }
  if ("error" in o) throw new ToolFailure(o.error);
  return { text: libPublishText(cmd, o).trimEnd(), data: libPublishJson(cmd, o) };
}

export interface LibTakeArgs {
  project: string;
  name: string;
  lib?: string;
  dryRun?: boolean;
  discard?: boolean;
}

export function libTakeTool(roots: Roots, args: LibTakeArgs): ToolOutput {
  const project = projectFolder(roots, args.project);
  const lib = args.lib !== undefined ? roots.resolve(args.lib) : undefined;
  const cmd = { project, name: args.name, dryRun: args.dryRun === true, discard: args.discard === true, json: true, ...(lib !== undefined ? { lib } : {}) };
  let o: ReturnType<typeof runLibTake>;
  try {
    o = runLibTake(cmd);
  } catch (e) {
    throw new ToolFailure(`cannot take ${args.name}: ${errorText(e)}`);
  }
  if ("error" in o) throw new ToolFailure(o.error);
  const text = libTakeText(cmd, o).replace("Run again with --discard", "Call again with discard: true (only when the user agrees)").trimEnd();
  if (o.refused) throw new ToolFailure(text, libTakeJson(cmd, o));
  return { text, data: libTakeJson(cmd, o) };
}

export interface LibBaseArgs {
  project: string;
  name?: string;
  all?: boolean;
  lib?: string;
  dryRun?: boolean;
}

export function libBaseTool(roots: Roots, args: LibBaseArgs): ToolOutput {
  const project = projectFolder(roots, args.project);
  const all = args.all === true;
  if (all === (args.name !== undefined)) throw new ToolFailure("xln_lib_base takes a name or all: true (one of them)");
  const lib = args.lib !== undefined ? roots.resolve(args.lib) : undefined;
  const cmd = { project, all, dryRun: args.dryRun === true, json: true, ...(all ? {} : { name: args.name! }), ...(lib !== undefined ? { lib } : {}) };
  let o: ReturnType<typeof runLibBase>;
  try {
    o = runLibBase(cmd);
  } catch (e) {
    throw new ToolFailure(`cannot record the library base: ${errorText(e)}`);
  }
  if ("error" in o) throw new ToolFailure(o.error);
  const text = libBaseText(cmd, o).trimEnd();
  if (libBaseRefused(cmd, o)) throw new ToolFailure(text, libBaseJson(cmd, o));
  return { text, data: libBaseJson(cmd, o) };
}
