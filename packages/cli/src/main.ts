// xln command line. All the work is in @xln/core; this file reads and writes files and
// prints. `--json` output is for agents and scripts.

import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { basename, dirname, extname, join, resolve } from "node:path";
import {
  audit,
  auditOptions,
  buildGraph,
  CONFIG_FILE,
  defaultConfigText,
  parseConfig,
  type AuditOptions,
  type AuditSettings,
  CHECK_IDS,
  cellValueMap,
  renderAuditReport,
  type AuditReport,
  type AuditSeverity,
  type CheckId,
  isLocked,
  LOCK_FILE,
  NAMES_DIR,
  pullProject,
  readWorkbook,
  orderText,
  renderFormulaView,
  sheetCalcView,
  sheetFormulaView,
  workbookFormulaView,
  workbookNameIndex,
  type DependencyGraph,
  type FormulaOrder,
  type FormulaViewLine,
  type GraphNode,
  type PullReport,
  type UnbuiltEdit,
  unbuiltEdits,
  rewrittenFiles,
  formatUnbuiltEdit,
  parseLockfile,
  sourceFindings,
  spillMap,
  type SourceFinding,
  type ProvenanceStatus,
} from "@xln/core";
import { buildJson, buildText, readProjectTree, runBuild, runVerify, verifyText, type BuildCommand, type BuildOutcome, type VerifyCommand } from "./build.js";
import { applyJson, applyText, runApply, type ApplyCommand, type ApplyOutcome } from "./apply.js";
import { excelControl } from "./excel.js";
import { aroundExcel, reopenText, type ReopenReport } from "./reopen.js";
import { mainLib, projectWorkbook, readProjectFiles } from "./lib.js";
import { renameJson, renameText, runRename, type RenameCommand, type RenameOutcome } from "./rename.js";

export { runBuild, runVerify, buildJson, buildText, readProjectTree, type BuildCommand, type BuildOutcome, type VerifyCommand } from "./build.js";
// For in-process callers (the MCP server): the same functions the commands run.
export { runLibStatus, type LibStatusCommand } from "./lib.js";

export interface Io {
  out: (s: string) => void;
  err: (s: string) => void;
}

const defaultIo: Io = {
  out: (s) => process.stdout.write(s),
  err: (s) => process.stderr.write(s),
};

const USAGE = `usage: xln pull <workbook.xlsx> [--out <dir>] [--json] [--width <n>] [--discard]
       xln formulas <workbook.xlsx> [--sheet <name>] [--order appearance|calculation] [--workbook] [--json]
       xln graph <workbook.xlsx> [--json]
       xln check <workbook.xlsx | project-folder> [--json] [--only C2,C9] [--severity error|warning|info|hint] [--census-exclude <glob,...>] [--config <file>]
       xln build <workbook.xlsx> [--project <dir>] [--out <file>] [--dry-run] [--force] [--embed | --no-embed] [--no-tags] [--reopen [--discard]] [--json]
       xln apply <workbook.xlsx> <changes.json> [--out <file>] [--dry-run] [--reopen [--discard]] [--json]
       xln verify <workbook.xlsx> [--before <file>] [--tolerance <x>] [--json]
       xln rename <project | workbook.xlsx> <Old> <New> [--dry-run] [--json]
       xln lib status <workbook.xlsx | project> [--lib <dir>] [--json] [--no-diff]
       xln lib publish <project> <Name> [--lib <dir>] [--dry-run] [--json]
       xln lib take <project> <Name> [--lib <dir>] [--dry-run] [--discard] [--json]
       xln lib base <project> <Name | --all> [--lib <dir>] [--dry-run] [--json]
       xln --help

  pull      Read the workbook's defined names and write a project:
            names/*.xln, names/sheets/*.xln, workbook.manifest.json, xln.lock.json,
            and xln.config.json (audit settings) when the project has none yet.
            Default folder: <workbook>.xln next to the workbook. The workbook is only read.
            Every pull is fresh: names/**, the lockfile and the manifest are written as
            the workbook has them now (other .xln files there are removed; names go to
            their module's or sheet's file); xln.config.json is kept. A pull refuses
            (exit 1, nothing written) while the project has source edits not built yet,
            and lists them: build them first, or pull with --discard.
    --out     folder to write the project into
    --json    print the report as JSON instead of a summary
    --width   line width above which definitions are pretty-printed (default 100)
    --discard replace the project even though it has source edits not built (they are lost)
    --fresh, --live  accepted and ignored: every pull is fresh

  formulas  Print the cell formulas of each sheet in order of appearance (row by row):
            the names defined as the cell (or its spill), address, kind, formula as Excel
            shows it, saved value. The workbook is only read.
    --sheet   only this sheet
    --order   calculation: each formula after what it reads (inputs first), with its
              level and circular references marked (default: appearance)
    --workbook  all sheets in one list, in calculation order (dependencies cross sheets)
    --json    print the lines as JSON (names on the cell as lhs; names used, with their
              spans in the formula; in calculation order also level, cycle, dependsOn)

  graph     Summarise the cell dependency graph: nodes by kind, edges, circular
            references, references that cannot be followed (INDIRECT, computed OFFSET,
            other workbooks, #REF!), fixed references into a spill (C9), unused names (C10)
            and name cycles (C12). The workbook is only read.
    --json    print the summary as JSON

  check     Audit the workbook (checks C1-C15): syntax, prefix health, built-in
            collisions, unresolved references, unqualified sheet-scoped reads, arity,
            Excel limits, fixed references into spills, unused names, copy drift,
            cycles, constants in LAMBDAs, and the copy of the names that Microsoft's
            Advanced Formula Environment (AFE, Excel Labs) keeps in the workbook: its
            modules, and the names where its text and the workbook differ (C14);
            names whose label cell no longer gives them (C15); then
            the name census (C8) and the spill census (C9). With its project (<workbook>.xln beside it, or the project
            folder given in place of the workbook), also the checks the editor runs as
            you type, on the names files at their file:line, with the quick fix the
            editor offers. Nothing is written. Exit code 0: no errors; 1: errors found
            (in the workbook or the source); 2: bad usage, or the workbook cannot be read.
    --json            the report as JSON (findings, census, spills, source), for agents
    --only            run only these checks (comma-separated, e.g. C2,C9)
    --severity        report only findings at least this severe: error, warning,
                      info or hint (default: all, the source's hints included; hint
                      lists everything too)
    --census-exclude  the check harness, replacing the config's: names C10 does not
                      report and the C8 tiers leave out (glob on Sheet!Name or
                      Module.Name, e.g. 'Check!*,CHK.*')
    --config          audit settings to use (default: <workbook>.xln/xln.config.json
                      when it exists); flags win over the file

  build     Write the project's names and cell formulas into the workbook. Compares the
            source with the last pull (xln.lock.json) and the workbook, per name and per
            cell statement (Name @C6 = …; @C5 = …;): what changed only in source is
            written; what changed in Excel is kept; a change on both sides is a
            conflict and nothing is written. A named cell follows its name when Excel
            moved it; an unnamed one is its address, which is read-only. What a name
            covers is the address's: Name @C6# is the spill, Name @C6 the cell alone
            (add or remove the #; nothing moves a name on its own). Filling a slot
            (Name @C6 = ;) writes the formula into the empty cell. The result is
            read back before and after writing. Refuses while Excel has the workbook
            open (~$ file). The previous file is kept as <workbook>.backup.xlsx. A
            rename (xln rename, or @renamed(OldName) before the new name) also rewrites
            the name in the workbook's cell formulas, conditional formats, validations
            and other names, token by token; once written, the build removes the
            @renamed line from the source (the one source edit a build makes).
            Exit codes: 0 built or up to date; 1 refused (source errors, conflicts,
            names still used by cells); 2 bad usage or unreadable; 3 Excel has the file
            open; 4 read-back failed (nothing written, or the original restored).
    --project  project folder (default <workbook>.xln)
    --out      write the built workbook to this file instead (the lockfile is not updated)
    --dry-run  print the change set without writing
    --force    write even when nothing changed (sets fullCalcOnLoad)
    --embed    embed the project source in the workbook: names/**/*.xln, the lockfile
               and the config in a custom XML part, an archive copy for a workbook
               handed on (pull does not read it). Off by default;
               "build": {"embed": true} in xln.config.json turns it on for the
               project. A build without it leaves an existing part as it is
    --no-embed do not embed, even when xln.config.json says so
    --no-tags  do not tag module names' comments with their provenance
               ([xln FN 1.2 #3f9a1c]: module, @version from the module file, source hash)
    --reopen   desktop Excel (macOS, Windows): close the workbook in Excel without
               saving, build, open it again. Refused (exit 5) while Excel has unsaved
               changes in it; exit 6 when Excel does not reopen it (a repair prompt).
    --discard  with --reopen: close it even with unsaved changes (they are lost)
    --json     the change set, conflicts, problems and read-back as JSON

  apply     Apply a change set (JSON: the changeSet that build --json prints, or a list
            of changes) to the workbook, without a project. Cell changes are written as
            probe F8 found Excel accepts them: dynamic-array form, old spill areas
            emptied, shared groups un-shared, slots inserted, calcChain.xml dropped.
            Same safety as build: read back before and after writing, refused while
            Excel has the file open, the previous file kept as <workbook>.backup.xlsx.
            Exit codes: 0 applied (or dry run); 2 bad usage or input; 3 Excel has the file
            open; 4 the changes do not apply or did not read back; 5, 6 as build --reopen.
    --out, --dry-run, --reopen, --discard, --json   as for build

  verify    After opening the built workbook in Excel and saving it, compare its cached
            values cell by cell with the copy before the build (default the backup
            <workbook>.backup.xlsx). Exit code 0: no cell changed; 1: some changed;
            2: unreadable.
    --before     the copy before the build
    --tolerance  relative difference under which numbers count as equal (default 0)
    --json       the changed cells as JSON

  rename    Rename a defined name in the project's source (Old: Name, or Sheet!Name for
            a sheet's local name): its statement gets the new name and @renamed(Old)
            above it (the record the build reads), and every formula of the project's
            .xln files that reads it gets the new name (strings, LET/LAMBDA variables
            and other scopes' names of that spelling untouched). Then xln build renames
            the name in the workbook and rewrites it in the cells, token by token,
            and removes the @renamed line.
            Refused (exit 1, nothing written) when the new name is not valid or taken,
            when a formula would then read something else (a LET/LAMBDA variable, a
            sheet's local name), or when the next build would refuse it: a chart, a
            Table column, a pivot table, a hyperlink or a form control reads the name
            (rename those in Excel's Name Manager instead, then pull). The workbook is
            only read. Exit 0 renamed (or dry run), 1 refused, 2 bad usage.
    --dry-run  print what would change, write nothing
    --json     the edits and the build's view as JSON

  lib status  Compare the library (a folder of .lambda files) with a workbook's names, or
            with a project's source (and the workbook beside it, if there: what is not
            built yet). Three-way on the copy's library base, the library version it
            came from (@from(lib #…) in the source, lib#… in the workbook's [xln …]
            tag): identical (same definition, layout and number spelling aside);
            outdated (the copy is its base, the library moved); modified (edited here,
            the library is still the base); both changed (each side moved: the diffs
            from the base, whose text the project keeps in library-bases/, else from
            the workbook or its backup); differs
            (no base recorded); missing. Plus local only: LAMBDAs of a library module
            (FN.*) the library does not have; other modules' LAMBDAs are counted on one
            line. Diffs: - the copy, + the library. Nothing is written. Exit 0, or 2.
    --lib      the library folder; default "library" in the project's xln.config.json
               (relative to the project folder; ~ is the home folder)
    --json     the report as JSON
    --no-diff  leave the diffs out
  lib publish  Write a project LAMBDA to the library: a new <Name>.lambda with a generated
            header (# name, # summary: the doc comment's first sentence, # params), or
            the existing file with its header fields and rationale kept and the
            definition (and summary, params, @param descriptions, when they changed)
            replaced. Prints the file's diff. The project's entry then records the
            published version as its base, @from(lib #…) (build to carry it into the
            workbook).
    --dry-run  print the diff, write nothing
  lib take  Take the library's version of a function into the project's source: its
            definition and doc comment, and @from(lib #…). Refused (exit 1) when the
            copy was edited since its base (modified, both changed) or has no base
            (differs), unless --discard: the edit is lost. Build to write it.
    --dry-run  print the change, write nothing
    --discard  take it even though it discards the copy's own edit
  lib base  Record the library base of a function identical to the library that records
            none (inserted or published before @from existed): writes @from(lib #…) on
            its entry, as Insert, Take and Publish do; --all: every such entry. Prints
            what it wrote; a function not identical, or with a base, is skipped with the
            reason (exit 1 when the one named is). Nothing else does it: lib status says
            "no base recorded" on these.
    --dry-run  print what it would record, write nothing
    --all      every identical function without a base
  Publish, take and base keep the base's text in the project's library-bases/<hash>.json
  (for the three-way diff); a pull leaves that folder alone and nothing prunes it.
`;

export interface PullCommand {
  workbook: string;
  out?: string;
  json: boolean;
  width?: number;
  /** `--discard`: replace the project even when it holds source edits not built yet. */
  discard?: boolean;
  /** `--fresh` or `--live` was given: accepted, no longer needed (every pull is fresh). */
  freshFlag?: string;
}

export interface PullOutcome {
  out: string;
  written: string[];
  /** Messages about the surroundings: Excel has the file open, files removed, a deprecated flag. */
  notices: string[];
  /** Absent when the pull refused. */
  report?: PullReport;
  /** Files of the project left as they were (the workbook says the same). */
  unchanged?: string[];
  /** D6: tagged names and whether they still match their tag. */
  provenance?: ProvenanceStatus[];
  /** Source edits not built yet: the pull refused to replace them (unless `--discard`, then they are the ones given up). */
  unbuilt: UnbuiltEdit[];
  /** Names files that differ from the pull's in layout or comments alone (M3e): rewritten, or they would be. */
  rewritten: string[];
  /** True when the pull refused (unbuilt edits and no `--discard`): nothing was written. */
  refused?: boolean;
}

/** Default project folder: `lbo.xlsx` → `lbo.xln` beside it. */
export function defaultOut(workbook: string): string {
  const ext = extname(workbook);
  return join(dirname(workbook), basename(workbook, ext) + ".xln");
}

export function pull(cmd: PullCommand): PullOutcome {
  const workbook = resolve(cmd.workbook);
  const out = resolve(cmd.out ?? defaultOut(workbook));
  const notices: string[] = [];
  if (isLocked(workbook, readdirSync(dirname(workbook)))) {
    notices.push(`Excel has ${basename(workbook)} open (~$ file present): unsaved changes are not in the file and are not pulled.`);
  }
  if (cmd.freshFlag) notices.push(`${cmd.freshFlag} is no longer needed: every pull is fresh (it stays accepted)`);
  const bytes = new Uint8Array(readFileSync(workbook));
  // Every pull is fresh (2026-10-06): first, what it would replace that was never built.
  const project = readProject(out);
  const unbuilt = project ? unbuiltEdits({ workbook: bytes, fileName: workbook, files: project }) : [];
  const { files, report, provenance } = pullProject(bytes, workbook, cmd.width !== undefined ? { width: cmd.width } : {});
  // Besides the edits: files rewritten for their layout or `//` comments alone (M3e). A note, not a refusal.
  const rewritten = project ? rewrittenFiles(project, files, unbuilt) : [];
  if (rewritten.length) notices.push(`layout or comments only (a pull rewrites them as the workbook has them): ${rewritten.join(", ")}`);
  if (unbuilt.length > 0 && !cmd.discard) return { out, written: [], notices, unbuilt, rewritten, refused: true };

  // names/** is replaced (with the lockfile and the manifest): files the pull does not write
  // go. Kept: xln.config.json and library-bases/ (the library bases' texts, written only by
  // the library actions).
  const stale: string[] = [];
  const walk = (dir: string, rel: string, depth: number): void => {
    if (depth > 6 || !existsSync(dir)) return;
    for (const e of readdirSync(dir, { withFileTypes: true })) {
      const r = `${rel}/${e.name}`;
      if (e.isDirectory()) walk(join(dir, e.name), r, depth + 1);
      else if (e.name.endsWith(".xln") && !(r in files)) stale.push(r);
    }
  };
  walk(join(out, NAMES_DIR), NAMES_DIR, 0);
  for (const rel of stale) rmSync(join(out, ...rel.split("/")));
  if (stale.length) notices.push(`removed ${stale.join(", ")} (not written by this pull)`);
  if (unbuilt.length) notices.push(`--discard: ${plural(unbuilt.length, "source edit")} not built replaced by the workbook's version`);
  // The audit settings are the author's: written once, never replaced.
  if (existsSync(join(out, CONFIG_FILE))) delete files[CONFIG_FILE];
  else files[CONFIG_FILE] ??= defaultConfigText();
  const written: string[] = [];
  const unchanged: string[] = [];
  for (const [rel, text] of Object.entries(files)) {
    const path = join(out, ...rel.split("/"));
    if (existsSync(path) && readFileSync(path, "utf8") === text) {
      unchanged.push(rel);
      continue;
    }
    mkdirSync(dirname(path), { recursive: true });
    writeFileSync(path, text, "utf8");
    written.push(rel);
  }
  return { out, written, unchanged, notices, report, provenance, unbuilt, rewritten };
}

/** The project folder's names files and lockfile (path → text), or undefined when it has no names files. */
function readProject(out: string): Record<string, string> | undefined {
  const names = join(out, NAMES_DIR);
  if (!existsSync(names)) return undefined;
  const files: Record<string, string> = {};
  const walk = (dir: string, rel: string, depth: number): void => {
    if (depth > 6) return;
    for (const e of readdirSync(dir, { withFileTypes: true })) {
      const p = join(dir, e.name);
      const r = `${rel}/${e.name}`;
      if (e.isDirectory()) walk(p, r, depth + 1);
      else if (e.name.endsWith(".xln")) files[r] = readFileSync(p, "utf8");
    }
  };
  walk(names, NAMES_DIR, 0);
  if (Object.keys(files).length === 0) return undefined;
  const lock = join(out, LOCK_FILE);
  if (existsSync(lock)) files[LOCK_FILE] = readFileSync(lock, "utf8");
  return files;
}

/** Why a pull refused: the source edits it would replace, and what to do. */
export function refusedText(o: PullOutcome, workbook: string): string {
  const lines = [`xln pull ${basename(workbook)}: refused: the project has ${plural(o.unbuilt.length, "source edit")} not built yet, which a pull would replace:`];
  for (const e of o.unbuilt.slice(0, 40)) lines.push(`  ${formatUnbuiltEdit(e)}`);
  if (o.unbuilt.length > 40) lines.push(`  … ${o.unbuilt.length - 40} more (--json lists all)`);
  lines.push("Build them first (xln build), then pull; or pull with --discard to replace them with the workbook's version.");
  for (const n of o.notices) lines.push(`note: ${n}`);
  return lines.join("\n") + "\n";
}

function plural(n: number, word: string): string {
  return `${n} ${word}${n === 1 ? "" : "s"}`;
}

export function summary(o: PullOutcome): string {
  const r = o.report!;
  const lines: string[] = [];
  lines.push(`xln pull ${r.workbook} → ${o.out}`);
  const per = Object.entries(r.byScope.perSheet).map(([s, n]) => `${s} ${n}`).join(", ");
  lines.push(`  ${plural(r.names, "name")}: ${r.byScope.workbook} workbook-scoped, ${r.byScope.sheet} sheet-scoped${per ? ` (${per})` : ""}`);
  const kinds = Object.entries(r.byKind).filter(([, n]) => n > 0).map(([k, n]) => `${k} ${n}`).join(" · ");
  lines.push(`  by kind: ${kinds || "none"}`);
  lines.push(`  modules: ${r.modules.map((m) => `${m.module} (${m.names})`).join(" · ") || "none"}`);
  if (r.sheetFiles.length) lines.push(`  sheet files: ${r.sheetFiles.map((f) => `${f.sheet} (${f.names})`).join(" · ")}`);
  const c = r.cells;
  if (c.named + c.slots + c.unnamed > 0) {
    lines.push(`  cell statements: ${c.named} named · ${c.slots} slot${c.slots === 1 ? "" : "s"} · ${c.unnamed} unnamed (${plural(c.blocks, "block")} over ${c.blockCells} cells)`);
  }
  const extra = [
    r.hidden ? `${r.hidden} hidden` : "",
    r.builtIns.length ? `built-in names in manifest: ${r.builtIns.join(", ")}` : "",
    r.helpersIgnored.length ? `${plural(r.helpersIgnored.length, "_xl helper name")} ignored` : "",
    r.unparsedFormulas ? `${plural(r.unparsedFormulas, "cell formula")} not indexed (does not parse)` : "",
  ].filter(Boolean);
  if (extra.length) lines.push(`  ${extra.join(" · ")}`);
  const edited = (o.provenance ?? []).filter((p) => p.state === "edited");
  if (o.provenance?.length) lines.push(`  provenance tags: ${o.provenance.length}, ${edited.length} edited in Excel since the build${edited.length ? ` (${edited.slice(0, 10).map((p) => p.key).join(", ")}${edited.length > 10 ? ", …" : ""})` : ""}`);
  lines.push(o.written.length ? `  wrote ${plural(o.written.length, "file")}: ${o.written.join(", ")}` : "  wrote nothing: the project already says this");
  if (o.written.length && o.unchanged?.length) lines.push(`  ${plural(o.unchanged.length, "file")} unchanged`);
  for (const n of r.notes ?? []) lines.push(`  note: ${n}`);
  const warn = [...o.notices, ...r.warnings];
  if (warn.length) {
    lines.push(`  ${plural(warn.length, "warning")}:`);
    for (const w of warn.slice(0, 20)) lines.push(`    ${w}`);
    if (warn.length > 20) lines.push(`    … ${warn.length - 20} more (--json lists all)`);
  }
  return lines.join("\n") + "\n";
}

export interface FormulasCommand {
  workbook: string;
  sheet?: string;
  json: boolean;
  /** Default appearance; `--workbook` implies calculation. */
  order?: FormulaOrder;
  /** All sheets in one list (calculation order). */
  all?: boolean;
}

export interface GraphCommand {
  workbook: string;
  json: boolean;
}

export interface CheckCommand {
  workbook: string;
  json: boolean;
  only?: CheckId[];
  /** The least severity listed; `hint` lists everything (the audit has no hints: it lists its info findings). */
  severity?: AuditSeverity | "hint";
  censusExclude?: string[];
  /** Audit settings file; default `<workbook>.xln/xln.config.json` when it exists. */
  config?: string;
}

type Command = ({ cmd: "pull" } & PullCommand) | ({ cmd: "formulas" } & FormulasCommand) | ({ cmd: "graph" } & GraphCommand) | ({ cmd: "check" } & CheckCommand);

function parseArgs(args: string[]): Command | string {
  const [cmd, ...rest] = args;
  if (cmd !== "pull" && cmd !== "formulas" && cmd !== "graph" && cmd !== "check") return cmd === undefined || cmd === "-h" || cmd === "--help" ? "" : `unknown command '${cmd}'`;
  const c: { workbook: string; json: boolean; out?: string; width?: number; sheet?: string; order?: FormulaOrder; all?: boolean; discard?: boolean; freshFlag?: string; only?: CheckId[]; severity?: AuditSeverity | "hint"; censusExclude?: string[]; config?: string } = { workbook: "", json: false };
  const valued = cmd === "pull" ? ["--out", "--width"] : cmd === "formulas" ? ["--sheet", "--order"] : cmd === "check" ? ["--only", "--severity", "--census-exclude", "--config"] : [];
  for (let k = 0; k < rest.length; k++) {
    const a = rest[k]!;
    if (a === "--json") c.json = true;
    else if (a === "--workbook" && cmd === "formulas") c.all = true;
    else if ((a === "--fresh" || a === "--live") && cmd === "pull") c.freshFlag = a;
    else if (a === "--discard" && cmd === "pull") c.discard = true;
    else if (valued.includes(a)) {
      const v = rest[++k];
      if (v === undefined) return `${a} needs a value`;
      if (a === "--out") c.out = v;
      else if (a === "--only") {
        const ids = v.split(",").map((x) => x.trim().toUpperCase()).filter(Boolean);
        const bad = ids.filter((x) => !(CHECK_IDS as readonly string[]).includes(x));
        if (bad.length || ids.length === 0) return `--only takes checks C1 to ${CHECK_IDS[CHECK_IDS.length - 1]} (got ${v})`;
        c.only = ids as CheckId[];
      } else if (a === "--severity") {
        if (v !== "error" && v !== "warning" && v !== "info" && v !== "hint") return "--severity must be error, warning, info or hint";
        c.severity = v;
      } else if (a === "--census-exclude") c.censusExclude = v.split(",").map((x) => x.trim()).filter(Boolean);
      else if (a === "--config") c.config = v;
      else if (a === "--sheet") c.sheet = v;
      else if (a === "--order") {
        if (v !== "appearance" && v !== "calculation") return "--order must be appearance or calculation";
        c.order = v;
      } else {
        const w = Number(v);
        if (!Number.isInteger(w) || w < 20) return "--width must be an integer of at least 20";
        c.width = w;
      }
    } else if (a.startsWith("-")) return `unknown option '${a}'`;
    else if (c.workbook === "") c.workbook = a;
    else return `unexpected argument '${a}'`;
  }
  if (c.workbook === "") return `${cmd} needs a workbook`;
  if (c.all && c.sheet !== undefined) return "--workbook and --sheet exclude each other";
  if (c.all && c.order === "appearance") return "--workbook lists formulas in calculation order only";
  return { cmd, ...c };
}

export interface FormulasOutcome {
  workbook: string;
  order: FormulaOrder;
  /** One entry per sheet; with `--workbook`, one entry for all (`sheet` is undefined). */
  sheets: { sheet: string | undefined; lines: FormulaViewLine[]; text: string }[];
}

/** The formula view of every sheet (or one, or the whole workbook): lines and rendered text. */
export function formulas(cmd: FormulasCommand): FormulasOutcome {
  const name = basename(cmd.workbook);
  const wb = readWorkbook(new Uint8Array(readFileSync(cmd.workbook)));
  const order: FormulaOrder = cmd.all ? "calculation" : (cmd.order ?? "appearance");
  const index = workbookNameIndex(wb);
  const cache = new Map<string, unknown>();
  const graph = order === "calculation" ? buildGraph(wb, index, { cache }) : undefined;
  if (cmd.all) {
    const lines = workbookFormulaView(wb, index, { graph: graph! });
    return { workbook: name, order, sheets: [{ sheet: undefined, lines, text: renderFormulaView(lines, { sheet: "", workbook: name, order: orderText(order), levels: true, sheets: true }).text }] };
  }
  let sheets = wb.sheets.filter((s) => s.kind === "worksheet" || s.kind === "macrosheet" || s.formulas.length > 0);
  if (cmd.sheet !== undefined) {
    const want = cmd.sheet.toLowerCase();
    sheets = wb.sheets.filter((s) => s.name.toLowerCase() === want);
    if (sheets.length === 0) throw new Error(`no sheet '${cmd.sheet}'; the workbook has ${wb.sheets.map((s) => s.name).join(", ")}`);
  }
  return {
    workbook: name,
    order,
    sheets: sheets.map((s) => {
      const lines = graph ? sheetCalcView(wb, s.name, index, { graph }) : sheetFormulaView(wb, s.name, index, { cache });
      const opts = graph ? { order: orderText(order), levels: true } : {};
      return { sheet: s.name, lines, text: renderFormulaView(lines, { sheet: s.name, workbook: name, ...opts }).text };
    }),
  };
}

export interface GraphSummary {
  workbook: string;
  nodes: { formulas: number; inputs: number; names: number };
  edges: number;
  /** Circular references, members in sheet order. */
  cycles: { id: number; members: string[] }[];
  /** LAMBDA names calling themselves or each other (allowed). */
  recursions: string[][];
  /** References that cannot be followed from the file alone. */
  flagged: { node: string; kind: string; reason: string; text: string }[];
  /** C9: fixed references into a spill. */
  spillRefs: { at: string; ref: string; use: string; fit: "exact" | "part" | "beyond" }[];
  /** C10. */
  unusedNames: string[];
  usedOnlyByUnusedNames: string[];
  /** C12 (recursive: LAMBDAs only). */
  nameCycles: { members: string[]; recursive: boolean }[];
  /** Longest chain of formulas. */
  maxLevel: number;
  buildMs: number;
}

/** The dependency graph of a workbook, summarised. */
export function graphSummary(cmd: GraphCommand): { summary: GraphSummary; graph: DependencyGraph } {
  const wb = readWorkbook(new Uint8Array(readFileSync(cmd.workbook)));
  const t0 = performance.now();
  const g = buildGraph(wb);
  const buildMs = performance.now() - t0;
  const s = g.stats();
  const label = (n: GraphNode) => n.label;
  const u = g.unusedNames();
  const summary: GraphSummary = {
    workbook: basename(cmd.workbook),
    nodes: { formulas: s.formulas, inputs: s.inputs, names: s.names },
    edges: s.edges,
    cycles: g.cycles.map((c) => ({ id: c.id, members: c.members.map(label) })),
    recursions: g.recursions.map((c) => c.members.map(label)),
    flagged: g.flagged().flatMap((n) => n.flags.map((f) => ({ node: n.label, kind: f.kind, reason: f.reason, text: f.text }))),
    spillRefs: g.spillRefs.map((f) => ({ at: f.node.label, ref: f.ref, use: f.use, fit: f.fit })),
    unusedNames: u.unused.map(label),
    usedOnlyByUnusedNames: u.onlyByUnused.map(label),
    nameCycles: g.nameCycles().map((c) => ({ members: c.members.map(label), recursive: c.recursive })),
    maxLevel: Math.max(0, ...g.nodes.map((n) => n.level)),
    buildMs: Math.round(buildMs * 10) / 10,
  };
  return { summary, graph: g };
}

function list(items: string[], max = 12): string {
  return items.length <= max ? items.join(", ") : `${items.slice(0, max).join(", ")}, … (${items.length - max} more)`;
}

const FIT: Record<GraphSummary["spillRefs"][number]["fit"], string> = { exact: "", part: " (the reference covers part of the spill)", beyond: " (the reference also reaches cells outside the spill)" };

export function graphText(s: GraphSummary): string {
  const out: string[] = [];
  out.push(`xln graph ${s.workbook}`);
  out.push(`  nodes: ${plural(s.nodes.formulas, "formula")} · ${plural(s.nodes.inputs, "input range")} · ${plural(s.nodes.names, "name")}; ${plural(s.edges, "edge")}; longest chain ${s.maxLevel} (built in ${s.buildMs} ms)`);
  if (s.cycles.length === 0) out.push("  circular references: none");
  else {
    out.push(`  circular references: ${s.cycles.length}`);
    for (const c of s.cycles) out.push(`    ↻${c.id} (${c.members.length}): ${list(c.members, 100)}`);
  }
  if (s.recursions.length) out.push(`  recursive LAMBDAs (allowed): ${s.recursions.map((r) => r.join(" ↔ ")).join("; ")}`);
  const by = (k: string) => s.flagged.filter((f) => f.kind === k);
  out.push(`  not followed: ${by("dynamic").length} dynamic, ${by("external").length} external, ${by("broken").length} broken`);
  for (const f of s.flagged.slice(0, 30)) out.push(`    ${f.kind} ${f.node}: ${f.reason}${f.text && !f.reason.includes(f.text) ? ` (${f.text})` : ""}`);
  if (s.flagged.length > 30) out.push(`    … ${s.flagged.length - 30} more (--json lists all)`);
  out.push(`  C9 fixed references into a spill: ${s.spillRefs.length === 0 ? "none" : s.spillRefs.length}`);
  for (const f of s.spillRefs.slice(0, 30)) out.push(`    ${f.at}: ${f.ref} → use ${f.use}${FIT[f.fit]}`);
  if (s.spillRefs.length > 30) out.push(`    … ${s.spillRefs.length - 30} more (--json lists all)`);
  out.push(`  C10 unused names: ${s.unusedNames.length === 0 ? "none" : `${s.unusedNames.length}: ${list(s.unusedNames, 40)}`}`);
  if (s.usedOnlyByUnusedNames.length) out.push(`      used only by unused names: ${list(s.usedOnlyByUnusedNames, 40)}`);
  const bad = s.nameCycles.filter((c) => !c.recursive);
  out.push(`  C12 name cycles: ${bad.length === 0 ? "none" : bad.map((c) => c.members.join(" → ")).join("; ")}`);
  return out.join("\n") + "\n";
}

function runGraph(cmd: GraphCommand, io: Io): number {
  let s: GraphSummary;
  try {
    s = graphSummary(cmd).summary;
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    if (cmd.json) io.out(JSON.stringify({ ok: false, error: msg }, null, 2) + "\n");
    else io.err(`xln: cannot read ${cmd.workbook}: ${msg}\n`);
    return 1;
  }
  io.out(cmd.json ? JSON.stringify({ ok: true, ...s }, null, 2) + "\n" : graphText(s));
  return 0;
}

export interface LoadedConfig {
  path: string;
  settings: AuditSettings | undefined;
  /** What in the file could not be used (left out). */
  problems: string[];
  /** For information: a setting that no longer exists, ignored. */
  notes: string[];
}

/** The project's audit settings: `--config`, or `<workbook>.xln/xln.config.json` (or the project's given) when it exists. */
export function loadConfig(workbook: string, config?: string, project?: string): LoadedConfig | undefined {
  const path = resolve(config ?? join(project ?? defaultOut(resolve(workbook)), CONFIG_FILE));
  if (!existsSync(path)) {
    if (config !== undefined) throw new Error(`no such config file: ${config}`);
    return undefined;
  }
  const { config: parsed, problems, notes } = parseConfig(readFileSync(path, "utf8"));
  return { path, settings: parsed.audit, problems, notes };
}

/** Audit options: the project's settings, then the flags on top. */
export function checkOptions(cmd: CheckCommand, settings: AuditSettings | undefined): AuditOptions {
  const flags: AuditOptions = {
    workbook: basename(cmd.workbook),
    ...(cmd.only ? { only: cmd.only } : {}),
    ...(cmd.severity ? { minSeverity: cmd.severity === "hint" ? "info" : cmd.severity } : {}),
    ...(cmd.censusExclude ? { harness: cmd.censusExclude } : {}),
  };
  return auditOptions(settings, flags);
}

/** What the source checker found in a project's files (the editor's live checks). */
export interface SourceCheck {
  /** The project folder. */
  project: string;
  counts: Record<SourceFinding["severity"], number>;
  findings: SourceFinding[];
}

const SOURCE_RANK: Record<SourceFinding["severity"], number> = { error: 3, warning: 2, info: 1, hint: 0 };

/**
 * The audit of a workbook (C1–C15, census sections), with the settings it ran with; and,
 * when the workbook has its project (`<workbook>.xln`, or the project folder given in its
 * place), the source checker's findings on the project's files: the same the editor shows
 * as you type, so the command line and the editor agree (feedback 2026-10-07).
 */
export function check(cmd: CheckCommand): AuditReport & { config?: { path: string; problems: string[]; notes?: string[] }; source?: SourceCheck } {
  const given = resolve(cmd.workbook);
  let workbook = given;
  let project: string | undefined;
  if (existsSync(given) && statSync(given).isDirectory()) {
    project = given;
    const wbPath = projectWorkbook(project, readProjectFiles(project));
    if (!wbPath) throw new Error(`${cmd.workbook} is not a project folder with its workbook beside it (the manifest names the workbook)`);
    workbook = wbPath;
  } else {
    const p = defaultOut(given);
    if (existsSync(join(p, NAMES_DIR))) project = p;
  }
  const loaded = loadConfig(workbook, cmd.config, project);
  const bytes = new Uint8Array(readFileSync(workbook));
  const wb = readWorkbook(bytes);
  const r = audit(wb, { ...checkOptions({ ...cmd, workbook }, loaded?.settings), ...cellValuesOf(bytes) });
  const out: ReturnType<typeof check> = { ...r };
  if (loaded) out.config = { path: loaded.path, problems: loaded.problems, ...(loaded.notes.length ? { notes: loaded.notes } : {}) };
  if (project) {
    const files = readProjectTree(project);
    let lock;
    try {
      lock = files[LOCK_FILE] !== undefined ? parseLockfile(files[LOCK_FILE]) : undefined;
    } catch {
      // Not a lockfile: the checker runs without the last pull (the build reports it).
    }
    const min = cmd.severity === undefined ? 0 : SOURCE_RANK[cmd.severity];
    const findings = sourceFindings(files, { sheets: wb.sheets.map((s) => s.name), tables: wb.tables.map((t) => t.displayName), spills: spillMap(wb.sheets), links: wb.externalLinks, ...(lock ? { lock } : {}) }).filter((f) => SOURCE_RANK[f.severity] >= min);
    const counts = { error: 0, warning: 0, info: 0, hint: 0 };
    for (const f of findings) counts[f.severity]++;
    out.source = { project, counts, findings };
  }
  return out;
}

/** C15 reads the labels: the cells' values, when they can be read (without them C15 says nothing). */
function cellValuesOf(bytes: Uint8Array): Pick<AuditOptions, "values"> {
  try {
    return { values: cellValueMap(bytes) };
  } catch {
    return {};
  }
}

/** The source checker's findings as text: one line each, the quick fix the editor offers named. */
export function sourceCheckText(s: SourceCheck): string {
  const c = s.counts;
  const lines = [`source ${s.project}: ${c.error} error${c.error === 1 ? "" : "s"}, ${c.warning} warning${c.warning === 1 ? "" : "s"}, ${c.info} info, ${c.hint} hint${c.hint === 1 ? "" : "s"} (the editor's checks as you type)`];
  for (const f of s.findings) lines.push(`  ${f.severity.padEnd(7)} ${f.file}:${f.line}:${f.column} ${f.code}: ${f.message}${f.fix ? ` [quick fix: ${f.fix.title}]` : ""}`);
  return lines.join("\n") + "\n";
}

function runCheck(cmd: CheckCommand, io: Io): number {
  let r: ReturnType<typeof check>;
  try {
    r = check(cmd);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    if (cmd.json) io.out(JSON.stringify({ ok: false, error: msg }, null, 2) + "\n");
    else io.err(`xln: cannot check ${cmd.workbook}: ${msg}\n`);
    return 2;
  }
  if (cmd.json) io.out(JSON.stringify({ ok: true, ...r }, null, 2) + "\n");
  else {
    if (r.config) {
      io.out(`settings: ${r.config.path}\n`);
      for (const p of r.config.problems) io.err(`xln: ${r.config.path}: ${p}\n`);
      for (const n of r.config.notes ?? []) io.out(`note: ${n}\n`);
    }
    io.out(renderAuditReport(r, { maxPerCheck: 100 }).text);
    if (r.source) io.out("\n" + sourceCheckText(r.source));
  }
  return r.counts.error > 0 || (r.source?.counts.error ?? 0) > 0 ? 1 : 0;
}

function runFormulas(cmd: FormulasCommand, io: Io): number {
  let o: FormulasOutcome;
  try {
    o = formulas(cmd);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    if (cmd.json) io.out(JSON.stringify({ ok: false, error: msg }, null, 2) + "\n");
    else io.err(`xln: cannot read formulas of ${cmd.workbook}: ${msg}\n`);
    return 1;
  }
  if (cmd.json) {
    const sheets = o.sheets.map(({ sheet, lines }) => (sheet === undefined ? { lines } : { sheet, lines }));
    io.out(JSON.stringify({ ok: true, workbook: o.workbook, order: o.order, sheets }, null, 2) + "\n");
  } else {
    io.out(o.sheets.map((s) => s.text).join("\n"));
  }
  return 0;
}

function parseBuildArgs(cmd: "build" | "verify" | "apply", rest: string[]): BuildCommand | VerifyCommand | ApplyCommand | string {
  const c: { workbook: string; changes: string; json: boolean; project?: string; out?: string; before?: string; tolerance?: number; dryRun: boolean; force: boolean; reopen: boolean; discard: boolean; embed?: boolean; tags: boolean } = {
    workbook: "",
    changes: "",
    json: false,
    dryRun: false,
    force: false,
    reopen: false,
    discard: false,
    tags: true,
  };
  const valued = cmd === "build" ? ["--project", "--out"] : cmd === "apply" ? ["--out"] : ["--before", "--tolerance"];
  const writes = cmd === "build" || cmd === "apply";
  for (let k = 0; k < rest.length; k++) {
    const a = rest[k]!;
    if (a === "--json") c.json = true;
    else if (writes && a === "--dry-run") c.dryRun = true;
    else if (cmd === "build" && a === "--force") c.force = true;
    else if (cmd === "build" && a === "--rescope-slots") return "--rescope-slots is gone (M3d): a build never changes a scope on its own. To make a workbook name on a sheet's cell local, remove the @workbook line above it in names/sheets/<Sheet>.xln, then build";
    else if (cmd === "build" && a === "--no-embed") c.embed = false;
    else if (cmd === "build" && a === "--embed") c.embed = true;
    else if (cmd === "build" && a === "--no-tags") c.tags = false;
    else if (writes && a === "--reopen") c.reopen = true;
    else if (writes && a === "--discard") c.discard = true;
    else if (valued.includes(a)) {
      const v = rest[++k];
      if (v === undefined) return `${a} needs a value`;
      if (a === "--project") c.project = v;
      else if (a === "--out") c.out = v;
      else if (a === "--before") c.before = v;
      else {
        const t = Number(v);
        if (!(t >= 0)) return "--tolerance must be a number of at least 0";
        c.tolerance = t;
      }
    } else if (a.startsWith("-")) return `unknown option '${a}'`;
    else if (c.workbook === "") c.workbook = a;
    else if (cmd === "apply" && c.changes === "") c.changes = a;
    else return `unexpected argument '${a}'`;
  }
  if (c.workbook === "") return `${cmd} needs a workbook`;
  if (cmd === "apply" && c.changes === "") return "apply needs a change set file";
  if (c.discard && !c.reopen) return "--discard goes with --reopen";
  if (c.reopen && c.dryRun) return "--reopen and --dry-run do not go together";
  return c;
}

/** With `--reopen`: `work` runs between closing the target in Excel and opening it again. */
function withReopen<T extends { exit: number }>(cmd: { workbook: string; out?: string; reopen?: boolean; discard?: boolean }, work: () => T): { result: T | undefined; exit: number; reopen?: ReopenReport } {
  if (!cmd.reopen) {
    const result = work();
    return { result, exit: result.exit };
  }
  const r = aroundExcel(cmd.out ?? cmd.workbook, excelControl(), cmd.discard ?? false, work);
  return { result: r.result, exit: r.exit ?? r.result?.exit ?? 2, reopen: r.reopen };
}

async function mainApply(rest: string[], io: Io): Promise<number> {
  const parsed = parseBuildArgs("apply", rest) as ApplyCommand | string;
  if (typeof parsed === "string") {
    io.err(`xln: ${parsed}\n${USAGE}`);
    return 2;
  }
  for (const f of [parsed.workbook, parsed.changes]) {
    if (!existsSync(f)) {
      io.err(`xln: no such file: ${f}\n`);
      return 2;
    }
  }
  const run = (): ApplyOutcome => {
    try {
      return runApply(parsed);
    } catch (e) {
      return { exit: 2, message: e instanceof Error ? e.message : String(e), changes: [] };
    }
  };
  const { result, exit, reopen } = withReopen(parsed, run);
  if (parsed.json) io.out(JSON.stringify({ ...(result ? (applyJson(result) as object) : { ok: false }), exit, reopen: reopen ?? null }, null, 2) + "\n");
  else {
    const text = (result ? applyText(parsed, result) : `xln apply ${parsed.workbook}: nothing done\n`) + (reopen ? reopenText(reopen) : "");
    (exit === 0 ? io.out : io.err)(text);
  }
  return exit;
}

async function mainBuild(cmd: "build" | "verify", rest: string[], io: Io): Promise<number> {
  const parsed = parseBuildArgs(cmd, rest);
  if (typeof parsed === "string") {
    io.err(`xln: ${parsed}\n${USAGE}`);
    return 2;
  }
  if (!existsSync(parsed.workbook)) {
    io.err(`xln: no such file: ${parsed.workbook}\n`);
    return 2;
  }
  if (cmd === "build") {
    const b = parsed as BuildCommand;
    const run = (): BuildOutcome => {
      try {
        return runBuild(b, defaultOut);
      } catch (e) {
        return { exit: 2, projectFiles: [], message: e instanceof Error ? e.message : String(e) };
      }
    };
    const { result: o, exit, reopen } = withReopen(b, run);
    if (b.json) io.out(JSON.stringify({ ...(o ? (buildJson(o) as object) : { ok: false }), ok: exit === 0, exit, reopen: reopen ?? null }, null, 2) + "\n");
    else {
      let text = o ? buildText(b, o) : `xln build ${b.workbook}: nothing done\n`;
      // The verify hint follows the reopen report: Excel has the file open now.
      if (reopen) text += reopenText(reopen);
      (exit === 0 ? io.out : io.err)(text);
    }
    return exit;
  }
  const v = parsed as VerifyCommand;
  try {
    const { report, before, warnings } = runVerify(v);
    if (v.json) io.out(JSON.stringify({ ok: report.changed.length === 0, before, warnings, ...report }, null, 2) + "\n");
    else io.out(verifyText(v, before, report, warnings));
    return report.changed.length === 0 && report.sheetsAdded.length === 0 && report.sheetsRemoved.length === 0 ? 0 : 1;
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    if (v.json) io.out(JSON.stringify({ ok: false, error: msg }, null, 2) + "\n");
    else io.err(`xln: cannot verify ${v.workbook}: ${msg}\n`);
    return 2;
  }
}

function mainRename(rest: string[], io: Io): number {
  const positional: string[] = [];
  let dryRun = false;
  let json = false;
  for (const a of rest) {
    if (a === "--dry-run") dryRun = true;
    else if (a === "--json") json = true;
    else if (a.startsWith("--")) {
      io.err(`xln: unknown option '${a}'\n${USAGE}`);
      return 2;
    } else positional.push(a);
  }
  if (positional.length !== 3) {
    io.err(`xln: rename needs a project (or workbook), the name and the new name\n${USAGE}`);
    return 2;
  }
  const cmd: RenameCommand = { project: positional[0]!, name: positional[1]!, to: positional[2]!, dryRun, json };
  let o: RenameOutcome;
  try {
    o = runRename(cmd);
  } catch (e) {
    o = { exit: 2, project: cmd.project, written: [], error: e instanceof Error ? e.message : String(e) };
  }
  if (json) io.out(JSON.stringify(renameJson(o), null, 2) + "\n");
  else (o.exit === 0 ? io.out : io.err)(renameText(cmd, o));
  return o.exit;
}

export async function main(args: string[], io: Io = defaultIo): Promise<number> {
  if (args[0] === "build" || args[0] === "verify") return mainBuild(args[0], args.slice(1), io);
  if (args[0] === "rename") return mainRename(args.slice(1), io);
  if (args[0] === "apply") return mainApply(args.slice(1), io);
  if (args[0] === "lib") return mainLib(args.slice(1), io);
  const cmd = parseArgs(args);
  if (typeof cmd === "string") {
    if (cmd) io.err(`xln: ${cmd}\n`);
    io.err(USAGE);
    return cmd ? 2 : 0;
  }
  if (!existsSync(cmd.workbook)) {
    io.err(`xln: no such file: ${cmd.workbook}\n`);
    return cmd.cmd === "check" ? 2 : 1;
  }
  if (cmd.cmd === "check") return runCheck(cmd, io);
  if (cmd.cmd === "formulas") return runFormulas(cmd, io);
  if (cmd.cmd === "graph") return runGraph(cmd, io);
  let outcome: PullOutcome;
  try {
    outcome = pull(cmd);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    if (cmd.json) io.out(JSON.stringify({ ok: false, error: msg }, null, 2) + "\n");
    else io.err(`xln: cannot pull ${cmd.workbook}: ${msg}\n`);
    return 1;
  }
  if (outcome.refused) {
    if (cmd.json) io.out(JSON.stringify({ ok: false, refused: true, out: outcome.out, unbuilt: outcome.unbuilt, rewritten: outcome.rewritten, notices: outcome.notices }, null, 2) + "\n");
    else io.err(refusedText(outcome, cmd.workbook));
    return 1;
  }
  if (cmd.json) {
    io.out(JSON.stringify({ ok: true, out: outcome.out, files: outcome.written, notices: outcome.notices, report: outcome.report, unchanged: outcome.unchanged ?? [], discarded: outcome.unbuilt, rewritten: outcome.rewritten, provenance: outcome.provenance ?? [] }, null, 2) + "\n");
  } else {
    io.out(summary(outcome));
  }
  return 0;
}
