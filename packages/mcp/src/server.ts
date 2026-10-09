// xln as an MCP server: thirteen tools over the CLI's own functions. The descriptions are the
// interface an agent reads, so they say when to use each tool and what it will not do.

import { McpServer, type CallToolResult } from "@modelcontextprotocol/server";
import * as z from "zod";
import { Roots } from "./paths.js";
import {
  buildPlanTool,
  buildTool,
  checkTool,
  formulasTool,
  graphTool,
  libBaseTool,
  libPublishTool,
  libStatusTool,
  libTakeTool,
  namesTool,
  pullTool,
  renameTool,
  verifyTool,
  ToolFailure,
  type ToolOutput,
} from "./tools.js";

export { Roots, PathError } from "./paths.js";
export { ToolFailure, type ToolOutput } from "./tools.js";

export const SERVER_NAME = "xln";
export const SERVER_VERSION = "0.0.0";

export interface ServerOptions {
  /** Folders the tools may read and write under; relative paths resolve against the first. */
  roots: string[];
}

const INSTRUCTIONS = `xln treats an Excel workbook's defined names, LAMBDAs and cell formulas as source code, with no Excel running.
Workflow: xln_check a workbook for a deterministic verdict (run it after anything edits a workbook); xln_pull it into a project folder (names/*.xln text files you can read and edit); read the model as code with xln_names and xln_formulas (paged; filter by sheet, query or names), and xln_graph for cycles and references that cannot be followed; edit the names files (to rename a name, use xln_rename, not a hand edit: it records the rename and rewrites every reader); xln_check the project folder; xln_build_plan to see the change set; show it to the user; xln_build with confirm: true and the plan's planId to write it. After the user opens the built workbook in Excel and saves it, xln_verify compares its values with the backup: a build meant not to change numbers should show none changed.
LAMBDA library: xln_lib_status compares with the library; xln_lib_take takes the library's version, xln_lib_base records the base of an identical copy, xln_lib_publish writes a copy into the shared library (dryRun first, then confirm: true once the user approves).
xln never changes names or scopes on its own, never merges a conflict silently, refuses to write while Excel has the file open, keeps <workbook>.backup.xlsx, and reads every build back.
Paths are absolute or relative to the first allowed root; paths outside the roots are refused.`;

/**
 * The summary goes into the structured data too (`summary`, first): the spec makes the
 * text block the serialized structured content "for backwards compatibility", and a client
 * may hand the model only `structuredContent` when it is present (Claude Code did in the
 * MCP trial of 2026-10-09: the agent saw the JSON, not the summary). Clients that read
 * only `content` get the summary and the JSON without it.
 */
function result(o: ToolOutput): CallToolResult {
  return {
    content: [
      { type: "text", text: o.text },
      { type: "text", text: JSON.stringify(o.data) },
    ],
    structuredContent: { summary: o.text, ...o.data },
  };
}

function failure(e: unknown): CallToolResult {
  const message = e instanceof Error ? e.message : String(e);
  const data = e instanceof ToolFailure ? e.data : {};
  return { isError: true, content: [{ type: "text", text: message }], structuredContent: { ok: false, error: message, ...data } };
}

function run(f: () => ToolOutput): CallToolResult {
  try {
    return result(f());
  } catch (e) {
    return failure(e);
  }
}

const detail = z
  .enum(["compact", "full"])
  .optional()
  .describe("'compact' (default): the change set without stored forms, provenance-only updates counted; 'full': the CLI's change set as it is");

const path = (what: string) => z.string().min(1).describe(`${what}: absolute, or relative to the server's first root`);

export function createServer(opts: ServerOptions): McpServer {
  const roots = new Roots(opts.roots);
  const server = new McpServer({ name: SERVER_NAME, version: SERVER_VERSION }, { capabilities: { tools: {} }, instructions: INSTRUCTIONS });

  server.registerTool(
    "xln_check",
    {
      title: "Check a workbook or project",
      description:
        "Audit an Excel workbook (.xlsx/.xlsm) or an xln project folder and return a deterministic verdict: the findings of checks C1-C13 " +
        "(syntax, missing _xlfn. prefixes that give #NAME?, built-in collisions, unresolved references and #REF!, unqualified sheet-scoped reads, " +
        "LAMBDA arity, Excel limits, fixed references into spills, unused names, copy drift, cycles, constants in LAMBDAs), plus the project's " +
        "source findings (file:line, with the editor's quick fix) when the workbook has its project beside it or a project folder is given. " +
        "Same findings as `xln check --json`. Use it after you or any tool edits a workbook or its names files, and before building. " +
        "verdict 'errors' means fix before building. Reads only.",
      inputSchema: z.object({
        path: path("Workbook file, or project folder (<workbook>.xln)"),
        only: z.array(z.string()).optional().describe("Run only these checks, e.g. [\"C2\", \"C6\"]"),
        severity: z.enum(["error", "warning", "info", "hint"]).optional().describe("Report only findings at least this severe (default: all)"),
        detail: z.enum(["summary", "full"]).optional().describe("'full' adds the whole name census and spill census (large on big workbooks); default 'summary'"),
        maxFindings: z.number().int().min(1).max(5000).optional().describe("Findings returned per list (default 200; counts are always complete)"),
      }),
      annotations: { readOnlyHint: true, openWorldHint: false },
    },
    async (args) => run(() => checkTool(roots, args)),
  );

  server.registerTool(
    "xln_names",
    {
      title: "List or show defined names",
      description:
        "List the defined names of a workbook (as a pull would write them, nothing written) or of a project folder (its names/*.xln source, " +
        "unbuilt edits included): name, scope (sheet or workbook), kind (constant, range, spill, table, formula, lambda, cell, slot), " +
        "definition as Excel displays it (for a named cell, the cell's formula), doc comment, hidden flag, cell address, LAMBDA parameters, " +
        "and the file:line where it is in the project. Use it to find or read names instead of opening files; filter with query, names, kind or scope. Reads only.",
      inputSchema: z.object({
        path: path("Workbook file, or project folder"),
        query: z.string().optional().describe("Case-insensitive text to find in the name, definition or doc comment"),
        names: z.array(z.string()).optional().describe("Exactly these names (Name or Sheet!Name, case-insensitive)"),
        kind: z.enum(["constant", "range", "spill", "table", "formula", "lambda", "unparsed", "cell", "slot"]).optional(),
        scope: z.string().optional().describe("'workbook', or a sheet name for that sheet's local names"),
        limit: z.number().int().min(1).max(5000).optional().describe("Names returned (default 100)"),
        offset: z.number().int().min(0).optional().describe("Skip this many matches (paging)"),
      }),
      annotations: { readOnlyHint: true, openWorldHint: false },
    },
    async (args) => run(() => namesTool(roots, args)),
  );

  server.registerTool(
    "xln_pull",
    {
      title: "Pull a workbook into a project folder",
      description:
        "Read the workbook's names and cell formulas and write them as a project folder (default <workbook>.xln beside it): names/*.xln modules, " +
        "names/sheets/<Sheet>.xln, workbook.manifest.json, xln.lock.json. The workbook is only read. Use it to get editable source, and after the " +
        "workbook changed in Excel or another tool, so the text diff shows what changed. Every pull is fresh: it refuses (tool error, nothing written) " +
        "while the project has source edits not built yet, and lists them. Build them first; pass discard: true only when the user agrees to lose them.",
      inputSchema: z.object({
        workbook: path("Workbook file (.xlsx or .xlsm)"),
        out: z.string().optional().describe("Project folder to write (default <workbook>.xln beside the workbook)"),
        discard: z.boolean().optional().describe("Replace the project even though it has source edits not built (they are lost). Default false"),
      }),
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async (args) => run(() => pullTool(roots, args)),
  );

  server.registerTool(
    "xln_build_plan",
    {
      title: "Plan a build (dry run)",
      description:
        "Compare the project's source with the last pull (lockfile) and the workbook, three-way, and return the change set a build would write " +
        "(names to set, rename, re-scope or delete; cell formulas to replace), conflicts (changed in Excel and in source), refusals (source errors, " +
        "names still used by cells) and warnings. Nothing is written. Call it before xln_build and show the plan to the user (summary is the plan as text); it returns a planId for xln_build. " +
        "By default the change set is compact: each change's display text (what the build writes), previous cell formulas as displayed, a cell written on one line " +
        "as in the workbook marked by layout, and updates of the provenance tag alone folded into changeSet.provenanceOnly; detail: 'full' gives the CLI's change set with stored forms.",
      inputSchema: z.object({
        workbook: path("Workbook file"),
        project: z.string().optional().describe("Project folder (default <workbook>.xln beside the workbook)"),
        detail: detail,
      }),
      annotations: { readOnlyHint: true, openWorldHint: false },
    },
    async (args) => run(() => buildPlanTool(roots, args)),
  );

  server.registerTool(
    "xln_build",
    {
      title: "Build: write the project into the workbook",
      description:
        "Write the project's names and cell formulas into the workbook: the plan of xln_build_plan. Only with confirm: true, after the user approved " +
        "that plan; pass its planId so the build refuses if the plan changed since. Safety kept from `xln build`: refused (nothing written) on source " +
        "errors or conflicts, and while Excel has the file open (~$ lock file); the previous file is kept as <workbook>.backup.xlsx; the result is " +
        "read back before and after writing and the original restored if it does not match. Never forces and never discards Excel's unsaved changes. " +
        "Once written, it removes from the names files the @renamed(Old) lines of the renames it built, and of renames built earlier (renamedRemoved): the one source edit a build makes. " +
        "For a renamed name whose label cells still read the old name, the result ends with a label notice (labelNotices): the cells, and the " +
        "Find & Replace that fixes them in Excel; xln never writes cell values, so show it to the user. " +
        "After it, the user opens the workbook in Excel and saves; then xln_pull shows what Excel recomputed.",
      inputSchema: z.object({
        workbook: path("Workbook file"),
        project: z.string().optional().describe("Project folder (default <workbook>.xln beside the workbook)"),
        out: z.string().optional().describe("Write the built workbook to this file instead; the original and the lockfile stay as they are"),
        confirm: z.literal(true).describe("Must be true: the user approved the plan"),
        planId: z.string().optional().describe("The planId from xln_build_plan: refuse if the change set is no longer that plan"),
        detail: detail,
      }),
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false },
    },
    async (args) => run(() => buildTool(roots, args)),
  );

  server.registerTool(
    "xln_lib_status",
    {
      title: "Compare with a LAMBDA library",
      description:
        "Compare a workbook's or project's LAMBDAs with a library folder of .lambda files: per library function identical, outdated (the library moved), " +
        "modified (edited here), both changed, differs (no base recorded) or missing, plus local-only functions, with diffs (- the copy, + the library). " +
        "The library is the lib argument or \"library\" in the project's xln.config.json. Reads only.",
      inputSchema: z.object({
        path: path("Workbook file, or project folder"),
        lib: z.string().optional().describe("Library folder (default: the project's xln.config.json \"library\")"),
        diffs: z.boolean().optional().describe("Include the diffs (default true)"),
      }),
      annotations: { readOnlyHint: true, openWorldHint: false },
    },
    async (args) => run(() => libStatusTool(roots, args)),
  );

  server.registerTool(
    "xln_formulas",
    {
      title: "Read the cell formulas as code",
      description:
        "The workbook's cell formulas as a reviewable listing, as `xln formulas --json` gives it: per formula its sheet, cell (C6# for a dynamic array), " +
        "kind (normal, shared, array, dynamic-array, data-table), saved extent, the names defined as that cell or spill (defines), the formula as Excel " +
        "displays it, the value saved with the file, the names and cell references it reads. order 'calculation' (or workbookWide: all sheets in one list) " +
        "adds level (longest chain from the inputs), cycle (circular reference number) and dependsOn. Paged: a big model has hundreds of formulas, so " +
        "filter with sheet, query (text in the address, formula or defined name) or names (formulas defining or reading them), and page with limit/offset. " +
        "Use it to read the model before editing it. Reads only.",
      inputSchema: z.object({
        workbook: path("Workbook file"),
        sheet: z.string().optional().describe("Only this sheet"),
        order: z.enum(["appearance", "calculation"]).optional().describe("Per sheet: row by row (default), or each formula after what it reads"),
        workbookWide: z.boolean().optional().describe("All sheets as one list in calculation order (`--workbook`); excludes sheet"),
        query: z.string().optional().describe("Case-insensitive text to find in Sheet!Cell, the formula or a defined name"),
        names: z.array(z.string()).optional().describe("Only formulas that define or read these names (Name or Sheet!Name, case-insensitive)"),
        detail: z.enum(["compact", "full"]).optional().describe("'full': each line as the CLI's JSON has it (stored text, spans, value object); default 'compact'"),
        limit: z.number().int().min(1).max(5000).optional().describe("Formulas returned (default 50)"),
        offset: z.number().int().min(0).optional().describe("Skip this many matches (paging)"),
      }),
      annotations: { readOnlyHint: true, openWorldHint: false },
    },
    async (args) => run(() => formulasTool(roots, args)),
  );

  server.registerTool(
    "xln_graph",
    {
      title: "Dependency summary",
      description:
        "The workbook's dependency graph over cells and names, summarised as `xln graph --json` gives it: node and edge counts, longest chain, " +
        "circular references (members), recursive LAMBDAs (allowed), references that cannot be followed from the file (dynamic: INDIRECT/OFFSET; " +
        "external; broken), C9 fixed references into a spill (with the x# to use), C10 unused names (and names used only by unused ones), C12 name " +
        "cycles. Use it to find what to look at before an edit, or why a value cannot be traced. Reads only.",
      inputSchema: z.object({
        workbook: path("Workbook file"),
        maxItems: z.number().int().min(1).max(5000).optional().describe("Items returned per list (default 200; totals are always complete)"),
      }),
      annotations: { readOnlyHint: true, openWorldHint: false },
    },
    async (args) => run(() => graphTool(roots, args)),
  );

  server.registerTool(
    "xln_verify",
    {
      title: "Verify a build changed no values",
      description:
        "Compare the values saved in two copies of a workbook, cell by cell, as `xln verify --json` does: the copy before a build (default the build's " +
        "backup <workbook>.backup.xlsx) and the workbook after the user opened the built file in Excel and saved it (Excel recalculates it on open). " +
        "Every worksheet cell with a saved value counts, formula or not (a spill's cells too); numbers within the relative tolerance (default 0, exact) " +
        "count as equal, text, booleans and errors must match exactly; sheets on one side only are listed. verdict 'same': no cell changed. " +
        "warnings say when a side's values were not calculated by Excel (a built file not yet saved in Excel has none), which makes the comparison " +
        "empty or partial: ask the user to open and save it first. Use it after a build meant not to change numbers (renames, refactors). Reads only.",
      inputSchema: z.object({
        workbook: path("The workbook after the build, saved by Excel"),
        before: z.string().optional().describe("The copy before the build (default <workbook>.backup.xlsx beside it)"),
        tolerance: z.number().min(0).optional().describe("Relative difference under which numbers count as equal (default 0)"),
        maxChanges: z.number().int().min(1).max(10000).optional().describe("Changed cells returned (default 200; changedTotal is complete)"),
      }),
      annotations: { readOnlyHint: true, openWorldHint: false },
    },
    async (args) => run(() => verifyTool(roots, args)),
  );

  server.registerTool(
    "xln_rename",
    {
      title: "Rename a defined name in the source",
      description:
        "Rename a defined name in a project's source, as `xln rename` does: the statement gets the new name with @renamed(Old) above it (the record the " +
        "build reads; without it a rename is a deletion plus a new name), and every formula in the project's .xln files that reads it is rewritten, token " +
        "by token (strings, LET/LAMBDA variables and other scopes' names of that spelling untouched). Use it instead of editing names by hand. " +
        "With the workbook beside the project, it plans the next build first and refuses (nothing written) a rename that build would refuse " +
        "(a chart or Table column reads the name, the new name would be captured). Writes only the project's names files, like a hand edit; the " +
        "workbook is only read and changes only through xln_build_plan and xln_build with the user's approval. dryRun: true lists every edit " +
        "(file, line, old text) without writing.",
      inputSchema: z.object({
        path: path("Project folder, or the workbook (its <workbook>.xln beside it)"),
        name: z.string().min(1).describe("The name: Name, or Sheet!Name for a sheet's local name"),
        to: z.string().min(1).describe("The new name"),
        dryRun: z.boolean().optional().describe("List the edits, write nothing (default false)"),
      }),
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    },
    async (args) => run(() => renameTool(roots, args)),
  );

  const lib = z.string().optional().describe("Library folder (default: the project's xln.config.json \"library\")");

  server.registerTool(
    "xln_lib_publish",
    {
      title: "Publish a LAMBDA to the library",
      description:
        "Write a workbook-scoped LAMBDA of the project into the library folder as <Name>.lambda (created or updated), as `xln lib publish` does, and " +
        "record in the project that published version as the name's base (@from(lib #hash), its text kept in library-bases/). The library is shared " +
        "with other workbooks: call it with dryRun: true first, show the user the diff, and only then with confirm: true. Refused when the library " +
        "folder is not under an allowed root. Writes the library file and the project's names file; never the workbook (build for that).",
      inputSchema: z.object({
        project: path("Project folder"),
        name: z.string().min(1).describe("The LAMBDA's name (workbook-scoped)"),
        lib,
        dryRun: z.boolean().optional().describe("Show the diff, write nothing (default false)"),
        confirm: z.boolean().optional().describe("Must be true to write (not needed with dryRun): the user approved the diff"),
      }),
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: false },
    },
    async (args) => run(() => libPublishTool(roots, args)),
  );

  server.registerTool(
    "xln_lib_take",
    {
      title: "Take the library's version of a LAMBDA",
      description:
        "Replace a LAMBDA in the project's source with the library's version (definition and doc comment) and record it as the base (@from(lib #hash)), " +
        "as `xln lib take` does. Refused (tool error, nothing written, with the diff) when the copy was edited here (modified, both changed) or records " +
        "no base so an edit cannot be ruled out (differs); pass discard: true only when the user agrees to lose that edit, or publish it instead. " +
        "Writes only the project's source; build to carry it into the workbook.",
      inputSchema: z.object({
        project: path("Project folder"),
        name: z.string().min(1).describe("The library function's name"),
        lib,
        dryRun: z.boolean().optional().describe("Show the diff, write nothing (default false)"),
        discard: z.boolean().optional().describe("Take it even though the copy has (or may have) an edit of its own, which is lost. Default false"),
      }),
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async (args) => run(() => libTakeTool(roots, args)),
  );

  server.registerTool(
    "xln_lib_base",
    {
      title: "Record a LAMBDA's library base",
      description:
        "Record @from(lib #hash) on a LAMBDA identical to the library's that records no base yet (all: true for every such one), with the base's text kept " +
        "in library-bases/, as `xln lib base` does: later status can then tell outdated from modified. Nothing else records a base. Refused for a named " +
        "function that is not identical or already has a base (the reason given). Writes only the project's source; build to carry it into the workbook.",
      inputSchema: z.object({
        project: path("Project folder"),
        name: z.string().min(1).optional().describe("The function (or all: true)"),
        all: z.boolean().optional().describe("Every function identical to the library without a base"),
        lib,
        dryRun: z.boolean().optional().describe("List what would be recorded, write nothing (default false)"),
      }),
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
    },
    async (args) => run(() => libBaseTool(roots, args)),
  );

  return server;
}
