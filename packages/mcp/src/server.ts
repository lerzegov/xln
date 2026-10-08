// xln as an MCP server: six tools over the CLI's own functions. The descriptions are the
// interface an agent reads, so they say when to use each tool and what it will not do.

import { McpServer, type CallToolResult } from "@modelcontextprotocol/server";
import * as z from "zod";
import { Roots } from "./paths.js";
import { buildPlanTool, buildTool, checkTool, libStatusTool, namesTool, pullTool, ToolFailure, type ToolOutput } from "./tools.js";

export { Roots, PathError } from "./paths.js";
export { ToolFailure, type ToolOutput } from "./tools.js";

export const SERVER_NAME = "xln";
export const SERVER_VERSION = "0.0.0";

export interface ServerOptions {
  /** Folders the tools may read and write under; relative paths resolve against the first. */
  roots: string[];
}

const INSTRUCTIONS = `xln treats an Excel workbook's defined names, LAMBDAs and cell formulas as source code, with no Excel running.
Workflow: xln_check a workbook for a deterministic verdict (run it after anything edits a workbook); xln_pull it into a project folder (names/*.xln text files you can read and edit); edit those files; xln_check the project folder; xln_build_plan to see the change set; show it to the user; xln_build with confirm: true and the plan's planId to write it.
xln never changes names or scopes on its own, never merges a conflict silently, refuses to write while Excel has the file open, keeps <workbook>.backup.xlsx, and reads every build back.
Paths are absolute or relative to the first allowed root; paths outside the roots are refused.`;

function result(o: ToolOutput): CallToolResult {
  return {
    content: [
      { type: "text", text: o.text },
      { type: "text", text: JSON.stringify(o.data) },
    ],
    structuredContent: o.data,
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
        "names still used by cells) and warnings. Nothing is written. Call it before xln_build and show the plan to the user; it returns a planId for xln_build.",
      inputSchema: z.object({
        workbook: path("Workbook file"),
        project: z.string().optional().describe("Project folder (default <workbook>.xln beside the workbook)"),
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

  return server;
}
