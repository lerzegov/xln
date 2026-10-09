# @xln/mcp

`xln-mcp`: an [MCP](https://modelcontextprotocol.io) server that lets an AI agent
(Claude Code, Claude Desktop, other MCP clients) use xln as its **critic and ledger**.
The agent edits the workbook or its names files. xln tells it, deterministically, what is
wrong (`xln_check`, `xln_graph`). It shows every name, LAMBDA and cell formula as
reviewable text (`xln_pull`, `xln_names`, `xln_formulas`), renames explicitly
(`xln_rename`), and writes source back only through the CLI's guarded build
(`xln_build_plan`, then `xln_build` after the user approves); `xln_verify` then shows
whether any value moved. The library tools keep a LAMBDA library in step
(`xln_lib_status`, `xln_lib_take`, `xln_lib_base`, `xln_lib_publish`).

It is a thin layer over `@xln/cli`. Each tool calls the same function the CLI command
runs, in process (no shelling out), so verdicts and messages are the CLI's. Transport:
stdio.

## Install

From a checkout of the repository:

```
npm install
npm run build
```

This gives `packages/mcp/bin/xln-mcp.js`, also on the path inside the workspace as
`npx xln-mcp`. Run `xln-mcp --help` for the options.

```
xln-mcp [--root <folder>]...
```

`--root` is a folder the tools may read and write under. Repeat it for several folders.
Relative paths in tool calls resolve against the first root. The default root is the
folder the server is started in.

### Claude Code

```
claude mcp add xln -- npx xln-mcp --root ~/models
```

Use `npx xln-mcp` from inside the repository, or after the package is published or
`npm link`ed. Anywhere else, give node the script:

```
claude mcp add xln -- node /path/to/excel_dim/packages/mcp/bin/xln-mcp.js --root ~/models
```

### Claude Desktop

In `claude_desktop_config.json` (Settings → Developer → Edit Config):

```json
{
  "mcpServers": {
    "xln": {
      "command": "node",
      "args": ["/path/to/excel_dim/packages/mcp/bin/xln-mcp.js", "--root", "/Users/me/models"]
    }
  }
}
```

Claude Desktop starts servers with no useful working folder, so always pass `--root`
with an absolute path.

## Tools

Every result carries a short text summary and structured JSON (`structuredContent`, and
the same JSON as a second text block for clients that read only text). The summary is
also the first field of the structured JSON (`summary`): the spec makes the text block a
serialization of `structuredContent` "for backwards compatibility", so a client may pass
the model only `structuredContent` when it is present. Claude Code did in the MCP trial
of 2026-10-09 (the agent saw the plan's JSON, not its text summary; issues on Claude Code's
tracker report the same), so nothing an agent needs is only in the text block. Failures are MCP
tool errors (`isError`) with the CLI's wording: an unreadable file, a path outside the
roots, a refused pull or build.

| Tool | Writes | Arguments | What it returns |
|---|---|---|---|
| `xln_check` | no | `path` (workbook or project folder), `only?` (`["C2","C6"]`), `severity?` (`error`/`warning`/`info`/`hint`), `detail?` (`summary`/`full`), `maxFindings?` (200) | `verdict` (`errors`/`warnings`/`clean`), `ok`, `counts`, `byCheck`, the findings of C1–C13 as `xln check --json` gives them, the project's `source` findings (file:line, quick fix) when there is a project, and the census as a summary. `detail: "full"` adds the whole name census and spill census (about 250 kB on lbo-ep03r) |
| `xln_names` | no | `path` (workbook or project folder), `query?`, `names?`, `kind?`, `scope?` (`workbook` or a sheet), `limit?` (100), `offset?` | Each name's `key`, `scope`, `kind` (constant, range, spill, table, formula, lambda, cell, slot), `definition` as Excel displays it (a named cell's formula), `doc`, `hidden`, `cell`, `params`, `libBase`, and `file` (`names/FN.xln:12`). A workbook is read as a pull would write it, without writing anything. A project folder is its source, unbuilt edits included |
| `xln_pull` | project folder | `workbook`, `out?` (default `<workbook>.xln`), `discard?` (false) | The `xln pull --json` result: files written, report, notices. Refused, with the unbuilt edits listed, while the project has source edits not built yet |
| `xln_build_plan` | no | `workbook`, `project?`, `detail?` (`compact`/`full`) | The `xln build --dry-run --json` result: `changeSet`, `conflicts`, `problems`, `excelChanges`, `status` (`refused` is a result here, with its reasons), plus `willWrite`, `excelOpen` and `planId`. `compact` (default): each change without its stored form (`op`, name and scope or sheet and range, `display`, `fields`, `comment`), a cell's `previous` formula as Excel displays it, `layout: "one line, as in the workbook"` on a cell written on one line, embedded source as its file list, and updates of the provenance tag alone folded into `changeSet.provenanceOnly` (`count`, `names`). `full`: the CLI's change set. The `planId` is the same either way. On lbo-ep03r with 16 real changes and 22 tag updates: 9 kB of structured data instead of 22 kB |
| `xln_build` | workbook, backup, lockfile | `workbook`, `project?`, `out?`, `confirm` (must be `true`), `planId?`, `detail?` | The `xln build --json` result: `written`, `backup`, `readBack`, project files updated; its change set compact like the plan's unless `detail: "full"` |
| `xln_lib_status` | no | `path` (workbook or project folder), `lib?` (default the project's `xln.config.json` `"library"`), `diffs?` (true) | The `xln lib status --json` result: per library function identical, outdated, modified, both changed, differs, missing; local only; diffs |
| `xln_formulas` | no | `workbook`, `sheet?`, `order?` (`appearance`/`calculation`), `workbookWide?` (all sheets in calculation order, the CLI's `--workbook`), `query?`, `names?`, `detail?` (`compact`/`full`), `limit?` (50), `offset?` | The `xln formulas --json` view, paged: `total`, `matched`, per-sheet counts, and per formula `sheet`, `cell` (`C6#` for a dynamic array), `kind`, `extent`, `defines` (the names on the cell or spill), `formula` as displayed, saved `value`, `reads` (names), `refs`; in calculation order also `level`, `cycle`, `dependsOn`. `detail: "full"` gives the CLI's lines (stored text, spans, value object). `query` matches `Sheet!Cell`, the formula or a defined name; `names` keeps the formulas that define or read them |
| `xln_graph` | no | `workbook`, `maxItems?` (200) | The `xln graph --json` summary: counts, longest chain, circular references, recursive LAMBDAs, references not followed (dynamic, external, broken), C9 fixed references into a spill, C10 unused names, C12 name cycles; `totals` per list |
| `xln_verify` | no | `workbook` (after the build, saved by Excel), `before?` (default `<workbook>.backup.xlsx`), `tolerance?` (0), `maxChanges?` (200) | The `xln verify --json` result: every worksheet cell with a saved value compared (numbers within the relative tolerance, everything else exactly), `changed` cells, sheets on one side only, `warnings` when a side's values were not calculated by Excel, and `verdict` (`same`/`changed`), `changedTotal` |
| `xln_rename` | project's names files | `path` (project folder, or the workbook: its project beside it), `name` (`Name` or `Sheet!Name`), `to`, `dryRun?` (false) | The `xln rename --json` result: the edits (with `line` and the text replaced), files `written`, what the next build will do, the label notice; refused (nothing written) when the next build would refuse the rename |
| `xln_lib_take` | project's names files | `project`, `name`, `lib?`, `dryRun?`, `discard?` (false) | The `xln lib take --json` result: the diff, the state before, the base recorded. A copy edited here (modified, both changed) or without a base (differs) is refused unless `discard: true` |
| `xln_lib_base` | project's names files | `project`, `name?` or `all?`, `lib?`, `dryRun?` | The `xln lib base --json` result: the bases recorded (`@from(lib #…)`), those skipped and why |
| `xln_lib_publish` | the library file, project's names file | `project`, `name`, `lib?`, `dryRun?`, `confirm?` (true to write) | The `xln lib publish --json` result: the library file created or updated, the diff, the base recorded in the project. Writes only with `confirm: true`, and only into a library under a root |

The server's instructions give an agent the workflow: check, pull, read the model
(`xln_names`, `xln_formulas`, `xln_graph`), edit `names/*.xln` (renames with
`xln_rename`), check the project, plan, show the plan to the user, build, then
`xln_verify` once the user has opened and saved the built file in Excel; and the library
tools for keeping LAMBDAs in step with a shared library.

## Safety model

- **Roots.** Every path an agent passes, and every path a tool derives from it (the
  default project folder, `out`, the backup `xln_verify` reads, the workbook beside a
  project `xln_rename` plans against), must lie under a `--root`. Symbolic links are
  resolved first, so a link inside a root cannot lead out of it.
- **The library folder.** The library a project's `xln.config.json` names is *read*
  wherever it is (`xln_lib_status`, `xln_lib_take`, `xln_lib_base`), as the CLI does: a
  shared library often lives outside the models' folder. *Writing* it is another matter:
  `xln_lib_publish` refuses unless the library folder (given or configured, after
  resolving links) lies under a root, even for a dry run. To let an agent publish, start
  the server with the library as one of its roots (`--root ~/models --root ~/lib`).
- **Reads by default.** Seven of the thirteen tools only read. `xln_pull` writes only
  the project folder. It refuses to replace source edits not built yet unless the call
  says `discard: true`, and the tool description tells the agent to pass that only when
  the user agrees.
- **Source edits are the agent's to make.** `xln_rename`, `xln_lib_take` and
  `xln_lib_base` write only the project's names files (and `library-bases/`), which the
  agent may edit by hand anyway, so they take no `confirm`; `dryRun: true` shows every
  edit first. Their effect reaches the workbook only through a build, which the user
  approves. `xln_rename` refuses a rename the next build would refuse, and
  `xln_lib_take` keeps the CLI's refusal of a copy with an edit of its own (modified,
  both changed, differs) unless `discard: true`, which its description reserves for when
  the user agrees.
- **The library needs consent.** The library is shared with other workbooks, so
  `xln_lib_publish` is consented like a build: a dry run (`dryRun: true`) shows the
  diff, and only `confirm: true` writes the library file.
- **Build needs consent.** `xln_build` takes `confirm: true` (the schema has no other
  value). Its description asks the agent to show the user the plan from `xln_build_plan`
  first. With the plan's `planId`, the build refuses when the change set is no longer
  that plan (the source or the workbook moved since).
- **Every CLI safety stays.** These come from `runBuild` itself, the CLI's code: the
  three-way check against the lockfile (a conflict refuses, nothing is merged); refusal
  on source errors and on names still used by cells; refusal while Excel has the file
  open (`~$` lock file); the previous file kept as `<workbook>.backup.xlsx`; the read-back
  before and after writing, with the original restored on a mismatch.
- **Never forced.** The server has no `--force`, no `--reopen` (it never drives Excel)
  and no way to discard Excel's unsaved changes.

## Runtime dependencies

| Package | Why |
|---|---|
| [`@modelcontextprotocol/server`](https://www.npmjs.com/package/@modelcontextprotocol/server) | The official MCP TypeScript SDK, v2 (the current stable line; it replaces the single `@modelcontextprotocol/sdk` of v1, which is in maintenance and stops at the 2025-11-25 spec). Server, tool registration, stdio transport, and the 2025/2026 protocol negotiation. Pure JavaScript; its only dependencies are `zod` and `@modelcontextprotocol/core` (also pure JavaScript). |
| [`zod`](https://zod.dev) | The tools' input schemas, which the SDK turns into JSON Schema for clients and validates calls against. Pure JavaScript, no dependencies. It is a dependency of the SDK too. |
| `@xln/cli`, `@xln/core` | The workspace's own packages: the commands' functions and the engine. |

Development only: `@modelcontextprotocol/client`, the SDK's client, which the tests use to
call the server over the SDK's in-memory transport.

## Tests

`packages/mcp/test/server.test.ts` runs the server in process and calls every tool on
copies of `probes/results` workbooks in a temporary root: the findings equal
`xln check`'s, the pull guard, a build refused for a missing confirm, a stale plan, Excel's
lock file and a source error, a build with its backup, a build to a copy, paths outside
the roots and through a link; the formula view's paging and filters, the graph summary,
verify against the backup and a given copy, a rename's dry run and write, take refused
without `discard`, base, and publish refused without `confirm` or with a library outside
the roots (configured, or through a link). `corpus.test.ts` (with `XLN_CORPUS`) runs
check, names, pull, plan, formulas, graph and verify on a copy of every corpus workbook.
It also checks that a summary check and a default formulas page stay under 80 kB (a
formulas page is about 35 kB on lbo-ep03r; its full view runs to 900 kB). On lbo-ep03r it plans the MCP trial's edit (FN.LAG in 18 places) and holds the compact
plan under 12 kB, less than half the full one.
