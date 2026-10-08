# @xln/mcp

`xln-mcp`: an [MCP](https://modelcontextprotocol.io) server that lets an AI agent
(Claude Code, Claude Desktop, other MCP clients) use xln as its **critic and ledger**.
The agent edits the workbook or its names files. xln tells it, deterministically, what is
wrong (`xln_check`). It shows every name and LAMBDA as reviewable text (`xln_pull`,
`xln_names`), and writes source back only through the CLI's guarded build
(`xln_build_plan`, then `xln_build` after the user approves).

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
the same JSON as a second text block for clients that read only text). Failures are MCP
tool errors (`isError`) with the CLI's wording: an unreadable file, a path outside the
roots, a refused pull or build.

| Tool | Writes | Arguments | What it returns |
|---|---|---|---|
| `xln_check` | no | `path` (workbook or project folder), `only?` (`["C2","C6"]`), `severity?` (`error`/`warning`/`info`/`hint`), `detail?` (`summary`/`full`), `maxFindings?` (200) | `verdict` (`errors`/`warnings`/`clean`), `ok`, `counts`, `byCheck`, the findings of C1–C13 as `xln check --json` gives them, the project's `source` findings (file:line, quick fix) when there is a project, and the census as a summary. `detail: "full"` adds the whole name census and spill census (about 250 kB on lbo-ep03r) |
| `xln_names` | no | `path` (workbook or project folder), `query?`, `names?`, `kind?`, `scope?` (`workbook` or a sheet), `limit?` (100), `offset?` | Each name's `key`, `scope`, `kind` (constant, range, spill, table, formula, lambda, cell, slot), `definition` as Excel displays it (a named cell's formula), `doc`, `hidden`, `cell`, `params`, `libBase`, and `file` (`names/FN.xln:12`). A workbook is read as a pull would write it, without writing anything. A project folder is its source, unbuilt edits included |
| `xln_pull` | project folder | `workbook`, `out?` (default `<workbook>.xln`), `discard?` (false) | The `xln pull --json` result: files written, report, notices. Refused, with the unbuilt edits listed, while the project has source edits not built yet |
| `xln_build_plan` | no | `workbook`, `project?` | The `xln build --dry-run --json` result: `changeSet`, `conflicts`, `problems`, `excelChanges`, `status` (`refused` is a result here, with its reasons), plus `willWrite`, `excelOpen` and `planId` |
| `xln_build` | workbook, backup, lockfile | `workbook`, `project?`, `out?`, `confirm` (must be `true`), `planId?` | The `xln build --json` result: `written`, `backup`, `readBack`, project files updated |
| `xln_lib_status` | no | `path` (workbook or project folder), `lib?` (default the project's `xln.config.json` `"library"`), `diffs?` (true) | The `xln lib status --json` result: per library function identical, outdated, modified, both changed, differs, missing; local only; diffs |

The server's instructions give an agent the workflow: check, pull, edit
`names/*.xln`, check the project, plan, show the plan to the user, then build.

## Safety model

- **Roots.** Every path an agent passes, and every path a tool derives from it (the
  default project folder, `out`), must lie under a `--root`. Symbolic links are resolved
  first, so a link inside a root cannot lead out of it. One exception: the library folder
  that a project's `xln.config.json` names is read (never written) wherever it is, as
  `xln lib status` does.
- **Reads by default.** Four of the six tools only read. `xln_pull` writes only the
  project folder. It refuses to replace source edits not built yet unless the call says
  `discard: true`, and the tool description tells the agent to pass that only when the
  user agrees.
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
the roots and through a link. `corpus.test.ts` (with `XLN_CORPUS`) runs check, names,
pull and plan on a copy of every corpus workbook. It also checks that a summary check
stays under 80 kB.
