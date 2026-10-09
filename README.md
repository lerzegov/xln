# xln — Excel names as code

**A better Name Manager, as code.** A tool (`xln`) that reads an existing
Excel workbook, turns its defined names and LAMBDA functions into a text project in
VS Code, audits them, and writes edits back into the file. It works with no add-in, no
macros and no running Excel, on macOS and Windows.

Status: **0.1.3** (first public release 0.1.0, 2026-10-08). The whole loop works: pull a
workbook's names into text, edit and check them in VS Code (desktop or vscode.dev), audit
the workbook, build the names and cell formulas back into the file; plus a LAMBDA library
across workbooks, a command line (`xln`) and an MCP server for agents.

**Install:** the extension is on the
[VS Code Marketplace](https://marketplace.visualstudio.com/items?itemName=lerzegov-xln.xln)
and on [Open VSX](https://open-vsx.org/extension/lerzegov-xln/xln) (search **xln** in the
Extensions view, in desktop VS Code or in vscode.dev).

**Using xln:** the [user guide](docs/USER-GUIDE.md) (for modellers) and the
[`.xln` language specification](docs/LANGUAGE.md) (for contributors and power users).
Coming from Excel Labs' Advanced Formula Environment: [what AFE users asked for](docs/AFE-USERS.md),
request by request, and how to move.

## What is here

| Path | Content |
|---|---|
| [`docs/USER-GUIDE.md`](docs/USER-GUIDE.md) | **User guide:** install, the Excel ↔ editor loop, scope, LAMBDAs, the library, messages, cheat sheet |
| [`docs/AFE-USERS.md`](docs/AFE-USERS.md) | **For AFE (Excel Labs) users:** the requests on Microsoft's trackers, what xln does about each, and moving an AFE workbook to xln |
| [`docs/LANGUAGE.md`](docs/LANGUAGE.md) | **The `.xln` language specification** (v0.2): project layout, grammar, cell statements, scope, stored forms, checks, library format, open issues |
| [`packages/`](packages/) | The code: `core` (pure TypeScript, Node and browser), `cli` (the `xln` command), `vscode` (the extension, desktop and web), `mcp` (an MCP server over the CLI) |
| [`CONTRIBUTING.md`](CONTRIBUTING.md) | **How to contribute:** catalogue entries, rules, workbooks; code conventions and the measured Excel facts the code respects |
| [`probes/`](probes/) | Scripts that measured what Excel allows without an add-in, and their results. See `probes/README.md` |

## Open a workbook folder with the extension loaded

```
npm run open -- ~/Desktop/my-model    # desktop VS Code, no clicks; later just: npm run open
npm run open -- --web                 # VS Code for the Web, local, in Chrome, no clicks
npm run open -- --vscode-dev          # the real vscode.dev; then File > Open Recent > the folder
```

The folder is remembered. `--web` runs the same web bundle as vscode.dev but its file
system keeps changes in memory, so missing projects are pulled to disk first.
`--vscode-dev` is the only route that tests Chrome's real local-folder access; Chrome
lets only a person pick the folder, so that click cannot be automated.

## Try a development build in the browser (vscode.dev, Chrome)

```
npm install
npm run web
```

`npm run web` builds the extension, serves it at `https://localhost:5443` and opens
vscode.dev in Chrome. Keep it running: vscode.dev fetches the extension from it every
session. The first time in a Chrome profile, install it once with *Developer: Install
Extension from Location...* → `https://localhost:5443`. One-time setup per Mac:
`brew install mkcert && mkcert -install`. Details: `packages/vscode/DEVELOPMENT.md`.

## Why it works this way

- **Prior art.** Microsoft's Advanced Formula Environment (Excel Labs) has the right
  shape, but keeps its source inside the workbook, needs an add-in, and has had no
  updates since early 2024 ([AFE users' requests](docs/AFE-USERS.md)).
- **No add-in.** Many organisations, universities included, do not allow add-ins. xln
  works on the workbook file itself: patching the `.xlsx` keeps comments, sheet scope and
  everything it does not touch, and Excel recalculates on opening. What Excel accepts was
  measured in workbooks Excel saved, on Mac and Windows (`probes/`).
- **VS Code and the browser.** A TypeScript core runs in desktop VS Code and in
  vscode.dev, so it works on lab PCs where nothing can be installed.
- **Names, not a new modelling tool.** Excel owns layout and labels; xln owns names and
  the formulas of named cells, and audits the rest.

## Contributing

Most contributions are data, not code. Details, code conventions and the measured Excel
facts: [`CONTRIBUTING.md`](CONTRIBUTING.md).

- **Function catalogue entries:** `packages/core/src/lang/catalogue-data.ts`, one Excel
  function per line (name, stored prefix, argument counts, version); a `?prefix` mark
  where the stored prefix is not yet confirmed in a saved file. Confirming one from a
  workbook Excel saved is a welcome first contribution.
- **Audit rules:** messages, hints and severities are data in
  `packages/core/src/audit/rules.ts`.
- **"Broke on my Excel" cases:** open an issue with the smallest workbook that shows it,
  the Excel version and system, and what you did. Workbooks Excel saved are the evidence
  the tool rests on (`probes/README.md` records every measured fact).
- **Code:** `npm install`, `npm run build`, `npm test`. Read `CONTRIBUTING.md` first:
  the core must run in the browser, and Excel's stored forms are measured, not guessed.

Licence: MIT ([`LICENSE`](LICENSE)).
