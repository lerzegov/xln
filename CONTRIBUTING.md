# Contributing to xln

Most contributions are data, not code:

- **Function catalogue entries:** `packages/core/src/lang/catalogue-data.ts`, one Excel
  function per line (name, stored prefix, argument counts, version); a `?prefix` mark
  where the stored prefix is not yet confirmed in a saved file. Confirming one from a
  workbook Excel saved is a welcome first contribution.
- **Audit rules:** messages, hints and severities are data in
  `packages/core/src/audit/rules.ts`.
- **"Broke on my Excel" cases:** open an issue with the smallest workbook that shows it,
  the Excel version and system, and what you did. Workbooks Excel saved are the evidence
  the tool rests on (`probes/README.md` records every measured fact).

## Code

- npm workspaces: `packages/core` (pure library), `packages/cli` (the `xln` command),
  `packages/vscode` (the extension, desktop and web), `packages/mcp` (an MCP server for
  agents, over the CLI's functions).
- TypeScript, strict, ESM. Build `npm run build`; test `npm test` (Vitest); the
  extension's suites `npm run test:desktop -w xln` and `npm run test:web -w xln`.
  Extension development: `packages/vscode/DEVELOPMENT.md`.
- **`packages/core` must run in the browser** (vscode.dev). No Node built-ins, no
  `Buffer`, no `process`. Bytes in and out are `Uint8Array`, text is `string`.
  `packages/core/test/purity.test.ts` enforces this.
- Runtime dependencies: pure JavaScript, small, no native modules. Name each one with
  its reason in the package README.
- No regular expressions over XML or formula text where structure matters: use a
  tokenizer. (An early probe corrupted a sheet with a regex over `<f>` elements.)
- Comments explain why, not what. Match the surrounding code.

## Excel facts the code must respect

Measured in workbooks Excel saved; details and the probes in `probes/README.md`.

- Stored forms: `_xlfn.`, `_xlfn._xlws.`, `_xlpm.` (LAMBDA/LET parameters), `_xleta.`,
  `_xlop.p` for an optional LAMBDA parameter `[p]` in the parameter list (uses in the body
  stay `_xlpm.p`; brackets in the stored form make Excel drop the name; same in cells and
  nested LAMBDAs, probe F9), `ANCHORARRAY(x)` for `x#`. A modern function without its
  prefix becomes `#NAME?` and Excel re-saves it as `_xludf.` permanently.
- `localSheetId` is the 0-based sheet **position** in `<sheets>`, not `sheetId`.
- A name's own sheet is qualified on entry (`'S1'!$B$1#`).
- Excel re-spaces formulas: compare definitions modulo whitespace.
- Line breaks inside a definition are stored as CR LF; XML parsers normalise to LF.
- Shared formulas: only the master `<f t="shared" ref=… si=…>` carries text; children
  are `<f t="shared" si=…/>`.
- Hidden `_xl*` helper names are not stored in the file; ignore any that appear.
- Cell formulas (F8, Mac): a plain `<f>` has legacy implicit-intersection semantics
  (`SUM(A1:A3*2)` gives `#VALUE!`); dynamic-array form is `cm="<XLDAPR record>"` plus
  `<f t="array" ref=…>`, and `t="array"` without `cm` is a legacy CSE array. Excel
  accepts the dynamic-array form for scalars too.
- A spill anchor written with `ref` = the anchor alone needs the old spill's ghost
  `<c>` cells deleted, or it shows `#SPILL!`.
- After any cell-formula change, drop `calcChain.xml` (a stale one, even pruned, makes
  Excel repair the file) and set `fullCalcOnLoad="1"` (otherwise Excel shows stale
  values when the file carries the current calcId).

## Tests and fixtures

- In-repo fixtures: `probes/results/*.xlsx` (Excel-saved on Mac and Windows), kept byte
  for byte as Excel saved them.
- Tests that need a larger corpus of workbooks read the folder in the `XLN_CORPUS`
  environment variable (workbooks under `*/dist/*.xlsx`) and skip when it is unset.

Licence: MIT ([`LICENSE`](LICENSE)). Contributions are under the same licence.
