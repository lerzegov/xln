# Changelog

## 0.1.2 (2026-10-09)

- Pulling twice in a row no longer refuses on a sheet name over a deleted reference: the
  pull now writes it with its sheet (`IS!#REF!#`), so the second pull finds no edits.
  A project pulled with 0.1.1 that has such a line refuses its next pull once: pull with
  *Discard and pull*, or add the sheet by hand.
- The build plan shows updates of the provenance tag alone as one line ("update the
  provenance tag of 22 module names (comment only)") instead of one line per name.
- The formula view shows a cell the build wrote, and Excel has not calculated since, as
  "not calculated since the build" instead of an empty value; its hover says to open and
  save the workbook in Excel. The spill census counts such cells apart.
- *Audit workbook*: a spill (`x#`) of a deleted reference reads "#REF!: the reference was
  deleted".

## 0.1.1 (2026-10-08)

- *Build and reopen in Excel* opens the workbook in Excel when there is nothing to
  build; before, it only said "up to date".

## 0.1.0 (2026-10-08)

The first public release. It covers the whole loop: pull a workbook's names into text,
edit and check them, build them back into the workbook.

**Pull**
- *xln: Pull workbook* writes `<workbook>.xln/` beside an `.xlsx` or `.xlsm`: one file
  per sheet (`names/sheets/<Sheet>.xln`) with the sheet's named cells and local names in
  sheet order, module files for LAMBDAs and other names, a lockfile and a manifest.
- Every pull is fresh: the project is written as the workbook has it now. When the
  project holds edits not built yet, the pull lists them and offers *Build first* or
  *Discard and pull*.
- Names keep their scope, comment and hidden state; modern functions are shown without
  their stored prefixes (`_xlfn.`, `_xlpm.`, `_xlop.`) and `ANCHORARRAY(x)` as `x#`.

**Edit**
- The `.xln` language (spec in `docs/LANGUAGE.md`, v0.2): named formulas, LAMBDAs with
  doc comments and `@param`, named cell statements (`Sales @C3# = …`), slots, unnamed
  formula cells, `@workbook` per name, modules, `@from` provenance.
- Syntax highlighting, completion of names in scope, LAMBDA variables and Excel
  functions, signature help for LAMBDAs and for the 525 catalogue functions.
- Checks as you type, the same checker the build uses: unknown names and functions,
  argument counts, another sheet's local name read without its sheet, a missing `#` on a
  spill, a workbook name on a cell read only on its sheet; quick fixes where the edit is
  certain.
- *xln: New module*, *xln: Search names*, *Show cell usages*, *Reveal name*.

**Check**
- *xln: Audit workbook*: checks C1–C13 (syntax, stored prefixes, names shadowing
  built-ins, broken references, unqualified reads, LAMBDA arity, formula length and
  nesting limits, fixed references into spills, unused names, copy drift, name cycles,
  constants in LAMBDA bodies), the name census and the spill census. Findings go to the
  Problems panel. Severities and the check harness are set in `xln.config.json`.
- An unused name on cells (a result shown on the sheet) is reported as info; unused
  constants and LAMBDAs as warnings.
- C15, info: a name whose label cell no longer gives it. After a rename of a name made
  with Create from Selection (`Gross_income` → `Gross_ind_income`), the label in Excel
  still reads "Gross income"; xln never writes cell values, so the audit says which cell
  to edit and the probable old name. Only where the names beside it match their labels.
  Its hint gives the Find & Replace that fixes the label in Excel.
- C14, info: the module store of Excel Labs' Advanced Formula Environment (AFE), and each
  name whose AFE text differs from the workbook's, with both texts. Info, not a warning:
  the build warns when it changes a name AFE's modules define; `xln apply` says nothing.
- `xlm-name`, a warning: a new LAMBDA named like an Excel 4.0 macro function that is not
  a worksheet function (`Group`, `Get.Cell`, `Evaluate`, `Files`; the list is [MS-XLS]
  Ftab): Excel may call the macro function or refuse the name (AFE issue #10, not
  measured). The command equivalents (Cetab: `Open`, `Save`, `Copy`, `Table`, `Scale`)
  get no warning: they are ordinary words. `xln rename` to such a name prints the warning and renames.

**Views**
- Formula view of a sheet or the workbook: the names on each cell, in order of
  appearance or of calculation, with spill extents and links to definitions.

**Build**
- *xln: Build workbook* writes edited names and the formulas of existing cells back into
  the `.xlsx`, after a dialog listing every change. Only the changed parts are rewritten;
  everything else is copied byte for byte.
- Safety: Excel's lock file stops a desktop build (or *Close in Excel and build*); a
  backup of the previous file; a read-back that restores the original on any mismatch;
  a name or cell changed in both Excel and xln stops the build and shows both versions.
- *xln: Build and reopen in Excel* (desktop).
- A build that writes a rename removes its `@renamed(Old)` line from the source (saving
  an open editor), the only change a build makes to `.xln` files: the next pull then
  changes nothing. A refused build and the browser's copy leave it.
- An `@renamed(Old)` whose rename is in the workbook already (left by an older build, a
  build in the browser, another tool, or written by hand) is a faded hint,
  `renamed-built`: *@renamed(Old): the rename is built; this line can go*, with the quick
  fix *Remove @renamed(Old)*; `xln check` lists it. The next build that writes the
  workbook removes it too, with the same output line as for its own renames; a build
  with nothing to write leaves the source alone.
- The label notice: after a build that renamed a name on cells, the output lists the text
  cells in its rows and columns that still read the old name ("Gross income",
  "Gross_income") and the Find & Replace that fixes them in Excel (one Find what /
  Replace with pair per text, spaces kept; Within Sheet, Match entire cell contents);
  cells that only resemble the old name are listed to check by eye. The toast offers
  *Show labels to fix* (*xln: Show labels to fix after a rename*): the notice in a
  document, and a pick that copies each text. `xln rename` gives the same notice for
  after the build, `xln build` and MCP `xln_build` print it.
- `@renamed('Cash Flow'!Old)`, a quoted sheet in the argument, parses (it said
  "'@renamed(' is not closed").
- In vscode.dev the build writes `<workbook>.xln.xlsx` and leaves the workbook alone.
- `xln verify` warns first when a side's values are not Excel's: a file never saved by
  Excel (Python-written, or built by xln and not opened since: `original.xlsx has no
  values Excel calculated (it was never saved by Excel): open and save it in Excel first,
  or the comparison is empty`), or one with formula cells a tool left without a value.
  Also `warnings`, `beforeValues`, `afterValues` in `--json`.

**Library**
- LAMBDA library folders: *Library status* (identical, outdated, modified, both
  changed), *Insert library function*, *Take the library's version*, *Show diff with
  the library*, *Publish to library*, *Record library base*; provenance as
  `@from(lib #hash)`.

**Platforms**
- One extension for desktop VS Code (macOS, Windows) and vscode.dev (Edge, Chrome),
  built from the same source. No add-in, nothing installed in Excel.
- The *xln* output panel logs every action, in the CLI's style: a header with the time
  (`14:02:31 xln build book.xlsx: 2 changes`), then the details: pull, build (target,
  changes, read-back, `@renamed` removed, label notice, refusal reasons, timing), Build
  and reopen, rename (F2, or why it was refused), audit counts, formula view, library
  status/insert/take/diff/publish/record base, new module, reload, quick fixes that edit
  other files, and every message shown on screen. The project-load line is written only
  when the counts change, the load is slow, or on *xln: Reload*.
- A warning when another copy of the extension is installed under another publisher id
  (the test builds' `xln.xln`): with two copies every menu entry and button shows twice.

**Documentation**
- `docs/AFE-USERS.md`: for users of Excel Labs' Advanced Formula Environment, the
  requests on Microsoft's trackers and what xln does about each; moving an AFE workbook.
