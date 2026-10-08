# xln VS Code extension: development notes

*`README.md` is the Marketplace listing; this file is for developers. Conventions:
[`CONTRIBUTING.md`](../../CONTRIBUTING.md).*

A Name Manager as code: pull a workbook's defined names into an `.xln` project, browse,
search and trace them, and build the edits back into the workbook. The same bundle runs
in desktop VS Code and in vscode.dev on a local folder. Only **xln: Build workbook**
writes an `.xlsx`.

For users: [the user guide](../../docs/USER-GUIDE.md); the language: [`docs/LANGUAGE.md`](../../docs/LANGUAGE.md).

One extension, two entry points built from the same source by esbuild:

| Field | Bundle | Host |
|---|---|---|
| `main` | `dist/node/extension.js` | desktop VS Code |
| `browser` | `dist/web/extension.js` | vscode.dev, github.dev, desktop web-worker host |

The web bundle is built with `platform: "browser"` and no polyfills, so a Node built-in
that slips in fails the build. Both bundles include `@xln/core`, compiled from its
TypeScript sources. All file access goes through `vscode.workspace.fs`.

## Features

| Req | What | How to reach it |
|---|---|---|
| A | **xln: Pull workbook**: runs the core's `pullProject` and writes `<workbook>.xln/` beside the `.xlsx`/`.xlsm`. **Every pull is fresh** (decided 2026-10-06; *Pull workbook (fresh)* is gone): `names/**`, the lockfile and the manifest are written as the workbook has them now (`.xln` files it does not write are removed, `xln.config.json` is kept; only files whose text changes are written), so a scope changed in Excel comes through as `@workbook` exactly per the workbook, and a name moved by hand goes back to its file. When the project holds **source edits not built yet** (unsaved editors included), a modal lists them (`names/sheets/Ratios.xln:11  Ratios!ROS: update Ratios!ROS (comment)`; the full list in the *xln* output channel) and offers **Build first** (runs *xln: Build workbook*, then pulls; it stops if the build does not finish, or in the browser, whose build writes a copy and leaves the edits unbuilt), **Discard and pull** (the edits are replaced; unsaved editors of the old files are reverted, so their text does not come back), or Cancel. Files whose text differs only in layout or `//` comments (no edit in them; line endings do not count) are a group of their own in that modal, *Layout or comments only: will be rewritten*; when they are the only difference, a modal asks **Pull anyway** or Cancel (M3e). With neither it pulls without asking. After the pull the editor tabs of files it removed are closed (one with unsaved text stays open, with a warning). The source a build embedded (D5) is never read: it is an archive copy | Explorer context menu on a workbook, or the Command Palette |
| D, E | **xln: Build workbook**: runs the core's `buildWorkbook` on the project and asks before writing (the modal lists the changes). Unsaved editors of the project's files are saved first, each save awaited, and the build plans from the text just saved (feedback 2026-10-07: after *Take the library's version*, a Build wrote nothing and said nothing). **Every Build ends with a message**: *built … : N changes (…)*, *up to date*, *cancelled*, the refusal, or the failure; never silence. **A refused build** shows a modal with its reasons (the first five, then *… and N more*; *Show details* opens the *xln* output, where the reasons come first, before the plan it did not write), and its blocking errors go to the Problems panel at their `file:line` as errors with source `xln build` (apart from the live checks; cleared by the next build, or when the file is edited). **Desktop** overwrites the workbook after the lock-file guard (`~$` re-listed right before writing), keeps `<name>.backup.xlsx`, re-reads what it wrote (restoring the original on a mismatch), and updates `xln.lock.json` and the manifest. It embeds the source in the workbook (D5) only when the project's `xln.config.json` says `"build": {"embed": true}`. **In the browser**, where Excel's `~$` file is invisible, it writes `<name>.xln.xlsx` beside the workbook and leaves the original and the lockfile alone. A name changed both in Excel and in source refuses the build; *Show conflicts* opens each as a diff (Excel's version left, `xln-conflict:` documents). It refuses exactly the source errors the editor shows (M3d: one checker) and what needs the workbook (conflicts, broken references). Nothing prompts about scopes or formulas: the build writes the source. **When Excel has the workbook open** (desktop), it offers **Close in Excel and build**: Excel closes it (asking before discarding unsaved changes), the build writes, and Excel opens it again (as *Build and reopen*). A file with no place in `names/` refuses the build (see *Files with no place in a project* below). The summary goes to the *xln* output channel | Explorer context menu on an `.xln` file or a project folder (or anything in it), editor context menu of an `.xln` file, Command Palette |
| E7 | **xln: Build and reopen in Excel** (desktop only; hidden in vscode.dev): the same build, but desktop Excel first closes the workbook **without saving** and opens it again after writing, through `osascript` on the Mac and a PowerShell COM script on Windows (**untested**). When Excel reports unsaved changes in the workbook it asks before discarding them. The reopen is checked after a pause (the workbook open under its path, Excel answering, no repair log): a repair prompt, a file-access request or a missing workbook is an error with its reason, never a success message. After a verified open Excel comes to the front (M3e; also after *Close in Excel and build*): `activate` on the Mac; on Windows the workbook is activated and Excel's window asked to the foreground, which Windows may answer by flashing the taskbar button. The scripts are the core's (`excelControl`); only the desktop bundle runs them (`src/excelHost.node.ts`, swapped in by `scripts/build.mjs`; the web bundle gets the stub `src/excelHost.ts`) | Explorer context menu on an `.xln` file or a project folder (or anything in it), editor context menu of an `.xln` file, Command Palette |
| B1 | Outline per `.xln` file: names in file order (in a file with old `@scope(Sheet)` / `@workbook` blocks, grouped under their block), symbol kind by name kind (LAMBDA = function, spill = array, …); folding of blocks, multi-line definitions and doc comments; colouring from the TextMate grammar | Outline view, breadcrumbs, *Go to Symbol in Editor* |
| B2 | Workspace symbols over every name; **xln: Search names**: a quick pick matching every word of the query in names, definitions and doc comments | *Go to Symbol in Workspace* (Ctrl/Cmd+T; in a browser, where Chrome takes Cmd+T, use Cmd/Ctrl+P and type `#`); Command Palette |
| B3 | Go to definition of a name used in a definition, resolved like Excel: `Sheet!X`, a bare `X` inside a sheet-scoped definition is that sheet's `X` first, LET/LAMBDA variables shadow names (they go to their declaration), a call spelled like a built-in calls the built-in | F12, Ctrl/Cmd+click |
| B3 | Find references: uses in `.xln` files **and** in the workbook: cells, conditional formats, data validations, Table columns (see *Cell usages* below) | Shift+F12, Shift+Alt+F12; **xln: Show cell usages** |
| D4, G | **Rename Symbol** on a defined name (its statement or a use; M5): the core's `renameInSource`, as `xln rename`: the new name with `@renamed(Old)` above it and the new token in every formula of the project's files that reads it, shown in the refactor preview before it is applied. Refused with a message for a LET/LAMBDA variable, an invalid or taken name, or a formula the new name would make read something else. The next build renames the name in the workbook and rewrites it in the cells (or refuses when a chart, Table column, pivot, hyperlink or form control reads it) | F2 |
| B3 | "Used by" and "uses" among names | *Show Call Hierarchy* (Shift+Alt+H): incoming = used by, outgoing = uses; also in the hover |
| B4 | Hover: kind, scope, arity, hidden flag, doc comment; for a spill name its anchor cell, the anchor formula, the last saved extent and the first cached value; the names it uses and how many names, cells, formats and validations use it | Mouse over a name |
| B6 b | **xln: Formula view (calculation order)**: the same document for one sheet in calculation order (each formula after what it reads, on any sheet and through names; otherwise in order of appearance), a first column with the level (the longest chain of formulas from the inputs; 1 reads only inputs) and `↻n` on the members of circular reference n, which stand together under a comment. **xln: Workbook formula view**: every formula of the workbook in one calculation order, each address with its sheet (`IS!C16#`); a cell reference goes to its line in the same document. **xln: Switch formula view order**: the editor title's button (list-tree icon) reopens a sheet's view in the other order at the same cell (from the workbook view: that cell's sheet in order of appearance) | Explorer context menu on a workbook; editor title or context menu of a formula view; Command Palette |
| B6 | Hover on a formula's address, in either order: kind, extent, value, in calculation order its level and cycle, and **reads** / **read by**: the names it reads or that read it, then the formulas, directly or through a name (`C6# (via AccountsReceivable_base)`), each linked to its line; input cells are listed unlinked; the first 12 of each, then a count | Mouse over an address in a formula view |
| C1 | Problems panel, from the core's checker (M3d: the same the build refuses on): `.xln` syntax and annotations, formula syntax and compile errors at their token, a key defined twice, a sheet the workbook does not have, addresses (bare in a sheet file; the lockfile's; quick fix *Restore the address*), a name Excel would refuse. **Scope** (set in Excel or in the source, no setting since 2026-10-06): a hint, advice only, on a workbook name on a sheet's cell that only that sheet reads (*workbook name on a cell of Ratios, read only on Ratios: remove @workbook to make it local*, quick fix *Remove @workbook*), and nothing when another sheet or a name of another scope reads it (it needs workbook scope); `xln check` lists the same findings; a scope change is noted as the re-scope the build makes (info). An old `"names": {"scope": …}` in `xln.config.json` is ignored (a note in the xln output channel). An old sheet file (`@scope`/`@workbook` blocks) gets an info with the quick fix **Convert to per-name @workbook**, an old module file **Convert to per-name @sheet**. A formula as the workbook already has it keeps its findings as warnings (the build leaves it), unless the source renamed, moved or deleted what it reads (an error, as the build refuses). The spill warnings (*spills over …, but X covers only …*, *'#' on a cell left empty*) and the warnings for a name a pull would move to another file (with quick fixes that move it, possibly into another file) are the checker's too | Problems panel, light bulb |
| M3c 2 | **Completion** inside formulas (after the `=` of a definition or cell statement, in LET/LAMBDA bodies), resolved like Excel: LET/LAMBDA variables in scope first, then the home sheet's local names bare, then workbook names, other sheets' local names as `IS!Sales_base` (quoted when Excel needs it: `'SCF recursive'!…`, `'S2'!…`; matched on the name, so typing `Sal` finds `IS!Sales_base`), module prefixes (`FN.`, which completes again to the module's members), then the catalogue's functions (not the internal ones) with arity and Excel version, parameter names when the catalogue has them. Each item says what the name is (kind, scope, LAMBDA parameters) and shows its doc comment and formula. Nothing completes in strings, comments, structured references, on the left-hand side or in a cell address | Typing (letters), `.` after a module prefix, `!` after a sheet name, Ctrl+Space |
| M3c 3 | **Signature help**: for a project LAMBDA its parameters with the doc comment's summary and `@param` lines; for a LAMBDA bound in a LET in scope its parameters; for a built-in the catalogue's parameter names (`FunctionInfo.params`), else `NAME(min..max args)`. A repeating tail (`...`) keeps the last parameter active | `(` and `,` in a call; Ctrl/Cmd+Shift+Space |
| M3c 4 | **Checks as you type** (source `xln`, on the source text, after a 250 ms pause in typing, every file of the project): a name nothing defines (with *Did you mean …?* for a name one edit or a swap away), another sheet's local name read without its sheet (error, quick fix *Qualify: IS!Sales*), a qualified name its sheet and the workbook lack, an unknown sheet, a call to a function neither the catalogue nor the project knows (the build refuses it), the argument count of project LAMBDAs (error) and of built-ins (warning, as the audit), a value called like a function; a cell called like a function (`C2(D2, E2)`, the cell holding a LAMBDA) or a name on such a cell is checked against the LAMBDA of its cell statement (argument count an error; a cell holding a value or a value formula a warning), a doc comment over the Name Manager's 255 characters (an error: the build refuses it); a name written in another case than its definition (`EbIT` for `EBIT`, which Excel accepts) is a hint with the fix *Match the name's spelling*. A problem the audit already reports (same rule at the same place, or on the same cell's line of the formula view) is left to the audit's diagnostic, which then offers the same quick fixes | Problems panel, light bulb |
| M3c 5 | **LAMBDA doc comments**: `/** Summary. @param x what x is @param [y] … */` above a definition (AFE's `/** */` plus `@param` tags); the hover and signature help show the summary and each parameter. The whole text is the Name Manager comment, so a pull gives it back. **xln: New module** asks for a prefix and writes `names/<Prefix>.xln` with a header and a sample LAMBDA documented that way; nothing reaches the workbook until a build. From an Explorer item it adds to that item's project; otherwise, with several projects, a pick offers the active editor's project first (M3e) | Hover, signature help; the Explorer's (+) button (when the workspace has a project), Explorer context menu of a project folder, its `names` folder or an `.xln` file; Command Palette |
| D7, D8 | **Cell statements** in `names/sheets/<Sheet>.xln` (`Name @C6 = …;`, slots `Name @C6 = ;`, unnamed `@C5 = …;` and blocks `@B40:G40 = …;`): coloured by the grammar, listed in the Outline in sheet order (`@C5` for an unnamed one), their formulas read like definitions (go to definition, references: a name's references include the cells whose statements read it). The address is bare (the sheet is the file's) and read-only: one that differs from the last pull (the lockfile) is an error, and a named cell's has a quick fix that restores it. Its `#` is not: `Name @C6#` puts the name on the spill, `Name @C6` on the cell, and the build writes a change of it as a name change; a named statement without `#` on a formula that spilled beyond its cell when last saved gets a warning (`Sales: the formula spills over C3:G3, but Sales covers only C3`) with the quick fix *Name the whole spill: @C3#* (not in a project pulled before the `#` was written, lockfile format 3, where the build reads a missing `#` on a name on the spill as `#`). **xln: Build workbook** lists cell changes (`set formula of BS!C8`) and shows a cell conflict as a diff of the two statements. A workbook-scoped name on a sheet's cell (Create from Selection) has `@workbook` on the line above it; removing that line makes it local, as a build's `rescope-name` (the *Make local* prompt is gone: nothing prompts during a build) | `.xln` editor; Problems panel (light bulb); the build command |
| C | **The audit** (checks C1–C15 of `@xln/core`; C15 with the workbook's cell values, read once per workbook version): every pulled project's workbook is audited when the projects load (reload, pull, file changes) and its findings go to the Problems panel, source `xln check`, code the rule (`C2.bare-prefix`), the hint in brackets; a name finding sits on its entry in the `.xln` files, on the part of the formula it is about when it is found there (`SEQUENCE`, `'S1'!#REF!`, `0.27`), else on the name; a cell finding on the cell's line of its sheet's formula view (`xln-formulas:`, rendered when the problem is opened); formats, validations, Table columns, charts and names the project no longer has on the workbook file. **xln: Audit workbook** opens a read-only report `xln-audit:/<workbook> (audit)` (the CLI's text: counts per check, findings by check, name and spill census) whose names go to their entry and whose cells open the formula view at that cell; it also audits a workbook that was never pulled (its cell findings then sit on the formula view, its name findings on the workbook). A project's `xln.config.json` applies (harness, rule severities, C13 constants, as for `xln check`; problems in it go to the output channel and are marked on their key in the file: a warning for a key that is unknown or in the wrong section, saying where it belongs, e.g. `audit.library: \`library\` is a top-level setting, not an audit setting: move it out of "audit"`); a pull writes the default one when there is none and a replacing pull keeps it | Explorer context menu on a workbook, editor context menu of an `.xln` file, Command Palette |
| B6 a | **xln: Formula view**: a read-only document per sheet, its cell formulas in order of appearance (row by row, left to right) with the names defined as each cell (its left-hand side: the cell, its spill `C6#` or its saved extent), kind, saved extent of spills, display form and saved value; names in a formula and on the left go to their entry in the project (F12, Ctrl/Cmd+click), hover and Find References as in `.xln` files; a cell reference goes to that cell's line; the Outline lists the entries as `Name — C6#` (the address alone when unnamed), so its filter finds a block by name; colouring from its own grammar (`syntaxes/xln-formulas.tmLanguage.json`, formula tokens from the `.xln` grammar) | Explorer context menu on a workbook (asks for the sheet); editor title (table icon) or context menu of `names/sheets/<Sheet>.xln` (that sheet); F12 on the address of a cell usage line, or the command with the cursor on one (opens at that cell); Command Palette |

**Menus (M3e).** The Explorer's context menu offers on a workbook (`.xlsx`, `.xlsm`):
*Pull workbook*, *Formula view*, *Formula view (calculation order)*, *Workbook formula
view*, *Audit workbook*; on an `.xln` file, a project folder (`*.xln`) or anything in one:
*Build workbook*, *Build and reopen in Excel* (desktop), *New module*. No build on a
workbook, no pull on an `.xln` file; the Command Palette keeps every command and resolves
its target as before. *Formula view* stays on a sheet file's editor title. The Explorer's
title bar has a (+) **New module** button when the workspace holds a project (context key
`xln.hasProject`). (*Inspect workbook* and *Write test file*, the W3 probe commands, are
unchanged.)

**The output log (feedback 2026-10-07).** Every command and action writes to the *xln*
output channel through `src/log.ts`: a header `HH:MM:SS xln <verb> <target>…` after a
blank line, then details indented by two spaces. Any line starting `xln ` is a header and
gets the time, so the shared formatters (`formatPullSummary`, `formatBuildSummary`, the
audit's and the library status's first lines from the core's renderers) print as in the
CLI. An `Activity`'s `info`/`warn`/`error` log the message, then show it, so early
returns and refusals leave a trace. The project-load line (`xln loaded …`) and the
workbook-read line are written only when the counts change, the load or read took over
200 ms, a project has load errors, or on *xln: Reload*: a load follows every save. Tests
read the lines through the API's `logLines()` (an OutputChannel cannot be read back). At
activation a second copy of the extension under another publisher id (the test builds'
`xln.xln`) gets a warning: VS Code runs both, so every menu entry shows twice.

**Project internals hidden and read-only (M3e).** The extension's configuration defaults
put `**/*.xln/workbook.manifest.json`, `**/*.xln/xln.lock.json` and `**/*.xln/library-bases`
(the library bases' texts, below) in `files.exclude`
(not in the Explorer) and in `files.readonlyInclude` (read-only in the editor);
`xln.config.json` stays visible and editable. **xln: Show project internals** writes
`false` for those patterns into the workspace's `files.exclude`; **xln: Hide project
internals** removes them again (the default applies). Workspace settings are written
only by these commands. The extension reads and writes the files through
`workspace.fs`, which neither setting restricts. Activation does not depend on the
hidden lockfile: it also fires on `**/*.xln/xln.config.json` and `**/*.xln/names/*.xln`.

**Files with no place in a project (M3e).** VS Code cannot refuse a file created in the
Explorer, so one created or renamed into a project's `names/` that has no place there gets
a warning at once (with *Delete this file*), an error diagnostic on the file (quick fix
*Delete this file*), and the build refuses it with the same message (the core checker's
`strayFile`): anything not `.xln` (hidden files such as `.DS_Store` aside), *there is no
sheet Foo in the workbook: sheets are created in Excel, then pulled* for a sheet file of
a sheet the workbook lacks, a folder other than `names/sheets/`, a module file not named
after a prefix (`my module.xln`; letters, digits and `_`, `~2` as a pull adds on a clash).

**Stale tabs (M3e).** When the projects load, clean tabs of files that no longer exist
below a project are closed; a pull does the same for the files it removed, and warns
about a tab with unsaved text, which it leaves open.

**Suggestions come from xln only.** The extension sets `"[xln]": {"editor.wordBasedSuggestions": "off"}`
(a configuration default): VS Code's word-based suggestions would otherwise offer any word
of the open files (`scope`, `workbook`, sheet names) where xln has nothing to complete.
No xln completion, snippet or quick fix inserts `@scope(…)`: the proposals the author saw
in sheet files (M3d finding) came from word-based suggestions or an AI completion such as
Copilot repeating the old files' blocks; no file a pull writes has them since 2026-10-07.

**Cell usages: why a virtual document.** A reference result needs a document and a
range, but cells live in the workbook. Each name gets a read-only document
`xln-cells:/<workbook>/<key> (usages)` with one line per place, e.g.
`'S1'!B3:B7: =A3*Rate  → 0.2  (formula of B3)`, and Find All
References returns those lines next to the uses in `.xln` files. One gesture then shows
every usage, in names and in cells, in VS Code's own references view, peek and search
editor, with the same keys in desktop and browser; there is no custom view to learn or
to keep in sync. The places come from `workbook.manifest.json` (the last pull, cells
merged into rectangles); the formula text and cached value are read from the workbook
beside the project when it is there (read once, kept while its size and date stay the
same), so the lines still list the places when the workbook has moved. The alternative,
a tree view "xln: Usages", would duplicate the references view and still need a
document to show a formula.

**Formula view: the same design.** `xln-formulas:/<workbook>/<Sheet> (formulas)?<workbook
URI>` is rendered by the core (`sheetFormulaView`, `renderFormulaView`), which also
returns where each name and reference sits in the text; the providers on that scheme
look up the offset under the cursor. Names resolve against the project's live
definitions when the workbook has been pulled (otherwise against the workbook's own
names, and the document says they are not linked). The workbook is read through
`XlnWorkspace.workbookAt` and kept while its size and date stay the same; the view is
rebuilt when the workbook changes or the projects reload. A sample (from `lbo-ep03r.xlsx`):

```
// Sheet BS: 96 formulas in order of appearance (row by row, left to right). Read-only view of lbo-ep03r.xlsx.
                                  C5#    (1×5)  = Model!Years                                                     → 2022 …
AccountsReceivable_base           C6#    (1×5)  = IN.ASMPT_base("AR days") * IS!Sales_base / 365                  → 12328.76712 …

CapitalExpenditures_base          C13#   (1×5)  = FN.SEEDROW(                                                     → 0 …
                                                      0,
                                                      IN.ASMPT_base("capex pct of sales change") * (IS!Sales_base - FN.PREV(IS!Sales_base))
                                                  )
```

In calculation order the views are `…/<Sheet> (calculation order)` and
`…/[workbook] (calculation order)` (`[` cannot occur in a sheet name). The dependency
graph (`buildGraph`, with the project's names when pulled) is built once per workbook
snapshot and project generation and shared by every view and hover of that workbook.
The levels show where a line sits in the chain; a line that moved ahead of the order of
appearance waits for something below it. From `lbo-ep03r.xlsx`:

```
// Sheet BS: 96 formulas in calculation order (each formula after what it reads). Read-only view of lbo-ep03r.xlsx.
2                                     C5#    (1×5)  = Model!Years                                         → 2022 …
5   AccountsReceivable_base           C6#    (1×5)  = IN.ASMPT_base("AR days") * IS!Sales_base / 365      → 12328.76712 …
…
2   DeltaFinancialDebt_base           C33#   (1×5)  = FN.DELTA(FinancialDebt_base)                        → #N/A …

16  Dividends_base                    C23#   (1×5)  = 'SCF recursive'!Dividends_base                      → 0 …
```

and in the workbook view of `lbo-ep03r-circ.xlsx`, the interest ↔ cash loop:

```
// ↻1: circular reference, 16 formulas here that depend on each other (Excel iterates them or reports it)
17 ↻1  InterestIncomeOnExcessCash_base   IS!C16#                (1×5)   = FN.SEEDROW(                  → 0 …
```

**Library (M4).** With `"library": "<folder of .lambda files>"` in the project's
`xln.config.json` (relative to the project folder; absolute or `~/…` on the desktop):
**xln: Library status** opens a read-only report `xln-lib:/<project> (library)`, the
CLI's text, with each library function identical, outdated, modified, both changed,
differs or missing, three-way on the copy's **library base**: `@from(lib #353921)` above
the name, the library version it came from, written only by Insert, Take, Publish and Record library base
(the build carries it in the Name Manager tag as `lib#353921`, pull writes it back;
hovering it says which library version it records; delete it and the copy has no base)
(and LAMBDAs of a library module the library lacks; other modules' are counted), diffs included; names link to the project, missing
ones to the library file. The report is of the source, so each function the workbook does
not have as the source does is marked *not built yet* there, in its summary line and on
its code lens (`library: identical · not built yet`): a status never looks done while the
workbook is not. It is a document rather than a tree view because the diffs are
text, it works the same in vscode.dev, and it matches the audit report. Completion in a
formula offers the library's functions the project lacks, marked *from library* (after
`FN.`; a module only the library has is offered as a prefix); accepting one inserts the
call and adds its definition with its doc comment (and the library functions it calls) to
`names/FN.xln`, created if missing, with its `@from`, as an unsaved edit: the build writes it.
**xln: Insert library function** does the same from a quick pick. On a module file's
entries, code lenses show the library state (the tooltip gives the base and library
versions) and what fits it: **outdated**: *Take the library's version* (replaces the
definition and doc comment in the source and records the base); **modified**: *Publish to
library*, and *Take the library's version (undo the edit)*, which first asks *Discard your
edit of FN.SPREAD?*; **both changed**: the diff opens as the three versions (what each side
changed since the base, when the workbook still has the base's text), and Take asks the
same question, naming the library's change too; no one-click Publish; **differs** (no
base): Publish, and Take after *FN.X has no library base: taking the library's version
replaces whatever this copy has. Continue?* (a local edit cannot be ruled out); **local only**: Publish. Clicking the state label opens
the diff. *Publish to library* opens the library file against what it would become,
warns when it would overwrite a library change the copy does not have, and writes it only
after a modal confirmation; the entry then records the published version as its base (an
unsaved edit). **identical** with no `@from` (inserted or published before the base
existed; it would read *differs* after the first change on either side): the lens reads
`library: identical · no base` and offers **Record library base**, which writes
`@from(lib #…)` (the library's version) as an unsaved edit; also a quick fix and the
command *xln: Record library base* at the cursor. The same actions are quick fixes on the
entry. Nothing is automatic: the status only says *no base recorded: Record library base*.

**The base's text** (author's idea, 2026-10-07). Insert, Take, Publish and Record library
base also keep the base definition in the project, `library-bases/<hash>.json` (stored and
display text, the library file, the hash), written at once: internal, hidden and read-only
like the lockfile. On **both changed** the three-way document then shows base → copy and
base → library from it (`the base #…'s text: from library-bases/….json`); without it,
from the workbook as before, else the copy against the library with the note. A pull
leaves the folder alone; nothing prunes it (an unreferenced file is a few hundred bytes).

**`@param` drift** (author's idea, 2026-10-07), in the shared checker, so live and in
`xln check` alike: an `@param x` in a LAMBDA's doc comment with no parameter `x` is a
warning (code `doc-param`) with the quick fixes *Rename @param x to y* (the parameter at
the same position, when that one is not documented already) and *Remove @param x*;
parameters left out while others are documented are a hint (`doc-param-missing`).
`@param [p]` and `@param p` both document an optional `[p]`.

**Live model.** The workspace model loads every project folder (a folder holding
`xln.lock.json`, `workbook.manifest.json` and `names/`), parses all `.xln` files with the
core's `parseModule` and each definition with the core's formula parser. Navigation
among names follows the text as edited (open editors feed in their unsaved text); cell
usages come from the manifest, so they reflect the last pull. The model is rebuilt on
file-watcher events, on save of an `.xln` file, after a pull, and on **xln: Reload**:
vscode.dev has no file watching for a local folder.

## Code

| File | What |
|---|---|
| `src/model/project.ts` | The project model, free of the vscode API: resolution, references, used by / uses, outline, folding, search, problems (offsets in, offsets out) |
| `src/model/usages.ts` | Text of the `xln-cells:` documents and of the hover |
| `src/model/formulaView.ts` | The formula view of a sheet for the editor: build it with the project's names, find the name, reference or cell entry at an offset, the entry of a cell |
| `src/formulaView.ts` | The `xln-formulas:` documents, their providers, the `xln-cells:` go-to-cell, and **xln: Formula view** |
| `src/model/manifest.ts`, `pull.ts`, `lines.ts` | Manifest reading, pull helpers, offset ↔ line/character |
| `src/xlnWorkspace.ts` | Finds and loads project folders through `workspace.fs`, keeps them current, reads the workbook lazily |
| `src/features.ts` | The VS Code providers on top of the model; diagnostics (problems and checks as you type) and their quick fixes |
| `src/model/editor.ts`, `src/editor.ts` | Writing help (M3c), vscode-free: the cursor's place in a formula, names in scope, completion lists, signature help, the checks as you type, the new module's text; the completion and signature help providers and **xln: New module** |
| `src/commands.ts` | Pull and search commands |
| `src/model/ux.ts`, `src/ux.ts` | M3e housekeeping (vscode-free part: which project an Explorer item or a pick means, stale tabs, the internals' `files.exclude` patterns); the `xln.hasProject` and `xln.internalsShown` context keys, closing stale tabs, files with no place in a project (warning, diagnostic, *Delete this file*), **xln: Show / Hide project internals** |
| `src/model/build.ts`, `src/build.ts` | Build target (desktop or browser), summary and conflict texts (vscode-free); **xln: Build workbook** and the conflict diffs |
| `src/model/library.ts`, `src/library.ts` | The library (M4), vscode-free: its location, completion items, entry states; the `xln-lib:` report, the code lenses and quick fixes, insert, take, diff, publish and record library base, the kept bases (`library-bases/`) |
| `src/model/audit.ts`, `src/audit.ts` | Where a finding lands in a project (vscode-free); the audit diagnostics, the `xln-audit:` report, its links and **xln: Audit workbook** |
| `src/extension.ts` | Activation; also the W3 browser-probe commands **xln: Inspect workbook** and **xln: Write test file** (`src/inspect.ts`) |
| `syntaxes/`, `language-configuration.json` | The `.xln` grammar (`samples/demo.xln` shows every construct) and the formula view's |

## Tests

- Vitest (root `npm test`): `test/model.test.ts` (resolution with sheet scope and
  shadowing, comments inside definitions, references, outline, folding, search,
  problems, usages, hover and the formula view model on `f7_base.xlsx`), `test/corpus.test.ts` (with
  `XLN_CORPUS`: every corpus workbook loads without problems and the live "uses" / "used
  by" equal the manifest's; the checks as you type find no error on any of them; completion
  in `lbo-ep03r`'s `BS.xln` lists `IS!Sales_base`), `test/grammar.test.ts` (both grammars), `test/inspect.test.ts`,
  `test/audit.test.ts` (where the findings of `probes/fixtures/traps.xlsx` land in its pulled project),
  `test/editor.test.ts` (M3c on a hand-made project: scope and order of completion, module
  and sheet prefixes, no completion in strings, comments, left-hand sides and addresses,
  signature help for LAMBDAs, built-ins with and without parameter names and LET-bound
  LAMBDAs, every check as you type with its fixes, `@param` in the hover, the new module),
  `test/library.test.ts` (M4: library items in completion, the library's location, the
  states the code lenses show),
  `test/menus.test.ts` (M3e: the `when` clauses of package.json evaluated for a workbook,
  an `.xln` file, a project folder, files and folders in one, in the browser; the (+)
  button; the internals' configuration defaults and palette entries), `test/ux.test.ts`
  (stale tabs, the project an Explorer item or a pick means, Show / Hide internals).
- `test:desktop` and `test:web` run `src/test/probe.suite.ts` (W3) and
  `src/test/browse.suite.ts`, which pulls `probes/results/f7_base.xlsx` in a fixture
  folder and checks pull, replace, a cancelled pull over an unbuilt edit, a pull that writes nothing, a pull closing the tab of a file it removed, a pull over a `//` comment alone (Cancel, Pull anyway), the internals hidden and read-only while the pull still writes them, Show / Hide project internals, outline, definition, references with a cell usage,
  call hierarchy, hover, search, workspace symbols, the formula view (a line, a name
  link, the name on the left of `E1#` (link, hover, references), a reference to a cell,
  hover, outline, references, from a cell usage to the view,
  from `names/sheets/S2.xln`) and diagnostics; then `src/test/editor.suite.ts` (M3c in the
  extension host: completion, signature help, checks as you type following an unsaved
  quick fix, **xln: New module**); then `src/test/library.suite.ts` (M4: a library folder
  written into the fixture, the status report, completion from the library and its insert,
  the code lenses, take the library's version, publish, the kept bases, Record library base); then `src/test/audit.suite.ts`, which pulls
  `probes/fixtures/traps.xlsx` (the seeded-trap workbook) and checks that its findings appear
  as diagnostics (on `_unmanaged.xln` and on formula view lines) and that **xln: Audit
  workbook** opens the report with links to names and cells; then `src/test/build.suite.ts`
  (build, Build first, a conflict as a diff, a file with no place in `names/` flagged and
  refused; on the desktop, with Excel's owner file present, *Close in Excel and build*
  declined writes nothing: the test answers the modal, `closeInExcel: false`, and never
  drives Excel).
- `test:vscode-dev` drives the real vscode.dev: install, open folder, the probe
  commands, then pull, outline, go to definition, hover and cell usages through the UI.
- `--smoke <workbook.xlsx>` after `test:desktop` or `test:web` runs a timing suite on a
  copy of that workbook (the original is only read), including the formula view of
  every sheet.

## Scripts (`npm run <script> -w xln` from the repo root)

| Script | What |
|---|---|
| `build` / `watch` | both bundles plus the two test bundles |
| `test:web` | suites in headless Chromium via `@vscode/test-web` |
| `test:desktop` | suites in desktop VS Code via `@vscode/test-electron` (downloads VS Code once) |
| `test:vscode-dev` | headless smoke test against the real vscode.dev (needs network) |
| `serve` | serve this folder at `https://localhost:5443` for *Install Extension from Location* (needs mkcert) |
| `package` | `.vsix` via vsce (writes `xln-<version>.vsix` here; `npx vsce package --no-dependencies -o /tmp/x.vsix` for a test package) |
| `icon` | renders `media/icon.svg` to `media/icon.png` (128×128) with Playwright's Chromium |

## Dependencies

Runtime (bundled):
- **@xln/core**: the workspace's own core library (pull, parser, module reader).
- **fflate**: unzip `.xlsx` in pure JavaScript (the probe commands use it directly; the
  core uses it too).

Development only:
- **esbuild**: bundles the two entry points.
- **@types/vscode**: API types, pinned to the minimum engine (`^1.110.0`).
- **@vscode/test-web**, **@vscode/test-electron**, **mocha**, **@types/mocha**: run the
  suites in the web and desktop extension hosts.
- **playwright**: drives vscode.dev for `test:vscode-dev` (already a dependency of test-web).
- **vscode-textmate**, **vscode-oniguruma**: tokenize the grammar in tests with
  VS Code's own engine.
- **@vscode/vsce**: packages the `.vsix`.
- **ovsx**: publishes to Open VSX (the release workflow, `.github/workflows/release.yml`).
