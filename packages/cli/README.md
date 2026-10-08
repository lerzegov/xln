# @xln/cli

The `xln` command: a thin Node wrapper around `@xln/core`. It reads and writes files
and prints; everything else happens in the core, which the VS Code extension shares.

```
xln pull <workbook.xlsx> [--out <dir>] [--json] [--width <n>] [--discard]
xln formulas <workbook.xlsx> [--sheet <name>] [--order appearance|calculation] [--workbook] [--json]
xln graph <workbook.xlsx> [--json]
xln check <workbook.xlsx | project> [--json] [--only C2,C9] [--severity error|warning|info|hint] [--census-exclude <glob,...>] [--config <file>]
xln build <workbook.xlsx> [--project <dir>] [--out <file>] [--dry-run] [--force] [--embed | --no-embed] [--no-tags] [--reopen [--discard]] [--json]
xln apply <workbook.xlsx> <changes.json> [--out <file>] [--dry-run] [--reopen [--discard]] [--json]
xln verify <workbook.xlsx> [--before <file>] [--tolerance <x>] [--json]
xln rename <project | workbook.xlsx> <Old> <New> [--dry-run] [--json]
xln lib status <workbook.xlsx | project> [--lib <dir>] [--json] [--no-diff]
xln lib publish <project> <Name> [--lib <dir>] [--dry-run] [--json]
xln lib take <project> <Name> [--lib <dir>] [--dry-run] [--discard] [--json]
xln lib base <project> <Name | --all> [--lib <dir>] [--dry-run] [--json]
```

`pull` only reads the workbook. It writes a project folder, by default `<workbook>.xln`
next to the workbook:

| Path | Content |
|---|---|
| `names/<Module>.xln` | names of one module (`FN.*`, `IN_*`), AFE syntax; its sheet-scoped members follow, each with `@sheet(Sheet)` above it |
| `names/_unmanaged.xln` | workbook-scoped names no module owns |
| `names/sheets/<Sheet>.xln` | one file per sheet: its **cell statements** in sheet order (below), then its local names that are not cell statements and that no module owns. Every name in it is local to the sheet unless `@workbook` is on the line above it |
| `workbook.manifest.json` | sheets, Tables, spill map, built-in names, other parts, where each name is used, and where each cell statement is |
| `xln.lock.json` | per name: hash of the stored definition (modulo whitespace and the spelling of numbers), of the comment, hidden flag; per cell statement: sheet, range and hash of its formula. One line per entry; format `xln.lock/4` (below) |
| `xln.config.json` | the project's audit settings (below); written only when the project has none, never overwritten |

**Scope** (M3d, decided 2026-10-06). In a sheet file the sheet is the file's: every name
is local to it unless the line above it says `@workbook` (an annotation like `@hidden`, for
the next statement only; both may stack), and addresses are bare (`@C3#`). In a module
file names are workbook-scoped unless the line above says `@sheet(Sheet)` (per name,
decided 2026-10-07). Sheet names that are not valid file names are percent-encoded
(`a<b` → `a%3Cb.xln`); a file name that cannot say its sheet (a `~2` added on a clash)
carries `@sheet(Name)` on a line of its own. Files written before then (module files with
`@scope(Sheet)` blocks, until the next `@scope(…)` or `@workbook`; sheet files before M3d
with blocks) are still read as they were; the extension offers *Convert to per-name
@sheet* / *@workbook*; a build never rewrites them, and the next pull writes the new form.
A sheet in `@sheet(…)`, `@scope(…)`, an address or a formula is quoted by Excel's rule
(`BS!`, `'S1'!`, `'SCF recursive'!`).

Scope is set in Excel or in the source, whichever is being edited (one at a time; the
author's decision, 2026-10-06): adding or removing `@workbook` is a scope change the build
applies (`rescope-name`, refused when a reference would break, D4), and a scope changed in
Excel's Name Manager (delete, then create again in the other scope) comes in with the next
pull. A workbook name on a sheet's cell that only its own sheet reads is a hint, advice
only (remove `@workbook` to make it local); when another sheet's formula or a name of
another scope reads it, it needs workbook scope and there is no finding. The M3d setting `"names": {"scope": …}` is gone: an old `xln.config.json`
that has it is read, and `xln check` notes that it is ignored.

**Cell statements** (M3b, D7/D8): every formula cell of a sheet appears in its sheet's
file, row by row and left to right, as `xln formulas` lists them:

```
@C5 = Model!Years;                        // an unnamed formula cell
/** Receivables at the end of the year. */
AccountsReceivable_base @C6 = IN.ASMPT_base("AR days") * IS!Sales_base / 365;   // a named cell
@B40:G40 = SUM(B30:B39);                  // a block filled across: the top-left cell's formula
@workbook
Revenue @C41 = ;                          // a slot: a workbook-scoped name on an empty cell
@C42 = Revenue * 2;

// Other names on BS.

Depreciation_base = CHOOSEROWS(FixedAssetRoll_base, 1);
@workbook
Opening = BS!$C$12:$G$12;                 // a workbook name fixed to BS's cells
```

- **Other names on the sheet**: its local names that no module owns, then the workbook
  names no module owns whose definition is fixed to its cells (one absolute cell, range
  or spill of that sheet: a cell inside a spill, a value cell, a range), each with
  `@workbook` above it and Excel's text (decided 2026-10-07). `_unmanaged.xln` keeps the
  other workbook names. Removing `@workbook` makes such a name local (`rescope-name`); the
  hint *Remove @workbook* applies as for slots.
- **Pull notes** a name *Create from Selection* took from a computed cell's current value
  (the cell just left of a row or above a column holds a formula, or lies in a spill) or
  from the corner of a selection with both Top row and Left column: `bullet04: named after
  the current value of Mortgage!A12, a formula result; the name stays as is when that value
  changes`. Info only (`report.valueLabels` in `--json`).

- A **named cell** is a name defined as exactly one cell of the sheet, `$C$6` or
  `$C$6#`, scoped to the sheet or to the workbook; its doc comment and `@hidden` stay
  above it. A **slot** is such a name on an empty cell (no formula, no value). Several
  names on one cell: the sheet-scoped one, then the `#` one, then the first by name is
  the statement; the others stay ordinary names.
- An **unnamed** cell is any other formula cell. Cells holding one formula filled across
  (a shared group, or equal after shifting relative references) form one statement over
  their range: row runs first, then runs over the same columns on consecutive rows into
  rectangles. A spill anchor is one cell; the cells it spilled into are not statements,
  nor are data tables.
- A workbook-scoped name on a sheet's cell (what *Create from Selection* makes) is
  written at its place with `@workbook` on the line above it. To make it local, remove
  that line and build (`--rescope-slots` is gone: no step changes a scope on its own).
- Module-prefixed names (`FN.*`, `IN_*`) that are cell statements live in their sheet's
  file, not in their module's: a cell's formula is read with the sheet.
- The address is set in Excel and read-only; edit the formula after `=`. One part of it
  is yours (decided 2026-10-05): a named statement's `#` says what the name covers.
  `Sales @C3# = SEQUENCE(1,5);` puts `Sales` on the spill (`IS!$C$3#`, stored
  `_xlfn.ANCHORARRAY(IS!$C$3)`), `Sales @C3 = …;` on the cell alone (`IS!$C$3`).
  A pull writes `#` exactly when the name has it, so a name on a 1×1 dynamic array
  defined as the plain cell is written without. Unnamed statements and blocks take no `#`.

**Every pull is fresh** (decided 2026-10-06, replacing the merge of 2026-10-05): there
is no live link between Excel and the editor, so the workbook is the only real state
when you pull. A pull writes `names/**`, the lockfile and the manifest as the workbook
has them now, in a fixed structure (modules by prefix, sheet files by sheet: a name moved
by hand into another file goes back to its place); `.xln` files it does not write are
removed; `xln.config.json` is kept. A scope changed in Excel (the Name Manager cannot
change one: delete the name, create it again in the other scope) comes through as it
is: `@workbook` above a name exactly when the workbook has it at workbook scope. Pull
never reads the source a build embedded (`--embed`, D5): that part is an archive copy for
a workbook handed on.

**The guard.** A pull would replace source edits not built yet, so it first compares the
project as the build does (source against the lockfile and the workbook as it is now;
without a lockfile, against the workbook itself). If there are any, it **refuses**:
exit 1, nothing written, and the list on stderr:

```
xln pull is-model.xlsx: refused: the project has 1 source edit not built yet, which a pull would replace:
  names/sheets/Ratios.xln:11  Ratios!ROS: update Ratios!ROS (comment)
Build them first (xln build), then pull; or pull with --discard to replace them with the workbook's version.
```

A source with errors (the build would refuse it) counts too, with its errors listed.
`--discard` pulls anyway (the summary says how many edits were replaced). Only layout
and `//` comments are not edits: a pull gives the workbook's text back without them. The
files a pull rewrites for that alone (no edit in them; line endings do not count) are
named in a note, `layout or comments only (a pull rewrites them as the workbook has
them): names/FN.xln` (M3e): no refusal of their own, also listed under a refusal.
`--fresh` and `--live` are still accepted, with a note: every pull is fresh. Excel
writing a number out in full when it saves (`1E-14` → `0.00000000000001`) is no edit.
Provenance tags (`[xln FN 1.2 #3f9a1c]`) are stripped from comments, but a tag's library
base (`lib#353921`) comes back as `@from(lib #353921)` above the name; the summary counts
the tagged names whose live definition or comment no longer matches their hash (edited in
Excel since the build).

It warns, but still pulls, when Excel's owner file `~$<workbook>` exists (unsaved
changes are not in the file). Files whose text does not change are not rewritten.

`--json` prints `{ ok, out, files, unchanged, notices, report, discarded, rewritten, provenance }`
for agents and scripts; a refused pull prints `{ ok: false, refused: true, out, unbuilt,
rewritten, notices }` (each edit: `key`, `file`, `line`, `what`). Exit codes:
0 done, 1 refused (source edits not built) or the workbook is missing or unreadable, 2 bad usage.

**Lockfile format 4** (`xln.lock/4`, 2026-10-05) says that the sheet files write the `#`
(above), so a missing `#` means the cell alone. A project pulled before (format 3 or
older) wrote `Name @C6` whatever the name covered: its build reads a missing `#` on a
name that is on the spill in the lockfile and in the workbook as `#` (so nothing turns
into a name change), and keeps writing format 3 while the source relies on that; the
next pull writes the `#` in (`'#' added, as the name is on the spill`) and format 4.

**Lockfile format 3** (`xln.lock/3`, the hashes formats 3 and 4 share): each hash is `h:` and 16 hex digits (64 bits of
SHA-256) over the definition's tokens without layout, every number literal taken by its
value (`1E-14`, `1e-14`, `0.00000000000001` hash alike; strings, references and `%` do
not change). One line per name and per cell:

```json
    "Rate": { "definition": "h:e11e195b21cd2df9", "comment": null, "hidden": false },
    "S1!C8": { "sheet": "S1", "range": "C8", "formula": "h:01e4ff9fe6d6d39b" },
```

Lockfiles of formats 1 and 2 (`sha256:` and 64 digits, numbers as written), in the
project or embedded in a workbook by an earlier xln, still work: each hash is compared
in the form it was made, and upgraded where the source still matches it. A build always
writes format-3 hashes; an old entry it cannot upgrade (a statement the source no longer has)
keeps its `sha256:` hash. One case stays a conflict: against an old lockfile, a formula
both edited in the source and respelled by Excel's save.

## `xln formulas`

Prints the cell formulas of every sheet (or of `--sheet`, any case) in order of
appearance: row by row, left to right, one entry per formula as Excel shows it, a spill
once at its anchor with its saved size, shared-formula children with their own text. The
workbook is only read. A defined name whose definition is exactly the cell, its spill
(`BS!$C$6#`) or its saved extent leads the line, like the left-hand side of an equation.
The layout is the core's `renderFormulaView` (see `@xln/core`):

```
// Sheet BS: 96 formulas in order of appearance (row by row, left to right). Read-only view of lbo-ep03r.xlsx.
// names · cell · kind · formula · saved value. Names: those defined as the cell, its spill (C6#) or its saved
// extent. C6# (6×1): dynamic array and the size it spilled to when saved; {r×c} legacy array;
// shared ×n / shared ← B2: shared formula and its master; table: data table.

                                  C5#    (1×5)  = Model!Years                                                     → 2022 …
AccountsReceivable_base           C6#    (1×5)  = IN.ASMPT_base("AR days") * IS!Sales_base / 365                  → 12328.76712 …
```

`--json` prints `{ ok, workbook, sheets: [{ sheet, lines }] }`, where each line carries
the cell, kind, extent, stored and displayed formula, saved value, and the names it reads
(`id`, `sheet`, resolved `key`, span in the displayed formula), and `lhs`: the names on
the cell (`key`, `name`, `scope`, `display`, `hidden`, `target`: `cell`, `spill` or `extent`). Exit codes as for `pull`;
an unknown `--sheet` exits 1 and lists the sheets.

`--order calculation` lists each sheet's formulas in calculation order, from the core's
dependency graph: every formula after what it reads (on any sheet, through names too),
otherwise in order of appearance, so a sheet laid out top-down keeps its order. A first
column gives the level (the longest chain of formulas from the inputs; 1 reads inputs
only), `↻n` marks the members of circular reference n, which stand together under a
comment. `--workbook` lists every formula of the workbook in one calculation order, each
address with its sheet (`IS!C16#`). In JSON, calculation-order lines also carry `level`,
`cycle` and `dependsOn` (what the formula reads directly, defined names first).

```
2                                     C5#    (1×5)  = Model!Years                                   → 2022 …
5   AccountsReceivable_base           C6#    (1×5)  = IN.ASMPT_base("AR days") * IS!Sales_base / 365  → 12328.76712 …
```

## `xln graph`

Summarises the dependency graph (`buildGraph` in `@xln/core`): nodes (formula blocks,
input ranges, defined names), edges, the longest chain, circular references with their
members, LAMBDA recursion (allowed, listed apart), references the file alone cannot
resolve (`dynamic`: INDIRECT, OFFSET with computed arguments; `external`: other
workbooks; a relative reference in a name; `broken`: `#REF!`, unknown names or Tables), C9
fixed references into a spill (`C10:G10` where `C10#` follows the spill; `INDEX(C10#, 3)`
for one cell of it; noted when the reference covers only part of the spill or reaches
past it), C10 unused names (and names used only by them) and C12 name cycles. Every
member of a circular reference is listed.

```
xln graph lbo-ep03r.xlsx
  nodes: 735 formulas · 100 input ranges · 539 names; 3071 edges; longest chain 36 (built in 69 ms)
  circular references: none
  not followed: 0 dynamic, 0 external, 0 broken
  C9 fixed references into a spill: none
  C10 unused names: 1: Check!SelfTest
  C12 name cycles: none
```

`--json` prints `{ ok, workbook, nodes, edges, cycles, recursions, flagged, spillRefs,
unusedNames, usedOnlyByUnusedNames, nameCycles, maxLevel, buildMs }`; each of `spillRefs`
is `{ at, ref, use, fit }` with `fit` one of `exact`, `part`, `beyond`.

## `xln check`

Audits the workbook (the core's `audit`, checks C1–C13 of the brief; the rules and their
severities are listed in `@xln/core`'s README). The workbook is only read. The report:
the counts per check, then the findings grouped by check (severity, place, message, `→`
hint), then the name census (C8: by kind and scope, coordinate tags and families, the
tiers of names standing in for dimensions, with how to read them) and the spill census
(C9: per sheet the spills named as `x#`, only by a fixed range, or unnamed; the spills
without an `x#` name). At most 100 findings per check are printed; `--json` has all.

**The source too.** When the workbook has its project (`<workbook>.xln` beside it, or the
project folder given in place of the workbook), `xln check` also runs the checks the
editor runs as you type on the project's names files (the core's `sourceFindings`, the
checker the build refuses on) and lists them after the report, one line each at
`file:line:column`, with the quick fix the editor offers; hints included unless a
`--severity` above `hint` is given. So the command line and the editor report the same findings (until
2026-10-07 `xln check` audited the workbook only, and a finding of the source checker,
such as *workbook-on-cell*, showed in the editor alone):

```
source /…/is-model.xln: 0 errors, 0 warnings, 0 info, 1 hint (the editor's checks as you type)
  hint    names/sheets/S1.xln:9:1 workbook-on-cell: workbook name on a cell of S1, read only on S1: remove @workbook to make it local to S1 (the build then moves it) [quick fix: Remove @workbook: make Spl local to S1]
```

`--json` adds `source: { project, counts, findings }`; an error in the source makes the
exit code 1.

```
xln check lbo-ep03r.xlsx: 1 warning
  C1 Syntax                            none
  …
  C10 Unused names                     1 warning
  C13 Constants in LAMBDA bodies       none

== C10 Unused names: 1 warning
  warning name Check!SelfTest: nothing reads Check!SelfTest: no formula, name, conditional format, data validation, Table column or chart

== C8 Name census
  539 names: 22 workbook-scoped, 517 sheet-scoped (Model 5, Assumpt 12, …)
  names standing in for dimensions: 526 of 539 (97.6%): T1 9 · T2 178 · T3 336 · T4 3
```

With the harness declared in `lbo-ep03r.xln/xln.config.json` (below) the same workbook
reports `none` and the published census, 433 of 539.

- `--only C2,C9`: run only these checks (the census sections are always there).
- `--severity warning`: leave out findings below that severity (`error`, `warning`, `info`, `hint`; `hint` lists everything).
- `--census-exclude 'Check!*,Model!Fix*'`: the check harness for this run, replacing the
  settings' (below): names C10 does not report and the C8 tiers leave out (a self-test,
  solver settings); globs on `Sheet!Name` or `Module.Name`, any case. With
  `Check!*,Model!Fix*,*!FixedPoint_*` lbo-ep03r gives the published 433 of 539.
- `--config <file>`: the settings file to use instead of `<workbook>.xln/xln.config.json`.
- `--json`: `{ ok, format: "xln-audit/1", workbook, counts, byCheck, findings, census,
  spills, checks }`; each finding `{ check, rule, severity, where, message, hint?, data? }`
  with `where` `{ kind, sheet?, name?, key?, ref?, range?, span?, text? }`. For agents
  that run `xln check` as a critic after editing a workbook.

**Project settings.** `xln check lbo.xlsx` reads `lbo.xln/xln.config.json` when it exists
(the report then starts with `settings: <path>`; what in the file cannot be used is said
on stderr and left out):

```json
{ "audit": { "harness": ["Check!*", "CHK.*", "Model!Fix*", "*!FixedPoint_*"],
             "rules": { "C10.unused": "warning", "C2.unknown-function": "off" },
             "constants": { "allow": [4, 1.5], "sentinelAbove": 1e90 } } }
```

`harness` is `--census-exclude` made persistent: the way to silence an intentional unused
name (an unused constant or LAMBDA stays a warning; an unused name on cells is info). `rules` sets a severity (`off`, `info`, `warning`, `error`)
per rule or check; `constants.allow` adds numbers to C13's defaults, and numbers from
`sentinelAbove` (default 1E+90) up are sentinels, not constants. Flags win over the file.
`xln pull` writes `{ "audit": { "harness": [], "rules": {}, "constants": { "allow": [] } } }`
when the project has no settings yet. JSON output adds `config: { path, problems }`.

The known keys are `audit` (`harness`, `rules`, `constants` with `allow` and
`sentinelAbove`), `build` (`embed`) and `library` (the core's `CONFIG_KEYS`). Any other
key is flagged, never silently dropped: a known setting in the wrong section says where it
goes (`audit.library: \`library\` is a top-level setting, not an audit setting: move it
out of "audit"`), a near miss names the setting meant (`libary: … did you mean
"library"?`), anything else lists the keys of its section. `xln check` says them on
stderr, `xln build` and `xln lib status` as `note: xln.config.json: …` lines.

Exit codes: **0** no errors (warnings and info allowed), **1** errors found, **2** bad
usage or a workbook that is missing or cannot be read (`--json` then prints `{ ok: false, error }`).

## `xln build`

Writes the project's names back into the workbook (M3a: D1–D4, D9, E1–E3, E5). It reads
every `names/**/*.xln`, the lockfile `xln.lock.json` (what the last pull or build saw)
and the workbook as it is now, and compares each name three ways:

| Source vs lock | Workbook vs lock | Result |
|---|---|---|
| changed | unchanged | written (create, update definition, comment, hidden flag, delete, scope change, rename) |
| unchanged | changed | Excel's version kept, reported as a note (pull again to update the source) |
| changed | changed, differently | **conflict**: reported with both versions, nothing written |

- **Rename** a name with `xln rename` (below) or by writing `@renamed(OldName)` before
  its new definition and the new name in every formula of the source that reads it.
  A rename in the same scope also rewrites the name's token in the workbook's cell
  formulas (a shared group's master), conditional formats, validations and other names
  (M5, stretch G): a token substitution, the formulas' form and metadata kept; a cell
  statement whose only edit is the renamed token is not written as a new formula. It is
  refused, with the places, when a chart, a Table column, a pivot table's source, a
  hyperlink, a form control or a link of the workbook to itself reads the name (rename it
  in Excel's Name Manager instead, then pull), and when the new name would make a formula
  read something else (`rename-capture`: a LET/LAMBDA variable, a sheet's local name).
  A rename that also changes the scope (`@renamed(Sheet!Old)`) rewrites nothing: its
  readers are the source's to change, as below. Once the workbook is written, the build
  removes the `@renamed` lines it applied from the source (not with `--out` or
  `--dry-run`): the one source edit a build makes. Adding or removing
  `@workbook` above a name in a sheet file, or `@sheet(Sheet)` above a module name (or
  moving a name to another sheet's file) changes its scope. A rename, deletion or scope
  change that would break
  a reference in cells, conditional formats, validations, Table columns, charts or other
  names is refused, with the list of places (D4, v1: cells are not rewritten). The places
  counted are those of the workbook **as this build leaves it**: a cell whose formula the
  same build sets or clears counts by its new formula, a name it updates or deletes by its
  new definition (or not at all), and the names it creates are there for the new
  formulas. So one build can move a cell from `ANA.GROW(…)` to a new `FN.GROW(…)`,
  create `FN.GROW` and delete `ANA.GROW` (feedback 2026-10-07). When a reader really
  remains, the message names it with its statement and says what to do:
  `deleting Fn refused: 1 place in the workbook refers to it by name and would break:
  cell S1!C10 (names/sheets/S1.xln:19). Change that formula in the source, or keep the name`.
- **A refused build** prints its reasons first (errors with `file:line`, then conflicts),
  then the plan it did not write, then warnings and notes.
- **Cell statements** (E6) are compared the same way, per statement, on the hash of
  the top-left formula (tokens without layout): changed only in the source →
  `set-cell-formula` (or `clear-cell-formula` for `= ;`); only in Excel → kept and
  reported; in both, differently → a conflict. A named cell follows its name: if Excel
  moved it (rows inserted), the build writes where the name now points and notes the
  move. An unnamed cell is its address: moved in Excel, it shows as one statement
  emptied and one cell added. An address that differs from the lockfile's refuses the
  build (the problem carries the fix). A source text equal to the workbook's formula
  decompiled is no edit, even if compiling it would add a missing prefix: cells are not
  repaired. Formula cells added in Excel, and unnamed statements missing from the
  source, are reported and left alone. A named statement removed from the source
  refuses the build: clear the cell with `Name @C6 = ;`, or delete the name in Excel's
  Name Manager and pull. A format-1 lockfile (no cells) is still read:
  cells are then checked against the workbook only.
- **The `#` of a named statement** is the one editable part of its address: adding or
  removing it is a `set-name` (definition `ANCHORARRAY` of the cell, or the cell), with
  the three-way check of any name (changed in Excel only: kept and reported; in both:
  the same way is no change, otherwise a conflict). No build moves a name to `#` on its
  own: a named statement without `#` on a formula whose saved spill is larger than one
  cell gets a warning, `Sales @C3: the formula spills over C3:G3, but Sales covers only
  C3 (write @C3# to name the spill)`, whose fix inserts the `#` (a 1×1 result gets none).
- **Slots**: filling `Name @C6 = ;` writes the cell and leaves the name on the cell;
  `Name @C6# = formula;` also puts the name on the spill (measured: `C6#` works on a
  cell that does not spill). A `#` on a slot left empty (`Name @C6# = ;`) is a warning
  and is not written until the cell gets a formula (a name on `C6#` of an empty cell is
  not measured).
- **The source's errors** are the checker's (M3d), the same the extension shows as you
  type: statement syntax and annotations, addresses (bare in a sheet file, the address
  the last pull saw), names defined twice or that Excel would refuse, the scope rules
  above, and every formula compiled (an unknown function, `_xlfn.X` not in the catalogue,
  a sheet the workbook lacks, another sheet's local name read bare, argument counts),
  each with its line and a message saying what to write instead. A formula as the
  workbook already has it (unchanged since the last pull) keeps its findings as warnings:
  the build leaves it. A file with no place in `names/` is refused too (M3e: anything not
  `.xln`, hidden files aside; a sheet file for a sheet the workbook lacks, *there is no
  sheet Foo in the workbook: sheets are created in Excel, then pulled*; a folder other
  than `sheets/`; a module file not named after a prefix, e.g. `my module.xln`).
  Besides these the build refuses only what needs the workbook:
  conflicts, a change that would break references (D4), a removed named statement, a
  comment over 255 characters once the provenance tag is added.
- **`--rescope-slots` is gone** (M3d): no step changes a scope on its own. Passing it is
  a usage error that says the way now: remove `@workbook` above the name, then build.
- New and renamed names must be legal Excel names and must not collide with a built-in
  function (`Fact` → `FACT`, probe T12).
- Only `<definedNames>` changes in `xl/workbook.xml`, plus `fullCalcOnLoad="1"` on
  `<calcPr>`; every other zip entry is copied byte for byte. An untouched name keeps its
  exact markup. Excel's own `_xl*` names are never touched. Cell changes in the change
  set (M3b) also patch the sheets they name, as described under `xln apply`.
- Before writing, the built bytes are read back and decompiled (E3): every source name
  must equal its source modulo whitespace, everything else must be as it was. After
  writing, the file is read again; on a mismatch the original is restored.
- Refuses while Excel has the file open (`~$<file>` beside it, E1). Keeps the previous
  file as `<workbook>.backup.xlsx` (E5). Then rewrites `xln.lock.json` and
  `workbook.manifest.json` (not the `.xln` files: their layout is the author's).
- **Embedded source** (D5, opt-in since 2026-10-05): `--embed`, or `"build": {"embed":
  true}` in `xln.config.json`, also writes the project into the workbook, as a custom
  XML part: `names/**/*.xln`, the lockfile it leaves, `xln.config.json`; for a workbook
  handed on without its project folder (an archive copy: pull does not read it). A build whose
  names and cells are up to date still writes once if the part is missing or stale.
  Excel keeps the part when it saves (measured, `probes/README.md` § F5). `--no-embed`
  overrides the config. A build without embedding leaves an existing part as it is
  (neither refreshed nor removed).
- **Provenance** (D6): names declared in a module file (`names/FN.xln`) carry
  `[xln FN 1.2 #3f9a1c]` at the end of their Name Manager comment: the module, the
  version from a `// @version 1.2` line in the module file's header (optional), and a
  hash of the definition and comment; then, when the entry says `@from(lib #353921)`,
  its library base: `[xln FN 1.2 #636cf1 lib#353921]`. Pull strips the tag and writes
  `@from` back; a tag never counts as an edit (a changed `@from` is a provenance change). A
  comment that would pass Excel's 255 characters with the tag is written without it
  (warning). `--no-tags` turns it off. A comment over 255 characters refuses the build:
  Excel will not open such a file (measured, F5).
- `--out <file>` writes the built workbook elsewhere and leaves the original and the
  lockfile alone; `--dry-run` prints the change set only; `--force` writes even with no
  change (only `fullCalcOnLoad` then); `--project <dir>` names the project folder.

`--json` prints `{ ok, exit, status, message, written, backup, projectFiles, changeSet,
conflicts, problems, excelChanges, unchanged, unchangedCells, readBack }`. The change set is the core's
`ChangeSet` (`xln.changes/1`): `rename-name`, `rescope-name`, `delete-name`, `set-name`
(with `stored` and `display` forms), in the order a backend applies them.

Exit codes: **0** built, up to date, or dry run; **1** refused (source errors, conflicts,
names still in use); **2** bad usage or unreadable workbook or project; **3** Excel has
the file open (the message points to `--reopen`); **4** read-back failed (nothing written, or the original restored);
**5** `--reopen`: Excel has unsaved changes in the workbook; **6** `--reopen`: built, but
Excel did not open it again (a repair prompt or another dialog: look in Excel).

### `--reopen` (E7, desktop Excel)

`xln build <workbook> --reopen` closes the workbook in Excel **without saving**, builds,
and opens it again, so the author need not close and reopen it by hand. On the Mac it
drives Excel with `osascript` (AppleScript); on Windows with a PowerShell script over
Excel's COM interface (**written without a Windows machine: untested**). The build reads
the file on disk, so changes Excel has not saved would be lost by the close: when Excel
reports unsaved changes the build is refused (exit 5); save them in Excel and pull, or
pass `--discard` to lose them. If Excel does not have the workbook open, it is simply
opened after the build. An open is given 90 s and then checked again after a pause (M3d):
the workbook must be open under its full path, Excel must answer within 15 s (a repair
prompt or a file-access request still up makes it time out), and no repair log may have
appeared. Anything else is exit 6 with the reason, never a success; xln does not retry.
After the check Excel comes to the front (M3e: AppleScript `activate` on the Mac; on
Windows the workbook is activated and Excel's window asked to the foreground, which
Windows may answer by flashing the taskbar button). On the Mac,
Excel's sandbox may ask once for access to a folder outside its own containers.

## `xln apply`

```
xln apply <workbook.xlsx> <changes.json> [--out <file>] [--dry-run] [--reopen [--discard]] [--json]
```

Applies a change set without a project: the `changeSet` that `xln build --json` prints,
or a list of changes. For hand-built change sets (the backend's Excel checks), agents,
and cell changes before the source language has them. Cell changes:

```json
[
  { "op": "set-cell-formula", "sheet": "BS", "range": "C5", "stored": "_xlfn.TAKE(Model!Years,,3)", "display": "TAKE(Model!Years,,3)" },
  { "op": "set-cell-formula", "sheet": "Model", "range": "K20:K21", "stored": "ROW()*3", "display": "ROW()*3" },
  { "op": "clear-cell-formula", "sheet": "Model", "range": "L24" }
]
```

A range is written cell by cell with the top-left formula moved like a fill. Every
formula is written in dynamic-array form (probe F8); a spill anchor's old area is emptied
(styles kept), a shared group whose master changes is un-shared, an empty cell gets its
`<c>` (and `<row>`), `xl/metadata.xml` gains the dynamic-array record if needed, and
`xl/calcChain.xml` is dropped (Excel rebuilds it; a stale one makes it repair the file).
The summary lists what was emptied, un-shared and inserted. Same safety as `build`
(read-back before and after writing, the `~$` guard, `<workbook>.backup.xlsx`); the
project's lockfile is not touched. Exit codes: **0** applied or dry run; **2** bad usage
or input; **3** Excel has the file open; **4** the changes do not apply (an unknown sheet,
part of a legacy array or a data table) or did not read back; **5**, **6** as for
`build --reopen`.

## `xln verify`

```
xln verify <workbook.xlsx> [--before <file>] [--tolerance <x>] [--json]
```

After opening the built workbook in Excel (which recalculates it on load) and saving it,
compares its cached values cell by cell with the copy before the build, by default the
backup `<workbook>.backup.xlsx` (E4). Every cell with a value counts, spill cells
included. A build of names that should not change numbers must show none. Exit codes:
**0** no cell changed, **1** some changed (each listed), **2** unreadable.

A side whose values Excel did not calculate is named first, as a warning (`warnings`,
with `beforeValues` and `afterValues`, in `--json`): `original.xlsx has no values Excel
calculated (it was never saved by Excel): open and save it in Excel first, or the
comparison is empty` for a file written by Python or xln and never opened (no value, or
`fullCalcOnLoad="1"` with only placeholder 0 or "" values, as XlsxWriter writes), and
`… was written after Excel last saved it: 26 of 49 formula cells have no value Excel
calculated …` for one a tool changed since (`fullCalcOnLoad="1"`, which Excel never
saves, and formula cells without a value).

## `xln rename` (M5)

```
xln rename <project | workbook.xlsx> <Old> <New> [--dry-run] [--json]
```

Renames a defined name in the project's source, explicitly (pull and build never rename
on their own). `Old` is `Name`, or `Sheet!Name` for a sheet's local name (a bare name is
the workbook's, else the only local one). The name's statement gets the new name and
`@renamed(Old)` above it, the record the next build reads (kept when an earlier rename
not built yet already says where the name is in the workbook; removed when renamed
back; none for a name not built yet); every formula of the project's `.xln` files that
reads it gets the new token (strings, LET/LAMBDA variables and other scopes' names of the
same spelling untouched). With the workbook beside the project the next build is planned
on the renamed source first: a rename it would refuse is refused here, nothing written,
with its reasons. Then `xln build` renames the name, rewrites it in the cells and, once
the workbook is written, removes the `@renamed` line (`removed @renamed(Rate) from
names/_unmanaged.xln: the rename is built`; `renamedRemoved` in `--json`). For a name
on cells, both commands then give the **label notice** (`labelNotice` in `xln rename
--json`, `labelNotices` in `xln build --json`): the text cells in the name's rows and
columns on its sheet whose whole text still gives the old name, typed text only, and the
Find & Replace that fixes them in Excel, one pair per text in the label's style; cells
that only resemble the old name are listed to check by eye, never in the replace. xln
never writes cell values. Exit
codes: **0** renamed (or dry run), **1** refused (unknown, invalid or taken name, a
capture, a reader the build cannot reach), **2** bad usage.

```
$ xln rename lbo.xln Rate Pace
xln rename Rate → Pace in lbo.xln
  names/_unmanaged.xln: the name, @renamed(Rate), 2 references
  names/sheets/S1.xln: 7 references
  the next build will rename Rate → Pace, rewriting it in 8 cell formulas, 1 conditional format, 1 validation, 2 names
  wrote names/_unmanaged.xln, names/sheets/S1.xln
  then: xln build lbo.xlsx
```

On the author's income statement (column A the readable text, B the name's text that
Create from Selection read, C:G the cells):

```
$ xln build is-model.xlsx
xln build is-model.xlsx: built is-model.xlsx: 1 change
  rename IS!Gross_income → Gross_ind_income, rewriting it in 1 cell formula
  …
  removed @renamed(Gross_income) from names/sheets/IS.xln: the rename is built
  Gross_income → Gross_ind_income: 2 labels still read the old name: IS!A5, IS!B5
  In Excel, on sheet IS: Find & Replace (Ctrl+H on Windows; Ctrl+H or ⌘⇧H on Mac; or Home → Find & Select → Replace)
    1. Find what:     Gross income
       Replace with:  Gross ind income
    2. Find what:     Gross_income
       Replace with:  Gross_ind_income
    Within: Sheet · Look in: Formulas · Match entire cell contents ✓ · then Replace All (for each pair)
```

## `xln lib` (M4)

```
xln lib status <workbook.xlsx | project> [--lib <dir>] [--json] [--no-diff]
xln lib publish <project> <Name> [--lib <dir>] [--dry-run] [--json]
xln lib take <project> <Name> [--lib <dir>] [--dry-run] [--discard] [--json]
xln lib base <project> <Name | --all> [--lib <dir>] [--dry-run] [--json]
```

The library is a folder of `.lambda` files (the author's `_shared/lib`): `--lib <dir>`,
else `"library"` in the project's `xln.config.json`, relative to the project folder or
absolute, `~` for the home folder (for a workbook, the project is `<workbook>.xln`).

**The library base** (decided 2026-10-07): a copy taken from the library records the
library version it came from, visibly, as `@from(lib #353921)` on the line above the
name. Only explicit library actions write it: *Insert* and *Take* (the library's
version), *Publish* (the version just published). The build carries it in the tag
(`lib#353921`), pull writes it back; delete it and the copy has no base. The version is
the first 6 hex digits of a hash of the stored definition, layout and number spelling
aside: `lib status` prints it as `library #…` and the copy's as `base #…`.

`lib status` reads (never writes) and reports each library function, three-way on the
copy, its base and the library: **identical** (the copy equals the library),
**outdated** (the copy is still its base; the library moved: take it), **modified** (edited
here; the library is still the base: publish it, or take to undo), **both changed** (each
side moved: the diffs from the base, base → copy and base → library, from the project's
kept `library-bases/<hash>.json`, else when the workbook or its `<name>.backup.xlsx`
still has the base's text; else the copy against the library),
**differs** (no base recorded: which side moved cannot be told) or **missing**, plus **local only**
LAMBDAs of a library module (`FN.*`) the library lacks (other modules' LAMBDAs, such as
the `IN.*` input readers, are counted on one line: not library candidates); diffs `-` the copy, `+` the library. On a project the
copies are the source (unsaved edits on disk) with their `@from`, and the tags shown come
from the workbook beside it when it is there. Since the source is not the workbook, each function the
workbook does not have as the source does is marked **not built yet** (`in the source,
not built yet`, `edited in the source, not built yet`, or `deleted in the source, still in
the workbook`), and the summary line counts them (`· not built yet 3`, then *The source
differs from is-model.xlsx: 3 functions are not built yet*), so a status never looks done
while the workbook is not (feedback 2026-10-07). Without the workbook beside the project
it says that what is built cannot be told. Run on a workbook it reports the workbook as
it is. Exit 0, or 2 (bad usage, no library). For the three LBO
workbooks against `_shared/lib` (2026-10-06):

| Workbook | identical | missing | local only (other module LAMBDAs) | outdated / modified / differs |
|---|---|---|---|---|
| `lbo-ep02.xlsx` | 13 | 4 (AVGNEG, AVGPART, AVGPOS, FIXPOINT) | 0 (`IN` 8) | 0 |
| `lbo-ep03.xlsx` | 12 | 5 (AVGPART, BACKDISC, FIXPOINT, NPV0, TAILROW) | 0 (`IN` 6) | 0 |
| `lbo-ep03r.xlsx` | 14 | 3 (BACKDISC, NPV0, TAILROW) | 0 (`IN` 8) | 0 |

(The Python build writes no provenance tags, so nothing in them has a base: a copy that
changed would show as differs. A tag written before 2026-10-07 has no `lib#` either:
the author's FN.SPREAD, renamed locally and built, read as outdated then and reads as
differs now; after *Take* and a build it reads identical, and a later local edit as
modified.)

`lib publish` writes one project LAMBDA to the library: a new `<Name>.lambda` with a
generated header, or the existing file with its header fields and rationale kept and the
definition (and summary, params, `@param` descriptions, when they changed) replaced. It
prints the file's diff; `--dry-run` prints it and writes nothing. A definition equal to
the library's modulo layout writes nothing. The project's entry then records the
published version, `@from(lib #…)` in its names file (build to carry it into the
workbook).

`lib take` gives a project's entry the library's definition and doc comment and records
that version, `@from(lib #…)`; build to write it. It refuses (exit 1, with the diff) a
copy that has an edit of its own, **modified** or **both changed**, or may have one,
**differs**, unless `--discard`, which loses that edit. `--dry-run` prints the change.

`lib base` (*Record library base*, author's idea 2026-10-07) writes `@from(lib #…)` on an
entry **identical** to the library that records no base (inserted or published before
the base existed: it would read *differs* after the first change on either side);
`--all` does it for every such entry. It prints each one written; a function that is not
identical, or already has a base, is skipped with the reason (exit 1 when the one named
is). `--dry-run` prints what it would record. Nothing else records one: `lib status` only
notes *no base recorded: Record library base* on these.

**The base's text.** `publish`, `take` and `base` also keep the base definition in the
project, `library-bases/<hash>.json` (stored and display text, library file, hash; the
JSON outputs list it as `kept`), which `lib status` reads first for the three-way diff
(on a workbook, from the project beside it). A pull writes and removes nothing there;
nothing prunes it.

## Dependencies

Runtime: `@xln/core` only. Development: `@types/node`.
