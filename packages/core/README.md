# @xln/core

The pure library behind `xln`: TypeScript, ESM, no Node built-ins, so the same code runs
in the CLI and in VS Code for the web. Bytes are `Uint8Array`, text is `string`.

## Layers

| Folder | Layer | Entry point |
|---|---|---|
| `src/file/` | L0: read the `.xlsx` package | `readWorkbook(bytes): WorkbookSnapshot`, `lockFileName`, `isLocked` |
| `src/lang/` | L1: formulas | `parse`, `decompile`, `compile`, `prettyPrint`, `shiftFormula(text, dRow, dCol)`, `formulaCursor(text, offset)` (M3c: the word being typed, its sheet, open calls and their argument, LET/LAMBDA variables in scope; read from tokens, since the text before a cursor is rarely a whole formula) |
| `src/project/` | L2: the workbook as a project | `pullProject(bytes, fileName) → { files, report, names, snapshot }`, `parseModule(text)`, `parseDocComment(doc) → { summary, params }` (M3c: `@param x …` tags of a LAMBDA's doc comment, which stays the whole Name Manager comment) |
| `src/view/` | B6: views of a sheet's cells | `sheetFormulaView(snapshot, sheet, names?) → FormulaViewLine[]`, `sheetCalcView`, `workbookFormulaView`, `renderFormulaView(lines, opts) → { text, entries }` |
| `src/graph/` | the cell dependency graph (B6 b; C9, C10, C12) | `buildGraph(snapshot, names?) → DependencyGraph` |
| `src/audit/` | L3: the audit, checks C1–C15 | `audit(snapshot, opts?) → AuditReport`, `renderAuditReport(report) → { text, links }`, `RULES` |
| `src/build/` | M3a: write names; M3b-2: write cell formulas | `buildWorkbook({ workbook, fileName, files }) → BuildResult`, `planBuild`, `applyChangeSet(bytes, changes)`, `applyChangeSetWithReport`, `readBack`, `verifyValues(before, after)`, `excelControl` |
| `src/check/` | M3d: the source model and the one checker of the editor and the build | `SourceModel`, `checkFile(model, path, ctx)`, `checkProject`, `invalidName` |
| `src/library/` | M4: the library of `.lambda` files | `parseLambdaFile`, `readLibrary`, `libraryStatus(lib, workbookCopies(wb) \| projectCopies(files, wb?))`, `renderLibStatus`, `libraryInsertion`, `libraryReplacement`, `publishLambda`, `libraryBaseRecordings`, the kept bases (`readBases`, `baseFiles`) |

### `check/` (M3d)

One checker for the editor and the build. `SourceModel` (`model.ts`) holds a project's
`.xln` files parsed by path (`parseSourceFile`), each formula parsed and every identifier
resolved as Excel resolves it (`occurrences`, `NameResolver`); the extension's project
model extends it. `checkFile(model, path, ctx)` / `checkProject(model, ctx)` (`check.ts`)
return problems with offsets, a code, a message that says what to write instead, and
quick fixes: syntax and annotations, a sheet file in the old block form (info, fix:
*Convert to per-name @workbook*), addresses (bare in a sheet file, the lockfile's,
`#` only on a named single cell), duplicates, names Excel would refuse (`invalidName`),
scope (a hint, advice only, on a workbook name on a sheet's cell that only that sheet
reads, with a fix that removes `@workbook`; none when another sheet or a name of another
scope reads it, since it then needs workbook scope; an info noting a scope change as the re-scope the build makes; no
setting since 2026-10-06: Excel and the source both set scope), and every formula compiled with
the real compiler (unknown functions, `_xlfn.X` not in the catalogue, unknown sheets,
another sheet's local name read bare, argument counts). Findings in a formula the
lockfile already has are warnings (the build leaves it). A LAMBDA's doc comment against
its parameters (2026-10-07): an `@param` naming no parameter is a warning (`doc-param`,
fixes *Rename @param x to y*, the parameter at its position when that one is not
documented already, and *Remove @param x*); parameters left out while others are
documented, a hint (`doc-param-missing`); `@param [p]` and `@param p` both match `[p]`
(`docParamSpans` in `project/doc.ts` gives the tags' offsets). `ctx`: the workbook's sheets
and Tables, the lockfile. `planBuild`
runs it (`sourceErrors`) and refuses on its errors.

### `build/`

A build is a **change set** (`changes.ts`, D9): plain JSON (`xln.changes/1`) with
`rename-name`, `rescope-name`, `delete-name` and `set-name` (the definition in `stored`
and `display` form, comment, hidden), sheets named by name. A backend applies it; the
file backend is `applyChangeSet(bytes, changes) → bytes`. A live AppleScript/COM backend
can take the same list. M3b adds `set-cell-formula` and `clear-cell-formula` (sheet by name, a cell or a range, the top-left formula in `stored` and `display` form).

- `planBuild` (`plan.ts`) reads the source (`source.ts`: every `names/**/*.xln`;
  `@renamed(Old)` marks a rename), compiles each name against the target set of names
  (prefixes, `_xlpm.`, home sheet qualified, `#` → `ANCHORARRAY`, CR LF), and compares
  source, lockfile and workbook per name (E2): changed in source only → a change;
  in Excel only → kept and reported; in both → a `Conflict` with both versions. The
  source's errors are the checker's (`check/`, below), the same the editor shows; the plan
  adds only what needs the workbook. Renames, deletions and scope
  changes that would leave a cell, format, validation, Table column, chart or kept name
  pointing at nothing (or at another name) are refused with the places (D4), counted on
  the workbook as the build leaves it: a cell the same build sets or clears counts by its
  new formula, the names it creates are there for it (feedback 2026-10-07). The refusal
  sits at the first reader's statement (`file:line`) and says to change the formula in
  the source or keep the name. `buildReportLines(result)` is the report body the CLI and
  the editor print (a refusal's reasons first), `refusalReasons(result)` one line each.
  `sourceFindings(files, ctx)` is the checker's every finding at `file:line:column`
  (what `xln check` lists for a project); `sourceErrors` its errors, which the build
  refuses on.
- **Rename across cells** (M5, stretch G). `project/rename.ts`: `renameInFormula(text,
  home, RenameContext)` rewrites a renamed name's token in one formula (stored or display
  form), resolving names as Excel does (a sheet's local name shadows the workbook's,
  `Sheet!Name`, `[0]!Name`, `Name#`/`ANCHORARRAY`; LET/LAMBDA variables, strings,
  structured references and other workbooks' names untouched; case ignored), and reports
  a `captured` reader instead of writing it. `build/renameSites.ts` sorts the workbook's
  readers of a rename into what the build rewrites (cells, formats, validations, names)
  and what refuses it (charts, Table columns, pivot sources, hyperlinks, form controls, a
  self-link, formulas that do not parse). The plan marks a same-scope rename with
  `references` (counts), and leaves cell statements and names whose only edit is the
  renamed token to it (`BuildPlan.renamedOnly`). The file backend's
  `applyReferenceRenames` runs first, splicing at the XML tokenizer's offsets of each
  formula text (`formulaText.ts`), dropping `calcChain.xml` after a cell formula changed;
  read-back checks that step on its own (`readbackRename.ts`). `build/renameSource.ts`:
  `renameInSource`/`renameInProject`, the source edits of `xln rename` and the editor's
  Rename Symbol (the statement, `@renamed(Old)` from the lockfile, every reader), with the
  same capture check on the source; `consumedRenamedEdits`, the removal of the
  `@renamed(…)` a build applied, and, given the lockfile the build read, of those an
  earlier build applied (`renameBuilt` in `check.ts`, the test the plan ignores them by
  and the checker's `renamed-built` hint uses) (LANGUAGE §9.6).
- `planCells` (`cells.ts`, E6, M3b-1) does the same per cell statement, on the hash of
  the top-left stored formula: a named cell's identity is its name (moved in Excel: the
  change goes where the name points now, noted as `moved`), an unnamed one's is its
  address; an address that differs from the lockfile's is an `address` error with a
  `fix`. A named cell's definition is not the source's to change: the plan keeps
  Excel's, except that writing a named cell moves its name to `'S'!$C$6#`. Excel-only
  edits are kept and noted, formula cells added in Excel are noted, a source text equal
  to the decompiled workbook formula is no edit (cells are not repaired). A scope changes
  only where the source changes it (M3d: `rescopeSlots`, `rescopeCandidates` and the
  rewritten sheet file are gone; a build never writes the source files, but for the
  `@renamed` notes of the renames it made, or that earlier builds made: `consumedRenamedEdits`, returned as
  `BuildResult.renamedConsumed` and `sourceFiles` for the caller to write with the lockfile).
  `BuildPlan` gains `cellsInSync`, `unchangedCells`;
  `buildWorkbook({…}, { dryRun: true })` stops after planning (status `planned`).
- `apply.ts` / `workbookXml.ts` patch `xl/workbook.xml` with the XML tokenizer, which
  now reports each token's offsets: untouched `<definedName>` elements keep their exact
  markup, changed ones are re-serialised (attributes in Excel's order, `>` as `&gt;`, line
  breaks in comments as `&#10;`), moved and new ones go to Excel's sort position (name,
  then scope sheet name, case ignored, workbook scope last). A missing `<definedNames>`
  is inserted before `<calcPr>` (schema order); `fullCalcOnLoad="1"` is set or added.
- `file/zip.ts` rewrites the archive at record level: every other entry's local record
  is copied byte for byte (fflate's `zipSync` would recompress them). ZIP64 and
  encrypted archives are refused.
- `readBack` (E3): same entries in the same order, byte-identical except the workbook
  part, which is identical outside `<definedNames>`/`<calcPr>`; every source name in
  sync decompiles to its source text modulo whitespace; untouched names unchanged.
- **Cell changes** (M3b-2, `set-cell-formula`, `clear-cell-formula`) follow probe F8's
  recipe (`probes/README.md` § F8). `sheetXml.ts` patches each named worksheet part at
  tokenizer offsets: every formula written in dynamic-array form (`cm` = the XLDAPR
  `cellMetadata` record, `<f t="array" ref="<cell>">`); the cell's `<v>`, `<is>`, `t`
  (and `vm`) removed, `s` and other attributes kept; a range written cell by cell with
  the top-left formula shifted (`shiftFormula`); an array anchor that changes has its
  old area emptied (cells with only an address removed, styled ones kept empty, so the
  spill keeps its formatting); a shared master that changes un-shares its group (the
  other members get the translated text and keep their value), an edited child gets its
  own formula; slots get a `<c>` (and a `<row>`) inserted in order; a clear removes
  formulas only (a typed value in its range stays). Part of a legacy
  (CSE) array or a data table is refused, as Excel refuses it. `packageXml.ts` finds the
  dynamic-array record in `xl/metadata.xml`, adds it to a part that holds only other
  types, or creates the part with its relationship and content type; drops
  `xl/calcChain.xml` with its relationship and content-type override (a stale chain
  makes Excel repair the file). `applyChangeSetWithReport` also says what it did
  (cells set, cleared, inserted, emptied spill cells, un-shared members).
- `file/zip.ts` `rewriteZip` replaces, removes and adds entries; every other entry is
  copied byte for byte.
- `readBack` for cell changes (`readbackCells.ts`), worked out from the original and the
  change list rather than from the patcher's report: each changed cell decompiles to the
  change's display text (moved to the cell) in dynamic-array form with no cached value
  and its style kept; cleared cells are empty; old spill areas are emptied and un-shared
  members hold the translated text with their value; every other cell on every sheet is
  unchanged (and each dynamic array still reads as one); the sheets are identical
  outside `<sheetData>`; calcChain is gone; the relationships and content types differ
  by nothing else than calcChain and a new metadata part.
- `excelControl(platform, exec)` (`excelControl.ts`, E7): the AppleScript and PowerShell
  (COM, **untested**) scripts that close a workbook in desktop Excel without saving,
  tell whether it has unsaved changes, and open it again (an open that does not leave a
  workbook of that name open counts as failed: a repair prompt). Pure: the caller runs
  the program (`exec`).
- `verifyValues` (E4) compares every cached cell value of two copies.
- **Embedded source** (D5, `embedXml.ts`, `embed.ts`; opt-in since 2026-10-05:
  `buildWorkbook(…, { embed: true })`, which the CLI and the extension pass for
  `--embed` or `"build": {"embed": true}` in `xln.config.json`, `embedSetting`; without
  it an existing part is left as it is): a build carries the project
  in a custom XML part, `customXml/itemN.xml`, root `<project xmlns="urn:xln:embedded-source:1"
  format="xln.embed/1">` with one `<file path sha256>` per file: every `names/**/*.xln`,
  the lockfile **as the build leaves it** and `xln.config.json` (not the manifest, which
  pull derives from the workbook). The part is an archive copy for a workbook handed on:
  pull never reads it (decided 2026-10-06, every pull is fresh); `readEmbeddedSource`
  stays for the build (a part that already says the same is left alone), its read-back
  and the tests. Text escapes `&`, `<`, `>`;
  control characters force base64, and a checksum per file catches damage. Files are
  stored with LF line ends and `eol="lf"|"crlf"` (mixed ones as base64): Excel re-saves
  the item with every LF as CR LF (F5), and the reader undoes that before checking. The part
  comes with `itemPropsN.xml` (the fixed xln datastore itemID
  `{6B1D2E7A-4C3F-4E8B-9A15-7F0C3D2B5E91}`, a schemaRef to the namespace), the item's
  relationship to it, a `customXml` relationship from the workbook part and the
  properties' content type, as Excel writes one. A later build finds the part by the
  namespace of its root (Excel renumbers items) and rewrites it in place, or leaves it
  when it already says the project; it is never duplicated. In the change set this is
  `set-embedded-source` (`files`: path → text), applied last on the patched package; a
  live backend that cannot write custom XML may skip it. Without `embed: true` there
  is no such change. Read-back checks the part decodes to exactly those files and
  that the package gained nothing else.
- **Provenance tags** (D6, `project/provenance.ts`, `tags.ts`): after planning, every
  name the build leaves in sync gets its comment as the file should carry it: a name
  declared in a module file (`names/FN.xln` → `FN`; not `_unmanaged.xln`, not
  `names/sheets/`) ends its comment's last line with `[xln FN 1.2 #3f9a1c]`: module,
  version from a `// @version 1.2` header line of the module file (none: `[xln FN
  #3f9a1c]`), and the first 6 hex digits of SHA-256 over the stored definition
  (normalised as the lockfile hashes it, numbers by value) and the comment without its
  tag; a tag made before numbers counted by value (`sourceHashV1`) is still accepted.
  A name whose source entry carries `@from(lib #353921)` (its **library base**, decided
  2026-10-07) gets it after the hash: `[xln FN 1.2 #636cf1 lib#353921]`; the `#` hash
  stays what was built, `lib#` is the library version the copy came from. Adding,
  changing or removing `@from` is therefore a `provenance` change too, never a
  definition change or a conflict. Pull writes `@from(lib #…)` back from the tag (`ProjectName.libBase`); a tag
  without `lib#` gives no annotation. The pull guard (`unbuiltEdits`) lists a `@from`
  the workbook's tag does not carry yet (only where the name has a tag). Plan, lockfile,
  pull and read-back compare comments with the tag stripped, so a tag change is never an
  edit or a conflict; a tag that is missing, stale or no longer due is a `set-name` with
  field `provenance`. A comment that would pass 255 characters (line breaks counted as
  CR LF) with its tag gets none, and the build warns; any comment over 255 refuses the
  build (`checkCommentLengths`: Excel will not open the file, F5). Comments are read with
  their `_xHHHH_` escapes decoded (`file/xstring.ts`): Excel saves a line break as
  `_x000a_`. Names Excel changed since the last
  pull keep their tag: its stale hash is how pull tells them apart.
  `buildWorkbook({…}, { provenance: false })` turns it off.
- `buildWorkbook` ties them together and returns the new lockfile (in-sync names from the
  built file; names Excel changed keep their old entry; names Excel created stay out)
  and manifest. `builtCopyName` (`x.xln.xlsx`, browser) and `backupName`
  (`x.backup.xlsx`) name the files beside the workbook.

### `file/`

`readWorkbook` returns plain data (no classes, no Maps): sheets in `<sheets>` order with
their formula cells, shared-formula groups, spill anchors and saved extents, conditional
formatting and data validation formulas; every defined name with its scope by sheet
position, hidden flag, comment and the definition exactly as stored (entities decoded,
CR LF kept); Tables with their columns; chart parts (`charts`: per chart its `<c:f>` formulas, with
the element each belongs to, `val`, `cat`, `tx`, …, and the sheet and drawing that show
it, worksheet or chartsheet; Office 2016 `chartEx` parts too); the parts a rename must
check (charts, pivots, external links, custom XML); the links to other workbooks
(`externalLinks`: per `<externalReference>` its number `n` of `[n]` in stored formulas, the
file's name, the stored target and the part, probe F10), which `decompile` and `compile`
take as `links` to show `[1]Sheet1!A1` as `[Other.xlsx]Sheet1!A1` and back (an
`external-link` error for a workbook the file has no link to; without `links`, kept as
written). Trim references (`A1.:.A10`) are syntax: stored `_xlfn._TRO_ALL(A1:A10)` and
the like (F10). Non-fatal oddities go to `warnings`; an unreadable package
throws `XlsxError` (or `XmlError` for malformed XML).

`foreignModuleStores` (`file/afe.ts`): another tool's copy of the names as module text,
read only. So far Microsoft's Advanced Formula Environment (AFE, Excel Labs): its custom
XML part, recognised by the namespace of its root element
(`http://schemas.advancedformulaenvironment.officeapps.live.com/…`; `AFEJSONBlob` in
`…/afejsonblob/1.0` is read: base64 of the UTF-16LE JSON store, modules under
`/projects/<Module>`, the names AFE exported, the locale), the item's ID and whether the
add-in's settings point at it; AFE 1.0's very hidden code sheet and the locale-detection
sheet by their fixed names. Format and sources: `probes/README.md` § AFE-saved workbooks.
Pull, check (C14) and build report it; nothing writes it.

The XML reader (`file/xml.ts`) is our own: generic parsers normalise CR LF to LF in text,
which would hide what Excel stored, and worksheets are streamed rather than built into
a tree. Shared-formula children are not expanded here: `formulaTextAt(sheet, f, shift)`
takes the reference shifter from the formula layer (`shiftFormula`).

### `lang/shiftFormula`

`shiftFormula(text, dRow, dCol)` is the formula Excel means `dRow` rows down and `dCol`
columns right of where `text` is written (a shared-formula child, a fill). It rewrites
`ref` nodes of the AST only, so strings, structured references, names (qualified or
not), function names and LET/LAMBDA variables never move. Relative parts move, `$`
parts stay; cells, areas (also the trim forms `:.` `.:` `.:.`), whole columns and whole
rows; references on other sheets and in other workbooks move like any other, as Excel
fills them. A reference moved off the sheet becomes `#REF!` (keeping its sheet:
`S2!#REF!`), and so does a spill of it (`A1#`). Layout is kept; text the parser rejects
is shifted on its tokens. `shiftAddress(address, dRow, dCol)` moves one A1 address.

### `lang/`: a cell called like a function

A cell can hold a LAMBDA (`C2: =LAMBDA(a, b, a+b)`), and Excel calls it through the
reference: `=C2(D2, E2)`, stored exactly so (no prefix). A name on the cell works too
(`Fn` = `S!$C$2`, `=Fn(1, 2)`), as does a name holding the LAMBDA itself. The parser reads
`C2(…)` as an `invoke` node whose callee is the `ref`, like `LAMBDA(x, x)(1)`; compile and
decompile treat the reference like any other (a sheet's name qualifies it), so the round
trip is exact. The rule for a word directly followed by `(` (no space: `C2 (D2)` is still
an intersection):

- it looks like a cell (`C2`, `$C$2`, `C$2`, `XFD1048576`) and has `$` or a sheet
  (`Sheet1!C2(`, `'S 1'!$C$2(`, even `Sheet1!LOG10(`): a called cell;
- it looks like a cell, bare, and the function catalogue has it: the built-in (Excel
  resolves `LOG10(` / `log10(` to the function; LOG10 is the only catalogue name shaped
  like a cell: `ATAN2`, `DAYS360` have four letters, `T` no row);
- it looks like a cell, bare, and the catalogue lacks it: a called cell;
- anything else: a function or a defined name called.

The checks (`checkStoredForm`, the compiler, the checker, the audit) never call it an
unknown function. C6 checks the call against the cell (`cellCallee` in
`project/classify.ts`): a cell whose formula is a LAMBDA (also as a LET's result) takes so
many arguments (`C6.lambda-arity`, an error as for a named LAMBDA); a cell holding a value
or nothing, or a formula whose result cannot be a LAMBDA (a constant, arithmetic, a
built-in other than those that can return an argument or a reference such as IF, CHOOSE,
INDEX, INDIRECT, XLOOKUP), is a warning `C6.not-a-lambda` (Excel gives #VALUE! or #CALC!);
any other formula is not judged. The checker reads the cell from its cell statement; with
no statement it says nothing unless the last pull recorded the sheet's cells (then the
cell holds a value or nothing). The dependency graph has the edge to the cell, as for
any reference.

### `project/`

`pullProject` does no file I/O: it returns `files` (project-relative path → text, LF line
endings) for the caller to write, plus a `report`.

- **Every pull is fresh** (decided 2026-10-06; the merge of 2026-10-05, `merge.ts`, is
  gone): there is no live link between Excel and the editor, so the workbook is the only
  real state when one pulls. `pullProject(bytes, name, { width })` writes the project
  from the workbook alone, in a fixed structure (modules by prefix, sheet files by
  sheet: a name moved by hand goes back to its place), and never reads the embedded
  part. The caller replaces `names/**`, the lockfile and the manifest (and keeps
  `xln.config.json`). Before it does, **`unbuiltEdits({ workbook, fileName, files })`**
  (`build/unbuilt.ts`) lists the source edits a pull would replace: the plan's
  `sourceEdits` against the lockfile and the workbook as it is (without a lockfile, the
  source against the workbook itself), each with its file and line and the change the
  build would make (`describeChange`), plus the source's errors (a source the build
  refuses was never built); an unreadable lockfile is one item. `formatUnbuiltEdit` gives
  one line. The CLI refuses while there are any (`--discard` replaces them); the
  extension asks (*Build first*, *Discard and pull*).

- **Classify** (A3, `classify.ts`) on the syntax tree: `constant` (literals, constant
  arrays), `range` (references, unions, intersections), `spill` (`x#`, with the anchor),
  `table` (`Tbl[...]`), `lambda` (arity, optional parameters), otherwise `formula`;
  `unparsed` if the parser rejects it (written as stored, with a warning). `_xlnm.`
  names go to the manifest only; other `_xl*` names are dropped.
- **Modules** (A4, `modules.ts`): dot prefix (`FN.PICK` → `FN`); or an upper-case
  underscore prefix (`IN_Rate` → `IN`) when the text after `_` starts upper-case and at
  least three names share it. Deterministic. Files under `names/`:
  - `<Module>.xln`: the module's names, workbook scope first, then its sheet-scoped
    members in one `@scope(Sheet)` block per sheet;
  - `_unmanaged.xln`: workbook-scoped names that no module owns;
  - `sheets/<Sheet>.xln`: one file per sheet (in sheet order): its cell statements in
    sheet order, then its local names that are not cell statements and that no module
    owns. No blocks (M3d): every name is local to the file's sheet unless `@workbook` is
    on the line above it, and addresses are bare.
- **Cell statements** (M3b, `statements.ts`): `cellStatements(wb, { values })` lists each
  worksheet's statements in sheet order: `named` (a name, sheet- or workbook-scoped,
  defined as one cell of the sheet, `$C$6` or `$C$6#`), `slot` (such a name on an empty
  cell: no formula, no value per `readCellValues`, not inside a spill) and `unnamed`
  (every other formula cell; cells equal after `shiftFormula` merge into row runs, then
  into rectangles: `@B40:G40`). Spill anchors are one cell; ghost cells and data tables
  are not statements. A name that is a statement gets `ProjectName.cell` and is written
  in its sheet's file whatever its module; a workbook-scoped one sits at its place with
  `@workbook` above it. `sheetFormulaCells` and `sameFilled` are the
  pieces the plan reuses.

  The sheet file name keeps spaces and letters (`SCF recursive.xln`) and percent-encodes
  (UTF-8) what file systems refuse or alter: `/ \ : * ? " < > |`, control characters, `%`
  itself, a leading `.`, a trailing `.` or space, and the first letter of a Windows device
  name (`CON` → `%43ON.xln`). `sheetFromFileName` reverses it. File names that macOS or
  Windows would take for the same file (case, Unicode normalisation: `fileSystemKey`)
  get `~2`, `~3`; such a file names its sheet with `@sheet(Name)`. `sheetOfPath` gives a
  sheet file's sheet, `parseSourceFile(path, text)` parses a project file by its path.
- **Module files** (`module.ts`): `/** comment */`, `@hidden`, then `name = formula;` in
  display form (A2, CR LF → LF in layout). In a module file `@scope(Sheet)` opens a block:
  every name after it is local to `Sheet` until the next `@scope(…)` or `@workbook`, which
  returns to workbook scope (a file starts in workbook scope). In a sheet file (M3d)
  `@workbook` is an annotation of the next statement (`form: "sheet"`); a sheet file
  with `@scope` lines is read in the old block form (`form: "sheet-blocks"`), and
  `convertSheetBlocks(text, sheet)` rewrites it (the extension's quick fix). The sheet is quoted only when this
  syntax needs it (blank, spaces, parentheses, quotes): `@scope(BS)`,
  `@scope('SCF recursive')`. `@hidden` applies to the next name only, as does
  `@from(lib #353921)`, the library base (`ModuleEntry.from`, 6 hex digits; `formatEntry`
  writes it after the doc comment and the other annotations; `parseLibBase` says why a
  malformed one is not one, and the checker reports it as an error, a doubled one too,
  and warns when it sits outside a module file, where no tag carries it). Annotations
  carry their offsets (`offset`, `end`). A definition longer
  than the line is pretty-printed (A5) unless Excel stored it on several lines already.
  `parseModule` reads the files back. For the editor it also records offsets: each
  entry's extent (`start` from its doc comment, `end` after `;`), the name (`offset`), the
  block directives (`scopes`), and where the formula text lies in the file
  (`formulaMap`, read with `formulaToSource` and `sourceToFormula`; comments inside a
  formula are removed from `formula`, so the map has a run per stretch between them).
- **Manifest** (A6, `manifest.ts`): sheets, Tables, spills (and 1×1 dynamic-array cells
  apart), built-in names, other parts, and per name its kind, module, the names it uses
  and where it is used: other names, cells (merged into rectangles), conditional formats,
  data validations, Table column formulas. Name resolution follows sheet scope and
  LET/LAMBDA shadowing (`nameUses` gives each use with its span; `NameResolver` resolves it).
  `parseModule` also reads cell statements (`Name @C6 = …;`, `Name @C6 = ;`, `@C5 = …;`,
  `@B40:G40 = …;`, `@'S 1'!C6`, `Name @C6# = …;` for a name on the spill): named ones
  are in `entries` with `cell`, all are in `cells` (unnamed ones have name `""`); `cell`
  records the address's offsets for diagnostics and quick fixes and its `spill` (`#`,
  refused on unnamed statements and ranges). An empty right-hand side is allowed only there.
- **Lockfile** (A7, `lockfile.ts`): per name, a hash of the definition's tokens without
  layout and of the comment. Format 2 adds `cells`: per statement (key: the name's, or
  `Sheet!Range` for an unnamed one) its sheet, range and formula hash (null for a slot).
  Format 3 hashes number literals by value (`canonicalNumber` in
  `lang/format.ts`: 15 significant digits, then JavaScript's shortest form; Excel saves
  `1E-14` as `0.00000000000001`), writes `h:` + 16 hex digits of SHA-256 (`hash.ts`,
  pure TypeScript) instead of `sha256:` + 64, and one line per entry (`lockfileText`).
  Formats 1 and 2 are still read: `definitionHashV2`/`commentHashV2` keep their hash;
  every comparison takes the function of the hash it compares with
  (`definitionHashLike`), and the plan upgrades an old hash the source still matches
  (`upgradeLockfile`). Format 4 (`LOCK_FORMAT`, same hashes) says that the sheet
  files write `#` for a name on the spill; for formats 1–3 (`explicitSpill` false) the
  plan reads a missing `#` on a name on the spill in the lockfile and the workbook as
  `#` (`BuildPlan.implicitSpill`), and the build keeps format 3 while it does. The
  manifest gives each named
  statement's `cell` and lists unnamed ones per sheet (`unnamedCells`).

### `view/`

`sheetFormulaView(wb, sheet, names = workbookNameIndex(wb), { cache })` lists the
sheet's formulas in **order of appearance**, row by row and left to right. One line per
formula Excel shows on its own: a dynamic array once at its anchor with the extent saved
with the file (`rows`, `cols`), a legacy array or data table once at its top-left cell,
every cell of a shared group with its own text (the master's moved by `shiftFormula`).
A formula cell inside another cell's array, spill or data table is not listed (Excel
saves an empty `<f/>` in spilled cells of a volatile formula). Each line has `kind`
(`normal`, `shared`, `array`, `dynamic-array`, `data-table`), `stored` (the text at that
cell), `formula` (display form: no stored prefixes, `x#`, LF line breaks; a data table
reads `TABLE(row input, column input)`), `value` and `valueText` (short: 15 significant
digits at most, strings quoted and cut at 40 characters, errors as Excel shows them),
`names` (each use with its key resolved with the sheet as home, sheet scope first, and
LET/LAMBDA variables shadowing names; `key` undefined for a name that does not exist)
and `refs` (cell references with the sheet they point at), both as spans in `formula`;
`lhs`, the defined names whose definition is exactly this line's location: the cell
(`$C$17`, with or without `$`), the anchor's spill (`$C$6#`, stored `ANCHORARRAY`) or
the saved extent as a fixed range (`$C$6:$G$6`, a legacy array's range), on this sheet
(the sheet written in the definition, or the name's own sheet when none is). A name over a
larger block, or a shared formula's group, is not listed per cell; `_xl*` and `_xlnm.`
names never are. Each has `key`, `display` (bare for this sheet's and workbook names,
`Sheet!Name` for another sheet's), `hidden` and `target` (`cell`, `spill`, `extent`),
sorted by `display`. The reference is read by `definitionTarget` (`project/classify.ts`),
from the parse tree; the names are parsed once per workbook.
`names` is any `NameIndex` (`NameResolver` is one), so the editor can resolve against a
project's live definitions. Cell formulas are decompiled without a home sheet: Excel
shows a cell's own-sheet qualifier when the author wrote one, so the view keeps it.

`renderFormulaView(lines, { sheet, workbook?, formulaWidth = 64, width = 110, nameColumn = true, nameWidth = 32 })`
returns the text and, per line, where its left-hand names, address, formula, names and
references sit in it (document offsets; `start` is the entry's first character). Columns:
the names on the cell (joined with `, `; as wide as the longest list up to `nameWidth`,
a longer list stands on a line of its own above its entry; no column when the sheet has
none or `nameColumn` is false), cell (`C6#` for a dynamic array: the reference to its whole
spill), kind (`(r×c)` dynamic array and its saved size, `{r×c}` legacy array,
`shared ×n` a group's master, `shared ← B2` a child, `table r×c` a data table),
`= formula`, `→ value` (` …` after the anchor's value of a spill: the other values are
not read). A formula longer than `formulaWidth` stays on its line when its pretty-printed
form is one line; otherwise it gets a block, pretty-printed to `width` (`prettyPrint` with
`pack`: short arguments share lines, as in `MAX(C6, C23, …)`; LET bindings one
per line) unless Excel stored it on several lines already. An empty row in the sheet,
and a block, are set off by a blank line.

**Calculation order** (B6 b, `calc.ts`). `sheetCalcView(wb, sheet, names?, { cache, graph })`
gives the same lines as `sheetFormulaView`, ordered by the dependency graph (below): each
formula after everything it reads, on any sheet and through names; among formulas ready
at the same time, order of appearance, so a sheet laid out top-down keeps its order and
only what must move moves. `workbookFormulaView(wb, names?, { graph })` lists every
formula of the workbook in one such order (ties: sheet order, then row, then column).
Each line then also carries `level` (the longest chain of formulas from the inputs:
1 reads only inputs; names add no level), `cycle` (the circular reference it belongs
to; its members stand together) and `dependsOn` (what it reads directly, as seen from
its sheet: defined names first in the order written, then cells). Pass `graph` to reuse
one graph across sheets. `renderFormulaView` takes `levels: true` (a first column with
the level, `5 ↻2` for a member of circular reference 2, and a comment line before each
circular reference) and `sheets: true` (addresses written with their sheet, `IS!C16#`,
for the workbook view).

### `graph/`

`buildGraph(wb, names = workbookNameIndex(wb), { cache })` builds the cell dependency
graph from the snapshot alone. An edge A → B reads "A depends on B". Nodes:

- **formula**: one per line of the formula view (a dynamic array with the extent it
  spilled to when saved, a legacy array or data table, or one cell of a normal or
  shared formula), key `Sheet!C6`, label `BS!C6#`;
- **input**: the cells formulas read that no formula fills, one node per distinct
  rectangle as written (`Assumpt!C36`, `S!A:A`, `tblAssumpt[base]`), key `in:Sheet!A1:B3`.
  Grouping by reference keeps the graph the size of the formulas (a whole column is one
  node, not a million) and names inputs the way the author wrote them; an area partly
  covered by formulas gets edges to those formulas plus one input node;
- **name**: every defined name (except `_xl*`), key `name:Sheet!Name` or `name:Name`.

Resolution works on the syntax tree: cells, areas (also `A1 : B9` and the trim forms),
whole columns and rows (one query in a per-sheet index of formula blocks), `x#` and
`ANCHORARRAY(x)` (the dynamic array anchored there; `#REF!` flagged if none), defined
names with sheet scope (via `names`, so a project's live definitions can be used), with
LET and LAMBDA variables shadowing names and making no edge, structured references
(`Tbl[Col]`, `[#Headers]`, `[#Totals]`, `[#All]`, `[[A]:[B]]`, `[@Col]` in the formula's
row) and a Table's own name (its data rows), 3-D references (`S1:S3!A1`, every sheet
between by position), cross-sheet references, `OFFSET` with literal arguments (the
moved rectangle), and data tables (their input cells, the formulas in the row above and
the values in the column to their left). What the file alone cannot resolve is recorded
on the node as a flag, never guessed: `dynamic` (`INDIRECT`, `OFFSET` with computed
arguments, whose first argument is then not an edge; a relative reference in a name,
which moves with the cell that reads it; an unqualified reference in a workbook-scoped
name), `external` (another workbook), `broken` (`#REF!`, an unknown name or Table or
column, a formula that does not parse). Each flag also has a stable `code` (`ref-deleted`, `unknown-name`, `relative-in-name`, … : `GraphFlagCode`) for tools such as the audit.

`DependencyGraph`:

- `nodes`, `edges`, `node(key)`, `nameNode(key)`, `formulaAt(sheet, cell)` (the block
  holding a cell), `precedents(n)` / `dependents(n)` (direct), `allPrecedents`,
  `allDependents`, `inputsOf(n)` (input cells and constant names it depends on);
- `cycles`: strongly connected components (Tarjan, iterative: chains run thousands deep)
  of two or more nodes, or a node reading itself; members in sheet, row, column order,
  names after formulas; each member's `cycle` holds its number. LAMBDA names calling
  themselves or each other are `recursions`, not cycles (Excel allows them);
- `level` on every node: 0 for inputs, one more than the highest formula it reads for a
  formula, the level of what it reads for a name; members of a cycle share one;
- `order({ sheet? })`: every node once, each after what it depends on (cycles together):
  Kahn's algorithm on the components with a heap keyed by sheet position, row, column, so
  the order is stable. With `sheet`, the best order for that sheet: other sheets'
  formulas, names and inputs come as early as they can and never hold a formula of the
  sheet back;
- `flagged()`, `stats()`;
- `spillRefs` (C9): fixed references to cells of a dynamic array (`C10:G10`) where the
  spill reference would follow it as it grows, in formulas and definitions: the formula
  or name, the reference as written, the spill, `use` (`C10#`, or `INDEX(C10#, 3)` for
  one cell) and `fit` (`exact`: the saved extent; `part`: inside it; `beyond`: reaching
  cells outside it too). A reference to the anchor cell alone is not a finding;
- `unusedNames()` (C10): `unused` (nothing reads it: no formula, no other used name, no
  conditional format, data validation, Table column formula or built-in name) and
  `onlyByUnused` (read only by unused names). Charts are not read here; the audit's C10
  counts them;
- `nameCycles()` (C12): cycles among names through their definitions alone, with
  `recursive` when all members are LAMBDAs.

The spill extents are those saved with the file: a spill that grew since is not seen
grown (FEASIBILITY §11.3). On the corpus a graph builds in 15–30 ms of CPU.

### `library/` (M4)

A library is a folder of `.lambda` files (the author's `_shared/lib`, which the Python
build also reads, so the format is read as it is). The caller lists the folder; the core
gets file name → text.

- `parseLambdaFile(path, text, known?)` → `{ fn?, problems }` (never throws; problems
  carry 1-based lines). The header is the run of `#` lines up to the first blank `#`: a
  field is `# key value` with one blank after `#`, a continuation line is indented
  further (`impromptu`, `example` keep their lines; `summary` is joined into one line).
  The `#` lines after it are the rationale; the first non-comment line starts the
  definition (display form). Checked: a name; a single LAMBDA that parses and compiles
  (an unknown function is an error, as Excel would poison it as `_xludf.`); `params`
  equal to the LAMBDA's parameters (as the Python reader checks); the file name against
  the name (warning). `fn` = `{ name, summary, params, paramDocs, fields, rationale,
  definition, definitionLine, doc, docShortened, path }`.
- `doc` is the doc comment xln writes into the project (and so the Name Manager
  comment): the summary, then `@param name` per parameter (with a description when a
  rationale line `@param name text` gives one: xln's addition, which publish writes and
  the Python build ignores). It stays within `LIBRARY_DOC_MAX` = 212 characters, which
  leaves room for the build's provenance tag with its library base
  (` [xln FN 1.2 #3f9a1c lib#353921]`, 32 characters, 11 more for longer module names and
  versions; it was 223 before the base): over it the summary is cut at a word with
  `…` (`docShortened`, and a warning); the definition is never touched. Of the 17 real
  files only FN.FIXPOINT's three-line summary is cut.
- `readLibrary(files)` → `Library` (`functions` by name, `problems`, `get(name)`,
  case-insensitive). `libraryClosure(lib, name, has)`: the function and the library
  functions it calls that `has` says are missing, callees first.
- `libraryStatus(lib, copies, meta)` with `workbookCopies(snapshot)` or
  `projectCopies(files, snapshot?)` → `LibStatusReport`. Per library function:
  three-way on the copy's **library base** (2026-10-07): the source's `@from(lib #…)`
  for a project, the tag's `lib#…` for a workbook; versions are `libraryHash` of the
  stored definition (6 hex digits, normalised like the lockfile's: layout and number
  spelling do not count; `libraryFunctionHash(fn)` for a library function,
  `definitionBase(display, name)` for any LAMBDA text). **identical** (the copy equals
  the library, whatever its base), **outdated** (the copy is its base; the library
  moved), **modified** (the copy moved; the library is still the base), **both
  changed** (both moved, differently), **differs** (no base recorded: which side moved
  cannot be told; the tag's own hash no longer decides, since after a local edit and a
  build it matched the edited copy and read as outdated), **missing**; plus **local only**, a module's
  LAMBDA (`FN.X`, or a name in `names/FN.xln`) the library lacks, of a module the library
  has (a publishing candidate); other modules' LAMBDAs (`IN.*` input readers) are only
  counted per module (`otherModules`, one line in the text). Only definitions
  decide: the doc comment is generated from the header and may be shortened. A
  project's copy is its source compiled again, its base its `@from`; the tag shown comes
  from the workbook it was built into (pass the snapshot). Items carry `base`,
  `libraryHash`, `copyHash`; a both-changed item carries `baseDiffs` (base → copy, base →
  library) when a version at hand has the base's hash (`LibCopy.known`, in this order:
  the project's kept `library-bases/<hash>.json` (`withBases`; `projectLibraryStatus`
  reads them among `files`, or takes `bases`), the workbook's copy for a project, a
  snapshot added with `withKnownVersions` / `projectLibraryStatus`'s `backup`, the CLI's
  `<name>.backup.xlsx`), and `baseSource` says which; `threeWayText(item)` renders it, or
  the copy against the library when the base's text is not at hand. An identical item
  with no base carries `noBase` and the note *no base recorded: Record library base*.
  `projectLibraryStatus(lib, files, snapshot?, meta)` is that for a project, and with the
  snapshot each item also says whether the workbook has it as the source does
  (`unbuilt`: `new`, `edited`, or `deleted` in the source and still in the workbook; the
  report's `unbuilt` counts them and its text says "not built yet"), so a status of the
  source never looks done while the workbook is not (feedback 2026-10-07). The
  library's git history is not used. `renderLibStatus(report)` → `{ text, links }` (the
  CLI's output and the editor's document; diffs `-` copy, `+` library, both sides
  pretty-printed the same way first), `libStatusJson(report)`.
- Edits (`edit.ts`), returned as text edits for the caller to show and apply:
  `libraryInsertion(lib, files, name)` → `{ names, edits: FileEdit[] }` (the closure,
  each at its name-order position among the workbook names of `names/<Prefix>.xln`,
  before any `@scope` block; the file created with a one-line header when missing;
  each with `@from(lib #…)`, its library version);
  `libraryReplacement(text, entry, fn)` (the entry's doc comment and definition become
  the library's, and `@from(lib #…)` is written or rewritten; other annotations, name and
  trailing comments stay); `libraryBaseEdit(text, entry, hash)` (the edit that records a
  base; Publish uses it with `publishedBase(formula, name)`, the version the library will
  read back); `publishLambda(source,
  existing?)` → `{ path, created, text, changed }`: a new file with `# name`,
  `# summary` (the doc comment's first sentence; the rest goes to the rationale with the
  `@param` descriptions), `# params`; or the existing file with only the definition
  replaced (and the summary, params and `@param` lines when they changed; a summary
  shortened with `…` counts as the library's). A definition equal modulo layout is not a
  change, so publishing an unchanged function writes nothing. Its `base` is the version
  the file holds afterwards. `libraryBaseRecordings(lib, files, names?)` (*Record library
  base*): the `@from` edits for identical entries with no base (all, or those named; the
  others skipped with the reason), each with its base's text.
- The kept bases (`bases.ts`, author's idea 2026-10-07): `@from` is only a hash, so the
  explicit actions that write it (Insert, Take, Publish, Record library base) also hand
  the caller a `LibraryBase` (`hash`, `name`, `library` file, `stored`, `display`;
  `libraryFunctionBase(fn)`, `libraryBase(name, display, file)`; `libraryInsertion`'s
  `bases`, `publishLambda`'s `base`), kept by the caller as `library-bases/<hash>.json`
  (`BASES_DIR`, `baseFilePath`, `baseFileText`, `baseFiles`). `parseBaseFile` refuses a
  file whose text does not give its hash; `readBases(files)` → by hash. A pull writes
  nothing there, so it keeps them; nothing prunes them.

### `audit/`

`audit(wb, opts)` runs checks C1–C13 of the brief (§4 C), C14 (another tool's copy of the names) and C15 (labels that no longer give their names; it needs `opts.values`, the cells' values from `cellValueMap(bytes)`, and says nothing without them), on the snapshot and
returns an `AuditReport` (plain data, `format: "xln-audit/1"`): `findings`, `counts` and
`byCheck` per severity, the name census (`census`, C8) and the spill census (`spills`,
C9). A finding is `{ check, rule, severity, where, message, hint?, data? }`; `where` is
`{ kind: name | cell | cf | dv | table | chart | part, sheet?, name?, key?, ref?, range?, span?, text? }`
(`span`/`text`: the part of the stored formula it is about; a shared formula reports once,
at its master, with its group as `range`). Findings are ordered by check, then place
(names, cells, formats, validations, Tables, charts; sheet order, row, column), then
rule, so the same workbook gives the same report on any machine.

| Check | Rules (default severity) | What fires them |
|---|---|---|
| C1 | `syntax` (error) | a definition, cell, format, validation or Table column formula that does not parse; line and column |
| C2 | `bare-prefix` (error), `poisoned` (error), `wrong-prefix` (warning), `unknown-function` (info) | the decompiler's prefix diagnostics on the stored form (probe F6) |
| C3 | `lambda` (error), `called` (warning) | a LAMBDA named like a built-in (T12: every call reaches the built-in); a non-LAMBDA name with a built-in's spelling that formulas call |
| C4 | `ref-deleted`, `unknown-name`, `table`, `no-anchor`, `broken`, `error-definition` (errors); `relative`, `unqualified-ref` (warnings); `external` (info) | the graph's flags (`GraphFlag.code`) on names and cells; a definition that is an error; a relative reference in a name (decision 2026-10-04: flagged, not followed) |
| C5 | `own-sheet`, `other-sheet` (errors) | an unqualified read, in a stored definition, of the definition's own sheet's name (Excel stores it qualified; unqualified it is `#NAME?` in functions, findings §11.3), or of a name that exists only on other sheets |
| C6 | `lambda-arity`, `not-a-function` (errors), `not-a-lambda`, `builtin-arity` (warnings) | calls in names and cells: a LAMBDA name, or a cell holding a LAMBDA (`C2(…)`, a name on C2 called), with too few or too many arguments (`[x]` optional); a constant or Table name called (error); a called cell that holds a value or a value formula, or a called range of several cells (`not-a-lambda`); a built-in outside the catalogue's arity |
| C7 | `length`, `nesting` (errors), `length-near`, `nesting-near` (warnings) | display length ≥ 8,192 (warn from 7,500; probe T15: 7,992 accepted, 8,192 rejected), function nesting over 64 (warn over 48); names and cells; `opts.limits` |
| C9 | `fixed-ref` (warning) | the graph's `spillRefs` in cells and in names: a name over a spill's fixed range is one (decision 2026-10-04) |
| C10 | `unused` (warning), `unused-cell`, `unused-hidden`, `only-by-unused` (info) | the graph's unused names, minus the names chart formulas read (`[0]!Name`, `'Sheet'!Name`), the harness (`opts.harness`), end results, and what those read. An **end result** (M3c) reads something and is read by a person: its definition computes from cells, names or Table columns, or it names cells holding formulas (`Unlevered_net_income @C11# = EBIT - Taxes`). A constant, an uncalled LAMBDA and a name on value or empty cells that nothing reads are still unused; the name on cells is reported as `unused-cell`, **info** (decided 2026-10-07: a person may read it on the sheet, like a displayed `SelfTest`), the others as `unused`, warning. A name meant to be read by a person or a solver that reads nothing itself (a solver cell) is silenced by listing it in the harness, not by lowering the rule |
| C11 | `drift` (warning) | families of names differing by a coordinate tag (C8): members whose formula differs from the majority beyond the tag; a name over a cell or spill is compared through its anchor's formula, relative references as offsets; constants are not compared. The same comparison decides C8's inferred tags (decided 2026-10-07): drift is one copy departing from copies that otherwise agree, so a family whose members all differ under a tag nothing else vouches for (`initial_amount`, `final_amount` in is-model: an opening and a closing balance) is not a family and gives no finding |
| C12 | `name-cycle` (error), `circular` (info) | the graph's name cycles (LAMBDA recursion allowed); circular references among cells |
| C13 | `constant` (info) | numbers in LAMBDA bodies outside `DEFAULT_CONSTANTS` (−1, 0, 1, 2, 0.5, 10, 100, 1000, 7, 12, 24, 52, 60, 360, 365, and their negatives) and below the sentinel magnitude (`DEFAULT_SENTINEL_ABOVE`, 1E+90: `1E+99`, `1E+300` are "infinity", not constants); `opts.constants.allow`, `opts.constants.sentinelAbove` |
| C14 | `afe-store`, `afe-absent`, `afe-not-compared`, `afe-unreadable`, `afe-code-sheet`, `afe-locale-sheet`, `afe-drift` (info) | the module store of Microsoft's Advanced Formula Environment (Excel Labs) in the workbook (`WorkbookSnapshot.foreignModuleStores`, `file/afe.ts`; `where.kind: "part"`): what it holds; names whose AFE text differs from the workbook's definition (`afeStatus`: display forms modulo whitespace, case of names, number spelling; a sibling named without its module counts as `<Module>.name`); names AFE defines that the workbook lacks; text in other separators (not compared); a store xln cannot read; AFE 1.0's code sheet; AFE's locale-detection sheet. xln never changes the store |
| C15 | `label-drift` (info) | `labelDrift` (`project/labels.ts`), with `opts.values`: a name on a row (column, single cell) whose label just left (above) is typed text that Create from Selection turns into another name (`labelName(v, shown?)`, Excel's rule as probe F11 measured it: trimmed, one `_` per character a name cannot hold, those at the end dropped, `_` in front of a first character that cannot start a name (`_2024_sales`, `_€uro`), `_` behind an A1-reference-like text or TRUE/FALSE (`Q1_`, `True_`), in front of an R1C1 one (`_R1C1`), no name from a number or past 255 characters, a date by its shown text; `labelNames` adds the common short-date spellings of a whole number and, not measured, leading invalid characters dropped), when other names from the same column (row) of labels over the same span match their labels, at least as many as those that don't. Not: a name a label of it gives (left, above, the corner), an alias on the same cells that the label gives, a computed label (the pull's note), a row or column of text (a header). `data.renamedFrom`: the label's name when no name has that spelling (decided 2026-10-07: renaming after Create from Selection). `data.fix`: the Find & Replace for the hint (`labelReplaceShort`). The build's label notice for every rename: `labelNotice(wb, values, { from, to, scope })`, `renameLabelNotices`, `labelNoticeLines` (`project/labelNotice.ts`) |

**Rules are data** (`rules.ts`, brief §10): per rule id its check, default severity, title,
and the message and hint as templates (`{fn} is stored without its {prefix} prefix…`)
filled from the finding's `data`. The checks decide where a rule fires; the table what it
says. `opts.rules` sets a severity per rule (`"C10.unused": "info"`) or per check
(`C13: "off"`); `opts.only` and `opts.minSeverity` filter.

**The harness** (`opts.harness`): globs on keys (`Check!*`, `CHK.*`, `Model!Fix*`,
`*!FixedPoint_*`; `Sheet!Name` or `Module.Name`, any case) for names that exist to be
read by a person or a solver. C10 treats them as roots (like chart series: neither they
nor what they read are unused) and the C8 tiers leave them out (with `census.exclude`).

**Project settings** (`config.ts`): a project folder may hold `xln.config.json`
(`CONFIG_FILE`), `{ "audit": { "harness": [globs], "rules": { "<rule or check>": "off" |
"info" | "warning" | "error" }, "constants": { "allow": [numbers], "sentinelAbove": n } } }`.
The core reads no files: `parseConfig(text) → { config, problems, notes, issues }` (never
throws; what cannot be used is left out and listed; a key that is not a setting where it
is written is a problem saying where it belongs, from `CONFIG_KEYS`, every setting and
its section: `audit.library: \`library\` is a top-level setting, not an audit setting:
move it out of "audit"`; `issues` carry each one's key path, and `configKeyRange(text,
path)` finds the key in the text for the editor's mark; `notes` say what is ignored on purpose: an old
`"names": {"scope": …}`, M3d's setting, removed 2026-10-06, `NAMES_SCOPE_NOTE`),
`auditOptions(config.audit, overrides) → AuditOptions` (`constants.allow` adds to the
defaults; overrides win: the CLI's flags), `defaultConfigText()` (what `pull` writes when
a project has none). The CLI and the extension read the file and pass the options to `audit`.

**C8 name census.** Counts by kind (the classifier's) and by scope; coordinate tags and
families: a tag is a suffix after the last `_` (or a prefix before the first) that names
share in at least two families (`opts.census.minFamilies`, or `opts.census.tags` to
name them); `X_baseExp` is the tag `base` inside a longer name (family `X_*Exp`). An
inferred tag must also be one copies agree on (`nameFamilies(members, opts, agrees)`,
`audit` passing C11's comparison): at least one family carrying it has two members with
the same formula up to the tag. When every family carrying it that can be compared
differs, the tag is dropped, its names are separate line items (T2), and the families
are formed again without it (another tag left on fewer than `minFamilies` families goes
too). A family with nothing to compare (constants, input cells: scenario inputs like
`Volume_base = 100`) neither keeps nor drops a tag, so a tag carried only by such
families stays, as before. Tags named in `opts.census.tags` are not judged. The
tiers of names standing in for dimensions generalise `scope_census.py`/`census.py`
(excel-models SUMMARY §3.2), reading what a name is, not its prefix: **T1** LAMBDAs that
turn a coordinate into a lookup (called with a text argument, or looking a parameter up
with XLOOKUP, MATCH, …); **T4** names whose value is a SEQUENCE (also shifted, or wrapped
in EOMONTH, DATE, …, or made by a LAMBDA) and the non-LAMBDA names it is built from;
**T2** one per line item, per (scope, stem); **T3** the line-item names beyond the first
of each item; not counted: other LAMBDAs (library) and `opts.census.exclude` (globs on
keys, e.g. a check harness). With the harness excluded
(`Check!*, CHK.*, Model!Fix*, *!FixedPoint_*`) it reproduces the published counts:
lbo-ep02 9·123·118·3 = 253 of 329, lbo-ep03 7·106·102·3 = 218 of 281, lbo-ep03r
9·143·278·3 = 433 of 539 (T2 104 spellings). Also `sheetDuplicates` (a short name on
several sheets) and the text `explanation`.

**C9 spill census.** Every dynamic array that spilled when saved (sheet, anchor, extent,
size, formula), with the names over it (`spill` `x#`, `extent`, `part`, `anchor`);
per sheet the counts named by `x#`, only by a fixed range, or unnamed; and the number of
dynamic arrays saved as one cell.

`renderAuditReport(report, { maxPerCheck?, width? })` is the text the CLI prints and the
editor shows: the counts per check, the findings grouped by check, then the two census
sections; `links` gives where each place (name, cell, spill) is written. On the corpus the
audit takes 30–75 ms per workbook (graph included).

## Runtime dependencies

| Package | Why |
|---|---|
| [`fflate`](https://github.com/101arrowz/fflate) | Unzip in the browser and in Node: pure JavaScript, no dependencies, a few kB minified for the parts we use (inflate, and deflate for the parts a build rewrites). Also decodes UTF-8 without relying on `TextDecoder` typings. |
