# The `.xln` language

*Specification, version 0.2 (2026-10-07). For contributors and power users. The user
guide for modellers is [`USER-GUIDE.md`](USER-GUIDE.md).*

This document says what an `.xln` project is, what each file may contain, and what each
construct means in the workbook. "Must" and "must not" are rules the tools enforce (an
error refuses the build); "should" marks advice (a warning or a hint). Where the code and
this text disagree, the code is the current behaviour and the disagreement is a bug in
one of them: report it. Behaviour that is odd but real is listed in
[§14 Open issues](#14-open-issues-found-while-writing), with the ones fixed since.

The examples come from runs of the CLI on copies of `probes/results/*.xlsx`,
`probes/fixtures/traps.xlsx` and of the author's `is-model.xlsx` (a one-sheet income
statement with a Ratios sheet).

Contents:
1. [Model](#1-model)
2. [Project layout](#2-project-layout)
3. [Lexical structure](#3-lexical-structure)
4. [Grammar](#4-grammar)
5. [Name statements and annotations](#5-name-statements-and-annotations)
6. [Cell statements](#6-cell-statements)
7. [Scope](#7-scope)
8. [Formulas](#8-formulas)
9. [What a build does with each construct](#9-what-a-build-does-with-each-construct)
10. [Checks](#10-checks)
11. [Comments, provenance and limits](#11-comments-provenance-and-limits)
12. [The library](#12-the-library)
13. [Versioning, and differences from AFE](#13-versioning-and-differences-from-afe)
14. [Open issues found while writing](#14-open-issues-found-while-writing) (most fixed 2026-10-07)

---

## 1. Model

A workbook has defined names (the Name Manager) and cells. **Excel owns the layout and the
labels**: sheets, rows, columns, text cells, and which cells exist. **xln owns the names
and the formulas of existing cells.** An `.xln` project is the text form of both:

- every defined name of the workbook, with its scope, hidden flag, comment and
  definition;
- every formula cell of every sheet, and every name that sits on an empty cell (a
  *slot*).

Two operations connect the two sides. Neither changes a name, a scope or a formula on its
own: the source shows the workbook as it is, and only the author's edits change it.

| Operation | Direction | What it does |
|---|---|---|
| **pull** | workbook → project | Writes the project from the workbook alone. Every pull is fresh: `names/**`, the lockfile and the manifest are recreated. |
| **build** | project → workbook | Compares source, lockfile and workbook three ways and writes what changed only in the source. |

There is no live link with Excel. One side is edited at a time; each pull or build
carries the changes across.

## 2. Project layout

`xln pull book.xlsx` writes the folder `book.xln/` next to the workbook:

```
book.xlsx
book.xln/
  names/
    FN.xln                 a module: names with the prefix FN.
    IN.xln                 a module: names with the prefix IN_ (see 2.1)
    _unmanaged.xln         workbook-scoped names no module owns, not on one sheet's cells
    sheets/
      IS.xln               the IS sheet: its formula cells, slots, names on its cells, local names
      Ratios.xln
  xln.config.json          settings (yours; pull keeps it)
  xln.lock.json            internal: what the last pull or build saw
  workbook.manifest.json   internal: sheets, Tables, spills, usages
  library-bases/           internal: library definitions a copy came from
    353921.json
book.backup.xlsx           written by a desktop build: the file before it
book.xln.xlsx              written by a browser build instead of book.xlsx
```

### 2.1 Which file a name goes to

The structure is fixed. A pull puts every name in exactly one file:

1. A name that is a **cell statement** (it is defined as one cell of a sheet, `$C$6` or
   `$C$6#`) goes to that sheet's file, whatever its prefix or scope (§6).
2. A name with a **module prefix** goes to `names/<Prefix>.xln`. A prefix is:
   - the text before the first `.` (`FN.PICK` → `FN`); or
   - an upper-case word before `_` (`IN_Rate` → `IN`) when the text after `_` starts
     with a capital and at least three names share the prefix.
   The module's sheet-scoped members follow its workbook-scoped ones, each with
   `@sheet(Sheet)` on the line above it (§7.2).
3. A **sheet-scoped** name with no prefix goes to `names/sheets/<Sheet>.xln`, after the
   sheet's cell statements.
4. A **workbook-scoped** name with no prefix whose definition is fixed to one sheet's
   cells goes to that sheet's file too, with `@workbook` above it (decided 2026-10-07).
   Fixed means one cell, one rectangular range or one spill (`x#`), with `$` on every row
   and column and the sheet written: `Mortgage!$B$12`, `Mortgage!$B$12:$G$12`. Inside a
   spill or not, on an empty cell, a value or a formula: what *Create from Selection*
   makes. Excel's text is kept (`amount = Mortgage!$B$12;`).
5. Any other **workbook-scoped** name with no prefix goes to `names/_unmanaged.xln`:
   constants, formulas, LAMBDAs, relative references, unions, references across sheets,
   3-D references, whole rows or columns.

Excel's hidden helper names (`_xlfn.*`, `_xlpm.*` and other `_xl*`) are not written.
Built-in names (`_xlnm.Print_Area`, …) go to the manifest only.

A name moved by hand to another file goes back at the next pull; the checker says so as
you type (`file-placement`, §10), with a quick fix that moves it where the pull would.
Files are ordered: workbook scope first, then sheets in sheet order, then by name; sheet
files list cell statements row by row, left to right.

### 2.2 File names

- Sheet files keep spaces and letters (`SCF recursive.xln`). Characters a file system
  refuses or alters (`/ \ : * ? " < > |`, control characters, `%`, a leading `.`, a
  trailing `.` or space, a Windows device name) are percent-encoded in UTF-8
  (`a<b` → `a%3Cb.xln`, `CON` → `%43ON.xln`).
- Two files that macOS or Windows would treat as one (case, Unicode normalisation) get
  `~2`, `~3`. A sheet file so renamed carries `@sheet(Name)` to say its sheet.
- A module file is named after its prefix: letters, digits and `_`, starting with a
  letter or `_` (`FIN.xln`; `~2` on a clash). Its header says `// module: FIN, pulled by
  xln from book.xlsx.`, then `// @version 1.2` when its names' provenance tags carry a
  version (§11.2).

### 2.3 What may be in `names/`

Only `.xln` files, in `names/` (modules) and `names/sheets/` (sheet files). A build
refuses, and the editor flags with *Delete this file*:
- any other file (hidden files such as `.DS_Store` aside);
- a folder other than `names/sheets/`, or a folder below it;
- a sheet file for a sheet the workbook does not have ("sheets are created in Excel, then
  pulled");
- a module file not named after a prefix (`my module.xln`).

### 2.4 Who writes which file

| File | Pull | Build | Author |
|---|---|---|---|
| `names/**/*.xln` | recreated from the workbook | never written | edits |
| `xln.lock.json` | recreated | rewritten (desktop build only) | never |
| `workbook.manifest.json` | recreated | rewritten (desktop build only) | never |
| `xln.config.json` | written only when missing | read | edits |
| `library-bases/*.json` | kept, never written or pruned | read | never (library actions write it) |

The extension hides the lockfile, the manifest and `library-bases/` in the Explorer and
opens them read-only (*xln: Show project internals* reveals them).

### 2.5 `xln.config.json`

All keys are optional. An unknown key, or a known one in the wrong section, is reported
with where it belongs (`audit.library: \`library\` is a top-level setting, not an audit
setting: move it out of "audit"`), never silently dropped.

```json
{
  "audit": {
    "harness": ["Check!*", "CHK.*"],
    "rules": { "C10.unused": "warning", "C13": "off" },
    "constants": { "allow": [4, 1.5], "sentinelAbove": 1e90 }
  },
  "build": { "embed": false },
  "library": "~/models/_shared/lib"
}
```

| Key | Meaning |
|---|---|
| `audit.harness` | Globs on `Sheet!Name` or `Module.Name`, any case: names that exist to be read by a person or a solver. C10 does not report them; the C8 tiers leave them out |
| `audit.rules` | Severity per rule (`C10.unused`) or per check (`C13`): `off`, `info`, `warning`, `error` |
| `audit.constants.allow` | Numbers C13 accepts in LAMBDA bodies, besides the defaults |
| `audit.constants.sentinelAbove` | Numbers from this magnitude up are sentinels, not constants (default `1E+90`) |
| `build.embed` | Embed the project in the workbook as a custom XML part on every build (default `false`). An archive copy: pull never reads it |
| `library` | The library folder (§12): relative to the project folder, absolute, or `~/…` |

An old `"names": {"scope": …}` (removed 2026-10-06) is read and ignored with a note.

### 2.6 The lockfile and the manifest

`xln.lock.json`, format `xln.lock/4`, one line per entry:

```json
{
  "format": "xln.lock/4",
  "workbook": "is-model.xlsx",
  "names": {
    "ANA.GROW": { "definition": "h:fd3c33563b4879ef", "comment": "h:48b5e48cece7930c", "hidden": false },
    …
  },
  "cells": {
    "S1!C8": { "sheet": "S1", "range": "C8", "formula": "h:01e4ff9fe6d6d39b" },
    …
  }
}
```

A name key is `Name` (workbook scope) or `Sheet!Name`. A hash is `h:` and 16 hex digits of
SHA-256 over the stored form's tokens without layout, every number by its value (`1E-14`
and `0.00000000000001` hash alike), the provenance tag stripped from comments. Format 4
says the sheet files write the `#` of a name on a spill (§6.3). Formats 1–3 are still
read; see the CLI README.

The manifest (`workbook.manifest.json`) holds the sheets, Tables, the spill map, built-in
names, and where each name is used (names, cells, conditional formats, validations,
Table columns). The editor uses it for navigation; it reflects the last pull.

### 2.7 `library-bases/`

One JSON file per library version a copy came from, named by its 6-digit hash: the hash,
the name, the library file, and the stored and display text. Insert, Take, Publish and
Record library base write it (§12.4). It lets `lib status` show what each side changed
since the base. A pull leaves the folder alone; nothing prunes it.

## 3. Lexical structure

- **Encoding:** UTF-8. Line ends LF or CR LF; a pull writes LF. CR is whitespace.
- **Whitespace** (blank, tab, CR, LF) separates tokens and is otherwise not significant,
  inside formulas too: definitions are compared modulo whitespace (Excel re-spaces
  formulas). Line breaks inside a multi-line definition are stored as CR LF in the
  workbook; the author's layout stays in the source.
- **Number literals** are written to the workbook as typed (`0.10`, `.5`, `1E3`). Excel
  keeps the number, not the spelling, and writes its own on save (`1E-14` becomes
  `0.00000000000001`; `0.10` presumably `0.1`): comparisons and hashes take a literal by
  its value rounded to 15 significant digits, so that is no change in Excel, and the
  next pull shows Excel's spelling.
- **Line comment:** `//` to the end of the line. Allowed between statements and inside a
  formula (outside strings).
- **Block comment:** `/* … */`, not nested. Inside a formula it counts as a blank.
- **Doc comment:** `/** … */` (but not `/**/`) before a name statement: it becomes the
  name's Name Manager comment (§11). In a multi-line doc comment the leading `*` of each
  line and one blank after it are removed, and so are a first and last line that are
  blank. `*/` inside the text is written `*\/`. A statement has one doc comment: a second
  one before the same name is a warning (*two doc comments: only the last is kept*).
- **Annotation:** `@` followed by a word (letters, digits, `_`), optionally `( … )` on the
  same line. A quoted argument `('…')` may contain `)`; `''` is a quote.
- **Statement end:** a statement ends at the first `;` outside strings (`"…"`), quoted
  sheet names (`'…'`), brackets (`[…]`) and array braces (`{…}`). So
  `"a;b"` and `{1,2;3,4}` do not end a statement.
- **Comments are layout.** Pull recreates every file from the workbook: `//` and `/* */`
  comments and hand layout do not survive a pull. Only doc comments do, because they
  travel through the Name Manager. The pull guard (§9.5) names files that would lose
  comments or layout, but does not refuse for them.

## 4. Grammar

EBNF-like; `{ x }` is zero or more, `[ x ]` optional, `|` alternatives. Whitespace and
comments may appear between any two tokens.

```ebnf
file          = { directive | statement } ;

directive     = "@sheet" "(" sheet ")"          (* sheet file: names its sheet (IS~2.xln) *)
              (* files written before 2026-10-07 (module files) or M3d (sheet files): *)
              | "@scope" "(" sheet ")"          (* names after it are local to sheet *)
              | "@workbook" ;                    (* module file: back to workbook scope *)

statement     = { doc | annotation } head "=" [ formula ] ";" ;
doc           = "/**" text "*/" ;
annotation    = "@hidden"
              | "@workbook"                      (* sheet file: this statement only *)
              | "@sheet" "(" sheet ")"           (* module file: this statement only *)
              | "@renamed" "(" [ [ sheet ] "!" ] name ")"
              | "@from" "(" "lib" "#" hex6 ")" ;

head          = name                             (* a name statement *)
              | name address                     (* a named cell statement or a slot *)
              | address ;                        (* an unnamed cell statement or block *)

address       = "@" [ sheetref "!" ] cell [ ":" cell ] [ "#" ] ;
cell          = [ "$" ] col [ "$" ] row ;        (* col: 1-3 letters; row: 1-7 digits *)
sheetref      = sheetname | "'" { char | "''" } "'" ;   (* quoted by Excel's rule, §7.4 *)
sheet         = sheetname | "'" { char | "''" } "'" ;

name          = Excel defined name, e.g. Sales, IS_Sales, FN.GROW ;
formula       = an Excel formula in display form (§8), without the leading "=" ;
hex6          = 6 hex digits ;
```

Constraints the grammar does not show:

- `formula` may be empty only in a cell statement (`Name @C6 = ;`, `@C5 = ;`).
- `#` follows a single cell, and only in a **named** cell statement.
- A named cell statement has one cell, never a range.
- An unnamed cell statement takes no doc comment and no annotation.
- In a sheet file of the current form the address is **bare** (`@C6`); the sheet is the
  file's. Elsewhere it names its sheet (`@BS!C6`), unless the name is local to a sheet
  (`@sheet(BS)` above it, or an old `@scope` block), whose sheet it is.
- `Name : type = …` (AFE's type declaration) is an error, *type declarations are not part
  of xln v1 (planned with the dimension layer)*, on the type. Types come with the
  dimension layer; no pull ever wrote one.
- A `doc` and the annotations may come in any order before the head; the pull writes doc
  comment, then `@workbook`, `@hidden`, `@from`. Each applies to the next statement only.
- A sheet file's `@sheet(Name)` and the old `@scope` and `@workbook` directives are lines
  of their own, not attached to a statement. In a module file `@sheet(Sheet)` is an
  annotation of the next statement.

## 5. Name statements and annotations

### 5.1 Name statements

```
Rate = 0.1;
Fn = LAMBDA(x, x*Rate);
Spl = 'S1'!$E$1#;
```

`Name = formula;` defines a name. Its scope comes from the file and its block (§7). The
formula is in display form (§8). A new name must be legal for Excel:

- 1–255 characters; it starts with a letter, `_` or `\`; it continues with letters,
  digits, marks, `_`, `\`, `.`, `?`;
- not `R` or `C`, not shaped like a cell (`OPT2`) or an R1C1 reference;
- not starting with `_xl`;
- for a LAMBDA, not spelled like a built-in function (`Fact` collides with `FACT`: a call
  reaches the built-in, probe T12). A value or a range may be (`Rate = 0.1;` is read as
  the name); calling one is the audit's C3.

A LAMBDA spelled like an Excel 4.0 macro function that is not a worksheet function
(`Group`, `Get.Cell`, `Evaluate`, `Files`) is a **warning**, `xlm-name`, not
an error: `Group is also an Excel 4.0 macro function: Excel may call that instead or
refuse the name (AFE #10, not measured)`. The list is the functions of [MS-XLS] Ftab
(`packages/core/src/lang/xlm-data.ts`). The command equivalents of Cetab (`Open`, `Save`,
`Save.As`, `Copy`, `Table`, `Scale`) are kept there as data but get no warning: they are
ordinary words and the warning was noise. A value gets no warning, as for built-ins. `xln rename` to such a name prints the same warning
and renames.

Names are case-insensitive, as in Excel. A name may be defined once per scope.

### 5.2 Annotations

| Annotation | Where | Meaning |
|---|---|---|
| `@hidden` | any file, before a name | The name is hidden in Excel's Name Manager |
| `@workbook` | sheet file, before a name | The name is workbook-scoped (§7.1). In a module file, where names are workbook-scoped anyway, a hint with *Remove @workbook* |
| `@sheet(Sheet)` | module file, before a name | The name is local to `Sheet` (§7.2). A sheet the workbook lacks is an error |
| `@renamed(Old)` | before the new name | This name was called `Old` (same scope); `@renamed(Sheet!Old)` names the old scope, `@renamed(!Old)` workbook scope. The build renames instead of deleting and creating, and in the same scope rewrites the name in the workbook's formulas (§9.6). `xln rename` and the editor's Rename Symbol write it; the build that makes the rename removes it, and so does any later build that writes the workbook; once the rename is built it is a hint, `renamed-built` |
| `@from(lib #353921)` | module file, before a name | The library base: the library version this copy came from (§12.3). Written by Insert, Take, Publish and Record library base |

They stack, one per line, in any order:

```
/**
 * Grows a value by a rate, compounded over a number of periods.
 * @param value the starting value
 * @param rate the growth rate per period (0.05 for 5%)
 * @param [periods] how many periods; 1 when omitted
 */
@hidden
@from(lib #b653b0)
ANA.GROW = LAMBDA(value, rate, [periods],
    value * (1 + rate) ^ IF(ISOMITTED(periods), 1, periods)
);
```

An annotation xln does not know is a warning and is ignored; the message lists those of
the file kind (`@deprecated is not an annotation xln knows (@hidden, @sheet(Sheet),
@renamed(Old), @from(lib #…)); it is ignored` in a module file; `@hidden, @workbook,
@renamed(Old), @from(lib #…); and on a line of its own @sheet(Name)` in a sheet file). A doc comment or annotation with no statement after it is an error.
`@from` outside a module file is a warning: no provenance tag carries it, so a pull
drops it. `@from` given twice, or malformed, is an error.

### 5.3 File header directives

- `// @version 1.2`: in a module file, a line comment among the comment lines that open
  the file (before the first statement or doc comment). The build puts the version in the
  provenance tag (§11.2); a pull writes it back from the tags.
- `@sheet(Name)`: in a sheet file whose name cannot say its sheet (`IS~2.xln`), on a line
  of its own.
- `@scope(Sheet)` / `@workbook`: the block directives of files written before 2026-10-07
  (§7.3). `@scope` in `_unmanaged.xln` is a warning: a pull puts those names in their
  sheet's file.

## 6. Cell statements

Cell statements live in sheet files. Every formula cell of the sheet appears, in sheet
order (row by row, left to right), and every name on an empty cell. After them come the
sheet's other names that no module owns, under the line `// Other names on <Sheet>.`:
first its local names (constants, ranges, LAMBDAs), then the workbook names fixed to its
cells (§2.1), each with `@workbook` above it:

```
// Other names on Mortgage.

Mortgage_list = Mortgage_terms[Mortgage];
@workbook
amount = Mortgage!$B$12;
@workbook
Mortgage = Mortgage!$B$12:$G$12;
```

From the author's `is-model` (pulled, unchanged):

```
// Sheet IS, pulled by xln from is-model.xlsx: its formula cells in sheet order, then
// …

@workbook
years @C2# =  SEQUENCE(1, 5, 2025);
Sales @C3# = SEQUENCE(1,5,20000,3400);
COGS @C4# = Sales * 0.6;
Gross_income @C5# = Sales - COGS;
…
Unlevered_net_income @C11# = EBIT - Taxes;
```

### 6.1 Forms

| Form | Meaning |
|---|---|
| `Name @C6 = formula;` | A named cell: the name is defined as the cell (`S!$C$6`) and the cell holds the formula |
| `Name @C6# = formula;` | The name is defined as the cell's spill (`S!$C$6#`, stored `_xlfn.ANCHORARRAY(S!$C$6)`) |
| `Name @C6 = ;` | A **slot**: a name on an empty cell (no formula, no value), waiting for a formula |
| `@C5 = formula;` | An unnamed formula cell |
| `@B40:G40 = formula;` | A **block**: cells holding one formula filled across. The formula is the top-left cell's; the others are its fill |
| `@C5 = ;` / `Name @C6 = ;` on a formula cell | Clears the cell's formula |

A pull makes a block of cells that form a shared-formula group or are equal after
shifting relative references: row runs first, then runs over the same columns on
consecutive rows into rectangles. A spill anchor is one cell; the cells it spills into
are not statements, nor are data tables.

When several names sit on one cell, one is the statement (the sheet-scoped one, then the
`#` one, then the first by name); the others stay ordinary name statements.

### 6.2 What is read-only

- **The address is set in Excel.** Changing it is an error with the quick fix *Restore the
  address* (`Far: the address is set in Excel and read-only; the last pull had @B9`). To
  move a cell, move it in Excel and pull.
- **A new cell statement cannot be added in the source.** An address the last pull did not
  have is an error (`no cell statement at Ratios!C20 in the last pull: addresses are set
  in Excel (write the formula in Excel, then pull)`). A name on a new cell is made in
  Excel (Create from Selection, the Name Box, the Name Manager) and pulled.
- **Removing a named statement** from the source refuses the build: clear the cell with
  `Name @C6 = ;`, or delete the name in Excel and pull.
- **Removing an unnamed statement** writes nothing: the cell is left as it is, and the
  build reports it.
- Editable: the formula after `=`, the doc comment, `@hidden`, `@workbook` (a scope
  change), the name with `@renamed` (§9.6: `xln rename` or Rename Symbol), and the `#`.

### 6.3 The `#`

The `#` says what the name covers. It is the one editable part of the address.

- `Sales @C3# = SEQUENCE(1, 5, 20000, 3400);`: `Sales` is the whole spill, C3:G3.
  `COGS @C4# = Sales * 0.6;` then computes five years.
- `Sales @C3 = SEQUENCE(1, 5, 20000, 3400);`: `Sales` is C3 alone, so `Sales * 0.6` is
  one number (the first year).
- A pull writes `#` exactly when the workbook's name has it.
- Adding or removing it changes the name's definition (a `set-name`, with the three-way
  check of any name).
- A named statement without `#` whose formula spilled beyond the cell when the workbook
  was last saved gets a warning (from the shared checker: the editor, `xln check` and the
  build list it alike) with the quick fix *Name the whole spill: @C3#*:
  `Out @B1: the formula spills over B1:B3, but Out covers only B1 (write @B1# to name the
  spill)`. A 1×1 result gets none. Nothing moves the name on its own.
- `Name @C6# = ;` (a `#` on an empty slot the name is not on yet) is a warning, *'#' on a
  cell left empty is not written yet*, and the `#` is written only when the cell gets a
  formula.
- `#` on an unnamed statement or a range is an error.

Projects pulled before 2026-10-05 (lockfile format 3 or older) read a missing `#` on a
name that is on the spill as `#`; the next pull writes it in.

### 6.4 Addresses outside sheet files

In a module file (and in a sheet file written before M3d) an address names its sheet:
`Created @'Cash Flow'!$D$9 = 1;`. The sheet is quoted by Excel's rule (§7.4); either form
is read. For a name local to a sheet (`@sheet(Sheet)` above it, or an old `@scope` block)
the sheet may be left out. A pull never writes cell statements into module files.

## 7. Scope

A name is **workbook-scoped** (visible everywhere) or **local** to one sheet (Excel's
`localSheetId`, the 0-based position of the sheet in the workbook).

### 7.1 Sheet files

Every name in `names/sheets/<Sheet>.xln` is local to that sheet, unless `@workbook` is on
a line above it (with its doc comment and other annotations):

```
@workbook
years @C2# = SEQUENCE(1, 5, 2025);     // workbook-scoped
Sales @C3# = SEQUENCE(1,5,20000,3400); // local to IS
```

*Formulas → Create from Selection* makes workbook names, so a pulled slot made that way
has `@workbook` above it, and so has a name on cells inside a spill or on a range of the
sheet (`amount = Mortgage!$B$12;` under *Other names*). Adding or removing `@workbook` is a
scope change the build applies (`rescope-name`), shown as info while editing
(`Tax: the build moves it from workbook scope to sheet Slot (a scope change,
rescope-name)`). A workbook name on a sheet's cells that only that sheet reads gets a hint
with the quick fix *Remove @workbook*; when another sheet or a name of another scope
reads it, it needs workbook scope and gets nothing. Once local, a pull writes the
reference without its own sheet (`amount = $B$12;`), as for every local name; the quick
fix writes it that way too, so the pull after the build changes nothing.

Scope can also change in Excel: the Name Manager cannot change a scope, so delete the
name and create it again in the other scope; the next pull shows it. Edit one side at a
time.

### 7.2 Module files and `_unmanaged.xln`

Scope is per name, as in sheet files (decided 2026-10-07). A name in a module file is
workbook-scoped unless `@sheet(Sheet)` is on a line above it (with its doc comment and
other annotations): then it is local to `Sheet`. The annotation applies to the next
statement only.

```
// module: IN, …
IN_Rate = 0.03;                  // workbook scope

@sheet(Model)
IN_Horizon = 5;                  // local to Model

IN_Tax = 0.24;                   // workbook scope
```

The sheet in `@sheet(…)` is quoted by Excel's rule (§7.4): `@sheet('SCF recursive')`,
`@sheet('S2')`. Adding or removing `@sheet(…)` is a scope change the build applies
(`rescope-name`). `@workbook` in a module file says nothing new: a hint with *Remove
@workbook*. `_unmanaged.xln` holds workbook-scoped names: a name there with `@sheet(…)`
is a `file-placement` warning (a pull puts it in its sheet's file).

### 7.3 Files written before 2026-10-07: blocks

Older files have **blocks**: `@scope(Sheet)` makes every name after it local to `Sheet`
until the next `@scope(…)` or `@workbook`, a **directive** that returns to workbook scope.
They are still read, and build unchanged:

- a **module file** with `@scope` blocks: the editor shows an info with *Convert to
  per-name @sheet* (each name of a block gets `@sheet(Sheet)` above it, an unnamed cell
  statement names its sheet, the block lines go);
- a **sheet file** with a `@scope(…)` line (before 2026-10-06): *Convert to per-name
  @workbook*.

A build reads either as it is; the next pull writes the new form.

### 7.4 How names resolve in formulas

As in Excel:

- In a definition or cell of sheet S, a bare `X` is S's local `X` if there is one, else
  the workbook's `X`.
- Another sheet's local name must be qualified: `IS!Sales`. Read bare it is `#NAME?`;
  the checker says `Sales is local to IS, not to Ratios: unqualified it is #NAME? in
  Excel; write IS!Sales` and offers *Qualify: IS!Sales*.
- A workbook-scoped name is read bare from anywhere.
- LET and LAMBDA variables shadow names inside their body.
- A call spelled like a built-in calls the built-in (`Fact(5)` is `FACT(5)`).
- **Sheet quoting is Excel's rule, everywhere** (formulas stored and shown, `@sheet(…)`,
  `@scope(…)`, addresses): a sheet is bare when it is a plain word (letters, digits, `_`,
  `.`, starting with a letter or `_`) that does not read as a cell or an R1C1 reference,
  and quoted otherwise: `IS!X`, `BS!X`, `Model!X`, but `'S1'!X`, `'SCF recursive'!X`,
  `'0 Tables'!X`. Measured on the Excel-saved files (probes/README.md, *Sheet quoting*):
  Excel writes column-like names such as `BS` bare. A pull rewrites qualifiers to this
  rule; comparisons ignore optional quotes, so that is never an edit.

## 8. Formulas

### 8.1 Display form

Formulas in `.xln` files are written as Excel shows them in the formula bar, in English,
with `,` as the argument separator and `.` as the decimal point, whatever the locale.
Everything Excel accepts is accepted: references (`A1`, `$A$1`, `A:A`, `1:3`, 3-D
`S1:S3!A1`, other workbooks `[Other.xlsx]Sheet1!A1` and `Other.xlsx!Rate`, trim ranges
`A1:.B9`, `A1.:B9`, `A1.:.B9`, `A:.A`),
spills (`C6#`), implicit intersection (`@x`), intersection (space) and union (`(a, b)`),
structured references (`Tbl[Col]`, `Tbl[[#Headers],[A]:[B]]`, `[@Col]`), arrays
(`{1,2;3,4}`), errors (`#N/A`), `%`, strings with `""`.

A cell holding a LAMBDA can be called through its reference, `C2(D2, E2)`, or through a
name on it. A word directly followed by `(` is a called cell when it looks like a cell
and has a `$` or a sheet, or when it is bare and not a catalogue function (`LOG10(` is the
function). `C2 (D2)`, with a space, is an intersection.

### 8.2 Stored form, and what the compiler adds

The workbook stores formulas differently. The build compiles; the pull decompiles.

| Display (source) | Stored (file) | Rule |
|---|---|---|
| `SEQUENCE(1, 5)` | `_xlfn.SEQUENCE(1, 5)` | A function newer than Excel 2007 carries its catalogue prefix |
| `FILTER(…)`, `SORT(…)` | `_xlfn._xlws.FILTER(…)` | Two functions carry `_xlfn._xlws.` |
| `LAMBDA(x, x + 1)` | `_xlfn.LAMBDA(_xlpm.x, _xlpm.x + 1)` | LAMBDA and LET parameters and their uses carry `_xlpm.` |
| `LAMBDA(a, [b], …)` | `_xlfn.LAMBDA(_xlpm.a,_xlop.b, …)` | An optional parameter is `_xlop.b` in the list, without brackets; its uses stay `_xlpm.b` (probe F9) |
| `MAP(A1:A3, ABS)` | `_xlfn.MAP(A1:A3, _xleta.ABS)` | A built-in passed as a function (BYROW, BYCOL, MAP, REDUCE, SCAN, MAKEARRAY, GROUPBY, PIVOTBY) carries `_xleta.` |
| `C6#` | `_xlfn.ANCHORARRAY(C6)` | Spill reference |
| `@x` | `_xlfn.SINGLE(x)` | Implicit intersection |
| `Tbl[@Col]` | `Tbl[[#This Row],[Col]]` | This-row structured reference |
| `A1.:.A10`, `A1:.A10`, `A1.:A10` | `_xlfn._TRO_ALL(A1:A10)`, `_xlfn._TRO_TRAILING(A1:A10)`, `_xlfn._TRO_LEADING(A1:A10)` | Trim references (leading and trailing, trailing, leading blanks), whole columns and rows too (`A:.A`); the qualifier goes inside (`_xlfn._TRO_ALL(Sheet1!$A$1:$A$10)`). Measured in names and cells (probe F10) |
| `[Other.xlsx]Sheet1!$A$1`, `Other.xlsx!Rate` | `[1]Sheet1!$A$1`, `[1]!Rate` | Another workbook is stored by the number of its link (`<externalReferences>`), shown by the linked file's name as Excel shows it while that file is open (while it is closed Excel shows the full path; the display was not part of the probe) (quoted when the file or the sheet needs it: `'[My Book.xlsx]Sheet 1'!A1`; a folder before the name is ignored). Only Excel writes a link: a workbook the file has no link to is an error, *type the reference once in Excel, save, then pull*. xln never writes the link parts; a build copies them byte for byte (probe F10) |
| `Sales` in a name local to IS | `IS!Sales` | A sheet-scoped definition's own sheet is qualified (Excel qualifies on entry), quoted by Excel's rule (`'S1'!Sales`); a pull un-qualifies it |
| line breaks | CR LF | Excel stores CR LF |

The pull does the reverse, so the source never shows a prefix. A prefix written by hand
is accepted when it is the catalogue's (`_xlfn.LET(_xlpm.x, …)`); an unknown one
(`_xlfn.FOO`) is an error. A stored form missing its prefix (`SEQUENCE(` bare in the
file) becomes `#NAME?` and Excel re-saves it as `_xludf.SEQUENCE` for good; the audit
reports both (C2). Every stored text a build writes passes a strict check of Excel's
stored grammar before writing.

Cell formulas are written in dynamic-array form (`cm` and `t="array"`), which Excel
accepts for scalars too (probe F8).

### 8.3 LET and LAMBDA

```
FN.SPREAD = LAMBDA(v, periods,
    EXPAND(v, 1, periods, INDEX(v, 1, COLUMNS(v)))
);
ANA.GROW = LAMBDA(value, rate, [periods],
    value * (1 + rate) ^ IF(ISOMITTED(periods), 1, periods)
);
INLET = LET(f, LAMBDA(t, t * 3), f(2));
MAKEADDER = LAMBDA(n, LAMBDA(x, x + n));
IMMED = LAMBDA(x, x + 1)(2);
```

- `[p]` marks an optional parameter; test it with `ISOMITTED(p)`. A call may leave it
  out, or pass an empty argument (`ALLOPT(, 7)`).
- Argument counts are checked against the LAMBDA: `FN.PREV(row) takes 1 argument; it is
  given 2` (an error); built-ins against the catalogue (a warning).
- A LAMBDA bound in a LET, or returned by a LAMBDA, can be called (`f(2)`,
  `MAKEADDER(2)(3)`). LAMBDA names may call themselves (recursion is not a cycle).
- Long LET and LAMBDA definitions are pretty-printed by a pull (one binding per line),
  unless Excel already stored them on several lines.

### 8.4 The function catalogue

`packages/core/src/lang/catalogue-data.ts` lists 525 functions: name, stored prefix (none,
`_xlfn.`, `_xlfn._xlws.`), minimum and maximum arguments, first Excel version, flags
(`?prefix` for the 13 whose prefix is not confirmed; `internal` for `ANCHORARRAY` and
`SINGLE`). `catalogue-params.ts` gives the parameter names from Microsoft's reference for
signature help. A function not in the catalogue and not a defined name is an error: Excel
would store it as `_xludf.` (`XLOKUP(…): no built-in function and no LAMBDA of that name;
Excel would store it as _xludf.XLOKUP (#NAME?): did you mean XLOOKUP?`). The catalogue is
plain data: adding a function is a one-line change with a test.

## 9. What a build does with each construct

### 9.1 Three-way comparison

Per name and per cell statement, the build compares the source, the lockfile (what the
last pull or build saw) and the workbook as it is now:

| Source vs lockfile | Workbook vs lockfile | Result |
|---|---|---|
| changed | unchanged | written |
| unchanged | changed | Excel's version kept, reported (pull to see it) |
| changed | changed the same way | nothing to do |
| changed | changed differently | **conflict**: nothing written; shown as a diff |

A named cell follows its name: if Excel moved it (rows inserted), the build writes where
the name points now. An unnamed cell is its address. A source text equal to the
workbook's formula is no edit, even when compiling it would add a missing prefix.

### 9.2 Changes

| Source edit | Change |
|---|---|
| new name statement | create the name |
| edited formula, doc comment, `@hidden` | update the name |
| name statement removed | delete the name |
| `@renamed(Old)` above a new name | rename |
| `@workbook` added or removed (sheet file), `@sheet(…)` added, removed or changed (module file), name moved to another sheet file or `@scope` block | re-scope |
| `#` added or removed | update the name's definition |
| cell formula edited | set the cell's formula (a block: every cell, the top-left formula filled) |
| `= ;` on a formula cell | clear the cell's formula |
| slot filled | write the formula into the empty cell |
| `@from` changed | update the provenance tag only |

A deletion or re-scope (and a rename that also changes the scope) that would leave a
cell, conditional format, validation, Table column, chart or another name pointing at
nothing is refused with the places, counted as the workbook will be after this build (a
reader the same build rewrites does not count):

```
error names/sheets/S1.xln:19: deleting Fn refused: 1 place in the workbook refers to it
by name and would break: cell S1!C10 (names/sheets/S1.xln:19). Change that formula in
the source, or keep the name
```

A rename in the same scope rewrites its readers instead (§9.6).

### 9.3 What is written

Only `<definedNames>` in `xl/workbook.xml`, the sheets whose cells change, `metadata.xml`
when a dynamic-array record is needed, and `fullCalcOnLoad="1"`. `calcChain.xml` is
dropped after a cell change. Every other zip entry is copied byte for byte. The result is
read back and decompiled before and after writing; on a mismatch the original stays.

### 9.4 Safety

- Desktop: refused while Excel's owner file `~$book.xlsx` exists (Excel has it open); the
  editor offers *Close in Excel and build*. The previous file is kept as
  `book.backup.xlsx`.
- Browser (vscode.dev): Excel's owner file is invisible, so the build writes
  `book.xln.xlsx` and leaves `book.xlsx` and the lockfile alone.
- `xln verify` compares cached values before and after Excel recalculated and saved; it
  warns first when a side's values are not Excel's (never saved by Excel).

### 9.5 The pull guard

A pull would replace source edits not built yet, so it first lists them (the build's own
comparison) and refuses: CLI exit 1 with the list (`--discard` replaces them); the
editor asks *Build first* or *Discard and pull*. Changes of layout or `//` comments alone
are named but do not refuse. A rename is one line (`Pace: rename Rate → Pace, rewriting
it in 8 cell formulas, …`); the readers whose only edit is its token are not listed
apart.

An edit is unbuilt only where the source **differs from the workbook** as it is now; the
lockfile decides only which side moved. So after a browser build, once the author has put
`book.xln.xlsx` in place of `book.xlsx` (the lockfile still the one before the build),
build and pull agree: nothing to build, nothing lost.

### 9.6 Renaming

A rename is explicit: pull and build never rename on their own, and a name that vanishes
from the source while another appears is a deletion plus a creation. Two ways to say it,
the same result:

- **`xln rename <project> Old New`** (`Sheet!Old` for a sheet's local name), or **Rename
  Symbol** (F2) on the name in the editor, which shows the edits in a preview first.
  Either edits the source: the statement gets the new name with `@renamed(Old)` above
  it, and every formula of the project's `.xln` files that reads the name gets the new
  token. `@renamed` records where the name is in the workbook (the lockfile): renamed
  twice before a build, the first one stays; renamed back, it goes; a name not built yet
  gets none.
- **By hand**: `@renamed(Old)` above the new name, and the new name in every formula of
  the source that reads it. A formula left with the old name is an error, `in-use`:
  `RateX is not a name in scope …: the source renames RateX to RateY, so write RateY
  here (xln rename, or Rename Symbol in the editor, renames every reader); the build
  refuses until then`.

```
// names/_unmanaged.xln after `xln rename book.xln Rate Pace`
Fn = LAMBDA(x, x*Pace);
@renamed(Rate)
Pace = 0.1;
Rate2 = Pace*2;

// names/sheets/S1.xln
@C2 = "Rate is "&Pace;          // the string is not a name
@C3 = LET(Rate, 5, Rate*2);     // a LET variable, not the name
@C7 = INDIRECT("Rate");         // text: Excel does not rename it either
```

The build then renames the name (`rename-name`) and, in the same scope, **rewrites its
token in the workbook's formulas** as Excel's Name Manager does (probe F7): cell formulas
(a shared group's text at its master), conditional formats, validations and other names'
definitions. The token is resolved as Excel resolves it: a sheet's local name shadows the
workbook's on that sheet, `Sheet!Name` and `'Sheet 1'!Name` name the scope, `[0]!Name` is
this workbook's, `Name#` (stored `ANCHORARRAY(Name)`) is the name, case does not matter;
LET/LAMBDA variables of that spelling (`_xlpm.Rate`), strings, structured references
(`Tbl[Rate]`) and other workbooks' names (`[1]!Rate`) are left alone. Nothing else of
those formulas changes: no new cell, no dynamic-array conversion, no other text; a cell
statement whose only edit is the renamed token is not written as a new formula (one
edited further is, as any edited cell). After a cell formula changed, `calcChain.xml`
goes and `fullCalcOnLoad="1"` is set, as for any cell change. The report says how many
formulas: `rename Rate → Pace, rewriting it in 8 cell formulas, 1 conditional format, 1
validation, 2 names`. Read-back checks that every formula of the built file is the old
one with the token replaced, and nothing else changed.

**Refused**, with the places, nothing written:

| Where the name is read | Why |
|---|---|
| a chart (`xl/charts/*`), a Table column's formula, a pivot table's source, a hyperlink's location, a form control's link (`fmlaLink`, `fmlaRange`), a shape's text link, a link of the workbook to itself | the build does not rewrite those parts: `… the build rewrites the name in cell formulas, conditional formats, validations and other names, but cannot reach these places: rename it in Excel's Name Manager instead (it rewrites them too), then pull; or keep the name` |
| a formula that does not parse and spells the old name | it cannot be rewritten safely |
| a formula the new name would make read something else: a LET/LAMBDA variable of that spelling, a sheet's local name that would shadow it, an unknown name of the new spelling | `rename-capture`: `renaming Rate to g refused: in 1 place the new name would read something else: name Mul (would read g)` |

`xln rename` and Rename Symbol refuse the same captures in the source, and an invalid or
taken name (`RateX exists already`, a Table's name, a built-in's for a LAMBDA: `Growth`
collides with GROWTH). With the workbook beside the project, `xln rename` plans the next build first and
refuses, writing nothing, what that build would refuse.

A rename that also changes the scope (`@renamed(Sheet!Old)`, `@renamed('Cash Flow'!Old)`
with Excel's quoting, or `@renamed(!Old)` above a local name) rewrites nothing in the workbook: its readers mean something else once the
scope changes, so they are the source's to change, and the refusal of §9.2 applies.

**The build removes `@renamed`.** `@renamed(Old)` is a note of a pending change: it is
how the build tells a rename from a deletion plus a creation, so it stays while the
rename is not built. Once a build has written the workbook with the rename and its
read-back passed, the note is spent: the build removes each `@renamed(…)` it applied from
the source (its line, or the annotation alone when the line holds more), which is what
the next pull would write. This is the one exception to "the build never changes the
source": nothing else in the source is touched, and only spent notes: those of the
renames this build made (a scope change by `@renamed(Sheet!Old)` included) and those of
renames already built (below). A refused build, a dry
run, and a build written to another file (`xln build --out`, the browser's
`<name>.xln.xlsx`) leave the source as it is. The output says it: `removed
@renamed(Rate) from names/_unmanaged.xln: the rename is built`. The source then agrees
with the lockfile and the workbook: the next pull finds nothing unbuilt and rewrites
nothing.

A note can outlive its rename: left by an older build, by a browser build whose copy
then replaced the workbook, by another tool, or written by hand for a rename made in
Excel. Its rename is built when the lockfile (the last pull or build) has the name in
its scope and not the old one. The build ignores such a note; the checker marks it with
a hint, `renamed-built` (`@renamed(Rate): the rename is built; this line can go`, faded
in the editor, with the quick fix *Remove @renamed(Rate)*); and the next build that
writes the workbook removes it with the notes of its own renames, with the same output
line. A build with nothing to write (up to date) leaves the source alone, and so do a
refused build, a dry run and a build written to another file. A pull drops it too. A
note whose old name is still in the lockfile (the rename is pending) and a change of
spelling only (`@renamed(rate)` above `Rate`) are not spent: they stay.

**Labels.** xln never writes cell values: a label cell that still reads the old name
(the text Create from Selection named the cells from) stays as it is. For a renamed name
on one sheet's cells, `xln build` (and `xln rename`, for after the build) prints a
**label notice**: the cells of that sheet, in the name's rows and columns, whose typed
text gives the old name by Create from Selection's conversion ("Gross income" and
"Gross_income" for `Gross_income`), and Excel's Find & Replace that fixes them: one *Find
what* / *Replace with* pair per text, the replacement in the label's style ("Gross ind
income" for `Gross_ind_income`), with *Within: Sheet* and *Match entire cell contents*
(a formula's content starts with `=` and never matches). Text that only resembles the
old name is listed to check by eye, never in a pair. A name not on cells (a LAMBDA, a
constant) gets none.

Not reached by any check: other workbooks that link to this one by name (Excel's own
rename does not fix those either unless they are open), names inside text (`INDIRECT`),
VBA, and legacy (VML) form controls without a `ctrlProps` part.

## 10. Checks

One checker serves the editor (as you type) and the build: the build refuses exactly the
errors the editor shows, plus what needs the workbook (conflicts, broken references, a
removed named statement, a comment over 255 characters with its tag). `xln check` on a
project lists the same source findings after the workbook audit.

| Severity | Effect | Examples |
|---|---|---|
| **error** | The build refuses | syntax (a type declaration among them); an address changed or new; `#` on an unnamed statement; a name defined twice; a name Excel would refuse; unknown name, function or sheet (`@sheet(…)` included); another sheet's local name read bare; LAMBDA argument count; doc comment over 255; a file with no place in `names/`; a reader the source's rename, re-scope or deletion breaks (`in-use`; for a rename: a source formula still reading the old name, or a workbook reader the build cannot rewrite, §9.6); a rename the new name would capture somewhere (`rename-capture`) |
| **warning** | Built anyway | built-in argument count; a name without `#` on a spilling formula (`spill-uncovered`); `#` on an empty slot (`spill-empty`); an unknown annotation; `@from` outside a module file; a doc comment in a module file too long for the provenance tag the build adds (`provenance`: `the doc comment is 235 characters; with its provenance tag (29) it passes Excel's 255, so the build writes it without the tag and the workbook won't record its library base, @from(lib #353921): shorten it by 9 characters`); a doc comment's `@param` naming no parameter (`doc-param`, with *Rename @param x to y* and *Remove @param x*); two doc comments before one name (`doc-twice`); a name in a file a pull would not write it to (`file-placement`: `Rate has no ANA. prefix: a pull will put it in _unmanaged.xln; rename it ANA.Rate or move it there`; `@workbook` on a name that is not a cell in a sheet file), with *Move it to …* and, in a module file, *Rename it ANA.Rate*; `@scope` in `_unmanaged.xln` (`directive`); a new LAMBDA named like an Excel 4.0 macro function (`xlm-name`, §5.1) |
| **info** | Information | a scope change the build will make; a file with old `@scope` blocks (`old-blocks`, with the conversion) |
| **hint** | Dots under the text, not in the Problems panel | a name in another case than its definition (`EbIT` for `EBIT`); a workbook name on a cell only its sheet reads; parameters left out of the `@param` lines while others are documented (`doc-param-missing`); `@workbook` in a module file (`redundant-workbook`); an `@renamed(Old)` whose rename is built (`renamed-built`, faded, with *Remove @renamed(Old)*) |

One checker serves every severity: `xln check` (with `--severity error|warning|info|hint`),
the build (its errors refuse; its warnings are listed) and the editor report the same
findings. A finding in a formula the lockfile already has (unchanged since the last pull)
is downgraded from error to warning with `(as in the workbook since the last pull: the
build leaves it; fix it here to have it written)`, and the build does not list it, unless
what the formula reads at the last pull is gone because of the source (renamed, moved to
another scope, deleted): then it stays an error, `in-use`, as the build refuses it.

The workbook audit (`xln check`, the *Audit workbook* report, Problems with source
`xln check`) runs checks C1–C15 on the workbook itself; its rules are listed in
`packages/core/README.md`.

C15 (`C15.label-drift`, info) reads the cells' values: a name whose label cell no longer
gives it. Excel's *Create from Selection* names a row after the text just left of it (a
column after the text just above it), by the rule probe F11 measured: the text trimmed,
each character a name cannot hold becoming one `_` (`a - b` → `a___b`), those at the end
dropped (`Margin %` → `Margin`), `_` in front of a first character that cannot start a
name (`_2024_sales`, `_€uro`), `_` behind a text that reads as an A1 reference or as
TRUE or FALSE (`Q1_`, `R_`, `True_`) and in front of an R1C1 one (`_R1C1`); a number, or a
text over 255 characters, gives no name; a date gives the text it shows. After
`Gross_income` is renamed `Gross_ind_income` (F2, `xln rename`), the build renames the
name, but the label cell still reads "Gross income": xln never writes cell values, so it
says so (`IS!B5 reads "Gross income", which Create from Selection makes Gross_income; the
name on IS!C5:G5 is Gross_ind_income, probably renamed from Gross_income`). It fires only
where Create from Selection was evidently used: other names in the same column (or row)
of labels, over the same span, match their labels, at least as many as those that don't.
Names given by hand, a label that is a formula's result (the pull's note), an alias on
the same cells, a corner name and a header row of text give nothing. Its hint gives the
fix as Excel's Find & Replace: `in Excel, Find & Replace "Gross income" with "Gross ind
income" (Ctrl+H, ⌘⇧H on Mac; Match entire cell contents)`, the replacement in the
label's style.

The build gives the same advice for every rename it makes, before any audit: the
**label notice** (§9.6).

## 11. Comments, provenance and limits

### 11.1 Doc comments and `@param`

The doc comment is the Name Manager comment, whole. For a LAMBDA, `@param name text`
lines document the parameters (`@param [p]` or `@param p` for an optional one); the editor
shows them in hover and signature help. A line break counts as two characters (Excel
stores CR LF). Excel saves a line break in a comment as `_x000a_`; xln reads both.

### 11.2 The provenance tag

A build adds a tag to the comment of each name declared in a module file
(`names/FN.xln`; not `_unmanaged.xln`, not sheet files):

```
Grows a value by a rate, … [xln ANA 1.2 #40623a lib#b653b0]
                               │   │    │       └ library base, from @from(lib #b653b0)
                               │   │    └ hash of the stored definition and comment
                               │   └ version, from // @version 1.2 (optional)
                               └ module
```

A pull strips the tag; its `lib#` comes back as `@from(lib #…)`, and its version as the
module file's `// @version` when all the module's tagged names, or two in three, carry
the same one (a disagreement is a pull note). So a build right after a pull plans no
provenance update. A tag never counts as an
edit; a stale or missing tag is a provenance update. `--no-tags` turns tags off. The
pull's summary counts tagged names edited in Excel since the build (the hash no longer
matches).

### 11.3 Limits

| Limit | Value | Effect |
|---|---|---|
| Comment | 255 characters (line break = 2) | over it: error, Excel refuses the file (probe F5). If the tag would pass 255, the name is written without a tag: the checker's warning `provenance` on the doc comment (as you type, in `xln check`, and from Insert, Take, Publish and Record library base before they write), measured with the tag the build would write (module, version, hash, and `lib#` with `@from`) |
| Library doc comment | 212 characters | room for a tag of 32–43 characters; longer summaries are cut with `…` |
| Definition / cell formula | under 8,192 characters | error from 8,192, warning from 7,500 (C7) |
| Function nesting | 64 levels | error over 64, warning over 48 (C7) |
| Name | 255 characters | error |

## 12. The library

### 12.1 Files

A library is a folder of `.lambda` files, one LAMBDA each, as the author's Python build
reads them. The folder is `library` in `xln.config.json` or `--lib`.

```
# name       ANA.GROW
# summary    Grows a value by a rate, compounded over a number of periods.
# params     value, rate, [periods]
#
# @param value the starting value
# @param rate the growth rate per period (0.05 for 5%)
# @param [periods] how many periods; 1 when omitted

LAMBDA(value, rate, [periods],
    value * (1 + rate) ^ IF(ISOMITTED(periods), 1, periods)
)
```

- **Header:** the `#` lines up to the first blank `#` line. A field is `# key value` (one
  blank after `#`); a continuation line is indented further. Known fields: `name`
  (required), `summary`, `params` (must equal the LAMBDA's parameters), `example`,
  `impromptu`; others are kept.
- **Rationale:** the `#` lines after the header, free text. A line `@param x text`
  describes a parameter (xln's addition; the Python build ignores it).
- **Definition:** from the first line that is not a comment, in display form, one LAMBDA
  that must compile. The file name should be `<name>.lambda` (a warning otherwise).

xln writes the summary and the `@param` lines as the doc comment of the copy (at most 212
characters). The rationale stays in the library.

### 12.2 Versions

A library version is `libraryHash`: the first 6 hex digits of SHA-256 over the stored
definition, layout and number spelling aside. Comments do not count.

### 12.3 The library base, `@from`

`@from(lib #353921)` above a module entry records the library version the copy came
from. Only explicit actions write it: Insert, Take, Publish, Record library base. The
author may delete it (the copy then has no base). The build carries it in the tag as
`lib#353921`; a pull writes it back.

### 12.4 Statuses

Three-way on the copy, its base and the library:

| Status | Meaning | Actions offered |
|---|---|---|
| identical | the copy equals the library (whatever the base) | Record library base, when it has none |
| outdated | the copy is its base; the library moved | Take |
| modified | the copy moved; the library is still the base | Publish; Take (asks first: discards the edit) |
| both changed | all three differ | Show diff (base → copy, base → library); Take asks first; no one-click Publish |
| differs | no base recorded: which side moved cannot be told | Publish; Take (asks first) |
| missing | in the library, not in the project | Insert |
| local only | a LAMBDA of a library module (`FN.*`) the library lacks | Publish |

On a project, each function the workbook does not have as the source does is marked
*not built yet*. LAMBDAs of other modules are counted on one line, not listed.

The base's text, for the three-way diff, comes from `library-bases/<hash>.json`, else
from the workbook or its backup when they still hold it, else the diff shows copy against
library with a note.

## 13. Versioning, and differences from AFE

### 13.1 Versions of this specification

This is **version 0.2**. While the version is 0.x the language may change between
releases; each change is recorded here with its date. The lockfile carries its own format
(`xln.lock/4`), and the embedded part its own (`xln.embed/1`). A file has no version line:
old forms are recognised by their content (files with `@scope` blocks, lockfiles of
formats 1–3).

| Date | Change |
|---|---|
| 2026-10-04 | Module files, `@scope` blocks, `@hidden`, doc comments |
| 2026-10-05 | Cell statements, slots, blocks; explicit `#`; `@renamed`; provenance tags and `// @version` |
| 2026-10-06 | Sheet files: per-name `@workbook`, no blocks; `@sheet`; optional parameters stored as `_xlop.`; every pull fresh |
| 2026-10-07 | `@from(lib #…)`; `library-bases/`; `@param` checks (version 0.1) |
| 2026-10-07 | Workbook names fixed to one sheet's cells (a cell inside a spill, a range) are written in that sheet's file with `@workbook`, under *Other names on <Sheet>*; `_unmanaged.xln` keeps the rest |
| 2026-10-07 | Version 0.2: module files per-name `@sheet(Sheet)` (old `@scope` blocks read, with a conversion); type declarations an error; Excel's sheet-quoting rule everywhere; pull writes `// @version` back; checker warnings for names a pull would move, two doc comments, `@scope` in `_unmanaged.xln`, the spill warnings; broken readers are errors |
| 2026-10-07 | M5, no syntax change: a rename in the same scope rewrites the name's token in the workbook's formulas (§9.6); `xln rename` and Rename Symbol write `@renamed` and the readers; `rename-capture` |
| 2026-10-07 | No syntax change: a build that writes a rename removes its `@renamed(…)` from the source (§9.6), the one source edit a build makes |
| 2026-10-07 | `@renamed('Cash Flow'!Old)`: a quoted sheet in the argument parses; the label notice after a rename (§9.6), C15's hint gives the Find & Replace |
| 2026-10-07 | No syntax change: an `@renamed` whose rename is built is a hint, `renamed-built`, and a writing build removes it with its own (§9.6) |

### 13.2 Differences from AFE (Excel Labs) module syntax

xln reads the core of the Advanced Formula Environment's module syntax (`name = formula;`,
`//`, `/** */` doc comments synced to the Name Manager comment). The differences:

| Topic | AFE | xln |
|---|---|---|
| Where the source lives | inside the workbook, written by an add-in | text files next to the workbook; no add-in |
| Module names | a module `M` is a file of bare names, exported as `M.name` | names are written in full (`FN.GROW = …`); the file is chosen by the prefix. A bare name in a module file is a plain workbook name, and a warning (a pull moves it to `_unmanaged.xln`) |
| Sheet scope | not in modules (the Names tab) | per name: `@sheet(Sheet)` in module files, `@workbook` in sheet files |
| Type declarations | `name : type = …` | an error in v1 (planned with the dimension layer) |
| Cells | the Grid tab, one cell at a time | cell statements for every formula cell, in sheet order |
| Annotations | none | `@hidden`, `@workbook`, `@sheet`, `@renamed`, `@from` (and the old `@scope`) |
| `@param` | not defined | read by the editor and checked against the LAMBDA |
| Line ends, tabs | no `\t`, no `\r` | both accepted |
| Locale | translated on sync | always English separators in source |
| Names tab (two-way) / modules (one-way) | two sync modes | one: pull is fresh, build is three-way |
| Provenance | none | `[xln M 1.2 #hash lib#base]` tag |

## 14. Open issues found while writing

Found by reading the code and running the CLI on copies of fixtures (2026-10-07). Each
needed a decision or a fix. All but one are fixed (2026-10-07, branch `spec-fixes`); what
each was is kept below with what changed.

### Open

17. ~~Installing in vscode.dev needs a served extension.~~ Closed 2026-10-08: the
    extension is on the VS Code Marketplace (`lerzegov-xln.xln`), and vscode.dev installs
    it from the Extensions view.

### Fixed 2026-10-07

1. **`// @version` did not survive a pull.** Pull now writes `// @version x.y` into a
   module file's header when all its tagged names, or two in three, carry that version;
   a disagreement is a pull note. A build right after a pull plans no provenance update
   (§11.2).
2. **A name without its module's prefix in a module file** built as a module name and was
   moved by the next pull. Now a warning at the name, `Rate has no ANA. prefix: a pull
   will put it in _unmanaged.xln; rename it ANA.Rate or move it there`, with *Move it to
   names/_unmanaged.xln* and *Rename it ANA.Rate* (`file-placement`, §10). A name of another
   module pasted in is sent to its module's file.
3. **`: type` declarations** were parsed and silently ignored. Decided: not part of v1.
   `Name : type = …` is an error on the type, *type declarations are not part of xln v1
   (planned with the dimension layer)*; the grammar marks it invalid; `samples/demo.xln`
   has none. No pulled project ever had one.
4. **`@sheet(Name)` in a module file was silently ignored.** Since issue 5 it is the
   per-name scope annotation there. `@scope` in `_unmanaged.xln` is a warning naming it;
   the unknown-annotation message lists the annotations of the file kind.
5. **`@workbook` meant two things** (an annotation in sheet files, a block directive in
   module files). Decided: scope is per name everywhere. Module files use `@sheet(Sheet)`
   above a local name; pull writes no blocks; old `@scope`/`@workbook` blocks are still
   read (old projects build unchanged), with *Convert to per-name @sheet*; `@workbook` in
   a module file is a hint with *Remove @workbook* (§7.2, §7.3).
6. **`@workbook` on a non-cell name in a sheet file** built a workbook name that the next
   pull moved. Now a warning, *a pull puts workbook names that are not on one sheet's
   cells in _unmanaged.xln (or their module)*, with *Move it to names/_unmanaged.xln*.
   Since 2026-10-07 a workbook name fixed to the file's sheet's cells belongs there (§2.1);
   one in `_unmanaged.xln` or another sheet's file gets *`amount` is on cells of Mortgage: a
   pull puts it in names/sheets/Mortgage.xln (with @workbook above it)*, and the move fix
   carries `@workbook` along.
7. **Three quoting rules for sheet names.** Now one, Excel's, measured on the Excel-saved
   files: column-like names (`BS`, `IS`) bare, cell-like and non-word names quoted (§7.4).
   The compiler writes `IS!$C$3` (not `'IS'!$C$3`), the decompiler shows other sheets as
   Excel writes them, and `@sheet(…)`, `@scope(…)` and addresses follow the same rule.
8. **Two warnings lived outside the shared checker.** *The formula spills over …, but X
   covers only …* and *'#' on a cell left empty is not written yet* are the checker's:
   `xln check`, the build and the editor list them alike; the extension's copy is gone.
9. **After a browser build, pull and build disagreed.** An edit is unbuilt only where the
   source differs from the workbook (§9.5).
10. **Misleading downgrade after a rename or re-scope.** A reader whose name the source
    renamed, moved or deleted is an error (`in-use`) in the editor too, as the build
    refuses it; only formulas already broken at the last pull are downgraded (§10).
11. **Refusal wording** now reads in order: `renaming IS!COGS to Cost_of_sales refused`,
    `moving years to sheet IS refused`, `deleting Fn refused`.
12. **Stale CLI help:** `xln --help` lists every command (`lib base` among them) and the
    current flags; `build` no longer claims to move names to the spill.
13. **`xln check --severity hint`** is accepted (it lists everything).
14. **Two doc comments before one name** are a warning, *two doc comments: only the last
    is kept*, with *Remove this doc comment*. Doc comments and annotations in any order
    stay accepted.
15. **`samples/demo.xln`** is up to date: no "kept by pull" claim, no type declarations,
    per-name `@sheet`.
16. **The brief's A4** describes sheet files as they are (cell statements, per-name
    `@workbook`) and module files per-name `@sheet`.
