# xln user guide

*For modellers who know Excel well and have never needed a code editor. Version of
2026-10-07. The precise rules are in [`LANGUAGE.md`](LANGUAGE.md).*

Contents:
1. [What xln is, and what it is not](#1-what-xln-is-and-what-it-is-not)
2. [Install](#2-install)
3. [The loop, on a small income statement](#3-the-loop-on-a-small-income-statement)
4. [Spills and the `#`](#4-spills-and-the-)
5. [A second sheet: scope and `@workbook`](#5-a-second-sheet-scope-and-workbook)
6. [Modules and LAMBDAs](#6-modules-and-lambdas)
7. [How the editor helps](#7-how-the-editor-helps)
8. [The library](#8-the-library)
9. [Safety: what protects your workbook](#9-safety-what-protects-your-workbook)
10. [Common messages and what to do](#10-common-messages-and-what-to-do)
11. [Working with AI completion (Copilot)](#11-working-with-ai-completion-copilot)
12. [Working in the browser (vscode.dev)](#12-working-in-the-browser-vscodedev)
13. [Workbooks edited with AFE](#13-workbooks-edited-with-afe)
14. [FAQ](#14-faq)
15. [Cheat sheet](#15-cheat-sheet)

---

## 1. What xln is, and what it is not

Excel's Name Manager is where a modern workbook keeps its structure: the names of your
line items, your LAMBDA functions, your constants. But it is a small dialog with a
one-line box. xln shows all of that as text in VS Code, a free code editor, where you can
read it, search it, check it and edit it, and then writes your edits back into the
workbook.

**Excel and xln share the work:**

| Excel owns | xln owns |
|---|---|
| sheets, rows, columns | defined names (create, edit, delete, rename, scope, comment) |
| labels and other text | the formulas of existing cells |
| values, formats, charts | LAMBDA functions and their documentation |
| which cells exist and where | checks and the audit |

So you never add a row or type a label in xln. You do that in Excel, save, and let xln
read it.

**What you need:** VS Code on the desktop (macOS or Windows), or vscode.dev in Edge or
Chrome. **No add-in, no macros**, and Excel does not need to be running while you edit:
xln reads and writes the `.xlsx` file itself.

**What it is not:** it does not generate workbooks, it does not move cells, and it does
not keep a live link with Excel. You switch between the two: edit in Excel, or edit in
xln, one at a time.

## 2. Install

### Desktop VS Code (recommended for writing)

1. Install VS Code from code.visualstudio.com.
2. In VS Code: *Extensions* view (the four squares on the left) → search **xln** →
   install **xln — Excel names as code** (publisher Luca Erzegovesi,
   [Marketplace page](https://marketplace.visualstudio.com/items?itemName=lerzegov-xln.xln)).
   Updates arrive by themselves.
3. After an update, reload VS Code (*Developer: Reload Window*) if a window still runs the
   old version.
4. *File → Open Folder…* and open the folder that holds your workbook.

A `.vsix` file (from your teacher, or `npm run package -w xln` in the repository)
installs the same way: the `…` menu at the top of the Extensions view → **Install from
VSIX…**. Have only one xln installed: two copies show every menu twice.

### vscode.dev (when you cannot install anything)

On a lab PC that is wiped every few weeks, use vscode.dev in **Edge or Chrome** (not
Safari or Firefox: they cannot open a local folder).

1. Go to `https://vscode.dev`.
2. Install the extension: *Extensions* view → search **xln** → install **xln — Excel
   names as code**. vscode.dev remembers it in this browser profile; on a wiped PC,
   install it again (one click).
3. *Open Folder* and pick the folder with your workbook. Chrome asks to let the site view
   and then edit files: allow both.

What changes in the browser is in [§12](#12-working-in-the-browser-vscodedev).

### The command line (optional)

`xln` also runs in a terminal, with the same engine: `xln pull`, `xln build`,
`xln check`, `xln lib status`. You do not need it for anything in this guide. From a
checkout of the repository: `npm install`, `npm run build`, then
`node packages/cli/bin/xln.js --help`.

## 3. The loop, on a small income statement

We build a one-sheet income statement, `is-model.xlsx`: years across columns C to G,
line items down column B.

### Step 1: labels and slots, in Excel

On a sheet called `IS`, type the labels in column B: `Sales` in B3, `COGS` in B4,
`Gross income` in B5, `SGA expenses`, `EBITDA`, `Depreciation`, `EBIT`, `Taxes`,
`Unlevered net income` down to B11. Leave column C empty.

Now name the cells next to the labels. Select **B3:C11** (labels and the first-year
column only), then *Formulas → Create from Selection → Left column → OK*. Excel makes one
name per row, from the label: `Sales` on C3, `COGS` on C4, `Gross_income` on C5 (Excel
replaces the space). These names on empty cells are **slots**: places waiting for a
formula.

Select only one column of cells, C. If you select B3:G11, each name covers C:G as a
fixed range, which is not what you want here (see [§4](#4-spills-and-the-)).

Save the workbook. You may leave it open in Excel.

### Step 2: Pull

In VS Code, right-click `is-model.xlsx` in the Explorer → **Pull workbook**. xln writes
a folder `is-model.xln` next to the workbook. Open `is-model.xln/names/sheets/IS.xln`:

```
// Sheet IS, pulled by xln from is-model.xlsx: …

@workbook
Sales @C3 = ;
@workbook
COGS @C4 = ;
@workbook
Gross_income @C5 = ;
…
```

Read each line as "the name `Sales` sits on cell C3, which is empty". The `@workbook`
line says the name is visible from every sheet: Create from Selection makes names that
way. The address `@C3` belongs to Excel: you cannot change it here.

### Step 3: Write formulas

Type the formulas after the `=`, ending each with `;`. Years first: give row 2 a year header
too (in Excel, select C2, type `years` in the Name Box left of the formula bar and press
Enter; save and pull again), then:

```
@workbook
years @C2# = SEQUENCE(1, 5, 2025);
Sales @C3# = SEQUENCE(1, 5, 20000, 3400);
COGS @C4# = Sales * 0.6;
Gross_income @C5# = Sales - COGS;
SGA_expenses @C6# = Sales * 0.2;
EBITDA @C7# = Gross_income - SGA_expenses;
Depreciation @C8# = SEQUENCE(1, 5, 2000, 42);
EBIT @C9# = EBITDA - Depreciation;
Taxes @C10# = EBIT * 0.3;
Unlevered_net_income @C11# = EBIT - Taxes;
```

Notice the `#` after each address, and that `@workbook` is gone from all lines but
`years`. Both are explained below ([§4](#4-spills-and-the-), [§5](#5-a-second-sheet-scope-and-workbook)).
Write formulas exactly as in Excel's formula bar, in English, with commas between
arguments, whatever your Excel's language.

While you type, xln completes names (`Gro` → `Gross_income`) and functions, shows the
arguments of a function as you type its `(`, and underlines mistakes at once.

### Step 4: Build

Right-click the `is-model.xln` folder (or any `.xln` file) → **Build workbook**. A dialog
lists what will change (`fill IS!C3 (Sales)`, `move Sales to sheet IS`, …). Confirm.

If the workbook is open in Excel, xln offers **Close in Excel and build**: Excel closes
it (asking first if you have unsaved changes), xln writes the file, and Excel opens it
again. **Build and reopen in Excel** does the same in one step. In the end Excel comes to
the front with your workbook; when there is nothing to build, it just opens the workbook
in Excel.

### Step 5: Check in Excel

Look at the numbers. Excel recalculates everything when it opens the file. Then close
the workbook **without saving** if you want to go on editing in xln, or save it if you
changed something in Excel.

### Step 6: Again

From here, the loop is:

1. Change layout or labels in Excel (new rows, new slots), save → **Pull**.
2. Write or fix formulas in xln → **Build** (and reopen) → look in Excel.

**Edit one side at a time.** If you change a formula in Excel and the same formula in
xln, the build stops and shows both versions side by side.

## 4. Spills and the `#`

`SEQUENCE(1, 5, 20000, 3400)` gives five numbers, which Excel spills from C3 to G3. The
`#` in the address says what the **name** covers:

- `Sales @C3# = …`: `Sales` is the whole spill, C3:G3. Then `COGS @C4# = Sales * 0.6`
  computes five years, and spills too.
- `Sales @C3 = …`: `Sales` is the single cell C3. Then `Sales * 0.6` is one number, the
  first year, and nothing in the formula tells you.

So: **for a row of years, write the `#`.** If you forget, and the formula spilled the
last time the workbook was saved, xln warns on the address:

> `Sales @C3: the formula spills over C3:G3, but Sales covers only C3 (write @C3# to name the spill)`

and the light bulb offers *Name the whole spill: @C3#*. Adding or removing the `#` is an
ordinary edit; the next build moves the name.

Why not one name per year cell? Because a spilled block grows and shrinks with its
formula, and every reader follows it. A fixed range like `C3:G3` does not; the audit
reports fixed references into a spill (check C9) for that reason.

## 5. A second sheet: scope and `@workbook`

Add a sheet `Ratios` in Excel. Labels in B, slots in C made with Create from Selection,
plus a year header in C2 as a formula `=years`. Save, pull. Then write:

```
@C2 = years;
ROS @C4# = IS!EBIT / IS!Sales;
EBITDA_on_Sales @C8# = IS!EBITDA / IS!Sales;
```

### Local and workbook names

Each name is either **local** to one sheet or visible in the **whole workbook**:

- In `names/sheets/Ratios.xln`, every name is local to Ratios, unless the line above it
  says `@workbook`.
- A local name of another sheet is read with the sheet in front: `IS!Sales`. Written
  bare, `Sales` on Ratios is `#NAME?` in Excel. xln says so while you type:
  > `Sales is local to IS, not to Ratios: unqualified it is #NAME? in Excel; write IS!Sales`

  and the quick fix *Qualify: IS!Sales* writes it for you.
- A workbook name is read bare everywhere: `years` above.

Local names keep a model tidy: every sheet can have its own `Sales`, and the sheet in
front of a name tells the reader where it lives. That is why in step 3 we removed
`@workbook` from the IS lines and kept it on `years`, which every sheet reads.

### Changing scope

Both sides can change a scope:

- **In xln:** add or remove the `@workbook` line above the name, then build. While you
  edit, xln notes *the build moves it from workbook scope to sheet IS*.
- **In Excel:** the Name Manager cannot change a scope, so delete the name and create it
  again with the other scope; save and pull.

A hint (dotted underline) suggests *Remove @workbook* on a workbook name that only its
own sheet reads. It is advice: ignore it if you want the name workbook-wide. When another
sheet reads the name, there is no hint.

### Names on a sheet's cells

A workbook name that points at fixed cells of one sheet is in that sheet's file, under
`// Other names on <Sheet>.`, with `@workbook` above it. These are the names Create from
Selection makes on cells that are not formula cells or empty cells of their own: the cells
a formula spills into, cells holding a value, a range. On the Mortgage sheet, where
`A12 = XLOOKUP(…)` spills across A12:G12, Create from Selection (Top row) on A11:G12 gives:

```
// Other names on Mortgage.

Mortgage_list = Mortgage_terms[Mortgage];
@workbook
amount = Mortgage!$B$12;
@workbook
pmt_freq = Mortgage!$C$12;
```

They work like the slots above: remove `@workbook` to make one local to Mortgage (the
build moves it; the quick fix also drops `Mortgage!`, which a local name does not need),
or leave it. `names/_unmanaged.xln` keeps the workbook names that are not
tied to one sheet's cells: constants, formulas, references across sheets.

**A name taken from a value.** With Left column ticked too, Excel names the row B12:G12
after A12's current text, `bullet04`, a formula result; and after the corner cell A11,
`Mortgage`, for the whole block. When the lookup changes, A12 shows another text but the
name stays `bullet04`. The pull notes such names:
> `bullet04: named after the current value of Mortgage!A12, a formula result; the name stays as is when that value changes`

It changes nothing; delete the name in Excel if you do not want it.

## 6. Modules and LAMBDAs

A **module** is a family of names with a common prefix: `ANA.GROW`, `ANA.DOUBLE`. Each
module is one file, `names/ANA.xln`. Its names are workbook names.

### A new module

Right-click the project folder → **New module** (or the (+) button in the Explorer's
title bar), type the prefix `ANA`. xln writes `names/ANA.xln` with an example:

```
/**
 * Grows a value by a rate, compounded over a number of periods.
 * @param value the starting value
 * @param rate the growth rate per period (0.05 for 5%)
 * @param [periods] how many periods; 1 when omitted
 */
ANA.GROW = LAMBDA(value, rate, [periods],
    value * (1 + rate) ^ IF(ISOMITTED(periods), 1, periods)
);
```

Nothing reaches the workbook until you build. Then use it on Ratios:

```
Sales_growing @C7# = ANA.GROW(INDEX(IS!Sales, 1), 0.035, SEQUENCE(1, 5, 1, 1));
```

### Writing a LAMBDA

- Write the full name, prefix included: `ANA.GROW = …`, in `names/ANA.xln`.
- Spread long formulas over several lines; spaces and line breaks do not matter. The
  definition ends at `;`.
- **Doc comment:** the `/** … */` above the name becomes its comment in Excel's Name
  Manager (at most 255 characters). The first line says what the function does; one
  `@param name text` line per parameter. VS Code shows them when you hover the name and
  while you type its arguments. If an `@param` names no parameter (you renamed one), xln
  warns and offers *Rename @param …*.
- **Optional parameters:** write `[periods]` in the parameter list and test
  `ISOMITTED(periods)` in the body. A call may leave it out: `ANA.GROW(100, 0.1)`.
- Calling with the wrong number of arguments is an error: `FN.PREV(row) takes 1
  argument; it is given 2`.
- A name may be spelled like a built-in function (`Rate`, `Fact`) in Excel, but a call
  then reaches the built-in. xln refuses new names like that.

### What you see in Excel

The Name Manager shows `=LAMBDA(value,rate,[periods], …)` and your comment, ending with
a tag like `[xln ANA #40623a]`. The tag says which module wrote the name. xln adds it and
removes it again when it pulls; you never type it. A line `// @version 1.2` near the top
of the module file puts a version in the tag, and a pull writes it back.

A module name is visible from every sheet. To make one local to a sheet, put
`@sheet(IS)` on the line above it (like `@workbook` in a sheet file, it applies to that
name only). Files pulled before 2026-10-07 may have `@scope(IS)` blocks instead: they
still work, and the light bulb offers *Convert to per-name @sheet*.

## 7. How the editor helps

| Help | Where | What it does |
|---|---|---|
| Completion | typing in a formula, Ctrl+Space | Variables of LET/LAMBDA first, then the sheet's own names, workbook names, other sheets' names (`IS!Sales`), module prefixes (`FN.` then its members), Excel functions |
| Signature help | after `(` and `,` | The parameters of the function, with your `@param` text |
| Hover | mouse over a name | Kind, scope, its formula, its comment; for a spill, its cell, size and first value; who uses it |
| Go to definition | right-click a name → *Go to Definition*; Ctrl+click (⌘+click on a Mac); F12 | Jumps to the name's line |
| Find references | right-click → *Find All References*; Shift+F12 | Every use, in `.xln` files **and** in the workbook's cells, formats and validations |
| Rename | right-click a name (where it is defined or used) → *Rename Symbol*; F2 | Renames it everywhere in the project (Shift+Enter: preview first); the next build renames it in the workbook and in its cells (below) |
| Outline | Outline view | The names of a file in order |
| Search names | *xln: Search names* | Finds words in names, formulas and comments |
| Formula view | right-click the workbook, or the table icon on a sheet file | One read-only page per sheet: every formula, its address, its size and its saved value. *Calculation order* lists each formula after what it reads; *Workbook formula view* does the whole workbook |
| Audit | *xln: Audit workbook* | A report of checks C1–C15: wrong prefixes, broken references, unused names, fixed references into spills, copy drift, cycles, AFE's copy of the names, labels that no longer match their names; plus counts of names and spills. An unused constant or LAMBDA is a warning; an unused name on cells (a result shown on the sheet) is info |
| Live checks | as you type | Errors, warnings and hints, below |

**Function keys are shortcuts, never needed.** Every action above is on the right-click
menu. On a Mac, F2 and F12 often do not reach VS Code: the keyboard sends brightness or
volume unless you hold fn, and macOS may keep F12 (and fn+F12) for itself, depending on
the keyboard and on *System Settings → Keyboard → Keyboard Shortcuts* (Mission Control,
Function Keys). In vscode.dev the browser can take some keys too. Use the menu, or
⌘+click for Go to Definition.

**What xln did:** the *xln* output panel (View → Output, then *xln* in the list) logs
every action, desktop and browser alike: each pull, build, rename (F2), audit, formula
view, library action and new module gets a line with the time (`14:02:31 xln build
is-model.xlsx: 2 changes`) and the details below it: what changed, where it was written,
why it was refused, how long it took. When a message on screen has gone, look there.

### Errors, warnings, hints

The Problems panel (View → Problems) lists what xln found:

- **Errors** (red): the build will refuse. An unknown name or function, another sheet's
  name without its sheet, a changed address, a wrong number of arguments, a comment over
  255 characters. Fix them first.
- **Warnings** (yellow): the build goes ahead, but look. A name without `#` on a spill,
  a built-in function with an odd number of arguments, an `@param` that names nothing,
  two doc comments before one name, a name in a file the next pull would move it out of
  (`Rate has no ANA. prefix: a pull will put it in _unmanaged.xln`; the light bulb moves it
  or renames it). `xln check` and the build list the same warnings.
- **Info** (blue): news, such as *the build moves it from workbook scope to sheet IS*.
- **Hints** (dots under the text, not in the panel): suggestions, such as matching a
  name's spelling (`EbIT` → `EBIT`).

Most come with a **quick fix**: click the light bulb, or press Ctrl+. (Cmd+. on a Mac).

A problem in a formula you did not touch (it is so in the workbook) is shown as a
warning: the build leaves that formula as it is. Fix it in xln to have it written. But if
the formula broke because you renamed, moved or deleted the name it reads, it is an
error: the build refuses until you change the formula too (or keep the name).

### Renaming a name

Right-click the name, where it is defined or where a formula uses it, and choose
**Rename Symbol** (shortcut: F2, see the note on function keys above). Type the new name and press Enter: VS Code
changes the name's line, adds a line `@renamed(OldName)` above it, and changes every
formula of the project that uses it. Press **Shift+Enter** instead to see a preview of
these changes first. Then **Build**. The `@renamed` line tells the build that this is
the old name renamed, not a name deleted and another created; it stays until the rename
is built.

The build renames the name in Excel's Name Manager and changes it in every formula of
the workbook that uses it: cells, conditional formats, data validations, other names.
Only the name changes in those formulas, nothing else, just as when you rename in
Excel's Name Manager. Other things are left alone: a word in quotes (`"Rate is "`), a LET
or LAMBDA variable with the same spelling, a Table column `Tbl[Rate]`.

```
// before                          // after F2 Rate → Pace
Rate = 0.1;                        @renamed(Rate)
                                   Pace = 0.1;
@C1 = RateX + Rate;                @C1 = RateX + Pace;
@C2 = "Rate is " & Rate;           @C2 = "Rate is " & Pace;
@C3 = LET(Rate, 5, Rate * 2);      @C3 = LET(Rate, 5, Rate * 2);
```

The build removes the `@renamed` line once it has written the rename into the workbook,
and says so in the xln output (`removed @renamed(Rate) from names/_unmanaged.xln: the
rename is built`). An open editor gets the change and is saved; there is nothing to
undo, and the next pull changes nothing. If you rename the name back before building,
F2 removes the line itself; renamed back after a build, F2 writes a new
`@renamed(NewName)` and the next build removes it. A refused build leaves it, and so does
a build in the browser, which writes a copy (`<name>.xln.xlsx`) and not the workbook.
This is the only change a build makes to your `.xln` files.

An `@renamed` line can stay after its rename is in the workbook: after a build in the
browser whose copy you then used as the workbook, a build by an older version of xln,
or a line you wrote yourself for a rename made in Excel. The editor then shows it faded,
with a hint *@renamed(Rate): the rename is built; this line can go* and the quick fix
*Remove @renamed(Rate)* (Ctrl+. or ⌘.). You can also leave it: the next build that
writes the workbook removes it, with the same line in the output. A build with nothing
to write does not touch your files.

A name made with *Create from Selection* keeps its label: rename `Gross_income` to
`Gross_ind_income` and the cells beside it still read "Gross income" and "Gross_income",
since xln never writes cell values (Excel owns them). So the build tells you exactly how
to fix them in Excel, with Find & Replace (the **label notice**, in the xln output; the
message after the build has a *Show labels to fix* button):

```
Gross_income → Gross_ind_income: 2 labels still read the old name: IS!A5, IS!B5
In Excel, on sheet IS: Find & Replace (Ctrl+H on Windows; Ctrl+H or ⌘⇧H on Mac; or Home → Find & Select → Replace)
  1. Find what:     Gross income
     Replace with:  Gross ind income
  2. Find what:     Gross_income
     Replace with:  Gross_ind_income
  Within: Sheet · Look in: Formulas · Match entire cell contents ✓ · then Replace All (for each pair)
Check by eye (not in the replace): IS!H5 "Gross income (EUR)"
```

xln looks at the sheet of the name's cells, in the name's rows and columns, at text you
typed (not a formula's result). A cell is listed when its whole text gives the old name
the way Create from Selection converts a label (spaces and other characters become `_`,
as Excel does it: `Margin %` gives `Margin`, `Q1` gives `Q1_`),
so "Gross income" counts as well as "Gross_income"; there is one Find what / Replace
with pair per text, and the replacement keeps the label's style (spaces stay spaces).
Text that only resembles the old name ("Gross income (EUR)") is listed under *Check by
eye*, never in the replace. A LAMBDA or a constant has no label: no notice.

To apply it: in Excel, go to the sheet, open Replace (the shortcut above, or Home → Find
& Select → Replace; on the Mac also Edit → Find → Replace), click *Options*, set
**Within: Sheet** and tick **Match entire cell contents** (Look in is *Formulas*, the
only choice for Replace). Type or paste each pair and click *Replace All*. Match entire
cell contents is what keeps it safe: a formula cell's content starts with `=` and never
matches, and a longer text that contains the words is left alone. If the notice says
*Replace All also changes* another cell with the same text, use *Find Next* and
*Replace* instead to skip it. In VS Code, *Show labels to fix* (or *xln: Show labels to
fix after a rename* later) opens the notice and offers each text to copy. Save in Excel,
then pull.

`xln rename` (with the workbook beside the project) prints the same notice as a heads-up
for after the build. After the build, the audit also flags the label beside the name
(C15, info, on the name), with the same Find & Replace in its hint.

xln refuses the rename, and says where, when:
- the new name is taken, not a valid name, or spelled like an Excel function (`Growth`
  is GROWTH);
- a formula would then read something else (a LET variable `x` and the new name `x`);
- the workbook uses the name where xln cannot change it: a **chart**, a **Table column**
  formula, a **pivot table**, a **hyperlink** or a **form control**. Then rename it in
  Excel's Name Manager instead (it changes those too), save, and pull.

From the command line: `xln rename book.xln Rate Pace` does the same (`IS!Sales` for a
name local to IS), and checks first that the next build would accept it.

## 8. The library

A library is a folder of `.lambda` files, one function each, shared across your
workbooks. Tell a project where it is, in `is-model.xln/xln.config.json`:

```json
{ "library": "~/models/_shared/lib" }
```

### Using it

- **Insert:** type a library function's name in a formula (it appears in completion,
  marked *from library*), or run **xln: Insert library function**. xln adds the
  definition to `names/FN.xln` with its comment and a line `@from(lib #eae297)`.
- **Library status:** right-click the project → **Library status**. A report lists each
  library function and its state. In module files, a line above each function (a *code
  lens*) shows the same state and the actions that fit it.
- **Take the library's version:** replaces your copy with the library's.
- **Publish to library:** writes your copy to the library, after showing the change and
  asking.
- **Record library base:** for a copy equal to the library that has no `@from` line yet.

### `@from`: where a copy came from

`@from(lib #eae297)` above a function records which library version your copy started
from. xln writes it when you Insert, Take, Publish or Record; never otherwise. With it,
xln can tell who changed what:

| Status | Means | What you might do |
|---|---|---|
| identical | your copy equals the library | nothing |
| outdated | the library changed; your copy did not | Take |
| modified | you changed your copy; the library did not | Publish it, or Take to undo your edit |
| both changed | you and the library both changed it | Show diff, then decide |
| differs | no `@from`: xln cannot tell who changed it | Publish, or Take (xln asks first) |
| missing | in the library, not in your workbook | Insert, if you want it |
| local only | your function, with a library prefix, not in the library | Publish, if it is ready |

*Not built yet* next to a status means your source has it but the workbook does not yet:
build.

Take asks before replacing a copy you may have edited (*Discard your edit of FN.X?*).
Nothing moves between library, project and workbook unless you ask.

The workbook keeps `@from` in the name's tag, `[xln FN 1.2 #636cf1 lib#eae297]`, at the
end of its comment, and a comment holds at most 255 characters. A doc comment too long
to carry the tag is written without it, and then the workbook does not record where the
copy came from. The editor and `xln check` warn on the doc comment as you type (*the doc
comment is 235 characters; with its provenance tag (29) it passes Excel's 255, … shorten
it by 9 characters*), and Insert, Take, Publish and Record say the same before they
write. They write anyway: shorten the doc comment in your module file (the tag goes on
that comment, not on the library file's summary), then build.

## 9. Safety: what protects your workbook

- **Excel's lock file.** While Excel has the workbook open it keeps a hidden file
  `~$is-model.xlsx`. A desktop build sees it and stops (or offers *Close in Excel and
  build*).
- **A backup.** Every desktop build keeps the previous file as `is-model.backup.xlsx`.
- **Read-back.** After writing, xln reads the file again and compares every name and cell
  it wrote with your source. On a mismatch it puts the original back.
- **Only what changed.** A build rewrites the names and the cells you changed; every other
  part of the file is copied byte for byte.
- **The pull guard.** A pull would replace your source with the workbook's version, so
  if you have edits not built yet, it stops and lists them: **Build first** or **Discard
  and pull**.
- **Conflicts.** A name or cell changed both in Excel and in xln since the last pull stops
  the build and opens both versions side by side. Nothing is merged silently.
- **Verify** (command line): after you open and save the built workbook in Excel,
  `xln verify is-model.xlsx` compares every value with the backup and lists the cells
  that changed. A build that should not change numbers must show none. If either file
  was never saved by Excel (written by Python, or by xln and not opened since), verify
  says so first: its values are not Excel's, so open and save it in Excel, then verify.

## 10. Common messages and what to do

| Message (short) | Why | What to do |
|---|---|---|
| *X is local to IS, not to Ratios: unqualified it is #NAME? in Excel; write IS!X* | another sheet's name without its sheet | quick fix *Qualify* |
| *X is not a name in scope, a LET/LAMBDA variable or a function (did you mean Y?)* | a typo, or a name that does not exist yet | quick fix *Did you mean*, or create the name |
| *XLOKUP(…): no built-in function and no LAMBDA of that name; Excel would store it as _xludf.XLOKUP* | a misspelled function | quick fix *Did you mean XLOOKUP?* |
| *FN.PREV(row) takes 1 argument; it is given 2* | wrong number of arguments | fix the call |
| *Far: the address is set in Excel and read-only; the last pull had @B9* | you edited an address | quick fix *Restore the address*; move cells in Excel |
| *no cell statement at Ratios!C20 in the last pull* | you added a cell line by hand | write a first formula (or a name) in that cell in Excel, save, pull |
| *Sales @C3: the formula spills over C3:G3, but Sales covers only C3* | the name covers one cell of a spill | quick fix *Name the whole spill: @C3#* |
| *'#' on a cell left empty is not written yet* | `Name @C6# = ;` | give the slot a formula |
| *Excel has is-model.xlsx open (~$is-model.xlsx exists): nothing written* | the workbook is open (command line) | close it, or use Build and reopen |
| *refused: the project has 1 source edit not built yet, which a pull would replace* | you pulled with unbuilt edits | Build first, or Discard and pull |
| *deleting Fn refused: 1 place in the workbook refers to it by name and would break: cell S1!C10* | something still reads the name | change that formula too, in the same build, or keep the name |
| *… the source renames X to Y, so write Y here* | you renamed by hand and this formula still has the old name | write Y there; or undo and rename with F2, which changes every formula |
| *@renamed(X): the rename is built; this line can go* (hint, faded) | the rename is in the workbook already; the line was left by a browser build, an older xln, or by hand | quick fix *Remove @renamed(X)*; or leave it: the next build that writes removes it |
| *renaming X to Y refused: … cannot reach these places: chart …* | a chart, Table column, pivot, hyperlink or form control uses the name | rename it in Excel's Name Manager, save, pull |
| *renaming X to Y refused: … the new name would read something else* | a LET/LAMBDA variable or another sheet's name has the new spelling | choose another name |
| *the doc comment has 260 characters; the Name Manager takes 255* | comment too long | shorten it (a line break counts 2) |
| *the doc comment is 235 characters; with its provenance tag (29) it passes Excel's 255, so the build writes it without the tag …* | the comment fits, the tag xln adds does not | shorten it by the characters named; otherwise the workbook loses the tag (and its `@from`) |
| *@param count names no parameter of the LAMBDA* | doc and parameters disagree | quick fix *Rename @param* |
| *there is no sheet Foo in the workbook: sheets are created in Excel, then pulled* | a file you created in `names/sheets/` | create the sheet in Excel, or delete the file |
| *original.xlsx has no values Excel calculated (it was never saved by Excel)* (`xln verify`) | the file before the build was written by Python or xln and never opened in Excel | open it in Excel, save, and verify again; otherwise the comparison is empty |
| *this workbook has no link to Prices.xlsx (it has none): Excel writes a link when a formula first names the other workbook, xln does not; type the reference once in Excel, save, then pull* | a formula names another workbook (`[Prices.xlsx]Sheet1!A1`, `Prices.xlsx!Rate`) that this file does not link to yet | in Excel, type that reference once in any cell or name, save, pull; then the source can use it. A workbook already linked is written by its file's name and needs nothing |
| *Rate collides with the built-in function RATE* | name spelled like a function | pick another name (`GrowthRate`) |
| *Group is also an Excel 4.0 macro function: Excel may call that instead or refuse the name (AFE #10, not measured)* (warning) | a LAMBDA spelled like an old macro-sheet function (`Group`, `Get.Cell`, `Evaluate`, `Files`) | pick another name (`GroupRows`); the build writes it anyway. `xln rename` to such a name prints it too |
| *IS!B5 reads "Gross income", which Create from Selection makes Gross_income; the name on IS!C5:G5 is Gross_ind_income, probably renamed from Gross_income* (info, C15) | you renamed a name made from its label; the label in Excel still has the old text. xln never writes cell values | fix the label in Excel with the Find & Replace in the hint (Match entire cell contents; see *Renaming a name*), save and pull; or rename the name back; `"C15": "off"` under `audit.rules` in `xln.config.json` silences it |
| *the workbook carries modules of Microsoft's Advanced Formula Environment … this build changes 1 name that AFE's modules also define* | the workbook was edited with AFE (§13) | bring AFE's module in line before saving it from AFE, or leave AFE's modules unsaved |

## 11. Working with AI completion (Copilot)

GitHub Copilot and similar tools suggest grey "ghost text" as you type. They are useful
for long LET formulas and for comments. But they do not know your workbook:

- They invent names that look right (`Revenue` where yours is `Sales`). xln's live checks
  underline an unknown name at once: read the red squiggles before you build.
- They repeat patterns from the file, including old forms such as `@scope(IS)` blocks.
  xln no longer writes those (scope is per name: `@workbook` in sheet files, `@sheet(IS)`
  in module files); the light bulb converts them. They may also invent type
  declarations (`Periods : scalar = …`), which xln does not accept.
- They may change the address part of a line. The address is Excel's: xln reports any
  change as an error with *Restore the address*.
- xln's own completion (the list with icons) comes only from your project and Excel's
  function catalogue. Word-based suggestions are turned off for `.xln` files.

A good habit: accept ghost text only for the formula after `=`, then wait for the
Problems panel to be quiet before you build. An AI agent working on a workbook can run
`xln check` as a critic; it reports the same findings as the editor.

### AI agents: the MCP server

Agents such as Claude Code and Claude Desktop can use xln directly through its MCP
server, `xln-mcp`. It gives them six tools:
- `xln_check`: the same verdict as `xln check`;
- `xln_names`: list and search names;
- `xln_pull`;
- `xln_build_plan`: a dry run;
- `xln_build`;
- `xln_lib_status`.

A typical session: you ask the agent to add a LAMBDA, and it does four things.
1. It pulls the workbook and edits `names/FN.xln`.
2. It checks the project until there are no errors.
3. It shows you the plan.
4. It builds, but only after you say yes.

The `.xln` files in between are ordinary text, so you can read or diff what the agent
wrote before anything reaches the workbook.

To connect it to Claude Code, from a checkout where you ran `npm install` and
`npm run build`:

```
claude mcp add xln -- node /path/to/excel_dim/packages/mcp/bin/xln-mcp.js --root ~/models
```

`--root` is the folder the agent may work in. The server refuses every file outside it.
A build through the agent keeps every protection of section 9:
- the lock file;
- the backup;
- the read-back;
- conflicts that stop the build.

The agent also cannot build without an explicit confirmation, and the server never
closes or reopens Excel. The Claude Desktop configuration and the tool list are in
`packages/mcp/README.md`.

## 12. Working in the browser (vscode.dev)

Everything in this guide works in vscode.dev, with these differences:

- **The build writes a new file.** The browser cannot see Excel's lock file, so xln never
  overwrites your workbook there: it writes `is-model.xln.xlsx` next to it and leaves
  `is-model.xlsx` untouched. Open the new file in Excel to check it.
- **To continue from the built file:** close the workbook in Excel, replace
  `is-model.xlsx` with `is-model.xln.xlsx` (keep a copy of the old one), then Pull. The
  workbook now has your edits, so the pull goes ahead (and a build says *up to date*).
- **No Build and reopen**, and no *Close in Excel and build*: the browser cannot drive
  Excel.
- **No file watching:** after changing files outside VS Code, run **xln: Reload**.
- **Edge or Chrome only.** In Chrome, Cmd/Ctrl+T opens a browser tab; use Cmd/Ctrl+P and
  type `#` to search names.

For long editing sessions, desktop VS Code is smoother; the browser is for when you
cannot install anything.

## 13. Workbooks edited with AFE

Microsoft's Advanced Formula Environment (AFE, in the Excel Labs add-in) writes modules
in the same syntax as xln and pushes their names into the Name Manager. It also keeps
**its own copy of the modules** inside the workbook (a hidden part of the file; AFE 1.0
used a very hidden sheet). xln works from the Name Manager only, so in such a workbook
the names exist twice: in the Name Manager, and in AFE's copy.

What xln does:

- **Pull** reads the names from the Name Manager, as always, and adds a note: the
  workbook carries AFE modules (which ones, how many names), and how many of them differ
  from AFE's text. AFE's copy is never read into your project.
- **Check** (and *xln: Audit workbook*, and the Problems panel) has a check for it, C14:
  an info saying what AFE's copy holds; an **info on each name whose AFE text differs**
  from the workbook (whitespace, case and number spelling aside), with both texts; an
  info listing names AFE's modules define that the workbook does not have. They are info,
  not warnings: the build's warning is the moment that matters, and a difference found
  later is a state to know about.
- **Build** leaves AFE's copy exactly as it is: it never edits or deletes it. When the
  build changes names that AFE's modules also define, it **warns and lists them**: AFE's
  copy still has the old text.

What you do after a build that warned:

- If you no longer use AFE on this workbook, nothing: AFE's copy is inert while AFE is
  closed. xln will not remove it: AFE manages it.
- If you still use AFE, **do not save AFE's modules before bringing them in line with the
  Name Manager**: when AFE opens it shows its own text, and saving its modules writes that
  text back over the names, undoing your build. Edit the module in AFE to match (the C14
  findings show both versions), then save it; run `xln check` again to see it is clean.
- Keep one side as the place where you edit each name. A name you edit in xln, edit in
  xln; a name you prefer to edit in AFE, edit there and pull.

AFE writes its modules with your locale's separators (`;` in Italian Excel). xln then
cannot compare them with the names and says so (C14, "not compared").

## 14. FAQ

**Do my `//` comments survive?** Not a pull. Every pull writes the files again from the
workbook, and only what Excel stores comes back: names, formulas, doc comments (the Name
Manager comment). Put what you want to keep in doc comments. The pull tells you which
files would lose comments or layout.

**Can I add a row or a new line item from xln?** No. Insert the row and type the label in
Excel, name the cell (Create from Selection), save, pull. Then write its formula.

**Can I move a cell?** Move it in Excel (cut and paste, or insert rows). A named cell
follows its name; the next pull shows the new address.

**I selected B3:G11 for Create from Selection.** Each name then covers C:G as a fixed
range (`Sales = IS!$C$3:$G$3`) and is not a slot. Delete those names in Excel, select
B3:C11, and create them again.

**Why do some formulas show `@C5 = …;` with no name?** Formula cells without a name. You
can edit their formula too. A row of cells that hold one formula filled across shows once,
as `@B40:G40 = …;`.

**How do I clear a cell's formula?** Leave the right side empty: `Sales @C3 = ;`.

**How do I rename a name?** Right-click it → *Rename Symbol* (or F2), then Build: the name changes
everywhere, in the project and in the workbook's cells (§7, *Renaming a name*).

**How do I delete a name?** Delete its line (not for a name on a cell: delete that one
in Excel's Name Manager and pull).

**What is `xln.config.json`?** Your project settings: the library folder and the audit's
settings. A pull never overwrites it.

**What are `xln.lock.json` and `workbook.manifest.json`?** xln's memory of the last pull
or build. They are hidden; do not edit them (*xln: Show project internals* shows them).

**Can I put the project in git?** Yes: the `.xln` folder is plain text, made for diffs.

**Does it work with `.xlsm`?** Yes. Not with `.xlsb`, or with workbooks open for
co-authoring in OneDrive or SharePoint.

**My Excel is in Italian. Do I write `;` between arguments?** No. Always English names
and `,`, as Excel stores them. Excel shows them in your language.

## 15. Cheat sheet

```
FILES
  book.xln/names/sheets/IS.xln     the IS sheet: its cells, names on its cells, local names
  book.xln/names/FN.xln            module FN: FN.* names
  book.xln/names/_unmanaged.xln    other workbook names (not on one sheet's cells)
  book.xln/xln.config.json         settings (library, audit)

LINES
  Rate = 0.03;                      a name
  Sales @C3# = SEQUENCE(1,5,100);   a name on cell C3, covering its spill
  Sales @C3 = 100;                  a name on cell C3 alone
  Tax @C10 = ;                      a slot: a name on an empty cell
  @C5 = years;                      a formula cell without a name
  @B40:G40 = SUM(B30:B39);          one formula filled across
  /** Comment. @param x … */        Name Manager comment (max 255)
  @workbook                         next name: visible from every sheet (sheet files)
  @sheet(IS)                        next name: local to IS (module files)
  @hidden                           next name: hidden in the Name Manager
  @renamed(Old)                     next name: was called Old (F2 writes it)
  @from(lib #eae297)                next name: library version it came from
  // comment                        lost at the next pull

FORMULAS
  as in Excel's formula bar, in English, commas between arguments
  IS!Sales        another sheet's local name
  C3#             a spill        [p]   optional LAMBDA parameter
  end every line with ;          spaces and line breaks do not matter

THE LOOP
  Excel: labels, rows, Create from Selection (label column + one cell column), save
  VS Code: Pull workbook → write formulas → Build workbook (and reopen)
  Excel: check, close without saving (or save and Pull if you changed something)

PROBLEMS
  red error: build refuses     yellow warning: look     dots: hint     Ctrl+. : quick fix

COMMAND LINE (optional)
  xln pull book.xlsx          xln build book.xlsx [--dry-run] [--reopen]
  xln check book.xlsx         xln verify book.xlsx
  xln rename book.xln Old New
  xln lib status book.xln     xln lib take|publish|base book.xln Name
```
