# What Excel Labs (AFE) users asked for

*For users of Microsoft's Advanced Formula Environment (AFE, now part of the Excel Labs
add-in) who are looking for somewhere to go. As of 2026-10-08.*

> **A word of thanks.** xln was designed and built on the results of the AFE project: its
> module syntax, its doc comments, and the requests its users filed. We would have liked
> to work with AFE directly. One practical hurdle prevented it: our university's Office
> installations do not allow add-ins such as Excel Labs. So xln works on the workbook
> file instead, with no add-in, and reads and respects AFE's own copy of the modules
> where it finds one (§5).

AFE brought names as code to Excel: modules, `name = formula;`, `/** doc */` comments
synced to the Name Manager, a code editor for LAMBDAs. xln reads the same module syntax
on purpose. AFE did it first, and much of xln's format comes from it.

AFE still installs as part of Excel Labs. The last version users cite is 1.4, of January
2024. Microsoft has announced no successor, and the issue
[Future of AFE?](https://github.com/microsoft/Excel-Labs/issues/39) has no reply from
Microsoft. Over four years AFE users filed requests that a live add-in could not easily
meet: modules as files, Git, a VS Code extension, shared libraries, offline use. This
page goes through those requests and says, for each one, what xln does today, what it
does in another way, and what it does not do.

xln is young: the first public release, 0.1.0, came out on 2026-10-08 (now 0.1.1), on the
VS Code Marketplace and Open VSX. Everything below is true of the current code and its tests. Where xln does not answer a request, the page says so.

Contents:
1. [xln in one paragraph](#1-xln-in-one-paragraph)
2. [The requests, by theme](#2-the-requests-by-theme)
3. [What xln does differently from AFE](#3-what-xln-does-differently-from-afe)
4. [What xln does not do](#4-what-xln-does-not-do)
5. [Moving from AFE](#5-moving-from-afe)
6. [Sources](#6-sources)

## 1. xln in one paragraph

xln reads an `.xlsx` or `.xlsm` file directly and writes its defined names, LAMBDAs and
cell formulas as text files in a folder next to it (a *pull*). You edit them in VS Code
(desktop on macOS or Windows, or vscode.dev in Edge or Chrome), where they are checked as
you type. A *build* writes your edits back into the file, after a dialog that lists every
change. There is no add-in and nothing runs inside Excel; Excel does not need to be
running. The same engine runs as a command line (`xln pull`, `xln build`, `xln check`,
`xln lib …`) and as an MCP server for AI agents. MIT licence. The details are in the
[user guide](USER-GUIDE.md) and the [language specification](LANGUAGE.md).

## 2. The requests, by theme

Requests are on two trackers: the old AFE repository
([microsoft/advanced-formula-environment](https://github.com/microsoft/advanced-formula-environment/issues)),
cited as *AFE #n*, and the Excel Labs repository
([microsoft/Excel-Labs](https://github.com/microsoft/Excel-Labs/issues)), cited as
*Labs #n*. The date is the day the issue was opened. Each answer starts with **Yes**,
**Partly**, **Differently** or **No**.

### 2.1 Modules as files: export and import

| Request | xln today |
|---|---|
| [AFE #53](https://github.com/microsoft/advanced-formula-environment/issues/53), 2023-05-12: save modules to an external file, import them from the file system or HTTP, keep them in sync. Still open; a user asked again in September 2026 | **Yes** for files, **No** for HTTP. A pull writes every module as a text file (`names/FN.xln`); a build writes it back. Sync with a shared library: see 2.2. xln downloads nothing |
| [AFE #64](https://github.com/microsoft/advanced-formula-environment/issues/64), 2023-10-12: keep the functions in a plain text file in Git and import or export them as needed | **Yes.** The project folder is plain text. Pull is the export, build is the import |
| [Labs #61](https://github.com/microsoft/Excel-Labs/issues/61), 2025-11-01: import modules from a local file, because GitHub is blocked at work | **Yes.** Everything is local: the project folder, and the library folder (2.2). No GitHub account is needed |
| [AFE #36](https://github.com/microsoft/advanced-formula-environment/issues/36), 2022-12-05: how to load an external library from GitHub "or locally" | **Partly.** Local folders, yes. Import from a URL, no |
| [AFE #56](https://github.com/microsoft/advanced-formula-environment/issues/56), 2023-09-07: make a module from about 100 existing names already called `ModuleName.FunctionName` | **Yes.** A pull puts every name with a prefix before `.` into that module's file, with its comment as a `/** */` doc comment. No script needed |
| [AFE #46](https://github.com/microsoft/advanced-formula-environment/issues/46), 2023-04-03, and [AFE #91](https://github.com/microsoft/advanced-formula-environment/issues/91), 2024-08-05: no way to export functions made in the Names tab, or existing definitions, as a module | **Yes.** xln has one store, the Name Manager. Every name is pulled, whichever tool created it |

### 2.2 Library sharing and updates

| Request | xln today |
|---|---|
| [AFE #53](https://github.com/microsoft/advanced-formula-environment/issues/53), 2023-05-12: link a workbook to an organisation's vetted module and resync it when it changes | **Yes, differently.** A library is a folder of `.lambda` files, one function each. *Library status* (`xln lib status`) compares each copy in the workbook with the library: identical, outdated, modified, both changed, missing. *Insert*, *Take the library's version* and *Publish to library* move a function one way, only when you ask |
| [AFE #19](https://github.com/microsoft/advanced-formula-environment/issues/19), 2022-08-18: save to a gist, not only read from one. Microsoft replied that gists are not ideal for managing libraries | **Differently.** *Publish to library* (`xln lib publish`) writes your copy into the library folder. Sharing the folder (Git, a network drive) is up to you. No gists |
| [AFE #43](https://github.com/microsoft/advanced-formula-environment/issues/43), 2023-03-16: remember where a module came from, and replace old definitions on re-import instead of duplicating them | **Yes.** `@from(lib #eae297)` above a function records the library version the copy came from, and the workbook keeps it in the name's comment tag. *Take* replaces the copy in place |
| [AFE #59](https://github.com/microsoft/advanced-formula-environment/issues/59), 2023-09-11 (raw URLs, because gists cannot belong to an organisation); [AFE #79](https://github.com/microsoft/advanced-formula-environment/issues/79), 2024-04-18 (GitHub Enterprise gists); [AFE #80](https://github.com/microsoft/advanced-formula-environment/issues/80), 2024-04-18 (modules from repository packages) | **Differently.** xln fetches nothing over the network. Keep the library folder in any Git repository (an organisation's, GitHub Enterprise, an internal server) and clone or pull it with Git |
| [Labs #17](https://github.com/microsoft/Excel-Labs/issues/17), 2024-01-05: importing a gist into a module renames its functions `MODULE.FUNCTION`, which surprised the importer | **Differently.** In xln a name is written in full (`FN.GROW`) wherever it appears. Nothing is renamed on import |
| [AFE #100](https://github.com/microsoft/advanced-formula-environment/issues/100), 2026-06-20: copying a sheet to another workbook brings the names but not AFE's module, and some comments are lost | **Partly.** xln keeps no second store in the workbook, so there is no module to leave behind: a pull of the second workbook shows the names it has. xln does not restore comments that Excel drops when copying a sheet |

### 2.3 Git and versioning

| Request | xln today |
|---|---|
| [AFE #64](https://github.com/microsoft/advanced-formula-environment/issues/64), 2023-10-12: diffs, branches and refactorings, as with normal source code | **Yes.** The `.xln` folder is plain text made for `git diff`. Rename Symbol (F2) renames a name across the project, and the next build renames it in the workbook's cell formulas, conditional formats, data validations and other names |
| [AFE #53](https://github.com/microsoft/advanced-formula-environment/issues/53), comment of 2023-06-29: edit modules "using powerful code editors such as VSCode" and trace them in version control | **Yes.** As above. A line `// @version 1.2` in a module file puts a version into each name's comment tag (`[xln FN 1.2 #636cf1]`), so the workbook says which version it carries |

### 2.4 The editor and VS Code

| Request | xln today |
|---|---|
| [AFE #76](https://github.com/microsoft/advanced-formula-environment/issues/76), 2024-02-20: release the formula editor as a VS Code extension. Microsoft replied that it was more work than it could budget | **Yes.** xln is a VS Code extension, for the desktop and vscode.dev: highlighting, completion of names and functions, signature help, hover, go to definition, find references (in `.xln` files and in the workbook's cells), rename, outline. It is on the [VS Code Marketplace](https://marketplace.visualstudio.com/items?itemName=lerzegov-xln.xln) and [Open VSX](https://open-vsx.org/extension/lerzegov-xln/xln) as `lerzegov-xln.xln` |
| [Labs #39](https://github.com/microsoft/Excel-Labs/issues/39), comment of 2025-08-03: Excel's Name Manager shows a formula on one line and does not keep line breaks | **Yes.** Write a definition over several lines; the build stores the line breaks in the name, and the next pull shows them again |
| [Labs #34](https://github.com/microsoft/Excel-Labs/issues/34), 2024-07-30: open the editor in its own window, on a second monitor | **Yes.** VS Code is its own window |
| [Labs #14](https://github.com/microsoft/Excel-Labs/issues/14), 2023-10-21: Copilot in the editor | **Partly.** Copilot and similar tools work in `.xln` files as in any VS Code file; xln's checks underline the names they invent ([user guide §11](USER-GUIDE.md#11-working-with-ai-completion-copilot)). Agents can also call xln through its MCP server |
| [AFE #74](https://github.com/microsoft/advanced-formula-environment/issues/74), 2023-12-05, and [Labs #10](https://github.com/microsoft/Excel-Labs/issues/10), 2023-09-20: descriptions for each LAMBDA argument | **Partly.** `@param name text` lines in the doc comment show in VS Code's hover and signature help, and are checked against the parameters. In Excel they are part of the Name Manager comment, shown as one text: Excel has no per-argument descriptions for LAMBDAs |
| [AFE #21](https://github.com/microsoft/advanced-formula-environment/issues/21), 2022-10-21: enum-typed arguments with completion | **No** |
| [AFE #15](https://github.com/microsoft/advanced-formula-environment/issues/15), 2022-03-18: the editor for unnamed cell formulas too (AFE answered with its Grid tab) | **Yes.** Every formula cell of every sheet is a line in that sheet's file (`@C5 = …;`), in sheet order, and can be edited. A formula view lists them per sheet, in order of appearance or of calculation |
| [AFE #87](https://github.com/microsoft/advanced-formula-environment/issues/87), 2024-06-12: show cell values on hover | **Partly.** Hover on a spilled name shows its cell, size and first value, and the formula view shows each formula's value, as saved in the file. No live values |

### 2.5 Scope and the Name Manager

| Request | xln today |
|---|---|
| [AFE #40](https://github.com/microsoft/advanced-formula-environment/issues/40), 2023-02-24; [AFE #88](https://github.com/microsoft/advanced-formula-environment/issues/88), 2024-07-09; [Labs #71](https://github.com/microsoft/Excel-Labs/issues/71), 2026-04-28: support sheet-scoped names, and changing a name's scope | **Yes.** A name in a sheet's file is local to that sheet unless `@workbook` is above it; in a module file `@sheet(IS)` makes it local. Adding or removing the line changes the scope at the next build. Reading another sheet's local name without its sheet is an error, as it is `#NAME?` in Excel |
| [AFE #85](https://github.com/microsoft/advanced-formula-environment/issues/85), 2024-05-22: saving a module makes hidden names visible again | **Yes.** A hidden name is pulled with `@hidden` above it. A build changes the hidden flag only when you add or remove that line, and writes only the names you changed |
| [Labs #50](https://github.com/microsoft/Excel-Labs/issues/50), 2025-08-03: a description over Excel's limit can corrupt the workbook | **Yes.** A doc comment over 255 characters is an error as you type, and the build refuses it |
| [AFE #5](https://github.com/microsoft/advanced-formula-environment/issues/5), 2022-02-10: sync both ways with the Name Manager; a comment notes stale cell references when a cell moves in Excel | **Yes, differently.** There is no live sync. Every pull is written fresh from the file, so a moved cell shows its new address. A build compares source, workbook and the last pull: a name changed on both sides stops the build and shows both versions |
| [AFE #10](https://github.com/microsoft/advanced-formula-environment/issues/10), 2022-02-16: report names such as `Group` or `Table` (Excel 4.0 macro functions) that fail as LAMBDA names | **No.** xln refuses new names spelled like worksheet functions (`Rate`, `Fact`), but its catalogue does not hold the Excel 4.0 macro functions, so `Group = LAMBDA(…)` passes |

### 2.6 Checks and errors

| Request | xln today |
|---|---|
| [AFE #81](https://github.com/microsoft/advanced-formula-environment/issues/81), 2024-04-20: a better linter | **Partly.** Checks as you type (unknown names and functions, argument counts, unqualified local names, spills not covered by `#`, `@param` mismatches), and a workbook audit, checks C1–C15: missing stored prefixes that give `#NAME?`, broken references, LAMBDA arity, Excel's length and nesting limits, fixed references into spills, unused names, copy drift, name cycles, constants in LAMBDA bodies. No style rewrites (nested `IF` to `IFS`), no performance advice |
| [Labs #57](https://github.com/microsoft/Excel-Labs/issues/57), 2025-09-30: a save of large modules fails with 0x8007000E, and AFE cannot say which LAMBDA is at fault | **Yes, for the second part.** Errors are reported at their file and line before anything is written; the build refuses until they are fixed. xln writes the file itself, not through Excel's add-in API |
| [AFE #20](https://github.com/microsoft/advanced-formula-environment/issues/20), 2022-09-16: a failed sync leaves names in a bad state; asks for a forced re-sync | **Yes, differently.** A build either writes all its changes or none: it reads the written file back and restores the original on any mismatch, and a desktop build keeps the previous file as `<workbook>.backup.xlsx`. A pull is always a full re-read |

### 2.7 Formatting

| Request | xln today |
|---|---|
| [AFE #28](https://github.com/microsoft/advanced-formula-environment/issues/28), 2022-11-19; [Labs #37](https://github.com/microsoft/Excel-Labs/issues/37), 2024-09-07; [Labs #40](https://github.com/microsoft/Excel-Labs/issues/40), 2024-12-13: keep the user's own indentation and line breaks; do not reformat to the pane's width | **Yes.** xln does not reformat what you write. Your line breaks and indentation are stored in the name and come back at the next pull. Only a definition stored on one line and longer than 100 characters is pretty-printed by a pull (one binding per line); the workbook is not changed by that, since definitions are compared without regard to whitespace. There is no *Format Document* command |

### 2.8 Where it runs

| Request | xln today |
|---|---|
| [AFE #60](https://github.com/microsoft/advanced-formula-environment/issues/60), 2023-09-11: offline use; the add-in fails to load offline and disappears | **Yes** on the desktop. The extension and the command line make no network requests. vscode.dev needs a connection to load |
| [Labs #44](https://github.com/microsoft/Excel-Labs/issues/44), 2025-03-19: the Office Store is blocked by the organisation, so the add-in cannot be installed | **Partly.** Nothing is installed in Excel. You need VS Code, or vscode.dev in Edge or Chrome where nothing can be installed; there it installs from the Marketplace, inside the browser |
| Load failures, such as [AFE #98](https://github.com/microsoft/advanced-formula-environment/issues/98), 2025-10-29 | xln does not depend on Excel's add-in host. It needs Excel only to look at the result |
| [Labs #30](https://github.com/microsoft/Excel-Labs/issues/30), 2024-06-14: access to the source code. Microsoft: closed source, no plans to open it | **Yes.** xln is MIT-licensed, source included |

xln runs in desktop VS Code on macOS and Windows and in vscode.dev; it does not run
inside Excel, so not in Excel for the web. It reads `.xlsx` and `.xlsm`, not `.xlsb`, and
not workbooks open for co-authoring in OneDrive or SharePoint.

### 2.9 Locales

| Request | xln today |
|---|---|
| [AFE #3](https://github.com/microsoft/advanced-formula-environment/issues/3), 2022-02-09: support locales with `;` as separator, and localised function names (the most-upvoted AFE issue) | **Differently.** The file stores formulas in English with `,`, whatever Excel's language, and xln works on the file, so xln works with workbooks from any locale. But you write English function names and `,`; xln does not accept `;` or local function names in the source |
| [AFE #78](https://github.com/microsoft/advanced-formula-environment/issues/78), 2024-04-18: module import fails under German localisation | **Yes**, for the same reason: a module file and a library file are in English and mean the same in every locale |

### 2.10 Performance and size

| Request | xln today |
|---|---|
| [AFE #69](https://github.com/microsoft/advanced-formula-environment/issues/69), 2023-10-25: AFE and Excel become very slow with many LAMBDAs | **Not measured.** xln does not run inside Excel, so it does not slow Excel down while you edit. We have not published timings on large workbooks |

### 2.11 Debugging and live evaluation

| Request | xln today |
|---|---|
| [AFE #83](https://github.com/microsoft/advanced-formula-environment/issues/83), 2024-05-09, and [Labs #73](https://github.com/microsoft/Excel-Labs/issues/73), 2026-07-18: the debugger does not handle spilled arrays and fails on most formulas | **No.** xln has no evaluator and no debugger. It does not compute values; Excel does, when it opens the built file |

## 3. What xln does differently from AFE

- **One store: the Name Manager.** AFE keeps its modules in the workbook as a second copy
  of the names. xln keeps no copy: the project is written from the names, and written
  back into them. So `//` comments and blank lines in a module file do not survive a
  pull; the doc comment (`/** */`, the Name Manager comment) does. A pull tells you which
  files would lose comments or layout.
- **Names are written in full.** In AFE a module `M` is a file of bare names that become
  `M.name` in the Name Manager, and siblings call each other bare. In xln the file
  `names/M.xln` holds `M.name = …;` and calls are written `M.name(…)`. A bare name in a
  module file is a plain workbook name, with a warning that a pull will move it to
  `_unmanaged.xln`.
- **No live link.** You edit one side at a time: change layout and labels in Excel, save,
  pull; write formulas in xln, build, look in Excel. On the desktop, *Build and reopen in
  Excel* closes, writes and reopens the workbook in one step.
- **Cells as well as names.** Every formula cell is a line of its sheet's file. Excel owns
  the layout: xln does not add rows, move cells or write labels.
- **Smaller syntax differences.** No `name : type` declarations (an error); annotations
  that AFE does not have (`@hidden`, `@workbook`, `@sheet`, `@renamed`, `@from`); tabs
  and CR accepted. The full list is in [LANGUAGE.md §13.2](LANGUAGE.md#132-differences-from-afe-excel-labs-module-syntax).

## 4. What xln does not do

- Evaluate formulas, show live values, or debug.
- Import from a gist or a URL, or publish to one.
- Keep AFE's module copy in step with the names (section 5).
- Accept `;` separators or localised function names in the source.
- Define enum arguments, or put argument help into Excel's own formula tooltips.
- Flag Excel 4.0 macro function names (`Group`, `Table`) used as LAMBDA names.
- Run inside Excel, or work on `.xlsb` files or workbooks open for co-authoring.

## 5. Moving from AFE

### Open an AFE workbook

1. Save the workbook in Excel. It may stay open: a pull only reads the file. A desktop
   build keeps a backup of the previous file, but a copy of your own does no harm.
2. In VS Code, open the folder that holds it, right-click the workbook → **xln: Pull
   workbook** (or `xln pull book.xlsx`). xln writes `book.xln/` next to it.
3. The pull reads the names from the Name Manager and adds a note: the workbook carries
   AFE modules, which ones, and how many names they define.

What you find in the project:

- Names of an AFE module `M` (stored as `M.name`) are in `names/M.xln`, written in full.
- Names of AFE's `Workbook` module (stored without a prefix) are usually in
  `names/_unmanaged.xln`; in a sheet's file when they point at that sheet's cells.
- Doc comments come back as `/** */`. `//` comments, layout and anything else that lives
  only in AFE's module text do not: xln does not read AFE's text into the project.
- Sheet-scoped names, which AFE did not show, are in their sheet's file.

### What happens to AFE's modules and its custom XML part

AFE 1.1 and later keeps its modules in a custom XML part of the workbook (an
`AFEJSONBlob` element; AFE 1.0 used a very hidden sheet, `AFE_hidden_codesheet_49ddb8b8`,
see [AFE #41](https://github.com/microsoft/advanced-formula-environment/issues/41)).

- **xln never edits or deletes it.** Every build copies AFE's parts byte for byte, and
  the tests check it.
- **xln reads it to compare.** Check C14 (`xln check --only C14`, *xln: Audit workbook*,
  the Problems panel) says what AFE's copy holds, notes (as info, not a warning) each name
  whose AFE text differs from the workbook (whitespace, case and number spelling aside),
  and lists names AFE's modules define that the workbook does not have. AFE 1.0's code sheet and AFE's
  locale-detection sheet are reported as info and left in place.
- AFE writes its modules with your locale's separators. In a `;` locale xln cannot
  compare them, and C14 says *not compared*.

The format is not documented by Microsoft. xln's knowledge of it was measured in AFE's
own add-in code and in one AFE-shaped file from a third party, not yet in a workbook saved
by Excel with AFE on our machines ([probes/README.md](../probes/README.md#afe-saved-workbooks-research-2026-10-07-no-clean-afe-save-yet)).
If xln misreads your file, C14 reports the part as unreadable, and the build still copies
it unchanged.

### If you keep using AFE alongside

After a build that changes names AFE's modules also define, the names exist in two
versions: the new one in the Name Manager, the old one in AFE's copy. The build warns
and lists them.

- **Do not save AFE's modules before bringing them in line.** AFE shows its own text when
  it opens, and saving its modules writes that text back over the names, undoing the
  build. Edit the module in AFE to match (the C14 findings show both texts), then save
  it, and run `xln check` again.
- **Edit each name in one place.** A name you edit in xln, keep editing in xln; a name you
  prefer to edit in AFE, edit there, save, and pull.
- **If you stop using AFE**, there is nothing to do: AFE's copy is inert while AFE is
  closed. xln does not remove it.

`xln apply` (for scripts and agents) says nothing about AFE's copy: decided, no warning
there (2026-10-08). `xln check` lists the differences.

Not yet measured: what AFE shows when it opens a workbook right after an xln build.

### Bringing AFE module text you keep elsewhere

To bring a module you keep outside the workbook (a gist, a text file), paste it into
`names/M.xln` and add the prefix to each name and to each call of a sibling
(`add_one` → `M.add_one`). The checker warns on every name still without the prefix.
Then build. For functions shared across workbooks, a library of `.lambda` files
([LANGUAGE.md §12](LANGUAGE.md#12-the-library)) gives you status, take and publish.

## 6. Sources

- The issues of [microsoft/advanced-formula-environment](https://github.com/microsoft/advanced-formula-environment/issues)
  and [microsoft/Excel-Labs](https://github.com/microsoft/Excel-Labs/issues), read on
  2026-10-07.
- AFE's [CHANGELOG](https://github.com/microsoft/Excel-Labs/blob/main/advanced-formula-environment/CHANGELOG.md)
  and [module-import.md](https://github.com/microsoft/Excel-Labs/blob/main/advanced-formula-environment/documentation/module-import.md).
- For xln: the [user guide](USER-GUIDE.md) (§13 for AFE workbooks), the
  [language specification](LANGUAGE.md) and the measured facts in
  [probes/README.md](../probes/README.md).
