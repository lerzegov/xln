# xln — Excel names as code

xln shows an Excel workbook's defined names (line items, constants, LAMBDA functions)
and the formulas of its named cells as text in VS Code. You read, search, check and edit
them there, and xln writes the edits back into the `.xlsx`. It reads and writes the file
itself: Excel does not need to be running, and nothing is installed in Excel.

It runs in desktop VS Code and in [vscode.dev](https://vscode.dev) (Edge or Chrome, on a
local folder), so it works on lab PCs where nothing can be installed.

## Quick start

1. **Pull.** Open the folder that holds your workbook. Right-click `model.xlsx` in the
   Explorer → **xln: Pull workbook**. xln writes `model.xln/` beside it: one `.xln` file
   per sheet under `names/sheets/`, and module files for LAMBDAs and other names.
2. **Edit.** Write formulas after the `=` of each name, as in Excel's formula bar:

   ```
   Sales @C3# = SEQUENCE(1, 5, 20000, 3400);
   COGS @C4# = Sales * 0.6;
   Gross_income @C5# = Sales - COGS;
   ```

   Completion offers names in scope and Excel functions; signature help shows LAMBDA
   parameters and their doc comments.
3. **Check.** Mistakes are underlined as you type: unknown names and functions, wrong
   argument counts, another sheet's local name read without its sheet. **xln: Audit
   workbook** reports checks C1–C15 on the whole workbook: functions stored without the
   prefix Excel needs (they show `#NAME?`), broken references, unused names, fixed
   references into spills, copy drift, cycles, labels that no longer match a renamed name.
4. **Build.** Right-click the `.xln` folder → **xln: Build workbook**. A dialog lists
   what will change; confirm, then open the workbook in Excel. On the desktop, **xln:
   Build and reopen in Excel** closes, writes and reopens it in one step.

Change layout and labels in Excel, save, and pull again; write formulas in xln and build.
Edit one side at a time.

Also: a formula view of each sheet (names on cells, in order of appearance or of
calculation), cell usages of a name, name search, Rename Symbol (F2) that also renames
the name in the workbook's cell formulas on the next build, and a shared LAMBDA library
with provenance (`@from`).

## What xln never does

- **No add-in, no macros.** Nothing runs inside Excel.
- **No hidden transformations.** A pull writes the names as the workbook has them; a
  build writes only what you changed, and the dialog lists every change first. xln
  never renames a name or changes its scope on its own: the checker points at a problem,
  you make the edit.
- **No silent overwrite.** A desktop build stops while Excel has the file open, keeps
  the previous file as `model.backup.xlsx`, reads the result back and restores the
  original on a mismatch. A name or cell changed in both Excel and xln stops the build.
  In the browser the build writes a new file, `model.xln.xlsx`, and leaves yours alone.
- **No cells or layout.** xln does not add rows, move cells or write labels: Excel owns
  those.

## Documentation

- [User guide](../../docs/USER-GUIDE.md): install, the pull–edit–build loop on a small
  income statement, spills and `#`, scope, modules and LAMBDAs, the library, common
  messages, vscode.dev.
- [The `.xln` language](../../docs/LANGUAGE.md): the precise rules of the text format.
- [For AFE (Excel Labs) users](../../docs/AFE-USERS.md): the requests AFE users made,
  what xln does about each, and opening an AFE workbook with xln.
- [Changelog](CHANGELOG.md).

## Requirements

VS Code 1.110 or later, desktop (macOS, Windows) or vscode.dev in Edge or Chrome.
Workbooks in `.xlsx` or `.xlsm` format.

## Licence

MIT, © 2026 Luca Erzegovesi.
