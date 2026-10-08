# P0 probes: can the Name Manager be driven without an add-in?

The same tests (IDs `T01`–`T19`) are run on three channels so the results line up:

| Channel | File | How to run | Status |
|---|---|---|---|
| macOS: AppleScript → Excel | `mac/probe_names.applescript` | `bash probes/mac/run_mac.sh` | **Run 2026-10-03** on Excel 16.115, macOS 26.6.2, `it_IT`; see `results/` |
| Windows: PowerShell → Excel COM | `windows/probe_names.ps1` | double-click `windows/run_windows.cmd` | **Run 2026-10-04** on M365 2609, Windows 11, `it-IT` (university licence) |
| Office Scripts (web, Windows, Mac) | `officescripts/probe_names.ts` | blank OneDrive workbook → Automate → New Script → paste → Run | to run on the UniTN tenant |
| Windows: M3 check (files xln wrote, F8 files, `--reopen` script) | `windows/m3_check.ps1`, kit in `windows/m3kit/` | double-click `windows/run_m3_check.cmd` (`windows/M3-CHECKS.md`); then on the Mac `node probes/windows/check_m3_winsaved.mjs` | kit ready 2026-10-05, passes in Mac Excel (`mac_open_kit.mjs`); to run on Windows |

Each probe opens a **new blank workbook**, saves a copy so we can read the stored form
in `xl/workbook.xml`, closes it unsaved, and leaves your other workbooks alone. The
Windows probe starts its own Excel instance. Reports go to `results/`.

## Test list

| ID | Question |
|---|---|
| E01–E04 | (Windows) PowerShell language mode, execution policy, can the Excel COM object be created at all? |
| T01, T02 | Create a constant and a LAMBDA name; call the LAMBDA from a cell |
| T03 | The stored text, in English (`RefersTo`) and in the local language (`RefersToLocal`) |
| T04 | Set and read the Name Manager **comment** |
| T05 | **Sheet-scoped** names: (a) created on the sheet object, (b) created as `S2!Name`, (c) listed per sheet |
| T06 | A name over a spilled range (`S1!$B$1#`) and its shape |
| T07 | Enumerate all names (including Excel's hidden `_xlfn.`/`_xlpm.` names) |
| T08–T10 | Update, rename (do dependent formulas follow?), delete |
| T11 | Syntax error: what message does the caller get? |
| T12 | A name that collides with a built-in (`Fact` vs `FACT`) |
| T13 | Dotted, AFE-module-style name `Mod.Fn` |
| T14 | Line breaks inside a definition |
| T15 | Length limit (7,992 / 8,192 / 9,012 characters) |
| T16 | Evaluate an expression (`ROWS(P_Spill)`) without writing a cell |
| T17 | Non-ASCII name `Growλ` |
| T18 | Write a definition in the local language (`;` separator) |
| T19, T20 | Save a copy; (Windows) is the open file locked? |
| T22 | (Windows) Four ways to create a name with an English definition |
| X01 | Stored form in `xl/workbook.xml` |
| F01 | (Office Scripts) `fetch` from raw GitHub |

## macOS results (2026-10-03)

| ID | Result | Consequence for the tool |
|---|---|---|
| T01, T02 | ✔ `make new named item at wb with properties {name:…, references:…}` | Live write works |
| T03 | `references` and `reference local` **both** return English with `,`, even on an `it_IT` system whose decimal separator is `,` | No locale translation layer needed on the Mac |
| T04 | ✘ **No comment property**: "Non posso impostare comment of named item" | Doc comments cannot reach the Name Manager over AppleScript. Options: keep docs only in source; write the `comment` attribute at file level; or use Office Scripts (`setComment`) |
| T05 | ✘ `make … at sheet` ("Errore nei parametri"); ✔ `name:"S2!P_Local2"` at the workbook creates a sheet-scoped name (stored `localSheetId="1"`); ✘ `named items of sheet` cannot be read | Sheet scope works through the **name prefix**. Enumerate at the workbook level, where sheet names appear as `'S2'!P_Local2` |
| T06 | ✔ Excel qualifies on entry: `=S1!$B$1#` → `='S1'!$B$1#`, stored `_xlfn.ANCHORARRAY('S1'!$B$1)` | Read back after every write and store Excel's normalised form in the lockfile |
| T07 | ✔ `named item i of wb` works (the `every name` failure in SUMMARY D5 was a terminology clash); hidden `_xlfn.*`/`_xlpm.*` names are listed with `visible=false` | Inbound sync is possible on the Mac. Filter `_xl*` names |
| T08–T10 | ✔ update, ✔ rename (the cell formula follows: `=P_K` → `=P_K2`), ✔ delete (dependents become `#NAME?`) | Full CRUD. Rename should be its own operation, not delete + add |
| T11 | Rejected with the generic "Errore nei parametri" (localised, no position) | Diagnostics must come from our own parser. Excel only says yes or no |
| T12 | ⚠ `Fact` is **accepted**, but `=Fact(5)` becomes `=FACT(5)` = 120: the built-in wins silently | The compiler must refuse names that collide with built-in functions |
| T13 | ✔ `Mod.Fn` works | AFE module naming carries over |
| T14 | ✔ Newlines are kept, in the live object and in the file | Formatted source can round-trip, apart from small whitespace changes (see below) |
| T15 | 7,992 characters accepted; 8,192 rejected | Limit is just under 8,192. The compiler should check it |
| T16 | ✔ `evaluate name "ROWS(P_Spill)*10+COLUMNS(P_Spill)"` → 15 | Shape checks need no scratch cells |
| T17 | ✔ `Growλ` | Unicode names fine |
| T18 | ✘ `reference local` with `;` rejected | Consistent with T03: always write English |
| T19 | ✔ only with a **POSIX path and no file format** (HFS path → "Errore nei parametri"); the workbook is renamed on save, so close it by its new name | Noted in the script |

**Whitespace is not preserved exactly.** `LAMBDA(b, g, b*(1+g))` was stored as
`LAMBDA(b,g, b*(1+g))`. Sync must compare definitions **modulo whitespace** and keep the
author's formatting in the source file, never in Excel.

## File-level route (code-only): `filelevel/probe_filelevel.py`

This probe patches `results/probe_mac.xlsx` **as a zip, without Excel**, into
`results/probe_patched.xlsx`, then lets Excel open it once. Mac results (2026-10-03):
no repair dialog; stale cached values recomputed on load (`fullCalcOnLoad`, which Excel
drops on save); the `comment` attribute survives; `localSheetId` scope works; a custom
XML part carrying the module source survives a re-save; a name written without its
`_xlfn.` prefix becomes `#NAME?` and is permanently rewritten as `_xludf.`; the `~$`
lock file is present while the workbook is open.

**On Windows** the same checks are automated in `windows/probe_names.ps1` (v4, tests
F00–F07): it patches the copy it has just saved, opens it in a second Excel instance,
reads the cells, re-saves it and inspects the result.

A manual check on 2026-10-04 (`probe_patched.xlsx` and `probe_mac.xlsx` brought over from
the Mac) found:
- In **Protected View**, every cell that calls a LAMBDA name shows `#NAME?`; only the
  built-in `=FACT(5)` shows its value.
- After *Enable Editing* [*Abilita modifica*], every value is correct, including
  `A1` = 141. So the patched file recalculates on load on Windows too.
- Users who receive a workbook by mail or download will see this in Protected View;
  worth a line in the docs.

## Windows results (2026-10-04)

Microsoft 365 for enterprise, version 2609 build 16.0.20430.20032, Windows 11
(NT 10.0.26200), PowerShell 5.1, `it-IT`. The author's laptop with the university
Microsoft 365 licence: E02–E04 pass (FullLanguage, no Group Policy execution policy, COM
works). Final run:
`results/win-LAPTOP-TJ8P6I38-20261004-174721.txt` (script v3). The two earlier reports
show the trap below.

| ID | Result | Consequence for the tool |
|---|---|---|
| T22 | ✘ `Names.Add(name, refersTo)` from PowerShell parses `refersTo` in the **local** language: `=LAMBDA(x, x+1)` is rejected on `it-IT`. Forcing the thread culture or calling through `InvokeMember` with en-US does not help. ✔ Creating with `=0`, then setting the `.RefersTo` **property**, takes English. `.Formula2` also takes English | Any live Windows backend must create, then assign. No effect on `xln`, which works at file level |
| T01, T02 | ✔ (via T22 route b) | |
| T03 | `RefersTo` is English (`,`); `RefersToLocal` is **localised** (`LAMBDA(x; x+1)`). The Mac returns English for both | Read `RefersTo` only |
| T04 | ✔ **The comment can be set and read** over COM, stored as the `comment` attribute. The Mac cannot do this | |
| T05 | ✔ both ways: on the sheet's `Names`, and `S2!P_Local2` at the workbook; ✔ listed per sheet. Stored `localSheetId="1"` | |
| T06 | ✔ as Mac: `='S1'!$B$1#`, stored `_xlfn.ANCHORARRAY('S1'!$B$1)` | |
| T07 | ✔ hidden `_xlfn.*`/`_xlpm.*` helpers listed with `visible=False` and `=#NAME?`, but **not** in the saved `workbook.xml` | Filter `_xl*`, as on the Mac |
| T08–T10 | ✔ update, ✔ rename (cell follows: `=P_K2`), ✔ delete (`#NOME?`) | Same as Mac |
| T11 | Rejected, with a more specific but still position-free message ("manca una parentesi") | Diagnostics from our own parser |
| T12 | ⚠ as Mac: `Fact` accepted, `=Fact(5)` becomes `=FACT(5)` = 120 | Check C3 confirmed on both systems |
| T13, T14, T17 | ✔ `Mod.Fn`, newlines kept, `Growλ`. Whitespace normalised as on the Mac: `LAMBDA(b, g, …)` → `LAMBDA(b,g, …)` | |
| T15 | 7,992 accepted; 8,192 rejected ("troppi valori, riferimenti… nomi"); 9,012 rejected | Same limit as Mac |
| T16 | ✔ `Evaluate` works | |
| T18 | ✔ `RefersToLocal` with `;` accepted (the Mac rejects it) | |
| T20 | The open workbook file itself is **locked** (sharing violation) | A write fails anyway while Excel has it open; keep the `~$` check (E1) for the Mac and for a clear message |

**Line breaks are stored as raw CR LF on both systems.** The script sent LF only;
`results/probe_win.xlsx` stores `_xlfn.LAMBDA(_xlpm.x,\r\n  _xlpm.x*2)`, byte for byte as
`results/probe_mac.xlsx`. A conforming XML parser turns CR LF into LF on read, so the
decompiler sees LF. When `xln` writes a multi-line definition, it should emit CR LF, as
Excel does.

**File-level route on Windows (script v4, `results/win-LAPTOP-TJ8P6I38-20261004-181706.txt`):
identical to the Mac.**

| ID | Result |
|---|---|
| F01 | Opens with no repair; `~$probe_patched_win.xlsx` present while open |
| F02 | `A1` = **141** on open (cached 42): `fullCalcOnLoad` works; all other cells correct |
| F03 | `Z_Add2(1)` = 3; the patched `comment` attribute is read by Excel and kept on save |
| F04 | `Z_Loc` scoped to S2 by `localSheetId="1"` (sheet position) |
| F05, F06 | `Z_Bare` without `_xlfn.` gives `#NOME?` and is re-saved as `_xludf.SEQUENCE(1,3)`: the poisoning trap (C2) is the same on Windows |
| F06 | Excel drops `fullCalcOnLoad` on save (`<calcPr calcId="191029"/>`) |
| F07 | The custom XML part with the module source survives a Windows save |

Risk 8.2 of the brief ("Windows differs from the Mac at file level") is retired for
everything tested here.

## F7: rename a name inside cell formulas at file level (2026-10-04, macOS)

`filelevel/probe_f7_rename.py`. Excel builds a workbook full of rename traps; the same
renames are done twice, **live by Excel** (the oracle: the Name Manager rewrites every
dependent) and **by patching the zip**, token by token. Then Excel opens the patched
file, recalculates on load and re-saves it. Report: `results/f7-mac-20261004-182314.txt`;
files `results/f7_*.xlsx`.

Renames: `Rate` → `Growth` (workbook), `Fn` → `Fx` (workbook LAMBDA), `S2!Loc` → `Spot`
(sheet scope, with a workbook-level `Loc` that must not change).

| Check | Result |
|---|---|
| Formulas, patched vs Excel's own rename: 28 cells, 7 names, 1 CF rule, 1 DV rule | **0 differences** |
| Cached values after Excel recalculated the patched file vs before the rename | **0 changed** (28 cells) |
| Formulas Excel rewrote when re-saving the patched file | **0** |
| Shared formulas (13 cells in 2 groups) | Only the master carries text: renaming it renames the group. Children (`<f t="shared" si="0"/>`) untouched |
| Spill anchor (`cm="1"`, `t="array" ref="E1:E3"`) | Token rename only; metadata untouched and accepted |
| `calcChain.xml` | Left as is; it lists cells, not formula text. Accepted, and still present after re-save |

Traps that both Excel and the patch handle the same way:
- `RateX` is not touched;
- `"Rate is "&Rate` changes the name but not the string;
- `LET(Rate, …)` is stored as `_xlpm.Rate` and not touched;
- `rate*3` is renamed regardless of case;
- `Loc` on S1 (the workbook-level name) is kept, while `Loc` on S2 and `'S2'!Loc` everywhere become `Spot`;
- the name inside `Rate2`'s and `Fx`'s definitions is renamed;
- the name is renamed in conditional formatting and data validation formulas;
- `INDIRECT("Rate")` is not renamed by Excel either (here it was `#REF!` already, since
  `Rate` is a constant). The audit should flag names that appear inside strings.

**Lesson from the first run (a bug in the probe, now fixed).** The patcher's pattern for `<f>…</f>`
also matched the self-closing shared children and swallowed the cells up to the next `</f>`.
Excel showed its **recovery dialog** ("recover the contents?"), the author accepted, and
Excel saved sheet S1 without those cells. The comparison also caught it. So:
- the patcher must use a real XML tokenizer, not a regular expression;
- a broken patch does not fail quietly on the Mac: the user sees the recovery dialog.
  But read-back (E3) must catch damage *before* the user opens the file, and `verify`
  (E4) must catch whatever a repair throws away.

**Verdict:** F7 passes. Stretch G (rename across cells) is feasible as a token substitution.
Not yet covered: names used in charts (`xl/charts/*.xml`), pivot tables, Table column
formulas (`xl/tables/*.xml` `calculatedColumnFormula`), and `xl/externalLinks`. G must
either rename in those too, or refuse when the name appears there.
The file-level behaviour on Windows matched the Mac in F00–F07, so a Windows rerun of F7 is
optional.

## F8: replace cell formulas at file level (2026-10-05, macOS)

`filelevel/probe_f8_cells.py` (run with `/opt/homebrew/bin/python3`). Excel builds
`results/f8_base.xlsx` (sheets N, D, Slot, Sh) and `results/f8_plain_base.xlsx` (no
formulas at all, so no `metadata.xml` and no `calcChain.xml`). The same edits are made
**live by Excel** (`formula2`; the oracle, `f8_oracle.xlsx`) and **by patching the zip**
with a real XML parser (`xml.dom.minidom`). Excel opens each patched file; the script
reads `formula2` and the displayed text of 86 cells before anything else happens, then
re-saves. Report: `results/f8-mac-20261005-091726.txt`. Patched and re-saved files:
`f8_p1*.xlsx`, `f8_q1*.xlsx`; the stale-calcChain file that needs repair: `f8_p2.xlsx`;
Excel's repair logs: `f8-repair-p2.xml`, `f8-repair-p2b.xml`.

Workbook-wide variants (cases 6 and 7), the same cell edits in each:

| Variant | calcChain | Edited cells' `<v>` | `fullCalcOnLoad` | calcId | Opens | Values on open |
|---|---|---|---|---|---|---|
| **p1** | dropped | none | `1` | as saved (181029) | ✔ | **= oracle** (except the deliberate traps below) |
| p3 | dropped | stale kept | `1` | 181029 | ✔ | = p1 |
| p4 | dropped | none | — | 181029 | ✔ | = p1 (Excel recalculates anyway: older calcId) |
| p5 | dropped | stale kept | — | 181029 | ✔ | = p1 (same reason) |
| p6 | dropped | stale kept | — | **191029** (current) | ✔ | ✘ 45 cells wrong: no recalculation, stale values shown |
| p7 | dropped | none | — | 191029 | ✔ | ✘ 39 wrong: plain formulas without `<v>` are computed, array formulas stay blank, dependents keep stale values |
| p2 | **stale, kept** | none | `1` | 181029 | ✘ **REPAIR** | (after *Yes*: values right) |
| p2b | stale minus entries of cells that lost their formula | none | `1` | 181029 | ✘ **REPAIR** | |
| p2c | p2b minus entries whose `a="1"` flag changed | | | | not run (each repair needs a human click) | |

Repair log, both p2 and p2b: `Removed Records: Formula from /xl/calcChain.xml part
(Calculation properties)`. A new workbook saved through AppleScript carries
`calcId="181029"`; every re-save writes `191029`. Real workbooks have the current calcId,
so **p6/p7 are the realistic cases: `fullCalcOnLoad="1"` is required.**

Cell cases (variant p1; values compared with the oracle):

| Case | Patch | Opens, values | Excel re-saves |
|---|---|---|---|
| 1 normal cell | `<f>SUM(A1:A3)*100</f>`, `<v>` removed | ✔ 600 | same text, `<v>` added |
| 1 LET in a cell | `_xlfn.LET(_xlpm.x,A1*10,_xlpm.x+1)` | ✔ 11 | same |
| 1 trap: array-evaluating formula written plain | `<f>SUM(A1:A3*2)</f>` (no `cm`) | ✘ `#VALUE!`; `formula2` shows `=SUM(@A1:A3*2)` | kept as legacy, `t="e"` |
| 1 same formula in dynamic-array form | `cm="1"` + `<f t="array" ref="B5">` | ✔ 12 (Excel's own form) | same |
| 1 scalar in dynamic-array form | `cm="1"` + `<f t="array" ref="B6">1+2</f>` | ✔ 3, `formula2` `=1+2`, no braces | kept in array form (Excel itself writes `1+2` plain) |
| 1 trap: LAMBDA call written plain | `<f>_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)(A3)</f>` | value 6, but `formula2` shows `=@LAMBDA(x,@x*2)(A3)`; Excel stores it with `cm="1"` | kept plain |
| 1 clear a formula | `<f>`, `<v>`, `t`, `cm` removed (empty `<c r>` left) | ✔ | element dropped |
| 2 bigger spill (`SEQUENCE(3)`→`(5)`) with names on `C6#` | `ref="C6"` (anchor only), old spill's ghost cells deleted | ✔ C6:C10, `SUM(Spill)`=15, `ROWS(C6#)`=5 | `ref="C6:C10"`, `cm="1"`, ghosts written back |
| 2 smaller spill (4→2) | same | ✔ F6:F7, `SUM(Spill2)`=3 | `ref="F6:F7"` |
| 2 old `ref` and ghost cells kept (bigger and smaller) | text only | ✔ Excel treats the old range as the spill and resizes it | correct ref, stale ghosts removed |
| 2 trap: `ref` reduced to the anchor but ghosts kept | | ✘ **`#SPILL!`**: the ghosts are now constants in the way | |
| 2 spill → scalar | `<f>42</f>`, `cm` and `t="array"` removed, ghosts deleted | ✔ | |
| 2 scalar → spill | `cm="1"`, `<f t="array" ref="P6">` | ✔ P6:P8 | |
| 3 slot (Create from Selection, no `<c>` in the file) → spill | new `<c r="B1" cm="1"><f t="array" ref="B1">…` inserted in column order | ✔ B1:D1 | |
| 3 slot → scalar; slot in a row absent from the file | new `<row r="9">` inserted in row order, no `spans` | ✔ | |
| 3 slot name moved to the spill | `<definedName>` `Slot!$B$1` → `_xlfn.ANCHORARRAY(Slot!$B$1)` | ✔ `SUM(Revenue)`=60 | same |
| 4 number cell → formula | `<f>`, `<v>` removed | ✔ | |
| 4 shared-string cell → formula (text result) | `t="s"` and `<v>` removed; the string stays orphaned in `sharedStrings.xml` | ✔ `AB1` | `t="str"`; string pruned |
| 4 text cell → spill | `cm="1"`, `t="array"` | ✔ | |
| 5 edit a shared child | child gets its own `<f>`; master and its `ref` untouched | ✔ | identical to the oracle |
| 5 edit a master, promote the next child | C2 gets `<f t="shared" ref="C2:C5" si="1">A2+1</f>` (text **translated** to C2) | ✔ | Excel's own: promotes C2 but keeps `ref="C1:C5"`; both accepted |
| 5 edit a master, un-share the group | each child gets its own translated `<f>` | ✔ | kept un-shared (Excel's own keeps a group) |
| 5 edit the last child, shrink the master's `ref` | `ref` E1:E5 → E1:E4 | ✔ | kept; Excel's own keeps E1:E5 |
| 7 dynamic array in a workbook without `metadata.xml` (q1) | create `xl/metadata.xml` (XLDAPR record), its rel and content-type override | ✔ = oracle | same `cellMetadata` as Excel writes |
| 7 trap: `t="array"` without `cm` (q2) | | ✘ a legacy CSE array: B1 only, no spill, `SUM(B1#)`=1 | |

Also measured: a stale `<dimension>` (`A1:A6` on a sheet now using F9) and stale row
`spans` are accepted and fixed on save; Excel renumbers `si` on save; it drops
`fullCalcOnLoad` on save and rebuilds `calcChain.xml`.

**Recipe for M3b (file backend, cell formulas):**
- Write the stored form: `_xlfn.`/`_xlfn._xlws.` prefixes, `_xlpm.` parameters,
  `_xlfn.ANCHORARRAY(x)` for `x#`, XML-escaped. Parse and serialise with an XML parser.
- Write **every** replaced formula in dynamic-array form: `<c … cm="N"><f t="array"
  ref="<anchor>">text</f></c>`, where N is the 1-based `cellMetadata` record pointing at
  the `XLDAPR` (`fDynamic="1"`) type. Plain `<f>` has legacy (implicit-intersection)
  semantics, and `t="array"` without `cm` is a legacy CSE array. (Excel writes plain only
  when the two coincide; xln cannot tell without type analysis, and the array form is
  accepted for scalars.)
- If `xl/metadata.xml` is absent, create it with the XLDAPR record (`metadataTypes`,
  `futureMetadata` with `xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"`,
  `cellMetadata` with one `<bk><rc t="1" v="0"/></bk>`), plus its relationship
  (`…/relationships/sheetMetadata`) and the content-type override.
- Remove the cell's `<v>`, `<is>` and `t`; keep `s` (style) and every other attribute.
- Spill anchors: `ref` = the anchor alone, and **delete the ghost cells** of the old
  `ref` (all but the anchor). Never leave ghosts behind a reduced `ref`. The new spill
  size need not (and cannot) be known.
- Slots: insert `<row>` and `<c>` in ascending order. `dimension` and `spans` need not
  be updated. A name on a slot that now spills may be redefined to `ANCHORARRAY`.
- Clearing a formula: remove `<f>`, `<v>`, `t`, `cm`.
- Shared formulas: editing a child = give it its own formula, master untouched. Editing a
  master = promote the next member (`t="shared"`, same `si`, a `ref` covering the rest,
  the master text **translated** to the new cell) or un-share the whole group (every
  member gets its translated text). Both need a relative-reference shifter on the W2
  tokenizer; un-sharing is the simpler rule.
- **Delete `xl/calcChain.xml`**, its relationship and its content-type override,
  whenever any cell formula changes. A stale chain triggers the repair prompt, even
  pruned of the cells that lost their formula.
- Set `<calcPr fullCalcOnLoad="1">`. Without it Excel shows stale values (or blanks for
  array formulas) when the file carries the current calcId.

**Traps for the tool and for the probes:**
- The repair prompt does not always block AppleScript: in one run `open workbook`
  returned with no workbook open, the script went on sending events, `quit` was refused
  ("cancelled by user") and Excel **crashed** (Microsoft Error Reporting). The probe now
  checks that a workbook of the expected name opened, times out after 60–90 s, and
  opens the repair variants only with `--with-repair`.
- Repairs are visible afterwards only if someone clicks *Yes*: Excel then writes
  `~/Library/Containers/com.microsoft.Excel/Data/tmp/Repair Result to <file><n>.xml`.
  `--reopen` or E3 could read it. System Events cannot press *No* without Accessibility
  rights for the terminal, which this machine does not grant.
- macOS prompted the author to install the Xcode command-line tools during this session.
  The probe itself runs only `osascript` and a Python interpreter; the likely trigger is
  a one-off `sdef` call (reading Excel's AppleScript dictionary) made while writing it.
  Run the probe with `/opt/homebrew/bin/python3`, never `/usr/bin/python3`.

**Verdict:** F8 passes. Cell formulas can be replaced at file level, including spills,
slots, value and text cells and shared groups, with no repair and the right values on
open, provided the recipe above is followed. **Windows:** rerun the same files
(`f8_p1.xlsx`, `f8_q1.xlsx`, `f8_p2.xlsx`) through the probe script to confirm (open,
values, repair on `f8_p2`).

## F9: LAMBDA stored forms (2026-10-06, macOS)

**Why:** our compiler wrote an optional parameter as `[_xlpm.periods]`; Excel repaired the
file ("Removed Records: Named range from /xl/workbook.xml") and dropped the name, and our
read-back passed because our decompiler read our own form back. Script
`filelevel/probe_f9_lambda.py`: Excel (Mac, it-IT) makes a new workbook, enters 13 names
(`make new named item … references:"=…"`) and 27 cell formulas (`formula2`), calculates
and saves `results/f9_lambda_mac.xlsx`; output in `results/f9-mac-20261006.txt`. Also
`results/lambda_optional_mac.xlsx` (the author's `ANA.GROW`, saved by Excel).

| Typed | Stored |
|---|---|
| `LAMBDA(value, rate, [periods], … ISOMITTED(periods) …)` | `_xlfn.LAMBDA(_xlpm.value,_xlpm.rate,_xlop.periods, … _xlfn.ISOMITTED(_xlpm.periods) …)` |
| `LAMBDA([x], [y], …)` (only optional) | `_xlfn.LAMBDA(_xlop.x,_xlop.y, …)` |
| `LAMBDA(x, [k], LET(m, …, LAMBDA(y, [z], …)))` | `_xlfn.LAMBDA(_xlpm.x,_xlop.k, _xlfn.LET(_xlpm.m, …, _xlfn.LAMBDA(_xlpm.y,_xlop.z, …)))` |
| `LAMBDA(x, LET(y, x * 2, z, y + 1, y * z))` | `_xlfn.LAMBDA(_xlpm.x, _xlfn.LET(_xlpm.y, _xlpm.x * 2, _xlpm.z, …))` |
| `LET(f, LAMBDA(t, t * 3), f(2))` | `_xlfn.LET(_xlpm.f, _xlfn.LAMBDA(_xlpm.t, _xlpm.t * 3), _xlpm.f(2))` |
| `LAMBDA(x, x + 1)(2)`, `MAKEADDER(2)(3)` | `_xlfn.LAMBDA(_xlpm.x, _xlpm.x + 1)(2)`, `MAKEADDER(2)(3)` |
| `MAP(A1:A3, ABS)`, `BYROW(A1:A3, SUM)` | `_xlfn.MAP(A1:A3, _xleta.ABS)`, `_xlfn.BYROW(A1:A3, _xleta.SUM)` (names and cells) |
| `LAMBDA(n, IF(n <= 1, 1, n * FACT(n - 1)))` as `FACT` | the recursive call stays `FACT(…)`, no prefix |
| `ANA.GROW(100, 0.1, 2)` (a name with a dot) | `ANA.GROW(100, 0.1, 2)` |

- **An optional parameter is `_xlop.name` in the parameter list, without brackets; every use
  in the body is `_xlpm.name`.** Same in cells, in nested LAMBDAs and in a LAMBDA called
  immediately (`LAMBDA(_xlop.p, …)()`). `[_xlpm.p]` is not a stored form: Excel drops the
  name on open.
- Excel removes the spaces between parameters (`_xlpm.a,_xlop.b,_xlop.d, body`) and keeps
  the rest of the author's spacing.
- Values: `ANA.GROW(100,0.1,2)` = 121, `ANA.GROW(100,0.1)` = 110, `ALLOPT()` = 2,
  `ALLOPT(, 7)` = 7 (an empty argument counts as omitted), `FACT(5)` = 120,
  `OPTNEST(3, 4)(1, 5)` = 18.
- Cells holding these formulas are stored in dynamic-array form (`cm="1"`,
  `t="array"`) except those that cannot return an array (`FACT(5)`, `SUM(ETAMAP)`): plain `<f>`.
- A name that reads as a cell address (`OPT2`: column OPT, row 2) is refused by the Name
  Manager ("Errore nei parametri").

**Check of the fix (2026-10-06, Mac):** a copy of `f7_base.xlsx` built by xln with the *New
module* sample (`ANA.GROW`, stored `…,_xlop.periods,…`) and three cell changes on S1
(`ANA.GROW(100, 0.1, 2)`, `ANA.GROW(100, 0.1)`, `LAMBDA(x, [y], x + IF(ISOMITTED(y), 10, y))(1)`)
opens in Excel without a repair (no repair log) and computes 121, 110 and 11; the Name
Manager shows `=LAMBDA(value,rate,[periods], …)`.

## Sheet quoting in stored formulas (survey, 2026-10-07)

Read-only survey of every Excel-saved workbook at hand (the 19 `results/*.xlsx`,
`fixtures/traps.xlsx`, the 7 corpus `dist/*.xlsx`, the author's `is-model.xlsx` and
`lbo-play.xlsx`; `docProps/app.xml` says Microsoft Excel for all, Mac and Windows), counting
each sheet qualifier in front of `!` as quoted or bare. Cell formulas, which Excel writes
from its own parse on every save:

| Sheet | quoted | bare |
|---|---|---|
| `BS`, `IS`, `SCF`, `FCF` (look like columns) | 0 | 495 |
| `Model`, `Check`, `Primer`, `Traps`, `SalesCOGS`, `FinDebt`, `LastActuals` | 0 | 2,061 |
| `S1`, `S2` (look like cells) | 16 | 0 |
| `SCF recursive`, `0 Tables` … `8 Audit` (blank, leading digit) | 1,164 | 0 |

Defined names keep the text their writer stored (the Python build quoted `'BS'!`; names
made in Excel store `D!`, `P!`, `APV!`, `FTE!` bare). So Excel quotes a sheet name only
when it is not a plain word (letters, digits, `_`, `.`, starting with a letter or `_`) or
reads as a cell; a column-like name is bare. R1C1-like names (`R`, `C`, `R1C1`) and
`TRUE`/`FALSE` were not met; xln quotes them, as Excel's documentation does. xln writes
this rule everywhere (`sheetNeedsQuotes`, spec §14 issue 7).

## F5 (D5, D6): the embedded source and provenance tags through Excel's save (2026-10-05, macOS)

Excel 16.115 on the Mac, driven through AppleScript; copies of `lbo-ep03r.xlsx` (corpus,
`dist/`) in `~/Library/Group Containers/UBF8T346G9.Office/xln-d5/`. `xln pull`, edits in
`names/FN.xln` (`// @version 1.2`, a two-line doc comment) and `names/IN.xln` (a short
doc, a 240-character doc), `xln build` with the source embedded and module names tagged:
13 names files, lockfile and config in `customXml/item1.xml` (347 KB), 21 module names
tagged. Opened in Excel, saved, closed; pulled again; edited, built, opened, saved and
pulled a second time.

| Check | Result |
|---|---|
| Open the built file | ✔ no repair prompt (workbook of that name open within 1 s), both rounds |
| The custom XML part after Excel's save | ✔ **kept**: same part name (`customXml/item1.xml`), same `itemProps1.xml` (our itemID and schemaRef), same relationship and content type. Excel **re-serialises the item**: every LF in text becomes CR LF (9,142 lines: 347,211 → 355,974 bytes), the XML declaration loses its line break; `&lt;`/`&gt;` kept |
| Restore after the save (first build, LF stored as is) | ✘ every file failed its checksum because of the CR LF. Fixed: files are stored with LF and `eol="lf"\|"crlf"`, read back with CR LF → LF before `eol` is applied; mixed line ends go as base64. With the fix, the same Excel-saved file restores |
| Pull of the Excel-saved file | ✔ `names/**` (13 files), `xln.lock.json` and `xln.config.json` **byte-identical** to the project that was built; only the manifest differs (derived from the workbook) |
| Second build on the Excel-saved file | ✔ the part is found by its namespace and rewritten in place (still one `customXml/item*.xml`); Excel opens and saves it again; the second pull is again byte-identical; `xln verify`: 12,078 cells, 0 changed |
| Provenance tags in Excel | ✔ kept on save: `comment="Picks a value by key. [xln FN 1.2 #41bad1]"`. AppleScript cannot read a name's comment on the Mac (`comment of named item` is `missing value`, as T04), so the check is on the saved file |
| A line break in a comment | Written by xln as `&#10;`, accepted; **Excel saves it as `_x000a_`** (an ST_Xstring escape). The reader now decodes `_xHHHH_` in comments (and the writer escapes a literal `_xHHHH_` as `_x005F_xHHHH_`); before the fix the comment read as changed in Excel |
| Comment length | A name comment of **255** characters opens; **256** and 300: Excel refuses the file (`open workbook` fails with "Errore nei parametri", no workbook opened, no repair log). A build now refuses a comment over 255 (error `comment-length`, a line break counted as 2); a tag that would pass 255 is left out with a warning |
| Drift reported by pull after the save | 3 cells on `Check` (C726, D726, B727): Excel rewrote the literal `1E-14` as `0.00000000000001` when saving. Not xln's change, but the lockfile hashed numbers as written, so it showed as an edit in Excel. Fixed by lockfile format 3 (numbers hashed by value): the same file now pulls with no drift. `1E+99` and `1E+300` in the names `FN.DEV` and `FN.FIXPOINT` came back as written |

The first open of the first file, made with `open (POSIX file p)` instead of
`open workbook workbook file name p`, timed out after 90 s with no workbook open; the
same command form on a good file returned at once without opening it. That file also
had the 300-character comment, so it could not have opened either. Use the
`excelControl` form.

**Windows to check:** open a built file with the embedded part and tags, save, and pull
it again: does Windows Excel also rewrite the item's line ends (the reader copes either
way), keep `itemProps`, and write `_x000a_` in comments? Does it refuse a 256-character
comment the same way (and how: repair prompt or error)?

## M3 on Windows (2026-10-05, Windows 11, M365 build 20430, it-IT)

`windows/run_m3_check.cmd` on the kit `windows/m3kit/` (run from the Mac's `probes` folder
over SMB); report `results/m3-win-LAPTOP-TJ8P6I38-20261005-142920.txt`, Excel's saved
copies in `results/m3win/`, compared on the Mac with `windows/check_m3_winsaved.mjs`.

- **Part 1, the 7 clean kit files** (F8 p1 and q1, M3a names, M3b build and apply, D5/D6,
  a 255-character comment): every formula, value, name, comment and tag as expected, no
  prompt. Saved copies: the custom XML part is kept and its line ends become CR LF (as on
  the Mac); comment line breaks are stored as `_x000a_`; a pull restores the project with
  no drift. The only value that differs from Mac Excel's copy is `="Rate is "&Rate`:
  number-to-text follows the locale (`0,2` on it-IT), a property of the formula.
- **Files that need repair** (`c256_comment`, `f8_p2`), opened through COM with
  `Workbooks.Open(path)` and `DisplayAlerts` on: Excel raises error 1004 ("Impossibile
  trovare la proprietà Open per la classe Workbooks") within 0.3 s, **no prompt**. Under
  automation Windows Excel refuses such a file instead of offering repair. (The report's
  K08 PASS / K09 FAIL lines are labels for a prompt that never came.)
- **`--reopen` script:** state, close, discard and "opened from Explorer" work (R01, R03,
  R06–R11; Excel started from Explorer is reachable once the console has the focus).
  **Open failed on every file** (R02, R12, R15): the 15-argument
  `Workbooks.Open(…, [Type]::Missing …, CorruptLoad)` call raises error 1004 even on clean
  files. Fixed 2026-10-05 to the plain `Workbooks.Open($full)`.
- **Rerun of part 2** (`windows/run_m3_reopen.cmd`): opening worked, but a workbook
  opened from Explorer lives in **another Excel instance** than the one
  `GetActiveObject` returns, so state/close missed it (R08–R10). Fixed: the script also
  looks the workbook up in the Running Object Table by its path (any instance), and `open`
  does nothing when it is already open. **Second rerun
  (`results/m3-win-LAPTOP-TJ8P6I38-20261005-144201.txt`): 16 PASS on local files**:
  state with Excel not running, open (Excel started by the script, which stays open), state
  saved/unsaved, close without saving (file unchanged), Explorer-opened workbook, open of the
  rewritten file (the new file, with its embedded part), a file needing repair is not
  reported as opened.
- **OneDrive (R17, R18): not found.** Excel names a workbook opened from the OneDrive folder
  by its web address, not its path, so `--reopen` cannot find it; the run was also
  disturbed (Excel asked to sign in, could not save to OneDrive, the author saved a copy).
  Files open for co-authoring in OneDrive are out of scope (brief §3); the lock-file guard
  still keeps a build from writing while Excel has the file open. Keep xln workbooks in a
  local folder.

## AFE-saved workbooks (research 2026-10-07; no clean AFE save yet)

Microsoft's Advanced Formula Environment (AFE, in the Excel Labs add-in) keeps its own
copy of the names it manages, as module text, inside the workbook. Microsoft does not
document the format. What follows was **measured** in AFE's production JavaScript and in
one AFE-shaped file, **not** in a workbook saved by Excel with AFE on our machines. Each
line says which.

**Sources**

- *Bundle* (measured): `https://advancedformulaenvironment.officeapps.live.com/taskpane.b9a14c33c199e850186a.bundle.js`,
  the add-in's live task-pane code, fetched 2026-10-07 (10.6 MB, SHA-256 `65b80201…d3d656`).
  The host was found through the console log attached to
  [Excel-Labs #32](https://github.com/microsoft/Excel-Labs/issues/32); the file name is also in
  the Wayback index (snapshot 2025-11-14). The store is webpack module 93224, the sync
  manager 71706. Not kept in the repository (Microsoft's code).
- *Sample* (measured, with a caveat): `ozzit.xlsx` from
  [github.com/ryanduguid/Ozzit](https://github.com/ryanduguid/Ozzit) (SHA-256 `8c42e227…bb760`).
  It is **not** a plain Excel save: the repository's `tools/sync_afe_store.py` rewrites the
  blob and re-zips the file. Its part layout is AFE's; its module paths (`/projects/Dates`)
  do not match its names' prefix (`oz.…`). Not kept in the repository (third-party file).
- *Reported*: the AFE [CHANGELOG](https://github.com/microsoft/Excel-Labs/blob/main/advanced-formula-environment/CHANGELOG.md)
  and [FAQ](https://github.com/microsoft/Excel-Labs/blob/main/advanced-formula-environment/FAQ.md);
  AFE issues [#30](https://github.com/microsoft/advanced-formula-environment/issues/30) and
  [#41](https://github.com/microsoft/advanced-formula-environment/issues/41).

**The custom XML part (AFE 1.1 and later)**

| Fact | Source |
|---|---|
| Part `customXml/itemN.xml` with root `<AFEJSONBlob xmlns="http://schemas.advancedformulaenvironment.officeapps.live.com/afejsonblob/1.0">`, no XML declaration and no child elements | bundle (`NODE_NAME="AFEJSONBlob"`, the namespace literal), sample |
| Its text is base64 of the **UTF-16LE** bytes (no BOM) of `JSON.stringify(store)` | bundle (writer: `Uint16Array` of char codes → bytes → `btoa`), sample decoded |
| `store = { schema: ".../afeprojects/0.2", files: [{ path: "/projects/<Module>", text }], projectNames: [...], locale: {...} }`; schema 0.1 is upgraded to 0.2 on load | bundle, sample (7 modules, 142 `projectNames`, `localeName: "en-au"`) |
| Module text: plain text, LF line breaks, in AFE's syntax (`name = definition;`, `/** doc */`, `//` comments); written with the locale's separators (`locale.listSeparator`, `decimalSeparator`) | bundle (default Workbook module text; `targetLocaleSpec`), sample (en-au: `,` and `.`) |
| Module `Workbook` (`PrimaryModule`): names exported unprefixed; any other module `M`: names exported as `M.name` | bundle (reported by our research run from the export code; the sample does not follow it, see caveat) |
| `projectNames`: the names AFE exported to the Name Manager; the Names tab hides them; when AFE commits its modules it adds and updates names from the modules and deletes names that were in `projectNames` but are no longer in a module | bundle (`emitPatch`, `namePatch.added/deleted`) |
| `customXml/itemPropsN.xml`: `ds:itemID="{GUID}"`, one `ds:schemaRef` to the AFEJSONBlob namespace; `customXml/_rels/itemN.xml.rels` → itemProps; workbook relationship of type `customXml`; content-type override for itemProps only (the item relies on `Default Extension="xml"`) | sample |
| AFE finds its part **by ID, not namespace**: the ID is kept in the add-in's document settings under `projectV0_1-56c6e055-265e-4713-816e-a646dbb708de` = `{"kind":"AFEJSONBlobNode","id":"{GUID}"}`, which Excel saves in `xl/webextensions/webextension1.xml` (`we:reference id="wa200003696"`), equal to `ds:itemID` | bundle (`customXmlParts.getItemOrNullObject(id)`, never `getByNamespace`), sample |

**Hidden sheets**

| Fact | Source |
|---|---|
| AFE 1.0 stored modules in a very hidden sheet named `AFE_hidden_codesheet_49ddb8b8` (header row with `rows`/`columns`; per module `namespace`, `name`, `start`, `code_0…`; text split across cells, joined with LF). AFE 1.1 converts it to the custom XML part and does not delete the sheet | bundle (`hiddenCodeSheetId`, `readFilesFromSheet`); CHANGELOG ("AFE will not delete the hidden sheet automatically") |
| Locale detection: a sheet named `e00eb4de3c8a421cba9b8f4cb8546ec` (31 hex digits: a GUID cut to Excel's sheet-name limit), where AFE writes test formulas (`=SUM(,)`, `={1,2;3,4}`) and reads them back localised. The current bundle creates it **hidden** and **deletes it** after use; the FAQ still says "very hidden" and that it can be deleted; older builds may have left it. With a protected workbook AFE uses an existing sheet `afe_scratch_sheet` instead. A temporary hidden name `aa5483da3aa04530a4c60d5c5971cb1b` is added and deleted the same way | bundle; FAQ; #30 (a user asked to keep the fixed name) |

**What xln does with it** (`packages/core/src/file/afe.ts`, `audit/afe.ts`): recognises a
custom XML item whose root is in the `http://schemas.advancedformulaenvironment.officeapps.live.com/`
namespace family (reads `AFEJSONBlob` 1.0, reports a later one as unreadable) and the two
sheet names above (not `afe_scratch_sheet`: a user's sheet); never writes any of them. The
build copies these parts byte for byte (tested, `test/file/afe.test.ts`). On the sample,
140 of 142 AFE entries matched the workbook's names modulo whitespace once compared on
display forms; the other two differed only in the case of an error literal (`#Value!`
in AFE's text, `#VALUE!` stored), which the comparison now ignores: 142 of 142. The sample
needed the `projectNames` list to pair entries with names (its prefix is not the module's).

**Synthetic fixtures** (`probes/fixtures/make-afe.mjs`): `afe-synthetic-v11.xlsx` (the
part, its properties, the web extension parts) and `afe-synthetic-v10.xlsx` (the code
sheet, layout not reproduced, and the locale sheet, both very hidden), from
`f9_lambda_mac.xlsx`. They follow the facts above and prove nothing about Excel.

**To confirm with a real AFE save** (drop the files into `probes/results/afe/`; the tests
there run on each):

1. Does a workbook saved by Excel (Mac and Windows) with AFE modules carry exactly one
   `AFEJSONBlob` part, linked from `webextension1.xml`, as above?
2. Is the module → name rule `Workbook` → unprefixed, `M` → `M.name`? How does AFE store
   a sheet-scoped name it did not create (AFE has none of its own)?
3. Right after the save, does every module entry equal its name modulo whitespace (the
   test expects so)?
4. After an xln build that changes one AFE name, open in Excel and AFE: does AFE show its
   old text, flag the difference, or re-sync from the Name Manager; and does saving a
   module in AFE write the old definition back over xln's?
5. Does Excel re-serialise the part on save (as it does xln's, § F5), and does AFE still
   read it?

## F10, F11: trim ranges, external references, Create from Selection (run 2026-10-07)

**Run by the author in Excel for Mac, 2026-10-07: no refusals, no messages, every C cell
as expected** (4, 6, 8, 4, 100, 42, 7, 71). Steps: `kits/F10-F11.md`; kits made by
`node probes/kits/make-kits.mjs` (plain packages written without Excel). Results:
`results/f10_trim_extref_mac.xlsx`, `results/f11_create_from_selection_mac.xlsx`; the tests
`packages/core/test/file/probe-f10.test.ts` and `probe-f11.test.ts` print the stored forms
next to the assumptions written in them (and in `kits/f11/labels.json`).

**F10, measured** (names and cells alike):

| Typed | Stored |
|---|---|
| `A1.:.A10`, `A1:.A10`, `A1.:A10` | `_xlfn._TRO_ALL(A1:A10)`, `_xlfn._TRO_TRAILING(A1:A10)`, `_xlfn._TRO_LEADING(A1:A10)` |
| `Sheet1!$A$1.:.$A$10` (a name) | `_xlfn._TRO_ALL(Sheet1!$A$1:$A$10)`: the qualifier inside |
| `SUM(A:.A)` | `SUM(_xlfn._TRO_TRAILING(A:A))` |
| `TRIMRANGE(A1:A10)` | `_xlfn.TRIMRANGE(A1:A10)` |
| `'[Other.xlsx]Sheet1'!$A$1` | `[1]Sheet1!$A$1` |
| `Other.xlsx!OtherVal` | `[1]!OtherVal` (in a cell: dynamic-array form, `cm="1"`) |

`[1]` is the first `<externalReference r:id>` of workbook.xml, whose part
`xl/externalLinks/externalLink1.xml` has `<externalBook r:id>` → the file's path (an
external relationship; Excel for Mac also wrote `xxl21:alternateUrls` with an absolute and
a relative path), the other file's sheet names, its defined names (`OtherVal` =
`'Sheet1'!$A$2`) and cached values. The file also kept a `calcChain.xml`.

**What xln does with it**: trim references are syntax (compile and decompile, exact round
trip on every stored form above); `[n]` is shown by the linked file's name
(`[Other.xlsx]Sheet1!$A$1`, `Other.xlsx!OtherVal`, quoted when the file or the sheet needs
it) and compiled back to `[n]`; a workbook the file has no link to is an error (*type the
reference once in Excel, save, then pull*: only Excel writes a link part). The link parts
are copied byte for byte. The `unmeasured-form` warning is gone. Not measured: how the
formula bar shows the reference (xln uses the form Excel shows while the other file is
open), and how Excel stores a trim operator between operands that are not one range token
(`A1:.INDEX(B:B,3)`: xln writes `_xlfn._TRO_TRAILING(A1:INDEX(B:B,3))`, the same function
around the range).

**F11, measured** (Create from Selection, Left column; `labelName` follows it, and so do
C15 and the label notice):

| Label | Name | Rule |
|---|---|---|
| `Gross  income`, `a - b`, `x/y/z`, `R&D` | `Gross__income`, `a___b`, `x_y_z`, `R_D` | one `_` per character a name cannot hold |
| `line⏎break` (stored CR LF) | `line_break` | a line break is one character |
| `  lead`, `trail  ` | `lead`, `trail` | trimmed |
| `Margin %` | `Margin` | characters a name cannot hold at the end are dropped |
| `Q1`, `Tax2024`, `R`, `C`, `rc`, `True` | `Q1_`, `Tax2024_`, `R_`, `C_`, `rc_`, `True_` | `_` **behind** a text that reads as an A1 reference, R, C, RC, TRUE or FALSE |
| `R1C1` | `_R1C1` | `_` in front of an R1C1 reference |
| `A1B`, `v1.2`, `Why?`, `back\slash`, `_under`, `Crescità`, `Café`, `中文` | unchanged | `.`, `?`, `\`, `_` and letters are name characters |
| `2024 sales`, `1st`, `€uro` | `_2024_sales`, `_1st`, `_€uro` | `_` in front of a first character that cannot start a name; `€` may follow but not start |
| `2024`, `3.5` (numbers) | none | a number gives no name |
| 45322 shown as a date | `_31_01_2024` | a date gives the text it shows (the author's locale: 31/01/2024) |
| `Total`, `TOTAL`, `Total` (rows 23–25, Yes to replace) | `Total` on row 25 | duplicates: the last wins, with its spelling |
| empty; 300 characters | none | no name |

Not measured: characters a name cannot hold at the **start** (`(a) b`): `labelName` makes
each one `_` (`_a__b`), `labelNames` also accepts them dropped (`a__b`); which other
symbols besides `€` Excel takes inside a name; the locale's date spelling (without the
shown text, `labelNames` accepts the common short-date spellings).


## Number literals (survey of saved files, 2026-10-07)

xln writes a number literal as the source spells it; Excel parses it to a number and
writes its own spelling when it saves. Since lockfile format 3 every comparison and hash
takes a literal by its value rounded to 15 significant digits (`canonicalNumber`), so a
re-spelling is no change in Excel; the next pull shows Excel's spelling.

| | Evidence |
|---|---|
| **Measured** | `1E-14` in cell formulas → `0.00000000000001` (lbo-ep03r: XlsxWriter's file vs the copy Excel for Mac 16.0300 saved, F5; 3 cells). `1E+99` and `1E+300` kept as written, in names and cells (same pair, and layers-demo). The other 734 literals of that pair and all 326 of layers-demo (XlsxWriter → Excel-saved v3) were already plain and came back byte-identical. No Excel-saved file in `probes/results/` or the corpus has a literal with a trailing zero, a leading `.` or `0`, a lower-case `e`, or more than 15 significant digits |
| **Assumed** | `1.50` → `1.5`, `.5` → `0.5`, `1E3` → `1000`, `1e-7` → `1E-07`, `00.1` → `0.1`; that names re-spell as cells do; that Excel rounds (not truncates) past 15 significant digits (`lockhash.test.ts` pins the consequence: if it truncates, such a literal reads as an edit in Excel after the save). xln does not write Excel's spelling itself: the thresholds for E notation are unmeasured, and any spelling opens |

## To measure (open questions for the next Excel session)

- **Number literals**: build a name and a cell formula with `=0.10+.5+1E3+1e-7+00.1`,
  `=1.50`, `=1E15`, `=1E16`, `=1E21`, `=0.0001`, `=1E-5`, `=1E-10`,
  `=0.1234567890123456789`, `=12345678901234567`, `=-1E-14`; open, save in Excel (Mac and
  Windows), read `xl/workbook.xml` and the sheet: the spelling of each, and whether the
  long ones are rounded or truncated to 15 digits.

- **XLM names** (AFE #10, warning `xlm-name`): in a blank workbook define `Group`,
  `Get.Cell`, `Evaluate`, `Files` and `Scale` as `=LAMBDA(x, x+1)`: does the Name Manager
  take each, and what does `=Group(1)` (and the others) give in a cell, Mac and Windows?
  Refused names become errors; if each calls the LAMBDA, the warning can go. Decided
  without measuring (author, 2026-10-08): the warning covers the macro functions (Ftab)
  only, not the command equivalents (Cetab: `Open`, `Save`, `Copy`, `Table`, `Scale`),
  which are ordinary words. `Scale` here is optional: it tells whether that was right.

## Office Scripts results

*Pending.* Copy the report file into `results/` when done.
