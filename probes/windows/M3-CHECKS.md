# M3 Windows check: what to do

Checks on Windows Excel what was checked on the Mac for M3: names written by `xln build`
(M3a), cell formulas written by `build` and `apply` (M3b), the embedded source and the
provenance tags (D5, D6), the F8 files, the 255-character comment limit, and the
PowerShell script behind `xln build --reopen` (E7). No Node needed on Windows. About 10
minutes, with a few questions to answer in the console.

## On Windows

1. **Get the kit.** In the repo folder: `git pull` (on the branch the coordinator names).
   `probes\windows\m3kit\` must exist. Use git, not a downloaded zip.
2. **Close Excel** (save your work first). Part 2 needs Excel closed; the script asks
   again if it is not.
3. **Double-click `probes\windows\run_m3_check.cmd`.** A console opens.
4. **Part 1** (2–3 minutes): an Excel window opens the kit workbooks one by one, reads
   them, saves a copy and closes them. Do not click in it. Then come two files that should
   make Excel complain, each in a new Excel window, after the console says so and you
   press Enter:
   - `c256_comment.xlsx` (a name comment one character over Excel's limit). The Mac
     refused it. If Excel says it cannot open the file: **OK**. If it asks
     *"Abbiamo riscontrato un problema in alcuni contenuti… Ripristinare…?"*
     (*We found a problem with some content…*): **Sì** (Yes), then **Chiudi** (Close).
   - `f8_p2.xlsx` (a stale calculation chain). Excel should ask the same question:
     **Sì**, then on the list of repairs **Chiudi**.

   After each, the console asks two things: whether Excel asked to repair (`y`/`n`) and
   what else it showed (type a few words, or Enter).
5. **Part 2** (3–4 minutes): the `--reopen` script opens and closes a workbook by itself
   several times. You may be asked to:
   - click once on the console window and press Enter (if the script cannot see Excel yet);
   - for `reopen_repair.xlsx`: when Excel asks whether to recover it, click **No** this
     time, within 60 seconds;
   - say whether to try a copy in your OneDrive folder (`y` is useful; it is deleted after).
6. **If the console beeps** and prints `>>> Excel has not answered…`, Excel is showing a
   dialog: switch to it (Alt+Tab) and answer as the message says. Nothing hangs forever.
7. **Send the results back** (the console prints these too):

   ```
   git add probes/results
   git commit -m "M3 Windows check results"
   git push
   ```

   That commits `probes/results/m3-win-<PC>-<time>.txt` (the report) and
   `probes/results/m3win/*_winsaved.xlsx` (Excel's saved copies).

If a window says scripts are disabled: open PowerShell in `probes\windows` and run
`powershell -NoProfile -ExecutionPolicy Bypass -File .\m3_check.ps1`. To skip part 2:
add `-SkipReopen`. If Excel stays open at the end, close it without saving `reopen_*.xlsx`.

## On the Mac, afterwards (coordinator)

```
git pull && npm run build
node probes/windows/check_m3_winsaved.mjs
```

It compares each `_winsaved` copy with its kit file and with Mac Excel's copy
(`m3kit/mac/`): `verify` (only the edited cells may change, and Windows must equal the
Mac), the expected values, names and comments (tags kept), the embedded part (kept, its
line ends) and a pull that must restore the kit project byte for byte; then it prints the
Windows report's summary and the `--reopen` lines (`R01`–`R18`).

## The kit

`make_m3_kit.mjs` writes `m3kit/` from in-repo fixtures only (`probes/results/f7_base.xlsx`,
`f8_base.xlsx` and the F8 files), through `xln pull`, edits to the project, `xln build`,
or `xln apply`; `mac_open_kit.mjs` is the same check in Mac Excel (it wrote `m3kit/mac/`).

| File | What | Expected on open |
|---|---|---|
| `f8_p1`, `f8_q1` | probe F8's patched files, as committed | clean; values as on the Mac |
| `m3a_names` | build ×2 on f7: multi-line LET with a two-line comment, `@hidden`, update, `@renamed`, delete, rescope | clean |
| `m3b_build` | build on f8: named spills bigger and smaller, unnamed cell, block, a whole shared group, slots filled (`Name @B1# = …`: names put on `#` by the statements), a clear | clean |
| `m3b_apply` | `apply` on f8: one shared child, a shared master, number and text cells to formulas, spill ↔ scalar, a slot, a clear | clean |
| `d5_module` | module `FN` with `// @version 1.2`, tags, a two-line comment, embedded source | clean; part kept |
| `c255_comment` | a 255-character comment | clean |
| `c256_comment` | 256 characters (the build refuses it; patched in the XML) | refused or repaired |
| `f8_p2` | stale `calcChain.xml` | repair prompt |

Each has `<name>.expected.json` (cells, names, expressions, the part, what may change);
built ones keep their project in `<name>.xln/`; `xln_excel.ps1` is the `--reopen` script
exactly as xln writes it (a test fails when it goes stale: rerun the generator).
