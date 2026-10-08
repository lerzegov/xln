#!/usr/bin/env node
// mac_open_kit.mjs: the Mac counterpart of m3_check.ps1, a sanity run of the M3 Windows kit
// before it goes to Windows. Mac Excel opens each kit workbook that should open cleanly
// (never the repair ones), reads the expected cells, names and expressions, saves a copy
// as `m3kit/mac/<name>_macsaved.xlsx` and closes it. The copies give check_m3_winsaved.mjs
// a dry run and a second oracle (Windows's values must equal the Mac's).
//
//   node probes/windows/mac_open_kit.mjs [name ...]
//
// Drives Excel through AppleScript like probe F8: one open per file with a time limit,
// a check that a workbook of that name opened, only our own workbook closed. Files are
// staged in Office's group container so Excel's sandbox asks for nothing.

import { spawnSync } from "node:child_process";
import { copyFileSync, existsSync, mkdirSync, readFileSync, rmSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));
const KIT = join(HERE, "m3kit");
const OUT = join(KIT, "mac");
const STAGE = join(homedir(), "Library/Group Containers/UBF8T346G9.Office/xln-m3kit");

const READ = String.raw`
on run argv
  set src to item 1 of argv
  set out to item 2 of argv
  set want to item 3 of argv
  set refs to paragraphs of (item 4 of argv)
  set exprs to paragraphs of (item 5 of argv)
  set TB to tab
  set RS to character id 30 -- records may hold line breaks (multi-line definitions)
  tell application "Microsoft Excel"
    with timeout of 60 seconds
      open workbook workbook file name src
    end timeout
    set wb to missing value
    repeat with i from 1 to (count of workbooks)
      if (name of workbook i) is want then set wb to workbook want
    end repeat
    if wb is missing value then return "NOT OPENED: no workbook named " & want
    set r to "sheets" & TB & (count of worksheets of wb) & RS
    repeat with x in refs
      if (x as text) is not "" then
        set AppleScript's text item delimiters to "!"
        set parts to text items of (x as text)
        set AppleScript's text item delimiters to ""
        set rg to range (item 2 of parts) of worksheet (item 1 of parts) of wb
        set r to r & "cell" & TB & (x as text) & TB & (formula2 of rg) & TB & (string value of rg) & RS
      end if
    end repeat
    repeat with i from 1 to (count of named items of wb)
      set n to named item i of wb
      set r to r & "name" & TB & (name of n) & TB & (references of n) & TB & (visible of n) & RS
    end repeat
    repeat with e in exprs
      if (e as text) is not "" then
        set v to evaluate name (e as text)
        set r to r & "eval" & TB & (e as text) & TB & (v as text) & RS
      end if
    end repeat
    save workbook as wb filename out
    set AppleScript's text item delimiters to "/"
    set nm to last text item of out
    set AppleScript's text item delimiters to ""
    close workbook nm saving no
  end tell
  return r
end run`;

function osa(script, args, timeoutMs) {
  const r = spawnSync("osascript", ["-", ...args], { input: script, encoding: "utf8", timeout: timeoutMs });
  if (r.error) return { out: "", error: r.error.code === "ETIMEDOUT" ? "TIMEOUT: Excel did not answer (a dialog may be open in Excel)" : r.error.message };
  return { out: r.stdout.trim(), error: r.status === 0 ? undefined : r.stderr.trim() };
}

// The same comparisons as m3_check.ps1: formulas modulo whitespace and sheet quotes,
// numbers within a relative 1e-9.
const squash = (f) => String(f ?? "").replace(/\s+/g, "").replace(/'([A-Za-z_][A-Za-z0-9_.]*)'!/g, "$1!").toLowerCase();
function sameValue(expected, shown) {
  if (expected === null) return shown === "";
  if (typeof expected === "object") return shown === expected.error;
  if (typeof expected === "number") {
    const n = Number(shown.replace(",", "."));
    return Number.isFinite(n) && Math.abs(n - expected) <= 1e-9 * Math.max(1, Math.abs(expected));
  }
  return shown === expected;
}

const kit = JSON.parse(readFileSync(join(KIT, "kit.json"), "utf8"));
const only = process.argv.slice(2);
mkdirSync(OUT, { recursive: true });
mkdirSync(STAGE, { recursive: true });
let fails = 0;
for (const f of kit.files) {
  const name = f.file.replace(/\.xlsx$/, "");
  if (f.repair !== false || (only.length && !only.includes(name))) continue;
  const exp = JSON.parse(readFileSync(join(KIT, f.expected), "utf8"));
  const src = join(STAGE, f.file);
  const saved = join(STAGE, `${name}_macsaved.xlsx`);
  copyFileSync(join(KIT, f.file), src);
  rmSync(saved, { force: true });
  const cells = exp.cells.map((c) => `${c.sheet}!${c.cell}`).join("\n");
  const exprs = exp.evaluate.map((e) => e.expr).join("\n");
  const r = osa(READ, [src, saved, f.file, cells, exprs], 150_000);
  console.log(`== ${f.file}`);
  if (r.error || !r.out.startsWith("sheets")) {
    console.log(`   FAIL ${r.error ?? r.out}`);
    fails++;
    // A failure halfway leaves our workbook open: close it (only it), unsaved.
    osa(`on run argv\ntell application "Microsoft Excel"\nrepeat with i from (count of workbooks) to 1 by -1\nif (name of workbook i) is item 1 of argv then close workbook i saving no\nend repeat\nend tell\nend run`, [f.file], 30_000);
    continue;
  }
  const rows = r.out.split("\x1e").filter((l) => l !== "").map((l) => l.split("\t"));
  const shownCells = new Map(rows.filter((x) => x[0] === "cell").map((x) => [x[1], { formula: x[2], text: x[3] ?? "" }]));
  const names = new Map(rows.filter((x) => x[0] === "name").map((x) => [x[1].toLowerCase(), { refersTo: x[2], visible: x[3] === "true" }]));
  const evals = new Map(rows.filter((x) => x[0] === "eval").map((x) => [x[1], x[2]]));
  const line = (ok, what) => { if (!ok) fails++; console.log(`   ${ok ? "PASS" : "FAIL"} ${what}`); };
  line(Number(rows[0][1]) === exp.sheets, `sheets ${rows[0][1]} (expected ${exp.sheets})`);
  for (const c of exp.cells) {
    const got = shownCells.get(`${c.sheet}!${c.cell}`);
    if (c.formula !== undefined) line(squash(got?.formula) === squash(c.formula), `${c.sheet}!${c.cell} formula2 '${got?.formula}' (expected '${c.formula}')`);
    if (c.value !== undefined) line(got !== undefined && sameValue(c.value, got.text), `${c.sheet}!${c.cell} shows '${got?.text}' (expected ${JSON.stringify(c.value)})`);
  }
  for (const n of exp.names) {
    // AppleScript lists a sheet-scoped name as 'S2'!TmpC.
    const key = n.name.includes("!") ? n.name.replace(/^([^!]+)!/, "'$1'!").toLowerCase() : n.name.toLowerCase();
    const got = names.get(key) ?? names.get(n.name.toLowerCase());
    if (n.absent) { line(!got, `name ${n.name} absent`); continue; }
    line(got && squash(got.refersTo) === squash(n.refersTo), `name ${n.name} refers to '${got?.refersTo}' (expected '${n.refersTo}')`);
    if (n.visible !== undefined) line(got && got.visible === n.visible, `name ${n.name} visible ${got?.visible} (expected ${n.visible}); comment not readable over AppleScript (T04)`);
  }
  for (const e of exp.evaluate) line(sameValue(e.value, evals.get(e.expr) ?? ""), `evaluate ${e.expr} -> '${evals.get(e.expr)}' (expected ${e.value})`);
  if (existsSync(saved)) copyFileSync(saved, join(OUT, `${name}_macsaved.xlsx`));
  line(existsSync(saved), `saved m3kit/mac/${name}_macsaved.xlsx`);
}
console.log(fails ? `${fails} check(s) failed` : "all checks passed");
process.exit(fails ? 1 : 0);
