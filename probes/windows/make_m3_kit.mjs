#!/usr/bin/env node
// make_m3_kit.mjs: builds the Windows check kit for M3 (M3a names, M3b cells, D5 embedded
// source, D6 tags, E7 --reopen) into probes/windows/m3kit/.
//
//   npm run build && node probes/windows/make_m3_kit.mjs
//
// Every kit workbook comes from an in-repo fixture (probes/results/*.xlsx), never from the
// corpus, and goes through the real pipeline: `xln pull`, edits to the project files,
// `xln build` (or `xln apply` for cell changes the source cannot express: a value cell
// turned into a formula, one member of a shared group). Each kit file gets
// `<name>.expected.json`: what probes/windows/m3_check.ps1 reads in Windows Excel and what
// it should find. Expected values are written by hand from the fixtures' inputs (an
// oracle independent of xln); expected comments come from the built file, since they are
// xln's own text (tags included). Built projects are kept beside their workbook
// (`<name>.xln/`) so that check_m3_winsaved.mjs can test the embedded part byte for byte.
//
// The generator is reproducible: the zip writer stamps new entries with the first
// entry's time, so a rerun gives the same bytes.

import { spawnSync } from "node:child_process";
import { copyFileSync, cpSync, existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { unzipSync, strFromU8, strToU8 } from "fflate";
import { compile, readWorkbook, replaceZipEntries, WINDOWS_EXCEL_SCRIPT, EMBED_NS } from "../../packages/core/dist/index.js";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, "../..");
const KIT = join(HERE, "m3kit");
const RESULTS = join(ROOT, "probes/results");
const XLN = join(ROOT, "packages/cli/bin/xln.js");
const STAGE = mkdtempSync(join(tmpdir(), "xln-m3kit-"));

function xln(args, { cwd = STAGE, allow = [0] } = {}) {
  const r = spawnSync(process.execPath, [XLN, ...args], { cwd, encoding: "utf8" });
  if (!allow.includes(r.status)) throw new Error(`xln ${args.join(" ")} exited ${r.status}\n${r.stdout}\n${r.stderr}`);
  return r.stdout;
}

/** Replaces one exact piece of a project file; fails if it is not there exactly once. */
function edit(file, from, to) {
  const text = readFileSync(file, "utf8");
  const n = text.split(from).length - 1;
  if (n !== 1) throw new Error(`${relative(STAGE, file)}: expected one '${from}', found ${n}`);
  writeFileSync(file, text.replace(from, () => to));
}

function stageCopy(fixture, name) {
  const p = join(STAGE, `${name}.xlsx`);
  copyFileSync(fixture, p);
  return p;
}

function entryText(bytes, name) {
  const files = unzipSync(bytes, { filter: (f) => f.name === name });
  if (!files[name]) throw new Error(`no ${name} in the zip`);
  return strFromU8(files[name]);
}

/** The project files the build embeds (D5): names/**, the lockfile, the config. */
function projectFiles(dir) {
  const out = [];
  const walk = (d) => {
    for (const e of readdirSync(d, { withFileTypes: true })) {
      const p = join(d, e.name);
      if (e.isDirectory()) walk(p);
      else {
        const rel = relative(dir, p).split(sep).join("/");
        if (rel.startsWith("names/") || rel === "xln.lock.json" || rel === "xln.config.json") out.push(rel);
      }
    }
  };
  walk(dir);
  return out.sort();
}

function keepProject(name) {
  const src = join(STAGE, `${name}.xln`);
  const dst = join(KIT, `${name}.xln`);
  for (const rel of projectFiles(src)) {
    mkdirSync(dirname(join(dst, rel)), { recursive: true });
    copyFileSync(join(src, rel), join(dst, rel));
  }
  return `${name}.xln`;
}

function nameComment(bytes, name) {
  const n = readWorkbook(bytes).definedNames.find((x) => x.name === name);
  if (!n) throw new Error(`built workbook has no name ${name}`);
  return n.comment ?? null;
}

const kitIndex = [];
function writeKitFile(name, stagedOrFixture, expected) {
  copyFileSync(stagedOrFixture, join(KIT, `${name}.xlsx`));
  const e = { file: `${name}.xlsx`, ...expected };
  writeFileSync(join(KIT, `${name}.expected.json`), JSON.stringify(e, null, 2) + "\n");
  kitIndex.push({ file: e.file, expected: `${name}.expected.json`, repair: e.repair });
  console.log(`  ${e.file}: ${e.title}`);
}

// Cells: { sheet, cell, formula?, value? }. `formula` is Excel's Formula2 ("" = no formula
// and no constant), compared modulo whitespace and sheet-name quotes; `value` is Value2:
// a number, a string, null for an empty cell, or { error: "#SPILL!" }.
const C = (sheet, cell, formula, value) => ({ sheet, cell, ...(formula !== undefined ? { formula } : {}), ...(value !== undefined ? { value } : {}) });
const err = (e) => ({ error: e });

// ---------------------------------------------------------------------------------------
// F8 files as they are (probe F8, patched by Python): expectations from the Mac run's
// report, column p1 (q1 for the plain workbook): what Mac Excel showed on open.
function f8Expectations(variant, section) {
  const text = readFileSync(join(RESULTS, "f8-mac-20261005-091726.txt"), "utf8");
  const lines = text.split("\n");
  const start = lines.findIndex((l) => l.startsWith(`## 5a ${section}:`));
  const col = { main: { p1: 1 }, plain: { q1: 1 } }[section][variant];
  const cells = [];
  const fdiff = new Map();
  for (let i = start + 1; i < lines.length && !lines[i].startsWith("## "); i++) {
    const l = lines[i];
    const diff = l.match(/^\s+(\w+) (\S+!\S+): oracle '(.*)' vs '(.*)'$/);
    if (diff) { if (diff[1] === variant) fdiff.set(diff[2], diff[4]); continue; }
    const row = l.match(/^\s+(\S+!\S+)\s+(.*?)\s+oracle f: ?(.*)$/);
    if (!row) continue;
    const shown = row[2].split("|")[col].trim().replace(/^!/, "");
    cells.push({ ref: row[1], shown, formula: row[3] });
  }
  return cells.map(({ ref, shown, formula }) => {
    const [sheet, cell] = ref.split("!");
    const f = fdiff.has(ref) ? fdiff.get(ref) : formula;
    const value = shown === "" ? null : shown.startsWith("#") ? err(shown) : Number.isFinite(Number(shown)) ? Number(shown) : shown;
    return C(sheet, cell, f, value);
  });
}

// ---------------------------------------------------------------------------------------
function main() {
  // Keep mac/ (Mac Excel's saved copies, written by mac_open_kit.mjs): the kit is
  // reproducible, so they stay valid across reruns.
  mkdirSync(KIT, { recursive: true });
  for (const e of readdirSync(KIT)) if (e !== "mac") rmSync(join(KIT, e), { recursive: true, force: true });
  console.log(`staging in ${STAGE}`);

  // -- F8 files ----------------------------------------------------------------------------
  const p1 = f8Expectations("p1", "main");
  writeKitFile("f8_p1", join(RESULTS, "f8_p1.xlsx"), {
    title: "F8 p1: cell formulas patched by the F8 recipe (Python), expected clean",
    from: "probes/results/f8_p1.xlsx (probe F8, as committed)",
    repair: false, sheets: 4, customXml: false,
    cells: p1, names: [], evaluate: [],
    // Mac Excel's own recalculation of the same file, saved: Windows must match it exactly.
    verify: { before: "../../results/f8_p1_resaved.xlsx", mayChange: [] },
  });
  writeKitFile("f8_q1", join(RESULTS, "f8_q1.xlsx"), {
    title: "F8 q1: dynamic array in a workbook without metadata.xml (created), expected clean",
    from: "probes/results/f8_q1.xlsx (probe F8, as committed)",
    repair: false, sheets: 1, customXml: false,
    cells: f8Expectations("q1", "plain"), names: [], evaluate: [],
    verify: { before: "../../results/f8_q1_resaved.xlsx", mayChange: [] },
  });

  // -- M3a: names, two builds on f7_base ---------------------------------------------------
  {
    const name = "m3a_names";
    const wb = stageCopy(join(RESULTS, "f7_base.xlsx"), name);
    xln(["pull", wb]);
    const U = join(STAGE, `${name}.xln/names/_unmanaged.xln`);
    const S2 = join(STAGE, `${name}.xln/names/sheets/S2.xln`);
    edit(U, "Rate = 0.1;", "Rate = 0.2;");
    edit(U, "RateX = 0.5;\n", [
      "RateX = 0.5;",
      "",
      "/**",
      " * Growth factor over n years.",
      " * Second line of the comment.",
      " */",
      "Grow = LAMBDA(g, n,",
      "  LET(f, 1 + g,",
      "    f ^ n));",
      "",
      "@hidden",
      "Secret = 42;",
      "",
      "TmpA = 11;",
      "TmpB = 22;",
      "TmpC = 33;",
      "",
    ].join("\n"));
    xln(["build", wb, "--embed"]);
    // Second build: rename, delete, rescope (names the first build created).
    edit(U, "TmpA = 11;\nTmpB = 22;\nTmpC = 33;\n", "@renamed(TmpA)\nRenamedA = 11;\n");
    edit(S2, "Loc = 7;\n", "Loc = 7;\nTmpC = 33;\n");
    xln(["build", wb, "--embed"]);
    const bytes = readFileSync(wb);
    writeKitFile(name, wb, {
      title: "M3a: names written by xln build (multi-line LET with a two-line comment, hidden, update, rename, delete, rescope)",
      from: "probes/results/f7_base.xlsx: pull; build 1 (Rate 0.1 -> 0.2, new Grow, Secret, TmpA-C); build 2 (@renamed(TmpA) RenamedA, TmpB deleted, TmpC moved into names/sheets/S2.xln: local to S2)",
      repair: false, sheets: 2, customXml: true, project: keepProject(name),
      cells: [
        C("S1", "A1", "=Rate*2", 0.4),
        C("S1", "C1", "=RateX+Rate", 0.7),
        C("S1", "E1", "=SEQUENCE(3)*Rate", 0.2),
        C("S1", "E3", "", 0.6000000000000001),
        C("S1", "C4", "=Fn(10)", 2),
        C("S1", "C10", "=Rate2+Fn(1)", 0.6000000000000001),
        C("S2", "A2", "=Rate+Loc", 7.2),
      ],
      names: [
        { name: "Grow", refersTo: "=LAMBDA(g,n,LET(f,1+g,f^n))", comment: nameComment(bytes, "Grow"), visible: true },
        { name: "Secret", refersTo: "=42", visible: false },
        { name: "Rate", refersTo: "=0.2", visible: true },
        { name: "RenamedA", refersTo: "=11", visible: true },
        { name: "S2!TmpC", refersTo: "=33", visible: true },
        { name: "TmpA", absent: true },
        { name: "TmpB", absent: true },
        { name: "TmpC", absent: true },
      ],
      evaluate: [
        { expr: "Grow(0.1,2)", value: 1.2100000000000002 },
        { expr: "Secret", value: 42 },
        { expr: "RenamedA", value: 11 },
        { expr: "S2!TmpC", value: 33 },
      ],
      // Rate changed: every cell that reads it may change from the fixture's cached values.
      verify: { before: "m3a_names.xlsx", mayChange: ["S1!A1:E11", "S2!A1:B5"] },
    });
  }

  // -- M3b: cell statements edited in the source, one build on f8_base ---------------------
  {
    const name = "m3b_build";
    const wb = stageCopy(join(RESULTS, "f8_base.xlsx"), name);
    xln(["pull", wb]);
    const sheets = join(STAGE, `${name}.xln/names/sheets`);
    edit(join(sheets, "D.xln"), "Spill @C6# = SEQUENCE(3);", "Spill @C6# = SEQUENCE(5);");    // named anchor, bigger
    edit(join(sheets, "D.xln"), "Spill2 @F6# = SEQUENCE(4);", "Spill2 @F6# = SEQUENCE(2);");  // named anchor, smaller
    edit(join(sheets, "D.xln"), "@E7 = SUM(L6#);", "@E7 = ;");                                    // clear
    edit(join(sheets, "N.xln"), "@B2 = A1*10;", "@B2 = A1*100;");                                 // unnamed cell
    edit(join(sheets, "N.xln"), "@B5:B6 = A1;", "@B5:B6 = A1+1000;");                             // block range
    edit(join(sheets, "Sh.xln"), "@C1:C5 = A1+1;", "@C1:C5 = A1+10;");                            // a whole shared group
    edit(join(sheets, "Slot.xln"), "Revenue @B1 = ;", "Revenue @B1# = SEQUENCE(1,3)*10;"); // slot -> spill, name -> # (explicit)
    edit(join(sheets, "Slot.xln"), "Costs @B2 = ;", "Costs @B2# = 5;");                   // slot -> scalar, name -> # (C6# on a 1x1 result)
    edit(join(sheets, "Slot.xln"), "Far @B9 = ;", "Far @B9# = SUM(Revenue)*2+Costs;");   // slot in a row absent from the file
    xln(["build", wb, "--embed"]);
    writeKitFile(name, wb, {
      title: "M3b: cell statements edited in the source and written by xln build (spills bigger/smaller, block, shared group, slots, clear)",
      from: "probes/results/f8_base.xlsx: pull; edit names/sheets/{D,N,Sh,Slot}.xln; build",
      repair: false, sheets: 4, customXml: true, project: keepProject(name),
      cells: [
        C("N", "B1", "=SUM(A1:A3)", 6),
        C("N", "B2", "=A1*100", 100),
        C("N", "B5", "=A1+1000", 1001),
        C("N", "B6", "=A2+1000", 1002),
        C("D", "C6", "=SEQUENCE(5)", 1),
        C("D", "C10", "", 5),
        C("D", "E1", "=SUM(Spill)", 15),
        C("D", "E4", "=ROWS(C6#)", 5),
        C("D", "F6", "=SEQUENCE(2)", 1),
        C("D", "F7", "", 2),
        C("D", "F8", "", null),
        C("D", "F9", "", null),
        C("D", "E2", "=SUM(Spill2)", 3),
        C("D", "E3", "=ROWS(F6#)", 2),
        C("D", "E5", "=SUM(H6#)", 6),
        C("D", "E7", "", null),
        C("Sh", "C1", "=A1+10", 11),
        C("Sh", "C3", "=A3+10", 13),
        C("Sh", "C5", "=A5+10", 15),
        C("Sh", "B3", "=A3*2", 6),
        C("Slot", "B1", "=SEQUENCE(1,3)*10", 10),
        C("Slot", "D1", "", 30),
        C("Slot", "B2", "=5", 5),
        C("Slot", "B9", "=SUM(Revenue)*2+Costs", 125),
      ],
      names: [
        { name: "Spill", refersTo: "=D!$C$6#", visible: true },
        { name: "Spill2", refersTo: "=D!$F$6#", visible: true },
        { name: "Revenue", refersTo: "=Slot!$B$1#", visible: true },
        { name: "Costs", refersTo: "=Slot!$B$2#", visible: true },
        { name: "Far", refersTo: "=Slot!$B$9#", visible: true },
        { name: "Tax", refersTo: "=Slot!$B$6", visible: true },
      ],
      evaluate: [{ expr: "SUM(Revenue)", value: 60 }],
      verify: { before: "m3b_build.xlsx", mayChange: ["N!B2:B6", "D!C6:C10", "D!E1:E7", "D!F6:F9", "Sh!C1:C5", "Slot!B1:D1", "Slot!B2", "Slot!B9"] },
    });
  }

  // -- M3b: a hand-built change set through xln apply on f8_base ---------------------------
  {
    const name = "m3b_apply";
    const wb = stageCopy(join(RESULTS, "f8_base.xlsx"), name);
    const set = (sheet, range, display) => ({ op: "set-cell-formula", sheet, range, stored: compile(display), display });
    const changes = [
      set("Sh", "C3", "A3*100"),            // one child of the shared group C1:C5
      set("Sh", "D1", "A1*30"),             // the master of D1:D5 (the group is un-shared)
      set("N", "C1", "SUM(A1:A3)*2"),       // number cell -> formula
      set("N", "C2", '"hi "&A1'),           // shared-string cell -> formula with a text result
      set("N", "C4", "SEQUENCE(2,2)"),      // text cell -> spill
      set("Slot", "B6", "7"),               // a slot filled, its name left on the cell
      set("D", "H6", "42"),                 // spill -> scalar
      set("D", "P6", "SEQUENCE(3)"),        // scalar -> spill
      { op: "clear-cell-formula", sheet: "D", range: "E6" },
    ];
    writeFileSync(join(STAGE, `${name}.changes.json`), JSON.stringify(changes, null, 2) + "\n");
    xln(["apply", wb, join(STAGE, `${name}.changes.json`)]);
    copyFileSync(join(STAGE, `${name}.changes.json`), join(KIT, `${name}.changes.json`));
    writeKitFile(name, wb, {
      title: "M3b: a change set through xln apply (shared child and master, value and text cells to formulas, slot, spill <-> scalar, clear)",
      from: "probes/results/f8_base.xlsx: xln apply m3b_apply.changes.json",
      repair: false, sheets: 4, customXml: false,
      cells: [
        C("Sh", "C2", "=A2+1", 3),
        C("Sh", "C3", "=A3*100", 300),
        C("Sh", "C4", "=A4+1", 5),
        C("Sh", "D1", "=A1*30", 30),
        C("Sh", "D2", "=A2*3", 6),
        C("Sh", "D5", "=A5*3", 15),
        C("N", "C1", "=SUM(A1:A3)*2", 12),
        C("N", "C2", '="hi "&A1', "hi 1"),
        C("N", "C4", "=SEQUENCE(2,2)", 1),
        C("N", "D5", "", 4),
        C("Slot", "B6", "=7", 7),
        C("D", "H6", "=42", 42),
        C("D", "H7", "", null),
        C("D", "E5", "=SUM(H6#)", 42),
        C("D", "P6", "=SEQUENCE(3)", 1),
        C("D", "P8", "", 3),
        C("D", "E6", "", null),
      ],
      names: [{ name: "Tax", refersTo: "=Slot!$B$6", visible: true }],
      evaluate: [],
      verify: { before: "m3b_apply.xlsx", mayChange: ["Sh!C3", "Sh!D1", "N!C1:D5", "Slot!B6", "D!H6:H8", "D!E5:E6", "D!P6:P8"] },
    });
  }

  // -- D5/D6: a module with a version, tags, a multi-line comment, the embedded part --------
  {
    const name = "d5_module";
    const wb = stageCopy(join(RESULTS, "f7_base.xlsx"), name);
    xln(["pull", wb]);
    writeFileSync(join(STAGE, `${name}.xln/names/FN.xln`), [
      "// module: FN, written for the M3 Windows check",
      "// @version 1.2",
      "",
      "/**",
      " * Doubles a number.",
      " * Second line: checked on Windows (D5, D6).",
      " */",
      "FN.Double = LAMBDA(x, x * 2);",
      "",
      "/** Adds the workbook's Rate. */",
      "FN.AddRate = LAMBDA(x,",
      "  x + Rate);",
      "",
    ].join("\n"));
    xln(["build", wb, "--embed"]);
    const bytes = readFileSync(wb);
    const cDouble = nameComment(bytes, "FN.Double");
    const cAdd = nameComment(bytes, "FN.AddRate");
    if (!/\[xln FN 1\.2 #[0-9a-f]+\]$/.test(cDouble ?? "")) throw new Error(`FN.Double has no tag: ${cDouble}`);
    writeKitFile(name, wb, {
      title: "D5/D6: module FN (// @version 1.2) built with the embedded source and provenance tags",
      from: "probes/results/f7_base.xlsx: pull; add names/FN.xln; build (embed and tags on)",
      repair: false, sheets: 2, customXml: true, project: keepProject(name),
      cells: [C("S1", "A1", "=Rate*2", 0.2), C("S1", "C4", "=Fn(10)", 1)],
      names: [
        { name: "FN.Double", refersTo: "=LAMBDA(x,x*2)", comment: cDouble, visible: true },
        { name: "FN.AddRate", refersTo: "=LAMBDA(x,x+Rate)", comment: cAdd, visible: true },
      ],
      evaluate: [{ expr: "FN.Double(21)", value: 42 }, { expr: "FN.AddRate(1)", value: 1.1 }],
      verify: { before: "d5_module.xlsx", mayChange: [] },
    });
  }

  // -- Comment length: 255 (Excel's limit, opens on the Mac) and 256 (refused on the Mac) ---
  {
    const name = "c255_comment";
    const wb = stageCopy(join(RESULTS, "f7_base.xlsx"), name);
    xln(["pull", wb]);
    const head = "C255 comment: the longest a name comment can be. ";
    const doc = (head + "abcdefghij".repeat(30)).slice(0, 254) + "!";
    if (doc.length !== 255) throw new Error("bad 255 comment");
    edit(join(STAGE, `${name}.xln/names/_unmanaged.xln`), "RateX = 0.5;\n", `RateX = 0.5;\n\n/** ${doc} */\nLongDoc = 1;\n`);
    xln(["build", wb, "--no-embed", "--no-tags"]);
    const bytes = readFileSync(wb);
    const got = nameComment(bytes, "LongDoc");
    if (got !== doc) throw new Error(`LongDoc comment came back as ${JSON.stringify(got)}`);
    writeKitFile(name, wb, {
      title: "Comment length 255: the limit, opens on the Mac",
      from: "probes/results/f7_base.xlsx: pull; LongDoc with a 255-character doc comment; build --no-embed --no-tags",
      repair: false, sheets: 2, customXml: false,
      cells: [C("S1", "A1", "=Rate*2", 0.2)],
      names: [{ name: "LongDoc", refersTo: "=1", comment: doc, visible: true }],
      evaluate: [{ expr: "LongDoc", value: 1 }],
      verify: { before: "c255_comment.xlsx", mayChange: [] },
    });

    // 256: the build refuses it (error comment-length), so patch the 255 file's XML.
    const n256 = "c256_comment";
    const refused = spawnSync(process.execPath, [XLN, "build", wb, "--dry-run", "--no-embed", "--no-tags"], { cwd: STAGE, encoding: "utf8" });
    edit(join(STAGE, `${name}.xln/names/_unmanaged.xln`), `/** ${doc} */`, `/** ${doc}Z */`);
    const r = spawnSync(process.execPath, [XLN, "build", wb, "--dry-run", "--no-embed", "--no-tags"], { cwd: STAGE, encoding: "utf8" });
    if (refused.status !== 0 || r.status !== 1 || !/255/.test(r.stdout + r.stderr)) throw new Error(`expected xln build to refuse a 256-character comment:\n${r.stdout}${r.stderr}`);
    const xml = entryText(bytes, "xl/workbook.xml");
    const from = `comment="${doc}"`;
    if (xml.split(from).length !== 2) throw new Error("the 255 comment is not in workbook.xml exactly once");
    const patched = replaceZipEntries(bytes, new Map([["xl/workbook.xml", strToU8(xml.replace(from, () => `comment="${doc}Z"`))]]));
    const p256 = join(STAGE, `${n256}.xlsx`);
    writeFileSync(p256, patched);
    if (nameComment(patched, "LongDoc")?.length !== 256) throw new Error("patched comment is not 256 characters");
    writeKitFile(n256, p256, {
      title: "Comment length 256: one over the limit (xln build refuses it; patched in the XML). On the Mac Excel refused to open it",
      from: "c255_comment.xlsx with one character added to LongDoc's comment attribute (direct XML patch)",
      repair: "maybe",
      repairHint: "Excel will probably refuse this file, or offer to repair it. If it asks whether to recover the contents, click Yes [Si], then Close [Chiudi]. If it shows an error, click OK.",
      sheets: 2, customXml: false,
      cells: [C("S1", "A1", "=Rate*2", 0.2)],
      names: [{ name: "LongDoc", refersTo: "=1", visible: true }],
      evaluate: [],
      verify: null,
    });
  }

  // -- F8 p2: stale calcChain, Excel repairs it (opened last) ------------------------------
  writeKitFile("f8_p2", join(RESULTS, "f8_p2.xlsx"), {
    title: "F8 p2: as p1 but with the stale calcChain.xml kept: Excel asks to repair it",
    from: "probes/results/f8_p2.xlsx (probe F8, as committed)",
    repair: true,
    repairHint: "Excel will say it found a problem with some content and ask whether to recover it. Click Yes [Si]. Then a second window lists what was repaired: click Close [Chiudi].",
    sheets: 4, customXml: false,
    cells: p1, names: [], evaluate: [],
    verify: { before: "../../results/f8_p1_resaved.xlsx", mayChange: [] },
  });

  // -- the --reopen script, byte for byte as xln build --reopen writes it -------------------
  writeFileSync(join(KIT, "xln_excel.ps1"), WINDOWS_EXCEL_SCRIPT, "utf8");

  const kit = {
    format: "xln.m3kit/1",
    note: "Written by probes/windows/make_m3_kit.mjs; read by m3_check.ps1 (Windows) and check_m3_winsaved.mjs (Mac).",
    embedNamespace: EMBED_NS,
    files: kitIndex,
    // --reopen: the workbook first opened must be one Excel saved (no fullCalcOnLoad), so that
    // 'saved' means what it says; then a build's output (fullCalcOnLoad="1") replaces it,
    // which shows whether Excel marks a workbook dirty after recalculating it on load.
    reopen: { workbook: "../../results/f8_p1_resaved.xlsx", rebuilt: "m3b_build.xlsx", repair: "f8_p2.xlsx", script: "xln_excel.ps1" },
  };
  writeFileSync(join(KIT, "kit.json"), JSON.stringify(kit, null, 2) + "\n");
  rmSync(STAGE, { recursive: true, force: true });
  let total = 0;
  for (const f of readdirSync(KIT)) if (f.endsWith(".xlsx")) total += readFileSync(join(KIT, f)).length;
  console.log(`kit: ${kitIndex.length} workbooks, ${Math.round(total / 1024)} KB in ${relative(ROOT, KIT)}`);
}

try {
  main();
} catch (e) {
  console.error(`make_m3_kit: ${e.message}\n(staging kept in ${STAGE})`);
  process.exit(1);
}
