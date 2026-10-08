#!/usr/bin/env node
// check_m3_winsaved.mjs: the Mac half of the M3 Windows check. Reads the copies Windows
// Excel saved (probes/results/m3win/<name>_winsaved.xlsx, written by m3_check.ps1) and
// compares them with the kit (probes/windows/m3kit/):
//
//   - verify: cached values against the kit file (only the cells in `verify.mayChange` may
//     differ) and against Mac Excel's saved copy (m3kit/mac/, nothing may differ);
//   - the expected cells' cached values (expected.json);
//   - names: present or absent, definition modulo whitespace, comment exactly (D6 tags);
//   - D5: the embedded part is still there, the line ends Excel wrote in it, and it still
//     carries the kit project byte for byte (an archive copy: pull does not read it);
//   - how comment line breaks were stored (`_x000a_`, `&#10;`).
//
//   npm run build
//   node probes/windows/check_m3_winsaved.mjs            Windows copies, after the author's run
//   node probes/windows/check_m3_winsaved.mjs --mac      dry run on Mac Excel's copies (m3kit/mac)

import { spawnSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, relative, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { unzipSync, strFromU8 } from "fflate";
import { readCellValues, readEmbeddedSource, readWorkbook, verifyValues, valueText, WINDOWS_EXCEL_SCRIPT } from "../../packages/core/dist/index.js";

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = join(HERE, "../..");
const KIT = join(HERE, "m3kit");
const XLN = join(ROOT, "packages/cli/bin/xln.js");
const mac = process.argv.includes("--mac");
const DIR = mac ? join(KIT, "mac") : join(ROOT, "probes/results/m3win");
const SUFFIX = mac ? "_macsaved" : "_winsaved";

let fails = 0;
let warns = 0;
const out = (s) => console.log(s);
const pass = (ok, what) => { if (!ok) fails++; out(`   ${ok ? "PASS" : "FAIL"} ${what}`); };
const info = (what) => out(`   INFO ${what}`);
const warn = (what) => { warns++; out(`   WARN ${what}`); };

// "S1!A1:E11" -> does it hold sheet/cell?
function colNum(s) { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
function parseCell(a) { const m = /^([A-Z]+)(\d+)$/.exec(a); return m ? { c: colNum(m[1]), r: Number(m[2]) } : undefined; }
function inRanges(ranges, sheet, cell) {
  const p = parseCell(cell);
  return ranges.some((spec) => {
    const [s, rg] = spec.split("!");
    if (s !== sheet || !p) return false;
    const [a, b = a] = rg.split(":").map(parseCell);
    return p.c >= a.c && p.c <= b.c && p.r >= a.r && p.r <= b.r;
  });
}

// Excel saves the newer errors (#SPILL!, #CALC!, ...) as `<v>#VALUE!</v>` plus a rich-value
// reference (measured on the Mac copies): the cached value of a #SPILL! cell reads #VALUE!.
const NEW_ERRORS = new Set(["#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#UNKNOWN!", "#BUSY!"]);
function sameValue(expected, got) {
  if (expected === null) return got === undefined || got === "";
  if (typeof expected === "object") return typeof got === "object" && (got.error === expected.error || (NEW_ERRORS.has(expected.error) && got.error === "#VALUE!"));
  if (typeof expected === "number") return typeof got === "number" && Math.abs(got - expected) <= 1e-9 * Math.max(1, Math.abs(expected));
  return got === expected;
}
// Stored definitions modulo whitespace and optional quotes around a simple sheet name
// (Excel re-spaces, and writes 'Slot'!$B$1 as Slot!$B$1).
const squash = (s) => String(s ?? "").replace(/\s+/g, "").replace(/'([A-Za-z_][A-Za-z0-9_.]*)'!/g, "$1!");
const scopeKey = (n) => (n.scope.kind === "sheet" ? `${n.scope.name}!${n.name}` : n.name).toLowerCase();

function entries(bytes) {
  return unzipSync(bytes);
}
function count(text, what) { return text.split(what).length - 1; }

function projectFiles(dir) {
  const files = new Map();
  const walk = (d) => {
    for (const e of readdirSync(d, { withFileTypes: true })) {
      const p = join(d, e.name);
      if (e.isDirectory()) walk(p);
      else {
        const rel = relative(dir, p).split(sep).join("/");
        if (rel.startsWith("names/") || rel === "xln.lock.json" || rel === "xln.config.json") files.set(rel, readFileSync(p));
      }
    }
  };
  walk(dir);
  return files;
}

// The part is an archive copy (pull does not read it since 2026-10-06): it must still
// carry the kit project byte for byte after Windows Excel saved the file.
function checkRestore(saved, project) {
  const emb = readEmbeddedSource(readFileSync(saved));
  if (!emb) { pass(false, "the embedded part is gone"); return; }
  if (emb.damaged.length) pass(false, `embedded files failing their checksum: ${emb.damaged.join(", ")}`);
  const want = projectFiles(join(KIT, project));
  const diffs = [];
  for (const [p, b] of want) {
    const g = emb.files[p];
    if (g === undefined) diffs.push(`${p} missing`);
    else if (Buffer.compare(Buffer.from(g, "utf8"), b) !== 0) diffs.push(`${p} differs (${Buffer.byteLength(g)} vs ${b.length} bytes${g.includes("\r") ? ", has CR" : ""})`);
  }
  for (const p of Object.keys(emb.files)) if (!want.has(p)) diffs.push(`${p} extra`);
  pass(diffs.length === 0, `the embedded part carries the kit project byte for byte (${want.size} files)${diffs.length ? ": " + diffs.join("; ") : ""}`);
  const r = spawnSync(process.execPath, [XLN, "pull", saved, "--out", mkdtempSync(join(tmpdir(), "xln-m3pull-")), "--json"], { encoding: "utf8" });
  if (r.status !== 0) { pass(false, `pull of the saved copy failed (exit ${r.status}): ${r.stderr.trim()}`); return; }
  const report = JSON.parse(r.stdout);
  const edited = (report.provenance ?? []).filter((p) => p.state === "edited");
  if ((report.provenance ?? []).length) pass(edited.length === 0, `provenance tags: ${report.provenance.length}, edited since the build: ${edited.length}`);
  rmSync(report.out, { recursive: true, force: true });
}

function checkFile(f) {
  const exp = JSON.parse(readFileSync(join(KIT, f.expected), "utf8"));
  const name = f.file.replace(/\.xlsx$/, "");
  const saved = join(DIR, `${name}${SUFFIX}.xlsx`);
  out(`== ${f.file}: ${exp.title}`);
  if (!existsSync(saved)) {
    if (f.repair === false) pass(false, `no ${relative(ROOT, saved)} (did Excel open it?)`);
    else info(`no saved copy (expected when Excel refused the file or ${mac ? "the Mac run skips repair files" : "the author declined the repair"})`);
    return;
  }
  const bytes = readFileSync(saved);
  const kitBytes = readFileSync(join(KIT, f.file));

  // verify: only the expected cells may change; Mac and Windows must agree.
  if (exp.verify) {
    const before = readFileSync(join(KIT, exp.verify.before));
    const v = verifyValues(before, bytes);
    const outside = v.changed.filter((c) => !inRanges(exp.verify.mayChange, c.sheet, c.cell));
    pass(outside.length === 0 && v.sheetsRemoved.length === 0, `verify against ${exp.verify.before}: ${v.cells} cells, ${v.changed.length} changed, ${outside.length} outside the edited cells${outside.length ? ": " + outside.slice(0, 8).map((c) => `${c.sheet}!${c.cell} ${valueText(c.before)} -> ${valueText(c.after)}`).join(", ") : ""}`);
    const macCopy = join(KIT, "mac", `${name}_macsaved.xlsx`);
    if (!mac && existsSync(macCopy)) {
      const m = verifyValues(readFileSync(macCopy), bytes, { tolerance: 1e-12 });
      pass(m.changed.length === 0, `same values as Mac Excel's saved copy: ${m.changed.length} differ${m.changed.length ? ": " + m.changed.slice(0, 8).map((c) => `${c.sheet}!${c.cell} mac ${valueText(c.before)} win ${valueText(c.after)}`).join(", ") : ""}`);
    }
  }

  // expected cells, as cached by Excel's save
  const values = new Map(readCellValues(bytes).map((s) => [s.sheet, s.cells]));
  const bad = [];
  for (const c of exp.cells) {
    if (!("value" in c)) continue;
    const got = values.get(c.sheet)?.get(c.cell);
    if (!sameValue(c.value, got)) bad.push(`${c.sheet}!${c.cell} ${valueText(got)} (expected ${c.value === null ? "(empty)" : typeof c.value === "object" ? c.value.error : JSON.stringify(c.value)})`);
  }
  pass(bad.length === 0, `${exp.cells.length} expected cell values${bad.length ? ": " + bad.join(", ") : ""}`);

  // names: present or absent, definition modulo whitespace against the kit file, comment exactly
  const wb = readWorkbook(bytes);
  const kitNames = new Map(readWorkbook(kitBytes).definedNames.map((n) => [scopeKey(n), n]));
  const names = new Map(wb.definedNames.filter((n) => !n.isXlPrefixed).map((n) => [scopeKey(n), n]));
  for (const e of exp.names) {
    const n = names.get(e.name.toLowerCase());
    if (e.absent) { pass(!n, `name ${e.name} absent`); continue; }
    if (!n) { pass(false, `name ${e.name} missing`); continue; }
    const k = kitNames.get(e.name.toLowerCase());
    if (k) pass(squash(n.definition) === squash(k.definition), `name ${e.name} stored ${JSON.stringify(n.definition)}${squash(n.definition) === squash(k.definition) ? "" : ` (kit ${JSON.stringify(k.definition)})`}`);
    if (e.visible !== undefined) pass(n.hidden === !e.visible, `name ${e.name} hidden=${n.hidden}`);
    if ("comment" in e) pass(n.comment === e.comment, `name ${e.name} comment ${JSON.stringify(n.comment)}${n.comment === e.comment ? "" : ` (expected ${JSON.stringify(e.comment)})`}`);
    else if (n.comment) info(`name ${e.name} comment (${n.comment.length} chars) ${JSON.stringify(n.comment.slice(0, 80))}`);
  }

  // raw XML: comment line breaks, the embedded part's line ends, calcChain
  const z = entries(bytes);
  const wbx = strFromU8(z["xl/workbook.xml"]);
  const comments = [...wbx.matchAll(/comment="([^"]*)"/g)].map((m) => m[1]);
  const multi = comments.filter((c) => /_x000[ad]_|&#1[03];|&#x[aAdD];|\n/.test(c));
  if (multi.length) info(`comment line breaks stored as: ${[...new Set(multi.flatMap((c) => c.match(/_x000[ad]_|&#1[03];|&#x[aAdD];|\r?\n/g)))].map((x) => JSON.stringify(x)).join(", ")}`);
  info(`calcChain.xml ${z["xl/calcChain.xml"] ? "present (rebuilt by Excel)" : "absent"}; ${(/<calcPr[^>]*>/.exec(wbx) ?? ["no calcPr"])[0]}`);
  const emb = readEmbeddedSource(bytes);
  if (exp.customXml) {
    pass(!!emb, `embedded source part ${emb ? `kept at ${emb.location.item}` : "missing"}`);
    if (emb) {
      const t = strFromU8(z[emb.location.item]);
      info(`${emb.location.item}: ${t.length} chars, ${count(t, "\r\n")} CR LF of ${count(t, "\n")} line ends (kit: ${count(strFromU8(entries(kitBytes)[emb.location.item] ?? new Uint8Array()), "\r\n")} CR LF)`);
      const props = Object.keys(z).filter((p) => /^customXml\/itemProps\d+\.xml$/.test(p));
      info(`itemProps: ${props.map((p) => `${p}${strFromU8(z[p]).includes("urn:xln:embedded-source:1") ? " (our schemaRef)" : ""}`).join(", ") || "none"}`);
    }
    if (emb && exp.project) checkRestore(saved, exp.project);
  } else if (emb) warn(`an embedded part appeared in a file built without one`);
}

const kit = JSON.parse(readFileSync(join(KIT, "kit.json"), "utf8"));
out(`M3 kit check of ${relative(ROOT, DIR)}/*${SUFFIX}.xlsx${mac ? " (dry run on Mac Excel's copies)" : ""}`);
const ps1 = readFileSync(join(KIT, kit.reopen.script), "utf8").replace(/\r\n/g, "\n");
if (ps1 !== WINDOWS_EXCEL_SCRIPT) warn(`m3kit/${kit.reopen.script} is not the current WINDOWS_EXCEL_SCRIPT: rerun make_m3_kit.mjs`);
for (const f of kit.files) checkFile(f);

if (!mac) {
  const res = join(ROOT, "probes/results");
  const reports = readdirSync(res).filter((f) => /^m3-win-.*\.txt$/.test(f)).sort((a, b) => statSync(join(res, a)).mtimeMs - statSync(join(res, b)).mtimeMs);
  if (reports.length) {
    const last = reports[reports.length - 1];
    const lines = readFileSync(join(res, last), "utf8").replace(/^﻿/, "").split(/\r?\n/);
    out(`== Windows report ${last}`);
    for (const l of lines.filter((l) => /^(S00|S01|R\d\d) \|/.test(l))) out(`   ${l}`);
  } else out("== no probes/results/m3-win-*.txt yet");
}
out(fails ? `${fails} check(s) failed, ${warns} warning(s)` : `all checks passed${warns ? `, ${warns} warning(s)` : ""}`);
process.exit(fails ? 1 : 0);
