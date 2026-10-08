// Probe F10 (kit: probes/kits/F10-F11.md): how Excel for Mac stores trim ranges and
// references to another workbook, in defined names and in cells. Skips until the author
// saves probes/results/f10_trim_extref_mac.xlsx. It prints every stored form next to the
// one assumed (all as assumed, run 2026-10-07) and xln's display of it. What xln does with
// them (trim references as syntax, `[n]` by the linked file's name) is tested in
// test/build/f10.test.ts.
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { decompileWithDiagnostics, equalModuloWhitespace, readWorkbook, type WorkbookLink } from "../../src/index.js";

const FILE = join(import.meta.dirname, "..", "..", "..", "..", "probes", "results", "f10_trim_extref_mac.xlsx");

// Hypotheses (not measured): trim ranges become hidden functions with the `_xlfn.` prefix,
// the area left plain; another workbook is `[n]`, the n-th <externalReference> of
// workbook.xml, whose part xl/externalLinks/externalLinkN.xml keeps the file's path and
// cached values. TRIMRANGE is an ordinary new function (`_xlfn.`).
const NAMES: Record<string, string> = {
  TrimAll: "_xlfn._TRO_ALL(Sheet1!$A$1:$A$10)",
  TrimEnd: "_xlfn._TRO_TRAILING(Sheet1!$A$1:$A$10)",
  TrimStart: "_xlfn._TRO_LEADING(Sheet1!$A$1:$A$10)",
  TrimFn: "_xlfn.TRIMRANGE(Sheet1!$A$1:$A$10)",
  Ext: "[1]Sheet1!$A$1",
  ExtName: "[1]!OtherVal",
};
// Sheet1's cells as typed (kit steps), the stored form assumed, the value they should show.
const CELLS: [cell: string, typed: string, stored: string, value: number][] = [
  ["C1", "ROWS(A1.:.A10)", "ROWS(_xlfn._TRO_ALL(A1:A10))", 4],
  ["C2", "ROWS(A1:.A10)", "ROWS(_xlfn._TRO_TRAILING(A1:A10))", 6],
  ["C3", "ROWS(A1.:A10)", "ROWS(_xlfn._TRO_LEADING(A1:A10))", 8],
  ["C4", "ROWS(TRIMRANGE(A1:A10))", "ROWS(_xlfn.TRIMRANGE(A1:A10))", 4],
  ["C5", "SUM(A:.A)", "SUM(_xlfn._TRO_TRAILING(A:A))", 100],
  ["C6", "'[Other.xlsx]Sheet1'!$A$1", "[1]Sheet1!$A$1", 42],
  ["C7", "Other.xlsx!OtherVal", "[1]!OtherVal", 7],
  ["C8", "ROWS(TrimAll)+ROWS(TrimEnd)+ROWS(TrimStart)+ROWS(TrimFn)+Ext+ExtName", "ROWS(TrimAll)+ROWS(TrimEnd)+ROWS(TrimStart)+ROWS(TrimFn)+Ext+ExtName", 71],
];

const mark = (got: string | undefined, want: string) => (got !== undefined && equalModuloWhitespace(got, want) ? "=" : "DIFF");
const shown = (stored: string | undefined, links: readonly WorkbookLink[]) => {
  if (stored === undefined) return "";
  const d = decompileWithDiagnostics(stored, { links });
  return d.text + d.diagnostics.map((g) => `  [${g.severity}: ${g.message}]`).join("");
};

describe.skipIf(!existsSync(FILE))("probe F10: trim ranges and external references (Excel for Mac)", () => {
  const bytes = existsSync(FILE) ? new Uint8Array(readFileSync(FILE)) : new Uint8Array();

  it("prints the stored forms against the hypotheses", () => {
    const wb = readWorkbook(bytes);
    const lines: string[] = ["", "F10 defined names: name | stored | assumed | = or DIFF | xln's decompile"];
    for (const [name, want] of Object.entries(NAMES)) {
      const d = wb.definedNames.find((n) => n.name === name);
      lines.push(`  ${name} | ${d?.definition ?? "(missing: refused?)"} | ${want} | ${mark(d?.definition, want)} | ${shown(d?.definition, wb.externalLinks)}`);
    }
    for (const d of wb.definedNames) if (!(d.name in NAMES)) lines.push(`  (other) ${d.name} | ${d.definition}`);
    const s1 = wb.sheets.find((s) => s.name === "Sheet1");
    lines.push("F10 cells on Sheet1: cell | typed | stored | assumed | = or DIFF | kind, cm, ref | cached value (expected) | xln's decompile");
    for (const [cell, typed, want, value] of CELLS) {
      const f = s1?.formulas.find((x) => x.cell === cell);
      lines.push(`  ${cell} | ${typed} | ${f?.text ?? "(no formula)"} | ${want} | ${mark(f?.text, want)} | ${f?.kind ?? ""} ${f?.cm ?? ""} ${f?.range ?? ""} | ${String(f?.value.value)} (${value}) | ${shown(f?.text, wb.externalLinks)}`);
    }
    // The external link parts: what xln would have to write for `[n]` to mean anything.
    const zip = unzipSync(bytes);
    const wbXml = strFromU8(zip["xl/workbook.xml"]!);
    const k = wbXml.indexOf("<externalReferences");
    lines.push(`F10 workbook.xml <externalReferences>: ${k < 0 ? "(none)" : wbXml.slice(k, wbXml.indexOf("</externalReferences>", k) + 21)}`);
    for (const p of Object.keys(zip).filter((p) => p.startsWith("xl/externalLinks/")).sort()) lines.push(`F10 part ${p}:\n${strFromU8(zip[p]!)}`);
    for (const p of Object.keys(zip).filter((p) => p.includes("metadata") || p.includes("calcChain"))) lines.push(`F10 part present: ${p}`);
    console.log(lines.join("\n"));
    expect(wb.sheets.map((s) => s.name)).toContain("Sheet1");
  });
});
