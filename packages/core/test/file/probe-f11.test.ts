// Probe F11 (kit: probes/kits/F10-F11.md): the names Excel for Mac's Create from Selection
// (Left column) makes from the labels of probes/kits/f11/labels.json, against the names
// assumed there and against `labelName` / `labelNames` (labels.ts, partly assumed). Skips
// until the author saves probes/results/f11_create_from_selection_mac.xlsx. The table marks
// each row where Excel differs from the assumption; the follow-up tightens `labelNames` to
// the measured rule.
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { cellValueMap, labelName, labelNames, readWorkbook } from "../../src/index.js";

const ROOT = join(import.meta.dirname, "..", "..", "..", "..", "probes");
const FILE = join(ROOT, "results", "f11_create_from_selection_mac.xlsx");

interface Label {
  row: number;
  label: string | number;
  padTo?: number;
  date?: true;
  /** The name assumed for this row's B cell (null: none, the label gives no name or a later row replaced it). */
  expect: string | null;
  note: string;
}
const short = (s: string) => (s.length > 40 ? `${s.slice(0, 20)}…(${s.length})` : s);
const LABELS: Label[] = JSON.parse(readFileSync(join(ROOT, "kits", "f11", "labels.json"), "utf8"));

describe("the F11 kit's assumptions", () => {
  it("lists the labels the kit writes, one per row", () => {
    expect(new Set(LABELS.map((l) => l.row)).size).toBe(LABELS.length);
    expect(LABELS.length).toBeGreaterThan(30);
  });
});

describe.skipIf(!existsSync(FILE))("probe F11: Create from Selection (Excel for Mac)", () => {
  // Read only when the file is there: a skipped describe's body still runs.
  const present = existsSync(FILE);
  const bytes = present ? new Uint8Array(readFileSync(FILE)) : new Uint8Array();
  const wb = present ? readWorkbook(bytes) : { definedNames: [], sheets: [] };
  const values = (present ? cellValueMap(bytes).get("Labels") : undefined) ?? new Map();
  /** The names on Labels!$B$row (Create from Selection writes the sheet qualified and absolute). */
  const namesOn = (row: number) =>
    wb.definedNames.filter((d) => {
      const def = d.definition.split("'").join("");
      return def === `Labels!$B$${row}`;
    }).map((d) => d.name);

  it("prints Excel's names next to the assumed ones and labelName/labelNames", () => {
    const lines = ["", "F11: row | label | Excel's name | assumed | = or DIFF | labelName | in labelNames? | note"];
    for (const l of LABELS) {
      const excel = namesOn(l.row);
      const got = excel.join(" ") || "(none)";
      const want = l.expect ?? "(none)";
      const v = values.get(`A${l.row}`);
      const mine = labelNames(v);
      const covered = excel.length === 0 ? "" : excel.every((n) => mine.some((m) => m.toLowerCase() === n.toLowerCase())) ? "yes" : "NO";
      const shownLabel = JSON.stringify(typeof l.label === "string" && l.padTo ? `${l.label}…(${l.padTo})` : l.label);
      lines.push(`  ${l.row} | ${shownLabel} | ${short(got)} | ${want} | ${got === want ? "=" : "DIFF"} | ${short(labelName(v) ?? "(none)")} | ${covered} | ${l.note}`);
    }
    const rows = new Set(LABELS.map((l) => `Labels!$B$${l.row}`));
    for (const d of wb.definedNames) if (!rows.has(d.definition.split("'").join(""))) lines.push(`  (other) ${d.name} = ${d.definition}`);
    console.log(lines.join("\n"));
    expect(wb.sheets.map((s) => s.name)).toContain("Labels");
  });

  // What the checker relies on (labelDrift, C15): every name Excel makes is one of the
  // spellings `labelNames` accepts for its label; otherwise C15 reports a false drift.
  it("every name Excel made is among labelNames' spellings of its label", () => {
    const missed = LABELS.flatMap((l) => namesOn(l.row).filter((n) => !labelNames(values.get(`A${l.row}`)).some((m) => m.toLowerCase() === n.toLowerCase())).map((n) => `row ${l.row}: ${n} (labelNames: ${labelNames(values.get(`A${l.row}`)).join(", ") || "none"})`));
    expect(missed).toEqual([]);
  });
});
