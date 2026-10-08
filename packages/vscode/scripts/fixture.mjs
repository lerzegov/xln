// The folder both automated probes open: a workbook with a stand-in Excel owner file
// next to it, a second workbook in a subfolder without one, the grammar sample, and
// book/f7_base.xlsx for the browse suite to pull, traps/traps.xlsx for the audit suite,
// and labels/labels.xlsx (f7_base with a label that no longer gives its name, C15).
import { copyFileSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { basename, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

export const root = join(dirname(fileURLToPath(import.meta.url)), "..");

export function makeFixture(name) {
  const results = join(root, "..", "..", "probes", "results");
  const dir = join(root, ".vscode-test", name);
  rmSync(dir, { recursive: true, force: true });
  mkdirSync(join(dir, "sub"), { recursive: true });
  copyFileSync(join(results, "probe_mac.xlsx"), join(dir, "probe_mac.xlsx"));
  copyFileSync(join(results, "probe_win.xlsx"), join(dir, "sub", "probe_win.xlsx"));
  copyFileSync(join(root, "samples", "demo.xln"), join(dir, "demo.xln"));
  mkdirSync(join(dir, "book"));
  copyFileSync(join(results, "f7_base.xlsx"), join(dir, "book", "f7_base.xlsx"));
  mkdirSync(join(dir, "traps"));
  copyFileSync(join(results, "..", "fixtures", "traps.xlsx"), join(dir, "traps", "traps.xlsx"));
  writeFileSync(join(dir, "~$probe_mac.xlsx"), "owner file stand-in");
  mkdirSync(join(dir, "labels"));
  writeFileSync(join(dir, "labels", "labels.xlsx"), labelDriftBook(readFileSync(join(results, "f7_base.xlsx"))));
  return dir;
}

/** f7_base named with Create from Selection over S2!D7:F9, then Gross_income renamed Gross_ind_income. */
function labelDriftBook(bytes) {
  const files = unzipSync(new Uint8Array(bytes));
  const cell = (r, text) => `<c r="${r}" t="inlineStr"><is><t>${text}</t></is></c>`;
  const rows = [["7", "Revenue"], ["8", "COGS"], ["9", "Gross income"]].map(([r, label]) => `<row r="${r}">${cell(`D${r}`, label)}<c r="E${r}"><v>1</v></c><c r="F${r}"><v>2</v></c></row>`);
  files["xl/worksheets/sheet2.xml"] = strToU8(strFromU8(files["xl/worksheets/sheet2.xml"]).replace("</sheetData>", `${rows.join("")}</sheetData>`));
  const defs = [["COGS", 8], ["Gross_ind_income", 9], ["Revenue", 7]].map(([n, r]) => `<definedName name="${n}" localSheetId="1">S2!$E$${r}:$F$${r}</definedName>`);
  files["xl/workbook.xml"] = strToU8(strFromU8(files["xl/workbook.xml"]).replace("<definedNames>", `<definedNames>${defs.join("")}`));
  return zipSync(files);
}

/**
 * `--smoke <workbook>` on the command line: a folder with a copy of that workbook and
 * the marker that makes the test runners run the smoke suite. Undefined without it.
 */
export function smokeFixture(name, argv = process.argv) {
  const k = argv.indexOf("--smoke");
  if (k < 0) return undefined;
  const src = argv[k + 1];
  if (!src) throw new Error("--smoke needs a workbook path");
  const dir = join(root, ".vscode-test", name);
  rmSync(dir, { recursive: true, force: true });
  mkdirSync(dir, { recursive: true });
  const file = basename(src);
  copyFileSync(src, join(dir, file));
  writeFileSync(join(dir, "xln-smoke.json"), JSON.stringify({ workbook: file }));
  return dir;
}
