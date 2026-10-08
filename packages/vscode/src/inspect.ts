// The browser probe's logic, kept free of the vscode API so Vitest can run it in Node.
// Bytes in, plain data out: the same rule as packages/core (no Node built-ins).
import { unzipSync } from "fflate";

export interface ZipEntry {
  name: string;
  size: number;
}

export interface WorkbookProbe {
  entries: ZipEntry[];
  /** Number of <definedName> elements in xl/workbook.xml; -1 when that part is missing. */
  definedNames: number;
}

export function probeXlsx(bytes: Uint8Array): WorkbookProbe {
  const files = unzipSync(bytes);
  const entries = Object.entries(files).map(([name, data]) => ({ name, size: data.length }));
  const workbook = files["xl/workbook.xml"];
  const definedNames = workbook ? countElements(new TextDecoder().decode(workbook), "definedName") : -1;
  return { entries, definedNames };
}

// A count only: the real <definedNames> parser belongs to core/src/file (W1).
// Matches "<tag" followed by a delimiter, so "<definedNames>" is not counted as "<definedName".
export function countElements(xml: string, tag: string): number {
  const open = "<" + tag;
  let count = 0;
  for (let i = xml.indexOf(open); i >= 0; i = xml.indexOf(open, i + open.length)) {
    const next = xml[i + open.length];
    if (next === ">" || next === "/" || next === " " || next === "\t" || next === "\n" || next === "\r") count++;
  }
  return count;
}

/** Excel's owner file for an open workbook: "~$" + the file name, in the same folder. */
export function lockFileName(fileName: string): string {
  return "~$" + fileName;
}

export function baseName(path: string): string {
  return path.slice(path.lastIndexOf("/") + 1);
}

export function dirName(path: string): string {
  const i = path.lastIndexOf("/");
  return i <= 0 ? "/" : path.slice(0, i);
}

export interface InspectReport extends WorkbookProbe {
  file: string;
  bytes: number;
  lockFile: string;
  lockFilePresent: boolean;
  /** Every "~$" entry in the folder, to see what the file system exposes at all. */
  ownerFilesInFolder: string[];
}

export function formatReport(r: InspectReport): string {
  const lines = [
    `Workbook: ${r.file} (${r.bytes} bytes)`,
    `Zip entries (${r.entries.length}):`,
    ...r.entries.map((e) => `  ${e.name}  ${e.size}`),
    `<definedName> elements in xl/workbook.xml: ${r.definedNames < 0 ? "workbook.xml missing" : r.definedNames}`,
    `Lock file ${r.lockFile}: ${r.lockFilePresent ? "PRESENT (workbook is open in Excel)" : "absent"}`,
    `"~$" files in the folder: ${r.ownerFilesInFolder.length ? r.ownerFilesInFolder.join(", ") : "none"}`,
  ];
  return lines.join("\n");
}
