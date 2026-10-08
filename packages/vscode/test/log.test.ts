// The output channel's detail lines for Rename Symbol and the audit (model/log.ts).
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { audit, LOCK_FILE, MANIFEST_FILE, parseLockfile, pullProject, readWorkbook, type SourceRename } from "@xln/core";
import { describe, expect, it } from "vitest";
import { auditLogLines, renameLogLines } from "../src/model/log.js";
import { parseManifest } from "../src/model/manifest.js";
import { isNamesFile, Project } from "../src/model/project.js";

const f7 = fileURLToPath(new URL("../../../probes/results/f7_base.xlsx", import.meta.url));
const traps = fileURLToPath(new URL("../../../probes/fixtures/traps.xlsx", import.meta.url));

describe("log lines", () => {
  it("a rename: edits by file, readers, @renamed added", () => {
    const files = pullProject(new Uint8Array(readFileSync(f7)), "f7_base.xlsx").files;
    const p = new Project("mem:/p", parseManifest(files[MANIFEST_FILE]!));
    for (const [path, text] of Object.entries(files)) if (isNamesFile(path)) p.setFile(path, text);
    p.lock = parseLockfile(files[LOCK_FILE]!);
    const lines = renameLogLines(p.renameEdits("Rate", "Pace") as SourceRename, "book/f7_base.xln");
    expect(lines[0]).toBe("book/f7_base.xln: 12 edit(s) in 3 file(s): names/_unmanaged.xln (4), names/sheets/S1.xln (7), names/sheets/S2.xln (1)");
    expect(lines[1]).toBe("formulas reading Rate: 10 reference(s) in " + (p.renameEdits("Rate", "Pace") as SourceRename).readers + " statement(s)");
    expect(lines).toContain("@renamed(Rate) added: the build renames the name in the workbook and rewrites it in the cells");
  });

  it("an audit: the CLI's header and only the checks that found something", () => {
    const r = audit(readWorkbook(new Uint8Array(readFileSync(traps))), { workbook: "traps.xlsx" });
    const [head, ...checks] = auditLogLines(r);
    expect(head).toMatch(/^xln check traps\.xlsx: \d+ error/);
    expect(checks.length).toBeGreaterThan(0);
    for (const c of checks) expect(c).toMatch(/^ {2}C\d+ .*\d+ (error|warning|info)/);
  });
});
