// The M3 Windows check kit (probes/windows/m3kit, written by make_m3_kit.mjs) must match the
// code it checks: the --reopen script the author runs on Windows is the one xln writes, and
// every expectation points at something the kit workbook has.
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { readEmbeddedSource, readWorkbook, WINDOWS_EXCEL_SCRIPT } from "@xln/core";

const KIT = join(import.meta.dirname, "..", "..", "..", "probes", "windows", "m3kit");
const kit = JSON.parse(readFileSync(join(KIT, "kit.json"), "utf8")) as {
  files: { file: string; expected: string; repair: boolean | string }[];
  reopen: { workbook: string; rebuilt: string; repair: string; script: string };
};

describe("M3 Windows kit", () => {
  it("carries the current --reopen PowerShell script (rerun probes/windows/make_m3_kit.mjs if not)", () => {
    const ps1 = readFileSync(join(KIT, kit.reopen.script), "utf8").replace(/\r\n/g, "\n");
    expect(ps1).toBe(WINDOWS_EXCEL_SCRIPT);
    for (const f of [kit.reopen.workbook, kit.reopen.rebuilt, kit.reopen.repair]) expect(existsSync(join(KIT, f))).toBe(true);
  });

  it("is small: under 15 workbooks and 1 MB", () => {
    expect(kit.files.length).toBeLessThan(15);
    const total = kit.files.reduce((n, f) => n + readFileSync(join(KIT, f.file)).length, 0);
    expect(total).toBeLessThan(1 << 20);
  });

  for (const f of kit.files) {
    it(`${f.file}: the expectations match the workbook`, () => {
      const exp = JSON.parse(readFileSync(join(KIT, f.expected), "utf8"));
      const wb = readWorkbook(readFileSync(join(KIT, f.file)));
      expect(wb.sheets.length).toBe(exp.sheets);
      const sheets = new Set(wb.sheets.map((s) => s.name));
      for (const c of exp.cells) expect(sheets.has(c.sheet)).toBe(true);
      const names = new Set(wb.definedNames.map((n) => (n.scope.kind === "sheet" ? `${n.scope.name}!${n.name}` : n.name).toLowerCase()));
      for (const n of exp.names) expect(names.has(n.name.toLowerCase())).toBe(!n.absent);
      expect(!!readEmbeddedSource(readFileSync(join(KIT, f.file)))).toBe(exp.customXml);
      if (exp.project) expect(existsSync(join(KIT, exp.project, "xln.lock.json"))).toBe(true);
      if (exp.verify) expect(existsSync(join(KIT, exp.verify.before))).toBe(true);
    });
  }
});
