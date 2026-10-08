// The New module sample builds cleanly and its LAMBDA is stored as Excel stores it (M3d: the
// optional parameter `[periods]` is `_xlop.periods`; `[_xlpm.periods]` made Excel drop the name).
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { buildWorkbook, checkStoredForm, equalModuloWhitespace, pullProject, readWorkbook } from "@xln/core";
import { newModuleText } from "../src/model/editor.js";

const base = fileURLToPath(new URL("../../../probes/results/f7_base.xlsx", import.meta.url));
// ANA.GROW typed in Excel's Name Manager and saved by Excel (Mac, 2026-10-06).
const excelSaved = fileURLToPath(new URL("../../../probes/results/lambda_optional_mac.xlsx", import.meta.url));

describe("New module sample", () => {
  it("builds, reads back, and stores ANA.GROW exactly as Excel does", () => {
    const bytes = new Uint8Array(readFileSync(base));
    const files = { ...pullProject(bytes, "f7_base.xlsx").files, "names/ANA.xln": newModuleText("ANA") };
    const r = buildWorkbook({ workbook: bytes, fileName: "f7_base.xlsx", files }, { embed: false, provenance: false });
    expect(r.status, [...r.plan.problems.map((p) => p.message), ...(r.readBack?.problems ?? [])].join("\n")).toBe("built");
    expect(r.readBack?.ok).toBe(true);
    const stored = readWorkbook(r.bytes!).definedNames.find((d) => d.name === "ANA.GROW")!.definition;
    const excel = readWorkbook(new Uint8Array(readFileSync(excelSaved))).definedNames.find((d) => d.name === "ANA.GROW")!.definition;
    expect(excel).toContain("_xlop.periods");
    expect(equalModuloWhitespace(stored, excel)).toBe(true);
    expect(checkStoredForm(stored)).toEqual([]);
  });
});
