import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { buildWorkbook, pullProject } from "@xln/core";
import { buildDiagnostics, buildTarget, conflictTexts, formatBuildSummary, refusalMessage } from "../src/model/build.js";

const f7 = fileURLToPath(new URL("../../../probes/results/f7_base.xlsx", import.meta.url));
const S1 = "names/sheets/S1.xln";

describe("build model", () => {
  it("desktop overwrites with a backup; the browser writes <name>.xln.xlsx and keeps the lockfile", () => {
    expect(buildTarget("lbo.xlsx", false)).toEqual({ file: "lbo.xlsx", overwrite: true, backup: "lbo.backup.xlsx", updateLock: true });
    expect(buildTarget("lbo.xlsm", true)).toEqual({ file: "lbo.xln.xlsm", overwrite: false, backup: undefined, updateLock: false });
  });

  it("a conflict as two .xln texts, Excel's left", () => {
    const t = conflictTexts({
      kind: "both-changed",
      key: "S2!Loc",
      message: "",
      excel: { name: "Loc", scope: "S2", display: "8", comment: null, hidden: false },
      source: { name: "Loc", scope: "S2", display: "9", comment: "doc", hidden: true, file: "names/sheets/S2.xln", line: 4 },
    });
    expect(t.excel).toBe("@sheet('S2')\nLoc = 8;\n");
    expect(t.source).toBe("/** doc */\n@sheet('S2')\n@hidden\nLoc = 9;\n");
    expect(conflictTexts({ kind: "deleted-in-source", key: "X", message: "", excel: { name: "X", scope: null, display: "1", comment: null, hidden: false } }).source).toMatch(/not in the source/);
  });

  it("a cell conflict as two cell statements", () => {
    const cell = { sheet: "BS", range: "C8" };
    const t = conflictTexts({
      kind: "both-changed",
      key: "BS!C8",
      message: "",
      excel: { name: "", scope: "BS", display: "Rate*30", comment: null, hidden: false, cell },
      source: { name: "", scope: "BS", display: "Rate*4", comment: null, hidden: false, cell, file: "names/sheets/BS.xln", line: 9 },
    });
    expect(t.excel).toBe("@BS!C8 = Rate*30;\n");
    expect(t.source).toBe("@BS!C8 = Rate*4;\n");
    const slot = conflictTexts({ kind: "created-in-both", key: "Costs", message: "", excel: { name: "Costs", scope: null, display: "7", comment: null, hidden: false, cell: { sheet: "Slot", range: "B2" } } });
    expect(slot.excel).toBe("Costs @Slot!B2 = 7;\n");
  });

  it("the summary lists cell changes", () => {
    const files = pullProject(new Uint8Array(readFileSync(f7)), "f7_base.xlsx").files;
    const edited = { ...files, [S1]: files[S1]!.replace("@C8 = Rate*3;", "@C8 = Rate*4;") };
    const r = buildWorkbook({ workbook: new Uint8Array(readFileSync(f7)), fileName: "f7_base.xlsx", files: edited }, { embed: true });
    expect(formatBuildSummary("f7_base.xlsx", r)).toEqual(["xln build f7_base.xlsx: 2 changes", "  set formula of S1!C8", "  embed the source (3 names files, xln.lock.json)"]);
  });

  it("a refusal: reasons first in the summary, a modal text with them, errors placed in their files (feedback 2026-10-07)", () => {
    const bytes = new Uint8Array(readFileSync(f7));
    const files = pullProject(bytes, "f7_base.xlsx").files;
    // Fn deleted while S1!C10 still calls it; C4 moved to a new function in the same build.
    const edited = {
      ...files,
      "names/_unmanaged.xln": files["names/_unmanaged.xln"]!.replace("Fn = LAMBDA(x, x*Rate);\n", ""),
      "names/FN.xln": "FN.TIMES = LAMBDA(x, x*Rate);\n",
      [S1]: files[S1]!.replace("@C4 = Fn(10);", "@C4 = FN.TIMES(10);"),
    };
    const r = buildWorkbook({ workbook: bytes, fileName: "f7_base.xlsx", files: edited }, { embed: false });
    expect(r.status).toBe("refused");
    const line = edited[S1]!.slice(0, edited[S1]!.indexOf("@C10 =")).split("\n").length;
    const lines = formatBuildSummary("f7_base.xlsx", r);
    expect(lines[0]).toBe("xln build f7_base.xlsx: refused: 0 conflict(s), 1 error(s); nothing written");
    expect(lines[1]!.startsWith(`  error ${S1}:${line}: deleting Fn refused: 1 place in the workbook refers to it`)).toBe(true);
    expect(lines[2]).toBe("  the plan, not written:");
    expect(lines.slice(3)).toContain("    delete Fn");
    expect(refusalMessage("f7_base.xlsx", r)).toBe(
      `Build of f7_base.xlsx refused: nothing was written.\n\n• ${S1}:${line}: deleting Fn refused: 1 place in the workbook refers to it by name and would break: cell S1!C10 (${S1}:${line}). Change that formula in the source, or keep the name`,
    );
    expect(buildDiagnostics(r)).toEqual([{ file: S1, line: line - 1, code: "in-use", message: r.plan.problems[0]!.message }]);
  });

  it("the modal lists the first reasons, then how many more", () => {
    const bytes = new Uint8Array(readFileSync(f7));
    const files = pullProject(bytes, "f7_base.xlsx").files;
    const bad = Array.from({ length: 7 }, (_, k) => `Bad${k} = Nope${k}(1);\n`).join("");
    const r = buildWorkbook({ workbook: bytes, fileName: "f7_base.xlsx", files: { ...files, "names/_unmanaged.xln": files["names/_unmanaged.xln"]! + bad } }, { embed: false });
    expect(r.status).toBe("refused");
    const text = refusalMessage("f7_base.xlsx", r, 3);
    expect(text.split("\n").filter((l) => l.startsWith("• "))).toHaveLength(3);
    expect(text.endsWith(`\n… and ${r.plan.problems.filter((p) => p.severity === "error").length - 3} more`)).toBe(true);
  });
});
