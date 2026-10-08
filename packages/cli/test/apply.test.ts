// `xln apply` with hand-built change sets on the F8 workbook, and the `--reopen` sequence
// with a stand-in for Excel (the real one is checked by hand: PLAN.md, M3b Excel check).
import { copyFileSync, existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { excelControl as coreExcelControl, readWorkbook, type ExcelControl, type ExcelWorkbookState } from "@xln/core";
import { main } from "../src/main.js";
import { aroundExcel, EXIT_REOPEN_FAILED, EXIT_UNSAVED } from "../src/reopen.js";

const F8 = join(import.meta.dirname, "..", "..", "..", "probes", "results", "f8_base.xlsx");
const tmp = mkdtempSync(join(tmpdir(), "xln-apply-"));
afterAll(() => rmSync(tmp, { recursive: true, force: true }));

function capture() {
  const io = { stdout: "", stderr: "", out: (s: string) => void (io.stdout += s), err: (s: string) => void (io.stderr += s) };
  return io;
}

function setup(name: string, changes: unknown): { wb: string; cs: string } {
  const wb = join(tmp, `${name}.xlsx`);
  copyFileSync(F8, wb);
  const cs = join(tmp, `${name}.json`);
  writeFileSync(cs, JSON.stringify(changes));
  return { wb, cs };
}

describe("xln apply", () => {
  it("applies cell changes, keeps a backup, reports spills emptied and groups un-shared", async () => {
    const { wb, cs } = setup("cells", {
      format: "xln.changes/1",
      workbook: "cells.xlsx",
      changes: [
        { op: "set-cell-formula", sheet: "D", range: "C6", stored: "_xlfn.SEQUENCE(2)", display: "SEQUENCE(2)" },
        { op: "set-cell-formula", sheet: "Sh", range: "C1", stored: "A1+1000", display: "A1+1000" },
        { op: "set-cell-formula", sheet: "Slot", range: "B9", stored: "Tax+1", display: "Tax+1" },
      ],
    });
    const io = capture();
    expect(await main(["apply", wb, cs], io)).toBe(0);
    expect(io.stdout).toContain("applied 3 changes");
    expect(io.stdout).toContain("D: old spill emptied: C7 C8");
    expect(io.stdout).toContain("Sh: un-shared: C2 C3 C4 C5");
    expect(io.stdout).toContain("calcChain.xml dropped");
    expect(existsSync(join(tmp, "cells.backup.xlsx"))).toBe(true);
    const f = readWorkbook(new Uint8Array(readFileSync(wb))).sheets[1]!.formulas.find((x) => x.cell === "C6")!;
    expect([f.kind, f.text, f.range]).toEqual(["dynamic-array", "_xlfn.SEQUENCE(2)", "C6"]);
  });

  it("refuses a change that does not apply: nothing written", async () => {
    const { wb, cs } = setup("bad", [{ op: "set-cell-formula", sheet: "Nope", range: "A1", stored: "1", display: "1" }]);
    const before = readFileSync(wb);
    const io = capture();
    expect(await main(["apply", wb, cs], io)).toBe(4);
    expect(io.stderr).toContain("no sheet 'Nope'");
    expect(readFileSync(wb).equals(before)).toBe(true);
  });

  it("refuses while Excel has the file open", async () => {
    const { wb, cs } = setup("locked", [{ op: "clear-cell-formula", sheet: "N", range: "B3" }]);
    writeFileSync(join(tmp, "~$locked.xlsx"), "");
    expect(await main(["apply", wb, cs], capture())).toBe(3);
  });

  it("--json and --dry-run", async () => {
    const { wb, cs } = setup("dry", [{ op: "clear-cell-formula", sheet: "N", range: "B3" }]);
    const before = readFileSync(wb);
    const io = capture();
    expect(await main(["apply", wb, cs, "--dry-run", "--json"], io)).toBe(0);
    const j = JSON.parse(io.stdout);
    expect(j.report.sheets.N.cleared).toEqual(["B3"]);
    expect(j.readBack.ok).toBe(true);
    expect(readFileSync(wb).equals(before)).toBe(true);
  });
});

/** A stand-in for desktop Excel that records what it was asked. */
function fakeExcel(state: ExcelWorkbookState, openOk = true): ExcelControl & { calls: string[] } {
  const calls: string[] = [];
  return {
    calls,
    state: () => (calls.push("state"), state),
    close: () => (calls.push("close"), { ok: true, message: "closed" }),
    open: () => (calls.push("open"), openOk ? { ok: true, message: "opened" } : { ok: false, message: "repair prompt" }),
  };
}

describe("--reopen", () => {
  const path = join(tmp, "reopen.xlsx");

  it("closes, runs the build, opens again", () => {
    const x = fakeExcel({ running: true, open: true, saved: true });
    const r = aroundExcel(path, x, false, () => (x.calls.push("build"), 0));
    expect(x.calls).toEqual(["state", "close", "build", "open"]);
    expect(r.exit).toBeUndefined();
  });

  it("refuses with unsaved changes in Excel unless --discard", () => {
    const x = fakeExcel({ running: true, open: true, saved: false });
    expect(aroundExcel(path, x, false, () => x.calls.push("build")).exit).toBe(EXIT_UNSAVED);
    expect(x.calls).toEqual(["state"]);
    const y = fakeExcel({ running: true, open: true, saved: false });
    aroundExcel(path, y, true, () => y.calls.push("build"));
    expect(y.calls).toEqual(["state", "close", "build", "open"]);
  });

  it("opens a workbook Excel did not have open; reports a failed open", () => {
    const x = fakeExcel({ running: false, open: false }, false);
    const r = aroundExcel(path, x, false, () => x.calls.push("build"));
    expect(x.calls).toEqual(["state", "build", "open"]);
    expect(r.exit).toBe(EXIT_REOPEN_FAILED);
    expect(r.reopen.opened).toBe("repair prompt");
  });

  it("needs desktop Excel", () => {
    expect(aroundExcel(path, undefined, false, () => 0).exit).toBe(2);
    expect(coreExcelControl("linux", () => ({ out: "", timedOut: false }))).toBeUndefined();
  });

  it("the core's control reads the scripts' answers", () => {
    const answers: Record<string, string> = { state: "unsaved", close: "closed", open: "missing" };
    const seen: string[][] = [];
    const c = coreExcelControl("win32", (program, _script, args) => (seen.push([program, ...args]), { out: answers[args[1]!]!, timedOut: false }))!;
    expect(c.state("C:\\m\\a.xlsx")).toEqual({ running: true, open: true, saved: false });
    expect(c.close("C:\\m\\a.xlsx").ok).toBe(true);
    expect(c.open("C:\\m\\a.xlsx").ok).toBe(false);
    expect(seen[0]).toEqual(["powershell", "-Action", "state", "-Path", "C:\\m\\a.xlsx"]);
    const mac = coreExcelControl("darwin", (program, _s, args) => (seen.push([program, ...args]), { out: "notrunning", timedOut: false }))!;
    expect(mac.state("/x/b.xlsx")).toEqual({ running: false, open: false });
    expect(seen.at(-1)).toEqual(["osascript", "/x/b.xlsx", "b.xlsx"]);
    const slow = coreExcelControl("darwin", () => ({ out: "", timedOut: true }))!;
    expect(slow.open("/x/b.xlsx").message).toMatch(/within 90 s/);
  });

  it("M3d: an open is checked again after a pause; a dialog still up, a repair log, or a workbook not there is a failure, never a success", () => {
    // The second look times out when a dialog blocks Excel (AppleScript error -1712).
    const blocked = coreExcelControl("darwin", () => ({ out: "", timedOut: false, error: "execution error: Microsoft Excel got an error: AppleEvent timed out. (-1712)" }))!;
    const b = blocked.open("/x/b.xlsx");
    expect(b.ok).toBe(false);
    expect(b.message).toMatch(/a dialog is waiting in Excel/);
    // Opened (the repaired copy has the same name), but Excel wrote a repair log meanwhile.
    const repaired = coreExcelControl("darwin", () => ({ out: "opened", timedOut: false }), () => ["/tmp/Repair Result to b0.xml"])!;
    expect(repaired.open("/x/b.xlsx")).toEqual({ ok: false, message: "Excel repaired the file on opening (log: /tmp/Repair Result to b0.xml): restore the backup and report this" });
    const missing = coreExcelControl("darwin", () => ({ out: "missing", timedOut: false }))!;
    expect(missing.open("/x/b.xlsx").ok).toBe(false);
    // The script itself waits, then looks for the workbook by its full path under a short time limit.
    let script = "";
    coreExcelControl("darwin", (_p, s) => ((script = s), { out: "opened", timedOut: false }))!.open("/x/b.xlsx");
    expect(script).toContain("delay 2");
    expect(script).toMatch(/with timeout of 15 seconds[\s\S]*full name of workbook i\) is p/);
  });

  it("M3e: after a verified open Excel comes to the front (Mac: activate after the check; Windows: the window, also when already open)", () => {
    let mac = "";
    coreExcelControl("darwin", (_p, s) => ((mac = s), { out: "opened", timedOut: false }))!.open("/x/b.xlsx");
    // The activation sits after the full-path check, never before it.
    expect(mac).toMatch(/full name of workbook i\) is p then\s+(--[^\n]*\n\s*)*activate\s+return "opened"/);
    expect(mac.indexOf("activate")).toBeGreaterThan(mac.indexOf("delay 2"));
    let win = "";
    coreExcelControl("win32", (_p, s) => ((win = s), { out: "opened", timedOut: false }))!.open("C:\\m\\a.xlsx");
    expect(win).toContain("function Show-Book");
    expect(win).toContain("$app.Visible = $true");
    expect(win).toMatch(/AppActivate/);
    expect(win.match(/Show-Book \$wb; 'opened'/g)?.length).toBe(2);
  });
});
