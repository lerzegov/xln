/// <reference types="node" />
// Desktop Excel control for "Build and reopen in Excel" (E7): the core's scripts run with
// `osascript` (macOS) or PowerShell (Windows, untested). Only the desktop bundle contains
// this module (scripts/build.mjs swaps it in for excelHost.ts).
import { spawnSync } from "node:child_process";
import { existsSync, mkdtempSync, readdirSync, statSync, writeFileSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join, parse as parsePath } from "node:path";
import { excelControl, type ExcelControl, type ExcelExec } from "@xln/core";

const exec: ExcelExec = (program, script, args, timeoutMs) => {
  let r;
  if (program === "osascript") {
    r = spawnSync("osascript", ["-", ...args], { input: script, encoding: "utf8", timeout: timeoutMs });
  } else {
    const file = join(mkdtempSync(join(tmpdir(), "xln-")), "excel.ps1");
    writeFileSync(file, script, "utf8");
    r = spawnSync("powershell.exe", ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", file, ...args], { encoding: "utf8", timeout: timeoutMs });
  }
  if (r.error && (r.error as NodeJS.ErrnoException).code === "ETIMEDOUT") return { out: "", timedOut: true };
  if (r.error) return { out: "", timedOut: false, error: r.error.message };
  const out = (r.stdout ?? "").trim();
  if (r.status !== 0) return { out, timedOut: false, error: (r.stderr ?? "").trim() || `${program} exited with ${r.status}` };
  return { out, timedOut: false };
};

// Excel for Mac writes a log here when someone answers its repair prompt with Yes (F8).
const REPAIR_DIR = join(homedir(), "Library/Containers/com.microsoft.Excel/Data/tmp");

function repairLogsSince(path: string, since: number): string[] {
  const stem = parsePath(path).name;
  if (!existsSync(REPAIR_DIR)) return [];
  return readdirSync(REPAIR_DIR)
    .filter((f) => f.startsWith(`Repair Result to ${stem}`) && f.endsWith(".xml"))
    .filter((f) => statSync(join(REPAIR_DIR, f)).mtimeMs >= since)
    .map((f) => join(REPAIR_DIR, f));
}

/** The home folder, for `~` in a library path (M4). */
export function homeDir(): string | undefined {
  return homedir();
}

export function excelHost(): ExcelControl | undefined {
  return excelControl(process.platform, exec, repairLogsSince);
}
