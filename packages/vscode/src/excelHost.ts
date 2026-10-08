// Desktop Excel control for "Build and reopen in Excel" (E7). This is the web build's
// version: vscode.dev cannot run programs, so there is no Excel to drive. The desktop
// bundle replaces this module with excelHost.node.ts (scripts/build.mjs).
import type { ExcelControl } from "@xln/core";

export function excelHost(): ExcelControl | undefined {
  return undefined;
}

/** The home folder, for `~` in a library path (M4): none in the browser. */
export function homeDir(): string | undefined {
  return undefined;
}
