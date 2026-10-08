// `--reopen` (E7): close the workbook in desktop Excel without saving, write it, open it
// again. The file on disk is what xln builds from, so edits Excel has not saved would be
// lost by the close: refused unless the author passes `--discard`.

import { readdirSync } from "node:fs";
import { basename, dirname, resolve } from "node:path";
import { isLocked } from "@xln/core";
import type { ExcelControl, ExcelWorkbookState } from "./excel.js";

export interface ReopenReport {
  /** The workbook Excel was asked about. */
  path: string;
  before?: ExcelWorkbookState;
  closed?: string;
  opened?: string;
  /** Why nothing was done. */
  refused?: string;
  /** The open failed (a repair prompt, a dialog, an error). */
  openFailed?: boolean;
}

export const EXIT_UNSAVED = 5;
export const EXIT_REOPEN_FAILED = 6;

function sleep(ms: number): void {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}

/** Waits until Excel's `~$` owner file is gone (Excel removes it after closing). */
function waitUnlocked(path: string, timeoutMs: number): boolean {
  const end = Date.now() + timeoutMs;
  for (;;) {
    if (!isLocked(path, readdirSync(dirname(path)))) return true;
    if (Date.now() > end) return false;
    sleep(200);
  }
}

/**
 * Runs `work` between closing `path` in Excel and opening it again. Returns `work`'s result
 * (undefined when refused before it ran) and what happened in Excel.
 */
export function aroundExcel<T>(path: string, control: ExcelControl | undefined, discard: boolean, work: () => T): { result: T | undefined; reopen: ReopenReport; exit?: number } {
  const target = resolve(path);
  const reopen: ReopenReport = { path: target };
  if (!control) {
    reopen.refused = "--reopen needs desktop Excel (macOS or Windows)";
    return { result: undefined, reopen, exit: 2 };
  }
  const state = control.state(target);
  reopen.before = state;
  if (state.open && state.saved === false && !discard) {
    reopen.refused = `Excel has unsaved changes in ${basename(target)}: save them in Excel (then pull, and build again), or pass --discard to lose them. Nothing written.`;
    return { result: undefined, reopen, exit: EXIT_UNSAVED };
  }
  if (state.open) {
    const c = control.close(target);
    reopen.closed = c.message;
    if (!c.ok) {
      reopen.refused = `${c.message}. Nothing written.`;
      return { result: undefined, reopen, exit: 3 };
    }
    if (!waitUnlocked(target, 10_000)) {
      reopen.refused = `Excel closed ${basename(target)} but its owner file is still there. Nothing written.`;
      return { result: undefined, reopen, exit: 3 };
    }
  }
  const result = work();
  const o = control.open(target);
  reopen.opened = o.message;
  if (!o.ok) {
    reopen.openFailed = true;
    return { result, reopen, exit: EXIT_REOPEN_FAILED };
  }
  return { result, reopen };
}

export function reopenText(r: ReopenReport): string {
  const out: string[] = [];
  if (r.before) out.push(`  Excel: ${r.before.open ? `had it open (${r.before.saved ? "saved" : "unsaved changes"})` : r.before.running ? "did not have it open" : "was not running"}`);
  if (r.closed) out.push(`  Excel: ${r.closed}`);
  if (r.opened) out.push(`  Excel: ${r.opened}`);
  if (r.refused) out.push(`  ${r.refused}`);
  return out.join("\n") + (out.length ? "\n" : "");
}
