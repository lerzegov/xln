// The marks in `xln.config.json` (feedback 2026-10-07): a setting the project cannot use
// (unknown, misplaced, of the wrong type) is a warning on its key, saying where it
// belongs; a setting that no longer exists is information. Plain TypeScript for Vitest.

import { configKeyRange, LineIndex, parseConfig } from "@xln/core";

export interface ConfigMark {
  /** 0-based lines and characters. */
  start: { line: number; character: number };
  end: { line: number; character: number };
  message: string;
  severity: "warning" | "info";
}

export function configMarks(text: string): ConfigMark[] {
  const lines = new LineIndex(text);
  return parseConfig(text).issues.map((i) => {
    // The deepest key of the path the text has (a problem about a value marks its key).
    let r: { start: number; end: number } | undefined;
    for (let n = i.path.length; n > 0 && !r; n--) r = configKeyRange(text, i.path.slice(0, n));
    const range = r ?? { start: 0, end: Math.min(1, text.length) };
    return { start: lines.position(range.start), end: lines.position(range.end), message: i.message, severity: i.kind === "problem" ? "warning" : "info" };
  });
}
