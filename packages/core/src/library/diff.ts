// A line diff for the library's reports: two definitions are re-laid out the same way
// first (`prettyPrint`), so the diff shows what changed and not how each side was spaced.

import { prettyPrint } from "../lang/format.js";

export interface DiffLine {
  op: " " | "-" | "+";
  text: string;
}

/** Longest common subsequence over lines: fine for definitions (tens of lines). */
export function lineDiff(a: readonly string[], b: readonly string[]): DiffLine[] {
  const n = a.length;
  const m = b.length;
  const lcs: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) lcs[i]![j] = a[i] === b[j] ? lcs[i + 1]![j + 1]! + 1 : Math.max(lcs[i + 1]![j]!, lcs[i]![j + 1]!);
  const out: DiffLine[] = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) {
      out.push({ op: " ", text: a[i]! });
      i++;
      j++;
    } else if (lcs[i + 1]![j]! >= lcs[i]![j + 1]!) out.push({ op: "-", text: a[i++]! });
    else out.push({ op: "+", text: b[j++]! });
  }
  while (i < n) out.push({ op: "-", text: a[i++]! });
  while (j < m) out.push({ op: "+", text: b[j++]! });
  return out;
}

/** A formula laid out for comparison: pretty-printed when it parses, else as written. */
export function comparableLayout(display: string, width = 80): string[] {
  try {
    return prettyPrint(display, { width }).split("\n");
  } catch {
    return display.split("\n").map((l) => l.trimEnd());
  }
}

/** The diff of two definitions in display form, `-` for `from`, `+` for `to`. */
export function definitionDiff(from: string, to: string): DiffLine[] {
  return lineDiff(comparableLayout(from), comparableLayout(to));
}

export function diffText(lines: readonly DiffLine[], indent = ""): string {
  return lines.map((l) => `${indent}${l.op} ${l.text}`.trimEnd()).join("\n");
}
