// Diagnostics with positions (requirement C1). Excel itself only says yes or no
// (probe T11), so every message here must say where and what.

import type { Span } from "./tokens.js";

export type Severity = "error" | "warning";

export interface Diagnostic extends Span {
  severity: Severity;
  /** Stable identifier for tooling: `syntax`, `unknown-function`, `bare-prefix`, … */
  code: string;
  message: string;
}

/** 1-based line and column of an offset; CR LF counts as one line break. */
export function lineCol(src: string, offset: number): { line: number; col: number } {
  let line = 1;
  let col = 1;
  for (let k = 0; k < offset && k < src.length; k++) {
    const c = src[k];
    if (c === "\n") {
      line++;
      col = 1;
    } else if (c === "\r") {
      if (src[k + 1] !== "\n") {
        line++;
        col = 1;
      }
    } else col++;
  }
  return { line, col };
}

/** `line 2, column 5: expected ')'` followed by the offending line and a caret. */
export function formatDiagnostic(src: string, d: Diagnostic): string {
  const { line, col } = lineCol(src, d.start);
  const lines = src.split(/\r\n|\r|\n/);
  const text = lines[line - 1] ?? "";
  const width = Math.max(1, Math.min(d.end - d.start, text.length - col + 1));
  return `line ${line}, column ${col}: ${d.message}\n  ${text}\n  ${" ".repeat(col - 1)}${"^".repeat(width)}`;
}

/** Thrown by `parse`, `compile` and `decompile` when the input cannot be handled. */
export class FormulaError extends Error {
  readonly diagnostics: Diagnostic[];
  readonly src: string;
  constructor(src: string, diagnostics: Diagnostic[]) {
    const first = diagnostics.find((d) => d.severity === "error") ?? diagnostics[0];
    super(first ? formatDiagnostic(src, first) : "formula error");
    this.name = "FormulaError";
    this.src = src;
    this.diagnostics = diagnostics;
  }
  get start(): number {
    return this.diagnostics[0]?.start ?? 0;
  }
}
