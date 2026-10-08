// The xln output channel, one writer for every feature. vscode.dev has no other place where
// the author can read what an action did (feedback 2026-10-07: only pulls left a trace), so
// every command and action writes here, in the pull summary's style: a header line
// `HH:MM:SS xln <verb> <target>…` flush left, then its details indented by two spaces.
//
// A header is any line that starts with `xln ` (the build and pull summaries come from
// formatters shared with the CLI, whose first line reads `xln build …`, `xln pull …`); the
// local time goes in front of it here, so the shared formatters stay as the CLI prints them.

import * as vscode from "vscode";

type Sink = (line: string) => void;

let sink: Sink | undefined;
/** The lines written, as the channel shows them, the last KEEP (tests read them: an OutputChannel cannot be read back). */
const kept: string[] = [];
const KEEP = 5000;

export function setLogSink(s: Sink | undefined): void {
  sink = s;
}

export function loggedLines(): string[] {
  return [...kept];
}

/** Local time `HH:MM:SS`. */
export function clock(d = new Date()): string {
  const p = (n: number) => String(n).padStart(2, "0");
  return `${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

/** One line to the channel; a header line gets the time in front. */
export function logLine(line: string): void {
  const out = line.startsWith("xln ") ? `${clock()} ${line}` : line;
  kept.push(out);
  if (kept.length > KEEP) kept.splice(0, kept.length - KEEP);
  sink?.(out);
}

type Show = (message: string, ...rest: unknown[]) => Thenable<string | undefined>;

/**
 * One user action in the log. The header is written on the first line (so an action that
 * ends before it learns its target still gets one), after a blank line. The message
 * helpers log what the notification says, then show it: an early return leaves the same
 * trace in the channel as on the screen.
 */
export class Activity {
  private headed = false;
  private readonly t0 = Date.now();

  constructor(
    readonly verb: string,
    public target?: string,
  ) {}

  /** Writes the header now, once; `rest` follows the target (`: 3 changes`). */
  head(rest = ""): this {
    if (!this.headed) {
      this.headed = true;
      logLine("");
      logLine(`xln ${this.verb}${this.target ? ` ${this.target}` : ""}${rest}`);
    }
    return this;
  }

  /** Lines from a shared formatter whose first line is the header (`xln build …`). */
  summary(lines: readonly string[]): void {
    if (this.headed) {
      for (const l of lines) logLine(l.startsWith("xln ") ? `  ${l}` : l);
      return;
    }
    this.headed = true;
    logLine("");
    for (const l of lines) logLine(l);
  }

  line(s: string): void {
    this.head();
    logLine(`  ${s}`);
  }

  lines(ls: readonly string[]): void {
    for (const l of ls) this.line(l);
  }

  /** Milliseconds since the action began. */
  ms(): number {
    return Date.now() - this.t0;
  }

  info(message: string, ...rest: unknown[]): Thenable<string | undefined> {
    return this.say("info", message, rest, vscode.window.showInformationMessage as Show);
  }

  warn(message: string, ...rest: unknown[]): Thenable<string | undefined> {
    return this.say("warning", message, rest, vscode.window.showWarningMessage as Show);
  }

  error(message: string, ...rest: unknown[]): Thenable<string | undefined> {
    return this.say("error", message, rest, vscode.window.showErrorMessage as Show);
  }

  private say(kind: string, message: string, rest: unknown[], show: Show): Thenable<string | undefined> {
    const text = message.startsWith("xln: ") ? message.slice(5) : message;
    this.line(kind === "info" ? text : `${kind}: ${text}`);
    return show(message, ...rest);
  }
}
