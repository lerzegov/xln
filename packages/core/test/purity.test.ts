// packages/core must run in vscode.dev: no Node built-ins anywhere in its sources.
import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";

const SRC = join(import.meta.dirname, "..", "src");
const FORBIDDEN = /from\s+["'](node:[^"']+|fs|path|os|child_process|crypto|stream|buffer|url|util|zlib)["']|\brequire\(|\bBuffer\b|\bprocess\./;

function files(dir: string): string[] {
  return readdirSync(dir).flatMap((n) => {
    const p = join(dir, n);
    return statSync(p).isDirectory() ? files(p) : p.endsWith(".ts") ? [p] : [];
  });
}

describe("core purity", () => {
  it("imports nothing from Node", () => {
    const bad = files(SRC).filter((f) => FORBIDDEN.test(readFileSync(f, "utf8")));
    expect(bad).toEqual([]);
  });
});
