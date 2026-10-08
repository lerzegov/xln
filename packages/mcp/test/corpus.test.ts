// The MCP tools on the real workbooks (XLN_CORPUS=<folder>, workbooks at */dist/*.xlsx),
// each copied into a temporary root first: the corpus is only read. A check's summary must
// stay small enough for an agent's context, and a fresh pull must plan as up to date.
import { copyFileSync, existsSync, mkdtempSync, readdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import { Client } from "@modelcontextprotocol/client";
import { InMemoryTransport } from "@modelcontextprotocol/server";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { check } from "@xln/cli";
import { createServer } from "../src/server.js";

const CORPUS = process.env["XLN_CORPUS"];

function workbooks(root: string): string[] {
  const out: string[] = [];
  for (const d of readdirSync(root, { withFileTypes: true })) {
    if (!d.isDirectory()) continue;
    const dist = join(root, d.name, "dist");
    if (!existsSync(dist)) continue;
    for (const f of readdirSync(dist)) if (f.endsWith(".xlsx") && !f.startsWith("~$")) out.push(join(dist, f));
  }
  return out.sort();
}

interface Result {
  isError?: boolean;
  content: { type: string; text?: string }[];
  structuredContent?: Record<string, unknown>;
}

describe.skipIf(!CORPUS)("xln-mcp on the corpus (XLN_CORPUS)", () => {
  const root = mkdtempSync(join(tmpdir(), "xln-mcp-corpus-"));
  let client: Client;
  const call = async (name: string, args: Record<string, unknown>) => (await client.callTool({ name, arguments: args })) as Result;

  beforeAll(async () => {
    const server = createServer({ roots: [root] });
    const [a, b] = InMemoryTransport.createLinkedPair();
    await server.connect(b);
    client = new Client({ name: "corpus", version: "1.0.0" });
    await client.connect(a);
  });
  afterAll(async () => {
    await client.close();
    rmSync(root, { recursive: true, force: true });
  });

  it("check, names, pull and an up-to-date plan on every workbook", { timeout: 600_000 }, async () => {
    const all = workbooks(CORPUS!);
    expect(all.length).toBeGreaterThan(0);
    for (const src of all) {
      const name = basename(src);
      copyFileSync(src, join(root, name));
      const c = await call("xln_check", { path: name });
      expect(c.isError, name).toBeFalsy();
      expect(c.structuredContent!["counts"], name).toEqual(check({ workbook: join(root, name), json: true }).counts);
      // The summary mode keeps the result an agent reads small (the full census runs to 250 kB).
      expect(JSON.stringify(c.structuredContent).length, name).toBeLessThan(80_000);

      const p = await call("xln_pull", { workbook: name });
      expect(p.isError, name).toBeFalsy();
      const total = (p.structuredContent!["report"] as { names: number }).names;
      const n = await call("xln_names", { path: name, limit: 1 });
      expect(n.structuredContent!["total"], name).toBe(total);

      const plan = await call("xln_build_plan", { workbook: name });
      expect(plan.isError, name).toBeFalsy();
      // A fresh pull plans no edits: at most the provenance tags a first build adds (D6).
      expect(plan.structuredContent!["status"], name).not.toBe("refused");
      expect(plan.structuredContent!["conflicts"], name).toEqual([]);
      const changes = (plan.structuredContent!["changeSet"] as { changes: { op: string; fields?: string[] }[] }).changes;
      const edits = changes.filter((ch) => ch.op !== "set-embedded-source" && !(ch.op === "set-name" && (ch.fields ?? []).every((f) => f === "comment" || f === "provenance")));
      expect(edits, `${name}: ${plan.content[0]!.text}`).toEqual([]);
    }
  });
});
