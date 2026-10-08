// The MCP server in process: a client over the SDK's in-memory transport calls each tool on
// copies of the probe workbooks in a temporary root.
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Client } from "@modelcontextprotocol/client";
import { InMemoryTransport } from "@modelcontextprotocol/server";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { readWorkbook } from "@xln/core";
import { check } from "@xln/cli";
import { createServer } from "../src/server.js";
import { parseArgs } from "../src/main.js";

const RESULTS = join(import.meta.dirname, "..", "..", "..", "probes", "results");
const root = mkdtempSync(join(tmpdir(), "xln-mcp-"));
const outside = mkdtempSync(join(tmpdir(), "xln-mcp-outside-"));
let client: Client;

interface Result {
  isError?: boolean;
  content: { type: string; text?: string }[];
  structuredContent?: Record<string, unknown>;
}

async function call(name: string, args: Record<string, unknown>): Promise<Result> {
  return (await client.callTool({ name, arguments: args })) as Result;
}

const text = (r: Result) => r.content.map((c) => c.text ?? "").join("\n");

function copy(fixture: string, as: string): string {
  const p = join(root, as);
  copyFileSync(join(RESULTS, fixture), p);
  return p;
}

beforeAll(async () => {
  const server = createServer({ roots: [root] });
  const [a, b] = InMemoryTransport.createLinkedPair();
  await server.connect(b);
  client = new Client({ name: "test", version: "1.0.0" });
  await client.connect(a);
});

afterAll(async () => {
  await client.close();
  rmSync(root, { recursive: true, force: true });
  rmSync(outside, { recursive: true, force: true });
});

describe("xln-mcp", () => {
  it("lists six tools; only pull and build write, and build needs confirm", async () => {
    const { tools } = await client.listTools();
    expect(tools.map((t) => t.name).sort()).toEqual(["xln_build", "xln_build_plan", "xln_check", "xln_lib_status", "xln_names", "xln_pull"]);
    const by = Object.fromEntries(tools.map((t) => [t.name, t]));
    for (const n of ["xln_check", "xln_names", "xln_build_plan", "xln_lib_status"]) expect(by[n]!.annotations?.readOnlyHint, n).toBe(true);
    expect(by["xln_build"]!.annotations?.destructiveHint).toBe(true);
    expect(by["xln_build"]!.inputSchema.required).toEqual(expect.arrayContaining(["workbook", "confirm"]));
  });

  it("xln_check gives the CLI's findings and a verdict; a project folder adds the source findings", async () => {
    const wb = copy("probe_win.xlsx", "checked.xlsx");
    const r = await call("xln_check", { path: "checked.xlsx" });
    expect(r.isError).toBeFalsy();
    const s = r.structuredContent!;
    const cli = check({ workbook: wb, json: true });
    expect(s["verdict"]).toBe("errors");
    expect(s["ok"]).toBe(false);
    expect(s["counts"]).toEqual(cli.counts);
    expect(s["findings"]).toEqual(JSON.parse(JSON.stringify(cli.findings)));
    expect(s["findingsTotal"]).toBe(cli.findings.length);
    expect(s["source"]).toBeUndefined();
    expect(text(r)).toMatch(/^xln check checked\.xlsx: 2 errors, 4 warnings \(exit code 1\)/);
    // The census is summarised unless asked for in full.
    expect((s["census"] as { total: number }).total).toBe(10);
    expect(s["spills"]).toBeUndefined();
    const full = await call("xln_check", { path: wb, detail: "full", only: ["C2"] });
    expect(full.structuredContent!["spills"]).toBeDefined();
    expect(full.structuredContent!["checks"]).toContain("C2");
    expect(full.structuredContent!["checks"]).not.toContain("C3");

    expect((await call("xln_pull", { workbook: "checked.xlsx" })).isError).toBeFalsy();
    const p = await call("xln_check", { path: "checked.xln" });
    expect(p.isError).toBeFalsy();
    expect((p.structuredContent!["source"] as { project: string }).project).toBe(join(root, "checked.xln"));

    const bad = await call("xln_check", { path: wb, only: ["C99"] });
    expect(bad.isError).toBe(true);
    expect(text(bad)).toContain("--only takes checks C1 to C13");
  });

  it("xln_check on a clean workbook: verdict clean, ok", async () => {
    copy("f7_base.xlsx", "clean.xlsx");
    const r = await call("xln_check", { path: "clean.xlsx" });
    expect(r.structuredContent!["verdict"]).toBe("clean");
    expect(r.structuredContent!["ok"]).toBe(true);
  });

  it("refuses paths outside the roots, and missing files, as tool errors", async () => {
    const away = join(outside, "x.xlsx");
    copyFileSync(join(RESULTS, "f7_base.xlsx"), away);
    for (const [tool, args] of [
      ["xln_check", { path: away }],
      ["xln_names", { path: "../" + "x.xlsx" }],
      ["xln_pull", { workbook: away }],
      ["xln_build_plan", { workbook: away }],
    ] as const) {
      const r = await call(tool, args);
      expect(r.isError, tool).toBe(true);
      expect(text(r), tool).toMatch(/outside the allowed roots/);
    }
    copy("f7_base.xlsx", "pulled-out.xlsx");
    const r = await call("xln_pull", { workbook: "pulled-out.xlsx", out: join(outside, "proj") });
    expect(r.isError).toBe(true);
    expect(existsSync(join(outside, "proj"))).toBe(false);
    // A link inside the root that leads out of it is outside.
    symlinkSync(outside, join(root, "link"));
    const linked = await call("xln_check", { path: "link/x.xlsx" });
    expect(linked.isError).toBe(true);
    expect(text(linked)).toMatch(/outside the allowed roots/);
    const missing = await call("xln_check", { path: "nope.xlsx" });
    expect(missing.isError).toBe(true);
    expect(text(missing)).toBe("no such file: nope.xlsx");
  });

  it("xln_names lists and filters a workbook's names and a project's source", async () => {
    copy("probe_win.xlsx", "names.xlsx");
    const all = await call("xln_names", { path: "names.xlsx" });
    expect(all.structuredContent!["total"]).toBe(10);
    expect(all.structuredContent!["source"]).toBe("workbook");
    const lambdas = await call("xln_names", { path: "names.xlsx", kind: "lambda" });
    expect(lambdas.structuredContent!["matched"]).toBe(7);
    const one = await call("xln_names", { path: "names.xlsx", names: ["p_add1"] });
    expect(one.structuredContent!["names"]).toEqual([
      { name: "P_Add1", key: "P_Add1", scope: null, kind: "lambda", definition: "LAMBDA(x, x+1)", doc: "probe comment", hidden: false, cell: null, params: ["x"], file: "names/P.xln:4" },
    ]);
    const local = await call("xln_names", { path: "names.xlsx", scope: "S2" });
    expect(local.structuredContent!["matched"]).toBe(2);
    expect((local.structuredContent!["names"] as { scope: string }[]).every((n) => n.scope === "S2")).toBe(true);
    const page = await call("xln_names", { path: "names.xlsx", limit: 2, offset: 1 });
    expect(page.structuredContent!["returned"]).toBe(2);
    expect(text(page)).toContain("showing 2–3");

    // A project's source, unbuilt edits included.
    expect((await call("xln_pull", { workbook: "names.xlsx" })).isError).toBeFalsy();
    const p = join(root, "names.xln", "names", "P.xln");
    writeFileSync(p, readFileSync(p, "utf8").replace("P_Add1 = LAMBDA(x, x+1);", "P_Add1 = LAMBDA(x, x+2);"));
    const src = await call("xln_names", { path: "names.xln", query: "x+2" });
    expect(src.structuredContent!["source"]).toBe("project");
    expect((src.structuredContent!["names"] as { key: string }[]).map((n) => n.key)).toEqual(["P_Add1"]);
  });

  it("xln_pull writes the project; refuses over unbuilt edits (nothing written) unless discard: true", async () => {
    copy("f7_base.xlsx", "pulled.xlsx");
    const r = await call("xln_pull", { workbook: "pulled.xlsx" });
    expect(r.isError).toBeFalsy();
    expect(r.structuredContent!["out"]).toBe(join(root, "pulled.xln"));
    expect(r.structuredContent!["files"]).toEqual(expect.arrayContaining(["xln.lock.json", "workbook.manifest.json"]));
    expect(text(r)).toContain("xln pull pulled.xlsx →");

    const f = join(root, "pulled.xln", "names", "_unmanaged.xln");
    const edited = readFileSync(f, "utf8").replace(/Rate = [0-9.]+;/, "Rate = 0.5;");
    writeFileSync(f, edited);
    const refused = await call("xln_pull", { workbook: "pulled.xlsx" });
    expect(refused.isError).toBe(true);
    expect(text(refused)).toContain("refused: the project has 1 source edit not built yet");
    expect(text(refused)).toContain("discard: true");
    expect(refused.structuredContent!["refused"]).toBe(true);
    expect((refused.structuredContent!["unbuilt"] as unknown[]).length).toBe(1);
    expect(readFileSync(f, "utf8")).toBe(edited);

    const discarded = await call("xln_pull", { workbook: "pulled.xlsx", discard: true });
    expect(discarded.isError).toBeFalsy();
    expect((discarded.structuredContent!["discarded"] as unknown[]).length).toBe(1);
    expect(readFileSync(f, "utf8")).not.toBe(edited);
  });

  it("xln_build_plan shows the change set; xln_build writes it only with confirm and the same plan, keeping the backup", async () => {
    const wb = copy("f7_base.xlsx", "built.xlsx");
    expect((await call("xln_pull", { workbook: "built.xlsx" })).isError).toBeFalsy();
    const f = join(root, "built.xln", "names", "_unmanaged.xln");
    writeFileSync(f, readFileSync(f, "utf8").replace(/Rate = [0-9.]+;/, "Rate = 0.25;"));
    const before = readFileSync(wb);

    const plan = await call("xln_build_plan", { workbook: "built.xlsx" });
    expect(plan.isError).toBeFalsy();
    const ps = plan.structuredContent!;
    expect(ps["willWrite"]).toBe(true);
    expect(ps["excelOpen"]).toBe(false);
    const changes = (ps["changeSet"] as { changes: { op: string; name?: string }[] }).changes;
    expect(changes.some((c) => c.op === "set-name" && c.name === "Rate")).toBe(true);
    expect(text(plan)).toContain("dry run: nothing written");
    expect(readFileSync(wb).equals(before)).toBe(true);
    const id = ps["planId"] as string;

    // No confirm: the schema refuses it, nothing written.
    const noConfirm = await call("xln_build", { workbook: "built.xlsx" });
    expect(noConfirm.isError).toBe(true);
    const falseConfirm = await call("xln_build", { workbook: "built.xlsx", confirm: false });
    expect(falseConfirm.isError).toBe(true);
    expect(readFileSync(wb).equals(before)).toBe(true);

    // The plan moved since it was shown: refused.
    const stale = await call("xln_build", { workbook: "built.xlsx", confirm: true, planId: "000000000000" });
    expect(stale.isError).toBe(true);
    expect(text(stale)).toContain("the plan changed since planId 000000000000");
    expect(readFileSync(wb).equals(before)).toBe(true);

    // Excel has it open: refused (E1).
    const lock = join(root, "~$built.xlsx");
    writeFileSync(lock, "");
    const locked = await call("xln_build", { workbook: "built.xlsx", confirm: true, planId: id });
    expect(locked.isError).toBe(true);
    expect(text(locked)).toContain("Excel has built.xlsx open");
    expect(locked.structuredContent!["exit"]).toBe(3);
    expect(readFileSync(wb).equals(before)).toBe(true);
    expect((await call("xln_build_plan", { workbook: "built.xlsx" })).structuredContent!["excelOpen"]).toBe(true);
    rmSync(lock);

    const built = await call("xln_build", { workbook: "built.xlsx", confirm: true, planId: id });
    expect(built.isError, text(built)).toBeFalsy();
    expect(built.structuredContent!["status"]).toBe("built");
    expect(built.structuredContent!["backup"]).toBe(join(root, "built.backup.xlsx"));
    expect(readFileSync(join(root, "built.backup.xlsx")).equals(before)).toBe(true);
    const names = readWorkbook(new Uint8Array(readFileSync(wb))).definedNames;
    expect(names.find((n) => n.name === "Rate")!.definition).toBe("0.25");

    const again = await call("xln_build_plan", { workbook: "built.xlsx" });
    expect(again.structuredContent!["status"]).toBe("up-to-date");
    expect(again.structuredContent!["willWrite"]).toBe(false);
  });

  it("a refused plan is a result (with the reasons); building it is a tool error with the CLI's wording", async () => {
    copy("f7_base.xlsx", "refused.xlsx");
    expect((await call("xln_pull", { workbook: "refused.xlsx" })).isError).toBeFalsy();
    const f = join(root, "refused.xln", "names", "_unmanaged.xln");
    writeFileSync(f, readFileSync(f, "utf8").replace(/Rate = [0-9.]+;/, "Rate = (1 + ;"));
    const plan = await call("xln_build_plan", { workbook: "refused.xlsx" });
    expect(plan.isError).toBeFalsy();
    expect(plan.structuredContent!["status"]).toBe("refused");
    expect(plan.structuredContent!["willWrite"]).toBe(false);
    expect(text(plan)).toContain("build refused: nothing written");
    const build = await call("xln_build", { workbook: "refused.xlsx", confirm: true });
    expect(build.isError).toBe(true);
    expect(text(build)).toContain("xln build refused.xlsx: build refused: nothing written");
  });

  it("xln_build removes the @renamed line of the rename it built, as xln build does", async () => {
    copy("f7_base.xlsx", "ren.xlsx");
    expect((await call("xln_pull", { workbook: "ren.xlsx" })).isError).toBeFalsy();
    const f = join(root, "ren.xln", "names", "_unmanaged.xln");
    const pulledText = readFileSync(f, "utf8");
    writeFileSync(f, pulledText.replace("RateX = 0.5;", "@renamed(RateX)\nRateY = 0.5;"));
    const s1 = join(root, "ren.xln", "names", "sheets", "S1.xln");
    writeFileSync(s1, readFileSync(s1, "utf8").replace("RateX+Rate", "RateY+Rate"));
    const plan = await call("xln_build_plan", { workbook: "ren.xlsx" });
    expect(readFileSync(f, "utf8")).toContain("@renamed(RateX)");
    const built = await call("xln_build", { workbook: "ren.xlsx", confirm: true, planId: plan.structuredContent!["planId"] });
    expect(built.isError, text(built)).toBeFalsy();
    expect(text(built)).toContain("removed @renamed(RateX) from names/_unmanaged.xln: the rename is built");
    expect(built.structuredContent!["renamedRemoved"]).toEqual(["removed @renamed(RateX) from names/_unmanaged.xln: the rename is built"]);
    expect(readFileSync(f, "utf8")).toBe(pulledText.replace("RateX = 0.5;", "RateY = 0.5;"));
    const pull = await call("xln_pull", { workbook: "ren.xlsx" });
    expect(pull.isError, text(pull)).toBeFalsy();
  });

  it("xln_build with out writes a copy and leaves the original and the lockfile alone", async () => {
    const wb = copy("f7_base.xlsx", "orig.xlsx");
    expect((await call("xln_pull", { workbook: "orig.xlsx" })).isError).toBeFalsy();
    const f = join(root, "orig.xln", "names", "_unmanaged.xln");
    writeFileSync(f, readFileSync(f, "utf8").replace(/Rate = [0-9.]+;/, "Rate = 0.75;"));
    const lock = readFileSync(join(root, "orig.xln", "xln.lock.json"), "utf8");
    const before = readFileSync(wb);
    const r = await call("xln_build", { workbook: "orig.xlsx", out: "copy.xlsx", confirm: true });
    expect(r.isError, text(r)).toBeFalsy();
    expect(readFileSync(wb).equals(before)).toBe(true);
    expect(readFileSync(join(root, "orig.xln", "xln.lock.json"), "utf8")).toBe(lock);
    expect(readWorkbook(new Uint8Array(readFileSync(join(root, "copy.xlsx")))).definedNames.find((n) => n.name === "Rate")!.definition).toBe("0.75");
    const away = await call("xln_build", { workbook: "orig.xlsx", out: join(outside, "copy.xlsx"), confirm: true });
    expect(away.isError).toBe(true);
    expect(existsSync(join(outside, "copy.xlsx"))).toBe(false);
  });

  it("xln_lib_status compares with a library folder", async () => {
    copy("probe_win.xlsx", "lib.xlsx");
    const lib = join(root, "lib");
    mkdirSync(lib);
    writeFileSync(join(lib, "P_Add1.lambda"), "# name       P_Add1\n# summary    Adds one.\n# params     x\n\nLAMBDA(x, x + 1)\n");
    writeFileSync(join(lib, "P_Gone.lambda"), "# name       P_Gone\n# summary    Missing here.\n# params     x\n\nLAMBDA(x, x)\n");
    const r = await call("xln_lib_status", { path: "lib.xlsx", lib: "lib" });
    expect(r.isError, text(r)).toBeFalsy();
    const counts = r.structuredContent!["counts"] as Record<string, number>;
    expect(counts["missing"]).toBe(1);
    expect(counts["identical"]).toBe(1);
    expect(text(r)).toContain("xln lib status: lib.xlsx (workbook)");
    const none = await call("xln_lib_status", { path: "lib.xlsx" });
    expect(none.isError).toBe(true);
    expect(text(none)).toContain("no library: pass --lib <dir>");
  });

  it("parses --root", () => {
    expect(parseArgs(["--root", "/a", "--root=/b"])).toEqual({ roots: ["/a", "/b"] });
    expect(parseArgs([])).toEqual({ roots: [process.cwd()] });
    expect(parseArgs(["--root"])).toEqual({ error: "--root needs a folder" });
    expect(parseArgs(["x"])).toEqual({ error: "unknown argument 'x'" });
  });
});
