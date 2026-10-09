// The MCP server in process: a client over the SDK's in-memory transport calls each tool on
// copies of the probe workbooks in a temporary root.
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Client } from "@modelcontextprotocol/client";
import { InMemoryTransport } from "@modelcontextprotocol/server";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { readWorkbook } from "@xln/core";
import { check, formulas, graphSummary } from "@xln/cli";
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
  it("lists thirteen tools; seven only read, build needs confirm, publish is destructive", async () => {
    const { tools } = await client.listTools();
    expect(tools.map((t) => t.name).sort()).toEqual([
      "xln_build",
      "xln_build_plan",
      "xln_check",
      "xln_formulas",
      "xln_graph",
      "xln_lib_base",
      "xln_lib_publish",
      "xln_lib_status",
      "xln_lib_take",
      "xln_names",
      "xln_pull",
      "xln_rename",
      "xln_verify",
    ]);
    const by = Object.fromEntries(tools.map((t) => [t.name, t]));
    const readOnly = ["xln_check", "xln_names", "xln_build_plan", "xln_lib_status", "xln_formulas", "xln_graph", "xln_verify"];
    for (const t of tools) expect(t.annotations?.readOnlyHint, t.name).toBe(readOnly.includes(t.name));
    expect(by["xln_build"]!.annotations?.destructiveHint).toBe(true);
    expect(by["xln_lib_publish"]!.annotations?.destructiveHint).toBe(true);
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

  it("xln_formulas: the formula view, compact and paged, filtered by sheet, query and names", async () => {
    const wb = copy("f7_base.xlsx", "formulas.xlsx");
    const cli = formulas({ workbook: wb, json: true });
    const all = cli.sheets.flatMap((s) => s.lines);
    const r = await call("xln_formulas", { workbook: "formulas.xlsx" });
    expect(r.isError, text(r)).toBeFalsy();
    const s = r.structuredContent!;
    expect(s["total"]).toBe(28);
    expect(s["sheets"]).toEqual({ S1: 22, S2: 6 });
    expect(s["order"]).toBe("appearance");
    const rows = s["lines"] as { sheet: string; cell: string; formula: string; reads: string[]; defines: string[]; value: string | null; extent?: string }[];
    expect(rows.map((x) => x.formula)).toEqual(all.map((l) => l.formula));
    expect(rows.find((x) => x.cell === "E1#")).toMatchObject({ sheet: "S1", kind: "dynamic-array", formula: "SEQUENCE(3)*Rate", defines: ["Spl"], extent: "E1:E3", reads: ["Rate"] });
    expect(text(r)).toContain("xln formulas formulas.xlsx: 28 formulas on 2 sheets (order of appearance, sheet by sheet)");
    expect(text(r)).toContain("Spl  'S1'!E1# (E1:E3) = SEQUENCE(3)*Rate  → 0.1");

    const page = await call("xln_formulas", { workbook: "formulas.xlsx", limit: 5, offset: 20 });
    expect(page.structuredContent!["returned"]).toBe(5);
    expect(text(page)).toContain("showing 21–25 (offset 25 for more)");
    const s2 = await call("xln_formulas", { workbook: "formulas.xlsx", sheet: "s2" });
    expect(s2.structuredContent!["total"]).toBe(6);
    const q = await call("xln_formulas", { workbook: "formulas.xlsx", query: "loc*row" });
    expect(q.structuredContent!["matched"]).toBe(3);
    // Names: the formulas that read Loc, the workbook's or S2's (C5 reads both).
    const loc = await call("xln_formulas", { workbook: "formulas.xlsx", names: ["Loc", "S2!Loc"] });
    expect((loc.structuredContent!["lines"] as { cell: string }[]).map((x) => x.cell)).toEqual(["C5", "A1", "B1", "A2", "A3", "A4", "A5"]);

    const calc = await call("xln_formulas", { workbook: "formulas.xlsx", workbookWide: true, limit: 100 });
    expect(calc.structuredContent!["order"]).toBe("calculation");
    expect((calc.structuredContent!["lines"] as { level?: number }[]).every((x) => typeof x.level === "number")).toBe(true);
    const full = await call("xln_formulas", { workbook: "formulas.xlsx", detail: "full", limit: 1 });
    expect(full.structuredContent!["lines"]).toEqual(JSON.parse(JSON.stringify(all.slice(0, 1))));

    const both = await call("xln_formulas", { workbook: "formulas.xlsx", workbookWide: true, sheet: "S1" });
    expect(both.isError).toBe(true);
    expect(text(both)).toContain("exclude each other");
    const noSheet = await call("xln_formulas", { workbook: "formulas.xlsx", sheet: "Nope" });
    expect(noSheet.isError).toBe(true);
    expect(text(noSheet)).toContain("no sheet 'Nope'; the workbook has S1, S2");
  });

  it("xln_graph: the dependency summary of xln graph --json", async () => {
    const wb = copy("f7_base.xlsx", "graph.xlsx");
    const r = await call("xln_graph", { workbook: "graph.xlsx" });
    expect(r.isError, text(r)).toBeFalsy();
    const { buildMs: _a, ...want } = graphSummary({ workbook: wb, json: true }).summary;
    const { buildMs: _b, totals, ok, ...got } = r.structuredContent! as Record<string, unknown>;
    expect(got).toEqual(JSON.parse(JSON.stringify(want)));
    expect(ok).toBe(true);
    expect((totals as Record<string, number>)["flagged"]).toBe(1);
    expect(text(r)).toContain("not followed: 1 dynamic, 0 external, 0 broken");
    const away = join(outside, "g.xlsx");
    copyFileSync(join(RESULTS, "f7_base.xlsx"), away);
    expect(text(await call("xln_graph", { workbook: away }))).toMatch(/outside the allowed roots/);
  });

  it("xln_verify compares saved values with the backup or a given copy", async () => {
    copy("f7_base.xlsx", "ver.xlsx");
    const none = await call("xln_verify", { workbook: "ver.xlsx" });
    expect(none.isError).toBe(true);
    expect(text(none)).toContain(`no copy before the build at ${join(root, "ver.backup.xlsx")}: pass before: <file>`);
    copy("f7_base.xlsx", "ver.backup.xlsx");
    const same = await call("xln_verify", { workbook: "ver.xlsx" });
    expect(same.isError, text(same)).toBeFalsy();
    expect(same.structuredContent!["verdict"]).toBe("same");
    expect(same.structuredContent!["changedTotal"]).toBe(0);
    expect(same.content[0]!.text).toBe("xln verify ver.xlsx against ver.backup.xlsx: 40 cells on 2 sheets, 0 changed");

    copy("f8_base.xlsx", "v8-before.xlsx");
    copy("f8_p1_resaved.xlsx", "v8.xlsx");
    const changed = await call("xln_verify", { workbook: "v8.xlsx", before: "v8-before.xlsx", maxChanges: 10 });
    expect(changed.isError).toBeFalsy();
    const c = changed.structuredContent!;
    expect(c["verdict"]).toBe("changed");
    expect(c["ok"]).toBe(false);
    expect(c["changedTotal"]).toBe(46);
    expect((c["changed"] as unknown[]).length).toBe(10);
    expect((c["changed"] as unknown[])[0]).toEqual({ sheet: "N", cell: "B1", before: 6, after: 600 });
    expect(text(changed)).toContain("46 changed\n  N!B1: 6 → 600");

    const away = join(outside, "before.xlsx");
    copyFileSync(join(RESULTS, "f8_base.xlsx"), away);
    expect(text(await call("xln_verify", { workbook: "v8.xlsx", before: away }))).toMatch(/outside the allowed roots/);
  });

  it("xln_rename: a dry run lists every edit, a rename writes only the source; the build plan renames", async () => {
    const wb = copy("f7_base.xlsx", "rn.xlsx");
    const bytes = readFileSync(wb);
    expect((await call("xln_pull", { workbook: "rn.xlsx" })).isError).toBeFalsy();
    const files = ["names/_unmanaged.xln", "names/sheets/S1.xln", "names/sheets/S2.xln"].map((f) => join(root, "rn.xln", ...f.split("/")));
    const before = files.map((f) => readFileSync(f, "utf8"));

    const dry = await call("xln_rename", { path: "rn.xln", name: "Rate", to: "Pace", dryRun: true });
    expect(dry.isError, text(dry)).toBeFalsy();
    expect(files.map((f) => readFileSync(f, "utf8"))).toEqual(before);
    expect(text(dry)).toContain("xln rename Rate → Pace in rn.xln (dry run: nothing written)");
    expect(text(dry)).toContain("dry run: nothing written; call again without dryRun");
    const edits = dry.structuredContent!["edits"] as { path: string; label: string; line: number; was: string; text: string }[];
    expect(edits.find((e) => e.label === "name")).toMatchObject({ path: "names/_unmanaged.xln", was: "Rate", text: "Pace", line: before[0]!.split("\n").indexOf("Rate = 0.1;") + 1 });
    expect(edits.filter((e) => e.label === "reference").every((e) => e.was === "Rate" && e.text === "Pace")).toBe(true);
    expect(dry.structuredContent!["written"]).toEqual([]);

    // The workbook form names the project beside it; the workbook is only read.
    const real = await call("xln_rename", { path: "rn.xlsx", name: "Rate", to: "Pace" });
    expect(real.isError, text(real)).toBeFalsy();
    expect(real.structuredContent!["written"]).toEqual(expect.arrayContaining(["names/_unmanaged.xln", "names/sheets/S1.xln"]));
    expect(readFileSync(files[0]!, "utf8")).toContain("@renamed(Rate)\nPace = 0.1;");
    expect(readFileSync(wb).equals(bytes)).toBe(true);
    const plan = await call("xln_build_plan", { workbook: "rn.xlsx" });
    const changes = (plan.structuredContent!["changeSet"] as { changes: { op: string; to?: string }[] }).changes;
    expect(changes.some((c) => c.op === "rename-name" && c.to === "Pace")).toBe(true);

    const taken = await call("xln_rename", { path: "rn.xln", name: "Pace", to: "Loc" });
    expect(taken.isError).toBe(true);
    expect(text(taken)).toMatch(/^xln rename Pace → Loc in rn\.xln: /);
    const missing = await call("xln_rename", { path: "rn.xln", name: "Nope", to: "Other" });
    expect(missing.isError).toBe(true);
    expect(text(missing)).toContain("no name Nope in the project");
    const away = await call("xln_rename", { path: join(outside, "p.xln"), name: "A", to: "B" });
    expect(text(away)).toMatch(/outside the allowed roots/);
  });

  it("xln_lib_take refuses a copy with (possibly) its own edit unless discard: true; xln_lib_base records a base", async () => {
    copy("probe_win.xlsx", "lt.xlsx");
    expect((await call("xln_pull", { workbook: "lt.xlsx" })).isError).toBeFalsy();
    const project = join(root, "lt.xln");
    const fn = join(project, "names", "FN.xln");
    writeFileSync(fn, "/** Doubles. */\nFN.TWICE = LAMBDA(x, x * 2);\n\nFN.INC = LAMBDA(x, x + 1);\n");
    const lib = join(root, "ltlib");
    mkdirSync(lib);
    writeFileSync(join(lib, "FN.TWICE.lambda"), "# name       FN.TWICE\n# summary    Doubles.\n# params     x\n\nLAMBDA(x, x * 2)\n");
    writeFileSync(join(lib, "FN.INC.lambda"), "# name       FN.INC\n# summary    Adds two.\n# params     x\n\nLAMBDA(x, x + 2)\n");
    const cfg = join(project, "xln.config.json");
    writeFileSync(cfg, JSON.stringify({ ...JSON.parse(readFileSync(cfg, "utf8")), library: "../ltlib" }));
    const text0 = readFileSync(fn, "utf8");

    const refused = await call("xln_lib_take", { project: "lt.xln", name: "FN.INC" });
    expect(refused.isError).toBe(true);
    expect(text(refused)).toContain("refused: FN.INC records no library base");
    expect(text(refused)).toContain("Call again with discard: true (only when the user agrees)");
    expect(refused.structuredContent!["refused"]).toBeDefined();
    expect(readFileSync(fn, "utf8")).toBe(text0);

    const dry = await call("xln_lib_take", { project: "lt.xln", name: "FN.INC", discard: true, dryRun: true });
    expect(dry.isError, text(dry)).toBeFalsy();
    expect(dry.structuredContent!["written"]).toBe(false);
    expect(readFileSync(fn, "utf8")).toBe(text0);
    const taken = await call("xln_lib_take", { project: "lt.xln", name: "FN.INC", discard: true });
    expect(taken.isError, text(taken)).toBeFalsy();
    expect(taken.structuredContent!["written"]).toBe(true);
    expect(readFileSync(fn, "utf8")).toMatch(/@from\(lib #[0-9a-f]{6}\)/);
    expect(readFileSync(fn, "utf8")).toContain("LAMBDA(x, x + 2)");

    // FN.TWICE is identical without a base: record it; a second time there is nothing to record.
    const base = await call("xln_lib_base", { project: "lt.xln", name: "FN.TWICE" });
    expect(base.isError, text(base)).toBeFalsy();
    expect((base.structuredContent!["recorded"] as { name: string }[]).map((x) => x.name)).toEqual(["FN.TWICE"]);
    const again = await call("xln_lib_base", { project: "lt.xln", name: "FN.TWICE" });
    expect(again.isError).toBe(true);
    expect(text(again)).toContain("xln lib base FN.TWICE: nothing to record");
    const both = await call("xln_lib_base", { project: "lt.xln", name: "FN.TWICE", all: true });
    expect(text(both)).toContain("a name or all: true");
    const status = await call("xln_lib_status", { path: "lt.xln" });
    expect((status.structuredContent!["counts"] as Record<string, number>)["identical"]).toBe(2);
  });

  it("xln_lib_publish: a dry run shows the diff, a write needs confirm, and the library must lie under a root", async () => {
    copy("probe_win.xlsx", "lp.xlsx");
    expect((await call("xln_pull", { workbook: "lp.xlsx" })).isError).toBeFalsy();
    const project = join(root, "lp.xln");
    writeFileSync(join(project, "names", "FN.xln"), "/** Halves. */\nFN.HALF = LAMBDA(x, x / 2);\n");
    const lib = join(root, "lplib");
    mkdirSync(lib);
    const file = join(lib, "FN.HALF.lambda");

    const dry = await call("xln_lib_publish", { project: "lp.xln", name: "FN.HALF", lib: "lplib", dryRun: true });
    expect(dry.isError, text(dry)).toBeFalsy();
    expect(text(dry)).toContain(`xln lib publish FN.HALF --dry-run: would create ${file}`);
    expect(dry.structuredContent!["written"]).toBe(false);
    expect(existsSync(file)).toBe(false);
    const noConfirm = await call("xln_lib_publish", { project: "lp.xln", name: "FN.HALF", lib: "lplib" });
    expect(noConfirm.isError).toBe(true);
    expect(text(noConfirm)).toContain("pass confirm: true once they approve");
    expect(existsSync(file)).toBe(false);
    const pub = await call("xln_lib_publish", { project: "lp.xln", name: "FN.HALF", lib: "lplib", confirm: true });
    expect(pub.isError, text(pub)).toBeFalsy();
    expect(readFileSync(file, "utf8")).toContain("LAMBDA(x, x / 2)");
    expect(readFileSync(join(project, "names", "FN.xln"), "utf8")).toMatch(/@from\(lib #[0-9a-f]{6}\)/);

    // A library outside the roots, named in the config or reached through a link: read, never written.
    const away = join(outside, "lib");
    mkdirSync(away);
    const cfg = join(project, "xln.config.json");
    writeFileSync(cfg, JSON.stringify({ ...JSON.parse(readFileSync(cfg, "utf8")), library: away }));
    const status = await call("xln_lib_status", { path: "lp.xln" });
    expect(status.isError, text(status)).toBeFalsy();
    for (const args of [{}, { dryRun: true }, { confirm: true }]) {
      const r = await call("xln_lib_publish", { project: "lp.xln", name: "FN.HALF", ...args });
      expect(r.isError).toBe(true);
      expect(text(r)).toBe(`the library folder ${away} is outside the allowed roots: xln_lib_publish writes into it, so it must lie under a --root (reading a library elsewhere stays allowed)`);
    }
    symlinkSync(away, join(root, "liblink"));
    const linked = await call("xln_lib_publish", { project: "lp.xln", name: "FN.HALF", lib: "liblink", confirm: true });
    expect(linked.isError).toBe(true);
    expect(text(linked)).toMatch(/outside the allowed roots/);
    expect(existsSync(join(away, "FN.HALF.lambda"))).toBe(false);
  });

  it("parses --root", () => {
    expect(parseArgs(["--root", "/a", "--root=/b"])).toEqual({ roots: ["/a", "/b"] });
    expect(parseArgs([])).toEqual({ roots: [process.cwd()] });
    expect(parseArgs(["--root"])).toEqual({ error: "--root needs a folder" });
    expect(parseArgs(["x"])).toEqual({ error: "unknown argument 'x'" });
  });
});
