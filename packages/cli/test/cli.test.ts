import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { replaceZipEntries, utf8 } from "@xln/core";
import { checkOptions, defaultOut, loadConfig, main } from "../src/main.js";

const PROBE = join(import.meta.dirname, "..", "..", "..", "probes", "results", "probe_win.xlsx");
const tmp = mkdtempSync(join(tmpdir(), "xln-cli-"));
afterAll(() => rmSync(tmp, { recursive: true, force: true }));

function capture() {
  const io = { stdout: "", stderr: "", out: (s: string) => void (io.stdout += s), err: (s: string) => void (io.stderr += s) };
  return io;
}

describe("xln pull", () => {
  it("writes the project next to the workbook by default and prints a summary", async () => {
    const wb = join(tmp, "model.xlsx");
    copyFileSync(PROBE, wb);
    const io = capture();
    expect(await main(["pull", wb], io)).toBe(0);
    const out = join(tmp, "model.xln");
    expect(defaultOut(wb)).toBe(out);
    for (const f of ["names/P.xln", "names/Mod.xln", "names/_unmanaged.xln", "workbook.manifest.json", "xln.lock.json"]) {
      expect(existsSync(join(out, f)), f).toBe(true);
    }
    expect(readFileSync(join(out, "names/P.xln"), "utf8")).toContain("P_Add1 = LAMBDA(x, x+1);");
    expect(io.stdout).toContain("10 names: 8 workbook-scoped, 2 sheet-scoped (S2 2)");
    expect(io.stdout).toContain("by kind: constant 2 · spill 1 · lambda 7");
    expect(io.stdout).toContain("modules: Mod (1) · P (6) · _unmanaged (2)");
    expect(io.stdout).toContain("cell statements: 1 named · 0 slots · 10 unnamed (0 blocks over 0 cells)");
    expect(io.stderr).toBe("");
    expect(JSON.parse(readFileSync(join(out, "xln.config.json"), "utf8"))).toEqual({ audit: { harness: [], rules: {}, constants: { allow: [] } } });
  });

  it("writes xln.config.json only when the project has none", async () => {
    const wb = join(tmp, "model.xlsx");
    const config = join(tmp, "model.xln", "xln.config.json");
    const mine = '{ "audit": { "harness": ["Check!*"] } }\n';
    writeFileSync(config, mine);
    const io = capture();
    expect(await main(["pull", wb, "--json"], io)).toBe(0);
    expect(JSON.parse(io.stdout).files).not.toContain("xln.config.json");
    expect(readFileSync(config, "utf8")).toBe(mine);
  });

  it("a second pull with nothing to replace writes nothing that did not change; warns on Excel's owner file", async () => {
    const wb = join(tmp, "model.xlsx");
    writeFileSync(join(tmp, "~$model.xlsx"), "");
    const io = capture();
    expect(await main(["pull", wb, "--json"], io)).toBe(0);
    const j = JSON.parse(io.stdout);
    expect(j.ok).toBe(true);
    expect(j.report.names).toBe(10);
    expect(j.files).toEqual([]);
    expect(j.discarded).toEqual([]);
    expect(j.notices.join("\n")).toMatch(/Excel has model.xlsx open/);
  });

  it("refuses (exit 1, nothing written) while the project has source edits not built, and lists them", async () => {
    const wb = join(tmp, "model.xlsx");
    const out = join(tmp, "model.xln");
    writeFileSync(join(out, "names", "Old.xln"), "X = 1;\n");
    const p = join(out, "names", "P.xln");
    writeFileSync(p, readFileSync(p, "utf8").replace("P_Add1 = LAMBDA(x, x+1);", "P_Add1 = LAMBDA(x, x+2);"));
    const before = readFileSync(p, "utf8");
    let io = capture();
    expect(await main(["pull", wb], io)).toBe(1);
    expect(io.stdout).toBe("");
    expect(io.stderr).toContain("xln pull model.xlsx: refused: the project has 2 source edits not built yet, which a pull would replace:");
    expect(io.stderr).toMatch(/names\/Old\.xln:1 {2}X: create X/);
    expect(io.stderr).toMatch(/names\/P\.xln:\d+ {2}P_Add1: update P_Add1 \(definition\)/);
    expect(io.stderr).toContain("Build them first (xln build), then pull; or pull with --discard");
    expect(readFileSync(p, "utf8")).toBe(before);
    expect(existsSync(join(out, "names", "Old.xln"))).toBe(true);
    io = capture();
    expect(await main(["pull", wb, "--json"], io)).toBe(1);
    const j = JSON.parse(io.stdout);
    expect(j).toMatchObject({ ok: false, refused: true });
    expect(j.unbuilt.map((e: { key: string }) => e.key).sort()).toEqual(["P_Add1", "X"]);
  });

  it("--discard replaces names/**, the lockfile and the manifest: files it does not write go, edits are lost, the config stays; --fresh is accepted with a note", async () => {
    const wb = join(tmp, "model.xlsx");
    const out = join(tmp, "model.xln");
    mkdirSync(join(out, "names", "sheets"), { recursive: true });
    writeFileSync(join(out, "names", "sheets", "Gone.xln"), "@scope(Gone)\nX = 1;\n");
    const p = join(out, "names", "P.xln");
    const config = readFileSync(join(out, "xln.config.json"), "utf8");
    const io = capture();
    expect(await main(["pull", wb, "--json", "--fresh", "--discard"], io)).toBe(0);
    const j = JSON.parse(io.stdout);
    expect(j.discarded.length).toBeGreaterThan(0);
    expect(j.notices.join("\n")).toMatch(/removed names\/Old.xln, names\/sheets\/Gone.xln/);
    expect(j.notices.join("\n")).toMatch(/--fresh is no longer needed: every pull is fresh/);
    expect(existsSync(join(out, "names", "Old.xln"))).toBe(false);
    expect(existsSync(join(out, "names", "sheets", "Gone.xln"))).toBe(false);
    expect(readFileSync(p, "utf8")).toContain("P_Add1 = LAMBDA(x, x+1);");
    expect(readFileSync(join(out, "xln.config.json"), "utf8")).toBe(config);
  });

  it("a hand-moved name goes back to its place: the structure is fixed", async () => {
    const wb = join(tmp, "model.xlsx");
    const out = join(tmp, "model.xln");
    const p = join(out, "names", "P.xln");
    const text = readFileSync(p, "utf8");
    writeFileSync(p, text.replace("P_Loc = LAMBDA(x, x+1);\n", ""));
    writeFileSync(join(out, "names", "Mine.xln"), "P_Loc = LAMBDA(x, x+1);\n");
    const io = capture();
    expect(await main(["pull", wb], io), io.stderr).toBe(0);
    expect(readFileSync(p, "utf8")).toBe(text);
    expect(existsSync(join(out, "names", "Mine.xln"))).toBe(false);
    // M3e: not an edit, but the pull says which files it rewrote.
    expect(io.stdout).toContain("layout or comments only (a pull rewrites them as the workbook has them): names/Mine.xln, names/P.xln");
  });

  it("M3e: a // comment alone is no refusal, but a note; with an edit too, the refusal notes it", async () => {
    const wb = join(tmp, "model.xlsx");
    const out = join(tmp, "model.xln");
    const p = join(out, "names", "P.xln");
    const text = readFileSync(p, "utf8");
    writeFileSync(p, "// my note\n" + text);
    let io = capture();
    expect(await main(["pull", wb, "--json"], io), io.stderr).toBe(0);
    expect(JSON.parse(io.stdout).rewritten).toEqual(["names/P.xln"]);
    expect(readFileSync(p, "utf8")).toBe(text);
    writeFileSync(p, "// my note\n" + text);
    writeFileSync(join(out, "names", "Old.xln"), "X = 1;\n");
    io = capture();
    expect(await main(["pull", wb], io)).toBe(1);
    expect(io.stderr).toContain("note: layout or comments only (a pull rewrites them as the workbook has them): names/P.xln");
    rmSync(join(out, "names", "Old.xln"));
    writeFileSync(p, text);
  });

  it("lists the per-sheet files in the summary", async () => {
    const out = join(tmp, "f7");
    const io = capture();
    expect(await main(["pull", join(PROBE, "..", "f7_base.xlsx"), "--out", out], io)).toBe(0);
    expect(io.stdout).toMatch(/sheet files: .*S2 \(\d+\)/);
    expect(readFileSync(join(out, "names/sheets/S2.xln"), "utf8")).toContain("// Other names on S2.\n\nLoc = 7;\n");
  });

  it("writes a workbook name on one sheet's cells in that sheet's file, and notes a name taken from a computed value (2026-10-07)", async () => {
    const f7 = new Uint8Array(readFileSync(join(PROBE, "..", "f7_base.xlsx")));
    const xml = strFromU8(unzipSync(f7)["xl/workbook.xml"]!);
    // S1!C2 = "Rate is "&Rate shows "Rate is 0.1": Create from Selection (Left column) would name D2:F2 after it.
    const named = xml.replace("<definedNames>", '<definedNames><definedName name="Rate_is_0.1">S1!$D$2:$F$2</definedName><definedName name="Blk">S1!$B$2:$B$4</definedName>');
    const wb = join(tmp, "labels.xlsx");
    writeFileSync(wb, replaceZipEntries(f7, new Map([["xl/workbook.xml", utf8(named)]])));
    const io = capture();
    expect(await main(["pull", wb], io)).toBe(0);
    // The dot makes it module Rate_is_0's (the module rule is unchanged); Blk has no module: S1's file.
    expect(readFileSync(join(tmp, "labels.xln", "names/Rate_is_0.xln"), "utf8")).toContain("\nRate_is_0.1 = 'S1'!$D$2:$F$2;\n");
    expect(readFileSync(join(tmp, "labels.xln", "names/sheets/S1.xln"), "utf8")).toContain("// Other names on S1.\n\n@workbook\nBlk = 'S1'!$B$2:$B$4;\n");
    expect(readFileSync(join(tmp, "labels.xln", "names/_unmanaged.xln"), "utf8")).not.toContain("Blk");
    expect(io.stdout).toContain("  note: Rate_is_0.1: named after the current value of 'S1'!C2, a formula result; the name stays as is when that value changes\n");
  });

  it("honours --out and --width", async () => {
    const out = join(tmp, "elsewhere");
    const io = capture();
    expect(await main(["pull", PROBE, "--out", out, "--width", "40"], io)).toBe(0);
    expect(readFileSync(join(out, "names/_unmanaged.xln"), "utf8")).toContain("Growλ = LAMBDA(b,g, b*(1+g));");
  });

  it("rejects bad usage and missing files", async () => {
    let io = capture();
    expect(await main(["frob"], io)).toBe(2);
    expect(io.stderr).toContain("unknown command 'frob'");
    io = capture();
    expect(await main(["pull"], io)).toBe(2);
    io = capture();
    expect(await main(["pull", join(tmp, "nope.xlsx")], io)).toBe(1);
    io = capture();
    const junk = join(tmp, "junk.xlsx");
    writeFileSync(junk, "not a zip");
    expect(await main(["pull", junk, "--json"], io)).toBe(1);
    expect(JSON.parse(io.stdout).ok).toBe(false);
    io = capture();
    expect(await main(["--help"], io)).toBe(0);
    expect(io.stderr).toContain("usage: xln pull");
    // Every command, current flags (spec §14 issue 12): the build no longer moves names to `#`.
    for (const c of ["xln formulas", "xln graph", "xln check", "xln build", "xln apply", "xln verify", "xln lib status", "xln lib publish", "xln lib take", "xln lib base"]) expect(io.stderr).toContain(`       ${c} `);
    for (const f of ["--severity error|warning|info|hint", "--no-tags", "--reopen", "--no-diff", "--all", "--tolerance", "--census-exclude"]) expect(io.stderr).toContain(f);
    expect(io.stderr).not.toContain("moves the name to the spill");
  });
});

describe("xln formulas", () => {
  const F7 = join(import.meta.dirname, "..", "..", "..", "probes", "results", "f7_base.xlsx");

  it("prints every sheet's formulas in order of appearance", async () => {
    const io = capture();
    expect(await main(["formulas", F7], io)).toBe(0);
    expect(io.stdout).toContain("// Sheet S1: 22 formulas in order of appearance");
    expect(io.stdout).toContain("// Sheet S2: 6 formulas");
    expect(io.stdout).toContain("     B3    shared ← B2  = A3*Rate               → 0.2");
    // Spl = 'S1'!$E$1#: the name leads the line of its anchor.
    expect(io.stdout).toContain("Spl  E1#   (3×1)        = SEQUENCE(3)*Rate      → 0.1 …");
  });

  it("gives the names on each cell in JSON (lhs)", async () => {
    const io = capture();
    expect(await main(["formulas", F7, "--sheet", "S1", "--json"], io)).toBe(0);
    const lines = JSON.parse(io.stdout).sheets[0].lines as { cell: string; lhs: unknown[] }[];
    expect(lines.find((l) => l.cell === "E1")!.lhs).toEqual([{ key: "Spl", name: "Spl", display: "Spl", hidden: false, target: "spill" }]);
    expect(lines.filter((l) => l.lhs.length > 0)).toHaveLength(1);
  });

  it("prints one sheet, or JSON lines with the names used", async () => {
    let io = capture();
    expect(await main(["formulas", F7, "--sheet", "s2"], io)).toBe(0);
    expect(io.stdout).not.toContain("Sheet S1");
    expect(io.stdout).toContain("A4  shared ← A3  = Loc*ROW()         → 28");
    io = capture();
    expect(await main(["formulas", F7, "--sheet", "S2", "--json"], io)).toBe(0);
    const j = JSON.parse(io.stdout);
    expect(j.ok).toBe(true);
    expect(j.sheets[0].lines[1]).toMatchObject({ cell: "B1", formula: "'S1'!A1+'S2'!Loc", names: [{ id: "Loc", sheet: "S2", key: "S2!Loc" }] });
  });

  it("prints a sheet in calculation order, with levels", async () => {
    const io = capture();
    expect(await main(["formulas", F7, "--sheet", "S1", "--order", "calculation"], io)).toBe(0);
    expect(io.stdout).toContain("// Sheet S1: 22 formulas in calculation order");
    // C6 = SUM(Spl) reads E1#: level 2, after E1.
    expect(io.stdout).toContain("2       C6                 = SUM(Spl)              → 0.6");
    expect(io.stdout.indexOf("E1#")).toBeLessThan(io.stdout.indexOf(" C6 "));
  });

  it("prints the whole workbook in calculation order, JSON with levels and what each line reads", async () => {
    let io = capture();
    expect(await main(["formulas", F7, "--workbook"], io)).toBe(0);
    expect(io.stdout).toContain("// Workbook f7_base.xlsx: 28 formulas on 2 sheets in calculation order");
    expect(io.stdout).toMatch(/'S2'!A1 /);
    io = capture();
    expect(await main(["formulas", F7, "--workbook", "--json"], io)).toBe(0);
    const j = JSON.parse(io.stdout);
    expect(j.order).toBe("calculation");
    const c6 = (j.sheets[0].lines as { sheet: string; cell: string; level: number; dependsOn: string[] }[]).find((l) => l.sheet === "S1" && l.cell === "C6")!;
    expect(c6).toMatchObject({ level: 2, dependsOn: ["Spl"] });
    io = capture();
    expect(await main(["formulas", F7, "--workbook", "--sheet", "S1"], io)).toBe(2);
    io = capture();
    expect(await main(["formulas", F7, "--order", "sideways"], io)).toBe(2);
  });

  it("names the sheets when --sheet matches none", async () => {
    const io = capture();
    expect(await main(["formulas", F7, "--sheet", "Nope"], io)).toBe(1);
    expect(io.stderr).toContain("no sheet 'Nope'; the workbook has S1, S2");
  });
});

describe("xln graph", () => {
  const F7 = join(import.meta.dirname, "..", "..", "..", "probes", "results", "f7_base.xlsx");

  it("summarises nodes, cycles, flagged references, C9, C10, C12", async () => {
    let io = capture();
    expect(await main(["graph", F7], io)).toBe(0);
    expect(io.stdout).toContain("xln graph f7_base.xlsx");
    expect(io.stdout).toMatch(/nodes: 28 formulas · 10 input ranges · 7 names; 45 edges/);
    expect(io.stdout).toContain("circular references: none");
    expect(io.stdout).toContain("not followed: 1 dynamic, 0 external, 0 broken");
    expect(io.stdout).toContain("C10 unused names: none");
    io = capture();
    expect(await main(["graph", F7, "--json"], io)).toBe(0);
    const j = JSON.parse(io.stdout);
    expect(j).toMatchObject({ ok: true, nodes: { formulas: 28, names: 7 }, cycles: [], spillRefs: [], unusedNames: [] });
    expect(j.flagged).toEqual([{ node: "'S1'!C7#", kind: "dynamic", reason: "INDIRECT: the reference is computed", text: 'INDIRECT("Rate")' }]);
  });

  it("lists unused names", async () => {
    const io = capture();
    expect(await main(["graph", PROBE, "--json"], io)).toBe(0);
    expect(JSON.parse(io.stdout).unusedNames).toEqual(["Fact", "P_Loc", "'S2'!P_Local2"]);
  });
});

describe("xln check", () => {
  const TRAPS = join(import.meta.dirname, "..", "..", "..", "probes", "fixtures", "traps.xlsx");
  const F7 = join(import.meta.dirname, "..", "..", "..", "probes", "results", "f7_base.xlsx");

  it("prints the report and exits 1 when there are errors", async () => {
    const io = capture();
    expect(await main(["check", TRAPS], io)).toBe(1);
    expect(io.stdout).toContain("xln check traps.xlsx: 9 errors, 6 warnings, 1 info");
    expect(io.stdout).toMatch(/C5 Unqualified sheet-scoped reads\s+1 error/);
    expect(io.stdout).toContain("== C6 Arity: 2 errors");
    expect(io.stdout).toContain("  error   cell 'S1'!A18: Fn(x) takes 1 argument; it is given 2");
    expect(io.stdout).toContain("== C8 Name census");
    expect(io.stdout).toContain("== C9 spill census");
    expect(io.stderr).toBe("");
  });

  it("prints JSON for agents, filters by check and severity", async () => {
    let io = capture();
    expect(await main(["check", TRAPS, "--json"], io)).toBe(1);
    const j = JSON.parse(io.stdout);
    expect(j).toMatchObject({ ok: true, format: "xln-audit/1", workbook: "traps.xlsx", counts: { error: 9, warning: 6, info: 1 } });
    expect(j.findings[0]).toMatchObject({ check: "C2", rule: "C2.bare-prefix", severity: "error", where: { kind: "name", name: "BareSeq" } });
    expect(j.census.tiers).toBeDefined();
    io = capture();
    expect(await main(["check", TRAPS, "--only", "c9,C10", "--severity", "warning", "--json"], io)).toBe(0);
    expect(JSON.parse(io.stdout).findings.map((f: { rule: string }) => f.rule)).toEqual(["C9.fixed-ref", "C9.fixed-ref", "C10.unused", "C10.unused"]);
    io = capture();
    expect(await main(["check", TRAPS, "--census-exclude", "Cost_*,Sales_*", "--json"], io)).toBe(1);
    expect(JSON.parse(io.stdout).census.tiers.excluded).toBe(4);
  });

  it("exits 0 on a workbook without errors", async () => {
    const io = capture();
    expect(await main(["check", F7], io)).toBe(0);
    expect(io.stdout).toContain("xln check f7_base.xlsx: ");
  });

  it("reports a name whose label cell no longer gives it (C15, 2026-10-07)", async () => {
    const f7 = new Uint8Array(readFileSync(F7));
    const files = unzipSync(f7);
    const cell = (r: string, text: string) => `<c r="${r}" t="inlineStr"><is><t>${text}</t></is></c>`;
    // Create from Selection over D7:F9 of S2, then Gross_income renamed Gross_ind_income.
    const rows = [
      ["7", "Revenue", 100, 110],
      ["8", "COGS", 60, 66],
      ["9", "Gross income", 40, 44],
    ].map(([r, label, e, f]) => `<row r="${r}">${cell(`D${r}`, label as string)}<c r="E${r}"><v>${e}</v></c><c r="F${r}"><v>${f}</v></c></row>`);
    const sheet = strFromU8(files["xl/worksheets/sheet2.xml"]!).replace("</sheetData>", `${rows.join("")}</sheetData>`);
    const book = strFromU8(files["xl/workbook.xml"]!).replace(
      "<definedNames>",
      '<definedNames><definedName name="COGS">S2!$E$8:$F$8</definedName><definedName name="Gross_ind_income">S2!$E$9:$F$9</definedName><definedName name="Revenue">S2!$E$7:$F$7</definedName>',
    );
    const wb = join(tmp, "label-drift.xlsx");
    writeFileSync(wb, replaceZipEntries(f7, new Map([["xl/workbook.xml", utf8(book)], ["xl/worksheets/sheet2.xml", utf8(sheet)]])));
    let io = capture();
    expect(await main(["check", wb, "--only", "C15"], io)).toBe(0);
    expect(io.stdout).toContain("info    name Gross_ind_income: 'S2'!D9 reads \"Gross income\", which Create from Selection makes Gross_income; the name on 'S2'!E9:F9 is Gross_ind_income, probably renamed from Gross_income\n");
    io = capture();
    expect(await main(["check", wb, "--only", "C15", "--json"], io)).toBe(0);
    const r = JSON.parse(io.stdout);
    expect(r.findings.map((f: { rule: string; where: { key: string } }) => `${f.rule} ${f.where.key}`)).toEqual(["C15.label-drift Gross_ind_income"]);
  });

  it("exits 2 on bad usage or a workbook it cannot read", async () => {
    let io = capture();
    expect(await main(["check", TRAPS, "--only", "C16"], io)).toBe(2);
    expect(io.stderr).toContain("--only takes checks C1 to C15");
    io = capture();
    expect(await main(["check", TRAPS, "--severity", "loud"], io)).toBe(2);
    io = capture();
    expect(await main(["check", join(tmp, "nope.xlsx")], io)).toBe(2);
    io = capture();
    writeFileSync(join(tmp, "junk.xlsx"), "not a zip");
    expect(await main(["check", join(tmp, "junk.xlsx")], io)).toBe(2);
  });
});

describe("xln check runs the editor's source checks on the project too (feedback 2026-10-07)", () => {
  const F7 = join(import.meta.dirname, "..", "..", "..", "probes", "results", "f7_base.xlsx");
  const dir = join(tmp, "source-check");
  const wb = join(dir, "f7.xlsx");
  const S1 = join(dir, "f7.xln", "names", "sheets", "S1.xln");
  const S2 = join(dir, "f7.xln", "names", "sheets", "S2.xln");
  mkdirSync(dir, { recursive: true });
  copyFileSync(F7, wb);

  it("on the workbook or on its project folder: the same findings, a hint with its quick fix", async () => {
    expect(await main(["pull", wb], capture())).toBe(0);
    const io = capture();
    expect(await main(["check", wb], io)).toBe(0);
    expect(io.stdout).toContain(`source ${join(dir, "f7.xln")}: 0 errors, 0 warnings, 0 info, 1 hint (the editor's checks as you type)\n`);
    expect(io.stdout).toMatch(/\n {2}hint {4}names\/sheets\/S1\.xln:\d+:1 workbook-on-cell: workbook name on a cell of S1, read only on S1: remove @workbook to make it local to S1 \(the build then moves it\) \[quick fix: Remove @workbook: make Spl local to S1\]\n/);
    const onProject = capture();
    expect(await main(["check", join(dir, "f7.xln"), "--json"], onProject)).toBe(0);
    const j = JSON.parse(onProject.stdout);
    expect(j.workbook).toBe("f7.xlsx");
    expect(j.source.findings.map((f: { severity: string; code: string }) => `${f.severity} ${f.code}`)).toEqual(["hint workbook-on-cell"]);
    // --severity leaves hints out, as it leaves out the audit's lesser findings.
    const severe = capture();
    await main(["check", wb, "--severity", "warning", "--json"], severe);
    expect(JSON.parse(severe.stdout).source.findings).toEqual([]);
    // --severity hint is accepted (spec §14 issue 13) and lists everything.
    const all = capture();
    expect(await main(["check", wb, "--severity", "hint", "--json"], all)).toBe(0);
    expect(JSON.parse(all.stdout).source.findings.map((f: { code: string }) => f.code)).toEqual(["workbook-on-cell"]);
  });

  it("another sheet reads the name unqualified: no finding; a source error makes the exit 1", async () => {
    writeFileSync(S2, readFileSync(S2, "utf8").replace("@A1 = Loc;", "@A1 = Loc + SUM(Spl);"));
    let io = capture();
    expect(await main(["check", wb, "--json"], io)).toBe(0);
    expect(JSON.parse(io.stdout).source.findings).toEqual([]);
    writeFileSync(S1, readFileSync(S1, "utf8").replace("@C8 = Rate*3;", "@C8 = Rate*;"));
    io = capture();
    expect(await main(["check", wb], io)).toBe(1);
    expect(io.stdout).toMatch(/\n {2}error {3}names\/sheets\/S1\.xln:\d+:\d+ syntax: /);
  });
});

describe("a misplaced key in xln.config.json is named, with where it goes (feedback 2026-10-07)", () => {
  const F7 = join(import.meta.dirname, "..", "..", "..", "probes", "results", "f7_base.xlsx");
  const dir = join(tmp, "config-keys");
  const wb = join(dir, "f7.xlsx");
  const config = join(dir, "f7.xln", "xln.config.json");
  const lib = join(dir, "lib");
  mkdirSync(lib, { recursive: true });
  copyFileSync(F7, wb);
  const MOVE = 'audit.library: `library` is a top-level setting, not an audit setting: move it out of "audit" (ignored here)';

  it("by build, lib status and check", async () => {
    expect(await main(["pull", wb], capture())).toBe(0);
    writeFileSync(config, JSON.stringify({ audit: { harness: [], library: "../lib" } }, null, 2));
    let io = capture();
    expect(await main(["build", wb, "--dry-run"], io)).toBe(0);
    expect(io.stdout).toContain(`  note: xln.config.json: ${MOVE}\n`);
    io = capture();
    expect(await main(["lib", "status", wb], io)).toBe(2);
    expect(io.stderr).toContain(MOVE);
    io = capture();
    await main(["check", wb, "--only", "C13"], io);
    expect(io.stderr).toContain(MOVE);
    // Moved where it belongs: no note, and the library is read.
    writeFileSync(config, JSON.stringify({ audit: { harness: [] }, library: "../lib" }, null, 2));
    io = capture();
    expect(await main(["lib", "status", wb], io)).toBe(0);
    expect(io.stdout).not.toContain("note:");
    expect(io.stdout).toContain(`against ${lib}`);
  });
});

describe("xln check with project settings", () => {
  const TRAPS = join(import.meta.dirname, "..", "..", "..", "probes", "fixtures", "traps.xlsx");
  const dir = join(tmp, "settings");
  const wb = join(dir, "traps.xlsx");
  const config = join(dir, "traps.xln", "xln.config.json");
  mkdirSync(join(dir, "traps.xln"), { recursive: true });
  copyFileSync(TRAPS, wb);
  const rules = (out: string): string[] => JSON.parse(out).findings.map((f: { rule: string; where: { key?: string } }) => `${f.rule} ${f.where.key ?? ""}`.trim());

  it("reads <workbook>.xln/xln.config.json: harness, rules, constants", async () => {
    writeFileSync(config, JSON.stringify({ audit: { harness: ["Unused"], rules: { "C10.unused": "info" }, constants: { allow: [0.27] } } }));
    const io = capture();
    expect(await main(["check", wb, "--only", "C10,C13", "--json"], io)).toBe(0);
    const j = JSON.parse(io.stdout);
    expect(j.config).toEqual({ path: config, problems: [] });
    expect(rules(io.stdout)).toEqual(["C10.unused Fact"]);
    expect(j.findings[0].severity).toBe("info");
    expect(j.census.tierNames.excluded).toEqual(["Unused"]);
  });

  it("lets the flags win, and --config point elsewhere", async () => {
    let io = capture();
    expect(await main(["check", wb, "--only", "C10", "--census-exclude", "Fact", "--json"], io)).toBe(0);
    expect(rules(io.stdout)).toEqual(["C10.unused Unused"]);
    const other = join(dir, "other.json");
    writeFileSync(other, JSON.stringify({ audit: { rules: { C10: "off" } } }));
    io = capture();
    expect(await main(["check", wb, "--only", "C10", "--config", other, "--json"], io)).toBe(0);
    expect(rules(io.stdout)).toEqual([]);
    expect(JSON.parse(io.stdout).config.path).toBe(other);
    io = capture();
    expect(await main(["check", wb, "--config", join(dir, "nope.json")], io)).toBe(2);
    expect(io.stderr).toContain("no such config file");
  });

  it("says what in the file it could not use, and names the file", async () => {
    writeFileSync(config, JSON.stringify({ audit: { harness: "Check!*", rules: { C99: "off", C13: "loud" }, extra: 1 } }));
    const io = capture();
    await main(["check", wb, "--only", "C13"], io);
    expect(io.stdout.startsWith(`settings: ${config}\n`)).toBe(true);
    expect(io.stderr).toContain("audit.harness: must be a list of glob patterns");
    expect(io.stderr).toContain("audit.rules.C99: no such rule or check");
    expect(io.stderr).toContain("audit.rules.C13: must be off, info, warning or error");
    expect(io.stderr).toContain("audit.extra: unknown setting");
    writeFileSync(config, "{ not json");
    expect(loadConfig(wb)!.problems[0]).toMatch(/^not JSON/);
  });

  it("an old names.scope is ignored with one note on stdout, no problem", async () => {
    writeFileSync(config, JSON.stringify({ names: { scope: "excel" } }));
    const io = capture();
    await main(["check", wb, "--only", "C13"], io);
    expect(io.stderr).toBe("");
    expect(io.stdout.split("\n").filter((l) => l.startsWith("note: "))).toEqual(["note: names.scope: no longer a setting (ignored): a scope set in Excel or in the source is carried to the other side by the next pull or build"]);
    expect(loadConfig(wb)!.problems).toEqual([]);
  });

  it("merges settings and flags into audit options", () => {
    const settings = { harness: ["Check!*"], rules: { C13: "off" as const }, constants: { allow: [4], sentinelAbove: 1e50 } };
    const o = checkOptions({ workbook: wb, json: false, severity: "warning" }, settings);
    expect(o).toMatchObject({ workbook: "traps.xlsx", minSeverity: "warning", harness: ["Check!*"], rules: { C13: "off" }, constants: { sentinelAbove: 1e50 } });
    expect(o.constants!.allow).toContain(4);
    expect(o.constants!.allow).toContain(365);
    expect(checkOptions({ workbook: wb, json: false, censusExclude: ["X*"] }, settings).harness).toEqual(["X*"]);
    expect(checkOptions({ workbook: wb, json: false }, undefined)).toEqual({ workbook: "traps.xlsx" });
    rmSync(config);
    expect(loadConfig(wb)).toBeUndefined();
  });
});

describe("a workbook with AFE's module store (synthetic fixture)", () => {
  const AFE = join(import.meta.dirname, "..", "..", "..", "probes", "fixtures", "afe-synthetic-v11.xlsx");

  it("pull notes it, check reports it (C14), build warns and leaves it as it is", async () => {
    const wb = join(tmp, "afe.xlsx");
    copyFileSync(AFE, wb);
    let io = capture();
    expect(await main(["pull", wb], io)).toBe(0);
    expect(io.stdout).toContain("note: the workbook carries Microsoft's Advanced Formula Environment (AFE) modules: customXml/item1.xml");
    io = capture();
    expect(await main(["check", wb, "--only", "C14"], io)).toBe(0);
    expect(io.stdout).toContain("C14 AFE's copy of the names");
    expect(io.stdout).toContain("part customXml/item1.xml: AFE's module store: AFE keeps its own copy of 4 names");
    const ana = join(tmp, "afe.xln", "names", "ANA.xln");
    writeFileSync(ana, readFileSync(ana, "utf8").replace("LAMBDA(x, x ^ 3)", "LAMBDA(x, x * x * x)"));
    io = capture();
    expect(await main(["build", wb, "--no-tags"], io)).toBe(0);
    expect(io.stdout).toContain("warning: the workbook carries modules of Microsoft's Advanced Formula Environment (Excel Labs) in customXml/item1.xml (AFE's module store); this build changes 1 name that AFE's modules also define (ANA.CUBE)");
    io = capture();
    expect(await main(["check", wb, "--only", "C14", "--json"], io)).toBe(0);
    const r = JSON.parse(io.stdout);
    expect(r.findings.map((f: { rule: string }) => f.rule)).toEqual(["C14.afe-store", "C14.afe-drift"]);
  });
});
