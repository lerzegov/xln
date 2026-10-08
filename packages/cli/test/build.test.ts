import { copyFileSync, existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { readWorkbook } from "@xln/core";
import { main } from "../src/main.js";

const F7 = join(import.meta.dirname, "..", "..", "..", "probes", "results", "f7_base.xlsx");
const tmp = mkdtempSync(join(tmpdir(), "xln-build-"));
afterAll(() => rmSync(tmp, { recursive: true, force: true }));

function capture() {
  const io = { stdout: "", stderr: "", out: (s: string) => void (io.stdout += s), err: (s: string) => void (io.stderr += s) };
  return io;
}

async function setup(name: string): Promise<{ wb: string; project: string; names: string }> {
  const wb = join(tmp, `${name}.xlsx`);
  copyFileSync(F7, wb);
  expect(await main(["pull", wb], capture())).toBe(0);
  const project = join(tmp, `${name}.xln`);
  return { wb, project, names: join(project, "names", "_unmanaged.xln") };
}

function setRate(file: string, value: string): void {
  writeFileSync(file, readFileSync(file, "utf8").replace(/Rate = [0-9.]+;/, `Rate = ${value};`));
}

describe("xln build: cell statements", () => {
  it("--dry-run --json shows the cell change; a build writes it; a scope changes by removing @workbook (--rescope-slots is gone)", async () => {
    const { wb, project } = await setup("cells");
    const s1 = join(project, "names", "sheets", "S1.xln");
    writeFileSync(s1, readFileSync(s1, "utf8").replace("@C8 = Rate*3;", "@C8 = Rate*4;"));
    const before = readFileSync(wb);
    let io = capture();
    expect(await main(["build", wb, "--dry-run", "--json"], io)).toBe(0);
    const j = JSON.parse(io.stdout);
    expect(j.changeSet.changes.filter((c: { op: string }) => c.op !== "set-embedded-source")).toEqual([{ op: "set-cell-formula", sheet: "S1", range: "C8", stored: "Rate*4", display: "Rate*4", previous: "Rate*3" }]);
    expect(j.rescopeCandidates).toBeUndefined();
    expect(readFileSync(wb).equals(before)).toBe(true);
    io = capture();
    expect(await main(["build", wb], io)).toBe(0);
    expect(io.stdout).toContain("set formula of S1!C8");
    expect(io.stdout).not.toContain("rescope");
    expect(readWorkbook(new Uint8Array(readFileSync(wb))).sheets[0]!.formulas.find((f) => f.cell === "C8")!.text).toBe("Rate*4");
    io = capture();
    expect(await main(["build", wb, "--rescope-slots"], io)).toBe(2);
    expect(io.stderr).toContain("--rescope-slots is gone (M3d)");
    // The way now: remove @workbook above the name; the build moves it and leaves the file as written.
    const text = readFileSync(s1, "utf8").replace("@workbook\nSpl @E1#", "Spl @E1#");
    writeFileSync(s1, text);
    io = capture();
    expect(await main(["build", wb], io)).toBe(0);
    expect(io.stdout).toContain("move Spl to sheet S1");
    expect(readFileSync(s1, "utf8")).toBe(text);
  });
});

describe("xln build", () => {
  it("nothing to do: up to date, nothing written", async () => {
    const { wb } = await setup("same");
    const before = readFileSync(wb);
    const io = capture();
    expect(await main(["build", wb, "--no-embed"], io)).toBe(0);
    expect(io.stdout).toContain("up to date");
    expect(readFileSync(wb).equals(before)).toBe(true);
    expect(existsSync(join(tmp, "same.backup.xlsx"))).toBe(false);
  });

  it("D5 is opt-in: a plain build embeds nothing; --embed or the config does; pull does not read the part", async () => {
    const { wb, project, names } = await setup("embed");
    // The author's layout and comments, which the live names cannot give back.
    writeFileSync(names, readFileSync(names, "utf8").replace("Rate = ", "// the growth rate, kept by the embedded source\nRate = "));
    let io = capture();
    expect(await main(["build", wb], io)).toBe(0);
    expect(io.stdout).toContain("up to date");
    io = capture();
    expect(await main(["build", wb, "--embed"], io)).toBe(0);
    expect(io.stdout).toContain("embed the source");
    io = capture();
    expect(await main(["build", wb, "--embed"], io)).toBe(0);
    expect(io.stdout).toContain("up to date");
    const source = readFileSync(names, "utf8");
    expect(source).toContain("the growth rate");
    // A colleague pulls the workbook into a new folder: the project comes from the workbook;
    // the part is an archive copy, which pull does not read.
    const elsewhere = join(tmp, "embed-colleague");
    io = capture();
    expect(await main(["pull", wb, "--out", elsewhere], io)).toBe(0);
    expect(readFileSync(join(elsewhere, "names", "_unmanaged.xln"), "utf8")).not.toContain("the growth rate");
    // --fresh and --live are accepted, with a note.
    io = capture();
    expect(await main(["pull", wb, "--out", join(tmp, "embed-live"), "--live"], io)).toBe(0);
    expect(io.stdout).toContain("--live is no longer needed: every pull is fresh");
    // The project config turns embedding on; --no-embed wins over it.
    const config = join(project, "xln.config.json");
    writeFileSync(config, JSON.stringify({ ...JSON.parse(readFileSync(config, "utf8")), build: { embed: true } }));
    setRate(names, "0.25");
    io = capture();
    expect(await main(["build", wb, "--no-embed", "--dry-run", "--json"], io)).toBe(0);
    expect(JSON.parse(io.stdout).changeSet.changes.map((c: { op: string }) => c.op)).toEqual(["set-name"]);
    io = capture();
    expect(await main(["build", wb, "--json"], io)).toBe(0);
    expect(JSON.parse(io.stdout).changeSet.changes.map((c: { op: string }) => c.op)).toEqual(["set-name", "set-embedded-source"]);
  });

  it("a pull brings in a name created in Excel, where a pull puts it; a // comment is no edit and goes", async () => {
    const { wb, project } = await setup("merge");
    const s1 = join(project, "names", "sheets", "S1.xln");
    writeFileSync(s1, readFileSync(s1, "utf8").replace("@C5 = ", "// a comment of the author's\n@C5 = "));
    // "Excel": Create from Selection on S1!D5 (a workbook-scoped name on an empty cell).
    const changes = join(tmp, "merge-changes.json");
    writeFileSync(changes, JSON.stringify([{ op: "set-name", name: "Errored_balance_base", scope: null, stored: "S1!$D$5", display: "S1!$D$5", comment: null, hidden: false, fields: ["created"] }]));
    expect(await main(["apply", wb, changes], capture())).toBe(0);
    const io = capture();
    expect(await main(["pull", wb], io), io.stderr).toBe(0);
    const text = readFileSync(s1, "utf8");
    expect(text).toContain("\n@C5 = 'S2'!Loc+Loc;\n@workbook\nErrored_balance_base @D5 = ;\n@C6 = ");
    expect(text).not.toContain("a comment of the author's");
    const b = capture();
    expect(await main(["build", wb], b)).toBe(0);
    expect(b.stdout).toContain("up to date");
  });

  it("the author's case: a scope changed in Excel (delete, create again in the other scope) comes through", async () => {
    const { wb, project } = await setup("rescope");
    const s1 = join(project, "names", "sheets", "S1.xln");
    expect(readFileSync(s1, "utf8")).toContain("@workbook\nSpl @E1# =");
    const changes = join(tmp, "rescope-changes.json");
    writeFileSync(changes, JSON.stringify({ format: "xln.changes/1", workbook: "rescope.xlsx", changes: [{ op: "rescope-name", name: "Spl", from: null, to: "S1" }] }));
    expect(await main(["apply", wb, changes], capture())).toBe(0);
    expect(await main(["pull", wb], capture())).toBe(0);
    expect(readFileSync(s1, "utf8")).toContain("@C1 = RateX+Rate;\nSpl @E1# = SEQUENCE(3)*Rate;\n");
    const b = capture();
    expect(await main(["build", wb], b)).toBe(0);
    expect(b.stdout).toContain("up to date");
  });

  it("builds: writes the workbook, keeps a backup, updates the lockfile; verify finds no changed value", async () => {
    const { wb, project, names } = await setup("edit");
    const original = readFileSync(wb);
    setRate(names, "0.2");
    const io = capture();
    expect(await main(["build", wb, "--json"], io)).toBe(0);
    const j = JSON.parse(io.stdout);
    expect(j).toMatchObject({ ok: true, status: "built", backup: join(tmp, "edit.backup.xlsx"), projectFiles: ["xln.lock.json", "workbook.manifest.json"] });
    expect(j.changeSet.changes).toMatchObject([{ op: "set-name", name: "Rate", stored: "0.2" }]);
    expect(j.readBack.ok).toBe(true);
    expect(readFileSync(join(tmp, "edit.backup.xlsx")).equals(original)).toBe(true);
    expect(readWorkbook(new Uint8Array(readFileSync(wb))).definedNames.find((d) => d.name === "Rate")!.definition).toBe("0.2");
    // Built again from the same source: nothing to do.
    expect(await main(["build", wb], capture())).toBe(0);
    // Without Excel the cached values are those of the backup.
    const v = capture();
    expect(await main(["verify", wb, "--json"], v)).toBe(0);
    expect(JSON.parse(v.stdout)).toMatchObject({ ok: true, changed: [] });
    expect(existsSync(join(project, "xln.lock.json"))).toBe(true);
  });

  it("--dry-run prints the change set and writes nothing", async () => {
    const { wb, names } = await setup("dry");
    const before = readFileSync(wb);
    setRate(names, "0.3");
    const io = capture();
    expect(await main(["build", wb, "--dry-run"], io)).toBe(0);
    expect(io.stdout).toContain("update Rate (definition)");
    expect(readFileSync(wb).equals(before)).toBe(true);
  });

  it("refuses while Excel has the workbook open (exit 3)", async () => {
    const { wb, names } = await setup("open");
    setRate(names, "0.4");
    writeFileSync(join(tmp, "~$open.xlsx"), "");
    const before = readFileSync(wb);
    const io = capture();
    expect(await main(["build", wb], io)).toBe(3);
    expect(io.stderr).toMatch(/Excel has open.xlsx open/);
    expect(readFileSync(wb).equals(before)).toBe(true);
  });

  it("--out writes a copy and leaves the original and the lockfile alone", async () => {
    const { wb, project, names } = await setup("copy");
    const before = readFileSync(wb);
    const lock = readFileSync(join(project, "xln.lock.json"), "utf8");
    setRate(names, "0.5");
    const out = join(tmp, "copy.xln.xlsx");
    expect(await main(["build", wb, "--out", out], capture())).toBe(0);
    expect(readFileSync(wb).equals(before)).toBe(true);
    expect(readFileSync(join(project, "xln.lock.json"), "utf8")).toBe(lock);
    expect(readWorkbook(new Uint8Array(readFileSync(out))).definedNames.find((d) => d.name === "Rate")!.definition).toBe("0.5");
  });

  it("refuses on a conflict (exit 1) and on a name cells still use", async () => {
    const { wb, names } = await setup("conflict");
    setRate(names, "0.6");
    // Rate changed "in Excel" meanwhile: build a copy with a different value and put it in place.
    const other = await setup("conflict-other");
    setRate(other.names, "0.7");
    expect(await main(["build", other.wb], capture())).toBe(0);
    copyFileSync(other.wb, wb);
    const io = capture();
    expect(await main(["build", wb, "--json"], io)).toBe(1);
    expect(JSON.parse(io.stdout).conflicts).toMatchObject([{ kind: "both-changed", key: "Rate", excel: { display: "0.7" }, source: { display: "0.6" } }]);

    const used = await setup("used");
    writeFileSync(used.names, readFileSync(used.names, "utf8").replace("RateX = 0.5;\n", ""));
    const io2 = capture();
    expect(await main(["build", used.wb], io2)).toBe(1);
    expect(io2.stderr).toMatch(/deleting RateX refused: 1 place in the workbook refers to it.*cell S1!C1 \(names\/sheets\/S1\.xln:\d+\)\. Change that formula in the source, or keep the name/);
    // The reason first, then the plan it did not write.
    const lines = io2.stderr.split("\n");
    expect(lines[1]).toMatch(/^ {2}error names\/sheets\/S1\.xln:\d+: deleting RateX refused/);
    expect(lines.slice(2, 4)).toEqual(["  the plan, not written:", "    delete RateX"]);
  });

  it("refuses a file that has no place in names/ (M3e); a hidden one does not count", async () => {
    const { wb, project } = await setup("stray");
    writeFileSync(join(project, "names", ".DS_Store"), "");
    expect(await main(["build", wb], capture())).toBe(0);
    writeFileSync(join(project, "names", "notes.txt"), "remember");
    const io = capture();
    expect(await main(["build", wb], io)).toBe(1);
    expect(io.stderr).toContain("names/notes.txt: only .xln files belong in names/");
  });

  it("verify reports a changed cell (exit 1)", async () => {
    const a = join(tmp, "va.xlsx");
    copyFileSync(F7, a);
    const b = join(tmp, "vb.xlsx");
    copyFileSync(join(import.meta.dirname, "..", "..", "..", "probes", "results", "probe_mac.xlsx"), b);
    const io = capture();
    expect(await main(["verify", b, "--before", a], io)).toBe(1);
    expect(io.stdout).toMatch(/changed/);
    expect(await main(["verify", join(tmp, "nobackup.xlsx")], capture())).toBe(2);
  });

  it("verify warns first when a side has no values Excel calculated", async () => {
    // f8_q1: written by a tool (fullCalcOnLoad), its formula cells carry no value.
    const a = join(tmp, "original.xlsx");
    copyFileSync(join(import.meta.dirname, "..", "..", "..", "probes", "results", "f8_q1.xlsx"), a);
    const b = join(tmp, "resaved.xlsx");
    copyFileSync(join(import.meta.dirname, "..", "..", "..", "probes", "results", "f8_q1_resaved.xlsx"), b);
    const want = "original.xlsx has no values Excel calculated (it was never saved by Excel): open and save it in Excel first, or the comparison is empty";
    const io = capture();
    await main(["verify", b, "--before", a], io);
    expect(io.stdout.startsWith(`warning: ${want}\nxln verify resaved.xlsx against original.xlsx: `)).toBe(true);
    const j = capture();
    await main(["verify", b, "--before", a, "--json"], j);
    expect(JSON.parse(j.stdout)).toMatchObject({ warnings: [want], beforeValues: { origin: "none", formulaCells: 2, withoutValue: 2, fullCalcOnLoad: true }, afterValues: { origin: "excel" } });
    // Both Excel's: no warning.
    const ok = capture();
    await main(["verify", b, "--before", F7, "--json"], ok);
    expect(JSON.parse(ok.stdout).warnings).toEqual([]);
  });
});
