// `xln lib status` and `xln lib publish` on a temporary library and project (the author's
// library is only ever copied: XLN_CORPUS's _shared/lib is read, never written).
import { copyFileSync, cpSync, existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { main } from "../src/main.js";
import { expandHome, fileDiff } from "../src/lib.js";

const PROBE = join(import.meta.dirname, "..", "..", "..", "probes", "results", "probe_win.xlsx");
const tmp = mkdtempSync(join(tmpdir(), "xln-lib-"));
afterAll(() => rmSync(tmp, { recursive: true, force: true }));

function capture() {
  const io = { stdout: "", stderr: "", out: (s: string) => void (io.stdout += s), err: (s: string) => void (io.stderr += s) };
  return io;
}

const LIB = join(tmp, "lib");
const lambda = (name: string, summary: string, def: string) => `# name       ${name}\n# summary    ${summary}\n# params     x\n# example    ${name}(1)\n#\n# Why ${name}.\n\n${def}\n`;

describe("xln lib", () => {
  const wb = join(tmp, "model.xlsx");
  const project = join(tmp, "model.xln");

  it("status of a workbook and of its project, the library from --lib or xln.config.json", async () => {
    mkdirSync(LIB);
    writeFileSync(join(LIB, "FN.TWICE.lambda"), lambda("FN.TWICE", "Doubles.", "LAMBDA(x, x * 2)"));
    writeFileSync(join(LIB, "FN.INC.lambda"), lambda("FN.INC", "Adds two.", "LAMBDA(x, x + 2)"));
    writeFileSync(join(LIB, "README.md"), "not a function");
    copyFileSync(PROBE, wb);
    expect(await main(["pull", wb], capture())).toBe(0);
    writeFileSync(join(project, "names", "FN.xln"), "/** Doubles. */\nFN.TWICE = LAMBDA(x, x*2);\n\nFN.INC = LAMBDA(x, x + 1);\n\nFN.OWN = LAMBDA(x, x - 1);\n");
    const b = capture();
    expect(await main(["build", wb], b), b.stderr).toBe(0);

    const io = capture();
    expect(await main(["lib", "status", wb, "--lib", LIB], io)).toBe(0);
    // FN.INC was written by hand, not taken from the library: no base, so it differs.
    expect(io.stdout).toContain("outdated 0 · modified 0 · both changed 0 · differs 1 · missing 0 · identical 1 · local only 1 (2 library functions)");
    expect(io.stdout).toContain("other module LAMBDAs, not library candidates: 5 (Mod 1, P 4)");
    expect(io.stdout).toContain("      - LAMBDA(x, x + 1)\n      + LAMBDA(x, x + 2)");

    // No --lib: the project's setting, relative to the project folder.
    const cfg = join(project, "xln.config.json");
    writeFileSync(cfg, JSON.stringify({ ...JSON.parse(readFileSync(cfg, "utf8")), library: "../lib" }));
    const j = capture();
    expect(await main(["lib", "status", project, "--json"], j)).toBe(0);
    const r = JSON.parse(j.stdout);
    expect(r.ok).toBe(true);
    expect(r.kind).toBe("project");
    expect(r.items.filter((i: { name: string }) => i.name.startsWith("FN.")).map((i: { name: string; state: string }) => `${i.name} ${i.state}`)).toEqual([
      "FN.INC differs",
      "FN.TWICE identical",
      "FN.OWN local-only",
    ]);
    expect(r.items[0].file).toBe("names/FN.xln");
  });

  it("publish: --dry-run shows the diff and writes nothing; then a new file and an update; status is identical after", async () => {
    const dry = capture();
    expect(await main(["lib", "publish", project, "FN.OWN", "--dry-run"], dry)).toBe(0);
    expect(dry.stdout).toContain(`would create ${join(LIB, "FN.OWN.lambda")} (new file)`);
    expect(dry.stdout).toContain("+ # name       FN.OWN");
    expect(existsSync(join(LIB, "FN.OWN.lambda"))).toBe(false);

    const io = capture();
    expect(await main(["lib", "publish", project, "fn.own"], io)).toBe(0);
    expect(readFileSync(join(LIB, "FN.OWN.lambda"), "utf8")).toBe("# name       FN.OWN\n# summary\n# params     x\n\nLAMBDA(x, x - 1)\n");

    const up = capture();
    expect(await main(["lib", "publish", project, "FN.INC", "--json"], up)).toBe(0);
    const u = JSON.parse(up.stdout);
    expect(u).toMatchObject({ ok: true, created: false, changed: ["definition"], written: true });
    // The header's other fields and the rationale stay.
    expect(readFileSync(join(LIB, "FN.INC.lambda"), "utf8")).toBe(lambda("FN.INC", "Adds two.", "LAMBDA(x, x + 1)"));

    // Publish records the published version as the copy's base, in the source.
    expect(u.base).toMatchObject({ file: "names/FN.xln", changed: true });
    const fnText = readFileSync(join(project, "names", "FN.xln"), "utf8");
    expect(fnText).toContain(`@from(lib #${u.base.hash})\nFN.INC = LAMBDA(x, x + 1);`);
    expect(fnText).toMatch(/@from\(lib #[0-9a-f]{6}\)\nFN\.OWN = /);
    expect(io.stdout).toMatch(/recorded @from\(lib #[0-9a-f]{6}\) on fn\.own in names\/FN\.xln/);

    const again = capture();
    expect(await main(["lib", "publish", project, "FN.INC"], again)).toBe(0);
    expect(again.stdout).toBe(`xln lib publish FN.INC: ${join(LIB, "FN.INC.lambda")} already has this definition; nothing to write\n`);

    const s = capture();
    expect(await main(["lib", "status", wb, "--lib", LIB, "--json"], s)).toBe(0);
    const states = JSON.parse(s.stdout).items.filter((i: { name: string }) => i.name.startsWith("FN.")).map((i: { name: string; state: string }) => `${i.name} ${i.state}`);
    expect(states).toEqual(["FN.INC identical", "FN.OWN identical", "FN.TWICE identical"]);
  });

  it("take: the library's version with its base; a copy with its own edit only with --discard; the states after builds", async () => {
    const statesOf = async (target: string): Promise<Record<string, string>> => {
      const s = capture();
      expect(await main(["lib", "status", target, "--lib", LIB, "--json"], s)).toBe(0);
      return Object.fromEntries(JSON.parse(s.stdout).items.filter((i: { name: string }) => i.name.startsWith("FN.")).map((i: { name: string; state: string }) => [i.name, i.state]));
    };
    const fnFile = join(project, "names", "FN.xln");
    const build = async () => {
      const b = capture();
      expect(await main(["build", wb], b), b.stderr + b.stdout).toBe(0);
    };
    await build();
    // The library moves on: FN.TWICE was never given a base (written by hand): differs.
    writeFileSync(join(LIB, "FN.TWICE.lambda"), lambda("FN.TWICE", "Doubles.", "LAMBDA(x, x + x)"));
    expect((await statesOf(project))["FN.TWICE"]).toBe("differs");
    const refused = capture();
    expect(await main(["lib", "take", project, "FN.TWICE"], refused)).toBe(1);
    expect(refused.stderr).toContain("refused: FN.TWICE records no library base");
    const taken = capture();
    expect(await main(["lib", "take", project, "FN.TWICE", "--discard"], taken)).toBe(0);
    expect(taken.stdout).toMatch(/updated names\/FN\.xln \(differs → the library's version, @from\(lib #[0-9a-f]{6}\)\)/);
    await build();
    expect((await statesOf(wb))["FN.TWICE"]).toBe("identical");
    // A local edit, built: modified (not outdated), and Take refuses to discard it.
    writeFileSync(fnFile, readFileSync(fnFile, "utf8").replace("FN.TWICE = LAMBDA(x, x + x);", "FN.TWICE = LAMBDA(x, 2 * x);"));
    await build();
    expect((await statesOf(wb))["FN.TWICE"]).toBe("modified");
    expect((await statesOf(project))["FN.TWICE"]).toBe("modified");
    const no = capture();
    expect(await main(["lib", "take", project, "FN.TWICE"], no)).toBe(1);
    expect(no.stderr).toContain("FN.TWICE was edited here since its library base (modified)");
    // The library changes too: both changed, with the base's text from the build's backup.
    writeFileSync(join(LIB, "FN.TWICE.lambda"), lambda("FN.TWICE", "Doubles.", "LAMBDA(x, x * 2 + 0)"));
    const s = capture();
    expect(await main(["lib", "status", project, "--lib", LIB], s)).toBe(0);
    expect(s.stdout).toMatch(/both changed \(1\)/);
    expect(s.stdout).toMatch(/here, since the base #[0-9a-f]{6} \(- base, \+ copy\):\n {8}- LAMBDA\(x, x \+ x\)\n {8}\+ LAMBDA\(x, 2 \* x\)/);
    // The local edit undone: the copy is its base again, the library moved: outdated.
    writeFileSync(fnFile, readFileSync(fnFile, "utf8").replace("FN.TWICE = LAMBDA(x, 2 * x);", "FN.TWICE = LAMBDA(x, x + x);"));
    expect((await statesOf(project))["FN.TWICE"]).toBe("outdated");
    const ok = capture();
    expect(await main(["lib", "take", project, "FN.TWICE", "--dry-run"], ok)).toBe(0);
    expect(ok.stdout).toContain("would update names/FN.xln (outdated → the library's version");
  });

  it("says what is wrong", async () => {
    const io = capture();
    expect(await main(["lib", "status", join(tmp, "nope.xlsx"), "--lib", LIB], io)).toBe(2);
    expect(io.stderr).toContain("no such file or folder");
    const n = capture();
    expect(await main(["lib", "publish", project, "FN.NOPE", "--lib", LIB], n)).toBe(2);
    expect(n.stderr).toContain("FN.NOPE is not a workbook-scoped name");
    const u = capture();
    expect(await main(["lib", "frobnicate"], u)).toBe(2);
    expect(u.stderr).toContain("usage: xln lib status");
    const nolib = capture();
    copyFileSync(PROBE, join(tmp, "other.xlsx"));
    expect(await main(["lib", "status", join(tmp, "other.xlsx")], nolib)).toBe(2);
    expect(nolib.stderr).toContain("no library: pass --lib <dir>");
  });

  it("helpers: ~ and the file diff", () => {
    expect(expandHome("~/x")).toMatch(/[/\\]x$/);
    expect(expandHome("a/~")).toBe("a/~");
    expect(fileDiff("a\nb\nc\nd\ne\n", "a\nb\nC\nd\ne\n")).toBe("  …\n  b\n- c\n+ C\n  d\n  …");
  });
});

const CORPUS = process.env["XLN_CORPUS"];
const REAL = CORPUS ? join(CORPUS, "_shared", "lib") : "";

describe.skipIf(!CORPUS || !existsSync(REAL))("xln lib on a copy of the real library (XLN_CORPUS)", () => {
  it("status of lbo-ep03r's pulled project, then publish to the copy and back to identical", async () => {
    const lib = join(tmp, "reallib");
    cpSync(REAL, lib, { recursive: true });
    const before = Object.fromEntries(readdirSync(REAL).map((f) => [f, readFileSync(join(REAL, f), "utf8")]));
    const wb = join(tmp, "lbo-ep03r.xlsx");
    copyFileSync(join(CORPUS!, "lbo-ep03r", "dist", "lbo-ep03r.xlsx"), wb);
    expect(await main(["pull", wb], capture())).toBe(0);
    const project = join(tmp, "lbo-ep03r.xln");
    const io = capture();
    expect(await main(["lib", "status", project, "--lib", lib, "--no-diff"], io)).toBe(0);
    // The library grows (the author's own publishes): missing are the 3 FN.* of M4 and whatever came since.
    const n = readdirSync(REAL).filter((f) => f.endsWith(".lambda")).length;
    expect(io.stdout).toContain(`outdated 0 · modified 0 · both changed 0 · differs 0 · missing ${n - 14} · identical 14 · local only 0 (${n} library functions)`);
    expect(io.stdout).toContain("other module LAMBDAs, not library candidates: 8 (IN 8)");
    // Publishing an unchanged function writes nothing; a changed one keeps the header.
    const same = capture();
    expect(await main(["lib", "publish", project, "FN.GROW", "--lib", lib], same)).toBe(0);
    expect(same.stdout).toContain("nothing to write");
    const fnFile = join(project, "names", "FN.xln");
    const src = readFileSync(fnFile, "utf8");
    writeFileSync(fnFile, src.replace("prev * (1 + g)", "prev * (1 + g) * 1"));
    const pub = capture();
    expect(await main(["lib", "publish", project, "FN.GROW", "--lib", lib], pub)).toBe(0);
    expect(pub.stdout).toContain("(definition)");
    const after = readFileSync(join(lib, "FN.GROW.lambda"), "utf8");
    expect(after.slice(0, after.indexOf("\nLAMBDA("))).toBe(before["FN.GROW.lambda"]!.slice(0, before["FN.GROW.lambda"]!.indexOf("\nLAMBDA(")));
    const s = capture();
    expect(await main(["lib", "status", project, "--lib", lib, "--json"], s)).toBe(0);
    expect(JSON.parse(s.stdout).items.find((i: { name: string }) => i.name === "FN.GROW").state).toBe("identical");
    // The real library is untouched.
    for (const [f, t] of Object.entries(before)) expect(readFileSync(join(REAL, f), "utf8")).toBe(t);
  });
});

describe("xln lib: room for the provenance tag (feedback 2026-10-07)", () => {
  const proj = join(tmp, "tagroom.xln");
  const lib = join(tmp, "tagroom-lib");
  // Fits alone (and with a tag without lib#), not with the base Publish records.
  const DOC = "x".repeat(235);

  it("publish warns that the doc comment with its tag passes 255, and writes anyway", async () => {
    mkdirSync(join(proj, "names"), { recursive: true });
    mkdirSync(lib);
    writeFileSync(join(proj, "names", "MTG.xln"), `/** ${DOC} */\nMTG.REPAYMENT = LAMBDA(x, x);\n`);
    const io = capture();
    expect(await main(["lib", "publish", proj, "MTG.REPAYMENT", "--lib", lib], io)).toBe(0);
    expect(existsSync(join(lib, "MTG.REPAYMENT.lambda"))).toBe(true);
    const hash = /@from\(lib #([0-9a-f]{6})\)/.exec(readFileSync(join(proj, "names", "MTG.xln"), "utf8"))![1];
    expect(io.stdout).toContain(
      `warning: names/MTG.xln: MTG.REPAYMENT: the doc comment is 235 characters; with its provenance tag (29) it passes Excel's 255, so the build writes it without the tag and the workbook won't record its library base, @from(lib #${hash}): shorten it by 9 characters\n`,
    );
    const j = capture();
    expect(await main(["lib", "publish", proj, "MTG.REPAYMENT", "--lib", lib, "--json"], j)).toBe(0);
    expect(JSON.parse(j.stdout).warnings).toEqual([expect.stringContaining("shorten it by 9 characters")]);
  });

  it("take and base say it too; a doc comment with room says nothing", async () => {
    // A library doc within its own limit still leaves too little room for a long version.
    writeFileSync(join(lib, "MTG.SCHEDULE.lambda"), `# name       MTG.SCHEDULE\n# summary    ${"Lays out the schedule ".repeat(12)}.\n# params     x\n\nLAMBDA(x, x + 1)\n`);
    writeFileSync(join(proj, "names", "MTG.xln"), `// @version 2026.10.07-release-candidate\n\n/** Short. */\nMTG.SCHEDULE = LAMBDA(x, x);\n`);
    const t = capture();
    expect(await main(["lib", "take", proj, "MTG.SCHEDULE", "--lib", lib, "--discard"], t)).toBe(0);
    expect(t.stdout).toMatch(/\nwarning: names\/MTG\.xln: MTG\.SCHEDULE: the doc comment is 208 characters \(a line break counts 2\); with its provenance tag \(58\) it passes Excel's 255/);

    writeFileSync(join(proj, "names", "MTG.xln"), `/** ${DOC} */\nMTG.SCHEDULE = LAMBDA(x, x + 1);\n`);
    const b = capture();
    expect(await main(["lib", "base", proj, "MTG.SCHEDULE", "--lib", lib], b), b.stderr).toBe(0);
    expect(b.stdout).toContain("warning: names/MTG.xln: MTG.SCHEDULE: the doc comment is 235 characters; with its provenance tag (29)");

    writeFileSync(join(proj, "names", "MTG.xln"), `/** Short. */\nMTG.REPAYMENT = LAMBDA(x, x);\n`);
    const ok = capture();
    expect(await main(["lib", "publish", proj, "MTG.REPAYMENT", "--lib", lib], ok)).toBe(0);
    expect(ok.stdout).not.toContain("warning:");
  });
});
