// The author's three ideas on the command line (FEEDBACK 2026-10-07): `library-bases/` kept
// by publish, take and base, read by status for the three-way diff, left alone by pull;
// `xln lib base` (Record library base); `@param` drift in `xln check`'s source findings.
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { main } from "../src/main.js";

const PROBE = join(import.meta.dirname, "..", "..", "..", "probes", "results", "probe_win.xlsx");
const tmp = mkdtempSync(join(tmpdir(), "xln-lib-ideas-"));
afterAll(() => rmSync(tmp, { recursive: true, force: true }));

function capture() {
  const io = { stdout: "", stderr: "", out: (s: string) => void (io.stdout += s), err: (s: string) => void (io.stderr += s) };
  return io;
}

const LIB = join(tmp, "lib");
const lambda = (name: string, summary: string, def: string) => `# name       ${name}\n# summary    ${summary}\n# params     x\n#\n# Why ${name}.\n\n${def}\n`;

describe("xln lib: kept bases and Record library base", () => {
  const wb = join(tmp, "model.xlsx");
  const project = join(tmp, "model.xln");
  const fnFile = join(project, "names", "FN.xln");
  const bases = () => (existsSync(join(project, "library-bases")) ? readdirSync(join(project, "library-bases")).sort() : []);

  it("status notes identical copies without a base; lib base records them, by name or --all, never anything else", async () => {
    mkdirSync(LIB);
    writeFileSync(join(LIB, "FN.TWICE.lambda"), lambda("FN.TWICE", "Doubles.", "LAMBDA(x, x * 2)"));
    writeFileSync(join(LIB, "FN.DBL.lambda"), lambda("FN.DBL", "Doubles too.", "LAMBDA(x, x + x)"));
    writeFileSync(join(LIB, "FN.INC.lambda"), lambda("FN.INC", "Adds two.", "LAMBDA(x, x + 2)"));
    copyFileSync(PROBE, wb);
    expect(await main(["pull", wb], capture())).toBe(0);
    const src = "/** Doubles. */\nFN.TWICE = LAMBDA(x, x*2);\n\nFN.DBL = LAMBDA(x, x + x);\n\nFN.INC = LAMBDA(x, x + 1);\n";
    writeFileSync(fnFile, src);

    const s = capture();
    expect(await main(["lib", "status", project, "--lib", LIB, "--json"], s)).toBe(0);
    const items = JSON.parse(s.stdout).items as { name: string; state: string; note: string; noBase?: boolean }[];
    expect(items.find((i) => i.name === "FN.TWICE")).toMatchObject({ state: "identical", noBase: true });
    expect(items.find((i) => i.name === "FN.TWICE")!.note).toContain("no base recorded: Record library base");
    // Status wrote nothing.
    expect(readFileSync(fnFile, "utf8")).toBe(src);
    expect(bases()).toEqual([]);

    // One that differs: refused with the reason, exit 1.
    const no = capture();
    expect(await main(["lib", "base", project, "FN.INC", "--lib", LIB], no)).toBe(1);
    expect(no.stderr).toContain("skipped FN.INC: differs, not identical to the library");

    // --dry-run: says what, writes nothing.
    const dry = capture();
    expect(await main(["lib", "base", project, "--all", "--lib", LIB, "--dry-run"], dry)).toBe(0);
    expect(dry.stdout).toContain("xln lib base --all --dry-run: would record the library base of 2 functions");
    expect(readFileSync(fnFile, "utf8")).toBe(src);
    expect(bases()).toEqual([]);

    const all = capture();
    expect(await main(["lib", "base", project, "--all", "--lib", LIB], all)).toBe(0);
    expect(all.stdout).toMatch(/recorded the library base of 2 functions; build to carry it into the workbook\n {2}FN\.DBL {2}@from\(lib #[0-9a-f]{6}\) {2}names\/FN\.xln {2}\(text kept in library-bases\/[0-9a-f]{6}\.json\)\n {2}FN\.TWICE /);
    const after = readFileSync(fnFile, "utf8");
    expect(after).toMatch(/^\/\*\* Doubles\. \*\/\n@from\(lib #[0-9a-f]{6}\)\nFN\.TWICE = LAMBDA\(x, x\*2\);\n\n@from\(lib #[0-9a-f]{6}\)\nFN\.DBL = LAMBDA\(x, x \+ x\);\n\nFN\.INC = LAMBDA\(x, x \+ 1\);\n$/);
    expect(bases().length).toBe(2);
    const again = capture();
    expect(await main(["lib", "base", project, "--all", "--lib", LIB, "--json"], again)).toBe(0);
    expect(JSON.parse(again.stdout)).toMatchObject({ ok: true, recorded: [], written: false });
  });

  it("take keeps the base's text; with the library changed and the copy edited, status diffs both sides from it", async () => {
    const t = capture();
    expect(await main(["lib", "take", project, "FN.INC", "--lib", LIB, "--discard", "--json"], t)).toBe(0);
    const taken = JSON.parse(t.stdout) as { base: string; kept: string[] };
    expect(taken.kept).toEqual([`library-bases/${taken.base}.json`]);
    expect(bases()).toContain(`${taken.base}.json`);
    // Both sides move; neither the workbook (never built) nor a backup has the base's text.
    writeFileSync(fnFile, readFileSync(fnFile, "utf8").replace("FN.INC = LAMBDA(x, x + 2)", "FN.INC = LAMBDA(x, x + 3)"));
    writeFileSync(join(LIB, "FN.INC.lambda"), lambda("FN.INC", "Adds two.", "LAMBDA(x, ROUND(x + 2, 0))"));
    const s = capture();
    expect(await main(["lib", "status", project, "--lib", LIB], s)).toBe(0);
    expect(s.stdout).toContain(`the base #${taken.base}'s text: from library-bases/${taken.base}.json`);
    expect(s.stdout).toContain(`here, since the base #${taken.base} (- base, + copy):\n        - LAMBDA(x, x + 2)\n        + LAMBDA(x, x + 3)`);
    expect(s.stdout).toContain("        - LAMBDA(x, x + 2)\n        + LAMBDA(x, ROUND(x + 2, 0))");
  });

  it("publish keeps the version it wrote; a pull leaves library-bases/ alone", async () => {
    writeFileSync(fnFile, readFileSync(fnFile, "utf8") + "\nFN.NEW = LAMBDA(x, x - 7);\n");
    const p = capture();
    expect(await main(["lib", "publish", project, "FN.NEW", "--lib", LIB, "--json"], p)).toBe(0);
    const pub = JSON.parse(p.stdout) as { base: { hash: string }; kept: string[] };
    expect(pub.kept).toEqual([`library-bases/${pub.base.hash}.json`]);
    const before = bases();
    expect(before).toContain(`${pub.base.hash}.json`);
    const pull = capture();
    expect(await main(["pull", wb, "--discard"], pull), pull.stderr).toBe(0);
    // names/** is the workbook's again (never built: no FN module), the kept bases stay.
    expect(existsSync(fnFile)).toBe(false);
    expect(bases()).toEqual(before);
  });
});

describe("xln check: @param drift in the source findings", () => {
  it("a stale @param is a warning naming the rename fix; a parameter left out a hint", async () => {
    const w = join(tmp, "doc.xlsx");
    copyFileSync(PROBE, w);
    expect(await main(["pull", w], capture())).toBe(0);
    writeFileSync(join(tmp, "doc.xln", "names", "FN.xln"), "/**\n * Spreads.\n * @param total the amount\n * @param periods how many\n */\nFN.SPREAD = LAMBDA(total, periodi, total / periodi);\n\n/** Adds. @param a */\nFN.ADD = LAMBDA(a, b, a + b);\n");
    const io = capture();
    await main(["check", w, "--json"], io);
    const findings = (JSON.parse(io.stdout) as { source: { findings: { severity: string; code: string; message: string; line: number; fix?: { title: string } }[] } }).source.findings;
    const doc = findings.filter((f) => f.code.startsWith("doc-param"));
    expect(doc.map((f) => [f.severity, f.code, f.line])).toEqual([
      ["warning", "doc-param", 4],
      ["hint", "doc-param-missing", 8],
    ]);
    expect(doc[0]!.fix?.title).toBe("Rename @param periods to periodi");
    expect(doc[1]!.message).toContain("but not b");
  });

  it("a doc comment with no room for its provenance tag: the checker's warning, as in the editor", async () => {
    const w = join(tmp, "tagroom.xlsx");
    copyFileSync(PROBE, w);
    expect(await main(["pull", w], capture())).toBe(0);
    writeFileSync(join(tmp, "tagroom.xln", "names", "MTG.xln"), `/** ${"x".repeat(230)} */\n@from(lib #353921)\nMTG.REPAYMENT = LAMBDA(x, x);\n`);
    const io = capture();
    await main(["check", w, "--json"], io);
    const findings = (JSON.parse(io.stdout) as { source: { findings: { severity: string; code: string; message: string; line: number }[] } }).source.findings;
    expect(findings.filter((f) => f.code === "provenance").map((f) => [f.severity, f.line, f.message])).toEqual([
      [
        "warning",
        1,
        "MTG.REPAYMENT: the doc comment is 230 characters; with its provenance tag (29) it passes Excel's 255, so the build writes it without the tag and the workbook won't record its library base, @from(lib #353921): shorten it by 4 characters",
      ],
    ]);
  });
});
