// `xln rename` (M5) on a copy of probe F7's workbook: the source edits, the build that
// rewrites the cells and removes the @renamed line it applied, a re-pull that changes
// nothing; refusals write nothing.
import { copyFileSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { readWorkbook, replaceZipEntries, utf8 } from "@xln/core";
import { main } from "../src/main.js";

const F7 = join(import.meta.dirname, "..", "..", "..", "probes", "results", "f7_base.xlsx");
const tmp = mkdtempSync(join(tmpdir(), "xln-rename-"));
afterAll(() => rmSync(tmp, { recursive: true, force: true }));

function capture() {
  const io = { stdout: "", stderr: "", out: (s: string) => void (io.stdout += s), err: (s: string) => void (io.stderr += s) };
  return io;
}

function tree(dir: string, rel = ""): Record<string, string> {
  const out: Record<string, string> = {};
  for (const e of readdirSync(join(dir, rel), { withFileTypes: true })) {
    const r = rel ? `${rel}/${e.name}` : e.name;
    if (e.isDirectory()) Object.assign(out, tree(dir, r));
    else out[r] = readFileSync(join(dir, r), "utf8");
  }
  return out;
}

describe("xln rename", () => {
  const wb = join(tmp, "f7.xlsx");
  const project = join(tmp, "f7.xln");

  it("renames in the source, the build rewrites the cells, a re-pull drops only @renamed", async () => {
    copyFileSync(F7, wb);
    expect(await main(["pull", wb], capture())).toBe(0);
    const before = tree(join(project, "names"));

    const dry = capture();
    expect(await main(["rename", project, "Rate", "Pace", "--dry-run"], dry), dry.stderr).toBe(0);
    expect(dry.stdout).toContain("(dry run: nothing written)");
    expect(dry.stdout).toContain("the next build will rename Rate → Pace, rewriting it in 8 cell formulas, 1 conditional format, 1 validation, 2 names");
    expect(tree(join(project, "names"))).toEqual(before);

    const io = capture();
    expect(await main(["rename", wb, "Rate", "Pace"], io), io.stderr).toBe(0);
    expect(io.stdout).toContain("names/_unmanaged.xln: the name, @renamed(Rate), 2 references");
    expect(io.stdout).toContain("wrote names/_unmanaged.xln, names/sheets/S1.xln, names/sheets/S2.xln");
    const renamed = tree(join(project, "names"));
    const withoutRenamed = Object.fromEntries(Object.entries(renamed).map(([k, t]) => [k, t.split("\n").filter((l) => !l.startsWith("@renamed(")).join("\n")]));

    // A dry run and a build into another file leave the source as it is.
    expect(await main(["build", wb, "--dry-run"], capture())).toBe(0);
    expect(tree(join(project, "names"))).toEqual(renamed);
    const copy = capture();
    expect(await main(["build", wb, "--out", join(tmp, "copy.xlsx")], copy), copy.stderr).toBe(0);
    expect(copy.stdout).not.toContain("removed @renamed");
    expect(tree(join(project, "names"))).toEqual(renamed);

    const b = capture();
    expect(await main(["build", wb], b), b.stderr).toBe(0);
    expect(b.stdout).toContain("rename Rate → Pace, rewriting it in 8 cell formulas");
    // The build consumes the @renamed note: the source is what the next pull writes.
    expect(b.stdout).toContain("  removed @renamed(Rate) from names/_unmanaged.xln: the rename is built\n");
    expect(tree(join(project, "names"))).toEqual(withoutRenamed);
    const s1 = readWorkbook(new Uint8Array(readFileSync(wb))).sheets[0]!;
    expect(s1.formulas.map((f) => f.text ?? "")).toContain('"Rate is "&Pace');
    expect(s1.formulas.map((f) => f.text ?? "")).toContain("_xlfn.LET(_xlpm.Rate, 5, _xlpm.Rate*2)");

    const again = capture();
    expect(await main(["build", wb], again)).toBe(0);
    expect(again.stdout).toContain("up to date");

    // The pull guard finds nothing unbuilt and no file to rewrite.
    const pull = capture();
    expect(await main(["pull", wb], pull), pull.stderr).toBe(0);
    expect(pull.stdout + pull.stderr).not.toMatch(/unbuilt|rewrit/i);
    const pulled = tree(join(project, "names"));
    expect(pulled).toEqual(withoutRenamed);
  });

  it("renamed back after a build: X → Y, build, Y → X, build leaves no @renamed", async () => {
    const before = tree(join(project, "names"));
    const there = capture();
    expect(await main(["rename", wb, "RateX", "RateY"], there), there.stderr).toBe(0);
    expect(await main(["build", wb], capture())).toBe(0);
    expect(Object.values(tree(join(project, "names"))).join("")).not.toContain("@renamed");
    const back = capture();
    expect(await main(["rename", wb, "RateY", "RateX"], back), back.stderr).toBe(0);
    expect(tree(join(project, "names"))["_unmanaged.xln"]).toContain("@renamed(RateY)\nRateX = 0.5;");
    const b = capture();
    expect(await main(["build", wb, "--json"], b), b.stderr).toBe(0);
    expect((JSON.parse(b.stdout) as { renamedRemoved: string[] }).renamedRemoved).toEqual(["removed @renamed(RateY) from names/_unmanaged.xln: the rename is built"]);
    expect(tree(join(project, "names"))).toEqual(before);
    expect(await main(["pull", wb], capture())).toBe(0);
    expect(tree(join(project, "names"))).toEqual(before);
  });

  it("refuses an invalid or taken name, and a capture, writing nothing", async () => {
    const before = tree(join(project, "names"));
    for (const [to, msg] of [
      ["A1", "cell reference"],
      ["RateX", "exists already"],
      ["x", "would change what this formula reads"],
    ] as const) {
      const io = capture();
      expect(await main(["rename", project, "Pace", to], io)).toBe(1);
      expect(io.stderr).toContain(msg);
    }
    const unknown = capture();
    expect(await main(["rename", project, "Nope", "Other"], unknown)).toBe(1);
    expect(unknown.stderr).toContain("no name Nope in the project");
    expect(tree(join(project, "names"))).toEqual(before);
  });

  it("--json lists the edits; bad usage is exit 2", async () => {
    const io = capture();
    expect(await main(["rename", project, "S2!Loc", "Spot", "--dry-run", "--json"], io)).toBe(0);
    const j = JSON.parse(io.stdout) as { ok: boolean; from: string; to: string; annotation: string; edits: { path: string; label: string }[]; build: string[] };
    expect(j).toMatchObject({ ok: true, from: "S2!Loc", to: "S2!Spot", annotation: "added", build: ["rename S2!Loc → Spot, rewriting it in 5 cell formulas"] });
    expect(j.edits.filter((e) => e.label === "reference").length).toBe(5);
    expect(await main(["rename", project, "Pace"], capture())).toBe(2);
  });
});

describe("the label notice after a rename (2026-10-07)", () => {
  // f7_base named with Create from Selection over S2: C the readable text, D the name's text, E:F the cells.
  function labelsBook(): Uint8Array {
    const f7 = new Uint8Array(readFileSync(F7));
    const files = unzipSync(f7);
    const cell = (r: string, text: string) => `<c r="${r}" t="inlineStr"><is><t>${text}</t></is></c>`;
    const rows = [
      ["7", "Revenue", "Revenue"],
      ["8", "Cost of goods", "COGS"],
      ["9", "Gross income", "Gross_income"],
    ].map(([r, text, label]) => `<row r="${r}">${cell(`C${r}`, text!)}${cell(`D${r}`, label!)}<c r="E${r}"><v>1</v></c><c r="F${r}"><v>2</v></c></row>`);
    const sheet = strFromU8(files["xl/worksheets/sheet2.xml"]!).replace("</sheetData>", `${rows.join("")}</sheetData>`);
    const defs = [["COGS", 8], ["Gross_income", 9], ["Revenue", 7]].map(([n, r]) => `<definedName name="${n}">S2!$E$${r}:$F$${r}</definedName>`);
    const book = strFromU8(files["xl/workbook.xml"]!).replace("<definedNames>", `<definedNames>${defs.join("")}`);
    return replaceZipEntries(f7, new Map([["xl/workbook.xml", utf8(book)], ["xl/worksheets/sheet2.xml", utf8(sheet)]]));
  }

  it("xln rename gives it for after the build; xln build prints it after the removed @renamed", async () => {
    const dir = join(tmp, "labels");
    mkdirSync(dir);
    const wb = join(dir, "labels.xlsx");
    writeFileSync(wb, labelsBook());
    expect(await main(["pull", wb], capture())).toBe(0);
    const notice = [
      "Gross_income → Gross_ind_income: 2 labels still read the old name: 'S2'!C9, 'S2'!D9",
      "In Excel, on sheet S2: Find & Replace (Ctrl+H on Windows; Ctrl+H or ⌘⇧H on Mac; or Home → Find & Select → Replace)",
      "  1. Find what:     Gross income",
      "     Replace with:  Gross ind income",
      "  2. Find what:     Gross_income",
      "     Replace with:  Gross_ind_income",
      "  Within: Sheet · Look in: Formulas · Match entire cell contents ✓ · then Replace All (for each pair)",
    ];
    const r = capture();
    expect(await main(["rename", wb, "Gross_income", "Gross_ind_income"], r), r.stderr).toBe(0);
    expect(r.stdout).toContain(["  after the build, in Excel (xln never writes cell values):", ...notice.map((l) => `    ${l}`)].join("\n") + "\n");
    const b = capture();
    expect(await main(["build", wb], b), b.stderr).toBe(0);
    expect(b.stdout).toContain(["  removed @renamed(Gross_income) from names/sheets/S2.xln: the rename is built", ...notice.map((l) => `  ${l}`)].join("\n") + "\n");
    // A rename of a name not on cells says nothing about labels.
    const j = capture();
    expect(await main(["rename", wb, "Rate", "Pace", "--dry-run", "--json"], j)).toBe(0);
    expect(JSON.parse(j.stdout).labelNotice).toEqual([]);
  });
});

// A note whose rename is in the workbook already (an older build, a browser build's copy, a
// hand): `xln check` gives the hint with its quick fix; a build with nothing to write leaves
// it; the next build that writes removes it with the same line as its own (2026-10-07).
describe("a spent @renamed left in the source", () => {
  it("xln check hints it; an up-to-date build leaves it; a writing build removes it", async () => {
    const dir = join(tmp, "stale");
    mkdirSync(dir);
    const wb = join(dir, "f7.xlsx");
    const project = join(dir, "f7.xln");
    copyFileSync(F7, wb);
    expect(await main(["pull", wb], capture())).toBe(0);
    expect(await main(["rename", wb, "Rate", "Pace"], capture())).toBe(0);
    expect(await main(["build", wb], capture())).toBe(0);
    // Put back as an older build left it.
    const u = join(project, "names", "_unmanaged.xln");
    const built = readFileSync(u, "utf8");
    writeFileSync(u, built.replace("Pace = 0.1;", "@renamed(Rate)\nPace = 0.1;"));
    const line = readFileSync(u, "utf8").split("\n").indexOf("@renamed(Rate)") + 1;

    const c = capture();
    expect(await main(["check", wb, "--only", "C1"], c), c.stderr).toBeLessThan(2);
    expect(c.stdout).toContain(`  hint    names/_unmanaged.xln:${line}:1 renamed-built: @renamed(Rate): the rename is built; this line can go [quick fix: Remove @renamed(Rate)]`);

    const same = capture();
    expect(await main(["build", wb], same), same.stderr).toBe(0);
    expect(same.stdout).toContain("up to date");
    expect(same.stdout).not.toContain("removed @renamed");
    expect(readFileSync(u, "utf8")).toContain("@renamed(Rate)\nPace = 0.1;");

    expect(await main(["rename", wb, "S2!Loc", "Spot"], capture())).toBe(0);
    const b = capture();
    expect(await main(["build", wb], b), b.stderr).toBe(0);
    expect(b.stdout).toContain("  removed @renamed(Rate) from names/_unmanaged.xln: the rename is built\n");
    expect(b.stdout).toContain("  removed @renamed(Loc) from names/sheets/S2.xln: the rename is built\n");
    expect(readFileSync(u, "utf8")).toBe(built);
    expect(Object.values(tree(join(project, "names"))).join("")).not.toContain("@renamed");
    const after = capture();
    expect(await main(["check", wb, "--only", "C1"], after)).toBeLessThan(2);
    expect(after.stdout).not.toContain("renamed-built");
  });
});

describe("xln rename to an Excel 4.0 macro function's name", () => {
  it("warns, and renames all the same", async () => {
    const wb = join(tmp, "xlm.xlsx");
    const project = join(tmp, "xlm.xln");
    copyFileSync(F7, wb);
    expect(await main(["pull", wb], capture())).toBe(0);
    const io = capture();
    expect(await main(["rename", wb, "Fn", "Evaluate", "--dry-run"], io), io.stderr).toBe(0);
    expect(io.stdout).toContain("  warning: Evaluate is also an Excel 4.0 macro function: Excel may call that instead or refuse the name (AFE #10, not measured)\n");
    const j = capture();
    expect(await main(["rename", wb, "Fn", "Evaluate", "--json"], j), j.stderr).toBe(0);
    const out = JSON.parse(j.stdout) as { ok: boolean; warnings: string[] };
    expect(out.ok).toBe(true);
    expect(out.warnings).toEqual(["Evaluate is also an Excel 4.0 macro function: Excel may call that instead or refuse the name (AFE #10, not measured)"]);
    expect(readFileSync(join(project, "names", "_unmanaged.xln"), "utf8")).toContain("@renamed(Fn)\nEvaluate = LAMBDA(");
    // A value is no call: renaming Rate to Files says nothing.
    const v = capture();
    expect(await main(["rename", wb, "Rate", "Files", "--dry-run"], v), v.stderr).toBe(0);
    expect(v.stdout).not.toContain("warning:");
  });
});
