import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { applyChangeSet, LOCK_FILE, MANIFEST_FILE, parseLockfile, pullProject, readWorkbook, rewrittenFiles, unbuiltEdits } from "@xln/core";
import { describe, expect, it } from "vitest";
import { addressAt, buildFormulaView, entryForCell, entryHover, entryLabel, entryLinks, lhsAt, nameAt, refAt, WORKBOOK_VIEW } from "../src/model/formulaView.js";
import { LineIndex } from "../src/model/lines.js";
import { parseManifest } from "../src/model/manifest.js";
import { isNamesFile, Project, type Loc } from "../src/model/project.js";
import { formatPullSummary, isProjectListing, isWorkbookName, projectFolderName, pullNotification, rewrittenQuestion, unbuiltQuestion } from "../src/model/pull.js";
import { formatValue, hoverMarkdown, renderUsages } from "../src/model/usages.js";

const f7 = fileURLToPath(new URL("../../../probes/results/f7_base.xlsx", import.meta.url));

function projectFrom(files: Record<string, string>): Project {
  const p = new Project("mem:/p", files[MANIFEST_FILE] ? parseManifest(files[MANIFEST_FILE]) : undefined);
  for (const [path, text] of Object.entries(files)) if (isNamesFile(path)) p.setFile(path, text);
  if (files[LOCK_FILE]) p.lock = parseLockfile(files[LOCK_FILE]);
  return p;
}

/** Offset of the `n`-th (0-based) occurrence of `needle` in a file, plus `delta`. */
function at(p: Project, path: string, needle: string, n = 0, delta = 1): number {
  const text = p.files.get(path)!.text;
  let i = -1;
  for (let k = 0; k <= n; k++) i = text.indexOf(needle, i + 1);
  if (i < 0) throw new Error(`${needle} not in ${path}`);
  return i + delta;
}

function textOf(p: Project, l: Loc | undefined): string {
  if (!l) return "<none>";
  return `${l.path}:` + p.files.get(l.path)!.text.slice(l.start, l.end);
}

describe("project model on f7_base.xlsx", () => {
  const bytes = new Uint8Array(readFileSync(f7));
  const pulled = pullProject(bytes, "f7_base.xlsx");
  const p = projectFrom(pulled.files);
  const U = "names/_unmanaged.xln";
  const S2 = "names/sheets/S2.xln";
  const S1 = "names/sheets/S1.xln";

  it("loads every name with its scope", () => {
    expect(p.defs.map((d) => d.key)).toEqual(["Fn", "Loc", "Rate", "Rate2", "RateX", "Spl", "S2!Loc"]);
  });

  it("outlines a file's names; a sheet file has no blocks (M3d)", () => {
    expect(p.outline(U).map((s) => s.name)).toEqual(["Fn", "Loc", "Rate", "Rate2", "RateX"]);
    const s2 = p.outline(S2);
    expect(s2.map((s) => [s.name, s.kind, s.children.length])).toEqual([
      ["@A1", "formula", 0],
      ["@B1", "formula", 0],
      ["@A2", "formula", 0],
      ["@A3:A5", "formula", 0],
      ["Loc", "constant", 0],
    ]);
    expect(p.outline(U).find((s) => s.name === "Fn")!.kind).toBe("lambda");
    // Spl = 'S1'!$E$1# is a cell statement in S1's file, @workbook above it, at E1's place.
    const s1 = p.outline(S1);
    expect(s1.map((s) => s.name).slice(0, 4)).toEqual(["@A1", "@C1", "Spl", "@B2:B11"]);
    expect(s1.find((s) => s.name === "Spl")!.detail).toBe("SEQUENCE(3)*Rate");
  });

  it("reads cell statements: names in their formulas, read-only addresses with a quick fix", () => {
    // Go to definition from an unnamed cell's formula.
    expect(textOf(p, p.definition(S1, at(p, S1, "A2*Rate", 0, 4)))).toBe(`${U}:Rate`);
    const spl = p.lookup("Spl")!;
    expect(spl.entry.cell).toMatchObject({ range: "E1" });
    expect(spl.entry).toMatchObject({ cellSheet: "S1", workbook: true });
    expect(spl.scope).toBeUndefined();
    // An edited address: an error, and a fix that puts the pulled address back.
    const q = projectFrom({ ...pulled.files, [S1]: pulled.files[S1]!.replace("Spl @E1# =", "Spl @E2# =").replace("@C8 = Rate*3;", "@C18 = Rate*3;") });
    const probs = q.problems(S1).filter((x) => x.code !== "workbook-on-cell");
    expect(probs.map((x) => x.message)).toEqual([
      "Spl: the address is set in Excel and read-only; the last pull had @E1",
      "no cell statement at S1!C18 in the last pull: addresses are set in Excel (write the formula in Excel, then pull)",
    ]);
    const fix = probs[0]!.fix!;
    const text = q.files.get(S1)!.text;
    expect(text.slice(fix.start, fix.end)).toBe("E2");
    expect(fix.text).toBe("E1");
    expect(probs[1]!.fix).toBeUndefined();
  });

  it("a named statement without # on a formula that spills: a warning on the address, with a quick fix that adds #", () => {
    expect(pulled.files[S1]).toContain("Spl @E1# = SEQUENCE(3)*Rate;");
    const notWorkbook = (x: { code?: string }) => x.code !== "workbook-on-cell";
    expect(p.problems(S1).filter(notWorkbook)).toEqual([]);
    const files = { ...pulled.files, [S1]: pulled.files[S1]!.replace("Spl @E1# =", "Spl @E1 =") };
    const q = projectFrom(files);
    const probs = q.problems(S1).filter(notWorkbook);
    expect(probs.map((x) => [x.severity, x.message])).toEqual([["warning", "Spl @E1: the formula spills over E1:E3, but Spl covers only E1 (write @E1# to name the spill)"]]);
    const text = q.files.get(S1)!.text;
    expect(text.slice(probs[0]!.start, probs[0]!.end)).toBe("@E1");
    const fix = probs[0]!.fix!;
    expect(fix.title).toBe("Name the whole spill: @E1#");
    const fixed = text.slice(0, fix.start) + fix.text + text.slice(fix.end);
    expect(fixed).toContain("Spl @E1# = SEQUENCE(3)*Rate;");
    q.setFile(S1, fixed);
    expect(q.problems(S1).filter(notWorkbook)).toEqual([]);
    // A project pulled before the # was written (lockfile format 3): the build reads the
    // missing # of a name already on the spill as #, so the editor does not warn.
    const old = new Project("mem:/old", parseManifest(pulled.files[MANIFEST_FILE]!));
    old.lock = { ...parseLockfile(pulled.files[LOCK_FILE]!), format: "xln.lock/3" };
    for (const [path, text] of Object.entries(files)) if (isNamesFile(path)) old.setFile(path, text);
    expect(old.problems(S1).filter(notWorkbook)).toEqual([]);
  });

  it("goes to the definition of a name used in a definition", () => {
    expect(textOf(p, p.definition(U, at(p, U, "Rate*2")))).toBe(`${U}:Rate`);
    const d = p.definition(U, at(p, U, "Rate*2"))!;
    expect(d.start).toBe(at(p, U, "Rate = 0.1", 0, 0));
  });

  it("finds references among names and the used-by / uses relation", () => {
    const refs = p.references("Rate").map((l) => textOf(p, l));
    // Names first, then the cell statements that read Rate (the LET variable and the string do not).
    expect(refs.slice(0, 2)).toEqual([`${U}:Rate`, `${U}:Rate`]);
    expect(refs.slice(2)).toEqual([`${S1}:Rate`, ...Array(6).fill(`${S1}:Rate`), `${S2}:Rate`]);
    expect(p.usedBy("Rate").map((u) => u.def.key)).toEqual(["Fn", "Rate2", "Spl"]);
    expect(p.uses(p.lookup("Rate2")!).map((u) => u.def.key)).toEqual(["Rate"]);
  });

  it("knows LAMBDA parameters are not names", () => {
    const h = p.hit(U, at(p, U, "x*Rate"));
    expect(h?.target !== "self" && h?.target.kind).toBe("local");
    expect(textOf(p, p.definition(U, at(p, U, "x*Rate")))).toBe(`${U}:x`);
  });

  it("lists cell usages from the manifest with formulas from the workbook", () => {
    const wb = readWorkbook(bytes);
    const doc = renderUsages("Rate", p.manifestName("Rate"), p.manifest, wb);
    const lines = doc.text.split("\n");
    const cells = doc.items.filter((i) => i.kind === "cell").map((i) => lines[i.line]!);
    expect(cells[0]).toBe("'S1'!A1: =Rate*2  → 0.2"); // quoted as Excel does: S1 looks like a cell
    expect(cells.some((l) => l.startsWith("'S2'!A2: ="))).toBe(true);
    expect(doc.items.map((i) => i.kind)).toContain("conditionalFormat");
    expect(doc.items.map((i) => i.kind)).toContain("dataValidation");
    for (const i of doc.items) expect(lines[i.line]!.slice(i.start, i.end)).toMatch(/^'S[12]'!/);
    // Without the workbook the places are still there.
    const bare = renderUsages("Rate", p.manifestName("Rate"), p.manifest, undefined);
    expect(bare.items.length).toBe(doc.items.length);
  });

  it("hovers with kind, scope, arity, spill anchor, extent and first value", () => {
    const wb = readWorkbook(bytes);
    const fn = hoverMarkdown(p, p.lookup("Fn")!, wb);
    expect(fn).toContain("Fn(x)");
    expect(fn).toContain("LAMBDA · workbook scope · arity 1");
    expect(fn).toContain("uses `Rate`");
    expect(fn).toContain("used in 2 cells");
    const spl = hoverMarkdown(p, p.lookup("Spl")!, wb);
    expect(spl).toContain("anchored at `'S1'!E1`");
    expect(spl).toContain("last saved extent: `E1:E3`");
    expect(spl).toMatch(/anchor formula: `=.+`/);
    expect(spl).toMatch(/first value \(cached\): `.+`/);
    expect(hoverMarkdown(p, p.lookup("S2!Loc")!, wb)).toContain("local to sheet S2");
  });

  it("searches names, definitions and doc comments", () => {
    expect(p.search("rate").map((h) => h.def.key).slice(0, 3)).toEqual(["Rate", "Rate2", "RateX"]);
    expect(p.search("sequence").map((h) => [h.def.key, h.matched])).toEqual([["Spl", ["formula"]]]);
    expect(p.search("lambda rate").map((h) => h.def.key)).toEqual(["Fn"]);
  });

  it("reports no problems in a fresh pull, but the hint on a workbook name read only on its own sheet", () => {
    for (const path of p.files.keys()) {
      expect(p.problems(path).map((x) => [x.severity, x.code, x.message]), path).toEqual(
        path === S1 ? [["hint", "workbook-on-cell", "workbook name on a cell of S1, read only on S1: remove @workbook to make it local to S1 (the build then moves it)"]] : [],
      );
    }
  });

  it("shows a shared child's own text in its cell usage, and keeps sheet and range for go-to-cell", () => {
    const wb = readWorkbook(bytes);
    const doc = renderUsages("S2!Loc", p.manifestName("S2!Loc"), p.manifest, wb);
    const lines = doc.text.split("\n");
    const cells = doc.items.filter((i) => i.kind === "cell");
    expect(cells.map((i) => [i.sheet, i.range])).toEqual([["S1", "C5"], ["S2", "A1:B1"], ["S2", "A2:A5"]]);
    const s1 = renderUsages("Rate", p.manifestName("Rate"), p.manifest, wb);
    expect(s1.text).toContain("'S1'!B3:B7: =A3*Rate  → 0.2  (formula of B3)");
    expect(lines.length).toBeGreaterThan(3);
  });

  it("formula view: names linked to the project, cells found inside spills", () => {
    const wb = readWorkbook(bytes);
    const view = buildFormulaView(wb, "s1", "f7_base.xlsx", p)!;
    expect(view.sheet).toBe("S1");
    expect(view.linked).toBe(true);
    const b3 = view.text.indexOf("A3*Rate");
    const hit = nameAt(view, b3 + 4)!;
    expect(hit.line.cell).toBe("B3");
    expect(hit.line.names[hit.k]!.key).toBe("Rate");
    expect(p.lookup(hit.line.names[hit.k]!.key!)!.file.path).toBe(U);
    expect(nameAt(view, b3 + 1)).toBeUndefined();
    const ref = refAt(view, b3 + 1)!;
    expect(ref.line.refs[ref.k]).toMatchObject({ sheet: "S1", address: "A3" });
    // E2 lies in E1's spill; C5 reads the S2-local Loc first.
    expect(view.lines[entryForCell(view, "E2")!.index]!.cell).toBe("E1");
    expect(view.lines[entryForCell(view, "$C$5")!.index]!.names.map((n) => n.key)).toEqual(["S2!Loc", "Loc"]);
    expect(entryForCell(view, "Z99")).toBeUndefined();
    const e1 = entryForCell(view, "E1")!;
    expect(addressAt(view, e1.address.start)).toBe(e1);
    expect(entryHover(view, e1)).toContain("spilled to `E1:E3` when saved (3×1)");
    expect(entryHover(view, entryForCell(view, "B3")!)).toContain("text stored in B2");
    // Spl = 'S1'!$E$1# leads E1's line: it resolves in the project, and labels the outline.
    const spl = view.text.indexOf("Spl  E1#");
    const left = lhsAt(view, spl + 1)!;
    expect(left.entry).toBe(e1);
    expect(left.line.lhs[left.k]).toMatchObject({ key: "Spl", display: "Spl", target: "spill" });
    expect(p.lookup(left.line.lhs[left.k]!.key)!.file.path).toBe(S1);
    expect(lhsAt(view, e1.address.start)).toBeUndefined();
    expect(entryLabel(view.lines[e1.index]!)).toBe("Spl — E1#");
    expect(entryLabel(view.lines[entryForCell(view, "B3")!.index]!)).toBe("B3");
    // Without a project the names resolve against the workbook's own, and say so.
    const loose = buildFormulaView(wb, "S2", "f7_base.xlsx", undefined)!;
    expect(loose.linked).toBe(false);
    expect(loose.text).toContain("Names are not linked");
    expect(buildFormulaView(wb, "Nope", "f7_base.xlsx", p)).toBeUndefined();
  });

  it("formula view in calculation order: levels, what a line reads and what reads it; the workbook view", () => {
    const wb = readWorkbook(bytes);
    const view = buildFormulaView(wb, "S1", "f7_base.xlsx", p, undefined, "calculation")!;
    expect(view.order).toBe("calculation");
    expect(view.text).toContain("// Sheet S1: 22 formulas in calculation order");
    const c6 = entryForCell(view, "C6")!;
    expect(view.lines[c6.index]).toMatchObject({ level: 2, dependsOn: ["Spl"] });
    expect(view.lines.findIndex((l) => l.cell === "E1")).toBeLessThan(c6.index);
    const links = entryLinks(view, c6)!;
    // Through the name to the block behind it, so the hover can lead to its line.
    expect(links.precedents).toEqual([
      { label: "Spl", kind: "name" },
      { label: "E1#", kind: "formula", sheet: "S1", cell: "E1", via: "Spl" },
    ]);
    expect(links.dependents).toEqual([]);
    const e1 = entryLinks(view, entryForCell(view, "E1")!)!;
    expect(e1.precedents.map((n) => n.label)).toEqual(["Rate"]);
    expect(e1.dependents.map((n) => [n.label, n.sheet, n.cell, n.via])).toEqual([
      ["Spl", undefined, undefined, undefined],
      ["C6", "S1", "C6", "Spl"],
      ["C9", "S1", "C9", undefined],
    ]);
    expect(entryHover(view, c6)).toContain("level 2");
    // The appearance view takes a graph for its hover too.
    const plain = buildFormulaView(wb, "S1", "f7_base.xlsx", p)!;
    expect(plain.graph).toBeUndefined();
    expect(entryLinks(plain, entryForCell(plain, "C6")!, view.graph)!.precedents).toHaveLength(2);
    // Every sheet: addresses carry their sheet; cells are found per sheet.
    const all = buildFormulaView(wb, WORKBOOK_VIEW, "f7_base.xlsx", p)!;
    expect(all.sheet).toBe(WORKBOOK_VIEW);
    expect(all.text).toContain("// Workbook f7_base.xlsx: 28 formulas on 2 sheets in calculation order");
    const s2a1 = entryForCell(all, "A1", "S2")!;
    expect(all.lines[s2a1.index]!.sheet).toBe("S2");
    expect(all.text.slice(s2a1.address.start - 5, s2a1.address.end)).toBe("'S2'!A1");
    expect(entryLabel(all.lines[s2a1.index]!, true)).toBe("S2!A1");
  });
});

describe("resolution", () => {
  const files = {
    [MANIFEST_FILE]: JSON.stringify({ format: "xln.manifest/1", workbook: "w.xlsx", sheets: [{ name: "BS", position: 0 }, { name: "P L", position: 1 }], tables: [], names: {} }),
    "names/_unmanaged.xln": [
      "Rate = 0.1;",
      "Loc = 1;",
      "/** Shadows Rate inside. */",
      "Shadow = LET(Rate, 2, x, Rate + Loc, LAMBDA(Loc, Loc * x + Rate)(3));",
      "Uses = 'P L'!Loc + BS!Loc + Loc + SUM(Rate);",
      "FN.Twice = LAMBDA(v, v * 2);",
      "Call = FN.Twice(Rate) + Fact(1);",
      "Fact = 5;",
      "Commented = Rate /* a note */ + Loc // tail",
      "  + Rate;",
    ].join("\n"),
    "names/sheets/BS.xln": ["@scope(BS)", "Loc = 2;", "Here = Loc + Rate;"].join("\n"),
    "names/sheets/P L.xln": ["@scope('P L')", "Loc = 3;", "@workbook", "Late = Loc;", "@scope(Nope)", "Bad = (1;"].join("\n"),
  };
  const p = projectFrom(files);
  const U = "names/_unmanaged.xln";
  const BS = "names/sheets/BS.xln";
  const PL = "names/sheets/P L.xln";
  const target = (path: string, needle: string, n = 0) => {
    const d = p.definition(path, at(p, path, needle, n));
    return d ? `${d.path}@${p.files.get(d.path)!.lines.position(d.start).line + 1}:${p.files.get(d.path)!.text.slice(d.start, d.end)}` : "<none>";
  };

  it("respects LET and LAMBDA shadowing", () => {
    // `Rate + Loc` inside LET: Rate is the LET variable, Loc the workbook name.
    expect(target(U, "Rate + Loc")).toBe(`${U}@4:Rate`);
    expect(p.definition(U, at(p, U, "Rate + Loc"))!.start).toBe(at(p, U, "LET(Rate", 0, 4));
    expect(target(U, "Loc,", 0)).toBe(`${U}@2:Loc`);
    // Inside the LAMBDA, Loc is its parameter; Rate is still the LET variable.
    expect(p.definition(U, at(p, U, "Loc * x"))!.start).toBe(at(p, U, "LAMBDA(Loc", 0, 7));
    expect(p.definition(U, at(p, U, "Rate)(3)"))!.start).toBe(at(p, U, "LET(Rate", 0, 4));
    // So Shadow does not use the workbook's Rate at all.
    expect(p.uses(p.lookup("Shadow")!).map((u) => u.def.key)).toEqual(["Loc"]);
  });

  it("respects sheet scope", () => {
    expect(p.definition(U, at(p, U, "'P L'!Loc", 0, 7))!.path).toBe(PL);
    expect(p.definition(U, at(p, U, "'P L'!Loc", 0, 2))).toBeUndefined(); // on the qualifier
    expect(p.definition(U, at(p, U, "BS!Loc", 0, 4))!.path).toBe(BS);
    expect(p.definition(U, at(p, U, "+ Loc + SUM", 0, 3))!.path).toBe(U);
    // In a BS-scoped definition a bare Loc is BS's own.
    expect(p.definition(BS, at(p, BS, "Loc + Rate"))!.path).toBe(BS);
    expect(p.definition(BS, at(p, BS, "Rate;"))!.path).toBe(U);
    // After @workbook, back to workbook scope.
    expect(p.definition(PL, at(p, PL, "= Loc", 0, 3))!.path).toBe(U);
    expect(p.references("BS!Loc").map((l) => l.path)).toEqual([U, BS]);
    expect(p.references("P L!Loc").map((l) => l.path)).toEqual([U]);
  });

  it("calls a LAMBDA name but never shadows a built-in", () => {
    expect(target(U, "FN.Twice(Rate)")).toBe(`${U}@6:FN.Twice`);
    expect(p.definition(U, at(p, U, "Fact(1)"))).toBeUndefined(); // FACT, the built-in (T12)
    expect(p.definition(U, at(p, U, "SUM("))).toBeUndefined();
  });

  it("maps positions around comments inside a definition", () => {
    expect(target(U, "Rate /*")).toBe(`${U}@1:Rate`);
    expect(target(U, "Loc // tail")).toBe(`${U}@2:Loc`);
    expect(target(U, "+ Rate;", 0).length).toBeGreaterThan(0);
    expect(p.definition(U, at(p, U, "  + Rate;", 0, 5))!.path).toBe(U);
    expect(p.definition(U, at(p, U, "a note"))).toBeUndefined();
  });

  it("lists every read of a local variable", () => {
    const locs = p.localReferences(U, at(p, U, "x + Rate"))!;
    expect(locs.map((l) => p.files.get(l.path)!.text.slice(l.start, l.end))).toEqual(["x", "x"]);
  });

  it("reports module and formula errors, duplicates and unknown sheets", () => {
    const probs = p.problems(PL).map((q) => [q.severity, q.message, p.files.get(PL)!.text.slice(q.start, q.end)]);
    // An old sheet file (blocks): an info, here without the conversion (a block of another sheet).
    expect(probs).toEqual([
      [
        "info",
        "this sheet file has @scope/@workbook blocks (written before M3d). Every name in a sheet file is local to P L unless @workbook is on the line above it; it cannot be converted as it is (line 5: @scope(Nope) holds names local to another sheet; move them to that sheet's file first)",
        "@scope('P L')",
      ],
      ["error", "@scope(Nope): the workbook has no sheet 'Nope'", "@scope(Nope)"],
      ["error", "Bad: expected ')' to close the '(' at column 1, found the end of the formula", ";"],
    ]);
    const m = new Project("mem:/m", undefined);
    m.setFile("names/m.xln", "m.A = 1;\nB 2;\nm.C = 3");
    expect(m.problems("names/m.xln").map((q) => [q.message, q.start])).toEqual([
      ["'B': expected '='", 11],
      ["'m.C': definition is not ended by ';'", 21],
    ]);
    const dup = new Project("mem:/d", undefined);
    dup.setFile("names/a.xln", "A = 1;\n");
    dup.setFile("names/b.xln", "a = 2;\n");
    expect(dup.problems("names/b.xln").map((q) => q.message)).toEqual(["a is also defined at names/a.xln:1: keep one (a name can be defined once per scope)"]);
  });

  it("follows edits: re-resolves other files when one changes", () => {
    const q = projectFrom(files);
    expect(q.definition(BS, at(q, BS, "Loc + Rate"))!.path).toBe(BS);
    q.setFile(BS, ["@scope(BS)", "Here = Loc + Rate;"].join("\n"));
    expect(q.definition(BS, at(q, BS, "Loc + Rate"))!.path).toBe(U);
    expect(q.usedBy("Loc").map((u) => u.def.key)).toContain("BS!Here");
  });

  it("folds multi-line definitions, doc comments and blocks", () => {
    const f = new Project("mem:/f", undefined);
    f.setFile("names/a.xln", ["@scope(S)", "", "/**", " * doc", " */", "A = LET(", "  x, 1,", "  x);", "B = 1;"].join("\n"));
    expect(f.folding("names/a.xln")).toEqual([
      { start: 0, end: 8 },
      { start: 5, end: 7 },
      { start: 2, end: 4, kind: "comment" },
    ]);
  });
});

describe("helpers", () => {
  it("names the project folder and recognises workbooks and projects", () => {
    expect(projectFolderName("lbo-ep03r.xlsx")).toBe("lbo-ep03r.xln");
    expect(projectFolderName("a.b.xlsm")).toBe("a.b.xln");
    expect(isWorkbookName("A.XLSX")).toBe(true);
    expect(isWorkbookName("~$A.xlsx")).toBe(false);
    expect(isWorkbookName("a.xls")).toBe(false);
    const dir = (name: string) => ({ name, isDirectory: true });
    const file = (name: string) => ({ name, isDirectory: false });
    expect(isProjectListing([file("xln.lock.json"), file("workbook.manifest.json"), dir("names")])).toBe(true);
    expect(isProjectListing([file("xln.lock.json"), dir("names")])).toBe(false);
  });

  it("summarises a pull", () => {
    const r = pullProject(new Uint8Array(readFileSync(f7)), "f7_base.xlsx");
    const lines = formatPullSummary(r.report, "f7_base.xln", Object.keys(r.files), []);
    expect(lines[1]).toBe("  7 names: 6 workbook-scoped, 1 sheet-scoped (S2 1)");
  });

  it("before a pull that would replace source edits not built: the question lists them; after Discard, the summary and the notification say so", () => {
    const bytes = new Uint8Array(readFileSync(f7));
    const files = { ...pullProject(bytes, "f7_base.xlsx").files };
    files["names/_unmanaged.xln"] = files["names/_unmanaged.xln"]!.replace("Rate2 = Rate*2;", "Rate2 = Rate*3;");
    const edits = unbuiltEdits({ workbook: bytes, fileName: "f7_base.xlsx", files });
    const q = unbuiltQuestion("f7_base.xlsx", edits);
    expect(q.message).toBe("f7_base.xln has 1 source edit(s) not built yet. A pull writes the project as f7_base.xlsx has it now and would replace them.");
    expect(q.detail).toBe("names/_unmanaged.xln:8  Rate2: update Rate2 (definition)");
    const r = pullProject(bytes, "f7_base.xlsx");
    const lines = formatPullSummary(r.report, "f7_base.xln", ["names/_unmanaged.xln"], [], edits);
    expect(lines).toContain("  replaced 1 source edit(s) not built (Discard and pull):");
    expect(lines).toContain("    names/_unmanaged.xln:8  Rate2: update Rate2 (definition)");
    expect(pullNotification("f7_base.xlsx", 7, 1)).toBe("xln: pulled 7 names from f7_base.xlsx into f7_base.xln. 1 source edit(s) not built were replaced.");
    expect(pullNotification("f7_base.xlsx", 7)).toBe("xln: pulled 7 names from f7_base.xlsx into f7_base.xln.");
    // A workbook with AFE's module store (synthetic fixture): said, and left alone.
    const afe = pullProject(new Uint8Array(readFileSync(fileURLToPath(new URL("../../../probes/fixtures/afe-synthetic-v11.xlsx", import.meta.url)))), "afe.xlsx").report;
    expect(pullNotification("afe.xlsx", afe.names, 0, afe.foreignModules)).toBe(
      "xln: pulled 13 names from afe.xlsx into afe.xln. It also carries Advanced Formula Environment (AFE) modules (Workbook, ANA), left as they are.",
    );
    // Names Create from Selection took from a computed value or a corner (2026-10-07): one sentence, the details in the output.
    const one = [{ key: "italian10", kind: "value" as const, cell: "Mortgage!A12", range: "Mortgage!B12:G12" }];
    expect(pullNotification("is-model.xlsx", 40, 0, [], one)).toBe(
      "xln: pulled 40 names from is-model.xlsx into is-model.xln. One name looks named after a cell's current value or a corner label (italian10): see the xln output.",
    );
    const two = [...one, { key: "Mortgage", kind: "corner" as const, cell: "Mortgage!A11", range: "Mortgage!B12:G12" }];
    expect(pullNotification("is-model.xlsx", 40, 0, [], two)).toContain(" 2 names look named after a cell's current value or a corner label (italian10, Mortgage): see the xln output.");
  });

  it("M3e: files differing in layout or comments alone: a group of their own in the question; alone, a question of their own", () => {
    const bytes = new Uint8Array(readFileSync(f7));
    const pulled = pullProject(bytes, "f7_base.xlsx").files;
    const files = { ...pulled };
    files["names/_unmanaged.xln"] = files["names/_unmanaged.xln"]!.replace("Rate2 = Rate*2;", "// twice\nRate2 = Rate*2;");
    files["names/sheets/S2.xln"] = files["names/sheets/S2.xln"]!.replace("Loc = 7;", "Loc = 8;");
    const edits = unbuiltEdits({ workbook: bytes, fileName: "f7_base.xlsx", files });
    const rewritten = rewrittenFiles(files, pulled, edits);
    expect(rewritten).toEqual(["names/_unmanaged.xln"]);
    const q = unbuiltQuestion("f7_base.xlsx", edits, rewritten);
    expect(q.detail.split("\n").slice(-3)).toEqual(["", "Layout or comments only: will be rewritten", "names/_unmanaged.xln"]);
    const alone = rewrittenQuestion("f7_base.xlsx", rewritten);
    expect(alone.message).toMatch(/^f7_base\.xln has no edits that are not built, but 1 file\(s\) differ in layout or comments/);
    expect(alone.detail).toBe("Layout or comments only: will be rewritten\nnames/_unmanaged.xln");
  });

  it("shows cached values as Excel does", () => {
    expect(formatValue({ type: "n", raw: "0.7000000000000001", value: 0.7000000000000001 })).toBe("0.7");
    expect(formatValue({ type: "b", raw: "1", value: true })).toBe("TRUE");
    expect(formatValue({ type: "s", raw: "0", value: "a" })).toBe('"a"');
    expect(formatValue({ type: "e", raw: "#N/A", value: "#N/A" })).toBe("#N/A");
    expect(formatValue({ type: "n", raw: undefined, value: undefined })).toBeUndefined();
  });

  it("converts offsets and positions", () => {
    const l = new LineIndex("ab\r\ncd\ne");
    expect(l.position(4)).toEqual({ line: 1, character: 0 });
    expect(l.offset({ line: 2, character: 1 })).toBe(8);
    expect(l.lineText(0)).toBe("ab");
    expect(l.lineCount).toBe(3);
  });
});
