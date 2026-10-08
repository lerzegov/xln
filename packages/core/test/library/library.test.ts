// M4: the `.lambda` reader, lib status on constructed cases (a real build writes the D6
// tags), and the insert, update and publish edits, round trip included.
import { describe, expect, it } from "vitest";
import {
  applyEdits,
  definitionBase,
  buildWorkbook,
  checkProject,
  libraryInsertion,
  libraryReplacement,
  libraryStatus,
  findEntry,
  parseLambdaFile,
  projectCopies,
  projectLibraryStatus,
  publishLambda,
  readLibrary,
  readWorkbook,
  renderLibStatus,
  SourceModel,
  workbookCopies,
  libraryDoc,
  commentLength,
  firstSentence,
  sameSummary,
  libStatusJson,
  parseConfig,
  type Library,
  type LibStatusReport,
} from "../../src/index.js";
import { fixture, pulled, withWorkbookXml, workbookXml } from "../build/helpers.js";

const GROW = [
  "# name       FN.GROW",
  "# summary    Growth chain: seed the first period,",
  "#            compound it thereafter.",
  "# params     seed, growth, periods",
  "# impromptu  x[year:first] = <seed>",
  "#            x[year:rest]  = x[year:prev] * (1 + <g>)",
  "# example    FN.GROW(10000, 0.05, 5)",
  "#            ->  {10000, 10500}",
  "#",
  "# Why it is written this way.",
  "#",
  "# @param seed the first period's value",
  "",
  "LAMBDA(seed, growth, periods,",
  "    seed * (1 + growth) ^ SEQUENCE(1, periods, 0)",
  ")",
  "",
].join("\n");

describe("parseLambdaFile", () => {
  it("reads every field, continuation lines, the rationale and @param descriptions", () => {
    const r = parseLambdaFile("lib/FN.GROW.lambda", GROW);
    expect(r.problems).toEqual([]);
    const fn = r.fn!;
    expect(fn.name).toBe("FN.GROW");
    expect(fn.summary).toBe("Growth chain: seed the first period, compound it thereafter.");
    expect(fn.params).toEqual(["seed", "growth", "periods"]);
    expect(fn.fields.map((f) => [f.key, f.line, f.endLine])).toEqual([
      ["name", 1, 1],
      ["summary", 2, 3],
      ["params", 4, 4],
      ["impromptu", 5, 6],
      ["example", 7, 8],
    ]);
    expect(fn.fields[3]!.value).toBe("x[year:first] = <seed>\nx[year:rest]  = x[year:prev] * (1 + <g>)");
    expect(fn.rationale).toBe("Why it is written this way.\n\n@param seed the first period's value");
    expect(fn.paramDocs).toEqual({ seed: "the first period's value" });
    expect(fn.definition).toBe("LAMBDA(seed, growth, periods,\n    seed * (1 + growth) ^ SEQUENCE(1, periods, 0)\n)");
    expect(fn.definitionLine).toBe(14);
    expect(fn.doc).toBe("Growth chain: seed the first period, compound it thereafter.\n@param seed the first period's value\n@param growth\n@param periods");
  });

  it("reads CR LF files the same", () => {
    const r = parseLambdaFile("FN.GROW.lambda", GROW.split("\n").join("\r\n"));
    expect(r.problems).toEqual([]);
    expect(r.fn!.definition).toBe(parseLambdaFile("FN.GROW.lambda", GROW).fn!.definition);
  });

  it("reports problems with line numbers", () => {
    const noName = parseLambdaFile("X.lambda", "# summary  s\n# params x\n\nLAMBDA(x, x)\n");
    expect(noName.fn).toBeUndefined();
    expect(noName.problems).toContainEqual({ severity: "error", path: "X.lambda", line: 1, message: "the header has no name (# name FN.X)" });

    const params = parseLambdaFile("FN.P.lambda", "# name FN.P\n# summary s\n# params a, b\n\nLAMBDA(a, c, a + c)\n");
    expect(params.problems).toEqual([{ severity: "error", path: "FN.P.lambda", line: 3, message: "the header declares params a, b but the LAMBDA takes a, c" }]);

    const bad = parseLambdaFile("FN.Q.lambda", "# name FN.Q\n# summary s\n# params a\n\nLAMBDA(a,\n  a +* )\n");
    expect(bad.fn).toBeUndefined();
    expect(bad.problems[0]!.line).toBe(6);
    expect(bad.problems[0]!.message).toMatch(/^the definition does not parse/);

    const notLambda = parseLambdaFile("FN.R.lambda", "# name FN.R\n# summary s\n\nSUM(1, 2)\n");
    expect(notLambda.problems).toContainEqual({ severity: "error", path: "FN.R.lambda", line: 4, message: "the definition must be a single LAMBDA" });

    const renamed = parseLambdaFile("lib/FN.OLD.lambda", "# name FN.NEW\n# summary s\n# params a\n\nLAMBDA(a, a)\n");
    expect(renamed.fn!.name).toBe("FN.NEW");
    expect(renamed.problems).toEqual([{ severity: "warning", path: "lib/FN.OLD.lambda", line: 1, message: "the file is named FN.OLD.lambda but the header names FN.NEW" }]);

    // An unknown function would be poisoned as _xludf. by Excel: the compiler says so.
    const unknown = parseLambdaFile("FN.U.lambda", "# name FN.U\n# summary s\n# params a\n\nLAMBDA(a,\n  NOSUCHFN(a))\n");
    expect(unknown.problems.some((p) => p.severity === "error" && p.line === 6 && /does not compile/.test(p.message))).toBe(true);
  });

  it("shortens a long doc comment, never the definition", () => {
    const long = "word ".repeat(80).trim() + ".";
    const d = libraryDoc(long, ["a", "b"]);
    expect(d.shortened).toBe(true);
    expect(commentLength(d.doc)).toBeLessThanOrEqual(223);
    expect(d.doc.endsWith("…\n@param a\n@param b")).toBe(true);
    expect(libraryDoc("Short.", ["a"])).toEqual({ doc: "Short.\n@param a", shortened: false });
  });
});

// ---- lib status ---------------------------------------------------------------------------

function lambdaFile(name: string, summary: string, def: string): string {
  return `# name       ${name}\n# summary    ${summary}\n# params     x\n#\n# Rationale of ${name}.\n\n${def}\n`;
}

/** The library of the constructed cases. */
function library(over: Record<string, string> = {}): Library {
  const files: Record<string, string> = {
    "FN.SAME.lambda": lambdaFile("FN.SAME", "Doubles.", "LAMBDA(x, x * 2)"),
    "FN.OUT.lambda": lambdaFile("FN.OUT", "Adds two.", "LAMBDA(x, x + 2)"),
    "FN.MOD.lambda": lambdaFile("FN.MOD", "Adds one to a double.", "LAMBDA(x, FN.SAME(x) + 1)"),
    "FN.DIFF.lambda": lambdaFile("FN.DIFF", "Halves.", "LAMBDA(x, x / 2)"),
    "FN.NEW.lambda": lambdaFile("FN.NEW", "Quadruples, through FN.TWICE.", "LAMBDA(x, FN.TWICE(FN.TWICE(x)))"),
    "FN.TWICE.lambda": lambdaFile("FN.TWICE", "Twice, the library's own.", "LAMBDA(x, 2 * x)"),
    "FN.BOTH.lambda": lambdaFile("FN.BOTH", "Adds thirty.", "LAMBDA(x, x + 30)"),
    ...over,
  };
  return readLibrary(files);
}

/** The library base of a definition (what `@from(lib #…)` records). */
const base = (def: string, name = "FN.X") => definitionBase(def, name);

// The library bases: FN.OUT came from a library version equal to the copy (the library
// moved since: outdated), FN.MOD from the library's current version (an edit here will
// make it modified), FN.BOTH from a version both sides left (both changed).
const FN_SOURCE = [
  "// module: FN",
  "",
  "/** Doubles. */",
  "FN.SAME = LAMBDA(x, x*2);",
  "",
  "/** Adds one. */",
  `@from(lib #${base("LAMBDA(x, x + 1)")})`,
  "FN.OUT = LAMBDA(x, x + 1);",
  "",
  `@from(lib #${base("LAMBDA(x, FN.SAME(x) + 1)")})`,
  "FN.MOD = LAMBDA(x, FN.SAME(x) + 1);",
  "",
  `@from(lib #${base("LAMBDA(x, x + 10)")})`,
  "FN.BOTH = LAMBDA(x, x + 20);",
  "",
  "FN.LOCAL = LAMBDA(x, x - 1);",
  "",
].join("\n");

/** probe_mac with names/FN.xln built into it (tags on), and the project after the build. */
function builtProject(provenance = true, source = FN_SOURCE): { bytes: Uint8Array; files: Record<string, string> } {
  const base = fixture("probe_mac.xlsx");
  const files = pulled(base);
  files["names/FN.xln"] = source;
  const r = buildWorkbook({ workbook: base, fileName: "book.xlsx", files }, { embed: false, provenance });
  expect(r.status, JSON.stringify(r.plan.problems)).toBe("built");
  // The lockfile and manifest after the build, as the CLI writes them.
  return { bytes: r.bytes!, files: { ...files, ...r.files } };
}

/** The states of the FN names (probe_mac has module LAMBDAs of its own, local only). */
function states(r: LibStatusReport): Record<string, string> {
  return Object.fromEntries(r.items.filter((i) => i.name.startsWith("FN.")).map((i) => [i.name, i.state]));
}

describe("lib status", () => {
  it("identical, outdated, modified, both changed, differs, missing and local only, in a workbook", () => {
    const { bytes } = builtProject();
    // FN.MOD edited in Excel after the build (its tag now says otherwise), FN.DIFF added in Excel (no tag).
    let xml = workbookXml(bytes);
    const at = xml.indexOf('name="FN.MOD"');
    const close = xml.indexOf("</definedName>", at);
    const head = xml.slice(0, close);
    const plus = head.lastIndexOf("+ 1");
    expect(plus).toBeGreaterThan(at);
    xml = xml.slice(0, plus) + "+ 5" + xml.slice(plus + 3);
    const end = xml.indexOf("</definedNames>");
    xml = xml.slice(0, end) + '<definedName name="FN.DIFF">_xlfn.LAMBDA(_xlpm.x,_xlpm.x/3)</definedName>' + xml.slice(end);
    const wb = readWorkbook(withWorkbookXml(bytes, xml));
    const r = libraryStatus(library(), workbookCopies(wb), { target: "book.xlsx", kind: "workbook", library: "lib" });
    expect(states(r)).toEqual({
      "FN.DIFF": "differs",
      "FN.MOD": "modified",
      "FN.NEW": "missing",
      "FN.OUT": "outdated",
      "FN.SAME": "identical",
      "FN.TWICE": "missing",
      "FN.BOTH": "both-changed",
      "FN.LOCAL": "local-only",
    });
    const mod = r.items.find((i) => i.name === "FN.MOD")!;
    // The tag records what was built and, after it, the library base the source's @from gave.
    expect(mod.tag).toBe(`[xln FN #${mod.tag!.slice(9, 15)} lib#${base("LAMBDA(x, FN.SAME(x) + 1)")}]`);
    expect(mod.base).toBe(mod.libraryHash);
    expect(mod.note).toContain("edited in Excel since the build that tagged it");
    expect(mod.diff!.filter((l) => l.op !== " ").map((l) => l.op + l.text)).toEqual(["-LAMBDA(x, FN.SAME(x) + 5)", "+LAMBDA(x, FN.SAME(x) + 1)"]);
    const text = renderLibStatus(r).text;
    expect(text).toContain("outdated 1 · modified 1 · both changed 1 · differs 1 · missing 2 · identical 1 · local only 1 (7 library functions)");
    // Both changed, the base's text not at hand in a workbook alone: the copy against the library.
    expect(text).toContain(`the base #${base("LAMBDA(x, x + 10)")}'s text is not at hand`);
    expect(text).toMatch(new RegExp(` {2}FN\\.OUT +\\[xln FN #[0-9a-f]{6} lib#${base("LAMBDA(x, x + 1)")}\\] {2}base #${base("LAMBDA(x, x + 1)")} library #${base("LAMBDA(x, x + 2)")}\\n`));
    expect(text).toContain("other module LAMBDAs, not library candidates: 4 (Mod 1, P 3)");
    expect(r.otherModules).toEqual({ Mod: 1, P: 3 });
    expect(text).toContain("      - LAMBDA(x, FN.SAME(x) + 5)\n      + LAMBDA(x, FN.SAME(x) + 1)");
    const json = libStatusJson(r) as { items: { name: string; state: string; diff?: string[] }[] };
    expect(json.items.find((i) => i.name === "FN.OUT")!.diff).toEqual(["- LAMBDA(x, x + 1)", "+ LAMBDA(x, x + 2)"]);
  });

  it("without tags nothing is outdated or modified: it differs", () => {
    const { bytes } = builtProject(false);
    const r = libraryStatus(library(), workbookCopies(readWorkbook(bytes)), { target: "book.xlsx", kind: "workbook", library: "lib" });
    expect(states(r)["FN.OUT"]).toBe("differs");
    expect(states(r)["FN.MOD"]).toBe("identical");
  });

  it("a project: its source against the library, the base from its @from", () => {
    const { bytes, files } = builtProject();
    const wb = readWorkbook(bytes);
    // An edit in the source not built yet: the source is no longer its base.
    const edited = { ...files, "names/FN.xln": files["names/FN.xln"]!.replace("FN.MOD = LAMBDA(x, FN.SAME(x) + 1)", "FN.MOD = LAMBDA(x, FN.SAME(x) + 7)") };
    const r = libraryStatus(library(), projectCopies(edited, wb), { target: "book.xln", kind: "project", library: "lib" });
    expect(states(r)).toMatchObject({ "FN.SAME": "identical", "FN.OUT": "outdated", "FN.MOD": "modified", "FN.BOTH": "both-changed", "FN.NEW": "missing", "FN.LOCAL": "local-only" });
    expect(r.items.find((i) => i.name === "FN.MOD")).toMatchObject({ file: "names/FN.xln", line: 11 });
    // The base is the source's: the workbook is not needed to tell the states apart.
    expect(states(libraryStatus(library(), projectCopies(edited), { target: "p", kind: "project", library: "lib" }))).toMatchObject({ "FN.OUT": "outdated", "FN.MOD": "modified" });
    // Without @from: differs.
    const noBase = files["names/FN.xln"]!.split("\n").filter((l) => !l.startsWith("@from")).join("\n");
    expect(states(libraryStatus(library(), projectCopies({ ...files, "names/FN.xln": noBase }), { target: "p", kind: "project", library: "lib" }))["FN.OUT"]).toBe("differs");
  });

  it("a project with its workbook says what of the source is not built yet (feedback 2026-10-07)", () => {
    const { bytes, files } = builtProject();
    const wb = readWorkbook(bytes);
    // Inserted from the library (FN.TWICE), edited (FN.MOD), deleted (FN.OUT): none built.
    const out = files["names/FN.xln"]!.indexOf("/** Adds one. */");
    const outEnd = files["names/FN.xln"]!.indexOf("FN.OUT = LAMBDA(x, x + 1);\n") + "FN.OUT = LAMBDA(x, x + 1);\n".length;
    const fn = (files["names/FN.xln"]!.slice(0, out) + files["names/FN.xln"]!.slice(outEnd)).replace("FN.MOD = LAMBDA(x, FN.SAME(x) + 1)", "FN.MOD = LAMBDA(x, FN.SAME(x) + 7)") + "\nFN.TWICE = LAMBDA(x, 2 * x);\n";
    const r = projectLibraryStatus(library(), { ...files, "names/FN.xln": fn }, wb, { target: "book.xln", library: "lib", workbook: "book.xlsx" });
    const unbuilt = Object.fromEntries(r.items.filter((i) => i.name.startsWith("FN.")).map((i) => [i.name, i.unbuilt ?? "-"]));
    expect(unbuilt).toEqual({ "FN.SAME": "-", "FN.OUT": "deleted", "FN.MOD": "edited", "FN.DIFF": "-", "FN.NEW": "-", "FN.TWICE": "new", "FN.BOTH": "-", "FN.LOCAL": "-" });
    expect(states(r)).toMatchObject({ "FN.TWICE": "identical", "FN.OUT": "missing", "FN.MOD": "modified" });
    expect(r.unbuilt).toBe(3);
    const text = renderLibStatus(r).text;
    expect(text.split("\n")[1]).toBe("outdated 0 · modified 1 · both changed 1 · differs 0 · missing 3 · identical 2 · local only 1 (7 library functions) · not built yet 3");
    expect(text.split("\n")[2]).toBe("The source differs from book.xlsx: 3 functions are not built yet (build to write them into the workbook).");
    expect(text).toMatch(/ {2}FN\.TWICE +library #[0-9a-f]{6} {2}names\/FN\.xln:\d+ {2}\(in the source, not built yet\)\n/);
    expect(text).toMatch(/ {2}FN\.OUT +library #[0-9a-f]{6} {2}\(deleted in the source, still in the workbook: not built yet\)\n/);
    expect(r.items.find((i) => i.name === "FN.TWICE")!.note).toBe("same definition as the library; no base recorded: Record library base; in the source, not built yet");
    expect((libStatusJson(r) as { unbuilt: number }).unbuilt).toBe(3);
    // Built as the source says: nothing left.
    const again = projectLibraryStatus(library(), files, wb, { target: "book.xln", library: "lib", workbook: "book.xlsx" });
    expect(again.unbuilt).toBe(0);
    expect(renderLibStatus(again).text).not.toContain("not built yet");
    // Without the workbook nothing can be told, and the report says so.
    const blind = projectLibraryStatus(library(), files, undefined, { target: "book.xln", library: "lib" });
    expect(blind.workbook).toBeNull();
    expect(renderLibStatus(blind).text).toContain("The workbook was not found beside the project: what is built cannot be told.");
  });

  it("numbers by value and layout do not count", () => {
    const lib = readLibrary({ "FN.E.lambda": lambdaFile("FN.E", "Tiny.", "LAMBDA(x,\n    x + 1E-14\n)") });
    const { bytes } = builtProject(true, "FN.E = LAMBDA(x, x + 0.00000000000001);\n");
    expect(states(libraryStatus(lib, workbookCopies(readWorkbook(bytes)), { target: "b", kind: "workbook", library: "lib" }))["FN.E"]).toBe("identical");
  });
});

// ---- insert, update, publish --------------------------------------------------------------

describe("insert from library", () => {
  it("adds the function and the library functions it calls, in name order, to names/FN.xln", () => {
    const files = { "names/FN.xln": FN_SOURCE };
    const ins = libraryInsertion(library(), files, "fn.new")!;
    expect(ins.names).toEqual(["FN.TWICE", "FN.NEW"]);
    expect(ins.edits.length).toBe(1);
    const text = applyEdits(FN_SOURCE, ins.edits[0]!.edits);
    // At its sort position among the file's names as they are (here not sorted): before the first that sorts after it.
    const order = [...text.matchAll(/^(FN\.[A-Z]+) =/gm)].map((m) => m[1]);
    expect(order).toEqual(["FN.NEW", "FN.SAME", "FN.OUT", "FN.MOD", "FN.BOTH", "FN.LOCAL", "FN.TWICE"]);
    expect(text).toContain(`/**\n * Quadruples, through FN.TWICE.\n * @param x\n */\n@from(lib #${base("LAMBDA(x, FN.TWICE(FN.TWICE(x)))")})\nFN.NEW = LAMBDA(x, FN.TWICE(FN.TWICE(x)));\n\n`);
    expect(text.endsWith(`@from(lib #${base("LAMBDA(x, 2 * x)")})\nFN.TWICE = LAMBDA(x, 2 * x);\n`)).toBe(true);
    const model = new SourceModel();
    model.setFile("names/FN.xln", text);
    expect(checkProject(model).filter((p) => p.severity === "error")).toEqual([]);
  });

  it("creates names/FN.xln when the project has no FN module", () => {
    const ins = libraryInsertion(library(), { "names/_unmanaged.xln": "Rate = 0.1;\n" }, "FN.MOD")!;
    expect(ins.names).toEqual(["FN.SAME", "FN.MOD"]);
    expect(ins.edits).toEqual([{ path: "names/FN.xln", create: true, edits: [{ start: 0, end: 0, text: expect.stringContaining("FN.MOD = LAMBDA(x, FN.SAME(x) + 1);\n\n/**\n * Doubles.") }] }]);
    expect(libraryInsertion(library(), {}, "FN.NOPE")).toBeUndefined();
  });

  it("inserted, built: identical", () => {
    const base = fixture("probe_mac.xlsx");
    const files = pulled(base);
    const ins = libraryInsertion(library(), files, "FN.NEW")!;
    for (const fe of ins.edits) files[fe.path] = applyEdits(fe.create ? "" : files[fe.path]!, fe.edits);
    const r = buildWorkbook({ workbook: base, fileName: "book.xlsx", files }, { embed: false });
    expect(r.status).toBe("built");
    const s = libraryStatus(library(), workbookCopies(readWorkbook(r.bytes!)), { target: "b", kind: "workbook", library: "lib" });
    expect(states(s)).toMatchObject({ "FN.NEW": "identical", "FN.TWICE": "identical" });
    // The doc comment went to the Name Manager, tagged.
    const d = readWorkbook(r.bytes!).definedNames.find((n) => n.name === "FN.NEW")!;
    expect(d.comment).toMatch(/^Quadruples, through FN\.TWICE\.\r?\n@param x \[xln FN #[0-9a-f]{6} lib#[0-9a-f]{6}\]$/);
  });
});

describe("update from library", () => {
  it("replaces the definition and the doc comment, keeps the rest", () => {
    const text = "@hidden\nFN.MOD = LAMBDA(x, FN.SAME(x) + 9); // trailing\n\n/** Old doc. */\nFN.OUT = LAMBDA(x,\n  x + 1\n);\n";
    const lib = library();
    let t = applyEdits(text, libraryReplacement(text, findEntry(text, "FN.OUT")!, lib.get("FN.OUT")!));
    expect(t).toBe(`@hidden\nFN.MOD = LAMBDA(x, FN.SAME(x) + 9); // trailing\n\n/**\n * Adds two.\n * @param x\n */\n@from(lib #${base("LAMBDA(x, x + 2)")})\nFN.OUT = LAMBDA(x, x + 2);\n`);
    t = applyEdits(t, libraryReplacement(t, findEntry(t, "FN.MOD")!, lib.get("FN.MOD")!));
    expect(t.startsWith(`/**\n * Adds one to a double.\n * @param x\n */\n@from(lib #${base("LAMBDA(x, FN.SAME(x) + 1)")})\n@hidden\nFN.MOD = LAMBDA(x, FN.SAME(x) + 1); // trailing\n`)).toBe(true);
  });

  it("outdated, taken, built: identical", () => {
    const { bytes, files } = builtProject();
    const lib = library();
    const text = files["names/FN.xln"]!;
    files["names/FN.xln"] = applyEdits(text, libraryReplacement(text, findEntry(text, "FN.OUT")!, lib.get("FN.OUT")!));
    const r = buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files }, { embed: false });
    expect(r.status).toBe("built");
    expect(states(libraryStatus(lib, workbookCopies(readWorkbook(r.bytes!)), { target: "b", kind: "workbook", library: "lib" }))["FN.OUT"]).toBe("identical");
  });
});

describe("publish to library", () => {
  it("a new file: generated header, the rest of the doc as rationale, the definition", () => {
    const p = publishLambda({ name: "FN.LOCAL", doc: "Subtracts one. Used for offsets.\n@param x the value", formula: "LAMBDA(x,\n    x - 1\n)" });
    expect(p.created).toBe(true);
    expect(p.path).toBe("FN.LOCAL.lambda");
    expect(p.text).toBe(
      ["# name       FN.LOCAL", "# summary    Subtracts one.", "# params     x", "#", "# Used for offsets.", "# @param x the value", "", "LAMBDA(x,", "    x - 1", ")", ""].join("\n"),
    );
    const back = parseLambdaFile(p.path, p.text);
    expect(back.problems).toEqual([]);
    expect(back.fn!.paramDocs).toEqual({ x: "the value" });
    expect(publishLambda({ name: "FN.K", doc: undefined, formula: "42" }).error).toMatch(/not a LAMBDA/);
  });

  it("an update keeps the header's other fields and the rationale", () => {
    const p = publishLambda(
      { name: "FN.GROW", doc: "Growth chain: seed the first period, compound it thereafter.\n@param seed the first period's value\n@param growth\n@param periods", formula: "LAMBDA(seed, growth, periods, seed * (1 + growth) ^ SEQUENCE(1, periods))" },
      { path: "lib/FN.GROW.lambda", text: GROW },
    );
    expect(p.changed).toEqual(["definition"]);
    expect(p.text).toBe(GROW.slice(0, GROW.indexOf("LAMBDA(")) + "LAMBDA(seed, growth, periods, seed * (1 + growth) ^ SEQUENCE(1, periods))\n");
    // Same definition modulo layout: nothing to write.
    const same = publishLambda({ name: "FN.GROW", doc: "Growth chain: seed the first period, compound it thereafter.", formula: "LAMBDA(seed,growth,periods, seed*(1+growth)^SEQUENCE(1,periods,0))" }, { path: "x", text: GROW });
    expect(same.changed).toEqual([]);
    expect(same.text).toBe(GROW);
  });

  it("an update replaces the summary and params when they changed", () => {
    const p = publishLambda({ name: "FN.GROW", doc: "A new summary.\n@param seed\n@param g the rate", formula: "LAMBDA(seed, g, seed * (1 + g))" }, { path: "x", text: GROW });
    expect(p.changed).toEqual(["definition", "summary", "params", "param docs"]);
    const back = parseLambdaFile("FN.GROW.lambda", p.text);
    expect(back.problems).toEqual([]);
    expect(back.fn!.summary).toBe("A new summary.");
    expect(back.fn!.params).toEqual(["seed", "g"]);
    expect(back.fn!.paramDocs).toEqual({ g: "the rate" });
    expect(back.fn!.fields.map((f) => f.key)).toEqual(["name", "summary", "params", "impromptu", "example"]);
    expect(back.fn!.rationale).toBe("Why it is written this way.\n\n@param g the rate");
  });

  it("a shortened summary is the library's: not a change", () => {
    expect(sameSummary("Iterate a whole-row model function to its fixed…", "Iterate a whole-row model function to its fixed point.")).toBe(true);
    expect(sameSummary("Other.", "Iterate.")).toBe(false);
    expect(firstSentence("One. Two three.  Four")).toEqual({ first: "One.", rest: "Two three. Four" });
  });

  it("published, then status: identical (round trip)", () => {
    const { bytes, files } = builtProject();
    const entry = findEntry(files["names/FN.xln"]!, "FN.LOCAL")!;
    const p = publishLambda({ name: entry.name, doc: entry.doc, formula: entry.formula });
    const lib = readLibrary({ ...Object.fromEntries(library().functions.map((f) => [f.path, ""])), [p.path]: p.text });
    const r = libraryStatus(lib, workbookCopies(readWorkbook(bytes)), { target: "b", kind: "workbook", library: "lib" });
    expect(states(r)["FN.LOCAL"]).toBe("identical");
    // An outdated one published back from the project: the library takes the project's.
    const lib2 = library();
    const out = findEntry(files["names/FN.xln"]!, "FN.OUT")!;
    const up = publishLambda({ name: out.name, doc: out.doc, formula: out.formula }, { path: lib2.get("FN.OUT")!.path, text: lambdaFile("FN.OUT", "Adds two.", "LAMBDA(x, x + 2)") });
    expect(up.changed).toEqual(["definition", "summary"]);
    const lib3 = library({ "FN.OUT.lambda": up.text });
    expect(states(libraryStatus(lib3, workbookCopies(readWorkbook(bytes)), { target: "b", kind: "workbook", library: "lib" }))["FN.OUT"]).toBe("identical");
  });
});

describe("xln.config.json library", () => {
  it("is a path, kept as written", () => {
    expect(parseConfig('{ "library": " ../../_shared/lib " }').config.library).toBe("../../_shared/lib");
    expect(parseConfig('{ "library": "~/lib" }').config.library).toBe("~/lib");
    expect(parseConfig('{ "library": 3 }').problems).toEqual(["library: must be the path of the library folder (a string)"]);
  });
});

describe("publish keeps an optional parameter's brackets in its @param line", () => {
  it("writes @param [periods] and reads the description back", async () => {
    const { publishLambda, parseLambdaFile } = await import("../../src/index.js");
    const r = publishLambda({
      name: "ANA.GROW",
      formula: "LAMBDA(value, rate, [periods], value * (1 + rate) ^ IF(ISOMITTED(periods), 1, periods))",
      doc: "Grows a value. @param value the start @param [periods] how many; 1 when omitted",
    });
    expect(r.text).toContain("# params     value, rate, [periods]");
    expect(r.text).toContain("# @param [periods] how many; 1 when omitted");
    expect(r.text).toContain("# @param value the start");
    const f = parseLambdaFile("ANA.GROW.lambda", r.text);
    expect(f.fn?.paramDocs).toEqual({ value: "the start", periods: "how many; 1 when omitted" });
  });
});
