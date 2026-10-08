// The author's three ideas after the library rerun (FEEDBACK 2026-10-07): the base's text
// kept in `library-bases/` for the three-way diff; `@param` names checked against the
// LAMBDA's parameters; *Record library base* on an identical copy without one.
import { describe, expect, it } from "vitest";
import {
  applyEdits,
  baseFilePath,
  baseFiles,
  baseFileText,
  BASES_DIR,
  buildWorkbook,
  checkProject,
  definitionBase,
  docParamSpans,
  findEntry,
  isBasePath,
  libraryBaseRecordings,
  libraryFunctionBase,
  libraryFunctionHash,
  libraryInsertion,
  libraryReplacement,
  libStatusJson,
  parseBaseFile,
  projectLibraryStatus,
  publishedBase,
  publishLambda,
  pullProject,
  readBases,
  readLibrary,
  readWorkbook,
  renderLibStatus,
  SourceModel,
  type Library,
} from "../../src/index.js";
import { fixture, pulled } from "../build/helpers.js";

const lambda = (name: string, summary: string, def: string, params = "x") => `# name       ${name}\n# summary    ${summary}\n# params     ${params}\n#\n# Why ${name}.\n\n${def}\n`;

const SPREAD = "LAMBDA(total, periods, total / periods)";
const SPREAD_LOCAL = "LAMBDA(total, periodi, total / periodi)";
const SPREAD_LIB2 = "LAMBDA(total, periods, ROUND(total / periods, 2))";
const libOf = (def: string): Library => readLibrary({ "FN.SPREAD.lambda": lambda("FN.SPREAD", "Spreads a total.", def, "total, periods") });

// ---- 1. the base's text kept ---------------------------------------------------------------------

describe("library-bases/: the base's text, kept by the actions that write @from", () => {
  it("a base file reads back; one whose text does not give its hash is not taken", () => {
    const fn = libOf(SPREAD).get("FN.SPREAD")!;
    const b = libraryFunctionBase(fn);
    expect(b).toMatchObject({ hash: libraryFunctionHash(fn), name: "FN.SPREAD", library: "FN.SPREAD.lambda", display: SPREAD });
    expect(b.stored).toContain("_xlpm.total");
    expect(baseFilePath(b.hash)).toBe(`${BASES_DIR}/${b.hash}.json`);
    expect(isBasePath(baseFilePath(b.hash))).toBe(true);
    expect(isBasePath("names/FN.xln")).toBe(false);
    expect(isBasePath(`${BASES_DIR}/x/y.json`)).toBe(false);
    expect(parseBaseFile(baseFileText(b))).toEqual(b);
    expect(parseBaseFile(baseFileText({ ...b, stored: b.stored.replace("/", "*") }))).toBeUndefined();
    expect(parseBaseFile("not json")).toBeUndefined();
    expect(readBases({ ...baseFiles([b]), "names/FN.xln": "", [`${BASES_DIR}/bad.json`]: "{}" })).toEqual(new Map([[b.hash, b]]));
  });

  it("Insert, Take and Publish hand over the base they record", () => {
    const lib = readLibrary({ "FN.INC.lambda": lambda("FN.INC", "Adds one.", "LAMBDA(x, x + 1)"), "FN.TWO.lambda": lambda("FN.TWO", "Adds two.", "LAMBDA(x, FN.INC(FN.INC(x)))") });
    const ins = libraryInsertion(lib, {}, "FN.TWO")!;
    expect(ins.bases.map((b) => [b.name, b.hash])).toEqual([
      ["FN.INC", libraryFunctionHash(lib.get("FN.INC")!)],
      ["FN.TWO", libraryFunctionHash(lib.get("FN.TWO")!)],
    ]);
    // Publish: the version the library file holds afterwards, laid out as written there.
    const p = publishLambda({ name: "FN.INC", doc: "Adds one.", formula: "LAMBDA(x,\n  x + 5\n)" }, { path: "FN.INC.lambda", text: lambda("FN.INC", "Adds one.", "LAMBDA(x, x + 1)") });
    expect(p.base).toMatchObject({ hash: publishedBase("LAMBDA(x,\n  x + 5\n)", "FN.INC"), library: "FN.INC.lambda", display: "LAMBDA(x,\n  x + 5\n)" });
    const fresh = publishLambda({ name: "FN.NEW", doc: "New.", formula: "LAMBDA(y, y)" });
    expect(fresh.base?.hash).toBe(definitionBase("LAMBDA(y, y)", "FN.NEW"));
    expect(publishLambda({ name: "FN.K", doc: undefined, formula: "1" }).base).toBeUndefined();
  });

  it("both changed: the diffs come from the kept base, without the workbook or a backup; missing, the note as before", () => {
    const lib0 = libOf(SPREAD);
    const base = libraryFunctionBase(lib0.get("FN.SPREAD")!);
    const lib = libOf(SPREAD_LIB2);
    const files = { "names/FN.xln": `@from(lib #${base.hash})\nFN.SPREAD = ${SPREAD_LOCAL};\n` };
    const without = projectLibraryStatus(lib, files, undefined, { target: "p", library: "lib" });
    expect(without.items[0]!.state).toBe("both-changed");
    expect(renderLibStatus(without).text).toContain(`the base #${base.hash}'s text is not at hand (not in ${BASES_DIR}/,`);
    // Among the project's files…
    const kept = projectLibraryStatus(lib, { ...files, ...baseFiles([base]) }, undefined, { target: "p", library: "lib" });
    const item = kept.items[0]!;
    expect(item).toMatchObject({ state: "both-changed", baseDisplay: SPREAD, baseSource: `${BASES_DIR}/${base.hash}.json` });
    const text = renderLibStatus(kept).text;
    expect(text).toContain(`the base #${base.hash}'s text: from ${BASES_DIR}/${base.hash}.json`);
    expect(text).toContain(`here, since the base #${base.hash} (- base, + copy):\n        - ${SPREAD}\n        + ${SPREAD_LOCAL}`);
    expect(text).toContain(`in the library, since the base #${base.hash} (- base, + library #${item.libraryHash}):\n        - ${SPREAD}\n        + ${SPREAD_LIB2}`);
    expect((libStatusJson(kept) as { items: { baseSource?: string }[] }).items[0]!.baseSource).toBe(`${BASES_DIR}/${base.hash}.json`);
    // … or read apart (the editor): the same.
    const apart = projectLibraryStatus(lib, files, undefined, { target: "p", library: "lib", bases: new Map([[base.hash, base]]) });
    expect(apart.items[0]!.baseDiffs).toEqual(item.baseDiffs);
    // Another version's file says nothing about this base.
    const other = libraryFunctionBase(lib.get("FN.SPREAD")!);
    expect(projectLibraryStatus(lib, { ...files, ...baseFiles([other]) }, undefined, { target: "p", library: "lib" }).items[0]!.baseDiffs).toBeUndefined();
  });

  it("the build ignores library-bases/, and a pull writes nothing there (so a fresh pull keeps it)", () => {
    const wb0 = fixture("probe_mac.xlsx");
    const lib = libOf(SPREAD);
    const fn = lib.get("FN.SPREAD")!;
    const ins = libraryInsertion(lib, {}, "FN.SPREAD")!;
    const files: Record<string, string> = { ...pulled(wb0), "names/FN.xln": applyEdits("", ins.edits[0]!.edits), ...baseFiles(ins.bases) };
    const r = buildWorkbook({ workbook: wb0, fileName: "book.xlsx", files }, { embed: false });
    expect(r.status, JSON.stringify(r.plan.problems)).toBe("built");
    expect(Object.keys(r.files).filter((p) => p.startsWith(BASES_DIR))).toEqual([]);
    const again = pullProject(r.bytes!, "book.xlsx");
    expect(Object.keys(again.files).some((p) => p.startsWith(BASES_DIR))).toBe(false);
    // The pulled project, with the kept folder beside it, still finds the base.
    const report = projectLibraryStatus(lib, { ...again.files, ...baseFiles(ins.bases) }, readWorkbook(r.bytes!), { target: "p", library: "lib" });
    expect(report.items[0]).toMatchObject({ state: "identical", base: libraryFunctionHash(fn) });
  });
});

// ---- 2. @param drift -----------------------------------------------------------------------------

function problems(text: string, path = "names/FN.xln") {
  const m = new SourceModel();
  m.setFile(path, text);
  return checkProject(m).filter((p) => p.code === "doc-param" || p.code === "doc-param-missing");
}

describe("@param names against the LAMBDA's parameters", () => {
  it("the feedback's case: a renamed parameter leaves @param periods behind; the fix renames it", () => {
    const text = `/**\n * Spreads a total.\n * @param total the amount\n * @param periods how many\n */\nFN.SPREAD = ${SPREAD_LOCAL};\n`;
    const ps = problems(text);
    expect(ps.map((p) => [p.severity, text.slice(p.start, p.end)])).toEqual([["warning", "periods"]]);
    expect(ps[0]!.message).toContain("@param periods names no parameter of the LAMBDA (its parameters: total, periodi)");
    expect(ps[0]!.key).toBe("FN.SPREAD");
    const [rename, remove] = ps[0]!.fixes!;
    expect(rename!.title).toBe("Rename @param periods to periodi");
    expect(applyEdits(text, [rename!])).toContain(" * @param periodi how many\n */");
    expect(problems(applyEdits(text, [rename!]))).toEqual([]);
    // Removing it leaves periodi undocumented while total is: a hint.
    expect(remove!.title).toBe("Remove @param periods");
    const removed = applyEdits(text, [remove!]);
    expect(removed).toBe(`/**\n * Spreads a total.\n * @param total the amount\n */\nFN.SPREAD = ${SPREAD_LOCAL};\n`);
    const hint = problems(removed);
    expect(hint.map((p) => [p.severity, p.code])).toEqual([["hint", "doc-param-missing"]]);
    expect(hint[0]!.message).toContain("documents some parameters but not periodi");
  });

  it("optional parameters: @param [p] and @param p both match [p]", () => {
    expect(problems("/** Doc. @param x the x @param [p] opt */\nFN.A = LAMBDA(x, [p], x);\n")).toEqual([]);
    expect(problems("/** Doc. @param x the x @param p opt */\nFN.A = LAMBDA(x, [p], x);\n")).toEqual([]);
    // Renamed in brackets: the name inside them is replaced.
    const text = "/** Doc. @param x the x @param [q] opt */\nFN.A = LAMBDA(x, [p], x);\n";
    const ps = problems(text);
    expect(ps[0]!.message).toContain("its parameters: x, [p]");
    expect(applyEdits(text, [ps[0]!.fixes![0]!])).toBe("/** Doc. @param x the x @param [p] opt */\nFN.A = LAMBDA(x, [p], x);\n");
    // Removed on one line.
    expect(applyEdits(text, [ps[0]!.fixes![1]!])).toBe("/** Doc. @param x the x */\nFN.A = LAMBDA(x, [p], x);\n");
  });

  it("no @param at all, or every parameter documented: nothing; case does not matter; not a LAMBDA: not checked", () => {
    expect(problems("/** Doc. */\nFN.A = LAMBDA(x, y, x + y);\n")).toEqual([]);
    expect(problems("/** Doc.\n @param X a\n @param y b */\nFN.A = LAMBDA(x, y, x + y);\n")).toEqual([]);
    expect(problems("/** Doc. @param x a */\nrate = 0.1;\n")).toEqual([]);
    expect(problems("/** Doc. @param x a */\nFN.A = (LAMBDA(x, y, x + y));\n").map((p) => p.code)).toEqual(["doc-param-missing"]);
  });

  it("the rename goes to the parameter at the same position, unless that one is documented; then only the removal", () => {
    // `@param b` at position 1 while `y` is documented at position 0 and `x` is not: no rename target at 1 that is free.
    const text = "/**\n * @param y first\n * @param b second\n */\nFN.A = LAMBDA(x, y, x + y);\n";
    const ps = problems(text);
    const warn = ps.find((p) => p.code === "doc-param")!;
    expect(warn.fixes!.map((f) => f.title)).toEqual(["Remove @param b"]);
    expect(applyEdits(text, [warn.fixes![0]!])).toBe("/**\n * @param y first\n */\nFN.A = LAMBDA(x, y, x + y);\n");
    expect(ps.find((p) => p.code === "doc-param-missing")!.message).toContain("but not x");
    // A LAMBDA with no parameters.
    expect(problems("/** @param x */\nFN.K = LAMBDA(1);\n")[0]!.message).toContain("it takes none");
  });

  it("the spans and removal ranges of @param tags", () => {
    const text = "/** Sum.\n * @param a one\n *   more\n * @param [b] two */";
    const spans = docParamSpans(text, 3, text.length - 2);
    expect(spans.map((s) => [s.name, s.bracketed, text.slice(s.nameStart, s.nameEnd)])).toEqual([
      ["a", false, "a"],
      ["b", true, "b"],
    ]);
    expect(text.slice(0, spans[0]!.removeStart) + text.slice(spans[0]!.removeEnd)).toBe("/** Sum.\n * @param [b] two */");
    expect(text.slice(0, spans[1]!.removeStart) + text.slice(spans[1]!.removeEnd)).toBe("/** Sum.\n * @param a one\n *   more */");
  });
});

// ---- 3. Record library base ----------------------------------------------------------------------

describe("Record library base: an identical copy without @from", () => {
  const lib = () =>
    readLibrary({
      "FN.INC.lambda": lambda("FN.INC", "Adds one.", "LAMBDA(x, x + 1)"),
      "FN.DBL.lambda": lambda("FN.DBL", "Doubles.", "LAMBDA(x, 2 * x)"),
      "FN.ODD.lambda": lambda("FN.ODD", "Odd.", "LAMBDA(x, x * 3)"),
      "FN.NEW.lambda": lambda("FN.NEW", "New.", "LAMBDA(x, x)"),
    });
  const files = () => ({
    "names/FN.xln": "/** Doubles. */\nFN.DBL = LAMBDA(x, 2*x);\n\n/** Adds one. */\n@hidden\nFN.INC = LAMBDA(x, x + 1);\n\nFN.ODD = LAMBDA(x, x * 4);\n",
  });

  it("status keeps them identical, with a note naming the action; never records one itself", () => {
    const r = projectLibraryStatus(lib(), files(), undefined, { target: "p", library: "lib" });
    const inc = r.items.find((i) => i.name === "FN.INC")!;
    expect(inc).toMatchObject({ state: "identical", noBase: true });
    expect(inc.note).toBe("same definition as the library; no base recorded: Record library base");
    expect(r.items.find((i) => i.name === "FN.ODD")!).toMatchObject({ state: "differs" });
    expect(r.items.find((i) => i.name === "FN.ODD")!.noBase).toBeUndefined();
  });

  it("all of them: @from on each, with its base's text; the others skipped with the reason", () => {
    const l = lib();
    const f = files();
    const all = libraryBaseRecordings(l, f);
    expect(all.recorded.map((x) => x.name)).toEqual(["FN.DBL", "FN.INC"]);
    expect(all.skipped).toEqual([]);
    const after = applyEdits(f["names/FN.xln"], all.recorded.map((x) => x.edit));
    const h = (n: string) => libraryFunctionHash(l.get(n)!);
    expect(after).toBe(`/** Doubles. */\n@from(lib #${h("FN.DBL")})\nFN.DBL = LAMBDA(x, 2*x);\n\n/** Adds one. */\n@hidden\n@from(lib #${h("FN.INC")})\nFN.INC = LAMBDA(x, x + 1);\n\nFN.ODD = LAMBDA(x, x * 4);\n`);
    expect(all.recorded[0]!.base).toEqual(libraryFunctionBase(l.get("FN.DBL")!));
    // Then: identical with a base, nothing left to record.
    const r = projectLibraryStatus(l, { "names/FN.xln": after }, undefined, { target: "p", library: "lib" });
    expect(r.items.find((i) => i.name === "FN.INC")).toMatchObject({ state: "identical", base: h("FN.INC") });
    expect(r.items.some((i) => i.noBase)).toBe(false);
    expect(libraryBaseRecordings(l, { "names/FN.xln": after }).recorded).toEqual([]);
  });

  it("one by name; refused for one not identical, one with a base, one missing or unknown", () => {
    const l = lib();
    const one = libraryBaseRecordings(l, files(), ["fn.inc"]);
    expect(one.recorded.map((x) => [x.name, x.path])).toEqual([["FN.INC", "names/FN.xln"]]);
    const no = libraryBaseRecordings(l, files(), ["FN.ODD", "FN.NEW", "FN.NOPE"]);
    expect(no.recorded).toEqual([]);
    expect(no.skipped.map((s) => s.name)).toEqual(["FN.ODD", "FN.NEW", "FN.NOPE"]);
    expect(no.skipped[0]!.reason).toContain("differs, not identical to the library");
    expect(no.skipped[1]!.reason).toBe("not a workbook name of the project");
    expect(no.skipped[2]!.reason).toBe("not in the library");
    const based = applyEdits(files()["names/FN.xln"], [one.recorded[0]!.edit]);
    expect(libraryBaseRecordings(l, { "names/FN.xln": based }, ["FN.INC"]).skipped[0]!.reason).toContain("already records a base");
  });

  it("Take on an entry gives the same @from as Record would", () => {
    const l = lib();
    const text = files()["names/FN.xln"];
    const rec = libraryBaseRecordings(l, { "names/FN.xln": text }, ["FN.INC"]).recorded[0]!;
    const taken = applyEdits(text, libraryReplacement(text, findEntry(text, "FN.INC")!, l.get("FN.INC")!));
    expect(taken).toContain(`@from(lib #${rec.hash})\nFN.INC`);
  });
});
