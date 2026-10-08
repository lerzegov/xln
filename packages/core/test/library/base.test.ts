// The library base (decided 2026-10-07): `@from(lib #…)` in the source, `lib#…` in the
// provenance tag, written only by Insert, Take and Publish, carried by the build, written
// back by pull; lib status three-way on it. The repro of the feedback entry "A local edit,
// once built, reads as outdated" is the flow at the end.
import { describe, expect, it } from "vitest";
import {
  applyEdits,
  buildWorkbook,
  checkProject,
  commentLength,
  COMMENT_MAX,
  definitionBase,
  findEntry,
  formatEntry,
  formatProvenanceTag,
  libraryBaseEdit,
  libraryDoc,
  libraryFunctionHash,
  libraryHash,
  libraryInsertion,
  libraryReplacement,
  LIBRARY_DOC_MAX,
  parseLibBase,
  parseModule,
  projectLibraryStatus,
  publishedBase,
  publishLambda,
  pullProject,
  readLibrary,
  readWorkbook,
  renderLibStatus,
  SourceModel,
  splitProvenance,
  unbuiltEdits,
  withProvenance,
  type Library,
  type LibStatusReport,
} from "../../src/index.js";
import { fixture, pulled } from "../build/helpers.js";

const lambda = (name: string, summary: string, def: string) => `# name       ${name}\n# summary    ${summary}\n# params     x\n#\n# Why ${name}.\n\n${def}\n`;

describe("@from(lib #…): parser and formatter", () => {
  it("reads the base among the other annotations, in any order", () => {
    const pm = parseModule("/** Doc. */\n@hidden\n@from(lib #35392A)\nFN.X = LAMBDA(x, x);\n\n@from( lib  #abcdef )\nFN.Y = LAMBDA(x, x);\n");
    expect(pm.diagnostics).toEqual([]);
    const [x, y] = pm.entries;
    expect(x).toMatchObject({ name: "FN.X", hidden: true, doc: "Doc.", from: "35392a" });
    expect(x!.annotations.map((a) => a.name)).toEqual(["hidden", "from"]);
    const a = x!.annotations[1]!;
    expect("/** Doc. */\n@hidden\n@from(lib #35392A)\nFN.X".slice(a.offset, a.end)).toBe("@from(lib #35392A)");
    expect(y!.from).toBe("abcdef");
    // In a sheet file it stacks with @workbook like the others.
    const sheet = parseModule("@workbook\n@from(lib #000001)\nrate = 0.1;\n", { sheet: "IS" });
    expect(sheet.entries[0]).toMatchObject({ workbook: true, from: "000001" });
  });

  it("a malformed one sets no base; the argument says why", () => {
    for (const bad of ["@from", "@from()", "@from(353921)", "@from(lib 353921)", "@from(lib #35392)", "@from(lib #35392g)", "@from(git #353921)"]) {
      const pm = parseModule(`${bad}\nFN.X = LAMBDA(x, x);\n`);
      expect(pm.entries[0]!.from, bad).toBeUndefined();
    }
    expect(parseLibBase("lib #353921")).toEqual({ hash: "353921" });
    expect(parseLibBase("lib #35392")).toMatchObject({ error: expect.stringContaining("6 hex digits") });
    expect(parseLibBase(undefined)).toMatchObject({ error: expect.stringContaining("@from needs the library version") });
  });

  it("is written after the doc comment and the other annotations, and reads back", () => {
    const text = formatEntry({ name: "FN.X", doc: "Doc.", hidden: true, from: "353921", formula: "LAMBDA(x, x)" });
    expect(text).toBe("/** Doc. */\n@hidden\n@from(lib #353921)\nFN.X = LAMBDA(x, x);");
    expect(parseModule(text).entries[0]).toMatchObject({ from: "353921", hidden: true, doc: "Doc." });
  });

  it("the checker: a malformed or doubled @from is an error; outside a module file a warning", () => {
    const problems = (path: string, text: string) => {
      const m = new SourceModel();
      m.setFile(path, text);
      return checkProject(m).filter((p) => p.code === "annotation").map((p) => [p.severity, text.slice(p.start, p.end), p.message] as const);
    };
    expect(problems("names/FN.xln", "@from(lib #353921)\nFN.X = LAMBDA(x, x);\n")).toEqual([]);
    const bad = problems("names/FN.xln", "@from(lib 353921)\nFN.X = LAMBDA(x, x);\n");
    expect(bad).toEqual([["error", "@from(lib 353921)", expect.stringContaining("write @from(lib #abc123)")]]);
    const twice = problems("names/FN.xln", "@from(lib #353921)\n@from(lib #353922)\nFN.X = LAMBDA(x, x);\n");
    expect(twice).toEqual([["error", "@from(lib #353922)", "@from is given twice: a copy has one library base; keep one"]]);
    const sheet = problems("names/sheets/IS.xln", "@from(lib #353921)\nrate = 0.1;\n");
    expect(sheet.map((p) => p[0])).toEqual(["warning"]);
    expect(sheet[0]![2]).toContain("no provenance tag carries it");
  });
});

describe("the provenance tag carries the base", () => {
  it("formats and splits [xln FN 1.2 #636cf1 lib#353921]; old tags read as before", () => {
    const t = { module: "FN", version: "1.2", hash: "636cf1", lib: "353921" };
    expect(formatProvenanceTag(t)).toBe("[xln FN 1.2 #636cf1 lib#353921]");
    expect(splitProvenance("Doc. [xln FN 1.2 #636cf1 lib#353921]")).toEqual({ comment: "Doc.", tag: t });
    expect(splitProvenance("[xln FN #636cf1 lib#353921]")).toEqual({ comment: undefined, tag: { module: "FN", hash: "636cf1", lib: "353921" } });
    expect(splitProvenance("Doc. [xln FN #636cf1]")).toEqual({ comment: "Doc.", tag: { module: "FN", hash: "636cf1" } });
    // Not a tag: kept as the author's text.
    expect(splitProvenance("Doc. [xln FN #636cf1 lib#xyz]")).toEqual({ comment: "Doc. [xln FN #636cf1 lib#xyz]" });
    expect(splitProvenance("Doc. [xln lib#353921]")).toEqual({ comment: "Doc. [xln lib#353921]" });
  });

  it("the doc comment of a library function leaves room for the tag with its base", () => {
    expect(LIBRARY_DOC_MAX).toBe(212);
    const long = "word ".repeat(80).trim() + ".";
    const d = libraryDoc(long, ["a", "b"]);
    expect(d.shortened).toBe(true);
    expect(commentLength(d.doc)).toBeLessThanOrEqual(212);
    // A module name of 10 characters and a version of 4: the tag still fits.
    const tagged = withProvenance(d.doc, { module: "FINANCEXYZ", version: "1.10", hash: "636cf1", lib: "353921" });
    expect(commentLength(tagged)).toBeLessThanOrEqual(COMMENT_MAX);
  });

  it("libraryHash: layout and number spelling do not count; the definition does", () => {
    expect(libraryHash("_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1E-14)")).toBe(libraryHash("_xlfn.LAMBDA(_xlpm.x, _xlpm.x + 0.00000000000001)"));
    expect(libraryHash("_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)")).not.toBe(libraryHash("_xlfn.LAMBDA(_xlpm.x,_xlpm.x+2)"));
    expect(definitionBase("LAMBDA(x,\n  x + 1)", "FN.A")).toBe(definitionBase("LAMBDA(x, x+1)", "FN.B"));
    expect(definitionBase("LAMBDA(x, x+1)", "FN.A")).toMatch(/^[0-9a-f]{6}$/);
  });
});

// ---- build, plan, pull ------------------------------------------------------------------------

const SOURCE = (from?: string) => `// module: FN\n\n/** Adds one. */\n${from ? `@from(lib #${from})\n` : ""}FN.INC = LAMBDA(x, x + 1);\n`;

function built(source: string, bytes = fixture("probe_mac.xlsx"), files = pulled(fixture("probe_mac.xlsx"))): { bytes: Uint8Array; files: Record<string, string> } {
  const f = { ...files, "names/FN.xln": source };
  const r = buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files: f }, { embed: false });
  expect(r.status, JSON.stringify(r.plan.problems)).toBe("built");
  return { bytes: r.bytes!, files: { ...f, ...r.files } };
}

const comment = (bytes: Uint8Array, name: string) => readWorkbook(bytes).definedNames.find((d) => d.name === name)!.comment;

describe("build and pull carry the base", () => {
  it("the build writes lib# after the hash; pull writes @from back (round trip)", () => {
    const { bytes } = built(SOURCE("353921"));
    expect(comment(bytes, "FN.INC")).toMatch(/^Adds one\. \[xln FN #[0-9a-f]{6} lib#353921\]$/);
    const back = pullProject(bytes, "book.xlsx").files["names/FN.xln"]!;
    expect(back).toContain("/** Adds one. */\n@from(lib #353921)\nFN.INC = LAMBDA(x, x + 1);");
    // Built again from the pulled project: nothing to do.
    const again = buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files: pullProject(bytes, "book.xlsx").files }, { embed: false });
    expect(again.status).toBe("up-to-date");
  });

  it("an old tag without lib#: no annotation", () => {
    const { bytes } = built(SOURCE());
    expect(comment(bytes, "FN.INC")).toMatch(/^Adds one\. \[xln FN #[0-9a-f]{6}\]$/);
    expect(pullProject(bytes, "book.xlsx").files["names/FN.xln"]).not.toContain("@from");
  });

  it("adding, changing or removing @from is a provenance change of the name, never a definition change or a conflict", () => {
    const first = built(SOURCE());
    const plan = (src: string) => {
      const r = buildWorkbook({ workbook: first.bytes, fileName: "book.xlsx", files: { ...first.files, "names/FN.xln": src } }, { embed: false, dryRun: true });
      expect(r.plan.conflicts).toEqual([]);
      return r.plan.changeSet.changes.filter((c) => c.op === "set-name").map((c) => c.op === "set-name" && [c.name, c.fields, c.comment]);
    };
    expect(plan(SOURCE("353921"))).toEqual([["FN.INC", ["provenance"], expect.stringMatching(/ lib#353921\]$/)]]);
    const second = built(SOURCE("353921"), first.bytes, first.files);
    const r = buildWorkbook({ workbook: second.bytes, fileName: "book.xlsx", files: { ...second.files, "names/FN.xln": SOURCE("aaaaaa") } }, { embed: false, dryRun: true });
    expect(r.plan.changeSet.changes.map((c) => c.op === "set-name" && c.fields)).toEqual([["provenance"]]);
    const removed = buildWorkbook({ workbook: second.bytes, fileName: "book.xlsx", files: { ...second.files, "names/FN.xln": SOURCE() } }, { embed: false, dryRun: true });
    expect(removed.plan.changeSet.changes.map((c) => c.op === "set-name" && [c.fields, c.comment])).toEqual([[["provenance"], expect.stringMatching(/^Adds one\. \[xln FN #[0-9a-f]{6}\]$/)]]);
  });

  it("a comment too long for the tag with its base: written without the tag, and the build says the base is not carried", () => {
    const doc = "x".repeat(230);
    const src = `/** ${doc} */\n@from(lib #353921)\nFN.INC = LAMBDA(x, x + 1);\n`;
    const base = fixture("probe_mac.xlsx");
    const r = buildWorkbook({ workbook: base, fileName: "book.xlsx", files: { ...pulled(base), "names/FN.xln": src } }, { embed: false });
    expect(r.status).toBe("built");
    expect(comment(r.bytes!, "FN.INC")).toBe(doc);
    // Said once: the checker's warning on the doc comment (the build's own is for plans without it).
    expect(r.plan.problems.filter((p) => p.code === "provenance").map((p) => p.message)).toEqual([
      "FN.INC: the doc comment is 230 characters; with its provenance tag (28) it passes Excel's 255, so the build writes it without the tag and the workbook won't record its library base, @from(lib #353921): shorten it by 3 characters",
    ]);
  });

  it("the pull guard lists a @from the workbook's tag does not carry yet", () => {
    const { bytes, files } = built(SOURCE("353921"));
    expect(unbuiltEdits({ workbook: bytes, fileName: "book.xlsx", files })).toEqual([]);
    const edits = unbuiltEdits({ workbook: bytes, fileName: "book.xlsx", files: { ...files, "names/FN.xln": SOURCE("aaaaaa") } });
    expect(edits).toEqual([{ key: "FN.INC", file: "names/FN.xln", line: 5, what: "@from(lib #aaaaaa) in the source, not built (the workbook's tag says lib#353921)" }]);
    const removed = unbuiltEdits({ workbook: bytes, fileName: "book.xlsx", files: { ...files, "names/FN.xln": SOURCE() } });
    expect(removed.map((e) => e.what)).toEqual(["@from(lib #353921) removed in the source, not built"]);
  });
});

// ---- the edits ----------------------------------------------------------------------------------

describe("who writes @from", () => {
  const lib = () => readLibrary({ "FN.INC.lambda": lambda("FN.INC", "Adds one.", "LAMBDA(x, x + 1)"), "FN.TWO.lambda": lambda("FN.TWO", "Adds two.", "LAMBDA(x, FN.INC(FN.INC(x)))") });

  it("Insert writes the library's version as the base of each function it adds", () => {
    const ins = libraryInsertion(lib(), {}, "FN.TWO")!;
    const text = applyEdits("", ins.edits[0]!.edits);
    expect(text).toContain(`@from(lib #${libraryFunctionHash(lib().get("FN.INC")!)})\nFN.INC = LAMBDA(x, x + 1);`);
    expect(text).toContain(`@from(lib #${libraryFunctionHash(lib().get("FN.TWO")!)})\nFN.TWO = `);
  });

  it("Take writes or rewrites it, wherever the entry's annotations are", () => {
    const fn = lib().get("FN.INC")!;
    const h = libraryFunctionHash(fn);
    const take = (text: string) => applyEdits(text, libraryReplacement(text, findEntry(text, "FN.INC")!, fn));
    expect(take("FN.INC = LAMBDA(x, x + 9);\n")).toBe(`/**\n * Adds one.\n * @param x\n */\n@from(lib #${h})\nFN.INC = LAMBDA(x, x + 1);\n`);
    expect(take("/** Old. */\n@from(lib #000000)\n@hidden\nFN.INC = LAMBDA(x, x + 9);\n")).toBe(`/**\n * Adds one.\n * @param x\n */\n@from(lib #${h})\n@hidden\nFN.INC = LAMBDA(x, x + 9);\n`.replace("x + 9", "x + 1"));
    expect(take("/** Old. */\n  FN.INC = LAMBDA(x, x + 9);\n")).toBe(`/**\n * Adds one.\n * @param x\n */\n  @from(lib #${h})\n  FN.INC = LAMBDA(x, x + 1);\n`);
    expect(take("/** Old. */ @hidden FN.INC = LAMBDA(x, x + 9);\n")).toBe(`/**\n * Adds one.\n * @param x\n */ @hidden @from(lib #${h}) FN.INC = LAMBDA(x, x + 1);\n`);
  });

  it("Publish: the published version, which the library then reads back with the same hash", () => {
    const text = "@from(lib #000000)\nFN.INC = LAMBDA(x,\n  x + 5\n);\n";
    const e = findEntry(text, "FN.INC")!;
    const p = publishLambda({ name: e.name, doc: e.doc, formula: e.formula }, { path: "FN.INC.lambda", text: lambda("FN.INC", "Adds one.", "LAMBDA(x, x + 1)") });
    const h = publishedBase(e.formula, e.name);
    expect(libraryFunctionHash(readLibrary({ "FN.INC.lambda": p.text }).get("FN.INC")!)).toBe(h);
    expect(applyEdits(text, [libraryBaseEdit(text, e, h)!])).toBe(`@from(lib #${h})\nFN.INC = LAMBDA(x,\n  x + 5\n);\n`);
    // Already recorded: nothing to do.
    const done = `@from(lib #${h})\nFN.INC = LAMBDA(x, x + 5);\n`;
    expect(libraryBaseEdit(done, findEntry(done, "FN.INC")!, h)).toBeUndefined();
  });
});

// ---- the five states, the feedback's repro as a flow --------------------------------------------

function states(r: LibStatusReport): Record<string, string> {
  return Object.fromEntries(r.items.filter((i) => i.name === "FN.SPREAD").map((i) => [i.name, i.state]));
}

describe("lib status three-way: the feedback's flow (FN.SPREAD, a parameter renamed locally)", () => {
  const SPREAD = "LAMBDA(total, periods, total / periods)";
  const SPREAD_LOCAL = "LAMBDA(total, periodi, total / periodi)";
  const libOf = (def: string): Library => readLibrary({ "FN.SPREAD.lambda": lambda("FN.SPREAD", "Spreads a total.", def).replace("# params     x", "# params     total, periods") });

  it("no base: differs; Take, build: identical; edit, build: modified; library changes: both changed; edit undone: outdated", () => {
    const wb0 = fixture("probe_mac.xlsx");
    let files: Record<string, string> = { ...pulled(wb0), "names/FN.xln": `/** Spreads a total. */\nFN.SPREAD = ${SPREAD_LOCAL};\n` };
    let lib = libOf(SPREAD);
    const status = (wb: Uint8Array | undefined, backup?: Uint8Array) =>
      projectLibraryStatus(lib, files, wb ? readWorkbook(wb) : undefined, { target: "book.xln", library: "lib", ...(backup ? { backup: readWorkbook(backup) } : {}) });
    const build = (wb: Uint8Array) => {
      const r = buildWorkbook({ workbook: wb, fileName: "book.xlsx", files }, { embed: false });
      expect(r.status, JSON.stringify(r.plan.problems)).toBe("built");
      files = { ...files, ...r.files };
      return r.bytes!;
    };
    // Built before the library base existed: the tag has no lib#.
    const wb1 = build(wb0);
    expect(states(status(wb1))).toEqual({ "FN.SPREAD": "differs" });

    // Take: the library's version and its base.
    const text = files["names/FN.xln"]!;
    files["names/FN.xln"] = applyEdits(text, libraryReplacement(text, findEntry(text, "FN.SPREAD")!, lib.get("FN.SPREAD")!));
    const libHash = libraryFunctionHash(lib.get("FN.SPREAD")!);
    expect(files["names/FN.xln"]).toContain(`@from(lib #${libHash})\nFN.SPREAD = ${SPREAD};`);
    const wb2 = build(wb1);
    expect(states(status(wb2))).toEqual({ "FN.SPREAD": "identical" });
    expect(comment(wb2, "FN.SPREAD")).toMatch(new RegExp(` lib#${libHash}\\]$`));

    // A local edit, built: modified, not outdated (the bug of 2026-10-07).
    files["names/FN.xln"] = files["names/FN.xln"]!.replace(SPREAD, SPREAD_LOCAL);
    const wb3 = build(wb2);
    expect(states(status(wb3))).toEqual({ "FN.SPREAD": "modified" });
    expect(comment(wb3, "FN.SPREAD")).toMatch(new RegExp(` lib#${libHash}\\]$`));

    // The library changes too: both changed; the base's text comes from the backup (wb2).
    lib = libOf("LAMBDA(total, periods, ROUND(total / periods, 2))");
    const both = status(wb3, wb2);
    expect(states(both)).toEqual({ "FN.SPREAD": "both-changed" });
    const item = both.items.find((i) => i.name === "FN.SPREAD")!;
    expect(item.base).toBe(libHash);
    expect(item.baseDisplay).toBe(SPREAD);
    const text3 = renderLibStatus(both).text;
    expect(text3).toContain(`here, since the base #${libHash} (- base, + copy):\n        - ${SPREAD}\n        + ${SPREAD_LOCAL}`);
    expect(text3).toContain(`in the library, since the base #${libHash} (- base, + library #${item.libraryHash}):`);
    // Without the backup the base's text is not at hand.
    expect(renderLibStatus(status(wb3)).text).toContain(`the base #${libHash}'s text is not at hand`);

    // The local edit undone: the copy is its base, the library moved: outdated.
    files["names/FN.xln"] = files["names/FN.xln"]!.replace(SPREAD_LOCAL, SPREAD);
    expect(states(status(wb3))).toEqual({ "FN.SPREAD": "outdated" });
    // The workbook alone tells the same, from its tag, once built.
    const wb4 = build(wb3);
    expect(states(status(wb4))).toEqual({ "FN.SPREAD": "outdated" });
  });

  it("identical whatever the base says (the same change on both sides), with a note", () => {
    const lib = libOf(SPREAD_LOCAL);
    const files = { "names/FN.xln": `@from(lib #${definitionBase(SPREAD, "FN.SPREAD")})\nFN.SPREAD = ${SPREAD_LOCAL};\n` };
    const r = projectLibraryStatus(lib, files, undefined, { target: "p", library: "lib" });
    expect(states(r)).toEqual({ "FN.SPREAD": "identical" });
    expect(r.items[0]!.note).toContain("its recorded base #");
  });
});
