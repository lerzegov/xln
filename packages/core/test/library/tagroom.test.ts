// Room for the provenance tag (feedback 2026-10-07, "Publish doesn't warn that the doc
// comment plus the tag will exceed 255 characters"): the checker, Insert, Take and Publish
// say it before the build drops the tag, measured with the build's own tag code.
import { describe, expect, it } from "vitest";
import {
  applyEdits,
  buildProvenanceTag,
  buildWorkbook,
  checkProject,
  commentLength,
  COMMENT_MAX,
  docTagOverflow,
  docTagWarnings,
  findEntry,
  libraryBaseEdit,
  libraryInsertion,
  libraryReplacement,
  provenanceTagLength,
  publishedBase,
  readLibrary,
  readWorkbook,
  SourceModel,
  splitProvenance,
} from "../../src/index.js";
import { fixture, pulled } from "../build/helpers.js";

const BASE = "353921";
const comment = (bytes: Uint8Array, name: string) => readWorkbook(bytes).definedNames.find((d) => d.name === name)!.comment;

describe("the tag the build would write", () => {
  it("module, version, hash and lib#: only module files get one", () => {
    expect(buildProvenanceTag("names/FN.xln", "// @version 1.2\nFN.X = 1;\n", BASE)).toEqual({ module: "FN", version: "1.2", hash: "000000", lib: BASE });
    // ` [xln FN 1.2 #3f9a1c lib#353921]`: what lambdaFile.ts's LIBRARY_DOC_MAX reckons with.
    expect(provenanceTagLength(buildProvenanceTag("names/FN.xln", "// @version 1.2\n", BASE)!)).toBe(32);
    expect(provenanceTagLength(buildProvenanceTag("names/FN.xln", "")!)).toBe(" [xln FN #000000]".length);
    expect(buildProvenanceTag("names/_unmanaged.xln", "", BASE)).toBeUndefined();
    expect(buildProvenanceTag("names/sheets/IS.xln", "", BASE)).toBeUndefined();
  });

  it("the build agrees to the character: at 255 with the tag it tags; one more and it warns", () => {
    const base = fixture("probe_mac.xlsx");
    const tag = provenanceTagLength(buildProvenanceTag("names/FN.xln", "", BASE)!);
    const src = (n: number) => `/** ${"x".repeat(n)} */\n@from(lib #${BASE})\nFN.INC = LAMBDA(x, x + 1);\n`;
    const fits = buildWorkbook({ workbook: base, fileName: "book.xlsx", files: { ...pulled(base), "names/FN.xln": src(COMMENT_MAX - tag) } }, { embed: false });
    expect(fits.status).toBe("built");
    expect(commentLength(comment(fits.bytes!, "FN.INC")!)).toBe(COMMENT_MAX);
    expect(splitProvenance(comment(fits.bytes!, "FN.INC")!).tag!.lib).toBe(BASE);
    expect(fits.plan.problems.filter((p) => p.code === "provenance")).toEqual([]);

    const over = buildWorkbook({ workbook: base, fileName: "book.xlsx", files: { ...pulled(base), "names/FN.xln": src(COMMENT_MAX - tag + 1) } }, { embed: false });
    expect(over.status).toBe("built");
    expect(comment(over.bytes!, "FN.INC")).toBe("x".repeat(COMMENT_MAX - tag + 1));
    expect(over.plan.problems.filter((p) => p.code === "provenance").map((p) => p.message)).toEqual([
      `FN.INC: the doc comment is 228 characters; with its provenance tag (28) it passes Excel's 255, so the build writes it without the tag and the workbook won't record its library base, @from(lib #${BASE}): shorten it by 1 character`,
    ]);
  });

  it("the message: line breaks count 2; without a base it is the module and hash that go", () => {
    expect(docTagOverflow("a\nb".padEnd(240, "c"), buildProvenanceTag("names/FN.xln", ""))!.message).toBe(
      "the doc comment is 241 characters (a line break counts 2); with its provenance tag (17) it passes Excel's 255, so the build writes it without the tag and the workbook won't record its module and source hash: shorten it by 3 characters",
    );
    expect(docTagOverflow("x".repeat(200), buildProvenanceTag("names/FN.xln", "", BASE))).toBeUndefined();
    // Over 255 alone is the checker's error, not this warning.
    expect(docTagOverflow("x".repeat(256), buildProvenanceTag("names/FN.xln", "", BASE))).toBeUndefined();
    expect(docTagOverflow("x".repeat(250), undefined)).toBeUndefined();
  });
});

describe("the live checker (editor and xln check)", () => {
  const check = (path: string, text: string) => {
    const m = new SourceModel();
    m.setFile(path, text);
    return checkProject(m).filter((p) => p.code === "provenance" || p.code === "comment-length");
  };

  it("a warning on the doc comment of a name with @from, and of any module name", () => {
    const text = `// @version 1.2\n\n/** ${"x".repeat(230)} */\n@from(lib #${BASE})\nMTG.REPAYMENT = LAMBDA(x, x);\n`;
    const [p, ...rest] = check("names/MTG.xln", text);
    expect(rest).toEqual([]);
    expect(p).toMatchObject({ severity: "warning", code: "provenance", key: "MTG.REPAYMENT" });
    expect(text.slice(p!.start, p!.end)).toBe(`/** ${"x".repeat(230)} */`);
    expect(p!.message).toBe(
      `MTG.REPAYMENT: the doc comment is 230 characters; with its provenance tag (33) it passes Excel's 255, so the build writes it without the tag and the workbook won't record its library base, @from(lib #${BASE}): shorten it by 8 characters`,
    );
    expect(check("names/MTG.xln", `/** ${"x".repeat(240)} */\nMTG.X = 1;\n`).map((q) => q.message)).toEqual([
      "MTG.X: the doc comment is 240 characters; with its provenance tag (18) it passes Excel's 255, so the build writes it without the tag and the workbook won't record its module and source hash: shorten it by 3 characters",
    ]);
  });

  it("nothing where the build writes no tag; over 255 alone stays the error", () => {
    expect(check("names/_unmanaged.xln", `/** ${"x".repeat(250)} */\nX = 1;\n`)).toEqual([]);
    expect(check("names/sheets/IS.xln", `/** ${"x".repeat(250)} */\nX = 1;\n`)).toEqual([]);
    expect(check("names/MTG.xln", `/** ${"x".repeat(256)} */\nMTG.X = 1;\n`).map((q) => [q.severity, q.code])).toEqual([["error", "comment-length"]]);
  });
});

describe("Insert, Take, Publish say it before they write", () => {
  const LONG_VERSION = "// @version 2026.10.07-release-candidate\n";
  const lambda = (summary: string) => `# name       FN.INC\n# summary    ${summary}\n# params     x\n\nLAMBDA(x, x + 1)\n`;
  // A library doc is kept within LIBRARY_DOC_MAX, which leaves room for a short module and version: not this one.
  const lib = readLibrary({ "FN.INC.lambda": lambda(`Adds one ${"and says so at length ".repeat(12)}.`) });

  it("Insert: the inserted doc comment with the module's tag", () => {
    const files = { "names/FN.xln": `${LONG_VERSION}\nFN.OWN = 1;\n` };
    const ins = libraryInsertion(lib, files, "FN.INC")!;
    const after = { ...files };
    for (const fe of ins.edits) after[fe.path as "names/FN.xln"] = applyEdits(fe.create ? "" : files[fe.path as "names/FN.xln"], fe.edits);
    const w = docTagWarnings(after, ins.names);
    expect(w.map((x) => [x.name, x.path])).toEqual([["FN.INC", "names/FN.xln"]]);
    expect(w[0]!.message).toContain(`won't record its library base, @from(lib #`);
    // The same function into a module without a version: room enough.
    const plain = libraryInsertion(lib, {}, "FN.INC")!;
    expect(docTagWarnings({ "names/FN.xln": applyEdits("", plain.edits[0]!.edits) }, ["FN.INC"])).toEqual([]);
  });

  it("Take: the library's doc comment replacing the copy's", () => {
    const text = `${LONG_VERSION}\n/** Short. */\nFN.INC = LAMBDA(x, x + 2);\n`;
    const after = applyEdits(text, libraryReplacement(text, findEntry(text, "FN.INC")!, lib.get("FN.INC")!));
    expect(docTagWarnings({ "names/FN.xln": text }, ["FN.INC"])).toEqual([]);
    expect(docTagWarnings({ "names/FN.xln": after }, ["FN.INC"]).map((x) => x.name)).toEqual(["FN.INC"]);
  });

  it("Publish: the project's doc comment (the workbook's), with the base publish records", () => {
    // The feedback's case: a long doc that fits alone and without @from's 11 characters, not with them.
    const doc = "x".repeat(COMMENT_MAX - " [xln MTG #000000]".length - 2);
    const text = `/** ${doc} */\nMTG.REPAYMENT = LAMBDA(x, x);\n`;
    const files = { "names/MTG.xln": text };
    expect(docTagWarnings(files, ["MTG.REPAYMENT"])).toEqual([]);
    const e = findEntry(text, "MTG.REPAYMENT")!;
    const hash = publishedBase(e.formula, e.name);
    const after = applyEdits(text, [libraryBaseEdit(text, e, hash)!]);
    expect(docTagWarnings({ "names/MTG.xln": after }, ["MTG.REPAYMENT"]).map((x) => x.message)).toEqual([
      `MTG.REPAYMENT: the doc comment is 235 characters; with its provenance tag (29) it passes Excel's 255, so the build writes it without the tag and the workbook won't record its library base, @from(lib #${hash}): shorten it by 9 characters`,
    ]);
  });
});
