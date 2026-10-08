// M4 in the editor's model: library functions in completion, the library's location, and
// the states the code lenses show.
import { applyEdits, definitionBase, libraryFunctionBase, readLibrary } from "@xln/core";
import { describe, expect, it } from "vitest";
import { completions } from "../src/model/editor.js";
import { entryArg, entryLibraryActions, entryStates, libraryBaseHover, libraryCompletions, libraryLocation, projectLibraryStatus, publishWarning, tagWarning } from "../src/model/library.js";
import { Project } from "../src/model/project.js";

const lambda = (name: string, summary: string, params: string, def: string) => `# name       ${name}\n# summary    ${summary}\n# params     ${params}\n\n${def}\n`;

const LIB = readLibrary({
  "FN.TWICE.lambda": lambda("FN.TWICE", "Doubles.", "x", "LAMBDA(x, 2 * x)"),
  "FN.QUAD.lambda": lambda("FN.QUAD", "Quadruples.", "x", "LAMBDA(x, FN.TWICE(FN.TWICE(x)))"),
  "FN.OWN.lambda": lambda("FN.OWN", "The library's.", "x", "LAMBDA(x, x - 2)"),
  "CHK.OK.lambda": lambda("CHK.OK", "Checks.", "x", "LAMBDA(x, x = 0)"),
});

function project(): Project {
  const p = new Project("mem:/p", undefined);
  p.setFile("names/FN.xln", "FN.OWN = LAMBDA(x, x - 1);\n");
  p.setFile("names/_unmanaged.xln", "Rate = 0.1;\nUse = FN.;\nBare = C;\n");
  return p;
}

describe("completion from the library", () => {
  it("after a module prefix: the library's functions the project lacks, marked", () => {
    const p = project();
    const text = p.files.get("names/_unmanaged.xln")!.text;
    const at = text.indexOf("FN.;") + 3;
    const items = completions(p, "names/_unmanaged.xln", at, LIB)!;
    const lib = items.filter((c) => c.library);
    expect(lib.map((c) => c.label).sort()).toEqual(["FN.QUAD", "FN.TWICE"]);
    const quad = lib.find((c) => c.label === "FN.QUAD")!;
    expect(quad.detail).toBe("from library · FN.QUAD(x)");
    expect(quad.library).toEqual({ name: "FN.QUAD", adds: ["FN.TWICE", "FN.QUAD"] });
    expect(quad.documentation).toContain("(and FN.TWICE, which it calls) to `names/FN.xln`");
    // The project's own FN.OWN is offered as the project's, not the library's.
    expect(items.filter((c) => c.label === "FN.OWN").map((c) => c.library)).toEqual([undefined]);
    // Without a library, nothing from it.
    expect(completions(p, "names/_unmanaged.xln", at)!.some((c) => c.library)).toBe(false);
  });

  it("a bare word: a module only the library has is offered as a prefix", () => {
    const p = project();
    const text = p.files.get("names/_unmanaged.xln")!.text;
    const at = text.indexOf("= C;") + 3;
    const mods = completions(p, "names/_unmanaged.xln", at, LIB)!.filter((c) => c.kind === "module").map((c) => `${c.label} ${c.detail}`);
    expect(mods).toEqual(["FN. module · 1 name", "CHK. module · from library"]);
    expect(libraryCompletions(p, LIB, "chk.").map((c) => c.fn.name)).toEqual(["CHK.OK"]);
  });
});

describe("library location", () => {
  it("relative, absolute, ~ (desktop only)", () => {
    expect(libraryLocation("../../_shared/lib", undefined)).toEqual({ relative: "../../_shared/lib" });
    expect(libraryLocation("/Users/a/lib", undefined)).toEqual({ absolute: "/Users/a/lib" });
    expect(libraryLocation("C:\\models\\lib", undefined)).toEqual({ absolute: "/C:/models/lib" });
    expect(libraryLocation("~/lib", "/Users/a")).toEqual({ absolute: "/Users/a/lib" });
    expect(libraryLocation("~/lib", undefined)).toMatchObject({ error: expect.stringContaining("cannot be expanded") });
  });
});

describe("entry states for the code lenses", () => {
  it("by file and name", () => {
    const p = project();
    const r = projectLibraryStatus(p, LIB, undefined, { target: "p", library: "lib" });
    const s = entryStates(r, "names/FN.xln");
    expect([...s.entries()].map(([k, i]) => `${k} ${i.state}`)).toEqual(["fn.own differs"]);
    expect(r.items.filter((i) => i.state === "missing").map((i) => i.name)).toEqual(["CHK.OK", "FN.QUAD", "FN.TWICE"]);
  });
});

describe("code lenses and quick fixes by state, three-way on the library base", () => {
  const base = (def: string) => definitionBase(def, "FN.X");
  const LIB2 = readLibrary({
    "FN.SAME.lambda": lambda("FN.SAME", "Same.", "x", "LAMBDA(x, x * 2)"),
    "FN.OUT.lambda": lambda("FN.OUT", "Out.", "x", "LAMBDA(x, x + 2)"),
    "FN.MOD.lambda": lambda("FN.MOD", "Mod.", "x", "LAMBDA(x, x + 3)"),
    "FN.BOTH.lambda": lambda("FN.BOTH", "Both.", "x", "LAMBDA(x, x + 30)"),
    "FN.DIFF.lambda": lambda("FN.DIFF", "Diff.", "x", "LAMBDA(x, x / 2)"),
  });
  const p = new Project("mem:/p", undefined);
  p.setFile(
    "names/FN.xln",
    [
      "FN.SAME = LAMBDA(x, x * 2);",
      `@from(lib #${base("LAMBDA(x, x + 1)")})`,
      "FN.OUT = LAMBDA(x, x + 1);",
      `@from(lib #${base("LAMBDA(x, x + 3)")})`,
      "FN.MOD = LAMBDA(x, x + 4);",
      `@from(lib #${base("LAMBDA(x, x + 10)")})`,
      "FN.BOTH = LAMBDA(x, x + 20);",
      "FN.DIFF = LAMBDA(x, x / 3);",
      "FN.LOCAL = LAMBDA(x, x);",
      "",
    ].join("\n"),
  );
  const items = entryStates(projectLibraryStatus(p, LIB2, undefined, { target: "p", library: "lib" }), "names/FN.xln");
  const actions = (name: string) => entryLibraryActions(items.get(name.toLowerCase())!);

  it("the five states and local only", () => {
    expect([...items.values()].map((i) => `${i.name} ${i.state}`)).toEqual(["FN.BOTH both-changed", "FN.DIFF differs", "FN.MOD modified", "FN.OUT outdated", "FN.SAME identical", "FN.LOCAL local-only"]);
  });

  it("identical without a base: the label, and Record library base (offered, never done)", () => {
    expect(actions("FN.SAME")).toEqual({ label: "library: identical · no base", tooltip: expect.stringContaining("no base recorded: Record library base"), diff: false, record: { title: "Record library base" } });
  });

  it("identical with a base: the label only", () => {
    const q = new Project("mem:/q", undefined);
    q.setFile("names/FN.xln", `@from(lib #${base("LAMBDA(x, x * 2)")})\nFN.SAME = LAMBDA(x, x * 2);\n`);
    const i = entryStates(projectLibraryStatus(q, LIB2, undefined, { target: "q", library: "lib" }), "names/FN.xln").get("fn.same")!;
    expect(entryLibraryActions(i)).toEqual({ label: "library: identical", tooltip: expect.stringContaining("same definition"), diff: false });
  });

  it("outdated: Take (no question), no Publish (it would put the older version back)", () => {
    const a = actions("FN.OUT");
    expect(a).toMatchObject({ label: "library: outdated", diff: true, take: { title: "Take the library's version" } });
    expect(a.take!.confirm).toBeUndefined();
    expect(a.publish).toBeUndefined();
    expect(a.tooltip).toContain(`(base #${base("LAMBDA(x, x + 1)")}, library #${base("LAMBDA(x, x + 2)")})`);
    expect(publishWarning(items.get("fn.out"))).toContain("publishing puts the older definition back");
  });

  it("modified: Publish, and Take to undo after a question naming the edit", () => {
    const a = actions("FN.MOD");
    expect(a).toMatchObject({ label: "library: modified", diff: true, publish: { title: "Publish to library" } });
    expect(a.take!.title).toBe("Take the library's version (undo the edit)");
    expect(a.take!.confirm).toMatch(/^Discard your edit of FN\.MOD\? /);
    expect(publishWarning(items.get("fn.mod"))).toBeUndefined();
  });

  it("both changed: Show diff, Take only after a question, no one-click Publish", () => {
    const a = actions("FN.BOTH");
    expect(a).toMatchObject({ label: "library: both changed", diff: true });
    expect(a.take!.confirm).toContain("Discard your edit of FN.BOTH? The library changed it too");
    expect(a.publish).toBeUndefined();
    expect(publishWarning(items.get("fn.both"))).toContain("publishing replaces the library's change");
  });

  it("differs (no base): Take after a question (a local edit cannot be ruled out), and Publish", () => {
    const a = actions("FN.DIFF");
    expect(a).toMatchObject({ label: "library: differs", diff: true, take: { title: "Take the library's version" }, publish: { title: "Publish to library" } });
    expect(a.take!.confirm).toBe("FN.DIFF has no library base: taking the library's version replaces whatever this copy has. Continue?");
    expect(actions("FN.LOCAL")).toMatchObject({ label: "library: not in the library", diff: false, publish: { title: "Publish to library" } });
  });

  it("the hover of @from says which library version", () => {
    const text = p.files.get("names/FN.xln")!.text;
    const at = text.indexOf("@from") + 3;
    expect(libraryBaseHover(p, "names/FN.xln", at)).toContain(`from the library version #${base("LAMBDA(x, x + 1)")}, the library definition FN.OUT came from`);
    expect(libraryBaseHover(p, "names/FN.xln", text.indexOf("FN.SAME"))).toBeUndefined();
    const q = new Project("mem:/q", undefined);
    q.setFile("names/FN.xln", "@from(lib 12)\nFN.X = LAMBDA(x, x);\n");
    expect(libraryBaseHover(q, "names/FN.xln", 2)).toContain("write @from(lib #abc123)");
  });
});

describe("the kept library bases in the editor's status", () => {
  it("both changed: base → copy and base → library from library-bases/, passed in apart", () => {
    const old = readLibrary({ "FN.X.lambda": lambda("FN.X", "X.", "x", "LAMBDA(x, x + 1)") });
    const now = readLibrary({ "FN.X.lambda": lambda("FN.X", "X.", "x", "LAMBDA(x, x + 2)") });
    const b = libraryFunctionBase(old.get("FN.X")!);
    const p = new Project("mem:/p", undefined);
    p.setFile("names/FN.xln", `@from(lib #${b.hash})\nFN.X = LAMBDA(x, x + 3);\n`);
    const without = projectLibraryStatus(p, now, undefined, { target: "p", library: "lib" }).items[0]!;
    expect(without.state).toBe("both-changed");
    expect(without.baseDiffs).toBeUndefined();
    const kept = projectLibraryStatus(p, now, undefined, { target: "p", library: "lib", bases: new Map([[b.hash, b]]) }).items[0]!;
    expect(kept.baseDisplay).toBe("LAMBDA(x, x + 1)");
    expect(kept.baseSource).toBe(`library-bases/${b.hash}.json`);
  });
});

describe("@param drift, live (the shared checker)", () => {
  it("a stale @param: a warning with the rename and remove fixes; a parameter left out: a hint", () => {
    const p = new Project("mem:/p", undefined);
    const text = "/**\n * Spreads.\n * @param total the amount\n * @param periods how many\n */\nFN.SPREAD = LAMBDA(total, periodi, total / periodi);\n";
    p.setFile("names/FN.xln", text);
    const probs = p.problems("names/FN.xln").filter((x) => x.code?.startsWith("doc-param"));
    expect(probs.map((x) => [x.severity, text.slice(x.start, x.end)])).toEqual([["warning", "periods"]]);
    expect(probs[0]!.fixes!.map((f) => f.title)).toEqual(["Rename @param periods to periodi", "Remove @param periods"]);
    p.setFile("names/FN.xln", applyEdits(text, [probs[0]!.fixes![1]!]));
    expect(p.problems("names/FN.xln").filter((x) => x.code?.startsWith("doc-param")).map((x) => [x.severity, x.code])).toEqual([["hint", "doc-param-missing"]]);
  });
});

describe("library commands from a menu", () => {
  it("a menu's Uri is not an entry: the command then works on the cursor", () => {
    // VS Code passes the document's Uri to editor/context commands (the author's Publish did nothing).
    expect(entryArg({ scheme: "file", path: "/p/names/FN.xln", fsPath: "/p/names/FN.xln" })).toBeUndefined();
    expect(entryArg(undefined)).toBeUndefined();
    expect(entryArg({ root: "file:///p", name: "FN.GROW" })).toEqual({ root: "file:///p", name: "FN.GROW" });
  });
});

describe("room for the provenance tag (feedback 2026-10-07)", () => {
  const doc = "x".repeat(235);

  it("live: the shared checker's warning on the doc comment", () => {
    const p = new Project("mem:/p", undefined);
    const text = `/** ${doc} */\n@from(lib #353921)\nMTG.REPAYMENT = LAMBDA(x, x);\n`;
    p.setFile("names/MTG.xln", text);
    const probs = p.problems("names/MTG.xln").filter((x) => x.code === "provenance");
    expect(probs.map((x) => [x.severity, text.slice(x.start, x.end)])).toEqual([["warning", `/** ${doc} */`]]);
    expect(probs[0]!.message).toContain("with its provenance tag (29) it passes Excel's 255");
  });

  it("Publish, Take, Insert, Record base: the warning to show before, on the files as the edit leaves them", () => {
    const text = `/** ${doc} */\nMTG.REPAYMENT = LAMBDA(x, x);\n`;
    const files = { "names/MTG.xln": text };
    expect(tagWarning(files, {}, ["MTG.REPAYMENT"])).toBeUndefined();
    const after = text.replace("MTG.REPAYMENT =", "@from(lib #353921)\nMTG.REPAYMENT =");
    expect(tagWarning(files, { "names/MTG.xln": after }, ["MTG.REPAYMENT"])).toBe(
      "MTG.REPAYMENT: the doc comment is 235 characters; with its provenance tag (29) it passes Excel's 255, so the build writes it without the tag and the workbook won't record its library base, @from(lib #353921): shorten it by 9 characters",
    );
    expect(tagWarning(files, { "names/MTG.xln": after }, ["MTG.OTHER"])).toBeUndefined();
  });
});
