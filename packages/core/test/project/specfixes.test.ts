// The open issues of docs/LANGUAGE.md §14 that the checker now reports (fixed 2026-10-07):
// each one with its reproduction from the spec.
import { describe, expect, it } from "vitest";
import { buildWorkbook, checkFile, unbuiltEdits, convertModuleBlocks, definitionHash, parseModule, parseSourceFile, sourceFindings, spillMap, LOCK_FILE, parseLockfile, pullProject, SourceModel, type CheckContext, type PullResult } from "../../src/index.js";
import { fixture, withWorkbookXml, workbookXml } from "../build/helpers.js";

const F7 = fixture("f7_base.xlsx");
const U = "names/_unmanaged.xln";
const S1 = "names/sheets/S1.xln";

function model(files: Record<string, string>): SourceModel {
  const m = new SourceModel();
  for (const [p, t] of Object.entries(files)) if (p.startsWith("names/") && p.endsWith(".xln")) m.setFile(p, t);
  return m;
}

function ctxOf(r: PullResult): CheckContext {
  return { sheets: r.snapshot.sheets.map((s) => s.name), tables: r.snapshot.tables.map((t) => t.displayName), lock: parseLockfile(r.files[LOCK_FILE]!) };
}

/** A file's problems as [severity, code, the text they sit on, message]. */
function problems(files: Record<string, string>, path: string, ctx: CheckContext = {}): [string, string, string, string][] {
  const m = model(files);
  return checkFile(m, path, ctx).map((p) => [p.severity, p.code ?? "", m.files.get(path)!.text.slice(p.start, p.end), p.message]);
}

/** Applies a problem's quick fix to its file. */
function applyFix(files: Record<string, string>, path: string, code: string): string {
  const m = model(files);
  const p = checkFile(m, path).find((x) => x.code === code)!;
  const fix = p.fix ?? p.fixes![0]!;
  const t = files[path]!;
  return t.slice(0, fix.start) + fix.text + t.slice(fix.end);
}

describe("issue 14: two doc comments before one name", () => {
  it("is a warning on the first, with a fix removing it; doc comment and annotations in any order stay accepted", () => {
    const files = { "names/FN.xln": "/** First. */\n@hidden\n/** Second. */\nFN.X = 1;\n\n@hidden\n/** Only. */\nFN.Y = 2;\n" };
    expect(problems(files, "names/FN.xln")).toEqual([["warning", "doc-twice", "/** First. */", "FN.X: two doc comments: only the last is kept (it is the Name Manager comment); merge them, or remove this one"]]);
    expect(applyFix(files, "names/FN.xln", "doc-twice")).toBe("@hidden\n/** Second. */\nFN.X = 1;\n\n@hidden\n/** Only. */\nFN.Y = 2;\n");
    const m = model(files);
    expect(m.defs.map((d) => [d.name, d.entry.doc, d.entry.hidden])).toEqual([
      ["FN.X", "Second.", true],
      ["FN.Y", "Only.", true],
    ]);
  });
});

/** Applies a fix with its edits in other files. */
function applyMove(files: Record<string, string>, path: string, title: string, ctx: CheckContext = {}): Record<string, string> {
  const m = model(files);
  const fix = checkFile(m, path, ctx).flatMap((p) => p.fixes ?? []).find((f) => f.title === title)!;
  expect(fix, title).toBeDefined();
  const out = { ...files };
  out[path] = files[path]!.slice(0, fix.start) + fix.text + files[path]!.slice(fix.end);
  for (const e of fix.elsewhere ?? []) out[e.path] = e.create ? e.text : files[e.path]!.slice(0, e.start) + e.text + files[e.path]!.slice(e.end);
  return out;
}

describe("issue 2: a name without its module's prefix in a module file", () => {
  it("is a warning at the name, with fixes moving it to _unmanaged.xln or renaming it", () => {
    const r = pullProject(F7, "book.xlsx");
    const files = { ...r.files, "names/ANA.xln": "// module: ANA\n// @version 1.0\n\nANA.GROW = 2;\n\nRate3 = 0.3;\n" };
    expect(problems(files, "names/ANA.xln", ctxOf(r))).toEqual([["warning", "file-placement", "Rate3", "Rate3 has no ANA. prefix: a pull will put it in _unmanaged.xln; rename it ANA.Rate3 or move it there"]]);
    const moved = applyMove(files, "names/ANA.xln", "Move it to names/_unmanaged.xln");
    expect(moved["names/ANA.xln"]).toBe("// module: ANA\n// @version 1.0\n\nANA.GROW = 2;\n");
    expect(moved[U]!.endsWith("RateX = 0.5;\n\nRate3 = 0.3;\n")).toBe(true);
    expect(problems(moved, "names/ANA.xln", ctxOf(r))).toEqual([]);
    expect(problems(moved, U, ctxOf(r))).toEqual([]);
    const renamed = applyMove(files, "names/ANA.xln", "Rename it ANA.Rate3");
    expect(renamed["names/ANA.xln"]).toContain("\nANA.Rate3 = 0.3;\n");
  });

  it("a name the last pull had is renamed with @renamed; an AFE module's name pasted in is said to belong elsewhere", () => {
    const r = pullProject(F7, "book.xlsx");
    const files = { ...r.files, "names/ANA.xln": "// module: ANA\n\nANA.GROW = 2;\nAFE.Twice = 2;\n" };
    files[U] = files[U]!.replace("Loc = 3;\n", "");
    files["names/ANA.xln"] += "\n/** Where. */\nLoc = 3;\n";
    const ps = problems(files, "names/ANA.xln", ctxOf(r));
    expect(ps.map((p) => p[3])).toEqual([
      "AFE.Twice has the prefix of module AFE, not ANA: a pull will put it in names/AFE.xln; move it there",
      "Loc has no ANA. prefix: a pull will put it in _unmanaged.xln; rename it ANA.Loc or move it there",
    ]);
    const renamed = applyMove(files, "names/ANA.xln", "Rename it ANA.Loc", ctxOf(r));
    expect(renamed["names/ANA.xln"]).toContain("\n@renamed(Loc)\n/** Where. */\nANA.Loc = 3;\n");
    const moved = applyMove(files, "names/ANA.xln", "Move it to names/AFE.xln");
    expect(moved["names/AFE.xln"]).toBe("// module: AFE\n\nAFE.Twice = 2;\n");
  });

  it("a pulled project has none of these warnings", () => {
    const r = pullProject(fixture("traps.xlsx"), "traps.xlsx");
    const m = model(r.files);
    for (const path of m.files.keys()) expect(checkFile(m, path, ctxOf(r)).filter((p) => p.code === "file-placement" || p.code === "directive")).toEqual([]);
  });
});

describe("issue 4: @sheet and @scope in the wrong file kind", () => {
  it("@sheet(S) in a module file is per-name scope (issue 5); a sheet the workbook lacks is an error", () => {
    const files = { "names/FN.xln": "@sheet(S1)\nFN.X = 1;\n\n@sheet(Nope)\nFN.Y = 2;\n" };
    expect(problems(files, "names/FN.xln", { sheets: ["S1"] })).toEqual([["error", "unknown-sheet", "@sheet(Nope)", "@sheet(Nope): the workbook has no sheet 'Nope'"]]);
  });

  it("@scope in _unmanaged.xln is a warning naming it", () => {
    const files = { [U]: "A = 1;\n\n@scope(S1)\nB = 2;\n" };
    expect(problems(files, U).filter((p) => p[1] === "directive")).toEqual([["warning", "directive", "@scope(S1)", "@scope(S1) in _unmanaged.xln, which holds the workbook-scoped names no module owns: a pull puts names local to S1 in names/sheets/S1.xln (or in their module's file); move them there"]]);
  });

  it("@sheet(S) above a name in _unmanaged.xln: the pull puts it in the sheet's file", () => {
    const files = { [U]: "A = 1;\n\n@sheet(S1)\nB = 2;\n" };
    expect(problems(files, U)).toEqual([["warning", "file-placement", "B", "B (local to S1) has no module prefix: a pull puts a name local to S1 that no module owns in names/sheets/S1.xln; move it there"]]);
  });

  it("an unknown annotation lists the annotations of the file kind", () => {
    expect(problems({ "names/FN.xln": "@deprecated\nFN.X = 1;\n" }, "names/FN.xln")[0]![3]).toBe("@deprecated is not an annotation xln knows (@hidden, @sheet(Sheet), @renamed(Old), @from(lib #…)); it is ignored");
    expect(problems({ [S1]: "@deprecated\nX = 1;\n" }, S1)[0]![3]).toBe("@deprecated is not an annotation xln knows (@hidden, @workbook, @renamed(Old), @from(lib #…); and on a line of its own @sheet(Name)); it is ignored");
  });
});

describe("issue 5: per-name scope in module files (decided 2026-10-07)", () => {
  it("a module file with old blocks builds as it is, with an info and a fix converting it per name", () => {
    const r = pullProject(fixture("probe_win.xlsx"), "book.xlsx");
    const P = "names/P.xln";
    const fresh = r.files[P]!;
    // The same file in the old block form: the build reads it the same, so nothing to do.
    const old = fresh.replace("@sheet('S2')\nP_Local = 5;\n\n@sheet('S2')\nP_Local2 = 6;\n", "@scope(S2)\n\nP_Local = 5;\nP_Local2 = 6;\n");
    expect(old).not.toBe(fresh);
    const files = { ...r.files, [P]: old };
    const b = buildWorkbook({ workbook: fixture("probe_win.xlsx"), fileName: "book.xlsx", files }, { provenance: false, embed: false });
    expect(b.status, b.plan.problems.map((p) => p.message).join("\n")).toBe("up-to-date");
    const ps = checkFile(model(files), P, ctxOf(r)).filter((p) => p.code === "old-blocks");
    expect(ps.map((p) => p.severity)).toEqual(["info"]);
    expect(ps[0]!.fix!.title).toBe("Convert to per-name @sheet");
    expect(convertModuleBlocks(old)).toEqual({ text: fresh.replace("@sheet('S2')\nP_Local = 5;\n\n@sheet('S2')\nP_Local2 = 6;\n", "@sheet('S2')\nP_Local = 5;\n@sheet('S2')\nP_Local2 = 6;\n") });
  });

  it("a block of cells and @workbook back: converted with their sheets", () => {
    const text = "// m\nA = 1;\n\n@scope('Cash Flow')\n/** Doc. */\nB = 2;\n@C5 = 3;\n\n@workbook\nC = 4;\n";
    const conv = convertModuleBlocks(text);
    expect(conv).toEqual({ text: "// m\nA = 1;\n\n/** Doc. */\n@sheet('Cash Flow')\nB = 2;\n@'Cash Flow'!C5 = 3;\n\nC = 4;\n" });
    const pm = parseModule((conv as { text: string }).text);
    expect(pm.entries.map((e) => [e.name, e.scope ?? null])).toEqual([["A", null], ["B", "Cash Flow"], ["C", null]]);
    expect(pm.cells.map((c) => c.cellSheet)).toEqual(["Cash Flow"]);
  });

  it("@workbook in a module file without blocks is a hint with a fix removing it", () => {
    const files = { "names/FN.xln": "FN.A = 1;\n@workbook\nFN.B = 2;\n" };
    expect(problems(files, "names/FN.xln")).toEqual([["hint", "redundant-workbook", "@workbook", "@workbook in a module file says nothing: a module name is workbook-scoped unless @sheet(Sheet) is above it"]]);
    expect(applyFix(files, "names/FN.xln", "redundant-workbook")).toBe("FN.A = 1;\nFN.B = 2;\n");
  });
});

describe("issue 3: type declarations are not part of xln v1", () => {
  it("are an error on the type; pulled projects never have them", () => {
    const files = { "names/FN.xln": "FN.Periods : scalar = 5;\n" };
    expect(problems(files, "names/FN.xln")).toEqual([["error", "syntax", "s", "type declarations are not part of xln v1 (planned with the dimension layer)"]]);
    for (const name of ["probe_win.xlsx", "f7_base.xlsx", "traps.xlsx"]) {
      for (const [path, text] of Object.entries(pullProject(fixture(name), name).files)) if (path.endsWith(".xln")) expect(parseSourceFile(path, text).diagnostics, `${name} ${path}`).toEqual([]);
    }
  });
});

describe("issue 6: @workbook on a name that is not a cell, in a sheet file", () => {
  it("is a warning with fixes moving it to _unmanaged.xln (without @workbook) or making it local", () => {
    const r = pullProject(F7, "book.xlsx");
    const files = { ...r.files };
    files["names/sheets/S2.xln"] += "\n/** A helper. */\n@workbook\n@hidden\nHelper = 2;\n";
    const ps = problems(files, "names/sheets/S2.xln", ctxOf(r));
    expect(ps.filter((p) => p[1] === "file-placement")).toEqual([
      ["warning", "file-placement", "Helper", "Helper: a pull puts workbook names that are not on one sheet's cells in _unmanaged.xln (or their module): move it to names/_unmanaged.xln, or remove @workbook to make it local to S2"],
    ]);
    const moved = applyMove(files, "names/sheets/S2.xln", "Move it to names/_unmanaged.xln");
    expect(moved["names/sheets/S2.xln"]).toBe(r.files["names/sheets/S2.xln"]!);
    expect(moved[U]!.endsWith("RateX = 0.5;\n\n/** A helper. */\n@hidden\nHelper = 2;\n")).toBe(true);
    expect(problems(moved, U, ctxOf(r)).filter((p) => p[1] === "file-placement")).toEqual([]);
  });

  it("with no _unmanaged.xln yet, the fix creates it", () => {
    const files = { [S1]: "@workbook\nHelper = 2;\n" };
    const moved = applyMove(files, S1, "Move it to names/_unmanaged.xln");
    expect(moved[S1]).toBe("");
    expect(moved[U]).toBe("// Workbook-scoped names that no module owns.\n\nHelper = 2;\n");
  });
});

describe("issue 10: readers broken by a rename or re-scope in the source are errors in the editor, as in the build", () => {
  it("a rename: the unchanged readers are errors, saying why; the build refuses on the same", () => {
    const r = pullProject(F7, "book.xlsx");
    const files = { ...r.files };
    files[U] = files[U]!.replace("Fn = LAMBDA(x, x*Rate);", "@renamed(Fn)\nTimes = LAMBDA(x, x*Rate);");
    const ps = problems(files, S1, ctxOf(r)).filter((p) => p[2] === "Fn");
    expect(ps.map((p) => [p[0], p[1]])).toEqual([
      ["error", "in-use"],
      ["error", "in-use"],
    ]);
    expect(ps[0]![3]).toContain(": the source renames Fn to Times, so write Times here (xln rename, or Rename Symbol in the editor, renames every reader); the build refuses until then");
  });

  it("a re-scope: a workbook name made local to S2 is an error where S1 reads it", () => {
    const r = pullProject(F7, "book.xlsx");
    const files = { ...r.files };
    files[U] = files[U]!.replace("Rate2 = Rate*2;\n", "");
    files["names/sheets/S2.xln"] += "Rate2 = Rate*2;\n";
    const ps = problems(files, S1, ctxOf(r)).filter((p) => p[2] === "Rate2");
    expect(ps.map((p) => [p[0], p[1]])).toEqual([["error", "in-use"]]);
  });

  it("a deleted name: its unchanged readers are errors too", () => {
    const r = pullProject(F7, "book.xlsx");
    const files = { ...r.files };
    files[U] = files[U]!.replace("Loc = 3;\n", "");
    // `'S2'!Loc+Loc` on S1: the bare Loc read the workbook's Loc, which the source deleted.
    const ps = problems(files, S1, ctxOf(r)).filter((p) => p[2] === "Loc");
    expect(ps.map((p) => p[0])).toEqual(["error"]);
  });

  it("a formula already broken at the last pull stays a warning (the build leaves it)", () => {
    const r = pullProject(F7, "book.xlsx");
    const files = { ...r.files };
    files[S1] = files[S1]!.replace("@C8 = Rate*3;", "@C8 = Nope*3;");
    const ctx = ctxOf(r);
    ctx.lock!.cells!["S1!C8"]!.formula = definitionHash("Nope*3");
    const ps = problems(files, S1, ctx).filter((p) => p[2] === "Nope");
    expect(ps.map((p) => p[0])).toEqual(["warning"]);
    expect(ps[0]![3]).toContain("(as in the workbook since the last pull: the build leaves it; fix it here to have it written)");
  });
});

describe("issue 8: the spill warnings come from the shared checker", () => {
  it("a name without # on a spilling formula: the checker, the build and xln check say the same", () => {
    const r = pullProject(F7, "book.xlsx");
    const ctx = { ...ctxOf(r), spills: spillMap(r.snapshot.sheets) };
    const files = { ...r.files, [S1]: r.files[S1]!.replace("Spl @E1# =", "Spl @E1 =") };
    const ps = problems(files, S1, ctx).filter((p) => p[1].startsWith("spill"));
    expect(ps).toEqual([["warning", "spill-uncovered", "@E1", "Spl @E1: the formula spills over E1:E3, but Spl covers only E1 (write @E1# to name the spill)"]]);
    expect(problems(r.files, S1, ctx).filter((p) => p[1].startsWith("spill"))).toEqual([]);
    const b = buildWorkbook({ workbook: F7, fileName: "book.xlsx", files }, { provenance: false, embed: false });
    expect(b.plan.problems.filter((p) => p.code.startsWith("spill")).map((p) => p.message)).toEqual([ps[0]![3]]);
    expect(sourceFindings(files, ctx).filter((x) => x.code.startsWith("spill")).map((x) => x.message)).toEqual([ps[0]![3]]);
  });

  it("# on a cell left empty that the name is not on yet", () => {
    const files = { [S1]: "Slot @C6# = ;\n" };
    expect(problems(files, S1)).toEqual([["warning", "spill-empty", "@C6#", "Slot @C6#: '#' on a cell left empty is not written yet; S1!Slot goes on C6# when the cell gets a formula"]]);
  });
});

describe("issue 1: // @version comes back with a pull", () => {
  const MAC = fixture("probe_mac.xlsx");
  const tagged = (version: string) => {
    const files = { ...pullProject(MAC, "book.xlsx").files };
    files["names/P.xln"] = files["names/P.xln"]!.replace("// module: P, pulled by xln from book.xlsx.\n", `// module: P, pulled by xln from book.xlsx.\n// @version ${version}\n`);
    const r = buildWorkbook({ workbook: MAC, fileName: "book.xlsx", files }, { embed: false });
    expect(r.status).toBe("built");
    return r.bytes!;
  };

  it("pull writes the version the module's tags carry; the next build plans no provenance update", () => {
    const built = tagged("1.2");
    const back = pullProject(built, "book.xlsx");
    expect(back.files["names/P.xln"]!.startsWith("// module: P, pulled by xln from book.xlsx.\n// @version 1.2\n\n")).toBe(true);
    expect(back.files["names/Mod.xln"]).not.toContain("@version");
    expect(back.report.notes).toEqual([]);
    const again = buildWorkbook({ workbook: built, fileName: "book.xlsx", files: back.files }, { embed: false });
    expect(again.status, again.plan.changeSet.changes.map((c) => JSON.stringify(c)).join("\n")).toBe("up-to-date");
  });

  it("a clear majority is written, with a note naming the others; no majority, a note and no version", () => {
    const built = tagged("1.2");
    const xml = workbookXml(built);
    const one = withWorkbookXml(built, xml.replace("[xln P 1.2 #", "[xln P 1.1 #"));
    const r1 = pullProject(one, "book.xlsx");
    expect(r1.files["names/P.xln"]).toContain("// @version 1.2\n");
    expect(r1.report.notes).toHaveLength(1);
    expect(r1.report.notes[0]).toMatch(/^module P: wrote \/\/ @version 1\.2, which \d+ of its \d+ tagged names carry; \S+ 1\.1 \(the next build tags them 1\.2\)$/);
    const n = (xml.match(/\[xln P 1\.2 #/g) ?? []).length;
    let half = xml;
    for (let k = 0; k < Math.ceil(n / 2); k++) half = half.replace("[xln P 1.2 #", "[xln P 1.1 #");
    const r2 = pullProject(withWorkbookXml(built, half), "book.xlsx");
    expect(r2.files["names/P.xln"]).not.toContain("@version");
    expect(r2.report.notes[0]).toMatch(/^module P: no \/\/ @version written: its tagged names disagree/);
  });
});

describe("issue 9: after a browser build, the workbook replaced by <name>.xln.xlsx", () => {
  it("pull does not refuse: the source equals the workbook, whatever the old lockfile says", () => {
    const files = { ...pullProject(F7, "book.xlsx").files };
    files[S1] = files[S1]!.replace("@C8 = Rate*3;", "@C8 = Rate*4;");
    files[U] = files[U]!.replace("RateX = 0.5;", "/** Half. */\nRateX = 0.5;");
    // Not built yet: two edits.
    expect(unbuiltEdits({ workbook: F7, fileName: "book.xlsx", files }).map((e) => e.key)).toEqual(["RateX", "S1!C8"]);
    // A browser build writes book.xln.xlsx and leaves the lockfile; the author puts it in place.
    const r = buildWorkbook({ workbook: F7, fileName: "book.xlsx", files }, { embed: false });
    expect(r.status).toBe("built");
    expect(unbuiltEdits({ workbook: r.bytes!, fileName: "book.xlsx", files })).toEqual([]);
    // The build agrees: nothing to do.
    expect(buildWorkbook({ workbook: r.bytes!, fileName: "book.xlsx", files }, { embed: false }).status).toBe("up-to-date");
    // A further edit after that is unbuilt again.
    const more = { ...files, [S1]: files[S1]!.replace("@C8 = Rate*4;", "@C8 = Rate*5;") };
    expect(unbuiltEdits({ workbook: r.bytes!, fileName: "book.xlsx", files: more }).map((e) => e.key)).toEqual(["S1!C8"]);
  });
});
