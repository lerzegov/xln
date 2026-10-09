// D5 (embedded source) and D6 (provenance tags), end to end on the probe workbooks:
//   - a build embeds the project in one custom XML part; a second build finds it by its
//     namespace and never adds another; read-back covers the part;
//   - the part carries the project as built; pull never reads it (2026-10-06);
//   - module names carry `[xln Mod #hash]`, pull strips it, and a tag is never an edit.
import { readdirSync } from "node:fs";
import { describe, expect, it } from "vitest";
import {
  applyChangeSet,
  buildReportLines,
  buildWorkbook,
  describeChanges,
  provenanceOnly,
  embeddedSourceXml,
  LOCK_FILE,
  parseEmbeddedSource,
  planBuild,
  parseLockfile,
  pullProject,
  readBack,
  readEmbeddedSource,
  readWorkbook,
  sourceHash,
  splitProvenance,
  stripProvenance,
  withProvenance,
  formatProvenanceTag,
  moduleOfPath,
  moduleVersion,
  type BuildResult,
} from "../../src/index.js";
import { Package } from "../../src/file/package.js";
import { edit, fixture, RESULTS, withWorkbookXml, workbookXml } from "./helpers.js";

const NAMES = [...readdirSync(RESULTS).filter((f) => f.endsWith(".xlsx")).sort(), "traps.xlsx"];

function build(bytes: Uint8Array, files: Record<string, string>, opts: { force?: boolean } = {}): BuildResult {
  return buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files }, { embed: true, ...opts });
}

function why(r: BuildResult): string {
  return [...r.plan.problems.map((p) => `${p.severity} ${p.code}: ${p.message}`), ...r.plan.conflicts.map((c) => c.message), ...(r.readBack?.problems ?? []), r.error ?? ""].join("\n");
}

function customXmlItems(bytes: Uint8Array): string[] {
  return new Package(bytes).names.filter((n) => /^customXml\/item\d+\.xml$/i.test(n));
}

/** The project after a build: the source as given, with the files the build rewrote. */
function after(files: Record<string, string>, r: BuildResult): Record<string, string> {
  return { ...files, ...r.files };
}

/** Project files a restore must give back exactly. */
function sourceFiles(files: Record<string, string>): Record<string, string> {
  return Object.fromEntries(Object.entries(files).filter(([p]) => (p.startsWith("names/") && p.endsWith(".xln")) || p === LOCK_FILE));
}

describe("embedded source: format", () => {
  it("keeps CR LF, CR, XML specials, non-ASCII and control characters byte for byte", () => {
    const files = {
      "names/a.xln": "a = 1;\r\nb = \"<&>\";\r\n",
      "names/b.xln": "lone\rcr ]]> λ 𝔁 \t tab",
      "names/c.xln": "bell\u0007 and nul\u0000",
    };
    const back = parseEmbeddedSource(embeddedSourceXml(files))!;
    expect(back.damaged).toEqual([]);
    expect(back.files).toEqual(files);
  });

  it("survives Excel's save, which writes every LF in the item as CR LF (measured, F5)", () => {
    const files = { "names/a.xln": "a = 1;\nb = 2;\n", "names/w.xln": "w = 1;\r\n", "names/m.xln": "mixed\r\nand\n" };
    const xml = embeddedSourceXml(files);
    const excel = xml.split("\r\n").join("\n").split("\n").join("\r\n");
    expect(parseEmbeddedSource(excel)).toMatchObject({ files, damaged: [] });
  });

  it("reports a file whose checksum fails instead of restoring it", () => {
    const xml = embeddedSourceXml({ "names/a.xln": "a = 1;\n" }).replace("a = 1;", "a = 2;");
    expect(parseEmbeddedSource(xml)).toMatchObject({ files: {}, damaged: ["names/a.xln"] });
  });
});

describe("provenance tags: format", () => {
  it("writes, splits and strips the tag on the last line", () => {
    const t = { module: "FN", version: "1.2", hash: "3f9a1c" };
    expect(formatProvenanceTag(t)).toBe("[xln FN 1.2 #3f9a1c]");
    expect(withProvenance("Doubles.\nSecond line", t)).toBe("Doubles.\nSecond line [xln FN 1.2 #3f9a1c]");
    expect(withProvenance(undefined, { module: "FN", hash: "3f9a1c" })).toBe("[xln FN #3f9a1c]");
    expect(splitProvenance("Doubles. [xln FN 1.2 #3f9a1c]")).toEqual({ comment: "Doubles.", tag: t });
    expect(stripProvenance("[xln FN #3f9a1c]")).toBeUndefined();
    // Not tags: other brackets, a tag not at the end, a tag broken over lines.
    expect(stripProvenance("see [ref 2]")).toBe("see [ref 2]");
    expect(stripProvenance("[xln FN #3f9a1c] more")).toBe("[xln FN #3f9a1c] more");
    expect(stripProvenance("[xln FN\n#3f9a1c]")).toBe("[xln FN\n#3f9a1c]");
    expect(stripProvenance("[xln a b c #3f9a1c]")).toBe("[xln a b c #3f9a1c]");
  });

  it("module from the file, version from its header, hash insensitive to layout", () => {
    expect(moduleOfPath("names/FN.xln")).toBe("FN");
    expect(moduleOfPath("names/FN~2.xln")).toBe("FN");
    expect(moduleOfPath("names/_unmanaged.xln")).toBeUndefined();
    expect(moduleOfPath("names/sheets/S1.xln")).toBeUndefined();
    expect(moduleVersion("// module: FN\n// @version 1.2\n\nFN.x = 1;\n")).toBe("1.2");
    expect(moduleVersion("FN.x = 1;\n// @version 1.2\n")).toBeUndefined();
    expect(sourceHash("_xlfn.LAMBDA(_xlpm.x,\r\n _xlpm.x*2)", "Doc")).toBe(sourceHash("_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)", "Doc [xln FN #000000]"));
    expect(sourceHash("1", "a")).not.toBe(sourceHash("1", "b"));
  });
});

describe("embed and restore on the probe workbooks", () => {
  for (const name of NAMES) {
    it(name, () => {
      const bytes = fixture(name);
      const files = { ...pullProject(bytes, "book.xlsx").files };
      const r1 = build(bytes, files);
      expect(r1.status, why(r1)).toBe("built");
      expect(r1.readBack!.ok, why(r1)).toBe(true);
      expect(r1.plan.changeSet.changes.at(-1)!.op).toBe("set-embedded-source");
      const items = customXmlItems(r1.bytes!);
      expect(items.filter((i) => readEmbeddedSource(r1.bytes!)!.location.item === i)).toHaveLength(1);
      const project = after(files, r1);

      // Nothing changed: up to date, the part says exactly the project.
      const r2 = build(r1.bytes!, project);
      expect(r2.status, why(r2)).toBe("up-to-date");

      // The part carries the project; a pull of the built file gives the same (from the workbook).
      expect(sourceFiles(readEmbeddedSource(r1.bytes!)!.files)).toEqual(sourceFiles(project));
      const pulled = pullProject(r1.bytes!, "book.xlsx");
      expect(sourceFiles(pulled.files)).toEqual(sourceFiles(project));

      // An edit: the part is rewritten in place, never duplicated.
      const firstNames = Object.keys(project).find((p) => p.startsWith("names/") && p.endsWith(".xln"))!;
      const edited = { ...project, [firstNames]: project[firstNames] + "\n// a note the live names cannot carry\n" };
      const r3 = build(r1.bytes!, edited);
      expect(r3.status, why(r3)).toBe("built");
      expect(customXmlItems(r3.bytes!)).toEqual(items);
      expect(readEmbeddedSource(r3.bytes!)!.files[firstNames]).toBe(edited[firstNames]);
      // Pull never reads the part: the note is the part's only.
      expect(pullProject(r3.bytes!, "book.xlsx").files[firstNames]).toBe(project[firstNames]);
    });
  }
});

describe("the part is an archive copy: pull reads the workbook (2026-10-06)", () => {
  it("a cell filled in Excel after the build comes in; the part keeps the built source", () => {
    const bytes = fixture("f7_base.xlsx");
    const files = { ...pullProject(bytes, "book.xlsx").files };
    edit(files, "names/sheets/S1.xln", "@C8 = Rate*3;", "@C8 = ;");
    const r = build(bytes, files);
    expect(r.status, why(r)).toBe("built");
    expect(parseLockfile(r.files![LOCK_FILE]!).cells!["S1!C8"]!.formula).toBeNull();
    const filled = applyChangeSet(r.bytes!, [{ op: "set-cell-formula", sheet: "S1", range: "C8", stored: "Rate*5", display: "Rate*5" }]);
    expect(pullProject(filled, "book.xlsx").files["names/sheets/S1.xln"]).toContain("@C8 = Rate*5;");
    expect(readEmbeddedSource(filled)!.files["names/sheets/S1.xln"]).toContain("@C8 = ;");
  });

  it("a name changed in Excel after the build: the pull writes Excel's definition (the part's comments are not read back)", () => {
    const bytes = fixture("probe_mac.xlsx");
    const files = { ...pullProject(bytes, "book.xlsx").files };
    edit(files, "names/P.xln", "P_Add1 = LAMBDA(x, x+1);", "// adds one\nP_Add1 = LAMBDA(x, x+1);");
    const r1 = build(bytes, files);
    expect(r1.status, why(r1)).toBe("built");
    // "Excel" edits P_Add1.
    const xml = workbookXml(r1.bytes!);
    const excel = withWorkbookXml(r1.bytes!, xml.replace("_xlpm.x+1)</definedName>", "_xlpm.x+2)</definedName>"));
    const p = pullProject(excel, "book.xlsx");
    expect(p.files["names/P.xln"]).toContain("\nP_Add1 = LAMBDA(x, x+2);");
    expect(p.files["names/P.xln"]).not.toContain("// adds one");
    expect(readEmbeddedSource(excel)!.files["names/P.xln"]).toContain("// adds one\nP_Add1 = LAMBDA(x, x+1);");
    // The tag's hash no longer matches: edited in Excel.
    expect(p.provenance.find((x) => x.key === "P_Add1")).toMatchObject({ state: "edited" });
    expect(p.provenance.find((x) => x.key === "Mod.Fn")).toMatchObject({ state: "unchanged" });
    // The pulled project agrees with the workbook: nothing from Excel, no conflict (the build re-tags P_Add1).
    const r2 = build(excel, { ...p.files });
    expect(r2.plan.excelChanges).toEqual([]);
    expect(r2.plan.conflicts).toEqual([]);
  });
});

describe("provenance tags (D6)", () => {
  const bytes = fixture("probe_mac.xlsx");
  const pulled = () => ({ ...pullProject(bytes, "book.xlsx").files });

  it("module names get the tag, others do not; pull strips it; the lockfile ignores it", () => {
    const files = pulled();
    edit(files, "names/Mod.xln", "// module: Mod, pulled by xln from book.xlsx.\n", "// module: Mod, pulled by xln from book.xlsx.\n// @version 1.2\n");
    edit(files, "names/Mod.xln", "Mod.Fn = ", "/** Times ten. */\nMod.Fn = ");
    const r = build(bytes, files);
    expect(r.status, why(r)).toBe("built");
    const names = readWorkbook(r.bytes!).definedNames;
    const fn = names.find((d) => d.name === "Mod.Fn")!;
    expect(fn.comment).toBe(`Times ten. [xln Mod 1.2 #${sourceHash(fn.definition, "Times ten.")}]`);
    expect(names.find((d) => d.name === "P_Add1")!.comment).toMatch(/^\[xln P #[0-9a-f]{6}\]$/);
    expect(names.find((d) => d.name === "Fact")!.comment).toBeUndefined();
    const back = pullProject(r.bytes!, "book.xlsx");
    expect(back.files["names/Mod.xln"]).toContain("/** Times ten. */\nMod.Fn = ");
    expect(back.files["names/Mod.xln"]).not.toContain("[xln");
    const lock = parseLockfile(r.files![LOCK_FILE]!);
    expect(lock.names["Mod.Fn"]).toEqual(parseLockfile(back.files[LOCK_FILE]!).names["Mod.Fn"]);
    expect(build(r.bytes!, after(files, r)).status).toBe("up-to-date");
  });

  it("the text report folds updates of the tag alone into one line; the change set keeps each (feedback 2026-10-08)", () => {
    const files = pulled();
    edit(files, "names/_unmanaged.xln", "Fact = LAMBDA(n, 1);", "Fact = LAMBDA(n, 2);");
    const r = build(bytes, files);
    expect(r.status, why(r)).toBe("built");
    const tagOnly = r.plan.changeSet.changes.filter(provenanceOnly);
    expect(tagOnly.length).toBeGreaterThan(1);
    const lines = buildReportLines(r);
    expect(lines.filter((l) => l.includes("(provenance)"))).toEqual([]);
    expect(lines).toContain("  update Fact (definition)");
    expect(lines).toContain(`  update the provenance tag of ${tagOnly.length} module names (comment only)`);
    expect(describeChanges(r.plan.changeSet.changes)).toHaveLength(r.plan.changeSet.changes.length - tagOnly.length + 1);
  });

  it("a tag changed or removed in Excel is not an edit: no conflict, no Excel change, the build puts it back", () => {
    const files = pulled();
    const r1 = build(bytes, files);
    const project = after(files, r1);
    const xml = workbookXml(r1.bytes!);
    const fn = readWorkbook(r1.bytes!).definedNames.find((d) => d.name === "Mod.Fn")!;
    const tampered = withWorkbookXml(r1.bytes!, xml.replace(` comment="${fn.comment}"`, ""));
    const plan = planBuild({ workbook: readWorkbook(tampered), fileName: "book.xlsx", files: project, lock: parseLockfile(project[LOCK_FILE]!) });
    expect(plan.conflicts).toEqual([]);
    expect(plan.excelChanges).toEqual([]);
    const r2 = build(tampered, project);
    expect(r2.status, why(r2)).toBe("built");
    expect(r2.plan.changeSet.changes.filter((c) => c.op === "set-name")).toEqual([expect.objectContaining({ name: "Mod.Fn", fields: ["provenance"], comment: fn.comment })]);
    expect(readWorkbook(r2.bytes!).definedNames.find((d) => d.name === "Mod.Fn")!.comment).toBe(fn.comment);
  });

  it("an edit in source changes the hash; a comment too long for its tag is written without it, with a warning", () => {
    const files = pulled();
    edit(files, "names/Mod.xln", "Mod.Fn = LAMBDA(x, x*10);", "Mod.Fn = LAMBDA(x, x*11);");
    edit(files, "names/P.xln", "P_Add1 = ", `/** ${"a".repeat(240)} */\nP_Add1 = `);
    const r = build(bytes, files);
    expect(r.status, why(r)).toBe("built");
    const names = readWorkbook(r.bytes!).definedNames;
    const fn = names.find((d) => d.name === "Mod.Fn")!;
    expect(splitProvenance(fn.comment!).tag!.hash).toBe(sourceHash(fn.definition, undefined));
    expect(names.find((d) => d.name === "P_Add1")!.comment).toBe("a".repeat(240));
    expect(r.plan.problems).toEqual([expect.objectContaining({ severity: "warning", code: "provenance", key: "P_Add1" })]);
  });

  it("a comment over 255 characters refuses the build: Excel would not open the file (F5)", () => {
    const files = pulled();
    edit(files, "names/_unmanaged.xln", "Fact = ", `/** ${"b".repeat(256)} */\nFact = `);
    const r = build(bytes, files);
    expect(r.status).toBe("refused");
    expect(r.plan.problems).toEqual([expect.objectContaining({ severity: "error", code: "comment-length", key: "Fact" })]);
  });

  it("a multi-line comment as Excel saves it (_x000a_) reads as the source's", () => {
    const files = pulled();
    edit(files, "names/Mod.xln", "Mod.Fn = ", "/**\n * Line one.\n * Line_x0041_ two.\n */\nMod.Fn = ");
    const r = build(bytes, files);
    expect(r.status, why(r)).toBe("built");
    const xml = workbookXml(r.bytes!);
    expect(xml).toContain('comment="Line one.&#10;Line_x005F_x0041_ two. [xln Mod #');
    const excel = withWorkbookXml(r.bytes!, xml.replace("Line one.&#10;", "Line one._x000a_"));
    expect(readWorkbook(excel).definedNames.find((d) => d.name === "Mod.Fn")!.comment).toMatch(/^Line one\.\nLine_x0041_ two\. \[xln Mod #/);
    const again = build(excel, after(files, r));
    expect(again.status, why(again)).toBe("up-to-date");
  });

  it("--no-tags and --no-embed builds leave comments and the package alone", () => {
    const r = buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files: pulled() }, { embed: false, provenance: false });
    expect(r.status).toBe("up-to-date");
  });

  it("read-back catches a part that does not say the project", () => {
    const files = pulled();
    const r = build(bytes, files);
    const changes = r.plan.changeSet.changes.map((c) => (c.op === "set-embedded-source" ? { ...c, files: { ...c.files, "names/extra.xln": "x = 1;\n" } } : c));
    expect(readBack(bytes, r.bytes!, changes, r.plan.inSync).problems).toContain("the embedded source differs from the project files it should carry");
  });
});
