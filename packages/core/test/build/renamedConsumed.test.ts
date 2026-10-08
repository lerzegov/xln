// A build consumes the `@renamed(…)` notes of the renames it makes (author's decision,
// 2026-10-07): the edits that remove them, only for the renames applied, and a project
// that reads as fully built afterwards (no unbuilt edit, a pull rewrites nothing).
import { describe, expect, it } from "vitest";
import { annotationRemoval, applySourceEdits, buildWorkbook, consumedRenamedEdits, LOCK_FILE, parseLockfile, pullProject, renameInProject, rewrittenFiles, sourceFindings, unbuiltEdits, type Change, type SourceRename } from "../../src/index.js";
import { build, edit, fixture, pulled, why } from "./helpers.js";

const F7 = fixture("f7_base.xlsx");
const U = "names/_unmanaged.xln";
const S1 = "names/sheets/S1.xln";
const S2 = "names/sheets/S2.xln";

function renamed(files: Record<string, string>, name: string, to: string): Record<string, string> {
  const r = renameInProject(files, name, to);
  if (typeof r === "string") throw new Error(r);
  return { ...files, ...r.files };
}

/** The project after a build that was written: its lockfile, manifest and the source without the spent notes. */
function afterBuild(files: Record<string, string>, r: ReturnType<typeof build>): Record<string, string> {
  return { ...files, ...(r.sourceFiles ?? {}), ...r.files };
}

function names(files: Record<string, string>): Record<string, string> {
  return Object.fromEntries(Object.entries(files).filter(([p]) => p.startsWith("names/")));
}

describe("annotationRemoval", () => {
  const at = (text: string, ann: string) => ({ offset: text.indexOf(ann), end: text.indexOf(ann) + ann.length });
  const cut = (text: string, ann: string) => {
    const r = annotationRemoval(text, at(text, ann));
    return text.slice(0, r.start) + text.slice(r.end);
  };
  it("the whole line when the annotation stands alone, indented or with CR LF", () => {
    expect(cut("A = 1;\n@renamed(Old)\nNew = 2;\n", "@renamed(Old)")).toBe("A = 1;\nNew = 2;\n");
    expect(cut("  @renamed(Old)  \n  New = 2;", "@renamed(Old)")).toBe("  New = 2;");
    expect(cut("A = 1;\r\n@renamed(Old)\r\nNew = 2;\r\n", "@renamed(Old)")).toBe("A = 1;\r\nNew = 2;\r\n");
  });
  it("on a shared line, the annotation and the blanks after it, or before it at the line's end", () => {
    expect(cut("@renamed(Old) @hidden\nNew = 2;", "@renamed(Old)")).toBe("@hidden\nNew = 2;");
    expect(cut("@hidden @renamed(Old)\nNew = 2;", "@renamed(Old)")).toBe("@hidden\nNew = 2;");
    expect(cut("@renamed(Old) New = 2;", "@renamed(Old)")).toBe("New = 2;");
    expect(cut("@hidden @renamed(Old)\r\nNew = 2;", "@renamed(Old)")).toBe("@hidden\r\nNew = 2;");
  });
});

describe("consumedRenamedEdits", () => {
  const rename = (from: string, to: string, scope: string | null = null): Change => ({ op: "rename-name", scope, from, to });

  it("removes only the annotations whose rename is in the change set", () => {
    const files = { [U]: "/** The pace. */\n@renamed(Rate)\nPace = 0.1;\n@renamed(Other)\nKept = 1;\n" };
    const e = consumedRenamedEdits(files, [rename("Rate", "Pace")]);
    expect(e.map((x) => ({ annotation: x.annotation, key: x.key, line: x.line, path: x.path }))).toEqual([{ annotation: "@renamed(Rate)", key: "Pace", line: 2, path: U }]);
    expect(applySourceEdits(files, e)).toEqual({ [U]: "/** The pace. */\nPace = 0.1;\n@renamed(Other)\nKept = 1;\n" });
    expect(consumedRenamedEdits(files, [])).toEqual([]);
    expect(consumedRenamedEdits(files, [rename("Rate", "Speed")])).toEqual([]);
  });

  it("the forms of LANGUAGE §9.6: the same scope, Sheet!Old, !Old, a scope change alone", () => {
    const files = {
      [S2]: "@renamed(Loc)\nSpot = 7;\n@renamed(!RateX)\nRateY = 0.5;\n@renamed(S1!Moved)\nMoved = 1;\n",
      [U]: "@renamed(S1!Tot)\nTotal = 1;\n",
    };
    const changes: Change[] = [
      rename("Loc", "Spot", "S2"),
      rename("RateX", "RateY", null),
      { op: "rescope-name", name: "RateY", from: null, to: "S2" },
      { op: "rescope-name", name: "Moved", from: "S1", to: "S2" },
      rename("Tot", "Total", "S1"),
      { op: "rescope-name", name: "Total", from: "S1", to: null },
    ];
    const e = consumedRenamedEdits(files, changes);
    expect(e.map((x) => `${x.path} ${x.annotation} ${x.key}`)).toEqual([`${U} @renamed(S1!Tot) Total`, `${S2} @renamed(Loc) S2!Spot`, `${S2} @renamed(!RateX) S2!RateY`, `${S2} @renamed(S1!Moved) S2!Moved`]);
    expect(applySourceEdits(files, e)).toEqual({ [S2]: "Spot = 7;\nRateY = 0.5;\nMoved = 1;\n", [U]: "Total = 1;\n" });
  });

  it("a change of spelling only (`@renamed(rate)` on Rate) is not a rename: it stays", () => {
    expect(consumedRenamedEdits({ [U]: "@renamed(rate)\nRate = 1;\n" }, [rename("x", "y")])).toEqual([]);
  });
});

describe("a build consumes the @renamed notes it applied (F7)", () => {
  it("rename → build: the notes go, nothing is left to build, a pull rewrites nothing", () => {
    let files = pulled(F7);
    files = renamed(files, "Rate", "Pace");
    files = renamed(files, "S2!Loc", "Spot");
    expect(files[U]).toContain("@renamed(Rate)\nPace = 0.1;");
    const r = build(F7, files);
    expect(r.status, why(r)).toBe("built");
    const lineOf = (path: string, ann: string) => files[path]!.split("\n").indexOf(ann) + 1;
    expect(r.renamedConsumed!.map((e) => `${e.path}:${e.line} ${e.annotation}`)).toEqual([`${U}:${lineOf(U, "@renamed(Rate)")} @renamed(Rate)`, `${S2}:${lineOf(S2, "@renamed(Loc)")} @renamed(Loc)`]);
    expect(Object.keys(r.sourceFiles!).sort()).toEqual([U, S2].sort());
    // The build's own files (lockfile, manifest) never hold the source.
    expect(Object.keys(r.files!).some((p) => p.startsWith("names/"))).toBe(false);

    const after = afterBuild(files, r);
    for (const t of Object.values(names(after))) expect(t).not.toContain("@renamed");
    // Fully built: no unbuilt edit, a second build has nothing to do, and the pull gives back the same files.
    expect(unbuiltEdits({ workbook: r.bytes!, fileName: "book.xlsx", files: after })).toEqual([]);
    expect(build(r.bytes!, after).status).toBe("up-to-date");
    const pull = pullProject(r.bytes!, "book.xlsx").files;
    expect(rewrittenFiles(after, pull, [])).toEqual([]);
    expect(names(pull)).toEqual(names(after));
  });

  it("renamed back after a build: X → Y, build, Y → X, build: no @renamed left", () => {
    const files0 = pulled(F7);
    const r1 = build(F7, renamed(files0, "Rate2", "RateTwo"));
    expect(r1.status, why(r1)).toBe("built");
    const p1 = afterBuild(renamed(files0, "Rate2", "RateTwo"), r1);
    expect(p1[U]).not.toContain("@renamed");
    const back = renameInProject(p1, "RateTwo", "Rate2") as SourceRename;
    expect(back.annotation).toBe("added");
    const files2 = { ...p1, ...back.files };
    expect(files2[U]).toContain("@renamed(RateTwo)\nRate2 = ");
    const r2 = build(r1.bytes!, files2);
    expect(r2.status, why(r2)).toBe("built");
    const p2 = afterBuild(files2, r2);
    expect(names(p2)).toEqual(names(files0));
    expect(unbuiltEdits({ workbook: r2.bytes!, fileName: "book.xlsx", files: p2 })).toEqual([]);
  });

  it("a refused build, a dry run and an up-to-date build leave the source alone", () => {
    const files = renamed(pulled(F7), "Rate", "Pace");
    expect(build(F7, files).renamedConsumed).toHaveLength(1);
    const refusing = { ...files };
    edit(refusing, S1, "@C8 = Pace*3;", "@C8 = Pace*(3;");
    const refused = build(F7, refusing);
    expect(refused.status).toBe("refused");
    expect(refused.renamedConsumed).toBeUndefined();
    expect(refused.sourceFiles).toBeUndefined();
    const dry = buildDry(files);
    expect(dry.status).toBe("planned");
    expect(dry.renamedConsumed).toBeUndefined();
    // A note an earlier build already applied: a build with nothing to write leaves it (the checker's hint and quick fix cover it).
    const r = build(F7, files);
    const stale = { ...files, ...r.files };
    const again = build(r.bytes!, stale);
    expect(again.status).toBe("up-to-date");
    expect(again.renamedConsumed).toBeUndefined();
    expect(again.sourceFiles).toBeUndefined();
  });

  it("the embedded source is the source after the removal", () => {
    const files = renamed(pulled(F7), "Rate", "Pace");
    const r = buildEmbed(files);
    expect(r.status, why(r)).toBe("built");
    const embedded = r.plan.changeSet.changes.find((c) => c.op === "set-embedded-source");
    expect(embedded && embedded.op === "set-embedded-source" ? embedded.files[U] : undefined).toBe(r.sourceFiles![U]);
  });
});

// A note whose rename is in the workbook already (an older build, a browser build's copy,
// another tool, a hand): the checker says it can go (`renamed-built`, with the quick fix),
// and the next build that writes the workbook removes it with its own (2026-10-07).
describe("a spent @renamed whose rename an earlier build made", () => {
  /** After `Rate → Pace` was built and its note kept: the source with the note, the lockfile with Pace. */
  function stale(): { bytes: Uint8Array; files: Record<string, string> } {
    const files = renamed(pulled(F7), "Rate", "Pace");
    const r = build(F7, files);
    expect(r.status, why(r)).toBe("built");
    return { bytes: r.bytes!, files: { ...files, ...r.files } };
  }
  const findings = (files: Record<string, string>) => sourceFindings(files, { lock: parseLockfile(files[LOCK_FILE]!) });

  it("the checker: a hint on the annotation, with the fix that removes its line", () => {
    const { files } = stale();
    const f = findings(files).filter((x) => x.code === "renamed-built");
    expect(f.map((x) => ({ severity: x.severity, message: x.message, file: x.file, key: x.key }))).toEqual([{ severity: "hint", message: "@renamed(Rate): the rename is built; this line can go", file: U, key: "Pace" }]);
    expect(f[0]!.line).toBe(files[U]!.split("\n").indexOf("@renamed(Rate)") + 1);
    const fix = f[0]!.fix!;
    expect(fix.title).toBe("Remove @renamed(Rate)");
    const text = files[U]!;
    expect(text.slice(0, fix.start) + fix.text + text.slice(fix.end)).toBe(text.replace("@renamed(Rate)\n", ""));
  });

  it("not on a pending rename, a change of spelling, an old name still in the workbook, or without a lockfile", () => {
    const pending = renamed(pulled(F7), "Rate", "Pace");
    expect(findings(pending).some((x) => x.code === "renamed-built")).toBe(false);
    const { files } = stale();
    const spelling = { ...files };
    edit(spelling, U, "@renamed(Rate)\nPace", "@renamed(pace)\nPace");
    expect(findings(spelling).some((x) => x.code === "renamed-built")).toBe(false);
    // Rate2 is still in the workbook: the note is not spent (the build reports it as it does).
    const still = { ...files };
    edit(still, U, "@renamed(Rate)\nPace", "@renamed(Rate2)\nPace");
    expect(findings(still).some((x) => x.code === "renamed-built")).toBe(false);
    expect(sourceFindings(files, {}).some((x) => x.code === "renamed-built")).toBe(false);
  });

  it("consumedRenamedEdits with the lockfile: the spent note, marked earlier; the others stay", () => {
    const { files } = stale();
    const lock = parseLockfile(files[LOCK_FILE]!);
    const e = consumedRenamedEdits(files, [], lock);
    expect(e.map((x) => `${x.annotation} ${x.key} ${x.earlier}`)).toEqual(["@renamed(Rate) Pace true"]);
    expect(consumedRenamedEdits({ [U]: "@renamed(rate)\nRate = 1;\n@renamed(Rate2)\nPace = 2;\n@renamed(Nope)\nNew = 3;\n" }, [], lock)).toEqual([]);
  });

  it("a writing build removes it with the notes of its own renames", () => {
    const { bytes, files } = stale();
    const next = renamed(files, "S2!Loc", "Spot");
    const r = build(bytes, next);
    expect(r.status, why(r)).toBe("built");
    expect(r.renamedConsumed!.map((x) => `${x.path} ${x.annotation} ${x.earlier ?? false}`)).toEqual([`${U} @renamed(Rate) true`, `${S2} @renamed(Loc) false`]);
    const after = afterBuild(next, r);
    for (const t of Object.values(names(after))) expect(t).not.toContain("@renamed");
    expect(build(r.bytes!, after).status).toBe("up-to-date");
    expect(names(pullProject(r.bytes!, "book.xlsx").files)).toEqual(names(after));
  });

  it("an up-to-date build and a dry run leave it; a forced build writes and removes it", () => {
    const { bytes, files } = stale();
    expect(build(bytes, files).renamedConsumed).toBeUndefined();
    const edited = { ...files };
    edit(edited, U, "Pace = 0.1;", "Pace = 0.2;");
    const dry = buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files: edited }, { dryRun: true, embed: false, provenance: false });
    expect(dry.status).toBe("planned");
    expect(dry.renamedConsumed).toBeUndefined();
    const forced = build(bytes, files, true);
    expect(forced.status, why(forced)).toBe("built");
    expect(forced.renamedConsumed!.map((x) => x.annotation)).toEqual(["@renamed(Rate)"]);
  });
});

function buildDry(files: Record<string, string>) {
  return buildWorkbook({ workbook: F7, fileName: "book.xlsx", files }, { dryRun: true, embed: false, provenance: false });
}
function buildEmbed(files: Record<string, string>) {
  return buildWorkbook({ workbook: F7, fileName: "book.xlsx", files }, { embed: true, provenance: false });
}
