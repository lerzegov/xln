// M3d: sheet files with a per-name `@workbook` (no blocks), the old block form still read
// and converted only on request, scope (no setting since 2026-10-06), and the one checker
// the editor and the build share (check.ts).
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import {
  applyChangeSet,
  buildWorkbook,
  checkFile,
  convertSheetBlocks,
  definitionHash,
  defaultConfigText,
  NAMES_SCOPE_NOTE,
  parseConfig,
  LOCK_FILE,
  parseLockfile,
  parseModule,
  parseSourceFile,
  unbuiltEdits,
  pullProject,
  SourceModel,
  type BuildResult,
  type CheckContext,
  type PullResult,
} from "../../src/index.js";
import { edit, fixture } from "../build/helpers.js";

const F7 = fixture("f7_base.xlsx");
const S1 = "names/sheets/S1.xln";
const S2 = "names/sheets/S2.xln";

function project(bytes: Uint8Array): PullResult {
  return pullProject(bytes, "book.xlsx");
}

function build(bytes: Uint8Array, files: Record<string, string>): BuildResult {
  return buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files }, { provenance: false, embed: false });
}

function why(r: BuildResult): string {
  return [...r.plan.problems.map((p) => `${p.severity} ${p.code}: ${p.message}`), ...r.plan.conflicts.map((c) => c.message)].join("\n");
}

function model(files: Record<string, string>): SourceModel {
  const m = new SourceModel();
  for (const [p, t] of Object.entries(files)) if (p.startsWith("names/") && p.endsWith(".xln")) m.setFile(p, t);
  return m;
}

function ctxOf(r: PullResult, files: Record<string, string>): CheckContext {
  return { sheets: r.snapshot.sheets.map((s) => s.name), tables: r.snapshot.tables.map((t) => t.displayName), lock: parseLockfile(files[LOCK_FILE]!) };
}

/** A file's problems as [severity, code, the text they sit on, message]. */
function problems(files: Record<string, string>, path: string, ctx: CheckContext): [string, string, string, string][] {
  const m = model(files);
  return checkFile(m, path, ctx).map((p) => [p.severity, p.code ?? "", m.files.get(path)!.text.slice(p.start, p.end), p.message]);
}

describe("the parser: sheet files and their annotations", () => {
  it("a sheet file: names local to its sheet unless @workbook is above them; @workbook stacks with a doc comment and @hidden", () => {
    const text = ["/** the years */", "@workbook", "@hidden", "years @C2 = SEQUENCE(1, 5, 2025);", "Sales @C3# = 1;", "@C4 = years;", "Other = 2;"].join("\n");
    const m = parseSourceFile("names/sheets/IS.xln", text);
    expect(m.form).toBe("sheet");
    expect(m.sheet).toBe("IS");
    expect(m.scopes).toEqual([]);
    expect(m.diagnostics).toEqual([]);
    expect(m.entries.map((e) => [e.name, e.scope, e.workbook ?? false, e.hidden, e.doc, e.cellSheet])).toEqual([
      ["years", undefined, true, true, "the years", "IS"],
      ["Sales", "IS", false, false, undefined, "IS"],
      ["Other", "IS", false, false, undefined, undefined],
    ]);
    expect(m.cells.map((e) => [e.name, e.cellSheet])).toEqual([["years", "IS"], ["Sales", "IS"], ["", "IS"]]);
    // The entry starts at its doc comment: removing it takes the annotations with it.
    expect(text.slice(m.entries[0]!.start, m.entries[0]!.start + 3)).toBe("/**");
  });

  it("refuses @workbook with an argument and above an unnamed cell", () => {
    expect(parseSourceFile("names/sheets/IS.xln", "@workbook(IS)\nX @C2 = 1;").diagnostics.map((d) => d.message)).toEqual(["@workbook takes no argument: write @workbook on the line above the name"]);
    expect(parseSourceFile("names/sheets/IS.xln", "@workbook\n@C2 = 1;").diagnostics.map((d) => d.message)).toEqual(["@workbook scopes a name, and @C2 has none: remove the @workbook line above it"]);
  });

  it("reads an old sheet file (blocks) as it was; a module file keeps its blocks", () => {
    const old = "@scope(IS)\n\n@workbook\n\nyears @IS!C2 = 1;\n\n@scope(IS)\n\nSales @C3# = 2;\n";
    const m = parseSourceFile("names/sheets/IS.xln", old);
    expect(m.form).toBe("sheet-blocks");
    expect(m.entries.map((e) => [e.name, e.scope, e.cellSheet])).toEqual([["years", undefined, "IS"], ["Sales", "IS", "IS"]]);
    const mod = parseSourceFile("names/FN.xln", "FN.A = 1;\n@scope(IS)\nFN.B = 2;\n@workbook\nFN.C = 3;\n");
    expect(mod.form).toBe("module");
    expect(mod.entries.map((e) => [e.name, e.scope])).toEqual([["FN.A", undefined], ["FN.B", "IS"], ["FN.C", undefined]]);
    // Without a path, a text is a module.
    expect(parseModule("@workbook\nX = 1;").form).toBe("module");
  });

  it("@sheet(Name) names the sheet of a file whose name cannot (a pull's ~2)", () => {
    const m = parseSourceFile("names/sheets/bs~2.xln", "// Sheet bs\n@sheet(bs)\n\nX = 1;\n");
    expect([m.form, m.sheet, m.entries[0]!.scope]).toEqual(["sheet", "bs", "bs"]);
  });
});

describe("convertSheetBlocks: the quick fix for an old sheet file", () => {
  const old = [
    "// Sheet IS, pulled by xln from is-model.xlsx.",
    "",
    "@scope(IS)",
    "",
    "@workbook",
    "",
    "years @IS!C2 =  SEQUENCE(1, 5, 2025);",
    "",
    "@scope(IS)",
    "",
    "Sales @C3# = SEQUENCE(1,5,20000,3400);",
    "COGS @C4# = Sales * 0.6;",
    "",
  ].join("\n");

  it("writes @workbook above each name of a @workbook block, drops the blocks and the sheet in addresses", () => {
    const c = convertSheetBlocks(old, "IS");
    expect(c).toEqual({
      text: ["// Sheet IS, pulled by xln from is-model.xlsx.", "", "@workbook", "years @C2 =  SEQUENCE(1, 5, 2025);", "", "Sales @C3# = SEQUENCE(1,5,20000,3400);", "COGS @C4# = Sales * 0.6;", ""].join("\n"),
    });
    const m = parseSourceFile("names/sheets/IS.xln", (c as { text: string }).text);
    expect(m.entries.map((e) => [e.name, e.scope])).toEqual([["years", undefined], ["Sales", "IS"], ["COGS", "IS"]]);
  });

  it("refuses a file holding another sheet's block, and says why", () => {
    expect(convertSheetBlocks("@scope(IS)\nA = 1;\n@scope(BS)\nB = 2;\n", "IS")).toEqual({ reason: "line 3: @scope(BS) holds names local to another sheet; move them to that sheet's file first" });
  });

  it("the checker offers it as an info with a whole-file fix", () => {
    const m = model({ "names/sheets/IS.xln": old });
    const p = checkFile(m, "names/sheets/IS.xln");
    expect(p.map((x) => [x.severity, x.code])).toEqual([["info", "old-blocks"]]);
    expect(p[0]!.fix).toEqual({ title: "Convert to per-name @workbook", start: 0, end: old.length, text: (convertSheetBlocks(old, "IS") as { text: string }).text });
  });
});

describe("pull", () => {
  it("a fresh pull writes the new form: @workbook above the workbook name on a cell, bare addresses, no blocks", () => {
    const r = project(F7);
    expect(r.files[S1]).toContain("@C1 = RateX+Rate;\n@workbook\nSpl @E1# = SEQUENCE(3)*Rate;\n@B2:B11 = A2*Rate;\n");
    expect(r.files[S1]).not.toMatch(/^@scope|!E1/m);
  });

  /** F7's S1 file as a pull before M3d wrote it. */
  const OLD_S1 = [
    "// Sheet S1, pulled by xln from book.xlsx (before M3d).",
    "",
    "@scope(S1)",
    "",
    "@A1 = Rate*2;",
    "@C1 = RateX+Rate;",
    "",
    "@workbook",
    "",
    "Spl @S1!E1# = SEQUENCE(3)*Rate;",
    "",
    "@scope(S1)",
    "",
    "@B2:B11 = A2*Rate;",
    "@C2 = \"Rate is \"&Rate;",
    "@C3 = LET(Rate, 5, Rate*2);",
    "@C4 = Fn(10);",
    "@C5 = 'S2'!Loc+Loc;",
    "@C6 = SUM(Spl);",
    "@C7 = INDIRECT(\"Rate\");",
    "@C8 = Rate*3;",
    "@C9 = ROWS(E1#)*Rate;",
    "@C10 = Rate2+Fn(1);",
    "",
  ].join("\n");

  it("an old sheet file builds as it is and is no edit for the pull's guard; a pull writes the new form", () => {
    const files = { ...project(F7).files, [S1]: OLD_S1 };
    expect(build(F7, files).status, why(build(F7, files))).toBe("up-to-date");
    const created = applyChangeSet(F7, [{ op: "set-name", name: "Errored_balance_base", scope: null, stored: "S1!$D$5", display: "S1!$D$5", comment: null, hidden: false, fields: ["created"] }]);
    expect(unbuiltEdits({ workbook: created, fileName: "book.xlsx", files })).toEqual([]);
    const p = pullProject(created, "book.xlsx");
    expect(p.files[S1]).toContain("@C5 = 'S2'!Loc+Loc;\n@workbook\nErrored_balance_base @D5 = ;\n@C6 = SUM(Spl);\n");
    expect(build(created, { ...files, ...p.files }).status).toBe("up-to-date");
  });

  it("a scope changed in Excel (deleted, created again in the other scope) comes through: @workbook exactly per the workbook", () => {
    const files = { ...project(F7).files };
    expect(files[S1]).toContain("@workbook\nSpl @E1# =");
    const local = applyChangeSet(F7, [{ op: "rescope-name", name: "Spl", from: null, to: "S1" }]);
    expect(unbuiltEdits({ workbook: local, fileName: "book.xlsx", files })).toEqual([]);
    const p = pullProject(local, "book.xlsx");
    expect(p.files[S1]).toContain("@C1 = RateX+Rate;\nSpl @E1# = SEQUENCE(3)*Rate;\n");
    expect(build(local, p.files).status).toBe("up-to-date");
    const back = applyChangeSet(local, [{ op: "rescope-name", name: "Spl", from: "S1", to: null }]);
    expect(unbuiltEdits({ workbook: back, fileName: "book.xlsx", files: p.files })).toEqual([]);
    const q = pullProject(back, "book.xlsx");
    expect(q.files[S1]).toBe(files[S1]);
    expect(build(back, q.files).status).toBe("up-to-date");
  });
});

describe("scope: the source sets it as Excel does (no setting, 2026-10-06)", () => {
  const r = project(F7);

  it("a workbook name on a sheet's cell read only there is a hint with a fix that removes @workbook; removing it is a re-scope (info)", () => {
    const files = { ...r.files };
    const w = problems(files, S1, ctxOf(r, files));
    expect(w).toEqual([["hint", "workbook-on-cell", "Spl", "workbook name on a cell of S1, read only on S1: remove @workbook to make it local to S1 (the build then moves it)"]]);
    const fix = checkFile(model(files), S1, ctxOf(r, files))[0]!.fix!;
    const fixed = files[S1]!.slice(0, fix.start) + fix.text + files[S1]!.slice(fix.end);
    expect(fixed).toContain("@C1 = RateX+Rate;\nSpl @E1# = SEQUENCE(3)*Rate;\n");
    files[S1] = fixed;
    expect(problems(files, S1, ctxOf(r, files))).toEqual([["info", "rescope", "Spl", "Spl: the build moves it from workbook scope to sheet S1 (a scope change, rescope-name)"]]);
    // The build: Spl is read by S1's own cells only, so the move is one rescope-name.
    const b = build(F7, files);
    expect(b.status, why(b)).toBe("built");
    expect(b.plan.changeSet.changes).toEqual([{ op: "rescope-name", name: "Spl", from: null, to: "S1" }]);
  });

  it('an old config\'s "names": {"scope": "excel"} is ignored: the same hint, and removing @workbook is the same re-scope', () => {
    const files = { ...r.files, "xln.config.json": '{ "names": { "scope": "excel" } }\n' };
    expect(problems(files, S1, ctxOf(r, files)).map((p) => p[1])).toEqual(["workbook-on-cell"]);
    edit(files, S1, "@workbook\nSpl @E1#", "Spl @E1#");
    expect(problems(files, S1, ctxOf(r, files)).map((p) => [p[0], p[1]])).toEqual([["info", "rescope"]]);
    const b = build(F7, files);
    expect(b.status, why(b)).toBe("built");
    expect(b.plan.changeSet.changes).toEqual([{ op: "rescope-name", name: "Spl", from: null, to: "S1" }]);
  });

  it("adding @workbook back is a re-scope the other way (info), with the advice (hint)", () => {
    const files = { ...r.files };
    edit(files, S1, "@workbook\nSpl @E1#", "Spl @E1#");
    const b = build(F7, files);
    const built = { ...files, ...b.files };
    expect(problems(built, S1, ctxOf(r, built))).toEqual([]);
    edit(built, S1, "Spl @E1#", "@workbook\nSpl @E1#");
    expect(problems(built, S1, ctxOf(r, built)).map((p) => [p[0], p[1], p[3]])).toEqual([
      ["hint", "workbook-on-cell", "workbook name on a cell of S1, read only on S1: remove @workbook to make it local to S1 (the build then moves it)"],
      ["info", "rescope", "Spl: the build moves it from sheet S1 to workbook scope (a scope change, rescope-name)"],
    ]);
    const b2 = build(b.bytes!, built);
    expect(b2.status, why(b2)).toBe("built");
    expect(b2.plan.changeSet.changes).toEqual([{ op: "rescope-name", name: "Spl", from: "S1", to: null }]);
  });
});

describe("workbook-on-cell: only a name read on its own sheet could be local (feedback 2026-10-07)", () => {
  const r = project(F7);
  const codes = (files: Record<string, string>) => problems(files, S1, ctxOf(r, files)).map((p) => p[1]);

  it("another sheet's cell reads it unqualified: it needs workbook scope, no finding", () => {
    const files = { ...r.files };
    edit(files, S2, "@A1 = Loc;", "@A1 = Loc + SUM(Spl);");
    expect(codes(files)).toEqual([]);
  });

  it("a name of another scope reads it: no finding", () => {
    const files = { ...r.files };
    files["names/_unmanaged.xln"] += "\nSplTotal = SUM(Spl);\n";
    expect(codes(files)).toEqual([]);
  });

  it("read from another sheet qualified (S1!Spl) still counts as read from there: no finding", () => {
    const files = { ...r.files };
    edit(files, S2, "@A1 = Loc;", "@A1 = Loc + SUM(S1!Spl);");
    expect(codes(files)).toEqual([]);
  });

  it("read only by its own sheet's cells (S1!C6 = SUM(Spl)): a hint with its quick fix", () => {
    const files = { ...r.files };
    const found = checkFile(model(files), S1, ctxOf(r, files));
    expect(found.map((p) => [p.severity, p.code, p.fix?.title])).toEqual([["hint", "workbook-on-cell", "Remove @workbook: make Spl local to S1"]]);
  });
});

describe("xln.config.json: names.scope is no longer a setting", () => {
  it("an old file with it is read, with one note and no problem; the rest of the file still counts", () => {
    const { config, problems: probs, notes } = parseConfig(JSON.stringify({ names: { scope: "excel" }, build: { embed: true } }));
    expect(config).toEqual({ build: { embed: true } });
    expect(probs).toEqual([]);
    expect(notes).toEqual([NAMES_SCOPE_NOTE]);
    expect(notes[0]).toMatch(/^names\.scope: no longer a setting \(ignored\)/);
  });

  it("another names.* key is an unknown setting; a file without names has no note", () => {
    expect(parseConfig('{ "names": { "other": 1 } }').problems).toEqual(["names.other: unknown setting (ignored)"]);
    expect(parseConfig(defaultConfigText()).notes).toEqual([]);
  });
});

describe("the checker: what is wrong, where, and what to write", () => {
  const r = project(F7);

  it("reports each wrong statement on its text, with the message the build gives too", () => {
    const files = { ...r.files };
    edit(files, S1, "@C1 = RateX+Rate;", "@C1 = SUM(RateX+Rate;");
    edit(files, S1, "@C2 = \"Rate is \"&Rate;", "@C2 = Ratios! + 1;");
    edit(files, S1, "@C3 = LET(Rate, 5, Rate*2);", "@C3 = FOO(1) + _xlfn.BAR(2);");
    edit(files, S1, "@C4 = Fn(10);", "@C4 = Nope!A1 + 1;");
    edit(files, S1, "@C8 = Rate*3;", "@C8 = OnlyS2 * 2;");
    edit(files, S1, "Spl @E1#", "Spl @S1!E1#");
    edit(files, S1, "@C10 = ", "@workbook\n@C10 = ");
    files[S1] += "@C40 = 1;\nNew @Z9 = ;\n";
    edit(files, S2, "Loc = 7;", "Loc = 7;\nOnlyS2 = 1;\nLoc = 8;");
    const ctx = ctxOf(r, files);
    expect(problems(files, S1, ctx)).toEqual([
      ["error", "syntax", "(", "@C1: this '(' is never closed"],
      ["hint", "workbook-on-cell", "Spl", "workbook name on a cell of S1, read only on S1: remove @workbook to make it local to S1 (the build then moves it)"],
      ["warning", "address", "@S1!E1#", "the sheet is the file's (S1): write @E1#"],
      ["error", "syntax", "Ratios!", "@C2: expected a reference or a name after '!'"],
      ["error", "unknown-function", "FOO", "FOO(…): no built-in function and no LAMBDA of that name; Excel would store it as _xludf.FOO (#NAME?)"],
      ["error", "unknown-function", "_xlfn.BAR", "@C3: '_xlfn.BAR': 'BAR' is not in the catalogue: Excel would show #NAME?; write a function the catalogue knows (without a prefix: the build adds it)"],
      ["error", "unknown-sheet", "Nope!A1", "Nope!A1: the workbook has no sheet 'Nope'"],
      ["error", "C5.other-sheet", "OnlyS2", "OnlyS2 is local to S2, not to S1: unqualified it is #NAME? in Excel; write 'S2'!OnlyS2"],
      ["error", "syntax", "@", "@workbook scopes a name, and @C10 has none: remove the @workbook line above it"],
      ["error", "address", "@C40", "no cell statement at S1!C40 in the last pull: addresses are set in Excel (write the formula in Excel, then pull)"],
      ["error", "address", "@Z9", "S1!New is not in the last pull: a cell is named in Excel (Name Manager or the Name Box), then pulled"],
    ]);
    expect(problems(files, S2, ctx).filter((p) => p[1] === "duplicate")).toEqual([
      ["error", "duplicate", "Loc", "S2!Loc is also defined at names/sheets/S2.xln:15: keep one (a name can be defined once per scope)"],
    ]);
    // The build refuses exactly these errors (the source's), at their lines.
    const b = build(F7, files);
    expect(b.status).toBe("refused");
    const checked = [...checkFile(model(files), S1, ctx), ...checkFile(model(files), S2, ctx)].filter((p) => p.severity === "error").map((p) => p.message).sort();
    expect(b.plan.problems.filter((p) => p.severity === "error" && p.file !== undefined).map((p) => p.message).sort()).toEqual(checked);
  });

  it("a formula as the workbook has it keeps its findings as warnings, unless the source broke what it reads", () => {
    const files = { ...r.files };
    // The workbook's own C4 reads Fn; with Fn gone from the source it breaks: the build
    // refuses on it, so the editor says error too (§14 issue 10, fixed 2026-10-07).
    edit(files, "names/_unmanaged.xln", "Fn = LAMBDA(x, x*Rate);\n", "");
    const p = problems(files, S1, ctxOf(r, files)).filter((x) => x[1] === "in-use");
    expect(p.map((x) => x[0])).toEqual(["error", "error"]);
    expect(p[0]![3]).toContain("the last pull had Fn, which the source renames, moves or deletes, so this formula breaks and the build refuses");
    // A name that was never there: as in the workbook, a warning.
    edit(files, S1, "@C8 = Rate*3;", "@C8 = Nope*3;");
    const lock = parseLockfile(files[LOCK_FILE]!);
    lock.cells!["S1!C8"]!.formula = definitionHash("Nope*3");
    const q = problems(files, S1, { ...ctxOf(r, files), lock }).filter((x) => x[2] === "Nope");
    expect(q.map((x) => x[0])).toEqual(["warning"]);
    expect(q[0]![3]).toContain("(as in the workbook since the last pull: the build leaves it; fix it here to have it written)");
  });

  it("a wrong sheet in a sheet file's address is an error that says where the statement belongs", () => {
    const files = { ...r.files };
    edit(files, S1, "Spl @E1#", "Spl @S2!E1#");
    expect(problems(files, S1, ctxOf(r, files)).filter((p) => p[1] === "address")).toEqual([
      ["error", "address", "@S2!E1#", "S2!E1 is a cell of S2, and this file holds S1's statements: write it in names/sheets/ for S2 (as @E1#)"],
    ]);
  });
});

// The author's test (2026-10-06): a sheet Ratios with names from Create from Selection
// (workbook scope) and a workbook name `years` on IS!C2. Reads a copy of is-model.xlsx from
// the folder in XLN_PLAY (a copy, never the author's folder itself); skipped without it.
describe.skipIf(!process.env["XLN_PLAY"])("the author's case (is-model.xlsx)", () => {
  const bytes = () => new Uint8Array(readFileSync(join(process.env["XLN_PLAY"]!, "is-model.xlsx")));

  it("pulls Ratios with @workbook above each Create-from-Selection name (a hint each); IS with @workbook above years, which Ratios reads: no finding; no blocks", () => {
    const wb = bytes();
    const r = pullProject(wb, "is-model.xlsx");
    const ratios = r.files["names/sheets/Ratios.xln"]!;
    const is = r.files["names/sheets/IS.xln"]!;
    for (const n of ["Net_operating_assets @C3", "ROS @C4", "Turnover @C5", "ROI @C6"]) expect(ratios).toContain(`@workbook\n${n} =`);
    expect(is).toMatch(/\n@workbook\nyears @C2 = /);
    expect(ratios + is).not.toMatch(/^@scope/m);
    const files = { ...r.files };
    const ctx = ctxOf(r, files);
    expect(problems(files, "names/sheets/Ratios.xln", ctx).map((p) => [p[0], p[1], p[2]])).toEqual([
      ["hint", "workbook-on-cell", "Net_operating_assets"],
      ["hint", "workbook-on-cell", "ROS"],
      ["hint", "workbook-on-cell", "Turnover"],
      ["hint", "workbook-on-cell", "ROI"],
    ]);
    expect(problems(files, "names/sheets/IS.xln", ctx).map((p) => [p[0], p[1], p[2]])).toEqual([]);
    expect(build(wb, files).status).toBe("up-to-date");
  });

  it("removing @workbook above ROS and building gives exactly one rescope-name", () => {
    const wb = bytes();
    const files = { ...pullProject(wb, "is-model.xlsx").files };
    edit(files, "names/sheets/Ratios.xln", "@workbook\nROS @C4", "ROS @C4");
    const b = build(wb, files);
    expect(b.status, why(b)).toBe("built");
    expect(b.plan.changeSet.changes).toEqual([{ op: "rescope-name", name: "ROS", from: null, to: "Ratios" }]);
  });
});
