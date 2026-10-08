// Lockfile format 3: number literals count by their value (Excel saves `1E-14` as
// `0.00000000000001`), hashes are `h:` and 16 hex digits, one line per entry; lockfiles of
// formats 1 and 2 keep working without spurious drift.
import { describe, expect, it } from "vitest";
import {
  applyChangeSet,
  buildLockfile,
  buildWorkbook,
  canonicalNumber,
  commentHash,
  definitionHash,
  definitionHashV2,
  equalModuloWhitespace,
  LOCK_FILE,
  lockfileJson,
  lockfileText,
  parseLockfile,
  pullProject,
  unbuiltEdits,
  sha256,
  sourceHash,
  sourceHashV1,
  stringifyJson,
  tokenKeys,
  type BuildResult,
  type Change,
} from "../../src/index.js";
import { edit, fixture, why, withWorkbookXml, workbookXml } from "../build/helpers.js";

const F7 = fixture("f7_base.xlsx");
const S1 = "names/sheets/S1.xln";
const UNMANAGED = "names/_unmanaged.xln";

describe("number literals count by their value", () => {
  it("canonical spelling", () => {
    expect(canonicalNumber("1E-14")).toBe("1e-14");
    expect(canonicalNumber("1e-14")).toBe("1e-14");
    expect(canonicalNumber("0.00000000000001")).toBe("1e-14");
    expect(canonicalNumber(".5")).toBe("0.5");
    expect(canonicalNumber("0.50")).toBe("0.5");
    expect(canonicalNumber("1.0")).toBe("1");
    expect(canonicalNumber("1.")).toBe("1");
    expect(canonicalNumber("007")).toBe("7");
    expect(canonicalNumber("1E+3")).toBe("1000");
    expect(canonicalNumber("1E+300")).toBe("1e+300");
    // Excel keeps 15 significant digits of a typed number.
    expect(canonicalNumber("0.1234567890123456789")).toBe(canonicalNumber("0.123456789012346"));
    expect(canonicalNumber("12345678901234567")).toBe(canonicalNumber("12345678901234600"));
  });

  it("equal where Excel takes them as equal", () => {
    const same: [string, string][] = [
      ["1E-14", "0.00000000000001"],
      ["FN.FIXPOINT(LAMBDA(x, 1 + x / 2), SEQUENCE(1, 3, 0, 0), 1E-14, 100)", "FN.FIXPOINT(LAMBDA(x,1+x/2),SEQUENCE(1,3,0,0),0.00000000000001,100)"],
      ["1e-14", "1E-14"],
      [".5*A1", "0.5*A1"],
      ["1.0+B2", "1+B2"],
      ["{1,2.0;3E0,4}", "{1,2;3,4}"],
      ["-1E-14", "-0.00000000000001"],
      ["50.0%", "50%"],
    ];
    for (const [a, b] of same) {
      expect(equalModuloWhitespace(a, b), `${a} = ${b}`).toBe(true);
      expect(definitionHash(a), `${a} = ${b}`).toBe(definitionHash(b));
    }
  });

  it("strings, references, sheet names, the percent operator and other numbers are untouched", () => {
    const differ: [string, string][] = [
      ['"1E-14"', '"0.00000000000001"'],
      ['"1.0"', '"1"'],
      ["50%", "0.5"],
      ["1.5", "1.6"],
      ["A1", "A10"],
      ["'1E5'!A1", "'100000'!A1"],
      ["'S1.0'!B2", "'S1'!B2"],
      ["1:3", "1:4"],
      ["R1C1+1", "R1C1+1.5"],
      ["T[[#This Row],[1.0]]", "T[[#This Row],[1]]"],
    ];
    for (const [a, b] of differ) {
      expect(equalModuloWhitespace(a, b), `${a} ≠ ${b}`).toBe(false);
      expect(definitionHash(a), `${a} ≠ ${b}`).not.toBe(definitionHash(b));
    }
    // A sheet name with digits keeps its spelling in the key.
    expect(tokenKeys("'1E5'!A1")).toEqual(["ref:|1E5|!A1"]);
  });
});

describe("hashes of format 3 and of formats 1 and 2", () => {
  it("format 3: h: and 16 hex digits of SHA-256 over the tokens, numbers by value", () => {
    expect(definitionHash("1 + 2.0")).toBe("h:" + sha256("number:1\nop:+\nnumber:2").slice(0, 16));
    expect(definitionHash("1+2")).toMatch(/^h:[0-9a-f]{16}$/);
    expect(commentHash("a\r\nb")).toBe("h:" + sha256("a\nb").slice(0, 16));
    expect(commentHash("")).toBeNull();
  });

  it("formats 1 and 2: sha256: and all 64 digits, numbers as written (unchanged)", () => {
    expect(definitionHashV2("1 + 2.0")).toBe("sha256:" + sha256("number:1\nop:+\nnumber:2.0"));
    expect(definitionHashV2("1E-14")).not.toBe(definitionHashV2("0.00000000000001"));
    expect(definitionHashV2("'BS'!$A$1")).toBe(definitionHashV2("BS!$A$1"));
  });

  it("provenance: the tag hash counts numbers by value; tags of the older hash still read", () => {
    expect(sourceHash("1E-14", "c")).toBe(sourceHash("0.00000000000001", "c"));
    expect(sourceHashV1("1E-14", "c")).not.toBe(sourceHashV1("0.00000000000001", "c"));
    // A definition with numbers already canonical has the same tag either way.
    expect(sourceHash("_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)", "d")).toBe(sourceHashV1("_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)", "d"));
  });
});

describe("the lockfile text", () => {
  it("one line per name and per cell; parses back", () => {
    const p = pullProject(F7, "book.xlsx");
    const text = p.files[LOCK_FILE]!;
    const lock = parseLockfile(text);
    expect(lock.format).toBe("xln.lock/4");
    expect(text).toBe(lockfileText(lock));
    expect(JSON.parse(text)).toEqual(lock);
    const lines = text.split("\n");
    expect(lines).toContain('    "Rate": { "definition": "' + lock.names["Rate"]!.definition + '", "comment": null, "hidden": false },');
    expect(lines).toContain('    "Spl": { "sheet": "S1", "range": "E1", "name": "Spl", "formula": "' + lock.cells!["Spl"]!.formula + '" },');
    expect(lines.length).toBe(Object.keys(lock.names).length + Object.keys(lock.cells!).length + 9);
  });
});

// ---------------------------------------------------------------------------------------
// Compatibility: a project or a part whose lockfile is format 2.

/** The lockfile an earlier xln (format 2) wrote for this workbook: `sha256:` hashes, the generic JSON layout. */
function v2LockText(bytes: Uint8Array): string {
  const p = pullProject(bytes, "book.xlsx");
  return stringifyJson(lockfileJson(buildLockfile("book.xlsx", p.names, p.statements, "v2"))) + "\n";
}

function build(bytes: Uint8Array, files: Record<string, string>, opts: { force?: boolean; embed?: boolean } = {}): BuildResult {
  return buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files }, { embed: false, provenance: false, ...opts });
}

/** F7 built from a source that writes `1E-14` in a name and in a cell, with a format-2 lockfile of the result. */
function withSmallNumbers(embed = false): { bytes: Uint8Array; files: Record<string, string> } {
  const files = { ...pullProject(F7, "book.xlsx").files };
  edit(files, UNMANAGED, "RateX = 0.5;", "RateX = 1E-14;");
  edit(files, S1, "@C8 = Rate*3;", "@C8 = Rate*1E-14;");
  const r = build(F7, files, { embed });
  expect(r.status, why(r)).toBe("built");
  const project = { ...files, ...r.files };
  project[LOCK_FILE] = v2LockText(r.bytes!);
  expect(parseLockfile(project[LOCK_FILE]).format).toBe("xln.lock/2");
  return { bytes: r.bytes!, files: project };
}

/** What Excel does on save: writes the number out in full. */
function excelRewrites(bytes: Uint8Array): Uint8Array {
  const xml = workbookXml(bytes);
  expect(xml).toContain(">1E-14</definedName>");
  const named = withWorkbookXml(bytes, xml.replace(">1E-14</definedName>", ">0.00000000000001</definedName>"));
  const cell: Change = { op: "set-cell-formula", sheet: "S1", range: "C8", stored: "Rate*0.00000000000001", display: "Rate*0.00000000000001" };
  return applyChangeSet(named, [cell]);
}

describe("format-2 lockfiles still work", () => {
  it("a format-2 lockfile of the same state: nothing to do, no drift", () => {
    const { bytes, files } = withSmallNumbers();
    const r = build(bytes, files);
    expect(r.status, why(r)).toBe("up-to-date");
    expect(r.plan.excelChanges).toEqual([]);
  });

  it("Excel writing 1E-14 out in full is no change, and the next lockfile has format-3 hashes, no sha256 left", () => {
    const { bytes, files } = withSmallNumbers();
    const excel = excelRewrites(bytes);
    const r = build(excel, files, { force: true });
    expect(r.plan.excelChanges, why(r)).toEqual([]);
    expect(r.plan.conflicts).toEqual([]);
    expect(r.plan.changeSet.changes).toEqual([]);
    expect(r.status, why(r)).toBe("built");
    const text = r.files![LOCK_FILE]!;
    const lock = parseLockfile(text);
    expect(lock.format).toBe("xln.lock/4");
    expect(text).not.toContain("sha256:");
    expect(lock.names["RateX"]!.definition).toBe(definitionHash("1E-14"));
    expect(lock.cells!["S1!C8"]!.formula).toBe(definitionHash("Rate*1E-14"));
    expect(build(excel, { ...files, ...r.files }).status).toBe("up-to-date");
  });

  it("a real edit in Excel is still one; the entry it keeps is upgraded (the source still matched it)", () => {
    const { bytes, files } = withSmallNumbers();
    const excel = applyChangeSet(excelRewrites(bytes), [{ op: "set-cell-formula", sheet: "S1", range: "C7", stored: "Rate*7", display: "Rate*7" }]);
    const r = build(excel, files, { force: true });
    expect(r.plan.excelChanges.map((x) => x.key)).toEqual(["S1!C7"]);
    const lock = parseLockfile(r.files![LOCK_FILE]!);
    expect(lock.format).toBe("xln.lock/4");
    expect(lock.cells!["S1!C7"]!.formula).toBe(definitionHash('INDIRECT("Rate")'));
    const r2 = build(excel, { ...files, ...r.files });
    expect(r2.plan.excelChanges.map((x) => x.key)).toEqual(["S1!C7"]);
    expect(r2.plan.changeSet.changes).toEqual([]);
  });

  it("an old entry the source no longer matches stays sha256: in a format-4 file and is compared in its own format", () => {
    const { bytes, files } = withSmallNumbers();
    // The source drops C10's statement (the build leaves the cell) and edits C8: C10's entry cannot be upgraded.
    edit(files, S1, "@C10 = Rate2+Fn(1);\n", "");
    edit(files, S1, "@C8 = Rate*1E-14;", "@C8 = Rate*3E-14;");
    const r = build(bytes, files, { force: true });
    expect(r.status, why(r)).toBe("built");
    const text = r.files![LOCK_FILE]!;
    expect(parseLockfile(text).format).toBe("xln.lock/4");
    expect(parseLockfile(text).cells!["S1!C10"]!.formula).toMatch(/^sha256:[0-9a-f]{64}$/);
    // Put back, the statement matches the workbook through its old hash.
    edit(files, S1, "@C9 = ROWS(E1#)*Rate;\n", "@C9 = ROWS(E1#)*Rate;\n@C10 = Rate2+Fn(1);\n");
    const r2 = build(r.bytes!, { ...files, ...r.files });
    expect(r2.status, why(r2)).toBe("up-to-date");
    expect(r2.plan.excelChanges).toEqual([]);
  });

  it("a source edit against a format-2 lockfile is a change, as before", () => {
    const { bytes, files } = withSmallNumbers();
    edit(files, S1, "@C8 = Rate*1E-14;", "@C8 = Rate*2E-14;");
    const r = build(bytes, files);
    expect(r.status, why(r)).toBe("built");
    expect(r.plan.changeSet.changes).toEqual([expect.objectContaining({ op: "set-cell-formula", range: "C8", stored: "Rate*2E-14" })]);
  });

  it("known limit: a source edit of a formula whose numbers Excel respelled, against a format-2 lockfile, is a conflict", () => {
    // The old hash is of `Rate*1E-14`; neither side has that text any more, so nothing
    // shows that Excel's `Rate*0.00000000000001` is the same state. The build refuses
    // (it never overwrites); a pull with format 3 settles it.
    const { bytes, files } = withSmallNumbers();
    edit(files, S1, "@C8 = Rate*1E-14;", "@C8 = Rate*2E-14;");
    const r = build(excelRewrites(bytes), files);
    expect(r.status).toBe("refused");
    expect(r.plan.conflicts.map((c) => c.key)).toEqual(["S1!C8"]);
  });
});

describe("the pull's guard with a format-2 lockfile", () => {
  it("Excel writing 1E-14 out in full is no source edit: nothing to refuse; a source edit is listed", () => {
    const { bytes, files } = withSmallNumbers();
    expect(unbuiltEdits({ workbook: excelRewrites(bytes), fileName: "book.xlsx", files })).toEqual([]);
    edit(files, S1, "@C8 = Rate*1E-14;", "@C8 = Rate*2E-14;");
    expect(unbuiltEdits({ workbook: excelRewrites(bytes), fileName: "book.xlsx", files }).map((e) => e.key)).toEqual(["S1!C8"]);
  });
});

describe("a literal Excel spells otherwise on save (format 4)", () => {
  // The build writes a literal as typed; Excel stores the number and writes its own
  // spelling on save. Measured for 1E-14 (written out in full) and 1E+99, 1E+300 (kept);
  // the other spellings here are assumed (probes/README.md, "To measure").
  const typed = "0.10 + .5 + 1E3 + 1e-7 + 00.1";
  const excels = "0.1+0.5+1000+1E-07+0.1";

  function builtThenSaved(name: string, cell: string) {
    const files = { ...pullProject(F7, "book.xlsx").files };
    edit(files, UNMANAGED, "RateX = 0.5;", `RateX = ${name};`);
    edit(files, S1, "@C8 = Rate*3;", `@C8 = Rate*(${cell});`);
    const r = build(F7, files);
    expect(r.status, why(r)).toBe("built");
    // The built file has the source's spelling, as typed.
    expect(workbookXml(r.bytes!)).toContain(`<definedName name="RateX">${name}</definedName>`);
    const project = { ...files, ...r.files };
    const respell = (n: string, c: string) =>
      applyChangeSet(r.bytes!, [
        { op: "set-name", name: "RateX", scope: null, stored: n, display: n, comment: null, hidden: false, fields: ["definition"] },
        { op: "set-cell-formula", sheet: "S1", range: "C8", stored: `Rate*(${c})`, display: `Rate*(${c})` },
      ]);
    return { project, respell };
  }

  it("is no change in Excel: nothing to build, no source edit for the pull's guard, the pull takes Excel's spelling", () => {
    const { project, respell } = builtThenSaved(typed, typed);
    const excel = respell(excels, excels);
    const r = build(excel, project);
    expect(r.plan.excelChanges, why(r)).toEqual([]);
    expect(r.plan.conflicts).toEqual([]);
    expect(r.status, why(r)).toBe("up-to-date");
    expect(unbuiltEdits({ workbook: excel, fileName: "book.xlsx", files: project })).toEqual([]);
    const pulled = pullProject(excel, "book.xlsx").files;
    expect(pulled[UNMANAGED]).toContain(`RateX = ${excels};`);
    expect(parseLockfile(pulled[LOCK_FILE]!).names["RateX"]).toEqual(parseLockfile(project[LOCK_FILE]!).names["RateX"]);
    expect(parseLockfile(pulled[LOCK_FILE]!).cells!["S1!C8"]).toEqual(parseLockfile(project[LOCK_FILE]!).cells!["S1!C8"]);
  });

  it("known limit: past 15 significant digits the key rounds; if Excel truncates instead, the save reads as an edit in Excel", () => {
    const long = "0.1234567890123456789";
    const { project, respell } = builtThenSaved(long, "1");
    expect(build(respell("0.123456789012346", "1"), project).plan.excelChanges).toEqual([]);
    expect(build(respell("0.123456789012345", "1"), project).plan.excelChanges.map((x) => x.key)).toEqual(["RateX"]);
  });
});
