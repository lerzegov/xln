// The pull's guard (decided 2026-10-06): every pull is fresh, so before it replaces the
// project it lists the source edits not built yet, as the build sees them.
import { describe, expect, it } from "vitest";
import { applyChangeSet, formatUnbuiltEdit, LOCK_FILE, pullProject, rewrittenFiles, unbuiltEdits } from "../../src/index.js";
import { build, edit, fixture } from "./helpers.js";

const F7 = fixture("f7_base.xlsx");
const S1 = "names/sheets/S1.xln";
const UNMANAGED = "names/_unmanaged.xln";

function project(bytes: Uint8Array = F7): Record<string, string> {
  return { ...pullProject(bytes, "book.xlsx").files };
}

function guard(files: Record<string, string>, bytes: Uint8Array = F7): string[] {
  return unbuiltEdits({ workbook: bytes, fileName: "/some/folder/book.xlsx", files }).map(formatUnbuiltEdit);
}

describe("unbuiltEdits", () => {
  it("nothing for a project as pulled, nor for no project", () => {
    expect(guard(project())).toEqual([]);
    expect(guard({})).toEqual([]);
  });

  it("lists each edit at its file and line, with the change the build would make", () => {
    const files = project();
    edit(files, UNMANAGED, "Rate2 = Rate*2;", "/** Twice the rate. */\nRate2 = Rate*2;");
    edit(files, UNMANAGED, "RateX = 0.5;\n", "");
    edit(files, S1, "@C8 = Rate*3;", "@C8 = Rate*4;");
    files[UNMANAGED] += "Fresh = 1;\n";
    expect(guard(files)).toEqual([
      "names/_unmanaged.xln:9  Rate2: update Rate2 (comment)",
      "names/_unmanaged.xln:10  Fresh: create Fresh",
      "names/sheets/S1.xln:19  S1!C8: set formula of S1!C8",
      "RateX: delete RateX",
    ]);
  });

  it("a rename and a scope change are one line each, under the name the source has", () => {
    const files = project();
    edit(files, UNMANAGED, "Rate2 = Rate*2;", "@renamed(Rate2)\nRateTwo = Rate*2;");
    edit(files, S1, "@workbook\nSpl @E1#", "Spl @E1#");
    // The cell reading it, renamed too (M5): no line of its own, the rename says it.
    edit(files, S1, "@C10 = Rate2+Fn(1);", "@C10 = RateTwo+Fn(1);");
    expect(guard(files)).toEqual(["names/_unmanaged.xln:9  RateTwo: rename Rate2 → RateTwo, rewriting it in 1 cell formula", "names/sheets/S1.xln:10  S1!Spl: move Spl to sheet S1"]);
  });

  it("a source the build would refuse holds edits that were never built: its errors are listed", () => {
    const files = project();
    edit(files, S1, "@C8 = Rate*3;", "@C8 = Rate*(3;");
    expect(guard(files)).toEqual([
      "names/sheets/S1.xln:19  @C8: expected ')' to close the '(' at column 6, found the end of the formula",
      "names/sheets/S1.xln:19  S1!C8: edited in the source, not built",
    ]);
  });

  it("Excel's own changes since the last pull are not source edits", () => {
    const files = project();
    const xl = applyChangeSet(F7, [{ op: "set-cell-formula", sheet: "S1", range: "C8", stored: "Rate*9", display: "Rate*9" }]);
    expect(guard(files, xl)).toEqual([]);
  });

  it("without a lockfile the source is compared with the workbook itself", () => {
    const files = project();
    delete files[LOCK_FILE];
    expect(guard(files)).toEqual([]);
    edit(files, S1, "@C8 = Rate*3;", "@C8 = Rate*4;");
    expect(guard(files)).toEqual(["names/sheets/S1.xln:19  S1!C8: set formula of S1!C8"]);
  });

  it("an unreadable lockfile is said, so nothing is replaced unseen", () => {
    const files = { ...project(), [LOCK_FILE]: "{ not json" };
    expect(guard(files)[0]).toMatch(/^xln\.lock\.json  the lockfile cannot be read/);
  });

  it("after the build, nothing is left", () => {
    const files = project();
    edit(files, S1, "@C8 = Rate*3;", "@C8 = Rate*4;");
    const r = build(F7, files);
    expect(guard({ ...files, ...r.files }, r.bytes!)).toEqual([]);
  });
});

describe("rewrittenFiles (M3e): layout or comments only, rewritten by a pull", () => {
  const pulled = project();
  const rewritten = (files: Record<string, string>) => rewrittenFiles(files, pulled, unbuiltEdits({ workbook: F7, fileName: "book.xlsx", files }));

  it("nothing for a project as pulled; line endings do not count", () => {
    expect(rewritten(project())).toEqual([]);
    const crlf = project();
    crlf[S1] = crlf[S1]!.split("\n").join("\r\n");
    expect(rewritten(crlf)).toEqual([]);
  });

  it("a // comment, blank lines, a definition over two lines: not an edit, but the file is rewritten", () => {
    const files = project();
    edit(files, UNMANAGED, "Rate2 = Rate*2;", "// twice the rate\n\n\nRate2 =\n    Rate * 2;");
    expect(guard(files)).toEqual([]);
    expect(rewritten(files)).toEqual([UNMANAGED]);
  });

  it("a file the pull does not write (no names in it) is listed; a file with an edit is the edit's group", () => {
    const files = project();
    files["names/Notes.xln"] = "// notes\n";
    edit(files, S1, "@C8 = Rate*3;", "// changed\n@C8 = Rate*4;");
    expect(guard(files)).toEqual(["names/sheets/S1.xln:20  S1!C8: set formula of S1!C8"]);
    expect(rewritten(files)).toEqual(["names/Notes.xln"]);
  });
});
