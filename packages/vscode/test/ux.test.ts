// M3e: the editor's housekeeping, vscode-free (src/model/ux.ts).
import { describe, expect, it } from "vitest";
import { internalsExclude, internalsHidden, INTERNAL_GLOBS, projectChoices, rootOf, staleTabs } from "../src/model/ux.js";

const A = "file:///w/book/a.xln";
const B = "file:///w/other/b.xln";

describe("stale tabs (after a pull, on loading a project)", () => {
  it("closes tabs of missing files below a project; keeps a dirty one; leaves everything else", () => {
    const gone = new Set([`${A}/names/Old.xln`, `${A}/names/Draft.xln`, "file:///w/elsewhere/gone.txt", `${B}/names/X.xln`]);
    const r = staleTabs(
      [
        { uri: `${A}/names/Old.xln`, dirty: false },
        { uri: `${A}/names/Draft.xln`, dirty: true },
        { uri: `${A}/names/FN.xln`, dirty: false },
        { uri: "file:///w/elsewhere/gone.txt", dirty: false },
        { uri: `${B}/names/X.xln`, dirty: false },
      ],
      [A],
      (u) => !gone.has(u),
    );
    expect(r).toEqual({ close: [`${A}/names/Old.xln`], keep: [`${A}/names/Draft.xln`] });
  });
});

describe("project internals (files.exclude)", () => {
  it("the manifest, the lockfile and the kept library bases, not xln.config.json", () => {
    expect(INTERNAL_GLOBS).toEqual(["**/*.xln/workbook.manifest.json", "**/*.xln/xln.lock.json", "**/*.xln/library-bases"]);
  });

  it("show: false over the default, other patterns kept; hide: the entries go, the default applies", () => {
    const shown = internalsExclude({ "**/.git": true }, true);
    expect(shown).toEqual({ changed: true, value: { "**/.git": true, "**/*.xln/workbook.manifest.json": false, "**/*.xln/xln.lock.json": false, "**/*.xln/library-bases": false } });
    expect(internalsExclude(shown.value, true).changed).toBe(false);
    expect(internalsExclude(shown.value, false)).toEqual({ changed: true, value: { "**/.git": true } });
    // Nothing else in the workspace setting: it is removed.
    expect(internalsExclude(internalsExclude(undefined, true).value, false)).toEqual({ changed: true, value: undefined });
    expect(internalsExclude(undefined, false)).toEqual({ changed: false, value: undefined });
  });

  it("hidden when every internal pattern is true in the effective setting", () => {
    expect(internalsHidden({ "**/*.xln/workbook.manifest.json": true, "**/*.xln/xln.lock.json": true, "**/*.xln/library-bases": true })).toBe(true);
    expect(internalsHidden({ "**/*.xln/workbook.manifest.json": true, "**/*.xln/xln.lock.json": true })).toBe(false);
    expect(internalsHidden({ "**/*.xln/workbook.manifest.json": true, "**/*.xln/xln.lock.json": false })).toBe(false);
    expect(internalsHidden(undefined)).toBe(false);
  });
});

describe("which project (New module)", () => {
  it("the active editor's project is offered first; otherwise the list as it is", () => {
    expect(projectChoices([A, B], B)).toEqual([B, A]);
    expect(projectChoices([A, B], undefined)).toEqual([A, B]);
    expect(projectChoices([A, B], "file:///w/elsewhere.xln")).toEqual([A, B]);
  });

  it("an Explorer item: the project folder itself, its names folder, a file below it", () => {
    expect(rootOf(A, [A, B])).toBe(A);
    expect(rootOf(A + "/", [A, B])).toBe(A);
    expect(rootOf(A + "/names", [A, B])).toBe(A);
    expect(rootOf(B + "/names/sheets/IS.xln", [A, B])).toBe(B);
    expect(rootOf("file:///w/book/a.xlnx/names", [A, B])).toBeUndefined();
    expect(rootOf("file:///w/book/a.xlsx", [A, B])).toBeUndefined();
  });
});
