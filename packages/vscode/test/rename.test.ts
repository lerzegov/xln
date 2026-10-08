// Rename Symbol (F2, M5) in the editor's project model: what the cursor may rename, and the
// edits (the same core function as `xln rename`), which the provider shows as a preview.
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { LOCK_FILE, MANIFEST_FILE, parseLockfile, pullProject, type SourceRename } from "@xln/core";
import { describe, expect, it } from "vitest";
import { parseManifest } from "../src/model/manifest.js";
import { isNamesFile, Project } from "../src/model/project.js";

const f7 = fileURLToPath(new URL("../../../probes/results/f7_base.xlsx", import.meta.url));
const U = "names/_unmanaged.xln";
const S1 = "names/sheets/S1.xln";

function project(): Project {
  const files = pullProject(new Uint8Array(readFileSync(f7)), "f7_base.xlsx").files;
  const p = new Project("mem:/p", parseManifest(files[MANIFEST_FILE]!));
  for (const [path, text] of Object.entries(files)) if (isNamesFile(path)) p.setFile(path, text);
  p.lock = parseLockfile(files[LOCK_FILE]!);
  return p;
}

const offsetOf = (p: Project, path: string, needle: string, delta = 1) => p.files.get(path)!.text.indexOf(needle) + delta;

describe("Rename Symbol on f7_base", () => {
  const p = project();

  it("prepares on a name's statement and on a use, refuses a LET variable and other text", () => {
    const self = p.renameTarget(U, offsetOf(p, U, "Rate = 0.1"));
    expect(typeof self === "object" && [self.def.key, p.files.get(U)!.text.slice(self.at.start, self.at.end)]).toEqual(["Rate", "Rate"]);
    const use = p.renameTarget(S1, offsetOf(p, S1, "A2*Rate", 4));
    expect(typeof use === "object" && [use.def.key, p.files.get(S1)!.text.slice(use.at.start, use.at.end)]).toEqual(["Rate", "Rate"]);
    expect(p.renameTarget(S1, offsetOf(p, S1, "LET(Rate", 5))).toMatch(/LET or LAMBDA variable/);
    expect(p.renameTarget(S1, offsetOf(p, S1, '"Rate is', 2))).toMatch(/put the cursor on one/);
  });

  it("the edits: the name, @renamed and every reader; refusals as messages", () => {
    const r = p.renameEdits("Rate", "Pace") as SourceRename;
    expect(r.edits.map((e) => e.label).filter((l) => l !== "reference")).toEqual(["annotation", "name"]);
    expect(r.references).toBe(10);
    expect(r.files[S1]).toContain("@C3 = LET(Rate, 5, Rate*2);");
    expect(p.renameEdits("Rate", "RateX")).toMatch(/exists already/);
    expect(p.renameEdits("Rate", "x")).toMatch(/would change what this formula reads/);
  });
});
