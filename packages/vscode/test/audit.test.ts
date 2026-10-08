// Where audit findings land in a pulled project (the editor's diagnostics), on the traps workbook.
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { audit, MANIFEST_FILE, pullProject, readWorkbook, type Finding } from "@xln/core";
import { describe, expect, it } from "vitest";
import { findingLoc, problemMessage, problemSeverity } from "../src/model/audit.js";
import { parseManifest } from "../src/model/manifest.js";
import { isNamesFile, Project } from "../src/model/project.js";

const traps = fileURLToPath(new URL("../../../probes/fixtures/traps.xlsx", import.meta.url));

describe("audit findings in a pulled project", () => {
  const bytes = new Uint8Array(readFileSync(traps));
  const pulled = pullProject(bytes, "traps.xlsx");
  const p = new Project("mem:/p", parseManifest(pulled.files[MANIFEST_FILE]!));
  for (const [path, text] of Object.entries(pulled.files)) if (isNamesFile(path)) p.setFile(path, text);
  const report = audit(readWorkbook(bytes), { workbook: "traps.xlsx" });
  const find = (rule: string): Finding => report.findings.find((f) => f.rule === rule)!;
  const text = (f: Finding) => {
    const l = findingLoc(p, f)!;
    return `${l.path}:${p.files.get(l.path)!.text.slice(l.start, l.end)}`;
  };

  it("lands on the part of the formula the finding is about", () => {
    expect(text(find("C2.bare-prefix"))).toBe("names/_unmanaged.xln:SEQUENCE");
    expect(text(find("C2.poisoned"))).toBe("names/_unmanaged.xln:_xludf.SEQUENCE");
    expect(text(find("C5.other-sheet"))).toBe("names/_unmanaged.xln:OnlySecond");
    expect(text(find("C4.ref-deleted"))).toBe("names/_unmanaged.xln:'S1'!#REF!");
    expect(text(find("C13.constant"))).toBe("names/_unmanaged.xln:0.27");
  });

  it("lands on the name when there is no part to point at, and nowhere for cells", () => {
    expect(text(find("C10.unused"))).toBe("names/_unmanaged.xln:Fact");
    expect(text(find("C12.name-cycle"))).toBe("names/_unmanaged.xln:CycA");
    expect(findingLoc(p, find("C6.lambda-arity"))).toBeUndefined();
  });

  it("words the diagnostic", () => {
    const f = find("C5.other-sheet");
    expect(problemMessage(f)).toBe("C5 OnlySecond exists only on 'S2', not at workbook scope: unqualified it is #NAME? (write 'S2'!OnlySecond)");
    expect(problemSeverity(f)).toBe("error");
    expect(problemSeverity(find("C13.constant"))).toBe("information");
  });
});
