import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { countElements, dirName, formatReport, lockFileName, probeXlsx } from "../src/inspect.js";

const RESULTS = join(import.meta.dirname, "..", "..", "..", "probes", "results");
const workbooks = readdirSync(RESULTS).filter((f) => f.endsWith(".xlsx"));

describe("browser probe logic", () => {
  it("counts <definedName> but not <definedNames>", () => {
    const xml = '<definedNames><definedName name="a">1</definedName>\n<definedName\tname="b"/></definedNames>';
    expect(countElements(xml, "definedName")).toBe(2);
    expect(countElements("<definedNames/>", "definedName")).toBe(0);
  });

  it("names the owner file and the folder", () => {
    expect(lockFileName("model.xlsx")).toBe("~$model.xlsx");
    expect(dirName("/a/b/model.xlsx")).toBe("/a/b");
    expect(dirName("/model.xlsx")).toBe("/");
  });

  it.each(workbooks)("unzips %s", (name) => {
    const probe = probeXlsx(new Uint8Array(readFileSync(join(RESULTS, name))));
    const parts = probe.entries.map((e) => e.name);
    expect(parts).toContain("[Content_Types].xml");
    expect(parts).toContain("xl/workbook.xml");
    expect(probe.definedNames).toBeGreaterThan(0);
    const text = formatReport({ ...probe, file: name, bytes: 1, lockFile: lockFileName(name), lockFilePresent: false, ownerFilesInFolder: [] });
    expect(text).toContain("xl/workbook.xml");
  });
});
