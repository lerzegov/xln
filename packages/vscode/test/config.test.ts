import { describe, expect, it } from "vitest";
import { configMarks } from "../src/model/config.js";

describe("marks in xln.config.json (feedback 2026-10-07)", () => {
  it("a misplaced key is a warning on the key, saying where it belongs; an old setting is information", () => {
    const text = '{\n  "audit": {\n    "harness": [],\n    "library": "../lib"\n  },\n  "names": { "scope": "excel" }\n}\n';
    const marks = configMarks(text);
    expect(marks.map((m) => [m.severity, m.start, m.end])).toEqual([
      ["info", { line: 5, character: 13 }, { line: 5, character: 20 }],
      ["warning", { line: 3, character: 4 }, { line: 3, character: 13 }],
    ]);
    expect(marks[1]!.message).toBe('audit.library: `library` is a top-level setting, not an audit setting: move it out of "audit" (ignored here)');
  });

  it("a bad value marks its key; text that is not JSON marks the start", () => {
    expect(configMarks('{ "build": { "embed": "yes" } }').map((m) => [m.start.character, m.end.character, m.message])).toEqual([[13, 20, "build.embed: must be true or false"]]);
    expect(configMarks("{ nope").map((m) => [m.start, m.severity])).toEqual([[{ line: 0, character: 0 }, "warning"]]);
    expect(configMarks('{ "library": "../lib" }')).toEqual([]);
  });
});
