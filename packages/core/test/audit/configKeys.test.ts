// xln.config.json keys (feedback 2026-10-07): the author wrote "library" inside "audit",
// and it was ignored with a note that did not say where it belongs.
import { describe, expect, it } from "vitest";
import { CONFIG_KEYS, configKeyRange, defaultConfigText, parseConfig } from "../../src/index.js";

describe("xln.config.json keys", () => {
  it("a top-level setting written inside audit: where it belongs, and nothing taken from it", () => {
    const text = JSON.stringify({ audit: { harness: [], rules: {}, constants: { allow: [] }, library: "../lib" } }, null, 2);
    const r = parseConfig(text);
    expect(r.config.library).toBeUndefined();
    expect(r.problems).toEqual(['audit.library: `library` is a top-level setting, not an audit setting: move it out of "audit" (ignored here)']);
    expect(r.issues).toEqual([{ kind: "problem", path: ["audit", "library"], message: r.problems[0] }]);
  });

  it("other misplaced keys: into their section, or out of it", () => {
    const r = parseConfig(JSON.stringify({ harness: ["X"], embed: true, audit: { embed: true, allow: [1], constants: { library: "x" } }, build: { rules: {} } }));
    expect(r.problems).toEqual([
      'harness: `harness` is an audit setting, not a top-level setting: move it into "audit" (ignored here)',
      'embed: `embed` is a build setting, not a top-level setting: move it into "build" (ignored here)',
      'build.rules: `rules` is an audit setting, not a build setting: move it into "audit" (ignored here)',
      'audit.embed: `embed` is a build setting, not an audit setting: move it into "build" (ignored here)',
      'audit.allow: `allow` is an audit.constants setting, not an audit setting: move it into "constants" (ignored here)',
      'audit.constants.library: `library` is a top-level setting, not an audit.constants setting: move it out of "constants" (ignored here)',
    ]);
  });

  it("unknown keys: a typo is named, otherwise the known keys are listed; the top level counts too", () => {
    const r = parseConfig(JSON.stringify({ libary: "../lib", colour: 1, audit: { rule: {} } }));
    expect(r.problems).toEqual([
      'libary: unknown setting (ignored); did you mean "library"?',
      "colour: unknown setting (ignored); the settings are audit, build, library",
      'audit.rule: unknown setting (ignored); did you mean "rules"?',
    ]);
  });

  it("the known keys are listed once, from the code; the default file and a full one are clean", () => {
    expect(Object.keys(CONFIG_KEYS).sort()).toEqual(["allow", "audit", "build", "constants", "embed", "harness", "library", "rules", "sentinelAbove"]);
    expect(parseConfig(defaultConfigText()).issues).toEqual([]);
    const full = { audit: { harness: ["Check!*"], rules: { C13: "off" }, constants: { allow: [4], sentinelAbove: 1e90 } }, build: { embed: true }, library: "../lib" };
    expect(parseConfig(JSON.stringify(full)).issues).toEqual([]);
  });

  it("finds a key in the text, nested, with strings and arrays before it", () => {
    const text = '{\n  "x": ["a\\"}", {"library": 1}],\n  "audit": {\n    "harness": [],\n    "library": "../lib"\n  }\n}\n';
    const r = configKeyRange(text, ["audit", "library"])!;
    expect(text.slice(r.start, r.end)).toBe('"library"');
    expect(text.slice(0, r.start).split("\n").length).toBe(5);
    expect(configKeyRange(text, ["audit", "nope"])).toBeUndefined();
    expect(configKeyRange("{ not json", ["audit"])).toBeUndefined();
  });
});
