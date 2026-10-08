// A project's audit settings, `xln.config.json` in the project folder (`lbo.xln/`):
//
//   { "audit": { "harness": ["Check!*", "CHK.*"],
//                "rules": { "C13": "off", "C10.unused": "info" },
//                "constants": { "allow": [4, 1.5], "sentinelAbove": 1e90 } },
//     "build": { "embed": true },
//     "library": "../../_shared/lib" }
//
// `build.embed` (D5, opt-in since 2026-10-05): every build embeds the project source in the
// workbook, as `xln build --embed` does; `--no-embed` overrides it.
//
// `names.scope` ("source" | "excel", M3d) is gone (author's decision, 2026-10-06): one
// editing environment is active at a time, and each sync carries a scope change made in
// either to the other. An old file that has it is read with a note; the setting is ignored.
//
// `library` (M4): the folder of `.lambda` files `xln lib` compares with, relative to the
// project folder or absolute (`~/…` is expanded by the CLI and the extension).
//
// A key that is not a setting where it is written is flagged, never silently dropped: a
// known setting in the wrong section says where it goes (`audit.library` → "move it out
// of audit"), a near miss names the setting meant, anything else lists the known keys
// (`CONFIG_KEYS`). The CLI prints these as notes; the editor marks the key.
//
// The core reads text and returns options; the CLI and the extension read the file.

import { DEFAULT_CONSTANTS } from "./audit.js";
import { RULES } from "./rules.js";
import { CHECK_IDS, type AuditOptions, type AuditSeverity } from "./types.js";

export const CONFIG_FILE = "xln.config.json";

export interface AuditSettings {
  /** Glob patterns on name keys: the check harness (C10 skips it, the C8 tiers leave it out). */
  harness?: string[];
  /** Severity per rule id (`C10.unused`) or check (`C13`), or `off`. */
  rules?: Record<string, AuditSeverity | "off">;
  /** C13: numbers allowed besides `DEFAULT_CONSTANTS`; the sentinel magnitude. */
  constants?: { allow?: number[]; sentinelAbove?: number };
}

export interface BuildSettings {
  /** D5: embed the project source in the workbook on every build (default false). */
  embed?: boolean;
}

export interface ProjectConfig {
  audit?: AuditSettings;
  build?: BuildSettings;
  /**
   * M4: the library folder (`.lambda` files), relative to the project folder or absolute,
   * as written: the CLI and the extension resolve it (and expand `~`), not the core.
   */
  library?: string;
}

/** What `xln pull` writes when the project has no `xln.config.json` yet. */
export function defaultConfigText(): string {
  const config: ProjectConfig = { audit: { harness: [], rules: {}, constants: { allow: [] } } };
  return JSON.stringify(config, null, 2) + "\n";
}

const SEVERITIES = new Set(["off", "error", "warning", "info"]);

function isObject(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

/** The note on an old file's `names.scope`. */
export const NAMES_SCOPE_NOTE = "names.scope: no longer a setting (ignored): a scope set in Excel or in the source is carried to the other side by the next pull or build";

/** Something in `xln.config.json` that could not be used (`problem`) or that is said for information (`note`), with the key path it is about. */
export interface ConfigIssue {
  kind: "problem" | "note";
  /** The keys from the top (`["audit", "library"]`); empty for the file as a whole. */
  path: string[];
  message: string;
}

/**
 * Every setting, where it belongs (its parent keys). The source of truth for the known
 * keys: a key elsewhere is flagged, with where it goes when it belongs somewhere else
 * (feedback 2026-10-07: `"library"` written inside `"audit"` was silently ignored).
 */
export const CONFIG_KEYS: Readonly<Record<string, readonly string[]>> = {
  audit: [],
  build: [],
  library: [],
  harness: ["audit"],
  rules: ["audit"],
  constants: ["audit"],
  allow: ["audit", "constants"],
  sentinelAbove: ["audit", "constants"],
  embed: ["build"],
};

/** The keys known at a place (`[]`: the top level). `names` (gone) is handled on its own. */
function keysAt(path: readonly string[]): string[] {
  return Object.entries(CONFIG_KEYS)
    .filter(([, home]) => home.join(".") === path.join("."))
    .map(([k]) => k);
}

function placeName(path: readonly string[]): string {
  return path.length === 0 ? "a top-level setting" : `${/^[aeiou]/.test(path.join(".")) ? "an" : "a"} ${path.join(".")} setting`;
}

/** Edit distance at most 1 (a typo), case ignored. */
function nearMiss(a: string, b: string): boolean {
  const x = a.toLowerCase();
  const y = b.toLowerCase();
  if (x === y) return true;
  if (Math.abs(x.length - y.length) > 1) return false;
  let i = 0;
  while (i < x.length && i < y.length && x[i] === y[i]) i++;
  if (x.length === y.length) return x.slice(i + 1) === y.slice(i + 1) || (x[i] === y[i + 1] && x[i + 1] === y[i] && x.slice(i + 2) === y.slice(i + 2));
  return x.length < y.length ? x.slice(i) === y.slice(i + 1) : x.slice(i + 1) === y.slice(i);
}

/** Why `key` under `path` is not used: it belongs elsewhere, is a typo, or is unknown. */
export function unknownKeyMessage(path: readonly string[], key: string): string {
  const at = [...path, key].join(".");
  const home = CONFIG_KEYS[key];
  if (home !== undefined) {
    const here = path.join(".");
    const there = home.join(".");
    let move: string;
    if (home.length < path.length && here.startsWith(there)) move = `move it out of "${path[path.length - 1]}"`;
    else if (home.length > path.length && (here === "" || there.startsWith(here + "."))) move = `move it into "${home.slice(path.length).join(".")}"`;
    else move = home.length === 0 ? "move it to the top level" : `move it into "${there}"`;
    return `${at}: \`${key}\` is ${placeName(home)}, not ${placeName(path)}: ${move} (ignored here)`;
  }
  const known = keysAt(path);
  const near = known.find((k) => nearMiss(k, key));
  if (near) return `${at}: unknown setting (ignored); did you mean "${near}"?`;
  return `${at}: unknown setting (ignored); ${path.length === 0 ? "the settings are" : `the ${path.join(".")} settings are`} ${known.join(", ")}`;
}

/**
 * Reads `xln.config.json` text. Never throws: what cannot be used is left out and said in
 * `problems` (`audit.rules.C99: no such rule or check`); `notes` are for information (a
 * setting that no longer exists). `issues` are both, with the key path each is about (for
 * a mark on the key in the editor, `configKeyRange`). A key that is not a setting where it
 * is written is a problem too, saying where it belongs (`CONFIG_KEYS`).
 */
export function parseConfig(text: string): { config: ProjectConfig; problems: string[]; notes: string[]; issues: ConfigIssue[] } {
  const issues: ConfigIssue[] = [];
  const problem = (path: string[], message: string): void => void issues.push({ kind: "problem", path, message });
  const done = (config: ProjectConfig) => ({
    config,
    problems: issues.filter((i) => i.kind === "problem").map((i) => i.message),
    notes: issues.filter((i) => i.kind === "note").map((i) => i.message),
    issues,
  });
  let raw: unknown;
  try {
    raw = JSON.parse(text);
  } catch (e) {
    problem([], `not JSON: ${e instanceof Error ? e.message : String(e)}`);
    return done({});
  }
  if (!isObject(raw)) {
    problem([], "the file must hold a JSON object");
    return done({});
  }
  const config: ProjectConfig = {};
  for (const k of Object.keys(raw)) if (k !== "names" && !keysAt([]).includes(k)) problem([k], unknownKeyMessage([], k));
  const b = raw["build"];
  if (b !== undefined) {
    if (!isObject(b)) problem(["build"], "build: must be an object");
    else {
      const build: BuildSettings = {};
      for (const k of Object.keys(b)) if (k !== "embed") problem(["build", k], unknownKeyMessage(["build"], k));
      if (b["embed"] !== undefined) {
        if (typeof b["embed"] === "boolean") build.embed = b["embed"];
        else problem(["build", "embed"], "build.embed: must be true or false");
      }
      config.build = build;
    }
  }
  const lib = raw["library"];
  if (lib !== undefined) {
    if (typeof lib === "string" && lib.trim() !== "") config.library = lib.trim();
    else problem(["library"], "library: must be the path of the library folder (a string)");
  }
  const nm = raw["names"];
  if (nm !== undefined) {
    if (!isObject(nm)) problem(["names"], "names: unknown setting (ignored)");
    else
      for (const k of Object.keys(nm)) {
        if (k === "scope") issues.push({ kind: "note", path: ["names", "scope"], message: NAMES_SCOPE_NOTE });
        else problem(["names", k], `names.${k}: unknown setting (ignored)`);
      }
  }
  const a = raw["audit"];
  if (a === undefined) return done(config);
  if (!isObject(a)) {
    problem(["audit"], "audit: must be an object");
    return done(config);
  }
  const audit: AuditSettings = {};
  for (const k of Object.keys(a)) if (k !== "harness" && k !== "rules" && k !== "constants") problem(["audit", k], unknownKeyMessage(["audit"], k));

  const h = a["harness"];
  if (h !== undefined) {
    if (Array.isArray(h) && h.every((x) => typeof x === "string")) audit.harness = (h as string[]).map((x) => x.trim()).filter(Boolean);
    else problem(["audit", "harness"], "audit.harness: must be a list of glob patterns (strings)");
  }

  const r = a["rules"];
  if (r !== undefined) {
    if (!isObject(r)) problem(["audit", "rules"], "audit.rules: must be an object of rule or check → off, info, warning or error");
    else {
      const rules: Record<string, AuditSeverity | "off"> = {};
      for (const [id, v] of Object.entries(r)) {
        const known = id in RULES || (CHECK_IDS as readonly string[]).includes(id);
        if (!known) problem(["audit", "rules", id], `audit.rules.${id}: no such rule or check (ignored)`);
        else if (typeof v !== "string" || !SEVERITIES.has(v)) problem(["audit", "rules", id], `audit.rules.${id}: must be off, info, warning or error (ignored)`);
        else rules[id] = v as AuditSeverity | "off";
      }
      audit.rules = rules;
    }
  }

  const c = a["constants"];
  if (c !== undefined) {
    if (!isObject(c)) problem(["audit", "constants"], "audit.constants: must be an object");
    else {
      const constants: NonNullable<AuditSettings["constants"]> = {};
      const allow = c["allow"];
      if (allow !== undefined) {
        if (Array.isArray(allow) && allow.every((x) => typeof x === "number" && Number.isFinite(x))) constants.allow = allow as number[];
        else problem(["audit", "constants", "allow"], "audit.constants.allow: must be a list of numbers");
      }
      const s = c["sentinelAbove"];
      if (s !== undefined) {
        if (typeof s === "number" && s > 0) constants.sentinelAbove = s;
        else problem(["audit", "constants", "sentinelAbove"], "audit.constants.sentinelAbove: must be a positive number");
      }
      for (const k of Object.keys(c)) if (k !== "allow" && k !== "sentinelAbove") problem(["audit", "constants", k], unknownKeyMessage(["audit", "constants"], k));
      audit.constants = constants;
    }
  }
  config.audit = audit;
  return done(config);
}

/**
 * Where a key is written in `xln.config.json` text: the offsets of the key's string
 * (quotes included) at `path`; the start of the file for `[]`; undefined when the text is
 * not JSON enough to tell. A small scanner over the JSON grammar, not a pattern.
 */
export function configKeyRange(text: string, path: readonly string[]): { start: number; end: number } | undefined {
  if (path.length === 0) return { start: 0, end: Math.min(1, text.length) };
  let i = 0;
  const ws = (): void => {
    while (i < text.length && " \t\r\n".includes(text[i]!)) i++;
  };
  const str = (): string | undefined => {
    if (text[i] !== '"') return undefined;
    const start = i++;
    while (i < text.length && text[i] !== '"') i += text[i] === "\\" ? 2 : 1;
    if (i >= text.length) return undefined;
    i++;
    try {
      return JSON.parse(text.slice(start, i)) as string;
    } catch {
      return undefined;
    }
  };
  // Skips one value; false when the text is not JSON.
  const skip = (): boolean => {
    ws();
    const ch = text[i];
    if (ch === '"') return str() !== undefined;
    if (ch === "{" || ch === "[") {
      const close = ch === "{" ? "}" : "]";
      i++;
      ws();
      if (text[i] === close) {
        i++;
        return true;
      }
      for (;;) {
        ws();
        if (ch === "{") {
          if (str() === undefined) return false;
          ws();
          if (text[i++] !== ":") return false;
        }
        if (!skip()) return false;
        ws();
        if (text[i] === ",") i++;
        else if (text[i] === close) {
          i++;
          return true;
        } else return false;
      }
    }
    const start = i;
    while (i < text.length && !",}] \t\r\n".includes(text[i]!)) i++;
    return i > start;
  };
  // Finds `path[depth]` in the object at `i`.
  const find = (depth: number): { start: number; end: number } | undefined => {
    ws();
    if (text[i] !== "{") return undefined;
    i++;
    for (;;) {
      ws();
      if (text[i] === "}") return undefined;
      const start = i;
      const key = str();
      if (key === undefined) return undefined;
      const end = i;
      ws();
      if (text[i++] !== ":") return undefined;
      if (key === path[depth]) {
        if (depth === path.length - 1) return { start, end };
        ws();
        return find(depth + 1);
      }
      if (!skip()) return undefined;
      ws();
      if (text[i] === ",") i++;
      else return undefined;
    }
  };
  return find(0);
}

/**
 * Audit options from a project's settings, then `overrides` on top (command-line flags win:
 * a `harness` or a rule given there replaces the project's). `constants.allow` adds to
 * `DEFAULT_CONSTANTS`.
 */
export function auditOptions(settings: AuditSettings | undefined, overrides: AuditOptions = {}): AuditOptions {
  const out: AuditOptions = {};
  const s = settings ?? {};
  if (s.harness && s.harness.length > 0) out.harness = s.harness;
  if (s.rules && Object.keys(s.rules).length > 0) out.rules = { ...s.rules };
  if (s.constants && (s.constants.allow?.length || s.constants.sentinelAbove !== undefined)) {
    const constants: NonNullable<AuditOptions["constants"]> = {};
    if (s.constants.allow?.length) constants.allow = [...DEFAULT_CONSTANTS, ...s.constants.allow];
    if (s.constants.sentinelAbove !== undefined) constants.sentinelAbove = s.constants.sentinelAbove;
    out.constants = constants;
  }
  const merged: AuditOptions = { ...out, ...overrides };
  if (out.rules && overrides.rules) merged.rules = { ...out.rules, ...overrides.rules };
  if (out.constants && overrides.constants) merged.constants = { ...out.constants, ...overrides.constants };
  return merged;
}

/** D5: whether a build embeds the source: the flag when given, else the project's `build.embed`, else no. */
export function embedSetting(flag: boolean | undefined, configText: string | undefined): boolean {
  if (flag !== undefined) return flag;
  if (configText === undefined) return false;
  return parseConfig(configText).config.build?.embed === true;
}
