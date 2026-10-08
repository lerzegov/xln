// Loads the function catalogue table (catalogue-data.ts), joins in the parameter names
// (catalogue-params.ts) and answers lookups.

import { CATALOGUE_TABLE } from "./catalogue-data.js";
import { PARAMS_TABLE } from "./catalogue-params.js";
import { XLM_FUNCTIONS } from "./xlm-data.js";

export type StoredPrefix = "" | "_xlfn." | "_xlfn._xlws.";

export interface FunctionInfo {
  /** Upper-case English name, as Excel displays it. */
  name: string;
  /** Prefix in the stored (file) form. */
  prefix: StoredPrefix;
  minArgs: number;
  /** 255 for variadic functions. */
  maxArgs: number;
  /** First Excel version: "2007" (or earlier), "2010", …, "2024", "365". */
  since: string;
  /** The stored prefix is not confirmed by a saved file or by Microsoft's list. */
  prefixUnsure: boolean;
  /** Written by Excel (SINGLE for `@`, ANCHORARRAY for `#`), not typed by users. */
  internal: boolean;
  /**
   * Parameter names as Microsoft's function reference gives them, for signature help
   * (M3c); optional ones in brackets (`[if_not_found]`), a repeating tail with `...`
   * (`number2, ...`). Absent until the catalogue has them for this function.
   *
   * One entry per argument position; `...` is an entry of its own. The group that
   * repeats is the run of entries before `...` back to (not including) the earlier
   * entry with the same name up to its trailing digits: `criteria_range1, criteria1,
   * [criteria_range2], [criteria2], ...` repeats two entries because `criteria2`
   * comes two places after `criteria1`; without such an earlier entry the group is
   * the one entry before `...`. Entries after `...` (LAMBDA's `calculation`) are the
   * last arguments. `paramIndex` maps an argument position to an entry.
   */
  params?: string[];
  /** The parameter names are not confirmed by Microsoft's page (see catalogue-params.ts). */
  paramsUnsure?: boolean;
}

export const MAX_ARGS = 255;

const PREFIXES: Record<string, StoredPrefix> = { "-": "", xlfn: "_xlfn.", xlws: "_xlfn._xlws." };
const VERSIONS = new Set(["2007", "2010", "2013", "2016", "2019", "2021", "2024", "365"]);

/** Parses the table; throws on a malformed line so a bad edit fails the tests at once. */
export function parseCatalogue(table: string): Map<string, FunctionInfo> {
  const out = new Map<string, FunctionInfo>();
  table.split("\n").forEach((raw, k) => {
    const line = raw.trim();
    if (line === "" || line.startsWith("#")) return;
    const cols = line.split(/\s+/);
    const where = `catalogue line ${k + 1} (${line})`;
    if (cols.length < 5 || cols.length > 6) throw new Error(`${where}: expected 5 or 6 columns`);
    const [name, pfx, min, max, since, flags = ""] = cols as [string, string, string, string, string, string?];
    if (name !== name.toUpperCase()) throw new Error(`${where}: name must be upper case`);
    const prefix = PREFIXES[pfx];
    if (prefix === undefined) throw new Error(`${where}: prefix must be -, xlfn or xlws`);
    const minArgs = Number(min);
    const maxArgs = max === "*" ? MAX_ARGS : Number(max);
    if (!Number.isInteger(minArgs) || !Number.isInteger(maxArgs) || minArgs < 0 || maxArgs < minArgs) {
      throw new Error(`${where}: bad arity`);
    }
    if (!VERSIONS.has(since)) throw new Error(`${where}: unknown version ${since}`);
    const flagSet = new Set(flags ? flags.split(",") : []);
    for (const f of flagSet) if (f !== "?prefix" && f !== "internal") throw new Error(`${where}: unknown flag ${f}`);
    if (out.has(name)) throw new Error(`${where}: duplicate`);
    out.set(name, {
      name,
      prefix,
      minArgs,
      maxArgs,
      since,
      prefixUnsure: flagSet.has("?prefix"),
      internal: flagSet.has("internal"),
    });
  });
  return out;
}

export interface ParamsEntry {
  params: string[];
  /** Not confirmed by a Microsoft page (`?`). */
  unsure: boolean;
  /** Microsoft's brackets disagree with the catalogue's arity, knowingly (`!arity`). */
  arityDiffers: boolean;
}

// Microsoft's names: letters, digits, `_`, `.`, and the apostrophe of `known_y's`.
const PARAM = /^(\[[A-Za-z][A-Za-z0-9_.']*\]|[A-Za-z][A-Za-z0-9_.']*|\.\.\.)$/;

/**
 * Parses the parameter table (catalogue-params.ts); throws on a malformed line, on a
 * name missing from `functions`, and on a list that disagrees with the arity there
 * unless the line says so with `!arity`.
 */
export function parseParams(table: string, functions: ReadonlyMap<string, FunctionInfo>): Map<string, ParamsEntry> {
  const out = new Map<string, ParamsEntry>();
  table.split("\n").forEach((raw, k) => {
    const line = raw.trim();
    if (line === "" || line.startsWith("#")) return;
    const where = `params line ${k + 1} (${line})`;
    const open = line.indexOf("(");
    const close = line.lastIndexOf(")");
    if (open <= 0 || close < open) throw new Error(`${where}: expected NAME(params)`);
    const name = line.slice(0, open);
    const flags = line.slice(close + 1).trim().split(/\s+/).filter((f) => f !== "");
    for (const f of flags) if (f !== "?" && f !== "!arity") throw new Error(`${where}: unknown flag ${f}`);
    const info = functions.get(name);
    if (!info) throw new Error(`${where}: ${name} is not in the catalogue`);
    if (info.internal) throw new Error(`${where}: ${name} is internal`);
    if (out.has(name)) throw new Error(`${where}: duplicate`);
    const inner = line.slice(open + 1, close).trim();
    const params = inner === "" ? [] : inner.split(",").map((p) => p.trim());
    for (const p of params) if (!PARAM.test(p)) throw new Error(`${where}: bad parameter \`${p}\``);
    if (params.filter((p) => p === "...").length > 1) throw new Error(`${where}: more than one \`...\``);
    if (params[0] === "...") throw new Error(`${where}: \`...\` repeats nothing`);
    const problem = arityProblem(params, info);
    const arityDiffers = flags.includes("!arity");
    if (problem && !arityDiffers) throw new Error(`${where}: ${problem}`);
    if (!problem && arityDiffers) throw new Error(`${where}: \`!arity\` but the arity agrees`);
    out.set(name, { params, unsure: flags.includes("?"), arityDiffers });
  });
  return out;
}

const baseName = (p: string): string => p.replace(/^\[|\]$/g, "").replace(/\d+$/, "");

/** Length of the group that `...` repeats; 0 when there is no `...`. */
function repeatLength(params: readonly string[]): number {
  const dots = params.indexOf("...");
  if (dots < 1) return 0;
  const last = baseName(params[dots - 1]!);
  for (let i = dots - 2; i >= 0; i--) if (baseName(params[i]!) === last) return dots - 1 - i;
  return 1;
}

/** Why `params` does not fit the catalogue's arity; undefined when it fits. */
export function arityProblem(params: readonly string[], info: Pick<FunctionInfo, "minArgs" | "maxArgs">): string | undefined {
  const dots = params.filter((p) => p === "...").length;
  const args = params.filter((p) => p !== "...");
  const required = args.filter((p) => !p.startsWith("[")).length;
  if (required !== info.minArgs) return `${required} required parameters, but MIN is ${info.minArgs}`;
  if (dots > 0) {
    if (info.maxArgs !== MAX_ARGS) return `\`...\` but MAX is ${info.maxArgs}`;
  } else if (args.length !== info.maxArgs) {
    return `${args.length} parameters, but MAX is ${info.maxArgs === MAX_ARGS ? "*" : info.maxArgs}`;
  }
  return undefined;
}

/**
 * The entry of `params` that names argument `argIndex` (0-based), for signature help;
 * -1 when there is none. With `argCount`, the entries after `...` (LAMBDA's
 * `calculation`) name the last arguments.
 */
export function paramIndex(params: readonly string[], argIndex: number, argCount?: number): number {
  const dots = params.indexOf("...");
  if (dots < 0) return argIndex < params.length ? argIndex : -1;
  const tail = params.length - dots - 1;
  const len = repeatLength(params);
  if (argCount !== undefined && tail > 0) {
    const fromEnd = argCount - 1 - argIndex;
    if (fromEnd >= 0 && fromEnd < tail) return params.length - 1 - fromEnd;
  }
  if (argIndex < dots) return argIndex;
  if (len === 0) return -1;
  return dots - len + ((argIndex - dots) % len);
}

let cache: Map<string, FunctionInfo> | undefined;

/** All functions, keyed by upper-case name, with their parameter names joined in. */
export function catalogue(): ReadonlyMap<string, FunctionInfo> {
  if (!cache) {
    const functions = parseCatalogue(CATALOGUE_TABLE);
    for (const [name, entry] of parseParams(PARAMS_TABLE, functions)) {
      const info = functions.get(name)!;
      info.params = entry.params;
      if (entry.unsure) info.paramsUnsure = true;
    }
    cache = functions;
  }
  return cache;
}

/** Looks up a function by its display name, case-insensitively. */
export function lookupFunction(name: string): FunctionInfo | undefined {
  return catalogue().get(name.toUpperCase());
}

/**
 * The built-in function a defined name collides with, if any (probe T12): a name
 * `Fact` is accepted by Excel, but `=Fact(5)` calls the built-in `FACT`.
 */
export function builtinCollision(definedName: string): FunctionInfo | undefined {
  return lookupFunction(definedName);
}

/** Parses an XLM table (xlm-data.ts): upper-case names; throws on a malformed line. */
export function parseXlm(table: string): Set<string> {
  const out = new Set<string>();
  table.split("\n").forEach((raw, k) => {
    const line = raw.trim();
    if (line === "" || line.startsWith("#")) return;
    const cols = line.split(/\s+/);
    const where = `xlm line ${k + 1} (${line})`;
    if (cols.length !== 2 || !/^0x[0-9A-F]{4}$/.test(cols[0]!)) throw new Error(`${where}: expected 0xNNNN NAME`);
    const name = cols[1]!;
    if (name !== name.toUpperCase() || !/^[A-Z][A-Z0-9.]*$/.test(name)) throw new Error(`${where}: bad name`);
    if (out.has(name)) throw new Error(`${where}: duplicate`);
    out.add(name);
  });
  return out;
}

let xlm: Set<string> | undefined;

/**
 * The Excel 4.0 macro function a defined name is spelled like, upper case, if any and if
 * it is not also a worksheet function (that is builtinCollision's). Not measured: AFE
 * issue #10 reports such LAMBDAs called as the macro function, or refused. Functions
 * only (Ftab): the command equivalents (`Open`, `Save.As`, `Copy`) are ordinary words,
 * and warning on them was noise (author, 2026-10-08).
 */
export function xlmCollision(definedName: string): string | undefined {
  xlm ??= parseXlm(XLM_FUNCTIONS);
  const up = definedName.toUpperCase();
  return xlm.has(up) && !lookupFunction(up) ? up : undefined;
}
