// A7: the lockfile records, per name, a hash of what was pulled (later: built). A later
// pull or build compares against it to tell "changed in Excel" from "changed in source"
// (E2). Definitions are hashed in a normalised form, so Excel re-spacing a formula
// (probe T14) or quoting a sheet name differently does not count as a change.
//
// Format 2 (M3b) adds `cells`: one entry per cell statement (E6), keyed by the name for a
// named cell or a slot (`BS!Revenue`) and by `Sheet!Range` for an unnamed one
// (`BS!B40:G40`), with the hash of the top-left cell's stored formula (null for a slot).
// Format 1 files (no `cells`) are still read; a build then checks cells two-way.
//
// Format 3 changes the hashes, not the fields:
// - A number literal counts by its value: Excel rewrites `1E-14` as `0.00000000000001`
//   when it saves, which formats 1 and 2 took for an edit made in Excel.
// - `h:` and the first 16 hex digits (64 bits) of SHA-256, in place of `sha256:` and all
//   64: the lockfile is embedded in the workbook (D5) and full hashes do not compress.
//   64 bits are plenty to tell one state of a name from the next.
// - One line per entry, so the file stays readable and a diff shows one line per change.
// Every hash says by its prefix how it was computed, and each comparison with the
// workbook uses the function of the hash it compares with. So a lockfile of format 1 or
// 2 reads without drift, and a build that keeps an old entry (a name changed in Excel)
// carries its `sha256:` hash into the format-3 (or 4) file it writes: the hash of a state
// the build no longer has cannot be recomputed in the new form.
//
// Format 4 (2026-10-05) changes neither hashes nor fields: it says that the sheet files
// write `#` when a named cell's name is on the spill (`Name @C6# = …;`), so a missing `#`
// means the cell alone (`explicitSpill`). Formats 1–3 read as before.

import { tokenKeys } from "../lang/format.js";
import { sha256 } from "./hash.js";
import type { Json } from "./json.js";
import { nameKey, type ProjectName } from "./types.js";

export const LOCK_FORMAT = "xln.lock/4";
/** Formats `parseLockfile` reads. */
export const LOCK_FORMATS_READ: readonly string[] = ["xln.lock/1", "xln.lock/2", "xln.lock/3", LOCK_FORMAT];
/** The last format whose project may predate the explicit `#` of named cell statements. */
export const LOCK_FORMAT_IMPLICIT_SPILL = "xln.lock/3";

/**
 * Whether the sheet files of this lockfile's project write a name on a spill as `@C6#`.
 * Format 4 says so; before it, pull wrote `Name @C6` whatever the name covered, and a build
 * read a missing `#` on a name that is on its spill (in the lockfile and in the workbook)
 * as the spill. A build keeps format 3 while the source still relies on that reading.
 */
export function explicitSpill(lock: Pick<Lockfile, "format">): boolean {
  return lock.format === LOCK_FORMAT;
}

const PREFIX = "h:";
const PREFIX_V2 = "sha256:";
const DIGITS = 16;

/** The normalised text a definition hash is computed from: its tokens without layout, numbers by value. */
export function normalizeDefinition(stored: string): string {
  return tokenKeys(stored).join("\n");
}

/** The normalised text of formats 1 and 2: numbers as written. */
export function normalizeDefinitionV2(stored: string): string {
  return tokenKeys(stored, { numbers: "text" }).join("\n");
}

function lfComment(comment: string): string {
  return comment.split("\r\n").join("\n");
}

export function definitionHash(stored: string): string {
  return PREFIX + sha256(normalizeDefinition(stored)).slice(0, DIGITS);
}

export function commentHash(comment: string | undefined): string | null {
  return comment === undefined || comment === "" ? null : PREFIX + sha256(lfComment(comment)).slice(0, DIGITS);
}

/** The definition hash of lockfile formats 1 and 2. */
export function definitionHashV2(stored: string): string {
  return PREFIX_V2 + sha256(normalizeDefinitionV2(stored));
}

/** The comment hash of lockfile formats 1 and 2. */
export function commentHashV2(comment: string | undefined): string | null {
  return comment === undefined || comment === "" ? null : PREFIX_V2 + sha256(lfComment(comment));
}

/** Whether a hash read from a lockfile was computed by formats 1 and 2. */
export function isV2Hash(hash: string | null): boolean {
  return hash !== null && hash.startsWith(PREFIX_V2);
}

/** `stored` hashed the way `like` was (a lockfile hash; null takes the current way). */
export function definitionHashLike(like: string | null, stored: string): string {
  return isV2Hash(like) ? definitionHashV2(stored) : definitionHash(stored);
}

/** `comment` hashed the way `like` was. */
export function commentHashLike(like: string | null, comment: string | undefined): string | null {
  return isV2Hash(like) ? commentHashV2(comment) : commentHash(comment);
}

export interface LockEntry {
  definition: string;
  comment: string | null;
  hidden: boolean;
}

/** A cell statement as the last pull or build saw it. */
export interface LockCell {
  sheet: string;
  /** `C6` or `B40:G40`: where the statement was. */
  range: string;
  /** The name's key, for a named cell or a slot. */
  name?: string;
  /** `definitionHash` of the top-left stored formula; null for an empty cell (a slot). */
  formula: string | null;
}

export interface Lockfile {
  format: string;
  workbook: string;
  names: Record<string, LockEntry>;
  /** Absent in format 1. */
  cells?: Record<string, LockCell>;
}

/** What `buildLockfile` needs of a cell statement (`CellStatement` has it). */
export interface LockableStatement {
  sheet: string;
  range: string;
  key?: string;
  stored?: string;
}

/** How a lockfile's hashes are computed: `v2` for formats 1 and 2. */
export type LockHashes = "v3" | "v2";

/** The lockfile key of a cell statement: the name's key, or `Sheet!Range` for an unnamed one. */
export function cellKey(s: { key?: string | undefined; sheet: string; range: string }): string {
  return s.key ?? `${s.sheet}!${s.range}`;
}

/** The lock entry of a cell statement. */
export function lockCell(s: LockableStatement, hashes: LockHashes = "v3"): LockCell {
  const h = hashes === "v2" ? definitionHashV2 : definitionHash;
  const c: LockCell = { sheet: s.sheet, range: s.range, formula: s.stored === undefined ? null : h(s.stored) };
  if (s.key !== undefined) c.name = s.key;
  return c;
}

/**
 * `names` must already be in project order; `cells` in sheet order. `hashes: "v2"` gives
 * the hashes of formats 1 and 2, to compare with a lockfile of those formats (the file
 * itself is always written in format 3).
 */
export function buildLockfile(fileName: string, names: readonly ProjectName[], cells: readonly LockableStatement[] = [], hashes: LockHashes = "v3"): Lockfile {
  const dh = hashes === "v2" ? definitionHashV2 : definitionHash;
  const ch = hashes === "v2" ? commentHashV2 : commentHash;
  const out: Record<string, LockEntry> = {};
  for (const n of names) {
    out[nameKey(n)] = { definition: dh(n.stored), comment: ch(n.comment), hidden: n.hidden };
  }
  const cellOut: Record<string, LockCell> = {};
  for (const s of cells) cellOut[cellKey(s)] = lockCell(s, hashes);
  return { format: hashes === "v2" ? "xln.lock/2" : LOCK_FORMAT, workbook: fileName, names: out, cells: cellOut };
}

export function lockfileJson(lock: Lockfile): Json {
  return lock as unknown as Json;
}

function entryLine(key: string, fields: Record<string, Json | undefined>): string {
  const parts = Object.entries(fields)
    .filter(([, v]) => v !== undefined)
    .map(([k, v]) => `${JSON.stringify(k)}: ${JSON.stringify(v)}`);
  return `    ${JSON.stringify(key)}: { ${parts.join(", ")} }`;
}

function block(name: string, lines: string[]): string {
  return lines.length === 0 ? `  ${JSON.stringify(name)}: {}` : `  ${JSON.stringify(name)}: {\n${lines.join(",\n")}\n  }`;
}

/** The text of `xln.lock.json`: JSON with one line per name and per cell, ending in a line break. */
export function lockfileText(lock: Lockfile): string {
  const parts = [`  "format": ${JSON.stringify(lock.format)}`, `  "workbook": ${JSON.stringify(lock.workbook)}`];
  parts.push(block("names", Object.entries(lock.names).map(([k, e]) => entryLine(k, { definition: e.definition, comment: e.comment, hidden: e.hidden }))));
  if (lock.cells !== undefined) {
    parts.push(block("cells", Object.entries(lock.cells).map(([k, c]) => entryLine(k, { sheet: c.sheet, range: c.range, name: c.name, formula: c.formula }))));
  }
  return `{\n${parts.join(",\n")}\n}\n`;
}

/** Reads `xln.lock.json`. Throws on anything that is not a lockfile of a format it reads. */
export function parseLockfile(text: string): Lockfile {
  let v: unknown;
  try {
    v = JSON.parse(text);
  } catch (e) {
    throw new Error(`lockfile is not JSON: ${(e as Error).message}`);
  }
  const o = v as Partial<Lockfile> | null;
  if (!o || typeof o !== "object" || !LOCK_FORMATS_READ.includes(o.format as string)) throw new Error(`lockfile format is not ${LOCK_FORMATS_READ.join(" or ")}`);
  if (typeof o.workbook !== "string" || !o.names || typeof o.names !== "object") throw new Error("lockfile lacks workbook or names");
  const names: Record<string, LockEntry> = {};
  for (const [k, e] of Object.entries(o.names)) {
    if (!e || typeof e.definition !== "string" || (e.comment !== null && typeof e.comment !== "string") || typeof e.hidden !== "boolean") {
      throw new Error(`lockfile entry ${k} is malformed`);
    }
    names[k] = { definition: e.definition, comment: e.comment, hidden: e.hidden };
  }
  const out: Lockfile = { format: o.format!, workbook: o.workbook, names };
  if (o.cells !== undefined) {
    if (!o.cells || typeof o.cells !== "object") throw new Error("lockfile cells are malformed");
    const cells: Record<string, LockCell> = {};
    for (const [k, e] of Object.entries(o.cells)) {
      if (!e || typeof e.sheet !== "string" || typeof e.range !== "string" || (e.formula !== null && typeof e.formula !== "string") || (e.name !== undefined && typeof e.name !== "string")) {
        throw new Error(`lockfile cell entry ${k} is malformed`);
      }
      cells[k] = e.name !== undefined ? { sheet: e.sheet, range: e.range, name: e.name, formula: e.formula } : { sheet: e.sheet, range: e.range, formula: e.formula };
    }
    out.cells = cells;
  }
  return out;
}

/** Splits a key `Sheet!Name` (or `Name`) into name and scope. Names cannot hold `!`. */
export function splitNameKey(key: string): { name: string; scope: string | undefined } {
  const i = key.lastIndexOf("!");
  return i < 0 ? { name: key, scope: undefined } : { name: key.slice(i + 1), scope: key.slice(0, i) };
}
