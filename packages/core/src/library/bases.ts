// The library bases' texts (author's idea, 2026-10-07): `@from(lib #…)` is only a hash, so a
// copy changed on both sides could only be diffed against the library unless some workbook
// or backup still had the base's text. Insert, Take, Publish and Record library base (the
// explicit actions that write `@from`) therefore also keep the base definition in the
// project, one file per version: `library-bases/<hash>.json`. The folder is internal, like
// the lockfile and the manifest (hidden and read-only in the editor); a pull leaves it alone
// (it writes `names/**`, the lockfile and the manifest only), and nothing prunes it: an
// entry no `@from` names any more costs a few hundred bytes.
//
// Keyed by the hash, a file is the same whoever writes it: writing it again changes nothing.

import { compileWithDiagnostics } from "../lang/transform.js";
import { stringifyJson } from "../project/json.js";
import { libraryHash } from "../project/provenance.js";
import type { LibraryFunction } from "./lambdaFile.js";

/** The project folder that keeps the library bases' texts. */
export const BASES_DIR = "library-bases";

/** One library version, as Insert, Take, Publish or Record library base saw it. */
export interface LibraryBase {
  /** `libraryHash` of `stored`: what `@from(lib #…)` records. */
  hash: string;
  /** The library function's name. */
  name: string;
  /** Its `.lambda` file in the library. */
  library: string;
  /** The definition in stored form (`_xlfn.`, `_xlpm.`). */
  stored: string;
  /** The definition as the library file writes it. */
  display: string;
}

/** The project path of a base's file: `library-bases/353921.json`. */
export function baseFilePath(hash: string): string {
  return `${BASES_DIR}/${hash}.json`;
}

/** Whether a project path is a base's file. */
export function isBasePath(path: string): boolean {
  return path.startsWith(BASES_DIR + "/") && path.endsWith(".json") && !path.slice(BASES_DIR.length + 1).includes("/");
}

/** The base a definition (display form) of `name` from library file `library` is. */
export function libraryBase(name: string, display: string, library: string): LibraryBase {
  const stored = compileWithDiagnostics(display, { names: [name], allowUnknownFunctions: true }).text;
  return { hash: libraryHash(stored), name, library, stored, display };
}

/** The base a library function's current version is (Insert, Take, Record library base). */
export function libraryFunctionBase(fn: LibraryFunction): LibraryBase {
  return libraryBase(fn.name, fn.definition, fn.path);
}

/** A base's file text. */
export function baseFileText(b: LibraryBase): string {
  return stringifyJson({ hash: b.hash, name: b.name, library: b.library, stored: b.stored, display: b.display }) + "\n";
}

/** A base's file read back; undefined when it is not one (or its text does not give its hash). */
export function parseBaseFile(text: string): LibraryBase | undefined {
  let v: unknown;
  try {
    v = JSON.parse(text);
  } catch {
    return undefined;
  }
  if (v === null || typeof v !== "object") return undefined;
  const o = v as Record<string, unknown>;
  const str = (k: string) => (typeof o[k] === "string" ? (o[k] as string) : undefined);
  const hash = str("hash");
  const stored = str("stored");
  const display = str("display");
  if (hash === undefined || stored === undefined || display === undefined) return undefined;
  // A hand-edited file must not stand in for a version it is not.
  if (libraryHash(stored) !== hash) return undefined;
  return { hash, name: str("name") ?? "", library: str("library") ?? "", stored, display };
}

/** The bases among a project's files (path → text), by hash. */
export function readBases(files: Readonly<Record<string, string>>): Map<string, LibraryBase> {
  const out = new Map<string, LibraryBase>();
  for (const [path, text] of Object.entries(files)) {
    if (!isBasePath(path)) continue;
    const b = parseBaseFile(text);
    if (b) out.set(b.hash, b);
  }
  return out;
}

/** The file writes that keep these bases (path → text), one per version. */
export function baseFiles(bases: Iterable<LibraryBase>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const b of bases) out[baseFilePath(b.hash)] = baseFileText(b);
  return out;
}
