// A library: every `.lambda` file of a folder, read with `parseLambdaFile`. The caller
// lists the folder and reads the files (CLI: node:fs; extension: vscode.workspace.fs);
// the core gets path → text. Library names compare case-insensitively, as Excel's do.

import { tokenize } from "../lang/tokens.js";
import { compileWithDiagnostics } from "../lang/transform.js";
import { libraryHash } from "../project/provenance.js";
import { parseLambdaFile, type LibraryFunction, type LibraryProblem } from "./lambdaFile.js";

export interface Library {
  /** By name, sorted. */
  functions: LibraryFunction[];
  problems: LibraryProblem[];
  get(name: string): LibraryFunction | undefined;
}

/** Whether a file of the library folder is a library function. */
export function isLambdaFile(path: string): boolean {
  return path.toLowerCase().endsWith(".lambda");
}

/** Reads the `.lambda` files among `files` (path → text). A name given twice keeps its first file. */
export function readLibrary(files: Readonly<Record<string, string>>): Library {
  const paths = Object.keys(files).filter(isLambdaFile).sort();
  // First the names, so that each definition compiles knowing what else the library defines.
  const names = paths.map((p) => parseLambdaFile(p, files[p]!).fn?.name).filter((n): n is string => n !== undefined);
  const byName = new Map<string, LibraryFunction>();
  const problems: LibraryProblem[] = [];
  for (const p of paths) {
    const r = parseLambdaFile(p, files[p]!, names);
    problems.push(...r.problems);
    if (!r.fn) continue;
    const k = r.fn.name.toLowerCase();
    const first = byName.get(k);
    if (first) {
      problems.push({ severity: "error", path: p, message: `${r.fn.name} is also defined in ${first.path}; this file is ignored` });
      continue;
    }
    byName.set(k, r.fn);
  }
  const functions = [...byName.values()].sort((a, b) => (a.name.toLowerCase() < b.name.toLowerCase() ? -1 : a.name.toLowerCase() > b.name.toLowerCase() ? 1 : 0));
  return { functions, problems, get: (name) => byName.get(name.toLowerCase()) };
}

/** A library function's definition in stored form (`_xlfn.`, `_xlpm.`), as a build would write it. */
export function libraryStored(fn: LibraryFunction, names: Iterable<string>): string {
  return compileWithDiagnostics(fn.definition, { names: [...names, fn.name], allowUnknownFunctions: true }).text;
}

/**
 * A library function's version: `libraryHash` of its stored form, what `@from(lib #…)`
 * records and `lib status` shows. Compiled on its own: the stored form of a library
 * definition does not depend on which other names exist (calls to names stay bare).
 */
export function libraryFunctionHash(fn: LibraryFunction): string {
  return definitionBase(fn.definition, fn.name);
}

/** `libraryHash` of a LAMBDA named `name` given in display form (a library file's, or a project's about to be published). */
export function definitionBase(display: string, name: string): string {
  return libraryHash(compileWithDiagnostics(display, { names: [name], allowUnknownFunctions: true }).text);
}

/**
 * The library functions `name` calls, directly or through others, that `has` says are
 * missing, in dependency order (callees first) and `name` last. Inserting a function
 * from the library brings these along, or the project would call names it lacks.
 */
export function libraryClosure(library: Library, name: string, has: (name: string) => boolean): LibraryFunction[] {
  const out: LibraryFunction[] = [];
  const seen = new Set<string>();
  const visit = (fn: LibraryFunction): void => {
    const k = fn.name.toLowerCase();
    if (seen.has(k)) return;
    seen.add(k);
    for (const dep of calledNames(fn.definition)) {
      const d = library.get(dep);
      if (d && !has(d.name)) visit(d);
    }
    out.push(fn);
  };
  const fn = library.get(name);
  if (fn) visit(fn);
  return out;
}

/** Unqualified defined names a formula mentions (calls and plain references). */
export function calledNames(display: string): string[] {
  const out: string[] = [];
  for (const t of tokenize(display)) if (t.kind === "name" && !t.qual && t.value !== undefined) out.push(t.value);
  return out;
}
