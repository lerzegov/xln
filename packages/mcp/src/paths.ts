// The allowed roots: an agent names files, and the server reads and writes only under the
// folders it was started with. Paths are compared after resolving symbolic links, so a
// link inside a root cannot lead out of it.

import { existsSync, realpathSync } from "node:fs";
import { basename, dirname, isAbsolute, join, relative, resolve, sep } from "node:path";

export class PathError extends Error {}

export class Roots {
  /** As given (resolved), for messages. */
  readonly given: string[];
  private readonly real: string[];

  constructor(roots: string[]) {
    if (roots.length === 0) throw new Error("at least one root is needed");
    this.given = roots.map((r) => resolve(r));
    for (const r of this.given) if (!existsSync(r)) throw new Error(`no such root folder: ${r}`);
    this.real = this.given.map((r) => realpathSync(r));
  }

  /**
   * An agent's path, absolute or relative to the first root, checked to lie inside a root.
   * It need not exist (a project folder a pull will create); its nearest existing ancestor
   * is what links are resolved on.
   */
  resolve(path: string): string {
    if (path.trim() === "") throw new PathError("empty path");
    const abs = isAbsolute(path) ? resolve(path) : resolve(this.given[0]!, path);
    this.check(abs);
    return abs;
  }

  /** Refuses a path (already absolute) outside every root. */
  check(abs: string): void {
    const real = realOf(abs);
    if (!this.real.some((r) => real === r || real.startsWith(r.endsWith(sep) ? r : r + sep))) {
      throw new PathError(`${abs} is outside the allowed roots (${this.given.join(", ")}): start xln-mcp with --root <folder> to allow it`);
    }
  }

  /** A path for messages: relative to the first root when inside it. */
  show(abs: string): string {
    const r = relative(this.given[0]!, abs);
    return r && !r.startsWith("..") && !isAbsolute(r) ? r : abs;
  }
}

/** The real path of `abs`, or of its nearest existing ancestor joined with the rest. */
function realOf(abs: string): string {
  let dir = abs;
  const rest: string[] = [];
  while (!existsSync(dir)) {
    const up = dirname(dir);
    if (up === dir) return abs;
    rest.unshift(basename(dir));
    dir = up;
  }
  return join(realpathSync(dir), ...rest);
}
