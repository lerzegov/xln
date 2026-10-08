// Where an audit finding belongs in the editor, without the vscode API: a name finding on
// its entry in the project's `.xln` files (on the part of the formula it is about, when it
// can be found there), a cell finding on the cell's line of the formula view. What the
// audit report document links to.

import { formulaToSource, stripPrefix, type Finding, type FindingWhere } from "@xln/core";
import type { Loc, NameDef, Project } from "./project.js";

/** Severity names as the Problems panel has them. */
export type ProblemSeverity = "error" | "warning" | "information";

export function problemSeverity(f: Finding): ProblemSeverity {
  return f.severity === "info" ? "information" : f.severity;
}

/** The text of a diagnostic: the message, then the hint. */
export function problemMessage(f: Finding): string {
  return `${f.check} ${f.message}${f.hint ? ` (${f.hint})` : ""}`;
}

/** The definition a name finding is about, in the project's sources. */
export function findingDef(project: Project, where: FindingWhere): NameDef | undefined {
  return where.kind === "name" && where.key !== undefined ? project.lookup(where.key) : undefined;
}

/**
 * Where a name finding sits in the project: the part of the formula it is about (the text
 * the audit saw in the stored definition, found again in the source's display form), or
 * else the name itself.
 */
export function findingLoc(project: Project, f: Finding): Loc | undefined {
  const def = findingDef(project, f.where);
  if (!def) return undefined;
  const text = f.where.text;
  if (text) {
    const formula = def.entry.formula.toLowerCase();
    // The stored form may carry prefixes the display form does not (`_xlfn.`), or the own sheet.
    const tries = [text, stripPrefix(text).base, text.slice(text.lastIndexOf("!") + 1)].filter((t) => t.length > 0);
    for (const t of tries) {
      const i = formula.indexOf(t.toLowerCase());
      if (i >= 0) {
        const start = formulaToSource(def.entry, i);
        const end = formulaToSource(def.entry, i + t.length - 1) + 1;
        return { path: def.file.path, start, end };
      }
    }
  }
  return project.nameLoc(def);
}
