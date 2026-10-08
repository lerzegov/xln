// D6 in a build: after the plan, give every name the build leaves in sync the comment it
// should carry in the file: its doc comment, plus the provenance tag when it comes from a
// module (project/provenance.ts). The plan compares comments without tags, so a tag is
// never an edit and never a conflict; this step only adds `set-name` changes (field
// `provenance`) for names whose tag is missing, stale or no longer due, and puts the tag
// on the comments of the `set-name` changes the plan made. Names Excel changed since the
// last pull are not in sync and keep whatever tag they have: a stale hash is what tells a
// later pull that the name was edited in Excel. A name's `@from(lib #…)` goes into its tag
// as `lib#…`, so a change of the annotation alone is a `provenance` change too.

import type { DefinedName, WorkbookSnapshot } from "../file/types.js";
import { buildProvenanceTag, commentLength, COMMENT_MAX, docTagOverflow, moduleOfPath, sourceHash, sourceHashV1, withProvenance } from "../project/provenance.js";
import { orderChanges, scopedKey, type Change, type Scope, type SetName } from "./changes.js";
import type { BuildPlan } from "./plan.js";
import { readSourceProject } from "./source.js";

function lkey(name: string, scope: Scope | undefined): string {
  return `${(scope ?? "").toLowerCase()}!${name.toLowerCase()}`;
}

function rawComment(d: DefinedName): string | null {
  return d.comment === undefined || d.comment === "" ? null : d.comment.split("\r\n").join("\n");
}

/**
 * Excel refuses to open a workbook whose name has a comment over 255 characters (measured
 * on the Mac, probes/README.md § F5: 255 opens, 256 does not). A build that would write one
 * is refused, tag or no tag.
 */
export function checkCommentLengths(plan: BuildPlan): void {
  for (const c of plan.changeSet.changes) {
    if (c.op !== "set-name" || c.comment === null || commentLength(c.comment) <= COMMENT_MAX) continue;
    // The source's own comment over the limit is already the checker's error; this one is the tag's.
    if (plan.problems.some((p) => p.code === "comment-length" && p.key?.toLowerCase() === scopedKey(c.name, c.scope).toLowerCase())) continue;
    plan.problems.push({
      severity: "error",
      code: "comment-length",
      message: `${scopedKey(c.name, c.scope)}: the comment has ${commentLength(c.comment)} characters (a line break counts 2); Excel refuses to open a workbook with a name comment over ${COMMENT_MAX}`,
      key: scopedKey(c.name, c.scope),
    });
  }
}

/** Adds the provenance tags to `plan` (in place). */
export function addProvenance(plan: BuildPlan, wb: WorkbookSnapshot, files: Readonly<Record<string, string>>): void {
  const source = readSourceProject(files);
  const sheetByLower = new Map(wb.sheets.map((s) => [s.name.toLowerCase(), s.name]));
  const fileOf = new Map<string, string>();
  // The library base (`@from(lib #…)`) goes into the tag as `lib#…`: adding, changing or
  // removing it is a provenance change, never a definition change or a conflict.
  const libOf = new Map<string, string>();
  for (const s of source.names) {
    const scope = s.scope === undefined ? null : (sheetByLower.get(s.scope.toLowerCase()) ?? s.scope);
    fileOf.set(lkey(s.name, scope), s.file);
    if (s.libBase !== undefined) libOf.set(lkey(s.name, scope), s.libBase);
  }
  const live = new Map<string, DefinedName>();
  for (const d of wb.definedNames) if (!d.isXlPrefixed && !d.scopeInvalid) live.set(lkey(d.name, d.scope.kind === "sheet" ? d.scope.name : null), d);
  const changes: Change[] = plan.changeSet.changes;
  const sets = new Map<string, SetName>();
  for (const c of changes) if (c.op === "set-name") sets.set(lkey(c.name, c.scope), c);

  for (const s of plan.inSync) {
    const key = lkey(s.name, s.scope);
    const file = fileOf.get(key);
    const module = file === undefined ? undefined : moduleOfPath(file);
    const set = sets.get(key);
    const d = live.get(key);
    const stored = set?.stored ?? d?.definition;
    if (stored === undefined) continue; // renamed without other edits: tagged by the next build
    let want = s.comment;
    // A tag written before numbers counted by their value, still right for this name: kept, not rewritten.
    let alsoRight: string | undefined;
    if (module !== undefined) {
      const lib = libOf.get(key);
      const tag = buildProvenanceTag(file!, files[file!] ?? "", lib, sourceHash(stored, s.comment))!;
      const tagged = withProvenance(s.comment, tag);
      alsoRight = withProvenance(s.comment, { ...tag, hash: sourceHashV1(stored, s.comment) });
      if (commentLength(tagged) <= COMMENT_MAX) want = tagged;
      else {
        // The checker said it already, on the doc comment (docTagOverflow); this is for a
        // plan whose problems come from elsewhere.
        const k = scopedKey(s.name, s.scope);
        const said = plan.problems.some((p) => p.code === "provenance" && p.key?.toLowerCase() === k.toLowerCase());
        const over = docTagOverflow(s.comment, tag);
        if (!said && over)
          plan.problems.push({ severity: "warning", code: "provenance", message: `${k}: ${over.message}`, ...(file !== undefined ? { file } : {}), key: k });
      }
    }
    if (set) {
      if (d && !set.fields.includes("created") && rawComment(d) !== want && !set.fields.includes("comment")) set.fields.push("provenance");
      set.comment = want;
      continue;
    }
    if (!d || rawComment(d) === want || (want !== s.comment && rawComment(d) === alsoRight)) continue;
    const add: SetName = { op: "set-name", name: s.name, scope: s.scope, stored: d.definition, display: s.display, comment: want, hidden: s.hidden, fields: ["provenance"] };
    changes.push(add);
    sets.set(key, add);
  }
  plan.changeSet.changes = orderChanges(changes);
}
