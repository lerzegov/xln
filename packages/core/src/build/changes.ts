// D9: a build is a list of changes, computed from the source project, the lockfile and the
// workbook, and applied by a backend. The file backend (apply.ts) patches the workbook's
// bytes; a live backend (AppleScript on the Mac, COM on Windows) can apply the same list
// to an open workbook later. So the list is plain JSON, names sheets by name (a backend
// finds the position), and carries each definition in both forms: `stored` for the file,
// `display` for Excel's own entry points (Name Manager, `RefersTo`).
//
// M3b adds cell-formula changes (`set-cell-formula`, `clear-cell-formula`): Excel owns the
// layout, so a cell change never adds or moves cells, it only replaces the formula of cells
// the workbook already has (or fills an empty named cell, a slot).

export const CHANGESET_FORMAT = "xln.changes/1";

/** A sheet name, or null for workbook scope. */
export type Scope = string | null;

/** What a `set-name` changes on an existing name, or `created`. */
export type NameField = "created" | "definition" | "comment" | "hidden" | "spelling" | "provenance";

/** Create a name or bring an existing one to this state. */
export interface SetName {
  op: "set-name";
  name: string;
  scope: Scope;
  /** The definition as stored in the file: prefixes, `_xlpm.`, qualified home sheet; CR LF line breaks. */
  stored: string;
  /** The definition as Excel displays it (and as the source has it), without the leading `=`. */
  display: string;
  /** The Name Manager comment as written to the file, with its provenance tag (D6); null for none. */
  comment: string | null;
  hidden: boolean;
  /** What differs from the workbook; `created` when the name is new. */
  fields: NameField[];
  /** The stored form changes although the source text did not: compile adds a prefix
   *  Excel needs (probe F6), so the build repairs the name. */
  repair?: true;
}

export interface DeleteName {
  op: "delete-name";
  name: string;
  scope: Scope;
}

/** Rename within names (D4). Applied before the `set-name` of the new name. */
export interface RenameName {
  op: "rename-name";
  scope: Scope;
  from: string;
  to: string;
  /**
   * Stretch G (M5): also rewrite the name's token wherever the workbook's formulas read it
   * (cell formulas, a shared group's master, conditional formats, validations, other
   * names' definitions), as Excel's Name Manager does; read on the names before every
   * rename of the set. Nothing else of those formulas changes. The counts are what the
   * plan found, for reports.
   */
  references?: RenameReferences;
}

/** How many formulas of each kind read the renamed name (stretch G). */
export interface RenameReferences {
  cells: number;
  formats: number;
  validations: number;
  names: number;
}

/** Move a name to another scope. Applied before the `set-name` of the name in its new scope. */
export interface RescopeName {
  op: "rescope-name";
  name: string;
  from: Scope;
  to: Scope;
}

/**
 * Give existing cells a formula (M3b, D7/D8). `range` is one cell (`C6`) or a block that
 * holds one formula filled across (`B40:G40`): `stored` and `display` are the top-left
 * cell's text, and a backend shifts relative references for the other cells. A spill
 * anchor is a single cell; Excel decides its extent. A backend writes every formula in
 * dynamic-array form (probe F8).
 */
export interface SetCellFormula {
  op: "set-cell-formula";
  sheet: string;
  range: string;
  /** The top-left cell's formula as stored in the file, without the leading `=`. */
  stored: string;
  /** As Excel displays it (and as the source has it), without the leading `=`. */
  display: string;
  /** The name on this cell or block, if any (identity for E6 and reports). */
  name?: string;
  /** The top-left cell's stored formula before the build; absent for a slot (empty cell). */
  previous?: string;
}

/** Remove the formulas of existing cells, leaving them empty (the source wrote `= ;`). */
export interface ClearCellFormula {
  op: "clear-cell-formula";
  sheet: string;
  range: string;
  name?: string;
  previous?: string;
}

/**
 * D5: carry the project's source in the workbook (a custom XML part, `embedXml.ts`), so a
 * workbook passed on still has its formatted modules, doc comments and lockfile. Replaces
 * the part a previous build embedded, or adds one. `files` maps project paths to their
 * text: every `names/**\/*.xln`, the lockfile as the build leaves it, and the audit config.
 * Not a change to the workbook's content: a live backend that cannot write custom XML
 * may skip it (the source is then embedded by the next file build).
 */
export interface SetEmbeddedSource {
  op: "set-embedded-source";
  files: Record<string, string>;
}

export type Change = RenameName | RescopeName | DeleteName | SetName | SetCellFormula | ClearCellFormula | SetEmbeddedSource;

export interface ChangeSet {
  format: typeof CHANGESET_FORMAT;
  /** The workbook's file name. */
  workbook: string;
  /** In the order to apply: renames, scope changes, deletions, sets, then cell formulas. */
  changes: Change[];
}

const ORDER: Record<Change["op"], number> = {
  "rename-name": 0,
  "rescope-name": 1,
  "delete-name": 2,
  "set-name": 3,
  "clear-cell-formula": 4,
  "set-cell-formula": 5,
  "set-embedded-source": 6,
};

/** Sorts changes into the order a backend applies them (stable within one kind). */
export function orderChanges(changes: readonly Change[]): Change[] {
  return changes
    .map((c, i) => ({ c, i }))
    .sort((a, b) => ORDER[a.c.op] - ORDER[b.c.op] || a.i - b.i)
    .map((x) => x.c);
}

/** `Sheet!Name` or `Name`, as in the lockfile and the manifest. */
export function scopedKey(name: string, scope: Scope | undefined): string {
  return scope === null || scope === undefined ? name : `${scope}!${name}`;
}

/** `12 cell formulas, 1 conditional format, 3 names`. */
export function referenceCounts(r: RenameReferences): string {
  const n = (k: number, one: string, many = `${one}s`) => (k === 0 ? "" : `${k} ${k === 1 ? one : many}`);
  const parts = [n(r.cells, "cell formula"), n(r.formats, "conditional format"), n(r.validations, "validation"), n(r.names, "name")].filter(Boolean);
  return parts.length ? parts.join(", ") : "no formula";
}

/** A one-line description of a change, for reports. */
export function describeChange(c: Change): string {
  switch (c.op) {
    case "rename-name":
      return `rename ${scopedKey(c.from, c.scope)} → ${c.to}${c.references ? `, rewriting it in ${referenceCounts(c.references)}` : ""}`;
    case "rescope-name":
      return `move ${scopedKey(c.name, c.from)} to ${c.to === null ? "workbook scope" : `sheet ${c.to}`}`;
    case "delete-name":
      return `delete ${scopedKey(c.name, c.scope)}`;
    case "set-name":
      return `${c.fields.includes("created") ? "create" : "update"} ${scopedKey(c.name, c.scope)}${c.fields.includes("created") ? "" : ` (${c.fields.join(", ")})`}${c.repair ? " [repairs a missing prefix]" : ""}`;
    case "set-cell-formula":
      return `${c.previous === undefined ? "fill" : "set formula of"} ${c.sheet}!${c.range}${c.name ? ` (${c.name})` : ""}`;
    case "clear-cell-formula":
      return `clear formula of ${c.sheet}!${c.range}${c.name ? ` (${c.name})` : ""}`;
    case "set-embedded-source": {
      const n = Object.keys(c.files).filter((p) => p.endsWith(".xln")).length;
      const others = Object.keys(c.files).filter((p) => !p.endsWith(".xln"));
      return `embed the source (${n} names file${n === 1 ? "" : "s"}${others.length ? `, ${others.join(", ")}` : ""})`;
    }
  }
}

/** A `set-name` whose only change is the provenance tag in its comment (D6). */
export function provenanceOnly(c: Change): boolean {
  return c.op === "set-name" && c.fields.length === 1 && c.fields[0] === "provenance";
}

/**
 * The changes described, one line each, for text reports, except updates of the
 * provenance tag alone: those are one line after the others (feedback 2026-10-08: the first
 * build of a workbook xln never built tags every module name, ~20 lines before the real
 * change). The change set itself (JSON) keeps every change.
 */
export function describeChanges(changes: readonly Change[]): string[] {
  const tags = changes.filter(provenanceOnly);
  if (tags.length < 2) return changes.map(describeChange);
  const out = changes.filter((c) => !provenanceOnly(c)).map(describeChange);
  out.push(`update the provenance tag of ${tags.length} module names (comment only)`);
  return out;
}
