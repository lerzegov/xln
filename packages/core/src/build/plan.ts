// Planning a build (D3, D4, D9, E2): the source project, the lockfile (what the last pull
// or build saw in the workbook) and the workbook as it is now give, per name, a three-way
// comparison:
//
//   source vs lock   changed in source (the author's edit)
//   workbook vs lock changed in Excel (an edit in the Name Manager since the last pull)
//
// Changed only in source → a change; only in Excel → kept, reported (the source is stale;
// pull again); in both, differently → a conflict, reported and never merged. A name is
// compared on its definition (modulo layout and sheet quotes, as the lockfile hashes it),
// its comment, its hidden flag and its spelling.

import { collectSites } from "../audit/sites.js";
import type { DefinedName, WorkbookSnapshot } from "../file/types.js";
import { walk, type Expr } from "../lang/ast.js";
import { checkProject, invalidName, renameBuilt, spillMap, strayFile, type CheckContext } from "../check/check.js";
import { SourceModel } from "../check/model.js";
import { equalModuloWhitespace } from "../lang/format.js";
import { tryParse } from "../lang/parser.js";

import { compileWithDiagnostics, decompile } from "../lang/transform.js";
import { commentHashLike, definitionHash, definitionHashLike, definitionHashV2, isV2Hash, splitNameKey, type LockCell, type LockEntry, type Lockfile } from "../project/lockfile.js";
import { stripProvenance } from "../project/provenance.js";
import { NameResolver, nameUses, type NameUse } from "../project/refs.js";
import { RenameContext, renameInFormula, type NameRename } from "../project/rename.js";
import { renameSites, type RenameSites } from "./renameSites.js";
import { parseCellRange, type ValueCells } from "../project/statements.js";
import { planCells, type CellState, type NamedCell } from "./cells.js";
import { CHANGESET_FORMAT, orderChanges, scopedKey, type Change, type ChangeSet, type NameField, type RenameName, type Scope } from "./changes.js";
import { isSourcePath, readSourceProject, type SourceCell, type SourceName, type SourceProblem } from "./source.js";

type Rect = { r1: number; c1: number; r2: number; c2: number };

/** One side's view of a name, for reports and for the editor's diff. */
export interface NameState {
  name: string;
  scope: Scope;
  /** Display form. */
  display: string;
  comment: string | null;
  hidden: boolean;
  /** For a cell statement (M3b): the cell it is written to; `display` is then its formula ("" for an empty cell), and `name` is "" for an unnamed cell. */
  cell?: { sheet: string; range: string };
}

export type ConflictKind =
  /** Changed in source and in Excel, differently. */
  | "both-changed"
  /** Deleted (or renamed, or moved) in source, changed in Excel. */
  | "deleted-in-source"
  /** Changed in source, deleted in Excel. */
  | "deleted-in-excel"
  /** New in source and new in Excel, differently. */
  | "created-in-both";

export interface Conflict {
  kind: ConflictKind;
  key: string;
  message: string;
  /** The name as the workbook has it now; absent when Excel deleted it. */
  excel?: NameState;
  /** The name as the source has it; absent when the source deleted it. */
  source?: NameState & { file: string; line: number };
}

/** A name Excel changed since the last pull while the source did not: the build keeps Excel's. */
export interface ExcelChange {
  /** `moved`: a named cell Excel moved (rows or columns inserted); its statement follows the name. */
  kind: "changed" | "created" | "deleted" | "moved";
  key: string;
  message: string;
}

export interface BuildPlan {
  changeSet: ChangeSet;
  conflicts: Conflict[];
  /** Errors stop the build; warnings do not. */
  problems: SourceProblem[];
  excelChanges: ExcelChange[];
  /** Source names that already match the workbook. */
  unchanged: number;
  /** Every source name the built workbook must match (read-back, E3): all but those Excel changed or that conflict. */
  inSync: NameState[];
  /** Cell statements that already match the workbook. */
  unchangedCells: number;
  /** Every cell statement the built workbook must match: all but those Excel changed or that conflict. */
  cellsInSync: CellState[];
  /** Lockfile key → format-3 hash, for the format-1/2 hashes the source still matches (`upgradeLockfile`). */
  lockUpgrades: { names: Map<string, string>; cells: Map<string, string> };
  /**
   * What the source changed since the lockfile, whatever Excel did: keys in lower case
   * (`sheet!name`, `name`, `sheet!range` for an unnamed cell; a renamed name under its
   * old and new key). Names (definition, comment, hidden, spelling, created, deleted,
   * renamed, moved) and cell statements (formula, created, removed) apart, since a named
   * cell has both. A pull refuses while there are any (build/unbuilt.ts). `spill`: the
   * named cell statements whose `#` the source added or removed (name keys).
   */
  sourceEdits: { names: Set<string>; cells: Set<string>; spill: Set<string> };
  /**
   * Named cell statements without `#` that this plan read as on the spill, since the
   * lockfile predates the explicit `#` (format 3 or older) and the name is on the spill in
   * it and in the workbook. While there are any, the build keeps the lockfile's format.
   */
  implicitSpill: string[];
  /**
   * Source edits the renames' reference rewrite carries (stretch G): names and cell
   * statements (keys in lower case) whose only edit is the renamed token. They are no
   * change of their own; the pull guard lists them under the rename.
   */
  renamedOnly: Set<string>;
}

export interface PlanInput {
  workbook: WorkbookSnapshot;
  /** The workbook's file name (for the change set). */
  fileName: string;
  /** Project files, path → text; `names/**\/*.xln` are read. */
  files: Readonly<Record<string, string>>;
  lock: Lockfile;
  /** Cells holding a value, per sheet (`readCellValues`): a slot that Excel filled with a value is a change in Excel. */
  values?: ValueCells;
  /** The workbook's bytes: a rename then also looks for readers only the package has (hyperlinks, form controls, pivot sources). */
  bytes?: Uint8Array;
}

/** lower-case `sheet!name` (sheet "" for workbook scope): names compare ignoring case. */
function lkey(name: string, scope: Scope | undefined): string {
  return `${(scope ?? "").toLowerCase()}!${name.toLowerCase()}`;
}

function scopeOf(d: DefinedName): Scope {
  return d.scope.kind === "sheet" ? d.scope.name : null;
}

/** A comment as compared: LF line breaks, the provenance tag stripped (D6: a tag is not an edit). */
function normComment(c: string | undefined | null): string | null {
  return c === undefined || c === null || c === "" ? null : (stripProvenance(c.split("\r\n").join("\n")) ?? null);
}

interface Src {
  s: SourceName;
  /** Scope with the workbook's spelling of the sheet. */
  scope: Scope;
  key: string;
  stored: string | undefined;
  /** Lock identity this name continues (itself, or where it was renamed or moved from). */
  base: string | undefined;
  move?: "rename" | "rescope";
  /** Problems reported only if the build writes this name. */
  deferred: SourceProblem[];
  /** A named cell statement: `s.formula` is the cell's formula, and the name's definition is
   *  its cell, set in Excel, or its spill when the address has `#` (`stored`). */
  cell?: true;
}

export function planBuild(input: PlanInput): BuildPlan {
  const { workbook: wb, lock } = input;
  const source = readSourceProject(input.files);
  // M3d: the source's errors are the checker's, the same the editor shows; the plan adds
  // only what needs the workbook (conflicts, references a change would break).
  // Its warnings too (one checker for every severity, §14 issue 8): the build lists them.
  const checked = sourceProblems(input.files, { sheets: wb.sheets.map((s) => s.name), tables: wb.tables.map((t) => t.displayName), lock, spills: spillMap(wb.sheets), links: wb.externalLinks ?? [] });
  const problems: SourceProblem[] = [];
  const conflicts: Conflict[] = [];
  const excelChanges: ExcelChange[] = [];
  const changes: Change[] = [];
  const inSync: NameState[] = [];
  let unchanged = 0;

  const sheetByLower = new Map(wb.sheets.map((s) => [s.name.toLowerCase(), s.name]));
  const cur = new Map<string, DefinedName>();
  for (const d of wb.definedNames) if (!d.isXlPrefixed && !d.scopeInvalid) cur.set(lkey(d.name, scopeOf(d)), d);
  const locked = new Map<string, { key: string; name: string; scope: Scope; entry: LockEntry }>();
  for (const [key, entry] of Object.entries(lock.names)) {
    const { name, scope } = splitNameKey(key);
    locked.set(lkey(name, scope), { key, name, scope: scope ?? null, entry });
  }

  // Source names, with the sheet spelled as in the workbook.
  // A scope changes only where the source changes it (M3d: no step re-scopes on its own).
  const srcs: Src[] = [];
  for (const s of source.names) {
    let scope: Scope = null;
    if (s.scope !== undefined) {
      const sheet = sheetByLower.get(s.scope.toLowerCase());
      if (sheet === undefined) {
        problems.push({ severity: "error", code: "unknown-sheet", message: `@scope(${s.scope}): the workbook has no such sheet`, file: s.file, line: s.line, key: scopedKey(s.name, s.scope) });
        continue;
      }
      scope = sheet;
    }
    const x: Src = { s, scope, key: lkey(s.name, scope), stored: undefined, base: undefined, deferred: [] };
    if (s.cell) x.cell = true;
    srcs.push(x);
  }
  const srcByKey = new Map(srcs.map((x) => [x.key, x]));

  // Identities: a name continues itself, or the name it was renamed from (`@renamed`), or
  // the same name in another scope when that one is gone from the source (a scope change).
  const claimed = new Set<string>();
  for (const x of srcs) {
    const r = x.s.renamedFrom;
    if (!r) continue;
    const oldScope: Scope = r.scope === undefined ? x.scope : r.scope === null ? null : (sheetByLower.get(r.scope.toLowerCase()) ?? r.scope);
    const old = lkey(r.name, oldScope);
    const where = { file: x.s.file, line: x.s.line, key: scopedKey(x.s.name, x.scope) };
    if (old === x.key) {
      // Same name ignoring case: a change of spelling, nothing to rename.
      continue;
    }
    // Already applied by an earlier build that left the annotation (a browser build's copy,
    // a build before 2026-10-07, or one put back by hand): nothing to do; the next pull drops it.
    if (renameBuilt(x.s.name, x.scope ?? undefined, r, (n, sc) => locked.has(lkey(n, sc)))) {
      x.base = x.key;
      claimed.add(x.key);
      continue;
    }
    if (!locked.has(old)) problems.push({ severity: "error", code: "rename", message: `@renamed(${scopedKey(r.name, oldScope)}): no such name in the last pull`, ...where });
    else if (srcByKey.has(old)) problems.push({ severity: "error", code: "rename", message: `@renamed(${scopedKey(r.name, oldScope)}): that name is still defined in the source`, ...where });
    else if (claimed.has(old)) problems.push({ severity: "error", code: "rename", message: `@renamed(${scopedKey(r.name, oldScope)}): another name is renamed from it too`, ...where });
    else {
      claimed.add(old);
      x.base = old;
      x.move = locked.get(old)!.name.toLowerCase() === x.s.name.toLowerCase() ? "rescope" : "rename";
    }
  }
  for (const x of srcs) {
    if (x.base !== undefined || (x.s.renamedFrom && lkey(x.s.renamedFrom.name, x.scope) !== x.key)) continue;
    if (locked.has(x.key)) {
      x.base = x.key;
      claimed.add(x.key);
    }
  }
  for (const x of srcs) {
    if (x.base !== undefined || cur.has(x.key)) continue;
    const n = x.s.name.toLowerCase();
    const candidates = [...locked.keys()].filter((k) => !claimed.has(k) && !srcByKey.has(k) && k.slice(k.indexOf("!") + 1) === n);
    if (candidates.length === 1) {
      x.base = candidates[0]!;
      x.move = "rescope";
      claimed.add(x.base);
    }
  }

  // Stretch G (M5): a rename in the same scope rewrites the name's token in the workbook's
  // formulas (cells, conditional formats, validations, names), read on the names as they
  // are now. A cell statement or a name whose only source edit is that rewrite is left to
  // it: no new formula, no dynamic-array conversion, the workbook's text kept.
  const refRenames: NameRename[] = [];
  for (const x of srcs) {
    if (x.move !== "rename" || x.base === undefined) continue;
    const d = cur.get(x.base);
    if (d && lkey("", scopeOf(d)) === lkey("", x.scope)) refRenames.push({ scope: scopeOf(d) ?? undefined, from: d.name, to: x.s.name });
  }
  const renameCtx = new RenameContext(
    [...cur.values()].map((d) => ({ name: d.name, scope: scopeOf(d) ?? undefined })),
    refRenames,
  );
  const sites = renameSites(wb, renameCtx, input.bytes, input.fileName);
  /** The workbook's formula with the renamed tokens rewritten; undefined when the renames do not touch it. */
  const renamedText = (stored: string, home: string | undefined): string | undefined => {
    if (renameCtx.empty) return undefined;
    const r = renameInFormula(stored, home, renameCtx);
    return r.count > 0 ? r.text : undefined;
  };

  // Compile every source name against the names the workbook will have.
  const allNames = srcs.map((x) => x.s.name);
  for (const d of cur.values()) if (!srcByKey.has(lkey(d.name, scopeOf(d)))) allNames.push(d.name);
  const localNames = new Map<string, string[]>();
  for (const x of srcs) {
    if (x.scope === null) continue;
    const l = localNames.get(x.scope.toLowerCase()) ?? [];
    l.push(x.s.name);
    localNames.set(x.scope.toLowerCase(), l);
  }
  const links = wb.externalLinks ?? [];
  const ctxOf = (scope: Scope) => (scope === null ? { names: allNames, links } : { homeSheet: scope, localNames: localNames.get(scope.toLowerCase()) ?? [], names: allNames, links });
  // A name's diagnostics count only when the build writes it: an unchanged name stays as
  // Excel has it, whatever its faults (the audit reports those).
  // Cells first (E6): what they read of a named cell's `#` feeds the names.
  const srcByLine = new Map(srcs.map((x) => [`${x.s.file}:${x.s.line}`, x]));
  const cellPlan = planCells({
    wb,
    lock,
    values: input.values,
    cells: source.cells,
    names: allNames,
    renamed: renamedText,
    named: (c): NamedCell | undefined => {
      const x = srcByLine.get(`${c.file}:${c.line}`);
      if (!x || x.s.name !== c.name) return undefined;
      const base = x.base ?? x.key;
      return { key: scopedKey(x.s.name, x.scope), name: x.s.name, scope: x.scope, current: cur.get(base), lockKey: x.base !== undefined ? locked.get(x.base)?.key : undefined };
    },
  });
  problems.push(...cellPlan.problems);
  conflicts.push(...cellPlan.conflicts);
  excelChanges.push(...cellPlan.excelChanges);
  const renamedOnly = new Set(cellPlan.renamedOnly);

  for (const x of srcs) {
    const where = { file: x.s.file, line: x.s.line, key: scopedKey(x.s.name, x.scope) };
    const defer = (severity: "error" | "warning", code: string, message: string) => x.deferred.push({ severity, code, message: `${where.key}: ${message}`, ...where });
    if (x.cell) {
      // The definition is the cell, which Excel owns, or its spill: the `#` is the source's.
      const d = cur.get(x.base ?? x.key);
      x.stored = cellPlan.spill.get(scopedKey(x.s.name, x.scope).toLowerCase())?.stored ?? d?.definition;
      if (!d) {
        if (x.base === undefined) problems.push({ severity: "error", code: "address", message: `${where.key}: a cell statement's name is created in Excel (name the cell there, then pull)`, ...where });
        // Deleted in Excel: compared with the lockfile below (the definition does not count).
        else x.stored ??= "";
      }
      continue;
    }
    try {
      const r = compileWithDiagnostics(x.s.formula, { ...ctxOf(x.scope), crlf: true });
      for (const d of r.diagnostics) if (d.severity === "error" || d.code !== "prefix-unsure") defer(d.severity === "error" ? "error" : "warning", d.code, d.message);
      // Compared on the text it would have, even with an unknown function in it.
      x.stored = r.diagnostics.some((d) => d.severity === "error") ? compileWithDiagnostics(x.s.formula, { ...ctxOf(x.scope), crlf: true, allowUnknownFunctions: true }).text : r.text;
    } catch (e) {
      // Does not parse: compared as written (an unparsed name is pulled as stored).
      defer("error", "syntax", (e as Error).message);
      x.stored = x.s.formula;
    }
    if (x.s.formula.length >= 8192) defer("error", "length", `${x.s.formula.length} characters; Excel refuses 8,192 or more (probe T15)`);
  }
  const accept = (x: Src): void => {
    problems.push(...x.deferred);
    changed.add(x.key);
  };
  const changed = new Set<string>();

  const curState = (d: DefinedName): NameState => {
    let display = d.definition;
    try {
      display = decompile(d.definition, ctxOf(scopeOf(d)));
    } catch {
      // Unparsed definitions are shown as stored.
    }
    return { name: d.name, scope: scopeOf(d), display: display.split("\r\n").join("\n"), comment: normComment(d.comment), hidden: d.hidden };
  };
  const srcState = (x: Src): NameState & { file: string; line: number } => {
    let display = x.s.formula;
    if (x.cell) {
      // A named cell's definition is its cell: the state the name must have is that definition.
      display = x.stored ?? "";
      try {
        display = decompile(display, ctxOf(x.scope));
      } catch {
        // Shown as stored.
      }
      display = display.split("\r\n").join("\n");
    }
    return { name: x.s.name, scope: x.scope, display, comment: normComment(x.s.doc), hidden: x.s.hidden, file: x.s.file, line: x.s.line };
  };
  /** `def` undefined: the definition does not count (a named cell's: Excel owns it). */
  const sameAsLock = (e: LockEntry, name: string, lockName: string, def: string | undefined, comment: string | null, hidden: boolean): boolean =>
    (def === undefined || e.definition === definitionHashLike(e.definition, def)) && e.comment === commentHashLike(e.comment, comment ?? undefined) && e.hidden === hidden && name === lockName;
  const sameState = (d: DefinedName, x: Src): boolean =>
    definitionHash(d.definition) === definitionHash(x.stored!) && normComment(d.comment) === normComment(x.s.doc) && d.hidden === x.s.hidden && d.name === x.s.name;

  const lockUpgrades = new Map<string, string>();
  const removed = new Map<string, { what: (key: string) => string; to: string | undefined; rewrite?: RenameSites }>(); // lkey in the workbook → what happens to it
  const created: Src[] = [];
  const nameEdits = new Set<string>();
  const spillOf = (x: Src) => (x.cell ? cellPlan.spill.get(scopedKey(x.s.name, x.scope).toLowerCase()) : undefined);
  for (const x of srcs) {
    const key = scopedKey(x.s.name, x.scope);
    const lockE = x.base !== undefined ? locked.get(x.base) : undefined;
    if (!lockE || x.move !== undefined || x.stored === undefined || spillOf(x)?.edited || !sameAsLock(lockE.entry, x.s.name, lockE.name, x.cell ? undefined : x.stored, normComment(x.s.doc), x.s.hidden)) {
      nameEdits.add(key.toLowerCase());
      if (lockE) nameEdits.add(lockE.key.toLowerCase());
    }
  }
  for (const x of srcs) {
    if (x.stored === undefined) continue;
    const key = scopedKey(x.s.name, x.scope);
    const where = { file: x.s.file, line: x.s.line, key };
    const lockE = x.base !== undefined ? locked.get(x.base) : undefined;
    if (lockE && !x.cell && isV2Hash(lockE.entry.definition) && definitionHashV2(x.stored) === lockE.entry.definition) {
      // The source still has what a format-1/2 lockfile locked: compare in format 3, so
      // Excel's spelling of a number (`1E-14` saved as `0.00000000000001`) is no change.
      lockE.entry = { ...lockE.entry, definition: definitionHash(x.stored) };
      lockUpgrades.set(lockE.key, lockE.entry.definition);
    }
    if (!lockE) {
      const d = cur.get(x.key);
      if (d) {
        if (sameState(d, x)) {
          unchanged++;
          inSync.push(srcState(x));
        } else conflicts.push({ kind: "created-in-both", key, message: `${key} is new in the source and was also created in Excel since the last pull, differently`, excel: curState(d), source: srcState(x) });
        continue;
      }
      const bad = invalidName(x.s.name, x.s.formula ?? "");
      if (bad) {
        problems.push({ severity: "error", code: "invalid-name", message: `${key}: ${bad}`, ...where });
        continue;
      }
      created.push(x);
      accept(x);
      changes.push({ op: "set-name", name: x.s.name, scope: x.scope, stored: x.stored, display: x.s.formula, comment: normComment(x.s.doc), hidden: x.s.hidden, fields: ["created"] });
      inSync.push(srcState(x));
      continue;
    }
    const d = cur.get(x.base!);
    // A named cell's definition counts through its `#` only (the cell is Excel's).
    const spill = spillOf(x);
    const srcChanged = spill?.srcChanged === true || x.move !== undefined || !sameAsLock(lockE.entry, x.s.name, lockE.name, x.cell ? undefined : x.stored, normComment(x.s.doc), x.s.hidden);
    if (!d) {
      if (srcChanged) conflicts.push({ kind: "deleted-in-excel", key, message: `${lockE.key} was deleted in Excel since the last pull, and changed in the source`, source: srcState(x) });
      else excelChanges.push({ kind: "deleted", key: lockE.key, message: `${lockE.key} was deleted in Excel since the last pull; the build leaves it deleted (pull again to drop it from the source)` });
      continue;
    }
    const xlChanged = spill?.xlChanged === true || !sameAsLock(lockE.entry, d.name, lockE.name, x.cell ? undefined : d.definition, normComment(d.comment), d.hidden);
    if (!srcChanged) {
      if (xlChanged) excelChanges.push({ kind: "changed", key: lockE.key, message: `${lockE.key} was changed in Excel since the last pull; the build keeps Excel's version (pull again to update the source)` });
      else {
        unchanged++;
        inSync.push(srcState(x));
      }
      continue;
    }
    if (xlChanged && !(x.move === undefined && sameState(d, x))) {
      conflicts.push({ kind: "both-changed", key, message: `${lockE.key} was changed in Excel since the last pull and in the source: resolve by hand (pull into a copy to see Excel's version)`, excel: curState(d), source: srcState(x) });
      continue;
    }
    if (x.move !== undefined) {
      const clash = cur.get(x.key);
      if (clash && clash !== d) {
        problems.push({ severity: "error", code: "rename", message: `${key}: the workbook already has a name ${scopedKey(clash.name, scopeOf(clash))}`, ...where });
        continue;
      }
      if (x.move === "rename") {
        const bad = invalidName(x.s.name, x.s.formula ?? "");
        if (bad) {
          problems.push({ severity: "error", code: "invalid-name", message: `${key}: ${bad}`, ...where });
          continue;
        }
        const rename: RenameName = { op: "rename-name", scope: scopeOf(d), from: d.name, to: x.s.name };
        const found = sites.get(x.base!);
        if (found && found.counts.cells + found.counts.formats + found.counts.validations + found.counts.names > 0) rename.references = found.counts;
        changes.push(rename);
        if (lkey("", scopeOf(d)) !== lkey("", x.scope)) changes.push({ op: "rescope-name", name: x.s.name, from: scopeOf(d), to: x.scope });
        removed.set(x.base!, { what: (k) => `renaming ${k} to ${x.s.name}`, to: undefined, rewrite: found });
      } else {
        changes.push({ op: "rescope-name", name: d.name, from: scopeOf(d), to: x.scope });
        removed.set(x.base!, { what: (k) => `moving ${k} to ${x.scope === null ? "workbook scope" : `sheet ${x.scope}`}`, to: x.key });
      }
    }
    const fields: NameField[] = [];
    // The workbook's definition after the renames' rewrite (stretch G): what the build keeps.
    const kept = (!x.cell && renamedText(d.definition, scopeOf(d) ?? undefined)) || d.definition;
    const defChanged = definitionHash(kept) !== definitionHash(x.stored);
    if (defChanged) fields.push("definition");
    if (normComment(d.comment) !== normComment(x.s.doc)) fields.push("comment");
    if (d.hidden !== x.s.hidden) fields.push("hidden");
    if (d.name !== x.s.name && x.move !== "rename") fields.push("spelling");
    if (fields.length > 0) {
      const set: Change = {
        op: "set-name",
        name: x.s.name,
        scope: x.scope,
        // An unchanged definition keeps Excel's text, spacing included.
        stored: defChanged ? x.stored : kept,
        display: x.cell ? srcState(x).display : x.s.formula,
        comment: normComment(x.s.doc),
        hidden: x.s.hidden,
        fields,
      };
      if (defChanged && !x.cell && equalModuloWhitespace(curState(d).display, x.s.formula)) set.repair = true;
      changes.push(set);
    }
    if (kept !== d.definition && !defChanged && x.move === undefined) for (const k of [key, lockE.key]) renamedOnly.add(k.toLowerCase());
    if (fields.length > 0 || x.move !== undefined) accept(x);
    inSync.push(srcState(x));
  }

  // A named cell statement removed from the source is refused, not taken as a delete (the
  // author's choice, 2026-10-05): the name belongs to the layout Excel owns, and a stray
  // deleted line must not cost a name. Clearing the cell is `Name @C6 = ;`.
  const cellNames = new Map<string, LockCell>();
  for (const c of Object.values(lock.cells ?? {})) {
    if (c.name === undefined) continue;
    const { name, scope } = splitNameKey(c.name);
    cellNames.set(lkey(name, scope), c);
  }
  // Names the source no longer has: delete them, unless Excel changed them meanwhile.
  for (const [k, l] of locked) {
    if (claimed.has(k)) continue;
    nameEdits.add(l.key.toLowerCase());
    const d = cur.get(k);
    if (!d) continue;
    const cell = cellNames.get(k);
    if (cell) {
      problems.push({
        severity: "error",
        code: "statement-removed",
        key: l.key,
        message: `${l.key}: the cell statement ${l.name} @${cell.sheet}!${cell.range} was removed from the source. To clear the cell, write ${l.name} @${cell.range} = ; — to remove the name, delete it in Excel (Name Manager), then pull`,
      });
      continue;
    }
    if (!sameAsLock(l.entry, d.name, l.name, d.definition, normComment(d.comment), d.hidden)) {
      conflicts.push({ kind: "deleted-in-source", key: l.key, message: `${l.key} is gone from the source but was changed in Excel since the last pull`, excel: curState(d) });
      continue;
    }
    changes.push({ op: "delete-name", name: d.name, scope: scopeOf(d) });
    removed.set(k, { what: (key) => `deleting ${key}`, to: undefined });
  }
  for (const [k, d] of cur) {
    if (!locked.has(k) && !srcByKey.has(k)) excelChanges.push({ kind: "created", key: scopedKey(d.name, scopeOf(d)), message: `${scopedKey(d.name, scopeOf(d))} was created in Excel since the last pull; the build keeps it (pull again to bring it into the source)` });
  }

  checkReferences(wb, srcs, cur, removed, created, changed, problems, cellPlan.changes, source.cells);
  changes.push(...cellPlan.changes);

  return {
    changeSet: { format: CHANGESET_FORMAT, workbook: input.fileName, changes: orderChanges(changes) },
    conflicts,
    problems: mergeProblems(checked, problems),
    excelChanges,
    unchanged,
    inSync,
    unchangedCells: cellPlan.unchanged,
    cellsInSync: cellPlan.inSync,
    lockUpgrades: { names: lockUpgrades, cells: cellPlan.lockUpgrades },
    sourceEdits: { names: nameEdits, cells: cellPlan.sourceEdits, spill: new Set([...cellPlan.spill].filter(([, s]) => s.edited).map(([k]) => k)) },
    implicitSpill: cellPlan.implicitSpill,
    renamedOnly,
  };
}

/**
 * The lockfile with the format-3 hash in place of each format-1/2 (`sha256:`) definition
 * or cell hash that the plan's source still matched. The match shows the locked state was
 * the source's, so its format-3 hash is known exactly; the other old hashes stay (they
 * are of a state no longer at hand) and are compared in their own format.
 */
export function upgradeLockfile(lock: Lockfile, plan: Pick<BuildPlan, "lockUpgrades">): Lockfile {
  const { names: un, cells: uc } = plan.lockUpgrades;
  if (un.size === 0 && uc.size === 0) return lock;
  const names: Record<string, LockEntry> = {};
  for (const [k, e] of Object.entries(lock.names)) names[k] = un.has(k) ? { ...e, definition: un.get(k)! } : e;
  const out: Lockfile = { ...lock, names };
  if (lock.cells !== undefined) {
    const cells: Record<string, LockCell> = {};
    for (const [k, c] of Object.entries(lock.cells)) cells[k] = uc.has(k) ? { ...c, formula: uc.get(k)! } : c;
    out.cells = cells;
  }
  return out;
}

/**
 * D4: a rename, a deletion or a scope change must not leave a reference behind. Cells,
 * conditional formats, validations, Table columns, charts and names kept from the
 * workbook refer to names by text, which the build does not rewrite (that is stretch G):
 * a reference that would no longer reach the same name refuses the build, with the list
 * of places.
 *
 * The readers counted are the workbook's as this build leaves it (feedback 2026-10-07):
 * a name the source defines is checked on its new text, a cell whose formula the build
 * sets or clears on its new formula (or not at all), and the names the build creates are
 * there for them. So one build can move a cell from an old function to a new one and
 * delete the old one.
 */
function checkReferences(
  wb: WorkbookSnapshot,
  srcs: readonly Src[],
  cur: ReadonlyMap<string, DefinedName>,
  removed: ReadonlyMap<string, { what: (key: string) => string; to: string | undefined; rewrite?: RenameSites }>,
  created: readonly Src[],
  changed: ReadonlySet<string>,
  problems: SourceProblem[],
  cellChanges: readonly Change[],
  sourceCells: readonly SourceCell[],
): void {
  const curKeys = [...cur.entries()].map(([k, d]) => ({ name: d.name, scope: scopeOf(d) ?? undefined, key: k }));
  const curResolver = new NameResolver(curKeys);
  const target = curKeys.filter((k) => !removed.has(k.key));
  for (const x of srcs) if (x.stored !== undefined && !target.some((t) => t.key === x.key)) target.push({ name: x.s.name, scope: x.scope ?? undefined, key: x.key });
  for (const x of created) if (!target.some((t) => t.key === x.key)) target.push({ name: x.s.name, scope: x.scope ?? undefined, key: x.key });
  const targetResolver = new NameResolver(target);
  const tables = new Set(wb.tables.map((t) => t.displayName.toLowerCase()));

  if (removed.size > 0) {
    const broken = new Map<string, string[]>();
    const note = (curKey: string, place: string): void => {
      const l = broken.get(curKey) ?? [];
      if (!l.includes(place)) l.push(place);
      broken.set(curKey, l);
    };
    const parsed = new Map<string, Expr | null>();
    const body = (text: string): Expr | undefined => {
      let b = parsed.get(text);
      if (b === undefined) parsed.set(text, (b = tryParse(text).formula?.body ?? null));
      return b ?? undefined;
    };
    /** `rewritable`: a formula a rename's reference rewrite reaches (stretch G): it does not break. */
    const check = (uses: NameUse[], home: string | undefined, place: string, rewritable = false): void => {
      for (const u of uses) {
        const was = curResolver.resolve(u, home);
        if (was === undefined || !removed.has(was)) continue;
        if (rewritable && removed.get(was)!.rewrite) continue;
        const now = targetResolver.resolve(u, home);
        if (now === undefined || now !== removed.get(was)!.to) note(was, place);
      }
    };
    const srcKeys = new Set(srcs.map((x) => x.key));
    // The cells this build gives a formula or clears: their old formulas are no readers.
    const rewritten = new Map<string, Rect[]>();
    for (const c of cellChanges) {
      if (c.op !== "set-cell-formula" && c.op !== "clear-cell-formula") continue;
      const r = parseCellRange(c.range);
      if (!r) continue;
      const l = rewritten.get(c.sheet.toLowerCase()) ?? [];
      l.push(r);
      rewritten.set(c.sheet.toLowerCase(), l);
    }
    const sharedMasters = new Set<string>();
    for (const s of wb.sheets) for (const f of s.formulas) if (f.kind === "shared-master") sharedMasters.add(`${s.name.toLowerCase()}!${f.cell}`);
    const isRewritten = (sheet: string, ref: string, range: string | undefined): boolean => {
      // A shared master's text is its whole group's; an array's lives at its anchor alone.
      const ext = parseCellRange(range !== undefined && sharedMasters.has(`${sheet.toLowerCase()}!${ref}`) ? range : ref);
      return ext !== undefined && (rewritten.get(sheet.toLowerCase()) ?? []).some((r) => r.r1 <= ext.r1 && r.c1 <= ext.c1 && ext.r2 <= r.r2 && ext.c2 <= r.c2);
    };
    // A rename that rewrites references (stretch G): cells, formats, validations and
    // names are rewritten, so they do not break; the places the rewrite cannot reach do.
    for (const [k, r] of removed) {
      if (!r.rewrite) continue;
      for (const p of [...r.rewrite.blocked, ...r.rewrite.unparsed.map((u) => `${u} (does not parse)`)]) note(k, p);
    }
    for (const site of collectSites(wb)) {
      if (site.where.kind === "name") {
        // A name the source defines is checked on its new text below.
        const k = lkey(site.where.name ?? "", site.where.sheet ?? null);
        if (srcKeys.has(k) || removed.has(k)) continue;
      }
      const w = site.where;
      if (w.kind === "cell" && isRewritten(w.sheet!, w.ref!, w.range)) continue;
      const b = body(site.stored);
      if (!b) continue;
      const place =
        w.kind === "cell" ? `cell ${w.sheet}!${w.ref}` : w.kind === "cf" ? `conditional format ${w.sheet}!${w.ref}` : w.kind === "dv" ? `validation ${w.sheet}!${w.ref}` : w.kind === "table" ? `Table column ${w.name}` : `name ${w.key}`;
      check(nameUses(b), site.home, place, w.kind !== "table");
    }
    // The new formulas of those cells read the workbook as the build leaves it.
    for (const c of cellChanges) {
      if (c.op !== "set-cell-formula") continue;
      const b = body(c.stored);
      if (b) check(nameUses(b), c.sheet, `cell ${c.sheet}!${c.range}`);
    }
    for (const chart of wb.charts) {
      for (const f of chart.formulas) {
        const b = body(f.text);
        if (!b) continue;
        const uses: NameUse[] = [];
        walk(b, (n) => {
          const q = n.kind === "name" ? n.qual : undefined;
          if (n.kind !== "name" || (q?.book !== undefined && q.book !== "[0]") || q?.sheet2 !== undefined) return;
          uses.push({ id: n.id, sheet: q?.sheet, span: n.span });
        });
        check(uses, chart.sheet?.name, `chart ${chart.part}`);
      }
    }
    // A cell's statement in the source, for the place a reader is changed.
    const statementOf = (place: string): SourceCell | undefined => {
      const m = /^cell (.+)!([A-Z]+[0-9]+)/.exec(place);
      if (!m) return undefined;
      const at = parseCellRange(m[2]!);
      if (!at) return undefined;
      return sourceCells.find((c) => {
        const r = parseCellRange(c.address.range);
        return c.address.sheet?.toLowerCase() === m[1]!.toLowerCase() && r !== undefined && r.r1 <= at.r1 && at.r1 <= r.r2 && r.c1 <= at.c1 && at.c1 <= r.c2;
      });
    };
    for (const [k, places] of broken) {
      const d = cur.get(k)!;
      const key = scopedKey(d.name, scopeOf(d));
      const one = places.length === 1;
      const located = places.map((p) => {
        const s = statementOf(p);
        return s ? `${p} (${s.file}:${s.line})` : p;
      });
      // Cells are edited in the source. Formats, validations, Table columns and charts are
      // Excel's, and a name counted here is one the source lacks (made in Excel since the
      // last pull; the source's names are checked on their own text).
      const excelOnly = places.filter((p) => !p.startsWith("cell ")).length;
      const rewrite = removed.get(k)!.rewrite;
      const advice = rewrite
        ? `the build rewrites the name in cell formulas, conditional formats, validations and other names, but cannot reach ${one ? "that place" : "these places"}: rename it in Excel's Name Manager instead (it rewrites them too), then pull; or keep the name`
        : excelOnly === 0
          ? `change ${one ? "that formula" : "these formulas"} in the source, or keep the name`
          : excelOnly === places.length
            ? `edit ${one ? "it" : "them"} in Excel and pull, or keep the name`
            : "change the cells in the source and the others in Excel (then pull), or keep the name";
      const own = srcs.find((x) => x.base === k && x.move !== undefined)?.s;
      const first = places.map(statementOf).find((s) => s !== undefined) ?? (rewrite ? own : undefined);
      problems.push({
        severity: "error",
        code: "in-use",
        message: `${removed.get(k)!.what(key)} refused: ${places.length} place${one ? "" : "s"} in the workbook ${one ? "refers" : "refer"} to it by name and would break: ${located.slice(0, 20).join(", ")}${places.length > 20 ? ", …" : ""}. ${advice[0]!.toUpperCase()}${advice.slice(1)}`,
        key,
        sites: places,
        ...(first ? { file: first.file, line: first.line } : {}),
      });
    }
    // Readers the rewrite would make read another name (a LET variable, a sheet's local name, an unknown name of the new spelling).
    for (const [k, r] of removed) {
      if (!r.rewrite || r.rewrite.captured.length === 0) continue;
      const d = cur.get(k)!;
      const key = scopedKey(d.name, scopeOf(d));
      const n = r.rewrite.captured.length;
      const at = srcs.find((x) => x.base === k);
      problems.push({
        severity: "error",
        code: "rename-capture",
        message: `${r.what(key)} refused: in ${n} place${n === 1 ? "" : "s"} the new name would read something else: ${r.rewrite.captured.slice(0, 20).join(", ")}${n > 20 ? ", …" : ""}. Choose another name, or rename the variable or name it collides with`,
        key,
        sites: r.rewrite.captured,
        ...(at ? { file: at.s.file, line: at.s.line } : {}),
      });
    }
  }

  // Source definitions must not refer to a name the build removes.
  for (const x of srcs) {
    if (x.stored === undefined) continue;
    const b = tryParse(x.s.formula).formula?.body;
    if (!b) continue;
    for (const u of nameUses(b)) {
      if (targetResolver.resolve(u, x.scope ?? undefined) !== undefined) continue;
      if (u.sheet === undefined && tables.has(u.id.toLowerCase())) continue;
      const was = curResolver.resolve(u, x.scope ?? undefined);
      const key = scopedKey(x.s.name, x.scope);
      const ref = u.sheet === undefined ? u.id : `${u.sheet}!${u.id}`;
      if (was !== undefined && removed.has(was)) {
        problems.push({ severity: "error", code: "in-use", message: `${key} refers to ${ref}, which this build removes (${removed.get(was)!.what(ref)})`, file: x.s.file, line: x.s.line, key });
      } else if (changed.has(x.key)) {
        problems.push({ severity: "warning", code: "unknown-name", message: `${key} refers to ${ref}, which is not a name in the workbook or the source: Excel shows #NAME?`, file: x.s.file, line: x.s.line, key });
      }
    }
  }
}

/**
 * The checker's errors and the plan's own problems, one per place. A reader that a rename,
 * re-scope or deletion in the source breaks is an error for both (`in-use`, §14 issue 10):
 * the plan says it once per name with every place and what to do, so its problem stands in
 * for the checker's per reader. Elsewhere the checker's finding at a statement wins.
 */
function mergeProblems(checked: readonly SourceProblem[], own: readonly SourceProblem[]): SourceProblem[] {
  const perName = (p: SourceProblem) => p.code === "in-use" && p.sites !== undefined;
  const kept = own.some(perName) ? checked.filter((c) => c.code !== "in-use") : [...checked];
  return [...kept, ...own.filter((p) => !(p.file !== undefined && kept.some((c) => c.severity === "error" && c.file === p.file && c.line === p.line)))];
}

/** A finding of the source checker placed in a project file (1-based line and column). */
export interface SourceFinding {
  severity: "error" | "warning" | "info" | "hint";
  code: string;
  message: string;
  file: string;
  line: number;
  column: number;
  key?: string;
  fix?: { title: string; start: number; end: number; text: string };
  /** In a formula as the workbook has it since the last pull (a warning the build leaves out). */
  inWorkbook?: true;
}

/**
 * Everything the source checker (check.ts) finds in a project's files, at every severity:
 * what the editor shows as you type, for `xln check` on a project, so the command line
 * and the editor report the same findings (feedback 2026-10-07).
 */
export function sourceFindings(files: Readonly<Record<string, string>>, ctx: CheckContext): SourceFinding[] {
  const model = new SourceModel();
  for (const [path, text] of Object.entries(files)) if (isSourcePath(path)) model.setFile(path, text);
  const out: SourceFinding[] = [];
  for (const p of checkProject(model, ctx)) {
    const at = model.files.get(p.path)!.lines.position(p.start);
    const f: SourceFinding = { severity: p.severity, code: p.code ?? "source", message: p.message, file: p.path, line: at.line + 1, column: at.character + 1, ...(p.key !== undefined ? { key: p.key } : {}), ...(p.inWorkbook ? { inWorkbook: true as const } : {}) };
    const fix = p.fix ?? p.fixes?.[0];
    if (fix) f.fix = { title: fix.title, start: fix.start, end: fix.end, text: fix.text };
    out.push(f);
  }
  // Other files below names/ (the caller lists them, text or not): the build reads none of
  // them, so one there is a mistake the author should hear about (M3e).
  for (const path of Object.keys(files).sort()) {
    if (isSourcePath(path)) continue;
    const message = strayFile(path, ctx.sheets);
    if (message !== undefined) out.push({ severity: "error", code: "stray-file", message, file: path, line: 1, column: 1 });
  }
  return out;
}

/** The checker's errors on a project's files (check.ts), as build problems. */
export function sourceErrors(files: Readonly<Record<string, string>>, ctx: CheckContext): SourceProblem[] {
  return sourceProblems(files, ctx).filter((p) => p.severity === "error");
}

/** The checker's errors and warnings on a project's files, as build problems. */
function sourceProblems(files: Readonly<Record<string, string>>, ctx: CheckContext): SourceProblem[] {
  const out: SourceProblem[] = [];
  for (const f of sourceFindings(files, ctx)) {
    if (f.severity !== "error" && (f.severity !== "warning" || f.inWorkbook)) continue;
    const sp: SourceProblem = { severity: f.severity, code: f.code, message: f.message, file: f.file, line: f.line, ...(f.key !== undefined ? { key: f.key } : {}) };
    if (f.fix) sp.fix = { start: f.fix.start, end: f.fix.end, text: f.fix.text };
    out.push(sp);
  }
  return out;
}
