// `lib status` (F2): each library function against its copy in a workbook or a project,
// three-way (decided 2026-10-07): the copy, its library base (the library version it came
// from: `@from(lib #…)` in the source, `lib#…` in the workbook's provenance tag) and the
// library's current definition. All compared in stored form, modulo layout and the
// spelling of numbers (as the lockfile hashes them).
//
//   identical     the copy's definition equals the library's (whatever its base says)
//   outdated      the copy is still its base; the library has a newer version
//   modified      the copy was edited since its base, which is still the library's version
//   both-changed  the copy and the library both moved away from the base, differently
//   differs       no base recorded (an old tag without `lib#`, or the author removed
//                 `@from`): which side moved cannot be told
//   missing       the library has it, the workbook or project does not
//   local only    a module LAMBDA (`FN.X`, `names/FN.xln`) the library does not have, of a
//                 module the library has (its prefix): a candidate for publishing. Other
//                 modules' LAMBDAs (`IN.*` input readers) are only counted, per module.
//
// Before the base, the tag's hash (what was built) stood in for it: after a local edit and
// a build the tag matched the edited copy, so an edit read as "outdated" and the status
// steered the author to take the library's version, throwing the edit away (feedback
// 2026-10-07). The tag still shows; only the base decides.
//
// A workbook's copy is its defined name, its base the tag's `lib#`; a project's is the name
// in its source, its base the source's `@from`. Only definitions decide the state: the doc
// comment of a library function is generated from its header and may be shortened.
//
// A project's status is its source's, so with the workbook at hand each item also says
// whether the workbook has it as the source does: "not built yet" when the source has it
// and the workbook lacks it or has it otherwise (definition, doc comment or library base),
// or when the source deleted it and the workbook still has it (feedback 2026-10-07: a
// status on the project looked done while every build had been refused).
//
// The base is only a hash. Its text, for the three-way diff of a copy changed on both
// sides, is found among the versions at hand (`LibCopy.known`): first the project's
// `library-bases/<hash>.json`, kept by the actions that wrote the `@from` (`withBases`);
// else the workbook's copy for a project, a backup the caller adds with
// `withKnownVersions`; without any the diff is the copy against the library.
//
// An identical copy with no base is fine today but reads "differs" as soon as either side
// changes: its note says so and names the explicit action, *Record library base*
// (`xln lib base`). Status never records one itself (no hidden transformations).

import type { WorkbookSnapshot } from "../file/types.js";
import { equalModuloWhitespace } from "../lang/format.js";
import { compileWithDiagnostics, decompileWithDiagnostics } from "../lang/transform.js";
import { classify } from "../project/classify.js";
import { definitionHash } from "../project/lockfile.js";
import { proposeModules } from "../project/modules.js";
import { formatProvenanceTag, libraryHash, moduleOfPath, sourceHash, sourceHashV1, splitProvenance, type ProvenanceTag } from "../project/provenance.js";
import { readSourceProject } from "../build/source.js";
import { definitionDiff, diffText, type DiffLine } from "./diff.js";
import { libraryStored, type Library } from "./library.js";
import { readBases, BASES_DIR, type LibraryBase } from "./bases.js";
import type { LibraryProblem } from "./lambdaFile.js";

export type LibState = "identical" | "outdated" | "modified" | "both-changed" | "differs" | "missing" | "local-only";

export const LIB_STATES: readonly LibState[] = ["outdated", "modified", "both-changed", "differs", "missing", "identical", "local-only"];

/** A version of a definition known besides the copy (the workbook's, a backup's): a candidate for the base's text. */
export interface KnownVersion {
  stored: string;
  display: string;
  /** Where it was found, for the report: `library-bases/…`, `the workbook`, `the backup`. */
  source?: string;
}

/** A name as the workbook or the project has it. */
export interface LibCopy {
  name: string;
  /** Stored form (for the comparison). */
  stored: string;
  /** Display form (for the diff). */
  display: string;
  /** The comment / doc comment without its tag. */
  comment: string | undefined;
  tag?: ProvenanceTag;
  /** Whether the tag's hash matches the copy (its definition and comment): false when edited in Excel since the build. */
  tagMatches?: boolean;
  /** The library base: the source's `@from(lib #…)` (project), the tag's `lib#…` (workbook). */
  base?: string;
  /** Other versions of this name at hand, where the base's text may be found. */
  known?: KnownVersion[];
  /** The module it belongs to (`FN`), if any. */
  module: string | undefined;
  lambda: boolean;
  /** For a project: the names file and line. */
  file?: string;
  line?: number;
  /** For a project with its workbook: the workbook lacks the name (`new`), or has it otherwise (`edited`): a build would write it. */
  unbuilt?: "new" | "edited";
}

/** How a project's copy stands against its workbook: the source has it, the build has not written it yet. */
export type Unbuilt = "new" | "edited" | "deleted";

const UNBUILT_NOTE: Record<Unbuilt, string> = {
  new: "in the source, not built yet",
  edited: "edited in the source, not built yet",
  deleted: "deleted in the source, still in the workbook: not built yet",
};

export interface LibStatusItem {
  name: string;
  state: LibState;
  /** The library file (absent for local only). */
  libraryPath?: string;
  /** Where the copy is (a project's names file and line). */
  file?: string;
  line?: number;
  /** The copy's provenance tag, as written in the comment. */
  tag?: string;
  /** The copy's library base (`@from(lib #…)` / `lib#…`), when it records one. */
  base?: string;
  /** The library's version (`libraryHash` of its definition): what Insert, Take and Publish record. */
  libraryHash?: string;
  /** The copy's version, on the same scale. */
  copyHash?: string;
  /** modified / differs / outdated / both-changed: the copy (`-`) against the library (`+`). */
  diff?: DiffLine[];
  /** The copy and the library definitions in display form, for a diff view. */
  copyDisplay?: string;
  libraryDisplay?: string;
  /** both-changed: the base's definition (display form), when a version at hand has its hash. */
  baseDisplay?: string;
  /** Where the base's text was found (`library-bases/353921.json`, `the workbook`, `the backup`). */
  baseSource?: string;
  /** identical, but no library base is recorded: *Record library base* would write one. */
  noBase?: boolean;
  /** both-changed with the base at hand: base (`-`) → copy (`+`), and base → library. */
  baseDiffs?: { copy: DiffLine[]; library: DiffLine[] };
  /** One line on why. */
  note: string;
  /** A project's copy the workbook does not have as the source does (the state is the source's). */
  unbuilt?: Unbuilt;
}

export interface LibStatusReport {
  /** What was compared: the workbook's file name or the project folder, as the caller names it. */
  target: string;
  kind: "workbook" | "project";
  /** The library folder, as the caller names it. */
  library: string;
  items: LibStatusItem[];
  counts: Record<LibState, number>;
  /** LAMBDAs of modules the library has none of (not library candidates), by module. */
  otherModules: Record<string, number>;
  /** Problems of the library's files. */
  problems: LibraryProblem[];
  /**
   * A project's: the workbook it is built into, compared with (`projectLibraryStatus`);
   * null when it was not at hand (what is built cannot be told). Absent for a workbook.
   */
  workbook?: string | null;
  /** A project's items the workbook does not have as the source does. */
  unbuilt?: number;
}

const lower = (s: string) => s.toLowerCase();

/** A workbook's definition in display form, for diffs. */
function displayOf(stored: string, allNames: readonly string[]): string {
  return decompileWithDiagnostics(stored, { names: [...allNames] }).text;
}

/** The workbook-scoped names of a workbook, as copies. */
export function workbookCopies(wb: WorkbookSnapshot): LibCopy[] {
  const regular = wb.definedNames.filter((d) => !d.isXlPrefixed && !d.isBuiltIn);
  const allNames = [...new Set(regular.map((d) => d.name))];
  const modules = proposeModules(allNames);
  const out: LibCopy[] = [];
  for (const d of regular) {
    if (d.scope.kind !== "workbook") continue;
    const raw = d.comment === undefined || d.comment === "" ? undefined : d.comment.split("\r\n").join("\n");
    const split = raw === undefined ? { comment: undefined } : splitProvenance(raw);
    const c: LibCopy = {
      name: d.name,
      stored: d.definition,
      display: displayOf(d.definition, allNames),
      comment: split.comment,
      module: modules.get(d.name),
      lambda: classify(d.definition).kind === "lambda",
    };
    if (split.tag) {
      c.tag = split.tag;
      c.tagMatches = sourceHash(d.definition, split.comment) === split.tag.hash || sourceHashV1(d.definition, split.comment) === split.tag.hash;
      if (split.tag.lib !== undefined) c.base = split.tag.lib;
    }
    out.push(c);
  }
  return out;
}

/**
 * The versions a snapshot has of the copies' names (a backup made by a build, the
 * workbook beside a project): candidates for the base's text in a three-way diff.
 */
export function withKnownVersions(copies: LibCopy[], wb: WorkbookSnapshot, source = "the backup"): LibCopy[] {
  const regular = wb.definedNames.filter((d) => !d.isXlPrefixed && !d.isBuiltIn && d.scope.kind === "workbook");
  const allNames = [...new Set(regular.map((d) => d.name))];
  const byName = new Map(regular.map((d) => [lower(d.name), d]));
  for (const c of copies) {
    const d = byName.get(lower(c.name));
    if (!d) continue;
    (c.known ??= []).push({ stored: d.definition, display: displayOf(d.definition, allNames), source });
  }
  return copies;
}

/**
 * The kept texts of the copies' bases (`library-bases/`, read with `readBases`), put first
 * among the versions at hand: they are the base by construction.
 */
export function withBases(copies: LibCopy[], bases: ReadonlyMap<string, LibraryBase>): LibCopy[] {
  for (const c of copies) {
    const b = c.base !== undefined ? bases.get(c.base) : undefined;
    if (b) (c.known ??= []).unshift({ stored: b.stored, display: b.display, source: `${BASES_DIR}/${b.hash}.json` });
  }
  return copies;
}

/**
 * The workbook-scoped names of a project's source, as copies; with the workbook it was
 * built into, each carries that workbook's tag for the name, checked against the source.
 */
export function projectCopies(files: Readonly<Record<string, string>>, wb?: WorkbookSnapshot): LibCopy[] {
  const source = readSourceProject(files);
  const allNames = [...new Set(source.names.map((n) => n.name))];
  const tags = new Map<string, ProvenanceTag>();
  const built = new Map<string, { stored: string; comment: string | undefined }>();
  const wbNames: string[] = [];
  if (wb) {
    for (const d of wb.definedNames) {
      if (d.isXlPrefixed || d.scope.kind !== "workbook") continue;
      wbNames.push(d.name);
      const split = d.comment ? splitProvenance(d.comment.split("\r\n").join("\n")) : { comment: undefined };
      built.set(lower(d.name), { stored: d.definition, comment: split.comment });
      if (split.tag) tags.set(lower(d.name), split.tag);
    }
  }
  const out: LibCopy[] = [];
  for (const n of source.names) {
    if (n.scope !== undefined || n.cell) continue;
    const stored = compileWithDiagnostics(n.formula, { names: allNames, allowUnknownFunctions: true }).text;
    const c: LibCopy = {
      name: n.name,
      stored,
      display: n.formula,
      comment: n.doc,
      module: moduleOfPath(n.file),
      lambda: classify(stored).kind === "lambda",
      file: n.file,
      line: n.line,
    };
    if (n.libBase !== undefined) c.base = n.libBase;
    const tag = tags.get(lower(n.name));
    if (tag) {
      c.tag = tag;
      c.tagMatches = sourceHash(stored, n.doc) === tag.hash || sourceHashV1(stored, n.doc) === tag.hash;
    }
    if (wb) {
      const b = built.get(lower(n.name));
      if (!b) c.unbuilt = "new";
      else {
        // A tag's library base differs from the source's `@from`: the next build rewrites
        // the tag. Without a tag (an old build, a comment too long for one) there is nothing to compare.
        const baseChanged = tag !== undefined && tag.lib !== n.libBase;
        if (definitionHash(b.stored) !== definitionHash(stored) || (b.comment ?? "") !== (n.doc ?? "").split("\r\n").join("\n") || baseChanged) c.unbuilt = "edited";
        c.known = [{ stored: b.stored, display: displayOf(b.stored, wbNames), source: "the workbook" }];
      }
    }
    out.push(c);
  }
  return out;
}

/** The workbook-scoped names of a workbook, lower case (`libraryStatus`'s `inWorkbook`). */
export function workbookNameSet(wb: WorkbookSnapshot): Set<string> {
  return new Set(wb.definedNames.filter((d) => !d.isXlPrefixed && d.scope.kind === "workbook").map((d) => lower(d.name)));
}

/**
 * `lib status` of a project: its source against the library, and, with the workbook it is
 * built into, what of it the workbook does not have yet (feedback 2026-10-07: a status on
 * the source looked done while the build had been refused). The bases' texts for a
 * three-way diff: the project's `library-bases/` (among `files`, or `bases` when the caller
 * read them apart), then the workbook, then `backup`, another snapshot of the workbook (the
 * build's `.backup.xlsx`).
 */
export function projectLibraryStatus(
  library: Library,
  files: Readonly<Record<string, string>>,
  wb: WorkbookSnapshot | undefined,
  meta: { target: string; library: string; workbook?: string; backup?: WorkbookSnapshot; bases?: ReadonlyMap<string, LibraryBase> },
): LibStatusReport {
  const copies = projectCopies(files, wb);
  if (meta.backup) withKnownVersions(copies, meta.backup);
  withBases(copies, meta.bases ?? readBases(files));
  const r = libraryStatus(library, copies, { target: meta.target, library: meta.library, kind: "project", ...(wb ? { inWorkbook: workbookNameSet(wb) } : {}) });
  return { ...r, workbook: wb ? (meta.workbook ?? "the workbook") : null };
}

/** The state of one library function given its copy (undefined: missing), three-way on the copy's base. */
export function libState(libStored: string, copy: LibCopy | undefined): LibState {
  if (!copy) return "missing";
  if (equalModuloWhitespace(copy.stored, libStored)) return "identical";
  if (copy.base === undefined) return "differs";
  if (libraryHash(copy.stored) === copy.base) return "outdated";
  if (libraryHash(libStored) === copy.base) return "modified";
  return "both-changed";
}

const NOTE: Record<LibState, string> = {
  identical: "same definition as the library",
  outdated: "unchanged since its library base; the library has a newer version (take it)",
  modified: "edited here since its library base, which is still the library's version (publish it, or take the library's version to undo the edit)",
  "both-changed": "edited here and in the library since its base: compare the three before taking or publishing",
  differs: "no library base recorded (@from in the source, lib# in the tag): cannot tell which side changed",
  missing: "in the library, not here",
  "local-only": "a LAMBDA of a library module the library does not have",
};

/** `lib status` over copies (`workbookCopies`, `projectCopies`). */
export function libraryStatus(
  library: Library,
  copies: readonly LibCopy[],
  meta: {
    target: string;
    kind: "workbook" | "project";
    library: string;
    /** For a project: the workbook's workbook-scoped names, lower case (`workbookNameSet`): a library function the source deleted and the workbook still has is not built yet. */
    inWorkbook?: ReadonlySet<string>;
  },
): LibStatusReport {
  const { inWorkbook, ...head } = meta;
  const byName = new Map(copies.map((c) => [lower(c.name), c]));
  const names = [...new Set([...copies.map((c) => c.name), ...library.functions.map((f) => f.name)])];
  const items: LibStatusItem[] = [];
  const unbuilt = (item: LibStatusItem, u: Unbuilt | undefined): void => {
    if (!u) return;
    item.unbuilt = u;
    item.note = `${item.note}; ${UNBUILT_NOTE[u]}`;
  };
  for (const fn of library.functions) {
    const copy = byName.get(lower(fn.name));
    const libStored = libraryStored(fn, names);
    const state = libState(libStored, copy);
    const libHash = libraryHash(libStored);
    const item: LibStatusItem = { name: fn.name, state, libraryPath: fn.path, libraryHash: libHash, note: NOTE[state] };
    if (copy) {
      if (copy.file !== undefined) item.file = copy.file;
      if (copy.line !== undefined) item.line = copy.line;
      if (copy.tag) item.tag = formatProvenanceTag(copy.tag);
      if (copy.base !== undefined) item.base = copy.base;
      item.copyHash = libraryHash(copy.stored);
      // The same change made on both sides leaves an older base behind: say so, nothing to do.
      if (state === "identical" && copy.base !== undefined && copy.base !== libHash) item.note += `; its recorded base #${copy.base} is older than the library's #${libHash}`;
      // Identical today, "differs" after the first change on either side: say what records a base.
      if (state === "identical" && copy.base === undefined) {
        item.noBase = true;
        item.note += meta.kind === "project" ? "; no base recorded: Record library base" : "; no base recorded: Record library base in the project (xln lib base), then build";
      }
      if (meta.kind === "workbook" && copy.tag && copy.tagMatches === false) item.note += "; edited in Excel since the build that tagged it";
      if (state !== "identical") {
        item.diff = definitionDiff(copy.display, fn.definition);
        item.copyDisplay = copy.display;
        item.libraryDisplay = fn.definition;
      }
      if (state === "both-changed") {
        const base = (copy.known ?? []).find((k) => libraryHash(k.stored) === copy.base);
        if (base) {
          item.baseDisplay = base.display;
          if (base.source !== undefined) item.baseSource = base.source;
          item.baseDiffs = { copy: definitionDiff(base.display, copy.display), library: definitionDiff(base.display, fn.definition) };
        }
      }
    }
    unbuilt(item, copy ? copy.unbuilt : inWorkbook?.has(lower(fn.name)) ? "deleted" : undefined);
    items.push(item);
  }
  const libModules = new Set(library.functions.map((f) => (f.name.indexOf(".") > 0 ? lower(f.name.slice(0, f.name.indexOf("."))) : "")).filter((m) => m !== ""));
  const otherModules: Record<string, number> = {};
  for (const c of copies) {
    if (!c.lambda || c.module === undefined || library.get(c.name)) continue;
    if (!libModules.has(lower(c.module))) {
      otherModules[c.module] = (otherModules[c.module] ?? 0) + 1;
      continue;
    }
    const item: LibStatusItem = { name: c.name, state: "local-only", note: NOTE["local-only"] };
    unbuilt(item, c.unbuilt);
    if (c.file !== undefined) item.file = c.file;
    if (c.line !== undefined) item.line = c.line;
    if (c.base !== undefined) item.base = c.base;
    items.push(item);
  }
  const counts = Object.fromEntries(LIB_STATES.map((s) => [s, 0])) as Record<LibState, number>;
  for (const i of items) counts[i.state]++;
  const report: LibStatusReport = { ...head, items, counts, otherModules, problems: library.problems };
  if (meta.kind === "project" && inWorkbook) report.unbuilt = items.filter((i) => i.unbuilt !== undefined).length;
  return report;
}

const LABEL: Record<LibState, string> = {
  identical: "identical",
  outdated: "outdated",
  modified: "modified",
  "both-changed": "both changed",
  differs: "differs",
  missing: "missing",
  "local-only": "local only",
};

export function libStateLabel(s: LibState): string {
  return LABEL[s];
}

/** The note of a state (the report's heading for it). */
export function libStateNote(s: LibState): string {
  return NOTE[s];
}

export interface LibStatusLink {
  start: number;
  end: number;
  name: string;
  /** `copy`: the name in the project; `library`: the `.lambda` file. */
  target: "copy" | "library";
  path?: string;
}

/**
 * A copy changed on both sides, as text: what the copy changed since the base and what the
 * library changed, each against the base; without the base's text, the copy against the
 * library. `indent` prefixes every line.
 */
export function threeWayText(i: LibStatusItem, indent = ""): string {
  const out: string[] = [];
  if (i.baseDiffs) {
    if (i.baseSource) out.push(`${indent}the base #${i.base}'s text: from ${i.baseSource}`);
    out.push(`${indent}here, since the base #${i.base} (- base, + copy):`);
    out.push(diffText(i.baseDiffs.copy, indent + "  "));
    out.push(`${indent}in the library, since the base #${i.base} (- base, + library #${i.libraryHash}):`);
    out.push(diffText(i.baseDiffs.library, indent + "  "));
  } else {
    out.push(`${indent}the base #${i.base}'s text is not at hand (not in ${BASES_DIR}/, and neither the workbook nor a backup has it); the copy (-) against the library (+):`);
    if (i.diff) out.push(diffText(i.diff, indent + "  "));
  }
  return out.join("\n");
}

/** The report as text (the CLI's output and the editor's read-only document), with the spans of names. */
export function renderLibStatus(r: LibStatusReport, opts: { diffs?: boolean } = {}): { text: string; links: LibStatusLink[] } {
  const diffs = opts.diffs ?? true;
  let text = "";
  const links: LibStatusLink[] = [];
  const line = (s = "") => {
    text += s + "\n";
  };
  const n = r.items.filter((i) => i.state !== "local-only").length;
  line(`xln lib status: ${r.target} (${r.kind}) against ${r.library}`);
  const unbuilt = r.unbuilt ?? 0;
  line(LIB_STATES.map((s) => `${LABEL[s]} ${r.counts[s]}`).join(" · ") + ` (${n} library function${n === 1 ? "" : "s"})` + (unbuilt ? ` · not built yet ${unbuilt}` : ""));
  // A project's status is its source's: say when the workbook is not there yet.
  if (unbuilt) line(`The source differs from ${r.workbook}: ${unbuilt} function${unbuilt === 1 ? " is" : "s are"} not built yet (build to write ${unbuilt === 1 ? "it" : "them"} into the workbook).`);
  else if (r.kind === "project" && r.workbook === null) line("The workbook was not found beside the project: what is built cannot be told.");
  if (r.problems.length) {
    line();
    line("Library problems:");
    for (const p of r.problems) line(`  ${p.severity === "error" ? "error  " : "warning"} ${p.path}${p.line !== undefined ? `:${p.line}` : ""}: ${p.message}`);
  }
  const width = Math.max(10, ...r.items.map((i) => i.name.length));
  for (const state of LIB_STATES) {
    const list = r.items.filter((i) => i.state === state);
    if (list.length === 0) continue;
    line();
    line(`${LABEL[state]} (${list.length}): ${NOTE[state]}`);
    for (const i of list) {
      const start = text.length + 2;
      links.push({ start, end: start + i.name.length, name: i.name, target: i.state === "missing" ? "library" : "copy", ...(i.libraryPath && i.state === "missing" ? { path: i.libraryPath } : {}) });
      const where = i.file !== undefined ? `${i.file}:${i.line}` : "";
      const versions = [i.base !== undefined ? `base #${i.base}` : "", i.libraryHash !== undefined ? `library #${i.libraryHash}` : ""].filter((s) => s !== "").join(" ");
      const extra = [i.tag ?? "", versions, where, i.unbuilt ? `(${UNBUILT_NOTE[i.unbuilt]})` : ""].filter((s) => s !== "").join("  ");
      line(`  ${i.name.padEnd(width)}${extra ? "  " + extra : ""}`.trimEnd());
      if (!diffs) continue;
      if (state === "both-changed") line(threeWayText(i, "      "));
      else if (i.diff && (state === "modified" || state === "differs" || state === "outdated")) line(diffText(i.diff, "      "));
    }
  }
  const others = Object.entries(r.otherModules);
  if (others.length) {
    const total = others.reduce((a, [, k]) => a + k, 0);
    line();
    line(`other module LAMBDAs, not library candidates: ${total} (${others.map(([m, k]) => `${m} ${k}`).join(", ")})`);
  }
  return { text, links };
}

/** Plain JSON for `--json`. */
export function libStatusJson(r: LibStatusReport): unknown {
  return {
    target: r.target,
    kind: r.kind,
    library: r.library,
    counts: r.counts,
    ...(r.workbook !== undefined ? { workbook: r.workbook } : {}),
    ...(r.unbuilt !== undefined ? { unbuilt: r.unbuilt } : {}),
    otherModules: r.otherModules,
    problems: r.problems,
    items: r.items.map((i) => ({
      name: i.name,
      state: i.state,
      ...(i.libraryPath !== undefined ? { libraryPath: i.libraryPath } : {}),
      ...(i.file !== undefined ? { file: i.file, line: i.line } : {}),
      ...(i.tag !== undefined ? { tag: i.tag } : {}),
      ...(i.base !== undefined ? { base: i.base } : {}),
      ...(i.libraryHash !== undefined ? { libraryHash: i.libraryHash } : {}),
      ...(i.copyHash !== undefined ? { copyHash: i.copyHash } : {}),
      ...(i.noBase ? { noBase: true } : {}),
      ...(i.baseSource !== undefined ? { baseSource: i.baseSource } : {}),
      note: i.note,
      ...(i.unbuilt !== undefined ? { unbuilt: i.unbuilt } : {}),
      ...(i.diff ? { diff: diffText(i.diff).split("\n") } : {}),
      ...(i.baseDiffs ? { baseDiff: { copy: diffText(i.baseDiffs.copy).split("\n"), library: diffText(i.baseDiffs.library).split("\n") } } : {}),
    })),
  };
}
