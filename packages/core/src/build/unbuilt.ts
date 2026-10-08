// The pull's guard (decided 2026-10-06): every pull is fresh, so it would replace source
// edits not built yet. Before it does, the project is compared as the build compares it
// (the plan's `sourceEdits`, the source against the lockfile and the workbook as it is
// now) and every such edit is listed; the pull then refuses unless told to discard them.
//
// Without a lockfile the source is compared with the workbook itself (the lockfile a pull
// would write): what differs from it is lost by a pull. With one, an edit counts only where
// the source also differs from the workbook as it is (a browser build's result put in
// place while the lockfile stayed the old one has none). A source that does not parse, or
// that the build would refuse, also holds edits that were never built: its errors are
// listed too.

import { readWorkbook } from "../file/workbook.js";
import { parseLockfile, type Lockfile } from "../project/lockfile.js";
import { moduleOfPath, splitProvenance } from "../project/provenance.js";
import { LOCK_FILE, pullProject } from "../project/pull.js";
import { valueCells } from "./build.js";
import { describeChange, scopedKey, type Change } from "./changes.js";
import { planBuild } from "./plan.js";
import { isSourcePath, readSourceProject } from "./source.js";

export interface UnbuiltEdit {
  /** The name's key (`Sheet!Name`, `Name`), `Sheet!Range` for an unnamed cell; undefined for an error not tied to one. */
  key?: string;
  /** Where the source has it; undefined when the source removed it. */
  file?: string;
  line?: number;
  /** What the source changed: the change the build would make, or what keeps it from making it. */
  what: string;
}

export interface UnbuiltInput {
  /** The workbook as it is now. */
  workbook: Uint8Array;
  fileName: string;
  /** The project files on disk (names files, lockfile), path → text. */
  files: Readonly<Record<string, string>>;
}

const low = (s: string): string => s.toLowerCase();

/** The keys a change touches, in lower case. */
function changeKeys(c: Change): string[] {
  switch (c.op) {
    case "set-name":
    case "delete-name":
      return [scopedKey(c.name, c.scope)];
    case "rename-name":
      return [scopedKey(c.from, c.scope), scopedKey(c.to, c.scope)];
    case "rescope-name":
      return [scopedKey(c.name, c.from), scopedKey(c.name, c.to)];
    case "set-cell-formula":
    case "clear-cell-formula":
      return [`${c.sheet}!${c.range}`, ...(c.name ? [c.name, `${c.sheet}!${c.name}`] : [])];
    case "set-embedded-source":
      return [];
  }
}

/** Source edits a pull would replace: empty when the project has none (or no names files). */
export function unbuiltEdits(input: UnbuiltInput): UnbuiltEdit[] {
  if (!Object.keys(input.files).some(isSourcePath)) return [];
  let lock: Lockfile | undefined;
  const lockText = input.files[LOCK_FILE];
  if (lockText !== undefined) {
    try {
      lock = parseLockfile(lockText);
    } catch (e) {
      return [{ file: LOCK_FILE, what: `the lockfile cannot be read (${(e as Error).message}): the source cannot be compared with the last pull` }];
    }
  }
  const fileName = input.fileName.slice(Math.max(input.fileName.lastIndexOf("/"), input.fileName.lastIndexOf("\\")) + 1);
  // The workbook as a lockfile: what a pull would write. Comparing the source with it says
  // what the source has that the workbook does not.
  const asWorkbook = parseLockfile(pullProject(input.workbook, fileName).files[LOCK_FILE]!);
  const wb = readWorkbook(input.workbook);
  const values = valueCells(input.workbook);
  const plan = planBuild({ workbook: wb, fileName, files: input.files, lock: lock ?? asWorkbook, values, bytes: input.workbook });
  // Unbuilt edits are those where the source differs from the workbook (spec §14 issue 9):
  // the lockfile decides only which side moved. After a browser build the author replaces
  // the workbook with <name>.xln.xlsx while the lockfile stays the one before the build; the
  // source and the workbook then agree, and a pull loses nothing.
  const differs = (() => {
    if (!lock) return undefined;
    const p = planBuild({ workbook: wb, fileName, files: input.files, lock: asWorkbook, values, bytes: input.workbook });
    return new Set([...p.sourceEdits.names, ...p.sourceEdits.cells, ...p.sourceEdits.spill]);
  })();

  const out: UnbuiltEdit[] = [];
  // Errors in the source: the build refuses it, so it is not what the workbook has. A
  // reference a change would break (`sites`) is placed at a reader, which is no edit: the
  // edit itself is listed below.
  for (const p of plan.problems) if (p.severity === "error" && p.file !== undefined && p.sites === undefined) out.push({ ...(p.key ? { key: p.key } : {}), file: p.file, ...(p.line !== undefined ? { line: p.line } : {}), what: p.message });

  const src = readSourceProject(input.files);
  const where = new Map<string, { key: string; file: string; line: number }>();
  for (const c of src.cells) {
    const key = c.name === undefined ? `${c.address.sheet ?? ""}!${c.address.range}` : scopedKey(c.name, c.scope);
    where.set(low(key), { key, file: c.file, line: c.line });
  }
  for (const n of src.names) {
    const key = scopedKey(n.name, n.scope);
    where.set(low(key), { key, file: n.file, line: n.line });
  }
  const what = new Map<string, string[]>();
  for (const c of plan.changeSet.changes) {
    for (const k of changeKeys(c)) {
      const l = what.get(low(k)) ?? [];
      if (!l.includes(describeChange(c))) l.push(describeChange(c));
      what.set(low(k), l);
    }
  }
  for (const c of plan.conflicts) {
    const l = what.get(low(c.key)) ?? [];
    l.push(c.message);
    what.set(low(c.key), l);
  }

  const base = lock ?? asWorkbook;
  const lockKeys = [...Object.keys(base.names), ...Object.keys(base.cells ?? {})];
  const keys = new Set([...plan.sourceEdits.names, ...plan.sourceEdits.cells, ...plan.sourceEdits.spill].filter((k) => !differs || differs.has(k)));
  // A renamed or moved name is listed once, under its key in the source.
  const saidThere = new Set([...keys].filter((k) => where.has(k)).flatMap((k) => what.get(k) ?? []));
  for (const k of keys) {
    // A reader whose only edit is a renamed token: the rename's line says it.
    if (plan.renamedOnly.has(k)) continue;
    const w = where.get(k);
    const said = what.get(k) ?? [];
    if (!w && said.length > 0 && said.every((s) => saidThere.has(s))) continue;
    const e: UnbuiltEdit = { key: w?.key ?? lockKeys.find((x) => low(x) === k) ?? k, what: said.length ? said.join("; ") : w ? "edited in the source, not built" : "removed from the source, not built" };
    if (w) {
      e.file = w.file;
      e.line = w.line;
    }
    out.push(e);
  }
  // A library base (`@from`) the workbook's tag does not carry yet: a pull would write the
  // tag's back. Only where the workbook's name has a tag: without one (an older build, a
  // comment too long for it) the base cannot travel, and the edit is listed above if any.
  const listed = new Set(out.map((e) => low(e.key ?? "")));
  const tags = new Map<string, string | undefined>();
  for (const d of wb.definedNames) {
    if (d.isXlPrefixed || d.comment === undefined || d.comment === "") continue;
    const tag = splitProvenance(d.comment.split("\r\n").join("\n")).tag;
    if (tag) tags.set(low(scopedKey(d.name, d.scope.kind === "sheet" ? d.scope.name : null)), tag.lib);
  }
  for (const n of src.names) {
    const key = scopedKey(n.name, n.scope);
    if (listed.has(low(key)) || !tags.has(low(key)) || moduleOfPath(n.file) === undefined) continue;
    const lib = tags.get(low(key));
    if (lib === n.libBase) continue;
    const what = n.libBase === undefined ? `@from(lib #${lib}) removed in the source, not built` : `@from(lib #${n.libBase}) in the source, not built (the workbook's tag says ${lib === undefined ? "no library base" : `lib#${lib}`})`;
    out.push({ key, file: n.file, line: n.line, what });
  }
  out.sort((a, b) => (a.file ?? "￿").localeCompare(b.file ?? "￿") || (a.line ?? 0) - (b.line ?? 0));
  return out;
}

/**
 * Names files a pull would rewrite although they hold no edit of `unbuilt`: what differs is
 * layout or comments (`//` lines, blank lines, a definition spread over lines by hand), or
 * the file is one the pull does not write. A pull is fresh, so these go too; the guard
 * says so without refusing (M3e). `files`: the project as on disk (unsaved editors over
 * it); `pulled`: the files the pull would write. Line endings do not count (a pull writes
 * LF; a checkout on Windows may have CR LF). Paths in order.
 */
export function rewrittenFiles(files: Readonly<Record<string, string>>, pulled: Readonly<Record<string, string>>, unbuilt: readonly UnbuiltEdit[]): string[] {
  const edited = new Set(unbuilt.map((e) => e.file).filter((f) => f !== undefined));
  const lf = (s: string) => s.split("\r\n").join("\n");
  return Object.keys(files)
    .filter((p) => isSourcePath(p) && !edited.has(p) && (pulled[p] === undefined || lf(pulled[p]!) !== lf(files[p]!)))
    .sort();
}

/** One line per edit: `names/sheets/IS.xln:12  Revenue: update Revenue (comment)`. */
export function formatUnbuiltEdit(e: UnbuiltEdit): string {
  const at = e.file === undefined ? "" : `${e.file}${e.line !== undefined ? `:${e.line}` : ""}  `;
  return `${at}${e.key !== undefined ? `${e.key}: ` : ""}${e.what}`;
}
