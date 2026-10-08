// The library (M4) in the editor, without the vscode API: where a project's library is,
// completion items for library functions the project lacks, and each module entry's
// state against the library (for the code lenses and quick fixes). The library itself is
// read by library.ts through vscode.workspace.fs; here it is a core `Library`.

import {
  docTagWarnings,
  libraryClosure,
  libStateLabel,
  parseLibBase,
  FROM,
  projectLibraryStatus as coreProjectStatus,
  moduleOfPath,
  type Library,
  type LibraryFunction,
  type LibStatusItem,
  type LibStatusReport,
  type LibraryBase,
  type WorkbookSnapshot,
} from "@xln/core";
import type { Project } from "./project.js";

const lower = (s: string) => s.toLowerCase();

/** The project's files, path → live text (unsaved edits included). */
export function projectFiles(project: Project): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [p, f] of project.files) out[p] = f.text;
  return out;
}

/**
 * Insert, Take, Publish and Record library base write `@from`, which the build carries in
 * the provenance tag; a doc comment too long for the tag gets none. The warning to show
 * before (core `docTagWarnings`), on the project's files with `changed` (path → text after
 * the edit): one text, undefined when every doc comment of `names` leaves room.
 */
export function tagWarning(files: Readonly<Record<string, string>>, changed: Readonly<Record<string, string>>, names: readonly string[]): string | undefined {
  const w = docTagWarnings({ ...files, ...changed }, names);
  return w.length === 0 ? undefined : w.map((x) => x.message).join(" ");
}

/** Whether the project defines a workbook-scoped name (any file). */
export function definesWorkbookName(project: Project, name: string): boolean {
  return project.scopeIndex().workbook.some((d) => lower(d.name) === lower(name));
}

/**
 * Where the library is, from `xln.config.json`'s `library`: a path relative to the project
 * folder, absolute, or under `~` (needs `home`). Paths use `/`; a Windows drive path
 * (`C:\lib`, `C:/lib`) counts as absolute.
 */
export function libraryLocation(setting: string, home: string | undefined): { relative: string } | { absolute: string } | { error: string } {
  let p = setting.trim().split("\\").join("/");
  if (p === "~" || p.startsWith("~/")) {
    if (home === undefined) return { error: `"library": "${setting}" starts with ~, which cannot be expanded here (vscode.dev): use a path relative to the project folder` };
    p = home.split("\\").join("/").replace(/\/$/, "") + p.slice(1);
  }
  if (p.startsWith("/")) return { absolute: p };
  if (p.length >= 2 && p[1] === ":" && /[A-Za-z]/.test(p[0]!)) return { absolute: "/" + p };
  return { relative: p };
}

/** `FN.GROW(seed, growth, periods)`. */
export function librarySignature(fn: LibraryFunction): string {
  return `${fn.name}(${fn.params.join(", ")})`;
}

export interface LibraryCompletion {
  fn: LibraryFunction;
  /** What accepting it adds to the project: the function and the library functions it calls that the project lacks. */
  adds: string[];
}

/**
 * The library functions to offer while `word` is typed in a formula of `project`: those
 * whose name starts with the word (case ignored) and the project does not define.
 */
export function libraryCompletions(project: Project, library: Library, word: string): LibraryCompletion[] {
  const w = lower(word);
  const out: LibraryCompletion[] = [];
  for (const fn of library.functions) {
    if (!lower(fn.name).startsWith(w) || definesWorkbookName(project, fn.name)) continue;
    out.push({ fn, adds: libraryClosure(library, fn.name, (n) => definesWorkbookName(project, n)).map((f) => f.name) });
  }
  return out;
}

/** The documentation of a library item (Markdown). */
export function libraryItemDoc(c: LibraryCompletion, file: string): string {
  const parts = [`*From the library* (\`${c.fn.path}\`)`];
  if (c.fn.summary) parts.push(c.fn.summary);
  parts.push("```xln\n" + c.fn.definition + "\n```");
  const others = c.adds.filter((n) => lower(n) !== lower(c.fn.name));
  parts.push(`Accepting it adds the definition${others.length ? ` (and ${others.join(", ")}, which it calls)` : ""} to \`${file}\`; xln build writes it into the workbook.`);
  return parts.join("\n\n");
}

/** Prefixes (`FN.`) of library functions the project lacks, for a module item when the project has none of that module. */
export function libraryModules(project: Project, library: Library): string[] {
  const out = new Set<string>();
  for (const fn of library.functions) {
    const dot = fn.name.indexOf(".");
    if (dot > 0 && !definesWorkbookName(project, fn.name)) out.add(fn.name.slice(0, dot + 1));
  }
  return [...out];
}

/**
 * The library status of a project (its live source), with the workbook's tags and what it
 * lacks of the source when it is given; `bases`: the project's kept `library-bases/`, for
 * the three-way diff.
 */
export function projectLibraryStatus(
  project: Project,
  library: Library,
  wb: WorkbookSnapshot | undefined,
  meta: { target: string; library: string; workbook?: string; bases?: ReadonlyMap<string, LibraryBase> },
): LibStatusReport {
  return coreProjectStatus(library, projectFiles(project), wb, meta);
}

/** The status items of a file's entries, by lower-case name: library functions, and LAMBDAs of a module file. */
export function entryStates(report: LibStatusReport, path: string): Map<string, LibStatusItem> {
  const out = new Map<string, LibStatusItem>();
  for (const i of report.items) if (i.file === path) out.set(lower(i.name), i);
  return out;
}

/**
 * What the editor offers on a module entry, by its library state (three-way on its base,
 * 2026-10-07). Take on a copy with an edit of its own asks first, naming the edit it
 * discards; Publish is not offered where it would overwrite a library change the copy
 * does not have (outdated, both changed): the publish command still asks before writing.
 */
export interface EntryLibraryActions {
  /** The code lens's label and tooltip. */
  label: string;
  tooltip: string;
  /** Clicking the label opens the diff (two-way, or three-way for both changed). */
  diff: boolean;
  /** "Take the library's version"; `confirm`: the question asked first, when the copy has (or may have) an edit of its own. */
  take?: { title: string; confirm?: string };
  publish?: { title: string };
  /** "Record library base": an identical copy that records none (author's idea, 2026-10-07). */
  record?: { title: string };
}

export function entryLibraryActions(item: LibStatusItem): EntryLibraryActions {
  const name = item.name;
  const state = item.state;
  const label = (state === "local-only" ? "library: not in the library" : `library: ${libStateLabel(state)}`) + (item.unbuilt ? " · not built yet" : "");
  const base = item.base !== undefined ? ` (base #${item.base}, library #${item.libraryHash})` : "";
  const out: EntryLibraryActions = { label, tooltip: item.note + base, diff: state === "outdated" || state === "modified" || state === "differs" || state === "both-changed" };
  const take = "Take the library's version";
  switch (state) {
    case "outdated":
      out.take = { title: take };
      break;
    case "modified":
      out.take = { title: `${take} (undo the edit)`, confirm: `Discard your edit of ${name}? Its definition and doc comment become the library's version #${item.libraryHash}.` };
      out.publish = { title: "Publish to library" };
      break;
    case "both-changed":
      out.take = { title: take, confirm: `Discard your edit of ${name}? The library changed it too (#${item.base} → #${item.libraryHash}); taking the library's version loses the edit made here. Show the diff first to compare the three.` };
      break;
    case "differs":
      // No base: a local edit cannot be ruled out, so Take asks too (coordinator, 2026-10-07).
      out.take = { title: take, confirm: `${name} has no library base: taking the library's version replaces whatever this copy has. Continue?` };
      out.publish = { title: "Publish to library" };
      break;
    case "local-only":
      out.publish = { title: "Publish to library" };
      break;
    case "identical":
      // Identical now, "differs" after the first change on either side: offered, never done.
      if (item.noBase) {
        out.label += " · no base";
        out.record = { title: "Record library base" };
      }
      break;
  }
  return out;
}

/** The question the Publish command adds before overwriting a library change the copy does not have; undefined when there is none. */
export function publishWarning(item: LibStatusItem | undefined): string | undefined {
  if (item?.state === "outdated") return `The library has a newer version of ${item.name} (#${item.libraryHash}) than this copy's base (#${item.base}): publishing puts the older definition back.`;
  if (item?.state === "both-changed") return `The library changed ${item.name} too since this copy's base (#${item.base} → #${item.libraryHash}): publishing replaces the library's change with this copy.`;
  return undefined;
}

/** The hover of an entry's `@from(lib #…)` at `offset` in a file (Markdown); undefined elsewhere. */
export function libraryBaseHover(project: Project, path: string, offset: number): string | undefined {
  const file = project.files.get(path);
  if (!file) return undefined;
  for (const e of file.parsed.entries) {
    const a = e.annotations.find((x) => x.name === FROM && x.offset <= offset && offset <= x.end);
    if (!a) continue;
    const b = parseLibBase(a.arg);
    if ("error" in b) return `**@from**: ${b.error}`;
    return [
      `**@from(lib #${b.hash})**: from the library version #${b.hash}, the library definition ${e.name} came from.`,
      "Written by *Insert library function*, *Take the library's version*, *Publish to library* and *Record library base*, which also keep the base's text in the project's `library-bases/` for the three-way diff; nothing else changes it (delete it and the copy has no base). " +
        `The build carries it in the Name Manager comment's tag (\`lib#${b.hash}\`) and pull writes it back.`,
      "*Library status* compares the copy, this base and the library: unchanged since the base and the library moved is **outdated**; edited here with the library still at the base is **modified**; both moved is **both changed**.",
    ].join("\n\n");
  }
  return undefined;
}

/** Whether a file is a module file, whose LAMBDAs may be published. */
export function isModuleFile(path: string): boolean {
  return moduleOfPath(path) !== undefined;
}

/** The entry a library command was given, if it was given one: from a menu VS Code passes the document's Uri instead. */
export function entryArg<T extends { root: string; name: string }>(arg: unknown): T | undefined {
  return arg !== null && typeof arg === "object" && typeof (arg as { root?: unknown }).root === "string" && typeof (arg as { name?: unknown }).name === "string" ? (arg as T) : undefined;
}
