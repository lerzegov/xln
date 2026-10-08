// `xln pull` without the file system: workbook bytes in, project files (path → text) out.
// The CLI writes them to disk; the extension writes them through vscode.workspace.fs.
//
// Every pull is fresh (decided 2026-10-06): the project is written from the workbook as
// it is, since there is no live link between Excel and the editor and the workbook is the
// only real state when one pulls. The structure is fixed (modules by prefix, sheet files
// by sheet). Pull never reads the source a build embedded (D5): that part is an archive
// copy for a workbook handed on. What a pull would replace, source edits not built yet,
// is the guard's to find first (build/unbuilt.ts).

import { readWorkbook } from "../file/workbook.js";
import type { DefinedName, ForeignModuleStore, WorkbookSnapshot } from "../file/types.js";
import { afeStatus, afeStoreLabel } from "../audit/afe.js";
import { prettyPrint } from "../lang/format.js";
import { decompileWithDiagnostics } from "../lang/transform.js";
import { classify, sheetCellsOf } from "./classify.js";
import { splitProvenance } from "./provenance.js";
import { provenanceStatus, type ProvenanceStatus, type ProvenanceTag } from "./provenance.js";
import { layoutToLf } from "./layout.js";
import { stringifyJson } from "./json.js";
import { buildLockfile, lockfileText } from "./lockfile.js";
import { buildManifest } from "./manifest.js";
import { formatCellAddress, formatEntry, formatSheetAnnotation, sheetOfPath } from "./module.js";
import { cellStatements, type CellStatement } from "./statements.js";
import { readCellValues, type CellValue } from "../build/verify.js";
import { valueLabels, type ValueLabel } from "./labels.js";
import { compareNames, fileSystemKey, moduleFileName, proposeModules, sheetFileName, SHEETS_DIR, UNMANAGED } from "./modules.js";
import { nameKey, type NameKind, type ProjectName } from "./types.js";

export { layoutToLf } from "./layout.js";

export const NAMES_DIR = "names";
export const MANIFEST_FILE = "workbook.manifest.json";
export const LOCK_FILE = "xln.lock.json";

export interface PullOptions {
  /** Definitions longer than this (in characters, on one line) are pretty-printed. Default 100. */
  width?: number;
}

export interface ModuleSummary {
  /** Module name; `_unmanaged` for the workbook-scoped names no module owns. */
  module: string;
  file: string;
  names: number;
  workbookScoped: number;
  sheetScoped: number;
}

/** Cell statements, counted (M3b). */
export interface CellCounts {
  /** Named formula cells (`Name @C6 = …;`). */
  named: number;
  /** Names on empty cells (`Name @C6 = ;`). */
  slots: number;
  /** Unnamed statements (`@C5 = …;`), blocks included. */
  unnamed: number;
  /** Unnamed statements over more than one cell (`@B40:G40 = …;`). */
  blocks: number;
  /** Cells those blocks cover. */
  blockCells: number;
}

/** A `names/sheets/<Sheet>.xln` file: one sheet's cell statements and the local names no module owns. */
export interface SheetFileSummary {
  sheet: string;
  file: string;
  /** Names written in the file: local names, named cells and slots. */
  names: number;
  cells: CellCounts;
}

export interface PullReport {
  workbook: string;
  /** Names written to modules (excludes built-ins and helpers). */
  names: number;
  byKind: Record<NameKind, number>;
  byScope: { workbook: number; sheet: number; perSheet: Record<string, number> };
  hidden: number;
  /** Module files, then `_unmanaged.xln` (when it has names). */
  modules: ModuleSummary[];
  /** Per-sheet files, in sheet order. */
  sheetFiles: SheetFileSummary[];
  /** Names in `_unmanaged.xln`. */
  unmanaged: number;
  /** Cell statements in all sheet files. */
  cells: CellCounts;
  /** `_xlnm.` names (print areas, filters): listed in the manifest, not in modules. */
  builtIns: string[];
  /** Other `_xl*` names Excel keeps for itself; dropped. */
  helpersIgnored: string[];
  /** Cell, conditional-format and validation formulas the manifest could not index. */
  unparsedFormulas: number;
  warnings: string[];
  /** Things a pull decided that the author may want to know (a module's `// @version` from its tags). */
  notes: string[];
  /**
   * Names Create from Selection seems to have taken from a computed cell's current value,
   * or from the corner of a two-way selection (one note each in `notes`).
   */
  valueLabels: Omit<ValueLabel, "note">[];
  /**
   * Other tools' copies of the names in the workbook (AFE's module store), as found; the
   * pull reads the names from the Name Manager, never from these.
   */
  foreignModules: ForeignModuleSummary[];
}

export interface ForeignModuleSummary {
  tool: "afe";
  kind: ForeignModuleStore["kind"];
  part: string;
  sheet?: string;
  /** Modules and how many names each defines, when readable. */
  modules?: { name: string; names: number }[];
  /** Names whose AFE text differs from the workbook's definition. */
  differs?: string[];
  /** Names AFE's modules define that the workbook does not have. */
  absent?: string[];
  unreadable?: string;
}

/** What a pull says about other tools' copies of the names: a summary and one note each. */
function foreignModules(wb: WorkbookSnapshot, notes: string[]): ForeignModuleSummary[] {
  const status = new Map(afeStatus(wb).map((s) => [s.store, s]));
  const out: ForeignModuleSummary[] = [];
  for (const store of wb.foreignModuleStores) {
    const sum: ForeignModuleSummary = { tool: store.tool, kind: store.kind, part: store.part };
    if (store.sheet !== undefined) sum.sheet = store.sheet;
    if (store.unreadable !== undefined) sum.unreadable = store.unreadable;
    const label = afeStoreLabel(store);
    const s = status.get(store);
    if (s) {
      sum.modules = s.modules;
      sum.differs = s.entries.filter((e) => e.state === "differs").map((e) => e.name);
      sum.absent = s.entries.filter((e) => e.state === "absent").map((e) => e.name);
      const parts = [`${s.entries.length} name${s.entries.length === 1 ? "" : "s"} in modules ${s.modules.map((m) => `${m.name} (${m.names})`).join(", ") || "none"}`];
      if (sum.differs.length) parts.push(`${sum.differs.length} differ from the workbook (${sum.differs.slice(0, 6).join(", ")}${sum.differs.length > 6 ? ", …" : ""})`);
      if (sum.absent.length) parts.push(`${sum.absent.length} not in the workbook`);
      if (s.notCompared) parts.push("not compared: other separators");
      notes.push(`the workbook carries Microsoft's Advanced Formula Environment (AFE) modules: ${label}, ${parts.join("; ")}. The project comes from the Name Manager; AFE's copy is left as it is (xln check, C14, lists the differences)`);
    } else if (store.kind === "locale-sheet") {
      notes.push(`${label}: AFE's scratch sheet, no names; left as it is`);
    } else {
      notes.push(`the workbook carries Microsoft's Advanced Formula Environment (AFE) modules: ${label}, which xln cannot read (${store.unreadable ?? "unknown format"}). The project comes from the Name Manager; AFE's copy is left as it is`);
    }
    out.push(sum);
  }
  return out;
}

export interface PullResult {
  /** Project-relative path (forward slashes) → file text (LF line endings). */
  files: Record<string, string>;
  report: PullReport;
  /** The names as classified, in project order. */
  names: ProjectName[];
  snapshot: WorkbookSnapshot;
  /** Every sheet's cell statements, sheets in order. */
  statements: CellStatement[];
  /** D6: names whose comment carries a provenance tag, and whether they still match it. */
  provenance: ProvenanceStatus[];
}

function baseName(path: string): string {
  const i = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
  return i < 0 ? path : path.slice(i + 1);
}

/**
 * The source text of a definition: as decompiled when it fits on the line after `name = `
 * or when the author laid it out on several lines; otherwise pretty-printed (A5).
 */
function sourceText(name: string, display: string, width: number): string {
  if (display.includes("\n") || name.length + 3 + display.length <= width) return display;
  return prettyPrint(display, { width });
}

export function pullProject(bytes: Uint8Array, fileName: string, opts: PullOptions = {}): PullResult {
  const width = opts.width ?? 100;
  const workbook = baseName(fileName);
  const wb = readWorkbook(bytes);
  const warnings = wb.warnings.map((w) => `file: ${w}`);

  const builtIns: DefinedName[] = [];
  const helpers: string[] = [];
  const regular: DefinedName[] = [];
  for (const d of wb.definedNames) {
    if (d.isBuiltIn) builtIns.push(d);
    else if (d.isXlPrefixed) helpers.push(d.name);
    else if (d.scopeInvalid) warnings.push(`${d.name}: localSheetId points at no sheet; not pulled`);
    else regular.push(d);
  }

  const allNames = regular.map((d) => d.name);
  const links = wb.externalLinks ?? [];
  const localBySheet = new Map<string, string[]>();
  for (const d of regular) {
    if (d.scope.kind !== "sheet") continue;
    let l = localBySheet.get(d.scope.name);
    if (!l) localBySheet.set(d.scope.name, (l = []));
    l.push(d.name);
  }
  const modules = proposeModules(allNames);

  // D6: each name's provenance tag, for the module's `// @version` (spec §14 issue 1).
  const tags = new Map<string, ProvenanceTag>();
  const names: ProjectName[] = regular.map((d) => {
    const scope = d.scope.kind === "sheet" ? d.scope.name : undefined;
    const classification = classify(d.definition);
    const label = scope === undefined ? d.name : `${scope}!${d.name}`;
    let display: string;
    if (classification.kind === "unparsed") {
      display = layoutToLf(d.definition);
      warnings.push(`${label}: does not parse (${classification.error}); written as stored`);
    } else {
      const ctx = scope === undefined ? { names: allNames, links } : { homeSheet: scope, localNames: localBySheet.get(scope) ?? [], names: allNames, links };
      const r = decompileWithDiagnostics(d.definition, ctx);
      display = sourceText(d.name, layoutToLf(r.text), width);
      for (const g of r.diagnostics) warnings.push(`${label}: ${g.message}`);
    }
    // D6: the provenance tag is the build's, not the author's: the source never shows it.
    const split = d.comment === undefined || d.comment === "" ? { comment: undefined } : splitProvenance(d.comment.split("\r\n").join("\n"));
    const comment = split.comment;
    // Its library base, though, is the author's (`@from(lib #…)`, written by Insert, Take
    // and Publish): the tag carried it through the workbook, and it comes back as written.
    const lib = split.tag?.lib;
    if (split.tag) tags.set(scope === undefined ? d.name : `${scope}!${d.name}`, split.tag);
    return {
      ...(lib !== undefined ? { libBase: lib } : {}),
      name: d.name,
      scope,
      scopePosition: d.scope.kind === "sheet" ? d.scope.position : undefined,
      hidden: d.hidden,
      comment,
      stored: d.definition,
      display,
      classification,
      module: modules.get(d.name),
    };
  });

  // Project order: workbook scope before sheets (in sheet order),
  // then by name. The same order is used in every file and in the manifest.
  const order = (a: ProjectName, b: ProjectName): number =>
    (a.scopePosition ?? -1) - (b.scopePosition ?? -1) || compareNames(a.name, b.name);
  names.sort(order);

  // Cell statements (M3b): every formula cell and every slot, by sheet. A name that is one
  // is written in its sheet's file, whatever its module.
  const values = new Map<string, Set<string>>();
  const cellValues = new Map<string, ReadonlyMap<string, CellValue>>();
  try {
    for (const s of readCellValues(bytes)) {
      values.set(s.sheet, new Set(s.cells.keys()));
      cellValues.set(s.sheet, s.cells);
    }
  } catch (e) {
    warnings.push(`cell values could not be read (${(e as Error).message}): names on cells holding a value may show as slots`);
  }
  const statements = cellStatements(wb, { values });
  const byKey = new Map(names.map((n) => [nameKey(n), n]));
  for (const list of statements.values()) {
    for (const s of list) {
      if (s.key === undefined) continue;
      const n = byKey.get(s.key);
      if (n) n.cell = { sheet: s.sheet, range: s.range };
    }
  }

  // Files: one per module (with its sheet-scoped members); `_unmanaged.xln` for the
  // workbook-scoped names no module owns; `sheets/<Sheet>.xln` for each sheet's cell
  // statements and its local names that no module owns, in sheet order. File names are
  // unique by `fileSystemKey`.
  const unmanagedPath = `${NAMES_DIR}/${moduleFileName(undefined)}`;
  const taken = new Set<string>([fileSystemKey(unmanagedPath)]);
  const unique = (path: string): string => {
    let p = path;
    for (let k = 2; taken.has(fileSystemKey(p)); k++) p = path.replace(/\.xln$/, `~${k}.xln`);
    taken.add(fileSystemKey(p));
    return p;
  };
  const plain = names.filter((n) => n.cell === undefined);
  // A workbook name no module owns whose definition is fixed to one sheet's cells (a cell
  // or range Create from Selection named, inside a spill or not) is written in that
  // sheet's file with `@workbook`, beside the sheet's own names (feedback 2026-10-07): the
  // author looks for it there, and removing `@workbook` makes it local. The rest of the
  // workbook names no module owns stay in `_unmanaged.xln`.
  const worksheets = wb.sheets.filter((s) => s.kind === "worksheet").map((s) => s.name);
  const onSheet = new Map<ProjectName, string>();
  for (const n of plain) {
    if (n.module !== undefined || n.scope !== undefined) continue;
    const sheet = sheetCellsOf(n.stored, worksheets);
    if (sheet !== undefined) onSheet.set(n, sheet);
  }
  /** The sheet file a name no module owns goes to; undefined for `_unmanaged.xln`. */
  const sheetOf = (n: ProjectName): string | undefined => n.scope ?? onSheet.get(n);
  const moduleList = [...new Set(plain.map((n) => n.module).filter((m): m is string => m !== undefined))].sort(compareNames);
  const moduleFile = new Map(moduleList.map((m) => [m, unique(`${NAMES_DIR}/${moduleFileName(m)}`)]));
  const sheetList = wb.sheets
    .filter((s) => (statements.get(s.name)?.length ?? 0) > 0 || plain.some((n) => n.module === undefined && sheetOf(n) === s.name))
    .map((s) => s.name);
  const sheetFile = new Map(sheetList.map((s) => [s, unique(`${NAMES_DIR}/${SHEETS_DIR}/${sheetFileName(s)}`)]));

  const files: Record<string, string> = {};
  const summaries: ModuleSummary[] = [];
  const sheetSummaries: SheetFileSummary[] = [];
  const fileByKey = new Map<string, string>();
  const emit = (path: string, members: readonly ProjectName[], head: string[]): void => {
    for (const n of members) fileByKey.set(nameKey(n), path);
    files[path] = renderFile(head, members);
  };
  const notes: string[] = [];
  const labels = valueLabels(wb, regular, cellValues);
  for (const l of labels) notes.push(l.note);
  for (const m of moduleList) {
    const members = plain.filter((n) => n.module === m);
    const path = moduleFile.get(m)!;
    const version = moduleVersionFromTags(m, members, tags, notes);
    emit(path, members, [`// module: ${m}, pulled by xln from ${workbook}.`, ...(version !== undefined ? [`// @version ${version}`] : [])]);
    summaries.push({
      module: m,
      file: path,
      names: members.length,
      workbookScoped: members.filter((n) => n.scope === undefined).length,
      sheetScoped: members.filter((n) => n.scope !== undefined).length,
    });
  }
  const unmanaged = plain.filter((n) => n.module === undefined && sheetOf(n) === undefined);
  if (unmanaged.length > 0) {
    emit(unmanagedPath, unmanaged, [
      "// Workbook-scoped names that no module owns and that are not tied to one sheet's cells,",
      `// pulled by xln from ${workbook}. A name on a sheet's cells is in that sheet's file`,
      "// (names/sheets/), a name with a module prefix (ANA.GROW) in its module's file.",
    ]);
    summaries.push({ module: UNMANAGED, file: unmanagedPath, names: unmanaged.length, workbookScoped: unmanaged.length, sheetScoped: 0 });
  }
  const totals: CellCounts = { named: 0, slots: 0, unnamed: 0, blocks: 0, blockCells: 0 };
  for (const s of sheetList) {
    const stmts = statements.get(s) ?? [];
    // The sheet's own names first, then the workbook names on its cells, each by name.
    const others = [...plain.filter((n) => n.module === undefined && n.scope === s), ...plain.filter((n) => n.module === undefined && n.scope === undefined && onSheet.get(n) === s)];
    const path = sheetFile.get(s)!;
    for (const n of others) fileByKey.set(nameKey(n), path);
    for (const st of stmts) if (st.key !== undefined) fileByKey.set(st.key, path);
    files[path] = renderSheetFile(s, workbook, stmts, others, byKey, width, path);
    const counts = countStatements(stmts);
    for (const k of Object.keys(totals) as (keyof CellCounts)[]) totals[k] += counts[k];
    sheetSummaries.push({ sheet: s, file: path, names: others.length + counts.named + counts.slots, cells: counts });
  }

  const allStatements = [...statements.values()].flat();
  const { manifest, unparsedFormulas } = buildManifest({ fileName: workbook, wb, names, files: fileByKey, builtIns, warnings, statements: allStatements });
  files[MANIFEST_FILE] = stringifyJson(manifest) + "\n";

  files[LOCK_FILE] = lockfileText(buildLockfile(workbook, names, allStatements));

  const byKind: Record<NameKind, number> = { constant: 0, range: 0, spill: 0, table: 0, formula: 0, lambda: 0, unparsed: 0 };
  const perSheet: Record<string, number> = {};
  for (const n of names) {
    byKind[n.classification.kind]++;
    if (n.scope !== undefined) perSheet[n.scope] = (perSheet[n.scope] ?? 0) + 1;
  }
  const report: PullReport = {
    workbook,
    names: names.length,
    byKind,
    byScope: { workbook: names.filter((n) => n.scope === undefined).length, sheet: names.filter((n) => n.scope !== undefined).length, perSheet },
    hidden: names.filter((n) => n.hidden).length,
    modules: summaries,
    sheetFiles: sheetSummaries,
    unmanaged: unmanaged.length,
    cells: totals,
    builtIns: builtIns.map((b) => (b.scope.kind === "sheet" ? `${b.scope.name}!${b.name}` : b.name)),
    helpersIgnored: helpers,
    unparsedFormulas,
    warnings,
    notes,
    valueLabels: labels.map(({ key, kind, cell, range }) => ({ key, kind, cell, range })),
    foreignModules: foreignModules(wb, notes),
  };
  return { files, report, names, snapshot: wb, statements: allStatements, provenance: provenanceStatus(wb) };
}

/**
 * The `// @version` a pull writes back into a module file (spec §14 issue 1): the version
 * its names' provenance tags carry, when all of them, or a clear majority (two in three),
 * carry the same one. The build put it there from the file's header, so the next build
 * plans no provenance update. Tags of another module, and names without a tag, do not
 * count. Disagreements are said in `notes`.
 */
function moduleVersionFromTags(module: string, members: readonly ProjectName[], tags: ReadonlyMap<string, ProvenanceTag>, notes: string[]): string | undefined {
  const byVersion = new Map<string, string[]>();
  for (const n of members) {
    const key = nameKey(n);
    const t = tags.get(key);
    if (!t || t.module.toLowerCase() !== module.toLowerCase()) continue;
    const v = t.version ?? "";
    byVersion.set(v, [...(byVersion.get(v) ?? []), key]);
  }
  const total = [...byVersion.values()].reduce((a, l) => a + l.length, 0);
  if (total === 0) return undefined;
  const ranked = [...byVersion].sort((a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0]));
  const [top, keys] = ranked[0]!;
  const others = ranked.slice(1);
  const say = (v: string) => (v === "" ? "no version" : v);
  const list = (l: readonly string[]) => (l.length > 5 ? `${l.slice(0, 5).join(", ")}, …` : l.join(", "));
  if (others.length === 0) return top === "" ? undefined : top;
  const detail = others.map(([v, l]) => `${list(l)} ${say(v)}`).join("; ");
  if (top !== "" && keys.length * 3 >= total * 2) {
    notes.push(`module ${module}: wrote // @version ${top}, which ${keys.length} of its ${total} tagged names carry; ${detail} (the next build tags them ${top})`);
    return top;
  }
  notes.push(`module ${module}: no // @version written: its tagged names disagree (${ranked.map(([v, l]) => `${say(v)}: ${l.length}`).join(", ")}; ${detail}); write the version in the file's header if it has one`);
  return undefined;
}

function countStatements(stmts: readonly CellStatement[]): CellCounts {
  const c: CellCounts = { named: 0, slots: 0, unnamed: 0, blocks: 0, blockCells: 0 };
  for (const s of stmts) {
    if (s.kind === "named") c.named++;
    else if (s.kind === "slot") c.slots++;
    else {
      c.unnamed++;
      if (s.rows * s.cols > 1) {
        c.blocks++;
        c.blockCells += s.rows * s.cols;
      }
    }
  }
  return c;
}

/**
 * One module file (or `_unmanaged.xln`). Members come in project order: workbook-scoped
 * names first, then each sheet's, each of those with `@sheet(Sheet)` above it (per name,
 * no blocks, 2026-10-07).
 */
function renderFile(head: readonly string[], members: readonly ProjectName[]): string {
  const out: string[] = [...head, ""];
  let prevBlock = false;
  for (const n of members) {
    const text = formatEntry({ name: n.name, scope: n.scope, hidden: n.hidden, from: n.libBase, doc: n.comment, formula: n.display });
    const block = n.scope !== undefined || n.comment !== undefined || n.libBase !== undefined || n.display.includes("\n");
    if ((block || prevBlock) && out[out.length - 1] !== "") out.push("");
    out.push(text);
    prevBlock = block;
  }
  return out.join("\n") + "\n";
}

/**
 * A sheet's file (M3d): its cell statements in sheet order, then its other local names.
 * Every name is local to the sheet unless `@workbook` is on the line above it (a
 * workbook-scoped name on one of the sheet's cells, what *Create from Selection* makes).
 * Addresses are bare: the sheet is the file's. No `@scope`/`@workbook` blocks.
 */
function renderSheetFile(
  sheet: string,
  workbook: string,
  stmts: readonly CellStatement[],
  others: readonly ProjectName[],
  byKey: ReadonlyMap<string, ProjectName>,
  width: number,
  path = "",
): string {
  const out: string[] = [
    `// Sheet ${sheet}, pulled by xln from ${workbook}: its formula cells in sheet order, then`,
    "// the other names on its cells or local to it that no module owns. Every name here is",
    `// local to ${sheet} unless the line above it says @workbook. A cell's address (@C6) is`,
    "// set in Excel and read-only; edit the formula after '='. `Name @C6 = ;` is a slot: a",
    "// name on an empty cell. `Name @C6#` names the cell's spill, `Name @C6` the cell alone:",
    "// add or remove the `#`.",
    "",
  ];
  // A file name that cannot say its sheet (a `~2` a clash on the file system added) names it.
  if (sheetOfPath(path) !== sheet) out.push(formatSheetAnnotation(sheet), "");
  let prevBlock = false;
  const put = (text: string, isBlock: boolean): void => {
    if ((isBlock || prevBlock) && out[out.length - 1] !== "") out.push("");
    out.push(text);
    prevBlock = isBlock;
  };
  for (const s of stmts) {
    const n = s.key !== undefined ? byKey.get(s.key) : undefined;
    const global = s.kind !== "unnamed" && s.scope === undefined;
    // `#` exactly when the name is on the spill (`'S'!$C$6#`): the address says what the name covers.
    const address = formatCellAddress(s.range, undefined, s.onSpill === true);
    const label = s.name === undefined ? `@${address}` : `${s.name} @${address}`;
    const formula = s.error !== undefined || s.display === "" ? s.display : sourceText(label, s.display, width);
    const text = formatEntry({ name: s.name ?? "", hidden: n?.hidden ?? false, from: n?.libBase, doc: n?.comment, formula, cell: address, workbook: global });
    const isBlock = n?.comment !== undefined || n?.libBase !== undefined || formula.includes("\n");
    put(text, isBlock);
  }
  if (others.length > 0) {
    if (stmts.length > 0) {
      if (out[out.length - 1] !== "") out.push("");
      out.push(`// Other names on ${sheet}.`, "");
      prevBlock = false;
    }
    for (const n of others) {
      const text = formatEntry({ name: n.name, hidden: n.hidden, from: n.libBase, doc: n.comment, formula: n.display, workbook: n.scope === undefined });
      const isBlock = n.comment !== undefined || n.libBase !== undefined || n.display.includes("\n");
      put(text, isBlock);
    }
  }
  return out.join("\n") + "\n";
}
