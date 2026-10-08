// Shared check for the probe and corpus tests: parse every written module back and
// compare with the workbook's names and formula cells.
import {
  compile,
  decompileWithDiagnostics,
  definitionTarget,
  equalModuloWhitespace,
  nameKey,
  parseSourceFile,
  parseCellRange,
  sameFilled,
  sheetCellsOf,
  sheetFormulaCells,
  type ModuleEntry,
  type PullResult,
} from "../../src/index.js";

export interface ProjectCheck {
  /** Problems; empty when the project is complete and loss-free. */
  failures: string[];
  /** Definitions stored without a needed prefix: compile repairs them, so they differ on purpose. */
  repaired: string[];
  entries: number;
  /** Cell statements written, and the formula cells they cover. */
  cellStatements: number;
  cellsCovered: number;
}

export function checkProject(r: PullResult): ProjectCheck {
  const failures: string[] = [];
  const repaired: string[] = [];
  const found = new Map<string, { entry: ModuleEntry; file: string }[]>();
  const cellEntries: { entry: ModuleEntry; file: string; sheet: string }[] = [];
  let entries = 0;
  const worksheets = r.snapshot.sheets.filter((s) => s.kind === "worksheet").map((s) => s.name);
  for (const [path, text] of Object.entries(r.files)) {
    if (!path.endsWith(".xln")) continue;
    const parsed = parseSourceFile(path, text);
    for (const d of parsed.diagnostics) failures.push(`${path}:${d.line}:${d.col}: ${d.message}`);
    if (path.startsWith("names/sheets/")) {
      // One sheet per file (M3d): no blocks, bare addresses; a workbook name on one of its
      // cells has @workbook above it, every other name is local to the file's sheet.
      if (parsed.form !== "sheet") failures.push(`${path}: read as ${parsed.form}, not as a sheet file`);
      if (parsed.scopes.length) failures.push(`${path}: has @scope/@workbook blocks`);
      const sheet = parsed.sheet;
      for (const e of parsed.entries) {
        // A workbook name on the sheet's cells: a cell statement, or a definition fixed to its cells (2026-10-07).
        const onSheet = e.cell !== undefined || (sheet !== undefined && sheetCellsOf(e.formula, worksheets)?.toLowerCase() === sheet.toLowerCase());
        if (e.scope !== sheet && !(e.scope === undefined && e.workbook && onSheet)) failures.push(`${path}: ${e.name} has scope ${e.scope} in the file of ${sheet}`);
        if (e.cell?.sheet !== undefined) failures.push(`${path}: ${e.name}'s address names its sheet`);
      }
      for (const e of parsed.cells) cellEntries.push({ entry: e, file: path, sheet: e.cellSheet! });
    } else {
      if (path === "names/_unmanaged.xln" && parsed.entries.some((e) => e.scope !== undefined)) failures.push(`${path}: holds sheet-scoped names`);
      if (parsed.cells.length) failures.push(`${path}: holds cell statements`);
    }
    for (const e of parsed.entries) {
      entries++;
      const key = nameKey({ name: e.name, scope: e.scope });
      const list = found.get(key) ?? [];
      list.push({ entry: e, file: path });
      found.set(key, list);
    }
  }

  const allNames = r.snapshot.definedNames.map((d) => d.name);
  const regular = r.snapshot.definedNames.filter((d) => !d.isXlPrefixed && !d.scopeInvalid);
  const keys = new Set<string>();
  for (const d of regular) {
    const scope = d.scope.kind === "sheet" ? d.scope.name : undefined;
    const key = nameKey({ name: d.name, scope });
    keys.add(key);
    const hits = found.get(key) ?? [];
    if (hits.length !== 1) {
      failures.push(`${key}: written ${hits.length} times (${hits.map((h) => h.file).join(", ")})`);
      continue;
    }
    const e = hits[0]!.entry;
    const comment = d.comment ? d.comment.split("\r\n").join("\n") : undefined;
    if (e.doc !== comment) failures.push(`${key}: comment ${JSON.stringify(e.doc)} ≠ ${JSON.stringify(comment)}`);
    if (e.hidden !== d.hidden) failures.push(`${key}: hidden ${e.hidden} ≠ ${d.hidden}`);
    if (e.cell) {
      // A named cell: the definition is the cell, the statement holds the cell's formula.
      const t = definitionTarget(d.definition);
      const at = parseCellRange(e.cell.range);
      if (!t || !at || t.r1 !== at.r1 || t.c1 !== at.c1 || t.r1 !== t.r2 || t.c1 !== t.c2) failures.push(`${key}: defined as ${d.definition}, written on @${e.cell.range}`);
      continue;
    }
    const local = regular.filter((x) => x.scope.kind === "sheet" && x.scope.name === scope).map((x) => x.name);
    const links = r.snapshot.externalLinks;
    const ctx = scope === undefined ? { names: allNames, links } : { homeSheet: scope, localNames: local, names: allNames, links };
    try {
      const { diagnostics } = decompileWithDiagnostics(d.definition, ctx);
      const back = compile(e.formula, ctx);
      if (diagnostics.some((g) => g.code === "bare-prefix")) repaired.push(`${key}: ${d.definition} → ${back}`);
      else if (!equalModuloWhitespace(back, d.definition)) {
        failures.push(`${key}: compiled source differs\n  stored:   ${d.definition}\n  source:   ${e.formula}\n  compiled: ${back}`);
      }
    } catch (err) {
      const p = r.names.find((n) => nameKey(n) === key);
      if (p?.classification.kind === "unparsed" && e.formula === p.display) continue;
      failures.push(`${key}: ${err instanceof Error ? err.message : String(err)}`);
    }
  }
  for (const key of found.keys()) if (!keys.has(key)) failures.push(`${key}: written but not in the workbook`);

  // Cells: every formula cell is covered by exactly one statement, whose formula compiles
  // to the cell's own (shifted for the cells of a block); a slot covers an empty cell.
  let cellsCovered = 0;
  const covered = new Map<string, string>();
  for (const { entry: e, file, sheet: sheetName } of cellEntries) {
    const label = `${file}:${e.line} @${e.cell!.range}`;
    const sheet = r.snapshot.sheets.find((s) => s.name.toLowerCase() === sheetName.toLowerCase());
    const rect = parseCellRange(e.cell!.range);
    if (!sheet || !rect) {
      failures.push(`${label}: no sheet ${sheetName} or bad range`);
      continue;
    }
    const fc = sheetFormulaCells(sheet);
    let compiled: string | undefined;
    if (e.formula !== "") {
      try {
        compiled = compile(e.formula, { names: allNames, links: r.snapshot.externalLinks, allowUnknownFunctions: true });
      } catch {
        compiled = e.formula;
      }
    }
    const top = fc.cells.get(`${colName(rect.c1)}${rect.r1}`);
    for (let row = rect.r1; row <= rect.r2; row++) {
      for (let col = rect.c1; col <= rect.c2; col++) {
        const addr = `${sheet.name}!${colName(col)}${row}`;
        if (covered.has(addr)) failures.push(`${label}: ${addr} is also covered by ${covered.get(addr)}`);
        covered.set(addr, label);
        cellsCovered++;
        const f = fc.cells.get(`${colName(col)}${row}`);
        if (e.formula === "") {
          if (f?.stored) failures.push(`${label}: a slot on a formula cell`);
          continue;
        }
        if (!f?.stored || !top?.stored) {
          failures.push(`${label}: ${addr} has no formula`);
          continue;
        }
        if (!sameFilled(top.stored, row - rect.r1, col - rect.c1, f.stored)) failures.push(`${label}: ${addr} is not the block's formula filled there`);
      }
    }
    if (compiled !== undefined && top?.stored !== undefined) {
      const bare = decompileWithDiagnostics(top.stored).diagnostics.some((g) => g.code === "bare-prefix");
      if (!bare && !equalModuloWhitespace(compiled, top.stored) && !equalModuloWhitespace(e.formula, top.stored)) {
        failures.push(`${label}: compiled source differs\n  stored:   ${top.stored}\n  source:   ${e.formula}\n  compiled: ${compiled}`);
      }
    }
  }
  for (const sheet of r.snapshot.sheets) {
    for (const [addr, f] of sheetFormulaCells(sheet).cells) {
      if (f.kind === "data-table" || !f.stored || f.stored.trim() === "") continue;
      if (!covered.has(`${sheet.name}!${addr}`)) failures.push(`${sheet.name}!${addr}: formula cell without a statement`);
    }
  }
  return { failures, repaired, entries, cellStatements: cellEntries.length, cellsCovered };
}

function colName(col: number): string {
  let s = "";
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
