// A6: workbook.manifest.json, the read-only context next to the modules: sheets, Tables,
// the spill map, and where each name is used (in other names, cells, conditional
// formats, data validations, Table column formulas). For the cell statements of M3b it
// records where each named cell is (`cell`) and the unnamed statements' ranges
// (`unnamedCells`), which the editor checks the read-only addresses against.

import type { DefinedName, WorkbookSnapshot } from "../file/types.js";
import type { Expr } from "../lang/ast.js";
import { tryParse } from "../lang/parser.js";
import { compressCells, extentSize } from "./cells.js";
import type { Json } from "./json.js";
import { NameResolver, nameUses } from "./refs.js";
import type { CellStatement } from "./statements.js";
import { nameKey, type ProjectName } from "./types.js";

export const MANIFEST_FORMAT = "xln.manifest/1";

interface Usage {
  names: Set<string>;
  cells: Map<string, Set<string>>;
  conditionalFormats: Map<string, Set<string>>;
  dataValidations: Map<string, Set<string>>;
  tableColumns: Set<string>;
}

function emptyUsage(): Usage {
  return { names: new Set(), cells: new Map(), conditionalFormats: new Map(), dataValidations: new Map(), tableColumns: new Set() };
}

function add(m: Map<string, Set<string>>, sheet: string, item: string): void {
  let s = m.get(sheet);
  if (!s) m.set(sheet, (s = new Set()));
  s.add(item);
}

function scopeText(n: DefinedName): string | null {
  return n.scope.kind === "sheet" ? n.scope.name : null;
}

export interface ManifestInput {
  fileName: string;
  wb: WorkbookSnapshot;
  /** In project order. */
  names: readonly ProjectName[];
  /** Module file of each name key. */
  files: ReadonlyMap<string, string>;
  builtIns: readonly DefinedName[];
  warnings: readonly string[];
  /** The cell statements written in the sheet files (M3b). */
  statements?: readonly CellStatement[];
}

export interface ManifestResult {
  manifest: Json;
  /** Formulas the manifest could not index (they do not parse). */
  unparsedFormulas: number;
}

/**
 * The unnamed cell statements of each sheet, as their ranges in sheet order: what the
 * editor checks an `@C5` address against (named cells carry `cell` in their name entry).
 */
function unnamedCells(statements: readonly CellStatement[]): Json {
  const out: Record<string, Json> = {};
  for (const s of statements) {
    if (s.kind !== "unnamed") continue;
    ((out[s.sheet] ??= []) as Json[]).push(s.range);
  }
  return out;
}

export function buildManifest(input: ManifestInput): ManifestResult {
  const { wb, names } = input;
  const keys = names.map((n) => ({ name: n.name, scope: n.scope, key: nameKey(n) }));
  const resolver = new NameResolver(keys);
  const usage = new Map<string, Usage>(keys.map((k) => [k.key, emptyUsage()]));
  const uses = new Map<string, string[]>();
  let unparsedFormulas = 0;

  // One parse per distinct formula text: shared formulas and copied blocks repeat a lot.
  const parsed = new Map<string, Expr | null>();
  const bodyOf = (text: string): Expr | undefined => {
    let b = parsed.get(text);
    if (b === undefined) {
      b = tryParse(text).formula?.body ?? null;
      if (b === null) unparsedFormulas++;
      parsed.set(text, b);
    }
    return b ?? undefined;
  };
  const usedNames = (text: string, home: string | undefined): string[] => {
    const body = bodyOf(text);
    if (!body) return [];
    const out: string[] = [];
    for (const u of nameUses(body)) {
      const k = resolver.resolve(u, home);
      if (k !== undefined && !out.includes(k)) out.push(k);
    }
    return out;
  };

  for (const n of names) {
    const k = nameKey(n);
    const us = usedNames(n.stored, n.scope);
    uses.set(k, us);
    for (const u of us) usage.get(u)!.names.add(k);
  }

  for (const sheet of wb.sheets) {
    const masters = new Map<number, string | undefined>();
    for (const g of sheet.sharedFormulas) masters.set(g.si, g.text);
    for (const f of sheet.formulas) {
      if (f.kind === "data-table") continue;
      // Names do not move when a shared formula is filled, so the master's text serves.
      const text = f.text ?? (f.si !== undefined ? masters.get(f.si) : undefined);
      if (text === undefined || text.trim() === "") continue; // Excel saves some empty <f/>
      for (const k of usedNames(text, sheet.name)) add(usage.get(k)!.cells, sheet.name, f.cell);
    }
    for (const cf of sheet.conditionalFormats) {
      for (const text of cf.formulas) for (const k of usedNames(text, sheet.name)) add(usage.get(k)!.conditionalFormats, sheet.name, cf.sqref);
    }
    for (const dv of sheet.dataValidations) {
      for (const text of [dv.formula1, dv.formula2]) {
        if (text === undefined || text.trim() === "") continue; // Excel saves some empty <f/>
        for (const k of usedNames(text, sheet.name)) add(usage.get(k)!.dataValidations, sheet.name, dv.sqref);
      }
    }
  }
  for (const t of wb.tables) {
    for (const c of t.columns) {
      for (const text of [c.calculatedColumnFormula, c.totalsRowFormula]) {
        if (text === undefined || text.trim() === "") continue; // Excel saves some empty <f/>
        for (const k of usedNames(text, t.sheet.name)) usage.get(k)!.tableColumns.add(`${t.displayName}[${c.name}]`);
      }
    }
  }

  const sheetOrder = new Map(wb.sheets.map((s) => [s.name, s.position]));
  const bySheet = (m: Map<string, Set<string>>, compress: boolean): Json | undefined => {
    if (m.size === 0) return undefined;
    const o: Record<string, Json> = {};
    for (const s of [...m.keys()].sort((a, b) => (sheetOrder.get(a) ?? 0) - (sheetOrder.get(b) ?? 0))) {
      o[s] = compress ? compressCells(m.get(s)!) : [...m.get(s)!];
    }
    return o;
  };

  const spillAt = (sheet: string | undefined, cell: string) =>
    wb.sheets.find((s) => s.name.toLowerCase() === (sheet ?? "").toLowerCase())?.spills.find((p) => p.anchor === cell);

  const nameEntries: Record<string, Json> = {};
  for (const n of names) {
    const k = nameKey(n);
    const c = n.classification;
    const u = usage.get(k)!;
    const usedBy: Record<string, Json | undefined> = {
      names: u.names.size ? [...u.names] : undefined,
      cells: bySheet(u.cells, true),
      conditionalFormats: bySheet(u.conditionalFormats, false),
      dataValidations: bySheet(u.dataValidations, false),
      tableColumns: u.tableColumns.size ? [...u.tableColumns] : undefined,
    };
    let spill: Json | undefined;
    if (c.anchor) {
      const sheet = c.anchor.sheet ?? n.scope;
      const sp = spillAt(sheet, c.anchor.cell);
      const size = sp ? extentSize(sp.extent) : undefined;
      spill = {
        sheet: sheet ?? null,
        anchor: c.anchor.cell,
        extent: sp?.extent ?? null,
        spilling: size ? size.rows * size.cols > 1 : null,
      };
    }
    nameEntries[k] = {
      name: n.name,
      scope: n.scope ?? null,
      kind: c.kind,
      arity: c.arity ? { required: c.arity.required, optional: c.arity.optional } : undefined,
      params: c.params,
      table: c.table,
      spill,
      error: c.error,
      hidden: n.hidden || undefined,
      module: n.module ?? null,
      cell: n.cell ? { sheet: n.cell.sheet, range: n.cell.range } : undefined,
      file: input.files.get(k)!,
      uses: uses.get(k)!.length ? uses.get(k)! : undefined,
      usedBy: Object.values(usedBy).some((v) => v !== undefined) ? (usedBy as Json) : undefined,
    };
  }

  const spills: Record<string, Json> = {};
  const dynamicArrayCells: Record<string, Json> = {};
  for (const s of wb.sheets) {
    const real: Json[] = [];
    const single: string[] = [];
    for (const sp of s.spills) {
      const size = extentSize(sp.extent);
      if (size && size.rows * size.cols === 1) single.push(sp.anchor);
      else real.push({ anchor: sp.anchor, extent: sp.extent, rows: size?.rows ?? null, cols: size?.cols ?? null });
    }
    if (real.length) spills[s.name] = real;
    if (single.length) dynamicArrayCells[s.name] = single;
  }

  const manifest: Json = {
    format: MANIFEST_FORMAT,
    workbook: input.fileName,
    sheets: wb.sheets.map((s) => ({ name: s.name, position: s.position, state: s.state, kind: s.kind, tables: s.tables })),
    tables: wb.tables.map((t) => ({
      name: t.displayName,
      sheet: t.sheet.name,
      ref: t.ref,
      headerRowCount: t.headerRowCount,
      totalsRowCount: t.totalsRowCount,
      columns: t.columns.map((c) => c.name),
      calculatedColumns: t.columns.some((c) => c.calculatedColumnFormula !== undefined)
        ? Object.fromEntries(t.columns.filter((c) => c.calculatedColumnFormula !== undefined).map((c) => [c.name, c.calculatedColumnFormula!]))
        : undefined,
    })),
    spills,
    dynamicArrayCells,
    unnamedCells: unnamedCells(input.statements ?? []),
    names: nameEntries,
    builtInNames: input.builtIns.map((b) => ({ name: b.name, scope: scopeText(b), definition: b.definition })),
    // Links to other workbooks (probe F10): what `[n]` in a stored formula names. Only when
    // there are some, so a manifest without links reads as before. Not the link's target:
    // it is often an absolute path naming the author's folders, and the project is shared.
    externalLinks: wb.externalLinks?.length ? wb.externalLinks.map((l) => ({ index: l.index, book: l.book })) : undefined,
    otherParts: wb.otherParts.map((p) => ({ kind: p.kind, path: p.path })),
    warnings: [...input.warnings],
  };
  return { manifest, unparsedFormulas };
}
