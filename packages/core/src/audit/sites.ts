// Every formula text the audit reads, with where it lives: defined names, cell formulas
// (a shared formula once, at its master, for its whole group), conditional formats, data
// validations and Table column formulas.

import type { DefinedName, WorkbookSnapshot } from "../file/types.js";
import { nameKey } from "../project/types.js";
import type { FindingWhere } from "./types.js";

export interface Site {
  where: FindingWhere;
  /** As stored in the file. */
  stored: string;
  /** The sheet unqualified names and references belong to: a cell's sheet, a sheet-scoped name's scope. */
  home: string | undefined;
  /** For a defined name. */
  name?: DefinedName;
}

/** Defined names the audit looks at: not Excel's own (`_xlnm.`, stray `_xlfn.` helpers), not with a broken scope. */
export function auditedNames(wb: WorkbookSnapshot): DefinedName[] {
  return wb.definedNames.filter((d) => !d.isXlPrefixed && !d.scopeInvalid);
}

export function scopeOf(d: DefinedName): string | undefined {
  return d.scope.kind === "sheet" ? d.scope.name : undefined;
}

export function nameWhere(d: DefinedName): FindingWhere {
  const scope = scopeOf(d);
  const w: FindingWhere = { kind: "name", name: d.name, key: nameKey({ name: d.name, scope }) };
  if (scope !== undefined) w.sheet = scope;
  return w;
}

export function collectSites(wb: WorkbookSnapshot): Site[] {
  const out: Site[] = [];
  for (const d of auditedNames(wb)) out.push({ where: nameWhere(d), stored: d.definition, home: scopeOf(d), name: d });
  for (const s of wb.sheets) {
    for (const f of s.formulas) {
      if (f.text === undefined || f.text.trim() === "" || f.kind === "data-table" || f.kind === "shared-child") continue;
      const where: FindingWhere = { kind: "cell", sheet: s.name, ref: f.cell };
      if (f.range !== undefined && f.range !== f.cell) where.range = f.range;
      out.push({ where, stored: f.text, home: s.name });
    }
    for (const cf of s.conditionalFormats) for (const text of cf.formulas) if (text.trim() !== "") out.push({ where: { kind: "cf", sheet: s.name, ref: cf.sqref }, stored: text, home: s.name });
    for (const dv of s.dataValidations) {
      for (const text of [dv.formula1, dv.formula2]) if (text !== undefined && text.trim() !== "") out.push({ where: { kind: "dv", sheet: s.name, ref: dv.sqref }, stored: text, home: s.name });
    }
  }
  for (const t of wb.tables) {
    for (const c of t.columns) {
      for (const text of [c.calculatedColumnFormula, c.totalsRowFormula]) {
        if (text !== undefined && text.trim() !== "") out.push({ where: { kind: "table", sheet: t.sheet.name, name: `${t.displayName}[${c.name}]` }, stored: text, home: t.sheet.name });
      }
    }
  }
  return out;
}
