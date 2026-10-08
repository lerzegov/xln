// Hand-made snapshots for the audit's unit tests: sheets of formula cells (a spill by its
// extent), names with scope and hidden flag, conditional formats, validations, Tables
// with column formulas, and charts.
import type { CellFormula, ChartPart, DefinedName, Sheet, Table, WorkbookSnapshot } from "../../src/index.js";

export type F = string | { text: string; spill?: string };

export interface Book {
  sheets: Record<string, Record<string, F>>;
  names?: { name: string; def: string; scope?: string; hidden?: boolean }[];
  cf?: { sheet: string; sqref: string; formula: string }[];
  dv?: { sheet: string; sqref: string; formula: string }[];
  tables?: { name: string; sheet: string; ref: string; columns: (string | { name: string; formula: string })[] }[];
  charts?: { sheet: string; formulas: string[] }[];
}

export function book(b: Book): WorkbookSnapshot {
  const order = Object.keys(b.sheets);
  const sheets: Sheet[] = order.map((name, position) => {
    const formulas: CellFormula[] = [];
    const spills: Sheet["spills"] = [];
    for (const [cell, f] of Object.entries(b.sheets[name]!)) {
      const x = typeof f === "string" ? { text: f } : f;
      const cf: CellFormula = { cell, kind: x.spill ? "dynamic-array" : "normal", text: x.text, attributes: {}, value: { type: "n", raw: "1", value: 1 } };
      if (x.spill) {
        cf.range = x.spill;
        spills.push({ anchor: cell, extent: x.spill });
      }
      formulas.push(cf);
    }
    return {
      name,
      sheetId: position + 1,
      position,
      state: "visible",
      kind: "worksheet",
      relId: `rId${position + 1}`,
      part: `xl/worksheets/sheet${position + 1}.xml`,
      formulas,
      sharedFormulas: [],
      spills,
      conditionalFormats: (b.cf ?? []).filter((c) => c.sheet === name).map((c) => ({ sqref: c.sqref, type: "expression", priority: 1, formulas: [c.formula], ext: false })),
      dataValidations: (b.dv ?? []).filter((d) => d.sheet === name).map((d) => ({ sqref: d.sqref, type: "list", formula1: d.formula, formula2: undefined, ext: false })),
      tables: (b.tables ?? []).filter((t) => t.sheet === name).map((t) => t.name),
    };
  });
  const definedNames: DefinedName[] = (b.names ?? []).map((n, index) => ({
    name: n.name,
    scope: n.scope === undefined ? { kind: "workbook" } : { kind: "sheet", position: order.indexOf(n.scope), name: n.scope },
    hidden: n.hidden ?? false,
    comment: undefined,
    definition: n.def,
    attributes: {},
    index,
    isXlPrefixed: n.name.toLowerCase().startsWith("_xl"),
    isBuiltIn: n.name.toLowerCase().startsWith("_xlnm."),
  }));
  const tables: Table[] = (b.tables ?? []).map((t, k) => ({
    id: k + 1,
    name: t.name,
    displayName: t.name,
    sheet: { position: order.indexOf(t.sheet), name: t.sheet },
    part: `xl/tables/table${k + 1}.xml`,
    ref: t.ref,
    headerRowCount: 1,
    totalsRowCount: 0,
    columns: t.columns.map((c, id) => (typeof c === "string" ? { id: id + 1, name: c } : { id: id + 1, name: c.name, calculatedColumnFormula: c.formula })),
  }));
  const charts: ChartPart[] = (b.charts ?? []).map((c, k) => ({
    part: `xl/charts/chart${k + 1}.xml`,
    kind: "chart",
    sheet: { position: order.indexOf(c.sheet), name: c.sheet },
    drawing: "xl/drawings/drawing1.xml",
    formulas: c.formulas.map((text) => ({ text, element: "val" })),
  }));
  return { sheets, definedNames, tables, charts, otherParts: [], parts: [], foreignModuleStores: [], workbookPart: "xl/workbook.xml", warnings: [] };
}
