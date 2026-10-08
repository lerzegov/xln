// Streaming reader for a worksheet part: formula cells, conditional formatting, data
// validation and table parts. Only formula cells are kept; plain values are skipped.
import { formatCell, parseCell } from "./cellref.js";
import type {
  CachedValue,
  CellFormula,
  CellValueType,
  ConditionalFormat,
  DataValidation,
  SharedFormulaGroup,
  Spill,
} from "./types.js";
import {
  attrNS,
  childElements,
  descendants,
  firstChild,
  ownText,
  readElement,
  XmlReader,
  type XmlElement,
} from "./xml.js";

export const REL_NS = [
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
] as const;

export interface WorksheetContext {
  sharedStrings: readonly string[];
  /** 1-based `cm` indices whose cell metadata marks a dynamic array. */
  dynamicArrayCm: ReadonlySet<number>;
  warn(message: string): void;
}

export interface WorksheetData {
  formulas: CellFormula[];
  sharedFormulas: SharedFormulaGroup[];
  spills: Spill[];
  conditionalFormats: ConditionalFormat[];
  dataValidations: DataValidation[];
  /** Relationship ids from `<tableParts>`, in order. */
  tablePartIds: string[];
}

const VALUE_TYPES: ReadonlySet<string> = new Set(["n", "s", "str", "b", "e", "inlineStr", "d"]);

export function readWorksheet(xml: string, part: string, ctx: WorksheetContext): WorksheetData {
  const out: WorksheetData = {
    formulas: [],
    sharedFormulas: [],
    spills: [],
    conditionalFormats: [],
    dataValidations: [],
    tablePartIds: [],
  };
  const groups = new Map<number, SharedFormulaGroup>();
  const reader = new XmlReader(xml);

  let t = reader.next();
  while (t && t.type !== "open") t = reader.next();
  if (!t) return out;
  // `t` is the root (<worksheet>, or <dialogsheet>/<macrosheet>, same children).
  for (let c = reader.next(); c; c = reader.next()) {
    if (c.type !== "open") continue;
    switch (c.local) {
      case "sheetData":
        if (!c.selfClosing) readSheetData(reader, part, ctx, out, groups);
        else reader.skip(c);
        break;
      case "conditionalFormatting":
        readConditionalFormatting(readElement(reader, c), out.conditionalFormats);
        break;
      case "dataValidations":
        for (const dv of childElements(readElement(reader, c), "dataValidation")) {
          out.dataValidations.push({
            sqref: dv.attrs["sqref"] ?? "",
            type: dv.attrs["type"],
            formula1: optText(firstChild(dv, "formula1")),
            formula2: optText(firstChild(dv, "formula2")),
            ext: false,
          });
        }
        break;
      case "tableParts":
        for (const tp of childElements(readElement(reader, c), "tablePart")) {
          const id = attrNS(tp, REL_NS, "id");
          if (id !== undefined) out.tablePartIds.push(id);
        }
        break;
      case "extLst":
        readExtensions(readElement(reader, c), out);
        break;
      default:
        reader.skip(c);
    }
  }

  for (const f of out.formulas) {
    if (f.kind !== "shared-child" || f.si === undefined) continue;
    const g = groups.get(f.si);
    if (g?.master !== undefined) f.master = g.master;
    else ctx.warn(`${part}: shared formula si=${f.si} in ${f.cell} has no master`);
  }
  out.sharedFormulas = [...groups.values()].sort((a, b) => a.si - b.si);
  return out;
}

function readSheetData(
  reader: XmlReader,
  part: string,
  ctx: WorksheetContext,
  out: WorksheetData,
  groups: Map<number, SharedFormulaGroup>,
): void {
  const base = reader.depth; // depth inside <sheetData>
  let row = 0;
  let col = 0;
  for (let t = reader.next(); t; t = reader.next()) {
    if (t.type === "close") {
      if (reader.depth < base) return;
      continue;
    }
    if (t.type !== "open") continue;
    if (t.local === "row") {
      const r = t.attrs["r"];
      row = r !== undefined ? Number(r) : row + 1;
      col = 0;
      continue; // descend into the row
    }
    if (t.local !== "c") {
      reader.skip(t);
      continue;
    }
    const cell = readElement(reader, t);
    const r = cell.attrs["r"];
    const addr = r !== undefined ? parseCell(r) : { row, col: col + 1 };
    if (!addr) {
      ctx.warn(`${part}: unreadable cell address "${r}"`);
      continue;
    }
    col = addr.col;
    const f = firstChild(cell, "f");
    if (!f) continue;
    const ref = formatCell(addr);
    const formula = readFormula(cell, f, ref, part, ctx);
    out.formulas.push(formula);
    if (formula.kind === "dynamic-array") {
      out.spills.push({ anchor: ref, extent: formula.range ?? ref });
    }
    if (formula.si !== undefined) {
      let g = groups.get(formula.si);
      if (!g) {
        g = { si: formula.si, master: undefined, range: undefined, text: undefined, cells: [] };
        groups.set(formula.si, g);
      }
      if (formula.kind === "shared-master") {
        if (g.master !== undefined) ctx.warn(`${part}: shared formula si=${formula.si} has two masters`);
        g.master = ref;
        g.range = formula.range;
        g.text = formula.text;
        g.cells.unshift(ref);
      } else {
        g.cells.push(ref);
      }
    }
  }
}

function readFormula(cell: XmlElement, f: XmlElement, ref: string, part: string, ctx: WorksheetContext): CellFormula {
  const ft = f.attrs["t"] ?? "normal";
  const range = f.attrs["ref"];
  const siRaw = f.attrs["si"];
  const cmRaw = cell.attrs["cm"];
  const cm = cmRaw !== undefined ? Number(cmRaw) : undefined;
  const attributes: Record<string, string> = {};
  for (const [k, v] of Object.entries(f.attrs)) {
    if (k !== "t" && k !== "ref" && k !== "si") attributes[k] = v;
  }
  const stored = ownText(f);
  let kind: CellFormula["kind"];
  switch (ft) {
    case "shared":
      kind = range !== undefined ? "shared-master" : "shared-child";
      break;
    case "array":
      kind = cm !== undefined && ctx.dynamicArrayCm.has(cm) ? "dynamic-array" : "array";
      break;
    case "dataTable":
      kind = "data-table";
      break;
    case "normal":
      kind = "normal";
      break;
    default:
      ctx.warn(`${part}: ${ref} has unknown formula type t="${ft}"`);
      kind = "normal";
  }
  const out: CellFormula = {
    cell: ref,
    kind,
    // A child normally has no text; keep any a non-Excel writer put there.
    text: kind === "shared-child" && stored === "" ? undefined : stored,
    attributes,
    value: readValue(cell, ref, part, ctx),
  };
  if (range !== undefined) out.range = range;
  if (siRaw !== undefined && ft === "shared") out.si = Number(siRaw);
  if (cm !== undefined) out.cm = cm;
  return out;
}

function readValue(cell: XmlElement, ref: string, part: string, ctx: WorksheetContext): CachedValue {
  const tRaw = cell.attrs["t"] ?? "n";
  const type: CellValueType = VALUE_TYPES.has(tRaw) ? (tRaw as CellValueType) : "n";
  if (type === "inlineStr") {
    const is = firstChild(cell, "is");
    const text = is ? richText(is) : undefined;
    return { type, raw: text, value: text };
  }
  const v = firstChild(cell, "v");
  const raw = v ? ownText(v) : undefined;
  if (raw === undefined) return { type, raw, value: undefined };
  switch (type) {
    case "n":
      return { type, raw, value: raw === "" ? undefined : Number(raw) };
    case "b":
      return { type, raw, value: raw === "1" || raw === "true" };
    case "s": {
      const s = ctx.sharedStrings[Number(raw)];
      if (s === undefined) ctx.warn(`${part}: ${ref} points at missing shared string ${raw}`);
      return { type, raw, value: s };
    }
    default:
      return { type, raw, value: raw };
  }
}

/** Text of a string item (`<si>`, `<is>`): plain `<t>` or rich-text runs, without phonetic runs. */
export function richText(el: XmlElement): string {
  let s = "";
  for (const c of childElements(el)) {
    if (c.local === "t") s += ownText(c);
    else if (c.local === "r") {
      const t = firstChild(c, "t");
      if (t) s += ownText(t);
    }
  }
  return s;
}

function optText(el: XmlElement | undefined): string | undefined {
  return el ? ownText(el) : undefined;
}

function optNumber(s: string | undefined): number | undefined {
  return s === undefined ? undefined : Number(s);
}

function readConditionalFormatting(cf: XmlElement, out: ConditionalFormat[]): void {
  const sqref = cf.attrs["sqref"] ?? "";
  for (const rule of childElements(cf, "cfRule")) {
    const formulas = childElements(rule, "formula").map(ownText);
    // Colour scales, data bars and icon sets can take formula thresholds.
    for (const d of descendants(rule)) {
      if (d.local === "cfvo" && d.attrs["type"] === "formula" && d.attrs["val"] !== undefined) {
        formulas.push(d.attrs["val"]);
      }
    }
    out.push({ sqref, type: rule.attrs["type"], priority: optNumber(rule.attrs["priority"]), formulas, ext: false });
  }
}

// x14 extension lists hold newer CF rules and data validations that point at other
// sheets: `<x14:conditionalFormatting><x14:cfRule>…<xm:f>…</xm:f></x14:cfRule><xm:sqref>`.
function readExtensions(extLst: XmlElement, out: WorksheetData): void {
  for (const d of descendants(extLst)) {
    if (d.local === "conditionalFormatting") {
      const sqref = optText(firstChild(d, "sqref")) ?? d.attrs["sqref"] ?? "";
      for (const rule of childElements(d, "cfRule")) {
        const formulas: string[] = [];
        for (const x of descendants(rule)) if (x.local === "f") formulas.push(ownText(x));
        out.conditionalFormats.push({
          sqref,
          type: rule.attrs["type"],
          priority: optNumber(rule.attrs["priority"]),
          formulas,
          ext: true,
        });
      }
    } else if (d.local === "dataValidation") {
      const f1 = firstChild(d, "formula1");
      const f2 = firstChild(d, "formula2");
      out.dataValidations.push({
        sqref: optText(firstChild(d, "sqref")) ?? d.attrs["sqref"] ?? "",
        type: d.attrs["type"],
        formula1: f1 ? optText(firstChild(f1, "f")) ?? ownText(f1) : undefined,
        formula2: f2 ? optText(firstChild(f2, "f")) ?? ownText(f2) : undefined,
        ext: true,
      });
    }
  }
}
