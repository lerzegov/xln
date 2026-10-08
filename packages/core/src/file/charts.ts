// Chart parts: the formulas their series, titles and labels read (`<c:f>`), and the sheet
// whose drawing holds each chart. A defined name can be used only in a chart (a series
// over `[0]!Sales`), so the audit must see these before it calls a name unused (C10), and
// a rename (stretch G) must find them.
import { Package, relTypeIs } from "./package.js";
import type { ChartFormula, ChartPart, Sheet } from "./types.js";
import { ownText, parseXml, type XmlElement } from "./xml.js";

const CHART_REL = "chart";
// Office 2016 charts (waterfall, histogram, ...) live in `chartEx` parts.
const CHART_EX_REL = "chartEx";

function isChartRel(type: string): "chart" | "chartEx" | undefined {
  if (relTypeIs(type, CHART_REL)) return "chart";
  if (relTypeIs(type, CHART_EX_REL)) return "chartEx";
  return undefined;
}

/** Every `<…:f>` formula in the chart, with the element that holds it (`val`, `cat`, `tx`, ...). */
function chartFormulas(root: XmlElement): ChartFormula[] {
  const out: ChartFormula[] = [];
  const visit = (el: XmlElement, parent: string): void => {
    for (const c of el.children) {
      if (typeof c === "string") continue;
      if (c.local === "f") {
        const text = ownText(c).trim();
        if (text !== "") out.push({ text, element: parent });
        continue;
      }
      // `numRef` / `strRef` / `multiLvlStrRef` only wrap the formula: name the element above them.
      const holder = c.local.endsWith("Ref") ? parent : c.local;
      visit(c, holder);
    }
  };
  visit(root, root.local);
  return out;
}

/**
 * The chart parts of the package, in package order, each with the sheet (worksheet or
 * chartsheet) whose drawing shows it. A chart no sheet's drawing reaches is listed with
 * no sheet.
 */
export function readCharts(pkg: Package, sheets: readonly Sheet[], warn: (m: string) => void): ChartPart[] {
  const owner = new Map<string, { sheet: Sheet; drawing: string; kind: "chart" | "chartEx" }>();
  for (const sheet of sheets) {
    if (sheet.part === undefined) continue;
    for (const r of pkg.rels(sheet.part)) {
      if (r.external || !relTypeIs(r.type, "drawing")) continue;
      const drawing = pkg.find(r.target);
      if (drawing === undefined) continue;
      for (const c of pkg.rels(drawing)) {
        const kind = isChartRel(c.type);
        const part = kind && !c.external ? pkg.find(c.target) : undefined;
        if (kind && part !== undefined && !owner.has(part)) owner.set(part, { sheet, drawing, kind });
      }
    }
  }
  const out: ChartPart[] = [];
  for (const path of pkg.names) {
    const lower = path.toLowerCase();
    if (!lower.startsWith("xl/charts/chart") || !lower.endsWith(".xml")) continue;
    const text = pkg.text(path);
    if (text === undefined) continue;
    let root: XmlElement;
    try {
      root = parseXml(text);
    } catch (e) {
      warn(`${path}: chart part is not readable XML (${(e as Error).message})`);
      continue;
    }
    const o = owner.get(path);
    const part: ChartPart = {
      part: path,
      kind: o?.kind ?? (lower.startsWith("xl/charts/chartex") ? "chartEx" : "chart"),
      formulas: chartFormulas(root),
    };
    if (o) {
      part.sheet = { position: o.sheet.position, name: o.sheet.name };
      part.drawing = o.drawing;
    }
    out.push(part);
  }
  return out;
}
