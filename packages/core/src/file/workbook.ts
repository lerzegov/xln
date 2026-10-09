// readWorkbook: bytes of an .xlsx in, a WorkbookSnapshot out.
import { findForeignModuleStores } from "./afe.js";
import { readCharts } from "./charts.js";
import { Package, relTypeIs, XlsxError, type Relationship } from "./package.js";
import type {
  DefinedName,
  ExternalLink,
  NameScope,
  OtherPart,
  OtherPartKind,
  Sheet,
  SheetKind,
  SheetState,
  Table,
  TableColumn,
  WorkbookSnapshot,
} from "./types.js";
import { readWorksheet, REL_NS, richText } from "./worksheet.js";
import { decodeXstring } from "./xstring.js";
import { attrNS, childElements, descendants, firstChild, ownText, parseXml, type XmlElement } from "./xml.js";

export function readWorkbook(bytes: Uint8Array): WorkbookSnapshot {
  const pkg = new Package(bytes);
  const warnings: string[] = [];
  const warn = (m: string) => warnings.push(m);

  const workbookPart = findWorkbookPart(pkg);
  const wbText = pkg.text(workbookPart);
  if (wbText === undefined) throw new XlsxError(`workbook part ${workbookPart} is missing`);
  const wb = parseXml(wbText);
  const wbRels = pkg.rels(workbookPart);
  const relById = new Map(wbRels.map((r) => [r.id, r]));

  const sharedStrings = readSharedStrings(pkg, wbRels);
  const dynamicArrayCm = readDynamicArrayMetadata(pkg, wbRels);

  const sheets: Sheet[] = [];
  const sheetsEl = firstChild(wb, "sheets");
  for (const s of sheetsEl ? childElements(sheetsEl, "sheet") : []) {
    const relId = attrNS(s, REL_NS, "id") ?? "";
    const rel = relById.get(relId);
    const name = s.attrs["name"] ?? "";
    let part: string | undefined;
    if (!rel) warn(`sheet "${name}": relationship ${relId} not found`);
    else if (rel.external) warn(`sheet "${name}": external target ${rel.target}`);
    else if (!pkg.has(rel.target)) warn(`sheet "${name}": part ${rel.target} is missing`);
    else part = pkg.find(rel.target);
    sheets.push({
      name,
      sheetId: Number(s.attrs["sheetId"]),
      position: sheets.length,
      state: sheetState(s.attrs["state"]),
      kind: rel ? sheetKind(rel.type) : "unknown",
      relId,
      part,
      formulas: [],
      sharedFormulas: [],
      spills: [],
      conditionalFormats: [],
      dataValidations: [],
      tables: [],
    });
  }

  const definedNames = readDefinedNames(wb, sheets, warn);
  const fullCalc = firstChild(wb, "calcPr")?.attrs["fullCalcOnLoad"];

  const tables: Table[] = [];
  for (const sheet of sheets) {
    if (sheet.part === undefined || sheet.kind === "chartsheet") continue;
    const text = pkg.text(sheet.part);
    if (text === undefined) continue;
    const data = readWorksheet(text, sheet.part, { sharedStrings, dynamicArrayCm, warn });
    sheet.formulas = data.formulas;
    sheet.sharedFormulas = data.sharedFormulas;
    sheet.spills = data.spills;
    sheet.conditionalFormats = data.conditionalFormats;
    sheet.dataValidations = data.dataValidations;

    const rels = pkg.rels(sheet.part).filter((r) => relTypeIs(r.type, "table") && !r.external);
    const byId = new Map(rels.map((r) => [r.id, r]));
    const ordered: Relationship[] = [];
    for (const id of data.tablePartIds) {
      const r = byId.get(id);
      if (r) ordered.push(r);
      else warn(`${sheet.part}: tablePart ${id} has no relationship`);
    }
    for (const r of rels) if (!ordered.includes(r)) ordered.push(r);
    for (const r of ordered) {
      const tText = pkg.text(r.target);
      if (tText === undefined) {
        warn(`${sheet.part}: table part ${r.target} is missing`);
        continue;
      }
      const table = readTable(parseXml(tText), pkg.find(r.target) ?? r.target, sheet);
      tables.push(table);
      sheet.tables.push(table.displayName);
    }
  }

  return {
    sheets,
    definedNames,
    tables,
    charts: readCharts(pkg, sheets, warn),
    externalLinks: readExternalLinks(pkg, wb, relById, warn),
    otherParts: classifyParts(pkg.names),
    parts: [...pkg.names],
    foreignModuleStores: findForeignModuleStores(pkg, sheets),
    workbookPart,
    ...(fullCalc === "1" || fullCalc === "true" ? { fullCalcOnLoad: true as const } : {}),
    warnings,
  };
}

/**
 * The links to other workbooks: each `<externalReference r:id>` names an externalLink part,
 * whose `<externalBook r:id>` names the file (an external relationship). Measured on probe
 * F10's file (Excel for Mac): the part also lists the file's sheets, names and cached values.
 */
function readExternalLinks(pkg: Package, wb: XmlElement, relById: ReadonlyMap<string, Relationship>, warn: (m: string) => void): ExternalLink[] {
  const out: ExternalLink[] = [];
  const refs = firstChild(wb, "externalReferences");
  for (const r of refs ? childElements(refs, "externalReference") : []) {
    const index = out.length + 1;
    const rel = relById.get(attrNS(r, REL_NS, "id") ?? "");
    const part = rel && !rel.external ? pkg.find(rel.target) : undefined;
    const text = part !== undefined ? pkg.text(part) : undefined;
    if (part === undefined || text === undefined) {
      warn(`external reference [${index}]: its part is missing`);
      out.push({ index, book: "", target: "", part: rel?.target ?? "" });
      continue;
    }
    const book = firstChild(parseXml(text), "externalBook");
    const target = book ? pkg.rels(part).find((x) => x.id === attrNS(book, REL_NS, "id"))?.target : undefined;
    if (target === undefined) warn(`external reference [${index}] (${part}): no workbook path`);
    out.push({ index, book: target !== undefined ? fileNameOf(target) : "", target: target ?? "", part });
  }
  return out;
}

/** `Other.xlsx` of `/Users/x/Other.xlsx`, `file:///C:\\dir\\Other.xlsx`, `My%20File.xlsx`. */
function fileNameOf(target: string): string {
  const cut = Math.max(target.lastIndexOf("/"), target.lastIndexOf("\\"));
  const name = target.slice(cut + 1);
  try {
    return decodeURIComponent(name);
  } catch {
    return name;
  }
}

function findWorkbookPart(pkg: Package): string {
  for (const r of pkg.rels("")) {
    if (relTypeIs(r.type, "officeDocument") && !r.external && pkg.has(r.target)) return pkg.find(r.target)!;
  }
  const fallback = pkg.find("xl/workbook.xml");
  if (fallback) return fallback;
  throw new XlsxError("no workbook part: not an .xlsx package");
}

function sheetState(s: string | undefined): SheetState {
  return s === "hidden" || s === "veryHidden" ? s : "visible";
}

function sheetKind(type: string): SheetKind {
  for (const k of ["worksheet", "chartsheet", "dialogsheet"] as const) if (relTypeIs(type, k)) return k;
  if (relTypeIs(type, "xlMacrosheet") || relTypeIs(type, "xlIntlMacrosheet")) return "macrosheet";
  return "unknown";
}

function isTrue(v: string | undefined): boolean {
  return v === "1" || v === "true";
}

const NAME_CORE_ATTRS = new Set(["name", "localSheetId", "hidden", "comment"]);

function readDefinedNames(wb: XmlElement, sheets: Sheet[], warn: (m: string) => void): DefinedName[] {
  const out: DefinedName[] = [];
  const container = firstChild(wb, "definedNames");
  if (!container) return out;
  for (const el of childElements(container, "definedName")) {
    const name = el.attrs["name"] ?? "";
    const lsid = el.attrs["localSheetId"];
    let scope: NameScope = { kind: "workbook" };
    let scopeInvalid = false;
    if (lsid !== undefined) {
      const position = Number(lsid);
      const sheet = Number.isInteger(position) ? sheets[position] : undefined;
      if (!sheet) {
        scopeInvalid = true;
        warn(`defined name "${name}": localSheetId ${lsid} is not a sheet position`);
      }
      scope = { kind: "sheet", position, name: sheet?.name ?? "" };
    }
    const attributes: Record<string, string> = {};
    for (const [k, v] of Object.entries(el.attrs)) if (!NAME_CORE_ATTRS.has(k)) attributes[k] = v;
    const lower = name.toLowerCase();
    const dn: DefinedName = {
      name,
      scope,
      hidden: isTrue(el.attrs["hidden"]),
      comment: el.attrs["comment"] === undefined ? undefined : decodeXstring(el.attrs["comment"]),
      definition: ownText(el),
      attributes,
      index: out.length,
      isXlPrefixed: lower.startsWith("_xl"),
      isBuiltIn: lower.startsWith("_xlnm."),
    };
    if (scopeInvalid) dn.scopeInvalid = true;
    out.push(dn);
  }
  return out;
}

function readSharedStrings(pkg: Package, wbRels: Relationship[]): string[] {
  const rel = wbRels.find((r) => relTypeIs(r.type, "sharedStrings") && !r.external);
  const text = rel ? pkg.text(rel.target) : undefined;
  if (text === undefined) return [];
  return childElements(parseXml(text), "si").map(richText);
}

// Dynamic-array formulas are marked by the cell's `cm` attribute: a 1-based index into
// <cellMetadata>, whose <rc t v> points at metadata type t (1-based) and record v
// (0-based) of the <futureMetadata> block of that type. Type XLDAPR, with
// dynamicArrayProperties fDynamic="1", is a dynamic array.
function readDynamicArrayMetadata(pkg: Package, wbRels: Relationship[]): Set<number> {
  const out = new Set<number>();
  const rel = wbRels.find((r) => relTypeIs(r.type, "sheetMetadata") && !r.external);
  const text = rel ? pkg.text(rel.target) : undefined;
  if (text === undefined) return out;
  const md = parseXml(text);
  const typesEl = firstChild(md, "metadataTypes");
  const types = typesEl ? childElements(typesEl, "metadataType").map((t) => t.attrs["name"] ?? "") : [];
  const future = new Map<string, XmlElement[]>();
  for (const fm of childElements(md, "futureMetadata")) {
    future.set(fm.attrs["name"] ?? "", childElements(fm, "bk"));
  }
  const cellMd = firstChild(md, "cellMetadata");
  if (!cellMd) return out;
  childElements(cellMd, "bk").forEach((bk, i) => {
    for (const rc of childElements(bk, "rc")) {
      const typeName = types[Number(rc.attrs["t"]) - 1];
      if (typeName !== "XLDAPR") continue;
      const record = future.get(typeName)?.[Number(rc.attrs["v"])];
      let dynamic = true; // XLDAPR without a readable record: trust the type
      if (record) {
        for (const d of descendants(record)) {
          if (d.local === "dynamicArrayProperties") dynamic = isTrue(d.attrs["fDynamic"]);
        }
      }
      if (dynamic) out.add(i + 1);
    }
  });
  return out;
}

function readTable(el: XmlElement, part: string, sheet: Sheet): Table {
  const columns: TableColumn[] = [];
  const cols = firstChild(el, "tableColumns");
  for (const c of cols ? childElements(cols, "tableColumn") : []) {
    const col: TableColumn = { id: Number(c.attrs["id"]), name: c.attrs["name"] ?? "" };
    const calc = firstChild(c, "calculatedColumnFormula");
    if (calc) {
      col.calculatedColumnFormula = ownText(calc);
      if (isTrue(calc.attrs["array"])) col.calculatedColumnArray = true;
    }
    const tot = firstChild(c, "totalsRowFormula");
    if (tot) col.totalsRowFormula = ownText(tot);
    const fn = c.attrs["totalsRowFunction"];
    if (fn !== undefined) col.totalsRowFunction = fn;
    columns.push(col);
  }
  const name = el.attrs["name"] ?? "";
  return {
    id: Number(el.attrs["id"]),
    name,
    displayName: el.attrs["displayName"] ?? name,
    sheet: { position: sheet.position, name: sheet.name },
    part,
    ref: el.attrs["ref"] ?? "",
    headerRowCount: el.attrs["headerRowCount"] !== undefined ? Number(el.attrs["headerRowCount"]) : 1,
    totalsRowCount: el.attrs["totalsRowCount"] !== undefined ? Number(el.attrs["totalsRowCount"]) : 0,
    columns,
  };
}

const PART_PREFIXES: [string, OtherPartKind][] = [
  ["xl/charts/chart", "chart"],
  ["xl/chartsheets/", "chartsheet"],
  ["xl/drawings/drawing", "drawing"],
  ["xl/pivottables/", "pivotTable"],
  ["xl/pivotcache/pivotcachedefinition", "pivotCache"],
  ["xl/externallinks/", "externalLink"],
  ["customxml/item", "customXml"],
  ["xl/querytables/", "queryTable"],
  ["xl/connections.xml", "connections"],
  ["xl/slicers/", "slicer"],
  ["xl/slicercaches/", "slicerCache"],
  ["xl/timelines/", "timeline"],
  ["xl/vbaproject.bin", "vbaProject"],
];

function classifyParts(names: readonly string[]): OtherPart[] {
  const out: OtherPart[] = [];
  for (const path of names) {
    const lower = path.toLowerCase();
    if (lower.includes("/_rels/") || lower.endsWith("/")) continue;
    if (lower.startsWith("customxml/itemprops")) continue;
    const hit = PART_PREFIXES.find(([p]) => lower.startsWith(p));
    if (hit) out.push({ kind: hit[1], path });
  }
  return out;
}
