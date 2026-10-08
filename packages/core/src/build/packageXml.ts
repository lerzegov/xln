// The small package parts a cell build touches besides the sheets: the workbook's
// relationships, `[Content_Types].xml` and `xl/metadata.xml`. Patched like the other parts,
// by splicing at tokenizer offsets, so whatever a change does not reach stays as written.

import { XmlReader, type XmlOpen } from "../file/xml.js";
import { escapeXmlAttr } from "./workbookXml.js";

interface Splice {
  start: number;
  end: number;
  text: string;
}

function splice(xml: string, edits: Splice[]): string {
  edits.sort((a, b) => a.start - b.start || a.end - b.end);
  let out = "";
  let at = 0;
  for (const e of edits) {
    out += xml.slice(at, e.start) + e.text;
    at = e.end;
  }
  return out + xml.slice(at);
}

interface Element {
  open: XmlOpen;
  start: number;
  end: number;
  /** Where the content ends: the close tag's start (the open tag's end when self-closing). */
  innerEnd: number;
  children: Element[];
}

/** The root and its children down to `maxDepth`, with offsets. */
function tree(xml: string, maxDepth: number): Element {
  const r = new XmlReader(xml);
  const stack: Element[] = [];
  let root: Element | undefined;
  for (let t = r.next(); t; t = r.next()) {
    if (t.type === "open") {
      const el: Element = { open: t, start: t.start, end: t.end, innerEnd: t.end, children: [] };
      if (stack.length > 0 && stack.length < maxDepth) stack[stack.length - 1]!.children.push(el);
      if (!root) root = el;
      stack.push(el);
    } else if (t.type === "close") {
      const el = stack.pop()!;
      el.innerEnd = t.start;
      el.end = t.end;
    }
  }
  if (!root) throw new Error("no root element");
  return root;
}

function prefixOf(name: string): string {
  const k = name.indexOf(":");
  return k < 0 ? "" : name.slice(0, k + 1);
}

/** `xml` with `text` added as the root's last child. */
function appendToRoot(xml: string, root: Element, text: string): string {
  if (root.open.selfClosing) {
    const tag = xml.slice(root.start, root.end - 2) + ">";
    return splice(xml, [{ start: root.start, end: root.end, text: tag + text + `</${root.open.name}>` }]);
  }
  return splice(xml, [{ start: root.innerEnd, end: root.innerEnd, text }]);
}

// ---------------------------------------------------------------------------------------
// Relationships

/** Removes every `<Relationship>` whose `Type` ends with `/<suffix>`; returns the targets removed. */
export function removeRelationships(xml: string, suffix: string): { xml: string; removed: string[] } {
  const root = tree(xml, 2);
  const removed: string[] = [];
  const edits: Splice[] = [];
  for (const c of root.children) {
    if (c.open.local === "Relationship" && (c.open.attrs["Type"] ?? "").endsWith("/" + suffix)) {
      edits.push({ start: c.start, end: c.end, text: "" });
      removed.push(c.open.attrs["Target"] ?? "");
    }
  }
  return { xml: edits.length ? splice(xml, edits) : xml, removed };
}

/** Adds a relationship with a fresh `rIdN`. */
export function addRelationship(xml: string, type: string, target: string): { xml: string; id: string } {
  const root = tree(xml, 2);
  const ids = new Set(root.children.map((c) => c.open.attrs["Id"]));
  let n = 1;
  while (ids.has(`rId${n}`)) n++;
  const id = `rId${n}`;
  const p = prefixOf(root.open.name);
  return { xml: appendToRoot(xml, root, `<${p}Relationship Id="${id}" Type="${escapeXmlAttr(type)}" Target="${escapeXmlAttr(target)}"/>`), id };
}

// ---------------------------------------------------------------------------------------
// [Content_Types].xml

/** Removes the `<Override>` for a part (`/xl/calcChain.xml`), matched ignoring case as OPC part names are. */
export function removeOverride(xml: string, partName: string): string {
  const root = tree(xml, 2);
  const want = partName.toLowerCase();
  const edits = root.children
    .filter((c) => c.open.local === "Override" && (c.open.attrs["PartName"] ?? "").toLowerCase() === want)
    .map((c) => ({ start: c.start, end: c.end, text: "" }));
  return edits.length ? splice(xml, edits) : xml;
}

export function addOverride(xml: string, partName: string, contentType: string): string {
  const root = tree(xml, 2);
  if (root.children.some((c) => c.open.local === "Override" && (c.open.attrs["PartName"] ?? "").toLowerCase() === partName.toLowerCase())) return xml;
  const p = prefixOf(root.open.name);
  return appendToRoot(xml, root, `<${p}Override PartName="${escapeXmlAttr(partName)}" ContentType="${escapeXmlAttr(contentType)}"/>`);
}

// ---------------------------------------------------------------------------------------
// xl/metadata.xml: the dynamic-array cell-metadata record

export const SHEET_METADATA_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
export const SHEET_METADATA_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
const XDA_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray";
const XDA_EXT = "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}";

// As Excel writes it (probes/results/f8_base.xlsx, and q1 in F8: Excel keeps it as is).
const XLDAPR_TYPE =
  'name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"';

function xldaprRecord(p: string, declare: boolean): string {
  return `<${p}bk><${p}extLst><${p}ext uri="${XDA_EXT}"><xda:dynamicArrayProperties${declare ? ` xmlns:xda="${XDA_NS}"` : ""} fDynamic="1" fCollapsed="0"/></${p}ext></${p}extLst></${p}bk>`;
}

/** A new `xl/metadata.xml` holding only the dynamic-array record; its `cm` is 1. */
export function newMetadataXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    `<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="${XDA_NS}">` +
    `<metadataTypes count="1"><metadataType ${XLDAPR_TYPE}/></metadataTypes>` +
    `<futureMetadata name="XLDAPR" count="1">${xldaprRecord("", false)}</futureMetadata>` +
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'
  );
}

function isTrue(v: string | undefined): boolean {
  return v === "1" || v === "true";
}

/** Whether a `<futureMetadata name="XLDAPR">` record says dynamic and not collapsed. */
function dynamicRecord(xml: string, bk: Element): boolean {
  // Read in the whole part, where the record's namespace prefixes are declared.
  const r = new XmlReader(xml);
  for (let t = r.next(); t; t = r.next()) {
    if (t.type !== "open" || t.start < bk.start) continue;
    if (t.start >= bk.end) break;
    if (t.local === "dynamicArrayProperties") return isTrue(t.attrs["fDynamic"]) && !isTrue(t.attrs["fCollapsed"]);
  }
  return false;
}

interface MetadataLayout {
  root: Element;
  types: Element | undefined;
  /** 1-based index of the XLDAPR type, if any. */
  xldapr: number | undefined;
  future: Element | undefined;
  cell: Element | undefined;
}

function metadataLayout(xml: string): MetadataLayout {
  const root = tree(xml, 4);
  const types = root.children.find((c) => c.open.local === "metadataTypes");
  const k = types ? types.children.filter((c) => c.open.local === "metadataType").findIndex((c) => c.open.attrs["name"] === "XLDAPR") : -1;
  return {
    root,
    types,
    xldapr: k < 0 ? undefined : k + 1,
    future: root.children.find((c) => c.open.local === "futureMetadata" && c.open.attrs["name"] === "XLDAPR"),
    cell: root.children.find((c) => c.open.local === "cellMetadata"),
  };
}

/** The 1-based `cellMetadata` record that marks a dynamic array, if the part has one. */
export function findDynamicArrayCm(xml: string): number | undefined {
  const l = metadataLayout(xml);
  if (l.xldapr === undefined || !l.future || !l.cell) return undefined;
  const records = l.future.children.filter((c) => c.open.local === "bk");
  const blocks = l.cell.children.filter((c) => c.open.local === "bk");
  for (let i = 0; i < blocks.length; i++) {
    const rcs = blocks[i]!.children.filter((c) => c.open.local === "rc");
    if (rcs.length !== 1) continue;
    const rc = rcs[0]!.open.attrs;
    if (Number(rc["t"]) !== l.xldapr) continue;
    const rec = records[Number(rc["v"])];
    if (rec && dynamicRecord(xml, rec)) return i + 1;
  }
  return undefined;
}

function withCount(el: Element, count: number): Splice {
  let s = `<${el.open.name}`;
  let placed = false;
  for (const [k, v] of Object.entries(el.open.attrs)) {
    if (k === "count") {
      s += ` count="${count}"`;
      placed = true;
    } else s += ` ${k}="${escapeXmlAttr(v)}"`;
  }
  if (!placed) s += ` count="${count}"`;
  return { start: el.start, end: el.open.end, text: s + (el.open.selfClosing ? "/>" : ">") };
}

/**
 * A metadata part that has a dynamic-array record: as it is when it has one, else with the
 * XLDAPR type, a record and a `cellMetadata` block added (a workbook whose metadata holds
 * only other types, rich values for instance). Returns the record's 1-based `cm`.
 */
export function ensureDynamicArrayCm(xml: string): { xml: string; cm: number; changed: boolean } {
  const found = findDynamicArrayCm(xml);
  if (found !== undefined) return { xml, cm: found, changed: false };
  const l = metadataLayout(xml);
  const p = prefixOf(l.root.open.name);
  const edits: Splice[] = [];
  const insertBefore = (names: string[]): number => {
    const next = l.root.children.find((c) => names.includes(c.open.local));
    return next ? next.start : l.root.innerEnd;
  };

  // The type.
  let typeIndex = l.xldapr;
  if (typeIndex === undefined) {
    const n = l.types ? l.types.children.filter((c) => c.open.local === "metadataType").length : 0;
    typeIndex = n + 1;
    const el = `<${p}metadataType ${XLDAPR_TYPE}/>`;
    if (!l.types) {
      const at = l.root.children[0]?.start ?? l.root.innerEnd;
      edits.push({ start: at, end: at, text: `<${p}metadataTypes count="1">${el}</${p}metadataTypes>` });
    } else if (l.types.open.selfClosing) {
      edits.push({ start: l.types.start, end: l.types.end, text: `<${p}metadataTypes count="1">${el}</${p}metadataTypes>` });
    } else {
      edits.push(withCount(l.types, n + 1));
      edits.push({ start: l.types.innerEnd, end: l.types.innerEnd, text: el });
    }
  }

  // The record.
  let recordIndex: number;
  if (l.future && !l.future.open.selfClosing) {
    recordIndex = l.future.children.filter((c) => c.open.local === "bk").length;
    edits.push(withCount(l.future, recordIndex + 1));
    edits.push({ start: l.future.innerEnd, end: l.future.innerEnd, text: xldaprRecord(p, true) });
  } else {
    recordIndex = 0;
    const block = `<${p}futureMetadata name="XLDAPR" count="1">${xldaprRecord(p, true)}</${p}futureMetadata>`;
    if (l.future) edits.push({ start: l.future.start, end: l.future.end, text: block });
    else {
      const at = insertBefore(["cellMetadata", "valueMetadata", "extLst"]);
      edits.push({ start: at, end: at, text: block });
    }
  }

  // The cell-metadata block.
  const rc = `<${p}bk><${p}rc t="${typeIndex}" v="${recordIndex}"/></${p}bk>`;
  let cm: number;
  if (l.cell && !l.cell.open.selfClosing) {
    cm = l.cell.children.filter((c) => c.open.local === "bk").length + 1;
    edits.push(withCount(l.cell, cm));
    edits.push({ start: l.cell.innerEnd, end: l.cell.innerEnd, text: rc });
  } else {
    cm = 1;
    const block = `<${p}cellMetadata count="1">${rc}</${p}cellMetadata>`;
    if (l.cell) edits.push({ start: l.cell.start, end: l.cell.end, text: block });
    else {
      const at = insertBefore(["valueMetadata", "extLst"]);
      edits.push({ start: at, end: at, text: block });
    }
  }
  return { xml: splice(xml, edits), cm, changed: true };
}
