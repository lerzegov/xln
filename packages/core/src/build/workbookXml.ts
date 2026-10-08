// Patching xl/workbook.xml: only `<definedNames>` and the `fullCalcOnLoad` attribute of
// `<calcPr>` change. The part is read with the XML tokenizer (never a regular expression:
// probe F7's first run corrupted a sheet that way), and the new text is spliced in at the
// tokens' offsets, so every other byte of the part stays as Excel wrote it.

import { XmlError, XmlReader, type XmlOpen } from "../file/xml.js";
import { encodeXstring } from "../file/xstring.js";

/** A `<definedName>` element as found in the part. */
export interface DefinedNameElement {
  /** Offsets of the whole element, open tag to close tag. */
  start: number;
  end: number;
  attrs: Record<string, string>;
  /** Text content, entities decoded, CR LF kept. */
  text: string;
}

export interface WorkbookXmlLayout {
  /** Qualified-name prefix of the root (`""`, or `x:` for a prefixed SpreadsheetML part). */
  prefix: string;
  /** Sheet names in `<sheets>` order: a name's `localSheetId` counts these. */
  sheets: string[];
  definedNames:
    | {
        start: number;
        end: number;
        /** Where the content starts and ends (between the tags). */
        innerStart: number;
        innerEnd: number;
        /** Written `<definedNames/>`: no inner range to splice into. */
        selfClosing: boolean;
        items: DefinedNameElement[];
      }
    | undefined;
  calcPr: { start: number; end: number; open: XmlOpen } | undefined;
  /** Where a missing `<definedNames>` goes: before the first element that the schema puts after it. */
  insertAt: number;
}

// CT_Workbook's sequence after definedNames (ECMA-376 Part 1, §18.2.27).
const AFTER_DEFINED_NAMES = new Set([
  "calcPr",
  "oleSize",
  "customWorkbookViews",
  "pivotCaches",
  "smartTagPr",
  "smartTagTypes",
  "webPublishing",
  "fileRecoveryPr",
  "webPublishObjects",
  "extLst",
]);

export function scanWorkbookXml(xml: string): WorkbookXmlLayout {
  const r = new XmlReader(xml);
  let depth = 0;
  let prefix = "";
  let rootClose = -1;
  let insertAt = -1;
  const sheets: string[] = [];
  let definedNames: WorkbookXmlLayout["definedNames"];
  let calcPr: WorkbookXmlLayout["calcPr"];
  /** The direct child of the root being read, and of `<definedNames>`. */
  let child: XmlOpen | undefined;
  let item: { open: XmlOpen; text: string } | undefined;
  for (let t = r.next(); t; t = r.next()) {
    if (t.type === "open") {
      depth++;
      if (depth === 1) {
        if (t.local !== "workbook") throw new XmlError(`root element is <${t.name}>, not <workbook>`, t.start);
        prefix = t.name.includes(":") ? t.name.slice(0, t.name.indexOf(":") + 1) : "";
      } else if (depth === 2) {
        child = t;
        if (insertAt < 0 && AFTER_DEFINED_NAMES.has(t.local)) insertAt = t.start;
        if (t.local === "calcPr") calcPr = { start: t.start, end: t.end, open: t };
        if (t.local === "definedNames") {
          if (definedNames) throw new XmlError("more than one <definedNames>", t.start);
          definedNames = { start: t.start, end: t.end, innerStart: t.end, innerEnd: t.end, selfClosing: t.selfClosing, items: [] };
        }
      } else if (depth === 3 && child?.local === "sheets" && t.local === "sheet") {
        sheets.push(t.attrs["name"] ?? "");
      } else if (depth === 3 && child?.local === "definedNames") {
        if (t.local !== "definedName") throw new XmlError(`unexpected <${t.name}> in <definedNames>`, t.start);
        item = { open: t, text: "" };
      } else if (depth > 3 && child?.local === "definedNames") {
        throw new XmlError(`unexpected <${t.name}> inside a <definedName>`, t.start);
      }
    } else if (t.type === "text") {
      if (item && depth === 3) item.text += t.text;
    } else {
      if (depth === 3 && item) {
        definedNames!.items.push({ start: item.open.start, end: t.end, attrs: item.open.attrs, text: item.text });
        item = undefined;
      } else if (depth === 2 && child) {
        if (child.local === "definedNames") {
          definedNames!.innerEnd = t.start;
          definedNames!.end = t.end;
        }
        child = undefined;
      } else if (depth === 1) {
        rootClose = t.start;
      }
      depth--;
    }
  }
  if (rootClose < 0) throw new XmlError("no <workbook> element", 0);
  return { prefix, sheets, definedNames, calcPr, insertAt: insertAt < 0 ? rootClose : insertAt };
}

/** Escapes element text. `>` too, as Excel writes it. CR and LF stay raw (Excel stores CR LF). */
export function escapeXmlText(s: string): string {
  let out = "";
  for (const c of s) {
    if (c === "&") out += "&amp;";
    else if (c === "<") out += "&lt;";
    else if (c === ">") out += "&gt;";
    else out += c;
  }
  return out;
}

/** Escapes an attribute value in double quotes. Line breaks and tabs become character
 *  references, or a reader's attribute normalisation would turn them into spaces. */
export function escapeXmlAttr(s: string): string {
  let out = "";
  for (const c of s) {
    if (c === "&") out += "&amp;";
    else if (c === "<") out += "&lt;";
    else if (c === ">") out += "&gt;";
    else if (c === '"') out += "&quot;";
    else if (c === "\r") out += "&#13;";
    else if (c === "\n") out += "&#10;";
    else if (c === "\t") out += "&#9;";
    else out += c;
  }
  return out;
}

/** A `<definedName>` to write. */
export interface DefinedNameOut {
  name: string;
  comment: string | undefined;
  /** 0-based sheet position, or undefined for workbook scope. */
  localSheetId: number | undefined;
  hidden: boolean;
  /** Other attributes (function, vbProcedure, …), kept from the original element. */
  other: Record<string, string>;
  definition: string;
}

/** Serialises one element, attributes in Excel's order: name, comment, others, localSheetId, hidden. */
export function definedNameXml(d: DefinedNameOut, prefix = ""): string {
  let s = `<${prefix}definedName name="${escapeXmlAttr(d.name)}"`;
  if (d.comment !== undefined && d.comment !== "") s += ` comment="${escapeXmlAttr(encodeXstring(d.comment))}"`;
  for (const [k, v] of Object.entries(d.other)) s += ` ${k}="${escapeXmlAttr(v)}"`;
  if (d.localSheetId !== undefined) s += ` localSheetId="${d.localSheetId}"`;
  if (d.hidden) s += ` hidden="1"`;
  return s + `>${escapeXmlText(d.definition)}</${prefix}definedName>`;
}

/** The `<calcPr>` open tag with `fullCalcOnLoad="1"`, other attributes kept in order. */
function calcPrTag(open: XmlOpen): string {
  let s = `<${open.name}`;
  let seen = false;
  for (const [k, v] of Object.entries(open.attrs)) {
    if (k === "fullCalcOnLoad") {
      seen = true;
      s += ` ${k}="1"`;
    } else s += ` ${k}="${escapeXmlAttr(v)}"`;
  }
  if (!seen) s += ` fullCalcOnLoad="1"`;
  return s + (open.selfClosing ? "/>" : ">");
}

/**
 * The part with new `<definedNames>` content (`inner`, already serialised; empty → the
 * element is removed) and `fullCalcOnLoad="1"` on `<calcPr>` (added when the part has none:
 * F4, Excel recalculates on load and drops the flag when it saves).
 */
export function patchWorkbookXml(xml: string, layout: WorkbookXmlLayout, inner: string): string {
  const p = layout.prefix;
  const dn = layout.definedNames;
  const block = inner === "" ? "" : `<${p}definedNames>${inner}</${p}definedNames>`;
  const edits: { start: number; end: number; text: string }[] = [];
  if (dn && inner !== "" && !dn.selfClosing) edits.push({ start: dn.innerStart, end: dn.innerEnd, text: inner });
  else if (dn) edits.push({ start: dn.start, end: dn.end, text: block });
  if (layout.calcPr) {
    if (!dn) edits.push({ start: layout.insertAt, end: layout.insertAt, text: block });
    edits.push({ start: layout.calcPr.start, end: layout.calcPr.end, text: calcPrTag(layout.calcPr.open) });
  } else {
    // calcPr comes right after definedNames in the schema.
    const calc = `<${p}calcPr fullCalcOnLoad="1"/>`;
    if (dn) edits.push({ start: dn.end, end: dn.end, text: calc });
    else edits.push({ start: layout.insertAt, end: layout.insertAt, text: block + calc });
  }
  edits.sort((a, b) => a.start - b.start || a.end - b.end);
  let out = "";
  let at = 0;
  for (const e of edits) {
    out += xml.slice(at, e.start) + e.text;
    at = e.end;
  }
  return out + xml.slice(at);
}
