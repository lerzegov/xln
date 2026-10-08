// D5, embedded source: the project's source files carried inside the workbook as a custom
// XML part (probe F5: Excel keeps such a part through open and save). This file holds the
// part's format, read and written as text; build/embed.ts puts the part into a package.
// The part is an archive copy for a workbook handed on: pull never reads it (decided
// 2026-10-06, every pull is fresh). `readEmbeddedSource` stays for the build (a part that
// already says the same is left as it is), its read-back, and the tests.
//
//   customXml/itemN.xml       <project xmlns="urn:xln:embedded-source:1" format="xln.embed/1">
//                               <file path="names/FN.xln" sha256="…">text</file> …
//   customXml/itemPropsN.xml  datastore item: the fixed xln itemID, a schemaRef to the namespace
//   customXml/_rels/itemN.xml.rels → itemPropsN.xml
//
// The part is found by the namespace of its root element, never by its number: Excel
// numbers custom XML items itself, and other add-ins keep their own items beside ours.
//
// Line ends are not the part's to keep: Excel re-saves the item with every LF in text
// written as CR LF (measured, probes/README.md § F5), and a conforming XML reader turns
// CR LF back into LF. So a file is stored with its line ends as LF and an `eol` attribute
// (`lf` or `crlf`) says which to restore; on reading, CR LF and lone CR become LF before
// `eol` is applied. A file whose line ends are mixed, or that holds a lone CR or a
// character XML 1.0 cannot carry (a control character), is stored as base64 of its UTF-8
// bytes instead. Text escapes `&`, `<`, `>`. Each file carries the SHA-256 of its UTF-8
// bytes, so a part damaged on the way is reported, not restored.

import { strFromU8 } from "fflate";
import { sha256Bytes, utf8 } from "../project/hash.js";
import { fromBase64, toBase64 } from "../file/base64.js";
import { Package, relTypeIs } from "../file/package.js";
import { childElements, ownText, parseXml } from "../file/xml.js";

/** Namespace of the root element: how the part is recognised. */
export const EMBED_NS = "urn:xln:embedded-source:1";
export const EMBED_FORMAT = "xln.embed/1";
/** The datastore item ID of the xln part. Fixed, so a rebuilt part keeps its identity. */
export const EMBED_ITEM_ID = "{6B1D2E7A-4C3F-4E8B-9A15-7F0C3D2B5E91}";

export const CUSTOM_XML_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
export const CUSTOM_XML_PROPS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
export const CUSTOM_XML_PROPS_CT = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
export const CUSTOM_XML_DS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";

export interface EmbeddedSource {
  /** Project-relative path → text, in the order stored. */
  files: Record<string, string>;
  /** Files whose text did not match their checksum, or could not be decoded: not in `files`. */
  damaged: string[];
  format: string;
}

function isXmlChar(code: number): boolean {
  return code === 0x9 || code === 0xa || code === 0xd || (code >= 0x20 && code <= 0xd7ff) || (code >= 0xe000 && code <= 0xfffd) || code >= 0x10000;
}

function needsBase64(text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    const c = text.codePointAt(i)!;
    if (c > 0xffff) i++;
    else if (c >= 0xd800 && c <= 0xdfff) return true;
    if (!isXmlChar(c)) return true;
  }
  return false;
}

function escapeText(s: string): string {
  let out = "";
  for (const c of s) {
    if (c === "&") out += "&amp;";
    else if (c === "<") out += "&lt;";
    else if (c === ">") out += "&gt;";
    else if (c === "\r") out += "&#13;";
    else out += c;
  }
  return out;
}

function escapeAttr(s: string): string {
  return escapeText(s).split('"').join("&quot;").split("\n").join("&#10;").split("\t").join("&#9;");
}

/** `lf` or `crlf` when every line end of `text` is that (a text with none is `lf`); undefined when mixed or with a lone CR. */
function lineEnds(text: string): "lf" | "crlf" | undefined {
  let lf = 0;
  let crlf = 0;
  for (let i = 0; i < text.length; i++) {
    if (text[i] === "\r") {
      if (text[i + 1] !== "\n") return undefined;
      crlf++;
      i++;
    } else if (text[i] === "\n") lf++;
  }
  if (lf > 0 && crlf > 0) return undefined;
  return crlf > 0 ? "crlf" : "lf";
}

/** Line ends as LF, whatever a writer made of them (CR LF, lone CR). */
function toLf(text: string): string {
  return text.split("\r\n").join("\n").split("\r").join("\n");
}

/** The custom XML item holding `files` (path → text), in the order given. */
export function embeddedSourceXml(files: Readonly<Record<string, string>>): string {
  let s = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<project xmlns="${EMBED_NS}" format="${EMBED_FORMAT}">`;
  for (const [path, text] of Object.entries(files)) {
    const bytes = utf8(text);
    const sum = sha256Bytes(bytes);
    const eol = lineEnds(text);
    if (eol === undefined || needsBase64(text)) s += `<file path="${escapeAttr(path)}" sha256="${sum}" encoding="base64">${toBase64(bytes)}</file>`;
    else s += `<file path="${escapeAttr(path)}" sha256="${sum}" eol="${eol}">${escapeText(toLf(text))}</file>`;
  }
  return s + "</project>";
}

/** The datastore-item properties part that goes with the item. */
export function embeddedSourcePropsXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\r\n' +
    `<ds:datastoreItem ds:itemID="${EMBED_ITEM_ID}" xmlns:ds="${CUSTOM_XML_DS_NS}">` +
    `<ds:schemaRefs><ds:schemaRef ds:uri="${EMBED_NS}"/></ds:schemaRefs></ds:datastoreItem>`
  );
}

/** Whether the text of a custom XML item is the xln part (its root is in the xln namespace). */
export function isEmbeddedSourceXml(xml: string): boolean {
  try {
    const root = parseXml(xml);
    return root.ns === EMBED_NS && root.local === "project";
  } catch {
    return false;
  }
}

/** Reads the files out of the xln part's text; undefined when the text is not one. */
export function parseEmbeddedSource(xml: string): EmbeddedSource | undefined {
  let root;
  try {
    root = parseXml(xml);
  } catch {
    return undefined;
  }
  if (root.ns !== EMBED_NS || root.local !== "project") return undefined;
  const out: EmbeddedSource = { files: {}, damaged: [], format: root.attrs["format"] ?? "" };
  for (const f of childElements(root, "file")) {
    const path = f.attrs["path"];
    if (path === undefined) continue;
    const raw = ownText(f);
    let bytes: Uint8Array | undefined;
    let text: string | undefined;
    if (f.attrs["encoding"] === "base64") {
      bytes = fromBase64(raw);
      text = bytes === undefined ? undefined : strFromU8(bytes);
    } else {
      text = toLf(raw);
      if (f.attrs["eol"] === "crlf") text = text.split("\n").join("\r\n");
      bytes = utf8(text);
    }
    if (bytes === undefined || text === undefined || (f.attrs["sha256"] !== undefined && f.attrs["sha256"] !== sha256Bytes(bytes))) {
      out.damaged.push(path);
      continue;
    }
    out.files[path] = text;
  }
  return out;
}

/** Where the xln part sits in a package. */
export interface EmbeddedPartLocation {
  /** The item part, e.g. `customXml/item1.xml`. */
  item: string;
  /** Its relationship Id in the workbook part's relationships; undefined when the part is not related (found by scanning). */
  relId: string | undefined;
}

function workbookPart(pkg: Package): string | undefined {
  for (const r of pkg.rels("")) if (relTypeIs(r.type, "officeDocument") && !r.external && pkg.has(r.target)) return pkg.find(r.target);
  return pkg.find("xl/workbook.xml");
}

/**
 * The xln part of a package: first among the custom XML items the workbook part relates
 * to, else any `customXml/…` item whose root is in the xln namespace (a part whose
 * relationship was lost still restores).
 */
export function findEmbeddedPart(pkg: Package): EmbeddedPartLocation | undefined {
  const wb = workbookPart(pkg);
  const related = new Set<string>();
  if (wb) {
    for (const r of pkg.rels(wb)) {
      if (r.external || !relTypeIs(r.type, "customXml")) continue;
      const part = pkg.find(r.target);
      if (!part) continue;
      related.add(part);
      const text = pkg.text(part);
      if (text !== undefined && isEmbeddedSourceXml(text)) return { item: part, relId: r.id };
    }
  }
  for (const name of pkg.names) {
    const n = name.toLowerCase();
    if (!n.startsWith("customxml/") || n.includes("/_rels/") || !n.endsWith(".xml") || n.includes("itemprops") || related.has(name)) continue;
    const text = pkg.text(name);
    if (text !== undefined && isEmbeddedSourceXml(text)) return { item: name, relId: undefined };
  }
  return undefined;
}

/** The embedded source of a workbook, with where it was found; undefined when it has none. */
export function readEmbeddedSource(bytes: Uint8Array | Package): (EmbeddedSource & { location: EmbeddedPartLocation }) | undefined {
  const pkg = bytes instanceof Package ? bytes : new Package(bytes);
  const location = findEmbeddedPart(pkg);
  if (!location) return undefined;
  const parsed = parseEmbeddedSource(pkg.text(location.item) ?? "");
  return parsed ? { ...parsed, location } : undefined;
}
