// Test-only helper: pulls defined-name definitions and cell formula texts out of an .xlsx,
// so the language tests do not depend on the file layer being built in parallel (W1).
// Node is allowed here; packages/core/src is not.
import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";

export interface DefinedNameText {
  name: string;
  localSheetId?: number;
  text: string;
}
export interface CellFormulaText {
  part: string;
  cell: string;
  text: string;
}
export interface WorkbookStrings {
  sheets: string[];
  names: DefinedNameText[];
  formulas: CellFormulaText[];
}

function unzip(bytes: Uint8Array): Map<string, Uint8Array> {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let p = bytes.length - 22; p >= 0; p--) {
    if (dv.getUint32(p, true) === 0x06054b50) {
      eocd = p;
      break;
    }
  }
  if (eocd < 0) throw new Error("not a zip");
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const out = new Map<string, Uint8Array>();
  const dec = new TextDecoder();
  for (let k = 0; k < count; k++) {
    const method = dv.getUint16(p + 10, true);
    const csize = dv.getUint32(p + 20, true);
    const nlen = dv.getUint16(p + 28, true);
    const elen = dv.getUint16(p + 30, true);
    const clen = dv.getUint16(p + 32, true);
    const local = dv.getUint32(p + 42, true);
    const name = dec.decode(bytes.subarray(p + 46, p + 46 + nlen));
    const lnlen = dv.getUint16(local + 26, true);
    const lelen = dv.getUint16(local + 28, true);
    const data = bytes.subarray(local + 30 + lnlen + lelen, local + 30 + lnlen + lelen + csize);
    out.set(name, method === 0 ? data : new Uint8Array(inflateRawSync(data)));
    p += 46 + nlen + elen + clen;
  }
  return out;
}

function decodeXml(s: string): string {
  return s.replace(/\r\n/g, "\n").replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/g, (_, e: string) => {
    if (e === "amp") return "&";
    if (e === "lt") return "<";
    if (e === "gt") return ">";
    if (e === "quot") return '"';
    if (e === "apos") return "'";
    return String.fromCodePoint(e[1] === "x" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
  });
}

interface Tag {
  name: string;
  attrs: Record<string, string>;
  selfClosing: boolean;
  closing: boolean;
  end: number;
}

/** A small tag scanner: handles quoted attributes and self-closing tags (the F7 lesson). */
function* tags(xml: string): Generator<Tag> {
  let p = 0;
  while ((p = xml.indexOf("<", p)) >= 0) {
    if (xml.startsWith("<?", p) || xml.startsWith("<!", p)) {
      p = xml.indexOf(">", p) + 1;
      continue;
    }
    let q = p + 1;
    const closing = xml[q] === "/";
    if (closing) q++;
    const n0 = q;
    while (q < xml.length && !" \t\r\n/>".includes(xml[q]!)) q++;
    const name = xml.slice(n0, q);
    const attrs: Record<string, string> = {};
    for (;;) {
      while (" \t\r\n".includes(xml[q]!)) q++;
      if (xml[q] === ">" || xml[q] === "/") break;
      const a0 = q;
      while (xml[q] !== "=") q++;
      const an = xml.slice(a0, q).trim();
      q++;
      while (" \t\r\n".includes(xml[q]!)) q++;
      const quote = xml[q]!;
      const v0 = q + 1;
      q = xml.indexOf(quote, v0);
      attrs[an.replace(/^.*:/, "")] = decodeXml(xml.slice(v0, q));
      q++;
    }
    const selfClosing = xml[q] === "/";
    const end = xml.indexOf(">", q) + 1;
    yield { name: name.replace(/^.*:/, ""), attrs, selfClosing, closing, end };
    p = end;
  }
}

function elementTexts(xml: string, element: string): { attrs: Record<string, string>; text: string }[] {
  const out: { attrs: Record<string, string>; text: string }[] = [];
  for (const t of tags(xml)) {
    if (t.name !== element || t.closing) continue;
    if (t.selfClosing) {
      out.push({ attrs: t.attrs, text: "" });
      continue;
    }
    const close = xml.indexOf("</", t.end);
    out.push({ attrs: t.attrs, text: decodeXml(xml.slice(t.end, close)) });
  }
  return out;
}

export function readWorkbookStrings(path: string): WorkbookStrings {
  const zip = unzip(new Uint8Array(readFileSync(path)));
  const dec = new TextDecoder();
  const wb = dec.decode(zip.get("xl/workbook.xml"));
  const sheets = elementTexts(wb, "sheet").map((s) => s.attrs.name ?? "");
  const names = elementTexts(wb, "definedName").map((d) => ({
    name: d.attrs.name ?? "",
    ...(d.attrs.localSheetId !== undefined ? { localSheetId: Number(d.attrs.localSheetId) } : {}),
    text: d.text,
  }));
  const formulas: CellFormulaText[] = [];
  for (const [part, data] of zip) {
    if (!part.startsWith("xl/worksheets/") || !part.endsWith(".xml")) continue;
    const xml = dec.decode(data);
    // Pair each <f> with the r= of its enclosing <c>.
    let cell = "";
    for (const t of tags(xml)) {
      if (t.name === "c" && !t.closing) cell = t.attrs.r ?? "";
      if (t.name === "f" && !t.closing && !t.selfClosing) {
        const close = xml.indexOf("</", t.end);
        const text = decodeXml(xml.slice(t.end, close));
        if (text) formulas.push({ part, cell, text });
      }
    }
  }
  return { sheets, names, formulas };
}
