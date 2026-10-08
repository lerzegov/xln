// A small XML pull parser for the parts of an OOXML package.
//
// Why our own: the file layer needs two things generic parsers get in the way of.
// (1) Text content is returned exactly as stored, without the XML end-of-line
// normalisation (CR LF -> LF): Excel keeps CR LF inside definitions and the write path
// must be able to compare against what is on disk. (2) Worksheets can be large, so the
// sheet reader streams tokens and only builds trees for small subtrees.
//
// Scope: well-formed XML 1.0 as Office writes it. Comments, processing instructions and
// a DOCTYPE are skipped (OOXML forbids DTDs, so no custom entities are expanded); CDATA
// is returned as text. Element nesting is checked. Attribute values are normalised as
// XML 1.0 §3.3.3 requires (literal tab, CR, LF become spaces; character references are
// kept), because that is what Excel itself reads.

export class XmlError extends Error {
  constructor(message: string, readonly offset: number) {
    super(`${message} (at offset ${offset})`);
    this.name = "XmlError";
  }
}

/** Offsets of a token in the source text: `src.slice(start, end)` is its markup. The write
 *  path patches a part by splicing at these offsets, so untouched bytes stay as they were. */
export interface XmlSpan {
  start: number;
  end: number;
}

export interface XmlOpen extends XmlSpan {
  type: "open";
  /** Qualified name as written, e.g. `x14:cfRule`. */
  name: string;
  /** Name without its prefix. */
  local: string;
  /** Namespace URI of the element ("" when unbound). */
  ns: string;
  /** Attributes by qualified name, entity-decoded. Namespace declarations included. */
  attrs: Record<string, string>;
  selfClosing: boolean;
  /** Resolves a prefix to its namespace URI in this element's scope. */
  resolve(prefix: string): string | undefined;
}

/** The close token of a self-closing element is empty: `start === end`, just after the tag. */
export interface XmlClose extends XmlSpan {
  type: "close";
  name: string;
  local: string;
}

/** Spans the raw text, entity references and CDATA sections included. */
export interface XmlText extends XmlSpan {
  type: "text";
  text: string;
}

export type XmlToken = XmlOpen | XmlClose | XmlText;

const PREDEFINED: Record<string, string> = { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" };

/** Decodes the five predefined entities and numeric character references. */
export function decodeEntities(s: string, offset = 0): string {
  let amp = s.indexOf("&");
  if (amp < 0) return s;
  let out = "";
  let last = 0;
  while (amp >= 0) {
    const semi = s.indexOf(";", amp + 1);
    if (semi < 0) throw new XmlError("unterminated entity reference", offset + amp);
    const ent = s.slice(amp + 1, semi);
    let rep: string | undefined;
    if (ent.startsWith("#x") || ent.startsWith("#X")) {
      rep = codePoint(ent.slice(2), 16);
    } else if (ent.startsWith("#")) {
      rep = codePoint(ent.slice(1), 10);
    } else {
      rep = PREDEFINED[ent];
    }
    if (rep === undefined) throw new XmlError(`unknown entity &${ent};`, offset + amp);
    out += s.slice(last, amp) + rep;
    last = semi + 1;
    amp = s.indexOf("&", last);
  }
  return out + s.slice(last);
}

function codePoint(digits: string, radix: number): string | undefined {
  if (digits.length === 0) return undefined;
  for (const ch of digits) if (Number.isNaN(parseInt(ch, radix))) return undefined;
  const n = parseInt(digits, radix);
  if (n > 0x10ffff) return undefined;
  return String.fromCodePoint(n);
}

function isSpace(c: number): boolean {
  return c === 0x20 || c === 0x09 || c === 0x0a || c === 0x0d;
}

function normaliseAttr(raw: string): string {
  // Literal whitespace only; references are decoded afterwards so &#10; survives.
  if (!raw.includes("\r") && !raw.includes("\n") && !raw.includes("\t")) return raw;
  let out = "";
  for (let i = 0; i < raw.length; i++) {
    const c = raw[i]!;
    if (c === "\r") {
      out += " ";
      if (raw[i + 1] === "\n") i++;
    } else if (c === "\n" || c === "\t") out += " ";
    else out += c;
  }
  return out;
}

const XML_NS = "http://www.w3.org/XML/1998/namespace";

interface Scope {
  map: Record<string, string>;
  parent: Scope | undefined;
}

function lookup(scope: Scope | undefined, prefix: string): string | undefined {
  for (let s = scope; s; s = s.parent) {
    const uri = s.map[prefix];
    if (uri !== undefined) return uri;
  }
  if (prefix === "xml") return XML_NS;
  return undefined;
}

function localOf(qname: string): string {
  const c = qname.indexOf(":");
  return c < 0 ? qname : qname.slice(c + 1);
}

function prefixOf(qname: string): string {
  const c = qname.indexOf(":");
  return c < 0 ? "" : qname.slice(0, c);
}

/** Streams tokens from an XML document. Adjacent text and CDATA are merged. */
export class XmlReader {
  private pos = 0;
  private readonly stack: { name: string; scope: Scope | undefined }[] = [];
  private pendingClose: XmlClose | undefined;
  private seenRoot = false;

  constructor(private readonly src: string) {
    if (src.charCodeAt(0) === 0xfeff) this.pos = 1;
  }

  /** Current nesting depth (number of open elements). */
  get depth(): number {
    return this.stack.length;
  }

  next(): XmlToken | undefined {
    if (this.pendingClose) {
      const t = this.pendingClose;
      this.pendingClose = undefined;
      return t;
    }
    const src = this.src;
    let text = "";
    let hasText = false;
    const textStart = this.pos;
    while (this.pos < src.length) {
      const lt = src.indexOf("<", this.pos);
      if (lt < 0) {
        const at = this.pos;
        const rest = src.slice(at);
        this.pos = src.length;
        if (this.stack.length > 0) {
          text += decodeEntities(rest, at);
          hasText = true;
        } else if (rest.trim() !== "") {
          throw new XmlError("text outside the root element", at);
        }
        break;
      }
      if (lt > this.pos) {
        const raw = src.slice(this.pos, lt);
        if (this.stack.length > 0) {
          text += decodeEntities(raw, this.pos);
          hasText = true;
        } else if (raw.trim() !== "") {
          throw new XmlError("text outside the root element", this.pos);
        }
        this.pos = lt;
      }
      // At '<'.
      if (src.startsWith("<![CDATA[", lt)) {
        const end = src.indexOf("]]>", lt + 9);
        if (end < 0) throw new XmlError("unterminated CDATA section", lt);
        text += src.slice(lt + 9, end);
        hasText = true;
        this.pos = end + 3;
        continue;
      }
      if (src.startsWith("<!--", lt)) {
        const end = src.indexOf("-->", lt + 4);
        if (end < 0) throw new XmlError("unterminated comment", lt);
        this.pos = end + 3;
        continue;
      }
      if (src.startsWith("<?", lt)) {
        const end = src.indexOf("?>", lt + 2);
        if (end < 0) throw new XmlError("unterminated processing instruction", lt);
        this.pos = end + 2;
        continue;
      }
      if (src.startsWith("<!DOCTYPE", lt)) {
        this.pos = this.skipDoctype(lt);
        continue;
      }
      // A markup token: flush text first, leave the tag for the next call.
      if (hasText && text !== "") return { type: "text", text, start: textStart, end: lt };
      text = "";
      hasText = false;
      return src.charCodeAt(lt + 1) === 0x2f ? this.readClose(lt) : this.readOpen(lt);
    }
    if (hasText && text !== "") return { type: "text", text, start: textStart, end: this.pos };
    if (this.stack.length > 0) {
      throw new XmlError(`unexpected end of document inside <${this.stack[this.stack.length - 1]!.name}>`, src.length);
    }
    if (!this.seenRoot) throw new XmlError("no root element", 0);
    return undefined;
  }

  /** Skips to the close tag matching an `open` token just returned, if it was not self-closing. */
  skip(open: XmlOpen): void {
    if (open.selfClosing) {
      this.next(); // the synthetic close
      return;
    }
    const target = this.stack.length - 1;
    for (;;) {
      const t = this.next();
      if (!t) throw new XmlError(`unexpected end of document inside <${open.name}>`, this.pos);
      if (t.type === "close" && this.stack.length === target) return;
    }
  }

  private skipDoctype(lt: number): number {
    let depth = 0;
    for (let i = lt + 9; i < this.src.length; i++) {
      const c = this.src[i];
      if (c === "[") depth++;
      else if (c === "]") depth--;
      else if (c === ">" && depth <= 0) return i + 1;
    }
    throw new XmlError("unterminated DOCTYPE", lt);
  }

  private readName(i: number): number {
    const src = this.src;
    let j = i;
    while (j < src.length) {
      const c = src.charCodeAt(j);
      if (isSpace(c) || c === 0x3e /* > */ || c === 0x2f /* / */ || c === 0x3d /* = */) break;
      j++;
    }
    if (j === i) throw new XmlError("expected a name", i);
    return j;
  }

  private readClose(lt: number): XmlClose {
    const src = this.src;
    const end = this.readName(lt + 2);
    const name = src.slice(lt + 2, end);
    let i = end;
    while (i < src.length && isSpace(src.charCodeAt(i))) i++;
    if (src[i] !== ">") throw new XmlError(`malformed close tag </${name}`, lt);
    this.pos = i + 1;
    const top = this.stack.pop();
    if (!top) throw new XmlError(`close tag </${name}> without an open element`, lt);
    if (top.name !== name) throw new XmlError(`</${name}> does not close <${top.name}>`, lt);
    return { type: "close", name, local: localOf(name), start: lt, end: i + 1 };
  }

  private readOpen(lt: number): XmlOpen {
    const src = this.src;
    if (this.stack.length === 0 && this.seenRoot) throw new XmlError("more than one root element", lt);
    const nameEnd = this.readName(lt + 1);
    const name = src.slice(lt + 1, nameEnd);
    const attrs: Record<string, string> = {};
    let decls: Record<string, string> | undefined;
    let i = nameEnd;
    let selfClosing = false;
    for (;;) {
      while (i < src.length && isSpace(src.charCodeAt(i))) i++;
      if (i >= src.length) throw new XmlError(`unterminated tag <${name}`, lt);
      const c = src.charCodeAt(i);
      if (c === 0x3e) {
        i++;
        break;
      }
      if (c === 0x2f) {
        if (src[i + 1] !== ">") throw new XmlError(`malformed tag <${name}`, i);
        selfClosing = true;
        i += 2;
        break;
      }
      const an = this.readName(i);
      const aname = src.slice(i, an);
      i = an;
      while (i < src.length && isSpace(src.charCodeAt(i))) i++;
      if (src[i] !== "=") throw new XmlError(`attribute ${aname} without a value`, i);
      i++;
      while (i < src.length && isSpace(src.charCodeAt(i))) i++;
      const q = src[i];
      if (q !== '"' && q !== "'") throw new XmlError(`attribute ${aname} value not quoted`, i);
      const close = src.indexOf(q, i + 1);
      if (close < 0) throw new XmlError(`unterminated value of attribute ${aname}`, i);
      const raw = src.slice(i + 1, close);
      if (raw.includes("<")) throw new XmlError(`'<' in value of attribute ${aname}`, i);
      const value = decodeEntities(normaliseAttr(raw), i + 1);
      i = close + 1;
      if (aname === "__proto__") continue;
      if (Object.prototype.hasOwnProperty.call(attrs, aname)) {
        throw new XmlError(`duplicate attribute ${aname}`, i);
      }
      attrs[aname] = value;
      if (aname === "xmlns") (decls ??= {})[""] = value;
      else if (aname.startsWith("xmlns:")) (decls ??= {})[aname.slice(6)] = value;
    }
    this.pos = i;
    this.seenRoot = true;
    const parentScope = this.stack.length ? this.stack[this.stack.length - 1]!.scope : undefined;
    const scope: Scope | undefined = decls ? { map: decls, parent: parentScope } : parentScope;
    const prefix = prefixOf(name);
    const ns = lookup(scope, prefix) ?? "";
    if (prefix !== "" && ns === "") throw new XmlError(`unbound namespace prefix ${prefix}:`, lt);
    if (selfClosing) {
      this.pendingClose = { type: "close", name, local: localOf(name), start: i, end: i };
    } else {
      this.stack.push({ name, scope });
    }
    return {
      type: "open",
      start: lt,
      end: i,
      name,
      local: localOf(name),
      ns,
      attrs,
      selfClosing,
      resolve: (p: string) => lookup(scope, p),
    };
  }
}

/** Value of the attribute with namespace `uri` and local name `local`, if present. */
export function attrNS(open: XmlOpen | XmlElement, uri: readonly string[], local: string): string | undefined {
  for (const [qname, value] of Object.entries(open.attrs)) {
    const c = qname.indexOf(":");
    if (c < 0 || qname.slice(c + 1) !== local) continue;
    const p = qname.slice(0, c);
    if (p === "xmlns") continue;
    const u = open.resolve(p);
    if (u !== undefined && uri.includes(u)) return value;
  }
  return undefined;
}

// ---------------------------------------------------------------------------------
// A minimal element tree, for the small parts (workbook, rels, tables, metadata) and
// for subtrees of a worksheet.

export interface XmlElement {
  name: string;
  local: string;
  ns: string;
  attrs: Record<string, string>;
  children: (XmlElement | string)[];
  resolve(prefix: string): string | undefined;
}

/** Builds the subtree rooted at `open` (just returned by `reader`). */
export function readElement(reader: XmlReader, open: XmlOpen): XmlElement {
  const root: XmlElement = {
    name: open.name,
    local: open.local,
    ns: open.ns,
    attrs: open.attrs,
    children: [],
    resolve: open.resolve,
  };
  const stack: XmlElement[] = [root];
  if (open.selfClosing) {
    reader.next(); // the synthetic close
    return root;
  }
  for (;;) {
    const t = reader.next();
    if (!t) throw new XmlError(`unexpected end of document inside <${open.name}>`, -1);
    const top = stack[stack.length - 1]!;
    if (t.type === "text") {
      top.children.push(t.text);
    } else if (t.type === "open") {
      const el: XmlElement = { name: t.name, local: t.local, ns: t.ns, attrs: t.attrs, children: [], resolve: t.resolve };
      top.children.push(el);
      stack.push(el);
    } else {
      stack.pop();
      if (stack.length === 0) return root;
    }
  }
}

/** Parses a whole document into a tree and returns its root element. */
export function parseXml(src: string): XmlElement {
  const reader = new XmlReader(src);
  let root: XmlElement | undefined;
  for (let t = reader.next(); t; t = reader.next()) {
    if (t.type === "open") root = readElement(reader, t);
  }
  if (!root) throw new XmlError("no root element", 0);
  return root;
}

export function childElements(el: XmlElement, local?: string): XmlElement[] {
  const out: XmlElement[] = [];
  for (const c of el.children) {
    if (typeof c !== "string" && (local === undefined || c.local === local)) out.push(c);
  }
  return out;
}

export function firstChild(el: XmlElement, local: string): XmlElement | undefined {
  for (const c of el.children) if (typeof c !== "string" && c.local === local) return c;
  return undefined;
}

/** Concatenated text of the element's direct text children (no descendants). */
export function ownText(el: XmlElement): string {
  let s = "";
  for (const c of el.children) if (typeof c === "string") s += c;
  return s;
}

/** Depth-first walk over descendants (excluding `el` itself). */
export function* descendants(el: XmlElement): Generator<XmlElement> {
  for (const c of el.children) {
    if (typeof c === "string") continue;
    yield c;
    yield* descendants(c);
  }
}
