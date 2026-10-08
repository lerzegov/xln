// Where a part keeps formula text, found with the XML tokenizer (probe F7's first run
// corrupted a sheet with a regular expression over `<f>`): the definitions in
// `<definedNames>` of the workbook part; in a worksheet, cell formulas (`<f>`, a shared
// group's text at its master only: the members' `<f t="shared" si=…/>` carry none),
// conditional formats (`<formula>`), validations (`<formula1>`, `<formula2>`) and the x14
// extension's `<xm:f>` (conditional formats, validations, sparklines). Each with the
// offsets of its raw text, so a rename splices new text in and leaves every other byte.

import { XmlReader, type XmlOpen } from "../file/xml.js";

export type FormulaPlace = "cell" | "format" | "validation" | "name" | "other";

export interface FormulaTextNode {
  /** Offsets of the raw text (entity references included) between the element's tags. */
  start: number;
  end: number;
  /** The text, entities decoded, line breaks as stored. */
  text: string;
  place: FormulaPlace;
  /** The element holding the text (`<f>`, `<definedName>`, …). */
  open: XmlOpen;
  /** For a cell: its `<c>` element's `r`. */
  ref?: string;
}

const SHEET_FORMULAS = new Set(["f", "formula", "formula1", "formula2"]);

/** The formula texts of a workbook part (`definedName`) or a worksheet part. */
export function formulaTextNodes(xml: string, part: "workbook" | "sheet"): FormulaTextNode[] {
  const r = new XmlReader(xml);
  const stack: XmlOpen[] = [];
  const out: FormulaTextNode[] = [];
  let cur: { open: XmlOpen; depth: number; start: number; end: number; text: string; place: FormulaPlace; ref?: string } | undefined;
  for (let t = r.next(); t; t = r.next()) {
    if (t.type === "open") {
      stack.push(t);
      const parent = stack[stack.length - 2];
      if (cur) continue; // nothing nests inside a formula's text
      let place: FormulaPlace | undefined;
      let ref: string | undefined;
      if (part === "workbook") {
        if (t.local === "definedName" && parent?.local === "definedNames") place = "name";
      } else if (SHEET_FORMULAS.has(t.local)) {
        const c = stack.find((s) => s.local === "c");
        if (c) {
          place = "cell";
          ref = c.attrs["r"];
        } else if (stack.some((s) => s.local === "conditionalFormatting")) place = "format";
        else if (stack.some((s) => s.local === "dataValidation")) place = "validation";
        else place = "other";
      }
      if (place !== undefined && !t.selfClosing) cur = { open: t, depth: stack.length, start: -1, end: -1, text: "", place, ...(ref !== undefined ? { ref } : {}) };
    } else if (t.type === "text") {
      if (cur && stack.length === cur.depth) {
        if (cur.start < 0) cur.start = t.start;
        cur.end = t.end;
        cur.text += t.text;
      }
    } else {
      if (cur && stack.length === cur.depth) {
        if (cur.start >= 0) out.push({ start: cur.start, end: cur.end, text: cur.text, place: cur.place, open: cur.open, ...(cur.ref !== undefined ? { ref: cur.ref } : {}) });
        cur = undefined;
      }
      stack.pop();
    }
  }
  return out;
}

/** The part with every formula text cut out: two parts that differ only in formula texts give the same. */
export function withoutFormulaTexts(xml: string, part: "workbook" | "sheet"): string {
  let out = "";
  let at = 0;
  for (const n of formulaTextNodes(xml, part)) {
    out += xml.slice(at, n.start);
    at = n.end;
  }
  return out + xml.slice(at);
}
