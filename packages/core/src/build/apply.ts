// The file backend: applies a change set to the bytes of a workbook. Name changes touch
// only the workbook part (its `<definedNames>` and `fullCalcOnLoad`). Cell changes (M3b)
// also patch the sheet parts they name, ensure the dynamic-array record in
// `xl/metadata.xml` (creating the part, its relationship and content type if needed) and
// drop `xl/calcChain.xml` (probe F8). Every other zip entry is copied byte for byte
// (zip.ts), and an element no change touches keeps its exact markup, so a build with
// nothing to do leaves `<definedNames>` byte-identical.

import { decodeText, Package, relTypeIs, XlsxError } from "../file/package.js";
import { REL_NS } from "../file/worksheet.js";
import { attrNS, childElements, firstChild, parseXml, XmlError } from "../file/xml.js";
import { rewriteZip } from "../file/zip.js";
import { decodeXstring } from "../file/xstring.js";
import { utf8 } from "../project/hash.js";
import { RenameContext, renameInFormula, type ScopedName } from "../project/rename.js";
import { orderChanges, scopedKey, type Change, type ChangeSet, type ClearCellFormula, type RenameName, type RenameReferences, type Scope, type SetCellFormula, type SetEmbeddedSource } from "./changes.js";
import { formulaTextNodes, type FormulaPlace, type FormulaTextNode } from "./formulaText.js";
import { applyEmbeddedSource, type EmbedOutcome } from "./embed.js";
import { addOverride, addRelationship, ensureDynamicArrayCm, newMetadataXml, removeOverride, removeRelationships, SHEET_METADATA_CT, SHEET_METADATA_REL } from "./packageXml.js";
import { parseRange, patchSheetXml, SheetPatchError, type SheetCellOp, type SheetPatchReport } from "./sheetXml.js";
import { definedNameXml, escapeXmlText, patchWorkbookXml, scanWorkbookXml, type DefinedNameElement } from "./workbookXml.js";

export class ApplyError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ApplyError";
  }
}

const CORE = new Set(["name", "comment", "localSheetId", "hidden"]);

interface Slot {
  name: string;
  /** Sheet name, null for workbook scope; undefined when `localSheetId` names no sheet. */
  scope: Scope | undefined;
  hidden: boolean;
  comment: string | undefined;
  definition: string;
  other: Record<string, string>;
  /** The original markup, while nothing has changed the element. */
  raw: string | undefined;
  /** Renamed, moved to another scope or created: placed again in Excel's sort order. */
  placed: boolean;
}

function isTrue(v: string | undefined): boolean {
  return v === "1" || v === "true";
}

function sameScope(a: Scope | undefined, b: Scope): boolean {
  if (a === undefined) return false;
  return a === null || b === null ? a === b : a.toLowerCase() === b.toLowerCase();
}

/** Excel's order in `<definedNames>`: by name, then by scope sheet name, ignoring case; workbook scope
 *  after the sheet scopes (f7_base.xlsx: `Loc` on S2, then the workbook's `Loc`). */
function sortKey(s: Slot): [string, string] {
  return [s.name.toLowerCase(), s.scope === null || s.scope === undefined ? "\uffff" : s.scope.toLowerCase()];
}

function before(a: Slot, b: Slot): boolean {
  const [an, as] = sortKey(a);
  const [bn, bs] = sortKey(b);
  return an < bn || (an === bn && as < bs);
}

/** The workbook part's path, from the package relationships. */
export function workbookPartOf(pkg: Package): string {
  for (const r of pkg.rels("")) if (relTypeIs(r.type, "officeDocument") && !r.external && pkg.has(r.target)) return pkg.find(r.target)!;
  const fallback = pkg.find("xl/workbook.xml");
  if (fallback) return fallback;
  throw new XlsxError("no workbook part: not an .xlsx package");
}

/** Applies `changes` to the text of the workbook part. */
export function applyToWorkbookXml(xml: string, changes: readonly Change[]): string {
  const layout = scanWorkbookXml(xml);
  const sheetIndex = (scope: Scope): number | undefined => {
    if (scope === null) return undefined;
    const k = layout.sheets.findIndex((s) => s.toLowerCase() === scope.toLowerCase());
    if (k < 0) throw new ApplyError(`no sheet '${scope}' in the workbook`);
    return k;
  };
  const sheetName = (scope: Scope): Scope => (scope === null ? null : layout.sheets[sheetIndex(scope)!]!);

  const slots: Slot[] = (layout.definedNames?.items ?? []).map((e: DefinedNameElement) => {
    const lsid = e.attrs["localSheetId"];
    const pos = lsid === undefined ? undefined : Number(lsid);
    const other: Record<string, string> = {};
    for (const [k, v] of Object.entries(e.attrs)) if (!CORE.has(k)) other[k] = v;
    return {
      name: e.attrs["name"] ?? "",
      scope: pos === undefined ? null : Number.isInteger(pos) && layout.sheets[pos] !== undefined ? layout.sheets[pos]! : undefined,
      hidden: isTrue(e.attrs["hidden"]),
      comment: e.attrs["comment"] === undefined ? undefined : decodeXstring(e.attrs["comment"]),
      definition: e.text,
      other,
      raw: xml.slice(e.start, e.end),
      placed: false,
    };
  });

  const find = (name: string, scope: Scope): Slot | undefined => {
    const n = name.toLowerCase();
    return slots.find((s) => s.name.toLowerCase() === n && sameScope(s.scope, scope));
  };
  const must = (name: string, scope: Scope, what: string): Slot => {
    const s = find(name, scope);
    if (!s) throw new ApplyError(`${what}: the workbook has no name ${scopedKey(name, scope)}`);
    if (s.name.toLowerCase().startsWith("_xl")) throw new ApplyError(`${what}: ${s.name} is Excel's own name`);
    return s;
  };
  const free = (name: string, scope: Scope, what: string): void => {
    if (find(name, scope)) throw new ApplyError(`${what}: ${scopedKey(name, scope)} exists already`);
  };

  for (const c of orderChanges(changes)) {
    switch (c.op) {
      case "rename-name": {
        const s = must(c.from, c.scope, "rename");
        if (c.from.toLowerCase() !== c.to.toLowerCase()) free(c.to, c.scope, "rename");
        Object.assign(s, { name: c.to, raw: undefined, placed: true });
        break;
      }
      case "rescope-name": {
        const s = must(c.name, c.from, "scope change");
        free(c.name, c.to, "scope change");
        Object.assign(s, { scope: sheetName(c.to), raw: undefined, placed: true });
        break;
      }
      case "delete-name":
        slots.splice(slots.indexOf(must(c.name, c.scope, "delete")), 1);
        break;
      case "set-name": {
        const s = find(c.name, c.scope);
        const comment = c.comment ?? undefined;
        if (!s) {
          sheetIndex(c.scope);
          slots.push({ name: c.name, scope: sheetName(c.scope), hidden: c.hidden, comment, definition: c.stored, other: {}, raw: undefined, placed: true });
          break;
        }
        if (s.name.toLowerCase().startsWith("_xl")) throw new ApplyError(`set: ${s.name} is Excel's own name`);
        const same = s.name === c.name && s.hidden === c.hidden && (s.comment ?? "") === (comment ?? "") && s.definition === c.stored;
        if (!same) Object.assign(s, { name: c.name, hidden: c.hidden, comment, definition: c.stored, raw: undefined });
        break;
      }
      case "set-cell-formula":
      case "clear-cell-formula":
        break; // the sheet parts' business (applyCellChanges)
      case "set-embedded-source":
        break; // a package part of its own (embed.ts)
    }
  }

  // Names nothing moved keep their order; the others go where Excel would sort them.
  const ordered = slots.filter((s) => !s.placed);
  for (const s of slots.filter((x) => x.placed).sort((a, b) => (before(a, b) ? -1 : before(b, a) ? 1 : 0))) {
    let k = ordered.findIndex((o) => before(s, o));
    if (k < 0) k = ordered.length;
    ordered.splice(k, 0, s);
  }
  const inner = ordered
    .map(
      (s) =>
        s.raw ??
        definedNameXml(
          { name: s.name, comment: s.comment, localSheetId: s.scope === null || s.scope === undefined ? undefined : sheetIndex(s.scope), hidden: s.hidden, other: s.other, definition: s.definition },
          layout.prefix,
        ),
    )
    .join("");
  return patchWorkbookXml(xml, layout, inner);
}

/** Text encoded as the part it replaces was: UTF-8, with its BOM if it had one. */
function encodeLike(before: Uint8Array, text: string, part: string): Uint8Array {
  if (before.length >= 2 && (before[0] === 0xff || before[0] === 0xfe)) throw new XlsxError(`${part} is UTF-16: not supported for writing`);
  const bom = before[0] === 0xef && before[1] === 0xbb && before[2] === 0xbf;
  const encoded = utf8(text);
  if (!bom) return encoded;
  const out = new Uint8Array(3 + encoded.length);
  out.set([0xef, 0xbb, 0xbf]);
  out.set(encoded, 3);
  return out;
}

/** The `_rels` part of a part: `xl/workbook.xml` → `xl/_rels/workbook.xml.rels`. */
export function relsPartOf(part: string): string {
  const slash = part.lastIndexOf("/");
  return (slash < 0 ? "" : part.slice(0, slash + 1)) + "_rels/" + part.slice(slash + 1) + ".rels";
}

/** A relationship target for `part`, relative to the directory of `from`. */
function relativeTarget(from: string, part: string): string {
  const dir = from.slice(0, from.lastIndexOf("/") + 1);
  return dir !== "" && part.startsWith(dir) ? part.slice(dir.length) : "/" + part;
}

/** Worksheet parts by sheet name (lower case); `part` is undefined for a sheet that is not a worksheet. */
export function sheetPartsOf(pkg: Package, workbookPart: string): Map<string, { name: string; part: string | undefined }> {
  const wb = parseXml(pkg.text(workbookPart) ?? "");
  const rels = new Map(pkg.rels(workbookPart).map((r) => [r.id, r]));
  const out = new Map<string, { name: string; part: string | undefined }>();
  const sheets = firstChild(wb, "sheets");
  for (const s of sheets ? childElements(sheets, "sheet") : []) {
    const name = s.attrs["name"] ?? "";
    const rel = rels.get(attrNS(s, REL_NS, "id") ?? "");
    const part = rel && !rel.external && relTypeIs(rel.type, "worksheet") ? pkg.find(rel.target) : undefined;
    out.set(name.toLowerCase(), { name, part });
  }
  return out;
}

export interface ApplyReport {
  /** Per sheet name, what the cell patch did. */
  sheets: Record<string, SheetPatchReport>;
  /** The dynamic-array record: found in `xl/metadata.xml`, added to it, or a new part. */
  metadata: "created" | "extended" | "kept" | "none";
  /** `cm` of the written formulas. */
  cm?: number;
  /** The calculation chain dropped, if there was one. */
  calcChainDropped?: string;
  /** Stretch G: the references the renames rewrote. */
  references?: ReferenceRenameReport;
  /** D5: the embedded source part, and whether it was added, rewritten or already up to date. */
  embedded?: { item: string; outcome: EmbedOutcome };
}

interface PackageEdits {
  replace: Map<string, Uint8Array>;
  remove: Set<string>;
  add: Map<string, Uint8Array>;
}

/**
 * F8 case 6: after a cell formula changes, a stale calculation chain makes Excel repair the
 * file, even pruned; Excel rebuilds it. Takes the chain out of the workbook's relationships
 * (`rels`) and the content types (`ct`), and adds the part to `remove`.
 */
function dropCalcChain(pkg: Package, wbPart: string, rels: string, ct: string, remove: Set<string>): { rels: string; ct: string; part?: string } {
  const chain = pkg.rels(wbPart).filter((r) => relTypeIs(r.type, "calcChain") && !r.external);
  const out: { rels: string; ct: string; part?: string } = { rels: removeRelationships(rels, "calcChain").xml, ct };
  for (const r of chain) {
    const part = pkg.find(r.target);
    out.ct = removeOverride(out.ct, "/" + (part ?? r.target));
    if (part) {
      remove.add(part);
      out.part = part;
    }
  }
  return out;
}

/** Cell changes: sheet parts patched, the dynamic-array record ensured, the calculation chain dropped. */
function applyCellChanges(pkg: Package, wbPart: string, changes: readonly Change[], edits: PackageEdits): ApplyReport {
  const report: ApplyReport = { sheets: {}, metadata: "none" };
  const cellChanges = orderChanges(changes).filter((c): c is SetCellFormula | ClearCellFormula => c.op === "set-cell-formula" || c.op === "clear-cell-formula");
  if (cellChanges.length === 0) return report;

  const sheetParts = sheetPartsOf(pkg, wbPart);
  const bySheet = new Map<string, { name: string; part: string; ops: SheetCellOp[] }>();
  for (const c of cellChanges) {
    const s = sheetParts.get(c.sheet.toLowerCase());
    if (!s) throw new ApplyError(`${c.op} ${c.sheet}!${c.range}: no sheet '${c.sheet}' in the workbook`);
    if (!s.part) throw new ApplyError(`${c.op} ${c.sheet}!${c.range}: '${s.name}' is not a worksheet`);
    if (!parseRange(c.range)) throw new ApplyError(`${c.op} ${c.sheet}!${c.range}: '${c.range}' is not a cell or a range of cells`);
    let entry = bySheet.get(s.part);
    if (!entry) bySheet.set(s.part, (entry = { name: s.name, part: s.part, ops: [] }));
    entry.ops.push(c.op === "set-cell-formula" ? { op: "set", range: c.range, stored: c.stored } : { op: "clear", range: c.range });
  }

  const relsName = pkg.find(relsPartOf(wbPart));
  if (!relsName) throw new ApplyError(`${relsPartOf(wbPart)} is missing`);
  let rels = decodeText(pkg.raw(relsName)!);
  const ctName = pkg.find("[Content_Types].xml");
  if (!ctName) throw new ApplyError("[Content_Types].xml is missing");
  let ct = decodeText(pkg.raw(ctName)!);

  // The dynamic-array record (F8 case 7): found, added to the part, or a new part.
  const mdRel = pkg.rels(wbPart).find((r) => relTypeIs(r.type, "sheetMetadata") && !r.external);
  const mdPart = mdRel ? pkg.find(mdRel.target) : undefined;
  let cm: number;
  if (mdPart) {
    const before = pkg.raw(mdPart)!;
    const md = ensureDynamicArrayCm(decodeText(before));
    cm = md.cm;
    if (md.changed) edits.replace.set(mdPart, encodeLike(before, md.xml, mdPart));
    report.metadata = md.changed ? "extended" : "kept";
  } else {
    const dir = wbPart.slice(0, wbPart.lastIndexOf("/") + 1);
    let part = `${dir}metadata.xml`;
    for (let n = 2; pkg.has(part); n++) part = `${dir}metadata${n}.xml`;
    edits.add.set(part, utf8(newMetadataXml()));
    rels = addRelationship(rels, SHEET_METADATA_REL, relativeTarget(wbPart, part)).xml;
    ct = addOverride(ct, "/" + part, SHEET_METADATA_CT);
    cm = 1;
    report.metadata = "created";
  }
  report.cm = cm;

  for (const s of bySheet.values()) {
    const before = pkg.raw(s.part);
    if (!before) throw new ApplyError(`sheet part ${s.part} is missing`);
    let patched;
    try {
      patched = patchSheetXml(decodeText(before), s.ops, cm);
    } catch (e) {
      if (e instanceof SheetPatchError || e instanceof XmlError) throw new ApplyError(`sheet ${s.name}: ${e.message}`);
      throw e;
    }
    edits.replace.set(s.part, encodeLike(before, patched.xml, s.part));
    report.sheets[s.name] = patched.report;
  }

  const dropped = dropCalcChain(pkg, wbPart, rels, ct, edits.remove);
  rels = dropped.rels;
  ct = dropped.ct;
  if (dropped.part) report.calcChainDropped = dropped.part;
  const relsBefore = pkg.raw(relsName)!;
  if (rels !== decodeText(relsBefore)) edits.replace.set(relsName, encodeLike(relsBefore, rels, relsName));
  const ctBefore = pkg.raw(ctName)!;
  if (ct !== decodeText(ctBefore)) edits.replace.set(ctName, encodeLike(ctBefore, ct, ctName));
  return report;
}

/** What the reference rewrite of a change set's renames did (stretch G). */
export interface ReferenceRenameReport {
  counts: RenameReferences;
  /** Sheet names whose part changed. */
  sheets: string[];
  calcChainDropped?: string;
}

/**
 * Stretch G (M5), the first step of a build with renames that carry `references`: every
 * formula that reads a renamed name gets its token rewritten (`renameInFormula`), in the
 * definitions of `<definedNames>` and in the worksheets' cell formulas, conditional
 * formats and validations, all read on the names as they are before the renames. Only
 * the text of those formulas changes; the names themselves are renamed by the change set
 * afterwards. After a cell formula changes the calculation chain goes (F8). A formula the
 * rename cannot rewrite safely (it would read another name, or does not parse) throws:
 * the plan refuses those before.
 */
export function applyReferenceRenames(bytes: Uint8Array, renames: readonly RenameName[]): { bytes: Uint8Array; report: ReferenceRenameReport } {
  const report: ReferenceRenameReport = { counts: { cells: 0, formats: 0, validations: 0, names: 0 }, sheets: [] };
  if (renames.length === 0) return { bytes, report };
  const pkg = new Package(bytes);
  const wbPart = workbookPartOf(pkg);
  const wbBefore = pkg.raw(wbPart);
  if (!wbBefore) throw new XlsxError(`workbook part ${wbPart} is missing`);
  const wbXml = decodeText(wbBefore);
  const layout = scanWorkbookXml(wbXml);
  const scopeOfItem = (e: DefinedNameElement): string | null | undefined => {
    const lsid = e.attrs["localSheetId"];
    if (lsid === undefined) return null;
    return layout.sheets[Number(lsid)];
  };
  const names: ScopedName[] = [];
  for (const e of layout.definedNames?.items ?? []) {
    const scope = scopeOfItem(e);
    if (scope !== undefined) names.push({ name: e.attrs["name"] ?? "", scope: scope ?? undefined });
  }
  const ctx = new RenameContext(
    names,
    renames.map((r) => ({ scope: r.scope ?? undefined, from: r.from, to: r.to })),
  );
  const failures: string[] = [];
  const rewrite = (xml: string, part: "workbook" | "sheet", home: (n: FormulaTextNode) => string | undefined, where: (n: FormulaTextNode) => string): { xml: string; changed: Set<FormulaPlace> } => {
    const changed = new Set<FormulaPlace>();
    let out = "";
    let at = 0;
    for (const n of formulaTextNodes(xml, part)) {
      const r = renameInFormula(n.text, home(n), ctx);
      if (r.captured || r.unparsed) failures.push(`${where(n)}: ${r.unparsed ? "does not parse" : `would read another name (${r.captured!.join(", ")})`}`);
      if (r.count === 0) continue;
      out += xml.slice(at, n.start) + escapeXmlText(r.text);
      at = n.end;
      changed.add(n.place);
      const k = n.place === "cell" ? "cells" : n.place === "format" ? "formats" : n.place === "validation" ? "validations" : n.place === "name" ? "names" : undefined;
      if (k) report.counts[k]++;
    }
    return { xml: out + xml.slice(at), changed };
  };

  const edits: PackageEdits = { replace: new Map(), remove: new Set(), add: new Map() };
  const items = new Map((layout.definedNames?.items ?? []).map((e) => [e.start, e]));
  // A definedName's text node: its element is the one opening at the node's open token.
  const wb = rewrite(
    wbXml,
    "workbook",
    (n) => {
      const e = items.get(n.open.start);
      return e ? (scopeOfItem(e) ?? undefined) : undefined;
    },
    (n) => `name ${n.open.attrs["name"] ?? "?"}`,
  );
  if (wb.xml !== wbXml) edits.replace.set(wbPart, encodeLike(wbBefore, wb.xml, wbPart));

  let cells = false;
  for (const s of sheetPartsOf(pkg, wbPart).values()) {
    if (!s.part) continue;
    const before = pkg.raw(s.part);
    if (!before) continue;
    const xml = decodeText(before);
    const r = rewrite(xml, "sheet", () => s.name, (n) => `${n.place === "cell" ? "cell" : n.place === "format" ? "conditional format on" : n.place === "validation" ? "validation on" : "formula on"} ${s.name}${n.ref ? `!${n.ref}` : ""}`);
    if (r.xml === xml) continue;
    edits.replace.set(s.part, encodeLike(before, r.xml, s.part));
    report.sheets.push(s.name);
    if (r.changed.has("cell")) cells = true;
  }
  if (failures.length) throw new ApplyError(`renaming ${renames.map((r) => `${scopedKey(r.from, r.scope)} to ${r.to}`).join(", ")}: ${failures.slice(0, 10).join("; ")}${failures.length > 10 ? "; …" : ""}`);

  if (cells) {
    const relsName = pkg.find(relsPartOf(wbPart));
    const ctName = pkg.find("[Content_Types].xml");
    if (relsName && ctName) {
      const relsBefore = pkg.raw(relsName)!;
      const ctBefore = pkg.raw(ctName)!;
      const d = dropCalcChain(pkg, wbPart, decodeText(relsBefore), decodeText(ctBefore), edits.remove);
      if (d.rels !== decodeText(relsBefore)) edits.replace.set(relsName, encodeLike(relsBefore, d.rels, relsName));
      if (d.ct !== decodeText(ctBefore)) edits.replace.set(ctName, encodeLike(ctBefore, d.ct, ctName));
      if (d.part) report.calcChainDropped = d.part;
    }
  }
  if (edits.replace.size === 0 && edits.remove.size === 0) return { bytes, report };
  return { bytes: rewriteZip(bytes, edits), report };
}

/** The renames of a change set that rewrite references too. */
export function referenceRenames(changes: readonly Change[]): RenameName[] {
  return changes.filter((c): c is RenameName => c.op === "rename-name" && c.references !== undefined);
}

/**
 * Applies a change set to workbook bytes and returns the new bytes. Throws `ApplyError`
 * when a change does not fit the workbook (a name to update is missing, a sheet is unknown,
 * a cell is part of a legacy array or a data table).
 */
export function applyChangeSet(bytes: Uint8Array, set: ChangeSet | readonly Change[]): Uint8Array {
  return applyChangeSetWithReport(bytes, set).bytes;
}

/** `applyChangeSet`, with what the cell patches did. */
export function applyChangeSetWithReport(original: Uint8Array, set: ChangeSet | readonly Change[]): { bytes: Uint8Array; report: ApplyReport } {
  const changes = Array.isArray(set) ? (set as readonly Change[]) : (set as ChangeSet).changes;
  // Stretch G first: the references to renamed names, read on the names before the renames.
  const refs = referenceRenames(changes);
  const renamed = refs.length ? applyReferenceRenames(original, refs) : undefined;
  const bytes = renamed?.bytes ?? original;
  const pkg = new Package(bytes);
  const part = workbookPartOf(pkg);
  const before = pkg.raw(part);
  if (!before) throw new XlsxError(`workbook part ${part} is missing`);
  const edits: PackageEdits = { replace: new Map(), remove: new Set(), add: new Map() };
  // Keep the XML declaration's encoding: Excel writes UTF-8 (a BOM, if any, is kept too).
  edits.replace.set(part, encodeLike(before, applyToWorkbookXml(decodeText(before), changes), part));
  const report = applyCellChanges(pkg, part, changes, edits);
  if (renamed) report.references = renamed.report;
  let out = rewriteZip(bytes, edits);
  // The embedded source last, on the patched package: it adds its own relationship and content type.
  const embed = changes.find((c): c is SetEmbeddedSource => c.op === "set-embedded-source");
  if (embed) {
    const e = applyEmbeddedSource(out, embed.files);
    out = e.bytes;
    report.embedded = { item: e.item, outcome: e.outcome };
  }
  return { bytes: out, report };
}
