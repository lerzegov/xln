// E3 for cell changes: the built bytes, read again, must say what the cell changes meant
// and nothing else. Worked out from the original workbook and the change list alone, not
// from what the patcher reports, so a patcher bug cannot excuse itself:
//   - each changed cell holds its formula in dynamic-array form (`cm` naming the XLDAPR
//     record, `<f t="array" ref=cell>`), no cached value, and decompiles to the change's
//     display text moved to that cell (modulo whitespace); its other attributes are kept;
//   - each cleared cell has no formula, value or value type;
//   - an array whose anchor changed has its old area emptied (values and value types gone,
//     styles kept); a shared group whose master changed has every other member holding the
//     master's text moved to it, with its value kept;
//   - every other cell, on every sheet, is exactly as it was; so is each changed sheet
//     outside <sheetData>, and its rows' attributes;
//   - the calculation chain is gone (part, relationship, content type); the workbook's
//     relationships and the content types differ by nothing else than that and a new
//     metadata part.

import { columnName, parseCell } from "../file/cellref.js";
import { decodeText, Package, relTypeIs } from "../file/package.js";
import type { CellFormula, WorkbookSnapshot } from "../file/types.js";
import { childElements, parseXml } from "../file/xml.js";
import { equalModuloWhitespace } from "../lang/format.js";
import { shiftFormula } from "../lang/shift.js";
import { decompile } from "../lang/transform.js";
import { checkStoredForm } from "../lang/stored.js";
import { relsPartOf, sheetPartsOf, workbookPartOf } from "./apply.js";
import { orderChanges, type ClearCellFormula, type SetCellFormula } from "./changes.js";
import { outsideSheetData, parseRange, readSheetCells, rectCells, scanSheetXml, type CellSnapshot } from "./sheetXml.js";

export interface PackageFacts {
  /** Parts a cell build may rewrite. */
  mayChange: string[];
  /** Parts it removed (the calculation chain). */
  removed: Set<string>;
  /** Parts it added (a new metadata part). */
  added: Set<string>;
}

export function relations(pkg: Package, relsPart: string): Map<string, string> {
  const text = pkg.text(relsPart);
  const out = new Map<string, string>();
  if (text === undefined) return out;
  for (const r of childElements(parseXml(text), "Relationship")) out.set(r.attrs["Id"] ?? "", `${r.attrs["Type"] ?? ""} ${r.attrs["Target"] ?? ""} ${r.attrs["TargetMode"] ?? ""}`);
  return out;
}

export function contentTypes(pkg: Package): Map<string, string> {
  const out = new Map<string, string>();
  const text = pkg.text("[Content_Types].xml");
  if (text === undefined) return out;
  for (const e of childElements(parseXml(text))) {
    const key = e.local === "Override" ? `O ${(e.attrs["PartName"] ?? "").toLowerCase()}` : `D ${(e.attrs["Extension"] ?? "").toLowerCase()}`;
    out.set(key, e.attrs["ContentType"] ?? "");
  }
  return out;
}

export function sameMaps(a: Map<string, string>, b: Map<string, string>): boolean {
  if (a.size !== b.size) return false;
  for (const [k, v] of a) if (b.get(k) !== v) return false;
  return true;
}

/** Relationships and content-type overrides another step of the same build added (the embedded source, D5). */
export interface PackageIgnore {
  relIds: ReadonlySet<string>;
  /** Lower-case part names with their leading `/`. */
  overrides: ReadonlySet<string>;
}

/** The package-level facts of a cell build; problems found on the way are added to `problems`. */
export function cellPackageFacts(
  beforeBytes: Uint8Array,
  afterBytes: Uint8Array,
  changes: readonly (SetCellFormula | ClearCellFormula)[],
  problems: string[],
  ignore: PackageIgnore = { relIds: new Set(), overrides: new Set() },
): PackageFacts {
  const a = new Package(beforeBytes);
  const b = new Package(afterBytes);
  const wbPart = workbookPartOf(a);
  const relsPart = a.find(relsPartOf(wbPart)) ?? relsPartOf(wbPart);
  const facts: PackageFacts = { mayChange: [relsPart, a.find("[Content_Types].xml") ?? "[Content_Types].xml"], removed: new Set(), added: new Set() };

  const sheets = sheetPartsOf(a, wbPart);
  for (const c of changes) {
    const part = sheets.get(c.sheet.toLowerCase())?.part;
    if (part) facts.mayChange.push(part);
    else problems.push(`${c.sheet}!${c.range}: no worksheet '${c.sheet}' in the original`);
  }

  const relsA = a.rels(wbPart);
  const relsB = b.rels(wbPart);
  for (const r of relsA.filter((x) => relTypeIs(x.type, "calcChain") && !x.external)) {
    const part = a.find(r.target);
    if (part) facts.removed.add(part);
  }
  if (relsB.some((r) => relTypeIs(r.type, "calcChain"))) problems.push("the calculation chain is still referenced: Excel would repair the file (F8)");
  for (const p of b.names) if (p.toLowerCase().endsWith("calcchain.xml")) problems.push(`${p} is still in the archive: Excel would repair the file (F8)`);

  const mdA = relsA.find((r) => relTypeIs(r.type, "sheetMetadata") && !r.external);
  const mdB = relsB.find((r) => relTypeIs(r.type, "sheetMetadata") && !r.external);
  if (!mdB || !b.has(mdB.target)) problems.push("no metadata part: the formulas written cannot be dynamic arrays");
  else if (!mdA) {
    if (a.has(mdB.target)) problems.push(`the new metadata relationship points at an existing part ${mdB.target}`);
    facts.added.add(b.find(mdB.target)!);
  } else {
    const part = a.find(mdA.target);
    if (part) facts.mayChange.push(part);
    if (mdA.target.toLowerCase() !== mdB.target.toLowerCase()) problems.push("the metadata relationship points at another part");
  }

  // Relationships and content types: only the calculation chain gone and the metadata part added.
  const ra = relations(a, relsPart);
  const rb = relations(b, relsPart);
  for (const [id, v] of [...ra]) if (v.includes("/calcChain ")) ra.delete(id);
  if (!mdA && mdB) rb.delete(mdB.id);
  for (const id of ignore.relIds) if (!ra.has(id)) rb.delete(id);
  if (!sameMaps(ra, rb)) problems.push(`${relsPart} differs from the original by more than the calculation chain and the metadata part`);
  const ca = contentTypes(a);
  const cb = contentTypes(b);
  for (const p of facts.removed) ca.delete(`O /${p.toLowerCase()}`);
  for (const p of facts.added) cb.delete(`O /${p.toLowerCase()}`);
  for (const o of ignore.overrides) if (!ca.has(`O ${o}`)) cb.delete(`O ${o}`);
  if (!sameMaps(ca, cb)) problems.push("[Content_Types].xml differs from the original by more than the calculation chain and the metadata part");
  for (const p of facts.removed) if (contentTypes(b).has(`O /${p.toLowerCase()}`)) problems.push(`[Content_Types].xml still lists ${p}`);
  return facts;
}

type Expect =
  | { kind: "set"; display: string; stored: string; dRow: number; dCol: number }
  | { kind: "clear" }
  | { kind: "ghost" }
  | { kind: "unshare"; text: string };

const VALUE_ATTRS = new Set(["t", "cm", "vm"]);

function attrsWithout(attrs: Record<string, string>, drop: ReadonlySet<string>): string {
  return JSON.stringify(Object.entries(attrs).filter(([k]) => !drop.has(k)));
}

function sameCell(x: CellSnapshot, y: CellSnapshot): boolean {
  return JSON.stringify(x.attrs) === JSON.stringify(y.attrs) && JSON.stringify(x.f) === JSON.stringify(y.f) && x.v === y.v && x.is === y.is && x.other === y.other;
}

function sameFormula(x: CellFormula, y: CellFormula | undefined): boolean {
  return !!y && x.kind === y.kind && x.text === y.text && x.range === y.range && x.si === y.si && x.cm === y.cm;
}

/** Checks the cells; returns how many changed, emptied and un-shared cells it checked. */
export function checkCells(
  before: WorkbookSnapshot,
  after: WorkbookSnapshot,
  beforeBytes: Uint8Array,
  afterBytes: Uint8Array,
  changes: readonly (SetCellFormula | ClearCellFormula)[],
  problems: string[],
): number {
  const a = new Package(beforeBytes);
  const b = new Package(afterBytes);
  const names = after.definedNames.map((d) => d.name);
  let checked = 0;

  // What each sheet's cells must become; a later change to a cell wins.
  const bySheet = new Map<string, Map<string, Expect>>();
  for (const c of orderChanges(changes) as (SetCellFormula | ClearCellFormula)[]) {
    const sheet = before.sheets.find((s) => s.name.toLowerCase() === c.sheet.toLowerCase());
    const rect = parseRange(c.range);
    if (!sheet || !rect) continue;
    let m = bySheet.get(sheet.name);
    if (!m) bySheet.set(sheet.name, (m = new Map()));
    for (const x of rectCells(rect)) {
      m.set(columnName(x.col) + x.row, c.op === "set-cell-formula" ? { kind: "set", display: c.display, stored: c.stored, dRow: x.row - rect.r1, dCol: x.col - rect.c1 } : { kind: "clear" });
    }
  }

  for (const sheet of before.sheets) {
    const afterSheet = after.sheets.find((s) => s.name === sheet.name);
    const targets = bySheet.get(sheet.name);
    if (!targets) {
      // Untouched sheet: its part is byte-identical (checked with the archive); the metadata
      // its cells point at must still mean the same.
      const fb = new Map((afterSheet?.formulas ?? []).map((f) => [f.cell, f]));
      for (const f of sheet.formulas) if (!sameFormula(f, fb.get(f.cell))) problems.push(`${sheet.name}!${f.cell}: its formula reads differently after the build`);
      continue;
    }
    if (!sheet.part || !afterSheet?.part) {
      problems.push(`sheet ${sheet.name}: no worksheet part`);
      continue;
    }
    const xa = decodeText(a.raw(sheet.part)!);
    const xb = decodeText(b.raw(afterSheet.part)!);
    if (outsideSheetData(xa) !== outsideSheetData(xb)) problems.push(`sheet ${sheet.name} changed outside <sheetData>`);
    const rowsA = new Map(scanSheetXml(xa).rows.map((r) => [r.r, r]));
    for (const r of scanSheetXml(xb).rows) {
      const old = rowsA.get(r.r);
      if (old && JSON.stringify(old.open.attrs) !== JSON.stringify(r.open.attrs)) problems.push(`sheet ${sheet.name}: row ${r.r}'s attributes changed`);
      if (!old && !r.cells.every((c) => targets.get(c.ref)?.kind === "set")) problems.push(`sheet ${sheet.name}: row ${r.r} appeared with cells no change wrote`);
    }
    const ca = readSheetCells(xa);
    const cb = readSheetCells(xb);
    const fa = new Map(sheet.formulas.map((f) => [f.cell, f]));
    const fb = new Map(afterSheet.formulas.map((f) => [f.cell, f]));

    const expect = new Map<string, Expect>(targets);
    // A clear removes formulas only: a cell without one must stay as it was.
    for (const [ref, e] of targets) if (e.kind === "clear" && !ca.get(ref)?.f) expect.delete(ref);
    // Old array areas whose anchor changed.
    for (const f of sheet.formulas) {
      if ((f.kind !== "dynamic-array" && f.kind !== "array") || !targets.has(f.cell) || !f.range) continue;
      const rect = parseRange(f.range);
      if (!rect) continue;
      for (const x of rectCells(rect)) {
        const ref = columnName(x.col) + x.row;
        if (!expect.has(ref) && ca.has(ref) && !ca.get(ref)!.f) expect.set(ref, { kind: "ghost" });
      }
    }
    // Shared groups whose master changed.
    for (const g of sheet.sharedFormulas) {
      if (!g.master || g.text === undefined || !targets.has(g.master)) continue;
      const m = parseCell(g.master)!;
      for (const ref of g.cells) {
        if (expect.has(ref)) continue;
        const c = parseCell(ref)!;
        expect.set(ref, { kind: "unshare", text: shiftFormula(g.text, c.row - m.row, c.col - m.col) });
      }
    }

    for (const [ref, e] of expect) {
      const x = ca.get(ref);
      const y = cb.get(ref);
      const at = `${sheet.name}!${ref}`;
      checked++;
      switch (e.kind) {
        case "set": {
          if (!y?.f) {
            problems.push(`${at}: no formula in the built file`);
            break;
          }
          if (y.f.attrs["t"] !== "array" || y.f.attrs["ref"] !== ref) problems.push(`${at}: not written as a dynamic array anchored at the cell`);
          if (fb.get(ref)?.kind !== "dynamic-array") problems.push(`${at}: its cm does not name the dynamic-array metadata`);
          if (y.v !== undefined || y.is !== undefined || y.attrs["t"] !== undefined) problems.push(`${at}: keeps a cached value`);
          if (x && attrsWithout(x.attrs, VALUE_ATTRS) !== attrsWithout(y.attrs, VALUE_ATTRS)) problems.push(`${at}: its attributes (style) changed`);
          if (x && x.other !== y.other) problems.push(`${at}: lost or changed child elements`);
          let display: string;
          try {
            display = decompile(y.f.text, { names, links: after.externalLinks ?? [] });
          } catch (err) {
            problems.push(`${at}: the built formula does not parse: ${(err as Error).message}`);
            break;
          }
          // Excel's stored grammar, read on its own terms: our decompiler would read back a
          // wrong stored form of ours as if it were right (M3d: `[_xlpm.p]`).
          for (const d of checkStoredForm(y.f.text, { names })) problems.push(`${at}: not Excel's stored form: ${d.message}`);
          const want = shiftFormula(e.display, e.dRow, e.dCol);
          if (!equalModuloWhitespace(display, want)) problems.push(`${at}: the built file says ${display}, the change says ${want}`);
          break;
        }
        case "clear":
        case "ghost":
          if (y && (y.f || y.v !== undefined || y.is !== undefined || y.attrs["t"] !== undefined || y.attrs["cm"] !== undefined)) {
            problems.push(`${at}: ${e.kind === "clear" ? "not cleared" : "the old spill's value is still there"}`);
          }
          if (x && y && attrsWithout(x.attrs, VALUE_ATTRS) !== attrsWithout(y.attrs, VALUE_ATTRS)) problems.push(`${at}: its attributes (style) changed`);
          if (e.kind === "clear" && x && !y) problems.push(`${at}: the cell element was removed`);
          break;
        case "unshare": {
          const g = fb.get(ref);
          if (!x || !y || !g) {
            problems.push(`${at}: the shared formula's member is missing`);
            break;
          }
          if (g.kind !== "normal" || g.text !== e.text) problems.push(`${at}: un-shared as ${g.text ?? "(nothing)"}, expected ${e.text}`);
          if (x.v !== y.v || JSON.stringify(x.attrs) !== JSON.stringify(y.attrs) || x.is !== y.is || x.other !== y.other) problems.push(`${at}: its value or attributes changed`);
          break;
        }
      }
    }
    for (const [ref, x] of ca) {
      if (expect.has(ref)) continue;
      const y = cb.get(ref);
      if (!y) problems.push(`${sheet.name}!${ref} disappeared`);
      else if (!sameCell(x, y)) problems.push(`${sheet.name}!${ref} changed although no change touched it`);
      const f = fa.get(ref);
      if (f && !sameFormula(f, fb.get(ref))) problems.push(`${sheet.name}!${ref}: its formula reads differently after the build`);
    }
    for (const ref of cb.keys()) if (!ca.has(ref) && expect.get(ref)?.kind !== "set") problems.push(`${sheet.name}!${ref} appeared although no change wrote it`);
  }
  return checked;
}
