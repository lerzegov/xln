// E3 for the first step of a build with renames (stretch G): the references rewritten by
// `applyReferenceRenames`, read again with the workbook reader (not the patcher's
// scanner), must be exactly the formulas before with the renamed tokens replaced, and
// nothing else may differ:
//   - every defined name keeps its name, scope, comment, flags and attributes; its
//     definition is the old one renamed;
//   - every cell formula keeps its kind, extent, shared index and metadata; its text is
//     the old one renamed (a shared member still has none); so are the conditional
//     formats' and validations' formulas;
//   - the workbook part and the sheets differ only inside formula texts; every other zip
//     entry is byte-identical, except the calculation chain, which goes after a cell
//     formula changed (with its relationship and content type, and nothing else).

import { decodeText, Package, relTypeIs } from "../file/package.js";
import type { DefinedName, WorkbookSnapshot } from "../file/types.js";
import { readWorkbook } from "../file/workbook.js";
import { rawZipRecords, readZipLayout } from "../file/zip.js";
import { RenameContext, renameInFormula } from "../project/rename.js";
import { relsPartOf, sheetPartsOf, workbookPartOf } from "./apply.js";
import { scopedKey, type RenameName } from "./changes.js";
import { withoutFormulaTexts } from "./formulaText.js";
import { contentTypes, relations, sameMaps } from "./readbackCells.js";

function scopeOf(d: DefinedName): string | undefined {
  return d.scope.kind === "sheet" ? d.scope.name : undefined;
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** Checks `after` against `before` with `renames` applied to the references; adds what is wrong to `problems`. */
export function checkReferenceRenames(beforeBytes: Uint8Array, afterBytes: Uint8Array, renames: readonly RenameName[], problems: string[]): void {
  const at = (m: string) => problems.push(`renamed references: ${m}`);
  let a: WorkbookSnapshot;
  let b: WorkbookSnapshot;
  try {
    a = readWorkbook(beforeBytes);
    b = readWorkbook(afterBytes);
  } catch (e) {
    at(`the file cannot be read: ${(e as Error).message}`);
    return;
  }
  const ctx = new RenameContext(
    a.definedNames.filter((d) => !d.scopeInvalid).map((d) => ({ name: d.name, scope: scopeOf(d) })),
    renames.map((r) => ({ scope: r.scope ?? undefined, from: r.from, to: r.to })),
  );
  const renamed = (text: string, home: string | undefined) => renameInFormula(text, home, ctx).text;

  // Names: in file order, nothing but the definition's renamed tokens.
  if (a.definedNames.length !== b.definedNames.length) at(`${a.definedNames.length} names before, ${b.definedNames.length} after`);
  a.definedNames.forEach((d, i) => {
    const e = b.definedNames[i];
    const label = scopedKey(d.name, scopeOf(d) ?? null);
    if (!e) return;
    if (e.name !== d.name || scopeOf(e) !== scopeOf(d) || e.comment !== d.comment || e.hidden !== d.hidden || JSON.stringify(e.attributes) !== JSON.stringify(d.attributes)) at(`name ${label} changed beyond its definition`);
    const want = d.scopeInvalid ? d.definition : renamed(d.definition, scopeOf(d));
    if (e.definition !== want) at(`name ${label} reads ${e.definition}, expected ${want}`);
  });

  // Sheets: every formula the same but for the renamed tokens.
  for (const s of a.sheets) {
    const t = b.sheets.find((x) => x.name === s.name);
    if (!t) {
      at(`sheet ${s.name} is gone`);
      continue;
    }
    const fb = new Map(t.formulas.map((f) => [f.cell, f]));
    if (s.formulas.length !== t.formulas.length) at(`sheet ${s.name}: ${s.formulas.length} formula cells before, ${t.formulas.length} after`);
    for (const f of s.formulas) {
      const g = fb.get(f.cell);
      const where = `${s.name}!${f.cell}`;
      if (!g) {
        at(`${where} lost its formula`);
        continue;
      }
      if (g.kind !== f.kind || g.range !== f.range || g.si !== f.si || g.cm !== f.cm || JSON.stringify(g.attributes) !== JSON.stringify(f.attributes)) at(`${where}: the formula's kind, extent or attributes changed`);
      const want = f.text === undefined ? undefined : renamed(f.text, s.name);
      if (g.text !== want) at(`${where} reads ${g.text ?? "(no text)"}, expected ${want ?? "(no text)"}`);
      if (JSON.stringify(g.value) !== JSON.stringify(f.value)) at(`${where}: its cached value changed`);
    }
    const cfA = s.conditionalFormats.flatMap((c) => c.formulas.map((x) => renamed(x, s.name)));
    const cfB = t.conditionalFormats.flatMap((c) => c.formulas);
    if (JSON.stringify(cfA) !== JSON.stringify(cfB)) at(`sheet ${s.name}: the conditional formats' formulas are not the old ones renamed`);
    const dv = (x: string | undefined) => (x === undefined ? null : renamed(x, s.name));
    const dvA = s.dataValidations.map((v) => [dv(v.formula1), dv(v.formula2)]);
    const dvB = t.dataValidations.map((v) => [v.formula1 ?? null, v.formula2 ?? null]);
    if (JSON.stringify(dvA) !== JSON.stringify(dvB)) at(`sheet ${s.name}: the validations' formulas are not the old ones renamed`);
  }

  // The package: formula texts aside, every byte as it was; the calculation chain may go.
  try {
    const pa = new Package(beforeBytes);
    const pb = new Package(afterBytes);
    const wbPart = workbookPartOf(pa);
    const relsPart = pa.find(relsPartOf(wbPart)) ?? relsPartOf(wbPart);
    const ctPart = pa.find("[Content_Types].xml") ?? "[Content_Types].xml";
    const chain = new Set(pa.rels(wbPart).filter((r) => relTypeIs(r.type, "calcChain") && !r.external).map((r) => pa.find(r.target)).filter((p): p is string => p !== undefined));
    const sheetParts = new Map([...sheetPartsOf(pa, wbPart).values()].filter((s) => s.part !== undefined).map((s) => [s.part!, s.name]));
    const la = readZipLayout(beforeBytes);
    const lb = readZipLayout(afterBytes);
    const namesA = la.entries.map((e) => e.name);
    const namesB = lb.entries.map((e) => e.name);
    const dropped = namesA.filter((n) => !namesB.includes(n));
    if (dropped.some((n) => !chain.has(n)) || namesB.some((n) => !namesA.includes(n)) || namesA.filter((n) => namesB.includes(n)).join("\n") !== namesB.join("\n")) at("the archive's entries differ by more than the calculation chain");
    const ra = rawZipRecords(beforeBytes, la);
    const rb = rawZipRecords(afterBytes, lb);
    for (const [name, rec] of ra) {
      const other = rb.get(name);
      if (!other) continue;
      if (name === wbPart || sheetParts.has(name)) {
        const kind = name === wbPart ? "workbook" : "sheet";
        if (withoutFormulaTexts(decodeText(pa.raw(name)!), kind) !== withoutFormulaTexts(decodeText(pb.raw(name)!), kind)) at(`${name} changed outside its formulas`);
      } else if (name === relsPart || name === ctPart) {
        if (dropped.length === 0 && !bytesEqual(rec, other)) at(`${name} changed`);
      } else if (!bytesEqual(rec, other)) at(`zip entry ${name} is not byte-identical to the original`);
    }
    if (dropped.length) {
      const relA = relations(pa, relsPart);
      for (const [id, v] of [...relA]) if (v.includes("/calcChain ")) relA.delete(id);
      if (!sameMaps(relA, relations(pb, relsPart))) at(`${relsPart} differs by more than the calculation chain`);
      const ctA = contentTypes(pa);
      for (const p of dropped) ctA.delete(`O /${p.toLowerCase()}`);
      if (!sameMaps(ctA, contentTypes(pb))) at("[Content_Types].xml differs by more than the calculation chain");
    }
    const cellsChanged = a.sheets.some((s) => s.formulas.some((f) => f.text !== undefined && renamed(f.text, s.name) !== f.text));
    if (cellsChanged && chain.size > 0 && dropped.length === 0) at("a cell formula changed but the calculation chain is still there: Excel would repair the file (F8)");
  } catch (e) {
    at(`the archive cannot be compared: ${(e as Error).message}`);
  }
}
