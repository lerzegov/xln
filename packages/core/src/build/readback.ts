// E3, read-back: before the built bytes replace the workbook, read them again as Excel
// would and check that they say what the build meant, and nothing else:
//   - the archive has the same entries in the same order, each byte-identical except the
//     workbook part (and, for cell changes, the parts readbackCells.ts allows); the
//     workbook part is identical outside <definedNames> and <calcPr>;
//   - every source name in sync decompiles to its source text (modulo whitespace), with
//     its comment, hidden flag and scope;
//   - every name no change touched kept its exact definition, comment and attributes;
//   - deleted and renamed-away names are gone; `fullCalcOnLoad` is set;
//   - cell changes: readbackCells.ts;
//   - the embedded source (D5) says exactly the files the change carries, and the package
//     gained nothing else than its part; comments carry the provenance tag (D6) the build meant.

import { readWorkbook } from "../file/workbook.js";
import type { DefinedName, WorkbookSnapshot } from "../file/types.js";
import { rawZipRecords, readZipLayout } from "../file/zip.js";
import { decodeText, Package } from "../file/package.js";
import { equalModuloWhitespace } from "../lang/format.js";
import { compile, decompile, type FormulaContext } from "../lang/transform.js";
import { checkStoredForm } from "../lang/stored.js";
import { applyReferenceRenames, referenceRenames, workbookPartOf } from "./apply.js";
import { checkReferenceRenames } from "./readbackRename.js";
import { scopedKey, type Change, type ClearCellFormula, type Scope, type SetCellFormula, type SetEmbeddedSource } from "./changes.js";
import { embedPackageFacts } from "./embed.js";
import { stripProvenance } from "../project/provenance.js";
import { cellPackageFacts, checkCells, type PackageFacts } from "./readbackCells.js";
import type { NameState } from "./plan.js";
import { scanWorkbookXml } from "./workbookXml.js";

export interface ReadBackReport {
  ok: boolean;
  problems: string[];
  /** Names checked against the source. */
  checked: number;
  /** Cells checked against the cell changes (changed cells, emptied spill cells, un-shared members). */
  cellsChecked: number;
}

function lkey(name: string, scope: Scope | undefined): string {
  return `${(scope ?? "").toLowerCase()}!${name.toLowerCase()}`;
}

function scopeOf(d: DefinedName): Scope {
  return d.scope.kind === "sheet" ? d.scope.name : null;
}

function bytesEqual(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/** The workbook part with <definedNames> and the <calcPr> tag cut out. */
function outsideNames(bytes: Uint8Array): string {
  const pkg = new Package(bytes);
  const xml = decodeText(pkg.raw(workbookPartOf(pkg))!);
  const l = scanWorkbookXml(xml);
  const cuts: { start: number; end: number }[] = [];
  if (l.definedNames) cuts.push(l.definedNames);
  if (l.calcPr) cuts.push(l.calcPr);
  cuts.sort((a, b) => a.start - b.start);
  let out = "";
  let at = 0;
  for (const c of cuts) {
    out += xml.slice(at, c.start);
    at = c.end;
  }
  return out + xml.slice(at);
}

export function readBack(originalBytes: Uint8Array, afterBytes: Uint8Array, changes: readonly Change[], inSync: readonly NameState[]): ReadBackReport {
  const problems: string[] = [];
  // Stretch G: a build with renames that rewrite references does that first. Its result is
  // checked on its own (readbackRename.ts); the rest of the build is checked against it.
  let beforeBytes = originalBytes;
  const refs = referenceRenames(changes);
  if (refs.length) {
    try {
      beforeBytes = applyReferenceRenames(originalBytes, refs).bytes;
    } catch (e) {
      return { ok: false, problems: [`the renamed references cannot be rewritten: ${(e as Error).message}`], checked: 0, cellsChecked: 0 };
    }
    checkReferenceRenames(originalBytes, beforeBytes, refs, problems);
  }
  let after: WorkbookSnapshot;
  let before: WorkbookSnapshot;
  try {
    before = readWorkbook(beforeBytes);
    after = readWorkbook(afterBytes);
  } catch (e) {
    return { ok: false, problems: [`the built file cannot be read: ${(e as Error).message}`], checked: 0, cellsChecked: 0 };
  }

  // The archive: same entries, same order, other parts byte for byte. A cell build may
  // also change the sheets it names, the workbook's relationships, the content types and
  // the metadata part, add a metadata part, and must drop the calculation chain.
  const cellChanges = changes.filter((c): c is SetCellFormula | ClearCellFormula => c.op === "set-cell-formula" || c.op === "clear-cell-formula");
  try {
    const embed = changes.find((c): c is SetEmbeddedSource => c.op === "set-embedded-source");
    const efacts = embed ? embedPackageFacts(beforeBytes, afterBytes, embed.files, problems, cellChanges.length === 0) : undefined;
    const ignore = { relIds: efacts?.addedRelIds ?? new Set<string>(), overrides: efacts?.addedOverrides ?? new Set<string>() };
    const facts: PackageFacts | undefined = cellChanges.length ? cellPackageFacts(beforeBytes, afterBytes, cellChanges, problems, ignore) : undefined;
    const la = readZipLayout(beforeBytes);
    const lb = readZipLayout(afterBytes);
    const removed = facts?.removed ?? new Set<string>();
    const added = new Set([...(facts?.added ?? []), ...(efacts?.added ?? [])]);
    const mayChange = new Set([before.workbookPart, ...(facts?.mayChange ?? []), ...(efacts?.mayChange ?? [])]);
    const expected = [...la.entries.map((e) => e.name).filter((n) => !removed.has(n)), ...lb.entries.map((e) => e.name).filter((n) => added.has(n))];
    if (expected.join("\n") !== lb.entries.map((e) => e.name).join("\n")) problems.push("the archive's entries differ from the original's");
    const ra = rawZipRecords(beforeBytes, la);
    const rb = rawZipRecords(afterBytes, lb);
    for (const [name, rec] of ra) {
      if (mayChange.has(name) || removed.has(name)) continue;
      const other = rb.get(name);
      if (!other || !bytesEqual(rec, other)) problems.push(`zip entry ${name} is not byte-identical to the original`);
    }
    if (outsideNames(beforeBytes) !== outsideNames(afterBytes)) problems.push(`${before.workbookPart} changed outside <definedNames> and <calcPr>`);
  } catch (e) {
    problems.push(`the archive cannot be compared: ${(e as Error).message}`);
  }

  // fullCalcOnLoad.
  try {
    const pkg = new Package(afterBytes);
    const l = scanWorkbookXml(decodeText(pkg.raw(workbookPartOf(pkg))!));
    if (l.calcPr?.open.attrs["fullCalcOnLoad"] !== "1") problems.push("<calcPr> lacks fullCalcOnLoad=\"1\"");
  } catch (e) {
    problems.push(`the workbook part cannot be read: ${(e as Error).message}`);
  }

  const index = (wb: WorkbookSnapshot) => {
    const m = new Map<string, DefinedName>();
    for (const d of wb.definedNames) m.set(lkey(d.name, scopeOf(d)), d);
    return m;
  };
  const was = index(before);
  const now = index(after);

  // Names the changes touch: by their key before and after the build.
  const touchedBefore = new Set<string>();
  const touchedAfter = new Set<string>();
  const gone = new Set<string>();
  for (const c of changes) {
    if (c.op === "set-name") {
      touchedBefore.add(lkey(c.name, c.scope));
      touchedAfter.add(lkey(c.name, c.scope));
    } else if (c.op === "delete-name") {
      touchedBefore.add(lkey(c.name, c.scope));
      gone.add(lkey(c.name, c.scope));
    } else if (c.op === "rename-name") {
      touchedBefore.add(lkey(c.from, c.scope));
      gone.add(lkey(c.from, c.scope));
      touchedAfter.add(lkey(c.to, c.scope));
    } else if (c.op === "rescope-name") {
      touchedBefore.add(lkey(c.name, c.from));
      gone.add(lkey(c.name, c.from));
      touchedAfter.add(lkey(c.name, c.to));
    }
  }
  for (const c of changes) {
    if (c.op === "set-name") gone.delete(lkey(c.name, c.scope));
    if (c.op === "rename-name") gone.delete(lkey(c.to, c.scope));
    if (c.op === "rescope-name") gone.delete(lkey(c.name, c.to));
  }
  for (const k of gone) if (now.has(k)) problems.push(`${now.get(k)!.name} should be gone but is still in the file`);

  // Untouched names: exactly as they were (Excel's own `_xl*` names included).
  for (const [k, d] of was) {
    if (touchedBefore.has(k)) continue;
    const n = now.get(k);
    const label = scopedKey(d.name, scopeOf(d));
    if (!n) problems.push(`${label} disappeared`);
    else if (n.name !== d.name || n.definition !== d.definition || n.comment !== d.comment || n.hidden !== d.hidden || JSON.stringify(n.attributes) !== JSON.stringify(d.attributes)) {
      problems.push(`${label} changed although no change touched it`);
    }
  }
  for (const k of now.keys()) if (!was.has(k) && !touchedAfter.has(k)) problems.push(`${now.get(k)!.name} appeared although no change created it`);

  // Source names: the file must decompile to the source.
  const allNames = after.definedNames.map((d) => d.name);
  let checked = 0;
  for (const s of inSync) {
    const label = scopedKey(s.name, s.scope);
    const d = now.get(lkey(s.name, s.scope));
    if (!d) {
      problems.push(`${label} is not in the built file`);
      continue;
    }
    checked++;
    if (d.name !== s.name) problems.push(`${label} is spelled ${d.name} in the built file`);
    const ctx =
      s.scope === null
        ? { names: allNames, links: after.externalLinks ?? [] }
        : { homeSheet: s.scope, localNames: after.definedNames.filter((x) => x.scope.kind === "sheet" && x.scope.name === s.scope).map((x) => x.name), names: allNames, links: after.externalLinks ?? [] };
    let display: string;
    try {
      display = decompile(d.definition, ctx);
    } catch (e) {
      problems.push(`${label}: the built definition does not parse: ${(e as Error).message}`);
      continue;
    }
    if (!equalModuloWhitespace(display, s.display) && !sameAfterCompile(display, s.display, ctx)) problems.push(`${label}: the built file says ${display}, the source says ${s.display}`);
    const comment = stripProvenance(d.comment === undefined || d.comment === "" ? null : d.comment.split("\r\n").join("\n")) ?? null;
    if (comment !== s.comment) problems.push(`${label}: comment differs from the source`);
    if (d.hidden !== s.hidden) problems.push(`${label}: hidden flag differs from the source`);
  }
  // Every definition the build wrote, against Excel's stored grammar on its own terms: the
  // decompile above would read a wrong stored form of ours back as if it were right
  // (M3d: `[_xlpm.p]` passed, Excel dropped the name).
  for (const c of changes) {
    if (c.op !== "set-name") continue;
    const d = now.get(lkey(c.name, c.scope));
    if (!d) continue;
    for (const p of checkStoredForm(d.definition, { names: allNames })) problems.push(`${scopedKey(c.name, c.scope)}: not Excel's stored form: ${p.message}`);
  }
  // The comments the build wrote, provenance tag included (D6), exactly.
  for (const c of changes) {
    if (c.op !== "set-name") continue;
    const d = now.get(lkey(c.name, c.scope));
    const got = d?.comment === undefined || d.comment === "" ? null : d.comment.split("\r\n").join("\n");
    if (d && got !== c.comment) problems.push(`${scopedKey(c.name, c.scope)}: the comment written is not the one the build meant`);
  }
  let cellsChecked = 0;
  try {
    if (cellChanges.length) cellsChecked = checkCells(before, after, beforeBytes, afterBytes, cellChanges, problems);
  } catch (e) {
    problems.push(`the cells cannot be compared: ${(e as Error).message}`);
  }
  return { ok: problems.length === 0, problems, checked, cellsChecked };
}

/**
 * The source's text says what the built text says once compiled: a local name's own sheet
 * (`'S1'!$B$2`, as a workbook name on S1's cells reads before `@workbook` is removed)
 * decompiles without it (`$B$2`), and both store as `'S1'!$B$2`.
 */
function sameAfterCompile(display: string, source: string, ctx: FormulaContext): boolean {
  try {
    return equalModuloWhitespace(display, decompile(compile(source, { ...ctx, crlf: true, allowUnknownFunctions: true }), ctx));
  } catch {
    return false;
  }
}
