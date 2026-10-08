// Shared helpers for the build tests: fixtures, a pulled project as a file map, edits.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { decodeText, Package } from "../../src/file/package.js";
import { buildWorkbook, compile, equalModuloWhitespace, pullProject, replaceZipEntries, scanWorkbookXml, utf8, type BuildResult, type Change, type DefinedName, type NameState } from "../../src/index.js";

export const RESULTS = join(import.meta.dirname, "..", "..", "..", "..", "probes", "results");
export const FIXTURES = join(import.meta.dirname, "..", "..", "..", "..", "probes", "fixtures");

export function fixture(name: string): Uint8Array {
  const dir = name === "traps.xlsx" ? FIXTURES : RESULTS;
  return new Uint8Array(readFileSync(join(dir, name)));
}

export function pulled(bytes: Uint8Array, name = "book.xlsx"): Record<string, string> {
  return { ...pullProject(bytes, name).files };
}

/** The workbook part's text. */
export function workbookXml(bytes: Uint8Array): string {
  const pkg = new Package(bytes);
  return decodeText(pkg.raw("xl/workbook.xml")!);
}

/** The `<definedNames>` element as stored, or "" when there is none. */
export function definedNamesXml(bytes: Uint8Array): string {
  const xml = workbookXml(bytes);
  const dn = scanWorkbookXml(xml).definedNames;
  return dn ? xml.slice(dn.start, dn.end) : "";
}

/** The workbook with its workbook part replaced (to simulate edits made in Excel). */
export function withWorkbookXml(bytes: Uint8Array, xml: string): Uint8Array {
  return replaceZipEntries(bytes, new Map([["xl/workbook.xml", utf8(xml)]]));
}

/** Replaces `find` (which must occur once) in a project file. */
export function edit(files: Record<string, string>, path: string, find: string, replace: string): void {
  const text = files[path];
  if (text === undefined) throw new Error(`no file ${path}; have ${Object.keys(files).join(", ")}`);
  const at = text.indexOf(find);
  if (at < 0 || text.indexOf(find, at + 1) >= 0) throw new Error(`'${find}' is not in ${path} exactly once:\n${text}`);
  files[path] = text.slice(0, at) + replace + text.slice(at + find.length);
}

export function build(bytes: Uint8Array, files: Record<string, string>, force = false): BuildResult {
  return buildWorkbook({ workbook: bytes, fileName: "book.xlsx", files }, { force, embed: false, provenance: false });
}

/** Problems and conflicts of a result, as text (for failing assertions that explain themselves). */
export function why(r: BuildResult): string {
  return [...r.plan.problems.map((p) => `${p.severity} ${p.code}: ${p.message}`), ...r.plan.conflicts.map((c) => `conflict ${c.kind}: ${c.message}`), ...(r.readBack?.problems ?? []), r.error ?? ""].join("\n");
}
/** Every name rewritten from its pulled source text, as a change set. */
export function rewriteAll(bytes: Uint8Array): { changes: Change[]; inSync: NameState[] } {
  const r = pullProject(bytes, "book.xlsx");
  const all = r.names.map((n) => n.name);
  const changes: Change[] = [];
  const inSync: NameState[] = [];
  for (const n of r.names) {
    if (n.classification.kind === "unparsed") continue;
    const scope = n.scope ?? null;
    const links = r.snapshot.externalLinks;
    const ctx = n.scope === undefined ? { names: all, links } : { homeSheet: n.scope, localNames: r.names.filter((m) => m.scope === n.scope).map((m) => m.name), names: all, links };
    const stored = compile(n.display, { ...ctx, crlf: true, allowUnknownFunctions: true });
    changes.push({ op: "set-name", name: n.name, scope, stored, display: n.display, comment: n.comment ?? null, hidden: n.hidden, fields: ["definition"] });
    inSync.push({ name: n.name, scope, display: n.display, comment: n.comment ?? null, hidden: n.hidden });
  }
  return { changes, inSync };
}

/** Semantic equality of two lists of defined names; `repaired` names must have gained their prefix. */
export function sameNames(a: readonly DefinedName[], b: readonly DefinedName[], repaired: readonly string[] = []): string[] {
  const key = (d: DefinedName) => `${d.scope.kind === "sheet" ? d.scope.position : "-"}!${d.name.toLowerCase()}`;
  const bm = new Map(b.map((d) => [key(d), d]));
  const out: string[] = [];
  if (a.length !== b.length) out.push(`${a.length} names before, ${b.length} after`);
  for (const d of a) {
    const e = bm.get(key(d));
    if (!e) out.push(`${key(d)} missing`);
    else if (repaired.includes(d.name)) {
      if (!e.definition.startsWith("_xlfn.")) out.push(`${key(d)}: not repaired: ${e.definition}`);
    } else if (!equalModuloWhitespace(d.definition, e.definition)) out.push(`${key(d)}: ${d.definition.slice(0, 200)} → ${e.definition.slice(0, 200)}`);
    else if ((d.comment ?? "") !== (e.comment ?? "") || d.hidden !== e.hidden || d.name !== e.name) out.push(`${key(d)}: comment, hidden or spelling differ`);
  }
  return out;
}

