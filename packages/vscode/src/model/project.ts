// One xln project folder (`<workbook>.xln/`) as the editor sees it: the core's source model
// (its `.xln` files parsed, each formula parsed and resolved, M3d) with the manifest
// written by `xln pull`, the lockfile and the project's settings. Plain TypeScript
// without the vscode API, so Vitest can test it; positions are offsets into a file's text.
//
// Names are resolved from the live sources, not from the manifest: after an edit, go to
// definition and "used by" among names follow the text. Cell usages can only come from
// the manifest (the workbook is read-only here), so they reflect the last pull. The
// problems of a file are the core checker's (the build refuses exactly its errors), with
// the manifest's spill map as the workbook's.

import {
  checkFile,
  renameInSource,
  type SourceRename,
  sourceToFormula,
  SourceModel,
  type CheckContext,
  type Classification,
  type Lockfile,
  type Loc,
  type NameDef,
  type Occurrence,
  type Problem,
} from "@xln/core";
import { type Manifest, type ManifestName } from "./manifest.js";

export { defKey, occurrences, type Analysis, type Fix, type Loc, type NameDef, type Occurrence, type Problem, type SourceFile } from "@xln/core";

export const NAMES_DIR = "names";

export type SymbolKind = "block" | Classification["kind"];

export interface OutlineSymbol {
  name: string;
  detail: string;
  kind: SymbolKind;
  start: number;
  end: number;
  selectionStart: number;
  selectionEnd: number;
  children: OutlineSymbol[];
}

export interface SearchHit {
  def: NameDef;
  /** Where the query matched: some of `name`, `doc`, `formula`. */
  matched: ("name" | "doc" | "formula")[];
}

const lower = (s: string) => s.toLowerCase();

export class Project extends SourceModel {
  private manifestByKey: Map<string, ManifestName> | undefined;
  private usedByCache: Map<string, { def: NameDef; occ: Occurrence & { kind: "use" } }[]> | undefined;
  /** The lockfile (the last pull or build): the addresses and scopes the checker compares with. */
  lock: Lockfile | undefined;

  /**
   * @param root where the project lives, as the caller names folders (a URI string)
   * @param manifest undefined when `workbook.manifest.json` is missing or unreadable
   */
  constructor(
    readonly root: string,
    public manifest: Manifest | undefined,
  ) {
    super();
  }

  protected override invalidate(): void {
    super.invalidate();
    this.usedByCache = undefined;
  }

  /** What the checker knows of the workbook: its sheets and Tables (the manifest), the lockfile. */
  checkContext(): CheckContext {
    const ctx: CheckContext = {};
    if (this.manifest && this.manifest.sheets.length) ctx.sheets = this.manifest.sheets.map((s) => s.name);
    if (this.manifest) ctx.tables = this.manifest.tables.map((t) => t.name);
    if (this.lock) ctx.lock = this.lock;
    // The spills as the workbook was saved at the last pull: a name without `#` on one is a warning.
    if (this.manifest?.spills) ctx.spills = this.manifest.spills;
    // Links to other workbooks (no key: none): a formula naming another workbook needs one.
    if (this.manifest) ctx.links = this.manifest.externalLinks ?? [];
    return ctx;
  }
  get workbookName(): string | undefined {
    return this.manifest?.workbook || undefined;
  }
  /** The manifest's record of a key, if the last pull saw it. */
  manifestName(key: string): ManifestName | undefined {
    if (!this.manifest) return undefined;
    if (!this.manifestByKey) {
      this.manifestByKey = new Map(Object.entries(this.manifest.names).map(([k, v]) => [lower(k), v]));
    }
    return this.manifestByKey.get(lower(key));
  }

  /** What a cell statement's name is, as the last pull classified it: a cell or the spill of one. */
  protected override cellClassification(def: NameDef): Classification {
    if (def.name === "") return { kind: "formula" };
    const mn = this.manifestName(def.key);
    if (mn?.kind === "spill" && mn.spill) return { kind: "spill", anchor: { sheet: mn.spill.sheet ?? undefined, cell: mn.spill.anchor } };
    return { kind: mn?.kind === "spill" ? "spill" : "range" };
  }

  /** What the identifier at `offset` is: the name being defined, a use, or a local. */
  hit(path: string, offset: number): { def: NameDef; target: "self" | Occurrence } | undefined {
    const def = this.defAt(path, offset);
    if (!def) return undefined;
    if (def.name !== "" && offset >= def.entry.offset && offset <= def.entry.offset + def.name.length) return { def, target: "self" };
    const i = sourceToFormula(def.entry, offset);
    if (i === undefined) return undefined;
    const occ = this.analysis(def).occurrences.find((o) => o.span.start <= i && i <= o.span.end);
    return occ ? { def, target: occ } : undefined;
  }

  /** The key of the defined name at `offset` (its definition or a use), if any. */
  keyAt(path: string, offset: number): string | undefined {
    const h = this.hit(path, offset);
    if (!h) return undefined;
    if (h.target === "self") return h.def.key;
    return h.target.kind === "use" ? h.target.key : undefined;
  }

  /** Go to definition (B3): a defined name → its entry; a LET/LAMBDA variable → its declaration. */
  definition(path: string, offset: number): Loc | undefined {
    const h = this.hit(path, offset);
    if (!h) return undefined;
    if (h.target === "self") return this.nameLoc(h.def);
    const t = h.target;
    if (t.kind === "use") {
      const d = t.key === undefined ? undefined : this.lookup(t.key);
      return d && this.nameLoc(d);
    }
    return this.loc(h.def, t.kind === "local" ? t.binder : t.span);
  }

  /** Every read of a local variable, with its declaration first. */
  localReferences(path: string, offset: number): Loc[] | undefined {
    const h = this.hit(path, offset);
    if (!h || h.target === "self" || h.target.kind === "use") return undefined;
    const binder = h.target.kind === "local" ? h.target.binder : h.target.span;
    const occ = this.analysis(h.def).occurrences;
    return [
      this.loc(h.def, binder),
      ...occ.filter((o) => o.kind === "local" && o.binder.start === binder.start).map((o) => this.loc(h.def, o.span)),
    ];
  }

  private usedByIndex(): Map<string, { def: NameDef; occ: Occurrence & { kind: "use" } }[]> {
    if (!this.usedByCache) {
      const m = new Map<string, { def: NameDef; occ: Occurrence & { kind: "use" } }[]>();
      for (const def of [...this.defs, ...this.cellDefs]) {
        for (const occ of this.analysis(def).occurrences) {
          if (occ.kind !== "use" || occ.key === undefined) continue;
          const k = lower(occ.key);
          let l = m.get(k);
          if (!l) m.set(k, (l = []));
          l.push({ def, occ });
        }
      }
      this.usedByCache = m;
    }
    return this.usedByCache;
  }

  /**
   * Rename Symbol (F2, M5): the defined name at `offset` (its statement or a use) and the
   * identifier's place, or why it cannot be renamed here.
   */
  renameTarget(path: string, offset: number): { def: NameDef; at: Loc } | string {
    const h = this.hit(path, offset);
    if (!h) return "Rename Symbol renames a defined name: put the cursor on one";
    if (h.target === "self") return { def: h.def, at: this.nameLoc(h.def) };
    const t = h.target;
    if (t.kind !== "use") return "a LET or LAMBDA variable: rename it by hand in its formula (Rename Symbol renames defined names)";
    const def = t.key === undefined ? undefined : this.lookup(t.key);
    if (!def) return `${t.id} is not a name of the project`;
    return { def, at: this.loc(h.def, t.span) };
  }

  /**
   * The edits renaming the name `key` to `to`: its statement with `@renamed(Old)` and every
   * reader in the project's files (core `renameInSource`, as `xln rename`). The next build
   * renames the name in the workbook and rewrites it in the cells.
   */
  renameEdits(key: string, to: string): SourceRename | string {
    return renameInSource(this, key, to, { ...(this.lock ? { lock: this.lock } : {}), tables: this.manifest?.tables.map((t) => t.name) ?? [] });
  }

  /** Every place in the sources that reads the name `key`. */
  references(key: string): Loc[] {
    return (this.usedByIndex().get(lower(key)) ?? []).map(({ def, occ }) => this.loc(def, occ.span));
  }

  /** "Used by" among names: definitions that read `key`, each with the places it does (unnamed cells left out). */
  usedBy(key: string): { def: NameDef; locs: Loc[] }[] {
    const out = new Map<NameDef, Loc[]>();
    for (const { def, occ } of this.usedByIndex().get(lower(key)) ?? []) {
      if (def.name === "") continue;
      let l = out.get(def);
      if (!l) out.set(def, (l = []));
      l.push(this.loc(def, occ.span));
    }
    return [...out].map(([def, locs]) => ({ def, locs }));
  }

  /** "Uses": the names `def` reads, each with the places it does. */
  uses(def: NameDef): { def: NameDef; locs: Loc[] }[] {
    const out = new Map<NameDef, Loc[]>();
    for (const occ of this.analysis(def).occurrences) {
      if (occ.kind !== "use" || occ.key === undefined) continue;
      const target = this.lookup(occ.key);
      if (!target) continue;
      let l = out.get(target);
      if (!l) out.set(target, (l = []));
      l.push(this.loc(def, occ.span));
    }
    return [...out].map(([d, locs]) => ({ def: d, locs }));
  }


  /** Problems of one file: the core checker's, the same `xln check` lists and the build refuses on. */
  problems(path: string): Problem[] {
    if (!this.files.has(path)) return [];
    return checkFile(this, path, this.checkContext());
  }

  /** B1: the outline of a file, names grouped under their `@scope`/`@workbook` block. */
  outline(path: string): OutlineSymbol[] {
    const file = this.files.get(path);
    if (!file) return [];
    const top: OutlineSymbol[] = [];
    let block: OutlineSymbol | undefined;
    const directives = [...file.parsed.scopes];
    const defs = [...this.defs, ...this.cellDefs].filter((d) => d.file === file).sort((a, b) => a.entry.start - b.entry.start);
    let k = 0;
    const openBlock = (s: (typeof directives)[number]): void => {
      const name = s.scope === undefined ? "@workbook" : `@scope(${s.scope})`;
      block = {
        name,
        detail: "",
        kind: "block",
        start: s.offset,
        end: s.end,
        selectionStart: s.offset,
        selectionEnd: s.end,
        children: [],
      };
      top.push(block);
    };
    for (const def of defs) {
      while (k < directives.length && directives[k]!.offset < def.entry.start) openBlock(directives[k++]!);
      const a = this.analysis(def);
      const sym: OutlineSymbol = {
        name: def.name !== "" ? def.name : `@${def.entry.cell!.range}`,
        detail: symbolDetail(a.classification, def.entry.formula),
        kind: a.classification.kind,
        start: def.entry.start,
        end: def.entry.end,
        selectionStart: def.entry.offset,
        selectionEnd: def.name !== "" ? def.entry.offset + def.name.length : def.entry.cell!.end,
        children: [],
      };
      if (block) {
        block.children.push(sym);
        block.end = Math.max(block.end, sym.end);
        block.detail = `${block.children.length} name${block.children.length === 1 ? "" : "s"}`;
      } else top.push(sym);
    }
    while (k < directives.length) openBlock(directives[k++]!);
    return top;
  }

  /** Folding (B1): multi-line definitions, doc comments, and `@scope` blocks, as 0-based lines. */
  folding(path: string): { start: number; end: number; kind?: "comment" }[] {
    const file = this.files.get(path);
    if (!file) return [];
    const out: { start: number; end: number; kind?: "comment" }[] = [];
    const line = (o: number) => file.lines.position(o).line;
    for (const s of this.outline(path)) {
      if (s.kind === "block" && line(s.end) > line(s.start)) out.push({ start: line(s.start), end: line(s.end) });
    }
    for (const def of [...this.defs, ...this.cellDefs]) {
      if (def.file !== file) continue;
      const e = def.entry;
      if (line(e.end) > line(e.offset)) out.push({ start: line(e.offset), end: line(e.end) });
      if (e.start < e.offset && line(e.offset) - 1 > line(e.start)) out.push({ start: line(e.start), end: line(e.offset) - 1, kind: "comment" });
    }
    return out;
  }

  /**
   * B2: names whose name, doc comment or definition contains every word of `query`
   * (case-insensitive). Name matches rank first, then doc comments, then definitions.
   */
  search(query: string, limit = 500): SearchHit[] {
    const words = lower(query).split(/\s+/).filter((w) => w !== "");
    const hits: (SearchHit & { rank: number })[] = [];
    for (const def of this.defs) {
      const fields = { name: lower(def.key), doc: lower(def.entry.doc ?? ""), formula: lower(def.entry.formula) };
      const matched = new Set<SearchHit["matched"][number]>();
      let all = true;
      for (const w of words) {
        let any = false;
        for (const f of ["name", "doc", "formula"] as const) {
          if (fields[f].includes(w)) {
            matched.add(f);
            any = true;
          }
        }
        if (!any) {
          all = false;
          break;
        }
      }
      if (!all) continue;
      const rank = words.length === 0 ? 0 : matched.has("name") ? (lower(def.name).startsWith(words[0]!) ? 0 : 1) : matched.has("doc") ? 2 : 3;
      hits.push({ def, matched: [...matched], rank });
    }
    hits.sort((a, b) => a.rank - b.rank || a.def.name.localeCompare(b.def.name));
    return hits.slice(0, limit).map(({ def, matched }) => ({ def, matched }));
  }
}

/** One line for the outline and search: the kind and the start of the formula. */
export function symbolDetail(c: Classification, formula: string): string {
  const flat = formula.replace(/\s+/g, " ");
  const head = flat.length > 60 ? flat.slice(0, 59) + "…" : flat;
  if (c.kind === "lambda") return `λ(${(c.params ?? []).join(", ")})`;
  return head;
}

/** `names/…/x.xln` path of a file below the project root; undefined for other files. */
export function isNamesFile(path: string): boolean {
  return path.startsWith(NAMES_DIR + "/") && path.toLowerCase().endsWith(".xln");
}
