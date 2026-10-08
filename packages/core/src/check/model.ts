// The source of an xln project as the editor and the build both read it (M3d): its `.xln`
// files parsed (`parseSourceFile`), each definition's formula parsed, and every identifier
// resolved the way Excel resolves it. Positions are offsets into a file's text. The
// extension's project model extends this with the manifest (go to definition, usages,
// outline); the checker (check.ts) reads it as it is.
//
// Names are resolved from the sources alone, not from the workbook: after an edit,
// resolution follows the text.

import { leftSpine, type Expr } from "../lang/ast.js";
import { lookupFunction } from "../lang/catalogue.js";
import { stripPrefix, tryParse } from "../lang/parser.js";
import type { Qualifier, Span } from "../lang/tokens.js";
import { classify } from "../project/classify.js";
import { formulaToSource, parseSourceFile, sourceToFormula, type ModuleEntry, type ParsedModule } from "../project/module.js";
import { NameResolver } from "../project/refs.js";
import type { Classification } from "../project/types.js";
import { LineIndex } from "./lines.js";

export interface SourceFile {
  /** Project-relative, forward slashes: `names/_unmanaged.xln`. */
  path: string;
  text: string;
  parsed: ParsedModule;
  lines: LineIndex;
}

/** A definition in an `.xln` file (or an unnamed cell statement, `name` ""). */
export interface NameDef {
  /** `Sheet!Name` or `Name`, as written in the source (`@Sheet!C5` for an unnamed cell). */
  key: string;
  name: string;
  scope: string | undefined;
  file: SourceFile;
  entry: ModuleEntry;
}

/** A place in a project file, as offsets. */
export interface Loc {
  path: string;
  start: number;
  end: number;
}

/** What sits at a spot of a definition's formula. */
export type Occurrence =
  /** A defined name read here; `key` undefined when no name of that spelling exists. */
  | { kind: "use"; span: Span; id: string; sheet: string | undefined; key: string | undefined; call: boolean }
  /** A LET or LAMBDA variable read here; `binder` is where it is declared. */
  | { kind: "local"; span: Span; id: string; binder: Span }
  /** The declaration of a LET or LAMBDA variable. */
  | { kind: "binder"; span: Span; id: string };

export interface Analysis {
  classification: Classification;
  /** Formula parse errors, as formula-text spans. */
  errors: { span: Span; message: string }[];
  occurrences: Occurrence[];
  /** The formula's syntax tree; undefined when it does not parse (or for a slot). */
  body: Expr | undefined;
}

const lower = (s: string) => s.toLowerCase();

export function defKey(name: string, scope: string | undefined): string {
  return scope === undefined ? name : `${scope}!${name}`;
}

function external(q: Qualifier | undefined): boolean {
  return q !== undefined && (q.book !== undefined || q.sheet === undefined || q.sheet2 !== undefined);
}

/** The identifier at the end of a node's span: a qualifier comes before it. */
function idSpan(span: Span, id: string): Span {
  return { start: span.end - id.length, end: span.end };
}

/**
 * Every identifier of a formula with what it means. Mirrors the core's `nameUses` (the
 * manifest's view), and adds the LET/LAMBDA variables it skips, with their declarations.
 */
export function occurrences(body: Expr, resolve: (id: string, sheet: string | undefined) => string | undefined): Occurrence[] {
  const out: Occurrence[] = [];
  type Scope = ReadonlyMap<string, Span>;
  const bind = (scope: Map<string, Span>, ident: { text: string; span: Span }): void => {
    const base = stripPrefix(ident.text).base;
    const span = idSpan(ident.span, base);
    scope.set(lower(base), span);
    out.push({ kind: "binder", span, id: base });
  };
  const ref = (id: string, qual: Qualifier | undefined, span: Span, scope: Scope, isCall: boolean): void => {
    if (external(qual)) return;
    const { prefix, base } = stripPrefix(id);
    const s = idSpan(span, id);
    if (prefix === "_xlpm." || prefix === "_xleta." || (!qual && scope.has(lower(base)))) {
      const binder = scope.get(lower(base));
      if (binder) out.push({ kind: "local", span: s, id: base, binder });
      return;
    }
    if (prefix !== "") return; // `_xlfn.` / `_xludf.` spellings are functions
    if (isCall && !qual && lookupFunction(base) !== undefined) return; // built-ins win (T12)
    out.push({ kind: "use", span: s, id, sheet: qual?.sheet, key: resolve(id, qual?.sheet), call: isCall });
  };
  const visit = (node: Expr, scope: Scope): void => {
    switch (node.kind) {
      case "name":
        ref(node.id, node.qual, node.span, scope, false);
        return;
      case "call":
        ref(node.fn.text, node.fn.qual, node.fn.span, scope, true);
        for (const a of node.args) visit(a, scope);
        return;
      case "lambda": {
        const inner = new Map(scope);
        for (const p of node.params) bind(inner, p.name);
        visit(node.body, inner);
        return;
      }
      case "let": {
        const inner = new Map(scope);
        for (const b of node.bindings) {
          visit(b.value, new Map(inner));
          bind(inner, b.name);
        }
        visit(node.body, inner);
        return;
      }
      case "binary": {
        const { first, links } = leftSpine(node);
        visit(first, scope);
        for (const b of links) visit(b.right, scope);
        return;
      }
      case "array":
        for (const r of node.rows) for (const e of r) visit(e, scope);
        return;
      case "paren":
        visit(node.expr, scope);
        return;
      case "unary":
      case "postfix":
        visit(node.operand, scope);
        return;
      case "invoke":
        visit(node.callee, scope);
        for (const a of node.args) visit(a, scope);
        return;
      default:
        return;
    }
  };
  visit(body, new Map());
  return out;
}

/** Names in scope: every name once (the first definition of a key wins), local ones by lower-case sheet. */
export interface ScopeIndex {
  locals: Map<string, NameDef[]>;
  workbook: NameDef[];
}

export class SourceModel {
  readonly files = new Map<string, SourceFile>();
  private defsCache: NameDef[] | undefined;
  private cellDefsCache: NameDef[] | undefined;
  private byKey: Map<string, NameDef> | undefined;
  private resolver: NameResolver | undefined;
  private scopeCache: ScopeIndex | undefined;
  private readersCache: Map<string, NameDef[]> | undefined;
  /** Bumped whenever a file changes: resolution depends on every file. */
  protected generation = 0;
  private readonly analyses = new WeakMap<ModuleEntry, { generation: number; analysis: Analysis }>();
  private readonly parses = new WeakMap<ModuleEntry, { body: Expr | undefined; classification: Classification; errors: Analysis["errors"] }>();

  setFile(path: string, text: string): void {
    const old = this.files.get(path);
    if (old && old.text === text) return;
    this.files.set(path, { path, text, parsed: parseSourceFile(path, text), lines: new LineIndex(text) });
    this.invalidate();
  }

  deleteFile(path: string): void {
    if (this.files.delete(path)) this.invalidate();
  }

  /** Drops what depends on every file; a subclass with caches of its own extends it. */
  protected invalidate(): void {
    this.cellDefsCache = undefined;
    this.defsCache = undefined;
    this.byKey = undefined;
    this.resolver = undefined;
    this.scopeCache = undefined;
    this.readersCache = undefined;
    this.generation++;
  }

  /** Every definition, files in path order; a key defined twice appears twice. */
  get defs(): NameDef[] {
    if (!this.defsCache) {
      const out: NameDef[] = [];
      for (const path of [...this.files.keys()].sort()) {
        const file = this.files.get(path)!;
        for (const entry of file.parsed.entries) {
          out.push({ key: defKey(entry.name, entry.scope), name: entry.name, scope: entry.scope, file, entry });
        }
      }
      this.defsCache = out;
    }
    return this.defsCache;
  }

  /**
   * The unnamed cell statements (`@C5 = …;`, M3b), files in path order. They are not names:
   * `defs` and `lookup` leave them out, but their formulas are read like a definition's.
   * Key `@Sheet!C5`.
   */
  get cellDefs(): NameDef[] {
    if (!this.cellDefsCache) {
      const out: NameDef[] = [];
      for (const path of [...this.files.keys()].sort()) {
        const file = this.files.get(path)!;
        for (const entry of file.parsed.cells) {
          if (entry.name !== "") continue;
          out.push({ key: `@${entry.cellSheet ?? ""}!${entry.cell!.range}`, name: "", scope: entry.scope, file, entry });
        }
      }
      this.cellDefsCache = out;
    }
    return this.cellDefsCache;
  }

  /** The definition of a key (`Sheet!Name` or `Name`, any case); the first one when duplicated. */
  lookup(key: string): NameDef | undefined {
    if (!this.byKey) {
      this.byKey = new Map();
      for (const d of this.defs) if (!this.byKey.has(lower(d.key))) this.byKey.set(lower(d.key), d);
    }
    return this.byKey.get(lower(key));
  }

  /** The names in scope, the first definition of each key, local ones by sheet. */
  scopeIndex(): ScopeIndex {
    if (!this.scopeCache) {
      const scope: ScopeIndex = { locals: new Map(), workbook: [] };
      for (const d of this.defs) {
        if (this.lookup(d.key) !== d) continue;
        if (d.scope === undefined) scope.workbook.push(d);
        else {
          const k = lower(d.scope);
          let l = scope.locals.get(k);
          if (!l) scope.locals.set(k, (l = []));
          l.push(d);
        }
      }
      this.scopeCache = scope;
    }
    return this.scopeCache;
  }

  /** The sheets defining a local name `id`, `except` left out, in `order` (sheet names) where given. */
  sheetsWith(id: string, except: string | undefined, order: readonly string[] = []): string[] {
    const out: string[] = [];
    for (const [, list] of this.scopeIndex().locals) {
      for (const d of list) if (lower(d.name) === lower(id) && (except === undefined || lower(d.scope!) !== lower(except))) out.push(d.scope!);
    }
    const o = order.map(lower);
    const pos = (s: string) => {
      const k = o.indexOf(lower(s));
      return k < 0 ? o.length : k;
    };
    return out.sort((a, b) => pos(a) - pos(b));
  }

  /** The key a name read in a formula resolves to, the way Excel does (see `NameResolver`). */
  resolveName(id: string, sheet: string | undefined, home: string | undefined): string | undefined {
    if (!this.resolver) {
      this.resolver = new NameResolver(this.defs.map((d) => ({ name: d.name, scope: d.scope, key: d.key })));
    }
    const k = this.resolver.resolve({ id, sheet, span: { start: 0, end: 0 } }, home);
    return k === undefined ? undefined : this.lookup(k)?.key;
  }

  /** The definitions and cell statements whose formulas read the name `key` (any case). */
  readers(key: string): NameDef[] {
    if (!this.readersCache) {
      const map = new Map<string, NameDef[]>();
      for (const d of [...this.defs, ...this.cellDefs]) {
        const seen = new Set<string>();
        for (const o of this.analysis(d).occurrences) {
          if (o.kind !== "use" || o.key === undefined || seen.has(lower(o.key))) continue;
          seen.add(lower(o.key));
          const l = map.get(lower(o.key)) ?? [];
          l.push(d);
          map.set(lower(o.key), l);
        }
      }
      this.readersCache = map;
    }
    return this.readersCache.get(lower(key)) ?? [];
  }

  /** The sheet whose names a definition's formula reads bare: a cell's sheet, a local name's own. */
  homeOf(def: NameDef): string | undefined {
    return def.entry.cell ? (def.entry.cellSheet ?? def.scope) : def.scope;
  }

  /** The parsed formula of a definition: kind, errors and every identifier with its meaning. */
  analysis(def: NameDef): Analysis {
    const cached = this.analyses.get(def.entry);
    if (cached && cached.generation === this.generation) return cached.analysis;
    // Parsing depends on the entry alone and survives edits elsewhere; resolution does not.
    let parsed = this.parses.get(def.entry);
    if (!parsed) {
      const cell = def.entry.cell;
      if (cell && def.entry.formula === "") {
        // A slot (or a cell to clear): nothing to parse.
        parsed = { body: undefined, classification: this.cellClassification(def), errors: [] };
      } else {
        const { formula, diagnostics } = tryParse(def.entry.formula);
        parsed = {
          body: formula?.body,
          // A cell statement's name is its cell (a range or its spill); its formula is the cell's.
          classification: cell ? this.cellClassification(def) : classify(def.entry.formula),
          errors: diagnostics.filter((d) => d.severity === "error").map((d) => ({ span: { start: d.start, end: d.end }, message: d.message })),
        };
      }
      this.parses.set(def.entry, parsed);
    }
    const home = this.homeOf(def);
    const a: Analysis = {
      classification: parsed.classification,
      errors: parsed.errors,
      occurrences: parsed.body ? occurrences(parsed.body, (id, sheet) => this.resolveName(id, sheet, home)) : [],
      body: parsed.body,
    };
    this.analyses.set(def.entry, { generation: this.generation, analysis: a });
    return a;
  }

  /** What a cell statement's name is; the extension refines it from the manifest. */
  protected cellClassification(def: NameDef): Classification {
    return def.name === "" ? { kind: "formula" } : { kind: def.entry.cell?.spill ? "spill" : "range" };
  }

  /** Where a span of a definition's formula is in its file. */
  loc(def: NameDef, span: Span): Loc {
    const start = formulaToSource(def.entry, span.start);
    const end = span.end > span.start ? formulaToSource(def.entry, span.end - 1) + 1 : start;
    return { path: def.file.path, start, end };
  }

  /** Where a definition's name is written. */
  nameLoc(def: NameDef): Loc {
    return { path: def.file.path, start: def.entry.offset, end: def.entry.offset + def.name.length };
  }

  /** The definition (or unnamed cell statement) whose text (doc comment to `;`) contains `offset`. */
  defAt(path: string, offset: number): NameDef | undefined {
    const inside = (d: NameDef) => d.file.path === path && d.entry.start <= offset && offset <= d.entry.end;
    return this.defs.find(inside) ?? this.cellDefs.find(inside);
  }

  /** The formula index at a file offset inside a definition's formula text. */
  formulaIndexAt(def: NameDef, offset: number): number | undefined {
    return sourceToFormula(def.entry, offset);
  }
}
