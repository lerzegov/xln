// Rename across formulas (stretch G): rewrite the token of a renamed name in one formula,
// stored or display form, the way Excel's Name Manager rewrites its dependents (probe F7).
// A token substitution on the parsed formula, never a text search: strings, structured
// references, sheet names, LET/LAMBDA variables of the same spelling (`_xlpm.Rate`, or a
// `Rate` bound by a LET in display form) and names of other workbooks are left alone. A
// bare name is the local one of the formula's sheet when there is one, else the
// workbook's (`NameResolver`); `Sheet!Name` names the scope; `[0]!Name` is this
// workbook's name; case does not matter.
//
// After the edit the formula is parsed again and every name in it must reach what it
// reached before (the renamed one under its new name): a new name that a sheet's local
// name or a LET/LAMBDA variable would capture, or that an unknown name of that spelling
// would start to reach, is reported (`captured`), never written.

import { tryParse } from "../lang/parser.js";
import { tokenize, type Span } from "../lang/tokens.js";
import { NameResolver, nameUses, type NameUse } from "./refs.js";

/** A defined name the formulas can read: its spelling and its scope (a sheet, or undefined for workbook scope). */
export interface ScopedName {
  name: string;
  scope: string | undefined;
}

/** One rename: the name `from` in `scope` becomes `to`, same scope. */
export interface NameRename {
  scope: string | undefined;
  from: string;
  to: string;
}

const lower = (s: string) => s.toLowerCase();

/** lower-case `sheet!name`, sheet "" for workbook scope. */
export function scopedLowerKey(name: string, scope: string | undefined): string {
  return `${lower(scope ?? "")}!${lower(name)}`;
}

/** The names before and after a set of renames, to resolve a formula's names on both sides. */
export class RenameContext {
  readonly before: NameResolver;
  readonly after: NameResolver;
  /** Before key → the rename and its key after. */
  readonly renames = new Map<string, { rename: NameRename; afterKey: string }>();
  /** Old spellings, lower case (a quick test before parsing). */
  readonly oldIds = new Set<string>();
  /** Old and new spellings, lower case: a formula with neither cannot change. */
  readonly ids = new Set<string>();

  constructor(names: readonly ScopedName[], renames: readonly NameRename[]) {
    for (const r of renames) {
      const k = scopedLowerKey(r.from, r.scope);
      this.renames.set(k, { rename: r, afterKey: scopedLowerKey(r.to, r.scope) });
      this.oldIds.add(lower(r.from));
      this.ids.add(lower(r.from));
      this.ids.add(lower(r.to));
    }
    const keyed = (n: ScopedName) => ({ name: n.name, scope: n.scope, key: scopedLowerKey(n.name, n.scope) });
    this.before = new NameResolver(names.map(keyed));
    // After: each renamed name under its new spelling, keyed by its new key.
    this.after = new NameResolver(
      names.map((n) => {
        const r = this.renames.get(scopedLowerKey(n.name, n.scope));
        return r ? { name: r.rename.to, scope: n.scope, key: r.afterKey } : keyed(n);
      }),
    );
  }

  get empty(): boolean {
    return this.renames.size === 0;
  }
}

export interface FormulaRename {
  /** The formula with the renamed tokens replaced (the input when nothing changed or `captured`/`unparsed`). */
  text: string;
  /** Tokens replaced. */
  count: number;
  /** The spans replaced, in the input text. */
  spans: Span[];
  /** For each span, the renamed name's key before the rename (lower-case `sheet!name`). */
  keys: string[];
  /** Names whose reading the rename would change (`x` captured by a LET variable, a local name, …): nothing was replaced. */
  captured?: string[];
  /** The formula does not parse, yet holds an identifier spelled like a renamed name: nothing was replaced. */
  unparsed?: true;
}

/**
 * Rewrites the renamed names in one formula. `home` is the sheet the formula lives on (a
 * cell's, a conditional format's, a sheet-scoped name's own sheet); undefined for a
 * workbook-scoped name's definition.
 */
export function renameInFormula(text: string, home: string | undefined, ctx: RenameContext): FormulaRename {
  const none: FormulaRename = { text, count: 0, spans: [], keys: [] };
  if (ctx.empty) return none;
  // Cheap test first: no identifier spelled like an old or a new name, nothing can change.
  const toks = tokenize(text);
  const spelled = (set: ReadonlySet<string>) => toks.some((t) => t.kind === "name" && t.value !== undefined && set.has(lower(t.value)));
  if (!spelled(ctx.ids)) return none;
  const { formula } = tryParse(text);
  if (!formula) return spelled(ctx.oldIds) ? { ...none, unparsed: true } : none;

  const edits: { span: Span; text: string; key: string }[] = [];
  const uses = nameUses(formula.body);
  const expected: (string | undefined)[] = [];
  for (const u of uses) {
    const k = ctx.before.resolve(u, home);
    const r = k === undefined ? undefined : ctx.renames.get(k);
    if (r) {
      edits.push({ span: u.span, text: r.rename.to, key: k! });
      expected.push(r.afterKey);
    } else expected.push(k);
  }
  // `[0]!Name`: this workbook's workbook-scoped name (charts and links write it so).
  // `nameUses` leaves it out with the other workbooks' names.
  for (const t of toks) {
    if (t.kind !== "name" || t.qual?.book !== "[0]" || t.qual.sheet !== undefined || t.value === undefined) continue;
    const r = ctx.renames.get(scopedLowerKey(t.value, undefined));
    if (r) edits.push({ span: { start: t.end - t.value.length, end: t.end }, text: r.rename.to, key: scopedLowerKey(t.value, undefined) });
  }
  edits.sort((a, b) => a.span.start - b.span.start);
  let out = "";
  let at = 0;
  for (const e of edits) {
    out += text.slice(at, e.span.start) + e.text;
    at = e.span.end;
  }
  out += text.slice(at);

  // Every name must reach after the rename what it reached before.
  const again = tryParse(out).formula;
  const captured: string[] = [];
  if (!again) captured.push("(the formula no longer parses)");
  else {
    const after: NameUse[] = nameUses(again.body);
    if (after.length !== uses.length) {
      // A renamed use became a LET/LAMBDA variable of the new spelling.
      captured.push(...edits.map((e) => e.text));
    } else {
      after.forEach((u, i) => {
        if (ctx.after.resolve(u, home) !== expected[i]) captured.push(u.sheet === undefined ? u.id : `${u.sheet}!${u.id}`);
      });
    }
  }
  if (captured.length) return { ...none, captured: [...new Set(captured)] };
  if (edits.length === 0) return none;
  return { text: out, count: edits.length, spans: edits.map((e) => e.span), keys: edits.map((e) => e.key) };
}
