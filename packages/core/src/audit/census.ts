// C8: the name census. Counts by kind and by scope, the families of names that differ
// only by a coordinate tag (`Sales_base`, `Sales_payout`), and the tier count of names
// that stand in for dimensions the grid does not have.
//
// The tiers generalise `scope_census.py` / `census.py` (excel-models, SUMMARY §3.2),
// which read the author's prefixes (`FN.PICK`, `IN.*`, `T.*`) and scenario tags. Here
// every rule reads what a name *is* and how it is used, so it applies to any workbook:
//
//   T1 coordinate addressing  LAMBDAs that turn a coordinate into a runtime lookup: called
//                             with a text argument (`IN.SET("periods")`), or looking a
//                             parameter up (`XLOOKUP(key, …)`, MATCH, XMATCH, VLOOKUP, …)
//   T4 the axis itself        names whose value is a SEQUENCE (directly or through a
//                             LAMBDA: the header row), and the non-LAMBDA names it is built from
//   T2 line-item identity     every other name, counted once per item: per (scope, stem),
//                             where `Sales_base` and `Sales_payout` share the stem `Sales`
//   T3 axis duplication       the line-item names beyond the first of each item
//   not counted               the other LAMBDAs (a library), and names left out by
//                             `exclude` (a check harness, solver settings)

import type { WorkbookSnapshot } from "../file/types.js";
import type { DependencyGraph } from "../graph/model.js";
import { leftSpine, children, type Expr } from "../lang/ast.js";
import { stripPrefix, tryParse } from "../lang/parser.js";
import { definitionTarget } from "../project/classify.js";
import { nameKey, type Classification, type NameKind } from "../project/types.js";
import { formatCell } from "../file/cellref.js";
import type { NameIndex } from "../view/formulas.js";
import { auditedNames, scopeOf } from "./sites.js";
import type { CensusOptions, NameCensus, NameFamily } from "./types.js";
import { cmp, unwrap } from "./walk.js";

export const CENSUS_EXPLANATION = [
  "Names standing in for dimensions (the tiers of FEASIBILITY §11.2, made generic):",
  "T1 coordinate addressing: LAMBDAs that turn a coordinate into a runtime lookup (called with a text argument, or looking a parameter up with XLOOKUP, MATCH, ...);",
  "T2 line-item identity: one per line item, i.e. per (scope, stem) where names differing only by a coordinate tag (Sales_base, Sales_payout) share a stem;",
  "T3 axis duplication: the names of an item beyond its first, the same item re-spelled per coordinate;",
  "T4 the axis itself: names whose value is a SEQUENCE (the header row) and the names it is built from.",
  "Not counted: the other LAMBDAs (a library) and names left out on purpose (the harness: xln.config.json, --census-exclude).",
  "A coordinate tag is a suffix or prefix after or before '_' that at least two families share, unless the formulas differ in every family carrying it (initial_amount, final_amount: separate items, not copies); sheet duplicates are short names defined on several sheets, the sheet standing in for a dataset.",
].join(" ");

const LOOKUPS = new Set(["XLOOKUP", "XMATCH", "MATCH", "VLOOKUP", "HLOOKUP", "LOOKUP"]);
/** Functions that turn a SEQUENCE into the labels of an axis: dates, years, text. */
const AXIS_WRAPPERS = new Set(["EOMONTH", "EDATE", "DATE", "YEAR", "MONTH", "WORKDAY", "TEXT", "DATEVALUE"]);

/** Case-insensitive glob (`*` only) on a name key. */
export function globMatch(pattern: string, text: string): boolean {
  const p = pattern.toLowerCase();
  const t = text.toLowerCase();
  const parts = p.split("*");
  if (parts.length === 1) return p === t;
  if (!t.startsWith(parts[0]!) || !t.endsWith(parts[parts.length - 1]!)) return false;
  let at = parts[0]!.length;
  const end = t.length - parts[parts.length - 1]!.length;
  for (const mid of parts.slice(1, -1)) {
    const k = t.indexOf(mid, at);
    if (k < 0 || k + mid.length > end) return false;
    at = k + mid.length;
  }
  return at <= end;
}

interface Split {
  stem: string;
  tag: string;
}

function splitSuffix(name: string): Split | undefined {
  const i = name.lastIndexOf("_");
  if (i <= 0 || i === name.length - 1 || name.lastIndexOf(".") > i) return undefined;
  return { stem: name.slice(0, i), tag: name.slice(i + 1) };
}

function splitPrefix(name: string): Split | undefined {
  if (name.includes(".")) return undefined;
  const i = name.indexOf("_");
  if (i <= 0 || i === name.length - 1) return undefined;
  return { tag: name.slice(0, i), stem: name.slice(i + 1) };
}

interface Member {
  key: string;
  name: string;
  scope: string | undefined;
}

/**
 * Whether the members of a family agree up to the tag: `true` when two of them have the
 * same formula modulo the tags, `false` when every comparable member differs, `undefined`
 * when fewer than two can be compared (constants, input cells). C11's comparison.
 */
export type FamilyAgreement = (family: NameFamily, tags: ReadonlySet<string>) => boolean | undefined;

/**
 * Families of names differing only by a coordinate tag, suffix first, then prefix for the
 * names left. With `agrees`, an inferred tag (not one in `opts.tags`) stays a coordinate
 * only if a family carrying it agrees, or no family carrying it can be compared.
 */
export function nameFamilies(members: readonly Member[], opts: CensusOptions = {}, agrees?: FamilyAgreement): { families: NameFamily[]; coordinates: NameCensus["coordinates"] } {
  const minFamilies = opts.minFamilies ?? 2;
  const forced = opts.tags ? new Set(opts.tags.map((t) => t.toLowerCase())) : undefined;
  const families: NameFamily[] = [];
  const coordinates: NameCensus["coordinates"] = [];
  const taken = new Set<string>();
  for (const position of ["suffix", "prefix"] as const) {
    const split = position === "suffix" ? splitSuffix : splitPrefix;
    const parts = members
      .filter((m) => !taken.has(m.key))
      .map((m) => ({ m, s: split(m.name) }))
      .filter((x): x is { m: Member; s: Split } => x.s !== undefined);
    const famKey = (x: { m: Member; s: Split }) => `${(x.m.scope ?? "").toLowerCase()}!${x.s.stem.toLowerCase()}`;
    const group = (tags: ReadonlySet<string>) => {
      const by = new Map<string, { m: Member; s: Split }[]>();
      for (const x of parts) {
        if (!tags.has(x.s.tag.toLowerCase())) continue;
        const k = famKey(x);
        let g = by.get(k);
        if (!g) by.set(k, (g = []));
        if (!g.some((y) => y.s.tag.toLowerCase() === x.s.tag.toLowerCase())) g.push(x);
      }
      return [...by.values()].filter((g) => g.length >= 2);
    };
    const familiesOf = (tags: ReadonlySet<string>) => {
      const groups = group(tags);
      if (position === "suffix") {
        // `X_baseExp`, `X_payoutExp`: the coordinate inside the last part; the family's stem is `X_*Exp`.
        const left = parts.filter((x) => !tags.has(x.s.tag.toLowerCase()));
        const by = new Map<string, { m: Member; s: Split }[]>();
        for (const x of left) {
          const tail = x.s.tag.toLowerCase();
          let best = "";
          for (const t of tags) if (tail.startsWith(t) && tail.length > t.length && t.length > best.length) best = t;
          if (!best) continue;
          const s: Split = { stem: `${x.s.stem}_*${x.s.tag.slice(best.length)}`, tag: x.s.tag.slice(0, best.length) };
          const k = `${(x.m.scope ?? "").toLowerCase()}!${s.stem.toLowerCase()}`;
          let g = by.get(k);
          if (!g) by.set(k, (g = []));
          if (!g.some((y) => y.s.tag.toLowerCase() === s.tag.toLowerCase())) g.push({ m: x.m, s });
        }
        for (const g of by.values()) if (g.length >= 2) groups.push(g);
      }
      for (const g of groups) g.sort((a, b) => cmp(a.m.key.toLowerCase(), b.m.key.toLowerCase()));
      return groups;
    };
    let tags: Set<string>;
    if (forced) tags = forced;
    else {
      // Candidates: tags on two stems or more; coordinates: tags that families share.
      const stems = new Map<string, Set<string>>();
      for (const x of parts) {
        const t = x.s.tag.toLowerCase();
        let s = stems.get(t);
        if (!s) stems.set(t, (s = new Set()));
        s.add(famKey(x));
      }
      const candidates = new Set([...stems].filter(([, s]) => s.size >= 2).map(([t]) => t));
      const count = new Map<string, number>();
      for (const g of group(candidates)) for (const x of g) count.set(x.s.tag.toLowerCase(), (count.get(x.s.tag.toLowerCase()) ?? 0) + 1);
      tags = new Set([...count].filter(([, n]) => n >= minFamilies).map(([t]) => t));
      // `baseExp` is the coordinate `base` followed by more name (`BSdelta_baseExp`): not a coordinate of its own.
      for (const t of [...tags]) if ([...tags].some((u) => u !== t && t.startsWith(u))) tags.delete(t);
    }
    const asFamily = (g: { m: Member; s: Split }[]): NameFamily => ({ scope: g[0]!.m.scope, stem: g[0]!.s.stem, position, members: g.map((x) => x.m.key), tags: g.map((x) => x.s.tag) });
    let groups = familiesOf(tags);
    if (!forced && agrees) {
      // Copy drift is one copy departing from copies that otherwise agree. When every
      // comparable family carrying an inferred tag differs (`initial_amount`,
      // `final_amount`: an opening and a closing balance), the tag is not a coordinate and
      // its names are separate line items. Dropping a tag can leave another one on fewer
      // than `minFamilies` families: repeat until nothing changes.
      for (;;) {
        const agreeing = new Set<string>();
        const differing = new Set<string>();
        const count = new Map<string, number>();
        for (const g of groups) {
          const verdict = agrees(asFamily(g), tags);
          for (const x of g) {
            const t = x.s.tag.toLowerCase();
            count.set(t, (count.get(t) ?? 0) + 1);
            if (verdict === true) agreeing.add(t);
            else if (verdict === false) differing.add(t);
          }
        }
        const drop = [...tags].filter((t) => (differing.has(t) && !agreeing.has(t)) || (count.get(t) ?? 0) < minFamilies);
        if (drop.length === 0) break;
        for (const t of drop) tags.delete(t);
        groups = familiesOf(tags);
      }
    }
    const perTag = new Map<string, { tag: string; n: number }>();
    for (const g of groups) {
      families.push(asFamily(g));
      for (const x of g) {
        taken.add(x.m.key);
        const t = x.s.tag.toLowerCase();
        const e = perTag.get(t);
        if (e) e.n++;
        else perTag.set(t, { tag: x.s.tag, n: 1 });
      }
    }
    for (const { tag, n } of perTag.values()) coordinates.push({ tag, position, families: n });
  }
  coordinates.sort((a, b) => b.families - a.families || cmp(a.tag, b.tag));
  families.sort((a, b) => cmp(a.scope ?? "", b.scope ?? "") || cmp(a.stem.toLowerCase(), b.stem.toLowerCase()));
  return { families, coordinates };
}

export interface CensusInput {
  wb: WorkbookSnapshot;
  graph: DependencyGraph;
  names: NameIndex;
  classes: ReadonlyMap<string, Classification>;
  /** Keys of LAMBDA names called somewhere with a text literal among the arguments. */
  calledWithText: ReadonlySet<string>;
  opts: CensusOptions;
  /** C11's comparison of a family's formulas: an inferred tag whose families all differ is not a coordinate. */
  agrees?: FamilyAgreement;
}

export function nameCensus(input: CensusInput): NameCensus {
  const { wb, graph, names, classes, calledWithText, opts, agrees } = input;
  const defs = auditedNames(wb);
  const byKind: Record<NameKind, number> = { constant: 0, range: 0, spill: 0, table: 0, formula: 0, lambda: 0, unparsed: 0 };
  const perSheet = new Map<string, number>();
  let workbook = 0;
  const members: Member[] = [];
  for (const d of defs) {
    const scope = scopeOf(d);
    const key = nameKey({ name: d.name, scope });
    byKind[classes.get(key)?.kind ?? "unparsed"]++;
    if (scope === undefined) workbook++;
    else perSheet.set(scope, (perSheet.get(scope) ?? 0) + 1);
    members.push({ key, name: d.name, scope });
  }
  const sheets = wb.sheets.filter((s) => perSheet.has(s.name)).map((s) => ({ sheet: s.name, names: perSheet.get(s.name)! }));

  const exclude = opts.exclude ?? [];
  const excluded = new Set(members.filter((m) => exclude.some((p) => globMatch(p, m.key))).map((m) => m.key));
  const lambdas = new Set(members.filter((m) => classes.get(m.key)?.kind === "lambda").map((m) => m.key));

  // ---- T1: accessors -----------------------------------------------------------------
  const lambdaOf = new Map<string, { params: Set<string>; body: Expr }>();
  for (const d of defs) {
    const key = nameKey({ name: d.name, scope: scopeOf(d) });
    if (!lambdas.has(key)) continue;
    const body = tryParse(d.definition).formula?.body;
    const l = body && unwrap(body);
    if (l?.kind === "lambda") lambdaOf.set(key, { params: new Set(l.params.map((p) => stripPrefix(p.name.text).base.toLowerCase())), body: l.body });
  }
  const accessors = new Set<string>();
  for (const [key, l] of lambdaOf) {
    if (excluded.has(key)) continue;
    if (calledWithText.has(key) || looksUpParameter(l.body, l.params)) accessors.add(key);
  }

  // ---- T4: the axis ------------------------------------------------------------------
  const lambdaBody = (id: string, sheet: string | undefined, home: string | undefined): Expr | undefined => {
    const key = names.resolve({ id, sheet, span: { start: 0, end: 0 } }, home);
    return key === undefined ? undefined : lambdaOf.get(key)?.body;
  };
  // The axis: a SEQUENCE, shifted (`start + SEQUENCE(…)`), turned into dates or labels
  // (`EOMONTH(start, SEQUENCE(…))`), or made by a LAMBDA that returns one.
  const isSequence = (e: Expr, home: string | undefined, depth = 0): boolean => {
    const x = unwrap(e);
    if (depth > 8) return false;
    if (x.kind === "binary" && (x.op === "+" || x.op === "-")) return isSequence(x.left, home, depth + 1) !== isSequence(x.right, home, depth + 1);
    if (x.kind !== "call") return false;
    const { prefix, base: name } = stripPrefix(x.fn.text);
    if (prefix === "_xludf.") return false; // #NAME?: no axis
    const base = name.toUpperCase();
    if (base === "SEQUENCE") return true;
    if (AXIS_WRAPPERS.has(base)) return x.args.some((a) => isSequence(a, home, depth + 1));
    const body = lambdaBody(x.fn.text, x.fn.qual?.sheet, home);
    return body !== undefined && isSequence(body, home, depth + 1);
  };
  const axis = new Set<string>();
  for (const d of defs) {
    const scope = scopeOf(d);
    const key = nameKey({ name: d.name, scope });
    if (lambdas.has(key) || excluded.has(key)) continue;
    let value: { body: Expr; home: string | undefined } | undefined;
    const t = definitionTarget(d.definition);
    if (t) {
      const sheet = t.sheet ?? scope;
      const line = sheet !== undefined ? graph.formulaAt(sheet, formatCell({ row: t.r1, col: t.c1 }))?.line : undefined;
      const body = line && line.row === t.r1 && line.col === t.c1 ? tryParse(line.formula).formula?.body : undefined;
      if (body) value = { body, home: line!.sheet };
    } else {
      const body = tryParse(d.definition).formula?.body;
      if (body) value = { body, home: scope };
    }
    if (value && isSequence(value.body, value.home)) axis.add(key);
  }
  for (const key of [...axis]) {
    const node = graph.nameNode(key);
    if (!node) continue;
    for (const p of graph.allPrecedents(node)) {
      if (p.kind !== "name" || !p.name!.defined) continue;
      const k = p.name!.key;
      if (!lambdas.has(k) && !excluded.has(k) && !accessors.has(k)) axis.add(k);
    }
  }

  // ---- families, T2, T3 ----------------------------------------------------------------
  const { families, coordinates } = nameFamilies(members, opts, agrees);
  const stemOf = new Map<string, string>();
  for (const f of families) for (const k of f.members) stemOf.set(k, f.stem);
  const items = new Set<string>();
  const spellings = new Set<string>();
  let lineItems = 0;
  const library: string[] = [];
  for (const m of members) {
    if (excluded.has(m.key) || accessors.has(m.key) || axis.has(m.key)) continue;
    if (lambdas.has(m.key)) {
      library.push(m.key);
      continue;
    }
    lineItems++;
    const stem = (stemOf.get(m.key) ?? m.name).toLowerCase();
    items.add(`${(m.scope ?? "").toLowerCase()}!${stem}`);
    spellings.add(stem);
  }
  const T1 = accessors.size;
  const T2 = items.size;
  const T3 = lineItems - items.size;
  const T4 = axis.size;
  const total = T1 + T2 + T3 + T4;

  const scopesByName = new Map<string, { name: string; sheets: string[] }>();
  for (const m of members) {
    if (m.scope === undefined) continue;
    const k = m.name.toLowerCase();
    let e = scopesByName.get(k);
    if (!e) scopesByName.set(k, (e = { name: m.name, sheets: [] }));
    e.sheets.push(m.scope);
  }
  const position = new Map(wb.sheets.map((s) => [s.name, s.position]));
  const sheetDuplicates = [...scopesByName.values()]
    .filter((e) => e.sheets.length > 1)
    .map((e) => ({ name: e.name, sheets: e.sheets.sort((a, b) => (position.get(a) ?? 0) - (position.get(b) ?? 0)) }))
    .sort((a, b) => cmp(a.name.toLowerCase(), b.name.toLowerCase()));

  const byKey = (a: string, b: string) => cmp(a.toLowerCase(), b.toLowerCase()) || cmp(a, b);
  return {
    total: members.length,
    hidden: defs.filter((d) => d.hidden).length,
    byKind,
    byScope: { workbook, sheets },
    builtIns: wb.definedNames.filter((d) => d.isBuiltIn).length,
    coordinates,
    families,
    tiers: {
      T1,
      T2,
      T3,
      T4,
      total,
      of: members.length,
      percent: members.length === 0 ? 0 : Math.round((1000 * total) / members.length) / 10,
      T2spellings: spellings.size,
      library: library.length,
      excluded: excluded.size,
    },
    tierNames: {
      T1: [...accessors].sort(byKey),
      T4: [...axis].sort(byKey),
      library: library.sort(byKey),
      excluded: [...excluded].sort(byKey),
    },
    sheetDuplicates,
    explanation: CENSUS_EXPLANATION,
  };
}

/** A parameter used as the value looked up: `XLOOKUP(key, …)`, `MATCH(key, …)`. */
function looksUpParameter(body: Expr, params: ReadonlySet<string>): boolean {
  let hit = false;
  const visit = (n: Expr): void => {
    if (hit) return;
    if (n.kind === "call") {
      const base = stripPrefix(n.fn.text).base.toUpperCase();
      const first = n.args[0] && unwrap(n.args[0]);
      if (LOOKUPS.has(base) && first?.kind === "name" && !first.qual && params.has(stripPrefix(first.id).base.toLowerCase())) {
        hit = true;
        return;
      }
    }
    if (n.kind === "binary") {
      const { first, links } = leftSpine(n);
      visit(first);
      for (const b of links) visit(b.right);
      return;
    }
    for (const c of children(n)) visit(c);
  };
  visit(body);
  return hit;
}
