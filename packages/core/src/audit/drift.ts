// C11: copy drift in families of names that differ only by a coordinate tag. The members
// of `Sales_base`, `Sales_payout`, `Sales_deleverage` are one line item copied per
// scenario: their formulas should be the same up to the tag. A name over a cell or spill
// is compared through the formula at its anchor (relative references as offsets from
// it), any other name through its definition. The members that differ from the majority
// are reported, with the tokens where they differ. Constants (scenario inputs) are not compared.
// Drift is one copy departing from copies that otherwise agree: the census uses the same
// comparison (`familyAgreement`) to drop an inferred tag whose families all differ.

import { formatCell, parseCell } from "../file/cellref.js";
import type { DefinedName } from "../file/types.js";
import type { DependencyGraph } from "../graph/model.js";
import { significant, tokenize, type Token } from "../lang/tokens.js";
import { decompile } from "../lang/transform.js";
import { definitionTarget } from "../project/classify.js";
import type { Classification } from "../project/types.js";
import { layoutToLf } from "../project/pull.js";
import { addressLabel, sheetLabel } from "../view/render.js";
import { nameWhere, scopeOf } from "./sites.js";
import type { Finding, NameCensus, NameFamily } from "./types.js";

const MARK = "◇";
const MAX_SNIPPET = 60;

interface Body {
  key: string;
  name: string;
  tag: string;
  formula: string;
  tokens: Token[];
  norms: string[];
  /** `IS!C30#` when the formula is the anchor's, not the definition. */
  at: string | undefined;
  def: DefinedName;
}

function replaceTag(id: string, tags: ReadonlySet<string>, position: NameFamily["position"]): string {
  const low = id.toLowerCase();
  if (position === "suffix") {
    const i = low.lastIndexOf("_");
    if (i <= 0) return low;
    const tail = low.slice(i + 1);
    if (tags.has(tail)) return low.slice(0, i + 1) + MARK;
    // `X_baseExp`: the coordinate followed by more name.
    for (const t of tags) if (tail.startsWith(t) && tail.length > t.length) return low.slice(0, i + 1) + MARK + tail.slice(t.length);
  } else {
    const i = low.indexOf("_");
    if (i > 0 && tags.has(low.slice(0, i))) return MARK + low.slice(i);
  }
  return low;
}

/** `C6` read from the anchor `C5` → `R[1]C[0]`; `$C$6` stays absolute. */
function relativeAddress(address: string, anchor: { row: number; col: number }): string {
  return address
    .split(":")
    .map((part) => {
      const absCol = part.startsWith("$");
      const body = absCol ? part.slice(1) : part;
      const dollar = body.indexOf("$");
      const absRow = dollar > 0;
      const a = parseCell(part.split("$").join(""));
      if (!a) return part;
      const r = absRow ? `R${a.row}` : `R[${a.row - anchor.row}]`;
      const c = absCol ? `C${a.col}` : `C[${a.col - anchor.col}]`;
      return r + c;
    })
    .join(":");
}

function normalise(t: Token, tags: ReadonlySet<string>, position: NameFamily["position"], anchor: { row: number; col: number } | undefined): string {
  const q = t.qual?.raw ?? "";
  switch (t.kind) {
    case "name":
      return q + replaceTag(t.value ?? t.text, tags, position);
    case "ref":
      return anchor && (t.refKind === "cell" || t.refKind === "area") ? q + relativeAddress(t.value ?? t.text, anchor) : t.text;
    case "structref": {
      let inner = t.inner ?? "";
      // `tbl[base]`: the column is the coordinate; `tbl[[#This Row],[base]]`: one of its parts.
      if (tags.has(inner.trim().toLowerCase())) inner = MARK;
      for (const tag of tags) {
        const want = `[${tag}]`;
        let k = inner.toLowerCase().indexOf(want);
        while (k >= 0) {
          inner = inner.slice(0, k) + `[${MARK}]` + inner.slice(k + want.length);
          k = inner.toLowerCase().indexOf(want, k + 3);
        }
      }
      return `${replaceTag(t.value ?? "", tags, position)}[${inner}]`;
    }
    case "string":
      return tags.has((t.value ?? "").toLowerCase()) ? `"${MARK}"` : t.text;
    default:
      return t.text;
  }
}

function snippet(b: Body, from: number, to: number): string {
  if (from >= to) return "nothing";
  const text = b.formula.slice(b.tokens[from]!.start, b.tokens[to - 1]!.end).replace(/\s+/g, " ");
  return "`" + (text.length > MAX_SNIPPET ? text.slice(0, MAX_SNIPPET - 1) + "…" : text) + "`";
}

/** What a family's formulas are read from. */
export interface DriftContext {
  defs: ReadonlyMap<string, DefinedName>;
  classes: ReadonlyMap<string, Classification>;
  localNames: ReadonlyMap<string, string[]>;
  graph: DependencyGraph;
}

/** The members of a family that can be compared, their tokens normalised up to the tags. */
function familyBodies(fam: NameFamily, tags: ReadonlySet<string>, ctx: DriftContext): Body[] {
  const { defs, classes, localNames, graph } = ctx;
  const bodies: Body[] = [];
  fam.members.forEach((key, k) => {
    const def = defs.get(key);
    // Inputs differ by design (`Volume_base = 100`, `Volume_high = 120`): only formulas are compared.
    if (!def || classes.get(key)?.kind === "constant") return;
    const scope = scopeOf(def);
    let formula: string | undefined;
    let anchor: { row: number; col: number } | undefined;
    let at: string | undefined;
    const t = definitionTarget(def.definition);
    if (t) {
      const sheet = t.sheet ?? scope;
      const node = sheet !== undefined ? graph.formulaAt(sheet, formatCell({ row: t.r1, col: t.c1 })) : undefined;
      const line = node?.line;
      if (!line || line.row !== t.r1 || line.col !== t.c1 || line.error !== undefined || line.formula === "") return; // over input cells
      formula = line.formula;
      anchor = { row: line.row, col: line.col };
      at = `${sheetLabel(line.sheet)}!${addressLabel(line)}`;
    } else {
      try {
        const home = scope === undefined ? {} : { homeSheet: scope, localNames: localNames.get(scope.toLowerCase()) ?? [] };
        formula = layoutToLf(decompile(def.definition, home));
      } catch {
        return; // C1 reports it
      }
    }
    const tokens = significant(tokenize(formula)).filter((x) => x.kind !== "eof");
    const norms = tokens.map((x) => normalise(x, tags, fam.position, anchor));
    bodies.push({ key, name: def.name, tag: fam.tags[k]!, formula, tokens, norms, at, def });
  });
  return bodies;
}

/** Bodies grouped by their formula up to the tags. */
function bySignature(bodies: readonly Body[]): Map<string, Body[]> {
  const groups = new Map<string, Body[]>();
  for (const b of bodies) {
    const sig = b.norms.join(" ");
    let g = groups.get(sig);
    if (!g) groups.set(sig, (g = []));
    g.push(b);
  }
  return groups;
}

/**
 * The census's test of an inferred tag (`FamilyAgreement`): two members with the same
 * formula up to the tags agree; members that all differ do not; fewer than two
 * comparable members (constants, input cells) say nothing.
 */
export function familyAgreement(ctx: DriftContext): (fam: NameFamily, tags: ReadonlySet<string>) => boolean | undefined {
  return (fam, tags) => {
    const bodies = familyBodies(fam, tags, ctx);
    if (bodies.length < 2) return undefined;
    return bySignature(bodies).size < bodies.length;
  };
}

export function copyDrift(census: NameCensus, ctx: DriftContext): Finding[] {
  const out: Finding[] = [];
  const tagsBy = {
    suffix: new Set(census.coordinates.filter((c) => c.position === "suffix").map((c) => c.tag.toLowerCase())),
    prefix: new Set(census.coordinates.filter((c) => c.position === "prefix").map((c) => c.tag.toLowerCase())),
  };
  for (const fam of census.families) {
    const bodies = familyBodies(fam, tagsBy[fam.position], ctx);
    if (bodies.length < 2) continue;
    const groups = bySignature(bodies);
    if (groups.size < 2) continue;
    // The majority is the reference; on a tie, the group of the family's first member.
    let ref: Body[] = [];
    for (const g of groups.values()) if (g.length > ref.length) ref = g;
    const rep = ref[0]!;
    for (const b of bodies) {
      if (ref.includes(b)) continue;
      let i = 0;
      while (i < b.norms.length && i < rep.norms.length && b.norms[i] === rep.norms[i]) i++;
      let j = 0;
      while (j < b.norms.length - i && j < rep.norms.length - i && b.norms[b.norms.length - 1 - j] === rep.norms[rep.norms.length - 1 - j]) j++;
      const here = snippet(b, i, b.norms.length - j);
      const there = snippet(rep, i, rep.norms.length - j);
      out.push({
        check: "C11",
        rule: "C11.drift",
        severity: "warning",
        where: nameWhere(b.def),
        message: "",
        data: {
          name: b.name,
          at: b.at ? ` (its formula at ${b.at})` : "",
          others: ref.map((x) => x.name).join(", "),
          position: fam.position,
          tag: b.tag,
          otherTag: rep.tag,
          here,
          there,
          family: fam.members,
        },
      });
    }
  }
  return out;
}
