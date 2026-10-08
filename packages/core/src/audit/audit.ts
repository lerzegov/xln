// L3 audit: checks C1–C13 of the brief (§4 C) over a workbook snapshot, from the saved
// file alone; C14 (AFE's copy of the names) and C15 (labels that no longer give their
// names, from the cells' values when the caller passes them). Formula-level checks (C1 syntax, C2 prefixes, C6 arity, C7 limits, C5,
// C13) read every formula text once; reference-level checks (C4, C9, C10, C12) read the
// dependency graph; C8 and C11 read the families of names. What each finding says and
// how loud it is comes from the rule table (`rules.ts`).

import { formatCell, parseCell } from "../file/cellref.js";
import type { DefinedName, WorkbookSnapshot } from "../file/types.js";
import { buildGraph } from "../graph/graph.js";
import type { DependencyGraph, GraphNode } from "../graph/model.js";
import { walk } from "../lang/ast.js";
import { builtinCollision, lookupFunction } from "../lang/catalogue.js";
import { lineCol } from "../lang/errors.js";
import { stripPrefix, tryParse } from "../lang/parser.js";
import { quoteSheet } from "../lang/tokens.js";
import { decompileWithDiagnostics } from "../lang/transform.js";
import { cellCallee, classify, definitionTarget } from "../project/classify.js";
import { labelDrift } from "../project/labels.js";
import { labelReplacement, labelReplaceShort } from "../project/labelNotice.js";
import { nameUses } from "../project/refs.js";
import { nameKey, type Classification } from "../project/types.js";
import { workbookNameIndex, type FormulaViewLine } from "../view/formulas.js";
import { afeStatus, afeStoreLabel } from "./afe.js";
import { globMatch, nameCensus } from "./census.js";
import { copyDrift, familyAgreement } from "./drift.js";
import { CHECK_TITLES, fill, RULES } from "./rules.js";
import { auditedNames, collectSites, nameWhere, scopeOf, type Site } from "./sites.js";
import { CHECK_IDS, type AuditCounts, type AuditOptions, type AuditReport, type AuditSeverity, type CheckId, type Finding, type FindingWhere, type SpillCensus, type SpillEntry } from "./types.js";
import { callUses, cmp, nestingDepth, numbersIn, outerLambdas, unwrap } from "./walk.js";

/**
 * C13: numbers a LAMBDA body may hold without a finding (and their negatives). Identities, halves,
 * powers of ten, and calendar counts (days in a week and year, months, hours, minutes).
 */
export const DEFAULT_CONSTANTS: readonly number[] = [-1, 0, 1, 2, 0.5, 10, 100, 1000, 7, 12, 24, 52, 60, 360, 365];

/** C13: numbers this large (in absolute value) are sentinels, "infinity" in a search or an iteration cap (1E+99, 1E+300). */
export const DEFAULT_SENTINEL_ABOVE = 1e90;

export const DEFAULT_LIMITS = { lengthWarn: 7500, lengthError: 8192, nestingWarn: 48, nestingError: 64 } as const;

const SEVERITY_RANK: Record<AuditSeverity, number> = { error: 0, warning: 1, info: 2 };
const KIND_RANK: Record<FindingWhere["kind"], number> = { part: -1, name: 0, cell: 1, cf: 2, dv: 3, table: 4, chart: 5 };
const FIT_TEXT = { exact: "its whole saved extent", part: "part of it", beyond: "and cells outside it" } as const;

type Draft = Omit<Finding, "severity" | "message"> & { severity?: AuditSeverity; message?: string };

function plural(n: number, word: string): string {
  return `${n} ${word}${n === 1 ? "" : "s"}`;
}

function wantText(min: number, max: number): string {
  if (max >= 255) return `at least ${plural(min, "argument")}`;
  if (min === max) return plural(min, "argument");
  return `${min} to ${max} arguments`;
}

/** The place of a formula node: its cell, or for a shared formula its master and group. */
function cellWhere(line: FormulaViewLine): FindingWhere {
  if (line.kind === "shared") {
    const w: FindingWhere = { kind: "cell", sheet: line.sheet, ref: line.master ?? line.cell };
    if (line.extent !== undefined) w.range = line.extent;
    return w;
  }
  const w: FindingWhere = { kind: "cell", sheet: line.sheet, ref: line.cell };
  if (line.extent !== undefined && line.extent !== line.cell) w.range = line.extent;
  return w;
}

export function audit(wb: WorkbookSnapshot, opts: AuditOptions = {}): AuditReport {
  const names = opts.names ?? workbookNameIndex(wb);
  const graph = opts.graph ?? buildGraph(wb, names);
  const limits = { ...DEFAULT_LIMITS, ...opts.limits };
  const allowed = new Set(opts.constants?.allow ?? DEFAULT_CONSTANTS);
  const sentinelAbove = opts.constants?.sentinelAbove ?? DEFAULT_SENTINEL_ABOVE;
  const harness = opts.harness ?? [];

  const defs = auditedNames(wb);
  const defByKey = new Map<string, DefinedName>();
  const classes = new Map<string, Classification>();
  const localNames = new Map<string, string[]>();
  /** lower-case name → scopes defining it ("" for the workbook). */
  const scopes = new Map<string, string[]>();
  for (const d of defs) {
    const scope = scopeOf(d);
    const key = nameKey({ name: d.name, scope });
    if (defByKey.has(key)) continue;
    defByKey.set(key, d);
    classes.set(key, classify(d.definition));
    if (scope !== undefined) {
      const l = localNames.get(scope.toLowerCase()) ?? [];
      l.push(d.name);
      localNames.set(scope.toLowerCase(), l);
    }
    const s = scopes.get(d.name.toLowerCase()) ?? [];
    s.push(scope ?? "");
    scopes.set(d.name.toLowerCase(), s);
  }
  const position = new Map(wb.sheets.map((s) => [s.name.toLowerCase(), s.position]));
  const hasLocal = (sheet: string, id: string) => (scopes.get(id.toLowerCase()) ?? []).some((s) => s !== "" && s.toLowerCase() === sheet.toLowerCase());
  const hasWorkbook = (id: string) => (scopes.get(id.toLowerCase()) ?? []).includes("");
  const sheetsWith = (id: string) => (scopes.get(id.toLowerCase()) ?? []).filter((s) => s !== "").sort((a, b) => (position.get(a.toLowerCase()) ?? 0) - (position.get(b.toLowerCase()) ?? 0));
  const nameWhereOf = (key: string): FindingWhere => {
    const d = defByKey.get(key);
    if (d) return nameWhere(d);
    const bang = key.lastIndexOf("!");
    return bang < 0 ? { kind: "name", name: key, key } : { kind: "name", sheet: key.slice(0, bang), name: key.slice(bang + 1), key };
  };
  const nodeWhere = (n: GraphNode): FindingWhere => (n.kind === "name" ? nameWhereOf(n.name!.key) : cellWhere(n.line!));

  const drafts: Draft[] = [];
  const add = (rule: string, where: FindingWhere, data: Record<string, unknown>): void => {
    const spec = RULES[rule];
    if (!spec) throw new Error(`audit: no rule ${rule}`);
    drafts.push({ check: spec.check, rule, where, data });
  };
  const at = (site: Site, span: { start: number; end: number }): FindingWhere => ({ ...site.where, span: { start: span.start, end: span.end }, text: site.stored.slice(span.start, span.end) });

  /** Lower-case sheet → cell (`C2`) → its formula text (a shared child's: its master's). */
  const cellFormulas = new Map<string, Map<string, string | undefined>>();
  for (const s of wb.sheets) {
    if (s.kind !== "worksheet") continue;
    const m = new Map<string, string | undefined>();
    for (const f of s.formulas) m.set(f.cell, f.text);
    for (const f of s.formulas) if (f.kind === "shared-child" && f.master !== undefined) m.set(f.cell, m.get(f.master));
    cellFormulas.set(s.name.toLowerCase(), m);
  }
  /**
   * C6 for a cell called like a function (`C2(x, y)`, or a name on C2 called): the LAMBDA
   * the cell holds takes so many arguments; a cell holding anything else gives #VALUE! or
   * #CALC!. A formula that may give a LAMBDA (`=IF(…)`, `=Fn`) is not judged.
   */
  const cellCall = (site: Site, span: { start: number; end: number }, fn: string, sheet: string, cell: string, n: number): void => {
    const m = cellFormulas.get(sheet.toLowerCase());
    if (!m) return;
    const callee = cellCallee(m.get(cell));
    const where = `${quoteSheet(sheet)}!${cell}`;
    if (callee.kind === "lambda" && callee.arity) {
      const { required, optional } = callee.arity;
      if (n < required || n > required + optional) {
        add("C6.lambda-arity", at(site, span), { fn, params: (callee.params ?? []).join(", "), want: wantText(required, required + optional), n });
      }
    } else if (callee.kind === "not-lambda") {
      add("C6.not-a-lambda", at(site, span), { fn, cell: where, holds: callee.holds === "value" ? "holds a value or nothing" : "has a formula that gives a value", n });
    }
  };

  // ---- formula-level checks, every formula text once ----------------------------------
  const calledWithText = new Set<string>();
  /** Defined names (by key) called as `Name(…)` but spelled like a built-in: calls that reach the built-in. */
  const collisionCalls = new Map<string, number>();
  /** C5 findings, so C4 does not report the same unknown name again. */
  const c5 = new Set<string>();
  for (const site of collectSites(wb)) {
    const parsed = tryParse(site.stored);
    if (!parsed.formula) {
      const d = parsed.diagnostics[0];
      const span = d ? { start: d.start, end: d.end } : { start: 0, end: site.stored.length };
      const { line, col } = lineCol(site.stored, span.start);
      add("C1.syntax", at(site, span), { error: d?.message ?? "does not parse", line, col });
      continue;
    }
    const body = parsed.formula.body;

    // C2 and C7 (length of the formula as Excel shows it).
    const links = wb.externalLinks ?? [];
    const ctx = site.name && site.home !== undefined ? { homeSheet: site.home, localNames: localNames.get(site.home.toLowerCase()) ?? [], links } : { links };
    let display = site.stored;
    try {
      const r = decompileWithDiagnostics(site.stored, ctx);
      display = r.text;
      for (const d of r.diagnostics) {
        const text = site.stored.slice(d.start, d.end);
        const { prefix, base } = stripPrefix(text);
        const info = lookupFunction(base);
        if (d.code === "bare-prefix" && info) add("C2.bare-prefix", at(site, d), { fn: info.name, prefix: info.prefix });
        else if (d.code === "poisoned") add("C2.poisoned", at(site, d), { text, fn: base.toUpperCase() });
        else if (d.code === "wrong-prefix" && info) add("C2.wrong-prefix", at(site, d), { text, got: prefix, want: info.prefix || "no prefix" });
        else if (d.code === "unknown-function") add("C2.unknown-function", at(site, d), { text, fn: base });
      }
    } catch {
      // parsed above; decompile cannot fail on it
    }
    const length = display.replace(/\r\n/g, "\n").length;
    if (length >= limits.lengthError) add("C7.length", site.where, { length });
    else if (length >= limits.lengthWarn) add("C7.length-near", site.where, { length });
    const depth = nestingDepth(body);
    if (depth > limits.nestingError) add("C7.nesting", site.where, { depth });
    else if (depth > limits.nestingWarn) add("C7.nesting-near", site.where, { depth });

    // C6 (and what C3 and the census need from calls).
    for (const c of callUses(body)) {
      const n = c.args.length;
      if (c.kind === "builtin") {
        if (n < c.info.minArgs || n > c.info.maxArgs) add("C6.builtin-arity", at(site, c.span), { fn: c.info.name, want: wantText(c.info.minArgs, c.info.maxArgs), n });
        const key = names.resolve({ id: c.id, sheet: undefined, span: c.span }, site.home);
        if (key !== undefined && defByKey.has(key)) collisionCalls.set(key, (collisionCalls.get(key) ?? 0) + 1);
        continue;
      }
      if (c.kind === "cell") {
        const sheet = c.sheet ?? site.home;
        const cell = parseCell(c.ref.address);
        if (sheet !== undefined && cell) cellCall(site, c.span, site.stored.slice(c.span.start, c.span.end), sheet, formatCell(cell), n);
        continue;
      }
      const key = names.resolve({ id: c.id, sheet: c.sheet, span: c.span }, site.home);
      const cls = key === undefined ? undefined : classes.get(key);
      if (!cls || key === undefined) continue;
      if (cls.kind === "lambda" && cls.arity) {
        const { required, optional } = cls.arity;
        if (n < required || n > required + optional) {
          add("C6.lambda-arity", at(site, c.span), { fn: c.id, params: (cls.params ?? []).join(", "), want: wantText(required, required + optional), n });
        }
        if (c.args.some((a) => unwrap(a).kind === "string")) calledWithText.add(key);
      } else if (cls.kind === "range" || cls.kind === "spill") {
        // A name on a cell called: the LAMBDA the cell holds (`Fn` = `S!$C$2`, `Fn(x, y)`).
        const d = defByKey.get(key)!;
        const t = definitionTarget(d.definition);
        const sheet = t?.sheet ?? scopeOf(d);
        if (t && !t.spill && t.r1 === t.r2 && t.c1 === t.c2 && sheet !== undefined) {
          cellCall(site, c.span, c.id, sheet, formatCell({ row: t.r1, col: t.c1 }), n);
        } else {
          add("C6.not-a-lambda", at(site, c.span), { fn: c.id, cell: c.id, holds: cls.kind === "spill" ? "is a spill range" : "is a range of several cells", n });
        }
      } else if (cls.kind === "constant" || cls.kind === "table") {
        add("C6.not-a-function", at(site, c.span), { fn: c.id, kind: cls.kind, n });
      }
    }

    if (!site.name) continue;
    const key = site.where.key!;
    // C5: unqualified reads of sheet-scoped names inside a definition.
    for (const u of nameUses(body)) {
      if (u.sheet !== undefined) continue;
      const home = site.home;
      if (home !== undefined && hasLocal(home, u.id)) {
        add("C5.own-sheet", at(site, u.span), { id: u.id, sheet: home, qualified: `${quoteSheet(home)}!${u.id}` });
        c5.add(`${key}\u0000${u.id.toLowerCase()}`);
      } else if (!hasWorkbook(u.id)) {
        const elsewhere = sheetsWith(u.id);
        if (elsewhere.length === 0) continue;
        add("C5.other-sheet", at(site, u.span), {
          id: u.id,
          sheets: elsewhere.map((s) => quoteSheet(s)).join(", "),
          home: home === undefined ? "at workbook scope" : `on ${quoteSheet(home)}`,
          qualified: `${quoteSheet(elsewhere[0]!)}!${u.id}`,
        });
        c5.add(`${key}\u0000${u.id.toLowerCase()}`);
      }
    }
    // C13: numbers in LAMBDA bodies.
    const found = outerLambdas(body).flatMap((l) => numbersIn(l.body, site.stored)).filter((x) => !allowed.has(x.value) && !allowed.has(-x.value) && !(Math.abs(x.value) >= sentinelAbove));
    if (found.length > 0) {
      const values = [...new Set(found.map((x) => x.text))];
      add("C13.constant", at(site, found[0]!.span), { values: values.join(", "), count: found.length });
    }
  }

  // ---- C3: names that lose to a built-in ------------------------------------------------
  for (const [key, d] of defByKey) {
    const info = builtinCollision(d.name);
    if (!info) continue;
    if (classes.get(key)?.kind === "lambda") add("C3.lambda", nameWhere(d), { name: d.name, fn: info.name });
    else if ((collisionCalls.get(key) ?? 0) > 0) add("C3.called", nameWhere(d), { name: d.name, fn: info.name, calls: collisionCalls.get(key) });
  }

  // ---- C4: what the graph could not resolve; names defined as an error -----------------
  const errorDefs = new Set<string>();
  for (const [key, d] of defByKey) {
    const body = tryParse(d.definition).formula?.body;
    const e = body && unwrap(body);
    if (e && (e.kind === "error" || (e.kind === "ref" && e.refKind === "error"))) {
      errorDefs.add(key);
      add("C4.error-definition", nameWhere(d), { value: d.definition.trim() });
    }
  }
  const grouped = new Set<string>();
  for (const n of graph.flagged()) {
    if (n.kind === "input") continue;
    if (n.kind === "name" && (errorDefs.has(n.name!.key) || !n.name!.defined)) continue;
    const where = nodeWhere(n);
    for (const f of n.flags) {
      let rule: string | undefined;
      const data: Record<string, unknown> = { text: f.text, reason: f.reason };
      switch (f.code) {
        case "ref-deleted":
          rule = "C4.ref-deleted";
          break;
        case "unknown-name": {
          if (n.kind === "name" && c5.has(`${n.name!.key}\u0000${f.text.toLowerCase()}`)) break;
          rule = "C4.unknown-name";
          const elsewhere = sheetsWith(stripPrefix(f.text).base);
          if (elsewhere.length > 0) data["elsewhere"] = elsewhere;
          break;
        }
        case "table":
          rule = "C4.table";
          break;
        case "no-anchor":
          rule = "C4.no-anchor";
          break;
        case "no-sheet":
        case "unreadable":
          rule = "C4.broken";
          break;
        case "relative-in-name":
          rule = "C4.relative";
          break;
        case "unqualified-in-name":
          rule = "C4.unqualified-ref";
          break;
        case "external":
          rule = "C4.external";
          break;
        default:
          break; // unparsable (C1), INDIRECT and OFFSET (not findings)
      }
      if (!rule) continue;
      // A shared formula reports once for its group.
      if (n.kind === "formula" && n.line!.kind === "shared") {
        const g = `${rule}\u0000${where.sheet}\u0000${where.ref}`;
        if (grouped.has(g)) continue;
        grouped.add(g);
      }
      add(rule, { ...where, text: f.text }, data);
    }
  }

  // ---- C9: fixed references into a spill; the spill census ------------------------------
  const seenC9 = new Set<string>();
  for (const f of graph.spillRefs) {
    const where = nodeWhere(f.node);
    if (f.node.kind === "formula" && f.node.line!.kind === "shared") {
      const g = `${where.sheet}\u0000${where.ref}\u0000${f.spill.key}`;
      if (seenC9.has(g)) continue;
      seenC9.add(g);
    }
    add("C9.fixed-ref", { ...where, text: f.ref }, { ref: f.ref, spill: f.spill.label, use: f.use, fit: FIT_TEXT[f.fit] });
  }
  const spills = spillCensus(wb, graph, defs);

  // ---- C10: unused names, charts counted as uses ----------------------------------------
  const chartUsed = new Set<number>();
  for (const chart of wb.charts) {
    for (const cf of chart.formulas) {
      const body = tryParse(cf.text).formula?.body;
      if (!body) continue;
      walk(body, (x) => {
        const q = x.kind === "name" ? x.qual : undefined;
        if (x.kind !== "name" || (q?.book !== undefined && q.book !== "[0]") || q?.sheet2 !== undefined) return;
        const key = names.resolve({ id: x.id, sheet: q?.sheet, span: x.span }, chart.sheet?.name);
        const node = key !== undefined ? graph.nameNode(key) : undefined;
        if (node) chartUsed.add(node.id);
      });
    }
  }
  // The harness is read by a person or a solver: a root, like a chart.
  const roots = new Set<number>(chartUsed);
  if (harness.length > 0) {
    for (const key of defByKey.keys()) {
      if (!harness.some((p) => globMatch(p, key))) continue;
      const node = graph.nameNode(key);
      if (node) roots.add(node.id);
    }
  }
  // An end result (the last line of a statement, `Net_income @C11# = EBIT - Taxes`) reads
  // something and is read by a person: a root too. A name with no inputs that nothing
  // reads (a constant, an uncalled LAMBDA, a name on a value cell) is still unused.
  const u = graph.unusedNames();
  for (const n of u.unused) if (endResult(graph, n, classes.get(n.name!.key))) roots.add(n.id);
  const live = new Set<number>(roots);
  for (const id of roots) for (const p of graph.allPrecedents(id)) live.add(p.id);
  const unused = u.unused.filter((n) => !live.has(n.id));
  // A name on cells that nothing reads may still be read by a person (a displayed check
  // like SelfTest): info, not a warning (decided 2026-10-07).
  const onCells = (n: GraphNode): boolean => {
    const k = classes.get(n.name!.key)?.kind;
    return k === "range" || k === "spill";
  };
  for (const n of unused) {
    const rule = n.name!.hidden ? "C10.unused-hidden" : onCells(n) ? "C10.unused-cell" : "C10.unused";
    add(rule, nameWhereOf(n.name!.key), { name: n.label, hidden: n.name!.hidden });
  }
  for (const n of u.onlyByUnused) {
    if (live.has(n.id)) continue;
    const by = graph.dependents(n).filter((d) => d.kind === "name" && d.id !== n.id).map((d) => d.label);
    add("C10.only-by-unused", nameWhereOf(n.name!.key), { name: n.label, by: by.join(", ") });
  }

  // ---- C11: copy drift; C8 census -------------------------------------------------------
  const censusOpts = harness.length > 0 ? { ...opts.census, exclude: [...(opts.census?.exclude ?? []), ...harness] } : (opts.census ?? {});
  const driftCtx = { defs: defByKey, classes, localNames, graph };
  const census = nameCensus({ wb, graph, names, classes, calledWithText, opts: censusOpts, agrees: familyAgreement(driftCtx) });
  for (const f of copyDrift(census, driftCtx)) drafts.push(f);

  // ---- C12: name cycles; circular references among cells ----------------------------------
  for (const c of graph.nameCycles()) {
    if (c.recursive) continue;
    add("C12.name-cycle", nameWhereOf(c.members[0]!.name!.key), { members: c.members.map((m) => m.label).join(" ↔ "), keys: c.members.map((m) => m.name!.key) });
  }
  for (const c of graph.cycles) {
    const formulas = c.members.filter((m) => m.kind === "formula");
    if (formulas.length === 0) continue;
    const through = c.members.filter((m) => m.kind === "name").map((m) => m.label);
    const shown = formulas.slice(0, 6).map((m) => m.label).join(", ") + (formulas.length > 6 ? ", …" : "");
    add("C12.circular", cellWhere(formulas[0]!.line!), {
      count: formulas.length,
      members: shown,
      through: through.length ? `, through ${through.slice(0, 6).join(", ")}${through.length > 6 ? ", …" : ""}` : "",
      cycle: c.id,
    });
  }

  // ---- C14: AFE's own copy of the names (file/afe.ts, audit/afe.ts) ----------------------
  for (const d of afeDrafts(wb, nameWhereOf)) drafts.push(d);

  // ---- C15: names whose label cell no longer gives them (project/labels.ts) -------------
  if (opts.values) {
    for (const l of labelDrift(wb, defs, opts.values)) {
      const matching = l.matching.slice(0, 4).join(", ") + (l.matching.length > 4 ? ", …" : "");
      add("C15.label-drift", nameWhereOf(l.key), {
        cell: l.cell,
        label: l.label,
        converted: l.converted,
        range: l.range,
        name: l.name,
        was: l.renamedFrom !== undefined ? `, probably renamed from ${l.renamedFrom}` : "",
        line: l.along === "row" ? "column" : "row",
        matching,
        fix: labelReplaceShort(l.label, labelReplacement(l.label, l.converted, l.name)),
        ...(l.renamedFrom !== undefined ? { renamedFrom: l.renamedFrom } : {}),
      });
    }
  }

  return finish(drafts, opts, position, census, spills);
}

/** C14: AFE's stores, what xln could read of them, and where their text and the names disagree. */
function afeDrafts(wb: WorkbookSnapshot, nameWhereOf: (key: string) => FindingWhere): Draft[] {
  const out: Draft[] = [];
  const add = (rule: string, where: FindingWhere, data: Record<string, unknown>) => out.push({ check: "C14", rule, where, data });
  const status = new Map(afeStatus(wb).map((s) => [s.store, s]));
  for (const store of wb.foreignModuleStores) {
    const where: FindingWhere = { kind: "part", ref: store.part || (store.sheet ?? "") };
    if (store.sheet !== undefined) where.sheet = store.sheet;
    // A part finding prints the part already: say what it is.
    const what = store.kind === "custom-xml" ? "AFE's module store" : afeStoreLabel(store);
    if (store.kind === "locale-sheet") {
      add("C14.afe-locale-sheet", where, { what, sheet: store.sheet });
      continue;
    }
    const s = status.get(store);
    if (!s) {
      add(store.kind === "code-sheet" ? "C14.afe-code-sheet" : "C14.afe-unreadable", where, { what, sheet: store.sheet, reason: store.unreadable ?? "" });
      continue;
    }
    const modules = s.modules.map((m) => `${m.name} (${m.names})`).join(", ");
    const names = s.entries.length;
    const via = s.viaExportedNames ? ` (${s.viaExportedNames} found in the workbook through AFE's list of exported names, not as Module.name)` : "";
    add("C14.afe-store", where, { what, modules: modules || "none", names, plural: names === 1 ? "" : "s", via, modulesList: s.modules.map((m) => m.name), itemId: store.itemId, linked: store.linked });
    if (s.notCompared) add("C14.afe-not-compared", where, { what, reason: s.notCompared });
    for (const e of s.entries) {
      if (e.state !== "differs") continue;
      add("C14.afe-drift", nameWhereOf(e.name), { module: e.module, afe: oneLine(e.afe), workbook: oneLine(e.workbook ?? "") });
    }
    const absent = new Map<string, string[]>();
    for (const e of s.entries) if (e.state === "absent") absent.set(e.module, [...(absent.get(e.module) ?? []), e.name]);
    for (const [module, list] of absent) {
      add("C14.afe-absent", where, { module, count: list.length, plural: list.length === 1 ? "" : "s", names: list.slice(0, 8).join(", ") + (list.length > 8 ? ", …" : ""), list });
    }
  }
  return out;
}

/** A definition on one line, shortened for a message. */
function oneLine(s: string): string {
  const flat = s.split(/\s+/).join(" ").trim();
  return flat.length > 120 ? flat.slice(0, 119) + "…" : flat;
}

function finish(drafts: Draft[], opts: AuditOptions, position: Map<string, number>, census: AuditReport["census"], spills: SpillCensus): AuditReport {
  const only = opts.only ? new Set<CheckId>(opts.only) : undefined;
  const min = SEVERITY_RANK[opts.minSeverity ?? "info"];
  const findings: Finding[] = [];
  for (const d of drafts) {
    const spec = RULES[d.rule]!;
    const override = opts.rules?.[d.rule] ?? opts.rules?.[d.check];
    if (override === "off") continue;
    const severity = override ?? spec.severity;
    if (only && !only.has(d.check)) continue;
    if (SEVERITY_RANK[severity] > min) continue;
    const data = d.data ?? {};
    const f: Finding = { check: d.check, rule: d.rule, severity, where: d.where, message: fill(spec.message, data) };
    if (spec.hint) f.hint = fill(spec.hint, data);
    if (d.data && Object.keys(d.data).length > 0) f.data = d.data;
    findings.push(f);
  }
  const key = (f: Finding): (string | number)[] => {
    const w = f.where;
    const pos = w.sheet === undefined ? -1 : (position.get(w.sheet.toLowerCase()) ?? 999);
    const cell = w.kind === "cell" && w.ref ? parseCell(w.ref) : undefined;
    return [
      CHECK_IDS.indexOf(f.check),
      KIND_RANK[w.kind],
      pos,
      cell?.row ?? 0,
      cell?.col ?? 0,
      (w.name ?? w.ref ?? "").toLowerCase(),
      w.name ?? w.ref ?? "",
      f.rule,
      w.span?.start ?? -1,
      f.message,
    ];
  };
  const keyed = findings.map((f) => ({ f, k: key(f) }));
  keyed.sort((a, b) => {
    for (let i = 0; i < a.k.length; i++) {
      const x = a.k[i]!;
      const y = b.k[i]!;
      if (x === y) continue;
      return x < y ? -1 : 1;
    }
    return 0;
  });
  const zero = (): AuditCounts => ({ error: 0, warning: 0, info: 0 });
  const counts = zero();
  const byCheck = Object.fromEntries(CHECK_IDS.map((c) => [c, zero()])) as Record<CheckId, AuditCounts>;
  for (const { f } of keyed) {
    counts[f.severity]++;
    byCheck[f.check][f.severity]++;
  }
  return {
    format: "xln-audit/1",
    workbook: opts.workbook,
    counts,
    byCheck,
    findings: keyed.map((x) => x.f),
    census,
    spills,
    checks: CHECK_IDS.filter((c) => !only || only.has(c)),
  };
}

/** C9: every dynamic array that spilled when saved, and the names defined over it. */
function spillCensus(wb: WorkbookSnapshot, graph: DependencyGraph, defs: readonly DefinedName[]): SpillCensus {
  const over = new Map<string, SpillEntry["names"]>();
  for (const d of defs) {
    const t = definitionTarget(d.definition);
    if (!t) continue;
    const scope = scopeOf(d);
    const sheet = t.sheet ?? scope;
    if (sheet === undefined) continue;
    const node = graph.formulaAt(sheet, formatCell({ row: t.r1, col: t.c1 }));
    if (!node || node.line!.kind !== "dynamic-array") continue;
    const s = node.rect!;
    let how: SpillEntry["names"][number]["how"];
    if (t.spill) {
      if (s.r1 !== t.r1 || s.c1 !== t.c1) continue;
      how = "spill";
    } else if (t.r1 === s.r1 && t.c1 === s.c1 && t.r2 === s.r2 && t.c2 === s.c2) how = "extent";
    else if (t.r1 === s.r1 && t.c1 === s.c1 && t.r2 === t.r1 && t.c2 === t.c1) how = "anchor";
    else if (t.r2 <= s.r2 && t.c2 <= s.c2) how = "part";
    else continue;
    const list = over.get(node.key) ?? [];
    list.push({ key: nameKey({ name: d.name, scope }), how });
    over.set(node.key, list);
  }
  const spills: SpillEntry[] = [];
  let singleCell = 0;
  const bySheet = new Map<string, SpillCensus["bySheet"][number]>();
  for (const n of graph.nodes) {
    if (n.kind !== "formula" || n.line!.kind !== "dynamic-array") continue;
    const l = n.line!;
    const rows = l.rows ?? 1;
    const cols = l.cols ?? 1;
    if (rows * cols <= 1) {
      singleCell++;
      continue;
    }
    const names = (over.get(n.key) ?? []).sort((a, b) => cmp(a.key.toLowerCase(), b.key.toLowerCase()));
    spills.push({ sheet: l.sheet, anchor: l.cell, extent: l.extent ?? l.cell, rows, cols, formula: l.formula, names });
    const s = bySheet.get(l.sheet) ?? { sheet: l.sheet, spills: 0, spillNamed: 0, fixedNamed: 0, unnamed: 0 };
    s.spills++;
    if (names.some((x) => x.how === "spill")) s.spillNamed++;
    else if (names.length > 0) s.fixedNamed++;
    else s.unnamed++;
    bySheet.set(l.sheet, s);
  }
  const pos = new Map(wb.sheets.map((s) => [s.name, s.position]));
  const ord = (sheet: string) => pos.get(sheet) ?? 0;
  spills.sort((a, b) => ord(a.sheet) - ord(b.sheet) || cellOrder(a.anchor, b.anchor));
  return { spills, singleCell, bySheet: [...bySheet.values()].sort((a, b) => ord(a.sheet) - ord(b.sheet)) };
}

function cellOrder(a: string, b: string): number {
  const x = parseCell(a);
  const y = parseCell(b);
  return (x?.row ?? 0) - (y?.row ?? 0) || (x?.col ?? 0) - (y?.col ?? 0);
}

/**
 * C10: whether an unused name is an end result, one that reads something. Its definition
 * computes from cells, names or Table columns, or it names cells holding formulas. A
 * reference to value cells only (an input range) and a LAMBDA (nothing calls it) are not.
 */
function endResult(graph: DependencyGraph, n: GraphNode, cls: Classification | undefined): boolean {
  if (n.name!.lambda || cls?.kind === "lambda") return false;
  const reads = graph.precedents(n).filter((p) => p.id !== n.id);
  if (reads.some((p) => p.kind === "formula" || p.kind === "name")) return true;
  return reads.length > 0 && (cls?.kind === "formula" || cls?.kind === "table");
}

/** The rule table's title of a check, for reports. */
export function checkTitle(c: CheckId): string {
  return CHECK_TITLES[c];
}

