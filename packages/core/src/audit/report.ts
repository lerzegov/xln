// The audit report as text, for the CLI and the editor's read-only report document:
// a summary per check, the findings grouped by check, then the name census (C8) and the
// spill census (C9). It also returns where each place (a name, a cell) is written in the
// text, so the editor can link it.

import { sheetLabel } from "../view/render.js";
import { CHECK_TITLES } from "./rules.js";
import { CHECK_IDS, type AuditCounts, type AuditReport, type Finding, type FindingWhere } from "./types.js";

export interface AuditLink {
  start: number;
  end: number;
  where: FindingWhere;
}

export interface RenderedAudit {
  text: string;
  /** Places (names, cells, formats, ...) in `text`, in order. */
  links: AuditLink[];
}

export interface RenderAuditOptions {
  /** Findings shown per check before "… n more" (default: all). */
  maxPerCheck?: number;
  /** Line width the census explanation is wrapped to (default 100). */
  width?: number;
}

function plural(n: number, word: string): string {
  return `${n} ${word}${n === 1 ? "" : "s"}`;
}

function countsText(c: AuditCounts): string {
  const parts = [c.error ? plural(c.error, "error") : "", c.warning ? plural(c.warning, "warning") : "", c.info ? `${c.info} info` : ""].filter(Boolean);
  return parts.length ? parts.join(", ") : "none";
}

/** How a place reads in a report: `name BS!Sales_base`, `cell IS!C6`, `validation S1!G1`. */
export function placeText(w: FindingWhere): { prefix: string; label: string; suffix: string } {
  const on = (s: string | undefined) => (s === undefined ? "" : `${sheetLabel(s)}!`);
  switch (w.kind) {
    case "name":
      return { prefix: "name ", label: `${on(w.sheet)}${w.name ?? w.key ?? ""}`, suffix: "" };
    case "cell":
      return { prefix: "cell ", label: `${on(w.sheet)}${w.ref ?? ""}`, suffix: w.range && w.range !== w.ref ? ` (${w.range})` : "" };
    case "cf":
      return { prefix: "conditional format ", label: `${on(w.sheet)}${w.ref ?? ""}`, suffix: "" };
    case "dv":
      return { prefix: "data validation ", label: `${on(w.sheet)}${w.ref ?? ""}`, suffix: "" };
    case "table":
      return { prefix: "Table column ", label: w.name ?? "", suffix: "" };
    case "chart":
      return { prefix: "chart ", label: w.ref ?? "", suffix: w.sheet ? ` on ${w.sheet}` : "" };
    case "part":
      return { prefix: "part ", label: w.ref ?? "", suffix: "" };
  }
}

function wrap(text: string, width: number, indent: string): string[] {
  const out: string[] = [];
  let line = "";
  for (const word of text.split(" ")) {
    if (line && indent.length + line.length + 1 + word.length > width) {
      out.push(indent + line);
      line = word;
    } else line = line ? `${line} ${word}` : word;
  }
  if (line) out.push(indent + line);
  return out;
}

function list(items: readonly string[], max: number): string {
  return items.length <= max ? items.join(", ") : `${items.slice(0, max).join(", ")}, … (${items.length - max} more)`;
}

export function renderAuditReport(r: AuditReport, opts: RenderAuditOptions = {}): RenderedAudit {
  const max = opts.maxPerCheck ?? Infinity;
  const width = opts.width ?? 100;
  let text = "";
  const links: AuditLink[] = [];
  const line = (s = "") => {
    text += s + "\n";
  };

  line(`xln check${r.workbook ? ` ${r.workbook}` : ""}: ${countsText(r.counts)}`);
  for (const c of CHECK_IDS) {
    if (!r.checks.includes(c)) continue;
    const label = `${c} ${CHECK_TITLES[c]}`;
    const note = c === "C8" ? "census below" : c === "C9" ? `${countsText(r.byCheck[c])}; spill census below` : countsText(r.byCheck[c]);
    line(`  ${label.padEnd(36)} ${note}`);
  }

  for (const c of CHECK_IDS) {
    const fs = r.findings.filter((f) => f.check === c);
    if (fs.length === 0) continue;
    line();
    line(`== ${c} ${CHECK_TITLES[c]}: ${countsText(r.byCheck[c])}`);
    for (const f of fs.slice(0, max)) finding(f);
    if (fs.length > max) line(`  … ${fs.length - max} more (--json lists all)`);
  }

  function finding(f: Finding): void {
    const p = placeText(f.where);
    const head = `  ${f.severity.padEnd(8)}${p.prefix}`;
    const start = text.length + head.length;
    links.push({ start, end: start + p.label.length, where: f.where });
    line(`${head}${p.label}${p.suffix}: ${f.message}`);
    if (f.hint) line(`          → ${f.hint}`);
  }

  // ---- C8 ------------------------------------------------------------------------------
  const n = r.census;
  line();
  line(`== C8 ${CHECK_TITLES.C8}`);
  const per = n.byScope.sheets.map((s) => `${s.sheet} ${s.names}`).join(", ");
  line(`  ${plural(n.total, "name")}: ${n.byScope.workbook} workbook-scoped, ${n.total - n.byScope.workbook} sheet-scoped${per ? ` (${per})` : ""}`);
  const kinds = Object.entries(n.byKind).filter(([, v]) => v > 0).map(([k, v]) => `${k} ${v}`).join(" · ");
  line(`  by kind: ${kinds || "none"}${n.hidden ? ` · ${n.hidden} hidden` : ""}${n.builtIns ? ` · ${n.builtIns} built-in (_xlnm.) not counted` : ""}`);
  if (n.coordinates.length) {
    line(`  coordinate tags: ${list(n.coordinates.map((c) => `${c.position === "suffix" ? "_" + c.tag : c.tag + "_"} (${c.families} famil${c.families === 1 ? "y" : "ies"})`), 12)}`);
    line(`  families of names differing only by a tag: ${n.families.length}`);
  } else line("  coordinate tags: none (no families of names differing only by a suffix or prefix)");
  const t = n.tiers;
  line(`  names standing in for dimensions: ${t.total} of ${t.of} (${t.percent}%): T1 ${t.T1} · T2 ${t.T2} · T3 ${t.T3} · T4 ${t.T4}`);
  line(`    T2 as distinct spellings: ${t.T2spellings}; not counted: ${plural(t.library, "library LAMBDA")}, ${t.excluded} excluded`);
  if (n.tierNames.T1.length) line(`    T1 coordinate addressing: ${list(n.tierNames.T1, 20)}`);
  if (n.tierNames.T4.length) line(`    T4 the axis: ${list(n.tierNames.T4, 20)}`);
  if (n.sheetDuplicates.length) line(`  short names on several sheets: ${n.sheetDuplicates.length} (${list(n.sheetDuplicates.map((d) => `${d.name} on ${d.sheets.join(", ")}`), 3)})`);
  for (const l of wrap(n.explanation, width, "  | ")) line(l);

  // ---- C9 census -------------------------------------------------------------------------
  const s = r.spills;
  line();
  line(`== C9 spill census: ${plural(s.spills.length, "dynamic array")} spilled when saved, ${s.singleCell} saved as one cell${s.uncalculated ? `, ${s.uncalculated} not calculated since the build (open and save in Excel to count them)` : ""}`);
  for (const b of s.bySheet) line(`  ${b.sheet}: ${plural(b.spills, "spill")}: ${b.spillNamed} named as x#, ${b.fixedNamed} named only by a fixed range, ${b.unnamed} unnamed`);
  const notable = s.spills.filter((x) => !x.names.some((m) => m.how === "spill"));
  if (notable.length) {
    line("  spills without an x# name:");
    for (const x of notable.slice(0, max === Infinity ? notable.length : max)) {
      const head = "    ";
      const label = `${sheetLabel(x.sheet)}!${x.anchor}#`;
      const start = text.length + head.length;
      links.push({ start, end: start + label.length, where: { kind: "cell", sheet: x.sheet, ref: x.anchor, range: x.extent } });
      const flat = x.formula.replace(/\s+/g, " ");
      const names = x.names.length ? `   names: ${x.names.map((m) => `${m.key} (${m.how})`).join(", ")}` : "";
      line(`${head}${label} (${x.rows}×${x.cols}) = ${flat.length > 60 ? flat.slice(0, 59) + "…" : flat}${names}`);
    }
  }
  return { text, links };
}
