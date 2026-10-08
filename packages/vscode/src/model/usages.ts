// Workbook-side context for a name, as text: where cells, conditional formats, data
// validations and Table columns use it (B3), and what the hover says (B4). The places
// come from the manifest (the last pull); the formula texts and cached values come from
// the workbook itself when it can be read, because the manifest stores cell addresses
// only. Without the workbook the addresses are still listed.

import { decompile, formulaTextAt, parseDocComment, quoteSheet, shiftFormula, type CachedValue, type CellFormula, type Sheet, type WorkbookSnapshot } from "@xln/core";
import { usageCounts, type Manifest, type ManifestName } from "./manifest.js";
import type { Analysis, NameDef, Project } from "./project.js";

export interface UsageLine {
  /** 0-based line in the document. */
  line: number;
  /** Where the address sits on the line, for the reference location. */
  start: number;
  end: number;
  kind: "cell" | "conditionalFormat" | "dataValidation" | "tableColumn";
  /** For a cell, conditional format or data validation: the sheet and the range or sqref. */
  sheet?: string;
  range?: string;
}

export interface UsageDocument {
  text: string;
  items: UsageLine[];
}

const cellIndex = new WeakMap<Sheet, Map<string, CellFormula>>();

function formulaAt(sheet: Sheet, cell: string): CellFormula | undefined {
  let m = cellIndex.get(sheet);
  if (!m) {
    m = new Map(sheet.formulas.map((f) => [f.cell, f]));
    cellIndex.set(sheet, m);
  }
  return m.get(cell.split("$").join("").toUpperCase());
}

function sheetOf(wb: WorkbookSnapshot | undefined, name: string): Sheet | undefined {
  const l = name.toLowerCase();
  return wb?.sheets.find((s) => s.name.toLowerCase() === l);
}

/** Display form of a stored formula; the stored text if it does not decompile. */
export function display(stored: string): string {
  try {
    return decompile(stored);
  } catch {
    return stored;
  }
}

export function formatValue(v: CachedValue | undefined): string | undefined {
  if (!v || v.value === undefined) return undefined;
  if (typeof v.value === "boolean") return v.value ? "TRUE" : "FALSE";
  // Excel shows at most 15 significant digits: 0.7, not 0.7000000000000001.
  if (typeof v.value === "number") return String(Number(v.value.toPrecision(15)));
  if (v.type === "e") return v.value;
  return JSON.stringify(v.value);
}

/** The formula of a cell as Excel shows it, with its cached value; a shared child's text is moved from its master's. */
export function cellFormula(wb: WorkbookSnapshot | undefined, sheetName: string, cell: string): { text: string; value?: string; note?: string } | undefined {
  const sheet = sheetOf(wb, sheetName);
  const f = sheet && formulaAt(sheet, cell);
  if (!sheet || !f) return undefined;
  const stored = formulaTextAt(sheet, f, shiftFormula);
  if (stored === undefined) return undefined;
  const out: { text: string; value?: string; note?: string } = { text: "=" + display(stored).replace(/\r?\n/g, " ") };
  const value = formatValue(f.value);
  if (value !== undefined) out.value = value;
  if (f.kind === "shared-child" && f.master) out.note = `shared formula of ${f.master}`;
  return out;
}

/** Usage lines stay readable in the references view: a long formula is cut. */
const LINE_MAX = 240;

function clip(text: string): string {
  return text.length > LINE_MAX ? text.slice(0, LINE_MAX - 1) + "…" : text;
}

function topLeft(range: string): string {
  return range.split(":")[0]!;
}

/** The read-only document behind `xln-cells:`: one line per place that uses the name. */
export function renderUsages(key: string, mn: ManifestName | undefined, manifest: Manifest | undefined, wb: WorkbookSnapshot | undefined): UsageDocument {
  const lines: string[] = [];
  const items: UsageLine[] = [];
  const workbook = manifest?.workbook || "the workbook";
  lines.push(`// Where ${key} is used in ${workbook}: cells, conditional formats, data validations, Table columns.`);
  lines.push(`// Places from workbook.manifest.json (the last pull); ${wb ? "formulas and cached values read from the workbook" : "the workbook could not be read, so no formulas"}. Read-only.`);
  const u = mn?.usedBy;
  if (!mn) lines.push("", "Not in the manifest: the name was added after the last pull, or the project was not pulled.");
  const counts = usageCounts(u);

  const item = (kind: UsageLine["kind"], address: string, rest: string, sheet?: string, range?: string): void => {
    items.push({ line: lines.length, start: 0, end: address.length, kind, ...(sheet !== undefined ? { sheet, range } : {}) });
    lines.push(address + rest);
  };

  const section = (title: string, n: number): void => {
    lines.push("", `${title} (${n})`);
  };

  if (u?.cells) {
    section("Cells", counts.cells);
    for (const [sheet, ranges] of Object.entries(u.cells)) {
      for (const r of ranges) {
        const f = cellFormula(wb, sheet, topLeft(r));
        let rest = "";
        if (f) {
          rest = `: ${clip(f.text)}`;
          if (f.value !== undefined) rest += `  → ${f.value}`;
          // A rectangle may join cells with different formulas: name the one shown.
          if (r.includes(":")) rest += `  (formula of ${topLeft(r)})`;
          else if (f.note) rest += `  (${f.note})`;
        }
        item("cell", `${qualify(sheet)}!${r}`, rest, sheet, r);
      }
    }
  }
  if (u?.conditionalFormats) {
    section("Conditional formats", counts.conditionalFormats);
    for (const [sheet, sqrefs] of Object.entries(u.conditionalFormats)) {
      const s = sheetOf(wb, sheet);
      for (const sq of sqrefs) {
        const texts = s?.conditionalFormats.filter((c) => c.sqref === sq).flatMap((c) => c.formulas) ?? [];
        item("conditionalFormat", `${qualify(sheet)}!${sq}`, texts.length ? `: ${clip(texts.map((t) => "=" + display(t)).join(" ; "))}` : "", sheet, sq);
      }
    }
  }
  if (u?.dataValidations) {
    section("Data validations", counts.dataValidations);
    for (const [sheet, sqrefs] of Object.entries(u.dataValidations)) {
      const s = sheetOf(wb, sheet);
      for (const sq of sqrefs) {
        const dv = s?.dataValidations.find((d) => d.sqref === sq);
        const texts = [dv?.formula1, dv?.formula2].filter((t): t is string => t !== undefined && t.trim() !== "");
        item("dataValidation", `${qualify(sheet)}!${sq}`, texts.length ? `: ${clip(texts.map((t) => "=" + display(t)).join(" ; "))}` : "", sheet, sq);
      }
    }
  }
  if (u?.tableColumns) {
    section("Table columns", counts.tableColumns);
    for (const tc of u.tableColumns) {
      const open = tc.indexOf("[");
      const table = manifest?.tables.find((t) => t.name === tc.slice(0, open));
      const col = tc.slice(open + 1, -1);
      const text = table?.calculatedColumns?.[col];
      item("tableColumn", tc, text !== undefined ? `: ${clip("=" + display(text))}` : "");
    }
  }
  if (mn && items.length === 0) lines.push("", "No cell, conditional format, data validation or Table column uses it.");
  return { text: lines.join("\n") + "\n", items };
}

/** `Sheet` or `'My sheet'`, the way a reference in a formula writes it. */
function qualify(sheet: string): string {
  return quoteSheet(sheet);
}

const KIND_LABEL: Record<string, string> = {
  constant: "constant",
  range: "range",
  spill: "spill",
  table: "Table reference",
  formula: "formula",
  lambda: "LAMBDA",
  unparsed: "does not parse",
};

function plural(n: number, one: string, many = one + "s"): string {
  return `${n} ${n === 1 ? one : many}`;
}

/** The hover of a defined name (B4), as Markdown. */
export function hoverMarkdown(project: Project, def: NameDef, wb: WorkbookSnapshot | undefined): string {
  const a: Analysis = project.analysis(def);
  const c = a.classification;
  const mn = project.manifestName(def.key);
  const out: string[] = [];
  const scope = def.scope === undefined ? "workbook scope" : `local to sheet ${def.scope}`;
  const head = c.kind === "lambda" ? `${def.name}(${(c.params ?? []).join(", ")})` : def.name;
  out.push("```xln\n" + head + "\n```");
  const facts = [KIND_LABEL[c.kind] ?? c.kind, scope];
  if (c.arity) facts.push(`arity ${c.arity.required}${c.arity.optional ? `–${c.arity.required + c.arity.optional}` : ""}`);
  if (def.entry.hidden) facts.push("hidden");
  if (c.table) facts.push(`Table ${c.table}`);
  out.push(facts.join(" · "));
  if (def.entry.doc) {
    // `@param x what x is` lines (M3c) as a parameter list under the summary.
    const doc = parseDocComment(def.entry.doc);
    if (doc.summary) out.push(doc.summary);
    if (doc.params.length) out.push(doc.params.map((p) => `*@param* \`${p.name}\` — ${p.text}`).join("  \n"));
  }

  if (c.kind === "spill" || mn?.spill) {
    const sheet = mn?.spill?.sheet ?? c.anchor?.sheet ?? def.entry.cellSheet ?? def.scope;
    const anchor = mn?.spill?.anchor ?? c.anchor?.cell;
    if (sheet && anchor) {
      const f = cellFormula(wb, sheet, anchor);
      const lines = [`**Spill** anchored at \`${qualify(sheet)}!${anchor}\``];
      if (f) lines.push(`anchor formula: \`${f.text}\``);
      const extent = mn?.spill?.extent;
      if (extent) lines.push(`last saved extent: \`${extent}\``);
      if (f?.value !== undefined) lines.push(`first value (cached): \`${f.value}\``);
      out.push(lines.join("  \n"));
    }
  }

  const usedBy = project.usedBy(def.key);
  const uses = project.uses(def);
  const rel: string[] = [];
  if (uses.length) rel.push(`uses ${uses.map((x) => `\`${x.def.key}\``).join(", ")}`);
  if (usedBy.length) rel.push(`used by ${plural(usedBy.length, "name")}`);
  if (mn) {
    const n = usageCounts(mn.usedBy);
    const parts = [
      n.cells ? plural(n.cells, "cell") : "",
      n.conditionalFormats ? plural(n.conditionalFormats, "conditional format") : "",
      n.dataValidations ? plural(n.dataValidations, "data validation") : "",
      n.tableColumns ? plural(n.tableColumns, "Table column") : "",
    ].filter((p) => p !== "");
    rel.push(parts.length ? `used in ${parts.join(", ")}` : "used in no cell");
  } else rel.push("not in the manifest (added after the last pull?)");
  out.push(rel.join(" · "));
  return out.join("\n\n");
}
