// Shared round-trip check for the fixture and corpus tests.
import {
  compile,
  decompileWithDiagnostics,
  equalModuloWhitespace,
  parse,
  prettyPrint,
  walk,
  stripPrefix,
  tokenize,
  type FormulaContext,
} from "../../src/index.js";
import type { WorkbookStrings } from "./xlsx-strings.js";

export interface RoundTripReport {
  definitions: number;
  formulas: number;
  /** Stored text with a modern function lacking its prefix (F6): compile repairs it, so it differs. */
  repaired: string[];
  failures: string[];
  /** Equal only when optional sheet quotes are ignored (`Assumpt!X` vs `'Assumpt'!X`). */
  quotingOnly: string[];
  /** Every function name called anywhere (stored base names, upper case), minus defined names and variables. */
  calledFunctions: Set<string>;
}

/** Strict comparison: token texts, layout whitespace dropped. */
function sameTokenText(a: string, b: string): boolean {
  const key = (s: string) =>
    tokenize(s)
      .filter((t) => t.kind !== "ws")
      .map((t) => (t.kind === "isect" ? " " : t.text))
      .join("\u0000");
  return key(a) === key(b);
}

/** Hidden helper names Excel lists live but does not save; ignore any that appear (CONTRIBUTING.md). */
function isHelperName(name: string): boolean {
  return name.startsWith("_xlfn.") || name.startsWith("_xlpm.");
}

export function roundTrip(wb: WorkbookStrings, label: string): RoundTripReport {
  const report: RoundTripReport = { definitions: 0, formulas: 0, repaired: [], failures: [], quotingOnly: [], calledFunctions: new Set() };
  const allNames = wb.names.map((n) => n.name);
  const lowerNames = new Set(allNames.map((n) => n.toLowerCase()));

  const check = (where: string, stored: string, ctx: FormulaContext): void => {
    try {
      const { text: display, diagnostics } = decompileWithDiagnostics(stored, ctx);
      const back = compile(display, ctx);
      const again = decompileWithDiagnostics(display, ctx).text;
      const stable = compile(stored, ctx);
      const repairing = diagnostics.some((d) => d.code === "bare-prefix");
      if (repairing) {
        report.repaired.push(`${where}: ${stored}  →  ${back}`);
      } else if (!equalModuloWhitespace(back, stored)) {
        report.failures.push(`${where}: round trip\n  stored:   ${stored}\n  display:  ${display}\n  compiled: ${back}`);
        return;
      } else if (!equalModuloWhitespace(stable, stored)) {
        report.failures.push(`${where}: compile is not idempotent\n  stored:   ${stored}\n  compiled: ${stable}`);
        return;
      }
      if (!repairing && !sameTokenText(back, stored)) {
        report.quotingOnly.push(`${where}: ${stored}  →  ${back}`);
      }
      // The pretty-printed display text must mean the same and compile to the same.
      const pretty = prettyPrint(display, { width: 60 });
      if (!equalModuloWhitespace(pretty, display) || !equalModuloWhitespace(compile(pretty, ctx), back)) {
        report.failures.push(`${where}: pretty-print changed the formula\n  display: ${display}\n  pretty:  ${pretty}`);
      }
      if (again !== display) {
        report.failures.push(`${where}: decompile is not idempotent\n  display: ${display}\n  again:   ${again}`);
      }
      walk(parse(stored).body, (n) => {
        if (n.kind !== "call" || n.fn.qual) return;
        const { prefix, base } = stripPrefix(n.fn.text);
        if (prefix === "_xlpm." || prefix === "_xludf.") return;
        if (prefix === "" && lowerNames.has(base.toLowerCase())) return;
        report.calledFunctions.add(base.toUpperCase());
      });
    } catch (e) {
      report.failures.push(`${where}: ${(e as Error).message}\n  stored: ${stored}`);
    }
  };

  for (const n of wb.names) {
    if (isHelperName(n.name)) continue;
    report.definitions++;
    const ctx: FormulaContext = { names: allNames };
    if (n.localSheetId !== undefined) {
      ctx.homeSheet = wb.sheets[n.localSheetId];
      ctx.localNames = wb.names.filter((m) => m.localSheetId === n.localSheetId).map((m) => m.name);
    }
    check(`${label} name ${n.name}${n.localSheetId !== undefined ? ` @${ctx.homeSheet}` : ""}`, n.text, ctx);
  }
  for (const f of wb.formulas) {
    report.formulas++;
    check(`${label} ${f.part} ${f.cell}`, f.text, { names: allNames });
  }
  return report;
}
