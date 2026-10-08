// The audit's rules as data (brief §10: contributors tune rules, not code): per rule its
// check, default severity, a one-line title, and the message and hint as templates whose
// `{field}`s are filled from the finding's data. A check's code decides *where* a rule
// fires; this table decides what it says and how loud it is. `AuditOptions.rules`
// overrides the severity per rule or per check.

import type { AuditSeverity, CheckId } from "./types.js";

export interface RuleSpec {
  check: CheckId;
  severity: AuditSeverity;
  title: string;
  message: string;
  hint?: string;
}

export const CHECK_TITLES: Record<CheckId, string> = {
  C1: "Syntax",
  C2: "Prefix health",
  C3: "Built-in collision",
  C4: "Unresolved references",
  C5: "Unqualified sheet-scoped reads",
  C6: "Arity",
  C7: "Excel limits",
  C8: "Name census",
  C9: "Fixed references into a spill",
  C10: "Unused names",
  C11: "Copy drift",
  C12: "Cycles",
  C13: "Constants in LAMBDA bodies",
  C14: "AFE's copy of the names",
  C15: "Labels that no longer match their names",
};

export const RULES: Record<string, RuleSpec> = {
  "C1.syntax": {
    check: "C1",
    severity: "error",
    title: "the formula does not parse",
    message: "does not parse: {error} (line {line}, column {col})",
  },
  "C2.bare-prefix": {
    check: "C2",
    severity: "error",
    title: "a modern function stored without its prefix",
    message: "{fn} is stored without its {prefix} prefix: Excel shows #NAME? and re-saves it as _xludf.{fn} for good (probe F6)",
    hint: "store it as {prefix}{fn} (xln build does), or retype the formula in Excel",
  },
  "C2.poisoned": {
    check: "C2",
    severity: "error",
    title: "a call Excel has already poisoned as _xludf.",
    message: "{text}: Excel did not recognise {fn} when it loaded the formula and stored it as an unknown user function; it stays #NAME? until rewritten (probe F6)",
    hint: "rewrite it as {fn}(…) with its proper prefix",
  },
  "C2.wrong-prefix": {
    check: "C2",
    severity: "warning",
    title: "a function stored with an unexpected prefix",
    message: "{text} is stored with {got}, but the catalogue says {want}",
  },
  "C2.unknown-function": {
    check: "C2",
    severity: "info",
    title: "a prefixed function the catalogue does not know",
    message: "{text}: {fn} is not in xln's function catalogue (kept as stored)",
    hint: "if Excel knows it, add it to the catalogue (packages/core/src/lang/catalogue-data.ts)",
  },
  "C3.lambda": {
    check: "C3",
    severity: "error",
    title: "a LAMBDA named like a built-in function",
    message: "the LAMBDA {name} has the name of the built-in {fn}: every call {name}(…) calls {fn} instead, silently (probe T12)",
    hint: "rename it (e.g. {name}_ or a module prefix: FN.{name})",
  },
  "C3.called": {
    check: "C3",
    severity: "warning",
    title: "a name called like the built-in it is named after",
    message: "{name} has the name of the built-in {fn}, and {calls} formula(s) call {name}(…): they call {fn}, not the name (probe T12)",
    hint: "rename the name",
  },
  "C4.ref-deleted": {
    check: "C4",
    severity: "error",
    title: "#REF!: a reference to deleted cells",
    message: "{text}: the reference points at deleted cells (#REF!)",
  },
  "C4.unknown-name": {
    check: "C4",
    severity: "error",
    title: "a name nothing defines",
    message: "{text} is not a defined name, a Table or a function: #NAME?",
  },
  "C4.table": {
    check: "C4",
    severity: "error",
    title: "an unknown Table or column",
    message: "{text}: {reason}",
  },
  "C4.no-anchor": {
    check: "C4",
    severity: "error",
    title: "a spill reference on a cell that does not spill",
    message: "{text}: no dynamic array is anchored at that cell (Excel gives #REF!)",
  },
  "C4.broken": {
    check: "C4",
    severity: "error",
    title: "a reference that cannot be resolved",
    message: "{text}: {reason}",
  },
  "C4.error-definition": {
    check: "C4",
    severity: "error",
    title: "a name defined as an error",
    message: "the definition is the error {value}",
  },
  "C4.relative": {
    check: "C4",
    severity: "warning",
    title: "a relative reference in a name",
    message: "{text} has no $: stored relative to A1, it points somewhere else from each cell that reads the name (not followed)",
    hint: "make it absolute ($A$1) unless the name is meant to move with the cell",
  },
  "C4.unqualified-ref": {
    check: "C4",
    severity: "warning",
    title: "a reference without a sheet in a workbook-scoped name",
    message: "{text} has no sheet: it reads whichever sheet is active",
    hint: "qualify it ('Sheet'!{text})",
  },
  "C4.external": {
    check: "C4",
    severity: "info",
    title: "a reference to another workbook",
    message: "{text} reads another workbook (not followed)",
  },
  "C5.own-sheet": {
    check: "C5",
    severity: "error",
    title: "the definition's own sheet's name read unqualified",
    message: "{id} is {sheet}'s own name but is written without its sheet: Excel stores it as {qualified}; unqualified it is #NAME? inside functions, and as soon as another sheet defines {id} (findings §11.3)",
    hint: "write {qualified}",
  },
  "C5.other-sheet": {
    check: "C5",
    severity: "error",
    title: "another sheet's local name read unqualified",
    message: "{id} exists only on {sheets}, not {home}: unqualified it is #NAME?",
    hint: "write {qualified}",
  },
  "C6.lambda-arity": {
    check: "C6",
    severity: "error",
    title: "a LAMBDA called with the wrong number of arguments",
    message: "{fn}({params}) takes {want}; it is given {n}",
  },
  "C6.not-a-function": {
    check: "C6",
    severity: "error",
    title: "a name called that is not a LAMBDA",
    message: "{fn} is a {kind}, not a LAMBDA, but is called with {n} argument(s)",
  },
  "C6.not-a-lambda": {
    check: "C6",
    severity: "warning",
    title: "a cell called like a function that does not hold a LAMBDA",
    message: "{fn} is called with {n} argument(s), but {cell} {holds}, not a LAMBDA: Excel gives #VALUE! or #CALC!",
  },
  "C6.builtin-arity": {
    check: "C6",
    severity: "warning",
    title: "a built-in function called with the wrong number of arguments",
    message: "{fn} takes {want}; it is given {n}",
    hint: "the arities come from xln's catalogue; if Excel disagrees, correct catalogue-data.ts",
  },
  "C7.length": {
    check: "C7",
    severity: "error",
    title: "over Excel's formula length limit",
    message: "{length} characters: Excel refuses formulas of 8,192 characters or more (probe T15: 7,992 accepted, 8,192 rejected)",
  },
  "C7.length-near": {
    check: "C7",
    severity: "warning",
    title: "near Excel's formula length limit",
    message: "{length} characters, close to Excel's limit (probe T15: 7,992 accepted, 8,192 rejected)",
    hint: "split it: move parts into LET variables backed by other names, or into a LAMBDA",
  },
  "C7.nesting": {
    check: "C7",
    severity: "error",
    title: "over Excel's nesting limit",
    message: "functions nested {depth} deep: Excel allows 64 levels",
  },
  "C7.nesting-near": {
    check: "C7",
    severity: "warning",
    title: "near Excel's nesting limit",
    message: "functions nested {depth} deep, close to Excel's 64 levels",
  },
  "C9.fixed-ref": {
    check: "C9",
    severity: "warning",
    title: "a fixed reference into a spilled range",
    message: "{ref} is a fixed reference into the spill {spill} ({fit}): it does not follow the spill when it grows or shrinks",
    hint: "use {use}",
  },
  "C10.unused": {
    check: "C10",
    severity: "warning",
    title: "a name nothing uses",
    message: "nothing reads {name}: no formula, name, conditional format, data validation, Table column or chart",
  },
  "C10.unused-cell": {
    check: "C10",
    severity: "info",
    title: "a name on cells nothing reads",
    message: "nothing reads {name}: it names cells, so a person may read it on the sheet",
  },
  "C10.unused-hidden": {
    check: "C10",
    severity: "info",
    title: "a hidden name nothing uses",
    message: "nothing reads the hidden name {name} (an add-in or an old feature may have left it)",
  },
  "C10.only-by-unused": {
    check: "C10",
    severity: "info",
    title: "a name read only by unused names",
    message: "{name} is read only by unused names ({by})",
  },
  "C11.drift": {
    check: "C11",
    severity: "warning",
    title: "copy drift in a family of names",
    message: "{name}{at} differs from {others} beyond the {position} ({tag} vs {otherTag}): {here} here, {there} there",
    hint: "make it the same as the others, or say in its comment why it differs",
  },
  "C12.name-cycle": {
    check: "C12",
    severity: "error",
    title: "names that read each other",
    message: "name cycle: {members} read each other through their definitions",
  },
  "C12.circular": {
    check: "C12",
    severity: "info",
    title: "a circular reference among cells",
    message: "circular reference of {count} formula(s): {members}{through}",
    hint: "Excel needs iterative calculation for it; a fixed point (FN.FIXPOINT) or a recursive LAMBDA removes the loop",
  },
  "C13.constant": {
    check: "C13",
    severity: "info",
    title: "hard-coded numbers in a LAMBDA body",
    message: "{values} hard-coded in the LAMBDA body",
    hint: "pass it as a parameter or read it from a named input",
  },
  "C14.afe-store": {
    check: "C14",
    severity: "info",
    title: "the workbook carries modules of Microsoft's Advanced Formula Environment (Excel Labs)",
    message: "{what}: AFE keeps its own copy of {names} name{plural} as module text here, modules {modules}{via}. xln reads and writes the names in the Name Manager and leaves this copy exactly as it is",
    hint: "where AFE's text and the names differ is listed under C14; xln build warns when it changes a name AFE's modules define",
  },
  "C14.afe-drift": {
    check: "C14",
    // Info (author, 2026-10-08): the build's afe-modules warning already speaks when xln
    // makes the two copies differ; a drift found later is a state to know, not a fault.
    severity: "info",
    title: "AFE's module text differs from the name in the workbook",
    message: "AFE's module {module} has {afe}; the workbook has {workbook}",
    hint: "AFE shows its own text when it opens and may write it back over this name when its modules are saved from AFE. Before saving modules in AFE, bring the module in line with the Name Manager (edit it in AFE); or leave AFE's modules unsaved",
  },
  "C14.afe-absent": {
    check: "C14",
    severity: "info",
    title: "names AFE's module defines that the workbook does not have",
    message: "AFE's module {module} defines {count} name{plural} the workbook does not have: {names}",
    hint: "saving the module from AFE may create them again",
  },
  "C14.afe-not-compared": {
    check: "C14",
    severity: "info",
    title: "AFE's module text was not compared with the names",
    message: "{what}: not compared with the names: {reason}",
  },
  "C14.afe-unreadable": {
    check: "C14",
    severity: "info",
    title: "a store of AFE that xln cannot read",
    message: "{what}: AFE's copy of the names, which xln cannot read: {reason}. xln leaves it exactly as it is",
  },
  "C14.afe-code-sheet": {
    check: "C14",
    severity: "info",
    title: "AFE 1.0's hidden code sheet",
    message: "{what}: AFE 1.0 kept its module text in this sheet; AFE 1.1 and later convert it to a custom XML part and leave the sheet. xln does not read it and leaves it as it is",
    hint: "once AFE has converted it, the sheet can be deleted in Excel (it is very hidden: unhide it with VBA, or leave it)",
  },
  "C14.afe-locale-sheet": {
    check: "C14",
    severity: "info",
    title: "AFE's locale-detection sheet",
    message: "{what}: a scratch sheet AFE uses to find the argument separator; it holds no names or modules",
    hint: "it can be deleted in Excel (AFE's FAQ)",
  },
  "C15.label-drift": {
    check: "C15",
    severity: "info",
    title: "a name whose label cell no longer gives it",
    message: '{cell} reads "{label}", which Create from Selection makes {converted}; the name on {range} is {name}{was}',
    hint: "xln never writes cell values: {fix}, or rename the name back. The other names in its {line} of labels match theirs ({matching})",
  },
};

/** Fills `{field}` from `data`; a field with no value is left as written. */
export function fill(template: string, data: Record<string, unknown>): string {
  let out = "";
  let k = 0;
  while (k < template.length) {
    const open = template.indexOf("{", k);
    if (open < 0) break;
    const close = template.indexOf("}", open + 1);
    if (close < 0) break;
    const field = template.slice(open + 1, close);
    const v = data[field];
    out += template.slice(k, open) + (v === undefined ? template.slice(open, close + 1) : String(v));
    k = close + 1;
  }
  return out + template.slice(k);
}
