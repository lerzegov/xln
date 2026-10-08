// Tokenizes samples/demo.xln with the same engine VS Code uses (vscode-textmate over
// Oniguruma), so the grammar is checked without opening an editor.
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { join } from "node:path";
import { readWorkbook, renderFormulaView, sheetFormulaView } from "@xln/core";
import { beforeAll, describe, expect, it } from "vitest";
import oniguruma from "vscode-oniguruma";
import textmate from "vscode-textmate";

const require = createRequire(import.meta.url);
const ROOT = join(import.meta.dirname, "..");

interface Token {
  line: number;
  text: string;
  scopes: string[];
}

let tokens: Token[];

const GRAMMARS: Record<string, string> = { "source.xln": "xln.tmLanguage.json", "source.xln-formulas": "xln-formulas.tmLanguage.json" };

async function tokenize(source: string, scope = "source.xln"): Promise<Token[]> {
  const wasm = readFileSync(join(require.resolve("vscode-oniguruma"), "..", "onig.wasm"));
  await oniguruma.loadWASM(wasm.buffer.slice(wasm.byteOffset, wasm.byteOffset + wasm.byteLength) as ArrayBuffer);
  const registry = new textmate.Registry({
    onigLib: Promise.resolve({
      createOnigScanner: (p: string[]) => new oniguruma.OnigScanner(p),
      createOnigString: (s: string) => new oniguruma.OnigString(s),
    }),
    loadGrammar: async (name: string) =>
      GRAMMARS[name] ? textmate.parseRawGrammar(readFileSync(join(ROOT, "syntaxes", GRAMMARS[name]), "utf8"), GRAMMARS[name]) : null,
  });
  const grammar = await registry.loadGrammar(scope);
  if (!grammar) throw new Error("grammar did not load");
  const out: Token[] = [];
  let state = textmate.INITIAL;
  source.split("\n").forEach((text, line) => {
    const r = grammar.tokenizeLine(text, state);
    for (const t of r.tokens) out.push({ line: line + 1, text: text.slice(t.startIndex, t.endIndex), scopes: t.scopes });
    state = r.ruleStack;
  });
  return out;
}

/** Scope lists of every token with this exact text. */
function scopesOf(text: string, among: Token[] = tokens): string[][] {
  return among.filter((t) => t.text === text).map((t) => t.scopes);
}

function expectScope(text: string, scope: string, among: Token[] = tokens): void {
  const all = scopesOf(text, among);
  expect(all.length, `token ${JSON.stringify(text)} not found`).toBeGreaterThan(0);
  expect(
    all.some((s) => s.some((x) => x.startsWith(scope))),
    `${JSON.stringify(text)} has scopes ${JSON.stringify(all)}, expected ${scope}`,
  ).toBe(true);
}

beforeAll(async () => {
  tokens = await tokenize(readFileSync(join(ROOT, "samples", "demo.xln"), "utf8"));
});

describe("xln grammar on samples/demo.xln", () => {
  it("leaves no text uncoloured", () => {
    const bare = tokens.filter((t) => {
      if (t.text.trim() === "") return false;
      const inner = t.scopes[t.scopes.length - 1]!;
      return inner === "source.xln" || inner.startsWith("meta.");
    });
    expect(bare.map((t) => `${t.line}: ${JSON.stringify(t.text)} ${t.scopes.join(" ")}`)).toEqual([]);
  });

  it("module syntax", () => {
    expectScope("/**", "comment.block.documentation");
    expectScope("//", "punctuation.definition.comment");
    expectScope("Timeline", "entity.name.function");
    expectScope("Assumptions.", "entity.name.namespace.module");
    expectScope("Growth", "entity.name.function");
    expectScope("NPVrow", "entity.name.function");
    expectScope("=", "keyword.operator.assignment");
    expectScope(";", "punctuation.terminator.statement");
  });

  it("a type declaration (not xln v1) is marked invalid", async () => {
    const t = await tokenize("Periods : scalar = 5;");
    expectScope("scalar", "invalid.illegal.type-declaration", t);
  });

  it("annotations", () => {
    expectScope("sheet", "storage.type.annotation.sheet");
    expectScope("scope", "storage.type.annotation.scope");
    expectScope("workbook", "storage.type.annotation.workbook");
    expectScope("hidden", "storage.type.annotation");
    expectScope("from", "storage.type.annotation.from");
    expectScope("lib", "entity.name.namespace.library");
    expectScope("3f9a1c", "constant.numeric.hex.library-version");
    expectScope("Model", "entity.name.namespace.sheet");
    expectScope("'Cash Flow'", "entity.name.namespace.sheet");
  });

  it("formulas", () => {
    expectScope("LAMBDA", "storage.type.function");
    expectScope("LET", "storage.type.function");
    expectScope("SEQUENCE", "support.function");
    expectScope("_xlfn.", "storage.modifier.prefix");
    expectScope("_xlfn._xlws.", "storage.modifier.prefix");
    expectScope("x", "variable.parameter");
    expectScope("FILTER", "support.function");
    expectScope('""', "constant.character.escape");
    expectScope("2.5", "constant.numeric");
    expectScope(".75", "constant.numeric");
    expectScope("4e-2", "constant.numeric");
    expectScope("%", "keyword.operator.percent");
    expectScope("TRUE", "constant.language.boolean");
    expectScope("#DIV/0!", "constant.language.error");
    expectScope("#NAME?", "constant.language.error");
    expectScope("#N/A", "constant.language.error");
    expectScope("<>", "keyword.operator.comparison");
    expectScope("&", "keyword.operator.concatenation");
  });

  it("@scope and @workbook mark scope blocks", () => {
    for (const t of ["scope", "workbook"]) {
      expect(scopesOf(t).every((s) => s.includes("meta.annotation.scope-block.xln")), t).toBe(true);
    }
    expect(scopesOf("hidden").every((s) => s.includes("meta.annotation.xln"))).toBe(true);
    // The definition after @workbook is a definition, not part of the directive.
    expectScope("Assumptions.", "entity.name.namespace.module");
  });

  it("cell statements: named, slot, unnamed, blocks, a qualified address", () => {
    for (const a of ["C4", "C5", "C6", "C7", "B40:G40", "$D$9", "D10"]) expectScope(a, "constant.other.reference.cell.address");
    // The # of a named statement's address: the name is on the spill.
    for (const name of ["Revenue", "Spilled"]) {
      const at = tokens.find((u) => u.text === name && tokens.some((v) => v.line === u.line && v.text === "@"))!.line;
      const line = tokens.filter((t) => t.line === at);
      const hash = line.findIndex((t) => t.text === "#");
      expect(line[hash - 1]!.scopes.at(-1)).toBe("constant.other.reference.cell.address.xln");
      expect(line[hash]!.scopes.at(-1)).toBe("keyword.operator.spill.address.xln");
    }
    expect(scopesOf("C5").every((s) => s.includes("meta.cell-statement.xln"))).toBe(true);
    expectScope("Receivables", "entity.name.function");
    expectScope("Slot", "entity.name.function");
    expectScope("Created", "entity.name.function");
    expectScope("'Cash Flow'", "entity.name.namespace.sheet");
    // The formula after an unnamed address is a formula; the slot's `;` ends it.
    expectScope("Years", "variable.other.name");
    const slot = tokens.filter((t) => t.line === tokens.find((u) => u.text === "Slot")!.line);
    expect(slot.at(-1)!.scopes.at(-1)).toBe("punctuation.terminator.statement.xln");
  });

  it("array constants keep the definition open across ';'", () => {
    const row = tokens.filter((t) => t.line === tokens.find((u) => u.text === "Arrays")!.line);
    expect(row.filter((t) => t.text === ";").map((t) => t.scopes.at(-1))).toEqual([
      "punctuation.separator.array.row.xln",
      "punctuation.terminator.statement.xln",
    ]);
  });

  it("references", () => {
    expectScope("Sheet1", "entity.name.namespace.sheet");
    expectScope("Sheet1:Sheet3", "entity.name.namespace.sheet");
    expectScope("[1]Ext", "entity.name.namespace.sheet");
    expectScope("!", "punctuation.separator.sheet");
    expectScope("$B$3", "constant.other.reference.cell");
    expectScope("A1", "constant.other.reference.cell");
    expectScope("$A", "constant.other.reference.column");
    expectScope("$C", "constant.other.reference.column");
    expectScope("1", "constant.numeric");
    expectScope("3", "constant.other.reference.row");
    expectScope(".:", "keyword.operator.range.trim");
    expectScope(":.", "keyword.operator.range.trim");
    expectScope(".:.", "keyword.operator.range.trim");
    expect(scopesOf("B100").map((s) => s.at(-1))).toEqual(Array(3).fill("constant.other.reference.cell.xln"));
    expectScope("#", "keyword.operator.spill");
    expectScope("Revenue", "variable.other.name");
    expectScope("Fin.", "entity.name.namespace.module");
  });

  it("structured references", () => {
    expectScope("Sales", "entity.name.type.table");
    expectScope("Amount", "variable.other.member.column");
    expectScope("#Headers", "support.constant.structured-reference");
    expectScope("#This Row", "support.constant.structured-reference");
    expectScope("#All", "support.constant.structured-reference");
    expectScope("''", "constant.character.escape");
    expectScope("@", "keyword.operator.this-row");
  });

  it("non-ASCII names", () => {
    expectScope("Crescità", "variable.other.name");
    expectScope("Δ", "variable.other.name");
  });
});

describe("formula view grammar on the view of f7_base.xlsx", () => {
  let view: Token[];
  let text: string;
  beforeAll(async () => {
    const wb = readWorkbook(new Uint8Array(readFileSync(join(ROOT, "..", "..", "probes", "results", "f7_base.xlsx"))));
    // A narrow formula column turns the longer formulas into blocks, with open parentheses.
    text = renderFormulaView(sheetFormulaView(wb, "S1"), { sheet: "S1", formulaWidth: 12, width: 40 }).text;
    view = await tokenize(text, "source.xln-formulas");
  });

  it("colours the columns: cell, kind, value; formulas as in .xln files", () => {
    expectScope("// Sheet S1: 22 formulas in order of appearance (row by row, left to right). Read-only view.", "comment.line", view);
    expectScope("E1", "entity.name.tag.cell", view);
    expectScope("#", "keyword.operator.spill", view);
    expectScope("(3×1)", "storage.modifier.kind", view);
    expectScope("shared ← B2", "storage.modifier.kind", view);
    expectScope("→", "punctuation.separator.value", view);
    expectScope("0.1 …", "constant.other.value", view);
    expectScope('"Rate is 0.1"', "constant.other.value", view);
    expectScope("SEQUENCE", "support.function", view);
    expectScope("Rate", "variable.other.name", view);
    expectScope("LET", "storage.type.function", view);
    expectScope("A3", "constant.other.reference.cell", view);
    expectScope("Spl", "entity.name.function.lhs", view);
  });

  it("colours the names on a cell, in the column or on a line of their own", async () => {
    const block = [
      "AccountsReceivable_base, 'My Sheet'!X  C6#   (1×5)  = IN.ASMPT_base(\"AR days\") * 2  → 1 …",
      "                                       C17          = SUM(C6:C11)  → 742",
      "AVeryLongNameThatDoesNotFitTheNameColumnAtAll",
      "                                       E4           = IF(",
      "                                                          A1 = 0,",
      "                                                          1, 2)",
    ].join("\n");
    const t = await tokenize(block, "source.xln-formulas");
    expectScope("AccountsReceivable_base", "entity.name.function.lhs", t);
    expectScope("'My Sheet'", "entity.name.namespace.sheet", t);
    expectScope("X", "entity.name.function.lhs", t);
    expectScope("C6", "entity.name.tag.cell", t);
    expectScope("C17", "entity.name.tag.cell", t);
    expectScope("AVeryLongNameThatDoesNotFitTheNameColumnAtAll", "entity.name.function.lhs", t);
    expectScope("E4", "entity.name.tag.cell", t);
    // A comparison in a continued formula is not an entry.
    expect(t.filter((x) => x.line === 5 && x.text === "A1").map((x) => x.scopes.at(-1))).toEqual(["constant.other.reference.cell.xln"]);
  });

  it("colours calculation order: level, cycle mark, sheet-qualified cells", async () => {
    const block = [
      "2                                  C5#    (1×5)  = Model!Years  → 2022 …",
      "17 ↻1  InterestIncome_base         IS!C16#  (1×5)  = A1  → 0 …",
      "3      'SCF recursive'!C12#  (1×5)  = B2  → 1",
    ].join("\n");
    const t = await tokenize(block, "source.xln-formulas");
    expectScope("2", "constant.numeric.level", t);
    expectScope("17 ↻1", "constant.numeric.level", t);
    expectScope("InterestIncome_base", "entity.name.function.lhs", t);
    expectScope("IS", "entity.name.namespace.sheet", t);
    expectScope("C16", "entity.name.tag.cell", t);
    expectScope("'SCF recursive'", "entity.name.namespace.sheet", t);
    expectScope("C12", "entity.name.tag.cell", t);
    expectScope("C5", "entity.name.tag.cell", t);
  });

  it("finds the value column inside a formula left open on its first line", async () => {
    const block = ["C14#  (2×5)  = LET(                  → 0 …", "                  dep, FN.PREV(Nfa_base),", "                  dep * 2", "              )", "C16   = A1  → 3"].join("\n");
    const t = await tokenize(block, "source.xln-formulas");
    expect(t.filter((x) => x.line === 1 && x.text === "0 …").map((x) => x.scopes.at(-1))).toEqual(["constant.other.value.xln-formulas"]);
    expectScope("FN.PREV", "support.function", t);
    expectScope("Nfa_base", "variable.other.name", t);
    expectScope("C16", "entity.name.tag.cell", t);
    expectScope("3", "constant.other.value", t);
  });
});
