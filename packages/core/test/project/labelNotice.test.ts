// The label notice (2026-10-07): after a rename, the text cells that still read the old
// name and the Find & Replace that fixes them in Excel. Hand-made snapshots, values beside.
import { describe, expect, it } from "vitest";
import { labelNotice, labelNoticeLines, labelReplacement, type CellValue } from "../../src/index.js";
import { book, type Book } from "../audit/book.js";

type Values = Record<string, Record<string, CellValue>>;
const valueMap = (v: Values) => new Map(Object.entries(v).map(([s, cells]) => [s, new Map(Object.entries(cells))]));

// The author's layout: A the readable text, B the name's text (Create from Selection), C:E the cells.
const IS: Book = {
  sheets: { IS: { C4: "C2-C3", D4: "D2-D3", E4: "E2-E3", B9: "B4" } },
  names: [
    { name: "Revenue", def: "IS!$C$2:$E$2" },
    { name: "COGS", def: "IS!$C$3:$E$3" },
    { name: "Gross_income", def: "IS!$C$4:$E$4" },
    { name: "Fn", def: "LAMBDA(x,x*2)" },
  ],
};
const VALUES: Values = {
  IS: {
    A2: "Revenue", B2: "Revenue", A3: "Cost of goods", B3: "COGS",
    A4: "Gross income", B4: "Gross_income", C4: 40, D4: 44, E4: 48,
    C1: "Gross income (EUR)",
    // A formula's result (B9 = B4) reads the same text: not a label to replace.
    B9: "Gross_income",
    // Same text, away from the name's rows and columns.
    H20: "Gross income",
  },
  Other: { A4: "Gross_income" },
};

const notice = (from: string, to: string, scope: string | null = null, v: Values = VALUES, b: Book = IS) => labelNotice(book(b), valueMap(v), { from, to, scope });

describe("labelNotice", () => {
  it("finds the texts that give the old name, one pair per text, in the label's style", () => {
    const n = notice("Gross_income", "Gross_ind_income")!;
    expect(n.sheet).toBe("IS");
    expect(n.stale).toEqual([
      { cell: "IS!A4", text: "Gross income" },
      { cell: "IS!B4", text: "Gross_income" },
    ]);
    expect(n.pairs).toEqual([
      { find: "Gross income", replace: "Gross ind income", cells: ["IS!A4"], alsoOnSheet: ["IS!H20"] },
      { find: "Gross_income", replace: "Gross_ind_income", cells: ["IS!B4"], alsoOnSheet: [] },
    ]);
    // Resembles, not a conversion of the whole text: listed only. Formula results and other sheets: nothing.
    expect(n.maybe).toEqual([{ cell: "IS!C1", text: "Gross income (EUR)" }]);
  });

  it("prints the shared notice", () => {
    expect(labelNoticeLines(notice("Gross_income", "Gross_ind_income")!)).toEqual([
      "Gross_income → Gross_ind_income: 2 labels still read the old name: IS!A4, IS!B4",
      "In Excel, on sheet IS: Find & Replace (Ctrl+H on Windows; Ctrl+H or ⌘⇧H on Mac; or Home → Find & Select → Replace)",
      "  1. Find what:     Gross income",
      "     Replace with:  Gross ind income",
      "  2. Find what:     Gross_income",
      "     Replace with:  Gross_ind_income",
      "  Within: Sheet · Look in: Formulas · Match entire cell contents ✓ · then Replace All (for each pair)",
      '  Replace All also changes IS!H20 ("Gross income", away from the name\'s rows and columns): to leave it, use Find Next and Replace instead',
      'Check by eye (not in the replace): IS!C1 "Gross income (EUR)"',
    ]);
  });

  it("one text, one pair, no numbering", () => {
    const lines = labelNoticeLines(notice("COGS", "Cost_of_sales")!);
    expect(lines).toEqual([
      "COGS → Cost_of_sales: 1 label still reads the old name: IS!B3",
      "In Excel, on sheet IS: Find & Replace (Ctrl+H on Windows; Ctrl+H or ⌘⇧H on Mac; or Home → Find & Select → Replace)",
      "  Find what:     COGS",
      "  Replace with:  Cost_of_sales",
      "  Within: Sheet · Look in: Formulas · Match entire cell contents ✓ · then Replace All",
    ]);
  });

  it("a name not on cells (a LAMBDA) and a name not in the workbook give nothing", () => {
    expect(notice("Fn", "Fn2")).toBeUndefined();
    expect(notice("Nope", "Nope2")).toBeUndefined();
  });

  it("a label that is a formula's result is not stale", () => {
    const b: Book = { sheets: { S: { A1: '"Rate"' } }, names: [{ name: "Rate", def: "S!$B$1" }] };
    expect(notice("Rate", "Pace", null, { S: { A1: "Rate", B1: 0.1 } }, b)).toBeUndefined();
  });

  it("a sheet's local name, a quoted sheet, a column under its label", () => {
    const b: Book = { sheets: { "Cash Flow": {} }, names: [{ name: "Capex", def: "'Cash Flow'!$B$2:$B$6", scope: "Cash Flow" }] };
    const n = notice("Capex", "Capital_expenditure", "Cash Flow", { "Cash Flow": { B1: "Capex", B2: 1 } }, b)!;
    expect(n.stale).toEqual([{ cell: "'Cash Flow'!B1", text: "Capex" }]);
    expect(n.pairs[0]!.replace).toBe("Capital_expenditure");
    expect(notice("Capex", "X", null, { "Cash Flow": { B1: "Capex" } }, b)).toBeUndefined();
  });

  it("only resembling texts: says so and lists them", () => {
    const n = notice("Gross_income", "Gross_ind_income", null, { IS: { A4: "Gross income, adjusted" } })!;
    expect(n.stale).toEqual([]);
    expect(labelNoticeLines(n)).toEqual(["Gross_income → Gross_ind_income: no label reads the old name exactly", 'Check by eye (not in the replace): IS!A4 "Gross income, adjusted"']);
  });
});

describe("labelReplacement", () => {
  it("keeps the label's style", () => {
    expect(labelReplacement("Gross income", "Gross_income", "Gross_ind_income")).toBe("Gross ind income");
    expect(labelReplacement("Gross_income", "Gross_income", "Gross_ind_income")).toBe("Gross_ind_income");
    expect(labelReplacement("gross_income", "Gross_income", "Net")).toBe("Net");
    expect(labelReplacement("Gross-income", "Gross_income", "Gross_ind_income")).toBe("Gross-ind-income");
    expect(labelReplacement("2024 sales", "_2024_sales", "_2025_sales")).toBe("2025 sales");
    // Mixed separators: no style to keep, the name itself.
    expect(labelReplacement("Gross income-x", "Gross_income_x", "A_b_c")).toBe("A_b_c");
  });

  it("past the `_` Excel adds behind a reference-like label and the characters it drops at the end (F11)", () => {
    expect(labelReplacement("Q1", "Q1_", "Q2_")).toBe("Q2");
    expect(labelReplacement("Q1 sales", "Q1_sales", "Q2_sales")).toBe("Q2 sales");
    expect(labelReplacement("Margin %", "Margin", "Gross_margin")).toBe("Gross_margin");
    expect(labelReplacement("Net margin %", "Net_margin", "Gross_margin")).toBe("Gross margin");
    expect(labelReplacement("€uro rate", "_€uro_rate", "_€uro_spot")).toBe("€uro spot");
  });
});
