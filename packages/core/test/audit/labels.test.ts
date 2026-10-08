// C15: a name whose label cell no longer gives it (decided 2026-10-07). The author names
// cells with Create from Selection, then renames a name (F2, xln rename); the label in
// Excel still reads the old text. Hand-made snapshots, with the cells' values beside them.
import { describe, expect, it } from "vitest";
import { audit, labelName, labelNames, type AuditOptions, type CellValue, type Finding } from "../../src/index.js";
import { book, type Book } from "./book.js";

type Values = Record<string, Record<string, CellValue>>;

const valueMap = (v: Values) => new Map(Object.entries(v).map(([s, cells]) => [s, new Map(Object.entries(cells))]));

function c15(b: Book, v: Values | undefined, opts: AuditOptions = {}): Finding[] {
  return audit(book(b), { ...opts, ...(v ? { values: valueMap(v) } : {}) }).findings.filter((f) => f.check === "C15");
}

const keys = (fs: Finding[]) => fs.map((f) => f.where.key);

// An income statement made with Create from Selection (Left column) over A2:D5, then
// Gross_income renamed Gross_ind_income.
const IS: Book = {
  sheets: { IS: { B4: "B2-B3", C4: "C2-C3", D4: "D2-D3", B5: "B4*0.7", C5: "C4*0.7", D5: "D4*0.7" } },
  names: [
    { name: "Revenue", def: "IS!$B$2:$D$2" },
    { name: "COGS", def: "IS!$B$3:$D$3" },
    { name: "Gross_ind_income", def: "IS!$B$4:$D$4" },
    { name: "Net_income", def: "IS!$B$5:$D$5" },
  ],
};
const IS_VALUES: Values = { IS: { A2: "Revenue", A3: "COGS", A4: "Gross income", A5: "Net income", B2: 100, C2: 110, D2: 120, B3: 60, C3: 66, D3: 72 } };

describe("C15 labels that no longer match their names", () => {
  it("reports a renamed name whose label still reads the old text, with the probable old name", () => {
    const fs = c15(IS, IS_VALUES);
    expect(keys(fs)).toEqual(["Gross_ind_income"]);
    const f = fs[0]!;
    expect(f.rule).toBe("C15.label-drift");
    expect(f.severity).toBe("info");
    expect(f.where).toMatchObject({ kind: "name", name: "Gross_ind_income" });
    expect(f.message).toBe('IS!A4 reads "Gross income", which Create from Selection makes Gross_income; the name on IS!B4:D4 is Gross_ind_income, probably renamed from Gross_income');
    expect(f.hint).toBe('xln never writes cell values: in Excel, Find & Replace "Gross income" with "Gross ind income" (Ctrl+H, ⌘⇧H on Mac; Match entire cell contents), or rename the name back. The other names in its column of labels match theirs (Revenue, COGS, Net_income)');
    expect(f.data).toMatchObject({ cell: "IS!A4", label: "Gross income", converted: "Gross_income", renamedFrom: "Gross_income" });
  });

  it("says no old name when a name of the label's spelling exists", () => {
    const b: Book = { ...IS, names: [...IS.names!, { name: "Gross_income", def: "42" }] };
    const f = c15(b, IS_VALUES)[0]!;
    expect(f.message).toBe('IS!A4 reads "Gross income", which Create from Selection makes Gross_income; the name on IS!B4:D4 is Gross_ind_income');
    expect(f.data).not.toHaveProperty("renamedFrom");
  });

  it("Top row: a column of names under a row of labels", () => {
    const b: Book = {
      sheets: { P: {} },
      names: [
        { name: "Price", def: "P!$B$2:$B$9" },
        { name: "Qty", def: "P!$C$2:$C$9" },
        { name: "Amount_due", def: "P!$D$2:$D$9" },
      ],
    };
    const fs = c15(b, { P: { B1: "Price", C1: "Qty", D1: "Amount" } });
    expect(keys(fs)).toEqual(["Amount_due"]);
    expect(fs[0]!.message).toContain("P!D1 reads \"Amount\"");
    expect(fs[0]!.hint).toContain("in its row of labels");
  });

  it("names that never matched their labels (named by hand) give nothing", () => {
    const b: Book = {
      sheets: { IS: {} },
      names: [
        { name: "Rev", def: "IS!$B$2:$D$2" },
        { name: "Cogs_total", def: "IS!$B$3:$D$3" },
        { name: "GI", def: "IS!$B$4:$D$4" },
      ],
    };
    expect(c15(b, { IS: { A2: "Revenue (EUR)", A3: "Cost of goods", A4: "Gross income" } })).toEqual([]);
  });

  it("more names that differ than names that match: nothing (not enough evidence)", () => {
    const b: Book = {
      sheets: { IS: {} },
      names: [
        { name: "Revenue", def: "IS!$B$2:$D$2" },
        { name: "Cogs_total", def: "IS!$B$3:$D$3" },
        { name: "GI", def: "IS!$B$4:$D$4" },
      ],
    };
    expect(c15(b, { IS: { A2: "Revenue", A3: "Cost of goods", A4: "Gross income" } })).toEqual([]);
  });

  it("evidence comes from the same column of labels and the same span only", () => {
    const b: Book = {
      sheets: { IS: {} },
      names: [
        { name: "Revenue", def: "IS!$B$2:$D$2" },
        { name: "COGS", def: "IS!$B$3:$D$3" },
        // Another span (B:E), and another label column (G): no matching name beside them.
        { name: "Gross_ind_income", def: "IS!$B$4:$E$4" },
        { name: "Margin", def: "IS!$H$2:$J$2" },
      ],
    };
    expect(c15(b, { IS: { A2: "Revenue", A3: "COGS", A4: "Gross income", G2: "Gross margin" } })).toEqual([]);
  });

  it("an alias on the same cells, a corner name, a computed label, a header row of text: nothing", () => {
    const b: Book = {
      sheets: { IS: { A6: '"Tax"&"es"' } },
      names: [
        { name: "Revenue", def: "IS!$B$2:$D$2" },
        { name: "Sales", def: "IS!$B$2:$D$2" }, // alias: Revenue, on the same cells, matches A2
        { name: "COGS", def: "IS!$B$3:$D$3" },
        { name: "Block", def: "IS!$B$2:$D$3" }, // two-way corner (A1)
        { name: "Tax_paid", def: "IS!$B$6:$D$6" }, // A6 is a formula: the pull's note, not C15
        { name: "Years", def: "IS!$B$1:$D$1" }, // a header row of text under "Line item"
        { name: "Net", def: "IS!$B$7:$D$7" },
        { name: "Other", def: "IS!$B$8:$D$8" },
      ],
    };
    const v: Values = { IS: { A1: "Block", A2: "Revenue", A3: "COGS", A6: "Taxes", A7: "Net", A8: "Other", B1: "FY24", C1: "FY25", D1: "FY26", B2: 1, B3: 2 } };
    expect(c15(b, v)).toEqual([]);
    // Without the alias, Sales is reported.
    expect(keys(c15({ ...b, names: b.names!.filter((n) => n.name !== "Revenue") }, v))).toEqual(["Sales"]);
  });

  it("single cells: the label to the left or above", () => {
    const b: Book = {
      sheets: { In: {} },
      names: [
        { name: "Tax_rate", def: "In!$B$2" },
        { name: "Start_year", def: "In!$B$3" },
        { name: "Growth", def: "In!$B$4" },
      ],
    };
    const fs = c15(b, { In: { A2: "Tax rate", A3: "Start year", A4: "Growth rate", B2: 0.24, B3: 2025, B4: 0.03 } });
    expect(keys(fs)).toEqual(["Growth"]);
    expect(fs[0]!.data).toMatchObject({ cell: "In!A4", converted: "Growth_rate" });
  });

  it("a spill name reads the label before its anchor, over the saved extent", () => {
    const b: Book = {
      sheets: { S: { B2: { text: "SEQUENCE(1,3)", spill: "B2:D2" }, B3: { text: "SEQUENCE(1,3)", spill: "B3:D3" } } },
      names: [
        { name: "Years", def: "_xlfn.ANCHORARRAY(S!$B$2)" },
        { name: "Volume", def: "_xlfn.ANCHORARRAY(S!$B$3)" },
      ],
    };
    expect(keys(c15(b, { S: { A2: "Years", A3: "Units", B2: 1, C2: 2, D2: 3, B3: 1, C3: 2, D3: 3 } }))).toEqual(["Volume"]);
  });

  it("needs the cells' values; respects the rule switches", () => {
    expect(c15(IS, undefined)).toEqual([]);
    expect(c15(IS, IS_VALUES, { rules: { C15: "off" } })).toEqual([]);
    expect(c15(IS, IS_VALUES, { rules: { "C15.label-drift": "warning" } })[0]!.severity).toBe("warning");
    expect(c15(IS, IS_VALUES, { minSeverity: "warning" })).toEqual([]);
    expect(c15(IS, IS_VALUES, { only: ["C15"] })).toHaveLength(1);
  });
});

describe("labelName: Create from Selection's rule as probe F11 measured it", () => {
  // Label → Excel's name (Excel for Mac, probes/results/f11_create_from_selection_mac.xlsx).
  const F11: [CellValue, string | undefined][] = [
    ["Gross  income", "Gross__income"],
    ["Gross income", "Gross_income"],
    ["a - b", "a___b"],
    ["x/y/z", "x_y_z"],
    ["Q1", "Q1_"],
    ["Tax2024", "Tax2024_"],
    ["R", "R_"],
    ["C", "C_"],
    ["R1C1", "_R1C1"],
    ["A1B", "A1B"],
    ["rc", "rc_"],
    ["2024 sales", "_2024_sales"],
    ["  lead", "lead"],
    ["trail  ", "trail"],
    ["Crescità", "Crescità"],
    ["Café", "Café"],
    ["Margin %", "Margin"],
    ["R&D", "R_D"],
    ["v1.2", "v1.2"],
    ["Why?", "Why?"],
    ["back\\slash", "back\\slash"],
    ["_under", "_under"],
    ["Total", "Total"],
    ["", undefined],
    [2024, undefined],
    [3.5, undefined],
    ["True", "True_"],
    ["line\r\nbreak", "line_break"],
    ["€uro", "_€uro"],
    ["中文", "中文"],
    ["1st", "_1st"],
    ["Long" + "x".repeat(296), undefined],
  ];
  it.each(F11)("%j → %s", (label, name) => {
    expect(labelName(label)).toBe(name);
    // A whole number may be a date shown as one: labelNames tries those spellings.
    if (typeof label === "string") expect(labelNames(label)[0]).toBe(name);
  });

  it("a date gives the text it shows, when the caller knows it", () => {
    expect(labelName(45322, "31/01/2024")).toBe("_31_01_2024");
    expect(labelName(45322)).toBeUndefined();
    // Without the shown text, the common short-date spellings are all accepted.
    expect(labelNames(45322)).toEqual(expect.arrayContaining(["_31_01_2024", "_1_31_2024", "_2024_01_31", "_31.01.2024"]));
    expect(labelNames(3.5)).toEqual([]);
  });

  it("characters a name cannot hold, at the start: `_` each, or dropped (not measured, both accepted)", () => {
    expect(labelNames("(a) b")).toEqual(["_a__b", "a__b"]);
    expect(labelNames("%")).toEqual([]);
  });

  it("no name from an empty cell, an error or a boolean", () => {
    for (const v of ["", "   ", { error: "#N/A" }, true, undefined]) expect(labelNames(v as CellValue)).toEqual([]);
  });
});
