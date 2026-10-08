// One check at a time, on hand-made snapshots.
import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { audit, globMatch, nameFamilies, readWorkbook, type AuditOptions, type Finding } from "../../src/index.js";
import { book, type Book } from "./book.js";

const run = (b: Book, opts: AuditOptions = {}) => audit(book(b), opts);
const rules = (b: Book, check: string, opts: AuditOptions = {}): string[] =>
  run(b, opts)
    .findings.filter((f) => f.check === check)
    .map((f) => `${f.rule} ${f.where.kind === "name" ? f.where.key : `${f.where.sheet}!${f.where.ref ?? f.where.name}`}`);
const one = (b: Book, rule: string, opts: AuditOptions = {}): Finding => {
  const fs = run(b, opts).findings.filter((f) => f.rule === rule);
  expect(fs, rule).toHaveLength(1);
  return fs[0]!;
};

describe("C1 syntax", () => {
  it("reports where a definition or a cell stops parsing", () => {
    const b: Book = { sheets: { S: { A1: "SUM(1,", A2: "1+1" } }, names: [{ name: "Bad", def: "(1+\r\n2" }] };
    expect(rules(b, "C1")).toEqual(["C1.syntax Bad", "C1.syntax S!A1"]);
  });
  it("gives line and column", () => {
    const f = run({ sheets: { S: {} }, names: [{ name: "Bad", def: "(1+\r\n2" }] }).findings[0]!;
    expect(f.data).toMatchObject({ line: 2, col: 2 });
    expect(f.severity).toBe("error");
  });
});

describe("C2 prefix health", () => {
  it("finds bare modern functions, poisoned calls, wrong and unknown prefixes, everywhere", () => {
    const b: Book = {
      sheets: { S: { A1: "FILTER(B1:B3,B1:B3>0)", A2: "_xlfn.SUM(1)", A3: "_xlfn.FOOBAR(1)" } },
      cf: [{ sheet: "S", sqref: "B1:B3", formula: "_xludf.XLOOKUP(1,B1:B3,B1:B3)" }],
      names: [{ name: "Ok", def: "_xlfn.SEQUENCE(3)" }],
    };
    expect(rules(b, "C2")).toEqual(["C2.bare-prefix S!A1", "C2.wrong-prefix S!A2", "C2.unknown-function S!A3", "C2.poisoned S!B1:B3"]);
    expect(run(b).findings.find((f) => f.rule === "C2.unknown-function")!.severity).toBe("info");
  });
});

describe("C3 built-in collision", () => {
  it("errs on a LAMBDA, warns on a name called like the built-in, ignores one used as a value", () => {
    const b: Book = {
      sheets: { S: { A1: "Days(1,2)", A2: "Rate*2" } },
      names: [
        { name: "Fact", def: "_xlfn.LAMBDA(_xlpm.n, _xlpm.n)" },
        { name: "Days", def: "5" },
        { name: "Rate", def: "0.1" },
      ],
    };
    expect(rules(b, "C3")).toEqual(["C3.called Days", "C3.lambda Fact"]);
    expect(one(b, "C3.called").message).toContain("1 formula(s) call Days(…): they call DAYS");
  });
});

describe("C4 unresolved references", () => {
  const b: Book = {
    sheets: { S: { A1: "Nope+1", A2: "Elsewhere", B1: { text: "SEQUENCE(3)", spill: "B1:B3" } }, T: {} },
    names: [
      { name: "Err", def: "#N/A" },
      { name: "NoSpill", def: "_xlfn.ANCHORARRAY('S'!$C$9)" },
      { name: "NoTable", def: "SUM(tblX[a])" },
      { name: "Floating", def: "$A$1" },
      { name: "Ext", def: "[1]Sheet1!$A$1" },
      { name: "Elsewhere", def: "1", scope: "T" },
    ],
  };
  it("maps what the graph cannot resolve to rules", () => {
    expect(rules(b, "C4")).toEqual([
      "C4.error-definition Err",
      "C4.external Ext",
      "C4.unqualified-ref Floating",
      "C4.no-anchor NoSpill",
      "C4.table NoTable",
      "C4.unknown-name S!A1",
      "C4.unknown-name S!A2",
    ]);
  });
  it("says where an unknown name does exist", () => {
    const f = run(b).findings.find((x) => x.rule === "C4.unknown-name" && x.where.ref === "A2")!;
    expect(f.data).toMatchObject({ elsewhere: ["T"] });
  });
});

describe("C5 unqualified sheet-scoped reads", () => {
  it("errs on the own sheet's name unqualified, and on another sheet's name that only exists there", () => {
    const b: Book = {
      sheets: { S1: { A1: "Y+Z+W" }, S2: {} },
      names: [
        { name: "X", def: "1", scope: "S1" },
        { name: "Y", def: "X*2", scope: "S1" },
        { name: "Q", def: "'S1'!X*2", scope: "S1" },
        { name: "Z", def: "X*3" },
        { name: "V", def: "1", scope: "S2" },
        { name: "V", def: "2" },
        { name: "W", def: "V+1" },
      ],
    };
    expect(rules(b, "C5")).toEqual(["C5.other-sheet Z", "C5.own-sheet S1!Y"]);
    expect(one(b, "C5.own-sheet").hint).toBe("write 'S1'!X");
    // C4 does not report the same unknown name again.
    expect(rules(b, "C4")).toEqual([]);
  });
});

describe("C6 arity", () => {
  it("checks LAMBDA calls in names and cells, optional parameters, non-LAMBDA calls and built-ins", () => {
    const b: Book = {
      sheets: { S: { A1: "Two(1)", A2: "Two(1,2)", A3: "Opt(1,2)", A4: "Opt(1,2,3,4)", A5: "Val(1)", A6: "ROUND(1)" } },
      names: [
        { name: "Two", def: "_xlfn.LAMBDA(_xlpm.a,_xlpm.b, _xlpm.a+_xlpm.b)" },
        { name: "Opt", def: "_xlfn.LAMBDA(_xlpm.a,[_xlpm.b],[_xlpm.c], _xlpm.a)" },
        { name: "Val", def: "'S'!$B$1" },
        { name: "Uses", def: "Two(1,2,3)" },
        { name: "Shadow", def: "_xlfn.LET(_xlpm.Two, 1, _xlpm.Two)" },
      ],
    };
    // `Val` is a name on the empty cell S!B1: calling it is #VALUE!, a warning.
    expect(rules(b, "C6")).toEqual(["C6.lambda-arity Uses", "C6.lambda-arity S!A1", "C6.lambda-arity S!A4", "C6.not-a-lambda S!A5", "C6.builtin-arity S!A6"]);
    expect(run(b).findings.find((f) => f.where.ref === "A4")!.message).toBe("Opt(a, [b], [c]) takes 1 to 3 arguments; it is given 4");
    expect(run(b).findings.find((f) => f.rule === "C6.builtin-arity")!.severity).toBe("warning");
  });
  it("checks cells called like functions, and names on them (the cell holds a LAMBDA)", () => {
    const b: Book = {
      sheets: {
        S: {
          C2: "_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)",
          C3: "1+2",
          C4: "_xlfn.LET(_xlpm.k,2,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*_xlpm.k))",
          C5: "IF(TRUE,C2,C4)",
          F2: "C2(D2,E2)",
          F3: "$C$2(1)",
          F4: "C3(1,2)",
          F5: "B9(1)",
          F6: "C4(1)",
          F7: "C5(1,2)",
          F8: "T!A1(1)",
          G2: "Fake(D2,E2)",
          G3: "Fake(1,2,3)",
          G4: "Num(1)",
        },
        T: { A1: "_xlfn.LAMBDA(_xlpm.x,_xlpm.x)" },
      },
      names: [
        { name: "Fake", def: "'S'!$C$2" },
        { name: "Num", def: "'S'!$C$3" },
      ],
    };
    expect(rules(b, "C6")).toEqual(["C6.lambda-arity S!F3", "C6.lambda-arity S!G3", "C6.not-a-lambda S!F4", "C6.not-a-lambda S!G4", "C6.not-a-lambda S!F5"]);
    expect(run(b).findings.find((f) => f.where.ref === "F3")!.message).toBe("$C$2(a, b) takes 2 arguments; it is given 1");
    const f4 = run(b).findings.find((f) => f.where.ref === "F4")!;
    expect(f4.severity).toBe("warning");
    expect(f4.message).toBe("C3 is called with 2 argument(s), but S!C3 has a formula that gives a value, not a LAMBDA: Excel gives #VALUE! or #CALC!");
    expect(run(b).findings.find((f) => f.where.ref === "F5")!.message).toContain("S!B9 holds a value or nothing");
    expect(run(b).findings.find((f) => f.where.ref === "G3")!.severity).toBe("error");
    // Not unknown functions either.
    expect(rules(b, "C2")).toEqual([]);
  });
});

describe("C6 on a constructed workbook: a called cell that holds a number", () => {
  const M = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const rels = (entries: [string, string, string][]) =>
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries
      .map(([id, type, target]) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`)
      .join("")}</Relationships>`;
  const parts: Record<string, string> = {
    "_rels/.rels": rels([["rId1", "officeDocument", "xl/workbook.xml"]]),
    "xl/workbook.xml": `<workbook ${M} ${R}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Num">S!$C$2</definedName><definedName name="Fn">S!$C$3</definedName></definedNames></workbook>`,
    "xl/_rels/workbook.xml.rels": rels([["rId1", "worksheet", "worksheets/sheet1.xml"]]),
    "xl/worksheets/sheet1.xml": `<worksheet ${M} ${R}><sheetData>
      <row r="2"><c r="C2"><v>5</v></c><c r="F2"><f>C2(1)</f><v>0</v></c><c r="G2"><f>Num(1)</f><v>0</v></c></row>
      <row r="3"><c r="C3" t="e"><f>_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)</f><v>#CALC!</v></c><c r="F3"><f>C3(1)</f><v>2</v></c><c r="G3"><f>Fn(1,2)</f><v>0</v></c></row>
    </sheetData></worksheet>`,
  };
  const wb = readWorkbook(zipSync(Object.fromEntries(Object.entries(parts).map(([k, v]) => [k, strToU8(v)]))));
  it("warns on the number, checks the LAMBDA's arity, reports no unknown function", () => {
    const fs = audit(wb).findings.filter((f) => f.check === "C6" || f.check === "C2");
    expect(fs.map((f) => `${f.severity} ${f.rule} ${f.where.ref}`)).toEqual(["warning C6.not-a-lambda F2", "warning C6.not-a-lambda G2", "error C6.lambda-arity G3"]);
    expect(fs[0]!.message).toBe("C2 is called with 1 argument(s), but S!C2 holds a value or nothing, not a LAMBDA: Excel gives #VALUE! or #CALC!");
  });
});

describe("C7 limits", () => {
  it("warns near and errs over the length limit, for names and cells", () => {
    const b: Book = { sheets: { S: { A1: "1" + "+1".repeat(3800) } }, names: [{ name: "L", def: "1" + "+1".repeat(4100) }] };
    expect(rules(b, "C7")).toEqual(["C7.length L", "C7.length-near S!A1"]);
  });
  it("counts nesting of functions", () => {
    const nest = (n: number) => "ABS(".repeat(n) + "1" + ")".repeat(n);
    const b: Book = { sheets: { S: { A1: nest(50), A2: nest(65), A3: nest(40) } } };
    expect(rules(b, "C7")).toEqual(["C7.nesting-near S!A1", "C7.nesting S!A2"]);
    expect(rules(b, "C7", { limits: { nestingWarn: 30 } })).toEqual(["C7.nesting-near S!A1", "C7.nesting S!A2", "C7.nesting-near S!A3"]);
  });
});

describe("C8 census", () => {
  it("counts kinds and scopes, finds families and the tiers", () => {
    const b: Book = {
      sheets: { S: { B1: { text: "_xlfn.SEQUENCE(1,3,Start)", spill: "B1:D1" }, B2: "Get(\"x\")" }, T: {} },
      names: [
        { name: "Years", def: "_xlfn.ANCHORARRAY('S'!$B$1)" },
        { name: "Start", def: "2024" },
        { name: "Get", def: "_xlfn.LAMBDA(_xlpm.k, _xlfn.XLOOKUP(_xlpm.k, 'S'!$A$1:$A$3, 'S'!$B$1:$B$3))" },
        { name: "Pick", def: "_xlfn.LAMBDA(_xlpm.k, _xlpm.k)" },
        { name: "Sales_base", def: "1" },
        { name: "Sales_high", def: "2" },
        { name: "Cost_base", def: "1" },
        { name: "Cost_high", def: "2" },
        { name: "Cost_low", def: "0" },
        { name: "Item", def: "1", scope: "S" },
        { name: "Item", def: "1", scope: "T" },
        { name: "Chk", def: "1", scope: "T" },
      ],
    };
    const c = run(b, { census: { exclude: ["T!Chk"] } }).census;
    expect(c.total).toBe(12);
    expect(c.byKind).toMatchObject({ constant: 9, spill: 1, lambda: 2 });
    expect(c.byScope).toEqual({ workbook: 9, sheets: [{ sheet: "S", names: 1 }, { sheet: "T", names: 2 }] });
    expect(c.coordinates.map((x) => x.tag)).toEqual(["base", "high"]);
    expect(c.families.map((f) => `${f.stem}: ${f.members.join(" ")}`)).toEqual(["Cost: Cost_base Cost_high", "Sales: Sales_base Sales_high"]);
    expect(c.tierNames).toEqual({ T1: ["Get"], T4: ["Start", "Years"], library: ["Pick"], excluded: ["T!Chk"] });
    // Line items: Sales (2 names), Cost (2 + Cost_low on its own), Item on S and on T.
    expect(c.tiers).toMatchObject({ T1: 1, T2: 5, T3: 2, T4: 2, total: 10, of: 12, T2spellings: 4, library: 1, excluded: 1 });
    expect(c.sheetDuplicates).toEqual([{ name: "Item", sheets: ["S", "T"] }]);
  });
  it("can be told the tags, finds prefix families and composite suffixes", () => {
    const m = (name: string) => ({ key: name, name, scope: undefined });
    const fam = nameFamilies(["Q1_Gross_Sales", "Q2_Gross_Sales", "Q1_Unit_Cost", "Q2_Unit_Cost", "A_x", "A_y", "B_xExp", "B_yExp", "B_x", "B_y"].map(m));
    expect(fam.families.map((f) => `${f.position} ${f.stem}`)).toEqual(["suffix A", "suffix B", "suffix B_*Exp", "prefix Gross_Sales", "prefix Unit_Cost"]);
    expect(nameFamilies(["a_x", "a_y"].map(m), { tags: ["x", "y"] }).families).toHaveLength(1);
    expect(nameFamilies(["a_x", "a_y"].map(m)).families).toHaveLength(0);
  });
  it("matches glob patterns case-insensitively", () => {
    expect(globMatch("check!*", "Check!Tol")).toBe(true);
    expect(globMatch("*!FixedPoint_*", "SCF!FixedPoint_base")).toBe(true);
    expect(globMatch("FN.*", "IN.SET")).toBe(false);
    expect(globMatch("a*b*c", "abc")).toBe(true);
    expect(globMatch("a*b*c", "ac")).toBe(false);
  });
});

describe("C9 fixed references into a spill", () => {
  it("covers cells and names, the whole extent, a part, one cell and beyond", () => {
    const b: Book = {
      sheets: { S: { B1: { text: "_xlfn.SEQUENCE(1,3)", spill: "B1:D1" }, A2: "SUM(B1:D1)", A3: "C1*2", A4: "SUM(A1:D1)", A5: "B1" } },
      names: [{ name: "Part", def: "'S'!$B$1:$C$1" }, { name: "Whole", def: "_xlfn.ANCHORARRAY('S'!$B$1)" }],
    };
    expect(rules(b, "C9")).toEqual(["C9.fixed-ref Part", "C9.fixed-ref S!A2", "C9.fixed-ref S!A3", "C9.fixed-ref S!A4"]);
    expect(run(b).findings.find((f) => f.where.ref === "A3")!.hint).toBe("use INDEX(B1#, 2)");
    const r = run(b).spills;
    expect(r.spills[0]!.names).toEqual([{ key: "Part", how: "part" }, { key: "Whole", how: "spill" }]);
    expect(r.bySheet).toEqual([{ sheet: "S", spills: 1, spillNamed: 1, fixedNamed: 0, unnamed: 0 }]);
  });
});

describe("C10 unused names", () => {
  it("counts formats, validations, Table columns, charts and other names as uses", () => {
    const b: Book = {
      sheets: { S: { A1: "1" }, C: {} },
      tables: [{ name: "tbl", sheet: "S", ref: "D1:E3", columns: ["a", { name: "b", formula: "tbl[[#This Row],[a]]*InTable" }] }],
      cf: [{ sheet: "S", sqref: "A1", formula: "A1>InCf" }],
      dv: [{ sheet: "S", sqref: "A1", formula: "InDv" }],
      charts: [{ sheet: "C", formulas: ["[0]!InChart", "'S'!Local"] }],
      names: [
        { name: "InTable", def: "1" },
        { name: "InCf", def: "1" },
        { name: "InDv", def: "'S'!$A$1:$A$2" },
        { name: "InChart", def: "Under*2" },
        { name: "Under", def: "1" },
        { name: "Local", def: "'S'!$A$1", scope: "S" },
        { name: "Dead", def: "_xlfn.LAMBDA(_xlpm.x, _xlpm.x+Deader)" },
        { name: "Deader", def: "1" },
        { name: "Ghost", def: "1", hidden: true },
      ],
    };
    expect(rules(b, "C10")).toEqual(["C10.unused Dead", "C10.only-by-unused Deader", "C10.unused-hidden Ghost"]);
    expect(run(b).findings.find((f) => f.rule === "C10.only-by-unused")!.message).toBe("Deader is read only by unused names (Dead)");
  });

  it("skips the harness and what only it reads, and leaves the harness out of the census tiers", () => {
    const b: Book = {
      sheets: { Check: {}, Model: { A1: "1" } },
      names: [
        { name: "SelfTest", def: "AND(Probe=1, Probe2=2)", scope: "Check" },
        { name: "Probe", def: "1" },
        { name: "Probe2", def: "2" },
        { name: "CHK.All", def: "1" },
        { name: "FixCap", def: "100", scope: "Model" },
        { name: "FixedPoint_tol", def: "0.001", scope: "Model" },
        { name: "Dead", def: "1" },
      ],
    };
    // SelfTest reads Probe and Probe2: an end result, not unused (and they are read).
    expect(rules(b, "C10").sort()).toEqual(["C10.unused CHK.All", "C10.unused Dead", "C10.unused Model!FixCap", "C10.unused Model!FixedPoint_tol"].sort());
    const harness = ["Check!*", "chk.*", "Model!Fix*", "*!FixedPoint_*"];
    expect(rules(b, "C10", { harness })).toEqual(["C10.unused Dead"]);
    const c = run(b, { harness, census: { exclude: ["Dead"] } }).census;
    expect(c.tierNames.excluded).toEqual(["Check!SelfTest", "CHK.All", "Dead", "Model!FixCap", "Model!FixedPoint_tol"]);
    expect(c.tiers.excluded).toBe(5);
  });
});

describe("C10 end results", () => {
  it("does not report a name that reads something, nor what it reads; keeps constants and LAMBDAs, and names on value cells as info", () => {
    const b: Book = {
      sheets: { IS: { C9: { text: "EBIT", spill: "C9:G9" }, C10: { text: "Taxes", spill: "C10:G10" }, C11: { text: "EbIT-Taxes", spill: "C11:G11" } }, In: {} },
      tables: [{ name: "tblIn", sheet: "In", ref: "A1:B4", columns: ["a", "b"] }],
      names: [
        { name: "EBIT", def: "1" },
        { name: "Taxes", def: "'IS'!$C$9#*0.3" },
        { name: "Unlevered_net_income", def: "_xlfn.ANCHORARRAY('IS'!$C$11)" },
        { name: "Total", def: "SUM('In'!$A$2:$A$4)" },
        { name: "ColSum", def: "SUM(tblIn[b])" },
        { name: "Inputs", def: "'In'!$A$2:$A$4" },
        { name: "Empty", def: "'In'!$H$1" },
        { name: "Unused", def: "0.3" },
        { name: "Fn", def: "_xlfn.LAMBDA(_xlpm.x, _xlpm.x*Rate)" },
        { name: "Rate", def: "2" },
      ],
    };
    // The author's single-sheet IS: Unlevered_net_income @C11# = EbIT - Taxes reads cells.
    expect(rules(b, "C10").sort()).toEqual(["C10.unused-cell Empty", "C10.unused Fn", "C10.unused-cell Inputs", "C10.unused Unused", "C10.only-by-unused Rate"].sort());
    // A name on cells may be read by a person on the sheet: info, not a warning.
    const sev = run(b).findings.filter((f) => f.check === "C10").map((f) => `${f.severity} ${f.where.key}`).sort();
    expect(sev).toEqual(["info Empty", "info Inputs", "info Rate", "warning Fn", "warning Unused"]);
  });
});

describe("C11 copy drift", () => {
  it("compares anchor formulas with relative references as offsets", () => {
    const b: Book = {
      sheets: { S: { C6: "C5*2", C16: "C15*2", C26: "C25*3", C36: "C35*2" } },
      names: [
        { name: "X_base", def: "'S'!$C$6" },
        { name: "X_high", def: "'S'!$C$16" },
        { name: "X_low", def: "'S'!$C$26" },
        { name: "X_mid", def: "'S'!$C$36" },
        { name: "Y_base", def: "X_base*Rate_base" },
        { name: "Y_high", def: "X_high*Rate_high" },
        { name: "Y_low", def: "X_low*Rate_low" },
        { name: "Y_mid", def: "X_mid*Rate_mid" },
        { name: "Rate_base", def: "0.1" },
        { name: "Rate_high", def: "0.2" },
      ],
    };
    const f = one(b, "C11.drift");
    expect(f.where.key).toBe("X_low");
    expect(f.message).toBe("X_low (its formula at S!C26) differs from X_base, X_high, X_mid beyond the suffix (low vs base): `3` here, `2` there");
  });
  it("finds drift in prefix families and Table columns named by the tag", () => {
    const b: Book = {
      sheets: { S: {} },
      tables: [{ name: "tbl", sheet: "S", ref: "A1:C3", columns: ["base", "high"] }],
      names: [
        { name: "base_Net_Sales", def: "SUM(tbl[base])" },
        { name: "high_Net_Sales", def: "SUM(tbl[high])" },
        { name: "base_Unit_Cost", def: "MAX(tbl[base])" },
        { name: "high_Unit_Cost", def: "MIN(tbl[high])" },
      ],
    };
    expect(rules(b, "C11")).toEqual(["C11.drift high_Unit_Cost"]);
    expect(one(b, "C11.drift").message).toContain("`MIN` here, `MAX` there");
  });

  // FEEDBACK 2026-10-07, is-model's Mortgage sheet: an opening and a closing balance.
  const mortgage: Book = {
    sheets: {
      Mortgage: {
        B17: { text: "_xlfn.VSTACK(amount,_xlfn.DROP(final_amount,-1,0))", spill: "B17:B28" },
        E17: { text: "_xlfn.SEQUENCE(12,1,amount/12,0)", spill: "E17:E28" },
        G17: { text: "_xlfn.SCAN(amount,repayment,_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a-_xlpm.b))", spill: "G17:G28" },
      },
    },
    names: [
      { name: "amount", def: "'Mortgage'!$B$12" },
      { name: "initial_amount", def: "_xlfn.ANCHORARRAY('Mortgage'!$B$17)", scope: "Mortgage" },
      { name: "repayment", def: "_xlfn.ANCHORARRAY('Mortgage'!$E$17)", scope: "Mortgage" },
      { name: "final_amount", def: "_xlfn.ANCHORARRAY('Mortgage'!$G$17)", scope: "Mortgage" },
      // Over the cells of another spill: nothing to compare.
      { name: "initial_amount_f", def: "'Mortgage'!$J$17" },
      { name: "final_amount_f", def: "'Mortgage'!$O$17" },
    ],
  };
  it("does not take a prefix whose families all differ for a coordinate (initial_, final_)", () => {
    const r = run(mortgage);
    expect(r.findings.filter((f) => f.check === "C11")).toEqual([]);
    expect(r.census.coordinates).toEqual([]);
    expect(r.census.families).toEqual([]);
    // Four line items, not two items spelled twice.
    expect(r.census.tiers).toMatchObject({ T3: 0 });
  });
  it("still warns on all-differing families when the tags are forced", () => {
    const opts: AuditOptions = { census: { tags: ["initial", "final"] } };
    expect(rules(mortgage, "C11", opts)).toEqual(["C11.drift Mortgage!initial_amount"]);
    expect(run(mortgage, opts).census.coordinates.map((c) => c.tag)).toEqual(["final", "initial"]);
  });
  it("keeps a coordinate whose copies agree, and the member that drifts", () => {
    const b: Book = {
      sheets: { S: {} },
      names: [
        { name: "Vol_base", def: "100" },
        { name: "Vol_payout", def: "120" },
        { name: "Price", def: "3" },
        { name: "Sales_base", def: "Vol_base*Price" },
        { name: "Sales_payout", def: "Vol_payout*Price" },
        { name: "Cost_base", def: "Sales_base*0.6" },
        { name: "Cost_payout", def: "Sales_payout*0.6" },
        { name: "Tax_base", def: "(Sales_base-Cost_base)*0.3" },
        { name: "Tax_payout", def: "(Sales_payout-Cost_payout)*0.25" },
      ],
    };
    expect(rules(b, "C11")).toEqual(["C11.drift Tax_payout"]);
    expect(run(b).census.coordinates.map((c) => `${c.tag} ${c.families}`)).toEqual(["base 4", "payout 4"]);
  });
  it("judges inferred tags with the comparison it is given", () => {
    const m = (name: string) => ({ key: name, name, scope: undefined });
    const names = ["A_x", "A_y", "B_x", "B_y", "C_x", "C_z", "D_x", "D_z", "E_w", "E_v", "F_w", "F_v"].map(m);
    // C's and D's names all differ (so z goes, and the prefixes C_, D_ the names left would
    // take); over w and v nothing can be compared; A and B agree.
    const agrees = (f: { stem: string; members: string[] }) => (f.stem === "E" || f.stem === "F" ? undefined : !f.members.some((k) => k.startsWith("C") || k.startsWith("D")));
    const r = nameFamilies(names, {}, agrees);
    expect(r.coordinates.map((c) => `${c.tag} ${c.families}`)).toEqual(["v 2", "w 2", "x 2", "y 2"]);
    expect(r.families.map((f) => f.stem)).toEqual(["A", "B", "E", "F"]);
    expect(nameFamilies(names, { tags: ["x", "y", "z"] }, agrees).families.map((f) => f.stem)).toEqual(["A", "B", "C", "D"]);
  });
});

describe("C12 cycles", () => {
  it("errs on names reading each other, notes circular references among cells, allows LAMBDA recursion", () => {
    const b: Book = {
      sheets: { S: { A1: "B1+1", B1: "A1+1" } },
      names: [
        { name: "P", def: "Q" },
        { name: "Q", def: "P" },
        { name: "Rec", def: "_xlfn.LAMBDA(_xlpm.n, IF(_xlpm.n<1, 0, Rec(_xlpm.n-1)))" },
      ],
    };
    expect(rules(b, "C12")).toEqual(["C12.name-cycle P", "C12.circular S!A1"]);
    expect(one(b, "C12.circular").severity).toBe("info");
  });
});

describe("C13 constants in LAMBDA bodies", () => {
  it("lists the numbers outside the allow-list, negatives and percents included", () => {
    const b: Book = {
      sheets: { S: { A1: "_xlfn.LAMBDA(_xlpm.x, _xlpm.x*0.3)(1)" } },
      names: [
        { name: "Ok", def: "_xlfn.LAMBDA(_xlpm.x, _xlpm.x*365/12-1)" },
        { name: "Tax", def: "_xlfn.LAMBDA(_xlpm.x, _xlpm.x*0.27 + -0.27 + 5%)" },
        { name: "NotLambda", def: "0.27" },
      ],
    };
    expect(rules(b, "C13")).toEqual(["C13.constant Tax"]);
    expect(one(b, "C13.constant").message).toBe("0.27, -0.27, 5% hard-coded in the LAMBDA body");
    expect(rules(b, "C13", { constants: { allow: [0.27, 0.05] } })).toEqual(["C13.constant Ok"]);
    expect(rules(b, "C13", { constants: { allow: [] } })).toEqual(["C13.constant Ok", "C13.constant Tax"]);
  });

  it("allows sentinels from 1E+90 up, configurable", () => {
    const b: Book = {
      sheets: { S: {} },
      names: [
        { name: "FN.DEV", def: "_xlfn.LAMBDA(_xlpm.x, IF(_xlpm.x, 1E+99, -1E+99))" },
        { name: "FN.FIXPOINT", def: "_xlfn.LAMBDA(_xlpm.x, MIN(_xlpm.x, 1E+300))" },
        { name: "Big", def: "_xlfn.LAMBDA(_xlpm.x, _xlpm.x*1E+89)" },
      ],
    };
    expect(rules(b, "C13")).toEqual(["C13.constant Big"]);
    expect(rules(b, "C13", { constants: { sentinelAbove: 1e100 } })).toEqual(["C13.constant Big", "C13.constant FN.DEV"]);
    expect(rules(b, "C13", { constants: { sentinelAbove: 1e80 } })).toEqual([]);
  });
});

describe("ordering and summary", () => {
  it("orders by check, place kind, sheet, row, column", () => {
    const b: Book = {
      sheets: { B: { A2: "SUM(1,", A1: "SUM(1," }, A: { C1: "SUM(1," } },
      names: [{ name: "z", def: "(" }, { name: "a", def: "(", scope: "A" }],
    };
    expect(rules(b, "C1")).toEqual(["C1.syntax z", "C1.syntax A!a", "C1.syntax B!A1", "C1.syntax B!A2", "C1.syntax A!C1"]);
    const r = run(b);
    expect(r.counts.error).toBe(5);
    expect(r.byCheck.C1).toEqual({ error: 5, warning: 0, info: 0 });
  });
});
