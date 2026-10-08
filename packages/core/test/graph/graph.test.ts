// The dependency graph on small synthetic workbooks: one reference kind at a time.
import { describe, expect, it } from "vitest";
import {
  buildGraph,
  parseStructInner,
  renderFormulaView,
  sheetCalcView,
  workbookFormulaView,
  type CellFormula,
  type DefinedName,
  type DependencyGraph,
  type Sheet,
  type Table,
  type WorkbookSnapshot,
} from "../../src/index.js";

type F = string | { text: string; spill?: string; kind?: CellFormula["kind"] };

interface Book {
  sheets: Record<string, Record<string, F>>;
  names?: { name: string; def: string; scope?: string }[];
  tables?: { name: string; sheet: string; ref: string; columns: string[]; totals?: boolean }[];
}

function book(b: Book): WorkbookSnapshot {
  const order = Object.keys(b.sheets);
  const sheets: Sheet[] = order.map((name, position) => {
    const formulas: CellFormula[] = [];
    const spills: Sheet["spills"] = [];
    for (const [cell, f] of Object.entries(b.sheets[name]!)) {
      const x = typeof f === "string" ? { text: f } : f;
      const kind = x.kind ?? (x.spill ? "dynamic-array" : "normal");
      const cf: CellFormula = { cell, kind, text: x.text, attributes: {}, value: { type: "n", raw: "1", value: 1 } };
      if (x.spill) {
        cf.range = x.spill;
        spills.push({ anchor: cell, extent: x.spill });
      }
      formulas.push(cf);
    }
    return { name, sheetId: position + 1, position, state: "visible", kind: "worksheet", relId: `rId${position + 1}`, part: `xl/worksheets/sheet${position + 1}.xml`, formulas, sharedFormulas: [], spills, conditionalFormats: [], dataValidations: [], tables: [] };
  });
  const definedNames: DefinedName[] = (b.names ?? []).map((n, index) => {
    const position = n.scope === undefined ? -1 : order.indexOf(n.scope);
    return {
      name: n.name,
      scope: n.scope === undefined ? { kind: "workbook" } : { kind: "sheet", position, name: n.scope },
      hidden: false,
      comment: undefined,
      definition: n.def,
      attributes: {},
      index,
      isXlPrefixed: false,
      isBuiltIn: false,
    };
  });
  const tables: Table[] = (b.tables ?? []).map((t, k) => ({
    id: k + 1,
    name: t.name,
    displayName: t.name,
    sheet: { position: order.indexOf(t.sheet), name: t.sheet },
    part: `xl/tables/table${k + 1}.xml`,
    ref: t.ref,
    headerRowCount: 1,
    totalsRowCount: t.totals ? 1 : 0,
    columns: t.columns.map((name, id) => ({ id: id + 1, name })),
  }));
  return { sheets, definedNames, tables, charts: [], otherParts: [], parts: [], foreignModuleStores: [], workbookPart: "xl/workbook.xml", warnings: [] };
}

/** Labels of what `key` (a formula cell `S!A1` or a name `name:X`) depends on. */
function prec(g: DependencyGraph, key: string): string[] {
  const n = g.node(key);
  if (!n) throw new Error(`no node ${key}`);
  return g.precedents(n).map((p) => p.label).sort();
}

describe("buildGraph: references", () => {
  it("cells, areas and inputs: formulas become edges, the rest one input node per rectangle", () => {
    const g = buildGraph(book({ sheets: { S: { A1: "1+1", A2: "A1*2", A3: "SUM(A1:A2)", A4: "SUM(A1:B2)", A5: "$B$9+B9" } } }));
    expect(prec(g, "S!A2")).toEqual(["S!A1"]);
    expect(prec(g, "S!A3")).toEqual(["S!A1", "S!A2"]); // fully covered: no input node
    expect(prec(g, "S!A4")).toEqual(["S!A1", "S!A1:B2", "S!A2"]);
    expect(prec(g, "S!A5")).toEqual(["S!B9"]);
    expect(g.node("in:S!B9")!.kind).toBe("input");
  });

  it("a cell called like a function reads the cell (its LAMBDA)", () => {
    const g = buildGraph(book({ sheets: { S: { C2: "_xlfn.LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)", F2: "C2(D2,E2)", F3: "$C$2(1,2)+T!B1(1)" }, T: { B1: "_xlfn.LAMBDA(_xlpm.x,_xlpm.x)" } } }));
    expect(prec(g, "S!F2")).toEqual(["S!C2", "S!D2", "S!E2"]);
    expect(prec(g, "S!F3")).toEqual(["S!C2", "T!B1"]);
  });

  it("whole columns and rows: one query, one input node", () => {
    const sheet: Record<string, F> = { C1: "SUM(A:A)", D1: "SUM(2:2)" };
    for (let r = 1; r <= 200; r++) sheet[`A${r + 2}`] = `${r}`;
    const g = buildGraph(book({ sheets: { S: sheet } }));
    expect(g.precedents(g.node("S!C1")!).length).toBe(201);
    expect(prec(g, "S!C1")).toContain("S!A:A");
    expect(prec(g, "S!D1")).toEqual(["S!2:2"]);
  });

  it("x# goes to the dynamic array; a fixed area over it is a C9 finding", () => {
    const g = buildGraph(
      book({
        sheets: {
          S: { C1: { text: "_xlfn.SEQUENCE(1,5)", spill: "C1:G1" }, A2: "SUM(C1#)", A3: "SUM(C1:G1)", A4: "D1*2", A5: "C1", A6: "SUM(_xlfn.ANCHORARRAY(C1))" },
          T: { A1: "SUM(S!C1:E1)", A2: "Z9#" },
        },
      }),
    );
    expect(prec(g, "S!A2")).toEqual(["S!C1#"]);
    expect(prec(g, "S!A6")).toEqual(["S!C1#"]);
    expect(prec(g, "S!A3")).toEqual(["S!C1#"]);
    expect(g.spillRefs.map((f) => [f.node.label, f.ref, f.use, f.fit])).toEqual([
      ["S!A3", "C1:G1", "C1#", "exact"],
      ["S!A4", "D1", "INDEX(C1#, 2)", "part"],
      ["T!A1", "S!C1:E1", "S!C1#", "part"],
    ]);
    // `C1` alone is the anchor's first value, not a fixed reference into the spill.
    expect(g.spillRefs.some((f) => f.node.label === "S!A5")).toBe(false);
    expect(g.node("T!A2")!.flags).toMatchObject([{ kind: "broken" }]);
  });

  it("names: cell, range, spill, formula and LAMBDA names, with sheet scope", () => {
    const g = buildGraph(
      book({
        sheets: { S: { C1: { text: "_xlfn.SEQUENCE(3)", spill: "C1:C3" }, A1: "Rate*Total", A2: "Fn(1)+Spill", A3: "Loc" }, T: { A1: "Loc" } },
        names: [
          { name: "Rate", def: "S!$B$1" },
          { name: "Spill", def: "_xlfn.ANCHORARRAY(S!$C$1)" },
          { name: "Total", def: "SUM(Spill)*Rate" },
          { name: "Fn", def: "_xlfn.LAMBDA(_xlpm.x, _xlpm.x*Rate+x)" },
          { name: "x", def: "S!$B$2" },
          { name: "Loc", def: "S!$B$3", scope: "S" },
          { name: "Loc", def: "T!$B$3" },
        ],
      }),
    );
    expect(prec(g, "S!A1")).toEqual(["Rate", "Total"]);
    expect(prec(g, "name:Rate")).toEqual(["S!B1"]);
    expect(prec(g, "name:Spill")).toEqual(["S!C1#"]);
    expect(prec(g, "name:Total")).toEqual(["Rate", "Spill"]);
    // A LAMBDA depends on what its body reads, not on its parameter (`x` the parameter
    // shadows the name x; the second, outside the LAMBDA's own scope... is inside: no edge).
    expect(prec(g, "name:Fn")).toEqual(["Rate"]);
    expect(prec(g, "S!A2")).toEqual(["Fn", "Spill"]);
    expect(prec(g, "S!A3")).toEqual(["S!Loc"]);
    expect(prec(g, "T!A1")).toEqual(["Loc"]);
    expect(g.node("S!A1")!.level).toBe(2); // C1 (1) → Spill → Total → A1; names add no level
  });

  it("C9: a reference reaching past the spill, a cell of a 2-D spill, a name over a spill", () => {
    const g = buildGraph(
      book({
        sheets: { S: { B2: { text: "_xlfn.SEQUENCE(2,3)", spill: "B2:D3" }, A5: "SUM(A1:D3)", A6: "C3", A7: "B2#", A8: "Whole+Fixed" } },
        names: [
          { name: "Whole", def: "_xlfn.ANCHORARRAY(S!$B$2)" },
          { name: "Fixed", def: "S!$B$2:$D$3" },
        ],
      }),
    );
    expect(g.spillRefs.map((f) => [f.node.label, f.ref, f.use, f.fit])).toEqual([
      ["S!A5", "A1:D3", "B2#", "beyond"],
      ["S!A6", "C3", "INDEX(B2#, 2, 2)", "part"],
      ["Fixed", "S!$B$2:$D$3", "S!B2#", "exact"],
    ]);
  });

  it("a Table's own name reads its data rows", () => {
    const g = buildGraph(
      book({
        sheets: { S: { A1: "SUM(tbl)", A2: "Tot*2" } },
        names: [{ name: "Tot", def: "SUM(tbl)" }],
        tables: [{ name: "tbl", sheet: "S", ref: "B2:C5", columns: ["Item", "Amt"] }],
      }),
    );
    expect(prec(g, "S!A1")).toEqual(["tbl"]);
    expect(g.node("in:S!B3:C5")!.label).toBe("tbl");
    expect(prec(g, "name:Tot")).toEqual(["tbl"]);
    expect(g.flagged()).toEqual([]);
  });

  it("a data table reads its input cells and the formulas in the row above and the column to its left", () => {
    const wb = book({ sheets: { S: { A1: "1", B2: "A1*10", C2: "A1*20", B3: { text: "", kind: "data-table" } } } });
    // One-variable table B3:C5 with row input A1: formulas B2:C2 above, values A3:A5 to the left.
    const f = wb.sheets[0]!.formulas.find((x) => x.cell === "B3")!;
    f.range = "B3:C5";
    f.attributes = { t: "dataTable", ref: "B3:C5", dt2D: "0", dtr: "0", r1: "A1" };
    const g = buildGraph(wb);
    expect(prec(g, "S!B3")).toEqual(["S!A1", "S!A2:C2", "S!A3:A5", "S!B2", "S!C2"]);
    const order = g.order().map((n) => n.label);
    expect(order.indexOf("S!B3")).toBeGreaterThan(order.indexOf("S!C2"));
  });

  it("a relative reference in a name is flagged, not resolved from A1", () => {
    const g = buildGraph(
      book({
        sheets: { S: { A1: "Left+Fixed" } },
        names: [
          { name: "Left", def: "S!B1" },
          { name: "Fixed", def: "S!$B$1+SUM(S!$C:$C)+SUM(S!$2:$2)" },
          { name: "Mixed", def: "S!$B1" },
        ],
      }),
    );
    expect(g.node("name:Left")!.flags).toMatchObject([{ kind: "dynamic", text: "S!B1" }]);
    expect(prec(g, "name:Left")).toEqual([]);
    expect(g.node("name:Fixed")!.flags).toEqual([]);
    expect(prec(g, "name:Fixed")).toEqual(["S!2:2", "S!B1", "S!C:C"]);
    expect(g.node("name:Mixed")!.flags).toMatchObject([{ kind: "dynamic" }]);
  });

  it("LET variables shadow names and make no edge", () => {
    const g = buildGraph(book({ sheets: { S: { A1: "_xlfn.LET(_xlpm.Rate, 2, _xlpm.Rate*B1)" } }, names: [{ name: "Rate", def: "S!$B$9" }] }));
    expect(prec(g, "S!A1")).toEqual(["S!B1"]);
  });

  it("structured references: columns, specials, this row", () => {
    const g = buildGraph(
      book({
        sheets: { S: { A1: "SUM(tbl[Amt])", A2: "ROWS(tbl[#All])", A3: "SUM(tbl[[#Headers],[Item]:[Amt]])", D3: "tbl[[#This Row],[Amt]]*2", E3: "[@Amt]+[@[Item]]", A4: "tbl[Nope]" } },
        tables: [{ name: "tbl", sheet: "S", ref: "B2:D5", columns: ["Item", "Amt", "Twice"] }],
      }),
    );
    expect(prec(g, "S!A1")).toEqual(["tbl[Amt]"]);
    expect(g.node("in:S!C3:C5")!.label).toBe("tbl[Amt]");
    expect(prec(g, "S!A2")).toEqual(["S!D3", "tbl[#All]"]);
    expect(prec(g, "S!A3")).toEqual(["tbl[[#Headers],[Item]:[Amt]]"]);
    expect(prec(g, "S!D3")).toEqual(["S!C3"]);
    // E3 is outside the Table (B2:D5): `[@Amt]` has no Table there.
    expect(g.node("S!E3")!.flags[0]).toMatchObject({ kind: "broken", reason: "table reference outside a Table" });
    expect(g.node("S!A4")!.flags[0]).toMatchObject({ kind: "broken", reason: "Table tbl has no column 'Nope'" });
    expect(parseStructInner("@[Item]")).toEqual({ areas: ["#this row"], columns: ["Item"] });
    expect(parseStructInner("[#Data],[A']]:[B]")).toEqual({ areas: ["#data"], columns: ["A]", "B"] });
  });

  it("3-D and cross-sheet references", () => {
    const g = buildGraph(book({ sheets: { S1: { A1: "1" }, S2: { A1: "2" }, S3: { A1: "SUM(S1:S2!A1)", A2: "'S2'!A1 + 'S1'!B1" } } }));
    expect(prec(g, "S3!A1")).toEqual(["'S1'!A1", "'S2'!A1"]);
    expect(prec(g, "S3!A2")).toEqual(["'S1'!B1", "'S2'!A1"]);
  });

  it("flags what it cannot resolve: INDIRECT, computed OFFSET, other workbooks, #REF!", () => {
    const g = buildGraph(
      book({ sheets: { S: { A1: 'INDIRECT("B"&B1)', A2: "OFFSET(A1,B1,0)", A3: "OFFSET(B1,1,1,2,1)", A4: "[1]Other!A1", A5: "#REF!+1", A6: "S!#REF!", A7: "Nope*2" } } }),
    );
    expect(g.node("S!A1")!.flags).toMatchObject([{ kind: "dynamic" }]);
    expect(prec(g, "S!A1")).toEqual(["S!B1"]);
    // OFFSET's first argument is a position: no edge to A1 (that would be a false cycle).
    expect(g.node("S!A2")!.flags).toMatchObject([{ kind: "dynamic", reason: "OFFSET with computed arguments" }]);
    expect(prec(g, "S!A2")).toEqual(["S!B1"]);
    expect(prec(g, "S!A3")).toEqual(["S!C2:C3"]);
    expect(g.node("S!A4")!.flags).toMatchObject([{ kind: "external" }]);
    expect(g.node("S!A5")!.flags).toMatchObject([{ kind: "broken" }]);
    expect(g.node("S!A6")!.flags).toMatchObject([{ kind: "broken" }]);
    expect(g.node("S!A7")!.flags).toMatchObject([{ kind: "broken", reason: "unknown name 'Nope' (#NAME?)" }]);
    expect(g.stats()).toMatchObject({ dynamic: 2, external: 1, broken: 3 });
  });
});

describe("buildGraph: cycles, order, levels", () => {
  it("finds a self-loop and a two-sheet cycle; LAMBDA recursion is not a cycle", () => {
    const g = buildGraph(
      book({
        sheets: { S: { A1: "A1+1", B1: "T!B1+1", C1: "MyFact(3)" }, T: { B1: "S!B1*2" } },
        names: [{ name: "MyFact", def: "_xlfn.LAMBDA(_xlpm.n, IF(_xlpm.n<2, 1, _xlpm.n*MyFact(_xlpm.n-1)))" }],
      }),
    );
    expect(g.cycles.map((c) => c.members.map((m) => m.label))).toEqual([["S!A1"], ["S!B1", "T!B1"]]);
    expect(g.node("T!B1")!.cycle).toBe(2);
    expect(g.recursions.map((c) => c.members.map((m) => m.label))).toEqual([["MyFact"]]);
    expect(g.nameCycles()).toMatchObject([{ recursive: true }]);
    const order = g.order().filter((n) => n.kind === "formula").map((n) => n.label);
    expect(order.indexOf("T!B1")).toBe(order.indexOf("S!B1") + 1); // kept together
  });

  it("keeps a well-ordered sheet in order of appearance, and moves only what must move", () => {
    const wb = book({ sheets: { S: { A1: "1", A2: "A1+1", A3: "A5*2", A4: "A2+1", A5: "A1*3" } } });
    expect(sheetCalcView(wb, "S").map((l) => l.cell)).toEqual(["A1", "A2", "A4", "A5", "A3"]);
    const ordered = book({ sheets: { S: { A1: "1", B1: "A1", A2: "B1", B2: "A2+A1" } } });
    expect(sheetCalcView(ordered, "S").map((l) => l.cell)).toEqual(["A1", "B1", "A2", "B2"]);
  });

  it("a sheet's order is not held back by other sheets", () => {
    // S!A1 needs T!A1 (ready at once); S!A2 is ready at once too: A1 still comes first.
    const wb = book({ sheets: { S: { A1: "T!A1", A2: "1" }, T: { A1: "2" } } });
    expect(sheetCalcView(wb, "S").map((l) => l.cell)).toEqual(["A1", "A2"]);
    expect(workbookFormulaView(wb).map((l) => `${l.sheet}!${l.cell}`)).toEqual(["S!A2", "T!A1", "S!A1"]);
  });

  it("levels: inputs 0, formulas one more than what they read, names transparent", () => {
    const g = buildGraph(book({ sheets: { S: { A1: "B1", A2: "A1+X", A3: "A2" } }, names: [{ name: "X", def: "S!$A$1" }] }));
    expect(["in:S!B1", "S!A1", "name:X", "S!A2", "S!A3"].map((k) => g.node(k)!.level)).toEqual([0, 1, 1, 2, 3]);
  });

  it("every edge outside a cycle points backwards in the order", () => {
    const g = buildGraph(book({ sheets: { S: { A1: "B1", A2: "A3", A3: "A1", B5: "A2+A3" } } }));
    const at = new Map(g.order().map((n, k) => [n.id, k]));
    for (const n of g.nodes) for (const p of g.precedents(n)) expect(at.get(p.id)!).toBeLessThan(at.get(n.id)!);
  });
});

describe("buildGraph: audit helpers", () => {
  it("unused names, and names used only by unused names (C10)", () => {
    const g = buildGraph(
      book({
        sheets: { S: { A1: "Used*2" } },
        names: [
          { name: "Used", def: "S!$B$1" },
          { name: "Dead", def: "Helper+1" },
          { name: "Helper", def: "S!$B$2" },
          { name: "Lonely", def: "3" },
        ],
      }),
    );
    const u = g.unusedNames();
    expect(u.unused.map((n) => n.label)).toEqual(["Dead", "Lonely"]);
    expect(u.onlyByUnused.map((n) => n.label)).toEqual(["Helper"]);
  });

  it("names used by conditional formats are used", () => {
    const wb = book({ sheets: { S: { A1: "1" } }, names: [{ name: "Limit", def: "S!$B$1" }] });
    wb.sheets[0]!.conditionalFormats.push({ sqref: "A1", type: "expression", priority: 1, formulas: ["A1>Limit"], ext: false });
    expect(buildGraph(wb).unusedNames().unused).toEqual([]);
  });

  it("name-level cycles (C12)", () => {
    const g = buildGraph(book({ sheets: { S: {} }, names: [{ name: "P", def: "Q+1" }, { name: "Q", def: "P*2" }] }));
    expect(g.nameCycles().map((c) => [c.members.map((m) => m.label), c.recursive])).toEqual([[["P", "Q"], false]]);
    expect(g.cycles).toHaveLength(1);
  });

  it("precedents, dependents and inputs", () => {
    const g = buildGraph(book({ sheets: { S: { A1: "B1+K", A2: "A1*2", A3: "A2+A1" } }, names: [{ name: "K", def: "0.5" }] }));
    expect(g.dependents(g.node("S!A1")!).map((n) => n.label)).toEqual(["S!A2", "S!A3"]);
    expect(g.inputsOf(g.node("S!A3")!).map((n) => n.label)).toEqual(["K", "S!B1"]);
  });
});

describe("calculation-order view", () => {
  it("renders levels, cycles and sheets", () => {
    const wb = book({ sheets: { S: { A1: "1", A2: "T!A1+A1", A3: "A3+1" }, T: { A1: "S!A1*2" } }, names: [{ name: "One", def: "S!$A$1" }] });
    const lines = workbookFormulaView(wb);
    expect(lines.map((l) => [l.sheet, l.cell, l.level, l.cycle])).toEqual([
      ["S", "A1", 1, undefined],
      ["S", "A3", 1, 1],
      ["T", "A1", 2, undefined],
      ["S", "A2", 3, undefined],
    ]);
    expect(lines[3]!.dependsOn).toEqual(["A1", "T!A1"]);
    const r = renderFormulaView(lines, { sheet: "", workbook: "w.xlsx", order: "calculation order", levels: true, sheets: true });
    expect(r.text).toContain("// Workbook w.xlsx: 4 formulas on 2 sheets in calculation order");
    expect(r.text).toContain("// ↻1: circular reference");
    for (const e of r.entries) {
      const l = lines[e.index]!;
      expect(r.text.slice(e.address.start, e.address.end)).toBe(l.cell);
      expect(r.text.slice(e.lhs[0]?.start ?? 0, e.lhs[0]?.end ?? 0)).toBe(l.lhs[0]?.display ?? "");
    }
    expect(r.text).toMatch(/\n1 ↻1 {2}\s*S!A3 /);
  });
});
