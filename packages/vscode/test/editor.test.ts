// M3c: completion, signature help and the checks as you type, on a hand-made project.
import { catalogue } from "@xln/core";
import { afterEach, describe, expect, it } from "vitest";
import { completions, formulaSite, liveProblems, modulePrefixProblem, newModuleText, oneEditAway, sheetPrefix, signatureHelp, type Completion } from "../src/model/editor.js";
import { Project } from "../src/model/project.js";
import { hoverMarkdown } from "../src/model/usages.js";

const IS = "names/sheets/IS.xln";
const BS = "names/sheets/BS.xln";
const SCF = "names/sheets/SCF recursive.xln";
const FN = "names/FN.xln";
const U = "names/_unmanaged.xln";

const FILES: Record<string, string> = {
  [IS]: [
    "@scope(IS)",
    "Sales @C6# = SEQUENCE(1, 5) * 100;",
    "EBIT @C9# = Sales * 0.2;",
    "Taxes @C10# = EBIT * 0.3;",
    "Unlevered_net_income @C11# = EbIT - Taxes;",
    "",
  ].join("\n"),
  [BS]: ["@scope(BS)", "Cash @C5# = Sales * 0.1;", "Debt @C6# = IS!Sales + Csh;", "@C7 = Cash + Debt;", ""].join("\n"),
  [SCF]: ["@scope('SCF recursive')", "ExcessCash @C5# = 0;", ""].join("\n"),
  [FN]: [
    "/**",
    " * Growth of a row.",
    " * @param row a row of values",
    " * @param [base] the starting value",
    " */",
    "FN.GROWTH = LAMBDA(row, [base], row / base - 1);",
    "FN.TWICE = LAMBDA(x, 2 * x);",
    "",
  ].join("\n"),
  [U]: ["Rate = 0.1;", "Pct = 0.5;", "Check = FN.TWICE(1, 2) + SUMM(1) + ROUND(1) + Pct(2) + BYROW(Rate, SUM);", "Typo = Rat + FN.TWICE(Rate);", ""].join("\n"),
};

function project(files: Record<string, string> = FILES): Project {
  const p = new Project("mem:/p", undefined);
  for (const [path, text] of Object.entries(files)) p.setFile(path, text);
  return p;
}

/** The project with `|` written into `path` (removed), and the cursor's offset. */
function cursor(path: string, text: string): { p: Project; offset: number } {
  const offset = text.indexOf("|");
  const p = project({ ...FILES, [path]: text.slice(0, offset) + text.slice(offset + 1) });
  return { p, offset };
}

const labels = (items: Completion[] | undefined) => (items ?? []).map((i) => i.label);

describe("completion", () => {
  it("lists the home sheet's names bare, other sheets' qualified, workbook names, modules and functions, in that order", () => {
    const { p, offset } = cursor(BS, "@scope(BS)\nCash @C5# = |Sales * 0.1;\nDebt @C6# = 1;\n");
    const items = completions(p, BS, offset)!;
    const l = labels(items);
    expect(l).toEqual(expect.arrayContaining(["Cash", "Debt", "IS!Sales", "IS!EBIT", "'SCF recursive'!ExcessCash", "Rate", "Check", "FN.", "SUM", "XLOOKUP"]));
    expect(l).not.toContain("Sales");
    expect(l).not.toContain("FN.GROWTH"); // behind its module prefix
    expect(l).not.toContain("ANCHORARRAY"); // written by Excel, not typed
    const sorted = [...items].sort((a, b) => a.sortText.localeCompare(b.sortText)).map((i) => i.label);
    expect(sorted.indexOf("Cash")).toBeLessThan(sorted.indexOf("IS!Sales"));
    expect(sorted.indexOf("IS!Sales")).toBeLessThan(sorted.indexOf("SUM"));
    const q = items.find((i) => i.label === "IS!Sales")!;
    expect(q).toMatchObject({ insertText: "IS!Sales", filterText: "Sales", kind: "name", nameKind: "range" });
    expect(items.find((i) => i.label === "FN.")).toMatchObject({ kind: "module", retrigger: true, detail: "module · 2 names" });
    expect(items.find((i) => i.label === "XLOOKUP")!.detail).toMatch(/^XLOOKUP\(.*\) · Excel 2021$/);
    // The range is the word being typed.
    const typed = cursor(BS, "@scope(BS)\nCash @C5# = Sal|;\n");
    const it0 = completions(typed.p, BS, typed.offset)!.find((i) => i.label === "IS!Sales")!;
    expect(typed.p.files.get(BS)!.text.slice(it0.start, it0.end)).toBe("Sal");
  });

  it("puts LET and LAMBDA variables first", () => {
    const { p, offset } = cursor(U, "X = LET(rate, 0.1, f, LAMBDA(v, v * r|));\n");
    const items = completions(p, U, offset)!;
    const vars = items.filter((i) => i.kind === "variable");
    expect(vars.map((v) => [v.label, v.detail])).toEqual([
      ["v", "LAMBDA parameter"],
      ["rate", "LET variable"],
    ]);
    expect(vars.every((v) => v.sortText.startsWith("0_"))).toBe(true);
    // A workbook formula has no home sheet: every local name is qualified.
    expect(labels(items)).toContain("BS!Cash");
  });

  it("after a module prefix, its members; after Sheet!, that sheet's local names", () => {
    let c = cursor(U, "X = FN.|;\n");
    let items = completions(c.p, U, c.offset)!;
    expect(labels(items).sort()).toEqual(["FN.GROWTH", "FN.TWICE"]);
    expect(c.p.files.get(U)!.text.slice(items[0]!.start, items[0]!.end)).toBe("FN.");
    c = cursor(U, "X = NORM.|;\n");
    expect(labels(completions(c.p, U, c.offset))).toContain("NORM.DIST");
    c = cursor(BS, "@scope(BS)\nCash @C5# = IS!|;\n");
    expect(labels(completions(c.p, BS, c.offset)).sort()).toEqual(["EBIT", "Sales", "Taxes", "Unlevered_net_income"]);
    c = cursor(BS, "@scope(BS)\nCash @C5# = 'SCF recursive'!Ex|;\n");
    expect(labels(completions(c.p, BS, c.offset))).toEqual(["ExcessCash"]);
  });

  it("offers nothing in strings, comments, on the left-hand side or in a cell address", () => {
    for (const t of ['X = IF(Rate > 0, "Sa|", 1);', "X = 1 + /* Sa| */ 2;", "X = 1; // Sa|", "Ra|te = 0.1;", "Cash @C|5# = 1;"]) {
      const { p, offset } = cursor(U, t + "\n");
      expect(completions(p, U, offset), t).toBeUndefined();
    }
    // After `=`, also with nothing typed yet (a slot being filled).
    const s = cursor(BS, "@scope(BS)\nCash @C5# = |;\n");
    expect(labels(completions(s.p, BS, s.offset))).toContain("Cash");
    const blank = cursor(BS, "@scope(BS)\nCash @C5# =  |  ;\n");
    expect(labels(completions(blank.p, BS, blank.offset))).toContain("Cash");
  });

  it("quotes a sheet only when a formula needs it", () => {
    expect(sheetPrefix("IS")).toBe("IS!");
    expect(sheetPrefix("SCF recursive")).toBe("'SCF recursive'!");
    expect(sheetPrefix("A1")).toBe("'A1'!");
    expect(sheetPrefix("It's")).toBe("'It''s'!");
  });

  it("knows the formula index around blanks, and not on the left-hand side", () => {
    const p = project();
    const text = p.files.get(BS)!.text;
    expect(formulaSite(p, BS, text.indexOf("Sales * 0.1"))!.index).toBe(0);
    expect(formulaSite(p, BS, text.indexOf("Cash @"))).toBeUndefined();
    expect(formulaSite(p, BS, text.indexOf("@C5#") + 2)).toBeUndefined();
  });
});

describe("signature help", () => {
  const saved = catalogue().get("ROUND")!.params;
  afterEach(() => {
    const info = catalogue().get("ROUND")!;
    if (saved === undefined) delete info.params;
    else info.params = saved;
  });

  it("shows a project LAMBDA's parameters and its doc comment's @param lines", () => {
    let c = cursor(U, "X = FN.GROWTH(|);\n");
    let s = signatureHelp(c.p, U, c.offset)!;
    expect(s.label).toBe("FN.GROWTH(row, [base])");
    expect(s.documentation).toBe("Growth of a row.");
    expect(s.params).toEqual([
      { label: "row", documentation: "a row of values" },
      { label: "[base]", documentation: "the starting value" },
    ]);
    expect(s.active).toBe(0);
    c = cursor(U, "X = FN.GROWTH(SUM(1, 2), |);\n");
    s = signatureHelp(c.p, U, c.offset)!;
    expect(s.active).toBe(1);
    // Inside a nested call, the inner one.
    c = cursor(U, "X = FN.GROWTH(SUM(1, |), 2);\n");
    expect(signatureHelp(c.p, U, c.offset)!.label).toMatch(/^SUM\(/);
  });

  it("falls back to the arity for a built-in without parameter names, and uses them when the catalogue has them", () => {
    const info = catalogue().get("ROUND")!;
    delete info.params;
    let c = cursor(U, "X = ROUND(Rate, |);\n");
    expect(signatureHelp(c.p, U, c.offset)).toMatchObject({ label: "ROUND(2 args)", params: [] });
    info.params = ["number", "num_digits"];
    const s = signatureHelp(c.p, U, c.offset)!;
    expect(s).toMatchObject({ label: "ROUND(number, num_digits)", active: 1 });
    // A repeating tail: every further argument is the last parameter.
    info.params = ["number1", "[number2]", "..."];
    c = cursor(U, "X = ROUND(1, 2, 3, |);\n");
    expect(signatureHelp(c.p, U, c.offset)!.active).toBe(1);
  });

  it("uses the catalogue's parameter names, repeating a group of two", () => {
    let c = cursor(U, "X = XLOOKUP(Rate, Rate, Rate, |);\n");
    expect(signatureHelp(c.p, U, c.offset)).toMatchObject({
      label: "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])",
      active: 3,
    });
    c = cursor(U, "X = SUMIFS(Rate, Rate, 1, Rate, 2, |);\n");
    let s = signatureHelp(c.p, U, c.offset)!;
    expect(s.params[s.active]!.label).toBe("[criteria_range2]");
    c = cursor(U, "X = SUMIFS(Rate, Rate, 1, Rate, 2, Rate, |);\n");
    s = signatureHelp(c.p, U, c.offset)!;
    expect(s.params[s.active]!.label).toBe("[criteria2]");
  });

  it("knows a LAMBDA bound in a LET, and nothing outside a call", () => {
    let c = cursor(U, "X = LET(sq, LAMBDA(x, [y], x * x), sq(1, |));\n");
    expect(signatureHelp(c.p, U, c.offset)).toMatchObject({ label: "sq(x, [y])", active: 1 });
    c = cursor(U, "X = Rate + |;\n");
    expect(signatureHelp(c.p, U, c.offset)).toBeUndefined();
    c = cursor(U, 'X = SUM("a, |");\n');
    expect(signatureHelp(c.p, U, c.offset)).toBeUndefined();
  });
});

describe("checks as you type", () => {
  const p = project();
  const show = (path: string) =>
    liveProblems(p, path).map((x) => `${x.severity} ${x.code} ${p.files.get(path)!.text.slice(x.start, x.end)}: ${x.message}`);

  it("an unqualified read of another sheet's local name, with a fix that qualifies it", () => {
    const probs = liveProblems(p, BS);
    const c5 = probs.find((x) => x.code === "C5.other-sheet")!;
    expect(c5.message).toBe("Sales is local to IS, not to BS: unqualified it is #NAME? in Excel; write IS!Sales");
    expect(p.files.get(BS)!.text.slice(c5.start, c5.end)).toBe("Sales");
    expect(c5.fixes).toEqual([{ title: "Qualify: IS!Sales", start: c5.start, end: c5.end, text: "IS!Sales" }]);
  });

  it("an unknown name with a did-you-mean fix; a different case is a hint only", () => {
    expect(show(BS)).toContain("error C4.unknown-name Csh: Csh is not a name in scope, a LET/LAMBDA variable or a function: #NAME? in Excel (did you mean Cash?)");
    expect(liveProblems(p, BS).find((x) => x.code === "C4.unknown-name")!.fixes!.map((f) => f.text)).toEqual(["Cash"]);
    const hint = liveProblems(p, IS).find((x) => x.code === "spelling")!;
    expect(hint).toMatchObject({ severity: "hint", message: "EbIT is defined as EBIT: Excel reads it case-insensitively and shows it as EBIT" });
    expect(hint.fixes![0]).toMatchObject({ title: "Match the name's spelling: EBIT", text: "EBIT" });
    // The author's case raises nothing else.
    expect(liveProblems(p, IS).filter((x) => x.severity !== "hint")).toEqual([]);
    expect(show(U)).toContain("error C4.unknown-name Rat: Rat is not a name in scope, a LET/LAMBDA variable or a function: #NAME? in Excel (did you mean Rate?)");
  });

  it("unknown functions, argument counts of LAMBDAs and built-ins, a value called like a function", () => {
    expect(show(U)).toEqual([
      "error C6.lambda-arity FN.TWICE: FN.TWICE(x) takes 1 argument; it is given 2",
      "error unknown-function SUMM: SUMM(…): no built-in function and no LAMBDA of that name; Excel would store it as _xludf.SUMM (#NAME?): did you mean SUM?",
      "warning C6.builtin-arity ROUND: ROUND takes 2 arguments; it is given 1",
      "error C6.not-a-function Pct: Pct is a constant, not a LAMBDA, but is called with 1 argument",
      "error C4.unknown-name Rat: Rat is not a name in scope, a LET/LAMBDA variable or a function: #NAME? in Excel (did you mean Rate?)",
    ]);
    expect(liveProblems(p, U).find((x) => x.code === "unknown-function")!.fixes!.map((f) => f.text)).toContain("SUM");
  });

  it("a cell called like a function, and a name on a cell called: the LAMBDA the cell holds", () => {
    const ML = "names/sheets/MyLambda.xln";
    const files = {
      [FN]: "a_plus_b = LAMBDA(a, b, a + b);\n",
      [ML]: [
        "@workbook",
        "a_plus_b_fake @C2 = LAMBDA(a,b,a+b);",
        "Num @C3 = 1+2;",
        "@C4 = LET(k, 2, LAMBDA(x, x * k));",
        "@F2 = C2(D2,E2) + $C$2(1, 2) + MyLambda!C2(1, 2) + C4(1);",
        "@F3 = $C$2(1);",
        "@F4 = C3(1,2);",
        "@F5 = B9(1);",
        "@G2 = a_plus_b_fake(D2,E2) + a_plus_b(D2,E2) + LOG10(100);",
        "@G3 = a_plus_b_fake(1,2,3);",
        "@G4 = Num(1);",
        "",
      ].join("\n"),
    };
    const q = project(files);
    const shown = () => liveProblems(q, ML).map((x) => `${x.severity} ${x.code} ${q.files.get(ML)!.text.slice(x.start, x.end)}: ${x.message}`);
    expect(shown()).toEqual([
      "error C6.lambda-arity $C$2: $C$2(a, b) takes 2 arguments; it is given 1",
      "warning C6.not-a-lambda C3: C3 is called with 2 arguments, but MyLambda!C3 has a formula that gives a value, not a LAMBDA: Excel gives #VALUE! or #CALC!",
      "error C6.lambda-arity a_plus_b_fake: a_plus_b_fake(a, b) takes 2 arguments; it is given 3",
      "warning C6.not-a-lambda Num: Num is called with 1 argument, but MyLambda!C3 has a formula that gives a value, not a LAMBDA: Excel gives #VALUE! or #CALC!",
    ]);
    // B9 has no statement: without the workbook's cells nothing is known of it; with the
    // last pull's, it holds a value or nothing.
    q.manifest = { format: "xln.manifest/1", workbook: "w.xlsx", sheets: [{ name: "MyLambda", position: 0 }], tables: [], names: {} };
    q.lock = { format: "xln.lock/2", workbook: "w.xlsx", names: {}, cells: {} };
    expect(shown()).toContain("warning C6.not-a-lambda B9: B9 is called with 1 argument, but MyLambda!B9 holds a value or nothing, not a LAMBDA: Excel gives #VALUE! or #CALC!");
  });

  it("resolves LET/LAMBDA variables, qualified names, functions passed as values and Tables", () => {
    const q = project({
      ...FILES,
      [U]: "A = LET(x, 1, f, LAMBDA(y, x + y), f(2)) + IS!Sales + BYROW(BS!Cash, SUM) + SUM(tblIn);\nB = NoSheet!Sales + IS!Nope;\n",
    });
    q.manifest = { format: "xln.manifest/1", workbook: "w.xlsx", sheets: [{ name: "IS", position: 0 }, { name: "BS", position: 1 }], tables: [{ name: "tblIn", sheet: "IS", ref: "A1:B3", columns: ["a", "b"] }], names: {} };
    const msgs = liveProblems(q, U).map((x) => x.message);
    expect(msgs).toEqual(["NoSheet!Sales: the workbook has no sheet NoSheet", "IS!Nope: no name Nope on IS or in the workbook: #NAME? in Excel"]);
  });

  it("a doc comment over the Name Manager's 255 characters is an error (the build refuses it)", () => {
    const long = "x".repeat(260);
    const q = project({ [U]: `/** ${long} */\nLong = 1;\n` });
    const w = liveProblems(q, U);
    expect(w.map((x) => [x.severity, x.code])).toEqual([["error", "comment-length"]]);
    expect(q.files.get(U)!.text.slice(w[0]!.start, w[0]!.end)).toBe(`/** ${long} */`);
  });

  it("near misses: one edit or a swap, not case", () => {
    expect(oneEditAway("Csh", "Cash")).toBe(true);
    expect(oneEditAway("Cahs", "Cash")).toBe(true);
    expect(oneEditAway("Cbsh", "Cash")).toBe(true);
    expect(oneEditAway("CASH", "Cash")).toBe(false);
    expect(oneEditAway("Ch", "Cash")).toBe(false);
  });
});

describe("LAMBDA authoring", () => {
  it("shows @param lines in the hover", () => {
    const p = project();
    const md = hoverMarkdown(p, p.lookup("FN.GROWTH")!, undefined);
    expect(md).toContain("Growth of a row.");
    expect(md).toContain("*@param* `row` — a row of values");
    expect(md).toContain("*@param* `base` — the starting value");
  });

  it("New module: checks the prefix and writes a module with a documented sample LAMBDA", () => {
    const p = project();
    expect(modulePrefixProblem(p, "FN")).toBe("names/FN.xln exists already");
    expect(modulePrefixProblem(p, "fin")).toBeUndefined();
    expect(modulePrefixProblem(p, "A1")).toMatch(/reads as a cell/);
    expect(modulePrefixProblem(p, "F.N")).toMatch(/no dot/);
    expect(modulePrefixProblem(p, "")).toMatch(/type a prefix/);
    const text = newModuleText("FIN");
    const q = project({ ...FILES, "names/FIN.xln": text });
    const d = q.lookup("FIN.GROW")!;
    expect(q.problems("names/FIN.xln")).toEqual([]);
    expect(liveProblems(q, "names/FIN.xln")).toEqual([]);
    expect(q.analysis(d).classification).toMatchObject({ kind: "lambda", params: ["value", "rate", "[periods]"] });
    expect(d.entry.doc).toContain("@param rate the growth rate per period");
    const c = cursor("names/FIN.xln", text + "X = FIN.GROW(1, |);\n");
    const p2 = c.p;
    p2.setFile("names/FIN.xln", text + "X = FIN.GROW(1, );\n");
    const s = signatureHelp(p2, "names/FIN.xln", c.offset)!;
    expect(s.params.map((x) => x.documentation)).toEqual(["the starting value", "the growth rate per period (0.05 for 5%)", "how many periods; 1 when omitted"]);
    expect(modulePrefixProblem(q, "FIN")).toBe("names/FIN.xln exists already");
  });
});
