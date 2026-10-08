// Stretch G (M5), the file backend: a change set's renames with `references` rewrite the
// token in cell formulas, conditional formats, validations and other names, and the
// result matches Excel's own rename (probe F7's oracle) formula for formula.
import { describe, expect, it } from "vitest";
import { applyChangeSet, applyChangeSetWithReport, readBack, readWorkbook, type Change, type WorkbookSnapshot } from "../../src/index.js";
import { Package } from "../../src/file/package.js";
import { fixture } from "./helpers.js";

const BASE = fixture("f7_base.xlsx");
const ORACLE = fixture("f7_oracle.xlsx");
const REFS = { cells: 0, formats: 0, validations: 0, names: 0 };
const F7_RENAMES: Change[] = [
  { op: "rename-name", scope: null, from: "Rate", to: "Growth", references: REFS },
  { op: "rename-name", scope: null, from: "Fn", to: "Fx", references: REFS },
  { op: "rename-name", scope: "S2", from: "Loc", to: "Spot", references: REFS },
];

/** Every formula of a workbook by place: names by scope and name, cells, CF and DV in order. */
function formulas(wb: WorkbookSnapshot): Map<string, string> {
  const out = new Map<string, string>();
  for (const d of wb.definedNames) if (!d.isXlPrefixed) out.set(`name ${d.scope.kind === "sheet" ? d.scope.name + "!" : ""}${d.name}`, d.definition);
  for (const s of wb.sheets) {
    for (const f of s.formulas) out.set(`cell ${s.name}!${f.cell}`, f.text ?? "(shared child)");
    s.conditionalFormats.forEach((c, i) => c.formulas.forEach((x, j) => out.set(`cf ${s.name}#${i}.${j}`, x)));
    s.dataValidations.forEach((v, i) => out.set(`dv ${s.name}#${i}`, `${v.formula1 ?? ""}|${v.formula2 ?? ""}`));
  }
  return out;
}

describe("rename across cells: the file backend (F7)", () => {
  it("matches Excel's own rename on every formula of f7_base.xlsx", () => {
    const { bytes, report } = applyChangeSetWithReport(BASE, F7_RENAMES);
    const got = formulas(readWorkbook(bytes));
    const want = formulas(readWorkbook(ORACLE));
    const diffs = [...new Set([...got.keys(), ...want.keys()])].filter((k) => got.get(k) !== want.get(k)).map((k) => `${k}: got ${got.get(k)}, Excel ${want.get(k)}`);
    expect(diffs).toEqual([]);
    expect(report.references).toMatchObject({ counts: { cells: 14, formats: 1, validations: 1, names: 2 }, sheets: ["S1", "S2"] });
    expect(report.references!.calcChainDropped).toBe("xl/calcChain.xml");
  });

  it("reads back: the references step on its own, then the renames", () => {
    const after = applyChangeSet(BASE, F7_RENAMES);
    const rb = readBack(BASE, after, F7_RENAMES, []);
    expect(rb.problems).toEqual([]);
    // fullCalcOnLoad and no calcChain (F8: a stale chain makes Excel repair the file).
    const pkg = new Package(after);
    expect(pkg.has("xl/calcChain.xml")).toBe(false);
    expect(pkg.text("xl/workbook.xml")).toContain('fullCalcOnLoad="1"');
    expect(pkg.text("xl/_rels/workbook.xml.rels")).not.toContain("calcChain");
    expect(pkg.text("[Content_Types].xml")).not.toContain("calcChain");
  });

  it("read-back catches a reference left behind or a formula changed beyond the token", () => {
    const after = applyChangeSet(BASE, F7_RENAMES);
    // Pretend the change set renamed only Rate: Fx and Spot then appear and Fn/Loc readers changed unasked.
    const one: Change[] = [F7_RENAMES[0]!];
    expect(readBack(BASE, after, one, []).ok).toBe(false);
  });

  it("without `references` a rename touches only the name (D4 as before)", () => {
    const plain: Change[] = [{ op: "rename-name", scope: null, from: "RateX", to: "RateY" }];
    const after = applyChangeSet(BASE, plain);
    const wb = readWorkbook(after);
    expect(wb.sheets[0]!.formulas.find((f) => f.cell === "C1")!.text).toBe("RateX+Rate");
    expect(new Package(after).has("xl/calcChain.xml")).toBe(true);
  });

  it("a rename whose new name a reader would capture is not written", () => {
    const fine: Change[] = [{ op: "rename-name", scope: null, from: "Rate", to: "Pace", references: REFS }, { op: "rename-name", scope: null, from: "RateX", to: "x", references: REFS }];
    expect(() => applyChangeSet(BASE, fine)).not.toThrow(); // no LAMBDA with a variable x reads RateX
    // Fn = LAMBDA(x, x*Rate): renaming Rate to x would make it read its own variable.
    const capture: Change[] = [{ op: "rename-name", scope: null, from: "Rate", to: "x", references: REFS }];
    expect(() => applyChangeSet(BASE, capture)).toThrow(/name Fn: would read another name \(x\)/);
    // Rate → Loc: on S2 the local Loc would capture Rate's readers (S2!A2 = Rate+Loc).
    const local: Change[] = [{ op: "rename-name", scope: null, from: "Rate", to: "Loc", references: REFS }];
    expect(() => applyChangeSet(BASE, local)).toThrow(/cell S2!A2: would read another name/);
  });
});
