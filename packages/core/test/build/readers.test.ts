// D4's reader check counts the workbook as the build leaves it (feedback 2026-10-07): the
// author moved a cell from an old LAMBDA (ANA.GROW) to a new one (FN.GROW) and deleted
// the old one in one build, and the build refused, citing the very cell it rewrote. The
// situation is rebuilt here on the F7 probe workbook: `Fn` is a workbook LAMBDA that
// S1!C4 (`Fn(10)`) and S1!C10 (`Rate2+Fn(1)`) call.
import { describe, expect, it } from "vitest";
import { readWorkbook } from "../../src/index.js";
import { build, edit, fixture, pulled, why } from "./helpers.js";

const F7 = fixture("f7_base.xlsx");
const U = "names/_unmanaged.xln";
const S1 = "names/sheets/S1.xln";

/** The F7 source with `Fn` replaced by a new module function `FN.TIMES`, created in the same build. */
function moved(): Record<string, string> {
  const files = pulled(F7);
  edit(files, U, "Fn = LAMBDA(x, x*Rate);\n", "");
  files["names/FN.xln"] = "// module: FN\n\nFN.TIMES = LAMBDA(x, x*Rate);\n";
  return files;
}

function lineOf(text: string, find: string): number {
  return text.slice(0, text.indexOf(find)).split("\n").length;
}

describe("D4 readers after the build", () => {
  it("deletes a name whose every reader the same build moves to a name it creates", () => {
    const files = moved();
    edit(files, S1, "@C4 = Fn(10);", "@C4 = FN.TIMES(10);");
    edit(files, S1, "@C10 = Rate2+Fn(1);", "@C10 = Rate2+FN.TIMES(1);");
    const r = build(F7, files);
    expect(r.status, why(r)).toBe("built");
    expect(r.plan.changeSet.changes.map((c) => c.op)).toEqual(["delete-name", "set-name", "set-cell-formula", "set-cell-formula"]);
    const wb = readWorkbook(r.bytes!);
    expect(wb.definedNames.some((d) => d.name === "Fn")).toBe(false);
    expect(wb.definedNames.find((d) => d.name === "FN.TIMES")?.definition).toBe("_xlfn.LAMBDA(_xlpm.x, _xlpm.x*Rate)");
    expect(wb.sheets[0]!.formulas.find((f) => f.cell === "C4")?.text).toBe("FN.TIMES(10)");
  });

  it("still refuses while a reader remains, with the place in the source and advice that can be followed", () => {
    const files = moved();
    edit(files, S1, "@C4 = Fn(10);", "@C4 = FN.TIMES(10);");
    const r = build(F7, files);
    expect(r.status).toBe("refused");
    const line = lineOf(files[S1]!, "@C10 =");
    expect(r.plan.problems).toEqual([
      {
        severity: "error",
        code: "in-use",
        key: "Fn",
        sites: ["cell S1!C10"],
        file: S1,
        line,
        message: `deleting Fn refused: 1 place in the workbook refers to it by name and would break: cell S1!C10 (${S1}:${line}). Change that formula in the source, or keep the name`,
      },
    ]);
  });

  it("counts a rewritten cell by its new formula: one that still reads the deleted name refuses", () => {
    const files = moved();
    edit(files, S1, "@C4 = Fn(10);", "@C4 = Fn(20);");
    edit(files, S1, "@C10 = Rate2+Fn(1);", "@C10 = Rate2+FN.TIMES(1);");
    const r = build(F7, files);
    expect(r.status).toBe("refused");
    // The source checker sees it first (the source no longer defines Fn), at the statement.
    expect(r.plan.problems).toMatchObject([{ severity: "error", file: S1, line: lineOf(files[S1]!, "@C4 =") }]);
    expect(r.plan.problems[0]!.message).toContain("Fn");
  });

  it("a cell the build clears is no reader", () => {
    const files = moved();
    edit(files, S1, "@C4 = Fn(10);", "@C4 = ;");
    edit(files, S1, "@C10 = Rate2+Fn(1);", "@C10 = Rate2;");
    const r = build(F7, files);
    expect(r.status, why(r)).toBe("built");
    expect(r.plan.changeSet.changes.map((c) => c.op)).toContain("clear-cell-formula");
  });

  it("several readers: the message says places … refer, and lists them", () => {
    const files = moved();
    const r = build(F7, files);
    expect(r.status).toBe("refused");
    const p = r.plan.problems.find((x) => x.code === "in-use")!;
    expect(p.sites).toEqual(["cell S1!C4", "cell S1!C10"]);
    expect(p.message).toContain("2 places in the workbook refer to it by name");
    expect(p.message).toContain("Change these formulas in the source, or keep the name");
  });
});

describe("refusals say what the build would do, in natural order (spec §14 issue 11)", () => {
  it("renaming Fn to Times refused while the source's cells still read Fn (M5: the build rewrites the workbook's)", () => {
    const files = pulled(F7);
    edit(files, U, "Fn = LAMBDA(x, x*Rate);", "@renamed(Fn)\nTimes = LAMBDA(x, x*Rate);");
    const r = build(F7, files);
    expect(r.status).toBe("refused");
    const p = r.plan.problems.filter((x) => x.code === "in-use");
    expect(p.map((x) => x.line)).toEqual([15, 21]);
    for (const x of p) expect(x.message).toContain("the source renames Fn to Times, so write Times here");
  });

  it("moving Rate2 to sheet S2 refused", () => {
    const files = pulled(F7);
    edit(files, U, "Rate2 = Rate*2;\n", "");
    edit(files, "names/sheets/S2.xln", "Loc = 7;", "Loc = 7;\nRate2 = Rate*2;");
    const r = build(F7, files);
    expect(r.status).toBe("refused");
    const p = r.plan.problems.find((x) => x.code === "in-use")!;
    expect(p.message).toMatch(/^moving Rate2 to sheet S2 refused: 1 place in the workbook refers to it by name/);
  });
});
