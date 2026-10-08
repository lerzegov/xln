import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { formulaTextAt, readWorkbook } from "../../src/file/index.js";
import { shiftAddress, shiftFormula } from "../../src/lang/index.js";

const RESULTS = join(import.meta.dirname, "..", "..", "..", "..", "probes", "results");

describe("shiftFormula", () => {
  const cases: [string, number, number, string][] = [
    ["A2*Rate", 1, 0, "A3*Rate"],
    ["$A$1+A$1+$A1+A1", 1, 1, "$A$1+B$1+$A2+B2"],
    ["a1+b$2", 2, 2, "C3+D$2"],
    ["SUM(A1:B2)", 2, 1, "SUM(B3:C4)"],
    ["SUM($A$1:B2)", 2, 1, "SUM($A$1:C4)"],
    ["SUM(A:A)", 5, 1, "SUM(B:B)"],
    ["SUM($A:B)", 7, 1, "SUM($A:C)"],
    ["SUM(1:3)", 2, 5, "SUM(3:5)"],
    ["SUM($1:2)", 2, 0, "SUM($1:4)"],
    ["A1#", 1, 0, "A2#"],
    ["ROWS(_xlfn.ANCHORARRAY(E1))", 0, 1, "ROWS(_xlfn.ANCHORARRAY(F1))"],
    ["SUM(A1:.B5)", 1, 0, "SUM(A2:.B6)"],
    ["SUM(A1.:.B5)", 1, 0, "SUM(A2.:.B6)"],
    ["SUM(A1:INDEX(B:B, 3))", 1, 1, "SUM(B2:INDEX(C:C, 3))"],
    ["A1 B1:C3", 1, 0, "A2 B2:C4"],
    // Strings, structured references, names and functions stay.
    ['"A1"&A1', 1, 0, '"A1"&A2'],
    ["Tbl[Col1]+[@A1]+Tbl[[#This Row],[B2]]", 3, 3, "Tbl[Col1]+[@A1]+Tbl[[#This Row],[B2]]"],
    ["LOG10(A1)+Rate1+FN.PREV(B1)", 1, 0, "LOG10(A2)+Rate1+FN.PREV(B2)"],
    ["_xlfn.LET(_xlpm.x, A1, _xlpm.x+B1)", 1, 0, "_xlfn.LET(_xlpm.x, A2, _xlpm.x+B2)"],
    // Sheet-qualified names stay; references on other sheets move, quoting kept.
    ["S2!Loc+'My S'!A1+'S2'!$B$1", 1, 0, "S2!Loc+'My S'!A2+'S2'!$B$1"],
    ["'S1:S3'!A1+S1:S3!B1", 1, 0, "'S1:S3'!A2+S1:S3!B2"],
    // Excel fills references into other workbooks the same way.
    ["[1]Sheet1!A1+[1]!Name", 1, 0, "[1]Sheet1!A2+[1]!Name"],
    // Off the sheet: #REF!, keeping the sheet; a spill of it too.
    ["A1+1", -1, 0, "#REF!+1"],
    ["S2!A1", 0, -1, "S2!#REF!"],
    ["SUM(A1#)", -1, 0, "SUM(#REF!)"],
    ["SUM(A1:B2)", 0, -1, "SUM(#REF!)"],
    ["XFD1", 0, 1, "#REF!"],
    ["A1048576", 1, 0, "#REF!"],
    ["SUM(A:B)", 0, -1, "SUM(#REF!)"],
    ["SUM(1:2)", -1, 0, "SUM(#REF!)"],
    ["$A$1+A$1", -5, 0, "$A$1+A$1"],
    ["Sheet1!#REF!+A1", 1, 0, "Sheet1!#REF!+A2"],
    // Layout kept.
    ["SUM( A1 ,\r\n  B1 )", 1, 0, "SUM( A2 ,\r\n  B2 )"],
    ["=A1", 1, 0, "=A2"],
  ];
  for (const [src, dr, dc, want] of cases) {
    it(`${src} by (${dr}, ${dc})`, () => expect(shiftFormula(src, dr, dc)).toBe(want));
  }

  it("shifts text the parser rejects with the tokenizer", () => {
    expect(shiftFormula("A1+(B1", 1, 0)).toBe("A2+(B2");
  });

  it("is undone by the opposite move when nothing falls off", () => {
    for (const [src] of cases.slice(0, 22).filter(([s]) => !s.startsWith("a1"))) {
      expect(shiftFormula(shiftFormula(src, 3, 2), -3, -2)).toBe(src);
    }
  });

  it("shiftAddress moves the relative ends of every kind of address", () => {
    expect(shiftAddress("B$3:$C4", 1, 1)).toBe("C$3:$C5");
    expect(shiftAddress("A1", -1, 0)).toBeUndefined();
  });
});

describe("shared formulas in Excel-saved fixtures", () => {
  it("children of f7_base read their own row (B2:B11, A3:A5)", () => {
    const wb = readWorkbook(new Uint8Array(readFileSync(join(RESULTS, "f7_base.xlsx"))));
    const text = (sheet: string, cell: string) => {
      const s = wb.sheets.find((x) => x.name === sheet)!;
      return formulaTextAt(s, s.formulas.find((f) => f.cell === cell)!, shiftFormula);
    };
    expect(text("S1", "B2")).toBe("A2*Rate");
    expect(text("S1", "B3")).toBe("A3*Rate");
    expect(text("S1", "B11")).toBe("A11*Rate");
    expect(text("S2", "A5")).toBe("Loc*ROW()");
    // The oracle (the same workbook after Excel renamed Rate → Growth) has the same shape.
    const oracle = readWorkbook(new Uint8Array(readFileSync(join(RESULTS, "f7_oracle.xlsx"))));
    const s1 = oracle.sheets[0]!;
    expect(formulaTextAt(s1, s1.formulas.find((f) => f.cell === "B7")!, shiftFormula)).toBe("A7*Growth");
  });

  it("the saved values agree with the shifted text (B = A × Rate, A = row number)", () => {
    const wb = readWorkbook(new Uint8Array(readFileSync(join(RESULTS, "f7_base.xlsx"))));
    const s = wb.sheets[0]!;
    for (const f of s.formulas.filter((x) => x.kind.startsWith("shared"))) {
      const t = formulaTextAt(s, f, shiftFormula)!;
      const row = Number(t.slice(1, t.indexOf("*")));
      expect(f.value.value as number).toBeCloseTo((row - 1) * 0.1, 10);
    }
  });
});
