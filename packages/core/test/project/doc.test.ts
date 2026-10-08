import { describe, expect, it } from "vitest";
import { docText, formatDoc, paramDoc, parseDocComment, parseModule } from "../../src/index.js";

describe("doc comments with @param", () => {
  it("splits the summary from the parameters, inline or on lines", () => {
    expect(parseDocComment("Adds. @param x the first @param y the second")).toEqual({
      summary: "Adds.",
      params: [
        { name: "x", text: "the first" },
        { name: "y", text: "the second" },
      ],
    });
    const multi = parseDocComment("Growth over a row.\n\n@param row a row of values\n  over several lines\n@param [base] - optional base");
    expect(multi.summary).toBe("Growth over a row.");
    expect(multi.params).toEqual([
      { name: "row", text: "a row of values over several lines" },
      { name: "base", text: "optional base" },
    ]);
    expect(paramDoc(multi, "[base]")).toBe("optional base");
    expect(paramDoc(multi, "ROW")).toBe("a row of values over several lines");
    expect(paramDoc(multi, "other")).toBeUndefined();
  });

  it("leaves text without tags (and an e-mail-like @param) alone", () => {
    expect(parseDocComment("Plain AFE comment.")).toEqual({ summary: "Plain AFE comment.", params: [] });
    expect(parseDocComment("see x@param y").params).toEqual([]);
  });

  it("survives the module round trip: the whole text is the Name Manager comment", () => {
    const doc = "Adds two numbers.\n@param a the first\n@param b the second";
    const text = `${formatDoc(doc)}\nFN.ADD = LAMBDA(a, b, a + b);\n`;
    expect(parseModule(text).entries[0]!.doc).toBe(doc);
    expect(docText(" Sum. @param a x ")).toBe("Sum. @param a x");
  });
});
