import { describe, expect, it } from "vitest";
import { activeCall, formulaCursor } from "../../src/index.js";

/** The cursor where `|` is written. */
function at(text: string) {
  const k = text.indexOf("|");
  return formulaCursor(text.slice(0, k) + text.slice(k + 1), k);
}

describe("formulaCursor", () => {
  it("finds the identifier being typed, with its module prefix", () => {
    expect(at("EBIT - Ta|").word.text).toBe("Ta");
    expect(at("FN.SE|").word).toMatchObject({ start: 0, text: "FN.SE" });
    expect(at("FN.|").word.text).toBe("FN.");
    expect(at("1 + |").word.text).toBe("");
    expect(at("Sales| * 2").word.text).toBe("Sales");
  });

  it("reads the sheet before !, quoted or not", () => {
    expect(at("IS!Sa|")).toMatchObject({ sheet: "IS", qualifierStart: 0, word: { text: "Sa" } });
    expect(at("1+IS!|")).toMatchObject({ sheet: "IS", qualifierStart: 2, word: { text: "" }, inert: false });
    expect(at("'SCF recursive'!Ex|")).toMatchObject({ sheet: "SCF recursive", qualifierStart: 0 });
    expect(at("'It''s'!|").sheet).toBe("It's");
  });

  it("is inert inside strings, structured references, quoted sheets, numbers and $ addresses", () => {
    expect(at('IF(x = "ab|').inert).toBe(true);
    expect(at("tbl[co|").inert).toBe(true);
    expect(at("'SCF rec|").inert).toBe(true);
    expect(at("1.5E|").inert).toBe(true);
    expect(at("$A|").inert).toBe(true);
    expect(at('IF(x = "ab", Sa|').inert).toBe(false);
    expect(at("Sa|").inert).toBe(false);
  });

  it("tracks open calls and the argument at the cursor", () => {
    const c = at("IF(Sales > 0, SUM(A1, |");
    expect(c.frames.map((f) => f.fn)).toEqual(["IF", "SUM"]);
    expect(activeCall(c)).toMatchObject({ frame: { fn: "SUM" }, index: 1 });
    // A group or an array inside a call: still the call's argument.
    const g = at("SUM(1, (2 + |");
    expect(activeCall(g)).toMatchObject({ frame: { fn: "SUM" }, index: 1 });
    expect(activeCall(at("INDEX({1,2,3}, |"))!.index).toBe(1);
    // Closed calls are gone; a qualified LAMBDA name keeps its sheet.
    expect(activeCall(at("SUM(1, 2) + MAX(|"))!.frame.fn).toBe("MAX");
    expect(activeCall(at("'S 1'!Fn(1, |"))!.frame).toMatchObject({ fn: "Fn", sheet: "S 1" });
    expect(activeCall(at("Sales|"))).toBeUndefined();
  });

  it("lists LET and LAMBDA variables in scope, innermost first", () => {
    expect(at("LET(a, 1, b, a + |").locals.map((l) => l.name)).toEqual(["a"]);
    expect(at("LET(a, 1, b, 2, |").locals.map((l) => l.name)).toEqual(["b", "a"]);
    expect(at("LET(a, 1|").locals).toEqual([]);
    expect(at("LAMBDA(x, [y], x + |").locals.map((l) => [l.name, l.kind])).toEqual([
      ["y", "lambda"],
      ["x", "lambda"],
    ]);
    const nested = at("LET(rate, 0.1, f, LAMBDA(v, v * |");
    expect(nested.locals.map((l) => l.name)).toEqual(["v", "rate"]);
    // Closed scopes are gone; `_xlpm.` is dropped.
    expect(at("LET(a, 1, a) + |").locals).toEqual([]);
    expect(at("LAMBDA(_xlpm.n, |").locals.map((l) => l.name)).toEqual(["n"]);
    const v = at("LET(sq, LAMBDA(x, x*x), sq(|").locals[0]!;
    expect(v.name).toBe("sq");
    expect("LET(sq, LAMBDA(x, x*x), sq(".slice(v.value!.start, v.value!.end).trim()).toBe("LAMBDA(x, x*x)");
  });
});
