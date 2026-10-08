// The seeded-trap workbook (brief §7): probes/fixtures/traps.xlsx, made without Excel by
// probes/fixtures/make-traps.mjs from an Excel-saved probe file. The audit must catch every
// seeded trap, exactly, and nothing else.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { audit, readWorkbook, renderAuditReport, type Finding } from "../../src/index.js";
// @ts-ignore: a plain ES module, run by node to make the fixture
import { BASE, makeTraps, OUT, TRAP_NAMES } from "../../../../probes/fixtures/make-traps.mjs";

const bytes = new Uint8Array(readFileSync(OUT as string));
const wb = readWorkbook(bytes);
const report = audit(wb, { workbook: "traps.xlsx" });

const place = (f: Finding) => (f.where.kind === "name" ? `name ${f.where.key}` : `${f.where.kind} ${f.where.sheet}!${f.where.ref}`);

describe("traps.xlsx", () => {
  it("is what make-traps.mjs makes from f7_base.xlsx", () => {
    const again = unzipSync(makeTraps(new Uint8Array(readFileSync(BASE as string))) as Uint8Array);
    const committed = unzipSync(bytes);
    expect(Object.keys(again).sort()).toEqual(Object.keys(committed).sort());
    for (const k of Object.keys(committed)) expect(Buffer.from(again[k]!).equals(Buffer.from(committed[k]!)), k).toBe(true);
  });

  it("reads as a clean package with the seeded names and the chart", () => {
    expect(wb.warnings).toEqual([]);
    for (const [n] of TRAP_NAMES as [string][]) expect(wb.definedNames.some((d) => d.name === n), n).toBe(true);
    expect(wb.charts).toEqual([
      {
        part: "xl/charts/chart1.xml",
        kind: "chart",
        sheet: { position: 0, name: "S1" },
        drawing: "xl/drawings/drawing1.xml",
        formulas: [
          { text: "'S1'!$A$2:$A$11", element: "cat" },
          { text: "[0]!ChartOnly", element: "val" },
        ],
      },
    ]);
  });

  it("catches every seeded trap, and nothing else", () => {
    expect(report.findings.map((f) => `${f.rule} ${place(f)}`)).toEqual([
      "C2.bare-prefix name BareSeq",
      "C2.poisoned name Poisoned",
      "C3.lambda name Fact",
      "C4.ref-deleted name Broken",
      "C4.relative name RelRef",
      "C5.other-sheet name ReadsLocal",
      "C6.lambda-arity cell S1!A18",
      "C6.lambda-arity cell S1!A20",
      "C7.length name Long",
      "C9.fixed-ref name SpillFixed",
      "C9.fixed-ref cell S1!A13",
      "C10.unused name Fact",
      "C10.unused name Unused",
      "C11.drift name Margin_high",
      "C12.name-cycle name CycA",
      "C13.constant name Tax",
    ]);
    expect(report.counts).toEqual({ error: 9, warning: 6, info: 1 });
  });

  it("says what each trap is, where", () => {
    const by = (rule: string, key?: string) => report.findings.find((f) => f.rule === rule && (key === undefined || f.where.key === key || f.where.ref === key))!;
    expect(by("C2.bare-prefix").where).toMatchObject({ span: { start: 0, end: 8 }, text: "SEQUENCE" });
    expect(by("C2.bare-prefix").message).toContain("SEQUENCE is stored without its _xlfn. prefix");
    expect(by("C5.other-sheet").message).toBe("OnlySecond exists only on 'S2', not at workbook scope: unqualified it is #NAME?");
    expect(by("C5.other-sheet").hint).toBe("write 'S2'!OnlySecond");
    expect(by("C6.lambda-arity", "A18").message).toBe("Fn(x) takes 1 argument; it is given 2");
    expect(by("C6.lambda-arity", "A20").message).toBe("Opt(a, [b]) takes 1 to 2 arguments; it is given 3");
    expect(by("C7.length").data).toEqual({ length: 8301 });
    expect(by("C9.fixed-ref", "SpillFixed").hint).toBe("use 'S1'!E1#");
    expect(by("C9.fixed-ref", "A13").hint).toBe("use E1#");
    expect(by("C11.drift").message).toBe("Margin_high differs from Margin_base beyond the suffix (high vs base): `*1.1` here, nothing there");
    expect(by("C12.name-cycle").message).toBe("name cycle: CycA ↔ CycB read each other through their definitions");
    expect(by("C13.constant").message).toBe("0.27 hard-coded in the LAMBDA body");
  });

  it("counts a name used only in a chart as used, and Opt(1) as a good call", () => {
    expect(report.findings.some((f) => f.where.key === "ChartOnly")).toBe(false);
    expect(report.findings.some((f) => f.where.ref === "A19")).toBe(false);
    // Without the chart, ChartOnly would be unused (it names cells: info).
    const noChart = audit({ ...wb, charts: [] });
    expect(noChart.findings.filter((f) => f.check === "C10").map((f) => `${f.rule} ${f.where.key}`)).toEqual(["C10.unused-cell ChartOnly", "C10.unused Fact", "C10.unused Unused"]);
  });

  it("takes the census and the spill census", () => {
    const c = report.census;
    expect(c.total).toBe(31);
    expect(c.byScope).toEqual({ workbook: 29, sheets: [{ sheet: "S2", names: 2 }] });
    expect(c.coordinates).toEqual([
      { tag: "base", position: "suffix", families: 4 },
      { tag: "high", position: "suffix", families: 4 },
    ]);
    expect(c.families.map((f) => f.stem)).toEqual(["Cost", "Margin", "Sales", "Volume"]);
    expect(c.tiers).toMatchObject({ T1: 0, T3: 4, T4: 1, library: 4 });
    expect(c.tierNames.T4).toEqual(["BareSeq"]);
    expect(report.spills.spills).toEqual([
      {
        sheet: "S1",
        anchor: "E1",
        extent: "E1:E3",
        rows: 3,
        cols: 1,
        formula: "SEQUENCE(3)*Rate",
        names: [
          { key: "SpillFixed", how: "extent" },
          { key: "Spl", how: "spill" },
        ],
      },
    ]);
  });

  it("filters by check and severity, and overrides severities", () => {
    expect(audit(wb, { only: ["C2", "C9"] }).findings.map((f) => f.check)).toEqual(["C2", "C2", "C9", "C9"]);
    expect(audit(wb, { minSeverity: "warning" }).counts).toEqual({ error: 9, warning: 6, info: 0 });
    const r = audit(wb, { rules: { C13: "off", "C10.unused": "info", "C4.relative": "error" } });
    expect(r.findings.some((f) => f.check === "C13")).toBe(false);
    expect(r.findings.filter((f) => f.rule === "C10.unused").every((f) => f.severity === "info")).toBe(true);
    expect(r.counts).toEqual({ error: 10, warning: 3, info: 2 });
  });

  it("is stable: the same report twice, and as text", () => {
    expect(JSON.stringify(audit(readWorkbook(bytes), { workbook: "traps.xlsx" }))).toBe(JSON.stringify(report));
    const { text, links } = renderAuditReport(report);
    expect(text).toContain("xln check traps.xlsx: 9 errors, 6 warnings, 1 info");
    expect(text).toContain("  error   name ReadsLocal: OnlySecond exists only on 'S2'");
    expect(text).toContain("names standing in for dimensions: ");
    expect(text).toContain("== C9 spill census: 1 dynamic array spilled when saved");
    for (const l of links) expect(text.slice(l.start, l.end).length).toBeGreaterThan(0);
    expect(links.map((l) => text.slice(l.start, l.end))).toContain("'S1'!A18");
    expect(links.map((l) => text.slice(l.start, l.end))).toContain("ReadsLocal");
  });
});
