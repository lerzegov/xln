// M5 end to end: `renameInProject` edits the source (the name, `@renamed(Old)`, every
// reader), the build renames the name and rewrites the token in the workbook (stretch
// G), a second build has nothing to do, and a fresh pull gives back the renamed source.
import { describe, expect, it } from "vitest";
import { LOCK_FILE, pullProject, readWorkbook, renameInProject, unbuiltEdits, formatUnbuiltEdit, type SourceRename, type WorkbookSnapshot, rewriteZip, utf8 } from "../../src/index.js";
import { Package } from "../../src/file/package.js";
import { build, edit, fixture, pulled, why } from "./helpers.js";

const F7 = fixture("f7_base.xlsx");
const ORACLE = fixture("f7_oracle.xlsx");
const U = "names/_unmanaged.xln";
const S1 = "names/sheets/S1.xln";

function renamed(files: Record<string, string>, name: string, to: string): Record<string, string> {
  const r = renameInProject(files, name, to);
  if (typeof r === "string") throw new Error(r);
  return { ...files, ...r.files };
}

/** Every cell formula and definition, by place. */
function formulas(wb: WorkbookSnapshot): Map<string, string> {
  const out = new Map<string, string>();
  for (const d of wb.definedNames) if (!d.isXlPrefixed) out.set(`name ${d.scope.kind === "sheet" ? d.scope.name + "!" : ""}${d.name}`, d.definition);
  for (const s of wb.sheets) {
    for (const f of s.formulas) out.set(`cell ${s.name}!${f.cell}`, `${f.kind} ${f.text ?? ""} ${f.range ?? ""} ${f.cm ?? ""}`);
    s.conditionalFormats.forEach((c, i) => out.set(`cf ${s.name}#${i}`, c.formulas.join("|")));
    s.dataValidations.forEach((v, i) => out.set(`dv ${s.name}#${i}`, `${v.formula1 ?? ""}|${v.formula2 ?? ""}`));
  }
  return out;
}

function diff(a: Map<string, string>, b: Map<string, string>): string[] {
  return [...new Set([...a.keys(), ...b.keys()])].filter((k) => a.get(k) !== b.get(k)).map((k) => `${k}: ${a.get(k)} | ${b.get(k)}`);
}

/** Pull files without the lockfile and manifest, and without `@renamed` lines (a pull does not write them). */
function sourceOnly(files: Record<string, string>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [p, t] of Object.entries(files)) {
    if (!p.endsWith(".xln")) continue;
    out[p] = t
      .split("\n")
      .filter((l) => !l.trim().startsWith("@renamed("))
      .join("\n");
  }
  return out;
}

describe("renameInProject: the source edits", () => {
  it("renames the statement, adds @renamed(Old) and rewrites every reader, F7's traps kept", () => {
    const files = pulled(F7);
    const r = renameInProject(files, "Rate", "Pace") as SourceRename;
    expect(typeof r).toBe("object");
    expect(r).toMatchObject({ from: "Rate", to: "Pace", annotation: "added" });
    const u = r.files[U]!;
    expect(u).toContain("Fn = LAMBDA(x, x*Pace);");
    expect(u).toContain("@renamed(Rate)\nPace = 0.1;");
    expect(u).toContain("Rate2 = Pace*2;");
    expect(u).toContain("RateX = 0.5;");
    const s1 = r.files[S1]!;
    for (const line of ["@A1 = Pace*2;", "@C1 = RateX+Pace;", "Spl @E1# = SEQUENCE(3)*Pace;", "@B2:B11 = A2*Pace;", '@C2 = "Rate is "&Pace;', "@C3 = LET(Rate, 5, Rate*2);", '@C7 = INDIRECT("Rate");', "@C8 = Pace*3;", "@C9 = ROWS(E1#)*Pace;", "@C10 = Rate2+Fn(1);"]) {
      expect(s1).toContain(line);
    }
    // S2: Rate+Loc.
    expect(r.files["names/sheets/S2.xln"]).toContain("Pace+Loc");
  });

  it("a sheet-local name: Sheet!Name, and only its readers", () => {
    const r = renameInProject(pulled(F7), "S2!Loc", "Spot") as SourceRename;
    expect(r.from).toBe("S2!Loc");
    expect(r.files[S1]).toContain("@C5 = 'S2'!Spot+Loc;");
    const s2 = r.files["names/sheets/S2.xln"]!;
    expect(s2).toContain("@renamed(Loc)\nSpot = 7;");
    for (const line of ["@A1 = Spot;", "@B1 = 'S1'!A1+'S2'!Spot;", "@A2 = Rate+Spot;", "@A3:A5 = Spot*ROW();"]) expect(s2).toContain(line);
    expect(r.files[U]).toBeUndefined(); // the workbook's Loc and its readers stay
  });

  it("twice before a build keeps the first @renamed; renamed back drops it", () => {
    const once = renamed(pulled(F7), "Rate", "Pace");
    const twice = renameInProject(once, "Pace", "Speed") as SourceRename;
    expect(twice.annotation).toBe("kept");
    expect(twice.files[U]).toContain("@renamed(Rate)\nSpeed = 0.1;");
    const back = renameInProject({ ...once, ...twice.files }, "Speed", "Rate") as SourceRename;
    expect(back.annotation).toBe("removed");
    expect(back.files[U]).toContain("\nRate = 0.1;");
    expect(back.files[U]).not.toContain("@renamed");
    expect({ ...once, ...twice.files, ...back.files }).toEqual(pulled(F7));
  });

  it("a name not built yet gets no @renamed", () => {
    const files = pulled(F7);
    files[U] += "Fresh = 1;\nUsesFresh = Fresh+1;\n";
    const r = renameInProject(files, "Fresh", "Newer") as SourceRename;
    expect(r.annotation).toBe("none");
    expect(r.files[U]).toContain("Newer = 1;\nUsesFresh = Newer+1;");
  });

  it("refuses: unknown, ambiguous, invalid, taken, a Table's name, a capture", () => {
    const files = pulled(F7);
    expect(renameInProject(files, "Nope", "X1y")).toBe("no name Nope in the project");
    expect(renameInProject(files, "Rate", "A1")).toMatch(/cell reference/);
    expect(renameInProject(files, "Fn", "Sum")).toMatch(/built-in function/);
    // A value spelled like a built-in is read as the name (probe T12 is about calls).
    expect(renameInProject(files, "Rate", "Sum")).toMatchObject({ to: "Sum" });
    expect(renameInProject(files, "Rate", "RateX")).toMatch(/^RateX exists already \(names\/_unmanaged\.xln:\d+\)/);
    expect(renameInProject(files, "Rate", "Sales", { tables: ["Sales"] })).toMatch(/name of a Table/);
    // Fn = LAMBDA(x, x*Rate): Rate → x would read the variable.
    expect(renameInProject(files, "Rate", "x")).toMatch(/would change what this formula reads.*_unmanaged\.xln:\d+ \(Fn\)/);
    // The workbook's Rate → Loc: S2's own Loc would capture S2's readers.
    expect(renameInProject(files, "Rate", "Loc")).toMatch(/exists already/);
    expect(renameInProject(files, "Rate", "Spot")).toMatchObject({ to: "Spot" });
    const s2 = renamed(files, "S2!Loc", "Spot");
    expect(renameInProject(s2, "Rate", "Spot")).toMatch(/would change what this formula reads.*S2\.xln/);
  });
});

describe("rename across cells: build, rebuild, pull (F7)", () => {
  it("renamed, built, renamed back to a built-in's spelling (Rate, RATE), built: the workbook's names again", () => {
    const once = build(F7, renamed(pulled(F7), "Rate", "Pace"));
    expect(once.status, why(once)).toBe("built");
    const files = pulled(once.bytes!);
    const back = build(once.bytes!, renamed(files, "Pace", "Rate"));
    expect(back.status, why(back)).toBe("built");
    const names = (wb: WorkbookSnapshot) => wb.definedNames.filter((d) => !d.isXlPrefixed).map((d) => d.name).sort();
    expect(names(readWorkbook(back.bytes!))).toEqual(names(readWorkbook(F7)));
  });

  it("the three renames of probe F7 in one build give Excel's own result", () => {
    let files = pulled(F7);
    files = renamed(files, "Rate", "Pace");
    files = renamed(files, "Fn", "Fx");
    files = renamed(files, "S2!Loc", "Spot");
    const r = build(F7, files);
    expect(r.status, why(r)).toBe("built");
    expect(r.readBack!.problems).toEqual([]);
    // Only renames: no cell is set, nothing goes to the dynamic-array form.
    expect(r.plan.changeSet.changes.map((c) => c.op)).toEqual(["rename-name", "rename-name", "rename-name"]);
    expect(r.plan.changeSet.changes[0]).toMatchObject({ from: "Fn", to: "Fx", references: { cells: 2, names: 0 } });
    // Excel's oracle renamed Rate to Growth, which xln refuses as a new name (GROWTH is a
    // built-in: a call would reach it, T12); Pace stands in for it.
    const oracle = new Map([...formulas(readWorkbook(ORACLE))].map(([k, v]) => [k.split("Growth").join("Pace"), v.split("Growth").join("Pace")]));
    expect(diff(formulas(readWorkbook(r.bytes!)), oracle)).toEqual([]);
    expect(new Package(r.bytes!).has("xl/calcChain.xml")).toBe(false);
    expect(new Package(r.bytes!).text("xl/workbook.xml")).toContain('fullCalcOnLoad="1"');

    // Built: nothing left to build, and the pull guard sees no unbuilt edit.
    const after = { ...files, ...r.files };
    const again = build(r.bytes!, after);
    expect(again.status, why(again)).toBe("up-to-date");
    expect(unbuiltEdits({ workbook: r.bytes!, fileName: "book.xlsx", files: after })).toEqual([]);
    // A fresh pull gives back the source, without the @renamed lines.
    expect(sourceOnly(pulled(r.bytes!))).toEqual(sourceOnly(after));
  });

  it("before the build, the pull guard lists the rename once", () => {
    const files = renamed(pulled(F7), "Rate2", "RateTwo");
    expect(unbuiltEdits({ workbook: F7, fileName: "book.xlsx", files }).map(formatUnbuiltEdit)).toEqual(["names/_unmanaged.xln:9  RateTwo: rename Rate2 → RateTwo, rewriting it in 1 cell formula"]);
  });

  it("a renamed reader edited further is set as usual, the others only renamed", () => {
    const files = renamed(pulled(F7), "Rate", "Pace");
    edit(files, S1, "@C8 = Pace*3;", "@C8 = Pace*4;");
    const r = build(F7, files);
    expect(r.status, why(r)).toBe("built");
    expect(r.plan.changeSet.changes.map((c) => (c.op === "set-cell-formula" ? `${c.op} ${c.range}` : c.op))).toEqual(["rename-name", "set-cell-formula C8"]);
    const wb = readWorkbook(r.bytes!);
    const s1 = wb.sheets.find((s) => s.name === "S1")!;
    expect(s1.formulas.find((f) => f.cell === "C8")).toMatchObject({ kind: "dynamic-array", text: "Pace*4" });
    expect(s1.formulas.find((f) => f.cell === "A1")).toMatchObject({ kind: "normal", text: "Pace*2" });
    expect(s1.formulas.find((f) => f.cell === "B2")).toMatchObject({ kind: "shared-master", text: "A2*Pace" });
  });

  it("refuses when a cell statement in the source still reads the old name", () => {
    const files = pulled(F7);
    edit(files, U, "Rate2 = Rate*2;", "@renamed(Rate2)\nRateTwo = Rate*2;");
    const r = build(F7, files);
    expect(r.status).toBe("refused");
    expect(r.plan.problems.find((p) => p.code === "in-use")!.message).toContain("the source renames Rate2 to RateTwo, so write RateTwo here");
  });

  it("refuses a rename that Excel's formulas would read differently, with the places", () => {
    // A name made in Excel since the pull (not in the source) whose LAMBDA has a variable g.
    const files = renamed(pulled(F7), "Rate", "g");
    const wbXml = new Package(F7).text("xl/workbook.xml")!;
    const withG = wbXml.replace('<definedName name="Loc"', '<definedName name="Mul">_xlfn.LAMBDA(_xlpm.g, _xlpm.g*Rate)</definedName><definedName name="Loc"');
    const wb = rewriteZip(F7, { replace: new Map([["xl/workbook.xml", utf8(withG)]]) });
    const r = build(wb, files);
    expect(r.status).toBe("refused");
    const p = r.plan.problems.find((x) => x.code === "rename-capture")!;
    expect(p.message).toMatch(/^renaming Rate to g refused: in 1 place the new name would read something else: name Mul \(would read g\)/);
  });

  it("refuses when a chart or a Table column reads the name, and says to rename in Excel", () => {
    const files = renamed(pulled(F7), "Rate", "Pace");
    const chart = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>[0]!Rate</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>';
    const wb = rewriteZip(F7, { add: new Map([["xl/charts/chart1.xml", utf8(chart)]]) });
    const r = build(wb, files);
    expect(r.status).toBe("refused");
    const p = r.plan.problems.find((x) => x.code === "in-use")!;
    expect(p.sites).toEqual(["chart xl/charts/chart1.xml"]);
    expect(p.message).toContain("rename it in Excel's Name Manager instead");
  });

  it("refuses for a pivot table's source, a hyperlink, a form control", () => {
    const files = renamed(pulled(F7), "Rate", "Pace");
    const pivot = '<?xml version="1.0" encoding="UTF-8"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource name="Rate"/></cacheSource></pivotCacheDefinition>';
    const ctrl = '<?xml version="1.0" encoding="UTF-8"?><formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="Spin" fmlaLink="Rate"/>';
    const sheet = new Package(F7).text("xl/worksheets/sheet2.xml")!.replace("<pageMargins", '<hyperlinks><hyperlink ref="D1" location="Rate" display="go"/></hyperlinks><pageMargins');
    const wb = rewriteZip(F7, {
      replace: new Map([["xl/worksheets/sheet2.xml", utf8(sheet)]]),
      add: new Map([
        ["xl/pivotCache/pivotCacheDefinition1.xml", utf8(pivot)],
        ["xl/ctrlProps/ctrlProp1.xml", utf8(ctrl)],
      ]),
    });
    const r = build(wb, files);
    expect(r.status).toBe("refused");
    expect(r.plan.problems.find((x) => x.code === "in-use")!.sites).toEqual(["hyperlink on S2!D1", "pivot table source xl/pivotCache/pivotCacheDefinition1.xml", "form control xl/ctrlProps/ctrlProp1.xml (fmlaLink)"]);
  });

  it("a rename that also moves the name is D4's as before: no token rewrite; readers are the source's to change", () => {
    const files = pulled(F7);
    // RateX to sheet S1 under a new name.
    edit(files, U, "RateX = 0.5;\n", "");
    files[S1] += "@renamed(!RateX)\nRateY = 0.5;\n";
    const refused = build(F7, files);
    expect(refused.status).toBe("refused");
    expect(refused.plan.problems.find((p) => p.code === "in-use" && p.sites)!.sites).toEqual(["cell S1!C1"]);
    // The reader changed in the source: written as a new formula (dynamic-array form).
    edit(files, S1, "@C1 = RateX+Rate;", "@C1 = RateY+Rate;");
    const r = build(F7, files);
    expect(r.status, why(r)).toBe("built");
    expect(r.plan.changeSet.changes.find((c) => c.op === "rename-name")).toEqual({ op: "rename-name", scope: null, from: "RateX", to: "RateY" });
    expect(r.plan.changeSet.changes.map((c) => c.op)).toContain("set-cell-formula");
  });

  it("the lockfile after the build records the renamed cells, so the next pull's guard is quiet", () => {
    const files = renamed(pulled(F7), "Rate", "Pace");
    const r = build(F7, files);
    const lock = JSON.parse(r.files![LOCK_FILE]!) as { names: Record<string, unknown> };
    expect(Object.keys(lock.names)).toContain("Pace");
    expect(Object.keys(lock.names)).not.toContain("Rate");
    expect(pullProject(r.bytes!, "book.xlsx").files[LOCK_FILE]).toBe(r.files![LOCK_FILE]);
  });
});
