// Every edit type of D3 and D4 on the probe workbooks, each read back (E3), plus the
// three-way check (E2) and the change set's shape (D9).
import { describe, expect, it } from "vitest";
import {
  applyChangeSet,
  definitionHash,
  LOCK_FILE,
  parseLockfile,
  pullProject,
  rawZipRecords,
  readWorkbook,
  scanWorkbookXml,
  type DefinedName,
} from "../../src/index.js";
import { build, definedNamesXml, edit, fixture, pulled, why, withWorkbookXml, workbookXml } from "./helpers.js";

const F7 = fixture("f7_base.xlsx");
const MAC = fixture("probe_mac.xlsx");
const U = "names/_unmanaged.xln";

function names(bytes: Uint8Array): DefinedName[] {
  return readWorkbook(bytes).definedNames;
}

function find(bytes: Uint8Array, name: string, sheet?: string): DefinedName | undefined {
  return names(bytes).find((d) => d.name === name && (sheet === undefined ? d.scope.kind === "workbook" : d.scope.kind === "sheet" && d.scope.name === sheet));
}

function built(bytes: Uint8Array, files: Record<string, string>) {
  const r = build(bytes, files);
  expect(r.status, why(r)).toBe("built");
  expect(r.readBack!.ok, r.readBack!.problems.join("\n")).toBe(true);
  // A second build from the same source, against the built file and the new lockfile, has nothing to do.
  const again = build(r.bytes!, { ...files, ...r.files });
  expect(again.status, why(again)).toBe("up-to-date");
  return r;
}

describe("D3 edits", () => {
  it("updates a definition; nothing else in <definedNames> moves", () => {
    const files = pulled(F7);
    edit(files, U, "Rate = 0.1;", "Rate = 0.25;");
    const r = built(F7, files);
    expect(r.plan.changeSet.changes).toEqual([{ op: "set-name", name: "Rate", scope: null, stored: "0.25", display: "0.25", comment: null, hidden: false, fields: ["definition"] }]);
    expect(definedNamesXml(r.bytes!)).toBe(definedNamesXml(F7).replace('<definedName name="Rate">0.1<', '<definedName name="Rate">0.25<'));
  });

  it("creates names: prefixes, _xlpm., home-sheet qualification, # → ANCHORARRAY, CR LF, sorted place", () => {
    const files = pulled(F7);
    files[U] += "\nAddRate = LAMBDA(v,\n    LET(r, Rate, v * (1 + r))\n);\nSeq = SEQUENCE(3) * Rate;\n";
    files["names/sheets/S2.xln"] += "\nLocSpill = $E$1# + A1;\n";
    const r = built(F7, files);
    expect(r.plan.changeSet.changes.map((c) => c.op === "set-name" && c.fields)).toEqual([["created"], ["created"], ["created"]]);
    expect(find(r.bytes!, "AddRate")!.definition).toBe("_xlfn.LAMBDA(_xlpm.v,\r\n    _xlfn.LET(_xlpm.r, Rate, _xlpm.v * (1 + _xlpm.r))\r\n)");
    expect(find(r.bytes!, "Seq")!.definition).toBe("_xlfn.SEQUENCE(3) * Rate");
    expect(find(r.bytes!, "LocSpill", "S2")!.definition).toBe("_xlfn.ANCHORARRAY('S2'!$E$1) + 'S2'!A1");
    // Raw CR LF in the part, as Excel writes it; Excel's order: by name, case ignored.
    expect(workbookXml(r.bytes!)).toContain("_xlfn.LAMBDA(_xlpm.v,\r\n    _xlfn.LET(");
    expect(names(r.bytes!).map((d) => d.name)).toEqual(["AddRate", "Fn", "Loc", "Loc", "LocSpill", "Rate", "Rate2", "RateX", "Seq", "Spl"]);
    expect(find(r.bytes!, "LocSpill", "S2")!.scope).toEqual({ kind: "sheet", position: 1, name: "S2" });
  });

  it("changes a comment (multi-line, XML specials) and the hidden flag", () => {
    const files = pulled(F7);
    edit(files, U, "RateX = 0.5;", '/**\n * Rate "x" <b> & co\n * second line\n */\n@hidden\nRateX = 0.5;');
    const r = built(F7, files);
    expect(r.plan.changeSet.changes).toMatchObject([{ op: "set-name", name: "RateX", fields: ["comment", "hidden"], stored: "0.5" }]);
    const d = find(r.bytes!, "RateX")!;
    expect(d.comment).toBe('Rate "x" <b> & co\nsecond line');
    expect(d.hidden).toBe(true);
    expect(definedNamesXml(r.bytes!)).toContain('<definedName name="RateX" comment="Rate &quot;x&quot; &lt;b&gt; &amp; co&#10;second line" hidden="1">0.5</definedName>');
    // And back: a pull of the built file gives the same source text.
    expect(pullProject(r.bytes!, "book.xlsx").files[U]).toContain('/**\n * Rate "x" <b> & co\n * second line\n */\n@hidden\nRateX = 0.5;');
  });

  it("deletes an unused name", () => {
    const files = pulled(MAC);
    edit(files, U, "Fact = LAMBDA(n, 1);\n", "");
    const r = built(MAC, files);
    expect(r.plan.changeSet.changes).toEqual([{ op: "delete-name", name: "Fact", scope: null }]);
    expect(find(r.bytes!, "Fact")).toBeUndefined();
    expect(names(r.bytes!).length).toBe(names(MAC).length - 1);
  });

  it("refuses to delete a name cells use, listing the cells", () => {
    const files = pulled(F7);
    edit(files, U, "RateX = 0.5;\n", "");
    const r = build(F7, files);
    expect(r.status).toBe("refused");
    expect(r.plan.problems).toMatchObject([{ severity: "error", code: "in-use", key: "RateX", sites: ["cell S1!C1"] }]);
  });

  it("changes scope: sheet → workbook", () => {
    const files = pulled(MAC);
    edit(files, "names/P.xln", "@sheet('S2')\nP_Local2 = 6;", "P_Local2 = 6;");
    const r = built(MAC, files);
    expect(r.plan.changeSet.changes).toEqual([{ op: "rescope-name", name: "P_Local2", from: "S2", to: null }]);
    expect(find(r.bytes!, "P_Local2")!.definition).toBe("6");
    expect(find(r.bytes!, "P_Local2", "S2")).toBeUndefined();
  });

  it("refuses a scope change that would make cells read another name", () => {
    // S2's Loc moved to S1: S2's cells would read the workbook's Loc instead.
    const files = pulled(F7);
    edit(files, "names/sheets/S2.xln", "Loc = 7;\n", "");
    files["names/sheets/S1.xln"] += "\nLoc = 7;\n";
    const r = build(F7, files);
    expect(r.status).toBe("refused");
    const p = r.plan.problems.find((x) => x.code === "in-use")!;
    expect(p.key).toBe("S2!Loc");
    expect(p.sites).toContain("cell S2!A1");
  });
});

describe("D4 rename within names", () => {
  it("renames an unused name with @renamed, and its doc follows", () => {
    const files = pulled(MAC);
    edit(files, U, "Fact = LAMBDA(n, 1);", "@renamed(Fact)\nFactOne = LAMBDA(n, 1);");
    const r = built(MAC, files);
    expect(r.plan.changeSet.changes).toEqual([{ op: "rename-name", scope: null, from: "Fact", to: "FactOne" }]);
    expect(find(r.bytes!, "Fact")).toBeUndefined();
    expect(find(r.bytes!, "FactOne")!.definition).toBe("_xlfn.LAMBDA(_xlpm.n, 1)");
    expect(names(r.bytes!).map((d) => d.name)).toEqual(["FactOne", "Growλ", "Mod.Fn", "P_Add1", "P_Local2", "P_Long3990", "P_Multi", "P_Spill"]);
  });

  it("renames and edits in one go", () => {
    const files = pulled(MAC);
    edit(files, U, "Fact = LAMBDA(n, 1);", "@renamed(Fact)\nOne = LAMBDA(n, 1 + 0);");
    const r = built(MAC, files);
    expect(r.plan.changeSet.changes).toMatchObject([
      { op: "rename-name", from: "Fact", to: "One" },
      { op: "set-name", name: "One", fields: ["definition"], stored: "_xlfn.LAMBDA(_xlpm.n, 1 + 0)" },
    ]);
  });

  it("refuses a rename while a cell statement still reads the old name; with it renamed, rewrites the cell's token", () => {
    const files = pulled(MAC);
    edit(files, "names/P.xln", "P_Add1 = LAMBDA(x, x+1);", "@renamed(P_Add1)\nP_Inc = LAMBDA(x, x+1);");
    const r = build(MAC, files);
    expect(r.status).toBe("refused");
    expect(r.plan.problems).toMatchObject([{ code: "in-use", file: "names/sheets/S1.xln" }]);
    expect(r.plan.problems[0]!.message).toContain("the source renames P_Add1 to P_Inc, so write P_Inc here");
    // M5: the reader renamed in the source too, the build rewrites the cell's token.
    files["names/sheets/S1.xln"] = files["names/sheets/S1.xln"]!.split("P_Add1(").join("P_Inc(");
    const ok = built(MAC, files);
    expect(ok.plan.changeSet.changes).toMatchObject([{ op: "rename-name", from: "P_Add1", to: "P_Inc", references: { cells: 1 } }]);
  });

  it("refuses a rename when another source name still uses the old name", () => {
    const files = pulled(F7);
    edit(files, U, "RateX = 0.5;", "RateX = 0.5;\nUsesX = RateX * 2;");
    const first = built(F7, files);
    const files2 = { ...files, ...first.files };
    edit(files2, U, "RateX = 0.5;", "@renamed(RateX)\nRateY = 0.5;");
    const r = build(first.bytes!, files2);
    expect(r.status).toBe("refused");
    expect(r.plan.problems.map((p) => p.code)).toContain("in-use");
    expect(r.plan.problems.some((p) => p.message.startsWith("RateX is not a name") && p.message.includes("the source renames RateX to RateY, so write RateY here"))).toBe(true);
  });

  it("refuses invalid and colliding names", () => {
    const files = pulled(F7);
    // Sum = 1 is a value: read as the name; only a LAMBDA spelled like a built-in is refused (T12).
    files[U] += "\nSum = 1;\nFact = LAMBDA(x, x);\nA1 = 2;\nR2C3 = 3;\n";
    const r = build(F7, files);
    expect(r.plan.problems.filter((p) => p.code === "invalid-name").map((p) => p.key)).toEqual(["Fact", "A1", "R2C3"]);
  });

  it("warns on a LAMBDA named like an Excel 4.0 macro function, and builds it", () => {
    const files = pulled(F7);
    // Group and Get.Cell are macro functions; Save.As? and Open are commands (no warning);
    // a value named Files is no call.
    files[U] += "\nGroup = LAMBDA(x, x);\nget.cell = LAMBDA(x, x);\nSave.As? = LAMBDA(x, x);\nOpen = LAMBDA(x, x);\nFiles = 1;\nGroupBy2 = LAMBDA(x, x);\n";
    const r = build(F7, files);
    const xlm = r.plan.problems.filter((p) => p.code === "xlm-name");
    expect(xlm.map((p) => [p.key, p.severity])).toEqual([["Group", "warning"], ["get.cell", "warning"]]);
    expect(xlm[0]!.message).toBe("Group is also an Excel 4.0 macro function: Excel may call that instead or refuse the name (AFE #10, not measured)");
    expect(r.status, why(r)).toBe("built");
  });
});

describe("E2 three-way check", () => {
  /** The workbook after "an edit in Excel": Rate set to 0.3 behind the project's back. */
  const excelEdited = () => applyChangeSet(F7, [{ op: "set-name", name: "Rate", scope: null, stored: "0.3", display: "0.3", comment: null, hidden: false, fields: ["definition"] }]);

  it("a name changed in Excel only: kept, reported, nothing to build", () => {
    const r = build(excelEdited(), pulled(F7));
    expect(r.status).toBe("up-to-date");
    expect(r.plan.excelChanges).toMatchObject([{ kind: "changed", key: "Rate" }]);
  });

  it("changed in Excel and in source: a conflict with both versions, never merged", () => {
    const files = pulled(F7);
    edit(files, U, "Rate = 0.1;", "Rate = 0.2;");
    const r = build(excelEdited(), files);
    expect(r.status).toBe("refused");
    expect(r.bytes).toBeUndefined();
    expect(r.plan.conflicts).toMatchObject([{ kind: "both-changed", key: "Rate", excel: { display: "0.3" }, source: { display: "0.2", file: U } }]);
  });

  it("changed the same way on both sides: no conflict", () => {
    const files = pulled(F7);
    edit(files, U, "Rate = 0.1;", "Rate = 0.3;");
    expect(build(excelEdited(), files).status).toBe("up-to-date");
  });

  it("deleted in source, changed in Excel: a conflict", () => {
    const files = pulled(F7);
    edit(files, U, "Rate2 = Rate*2;\n", "");
    const xl = applyChangeSet(F7, [{ op: "set-name", name: "Rate2", scope: null, stored: "Rate*3", display: "Rate*3", comment: null, hidden: false, fields: ["definition"] }]);
    expect(build(xl, files).plan.conflicts).toMatchObject([{ kind: "deleted-in-source", key: "Rate2" }]);
  });

  it("changed in source, deleted in Excel: a conflict", () => {
    const files = pulled(MAC);
    edit(files, U, "Fact = LAMBDA(n, 1);", "Fact = LAMBDA(n, 2);");
    const xl = applyChangeSet(MAC, [{ op: "delete-name", name: "Fact", scope: null }]);
    expect(build(xl, files).plan.conflicts).toMatchObject([{ kind: "deleted-in-excel", key: "Fact" }]);
  });

  it("a name created in Excel is kept and stays out of the lockfile", () => {
    const xl = applyChangeSet(F7, [{ op: "set-name", name: "NewInExcel", scope: null, stored: "1", display: "1", comment: null, hidden: false, fields: ["created"] }]);
    const files = pulled(F7);
    edit(files, U, "Rate = 0.1;", "Rate = 0.2;");
    const r = built(xl, files);
    expect(r.plan.excelChanges).toMatchObject([{ kind: "created", key: "NewInExcel" }]);
    expect(find(r.bytes!, "NewInExcel")!.definition).toBe("1");
    expect(Object.keys(parseLockfile(r.files![LOCK_FILE]!).names)).not.toContain("NewInExcel");
  });

  it("the lockfile after a build records what was built", () => {
    const files = pulled(F7);
    edit(files, U, "Rate = 0.1;", "Rate = 0.2;");
    const r = built(F7, files);
    expect(parseLockfile(r.files![LOCK_FILE]!).names["Rate"]!.definition).toBe(definitionHash("0.2"));
  });
});

describe("workbook part edge cases", () => {
  it("a workbook without <definedNames>: the element is inserted before <calcPr>", () => {
    const all = names(F7).map((d) => ({ op: "delete-name" as const, name: d.name, scope: d.scope.kind === "sheet" ? d.scope.name : null }));
    const bare = applyChangeSet(F7, all);
    expect(definedNamesXml(bare)).toBe("");
    const files = pulled(bare);
    files[U] = "Alpha = 1;\n";
    const r = built(bare, files);
    const xml = workbookXml(r.bytes!);
    expect(xml).toMatch(/<\/sheets><definedNames><definedName name="Alpha">1<\/definedName><\/definedNames><calcPr [^>]*fullCalcOnLoad="1"\/>/);
  });

  it("Excel's own _xl names are kept untouched, wherever they are", () => {
    const xml = workbookXml(F7).replace("<definedNames>", '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">\'S1\'!$A$1:$C$9</definedName>');
    const withPrint = withWorkbookXml(F7, xml);
    const files = pulled(withPrint);
    edit(files, U, "Rate = 0.1;", "Rate = 0.2;");
    const r = built(withPrint, files);
    expect(definedNamesXml(r.bytes!)).toContain('<definedName name="_xlnm.Print_Area" localSheetId="0">\'S1\'!$A$1:$C$9</definedName>');
    expect(r.plan.problems.filter((p) => p.severity === "error")).toEqual([]);
  });

  it("a source name spelled _xl… is refused", () => {
    const files = pulled(F7);
    files[U] += "\n_xlnm.Print_Titles = 1;\n";
    expect(build(F7, files).plan.problems.map((p) => p.code)).toContain("invalid-name");
  });

  it("other zip entries are copied byte for byte", () => {
    const files = pulled(F7);
    edit(files, U, "Rate = 0.1;", "Rate = 0.2;");
    const r = built(F7, files);
    const a = rawZipRecords(F7);
    const b = rawZipRecords(r.bytes!);
    for (const [k, v] of a) if (k !== "xl/workbook.xml") expect(Buffer.compare(Buffer.from(v), Buffer.from(b.get(k)!)), k).toBe(0);
  });

  it("leaves its input untouched, also a Node Buffer (whose slice is a view)", () => {
    const input = Buffer.from(F7);
    const files = pulled(F7);
    edit(files, U, "Rate = 0.1;", "Rate = 0.2;");
    const r = build(input, files);
    expect(r.status, why(r)).toBe("built");
    expect(Buffer.compare(input, Buffer.from(F7))).toBe(0);
  });

  it("scanWorkbookXml finds sheets, names and calcPr with offsets", () => {
    const xml = workbookXml(F7);
    const l = scanWorkbookXml(xml);
    expect(l.sheets).toEqual(["S1", "S2"]);
    expect(l.definedNames!.items.map((i) => i.attrs["name"])).toEqual(["Fn", "Loc", "Loc", "Rate", "Rate2", "RateX", "Spl"]);
    expect(xml.slice(l.calcPr!.start, l.calcPr!.end)).toMatch(/^<calcPr [^>]*\/>$/);
  });
});
