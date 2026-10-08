// probe_names.ts — P0 probe: the Name Manager through Office Scripts (no add-in).
//
// 1. Create a NEW blank workbook in OneDrive (Excel on the web, Windows or Mac).
// 2. Automate tab -> New Script -> replace everything with this file -> Run.
//    (No Automate tab at all = Office Scripts is disabled for you: that is result E01.)
// 3. Results appear on a sheet called ProbeResults and in the console. Copy them back
//    into probes/results/officescripts-<platform>.txt.
//
// Test IDs match probes/mac and probes/windows. Some tests have no Office Scripts
// equivalent (rename: NamedItem has no setName; evaluate: no Application.Evaluate) and
// are reported as N/A.

async function main(workbook: ExcelScript.Workbook) {
  const rpt: string[][] = [];
  const add = (tid: string, verdict: string, det: string) => { rpt.push([tid, verdict, det]); console.log(`${tid} | ${verdict} | ${det}`); };
  const err = (e: unknown) => (e instanceof Error ? e.message : String(e));
  const sheet = (n: string) => workbook.getWorksheet(n) ?? workbook.addWorksheet(n);
  const s1 = sheet("S1"), s2 = sheet("S2");
  const calc = () => workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  const cellText = (ws: ExcelScript.Worksheet, a: string) => ws.getRange(a).getText();

  add("E01", "PASS", "Office Scripts runs (Automate tab present)");
  add("T00", "INFO", `calculation mode ${workbook.getApplication().getCalculationMode()}`);

  // T01 workbook-scoped constant
  try { workbook.addNamedItem("P_K", "=10"); add("T01", "PASS", `created P_K; formula=${workbook.getNamedItem("P_K").getFormula()}`); }
  catch (e) { add("T01", "FAIL", err(e)); }

  // T02 LAMBDA name called from a cell; T04 comment set at creation
  try {
    workbook.addNamedItem("P_Add1", "=LAMBDA(x, x+1)", "probe comment");
    s1.getRange("A1").setFormula("=P_Add1(41)"); calc();
    const v = s1.getRange("A1").getValue();
    add("T02", v === 42 ? "PASS" : "FAIL", `=P_Add1(41) -> ${v}`);
  } catch (e) { add("T02", "FAIL", err(e)); }

  // T03 stored text
  try { add("T03", "INFO", `formula=${workbook.getNamedItem("P_Add1").getFormula()}`); } catch (e) { add("T03", "FAIL", err(e)); }

  // T04 comment read back, then changed
  try {
    const n = workbook.getNamedItem("P_Add1");
    const before = n.getComment(); n.setComment("probe comment 2");
    add("T04", "PASS", `comment at creation '${before}', after set '${n.getComment()}'`);
  } catch (e) { add("T04", "FAIL", err(e)); }

  // T05 sheet-scoped names
  try {
    s2.addNamedItem("P_Local", "=5");
    s2.getRange("A1").setFormula("=P_Local"); s1.getRange("A3").setFormula("=S2!P_Local"); calc();
    add("T05a", "INFO", `make at sheet: S2!A1=${cellText(s2, "A1")} S1!A3=${cellText(s1, "A3")}`);
  } catch (e) { add("T05a", "FAIL", err(e)); }
  try { add("T05c", "INFO", `names of S2: ${s2.getNames().map(n => `${n.getName()}(${n.getScope()})`).join("; ")}`); }
  catch (e) { add("T05c", "FAIL", err(e)); }

  // T06 name over a spilled range
  try {
    s1.getRange("B1").setFormula("=SEQUENCE(1,5)");
    workbook.addNamedItem("P_Spill", "=S1!$B$1#");
    s1.getRange("C2").setFormula("=COLUMNS(P_Spill)"); calc();
    add("T06", "INFO", `formula=${workbook.getNamedItem("P_Spill").getFormula()}  COLUMNS->${cellText(s1, "C2")}`);
  } catch (e) { add("T06", "FAIL", err(e)); }

  // T08 update
  try {
    s1.getRange("A4").setFormula("=P_K");
    workbook.getNamedItem("P_K").setFormula("=20"); calc();
    add("T08", "INFO", `after update S1!A4=${cellText(s1, "A4")} (expect 20)`);
  } catch (e) { add("T08", "FAIL", err(e)); }

  add("T09", "N/A", "NamedItem has no setName in Office Scripts (rename = delete + add)");

  // T10 delete
  try { workbook.getNamedItem("P_K").delete(); calc(); add("T10", "INFO", `deleted; S1!A4 shows ${cellText(s1, "A4")}`); }
  catch (e) { add("T10", "FAIL", err(e)); }

  // T11 syntax error
  try { workbook.addNamedItem("P_Bad", "=LAMBDA(x, x+"); add("T11", "INFO", "accepted (!)"); }
  catch (e) { add("T11", "INFO", `rejected with: ${err(e)}`); }

  // T12 collision with a built-in
  try {
    workbook.addNamedItem("Fact", "=LAMBDA(n, 1)");
    s1.getRange("A5").setFormula("=Fact(5)"); calc();
    add("T12", "INFO", `=Fact(5) -> ${cellText(s1, "A5")} (1 = name wins, 120 = built-in wins); formula reads ${s1.getRange("A5").getFormula()}`);
  } catch (e) { add("T12", "INFO", `rejected: ${err(e)}`); }

  // T13 dotted name
  try {
    workbook.addNamedItem("Mod.Fn", "=LAMBDA(x, x*10)");
    s1.getRange("A6").setFormula("=Mod.Fn(2)"); calc();
    add("T13", "INFO", `=Mod.Fn(2) -> ${cellText(s1, "A6")}`);
  } catch (e) { add("T13", "FAIL", err(e)); }

  // T14 line breaks
  try {
    workbook.addNamedItem("P_Multi", "=LAMBDA(x,\n  x*2)");
    const r = workbook.getNamedItem("P_Multi").getFormula();
    s1.getRange("A7").setFormula("=P_Multi(4)"); calc();
    add("T14", "INFO", `${r.includes("\n") ? "newline KEPT" : "newline LOST"}; =P_Multi(4) -> ${cellText(s1, "A7")}`);
  } catch (e) { add("T14", "FAIL", err(e)); }

  // T15 length limit
  for (const k of [3990, 4090, 4500]) {
    const f = "=LAMBDA(x,x" + "+1".repeat(k) + ")", nm = `P_Long${k}`;
    try {
      workbook.addNamedItem(nm, f);
      s1.getRange("A9").setFormula(`=${nm}(0)`); calc();
      add("T15", "INFO", `${f.length} chars accepted; ->${cellText(s1, "A9")}`);
    } catch (e) { add("T15", "INFO", `${f.length} chars rejected: ${err(e)}`); }
  }

  add("T16", "N/A", "no Application.Evaluate in Office Scripts (use a scratch cell)");

  // T17 non-ASCII name
  try {
    workbook.addNamedItem("Growλ", "=LAMBDA(b, g, b*(1+g))");
    s1.getRange("A10").setFormula("=Growλ(100, 0.1)"); calc();
    add("T17", "INFO", `=Growλ(100,0.1) -> ${cellText(s1, "A10")}`);
  } catch (e) { add("T17", "FAIL", err(e)); }

  // T18 local language
  try { workbook.addNamedItemFormulaLocal("P_Loc", "=LAMBDA(x; x+1)"); add("T18", "INFO", `local with ';' accepted; formula=${workbook.getNamedItem("P_Loc").getFormula()}`); }
  catch (e) { add("T18", "INFO", `local with ';' rejected: ${err(e)}`); }

  // T07 enumerate
  try {
    const all = workbook.getNames();
    add("T07", "INFO", `${all.length} workbook-scoped names`);
    for (const n of all) add("T07", "ITEM", `${n.getName()} | visible=${n.getVisible()} | type=${n.getType()} | comment=${n.getComment()} | ${n.getFormula().slice(0, 80)}`);
  } catch (e) { add("T07", "FAIL", err(e)); }

  // F01 can a script pull a module from the web? (raw GitHub sends CORS headers)
  try {
    const r = await fetch("https://raw.githubusercontent.com/microsoft/advanced-formula-environment/main/README.md");
    const t = await r.text();
    add("F01", r.ok ? "PASS" : "FAIL", `fetch raw.githubusercontent.com -> HTTP ${r.status}, ${t.length} chars`);
  } catch (e) { add("F01", "FAIL", `fetch: ${err(e)}`); }

  // write results to a sheet
  const out = workbook.getWorksheet("ProbeResults") ?? workbook.addWorksheet("ProbeResults");
  out.getRange().clear();
  out.getRangeByIndexes(0, 0, rpt.length, 3).setValues(rpt);
  out.activate();
}
