// W5: the Name Manager features end to end, in the web extension host (@vscode/test-web)
// and in desktop VS Code (@vscode/test-electron). Pulls `book/f7_base.xlsx` of the
// fixture folder (scripts/fixture.mjs), then drives the features through VS Code's own
// commands, as the editor would.
import * as vscode from "vscode";
import type { PullOutcome } from "../commands.js";
import type { LoadStats } from "../xlnWorkspace.js";

function check(cond: unknown, msg: string): asserts cond {
  if (!cond) throw new Error(msg);
}

export function defineBrowseSuite(): void {
  suite("xln browse", () => {
    const folder = () => vscode.workspace.workspaceFolders![0]!.uri;
    const project = () => vscode.Uri.joinPath(folder(), "book", "f7_base.xln");
    const unmanaged = () => vscode.Uri.joinPath(project(), "names", "_unmanaged.xln");

    async function open(uri: vscode.Uri): Promise<vscode.TextDocument> {
      return vscode.workspace.openTextDocument(uri);
    }
    function pos(doc: vscode.TextDocument, needle: string, delta = 1): vscode.Position {
      const i = doc.getText().indexOf(needle);
      check(i >= 0, `${needle} not in ${doc.uri.path}`);
      return doc.positionAt(i + delta);
    }

    test("pulls a workbook into <name>.xln beside it, and replaces it on request", async () => {
      const wb = vscode.Uri.joinPath(folder(), "book", "f7_base.xlsx");
      const r = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", wb, { replace: true });
      check(r, "no pull outcome");
      check(r.names === 7, `names = ${r.names}`);
      for (const f of ["names/_unmanaged.xln", "names/sheets/S2.xln", "workbook.manifest.json", "xln.lock.json"]) check(r.written.includes(f), `${f} not written`);
      const text = new TextDecoder().decode(await vscode.workspace.fs.readFile(unmanaged()));
      check(text.includes("Rate2 = Rate*2;"), "unexpected _unmanaged.xln");
      // A second pull replaces the folder: a stray file in it disappears.
      const stray = vscode.Uri.joinPath(project(), "names", "stray.xln");
      await vscode.workspace.fs.writeFile(stray, new TextEncoder().encode("Stray = 1;\n"));
      const again = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", wb, { replace: true });
      check(again, "no second pull");
      const left = (await vscode.workspace.fs.readDirectory(vscode.Uri.joinPath(project(), "names"))).map(([n]) => n);
      check(!left.includes("stray.xln"), `stray.xln survived: ${left.join(", ")}`);
      // A source edit not built (the stray name): cancelling the question writes nothing and keeps it.
      await vscode.workspace.fs.writeFile(stray, new TextEncoder().encode("Stray = 1;\n"));
      const keep = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", wb, { unbuilt: "cancel" });
      check(keep === undefined, "a cancelled pull wrote");
      const still = (await vscode.workspace.fs.readDirectory(vscode.Uri.joinPath(project(), "names"))).map(([n]) => n);
      check(still.includes("stray.xln"), "a cancelled pull removed the edit");
      await vscode.workspace.fs.delete(stray);
      // No edit not built: the pull asks nothing; the workbook says the same, nothing written.
      const plain = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", wb);
      check(plain?.written.length === 0 && plain.discarded.length === 0, `plain pull: ${JSON.stringify(plain?.written)}`);
    });

    test("a pull closes the tabs of the files it removed (M3e)", async () => {
      const wb = vscode.Uri.joinPath(folder(), "book", "f7_base.xlsx");
      const notes = vscode.Uri.joinPath(project(), "names", "Notes.xln");
      await vscode.workspace.fs.writeFile(notes, new TextEncoder().encode("// notes, no names\n"));
      await vscode.window.showTextDocument(await open(notes), { preview: false });
      const tabOpen = () => vscode.window.tabGroups.all.some((g) => g.tabs.some((t) => t.input instanceof vscode.TabInputText && t.input.uri.toString() === notes.toString()));
      check(tabOpen(), "no tab for Notes.xln");
      const r = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", wb, { replace: true });
      check(r?.tabs.closed.includes(notes.toString()), `tabs ${JSON.stringify(r?.tabs)}`);
      check(!tabOpen(), "the tab of the removed file is still open");
    });

    test("a pull over a // comment alone asks (Pull anyway / Cancel), not refuses (M3e)", async () => {
      const wb = vscode.Uri.joinPath(folder(), "book", "f7_base.xlsx");
      const text = new TextDecoder().decode(await vscode.workspace.fs.readFile(unmanaged()));
      await vscode.workspace.fs.writeFile(unmanaged(), new TextEncoder().encode("// my note\n" + text));
      const cancelled = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", wb, { layout: "cancel" });
      check(cancelled === undefined, "a cancelled pull went on");
      check(new TextDecoder().decode(await vscode.workspace.fs.readFile(unmanaged())).startsWith("// my note\n"), "a cancelled pull rewrote the file");
      const r = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", wb, { layout: "pull" });
      check(r && r.discarded.length === 0 && r.rewritten.join() === "names/_unmanaged.xln", `pull ${JSON.stringify(r?.rewritten)}`);
      check(new TextDecoder().decode(await vscode.workspace.fs.readFile(unmanaged())) === text, "the comment is still there");
    });

    test("internals hidden and read-only in the editor; xln still writes them; Show / Hide project internals (M3e)", async () => {
      const files = vscode.workspace.getConfiguration("files");
      const lock = "**/*.xln/xln.lock.json";
      check(files.get<Record<string, boolean>>("exclude")?.[lock] === true, "the lockfile is not hidden by default");
      check(files.get<Record<string, boolean>>("readonlyInclude")?.[lock] === true, "the lockfile is not read-only by default");
      // The pull above wrote the lockfile and the manifest through workspace.fs; so does a pull now.
      const wb = vscode.Uri.joinPath(folder(), "book", "f7_base.xlsx");
      const lockUri = vscode.Uri.joinPath(project(), "xln.lock.json");
      const before = new TextDecoder().decode(await vscode.workspace.fs.readFile(lockUri));
      await vscode.workspace.fs.writeFile(lockUri, new TextEncoder().encode(before.replace("{", "{ ")));
      const r = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", wb, { replace: true });
      check(r?.written.includes("xln.lock.json"), `lockfile not rewritten: ${JSON.stringify(r?.written)}`);
      check(new TextDecoder().decode(await vscode.workspace.fs.readFile(lockUri)) === before, "lockfile differs after the pull");
      const shown = await vscode.commands.executeCommand<boolean>("xln.showInternals");
      try {
        check(shown === false, "still hidden after Show project internals");
        check(vscode.workspace.getConfiguration("files").inspect<Record<string, boolean>>("exclude")?.workspaceValue?.[lock] === false, "no workspace setting");
      } finally {
        const hidden = await vscode.commands.executeCommand<boolean>("xln.hideInternals");
        check(hidden === true, "not hidden after Hide project internals");
      }
      check(vscode.workspace.getConfiguration("files").inspect<Record<string, boolean>>("exclude")?.workspaceValue?.[lock] === undefined, "the workspace setting stayed");
    });

    test("loads the project", async () => {
      const s = await vscode.commands.executeCommand<LoadStats>("xln.reload");
      check(s.projects === 1 && s.names === 7, `stats ${JSON.stringify(s)}`);
    });

    test("outline: names; a sheet file has no blocks (M3d)", async () => {
      const doc = await open(unmanaged());
      const syms = await vscode.commands.executeCommand<vscode.DocumentSymbol[]>("vscode.executeDocumentSymbolProvider", doc.uri);
      check(syms.map((s) => s.name).join() === "Fn,Loc,Rate,Rate2,RateX", `symbols ${syms.map((s) => s.name).join()}`);
      check(syms[0]!.kind === vscode.SymbolKind.Function, "Fn is not a function symbol");
      const s2 = await open(vscode.Uri.joinPath(project(), "names", "sheets", "S2.xln"));
      const blocks = await vscode.commands.executeCommand<vscode.DocumentSymbol[]>("vscode.executeDocumentSymbolProvider", s2.uri);
      check(blocks.some((b) => b.name === "Loc") && !blocks.some((b) => b.name.startsWith("@scope")), "S2 outline");
    });

    test("go to definition of a name used in a definition", async () => {
      const doc = await open(unmanaged());
      const defs = await vscode.commands.executeCommand<(vscode.Location | vscode.LocationLink)[]>("vscode.executeDefinitionProvider", doc.uri, pos(doc, "Rate*2"));
      check(defs.length === 1, `definitions: ${defs.length}`);
      const d = defs[0]!;
      const range = "targetRange" in d ? d.targetSelectionRange ?? d.targetRange : d.range;
      check(range.start.line === pos(doc, "Rate = 0.1", 0).line, `went to line ${range.start.line}`);
      // A LAMBDA parameter goes to its declaration, not to a name.
      const local = await vscode.commands.executeCommand<vscode.Location[]>("vscode.executeDefinitionProvider", doc.uri, pos(doc, "x*Rate"));
      check(local.length === 1 && doc.getText(local[0]!.range) === "x", "parameter definition");
    });

    test("references: in names and in cells (xln-cells:)", async () => {
      const doc = await open(unmanaged());
      const refs = await vscode.commands.executeCommand<vscode.Location[]>("vscode.executeReferenceProvider", doc.uri, pos(doc, "Rate = 0.1"));
      const inNames = refs.filter((r) => r.uri.toString() === doc.uri.toString());
      const inCells = refs.filter((r) => r.uri.scheme === "xln-cells");
      check(inNames.length === 3, `references in names: ${inNames.length}`); // declaration + Fn + Rate2
      check(inCells.length >= 5, `cell usages: ${inCells.length}`);
      const cells = await open(inCells[0]!.uri);
      const line = cells.lineAt(inCells[0]!.range.start.line).text;
      check(line === "'S1'!A1: =Rate*2  → 0.2", `first cell usage: ${line}`);
      check(cells.getText().includes("Conditional formats (1)"), "no conditional format usage");
    });

    test("used by / uses through the call hierarchy", async () => {
      const doc = await open(unmanaged());
      const items = await vscode.commands.executeCommand<vscode.CallHierarchyItem[]>("vscode.prepareCallHierarchy", doc.uri, pos(doc, "Rate = 0.1"));
      check(items.length === 1, "no hierarchy item");
      const incoming = await vscode.commands.executeCommand<vscode.CallHierarchyIncomingCall[]>("vscode.provideIncomingCalls", items[0]);
      check(incoming.map((c) => c.from.name).sort().join() === "Fn,Rate2,Spl", `used by ${incoming.map((c) => c.from.name).join()}`);
      const fn = await vscode.commands.executeCommand<vscode.CallHierarchyItem[]>("vscode.prepareCallHierarchy", doc.uri, pos(doc, "Fn ="));
      const outgoing = await vscode.commands.executeCommand<vscode.CallHierarchyOutgoingCall[]>("vscode.provideOutgoingCalls", fn[0]);
      check(outgoing.map((c) => c.to.name).join() === "Rate", "Fn uses");
    });

    test("hover: kind, scope, spill anchor, extent, first value", async () => {
      const doc = await open(unmanaged());
      // Spl names E1's spill: a cell statement in S1's file (M3b).
      const s1 = await open(vscode.Uri.joinPath(project(), "names", "sheets", "S1.xln"));
      const text = async (needle: string, d = doc) => {
        const hovers = await vscode.commands.executeCommand<vscode.Hover[]>("vscode.executeHoverProvider", d.uri, pos(d, needle));
        return hovers.flatMap((h) => h.contents.map((c) => (typeof c === "string" ? c : c.value))).join("\n");
      };
      const spl = await text("Spl @", s1);
      check(spl.includes("last saved extent: `E1:E3`"), `Spl hover: ${spl}`);
      check(spl.includes("first value (cached)"), `Spl hover: ${spl}`);
      const fn = await text("Fn =");
      check(fn.includes("LAMBDA · workbook scope · arity 1"), `Fn hover: ${fn}`);
    });

    test("search: names, definitions, doc comments; workspace symbols", async () => {
      const keys = await vscode.commands.executeCommand<string[]>("xln.searchNames", "lambda rate");
      check(keys?.join() === "Fn", `search: ${keys}`);
      const syms = await vscode.commands.executeCommand<vscode.SymbolInformation[]>("vscode.executeWorkspaceSymbolProvider", "RateX");
      check(syms.some((s) => s.name === "RateX"), "RateX not among workspace symbols");
    });

    test("formula view: a sheet's formulas, names linked, go to cell", async () => {
      const wb = vscode.Uri.joinPath(folder(), "book", "f7_base.xlsx");
      const opened = await vscode.commands.executeCommand<string | undefined>("xln.formulaView", { workbook: wb.toString(), sheet: "S1", cell: "B3" });
      check(opened, "no formula view");
      const view = await open(vscode.Uri.parse(opened));
      check(view.uri.scheme === "xln-formulas" && view.languageId === "xln-formulas", `view ${view.uri.toString()} in ${view.languageId}`);
      check(vscode.window.activeTextEditor?.selection.active.line === pos(view, "B3    shared", 0).line, "not opened at B3");
      const line = view.lineAt(pos(view, "B3    shared", 0).line).text;
      check(line === "     B3    shared ← B2  = A3*Rate               → 0.2", `B3 line: ${line}`);
      // A name goes to its entry in the project; a reference to the line of its cell.
      const defs = await vscode.commands.executeCommand<vscode.Location[]>("vscode.executeDefinitionProvider", view.uri, pos(view, "A3*Rate", 4));
      check(defs.length === 1 && defs[0]!.uri.toString() === unmanaged().toString(), `Rate goes to ${defs.map((d) => d.uri.toString())}`);
      const names = await open(unmanaged());
      check(defs[0]!.range.start.line === pos(names, "Rate = 0.1", 0).line, "Rate: wrong line");
      const toCell = await vscode.commands.executeCommand<vscode.Location[]>("vscode.executeDefinitionProvider", view.uri, pos(view, "ROWS(E1#)", 5));
      check(toCell.length === 1 && view.getText(toCell[0]!.range) === "E1", "E1# does not go to E1");
      const hovers = await vscode.commands.executeCommand<vscode.Hover[]>("vscode.executeHoverProvider", view.uri, pos(view, "A3*Rate", 4));
      const hover = hovers.flatMap((h) => h.contents.map((c) => (typeof c === "string" ? c : c.value))).join("\n");
      check(hover.includes("constant · workbook scope"), `hover: ${hover}`);
      const syms = await vscode.commands.executeCommand<vscode.DocumentSymbol[]>("vscode.executeDocumentSymbolProvider", view.uri);
      check(syms.length === 22 && syms[2]!.name === "Spl — E1#" && syms[0]!.name === "A1", `symbols ${syms.map((s) => s.name).join()}`);
      // The name on the left of E1# (Spl = 'S1'!$E$1#) goes to its entry, hovers, and has references.
      const lhs = await vscode.commands.executeCommand<vscode.Location[]>("vscode.executeDefinitionProvider", view.uri, pos(view, "Spl  E1#", 1));
      const s1 = await open(vscode.Uri.joinPath(project(), "names", "sheets", "S1.xln"));
      check(lhs.length === 1 && lhs[0]!.uri.toString() === s1.uri.toString(), `Spl goes to ${lhs.map((d) => d.uri.toString())}`);
      check(lhs[0]!.range.start.line === pos(s1, "Spl @", 0).line, "Spl: wrong line");
      const splHover = await vscode.commands.executeCommand<vscode.Hover[]>("vscode.executeHoverProvider", view.uri, pos(view, "Spl  E1#", 1));
      const splText = splHover.flatMap((h) => h.contents.map((c) => (typeof c === "string" ? c : c.value))).join("\n");
      check(splText.includes("last saved extent: `E1:E3`"), `Spl hover: ${splText}`);
      const splRefs = await vscode.commands.executeCommand<vscode.Location[]>("vscode.executeReferenceProvider", view.uri, pos(view, "Spl  E1#", 1));
      check(splRefs.some((r) => r.uri.toString() === s1.uri.toString()), "references of Spl");
      const refs = await vscode.commands.executeCommand<vscode.Location[]>("vscode.executeReferenceProvider", view.uri, pos(view, "A3*Rate", 4));
      check(refs.some((r) => r.uri.scheme === "xln-cells") && refs.some((r) => r.uri.toString() === unmanaged().toString()), "references of Rate");
      // From a cell usage line, F12 on the address opens the view at that cell.
      const usage = refs.find((r) => r.uri.scheme === "xln-cells")!;
      const back = await vscode.commands.executeCommand<vscode.Location[]>("vscode.executeDefinitionProvider", usage.uri, usage.range.start.translate(0, 2));
      check(back.length === 1 && back[0]!.uri.scheme === "xln-formulas", "cell usage does not lead to the formula view");
      const target = await open(back[0]!.uri);
      check(target.getText(back[0]!.range) === "A1", `cell usage went to ${target.getText(back[0]!.range)}`);
      // From names/sheets/S2.xln: that sheet, without asking.
      const s2 = await vscode.commands.executeCommand<string | undefined>("xln.formulaView", vscode.Uri.joinPath(project(), "names", "sheets", "S2.xln"));
      check(s2 !== undefined && vscode.Uri.parse(s2).path.endsWith("/S2 (formulas)"), `S2 view: ${s2}`);
    });

    test("formula view in calculation order: levels, hover links, switching order, the workbook view", async () => {
      const wb = vscode.Uri.joinPath(folder(), "book", "f7_base.xlsx");
      const opened = await vscode.commands.executeCommand<string | undefined>("xln.formulaViewCalc", { workbook: wb.toString(), sheet: "S1", cell: "C6" });
      check(opened && vscode.Uri.parse(opened).path.endsWith("/S1 (calculation order)"), `calc view: ${opened}`);
      const view = await open(vscode.Uri.parse(opened));
      check(view.languageId === "xln-formulas", `language ${view.languageId}`);
      check(view.getText().includes("// Sheet S1: 22 formulas in calculation order"), "header");
      const c6 = pos(view, "C6    ", 0);
      check(view.lineAt(c6.line).text.startsWith("2       C6"), `C6 line: ${view.lineAt(c6.line).text}`);
      check(view.getText().indexOf("E1#") < view.offsetAt(c6), "E1 not before C6");
      check(vscode.window.activeTextEditor?.selection.active.line === c6.line, "not opened at C6");
      const hovers = await vscode.commands.executeCommand<vscode.Hover[]>("vscode.executeHoverProvider", view.uri, c6.translate(0, 1));
      const hover = hovers.flatMap((h) => h.contents.map((c) => (typeof c === "string" ? c : c.value))).join("\n");
      check(hover.includes("level 2") && hover.includes("reads 2: `Spl`, [`E1#`](command:xln.formulaView?") && hover.includes(") (via `Spl`)") && hover.includes("read by: nothing"), `hover: ${hover}`);
      const e1 = await vscode.commands.executeCommand<vscode.Hover[]>("vscode.executeHoverProvider", view.uri, pos(view, "E1#", 0).translate(0, 1));
      const e1Text = e1.flatMap((h) => h.contents.map((c) => (typeof c === "string" ? c : c.value))).join("\n");
      check(e1Text.includes("command:xln.formulaView?"), `E1 hover has no links: ${e1Text}`);
      // The editor title's button: the same sheet in order of appearance, at the same cell.
      const back = await vscode.commands.executeCommand<string | undefined>("xln.formulaViewToggleOrder");
      check(back && vscode.Uri.parse(back).path.endsWith("/S1 (formulas)"), `toggled: ${back}`);
      const plain = await open(vscode.Uri.parse(back));
      check(vscode.window.activeTextEditor?.selection.active.line === pos(plain, "C6    ", 0).line, "toggle lost the cell");
      // Every sheet: a reference goes to its line in the same document.
      const all = await vscode.commands.executeCommand<string | undefined>("xln.workbookFormulaView", { workbook: wb.toString() });
      check(all && vscode.Uri.parse(all).path.endsWith("/[workbook] (calculation order)"), `workbook view: ${all}`);
      const allDoc = await open(vscode.Uri.parse(all));
      check(allDoc.getText().includes("// Workbook f7_base.xlsx: 28 formulas on 2 sheets"), "workbook header");
      const toCell = await vscode.commands.executeCommand<vscode.Location[]>("vscode.executeDefinitionProvider", allDoc.uri, pos(allDoc, "ROWS(E1#)", 5));
      check(toCell.length === 1 && toCell[0]!.uri.toString() === allDoc.uri.toString() && allDoc.getText(toCell[0]!.range) === "E1", "E1# does not go to E1 in the workbook view");
      const syms = await vscode.commands.executeCommand<vscode.DocumentSymbol[]>("vscode.executeDocumentSymbolProvider", allDoc.uri);
      check(syms.length === 28 && syms.some((s) => s.name === "Spl — S1!E1#"), `symbols ${syms.map((s) => s.name).join()}`);
    });

    test("diagnostics from module and formula syntax", async () => {
      const bad = vscode.Uri.joinPath(project(), "names", "bad.xln");
      await vscode.workspace.fs.writeFile(bad, new TextEncoder().encode("bad.Bad = (1;\nbad.Worse 2;\n"));
      await vscode.commands.executeCommand("xln.reload");
      const diags = vscode.languages.getDiagnostics(bad);
      check(diags.length === 2, `diagnostics: ${diags.map((d) => d.message).join(" | ")}`);
      await vscode.workspace.fs.delete(bad);
      await vscode.commands.executeCommand("xln.reload");
      check(vscode.languages.getDiagnostics(bad).length === 0, "diagnostics not cleared");
    });
  });
}
