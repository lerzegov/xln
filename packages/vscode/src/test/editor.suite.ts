// M3c: writing help end to end, in the web extension host and in desktop VS Code, on the
// project the browse suite pulled (book/f7_base.xln): completion, signature help, checks
// as you type (following unsaved edits) with their quick fixes, and xln: New module.
import * as vscode from "vscode";

function check(cond: unknown, msg: string): asserts cond {
  if (!cond) throw new Error(msg);
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

async function until<T>(get: () => T | undefined, what: string, ms = 5000): Promise<T> {
  for (let t = 0; t < ms; t += 100) {
    const v = get();
    if (v !== undefined && v !== false) return v;
    await sleep(100);
  }
  throw new Error(`timed out: ${what}`);
}

export function defineEditorSuite(): void {
  suite("xln editor (M3c)", () => {
    const folder = () => vscode.workspace.workspaceFolders![0]!.uri;
    const project = () => vscode.Uri.joinPath(folder(), "book", "f7_base.xln");
    const scratch = () => vscode.Uri.joinPath(project(), "names", "m3c.xln");
    const live = (uri: vscode.Uri) => vscode.languages.getDiagnostics(uri).filter((d) => d.source === "xln");

    function at(doc: vscode.TextDocument, needle: string, delta: number): vscode.Position {
      const i = doc.getText().indexOf(needle);
      check(i >= 0, `${needle} not in ${doc.uri.path}`);
      return doc.positionAt(i + delta);
    }

    suiteSetup(async () => {
      await vscode.workspace.fs.writeFile(scratch(), new TextEncoder().encode("M3c = Rat + Fn(1, 2) + LET(k, 2, k);\nM3d = Ra;\n"));
      await vscode.commands.executeCommand("xln.reload");
    });

    suiteTeardown(async () => {
      await vscode.commands.executeCommand("workbench.action.closeAllEditors");
      for (const f of [scratch(), vscode.Uri.joinPath(project(), "names", "FIN.xln")]) {
        try {
          await vscode.workspace.fs.delete(f);
        } catch {
          // not written
        }
      }
      await vscode.commands.executeCommand("xln.reload");
    });

    test("completion lists names in scope, other sheets' local names qualified, and functions", async () => {
      const doc = await vscode.workspace.openTextDocument(scratch());
      const list = await vscode.commands.executeCommand<vscode.CompletionList>("vscode.executeCompletionItemProvider", doc.uri, at(doc, "M3d = Ra", 8));
      const labels = list.items.map((i) => (typeof i.label === "string" ? i.label : i.label.label));
      for (const l of ["Rate", "Rate2", "Fn", "'S2'!Loc", "SUM"]) check(labels.includes(l), `no ${l} in ${labels.slice(0, 20).join(", ")}`);
      const rate = list.items.find((i) => (typeof i.label === "string" ? i.label : i.label.label) === "Rate")!;
      check(rate.kind === vscode.CompletionItemKind.Constant, `Rate kind ${rate.kind}`);
      // Nothing on the left-hand side.
      const lhs = await vscode.commands.executeCommand<vscode.CompletionList>("vscode.executeCompletionItemProvider", doc.uri, at(doc, "M3d", 2));
      check(!lhs.items.some((i) => (typeof i.label === "string" ? i.label : i.label.label) === "'S2'!Loc"), "completion on the left-hand side");
    });

    test("signature help for a project LAMBDA", async () => {
      const doc = await vscode.workspace.openTextDocument(scratch());
      const help = await vscode.commands.executeCommand<vscode.SignatureHelp>("vscode.executeSignatureHelpProvider", doc.uri, at(doc, "Fn(1, 2)", 6), ",");
      check(help && help.signatures[0]?.label.startsWith("Fn("), `signature ${help?.signatures[0]?.label}`);
      check(help.activeParameter === 1 || help.signatures[0]!.parameters.length === 1, `active ${help.activeParameter}`);
    });

    test("checks as you type: diagnostics, quick fixes, and unsaved edits", async () => {
      const doc = await vscode.workspace.openTextDocument(scratch());
      await vscode.window.showTextDocument(doc);
      const first = await until(() => {
        const d = live(doc.uri);
        return d.some((x) => x.code === "C4.unknown-name") && d.some((x) => x.code === "C6.lambda-arity") ? d : undefined;
      }, "live diagnostics");
      const rat = first.find((x) => x.code === "C4.unknown-name" && doc.getText(x.range) === "Rat");
      check(rat, `diagnostics: ${first.map((x) => `${x.code} ${doc.getText(x.range)}`).join(", ")}`);
      const actions = await vscode.commands.executeCommand<vscode.CodeAction[]>("vscode.executeCodeActionProvider", doc.uri, rat.range);
      const fix = actions.find((a) => a.title === "Did you mean Rate?");
      check(fix?.edit, `actions: ${actions.map((a) => a.title).join(", ")}`);
      await vscode.workspace.applyEdit(fix.edit);
      // The edit is unsaved: the checks follow the editor's text.
      await until(() => !live(doc.uri).some((x) => doc.getText(x.range) === "Rat" || doc.getText(x.range) === "Rate") || undefined, "the fixed name's diagnostic to go");
      check(doc.isDirty, "the fix saved the file");
      await vscode.commands.executeCommand("workbench.action.files.revert");
    });

    test("a spent @renamed: a faded hint, and the quick fix that removes its line", async () => {
      // The lockfile has Rate and no OldRate: the note's rename is built.
      const doc = await vscode.workspace.openTextDocument(vscode.Uri.joinPath(project(), "names", "_unmanaged.xln"));
      await vscode.window.showTextDocument(doc);
      const before = doc.getText();
      const i = before.startsWith("Rate = ") ? 0 : before.indexOf("\nRate = ") + 1;
      check(i > 0 || before.startsWith("Rate = "), `no Rate statement in ${before.slice(0, 200)}`);
      const edit = new vscode.WorkspaceEdit();
      edit.insert(doc.uri, doc.positionAt(i), "@renamed(OldRate)\n");
      await vscode.workspace.applyEdit(edit);
      try {
        const d = await until(() => live(doc.uri).find((x) => x.code === "renamed-built"), "the renamed-built hint");
        check(d.message === "@renamed(OldRate): the rename is built; this line can go", `message: ${d.message}`);
        check(d.severity === vscode.DiagnosticSeverity.Hint, `severity ${d.severity}`);
        check(d.tags?.includes(vscode.DiagnosticTag.Unnecessary), "not faded");
        check(doc.getText(d.range) === "@renamed(OldRate)", `range: ${doc.getText(d.range)}`);
        const actions = await vscode.commands.executeCommand<vscode.CodeAction[]>("vscode.executeCodeActionProvider", doc.uri, d.range);
        const fix = actions.find((a) => a.title === "Remove @renamed(OldRate)");
        check(fix?.edit, `actions: ${actions.map((a) => a.title).join(", ")}`);
        await vscode.workspace.applyEdit(fix.edit);
        check(doc.getText() === before, "the fix did not restore the text");
        await until(() => !live(doc.uri).some((x) => x.code === "renamed-built") || undefined, "the hint to go");
      } finally {
        await vscode.commands.executeCommand("workbench.action.files.revert");
      }
    });

    test("xln: New module writes names/<Prefix>.xln with a documented sample LAMBDA", async () => {
      const uri = await vscode.commands.executeCommand<string | undefined>("xln.newModule", { root: project().toString(), prefix: "FIN" });
      check(uri && uri.endsWith("/names/FIN.xln"), `new module at ${uri}`);
      const doc = await vscode.workspace.openTextDocument(vscode.Uri.parse(uri));
      check(doc.getText().includes("FIN.GROW = LAMBDA(value, rate, [periods],"), "no sample LAMBDA");
      check(doc.getText().includes("@param rate"), "no @param");
      const hover = await vscode.commands.executeCommand<vscode.Hover[]>("vscode.executeHoverProvider", doc.uri, at(doc, "FIN.GROW =", 2));
      const md = hover.flatMap((h) => h.contents).map((c) => (typeof c === "string" ? c : "value" in c ? c.value : "")).join("\n");
      check(md.includes("`rate`"), `hover: ${md.slice(0, 200)}`);
      check(live(doc.uri).length === 0, `diagnostics on the new module: ${live(doc.uri).map((d) => d.message).join(" | ")}`);
      // A second module with that prefix is refused.
      const again = await vscode.commands.executeCommand<string | undefined>("xln.newModule", { root: project().toString(), prefix: "FIN" });
      check(again === undefined, "a second FIN module");
    });
  });
}
