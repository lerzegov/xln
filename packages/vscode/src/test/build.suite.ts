// M3a: `xln: Build workbook` end to end, on the project the browse suite pulled
// (`book/f7_base.xln`). Desktop overwrites the workbook and keeps a backup; the browser
// writes `f7_base.xln.xlsx` beside it and leaves the original alone (E1).
import * as vscode from "vscode";
import { readWorkbook } from "@xln/core";
import type { BuildOutcome } from "../build.js";
import type { PullOutcome } from "../commands.js";
import type { XlnApi } from "../extension.js";

function check(cond: unknown, msg: string): asserts cond {
  if (!cond) throw new Error(msg);
}

export function defineBuildSuite(web: boolean): void {
  suite(`xln build (${web ? "web" : "desktop"})`, () => {
    const folder = () => vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, "book");
    const wb = () => vscode.Uri.joinPath(folder(), "f7_base.xlsx");
    const unmanaged = () => vscode.Uri.joinPath(folder(), "f7_base.xln", "names", "_unmanaged.xln");
    const read = async (u: vscode.Uri) => vscode.workspace.fs.readFile(u);
    const rate = (bytes: Uint8Array) => readWorkbook(bytes).definedNames.find((d) => d.name === "Rate")?.definition;
    const setRate = async (v: string) => {
      const t = new TextDecoder().decode(await read(unmanaged()));
      await vscode.workspace.fs.writeFile(unmanaged(), new TextEncoder().encode(t.replace(/Rate = [0-9.]+;/, `Rate = ${v};`)));
    };

    test("a fresh pull is up to date", async () => {
      const p = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", wb(), { replace: true });
      check(p, "no pull");
      const r = await vscode.commands.executeCommand<BuildOutcome | undefined>("xln.buildWorkbook", wb(), { confirm: false });
      check(r?.status === "up-to-date", `status ${r?.status}`);
    });

    test("builds an edited definition", async () => {
      const before = await read(wb());
      await setRate("0.2");
      const r = await vscode.commands.executeCommand<BuildOutcome | undefined>("xln.buildWorkbook", wb(), { confirm: false });
      check(r?.status === "built" && r.changes === 1, `outcome ${JSON.stringify(r)}`);
      if (web) {
        check(r.written!.endsWith("/book/f7_base.xln.xlsx"), `written ${r.written}`);
        check(rate(await read(vscode.Uri.parse(r.written!))) === "0.2", "copy lacks the change");
        check(rate(await read(wb())) === "0.1", "the browser overwrote the workbook");
      } else {
        check(r.written === wb().toString(), `written ${r.written}`);
        check(rate(await read(wb())) === "0.2", "workbook lacks the change");
        const backup = await read(vscode.Uri.joinPath(folder(), "f7_base.backup.xlsx"));
        check(backup.length === before.length && backup.every((x, i) => x === before[i]), "backup differs from the original");
        const again = await vscode.commands.executeCommand<BuildOutcome | undefined>("xln.buildWorkbook", wb(), { confirm: false });
        check(again?.status === "up-to-date", `rebuild ${again?.status}`);
      }
    });

    if (!web) {
      // With Excel's owner file the build offers *Close in Excel and build* (a modal, answered
      // here: declined). Accepting it would drive the real Excel, which a test must not.
      test("while Excel's owner file exists, Close in Excel and build declined writes nothing", async () => {
        await setRate("0.3");
        const lock = vscode.Uri.joinPath(folder(), "~$f7_base.xlsx");
        await vscode.workspace.fs.writeFile(lock, new Uint8Array());
        try {
          const r = await vscode.commands.executeCommand<BuildOutcome | undefined>("xln.buildWorkbook", wb(), { confirm: false, closeInExcel: false });
          check(r?.status === "locked", `status ${r?.status}`);
          check(r.lines.some((l) => l.toLowerCase().includes("nothing written")), `lines ${r.lines.join(" | ")}`);
          check(rate(await read(wb())) === "0.2", "written while locked");
        } finally {
          await vscode.workspace.fs.delete(lock);
        }
      });
    }

    test("a pull over an edit not built: Build first builds, then pulls (desktop); the browser's copy leaves the edit, so the pull stops", async () => {
      await vscode.commands.executeCommand("xln.pullWorkbook", wb(), { replace: true });
      const original = await read(wb());
      await setRate("0.4");
      const p = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", wb(), { unbuilt: "build", build: { confirm: false } });
      const text = new TextDecoder().decode(await read(unmanaged()));
      if (web) {
        check(p === undefined, "the pull went on although the workbook lacks the edit");
        check(text.includes("Rate = 0.4;"), "the edit was replaced");
      } else {
        check(p !== undefined && p.discarded.length === 0, `pull ${JSON.stringify(p)}`);
        check(rate(await read(wb())) === "0.4" && text.includes("Rate = 0.4;"), "the build did not come first");
      }
      await vscode.workspace.fs.writeFile(wb(), original);
      await vscode.commands.executeCommand("xln.pullWorkbook", wb(), { replace: true });
    });

    test("Rename Symbol, then Build: the build removes @renamed (desktop) and a pull is not refused; the browser's copy leaves it", async () => {
      await vscode.commands.executeCommand("xln.pullWorkbook", wb(), { replace: true });
      const original = await read(wb());
      try {
        const doc = await vscode.workspace.openTextDocument(unmanaged());
        await vscode.window.showTextDocument(doc);
        const at = doc.getText().indexOf("RateX = 0.5;");
        check(at >= 0, "RateX not in _unmanaged.xln");
        const edit = await vscode.commands.executeCommand<vscode.WorkspaceEdit>("vscode.executeDocumentRenameProvider", doc.uri, doc.positionAt(at + 1), "RateY");
        check(edit && edit.size > 0, "no rename edit");
        check(await vscode.workspace.applyEdit(edit), "rename edit not applied");
        check(doc.getText().includes("@renamed(RateX)\nRateY = 0.5;"), `after F2: ${doc.getText()}`);
        const r = await vscode.commands.executeCommand<BuildOutcome | undefined>("xln.buildWorkbook", wb(), { confirm: false });
        check(r?.status === "built", `outcome ${JSON.stringify(r)}`);
        const disk = new TextDecoder().decode(await read(unmanaged()));
        if (web) {
          check(r.renamedRemoved === undefined && doc.getText().includes("@renamed(RateX)"), "the browser's build changed the source");
          return;
        }
        check(JSON.stringify(r.renamedRemoved) === JSON.stringify(["removed @renamed(RateX) from names/_unmanaged.xln: the rename is built"]), `removed ${JSON.stringify(r.renamedRemoved)}`);
        check(r.lines.some((l) => l === "  removed @renamed(RateX) from names/_unmanaged.xln: the rename is built"), `lines ${r.lines.join(" | ")}`);
        check(!doc.getText().includes("@renamed") && doc.getText().includes("\nRateY = 0.5;"), `editor: ${doc.getText()}`);
        check(!doc.isDirty && disk === doc.getText(), "the editor's text is not saved");
        // Fully built: the pull finds nothing unbuilt and no file to rewrite (both answers here would cancel it).
        const p = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", wb(), { unbuilt: "cancel", layout: "cancel" });
        check(p !== undefined && p.discarded.length === 0 && p.rewritten.length === 0, `pull ${JSON.stringify(p)}`);
        check(new TextDecoder().decode(await read(unmanaged())) === disk, "the pull rewrote _unmanaged.xln");
      } finally {
        await vscode.workspace.fs.writeFile(wb(), original);
        await vscode.commands.executeCommand("xln.pullWorkbook", wb(), { replace: true });
        await vscode.commands.executeCommand("workbench.action.closeAllEditors");
      }
    });

    test("the xln output channel logs F2, the build, the audit and the formula view: a timed header, then details", async () => {
      await vscode.commands.executeCommand("xln.pullWorkbook", wb(), { replace: true });
      const original = await read(wb());
      const api = vscode.extensions.all.find((e) => e.packageJSON?.name === "xln")!.exports as XlnApi;
      const from = api.logLines().length;
      const head = (s: string) => new RegExp(`^\\d\\d:\\d\\d:\\d\\d xln ${s}`);
      // The test before renamed RateX in S1.xln too and restored the files on disk: VS Code
      // catches its (closed) document up with the disk on its own time; wait for it, then reload.
      const s1 = vscode.Uri.joinPath(folder(), "f7_base.xln", "names", "sheets", "S1.xln");
      const onDisk = new TextDecoder().decode(await read(s1));
      for (let k = 0; k < 50 && vscode.workspace.textDocuments.some((d) => d.uri.toString() === s1.toString() && d.getText() !== onDisk); k++) await new Promise((r) => setTimeout(r, 100));
      await vscode.commands.executeCommand("xln.reload");
      try {
        const doc = await vscode.workspace.openTextDocument(unmanaged());
        await vscode.window.showTextDocument(doc);
        const at = doc.getText().indexOf("RateX = 0.5;");
        // F2 on a non-name first: refused, and the reason is logged.
        await vscode.commands.executeCommand("vscode.executeDocumentRenameProvider", doc.uri, doc.positionAt(at + "RateX = 0".length), "Nope").then(
          () => undefined,
          () => undefined,
        );
        const edit = await vscode.commands.executeCommand<vscode.WorkspaceEdit>("vscode.executeDocumentRenameProvider", doc.uri, doc.positionAt(at + 1), "RateY");
        check(edit && (await vscode.workspace.applyEdit(edit)), "rename edit not applied");
        const r = await vscode.commands.executeCommand<BuildOutcome | undefined>("xln.buildWorkbook", wb(), { confirm: false });
        check(r?.status === "built", `outcome ${JSON.stringify(r)}`);
        check(await vscode.commands.executeCommand<string | undefined>("xln.auditWorkbook", wb()), "no audit report");
        check(await vscode.commands.executeCommand<string | undefined>("xln.formulaView", { workbook: wb().toString(), sheet: "S1" }), "no formula view");
        const log = api.logLines().slice(from);
        const after = (re: RegExp) => {
          const k = log.findIndex((l) => re.test(l));
          check(k >= 0, `no ${re} in the log:\n${log.join("\n")}`);
          check(log[k - 1] === "", `no blank line before ${log[k]}`);
          return log.slice(k + 1, log.findIndex((l, j) => j > k && l === "") >>> 0);
        };
        const refused = after(head("rename at names/_unmanaged\\.xln:\\d+$"));
        check(refused[0]?.startsWith("  refused: "), `refusal ${refused.join(" | ")}`);
        const renamed = after(head("rename RateX → RateY$"));
        check(/^ {2}book\/f7_base\.xln: \d+ edit\(s\) in \d+ file\(s\): names\/_unmanaged\.xln/.test(renamed[0] ?? ""), `rename ${renamed.join(" | ")}`);
        check(renamed.includes("  @renamed(RateX) added: the build renames the name in the workbook and rewrites it in the cells"), `rename ${renamed.join(" | ")}`);
        const built = after(head("build f7_base\\.xlsx: 1 change$"));
        check(built.some((l) => l.startsWith("  target: ")), `build ${built.join(" | ")}`);
        check(built.some((l) => l.startsWith("  rename-name RateX") || l.includes("RateX → RateY") || l.includes("RateY")), `build ${built.join(" | ")}`);
        check(built.some((l) => l.startsWith("  read back ")) && built.some((l) => l.startsWith(`  wrote book/${web ? "f7_base.xln.xlsx" : "f7_base.xlsx"}`)), `build ${built.join(" | ")}`);
        check(/^ {2}built in \d+ ms$/.test(built[built.length - 1] ?? ""), `build ${built.join(" | ")}`);
        const audited = after(head("audit f7_base\\.xlsx: "));
        check(audited.some((l) => / finding\(s\) in the Problems panel; report: .*; \d+ ms$/.test(l)), `audit ${audited.join(" | ")}`);
        const view = after(head("formula view f7_base\\.xlsx 'S1'$"));
        check(/^ {2}appearance order: \d+ formula cell\(s\)/.test(view[0] ?? ""), `formula view ${view.join(" | ")}`);
      } finally {
        await vscode.workspace.fs.writeFile(wb(), original);
        await vscode.commands.executeCommand("xln.pullWorkbook", wb(), { replace: true });
        await vscode.commands.executeCommand("workbench.action.closeAllEditors");
      }
    });

    test("Rename Symbol on a name made from its label, then Build: the output and the label notice say how to fix the label in Excel", async () => {
      const labels = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, "labels");
      const book = vscode.Uri.joinPath(labels, "labels.xlsx");
      const s2 = vscode.Uri.joinPath(labels, "labels.xln", "names", "sheets", "S2.xln");
      await vscode.commands.executeCommand("xln.pullWorkbook", book, { replace: true });
      const original = await read(book);
      try {
        const doc = await vscode.workspace.openTextDocument(s2);
        await vscode.window.showTextDocument(doc);
        const at = doc.getText().indexOf("COGS = ");
        check(at >= 0, `COGS not in S2.xln: ${doc.getText()}`);
        const edit = await vscode.commands.executeCommand<vscode.WorkspaceEdit>("vscode.executeDocumentRenameProvider", doc.uri, doc.positionAt(at + 1), "Cost_of_sales");
        check(edit && edit.size > 0 && (await vscode.workspace.applyEdit(edit)), "rename edit not applied");
        const r = await vscode.commands.executeCommand<BuildOutcome | undefined>("xln.buildWorkbook", book, { confirm: false });
        check(r?.status === "built", `outcome ${JSON.stringify(r)}`);
        const notice = [
          "COGS → Cost_of_sales: 1 label still reads the old name: 'S2'!D8",
          "In Excel, on sheet S2: Find & Replace (Ctrl+H on Windows; Ctrl+H or ⌘⇧H on Mac; or Home → Find & Select → Replace)",
          "  Find what:     COGS",
          "  Replace with:  Cost_of_sales",
          "  Within: Sheet · Look in: Formulas · Match entire cell contents ✓ · then Replace All",
        ];
        check(JSON.stringify(r.labelNotices) === JSON.stringify([notice]), `notices ${JSON.stringify(r.labelNotices)}`);
        const joined = r.lines.join("\n");
        check(joined.includes(notice.map((l) => `  ${l}`).join("\n")), `lines ${r.lines.join(" | ")}`);
        const uri = await vscode.commands.executeCommand<string | undefined>("xln.showLabelNotice", book, { copy: false });
        check(uri?.startsWith("xln-labels:"), `notice uri ${uri}`);
        const text = (await vscode.workspace.openTextDocument(vscode.Uri.parse(uri!))).getText();
        check(text.includes("  Find what:     COGS\n  Replace with:  Cost_of_sales\n"), `notice document: ${text}`);
      } finally {
        await vscode.workspace.fs.writeFile(book, original);
        await vscode.commands.executeCommand("xln.pullWorkbook", book, { replace: true });
        await vscode.commands.executeCommand("workbench.action.closeAllEditors");
      }
    });

    test("a file with no place in names/ is flagged and refused by the build (M3e)", async () => {
      await vscode.commands.executeCommand("xln.pullWorkbook", wb(), { replace: true });
      const names = vscode.Uri.joinPath(folder(), "f7_base.xln", "names");
      const foo = vscode.Uri.joinPath(names, "sheets", "Foo.xln");
      const notes = vscode.Uri.joinPath(names, "notes.txt");
      // As the Explorer creates them: onDidCreateFiles fires (a WorkspaceEdit does the same).
      const edit = new vscode.WorkspaceEdit();
      edit.createFile(foo, { contents: new TextEncoder().encode("X @A1 = 1;\n") });
      edit.createFile(notes, { contents: new TextEncoder().encode("remember\n") });
      check(await vscode.workspace.applyEdit(edit), "files not created");
      try {
        await vscode.commands.executeCommand("xln.reload");
        const on = (u: vscode.Uri) => vscode.languages.getDiagnostics(u).filter((d) => d.code === "stray-file").map((d) => d.message);
        check(on(foo).some((m) => m.startsWith("there is no sheet Foo in the workbook: sheets are created in Excel, then pulled")), `Foo.xln: ${on(foo).join(" | ")}`);
        check(on(notes).some((m) => m.includes("only .xln files belong in names/")), `notes.txt: ${on(notes).join(" | ")}`);
        const r = await vscode.commands.executeCommand<BuildOutcome | undefined>("xln.buildWorkbook", wb(), { confirm: false });
        check(r?.status === "refused" && r.errors.some((e) => e.startsWith("there is no sheet Foo")) && r.errors.some((e) => e.includes("notes.txt")), `outcome ${JSON.stringify(r)}`);
        // The refusal's errors are in the Problems panel too, as the build's, apart from the live checks.
        const fromBuild = vscode.languages.getDiagnostics(foo).filter((d) => d.source === "xln build");
        check(fromBuild.length > 0 && fromBuild.every((d) => d.severity === vscode.DiagnosticSeverity.Error && d.message.startsWith("build refused: ")), `build diagnostics ${JSON.stringify(fromBuild)}`);
        // The output lists the reasons before the plan.
        check(r.lines[1]?.startsWith("  error "), `lines ${r.lines.join(" | ")}`);
      } finally {
        await vscode.workspace.fs.delete(foo);
        await vscode.workspace.fs.delete(notes);
        await vscode.commands.executeCommand("xln.reload");
      }
    });

    test("a conflict refuses the build and opens as a diff", async () => {
      // "Excel" changes Rate: the pulled-from workbook gets another value behind the project.
      await vscode.commands.executeCommand("xln.pullWorkbook", wb(), { replace: true });
      const original = await read(wb());
      const text = new TextDecoder().decode(await read(unmanaged()));
      await setRate("0.9");
      const r1 = await vscode.commands.executeCommand<BuildOutcome | undefined>("xln.buildWorkbook", wb(), { confirm: false });
      check(r1?.status === "built", `first ${r1?.status}`);
      const built = await read(vscode.Uri.parse(r1.written!));
      // Back to the original project and lock, with the built bytes as "Excel's" edit.
      await vscode.workspace.fs.writeFile(wb(), original);
      await vscode.commands.executeCommand("xln.pullWorkbook", wb(), { replace: true });
      await vscode.workspace.fs.writeFile(wb(), built);
      await vscode.workspace.fs.writeFile(unmanaged(), new TextEncoder().encode(text.replace(/Rate = [0-9.]+;/, "Rate = 0.5;")));
      const r = await vscode.commands.executeCommand<BuildOutcome | undefined>("xln.buildWorkbook", wb(), { confirm: false });
      check(r?.status === "refused" && r.conflicts.length === 1 && r.conflicts[0]!.kind === "both-changed", `outcome ${JSON.stringify(r)}`);
      await vscode.workspace.fs.writeFile(wb(), original);
    });
  });
}
