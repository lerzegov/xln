// Timings on a real workbook, in a real extension host: runs instead of the other suites
// when the workspace holds `xln-smoke.json` ({ "workbook": "<file>.xlsx" }), which
// `npm run test:desktop -- --smoke <xlsx>` and `npm run test:web -- --smoke <xlsx>` write
// next to a copy of the workbook (scripts/fixture.mjs). The corpus is never written to.
import * as vscode from "vscode";
import type { PullOutcome } from "../commands.js";
import type { LoadStats } from "../xlnWorkspace.js";

export const SMOKE_FILE = "xln-smoke.json";

function check(cond: unknown, msg: string): asserts cond {
  if (!cond) throw new Error(msg);
}

export async function smokeConfig(): Promise<{ workbook: string } | undefined> {
  const folder = vscode.workspace.workspaceFolders?.[0]?.uri;
  if (!folder) return undefined;
  try {
    return JSON.parse(new TextDecoder().decode(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(folder, SMOKE_FILE))));
  } catch {
    return undefined;
  }
}

export function defineSmokeSuite(cfg: { workbook: string }): void {
  suite(`xln smoke on ${cfg.workbook}`, () => {
    const folder = () => vscode.workspace.workspaceFolders![0]!.uri;
    const host = vscode.env.uiKind === vscode.UIKind.Web ? "web" : "desktop";
    const log = (s: string) => console.log(`    [${host}] ${s}`);
    let out: vscode.Uri;

    test("pull", async () => {
      const t0 = Date.now();
      const r = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", vscode.Uri.joinPath(folder(), cfg.workbook), { replace: true });
      check(r, "no pull");
      out = vscode.Uri.parse(r.out);
      log(`pull: ${r.names} names, ${r.written.length} files; read+pull ${r.pullMs} ms, write ${r.writeMs} ms, command total ${Date.now() - t0} ms (includes reload)`);
    });

    test("model load", async () => {
      const t0 = Date.now();
      const s = await vscode.commands.executeCommand<LoadStats>("xln.reload");
      log(`reload: ${s.projects} project, ${s.files} files, ${s.names} names in ${s.ms} ms (command ${Date.now() - t0} ms)`);
      check(s.names > 0, "no names");
    });

    test("outline, hover, references with cell usages, search", async () => {
      const names = vscode.Uri.joinPath(out, "names");
      const files = (await vscode.workspace.fs.readDirectory(names)).filter(([n]) => n.endsWith(".xln")).map(([n]) => n);
      let t0 = Date.now();
      const doc = await vscode.workspace.openTextDocument(vscode.Uri.joinPath(names, files[0]!));
      const syms = await vscode.commands.executeCommand<vscode.DocumentSymbol[]>("vscode.executeDocumentSymbolProvider", doc.uri);
      log(`outline of ${files[0]}: ${syms.length} symbols in ${Date.now() - t0} ms`);
      const sym = syms.find((s) => s.children.length === 0) ?? syms[0]!.children[0]!;
      t0 = Date.now();
      const hovers = await vscode.commands.executeCommand<vscode.Hover[]>("vscode.executeHoverProvider", doc.uri, sym.selectionRange.start);
      log(`first hover on ${sym.name} (reads the workbook): ${Date.now() - t0} ms, ${hovers.length} hover(s)`);
      t0 = Date.now();
      await vscode.commands.executeCommand<vscode.Hover[]>("vscode.executeHoverProvider", doc.uri, sym.selectionRange.start);
      log(`second hover: ${Date.now() - t0} ms`);
      t0 = Date.now();
      const refs = await vscode.commands.executeCommand<vscode.Location[]>("vscode.executeReferenceProvider", doc.uri, sym.selectionRange.start);
      log(`references of ${sym.name}: ${refs.length} (${refs.filter((r) => r.uri.scheme === "xln-cells").length} in cells) in ${Date.now() - t0} ms`);
      t0 = Date.now();
      const keys = await vscode.commands.executeCommand<string[]>("xln.searchNames", "base");
      log(`search "base": ${keys?.length} names in ${Date.now() - t0} ms`);
      t0 = Date.now();
      const ws = await vscode.commands.executeCommand<vscode.SymbolInformation[]>("vscode.executeWorkspaceSymbolProvider", "sales");
      log(`workspace symbols "sales": ${ws.length} in ${Date.now() - t0} ms`);
    });

    test("formula view of every sheet", async () => {
      const wb = vscode.Uri.joinPath(folder(), cfg.workbook);
      const listing = await vscode.workspace.fs.readDirectory(vscode.Uri.joinPath(out, "names", "sheets"));
      const sheets = listing.map(([n]) => decodeURIComponent(n.replace(/\.xln$/, "")));
      const t0 = Date.now();
      let lines = 0;
      for (const sheet of sheets) {
        const uri = await vscode.commands.executeCommand<string | undefined>("xln.formulaView", { workbook: wb.toString(), sheet });
        check(uri, `no view of ${sheet}`);
        const doc = await vscode.workspace.openTextDocument(vscode.Uri.parse(uri));
        lines += doc.lineCount;
      }
      log(`formula view of ${sheets.length} sheets with a sheet file: ${lines} lines in ${Date.now() - t0} ms (opening editors included)`);
    });
  });
}
