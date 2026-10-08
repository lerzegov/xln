// The audit in the editor, in the web extension host and in desktop VS Code: pulls
// `traps/traps.xlsx` of the fixture folder (the seeded-trap workbook), then checks that
// its findings appear as diagnostics (names on their `.xln` entry, cells on their formula
// view line) and that `xln: Audit workbook` opens the report with links.
import * as vscode from "vscode";
import type { PullOutcome } from "../commands.js";
import type { XlnApi } from "../extension.js";

// Found by package name, not by id: the publisher id is set when the Marketplace account exists.
function xlnApi(): XlnApi {
  return vscode.extensions.all.find((e) => e.packageJSON?.name === "xln")!.exports as XlnApi;
}

function check(cond: unknown, msg: string): asserts cond {
  if (!cond) throw new Error(msg);
}

export function defineAuditSuite(): void {
  suite("xln audit", () => {
    const folder = () => vscode.workspace.workspaceFolders![0]!.uri;
    const workbook = () => vscode.Uri.joinPath(folder(), "traps", "traps.xlsx");
    const unmanaged = () => vscode.Uri.joinPath(folder(), "traps", "traps.xln", "names", "_unmanaged.xln");
    const code = (d: vscode.Diagnostic) => (typeof d.code === "object" ? String(d.code.value) : String(d.code));

    test("findings of the traps workbook appear as diagnostics", async () => {
      const r = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", workbook(), { replace: true });
      check(r, "no pull outcome");
      await vscode.commands.executeCommand("xln.reload");
      const api = xlnApi();
      await api.refreshAudit();

      const doc = await vscode.workspace.openTextDocument(unmanaged());
      const diags = vscode.languages.getDiagnostics(unmanaged()).filter((d) => d.source === "xln check");
      const bare = diags.find((d) => code(d) === "C2.bare-prefix");
      check(bare, `no C2 diagnostic: ${diags.map((d) => code(d)).join(", ")}`);
      check(doc.getText(bare.range) === "SEQUENCE", `C2 on '${doc.getText(bare.range)}'`);
      check(bare.severity === vscode.DiagnosticSeverity.Error, "C2 is not an error");
      for (const rule of ["C2.poisoned", "C3.lambda", "C5.other-sheet", "C7.length", "C11.drift", "C12.name-cycle", "C13.constant"]) {
        check(diags.some((d) => code(d) === rule), `no ${rule} on _unmanaged.xln`);
      }
      // Cell findings sit on the cell's line of its formula view.
      const all = vscode.languages.getDiagnostics();
      const arity = all.flatMap(([uri, ds]) => ds.filter((d) => code(d) === "C6.lambda-arity").map((d) => ({ uri, d })));
      check(arity.length === 2, `C6 diagnostics: ${arity.map((x) => `${x.d.source} ${x.uri.toString()} ${x.d.range.start.line}`).join(" | ")}`);
      check(arity.every((x) => x.uri.scheme === "xln-formulas"), `C6 on ${arity.map((x) => x.uri.toString()).join(", ")}`);
      const view = await vscode.workspace.openTextDocument(arity[0]!.uri);
      check(view.getText(arity[0]!.d.range) === "A18", `C6 on '${view.getText(arity[0]!.d.range)}'`);
    });

    test("xln: Audit workbook opens the report, with links to names and cells", async () => {
      const from = xlnApi().logLines().length;
      const uri = await vscode.commands.executeCommand<string | undefined>("xln.auditWorkbook", workbook());
      check(uri, "no report");
      // The output channel: the CLI's counts under a timed `xln audit` header, the checks that found something.
      const log = xlnApi().logLines().slice(from);
      const k = log.findIndex((l) => /^\d\d:\d\d:\d\d xln audit traps\.xlsx: 9 errors, 6 warnings, 1 info$/.test(l));
      check(k > 0 && log[k - 1] === "", `log:\n${log.join("\n")}`);
      const end = log.indexOf("", k + 1);
      const body = log.slice(k + 1, end < 0 ? undefined : end);
      check(body.length > 1 && body.slice(0, -1).every((l) => /^ {2}C\d+ .*\d+ (error|warning|info)/.test(l)), `log:\n${log.join("\n")}`);
      check(/^ {2}\d+ finding\(s\) in the Problems panel; report: traps\.xlsx \(audit\); \d+ ms$/.test(body[body.length - 1]!), `log:\n${log.join("\n")}`);
      const doc = await vscode.workspace.openTextDocument(vscode.Uri.parse(uri));
      check(doc.getText().startsWith("xln check traps.xlsx: 9 errors, 6 warnings, 1 info"), `report: ${doc.getText().slice(0, 80)}`);
      check(doc.getText().includes("== C8 Name census"), "no census");
      const links = await vscode.commands.executeCommand<vscode.DocumentLink[]>("vscode.executeLinkProvider", doc.uri);
      const texts = links.map((l) => doc.getText(l.range));
      check(texts.includes("ReadsLocal") && texts.includes("'S1'!A18"), `links: ${texts.join(", ")}`);
      const name = links.find((l) => doc.getText(l.range) === "ReadsLocal")!;
      check(name.target?.scheme === "command" && name.target.path === "xln.revealName", `name link ${name.target?.toString()}`);
      const shown = await vscode.commands.executeCommand<string | undefined>("xln.revealName", { root: vscode.Uri.joinPath(folder(), "traps", "traps.xln").toString(), key: "ReadsLocal" });
      check(shown === unmanaged().toString(), `revealed ${shown}`);
    });

    test("a name whose label cell no longer gives it (C15) shows on its entry, from the workbook's cell values", async () => {
      const book = vscode.Uri.joinPath(folder(), "labels", "labels.xlsx");
      const r = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", book, { replace: true });
      check(r, "no pull outcome");
      await vscode.commands.executeCommand("xln.reload");
      await xlnApi().refreshAudit();
      const s2 = vscode.Uri.joinPath(folder(), "labels", "labels.xln", "names", "sheets", "S2.xln");
      const doc = await vscode.workspace.openTextDocument(s2);
      const diags = vscode.languages.getDiagnostics(s2).filter((d) => d.source === "xln check" && code(d) === "C15.label-drift");
      check(diags.length === 1, `C15 diagnostics: ${diags.map((d) => d.message).join(" | ")}`);
      const d = diags[0]!;
      check(doc.getText(d.range).includes("Gross_ind_income"), `C15 on '${doc.getText(d.range)}'`);
      check(d.message.includes(`'S2'!D9 reads "Gross income"`) && d.message.includes("probably renamed from Gross_income"), `message: ${d.message}`);
      check(d.severity === vscode.DiagnosticSeverity.Information, "C15 is not info");
    });

    test("the project's xln.config.json applies, and a replacing pull keeps it", async () => {
      const config = vscode.Uri.joinPath(folder(), "traps", "traps.xln", "xln.config.json");
      const original = await vscode.workspace.fs.readFile(config);
      check(JSON.parse(new TextDecoder().decode(original)).audit.harness.length === 0, "the pull did not write the default settings");
      const mine = new TextEncoder().encode(JSON.stringify({ audit: { harness: ["Unused", "Fact"], rules: { C13: "off" } } }));
      await vscode.workspace.fs.writeFile(config, mine);
      try {
        const r = await vscode.commands.executeCommand<PullOutcome | undefined>("xln.pullWorkbook", workbook(), { replace: true });
        check(r && !r.written.includes("xln.config.json"), "the pull rewrote xln.config.json");
        check(new TextDecoder().decode(await vscode.workspace.fs.readFile(config)) === new TextDecoder().decode(mine), "the pull lost xln.config.json");
        const api = xlnApi();
        await api.refreshAudit();
        const diags = vscode.languages.getDiagnostics(unmanaged()).filter((d) => d.source === "xln check");
        for (const rule of ["C10.unused", "C13.constant"]) check(!diags.some((d) => code(d) === rule), `${rule} despite the settings`);
        const uri = await vscode.commands.executeCommand<string | undefined>("xln.auditWorkbook", workbook());
        const doc = await vscode.workspace.openTextDocument(vscode.Uri.parse(uri!));
        // The report was open from the test above: VS Code re-reads it after the change event.
        const want = "xln check traps.xlsx: 9 errors, 4 warnings\n";
        for (let k = 0; k < 50 && !doc.getText().startsWith(want); k++) await new Promise((r) => setTimeout(r, 100));
        check(doc.getText().startsWith(want), `report: ${doc.getText().slice(0, 80)}`);
      } finally {
        await vscode.workspace.fs.writeFile(config, original);
        await vscode.commands.executeCommand("xln.reload");
      }
    });
  });
}
