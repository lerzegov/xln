// The browser probe, automated. Runs in the web extension host (headless Chromium via
// @vscode/test-web) and in desktop VS Code (@vscode/test-electron), on the fixture
// folder built by scripts/fixture.mjs.
//
// Under test-web the folder is served by test-web's own file system provider (writes
// stay in memory), not by the browser's File System Access API: that run proves the
// extension code in a web worker, not vscode.dev's access to a local folder.
import * as vscode from "vscode";
import type { InspectReport } from "../inspect.js";
import type { WriteTestResult } from "../extension.js";

function check(cond: unknown, msg: string): asserts cond {
  if (!cond) throw new Error(msg);
}

export function defineProbeSuite(expectWeb: boolean): void {
  suite("xln browser probe", () => {
    const folder = () => vscode.workspace.workspaceFolders![0]!.uri;

    test(`runs in the ${expectWeb ? "web" : "desktop"} extension host`, () => {
      check((vscode.env.uiKind === vscode.UIKind.Web) === expectWeb, `uiKind = ${vscode.env.uiKind}`);
    });

    test("reads and unzips a workbook, sees the ~$ lock file", async () => {
      const uri = vscode.Uri.joinPath(folder(), "probe_mac.xlsx");
      const r = await vscode.commands.executeCommand<InspectReport | undefined>("xln.inspectWorkbook", uri);
      check(r, "no report");
      check(r.entries.some((e) => e.name === "xl/workbook.xml"), "xl/workbook.xml missing");
      check(r.definedNames > 0, `definedNames = ${r.definedNames}`);
      check(r.lockFilePresent, `lock file not seen; owner files: ${r.ownerFilesInFolder.join(", ")}`);
    });

    test("reports no lock file when there is none", async () => {
      const uri = vscode.Uri.joinPath(folder(), "sub", "probe_win.xlsx");
      const r = await vscode.commands.executeCommand<InspectReport | undefined>("xln.inspectWorkbook", uri);
      check(r, "no report");
      check(!r.lockFilePresent, "unexpected lock file");
    });

    test("writes a file and reads it back", async () => {
      const r = await vscode.commands.executeCommand<WriteTestResult | undefined>("xln.writeTestFile", folder());
      check(r, "no result");
      check(r.ok, `read back differs: ${JSON.stringify(r)}`);
      check(r.listed, "written file not listed by readDirectory");
    });

    test("opens .xln files with the xln language", async () => {
      const doc = await vscode.workspace.openTextDocument(vscode.Uri.joinPath(folder(), "demo.xln"));
      check(doc.languageId === "xln", `languageId = ${doc.languageId}`);
    });
  });
}
