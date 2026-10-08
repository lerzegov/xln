// One source for both entry points: esbuild bundles it for Node (desktop) and for a web
// worker (vscode.dev). All file access goes through vscode.workspace.fs, which is the only
// API that reaches a local folder opened in the browser (File System Access API).
import * as vscode from "vscode";
import { AuditFeature } from "./audit.js";
import { BuildFeature } from "./build.js";
import { EditorFeature } from "./editor.js";
import { pullWorkbook, searchCommand, type PullOptions } from "./commands.js";
import { Features } from "./features.js";
import { FormulaViews } from "./formulaView.js";
import { LibraryFeature } from "./library.js";
import { UxFeature } from "./ux.js";
import { baseName, dirName, formatReport, lockFileName, probeXlsx, type InspectReport } from "./inspect.js";
import { XlnWorkspace, type LoadStats } from "./xlnWorkspace.js";
import { Activity, loggedLines, logLine, setLogSink } from "./log.js";

let output: vscode.OutputChannel;

/** What tests and other extensions can reach. */
export interface XlnApi {
  workspace: XlnWorkspace;
  /** Audits every pulled project's workbook again (tests wait on it). */
  refreshAudit(): Promise<void>;
  /** The lines the xln output channel has shown, the last few thousand (tests). */
  logLines(): string[];
}

export function activate(context: vscode.ExtensionContext): XlnApi {
  output = vscode.window.createOutputChannel("xln");
  setLogSink((line) => output.appendLine(line));
  const log = logLine;
  warnOtherInstalls(context);
  const ws = new XlnWorkspace(log);
  const features = new Features(ws);
  features.register(context);
  const formulas = new FormulaViews(ws, features);
  formulas.register(context);
  const audits = new AuditFeature(ws, formulas);
  audits.register(context);
  // The checks as you type leave out what the audit says: republish when it has.
  features.cellFinding = (wb, rule, sheet, ref) => audits.hasCellFinding(wb, rule, sheet, ref);
  context.subscriptions.push(audits.onDidPublish(() => features.scheduleDiagnostics(false)));
  const library = new LibraryFeature(ws);
  library.register(context);
  new EditorFeature(ws, library).register(context);
  new BuildFeature(ws, () => output.show(true)).register(context);
  new UxFeature(ws).register(context);
  context.subscriptions.push(
    output,
    ws,
    vscode.commands.registerCommand("xln.inspectWorkbook", (uri?: vscode.Uri) => run("inspect", () => inspectWorkbook(uri))),
    vscode.commands.registerCommand("xln.writeTestFile", (uri?: vscode.Uri) => run("write test file", () => writeTestFile(uri))),
    vscode.commands.registerCommand("xln.pullWorkbook", (uri?: vscode.Uri, opts?: PullOptions) =>
      run("pull", () => pullWorkbook(ws, log, uri instanceof vscode.Uri ? uri : undefined, opts)),
    ),
    vscode.commands.registerCommand("xln.reload", (): Promise<LoadStats | undefined> => run("reload", () => ws.reload(true))),
    vscode.commands.registerCommand("xln.searchNames", (query?: string) => run("search", () => searchCommand(ws, typeof query === "string" ? query : undefined))),
    vscode.commands.registerCommand("xln.showUsages", (arg?: { root: string; key: string }) => run("usages", () => features.showUsages(arg))),
    vscode.commands.registerCommand("xln.formulaView", (arg?: unknown) => run("formula view", () => formulas.open(arg))),
    vscode.commands.registerCommand("xln.formulaViewCalc", (arg?: unknown) => run("formula view", () => formulas.open(arg, "calculation"))),
    vscode.commands.registerCommand("xln.workbookFormulaView", (arg?: unknown) => run("formula view", () => formulas.open(arg, "calculation", true))),
    vscode.commands.registerCommand("xln.auditWorkbook", (arg?: unknown) => run("audit", () => audits.open(arg))),
    vscode.commands.registerCommand("xln.formulaViewToggleOrder", (arg?: unknown) => run("formula view", () => formulas.toggleOrder(arg))),
  );
  void ws.ready();
  return { workspace: ws, refreshAudit: () => audits.refresh(), logLines: loggedLines };
}

export function deactivate(): void {}

// Commands return their result so the web test can check it; errors go to the channel,
// because in vscode.dev the channel is the only place the author can read them.
async function run<T>(verb: string, task: () => Promise<T | undefined>): Promise<T | undefined> {
  try {
    return await task();
  } catch (err) {
    const a = new Activity(verb).head(": failed");
    if (err instanceof Error && err.stack) a.lines(err.stack.split("\n"));
    else a.line(String(err));
    output.show(true);
    vscode.window.showErrorMessage(`xln: ${err instanceof Error ? err.message : String(err)}`);
    return undefined;
  }
}

/**
 * The same extension under another publisher id (the test builds were `xln.xln` before the
 * publisher became `lerzegov-xln`): VS Code runs both, so every menu entry and editor button
 * shows twice, there are two "xln" output channels, and the second copy to activate fails
 * on its first command id. Said once, in the channel and as a warning.
 */
function warnOtherInstalls(context: vscode.ExtensionContext): void {
  const self = context.extension.id.toLowerCase();
  const others = vscode.extensions.all.filter((e) => e.id.toLowerCase() !== self && e.id.toLowerCase().endsWith(".xln") && (e.packageJSON as { name?: unknown } | undefined)?.name === "xln");
  if (others.length === 0) return;
  const ids = others.map((e) => e.id).join(", ");
  void new Activity("extension", context.extension.id).warn(
    `xln: another copy of this extension is installed (${ids}): menus and buttons show twice and the copies get in each other's way. Uninstall ${ids} in the Extensions view, then reload the window.`,
  );
}

function header(title: string, uri: vscode.Uri): void {
  logLine("");
  logLine(`xln ${title} ${uri.toString()}`);
  logLine(`  host: ${vscode.env.uiKind === vscode.UIKind.Web ? "web" : "desktop"} (${vscode.env.appHost})`);
}

async function inspectWorkbook(uri?: vscode.Uri): Promise<InspectReport | undefined> {
  const target = uri?.path.toLowerCase().endsWith(".xlsx") ? uri : await pickWorkbook();
  if (!target) return undefined;
  header("inspect", target);

  const bytes = await vscode.workspace.fs.readFile(target);
  const probe = probeXlsx(bytes);

  const folder = target.with({ path: dirName(target.path) });
  const name = baseName(target.path);
  const listing = await vscode.workspace.fs.readDirectory(folder);
  const ownerFiles = listing.map(([n]) => n).filter((n) => n.startsWith("~$"));

  const report: InspectReport = {
    ...probe,
    file: name,
    bytes: bytes.length,
    lockFile: lockFileName(name),
    lockFilePresent: ownerFiles.includes(lockFileName(name)),
    ownerFilesInFolder: ownerFiles,
  };
  for (const l of formatReport(report).split("\n")) logLine(`  ${l}`);
  output.show(true);
  return report;
}

// readDirectory rather than workspace.findFiles: findFiles depends on a search provider,
// which a browser-opened local folder may not have; readDirectory is the same API the
// probe is testing anyway.
async function pickWorkbook(): Promise<vscode.Uri | undefined> {
  const found: vscode.Uri[] = [];
  for (const folder of vscode.workspace.workspaceFolders ?? []) await walk(folder.uri, 0, found);
  if (found.length === 0) {
    vscode.window.showWarningMessage("xln: no .xlsx files in the workspace.");
    return undefined;
  }
  const items = found.map((uri) => ({ label: baseName(uri.path), description: vscode.workspace.asRelativePath(uri), uri }));
  const choice = await vscode.window.showQuickPick(items, { placeHolder: "Workbook to inspect" });
  return choice?.uri;
}

const SKIP = new Set(["node_modules", ".git", ".vscode-test", "dist"]);

async function walk(dir: vscode.Uri, depth: number, found: vscode.Uri[]): Promise<void> {
  if (depth > 6 || found.length >= 500) return;
  for (const [name, type] of await vscode.workspace.fs.readDirectory(dir)) {
    const child = vscode.Uri.joinPath(dir, name);
    if (type & vscode.FileType.Directory) {
      if (!SKIP.has(name)) await walk(child, depth + 1, found);
    } else if (name.toLowerCase().endsWith(".xlsx") && !name.startsWith("~$")) {
      found.push(child);
    }
  }
}

export interface WriteTestResult {
  file: string;
  written: string;
  readBack: string;
  ok: boolean;
  listed: boolean;
}

async function writeTestFile(uri?: vscode.Uri): Promise<WriteTestResult | undefined> {
  const folder = await targetFolder(uri);
  if (!folder) {
    vscode.window.showWarningMessage("xln: open a folder first.");
    return undefined;
  }
  const file = vscode.Uri.joinPath(folder, "xln-write-test.txt");
  header("write test file", file);

  const written = `xln write test ${new Date().toISOString()} from ${vscode.env.appHost}\n`;
  await vscode.workspace.fs.writeFile(file, new TextEncoder().encode(written));
  const readBack = new TextDecoder().decode(await vscode.workspace.fs.readFile(file));
  const listed = (await vscode.workspace.fs.readDirectory(folder)).some(([n]) => n === "xln-write-test.txt");

  const result = { file: file.toString(), written, readBack, ok: readBack === written, listed };
  logLine(`  wrote ${written.length} chars; read back ${result.ok ? "IDENTICAL" : "DIFFERENT: " + JSON.stringify(readBack)}`);
  logLine(`  listed by readDirectory: ${listed ? "yes" : "NO"}`);
  output.show(true);
  return result;
}

async function targetFolder(uri?: vscode.Uri): Promise<vscode.Uri | undefined> {
  if (uri) {
    const stat = await vscode.workspace.fs.stat(uri);
    return stat.type & vscode.FileType.Directory ? uri : uri.with({ path: dirName(uri.path) });
  }
  return vscode.workspace.workspaceFolders?.[0]?.uri;
}
