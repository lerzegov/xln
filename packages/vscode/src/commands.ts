// `xln: Pull workbook` and `xln: Search names`. The pull only reads the workbook; it
// writes the project folder beside it through vscode.workspace.fs (A, E: the build is the
// only command of this extension that writes an .xlsx).
//
// Every pull is fresh (decided 2026-10-06): names/**, the lockfile and the manifest are
// written as the workbook has them. When the project holds source edits not built yet, a
// modal lists them and offers Build first (the build, then the pull) or Discard and pull.
// Files whose layout or `//` comments alone differ from what the pull writes are listed
// too, as a group of their own; alone they get a modal with Pull anyway (M3e).

import * as vscode from "vscode";
import { CONFIG_FILE, defaultConfigText, formatUnbuiltEdit, isLocked, isSourcePath, LOCK_FILE, NAMES_DIR, pullProject, rewrittenFiles, unbuiltEdits, type UnbuiltEdit } from "@xln/core";
import type { BuildCommandOptions, BuildOutcome } from "./build.js";
import { baseName, dirName } from "./inspect.js";
import type { NameDef } from "./model/project.js";
import { formatPullSummary, isWorkbookName, projectFolderName, pullNotification, rewrittenQuestion, unbuiltQuestion } from "./model/pull.js";
import { Activity } from "./log.js";
import { closeStaleTabs } from "./ux.js";
import type { ProjectHandle, XlnWorkspace } from "./xlnWorkspace.js";

export interface PullOptions {
  /** The answer to the modal on source edits not built yet, without asking (tests). */
  unbuilt?: "build" | "discard" | "cancel";
  /** Older form of `unbuilt`: true discards, false cancels. */
  replace?: boolean;
  /** The answer to the modal on files rewritten for layout or comments alone (no edit not built), without asking (tests). */
  layout?: "pull" | "cancel";
  /** For "build": the build's options (tests: `{ confirm: false }`). */
  build?: BuildCommandOptions;
}

export interface PullOutcome {
  /** The project folder. */
  out: string;
  /** Files written (those the workbook says differently). */
  written: string[];
  names: number;
  notices: string[];
  /** Source edits not built that the pull replaced ("Discard and pull"). */
  discarded: UnbuiltEdit[];
  /** Names files rewritten although they held no edit not built: layout or comments only (M3e). */
  rewritten: string[];
  /** Editor tabs of files the pull removed: closed, or kept open with unsaved text. */
  tabs: { closed: string[]; kept: string[] };
  /** Time to read the workbook and compute the project, and to write it. */
  pullMs: number;
  writeMs: number;
}

const SKIP = new Set(["node_modules", ".git", ".vscode-test", ".vscode-test-web", "dist"]);

async function findWorkbooks(dir: vscode.Uri, depth: number, found: vscode.Uri[]): Promise<void> {
  if (depth > 6 || found.length >= 500) return;
  for (const [name, type] of await vscode.workspace.fs.readDirectory(dir)) {
    const child = vscode.Uri.joinPath(dir, name);
    if (type & vscode.FileType.Directory) {
      if (!SKIP.has(name) && !name.endsWith(".xln")) await findWorkbooks(child, depth + 1, found);
    } else if (isWorkbookName(name)) found.push(child);
  }
}

export async function pickWorkbook(placeHolder = "Workbook to pull", verb = "pull"): Promise<vscode.Uri | undefined> {
  const found: vscode.Uri[] = [];
  for (const folder of vscode.workspace.workspaceFolders ?? []) await findWorkbooks(folder.uri, 0, found);
  if (found.length === 0) {
    void new Activity(verb).warn("xln: no .xlsx or .xlsm files in the workspace.");
    return undefined;
  }
  const items = found.map((uri) => ({ label: baseName(uri.path), description: vscode.workspace.asRelativePath(uri), uri }));
  return (await vscode.window.showQuickPick(items, { placeHolder }))?.uri;
}

async function exists(uri: vscode.Uri): Promise<boolean> {
  try {
    await vscode.workspace.fs.stat(uri);
    return true;
  } catch {
    return false;
  }
}

export async function pullWorkbook(ws: XlnWorkspace, log: (s: string) => void, uri?: vscode.Uri, opts: PullOptions = {}): Promise<PullOutcome | undefined> {
  const target = uri && isWorkbookName(baseName(uri.path)) ? uri : await pickWorkbook();
  if (!target) return undefined;
  const name = baseName(target.path);
  const folder = target.with({ path: dirName(target.path) });
  const out = vscode.Uri.joinPath(folder, projectFolderName(name));

  const notices: string[] = [];
  // Re-list right before acting: vscode.dev has no file watching for a local folder.
  const listing = (await vscode.workspace.fs.readDirectory(folder)).map(([n]) => n);
  if (isLocked(name, listing)) {
    notices.push(`Excel has ${name} open (~$ file present): unsaved changes are not in the file and are not pulled.`);
  } else if (vscode.env.uiKind === vscode.UIKind.Web) {
    notices.push("in the browser Excel's ~$ lock file is not visible on a local folder: save the workbook in Excel before pulling, or unsaved changes are missed.");
  }

  // Every pull is fresh (2026-10-06). First, what it would replace that was never built:
  // the files on disk, with the text of unsaved editors over them.
  let bytes = await vscode.workspace.fs.readFile(target);
  let project = (await exists(out)) ? await readProjectSource(out) : undefined;
  let unbuilt = project ? unbuiltEdits({ workbook: bytes, fileName: name, files: project }) : [];
  let discarded: UnbuiltEdit[] = [];
  const t0 = Date.now();
  let pulled = pullProject(bytes, name);
  let pullMs = Date.now() - t0;
  // Besides the edits: files the pull rewrites for their layout or comments alone (M3e).
  const rewritten = project ? rewrittenFiles(project, pulled.files, unbuilt) : [];
  const logRewritten = () => {
    if (!rewritten.length) return;
    log(`  layout or comments only, rewritten by the pull: ${rewritten.join(", ")}`);
  };
  if (unbuilt.length > 0) {
    const answer = opts.unbuilt ?? (opts.replace === true ? "discard" : opts.replace === false ? "cancel" : await askUnbuilt(name, unbuilt, rewritten));
    log("");
    log(`xln pull ${name}: ${unbuilt.length} source edit(s) not built yet:`);
    for (const e of unbuilt) log(`  ${formatUnbuiltEdit(e)}`);
    logRewritten();
    if (answer === "cancel") {
      log("  pull cancelled: nothing written");
      return undefined;
    }
    if (answer === "build") {
      const built = await vscode.commands.executeCommand<BuildOutcome | undefined>("xln.buildWorkbook", target, opts.build);
      if (!built || (built.status !== "built" && built.status !== "up-to-date")) {
        log(`  pull stopped: the build did not finish (${built?.status ?? "cancelled"}); nothing pulled`);
        vscode.window.showWarningMessage(`xln: the build of ${name} did not finish (${built?.status ?? "cancelled"}): nothing pulled. See the xln output.`);
        return undefined;
      }
      bytes = await vscode.workspace.fs.readFile(target);
      project = await readProjectSource(out);
      unbuilt = project ? unbuiltEdits({ workbook: bytes, fileName: name, files: project }) : [];
      if (unbuilt.length > 0) {
        log("  pull stopped: after the build the project still has edits not built (a browser build writes a copy beside the workbook):");
        for (const e of unbuilt) log(`    ${formatUnbuiltEdit(e)}`);
        vscode.window.showWarningMessage(`xln: after the build ${projectFolderName(name)} still has ${unbuilt.length} edit(s) not in ${name}: nothing pulled. See the xln output.`);
        return undefined;
      }
      const t2 = Date.now();
      pulled = pullProject(bytes, name);
      pullMs = Date.now() - t2;
    } else discarded = unbuilt;
  } else if (rewritten.length > 0) {
    // Nothing would be lost to Excel, but hand-made layout and comments would: asked, not refused.
    const answer = opts.layout ?? (opts.replace === true ? "pull" : opts.replace === false ? "cancel" : await askRewritten(name, rewritten));
    log("");
    log(`xln pull ${name}:`);
    logRewritten();
    if (answer === "cancel") {
      log("  pull cancelled: nothing written");
      return undefined;
    }
  }
  if (rewritten.length) notices.push(`layout or comments only, rewritten as the workbook has them: ${rewritten.join(", ")}`);
  // Unsaved editors of the old files would feed their text back into the project (and
  // show the old errors): they are reverted first, as their edits are built or given up.
  await revertEditorsUnder(vscode.Uri.joinPath(out, NAMES_DIR));

  // The audit settings are the author's: kept, written once. library-bases/ (the library
  // bases' texts) is kept too: the pull writes and removes nothing outside names/**, the
  // lockfile and the manifest.
  let config: Uint8Array | undefined;
  try {
    config = await vscode.workspace.fs.readFile(vscode.Uri.joinPath(out, CONFIG_FILE));
  } catch {
    config = undefined;
  }

  const { files, report } = pulled;
  if (!config) files[CONFIG_FILE] ??= defaultConfigText();
  else delete files[CONFIG_FILE];

  const t1 = Date.now();
  // names/** is replaced: the .xln files this pull does not write go.
  const removed: string[] = [];
  for (const rel of Object.keys((await readProjectSource(out, false)) ?? {})) {
    if (rel.endsWith(".xln") && !(rel in files)) {
      await vscode.workspace.fs.delete(vscode.Uri.joinPath(out, ...rel.split("/")), { useTrash: false });
      removed.push(rel);
    }
  }
  if (removed.length) notices.push(`removed ${removed.join(", ")} (not written by this pull)`);
  const written: string[] = [];
  const dirs = new Set<string>();
  for (const rel of Object.keys(files)) {
    const parts = rel.split("/");
    for (let k = 1; k < parts.length; k++) dirs.add(parts.slice(0, k).join("/"));
  }
  await vscode.workspace.fs.createDirectory(out);
  for (const d of [...dirs].sort()) await vscode.workspace.fs.createDirectory(vscode.Uri.joinPath(out, ...d.split("/")));
  const enc = new TextEncoder();
  const dec = new TextDecoder();
  for (const [rel, text] of Object.entries(files)) {
    const uri = vscode.Uri.joinPath(out, ...rel.split("/"));
    // A file the workbook says the same about is left alone (open editors, timestamps, git).
    if ((await exists(uri)) && dec.decode(await vscode.workspace.fs.readFile(uri)) === text) continue;
    await vscode.workspace.fs.writeFile(uri, enc.encode(text));
    written.push(rel);
  }
  const writeMs = Date.now() - t1;

  log("");
  for (const line of formatPullSummary(report, vscode.workspace.asRelativePath(out), written, notices, discarded)) log(line);
  log(`  read and pulled in ${pullMs} ms, wrote in ${writeMs} ms`);
  await ws.reload();
  // Tabs of the files this pull removed would show a file that is gone (M3e).
  const tabs = await closeStaleTabs([out], true);

  void vscode.window
    .showInformationMessage(pullNotification(name, report.names, discarded.length, report.foreignModules, report.valueLabels), "Search names")
    .then((c) => c && vscode.commands.executeCommand("xln.searchNames"));
  return { out: out.toString(), written, names: report.names, notices, discarded, rewritten, tabs, pullMs, writeMs };
}

/** The modal before a pull that would replace source edits not built yet. */
async function askUnbuilt(workbook: string, edits: readonly UnbuiltEdit[], rewritten: readonly string[]): Promise<"build" | "discard" | "cancel"> {
  const q = unbuiltQuestion(workbook, edits, rewritten);
  const choice = await vscode.window.showWarningMessage(q.message, { modal: true, detail: q.detail }, "Build first", "Discard and pull");
  return choice === "Build first" ? "build" : choice === "Discard and pull" ? "discard" : "cancel";
}

/** The modal before a pull that would rewrite files for their layout or comments alone. */
async function askRewritten(workbook: string, rewritten: readonly string[]): Promise<"pull" | "cancel"> {
  const q = rewrittenQuestion(workbook, rewritten);
  return (await vscode.window.showWarningMessage(q.message, { modal: true, detail: q.detail }, "Pull anyway")) === "Pull anyway" ? "pull" : "cancel";
}

/** Reverts every unsaved editor of a file below `dir` (its text on disk comes back). */
async function revertEditorsUnder(dir: vscode.Uri): Promise<void> {
  const prefix = dir.toString() + "/";
  for (const doc of vscode.workspace.textDocuments) {
    if (!doc.isDirty || !doc.uri.toString().startsWith(prefix)) continue;
    await vscode.window.showTextDocument(doc, { preview: false });
    await vscode.commands.executeCommand("workbench.action.files.revert");
  }
}

/**
 * The project's names files and lockfile, path → text; undefined when it has no names
 * files. `editors`: the text of an unsaved editor instead of the file's.
 */
async function readProjectSource(root: vscode.Uri, editors = true): Promise<Record<string, string> | undefined> {
  const files: Record<string, string> = {};
  const dec = new TextDecoder();
  const dirty = new Map<string, string>();
  if (editors) for (const d of vscode.workspace.textDocuments) if (d.isDirty) dirty.set(d.uri.toString(), d.getText());
  const walk = async (dir: vscode.Uri, rel: string, depth: number): Promise<void> => {
    if (depth > 6) return;
    let entries: [string, vscode.FileType][];
    try {
      entries = await vscode.workspace.fs.readDirectory(dir);
    } catch {
      return;
    }
    for (const [n, type] of entries) {
      const path = rel ? `${rel}/${n}` : n;
      const child = vscode.Uri.joinPath(dir, n);
      if (type & vscode.FileType.Directory) await walk(child, path, depth + 1);
      else if (isSourcePath(path)) files[path] = dirty.get(child.toString()) ?? dec.decode(await vscode.workspace.fs.readFile(child));
    }
  };
  await walk(vscode.Uri.joinPath(root, NAMES_DIR), NAMES_DIR, 0);
  if (Object.keys(files).length === 0) return undefined;
  try {
    files[LOCK_FILE] = dec.decode(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(root, LOCK_FILE)));
  } catch {
    // No lockfile: the source is compared with the workbook itself.
  }
  return files;
}

interface NamePick extends vscode.QuickPickItem {
  handle: ProjectHandle;
  def: NameDef;
}

function oneLine(s: string, max = 120): string {
  const t = s.split(/\s+/).join(" ");
  return t.length > max ? t.slice(0, max - 1) + "…" : t;
}

/** The names matching `query` in every project, best first. */
export function searchNames(ws: XlnWorkspace, query: string): NamePick[] {
  const out: NamePick[] = [];
  const many = ws.projects.length > 1;
  for (const handle of ws.projects) {
    for (const hit of handle.project.search(query, 200)) {
      const d = hit.def;
      const where = `${d.scope === undefined ? "" : d.scope + " · "}${d.file.path.slice("names/".length)}${many ? ` · ${handle.project.workbookName ?? ""}` : ""}`;
      const doc = d.entry.doc ? oneLine(d.entry.doc, 80) + "  —  " : "";
      out.push({ label: d.name, description: where, detail: doc + oneLine(d.entry.formula), alwaysShow: true, handle, def: d });
    }
  }
  return out;
}

/**
 * B2: a quick pick over names, definitions and doc comments. With a `query` argument it
 * returns the matching keys without showing anything (for tests and other commands).
 */
export async function searchCommand(ws: XlnWorkspace, query?: string): Promise<string[] | undefined> {
  await ws.ready();
  if (query !== undefined) return searchNames(ws, query).map((p) => p.def.key);
  const qp = vscode.window.createQuickPick<NamePick>();
  qp.placeholder = "Search names, definitions and doc comments (words in any order)";
  qp.matchOnDescription = false;
  qp.matchOnDetail = false;
  const update = () => (qp.items = searchNames(ws, qp.value));
  qp.onDidChangeValue(update);
  update();
  const chosen = await new Promise<NamePick | undefined>((resolve) => {
    qp.onDidAccept(() => resolve(qp.selectedItems[0]));
    qp.onDidHide(() => resolve(undefined));
    qp.show();
  });
  qp.dispose();
  if (!chosen) return undefined;
  const loc = chosen.handle.project.nameLoc(chosen.def);
  const lines = chosen.def.file.lines;
  const a = lines.position(loc.start);
  const b = lines.position(loc.end);
  await vscode.window.showTextDocument(ws.fileUri(chosen.handle, loc.path), { selection: new vscode.Range(a.line, a.character, b.line, b.character) });
  return [chosen.def.key];
}
