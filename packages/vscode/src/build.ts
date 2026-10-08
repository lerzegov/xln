// `xln: Build workbook`: writes the project's names into its workbook through
// vscode.workspace.fs. The core plans, patches and reads back (E3 in memory); here: the
// target (desktop overwrites, browser writes `<name>.xln.xlsx`, E1), the lock-file guard,
// the backup (E5), the re-read of what was written, and conflicts shown as diffs (E2).

import * as vscode from "vscode";
import { buildWorkbook, cellValueMap, CONFIG_FILE, describeChange, embedSetting, isLocked, isSourcePath, labelNoticeLines, LOCK_FILE, NAMES_DIR, readBack, readWorkbook, refusalReasons, renameLabelNotices, renamedConsumedLine, type BuildResult, type Change, type Conflict, type LabelNotice } from "@xln/core";
import { excelHost } from "./excelHost.js";
import { Activity } from "./log.js";
import { baseName, dirName } from "./inspect.js";
import { buildDiagnostics, buildTarget, conflictTexts, formatBuildSummary, refusalMessage } from "./model/build.js";
import type { ProjectHandle, XlnWorkspace } from "./xlnWorkspace.js";

export const CONFLICT_SCHEME = "xln-conflict";
/** The label notice of the last build that renamed names, as a read-only document. */
export const LABELS_SCHEME = "xln-labels";
const CLOSE_AND_BUILD = "Close in Excel and build";
const SHOW_LABELS = "Show labels to fix";

export interface BuildOutcome {
  status: BuildResult["status"] | "locked" | "write-failed" | "excel-unsaved" | "excel-failed" | "cancelled" | "failed";
  /** "Build and reopen in Excel": what Excel did. */
  excel?: string[];
  /** The file written, as a URI string. */
  written?: string;
  backup?: string;
  /** The `@renamed(…)` annotations the build removed from the source, one line per file. */
  renamedRemoved?: string[];
  /** The label notice, one block of lines per renamed name whose label cells still read its old name. */
  labelNotices?: string[][];
  changes: number;
  conflicts: Conflict[];
  errors: string[];
  lines: string[];
}

export interface BuildCommandOptions {
  /** Do not ask before writing (tests). */
  confirm?: boolean;
  /** E7 (desktop): close the workbook in Excel without saving, build, open it again. */
  reopen?: boolean;
  /** With `reopen`: close it even with unsaved changes in Excel, without asking. */
  discard?: boolean;
  /** When the workbook is open in Excel, the answer to *Close in Excel and build* without asking (tests): true closes it and builds, false declines. */
  closeInExcel?: boolean;
}

/** The changes, described, the first `max`, then how many more. */
function changeList(changes: readonly Change[], max = 12, sep = "\n"): string {
  const shown = changes.slice(0, max).map(describeChange);
  if (changes.length > max) shown.push(`… and ${changes.length - max} more`);
  return shown.join(sep);
}

function sleep(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  return a.length === b.length && a.every((x, i) => x === b[i]);
}

export class BuildFeature {
  private readonly texts = new Map<string, string>();
  /** The last build's label notices, by workbook URI. */
  private readonly notices = new Map<string, { name: string; notices: LabelNotice[] }>();
  /** The last build's blocking errors, apart from the live checks (source `xln build`). */
  private readonly diagnostics = vscode.languages.createDiagnosticCollection("xln build");

  constructor(
    private readonly ws: XlnWorkspace,
    private readonly showOutput: () => void = () => void vscode.commands.executeCommand("workbench.action.output.toggleOutput"),
  ) {}

  register(context: vscode.ExtensionContext): void {
    context.subscriptions.push(
      this.diagnostics,
      // A file edited after the build: its build errors may be fixed; the next build says.
      vscode.workspace.onDidChangeTextDocument((e) => {
        if (e.contentChanges.length && this.diagnostics.has(e.document.uri)) this.diagnostics.delete(e.document.uri);
      }),
      vscode.workspace.registerTextDocumentContentProvider(CONFLICT_SCHEME, { provideTextDocumentContent: (uri) => this.texts.get(uri.toString()) ?? "" }),
      vscode.workspace.registerTextDocumentContentProvider(LABELS_SCHEME, { provideTextDocumentContent: (uri) => this.labelText(uri.query) }),
      vscode.commands.registerCommand("xln.showLabelNotice", (arg?: unknown, opts?: { copy?: boolean }) => this.showLabelNotice(arg, opts)),
      vscode.commands.registerCommand("xln.buildWorkbook", (arg?: unknown, opts?: BuildCommandOptions) => this.run(arg, opts)),
      vscode.commands.registerCommand("xln.buildAndReopen", (arg?: unknown, opts?: BuildCommandOptions) => this.run(arg, { ...opts, reopen: true })),
    );
  }

  private async pickProject(arg: unknown, act: Activity): Promise<ProjectHandle | undefined> {
    await this.ws.ready();
    const uri = arg instanceof vscode.Uri ? arg : vscode.window.activeTextEditor?.document.uri;
    if (uri) {
      const byWorkbook = this.ws.handleForWorkbook(uri);
      if (byWorkbook) return byWorkbook;
      const at = this.ws.handleAt(uri);
      if (at) return at;
    }
    const withBook = this.ws.projects.filter((h) => h.workbookUri);
    if (withBook.length === 1) return withBook[0];
    if (withBook.length === 0) {
      act.warn("xln: no pulled project with its workbook in the workspace. Pull a workbook first.");
      return undefined;
    }
    const pick = await vscode.window.showQuickPick(
      withBook.map((h) => ({ label: baseName(h.workbookUri!.path), description: vscode.workspace.asRelativePath(h.uri), h })),
      { placeHolder: "Workbook to build" },
    );
    return pick?.h;
  }

  private async readFiles(root: vscode.Uri): Promise<Record<string, string>> {
    const files: Record<string, string> = {};
    const dec = new TextDecoder();
    const walk = async (dir: vscode.Uri, rel: string, depth: number): Promise<void> => {
      if (depth > 6) return;
      for (const [name, type] of await vscode.workspace.fs.readDirectory(dir)) {
        const path = rel ? `${rel}/${name}` : name;
        const child = vscode.Uri.joinPath(dir, name);
        if (type & vscode.FileType.Directory) await walk(child, path, depth + 1);
        else if (isSourcePath(path) || path === LOCK_FILE || path === CONFIG_FILE) files[path] = dec.decode(await vscode.workspace.fs.readFile(child));
        // Any other file below names/ is listed, unread: the checker refuses it (M3e).
        else if (path.startsWith(`${NAMES_DIR}/`)) files[path] = "";
      }
    };
    await walk(root, "", 0);
    return files;
  }

  /**
   * Every Build ends with something the author sees (feedback 2026-10-07: a build that
   * wrote nothing said nothing): the built message, "up to date", the refusal modal, or
   * the reason it stopped. An unexpected failure is shown too, not left to the host.
   */
  async run(arg?: unknown, opts: BuildCommandOptions = {}): Promise<BuildOutcome | undefined> {
    const act = new Activity(opts.reopen ? "build and reopen" : "build");
    try {
      return await this.runBuild(arg, opts, act);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      act.line(`failed: ${msg}; nothing written`);
      void vscode.window.showErrorMessage(`xln: the build failed: ${msg}. Nothing was written.`, "Show details").then((c) => c && this.showOutput());
      return { status: "failed", changes: 0, conflicts: [], errors: [msg], lines: [`xln build: failed: ${msg}`] };
    }
  }

  /** The open documents of a project's names files and config, by project path. */
  private projectDocuments(root: vscode.Uri): Map<string, vscode.TextDocument> {
    const out = new Map<string, vscode.TextDocument>();
    const prefix = root.toString() + "/";
    for (const d of vscode.workspace.textDocuments) {
      const s = d.uri.toString();
      if (!s.startsWith(prefix)) continue;
      const path = decodeURIComponent(s.slice(prefix.length));
      if (isSourcePath(path) || path === CONFIG_FILE) out.set(path, d);
    }
    return out;
  }

  private async runBuild(arg: unknown, opts: BuildCommandOptions, act: Activity): Promise<BuildOutcome | undefined> {
    let excel = opts.reopen ? excelHost() : undefined;
    if (opts.reopen && (!excel || vscode.env.uiKind === vscode.UIKind.Web)) {
      act.error("xln: Build and reopen needs desktop VS Code and desktop Excel (macOS or Windows). Use Build workbook.");
      return undefined;
    }
    const handle = await this.pickProject(arg, act);
    if (!handle) return undefined;
    if (!handle.workbookUri) {
      act.error(`xln: ${vscode.workspace.asRelativePath(handle.uri)} has no workbook beside it (its manifest names the workbook): nothing to build into.`);
      return undefined;
    }
    if (opts.reopen && handle.workbookUri.scheme !== "file") {
      act.error("xln: Build and reopen works on a workbook in a local folder.");
      return undefined;
    }
    // Unsaved edits are saved first, each save awaited; the project is matched by its
    // folder, not by the handle object a reload (which a save schedules) replaces.
    const saved = new Map<string, vscode.TextDocument>();
    for (const [path, d] of this.projectDocuments(handle.uri)) {
      if (!d.isDirty) continue;
      if (!(await d.save())) {
        act.error(`xln: ${vscode.workspace.asRelativePath(d.uri)} could not be saved: nothing built.`);
        return undefined;
      }
      saved.set(path, d);
    }
    const wbUri = handle.workbookUri;
    const name = baseName(wbUri.path);
    const folder = wbUri.with({ path: dirName(wbUri.path) });
    const web = vscode.env.uiKind === vscode.UIKind.Web;
    const target = buildTarget(name, web);
    act.target = name;

    const original = await vscode.workspace.fs.readFile(wbUri);
    const files = await this.readFiles(handle.uri);
    // The text just saved, as the editor has it: the build never plans from the file as
    // it was before the save, however the disk caught up.
    for (const [path, d] of saved) files[path] = d.getText();
    // D5: the project config says whether the source is embedded (opt-in).
    const embed = embedSetting(undefined, files[CONFIG_FILE]);
    // Nothing here changes a scope or a formula on its own (M3d): the build writes the source.
    const result = buildWorkbook({ workbook: original, fileName: name, files }, { embed });
    const lines = formatBuildSummary(name, result);
    if (saved.size) lines.push(`  saved first: ${[...saved.keys()].join(", ")}`);
    if (result.status === "built") lines.splice(1, 0, `  target: ${target.overwrite ? `${target.file}, overwritten (previous file kept as ${target.backup})` : `${target.file}, a new file beside ${name} (the browser never overwrites the workbook)`}`);
    const outcome: BuildOutcome = {
      status: result.status,
      changes: result.plan.changeSet.changes.length,
      conflicts: result.plan.conflicts,
      errors: result.plan.problems.filter((p) => p.severity === "error").map((p) => p.message),
      lines,
    };
    const finish = (o: BuildOutcome): BuildOutcome => {
      // The shared summary says `xln build …`; Build and reopen says what was asked.
      const [first, ...rest] = o.lines;
      const head = opts.reopen && first?.startsWith("xln build ") ? `xln build and reopen ${first.slice("xln build ".length)}` : first;
      act.summary([...(head === undefined ? [] : [head]), ...rest, `  ${o.status} in ${act.ms()} ms`]);
      return o;
    };
    this.publishDiagnostics(handle, result);

    if (result.status === "refused") {
      finish(outcome);
      // Tests and scripted builds (no questions asked) get the message without the modal.
      void this.showRefusal(handle, name, result, opts.confirm !== false);
      return outcome;
    }
    if (result.status === "read-back-failed") {
      finish(outcome);
      vscode.window.showErrorMessage(`xln: the build of ${name} did not read back as intended; nothing was written (see the xln output).`);
      return outcome;
    }
    if (result.status === "up-to-date") {
      // Build and reopen asked to see the workbook in Excel: nothing to write is no reason
      // not to (author, 2026-10-08: the command seemed to do nothing).
      if (excel) {
        const o = excel.open(wbUri.fsPath);
        outcome.excel = [o.message];
        lines.push(`  Excel: ${o.message}`);
        if (!o.ok) vscode.window.showErrorMessage(`xln: ${name} is up to date, but ${o.message}.`);
        else vscode.window.showInformationMessage(`xln: ${name} is up to date (nothing written); opened in Excel.`);
        return finish(outcome);
      }
      vscode.window.showInformationMessage(`xln: ${name} is up to date: the workbook already matches the source; nothing written.`);
      return finish(outcome);
    }

    // The workbook open in Excel: offer to close it, build, and open it again (E7), rather
    // than stop. Excel's unsaved changes are asked about before anything is closed.
    let confirmed = opts.confirm === false;
    if (!excel && target.overwrite && isLocked(target.file, (await vscode.workspace.fs.readDirectory(folder)).map(([n]) => n))) {
      const host = wbUri.scheme === "file" ? excelHost() : undefined;
      if (!host) {
        outcome.status = "locked";
        lines.push(`  Excel has ${target.file} open (~$ file present): close it in Excel, then build again. Nothing written.`);
        vscode.window.showErrorMessage(`xln: Excel has ${target.file} open. Close it, then build again.`);
        return finish(outcome);
      }
      const choice = opts.closeInExcel !== undefined
        ? (opts.closeInExcel ? CLOSE_AND_BUILD : undefined)
        : await vscode.window.showWarningMessage(
            `Excel has ${target.file} open. Close it in Excel, write ${outcome.changes} change(s) into it (the previous file is kept as ${target.backup}), and open it again?`,
            { modal: true },
            CLOSE_AND_BUILD,
          );
      if (choice !== CLOSE_AND_BUILD) {
        outcome.status = "locked";
        lines.push(`  Excel has ${target.file} open: not closed, nothing written.`);
        return finish(outcome);
      }
      excel = host;
      lines.push(`  Excel had ${target.file} open: Close in Excel and build (it is opened again after the write)`);
      confirmed = true;
    }

    if (!confirmed) {
      const what = target.overwrite ? `Write ${outcome.changes} change(s) into ${name}? The previous file is kept as ${target.backup}.` : `Write ${outcome.changes} change(s) into a new file ${target.file} beside ${name}? (In the browser the workbook itself is never overwritten.)`;
      const afe = result.plan.problems.find((p) => p.code === "afe-modules");
      const detail = changeList(result.plan.changeSet.changes) + (afe ? `\n\nWarning: ${afe.message}.` : "");
      if ((await vscode.window.showWarningMessage(what, { modal: true, detail }, "Build")) !== "Build") {
        outcome.status = "cancelled";
        lines.push("  cancelled: nothing written.");
        void vscode.window.showInformationMessage(`xln: build of ${name} cancelled: nothing written.`);
        return finish(outcome);
      }
    }

    const targetUri = vscode.Uri.joinPath(folder, target.file);
    if (excel) {
      // E7: Excel lets go of the file (without saving) before it is written.
      const path = targetUri.fsPath;
      const state = excel.state(path);
      outcome.excel = [];
      if (state.open && state.saved === false && !opts.discard) {
        const choice = await vscode.window.showWarningMessage(
          `Excel has unsaved changes in ${target.file}. Closing it for the build loses them. (To keep them: save in Excel, pull, and build again.)`,
          { modal: true },
          "Discard Excel's changes and build",
        );
        if (choice !== "Discard Excel's changes and build") {
          outcome.status = "excel-unsaved";
          lines.push(`  Excel has unsaved changes in ${target.file}: nothing written.`);
          return finish(outcome);
        }
      }
      if (state.open) {
        const c = excel.close(path);
        outcome.excel.push(c.message);
        lines.push(`  Excel: ${c.message}`);
        if (!c.ok) {
          outcome.status = "excel-failed";
          vscode.window.showErrorMessage(`xln: ${c.message}. Nothing written.`);
          return finish(outcome);
        }
        // Excel removes its owner file shortly after closing.
        for (let k = 0; k < 50 && isLocked(target.file, (await vscode.workspace.fs.readDirectory(folder)).map(([n]) => n)); k++) await sleep(200);
      }
    }
    // E1: re-list right before writing (vscode.dev has no file watching).
    const listing = (await vscode.workspace.fs.readDirectory(folder)).map(([n]) => n);
    if (isLocked(target.file, listing)) {
      outcome.status = "locked";
      lines.push(`  Excel has ${target.file} open (~$ file present): close it in Excel, then build again. Nothing written.`);
      vscode.window.showErrorMessage(`xln: Excel has ${target.file} open. Close it, then build again.`);
      return finish(outcome);
    }
    let backupUri: vscode.Uri | undefined;
    if (target.backup) {
      backupUri = vscode.Uri.joinPath(folder, target.backup);
      await vscode.workspace.fs.writeFile(backupUri, original);
      outcome.backup = backupUri.toString();
    }
    const bytes = result.bytes!;
    await vscode.workspace.fs.writeFile(targetUri, bytes);
    const onDisk = await vscode.workspace.fs.readFile(targetUri);
    const rb = readBack(original, onDisk, result.plan.changeSet.changes, result.plan.inSync);
    if (!sameBytes(onDisk, bytes) || !rb.ok) {
      if (backupUri && target.overwrite) await vscode.workspace.fs.writeFile(targetUri, original);
      outcome.status = "write-failed";
      lines.push(`  the written file did not read back as built${target.overwrite ? "; the original was restored" : ""}: ${rb.problems.join("; ")}`);
      if (excel) lines.push(`  Excel: ${excel.open(targetUri.fsPath).message}`);
      vscode.window.showErrorMessage(`xln: writing ${target.file} failed the read-back; see the xln output.`);
      return finish(outcome);
    }
    outcome.written = targetUri.toString();
    lines.push(`  read back ${target.file} from disk: as built (${outcome.changes} change(s) present, nothing else changed)`);
    const enc = new TextEncoder();
    if (target.updateLock) {
      for (const [rel, text] of Object.entries(result.files ?? {})) await vscode.workspace.fs.writeFile(vscode.Uri.joinPath(handle.uri, ...rel.split("/")), enc.encode(text));
    }
    lines.push(`  wrote ${vscode.workspace.asRelativePath(targetUri)}${backupUri ? `; previous file kept as ${target.backup}` : ""}`);
    // Only where the lockfile follows: a browser build's copy leaves the source as it is.
    if (target.updateLock) {
      const removed = await this.removeRenamed(handle.uri, result, files);
      lines.push(...removed.map((r) => `  ${r.line}`));
      const done = removed.filter((r) => r.removed);
      if (done.length) outcome.renamedRemoved = done.map((r) => r.line);
    }
    if (!target.updateLock) lines.push(`  ${name} and the lockfile are unchanged: open ${target.file} in Excel to check it.`);
    // Labels the rename left behind: xln never writes cell values, so it says how to fix them in Excel.
    const notices = labelNoticesOf(original, result.plan.changeSet.changes);
    if (notices.length) {
      outcome.labelNotices = notices.map(labelNoticeLines);
      for (const block of outcome.labelNotices) lines.push(...block.map((l) => `  ${l}`));
      this.notices.set(wbUri.toString(), { name, notices });
    } else this.notices.delete(wbUri.toString());
    if (excel) {
      const o = excel.open(targetUri.fsPath);
      outcome.excel!.push(o.message);
      lines.push(`  Excel: ${o.message}`);
      if (!o.ok) {
        outcome.status = "excel-failed";
        finish(outcome);
        await this.ws.reload();
        vscode.window.showErrorMessage(`xln: built ${target.file}, but ${o.message}`);
        return outcome;
      }
    }
    finish(outcome);
    // Said before the reload, so nothing after the write can swallow it.
    const afeNote = result.plan.problems.some((p) => p.code === "afe-modules") ? " AFE's modules in the workbook still have the old text: do not save them from AFE before bringing them in line (see Show details)." : "";
    const next = (excel ? "Excel has it open: check, save; then run xln verify." : "Open it in Excel, check, save; then run xln verify.") + afeNote;
    const spent = outcome.renamedRemoved?.length ? " The @renamed lines of the built renames are removed from the source." : "";
    const stale = notices.reduce((k, n) => k + n.stale.length, 0);
    const labels = notices.length ? ` ${stale ? `${stale} label cell${stale === 1 ? "" : "s"} in Excel still read${stale === 1 ? "s" : ""} the old name` : "Some label cells resemble the old name"}: ${SHOW_LABELS} says how to fix them with Find & Replace.` : "";
    const buttons = notices.length ? [SHOW_LABELS, "Show details"] : ["Show details"];
    void vscode.window.showInformationMessage(`xln: built ${target.file}: ${outcome.changes} change${outcome.changes === 1 ? "" : "s"} (${changeList(result.plan.changeSet.changes, 4, ", ")}).${spent}${labels} ${next}`, ...buttons).then((c) => {
      if (c === SHOW_LABELS) void this.showLabelNotice(wbUri);
      else if (c) this.showOutput();
    });
    await this.ws.reload();
    return outcome;
  }

  /** The label notice document's text for a workbook (its URI), or a line saying there is none. */
  private labelText(workbook: string): string {
    const n = this.notices.get(workbook);
    if (!n) return "No label notice: the last build of this workbook renamed no name with label cells left to fix.\n";
    const head = [
      `Labels to fix in ${n.name} after the build's renames. xln never writes cell values: fix them in Excel, save, then pull.`,
      "To copy a Find what or Replace with text: xln: Show labels to fix after a rename (it offers each one).",
      "",
    ];
    return [...head, ...n.notices.flatMap((x) => [...labelNoticeLines(x), ""])].join("\n");
  }

  /**
   * The label notice of the last build of a workbook, in a read-only document; then, unless
   * `copy` is false, a quick pick that copies one Find what or Replace with text.
   */
  async showLabelNotice(arg?: unknown, opts: { copy?: boolean } = {}): Promise<string | undefined> {
    const act = new Activity("labels to fix");
    let key = arg instanceof vscode.Uri ? arg.toString() : typeof arg === "string" ? arg : undefined;
    if (key === undefined) {
      const known = [...this.notices.keys()];
      if (known.length === 0) {
        void act.info("xln: no label notice: no build in this session renamed a name whose label cells still read its old name.");
        return undefined;
      }
      key = known.length === 1 ? known[0]! : (await vscode.window.showQuickPick(known.map((k) => ({ label: this.notices.get(k)!.name, k })), { placeHolder: "Workbook" }))?.k;
      if (key === undefined) return undefined;
    }
    const n = this.notices.get(key);
    const uri = vscode.Uri.from({ scheme: LABELS_SCHEME, path: `/${n?.name ?? "workbook"} labels to fix.txt`, query: key });
    await vscode.window.showTextDocument(await vscode.workspace.openTextDocument(uri), { preview: true });
    act.target = n?.name;
    act.line(n ? `opened: ${n.notices.map((x) => `${x.from} → ${x.to}`).join(", ")}` : "opened: no label notice for this workbook");
    if (opts.copy === false || !n) return uri.toString();
    const items = n.notices.flatMap((x) =>
      x.pairs.flatMap((p) => [
        { label: p.find, description: `Find what (${x.from} → ${x.to})`, text: p.find, box: "Find what" },
        { label: p.replace, description: "Replace with", text: p.replace, box: "Replace with" },
      ]),
    );
    if (items.length === 0) return uri.toString();
    const pick = await vscode.window.showQuickPick(items, { placeHolder: "Copy a text for Excel's Find & Replace (Esc to close)" });
    if (pick) {
      await vscode.env.clipboard.writeText(pick.text);
      act.line(`copied for ${pick.box}: ${pick.text}`);
      vscode.window.setStatusBarMessage(`xln: copied "${pick.text}": paste it in Excel's ${pick.box} box`, 5000);
    }
    return uri.toString();
  }

  /**
   * The one source edit a build makes (2026-10-07): the `@renamed(…)` notes of the renames
   * it has just written, or that an earlier build wrote, are spent, and go. An open
   * document gets a WorkspaceEdit (so its undo history keeps it) and is saved, as the
   * build saved it before planning; a file not open is written on disk. A file changed since the build read it is left alone: the
   * next pull drops the note. Returns the lines for the output.
   */
  private async removeRenamed(root: vscode.Uri, r: BuildResult, read: Readonly<Record<string, string>>): Promise<{ line: string; removed: boolean }[]> {
    if (!r.renamedConsumed?.length || !r.sourceFiles) return [];
    const docs = this.projectDocuments(root);
    const out: { line: string; removed: boolean }[] = [];
    const enc = new TextEncoder();
    for (const [path, text] of Object.entries(r.sourceFiles)) {
      const uri = vscode.Uri.joinPath(root, ...path.split("/"));
      const doc = docs.get(path);
      const now = doc ? doc.getText() : new TextDecoder().decode(await vscode.workspace.fs.readFile(uri));
      const anns = r.renamedConsumed.filter((e) => e.path === path);
      if (now !== read[path]) {
        out.push({ line: `left ${anns.map((e) => e.annotation).join(", ")} in ${path}: the file changed during the build (the next pull drops it)`, removed: false });
        continue;
      }
      if (doc) {
        const edit = new vscode.WorkspaceEdit();
        for (const e of anns) edit.delete(uri, new vscode.Range(doc.positionAt(e.start), doc.positionAt(e.end)));
        if (!(await vscode.workspace.applyEdit(edit)) || !(await doc.save())) {
          out.push({ line: `could not remove ${anns.map((e) => e.annotation).join(", ")} from ${path}: remove it by hand (or pull)`, removed: false });
          continue;
        }
      } else await vscode.workspace.fs.writeFile(uri, enc.encode(text));
      out.push({ line: renamedConsumedLine(path, anns.map((e) => e.annotation)), removed: true });
    }
    return out;
  }

  /** The build's blocking errors in the Problems panel, at their file:line (replacing the last build's for this project). */
  private publishDiagnostics(handle: ProjectHandle, r: BuildResult): void {
    const root = handle.uri.toString();
    this.diagnostics.forEach((uri) => {
      if (uri.toString().startsWith(root + "/")) this.diagnostics.delete(uri);
    });
    const byFile = new Map<string, vscode.Diagnostic[]>();
    for (const d of buildDiagnostics(r)) {
      const diag = new vscode.Diagnostic(new vscode.Range(d.line, 0, d.line, Number.MAX_SAFE_INTEGER), `build refused: ${d.message}`, vscode.DiagnosticSeverity.Error);
      diag.source = "xln build";
      diag.code = d.code;
      const l = byFile.get(d.file) ?? [];
      l.push(diag);
      byFile.set(d.file, l);
    }
    for (const [file, list] of byFile) this.diagnostics.set(vscode.Uri.joinPath(handle.uri, ...file.split("/")), list);
  }

  /**
   * A refusal cannot be missed (feedback 2026-10-07): a modal with the reasons, *Show
   * details* opening the output. E2: conflicts are never merged; each can be opened as a
   * diff, Excel's version left.
   */
  private async showRefusal(handle: ProjectHandle, workbook: string, r: BuildResult, modal: boolean): Promise<void> {
    const n = r.plan.conflicts.length;
    const actions = ["Show details", ...(n ? ["Show conflicts"] : [])];
    const [head, ...rest] = refusalMessage(workbook, r).split("\n\n");
    const choice = modal
      ? await vscode.window.showErrorMessage(`xln: ${head}`, { modal: true, detail: rest.join("\n\n") }, ...actions)
      : await vscode.window.showErrorMessage(`xln: ${head} ${refusalReasons(r)[0] ?? ""}`, ...actions);
    if (choice === "Show conflicts") {
      for (const c of r.plan.conflicts) await this.showConflict(handle, c);
    } else if (choice === "Show details") {
      this.showOutput();
    }
  }

  async showConflict(handle: ProjectHandle, c: Conflict): Promise<void> {
    const { excel, source } = conflictTexts(c);
    const key = encodeURIComponent(c.key);
    const left = vscode.Uri.from({ scheme: CONFLICT_SCHEME, path: `/${key}.excel.xln`, query: handle.uri.toString() });
    const right = vscode.Uri.from({ scheme: CONFLICT_SCHEME, path: `/${key}.source.xln`, query: handle.uri.toString() });
    this.texts.set(left.toString(), excel);
    this.texts.set(right.toString(), source);
    await vscode.commands.executeCommand("vscode.diff", left, right, `${c.key}: Excel ↔ source (conflict)`);
  }
}

/** The label notices of a build's renames, on the workbook before the build (the build leaves its labels as they are). Advice only: a workbook it cannot read gives none. */
function labelNoticesOf(original: Uint8Array, changes: readonly Change[]): LabelNotice[] {
  if (!changes.some((c) => c.op === "rename-name")) return [];
  const values = () => {
    try {
      return cellValueMap(original);
    } catch {
      return undefined;
    }
  };
  try {
    return renameLabelNotices(readWorkbook(original), values, changes);
  } catch {
    return [];
  }
}
