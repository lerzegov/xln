// M3e, the editor's housekeeping around the projects: the `xln.hasProject` context key
// (the Explorer's (+) New module button), tabs of files a pull removed, and files that
// have no place in a project (VS Code cannot refuse a file created in the Explorer: the
// user is told at once, the checker flags it and the build refuses it).

import * as vscode from "vscode";
import { strayFile } from "@xln/core";
import { internalsExclude, internalsHidden, staleTabs } from "./model/ux.js";
import { Activity } from "./log.js";
import type { XlnWorkspace } from "./xlnWorkspace.js";

async function exists(uri: vscode.Uri): Promise<boolean> {
  try {
    await vscode.workspace.fs.stat(uri);
    return true;
  } catch {
    return false;
  }
}

/**
 * Closes the editor tabs of files below `roots` (project folders) that no longer exist.
 * A tab with unsaved text is left open (its text is nobody else's); with `warn` the user is
 * told. Returns the URIs closed and kept.
 */
export async function closeStaleTabs(roots: readonly vscode.Uri[], warn: boolean): Promise<{ closed: string[]; kept: string[] }> {
  const tabs = new Map<string, vscode.Tab[]>();
  for (const group of vscode.window.tabGroups.all) {
    for (const tab of group.tabs) {
      if (!(tab.input instanceof vscode.TabInputText)) continue;
      const k = tab.input.uri.toString();
      tabs.set(k, [...(tabs.get(k) ?? []), tab]);
    }
  }
  const rootStrings = roots.map((r) => r.toString());
  const missing = new Set<string>();
  for (const [k, list] of tabs) {
    const uri = (list[0]!.input as vscode.TabInputText).uri;
    if (rootStrings.some((r) => k.startsWith(r + "/")) && !(await exists(uri))) missing.add(k);
  }
  const { close, keep } = staleTabs(
    [...tabs].map(([uri, list]) => ({ uri, dirty: list.some((t) => t.isDirty) })),
    rootStrings,
    (u) => !missing.has(u),
  );
  if (close.length) await vscode.window.tabGroups.close(close.flatMap((u) => tabs.get(u)!));
  if (keep.length && warn) {
    const names = keep.map((u) => vscode.workspace.asRelativePath(vscode.Uri.parse(u))).join(", ");
    void new Activity("tabs").warn(`xln: ${names} no longer exist${keep.length === 1 ? "s" : ""} but ha${keep.length === 1 ? "s" : "ve"} unsaved changes: left open. Save elsewhere or close ${keep.length === 1 ? "it" : "them"}.`);
  }
  return { closed: close, kept: keep };
}

const DELETE = "Delete this file";

export class UxFeature {
  constructor(private readonly ws: XlnWorkspace) {}

  register(context: vscode.ExtensionContext): void {
    context.subscriptions.push(
      this.ws.onDidChange((e) => !e.edit && this.loaded()),
      vscode.workspace.onDidCreateFiles((e) => void this.arrived(e.files)),
      vscode.workspace.onDidRenameFiles((e) => void this.arrived(e.files.map((f) => f.newUri))),
      vscode.commands.registerCommand("xln.showInternals", () => this.internals(true)),
      vscode.commands.registerCommand("xln.hideInternals", () => this.internals(false)),
      vscode.workspace.onDidChangeConfiguration((e) => e.affectsConfiguration("files.exclude") && this.internalsKey()),
      vscode.languages.registerCodeActionsProvider(
        { pattern: "**/names/**" },
        { provideCodeActions: (doc, _range, ctx) => this.deleteActions(doc.uri, ctx.diagnostics) },
        { providedCodeActionKinds: [vscode.CodeActionKind.QuickFix] },
      ),
    );
    this.internalsKey();
  }

  /**
   * `xln: Show project internals` / `Hide project internals`: the manifest and the lockfile
   * in the Explorer or not, for this workspace only (its settings are written only here,
   * when the user asks). Read-only in the editor either way (`files.readonlyInclude`): xln
   * writes them through workspace.fs, which that setting does not restrict. Returns
   * whether they are hidden now.
   */
  async internals(show: boolean): Promise<boolean> {
    const cfg = vscode.workspace.getConfiguration("files");
    const { changed, value } = internalsExclude(cfg.inspect<Record<string, unknown>>("exclude")?.workspaceValue, show);
    if (changed) await cfg.update("exclude", value, vscode.ConfigurationTarget.Workspace);
    new Activity(show ? "show internals" : "hide internals").line(
      changed ? `files.exclude of this workspace updated: the manifest and the lockfile are ${show ? "shown" : "hidden"} in the Explorer` : `nothing to change: already ${show ? "shown" : "hidden"}`,
    );
    return this.internalsKey();
  }

  /** The `xln.internalsShown` context key (which of the two commands the palette offers). */
  private internalsKey(): boolean {
    const hidden = internalsHidden(vscode.workspace.getConfiguration("files").get<Record<string, unknown>>("exclude"));
    void vscode.commands.executeCommand("setContext", "xln.internalsShown", !hidden);
    return hidden;
  }

  /** After every load of the projects. */
  private loaded(): void {
    void vscode.commands.executeCommand("setContext", "xln.hasProject", this.ws.projects.length > 0);
    // Tabs left over from files removed while VS Code was closed, or by a pull elsewhere.
    if (this.ws.projects.length) void closeStaleTabs(this.ws.projects.map((h) => h.uri), false);
  }

  /**
   * Files created or renamed in the editor: one that has no place in its project gets a
   * warning now (the error diagnostic follows with the reload). Returns the messages.
   */
  async arrived(uris: readonly vscode.Uri[]): Promise<string[]> {
    await this.ws.ready();
    const out: string[] = [];
    let reload = false;
    for (const uri of uris) {
      const handle = this.ws.handleAt(uri);
      if (!handle) continue;
      const path = decodeURIComponent(uri.toString().slice(handle.uri.toString().length + 1));
      if (!path.startsWith("names/")) continue;
      reload = true;
      try {
        // A folder alone is harmless: the files put in it are judged.
        if ((await vscode.workspace.fs.stat(uri)).type & vscode.FileType.Directory) continue;
      } catch {
        continue;
      }
      const message = strayFile(path, handle.project.checkContext().sheets);
      if (message === undefined) continue;
      out.push(message);
      void new Activity("stray file", vscode.workspace.asRelativePath(uri)).warn(`xln: ${message}`, DELETE).then(async (c) => {
        if (c !== DELETE) return;
        await vscode.workspace.fs.delete(uri, { useTrash: true });
        new Activity("stray file", vscode.workspace.asRelativePath(uri)).line("deleted (to the trash)");
      });
    }
    // The watcher sees .xln files only; vscode.dev sees nothing.
    if (reload) this.ws.scheduleReload();
    return out;
  }

  /** "Delete this file" on a file the checker says has no place in the project. */
  private deleteActions(uri: vscode.Uri, diagnostics: readonly vscode.Diagnostic[]): vscode.CodeAction[] {
    const d = diagnostics.find((x) => x.source === "xln" && x.code === "stray-file");
    if (!d) return [];
    const a = new vscode.CodeAction(DELETE, vscode.CodeActionKind.QuickFix);
    a.edit = new vscode.WorkspaceEdit();
    a.edit.deleteFile(uri, { ignoreIfNotExists: true });
    a.diagnostics = [d];
    return [a];
  }
}
