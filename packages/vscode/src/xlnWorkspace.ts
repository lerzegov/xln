// Finds and loads every xln project folder in the workspace through vscode.workspace.fs,
// the only file API that reaches a local folder in vscode.dev. There, a local folder has
// no file watching, so besides the watcher the model is rebuilt on save of an `.xln`
// file, after a pull, and on `xln: Reload`; open editors feed their unsaved text in.

import * as vscode from "vscode";
import { cellValueMap, CONFIG_FILE, LOCK_FILE, MANIFEST_FILE, parseConfig, parseLockfile, readWorkbook, type AuditSettings, type CellValue, type Lockfile, type WorkbookSnapshot } from "@xln/core";
import { parseManifest } from "./model/manifest.js";
import { isNamesFile, NAMES_DIR, Project } from "./model/project.js";
import { isProjectListing } from "./model/pull.js";
import { rootOf } from "./model/ux.js";

export interface ProjectHandle {
  uri: vscode.Uri;
  project: Project;
  /** The workbook beside the project folder, from the manifest. */
  workbookUri: vscode.Uri | undefined;
  /** The audit settings of `xln.config.json` (harness, rule severities, constants), if the project has one. */
  audit: AuditSettings | undefined;
  /** M4: `library` of `xln.config.json`, as written (library.ts resolves it). */
  library: string | undefined;
  /** Files below `names/` that are not `.xln` (hidden ones aside): the checker flags them (M3e). */
  strays: string[];
  /** `xln.config.json` as read, for the marks on what in it cannot be used. */
  configText?: string;
  /** Problems met while loading (unreadable manifest, unreadable file, unusable settings). */
  loadErrors: string[];
}

export interface LoadStats {
  projects: number;
  files: number;
  names: number;
  ms: number;
}

const SKIP = new Set(["node_modules", ".git", ".vscode-test", ".vscode-test-web", "dist"]);
const MAX_DEPTH = 6;

function parentOf(uri: vscode.Uri): vscode.Uri {
  const p = uri.path.replace(/\/+$/, "");
  return uri.with({ path: p.slice(0, p.lastIndexOf("/")) || "/" });
}

export class XlnWorkspace implements vscode.Disposable {
  projects: ProjectHandle[] = [];
  /** `edit`: an open editor's text changed (typing); otherwise files were (re)loaded. */
  private readonly changed = new vscode.EventEmitter<{ edit: boolean }>();
  readonly onDidChange = this.changed.event;
  private loading: Promise<LoadStats> | undefined;
  private pending: ReturnType<typeof setTimeout> | undefined;
  private readonly snapshots = new Map<string, { stamp: string; snapshot: Promise<WorkbookSnapshot | undefined> }>();
  /** The cells' values of a snapshot read here, computed on first use (the audit's C15 reads labels). */
  private readonly cellValues = new WeakMap<WorkbookSnapshot, () => ReadonlyMap<string, ReadonlyMap<string, CellValue>> | undefined>();
  private readonly disposables: vscode.Disposable[] = [];

  constructor(private readonly log: (line: string) => void) {
    const watcher = vscode.workspace.createFileSystemWatcher(`**/{*.xln,${MANIFEST_FILE},${LOCK_FILE},${CONFIG_FILE}}`);
    this.disposables.push(
      this.changed,
      watcher,
      watcher.onDidCreate(() => this.scheduleReload()),
      watcher.onDidDelete(() => this.scheduleReload()),
      watcher.onDidChange((uri) => (uri.path.endsWith(".xln") ? this.refreshFile(uri) : this.scheduleReload())),
      vscode.workspace.onDidSaveTextDocument((d) => {
        if (d.uri.path.endsWith(".xln") || d.uri.path.endsWith(MANIFEST_FILE) || d.uri.path.endsWith(`/${CONFIG_FILE}`)) this.scheduleReload();
      }),
      vscode.workspace.onDidChangeTextDocument((e) => this.syncDocument(e.document)),
      vscode.workspace.onDidOpenTextDocument((d) => this.syncDocument(d)),
      vscode.workspace.onDidChangeWorkspaceFolders(() => this.scheduleReload()),
    );
  }

  dispose(): void {
    if (this.pending) clearTimeout(this.pending);
    for (const d of this.disposables) d.dispose();
  }

  /** The first load, or the one running now. */
  ready(): Promise<LoadStats> {
    return this.loading ?? this.reload();
  }

  scheduleReload(): void {
    if (this.pending) clearTimeout(this.pending);
    this.pending = setTimeout(() => {
      this.pending = undefined;
      void this.reload();
    }, 300);
  }

  /** The counts of the last load the log reported (see `reload`). */
  private logged: string | undefined;

  /**
   * Re-reads every project folder in the workspace. The log line: a load follows every save
   * of an .xln file, a pull, a build and a new module, so a line for each would bury the
   * actions. It is written when the counts (projects, files, names) differ from the last
   * line written, when the load took over 200 ms (worth knowing), when a project has load
   * errors, and always for `xln: Reload` (`report`, the user asked).
   */
  reload(report = false): Promise<LoadStats> {
    const run = (async (): Promise<LoadStats> => {
      const t0 = Date.now();
      const found: vscode.Uri[] = [];
      for (const folder of vscode.workspace.workspaceFolders ?? []) await this.findProjects(folder.uri, 0, found);
      const handles = await Promise.all(found.map((uri) => this.loadProject(uri)));
      this.projects = handles;
      for (const doc of vscode.workspace.textDocuments) {
        // A document VS Code still holds for a file deleted on disk does not bring it back.
        const at = this.locate(doc.uri);
        if (at && !at.handle.project.files.has(at.path) && !doc.isDirty) continue;
        this.syncDocument(doc, false);
      }
      const stats: LoadStats = {
        projects: handles.length,
        files: handles.reduce((a, h) => a + h.project.files.size, 0),
        names: handles.reduce((a, h) => a + h.project.defs.length, 0),
        ms: Date.now() - t0,
      };
      const counts = `${stats.projects} project(s), ${stats.files} .xln files, ${stats.names} names`;
      const errors = handles.some((h) => h.loadErrors.length > 0);
      if (report || errors || stats.ms > 200 || counts !== this.logged) {
        this.logged = counts;
        this.log("");
        this.log(`xln ${report ? "reload" : "loaded"} ${counts} in ${stats.ms} ms`);
        for (const h of handles) for (const e of h.loadErrors) this.log(`  ${vscode.workspace.asRelativePath(h.uri)}: ${e}`);
      }
      this.changed.fire({ edit: false });
      return stats;
    })();
    this.loading = run;
    return run;
  }

  private async findProjects(dir: vscode.Uri, depth: number, found: vscode.Uri[]): Promise<void> {
    if (depth > MAX_DEPTH) return;
    let entries: [string, vscode.FileType][];
    try {
      entries = await vscode.workspace.fs.readDirectory(dir);
    } catch {
      return;
    }
    const listing = entries.map(([name, type]) => ({ name, isDirectory: (type & vscode.FileType.Directory) !== 0 }));
    if (isProjectListing(listing)) {
      found.push(dir);
      return;
    }
    await Promise.all(
      listing.filter((e) => e.isDirectory && !SKIP.has(e.name)).map((e) => this.findProjects(vscode.Uri.joinPath(dir, e.name), depth + 1, found)),
    );
  }

  private async loadProject(uri: vscode.Uri): Promise<ProjectHandle> {
    const loadErrors: string[] = [];
    let manifest;
    try {
      manifest = parseManifest(new TextDecoder().decode(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(uri, MANIFEST_FILE))));
    } catch (err) {
      loadErrors.push(`${MANIFEST_FILE}: ${err instanceof Error ? err.message : String(err)}`);
    }
    let audit: AuditSettings | undefined;
    let library: string | undefined;
    let configText: string | undefined;
    try {
      configText = new TextDecoder().decode(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(uri, CONFIG_FILE)));
    } catch {
      // no settings: the audit's defaults
    }
    if (configText !== undefined) {
      const { config, problems, notes } = parseConfig(configText);
      audit = config.audit;
      library = config.library;
      for (const p of [...problems, ...notes]) loadErrors.push(`${CONFIG_FILE}: ${p}`);
    }
    let lock: Lockfile | undefined;
    try {
      lock = parseLockfile(new TextDecoder().decode(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(uri, LOCK_FILE))));
    } catch {
      // no lockfile, or not one: the build reports it
    }
    const project = new Project(uri.toString(), manifest);
    // The checker compares with the last pull (addresses, scopes).
    project.lock = lock;
    const paths: string[] = [];
    const strays: string[] = [];
    await this.listXln(vscode.Uri.joinPath(uri, NAMES_DIR), NAMES_DIR, paths, 0, strays);
    const texts = await Promise.all(
      paths.map(async (p) => {
        try {
          return new TextDecoder().decode(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(uri, ...p.split("/"))));
        } catch (err) {
          loadErrors.push(`${p}: ${err instanceof Error ? err.message : String(err)}`);
          return undefined;
        }
      }),
    );
    paths.forEach((p, k) => {
      if (texts[k] !== undefined) project.setFile(p, texts[k]!);
    });
    const workbookUri = manifest?.workbook ? vscode.Uri.joinPath(parentOf(uri), manifest.workbook) : undefined;
    return { uri, project, workbookUri, audit, library, strays, loadErrors, ...(configText !== undefined ? { configText } : {}) };
  }

  private async listXln(dir: vscode.Uri, rel: string, out: string[], depth: number, strays: string[]): Promise<void> {
    if (depth > MAX_DEPTH) return;
    let entries: [string, vscode.FileType][];
    try {
      entries = await vscode.workspace.fs.readDirectory(dir);
    } catch {
      return;
    }
    for (const [name, type] of entries) {
      const path = `${rel}/${name}`;
      if (type & vscode.FileType.Directory) await this.listXln(vscode.Uri.joinPath(dir, name), path, out, depth + 1, strays);
      else if (isNamesFile(path)) out.push(path);
      else if (!name.startsWith(".")) strays.push(path);
    }
  }

  /** The project holding `uri`, and the file's project-relative path. */
  locate(uri: vscode.Uri): { handle: ProjectHandle; path: string } | undefined {
    const s = uri.toString();
    for (const handle of this.projects) {
      const root = handle.uri.toString() + "/";
      if (s.startsWith(root)) return { handle, path: decodeURIComponent(s.slice(root.length)) };
    }
    return undefined;
  }

  /** The project at `uri`: the project folder itself or anything below it (an Explorer item). */
  handleAt(uri: vscode.Uri): ProjectHandle | undefined {
    const root = rootOf(
      uri.toString(),
      this.projects.map((h) => h.uri.toString()),
    );
    return root === undefined ? undefined : this.handleFor(root);
  }

  fileUri(handle: ProjectHandle, path: string): vscode.Uri {
    return vscode.Uri.joinPath(handle.uri, ...path.split("/"));
  }

  handleFor(root: string): ProjectHandle | undefined {
    return this.projects.find((h) => h.uri.toString() === root);
  }

  /** Feeds an open editor's text in now: a provider may run before the change event. */
  sync(doc: vscode.TextDocument): void {
    this.syncDocument(doc, false);
  }

  /** An open editor's text replaces the file's: features follow unsaved edits. */
  private syncDocument(doc: vscode.TextDocument, fire = true): void {
    if (doc.languageId !== "xln" && !doc.uri.path.endsWith(".xln")) return;
    const at = this.locate(doc.uri);
    if (!at || !isNamesFile(at.path)) return;
    at.handle.project.setFile(at.path, doc.getText());
    if (fire) this.changed.fire({ edit: true });
  }

  private async refreshFile(uri: vscode.Uri): Promise<void> {
    const at = this.locate(uri);
    if (!at) return this.scheduleReload();
    if (vscode.workspace.textDocuments.some((d) => d.uri.toString() === uri.toString() && d.isDirty)) return;
    try {
      at.handle.project.setFile(at.path, new TextDecoder().decode(await vscode.workspace.fs.readFile(uri)));
      this.changed.fire({ edit: false });
    } catch {
      this.scheduleReload();
    }
  }

  /**
   * The workbook the project was pulled from, read lazily and kept while its size and
   * modification time stay the same. Undefined when it is missing or unreadable: the
   * features then work from the manifest alone.
   */
  async workbook(handle: ProjectHandle): Promise<WorkbookSnapshot | undefined> {
    return handle.workbookUri ? this.workbookAt(handle.workbookUri) : undefined;
  }

  /** The project pulled from the workbook at `uri`, if one is loaded. */
  handleForWorkbook(uri: vscode.Uri): ProjectHandle | undefined {
    const s = uri.toString();
    return this.projects.find((h) => h.workbookUri?.toString() === s);
  }

  /** Any workbook, read lazily and cached the same way (the formula view of a workbook never pulled). */
  async workbookAt(uri: vscode.Uri): Promise<WorkbookSnapshot | undefined> {
    const key = uri.toString();
    let stamp: string;
    try {
      const st = await vscode.workspace.fs.stat(uri);
      stamp = `${st.mtime}:${st.size}`;
    } catch (err) {
      this.noteReadError(uri, "stat", err);
      return undefined;
    }
    const hit = this.snapshots.get(key);
    if (hit && hit.stamp === stamp) return hit.snapshot;
    const snapshot = (async () => {
      try {
        const t0 = Date.now();
        let bytes: Uint8Array | undefined = await vscode.workspace.fs.readFile(uri);
        const wb = readWorkbook(bytes);
        let values: ReadonlyMap<string, ReadonlyMap<string, CellValue>> | undefined;
        this.cellValues.set(wb, () => {
          if (bytes !== undefined) {
            try {
              values = cellValueMap(bytes);
            } catch {
              // Without the values C15 says nothing; the rest of the audit stands.
            }
            bytes = undefined;
          }
          return values;
        });
        // A read follows every change of the workbook (a build, a save in Excel): logged when
        // it is slow enough to notice, as the loads are.
        const ms = Date.now() - t0;
        if (ms > 200) {
          this.log("");
          this.log(`xln read ${vscode.workspace.asRelativePath(uri)}: formulas and cached values in ${ms} ms`);
        }
        this.readErrors.delete(key);
        return wb;
      } catch (err) {
        // A failure is not cached: the next request tries again (permission granted, Excel closed...).
        this.snapshots.delete(key);
        this.noteReadError(uri, "read", err);
        return undefined;
      }
    })();
    this.snapshots.set(key, { stamp, snapshot });
    return snapshot;
  }

  /** The cells' values of a snapshot `workbookAt` returned (sheet → address → value); undefined when they cannot be read. */
  valuesOf(wb: WorkbookSnapshot): ReadonlyMap<string, ReadonlyMap<string, CellValue>> | undefined {
    return this.cellValues.get(wb)?.();
  }

  /** Why the last `workbookAt(uri)` failed, for the error message shown to the user. */
  readError(uri: vscode.Uri): string | undefined {
    return this.readErrors.get(uri.toString());
  }

  private readonly readErrors = new Map<string, string>();

  private noteReadError(uri: vscode.Uri, step: "stat" | "read", err: unknown): void {
    const why = `${step} failed: ${err instanceof Error ? `${err.name}: ${err.message}` : String(err)}`;
    this.readErrors.set(uri.toString(), why);
    this.log("");
    this.log(`xln read ${vscode.workspace.asRelativePath(uri)}: could not read it (${why})`);
  }
}
