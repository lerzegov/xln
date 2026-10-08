// The library in the editor (M4). Nothing moves between library, project and workbook
// without an explicit action, and each one is a source edit the user sees or a write the
// user confirms:
//
// - `xln: Library status`: a read-only report document `xln-lib:/<project> (library)`,
//   rendered by the core like the CLI's (names link to the project or the library file).
//   A document rather than a tree view: the diffs are text and read best as text, the
//   audit report already works this way (links, refresh, vscode.dev), and it costs no
//   view container next to the menus M3e is rearranging.
// - Completion offers library functions the project lacks ("from library"); accepting
//   one inserts the call and adds the definitions to `names/FN.xln` (created if missing)
//   as an unsaved edit. `xln: Insert library function` does the same from a quick pick.
// - Code lenses (and quick fixes) on module entries: the entry's library state, three-way
//   on its library base `@from(lib #…)` (model/library.ts `entryLibraryActions`);
//   "Take the library's version" (replaces the definition and doc comment in the source
//   and records the base; on a copy edited since its base it first asks, naming the edit
//   it discards); "Show diff with library" (three-way for "both changed"); "Publish to
//   library" (diff shown, then a modal confirmation before the `.lambda` file is written;
//   the entry then records the published version as its base, an unsaved edit); "Record
//   library base" on a copy identical to the library that records none (an unsaved edit,
//   never done on its own).
// - Insert, Take, Publish and Record library base keep the base's text in the project's
//   `library-bases/<hash>.json` (written at once: internal, hidden and read-only like the
//   lockfile), so "both changed" diffs both sides from the base. A pull leaves the folder
//   alone; nothing prunes it.
//
// The library folder is `library` in the project's `xln.config.json`; it is read through
// vscode.workspace.fs (so a folder inside the opened workspace also works in vscode.dev),
// on first use and again after a reload, a publish, or `xln: Library status`.

import * as vscode from "vscode";
import {
  applyEdits,
  comparableLayout,
  findEntry,
  isLambdaFile,
  libraryInsertion,
  libraryReplacement,
  libraryBaseEdit,
  publishedBase,
  threeWayText,
  publishLambda,
  readLibrary,
  renderLibStatus,
  libStateLabel,
  baseFiles,
  BASES_DIR,
  isBasePath,
  libraryBaseRecordings,
  libraryFunctionBase,
  readBases,
  type LibraryBase,
  type Library,
  type LibStatusItem,
  type LibStatusReport,
  type ModuleEntry,
} from "@xln/core";
import { homeDir } from "./excelHost.js";
import { formulaSite } from "./model/editor.js";
import { definesWorkbookName, entryArg, entryLibraryActions, entryStates, isModuleFile, libraryLocation, projectFiles, projectLibraryStatus, publishWarning, tagWarning } from "./model/library.js";
import { baseName } from "./inspect.js";
import { Activity } from "./log.js";
import type { ProjectHandle, XlnWorkspace } from "./xlnWorkspace.js";

export const LIB_SCHEME = "xln-lib";
const REPORT_SUFFIX = " (library)";
const SELECTOR: vscode.DocumentSelector = [{ language: "xln" }];

interface Loaded {
  uri: vscode.Uri;
  library: Library;
}

/** What tests get back from the commands. */
export interface InsertResult {
  names: string[];
  files: string[];
  /** A doc comment inserted leaves no room for the provenance tag (`tagWarning`), as shown. */
  warning?: string;
}

export class LibraryFeature {
  private readonly libraries = new Map<string, Promise<Loaded | { error: string }>>();
  /** The libraries read already, for completion (which cannot wait). */
  private readonly ready = new Map<string, Library>();
  private readonly lensesChanged = new vscode.EventEmitter<void>();
  private readonly docsChanged = new vscode.EventEmitter<vscode.Uri>();
  /** Texts of the virtual documents (diff sides, proposed library files), by URI. */
  private readonly virtual = new Map<string, string>();
  private readonly reports = new Map<string, LibStatusReport>();
  /** Each project's kept library bases (`library-bases/`), read once until a reload or a write. */
  private readonly bases = new Map<string, Promise<Map<string, LibraryBase>>>();

  constructor(private readonly ws: XlnWorkspace) {}

  register(context: vscode.ExtensionContext): void {
    context.subscriptions.push(
      this.lensesChanged,
      this.docsChanged,
      vscode.workspace.registerTextDocumentContentProvider(LIB_SCHEME, { onDidChange: this.docsChanged.event, provideTextDocumentContent: (uri) => this.text(uri) }),
      vscode.languages.registerDocumentLinkProvider({ scheme: LIB_SCHEME }, { provideDocumentLinks: (d) => this.links(d) }),
      vscode.languages.registerCodeLensProvider(SELECTOR, { onDidChangeCodeLenses: this.lensesChanged.event, provideCodeLenses: (d) => this.lenses(d) }),
      vscode.languages.registerCodeActionsProvider(SELECTOR, { provideCodeActions: (d, r) => this.actions(d, r) }, { providedCodeActionKinds: [vscode.CodeActionKind.QuickFix] }),
      vscode.commands.registerCommand("xln.libraryStatus", (arg?: unknown) => this.status(arg)),
      vscode.commands.registerCommand("xln.insertLibraryFunction", (arg?: unknown) => this.insertCommand(arg)),
      vscode.commands.registerCommand("xln.insertLibraryDefinition", (arg?: { root: string; name: string }) => this.insertDefinition(arg)),
      // From a code lens or quick fix the argument names the entry; from a menu VS Code passes
      // the document's Uri, which says nothing about the entry: then the cursor decides.
      vscode.commands.registerCommand("xln.takeLibraryVersion", (arg?: unknown) => this.take(entryArg<{ root: string; path: string; name: string }>(arg) ?? this.atCursor())),
      vscode.commands.registerCommand("xln.showLibraryDiff", (arg?: unknown) => this.diff(entryArg<{ root: string; name: string }>(arg) ?? this.atCursor())),
      vscode.commands.registerCommand("xln.publishToLibrary", (arg?: unknown) => this.publish(entryArg<{ root: string; name: string; confirm?: boolean }>(arg) ?? this.atCursor())),
      vscode.commands.registerCommand("xln.recordLibraryBase", (arg?: unknown) => this.recordLibraryBase(entryArg<{ root: string; path: string; name: string }>(arg) ?? this.atCursor())),
      this.ws.onDidChange((e) => {
        if (!e.edit) {
          this.libraries.clear();
          this.ready.clear();
          this.bases.clear();
        }
        this.lensesChanged.fire();
      }),
    );
  }

  // ---- the library of a project ---------------------------------------------------------

  /** The library folder of a project, or why there is none. */
  libraryUri(handle: ProjectHandle): vscode.Uri | { error: string } {
    if (handle.library === undefined) return { error: `no library: set "library" in ${vscode.workspace.asRelativePath(vscode.Uri.joinPath(handle.uri, "xln.config.json"))} (a folder of .lambda files, relative to the project folder)` };
    const loc = libraryLocation(handle.library, homeDir());
    if ("error" in loc) return loc;
    if ("absolute" in loc) return handle.uri.with({ path: loc.absolute });
    return vscode.Uri.joinPath(handle.uri, ...loc.relative.split("/").filter((s) => s !== "" && s !== "."));
  }

  /** The project's library, read once and kept until a reload. */
  load(handle: ProjectHandle, fresh = false): Promise<Loaded | { error: string }> {
    const key = handle.uri.toString();
    if (fresh) {
      this.libraries.delete(key);
      this.ready.delete(key);
    }
    let p = this.libraries.get(key);
    if (!p) {
      p = (async () => {
        const uri = this.libraryUri(handle);
        if (!(uri instanceof vscode.Uri)) return uri;
        let entries: [string, vscode.FileType][];
        try {
          entries = await vscode.workspace.fs.readDirectory(uri);
        } catch (err) {
          return { error: `cannot read the library folder ${uri.toString(true)}: ${err instanceof Error ? err.message : String(err)}` };
        }
        const files: Record<string, string> = {};
        await Promise.all(
          entries
            .filter(([n, t]) => isLambdaFile(n) && (t & vscode.FileType.File) !== 0)
            .map(async ([n]) => {
              files[n] = new TextDecoder().decode(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(uri, n)));
            }),
        );
        const library = readLibrary(files);
        this.ready.set(key, library);
        return { uri, library };
      })();
      this.libraries.set(key, p);
      void p.then(() => this.lensesChanged.fire());
    }
    return p;
  }

  /** The library if it is loaded already (completion cannot wait); starts loading it otherwise. */
  loadedLibrary(handle: ProjectHandle | undefined): Library | undefined {
    if (!handle || handle.library === undefined) return undefined;
    const key = handle.uri.toString();
    if (!this.libraries.has(key)) void this.load(handle);
    return this.ready.get(key);
  }

  private async loaded(handle: ProjectHandle, fresh = false): Promise<Loaded> {
    const r = await this.load(handle, fresh);
    if ("error" in r) throw new Error(r.error);
    return r;
  }

  private async report(handle: ProjectHandle, fresh = false): Promise<{ report: LibStatusReport; lib: Loaded }> {
    const lib = await this.loaded(handle, fresh);
    const wb = await this.ws.workbook(handle);
    const bases = await this.keptBases(handle, fresh);
    const report = projectLibraryStatus(handle.project, lib.library, wb, {
      target: baseName(handle.uri.path),
      library: vscode.workspace.asRelativePath(lib.uri),
      bases,
      ...(handle.workbookUri ? { workbook: baseName(handle.workbookUri.path) } : {}),
    });
    return { report, lib };
  }

  // ---- the kept bases -----------------------------------------------------------------------

  /** The project's `library-bases/`, by hash (empty when it has none). */
  private keptBases(handle: ProjectHandle, fresh = false): Promise<Map<string, LibraryBase>> {
    const key = handle.uri.toString();
    if (fresh) this.bases.delete(key);
    let p = this.bases.get(key);
    if (!p) {
      p = (async () => {
        const dir = vscode.Uri.joinPath(handle.uri, BASES_DIR);
        const files: Record<string, string> = {};
        try {
          for (const [n, t] of await vscode.workspace.fs.readDirectory(dir)) {
            const rel = `${BASES_DIR}/${n}`;
            if ((t & vscode.FileType.File) !== 0 && isBasePath(rel)) files[rel] = new TextDecoder().decode(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(dir, n)));
          }
        } catch {
          // no folder yet
        }
        return readBases(files);
      })();
      this.bases.set(key, p);
    }
    return p;
  }

  /**
   * Keeps the bases' texts the library actions record, in `library-bases/` (a file per
   * version, the same whoever writes it: an existing one is left alone). Written at once,
   * not as an unsaved edit: the folder is internal, and an unreferenced file is harmless.
   */
  private async keepBases(handle: ProjectHandle, bases: readonly LibraryBase[]): Promise<string[]> {
    const written: string[] = [];
    const have = await this.keptBases(handle);
    for (const [rel, text] of Object.entries(baseFiles(bases.filter((b) => !have.has(b.hash))))) {
      const uri = vscode.Uri.joinPath(handle.uri, ...rel.split("/"));
      await vscode.workspace.fs.createDirectory(vscode.Uri.joinPath(handle.uri, BASES_DIR));
      await vscode.workspace.fs.writeFile(uri, new TextEncoder().encode(text));
      written.push(rel);
    }
    if (written.length) this.bases.delete(handle.uri.toString());
    return written;
  }

  private async handleOf(arg: unknown): Promise<ProjectHandle | undefined> {
    await this.ws.ready();
    if (arg && typeof arg === "object" && typeof (arg as { root?: unknown }).root === "string") return this.ws.handleFor((arg as { root: string }).root);
    if (arg instanceof vscode.Uri) return this.ws.handleFor(arg.toString()) ?? this.ws.locate(arg)?.handle ?? this.ws.handleForWorkbook(arg);
    const ed = vscode.window.activeTextEditor?.document.uri;
    if (ed?.scheme === LIB_SCHEME) return this.ws.handleFor(ed.query);
    const h = ed ? this.ws.locate(ed)?.handle : undefined;
    if (h) return h;
    if (this.ws.projects.length === 1) return this.ws.projects[0];
    if (this.ws.projects.length === 0) return undefined;
    const pick = await vscode.window.showQuickPick(
      this.ws.projects.map((p) => ({ label: vscode.workspace.asRelativePath(p.uri), p })),
      { placeHolder: "Project" },
    );
    return pick?.p;
  }

  /** From the command palette: the workbook name whose entry the cursor is in. */
  private atCursor(): { root: string; path: string; name: string } | undefined {
    const ed = vscode.window.activeTextEditor;
    const at = ed ? this.ws.locate(ed.document.uri) : undefined;
    if (!ed || !at) {
      void new Activity("library").info("xln: put the cursor on a name in a module file (names/FN.xln) first.");
      return undefined;
    }
    this.ws.sync(ed.document);
    const offset = ed.document.offsetAt(ed.selection.active);
    const e = at.handle.project.files.get(at.path)?.parsed.entries.find((x) => x.scope === undefined && !x.cell && x.start <= offset && offset <= x.end);
    if (!e) {
      void new Activity("library").info("xln: the cursor is not on a workbook name's definition.");
      return undefined;
    }
    return { root: at.handle.uri.toString(), path: at.path, name: e.name };
  }

  // ---- the report document --------------------------------------------------------------

  private reportUri(handle: ProjectHandle): vscode.Uri {
    return vscode.Uri.from({ scheme: LIB_SCHEME, path: `/${baseName(handle.uri.path)}${REPORT_SUFFIX}`, query: handle.uri.toString() });
  }

  /** `xln: Library status`: opens the report of a project; returns its URI and the report. */
  async status(arg?: unknown): Promise<{ uri: string; report: LibStatusReport } | undefined> {
    const act = new Activity("lib status");
    const handle = await this.handleOf(arg);
    if (!handle) {
      void act.info("xln: pull a workbook first; the library status is a project's.");
      return undefined;
    }
    act.target = vscode.workspace.asRelativePath(handle.uri);
    let r;
    try {
      r = await this.report(handle, true);
    } catch (err) {
      void act.error(`xln: ${err instanceof Error ? err.message : String(err)}`);
      return undefined;
    }
    // The CLI's head (`xln lib status: …` and the counts), the lists stay in the report.
    const [head, ...rest] = renderLibStatus(r.report).text.split("\n\n")[0]!.split("\n");
    act.summary([head!, ...rest.map((l) => `  ${l}`)]);
    const uri = this.reportUri(handle);
    this.reports.set(uri.toString(), r.report);
    this.docsChanged.fire(uri);
    await vscode.window.showTextDocument(uri, { preview: true });
    return { uri: uri.toString(), report: r.report };
  }

  private text(uri: vscode.Uri): string {
    if (uri.path.endsWith(REPORT_SUFFIX)) {
      const r = this.reports.get(uri.toString());
      return r ? renderLibStatus(r).text : "// Run xln: Library status again.\n";
    }
    return this.virtual.get(uri.toString()) ?? "";
  }

  private async links(doc: vscode.TextDocument): Promise<vscode.DocumentLink[]> {
    const r = this.reports.get(doc.uri.toString());
    const handle = this.ws.handleFor(doc.uri.query);
    if (!r || !handle) return [];
    const lib = await this.load(handle);
    // A report refreshed before the document's text: the offsets would clamp to empty ranges,
    // which DocumentLink rejects. VS Code asks again once the content changes.
    const rendered = renderLibStatus(r);
    if (doc.getText() !== rendered.text) return [];
    const out: vscode.DocumentLink[] = [];
    for (const l of rendered.links) {
      if (l.end <= l.start) continue;
      const range = new vscode.Range(doc.positionAt(l.start), doc.positionAt(l.end));
      if (l.target === "library" && l.path && !("error" in lib)) {
        const link = new vscode.DocumentLink(range, vscode.Uri.joinPath(lib.uri, l.path));
        link.tooltip = "Open the library file";
        out.push(link);
      } else if (l.target === "copy") {
        const link = new vscode.DocumentLink(range, commandUri("xln.revealName", { root: handle.uri.toString(), key: l.name }));
        link.tooltip = "Go to the name in the project";
        out.push(link);
      }
    }
    return out;
  }

  // ---- code lenses and quick fixes ----------------------------------------------------------

  /** Each library-related entry of a module file, with its state. */
  private async entries(doc: vscode.TextDocument): Promise<{ handle: ProjectHandle; path: string; items: { entry: ModuleEntry; item: LibStatusItem | undefined }[] } | undefined> {
    this.ws.sync(doc);
    const at = this.ws.locate(doc.uri);
    if (!at || !isModuleFile(at.path) || at.handle.library === undefined) return undefined;
    let r;
    try {
      r = await this.report(at.handle);
    } catch {
      return undefined;
    }
    const states = entryStates(r.report, at.path);
    const file = at.handle.project.files.get(at.path);
    if (!file) return undefined;
    const items = file.parsed.entries.filter((e) => e.scope === undefined && !e.cell).map((entry) => ({ entry, item: states.get(entry.name.toLowerCase()) }));
    return { handle: at.handle, path: at.path, items: items.filter((i) => i.item !== undefined) };
  }

  private async lenses(doc: vscode.TextDocument): Promise<vscode.CodeLens[]> {
    const e = await this.entries(doc);
    if (!e) return [];
    const out: vscode.CodeLens[] = [];
    const root = e.handle.uri.toString();
    for (const { entry, item } of e.items) {
      const pos = doc.positionAt(entry.offset);
      const range = new vscode.Range(pos, pos);
      const a = entryLibraryActions(item!);
      out.push(new vscode.CodeLens(range, { title: a.label, tooltip: a.tooltip, command: a.diff ? "xln.showLibraryDiff" : "", arguments: [{ root, name: entry.name }] }));
      if (a.take) out.push(new vscode.CodeLens(range, { title: a.take.title, command: "xln.takeLibraryVersion", arguments: [{ root, path: e.path, name: entry.name }] }));
      if (a.publish) out.push(new vscode.CodeLens(range, { title: a.publish.title, command: "xln.publishToLibrary", arguments: [{ root, name: entry.name }] }));
      if (a.record) out.push(new vscode.CodeLens(range, { title: a.record.title, tooltip: "Writes @from(lib #…), the library's version this copy is, as an unsaved edit: then a change on either side reads as outdated or modified, not differs", command: "xln.recordLibraryBase", arguments: [{ root, path: e.path, name: entry.name }] }));
    }
    return out;
  }

  private async actions(doc: vscode.TextDocument, range: vscode.Range | vscode.Selection): Promise<vscode.CodeAction[]> {
    const e = await this.entries(doc);
    if (!e) return [];
    const at = doc.offsetAt(range.start);
    const hit = e.items.find((i) => i.entry.start <= at && at <= i.entry.end);
    if (!hit) return [];
    const root = e.handle.uri.toString();
    const out: vscode.CodeAction[] = [];
    const state = hit.item!.state;
    const actions = entryLibraryActions(hit.item!);
    if (actions.diff) {
      const d = new vscode.CodeAction(`Show the diff of ${hit.entry.name} with the library (${libStateLabel(state)})`, vscode.CodeActionKind.QuickFix);
      d.command = { title: d.title, command: "xln.showLibraryDiff", arguments: [{ root, name: hit.entry.name }] };
      out.push(d);
    }
    if (actions.take) {
      const a = new vscode.CodeAction(`Take the library's version of ${hit.entry.name} (${libStateLabel(state)}${actions.take.confirm ? (state === "differs" ? ": replaces this copy" : ": discards your edit") : ""})`, vscode.CodeActionKind.QuickFix);
      a.command = { title: a.title, command: "xln.takeLibraryVersion", arguments: [{ root, path: e.path, name: hit.entry.name }] };
      out.push(a);
    }
    if (actions.publish) {
      const p = new vscode.CodeAction(`Publish ${hit.entry.name} to the library`, vscode.CodeActionKind.QuickFix);
      p.command = { title: p.title, command: "xln.publishToLibrary", arguments: [{ root, name: hit.entry.name }] };
      out.push(p);
    }
    if (actions.record) {
      const r = new vscode.CodeAction(`Record library base of ${hit.entry.name} (identical, no @from)`, vscode.CodeActionKind.QuickFix);
      r.command = { title: r.title, command: "xln.recordLibraryBase", arguments: [{ root, path: e.path, name: hit.entry.name }] };
      out.push(r);
    }
    return out;
  }

  // ---- record library base ----------------------------------------------------------------

  /**
   * "Record library base": `@from(lib #…)` on an entry identical to the library that records
   * none, as an unsaved edit, and the base's text kept. Returns the hash recorded; undefined
   * (with a message saying why) when the entry is not such a copy.
   */
  async recordLibraryBase(arg?: { root: string; path: string; name: string }): Promise<string | undefined> {
    const act = new Activity("lib record base", arg?.name);
    const handle = arg ? this.ws.handleFor(arg.root) : undefined;
    if (!handle || !arg) return noTarget(arg, act), undefined;
    const { library } = await this.loaded(handle);
    // The edit's offsets are the document's: its live text, unsaved edits included.
    const at = this.ws.fileUri(handle, arg.path);
    const open = vscode.workspace.textDocuments.find((d) => d.uri.toString() === at.toString());
    if (open) this.ws.sync(open);
    const r = libraryBaseRecordings(library, projectFiles(handle.project), [arg.name]);
    const x = r.recorded[0];
    if (!x) {
      void act.info(`xln: ${arg.name}: ${r.skipped[0]?.reason ?? "nothing to record"}.`);
      return undefined;
    }
    const uri = this.ws.fileUri(handle, x.path);
    const doc = await vscode.workspace.openTextDocument(uri);
    if (doc.getText() !== projectFiles(handle.project)[x.path]) return void act.warn(`xln: ${x.path} changed meanwhile; try again.`), undefined;
    const edit = new vscode.WorkspaceEdit();
    edit.replace(uri, new vscode.Range(doc.positionAt(x.edit.start), doc.positionAt(x.edit.end)), x.edit.text);
    const warning = tagWarning(projectFiles(handle.project), { [x.path]: applyEdits(doc.getText(), [x.edit]) }, [x.name]);
    if (!(await vscode.workspace.applyEdit(edit))) return undefined;
    await this.keepBases(handle, [x.base]);
    this.lensesChanged.fire();
    void act.info(`xln: recorded @from(lib #${x.hash}) on ${x.name} (not saved; build to carry it into the workbook).`);
    if (warning) void act.warn(`xln: ${warning}`);
    return x.hash;
  }

  // ---- insert -------------------------------------------------------------------------------

  /**
   * Adds library function `name` (and the library functions it calls that the project
   * lacks) to the project's source, as an edit of the module file (created if missing)
   * the user sees and then builds. Completion runs it after inserting the call.
   */
  async insertDefinition(arg?: { root: string; name: string }): Promise<InsertResult | undefined> {
    const act = new Activity("lib insert", arg?.name);
    const handle = arg ? this.ws.handleFor(arg.root) : undefined;
    if (!handle || !arg) return undefined;
    const { library } = await this.loaded(handle);
    const files = projectFiles(handle.project);
    const ins = libraryInsertion(library, files, arg.name, (n) => definesWorkbookName(handle.project, n));
    if (!ins || ins.names.length === 0) return { names: [], files: [] };
    const changed: Record<string, string> = {};
    for (const fe of ins.edits) changed[fe.path] = applyEdits(fe.create ? "" : (files[fe.path] ?? ""), fe.edits);
    const warning = tagWarning(files, changed, ins.names);
    const edit = new vscode.WorkspaceEdit();
    for (const fe of ins.edits) {
      const uri = this.ws.fileUri(handle, fe.path);
      if (fe.create) {
        edit.createFile(uri, { ignoreIfExists: true });
        edit.insert(uri, new vscode.Position(0, 0), applyEdits("", fe.edits));
      } else {
        const doc = await vscode.workspace.openTextDocument(uri);
        for (const x of fe.edits) edit.replace(uri, new vscode.Range(doc.positionAt(x.start), doc.positionAt(x.end)), x.text);
      }
    }
    if (!(await vscode.workspace.applyEdit(edit))) throw new Error(`could not add ${ins.names.join(", ")} to the source`);
    await this.keepBases(handle, ins.bases);
    const where = ins.edits.map((f) => f.path).join(", ");
    void act.info(`xln: added ${ins.names.join(", ")} from the library to ${where} (not saved; xln build writes it into the workbook).`, "Open").then((b) => {
      if (b) void vscode.window.showTextDocument(this.ws.fileUri(handle, ins.edits[0]!.path), { preview: true });
    });
    if (warning) void act.warn(`xln: ${warning}`);
    return { names: ins.names, files: ins.edits.map((f) => f.path), ...(warning ? { warning } : {}) };
  }

  /** `xln: Insert library function`: a quick pick of the library functions the project lacks. */
  async insertCommand(arg?: unknown): Promise<InsertResult | undefined> {
    const act = new Activity("lib insert");
    const handle = await this.handleOf(arg);
    if (!handle) return undefined;
    const { library } = await this.loaded(handle, true);
    let name = arg && typeof arg === "object" && typeof (arg as { name?: unknown }).name === "string" ? (arg as { name: string }).name : undefined;
    if (name === undefined) {
      const items = library.functions
        .filter((f) => !definesWorkbookName(handle.project, f.name))
        .map((f) => ({ label: f.name, description: f.params.join(", "), detail: f.summary }));
      if (items.length === 0) {
        void act.info("xln: the project has every function of the library already.");
        return undefined;
      }
      name = (await vscode.window.showQuickPick(items, { placeHolder: "Library function to add to the project", matchOnDetail: true }))?.label;
      if (name === undefined) return undefined;
    }
    // With the cursor in a formula of this project, the call goes there too.
    const ed = vscode.window.activeTextEditor;
    const at = ed ? this.ws.locate(ed.document.uri) : undefined;
    if (ed && at?.handle === handle && formulaSite(handle.project, at.path, ed.document.offsetAt(ed.selection.active))) {
      await ed.edit((b) => b.replace(ed.selection, name!));
    }
    return this.insertDefinition({ root: handle.uri.toString(), name });
  }

  // ---- update -------------------------------------------------------------------------------

  /**
   * "Take the library's version": the entry's definition and doc comment become the
   * library's and `@from(lib #…)` records that version, in the source (unsaved). On a copy
   * edited since its base (modified, both changed), or with no base (differs), it first asks, naming the edit it
   * discards; `confirm: true` (tests) answers yes, `false` no.
   */
  async take(arg?: { root: string; path: string; name: string; confirm?: boolean }): Promise<boolean> {
    const act = new Activity("lib take", arg?.name);
    const handle = arg ? this.ws.handleFor(arg.root) : undefined;
    if (!handle || !arg) return noTarget(arg, act), false;
    const { library } = await this.loaded(handle);
    const fn = library.get(arg.name);
    const uri = this.ws.fileUri(handle, arg.path);
    const doc = await vscode.workspace.openTextDocument(uri);
    const entry = findEntry(doc.getText(), arg.name);
    if (!fn || !entry) {
      void act.info(`xln: ${arg.name} is not in ${fn ? arg.path : "the library"}.`);
      return false;
    }
    const item = (await this.report(handle)).report.items.find((i) => i.name.toLowerCase() === fn.name.toLowerCase());
    const ask = item ? entryLibraryActions(item).take?.confirm : undefined;
    const replacement = libraryReplacement(doc.getText(), entry, fn);
    const warning = tagWarning(projectFiles(handle.project), { [arg.path]: applyEdits(doc.getText(), replacement) }, [fn.name]);
    if (ask) {
      const discard = item?.state === "differs" ? "Take the library's version" : "Discard and take";
      const ok = arg.confirm ?? (await vscode.window.showWarningMessage(`${ask}${warning ? ` Also: ${warning}.` : ""}`, { modal: true }, discard)) === discard;
      if (!ok) {
        void act.info(`xln: ${fn.name} kept as it is.`);
        return false;
      }
    }
    const edit = new vscode.WorkspaceEdit();
    for (const x of replacement) edit.replace(uri, new vscode.Range(doc.positionAt(x.start), doc.positionAt(x.end)), x.text);
    if (!(await vscode.workspace.applyEdit(edit))) return false;
    await this.keepBases(handle, [libraryFunctionBase(fn)]);
    act.line(`took the library's version into ${arg.path}, @from(lib #…) recorded (not saved; build to carry it into the workbook)`);
    if (warning && !ask) void act.warn(`xln: ${warning}`);
    return true;
  }

  // ---- diffs and publish ----------------------------------------------------------------------

  private virtualUri(kind: string, name: string, root: string): vscode.Uri {
    return vscode.Uri.from({ scheme: LIB_SCHEME, path: `/${kind}/${name}`, query: root });
  }

  private setVirtual(uri: vscode.Uri, text: string): void {
    this.virtual.set(uri.toString(), text);
    this.docsChanged.fire(uri);
  }

  /** The project's entry of `name` (workbook scope, a module or other names file). */
  private projectEntry(handle: ProjectHandle, name: string): { path: string; entry: ModuleEntry } | undefined {
    for (const [path, f] of handle.project.files) {
      const e = f.parsed.entries.find((x) => x.scope === undefined && !x.cell && x.name.toLowerCase() === name.toLowerCase());
      if (e) return { path, entry: e };
    }
    return undefined;
  }

  /** The project's definition against the library's, both laid out the same way. */
  async diff(arg?: { root: string; name: string }): Promise<string | undefined> {
    const act = new Activity("lib diff", arg?.name);
    const handle = arg ? this.ws.handleFor(arg.root) : undefined;
    if (!handle || !arg) return noTarget(arg, act), undefined;
    const { library } = await this.loaded(handle);
    const fn = library.get(arg.name);
    const pe = this.projectEntry(handle, arg.name);
    if (!fn) return void act.warn(`xln: ${arg.name} is not in the library.`), undefined;
    if (!pe) return void act.warn(`xln: ${arg.name} is not in the project.`), undefined;
    // Changed on both sides: the three versions, as what each side changed since the base.
    const item = (await this.report(handle)).report.items.find((i) => i.name.toLowerCase() === fn.name.toLowerCase());
    if (item?.state === "both-changed") {
      const uri = this.virtualUri("three-way", `${fn.name} (base, project, library)`, arg.root);
      this.setVirtual(uri, `// ${fn.name}: changed here and in the library since its base #${item.base} (library #${item.libraryHash}).\n${threeWayText(item)}\n`);
      await vscode.window.showTextDocument(uri, { preview: true });
      act.line(`three-way text opened: base #${item.base}, project, library #${item.libraryHash}`);
      return uri.toString();
    }
    const left = this.virtualUri("project", `${fn.name}.xln`, arg.root);
    const right = this.virtualUri("library", `${fn.name}.lambda`, arg.root);
    this.setVirtual(left, comparableLayout(pe.entry.formula).join("\n") + "\n");
    this.setVirtual(right, comparableLayout(fn.definition).join("\n") + "\n");
    await vscode.commands.executeCommand("vscode.diff", left, right, `${fn.name}: project ↔ library`, { preview: true });
    act.line(`diff opened: project ↔ library (${item ? libStateLabel(item.state) : "no state"})`);
    return right.toString();
  }

  /**
   * "Publish to library": shows the library file against what it would become, then asks
   * before writing it. `confirm: true` (tests) skips the question.
   */
  async publish(arg?: { root: string; name: string; confirm?: boolean }): Promise<{ file: string; written: boolean; changed: string[]; tagWarning?: string } | undefined> {
    const act = new Activity("lib publish", arg?.name);
    const handle = arg ? this.ws.handleFor(arg.root) : undefined;
    if (!handle || !arg) return noTarget(arg, act), undefined;
    const lib = await this.loaded(handle, true);
    const pe = this.projectEntry(handle, arg.name);
    if (!pe) return void act.warn(`xln: ${arg.name} is not a definition in this project.`), undefined;
    const fn = lib.library.get(pe.entry.name);
    const fileUri = vscode.Uri.joinPath(lib.uri, fn?.path ?? `${pe.entry.name}.lambda`);
    const before = fn ? new TextDecoder().decode(await vscode.workspace.fs.readFile(fileUri)) : "";
    const r = publishLambda({ name: pe.entry.name, doc: pe.entry.doc, formula: pe.entry.formula }, fn ? { path: fn.path, text: before } : undefined);
    if (r.error) {
      void act.error(`xln: ${r.error}`);
      return undefined;
    }
    // The build tags the project's doc comment (the workbook's), with the base publish records.
    const files = projectFiles(handle.project);
    const own = files[pe.path] ?? "";
    const based = findEntry(own, pe.entry.name);
    const baseEdit = based ? libraryBaseEdit(own, based, publishedBase(pe.entry.formula, pe.entry.name)) : undefined;
    const tagged = tagWarning(files, { [pe.path]: baseEdit ? applyEdits(own, [baseEdit]) : own }, [pe.entry.name]);
    const also = tagged ? { tagWarning: tagged } : {};
    if (r.changed.length === 0) {
      const recorded = await this.recordBase(handle, pe, arg.name, r.base);
      void act.info(`xln: the library's ${r.path} already has this definition${recorded ? `; recorded @from(lib #${recorded}) on ${pe.entry.name} (not saved)` : ""}.`);
      if (tagged) void act.warn(`xln: ${tagged}`);
      return { file: fileUri.toString(), written: false, changed: [], ...also };
    }
    const warning = publishWarning((await this.report(handle)).report.items.find((i) => i.name.toLowerCase() === pe.entry.name.toLowerCase()));
    const proposed = this.virtualUri("publish", r.path, arg.root);
    this.setVirtual(proposed, r.text);
    const left = fn ? fileUri : this.virtualUri("empty", r.path, arg.root);
    if (!fn) this.setVirtual(left, "");
    await vscode.commands.executeCommand("vscode.diff", left, proposed, `${r.path}: library ↔ after publishing`, { preview: true });
    let ok = arg.confirm === true;
    if (!ok) {
      const write = r.created ? "Create" : "Update";
      const answer = await vscode.window.showWarningMessage(
        `${warning ? warning + " " : ""}${r.created ? "Create" : "Update"} ${r.path} in the library (${vscode.workspace.asRelativePath(lib.uri)})? ${r.created ? "A new file with a generated header." : `Changes: ${r.changed.join(", ")}; the header's other fields and the rationale stay.`} ${pe.entry.name} then records the published version as its base, @from(lib #…).${tagged ? ` Note: ${tagged}.` : ""}`,
        { modal: true },
        write,
      );
      ok = answer === write;
    }
    if (!ok) return act.line(`${r.path}: not written (declined)`), { file: fileUri.toString(), written: false, changed: r.changed, ...also };
    await vscode.workspace.fs.writeFile(fileUri, new TextEncoder().encode(r.text));
    const recorded = await this.recordBase(handle, pe, arg.name, r.base);
    await this.load(handle, true);
    this.lensesChanged.fire();
    void act.info(`xln: ${r.created ? "created" : "updated"} ${r.path} in the library${recorded ? `; ${pe.entry.name} records it as its base, @from(lib #${recorded}) (not saved; build to carry it into the workbook)` : ""}.`);
    if (tagged && arg.confirm === true) void act.warn(`xln: ${tagged}`);
    return { file: fileUri.toString(), written: true, changed: r.changed, ...also };
  }

  /** After a publish: the project's entry records the published version as its base (an unsaved edit), whose text is kept. Returns the hash when it changed the entry. */
  private async recordBase(handle: ProjectHandle, pe: { path: string; entry: ModuleEntry }, name: string, base: LibraryBase | undefined): Promise<string | undefined> {
    if (base) await this.keepBases(handle, [base]);
    const hash = publishedBase(pe.entry.formula, pe.entry.name);
    const uri = this.ws.fileUri(handle, pe.path);
    const doc = await vscode.workspace.openTextDocument(uri);
    const entry = findEntry(doc.getText(), name);
    const x = entry ? libraryBaseEdit(doc.getText(), entry, hash) : undefined;
    if (!x) return undefined;
    const edit = new vscode.WorkspaceEdit();
    edit.replace(uri, new vscode.Range(doc.positionAt(x.start), doc.positionAt(x.end)), x.text);
    return (await vscode.workspace.applyEdit(edit)) ? hash : undefined;
  }
}

function commandUri(command: string, arg: unknown): vscode.Uri {
  return vscode.Uri.parse(`command:${command}?${encodeURIComponent(JSON.stringify([arg]))}`);
}

/** Every early return says why: a silent command reads as a broken one. */
function noTarget(arg: { root: string } | undefined, act: Activity): void {
  if (arg) void act.warn("xln: this file is not in a pulled xln project.");
}
