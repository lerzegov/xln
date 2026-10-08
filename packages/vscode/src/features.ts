// The Name Manager features (B1–B4) on top of the project model: outline, folding,
// workspace symbols, go to definition, references (names and cells), call hierarchy
// for "uses" / "used by", hover, diagnostics, and the read-only `xln-cells:` documents.
//
// Cell usages: a reference result needs a document and a range, but cells live in the
// workbook. Each name gets a read-only virtual document `xln-cells:/<workbook>/<key>`
// with one line per place (`'S1'!B3:B7: =Rate*A3  → 0.3`), and Find All References
// returns those lines next to the uses in `.xln` files. One gesture then shows every
// usage, in names and in cells, in VS Code's own references view, peek and search
// editor, in desktop and in the browser alike; no custom view to learn or keep in sync.

import * as vscode from "vscode";
import { CONFIG_FILE, strayFile } from "@xln/core";
import { configMarks } from "./model/config.js";
import { liveProblems } from "./model/editor.js";
import type { Loc, NameDef, OutlineSymbol, Problem, Project, SymbolKind } from "./model/project.js";
import { isNamesFile, Project as ProjectModel } from "./model/project.js";
import { hoverMarkdown, renderUsages, type UsageDocument, type UsageLine } from "./model/usages.js";
import { libraryBaseHover } from "./model/library.js";
import { renameLogLines } from "./model/log.js";
import { Activity } from "./log.js";
import type { ProjectHandle, XlnWorkspace } from "./xlnWorkspace.js";

export const CELLS_SCHEME = "xln-cells";
/** The tab title of a usages document reads `Rate (usages)`. */
const USAGES_SUFFIX = " (usages)";
const SELECTOR: vscode.DocumentSelector = [{ language: "xln" }];

const SYMBOL_KIND: Record<SymbolKind, vscode.SymbolKind> = {
  block: vscode.SymbolKind.Namespace,
  constant: vscode.SymbolKind.Constant,
  range: vscode.SymbolKind.Field,
  spill: vscode.SymbolKind.Array,
  table: vscode.SymbolKind.Struct,
  formula: vscode.SymbolKind.Variable,
  lambda: vscode.SymbolKind.Function,
  unparsed: vscode.SymbolKind.Null,
};

interface Located {
  handle: ProjectHandle | undefined;
  project: Project;
  path: string;
}

export class Features {
  private readonly diagnostics = vscode.languages.createDiagnosticCollection("xln");
  /** Quick fixes of published diagnostics (an edited cell address restored, a name qualified), the first preferred. */
  private readonly fixes = new WeakMap<vscode.Diagnostic, FileFix[]>();
  private pending: ReturnType<typeof setTimeout> | undefined;
  /** The audit's cell findings (audit.ts), so a cell statement's live problem is not said twice. */
  cellFinding: ((workbook: vscode.Uri, rule: string, sheet: string, ref: string) => boolean) | undefined;
  /** Rendered `xln-cells:` documents, for the reference locations into them. */
  private readonly usageDocs = new Map<string, UsageDocument>();
  private readonly usageChanged = new vscode.EventEmitter<vscode.Uri>();

  constructor(private readonly ws: XlnWorkspace) {}

  register(context: vscode.ExtensionContext): void {
    const ws = this.ws;
    context.subscriptions.push(
      this.diagnostics,
      this.usageChanged,
      ws.onDidChange((e) => this.scheduleDiagnostics(e.edit)),
      // Typing in xln.config.json: its marks follow (a save reloads the project anyway).
      vscode.workspace.onDidChangeTextDocument((e) => {
        if (e.document.uri.path.endsWith(`/${CONFIG_FILE}`)) this.scheduleDiagnostics(true);
      }),
      { dispose: () => this.pending && clearTimeout(this.pending) },
      vscode.languages.registerDocumentSymbolProvider(SELECTOR, { provideDocumentSymbols: (d) => this.symbols(d) }, { label: "xln" }),
      vscode.languages.registerFoldingRangeProvider(SELECTOR, { provideFoldingRanges: (d) => this.folding(d) }),
      vscode.languages.registerDefinitionProvider(SELECTOR, { provideDefinition: (d, p) => this.definition(d, p) }),
      vscode.languages.registerReferenceProvider(SELECTOR, { provideReferences: (d, p, c) => this.references(d, p, c.includeDeclaration) }),
      vscode.languages.registerHoverProvider(SELECTOR, { provideHover: (d, p) => this.hover(d, p) }),
      vscode.languages.registerRenameProvider(SELECTOR, {
        prepareRename: (d, p) => this.prepareRename(d, p),
        provideRenameEdits: (d, p, n) => this.renameEdits(d, p, n),
      }),
      vscode.languages.registerWorkspaceSymbolProvider({ provideWorkspaceSymbols: (q) => this.workspaceSymbols(q) }),
      vscode.languages.registerCallHierarchyProvider(SELECTOR, {
        prepareCallHierarchy: (d, p) => this.prepareCalls(d, p),
        provideCallHierarchyIncomingCalls: (item) => this.incoming(item),
        provideCallHierarchyOutgoingCalls: (item) => this.outgoing(item),
      }),
      vscode.languages.registerCodeActionsProvider(SELECTOR, { provideCodeActions: (d, _r, c) => this.codeActions(d, c) }, { providedCodeActionKinds: [vscode.CodeActionKind.QuickFix] }),
      vscode.workspace.registerTextDocumentContentProvider(CELLS_SCHEME, {
        onDidChange: this.usageChanged.event,
        provideTextDocumentContent: (uri) => this.usageText(uri),
      }),
      vscode.commands.registerCommand(LOG_FIX, (title: string, files: string[]) => new Activity("quick fix", title).line(`edited ${files.join(", ")} (unsaved)`)),
    );
  }

  // ---- locating -------------------------------------------------------------------

  /** The project a document belongs to; a loose `.xln` file gets a project of its own. */
  private locate(doc: vscode.TextDocument): Located {
    const at = this.ws.locate(doc.uri);
    if (at && at.handle.project.files.has(at.path)) return { handle: at.handle, project: at.handle.project, path: at.path };
    const project = new ProjectModel(doc.uri.toString(), undefined);
    project.setFile("names/this.xln", doc.getText());
    return { handle: undefined, project, path: "names/this.xln" };
  }

  private uriOf(l: Located, path: string): vscode.Uri {
    return l.handle ? this.ws.fileUri(l.handle, path) : vscode.Uri.parse(l.project.root);
  }

  private range(project: Project, path: string, start: number, end: number): vscode.Range {
    const lines = project.files.get(path)!.lines;
    const a = lines.position(start);
    const b = lines.position(end);
    return new vscode.Range(a.line, a.character, b.line, b.character);
  }

  private location(l: Located, loc: Loc): vscode.Location {
    return new vscode.Location(this.uriOf(l, loc.path), this.range(l.project, loc.path, loc.start, loc.end));
  }

  private offset(l: Located, pos: vscode.Position): number {
    return l.project.files.get(l.path)!.lines.offset(pos);
  }

  // ---- B1 ---------------------------------------------------------------------------

  private symbols(doc: vscode.TextDocument): vscode.DocumentSymbol[] {
    const l = this.locate(doc);
    const conv = (s: OutlineSymbol): vscode.DocumentSymbol => {
      const sym = new vscode.DocumentSymbol(
        s.name,
        s.detail,
        SYMBOL_KIND[s.kind],
        this.range(l.project, l.path, s.start, s.end),
        this.range(l.project, l.path, s.selectionStart, s.selectionEnd),
      );
      sym.children = s.children.map(conv);
      return sym;
    };
    return l.project.outline(l.path).map(conv);
  }

  private folding(doc: vscode.TextDocument): vscode.FoldingRange[] {
    const l = this.locate(doc);
    return l.project
      .folding(l.path)
      .map((f) => new vscode.FoldingRange(f.start, f.end, f.kind === "comment" ? vscode.FoldingRangeKind.Comment : undefined));
  }

  // ---- B2 ---------------------------------------------------------------------------

  private workspaceSymbols(query: string): vscode.SymbolInformation[] {
    const q = query.toLowerCase();
    // VS Code ranks and filters the result itself; hand it every name whose letters
    // appear in order, as its fuzzy matcher expects.
    const fuzzy = (s: string): boolean => {
      let k = 0;
      for (const c of s.toLowerCase()) if (c === q[k]) k++;
      return k === q.length;
    };
    const out: vscode.SymbolInformation[] = [];
    for (const handle of this.ws.projects) {
      const l: Located = { handle, project: handle.project, path: "" };
      for (const def of handle.project.defs) {
        if (!fuzzy(def.name)) continue;
        const kind = SYMBOL_KIND[handle.project.analysis(def).classification.kind];
        const container = `${def.scope === undefined ? "" : def.scope + " · "}${def.file.path.slice("names/".length)}`;
        out.push(new vscode.SymbolInformation(def.name, kind, container, this.location(l, handle.project.nameLoc(def))));
      }
    }
    return out;
  }

  // ---- B3 ---------------------------------------------------------------------------

  private definition(doc: vscode.TextDocument, pos: vscode.Position): vscode.Location | undefined {
    const l = this.locate(doc);
    const d = l.project.definition(l.path, this.offset(l, pos));
    return d && this.location(l, d);
  }

  private async references(doc: vscode.TextDocument, pos: vscode.Position, includeDeclaration: boolean): Promise<vscode.Location[]> {
    const l = this.locate(doc);
    const offset = this.offset(l, pos);
    const locals = l.project.localReferences(l.path, offset);
    if (locals) return (includeDeclaration ? locals : locals.slice(1)).map((x) => this.location(l, x));
    const key = l.project.keyAt(l.path, offset);
    const def = key === undefined ? undefined : l.project.lookup(key);
    if (!def) return [];
    return this.nameReferences(l, def, includeDeclaration);
  }

  // ---- Rename Symbol (F2, M5) ---------------------------------------------------------

  private prepareRename(doc: vscode.TextDocument, pos: vscode.Position): { range: vscode.Range; placeholder: string } {
    const l = this.locate(doc);
    const t = l.project.renameTarget(l.path, this.offset(l, pos));
    if (typeof t === "string") {
      // VS Code shows the reason at the cursor; the log keeps it.
      new Activity("rename", `at ${l.path}:${pos.line + 1}`).line(`refused: ${t}`);
      throw new Error(t);
    }
    return { range: this.range(l.project, t.at.path, t.at.start, t.at.end), placeholder: t.def.name };
  }

  /**
   * The source edits of a rename (core `renameInSource`, as `xln rename`). Applied at once,
   * as any rename (the author found the unticked preview boxes a trap, 2026-10-07); Shift+Enter
   * still previews. Nothing is written to the workbook: the next build renames the name
   * there and rewrites it in the cells.
   */
  private renameEdits(doc: vscode.TextDocument, pos: vscode.Position, newName: string): vscode.WorkspaceEdit {
    const l = this.locate(doc);
    const t = l.project.renameTarget(l.path, this.offset(l, pos));
    const act = new Activity("rename", typeof t === "string" ? `at ${l.path}:${pos.line + 1}` : `${t.def.key} → ${newName}`);
    if (typeof t === "string") {
      act.line(`refused: ${t}`);
      throw new Error(t);
    }
    const r = l.project.renameEdits(t.def.key, newName);
    if (typeof r === "string") {
      act.line(`refused: ${r}`);
      throw new Error(r);
    }
    act.lines(renameLogLines(r, l.handle ? vscode.workspace.asRelativePath(l.handle.uri) : undefined));
    const edit = new vscode.WorkspaceEdit();
    const what: Record<string, string> = {
      name: `the name ${r.from} → ${newName}`,
      reference: `formulas reading ${r.from}`,
      annotation: r.annotation === "removed" ? "@renamed removed (renamed back)" : `@renamed: the build renames ${r.from} in the workbook and rewrites it in the cells`,
    };
    for (const e of r.edits) {
      const at: Located = { ...l, path: e.path };
      edit.replace(this.uriOf(at, e.path), this.range(l.project, e.path, e.start, e.end), e.text, { needsConfirmation: false, label: what[e.label]! });
    }
    return edit;
  }

  /** Every use of a name: in `.xln` files, then in the workbook (the `xln-cells:` lines). */
  async nameReferences(l: { handle: ProjectHandle | undefined; project: Project }, def: NameDef, includeDeclaration: boolean): Promise<vscode.Location[]> {
    const at: Located = { ...l, path: def.file.path };
    const out = l.project.references(def.key).map((x) => this.location(at, x));
    if (includeDeclaration) out.unshift(this.location(at, l.project.nameLoc(def)));
    if (l.handle) out.push(...(await this.cellLocations(l.handle, def)));
    return out;
  }

  /** Where a name is defined, as a location in its `.xln` file. */
  nameLocation(handle: ProjectHandle, def: NameDef): vscode.Location {
    return this.location({ handle, project: handle.project, path: def.file.path }, handle.project.nameLoc(def));
  }

  /** The hover of a defined name (B4), with the link to its cell usages. */
  async nameHover(handle: ProjectHandle | undefined, project: Project, def: NameDef): Promise<vscode.MarkdownString> {
    const wb = handle ? await this.ws.workbook(handle) : undefined;
    const md = new vscode.MarkdownString(hoverMarkdown(project, def, wb));
    if (handle) {
      const args = encodeURIComponent(JSON.stringify([{ root: handle.uri.toString(), key: def.key }]));
      md.appendMarkdown(`\n\n[Show cell usages](command:xln.showUsages?${args})`);
      md.isTrusted = { enabledCommands: ["xln.showUsages"] };
    }
    return md;
  }

  /** The usage line of an `xln-cells:` document at a 0-based line, with its project. */
  async usageItemAt(uri: vscode.Uri, line: number): Promise<{ handle: ProjectHandle; item: UsageLine } | undefined> {
    const handle = this.ws.handleFor(uri.query);
    if (!handle) return undefined;
    const doc = this.usageDocs.get(uri.toString()) ?? (await this.usageDocument(uri));
    const item = doc?.items.find((i) => i.line === line);
    return item && { handle, item };
  }

  usageUri(handle: ProjectHandle, key: string): vscode.Uri {
    const workbook = handle.project.workbookName ?? "workbook";
    return vscode.Uri.from({ scheme: CELLS_SCHEME, path: `/${workbook}/${key}${USAGES_SUFFIX}`, query: handle.uri.toString() });
  }

  private async usageDocument(uri: vscode.Uri): Promise<UsageDocument | undefined> {
    const handle = this.ws.handleFor(uri.query);
    if (!handle) return undefined;
    const last = uri.path.slice(uri.path.lastIndexOf("/") + 1);
    const key = last.endsWith(USAGES_SUFFIX) ? last.slice(0, -USAGES_SUFFIX.length) : last;
    const project = handle.project;
    const doc = renderUsages(project.lookup(key)?.key ?? key, project.manifestName(key), project.manifest, await this.ws.workbook(handle));
    this.usageDocs.set(uri.toString(), doc);
    return doc;
  }

  private async usageText(uri: vscode.Uri): Promise<string> {
    return (await this.usageDocument(uri))?.text ?? `// ${uri.path}: no xln project at ${uri.query} (run xln: Reload?)\n`;
  }

  /** The `xln-cells:` lines of a name, as reference locations. */
  private async cellLocations(handle: ProjectHandle, def: NameDef): Promise<vscode.Location[]> {
    const uri = this.usageUri(handle, def.key);
    const doc = await this.usageDocument(uri);
    if (!doc) return [];
    this.usageChanged.fire(uri);
    return doc.items.map((i) => new vscode.Location(uri, new vscode.Range(i.line, i.start, i.line, i.end)));
  }

  /** `xln: Show usages`: opens the cell usages of the name at the cursor (or of `key`). */
  async showUsages(arg?: { root: string; key: string }): Promise<vscode.Uri | undefined> {
    let handle: ProjectHandle | undefined;
    let key: string | undefined;
    if (arg) {
      handle = this.ws.handleFor(arg.root);
      key = arg.key;
    } else {
      const ed = vscode.window.activeTextEditor;
      if (ed) {
        const l = this.locate(ed.document);
        handle = l.handle;
        key = l.project.keyAt(l.path, this.offset(l, ed.selection.active));
      }
    }
    if (!handle || key === undefined) {
      void new Activity("usages").info("xln: put the cursor on a name in a pulled project's .xln file.");
      return undefined;
    }
    const uri = this.usageUri(handle, handle.project.lookup(key)?.key ?? key);
    const doc = await this.usageDocument(uri);
    new Activity("usages", handle.project.lookup(key)?.key ?? key).line(doc ? `${doc.items.length} place(s) in names and cells` : "no usages document: the project is not loaded");
    this.usageChanged.fire(uri);
    await vscode.window.showTextDocument(uri, { preview: true });
    return uri;
  }

  // "Uses" and "used by" among names, as VS Code's call hierarchy.
  private item(project: Project, def: NameDef, l: Located): vscode.CallHierarchyItem {
    const loc = project.nameLoc(def);
    const range = this.range(project, def.file.path, def.entry.start, def.entry.end);
    const sel = this.range(project, loc.path, loc.start, loc.end);
    const kind = SYMBOL_KIND[project.analysis(def).classification.kind];
    // The detail carries the scope, so the item alone tells the key (see fromItem).
    return new vscode.CallHierarchyItem(kind, def.name, def.scope ?? "", this.uriOf(l, def.file.path), range, sel);
  }

  private prepareCalls(doc: vscode.TextDocument, pos: vscode.Position): vscode.CallHierarchyItem | undefined {
    const l = this.locate(doc);
    const key = l.project.keyAt(l.path, this.offset(l, pos));
    const def = key === undefined ? undefined : l.project.lookup(key);
    return def && this.item(l.project, def, l);
  }

  private fromItem(item: vscode.CallHierarchyItem): { l: Located; def: NameDef } | undefined {
    const handle = this.ws.locate(item.uri)?.handle;
    if (!handle) return undefined;
    const def = handle.project.lookup(item.detail ? `${item.detail}!${item.name}` : item.name);
    return def && { l: { handle, project: handle.project, path: def.file.path }, def };
  }

  private incoming(item: vscode.CallHierarchyItem): vscode.CallHierarchyIncomingCall[] {
    const x = this.fromItem(item);
    if (!x) return [];
    return x.l.project
      .usedBy(x.def.key)
      .map((u) => new vscode.CallHierarchyIncomingCall(this.item(x.l.project, u.def, x.l), u.locs.map((r) => this.range(x.l.project, r.path, r.start, r.end))));
  }

  private outgoing(item: vscode.CallHierarchyItem): vscode.CallHierarchyOutgoingCall[] {
    const x = this.fromItem(item);
    if (!x) return [];
    return x.l.project
      .uses(x.def)
      .map((u) => new vscode.CallHierarchyOutgoingCall(this.item(x.l.project, u.def, x.l), u.locs.map((r) => this.range(x.l.project, r.path, r.start, r.end))));
  }

  // ---- B4 ---------------------------------------------------------------------------

  private async hover(doc: vscode.TextDocument, pos: vscode.Position): Promise<vscode.Hover | undefined> {
    const l = this.locate(doc);
    const offset = this.offset(l, pos);
    const base = libraryBaseHover(l.project, l.path, offset);
    if (base !== undefined) return new vscode.Hover(new vscode.MarkdownString(base));
    const key = l.project.keyAt(l.path, offset);
    const def = key === undefined ? undefined : l.project.lookup(key);
    if (!def) return undefined;
    return new vscode.Hover(await this.nameHover(l.handle, l.project, def));
  }

  // ---- diagnostics ------------------------------------------------------------------

  private codeActions(doc: vscode.TextDocument, context: vscode.CodeActionContext): vscode.CodeAction[] {
    const out: vscode.CodeAction[] = [];
    let live: { fixes: FileFix[]; code: string | undefined; range: vscode.Range }[] | undefined;
    for (const d of context.diagnostics) {
      let fixes = this.fixes.get(d);
      // An audit finding the checks as you type also make (and left to the audit): their fixes.
      if (!fixes && d.source === "xln check" && typeof d.code === "string" && LIVE_CODES.has(d.code)) {
        live ??= this.liveFixes(doc);
        fixes = live.find((x) => x.code === d.code && x.range.intersection(d.range) !== undefined)?.fixes;
      }
      (fixes ?? []).forEach((fix, k) => {
        const a = new vscode.CodeAction(fix.title, vscode.CodeActionKind.QuickFix);
        a.edit = new vscode.WorkspaceEdit();
        a.edit.replace(fix.uri, fix.range, fix.text);
        for (const m of fix.more ?? []) {
          if (m.range) a.edit.replace(m.uri, m.range, m.text);
          else a.edit.createFile(m.uri, { ignoreIfExists: true, contents: new TextEncoder().encode(m.text) });
        }
        // A fix that edits other files too is logged (VS Code runs the command after the edit).
        if (fix.more?.length) {
          const files = [fix.uri, ...fix.more.map((m) => m.uri)].map((u) => `${vscode.workspace.asRelativePath(u)}${fix.more!.some((m) => m.uri === u && !m.range) ? " (created)" : ""}`);
          a.command = { title: fix.title, command: LOG_FIX, arguments: [fix.title, files] };
        }
        a.diagnostics = [d];
        a.isPreferred = k === 0;
        out.push(a);
      });
    }
    return out;
  }

  /** The live problems of a document with their fixes, for the audit's diagnostics on it. */
  private liveFixes(doc: vscode.TextDocument): { fixes: FileFix[]; code: string | undefined; range: vscode.Range }[] {
    const at = this.ws.locate(doc.uri);
    if (!at || !isNamesFile(at.path) || !at.handle.project.files.has(at.path)) return [];
    const project = at.handle.project;
    return liveProblems(project, at.path).map((p) => ({
      code: p.code,
      range: this.range(project, at.path, p.start, p.end),
      fixes: (p.fixes ?? []).map((f) => ({ uri: doc.uri, range: this.range(project, at.path, f.start, f.end), title: f.title, text: f.text })),
    }));
  }

  /**
   * Republishes after a change: at once after a load, and after a pause in typing (the
   * checks as you type read every file of the project).
   */
  scheduleDiagnostics(edit = true): void {
    if (this.pending) clearTimeout(this.pending);
    this.pending = undefined;
    if (!edit) return this.publishDiagnostics();
    this.pending = setTimeout(() => {
      this.pending = undefined;
      this.publishDiagnostics();
    }, DEBOUNCE_MS);
  }

  publishDiagnostics(): void {
    this.diagnostics.clear();
    for (const handle of this.ws.projects) {
      for (const path of handle.project.files.keys()) {
        if (!isNamesFile(path)) continue;
        const uri = this.ws.fileUri(handle, path);
        // The audit's findings on this file: a live problem saying the same is left out.
        const audit = vscode.languages.getDiagnostics(uri).filter((d) => d.source === "xln check");
        const dup = (p: Problem, range: vscode.Range): boolean =>
          p.code !== undefined && LIVE_CODES.has(p.code) && audit.some((d) => d.code === p.code && d.range.intersection(range) !== undefined);
        // ... and on the cell's line of the formula view, for a cell statement.
        const dupCell = (p: Problem): boolean => {
          if (p.code === undefined || !LIVE_CODES.has(p.code) || !handle.workbookUri || !this.cellFinding) return false;
          const def = handle.project.defAt(path, p.start);
          const c = def?.entry.cell;
          const sheet = c && def.entry.cellSheet;
          return sheet !== undefined && this.cellFinding(handle.workbookUri, p.code, sheet, c!.range.split(":")[0]!);
        };
        const problems = handle.project.problems(path);
        const list: vscode.Diagnostic[] = [];
        for (const p of problems) {
          const range = this.range(handle.project, path, p.start, p.end);
          if (dup(p, range) || dupCell(p)) continue;
          const d = new vscode.Diagnostic(range, p.message, SEVERITY[p.severity]);
          d.source = "xln";
          if (p.code !== undefined) d.code = p.code;
          if (p.unnecessary) d.tags = [vscode.DiagnosticTag.Unnecessary];
          const fixes = [...(p.fix ? [p.fix] : []), ...(p.fixes ?? [])];
          if (fixes.length) {
            this.fixes.set(
              d,
              fixes.map((f) => ({
                uri,
                range: this.range(handle.project, path, f.start, f.end),
                title: f.title,
                text: f.text,
                ...(f.elsewhere ? { more: f.elsewhere.map((m) => ({ uri: this.ws.fileUri(handle, m.path), text: m.text, ...(m.create || !handle.project.files.has(m.path) ? {} : { range: this.range(handle.project, m.path, m.start, m.end) }) })) } : {}),
              })),
            );
          }
          list.push(d);
        }
        if (list.length) this.diagnostics.set(uri, list);
      }
      // xln.config.json: what in it cannot be used, on its key (the open editor's text first).
      const configUri = vscode.Uri.joinPath(handle.uri, CONFIG_FILE);
      const configText = vscode.workspace.textDocuments.find((d) => d.uri.toString() === configUri.toString())?.getText() ?? handle.configText;
      if (configText !== undefined) {
        const marks = configMarks(configText).map((m) => {
          const d = new vscode.Diagnostic(new vscode.Range(m.start.line, m.start.character, m.end.line, m.end.character), m.message, m.severity === "warning" ? vscode.DiagnosticSeverity.Warning : vscode.DiagnosticSeverity.Information);
          d.source = "xln";
          d.code = "config";
          return d;
        });
        if (marks.length) this.diagnostics.set(configUri, marks);
      }
      // Files below names/ that are not .xln: the build refuses them (M3e).
      const sheets = handle.project.checkContext().sheets;
      for (const path of handle.strays) {
        const message = strayFile(path, sheets);
        if (message === undefined) continue;
        const d = new vscode.Diagnostic(new vscode.Range(0, 0, 0, 0), message, vscode.DiagnosticSeverity.Error);
        d.source = "xln";
        d.code = "stray-file";
        this.diagnostics.set(this.ws.fileUri(handle, path), [d]);
      }
    }
  }
}

/** A quick fix as an edit of a document. */
interface FileFix {
  uri: vscode.Uri;
  range: vscode.Range;
  title: string;
  text: string;
  /** Edits of other files made with it (a statement moved to another file); no range: a file to create. */
  more?: { uri: vscode.Uri; range?: vscode.Range; text: string }[];
}

const DEBOUNCE_MS = 250;

/** Internal (not in the palette): logs a quick fix that edited more than its file. */
const LOG_FIX = "xln.logQuickFix";

const SEVERITY: Record<Problem["severity"], vscode.DiagnosticSeverity> = {
  error: vscode.DiagnosticSeverity.Error,
  warning: vscode.DiagnosticSeverity.Warning,
  info: vscode.DiagnosticSeverity.Information,
  hint: vscode.DiagnosticSeverity.Hint,
};

/** Codes the checks as you type share with the audit (`xln check`). */
const LIVE_CODES = new Set(["C4.unknown-name", "C5.other-sheet", "C6.lambda-arity", "C6.builtin-arity", "C6.not-a-function"]);
