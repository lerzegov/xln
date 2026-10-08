// B6: the formula view of a sheet, a read-only virtual document
// `xln-formulas:/<workbook>/<Sheet> (formulas)?<workbook URI>` in order of appearance
// (a), `…/<Sheet> (calculation order)?…` in calculation order (b), and
// `…/[workbook] (calculation order)?…` for every sheet at once. Same design as the
// `xln-cells:` usages: plain text that VS Code's own features work on (outline, search,
// go to definition, hover, references), identical in desktop VS Code and vscode.dev.
//
// - Names in a formula, and the names on the left of a line (those defined as that cell,
//   its spill or its extent), go to their entry in the pulled project (F12, Ctrl/Cmd+click),
//   hover like names in `.xln` files, and Find References lists their uses.
// - The outline lists the entries as `Name — C6#` (the address alone when unnamed), so
//   the outline's filter finds a block by its name.
// - A cell reference in a formula goes to the line of that cell (or of the spill holding
//   it), on this sheet or another.
// - On a cell usage line (`xln-cells:`), F12 on the address opens this view at that cell.
// - Hover on an address: what the formula reads and what reads it (from the dependency
//   graph), linked to their lines. In calculation order also its level and cycle.
// - The editor title's button switches a sheet between the two orders.
// The workbook is read lazily and kept while its size and date stay the same (W5).

import * as vscode from "vscode";
import { sheetFromFileName, SHEETS_DIR, type DependencyGraph, type FormulaOrder, type WorkbookSnapshot } from "@xln/core";
import { pickWorkbook } from "./commands.js";
import { Activity } from "./log.js";
import { CELLS_SCHEME, type Features } from "./features.js";
import { baseName } from "./inspect.js";
import { addressAt, buildFormulaView, entryForCell, entryHover, entryLabel, entryLinks, graphFor, lhsAt, nameAt, refAt, WORKBOOK_VIEW, type FormulaViewDoc, type HoverNode } from "./model/formulaView.js";
import { LineIndex } from "./model/lines.js";
import { isWorkbookName } from "./model/pull.js";
import type { ProjectHandle, XlnWorkspace } from "./xlnWorkspace.js";

export const FORMULAS_SCHEME = "xln-formulas";
export const FORMULAS_LANGUAGE = "xln-formulas";
/** The tab title of a formula view reads `BS (formulas)`, or `BS (calculation order)`. */
const SUFFIX = " (formulas)";
const CALC_SUFFIX = " (calculation order)";
/** How many precedents and dependents a hover lists before "and n more". */
const HOVER_ITEMS = 12;

/** Arguments of `xln.formulaView` when called from code (tests, links). */
export interface FormulaViewArgs {
  /** The workbook's URI, as a string. */
  workbook: string;
  sheet?: string;
  /** Open at this cell (or at the spill or array holding it). */
  cell?: string;
  /** Default: order of appearance. */
  order?: FormulaOrder;
  /** Every sheet in calculation order (`sheet` then only says where to open). */
  all?: boolean;
}

interface Built {
  doc: FormulaViewDoc;
  wb: WorkbookSnapshot;
  generation: number;
  lines: LineIndex;
}

const SYMBOL: Record<string, vscode.SymbolKind> = {
  normal: vscode.SymbolKind.Variable,
  shared: vscode.SymbolKind.Variable,
  array: vscode.SymbolKind.Array,
  "dynamic-array": vscode.SymbolKind.Array,
  "data-table": vscode.SymbolKind.Struct,
};

export class FormulaViews {
  private readonly built = new Map<string, Built>();
  private readonly parses = new WeakMap<WorkbookSnapshot, Map<string, unknown>>();
  private readonly graphs = new WeakMap<WorkbookSnapshot, { generation: number; graph: DependencyGraph }>();
  private readonly changed = new vscode.EventEmitter<vscode.Uri>();
  /** Bumped when the projects reload: names may resolve differently. */
  private generation = 0;

  constructor(
    private readonly ws: XlnWorkspace,
    private readonly features: Features,
  ) {}

  register(context: vscode.ExtensionContext): void {
    const sel: vscode.DocumentSelector = { scheme: FORMULAS_SCHEME };
    context.subscriptions.push(
      this.changed,
      vscode.workspace.registerTextDocumentContentProvider(FORMULAS_SCHEME, {
        onDidChange: this.changed.event,
        provideTextDocumentContent: (uri) => this.text(uri),
      }),
      vscode.languages.registerDefinitionProvider(sel, { provideDefinition: (d, p) => this.definition(d, p) }),
      vscode.languages.registerHoverProvider(sel, { provideHover: (d, p) => this.hover(d, p) }),
      vscode.languages.registerReferenceProvider(sel, { provideReferences: (d, p, c) => this.references(d, p, c.includeDeclaration) }),
      vscode.languages.registerDocumentSymbolProvider(sel, { provideDocumentSymbols: (d) => this.symbols(d) }, { label: "xln formulas" }),
      vscode.languages.registerDefinitionProvider({ scheme: CELLS_SCHEME }, { provideDefinition: (d, p) => this.usageDefinition(d, p) }),
      this.ws.onDidChange(() => this.refresh()),
    );
  }

  /** The view of `sheet` (`WORKBOOK_VIEW`: every sheet, calculation order). */
  uriFor(workbook: vscode.Uri, sheet: string, order: FormulaOrder = "appearance"): vscode.Uri {
    const suffix = order === "calculation" || sheet === WORKBOOK_VIEW ? CALC_SUFFIX : SUFFIX;
    return vscode.Uri.from({ scheme: FORMULAS_SCHEME, path: `/${baseName(workbook.path)}/${sheet}${suffix}`, query: workbook.toString() });
  }

  private target(uri: vscode.Uri): { workbook: vscode.Uri; sheet: string; order: FormulaOrder } {
    const last = uri.path.slice(uri.path.lastIndexOf("/") + 1);
    const workbook = vscode.Uri.parse(uri.query);
    if (last.endsWith(CALC_SUFFIX)) return { workbook, sheet: last.slice(0, -CALC_SUFFIX.length), order: "calculation" };
    return { workbook, sheet: last.endsWith(SUFFIX) ? last.slice(0, -SUFFIX.length) : last, order: "appearance" };
  }

  /** The dependency graph of a snapshot, with the project's names; kept until either changes. */
  private graph(wb: WorkbookSnapshot, workbook: vscode.Uri): DependencyGraph {
    const hit = this.graphs.get(wb);
    if (hit && hit.generation === this.generation) return hit.graph;
    const graph = graphFor(wb, this.ws.handleForWorkbook(workbook)?.project, this.parsesOf(wb));
    this.graphs.set(wb, { generation: this.generation, graph });
    return graph;
  }

  private parsesOf(wb: WorkbookSnapshot): Map<string, unknown> {
    let cache = this.parses.get(wb);
    if (!cache) this.parses.set(wb, (cache = new Map()));
    return cache;
  }

  private refresh(): void {
    this.generation++;
    for (const d of vscode.workspace.textDocuments) if (d.uri.scheme === FORMULAS_SCHEME) this.changed.fire(d.uri);
  }

  /** The view behind a URI, rebuilt when the workbook or the projects changed. */
  private async build(uri: vscode.Uri): Promise<Built | undefined> {
    const { workbook, sheet, order } = this.target(uri);
    const wb = await this.ws.workbookAt(workbook);
    if (!wb) return undefined;
    const key = uri.toString();
    const hit = this.built.get(key);
    if (hit && hit.wb === wb && hit.generation === this.generation) return hit;
    const graph = order === "calculation" ? this.graph(wb, workbook) : undefined;
    const doc = buildFormulaView(wb, sheet, baseName(workbook.path), this.ws.handleForWorkbook(workbook)?.project, this.parsesOf(wb), order, graph);
    if (!doc) return undefined;
    const b: Built = { doc, wb, generation: this.generation, lines: new LineIndex(doc.text) };
    this.built.set(key, b);
    return b;
  }

  private async text(uri: vscode.Uri): Promise<string> {
    const b = await this.build(uri);
    if (b) return b.doc.text;
    const { workbook, sheet } = this.target(uri);
    if (sheet === WORKBOOK_VIEW) return `// No formula view: ${workbook.toString()} cannot be read.\n`;
    return `// No formula view: ${workbook.toString()} cannot be read, or it has no sheet '${sheet}'.\n`;
  }

  private range(b: Built, start: number, end: number): vscode.Range {
    const a = b.lines.position(start);
    const z = b.lines.position(end);
    return new vscode.Range(a.line, a.character, z.line, z.character);
  }

  /** The location of `cell` in the view of `sheet` (its own line, or the spill holding it), in `order`. */
  async cellLocation(workbook: vscode.Uri, sheet: string, cell: string, order: FormulaOrder = "appearance"): Promise<vscode.Location | undefined> {
    const uri = this.uriFor(workbook, sheet, order);
    const b = await this.build(uri);
    const e = b && entryForCell(b.doc, cell);
    return e && new vscode.Location(uri, this.range(b, e.address.start, e.address.end));
  }

  /** A cell's line as seen from the view `from`: the same document for the workbook view, else that sheet's view in the same order. */
  private async lineFrom(from: vscode.Uri, sheet: string, cell: string): Promise<vscode.Location | undefined> {
    const t = this.target(from);
    if (t.sheet !== WORKBOOK_VIEW) return this.cellLocation(t.workbook, sheet, cell, t.order);
    const b = await this.build(from);
    const e = b && entryForCell(b.doc, cell, sheet);
    return e && new vscode.Location(from, this.range(b, e.address.start, e.address.end));
  }

  private handle(uri: vscode.Uri): ProjectHandle | undefined {
    return this.ws.handleForWorkbook(this.target(uri).workbook);
  }

  /** The defined name under the cursor (in a formula or on the left), resolved in the pulled project. */
  private async nameDef(document: vscode.TextDocument, pos: vscode.Position) {
    const b = await this.build(document.uri);
    if (!b) return undefined;
    const offset = document.offsetAt(pos);
    let use: { id: string; key: string | undefined } | undefined;
    const hit = nameAt(b.doc, offset);
    if (hit) use = hit.line.names[hit.k]!;
    else {
      const left = lhsAt(b.doc, offset);
      const n = left && left.line.lhs[left.k]!;
      if (n) use = { id: n.display, key: n.key };
    }
    if (!use) return undefined;
    const handle = this.handle(document.uri);
    const def = use.key !== undefined ? handle?.project.lookup(use.key) : undefined;
    return { b, use, handle, def };
  }

  private async definition(document: vscode.TextDocument, pos: vscode.Position): Promise<vscode.Location | undefined> {
    const n = await this.nameDef(document, pos);
    if (n) return n.handle && n.def ? this.features.nameLocation(n.handle, n.def) : undefined;
    const b = await this.build(document.uri);
    const r = b && refAt(b.doc, document.offsetAt(pos));
    if (!r) return undefined;
    const ref = r.line.refs[r.k]!;
    if (ref.refKind !== "cell" && ref.refKind !== "area") return undefined;
    return this.lineFrom(document.uri, ref.sheet, ref.address.split(":")[0]!.split(".")[0]!);
  }

  private async hover(document: vscode.TextDocument, pos: vscode.Position): Promise<vscode.Hover | undefined> {
    const n = await this.nameDef(document, pos);
    if (n) {
      if (n.def) return new vscode.Hover(await this.features.nameHover(n.handle, n.handle!.project, n.def));
      const why = n.b.doc.linked ? "no such name in the project" : "the workbook is not pulled, so names are not linked";
      return new vscode.Hover(new vscode.MarkdownString(`\`${n.use.id}\`: ${n.use.key === undefined ? "no such defined name (Excel shows #NAME?)" : why}`));
    }
    const b = await this.build(document.uri);
    const e = b && addressAt(b.doc, document.offsetAt(pos));
    if (!b || !e) return undefined;
    const md = new vscode.MarkdownString(entryHover(b.doc, e));
    const graph = b.doc.graph ?? this.graph(b.wb, this.target(document.uri).workbook);
    const links = entryLinks(b.doc, e, graph);
    if (links) {
      md.isTrusted = { enabledCommands: ["xln.formulaView"] };
      md.appendMarkdown("\n\n" + this.linkList("reads", links.precedents, document.uri));
      md.appendMarkdown("\n\n" + this.linkList("read by", links.dependents, document.uri));
    }
    return new vscode.Hover(md);
  }

  /** `reads 3: Rate, [C6#](…), …`: names as code, cells linked to their line. */
  private linkList(what: string, nodes: HoverNode[], from: vscode.Uri): string {
    if (nodes.length === 0) return `${what}: nothing`;
    const t = this.target(from);
    const shown = nodes.slice(0, HOVER_ITEMS).map((n) => {
      const text = "`" + n.label.split("`").join("'") + "`";
      const via = n.via !== undefined ? ` (via \`${n.via.split("`").join("'")}\`)` : "";
      if (n.sheet === undefined || n.cell === undefined) return text + via;
      const args: FormulaViewArgs = { workbook: t.workbook.toString(), sheet: n.sheet, cell: n.cell, order: t.order, all: t.sheet === WORKBOOK_VIEW };
      return `[${text}](command:xln.formulaView?${encodeURIComponent(JSON.stringify([args]))})${via}`;
    });
    const more = nodes.length > HOVER_ITEMS ? `, … and ${nodes.length - HOVER_ITEMS} more` : "";
    return `${what} ${nodes.length}: ${shown.join(", ")}${more}`;
  }

  private async references(document: vscode.TextDocument, pos: vscode.Position, includeDeclaration: boolean): Promise<vscode.Location[]> {
    const n = await this.nameDef(document, pos);
    if (!n?.def || !n.handle) return [];
    return this.features.nameReferences({ handle: n.handle, project: n.handle.project }, n.def, includeDeclaration);
  }

  private async symbols(document: vscode.TextDocument): Promise<vscode.DocumentSymbol[]> {
    const b = await this.build(document.uri);
    if (!b) return [];
    const all = b.doc.sheet === WORKBOOK_VIEW;
    return b.doc.entries.map((e) => {
      const l = b.doc.lines[e.index]!;
      const label = entryLabel(l, all);
      const flat = l.formula.replace(/\s+/g, " ");
      const detail = "=" + (flat.length > 60 ? flat.slice(0, 59) + "…" : flat);
      const full = new vscode.Range(e.line, 0, e.lastLine, b.lines.lineText(e.lastLine).length);
      return new vscode.DocumentSymbol(label, detail, SYMBOL[l.kind] ?? vscode.SymbolKind.Variable, full, this.range(b, e.address.start, e.address.end));
    });
  }

  /** F12 on the address of a cell usage line: the formula view at that cell. */
  private async usageDefinition(document: vscode.TextDocument, pos: vscode.Position): Promise<vscode.Location | undefined> {
    const u = await this.features.usageItemAt(document.uri, pos.line);
    if (!u || u.item.kind !== "cell" || !u.item.sheet || !u.item.range || !u.handle.workbookUri) return undefined;
    if (pos.character > u.item.end) return undefined;
    return this.cellLocation(u.handle.workbookUri, u.item.sheet, u.item.range.split(":")[0]!);
  }

  // ---- the command ------------------------------------------------------------------

  /**
   * `xln: Formula view`. From a workbook in the Explorer it asks for the sheet; from a
   * `names/sheets/<Sheet>.xln` file it opens that sheet; from a cell usage line it opens
   * at that cell; otherwise it asks for the workbook and the sheet. Returns the view's URI.
   * `order` and `all` (every sheet, calculation order) come from the command or `arg`.
   */
  async open(arg?: unknown, order: FormulaOrder = "appearance", all = false): Promise<string | undefined> {
    await this.ws.ready();
    let workbook: vscode.Uri | undefined;
    let sheet: string | undefined;
    let cell: string | undefined;
    let from: vscode.Uri | undefined;
    if (arg instanceof vscode.Uri) from = arg;
    else if (arg && typeof arg === "object" && typeof (arg as FormulaViewArgs).workbook === "string") {
      const a = arg as FormulaViewArgs;
      workbook = vscode.Uri.parse(a.workbook);
      sheet = a.sheet;
      cell = a.cell;
      if (a.order) order = a.order;
      if (a.all) all = true;
    } else {
      const ed = vscode.window.activeTextEditor;
      if (ed?.document.uri.scheme === CELLS_SCHEME) {
        const u = await this.features.usageItemAt(ed.document.uri, ed.selection.active.line);
        if (u?.handle.workbookUri) {
          workbook = u.handle.workbookUri;
          sheet = u.item.sheet;
          cell = u.item.kind === "cell" ? u.item.range?.split(":")[0] : undefined;
        }
      } else if (ed?.document.uri.scheme === FORMULAS_SCHEME) {
        const t = this.target(ed.document.uri);
        workbook = t.workbook;
        if (all && t.sheet !== WORKBOOK_VIEW) sheet = t.sheet;
      } else if (ed) from = ed.document.uri;
    }
    if (from && !workbook) {
      if (isWorkbookName(baseName(from.path))) workbook = from;
      else {
        const at = this.ws.locate(from);
        workbook = at?.handle.workbookUri;
        if (at && at.path.startsWith(`names/${SHEETS_DIR}/`)) {
          const file = at.handle.project.files.get(at.path);
          sheet = file?.parsed.sheet ?? file?.parsed.scopes.find((s) => s.scope !== undefined)?.scope ?? sheetFromFileName(baseName(at.path));
        }
      }
    }
    workbook ??= await this.pickWorkbook();
    if (!workbook) return undefined;
    const t0 = Date.now();
    const act = new Activity("formula view", baseName(workbook.path));
    const wb = await this.ws.workbookAt(workbook);
    const readMs = Date.now() - t0;
    if (!wb) {
      // In vscode.dev only the opened folder is reachable: a project opened on its own cannot
      // see the workbook next to it. Say so instead of showing the browser's raw error.
      const name = baseName(workbook.path);
      void act.error(
        vscode.workspace.getWorkspaceFolder(workbook) === undefined
          ? `xln: ${name} is outside the open folder, so it cannot be read. Open the folder that contains both ${name} and its .xln project folder.`
          : `xln: cannot read ${name} (${this.ws.readError(workbook) ?? "unknown reason"}).`,
      );
      return undefined;
    }
    const known = sheet === undefined ? undefined : wb.sheets.find((s) => s.name.toLowerCase() === sheet!.toLowerCase());
    if (!all) {
      sheet = known?.name ?? (await this.pickSheet(wb, baseName(workbook.path)));
      if (sheet === undefined) return undefined;
    }

    const uri = all ? this.uriFor(workbook, WORKBOOK_VIEW) : this.uriFor(workbook, sheet!, order);
    // Timed from here: the pickers wait on the user.
    const t1 = Date.now();
    const b = await this.build(uri);
    act.target = `${baseName(workbook.path)} ${all ? "(every sheet)" : `'${sheet}'`}`;
    act.line(
      b
        ? `${all ? "calculation" : order} order: ${b.doc.lines.length} formula cell(s)${b.doc.linked ? ", names linked to the project" : ""}${cell !== undefined ? `; at ${cell}` : ""}; ${readMs + Date.now() - t1} ms`
        : "no view: the workbook cannot be read, or it has no such sheet",
    );
    const doc = await vscode.workspace.openTextDocument(uri);
    if (doc.languageId !== FORMULAS_LANGUAGE) await vscode.languages.setTextDocumentLanguage(doc, FORMULAS_LANGUAGE);
    const e = b && cell !== undefined ? entryForCell(b.doc, cell, all ? known?.name : undefined) : undefined;
    const selection = b && e ? this.range(b, e.address.start, e.address.end) : undefined;
    await vscode.window.showTextDocument(uri, { preview: true, ...(selection ? { selection } : {}) });
    return uri.toString();
  }

  /**
   * `xln: Switch formula view order`: the active sheet view in the other order, at the
   * same cell. Returns the new view's URI.
   */
  async toggleOrder(arg?: unknown): Promise<string | undefined> {
    const uri = arg instanceof vscode.Uri ? arg : vscode.window.activeTextEditor?.document.uri;
    if (!uri || uri.scheme !== FORMULAS_SCHEME) return this.open(undefined, "calculation");
    const t = this.target(uri);
    const ed = vscode.window.activeTextEditor;
    let cell: string | undefined;
    if (ed && ed.document.uri.toString() === uri.toString()) {
      const b = await this.build(uri);
      const e = b && entryAtLine(b, ed.selection.active.line);
      if (e) cell = b.doc.lines[e.index]!.cell;
      if (t.sheet === WORKBOOK_VIEW && e) return this.open({ workbook: t.workbook.toString(), sheet: b.doc.lines[e.index]!.sheet, cell, order: "appearance" });
    }
    if (t.sheet === WORKBOOK_VIEW) return this.open({ workbook: t.workbook.toString(), order: "appearance" });
    const args: FormulaViewArgs = { workbook: t.workbook.toString(), sheet: t.sheet, order: t.order === "calculation" ? "appearance" : "calculation" };
    if (cell !== undefined) args.cell = cell;
    return this.open(args);
  }

  private async pickWorkbook(): Promise<vscode.Uri | undefined> {
    const pulled = this.ws.projects.map((h) => h.workbookUri).filter((u): u is vscode.Uri => u !== undefined);
    if (pulled.length === 1) return pulled[0];
    return pickWorkbook("Workbook whose formulas to show", "formula view");
  }

  private async pickSheet(wb: WorkbookSnapshot, name: string): Promise<string | undefined> {
    const items = wb.sheets
      .filter((s) => s.kind !== "chartsheet")
      .map((s) => ({ label: s.name, description: `${s.formulas.length} formula cell${s.formulas.length === 1 ? "" : "s"}${s.state !== "visible" ? ` · ${s.state}` : ""}` }));
    return (await vscode.window.showQuickPick(items, { placeHolder: `Sheet of ${name}` }))?.label;
  }
}

/** The entry whose lines include document line `line`. */
function entryAtLine(b: Built, line: number) {
  let found: Built["doc"]["entries"][number] | undefined;
  for (const e of b.doc.entries) {
    if (e.line > line) break;
    found = e;
  }
  return found;
}
