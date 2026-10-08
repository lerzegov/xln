// Writing help in `.xln` files (M3c): completion and signature help from the project
// (model/editor.ts), and `xln: New module`. The checks as you type are published with
// the other diagnostics (features.ts).

import * as vscode from "vscode";
import { completions, modulePrefixProblem, newModuleText, signatureHelp, type Completion } from "./model/editor.js";
import { NAMES_DIR, Project as ProjectModel, type Project } from "./model/project.js";
import type { LibraryFeature } from "./library.js";
import { projectChoices } from "./model/ux.js";
import { Activity } from "./log.js";
import type { ProjectHandle, XlnWorkspace } from "./xlnWorkspace.js";

const SELECTOR: vscode.DocumentSelector = [{ language: "xln" }];

const NAME_KIND: Record<string, vscode.CompletionItemKind> = {
  constant: vscode.CompletionItemKind.Constant,
  range: vscode.CompletionItemKind.Field,
  spill: vscode.CompletionItemKind.Field,
  table: vscode.CompletionItemKind.Struct,
  formula: vscode.CompletionItemKind.Value,
  lambda: vscode.CompletionItemKind.Function,
  unparsed: vscode.CompletionItemKind.Text,
};

export class EditorFeature {
  constructor(
    private readonly ws: XlnWorkspace,
    /** M4: library functions in completion. */
    private readonly library?: LibraryFeature,
  ) {}

  register(context: vscode.ExtensionContext): void {
    context.subscriptions.push(
      // Letters trigger suggestions by themselves; `.` after a module prefix, `!` after a sheet.
      vscode.languages.registerCompletionItemProvider(SELECTOR, { provideCompletionItems: (d, p) => this.complete(d, p) }, ".", "!"),
      vscode.languages.registerSignatureHelpProvider(
        SELECTOR,
        { provideSignatureHelp: (d, p) => this.signature(d, p) },
        { triggerCharacters: ["(", ","], retriggerCharacters: [")"] },
      ),
      vscode.commands.registerCommand("xln.newModule", (arg?: unknown) => this.newModule(arg)),
    );
  }

  /** The project and project path of a document (a loose `.xln` file gets a project of its own). */
  private locate(doc: vscode.TextDocument): { project: Project; path: string; handle: ProjectHandle | undefined } {
    this.ws.sync(doc);
    const at = this.ws.locate(doc.uri);
    if (at && at.handle.project.files.has(at.path)) return { project: at.handle.project, path: at.path, handle: at.handle };
    const project = new ProjectModel(doc.uri.toString(), undefined);
    project.setFile("names/this.xln", doc.getText());
    return { project, path: "names/this.xln", handle: undefined };
  }

  private complete(doc: vscode.TextDocument, pos: vscode.Position): vscode.CompletionList | undefined {
    const l = this.locate(doc);
    const lines = l.project.files.get(l.path)!.lines;
    const items = completions(l.project, l.path, lines.offset(pos), this.library?.loadedLibrary(l.handle));
    if (!items) return undefined;
    const range = (c: Completion) => {
      const a = lines.position(c.start);
      const b = lines.position(c.end);
      return new vscode.Range(a.line, a.character, b.line, b.character);
    };
    return new vscode.CompletionList(
      items.map((c) => {
        const item = new vscode.CompletionItem({ label: c.label, description: c.detail }, kindOf(c));
        item.insertText = c.insertText;
        item.filterText = c.filterText;
        item.sortText = c.sortText;
        item.detail = c.detail;
        if (c.documentation) item.documentation = new vscode.MarkdownString(c.documentation);
        item.range = range(c);
        if (c.retrigger) item.command = { title: "Suggest", command: "editor.action.triggerSuggest" };
        // From the library: once the call is in, its definition goes into the module file.
        if (c.library && l.handle) item.command = { title: "Add from library", command: "xln.insertLibraryDefinition", arguments: [{ root: l.handle.uri.toString(), name: c.library.name }] };
        return item;
      }),
      false,
    );
  }

  private signature(doc: vscode.TextDocument, pos: vscode.Position): vscode.SignatureHelp | undefined {
    const l = this.locate(doc);
    const s = signatureHelp(l.project, l.path, l.project.files.get(l.path)!.lines.offset(pos));
    if (!s) return undefined;
    const info = new vscode.SignatureInformation(s.label, s.documentation ? new vscode.MarkdownString(s.documentation) : undefined);
    // Parameters as label offsets: a name can occur twice in the label (`x`, `[x]`).
    let from = s.label.indexOf("(") + 1;
    info.parameters = s.params.map((p) => {
      const k = s.label.indexOf(p.label, from);
      from = k + p.label.length;
      return new vscode.ParameterInformation([k, k + p.label.length], p.documentation ? new vscode.MarkdownString(p.documentation) : undefined);
    });
    const help = new vscode.SignatureHelp();
    help.signatures = [info];
    help.activeSignature = 0;
    help.activeParameter = s.params.length ? Math.min(s.active, s.params.length - 1) : 0;
    return help;
  }

  /**
   * `xln: New module`: asks for a prefix and writes `names/<Prefix>.xln` with a header and a
   * sample LAMBDA. Nothing reaches the workbook until a build. Returns the file's URI.
   * `arg`: `{ root, prefix }` (tests); an Explorer item (its project: the context menu of
   * a project folder, its `names` folder or a file in it); or nothing (the Explorer's (+)
   * button, the Command Palette): the only project, else a pick that offers the active
   * editor's project first.
   */
  async newModule(arg?: unknown): Promise<string | undefined> {
    await this.ws.ready();
    const given = arg && typeof arg === "object" && !(arg instanceof vscode.Uri) ? (arg as { root?: unknown; prefix?: unknown }) : {};
    let handle = typeof given.root === "string" ? this.ws.handleFor(given.root) : arg instanceof vscode.Uri ? this.ws.handleAt(arg) : undefined;
    if (!handle && this.ws.projects.length === 1) handle = this.ws.projects[0];
    if (!handle && this.ws.projects.length > 1) {
      const ed = vscode.window.activeTextEditor?.document.uri;
      const active = ed ? this.ws.handleAt(ed)?.uri.toString() : undefined;
      const order = projectChoices(
        this.ws.projects.map((h) => h.uri.toString()),
        active,
      );
      const pick = await vscode.window.showQuickPick(
        order.map((root) => {
          const h = this.ws.handleFor(root)!;
          return { label: vscode.workspace.asRelativePath(h.uri), description: root === active ? "active editor" : undefined, h };
        }),
        { placeHolder: "Project to add the module to" },
      );
      handle = pick?.h;
    }
    if (!handle) {
      void new Activity("new module").info("xln: pull a workbook first; a module belongs to a project folder.");
      return undefined;
    }
    const project = handle.project;
    let prefix = typeof given.prefix === "string" ? given.prefix : undefined;
    if (prefix === undefined) {
      prefix = await vscode.window.showInputBox({
        title: "xln: New module",
        prompt: "Prefix of the module's names (FIN gives FIN.NPV, ...); the file is names/<Prefix>.xln",
        placeHolder: "FIN",
        validateInput: (v) => modulePrefixProblem(project, v.trim()),
      });
      if (prefix === undefined) return undefined;
      prefix = prefix.trim();
    }
    const problem = modulePrefixProblem(project, prefix);
    if (problem) {
      void new Activity("new module", prefix).error(`xln: ${problem}`);
      return undefined;
    }
    const uri = vscode.Uri.joinPath(handle.uri, NAMES_DIR, `${prefix}.xln`);
    await vscode.workspace.fs.writeFile(uri, new TextEncoder().encode(newModuleText(prefix)));
    new Activity("new module", prefix).line(`wrote ${vscode.workspace.asRelativePath(uri)} (a header and a sample LAMBDA; a build writes it into the workbook)`);
    // A local folder in vscode.dev has no file watching: load the new file now.
    await this.ws.reload();
    const doc = await vscode.workspace.openTextDocument(uri);
    const editor = await vscode.window.showTextDocument(doc, { preview: false });
    const line = doc.getText().split("\n").findIndex((t) => t.startsWith(`${prefix}.`));
    if (line >= 0) editor.selection = new vscode.Selection(line, 0, line, prefix.length + 1);
    return uri.toString();
  }
}

function kindOf(c: Completion): vscode.CompletionItemKind {
  switch (c.kind) {
    case "variable":
      return vscode.CompletionItemKind.Variable;
    case "module":
      return vscode.CompletionItemKind.Module;
    case "function":
      return vscode.CompletionItemKind.Function;
    default:
      return NAME_KIND[c.nameKind ?? "formula"] ?? vscode.CompletionItemKind.Value;
  }
}
