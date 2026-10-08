// The audit (C1–C15) in the editor: findings as diagnostics in the Problems panel, and
// `xln: Audit workbook`, a read-only report document `xln-audit:/<workbook> (audit)?<URI>`
// rendered by the core (`renderAuditReport`), whose names and cells are links.
//
// - A name finding is attached to the name's entry in the pulled project's `.xln` files
//   (on the part of the formula it is about when it can be found there).
// - A cell finding is attached to the cell's line in the formula view of its sheet
//   (`xln-formulas:`), which VS Code renders when the problem is opened.
// - Formats, validations, Table columns and charts, and names the project no longer has,
//   are attached to the workbook file.
// A pulled project's `xln.config.json` (harness, rule severities, constants) applies.
// Every pulled project is audited from its workbook when the projects load (reload, pull,
// file changes) and when its workbook changes; the command audits any workbook.

import * as vscode from "vscode";
import { audit, auditOptions, renderAuditReport, type AuditReport, type Finding, type FindingWhere } from "@xln/core";
import { pickWorkbook } from "./commands.js";
import type { FormulaViews } from "./formulaView.js";
import { baseName } from "./inspect.js";
import { findingDef, findingLoc, problemMessage, problemSeverity } from "./model/audit.js";
import { auditLogLines } from "./model/log.js";
import { Activity } from "./log.js";
import { isWorkbookName } from "./model/pull.js";
import type { ProjectHandle, XlnWorkspace } from "./xlnWorkspace.js";

export const AUDIT_SCHEME = "xln-audit";
const SUFFIX = " (audit)";

const SEVERITY: Record<ReturnType<typeof problemSeverity>, vscode.DiagnosticSeverity> = {
  error: vscode.DiagnosticSeverity.Error,
  warning: vscode.DiagnosticSeverity.Warning,
  information: vscode.DiagnosticSeverity.Information,
};

interface Audited {
  report: AuditReport;
  /** The diagnostics of this workbook, by document URI. */
  diagnostics: Map<string, { uri: vscode.Uri; list: vscode.Diagnostic[] }>;
}

export class AuditFeature {
  private readonly collection = vscode.languages.createDiagnosticCollection("xln-audit");
  private readonly audited = new Map<string, Audited>();
  private readonly changed = new vscode.EventEmitter<vscode.Uri>();
  private readonly published = new vscode.EventEmitter<void>();
  /** After the findings are (re)published: the checks as you type leave out what the audit already says. */
  readonly onDidPublish = this.published.event;
  private pending: ReturnType<typeof setTimeout> | undefined;
  private running: Promise<void> = Promise.resolve();

  constructor(
    private readonly ws: XlnWorkspace,
    private readonly formulas: FormulaViews,
  ) {}

  register(context: vscode.ExtensionContext): void {
    context.subscriptions.push(
      this.collection,
      this.changed,
      this.published,
      vscode.workspace.registerTextDocumentContentProvider(AUDIT_SCHEME, { onDidChange: this.changed.event, provideTextDocumentContent: (uri) => this.text(uri) }),
      vscode.languages.registerDocumentLinkProvider({ scheme: AUDIT_SCHEME }, { provideDocumentLinks: (d) => this.links(d) }),
      vscode.commands.registerCommand("xln.revealName", (arg?: { root: string; key: string }) => this.revealName(arg)),
      this.ws.onDidChange(() => this.schedule()),
      { dispose: () => this.pending && clearTimeout(this.pending) },
    );
  }

  private schedule(): void {
    if (this.pending) clearTimeout(this.pending);
    this.pending = setTimeout(() => {
      this.pending = undefined;
      void this.refresh();
    }, 400);
  }

  /** Audits every pulled project's workbook again and republishes the diagnostics. */
  refresh(): Promise<void> {
    this.running = this.running.then(async () => {
      const keep = new Set<string>();
      for (const h of this.ws.projects) {
        if (!h.workbookUri) continue;
        keep.add(h.workbookUri.toString());
        await this.auditWorkbook(h.workbookUri);
      }
      // Workbooks audited by the command stay; projects that disappeared lose their diagnostics.
      for (const key of [...this.audited.keys()]) {
        if (!keep.has(key) && !this.commandAudited.has(key)) this.audited.delete(key);
      }
      this.publish();
    });
    return this.running;
  }

  private readonly commandAudited = new Set<string>();

  /** Reads and audits one workbook; undefined if it cannot be read. */
  private async auditWorkbook(workbook: vscode.Uri): Promise<Audited | undefined> {
    const wb = await this.ws.workbookAt(workbook);
    if (!wb) {
      this.audited.delete(workbook.toString());
      return undefined;
    }
    const handle = this.ws.handleForWorkbook(workbook);
    // The project's xln.config.json: its harness, rule severities and constants.
    const values = this.ws.valuesOf(wb);
    const report = audit(wb, auditOptions(handle?.audit, { workbook: baseName(workbook.path), ...(values ? { values } : {}) }));
    const diagnostics = new Map<string, { uri: vscode.Uri; list: vscode.Diagnostic[] }>();
    const put = (uri: vscode.Uri, range: vscode.Range, f: Finding): void => {
      const d = new vscode.Diagnostic(range, problemMessage(f), SEVERITY[problemSeverity(f)]);
      d.source = "xln check";
      d.code = f.rule;
      const k = uri.toString();
      const e = diagnostics.get(k) ?? { uri, list: [] };
      e.list.push(d);
      diagnostics.set(k, e);
    };
    for (const f of report.findings) {
      const at = await this.place(handle, workbook, f.where, f);
      if (at) put(at.uri, at.range, f);
      else put(workbook, new vscode.Range(0, 0, 0, 0), f);
    }
    const a: Audited = { report, diagnostics };
    this.audited.set(workbook.toString(), a);
    return a;
  }

  /** The document and range of a place: a name's entry, or a cell's formula view line. */
  private async place(handle: ProjectHandle | undefined, workbook: vscode.Uri, where: FindingWhere, f?: Finding): Promise<vscode.Location | undefined> {
    if (where.kind === "name" && handle) {
      const def = findingDef(handle.project, where);
      const loc = f ? findingLoc(handle.project, f) : def && handle.project.nameLoc(def);
      if (!loc) return undefined;
      const lines = handle.project.files.get(loc.path)!.lines;
      const a = lines.position(loc.start);
      const b = lines.position(loc.end);
      return new vscode.Location(this.ws.fileUri(handle, loc.path), new vscode.Range(a.line, a.character, b.line, b.character));
    }
    if (where.kind === "cell" && where.sheet !== undefined && where.ref !== undefined) {
      return this.formulas.cellLocation(workbook, where.sheet, where.ref);
    }
    return undefined;
  }

  /**
   * Whether the last audit of a workbook has a finding of `rule` on a cell (`ref`, or the
   * master of its shared formula). The checks as you type leave such a problem of a cell
   * statement to the audit, which shows it on the cell's line of the formula view.
   */
  hasCellFinding(workbook: vscode.Uri, rule: string, sheet: string, ref: string): boolean {
    const r = this.audited.get(workbook.toString())?.report;
    if (!r) return false;
    const s = sheet.toLowerCase();
    return r.findings.some((f) => f.rule === rule && f.where.kind === "cell" && f.where.sheet?.toLowerCase() === s && (f.where.ref === ref || f.where.range?.split(":")[0] === ref));
  }

  private publish(): void {
    this.collection.clear();
    const merged = new Map<string, { uri: vscode.Uri; list: vscode.Diagnostic[] }>();
    for (const a of this.audited.values()) {
      for (const [k, e] of a.diagnostics) {
        const m = merged.get(k) ?? { uri: e.uri, list: [] };
        m.list.push(...e.list);
        merged.set(k, m);
      }
    }
    for (const { uri, list } of merged.values()) this.collection.set(uri, list);
    this.published.fire();
  }

  // ---- the report document ------------------------------------------------------------

  reportUri(workbook: vscode.Uri): vscode.Uri {
    return vscode.Uri.from({ scheme: AUDIT_SCHEME, path: `/${baseName(workbook.path)}${SUFFIX}`, query: workbook.toString() });
  }

  private async report(uri: vscode.Uri): Promise<AuditReport | undefined> {
    const workbook = vscode.Uri.parse(uri.query);
    return (this.audited.get(workbook.toString()) ?? (await this.auditWorkbook(workbook)))?.report;
  }

  private async text(uri: vscode.Uri): Promise<string> {
    const r = await this.report(uri);
    return r ? renderAuditReport(r).text : `// No audit: ${uri.query} cannot be read.\n`;
  }

  private async links(doc: vscode.TextDocument): Promise<vscode.DocumentLink[]> {
    const r = await this.report(doc.uri);
    if (!r) return [];
    const workbook = vscode.Uri.parse(doc.uri.query);
    const handle = this.ws.handleForWorkbook(workbook);
    const out: vscode.DocumentLink[] = [];
    for (const l of renderAuditReport(r).links) {
      const w = l.where;
      let target: vscode.Uri | undefined;
      let tooltip: string | undefined;
      if (w.kind === "name" && handle && w.key !== undefined) {
        target = commandUri("xln.revealName", { root: handle.uri.toString(), key: w.key });
        tooltip = "Go to the name in the project";
      } else if (w.kind === "cell" && w.sheet !== undefined && w.ref !== undefined) {
        target = commandUri("xln.formulaView", { workbook: workbook.toString(), sheet: w.sheet, cell: w.ref });
        tooltip = "Show the cell in the formula view";
      }
      if (!target) continue;
      const link = new vscode.DocumentLink(new vscode.Range(doc.positionAt(l.start), doc.positionAt(l.end)), target);
      link.tooltip = tooltip;
      out.push(link);
    }
    return out;
  }

  /** `xln.revealName`: opens a pulled project's `.xln` file at a name. */
  private async revealName(arg?: { root: string; key: string }): Promise<string | undefined> {
    const handle = arg && this.ws.handleFor(arg.root);
    const loc = handle && (await this.place(handle, handle.workbookUri ?? handle.uri, { kind: "name", key: arg.key }));
    if (!loc) {
      void new Activity("reveal name", arg?.key).info(`xln: ${arg?.key ?? "the name"} is not in a loaded project.`);
      return undefined;
    }
    await vscode.window.showTextDocument(loc.uri, { selection: loc.range, preview: true });
    return loc.uri.toString();
  }

  /**
   * `xln: Audit workbook`: from a workbook in the Explorer, a file of a pulled project, or a
   * pick. Publishes the workbook's diagnostics and opens the report; returns its URI.
   */
  async open(arg?: unknown): Promise<string | undefined> {
    await this.ws.ready();
    let workbook: vscode.Uri | undefined;
    if (arg instanceof vscode.Uri) {
      workbook = isWorkbookName(baseName(arg.path)) ? arg : this.ws.locate(arg)?.handle.workbookUri;
    } else if (arg && typeof arg === "object" && typeof (arg as { workbook?: unknown }).workbook === "string") {
      workbook = vscode.Uri.parse((arg as { workbook: string }).workbook);
    } else {
      const ed = vscode.window.activeTextEditor?.document.uri;
      if (ed) workbook = ed.scheme === AUDIT_SCHEME ? vscode.Uri.parse(ed.query) : this.ws.locate(ed)?.handle.workbookUri;
    }
    if (!workbook) {
      const pulled = this.ws.projects.map((h) => h.workbookUri).filter((u): u is vscode.Uri => u !== undefined);
      workbook = pulled.length === 1 ? pulled[0] : await pickWorkbook("Workbook to audit", "audit");
    }
    if (!workbook) return undefined;
    const act = new Activity("audit", baseName(workbook.path));
    await this.running;
    const a = await this.auditWorkbook(workbook);
    if (!a) {
      void act.error(`xln: cannot read ${baseName(workbook.path)} (${this.ws.readError(workbook) ?? "unknown reason"}).`);
      return undefined;
    }
    this.commandAudited.add(workbook.toString());
    this.publish();
    const uri = this.reportUri(workbook);
    // The CLI's summary (`xln check …`), under the command's name.
    const [head, ...checks] = auditLogLines(a.report);
    const counts = head!.slice("xln check".length);
    act.summary([`xln audit${counts.endsWith(": none") ? counts.slice(0, -"none".length) + "no findings" : counts}`, ...checks]);
    const placed = [...a.diagnostics.values()].reduce((n, e) => n + e.list.length, 0);
    act.line(`${placed} finding(s) in the Problems panel; report: ${uri.path.slice(1)}; ${act.ms()} ms`);
    this.changed.fire(uri);
    await vscode.window.showTextDocument(uri, { preview: true });
    return uri.toString();
  }
}

function commandUri(command: string, arg: unknown): vscode.Uri {
  return vscode.Uri.parse(`command:${command}?${encodeURIComponent(JSON.stringify([arg]))}`);
}
