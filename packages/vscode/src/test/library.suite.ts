// M4: the library end to end, in the web extension host and in desktop VS Code, on the
// project the browse suite pulled (book/f7_base.xln) with a library folder written into
// the test workspace (lib/): status, completion from the library, insert, the code
// lenses, take the library's version, publish; the kept bases (library-bases/) and
// Record library base.
import * as vscode from "vscode";
import { readWorkbook } from "@xln/core";

function check(cond: unknown, msg: string): asserts cond {
  if (!cond) throw new Error(msg);
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

async function until<T>(get: () => Promise<T | undefined> | T | undefined, what: string, ms = 5000): Promise<T> {
  for (let t = 0; t < ms; t += 100) {
    const v = await get();
    if (v !== undefined && v !== false) return v;
    await sleep(100);
  }
  throw new Error(`timed out: ${what}`);
}

const enc = (s: string) => new TextEncoder().encode(s);
const dec = (b: Uint8Array) => new TextDecoder().decode(b);
const lambda = (name: string, summary: string, def: string) => `# name       ${name}\n# summary    ${summary}\n# params     x\n#\n# Why ${name}.\n\n${def}\n`;

interface StatusResult {
  uri: string;
  report: { items: { name: string; state: string; base?: string; libraryHash?: string }[] };
}

export function defineLibrarySuite(): void {
  suite("xln library (M4)", () => {
    const folder = () => vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri);
    const project = () => vscode.Uri.joinPath(folder(), "book", "f7_base.xln");
    const lib = () => vscode.Uri.joinPath(folder(), "lib");
    const config = () => vscode.Uri.joinPath(project(), "xln.config.json");
    const fnFile = () => vscode.Uri.joinPath(project(), "names", "FN.xln");
    const scratch = () => vscode.Uri.joinPath(project(), "names", "m4.xln");
    const basesDir = () => vscode.Uri.joinPath(project(), "library-bases");
    const keptBases = async (): Promise<string[]> => {
      try {
        return (await vscode.workspace.fs.readDirectory(basesDir())).map(([n]) => n).sort();
      } catch {
        return [];
      }
    };
    const root = () => project().toString();
    let savedConfig = "";

    const states = async (): Promise<Record<string, string>> => {
      const r = await vscode.commands.executeCommand<StatusResult>("xln.libraryStatus", { root: root() });
      check(r, "no status");
      return Object.fromEntries(r.report.items.filter((i) => i.name.startsWith("FN.")).map((i) => [i.name, i.state]));
    };

    suiteSetup(async () => {
      await vscode.workspace.fs.createDirectory(lib());
      await vscode.workspace.fs.writeFile(vscode.Uri.joinPath(lib(), "FN.TWICE.lambda"), enc(lambda("FN.TWICE", "Doubles.", "LAMBDA(x, 2 * x)")));
      await vscode.workspace.fs.writeFile(vscode.Uri.joinPath(lib(), "FN.QUAD.lambda"), enc(lambda("FN.QUAD", "Quadruples.", "LAMBDA(x, FN.TWICE(FN.TWICE(x)))")));
      savedConfig = dec(await vscode.workspace.fs.readFile(config()));
      await vscode.workspace.fs.writeFile(config(), enc(JSON.stringify({ ...JSON.parse(savedConfig), library: "../../lib" }, null, 2)));
      await vscode.workspace.fs.writeFile(scratch(), enc("M4 = FN.;\n"));
      await vscode.commands.executeCommand("xln.reload");
    });

    suiteTeardown(async () => {
      await vscode.commands.executeCommand("workbench.action.closeAllEditors");
      await vscode.workspace.fs.writeFile(config(), enc(savedConfig));
      for (const f of [fnFile(), scratch(), lib(), basesDir()]) {
        try {
          await vscode.workspace.fs.delete(f, { recursive: true });
        } catch {
          // not written
        }
      }
      await vscode.commands.executeCommand("xln.reload");
    });

    test("status: the library's functions are missing from the project", async () => {
      check(JSON.stringify(await states()) === JSON.stringify({ "FN.QUAD": "missing", "FN.TWICE": "missing" }), "states");
      const doc = vscode.window.activeTextEditor?.document;
      check(doc?.uri.scheme === "xln-lib" && doc.getText().includes("missing (2)"), `report: ${doc?.getText()}`);
      const links = await vscode.commands.executeCommand<vscode.DocumentLink[]>("vscode.executeLinkProvider", doc.uri);
      check(links.some((l) => l.target?.path.endsWith("/FN.TWICE.lambda")), `links: ${links.map((l) => l.target?.toString()).join(", ")}`);
    });

    test("completion offers them, marked from library, and accepting adds the definitions", async () => {
      const doc = await vscode.workspace.openTextDocument(scratch());
      const pos = doc.positionAt(doc.getText().indexOf("FN.") + 3);
      const item = await until(async () => {
        const list = await vscode.commands.executeCommand<vscode.CompletionList>("vscode.executeCompletionItemProvider", doc.uri, pos);
        return list.items.find((i) => (typeof i.label === "string" ? i.label : i.label.label) === "FN.QUAD");
      }, "FN.QUAD from the library in completion");
      check(String(item.detail).startsWith("from library"), `detail ${item.detail}`);
      check(item.command?.command === "xln.insertLibraryDefinition", `command ${item.command?.command}`);
      const r = await vscode.commands.executeCommand<{ names: string[]; files: string[] }>(item.command!.command, ...(item.command!.arguments ?? []));
      check(r && r.names.join(",") === "FN.TWICE,FN.QUAD" && r.files.join(",") === "names/FN.xln", `inserted ${JSON.stringify(r)}`);
      const fn = await vscode.workspace.openTextDocument(fnFile());
      check(fn.getText().includes("FN.QUAD = LAMBDA(x, FN.TWICE(FN.TWICE(x)));") && fn.getText().includes(" * Doubles."), fn.getText());
      await fn.save();
      check(JSON.stringify(await states()) === JSON.stringify({ "FN.QUAD": "identical", "FN.TWICE": "identical" }), "identical after insert");
      // Insert kept the two bases' texts (for a three-way diff later).
      const st = await vscode.commands.executeCommand<StatusResult>("xln.libraryStatus", { root: root() });
      const want = st!.report.items.filter((i) => i.name.startsWith("FN.")).map((i) => `${i.libraryHash}.json`).sort();
      check(JSON.stringify(await keptBases()) === JSON.stringify(want), `kept ${JSON.stringify(await keptBases())}, want ${JSON.stringify(want)}`);
    });

    test("an edited entry is modified (its base is the library's): lenses, take the library's version after a question", async () => {
      const fn = await vscode.workspace.openTextDocument(fnFile());
      const edit = new vscode.WorkspaceEdit();
      const at = fn.getText().indexOf("2 * x");
      edit.replace(fn.uri, new vscode.Range(fn.positionAt(at), fn.positionAt(at + 5)), "x + x");
      await vscode.workspace.applyEdit(edit);
      await fn.save();
      // Insert recorded the library's version as the base: an edit here reads as modified.
      check(fn.getText().includes("@from(lib #"), fn.getText());
      check((await states())["FN.TWICE"] === "modified", "modified after the edit");
      const lenses = await until(async () => {
        const l = await vscode.commands.executeCommand<vscode.CodeLens[]>("vscode.executeCodeLensProvider", fn.uri);
        // Not in the workbook yet (never built): the lens says so too.
        return l.some((x) => x.command?.title === "library: modified · not built yet") ? l : undefined;
      }, "library lenses");
      const titles = lenses.map((l) => l.command?.title);
      check(titles.includes("Take the library's version (undo the edit)") && titles.includes("Publish to library"), titles.join(" | "));
      const take = lenses.find((l) => l.command?.title === "Take the library's version (undo the edit)")!;
      const arg = take.command!.arguments![0] as object;
      // Answered no: the edit stays.
      check(!(await vscode.commands.executeCommand<boolean>(take.command!.command, { ...arg, confirm: false })), "take refused");
      check(fn.getText().includes("FN.TWICE = LAMBDA(x, x + x);"), fn.getText());
      check(await vscode.commands.executeCommand<boolean>(take.command!.command, { ...arg, confirm: true }), "take");
      check(fn.getText().includes("FN.TWICE = LAMBDA(x, 2 * x);"), fn.getText());
      await fn.save();
      check((await states())["FN.TWICE"] === "identical", "identical after take");
    });

    test("publish writes the library file, keeping its header; then identical", async () => {
      const fn = await vscode.workspace.openTextDocument(fnFile());
      const edit = new vscode.WorkspaceEdit();
      const at = fn.getText().indexOf("2 * x");
      edit.replace(fn.uri, new vscode.Range(fn.positionAt(at), fn.positionAt(at + 5)), "x * 2");
      await vscode.workspace.applyEdit(edit);
      await fn.save();
      const r = await vscode.commands.executeCommand<{ written: boolean; changed: string[] }>("xln.publishToLibrary", { root: root(), name: "FN.TWICE", confirm: true });
      check(r?.written && r.changed.join(",") === "definition", `publish ${JSON.stringify(r)}`);
      const text = dec(await vscode.workspace.fs.readFile(vscode.Uri.joinPath(lib(), "FN.TWICE.lambda")));
      check(text === lambda("FN.TWICE", "Doubles.", "LAMBDA(x, x * 2)"), text);
      // The entry records the published version as its base (an unsaved edit).
      const st = await vscode.commands.executeCommand<StatusResult>("xln.libraryStatus", { root: root() });
      const item = st!.report.items.find((i) => i.name === "FN.TWICE")!;
      check(item.state === "identical" && item.base === item.libraryHash, `after publish ${JSON.stringify(item)}`);
      check(fn.getText().includes(`@from(lib #${item.libraryHash})`), fn.getText());
      await fn.save();
    });

    // Feedback 2026-10-07: Take, save, Build wrote nothing and said nothing; a Build with the
    // Take unsaved did the same. Each Build here must write what the editor shows.
    test("Take the library's version, save, Build: written; an unsaved edit: Build saves it and writes it", async () => {
      const web = vscode.env.uiKind === vscode.UIKind.Web;
      const wb = vscode.Uri.joinPath(folder(), "book", "f7_base.xlsx");
      const lock = vscode.Uri.joinPath(project(), "xln.lock.json");
      const written = web ? vscode.Uri.joinPath(folder(), "book", "f7_base.xln.xlsx") : wb;
      const before = await vscode.workspace.fs.readFile(wb);
      const lockBefore = await vscode.workspace.fs.readFile(lock);
      const definition = async (name: string) => readWorkbook(await vscode.workspace.fs.readFile(written)).definedNames.find((d) => d.name === name)?.definition;
      const replace = async (doc: vscode.TextDocument, from: string, to: string) => {
        const at = doc.getText().indexOf(from);
        check(at >= 0, `${from} not in ${doc.getText()}`);
        const edit = new vscode.WorkspaceEdit();
        edit.replace(doc.uri, new vscode.Range(doc.positionAt(at), doc.positionAt(at + from.length)), to);
        check(await vscode.workspace.applyEdit(edit), "edit");
      };
      try {
        // The completion test left `M4 = FN.QUAD…` half written: a plain constant, unsaved.
        const scratchDoc = await vscode.workspace.openTextDocument(scratch());
        const all = new vscode.WorkspaceEdit();
        all.replace(scratch(), new vscode.Range(scratchDoc.positionAt(0), scratchDoc.positionAt(scratchDoc.getText().length)), "Four = 4;\n");
        await vscode.workspace.applyEdit(all);
        const fn = await vscode.workspace.openTextDocument(fnFile());
        await vscode.window.showTextDocument(fn);
        await replace(fn, "x * 2", "x + x");
        await fn.save();
        const lenses = await until(async () => {
          const l = await vscode.commands.executeCommand<vscode.CodeLens[]>("vscode.executeCodeLensProvider", fn.uri);
          return l.some((x) => x.command?.title?.startsWith("Take the library's version")) ? l : undefined;
        }, "the Take lens");
        const take = lenses.find((l) => l.command?.title?.startsWith("Take the library's version"))!;
        check(await vscode.commands.executeCommand<boolean>(take.command!.command, { ...(take.command!.arguments![0] as object), confirm: true }), "take");
        await fn.save();
        const r1 = await vscode.commands.executeCommand<{ status: string; written?: string } | undefined>("xln.buildWorkbook", fn.uri, { confirm: false });
        check(r1?.status === "built", `first build ${JSON.stringify(r1)}`);
        check((await definition("FN.TWICE"))?.includes("_xlpm.x * 2"), `FN.TWICE in the workbook: ${await definition("FN.TWICE")}`);

        // Unsaved: the Build saves it first, then writes it.
        await replace(fn, "x * 2", "x * 3");
        check(fn.isDirty, "the edit is unsaved");
        const r2 = await vscode.commands.executeCommand<{ status: string } | undefined>("xln.buildWorkbook", fn.uri, { confirm: false });
        check(r2?.status === "built", `second build ${JSON.stringify(r2)}`);
        check(!fn.isDirty && dec(await vscode.workspace.fs.readFile(fnFile())).includes("x * 3"), "not saved");
        check((await definition("FN.TWICE"))?.includes("_xlpm.x * 3"), `FN.TWICE in the workbook: ${await definition("FN.TWICE")}`);
      } finally {
        await vscode.workspace.fs.writeFile(wb, before);
        await vscode.workspace.fs.writeFile(lock, lockBefore);
        if (web) {
          try {
            await vscode.workspace.fs.delete(written);
          } catch {
            // not written
          }
        } else {
          try {
            await vscode.workspace.fs.delete(vscode.Uri.joinPath(folder(), "book", "f7_base.backup.xlsx"));
          } catch {
            // no backup
          }
        }
      }
    });

    // Author's idea, 2026-10-07: a copy identical to the library with no @from (inserted
    // before the base existed) gets one only from an explicit action.
    test("Record library base: a lens on an identical entry without @from; the base written unsaved, its text kept", async () => {
      await vscode.workspace.fs.writeFile(vscode.Uri.joinPath(lib(), "FN.HAND.lambda"), enc(lambda("FN.HAND", "By hand.", "LAMBDA(x, x - 1)")));
      const fn = await vscode.workspace.openTextDocument(fnFile());
      const edit = new vscode.WorkspaceEdit();
      edit.insert(fn.uri, fn.positionAt(fn.getText().length), "\nFN.HAND = LAMBDA(x, x - 1);\n");
      await vscode.workspace.applyEdit(edit);
      await fn.save();
      const st = await vscode.commands.executeCommand<StatusResult & { report: { items: { note?: string }[] } }>("xln.libraryStatus", { root: root() });
      const item = st!.report.items.find((i) => i.name === "FN.HAND") as { state: string; base?: string; libraryHash?: string; note?: string } | undefined;
      check(item?.state === "identical" && item.base === undefined && item.note?.includes("no base recorded: Record library base"), `status ${JSON.stringify(item)}`);
      // Status itself wrote nothing.
      check(!fn.getText().includes("@from(lib #" + item.libraryHash + ")\nFN.HAND"), fn.getText());
      const lens = await until(async () => {
        const l = await vscode.commands.executeCommand<vscode.CodeLens[]>("vscode.executeCodeLensProvider", fn.uri);
        return l.find((x) => x.command?.title === "Record library base");
      }, "the Record library base lens");
      const hash = await vscode.commands.executeCommand<string>(lens.command!.command, ...(lens.command!.arguments ?? []));
      check(hash === item.libraryHash, `recorded ${hash}`);
      check(fn.isDirty && fn.getText().includes(`@from(lib #${hash})\nFN.HAND = LAMBDA(x, x - 1);`), fn.getText());
      check((await keptBases()).includes(`${hash}.json`), `kept ${JSON.stringify(await keptBases())}`);
      await fn.save();
      const after = await vscode.commands.executeCommand<StatusResult>("xln.libraryStatus", { root: root() });
      const hand = after!.report.items.find((i) => i.name === "FN.HAND")!;
      check(hand.state === "identical" && hand.base === hash, `after ${JSON.stringify(hand)}`);
    });
  });
}
