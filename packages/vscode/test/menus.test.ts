// M3e: which xln commands each menu offers where. The `when` clauses of package.json are
// evaluated here by a small evaluator of the subset they use (`==`, `!=`, `=~`, `!`, `&&`,
// `||`, parentheses), on the context keys VS Code sets for an Explorer item.
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

const pkg = JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8")) as {
  contributes: { menus: Record<string, { command: string; when?: string }[]>; commands: { command: string; icon?: string }[]; configurationDefaults: Record<string, unknown> };
};

type Ctx = Record<string, string | boolean | undefined>;

/** Evaluates a when clause of the subset package.json uses. */
export function evalWhen(when: string | undefined, ctx: Ctx): boolean {
  if (when === undefined) return true;
  const toks: string[] = [];
  let i = 0;
  while (i < when.length) {
    const c = when[i]!;
    if (c === " ") i++;
    else if (c === "(" || c === ")") toks.push(c), i++;
    else if (when.startsWith("&&", i) || when.startsWith("||", i) || when.startsWith("==", i) || when.startsWith("!=", i) || when.startsWith("=~", i)) toks.push(when.slice(i, i + 2)), (i += 2);
    else if (c === "!") toks.push("!"), i++;
    else if (c === "/") {
      // A regular expression literal runs to the next unescaped `/`.
      let j = i + 1;
      while (j < when.length && when[j] !== "/") j += when[j] === "\\" ? 2 : 1;
      toks.push(when.slice(i, j + 1));
      i = j + 1;
    } else {
      let j = i;
      while (j < when.length && !" ()!&|=".includes(when[j]!)) j++;
      toks.push(when.slice(i, j));
      i = j;
    }
  }
  let k = 0;
  const value = (t: string): string | boolean | undefined => (t === "true" ? true : t === "false" ? false : t in ctx ? ctx[t] : undefined);
  const atom = (): boolean => {
    const t = toks[k++]!;
    if (t === "!") return !atom();
    if (t === "(") {
      const v = or();
      k++;
      return v;
    }
    const op = toks[k];
    if (op === "==" || op === "!=") {
      const rhs = toks[k + 1]!;
      k += 2;
      const eq = String(value(t)) === rhs;
      return op === "==" ? eq : !eq;
    }
    if (op === "=~") {
      const lit = toks[k + 1]!;
      k += 2;
      const re = new RegExp(lit.slice(1, lit.lastIndexOf("/")), lit.slice(lit.lastIndexOf("/") + 1));
      return re.test(String(value(t) ?? ""));
    }
    return !!value(t);
  };
  const and = (): boolean => {
    let v = atom();
    while (toks[k] === "&&") (k++, (v = atom() && v));
    return v;
  };
  const or = (): boolean => {
    let v = and();
    while (toks[k] === "||") (k++, (v = and() || v));
    return v;
  };
  return or();
}

/** The context keys of an Explorer item at `path`. */
function item(path: string, folder = false, extra: Ctx = {}): Ctx {
  const name = path.slice(path.lastIndexOf("/") + 1);
  const dot = name.lastIndexOf(".");
  return { resourcePath: path, resourceFilename: name, resourceExtname: dot > 0 ? name.slice(dot) : "", explorerResourceIsFolder: folder, isWeb: false, ...extra };
}

const xln = (cmds: string[]) => cmds.filter((c) => c.startsWith("xln.") && c !== "xln.inspectWorkbook" && c !== "xln.writeTestFile").sort();
const shown = (menu: string, ctx: Ctx) => xln(pkg.contributes.menus[menu]!.filter((m) => evalWhen(m.when, ctx)).map((m) => m.command));

describe("menus (M3e)", () => {
  it("the evaluator reads the clauses as VS Code does", () => {
    expect(evalWhen("a == .x || b", { a: ".x" })).toBe(true);
    expect(evalWhen("!isWeb && (a == 1 || b == 2)", { isWeb: false, b: "2" })).toBe(true);
    expect(evalWhen("!isWeb && (a == 1 || b == 2)", { isWeb: true, b: "2" })).toBe(false);
    expect(evalWhen("p =~ /\\.xln\\//", { p: "/w/b.xln/names" })).toBe(true);
  });

  it("a workbook: pull, the formula views, audit; no build, no new module", () => {
    for (const p of ["/w/book/lbo.xlsx", "/w/book/lbo.xlsm"]) {
      expect(shown("explorer/context", item(p))).toEqual(["xln.auditWorkbook", "xln.formulaView", "xln.formulaViewCalc", "xln.libraryStatus", "xln.pullWorkbook", "xln.workbookFormulaView"]);
    }
  });

  it("an .xln file, a project folder, a folder or file inside one: build, build and reopen, new module; no pull", () => {
    const want = ["xln.buildAndReopen", "xln.buildWorkbook", "xln.insertLibraryFunction", "xln.libraryStatus", "xln.newModule"];
    expect(shown("explorer/context", item("/w/book/lbo.xln/names/FN.xln"))).toEqual(want);
    expect(shown("explorer/context", item("/w/book/lbo.xln", true))).toEqual(want);
    expect(shown("explorer/context", item("/w/book/lbo.xln/names", true))).toEqual(want);
    expect(shown("explorer/context", item("/w/book/lbo.xln/names/sheets", true))).toEqual(want);
    expect(shown("explorer/context", item("/w/book/lbo.xln/xln.config.json"))).toEqual(want);
    expect(shown("explorer/context", item("/w/loose.xln"))).toEqual(want);
  });

  it("an .xln editor: library status, insert library function, publish to library (M4)", () => {
    const lib = shown("editor/context", { editorLangId: "xln" }).filter((c) => c.toLowerCase().includes("library"));
    expect(lib).toEqual(["xln.insertLibraryFunction", "xln.libraryStatus", "xln.publishToLibrary"]);
  });

  it("in the browser: no build and reopen", () => {
    expect(shown("explorer/context", item("/w/book/lbo.xln", true, { isWeb: true }))).toEqual(["xln.buildWorkbook", "xln.insertLibraryFunction", "xln.libraryStatus", "xln.newModule"]);
  });

  it("anything else: nothing of xln's", () => {
    expect(shown("explorer/context", item("/w/book", true))).toEqual([]);
    expect(shown("explorer/context", item("/w/notes.txt"))).toEqual([]);
    expect(shown("explorer/context", item("/w/book/lbo.xlsx.bak"))).toEqual([]);
  });

  it("the Explorer's (+) New module button, only when the workspace has a project", () => {
    const view = (extra: Ctx) => shown("view/title", { view: "workbench.explorer.fileView", ...extra });
    expect(view({ "xln.hasProject": true })).toEqual(["xln.newModule"]);
    expect(view({ "xln.hasProject": false })).toEqual([]);
    expect(shown("view/title", { view: "outline", "xln.hasProject": true })).toEqual([]);
    expect(pkg.contributes.commands.find((c) => c.command === "xln.newModule")?.icon).toBe("$(add)");
  });

  it("internals: the manifest, the lockfile and library-bases/ excluded and read-only by default; xln.config.json neither", () => {
    const d = pkg.contributes.configurationDefaults as Record<string, Record<string, boolean>>;
    const want = { "**/*.xln/workbook.manifest.json": true, "**/*.xln/xln.lock.json": true };
    expect(d["files.exclude"]).toEqual({ ...want, "**/*.xln/library-bases": true });
    expect(d["files.readonlyInclude"]).toEqual({ ...want, "**/*.xln/library-bases/**": true });
    expect(JSON.stringify(d)).not.toContain("xln.config.json");
  });

  it("the palette offers Show or Hide project internals, whichever applies", () => {
    const pal = (shownNow: boolean) => shown("commandPalette", { "xln.internalsShown": shownNow }).filter((c) => c.includes("Internals"));
    expect(pal(false)).toEqual(["xln.showInternals"]);
    expect(pal(true)).toEqual(["xln.hideInternals"]);
  });

  it("the formula view stays on a sheet file's editor title", () => {
    const title = (path: string) => shown("editor/title", { ...item(path), resourceScheme: "file" });
    expect(title("/w/book/lbo.xln/names/sheets/IS.xln")).toEqual(["xln.formulaView"]);
    expect(title("/w/book/lbo.xln/names/FN.xln")).toEqual([]);
  });
});
