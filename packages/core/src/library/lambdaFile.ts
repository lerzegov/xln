// M4: a library is a folder of `.lambda` files, one LAMBDA each (the author's
// `_shared/lib`, which the Python build also reads, so the format is read as it is):
//
//   # name       FN.GROW
//   # summary    Growth chain: seed the first period, compound it thereafter.
//   # params     seed, growth, periods
//   # impromptu  x[year:first] = <seed>
//   #            x[year:rest]  = x[year:prev] * (1 + <g>)
//   # example    FN.GROW(10000, 0.05, 5)
//   #
//   # Free rationale paragraphs, as many as needed.
//
//   LAMBDA(seed, growth, periods,
//       ...)
//
// The header is the run of `#` lines up to the first blank `#` line: a field starts with
// `# key value` (one space after `#`), a continuation line is indented further. The
// rationale is the `#` lines after that, the definition (display form) everything from
// the first line that is not a comment. The header's `params` is checked against the
// LAMBDA's parameters, as the Python reader does. A rationale line `@param x text` gives
// x a description (xln's addition, written by publish; the Python build ignores the
// rationale).
//
// The header's summary and params become the doc comment xln writes into the project
// (and so the Name Manager comment): `summary` then one `@param` line per parameter.

import { parseDocComment } from "../project/doc.js";
import { COMMENT_MAX, commentLength } from "../project/provenance.js";
import { compileWithDiagnostics } from "../lang/transform.js";
import { tryParse, stripPrefix } from "../lang/parser.js";

export interface LibraryField {
  key: string;
  /** Continuation lines joined with LF, each without the field's indentation. */
  value: string;
  /** 1-based line of the `# key` line. */
  line: number;
  /** 1-based line of the field's last line. */
  endLine: number;
}

export interface LibraryProblem {
  severity: "error" | "warning";
  /** The `.lambda` file, as given to the reader. */
  path: string;
  /** 1-based; absent for a problem of the whole file. */
  line?: number;
  message: string;
}

export interface LibraryFunction {
  name: string;
  /** The `summary` field, continuation lines joined with a space. */
  summary: string;
  /** The LAMBDA's parameters (checked against the `params` field), `[x]` for an optional one. */
  params: string[];
  /** Descriptions of parameters, from `@param` lines of the rationale. */
  paramDocs: Record<string, string>;
  /** Every header field in file order, `name`, `summary` and `params` included. */
  fields: LibraryField[];
  /** The rationale paragraphs, `# ` removed (a blank `#` line is an empty line). */
  rationale: string;
  /** The LAMBDA in display form, as written (LF line breaks). */
  definition: string;
  /** 1-based line where the definition starts. */
  definitionLine: number;
  /** The doc comment for the project: summary and `@param` lines, within the Name Manager's limit. */
  doc: string;
  /** The doc comment had to be shortened to fit (see `libraryDoc`). */
  docShortened: boolean;
  /** The file, as given to the reader. */
  path: string;
}

export interface ParsedLambdaFile {
  fn?: LibraryFunction;
  problems: LibraryProblem[];
}

/**
 * What a doc comment written from the library may take of the Name Manager's 255
 * characters: the rest is left for the build's provenance tag with its library base
 * (` [xln FN 1.2 #3f9a1c lib#353921]`, 32 characters with a two-letter module and a
 * three-character version; 11 more are left for longer ones). A comment too long to
 * carry the tag goes without it, and then the workbook does not record the base; Insert
 * and Take measure the actual tag (`docTagWarnings`) and say so before they write.
 */
export const LIBRARY_DOC_MAX = COMMENT_MAX - 43;

function isBlank(s: string): boolean {
  return s.trim() === "";
}

function isWs(c: string | undefined): boolean {
  return c === " " || c === "\t";
}

function isKey(s: string): boolean {
  if (s === "") return false;
  for (let k = 0; k < s.length; k++) {
    const c = s[k]!;
    const ok = (c >= "a" && c <= "z") || (k > 0 && ((c >= "0" && c <= "9") || c === "_" || c === "-"));
    if (!ok) return false;
  }
  return true;
}

/** The file name a function is published under: `FN.GROW.lambda`. */
export function lambdaFileName(name: string): string {
  return `${name}.lambda`;
}

/** The LAMBDA's parameters, `[x]` for an optional one; undefined when the text is not one LAMBDA. */
export function lambdaParameters(display: string): string[] | undefined {
  let e = tryParse(display).formula?.body;
  while (e && e.kind === "paren") e = e.expr;
  if (!e || e.kind !== "lambda") return undefined;
  return e.params.map((p) => {
    const n = stripPrefix(p.name.text).base;
    return p.optional ? `[${n}]` : n;
  });
}

function bare(p: string): string {
  return p.startsWith("[") && p.endsWith("]") ? p.slice(1, -1) : p;
}

/** Collapses runs of blanks and line breaks into one space. */
export function collapseSpace(s: string): string {
  return s.split(/\s+/).filter((w) => w !== "").join(" ");
}

/**
 * The doc comment of a library function: its summary, then `@param name text` per
 * parameter. Over `max` characters (a line break counts 2, as in the Name Manager) the
 * summary is cut at a word with `…`; if the `@param` lines alone are too long they go.
 * The definition is never touched.
 */
export function libraryDoc(summary: string, params: readonly string[], paramDocs: Readonly<Record<string, string>> = {}, max = LIBRARY_DOC_MAX): { doc: string; shortened: boolean } {
  const paramLines = params.map((p) => {
    const d = paramDocs[bare(p).toLowerCase()];
    return d ? `@param ${p} ${d}` : `@param ${p}`;
  });
  const join = (s: string, ps: string[]) => [s, ...ps].filter((l) => l !== "").join("\n");
  const full = join(summary, paramLines);
  if (commentLength(full) <= max) return { doc: full, shortened: false };
  let lines = paramLines;
  // Descriptions go before the parameter names do: the names drive signature help.
  if (commentLength(join("", lines)) + 12 > max) lines = params.map((p) => `@param ${p}`);
  if (commentLength(join("", lines)) + 12 > max) lines = [];
  const room = max - (lines.length ? commentLength(join("", lines)) + 2 : 0);
  if (commentLength(summary) <= room) return { doc: join(summary, lines), shortened: true };
  const words = summary.split(" ");
  let cut = "";
  for (const w of words) {
    const next = cut === "" ? w : `${cut} ${w}`;
    if (commentLength(next) + 1 > room) break;
    cut = next;
  }
  return { doc: join(cut === "" ? summary.slice(0, Math.max(0, room - 1)) + "…" : cut + "…", lines), shortened: true };
}

/**
 * Reads one `.lambda` file. Never throws: what is wrong comes back in `problems`, with
 * line numbers; `fn` is absent when the file cannot be used (no name, no LAMBDA).
 * `known`: other names the definition may call (the rest of the library), for the
 * compile check.
 */
export function parseLambdaFile(path: string, text: string, known: Iterable<string> = []): ParsedLambdaFile {
  const problems: LibraryProblem[] = [];
  const problem = (severity: "error" | "warning", message: string, line?: number) =>
    problems.push(line === undefined ? { severity, path, message } : { severity, path, line, message });
  const lines = text.split("\n").map((l) => (l.endsWith("\r") ? l.slice(0, -1) : l));

  const fields: LibraryField[] = [];
  const rationale: string[] = [];
  let k = 0;
  // Blank lines before the header.
  while (k < lines.length && isBlank(lines[k]!)) k++;
  let inHeader = true;
  let current: { field: LibraryField; indent: number; parts: string[] } | undefined;
  const close = () => {
    if (!current) return;
    current.field.value = current.parts.join("\n");
    fields.push(current.field);
    current = undefined;
  };
  for (; k < lines.length; k++) {
    const raw = lines[k]!;
    const t = raw.trimStart();
    if (!t.startsWith("#")) {
      if (isBlank(raw)) {
        // A blank line ends the comment block only when the definition follows.
        let j = k;
        while (j < lines.length && isBlank(lines[j]!)) j++;
        if (j < lines.length && lines[j]!.trimStart().startsWith("#")) {
          if (!inHeader) rationale.push("");
          continue;
        }
      }
      break;
    }
    const content = t.slice(1);
    if (inHeader) {
      if (isBlank(content)) {
        close();
        inHeader = false;
        continue;
      }
      if (content.startsWith(" ") && !isWs(content[1])) {
        const body = content.slice(1);
        let e = 0;
        while (e < body.length && !isWs(body[e])) e++;
        const key = body.slice(0, e);
        if (isKey(key)) {
          close();
          let v = e;
          while (isWs(body[v])) v++;
          current = { field: { key, value: "", line: k + 1, endLine: k + 1 }, indent: 1 + v, parts: [body.slice(v).trimEnd()] };
          continue;
        }
      }
      if (current) {
        // A continuation line: indented up to the value's column (or less, in which case all of it goes).
        let i = 0;
        while (i < current.indent && isWs(content[i])) i++;
        current.parts.push(content.slice(i).trimEnd());
        current.field.endLine = k + 1;
        continue;
      }
      problem("warning", `not a header field (# key value); read as rationale`, k + 1);
      inHeader = false;
      rationale.push(content.startsWith(" ") ? content.slice(1) : content);
      continue;
    }
    rationale.push(content.startsWith(" ") ? content.slice(1) : content);
  }
  close();
  while (rationale.length && rationale[rationale.length - 1] === "") rationale.pop();
  while (rationale.length && rationale[0] === "") rationale.shift();

  while (k < lines.length && isBlank(lines[k]!)) k++;
  const definitionLine = k + 1;
  const defLines = lines.slice(k);
  while (defLines.length && isBlank(defLines[defLines.length - 1]!)) defLines.pop();
  const definition = defLines.map((l) => l.trimEnd()).join("\n");

  const field = (key: string) => fields.find((f) => f.key === key);
  for (const f of fields) {
    if (fields.filter((g) => g.key === f.key).length > 1 && field(f.key) !== f) problem("warning", `field ${f.key} given twice; the first one counts`, f.line);
  }
  const nameField = field("name");
  const name = nameField?.value.split("\n")[0]!.trim() ?? "";
  if (!nameField || name === "") problem("error", "the header has no name (# name FN.X)", nameField?.line ?? 1);
  if (definition === "") problem("error", "no definition after the header (a blank line, then the LAMBDA)", definitionLine);
  if (name !== "") {
    const file = path.slice(Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\")) + 1);
    if (file.toLowerCase().endsWith(".lambda") && file.slice(0, -".lambda".length).toLowerCase() !== name.toLowerCase()) {
      problem("warning", `the file is named ${file} but the header names ${name}`, nameField!.line);
    }
  }

  let params: string[] = [];
  if (definition !== "") {
    const parsed = tryParse(definition);
    const actual = lambdaParameters(definition);
    if (parsed.diagnostics.length) {
      const d = parsed.diagnostics[0]!;
      problem("error", `the definition does not parse: ${d.message}`, definitionLine + lineOf(definition, d.start));
    } else if (!actual) {
      problem("error", "the definition must be a single LAMBDA", definitionLine);
    } else {
      params = actual;
      const pf = field("params");
      if (pf) {
        const declared = pf.value
          .split("\n")[0]!
          .split(",")
          .map((s) => s.trim())
          .filter((s) => s !== "");
        const same = declared.length === actual.length && declared.every((p, i) => bare(p).toLowerCase() === bare(actual[i]!).toLowerCase());
        if (!same) problem("error", `the header declares params ${declared.join(", ") || "(none)"} but the LAMBDA takes ${actual.join(", ") || "none"}`, pf.line);
      } else problem("warning", "the header has no params line", nameField?.line ?? 1);
      const c = compileWithDiagnostics(definition, { names: [...known, name] });
      for (const d of c.diagnostics) {
        if (d.severity !== "error") continue;
        problem("error", `the definition does not compile: ${d.message}`, definitionLine + lineOf(definition, d.start));
      }
    }
  }
  if (!field("summary")) problem("warning", "the header has no summary", nameField?.line ?? 1);

  if (name === "" || definition === "" || !lambdaParameters(definition)) return { problems };

  const summary = collapseSpace(field("summary")?.value ?? "");
  const paramDocs: Record<string, string> = {};
  const rationaleText = rationale.join("\n");
  for (const p of parseDocComment(paramLinesOf(rationale)).params) paramDocs[p.name.toLowerCase()] = p.text;
  const { doc, shortened } = libraryDoc(summary, params, paramDocs);
  if (shortened) problem("warning", `the doc comment (summary and @param lines) is over ${LIBRARY_DOC_MAX} characters: shortened for the Name Manager`, field("summary")?.line);
  return {
    fn: { name, summary, params, paramDocs, fields, rationale: rationaleText, definition, definitionLine, doc, docShortened: shortened, path },
    problems,
  };
}

/** The rationale's `@param` lines (and their continuations), for `parseDocComment`. */
function paramLinesOf(rationale: string[]): string {
  const out: string[] = [];
  let on = false;
  for (const l of rationale) {
    const t = l.trim();
    if (t.startsWith("@param ")) {
      on = true;
      out.push(t);
    } else if (on && t !== "" && (l.startsWith(" ") || l.startsWith("\t"))) out.push(t);
    else on = false;
  }
  return out.join("\n");
}

function lineOf(text: string, offset: number): number {
  let n = 0;
  for (let i = 0; i < offset && i < text.length; i++) if (text[i] === "\n") n++;
  return n;
}
