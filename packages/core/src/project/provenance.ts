// D6, provenance: a name that comes from a module carries a tag at the end of its Name
// Manager comment, `[xln FN 1.2 #3f9a1c]`: the module, its version when the module file
// declares one, and a short hash of the name's source (definition and comment). Excel
// shows the tag to whoever opens the Name Manager; the source never does, because pull
// strips it. The hash lets a later pull (or `lib status`) tell a name unchanged since the
// build from one edited in Excel.
//
// - The module is the module file the name is declared in: `names/FN.xln` gives `FN`
//   (pull's `~2` suffix for a clashing file name is dropped). `_unmanaged.xln` and the
//   sheet files under `names/sheets/` are not modules: their names get no tag.
// - The version comes from a header line of the module file, before its first
//   definition: `// @version 1.2`. Without one the tag has no version: `[xln FN #3f9a1c]`.
// - The hash is the first 6 hex digits of SHA-256 over the definition in stored form,
//   normalised as the lockfile hashes it (layout, optional sheet quotes and the spelling
//   of a number do not count), and the comment without its tag. A tag change is therefore never a change of the
//   comment: plan, lockfile and read-back compare comments with the tag stripped.
// - The library base (decided 2026-10-07): a name taken from the library records, in
//   its source, the library definition it came from: `@from(lib #353921)` on the entry,
//   written only by Insert, Take and Publish. The build carries it in the tag,
//   `[xln FN 1.2 #636cf1 lib#353921]`, and pull writes it back, so `lib status` can tell
//   "the library moved" (the copy is still its base) from "edited here" (the library is
//   still the base). The tag's `#hash` stays what was built; `lib#` is the base. The base
//   hash is `libraryHash`: the definition alone in stored form, normalised like the
//   lockfile's (layout and the spelling of numbers do not count). A tag without `lib#`
//   (older builds, or a copy with no base) reads as before.
// - Excel's Name Manager takes comments of at most 255 characters. A comment that would
//   pass that with its tag gets no tag, and the build warns.

import { normalizeDefinition, normalizeDefinitionV2 } from "./lockfile.js";
import { sha256 } from "./hash.js";
import type { WorkbookSnapshot } from "../file/types.js";

export interface ProvenanceTag {
  module: string;
  version?: string;
  /** Hex digits of the source hash. */
  hash: string;
  /** Hex digits of the library base (`lib#…`): the library definition the copy came from. */
  lib?: string;
}

/** The Name Manager's limit on a comment (characters; a line break counts as CR LF). */
export const COMMENT_MAX = 255;
const HASH_DIGITS = 6;
const OPEN = "[xln ";
const LIB = "lib#";

function isHex(s: string): boolean {
  if (s.length === 0) return false;
  for (const c of s) if (!((c >= "0" && c <= "9") || (c >= "a" && c <= "f"))) return false;
  return true;
}

/** A token the tag can carry: no blanks, brackets or `#` (they would make it ambiguous). */
function tokenOk(s: string): boolean {
  if (s.length === 0 || s.startsWith("#")) return false;
  for (const c of s) if (c === " " || c === "\t" || c === "\r" || c === "\n" || c === "[" || c === "]") return false;
  return true;
}

export function formatProvenanceTag(t: ProvenanceTag): string {
  return `${OPEN}${t.module}${t.version !== undefined ? " " + t.version : ""} #${t.hash}${t.lib !== undefined ? ` ${LIB}${t.lib}` : ""}]`;
}

/** A comment split into the author's text (undefined when nothing is left) and its trailing tag. */
export function splitProvenance(comment: string): { comment: string | undefined; tag?: ProvenanceTag } {
  const none = { comment: comment === "" ? undefined : comment };
  if (!comment.endsWith("]")) return none;
  const at = comment.lastIndexOf(OPEN);
  if (at < 0) return none;
  const inner = comment.slice(at + OPEN.length, -1);
  if (inner.includes("\n") || inner.includes("\r")) return none;
  const parts = inner.split(" ");
  let lib: string | undefined;
  const tail = parts[parts.length - 1]!;
  if (tail.startsWith(LIB)) {
    if (!isHex(tail.slice(LIB.length))) return none;
    lib = tail.slice(LIB.length);
    parts.pop();
  }
  if (parts.length < 2 || parts.length > 3) return none;
  const last = parts[parts.length - 1]!;
  if (!last.startsWith("#") || !isHex(last.slice(1))) return none;
  if (!parts.slice(0, -1).every(tokenOk)) return none;
  const tag: ProvenanceTag = { module: parts[0]!, hash: last.slice(1) };
  if (parts.length === 3) tag.version = parts[1]!;
  if (lib !== undefined) tag.lib = lib;
  let rest = comment.slice(0, at);
  while (rest.endsWith(" ") || rest.endsWith("\t")) rest = rest.slice(0, -1);
  return { comment: rest === "" ? undefined : rest, tag };
}

/** The comment without its provenance tag; undefined when nothing else is in it. */
export function stripProvenance(comment: string | undefined | null): string | undefined {
  if (comment === undefined || comment === null || comment === "") return undefined;
  return splitProvenance(comment).comment;
}

/** The comment with `tag` at the end of its last line. */
export function withProvenance(comment: string | undefined | null, tag: ProvenanceTag): string {
  const t = formatProvenanceTag(tag);
  return comment === undefined || comment === null || comment === "" ? t : `${comment} ${t}`;
}

/** Length of a comment as Excel counts it (a line break is CR LF). */
export function commentLength(comment: string): number {
  let n = 0;
  for (const c of comment.split("\r\n").join("\n")) n += c === "\n" ? 2 : 1;
  return n;
}

function provenanceHash(normalized: string, comment: string | undefined | null): string {
  const c = stripProvenance(comment === null || comment === undefined ? undefined : comment.split("\r\n").join("\n")) ?? "";
  return sha256(`xln-provenance/1\n${normalized}\n\u0000${c}`).slice(0, HASH_DIGITS);
}

/** Short hash of a name's source: its stored definition (normalised) and its comment without tag. */
export function sourceHash(stored: string, comment: string | undefined | null): string {
  return provenanceHash(normalizeDefinition(stored), comment);
}

/**
 * The hash of tags written before lockfile format 3, which took number literals as
 * written. It differs from `sourceHash` only for a definition with a number not in its
 * canonical spelling (`1E-14`, `1.0`), so the text it hashes keeps its `/1` header: the
 * other tags stay valid.
 */
export function sourceHashV1(stored: string, comment: string | undefined | null): string {
  return provenanceHash(normalizeDefinitionV2(stored), comment);
}

/**
 * The library base of a definition: the first 6 hex digits of SHA-256 over the stored
 * form, normalised as the lockfile hashes it. Comments do not count: a library function's
 * doc comment is generated from its header (and may be shortened), so only definitions
 * decide what is the same version.
 */
export function libraryHash(stored: string): string {
  return sha256(`xln-lib/1\n${normalizeDefinition(stored)}`).slice(0, HASH_DIGITS);
}

/** The annotation that records a library base: `@from(lib #353921)`. */
export function formatLibBase(hash: string): string {
  return `@${FROM}(lib #${hash})`;
}

/** `@from(…)`: the library base of an entry. */
export const FROM = "from";

/** The argument of `@from(…)`: `lib #353921` gives the hash (lower case); otherwise why not. */
export function parseLibBase(arg: string | undefined): { hash: string } | { error: string } {
  const usage = "write @from(lib #abc123), the library version this copy came from (Insert, Take and Publish write it)";
  if (arg === undefined || arg.trim() === "") return { error: `@from needs the library version: ${usage}` };
  const t = arg.trim();
  const rest = t.startsWith("lib") ? t.slice(3).trim() : "";
  if (!rest.startsWith("#")) return { error: `@from(${t}): ${usage}` };
  const h = rest.slice(1).trim().toLowerCase();
  if (h.length !== HASH_DIGITS || !isHex(h)) return { error: `@from(${t}): the version is ${HASH_DIGITS} hex digits after '#'; ${usage}` };
  return { hash: h };
}

/** The module a names file stands for; undefined for `_unmanaged.xln`, sheet files and non-module files. */
export function moduleOfPath(path: string): string | undefined {
  if (!path.startsWith("names/") || !path.endsWith(".xln")) return undefined;
  if (path.startsWith("names/sheets/")) return undefined;
  let base = path.slice(path.lastIndexOf("/") + 1, -4);
  const tilde = base.lastIndexOf("~");
  if (tilde > 0 && /^[0-9]+$/.test(base.slice(tilde + 1))) base = base.slice(0, tilde);
  if (base === "_unmanaged" || !tokenOk(base)) return undefined;
  return base;
}

/** `// @version 1.2` among the comment lines that open a module file; undefined without one. */
export function moduleVersion(text: string): string | undefined {
  for (const raw of text.split("\n")) {
    const line = raw.trim();
    if (line === "") continue;
    if (!line.startsWith("//")) return undefined;
    const body = line.slice(2).trim();
    if (!(body.startsWith("@version ") || body.startsWith("@version\t"))) continue;
    const v = body.slice("@version".length).trim().split(/\s+/)[0] ?? "";
    return tokenOk(v) ? v : undefined;
  }
  return undefined;
}

export interface ProvenanceStatus {
  key: string;
  tag: ProvenanceTag;
  /** `unchanged`: the live name still hashes to the tag; `edited`: changed in Excel since the build that tagged it. */
  state: "unchanged" | "edited";
}

/** D6: the names whose comment carries a provenance tag, and whether they still match it. */
export function provenanceStatus(wb: WorkbookSnapshot): ProvenanceStatus[] {
  const out: ProvenanceStatus[] = [];
  for (const d of wb.definedNames) {
    if (d.isXlPrefixed || d.comment === undefined || d.comment === "") continue;
    const split = splitProvenance(d.comment.split("\r\n").join("\n"));
    if (!split.tag) continue;
    const key = d.scope.kind === "sheet" ? `${d.scope.name}!${d.name}` : d.name;
    // A tag from before numbers counted by their value may hold the older hash.
    const same = sourceHash(d.definition, split.comment) === split.tag.hash || sourceHashV1(d.definition, split.comment) === split.tag.hash;
    out.push({ key, tag: split.tag, state: same ? "unchanged" : "edited" });
  }
  return out;
}

// ---- room for the tag ---------------------------------------------------------------------
// The build drops the tag of a comment that would pass 255 characters with it (tags.ts),
// and could only say so then (feedback 2026-10-07: Publish succeeded, the next build
// warned). The checker, Publish, Insert and Take say it before, measuring the doc comment
// against the tag the build would write, made by the code the build uses.

/**
 * The provenance tag the build writes for a name declared in names file `path` (text
 * `fileText`, for its `@version`), with library base `lib`; undefined when the build writes
 * none (not a module file). `hash` defaults to a stand-in of the hash's length, which is
 * all a measure needs.
 */
export function buildProvenanceTag(path: string, fileText: string, lib?: string, hash = "0".repeat(HASH_DIGITS)): ProvenanceTag | undefined {
  const module = moduleOfPath(path);
  if (module === undefined) return undefined;
  const version = moduleVersion(fileText);
  return { module, ...(version !== undefined ? { version } : {}), hash, ...(lib !== undefined ? { lib } : {}) };
}

/** The characters a tag adds to a doc comment: the blank before it and the tag. */
export function provenanceTagLength(tag: ProvenanceTag): number {
  return commentLength(withProvenance("x", tag)) - 1;
}

/** A doc comment too long to carry its tag. */
export interface DocTagOverflow {
  /** The doc comment's length as Excel counts it. */
  docLength: number;
  /** What the tag adds (`provenanceTagLength`). */
  tagLength: number;
  /** Characters to cut for the tag to fit. */
  over: number;
  /** The warning, without the name in front. */
  message: string;
}

/**
 * Whether doc comment `doc` leaves room for `tag` within the Name Manager's 255 characters:
 * undefined when it does, when there is no tag, or when the doc alone passes 255 (an error
 * of its own, the checker's `comment-length`).
 */
export function docTagOverflow(doc: string | undefined | null, tag: ProvenanceTag | undefined): DocTagOverflow | undefined {
  if (tag === undefined || doc === undefined || doc === null || doc === "") return undefined;
  const docLength = commentLength(doc);
  const tagLength = provenanceTagLength(tag);
  if (docLength > COMMENT_MAX || docLength + tagLength <= COMMENT_MAX) return undefined;
  const over = docLength + tagLength - COMMENT_MAX;
  const breaks = doc.includes("\n") ? " (a line break counts 2)" : "";
  const loses = tag.lib !== undefined ? `its library base, ${formatLibBase(tag.lib)}` : "its module and source hash";
  const message = `the doc comment is ${docLength} characters${breaks}; with its provenance tag (${tagLength}) it passes Excel's ${COMMENT_MAX}, so the build writes it without the tag and the workbook won't record ${loses}: shorten it by ${over} character${over === 1 ? "" : "s"}`;
  return { docLength, tagLength, over, message };
}
