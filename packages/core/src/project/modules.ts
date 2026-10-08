// A4: propose modules from the names' spelling. The rule is deliberately simple and
// deterministic, so a re-pull of the same workbook gives the same files:
//
// 1. Dot prefix. A name with a dot belongs to the module named by the text before its
//    first dot: `FN.PICK` → `FN`, `Mod.Sub.f` → `Mod`. One such name is enough.
// 2. Upper-case underscore prefix. `IN_Rate` → `IN` when the prefix is 1–8 upper-case
//    ASCII letters or digits starting with a letter, the text after `_` starts with an
//    upper-case letter, and at least three names (any scope) share the prefix. The
//    second condition keeps suffix families such as `BS_base`, `BS_payout` out: there the
//    upper-case part is the subject, not a namespace.
// 3. Everything else has no module. Its workbook-scoped names go to `_unmanaged.xln`,
//    its sheet-scoped names to `sheets/<Sheet>.xln`, one file per sheet (`sheetFileName`).
//
// Module names compare case-insensitively (Excel names do, and so do the default file
// systems on macOS and Windows); the spelling used is the most common one, ties broken
// by sort order.

import { strFromU8 } from "fflate";
import { utf8 } from "./hash.js";

export const UNMANAGED = "_unmanaged";

/** Folder (inside `names/`) of the per-sheet files. */
export const SHEETS_DIR = "sheets";

/** Minimum number of names sharing an underscore prefix before it counts as a module. */
export const UNDERSCORE_MIN = 3;

function isUpperAscii(c: string | undefined): boolean {
  return c !== undefined && c >= "A" && c <= "Z";
}
function isDigit(c: string | undefined): boolean {
  return c !== undefined && c >= "0" && c <= "9";
}

function dotPrefix(name: string): string | undefined {
  const i = name.indexOf(".");
  return i > 0 ? name.slice(0, i) : undefined;
}

function underscorePrefix(name: string): string | undefined {
  const i = name.indexOf("_");
  if (i < 1 || i > 8) return undefined;
  const p = name.slice(0, i);
  if (!isUpperAscii(p[0])) return undefined;
  for (const c of p) if (!isUpperAscii(c) && !isDigit(c)) return undefined;
  return isUpperAscii(name[i + 1]) ? p : undefined;
}

function byCodePoint(a: string, b: string): number {
  return a < b ? -1 : a > b ? 1 : 0;
}

/** Sort key used everywhere in the project: case-insensitive, then exact spelling. */
export function compareNames(a: string, b: string): number {
  return byCodePoint(a.toLowerCase(), b.toLowerCase()) || byCodePoint(a, b);
}

/**
 * The module of each name (undefined → `_unmanaged`), for a list of distinct names.
 * Same input, same answer: the result does not depend on input order.
 */
export function proposeModules(names: readonly string[]): Map<string, string | undefined> {
  const sorted = [...new Set(names)].sort(compareNames);
  const counts = new Map<string, number>();
  const seen = new Set<string>();
  for (const n of sorted) {
    const k = n.toLowerCase();
    if (seen.has(k)) continue;
    seen.add(k);
    const p = dotPrefix(n) === undefined ? underscorePrefix(n) : undefined;
    if (p) counts.set(p, (counts.get(p) ?? 0) + 1);
  }
  const raw = new Map<string, string | undefined>();
  // lower-case module → spelling → number of names using it
  const spellings = new Map<string, Map<string, number>>();
  for (const n of sorted) {
    let m = dotPrefix(n);
    if (m === undefined) {
      const p = underscorePrefix(n);
      if (p && (counts.get(p) ?? 0) >= UNDERSCORE_MIN) m = p;
    }
    if (m !== undefined && m.toLowerCase() === UNMANAGED) m = undefined;
    raw.set(n, m);
    if (m !== undefined) {
      const k = m.toLowerCase();
      const s = spellings.get(k) ?? new Map<string, number>();
      s.set(m, (s.get(m) ?? 0) + 1);
      spellings.set(k, s);
    }
  }
  const out = new Map<string, string | undefined>();
  for (const [n, m] of raw) {
    if (m === undefined) {
      out.set(n, undefined);
      continue;
    }
    const s = [...spellings.get(m.toLowerCase())!].sort((a, b) => b[1] - a[1] || compareNames(a[0], b[0]));
    out.set(n, s[0]![0]);
  }
  return out;
}

const UNSAFE_FILE_CHARS = new Set(["\\", "/", ":", "*", "?", '"', "<", ">", "|"]);

/** File name (without folder) for a module: characters Windows forbids become `_`. */
export function moduleFileName(module: string | undefined): string {
  if (module === undefined) return UNMANAGED + ".xln";
  let s = "";
  for (const c of module) s += UNSAFE_FILE_CHARS.has(c) || c < " " ? "_" : c;
  return s + ".xln";
}

// Sheet files (A4): `names/sheets/<Sheet>.xln`. The file name is only a place to look;
// the scope is the `@scope(...)` inside the file. Excel already forbids `\ / ? * [ ] :`
// in sheet names but allows `" < > |`, `%`, a leading `.` and the Windows device names,
// so the mapping percent-encodes (UTF-8, upper-case hex):
//   - the characters Windows forbids in file names, control characters, and `%` itself
//     (which keeps the mapping reversible);
//   - a leading `.` (a hidden file on macOS and Linux) and a trailing `.` or space
//     (Windows drops them);
//   - the first letter of a Windows device name (`CON`, `PRN`, `AUX`, `NUL`, `COM1`–`COM9`,
//     `LPT1`–`LPT9`), reserved with any extension: `CON` → `%43ON.xln`.
// Spaces and non-ASCII letters stay: `SCF recursive` → `SCF recursive.xln`.
// Excel keeps sheet names unique ignoring case, so the mapping alone cannot make two files
// that differ only in case; `pullProject` still compares `fileSystemKey` (case and
// Unicode normalisation, which APFS ignores and Excel does not) and adds `~2` on a clash.

const DEVICE_NAMES = new Set(["CON", "PRN", "AUX", "NUL", ...[1, 2, 3, 4, 5, 6, 7, 8, 9].flatMap((d) => [`COM${d}`, `LPT${d}`])]);

function percent(c: string): string {
  let s = "";
  for (const b of utf8(c)) s += "%" + b.toString(16).toUpperCase().padStart(2, "0");
  return s;
}

/** File name (without folder) for a sheet's own names; `sheetFromFileName` reverses it. */
export function sheetFileName(sheet: string): string {
  const chars = [...sheet];
  const out = chars.map((c) => (UNSAFE_FILE_CHARS.has(c) || c < " " || c === "\x7f" || c === "%" ? percent(c) : c));
  if (chars[0] === "." || DEVICE_NAMES.has(sheet.toUpperCase())) out[0] = percent(chars[0]!);
  const last = chars.length - 1;
  if (last >= 0 && (chars[last] === "." || chars[last] === " ")) out[last] = percent(chars[last]!);
  return out.join("") + ".xln";
}

function hexValue(c: string | undefined): number {
  return c === undefined ? -1 : "0123456789abcdef".indexOf(c.toLowerCase());
}

/** The sheet a `sheetFileName` stands for; undefined when the text is not one. */
export function sheetFromFileName(file: string): string | undefined {
  if (!file.endsWith(".xln")) return undefined;
  const s = file.slice(0, -4);
  const parts: string[] = [];
  let bytes: number[] = [];
  const flush = (): void => {
    if (bytes.length) parts.push(strFromU8(new Uint8Array(bytes)));
    bytes = [];
  };
  for (let i = 0; i < s.length; i++) {
    if (s[i] === "%") {
      const hi = hexValue(s[i + 1]);
      const lo = hexValue(s[i + 2]);
      if (hi < 0 || lo < 0) return undefined;
      bytes.push(hi * 16 + lo);
      i += 2;
    } else {
      flush();
      parts.push(s[i]!);
    }
  }
  flush();
  const sheet = parts.join("");
  // Invalid UTF-8 decodes to U+FFFD and needless escapes decode too: only an exact round
  // trip (hex digits in either case) counts.
  return sheetFileName(sheet).toLowerCase() === file.toLowerCase() ? sheet : undefined;
}

/**
 * Two file names that the default macOS and Windows file systems treat as the same file
 * get the same key: case-insensitive, and insensitive to Unicode normalisation (APFS).
 */
export function fileSystemKey(file: string): string {
  return file.normalize("NFC").toLowerCase();
}
