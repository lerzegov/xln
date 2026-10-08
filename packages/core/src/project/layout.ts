// Line breaks in a formula's layout. Excel stores them as CR LF in definitions;
// the project's files use LF. Strings keep theirs: that is their value.

import { tokenize } from "../lang/tokens.js";

/** Line breaks in layout whitespace → LF. Strings keep theirs: that is their value. */
export function layoutToLf(src: string): string {
  if (!src.includes("\r")) return src;
  return tokenize(src)
    .map((t) => (t.kind === "ws" || t.kind === "isect" ? t.text.split("\r\n").join("\n") : t.text))
    .join("");
}
