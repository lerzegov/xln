// ST_Xstring escapes (ECMA-376 Part 1, 22.9.2.19): in some attribute values Excel writes a
// character as `_xHHHH_`, and a literal `_xHHHH_` as `_x005F_xHHHH_`. Measured on a name's
// comment (probes/README.md § F5): a line break written by xln as `&#10;` comes back from
// Excel's save as `_x000a_`. A reader that did not decode it would see the comment changed.

function isHex4(s: string, at: number): boolean {
  for (let k = at; k < at + 4; k++) {
    const c = s[k];
    if (c === undefined || !((c >= "0" && c <= "9") || (c >= "a" && c <= "f") || (c >= "A" && c <= "F"))) return false;
  }
  return true;
}

/** Whether `s` has an escape at `i`: `_x` + four hex digits + `_`. */
function escapeAt(s: string, i: number): boolean {
  return s[i] === "_" && s[i + 1] === "x" && isHex4(s, i + 2) && s[i + 6] === "_";
}

/** Decodes `_xHHHH_` escapes. */
export function decodeXstring(s: string): string {
  if (!s.includes("_x")) return s;
  let out = "";
  for (let i = 0; i < s.length; i++) {
    if (escapeAt(s, i)) {
      out += String.fromCharCode(parseInt(s.slice(i + 2, i + 6), 16));
      i += 6;
    } else out += s[i];
  }
  return out;
}

/** Escapes text that would read as an escape (`_x0041_` → `_x005F_x0041_`); everything else as it is. */
export function encodeXstring(s: string): string {
  if (!s.includes("_x")) return s;
  let out = "";
  for (let i = 0; i < s.length; i++) out += escapeAt(s, i) ? "_x005F_" : s[i];
  return out;
}
