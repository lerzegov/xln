// Base64 in pure TypeScript (the core runs in the browser too, and atob/btoa
// work on binary strings, not bytes). Used by the embedded source (build/embedXml.ts) and
// to read the store of Microsoft's Advanced Formula Environment (file/afe.ts).

const B64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
const INDEX = new Map([...B64].map((c, i) => [c, i]));

export function toBase64(bytes: Uint8Array): string {
  let s = "";
  for (let i = 0; i < bytes.length; i += 3) {
    const a = bytes[i]!;
    const b = bytes[i + 1];
    const c = bytes[i + 2];
    s += B64[a >> 2]! + B64[((a & 3) << 4) | ((b ?? 0) >> 4)]!;
    s += b === undefined ? "=" : B64[((b & 15) << 2) | ((c ?? 0) >> 6)]!;
    s += c === undefined ? "=" : B64[c & 63]!;
  }
  return s;
}

/** The bytes of a base64 text (layout whitespace ignored); undefined when it is not base64. */
export function fromBase64(s: string): Uint8Array | undefined {
  const clean = [...s].filter((c) => c !== " " && c !== "\n" && c !== "\r" && c !== "\t").join("");
  if (clean.length % 4 !== 0) return undefined;
  const out = new Uint8Array((clean.length / 4) * 3);
  let n = 0;
  for (let i = 0; i < clean.length; i += 4) {
    const v = [0, 1, 2, 3].map((k) => (clean[i + k] === "=" ? -2 : (INDEX.get(clean[i + k]!) ?? -1)));
    if (v.some((x) => x === -1)) return undefined;
    out[n++] = (v[0]! << 2) | (v[1]! >> 4);
    if (v[2]! >= 0) out[n++] = ((v[1]! & 15) << 4) | (v[2]! >> 2);
    if (v[3]! >= 0) out[n++] = ((v[2]! & 3) << 6) | v[3]!;
  }
  return out.subarray(0, n);
}
