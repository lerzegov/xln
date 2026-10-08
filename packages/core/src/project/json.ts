// Stable, diff-friendly JSON: two-space indentation, but arrays of plain values on one
// line (a list of 200 cell ranges should not take 200 lines), and so are small objects of
// plain values. Object keys keep the order
// the caller built them in, which the callers make deterministic.

export type Json = string | number | boolean | null | Json[] | { [k: string]: Json | undefined };

function flat(v: Json): boolean {
  return v === null || typeof v !== "object";
}

export function stringifyJson(value: Json, indent = ""): string {
  if (flat(value)) return JSON.stringify(value);
  const next = indent + "  ";
  if (Array.isArray(value)) {
    if (value.length === 0) return "[]";
    if (value.every(flat)) return "[" + value.map((v) => JSON.stringify(v)).join(", ") + "]";
    return "[\n" + value.map((v) => next + stringifyJson(v, next)).join(",\n") + "\n" + indent + "]";
  }
  const entries = Object.entries(value as { [k: string]: Json | undefined }).filter(([, v]) => v !== undefined) as [string, Json][];
  if (entries.length === 0) return "{}";
  if (entries.every(([, v]) => flat(v) || (Array.isArray(v) && v.length === 0))) {
    const one = "{ " + entries.map(([k, v]) => JSON.stringify(k) + ": " + JSON.stringify(v)).join(", ") + " }";
    if (indent.length + one.length <= 100) return one;
  }
  return "{\n" + entries.map(([k, v]) => next + JSON.stringify(k) + ": " + stringifyJson(v, next)).join(",\n") + "\n" + indent + "}";
}
