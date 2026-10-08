// The OPC package: zip entries, text decoding and relationships.
import { strFromU8, unzipSync } from "fflate";
import { childElements, parseXml } from "./xml.js";

export class XlsxError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "XlsxError";
  }
}

export interface Relationship {
  id: string;
  type: string;
  /** Resolved package path (no leading slash); for external targets, the raw target. */
  target: string;
  external: boolean;
}

/** A read-only view of the zip with case-insensitive part lookup (OPC part names are). */
export class Package {
  readonly names: string[] = [];
  private readonly data = new Map<string, Uint8Array>();
  private readonly byLower = new Map<string, string>();

  constructor(bytes: Uint8Array) {
    let files: Record<string, Uint8Array>;
    try {
      // Only XML parts are decompressed; media, vbaProject.bin and the like are listed only.
      files = unzipSync(bytes, {
        filter: (f) => {
          this.names.push(f.name);
          this.byLower.set(f.name.toLowerCase(), f.name);
          const n = f.name.toLowerCase();
          return n.endsWith(".xml") || n.endsWith(".rels");
        },
      });
    } catch (e) {
      throw new XlsxError(`not a readable zip archive: ${(e as Error).message}`);
    }
    for (const [name, d] of Object.entries(files)) this.data.set(name, d);
  }

  /** The stored name of a part, matched case-insensitively. */
  find(path: string): string | undefined {
    return this.byLower.get(path.toLowerCase());
  }

  has(path: string): boolean {
    return this.find(path) !== undefined;
  }

  /** The part's bytes (XML and .rels parts only; others are listed, not decompressed). */
  raw(path: string): Uint8Array | undefined {
    const name = this.find(path);
    return name === undefined ? undefined : this.data.get(name);
  }

  text(path: string): string | undefined {
    const name = this.find(path);
    if (name === undefined) return undefined;
    const d = this.data.get(name);
    return d === undefined ? undefined : decodeText(d);
  }

  /** Relationships of `part` (from `<dir>/_rels/<file>.rels`); empty when there are none. */
  rels(part: string): Relationship[] {
    const slash = part.lastIndexOf("/");
    const dir = slash < 0 ? "" : part.slice(0, slash);
    const file = part.slice(slash + 1);
    const relsPath = (dir ? dir + "/" : "") + "_rels/" + file + ".rels";
    const text = this.text(relsPath);
    if (text === undefined) return [];
    const out: Relationship[] = [];
    for (const r of childElements(parseXml(text), "Relationship")) {
      const id = r.attrs["Id"];
      const type = r.attrs["Type"];
      const target = r.attrs["Target"];
      if (id === undefined || type === undefined || target === undefined) continue;
      const external = r.attrs["TargetMode"] === "External";
      out.push({ id, type, external, target: external ? target : resolveTarget(dir, target) });
    }
    return out;
  }
}

/** Decodes a part's bytes: UTF-8 (with or without BOM) or UTF-16 with a BOM. */
export function decodeText(d: Uint8Array): string {
  if (d.length >= 2 && ((d[0] === 0xff && d[1] === 0xfe) || (d[0] === 0xfe && d[1] === 0xff))) {
    const le = d[0] === 0xff;
    let s = "";
    for (let i = 2; i + 1 < d.length; i += 2) {
      s += String.fromCharCode(le ? d[i]! | (d[i + 1]! << 8) : (d[i]! << 8) | d[i + 1]!);
    }
    return s;
  }
  const s = strFromU8(d);
  return s.charCodeAt(0) === 0xfeff ? s.slice(1) : s;
}

/**
 * Resolves a relationship target against the source part's directory. Absolute targets
 * (`/xl/worksheets/sheet1.xml`) are package-rooted; `..` and `.` segments are folded;
 * percent-escapes are decoded (OPC part names are URIs).
 */
export function resolveTarget(sourceDir: string, target: string): string {
  let t = target;
  try {
    t = decodeURIComponent(t);
  } catch {
    // Not valid percent-encoding: use as written.
  }
  const segs = t.startsWith("/") ? [] : sourceDir.split("/").filter((s) => s !== "");
  for (const s of t.split("/")) {
    if (s === "" || s === ".") continue;
    if (s === "..") segs.pop();
    else segs.push(s);
  }
  return segs.join("/");
}

/** True if a relationship type URI ends with `/<suffix>` (covers transitional and strict). */
export function relTypeIs(type: string, suffix: string): boolean {
  return type.endsWith("/" + suffix);
}
