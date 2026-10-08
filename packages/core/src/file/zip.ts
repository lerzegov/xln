// Replacing, removing and adding entries of a zip archive without touching the others
// (requirement D2: every part other than the patched ones is copied byte for byte). fflate's zipSync would
// re-compress every entry, so the archive is rewritten here at record level: each
// untouched entry's local record (header, compressed data, data descriptor) is copied as
// it was, and only the central directory's offsets are updated.
//
// Scope: the archives Excel writes. ZIP64, encryption and multi-disk archives are refused.

import { deflateSync, strFromU8, strToU8 } from "fflate";
import { XlsxError } from "./package.js";

const SIG_LOCAL = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const FLAG_DESCRIPTOR = 0x0008;
const FLAG_UTF8 = 0x0800;
const FLAG_ENCRYPTED = 0x0001;

export interface ZipEntry {
  /** As fflate reads it: UTF-8 when the flag says so, else Latin-1. */
  name: string;
  /** The name's bytes as stored. */
  nameBytes: Uint8Array;
  method: number;
  flags: number;
  crc: number;
  compressedSize: number;
  size: number;
  /** Offset of the local header. */
  localOffset: number;
  /** The central directory record, as stored. */
  central: Uint8Array;
}

export interface ZipLayout {
  /** In central-directory order. */
  entries: ZipEntry[];
  centralStart: number;
  eocdOffset: number;
  /** The end-of-central-directory record with its comment, as stored. */
  eocd: Uint8Array;
}

function u16(b: Uint8Array, o: number): number {
  return b[o]! | (b[o + 1]! << 8);
}

function u32(b: Uint8Array, o: number): number {
  return (b[o]! | (b[o + 1]! << 8) | (b[o + 2]! << 16) | (b[o + 3]! << 24)) >>> 0;
}

function put16(b: Uint8Array, o: number, v: number): void {
  b[o] = v & 0xff;
  b[o + 1] = (v >>> 8) & 0xff;
}

function put32(b: Uint8Array, o: number, v: number): void {
  b[o] = v & 0xff;
  b[o + 1] = (v >>> 8) & 0xff;
  b[o + 2] = (v >>> 16) & 0xff;
  b[o + 3] = (v >>> 24) & 0xff;
}

/** A copy of `b[start, end)`. Not `slice`: Node's byte arrays (what VS Code's file API returns on
 *  the desktop) answers `slice` with a view, and patching that view would corrupt the input. */
function copy(b: Uint8Array, start: number, end: number): Uint8Array {
  return new Uint8Array(b.subarray(start, end));
}

function fail(message: string): never {
  throw new XlsxError(`zip: ${message}`);
}

/** Reads the central directory. */
export function readZipLayout(bytes: Uint8Array): ZipLayout {
  // The EOCD record is the last thing in the file, followed by a comment of at most 64 KiB.
  let eocd = -1;
  for (let o = bytes.length - 22; o >= Math.max(0, bytes.length - 22 - 0xffff); o--) {
    if (u32(bytes, o) === SIG_EOCD && o + 22 + u16(bytes, o + 20) === bytes.length) {
      eocd = o;
      break;
    }
  }
  if (eocd < 0) fail("no end-of-central-directory record");
  if (u16(bytes, eocd + 4) !== 0 || u16(bytes, eocd + 6) !== 0) fail("multi-disk archives are not supported");
  const count = u16(bytes, eocd + 10);
  const cdSize = u32(bytes, eocd + 12);
  const cdStart = u32(bytes, eocd + 16);
  if (count === 0xffff || cdSize === 0xffffffff || cdStart === 0xffffffff) fail("ZIP64 archives are not supported");
  if (cdStart + cdSize > eocd) fail("central directory overlaps its end record");
  const entries: ZipEntry[] = [];
  let o = cdStart;
  for (let k = 0; k < count; k++) {
    if (u32(bytes, o) !== SIG_CENTRAL) fail(`bad central directory record ${k}`);
    const flags = u16(bytes, o + 8);
    const nameLen = u16(bytes, o + 28);
    const extraLen = u16(bytes, o + 30);
    const commentLen = u16(bytes, o + 32);
    const compressedSize = u32(bytes, o + 20);
    const size = u32(bytes, o + 24);
    const localOffset = u32(bytes, o + 42);
    if (compressedSize === 0xffffffff || size === 0xffffffff || localOffset === 0xffffffff) fail("ZIP64 entries are not supported");
    if (flags & FLAG_ENCRYPTED) fail("encrypted entries are not supported");
    const nameBytes = copy(bytes, o + 46, o + 46 + nameLen);
    const end = o + 46 + nameLen + extraLen + commentLen;
    entries.push({
      name: strFromU8(nameBytes, !(flags & FLAG_UTF8)),
      nameBytes,
      method: u16(bytes, o + 10),
      flags,
      crc: u32(bytes, o + 16),
      compressedSize,
      size,
      localOffset,
      central: copy(bytes, o, end),
    });
    o = end;
  }
  for (const e of entries) if (u32(bytes, e.localOffset) !== SIG_LOCAL) fail(`entry ${e.name}: no local header at ${e.localOffset}`);
  return { entries, centralStart: cdStart, eocdOffset: eocd, eocd: copy(bytes, eocd, bytes.length) };
}

/**
 * Each entry's local record exactly as stored: header, data and data descriptor. A record
 * runs to the next record (by offset) or to the central directory.
 */
export function rawZipRecords(bytes: Uint8Array, layout: ZipLayout = readZipLayout(bytes)): Map<string, Uint8Array> {
  const byOffset = [...layout.entries].sort((a, b) => a.localOffset - b.localOffset);
  const out = new Map<string, Uint8Array>();
  byOffset.forEach((e, k) => {
    const end = k + 1 < byOffset.length ? byOffset[k + 1]!.localOffset : layout.centralStart;
    out.set(e.name, bytes.subarray(e.localOffset, end));
  });
  return out;
}

let crcTable: Uint32Array | undefined;

export function crc32(data: Uint8Array): number {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c >>> 0;
    }
  }
  let c = 0xffffffff;
  for (let i = 0; i < data.length; i++) c = crcTable[(c ^ data[i]!) & 0xff]! ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/**
 * A copy of the archive in which the named entries have new (uncompressed) contents.
 * Every other entry's record is copied byte for byte and keeps its place; so does the
 * archive comment. Replacing an entry that does not exist is an error.
 */
export function replaceZipEntries(bytes: Uint8Array, replacements: ReadonlyMap<string, Uint8Array>): Uint8Array {
  return rewriteZip(bytes, { replace: replacements });
}

export interface ZipEdits {
  /** Entries given new contents, in place. */
  replace?: ReadonlyMap<string, Uint8Array>;
  /** Entries removed (a cell build drops `xl/calcChain.xml`). */
  remove?: Iterable<string>;
  /** New entries, appended after the others in this order. */
  add?: ReadonlyMap<string, Uint8Array>;
}

/** A deflated local record and its central-directory record, for new or replaced contents. */
function packEntry(
  data: Uint8Array,
  nameBytes: Uint8Array,
  at: number,
  base: { flags: number; versionNeeded: number; time: number; date: number; central?: Uint8Array },
): { record: Uint8Array; central: Uint8Array } {
  const packed = deflateSync(data, { level: 6 });
  const crc = crc32(data);
  const flags = base.flags & ~FLAG_DESCRIPTOR;
  const record = new Uint8Array(30 + nameBytes.length + packed.length);
  put32(record, 0, SIG_LOCAL);
  put16(record, 4, Math.max(20, base.versionNeeded));
  put16(record, 6, flags);
  put16(record, 8, 8);
  put16(record, 10, base.time);
  put16(record, 12, base.date);
  put32(record, 14, crc);
  put32(record, 18, packed.length);
  put32(record, 22, data.length);
  put16(record, 26, nameBytes.length);
  put16(record, 28, 0);
  record.set(nameBytes, 30);
  record.set(packed, 30 + nameBytes.length);
  let c: Uint8Array;
  if (base.central) {
    c = copy(base.central, 0, base.central.length);
    put16(c, 6, Math.max(20, u16(c, 6)));
  } else {
    c = new Uint8Array(46 + nameBytes.length);
    put32(c, 0, SIG_CENTRAL);
    put16(c, 4, 20); // made by: MS-DOS attributes, spec 2.0
    put16(c, 6, 20);
    put16(c, 12, base.time);
    put16(c, 14, base.date);
    put16(c, 28, nameBytes.length);
    c.set(nameBytes, 46);
  }
  put16(c, 8, flags);
  put16(c, 10, 8);
  put32(c, 16, crc);
  put32(c, 20, packed.length);
  put32(c, 24, data.length);
  put32(c, 42, at);
  return { record, central: c };
}

function isAscii(s: string): boolean {
  for (let i = 0; i < s.length; i++) if (s.charCodeAt(i) > 0x7e || s.charCodeAt(i) < 0x20) return false;
  return true;
}

/**
 * A copy of the archive with entries replaced, removed and added. Every entry not named
 * is copied byte for byte (local record and central record) and keeps its relative
 * order; so does the archive comment. Replacing or removing an entry that does not exist,
 * or adding one that does, is an error.
 */
export function rewriteZip(bytes: Uint8Array, edits: ZipEdits): Uint8Array {
  const layout = readZipLayout(bytes);
  const raw = rawZipRecords(bytes, layout);
  const replace = edits.replace ?? new Map<string, Uint8Array>();
  const remove = new Set(edits.remove ?? []);
  const add = edits.add ?? new Map<string, Uint8Array>();
  const exists = (name: string) => layout.entries.some((e) => e.name === name);
  for (const name of replace.keys()) if (!exists(name)) fail(`no entry ${name} to replace`);
  for (const name of remove) if (!exists(name)) fail(`no entry ${name} to remove`);
  for (const name of add.keys()) if (exists(name)) fail(`entry ${name} exists already`);
  for (const name of replace.keys()) if (remove.has(name)) fail(`entry ${name} both replaced and removed`);

  const byOffset = [...layout.entries].sort((a, b) => a.localOffset - b.localOffset).filter((e) => !remove.has(e.name));
  const records: Uint8Array[] = [];
  const newCentral = new Map<ZipEntry, Uint8Array>();
  let at = 0;
  for (const e of byOffset) {
    const data = replace.get(e.name);
    let record: Uint8Array;
    if (data === undefined) {
      record = raw.get(e.name)!;
      const c = copy(e.central, 0, e.central.length);
      put32(c, 42, at);
      newCentral.set(e, c);
    } else {
      const local = bytes.subarray(e.localOffset, e.localOffset + 30);
      // Time and date as they were.
      const packed = packEntry(data, e.nameBytes, at, { flags: e.flags, versionNeeded: u16(local, 4), time: u16(local, 10), date: u16(local, 12), central: e.central });
      record = packed.record;
      newCentral.set(e, packed.central);
    }
    records.push(record);
    at += record.length;
  }
  const central = layout.entries.filter((e) => !remove.has(e.name)).map((e) => newCentral.get(e)!);
  // New entries take the time stamp of the first entry, so a build is reproducible.
  const first = layout.entries[0];
  const stamp = first ? { time: u16(first.central, 12), date: u16(first.central, 14) } : { time: 0, date: 0x21 };
  for (const [name, data] of add) {
    const utf = !isAscii(name);
    const packed = packEntry(data, strToU8(name), at, { flags: utf ? FLAG_UTF8 : 0, versionNeeded: 20, ...stamp });
    records.push(packed.record);
    central.push(packed.central);
    at += packed.record.length;
  }
  if (central.length > 0xfffe) fail("too many entries");
  const centralStart = at;
  const cdSize = central.reduce((n, c) => n + c.length, 0);
  const eocd = copy(layout.eocd, 0, layout.eocd.length);
  put16(eocd, 8, central.length);
  put16(eocd, 10, central.length);
  put32(eocd, 12, cdSize);
  put32(eocd, 16, centralStart);

  const out = new Uint8Array(centralStart + cdSize + eocd.length);
  let o = 0;
  for (const r of records) {
    out.set(r, o);
    o += r.length;
  }
  for (const c of central) {
    out.set(c, o);
    o += c.length;
  }
  out.set(eocd, o);
  return out;
}
