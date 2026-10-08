// C14: the copy of names that Microsoft's Advanced Formula Environment (AFE, Excel Labs)
// keeps in the workbook (file/afe.ts), set against the names themselves. xln writes the
// Name Manager's names; AFE keeps module text in its own part and shows that text when it
// opens. When the two disagree, AFE shows the old text and, when its modules are saved
// from AFE, may write that text back over the names. xln never edits AFE's part (no
// hidden transformations): it says where the two differ, and the user decides.
//
// A module's entry `X` is the name `<Module>.X` (`X` for the Workbook module). Inside a
// module AFE lets a sibling be named without its module (`CUBE` for `ANA.CUBE`), so such
// names are qualified before comparing. Comparison is modulo whitespace, case of names,
// and the spelling of numbers, on the display forms (prefixes and `_xlpm.` decompiled).

import type { DefinedName, ForeignModuleStore, WorkbookSnapshot } from "../file/types.js";
import { afeExportedName } from "../file/afe.js";
import { canonicalNumber } from "../lang/format.js";
import { tokenize } from "../lang/tokens.js";
import { decompile } from "../lang/transform.js";
import { parseModule } from "../project/module.js";

export type AfeEntryState =
  /** Same definition (modulo whitespace) in AFE's module and in the workbook. */
  | "same"
  /** Both have the name; the definitions differ. */
  | "differs"
  /** AFE's module defines it; the workbook has no such workbook-scoped name. */
  | "absent"
  /** Not compared: the module text is written with other separators than the file's. */
  | "not-compared";

export interface AfeEntryStatus {
  module: string;
  /** The entry's name in the module (`CUBE`). */
  entry: string;
  /** Its name in the Name Manager (`ANA.CUBE`); the workbook's spelling when it has the name. */
  name: string;
  state: AfeEntryState;
  /** AFE's definition, as its module writes it (comments removed, trimmed). */
  afe: string;
  /** The workbook's definition in display form, when it has the name. */
  workbook?: string;
}

export interface AfeStoreStatus {
  store: ForeignModuleStore;
  /** Per module, in store order: the names its text defines. */
  modules: { name: string; names: number }[];
  entries: AfeEntryStatus[];
  /** Why entries were not compared (`not-compared`). */
  notCompared?: string;
  /** Entries found in the workbook through AFE's list of exported names rather than as `<Module>.<name>`. */
  viaExportedNames?: number;
}

/** Comparable tokens of a formula. `local` (AFE's side): a module's entries, lower case, → their names in the Name Manager. */
function keys(src: string, local?: ReadonlyMap<string, string>): string[] {
  const out: string[] = [];
  for (const t of tokenize(src.split("\r\n").join("\n"))) {
    if (t.kind === "ws" || t.kind === "eof") continue;
    if (t.kind === "number") out.push(`n:${canonicalNumber(t.text)}`);
    else if (t.kind === "name") {
      const v = (t.value ?? t.text).toLowerCase();
      const q = t.qual ? `${(t.qual.sheet ?? "").toLowerCase()}!` : "";
      out.push(`name:${q}${!t.qual && local ? (local.get(v) ?? v) : v}`);
    } else if (t.kind === "ref") out.push(`ref:${(t.qual?.sheet ?? "").toLowerCase()}!${(t.value ?? t.text).toUpperCase()}`);
    else if (t.kind === "error" || t.kind === "bool") out.push(`${t.kind}:${t.text.toUpperCase()}`); // Excel stores #VALUE!, TRUE
    else out.push(`${t.kind}:${t.text}`);
  }
  return out;
}

function same(a: string[], b: string[]): boolean {
  return a.length === b.length && a.every((k, i) => k === b[i]);
}

function display(d: DefinedName): string {
  try {
    return decompile(d.definition);
  } catch {
    return d.definition;
  }
}

/** Every readable AFE store of the workbook, its entries set against the workbook's names. */
export function afeStatus(wb: WorkbookSnapshot): AfeStoreStatus[] {
  const byName = new Map<string, DefinedName>();
  for (const d of wb.definedNames) if (d.scope.kind === "workbook" && !byName.has(d.name.toLowerCase())) byName.set(d.name.toLowerCase(), d);
  const out: AfeStoreStatus[] = [];
  for (const store of wb.foreignModuleStores) {
    if (store.tool !== "afe" || store.kind !== "custom-xml" || !store.modules) continue;
    const sep = store.locale?.listSeparator;
    const dec = store.locale?.decimalSeparator;
    const notCompared = (sep !== undefined && sep !== ",") || (dec !== undefined && dec !== ".") ? `AFE's modules are written with "${sep ?? ","}" between arguments and "${dec ?? "."}" for decimals (locale ${store.locale?.localeName ?? "unknown"}); the file stores "," and "."` : undefined;
    const status: AfeStoreStatus = { store, modules: [], entries: [] };
    const lastSegment = new Map<string, string[]>();
    for (const n of store.exportedNames ?? []) {
      const k = n.slice(n.lastIndexOf(".") + 1).toLowerCase();
      lastSegment.set(k, [...(lastSegment.get(k) ?? []), n]);
    }
    if (notCompared) status.notCompared = notCompared;
    for (const m of store.modules) {
      const parsed = parseModule(m.text);
      status.modules.push({ name: m.name, names: parsed.entries.length });
      // Each entry's name in the Name Manager first: inside the module a sibling may be named without its module.
      const exportedOf = new Map<string, string>();
      for (const e of parsed.entries) {
        let exported = afeExportedName(m.name, e.name);
        if (!byName.has(exported.toLowerCase())) {
          // A store whose names were exported under another prefix (seen in a third-party
          // AFE-shaped file: module Dates, names oz.*): AFE's own list of the names it
          // exported says which, when exactly one of them ends in this entry's name.
          const hits = (lastSegment.get(e.name.toLowerCase()) ?? []).filter((n) => byName.has(n.toLowerCase()));
          if (hits.length === 1) {
            exported = hits[0]!;
            status.viaExportedNames = (status.viaExportedNames ?? 0) + 1;
          }
        }
        exportedOf.set(e.name.toLowerCase(), exported.toLowerCase());
      }
      for (const e of parsed.entries) {
        const d = byName.get(exportedOf.get(e.name.toLowerCase())!);
        const entry: AfeEntryStatus = { module: m.name, entry: e.name, name: d?.name ?? afeExportedName(m.name, e.name), state: "absent", afe: e.formula };
        if (d) {
          entry.workbook = display(d);
          entry.state = notCompared ? "not-compared" : same(keys(e.formula, exportedOf), keys(entry.workbook)) ? "same" : "differs";
        }
        status.entries.push(entry);
      }
    }
    out.push(status);
  }
  return out;
}

/** What the store is, for messages: `customXml/item1.xml (AFE 1.1+ module store)`. */
export function afeStoreLabel(s: ForeignModuleStore): string {
  if (s.kind === "custom-xml") return `${s.part} (AFE's module store)`;
  if (s.kind === "code-sheet") return `very hidden sheet ${s.sheet} (AFE 1.0's module store)`;
  return `${s.state === "visible" ? "" : `${s.state === "veryHidden" ? "very " : ""}hidden `}sheet ${s.sheet} (AFE's locale detection)`;
}

/** The names a change set touches that AFE's text defines too: the names AFE will disagree with after the build. */
export function afeNamesTouched(wb: WorkbookSnapshot, touched: Iterable<string>): string[] {
  const afe = new Map<string, string>();
  for (const s of afeStatus(wb)) for (const e of s.entries) afe.set(e.name.toLowerCase(), e.name);
  for (const s of wb.foreignModuleStores) for (const n of s.exportedNames ?? []) if (!afe.has(n.toLowerCase())) afe.set(n.toLowerCase(), n);
  const out: string[] = [];
  for (const t of touched) {
    const hit = afe.get(t.toLowerCase());
    if (hit !== undefined && !out.includes(hit)) out.push(hit);
  }
  return out;
}
