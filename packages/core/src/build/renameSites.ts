// Where the workbook reads a name that a build renames (stretch G, M5), sorted into what
// the build rewrites (cell formulas, conditional formats, validations, other names'
// definitions: `applyReferenceRenames`) and what it cannot reach, which refuses the
// rename with the list of places (PLAN.md M5): charts, Table column formulas, pivot
// tables' sources, hyperlinks, form controls, a link of the workbook to itself. Also a
// formula the rewrite would make read another name (`captured`), or that does not parse
// and may hold the name.
//
// Renaming in Excel's Name Manager rewrites all of those; the refusal says so.

import { Package } from "../file/package.js";
import type { WorkbookSnapshot } from "../file/types.js";
import { XmlReader } from "../file/xml.js";
import { tokenize } from "../lang/tokens.js";
import { RenameContext, renameInFormula } from "../project/rename.js";
import { collectSites } from "../audit/sites.js";
import { sheetPartsOf, workbookPartOf } from "./apply.js";
import type { RenameReferences } from "./changes.js";

export interface RenameSites {
  /** Formulas the build rewrites, by kind. */
  counts: RenameReferences;
  /** Places the build cannot rewrite that read the name. */
  blocked: string[];
  /** Formulas whose rewrite would read another name, with the name. */
  captured: string[];
  /** Formulas that do not parse and spell the old name. */
  unparsed: string[];
}

/** A formula-like text elsewhere in the package, with the sheet it is read on. */
interface OtherReader {
  place: string;
  text: string;
  home: string | undefined;
}

const lower = (s: string) => s.toLowerCase();

function attrsOf(xml: string, want: (local: string) => boolean, each: (local: string, attrs: Record<string, string>, text: string) => void): void {
  const r = new XmlReader(xml);
  const stack: { local: string; attrs: Record<string, string>; text: string }[] = [];
  for (let t = r.next(); t; t = r.next()) {
    if (t.type === "open") stack.push({ local: t.local, attrs: t.attrs, text: "" });
    else if (t.type === "text") {
      const top = stack[stack.length - 1];
      if (top) top.text += t.text;
    } else {
      const e = stack.pop();
      if (e && want(e.local)) each(e.local, e.attrs, e.text);
    }
  }
}

function baseName(p: string): string {
  const i = Math.max(p.lastIndexOf("/"), p.lastIndexOf("\\"));
  return i < 0 ? p : p.slice(i + 1);
}

/**
 * Texts outside the formulas the snapshot reads that may name a defined name: hyperlink
 * locations, form controls' links (`fmlaLink`, `fmlaRange`, …), shapes' text links,
 * pivot caches' source names, and the names a link of the workbook to itself reads.
 */
export function otherNameReaders(bytes: Uint8Array, fileName: string | undefined): OtherReader[] {
  const out: OtherReader[] = [];
  let pkg: Package;
  try {
    pkg = new Package(bytes);
  } catch {
    return out;
  }
  const wbPart = workbookPartOf(pkg);
  for (const s of sheetPartsOf(pkg, wbPart).values()) {
    const xml = s.part ? pkg.text(s.part) : undefined;
    if (!xml) continue;
    try {
      attrsOf(xml, (l) => l === "hyperlink", (_l, a) => {
        if (a["location"]) out.push({ place: `hyperlink on ${s.name}${a["ref"] ? `!${a["ref"]}` : ""}`, text: a["location"], home: s.name });
      });
    } catch {
      // An unreadable sheet fails elsewhere.
    }
  }
  for (const path of pkg.names) {
    const lp = lower(path);
    const xml = lp.endsWith(".xml") ? pkg.text(path) : undefined;
    if (!xml) continue;
    try {
      if (lp.startsWith("xl/ctrlprops/")) {
        attrsOf(xml, (l) => l === "formControlPr", (_l, a) => {
          for (const [k, v] of Object.entries(a)) if (k.startsWith("fmla") && v !== "") out.push({ place: `form control ${path} (${k})`, text: v, home: undefined });
        });
      } else if (lp.startsWith("xl/drawings/") && !lp.includes("/_rels/")) {
        attrsOf(xml, () => true, (_l, a) => {
          if (a["textlink"]) out.push({ place: `shape text link in ${path}`, text: a["textlink"], home: undefined });
        });
      } else if (lp.startsWith("xl/pivotcache/pivotcachedefinition")) {
        attrsOf(xml, (l) => l === "worksheetSource", (_l, a) => {
          if (a["name"]) out.push({ place: `pivot table source ${path}`, text: a["name"], home: a["sheet"] });
        });
      } else if (lp.startsWith("xl/externallinks/") && !lp.includes("/_rels/") && fileName !== undefined) {
        const target = pkg.rels(path).find((r) => r.external)?.target;
        if (target === undefined || lower(baseName(decodeURIComponent(target))) !== lower(baseName(fileName))) continue;
        attrsOf(xml, (l) => l === "definedName", (_l, a) => {
          if (a["name"]) out.push({ place: `link of the workbook to itself ${path}`, text: a["name"], home: undefined });
        });
      }
    } catch {
      // A part that does not read as XML holds no reference we could rewrite either.
    }
  }
  return out;
}

/**
 * What the renames in `ctx` meet in the workbook, per renamed name (its lower-case key
 * before the rename). `bytes` adds the readers only the package has (hyperlinks, form
 * controls, pivot sources, a self-link); without it those are not looked at.
 */
export function renameSites(wb: WorkbookSnapshot, ctx: RenameContext, bytes?: Uint8Array, fileName?: string): Map<string, RenameSites> {
  const out = new Map<string, RenameSites>();
  for (const k of ctx.renames.keys()) out.set(k, { counts: { cells: 0, formats: 0, validations: 0, names: 0 }, blocked: [], captured: [], unparsed: [] });
  if (ctx.empty) return out;
  const all = [...ctx.renames.keys()];
  const byNewId = (id: string): string[] => all.filter((k) => lower(ctx.renames.get(k)!.rename.to) === lower(id.slice(id.lastIndexOf("!") + 1)));
  const spelling = (text: string): string[] => {
    const ids = new Set(tokenize(text).filter((t) => t.kind === "name" && t.value !== undefined).map((t) => lower(t.value!)));
    return all.filter((k) => ids.has(lower(ctx.renames.get(k)!.rename.from)));
  };
  const note = (keys: Iterable<string>, field: "blocked" | "captured" | "unparsed", place: string): void => {
    for (const k of keys) {
      const l = out.get(k)![field];
      if (!l.includes(place)) l.push(place);
    }
  };
  /** One formula: counted when the build rewrites it, else the place blocks the rename. */
  const visit = (text: string, home: string | undefined, place: string, kind: keyof RenameReferences | undefined): void => {
    const r = renameInFormula(text, home, ctx);
    if (r.unparsed) {
      note(spelling(text), "unparsed", place);
      return;
    }
    if (r.captured) {
      for (const id of r.captured) {
        const ks = byNewId(id);
        note(ks.length ? ks : all, "captured", `${place} (would read ${id})`);
      }
      return;
    }
    const keys = new Set(r.keys);
    if (keys.size === 0) return;
    if (kind === undefined) note(keys, "blocked", place);
    else for (const k of keys) out.get(k)!.counts[kind]++;
  };

  for (const site of collectSites(wb)) {
    const w = site.where;
    if (w.kind === "name") visit(site.stored, site.home, `name ${w.key}`, "names");
    else if (w.kind === "cell") visit(site.stored, site.home, `cell ${w.sheet}!${w.ref}`, "cells");
    else if (w.kind === "cf") visit(site.stored, site.home, `conditional format ${w.sheet}!${w.ref}`, "formats");
    else if (w.kind === "dv") visit(site.stored, site.home, `validation ${w.sheet}!${w.ref}`, "validations");
    else if (w.kind === "table") visit(site.stored, site.home, `Table column ${w.name}`, undefined);
  }
  for (const chart of wb.charts) for (const f of chart.formulas) visit(f.text, chart.sheet?.name, `chart ${chart.part}`, undefined);
  if (bytes) for (const o of otherNameReaders(bytes, fileName)) visit(o.text, o.home, o.place, undefined);
  return out;
}

