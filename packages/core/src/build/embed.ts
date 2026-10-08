// D5 in the package: put the xln custom XML part (embedXml.ts) into a workbook, or bring
// the one a previous build put there up to date. A workbook never gets a second xln part:
// the part is found by its namespace, wherever Excel has numbered it. A new part comes
// with its item properties, the item's relationship to them, a relationship from the
// workbook part and the content type of the properties part, as Excel itself writes a
// custom XML part. Every other entry is copied byte for byte (zip.ts).

import { decodeText, Package, relTypeIs } from "../file/package.js";
import { childElements, parseXml } from "../file/xml.js";
import { rewriteZip } from "../file/zip.js";
import { utf8 } from "../project/hash.js";
import { workbookPartOf, relsPartOf } from "./apply.js";
import {
  CUSTOM_XML_PROPS_CT,
  CUSTOM_XML_PROPS_REL,
  CUSTOM_XML_REL,
  embeddedSourcePropsXml,
  embeddedSourceXml,
  findEmbeddedPart,
  readEmbeddedSource,
} from "./embedXml.js";
import { addOverride, addRelationship } from "./packageXml.js";
import { contentTypes, relations, sameMaps } from "./readbackCells.js";

export type EmbedOutcome = "created" | "updated" | "unchanged";

/** A relationship target for `part` from the part `from` (`xl/workbook.xml` → `../customXml/item1.xml`). */
function relativeFrom(from: string, part: string): string {
  const up = from.split("/").length - 1;
  return "../".repeat(up) + part;
}

function hasXmlDefault(ct: string): boolean {
  return childElements(parseXml(ct)).some((e) => e.local === "Default" && (e.attrs["Extension"] ?? "").toLowerCase() === "xml");
}

/** Embeds `files` in the workbook; returns the new bytes (the same bytes when the part already says this). */
export function applyEmbeddedSource(bytes: Uint8Array, files: Readonly<Record<string, string>>): { bytes: Uint8Array; outcome: EmbedOutcome; item: string } {
  const pkg = new Package(bytes);
  const xml = embeddedSourceXml(files);
  const found = findEmbeddedPart(pkg);
  if (found) {
    if (pkg.text(found.item) === xml) return { bytes, outcome: "unchanged", item: found.item };
    return { bytes: rewriteZip(bytes, { replace: new Map([[found.item, utf8(xml)]]) }), outcome: "updated", item: found.item };
  }

  const wbPart = workbookPartOf(pkg);
  let n = 1;
  const taken = (k: number) => pkg.has(`customXml/item${k}.xml`) || pkg.has(`customXml/itemProps${k}.xml`) || pkg.has(`customXml/_rels/item${k}.xml.rels`);
  while (taken(n)) n++;
  const item = `customXml/item${n}.xml`;
  const props = `customXml/itemProps${n}.xml`;
  const itemRels = `customXml/_rels/item${n}.xml.rels`;
  const add = new Map<string, Uint8Array>([
    [item, utf8(xml)],
    [props, utf8(embeddedSourcePropsXml())],
    [
      itemRels,
      utf8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          `<Relationship Id="rId1" Type="${CUSTOM_XML_PROPS_REL}" Target="itemProps${n}.xml"/></Relationships>`,
      ),
    ],
  ]);
  const replace = new Map<string, Uint8Array>();

  const relsName = pkg.find(relsPartOf(wbPart));
  if (!relsName) throw new Error(`${relsPartOf(wbPart)} is missing`);
  const rels = addRelationship(decodeText(pkg.raw(relsName)!), CUSTOM_XML_REL, relativeFrom(wbPart, item)).xml;
  replace.set(relsName, utf8(rels));

  const ctName = pkg.find("[Content_Types].xml");
  if (!ctName) throw new Error("[Content_Types].xml is missing");
  let ct = decodeText(pkg.raw(ctName)!);
  ct = addOverride(ct, "/" + props, CUSTOM_XML_PROPS_CT);
  if (!hasXmlDefault(ct)) ct = addOverride(ct, "/" + item, "application/xml");
  replace.set(ctName, utf8(ct));
  return { bytes: rewriteZip(bytes, { replace, add }), outcome: "created", item };
}

/** What a build that embeds the source may change in the package, for read-back. */
export interface EmbedFacts {
  mayChange: string[];
  added: Set<string>;
  /** Relationship Ids in the workbook part's relationships that the embedding added. */
  addedRelIds: Set<string>;
  /** Content-type overrides (lower-case part names with `/`) that the embedding added. */
  addedOverrides: Set<string>;
}

function sameFiles(a: Readonly<Record<string, string>>, b: Readonly<Record<string, string>>): boolean {
  const ka = Object.keys(a);
  const kb = Object.keys(b);
  return ka.length === kb.length && ka.every((k, i) => k === kb[i] && a[k] === b[k]);
}

/** Checks the embedded part in the built bytes says exactly `files`; returns what the embedding may have touched. */
export function embedPackageFacts(beforeBytes: Uint8Array, afterBytes: Uint8Array, files: Readonly<Record<string, string>>, problems: string[], alone: boolean): EmbedFacts {
  const a = new Package(beforeBytes);
  const b = new Package(afterBytes);
  const facts: EmbedFacts = { mayChange: [], added: new Set(), addedRelIds: new Set(), addedOverrides: new Set() };
  const got = readEmbeddedSource(b);
  if (!got) {
    problems.push("the embedded source part is missing from the built file");
    return facts;
  }
  if (got.damaged.length) problems.push(`the embedded source has damaged files: ${got.damaged.join(", ")}`);
  if (!sameFiles(got.files, files)) problems.push("the embedded source differs from the project files it should carry");
  if (got.location.relId === undefined) problems.push("the embedded source part is not related from the workbook part");
  const was = findEmbeddedPart(a);
  if (was) {
    if (was.item !== got.location.item) problems.push(`the embedded source moved from ${was.item} to ${got.location.item}`);
    facts.mayChange.push(was.item);
    return facts;
  }
  // A new part: the item, its properties and their relationship, one relationship from the workbook, one content type.
  const item = got.location.item;
  facts.added.add(item);
  const dir = item.slice(0, item.lastIndexOf("/") + 1);
  const relsPart = `${dir}_rels/${item.slice(dir.length)}.rels`;
  const itemRels = b.rels(item);
  if (!b.has(relsPart) || a.has(relsPart)) problems.push(`${relsPart} should be a new part`);
  else facts.added.add(b.find(relsPart)!);
  const propsRel = itemRels.find((r) => relTypeIs(r.type, "customXmlProps"));
  if (!propsRel || !b.has(propsRel.target) || a.has(propsRel.target)) problems.push(`the item properties of ${item} should be a new part`);
  else {
    facts.added.add(b.find(propsRel.target)!);
    facts.addedOverrides.add("/" + propsRel.target.toLowerCase());
  }
  facts.addedOverrides.add("/" + item.toLowerCase());
  const wbPart = workbookPartOf(b);
  facts.addedRelIds.add(got.location.relId ?? "");
  const relsName = a.find(relsPartOf(wbPart));
  if (relsName) facts.mayChange.push(relsName);
  const ctName = a.find("[Content_Types].xml");
  if (ctName) facts.mayChange.push(ctName);
  // Without cell changes nothing else touches these two parts: they must differ by the new entries only.
  // (With cell changes, the cell check compares them, leaving these entries out.)
  if (alone && relsName) {
    const ra = relations(a, relsName);
    const rb = relations(b, relsName);
    for (const id of facts.addedRelIds) if (!ra.has(id)) rb.delete(id);
    if (!sameMaps(ra, rb)) problems.push(`${relsName} differs from the original by more than the embedded source`);
    const ca = contentTypes(a);
    const cb = contentTypes(b);
    for (const o of facts.addedOverrides) if (!ca.has(`O ${o}`)) cb.delete(`O ${o}`);
    if (!sameMaps(ca, cb)) problems.push("[Content_Types].xml differs from the original by more than the embedded source");
  }
  return facts;
}
