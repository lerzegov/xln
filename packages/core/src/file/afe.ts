// Microsoft's Advanced Formula Environment (AFE, in the Excel Labs add-in) keeps its own
// copy of the names it manages, as module text, inside the workbook. xln reads names from
// <definedNames> only; this file recognises AFE's copy so that pull, check and build can
// say it is there and where it differs. Nothing here changes a byte of it (no hidden
// transformations: AFE's part stays as AFE wrote it).
//
// The format, from the add-in's own JavaScript bundle and an AFE-shaped sample file
// (probes/README.md § AFE-saved workbooks):
//
//   AFE 1.1+   customXml/itemN.xml
//                <AFEJSONBlob xmlns="http://schemas.advancedformulaenvironment.officeapps.live.com/afejsonblob/1.0">
//                  base64 of the UTF-16LE bytes of JSON.stringify(store)
//                </AFEJSONBlob>
//              store = { schema: ".../afeprojects/0.2" (or 0.1),
//                        files: [{ path: "/projects/<Module>", text }],
//                        projectNames: [the names AFE exported to the Name Manager],
//                        locale: { listSeparator, decimalSeparator, localeName, ... } }
//              AFE finds the part by its custom XML part ID, kept in the add-in's
//              settings (xl/webextensions/webextensionN.xml, property `projectV0_1-…`).
//   AFE 1.0    a very hidden sheet "AFE_hidden_codesheet_49ddb8b8" holding the module
//              text in cells; AFE 1.1 converts it to the part and leaves the sheet.
//   Locale     AFE detects the argument separator on a scratch sheet named
//              "e00eb4de3c8a421cba9b8f4cb8546ec" (a GUID cut to Excel's 31 characters),
//              deleted after use by current builds; older ones left it very hidden.
//
// A module's names reach the Name Manager as `<Module>.<name>`, except the `Workbook`
// module's, which are unprefixed.

import { fromBase64 } from "./base64.js";
import type { Package } from "./package.js";
import type { ForeignModuleStore, Sheet } from "./types.js";
import { descendants, ownText, parseXml } from "./xml.js";

/** The namespace of AFE's custom XML part (measured, AFE bundle 2025-11 and a sample file). */
export const AFE_BLOB_NS = "http://schemas.advancedformulaenvironment.officeapps.live.com/afejsonblob/1.0";
export const AFE_BLOB_ROOT = "AFEJSONBlob";
/** Every AFE schema URI starts so: a part in a later AFE namespace is still recognised as AFE's. */
export const AFE_NS_PREFIX = "http://schemas.advancedformulaenvironment.officeapps.live.com/";
/** The store schemas this reader knows (0.1 is upgraded by AFE itself to 0.2 on load). */
export const AFE_PROJECT_SCHEMAS = [
  "http://schemas.advancedformulaenvironment.officeapps.live.com/afeprojects/0.1",
  "http://schemas.advancedformulaenvironment.officeapps.live.com/afeprojects/0.2",
];
/** The module whose names are not prefixed in the Name Manager. */
export const AFE_PRIMARY_MODULE = "Workbook";
export const AFE_CODE_SHEET = "AFE_hidden_codesheet_49ddb8b8";
export const AFE_LOCALE_SHEET = "e00eb4de3c8a421cba9b8f4cb8546ec";
/** The settings key under which AFE keeps the ID of its part. */
export const AFE_SETTINGS_KEY = "projectV0_1-56c6e055-265e-4713-816e-a646dbb708de";

/** UTF-16LE bytes as text (AFE encodes its JSON so before base64). */
function utf16le(bytes: Uint8Array): string {
  let s = "";
  const CHUNK = 8192;
  for (let i = 0; i + 1 < bytes.length; i += CHUNK * 2) {
    const codes: number[] = [];
    for (let j = i; j + 1 < bytes.length && j < i + CHUNK * 2; j += 2) codes.push(bytes[j]! | (bytes[j + 1]! << 8));
    s += String.fromCharCode(...codes);
  }
  return s;
}

function isObject(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

/** The module name of a store path: `/projects/FN` → `FN`. */
export function afeModuleName(path: string): string {
  const parts = path.split("/").filter((p) => p !== "");
  return parts[parts.length - 1] ?? path;
}

/** The name a module entry has in the Name Manager. */
export function afeExportedName(module: string, name: string): string {
  return module === AFE_PRIMARY_MODULE ? name : `${module}.${name}`;
}

/** Reads the store out of the part's text. `unreadable` says why when it cannot. */
export function parseAfeBlob(xml: string): Pick<ForeignModuleStore, "schema" | "modules" | "exportedNames" | "locale" | "unreadable"> {
  let root;
  try {
    root = parseXml(xml);
  } catch (e) {
    return { unreadable: `the part is not well-formed XML (${(e as Error).message})` };
  }
  if (root.ns !== AFE_BLOB_NS || root.local !== AFE_BLOB_ROOT) return { unreadable: `a later AFE format (${root.local} in ${root.ns}), which xln does not read` };
  const bytes = fromBase64(ownText(root));
  if (!bytes) return { unreadable: "its content is not base64" };
  let store: unknown;
  try {
    store = JSON.parse(utf16le(bytes));
  } catch {
    return { unreadable: "its content is not the JSON AFE writes" };
  }
  if (!isObject(store)) return { unreadable: "its content is not the JSON AFE writes" };
  const schema = typeof store["schema"] === "string" ? store["schema"] : undefined;
  if (schema === undefined || !AFE_PROJECT_SCHEMAS.includes(schema)) return { ...(schema ? { schema } : {}), unreadable: `an AFE store schema xln does not know (${schema ?? "none"})` };
  const files = Array.isArray(store["files"]) ? store["files"] : [];
  const modules = files.filter(isObject).filter((f) => typeof f["path"] === "string" && typeof f["text"] === "string").map((f) => ({ name: afeModuleName(f["path"] as string), path: f["path"] as string, text: f["text"] as string }));
  const exportedNames = Array.isArray(store["projectNames"]) ? store["projectNames"].filter((n): n is string => typeof n === "string") : [];
  const out: ReturnType<typeof parseAfeBlob> = { schema, modules, exportedNames };
  const loc = store["locale"];
  if (isObject(loc)) {
    const locale: NonNullable<ForeignModuleStore["locale"]> = {};
    for (const k of ["listSeparator", "decimalSeparator", "localeName"] as const) if (typeof loc[k] === "string") locale[k] = loc[k] as string;
    out.locale = locale;
  }
  return out;
}

/** The custom XML part IDs AFE's settings point at (`{FA35…}`), from the add-in's web extension parts. */
function settingsItemIds(pkg: Package): Set<string> {
  const out = new Set<string>();
  for (const name of pkg.names) {
    const n = name.toLowerCase();
    if (!n.startsWith("xl/webextensions/webextension") || !n.endsWith(".xml")) continue;
    let root;
    try {
      root = parseXml(pkg.text(name) ?? "");
    } catch {
      continue;
    }
    for (const el of descendants(root)) {
      if (el.local !== "property" || el.attrs["name"] !== AFE_SETTINGS_KEY) continue;
      try {
        const v = JSON.parse(el.attrs["value"] ?? "");
        if (isObject(v) && typeof v["id"] === "string") out.add(v["id"].toUpperCase());
      } catch {
        // A setting AFE did not write: no link.
      }
    }
  }
  return out;
}

/** The datastore item ID in a custom XML item's properties part, upper case. */
function itemIdOf(pkg: Package, item: string): string | undefined {
  for (const r of pkg.rels(item)) {
    if (!r.type.endsWith("/customXmlProps")) continue;
    try {
      const root = parseXml(pkg.text(r.target) ?? "");
      for (const [k, v] of Object.entries(root.attrs)) if (k === "itemID" || k.endsWith(":itemID")) return v.toUpperCase();
    } catch {
      return undefined;
    }
  }
  return undefined;
}

/**
 * AFE's copies of its modules in a package: its custom XML part (recognised by the
 * namespace of the part's root element, whatever Excel numbered it) and the sheets of
 * AFE 1.0 and of its locale detection (recognised by their fixed names).
 */
export function findForeignModuleStores(pkg: Package, sheets: readonly Sheet[]): ForeignModuleStore[] {
  const out: ForeignModuleStore[] = [];
  let linked: Set<string> | undefined;
  for (const name of pkg.names) {
    const n = name.toLowerCase();
    if (!n.startsWith("customxml/") || n.includes("/_rels/") || !n.endsWith(".xml") || n.includes("itemprops")) continue;
    const text = pkg.text(name);
    if (text === undefined) continue;
    let root;
    try {
      root = parseXml(text);
    } catch {
      continue;
    }
    if (!root.ns.startsWith(AFE_NS_PREFIX)) continue;
    const store: ForeignModuleStore = { tool: "afe", kind: "custom-xml", part: name, namespace: root.ns, ...parseAfeBlob(text) };
    const id = itemIdOf(pkg, name);
    if (id !== undefined) {
      store.itemId = id;
      linked ??= settingsItemIds(pkg);
      store.linked = linked.has(id);
    }
    out.push(store);
  }
  for (const s of sheets) {
    const lower = s.name.toLowerCase();
    const kind = lower === AFE_CODE_SHEET.toLowerCase() ? "code-sheet" : lower === AFE_LOCALE_SHEET ? "locale-sheet" : undefined;
    if (!kind) continue;
    const store: ForeignModuleStore = { tool: "afe", kind, part: s.part ?? "", sheet: s.name, state: s.state };
    if (kind === "code-sheet") store.unreadable = "AFE 1.0 kept the module text in the cells of this sheet; xln does not read them";
    out.push(store);
  }
  return out;
}

/** Whether a custom XML item's text belongs to AFE (for callers holding the text alone). */
export function isAfeXml(xml: string): boolean {
  try {
    return parseXml(xml).ns.startsWith(AFE_NS_PREFIX);
  } catch {
    return false;
  }
}
