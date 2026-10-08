// M3e, the editor's housekeeping without the vscode API: which project a command means,
// which tabs point at files a pull removed, and the settings that hide a project's
// internals. The vscode side (ux.ts) only feeds URIs in as strings.

import { BASES_DIR, LOCK_FILE, MANIFEST_FILE } from "@xln/core";

/**
 * Projects in the order a pick offers them: the active editor's first (the default), then
 * the others as listed. `roots` and `active` are project folder URIs.
 */
export function projectChoices(roots: readonly string[], active: string | undefined): string[] {
  if (active === undefined || !roots.includes(active)) return [...roots];
  return [active, ...roots.filter((r) => r !== active)];
}

/** The project folder of a URI string at or below one of `roots`; undefined elsewhere. */
export function rootOf(uri: string, roots: readonly string[]): string | undefined {
  const s = uri.endsWith("/") ? uri.slice(0, -1) : uri;
  return roots.find((r) => s === r || s.startsWith(r + "/"));
}

export interface TabInfo {
  /** The tab's file URI. */
  uri: string;
  dirty: boolean;
}

/**
 * The tabs to close after a pull or on loading a project: those below a project folder of
 * `roots` whose file is gone (`exists` false). A dirty one is kept and reported instead:
 * closing it would lose text nobody saved.
 */
export function staleTabs(tabs: readonly TabInfo[], roots: readonly string[], exists: (uri: string) => boolean): { close: string[]; keep: string[] } {
  const close: string[] = [];
  const keep: string[] = [];
  for (const t of tabs) {
    const root = rootOf(t.uri, roots);
    if (root === undefined || t.uri === root || exists(t.uri)) continue;
    (t.dirty ? keep : close).push(t.uri);
  }
  return { close, keep };
}

/**
 * The glob patterns of a project's internal files: written by xln, not for editing. The
 * library bases' texts (`library-bases/`, kept by the library actions) are one more.
 */
export const INTERNAL_GLOBS: readonly string[] = [`**/*.xln/${MANIFEST_FILE}`, `**/*.xln/${LOCK_FILE}`, `**/*.xln/${BASES_DIR}`];

/**
 * The workspace's own `files.exclude` (`current`, its workspace value) with the internals
 * shown or hidden again. Shown: the patterns set to `false` there, over the extension's
 * default `true`. Hidden: those entries removed, so the default applies again. Every other
 * pattern is kept. `value` undefined removes the workspace setting (nothing left in it).
 */
export function internalsExclude(current: Readonly<Record<string, unknown>> | undefined, show: boolean): { changed: boolean; value: Record<string, unknown> | undefined } {
  const next: Record<string, unknown> = { ...(current ?? {}) };
  let changed = false;
  for (const g of INTERNAL_GLOBS) {
    if (show && next[g] !== false) {
      next[g] = false;
      changed = true;
    } else if (!show && g in next) {
      delete next[g];
      changed = true;
    }
  }
  return { changed, value: Object.keys(next).length ? next : undefined };
}

/** Whether `files.exclude` (the effective value) hides the internals. */
export function internalsHidden(exclude: Readonly<Record<string, unknown>> | undefined): boolean {
  return INTERNAL_GLOBS.every((g) => exclude?.[g] === true);
}
