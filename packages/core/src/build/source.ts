// The source side of a build: every `names/**/*.xln` file of a project, parsed into the
// names and the cell statements it declares. A name in a sheet file is local to the file's
// sheet unless annotated `@workbook`; elsewhere its scope is its `@scope` block's. A cell's
// sheet is its address's, else the sheet file's or the block's. A name declared twice is
// an error.

import { parseSourceFile, type ModuleEntry, type ParsedModule } from "../project/module.js";
import { NAMES_DIR } from "../project/pull.js";
import { FROM, parseLibBase } from "../project/provenance.js";

/** `@renamed(Old)`: this name was called Old (D4). `@renamed(Sheet!Old)` names the old
 *  scope when it differs; `@renamed(!Old)` is workbook scope. */
export const RENAMED = "renamed";

/** Where a cell statement's address is written in its file, for diagnostics and quick fixes. */
export interface SourceAddress {
  /** The sheet of the cell: written in the address (`@BS!C6`), or a sheet file's; undefined otherwise. */
  sheet: string | undefined;
  /** Whether the address names its sheet (`@BS!C6`). */
  written?: boolean;
  /** `C6` or `B40:G40`, upper case, no `$`. */
  range: string;
  /** Offsets in the file: the range's text (after `Sheet!`, before a `#`). */
  rangeStart: number;
  rangeEnd: number;
  /** `@C6#`: the name covers the cell's spill. */
  spill: boolean;
}

export interface SourceName {
  name: string;
  /** The sheet of its `@scope` block, as written; undefined for workbook scope. */
  scope: string | undefined;
  hidden: boolean;
  /** The doc comment: the Name Manager comment. */
  doc: string | undefined;
  /** Display form, comments removed. For a named cell statement: the cell's formula. */
  formula: string;
  /** From `@renamed(…)`: the name's previous name and scope (undefined: same scope). */
  renamedFrom?: { name: string; scope: string | null | undefined };
  /** From `@from(lib #…)`: the library base, 6 hex digits; the build writes it into the provenance tag (`lib#…`). */
  libBase?: string;
  /** Set for a named cell statement (`Name @C6 = …;`): the name's definition is that cell. */
  cell?: SourceAddress;
  file: string;
  line: number;
}

/** A cell statement (M3b): `Name @C6 = …;`, `Name @C6 = ;` or `@C5 = …;`. */
export interface SourceCell {
  /** The name of a named cell or a slot; undefined for an unnamed cell. */
  name: string | undefined;
  /** The `@scope` block's sheet; undefined in a `@workbook` block. */
  scope: string | undefined;
  address: SourceAddress;
  /** Display form, comments removed; "" for a slot or a cleared cell. */
  formula: string;
  file: string;
  line: number;
}

export interface SourceProblem {
  severity: "error" | "warning";
  code: string;
  message: string;
  file?: string;
  line?: number;
  /** The name concerned, as `Sheet!Name` or `Name`. */
  key?: string;
  /** Places that use the name (cells, formats, validations, Table columns, charts, names). */
  sites?: string[];
  /** A text edit that fixes the problem in `file`: replace `start`…`end` with `text`. */
  fix?: { start: number; end: number; text: string };
}

export interface SourceProject {
  names: SourceName[];
  /** Every cell statement, files in path order, each file in text order. */
  cells: SourceCell[];
  problems: SourceProblem[];
}

/** Whether a project path is a names file. */
export function isSourcePath(path: string): boolean {
  return path.startsWith(NAMES_DIR + "/") && path.endsWith(".xln");
}

/** A cell statement's address; in a sheet file of the new form the sheet is the file's, written or not. */
function addressOf(e: ModuleEntry, pm: ParsedModule): SourceAddress | undefined {
  const c = e.cell;
  const sheet = c?.sheet ?? (pm.form === "sheet" ? e.cellSheet : undefined);
  return c && { sheet, written: c.sheet !== undefined, range: c.range, rangeStart: c.rangeStart, rangeEnd: c.spill ? c.end - 1 : c.end, spill: c.spill };
}

export function readSourceProject(files: Readonly<Record<string, string>>): SourceProject {
  const names: SourceName[] = [];
  const cells: SourceCell[] = [];
  const problems: SourceProblem[] = [];
  const seen = new Map<string, SourceName>();
  for (const file of Object.keys(files).filter(isSourcePath).sort()) {
    const parsed = parseSourceFile(file, files[file]!);
    for (const d of parsed.diagnostics) problems.push({ severity: "error", code: "syntax", message: d.message, file, line: d.line });
    for (const e of parsed.cells) {
      cells.push({ name: e.name === "" ? undefined : e.name, scope: e.scope, address: addressOf(e, parsed)!, formula: e.formula, file, line: e.line });
    }
    for (const e of parsed.entries) {
      const n: SourceName = { name: e.name, scope: e.scope, hidden: e.hidden, doc: e.doc, formula: e.formula, file, line: e.line };
      const address = addressOf(e, parsed);
      if (address) n.cell = address;
      if (e.from !== undefined) n.libBase = e.from;
      const known = new Set(["hidden", "workbook", RENAMED, FROM]);
      for (const a of e.annotations) {
        if (!known.has(a.name)) {
          problems.push({ severity: "warning", code: "annotation", message: `@${a.name} is not known to the build; ignored`, file, line: a.line });
        } else if (a.name === FROM) {
          const b = parseLibBase(a.arg);
          if ("error" in b) problems.push({ severity: "error", code: "annotation", message: b.error, file, line: a.line });
        } else if (a.name === RENAMED) {
          const arg = (a.arg ?? "").trim();
          if (arg === "") {
            problems.push({ severity: "error", code: "annotation", message: "@renamed needs the previous name: @renamed(OldName)", file, line: a.line });
            continue;
          }
          const bang = arg.lastIndexOf("!");
          if (bang < 0) n.renamedFrom = { name: arg, scope: undefined };
          else {
            let sheet = arg.slice(0, bang).trim();
            if (sheet.length >= 2 && sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).split("''").join("'");
            n.renamedFrom = { name: arg.slice(bang + 1).trim(), scope: sheet === "" ? null : sheet };
          }
        }
      }
      const key = `${(e.scope ?? "").toLowerCase()}!${e.name.toLowerCase()}`;
      const first = seen.get(key);
      if (first) {
        problems.push({
          severity: "error",
          code: "duplicate",
          message: `${e.scope === undefined ? e.name : `${e.scope}!${e.name}`} is also defined in ${first.file}:${first.line}`,
          file,
          line: e.line,
        });
        continue;
      }
      seen.set(key, n);
      names.push(n);
    }
  }
  return { names, cells, problems };
}
