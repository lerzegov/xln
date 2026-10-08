// The project layer's picture of a workbook's names: what `pullProject` computes before
// it renders files. Plain data, so the extension can keep it and serialise it.

/** A3: what a definition is, decided on its AST. */
export type NameKind = "constant" | "range" | "spill" | "table" | "formula" | "lambda" | "unparsed";

export interface Arity {
  /** Parameters without brackets. */
  required: number;
  /** Parameters written `[p]`. */
  optional: number;
}

export interface Classification {
  kind: NameKind;
  /** For `lambda`. */
  arity?: Arity;
  /** For `lambda`: parameter names as displayed, optional ones in brackets. */
  params?: string[];
  /** For `spill`: the anchor cell; `sheet` undefined when written without a sheet. */
  anchor?: { sheet: string | undefined; cell: string };
  /** For `table`: the Table name. */
  table?: string;
  /** For `unparsed`: the parser's message. */
  error?: string;
}

/** A defined name as the project sees it. */
export interface ProjectName {
  name: string;
  /** Sheet name for a sheet-scoped name; undefined for workbook scope. */
  scope: string | undefined;
  /** 0-based sheet position (`localSheetId`) for a sheet-scoped name. */
  scopePosition: number | undefined;
  hidden: boolean;
  /** Excel's comment, CR LF turned into LF, without its provenance tag; undefined when empty. */
  comment: string | undefined;
  /** As stored in the file (CR LF kept). */
  stored: string;
  /** Display form (A2), CR LF turned into LF in layout. */
  display: string;
  classification: Classification;
  /** The module that owns it; undefined when none does (`_unmanaged.xln` or `sheets/<Sheet>.xln`). */
  module: string | undefined;
  /** Set when the name is a cell statement (a named formula cell or a slot): it is then
   *  written in its sheet's file as `Name @C6 = formula;`, whatever its module. */
  cell?: { sheet: string; range: string };
  /** The library base its provenance tag records (`lib#…`): written back as `@from(lib #…)`. */
  libBase?: string;
}

/** `Sheet!Name` for a sheet-scoped name, `Name` for a workbook-scoped one. */
export function nameKey(n: { name: string; scope: string | undefined }): string {
  return n.scope === undefined ? n.name : `${n.scope}!${n.name}`;
}
