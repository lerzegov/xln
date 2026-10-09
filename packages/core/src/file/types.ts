// The read-only picture of a workbook that the file layer hands to the layers above.
// Plain data (arrays and records, no Maps or classes) so it can be serialised as is.

export type SheetState = "visible" | "hidden" | "veryHidden";

/** What the workbook relationship says the sheet part is. */
export type SheetKind = "worksheet" | "chartsheet" | "dialogsheet" | "macrosheet" | "unknown";

/** A sheet named by its position in `<sheets>` (0-based, what `localSheetId` counts). */
export interface SheetRef {
  position: number;
  name: string;
}

export interface Sheet {
  name: string;
  sheetId: number;
  /** 0-based index in `<sheets>`. */
  position: number;
  state: SheetState;
  kind: SheetKind;
  relId: string;
  /** Package path of the sheet part, e.g. `xl/worksheets/sheet1.xml`; undefined if the part is missing. */
  part: string | undefined;
  formulas: CellFormula[];
  sharedFormulas: SharedFormulaGroup[];
  /** Dynamic-array anchors and the extent they spilled to when the file was saved. */
  spills: Spill[];
  conditionalFormats: ConditionalFormat[];
  dataValidations: DataValidation[];
  /** Table names (`displayName`) on this sheet, in `<tableParts>` order. */
  tables: string[];
}

export type NameScope = { kind: "workbook" } | ({ kind: "sheet" } & SheetRef);

export interface DefinedName {
  name: string;
  scope: NameScope;
  hidden: boolean;
  comment: string | undefined;
  /**
   * The definition exactly as stored: XML entities decoded, nothing else. CR LF line
   * breaks are kept, and stored prefixes (`_xlfn.`, `_xlpm.`, ...) are not removed.
   */
  definition: string;
  /** Every attribute other than name, localSheetId, hidden and comment, raw. */
  attributes: Record<string, string>;
  /** Position in `<definedNames>` (file order). */
  index: number;
  /** Starts with `_xl` (any case): Excel's own names (`_xlnm.Print_Area`, stray `_xlfn.*`). */
  isXlPrefixed: boolean;
  /** Excel built-in name (`_xlnm.` prefix: print area, print titles, filter database...). */
  isBuiltIn: boolean;
  /** `localSheetId` pointed at a position with no sheet. */
  scopeInvalid?: true;
}

export interface TableColumn {
  id: number;
  name: string;
  calculatedColumnFormula?: string;
  /** The calculated column formula is an array formula. */
  calculatedColumnArray?: true;
  totalsRowFormula?: string;
  totalsRowFunction?: string;
}

export interface Table {
  id: number;
  name: string;
  displayName: string;
  sheet: SheetRef;
  part: string;
  ref: string;
  headerRowCount: number;
  totalsRowCount: number;
  columns: TableColumn[];
}

export type FormulaKind =
  /** An ordinary single-cell formula. */
  | "normal"
  /** First cell of a shared-formula group; carries the text and the group's `ref`. */
  | "shared-master"
  /** Other cells of a group: no text of their own, see `master`. */
  | "shared-child"
  /** Legacy (Ctrl+Shift+Enter) array formula over `range`. */
  | "array"
  /** Dynamic-array formula (`t="array"` plus `cm` pointing at XLDAPR metadata). */
  | "dynamic-array"
  /** What-if data table (`t="dataTable"`). */
  | "data-table";

export type CellValueType = "n" | "s" | "str" | "b" | "e" | "inlineStr" | "d";

export interface CachedValue {
  /** The cell's `t` attribute (default `n`). */
  type: CellValueType;
  /** The `<v>` text as stored (for `s`, the shared-string index); undefined if absent. */
  raw: string | undefined;
  /**
   * Decoded value: number for `n`, boolean for `b`, the string for `s`/`str`/`inlineStr`,
   * the error text (`#REF!`) for `e`, the ISO text for `d`; undefined when no value stored.
   */
  value: number | string | boolean | undefined;
}

export interface CellFormula {
  /** Cell address, e.g. `B3`. */
  cell: string;
  kind: FormulaKind;
  /** Formula text as stored (no leading `=`); undefined for a shared child. */
  text: string | undefined;
  /** The `ref` attribute: group extent (shared master), array extent (array, dynamic-array). */
  range?: string;
  /** Shared-formula group index (`si`), for masters and children. */
  si?: number;
  /** For a shared child: the master's cell address. Undefined if the file has no master for `si`. */
  master?: string;
  /** The cell's `cm` attribute (1-based cell metadata index), when present. */
  cm?: number;
  /** Other attributes of `<f>` (aca, ca, dt2D, dtr, r1, r2, del1, del2, bx), raw. */
  attributes: Record<string, string>;
  value: CachedValue;
}

export interface SharedFormulaGroup {
  si: number;
  /** Master cell address; undefined if no master was found (malformed file). */
  master: string | undefined;
  /** The master's `ref`. */
  range: string | undefined;
  /** The master's text. */
  text: string | undefined;
  /** Every cell in the group, master first, in file order. */
  cells: string[];
}

export interface Spill {
  /** Anchor cell, e.g. `E1`. */
  anchor: string;
  /** Extent saved with the file, e.g. `E1:E3` (a single cell when it did not spill). */
  extent: string;
}

export interface ConditionalFormat {
  sqref: string;
  type: string | undefined;
  priority: number | undefined;
  /** Rule formulas in document order (`<formula>`, or `<xm:f>` in the x14 extension). */
  formulas: string[];
  /** Read from the `x14` extension list rather than the main element. */
  ext: boolean;
}

export interface DataValidation {
  sqref: string;
  type: string | undefined;
  formula1: string | undefined;
  formula2: string | undefined;
  ext: boolean;
}

export type OtherPartKind =
  | "chart"
  | "chartsheet"
  | "drawing"
  | "pivotTable"
  | "pivotCache"
  | "externalLink"
  | "customXml"
  | "queryTable"
  | "connections"
  | "slicer"
  | "slicerCache"
  | "timeline"
  | "vbaProject"
  | "richData";

export interface OtherPart {
  kind: OtherPartKind;
  path: string;
}

/** A formula in a chart part (`<c:f>`): a series' values or categories, a title, a label. */
export interface ChartFormula {
  /** As stored: `'S1'!$B$2:$B$9`, `[0]!Sales` (a workbook-scoped name), `'S1'!Sales`. */
  text: string;
  /** The element it belongs to, past the `numRef`/`strRef` wrapper: `val`, `cat`, `tx`, `xVal`, `yVal`, ... */
  element: string;
}

export interface ChartPart {
  /** Package path, e.g. `xl/charts/chart1.xml`. */
  part: string;
  /** `chartEx`: an Office 2016 chart (waterfall, histogram, ...). */
  kind: "chart" | "chartEx";
  /** The sheet whose drawing shows the chart (a worksheet or a chartsheet); undefined if none does. */
  sheet?: SheetRef;
  /** The drawing part that places it. */
  drawing?: string;
  /** In document order. */
  formulas: ChartFormula[];
}

/** One module of a foreign module store, as its tool stored it. */
export interface ForeignModule {
  /** `FN`; AFE's primary module is `Workbook` (its names are not prefixed). */
  name: string;
  /** Path in the store (`/projects/FN`). */
  path: string;
  /** The module text as stored (LF line breaks). */
  text: string;
}

/**
 * A copy of names as module text that another tool keeps inside the workbook: so far
 * Microsoft's Advanced Formula Environment (AFE, Excel Labs). xln reads it, never changes it.
 */
export interface ForeignModuleStore {
  tool: "afe";
  /** `custom-xml`: AFE 1.1+'s part; `code-sheet`: AFE 1.0's very hidden sheet; `locale-sheet`: AFE's separator-detection sheet. */
  kind: "custom-xml" | "code-sheet" | "locale-sheet";
  /** Package path of the custom XML item or of the sheet's part ("" when the sheet has none). */
  part: string;
  /** For a sheet: its name and visibility. */
  sheet?: string;
  state?: SheetState;
  /** For the custom XML part: its root element's namespace. */
  namespace?: string;
  /** The datastore item ID (`{FA35…}`), from the item's properties. */
  itemId?: string;
  /** Whether AFE's settings in the workbook (its web extension part) point at this item. */
  linked?: boolean;
  /** The store's schema URI (`…/afeprojects/0.2`). */
  schema?: string;
  modules?: ForeignModule[];
  /** The names AFE says it exported to the Name Manager (`projectNames`). */
  exportedNames?: string[];
  /** The separators the module text is written with. */
  locale?: { listSeparator?: string; decimalSeparator?: string; localeName?: string };
  /** Why the content could not be read; absent when it was (or, for a locale sheet, there is none). */
  unreadable?: string;
}

/**
 * A link to another workbook (`<externalReference>` of workbook.xml, probe F10): stored
 * formulas name the file by its number, `[1]Sheet1!$A$1`, `[1]!Name`. Excel writes the part
 * (the file's path and cached values) when a formula first names the file; xln keeps it.
 */
export interface ExternalLink {
  /** 1-based position in `<externalReferences>`: the `n` of `[n]`. */
  index: number;
  /** The file name, `Other.xlsx` (no folder): what formulas show in place of `[n]`. */
  book: string;
  /** The link's target as stored (an absolute or relative path, a URL). */
  target: string;
  /** Package path of the link part, `xl/externalLinks/externalLink1.xml`. */
  part: string;
}

export interface WorkbookSnapshot {
  /** In `<sheets>` order. */
  sheets: Sheet[];
  /** In file order; nothing skipped. */
  definedNames: DefinedName[];
  tables: Table[];
  /** Chart parts and the formulas they read (series, titles, labels), in package order. */
  charts: ChartPart[];
  /** Links to other workbooks, in `<externalReferences>` order. */
  externalLinks: ExternalLink[];
  /** Parts a later layer must check before changing names (charts, pivots, external links...). */
  otherParts: OtherPart[];
  /** Every entry in the zip, in archive order. */
  parts: string[];
  /** Other tools' copies of names as module text (AFE's custom XML part, its hidden sheets); read only. */
  foreignModuleStores: ForeignModuleStore[];
  /** Package path of the workbook part (normally `xl/workbook.xml`). */
  workbookPart: string;
  /**
   * `<calcPr fullCalcOnLoad="1">`: the file asks Excel to recalculate everything on open.
   * Excel drops it when it saves (probe F06), so it marks a file a tool wrote (an xln build)
   * that Excel has not saved since.
   */
  fullCalcOnLoad?: true;
  /** Non-fatal oddities met while reading (missing parts, dangling references). */
  warnings: string[];
}
