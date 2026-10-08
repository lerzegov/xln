// L2 project: classify names, propose modules, write names/*.xln, the manifest and the
// lockfile; parse .xln modules back.
export { pullProject, layoutToLf, NAMES_DIR, MANIFEST_FILE, LOCK_FILE } from "./pull.js";
export type { PullOptions, PullReport, PullResult, ModuleSummary, SheetFileSummary, CellCounts, ForeignModuleSummary } from "./pull.js";
export { classify, cellCallee, definitionTarget, sheetCellsOf, type CellCallee, type DefinitionTarget } from "./classify.js";
export { proposeModules, moduleFileName, sheetFileName, sheetFromFileName, fileSystemKey, compareNames, UNMANAGED, SHEETS_DIR, UNDERSCORE_MIN } from "./modules.js";
export { parseModule, parseSourceFile, sheetOfPath, convertSheetBlocks, convertModuleBlocks, formatEntry, formatScope, formatSheetAnnotation, scopeNeedsQuotes, formatDoc, docText, formulaToSource, sourceToFormula, scanCellTarget, formatCellAddress } from "./module.js";
export type { ModuleEntry, Annotation, ModuleDiagnostic, ParsedModule, ScopeDirective, WritableEntry, CellTarget, ParseOptions } from "./module.js";
export { nameUses, NameResolver, type NameUse } from "./refs.js";
export { labelName, labelNames, labelDrift, valueLabels, type LabelDrift, type ValueLabel } from "./labels.js";
export { labelNotice, labelNoticeLines, renameLabelNotices, labelReplacement, labelReplaceShort, labelGives, EXCEL_REPLACE_KEYS, type LabelNotice, type LabelRename, type LabelCell, type LabelReplace } from "./labelNotice.js";
export { buildManifest, MANIFEST_FORMAT } from "./manifest.js";
export { buildLockfile, definitionHash, commentHash, definitionHashV2, commentHashV2, definitionHashLike, commentHashLike, isV2Hash, normalizeDefinition, normalizeDefinitionV2, parseLockfile, splitNameKey, lockfileJson, lockfileText, lockCell, cellKey, LOCK_FORMAT, LOCK_FORMATS_READ, LOCK_FORMAT_IMPLICIT_SPILL, explicitSpill } from "./lockfile.js";
export type { Lockfile, LockEntry, LockCell, LockableStatement, LockHashes } from "./lockfile.js";
export { compressCells, extentSize } from "./cells.js";
export { cellStatements, sheetFormulaCells, sameFilled, rangeText, parseCellRange } from "./statements.js";
export type { CellStatement, CellStatementKind, CellStatementOptions, FormulaCell, SheetFormulaCells, ValueCells } from "./statements.js";
export { sha256, sha256Bytes, utf8 } from "./hash.js";
export { formatProvenanceTag, splitProvenance, stripProvenance, withProvenance, sourceHash, sourceHashV1, moduleOfPath, moduleVersion, commentLength, COMMENT_MAX, libraryHash, formatLibBase, parseLibBase, FROM, buildProvenanceTag, provenanceTagLength, docTagOverflow, type DocTagOverflow } from "./provenance.js";
export type { ProvenanceTag } from "./provenance.js";
export { provenanceStatus, type ProvenanceStatus } from "./provenance.js";
export { stringifyJson, type Json } from "./json.js";
export { parseDocComment, paramDoc, docParamSpans, type DocComment, type DocParam, type DocParamSpan } from "./doc.js";
export { nameKey } from "./types.js";
export type * from "./types.js";
export { renameInFormula, RenameContext, scopedLowerKey, type FormulaRename, type NameRename, type ScopedName } from "./rename.js";
