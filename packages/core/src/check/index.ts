// The source model and the checker the editor and the build share (M3d).
export { LineIndex, type LinePos } from "./lines.js";
export { SourceModel, occurrences, defKey, type SourceFile, type NameDef, type Loc, type Occurrence, type Analysis, type ScopeIndex } from "./model.js";
export { checkFile, checkProject, spillMap, strayFile, invalidName, xlmNameWarning, renameBuilt, renamedFrom, renamedTo, sheetPrefix, oneEditAway, KIND_LABEL, LIVE_CODES, type CheckContext, type Fix, type Problem } from "./check.js";
