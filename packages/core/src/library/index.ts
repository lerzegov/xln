// M4 library governance: `.lambda` files read as they are, `lib status` against a workbook
// or a project, and the text edits of insert, update and publish.
export { parseLambdaFile, libraryDoc, lambdaFileName, lambdaParameters, collapseSpace, LIBRARY_DOC_MAX } from "./lambdaFile.js";
export type { LibraryFunction, LibraryField, LibraryProblem, ParsedLambdaFile } from "./lambdaFile.js";
export { readLibrary, isLambdaFile, libraryStored, libraryClosure, libraryFunctionHash, definitionBase, calledNames, type Library } from "./library.js";
export { lineDiff, definitionDiff, diffText, comparableLayout, type DiffLine } from "./diff.js";
export { libraryStatus, workbookCopies, projectCopies, projectLibraryStatus, workbookNameSet, withKnownVersions, withBases, libState, renderLibStatus, libStatusJson, libStateLabel, libStateNote, threeWayText, LIB_STATES } from "./status.js";
export type { LibState, LibCopy, LibStatusItem, LibStatusReport, LibStatusLink, KnownVersion } from "./status.js";
export { libraryInsertion, libraryReplacement, libraryBaseEdit, libraryBaseRecordings, publishedBase, publishLambda, applyEdits, moduleFileOf, findEntry, formulaRange, firstSentence, sameSummary, libraryModuleHeader, docTagWarnings } from "./edit.js";
export type { TextEdit, FileEdit, PublishSource, PublishResult, BaseRecording, DocTagWarning } from "./edit.js";
export { BASES_DIR, baseFilePath, isBasePath, libraryBase, libraryFunctionBase, baseFileText, parseBaseFile, readBases, baseFiles } from "./bases.js";
export type { LibraryBase } from "./bases.js";
