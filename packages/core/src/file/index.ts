// L0 file layer: read an .xlsx package into a WorkbookSnapshot.
export { readWorkbook } from "./workbook.js";
export {
  AFE_BLOB_NS,
  AFE_BLOB_ROOT,
  AFE_CODE_SHEET,
  AFE_LOCALE_SHEET,
  AFE_PRIMARY_MODULE,
  AFE_PROJECT_SCHEMAS,
  AFE_SETTINGS_KEY,
  afeExportedName,
  afeModuleName,
  findForeignModuleStores,
  isAfeXml,
  parseAfeBlob,
} from "./afe.js";
export { fromBase64, toBase64 } from "./base64.js";
export { XlsxError } from "./package.js";
export { XmlError } from "./xml.js";
export { readZipLayout, rawZipRecords, replaceZipEntries, rewriteZip, crc32, type ZipEntry, type ZipLayout, type ZipEdits } from "./zip.js";
export { lockFileName, isLocked } from "./lock.js";
export { formulaTextAt, type ReferenceShifter } from "./shared.js";
export { parseCell, formatCell, columnName, type CellAddress } from "./cellref.js";
export type * from "./types.js";
