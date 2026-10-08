// Excel's owner file: while a workbook is open, Excel keeps `~$<file name>` next to it
// (macOS and Windows, measured in probes T19/T20 and the file-level probe).
// The caller lists the folder; this module stays free of file-system access.

function baseName(path: string): string {
  const i = Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\"));
  return i < 0 ? path : path.slice(i + 1);
}

/** The owner-file name Excel creates for `fileName` (a bare name or a path). */
export function lockFileName(fileName: string): string {
  return "~$" + baseName(fileName);
}

/**
 * True if `siblingNames` (the names in the workbook's folder) contain its owner file.
 * Compared case-insensitively: the default file systems on both systems are.
 */
export function isLocked(fileName: string, siblingNames: readonly string[]): boolean {
  const lock = lockFileName(fileName).toLowerCase();
  return siblingNames.some((n) => baseName(n).toLowerCase() === lock);
}
