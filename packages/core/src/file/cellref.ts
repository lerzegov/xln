// A1 cell addresses as stored in the `r` attribute (no `$`, no sheet).

export interface CellAddress {
  /** 1-based. */
  row: number;
  /** 1-based (A = 1). */
  col: number;
}

export function columnName(col: number): string {
  let s = "";
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  }
  return s;
}

export function formatCell(a: CellAddress): string {
  return columnName(a.col) + a.row;
}

/** Parses `B3` (also accepts `$B$3`); undefined if it is not a single-cell address. */
export function parseCell(ref: string): CellAddress | undefined {
  let i = 0;
  if (ref[i] === "$") i++;
  let col = 0;
  const start = i;
  while (i < ref.length) {
    const c = ref.charCodeAt(i) & ~0x20; // upper-case ASCII letters
    if (c < 65 || c > 90) break;
    col = col * 26 + (c - 64);
    i++;
  }
  if (i === start || i - start > 3) return undefined;
  if (ref[i] === "$") i++;
  const digits = ref.slice(i);
  if (digits.length === 0 || digits.length > 7) return undefined;
  for (const ch of digits) if (ch < "0" || ch > "9") return undefined;
  const row = Number(digits);
  if (row < 1) return undefined;
  return { row, col };
}
