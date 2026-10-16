/** A1-style addressing: column letters, cell and range references. */

export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

export interface CellAddr { r: number; c: number }
/** Zero-based, inclusive range. */
export interface Range { r1: number; c1: number; r2: number; c2: number }

export function colName(c: number): string {
  let s = '';
  let n = c + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function colIndex(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) {
    const k = ch.charCodeAt(0) - 64;
    if (k < 1 || k > 26) return -1;
    n = n * 26 + k;
  }
  return n - 1;
}

export function cellName(r: number, c: number, absR = false, absC = false): string {
  return `${absC ? '$' : ''}${colName(c)}${absR ? '$' : ''}${r + 1}`;
}

export function rangeName(rg: Range): string {
  if (rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1) return `${colName(rg.c1)}:${colName(rg.c2)}`;
  if (rg.c1 === 0 && rg.c2 >= MAX_COLS - 1) return `${rg.r1 + 1}:${rg.r2 + 1}`;
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return cellName(rg.r1, rg.c1);
  return `${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}`;
}

export function absRangeName(rg: Range): string {
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) return cellName(rg.r1, rg.c1, true, true);
  return `${cellName(rg.r1, rg.c1, true, true)}:${cellName(rg.r2, rg.c2, true, true)}`;
}

const CELL_RE = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/;

export function parseCell(s: string): CellAddr | null {
  const m = CELL_RE.exec(s.trim());
  if (!m) return null;
  const c = colIndex(m[1]);
  const r = parseInt(m[2], 10) - 1;
  if (c < 0 || c >= MAX_COLS || r < 0 || r >= MAX_ROWS) return null;
  return { r, c };
}

/** Parses "A1", "A1:B5", "A:C", "3:7" (no sheet prefix). */
export function parseRange(s: string): Range | null {
  const t = s.trim().replace(/\$/g, '');
  const parts = t.split(':');
  if (parts.length === 1) {
    const a = parseCell(parts[0]);
    return a ? { r1: a.r, c1: a.c, r2: a.r, c2: a.c } : null;
  }
  if (parts.length !== 2) return null;
  const [p, q] = parts;
  if (/^[A-Za-z]{1,3}$/.test(p) && /^[A-Za-z]{1,3}$/.test(q)) {
    const c1 = colIndex(p), c2 = colIndex(q);
    if (c1 < 0 || c2 < 0 || c1 >= MAX_COLS || c2 >= MAX_COLS) return null;
    return { r1: 0, c1: Math.min(c1, c2), r2: MAX_ROWS - 1, c2: Math.max(c1, c2) };
  }
  if (/^\d+$/.test(p) && /^\d+$/.test(q)) {
    const r1 = +p - 1, r2 = +q - 1;
    if (r1 < 0 || r2 < 0 || r1 >= MAX_ROWS || r2 >= MAX_ROWS) return null;
    return { r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: MAX_COLS - 1 };
  }
  const a = parseCell(p), b = parseCell(q);
  if (!a || !b) return null;
  return normRange({ r1: a.r, c1: a.c, r2: b.r, c2: b.c });
}

export function normRange(rg: Range): Range {
  return { r1: Math.min(rg.r1, rg.r2), c1: Math.min(rg.c1, rg.c2), r2: Math.max(rg.r1, rg.r2), c2: Math.max(rg.c1, rg.c2) };
}

export function inRange(rg: Range, r: number, c: number): boolean {
  return r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2;
}

export function intersects(a: Range, b: Range): boolean {
  return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
}

export function intersection(a: Range, b: Range): Range | null {
  if (!intersects(a, b)) return null;
  return { r1: Math.max(a.r1, b.r1), c1: Math.max(a.c1, b.c1), r2: Math.min(a.r2, b.r2), c2: Math.min(a.c2, b.c2) };
}

export function union(a: Range, b: Range): Range {
  return { r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1), r2: Math.max(a.r2, b.r2), c2: Math.max(a.c2, b.c2) };
}

export function rangeSize(rg: Range): { rows: number; cols: number; cells: number } {
  const rows = rg.r2 - rg.r1 + 1, cols = rg.c2 - rg.c1 + 1;
  return { rows, cols, cells: rows * cols };
}

export function sameRange(a: Range, b: Range): boolean {
  return a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;
}

/** Map key for a cell: fits in a double exactly (rows < 2^20, cols < 2^14). */
export function cellKey(r: number, c: number): number { return r * MAX_COLS + c; }
export function keyRow(k: number): number { return Math.floor(k / MAX_COLS); }
export function keyCol(k: number): number { return k % MAX_COLS; }

/** Quotes a sheet name for use in a reference when needed. */
export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !parseCell(name) && !/^R\d*C\d*$/i.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}

/** Splits "Sheet1!A1:B2" or "'My Sheet'!A1" into sheet name and range part. */
export function splitSheetRef(s: string): { sheet: string | null; ref: string } {
  const t = s.trim();
  const m = /^(?:'((?:[^']|'')+)'|([^!'"]+))!(.+)$/.exec(t);
  if (!m) return { sheet: null, ref: t };
  return { sheet: (m[1] ?? m[2]).replace(/''/g, "'"), ref: m[3] };
}

export function eachCell(rg: Range, fn: (r: number, c: number) => void | false): void {
  for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) if (fn(r, c) === false) return;
}
