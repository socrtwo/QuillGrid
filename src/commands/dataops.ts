/**
 * Data-tab logic shared by the Data commands: sorting (multi-key, by color/icon, custom lists), AutoFilter and
 * table filter evaluation, Advanced Filter criteria ranges, Text to Columns, Flash Fill program inference,
 * Remove Duplicates, Subtotals, outlines, Consolidate and the Transform steps. Everything here is free of UI
 * so it can be unit tested; mutating helpers must be called inside book.transact().
 */
import type { Sheet } from '../model/workbook';
import type { Cell, CellStyle, FilterCriteria, Scalar, AutoFilter, TableDef } from '../model/types';
import { CellError } from '../model/types';
import { cellKey, keyRow, keyCol, colName, type Range } from '../model/address';
import { formatValue } from '../model/numfmt';
import { serialToParts, partsToSerial, jsDateToSerial, MONTHS } from '../model/dates';
import { shiftFormula } from '../calc/parser';
import { criteria as criteriaFn, wildcard, parseNumeric } from '../calc/values';
import { convertField } from '../io/csv';

// --------------------------------------------------------------------------------------------
// Small helpers

export const isBlank = (v: Scalar | undefined): boolean => v == null || v === '';

/** Displayed text of a cell (number format applied), as the filter list and Remove Duplicates use. */
export function displayText(sheet: Sheet, r: number, c: number): string {
  const cell = sheet.get(r, c);
  if (!cell || cell.v == null) return '';
  return formatValue(cell.v, sheet.styleOf(r, c).numFmt, { date1904: sheet.book.calc.date1904 }).text;
}

const normColor = (c: string | undefined | null): string => (c || '').toLowerCase();

export function rowsOf(r1: number, r2: number): number[] { const out: number[] = []; for (let r = r1; r <= r2; r++) out.push(r); return out; }

// --------------------------------------------------------------------------------------------
// Sorting

export interface SortKey {
  /** Absolute column (or row, when sorting left to right) to sort by. */
  index: number;
  on?: 'values' | 'fill' | 'font' | 'icon';
  desc?: boolean;
  /** Custom list order (Sort ▸ Order ▸ Custom List). */
  list?: string[];
  /** Cell/font color for 'fill' / 'font' keys ('' = no color). */
  color?: string;
  /** Icon for 'icon' keys. */
  icon?: { set: string; index: number };
  /** Put the matching color/icon at the bottom instead of the top. */
  bottom?: boolean;
}

export interface SortOptions {
  header?: boolean;
  /** Sort left to right (columns) instead of top to bottom. */
  byCols?: boolean;
  caseSensitive?: boolean;
  /** Leave hidden rows (filtered out) in place. */
  skipHidden?: boolean;
  iconAt?: (r: number, c: number) => { set: string; index: number } | null;
}

const typeRank = (v: Scalar): number => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);

/** Text comparison used by Sort: case-insensitive, accent-aware; case-sensitive puts lowercase first on ties. */
export function compareText(a: string, b: string, caseSensitive = false): number {
  const x = a.localeCompare(b, 'en', { sensitivity: 'accent' });
  if (x || !caseSensitive) return x;
  if (a === b) return 0;
  for (let i = 0; i < Math.min(a.length, b.length); i++) {
    if (a[i] === b[i]) continue;
    const la = a[i] === a[i].toLowerCase(), lb = b[i] === b[i].toLowerCase();
    if (la !== lb) return la ? -1 : 1;
  }
  return a.length - b.length;
}

/**
 * Compares two values the way a spreadsheet sort does: numbers < text < logical < errors in ascending order (the
 * type order reverses in descending order) and blanks always last.
 */
export function compareForSort(a: Scalar, b: Scalar, desc = false, caseSensitive = false, list?: string[]): number {
  const ba = isBlank(a), bb = isBlank(b);
  if (ba || bb) return ba && bb ? 0 : ba ? 1 : -1;
  const sgn = desc ? -1 : 1;
  if (list && list.length) {
    const low = list.map((x) => x.toLowerCase());
    const ia = low.indexOf(String(a).toLowerCase()), ib = low.indexOf(String(b).toLowerCase());
    if (ia >= 0 || ib >= 0) {
      if (ia < 0) return 1;
      if (ib < 0) return -1;
      return (ia - ib) * sgn;
    }
  }
  const ra = typeRank(a), rb = typeRank(b);
  if (ra !== rb) return (ra - rb) * sgn;
  let x = 0;
  if (typeof a === 'number') x = a - (b as number);
  else if (typeof a === 'string') x = compareText(a, b as string, caseSensitive);
  else if (typeof a === 'boolean') x = a === b ? 0 : a ? 1 : -1;
  else if (a instanceof CellError) x = 0;
  return Math.sign(x) * sgn;
}

/** Throws when a sort cannot proceed because of merged cells or multi-row arrays in the area. */
export function checkSortable(sheet: Sheet, area: Range, byCols = false): void {
  for (const m of sheet.merges) {
    if (m.r2 < area.r1 || m.r1 > area.r2 || m.c2 < area.c1 || m.c1 > area.c2) continue;
    const spansSortAxis = byCols ? m.c1 !== m.c2 : m.r1 !== m.r2;
    if (spansSortAxis) throw new Error('To sort, all the merged cells in the range need to be the same size (unmerge them first).');
  }
  sheet.forEachIn(area, (_r, _c, cell) => {
    if (cell.arr && (byCols ? cell.arr.c1 !== cell.arr.c2 : cell.arr.r1 !== cell.arr.r2)) throw new Error("You can't change part of an array. Sorting would split an array formula.");
  });
}

/**
 * Sorts a range in place. Whole cells move (values, formulas with relative references adjusted by the row/column
 * offset, styles, hyperlinks, comments and single-row merges). Returns the new order of source rows/columns.
 */
export function sortRange(sheet: Sheet, rg: Range, keys: SortKey[], opts: SortOptions = {}): number[] {
  const byCols = !!opts.byCols;
  const first = byCols ? rg.c1 + (opts.header ? 1 : 0) : rg.r1 + (opts.header ? 1 : 0);
  const last = byCols ? rg.c2 : rg.r2;
  if (last <= first) return [];
  const area: Range = byCols ? { ...rg, c1: first } : { ...rg, r1: first };
  checkSortable(sheet, area, byCols);
  let lines = rowsOf(first, last);
  if (opts.skipHidden && !byCols) lines = lines.filter((r) => !sheet.hiddenRows.has(r));
  if (opts.skipHidden && byCols) lines = lines.filter((c) => !sheet.hiddenCols.has(c));
  const at = (line: number, key: SortKey) => (byCols ? { r: key.index, c: line } : { r: line, c: key.index });
  const keyVal = (line: number, key: SortKey): number | Scalar => {
    const { r, c } = at(line, key);
    const on = key.on || 'values';
    if (on === 'values') return sheet.value(r, c);
    let hit = false;
    if (on === 'fill') { const st = sheet.styleOf(r, c); hit = normColor(st.fill) === normColor(key.color); }
    else if (on === 'font') { const st = sheet.styleOf(r, c); hit = normColor(st.color) === normColor(key.color); }
    else { const ic = opts.iconAt?.(r, c); hit = !!ic && !!key.icon && ic.set === key.icon.set && ic.index === key.icon.index; }
    return hit === !key.bottom ? 0 : 1;
  };
  const decorated = lines.map((line, i) => ({ line, i, k: keys.map((key) => keyVal(line, key)) }));
  decorated.sort((a, b) => {
    for (let j = 0; j < keys.length; j++) {
      const key = keys[j];
      const x = (key.on || 'values') === 'values' ? compareForSort(a.k[j] as Scalar, b.k[j] as Scalar, key.desc, opts.caseSensitive, key.list) : (a.k[j] as number) - (b.k[j] as number);
      if (x) return x;
    }
    return a.i - b.i;
  });
  const pairs: [number, number][] = decorated.map((d, i) => [lines[i], d.line]);
  if (pairs.every(([d, s]) => d === s)) return decorated.map((d) => d.line);
  if (byCols) rewriteCols(sheet, rg.r1, rg.r2, lines, pairs);
  else rewriteRows(sheet, rg.c1, rg.c2, lines, pairs);
  return decorated.map((d) => d.line);
}

/**
 * Rewrites rows of columns c1..c2: every row in `rows` is cleared, then each [dest, src] pair copies the source
 * row's cells (as they were before) to dest, adjusting relative references by the row offset. Comments and
 * single-row merges move with their cells. Records the sheet for undo.
 */
export function rewriteRows(sheet: Sheet, c1: number, c2: number, rows: number[], pairs: [number, number][]): void {
  rewriteLines(sheet, false, c1, c2, rows, pairs);
}

export function rewriteCols(sheet: Sheet, r1: number, r2: number, cols: number[], pairs: [number, number][]): void {
  rewriteLines(sheet, true, r1, r2, cols, pairs);
}

function rewriteLines(sheet: Sheet, byCols: boolean, a: number, b: number, lines: number[], pairs: [number, number][]): void {
  const book = sheet.book;
  book.recordSheet(sheet);
  const pos = (line: number, x: number) => (byCols ? { r: x, c: line } : { r: line, c: x });
  const snap = new Map<number, { cells: Map<number, Cell>; comments: Map<number, any> }>();
  for (const line of lines) {
    const cells = new Map<number, Cell>(), comments = new Map<number, any>();
    for (let x = a; x <= b; x++) {
      const { r, c } = pos(line, x);
      const cell = sheet.get(r, c);
      if (cell) cells.set(x, { ...cell });
      const cm = sheet.comments.get(cellKey(r, c));
      if (cm) comments.set(x, cm);
    }
    snap.set(line, { cells, comments });
  }
  const lineSet = new Set(lines);
  const movedMerges: { line: number; m: Range }[] = [];
  sheet.merges = sheet.merges.filter((m) => {
    const line = byCols ? m.c1 : m.r1;
    const single = byCols ? m.c1 === m.c2 : m.r1 === m.r2;
    const inside = byCols ? m.r1 >= a && m.r2 <= b : m.c1 >= a && m.c2 <= b;
    if (single && inside && lineSet.has(line)) { movedMerges.push({ line, m }); return false; }
    return true;
  });
  for (const line of lines) for (let x = a; x <= b; x++) {
    const { r, c } = pos(line, x);
    const cur = sheet.get(r, c);
    if (cur) sheet.set(r, c, undefined);
    sheet.comments.delete(cellKey(r, c));
  }
  for (const [dest, src] of pairs) {
    const s = snap.get(src);
    if (!s) continue;
    const d = dest - src;
    for (const [x, cell0] of s.cells) {
      const cell: Cell = { ...cell0 };
      if (cell.sp != null && cell.f == null) { delete cell.sp; delete cell.v; if (!cell.s) continue; }
      delete cell.sp;
      if (cell.f != null && d) cell.f = byCols ? shiftFormula(cell.f, 0, d) : shiftFormula(cell.f, d, 0);
      if (cell.arr && d) cell.arr = byCols ? { ...cell.arr, c1: cell.arr.c1 + d, c2: cell.arr.c2 + d } : { ...cell.arr, r1: cell.arr.r1 + d, r2: cell.arr.r2 + d };
      const { r, c } = pos(dest, x);
      sheet.set(r, c, cell);
    }
    for (const [x, cm] of s.comments) { const { r, c } = pos(dest, x); sheet.comments.set(cellKey(r, c), cm); }
    for (const mm of movedMerges) if (mm.line === src) sheet.merges.push(byCols ? { ...mm.m, c1: dest, c2: dest } : { ...mm.m, r1: dest, r2: dest });
  }
  book.pending.structural = true;
}

// --------------------------------------------------------------------------------------------
// Filters

export type DatePeriod =
  | 'today' | 'yesterday' | 'tomorrow' | 'thisWeek' | 'lastWeek' | 'nextWeek' | 'thisMonth' | 'lastMonth' | 'nextMonth'
  | 'thisQuarter' | 'lastQuarter' | 'nextQuarter' | 'thisYear' | 'lastYear' | 'nextYear' | 'yearToDate'
  | 'Q1' | 'Q2' | 'Q3' | 'Q4' | 'M1' | 'M2' | 'M3' | 'M4' | 'M5' | 'M6' | 'M7' | 'M8' | 'M9' | 'M10' | 'M11' | 'M12';

export const DATE_PERIODS: { id: DatePeriod; label: string }[] = [
  { id: 'tomorrow', label: 'Tomorrow' }, { id: 'today', label: 'Today' }, { id: 'yesterday', label: 'Yesterday' },
  { id: 'nextWeek', label: 'Next Week' }, { id: 'thisWeek', label: 'This Week' }, { id: 'lastWeek', label: 'Last Week' },
  { id: 'nextMonth', label: 'Next Month' }, { id: 'thisMonth', label: 'This Month' }, { id: 'lastMonth', label: 'Last Month' },
  { id: 'nextQuarter', label: 'Next Quarter' }, { id: 'thisQuarter', label: 'This Quarter' }, { id: 'lastQuarter', label: 'Last Quarter' },
  { id: 'nextYear', label: 'Next Year' }, { id: 'thisYear', label: 'This Year' }, { id: 'lastYear', label: 'Last Year' },
  { id: 'yearToDate', label: 'Year to Date' },
];

/** Half-open serial interval [from, to) for a date period relative to `today` (a serial). */
export function periodRange(p: DatePeriod, today: number, date1904 = false): [number, number] {
  const t = Math.floor(today);
  const { y, m, dow } = serialToParts(t, date1904);
  const ser = (yy: number, mm: number, dd: number) => {
    while (mm > 12) { mm -= 12; yy++; }
    while (mm < 1) { mm += 12; yy--; }
    return partsToSerial(yy, mm, dd, date1904);
  };
  const q = Math.floor((m - 1) / 3);
  switch (p) {
    case 'today': return [t, t + 1];
    case 'yesterday': return [t - 1, t];
    case 'tomorrow': return [t + 1, t + 2];
    case 'thisWeek': return [t - dow, t - dow + 7];
    case 'lastWeek': return [t - dow - 7, t - dow];
    case 'nextWeek': return [t - dow + 7, t - dow + 14];
    case 'thisMonth': return [ser(y, m, 1), ser(y, m + 1, 1)];
    case 'lastMonth': return [ser(y, m - 1, 1), ser(y, m, 1)];
    case 'nextMonth': return [ser(y, m + 1, 1), ser(y, m + 2, 1)];
    case 'thisQuarter': return [ser(y, q * 3 + 1, 1), ser(y, q * 3 + 4, 1)];
    case 'lastQuarter': return [ser(y, q * 3 - 2, 1), ser(y, q * 3 + 1, 1)];
    case 'nextQuarter': return [ser(y, q * 3 + 4, 1), ser(y, q * 3 + 7, 1)];
    case 'thisYear': return [ser(y, 1, 1), ser(y + 1, 1, 1)];
    case 'lastYear': return [ser(y - 1, 1, 1), ser(y, 1, 1)];
    case 'nextYear': return [ser(y + 1, 1, 1), ser(y + 2, 1, 1)];
    case 'yearToDate': return [ser(y, 1, 1), t + 1];
    default: return [NaN, NaN];
  }
}

function inPeriod(p: string, v: number, today: number, date1904: boolean): boolean {
  const d = Math.floor(v);
  if (/^Q[1-4]$/.test(p)) { const { m } = serialToParts(d, date1904); return Math.floor((m - 1) / 3) + 1 === +p[1]; }
  if (/^M\d+$/.test(p)) return serialToParts(d, date1904).m === +p.slice(1);
  const [a, b] = periodRange(p as DatePeriod, today, date1904);
  return d >= a && d < b;
}

/** Per-column statistics that some criteria need (Top 10, above/below average). */
export interface ColumnStats { avg: number; sorted: number[] }

export function columnStats(values: Scalar[]): ColumnStats {
  const nums = values.filter((v): v is number => typeof v === 'number');
  const avg = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : NaN;
  return { avg, sorted: [...nums].sort((a, b) => b - a) };
}

export interface CriteriaContext { stats?: ColumnStats; today?: number; date1904?: boolean; style?: CellStyle }

/** Tests one custom condition (Custom AutoFilter). Ops: eq ne gt ge lt le begins notBegins ends notEnds contains notContains period blank notBlank. */
export function customMatch(op: string, value: string, v: Scalar, text: string, ctx: CriteriaContext = {}): boolean {
  const low = text.toLowerCase();
  const val = value ?? '';
  const n = typeof v === 'number' ? v : null;
  const vn = parseNumeric(val);
  const asDate = vn == null ? dateSerial(val) : null;
  const target = vn ?? asDate;
  switch (op) {
    case 'blank': return isBlank(v);
    case 'notBlank': return !isBlank(v);
    case 'eq':
    case 'ne': {
      let hit: boolean;
      if (val === '') hit = isBlank(v);
      else if (target != null && n != null) hit = Math.abs(n - target) < 1e-9 || low === val.toLowerCase();
      else if (/[*?]/.test(val)) hit = wildcard(val).test(text);
      else hit = low === val.toLowerCase();
      return op === 'eq' ? hit : !hit;
    }
    case 'gt': case 'ge': case 'lt': case 'le': {
      let x: number;
      if (target != null) { if (n == null) return false; x = n - target; }
      else { if (isBlank(v)) return false; x = compareText(text, val); }
      return op === 'gt' ? x > 0 : op === 'ge' ? x >= 0 : op === 'lt' ? x < 0 : x <= 0;
    }
    case 'begins': return wildcard(val + '*').test(text);
    case 'notBegins': return !wildcard(val + '*').test(text);
    case 'ends': return wildcard('*' + val).test(text);
    case 'notEnds': return !wildcard('*' + val).test(text);
    case 'contains': return wildcard('*' + val + '*').test(text);
    case 'notContains': return !wildcard('*' + val + '*').test(text);
    case 'period': return n != null && inPeriod(val, n, ctx.today ?? jsDateToSerial(new Date()), !!ctx.date1904);
  }
  return true;
}

function dateSerial(s: string): number | null {
  if (!/\d/.test(s) || !/[-/.]|[a-z]{3}/i.test(s)) return null;
  const c = convertField(s.trim());
  return typeof c.v === 'number' && c.fmt ? c.v : null;
}

/** Whether a cell passes an AutoFilter column criteria. */
export function criteriaPasses(crit: FilterCriteria, v: Scalar, text: string, ctx: CriteriaContext = {}): boolean {
  if (crit.values) {
    if (isBlank(v)) { if (!crit.blanks) return false; }
    else if (!crit.values.some((x) => x.toLowerCase() === text.toLowerCase())) return false;
  } else if (crit.blanks === false && isBlank(v)) return false;
  if (crit.custom && crit.custom.length) {
    const res = crit.custom.map((cc) => customMatch(cc.op, cc.value, v, text, ctx));
    if (crit.and ? !res.every(Boolean) : !res.some(Boolean)) return false;
  }
  if (crit.top10) {
    if (typeof v !== 'number' || !ctx.stats) return false;
    const s = ctx.stats.sorted;
    const k = Math.max(1, crit.top10.percent ? Math.floor((s.length * crit.top10.n) / 100) : Math.floor(crit.top10.n));
    if (!s.length) return false;
    if (crit.top10.top) return v >= s[Math.min(k, s.length) - 1];
    return v <= s[Math.max(0, s.length - k)];
  }
  if (crit.dynamic) {
    if (typeof v !== 'number' || !ctx.stats || !isFinite(ctx.stats.avg)) return false;
    return crit.dynamic === 'aboveAverage' ? v > ctx.stats.avg : v < ctx.stats.avg;
  }
  if (crit.color) {
    const st = ctx.style || {};
    if (crit.color.fill !== undefined && normColor(st.fill) !== normColor(crit.color.fill)) return false;
    if (crit.color.font !== undefined && normColor(st.color) !== normColor(crit.color.font)) return false;
  }
  return true;
}

/** A filtered area: data rows (no header, no totals) and per-column criteria keyed by offset from c1. */
export interface FilterArea { r1: number; r2: number; c1: number; c2: number; columns: Record<number, FilterCriteria> }

export function tableDataRange(t: TableDef): Range {
  return { r1: t.range.r1 + (t.opts.header ? 1 : 0), c1: t.range.c1, r2: t.range.r2 - (t.opts.totals ? 1 : 0), c2: t.range.c2 };
}

/** All filtered areas of a sheet: the AutoFilter range and every table with criteria. */
export function filterAreas(sheet: Sheet): FilterArea[] {
  const out: FilterArea[] = [];
  const f = sheet.filter;
  if (f) out.push({ r1: f.range.r1 + 1, r2: f.range.r2, c1: f.range.c1, c2: f.range.c2, columns: f.columns || {} });
  for (const t of sheet.tables) if (t.filter) { const d = tableDataRange(t); out.push({ ...d, columns: t.filter.columns || {} }); }
  return out;
}

/**
 * Evaluates whether each data row of an area passes; `skipCol` ignores one column's criteria (used to list the
 * values a column's checkbox list offers under the other columns' filters).
 */
export function areaPasses(sheet: Sheet, area: FilterArea, opts: { today?: number; skipCol?: number } = {}): Map<number, boolean> {
  const out = new Map<number, boolean>();
  const cols = Object.keys(area.columns).map(Number).filter((k) => k !== opts.skipCol && area.columns[k]);
  const ctxs = new Map<number, CriteriaContext>();
  for (const k of cols) {
    const crit = area.columns[k];
    const ctx: CriteriaContext = { today: opts.today, date1904: sheet.book.calc.date1904 };
    if (crit.top10 || crit.dynamic) { const vals: Scalar[] = []; for (let r = area.r1; r <= area.r2; r++) vals.push(sheet.value(r, area.c1 + k)); ctx.stats = columnStats(vals); }
    ctxs.set(k, ctx);
  }
  for (let r = area.r1; r <= area.r2; r++) {
    let ok = true;
    for (const k of cols) {
      const c = area.c1 + k;
      const crit = area.columns[k];
      const ctx = ctxs.get(k)!;
      if (crit.color) ctx.style = sheet.styleOf(r, c);
      if (!criteriaPasses(crit, sheet.value(r, c), displayText(sheet, r, c), ctx)) { ok = false; break; }
    }
    out.set(r, ok);
  }
  return out;
}

/**
 * Recomputes hidden rows for every filtered area on the sheet (AutoFilter and tables). Rows that fail any area's
 * criteria are hidden; other rows of those areas are shown. Records the sheet for undo. Returns counts.
 */
export function applyFilters(sheet: Sheet, opts: { today?: number } = {}): { shown: number; total: number } {
  sheet.book.recordSheet(sheet);
  const fails = new Map<number, boolean>();
  let total = 0;
  for (const area of filterAreas(sheet)) {
    const res = areaPasses(sheet, area, opts);
    for (const [r, ok] of res) { fails.set(r, (fails.get(r) || false) || !ok); }
  }
  let shown = 0;
  for (const [r, fail] of fails) {
    total++;
    if (fail) sheet.hiddenRows.add(r); else { sheet.hiddenRows.delete(r); shown++; }
  }
  sheet.book.pending.structural = true;
  return { shown, total };
}

/** Grows an AutoFilter range down over contiguous data rows that were added under it. */
export function extendFilterRange(sheet: Sheet, rg: Range): Range {
  let r2 = rg.r2;
  const hasData = (r: number) => { for (let c = rg.c1; c <= rg.c2; c++) { const x = sheet.get(r, c); if (x && (x.v != null && x.v !== '' || x.f != null)) return true; } return false; };
  while (r2 < 1048575 && hasData(r2 + 1)) r2++;
  return { ...rg, r2 };
}

/** Distinct displayed values of a column of an area (rows passing the other columns' criteria), sorted. */
export function distinctColumnValues(sheet: Sheet, area: FilterArea, col: number): { values: string[]; blanks: boolean; fills: string[]; fonts: string[]; kind: 'number' | 'date' | 'text' } {
  const pass = areaPasses(sheet, area, { skipCol: col });
  const seen = new Map<string, Scalar>();
  let blanks = false, nums = 0, dates = 0, texts = 0;
  const fills = new Set<string>(), fonts = new Set<string>();
  const c = area.c1 + col;
  for (let r = area.r1; r <= area.r2; r++) {
    if (!pass.get(r)) continue;
    const v = sheet.value(r, c);
    const st = sheet.styleOf(r, c);
    fills.add(normColor(st.fill)); fonts.add(normColor(st.color));
    if (isBlank(v)) { blanks = true; continue; }
    const t = displayText(sheet, r, c);
    const key = t.toLowerCase();
    if (!seen.has(key)) seen.set(key, v);
    if (typeof v === 'number') { const fmt = sheet.styleOf(r, c).numFmt; if (fmt && /[dmy]/i.test(fmt.replace(/"[^"]*"|\[[^\]]*\]/g, '')) && !/^[#0,.%]*$/.test(fmt)) dates++; else nums++; } else texts++;
  }
  const entries = [...seen.entries()].map(([k, v]) => ({ k, v, t: '' }));
  const textOf = new Map<string, string>();
  for (let r = area.r1; r <= area.r2; r++) { const t = displayText(sheet, r, c); if (t && !textOf.has(t.toLowerCase())) textOf.set(t.toLowerCase(), t); }
  for (const e of entries) e.t = textOf.get(e.k) || e.k;
  entries.sort((a, b) => compareForSort(a.v, b.v) || compareText(a.t, b.t));
  const kind = dates > nums && dates >= texts ? 'date' : nums >= texts && nums > 0 ? 'number' : 'text';
  return { values: entries.map((e) => e.t), blanks, fills: [...fills], fonts: [...fonts], kind };
}

/** Distinct values in the column around (r, c) for Pick From Drop-down List. */
export function pickListValues(sheet: Sheet, r: number, c: number): string[] {
  const vals = new Map<string, string>();
  const scan = (dir: 1 | -1) => {
    for (let rr = r + dir; rr >= 0 && rr < 1048576; rr += dir) {
      const v = sheet.value(rr, c);
      if (isBlank(v)) break;
      if (typeof v === 'string') { const k = v.toLowerCase(); if (!vals.has(k)) vals.set(k, v); }
    }
  };
  scan(-1); scan(1);
  return [...vals.values()].sort((a, b) => compareText(a, b));
}

// --------------------------------------------------------------------------------------------
// Advanced Filter

export interface CriteriaTest { (r: number): boolean }

/**
 * Builds the row test of an Advanced Filter criteria range: the first row holds field names; each later row is a
 * set of AND conditions, rows are OR'ed. A condition without a matching field name and holding a formula is a
 * computed criterion evaluated relative to the first data row. `evalAt(f, r, c, baseR, baseC)` evaluates formulas.
 */
export function advancedCriteria(sheet: Sheet, list: Range, crit: Range, critSheet: Sheet, evalAt: (f: string, r: number, c: number, baseR: number, baseC: number) => Scalar): CriteriaTest {
  const headers = new Map<string, number>();
  for (let c = list.c1; c <= list.c2; c++) { const t = displayText(sheet, list.r1, c).trim().toLowerCase(); if (t && !headers.has(t)) headers.set(t, c); }
  const rows: ((r: number) => boolean)[][] = [];
  for (let cr = crit.r1 + 1; cr <= crit.r2; cr++) {
    const conds: ((r: number) => boolean)[] = [];
    for (let cc = crit.c1; cc <= crit.c2; cc++) {
      const cell = critSheet.get(cr, cc);
      if (!cell || (cell.f == null && isBlank(cell.v ?? null))) continue;
      const name = displayText(critSheet, crit.r1, cc).trim().toLowerCase();
      const col = headers.get(name);
      if (col == null) {
        if (cell.f != null) {
          const f = cell.f;
          conds.push((r) => { const x = evalAt(f, r, list.c1, list.r1 + 1, list.c1); return x === true || (typeof x === 'number' && x !== 0); });
        } else conds.push(() => false);
        continue;
      }
      const raw = cell.f != null ? cell.v ?? null : cell.v ?? null;
      conds.push(conditionTest(sheet, col, raw));
    }
    rows.push(conds);
  }
  if (!rows.length) return () => true;
  return (r) => rows.some((conds) => conds.every((t) => t(r)));
}

function conditionTest(sheet: Sheet, col: number, raw: Scalar): (r: number) => boolean {
  if (typeof raw === 'number' || typeof raw === 'boolean') { const t = criteriaFn(raw); return (r) => t(sheet.value(r, col)); }
  const s = String(raw ?? '');
  if (/^(<=|>=|<>|<|>)/.test(s)) { const t = criteriaFn(s); return (r) => t(sheet.value(r, col)); }
  if (s.startsWith('=')) {
    const rhs = s.slice(1);
    if (rhs === '') return (r) => isBlank(sheet.value(r, col));
    const t = criteriaFn('=' + rhs);
    return (r) => t(sheet.value(r, col)) || displayText(sheet, r, col).toLowerCase() === rhs.toLowerCase();
  }
  // Plain text: "begins with" (wildcards allowed).
  const re = wildcard(s + '*');
  const n = parseNumeric(s);
  return (r) => { const v = sheet.value(r, col); if (n != null && typeof v === 'number') return v === n; return re.test(displayText(sheet, r, col)); };
}

// --------------------------------------------------------------------------------------------
// Text to Columns

export interface SplitOptions {
  mode: 'delimited' | 'fixed';
  delimiters?: string[];
  consecutive?: boolean;
  qualifier?: string;
  breaks?: number[];
}

/** Splits one line into fields. */
export function splitLine(s: string, o: SplitOptions): string[] {
  if (o.mode === 'fixed') {
    const br = [...new Set((o.breaks || []).filter((b) => b > 0))].sort((a, b) => a - b);
    const out: string[] = [];
    let last = 0;
    for (const b of br) { out.push(s.slice(last, b)); last = b; }
    out.push(s.slice(last));
    return out;
  }
  const delims = new Set(o.delimiters || []);
  const q = o.qualifier || '';
  const out: string[] = [];
  let field = '', i = 0, fieldStart = true, prevDelim = false;
  while (i < s.length) {
    const ch = s[i];
    if (q && ch === q && fieldStart) {
      i++;
      for (;;) {
        if (i >= s.length) break;
        if (s[i] === q) { if (s[i + 1] === q) { field += q; i += 2; continue; } i++; break; }
        field += s[i++];
      }
      fieldStart = false; prevDelim = false;
      continue;
    }
    if (delims.has(ch)) {
      if (o.consecutive && prevDelim) { i++; continue; }
      out.push(field); field = ''; fieldStart = true; prevDelim = true; i++;
      continue;
    }
    field += ch; fieldStart = false; prevDelim = false; i++;
  }
  out.push(field);
  return out;
}

/** Suggested fixed-width break positions: columns where every line has a space followed by text. */
export function suggestBreaks(lines: string[]): number[] {
  const ls = lines.filter((l) => l.trim());
  if (!ls.length) return [];
  const max = Math.max(...ls.map((l) => l.length));
  const out: number[] = [];
  for (let p = 1; p < max; p++) {
    if (ls.every((l) => p >= l.length || (l[p - 1] === ' ' && l[p] !== ' ')) && ls.some((l) => p < l.length)) out.push(p);
  }
  return out;
}

export type ColumnFormat = 'general' | 'text' | 'skip' | 'MDY' | 'DMY' | 'YMD' | 'MYD' | 'DYM' | 'YDM';

/** Converts a split field per its column data format. */
export function convertSplitField(s: string, fmt: ColumnFormat, date1904 = false): { v: Scalar; fmt?: string } {
  if (fmt === 'text') return { v: s === '' ? null : s, fmt: '@' };
  const t = s.trim();
  if (t === '') return { v: null };
  if (fmt === 'general') {
    const c = convertField(t);
    if (c.f != null) return { v: s };
    return { v: c.v, fmt: c.fmt };
  }
  if (fmt === 'skip') return { v: null };
  const parts = t.split(/[\s/\-.,]+/).filter(Boolean);
  if (parts.length < 3) return { v: s };
  const order = fmt.split('');
  let y = NaN, m = NaN, d = NaN;
  order.forEach((k, i) => {
    const p = parts[i];
    let n = +p;
    if (k === 'M' && !Number.isFinite(n)) { const mi = MONTHS.findIndex((x) => x.toLowerCase().startsWith(p.toLowerCase().slice(0, 3))); n = mi >= 0 ? mi + 1 : NaN; }
    if (k === 'Y') y = n < 100 ? (n < 30 ? 2000 + n : 1900 + n) : n; else if (k === 'M') m = n; else d = n;
  });
  if (!(m >= 1 && m <= 12 && d >= 1 && d <= 31 && y >= 1900)) return { v: s };
  let serial = partsToSerial(y, m, d, date1904);
  const tm = /(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(t);
  if (tm) serial += (+tm[1] * 3600 + +tm[2] * 60 + +(tm[3] || 0)) / 86400;
  return { v: serial, fmt: tm ? 'm/d/yyyy h:mm' : 'm/d/yyyy' };
}

// --------------------------------------------------------------------------------------------
// Flash Fill: programs are concatenations of literals and substring expressions over the source columns.

type CaseT = 'none' | 'upper' | 'lower' | 'proper';
type TokT = 'word' | 'alpha' | 'digit';
export type FFExpr =
  | { k: 'all'; col: number; cs: CaseT }
  | { k: 'tok'; col: number; t: TokT; i: number; cs: CaseT }
  | { k: 'span'; col: number; t: TokT; i: number; j: number; cs: CaseT }
  | { k: 'init'; col: number; t: TokT; i: number; cs: CaseT }
  | { k: 'fld'; col: number; d: string; i: number; j: number; cs: CaseT }
  | { k: 'pre' | 'suf'; col: number; n: number; cs: CaseT };
export type FFPart = { lit: string } | { e: FFExpr };
export interface FFProgram { parts: FFPart[]; cost: number }

const TOK_RE: Record<TokT, RegExp> = { word: /[\p{L}\p{N}]+(?:['’][\p{L}]+)*/gu, alpha: /\p{L}+/gu, digit: /\p{N}+/gu };

function tokenSpans(s: string, t: TokT): [number, number][] {
  const re = new RegExp(TOK_RE[t].source, 'gu');
  const out: [number, number][] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(s))) out.push([m.index, m.index + m[0].length]);
  return out;
}

const at = (i: number, n: number) => (i >= 0 ? (i < n ? i : -1) : n + i >= 0 ? n + i : -1);

function applyCase(s: string, cs: CaseT): string {
  switch (cs) {
    case 'upper': return s.toUpperCase();
    case 'lower': return s.toLowerCase();
    case 'proper': return s.toLowerCase().replace(/(^|[^\p{L}'’])(\p{L})/gu, (_m, a, b) => a + b.toUpperCase());
    default: return s;
  }
}

export function evalExpr(e: FFExpr, src: string[]): string | null {
  const s = src[e.col] ?? '';
  let out: string | null = null;
  switch (e.k) {
    case 'all': out = s; break;
    case 'tok': case 'init': {
      const sp = tokenSpans(s, e.t); const i = at(e.i, sp.length);
      if (i < 0) return null;
      out = e.k === 'tok' ? s.slice(sp[i][0], sp[i][1]) : s[sp[i][0]];
      break;
    }
    case 'span': {
      const sp = tokenSpans(s, e.t); const i = at(e.i, sp.length), j = at(e.j, sp.length);
      if (i < 0 || j < 0 || j <= i) return null;
      out = s.slice(sp[i][0], sp[j][1]);
      break;
    }
    case 'fld': {
      const f = s.split(e.d); if (f.length < 2) return null;
      const i = at(e.i, f.length), j = at(e.j, f.length);
      if (i < 0 || j < 0 || j < i) return null;
      out = f.slice(i, j + 1).join(e.d).trim();
      break;
    }
    case 'pre': if (s.length < e.n) return null; out = s.slice(0, e.n); break;
    case 'suf': if (s.length < e.n) return null; out = s.slice(s.length - e.n); break;
  }
  if (out == null || out === '') return null;
  return applyCase(out, e.cs);
}

function exprCost(e: FFExpr): number {
  const cs = e.cs === 'none' ? 0 : 0.15;
  const idx = (i: number) => (i >= 0 ? 0.02 * i : 0.05 + 0.02 * (-i - 1));
  const tt = (t: TokT) => (t === 'word' ? 0 : 0.03);
  switch (e.k) {
    case 'all': return 1 + cs;
    case 'tok': return 1.1 + idx(e.i) + tt(e.t) + cs;
    case 'init': return 1.15 + idx(e.i) + tt(e.t) + cs;
    case 'span': return 1.3 + idx(e.i) + idx(e.j) + tt(e.t) + cs;
    case 'fld': return 1.2 + idx(e.i) + 0.05 * (e.j - e.i) + cs;
    default: return 1.7 + 0.01 * e.n + cs;
  }
}

const LIT_START = 0.6, LIT_CHAR = 0.6;

function candidateExprs(rows: string[][], cols: number[]): FFExpr[] {
  const out: FFExpr[] = [];
  const cases: CaseT[] = ['none', 'upper', 'lower', 'proper'];
  for (const col of cols) {
    const strs = rows.map((r) => r[col] ?? '');
    const base: Omit<FFExpr, 'cs'>[] = [{ k: 'all', col } as any];
    for (const t of ['word', 'alpha', 'digit'] as TokT[]) {
      const m = Math.min(8, Math.max(0, ...strs.map((s) => tokenSpans(s, t).length)));
      const idxs: number[] = [];
      for (let i = 0; i < m; i++) idxs.push(i);
      for (let i = 1; i <= m; i++) idxs.push(-i);
      for (const i of idxs) { base.push({ k: 'tok', col, t, i } as any); base.push({ k: 'init', col, t, i } as any); }
      for (const i of idxs) for (const j of idxs) if ((i >= 0 && j >= 0 && j > i) || (i >= 0 && j < 0) || (i < 0 && j < 0 && j > i)) base.push({ k: 'span', col, t, i, j } as any);
    }
    const delims = new Set<string>();
    for (const s of strs) for (const ch of s) if (/[\s,;:/\\|@._\-()#+]/.test(ch)) delims.add(ch);
    for (const d of delims) {
      const m = Math.min(8, Math.max(...strs.map((s) => s.split(d).length)));
      const idxs: number[] = [];
      for (let i = 0; i < m; i++) idxs.push(i);
      for (let i = 1; i <= m; i++) idxs.push(-i);
      for (const i of idxs) for (const j of idxs) if (i === j || (i >= 0 && j >= 0 && j > i) || (i >= 0 && j < 0) || (i < 0 && j < 0 && j > i)) base.push({ k: 'fld', col, d, i, j } as any);
    }
    for (let n = 1; n <= 6; n++) { base.push({ k: 'pre', col, n } as any); base.push({ k: 'suf', col, n } as any); }
    for (const b of base) for (const cs of cases) out.push({ ...(b as any), cs });
  }
  return out;
}

/**
 * Infers the simplest program consistent with every example (source fields -> output). Returns null when no
 * program explains all examples.
 */
export function inferFlashFill(examples: { src: string[]; out: string }[], cols?: number[]): FFProgram | null {
  if (!examples.length || examples.some((e) => e.out === '')) return null;
  const nCols = Math.max(...examples.map((e) => e.src.length));
  const useCols = cols || rowsOf(0, nCols - 1);
  const exprs = candidateExprs(examples.map((e) => e.src), useCols);
  // Evaluate every expression on every example; keep the cheapest per distinct value vector.
  const byKey = new Map<string, { e: FFExpr; vals: string[]; cost: number }>();
  for (const e of exprs) {
    const vals: string[] = [];
    let ok = true;
    for (const ex of examples) { const v = evalExpr(e, ex.src); if (v == null || !ex.out.includes(v)) { ok = false; break; } vals.push(v); }
    if (!ok) continue;
    const key = vals.join('\u0001');
    const cost = exprCost(e);
    const cur = byKey.get(key);
    if (!cur || cost < cur.cost) byKey.set(key, { e, vals, cost });
  }
  const cands = [...byKey.values()];
  const outs = examples.map((e) => e.out);
  const memo = new Map<string, { cost: number; parts: FFPart[] } | null>();
  const solve = (pos: number[], inLit: boolean): { cost: number; parts: FFPart[] } | null => {
    if (pos.every((p, i) => p === outs[i].length)) return { cost: 0, parts: [] };
    const key = pos.join(',') + (inLit ? 'L' : '');
    if (memo.has(key)) return memo.get(key)!;
    memo.set(key, null);
    let best: { cost: number; parts: FFPart[] } | null = null;
    for (const c of cands) {
      if (!c.vals.every((v, i) => outs[i].startsWith(v, pos[i]))) continue;
      const rest = solve(pos.map((p, i) => p + c.vals[i].length), false);
      if (rest && (!best || c.cost + rest.cost < best.cost)) best = { cost: c.cost + rest.cost, parts: [{ e: c.e }, ...rest.parts] };
    }
    if (pos.every((p, i) => p < outs[i].length) && outs.every((o, i) => o[pos[i]] === outs[0][pos[0]])) {
      const ch = outs[0][pos[0]];
      const rest = solve(pos.map((p) => p + 1), true);
      if (rest) {
        const cost = (inLit ? 0 : LIT_START) + LIT_CHAR + rest.cost;
        if (!best || cost < best.cost - 1e-9) {
          // A literal right after a literal step continues the same literal part.
          const head = rest.parts[0];
          const parts = head && 'lit' in head ? [{ lit: ch + head.lit }, ...rest.parts.slice(1)] : [{ lit: ch }, ...rest.parts];
          best = { cost, parts };
        }
      }
    }
    memo.set(key, best);
    return best;
  };
  const res = solve(outs.map(() => 0), false);
  if (!res) return null;
  // Merge adjacent literals (defensive) and return.
  const parts: FFPart[] = [];
  for (const p of res.parts) { const last = parts[parts.length - 1]; if ('lit' in p && last && 'lit' in last) last.lit += p.lit; else parts.push('lit' in p ? { lit: p.lit } : p); }
  return { parts, cost: res.cost };
}

export function runFlashFill(p: FFProgram, src: string[]): string | null {
  let out = '';
  for (const part of p.parts) {
    if ('lit' in part) out += part.lit;
    else { const v = evalExpr(part.e, src); if (v == null) return null; out += v; }
  }
  return out;
}

export function describeFlashFill(p: FFProgram, colLabel: (c: number) => string): string {
  return p.parts.map((x) => ('lit' in x ? JSON.stringify(x.lit) : `${x.e.k}(${colLabel(x.e.col)})`)).join(' & ');
}

// --------------------------------------------------------------------------------------------
// Remove Duplicates

/** Rows (absolute) of a range whose values in `cols` repeat an earlier row (displayed values, case-insensitive). */
export function duplicateRows(sheet: Sheet, rg: Range, cols: number[], header: boolean): number[] {
  const seen = new Set<string>();
  const dups: number[] = [];
  for (let r = rg.r1 + (header ? 1 : 0); r <= rg.r2; r++) {
    const key = cols.map((c) => displayText(sheet, r, c).toLowerCase()).join('\u0001');
    if (seen.has(key)) dups.push(r); else seen.add(key);
  }
  return dups;
}

/** Removes rows of a range (cells within the range shift up). */
export function removeRowsInRange(sheet: Sheet, rg: Range, remove: number[], header: boolean): void {
  const rm = new Set(remove);
  const rows = rowsOf(rg.r1 + (header ? 1 : 0), rg.r2);
  const keep = rows.filter((r) => !rm.has(r));
  rewriteRows(sheet, rg.c1, rg.c2, rows, keep.map((src, i) => [rows[i], src]));
}

// --------------------------------------------------------------------------------------------
// Subtotals and outlines

export const SUBTOTAL_FUNCS: { code: number; name: string; label: string }[] = [
  { code: 9, name: 'Sum', label: 'Total' }, { code: 3, name: 'Count', label: 'Count' }, { code: 1, name: 'Average', label: 'Average' },
  { code: 4, name: 'Max', label: 'Max' }, { code: 5, name: 'Min', label: 'Min' }, { code: 6, name: 'Product', label: 'Product' },
  { code: 2, name: 'Count Numbers', label: 'Count' }, { code: 7, name: 'StdDev', label: 'StdDev' }, { code: 8, name: 'StdDevp', label: 'StdDevp' },
  { code: 10, name: 'Var', label: 'Var' }, { code: 11, name: 'Varp', label: 'Varp' },
];

const isSubtotalRow = (sheet: Sheet, r: number, c1: number, c2: number) => {
  for (let c = c1; c <= c2; c++) { const f = sheet.get(r, c)?.f; if (f && /\bSUBTOTAL\s*\(/i.test(f)) return true; }
  return false;
};

/**
 * Inserts subtotal rows into a list (first row = headers) at each change in `groupCol`. Returns the new list range.
 * Must run inside a transaction.
 */
export function addSubtotals(sheet: Sheet, list: Range, opts: { groupCol: number; fn: number; cols: number[]; replace: boolean; pageBreaks: boolean; below: boolean; boldStyle?: number }, shift: (at: number, n: number) => void): Range {
  const book = sheet.book;
  book.recordSheet(sheet);
  let rg = { ...list };
  if (opts.replace) rg = removeSubtotals(sheet, rg, shift);
  const fn = SUBTOTAL_FUNCS.find((f) => f.code === opts.fn) || SUBTOTAL_FUNCS[0];
  const first = rg.r1 + 1;
  // Groups between existing subtotal rows (nested subtotals keep their outer groups).
  const groups: [number, number][] = [];
  let start = -1, prevKey: string | null = null;
  let hasGrand = false;
  for (let r = first; r <= rg.r2; r++) {
    if (isSubtotalRow(sheet, r, rg.c1, rg.c2)) {
      if (/^grand /i.test(String(sheet.value(r, opts.groupCol) ?? '')) || rowLabelIsGrand(sheet, r, rg)) hasGrand = true;
      if (start >= 0) groups.push([start, r - 1]);
      start = -1; prevKey = null;
      continue;
    }
    const key = displayText(sheet, r, opts.groupCol).toLowerCase();
    if (start < 0) { start = r; prevKey = key; continue; }
    if (key !== prevKey) { groups.push([start, r - 1]); start = r; prevKey = key; }
  }
  if (start >= 0) groups.push([start, rg.r2]);
  if (!groups.length) return rg;
  const baseLevel = (r: number) => sheet.rowLevel.get(r) || 0;
  const lab = fn.label;
  let lastRow = rg.r2;
  for (let gi = groups.length - 1; gi >= 0; gi--) {
    const [a, b] = groups[gi];
    const at = opts.below ? b + 1 : a;
    const lvl = baseLevel(a);
    shift(at, 1);
    lastRow++;
    const da = opts.below ? a : a + 1, db = opts.below ? b : b + 1;
    const name = displayText(sheet, da, opts.groupCol);
    for (let c = rg.c1; c <= rg.c2; c++) sheet.set(at, c, undefined);
    sheet.set(at, opts.groupCol, { v: `${name} ${lab}`, ...(opts.boldStyle ? { s: opts.boldStyle } : {}) });
    for (const c of opts.cols) {
      const ref = `${colName(c)}${da + 1}:${colName(c)}${db + 1}`;
      const cur = sheet.get(da, c);
      const s = cur?.s;
      sheet.set(at, c, { f: `SUBTOTAL(${fn.code},${ref})`, ...(s ? { s } : {}) });
    }
    // First subtotals: subtotal rows are level 1 and details level 2; nested ones go one level deeper.
    const base = Math.min(7, lvl === 0 ? 1 : lvl);
    sheet.rowLevel.set(at, base);
    for (let r = da; r <= db; r++) sheet.rowLevel.set(r, base + 1);
    if (opts.pageBreaks && opts.below) sheet.rowBreaks.push(at + 1);
  }
  if (!hasGrand) {
    const gAt = opts.below ? lastRow + 1 : first;
    shift(gAt, 1);
    lastRow++;
    const dFirst = opts.below ? first : first + 1, dLast = opts.below ? lastRow - 1 : lastRow;
    for (let c = rg.c1; c <= rg.c2; c++) sheet.set(gAt, c, undefined);
    sheet.set(gAt, opts.groupCol, { v: `Grand ${lab}`, ...(opts.boldStyle ? { s: opts.boldStyle } : {}) });
    for (const c of opts.cols) {
      const s = sheet.get(dFirst, c)?.s;
      sheet.set(gAt, c, { f: `SUBTOTAL(${fn.code},${colName(c)}${dFirst + 1}:${colName(c)}${dLast + 1})`, ...(s ? { s } : {}) });
    }
    sheet.rowLevel.delete(gAt);
  }
  if (opts.pageBreaks) sheet.rowBreaks = [...new Set(sheet.rowBreaks)].filter((r) => r <= lastRow).sort((a, b) => a - b);
  book.pending.structural = true;
  return { ...rg, r2: lastRow };
}

function rowLabelIsGrand(sheet: Sheet, r: number, rg: Range): boolean {
  for (let c = rg.c1; c <= rg.c2; c++) { const v = sheet.value(r, c); if (typeof v === 'string' && /^grand /i.test(v)) return true; }
  return false;
}

/** Removes subtotal rows (rows holding SUBTOTAL formulas) from a list and clears its outline. Returns the new range. */
export function removeSubtotals(sheet: Sheet, list: Range, shift: (at: number, n: number) => void): Range {
  sheet.book.recordSheet(sheet);
  let r2 = list.r2;
  for (let r = list.r2; r > list.r1; r--) {
    if (isSubtotalRow(sheet, r, list.c1, list.c2)) { shift(r, -1); r2--; }
  }
  for (let r = list.r1; r <= r2; r++) { sheet.rowLevel.delete(r); sheet.hiddenRows.delete(r); }
  sheet.rowBreaks = sheet.rowBreaks.filter((b) => b <= list.r1 || b > r2 + 1);
  return { ...list, r2 };
}

export type OutlineAxis = 'row' | 'col';

const levels = (sheet: Sheet, axis: OutlineAxis) => (axis === 'row' ? sheet.rowLevel : sheet.colLevel);
const hidden = (sheet: Sheet, axis: OutlineAxis) => (axis === 'row' ? sheet.hiddenRows : sheet.hiddenCols);

/** Groups (delta = 1) or ungroups (delta = -1) indices a..b. Returns false when nothing could change. */
export function changeLevel(sheet: Sheet, axis: OutlineAxis, a: number, b: number, delta: 1 | -1): boolean {
  const m = levels(sheet, axis);
  let changed = false;
  for (let i = a; i <= b; i++) {
    const cur = m.get(i) || 0;
    const next = Math.max(0, Math.min(8, cur + delta));
    if (next === cur) continue;
    changed = true;
    if (next) m.set(i, next); else { m.delete(i); hidden(sheet, axis).delete(i); }
  }
  return changed;
}

export function maxLevel(sheet: Sheet, axis: OutlineAxis): number {
  let mx = 0;
  for (const v of levels(sheet, axis).values()) mx = Math.max(mx, v);
  return mx;
}

/** Shows outline level n (1 = only the top summary): hides every index whose level is >= n. */
export function showOutlineLevel(sheet: Sheet, axis: OutlineAxis, n: number): void {
  const h = hidden(sheet, axis);
  for (const [i, lv] of levels(sheet, axis)) { if (lv >= n) h.add(i); else h.delete(i); }
}

/** The detail block that a summary/detail index controls: [a, b, level]. */
export function detailBlock(sheet: Sheet, axis: OutlineAxis, idx: number): [number, number, number] | null {
  const m = levels(sheet, axis);
  const lv = (i: number) => (i < 0 ? 0 : m.get(i) || 0);
  let target: number, a: number;
  if (lv(idx - 1) > lv(idx)) { target = lv(idx) + 1; a = idx - 1; }        // summary row below its detail
  else if (lv(idx) > 0) { target = lv(idx); a = idx; }                      // inside a group
  else if (lv(idx + 1) > lv(idx)) { target = lv(idx) + 1; a = idx + 1; }    // summary row above its detail
  else return null;
  let b = a;
  while (a > 0 && lv(a - 1) >= target) a--;
  while (lv(b + 1) >= target) b++;
  return [a, b, target];
}

/** Hide Detail / Show Detail for the group an index belongs to (or summarizes). */
export function setDetail(sheet: Sheet, axis: OutlineAxis, idx: number, show: boolean): boolean {
  const blk = detailBlock(sheet, axis, idx);
  if (!blk) return false;
  const [a, b] = blk;
  const h = hidden(sheet, axis);
  for (let i = a; i <= b; i++) { if (show) h.delete(i); else h.add(i); }
  return true;
}

/**
 * Auto Outline: rows (or columns) summarized by SUM/SUBTOTAL formulas that refer to the contiguous block right
 * above (left of) them become a group; larger groups nest the smaller ones.
 */
export function autoOutline(sheet: Sheet, rg: Range): { rows: number; cols: number } {
  const rowGroups: [number, number][] = [], colGroups: [number, number][] = [];
  sheet.forEachIn(rg, (r, c, cell) => {
    if (!cell.f || !/^\s*(SUM|SUBTOTAL)\s*\(/i.test(cell.f)) return;
    const refs = [...cell.f.matchAll(/\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)/g)];
    for (const m of refs) {
      const c1 = colIdx(m[1]), r1 = +m[2] - 1, c2 = colIdx(m[3]), r2 = +m[4] - 1;
      if (c1 === c && c2 === c && r2 === r - 1 && r1 <= r2) rowGroups.push([r1, r2]);
      else if (r1 === r && r2 === r && c2 === c - 1 && c1 <= c2) colGroups.push([c1, c2]);
    }
  });
  const uniq = (gs: [number, number][]) => [...new Map(gs.map((g) => [g.join(':'), g])).values()].sort((a, b) => (b[1] - b[0]) - (a[1] - a[0]));
  const rg2 = uniq(rowGroups), cg2 = uniq(colGroups);
  if (rg2.length) { sheet.rowLevel.clear(); for (const [a, b] of rg2) changeLevel(sheet, 'row', a, b, 1); }
  if (cg2.length) { sheet.colLevel.clear(); for (const [a, b] of cg2) changeLevel(sheet, 'col', a, b, 1); }
  return { rows: rg2.length, cols: cg2.length };
}

function colIdx(name: string): number { let n = 0; for (const ch of name) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; }

// --------------------------------------------------------------------------------------------
// Consolidate and aggregation

export const AGG_FUNCS = ['Sum', 'Count', 'Average', 'Max', 'Min', 'Product', 'Count Numbers', 'StdDev', 'StdDevp', 'Var', 'Varp'] as const;
export type AggName = typeof AGG_FUNCS[number];

const AGG_FORMULA: Record<AggName, string> = { Sum: 'SUM', Count: 'COUNTA', Average: 'AVERAGE', Max: 'MAX', Min: 'MIN', Product: 'PRODUCT', 'Count Numbers': 'COUNT', StdDev: 'STDEV.S', StdDevp: 'STDEV.P', Var: 'VAR.S', Varp: 'VAR.P' };
export const aggFormulaName = (n: AggName) => AGG_FORMULA[n];

export function aggregate(fn: AggName, vals: Scalar[]): Scalar {
  const nums = vals.filter((v): v is number => typeof v === 'number');
  const n = nums.length;
  const sum = nums.reduce((a, b) => a + b, 0);
  const mean = n ? sum / n : NaN;
  const ss = nums.reduce((a, b) => a + (b - mean) ** 2, 0);
  switch (fn) {
    case 'Sum': return sum;
    case 'Count': return vals.filter((v) => !isBlank(v)).length;
    case 'Count Numbers': return n;
    case 'Average': return n ? mean : new CellError('#DIV/0!');
    case 'Max': return n ? Math.max(...nums) : 0;
    case 'Min': return n ? Math.min(...nums) : 0;
    case 'Product': return n ? nums.reduce((a, b) => a * b, 1) : 0;
    case 'StdDev': return n > 1 ? Math.sqrt(ss / (n - 1)) : new CellError('#DIV/0!');
    case 'StdDevp': return n ? Math.sqrt(ss / n) : new CellError('#DIV/0!');
    case 'Var': return n > 1 ? ss / (n - 1) : new CellError('#DIV/0!');
    case 'Varp': return n ? ss / n : new CellError('#DIV/0!');
  }
}

export interface ConsolidateSource { values: Scalar[][]; ref?: (i: number, j: number) => string }

/**
 * Consolidates source blocks by position or by labels (top row and/or left column). Returns the output matrix
 * (with labels when used) and, per output data cell, the source cell references that contributed.
 */
export function consolidate(sources: ConsolidateSource[], fn: AggName, top: boolean, left: boolean): { out: Scalar[][]; refs: (string[] | null)[][] } {
  const rowKeys: string[] = [], colKeys: string[] = [];
  const rowLabel = new Map<string, Scalar>(), colLabel = new Map<string, Scalar>();
  const cells = new Map<string, { vals: Scalar[]; refs: string[] }>();
  const keyOf = (v: Scalar) => String(v ?? '').trim().toLowerCase();
  for (const src of sources) {
    const m = src.values;
    for (let i = top ? 1 : 0; i < m.length; i++) {
      const rk = left ? keyOf(m[i][0]) : String(i - (top ? 1 : 0));
      if (left && rk === '') continue;
      if (!rowLabel.has(rk)) { rowLabel.set(rk, left ? m[i][0] : null); rowKeys.push(rk); }
      for (let j = left ? 1 : 0; j < (m[i]?.length || 0); j++) {
        const ck = top ? keyOf(m[0][j]) : String(j - (left ? 1 : 0));
        if (top && ck === '') continue;
        if (!colLabel.has(ck)) { colLabel.set(ck, top ? m[0][j] : null); colKeys.push(ck); }
        const k = rk + '\u0001' + ck;
        let e = cells.get(k);
        if (!e) { e = { vals: [], refs: [] }; cells.set(k, e); }
        e.vals.push(m[i][j]);
        if (src.ref) e.refs.push(src.ref(i, j));
      }
    }
  }
  if (!top) colKeys.sort((a, b) => +a - +b);
  if (!left) rowKeys.sort((a, b) => +a - +b);
  const out: Scalar[][] = [], refs: (string[] | null)[][] = [];
  if (top) { out.push([...(left ? [null] : []), ...colKeys.map((k) => colLabel.get(k) ?? null)]); refs.push(out[0].map(() => null)); }
  for (const rk of rowKeys) {
    const row: Scalar[] = left ? [rowLabel.get(rk) ?? null] : [];
    const rrow: (string[] | null)[] = left ? [null] : [];
    for (const ck of colKeys) {
      const e = cells.get(rk + '\u0001' + ck);
      row.push(e ? aggregate(fn, e.vals) : null);
      rrow.push(e ? e.refs : null);
    }
    out.push(row); refs.push(rrow);
  }
  return { out, refs };
}

// --------------------------------------------------------------------------------------------
// Transform (Get Data from Table/Range)

export type TransformStep =
  | { kind: 'filter'; col: string; op: string; value: string }
  | { kind: 'removeCols'; cols: string[] }
  | { kind: 'keepCols'; cols: string[] }
  | { kind: 'changeType'; col: string; type: 'number' | 'text' | 'date' | 'boolean' }
  | { kind: 'split'; col: string; delimiter: string }
  | { kind: 'unpivot'; keep: string[]; attr?: string; value?: string }
  | { kind: 'groupBy'; keys: string[]; col: string; fn: AggName; as?: string }
  | { kind: 'sort'; col: string; desc: boolean }
  | { kind: 'removeDuplicates' }
  | { kind: 'removeBlankRows' }
  | { kind: 'rename'; col: string; to: string };

export interface TableData { headers: string[]; rows: Scalar[][] }

export function describeStep(s: TransformStep): string {
  switch (s.kind) {
    case 'filter': return `Filter rows: ${s.col} ${s.op} ${s.value}`;
    case 'removeCols': return `Remove columns: ${s.cols.join(', ')}`;
    case 'keepCols': return `Keep columns: ${s.cols.join(', ')}`;
    case 'changeType': return `Change type: ${s.col} → ${s.type}`;
    case 'split': return `Split column ${s.col} by "${s.delimiter}"`;
    case 'unpivot': return `Unpivot other columns (keep ${s.keep.join(', ') || 'none'})`;
    case 'groupBy': return `Group by ${s.keys.join(', ')}: ${s.fn} of ${s.col}`;
    case 'sort': return `Sort by ${s.col} ${s.desc ? 'descending' : 'ascending'}`;
    case 'removeDuplicates': return 'Remove duplicate rows';
    case 'removeBlankRows': return 'Remove blank rows';
    case 'rename': return `Rename ${s.col} → ${s.to}`;
  }
}

export function applySteps(data: TableData, steps: TransformStep[]): TableData {
  let { headers, rows } = { headers: [...data.headers], rows: data.rows.map((r) => [...r]) };
  const idx = (name: string) => { const i = headers.findIndex((h) => h.toLowerCase() === name.toLowerCase()); if (i < 0) throw new Error(`The column "${name}" wasn't found.`); return i; };
  for (const s of steps) {
    switch (s.kind) {
      case 'filter': {
        const i = idx(s.col);
        rows = rows.filter((r) => customMatch(s.op, s.value, r[i] ?? null, r[i] == null ? '' : String(r[i])));
        break;
      }
      case 'removeCols': case 'keepCols': {
        const set = new Set(s.cols.map(idx));
        const keep = headers.map((_, i) => (s.kind === 'removeCols' ? !set.has(i) : set.has(i)));
        headers = headers.filter((_, i) => keep[i]);
        rows = rows.map((r) => r.filter((_, i) => keep[i]));
        break;
      }
      case 'changeType': {
        const i = idx(s.col);
        for (const r of rows) {
          const v = r[i];
          if (v == null || v === '') continue;
          if (s.type === 'text') r[i] = typeof v === 'number' ? String(v) : String(v);
          else if (s.type === 'number') { const n = typeof v === 'number' ? v : parseNumeric(String(v)); r[i] = n == null ? new CellError('#VALUE!') : n; }
          else if (s.type === 'boolean') r[i] = typeof v === 'boolean' ? v : /^(true|yes|1)$/i.test(String(v));
          else if (s.type === 'date') { if (typeof v !== 'number') { const c = convertField(String(v)); r[i] = typeof c.v === 'number' ? c.v : new CellError('#VALUE!'); } }
        }
        break;
      }
      case 'split': {
        const i = idx(s.col);
        const parts = rows.map((r) => (r[i] == null ? [''] : String(r[i]).split(s.delimiter)));
        const n = Math.max(1, ...parts.map((p) => p.length));
        const names = rowsOf(1, n).map((k) => `${headers[i]}.${k}`);
        headers = [...headers.slice(0, i), ...names, ...headers.slice(i + 1)];
        rows = rows.map((r, ri) => [...r.slice(0, i), ...rowsOf(0, n - 1).map((k) => { const t = parts[ri][k]; if (t == null || t === '') return null; const c = convertField(t.trim()); return c.f != null ? t : c.v; }), ...r.slice(i + 1)]);
        break;
      }
      case 'unpivot': {
        const keep = s.keep.map(idx);
        const others = headers.map((_, i) => i).filter((i) => !keep.includes(i));
        const out: Scalar[][] = [];
        for (const r of rows) for (const o of others) { if (r[o] == null || r[o] === '') continue; out.push([...keep.map((k) => r[k]), headers[o], r[o]]); }
        headers = [...keep.map((k) => headers[k]), s.attr || 'Attribute', s.value || 'Value'];
        rows = out;
        break;
      }
      case 'groupBy': {
        const ks = s.keys.map(idx), vi = idx(s.col);
        const groups = new Map<string, { key: Scalar[]; vals: Scalar[] }>();
        for (const r of rows) {
          const key = ks.map((k) => r[k]);
          const kk = key.map((x) => String(x ?? '').toLowerCase()).join('\u0001');
          let g = groups.get(kk);
          if (!g) { g = { key, vals: [] }; groups.set(kk, g); }
          g.vals.push(r[vi]);
        }
        headers = [...ks.map((k) => headers[k]), s.as || `${s.fn} of ${headers[vi]}`];
        rows = [...groups.values()].map((g) => [...g.key, aggregate(s.fn, g.vals)]);
        break;
      }
      case 'sort': { const i = idx(s.col); rows = rows.map((r, k) => ({ r, k })).sort((a, b) => compareForSort(a.r[i] ?? null, b.r[i] ?? null, s.desc) || a.k - b.k).map((x) => x.r); break; }
      case 'removeDuplicates': { const seen = new Set<string>(); rows = rows.filter((r) => { const k = r.map((x) => String(x ?? '').toLowerCase()).join('\u0001'); if (seen.has(k)) return false; seen.add(k); return true; }); break; }
      case 'removeBlankRows': rows = rows.filter((r) => r.some((x) => !isBlank(x))); break;
      case 'rename': { const i = idx(s.col); headers[i] = s.to; break; }
    }
  }
  return { headers, rows };
}

// --------------------------------------------------------------------------------------------
// Misc

/** Ranges of `a` not covered by `b`. */
export function subtractRange(a: Range, b: Range): Range[] {
  if (b.r2 < a.r1 || b.r1 > a.r2 || b.c2 < a.c1 || b.c1 > a.c2) return [a];
  const out: Range[] = [];
  if (b.r1 > a.r1) out.push({ r1: a.r1, c1: a.c1, r2: b.r1 - 1, c2: a.c2 });
  if (b.r2 < a.r2) out.push({ r1: b.r2 + 1, c1: a.c1, r2: a.r2, c2: a.c2 });
  const r1 = Math.max(a.r1, b.r1), r2 = Math.min(a.r2, b.r2);
  if (b.c1 > a.c1) out.push({ r1, c1: a.c1, r2, c2: b.c1 - 1 });
  if (b.c2 < a.c2) out.push({ r1, c1: b.c2 + 1, r2, c2: a.c2 });
  return out;
}

export function keyToRC(k: number): { r: number; c: number } { return { r: keyRow(k), c: keyCol(k) }; }

export type { AutoFilter };
