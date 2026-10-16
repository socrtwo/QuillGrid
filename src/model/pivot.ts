/**
 * PivotTable computation (pure; no DOM, no workbook access).
 *
 * computePivot(def, source) groups the source rows by the row and column fields, aggregates every value
 * field for every combination (including subtotals and grand totals), applies "show values as", and lays
 * the result out as a grid of cells in compact, outline or tabular form, ready to be written to a sheet.
 *
 * Terminology follows the spreadsheet convention: "grand totals for rows" (grandRows) is the Grand Total
 * column at the right (it totals each row); "grand totals for columns" (grandCols) is the Grand Total row at
 * the bottom (it totals each column).
 */
import { CellError, type PivotDef, type Scalar } from './types';
import { serialToParts, MONTHS } from './dates';

export type PivotFn = PivotDef['values'][number]['fn'];
export type ShowAs = NonNullable<PivotDef['values'][number]['showAs']>;
export type DateGroup = 'years' | 'quarters' | 'months' | 'days';

export interface PivotSource { headers: string[]; rows: Scalar[][] }

export type PivotCellStyle = 'header' | 'total' | 'grand' | 'label' | 'caption' | 'value' | 'filter';
export interface PivotCellObj {
  v: Scalar;
  s?: PivotCellStyle;
  /** Label indent level (compact layout). */
  indent?: number;
  /** Source field of a label (used to copy its number format). */
  field?: string;
  /** Value field index of a value cell. */
  vi?: number;
  /** The value is a ratio shown as a percentage. */
  pct?: boolean;
  /** Group (non-leaf) row label or subtotal: shown bold. */
  bold?: boolean;
}
export type PivotCell = Scalar | PivotCellObj;

export interface PivotResult {
  cells: PivotCell[][];
  rows: number;
  cols: number;
  /** Rows taken by report filters (plus the blank separator row). */
  filterRows: number;
  /** Header rows of the table body (after the filter rows). */
  headerRows: number;
  /** Number of label columns at the left of the body. */
  labelCols: number;
  /** For value cells, "r,c" → the row and column item keys behind them (drill-down). */
  drill: Map<string, { row: string[]; col: string[] }>;
  /** Fields expanded from date grouping, in display order. */
  rowFields: string[];
  colFields: string[];
  /** Number of source rows that passed the filters. */
  records: number;
}

// --------------------------------------------------------------------------------------------
// Field items

interface FieldSpec { name: string; col: number; group?: DateGroup; base: string }
interface ItemKey { key: string; label: Scalar; sort: number | string; blank?: boolean }

const BLANK_KEY = '\u0000blank';
const MON3 = MONTHS.map((m) => m.slice(0, 3));

export function itemKey(v: Scalar, group?: DateGroup, date1904 = false): ItemKey {
  if (v == null || v === '') return { key: BLANK_KEY, label: '(blank)', sort: '￿', blank: true };
  if (group) {
    if (typeof v !== 'number') return itemKey(v, undefined, date1904);
    const p = serialToParts(Math.floor(v), date1904);
    switch (group) {
      case 'years': return { key: 'y:' + p.y, label: p.y, sort: p.y };
      case 'quarters': { const q = Math.floor((p.m - 1) / 3) + 1; return { key: 'q:' + q, label: 'Qtr' + q, sort: q }; }
      case 'months': return { key: 'm:' + p.m, label: MON3[p.m - 1], sort: p.m };
      case 'days': return { key: 'd:' + p.m + '-' + p.d, label: `${p.d}-${MON3[p.m - 1]}`, sort: p.m * 100 + p.d };
    }
  }
  if (typeof v === 'number') return { key: 'n:' + v, label: v, sort: v };
  if (typeof v === 'boolean') return { key: 'b:' + v, label: v, sort: v ? 'TRUE' : 'FALSE' };
  if (v instanceof CellError) return { key: 'e:' + v.code, label: v, sort: '￰' + v.code };
  return { key: 's:' + v.toLowerCase(), label: v, sort: v.toLowerCase() };
}

/** Text shown for an item in filter lists (and stored in filterValues). */
export function itemText(k: ItemKey): string {
  if (k.blank) return '(blank)';
  const l = k.label;
  if (typeof l === 'boolean') return l ? 'TRUE' : 'FALSE';
  if (l instanceof CellError) return l.code;
  return String(l);
}

function cmpSort(a: number | string, b: number | string): number {
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'number') return -1;
  if (typeof b === 'number') return 1;
  return a < b ? -1 : a > b ? 1 : 0;
}

export const GROUP_LABEL: Record<DateGroup, string> = { years: 'Years', quarters: 'Quarters', months: 'Months', days: 'Days' };

/** Expands date-grouped fields into one virtual field per grouping level. */
function expandFields(def: PivotDef, names: string[], headers: string[]): FieldSpec[] {
  const out: FieldSpec[] = [];
  for (const n of names) {
    const col = findField(headers, n);
    if (col < 0) continue;
    const groups = def.dateGroups?.[n];
    if (groups?.length) {
      const order: DateGroup[] = ['years', 'quarters', 'months', 'days'];
      for (const g of order.filter((x) => groups.includes(x))) out.push({ name: `${GROUP_LABEL[g]} (${headers[col]})`, col, group: g, base: headers[col] });
    } else out.push({ name: headers[col], col, base: headers[col] });
  }
  return out;
}

export function findField(headers: string[], name: string): number {
  const i = headers.indexOf(name);
  if (i >= 0) return i;
  const low = name.toLowerCase();
  return headers.findIndex((x) => x.toLowerCase() === low);
}

/** Unique items of a field (texts, sorted), for filter lists. */
export function pivotItems(source: PivotSource, field: string, group?: DateGroup): string[] {
  const col = findField(source.headers, field);
  if (col < 0) return [];
  const seen = new Map<string, ItemKey>();
  for (const row of source.rows) { const k = itemKey(row[col] ?? null, group); if (!seen.has(k.key)) seen.set(k.key, k); }
  return [...seen.values()].sort((a, b) => cmpSort(a.sort, b.sort)).map(itemText);
}

/** Indexes of the source rows that pass the report filters and item filters. */
export function filteredRows(def: PivotDef, source: PivotSource): number[] {
  const checks: { col: number; keep: Set<string> }[] = [];
  for (const [f, list] of Object.entries(def.filterValues || {})) {
    if (!list) continue;
    const col = findField(source.headers, f);
    if (col < 0) continue;
    checks.push({ col, keep: new Set(list.map((x) => x.toLowerCase())) });
  }
  const out: number[] = [];
  for (let i = 0; i < source.rows.length; i++) {
    const row = source.rows[i];
    if (row.every((x) => x == null || x === '')) continue;
    if (checks.every((ch) => ch.keep.has(itemText(itemKey(row[ch.col] ?? null)).toLowerCase()))) out.push(i);
  }
  return out;
}

// --------------------------------------------------------------------------------------------
// Aggregation

interface Acc { sum: number; count: number; nums: number[]; product: number; min: number; max: number }
const newAcc = (): Acc => ({ sum: 0, count: 0, nums: [], product: 1, min: Infinity, max: -Infinity });

function addTo(a: Acc, v: Scalar): void {
  if (v == null || v === '') return;
  a.count++;
  if (typeof v === 'number') { a.sum += v; a.nums.push(v); a.product *= v; if (v < a.min) a.min = v; if (v > a.max) a.max = v; }
}

export function aggregate(fn: PivotFn, a: Acc | undefined): Scalar {
  if (!a) return null;
  const n = a.nums.length;
  switch (fn) {
    case 'sum': return a.sum;
    case 'count': return a.count;
    case 'countNums': return n;
    case 'average': return n ? a.sum / n : new CellError('#DIV/0!');
    case 'max': return n ? a.max : 0;
    case 'min': return n ? a.min : 0;
    case 'product': return n ? a.product : 0;
    case 'stdDev': case 'stdDevP': case 'var': case 'varP': {
      const pop = fn === 'stdDevP' || fn === 'varP';
      if (n < (pop ? 1 : 2)) return new CellError('#DIV/0!');
      const mean = a.sum / n;
      const ss = a.nums.reduce((s, x) => s + (x - mean) ** 2, 0);
      const variance = ss / (pop ? n : n - 1);
      return fn.startsWith('var') ? variance : Math.sqrt(variance);
    }
  }
  return null;
}

export const FN_LABELS: Record<PivotFn, string> = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count Numbers', stdDev: 'StdDev', stdDevP: 'StdDevp', var: 'Var', varP: 'Varp' };
export const SHOW_AS_LABELS: Record<ShowAs, string> = { normal: 'No Calculation', percentOfTotal: '% of Grand Total', percentOfCol: '% of Column Total', percentOfRow: '% of Row Total', runningTotal: 'Running Total In', rank: 'Rank Largest to Smallest' };

export function valueName(v: PivotDef['values'][number]): string {
  return v.name || `${FN_LABELS[v.fn]} of ${v.field}`;
}

// --------------------------------------------------------------------------------------------
// Item trees

interface TreeNode { key: string; item: ItemKey; path: string[]; children: TreeNode[]; childMap: Map<string, TreeNode>; parent: TreeNode | null; depth: number }

function newNode(key: string, item: ItemKey, parent: TreeNode | null): TreeNode {
  return { key, item, path: parent ? [...parent.path, key] : [], children: [], childMap: new Map(), parent, depth: parent ? parent.depth + 1 : -1 };
}

function insertPath(root: TreeNode, items: ItemKey[]): void {
  let n = root;
  for (const it of items) {
    let ch = n.childMap.get(it.key);
    if (!ch) { ch = newNode(it.key, it, n); n.childMap.set(it.key, ch); n.children.push(ch); }
    n = ch;
  }
}

function sortTree(n: TreeNode, specs: FieldSpec[], def: PivotDef): void {
  if (!n.children.length) return;
  const spec = specs[n.depth + 1];
  const desc = spec && (def.sort?.[spec.name] === 'desc' || (!spec.group && def.sort?.[spec.base] === 'desc')) ? -1 : 1;
  n.children.sort((a, b) => {
    if (a.item.blank !== b.item.blank) return a.item.blank ? 1 : -1;
    return desc * cmpSort(a.item.sort, b.item.sort);
  });
  for (const c of n.children) sortTree(c, specs, def);
}

const SEP = '\u0001';
const pathKey = (p: string[]) => p.join(SEP);

// --------------------------------------------------------------------------------------------
// Layout

interface ColSlot { path: string[]; kind: 'item' | 'subtotal' | 'grand'; node: TreeNode | null }
interface RowEntry { kind: 'item' | 'group' | 'subtotal' | 'grand'; node: TreeNode | null; path: string[] }

export function computePivot(def: PivotDef, source: PivotSource, opts: { date1904?: boolean } = {}): PivotResult {
  const headers = source.headers;
  const rowSpecs = expandFields(def, def.rows, headers);
  const colSpecs = expandFields(def, def.cols, headers);
  const vals = def.values.filter((v) => findField(headers, v.field) >= 0);
  const valCols = vals.map((v) => findField(headers, v.field));
  const nV = vals.length;
  const drill = new Map<string, { row: string[]; col: string[] }>();

  // Report filter rows
  const cells: PivotCell[][] = [];
  for (const f of def.filters) {
    if (findField(headers, f) < 0) continue;
    const sel = def.filterValues?.[f];
    cells.push([{ v: f, s: 'filter' }, { v: !sel ? '(All)' : sel.length === 1 ? sel[0] : '(Multiple Items)', s: 'filter' }]);
  }
  const filterRows = cells.length ? cells.length + 1 : 0;
  if (cells.length) cells.push([]);

  if (!rowSpecs.length && !colSpecs.length && !nV) {
    cells.push([{ v: def.name ? `${def.name}: choose fields in the PivotTable Fields pane` : 'Choose fields in the PivotTable Fields pane', s: 'caption' }]);
    return finish(cells, filterRows, 0, 0, drill, [], [], 0);
  }

  // Group and aggregate
  const rowRoot = newNode('', { key: '', label: '', sort: '' }, null);
  const colRoot = newNode('', { key: '', label: '', sort: '' }, null);
  const buckets = new Map<string, Acc[]>();
  const rowsIdx = filteredRows(def, source);
  for (const i of rowsIdx) {
    const row = source.rows[i];
    const rItems = rowSpecs.map((s) => itemKey(row[s.col] ?? null, s.group, opts.date1904));
    const cItems = colSpecs.map((s) => itemKey(row[s.col] ?? null, s.group, opts.date1904));
    insertPath(rowRoot, rItems);
    insertPath(colRoot, cItems);
    const rKeys = rItems.map((x) => x.key), cKeys = cItems.map((x) => x.key);
    for (let a = 0; a <= rKeys.length; a++) {
      const rk = pathKey(rKeys.slice(0, a));
      for (let b = 0; b <= cKeys.length; b++) {
        const k = rk + '\u0002' + pathKey(cKeys.slice(0, b));
        let accs = buckets.get(k);
        if (!accs) { accs = vals.map(newAcc); buckets.set(k, accs); }
        for (let vi = 0; vi < nV; vi++) addTo(accs[vi], row[valCols[vi]] ?? null);
      }
    }
  }
  sortTree(rowRoot, rowSpecs, def);
  sortTree(colRoot, colSpecs, def);

  const raw = (rp: string[], cp: string[], vi: number): Scalar => aggregate(vals[vi].fn, buckets.get(pathKey(rp) + '\u0002' + pathKey(cp))?.[vi]);
  const asNum = (x: Scalar) => (typeof x === 'number' ? x : 0);
  const shown = (entry: RowEntry, slot: ColSlot, vi: number): { v: Scalar; pct?: boolean } => {
    const rp = entry.path, cp = slot.path;
    const v = raw(rp, cp, vi);
    const mode = vals[vi].showAs || 'normal';
    if (mode === 'normal' || v == null || v instanceof CellError) return { v };
    const ratio = (d: Scalar) => (typeof d === 'number' && d !== 0 && typeof v === 'number' ? v / d : new CellError('#DIV/0!'));
    switch (mode) {
      case 'percentOfTotal': return { v: ratio(raw([], [], vi)), pct: true };
      case 'percentOfRow': return { v: ratio(raw(rp, [], vi)), pct: true };
      case 'percentOfCol': return { v: ratio(raw([], cp, vi)), pct: true };
      case 'runningTotal': case 'rank': {
        const node = entry.node;
        if (!node || entry.kind === 'grand') return { v: null };
        const sibs = node.parent ? node.parent.children : [node];
        if (mode === 'runningTotal') {
          let acc = 0;
          for (const s of sibs) { acc += asNum(raw(s.path, cp, vi)); if (s === node) break; }
          return { v: acc };
        }
        let rank = 1;
        for (const s of sibs) { const x = raw(s.path, cp, vi); if (typeof x === 'number' && typeof v === 'number' && x > v) rank++; }
        return { v: rank };
      }
    }
    return { v };
  };

  // Column slots
  const slots: ColSlot[] = [];
  if (!colSpecs.length) slots.push({ path: [], kind: 'item', node: null });
  else {
    const walkCols = (n: TreeNode) => {
      for (const ch of n.children) {
        if (ch.children.length) {
          walkCols(ch);
          if (def.subtotals) slots.push({ path: ch.path, kind: 'subtotal', node: ch });
        } else slots.push({ path: ch.path, kind: 'item', node: ch });
      }
    };
    walkCols(colRoot);
    if (def.grandRows) slots.push({ path: [], kind: 'grand', node: null });
  }
  const perSlot = Math.max(1, nV);

  // Row entries
  const entries: RowEntry[] = [];
  const layout = def.layout || 'compact';
  if (!rowSpecs.length) entries.push({ kind: 'grand', node: null, path: [] });
  else {
    const walkRows = (n: TreeNode) => {
      for (const ch of n.children) {
        if (ch.children.length) {
          if (layout !== 'tabular') entries.push({ kind: 'group', node: ch, path: ch.path });
          walkRows(ch);
          if (layout === 'tabular' && def.subtotals) entries.push({ kind: 'subtotal', node: ch, path: ch.path });
        } else entries.push({ kind: 'item', node: ch, path: ch.path });
      }
    };
    walkRows(rowRoot);
    if (def.grandCols) entries.push({ kind: 'grand', node: null, path: [] });
  }

  const labelCols = rowSpecs.length ? (layout === 'compact' ? 1 : rowSpecs.length) : colSpecs.length ? 1 : 0;
  const totalCols = labelCols + slots.length * perSlot;
  const blankRow = (): PivotCell[] => new Array(totalCols).fill(null);

  // Header rows
  const headerStart = cells.length;
  const vName = (vi: number) => valueName(vals[vi]);
  const labelHeader = (row: PivotCell[]) => {
    if (!rowSpecs.length) return;
    if (layout === 'compact') row[0] = { v: 'Row Labels', s: 'header' };
    else rowSpecs.forEach((s, i) => { row[i] = { v: s.name, s: 'header' }; });
  };
  const itemLabel = (n: TreeNode) => n.item.label;
  if (colSpecs.length) {
    const cap = blankRow();
    if (nV === 1) cap[0] = { v: vName(0), s: 'caption' };
    cap[labelCols] = { v: 'Column Labels', s: 'header' };
    cells.push(cap);
    for (let L = 0; L < colSpecs.length; L++) {
      const row = blankRow();
      if (L === colSpecs.length - 1 && nV <= 1) labelHeader(row);
      let prev: string | null = null;
      slots.forEach((slot, si) => {
        const at = labelCols + si * perSlot;
        if (slot.kind === 'grand') { if (L === 0) row[at] = { v: nV > 1 ? '' : 'Grand Total', s: 'header' }; return; }
        if (slot.kind === 'subtotal') { if (slot.path.length - 1 === L) row[at] = { v: `${itemText(slot.node!.item)} Total`, s: 'header', bold: true }; return; }
        if (L < slot.path.length) {
          const k = pathKey(slot.path.slice(0, L + 1));
          if (k !== prev) {
            let n: TreeNode = slot.node!;
            while (n.depth > L) n = n.parent!;
            row[at] = { v: itemLabel(n), s: 'header', field: colSpecs[L].group ? undefined : colSpecs[L].base };
            prev = k;
          }
        }
      });
      cells.push(row);
    }
    if (nV > 1) {
      const row = blankRow();
      labelHeader(row);
      slots.forEach((slot, si) => {
        for (let vi = 0; vi < nV; vi++) {
          const at = labelCols + si * perSlot + vi;
          const nm = vName(vi);
          row[at] = { v: slot.kind === 'grand' ? `Total ${nm}` : slot.kind === 'subtotal' ? `${itemText(slot.node!.item)} ${nm}` : nm, s: 'header' };
        }
      });
      cells.push(row);
    }
  } else if (nV || rowSpecs.length) {
    const row = blankRow();
    labelHeader(row);
    for (let vi = 0; vi < nV; vi++) row[labelCols + vi] = { v: vName(vi), s: 'header' };
    cells.push(row);
  }
  const headerRows = cells.length - headerStart;

  // Body rows
  for (const e of entries) {
    const row = blankRow();
    const r = cells.length;
    const isGrand = e.kind === 'grand';
    const isTotal = e.kind === 'subtotal' || e.kind === 'group';
    const style: PivotCellStyle = isGrand ? 'grand' : e.kind === 'subtotal' ? 'total' : 'value';
    if (labelCols) {
      if (isGrand) row[0] = { v: rowSpecs.length ? 'Grand Total' : nV === 1 ? vName(0) : 'Total', s: rowSpecs.length ? 'grand' : 'label', bold: !!rowSpecs.length };
      else {
        const n = e.node!;
        const spec = rowSpecs[n.depth];
        const fieldName = spec && !spec.group ? spec.base : undefined;
        if (layout === 'compact') row[0] = { v: itemLabel(n), s: 'label', indent: n.depth, field: fieldName, bold: e.kind === 'group' };
        else if (e.kind === 'subtotal') row[n.depth] = { v: `${itemText(n.item)} Total`, s: 'total', bold: true };
        else {
          row[n.depth] = { v: itemLabel(n), s: 'label', field: fieldName, bold: e.kind === 'group' };
          if (layout === 'tabular') {
            // Ancestors are shown on the first row of their group.
            let a: TreeNode = n;
            while (a.parent && a.parent.depth >= 0 && a.parent.children[0] === a) {
              a = a.parent;
              const sp = rowSpecs[a.depth];
              row[a.depth] = { v: itemLabel(a), s: 'label', field: sp && !sp.group ? sp.base : undefined };
            }
          }
        }
      }
    }
    const showValues = e.kind !== 'group' || def.subtotals;
    if (showValues && nV) {
      slots.forEach((slot, si) => {
        for (let vi = 0; vi < nV; vi++) {
          const c = labelCols + si * perSlot + vi;
          const x = shown(e, slot, vi);
          const totalish = isGrand || slot.kind === 'grand';
          row[c] = { v: x.v, s: totalish ? 'grand' : style, vi, pct: x.pct, bold: isTotal || slot.kind === 'subtotal' };
          drill.set(`${r},${c}`, { row: e.path, col: slot.path });
        }
      });
    } else if (isGrand || e.kind === 'subtotal') {
      for (let c = labelCols; c < totalCols; c++) row[c] = { v: null, s: isGrand ? 'grand' : 'total' };
    }
    if (isGrand && labelCols) for (let c = 1; c < labelCols; c++) row[c] = { v: null, s: 'grand' };
    cells.push(row);
  }
  return finish(cells, filterRows, headerRows, labelCols, drill, rowSpecs.map((s) => s.name), colSpecs.map((s) => s.name), rowsIdx.length);
}

function finish(cells: PivotCell[][], filterRows: number, headerRows: number, labelCols: number, drill: PivotResult['drill'], rowFields: string[], colFields: string[], records: number): PivotResult {
  const cols = Math.max(1, ...cells.map((r) => r.length));
  return { cells, rows: cells.length, cols, filterRows, headerRows, labelCols, drill, rowFields, colFields, records };
}

/** Source row indexes behind a value cell (row and column item key paths, from PivotResult.drill). */
export function drillRows(def: PivotDef, source: PivotSource, row: string[], col: string[], opts: { date1904?: boolean } = {}): number[] {
  const rowSpecs = expandFields(def, def.rows, source.headers);
  const colSpecs = expandFields(def, def.cols, source.headers);
  return filteredRows(def, source).filter((i) => {
    const r = source.rows[i];
    for (let k = 0; k < row.length; k++) if (itemKey(r[rowSpecs[k].col] ?? null, rowSpecs[k].group, opts.date1904).key !== row[k]) return false;
    for (let k = 0; k < col.length; k++) if (itemKey(r[colSpecs[k].col] ?? null, colSpecs[k].group, opts.date1904).key !== col[k]) return false;
    return true;
  });
}

/** Plain value of a pivot cell. */
export function cellValue(c: PivotCell | undefined): Scalar {
  if (c == null) return null;
  if (typeof c === 'object' && !(c instanceof CellError)) return c.v;
  return c as Scalar;
}

/** Guesses a sensible field placement: numbers go to Values, everything else to Rows. */
export function isNumericField(source: PivotSource, field: string): boolean {
  const col = findField(source.headers, field);
  if (col < 0) return false;
  let nums = 0, other = 0;
  for (const r of source.rows.slice(0, 500)) { const v = r[col]; if (typeof v === 'number') nums++; else if (v != null && v !== '') other++; }
  return nums > 0 && nums >= other * 3;
}

/** Normalizes a header row: blanks become "ColumnN", duplicates get a numeric suffix. */
export function normalizeHeaders(raw: Scalar[]): string[] {
  const out: string[] = [];
  raw.forEach((h, i) => {
    let name = h == null || h === '' ? `Column${i + 1}` : h instanceof CellError ? h.code : String(h);
    if (out.some((x) => x.toLowerCase() === name.toLowerCase())) {
      let k = 2;
      while (out.some((x) => x.toLowerCase() === `${name}${k}`.toLowerCase())) k++;
      name = `${name}${k}`;
    }
    out.push(name);
  });
  return out;
}
