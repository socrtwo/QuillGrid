/**
 * PivotTables: create (and recommended), the PivotTable Fields pane, refresh, options, layout, totals,
 * change source, clear, move, delete and drill-down. The computation lives in src/model/pivot.ts; this
 * module reads the source data, writes the computed grid to the output sheet and keeps PivotDef in
 * sheet.pivots of the sheet that holds the output.
 */
import { def, getCommand } from './registry';
import type { App } from '../app';
import { Sheet, newId, type Workbook } from '../model/workbook';
import type { PivotDef, CellStyle, Scalar } from '../model/types';
import { CellError } from '../model/types';
import {
  computePivot, drillRows, isNumericField, normalizeHeaders, pivotItems, valueName, cellValue, findField,
  FN_LABELS, SHOW_AS_LABELS, GROUP_LABEL, type PivotSource, type PivotResult, type PivotCellObj, type PivotFn, type ShowAs, type DateGroup,
} from '../model/pivot';
import { openDialog, toast, showMenu, confirmDialog, alertDialog, type DialogControl } from '../ui/dialog';
import { h, icon, clear } from '../ui/dom';
import { registerPane, typingIn } from '../ui/panes';
import { activeCell, dataRegion } from './ops';
import { resolveRangeText, rangeRefText } from '../model/structure';
import { cellName, rangeName, intersects, quoteSheet, absRangeName, type Range } from '../model/address';
import { formatValue, isDateFormat } from '../model/numfmt';

const TAB = 'PivotTable Analyze';

// ============================================================================================
// Source data

export interface SourceData extends PivotSource { sheet: Sheet; range: Range; formats: (string | undefined)[] }

/** Reads a PivotTable source: a table name or a range reference whose first row holds the headers. */
export function readSource(book: Workbook, text: string, fallback: Sheet): SourceData | string {
  const t = text.trim().replace(/^=/, '');
  if (!t) return 'Choose the data to analyze.';
  let sheet: Sheet | null = null;
  let range: Range | null = null;
  let lastRow: number | null = null;
  for (const s of book.sheets) {
    const tb = s.tables.find((x) => x.name.toLowerCase() === t.toLowerCase());
    if (tb) { sheet = s; range = tb.range; if (tb.opts.totals) lastRow = tb.range.r2 - 1; break; }
  }
  if (!sheet) {
    const res = resolveRangeText(book, t, fallback);
    if (!res) return `"${t}" is not a valid reference or table name.`;
    sheet = res.sheet; range = res.range;
  }
  const used = sheet.usedRange(false);
  if (!used) return 'The source range is empty.';
  const rg: Range = { r1: range!.r1, c1: range!.c1, r2: Math.min(lastRow ?? range!.r2, used.r2), c2: Math.min(range!.c2, used.c2) };
  if (rg.r2 <= rg.r1 || rg.c2 < rg.c1) return 'The source needs a header row and at least one row of data.';
  const rawHeaders: Scalar[] = [];
  for (let c = rg.c1; c <= rg.c2; c++) rawHeaders.push(sheet.value(rg.r1, c));
  if (rawHeaders.every((x) => x == null || x === '')) return 'The PivotTable field names are not valid. The first row of the source must contain column headings.';
  const headers = normalizeHeaders(rawHeaders);
  const rows: Scalar[][] = [];
  for (let r = rg.r1 + 1; r <= rg.r2; r++) { const row: Scalar[] = []; for (let c = rg.c1; c <= rg.c2; c++) row.push(sheet.value(r, c)); rows.push(row); }
  const formats = headers.map((_, i) => sheet!.styleOf(rg.r1 + 1, rg.c1 + i).numFmt);
  return { sheet, range: rg, headers, rows, formats };
}

function defaultSourceText(app: App): string {
  const s = app.sheet;
  const { r, c } = activeCell(app);
  const t = s.tables.find((x) => r >= x.range.r1 && r <= x.range.r2 && c >= x.range.c1 && c <= x.range.c2);
  if (t) return t.name;
  const rg = dataRegion(app);
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && s.value(rg.r1, rg.c1) == null) return '';
  return rangeRefText(s, rg, true);
}

// ============================================================================================
// Locating PivotTables

let lastPivot: { sheetId: string; id: string } | null = null;

export function pivotAt(sheet: Sheet, r: number, c: number): PivotDef | undefined {
  return sheet.pivots.find((p) => p.lastRange && r >= p.lastRange.r1 && r <= p.lastRange.r2 && c >= p.lastRange.c1 && c <= p.lastRange.c2);
}

function findPivot(book: Workbook, sheetId: string, id: string): { sheet: Sheet; pv: PivotDef } | null {
  const sheet = book.sheetById(sheetId);
  const pv = sheet?.pivots.find((p) => p.id === id);
  return sheet && pv ? { sheet, pv } : null;
}

/** The PivotTable at the active cell, else the one used most recently on this workbook. */
function currentPivot(app: App, quiet = false): { sheet: Sheet; pv: PivotDef } | null {
  const s = app.sheet;
  const { r, c } = activeCell(app);
  const at = pivotAt(s, r, c);
  if (at) { lastPivot = { sheetId: s.id, id: at.id }; return { sheet: s, pv: at }; }
  const prev = lastPivot && findPivot(app.book, lastPivot.sheetId, lastPivot.id);
  if (prev && prev.sheet === s) return prev;
  if (!quiet) toast('Select a cell in a PivotTable first.');
  return null;
}

export function uniquePivotName(book: Workbook): string {
  const taken = new Set(book.sheets.flatMap((s) => s.pivots.map((p) => (p.name || '').toLowerCase())));
  let i = 1;
  while (taken.has(`pivottable${i}`)) i++;
  return `PivotTable${i}`;
}

// ============================================================================================
// Writing the output

const HEADER_FILL = '#dde7f3';
const LINE = '#8eaadb';

function styleFor(book: Workbook, cell: PivotCellObj, pv: PivotDef, src: SourceData): number {
  const st: CellStyle = {};
  switch (cell.s) {
    case 'header': st.bold = true; st.fill = HEADER_FILL; st.bb = { style: 'thin', color: LINE }; break;
    case 'caption': st.bold = true; break;
    case 'filter': st.bold = true; break;
    case 'grand': st.bold = true; st.fill = HEADER_FILL; st.bt = { style: 'thin', color: LINE }; break;
    case 'total': st.bold = true; break;
    case 'label': if (cell.indent) st.indent = cell.indent; if (cell.bold) st.bold = true; break;
    case 'value': if (cell.bold) st.bold = true; break;
  }
  if (cell.field) {
    const i = findField(src.headers, cell.field);
    const f = i >= 0 ? src.formats[i] : undefined;
    if (f) st.numFmt = f;
    if (cell.s === 'label' || cell.s === 'header') st.hAlign = 'left';
  }
  if (cell.vi != null) {
    const vf = pv.values.filter((v) => findField(src.headers, v.field) >= 0)[cell.vi];
    if (vf?.numFmt) st.numFmt = vf.numFmt;
    else if (cell.pct) st.numFmt = '0.00%';
    else if (vf && vf.fn !== 'count' && vf.fn !== 'countNums' && (vf.showAs || 'normal') !== 'rank') {
      const i = findField(src.headers, vf.field);
      const f = i >= 0 ? src.formats[i] : undefined;
      if (f && !isDateFormat(f)) st.numFmt = f;
    }
  }
  return book.styleId(st);
}

function areaOf(pv: PivotDef, res: PivotResult): Range {
  return { r1: pv.target.r, c1: pv.target.c, r2: pv.target.r + res.rows - 1, c2: pv.target.c + res.cols - 1 };
}

/** Clears the old output and writes the result. Call inside a transaction after recordSheet(sheet). */
function writePivot(app: App, sheet: Sheet, pv: PivotDef, res: PivotResult, src: SourceData): void {
  const book = app.book;
  if (pv.lastRange) {
    const old = pv.lastRange;
    const keys: [number, number][] = [];
    sheet.forEachIn(old, (r, c) => keys.push([r, c]));
    for (const [r, c] of keys) sheet.set(r, c, undefined);
  }
  const area = areaOf(pv, res);
  for (let i = 0; i < res.cells.length; i++) {
    const row = res.cells[i];
    for (let j = 0; j < res.cols; j++) {
      const x = row[j];
      const r = area.r1 + i, c = area.c1 + j;
      if (x == null) { if (sheet.get(r, c)) sheet.set(r, c, undefined); continue; }
      const obj: PivotCellObj = typeof x === 'object' && !(x instanceof CellError) ? x : { v: x as Scalar };
      const s = styleFor(book, obj, pv, src);
      const v = obj.v;
      sheet.set(r, c, v == null && !s ? undefined : { v: v ?? undefined, ...(s ? { s } : {}) } as any);
    }
  }
  pv.lastRange = area;
  // Fit the columns to their contents (only widening).
  for (let j = 0; j < res.cols; j++) {
    let w = 0;
    for (let i = 0; i < res.rows; i++) {
      const x = res.cells[i]?.[j];
      const v = cellValue(x);
      if (v == null) continue;
      const obj = typeof x === 'object' && x && !(x instanceof CellError) ? (x as PivotCellObj) : null;
      if (obj?.s === 'caption' && res.rows <= res.filterRows + 1) continue;
      const t = typeof v === 'number' ? formatValue(v, obj?.pct ? '0.00%' : undefined).text : String(v instanceof CellError ? v.code : v);
      w = Math.max(w, t.length * 7.2 + 14 + (obj?.indent || 0) * 10 + (obj?.s === 'header' && i === res.filterRows ? 16 : 0));
    }
    const c = area.c1 + j;
    const want = Math.round(Math.min(320, Math.max(sheet.defaultColW, w)));
    if (want > sheet.colWidth(c)) sheet.colW.set(c, want);
  }
}

/** Data cells (not part of this PivotTable) that a new output area would overwrite. */
function overlap(sheet: Sheet, pv: PivotDef, area: Range): { data: boolean; pivot: PivotDef | null } {
  const other = sheet.pivots.find((p) => p !== pv && p.lastRange && intersects(p.lastRange, area)) || null;
  let data = false;
  sheet.forEachIn(area, (r, c, cell) => {
    if (data) return;
    if (pv.lastRange && r >= pv.lastRange.r1 && r <= pv.lastRange.r2 && c >= pv.lastRange.c1 && c <= pv.lastRange.c2) return;
    if (cell.v != null && cell.v !== '' || cell.f != null) data = true;
  });
  return { data, pivot: other };
}

function guard(app: App, sheet: Sheet): boolean {
  if (app.isReadOnly()) { toast('This workbook is read-only.'); return false; }
  if (sheet.protection.on && !sheet.protection.allow.includes('pivotTables')) { toast('The sheet is protected. Unprotect it to change the PivotTable.'); return false; }
  return true;
}

/**
 * Applies an optional change to a PivotTable definition, recomputes it from its source and rewrites
 * the output (one undoable step). Asks before overwriting other data.
 */
export async function applyPivot(app: App, sheet: Sheet, pv: PivotDef, mutate?: (d: PivotDef) => void, label = 'PivotTable'): Promise<boolean> {
  if (!guard(app, sheet)) return false;
  const draft: PivotDef = JSON.parse(JSON.stringify(pv));
  mutate?.(draft);
  const src = readSource(app.book, draft.source, sheet);
  if (typeof src === 'string') { toast(src); return false; }
  const res = computePivot(draft, src, { date1904: app.book.calc.date1904 });
  const area = areaOf(draft, res);
  const ov = overlap(sheet, pv, area);
  if (ov.pivot) { await alertDialog('PivotTable', 'A PivotTable cannot overlap another PivotTable. Move one of them first.'); return false; }
  if (ov.data && !(await confirmDialog('PivotTable', `There is already data in ${quoteSheet(sheet.name)}!${rangeName(area)}. Do you want to replace it?`, 'Replace'))) return false;
  app.book.transact(label, () => {
    app.book.recordSheet(sheet);
    for (const k of Object.keys(pv) as (keyof PivotDef)[]) delete (pv as any)[k];
    Object.assign(pv, draft);
    writePivot(app, sheet, pv, res, src);
  });
  lastPivot = { sheetId: sheet.id, id: pv.id };
  return true;
}

// ============================================================================================
// Create

function newDef(book: Workbook, source: string, target: { sheet: string; r: number; c: number }): PivotDef {
  return { id: newId('pv'), name: uniquePivotName(book), source, target, rows: [], cols: [], filters: [], values: [], layout: 'compact', grandRows: true, grandCols: true, subtotals: true };
}

/** Creates a PivotTable (on a new sheet or at a location) and renders it. */
function createPivot(app: App, pvInit: Partial<PivotDef>, source: string, dest: { newSheet: true } | { sheet: Sheet; r: number; c: number }): { sheet: Sheet; pv: PivotDef } | null {
  const book = app.book;
  const src = readSource(book, source, app.sheet);
  if (typeof src === 'string') { toast(src); return null; }
  let target: Sheet;
  let r = 2, c = 0;
  let pv!: PivotDef;
  book.transact('Insert PivotTable', () => {
    if ('newSheet' in dest) {
      book.recordBook();
      target = new Sheet(book, book.uniqueSheetName());
      book.sheets.splice(book.activeSheet, 0, target);
    } else { target = dest.sheet; r = dest.r; c = dest.c; }
    book.recordSheet(target);
    pv = { ...newDef(book, source, { sheet: target.name, r, c }), ...pvInit, id: newId('pv'), target: { sheet: target.name, r, c } };
    if (!pv.name) pv.name = uniquePivotName(book);
    target.pivots.push(pv);
    writePivot(app, target, pv, computePivot(pv, src, { date1904: book.calc.date1904 }), src);
  });
  const idx = book.sheets.indexOf(target!);
  app.exec('sheet.activate', idx);
  app.grid.select(r, c);
  lastPivot = { sheetId: target!.id, id: pv.id };
  return { sheet: target!, pv };
}

async function createDialog(app: App, arg?: { source?: string; dest?: 'new' | 'existing'; location?: string } & Partial<PivotDef>): Promise<void> {
  if (app.isReadOnly()) { toast('This workbook is read-only.'); return; }
  const book = app.book;
  // Direct (macro) use: everything given.
  if (arg?.source && arg.dest) {
    const dest = arg.dest === 'existing' && arg.location ? resolveRangeText(book, arg.location, app.sheet) : null;
    const { source, dest: _d, location: _l, ...init } = arg;
    void _d; void _l;
    createPivot(app, init, source, dest ? { sheet: dest.sheet, r: dest.range.r1, c: dest.range.c1 } : { newSheet: true });
    return;
  }
  const r = await openDialog({
    title: 'Create PivotTable',
    fields: [
      { key: 'h1', type: 'heading', label: 'Choose the data that you want to analyze' },
      { key: 'source', label: 'Table or range', type: 'text', value: arg?.source || defaultSourceText(app), autofocus: true, help: 'A table name or a range such as Sheet1!$A$1:$D$100. The first row must contain column headings.' },
      { key: 'h2', type: 'heading', label: 'Choose where to place the PivotTable' },
      { key: 'dest', type: 'radio', options: [{ value: 'new', label: 'New Worksheet' }, { value: 'existing', label: 'Existing Worksheet' }], value: 'new' },
      { key: 'location', label: 'Location', type: 'text', value: `${quoteSheet(app.sheet.name)}!${cellName(app.sheet.sel.r, app.sheet.sel.c, true, true)}`, visibleIf: (v) => v.dest === 'existing' },
    ],
    validate: (v) => {
      const s = readSource(book, String(v.source || ''), app.sheet);
      if (typeof s === 'string') return s;
      if (v.dest === 'existing' && !resolveRangeText(book, String(v.location || ''), app.sheet)) return 'The destination reference is not valid.';
      return null;
    },
  });
  if (!r) return;
  let dest: { newSheet: true } | { sheet: Sheet; r: number; c: number } = { newSheet: true };
  if (r.dest === 'existing') {
    const loc = resolveRangeText(book, String(r.location), app.sheet)!;
    const area = { r1: loc.range.r1, c1: loc.range.c1, r2: loc.range.r1 + 2, c2: loc.range.c1 + 2 };
    if (loc.sheet.pivots.some((p) => p.lastRange && intersects(p.lastRange, area))) { toast('A PivotTable cannot overlap another PivotTable.'); return; }
    if (!guard(app, loc.sheet)) return;
    dest = { sheet: loc.sheet, r: loc.range.r1, c: loc.range.c1 };
  }
  const made = createPivot(app, {}, String(r.source).trim(), dest);
  if (made) app.ui.openPane('pivotFields', { sheetId: made.sheet.id, id: made.pv.id });
}

// ============================================================================================
// Recommended PivotTables

export function recommendPivots(src: PivotSource): { title: string; def: Partial<PivotDef> }[] {
  const cats: string[] = [], nums: string[] = [];
  for (const hd of src.headers) {
    if (isNumericField(src, hd)) nums.push(hd);
    else {
      const items = pivotItems(src, hd);
      if (items.length >= 1 && items.length <= Math.max(50, src.rows.length / 2)) cats.push(hd);
    }
  }
  const out: { title: string; def: Partial<PivotDef> }[] = [];
  const add = (title: string, d: Partial<PivotDef>) => { if (out.length < 6 && !out.some((x) => x.title === title)) out.push({ title, def: d }); };
  for (const cat of cats.slice(0, 3)) for (const n of nums.slice(0, 2)) add(`Sum of ${n} by ${cat}`, { rows: [cat], values: [{ field: n, fn: 'sum' }] });
  if (cats.length >= 2 && nums.length) add(`Sum of ${nums[0]} by ${cats[0]} and ${cats[1]}`, { rows: [cats[0]], cols: [cats[1]], values: [{ field: nums[0], fn: 'sum' }] });
  for (const cat of cats.slice(0, 2)) add(`Count of ${cat} by ${cat}`, { rows: [cat], values: [{ field: cat, fn: 'count' }] });
  if (cats.length && nums.length) add(`Average of ${nums[0]} by ${cats[0]}`, { rows: [cats[0]], values: [{ field: nums[0], fn: 'average' }] });
  if (!cats.length && nums.length) for (const n of nums.slice(0, 3)) add(`Sum of ${n}`, { values: [{ field: n, fn: 'sum' }] });
  return out;
}

function previewTable(res: PivotResult, maxRows = 14): HTMLElement {
  const tbl = h('table', { class: 'ql-pvt-preview' });
  res.cells.slice(0, maxRows).forEach((row) => {
    const tr = h('tr');
    for (let j = 0; j < Math.min(res.cols, 8); j++) {
      const x = row[j];
      const v = cellValue(x);
      const obj = typeof x === 'object' && x && !(x instanceof CellError) ? (x as PivotCellObj) : null;
      const text = v == null ? '' : typeof v === 'number' ? formatValue(v, obj?.pct ? '0.0%' : Number.isInteger(v) ? '#,##0' : '#,##0.00').text : String(v instanceof CellError ? v.code : v);
      tr.appendChild(h('td', { class: (obj?.s ? 'ql-pvt-' + obj.s : '') + (typeof v === 'number' && obj?.vi != null ? ' ql-num' : ''), style: obj?.indent ? `padding-left:${obj.indent * 12 + 4}px` : '' }, text));
    }
    tbl.appendChild(tr);
  });
  return tbl;
}

async function recommendedDialog(app: App, arg?: { source?: string }): Promise<void> {
  if (app.isReadOnly()) { toast('This workbook is read-only.'); return; }
  const sourceText = arg?.source || defaultSourceText(app);
  const src = readSource(app.book, sourceText, app.sheet);
  if (typeof src === 'string') { toast(src); return; }
  const recs = recommendPivots(src);
  if (!recs.length) { toast('No recommendations for this data. Use PivotTable to build one yourself.'); return createDialog(app, { source: sourceText }); }
  let chosen = 0;
  const content = h('div', { class: 'ql-pvt-rec' });
  const list = h('div', { class: 'ql-pvt-rec-list', role: 'listbox' });
  const preview = h('div', { class: 'ql-pvt-rec-preview' });
  const base = (d: Partial<PivotDef>): PivotDef => ({ ...newDef(app.book, sourceText, { sheet: '', r: 0, c: 0 }), ...d });
  const render = () => {
    clear(list);
    recs.forEach((rc, i) => {
      list.appendChild(h('button', { class: 'ql-pvt-rec-item' + (i === chosen ? ' ql-on' : ''), type: 'button', role: 'option', 'aria-selected': String(i === chosen), onclick: () => { chosen = i; render(); }, ondblclick: () => ctl?.close({ ok: true }) }, icon('pivot_table_chart'), ' ', rc.title));
    });
    clear(preview);
    preview.append(h('div', { class: 'ql-label' }, recs[chosen].title), previewTable(computePivot(base(recs[chosen].def), src, { date1904: app.book.calc.date1904 })));
  };
  let ctl: DialogControl | null = null;
  content.append(list, preview);
  render();
  const r = await openDialog({
    title: 'Recommended PivotTables', content, wide: true, ok: 'OK',
    onOpen: (c) => { ctl = c; },
    buttons: [{ label: 'Blank PivotTable', onClick: (_v, c) => { c.close({ blank: true }); } }],
  });
  if (!r) return;
  if (r.blank) { await createDialog(app, { source: sourceText }); return; }
  const made = createPivot(app, recs[chosen].def, sourceText, { newSheet: true });
  if (made) app.ui.openPane('pivotFields', { sheetId: made.sheet.id, id: made.pv.id });
}

// ============================================================================================
// Field settings dialogs

const FN_LIST = Object.keys(FN_LABELS) as PivotFn[];
const SHOW_AS_LIST = Object.keys(SHOW_AS_LABELS) as ShowAs[];

async function valueSettings(app: App, sheet: Sheet, pv: PivotDef, index: number): Promise<void> {
  const v = pv.values[index];
  if (!v) return;
  const r = await openDialog({
    title: 'Value Field Settings',
    fields: [
      { key: 'info', type: 'info', label: `Source name: ${v.field}` },
      { key: 'name', label: 'Custom name', type: 'text', value: valueName(v) },
      { key: 'fn', label: 'Summarize value field by', type: 'list', rows: 6, options: FN_LIST.map((f) => ({ value: f, label: FN_LABELS[f] })), value: v.fn },
      { key: 'showAs', label: 'Show values as', type: 'select', options: SHOW_AS_LIST.map((s) => ({ value: s, label: SHOW_AS_LABELS[s] })), value: v.showAs || 'normal' },
      { key: 'numFmt', label: 'Number format', type: 'combo', options: ['General', '0', '0.00', '#,##0', '#,##0.00', '$#,##0.00', '0%', '0.00%', 'm/d/yyyy'], value: v.numFmt || '', placeholder: 'Automatic' },
    ],
    onChange: (vals, ctl, key) => {
      if (key === 'fn') { const auto = `${FN_LABELS[v.fn]} of ${v.field}`; const cur = String(ctl.get('name') || ''); if (!cur || cur === auto || /^(Sum|Count|Average|Max|Min|Product|Count Numbers|StdDev|StdDevp|Var|Varp) of /.test(cur)) ctl.set('name', `${FN_LABELS[vals.fn as PivotFn]} of ${v.field}`); }
    },
  });
  if (!r) return;
  await applyPivot(app, sheet, pv, (d) => {
    const x = d.values[index];
    x.fn = (r.fn || x.fn) as PivotFn;
    x.showAs = r.showAs === 'normal' ? undefined : r.showAs;
    const auto = `${FN_LABELS[x.fn]} of ${x.field}`;
    x.name = r.name && r.name !== auto ? String(r.name) : undefined;
    x.numFmt = r.numFmt && r.numFmt !== 'General' ? String(r.numFmt) : undefined;
  }, 'Value Field Settings');
}

async function filterItemsDialog(app: App, sheet: Sheet, pv: PivotDef, field: string): Promise<void> {
  const src = readSource(app.book, pv.source, sheet);
  if (typeof src === 'string') { toast(src); return; }
  const items = pivotItems(src, field);
  const cur = pv.filterValues?.[field];
  const r = await openDialog({
    title: `Filter: ${field}`,
    help: 'Select the items to show. Select all items to remove the filter.',
    fields: [
      { key: 'items', label: 'Items', type: 'list', multiple: true, rows: Math.min(14, Math.max(5, items.length)), options: items, value: cur || items },
      { key: 'all', label: 'Select All', type: 'button', onClick: (_v, ctl) => ctl.set('items', items) },
    ],
    validate: (v) => (!v.items?.length ? 'Select at least one item.' : null),
  });
  if (!r) return;
  const sel: string[] = r.items || [];
  await applyPivot(app, sheet, pv, (d) => {
    d.filterValues ||= {};
    if (sel.length === items.length) delete d.filterValues[field]; else d.filterValues[field] = sel;
  }, 'Filter PivotTable');
}

async function groupDatesDialog(app: App, sheet: Sheet, pv: PivotDef, field: string): Promise<void> {
  const cur = pv.dateGroups?.[field] || [];
  const r = await openDialog({
    title: `Grouping: ${field}`, help: 'Group dates by:',
    fields: (['years', 'quarters', 'months', 'days'] as DateGroup[]).map((g) => ({ key: g, label: GROUP_LABEL[g], type: 'checkbox' as const, value: cur.includes(g) })),
  });
  if (!r) return;
  const groups = (['years', 'quarters', 'months', 'days'] as DateGroup[]).filter((g) => r[g]);
  await applyPivot(app, sheet, pv, (d) => { d.dateGroups ||= {}; if (groups.length) d.dateGroups[field] = groups; else delete d.dateGroups[field]; }, 'Group Dates');
}

// ============================================================================================
// PivotTable Fields pane

type Area = 'filters' | 'cols' | 'rows' | 'values';
const AREA_LABEL: Record<Area, string> = { filters: 'Filters', cols: 'Columns', rows: 'Rows', values: 'Values' };
const AREA_ICON: Record<Area, string> = { filters: 'filter_alt', cols: 'view_column', rows: 'table_rows', values: 'functions' };

let deferUpdate = false;
let pendingDef: { key: string; def: PivotDef } | null = null;

/** Moves a field to an area (pure on a definition). index = position in the target list. */
export function placeField(d: PivotDef, field: string, to: Area | null, opts: { from?: Area; fromIndex?: number; index?: number; fn?: PivotFn } = {}): void {
  if (opts.from === 'values' && opts.fromIndex != null) d.values.splice(opts.fromIndex, 1);
  else if (opts.from && opts.from !== 'values') d[opts.from] = d[opts.from].filter((x) => x !== field);
  if (to === null) {
    // Unchecking a field removes it from every area.
    d.rows = d.rows.filter((x) => x !== field); d.cols = d.cols.filter((x) => x !== field); d.filters = d.filters.filter((x) => x !== field);
    d.values = d.values.filter((x) => x.field !== field);
    if (d.filterValues) delete d.filterValues[field];
    return;
  }
  if (to === 'values') {
    const item = { field, fn: opts.fn || 'sum' as PivotFn };
    d.values.splice(opts.index ?? d.values.length, 0, item);
    return;
  }
  // A field can be in only one of rows / columns / filters.
  for (const a of ['rows', 'cols', 'filters'] as const) d[a] = d[a].filter((x) => x !== field);
  d[to].splice(Math.min(opts.index ?? d[to].length, d[to].length), 0, field);
}

function paneTarget(app: App, opts: any): { sheet: Sheet; pv: PivotDef } | null {
  const s = app.sheet;
  const at = pivotAt(s, s.sel.r, s.sel.c);
  if (at) { lastPivot = { sheetId: s.id, id: at.id }; return { sheet: s, pv: at }; }
  if (opts?.sheetId && opts?.id) { const f = findPivot(app.book, opts.sheetId, opts.id); if (f && f.sheet === s) return f; }
  if (lastPivot) { const f = findPivot(app.book, lastPivot.sheetId, lastPivot.id); if (f && f.sheet === s) return f; }
  return null;
}

/** True while a text box in the pane has focus (checkboxes do not block redraws). */
function editingIn(el: HTMLElement): boolean {
  const a = document.activeElement as HTMLInputElement | null;
  return !!a && el.contains(a) && typingIn(el) && !(a.tagName === 'INPUT' && (a.type === 'checkbox' || a.type === 'radio'));
}

function renderFieldsPane(app: App, el: HTMLElement, opts: any): void {
  const t = paneTarget(app, opts);
  const sig = JSON.stringify([app.tab.id, app.book.version, t?.sheet.id, t?.pv, deferUpdate, pendingDef?.key]);
  if (el.dataset.sig === sig || editingIn(el)) return;
  const search = el.querySelector<HTMLInputElement>('.ql-pvt-search')?.value || '';
  el.dataset.sig = sig;
  clear(el);
  if (!t) {
    el.append(h('p', { class: 'ql-help' }, 'Select a cell in a PivotTable to choose its fields.'),
      h('button', { class: 'ql-btn', type: 'button', onclick: () => app.exec('pivot.create') }, icon('pivot_table_chart'), ' Insert PivotTable…'));
    return;
  }
  const { sheet, pv } = t;
  const key = sheet.id + ':' + pv.id;
  if (pendingDef && pendingDef.key !== key) pendingDef = null;
  const d: PivotDef = deferUpdate && pendingDef ? pendingDef.def : pv;
  const src = readSource(app.book, pv.source, sheet);
  const change = (fn: (x: PivotDef) => void, label = 'PivotTable Fields') => {
    if (deferUpdate) {
      if (!pendingDef) pendingDef = { key, def: JSON.parse(JSON.stringify(pv)) };
      fn(pendingDef.def);
      el.dataset.sig = '';
      renderFieldsPane(app, el, opts);
      return;
    }
    void applyPivot(app, sheet, pv, fn, label);
  };
  el.appendChild(h('div', { class: 'ql-pvt-head' }, h('b', null, pv.name || 'PivotTable'), h('span', { class: 'ql-help' }, ` · ${pv.source}`)));
  if (typeof src === 'string') { el.append(h('p', { class: 'ql-warn' }, src), h('button', { class: 'ql-btn', type: 'button', onclick: () => app.exec('pivot.changeSource') }, 'Change Data Source…')); return; }
  const used = new Set([...d.rows, ...d.cols, ...d.filters, ...d.values.map((v) => v.field)].map((x) => x.toLowerCase()));
  const searchBox = h('input', { class: 'ql-input ql-pvt-search', type: 'search', placeholder: 'Search fields', value: search, 'aria-label': 'Search fields' }) as HTMLInputElement;
  const fieldList = h('div', { class: 'ql-pvt-fields' });
  const drawFields = () => {
    clear(fieldList);
    const q = searchBox.value.trim().toLowerCase();
    for (const f of src.headers) {
      if (q && !f.toLowerCase().includes(q)) continue;
      const cb = h('input', { type: 'checkbox', checked: used.has(f.toLowerCase()), 'aria-label': f }) as HTMLInputElement;
      cb.addEventListener('change', () => change((x) => {
        if (cb.checked) placeField(x, f, isNumericField(src, f) ? 'values' : 'rows', { fn: 'sum' });
        else placeField(x, f, null);
      }));
      const more = h('button', { class: 'ql-iconbtn ql-sm', type: 'button', 'aria-label': `More options for ${f}`, title: 'Add to…' }, icon('more_vert'));
      more.addEventListener('click', () => showMenu([
        { id: 'filters', label: 'Add to Filters', icon: AREA_ICON.filters }, { id: 'rows', label: 'Add to Rows', icon: AREA_ICON.rows },
        { id: 'cols', label: 'Add to Columns', icon: AREA_ICON.cols }, { id: 'values', label: 'Add to Values', icon: AREA_ICON.values },
        { separator: true }, { id: 'filter', label: 'Filter Items…', icon: 'filter_list' },
        ...(src.formats[findField(src.headers, f)] && isDateFormat(src.formats[findField(src.headers, f)]) ? [{ id: 'group', label: 'Group Dates…', icon: 'date_range' }] : []),
      ], {
        anchor: more, onPick: (e) => {
          if (e.id === 'filter') void filterItemsDialog(app, sheet, pv, f);
          else if (e.id === 'group') void groupDatesDialog(app, sheet, pv, f);
          else if (e.id) change((x) => placeField(x, f, e.id as Area, { fn: isNumericField(src, f) ? 'sum' : 'count' }));
        },
      }));
      const row = h('label', { class: 'ql-pvt-field', draggable: 'true' }, cb, h('span', { class: 'ql-pvt-fname' }, f), d.filterValues?.[f] ? icon('filter_alt', 'ql-pvt-filtered') : null, more);
      row.addEventListener('dragstart', (e) => { e.dataTransfer?.setData('text/plain', JSON.stringify({ field: f })); });
      fieldList.appendChild(row);
    }
  };
  searchBox.addEventListener('input', drawFields);
  drawFields();
  el.append(h('div', { class: 'ql-label' }, 'Choose fields to add to the report'), searchBox, fieldList);
  // Areas
  el.appendChild(h('div', { class: 'ql-label' }, 'Drag or move fields between the areas below'));
  const areas = h('div', { class: 'ql-pvt-areas' });
  for (const area of ['filters', 'cols', 'rows', 'values'] as Area[]) {
    const box = h('div', { class: 'ql-pvt-area', 'data-area': area }, h('div', { class: 'ql-pvt-area-h' }, icon(AREA_ICON[area]), ' ', AREA_LABEL[area]));
    box.addEventListener('dragover', (e) => { e.preventDefault(); box.classList.add('ql-drop'); });
    box.addEventListener('dragleave', () => box.classList.remove('ql-drop'));
    box.addEventListener('drop', (e) => {
      e.preventDefault(); box.classList.remove('ql-drop');
      try {
        const info = JSON.parse(e.dataTransfer?.getData('text/plain') || '{}');
        if (!info.field) return;
        change((x) => placeField(x, info.field, area, { from: info.from, fromIndex: info.index, fn: info.fn || (isNumericField(src, info.field) ? 'sum' : 'count') }));
      } catch { /* ignore */ }
    });
    const items: { label: string; field: string; i: number }[] = area === 'values' ? d.values.map((v, i) => ({ label: valueName(v), field: v.field, i })) : d[area].map((f, i) => ({ label: f, field: f, i }));
    if (area === 'cols' && d.values.length > 1) items.push({ label: 'Σ Values', field: '', i: -1 });
    for (const it of items) {
      if (it.i < 0) { box.appendChild(h('div', { class: 'ql-pvt-chip ql-pvt-chip-fixed' }, it.label)); continue; }
      const list = area === 'values' ? d.values : d[area];
      const chip = h('div', { class: 'ql-pvt-chip', draggable: 'true' },
        h('span', { class: 'ql-pvt-chip-label' }, it.label),
        h('button', { class: 'ql-iconbtn ql-sm', type: 'button', 'aria-label': 'Move up', disabled: it.i === 0, onclick: () => change((x) => { const l: any[] = area === 'values' ? x.values : x[area]; const [m] = l.splice(it.i, 1); l.splice(it.i - 1, 0, m); }) }, icon('arrow_upward')),
        h('button', { class: 'ql-iconbtn ql-sm', type: 'button', 'aria-label': 'Move down', disabled: it.i >= list.length - 1, onclick: () => change((x) => { const l: any[] = area === 'values' ? x.values : x[area]; const [m] = l.splice(it.i, 1); l.splice(it.i + 1, 0, m); }) }, icon('arrow_downward')));
      const menuBtn = h('button', { class: 'ql-iconbtn ql-sm', type: 'button', 'aria-label': `Options for ${it.label}` }, icon('arrow_drop_down'));
      menuBtn.addEventListener('click', () => {
        const entries = [
          ...(['filters', 'rows', 'cols', 'values'] as Area[]).filter((a) => a !== area).map((a) => ({ id: 'move:' + a, label: `Move to ${AREA_LABEL[a]}`, icon: AREA_ICON[a] })),
          { separator: true },
          ...(area === 'values' ? [{ id: 'settings', label: 'Value Field Settings…', icon: 'settings' }] : [
            { id: 'filter', label: 'Filter Items…', icon: 'filter_list' },
            ...(area !== 'filters' ? [{ id: 'sortAsc', label: 'Sort A to Z', icon: 'arrow_upward' }, { id: 'sortDesc', label: 'Sort Z to A', icon: 'arrow_downward' }] : []),
            ...(area !== 'filters' && isDateFormat(src.formats[findField(src.headers, it.field)]) ? [{ id: 'group', label: 'Group Dates…', icon: 'date_range' }] : []),
          ]),
          { id: 'remove', label: 'Remove Field', icon: 'close' },
        ];
        showMenu(entries, {
          anchor: menuBtn, onPick: (e) => {
            if (e.id?.startsWith('move:')) { const to = e.id.slice(5) as Area; change((x) => placeField(x, it.field, to, { from: area, fromIndex: it.i, fn: isNumericField(src, it.field) ? 'sum' : 'count' })); }
            else if (e.id === 'settings') void valueSettings(app, sheet, pv, it.i);
            else if (e.id === 'filter') void filterItemsDialog(app, sheet, pv, it.field);
            else if (e.id === 'group') void groupDatesDialog(app, sheet, pv, it.field);
            else if (e.id === 'sortAsc' || e.id === 'sortDesc') change((x) => { x.sort ||= {}; if (e.id === 'sortDesc') x.sort[it.field] = 'desc'; else delete x.sort[it.field]; }, 'Sort PivotTable');
            else if (e.id === 'remove') change((x) => { if (area === 'values') x.values.splice(it.i, 1); else { x[area].splice(it.i, 1); } });
          },
        });
      });
      chip.appendChild(menuBtn);
      if (area === 'values') chip.addEventListener('dblclick', () => void valueSettings(app, sheet, pv, it.i));
      chip.addEventListener('dragstart', (e) => { e.stopPropagation(); e.dataTransfer?.setData('text/plain', JSON.stringify({ field: it.field, from: area, index: it.i, fn: area === 'values' ? d.values[it.i].fn : undefined })); });
      box.appendChild(chip);
    }
    if (!items.length) box.appendChild(h('div', { class: 'ql-help ql-pvt-empty' }, 'Drop fields here'));
    areas.appendChild(box);
  }
  el.appendChild(areas);
  const deferCb = h('input', { type: 'checkbox', checked: deferUpdate, id: 'ql-pvt-defer' }) as HTMLInputElement;
  deferCb.addEventListener('change', () => { deferUpdate = deferCb.checked; if (!deferUpdate && pendingDef) { const p = pendingDef.def; pendingDef = null; void applyPivot(app, sheet, pv, (x) => Object.assign(x, p), 'PivotTable Fields'); } el.dataset.sig = ''; renderFieldsPane(app, el, opts); });
  el.appendChild(h('div', { class: 'ql-row ql-pvt-foot' },
    h('label', { class: 'ql-check', for: 'ql-pvt-defer' }, deferCb, ' Defer Layout Update'),
    h('span', { class: 'ql-spacer' }),
    h('button', { class: 'ql-btn', type: 'button', disabled: !deferUpdate || !pendingDef, onclick: () => { if (!pendingDef) return; const p = pendingDef.def; pendingDef = null; void applyPivot(app, sheet, pv, (x) => Object.assign(x, p), 'PivotTable Fields'); } }, 'Update')));
}

registerPane({
  id: 'pivotFields', title: 'PivotTable Fields',
  render(app, el, opts) { el.dataset.sig = ''; renderFieldsPane(app, el, opts); },
  update(app, el, opts) { renderFieldsPane(app, el, opts); },
});

// ============================================================================================
// Other actions

async function refresh(app: App, all = false): Promise<void> {
  const list: { sheet: Sheet; pv: PivotDef }[] = [];
  if (all) for (const s of app.book.sheets) for (const pv of s.pivots) list.push({ sheet: s, pv });
  else { const t = currentPivot(app); if (t) list.push(t); }
  if (!list.length) { if (all) toast('This workbook has no PivotTables.'); return; }
  let n = 0;
  for (const t of list) if (await applyPivot(app, t.sheet, t.pv, undefined, 'Refresh PivotTable')) n++;
  if (n) app.ui.status?.(n === 1 ? 'PivotTable refreshed' : `${n} PivotTables refreshed`);
}

async function optionsDialog(app: App): Promise<void> {
  const t = currentPivot(app);
  if (!t) return;
  const { sheet, pv } = t;
  const r = await openDialog({
    title: 'PivotTable Options',
    fields: [
      { key: 'name', label: 'PivotTable name', type: 'text', value: pv.name || '' },
      { key: 'layout', label: 'Report layout', type: 'select', options: [{ value: 'compact', label: 'Compact Form' }, { value: 'outline', label: 'Outline Form' }, { value: 'tabular', label: 'Tabular Form' }], value: pv.layout },
      { key: 'subtotals', label: 'Show subtotals', type: 'checkbox', value: pv.subtotals },
      { key: 'grandRows', label: 'Show grand totals for rows', type: 'checkbox', value: pv.grandRows },
      { key: 'grandCols', label: 'Show grand totals for columns', type: 'checkbox', value: pv.grandCols },
      { key: 'source', label: 'Data source', type: 'text', value: pv.source },
    ],
    validate: (v) => {
      const name = String(v.name || '').trim();
      if (!name) return 'Type a name for the PivotTable.';
      if (app.book.sheets.some((s) => s.pivots.some((p) => p !== pv && (p.name || '').toLowerCase() === name.toLowerCase()))) return 'Another PivotTable already uses this name.';
      const s = readSource(app.book, String(v.source || ''), sheet);
      return typeof s === 'string' ? s : null;
    },
  });
  if (!r) return;
  await applyPivot(app, sheet, pv, (d) => { d.name = String(r.name).trim(); d.layout = r.layout; d.subtotals = !!r.subtotals; d.grandRows = !!r.grandRows; d.grandCols = !!r.grandCols; d.source = String(r.source).trim(); }, 'PivotTable Options');
}

async function changeSource(app: App, arg?: string): Promise<void> {
  const t = currentPivot(app);
  if (!t) return;
  let text = typeof arg === 'string' ? arg : null;
  if (text == null) {
    const r = await openDialog({ title: 'Change PivotTable Data Source', fields: [{ key: 'source', label: 'Table or range', type: 'text', value: t.pv.source, autofocus: true }], validate: (v) => { const s = readSource(app.book, String(v.source || ''), t.sheet); return typeof s === 'string' ? s : null; } });
    if (!r) return;
    text = String(r.source).trim();
  }
  const src = readSource(app.book, text, t.sheet);
  if (typeof src === 'string') { toast(src); return; }
  await applyPivot(app, t.sheet, t.pv, (d) => {
    d.source = text!;
    const has = (f: string) => findField(src.headers, f) >= 0;
    d.rows = d.rows.filter(has); d.cols = d.cols.filter(has); d.filters = d.filters.filter(has); d.values = d.values.filter((v) => has(v.field));
  }, 'Change Data Source');
}

async function clearPivot(app: App): Promise<void> {
  const t = currentPivot(app);
  if (!t) return;
  await applyPivot(app, t.sheet, t.pv, (d) => { d.rows = []; d.cols = []; d.filters = []; d.values = []; d.filterValues = {}; d.dateGroups = {}; d.sort = {}; }, 'Clear PivotTable');
}

async function movePivot(app: App, arg?: string): Promise<void> {
  const t = currentPivot(app);
  if (!t) return;
  let loc = typeof arg === 'string' ? arg : null;
  if (loc == null) {
    const r = await openDialog({ title: 'Move PivotTable', fields: [{ key: 'loc', label: 'Location (top-left cell)', type: 'text', value: `${quoteSheet(t.sheet.name)}!${cellName(t.pv.target.r, t.pv.target.c, true, true)}`, autofocus: true }], validate: (v) => (resolveRangeText(app.book, String(v.loc || ''), t.sheet) ? null : 'The reference is not valid.') });
    if (!r) return;
    loc = String(r.loc);
  }
  const dest = resolveRangeText(app.book, loc, t.sheet);
  if (!dest) { toast('The reference is not valid.'); return; }
  const { sheet, pv } = t;
  if (dest.sheet === sheet) { await applyPivot(app, sheet, pv, (d) => { d.target = { sheet: sheet.name, r: dest.range.r1, c: dest.range.c1 }; }, 'Move PivotTable'); return; }
  if (!guard(app, sheet) || !guard(app, dest.sheet)) return;
  const src = readSource(app.book, pv.source, sheet);
  if (typeof src === 'string') { toast(src); return; }
  const draft: PivotDef = { ...JSON.parse(JSON.stringify(pv)), target: { sheet: dest.sheet.name, r: dest.range.r1, c: dest.range.c1 }, lastRange: undefined };
  const res = computePivot(draft, src, { date1904: app.book.calc.date1904 });
  const ov = overlap(dest.sheet, draft, areaOf(draft, res));
  if (ov.pivot) { toast('A PivotTable cannot overlap another PivotTable.'); return; }
  if (ov.data && !(await confirmDialog('Move PivotTable', 'There is already data in the destination. Replace it?', 'Replace'))) return;
  app.book.transact('Move PivotTable', () => {
    app.book.recordSheet(sheet); app.book.recordSheet(dest.sheet);
    if (pv.lastRange) { const keys: [number, number][] = []; sheet.forEachIn(pv.lastRange, (r, c) => keys.push([r, c])); for (const [r, c] of keys) sheet.set(r, c, undefined); }
    sheet.pivots = sheet.pivots.filter((p) => p !== pv);
    dest.sheet.pivots.push(draft);
    writePivot(app, dest.sheet, draft, res, src);
  });
  lastPivot = { sheetId: dest.sheet.id, id: draft.id };
  app.exec('sheet.activate', app.book.sheets.indexOf(dest.sheet));
  app.grid.select(dest.range.r1, dest.range.c1);
}

async function deletePivot(app: App): Promise<void> {
  const t = currentPivot(app);
  if (!t) return;
  const { sheet, pv } = t;
  if (!guard(app, sheet)) return;
  app.book.transact('Delete PivotTable', () => {
    app.book.recordSheet(sheet);
    if (pv.lastRange) { const keys: [number, number][] = []; sheet.forEachIn(pv.lastRange, (r, c) => keys.push([r, c])); for (const [r, c] of keys) sheet.set(r, c, undefined); }
    sheet.pivots = sheet.pivots.filter((p) => p !== pv);
  });
  if (lastPivot?.id === pv.id) lastPivot = null;
  if (app.ui.paneOpen('pivotFields')) app.ui.closePane();
}

async function drillDown(app: App, arg?: { r?: number; c?: number }): Promise<void> {
  const s = app.sheet;
  const r = arg?.r ?? s.sel.r, c = arg?.c ?? s.sel.c;
  const pv = pivotAt(s, r, c);
  if (!pv) { toast('Select a value cell in a PivotTable.'); return; }
  const src = readSource(app.book, pv.source, s);
  if (typeof src === 'string') { toast(src); return; }
  const res = computePivot(pv, src, { date1904: app.book.calc.date1904 });
  const info = res.drill.get(`${r - pv.target.r},${c - pv.target.c}`);
  if (!info) { toast('Select a value cell (not a label) to show the details behind it.'); return; }
  const rows = drillRows(pv, src, info.row, info.col, { date1904: app.book.calc.date1904 });
  if (app.isReadOnly()) { toast('This workbook is read-only.'); return; }
  const book = app.book;
  let target!: Sheet;
  book.transact('Show Details', () => {
    book.recordBook();
    target = new Sheet(book, book.uniqueSheetName());
    book.sheets.splice(book.activeSheet, 0, target);
    const hdr = book.styleId({ bold: true, fill: HEADER_FILL, bb: { style: 'thin', color: LINE } });
    src.headers.forEach((hd, j) => target.set(0, j, { v: hd, s: hdr }));
    rows.forEach((ri, i) => src.rows[ri].forEach((v, j) => {
      if (v == null) return;
      const f = src.formats[j];
      target.set(i + 1, j, { v, ...(f ? { s: book.styleId({ numFmt: f }) } : {}) });
    }));
    src.headers.forEach((hd, j) => target.colW.set(j, Math.min(260, Math.max(target.defaultColW, hd.length * 7.5 + 20))));
  });
  app.exec('sheet.activate', book.sheets.indexOf(target));
  app.grid.select(0, 0);
  app.ui.status?.(`${rows.length} source row${rows.length === 1 ? '' : 's'}`);
}

// ============================================================================================
// Command definitions

const openMenu = (app: App, id: string) => { const c = getCommand(id); if (c?.menu) showMenu(c.menu(app), { anchor: app.ui.anchorFor(id), onPick: (e) => { if (e.id) app.exec(e.id, e.arg); } }); };

def([
  {
    id: 'pivot.create', label: 'PivotTable', icon: 'pivot_table_chart', tab: 'Insert', group: 'Tables', edits: true, popular: true,
    tip: 'Summarize data with a PivotTable', keywords: 'pivot table summarize crosstab',
    run: (app, arg) => createDialog(app, arg),
  },
  {
    id: 'pivot.recommended', label: 'Recommended PivotTables', icon: 'auto_awesome', tab: 'Insert', group: 'Tables', edits: true,
    tip: 'See suggested PivotTables for your data and pick one', keywords: 'pivot table suggest',
    run: (app, arg) => recommendedDialog(app, arg),
  },
  { id: 'pivot.fields', label: 'Field List', icon: 'view_list', tab: TAB, group: 'PivotTable', noRecord: true, edits: false, tip: 'Show or hide the PivotTable Fields pane', run: (app) => app.ui.togglePane('pivotFields'), isActive: (app) => !!app.ui?.paneOpen?.('pivotFields') },
  { id: 'pivot.options', label: 'Options', icon: 'settings', tab: TAB, group: 'PivotTable', edits: true, tip: 'Name, layout, totals and data source of the PivotTable', run: (app) => optionsDialog(app) },
  { id: 'pivot.refresh', label: 'Refresh', icon: 'refresh', tab: TAB, group: 'Data', kind: 'split', edits: true, keys: ['Alt-F5'], tip: 'Read the source data again and update the PivotTable', run: (app, arg) => refresh(app, arg === 'all'), menu: () => [{ id: 'pivot.refresh', label: 'Refresh', icon: 'refresh' }, { id: 'pivot.refresh', arg: 'all', label: 'Refresh All PivotTables', icon: 'sync' }] },
  { id: 'pivot.refreshAll', label: 'Refresh All PivotTables', icon: 'sync', tab: TAB, group: 'Data', edits: true, keys: ['Mod-Alt-F5'], run: (app) => refresh(app, true) },
  { id: 'pivot.changeSource', label: 'Change Data Source', icon: 'dataset', tab: TAB, group: 'Data', edits: true, tip: 'Use a different range or table for this PivotTable', run: (app, arg) => changeSource(app, arg) },
  { id: 'pivot.clear', label: 'Clear All', icon: 'clear_all', tab: TAB, group: 'Actions', edits: true, tip: 'Remove all fields, filters and formatting from the PivotTable', run: (app) => clearPivot(app) },
  { id: 'pivot.move', label: 'Move PivotTable', icon: 'open_with', tab: TAB, group: 'Actions', edits: true, tip: 'Move the PivotTable to another location', run: (app, arg) => movePivot(app, arg) },
  { id: 'pivot.delete', label: 'Delete PivotTable', icon: 'delete', tab: TAB, group: 'Actions', edits: true, run: (app) => deletePivot(app) },
  {
    id: 'pivot.layout', label: 'Report Layout', icon: 'table_view', tab: TAB, group: 'Layout', kind: 'menu', edits: true,
    run: (app, mode) => { if (mode === 'compact' || mode === 'outline' || mode === 'tabular') { const t = currentPivot(app); if (t) return applyPivot(app, t.sheet, t.pv, (d) => { d.layout = mode; }, 'Report Layout'); return; } openMenu(app, 'pivot.layout'); },
    menu: (app) => { const t = currentPivot(app, true); const m = t?.pv.layout; return [{ id: 'pivot.layout', arg: 'compact', label: 'Show in Compact Form', checked: m === 'compact' }, { id: 'pivot.layout', arg: 'outline', label: 'Show in Outline Form', checked: m === 'outline' }, { id: 'pivot.layout', arg: 'tabular', label: 'Show in Tabular Form', checked: m === 'tabular' }]; },
  },
  {
    id: 'pivot.subtotals', label: 'Subtotals', icon: 'functions', tab: TAB, group: 'Layout', kind: 'menu', edits: true,
    run: (app, on) => { if (typeof on === 'boolean') { const t = currentPivot(app); if (t) return applyPivot(app, t.sheet, t.pv, (d) => { d.subtotals = on; }, 'Subtotals'); return; } openMenu(app, 'pivot.subtotals'); },
    menu: (app) => { const t = currentPivot(app, true); return [{ id: 'pivot.subtotals', arg: false, label: 'Do Not Show Subtotals', checked: t ? !t.pv.subtotals : false }, { id: 'pivot.subtotals', arg: true, label: 'Show All Subtotals', checked: !!t?.pv.subtotals }]; },
  },
  {
    id: 'pivot.grandTotals', label: 'Grand Totals', icon: 'summarize', tab: TAB, group: 'Layout', kind: 'menu', edits: true,
    run: (app, mode) => {
      const map: Record<string, [boolean, boolean]> = { off: [false, false], on: [true, true], rows: [true, false], cols: [false, true] };
      if (typeof mode === 'string' && map[mode]) { const t = currentPivot(app); if (t) return applyPivot(app, t.sheet, t.pv, (d) => { [d.grandRows, d.grandCols] = map[mode]; }, 'Grand Totals'); return; }
      openMenu(app, 'pivot.grandTotals');
    },
    menu: (app) => {
      const t = currentPivot(app, true);
      const gr = !!t?.pv.grandRows, gc = !!t?.pv.grandCols;
      return [
        { id: 'pivot.grandTotals', arg: 'off', label: 'Off for Rows and Columns', checked: !!t && !gr && !gc },
        { id: 'pivot.grandTotals', arg: 'on', label: 'On for Rows and Columns', checked: gr && gc },
        { id: 'pivot.grandTotals', arg: 'rows', label: 'On for Rows Only', checked: gr && !gc },
        { id: 'pivot.grandTotals', arg: 'cols', label: 'On for Columns Only', checked: !gr && gc },
      ];
    },
  },
  { id: 'pivot.drillDown', label: 'Show Details', icon: 'zoom_in_map', tab: TAB, group: 'Data', edits: true, tip: 'Show the source rows behind the selected value on a new sheet', keywords: 'drill down details', run: (app, arg) => drillDown(app, arg) },
  {
    id: 'pivot.valueSettings', label: 'Value Field Settings', icon: 'tune', tab: TAB, group: 'PivotTable', edits: true,
    tip: 'Change how the value at the active cell is summarized and shown',
    run: async (app, idx?: number) => {
      const t = currentPivot(app);
      if (!t) return;
      if (!t.pv.values.length) { toast('The PivotTable has no value fields.'); return; }
      let i = typeof idx === 'number' ? idx : 0;
      if (typeof idx !== 'number') {
        const src = readSource(app.book, t.pv.source, t.sheet);
        if (typeof src !== 'string') {
          const res = computePivot(t.pv, src);
          const cellObj = res.cells[app.sheet.sel.r - t.pv.target.r]?.[app.sheet.sel.c - t.pv.target.c];
          if (cellObj && typeof cellObj === 'object' && !(cellObj instanceof CellError) && (cellObj as PivotCellObj).vi != null) {
            const vi = (cellObj as PivotCellObj).vi!;
            const live = t.pv.values.filter((v) => findField(src.headers, v.field) >= 0)[vi];
            i = Math.max(0, t.pv.values.indexOf(live));
          }
        }
      }
      await valueSettings(app, t.sheet, t.pv, i);
    },
  },
]);

void absRangeName;
