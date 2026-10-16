/**
 * Data tab: Sort & Filter (AutoFilter and table header menus, Custom Sort, Advanced Filter), Data Tools (Text to
 * Columns, Flash Fill, Remove Duplicates, Data Validation, Consolidate), Outline (Group, Subtotal, detail
 * levels) and Get & Transform (text, web, workbook, clipboard and table/range queries). What-If tools live in
 * whatif.ts and the Analysis ToolPak in analysis.ts.
 */
import { def, type MenuEntry } from './registry';
import type { App } from '../app';
import { Sheet, type Workbook } from '../model/workbook';
import type { Cell, CellStyle, DataValidation, FilterCriteria, Scalar, TableDef } from '../model/types';
import { CellError } from '../model/types';
import { cellKey, keyRow, keyCol, colName, cellName, MAX_ROWS, MAX_COLS, type Range } from '../model/address';
import { resolveRangeText, rangeRefText, shiftCells } from '../model/structure';
import { parseDateTime } from '../model/dates';
import { formatValue, isDateFormat } from '../model/numfmt';
import { activeRange, clipToUsed, currentRegion, hasHeaderRow, isWholeCols, isWholeRows, addrText, selRanges } from './ops';
import { openDialog, toast, showMenu, alertDialog, confirmDialog, choiceDialog, promptDialog, setListOptions, type Field, type DialogControl } from '../ui/dialog';
import { h, clear, icon } from '../ui/dom';
import { BUILTIN_LISTS } from '../prefs';
import { validates, listItems } from './edit';
import {
  sortRange, compareForSort, extendFilterRange, applyFilters, distinctColumnValues, tableDataRange, pickListValues, displayText, isBlank,
  advancedCriteria, splitLine, suggestBreaks, convertSplitField, inferFlashFill, runFlashFill, duplicateRows, removeRowsInRange,
  addSubtotals, removeSubtotals, SUBTOTAL_FUNCS, changeLevel, maxLevel, showOutlineLevel, setDetail, autoOutline, consolidate, AGG_FUNCS,
  aggFormulaName, applySteps, describeStep, subtractRange, DATE_PERIODS, rowsOf,
  type SortKey, type FilterArea, type SplitOptions, type ColumnFormat, type AggName, type TransformStep, type TableData, type OutlineAxis,
} from './dataops';

// --------------------------------------------------------------------------------------------
// Shared helpers

const PROTECTED_MSG = 'The sheet is protected. Unprotect it (Review ▸ Unprotect Sheet) to make this change.';

/** True (with a message) when the active sheet's protection blocks the action. */
export function blocked(app: App, action?: string): boolean {
  const p = app.sheet.protection;
  if (p.on && !(action && p.allow.includes(action))) { toast(PROTECTED_MSG); return true; }
  return false;
}

function status(app: App, msg: string): void { try { app.ui?.status?.(msg); } catch { /* no UI */ } }

function tableAt(sheet: Sheet, r: number, c: number): TableDef | undefined {
  return sheet.tables.find((t) => r >= t.range.r1 && r <= t.range.r2 && c >= t.range.c1 && c <= t.range.c2);
}

function inRange(rg: Range, r: number, c: number): boolean { return r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2; }

const hasValue = (sheet: Sheet, r: number, c: number) => { const x = sheet.get(r, c); return !!x && ((x.v != null && x.v !== '') || x.f != null); };

function regionIsEmpty(sheet: Sheet, rg: Range): boolean {
  let any = false;
  sheet.forEachIn(rg, (r, c) => { if (!any && hasValue(sheet, r, c)) any = true; });
  return !any;
}

/** Screen point just below a cell (for drop-down menus). */
function cellPoint(app: App, r: number, c: number): { x: number; y: number } {
  try {
    const g = app.grid;
    const rect = g.el.getBoundingClientRect();
    const rc = g.cellRect(r, c);
    return { x: rect.left + rc.x, y: rect.top + rc.y + rc.h };
  } catch { return { x: 100, y: 100 }; }
}

/** Adds a sheet after the active one (inside a transaction) and returns it. */
export function addSheetAfter(book: Workbook, base: string): Sheet {
  book.recordBook();
  let name = base.replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Sheet';
  for (let i = 2; book.sheetByName(name); i++) name = `${base.slice(0, 27)} (${i})`;
  const s = new Sheet(book, name);
  book.sheets.splice(book.activeSheet + 1, 0, s);
  book.activeSheet = book.activeSheet + 1;
  return s;
}

function activateSheet(app: App, sheet: Sheet): void {
  const i = app.book.sheets.indexOf(sheet);
  if (i >= 0) app.exec('sheet.activate', i);
}

function boldStyle(book: Workbook, base = 0): number { return book.derivedStyle(base, { bold: true }); }

/** Writes a matrix of values (inside a transaction). Optional bold first row. */
export function writeMatrix(sheet: Sheet, r0: number, c0: number, m: Scalar[][], opts: { boldFirstRow?: boolean; boldFirstCol?: boolean; numFmt?: string } = {}): void {
  const book = sheet.book;
  m.forEach((row, i) => row.forEach((v, j) => {
    const cur = sheet.get(r0 + i, c0 + j);
    let s = cur?.s ?? 0;
    if ((opts.boldFirstRow && i === 0) || (opts.boldFirstCol && j === 0)) s = book.derivedStyle(s, { bold: true });
    if (opts.numFmt && typeof v === 'number' && !(opts.boldFirstRow && i === 0)) s = book.derivedStyle(s, { numFmt: opts.numFmt });
    if (v == null || v === '') sheet.set(r0 + i, c0 + j, s ? { s } : undefined);
    else sheet.set(r0 + i, c0 + j, { v, ...(s ? { s } : {}) });
  }));
}

// --------------------------------------------------------------------------------------------
// Sort target and sorting

interface SortTarget { sheet: Sheet; rg: Range; header: boolean; table?: TableDef; filter?: boolean }

/** The range a sort applies to: a table body, the AutoFilter range, or the selection/current region (expanded like Excel). */
async function sortTarget(app: App, ask = true): Promise<SortTarget | null> {
  const sheet = app.sheet;
  const { r, c } = sheet.sel;
  if (selRanges(app).length > 1) { toast("This command can't be used on multiple selections."); return null; }
  const t = tableAt(sheet, r, c);
  if (t) return { sheet, rg: { ...t.range, r2: t.range.r2 - (t.opts.totals ? 1 : 0) }, header: t.opts.header, table: t };
  if (sheet.filter && inRange(sheet.filter.range, r, c)) return { sheet, rg: extendFilterRange(sheet, sheet.filter.range), header: true, filter: true };
  const sel = activeRange(app);
  let rg: Range;
  if (sel.r1 === sel.r2 && sel.c1 === sel.c2) {
    rg = currentRegion(sheet, r, c);
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && !hasValue(sheet, r, c)) { toast("This can't be applied to the selected range. Select a single cell in a range and try again."); return null; }
  } else {
    rg = clipToUsed(sheet, sel);
    const region = currentRegion(sheet, r, c);
    const wider = region.c1 < rg.c1 || region.c2 > rg.c2;
    if (ask && wider && rg.c1 === rg.c2 && !isWholeRows(sel)) {
      const ch = await choiceDialog('Sort Warning', 'There is data next to your selection that will not be sorted. Expand the selection to include it, or sort only the selected cells?', ['Expand the selection', 'Continue with the current selection']);
      if (ch < 0) return null;
      if (ch === 0) rg = { r1: Math.min(region.r1, rg.r1), c1: region.c1, r2: Math.max(region.r2, rg.r2), c2: region.c2 };
    }
  }
  if (regionIsEmpty(sheet, rg)) { toast('There is no data to sort here.'); return null; }
  return { sheet, rg, header: hasHeaderRow(sheet, rg) };
}

function iconAtFn(app: App, sheet: Sheet) {
  return (r: number, c: number) => { try { return app.tab.cf.evaluate(sheet, r, c, app.book.version)?.icon || null; } catch { return null; } };
}

function runSort(app: App, tg: SortTarget, keys: SortKey[], opts: { header: boolean; byCols?: boolean; caseSensitive?: boolean }, label = 'Sort'): void {
  const sheet = tg.sheet;
  const skipHidden = !opts.byCols && (!!tg.filter || !!tg.table?.filter || (!!sheet.filter && inRange(sheet.filter.range, tg.rg.r1, tg.rg.c1)));
  app.book.transact(label, () => {
    sortRange(sheet, tg.rg, keys, { header: opts.header, byCols: opts.byCols, caseSensitive: opts.caseSensitive, skipHidden, iconAt: iconAtFn(app, sheet) });
    const k0 = keys[0];
    if (!opts.byCols && k0 && (k0.on || 'values') === 'values') {
      if (tg.table) { tg.table.filter = { columns: tg.table.filter?.columns || {}, sort: { col: k0.index - tg.table.range.c1, desc: !!k0.desc } }; }
      else if (tg.filter && sheet.filter) sheet.filter.sort = { col: k0.index - sheet.filter.range.c1, desc: !!k0.desc };
    }
  });
}

async function quickSort(app: App, desc: boolean): Promise<void> {
  if (blocked(app, 'sort')) return;
  const tg = await sortTarget(app);
  if (!tg) return;
  const c = Math.max(tg.rg.c1, Math.min(tg.rg.c2, app.sheet.sel.c));
  try { runSort(app, tg, [{ index: c, desc }], { header: tg.header }); } catch (e) { void alertDialog('Sort', (e as Error).message); }
}

async function sortByColor(app: App, which: 'fill' | 'font' = 'fill'): Promise<void> {
  if (blocked(app, 'sort')) return;
  const { r, c } = app.sheet.sel;
  const st = app.sheet.styleOf(r, c);
  const tg = await sortTarget(app, false);
  if (!tg) return;
  try { runSort(app, tg, [{ index: c, on: which, color: (which === 'fill' ? st.fill : st.color) || '' }], { header: tg.header }, 'Sort by Color'); } catch (e) { void alertDialog('Sort', (e as Error).message); }
}

/** Custom lists offered by the Order box. */
function customLists(app: App): string[][] { return [...BUILTIN_LISTS, ...(app.prefs.customLists || [])].map((l) => l.items); }

interface LevelState { index: number; on: 'values' | 'fill' | 'font' | 'icon'; order: string; color: string; bottom: boolean }

async function customSortDialog(app: App): Promise<void> {
  if (blocked(app, 'sort')) return;
  const tg = await sortTarget(app);
  if (!tg) return;
  const sheet = tg.sheet;
  let header = tg.header, byCols = false, caseSensitive = false;
  const levels: LevelState[] = [{ index: Math.max(tg.rg.c1, Math.min(tg.rg.c2, sheet.sel.c)), on: 'values', order: 'asc', color: '', bottom: false }];
  const lists = customLists(app);
  const content = h('div', { class: 'ql-data-sort' });
  const keyLabel = (i: number) => {
    if (byCols) return `Row ${i + 1}${header ? '' : ''}`;
    const t = header ? displayText(sheet, tg.rg.r1, i) : '';
    return t ? t : `Column ${colName(i)}`;
  };
  const keyOptions = () => (byCols ? rowsOf(tg.rg.r1, tg.rg.r2) : rowsOf(tg.rg.c1, tg.rg.c2));
  const colorsFor = (idx: number, which: 'fill' | 'font') => {
    const out = new Set<string>();
    const lines = byCols ? rowsOf(tg.rg.c1 + (header ? 1 : 0), tg.rg.c2) : rowsOf(tg.rg.r1 + (header ? 1 : 0), tg.rg.r2);
    for (const l of lines) { const st = byCols ? sheet.styleOf(idx, l) : sheet.styleOf(l, idx); out.add(((which === 'fill' ? st.fill : st.color) || '').toLowerCase()); }
    return [...out];
  };
  const render = () => {
    clear(content);
    const top = h('div', { class: 'ql-data-sort-bar' },
      h('button', { class: 'ql-btn', type: 'button', onclick: () => { levels.push({ index: keyOptions()[0], on: 'values', order: 'asc', color: '', bottom: false }); render(); } }, icon('add'), ' Add Level'),
      h('button', { class: 'ql-btn', type: 'button', onclick: () => { if (levels.length > 1) levels.pop(); render(); } }, icon('remove'), ' Delete Level'),
      h('button', { class: 'ql-btn', type: 'button', onclick: () => { levels.push({ ...levels[levels.length - 1] }); render(); } }, icon('content_copy'), ' Copy Level'),
    );
    const hdr = h('input', { type: 'checkbox', id: 'ql-sort-hdr', checked: header, disabled: byCols }) as HTMLInputElement;
    hdr.addEventListener('change', () => { header = hdr.checked; render(); });
    top.appendChild(h('label', { class: 'ql-check ql-data-inline', for: 'ql-sort-hdr' }, hdr, ' My data has headers'));
    content.appendChild(top);
    levels.forEach((lv, i) => {
      const row = h('div', { class: 'ql-data-level' });
      row.appendChild(h('span', { class: 'ql-data-level-lab' }, i === 0 ? 'Sort by' : 'Then by'));
      const kSel = h('select', { 'aria-label': 'Column' }) as HTMLSelectElement;
      for (const k of keyOptions()) kSel.appendChild(h('option', { value: String(k) }, keyLabel(k)));
      kSel.value = String(lv.index);
      kSel.addEventListener('change', () => { lv.index = +kSel.value; render(); });
      const onSel = h('select', { 'aria-label': 'Sort on' }) as HTMLSelectElement;
      for (const [v, l] of [['values', 'Cell Values'], ['fill', 'Cell Color'], ['font', 'Font Color'], ['icon', 'Conditional Formatting Icon']]) onSel.appendChild(h('option', { value: v }, l));
      onSel.value = lv.on;
      onSel.addEventListener('change', () => { lv.on = onSel.value as LevelState['on']; lv.color = ''; render(); });
      row.append(kSel, onSel);
      if (lv.on === 'values') {
        const oSel = h('select', { 'aria-label': 'Order' }) as HTMLSelectElement;
        const sample = byCols ? sheet.value(lv.index, tg.rg.c1 + (header ? 1 : 0)) : sheet.value(tg.rg.r1 + (header ? 1 : 0), lv.index);
        const num = typeof sample === 'number';
        oSel.appendChild(h('option', { value: 'asc' }, num ? 'Smallest to Largest' : 'A to Z'));
        oSel.appendChild(h('option', { value: 'desc' }, num ? 'Largest to Smallest' : 'Z to A'));
        lists.forEach((l, li) => { oSel.appendChild(h('option', { value: 'list:' + li }, l.slice(0, 4).join(', ') + (l.length > 4 ? ', …' : ''))); oSel.appendChild(h('option', { value: 'rlist:' + li }, l.slice(-4).reverse().join(', ') + ', … (reverse)')); });
        oSel.value = lv.order;
        oSel.addEventListener('change', () => { lv.order = oSel.value; });
        row.appendChild(oSel);
      } else if (lv.on === 'icon') {
        row.appendChild(h('span', { class: 'ql-help' }, 'Icon of the active cell on top'));
        const pos = h('select', { 'aria-label': 'Position' }) as HTMLSelectElement;
        pos.append(h('option', { value: 'top' }, 'On Top'), h('option', { value: 'bottom' }, 'On Bottom'));
        pos.value = lv.bottom ? 'bottom' : 'top';
        pos.addEventListener('change', () => { lv.bottom = pos.value === 'bottom'; });
        row.appendChild(pos);
      } else {
        const cSel = h('select', { 'aria-label': 'Color' }) as HTMLSelectElement;
        for (const col of colorsFor(lv.index, lv.on)) { const o = h('option', { value: col, style: col ? `background:${col}` : '' }, col || 'No Color'); cSel.appendChild(o); }
        if (!lv.color && cSel.options.length) lv.color = cSel.options[0].value;
        cSel.value = lv.color;
        cSel.addEventListener('change', () => { lv.color = cSel.value; });
        const pos = h('select', { 'aria-label': 'Position' }) as HTMLSelectElement;
        pos.append(h('option', { value: 'top' }, 'On Top'), h('option', { value: 'bottom' }, 'On Bottom'));
        pos.value = lv.bottom ? 'bottom' : 'top';
        pos.addEventListener('change', () => { lv.bottom = pos.value === 'bottom'; });
        row.append(cSel, pos);
      }
      content.appendChild(row);
    });
    const cs = h('input', { type: 'checkbox', id: 'ql-sort-case', checked: caseSensitive }) as HTMLInputElement;
    cs.addEventListener('change', () => { caseSensitive = cs.checked; });
    const orient = h('select', { 'aria-label': 'Orientation' }) as HTMLSelectElement;
    orient.append(h('option', { value: 'rows' }, 'Sort top to bottom'), h('option', { value: 'cols' }, 'Sort left to right'));
    orient.value = byCols ? 'cols' : 'rows';
    orient.addEventListener('change', () => {
      byCols = orient.value === 'cols';
      if (byCols) header = false;
      const opts = keyOptions();
      for (const l of levels) if (!opts.includes(l.index)) l.index = opts[0];
      render();
    });
    content.appendChild(h('div', { class: 'ql-data-sort-opts' }, h('label', { class: 'ql-check ql-data-inline', for: 'ql-sort-case' }, cs, ' Case sensitive'), orient));
  };
  render();
  const res = await openDialog({ title: 'Sort', content, wide: true });
  if (!res) return;
  const keys: SortKey[] = levels.map((lv) => {
    if (lv.on === 'values') {
      const m = /^(r?)list:(\d+)$/.exec(lv.order);
      if (m) return { index: lv.index, list: lists[+m[2]], desc: m[1] === 'r' };
      return { index: lv.index, desc: lv.order === 'desc' };
    }
    if (lv.on === 'icon') {
      const ic = iconAtFn(app, sheet)(sheet.sel.r, sheet.sel.c);
      return { index: lv.index, on: 'icon', icon: ic || { set: '', index: -1 }, bottom: lv.bottom };
    }
    return { index: lv.index, on: lv.on, color: lv.color, bottom: lv.bottom };
  });
  try { runSort(app, tg, keys, { header: byCols ? false : header, byCols, caseSensitive }); } catch (e) { void alertDialog('Sort', (e as Error).message); }
}

// --------------------------------------------------------------------------------------------
// Filters

interface FilterSource { kind: 'sheet' | 'table'; table?: TableDef; area: FilterArea; headerRow: number }

function filterSourceAt(sheet: Sheet, r: number, c: number): FilterSource | null {
  const f = sheet.filter;
  if (f && inRange(f.range, r, c)) {
    const rg = f.range;
    return { kind: 'sheet', area: { r1: rg.r1 + 1, r2: rg.r2, c1: rg.c1, c2: rg.c2, columns: f.columns || {} }, headerRow: rg.r1 };
  }
  const t = tableAt(sheet, r, c);
  if (t && t.opts.header) { const d = tableDataRange(t); return { kind: 'table', table: t, area: { ...d, columns: t.filter?.columns || {} }, headerRow: t.range.r1 }; }
  return null;
}

/** Sets (or clears, with null) one column's criteria and re-filters. */
function setCriteria(app: App, src: FilterSource, col: number, crit: FilterCriteria | null): void {
  const sheet = app.sheet;
  app.book.transact('Filter', () => {
    sheet.book.recordSheet(sheet);
    if (src.kind === 'sheet' && sheet.filter) {
      sheet.filter.range = extendFilterRange(sheet, sheet.filter.range);
      const cols = { ...(sheet.filter.columns || {}) };
      if (crit) cols[col] = crit; else delete cols[col];
      sheet.filter.columns = cols;
    } else if (src.table) {
      const t = src.table;
      const cols = { ...(t.filter?.columns || {}) };
      if (crit) cols[col] = crit; else delete cols[col];
      t.filter = { ...(t.filter || {}), columns: cols };
    }
    const res = applyFilters(sheet);
    status(app, `${res.shown} of ${res.total} records found`);
  });
}

function ensureFilter(app: App): FilterSource | null {
  const sheet = app.sheet;
  const { r, c } = sheet.sel;
  let src = filterSourceAt(sheet, r, c);
  if (src) return src;
  const t = tableAt(sheet, r, c);
  if (t) {
    app.book.transact('Filter', () => { sheet.book.recordSheet(sheet); t.opts = { ...t.opts, filterButton: true }; });
    return filterSourceAt(sheet, r, c);
  }
  const rg = currentRegion(sheet, r, c);
  if (regionIsEmpty(sheet, rg)) { toast("This can't be applied to the selected range. Select a single cell in a range and try again."); return null; }
  app.book.transact('Filter', () => { sheet.book.recordSheet(sheet); sheet.filter = { range: rg, columns: {} }; });
  src = filterSourceAt(sheet, r, c);
  return src;
}

const OPS_TEXT: [string, string][] = [['eq', 'equals'], ['ne', 'does not equal'], ['gt', 'is greater than'], ['ge', 'is greater than or equal to'], ['lt', 'is less than'], ['le', 'is less than or equal to'], ['begins', 'begins with'], ['notBegins', 'does not begin with'], ['ends', 'ends with'], ['notEnds', 'does not end with'], ['contains', 'contains'], ['notContains', 'does not contain']];
const OPS_DATE: [string, string][] = [['eq', 'equals'], ['ne', 'does not equal'], ['gt', 'is after'], ['ge', 'is after or equal to'], ['lt', 'is before'], ['le', 'is before or equal to']];

async function customFilterDialog(app: App, src: FilterSource, col: number, kind: 'text' | 'number' | 'date', op1 = 'eq', op2 = '', and = true, values: string[] = []): Promise<void> {
  const ops = kind === 'date' ? OPS_DATE : OPS_TEXT;
  const existing = src.area.columns[col]?.custom || [];
  const name = displayText(app.sheet, src.headerRow, src.area.c1 + col) || `Column ${colName(src.area.c1 + col)}`;
  const opOpts = [{ value: '', label: '' }, ...ops.map(([v, l]) => ({ value: v, label: l }))];
  const res = await openDialog({
    title: 'Custom AutoFilter',
    help: `Show rows where: ${name}. Use ? to represent any single character and * to represent any series of characters.`,
    fields: [
      { key: 'op1', type: 'select', label: 'Condition', options: opOpts, value: existing[0]?.op ?? op1, half: true },
      { key: 'v1', type: 'combo', label: 'Value', options: values, value: existing[0]?.value ?? '', half: true, autofocus: true },
      { key: 'and', type: 'radio', options: [{ value: 'and', label: 'And' }, { value: 'or', label: 'Or' }], value: (existing.length ? src.area.columns[col]?.and : and) ? 'and' : 'or' },
      { key: 'op2', type: 'select', label: 'Condition', options: opOpts, value: existing[1]?.op ?? op2, half: true },
      { key: 'v2', type: 'combo', label: 'Value', options: values, value: existing[1]?.value ?? '', half: true },
    ],
  });
  if (!res) return;
  const custom: { op: string; value: string }[] = [];
  if (res.op1) custom.push({ op: res.op1, value: String(res.v1 ?? '') });
  if (res.op2) custom.push({ op: res.op2, value: String(res.v2 ?? '') });
  setCriteria(app, src, col, custom.length ? { custom, and: res.and === 'and' } : null);
}

async function top10Dialog(app: App, src: FilterSource, col: number): Promise<void> {
  const cur = src.area.columns[col]?.top10;
  const res = await openDialog({
    title: 'Top 10 AutoFilter',
    fields: [
      { key: 'top', type: 'select', label: 'Show', options: [{ value: 'top', label: 'Top' }, { value: 'bottom', label: 'Bottom' }], value: cur && !cur.top ? 'bottom' : 'top', half: true },
      { key: 'n', type: 'number', label: 'Number', value: cur?.n ?? 10, min: 1, max: 500, half: true },
      { key: 'unit', type: 'select', label: 'Of', options: [{ value: 'items', label: 'Items' }, { value: 'percent', label: 'Percent' }], value: cur?.percent ? 'percent' : 'items' },
    ],
    validate: (v) => (!(v.n >= 1) ? 'Enter a number from 1 to 500.' : null),
  });
  if (!res) return;
  setCriteria(app, src, col, { top10: { top: res.top === 'top', percent: res.unit === 'percent', n: Math.round(res.n) } });
}

/** The AutoFilter / table header drop-down. */
function filterMenu(app: App, arg: { r: number; c: number }): void {
  const sheet = app.sheet;
  const src = filterSourceAt(sheet, arg.r, arg.c);
  if (!src || arg.r !== src.headerRow) return;
  const col = arg.c - src.area.c1;
  const info = distinctColumnValues(sheet, src.area, col);
  const crit = src.area.columns[col];
  const name = displayText(sheet, src.headerRow, arg.c) || `Column ${colName(arg.c)}`;
  const kind = info.kind;
  const sortLabels = kind === 'number' ? ['Sort Smallest to Largest', 'Sort Largest to Smallest'] : kind === 'date' ? ['Sort Oldest to Newest', 'Sort Newest to Oldest'] : ['Sort A to Z', 'Sort Z to A'];
  const colorEntries = (which: 'fill' | 'font', act: string): MenuEntry[] => (which === 'fill' ? info.fills : info.fonts).map((c) => ({ id: '__f', label: c ? (which === 'fill' ? 'Cell color' : 'Font color') : which === 'fill' ? 'No Fill' : 'Automatic', swatch: c || 'transparent', arg: { act, which, color: c } }));
  const hasColors = info.fills.length > 1 || info.fonts.length > 1;
  const colorSub = (act: string): MenuEntry[] => [
    ...(info.fills.length > 1 ? [{ heading: act === 'sortColor' ? 'Sort by Cell Color' : 'Filter by Cell Color' }, ...colorEntries('fill', act)] : []),
    ...(info.fonts.length > 1 ? [{ heading: act === 'sortColor' ? 'Sort by Font Color' : 'Filter by Font Color' }, ...colorEntries('font', act)] : []),
  ];
  const typed: MenuEntry[] = kind === 'date'
    ? [
      ...OPS_DATE.slice(0, 1).concat([['lt', 'Before…'], ['gt', 'After…']] as any).map(([op, l]) => ({ id: '__f', label: op === 'eq' ? 'Equals…' : l, arg: { act: 'custom', op } })),
      { id: '__f', label: 'Between…', arg: { act: 'custom', op: 'ge', op2: 'le' } },
      { separator: true },
      ...DATE_PERIODS.map((p) => ({ id: '__f', label: p.label, arg: { act: 'period', p: p.id } })),
      { label: 'All Dates in the Period', submenu: [
        ...['Q1', 'Q2', 'Q3', 'Q4'].map((q, i) => ({ id: '__f', label: `Quarter ${i + 1}`, arg: { act: 'period', p: q } })),
        { separator: true },
        ...['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'].map((m, i) => ({ id: '__f', label: m, arg: { act: 'period', p: `M${i + 1}` } })),
      ] },
      { separator: true },
      { id: '__f', label: 'Custom Filter…', arg: { act: 'custom', op: 'eq' } },
    ]
    : kind === 'number'
      ? [
        ...[['eq', 'Equals…'], ['ne', 'Does Not Equal…'], ['gt', 'Greater Than…'], ['ge', 'Greater Than Or Equal To…'], ['lt', 'Less Than…'], ['le', 'Less Than Or Equal To…']].map(([op, l]) => ({ id: '__f', label: l, arg: { act: 'custom', op } })),
        { id: '__f', label: 'Between…', arg: { act: 'custom', op: 'ge', op2: 'le' } },
        { id: '__f', label: 'Top 10…', arg: { act: 'top10' } },
        { id: '__f', label: 'Above Average', arg: { act: 'avg', which: 'aboveAverage' } },
        { id: '__f', label: 'Below Average', arg: { act: 'avg', which: 'belowAverage' } },
        { separator: true },
        { id: '__f', label: 'Custom Filter…', arg: { act: 'custom', op: 'eq' } },
      ]
      : [
        ...[['eq', 'Equals…'], ['ne', 'Does Not Equal…'], ['begins', 'Begins With…'], ['ends', 'Ends With…'], ['contains', 'Contains…'], ['notContains', 'Does Not Contain…']].map(([op, l]) => ({ id: '__f', label: l, arg: { act: 'custom', op } })),
        { separator: true },
        { id: '__f', label: 'Custom Filter…', arg: { act: 'custom', op: 'eq' } },
      ];
  const entries: MenuEntry[] = [
    { id: '__f', label: sortLabels[0], icon: 'arrow_upward', arg: { act: 'sort', desc: false } },
    { id: '__f', label: sortLabels[1], icon: 'arrow_downward', arg: { act: 'sort', desc: true } },
    ...(hasColors ? [{ label: 'Sort by Color', icon: 'palette', submenu: colorSub('sortColor') }] : []),
    { separator: true },
    { id: '__f', label: `Clear Filter From "${name}"`, icon: 'filter_alt_off', arg: { act: 'clear' } },
    ...(hasColors ? [{ label: 'Filter by Color', icon: 'format_color_fill', submenu: colorSub('filterColor') }] : []),
    { label: kind === 'number' ? 'Number Filters' : kind === 'date' ? 'Date Filters' : 'Text Filters', icon: 'filter_list', submenu: typed },
  ];
  // Checkbox list with search.
  const selected = new Set<string>(crit?.values ? crit.values.map((v) => v.toLowerCase()) : info.values.map((v) => v.toLowerCase()));
  let blanksOn = crit?.values ? !!crit.blanks : true;
  const listEl = h('div', { class: 'ql-data-checklist', role: 'group', 'aria-label': 'Values' });
  const search = h('input', { type: 'search', class: 'ql-data-search', placeholder: 'Search', 'aria-label': 'Search values' }) as HTMLInputElement;
  const shown = () => { const q = search.value.trim().toLowerCase(); return q ? info.values.filter((v) => v.toLowerCase().includes(q)) : info.values; };
  const MAX = 10000;
  const renderList = () => {
    clear(listEl);
    const vis = shown();
    const allOn = vis.every((v) => selected.has(v.toLowerCase())) && (!!search.value || !info.blanks || blanksOn);
    const box = (label: string, checked: boolean, onChange: (on: boolean) => void, cls = '') => {
      const cb = h('input', { type: 'checkbox', checked }) as HTMLInputElement;
      cb.addEventListener('change', () => onChange(cb.checked));
      return h('label', { class: 'ql-data-checkrow' + cls }, cb, h('span', null, label));
    };
    listEl.appendChild(box(search.value ? '(Select All Search Results)' : '(Select All)', allOn, (on) => { for (const v of vis) { if (on) selected.add(v.toLowerCase()); else selected.delete(v.toLowerCase()); } if (!search.value) blanksOn = on; renderList(); }, ' ql-data-all'));
    for (const v of vis.slice(0, MAX)) listEl.appendChild(box(v, selected.has(v.toLowerCase()), (on) => { if (on) selected.add(v.toLowerCase()); else selected.delete(v.toLowerCase()); renderList(); }));
    if (vis.length > MAX) listEl.appendChild(h('div', { class: 'ql-help' }, `Showing the first ${MAX} of ${vis.length} items.`));
    if (info.blanks && !search.value) listEl.appendChild(box('(Blanks)', blanksOn, (on) => { blanksOn = on; renderList(); }));
  };
  search.addEventListener('input', renderList);
  renderList();
  const ok = h('button', { class: 'ql-btn ql-primary', type: 'button' }, 'OK');
  const cancel = h('button', { class: 'ql-btn', type: 'button' }, 'Cancel');
  const extra = h('div', { class: 'ql-data-filterbox' }, search, listEl, h('div', { class: 'ql-data-filterfoot' }, cancel, ok));
  const closeAll = () => document.querySelectorAll('.ql-menu').forEach((m) => m.remove());
  ok.addEventListener('click', () => {
    closeAll();
    const q = search.value.trim();
    const vis = shown();
    let next: FilterCriteria | null;
    if (q) next = { values: vis.filter((v) => selected.has(v.toLowerCase())), blanks: false };
    else if (info.values.every((v) => selected.has(v.toLowerCase())) && (!info.blanks || blanksOn)) next = null;
    else next = { values: info.values.filter((v) => selected.has(v.toLowerCase())), blanks: blanksOn };
    if (next && !next.values!.length && !next.blanks) { toast('Select at least one value.'); return; }
    setCriteria(app, src, col, next);
  });
  cancel.addEventListener('click', closeAll);
  search.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); ok.click(); } });
  const pt = cellPoint(app, arg.r, arg.c);
  showMenu(entries, {
    x: pt.x, y: pt.y, extra, className: 'ql-data-filtermenu', focus: false,
    onPick: (e) => {
      const a = e.arg || {};
      const absCol = src.area.c1 + col;
      switch (a.act) {
        case 'sort': {
          if (blocked(app, 'sort')) return;
          const rg: Range = { r1: src.headerRow, c1: src.area.c1, r2: src.area.r2, c2: src.area.c2 };
          try { runSort(app, { sheet, rg, header: true, table: src.table, filter: src.kind === 'sheet' }, [{ index: absCol, desc: a.desc }], { header: true }); } catch (err) { void alertDialog('Sort', (err as Error).message); }
          break;
        }
        case 'sortColor': {
          const rg: Range = { r1: src.headerRow, c1: src.area.c1, r2: src.area.r2, c2: src.area.c2 };
          try { runSort(app, { sheet, rg, header: true, table: src.table, filter: src.kind === 'sheet' }, [{ index: absCol, on: a.which, color: a.color }], { header: true }, 'Sort by Color'); } catch (err) { void alertDialog('Sort', (err as Error).message); }
          break;
        }
        case 'clear': setCriteria(app, src, col, null); break;
        case 'filterColor': setCriteria(app, src, col, { color: a.which === 'fill' ? { fill: a.color } : { font: a.color } }); break;
        case 'custom': void customFilterDialog(app, src, col, kind, a.op, a.op2 || '', true, info.values); break;
        case 'top10': void top10Dialog(app, src, col); break;
        case 'avg': setCriteria(app, src, col, { dynamic: a.which }); break;
        case 'period': setCriteria(app, src, col, { custom: [{ op: 'period', value: a.p }] }); break;
      }
    },
  });
  setTimeout(() => search.focus(), 30);
}

function toggleFilter(app: App): void {
  const sheet = app.sheet;
  const { r, c } = sheet.sel;
  const t = tableAt(sheet, r, c);
  if (t) {
    app.book.transact('Filter', () => {
      sheet.book.recordSheet(sheet);
      const on = !t.opts.filterButton;
      t.opts = { ...t.opts, filterButton: on };
      if (!on && t.filter) { delete t.filter; for (let rr = t.range.r1; rr <= t.range.r2; rr++) sheet.hiddenRows.delete(rr); applyFilters(sheet); }
    });
    return;
  }
  if (sheet.filter) {
    const rg = sheet.filter.range;
    app.book.transact('Remove Filter', () => {
      sheet.book.recordSheet(sheet);
      for (let rr = rg.r1; rr <= rg.r2; rr++) sheet.hiddenRows.delete(rr);
      sheet.filter = null;
      applyFilters(sheet);
    });
    return;
  }
  const sel = activeRange(app);
  let rg = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? currentRegion(sheet, r, c) : clipToUsed(sheet, sel);
  if (rg.r2 === rg.r1 && !(sel.r1 === sel.r2 && sel.c1 === sel.c2)) rg = { ...rg, r2: currentRegion(sheet, r, c).r2 };
  if (regionIsEmpty(sheet, rg)) { toast("This can't be applied to the selected range. Select a single cell in a range and try again."); return; }
  app.book.transact('Filter', () => { sheet.book.recordSheet(sheet); sheet.filter = { range: rg, columns: {} }; });
}

/** Advanced Filter state per sheet (for Reapply and Clear). */
const advState = new WeakMap<Sheet, { list: string; crit: string; unique: boolean }>();

function clearFilters(app: App): void {
  const sheet = app.sheet;
  const t = tableAt(sheet, sheet.sel.r, sheet.sel.c);
  const adv = advState.get(sheet);
  if (!sheet.filter && !t?.filter && !adv) { toast('There is no filter to clear here.'); return; }
  app.book.transact('Clear Filter', () => {
    sheet.book.recordSheet(sheet);
    if (adv) {
      const res = resolveRangeText(app.book, adv.list, sheet);
      if (res) for (let r = res.range.r1; r <= res.range.r2; r++) sheet.hiddenRows.delete(r);
      advState.delete(sheet);
    }
    if (t?.filter) t.filter = { ...t.filter, columns: {} };
    else if (sheet.filter) sheet.filter.columns = {};
    applyFilters(sheet);
  });
}

function reapply(app: App): void {
  const sheet = app.sheet;
  const adv = advState.get(sheet);
  if (adv) { void runAdvanced(app, { action: 'inPlace', list: adv.list, crit: adv.crit, copyTo: '', unique: adv.unique }); return; }
  if (!sheet.filter && !sheet.tables.some((t) => t.filter)) { toast('There is no filter to reapply.'); return; }
  app.book.transact('Reapply', () => {
    sheet.book.recordSheet(sheet);
    if (sheet.filter) sheet.filter.range = extendFilterRange(sheet, sheet.filter.range);
    const res = applyFilters(sheet);
    status(app, `${res.shown} of ${res.total} records found`);
  });
}

async function runAdvanced(app: App, v: { action: string; list: string; crit: string; copyTo: string; unique: boolean }): Promise<void> {
  const book = app.book, sheet = app.sheet;
  const L = resolveRangeText(book, v.list, sheet);
  if (!L) { toast('The list range is not valid.'); return; }
  const C = v.crit.trim() ? resolveRangeText(book, v.crit, sheet) : null;
  if (v.crit.trim() && !C) { toast('The criteria range is not valid.'); return; }
  const list = clipToUsed(L.sheet, L.range);
  const lsheet = L.sheet;
  const engine = app.engine;
  const test = C ? advancedCriteria(lsheet, list, C.range, C.sheet, (f, r, c, br, bc) => engine.evaluateScalar(f, lsheet, r, c, { r: br, c: bc })) : () => true;
  const rows: number[] = [];
  const seen = new Set<string>();
  let copyCols = rowsOf(list.c1, list.c2);
  let dest: { sheet: Sheet; range: Range } | null = null;
  if (v.action === 'copy') {
    dest = resolveRangeText(book, v.copyTo, sheet);
    if (!dest) { toast('The Copy To reference is not valid.'); return; }
    const names: string[] = [];
    for (let c = dest.range.c1; c <= Math.min(dest.range.c2, dest.range.c1 + 200); c++) { const t = displayText(dest.sheet, dest.range.r1, c); if (!t) break; names.push(t.toLowerCase()); }
    if (names.length && dest.range.c2 > dest.range.c1) {
      const map = names.map((n) => copyCols.find((c) => displayText(lsheet, list.r1, c).toLowerCase() === n));
      if (map.some((x) => x == null)) { toast('The Copy To range has a field name that is not in the list.'); return; }
      copyCols = map as number[];
    }
  }
  for (let r = list.r1 + 1; r <= list.r2; r++) {
    if (!test(r)) continue;
    if (v.unique) { const k = copyCols.map((c) => displayText(lsheet, r, c).toLowerCase()).join('\u0001'); if (seen.has(k)) continue; seen.add(k); }
    rows.push(r);
  }
  if (v.action !== 'copy') {
    if (lsheet !== sheet) { toast('You can only filter in place on the active sheet.'); return; }
    const keep = new Set(rows);
    book.transact('Advanced Filter', () => {
      book.recordSheet(sheet);
      for (let r = list.r1 + 1; r <= list.r2; r++) { if (keep.has(r)) sheet.hiddenRows.delete(r); else sheet.hiddenRows.add(r); }
    });
    advState.set(sheet, { list: v.list, crit: v.crit, unique: v.unique });
    status(app, `${rows.length} of ${list.r2 - list.r1} records found`);
    return;
  }
  const d = dest!;
  book.transact('Advanced Filter', () => {
    const out = [list.r1, ...rows];
    out.forEach((r, i) => copyCols.forEach((c, j) => {
      const cell = lsheet.get(r, c);
      const next: Cell | undefined = cell ? { ...(cell.v != null ? { v: cell.v } : {}), ...(cell.s ? { s: cell.s } : {}) } : undefined;
      d.sheet.set(d.range.r1 + i, d.range.c1 + j, next && (next.v != null || next.s) ? next : undefined);
    }));
  });
  status(app, `${rows.length} records copied`);
}

async function advancedFilterDialog(app: App): Promise<void> {
  if (blocked(app, 'autoFilter')) return;
  const sheet = app.sheet;
  const rg = sheet.filter?.range || currentRegion(sheet, sheet.sel.r, sheet.sel.c);
  const prev = advState.get(sheet);
  const res = await openDialog({
    title: 'Advanced Filter',
    fields: [
      { key: 'action', type: 'radio', label: 'Action', options: [{ value: 'inPlace', label: 'Filter the list, in place' }, { value: 'copy', label: 'Copy to another location' }], value: 'inPlace' },
      { key: 'list', type: 'text', label: 'List range', value: prev?.list || rangeRefText(sheet, rg) },
      { key: 'crit', type: 'text', label: 'Criteria range', value: prev?.crit || '', placeholder: 'e.g. $H$1:$I$3 (field names in the first row)' },
      { key: 'copyTo', type: 'text', label: 'Copy to', value: '', visibleIf: (v) => v.action === 'copy' },
      { key: 'unique', type: 'checkbox', label: 'Unique records only', value: !!prev?.unique },
      { key: 'i', type: 'info', label: 'Criteria rows are combined with OR; conditions in one row with AND. Text like "ab" matches values beginning with ab; use =ab for an exact match, or >, <, <> comparisons. A formula under a heading that is not a field name is a computed criterion written for the first data row.' },
    ],
    validate: (v) => (v.action === 'copy' && !String(v.copyTo).trim() ? 'Enter a Copy To location.' : null),
  });
  if (!res) return;
  await runAdvanced(app, res as any);
}

// --------------------------------------------------------------------------------------------
// Text to Columns

const DATE_ORDERS: ColumnFormat[] = ['MDY', 'DMY', 'YMD', 'MYD', 'DYM', 'YDM'];

async function textToColumns(app: App): Promise<void> {
  if (blocked(app)) return;
  const sheet = app.sheet;
  const sel = clipToUsed(sheet, activeRange(app));
  if (selRanges(app).length > 1 || sel.c1 !== sel.c2) { void alertDialog('Text to Columns', 'Text to Columns can convert only one column at a time. The range can be many rows tall but no more than one column wide.'); return; }
  const rows = rowsOf(sel.r1, sel.r2);
  const lines = rows.map((r) => { const v = sheet.value(r, sel.c1); return typeof v === 'string' ? v : displayText(sheet, r, sel.c1); });
  if (!lines.some((l) => l)) { toast('There is no text to convert in the selection.'); return; }
  const sample = lines.slice(0, 200).join('\n');
  const guess = sample.includes('\t') ? 'tab' : sample.includes(';') && !sample.includes(',') ? 'semicolon' : sample.includes(',') ? 'comma' : 'space';
  const fixedGuess = !/[\t,;]/.test(sample) && /\S {2,}\S/.test(sample);
  let formats: ColumnFormat[] = [];
  const preview = h('div', { class: 'ql-data-preview' });
  const optsOf = (v: Record<string, any>): SplitOptions => {
    if (v.mode === 'fixed') return { mode: 'fixed', breaks: String(v.breaks || '').split(/[\s,;]+/).map(Number).filter((n) => n > 0) };
    const delims: string[] = [];
    if (v.tab) delims.push('\t');
    if (v.semicolon) delims.push(';');
    if (v.comma) delims.push(',');
    if (v.space) delims.push(' ');
    if (v.other && v.otherChar) delims.push(String(v.otherChar)[0]);
    return { mode: 'delimited', delimiters: delims, consecutive: !!v.consecutive, qualifier: v.qualifier === 'none' ? '' : v.qualifier };
  };
  const renderPreview = (v: Record<string, any>) => {
    const o = optsOf(v);
    const split = lines.slice(0, 12).map((l) => splitLine(l, o));
    const n = Math.max(1, ...split.map((s) => s.length));
    while (formats.length < n) formats.push('general');
    formats = formats.slice(0, Math.max(n, 1));
    clear(preview);
    const table = h('table', { class: 'ql-data-table' });
    const head = h('tr');
    for (let j = 0; j < n; j++) {
      const s = h('select', { 'aria-label': `Column ${j + 1} format` }) as HTMLSelectElement;
      for (const [val, lab] of [['general', 'General'], ['text', 'Text'], ...DATE_ORDERS.map((d) => [d, `Date (${d})`]), ['skip', 'Do not import (skip)']]) s.appendChild(h('option', { value: val }, lab));
      s.value = formats[j];
      s.addEventListener('change', () => { formats[j] = s.value as ColumnFormat; renderPreview(v); });
      head.appendChild(h('th', null, s));
    }
    table.appendChild(head);
    for (const row of split) {
      const tr = h('tr');
      for (let j = 0; j < n; j++) tr.appendChild(h('td', { class: formats[j] === 'skip' ? 'ql-data-skip' : '' }, row[j] ?? ''));
      table.appendChild(tr);
    }
    preview.appendChild(table);
  };
  const fields: Field[] = [
    { key: 'mode', type: 'radio', label: 'Original data type', options: [{ value: 'delimited', label: 'Delimited — characters such as commas or tabs separate fields' }, { value: 'fixed', label: 'Fixed width — fields are aligned in columns with spaces' }], value: fixedGuess ? 'fixed' : 'delimited' },
    { key: 'tab', type: 'checkbox', label: 'Tab', value: guess === 'tab', half: true, visibleIf: (v) => v.mode === 'delimited' },
    { key: 'semicolon', type: 'checkbox', label: 'Semicolon', value: guess === 'semicolon', half: true, visibleIf: (v) => v.mode === 'delimited' },
    { key: 'comma', type: 'checkbox', label: 'Comma', value: guess === 'comma', half: true, visibleIf: (v) => v.mode === 'delimited' },
    { key: 'space', type: 'checkbox', label: 'Space', value: guess === 'space', half: true, visibleIf: (v) => v.mode === 'delimited' },
    { key: 'other', type: 'checkbox', label: 'Other', value: false, half: true, visibleIf: (v) => v.mode === 'delimited' },
    { key: 'otherChar', type: 'text', label: 'Other delimiter', value: '', half: true, width: '4em', visibleIf: (v) => v.mode === 'delimited' },
    { key: 'consecutive', type: 'checkbox', label: 'Treat consecutive delimiters as one', value: false, visibleIf: (v) => v.mode === 'delimited' },
    { key: 'qualifier', type: 'select', label: 'Text qualifier', options: [{ value: '"', label: '"' }, { value: "'", label: "'" }, { value: 'none', label: '{none}' }], value: '"', visibleIf: (v) => v.mode === 'delimited' },
    { key: 'breaks', type: 'text', label: 'Break positions (characters from the left)', value: suggestBreaks(lines.slice(0, 200)).join(', '), placeholder: 'e.g. 10, 25', visibleIf: (v) => v.mode === 'fixed' },
    { key: 'detect', type: 'button', label: 'Detect Fixed-Width Breaks', onClick: (_v, ctl) => { ctl.set('mode', 'fixed'); ctl.set('breaks', suggestBreaks(lines.slice(0, 200)).join(', ')); renderPreview(ctl.values()); } },
    { key: 'dest', type: 'text', label: 'Destination', value: cellName(sel.r1, sel.c1, true, true) },
  ];
  const res = await openDialog({ title: 'Convert Text to Columns', fields, wide: true, content: h('div', null, h('div', { class: 'ql-label' }, 'Data preview (choose a format above each column)'), preview), onChange: (v) => renderPreview(v) });
  if (!res) return;
  const o = optsOf(res);
  if (o.mode === 'delimited' && !o.delimiters!.length) { toast('Choose at least one delimiter.'); return; }
  const dest = resolveRangeText(app.book, String(res.dest || ''), sheet);
  if (!dest) { toast('The destination reference is not valid.'); return; }
  const out = lines.map((l) => splitLine(l, o));
  const n = Math.max(...out.map((x) => x.length));
  const keepCols = rowsOf(0, n - 1).filter((j) => (formats[j] || 'general') !== 'skip');
  let occupied = false;
  for (let i = 0; i < out.length; i++) for (let k = 0; k < keepCols.length; k++) {
    const r = dest.range.r1 + i, c = dest.range.c1 + k;
    if (dest.sheet === sheet && c === sel.c1 && r >= sel.r1 && r <= sel.r2) continue;
    if (hasValue(dest.sheet, r, c)) occupied = true;
  }
  if (occupied && !(await confirmDialog('Text to Columns', 'There\'s already data here. Do you want to replace it?', 'OK', 'Cancel'))) return;
  const date1904 = app.book.calc.date1904;
  app.book.transact('Text to Columns', () => {
    out.forEach((parts, i) => keepCols.forEach((j, k) => {
      const r = dest.range.r1 + i, c = dest.range.c1 + k;
      const conv = convertSplitField(parts[j] ?? '', formats[j] || 'general', date1904);
      const cur = dest.sheet.get(r, c);
      let s = cur?.s;
      if (conv.fmt && conv.fmt !== 'General') s = app.book.derivedStyle(s ?? 0, { numFmt: conv.fmt });
      dest.sheet.set(r, c, conv.v == null ? (s ? { s } : undefined) : { v: conv.v, ...(s ? { s } : {}) });
    }));
  });
}

// --------------------------------------------------------------------------------------------
// Flash Fill

interface FlashPlan { col: number; rows: number[]; values: string[]; examples: number }

/** Works out what Flash Fill would write for the column of (r, c); null when there is no pattern. */
export function planFlashFill(sheet: Sheet, r: number, c: number): FlashPlan | null {
  let region = currentRegion(sheet, r, c);
  if (region.c1 === region.c2) return null;
  const srcCols = rowsOf(region.c1, region.c2).filter((x) => x !== c);
  const textOf = (rr: number, cc: number) => { const v = sheet.value(rr, cc); return v == null ? '' : typeof v === 'string' ? v : displayText(sheet, rr, cc); };
  const attempt = (first: number): FlashPlan | null => {
    const ex: { src: string[]; out: string }[] = [];
    const fill: number[] = [];
    for (let rr = first; rr <= region.r2; rr++) {
      const src = srcCols.map((cc) => textOf(rr, cc));
      if (src.every((s) => s === '')) continue;
      const t = textOf(rr, c);
      if (t) ex.push({ src, out: t }); else fill.push(rr);
    }
    if (!ex.length || !fill.length) return null;
    const prog = inferFlashFill(ex.slice(0, 12));
    if (!prog) return null;
    if (ex.slice(12).some((e) => runFlashFill(prog, e.src) !== e.out)) return null;
    const rows: number[] = [], values: string[] = [];
    for (const rr of fill) { const v = runFlashFill(prog, srcCols.map((cc) => textOf(rr, cc))); if (v != null) { rows.push(rr); values.push(v); } }
    return rows.length ? { col: c, rows, values, examples: ex.length } : null;
  };
  region = { ...region };
  return attempt(region.r1) || (hasHeaderRow(sheet, region) ? attempt(region.r1 + 1) : null);
}

function applyFlash(app: App, plan: FlashPlan): void {
  const sheet = app.sheet;
  app.book.transact('Flash Fill', () => plan.rows.forEach((r, i) => { const cur = sheet.get(r, plan.col); sheet.set(r, plan.col, { v: plan.values[i], ...(cur?.s ? { s: cur.s } : {}) }); }));
  toast(`Flash Fill changed ${plan.rows.length} cell${plan.rows.length === 1 ? '' : 's'}.`, { action: 'Undo', onAction: () => app.exec('edit.undo') });
}

// --------------------------------------------------------------------------------------------
// Remove Duplicates

async function removeDuplicatesDialog(app: App): Promise<void> {
  if (blocked(app)) return;
  const sheet = app.sheet;
  const { r, c } = sheet.sel;
  const t = tableAt(sheet, r, c);
  const sel = activeRange(app);
  let rg: Range = t ? { ...t.range, r2: t.range.r2 - (t.opts.totals ? 1 : 0) } : sel.r1 === sel.r2 && sel.c1 === sel.c2 ? currentRegion(sheet, r, c) : clipToUsed(sheet, sel);
  if (regionIsEmpty(sheet, rg)) { toast('Select a range of data first.'); return; }
  let header = t ? t.opts.header : hasHeaderRow(sheet, rg);
  const colOpts = (hdr: boolean) => rowsOf(rg.c1, rg.c2).map((cc) => ({ value: String(cc), label: (hdr && displayText(sheet, rg.r1, cc)) || `Column ${colName(cc)}` }));
  const all = rowsOf(rg.c1, rg.c2).map(String);
  const res = await openDialog({
    title: 'Remove Duplicates',
    help: 'To delete duplicate values, select one or more columns that contain duplicates.',
    fields: [
      { key: 'selAll', type: 'button', label: 'Select All', onClick: (_v, ctl) => ctl.set('cols', all) },
      { key: 'unselAll', type: 'button', label: 'Unselect All', onClick: (_v, ctl) => ctl.set('cols', []) },
      { key: 'header', type: 'checkbox', label: 'My data has headers', value: header },
      { key: 'cols', type: 'list', label: 'Columns', multiple: true, rows: Math.min(10, all.length + 1), options: colOpts(header), value: all },
    ],
    onChange: (v, ctl, key) => { if (key === 'header') { setListOptions(ctl, 'cols', colOpts(!!v.header)); ctl.set('cols', all); } },
    validate: (v) => (!v.cols?.length ? 'Select at least one column.' : null),
  });
  if (!res) return;
  header = !!res.header;
  const cols = (res.cols as string[]).map(Number);
  const dups = duplicateRows(sheet, rg, cols, header);
  const total = rg.r2 - rg.r1 + 1 - (header ? 1 : 0);
  if (dups.length) {
    app.book.transact('Remove Duplicates', () => {
      removeRowsInRange(sheet, rg, dups, header);
      if (t) { sheet.book.recordSheet(sheet); const tot = t.opts.totals; if (tot) { /* move totals row up */ const from = t.range.r2, to = t.range.r2 - dups.length; for (let cc = t.range.c1; cc <= t.range.c2; cc++) { const cell = sheet.get(from, cc); sheet.set(to, cc, cell ? { ...cell } : undefined); sheet.set(from, cc, undefined); } } t.range = { ...t.range, r2: t.range.r2 - dups.length }; rg = { ...rg }; }
    });
  }
  void alertDialog('Remove Duplicates', dups.length ? `${dups.length} duplicate value${dups.length === 1 ? '' : 's'} found and removed; ${total - dups.length} unique value${total - dups.length === 1 ? '' : 's'} remain.` : 'No duplicate values found.');
}

// --------------------------------------------------------------------------------------------
// Data Validation

const DV_TYPES = [
  { value: 'any', label: 'Any value' }, { value: 'whole', label: 'Whole number' }, { value: 'decimal', label: 'Decimal' }, { value: 'list', label: 'List' },
  { value: 'date', label: 'Date' }, { value: 'time', label: 'Time' }, { value: 'textLength', label: 'Text length' }, { value: 'custom', label: 'Custom' },
];
const DV_OPS = [
  { value: 'between', label: 'between' }, { value: 'notBetween', label: 'not between' }, { value: 'equal', label: 'equal to' }, { value: 'notEqual', label: 'not equal to' },
  { value: 'greaterThan', label: 'greater than' }, { value: 'lessThan', label: 'less than' }, { value: 'greaterThanOrEqual', label: 'greater than or equal to' }, { value: 'lessThanOrEqual', label: 'less than or equal to' },
];

function dvAt(sheet: Sheet, r: number, c: number): DataValidation | undefined {
  return sheet.validations.find((v) => v.ranges.some((rg) => inRange(rg, r, c)));
}

/** Formula text shown in the dialog for a stored validation formula. */
function dvShow(f: string | undefined, type: string, book: Workbook): string {
  if (f == null) return '';
  if (type === 'list') { const t = f.trim(); if (t.startsWith('"') && t.endsWith('"')) return t.slice(1, -1); return t.startsWith('=') ? t : '=' + t; }
  if ((type === 'date' || type === 'time') && /^-?\d+(\.\d+)?$/.test(f)) return formatValue(+f, type === 'date' ? 'm/d/yyyy' : 'h:mm:ss AM/PM', { date1904: book.calc.date1904 }).text;
  return /^-?\d+(\.\d+)?$/.test(f) ? f : '=' + f.replace(/^=/, '');
}

/** Converts dialog text to a stored validation formula. */
export function dvStore(text: string, type: string, date1904 = false): string | undefined {
  const t = String(text ?? '').trim();
  if (!t) return undefined;
  if (t.startsWith('=')) return t.slice(1);
  if (type === 'list') return `"${t.replace(/"/g, '')}"`;
  if (type === 'date' || type === 'time') { const d = parseDateTime(t, { date1904 }); if (d) return String(type === 'time' ? d.serial % 1 : d.serial); }
  return t;
}

async function validationDialog(app: App): Promise<void> {
  if (blocked(app)) return;
  const sheet = app.sheet;
  const ranges = selRanges(app).map((rg) => ({ ...rg }));
  const { r, c } = sheet.sel;
  const found = new Set<DataValidation>();
  for (const rg of ranges) for (const dv of sheet.validations) if (dv.ranges.some((x) => !(x.r2 < rg.r1 || x.r1 > rg.r2 || x.c2 < rg.c1 || x.c1 > rg.c2))) found.add(dv);
  let cur = dvAt(sheet, r, c);
  if (found.size > 1) {
    if (!(await confirmDialog('Data Validation', 'The selection contains more than one type of validation. Erase current settings and continue?', 'Yes', 'No'))) return;
    cur = undefined;
  } else if (found.size === 1 && !cur) {
    if (!(await confirmDialog('Data Validation', 'The selection contains some cells without Data Validation settings. Do you want to extend Data Validation to these cells?', 'Yes', 'No'))) return;
    cur = [...found][0];
  }
  const book = app.book;
  const type0 = cur?.type || 'any';
  let cleared = false;
  const fields: Field[] = [
    { key: 'tab', type: 'radio', options: [{ value: 'settings', label: 'Settings' }, { value: 'input', label: 'Input Message' }, { value: 'error', label: 'Error Alert' }], value: 'settings' },
    { key: 'type', type: 'select', label: 'Allow', options: DV_TYPES, value: type0, visibleIf: (v) => v.tab === 'settings' },
    { key: 'allowBlank', type: 'checkbox', label: 'Ignore blank', value: cur?.allowBlank !== false, visibleIf: (v) => v.tab === 'settings' && v.type !== 'any' },
    { key: 'operator', type: 'select', label: 'Data', options: DV_OPS, value: cur?.operator || 'between', visibleIf: (v) => v.tab === 'settings' && !['any', 'list', 'custom'].includes(v.type) },
    { key: 'f1', type: 'text', label: 'Minimum', value: dvShow(cur?.formula1, type0, book), visibleIf: (v) => v.tab === 'settings' && v.type !== 'any' && v.type !== 'list' && v.type !== 'custom' && ['between', 'notBetween'].includes(v.operator) },
    { key: 'f2', type: 'text', label: 'Maximum', value: dvShow(cur?.formula2, type0, book), visibleIf: (v) => v.tab === 'settings' && v.type !== 'any' && v.type !== 'list' && v.type !== 'custom' && ['between', 'notBetween'].includes(v.operator) },
    { key: 'fv', type: 'text', label: 'Value', value: dvShow(cur?.formula1, type0, book), visibleIf: (v) => v.tab === 'settings' && !['any', 'list', 'custom'].includes(v.type) && !['between', 'notBetween'].includes(v.operator) },
    { key: 'src', type: 'text', label: 'Source', value: type0 === 'list' ? dvShow(cur?.formula1, 'list', book) : '', placeholder: 'Yes, No, Maybe   or   =$A$1:$A$10', visibleIf: (v) => v.tab === 'settings' && v.type === 'list' },
    { key: 'dropdown', type: 'checkbox', label: 'In-cell dropdown', value: cur?.showDropdown !== false, visibleIf: (v) => v.tab === 'settings' && v.type === 'list' },
    { key: 'formula', type: 'text', label: 'Formula', value: type0 === 'custom' ? dvShow(cur?.formula1, 'custom', book) : '', placeholder: '=ISNUMBER(A1)', visibleIf: (v) => v.tab === 'settings' && v.type === 'custom' },
    { key: 'applyAll', type: 'checkbox', label: 'Apply these changes to all other cells with the same settings', value: false, visibleIf: (v) => v.tab === 'settings' && !!cur },
    { key: 'showInput', type: 'checkbox', label: 'Show input message when cell is selected', value: !!(cur?.inputMessage || cur?.inputTitle) || !cur, visibleIf: (v) => v.tab === 'input' },
    { key: 'inputTitle', type: 'text', label: 'Title', value: cur?.inputTitle || '', visibleIf: (v) => v.tab === 'input' },
    { key: 'inputMessage', type: 'textarea', label: 'Input message', value: cur?.inputMessage || '', visibleIf: (v) => v.tab === 'input' },
    { key: 'errorStyle', type: 'select', label: 'Style', options: [{ value: 'stop', label: 'Stop' }, { value: 'warning', label: 'Warning' }, { value: 'information', label: 'Information' }], value: cur?.errorStyle || 'stop', visibleIf: (v) => v.tab === 'error' },
    { key: 'errorTitle', type: 'text', label: 'Title', value: cur?.errorTitle || '', visibleIf: (v) => v.tab === 'error' },
    { key: 'errorMessage', type: 'textarea', label: 'Error message', value: cur?.errorMessage || '', visibleIf: (v) => v.tab === 'error' },
  ];
  const res = await openDialog({
    title: 'Data Validation', fields,
    buttons: [{ label: 'Clear All', onClick: (_v, ctl) => { cleared = true; ctl.close({ __clear: true }); } }],
    validate: (v) => {
      if (v.type === 'list' && !String(v.src).trim()) return 'Enter a Source: a list of items separated by commas, or a range reference.';
      if (v.type === 'custom' && !String(v.formula).trim()) return 'Enter a formula that returns TRUE for valid entries.';
      if (!['any', 'list', 'custom'].includes(v.type)) {
        const two = ['between', 'notBetween'].includes(v.operator);
        if (two && (!String(v.f1).trim() || !String(v.f2).trim())) return 'Enter a Minimum and a Maximum.';
        if (!two && !String(v.fv).trim()) return 'Enter a Value.';
      }
      return null;
    },
  });
  if (!res) return;
  const sel = ranges;
  app.book.transact('Data Validation', () => {
    book.recordSheet(sheet);
    if (cleared || res.__clear) {
      for (const dv of sheet.validations) dv.ranges = dv.ranges.flatMap((x) => sel.reduce<Range[]>((acc, s) => acc.flatMap((a) => subtractRange(a, s)), [x]));
      sheet.validations = sheet.validations.filter((dv) => dv.ranges.length);
      return;
    }
    const type = res.type as DataValidation['type'];
    const two = ['between', 'notBetween'].includes(res.operator);
    const d1904 = book.calc.date1904;
    const spec: Partial<DataValidation> = {
      type,
      operator: ['any', 'list', 'custom'].includes(type) ? undefined : res.operator,
      formula1: type === 'list' ? dvStore(res.src, 'list') : type === 'custom' ? dvStore(String(res.formula).startsWith('=') ? res.formula : '=' + res.formula, 'custom') : type === 'any' ? undefined : dvStore(two ? res.f1 : res.fv, type, d1904),
      formula2: !['any', 'list', 'custom'].includes(type) && two ? dvStore(res.f2, type, d1904) : undefined,
      allowBlank: !!res.allowBlank,
      showDropdown: type === 'list' ? !!res.dropdown : undefined,
      inputTitle: res.showInput ? res.inputTitle || undefined : undefined,
      inputMessage: res.showInput ? res.inputMessage || undefined : undefined,
      errorStyle: res.errorStyle,
      errorTitle: res.errorTitle || undefined,
      errorMessage: res.errorMessage || undefined,
    };
    if (res.applyAll && cur) {
      const target = sheet.validations.find((v) => v.id === cur!.id);
      if (target) Object.assign(target, spec, { ranges: [...target.ranges, ...sel] });
      return;
    }
    for (const dv of sheet.validations) dv.ranges = dv.ranges.flatMap((x) => sel.reduce<Range[]>((acc, s) => acc.flatMap((a) => subtractRange(a, s)), [x]));
    sheet.validations = sheet.validations.filter((dv) => dv.ranges.length);
    if (type === 'any' && !spec.inputMessage && !spec.inputTitle) return;
    sheet.validations.push({ id: `dv${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`, ranges: sel, ...spec } as DataValidation);
  });
}

function circleInvalid(app: App): void {
  const sheet = app.sheet;
  const out: { r: number; c: number }[] = [];
  for (const dv of sheet.validations) {
    if (dv.type === 'any') continue;
    for (const rg0 of dv.ranges) {
      const rg = clipToUsed(sheet, rg0);
      sheet.forEachIn(rg, (r, c, cell) => {
        if (out.length >= 255) return;
        if (cell.v == null && cell.f == null) return;
        if (!validates(app, sheet, r, c, cell.v ?? null, dv)) out.push({ r, c });
      });
    }
  }
  app.tab.circles = out;
  try { app.grid.invalidate(); } catch { /* no grid */ }
  toast(out.length ? `${out.length} invalid entr${out.length === 1 ? 'y' : 'ies'} circled.${out.length >= 255 ? ' (Only the first 255 are shown.)' : ''}` : 'No invalid data was found.');
}

// --------------------------------------------------------------------------------------------
// Consolidate

const consolidateRefs = new WeakMap<Workbook, { refs: string[]; fn: AggName; top: boolean; left: boolean; links: boolean }>();

async function consolidateDialog(app: App): Promise<void> {
  if (blocked(app)) return;
  const book = app.book, sheet = app.sheet;
  const prev = consolidateRefs.get(book);
  const refs: string[] = prev ? [...prev.refs] : [];
  const res = await openDialog({
    title: 'Consolidate',
    help: `Combines data from several ranges into the area starting at ${cellName(sheet.sel.r, sheet.sel.c)}.`,
    fields: [
      { key: 'fn', type: 'select', label: 'Function', options: [...AGG_FUNCS], value: prev?.fn || 'Sum' },
      { key: 'ref', type: 'text', label: 'Reference', value: rangeRefText(sheet, clipToUsed(sheet, activeRange(app))), placeholder: "Sheet2!$A$1:$C$10" },
      { key: 'add', type: 'button', label: 'Add', onClick: (v, ctl) => { const t = String(v.ref || '').trim(); if (!t) return; if (!resolveRangeText(book, t, sheet)) { toast('Reference isn\'t valid.'); return; } if (!refs.includes(t)) refs.push(t); setListOptions(ctl, 'refs', refs); } },
      { key: 'refs', type: 'list', label: 'All references', options: refs, rows: 5 },
      { key: 'del', type: 'button', label: 'Delete', onClick: (v, ctl) => { const i = refs.indexOf(v.refs); if (i >= 0) refs.splice(i, 1); setListOptions(ctl, 'refs', refs); } },
      { key: 'top', type: 'checkbox', label: 'Use labels in top row', value: !!prev?.top, half: true },
      { key: 'left', type: 'checkbox', label: 'Use labels in left column', value: !!prev?.left, half: true },
      { key: 'links', type: 'checkbox', label: 'Create links to source data', value: !!prev?.links },
    ],
    validate: () => (!refs.length ? 'Add at least one reference.' : null),
  });
  if (!res) return;
  consolidateRefs.set(book, { refs, fn: res.fn, top: !!res.top, left: !!res.left, links: !!res.links });
  const sources = refs.map((t) => {
    const x = resolveRangeText(book, t, sheet)!;
    const rg = clipToUsed(x.sheet, x.range);
    const values: Scalar[][] = [];
    for (let r = rg.r1; r <= rg.r2; r++) { const row: Scalar[] = []; for (let c = rg.c1; c <= rg.c2; c++) row.push(x.sheet.value(r, c)); values.push(row); }
    return { values, ref: (i: number, j: number) => rangeRefText(x.sheet, { r1: rg.r1 + i, c1: rg.c1 + j, r2: rg.r1 + i, c2: rg.c1 + j }, true, sheet) };
  });
  const { out, refs: cellRefs } = consolidate(sources, res.fn as AggName, !!res.top, !!res.left);
  const r0 = sheet.sel.r, c0 = sheet.sel.c;
  book.transact('Consolidate', () => {
    out.forEach((row, i) => row.forEach((v, j) => {
      const rr = r0 + i, cc = c0 + j;
      const cur = sheet.get(rr, cc);
      const s = cur?.s;
      const rf = cellRefs[i]?.[j];
      if (res.links && rf && rf.length) sheet.set(rr, cc, { f: `${aggFormulaName(res.fn as AggName)}(${rf.join(',')})`, ...(s ? { s } : {}) });
      else if (v == null) sheet.set(rr, cc, s ? { s } : undefined);
      else sheet.set(rr, cc, { v, ...(s ? { s } : {}) });
    }));
  });
}

// --------------------------------------------------------------------------------------------
// Data type helpers

function convertSelection(app: App, to: 'number' | 'date' | 'text'): void {
  if (blocked(app)) return;
  const sheet = app.sheet, book = app.book;
  let n = 0;
  book.transact(`Convert to ${to}`, () => {
    for (const rg0 of selRanges(app)) sheet.forEachIn(clipToUsed(sheet, rg0), (r, c, cell) => {
      if (cell.f != null || cell.v == null) return;
      const v = cell.v;
      if (to === 'number' && typeof v === 'string') {
        const t = v.trim().replace(/^'/, '').replace(/[ \s]/g, '');
        const neg = /^\((.*)\)$/.exec(t) || /^(.*)-$/.exec(t);
        const num = Number((neg ? '-' + neg[1] : t).replace(/[$€£¥,]/g, ''));
        const pct = /%$/.test(t) ? Number(t.slice(0, -1).replace(/[$,]/g, '')) / 100 : null;
        const x = pct != null && Number.isFinite(pct) ? pct : num;
        if (t !== '' && Number.isFinite(x)) { sheet.set(r, c, { ...cell, v: x }); n++; }
      } else if (to === 'date' && typeof v === 'string') {
        const d = parseDateTime(v.trim(), { dmy: app.prefs.dateOrder === 'dmy', date1904: book.calc.date1904 });
        if (d) { const st = sheet.styleOf(r, c); sheet.set(r, c, { ...cell, v: d.serial, s: isDateFormat(st.numFmt) ? cell.s : book.derivedStyle(cell.s ?? 0, { numFmt: d.format }) }); n++; }
      } else if (to === 'text' && typeof v !== 'string' && !(v instanceof CellError)) {
        const t = displayText(sheet, r, c);
        sheet.set(r, c, { ...cell, v: t, s: book.derivedStyle(cell.s ?? 0, { numFmt: '@' }) }); n++;
      }
    });
  });
  toast(n ? `Converted ${n} cell${n === 1 ? '' : 's'}.` : 'Nothing in the selection needed converting.');
}

// --------------------------------------------------------------------------------------------
// Outline

async function outlineAxis(app: App, arg: any, verb: string): Promise<{ axis: OutlineAxis; a: number; b: number } | null> {
  const rg = activeRange(app);
  if (arg?.axis) return arg.axis === 'col' ? { axis: 'col', a: rg.c1, b: rg.c2 } : { axis: 'row', a: rg.r1, b: rg.r2 };
  if (isWholeRows(rg) && !isWholeCols(rg)) return { axis: 'row', a: rg.r1, b: rg.r2 };
  if (isWholeCols(rg) && !isWholeRows(rg)) return { axis: 'col', a: rg.c1, b: rg.c2 };
  const res = await openDialog({ title: verb, fields: [{ key: 'axis', type: 'radio', label: verb, options: [{ value: 'row', label: 'Rows' }, { value: 'col', label: 'Columns' }], value: rg.c2 - rg.c1 > rg.r2 - rg.r1 ? 'col' : 'row' }] });
  if (!res) return null;
  return res.axis === 'col' ? { axis: 'col', a: rg.c1, b: Math.min(rg.c2, MAX_COLS - 1) } : { axis: 'row', a: rg.r1, b: Math.min(rg.r2, MAX_ROWS - 1) };
}

async function groupCmd(app: App, arg: any, delta: 1 | -1): Promise<void> {
  if (blocked(app, delta > 0 ? 'formatRows' : 'formatRows')) return;
  const t = await outlineAxis(app, arg, delta > 0 ? 'Group' : 'Ungroup');
  if (!t) return;
  const sheet = app.sheet;
  const lv = t.axis === 'row' ? sheet.rowLevel : sheet.colLevel;
  if (delta > 0) { for (let i = t.a; i <= t.b; i++) if ((lv.get(i) || 0) >= 8) { toast('Outlines can have at most 8 levels.'); return; } }
  let changed = false;
  app.book.transact(delta > 0 ? 'Group' : 'Ungroup', () => { sheet.book.recordSheet(sheet); changed = changeLevel(sheet, t.axis, t.a, Math.min(t.b, t.a + 1048575), delta); });
  if (!changed && delta < 0) toast('The selection isn\'t part of a group.');
}

function detailCmd(app: App, show: boolean, arg?: any): void {
  const sheet = app.sheet;
  const rg = activeRange(app);
  const axis: OutlineAxis = arg?.axis || (isWholeCols(rg) && !isWholeRows(rg) ? 'col' : 'row');
  let ok = false;
  app.book.transact(show ? 'Show Detail' : 'Hide Detail', () => {
    sheet.book.recordSheet(sheet);
    ok = setDetail(sheet, axis, axis === 'row' ? sheet.sel.r : sheet.sel.c, show);
    if (!ok && axis === 'row') ok = setDetail(sheet, 'col', sheet.sel.c, show);
  });
  if (!ok) toast(show ? 'There is no hidden detail for this cell. Select a summary cell of a group.' : 'This cell isn\'t in a group. Use Group first, or Subtotal to build an outline.');
}

function outlineLevelMenu(app: App): MenuEntry[] {
  const s = app.sheet;
  const out: MenuEntry[] = [];
  const mr = maxLevel(s, 'row'), mc = maxLevel(s, 'col');
  if (mr) { out.push({ heading: 'Row levels' }); for (let i = 1; i <= mr + 1; i++) out.push({ id: 'data.outlineLevel', label: `Show row level ${i}${i === mr + 1 ? ' (all)' : ''}`, arg: { axis: 'row', level: i } }); }
  if (mc) { out.push({ heading: 'Column levels' }); for (let i = 1; i <= mc + 1; i++) out.push({ id: 'data.outlineLevel', label: `Show column level ${i}${i === mc + 1 ? ' (all)' : ''}`, arg: { axis: 'col', level: i } }); }
  if (!out.length) out.push({ heading: 'This sheet has no outline' });
  return out;
}

async function outlineSettings(app: App): Promise<void> {
  const s = app.sheet;
  const mr = maxLevel(s, 'row'), mc = maxLevel(s, 'col');
  const lvBtns = (axis: OutlineAxis, m: number): Field[] => rowsOf(1, m + 1).map((i) => ({ key: `${axis}${i}`, type: 'button' as const, label: `${axis === 'row' ? 'Rows' : 'Columns'}: level ${i}`, half: true, onClick: (_v: any, ctl: DialogControl) => { app.exec('data.outlineLevel', { axis, level: i }); ctl.close({}); } }));
  await openDialog({
    title: 'Outline Settings', ok: 'Close', cancel: false,
    help: 'Summary rows are below their detail and summary columns to the right, as Subtotal and Auto Outline create them. Use the level buttons to show only the summaries you need.',
    fields: [
      ...(mr ? [{ key: 'h1', type: 'heading' as const, label: `Row outline (${mr} level${mr === 1 ? '' : 's'})` }, ...lvBtns('row', mr)] : []),
      ...(mc ? [{ key: 'h2', type: 'heading' as const, label: `Column outline (${mc} level${mc === 1 ? '' : 's'})` }, ...lvBtns('col', mc)] : []),
      ...(!mr && !mc ? [{ key: 'none', type: 'info' as const, label: 'This sheet has no outline yet.' }] : []),
      { key: 'h3', type: 'heading', label: 'Tools' },
      { key: 'auto', type: 'button', label: 'Create (Auto Outline)', half: true, onClick: (_v, ctl) => { ctl.close({}); app.exec('data.autoOutline'); } },
      { key: 'clr', type: 'button', label: 'Clear Outline', half: true, onClick: (_v, ctl) => { ctl.close({}); app.exec('data.clearOutline'); } },
      { key: 'show', type: 'button', label: 'Show Detail', half: true, onClick: (_v, ctl) => { ctl.close({}); app.exec('data.showDetail'); } },
      { key: 'hide', type: 'button', label: 'Hide Detail', half: true, onClick: (_v, ctl) => { ctl.close({}); app.exec('data.hideDetail'); } },
    ],
  });
}

async function subtotalDialog(app: App): Promise<void> {
  if (blocked(app)) return;
  const sheet = app.sheet;
  const { r, c } = sheet.sel;
  if (tableAt(sheet, r, c)) { void alertDialog('Subtotal', 'Subtotals can\'t be added inside a table. Convert the table to a range first (Table Design ▸ Convert to Range).'); return; }
  const sel = activeRange(app);
  const list = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? currentRegion(sheet, r, c) : clipToUsed(sheet, sel);
  if (regionIsEmpty(sheet, list) || list.r2 <= list.r1) { toast('Select a cell in a list with a header row first.'); return; }
  const cols = rowsOf(list.c1, list.c2).map((cc) => ({ value: String(cc), label: displayText(sheet, list.r1, cc) || `Column ${colName(cc)}` }));
  const numericCols = rowsOf(list.c1, list.c2).filter((cc) => typeof sheet.value(list.r1 + 1, cc) === 'number').map(String);
  let removeAll = false;
  const res = await openDialog({
    title: 'Subtotal',
    fields: [
      { key: 'group', type: 'select', label: 'At each change in', options: cols, value: String(c) },
      { key: 'fn', type: 'select', label: 'Use function', options: SUBTOTAL_FUNCS.map((f) => ({ value: String(f.code), label: f.name })), value: '9' },
      { key: 'cols', type: 'list', label: 'Add subtotal to', multiple: true, rows: Math.min(8, cols.length + 1), options: cols, value: numericCols.length ? [numericCols[numericCols.length - 1]] : [String(list.c2)] },
      { key: 'replace', type: 'checkbox', label: 'Replace current subtotals', value: true },
      { key: 'pageBreaks', type: 'checkbox', label: 'Page break between groups', value: false },
      { key: 'below', type: 'checkbox', label: 'Summary below data', value: true },
    ],
    buttons: [{ label: 'Remove All', onClick: (_v, ctl) => { removeAll = true; ctl.close({}); } }],
    validate: (v) => (!v.cols?.length ? 'Select at least one column to subtotal.' : null),
  });
  if (!res) return;
  const book = app.book;
  const shift = (at: number, n: number) => shiftCells(sheet, 'row', at, n);
  if (removeAll) {
    book.transact('Remove Subtotals', () => { removeSubtotals(sheet, list, shift); });
    return;
  }
  book.transact('Subtotal', () => {
    const out = addSubtotals(sheet, list, { groupCol: +res.group, fn: +res.fn, cols: (res.cols as string[]).map(Number), replace: !!res.replace, pageBreaks: !!res.pageBreaks, below: !!res.below, boldStyle: boldStyle(book, sheet.styleIdOf(list.r1 + 1, +res.group)) }, shift);
    sheet.sel = { r: out.r1, c: out.c1, ranges: [out] };
  });
}

// --------------------------------------------------------------------------------------------
// Get & Transform

interface StoredQuery { kind: 'web' | 'text'; name: string; url?: string; table?: number; sheet: string; r: number; c: number; rows: number; cols: number; time: number }
const Q_PREFIX = 'QuillLattice.Query.';

export function storedQueries(book: Workbook): StoredQuery[] {
  return Object.entries(book.props.custom || {}).filter(([k]) => k.startsWith(Q_PREFIX)).map(([, v]) => { try { return JSON.parse(v) as StoredQuery; } catch { return null; } }).filter((x): x is StoredQuery => !!x);
}

function saveQuery(book: Workbook, q: StoredQuery | null, name: string): void {
  book.recordBook();
  const custom = { ...(book.props.custom || {}) };
  if (q) custom[Q_PREFIX + name] = JSON.stringify(q); else delete custom[Q_PREFIX + name];
  book.props = { ...book.props, custom };
}

/** Copies cells from another workbook's sheet (styles re-interned) to (r0, c0). Inside a transaction. */
function copySheetCells(src: Sheet, dest: Sheet, r0: number, c0: number): { rows: number; cols: number } {
  const book = dest.book;
  const map = new Map<number, number>();
  const sid = (s: number | undefined) => { if (!s) return undefined; let x = map.get(s); if (x == null) { x = book.styleId(src.book.styles[s] || {}); map.set(s, x); } return x || undefined; };
  let rows = 0, cols = 0;
  for (const [k, cell] of src.cells) {
    const r = keyRow(k), c = keyCol(k);
    const next: Cell = { ...cell };
    delete next.sp; delete next.arr;
    const s = sid(cell.s);
    if (s) next.s = s; else delete next.s;
    dest.set(r0 + r, c0 + c, next);
    rows = Math.max(rows, r + 1); cols = Math.max(cols, c + 1);
  }
  for (const m of src.merges) dest.merges.push({ r1: m.r1 + r0, c1: m.c1 + c0, r2: m.r2 + r0, c2: m.c2 + c0 });
  for (const [c, w] of src.colW) if (!dest.colW.has(c + c0)) dest.colW.set(c + c0, w);
  return { rows, cols };
}

async function chooseDestination(app: App, title: string, extra: Field[] = []): Promise<Record<string, any> | null> {
  return openDialog({
    title,
    fields: [
      ...extra,
      { key: 'where', type: 'radio', label: 'Where do you want to put the data?', options: [{ value: 'new', label: 'New worksheet' }, { value: 'here', label: 'Existing worksheet, starting at' }], value: 'new' },
      { key: 'at', type: 'text', label: 'Cell', value: cellName(app.sheet.sel.r, app.sheet.sel.c, true, true), visibleIf: (v) => v.where === 'here' },
    ],
  });
}

async function fromText(app: App): Promise<void> {
  const { openDocument } = await import('../platform/native');
  const { parseCsv, detectDelimiter, decodeText } = await import('../io/index');
  const { fillSheet } = await import('../io/csv');
  const f = await openDocument('.csv,.tsv,.txt,.tab,.prn,.dat', ['text/csv', 'text/plain', 'text/tab-separated-values']);
  if (!f) return;
  const encs = [{ value: 'utf-8', label: 'Unicode (UTF-8)' }, { value: 'windows-1252', label: 'Western European (Windows)' }, { value: 'iso-8859-1', label: 'Western European (ISO)' }, { value: 'utf-16le', label: 'Unicode (UTF-16 LE)' }, { value: 'utf-16be', label: 'Unicode (UTF-16 BE)' }, { value: 'shift_jis', label: 'Japanese (Shift-JIS)' }, { value: 'gbk', label: 'Chinese Simplified (GBK)' }, { value: 'macintosh', label: 'Western European (Mac)' }];
  let text = decodeText(f.data);
  const det = detectDelimiter(text);
  const delimOpts = [{ value: ',', label: 'Comma' }, { value: ';', label: 'Semicolon' }, { value: '\t', label: 'Tab' }, { value: '|', label: 'Pipe' }, { value: ' ', label: 'Space' }, { value: 'other', label: 'Other…' }];
  const types: string[] = [];
  const preview = h('div', { class: 'ql-data-preview' });
  const parse = (v: Record<string, any>) => {
    const d = v.delim === 'other' ? String(v.other || ',')[0] : v.delim;
    let rows = parseCsv(text, { delimiter: d, quote: v.quote === 'none' ? '' : v.quote });
    const skip = Math.max(0, (+v.start || 1) - 1);
    rows = rows.slice(skip);
    return rows;
  };
  const render = (v: Record<string, any>) => {
    const rows = parse(v).slice(0, 15);
    const n = Math.max(1, ...rows.map((r) => r.length));
    while (types.length < n) types.push('general');
    clear(preview);
    const table = h('table', { class: 'ql-data-table' });
    const head = h('tr');
    for (let j = 0; j < n; j++) {
      const s = h('select', { 'aria-label': `Column ${j + 1} type` }) as HTMLSelectElement;
      for (const [val, lab] of [['general', 'General'], ['text', 'Text'], ['date', 'Date'], ['skip', 'Skip']]) s.appendChild(h('option', { value: val }, lab));
      s.value = types[j];
      s.addEventListener('change', () => { types[j] = s.value; render(v); });
      head.appendChild(h('th', null, s));
    }
    table.appendChild(head);
    rows.forEach((row, i) => { const tr = h('tr', { class: i === 0 && v.header ? 'ql-data-hdr' : '' }); for (let j = 0; j < n; j++) tr.appendChild(h('td', { class: types[j] === 'skip' ? 'ql-data-skip' : '' }, row[j] ?? '')); table.appendChild(tr); });
    preview.appendChild(table);
  };
  const res = await openDialog({
    title: `Import ${f.name}`, wide: true,
    fields: [
      { key: 'enc', type: 'select', label: 'File origin (encoding)', options: encs, value: 'utf-8', half: true },
      { key: 'delim', type: 'select', label: 'Delimiter', options: delimOpts, value: delimOpts.some((o) => o.value === det) ? det : ',', half: true },
      { key: 'other', type: 'text', label: 'Other delimiter', value: '', visibleIf: (v) => v.delim === 'other', width: '4em' },
      { key: 'quote', type: 'select', label: 'Text qualifier', options: [{ value: '"', label: '"' }, { value: "'", label: "'" }, { value: 'none', label: '{none}' }], value: '"', half: true },
      { key: 'start', type: 'number', label: 'Start import at row', value: 1, min: 1, half: true },
      { key: 'header', type: 'checkbox', label: 'My data has headers (keep the first row as text)', value: true },
      { key: 'dateOrder', type: 'select', label: 'Date order', options: [{ value: 'mdy', label: 'Month/Day/Year' }, { value: 'dmy', label: 'Day/Month/Year' }, { value: 'ymd', label: 'Year/Month/Day' }], value: app.prefs.dateOrder || 'mdy', half: true },
      { key: 'where', type: 'radio', label: 'Put the data in', options: [{ value: 'new', label: 'New worksheet' }, { value: 'here', label: 'Existing worksheet at the active cell' }], value: 'new' },
    ],
    content: h('div', null, h('div', { class: 'ql-label' }, 'Preview (choose a type above each column)'), preview),
    onChange: (v, _ctl, key) => { if (key === 'enc') text = decodeText(f.data, v.enc); render(v); },
  });
  if (!res) return;
  const rows = parse(res);
  const opts: any = { delimiter: res.delim === 'other' ? String(res.other || ',')[0] : res.delim, headerRows: res.header ? 1 : 0, columnTypes: types, dateOrder: res.dateOrder };
  const book = app.book;
  let target: Sheet = app.sheet;
  book.transact('Import Text', () => {
    if (res.where === 'new') target = addSheetAfter(book, f.name.replace(/\.[^.]*$/, ''));
    else book.recordSheet(target);
    const r0 = res.where === 'new' ? 0 : target.sel.r, c0 = res.where === 'new' ? 0 : target.sel.c;
    fillSheet(target, rows, opts, r0, c0);
    if (res.header) { const n = Math.max(...rows.slice(0, 1).map((r) => r.length), 0); for (let j = 0; j < n; j++) { const cur = target.get(r0, c0 + j); if (cur) target.cells.set(cellKey(r0, c0 + j), { ...cur, s: book.derivedStyle(cur.s ?? 0, { bold: true }) }); } }
  });
  if (res.where === 'new') activateSheet(app, target);
  toast(`Imported ${rows.length} rows.`);
}

async function fetchText(url: string): Promise<{ text: string; type: string }> {
  let resp: Response;
  try { resp = await fetch(url); } catch {
    throw new Error('The page could not be loaded. Many websites block requests from other apps (CORS), and offline devices cannot reach them. Try the desktop app, download the page or CSV and use From Text/CSV, or copy the table and use From Clipboard.');
  }
  if (!resp.ok) throw new Error(`The server answered ${resp.status} ${resp.statusText}.`);
  return { text: await resp.text(), type: resp.headers.get('content-type') || '' };
}

/** Loads a web table/CSV as a temporary workbook; returns its sheets. */
async function loadWeb(url: string): Promise<Workbook> {
  const { text, type } = await fetchText(url);
  if (/csv|text\/plain/i.test(type) || /\.(csv|txt|tsv)(\?|$)/i.test(url)) {
    const { parseCsv } = await import('../io/index');
    const { rowsToBook } = await import('../io/csv');
    return rowsToBook(parseCsv(text), {}, 'Data');
  }
  const { importHtml } = await import('../io/html');
  return importHtml(text, [], { allTables: true });
}

async function fromWeb(app: App): Promise<void> {
  const url = await promptDialog('From Web', 'URL', 'https://', { placeholder: 'https://example.com/page-with-a-table' });
  if (!url || !/^https?:\/\//i.test(url.trim())) { if (url) toast('Enter an address that starts with http:// or https://'); return; }
  let wb: Workbook;
  try { status(app, 'Loading…'); wb = await loadWeb(url.trim()); } catch (e) { void alertDialog('From Web', (e as Error).message); return; }
  const tables = wb.sheets.filter((s) => s.cells.size);
  if (!tables.length) { void alertDialog('From Web', 'No tables were found on that page.'); return; }
  const opts = tables.map((s, i) => { const u = s.usedRange(); return { value: String(i), label: `${s.name} — ${u ? `${u.r2 + 1} rows × ${u.c2 + 1} columns` : 'empty'}` }; });
  const preview = h('div', { class: 'ql-data-preview' });
  const showPrev = (i: number) => {
    clear(preview);
    const s = tables[i]; const u = s.usedRange();
    if (!u) return;
    const t = h('table', { class: 'ql-data-table' });
    for (let r = 0; r <= Math.min(u.r2, 9); r++) { const tr = h('tr'); for (let c = 0; c <= Math.min(u.c2, 9); c++) tr.appendChild(h('td', null, displayText(s, r, c))); t.appendChild(tr); }
    preview.appendChild(t);
  };
  const res = await openDialog({
    title: 'Navigator', wide: true,
    fields: [
      { key: 'table', type: 'list', label: 'Tables on the page', options: opts, value: '0', rows: Math.min(8, opts.length + 1) },
      { key: 'where', type: 'radio', label: 'Load to', options: [{ value: 'new', label: 'New worksheet' }, { value: 'here', label: 'Existing worksheet at the active cell' }], value: 'new' },
      { key: 'keep', type: 'checkbox', label: 'Save this query so Refresh All reloads it', value: true },
    ],
    content: h('div', null, h('div', { class: 'ql-label' }, 'Preview'), preview),
    onChange: (v) => showPrev(+(v.table ?? 0)),
  });
  if (!res) return;
  const idx = +(res.table ?? 0);
  const book = app.book;
  let target: Sheet = app.sheet;
  let size = { rows: 0, cols: 0 };
  const r0 = res.where === 'new' ? 0 : app.sheet.sel.r, c0 = res.where === 'new' ? 0 : app.sheet.sel.c;
  book.transact('From Web', () => {
    if (res.where === 'new') target = addSheetAfter(book, tables[idx].name);
    else book.recordSheet(target);
    size = copySheetCells(tables[idx], target, r0, c0);
    if (res.keep) {
      const taken = new Set(storedQueries(book).map((x) => x.name.toLowerCase()));
      let n = 1;
      while (taken.has(`web ${n}`)) n++;
      const name = `Web ${n}`;
      saveQuery(book, { kind: 'web', name, url: url.trim(), table: wb.sheets.indexOf(tables[idx]), sheet: target.name, r: r0, c: c0, rows: size.rows, cols: size.cols, time: Date.now() }, name);
    }
  });
  if (res.where === 'new') activateSheet(app, target);
}

async function refreshQuery(app: App, q: StoredQuery): Promise<boolean> {
  if (q.kind !== 'web' || !q.url) return false;
  const book = app.book;
  const sheet = book.sheetByName(q.sheet);
  if (!sheet) { toast(`The sheet "${q.sheet}" of query ${q.name} no longer exists.`); return false; }
  let wb: Workbook;
  try { wb = await loadWeb(q.url); } catch (e) { toast(`${q.name}: ${(e as Error).message}`); return false; }
  const src = wb.sheets[q.table ?? 0] || wb.sheets[0];
  if (!src) return false;
  book.transact('Refresh', () => {
    book.recordSheet(sheet);
    for (let r = q.r; r < q.r + q.rows; r++) for (let c = q.c; c < q.c + q.cols; c++) if (sheet.get(r, c)) sheet.set(r, c, undefined);
    const size = copySheetCells(src, sheet, q.r, q.c);
    saveQuery(book, { ...q, rows: size.rows, cols: size.cols, time: Date.now() }, q.name);
  });
  return true;
}

async function refreshAll(app: App): Promise<void> {
  const qs = storedQueries(app.book);
  let n = 0;
  for (const q of qs) if (await refreshQuery(app, q)) n++;
  app.calculate(true);
  status(app, qs.length ? `Refreshed ${n} of ${qs.length} quer${qs.length === 1 ? 'y' : 'ies'} and recalculated.` : 'Recalculated all formulas.');
  if (qs.length) toast(`Refreshed ${n} of ${qs.length} quer${qs.length === 1 ? 'y' : 'ies'}.`);
}

async function connectionsDialog(app: App): Promise<void> {
  const book = app.book;
  const qs = storedQueries(book);
  if (!qs.length) { void alertDialog('Queries & Connections', 'This workbook has no saved queries or connections.\nUse Data ▸ Get Data ▸ From Web to load a table that Refresh All can reload.'); return; }
  const opts = () => storedQueries(book).map((q) => ({ value: q.name, label: `${q.name} — ${q.url} → ${q.sheet}!${cellName(q.r, q.c)} (${q.rows}×${q.cols}, ${new Date(q.time).toLocaleString()})` }));
  await openDialog({
    title: 'Queries & Connections', ok: 'Close', cancel: false, wide: true,
    fields: [
      { key: 'q', type: 'list', label: 'Queries', options: opts(), rows: 6, value: qs[0].name },
      { key: 'refresh', type: 'button', label: 'Refresh', half: true, onClick: async (v) => { const q = storedQueries(book).find((x) => x.name === v.q); if (q && (await refreshQuery(app, q))) toast(`Refreshed ${q.name}.`); } },
      { key: 'del', type: 'button', label: 'Delete', half: true, onClick: (v, ctl) => { if (!v.q) return; book.transact('Delete Query', () => saveQuery(book, null, v.q)); setListOptions(ctl, 'q', opts()); } },
    ],
  });
}

async function fromWorkbook(app: App): Promise<void> {
  const { openDocument, OPEN_ACCEPT, OPEN_MIMES } = await import('../platform/native');
  const { importFile } = await import('../io/index');
  const f = await openDocument(OPEN_ACCEPT, OPEN_MIMES);
  if (!f) return;
  let wb: Workbook;
  try { wb = (await importFile(f.data, f.name)).book; } catch (e) { void alertDialog('From Workbook', (e as Error).message); return; }
  const res = await chooseDestination(app, 'Import From Workbook', [{ key: 'sheet', type: 'list', label: 'Sheet', options: wb.sheets.map((s, i) => ({ value: String(i), label: s.name })), value: '0', rows: Math.min(8, wb.sheets.length + 1) }]);
  if (!res) return;
  const src = wb.sheets[+(res.sheet ?? 0)];
  const book = app.book;
  let target = app.sheet;
  const at = resolveRangeText(book, res.at || 'A1', app.sheet);
  book.transact('Import Workbook', () => {
    if (res.where === 'new') target = addSheetAfter(book, src.name);
    else { if (at) target = at.sheet; book.recordSheet(target); }
    copySheetCells(src, target, res.where === 'new' ? 0 : at?.range.r1 ?? 0, res.where === 'new' ? 0 : at?.range.c1 ?? 0);
  });
  if (res.where === 'new') activateSheet(app, target);
}

async function fromClipboard(app: App): Promise<void> {
  const { readClipboardText } = await import('../platform/native');
  const { parseCsv, detectDelimiter } = await import('../io/index');
  const { fillSheet } = await import('../io/csv');
  let text: string | null = null;
  try { text = await readClipboardText(); } catch { text = null; }
  if (!text || !text.trim()) { void alertDialog('From Clipboard', 'The clipboard is empty, or the app isn\'t allowed to read it. Copy some text or a table and try again (or paste with Ctrl+V).'); return; }
  const d = text.includes('\t') ? '\t' : detectDelimiter(text);
  const rows = parseCsv(text, { delimiter: d });
  const res = await chooseDestination(app, `From Clipboard (${rows.length} rows)`);
  if (!res) return;
  const book = app.book;
  let target = app.sheet;
  const at = resolveRangeText(book, res.at || 'A1', app.sheet);
  book.transact('From Clipboard', () => {
    if (res.where === 'new') target = addSheetAfter(book, 'Clipboard');
    else { if (at) target = at.sheet; book.recordSheet(target); }
    fillSheet(target, rows, { delimiter: d }, res.where === 'new' ? 0 : at?.range.r1 ?? 0, res.where === 'new' ? 0 : at?.range.c1 ?? 0);
  });
  if (res.where === 'new') activateSheet(app, target);
}

function tableDataOf(sheet: Sheet, rg: Range, header: boolean): TableData {
  const headers = rowsOf(rg.c1, rg.c2).map((c, j) => (header ? displayText(sheet, rg.r1, c) : '') || `Column${j + 1}`);
  const seen = new Map<string, number>();
  for (let i = 0; i < headers.length; i++) { const k = headers[i].toLowerCase(); const n = seen.get(k) || 0; if (n) headers[i] = `${headers[i]}_${n + 1}`; seen.set(k, n + 1); }
  const rows: Scalar[][] = [];
  for (let r = rg.r1 + (header ? 1 : 0); r <= rg.r2; r++) rows.push(rowsOf(rg.c1, rg.c2).map((c) => sheet.value(r, c)));
  return { headers, rows };
}

async function fromTable(app: App): Promise<void> {
  const sheet = app.sheet;
  const { r, c } = sheet.sel;
  const t = tableAt(sheet, r, c);
  const sel = activeRange(app);
  const rg = t ? { ...t.range, r2: t.range.r2 - (t.opts.totals ? 1 : 0) } : sel.r1 === sel.r2 && sel.c1 === sel.c2 ? currentRegion(sheet, r, c) : clipToUsed(sheet, sel);
  if (regionIsEmpty(sheet, rg)) { toast('Select a cell in a table or range of data first.'); return; }
  const base = tableDataOf(sheet, rg, t ? t.opts.header : true);
  const steps: TransformStep[] = [];
  const preview = h('div', { class: 'ql-data-preview' });
  const stepsEl = h('ol', { class: 'ql-data-steps' });
  let cur = base;
  const renderPrev = () => {
    clear(preview); clear(stepsEl);
    stepsEl.appendChild(h('li', null, `Source: ${t ? t.name : rangeRefText(sheet, rg)}`));
    for (const s of steps) stepsEl.appendChild(h('li', null, describeStep(s)));
    const tbl = h('table', { class: 'ql-data-table' });
    tbl.appendChild(h('tr', { class: 'ql-data-hdr' }, ...cur.headers.map((x) => h('th', null, x))));
    for (const row of cur.rows.slice(0, 12)) tbl.appendChild(h('tr', null, ...row.map((v) => h('td', null, v == null ? '' : v instanceof CellError ? v.code : String(v)))));
    preview.appendChild(tbl);
    preview.appendChild(h('div', { class: 'ql-help' }, `${cur.rows.length} rows, ${cur.headers.length} columns`));
  };
  const kinds = [
    { value: 'filter', label: 'Filter rows' }, { value: 'removeCols', label: 'Remove columns' }, { value: 'keepCols', label: 'Keep only columns' }, { value: 'changeType', label: 'Change type' },
    { value: 'split', label: 'Split column by delimiter' }, { value: 'unpivot', label: 'Unpivot other columns' }, { value: 'groupBy', label: 'Group by' }, { value: 'sort', label: 'Sort' },
    { value: 'removeDuplicates', label: 'Remove duplicate rows' }, { value: 'removeBlankRows', label: 'Remove blank rows' }, { value: 'rename', label: 'Rename column' },
  ];
  const fieldsFor = (): Field[] => {
    const cols = cur.headers;
    const one = (v: any) => ['filter', 'changeType', 'split', 'sort', 'rename', 'groupBy'].includes(v.kind);
    return [
      { key: 'kind', type: 'select', label: 'Add a step', options: kinds, value: 'filter' },
      { key: 'col', type: 'select', label: 'Column', options: cols, value: cols[0], visibleIf: one, half: true },
      { key: 'op', type: 'select', label: 'Condition', options: OPS_TEXT.map(([v, l]) => ({ value: v, label: l })).concat([{ value: 'blank', label: 'is blank' }, { value: 'notBlank', label: 'is not blank' }]), value: 'eq', visibleIf: (v) => v.kind === 'filter', half: true },
      { key: 'value', type: 'text', label: 'Value', value: '', visibleIf: (v) => v.kind === 'filter' },
      { key: 'cols', type: 'list', label: 'Columns', multiple: true, options: cols, rows: Math.min(6, cols.length + 1), visibleIf: (v) => ['removeCols', 'keepCols', 'unpivot', 'groupBy'].includes(v.kind) },
      { key: 'type', type: 'select', label: 'New type', options: ['number', 'text', 'date', 'boolean'], value: 'number', visibleIf: (v) => v.kind === 'changeType', half: true },
      { key: 'delim', type: 'text', label: 'Delimiter', value: ',', visibleIf: (v) => v.kind === 'split', half: true },
      { key: 'fn', type: 'select', label: 'Aggregate (of Column)', options: [...AGG_FUNCS], value: 'Sum', visibleIf: (v) => v.kind === 'groupBy', half: true },
      { key: 'desc', type: 'checkbox', label: 'Descending', value: false, visibleIf: (v) => v.kind === 'sort' },
      { key: 'to', type: 'text', label: 'New name', value: '', visibleIf: (v) => v.kind === 'rename' },
      { key: 'hint', type: 'info', label: 'For Unpivot, select the columns to keep; for Group by, select the key columns.', visibleIf: (v) => v.kind === 'unpivot' || v.kind === 'groupBy' },
      { key: 'add', type: 'button', label: 'Add Step', half: true, onClick: (v, ctl) => {
        const s = stepFrom(v);
        if (!s) return;
        try { cur = applySteps(base, [...steps, s]); steps.push(s); ctl.rebuild(fieldsFor()); renderPrev(); } catch (e) { toast((e as Error).message); }
      } },
      { key: 'undo', type: 'button', label: 'Remove Last Step', half: true, onClick: (_v, ctl) => { steps.pop(); cur = applySteps(base, steps); ctl.rebuild(fieldsFor()); renderPrev(); } },
    ];
  };
  const stepFrom = (v: Record<string, any>): TransformStep | null => {
    const cols: string[] = v.cols || [];
    switch (v.kind) {
      case 'filter': return { kind: 'filter', col: v.col, op: v.op, value: String(v.value ?? '') };
      case 'removeCols': case 'keepCols': if (!cols.length) { toast('Select one or more columns.'); return null; } return { kind: v.kind, cols };
      case 'changeType': return { kind: 'changeType', col: v.col, type: v.type };
      case 'split': if (!v.delim) { toast('Enter a delimiter.'); return null; } return { kind: 'split', col: v.col, delimiter: String(v.delim) };
      case 'unpivot': return { kind: 'unpivot', keep: cols };
      case 'groupBy': if (!cols.length) { toast('Select the key columns to group by.'); return null; } return { kind: 'groupBy', keys: cols, col: v.col, fn: v.fn };
      case 'sort': return { kind: 'sort', col: v.col, desc: !!v.desc };
      case 'removeDuplicates': return { kind: 'removeDuplicates' };
      case 'removeBlankRows': return { kind: 'removeBlankRows' };
      case 'rename': if (!String(v.to).trim()) { toast('Enter a new name.'); return null; } return { kind: 'rename', col: v.col, to: String(v.to).trim() };
    }
    return null;
  };
  renderPrev();
  const res = await openDialog({ title: 'Transform Data', wide: true, ok: 'Close & Load', fields: fieldsFor(), content: h('div', { class: 'ql-data-transform' }, h('div', { class: 'ql-label' }, 'Applied steps'), stepsEl, preview) });
  if (!res) return;
  const book = app.book;
  let target!: Sheet;
  book.transact('Load Query', () => {
    target = addSheetAfter(book, t ? t.name : 'Query');
    writeMatrix(target, 0, 0, [cur.headers, ...cur.rows], { boldFirstRow: true });
  });
  activateSheet(app, target);
}

function editLinks(app: App): void {
  const found = new Set<string>();
  for (const sh of app.book.sheets) for (const cell of sh.cells.values()) {
    if (cell.f == null) continue;
    const f = cell.f.replace(/"(?:[^"]|"")*"/g, '');
    for (const m of f.matchAll(/'?\[([^\]]+)\]/g)) found.add(m[1]);
  }
  for (const n of app.book.names) for (const m of n.ref.matchAll(/'?\[([^\]]+)\]/g)) found.add(m[1]);
  if (!found.size) { void alertDialog('Edit Links', 'This workbook has no links to other workbooks.'); return; }
  void alertDialog('Edit Links', `This workbook refers to ${found.size} other workbook${found.size === 1 ? '' : 's'}:\n${[...found].join('\n')}\nLinks to other workbooks can't be updated in QuillLattice; formulas that use them keep their last values or show #REF!. Open the source workbook and copy the data in, or use Find & Replace to point the formulas at this workbook.`);
}

function relationships(app: App): void {
  const tables = app.book.sheets.flatMap((s) => s.tables.map((t) => `${t.name} — ${s.name}!${addrText(t.range)} (${t.columns.map((c) => c.name).join(', ')})`));
  void alertDialog('Relationships', `${tables.length ? `Tables in this workbook:\n${tables.join('\n')}` : 'This workbook has no tables. Use Insert ▸ Table to create one.'}\nQuillLattice doesn't keep a data model of related tables. To combine tables, look up related rows with XLOOKUP or use Get Data ▸ From Table/Range to group and reshape them.`);
}

// --------------------------------------------------------------------------------------------
// Commands

const T = 'Data';

def([
  { id: 'data.sortAZ', label: 'Sort A to Z', icon: 'arrow_upward', tab: T, group: 'Sort & Filter', popular: true, edits: true, keywords: 'ascending smallest oldest', tip: 'Sort the current region by the active column, lowest to highest.', run: (app) => quickSort(app, false) },
  { id: 'data.sortZA', label: 'Sort Z to A', icon: 'arrow_downward', tab: T, group: 'Sort & Filter', popular: true, edits: true, keywords: 'descending largest newest', tip: 'Sort the current region by the active column, highest to lowest.', run: (app) => quickSort(app, true) },
  { id: 'data.sort', label: 'Custom Sort', icon: 'sort', tab: T, group: 'Sort & Filter', edits: true, keywords: 'sort levels multiple keys custom list color', tip: 'Sort by several columns, by cell color, font color, icon or a custom list.', run: (app) => customSortDialog(app) },
  { id: 'data.sortByColor', label: 'Put Selected Cell Color on Top', icon: 'format_color_fill', tab: T, group: 'Sort & Filter', edits: true, run: (app) => sortByColor(app, 'fill') },
  { id: 'data.sortByFontColor', label: 'Put Selected Font Color on Top', icon: 'format_color_text', tab: T, group: 'Sort & Filter', edits: true, run: (app) => sortByColor(app, 'font') },
  {
    id: 'data.filter', label: 'Filter', icon: 'filter_alt', tab: T, group: 'Sort & Filter', kind: 'toggle', keys: ['Mod-Shift-L'], popular: true, edits: true, keywords: 'autofilter',
    tip: 'Turn on filter buttons for the current region or table.',
    isActive: (app) => { const s = app.sheet; const t = tableAt(s, s.sel.r, s.sel.c); return t ? !!t.opts.filterButton : !!s.filter; },
    run: (app) => { if (!blocked(app, 'autoFilter')) toggleFilter(app); },
  },
  { id: 'data.clearFilter', label: 'Clear Filter', icon: 'filter_alt_off', tab: T, group: 'Sort & Filter', edits: true, isEnabled: (app) => !!app.sheet.filter || app.sheet.tables.some((t) => t.filter) || advState.has(app.sheet), run: (app) => { if (!blocked(app, 'autoFilter')) clearFilters(app); } },
  { id: 'data.reapply', label: 'Reapply', icon: 'refresh', tab: T, group: 'Sort & Filter', keys: ['Mod-Alt-L'], edits: true, run: (app) => { if (!blocked(app, 'autoFilter')) reapply(app); } },
  { id: 'data.advancedFilter', label: 'Advanced Filter', icon: 'tune', tab: T, group: 'Sort & Filter', edits: true, keywords: 'criteria range unique records copy', run: (app) => advancedFilterDialog(app) },
  { id: 'data.filterMenu', label: 'Filter Drop-down', icon: 'arrow_drop_down', tab: T, group: 'Sort & Filter', internal: true, noRecord: true, run: (app, arg: { r: number; c: number }) => { if (arg) filterMenu(app, arg); else { const s = app.sheet; const src = filterSourceAt(s, s.sel.r, s.sel.c); if (src) filterMenu(app, { r: src.headerRow, c: s.sel.c }); } } },
  {
    id: 'data.filterBySelected', label: "Filter by Selected Cell's Value", icon: 'filter_alt', tab: T, group: 'Sort & Filter', edits: true,
    run: (app) => {
      if (blocked(app, 'autoFilter')) return;
      const s = app.sheet; const { r, c } = s.sel;
      const src = ensureFilter(app);
      if (!src) return;
      if (r === src.headerRow) { toast('Select a cell below the header row.'); return; }
      const v = s.value(r, c);
      setCriteria(app, src, c - src.area.c1, isBlank(v) ? { values: [], blanks: true } : { values: [displayText(s, r, c)], blanks: false });
    },
  },
  {
    id: 'data.filterByColor', label: "Filter by Selected Cell's Color", icon: 'format_color_fill', tab: T, group: 'Sort & Filter', edits: true,
    run: (app) => {
      if (blocked(app, 'autoFilter')) return;
      const s = app.sheet; const { r, c } = s.sel;
      const fill = s.styleOf(r, c).fill || '';
      const src = ensureFilter(app);
      if (src) setCriteria(app, src, c - src.area.c1, { color: { fill } });
    },
  },
  {
    id: 'data.sortFilterMenu', label: 'Sort & Filter', icon: 'filter_list', tab: 'Home', group: 'Editing', kind: 'menu', noRecord: true,
    run: () => undefined,
    menu: () => [
      { id: 'data.sortAZ', label: 'Sort A to Z', icon: 'arrow_upward' }, { id: 'data.sortZA', label: 'Sort Z to A', icon: 'arrow_downward' }, { id: 'data.sort', label: 'Custom Sort…', icon: 'sort' },
      { separator: true },
      { id: 'data.filter', label: 'Filter', icon: 'filter_alt' }, { id: 'data.clearFilter', label: 'Clear', icon: 'filter_alt_off' }, { id: 'data.reapply', label: 'Reapply', icon: 'refresh' },
    ],
  },
  {
    id: 'data.validationList', label: 'Validation List', icon: 'arrow_drop_down', tab: T, group: 'Data Tools', internal: true, noRecord: true,
    run: (app, arg?: { r: number; c: number }) => {
      const s = app.sheet;
      const r = arg?.r ?? s.sel.r, c = arg?.c ?? s.sel.c;
      const dv = s.validations.find((v) => v.type === 'list' && v.ranges.some((rg) => inRange(rg, r, c)));
      if (!dv) return;
      const items = listItems(app, s, dv, r, c);
      if (!items.length) { toast('The list for this cell is empty.'); return; }
      const cur = s.value(r, c);
      const pt = cellPoint(app, r, c);
      showMenu(items.slice(0, 1000).map((x) => ({ id: '__v', label: x instanceof CellError ? x.code : typeof x === 'number' ? formatValue(x, s.styleOf(r, c).numFmt).text : String(x), arg: x, checked: cur != null && compareForSort(cur, x) === 0 })), {
        x: pt.x, y: pt.y, className: 'ql-data-listmenu',
        onPick: (e) => { const v = e.arg; app.exec('cell.enter', { r, c, text: v instanceof CellError ? v.code : String(v) }); },
      });
    },
  },
  {
    id: 'data.pickFromList', label: 'Pick From Drop-down List', icon: 'list', tab: T, group: 'Data Tools', noRecord: true,
    run: (app) => {
      const s = app.sheet; const { r, c } = s.sel;
      const vals = pickListValues(s, r, c);
      if (!vals.length) { toast('There are no text entries above or below this cell to pick from.'); return; }
      const pt = cellPoint(app, r, c);
      showMenu(vals.map((v) => ({ id: '__v', label: v, arg: v })), { x: pt.x, y: pt.y, className: 'ql-data-listmenu', focus: true, onPick: (e) => app.exec('cell.enter', { r, c, text: String(e.arg) }) });
    },
  },
  // Data tools
  { id: 'data.textToColumns', label: 'Text to Columns', icon: 'view_column', tab: T, group: 'Data Tools', edits: true, keywords: 'split delimiter fixed width parse', run: (app) => textToColumns(app) },
  {
    id: 'data.flashFill', label: 'Flash Fill', icon: 'bolt', tab: T, group: 'Data Tools', keys: ['Mod-E'], edits: true, keywords: 'pattern autofill extract combine',
    tip: 'Fill the column from the examples you typed next to your data.',
    run: (app) => {
      if (blocked(app)) return;
      const s = app.sheet;
      const plan = planFlashFill(s, s.sel.r, s.sel.c);
      if (!plan) { void alertDialog('Flash Fill', 'Flash Fill couldn\'t find a pattern. Type one or two examples of the result next to your data (in the same rows), select a cell in that column and try again.\nIf the column isn\'t next to the data, Flash Fill can\'t see it.'); return; }
      applyFlash(app, plan);
    },
  },
  {
    id: 'data.flashFillSuggest', label: 'Flash Fill Suggestion', icon: 'bolt', tab: T, group: 'Data Tools', internal: true, noRecord: true,
    run: (app, arg?: { r: number; c: number }) => {
      if (!app.prefs.flashFill || app.isReadOnly()) return;
      const s = app.sheet;
      const r = arg?.r ?? s.sel.r, c = arg?.c ?? s.sel.c;
      const plan = planFlashFill(s, r, c);
      if (!plan || plan.examples < 2) return;
      toast(`Flash Fill can fill ${plan.rows.length} more cell${plan.rows.length === 1 ? '' : 's'} like your examples.`, { action: 'Fill', onAction: () => applyFlash(app, plan), ms: 6000 });
    },
  },
  { id: 'data.removeDuplicates', label: 'Remove Duplicates', icon: 'layers_clear', tab: T, group: 'Data Tools', edits: true, keywords: 'dedupe unique', run: (app) => removeDuplicatesDialog(app) },
  {
    id: 'data.validation', label: 'Data Validation', icon: 'rule', tab: T, group: 'Data Tools', kind: 'split', edits: true, keywords: 'dropdown list restrict input message error alert',
    run: (app) => validationDialog(app),
    menu: () => [{ id: 'data.validation', label: 'Data Validation…', icon: 'rule' }, { id: 'data.circleInvalid', label: 'Circle Invalid Data', icon: 'radio_button_unchecked' }, { id: 'data.clearCircles', label: 'Clear Validation Circles', icon: 'block' }],
  },
  { id: 'data.circleInvalid', label: 'Circle Invalid Data', icon: 'radio_button_unchecked', tab: T, group: 'Data Tools', noRecord: true, run: (app) => circleInvalid(app) },
  { id: 'data.clearCircles', label: 'Clear Validation Circles', icon: 'block', tab: T, group: 'Data Tools', noRecord: true, run: (app) => { app.tab.circles = []; try { app.grid.invalidate(); } catch { /* */ } } },
  { id: 'data.consolidate', label: 'Consolidate', icon: 'merge_type', tab: T, group: 'Data Tools', edits: true, keywords: 'combine ranges sum sheets', run: (app) => consolidateDialog(app) },
  { id: 'data.relationships', label: 'Relationships', icon: 'schema', tab: T, group: 'Data Tools', noRecord: true, run: (app) => relationships(app) },
  {
    id: 'data.dataTypes', label: 'Data Types', icon: 'category', tab: T, group: 'Data Types', kind: 'menu', noRecord: true, run: () => undefined,
    menu: () => [
      { id: 'data.convertToNumber', label: 'Convert Text to Numbers', icon: 'pin' }, { id: 'data.convertToDate', label: 'Convert Text to Dates', icon: 'calendar_month' },
      { id: 'data.convertToText', label: 'Convert to Text', icon: 'text_fields' }, { separator: true }, { id: 'data.textToColumns', label: 'Text to Columns…', icon: 'view_column' },
    ],
  },
  { id: 'data.convertToNumber', label: 'Convert Text to Numbers', icon: 'pin', tab: T, group: 'Data Types', edits: true, keywords: 'number stored as text', run: (app) => convertSelection(app, 'number') },
  { id: 'data.convertToDate', label: 'Convert Text to Dates', icon: 'calendar_month', tab: T, group: 'Data Types', edits: true, run: (app) => convertSelection(app, 'date') },
  { id: 'data.convertToText', label: 'Convert to Text', icon: 'text_fields', tab: T, group: 'Data Types', edits: true, run: (app) => convertSelection(app, 'text') },
  // Outline
  {
    id: 'data.group', label: 'Group', icon: 'unfold_less', tab: T, group: 'Outline', kind: 'split', keys: ['Alt-Shift-ArrowRight'], edits: true, keywords: 'outline',
    run: (app, arg) => groupCmd(app, arg, 1),
    menu: () => [{ id: 'data.group', label: 'Group…', icon: 'unfold_less' }, { id: 'data.autoOutline', label: 'Auto Outline', icon: 'account_tree' }],
  },
  {
    id: 'data.ungroup', label: 'Ungroup', icon: 'unfold_more', tab: T, group: 'Outline', kind: 'split', keys: ['Alt-Shift-ArrowLeft'], edits: true,
    run: (app, arg) => groupCmd(app, arg, -1),
    menu: () => [{ id: 'data.ungroup', label: 'Ungroup…', icon: 'unfold_more' }, { id: 'data.clearOutline', label: 'Clear Outline', icon: 'clear_all' }],
  },
  { id: 'data.subtotal', label: 'Subtotal', icon: 'functions', tab: T, group: 'Outline', edits: true, keywords: 'group totals outline', run: (app) => subtotalDialog(app) },
  { id: 'data.showDetail', label: 'Show Detail', icon: 'add_box', tab: T, group: 'Outline', edits: true, run: (app, arg) => detailCmd(app, true, arg) },
  { id: 'data.hideDetail', label: 'Hide Detail', icon: 'indeterminate_check_box', tab: T, group: 'Outline', edits: true, run: (app, arg) => detailCmd(app, false, arg) },
  {
    id: 'data.autoOutline', label: 'Auto Outline', icon: 'account_tree', tab: T, group: 'Outline', edits: true,
    run: (app) => {
      const s = app.sheet;
      const u = s.usedRange(false);
      if (!u) { toast('There is no data to outline.'); return; }
      let res = { rows: 0, cols: 0 };
      app.book.transact('Auto Outline', () => { s.book.recordSheet(s); res = autoOutline(s, u); });
      if (!res.rows && !res.cols) toast('Cannot create an outline: no summary formulas (SUM or SUBTOTAL of the cells right above or to the left) were found.');
    },
  },
  {
    id: 'data.clearOutline', label: 'Clear Outline', icon: 'clear_all', tab: T, group: 'Outline', edits: true,
    run: (app) => {
      const s = app.sheet;
      if (!s.rowLevel.size && !s.colLevel.size) { toast('This sheet has no outline.'); return; }
      app.book.transact('Clear Outline', () => {
        s.book.recordSheet(s);
        for (const r of s.rowLevel.keys()) s.hiddenRows.delete(r);
        for (const c of s.colLevel.keys()) s.hiddenCols.delete(c);
        s.rowLevel.clear(); s.colLevel.clear();
      });
    },
  },
  {
    id: 'data.outlineLevel', label: 'Show Outline Level', icon: 'format_list_numbered', tab: T, group: 'Outline', kind: 'menu', edits: true,
    menu: (app) => outlineLevelMenu(app),
    run: (app, arg?: { axis?: OutlineAxis; level?: number }) => {
      if (!arg?.level) return;
      const s = app.sheet;
      app.book.transact('Outline Level', () => { s.book.recordSheet(s); showOutlineLevel(s, arg.axis || 'row', arg.level!); });
    },
  },
  { id: 'data.outlineSettings', label: 'Outline Settings', icon: 'settings', tab: T, group: 'Outline', noRecord: true, run: (app) => outlineSettings(app) },
  // Get & Transform
  {
    id: 'data.getData', label: 'Get Data', icon: 'database', tab: T, group: 'Get & Transform Data', kind: 'menu', noRecord: true, run: () => undefined,
    menu: () => [
      { id: 'data.fromText', label: 'From Text/CSV…', icon: 'description' }, { id: 'data.fromWorkbook', label: 'From Workbook…', icon: 'grid_on' },
      { id: 'data.fromWeb', label: 'From Web…', icon: 'language' }, { id: 'data.fromTable', label: 'From Table/Range…', icon: 'table_chart' },
      { id: 'data.fromClipboard', label: 'From Clipboard', icon: 'content_paste' },
      { separator: true }, { id: 'data.connections', label: 'Queries & Connections', icon: 'hub' }, { id: 'data.refreshAll', label: 'Refresh All', icon: 'sync' },
    ],
  },
  { id: 'data.fromText', label: 'From Text/CSV', icon: 'description', tab: T, group: 'Get & Transform Data', edits: true, keywords: 'import csv text wizard', run: (app) => fromText(app) },
  { id: 'data.fromWorkbook', label: 'From Workbook', icon: 'grid_on', tab: T, group: 'Get & Transform Data', edits: true, keywords: 'import sheet another file', run: (app) => fromWorkbook(app) },
  { id: 'data.fromWeb', label: 'From Web', icon: 'language', tab: T, group: 'Get & Transform Data', edits: true, keywords: 'import html table url', run: (app) => fromWeb(app) },
  { id: 'data.fromTable', label: 'From Table/Range', icon: 'table_chart', tab: T, group: 'Get & Transform Data', edits: true, keywords: 'transform query unpivot group by', run: (app) => fromTable(app) },
  { id: 'data.fromClipboard', label: 'From Clipboard', icon: 'content_paste', tab: T, group: 'Get & Transform Data', edits: true, run: (app) => fromClipboard(app) },
  { id: 'data.refreshAll', label: 'Refresh All', icon: 'sync', tab: T, group: 'Queries & Connections', keys: ['Mod-Alt-F5'], run: (app) => refreshAll(app) },
  { id: 'data.connections', label: 'Queries & Connections', icon: 'hub', tab: T, group: 'Queries & Connections', noRecord: true, run: (app) => connectionsDialog(app) },
  { id: 'data.editLinks', label: 'Edit Links', icon: 'link', tab: T, group: 'Queries & Connections', noRecord: true, keywords: 'external references other workbooks', run: (app) => editLinks(app) },
]);

export { filterSourceAt, setCriteria, tableDataOf };
export type { CellStyle };
