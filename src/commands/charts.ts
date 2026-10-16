/**
 * Charts: creating charts from the selection (Insert ▸ Charts), the chart dialogs, and the Chart Design
 * contextual tab (elements, layouts, colors, styles, switch row/column, select data, change type,
 * move, save as picture).
 */
import { def, type MenuEntry } from './registry';
import type { App } from '../app';
import type { ChartDef, ChartSeries, ChartType } from '../model/types';
import { Sheet, newId } from '../model/workbook';
import { type Range, quoteSheet, splitSheetRef, parseRange, colName } from '../model/address';
import { rangeRefText, resolveRangeText, validSheetName } from '../model/structure';
import { isDateFormat } from '../model/numfmt';
import { activeRange, currentRegion, clipToUsed } from './ops';
import { openDialog, toast, setListOptions, promptDialog } from '../ui/dialog';
import { h } from '../ui/dom';
import { renderChartSvg, CHART_TYPE_LABELS, CHART_TYPES, DEFAULT_CHART_COLORS, mixColor } from '../charts/render';
import { addChart, findObject, galleryDialog, mutateObject, canEditObjects, type GalleryItem } from './objects';

// --------------------------------------------------------------------------------------------
// Building a chart from a range

export interface Layout { headerRow: boolean; catCols: number; catCol: boolean; seriesInRows: boolean }

const isText = (sheet: Sheet, r: number, c: number) => { const v = sheet.value(r, c); return typeof v === 'string' && v !== ''; };
const isNum = (sheet: Sheet, r: number, c: number) => typeof sheet.value(r, c) === 'number';
const isDate = (sheet: Sheet, r: number, c: number) => isNum(sheet, r, c) && isDateFormat(sheet.styleOf(r, c).numFmt);
const isEmpty = (sheet: Sheet, r: number, c: number) => { const v = sheet.value(r, c); return v == null || v === ''; };

/** Detects header row, category columns and series orientation the way spreadsheets do. */
export function detectLayout(sheet: Sheet, rg: Range, type?: ChartType): Layout {
  const rows = rg.r2 - rg.r1 + 1, cols = rg.c2 - rg.c1 + 1;
  const topLeftEmpty = isEmpty(sheet, rg.r1, rg.c1);
  let headerRow = false;
  if (rows > 1) {
    let text = 0, nums = 0, below = 0;
    for (let c = rg.c1; c <= rg.c2; c++) {
      if (isText(sheet, rg.r1, c)) text++;
      else if (isNum(sheet, rg.r1, c) && !isDate(sheet, rg.r1, c)) nums++;
      if (isNum(sheet, rg.r1 + 1, c)) below++;
    }
    headerRow = (text > 0 && nums === 0 && below > 0) || (topLeftEmpty && cols > 1 && [...Array(cols - 1)].every((_, i) => !isEmpty(sheet, rg.r1, rg.c1 + 1 + i)));
    if (!headerRow && text === cols && rows > 1) headerRow = true;
  }
  const d0 = rg.r1 + (headerRow ? 1 : 0);
  let catCols = 0;
  const xy = type === 'scatter' || type === 'scatterLines' || type === 'bubble';
  for (let c = rg.c1; c < rg.c2; c++) {
    let textish = true, any = false;
    for (let r = d0; r <= rg.r2; r++) {
      if (isEmpty(sheet, r, c)) continue;
      any = true;
      if (isNum(sheet, r, c) && !isDate(sheet, r, c)) { textish = false; break; }
      if (xy && isDate(sheet, r, c)) { textish = false; break; }
    }
    if (any && textish) catCols++;
    else break;
  }
  if (!catCols && topLeftEmpty && headerRow && cols > 1 && !xy) catCols = 1;
  const dataRows = rg.r2 - d0 + 1, dataCols = cols - catCols;
  const seriesInRows = !xy && dataRows < dataCols && dataRows > 0;
  return { headerRow, catCols, catCol: catCols > 0, seriesInRows };
}

const SINGLE_SERIES = new Set<ChartType>(['pie', 'doughnut', 'funnel', 'waterfall', 'pareto', 'treemap', 'sunburst']);

/** Builds a ChartDef (series formulas, categories, names) from a range. */
export function chartFromRange(sheet: Sheet, rg: Range, type: ChartType, opts: { seriesInRows?: boolean; secondaryLast?: boolean; comboTypes?: ChartType[] } = {}): ChartDef {
  const lay = detectLayout(sheet, rg, type);
  const inRows = opts.seriesInRows ?? lay.seriesInRows;
  const ref = (r1: number, c1: number, r2: number, c2: number) => rangeRefText(sheet, { r1, c1, r2, c2 }, true);
  const d0 = rg.r1 + (lay.headerRow ? 1 : 0);
  const c0 = rg.c1 + lay.catCols;
  const series: ChartSeries[] = [];
  const xy = type === 'scatter' || type === 'scatterLines' || type === 'bubble';
  if (xy) {
    // First numeric column is X (when there are at least two), the rest are Y (bubble: Y, size pairs).
    const cols: number[] = [];
    for (let c = rg.c1; c <= rg.c2; c++) cols.push(c);
    const hasX = cols.length > 1;
    const xCol = hasX ? cols[0] : -1;
    const ys = hasX ? cols.slice(1) : cols;
    const step = type === 'bubble' ? 2 : 1;
    for (let i = 0; i < ys.length; i += step) {
      const c = ys[i];
      const s: ChartSeries = { values: ref(d0, c, rg.r2, c) };
      if (hasX) s.categories = ref(d0, xCol, rg.r2, xCol);
      if (lay.headerRow) s.nameRef = ref(rg.r1, c, rg.r1, c);
      if (type === 'bubble' && ys[i + 1] != null) s.sizes = ref(d0, ys[i + 1], rg.r2, ys[i + 1]);
      series.push(s);
    }
  } else if (!inRows) {
    const cats = lay.catCols ? ref(d0, rg.c1, rg.r2, c0 - 1) : undefined;
    for (let c = c0; c <= rg.c2; c++) {
      const s: ChartSeries = { values: ref(d0, c, rg.r2, c) };
      if (cats) s.categories = cats;
      if (lay.headerRow) s.nameRef = ref(rg.r1, c, rg.r1, c);
      series.push(s);
    }
  } else {
    const cats = lay.headerRow ? ref(rg.r1, c0, rg.r1, rg.c2) : undefined;
    for (let r = d0; r <= rg.r2; r++) {
      const s: ChartSeries = { values: ref(r, c0, r, rg.c2) };
      if (cats) s.categories = cats;
      if (lay.catCols) s.nameRef = ref(r, c0 - 1, r, c0 - 1);
      series.push(s);
    }
  }
  const list = SINGLE_SERIES.has(type) ? series.slice(0, 1) : series;
  if (type === 'combo') list.forEach((s, i) => { s.type = opts.comboTypes?.[i] || (i === 0 || (list.length > 2 && i < list.length - 1) ? 'column' : 'line'); if (opts.secondaryLast && i === list.length - 1 && i > 0) s.secondary = true; });
  return {
    id: newId('ch'), type, series: list,
    at: { r: rg.r1, c: rg.c2 + 2, dx: 0, dy: 0, w: 480, h: 288 },
    sourceRange: rangeRefText(sheet, rg, true), seriesInRows: inRows,
    gridlines: !['pie', 'doughnut', 'treemap', 'sunburst', 'funnel'].includes(type),
  };
}

/** The range a chart should be made from: the selection, or the current region around one cell. */
export function chartSourceRange(app: App): Range | null {
  const s = app.sheet;
  let rg = activeRange(app);
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) rg = currentRegion(s, rg.r1, rg.c1);
  else rg = clipToUsed(s, rg);
  let any = false;
  s.forEachIn(rg, (_r, _c, cell) => { if (cell.v != null && cell.v !== '') any = true; });
  return any ? rg : null;
}

/** Recommends up to six chart types for the data. */
export function recommendTypes(sheet: Sheet, rg: Range | null): ChartType[] {
  const out: ChartType[] = [];
  const add = (...t: ChartType[]) => { for (const x of t) if (!out.includes(x)) out.push(x); };
  if (!rg) { add('column', 'line', 'pie', 'bar', 'area', 'scatter'); return out; }
  const lay = detectLayout(sheet, rg);
  const d0 = rg.r1 + (lay.headerRow ? 1 : 0);
  const dataRows = rg.r2 - d0 + 1;
  const numCols = rg.c2 - rg.c1 + 1 - lay.catCols;
  const nSeries = lay.seriesInRows ? dataRows : numCols;
  const nPoints = lay.seriesInRows ? numCols : dataRows;
  const dateCats = lay.catCols > 0 && isDate(sheet, d0, rg.c1);
  let allPositive = true;
  sheet.forEachIn({ r1: d0, c1: rg.c1 + lay.catCols, r2: rg.r2, c2: rg.c2 }, (_r, _c, cell) => { if (typeof cell.v === 'number' && cell.v < 0) allPositive = false; });
  if (lay.catCols >= 2) add('treemap', 'sunburst', 'column', 'bar');
  if (!lay.catCols && numCols >= 2 && !lay.seriesInRows) add('scatter', 'line', 'column', numCols >= 3 ? 'bubble' : 'scatterLines');
  if (!lay.catCols && numCols === 1 && nPoints >= 10) add('histogram', 'boxWhisker', 'line', 'column');
  if (dateCats) add('line', 'column', 'area', 'lineMarkers');
  if (nSeries === 1 && nPoints <= 8 && allPositive) add('column', 'pie', 'bar', 'doughnut', 'treemap');
  if (nSeries === 1 && !allPositive) add('column', 'waterfall', 'bar', 'line');
  if (nSeries > 1) add('column', 'line', 'bar', 'stackedColumn', 'area', nSeries >= 3 && nPoints >= 3 ? 'radar' : 'combo');
  add('column', 'line', 'bar', 'pie', 'area', 'scatter');
  return out.slice(0, 6);
}

/** Inserts a chart of the given type from the selection and selects it. */
export function insertChart(app: App, type: ChartType, opts: { secondaryLast?: boolean; comboTypes?: ChartType[]; seriesInRows?: boolean } = {}): ChartDef | null {
  if (!canEditObjects(app)) return null;
  const s = app.sheet;
  const rg = chartSourceRange(app);
  const ch: ChartDef = rg ? chartFromRange(s, rg, type, opts) : { id: newId('ch'), type, series: [], at: { r: s.sel.r, c: s.sel.c + 1, dx: 0, dy: 0, w: 480, h: 288 }, gridlines: true };
  // Place it beside the data when that is on screen, else in the middle of the view.
  const g = app.tab?.grid;
  if (g?.width) {
    const v = g.visible();
    const right = rg ? rg.c2 + 1 : s.sel.c + 1;
    if (!(right >= v.c1 && right + 3 <= v.c2 && (rg ? rg.r1 : s.sel.r) >= v.r1 && (rg ? rg.r1 : s.sel.r) <= v.r2)) {
      const z = g.zoom;
      Object.assign(ch.at, g.objects.anchorFor(Math.max(g.headerW + g.frozenW + 8, (g.width - ch.at.w * z) / 2), Math.max(g.headerH + g.frozenH + 8, (g.height - ch.at.h * z) / 2)));
    } else ch.at = { ...ch.at, r: rg ? rg.r1 : s.sel.r, c: right, dx: 12, dy: 0 };
  }
  return addChart(app, ch, `Insert ${CHART_TYPE_LABELS[type]} Chart`);
}

// --------------------------------------------------------------------------------------------
// Previews

/** SVG preview of a chart type for the selected data (or sample data when nothing is selected). */
export function chartPreview(app: App, type: ChartType, w = 180, h = 112, base?: ChartDef): string {
  const s = app.sheet;
  let ch: ChartDef;
  if (base) ch = { ...base, type, series: base.series.map((x) => ({ ...x, type: undefined })), title: '', legend: 'none', dataLabels: false, xTitle: undefined, yTitle: undefined };
  else {
    const rg = chartSourceRange(app);
    ch = rg ? chartFromRange(s, rg, type) : sampleChart(type);
    ch.title = ''; ch.legend = 'none';
  }
  if (type === 'combo' && ch.series.length > 1) ch.series = ch.series.map((x, i) => ({ ...x, type: i === ch.series.length - 1 ? 'line' : 'column' }));
  try { return renderChartSvg(app.tab, ch, w, h, s); } catch { return ''; }
}

function sampleChart(type: ChartType): ChartDef {
  return { id: 'sample', type, at: { r: 0, c: 0, dx: 0, dy: 0, w: 180, h: 112 }, series: [{ name: 'A', values: '{4,7,3,8,5}' }, { name: 'B', values: '{2,5,6,3,7}' }, { name: 'C', values: '{3,2,4,6,4}' }].slice(0, SINGLE_SERIES.has(type) ? 1 : 3) };
}

export const CHART_ICONS: Record<ChartType, string> = {
  column: 'bar_chart', stackedColumn: 'stacked_bar_chart', percentColumn: 'stacked_bar_chart', bar: 'align_horizontal_left', stackedBar: 'align_horizontal_left', percentBar: 'align_horizontal_left',
  line: 'show_chart', lineMarkers: 'timeline', stackedLine: 'stacked_line_chart', area: 'area_chart', stackedArea: 'area_chart', percentArea: 'area_chart',
  pie: 'pie_chart', doughnut: 'donut_large', scatter: 'scatter_plot', scatterLines: 'ssid_chart', bubble: 'bubble_chart', radar: 'radar', filledRadar: 'radar',
  histogram: 'equalizer', pareto: 'leaderboard', boxWhisker: 'candlestick_chart', waterfall: 'waterfall_chart', funnel: 'filter_alt', treemap: 'grid_view', sunburst: 'data_usage',
  stock: 'candlestick_chart', combo: 'insert_chart',
};

const CHART_GROUPS: [string, ChartType[]][] = [
  ['Column', ['column', 'stackedColumn', 'percentColumn']], ['Line', ['line', 'lineMarkers', 'stackedLine']], ['Pie', ['pie', 'doughnut']],
  ['Bar', ['bar', 'stackedBar', 'percentBar']], ['Area', ['area', 'stackedArea', 'percentArea']], ['X Y (Scatter)', ['scatter', 'scatterLines', 'bubble']],
  ['Stock', ['stock']], ['Radar', ['radar', 'filledRadar']], ['Treemap', ['treemap']], ['Sunburst', ['sunburst']], ['Histogram', ['histogram', 'pareto']],
  ['Box & Whisker', ['boxWhisker']], ['Waterfall', ['waterfall']], ['Funnel', ['funnel']], ['Combo', ['combo']],
];

function typeItems(app: App, types: ChartType[], base?: ChartDef, group = false): GalleryItem<ChartType>[] {
  const groupOf = (t: ChartType) => CHART_GROUPS.find(([, ts]) => ts.includes(t))?.[0];
  return types.map((t) => ({ value: t, label: CHART_TYPE_LABELS[t], html: chartPreview(app, t, 180, 112, base), group: group ? groupOf(t) : undefined }));
}

// --------------------------------------------------------------------------------------------
// Chart editing helpers

function chartOf(app: App, arg?: any): { def: ChartDef; sheet: Sheet } | null {
  const o = findObject(app, arg && typeof arg === 'object' && typeof arg.id === 'string' ? arg.id : undefined);
  return o?.kind === 'chart' ? { def: o.def, sheet: o.sheet } : null;
}

function setChart(app: App, arg: any, label: string, fn: (c: ChartDef) => void): boolean {
  const o = findObject(app, arg && typeof arg === 'object' && typeof arg.id === 'string' ? arg.id : undefined);
  if (!o || o.kind !== 'chart') { toast('Select a chart first.'); return false; }
  mutateObject(app, o, label, fn);
  return true;
}

/** Bounding range of a chart's series references (used when the chart has no stored source range). */
export function inferSourceRange(app: App, ch: ChartDef, home: Sheet): { sheet: Sheet; range: Range } | null {
  if (ch.sourceRange) { const r = resolveRangeText(app.book, ch.sourceRange.replace(/\$/g, ''), home); if (r) return r; }
  let sheet: Sheet | null = null;
  let box: Range | null = null;
  for (const s of ch.series) for (const ref of [s.values, s.categories, s.nameRef]) {
    if (!ref) continue;
    const r = resolveRangeText(app.book, ref.replace(/\$/g, ''), home);
    if (!r) continue;
    if (sheet && r.sheet !== sheet) return null;
    sheet = r.sheet;
    box = box ? { r1: Math.min(box.r1, r.range.r1), c1: Math.min(box.c1, r.range.c1), r2: Math.max(box.r2, r.range.r2), c2: Math.max(box.c2, r.range.c2) } : { ...r.range };
  }
  return sheet && box ? { sheet, range: box } : null;
}

const PALETTES: { id: string; label: string; colors: string[] | null }[] = [
  { id: 'theme', label: 'Colorful (Theme)', colors: null },
  { id: 'classic', label: 'Classic Office-Style', colors: ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'] },
  { id: 'bright', label: 'Bright', colors: ['#e6194b', '#3cb44b', '#ffe119', '#4363d8', '#f58231', '#911eb4'] },
  { id: 'pastel', label: 'Pastel', colors: ['#8ecae6', '#ffb4a2', '#b5e48c', '#cdb4db', '#ffd166', '#a2d2ff'] },
  { id: 'ocean', label: 'Ocean', colors: ['#03045e', '#0077b6', '#00b4d8', '#48cae4', '#90e0ef', '#caf0f8'] },
  { id: 'earth', label: 'Earth', colors: ['#7f5539', '#b08968', '#ddb892', '#606c38', '#283618', '#bc6c25'] },
  { id: 'grayscale', label: 'Grayscale', colors: ['#262626', '#595959', '#7f7f7f', '#a6a6a6', '#bfbfbf', '#d9d9d9'] },
  { id: 'colorblind', label: 'Color-Blind Safe', colors: ['#0072b2', '#e69f00', '#009e73', '#cc79a7', '#56b4e9', '#d55e00'] },
  ...DEFAULT_CHART_COLORS.map((c, i) => ({ id: 'mono' + i, label: `Monochromatic: ${['Blue', 'Orange', 'Gray', 'Gold', 'Light Blue', 'Green'][i]}`, colors: [c, mixColor(c, '#000000', 0.3), mixColor(c, '#ffffff', 0.3), mixColor(c, '#000000', 0.5), mixColor(c, '#ffffff', 0.55), mixColor(c, '#000000', 0.15)] })),
];

const CHART_STYLES: [number, string][] = [[1, 'Style 1: Flat'], [2, 'Style 2: Rounded Bars'], [3, 'Style 3: White Outlines'], [4, 'Style 4: Gradient'], [5, 'Style 5: Bold Lines'], [6, 'Style 6: Dark Outlines']];

const QUICK_LAYOUTS: { label: string; patch: Partial<ChartDef> }[] = [
  { label: 'Layout 1: Title and legend at right', patch: { legend: 'right', dataLabels: false, gridlines: true, xTitle: undefined, yTitle: undefined } },
  { label: 'Layout 2: Legend at top, data labels, no gridlines', patch: { legend: 'top', dataLabels: true, gridlines: false } },
  { label: 'Layout 3: Legend at bottom', patch: { legend: 'bottom', dataLabels: false, gridlines: true, xTitle: undefined, yTitle: undefined } },
  { label: 'Layout 4: Data labels, legend at bottom', patch: { legend: 'bottom', dataLabels: true, gridlines: true } },
  { label: 'Layout 5: Axis titles, no legend', patch: { legend: 'none', dataLabels: false, gridlines: true, xTitle: 'Axis Title', yTitle: 'Axis Title' } },
  { label: 'Layout 6: Vertical axis title, legend at right', patch: { legend: 'right', dataLabels: false, gridlines: true, yTitle: 'Axis Title', xTitle: undefined } },
  { label: 'Layout 7: Both axis titles, legend at right', patch: { legend: 'right', gridlines: true, xTitle: 'Axis Title', yTitle: 'Axis Title' } },
  { label: 'Layout 8: Minimal (no legend, no gridlines)', patch: { legend: 'none', gridlines: false, dataLabels: false, xTitle: undefined, yTitle: undefined } },
];

// --------------------------------------------------------------------------------------------
// Series editor (shared by Edit Chart and Select Data)

async function editSeries(app: App, s: ChartSeries, isCombo: boolean): Promise<ChartSeries | null> {
  const r = await openDialog({
    title: 'Edit Series',
    fields: [
      { key: 'name', label: 'Series name (text, or a cell reference such as =Sheet1!$B$1)', type: 'text', value: s.nameRef ? '=' + s.nameRef : s.name || '' },
      { key: 'values', label: 'Series values', type: 'text', value: s.values ? '=' + s.values : '', placeholder: "=Sheet1!$B$2:$B$10" },
      { key: 'categories', label: 'Category (X) axis labels', type: 'text', value: s.categories ? '=' + s.categories : '' },
      { key: 'sizes', label: 'Bubble sizes (bubble charts)', type: 'text', value: s.sizes ? '=' + s.sizes : '' },
      ...(isCombo ? [{ key: 'type', label: 'Chart type for this series', type: 'select' as const, options: ['column', 'stackedColumn', 'line', 'lineMarkers', 'area', 'stackedArea'].map((t) => ({ value: t, label: CHART_TYPE_LABELS[t as ChartType] })), value: s.type || 'column' }] : []),
      { key: 'secondary', label: 'Plot on secondary axis', type: 'checkbox', value: !!s.secondary },
      { key: 'useColor', label: 'Custom color', type: 'checkbox', value: !!s.color, half: true },
      { key: 'color', label: 'Color', type: 'color', value: s.color || '#2f6fb0', half: true, visibleIf: (v) => !!v.useColor },
    ],
    validate: (v) => (!String(v.values || '').replace(/^=/, '').trim() ? 'Enter the series values.' : null),
  });
  if (!r) return null;
  const strip = (x: string) => String(x || '').trim().replace(/^=/, '');
  const out: ChartSeries = { values: strip(r.values) };
  const nm = String(r.name || '').trim();
  if (nm.startsWith('=')) out.nameRef = strip(nm); else if (nm) out.name = nm;
  if (strip(r.categories)) out.categories = strip(r.categories);
  if (strip(r.sizes)) out.sizes = strip(r.sizes);
  if (isCombo && r.type) out.type = r.type;
  if (r.secondary) out.secondary = true;
  if (r.useColor) out.color = r.color;
  return out;
}

function seriesLabel(s: ChartSeries, i: number): string {
  return `${s.name || (s.nameRef ? '=' + s.nameRef : `Series${i + 1}`)}  —  ${s.values}`;
}

// --------------------------------------------------------------------------------------------
// Full chart editor

async function chartEditor(app: App, arg: any): Promise<void> {
  const found = chartOf(app, arg);
  if (!found) { toast('Select a chart first.'); return; }
  const { def: ch, sheet } = found;
  const work: ChartDef = JSON.parse(JSON.stringify(ch));
  const preview = (v: Record<string, any>) => {
    apply(work, v);
    return renderChartSvg(app.tab, work, 420, 250, sheet);
  };
  const apply = (c: ChartDef, v: Record<string, any>) => {
    c.type = v.type;
    c.title = v.autoTitle ? undefined : String(v.title ?? '');
    c.xTitle = String(v.xTitle || '') || undefined;
    c.yTitle = String(v.yTitle || '') || undefined;
    c.legend = v.legend;
    c.dataLabels = !!v.dataLabels;
    c.gridlines = !!v.gridlines;
    c.style = Number(v.style) || 1;
    const pal = PALETTES.find((p) => p.id === v.palette);
    c.colors = pal?.colors ? [...pal.colors] : undefined;
    c.trendline = v.trendline ? v.trendline : null;
    c.yMin = v.yMin === '' || v.yMin == null || isNaN(Number(v.yMin)) ? undefined : Number(v.yMin);
    c.yMax = v.yMax === '' || v.yMax == null || isNaN(Number(v.yMax)) ? undefined : Number(v.yMax);
    c.logScale = !!v.logScale;
  };
  const palId = ch.colors ? PALETTES.find((p) => JSON.stringify(p.colors) === JSON.stringify(ch.colors))?.id || 'theme' : 'theme';
  const refreshSeries = (ctl: any) => setListOptions(ctl, 'series', work.series.map((s, i) => ({ value: String(i), label: seriesLabel(s, i) })));
  const r = await openDialog({
    title: 'Edit Chart', wide: true, preview: true, previewStyle: 'display:flex;justify-content:center;background:#fff;overflow:auto',
    fields: [
      { key: 'type', label: 'Chart type', type: 'select', options: CHART_TYPES.map((t) => ({ value: t, label: CHART_TYPE_LABELS[t] })), value: ch.type, half: true },
      { key: 'style', label: 'Style', type: 'select', options: CHART_STYLES.map(([n, l]) => ({ value: String(n), label: l })), value: String(ch.style || 1), half: true },
      { key: 'autoTitle', label: 'Automatic title (series name)', type: 'checkbox', value: ch.title === undefined, half: true },
      { key: 'title', label: 'Chart title', type: 'text', value: ch.title ?? '', half: true, visibleIf: (v) => !v.autoTitle },
      { key: 'xTitle', label: 'Horizontal axis title', type: 'text', value: ch.xTitle || '', half: true },
      { key: 'yTitle', label: 'Vertical axis title', type: 'text', value: ch.yTitle || '', half: true },
      { key: 'legend', label: 'Legend', type: 'select', options: [{ value: 'right', label: 'Right' }, { value: 'top', label: 'Top' }, { value: 'left', label: 'Left' }, { value: 'bottom', label: 'Bottom' }, { value: 'none', label: 'None' }], value: ch.legend ?? (ch.series.length > 1 ? 'bottom' : 'none'), half: true },
      { key: 'palette', label: 'Colors', type: 'select', options: PALETTES.map((p) => ({ value: p.id, label: p.label })), value: palId, half: true },
      { key: 'dataLabels', label: 'Data labels', type: 'checkbox', value: !!ch.dataLabels, half: true },
      { key: 'gridlines', label: 'Gridlines', type: 'checkbox', value: ch.gridlines !== false, half: true },
      { key: 'trendline', label: 'Trendline', type: 'select', options: [{ value: '', label: 'None' }, { value: 'linear', label: 'Linear' }, { value: 'exponential', label: 'Exponential' }, { value: 'polynomial', label: 'Polynomial (order 2)' }, { value: 'movingAverage', label: 'Moving Average (2 periods)' }], value: ch.trendline || '', half: true },
      { key: 'logScale', label: 'Logarithmic value axis', type: 'checkbox', value: !!ch.logScale, half: true },
      { key: 'yMin', label: 'Value axis minimum (blank = auto)', type: 'text', value: ch.yMin ?? '', half: true },
      { key: 'yMax', label: 'Value axis maximum (blank = auto)', type: 'text', value: ch.yMax ?? '', half: true },
      { key: 'h1', label: 'Series', type: 'heading' },
      { key: 'series', label: '', type: 'list', rows: 4, options: work.series.map((s, i) => ({ value: String(i), label: seriesLabel(s, i) })) },
      {
        key: 'add', label: 'Add Series…', type: 'button', half: true, onClick: async (_v, ctl) => {
          const s = await editSeries(app, { values: '' }, work.type === 'combo');
          if (s) { work.series.push(s); refreshSeries(ctl); ctl.setPreview(preview(ctl.values())); }
        },
      },
      {
        key: 'edit', label: 'Edit Series…', type: 'button', half: true, onClick: async (v, ctl) => {
          const i = Number(v.series);
          if (!(i >= 0) || !work.series[i]) { toast('Choose a series in the list.'); return; }
          const s = await editSeries(app, work.series[i], work.type === 'combo');
          if (s) { work.series[i] = s; refreshSeries(ctl); ctl.setPreview(preview(ctl.values())); }
        },
      },
      {
        key: 'remove', label: 'Remove Series', type: 'button', half: true, onClick: (v, ctl) => {
          const i = Number(v.series);
          if (!(i >= 0) || !work.series[i]) { toast('Choose a series in the list.'); return; }
          work.series.splice(i, 1); refreshSeries(ctl); ctl.setPreview(preview(ctl.values()));
        },
      },
      {
        key: 'up', label: 'Move Up', type: 'button', half: true, onClick: (v, ctl) => {
          const i = Number(v.series);
          if (!(i > 0)) return;
          [work.series[i - 1], work.series[i]] = [work.series[i], work.series[i - 1]]; refreshSeries(ctl); ctl.set('series', String(i - 1)); ctl.setPreview(preview(ctl.values()));
        },
      },
    ],
    onChange: (v, ctl) => ctl.setPreview(preview(v)),
  });
  if (!r) return;
  setChart(app, { id: ch.id }, 'Edit Chart', (c) => { apply(work, r); Object.assign(c, { ...work, id: c.id, at: c.at }); for (const k of Object.keys(c) as (keyof ChartDef)[]) if ((c as any)[k] === undefined) delete (c as any)[k]; });
}

async function selectData(app: App, arg: any): Promise<void> {
  const found = chartOf(app, arg);
  if (!found) { toast('Select a chart first.'); return; }
  const { def: ch, sheet } = found;
  const src = inferSourceRange(app, ch, sheet);
  const work: ChartSeries[] = JSON.parse(JSON.stringify(ch.series));
  let rangeChanged = false;
  const refresh = (ctl: any) => setListOptions(ctl, 'series', work.map((s, i) => ({ value: String(i), label: seriesLabel(s, i) })));
  const r = await openDialog({
    title: 'Select Data Source', wide: true,
    fields: [
      { key: 'range', label: 'Chart data range', type: 'text', value: src ? '=' + rangeRefText(src.sheet, src.range, true) : '', placeholder: "=Sheet1!$A$1:$D$10", help: 'Changing the range rebuilds every series from it.' },
      { key: 'rows', label: 'Series in rows (switch row/column)', type: 'checkbox', value: !!ch.seriesInRows },
      { key: 'h', label: 'Legend Entries (Series)', type: 'heading' },
      { key: 'series', type: 'list', rows: 6, options: work.map((s, i) => ({ value: String(i), label: seriesLabel(s, i) })) },
      { key: 'add', label: 'Add…', type: 'button', half: true, onClick: async (_v, ctl) => { const s = await editSeries(app, { values: '' }, ch.type === 'combo'); if (s) { work.push(s); refresh(ctl); } } },
      { key: 'edit', label: 'Edit…', type: 'button', half: true, onClick: async (v, ctl) => { const i = Number(v.series); if (!work[i]) { toast('Choose a series in the list.'); return; } const s = await editSeries(app, work[i], ch.type === 'combo'); if (s) { work[i] = s; refresh(ctl); } } },
      { key: 'remove', label: 'Remove', type: 'button', half: true, onClick: (v, ctl) => { const i = Number(v.series); if (!work[i]) { toast('Choose a series in the list.'); return; } work.splice(i, 1); refresh(ctl); } },
      { key: 'blank', label: 'Hidden and empty cells: empty cells are shown as gaps.', type: 'info' },
    ],
    onChange: (_v, _c, key) => { if (key === 'range' || key === 'rows') rangeChanged = true; },
  });
  if (!r) return;
  if (rangeChanged && String(r.range || '').trim()) {
    const res = resolveRangeText(app.book, String(r.range).replace(/^=/, '').replace(/\$/g, ''), sheet);
    if (!res) { toast('That reference isn\'t valid.'); return; }
    const built = chartFromRange(res.sheet, res.range, ch.type, { seriesInRows: !!r.rows });
    setChart(app, { id: ch.id }, 'Select Data', (c) => { c.series = built.series; c.sourceRange = built.sourceRange; c.seriesInRows = !!r.rows; });
    return;
  }
  setChart(app, { id: ch.id }, 'Select Data', (c) => { c.series = work; });
}

async function moveChart(app: App, arg: any): Promise<void> {
  const o = findObject(app, arg && typeof arg === 'object' && typeof arg.id === 'string' ? arg.id : undefined);
  if (!o || o.kind !== 'chart') { toast('Select a chart first.'); return; }
  const book = app.book;
  let n = 1;
  while (book.sheetByName(`Chart${n}`)) n++;
  const r = await openDialog({
    title: 'Move Chart',
    fields: [
      { key: 'where', label: 'Choose where you want the chart to be placed:', type: 'radio', options: [{ value: 'new', label: 'New sheet' }, { value: 'obj', label: 'Object in' }], value: 'new' },
      { key: 'name', label: 'New sheet name', type: 'text', value: `Chart${n}`, visibleIf: (v) => v.where === 'new' },
      { key: 'sheet', label: 'Sheet', type: 'select', options: book.sheets.filter((s) => s.visibility === 'visible').map((s) => s.name), value: o.sheet.name, visibleIf: (v) => v.where === 'obj' },
    ],
    validate: (v) => (v.where === 'new' ? validSheetName(String(v.name)) || (book.sheetByName(String(v.name)) ? 'A sheet with that name already exists.' : null) : null),
  });
  if (!r) return;
  if (book.structureLocked && r.where === 'new') { toast('The workbook structure is protected.'); return; }
  const from = o.sheet;
  const chart = o.def;
  // Unqualified references must keep pointing at the original sheet.
  const qualify = (ref?: string) => (!ref || /!/.test(ref) || /^[{"]/.test(ref) || !parseRange(ref.replace(/\$/g, '')) ? ref : `${quoteSheet(from.name)}!${ref}`);
  const g = app.tab.grid;
  book.transact('Move Chart', () => {
    book.recordBook();
    const target = r.where === 'new' ? new Sheet(book, String(r.name)) : book.sheetByName(String(r.sheet))!;
    if (r.where === 'new') {
      book.sheets.splice(book.sheets.indexOf(from) + 1, 0, target);
      target.showGrid = false; target.showHeadings = false;
    }
    if (target === from) return;
    from.charts = from.charts.filter((c) => c.id !== chart.id);
    for (const s of chart.series) { s.values = qualify(s.values)!; s.categories = qualify(s.categories); s.nameRef = qualify(s.nameRef); s.sizes = qualify(s.sizes); }
    if (chart.sourceRange) chart.sourceRange = qualify(chart.sourceRange);
    if (r.where === 'new') chart.at = { r: 0, c: 0, dx: 0, dy: 0, w: Math.max(400, ((g?.width || 900) - 20) / (g?.zoom || 1)), h: Math.max(260, ((g?.height || 560) - 20) / (g?.zoom || 1)) };
    else chart.at = { ...chart.at, r: target.sel.r, c: target.sel.c, dx: 0, dy: 0 };
    target.charts.push(chart);
    book.activeSheet = book.sheets.indexOf(target);
  });
  app.tab.grid?.layout();
  app.tab.grid?.objects.select(chart.id);
}

async function saveImage(app: App, arg: any): Promise<void> {
  const found = chartOf(app, arg);
  if (!found) { toast('Select a chart first.'); return; }
  const { def: ch, sheet } = found;
  const r = await openDialog({
    title: 'Save as Picture',
    fields: [
      { key: 'name', label: 'File name', type: 'text', value: (ch.title || 'Chart').replace(/[\\/:*?"<>|]+/g, ' ').trim() || 'Chart' },
      { key: 'fmt', label: 'Format', type: 'radio', options: [{ value: 'png', label: 'PNG picture' }, { value: 'svg', label: 'SVG (scalable vector graphics)' }], value: 'png' },
      { key: 'scale', label: 'Resolution', type: 'select', options: [{ value: '1', label: '1× (screen)' }, { value: '2', label: '2× (sharp)' }, { value: '4', label: '4× (print)' }], value: '2', visibleIf: (v) => v.fmt === 'png' },
    ],
  });
  if (!r) return;
  const svg = renderChartSvg(app.tab, ch, ch.at.w, ch.at.h, sheet);
  const blob = r.fmt === 'svg' ? new Blob([svg], { type: 'image/svg+xml' }) : await svgToPng(svg, ch.at.w, ch.at.h, Number(r.scale) || 2);
  if (!blob) { toast('The picture could not be created.'); return; }
  const { saveAs } = await import('../platform/native');
  const name = `${r.name || 'Chart'}.${r.fmt}`;
  const res = await saveAs(blob, name).catch(() => null);
  if (res) toast(`Saved "${res.name}"`);
}

/** Rasterizes SVG markup into a PNG blob. */
export function svgToPng(svg: string, w: number, h: number, scale = 2): Promise<Blob | null> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const cv = document.createElement('canvas');
      cv.width = Math.round(w * scale); cv.height = Math.round(h * scale);
      const ctx = cv.getContext('2d');
      if (!ctx) { resolve(null); return; }
      ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, cv.width, cv.height);
      ctx.drawImage(img, 0, 0, cv.width, cv.height);
      cv.toBlob((b) => resolve(b), 'image/png');
    };
    img.onerror = () => resolve(null);
    img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
  });
}

async function chartTypeDialog(app: App, title: string, current?: ChartType, base?: ChartDef): Promise<ChartType | null> {
  return galleryDialog(title, typeItems(app, CHART_TYPES, base, true), { current, ok: 'OK' });
}

async function comboDialog(app: App): Promise<void> {
  const rg = chartSourceRange(app);
  if (!rg) { insertChart(app, 'combo'); return; }
  const tmp = chartFromRange(app.sheet, rg, 'combo');
  if (tmp.series.length < 2) { toast('A combo chart needs at least two series.'); return; }
  const names = tmp.series.map((s, i) => { try { const v = app.engine.evaluateScalar(s.nameRef || '""', app.sheet, 0, 0); return v ? String(v) : `Series${i + 1}`; } catch { return `Series${i + 1}`; } });
  const fields = names.flatMap((n, i) => [
    { key: 't' + i, label: n, type: 'select' as const, options: ['column', 'stackedColumn', 'line', 'lineMarkers', 'area', 'stackedArea'].map((t) => ({ value: t, label: CHART_TYPE_LABELS[t as ChartType] })), value: i === names.length - 1 ? 'line' : 'column', half: true },
    { key: 's' + i, label: 'Secondary axis', type: 'checkbox' as const, value: false, half: true },
  ]);
  const r = await openDialog({
    title: 'Custom Combination', fields, preview: true, previewStyle: 'display:flex;justify-content:center;background:#fff',
    onChange: (v, ctl) => { const c = { ...tmp, series: tmp.series.map((s, i) => ({ ...s, type: v['t' + i], secondary: !!v['s' + i] })) }; ctl.setPreview(renderChartSvg(app.tab, c, 380, 220, app.sheet)); },
  });
  if (!r) return;
  const ch = insertChart(app, 'combo', { comboTypes: names.map((_, i) => r['t' + i]) });
  if (ch) setChart(app, { id: ch.id }, 'Combo Chart', (c) => { c.series.forEach((s, i) => { if (r['s' + i]) s.secondary = true; }); });
}

// --------------------------------------------------------------------------------------------
// Commands

const chartMenu = (id: string, types: ChartType[], more = true): (() => MenuEntry[]) => () => [
  ...types.map((t) => ({ id, arg: t, label: CHART_TYPE_LABELS[t], icon: CHART_ICONS[t] })),
  ...(more ? [{ separator: true } as MenuEntry, { id: 'insert.chartDialog', label: 'More Charts…', icon: 'more_horiz' }] : []),
];

function insertCmd(id: string, label: string, icon: string, types: ChartType[], keywords: string, popular = false) {
  return {
    id, label, icon, tab: 'Insert', group: 'Charts', kind: 'split' as const, edits: true, popular, keywords,
    run: (app: App, arg?: any) => insertChart(app, (typeof arg === 'string' && (CHART_TYPES as string[]).includes(arg) ? arg : types[0]) as ChartType),
    menu: chartMenu(id, types),
  };
}

const hasChart = (app: App) => app.tab?.grid?.objects?.selectedObject()?.kind === 'chart';

def([
  insertCmd('insert.chartColumn', 'Column Chart', 'bar_chart', ['column', 'stackedColumn', 'percentColumn'], 'chart graph column vertical bar', true),
  insertCmd('insert.chartLine', 'Line Chart', 'show_chart', ['line', 'lineMarkers', 'stackedLine'], 'chart graph line trend', true),
  insertCmd('insert.chartPie', 'Pie Chart', 'pie_chart', ['pie', 'doughnut'], 'chart graph pie doughnut donut', true),
  insertCmd('insert.chartBar', 'Bar Chart', 'align_horizontal_left', ['bar', 'stackedBar', 'percentBar'], 'chart graph horizontal bar'),
  insertCmd('insert.chartArea', 'Area Chart', 'area_chart', ['area', 'stackedArea', 'percentArea'], 'chart graph area'),
  insertCmd('insert.chartScatter', 'Scatter Chart', 'scatter_plot', ['scatter', 'scatterLines', 'bubble'], 'chart graph scatter xy bubble'),
  insertCmd('insert.chartStatistic', 'Statistic Chart', 'equalizer', ['histogram', 'pareto', 'boxWhisker'], 'chart histogram pareto box whisker distribution'),
  insertCmd('insert.chartWaterfall', 'Waterfall, Funnel or Stock Chart', 'waterfall_chart', ['waterfall', 'funnel', 'stock'], 'chart waterfall funnel stock high low close'),
  insertCmd('insert.chartHierarchy', 'Hierarchy Chart', 'grid_view', ['treemap', 'sunburst'], 'chart treemap sunburst hierarchy'),
  insertCmd('insert.chartRadar', 'Radar Chart', 'radar', ['radar', 'filledRadar'], 'chart radar spider web'),
  {
    id: 'insert.chartCombo', label: 'Combo Chart', icon: 'insert_chart', tab: 'Insert', group: 'Charts', kind: 'split', edits: true, keywords: 'chart combination secondary axis',
    run: (app, arg) => {
      if (arg === 'custom') return comboDialog(app);
      if (arg === 'secondary') return insertChart(app, 'combo', { secondaryLast: true });
      if (arg === 'areaColumn') return insertChart(app, 'combo', { comboTypes: ['stackedArea', 'stackedArea', 'column', 'column', 'column', 'column'] });
      return insertChart(app, 'combo');
    },
    menu: () => [
      { id: 'insert.chartCombo', label: 'Clustered Column – Line', icon: 'insert_chart' },
      { id: 'insert.chartCombo', arg: 'secondary', label: 'Clustered Column – Line on Secondary Axis', icon: 'insert_chart' },
      { id: 'insert.chartCombo', arg: 'areaColumn', label: 'Stacked Area – Clustered Column', icon: 'area_chart' },
      { separator: true }, { id: 'insert.chartCombo', arg: 'custom', label: 'Create Custom Combo Chart…', icon: 'tune' },
    ],
  },
  {
    id: 'insert.recommendedCharts', label: 'Recommended Charts', icon: 'auto_graph', tab: 'Insert', group: 'Charts', edits: true, popular: true, keywords: 'chart suggest best',
    run: async (app) => {
      const rg = chartSourceRange(app);
      const recs = recommendTypes(app.sheet, rg);
      const t = await galleryDialog(rg ? 'Recommended Charts' : 'Recommended Charts (select data first for previews)', typeItems(app, recs), { help: rg ? 'Charts that suit the selected data. Double-click to insert.' : 'Select a range of data to see charts made from it.', current: recs[0], ok: 'Insert', extra: h('button', { class: 'ql-btn', type: 'button', onclick: (e: Event) => { (e.target as HTMLElement).closest('.ql-overlay')?.querySelector<HTMLElement>('.ql-dialog-head .ql-iconbtn')?.click(); setTimeout(() => app.exec('insert.chartDialog'), 150); } }, 'All Charts…') });
      if (t) insertChart(app, t);
    },
  },
  {
    id: 'insert.chartDialog', label: 'Insert Chart', icon: 'insert_chart', tab: 'Insert', group: 'Charts', edits: true, keywords: 'all charts chart types',
    run: async (app) => { const t = await chartTypeDialog(app, 'Insert Chart: All Charts'); if (t === 'combo') return comboDialog(app); if (t) insertChart(app, t); },
  },

  // ---- Chart Design
  { id: 'chart.edit', label: 'Edit Chart', icon: 'edit', tab: 'Chart Design', group: 'Chart Layouts', edits: true, run: (app, arg) => chartEditor(app, arg) },
  { id: 'chart.set', label: 'Set Chart Property', icon: 'tune', tab: 'Chart Design', group: 'Chart Layouts', internal: true, edits: true, run: (app, arg: { id?: string; patch: Partial<ChartDef> }) => setChart(app, arg, 'Chart Element', (c) => { for (const [k, v] of Object.entries(arg?.patch || {})) { if (v === undefined || v === null) delete (c as any)[k]; else (c as any)[k] = v; } }) },
  {
    id: 'chart.addElement', label: 'Add Chart Element', icon: 'add_chart', tab: 'Chart Design', group: 'Chart Layouts', kind: 'menu', edits: true, isEnabled: hasChart,
    run: async (app, arg) => {
      if (arg && typeof arg === 'object' && arg.ask) {
        const ch = chartOf(app)?.def;
        const v = await promptDialog(arg.ask === 'title' ? 'Chart Title' : 'Axis Title', 'Text', arg.ask === 'title' ? ch?.title || 'Chart Title' : arg.ask === 'x' ? ch?.xTitle || 'Axis Title' : ch?.yTitle || 'Axis Title');
        if (v == null) return;
        app.exec('chart.set', { patch: arg.ask === 'title' ? { title: v } : arg.ask === 'x' ? { xTitle: v || undefined } : { yTitle: v || undefined } });
        return;
      }
      if (arg && typeof arg === 'object' && arg.patch) app.exec('chart.set', arg);
    },
    menu: (app) => {
      const c = chartOf(app)?.def;
      const set = (patch: Partial<ChartDef>, label: string, checked?: boolean): MenuEntry => ({ id: 'chart.addElement', arg: { patch }, label, checked });
      return [
        { label: 'Axis Titles', icon: 'title', submenu: [{ id: 'chart.addElement', arg: { ask: 'x' }, label: 'Primary Horizontal…', checked: !!c?.xTitle }, { id: 'chart.addElement', arg: { ask: 'y' }, label: 'Primary Vertical…', checked: !!c?.yTitle }, set({ xTitle: undefined, yTitle: undefined } as any, 'None')] },
        { label: 'Chart Title', icon: 'text_fields', submenu: [set({ title: '' }, 'None', c?.title === ''), set({ title: undefined } as any, 'Automatic', c?.title === undefined), { id: 'chart.addElement', arg: { ask: 'title' }, label: 'Above Chart…', checked: !!c?.title }] },
        { label: 'Data Labels', icon: 'label', submenu: [set({ dataLabels: false }, 'None', !c?.dataLabels), set({ dataLabels: true }, 'Show', !!c?.dataLabels)] },
        { label: 'Gridlines', icon: 'grid_4x4', submenu: [set({ gridlines: true }, 'Primary Major Horizontal', c?.gridlines !== false), set({ gridlines: false }, 'None', c?.gridlines === false)] },
        { label: 'Legend', icon: 'list', submenu: (['none', 'right', 'top', 'left', 'bottom'] as const).map((p) => set({ legend: p }, p === 'none' ? 'None' : p[0].toUpperCase() + p.slice(1), c?.legend === p)) },
        { label: 'Trendline', icon: 'trending_up', submenu: [set({ trendline: null }, 'None', !c?.trendline), set({ trendline: 'linear' }, 'Linear', c?.trendline === 'linear'), set({ trendline: 'exponential' }, 'Exponential', c?.trendline === 'exponential'), set({ trendline: 'polynomial' }, 'Polynomial', c?.trendline === 'polynomial'), set({ trendline: 'movingAverage' }, 'Moving Average', c?.trendline === 'movingAverage')] },
        { label: 'Axes', icon: 'straighten', submenu: [set({ logScale: true }, 'Logarithmic Scale', !!c?.logScale), set({ logScale: false }, 'Linear Scale', !c?.logScale), { id: 'chart.edit', label: 'More Axis Options…' }] },
      ];
    },
  },
  {
    id: 'chart.quickLayout', label: 'Quick Layout', icon: 'dashboard', tab: 'Chart Design', group: 'Chart Layouts', kind: 'menu', edits: true, isEnabled: hasChart,
    run: (app, arg) => { const l = QUICK_LAYOUTS[Number(arg) || 0]; setChart(app, null, 'Quick Layout', (c) => { for (const [k, v] of Object.entries(l.patch)) { if (v === undefined) delete (c as any)[k]; else (c as any)[k] = v; } }); },
    menu: () => QUICK_LAYOUTS.map((l, i) => ({ id: 'chart.quickLayout', arg: i, label: l.label, icon: 'dashboard' })),
  },
  {
    id: 'chart.colors', label: 'Change Colors', icon: 'palette', tab: 'Chart Design', group: 'Chart Styles', kind: 'menu', edits: true, isEnabled: hasChart,
    run: (app, arg) => { const p = PALETTES.find((x) => x.id === arg) || PALETTES[0]; setChart(app, null, 'Chart Colors', (c) => { if (p.colors) c.colors = [...p.colors]; else delete c.colors; for (const s of c.series) delete s.color; }); },
    menu: (app) => { const cur = chartOf(app)?.def.colors; return PALETTES.map((p) => ({ id: 'chart.colors', arg: p.id, label: p.label, swatch: (p.colors || app.book.theme.colors.slice(4, 10))[0], checked: JSON.stringify(p.colors || null) === JSON.stringify(cur || null) })); },
  },
  {
    id: 'chart.style', label: 'Chart Styles', icon: 'style', tab: 'Chart Design', group: 'Chart Styles', kind: 'gallery', edits: true,
    run: (app, arg) => { const n = Number(arg) || 1; setChart(app, null, 'Chart Style', (c) => { c.style = n; }); },
    menu: (app) => { const cur = chartOf(app)?.def.style || 1; return CHART_STYLES.map(([n, l]) => ({ id: 'chart.style', arg: n, label: l.replace(/^Style \d+: /, ''), checked: n === cur, preview: `font-size:11px` })); },
  },
  {
    id: 'chart.switchRowCol', label: 'Switch Row/Column', icon: 'swap_horiz', tab: 'Chart Design', group: 'Data', edits: true, isEnabled: hasChart,
    run: (app, arg) => {
      const found = chartOf(app, arg);
      if (!found) { toast('Select a chart first.'); return; }
      const src = inferSourceRange(app, found.def, found.sheet);
      if (!src) { toast('Switching needs the chart\'s data in one range on one sheet. Use Select Data instead.'); return; }
      const built = chartFromRange(src.sheet, src.range, found.def.type, { seriesInRows: !found.def.seriesInRows });
      setChart(app, { id: found.def.id }, 'Switch Row/Column', (c) => { c.series = built.series; c.seriesInRows = built.seriesInRows; c.sourceRange = built.sourceRange; });
    },
  },
  { id: 'chart.selectData', label: 'Select Data', icon: 'dataset', tab: 'Chart Design', group: 'Data', edits: true, run: (app, arg) => selectData(app, arg) },
  {
    id: 'chart.changeType', label: 'Change Chart Type', icon: 'bar_chart', tab: 'Chart Design', group: 'Type', edits: true,
    run: async (app, arg) => {
      const found = chartOf(app, arg);
      if (!found) { toast('Select a chart first.'); return; }
      const t = typeof arg === 'string' && (CHART_TYPES as string[]).includes(arg) ? (arg as ChartType) : await chartTypeDialog(app, 'Change Chart Type', found.def.type, found.def);
      if (!t) return;
      setChart(app, { id: found.def.id }, 'Change Chart Type', (c) => {
        c.type = t;
        if (t !== 'combo') for (const s of c.series) delete s.type;
        else c.series.forEach((s, i) => { s.type = s.type || (i === c.series.length - 1 && i > 0 ? 'line' : 'column'); });
        c.gridlines = !['pie', 'doughnut', 'treemap', 'sunburst', 'funnel'].includes(t) && c.gridlines !== false;
      });
    },
  },
  { id: 'chart.move', label: 'Move Chart', icon: 'drive_file_move', tab: 'Chart Design', group: 'Location', edits: true, run: (app, arg) => moveChart(app, arg) },
  { id: 'chart.saveImage', label: 'Save as Picture', icon: 'image', tab: 'Chart Design', group: 'Location', noRecord: true, run: (app, arg) => saveImage(app, arg) },
  { id: 'chart.delete', label: 'Delete Chart', icon: 'delete', tab: 'Chart Design', group: 'Location', edits: true, isEnabled: hasChart, run: (app, arg) => { const c = chartOf(app, arg); if (!c) { toast('Select a chart first.'); return; } app.exec('object.delete', { id: c.def.id }); } },
  {
    id: 'chart.fromColumns', label: 'Chart from Columns', icon: 'insert_chart', tab: 'Insert', group: 'Charts', internal: true, edits: true,
    run: (app, arg: { type?: ChartType; range?: string }) => {
      if (arg?.range) {
        const res = resolveRangeText(app.book, String(arg.range).replace(/\$/g, ''), app.sheet);
        if (!res) return null;
        const ch = chartFromRange(res.sheet, res.range, arg.type || 'column');
        return addChart(app, ch);
      }
      return insertChart(app, arg?.type || 'column');
    },
  },
]);

void colName; void splitSheetRef;
