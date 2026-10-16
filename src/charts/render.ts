/**
 * Chart renderer: turns a ChartDef plus the workbook data it references into SVG markup.
 *
 * Pure string building (no DOM) so it is fast enough to run on every recalculation and works in
 * tests. Series data come from formula text (ranges, names, arrays) evaluated by the calc engine
 * in the context of the chart's sheet. Every piece of text is escaped.
 */
import type { BookTab } from '../app';
import type { ChartDef, ChartType, Scalar, Matrix } from '../model/types';
import { isError } from '../model/types';
import type { Sheet, Workbook } from '../model/workbook';
import { escHtml } from '../util/sanitize';
import { formatValue, isDateFormat } from '../model/numfmt';
import { resolveRangeText } from '../model/structure';

// --------------------------------------------------------------------------------------------
// Palette

export const DEFAULT_CHART_COLORS = ['#2f6fb0', '#e07b28', '#7a8a99', '#e3b505', '#4a90c2', '#5aa33f'];

function hexRgb(c: string): number[] | null {
  const m = /^#([0-9a-f]{6})$/i.exec(c) || /^#([0-9a-f]{3})$/i.exec(c);
  if (!m) return null;
  const s = m[1].length === 3 ? m[1].split('').map((x) => x + x).join('') : m[1];
  return [0, 2, 4].map((i) => parseInt(s.slice(i, i + 2), 16));
}
export function mixColor(a: string, b: string, t: number): string {
  const x = hexRgb(a), y = hexRgb(b);
  if (!x || !y) return a;
  return '#' + x.map((v, i) => Math.round(v + (y[i] - v) * t).toString(16).padStart(2, '0')).join('');
}
const safeCol = (c: unknown, fb: string) => (typeof c === 'string' && /^(#[0-9a-f]{3,8}|[a-z]{3,20}|rgba?\([\d\s.,%]+\))$/i.test(c) ? c : fb);

/** Series/point colors: the chart's own colors, else the theme accents, then darker and lighter variants. */
export function chartPalette(book: Workbook | null, chart: Pick<ChartDef, 'colors'>, n: number): string[] {
  const base = chart.colors?.length ? chart.colors.map((c, i) => safeCol(c, DEFAULT_CHART_COLORS[i % 6])) : (book?.theme?.colors?.slice(4, 10) || DEFAULT_CHART_COLORS).map((c, i) => safeCol(c, DEFAULT_CHART_COLORS[i]));
  const out: string[] = [];
  for (let i = 0; i < Math.max(n, 1); i++) {
    const c = base[i % base.length];
    const round = Math.floor(i / base.length);
    out.push(round === 0 ? c : round % 2 ? mixColor(c, '#000000', 0.3 * Math.ceil(round / 2)) : mixColor(c, '#ffffff', 0.35 * (round / 2)));
  }
  return out;
}

// --------------------------------------------------------------------------------------------
// Data extraction

export interface SeriesData {
  name: string;
  values: (number | null)[];
  xs?: (number | null)[];
  sizes?: (number | null)[];
  color: string;
  type: ChartType;
  secondary: boolean;
  fmt?: string;
  errors: number;
}

export interface ChartData {
  categories: string[];
  catRaw: Scalar[];
  /** Multi-level category labels (one array per point, outer level first) for hierarchy charts. */
  catLevels: string[][];
  catFmt?: string;
  series: SeriesData[];
  /** Colors for points (pie slices, treemap groups) in palette order. */
  palette: string[];
}

export function chartSheet(tab: Pick<BookTab, 'book'>, chart: ChartDef): Sheet {
  return tab.book.sheets.find((s) => s.charts.includes(chart) || s.charts.some((c) => c.id === chart.id)) || tab.book.sheet;
}

/** Evaluates a series formula (range, name, array constant) to a matrix; errors become a 1×1 error. */
export function evalRef(tab: Pick<BookTab, 'engine'>, ref: string | undefined, sheet: Sheet): Matrix {
  if (!ref || !String(ref).trim()) return [];
  try {
    return tab.engine.evaluateMatrix(String(ref).replace(/^=/, ''), sheet, 0, 0);
  } catch {
    return [];
  }
}

/** Number format of the first cell a series reference points at. */
export function refFormat(book: Workbook, ref: string | undefined, sheet: Sheet): string | undefined {
  if (!ref) return undefined;
  try {
    const r = resolveRangeText(book, String(ref).replace(/^=/, '').replace(/\$/g, ''), sheet);
    if (!r) return undefined;
    const f = r.sheet.styleOf(r.range.r1, r.range.c1).numFmt;
    return f && f !== 'General' ? f : undefined;
  } catch { return undefined; }
}

const flat = (m: Matrix): Scalar[] => {
  if (!m.length) return [];
  // A row vector or column vector flattens in reading order; a block reads column by column for values.
  if (m.length === 1 || m[0].length === 1) return m.flat();
  return m.flat();
};
const num = (v: Scalar): number | null => (typeof v === 'number' && Number.isFinite(v) ? v : typeof v === 'boolean' ? (v ? 1 : 0) : null);

export function scalarText(v: Scalar, fmt?: string, date1904 = false): string {
  if (v == null) return '';
  if (isError(v)) return v.code;
  if (typeof v === 'number') return formatValue(v, fmt, { date1904 }).text;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

/** Reads every series of a chart from the workbook. */
export function chartData(tab: Pick<BookTab, 'book' | 'engine'>, chart: ChartDef, sheet?: Sheet): ChartData {
  const sh = sheet || chartSheet(tab, chart);
  const book = tab.book;
  const d1904 = !!book.calc?.date1904;
  const pal = chartPalette(book, chart, chart.series.length);
  let categories: string[] = [];
  let catRaw: Scalar[] = [];
  let catLevels: string[][] = [];
  let catFmt: string | undefined;
  const firstCat = chart.series.find((s) => s.categories)?.categories;
  if (firstCat) {
    const m = evalRef(tab, firstCat, sh);
    catFmt = refFormat(book, firstCat, sh);
    const vertical = m.length >= (m[0]?.length || 0);
    if (m.length && m[0].length > 1 && m.length > 1) {
      // Multi-level categories: rows are points (vertical) or columns are points (horizontal).
      const pts = vertical ? m : m[0].map((_, j) => m.map((row) => row[j]));
      const last: string[] = [];
      catLevels = pts.map((row) => row.map((v, li) => {
        const t = scalarText(v, catFmt, d1904);
        if (t) { last[li] = t; last.length = li + 1; return t; }
        return last[li] ?? '';
      }));
      catRaw = pts.map((row) => row[row.length - 1]);
      categories = catLevels.map((l) => l[l.length - 1]);
    } else {
      catRaw = flat(m);
      categories = catRaw.map((v) => scalarText(v, catFmt, d1904));
      catLevels = categories.map((c) => [c]);
    }
  }
  const series: SeriesData[] = chart.series.map((s, i) => {
    const raw = flat(evalRef(tab, s.values, sh));
    let errors = 0;
    const values = raw.map((v) => { if (isError(v)) errors++; return num(v); });
    let name = s.name ?? '';
    if (s.nameRef) { const nm = flat(evalRef(tab, s.nameRef, sh)); if (nm.length) name = scalarText(nm[0], refFormat(book, s.nameRef, sh), d1904); }
    if (!name) name = `Series${i + 1}`;
    const out: SeriesData = { name, values, color: safeCol(s.color, pal[i]), type: s.type || chart.type, secondary: !!s.secondary, fmt: refFormat(book, s.values, sh), errors };
    if (s.categories && (chart.type === 'scatter' || chart.type === 'scatterLines' || chart.type === 'bubble')) out.xs = flat(evalRef(tab, s.categories, sh)).map(num);
    if (s.sizes) out.sizes = flat(evalRef(tab, s.sizes, sh)).map(num);
    return out;
  });
  const n = Math.max(0, ...series.map((s) => s.values.length));
  if (!categories.length) { categories = Array.from({ length: n }, (_, i) => String(i + 1)); catRaw = categories.map((_, i) => i + 1); catLevels = categories.map((c) => [c]); }
  return { categories, catRaw, catLevels, catFmt, series, palette: chartPalette(book, chart, Math.max(n, series.length)) };
}

// --------------------------------------------------------------------------------------------
// Scales

export interface Scale { min: number; max: number; step: number; ticks: number[]; log?: boolean }

/** "Nice" axis bounds and ticks with 1-2-5 steps (or powers of ten for a log axis). */
export function niceScale(min: number, max: number, maxTicks = 6, opts: { includeZero?: boolean; log?: boolean; fixedMin?: number; fixedMax?: number } = {}): Scale {
  if (!Number.isFinite(min) || !Number.isFinite(max)) { min = 0; max = 1; }
  if (opts.log) {
    const lo = opts.fixedMin && opts.fixedMin > 0 ? Math.log10(opts.fixedMin) : Math.floor(Math.log10(Math.max(min, 1e-12)));
    const hi = opts.fixedMax && opts.fixedMax > 0 ? Math.log10(opts.fixedMax) : Math.ceil(Math.log10(Math.max(max, 1e-12)));
    const a = Math.min(lo, hi - 1), b = Math.max(hi, lo + 1);
    const ticks: number[] = [];
    for (let e = Math.ceil(a); e <= Math.floor(b) && ticks.length < 40; e++) ticks.push(Math.pow(10, e));
    return { min: Math.pow(10, a), max: Math.pow(10, b), step: 10, ticks, log: true };
  }
  if (opts.includeZero) { min = Math.min(0, min); max = Math.max(0, max); }
  if (opts.fixedMin != null) min = opts.fixedMin;
  if (opts.fixedMax != null) max = opts.fixedMax;
  if (min === max) { if (min === 0) { max = 1; } else { const d = Math.abs(min) * 0.1; min -= d; max += d; if (opts.includeZero) { min = Math.min(0, min); max = Math.max(0, max); } } }
  if (min > max) [min, max] = [max, min];
  const rough = (max - min) / Math.max(1, maxTicks);
  const mag = Math.pow(10, Math.floor(Math.log10(rough)));
  const norm = rough / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  const lo = opts.fixedMin != null ? min : Math.floor(min / step + 1e-9) * step;
  let hi = opts.fixedMax != null ? max : Math.ceil(max / step - 1e-9) * step;
  if (hi <= lo) hi = lo + step;
  const ticks: number[] = [];
  const dec = Math.max(0, -Math.floor(Math.log10(step)) + 2);
  const first = Math.ceil(lo / step - 1e-9) * step;
  for (let v = first; v <= hi + step * 1e-6 && ticks.length < 200; v += step) ticks.push(+v.toFixed(Math.min(15, dec)));
  return { min: lo, max: hi, step, ticks };
}

const scalePos = (s: Scale, v: number, a: number, b: number) => {
  if (s.log) { const lv = Math.log10(Math.max(v, s.min)); const l0 = Math.log10(s.min), l1 = Math.log10(s.max); return a + ((lv - l0) / (l1 - l0 || 1)) * (b - a); }
  return a + ((v - s.min) / (s.max - s.min || 1)) * (b - a);
};

// --------------------------------------------------------------------------------------------
// Statistics helpers

export interface Bin { lo: number; hi: number; count: number; label: string }

/** Histogram bins: Scott's normal reference rule (as spreadsheets do automatically), or a given bin count/width. */
export function histogramBins(values: number[], opts: { binCount?: number; binWidth?: number; fmt?: (n: number) => string } = {}): Bin[] {
  const v = values.filter((x) => Number.isFinite(x)).sort((a, b) => a - b);
  if (!v.length) return [];
  const min = v[0], max = v[v.length - 1];
  const fmt = opts.fmt || ((n: number) => String(+n.toPrecision(6)));
  if (min === max) return [{ lo: min, hi: max, count: v.length, label: `[${fmt(min)}, ${fmt(max)}]` }];
  let width = opts.binWidth && opts.binWidth > 0 ? opts.binWidth : 0;
  if (!width) {
    if (opts.binCount && opts.binCount > 0) width = (max - min) / opts.binCount;
    else {
      const mean = v.reduce((a, b) => a + b, 0) / v.length;
      const sd = Math.sqrt(v.reduce((a, b) => a + (b - mean) ** 2, 0) / Math.max(1, v.length - 1));
      width = sd > 0 ? (3.49 * sd) / Math.cbrt(v.length) : (max - min) / 5;
    }
  }
  let count = Math.max(1, Math.ceil((max - min) / width - 1e-9));
  if (count > 500) { count = 500; width = (max - min) / count; }
  const bins: Bin[] = Array.from({ length: count }, (_, i) => ({ lo: min + i * width, hi: min + (i + 1) * width, count: 0, label: '' }));
  for (const x of v) {
    let i = Math.ceil((x - min) / width - 1e-9) - 1; // (lo, hi] intervals, first bin closed
    if (i < 0) i = 0;
    if (i >= count) i = count - 1;
    bins[i].count++;
  }
  bins.forEach((b, i) => { b.label = `${i === 0 ? '[' : '('}${fmt(b.lo)}, ${fmt(b.hi)}]`; });
  return bins;
}

export function quartiles(values: number[]): { min: number; q1: number; median: number; q3: number; max: number; mean: number; lowW: number; highW: number; outliers: number[] } | null {
  const v = values.filter((x) => Number.isFinite(x)).sort((a, b) => a - b);
  if (!v.length) return null;
  const q = (p: number) => { const pos = (v.length - 1) * p; const i = Math.floor(pos); return v[i] + (v[Math.min(i + 1, v.length - 1)] - v[i]) * (pos - i); };
  const q1 = q(0.25), q3 = q(0.75), iqr = q3 - q1;
  const lo = q1 - 1.5 * iqr, hi = q3 + 1.5 * iqr;
  const inside = v.filter((x) => x >= lo && x <= hi);
  return { min: v[0], q1, median: q(0.5), q3, max: v[v.length - 1], mean: v.reduce((a, b) => a + b, 0) / v.length, lowW: inside[0] ?? q1, highW: inside[inside.length - 1] ?? q3, outliers: v.filter((x) => x < lo || x > hi) };
}

export interface Rect { x: number; y: number; w: number; h: number }

/** Squarified treemap layout: returns one rectangle per item, in input order. Items must be ≥ 0. */
export function squarify(values: number[], box: Rect): Rect[] {
  const out: Rect[] = values.map(() => ({ x: box.x, y: box.y, w: 0, h: 0 }));
  const idx = values.map((v, i) => ({ v: Math.max(0, v || 0), i })).filter((x) => x.v > 0).sort((a, b) => b.v - a.v);
  const total = idx.reduce((a, b) => a + b.v, 0);
  if (!total || box.w <= 0 || box.h <= 0) return out;
  const scale = (box.w * box.h) / total;
  const items = idx.map((x) => ({ a: x.v * scale, i: x.i }));
  let { x, y, w, h } = box;
  const worst = (row: { a: number }[], side: number) => {
    const s = row.reduce((p, r) => p + r.a, 0);
    const mx = Math.max(...row.map((r) => r.a)), mn = Math.min(...row.map((r) => r.a));
    return Math.max((side * side * mx) / (s * s), (s * s) / (side * side * mn));
  };
  let row: { a: number; i: number }[] = [];
  const layout = (rw: { a: number; i: number }[]) => {
    const s = rw.reduce((p, r) => p + r.a, 0);
    if (w >= h) {
      const cw = s / h;
      let yy = y;
      for (const r of rw) { const rh = r.a / cw; out[r.i] = { x, y: yy, w: cw, h: rh }; yy += rh; }
      x += cw; w -= cw;
    } else {
      const rh = s / w;
      let xx = x;
      for (const r of rw) { const cw = r.a / rh; out[r.i] = { x: xx, y, w: cw, h: rh }; xx += cw; }
      y += rh; h -= rh;
    }
  };
  for (const it of items) {
    const side = Math.min(w, h);
    if (!row.length || worst([...row, it], side) <= worst(row, side)) row.push(it);
    else { layout(row); row = [it]; }
  }
  if (row.length) layout(row);
  return out;
}

/** Trendline points for (x, y) data: linear, exponential, polynomial (order 2) or moving average (period 2). */
export function trendline(kind: NonNullable<ChartDef['trendline']>, xs: number[], ys: number[], samples = 40): { x: number; y: number }[] {
  const pts = xs.map((x, i) => ({ x, y: ys[i] })).filter((p) => Number.isFinite(p.x) && Number.isFinite(p.y));
  if (pts.length < 2) return [];
  if (kind === 'movingAverage') {
    const out: { x: number; y: number }[] = [];
    for (let i = 1; i < pts.length; i++) out.push({ x: pts[i].x, y: (pts[i].y + pts[i - 1].y) / 2 });
    return out;
  }
  const x0 = Math.min(...pts.map((p) => p.x)), x1 = Math.max(...pts.map((p) => p.x));
  let fn: ((x: number) => number) | null = null;
  if (kind === 'linear') {
    const n = pts.length, mx = pts.reduce((a, p) => a + p.x, 0) / n, my = pts.reduce((a, p) => a + p.y, 0) / n;
    const sxx = pts.reduce((a, p) => a + (p.x - mx) ** 2, 0), sxy = pts.reduce((a, p) => a + (p.x - mx) * (p.y - my), 0);
    const b = sxx ? sxy / sxx : 0; const a = my - b * mx;
    fn = (x) => a + b * x;
  } else if (kind === 'exponential') {
    const pp = pts.filter((p) => p.y > 0);
    if (pp.length < 2) return [];
    const n = pp.length, mx = pp.reduce((a, p) => a + p.x, 0) / n, my = pp.reduce((a, p) => a + Math.log(p.y), 0) / n;
    const sxx = pp.reduce((a, p) => a + (p.x - mx) ** 2, 0), sxy = pp.reduce((a, p) => a + (p.x - mx) * (Math.log(p.y) - my), 0);
    const b = sxx ? sxy / sxx : 0; const a = my - b * mx;
    fn = (x) => Math.exp(a + b * x);
  } else {
    // Quadratic least squares via normal equations.
    let s0 = 0, s1 = 0, s2 = 0, s3 = 0, s4 = 0, t0 = 0, t1 = 0, t2 = 0;
    for (const p of pts) { const x = p.x; s0++; s1 += x; s2 += x * x; s3 += x ** 3; s4 += x ** 4; t0 += p.y; t1 += x * p.y; t2 += x * x * p.y; }
    const M = [[s0, s1, s2, t0], [s1, s2, s3, t1], [s2, s3, s4, t2]];
    for (let c = 0; c < 3; c++) {
      let piv = c; for (let r = c + 1; r < 3; r++) if (Math.abs(M[r][c]) > Math.abs(M[piv][c])) piv = r;
      [M[c], M[piv]] = [M[piv], M[c]];
      if (Math.abs(M[c][c]) < 1e-12) return trendline('linear', xs, ys, samples);
      for (let r = 0; r < 3; r++) if (r !== c) { const k = M[r][c] / M[c][c]; for (let j = c; j < 4; j++) M[r][j] -= k * M[c][j]; }
    }
    const a = M[0][3] / M[0][0], b = M[1][3] / M[1][1], cc = M[2][3] / M[2][2];
    fn = (x) => a + b * x + cc * x * x;
  }
  const out: { x: number; y: number }[] = [];
  for (let i = 0; i <= samples; i++) { const x = x0 + ((x1 - x0) * i) / samples; out.push({ x, y: fn!(x) }); }
  return out;
}

// --------------------------------------------------------------------------------------------
// SVG helpers

const F = (n: number) => (Number.isFinite(n) ? Math.round(n * 100) / 100 : 0).toString();
const FONT = 'Carlito, Calibri, Arial, sans-serif';
const textW = (s: string, size: number) => s.length * size * 0.55;
const TXT = '#404040', AXIS = '#bfbfbf', GRID = '#e3e3e3';

function text(x: number, y: number, s: string, o: { size?: number; anchor?: 'start' | 'middle' | 'end'; color?: string; bold?: boolean; rotate?: number; baseline?: 'middle' | 'hanging' | 'auto' } = {}): string {
  if (!s) return '';
  const size = o.size ?? 11;
  const dy = o.baseline === 'middle' ? size * 0.35 : o.baseline === 'hanging' ? size * 0.8 : 0;
  const tr = o.rotate ? ` transform="rotate(${F(o.rotate)} ${F(x)} ${F(y)})"` : '';
  return `<text x="${F(x)}" y="${F(y + dy)}" font-size="${F(size)}" text-anchor="${o.anchor || 'start'}" fill="${o.color || TXT}"${o.bold ? ' font-weight="bold"' : ''}${tr}>${escHtml(s)}</text>`;
}

function clip(s: string, maxW: number, size: number): string {
  if (textW(s, size) <= maxW) return s;
  const n = Math.max(1, Math.floor(maxW / (size * 0.55)) - 1);
  return s.slice(0, n) + '…';
}

function fmtTick(v: number, fmt: string | undefined, step: number, d1904: boolean): string {
  if (fmt) return formatValue(v, fmt, { date1904: d1904 }).text;
  const dec = step >= 1 || !Number.isFinite(step) ? 0 : Math.min(10, Math.ceil(-Math.log10(step) - 1e-9));
  const r = +v.toFixed(dec);
  const abs = Math.abs(r);
  if (abs >= 1e9 || (abs > 0 && abs < 1e-6)) return r.toExponential(1).replace('e+', 'E+').replace('e-', 'E-');
  return r.toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec });
}

interface LegendItem { label: string; color: string; kind: 'rect' | 'line' | 'marker' }

interface Frame { svg: string[]; plot: Rect; w: number; h: number }

/** Lays out the chart frame: background, title, legend; returns the remaining plot rectangle. */
function frame(chart: ChartDef, w: number, h: number, title: string, legend: LegendItem[], legendPos: ChartDef['legend']): Frame {
  const svg: string[] = [];
  svg.push(`<rect x="0.5" y="0.5" width="${F(w - 1)}" height="${F(h - 1)}" fill="#ffffff" stroke="#d9d9d9"/>`);
  let top = 10, bottom = h - 8, left = 10, right = w - 10;
  if (title) {
    const size = Math.max(10, Math.min(18, w / 24));
    svg.push(text(w / 2, top + size, clip(title, w - 20, size), { size, anchor: 'middle', color: '#262626' }));
    top += size + 10;
  }
  if (legend.length && legendPos && legendPos !== 'none') {
    const size = 11, sw = 10;
    if (legendPos === 'right' || legendPos === 'left') {
      const maxLabel = Math.min(w * 0.3, Math.max(...legend.map((l) => textW(l.label, size))) + sw + 10);
      const lh = 17;
      const totalH = legend.length * lh;
      let y = Math.max(top, (top + bottom) / 2 - totalH / 2);
      const x = legendPos === 'right' ? right - maxLabel : left;
      for (const it of legend) {
        if (y + lh > bottom) break;
        svg.push(legendSwatch(x, y + lh / 2, it), text(x + sw + 6, y + lh / 2, clip(it.label, maxLabel - sw - 8, size), { size, baseline: 'middle' }));
        y += lh;
      }
      if (legendPos === 'right') right -= maxLabel + 8; else left += maxLabel + 8;
    } else {
      // Horizontal legend, wrapped into rows.
      const rows: LegendItem[][] = [[]];
      let rowW = 0;
      const itemW = (it: LegendItem) => Math.min(w * 0.45, textW(it.label, size)) + sw + 18;
      for (const it of legend) { const iw = itemW(it); if (rowW + iw > w - 20 && rows[rows.length - 1].length) { rows.push([]); rowW = 0; } rows[rows.length - 1].push(it); rowW += iw; }
      const lh = 17;
      const rowsShown = rows.slice(0, 3);
      const blockH = rowsShown.length * lh;
      let y = legendPos === 'top' ? top : bottom - blockH;
      for (const row of rowsShown) {
        const tw = row.reduce((a, it) => a + itemW(it), 0);
        let x = (w - tw) / 2;
        for (const it of row) { svg.push(legendSwatch(x, y + lh / 2, it), text(x + sw + 5, y + lh / 2, clip(it.label, w * 0.45, size), { size, baseline: 'middle' })); x += itemW(it); }
        y += lh;
      }
      if (legendPos === 'top') top += blockH + 4; else bottom -= blockH + 4;
    }
  }
  return { svg, plot: { x: left, y: top, w: Math.max(10, right - left), h: Math.max(10, bottom - top) }, w, h };
}

function legendSwatch(x: number, y: number, it: LegendItem): string {
  if (it.kind === 'line') return `<line x1="${F(x - 2)}" y1="${F(y)}" x2="${F(x + 12)}" y2="${F(y)}" stroke="${it.color}" stroke-width="2.25"/>`;
  if (it.kind === 'marker') return `<circle cx="${F(x + 5)}" cy="${F(y)}" r="4" fill="${it.color}"/>`;
  return `<rect x="${F(x)}" y="${F(y - 5)}" width="10" height="10" fill="${it.color}"/>`;
}

function svgWrap(w: number, h: number, parts: string[]): string {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${F(w)}" height="${F(h)}" viewBox="0 0 ${F(w)} ${F(h)}" font-family="${FONT}" class="ql-chart-svg">${parts.join('')}</svg>`;
}

function message(w: number, h: number, title: string, msg: string): string {
  const parts = [`<rect x="0.5" y="0.5" width="${F(w - 1)}" height="${F(h - 1)}" fill="#ffffff" stroke="#d9d9d9"/>`];
  if (title) parts.push(text(w / 2, 26, clip(title, w - 20, 16), { size: 16, anchor: 'middle', color: '#262626' }));
  parts.push(text(w / 2, h / 2, clip(msg, w - 20, 12), { size: 12, anchor: 'middle', color: '#7f7f7f' }));
  return svgWrap(w, h, parts);
}

// --------------------------------------------------------------------------------------------
// Renderer

const STACKED = new Set<ChartType>(['stackedColumn', 'stackedBar', 'stackedLine', 'stackedArea']);
const PERCENT = new Set<ChartType>(['percentColumn', 'percentBar', 'percentArea']);
const BAR_H = new Set<ChartType>(['bar', 'stackedBar', 'percentBar']);
const COLUMNISH = new Set<ChartType>(['column', 'stackedColumn', 'percentColumn', 'bar', 'stackedBar', 'percentBar', 'histogram', 'pareto', 'waterfall']);
const LINEISH = new Set<ChartType>(['line', 'lineMarkers', 'stackedLine']);
const AREAISH = new Set<ChartType>(['area', 'stackedArea', 'percentArea']);

export const CHART_TYPE_LABELS: Record<ChartType, string> = {
  column: 'Clustered Column', stackedColumn: 'Stacked Column', percentColumn: '100% Stacked Column', bar: 'Clustered Bar', stackedBar: 'Stacked Bar', percentBar: '100% Stacked Bar',
  line: 'Line', lineMarkers: 'Line with Markers', stackedLine: 'Stacked Line', area: 'Area', stackedArea: 'Stacked Area', percentArea: '100% Stacked Area',
  pie: 'Pie', doughnut: 'Doughnut', scatter: 'Scatter', scatterLines: 'Scatter with Lines', bubble: 'Bubble', radar: 'Radar', filledRadar: 'Filled Radar',
  histogram: 'Histogram', pareto: 'Pareto', boxWhisker: 'Box and Whisker', waterfall: 'Waterfall', funnel: 'Funnel', treemap: 'Treemap', sunburst: 'Sunburst',
  stock: 'High-Low-Close', combo: 'Combo',
};
export const CHART_TYPES = Object.keys(CHART_TYPE_LABELS) as ChartType[];

/** Renders a chart as SVG markup sized w × h. */
export function renderChartSvg(tab: BookTab, chart: ChartDef, w: number, h: number, sheet?: Sheet): string {
  w = Math.max(40, Math.min(8000, Number(w) || 400)); h = Math.max(30, Math.min(8000, Number(h) || 260));
  let data: ChartData;
  try { data = chartData(tab, chart, sheet); } catch (e) { return message(w, h, chart.title || '', `Could not read the chart data: ${(e as Error).message}`); }
  const d1904 = !!tab.book.calc?.date1904;
  const autoTitle = chart.title ?? (data.series.length === 1 && !['scatter', 'scatterLines', 'bubble', 'histogram', 'boxWhisker'].includes(chart.type) ? data.series[0].name : '');
  const title = autoTitle || '';
  const hasValues = data.series.some((s) => s.values.some((v) => v != null));
  if (!data.series.length || !hasValues) {
    const errs = data.series.reduce((a, s) => a + s.errors, 0);
    return message(w, h, title, errs ? 'The chart data contain only errors.' : 'No data to chart. Use Select Data to choose a range.');
  }
  try {
    switch (chart.type) {
      case 'pie': case 'doughnut': return pieChart(chart, data, w, h, title, d1904);
      case 'radar': case 'filledRadar': return radarChart(chart, data, w, h, title, d1904);
      case 'funnel': return funnelChart(chart, data, w, h, title, d1904);
      case 'treemap': return treemapChart(chart, data, w, h, title, d1904);
      case 'sunburst': return sunburstChart(chart, data, w, h, title, d1904);
      case 'scatter': case 'scatterLines': case 'bubble': return xyChart(chart, data, w, h, title, d1904);
      default: return categoryChart(chart, data, w, h, title, d1904);
    }
  } catch (e) {
    return message(w, h, title, `Could not draw the chart: ${(e as Error).message}`);
  }
}

// ---- Axis drawing --------------------------------------------------------------------------

function valueAxisWidth(s: Scale, fmt: string | undefined, d1904: boolean): number {
  return Math.max(...s.ticks.map((t) => textW(fmtTick(t, fmt, s.step, d1904), 10))) + 8;
}

/** Category / value chart (column, bar, line, area, combo, histogram, pareto, waterfall, box & whisker, stock). */
function categoryChart(chart: ChartDef, data: ChartData, w: number, h: number, title: string, d1904: boolean): string {
  const type = chart.type;
  const horizontal = BAR_H.has(type);
  let cats = data.categories.slice();
  let series = data.series.map((s) => ({ ...s, values: s.values.slice() }));
  let valFmt = series[0]?.fmt;
  const percent = PERCENT.has(type);
  const stacked = STACKED.has(type) || percent;
  let gapRatio = 0.5;
  let secondaryFmt: string | undefined;
  let waterfall: { base: number; val: number; total: boolean }[] | null = null;
  let boxes: { cat: number; si: number; q: NonNullable<ReturnType<typeof quartiles>> }[] | null = null;

  if (type === 'histogram') {
    const vals = series.flatMap((s) => s.values).filter((v): v is number => v != null);
    const bins = histogramBins(vals, { fmt: (n) => fmtTick(n, valFmt, Math.abs(n) >= 1 ? 1 : 0.01, d1904) });
    cats = bins.map((b) => b.label);
    series = [{ ...series[0], name: series[0].name, values: bins.map((b) => b.count), type: 'column' }];
    valFmt = undefined;
    gapRatio = 0.02;
  } else if (type === 'pareto') {
    const s0 = series[0];
    const pairs = cats.map((c, i) => ({ c, v: s0.values[i] ?? 0 })).filter((p) => p.v != null).sort((a, b) => b.v - a.v);
    cats = pairs.map((p) => p.c);
    const total = pairs.reduce((a, p) => a + Math.max(0, p.v), 0) || 1;
    let run = 0;
    const cum = pairs.map((p) => { run += Math.max(0, p.v); return run / total; });
    series = [{ ...s0, values: pairs.map((p) => p.v), type: 'column' }, { ...s0, name: 'Cumulative %', values: cum, type: 'line', secondary: true, color: mixColor(s0.color, '#e07b28', 0.85), fmt: '0%' }];
    secondaryFmt = '0%';
    gapRatio = 0.1;
  } else if (type === 'waterfall') {
    const s0 = series[0];
    let run = 0;
    waterfall = s0.values.map((v, i) => {
      const total = /\b(sub)?totals?\b/i.test(cats[i] || '');
      if (total) { const r = { base: 0, val: v ?? run, total: true }; run = v ?? run; return r; }
      const r = { base: run, val: v ?? 0, total: false };
      run += v ?? 0;
      return r;
    });
    series = [s0];
  } else if (type === 'boxWhisker') {
    boxes = [];
    const hasCats = data.series.some((s) => s.values.length === cats.length) && new Set(cats).size < cats.length;
    const groups = hasCats ? [...new Set(cats)] : [''];
    series.forEach((s, si) => groups.forEach((g, gi) => {
      const vals = s.values.filter((v, i): v is number => v != null && (!hasCats || cats[i] === g));
      const q = quartiles(vals);
      if (q) boxes!.push({ cat: gi, si, q });
    }));
    cats = groups;
  }

  const n = Math.max(1, cats.length, ...series.map((s) => s.values.length));
  while (cats.length < n) cats.push(String(cats.length + 1));
  const seriesType = (s: SeriesData, i: number): ChartType => {
    if (type === 'combo') return s.type && s.type !== 'combo' ? s.type : i === 0 ? 'column' : 'line';
    if (type === 'pareto') return s.type;
    if (type === 'histogram' || type === 'waterfall') return 'column';
    return type;
  };
  const sTypes = series.map(seriesType);

  // Value extents per axis.
  const ext = [{ min: Infinity, max: -Infinity }, { min: Infinity, max: -Infinity }];
  const note = (ax: number, v: number) => { if (Number.isFinite(v)) { ext[ax].min = Math.min(ext[ax].min, v); ext[ax].max = Math.max(ext[ax].max, v); } };
  const axisOf = (s: SeriesData) => (s.secondary ? 1 : 0);
  const pctTotals: number[] = [];
  if (percent) for (let i = 0; i < n; i++) pctTotals[i] = series.reduce((a, s) => a + Math.abs(s.values[i] ?? 0), 0) || 1;
  if (waterfall) waterfall.forEach((p) => { note(0, p.base); note(0, p.base + (p.total ? 0 : p.val)); note(0, p.total ? p.val : p.base + p.val); note(0, 0); });
  else if (boxes) boxes.forEach((b) => { note(0, b.q.min); note(0, b.q.max); });
  else if (type === 'stock') series.forEach((s) => s.values.forEach((v) => v != null && note(0, v)));
  else {
    // Stacks per axis and per stack kind (column stacks separately from line stacks in combos).
    const stackKey = (t: ChartType) => (COLUMNISH.has(t) ? 'c' : AREAISH.has(t) ? 'a' : 'l');
    for (let i = 0; i < n; i++) {
      const pos: Record<string, number> = {}, neg: Record<string, number> = {};
      series.forEach((s, si) => {
        const v0 = s.values[i];
        if (v0 == null) return;
        const v = percent ? v0 / pctTotals[i] : v0;
        const t = sTypes[si];
        const k = axisOf(s) + stackKey(t);
        if (stacked && (STACKED.has(t) || PERCENT.has(t))) { if (v >= 0) pos[k] = (pos[k] || 0) + v; else neg[k] = (neg[k] || 0) + v; note(axisOf(s), v >= 0 ? pos[k] : neg[k]); }
        else note(axisOf(s), v);
      });
    }
  }
  const hasSecondary = series.some((s) => s.secondary);
  const anyColumn = sTypes.some((t) => COLUMNISH.has(t) || AREAISH.has(t)) || !!waterfall;
  const mkScale = (ax: number) => {
    let { min, max } = ext[ax];
    if (min === Infinity) { min = 0; max = 1; }
    const log = ax === 0 && !!chart.logScale && min > 0;
    const zero = anyColumn || (min >= 0 ? min < max * 5 / 6 : max > 0) || percent;
    const s = niceScale(min, max, Math.max(2, Math.floor((horizontal ? w : h) / 55)), { includeZero: zero && !log, log, fixedMin: ax === 0 ? chart.yMin : undefined, fixedMax: ax === 0 ? chart.yMax : undefined });
    if (percent && ax === 0 && chart.yMax == null) { s.max = Math.max(1, Math.min(s.max, 1)); s.ticks = s.ticks.filter((t) => t <= s.max + 1e-9); }
    return s;
  };
  const yS = mkScale(0);
  const y2S = hasSecondary ? (type === 'pareto' ? niceScale(0, 1, 5, { fixedMin: 0, fixedMax: 1 }) : mkScale(1)) : null;
  const yFmt = percent ? '0%' : valFmt;
  const y2Fmt = secondaryFmt || series.find((s) => s.secondary)?.fmt;

  const legendItems: LegendItem[] = waterfall
    ? [{ label: 'Increase', color: series[0].color, kind: 'rect' }, { label: 'Decrease', color: '#e07b28', kind: 'rect' }, { label: 'Total', color: '#7a8a99', kind: 'rect' }]
    : series.map((s, i) => ({ label: s.name, color: s.color, kind: LINEISH.has(sTypes[i]) || type === 'stock' ? 'line' as const : 'rect' as const }));
  const legendPos = chart.legend ?? (series.length > 1 || waterfall ? 'bottom' : 'none');
  const fr = frame(chart, w, h, title, type === 'histogram' ? [] : legendItems, legendPos);
  const svg = fr.svg;
  let { x: px, y: py, w: pw, h: ph } = fr.plot;

  // Axis titles
  const tSize = 11;
  if (chart.yTitle) { if (horizontal) { svg.push(text(px + pw / 2, py + ph, clip(chart.yTitle, pw, tSize), { size: tSize, anchor: 'middle', bold: true })); ph -= tSize + 6; } else { svg.push(text(px + tSize, py + ph / 2, clip(chart.yTitle, ph, tSize), { size: tSize, anchor: 'middle', rotate: -90, bold: true })); px += tSize + 6; pw -= tSize + 6; } }
  if (chart.xTitle) { if (horizontal) { svg.push(text(px + tSize, py + ph / 2, clip(chart.xTitle, ph, tSize), { size: tSize, anchor: 'middle', rotate: -90, bold: true })); px += tSize + 6; pw -= tSize + 6; } else { svg.push(text(px + pw / 2, py + ph, clip(chart.xTitle, pw, tSize), { size: tSize, anchor: 'middle', bold: true })); ph -= tSize + 6; } }

  // Category label space
  const labSize = 10;
  const catLabels = cats.map((c) => c);
  const maxCatW = Math.max(...catLabels.map((c) => textW(c, labSize)), 0);
  const vW = valueAxisWidth(yS, yFmt, d1904);
  const v2W = y2S ? valueAxisWidth(y2S, y2Fmt, d1904) : 0;
  let rotateCats = false;
  if (horizontal) {
    const cw = Math.min(maxCatW + 6, pw * 0.35);
    px += cw; pw -= cw;
    ph -= labSize + 8; // value labels below
  } else {
    px += vW; pw -= vW + v2W;
    const slot = pw / n;
    rotateCats = maxCatW > slot - 4 && n > 1;
    const labH = rotateCats ? Math.min(ph * 0.35, maxCatW * 0.72 + 6) : labSize + 8;
    ph -= labH;
  }
  if (pw < 10 || ph < 10) return svgWrap(w, h, svg);

  // Position helpers: band (category) and value.
  const band = (horizontal ? ph : pw) / n;
  const catPos = (i: number) => (horizontal ? py + band * (i + 0.5) : px + band * (i + 0.5));
  const valPos = (s: Scale, v: number) => (horizontal ? scalePos(s, v, px, px + pw) : scalePos(s, v, py + ph, py));

  // Gridlines + value axis labels
  const showGrid = chart.gridlines !== false;
  for (const t of yS.ticks) {
    const p = valPos(yS, t);
    if (horizontal) {
      if (showGrid) svg.push(`<line x1="${F(p)}" y1="${F(py)}" x2="${F(p)}" y2="${F(py + ph)}" stroke="${GRID}"/>`);
      svg.push(text(p, py + ph + labSize + 4, fmtTick(t, yFmt, yS.step, d1904), { size: labSize, anchor: 'middle' }));
    } else {
      if (showGrid) svg.push(`<line x1="${F(px)}" y1="${F(p)}" x2="${F(px + pw)}" y2="${F(p)}" stroke="${GRID}"/>`);
      svg.push(text(px - 5, p, fmtTick(t, yFmt, yS.step, d1904), { size: labSize, anchor: 'end', baseline: 'middle' }));
    }
  }
  if (y2S) for (const t of y2S.ticks) svg.push(text(px + pw + 5, valPos(y2S, t), fmtTick(t, y2Fmt, y2S.step, d1904), { size: labSize, baseline: 'middle' }));
  // Category axis line at zero (or the axis minimum)
  const base0 = Math.min(Math.max(0, yS.min), yS.max);
  const zp = valPos(yS, yS.log ? yS.min : base0);
  svg.push(horizontal ? `<line x1="${F(zp)}" y1="${F(py)}" x2="${F(zp)}" y2="${F(py + ph)}" stroke="${AXIS}"/>` : `<line x1="${F(px)}" y1="${F(zp)}" x2="${F(px + pw)}" y2="${F(zp)}" stroke="${AXIS}"/>`);
  // Category labels
  const labelStep = horizontal ? Math.max(1, Math.ceil((labSize + 2) / band)) : rotateCats ? Math.max(1, Math.ceil((labSize + 2) / band)) : 1;
  cats.forEach((c, i) => {
    if (i % labelStep) return;
    const p = catPos(i);
    if (horizontal) svg.push(text(px - 5, p, clip(c, Math.min(maxCatW + 6, w * 0.35), labSize), { size: labSize, anchor: 'end', baseline: 'middle' }));
    else if (rotateCats) svg.push(text(p, py + ph + 6, clip(c, ph * 0.5 + 40, labSize), { size: labSize, anchor: 'end', rotate: -45, baseline: 'hanging' }));
    else svg.push(text(p, py + ph + labSize + 4, clip(c, band, labSize), { size: labSize, anchor: 'middle' }));
  });

  const labels: string[] = [];
  const dl = !!chart.dataLabels;
  const lab = (x: number, y: number, v: number, fmt: string | undefined, anchor: 'start' | 'middle' | 'end' = 'middle', color = TXT) => { if (dl) labels.push(text(x, y, fmtTick(v, fmt, Math.abs(v) >= 10 || !fmt ? 1 : 0.01, d1904), { size: 9.5, anchor, color, baseline: 'middle' })); };
  const rect = (x0: number, y0: number, x1: number, y1: number, fill: string, extra = '') => `<rect x="${F(Math.min(x0, x1))}" y="${F(Math.min(y0, y1))}" width="${F(Math.max(0.5, Math.abs(x1 - x0)))}" height="${F(Math.max(0.5, Math.abs(y1 - y0)))}" fill="${fill}"${extra}/>`;

  // ---- Special category charts
  if (waterfall) {
    const bw = band * 0.62;
    let prevEnd: number | null = null;
    waterfall.forEach((p, i) => {
      const c = catPos(i);
      const a = p.total ? 0 : p.base, b = p.total ? p.val : p.base + p.val;
      const color = p.total ? '#7a8a99' : p.val >= 0 ? series[0].color : '#e07b28';
      svg.push(horizontal ? rect(valPos(yS, a), c - bw / 2, valPos(yS, b), c + bw / 2, color) : rect(c - bw / 2, valPos(yS, a), c + bw / 2, valPos(yS, b), color));
      if (prevEnd != null && !horizontal) svg.push(`<line x1="${F(c - band + bw / 2)}" y1="${F(valPos(yS, prevEnd))}" x2="${F(c - bw / 2)}" y2="${F(valPos(yS, prevEnd))}" stroke="#a6a6a6" stroke-dasharray="2 2"/>`);
      prevEnd = b;
      lab(horizontal ? valPos(yS, b) + 4 : c, horizontal ? c : valPos(yS, Math.max(a, b)) - 7, p.total ? p.val : p.val, valFmt, horizontal ? 'start' : 'middle');
    });
    return svgWrap(w, h, [...svg, ...labels]);
  }
  if (boxes) {
    const per = series.length;
    const slot = (band * 0.7) / per;
    for (const b of boxes) {
      const s = series[b.si];
      const c = catPos(b.cat) - (band * 0.7) / 2 + slot * (b.si + 0.5);
      const bw = slot * 0.7;
      const Y = (v: number) => valPos(yS, v);
      svg.push(`<line x1="${F(c)}" y1="${F(Y(b.q.lowW))}" x2="${F(c)}" y2="${F(Y(b.q.q1))}" stroke="#595959"/><line x1="${F(c)}" y1="${F(Y(b.q.q3))}" x2="${F(c)}" y2="${F(Y(b.q.highW))}" stroke="#595959"/>`);
      svg.push(`<line x1="${F(c - bw / 4)}" y1="${F(Y(b.q.lowW))}" x2="${F(c + bw / 4)}" y2="${F(Y(b.q.lowW))}" stroke="#595959"/><line x1="${F(c - bw / 4)}" y1="${F(Y(b.q.highW))}" x2="${F(c + bw / 4)}" y2="${F(Y(b.q.highW))}" stroke="#595959"/>`);
      svg.push(rect(c - bw / 2, Y(b.q.q1), c + bw / 2, Y(b.q.q3), s.color, ' fill-opacity="0.85" stroke="#404040" stroke-width="0.6"'));
      svg.push(`<line x1="${F(c - bw / 2)}" y1="${F(Y(b.q.median))}" x2="${F(c + bw / 2)}" y2="${F(Y(b.q.median))}" stroke="#ffffff" stroke-width="1.5"/>`);
      svg.push(`<path d="M${F(c - 3)},${F(Y(b.q.mean) - 3)} l6,6 m0,-6 l-6,6" stroke="#262626" stroke-width="1.3"/>`);
      for (const o of b.q.outliers) svg.push(`<circle cx="${F(c)}" cy="${F(Y(o))}" r="2.5" fill="none" stroke="${s.color}"/>`);
      lab(c + bw / 2 + 3, Y(b.q.median), b.q.median, valFmt, 'start');
    }
    return svgWrap(w, h, [...svg, ...labels]);
  }
  if (type === 'stock') {
    const k = series.length >= 4 ? 1 : 0; // O,H,L,C or H,L,C
    const H = series[k], L = series[k + 1], C = series[k + 2], O = k ? series[0] : null;
    const tick = Math.min(8, band * 0.3);
    for (let i = 0; i < n; i++) {
      const hi = H?.values[i], lo = L?.values[i], cl = C?.values[i], op = O?.values[i];
      const c = catPos(i);
      if (hi != null && lo != null) svg.push(`<line x1="${F(c)}" y1="${F(valPos(yS, hi))}" x2="${F(c)}" y2="${F(valPos(yS, lo))}" stroke="#404040" stroke-width="1.2"/>`);
      if (cl != null) svg.push(`<line x1="${F(c)}" y1="${F(valPos(yS, cl))}" x2="${F(c + tick)}" y2="${F(valPos(yS, cl))}" stroke="${H?.color || '#2f6fb0'}" stroke-width="2"/>`);
      if (op != null) svg.push(`<line x1="${F(c - tick)}" y1="${F(valPos(yS, op))}" x2="${F(c)}" y2="${F(valPos(yS, op))}" stroke="${H?.color || '#2f6fb0'}" stroke-width="2"/>`);
      if (cl != null) lab(c + tick + 3, valPos(yS, cl), cl, valFmt, 'start');
    }
    return svgWrap(w, h, [...svg, ...labels]);
  }

  // ---- Generic series drawing
  const colSeries = series.map((_, i) => i).filter((i) => COLUMNISH.has(sTypes[i]));
  const colStacked = stacked && colSeries.some((i) => STACKED.has(sTypes[i]) || PERCENT.has(sTypes[i]));
  const groups = colStacked ? 1 : Math.max(1, colSeries.length);
  const groupW = band * (1 / (1 + gapRatio * (type === 'histogram' ? 0 : 1)));
  const barW = Math.max(1, (groupW / groups) * (groups > 1 ? 0.92 : 1));
  const posStack = new Map<string, number[]>(), negStack = new Map<string, number[]>();
  const stackArr = (m: Map<string, number[]>, k: string) => { let a = m.get(k); if (!a) { a = new Array(n).fill(0); m.set(k, a); } return a; };
  const style = chart.style || 1;
  const outline = style === 3 ? ` stroke="#ffffff" stroke-width="1"` : style === 6 ? ` stroke="#404040" stroke-width="0.8"` : '';
  const defs: string[] = [];
  const fillFor = (color: string, id: string) => {
    if (style !== 4) return color;
    defs.push(`<linearGradient id="${id}" x1="0" y1="0" x2="${horizontal ? 1 : 0}" y2="${horizontal ? 0 : 1}"><stop offset="0" stop-color="${mixColor(color, '#ffffff', 0.35)}"/><stop offset="1" stop-color="${color}"/></linearGradient>`);
    return `url(#${id})`;
  };
  const gid = `c${String(chart.id || 'x').replace(/[^A-Za-z0-9]/g, '')}`;

  // Areas first (behind), then columns, then lines.
  const order = series.map((_, i) => i).sort((a, b) => (AREAISH.has(sTypes[a]) ? 0 : COLUMNISH.has(sTypes[a]) ? 1 : 2) - (AREAISH.has(sTypes[b]) ? 0 : COLUMNISH.has(sTypes[b]) ? 1 : 2));
  const areaBase = new Map<string, number[]>();
  for (const si of order) {
    const s = series[si];
    const t = sTypes[si];
    const sc = s.secondary && y2S ? y2S : yS;
    const isStack = STACKED.has(t) || PERCENT.has(t);
    const vOf = (i: number) => { const v = s.values[i]; return v == null ? null : PERCENT.has(t) ? v / pctTotals[i] : v; };
    const fmt = PERCENT.has(t) ? '0%' : s.secondary ? y2Fmt : s.fmt;
    if (COLUMNISH.has(t)) {
      const gi = colStacked ? 0 : colSeries.indexOf(si);
      const key = (s.secondary ? '2' : '1') + 'c';
      const fill = fillFor(s.color, `${gid}g${si}`);
      for (let i = 0; i < n; i++) {
        const v = vOf(i);
        if (v == null) continue;
        let a = sc.log ? sc.min : 0, b = v;
        if (isStack && colStacked) { const st = v >= 0 ? stackArr(posStack, key) : stackArr(negStack, key); a = st[i]; b = st[i] + v; st[i] = b; }
        const c0 = catPos(i) - groupW / 2 + (groupW / groups) * gi + ((groupW / groups) - barW) / 2;
        const A = valPos(sc, a), B = valPos(sc, b);
        svg.push(horizontal ? rect(A, c0, B, c0 + barW, fill, outline) : rect(c0, A, c0 + barW, B, fill, outline + (style === 2 ? ' rx="2"' : '')));
        if (isStack && colStacked) lab(horizontal ? (A + B) / 2 : c0 + barW / 2, horizontal ? c0 + barW / 2 : (A + B) / 2, v, fmt, 'middle', '#ffffff');
        else lab(horizontal ? B + (v >= 0 ? 4 : -4) : c0 + barW / 2, horizontal ? c0 + barW / 2 : B + (v >= 0 ? -7 : 7), v, fmt, horizontal ? (v >= 0 ? 'start' : 'end') : 'middle');
      }
    } else if (AREAISH.has(t)) {
      const key = (s.secondary ? '2' : '1') + 'a';
      const prev = areaBase.get(key) || new Array(n).fill(0);
      const tops: number[] = [];
      for (let i = 0; i < n; i++) tops[i] = (isStack ? prev[i] : 0) + (vOf(i) ?? 0);
      const pts = tops.map((v, i) => `${F(catPos(i))},${F(valPos(sc, v))}`);
      const basePts = (isStack ? prev : new Array(n).fill(sc.log ? sc.min : 0)).map((v, i) => `${F(catPos(i))},${F(valPos(sc, v))}`).reverse();
      svg.push(`<path d="M${pts.join(' L')} L${basePts.join(' L')} Z" fill="${fillFor(s.color, `${gid}g${si}`)}" fill-opacity="${isStack ? 0.9 : 0.6}"${outline}/>`);
      if (isStack) areaBase.set(key, tops);
      for (let i = 0; i < n; i++) { const v = vOf(i); if (v != null) lab(catPos(i), valPos(sc, tops[i]) - 7, v, fmt); }
    } else {
      // Lines (line, lineMarkers, stackedLine and combo lines)
      const key = (s.secondary ? '2' : '1') + 'l';
      const st = isStack ? stackArr(posStack, key) : null;
      const pts: ({ x: number; y: number; v: number } | null)[] = [];
      for (let i = 0; i < n; i++) {
        const v = vOf(i);
        if (v == null) { pts.push(null); continue; }
        const y = st ? (st[i] += v) : v;
        pts.push({ x: catPos(i), y: valPos(sc, y), v });
      }
      let d = '', pen = false;
      for (const p of pts) { if (!p) { pen = false; continue; } d += `${pen ? 'L' : 'M'}${F(p.x)},${F(p.y)} `; pen = true; }
      if (d) svg.push(`<path d="${d.trim()}" fill="none" stroke="${s.color}" stroke-width="${style === 5 ? 3 : 2.25}" stroke-linejoin="round" stroke-linecap="round"/>`);
      const markers = t === 'lineMarkers' || n <= 1;
      for (const p of pts) {
        if (!p) continue;
        if (markers) svg.push(`<circle cx="${F(p.x)}" cy="${F(p.y)}" r="3.5" fill="${s.color}" stroke="#ffffff" stroke-width="1"/>`);
        lab(p.x, p.y - 9, p.v, fmt);
      }
    }
    // Trendline (category charts use x = 1..n)
    if (chart.trendline && !isStack && type !== 'pareto') {
      const xs: number[] = [], ys: number[] = [];
      s.values.forEach((v, i) => { if (v != null) { xs.push(i + 1); ys.push(v); } });
      const tp = trendline(chart.trendline, xs, ys);
      if (tp.length > 1) svg.push(`<path d="M${tp.map((p) => { const ci = p.x - 1; const c = horizontal ? py + band * (ci + 0.5) : px + band * (ci + 0.5); const vv = valPos(sc, Math.max(Math.min(p.y, sc.max), sc.min)); return horizontal ? `${F(vv)},${F(c)}` : `${F(c)},${F(vv)}`; }).join(' L')}" fill="none" stroke="${mixColor(s.color, '#000000', 0.25)}" stroke-width="1.6" stroke-dasharray="5 4"/>`);
    }
  }
  if (defs.length) svg.unshift(`<defs>${defs.join('')}</defs>`);
  return svgWrap(w, h, [...svg, ...labels]);
}

// ---- Scatter and bubble ------------------------------------------------------------------------

function xyChart(chart: ChartDef, data: ChartData, w: number, h: number, title: string, d1904: boolean): string {
  const bubble = chart.type === 'bubble';
  const series = data.series.map((s) => {
    const xs = s.xs && s.xs.some((x) => x != null) ? s.xs : s.values.map((_, i) => i + 1);
    const pts = s.values.map((y, i) => ({ x: xs[i], y, size: bubble ? (s.sizes?.[i] ?? Math.abs(y ?? 0)) : 0 })).filter((p): p is { x: number; y: number; size: number } => p.x != null && p.y != null);
    return { ...s, pts };
  });
  const all = series.flatMap((s) => s.pts);
  if (!all.length) return message(w, h, title, 'No numeric data to plot.');
  const legendPos = chart.legend ?? (series.length > 1 ? 'bottom' : 'none');
  const fr = frame(chart, w, h, title, series.map((s) => ({ label: s.name, color: s.color, kind: 'marker' as const })), legendPos);
  const svg = fr.svg;
  let { x: px, y: py, w: pw, h: ph } = fr.plot;
  const xFmt = data.catFmt && isDateFormat(data.catFmt) ? data.catFmt : refFormat2(chart, data);
  const yFmt = series[0].fmt;
  const xs = all.map((p) => p.x), ys = all.map((p) => p.y);
  const xS = niceScale(Math.min(...xs), Math.max(...xs), Math.max(2, Math.floor(w / 80)), { includeZero: Math.min(...xs) >= 0 && Math.min(...xs) < Math.max(...xs) * 5 / 6 });
  const yS = niceScale(Math.min(...ys), Math.max(...ys), Math.max(2, Math.floor(h / 55)), { includeZero: Math.min(...ys) >= 0 && Math.min(...ys) < Math.max(...ys) * 5 / 6, log: !!chart.logScale && Math.min(...ys) > 0, fixedMin: chart.yMin, fixedMax: chart.yMax });
  const tSize = 11;
  if (chart.yTitle) { svg.push(text(px + tSize, py + ph / 2, clip(chart.yTitle, ph, tSize), { size: tSize, anchor: 'middle', rotate: -90, bold: true })); px += tSize + 6; pw -= tSize + 6; }
  if (chart.xTitle) { svg.push(text(px + pw / 2, py + ph, clip(chart.xTitle, pw, tSize), { size: tSize, anchor: 'middle', bold: true })); ph -= tSize + 6; }
  const vW = valueAxisWidth(yS, yFmt, d1904);
  px += vW; pw -= vW + 6; ph -= 18;
  if (pw < 10 || ph < 10) return svgWrap(w, h, svg);
  const X = (v: number) => scalePos(xS, v, px, px + pw), Y = (v: number) => scalePos(yS, v, py + ph, py);
  const grid = chart.gridlines !== false;
  for (const t of yS.ticks) { if (grid) svg.push(`<line x1="${F(px)}" y1="${F(Y(t))}" x2="${F(px + pw)}" y2="${F(Y(t))}" stroke="${GRID}"/>`); svg.push(text(px - 5, Y(t), fmtTick(t, yFmt, yS.step, d1904), { size: 10, anchor: 'end', baseline: 'middle' })); }
  for (const t of xS.ticks) { if (grid) svg.push(`<line x1="${F(X(t))}" y1="${F(py)}" x2="${F(X(t))}" y2="${F(py + ph)}" stroke="${GRID}"/>`); svg.push(text(X(t), py + ph + 14, fmtTick(t, xFmt, xS.step, d1904), { size: 10, anchor: 'middle' })); }
  svg.push(`<line x1="${F(px)}" y1="${F(py + ph)}" x2="${F(px + pw)}" y2="${F(py + ph)}" stroke="${AXIS}"/><line x1="${F(px)}" y1="${F(py)}" x2="${F(px)}" y2="${F(py + ph)}" stroke="${AXIS}"/>`);
  const maxSize = bubble ? Math.max(...all.map((p) => Math.abs(p.size)), 1e-9) : 1;
  const rMax = Math.max(4, Math.min(pw, ph) * 0.09);
  const labels: string[] = [];
  for (const s of series) {
    if (chart.type === 'scatterLines' && s.pts.length > 1) {
      const sorted = s.pts;
      svg.push(`<path d="M${sorted.map((p) => `${F(X(p.x))},${F(Y(p.y))}`).join(' L')}" fill="none" stroke="${s.color}" stroke-width="2"/>`);
    }
    for (const p of s.pts) {
      if (bubble) svg.push(`<circle cx="${F(X(p.x))}" cy="${F(Y(p.y))}" r="${F(Math.max(2, rMax * Math.sqrt(Math.abs(p.size) / maxSize)))}" fill="${s.color}" fill-opacity="0.7" stroke="${mixColor(s.color, '#000000', 0.2)}"/>`);
      else svg.push(`<circle cx="${F(X(p.x))}" cy="${F(Y(p.y))}" r="3.5" fill="${s.color}"/>`);
      if (chart.dataLabels) labels.push(text(X(p.x) + 6, Y(p.y) - 6, fmtTick(p.y, s.fmt, yS.step, d1904), { size: 9.5 }));
    }
    if (chart.trendline) {
      const tp = trendline(chart.trendline, s.pts.map((p) => p.x), s.pts.map((p) => p.y));
      if (tp.length > 1) svg.push(`<path d="M${tp.map((p) => `${F(X(p.x))},${F(Y(Math.max(Math.min(p.y, yS.max), yS.min)))}`).join(' L')}" fill="none" stroke="${mixColor(s.color, '#000000', 0.25)}" stroke-width="1.6" stroke-dasharray="5 4"/>`);
    }
  }
  return svgWrap(w, h, [...svg, ...labels]);
}

function refFormat2(_chart: ChartDef, data: ChartData): string | undefined { return data.catFmt; }

// ---- Pie and doughnut --------------------------------------------------------------------------

function pieChart(chart: ChartDef, data: ChartData, w: number, h: number, title: string, d1904: boolean): string {
  const s = data.series[0];
  const colors = s.values.map((_, i) => data.palette[i % data.palette.length]);
  const items = s.values.map((v, i) => ({ v: Math.max(0, v ?? 0), label: data.categories[i] ?? String(i + 1), color: colors[i] }));
  const total = items.reduce((a, it) => a + it.v, 0);
  const legendPos = chart.legend ?? 'right';
  const fr = frame(chart, w, h, title, items.map((it) => ({ label: it.label, color: it.color, kind: 'rect' as const })), legendPos);
  const svg = fr.svg;
  const { x, y, w: pw, h: ph } = fr.plot;
  if (!total) { svg.push(text(x + pw / 2, y + ph / 2, 'All values are zero', { anchor: 'middle', color: '#7f7f7f' })); return svgWrap(w, h, svg); }
  const cx = x + pw / 2, cy = y + ph / 2, r = Math.max(5, Math.min(pw, ph) / 2 - (chart.dataLabels ? 16 : 4));
  const ri = chart.type === 'doughnut' ? r * 0.55 : 0;
  let a = -Math.PI / 2;
  const labels: string[] = [];
  for (const it of items) {
    if (!it.v) continue;
    const da = (it.v / total) * Math.PI * 2;
    const a2 = a + da;
    const large = da > Math.PI ? 1 : 0;
    const p = (ang: number, rr: number) => `${F(cx + rr * Math.cos(ang))},${F(cy + rr * Math.sin(ang))}`;
    let d: string;
    if (da >= Math.PI * 2 - 1e-9) d = ri ? `M${p(0, r)} A${F(r)},${F(r)} 0 1 1 ${p(Math.PI, r)} A${F(r)},${F(r)} 0 1 1 ${p(0, r)} Z M${p(0, ri)} A${F(ri)},${F(ri)} 0 1 0 ${p(Math.PI, ri)} A${F(ri)},${F(ri)} 0 1 0 ${p(0, ri)} Z` : `M${p(0, r)} A${F(r)},${F(r)} 0 1 1 ${p(Math.PI, r)} A${F(r)},${F(r)} 0 1 1 ${p(0, r)} Z`;
    else if (ri) d = `M${p(a, r)} A${F(r)},${F(r)} 0 ${large} 1 ${p(a2, r)} L${p(a2, ri)} A${F(ri)},${F(ri)} 0 ${large} 0 ${p(a, ri)} Z`;
    else d = `M${F(cx)},${F(cy)} L${p(a, r)} A${F(r)},${F(r)} 0 ${large} 1 ${p(a2, r)} Z`;
    svg.push(`<path d="${d}" fill="${it.color}" stroke="#ffffff" stroke-width="1.2" fill-rule="evenodd"/>`);
    if (chart.dataLabels) {
      const mid = (a + a2) / 2;
      const rr = ri ? (r + ri) / 2 : r * 0.68;
      const pct = it.v / total;
      labels.push(text(cx + rr * Math.cos(mid), cy + rr * Math.sin(mid), pct >= 0.03 ? formatValue(pct, '0%').text : '', { size: 10, anchor: 'middle', color: '#ffffff', baseline: 'middle', bold: true }));
    }
    a = a2;
  }
  void d1904;
  return svgWrap(w, h, [...svg, ...labels]);
}

// ---- Radar -----------------------------------------------------------------------------------

function radarChart(chart: ChartDef, data: ChartData, w: number, h: number, title: string, d1904: boolean): string {
  const n = Math.max(3, data.categories.length);
  const legendPos = chart.legend ?? (data.series.length > 1 ? 'bottom' : 'none');
  const fr = frame(chart, w, h, title, data.series.map((s) => ({ label: s.name, color: s.color, kind: chart.type === 'filledRadar' ? 'rect' as const : 'line' as const })), legendPos);
  const svg = fr.svg;
  const { x, y, w: pw, h: ph } = fr.plot;
  const vals = data.series.flatMap((s) => s.values).filter((v): v is number => v != null);
  const S = niceScale(Math.min(0, ...vals), Math.max(...vals), 4, { includeZero: true, fixedMin: chart.yMin, fixedMax: chart.yMax });
  const cx = x + pw / 2, cy = y + ph / 2, R = Math.max(5, Math.min(pw, ph) / 2 - 22);
  const ang = (i: number) => -Math.PI / 2 + (i * 2 * Math.PI) / n;
  const rad = (v: number) => ((v - S.min) / (S.max - S.min || 1)) * R;
  for (const t of S.ticks) {
    const rr = rad(t);
    svg.push(`<path d="M${Array.from({ length: n }, (_, i) => `${F(cx + rr * Math.cos(ang(i)))},${F(cy + rr * Math.sin(ang(i)))}`).join(' L')} Z" fill="none" stroke="${GRID}"/>`);
    svg.push(text(cx + 3, cy - rr, fmtTick(t, data.series[0].fmt, S.step, d1904), { size: 9, color: '#7f7f7f', baseline: 'middle' }));
  }
  for (let i = 0; i < n; i++) {
    const ax = cx + R * Math.cos(ang(i)), ay = cy + R * Math.sin(ang(i));
    svg.push(`<line x1="${F(cx)}" y1="${F(cy)}" x2="${F(ax)}" y2="${F(ay)}" stroke="${GRID}"/>`);
    const lx = cx + (R + 10) * Math.cos(ang(i)), ly = cy + (R + 10) * Math.sin(ang(i));
    const anchor = Math.abs(Math.cos(ang(i))) < 0.2 ? 'middle' : Math.cos(ang(i)) > 0 ? 'start' : 'end';
    svg.push(text(lx, ly, clip(data.categories[i] ?? '', 90, 10), { size: 10, anchor, baseline: 'middle' }));
  }
  for (const s of data.series) {
    const pts = Array.from({ length: n }, (_, i) => { const v = s.values[i] ?? S.min; return `${F(cx + rad(v) * Math.cos(ang(i)))},${F(cy + rad(v) * Math.sin(ang(i)))}`; });
    svg.push(`<path d="M${pts.join(' L')} Z" fill="${chart.type === 'filledRadar' ? s.color : 'none'}" fill-opacity="0.35" stroke="${s.color}" stroke-width="2"/>`);
    if (chart.dataLabels) s.values.forEach((v, i) => { if (v != null) svg.push(text(cx + (rad(v) + 8) * Math.cos(ang(i)), cy + (rad(v) + 8) * Math.sin(ang(i)), fmtTick(v, s.fmt, S.step, d1904), { size: 9, anchor: 'middle', baseline: 'middle' })); });
  }
  return svgWrap(w, h, svg);
}

// ---- Funnel ------------------------------------------------------------------------------------

function funnelChart(chart: ChartDef, data: ChartData, w: number, h: number, title: string, d1904: boolean): string {
  const s = data.series[0];
  const fr = frame(chart, w, h, title, [], 'none');
  const svg = fr.svg;
  const { x, y, w: pw, h: ph } = fr.plot;
  const items = s.values.map((v, i) => ({ v: v ?? 0, label: data.categories[i] ?? '' })).filter((it) => it.v > 0);
  if (!items.length) return message(w, h, title, 'No positive values to show.');
  const max = Math.max(...items.map((it) => it.v));
  const labW = Math.min(pw * 0.3, Math.max(...items.map((it) => textW(it.label, 10))) + 8);
  const area = pw - labW;
  const bh = ph / items.length;
  items.forEach((it, i) => {
    const bw = (it.v / max) * area;
    const yy = y + i * bh;
    const xx = x + labW + (area - bw) / 2;
    svg.push(`<rect x="${F(xx)}" y="${F(yy + bh * 0.08)}" width="${F(Math.max(1, bw))}" height="${F(bh * 0.84)}" fill="${s.color}"/>`);
    svg.push(text(x + labW - 6, yy + bh / 2, clip(it.label, labW - 6, 10), { size: 10, anchor: 'end', baseline: 'middle' }));
    svg.push(text(x + labW + area / 2, yy + bh / 2, fmtTick(it.v, s.fmt, 1, d1904), { size: 10, anchor: 'middle', baseline: 'middle', color: bw > 40 ? '#ffffff' : TXT }));
  });
  return svgWrap(w, h, svg);
}

// ---- Treemap and sunburst ----------------------------------------------------------------------

interface HNode { name: string; value: number; kids: HNode[]; color?: string }

function hierarchy(data: ChartData): HNode {
  const root: HNode = { name: '', value: 0, kids: [] };
  const s = data.series[0];
  data.catLevels.forEach((path, i) => {
    const v = s.values[i];
    if (v == null || v <= 0) return;
    let node = root;
    const parts = path.length ? path : [data.categories[i] ?? String(i + 1)];
    for (const p of parts) {
      let k = node.kids.find((x) => x.name === p);
      if (!k) { k = { name: p, value: 0, kids: [] }; node.kids.push(k); }
      node = k;
    }
    node.value += v;
  });
  const sum = (n: HNode): number => (n.kids.length ? (n.value = n.kids.reduce((a, k) => a + sum(k), 0) + (n.kids.length ? 0 : n.value)) : n.value);
  sum(root);
  return root;
}

function treemapChart(chart: ChartDef, data: ChartData, w: number, h: number, title: string, d1904: boolean): string {
  const root = hierarchy(data);
  const pal = data.palette;
  root.kids.forEach((k, i) => { k.color = pal[i]; });
  const fr = frame(chart, w, h, title, root.kids.length > 1 && data.catLevels.some((l) => l.length > 1) ? root.kids.map((k) => ({ label: k.name, color: k.color!, kind: 'rect' as const })) : [], chart.legend ?? 'top');
  const svg = fr.svg;
  const fmt = data.series[0].fmt;
  const draw = (nodes: HNode[], box: Rect, depth: number, inherit?: string) => {
    const rects = squarify(nodes.map((n) => n.value), box);
    nodes.forEach((n, i) => {
      const r = rects[i];
      if (r.w <= 0 || r.h <= 0) return;
      const color = n.color || inherit || pal[i % pal.length];
      if (n.kids.length) {
        const head = depth === 0 && r.h > 40 ? 15 : 0;
        if (head) { svg.push(`<rect x="${F(r.x)}" y="${F(r.y)}" width="${F(r.w)}" height="${F(r.h)}" fill="${color}" stroke="#ffffff"/>`); svg.push(text(r.x + 4, r.y + 11, clip(n.name, r.w - 8, 10.5), { size: 10.5, color: '#ffffff', bold: true })); }
        draw(n.kids, { x: r.x, y: r.y + head, w: r.w, h: r.h - head }, depth + 1, color);
        return;
      }
      const shade = n.color ? color : mixColor(color, '#ffffff', ((i % 4) * 0.1));
      svg.push(`<rect x="${F(r.x)}" y="${F(r.y)}" width="${F(r.w)}" height="${F(r.h)}" fill="${shade}" stroke="#ffffff" stroke-width="1.5"/>`);
      if (r.w > 26 && r.h > 14) {
        svg.push(text(r.x + 4, r.y + 13, clip(n.name, r.w - 8, 10.5), { size: 10.5, color: '#ffffff' }));
        if (chart.dataLabels !== false && r.h > 28) svg.push(text(r.x + 4, r.y + 26, clip(fmtTick(n.value, fmt, 1, d1904), r.w - 8, 9.5), { size: 9.5, color: '#ffffff' }));
      }
    });
  };
  draw(root.kids, fr.plot, 0);
  return svgWrap(w, h, svg);
}

function sunburstChart(chart: ChartDef, data: ChartData, w: number, h: number, title: string, d1904: boolean): string {
  const root = hierarchy(data);
  const pal = data.palette;
  const fr = frame(chart, w, h, title, [], 'none');
  const svg = fr.svg;
  const { x, y, w: pw, h: ph } = fr.plot;
  const depthOf = (n: HNode): number => (n.kids.length ? 1 + Math.max(...n.kids.map(depthOf)) : 0);
  const levels = Math.max(1, depthOf(root));
  const cx = x + pw / 2, cy = y + ph / 2, R = Math.max(5, Math.min(pw, ph) / 2 - 2);
  const r0 = R * 0.18, ring = (R - r0) / levels;
  const total = root.value || 1;
  const arc = (a1: number, a2: number, ri: number, ro: number) => {
    const p = (ang: number, rr: number) => `${F(cx + rr * Math.cos(ang))},${F(cy + rr * Math.sin(ang))}`;
    if (a2 - a1 >= Math.PI * 2 - 1e-9) return `M${p(0, ro)} A${F(ro)},${F(ro)} 0 1 1 ${p(Math.PI, ro)} A${F(ro)},${F(ro)} 0 1 1 ${p(0, ro)} Z M${p(0, ri)} A${F(ri)},${F(ri)} 0 1 0 ${p(Math.PI, ri)} A${F(ri)},${F(ri)} 0 1 0 ${p(0, ri)} Z`;
    const large = a2 - a1 > Math.PI ? 1 : 0;
    return `M${p(a1, ro)} A${F(ro)},${F(ro)} 0 ${large} 1 ${p(a2, ro)} L${p(a2, ri)} A${F(ri)},${F(ri)} 0 ${large} 0 ${p(a1, ri)} Z`;
  };
  const walk = (nodes: HNode[], a0: number, depth: number, base?: string) => {
    let a = a0;
    nodes.forEach((n, i) => {
      const da = (n.value / total) * Math.PI * 2;
      const color = depth === 0 ? pal[i % pal.length] : mixColor(base!, '#ffffff', Math.min(0.6, 0.15 * depth + (i % 3) * 0.07));
      const ri = r0 + depth * ring, ro = ri + ring;
      svg.push(`<path d="${arc(a, a + da, ri, ro)}" fill="${color}" stroke="#ffffff" stroke-width="1.2" fill-rule="evenodd"/>`);
      const mid = a + da / 2, rr = (ri + ro) / 2;
      if (da * rr > 24) svg.push(text(cx + rr * Math.cos(mid), cy + rr * Math.sin(mid), clip(n.name, Math.min(ring - 4, da * rr), 9.5), { size: 9.5, anchor: 'middle', baseline: 'middle', color: depth === 0 ? '#ffffff' : '#262626' }));
      if (n.kids.length) walk(n.kids, a, depth + 1, color);
      a += da;
    });
  };
  walk(root.kids, -Math.PI / 2, 0);
  void d1904;
  return svgWrap(w, h, svg);
}
