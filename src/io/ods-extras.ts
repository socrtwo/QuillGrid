/**
 * OpenDocument spreadsheet features beyond cells and styles: validity rules
 * (table:content-validations), conditional formats (LibreOffice's calcext:conditional-formats
 * and the ODF style:map conditions of cell styles) and charts (embedded chart objects).
 * Used by the .ods/.fods reader (ods.ts) and through it by the OpenOffice.org 1.x reader.
 */
import { newId, type Sheet } from '../model/workbook';
import type { CellStyle, CFRule, ChartDef, ChartSeries, ChartType, DataValidation } from '../model/types';
import { keyRow, keyCol, cellKey, MAX_COLS, type Range } from '../model/address';
import { kids, kid, desc } from './index';

const NS = {
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  chart: 'urn:oasis:names:tc:opendocument:xmlns:chart:1.0',
  calcext: 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0',
  svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink: 'http://www.w3.org/1999/xlink',
} as const;

function a(el: Element | null | undefined, ns: keyof typeof NS, local: string): string | null {
  if (!el) return null;
  return el.getAttributeNS(NS[ns], local) ?? el.getAttribute(`${ns}:${local}`);
}

/** Splits "f(a;b)" style argument lists at top-level ',' or ';'. */
export function splitArgs(s: string): string[] {
  const out: string[] = [];
  let depth = 0, cur = '', q = false;
  for (const ch of s) {
    if (ch === '"') q = !q;
    if (!q && (ch === '(' || ch === '[' || ch === '{')) depth++;
    if (!q && (ch === ')' || ch === ']' || ch === '}')) depth--;
    if (!q && depth === 0 && (ch === ',' || ch === ';')) { out.push(cur.trim()); cur = ''; continue; }
    cur += ch;
  }
  if (cur.trim() || out.length) out.push(cur.trim());
  return out;
}

/** Cell keys -> covering rectangles (row runs merged downwards). */
export function cellsToRanges(keys: number[]): Range[] {
  const sorted = [...new Set(keys)].sort((x, y) => x - y);
  const runs: Range[] = [];
  for (const k of sorted) {
    const r = keyRow(k), c = keyCol(k);
    const last = runs[runs.length - 1];
    if (last && last.r1 === r && last.c2 === c - 1) last.c2 = c;
    else runs.push({ r1: r, c1: c, r2: r, c2: c });
  }
  const out: Range[] = [];
  for (const run of runs) {
    const prev = out.find((o) => o.c1 === run.c1 && o.c2 === run.c2 && o.r2 === run.r1 - 1);
    if (prev) prev.r2 = run.r2; else out.push({ ...run });
  }
  return out;
}

// --------------------------------------------------------------------------------------------
// Cell notes collected while reading rows

const notes = new WeakMap<Sheet, { validation: Map<string, number[]>; styleMap: Map<string, number[]> }>();
function noteOf(sheet: Sheet) { let n = notes.get(sheet); if (!n) { n = { validation: new Map(), styleMap: new Map() }; notes.set(sheet, n); } return n; }

/** Records that cells (r, c..c+rep-1) use a validity rule. */
export function noteValidation(sheet: Sheet, name: string, r: number, c: number, rep: number): void {
  const m = noteOf(sheet).validation;
  let list = m.get(name);
  if (!list) { list = []; m.set(name, list); }
  for (let k = 0; k < Math.min(rep, 1024) && c + k < MAX_COLS; k++) list.push(cellKey(r, c + k));
}

/** Records that cells use a cell style that carries style:map conditions. */
export function noteStyleMap(sheet: Sheet, style: string, r: number, c: number, rep: number): void {
  const m = noteOf(sheet).styleMap;
  let list = m.get(style);
  if (!list) { list = []; m.set(style, list); }
  for (let k = 0; k < Math.min(rep, 1024) && c + k < MAX_COLS; k++) list.push(cellKey(r, c + k));
}

// --------------------------------------------------------------------------------------------
// Validity

type ValidityRule = Omit<DataValidation, 'id' | 'ranges'>;

const OPS: Record<string, DataValidation['operator']> = { '=': 'equal', '!=': 'notEqual', '<>': 'notEqual', '>': 'greaterThan', '<': 'lessThan', '>=': 'greaterThanOrEqual', '<=': 'lessThanOrEqual' };

/** Parses an ODF validity condition into a rule. conv turns ODF expressions into formulas. */
export function parseValidity(cond: string, conv: (f: string) => string): ValidityRule | null {
  let s = cond.trim().replace(/^(of|oooc|ooow):/, '');
  const v: ValidityRule = { type: 'any' };
  const typeM = /^cell-content-is-(whole-number|decimal-number|date|time)\(\)\s*and\s*/.exec(s);
  if (typeM) {
    v.type = ({ 'whole-number': 'whole', 'decimal-number': 'decimal', date: 'date', time: 'time' } as const)[typeM[1] as 'date'];
    s = s.slice(typeM[0].length).replace(/^(of|oooc):/, '');
  }
  let m: RegExpExecArray | null;
  if ((m = /^cell-content-is-in-list\((.*)\)$/s.exec(s))) {
    v.type = 'list';
    const args = splitArgs(m[1]);
    if (args.length === 1 && !/^"/.test(args[0])) v.formula1 = conv(args[0]);
    else v.formula1 = '"' + args.map((x) => x.replace(/^"|"$/g, '').replace(/""/g, '"')).join(',') + '"';
    return v;
  }
  if ((m = /^is-true-formula\((.*)\)$/s.exec(s))) { v.type = 'custom'; v.formula1 = conv(m[1]); return v; }
  if ((m = /^cell-content(-text-length)?-is-(not-)?between\((.*)\)$/s.exec(s))) {
    if (m[1]) v.type = 'textLength'; else if (v.type === 'any') v.type = 'decimal';
    const args = splitArgs(m[3]);
    v.operator = m[2] ? 'notBetween' : 'between';
    v.formula1 = conv(args[0] || ''); v.formula2 = conv(args[1] || '');
    return v;
  }
  if ((m = /^cell-content(-text-length)?\(\)\s*(<=|>=|!=|<>|=|<|>)\s*(.*)$/s.exec(s))) {
    if (m[1]) v.type = 'textLength'; else if (v.type === 'any') v.type = 'decimal';
    v.operator = OPS[m[2]];
    v.formula1 = conv(m[3]);
    return v;
  }
  return v.type === 'any' ? null : v;
}

export function readValidations(spreadsheet: Element | null, conv: (f: string) => string): Map<string, ValidityRule> {
  const out = new Map<string, ValidityRule>();
  for (const cv of desc(spreadsheet, 'content-validation')) {
    const name = a(cv, 'table', 'name');
    if (!name) continue;
    const rule = parseValidity(a(cv, 'table', 'condition') || '', conv) || { type: 'any' as const };
    rule.allowBlank = a(cv, 'table', 'allow-empty-cell') !== 'false';
    if (rule.type === 'list') rule.showDropdown = a(cv, 'table', 'display-list') !== 'no';
    const help = kid(cv, 'help-message');
    if (help && a(help, 'table', 'display') !== 'false') {
      if (a(help, 'table', 'title')) rule.inputTitle = a(help, 'table', 'title')!;
      const t = kids(help, 'p').map((p) => p.textContent || '').join('\n');
      if (t) rule.inputMessage = t;
    }
    const err = kid(cv, 'error-message');
    if (err) {
      const mt = a(err, 'table', 'message-type');
      rule.errorStyle = mt === 'warning' ? 'warning' : mt === 'information' ? 'information' : 'stop';
      if (a(err, 'table', 'title')) rule.errorTitle = a(err, 'table', 'title')!;
      const t = kids(err, 'p').map((p) => p.textContent || '').join('\n');
      if (t) rule.errorMessage = t;
    }
    out.set(name, rule);
  }
  return out;
}

/** Adds the validity rules noted for the sheet's cells. */
export function applyValidations(sheet: Sheet, rules: Map<string, ValidityRule>): void {
  const n = notes.get(sheet);
  if (!n) return;
  for (const [name, keys] of n.validation) {
    const rule = rules.get(name);
    if (!rule || !keys.length) continue;
    sheet.validations.push({ id: newId('dv'), ranges: cellsToRanges(keys), ...rule });
  }
}

// --------------------------------------------------------------------------------------------
// Conditional formats

const CF_OPS: Record<string, CFRule['operator']> = { '=': 'equal', '!=': 'notEqual', '>': 'greaterThan', '<': 'lessThan', '>=': 'greaterThanOrEqual', '<=': 'lessThanOrEqual' };
const DATE_IS: Record<string, CFRule['timePeriod']> = { today: 'today', yesterday: 'yesterday', tomorrow: 'tomorrow', 'last-7-days': 'last7Days', 'this-week': 'thisWeek', 'last-week': 'lastWeek', 'next-week': 'nextWeek', 'this-month': 'thisMonth', 'last-month': 'lastMonth', 'next-month': 'nextMonth' };

/** Parses a calcext:value / style:map condition into rule fields (without ranges and style). */
export function parseCondition(value: string, conv: (f: string) => string): Partial<CFRule> | null {
  const s = value.trim().replace(/^cell-content\(\)\s*/, '');
  let m: RegExpExecArray | null;
  const str = (x: string) => x.trim().replace(/^"|"$/g, '').replace(/""/g, '"');
  if ((m = /^(?:cell-content-is-)?(not-)?between\((.*)\)$/s.exec(s))) { const args = splitArgs(m[2]); return { type: 'cellIs', operator: m[1] ? 'notBetween' : 'between', formulas: args.map(conv) }; }
  if ((m = /^(<=|>=|!=|=|<|>)\s*(.*)$/s.exec(s))) return { type: 'cellIs', operator: CF_OPS[m[1]], formulas: [conv(m[2])] };
  if ((m = /^(?:formula-is|is-true-formula)\((.*)\)$/s.exec(s))) return { type: 'expression', formulas: [conv(m[1])] };
  if ((m = /^begins-with\((.*)\)$/s.exec(s))) return { type: 'beginsWith', text: str(m[1]) };
  if ((m = /^ends-with\((.*)\)$/s.exec(s))) return { type: 'endsWith', text: str(m[1]) };
  if ((m = /^contains-text\((.*)\)$/s.exec(s))) return { type: 'containsText', text: str(m[1]) };
  if ((m = /^not-contains-text\((.*)\)$/s.exec(s))) return { type: 'notContainsText', text: str(m[1]) };
  if ((m = /^(top|bottom)-(elements|percent)\((.*)\)$/s.exec(s))) return { type: 'top10', rank: +m[3] || 10, ...(m[1] === 'bottom' ? { bottom: true } : {}), ...(m[2] === 'percent' ? { percent: true } : {}) };
  switch (s) {
    case 'duplicate': return { type: 'duplicateValues' };
    case 'unique': return { type: 'uniqueValues' };
    case 'above-average': return { type: 'aboveAverage', above: true };
    case 'below-average': return { type: 'aboveAverage', above: false };
    case 'above-equal-average': return { type: 'aboveAverage', above: true, equalAverage: true };
    case 'below-equal-average': return { type: 'aboveAverage', above: false, equalAverage: true };
    case 'error': return { type: 'containsErrors' };
    case 'no-error': return { type: 'notContainsErrors' };
  }
  return null;
}

const STOP_TYPES: Record<string, NonNullable<CFRule['stops']>[number]['type']> = { minimum: 'min', maximum: 'max', 'auto-minimum': 'min', 'auto-maximum': 'max', percent: 'percent', percentile: 'percentile', number: 'num', formula: 'formula' };
function stopOf(e: Element, conv: (f: string) => string): NonNullable<CFRule['stops']>[number] {
  const type = STOP_TYPES[a(e, 'calcext', 'type') || 'minimum'] || 'num';
  const v = a(e, 'calcext', 'value');
  const stop: NonNullable<CFRule['stops']>[number] = { type };
  if (type !== 'min' && type !== 'max' && v != null) stop.value = Number.isFinite(Number(v)) ? Number(v) : conv(v);
  return stop;
}

/**
 * Reads a table's calcext:conditional-formats. addr turns an ODF range address into ranges of
 * this sheet; styleOf gives the formatting of a named cell style.
 */
export function readConditionalFormats(table: Element, addr: (s: string) => Range[], styleOf: (name: string) => Partial<CellStyle>, conv: (f: string) => string, warnings: string[]): CFRule[] {
  const out: CFRule[] = [];
  let prio = 0;
  for (const cfs of kids(table, 'conditional-formats')) for (const cf of kids(cfs, 'conditional-format')) {
    const ranges = addr(a(cf, 'calcext', 'target-range-address') || '');
    if (!ranges.length) continue;
    for (const e of kids(cf)) {
      const r: CFRule = { id: newId('cf'), ranges, type: 'cellIs', priority: ++prio };
      if (e.localName === 'condition') {
        const p = parseCondition(a(e, 'calcext', 'value') || '', conv);
        if (!p) { if (!warnings.includes('Some conditional formats were not opened.')) warnings.push('Some conditional formats were not opened.'); continue; }
        Object.assign(r, p);
        const st = a(e, 'calcext', 'apply-style-name');
        if (st) r.style = styleOf(st);
      } else if (e.localName === 'date-is') {
        r.type = 'timePeriod';
        r.timePeriod = DATE_IS[a(e, 'calcext', 'date') || 'today'] || 'today';
        const st = a(e, 'calcext', 'style');
        if (st) r.style = styleOf(st);
      } else if (e.localName === 'color-scale') {
        const entries = kids(e, 'color-scale-entry');
        r.type = 'colorScale';
        r.stops = entries.map((x) => stopOf(x, conv));
        r.colors = entries.map((x) => a(x, 'calcext', 'color') || '#000000');
      } else if (e.localName === 'data-bar') {
        r.type = 'dataBar';
        r.barColor = a(e, 'calcext', 'positive-color') || '#638ec6';
        r.gradientBar = a(e, 'calcext', 'gradient') !== 'false';
        if (a(e, 'calcext', 'show-value') === 'false') r.showValue = false;
        r.stops = kids(e, 'formatting-entry').map((x) => stopOf(x, conv));
      } else if (e.localName === 'icon-set') {
        r.type = 'iconSet';
        r.iconSet = a(e, 'calcext', 'icon-set-type') || '3TrafficLights1';
        if (a(e, 'calcext', 'show-value') === 'false') r.showValue = false;
        r.stops = kids(e, 'formatting-entry').map((x) => stopOf(x, conv));
      } else continue;
      out.push(r);
    }
  }
  return out;
}

/** Converts ODF style:map conditions of cell styles used on this sheet into conditional formats. */
export function styleMapRules(sheet: Sheet, maps: (style: string) => { condition: string; apply: string }[], styleOf: (name: string) => Partial<CellStyle>, conv: (f: string) => string): CFRule[] {
  const n = notes.get(sheet);
  const out: CFRule[] = [];
  if (!n) return out;
  for (const [style, keys] of n.styleMap) {
    const ranges = cellsToRanges(keys);
    for (const m of maps(style)) {
      const p = parseCondition(m.condition.replace(/^(of|oooc):/, ''), conv);
      if (!p) continue;
      out.push({ id: newId('cf'), ranges, type: 'cellIs', ...p, style: styleOf(m.apply) } as CFRule);
    }
  }
  return out;
}

// --------------------------------------------------------------------------------------------
// Charts

/** Reads an ODF chart document (office:chart) into a chart definition (without its anchor). */
export function readChart(doc: Document | Element, addrToRef: (s: string) => string | null): Partial<ChartDef> | null {
  const chart = desc(doc, 'chart').find((c) => c.namespaceURI === NS.chart || c.prefix === 'chart');
  if (!chart) return null;
  const styles = new Map<string, Element>();
  for (const st of desc(doc, 'style')) { const n = a(st, 'style', 'name'); if (n) styles.set(n, st); }
  const props = (el: Element | null): Element | null => { const st = styles.get(a(el, 'chart', 'style-name') || ''); return st ? kid(st, 'chart-properties') : null; };
  const plot = kid(chart, 'plot-area');
  const pp = props(plot);
  const cls = (a(chart, 'chart', 'class') || 'chart:bar').replace(/^chart:/, '');
  const vertical = a(pp, 'chart', 'vertical') === 'true';
  const stacked = a(pp, 'chart', 'stacked') === 'true', percent = a(pp, 'chart', 'percentage') === 'true';
  const symbols = (el: Element | null) => { const t = a(props(el), 'chart', 'symbol-type'); return t != null && t !== 'none'; };
  const typeOf = (c: string, el: Element | null): ChartType => {
    switch (c) {
      case 'bar': return vertical ? (percent ? 'percentBar' : stacked ? 'stackedBar' : 'bar') : percent ? 'percentColumn' : stacked ? 'stackedColumn' : 'column';
      case 'line': return stacked || percent ? 'stackedLine' : symbols(el) || symbols(plot) ? 'lineMarkers' : 'line';
      case 'area': return percent ? 'percentArea' : stacked ? 'stackedArea' : 'area';
      case 'circle': return 'pie';
      case 'ring': return 'doughnut';
      case 'scatter': return a(props(el), 'chart', 'lines') === 'false' || a(pp, 'chart', 'lines') === 'false' ? 'scatter' : 'scatterLines';
      case 'radar': return 'radar';
      case 'filled-radar': return 'filledRadar';
      case 'bubble': return 'bubble';
      case 'stock': return 'stock';
    }
    return 'column';
  };
  const main = typeOf(cls, plot);
  const xAxis = kids(plot, 'axis').find((x) => a(x, 'chart', 'dimension') === 'x');
  const yAxis = kids(plot, 'axis').find((x) => a(x, 'chart', 'dimension') === 'y');
  const cats = a(kid(xAxis || null, 'categories'), 'table', 'cell-range-address');
  const series: ChartSeries[] = [];
  const types = new Set<ChartType>();
  for (const s of kids(plot, 'series')) {
    const vals = addrToRef(a(s, 'chart', 'values-cell-range-address') || '');
    if (!vals) continue;
    const ser: ChartSeries = { values: vals };
    const label = a(s, 'chart', 'label-cell-address');
    if (label) { const r = addrToRef(label); if (r) ser.nameRef = r; }
    const domain = kids(s, 'domain').map((d) => a(d, 'table', 'cell-range-address')).filter(Boolean) as string[];
    const catRef = domain.length ? addrToRef(domain[0]) : cats ? addrToRef(cats) : null;
    if (catRef) ser.categories = catRef;
    if (main === 'bubble' && domain.length > 1) { const x = addrToRef(domain[1]); if (x) { ser.sizes = ser.values; ser.values = vals; ser.categories = x; } }
    const sc = a(s, 'chart', 'class');
    const t = sc ? typeOf(sc.replace(/^chart:/, ''), s) : main;
    types.add(t);
    if (t !== main) ser.type = t;
    const fill = styles.get(a(s, 'chart', 'style-name') || '');
    const gp = fill ? kid(fill, 'graphic-properties') : null;
    const color = gp && (gp.getAttributeNS('urn:oasis:names:tc:opendocument:xmlns:drawing:1.0', 'fill-color') || gp.getAttribute('draw:fill-color'));
    if (color && /^#[0-9a-f]{6}$/i.test(color)) ser.color = color.toLowerCase();
    series.push(ser);
  }
  if (!series.length) return null;
  const def: Partial<ChartDef> = { type: types.size > 1 ? 'combo' : main, series };
  if (def.type === 'combo') for (const s of series) s.type = s.type || main;
  const title = kid(chart, 'title');
  if (title) def.title = kids(title, 'p').map((p) => p.textContent || '').join('\n');
  const legend = kid(chart, 'legend');
  if (!legend) def.legend = 'none';
  else {
    const pos = a(legend, 'chart', 'legend-position') || 'end';
    def.legend = pos === 'bottom' ? 'bottom' : pos === 'top' ? 'top' : pos === 'start' ? 'left' : 'right';
  }
  const axTitle = (ax: Element | undefined) => (ax ? kids(kid(ax, 'title'), 'p').map((p) => p.textContent || '').join('\n') : '');
  const xt = axTitle(xAxis), yt = axTitle(yAxis);
  if (xt) def.xTitle = xt;
  if (yt) def.yTitle = yt;
  def.gridlines = !!yAxis && kids(yAxis, 'grid').some((g) => a(g, 'chart', 'class') !== 'minor');
  const yp = props(yAxis || null);
  const mn = a(yp, 'chart', 'minimum'), mx = a(yp, 'chart', 'maximum');
  if (mn != null && Number.isFinite(+mn)) def.yMin = +mn;
  if (mx != null && Number.isFinite(+mx)) def.yMax = +mx;
  if (a(yp, 'chart', 'logarithmic') === 'true') def.logScale = true;
  if (desc(plot, 'data-label').length || a(pp, 'chart', 'data-label-number') === 'value') def.dataLabels = true;
  return def;
}
