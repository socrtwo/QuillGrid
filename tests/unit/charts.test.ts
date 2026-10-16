import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import { Workbook } from '../../src/model/workbook';
import { Engine } from '../../src/calc/engine';
import type { ChartDef, ChartType, DrawingObject } from '../../src/model/types';
import { renderChartSvg, chartData, niceScale, histogramBins, squarify, quartiles, trendline, CHART_TYPES, chartPalette } from '../../src/charts/render';
import { renderShapeSvg, SHAPES, shapeGeometry, diagramSvg, DIAGRAMS, equationRuns, WORDART_STYLES } from '../../src/charts/shapes';
import { chartFromRange, detectLayout } from '../../src/commands/charts';

function setup() {
  const book = new Workbook();
  const s = book.sheet;
  s.name = 'Sheet 1';
  const rows: any[][] = [
    ['Month', 'North', 'South', 'West'],
    ['Jan', 10, 7, 3],
    ['Feb', 14, -2, 5],
    ['Mar', 9, 11, 8],
    ['Total', 33, 16, 16],
  ];
  book.transact('t', () => rows.forEach((row, r) => row.forEach((v, c) => s.set(r, c, { v }))));
  book.transact('fmt', () => { const sid = book.styleId({ numFmt: '$#,##0' }); for (let r = 1; r <= 4; r++) s.patch(r, 1, { s: sid }); });
  const engine = new Engine(book);
  engine.rebuild(); engine.recalc(true);
  const tab = { book, engine, get sheet() { return book.sheet; } } as any;
  return { book, s, engine, tab };
}

const chart = (type: ChartType, extra: Partial<ChartDef> = {}): ChartDef => ({
  id: 'ch1', type, at: { r: 0, c: 6, dx: 0, dy: 0, w: 480, h: 300 },
  series: [
    { nameRef: "'Sheet 1'!$B$1", values: "'Sheet 1'!$B$2:$B$4", categories: "'Sheet 1'!$A$2:$A$4" },
    { nameRef: "'Sheet 1'!$C$1", values: "'Sheet 1'!$C$2:$C$4", categories: "'Sheet 1'!$A$2:$A$4" },
    { nameRef: "'Sheet 1'!$D$1", values: "'Sheet 1'!$D$2:$D$4", categories: "'Sheet 1'!$A$2:$A$4", sizes: "'Sheet 1'!$B$2:$B$4" },
  ],
  ...extra,
});

function wellFormed(svg: string): Document {
  const doc = new DOMParser().parseFromString(svg, 'image/svg+xml');
  expect(doc.getElementsByTagName('parsererror').length).toBe(0);
  expect(doc.documentElement.nodeName).toBe('svg');
  return doc;
}

describe('chart data', () => {
  it('reads series names, values, categories and formats from sheet-qualified refs', () => {
    const { tab, book } = setup();
    const ch = chart('column');
    book.sheet.charts.push(ch);
    const d = chartData(tab, ch);
    expect(d.categories).toEqual(['Jan', 'Feb', 'Mar']);
    expect(d.series.map((x) => x.name)).toEqual(['North', 'South', 'West']);
    expect(d.series[1].values).toEqual([7, -2, 11]);
    expect(d.series[0].fmt).toBe('$#,##0');
    expect(d.series[0].color).toBe(book.theme.colors[4]);
  });
  it('treats errors and text as gaps', () => {
    const { tab, book, s } = setup();
    book.transact('e', () => { s.set(2, 1, { f: '1/0' }); s.set(3, 1, { v: 'n/a' }); });
    tab.engine.recalc(true);
    const d = chartData(tab, chart('line'));
    expect(d.series[0].values).toEqual([10, null, null]);
    expect(d.series[0].errors).toBe(1);
  });
  it('palette extends past the theme accents', () => {
    const p = chartPalette(null, {}, 14);
    expect(new Set(p).size).toBe(14);
  });
});

describe('scales and layouts', () => {
  it('nice scale uses 1-2-5 steps and includes zero', () => {
    const s = niceScale(3, 97, 5, { includeZero: true });
    expect(s.min).toBe(0); expect(s.max).toBe(100); expect(s.step).toBe(20);
    expect(s.ticks).toEqual([0, 20, 40, 60, 80, 100]);
    const t = niceScale(0.12, 0.87, 5);
    expect(t.step).toBeCloseTo(0.2); expect(t.ticks[0]).toBeCloseTo(0);
    const n = niceScale(-35, 12, 5, { includeZero: true });
    expect(n.min).toBeLessThanOrEqual(-35); expect(n.max).toBeGreaterThanOrEqual(12);
    const l = niceScale(3, 4500, 5, { log: true });
    expect(l.ticks).toEqual([1, 10, 100, 1000, 10000]);
    const f = niceScale(0, 50, 5, { fixedMin: 10, fixedMax: 40 });
    expect(f.min).toBe(10); expect(f.max).toBe(40);
    const same = niceScale(5, 5, 5);
    expect(same.max).toBeGreaterThan(same.min);
  });
  it('histogram bins cover all values', () => {
    const vals = [1, 2, 2, 3, 3, 3, 4, 4, 5, 9, 10, 15];
    const bins = histogramBins(vals);
    expect(bins.reduce((a, b) => a + b.count, 0)).toBe(vals.length);
    expect(bins[0].lo).toBe(1);
    expect(bins[bins.length - 1].hi).toBeGreaterThanOrEqual(15);
    const fixed = histogramBins(vals, { binCount: 2 });
    expect(fixed.length).toBe(2);
    expect(fixed.map((b) => b.count)).toEqual([9, 3]);
    expect(histogramBins([])).toEqual([]);
    expect(histogramBins([4, 4])[0].count).toBe(2);
  });
  it('treemap squarify fills the box proportionally', () => {
    const vals = [6, 6, 4, 3, 2, 2, 1];
    const box = { x: 0, y: 0, w: 600, h: 400 };
    const rects = squarify(vals, box);
    const total = vals.reduce((a, b) => a + b, 0);
    rects.forEach((r, i) => expect(r.w * r.h).toBeCloseTo((vals[i] / total) * 600 * 400, 3));
    for (const r of rects) { expect(r.x).toBeGreaterThanOrEqual(-1e-6); expect(r.x + r.w).toBeLessThanOrEqual(600 + 1e-6); expect(r.y + r.h).toBeLessThanOrEqual(400 + 1e-6); }
    // Squarified: aspect ratios stay reasonable
    for (const r of rects) expect(Math.max(r.w / r.h, r.h / r.w)).toBeLessThan(4);
  });
  it('quartiles and trendlines', () => {
    const q = quartiles([1, 2, 3, 4, 5, 6, 7, 8, 100])!;
    expect(q.median).toBe(5); expect(q.outliers).toEqual([100]);
    const lin = trendline('linear', [1, 2, 3, 4], [2, 4, 6, 8]);
    expect(lin[lin.length - 1].y).toBeCloseTo(8);
    const ex = trendline('exponential', [0, 1, 2, 3], [1, 2, 4, 8]);
    expect(ex[ex.length - 1].y).toBeCloseTo(8, 4);
    const poly = trendline('polynomial', [0, 1, 2, 3], [0, 1, 4, 9]);
    expect(poly[poly.length - 1].y).toBeCloseTo(9, 4);
    expect(trendline('movingAverage', [1, 2, 3], [2, 4, 6]).map((p) => p.y)).toEqual([3, 5]);
  });
});

describe('renderChartSvg', () => {
  it('produces well-formed SVG for every chart type', () => {
    const { tab, book } = setup();
    for (const type of CHART_TYPES) {
      const ch = chart(type, { title: `A <${type}> & "chart"`, dataLabels: true, xTitle: 'X', yTitle: 'Y', trendline: 'linear', legend: 'right' });
      if (type === 'combo') ch.series[2].secondary = true;
      book.sheet.charts = [ch];
      const svg = renderChartSvg(tab, ch, 480, 300);
      const doc = wellFormed(svg);
      expect(doc.documentElement.textContent).toContain(`A <${type}> & "chart"`);
    }
  });
  it('handles empty, error and tiny inputs', () => {
    const { tab } = setup();
    wellFormed(renderChartSvg(tab, { ...chart('column'), series: [] }, 300, 200));
    wellFormed(renderChartSvg(tab, { ...chart('pie'), series: [{ values: '1/0' }] }, 300, 200));
    wellFormed(renderChartSvg(tab, chart('treemap'), 20, 10));
    wellFormed(renderChartSvg(tab, { ...chart('line'), series: [{ values: '{1,2,3}' }] }, 300, 200));
  });
  it('formats value axis labels with the source number format', () => {
    const { tab } = setup();
    const svg = renderChartSvg(tab, { ...chart('column'), series: [chart('column').series[0]] }, 480, 300);
    expect(svg).toMatch(/\$10/);
  });
  it('log scale, fixed bounds and multi-level hierarchy', () => {
    const { tab, book, s } = setup();
    wellFormed(renderChartSvg(tab, chart('line', { logScale: true, yMin: 1, yMax: 100 }), 400, 260));
    book.transact('h', () => { [['A', 'x', 3], ['A', 'y', 2], ['B', 'z', 5]].forEach((row, r) => row.forEach((v, c) => s.set(10 + r, c, { v }))); });
    tab.engine.recalc(true);
    const tm: ChartDef = { id: 't', type: 'treemap', at: { r: 0, c: 0, dx: 0, dy: 0, w: 400, h: 300 }, series: [{ values: "'Sheet 1'!$C$11:$C$13", categories: "'Sheet 1'!$A$11:$B$13" }] };
    const d = chartData(tab, tm);
    expect(d.catLevels).toEqual([['A', 'x'], ['A', 'y'], ['B', 'z']]);
    wellFormed(renderChartSvg(tab, tm, 400, 300));
    wellFormed(renderChartSvg(tab, { ...tm, type: 'sunburst' }, 400, 300));
  });
});

describe('chart from range', () => {
  it('detects headers, categories and series orientation', () => {
    const { s } = setup();
    const lay = detectLayout(s, { r1: 0, c1: 0, r2: 3, c2: 3 });
    expect(lay.headerRow).toBe(true); expect(lay.catCol).toBe(true); expect(lay.seriesInRows).toBe(false);
    const ch = chartFromRange(s, { r1: 0, c1: 0, r2: 3, c2: 3 }, 'column');
    expect(ch.series.length).toBe(3);
    expect(ch.series[0].values).toBe("'Sheet 1'!$B$2:$B$4");
    expect(ch.series[0].categories).toBe("'Sheet 1'!$A$2:$A$4");
    expect(ch.series[0].nameRef).toBe("'Sheet 1'!$B$1");
    const wide = chartFromRange(s, { r1: 0, c1: 0, r2: 1, c2: 3 }, 'column');
    expect(wide.series.length).toBe(1);
    const sc = chartFromRange(s, { r1: 0, c1: 1, r2: 3, c2: 3 }, 'scatter');
    expect(sc.series.length).toBe(2);
    expect(sc.series[0].categories).toBe("'Sheet 1'!$B$2:$B$4");
  });
});

describe('shapes', () => {
  it('has a large gallery and renders every shape as well-formed SVG', () => {
    expect(SHAPES.length).toBeGreaterThanOrEqual(40);
    for (const sh of SHAPES) {
      const d: DrawingObject = { id: 'd' + sh.id, kind: 'shape', shape: sh.id, at: { r: 0, c: 0, dx: 0, dy: 0, w: 140, h: 90 }, text: 'Hi <there>', fill: 'gradient:#ff0000:#0000ff', fx: { shadow: true, flipH: true } };
      wellFormed(renderShapeSvg(d));
      expect(shapeGeometry(sh.id, 100, 60).body.length).toBeGreaterThan(10);
    }
  });
  it('renders text boxes, decorative text, icons, diagrams, equations and slicers', () => {
    const base = { at: { r: 0, c: 0, dx: 0, dy: 0, w: 300, h: 160 } };
    for (const st of WORDART_STYLES) wellFormed(renderShapeSvg({ ...base, id: 'w', kind: 'wordart', text: 'Big & bold', font: { color: st.fill }, fx: { textOutline: st.outline, shadow: st.shadow, glow: st.glow } }));
    wellFormed(renderShapeSvg({ ...base, id: 't', kind: 'textbox', text: 'A long line of text that needs wrapping inside the box' }));
    wellFormed(renderShapeSvg({ ...base, id: 'i', kind: 'icon', shape: 'icon', text: 'home' }));
    wellFormed(renderShapeSvg({ ...base, id: 'e', kind: 'textbox', shape: 'equation', text: 'x^2 + y_1 = \\alpha' }));
    wellFormed(renderShapeSvg({ ...base, id: 's', kind: 'shape', shape: 'slicer', text: 'Region', fx: { items: ['East', 'West'], selected: ['East'] } }));
    wellFormed(renderShapeSvg({ ...base, id: 'o', kind: 'shape', shape: 'embed', text: 'report.pdf' }));
    for (const dg of DIAGRAMS) wellFormed(diagramSvg(dg.kind, 'One\n  Two\nThree\nFour', 400, 220));
    expect(equationRuns('x^2+a_(n+1)')).toEqual([{ t: 'x', k: 'n' }, { t: '2', k: 'sup' }, { t: '+a', k: 'n' }, { t: 'n+1', k: 'sub' }]);
    // unsafe colors are dropped
    const svg = renderShapeSvg({ ...base, id: 'x', kind: 'shape', shape: 'rect', fill: 'red" onload="alert(1)' });
    expect(svg).not.toContain('onload');
  });
});
