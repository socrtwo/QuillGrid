/**
 * Heritage commands: features from spreadsheets that came before (VisiCalc, Lotus 1-2-3, Symphony,
 * Improv, Agenda, Quattro Pro, Javelin, Multiplan, Excel for Macintosh, Resolver One) and ideas from
 * newer ones (Airtable and Grist typed columns, Causal-style uncertain inputs and sliders, cell edit
 * history). docs/HERITAGE.md has the survey behind them; the logic lives in ../heritage/*.
 */
import { def } from './registry';
import type { App } from '../app';
import type { Sheet } from '../model/workbook';
import type { Cell, CellStyle, ChartDef, DataValidation, Scalar } from '../model/types';
import { cellName, cellKey, rangeName, quoteSheet, colName, type Range } from '../model/address';
import { activeCell, activeRange, dataRegion, hasHeaderRow, cellText, setInput, clipToUsed, selRanges } from './ops';
import { addSheetWithRows, writeOut, checkWritable, workRanges, HEAD, BOLD } from './tools';
import { subtractRange } from './dataops';
import { resolveRangeText, rangeRefText, insertCellsShift } from '../model/structure';
import { shiftFormula } from '../calc/parser';
import { getFn } from '../calc/functions/registry';
import { newId } from '../model/workbook';
import { openDialog, toast, confirmDialog, type Field, type DialogControl } from '../ui/dialog';
import { h, icon, clear } from '../ui/dom';
import { registerPane, typingIn } from '../ui/panes';
import { openSlashMenu } from '../ui/slashMenu';
import { renderChartSvg } from '../charts/render';
import { lotusEntry, LOTUS_FUNCTIONS } from '../heritage/lotus';
import {
  toR1C1, formulaRules, inconsistentFormulas, explainTree, explainText, dependencyGraph, cellHistory, cellContentText, displayValue,
  MAP_STYLE, type ExplainNode,
} from '../heritage/audit';
import {
  matInverse, matMultiply, matSolve, matDeterminant, monteCarlo, parseDist, summarize, seeded, leafInputs, sliderRange,
  matchesCondition, COND_LABELS, COLUMN_TYPES, distinctChoices, typeMismatches, type Dist, type CondOp, type ColumnType, type Mat,
} from '../heritage/models';
import { sheetToCode, type CodeLang } from '../heritage/codegen';
import '../styles/heritage.css';

const T = 'Tools', F = 'Formulas', D = 'Data', V = 'View';

const goTo = (app: App, sheet: Sheet, r: number, c: number, range?: Range) => {
  const i = app.book.sheets.indexOf(sheet);
  if (i >= 0 && i !== app.book.activeSheet) app.exec('sheet.activate', i);
  app.grid.select(r, c, range ? ({ range } as any) : undefined);
  app.grid.scrollIntoView(r, c);
};
const addr = (app: App, sheet: Sheet, r: number, c: number) => (sheet === app.sheet ? cellName(r, c) : `${quoteSheet(sheet.name)}!${cellName(r, c)}`);

/** A nearby label for a cell: text to its left in the row, else above it in the column. */
export function labelFor(sheet: Sheet, r: number, c: number): string {
  for (let cc = c - 1; cc >= Math.max(0, c - 4); cc--) { const v = sheet.value(r, cc); if (typeof v === 'string' && v.trim()) return v.trim(); }
  for (let rr = r - 1; rr >= Math.max(0, r - 4); rr--) { const v = sheet.value(rr, c); if (typeof v === 'string' && v.trim()) return v.trim(); }
  return '';
}

// ============================================================================================
// 1-2-3 formula entry

/** Converts text cells that hold 1-2-3 formulas (@SUM(A1..A5), +A1*@AVG(...)) into live formulas. */
export function convertLotusCells(app: App, ranges: Range[]): number {
  const s = app.sheet;
  const names = app.book.sheets.map((x) => x.name);
  let n = 0;
  app.book.transact('Convert 1-2-3 Formulas', () => {
    for (const rg of ranges) s.forEachIn(rg, (r, c, cell) => {
      if (cell.f != null || typeof cell.v !== 'string') return;
      let t = cell.v.trim();
      if (!/^[@+=-]/.test(t)) t = '+' + t; // labels such as A1..A5 or (B2*3) typed without a prefix
      const f = lotusEntry(t, (nm) => !!getFn(nm), { sheetNames: names });
      if (!f) return;
      s.set(r, c, { ...(cell.s ? { s: cell.s } : {}), f });
      n++;
    });
  });
  return n;
}

// ============================================================================================
// Export as code

async function exportCode(app: App, arg?: CodeLang): Promise<void> {
  const sheet = app.sheet;
  if (!sheet.usedRange(false)) { toast('This sheet is empty.'); return; }
  let lang: CodeLang = arg === 'js' ? 'js' : 'python';
  const area = h('textarea', { class: 'ql-code-out', rows: '18', readonly: true, spellcheck: 'false', 'aria-label': 'Generated program' }) as HTMLTextAreaElement;
  const info = h('p', { class: 'ql-help' });
  const gen = () => {
    const res = sheetToCode(app.book, sheet, lang);
    area.value = res.code;
    info.textContent = `${res.formulas} formula${res.formulas === 1 ? '' : 's'} and ${res.cells - res.formulas} input${res.cells - res.formulas === 1 ? '' : 's'} from ${sheet.name}.` + (res.unsupported.length ? ` Not translated: ${res.unsupported.join(', ')}.` : ' Every formula was translated.');
    return res;
  };
  gen();
  const fileName = () => `${(sheet.name.replace(/[^\w-]+/g, '_') || 'sheet').toLowerCase()}.${lang === 'js' ? 'js' : 'py'}`;
  await openDialog({
    title: 'Export as Code', wide: true, ok: false, cancel: 'Close',
    help: 'The sheet as a program: each cell is a variable and each formula a line of code, computed in dependency order.',
    content: h('div', null, info, area),
    fields: [{ key: 'lang', label: 'Language', type: 'radio', value: lang, options: [{ value: 'python', label: 'Python' }, { value: 'js', label: 'JavaScript' }] }],
    onChange: (v, _ctl, key) => { if (key === 'lang' && v.lang !== lang) { lang = v.lang; gen(); } },
    buttons: [
      { label: 'Copy', onClick: async () => { const { writeClipboardText } = await import('../platform/native'); toast((await writeClipboardText(area.value)) ? 'Copied' : 'Copy failed'); } },
      { label: 'Save…', primary: true, onClick: async () => { const { saveAs } = await import('../platform/native'); await saveAs(new Blob([area.value], { type: 'text/plain' }), fileName()); } },
    ],
  });
}

// ============================================================================================
// Panes: Explain Value, Dependency Diagram, Formula Rules, Info Window, What-If Sliders

interface Target { sheet: Sheet; r: number; c: number }

/** The cell a pane describes: the pinned cell, or the active cell when following the selection. */
function paneTarget(app: App, el: HTMLElement): Target {
  if (el.dataset.pin) { const [id, r, c] = el.dataset.pin.split(':'); const sh = app.book.sheets.find((x) => x.id === id); if (sh) return { sheet: sh, r: +r, c: +c }; }
  return { sheet: app.sheet, r: app.sheet.sel.r, c: app.sheet.sel.c };
}

function pinToggle(app: App, el: HTMLElement, rerender: () => void): HTMLElement {
  const pinned = !!el.dataset.pin;
  return h('button', {
    class: 'ql-btn ql-sm', title: pinned ? 'Follow the selection again' : 'Keep showing this cell while you move around',
    onclick: () => { if (el.dataset.pin) delete el.dataset.pin; else { const t = paneTarget(app, el); el.dataset.pin = `${t.sheet.id}:${t.r}:${t.c}`; } el.dataset.sig = ''; rerender(); },
  }, icon(pinned ? 'push_pin' : 'keep'), pinned ? ' Unpin' : ' Pin');
}

function explainNodeEl(app: App, n: ExplainNode, depth: number): HTMLElement {
  const link = h('button', { class: 'ql-link', onclick: (e: Event) => { e.preventDefault(); goTo(app, n.sheet, n.r, n.c, n.range); } }, n.address);
  const line = [link, ' = ', h('span', { class: 'ql-x-val' }, n.value), n.note ? h('span', { class: 'ql-x-note' }, ` (${n.note})`) : null, n.formula ? h('div', { class: 'ql-x-f' }, n.formula) : null];
  if (!n.children.length) return h('div', { class: 'ql-x-leaf' }, ...line);
  const d = h('details', { open: depth < 2 }, h('summary', null, ...line), ...n.children.map((ch) => explainNodeEl(app, ch, depth + 1)));
  return d;
}

registerPane({
  id: 'explain', title: 'Explain Value',
  render(app, el) { el.dataset.sig = ''; this.update!(app, el, {}); },
  update(app, el) {
    const t = paneTarget(app, el);
    const sig = `${app.tab.id}:${app.book.version}:${app.engine.evaluations}:${t.sheet.id}:${t.r}:${t.c}:${el.dataset.pin || ''}`;
    if (el.dataset.sig === sig || typingIn(el)) return;
    el.dataset.sig = sig;
    clear(el);
    const tree = explainTree(app.engine, t.sheet, t.r, t.c);
    el.appendChild(h('div', { class: 'ql-row' }, pinToggle(app, el, () => this.update!(app, el, {})),
      h('button', { class: 'ql-btn ql-sm', onclick: async () => { const { writeClipboardText } = await import('../platform/native'); toast((await writeClipboardText(explainText(tree))) ? 'Copied' : 'Copy failed'); } }, icon('content_copy'), ' Copy')));
    if (!tree.formula) {
      el.appendChild(h('p', { class: 'ql-help' }, `${tree.address} holds ${tree.value === '(blank)' ? 'nothing' : 'a constant, not a formula'}. Select a formula to see where its value comes from.`));
      const deps = app.engine.dependentsOfCell(t.sheet, t.r, t.c);
      if (deps.length) el.appendChild(h('p', { class: 'ql-help' }, `It feeds ${deps.length} formula${deps.length === 1 ? '' : 's'}: `, ...deps.slice(0, 20).map((d) => h('button', { class: 'ql-link', onclick: () => goTo(app, d.sheet, d.r, d.c) }, addr(app, d.sheet, d.r, d.c)))));
      return;
    }
    el.appendChild(h('p', { class: 'ql-help' }, 'How this value is built: each formula, then the values it reads. Click an address to go there.'));
    el.appendChild(h('div', { class: 'ql-explain' }, explainNodeEl(app, tree, 0)));
  },
});

registerPane({
  id: 'depDiagram', title: 'Dependency Diagram',
  render(app, el) { el.dataset.sig = ''; el.dataset.levels = el.dataset.levels || '2'; this.update!(app, el, {}); },
  update(app, el) {
    const t = paneTarget(app, el);
    const levels = +(el.dataset.levels || 2);
    const sig = `${app.tab.id}:${app.book.version}:${t.sheet.id}:${t.r}:${t.c}:${levels}:${el.dataset.pin || ''}`;
    if (el.dataset.sig === sig || typingIn(el)) return;
    el.dataset.sig = sig;
    clear(el);
    const g = dependencyGraph(app.engine, t.sheet, t.r, t.c, levels);
    const sel = h('select', { 'aria-label': 'Levels', onchange: (e: Event) => { el.dataset.levels = (e.target as HTMLSelectElement).value; el.dataset.sig = ''; this.update!(app, el, {}); } },
      ...[1, 2, 3, 4].map((n) => h('option', { value: String(n), selected: n === levels }, `${n} level${n === 1 ? '' : 's'}`)));
    el.appendChild(h('div', { class: 'ql-row' }, pinToggle(app, el, () => this.update!(app, el, {})), sel));
    el.appendChild(h('p', { class: 'ql-help' }, 'Inputs at the top flow down into the results that use them. The selected cell is highlighted. Click a box to go to it.'));
    if (g.nodes.length === 1) { el.appendChild(h('p', { class: 'ql-help' }, `${g.nodes[0].label} has no precedents or dependents.`)); }
    // Levels run top to bottom (the pane is tall and narrow); a level wraps onto more rows when it is wide.
    const lv = [...new Set(g.nodes.map((n) => n.level))].sort((a, b) => a - b);
    const W = 98, H = 44, GX = 8, GY = 26;
    const avail = Math.max(W + 2 * GX, (el.clientWidth || 300) - 4);
    const perRow = Math.max(1, Math.floor((avail - GX) / (W + GX)));
    const byLevel = new Map<number, typeof g.nodes>();
    for (const n of g.nodes) { const a = byLevel.get(n.level) || []; a.push(n); byLevel.set(n.level, a); }
    const width = Math.min(perRow, Math.max(...[...byLevel.values()].map((a) => a.length))) * (W + GX) + GX;
    const pos = new Map<string, { x: number; y: number }>();
    let y = GY / 2;
    for (const l of lv) {
      const a = byLevel.get(l)!;
      for (let i = 0; i < a.length; i += perRow) {
        const row = a.slice(i, i + perRow);
        const off = (width - row.length * (W + GX) - GX) / 2;
        row.forEach((n, j) => pos.set(n.id, { x: GX + off + j * (W + GX), y }));
        y += H + GY / 2;
      }
      y += GY / 2;
    }
    const height = y;
    const NS = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('viewBox', `0 0 ${width} ${height}`); svg.setAttribute('width', String(width)); svg.setAttribute('height', String(height)); svg.setAttribute('role', 'img');
    svg.setAttribute('aria-label', `Dependency diagram of ${g.nodes.find((n) => n.level === 0)?.label}`);
    const mk = (tag: string, attrs: Record<string, string | number>, text?: string) => { const e = document.createElementNS(NS, tag); for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, String(v)); if (text != null) e.textContent = text; return e; };
    const defs = mk('defs', {});
    const marker = mk('marker', { id: 'ql-dg-arrow', viewBox: '0 0 10 10', refX: 9, refY: 5, markerWidth: 7, markerHeight: 7, orient: 'auto-start-reverse' });
    marker.appendChild(mk('path', { d: 'M0,0 L10,5 L0,10 z', fill: 'currentColor' }));
    defs.appendChild(marker); svg.appendChild(defs);
    for (const e of g.edges) {
      const a = pos.get(e.from), b = pos.get(e.to);
      if (!a || !b) continue;
      const x1 = a.x + W / 2, y1 = a.y + H, x2 = b.x + W / 2, y2 = b.y;
      svg.appendChild(mk('path', { d: `M${x1},${y1} C${x1},${y1 + GY / 2} ${x2},${y2 - GY / 2} ${x2},${y2}`, fill: 'none', stroke: 'currentColor', 'stroke-opacity': 0.45, 'marker-end': 'url(#ql-dg-arrow)' }));
    }
    for (const n of g.nodes) {
      const p = pos.get(n.id)!;
      const grp = mk('g', { class: 'ql-dg-node', transform: `translate(${p.x},${p.y})`, tabindex: 0, role: 'button' });
      const fill = n.level === 0 ? 'var(--accent-soft)' : n.formula ? 'var(--surface2)' : 'var(--surface)';
      grp.appendChild(mk('rect', { width: W, height: H, rx: 7, fill, stroke: n.level === 0 ? 'var(--accent)' : 'var(--border)', 'stroke-width': n.level === 0 ? 2 : 1 }));
      const clip = (s: string, n2: number) => (s.length > n2 ? s.slice(0, n2 - 1) + '…' : s);
      grp.appendChild(mk('text', { x: 8, y: 17, 'font-size': 12, 'font-weight': 600, fill: 'currentColor' }, clip(n.label, 13)));
      grp.appendChild(mk('text', { x: 8, y: 35, 'font-size': 12, fill: 'currentColor', 'fill-opacity': 0.75 }, clip(n.value, 13)));
      const title = mk('title', {}, `${n.label} = ${n.value}${n.formula ? `\n${n.formula}` : ''}`);
      grp.appendChild(title);
      grp.addEventListener('click', () => goTo(app, n.sheet, n.range.r1, n.range.c1, n.range.r1 === n.range.r2 && n.range.c1 === n.range.c2 ? undefined : n.range));
      grp.addEventListener('keydown', (ev) => { if ((ev as KeyboardEvent).key === 'Enter') goTo(app, n.sheet, n.range.r1, n.range.c1); });
      svg.appendChild(grp);
    }
    el.appendChild(h('div', { class: 'ql-depgraph' }, svg));
  },
});

registerPane({
  id: 'formulaRules', title: 'Formula Rules',
  render(app, el) { el.dataset.sig = ''; this.update!(app, el, {}); },
  update(app, el) {
    const s = app.sheet;
    const sig = `${app.tab.id}:${s.id}:${app.book.version}`;
    if (el.dataset.sig === sig || typingIn(el)) return;
    el.dataset.sig = sig;
    clear(el);
    const rules = formulaRules(s);
    const odd = inconsistentFormulas(s);
    const total = rules.reduce((a, r) => a + r.count, 0);
    el.appendChild(h('p', { class: 'ql-help' }, total ? `${total} formula cell${total === 1 ? '' : 's'} on ${s.name} use ${rules.length} distinct formula${rules.length === 1 ? '' : 's'}. Copies of one formula count once (they are equal in R1C1 notation).` : `${s.name} has no formulas.`));
    if (odd.length) {
      el.appendChild(h('h3', { class: 'ql-pane-h' }, icon('warning'), ` Unlike their neighbors (${odd.length})`));
      el.appendChild(h('div', { class: 'ql-row' }, ...odd.slice(0, 60).map((x) => h('button', { class: 'ql-link', onclick: () => goTo(app, s, x.r, x.c) }, cellName(x.r, x.c)))));
    }
    if (rules.length) el.appendChild(h('h3', { class: 'ql-pane-h' }, icon('rule'), ' Formulas'));
    for (const rule of rules.slice(0, 300)) {
      const rangesText = rule.ranges.slice(0, 6).map(rangeName).join(', ') + (rule.ranges.length > 6 ? ` +${rule.ranges.length - 6} more` : '');
      el.appendChild(h('div', { class: 'ql-rule' },
        h('code', null, '=' + rule.r1c1),
        h('div', { class: 'ql-help' }, `${rule.count} cell${rule.count === 1 ? '' : 's'}: `, h('button', { class: 'ql-link', onclick: () => { const rg = rule.ranges[0]; goTo(app, s, rg.r1, rg.c1, rg); } }, rangesText)),
        h('div', { class: 'ql-help' }, `as in ${cellName(rule.first.r, rule.first.c)}: =${rule.formula}`)));
    }
  },
});

/** Rows of the Info Window for a cell. */
export function cellInfoRows(app: App, sheet: Sheet, r: number, c: number): [string, string][] {
  const cell = sheet.get(r, c);
  const st = sheet.styleOf(r, c);
  const rows: [string, string][] = [];
  rows.push(['Cell', `${quoteSheet(sheet.name)}!${cellName(r, c)}`]);
  rows.push(['Value', displayValue(sheet, r, c)]);
  const v = cell?.v;
  rows.push(['Type', cell?.f != null ? 'Formula' : cell?.sp != null ? 'Spilled result' : v == null ? 'Blank' : typeof v === 'string' ? 'Text' : typeof v === 'number' ? 'Number' : typeof v === 'boolean' ? 'Logical' : 'Error']);
  if (cell?.f != null) { rows.push(['Formula', '=' + cell.f]); rows.push(['R1C1', '=' + toR1C1(cell.f, r, c)]); }
  if (cell?.arr) rows.push(['Array formula', rangeName(cell.arr)]);
  const spill = app.engine.spillRange(sheet, r, c);
  if (spill) rows.push(['Spills into', rangeName(spill)]);
  rows.push(['Number format', st.numFmt || 'General']);
  const font = [st.font || 'Default font', st.size ? `${st.size} pt` : '', st.bold ? 'bold' : '', st.italic ? 'italic' : '', st.underline ? 'underlined' : '', st.color || ''].filter(Boolean).join(', ');
  rows.push(['Font', font]);
  rows.push(['Alignment', [st.hAlign || 'general', st.vAlign || 'bottom', st.wrap ? 'wrap' : '', st.indent ? `indent ${st.indent}` : ''].filter(Boolean).join(', ')]);
  if (st.fill) rows.push(['Fill', st.fill]);
  rows.push(['Protection', `${st.locked === false ? 'Unlocked' : 'Locked'}${st.hidden ? ', formula hidden' : ''}${sheet.protection.on ? ' (sheet protected)' : ''}`]);
  const names = app.book.names.filter((n) => { const res = resolveRangeText(app.book, n.ref, sheet); return res && res.sheet === sheet && r >= res.range.r1 && r <= res.range.r2 && c >= res.range.c1 && c <= res.range.c2; }).map((n) => n.name);
  for (const t of sheet.tables) if (r >= t.range.r1 && r <= t.range.r2 && c >= t.range.c1 && c <= t.range.c2) names.push(`${t.name}[${t.columns[c - t.range.c1]?.name ?? ''}]`);
  if (names.length) rows.push(['Names', names.join(', ')]);
  const precs = app.engine.precedentsOf(sheet, r, c);
  if (precs.length) rows.push(['Precedents', precs.slice(0, 12).map((p) => (p.sheet === sheet ? '' : quoteSheet(p.sheet.name) + '!') + rangeName(p)).join(', ') + (precs.length > 12 ? ' …' : '')]);
  const deps = app.engine.dependentsOfCell(sheet, r, c);
  if (deps.length) rows.push(['Dependents', deps.slice(0, 12).map((d) => addr(app, d.sheet, d.r, d.c)).join(', ') + (deps.length > 12 ? ` … (${deps.length})` : '')]);
  const dv = sheet.validations.find((x) => x.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2));
  if (dv) rows.push(['Validation', `${dv.type}${dv.formula1 ? ' ' + dv.formula1 : ''}${dv.formula2 ? ' and ' + dv.formula2 : ''}`]);
  const cf = sheet.cf.filter((x) => x.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2)).length;
  if (cf) rows.push(['Conditional formats', String(cf)]);
  const note = sheet.comments.get(cellKey(r, c));
  if (note) rows.push([note.note ? 'Note' : 'Comment', note.text]);
  if (cell?.link) rows.push(['Link', cell.link]);
  return rows;
}

registerPane({
  id: 'info', title: 'Info Window',
  render(app, el) { el.dataset.sig = ''; this.update!(app, el, {}); },
  update(app, el) {
    const t = paneTarget(app, el);
    const sig = `${app.tab.id}:${app.book.version}:${app.engine.evaluations}:${t.sheet.id}:${t.r}:${t.c}:${el.dataset.pin || ''}`;
    if (el.dataset.sig === sig || typingIn(el)) return;
    el.dataset.sig = sig;
    clear(el);
    el.appendChild(h('div', { class: 'ql-row' }, pinToggle(app, el, () => this.update!(app, el, {}))));
    const table = h('table', { class: 'ql-info-table' });
    for (const [k, v] of cellInfoRows(app, t.sheet, t.r, t.c)) table.appendChild(h('tr', null, h('th', null, k), h('td', null, v)));
    el.appendChild(table);
  },
});

// What-If Sliders --------------------------------------------------------------------------------

registerPane({
  id: 'sliders', title: 'What-If Sliders',
  render(app, el, opts) {
    el.dataset.sig = '';
    const s = app.sheet;
    if (opts?.r != null) el.dataset.pin = `${opts.sheetId || s.id}:${opts.r}:${opts.c}`;
    this.update!(app, el, opts);
  },
  update(app, el) {
    const t = paneTarget(app, el);
    const out = el.querySelector<HTMLElement>('.ql-slider-out');
    const shown = displayValue(t.sheet, t.r, t.c);
    if (out && el.dataset.target === `${t.sheet.id}:${t.r}:${t.c}`) { out.textContent = shown; if (typingIn(el) || el.dataset.dragging) return; }
    const sig = `${app.tab.id}:${app.book.version}:${t.sheet.id}:${t.r}:${t.c}`;
    if (el.dataset.sig === sig || el.dataset.dragging) return;
    el.dataset.sig = sig;
    el.dataset.target = `${t.sheet.id}:${t.r}:${t.c}`;
    clear(el);
    const cell = t.sheet.get(t.r, t.c);
    if (cell?.f == null) { el.appendChild(h('p', { class: 'ql-help' }, 'Select a formula (a result such as a total or profit), then open What-If Sliders to get a slider for every number it depends on.')); return; }
    const lbl = labelFor(t.sheet, t.r, t.c);
    el.appendChild(h('div', { class: 'ql-help' }, `Result: ${addr(app, t.sheet, t.r, t.c)}${lbl ? ` (${lbl})` : ''}`));
    el.appendChild(h('div', { class: 'ql-slider-out', 'aria-live': 'polite' }, shown));
    const inputs = leafInputs(app.engine, t.sheet, t.r, t.c);
    if (!inputs.length) { el.appendChild(h('p', { class: 'ql-help' }, 'This formula does not depend on any number you typed in.')); return; }
    el.appendChild(h('p', { class: 'ql-help' }, 'Drag a slider to see the result change. Releasing it keeps the new value (Undo restores it).'));
    for (const x of inputs) {
      const v0 = x.sheet.value(x.r, x.c) as number;
      const rg = sliderRange(v0);
      const name = labelFor(x.sheet, x.r, x.c);
      const valEl = h('input', { class: 'ql-input', type: 'number', step: 'any', value: String(v0), style: 'width:7em', 'aria-label': `Value of ${addr(app, x.sheet, x.r, x.c)}` }) as HTMLInputElement;
      const slider = h('input', { type: 'range', min: String(rg.min), max: String(rg.max), step: String(rg.step), value: String(v0), 'aria-label': `Slider for ${name || addr(app, x.sheet, x.r, x.c)}` }) as HTMLInputElement;
      let original: Cell | undefined;
      const live = (n: number) => {
        if (original === undefined) original = x.sheet.get(x.r, x.c) ? { ...x.sheet.get(x.r, x.c)! } : undefined;
        el.dataset.dragging = '1';
        const cur = x.sheet.get(x.r, x.c);
        x.sheet.cells.set(cellKey(x.r, x.c), { ...(cur?.s ? { s: cur.s } : {}), v: n });
        app.book.noteChange(x.sheet, cellKey(x.r, x.c));
        app.engine.recalc();
        app.tab.cf.invalidate();
        app.grid.invalidate();
        const o = el.querySelector<HTMLElement>('.ql-slider-out');
        if (o) o.textContent = displayValue(t.sheet, t.r, t.c);
        valEl.value = String(n);
      };
      const commit = (n: number) => {
        delete el.dataset.dragging;
        const k = cellKey(x.r, x.c);
        if (original !== undefined || x.sheet.get(x.r, x.c)) {
          // Put the original back silently, then make the change as one undoable step.
          if (original) x.sheet.cells.set(k, original); else x.sheet.cells.delete(k);
        }
        original = undefined;
        if (app.isReadOnly()) { app.book.noteChange(x.sheet, k); app.engine.recalc(); app.refresh(); toast('This workbook is read-only.'); return; }
        app.book.transact(`What-If: ${cellName(x.r, x.c)}`, () => { const cur = x.sheet.get(x.r, x.c); x.sheet.set(x.r, x.c, { ...(cur?.s ? { s: cur.s } : {}), v: n }); });
      };
      slider.addEventListener('input', () => live(+slider.value));
      slider.addEventListener('change', () => commit(+slider.value));
      valEl.addEventListener('change', () => { const n = parseFloat(valEl.value); if (Number.isFinite(n)) { slider.value = String(n); live(n); commit(n); } });
      el.appendChild(h('div', { class: 'ql-slider' },
        h('div', { class: 'ql-slider-head' }, h('button', { class: 'ql-link', onclick: () => goTo(app, x.sheet, x.r, x.c) }, addr(app, x.sheet, x.r, x.c)), h('span', null, name)),
        slider, h('div', { class: 'ql-row' }, valEl, h('span', { class: 'ql-help' }, `${rg.min} to ${rg.max}`))));
    }
  },
});

// ============================================================================================
// Cell history

async function cellHistoryCmd(app: App): Promise<void> {
  const s = app.sheet;
  const { r, c } = activeCell(app);
  const list = cellHistory(app.book, s, r, c);
  const body = h('div', null);
  const now = s.get(r, c);
  body.appendChild(h('p', { class: 'ql-help' }, `Now: `, h('code', null, cellContentText(now)), now?.f != null ? ` = ${displayValue(s, r, c)}` : ''));
  if (!list.length) body.appendChild(h('p', { class: 'ql-help' }, 'No earlier contents in this session. Changes are remembered from when the workbook was opened, as far back as Undo goes.'));
  let done: (() => void) | null = null;
  list.forEach((e, i) => {
    body.appendChild(h('div', { class: 'ql-history-row' },
      h('div', null, h('div', null, h('code', null, cellContentText(e.before))), h('div', { class: 'ql-help' }, `${i === 0 ? 'Replaced by the latest change' : 'Replaced'} (${e.label}) with ${cellContentText(e.after)}`)),
      h('button', { class: 'ql-btn ql-sm', onclick: () => {
        if (app.isReadOnly()) { toast('This workbook is read-only.'); return; }
        if (!checkWritable(app, [{ r1: r, c1: c, r2: r, c2: c }])) return;
        app.book.transact('Restore Earlier Value', () => { const cur = s.get(r, c); s.set(r, c, e.before ? { ...e.before, ...(cur?.s != null ? { s: cur.s } : {}) } : cur?.s ? { s: cur.s } : undefined); });
        toast(`${cellName(r, c)} restored`);
        done?.();
      } }, 'Restore')));
  });
  await openDialog({ title: `History of ${cellName(r, c)}`, ok: false, cancel: 'Close', content: body, onOpen: (ctl) => { done = () => ctl.close(); } });
}

// ============================================================================================
// Chart slide show

function chartSlideShow(app: App): void {
  const slides: { sheet: Sheet; chart: ChartDef }[] = [];
  for (const sh of app.book.sheets) if (sh.visibility === 'visible') for (const ch of sh.charts) slides.push({ sheet: sh, chart: ch });
  if (!slides.length) { toast('This workbook has no charts. Insert ▸ Charts adds one.'); return; }
  let i = 0;
  const stage = h('div', { class: 'ql-slideshow-stage' });
  const cap = h('div', { class: 'ql-slideshow-cap' });
  const root = h('div', { class: 'ql-slideshow', role: 'dialog', 'aria-label': 'Chart slide show', tabindex: '-1' },
    stage, cap,
    h('button', { class: 'ql-slideshow-nav ql-slideshow-prev', 'aria-label': 'Previous chart', onclick: () => show(i - 1) }, '‹'),
    h('button', { class: 'ql-slideshow-nav ql-slideshow-next', 'aria-label': 'Next chart', onclick: () => show(i + 1) }, '›'),
    h('button', { class: 'ql-slideshow-close', onclick: () => close() }, 'Close (Esc)'));
  const show = (n: number) => {
    i = (n + slides.length) % slides.length;
    const { sheet, chart } = slides[i];
    const w = Math.min(window.innerWidth - 120, 1100), hgt = Math.min(window.innerHeight - 140, Math.round(w * 0.62));
    stage.innerHTML = renderChartSvg(app.tab, chart, Math.max(240, w), Math.max(160, hgt), sheet);
    cap.textContent = `${chart.title || 'Chart'} — ${sheet.name}  (${i + 1} of ${slides.length})`;
  };
  const onKey = (e: KeyboardEvent) => {
    if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter'].includes(e.key)) show(i + 1);
    else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(e.key)) show(i - 1);
    else if (e.key === 'Home') show(0);
    else if (e.key === 'End') show(slides.length - 1);
    else if (e.key === 'Escape') close();
    else return;
    e.preventDefault(); e.stopImmediatePropagation();
  };
  const close = () => { document.removeEventListener('keydown', onKey, true); window.removeEventListener('resize', onResize); root.remove(); app.focus(); };
  const onResize = () => show(i);
  stage.addEventListener('click', () => show(i + 1));
  document.addEventListener('keydown', onKey, true);
  window.addEventListener('resize', onResize);
  document.body.appendChild(root);
  root.focus();
  show(0);
}

// ============================================================================================
// Data Form

/** The list a record command works on: the current region with its header row. */
function listRegion(app: App): { rg: Range; header: boolean } | null {
  const rg = dataRegion(app);
  const s = app.sheet;
  if (!s.usedRange(false) || (rg.r1 === rg.r2 && rg.c1 === rg.c2 && s.value(rg.r1, rg.c1) == null)) return null;
  return { rg, header: hasHeaderRow(s, rg) || rg.r2 > rg.r1 };
}

async function dataForm(app: App): Promise<void> {
  const s = app.sheet;
  const lr = listRegion(app);
  if (!lr) { toast('Select a cell in a list with a header row first.'); return; }
  const rg = { ...lr.rg };
  const cols = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, j) => rg.c1 + j);
  const heads = cols.map((c) => { const v = s.value(rg.r1, c); return v == null || v === '' ? `Column ${colName(c)}` : String(v); });
  const first = rg.r1 + 1;
  let row = Math.min(Math.max(s.sel.r, first), Math.max(first, rg.r2));
  const count = () => Math.max(0, rg.r2 - first + 1);
  const fieldsFor = (): Field[] => cols.map((c, j) => (s.get(row, c)?.f != null
    ? { key: 'c' + j, label: `${heads[j]} (calculated): ${cellText(s, row, c)}`, type: 'info' }
    : { key: 'c' + j, label: heads[j], type: 'text', value: cellText(s, row, c, false) }));
  const setTitle = (ctl: DialogControl) => { const t = ctl.body.closest('.ql-dialog')?.querySelector('h2'); if (t) t.textContent = title(); };
  const title = () => (row > rg.r2 ? `New record (${count()} records)` : `Record ${row - first + 1} of ${count()}`);
  const save = (vals: Record<string, any>): boolean => {
    const changes: [number, string][] = [];
    cols.forEach((c, j) => {
      if (!(('c' + j) in vals)) return;
      const cur = s.get(row, c);
      if (cur?.f != null) return;
      const before = cellText(s, row, c, false);
      const now = String(vals['c' + j] ?? '');
      if (now !== before) changes.push([c, now]);
    });
    if (!changes.length) return false;
    if (app.isReadOnly()) { toast('This workbook is read-only.'); return false; }
    if (!checkWritable(app, changes.map(([c]) => ({ r1: row, c1: c, r2: row, c2: c })))) return false;
    app.book.transact('Data Form', () => {
      for (const [c, t] of changes) setInput(app, s, row, c, t);
      // A new record extends the list; formulas in the row above are carried down.
      if (row > rg.r2) {
        for (const c of cols) { const above = s.get(row - 1, c); if (above?.f != null && !s.get(row, c)?.f) s.set(row, c, { ...(above.s ? { s: above.s } : {}), f: shiftFormula(above.f, 1, 0) }); }
        rg.r2 = row;
      }
    });
    return true;
  };
  const go = (ctl: DialogControl, next: number) => {
    save(ctl.values());
    row = Math.max(first, Math.min(next, rg.r2 + 1));
    ctl.rebuild(fieldsFor());
    setTitle(ctl);
    app.grid.select(Math.min(row, rg.r2), cols[0]);
    app.grid.scrollIntoView(Math.min(row, rg.r2), cols[0]);
  };
  await openDialog({
    title: title(), ok: 'Close', cancel: false,
    fields: fieldsFor(),
    help: 'One record at a time. Changes are saved when you move to another record or close the form.',
    buttons: [
      { label: '◀ Previous', onClick: (_v, ctl) => { go(ctl, row - 1); } },
      { label: 'Next ▶', onClick: (_v, ctl) => { go(ctl, row + 1); } },
      { label: 'New', onClick: (_v, ctl) => { go(ctl, rg.r2 + 1); } },
      { label: 'Delete', onClick: async (_v, ctl) => {
        if (row > rg.r2) { go(ctl, rg.r2); return; }
        if (!(await confirmDialog('Delete Record', `Delete record ${row - first + 1}? The rows below move up.`, 'Delete'))) return;
        if (!checkWritable(app, [{ r1: row, c1: rg.c1, r2: rg.r2, c2: rg.c2 }])) return;
        app.book.transact('Data Form: Delete Record', () => insertCellsShift(s, { r1: row, c1: rg.c1, r2: row, c2: rg.c2 }, 'up'));
        rg.r2--;
        row = Math.min(row, Math.max(first, rg.r2));
        ctl.rebuild(fieldsFor());
        setTitle(ctl);
      } },
      { label: 'Find…', onClick: async (_v, ctl) => {
        const { promptDialog } = await import('../ui/dialog');
        const q = (await promptDialog('Find Record', 'Find the next record containing', ''))?.trim().toLowerCase();
        if (!q) return;
        for (let k = 1; k <= count(); k++) {
          const rr = first + ((row - first + k) % Math.max(1, count()));
          if (cols.some((c) => cellText(s, rr, c).toLowerCase().includes(q))) { go(ctl, rr); return; }
        }
        toast(`No record contains "${q}".`);
      } },
    ],
  }).then((res) => { if (res) save(res); });
}

// ============================================================================================
// Matrix (1-2-3 /Data Matrix)

type MatrixOp = 'invert' | 'multiply' | 'solve' | 'determinant';

function readMatrix(sheet: Sheet, rg: Range): Mat | string {
  const m: Mat = [];
  for (let r = rg.r1; r <= rg.r2; r++) {
    const row: number[] = [];
    for (let c = rg.c1; c <= rg.c2; c++) {
      const v = sheet.value(r, c);
      if (typeof v !== 'number') return `${cellName(r, c)} is not a number.`;
      row.push(v);
    }
    m.push(row);
  }
  return m;
}

/** Runs a matrix operation and writes the result as values (or as live formulas) at the output cell. */
export function matrixOp(app: App, o: { op: MatrixOp; a: string; b?: string; out: string; formulas?: boolean }): string | null {
  const book = app.book;
  const A = resolveRangeText(book, o.a, app.sheet);
  if (!A) return `"${o.a}" is not a range.`;
  const B = o.b ? resolveRangeText(book, o.b, app.sheet) : null;
  if ((o.op === 'multiply' || o.op === 'solve') && !B) return `"${o.b ?? ''}" is not a range.`;
  const out = resolveRangeText(book, o.out, app.sheet);
  if (!out) return `"${o.out}" is not a cell.`;
  const ma = readMatrix(A.sheet, A.range);
  if (typeof ma === 'string') return ma;
  const mb = B ? readMatrix(B.sheet, B.range) : null;
  if (typeof mb === 'string') return mb;
  let res: Mat | null = null;
  let formula = '';
  const ra = rangeRefText(A.sheet, A.range, true, out.sheet), rb = B ? rangeRefText(B.sheet, B.range, true, out.sheet) : '';
  switch (o.op) {
    case 'invert': res = matInverse(ma); formula = `MINVERSE(${ra})`; if (!res) return ma.length !== ma[0]?.length ? 'The matrix must be square.' : 'The matrix is singular (it has no inverse).'; break;
    case 'multiply': res = matMultiply(ma, mb!); formula = `MMULT(${ra},${rb})`; if (!res) return 'The first matrix needs as many columns as the second has rows.'; break;
    case 'solve': res = matSolve(ma, mb!); formula = `MMULT(MINVERSE(${ra}),${rb})`; if (!res) return 'The coefficients must be square and invertible, with as many rows as the constants.'; break;
    case 'determinant': { const d = matDeterminant(ma); if (d == null) return 'The matrix must be square.'; res = [[d]]; formula = `MDETERM(${ra})`; break; }
  }
  const r0 = out.range.r1, c0 = out.range.c1;
  const target: Range = { r1: r0, c1: c0, r2: r0 + res!.length - 1, c2: c0 + res![0].length - 1 };
  if (!checkWritable(app, [target], out.sheet)) return '';
  book.transact(`Matrix ${o.op[0].toUpperCase() + o.op.slice(1)}`, () => {
    if (o.formulas) { writeOut(out.sheet, r0, c0, { f: formula }); return; }
    res!.forEach((row, i) => row.forEach((v, j) => writeOut(out.sheet, r0 + i, c0 + j, +v.toPrecision(15))));
  });
  return null;
}

async function matrixCmd(app: App, arg?: MatrixOp | { op: MatrixOp; a: string; b?: string; out: string; formulas?: boolean }): Promise<void> {
  if (arg && typeof arg === 'object') { const e = matrixOp(app, arg); if (e) toast(e); return; }
  const op: MatrixOp = (arg as MatrixOp) || 'invert';
  const sel = clipToUsed(app.sheet, activeRange(app));
  const two = op === 'multiply' || op === 'solve';
  const defOut = cellName(sel.r1, sel.c2 + 2);
  const labels: Record<MatrixOp, [string, string]> = { invert: ['Matrix to invert', ''], multiply: ['First matrix', 'Second matrix'], solve: ['Coefficients (square matrix A)', 'Constants (b, one column per system)'], determinant: ['Square matrix', ''] };
  const res = await openDialog({
    title: { invert: 'Invert Matrix', multiply: 'Multiply Matrices', solve: 'Solve Linear Equations', determinant: 'Matrix Determinant' }[op],
    help: op === 'solve' ? 'Finds x in A·x = b and writes it at the output cell.' : 'The result is written at the output cell.',
    fields: [
      { key: 'a', label: labels[op][0], type: 'text', value: rangeName(sel) },
      ...(two ? [{ key: 'b', label: labels[op][1], type: 'text' as const, value: '' }] : []),
      { key: 'out', label: 'Output (top-left cell)', type: 'text', value: defOut },
      { key: 'formulas', label: 'Insert a live formula instead of values', type: 'checkbox', value: false },
    ],
  });
  if (!res) return;
  const e = matrixOp(app, { op, a: res.a, b: res.b, out: res.out, formulas: !!res.formulas });
  if (e) toast(e); else if (e === null) toast('Done');
}

// ============================================================================================
// Monte Carlo

export interface McSpec { inputs: { ref: string; dist: string }[]; outputs: string[]; iterations: number; seed?: number | null }

/** Runs a simulation and adds a results sheet. Returns an error message, or null. */
export function runMonteCarlo(app: App, spec: McSpec): string | null {
  const book = app.book;
  const inputs: { sheet: Sheet; r: number; c: number; dist: Dist }[] = [];
  for (const x of spec.inputs) {
    const p = resolveRangeText(book, x.ref, app.sheet);
    if (!p) return `"${x.ref}" is not a cell.`;
    const d = parseDist(x.dist);
    if (!d) return `"${x.dist}" is not a distribution. Use "10 to 20", "normal 100 15", "triangular 1 3 8" or "1, 2, 5".`;
    if (p.sheet.get(p.range.r1, p.range.c1)?.f != null) return `${x.ref} holds a formula; inputs must be constants.`;
    inputs.push({ sheet: p.sheet, r: p.range.r1, c: p.range.c1, dist: d });
  }
  const outputs: { sheet: Sheet; r: number; c: number; name: string }[] = [];
  for (const o of spec.outputs) {
    const p = resolveRangeText(book, o, app.sheet);
    if (!p) return `"${o}" is not a cell.`;
    if (p.sheet.get(p.range.r1, p.range.c1)?.f == null) return `${o} is not a formula; outputs must be results.`;
    outputs.push({ sheet: p.sheet, r: p.range.r1, c: p.range.c1, name: `${o.toUpperCase()}${labelFor(p.sheet, p.range.r1, p.range.c1) ? ` ${labelFor(p.sheet, p.range.r1, p.range.c1)}` : ''}` });
  }
  if (!outputs.length) return 'Choose at least one output.';
  const n = Math.max(10, Math.min(100000, Math.round(spec.iterations || 1000)));
  const rnd = spec.seed != null && Number.isFinite(spec.seed) ? seeded(spec.seed) : Math.random;
  const { values, errors } = monteCarlo(book, app.engine, inputs, outputs, n, rnd);
  const sums = values.map((v, j) => summarize(v, errors[j]));
  const stats: [string, (x: ReturnType<typeof summarize>) => number, string][] = [
    ['Mean', (x) => x.mean, '#,##0.00'], ['Standard deviation', (x) => x.sd, '#,##0.00'], ['Minimum', (x) => x.min, '#,##0.00'],
    ['5th percentile', (x) => x.p5, '#,##0.00'], ['25th percentile', (x) => x.p25, '#,##0.00'], ['Median', (x) => x.p50, '#,##0.00'],
    ['75th percentile', (x) => x.p75, '#,##0.00'], ['95th percentile', (x) => x.p95, '#,##0.00'], ['Maximum', (x) => x.max, '#,##0.00'],
    ['Runs', (x) => x.n, '0'], ['Runs with errors', (x) => x.errors, '0'],
  ];
  const rows: any[][] = [[{ v: 'Statistic', style: HEAD }, ...outputs.map((o) => ({ v: o.name, style: HEAD }))]];
  for (const [label, f, fmt] of stats) rows.push([{ v: label, style: BOLD }, ...sums.map((x) => (Number.isFinite(f(x)) ? { v: f(x), style: { numFmt: fmt } } : '#N/A'))]);
  rows.push([]);
  rows.push([{ v: 'Inputs', style: HEAD }, { v: 'Distribution', style: HEAD }]);
  spec.inputs.forEach((x) => rows.push([x.ref.toUpperCase(), x.dist]));
  const histStart = rows.length + 1;
  // Histogram of the first output (20 bins).
  const first = values[0];
  const hist: number[] = [];
  let lo = 0, width = 0;
  if (first.length) {
    lo = Math.min(...first); const hi = Math.max(...first);
    width = (hi - lo) / 20 || 1;
    for (let i = 0; i < 20; i++) hist.push(0);
    for (const v of first) hist[Math.min(19, Math.floor((v - lo) / width))]++;
    rows.push([]);
    rows.push([{ v: `Distribution of ${outputs[0].name} (bin up to)`, style: HEAD }, { v: 'Runs', style: HEAD }]);
    hist.forEach((cnt, i) => rows.push([{ v: lo + width * (i + 1), style: { numFmt: '#,##0.##' } }, cnt]));
  }
  const valuesCol = outputs.length + 10; // right of the chart
  addSheetWithRows(app, 'Simulation', rows, { colW: [190, ...outputs.map(() => 120)], freeze: 1, after: (sh) => {
    // Every run's results, so they can be analyzed further.
    writeOut(sh, 0, valuesCol, { v: 'Run', style: HEAD });
    outputs.forEach((o, j) => writeOut(sh, 0, valuesCol + 1 + j, { v: o.name, style: HEAD }));
    const len = Math.max(...values.map((v) => v.length));
    for (let i = 0; i < Math.min(len, 100000); i++) { writeOut(sh, i + 1, valuesCol, i + 1); values.forEach((v, j) => { if (i < v.length) writeOut(sh, i + 1, valuesCol + 1 + j, v[i]); }); }
    if (hist.length) {
      const top = histStart + 1;
      const chart: ChartDef = { id: newId('ch'), type: 'column', title: `Simulated ${outputs[0].name}`, legend: 'none', gridlines: true,
        series: [{ name: 'Runs', values: rangeRefText(sh, { r1: top, c1: 1, r2: top + 19, c2: 1 }), categories: rangeRefText(sh, { r1: top, c1: 0, r2: top + 19, c2: 0 }) }],
        at: { r: 1, c: outputs.length + 2, dx: 0, dy: 0, w: 440, h: 270 }, sourceRange: rangeRefText(sh, { r1: top - 1, c1: 0, r2: top + 19, c2: 1 }) };
      sh.charts.push(chart);
    }
  } });
  toast(`${n} runs simulated`);
  return null;
}

async function monteCarloCmd(app: App, arg?: McSpec): Promise<void> {
  if (arg) { const e = runMonteCarlo(app, arg); if (e) toast(e); return; }
  const s = app.sheet;
  const sel = selRanges(app);
  const outs: string[] = [];
  for (const rg of sel) s.forEachIn(clipToUsed(s, rg), (r, c, cell) => { if (cell.f != null && outs.length < 10) outs.push(cellName(r, c)); });
  const { r, c } = activeCell(app);
  if (!outs.length && s.get(r, c)?.f != null) outs.push(cellName(r, c));
  const base = outs.length ? resolveRangeText(app.book, outs[0], s) : null;
  const ins = base ? leafInputs(app.engine, base.sheet, base.range.r1, base.range.c1, 12) : [];
  const lines = ins.map((x) => { const v = x.sheet.value(x.r, x.c) as number; const a = +(v * 0.8).toPrecision(6), b = +(v * 1.2).toPrecision(6); return `${addr(app, x.sheet, x.r, x.c)}: ${Math.min(a, b)} to ${Math.max(a, b)}`; });
  const res = await openDialog({
    title: 'Monte Carlo Simulation', wide: true,
    help: 'Give uncertain inputs a range instead of one number, run the model many times, and see the spread of results. Formulas with RAND() vary too.',
    fields: [
      { key: 'inputs', label: 'Uncertain inputs (one per line: cell: distribution)', type: 'textarea', rows: 6, value: lines.join('\n'), placeholder: 'B2: 10 to 20\nB3: normal 100 15\nB4: triangular 1 3 8\nB5: 1, 2, 5', help: '"a to b" uniform · "normal mean sd" · "triangular min mode max" · "x, y, z" one of the listed values' },
      { key: 'outputs', label: 'Results to track (cells with formulas, comma-separated)', type: 'text', value: outs.join(', ') },
      { key: 'iterations', label: 'Runs', type: 'number', value: 1000, min: 10, max: 100000, half: true },
      { key: 'seed', label: 'Seed (blank = random)', type: 'number', value: '', half: true },
    ],
    validate: (v) => (!String(v.outputs || '').trim() ? 'Enter at least one result cell.' : null),
  });
  if (!res) return;
  const inputs = String(res.inputs || '').split(/\n/).map((l) => l.trim()).filter(Boolean).map((l) => { const i = l.indexOf(':'); return i < 0 ? { ref: l, dist: '' } : { ref: l.slice(0, i).trim(), dist: l.slice(i + 1).trim() }; });
  const e = runMonteCarlo(app, { inputs, outputs: String(res.outputs).split(/[,;]/).map((x) => x.trim()).filter(Boolean), iterations: Number(res.iterations) || 1000, seed: res.seed === null || res.seed === '' ? null : Number(res.seed) });
  if (e) toast(e);
}

// ============================================================================================
// Records: Delete Matching Records, Categorize, Column Type

/** Deletes the records (data rows of a list) that match a condition, closing up the gaps. Returns the number removed. */
export function deleteMatching(app: App, rg: Range, col: number, op: CondOp, value: string): number {
  const s = app.sheet;
  const first = rg.r1 + 1;
  const keep: number[] = [];
  let removed = 0;
  for (let r = first; r <= rg.r2; r++) { if (matchesCondition(s.value(r, col), op, value)) removed++; else keep.push(r); }
  if (!removed) return 0;
  app.book.transact('Delete Matching Records', () => {
    const rows = keep.map((r) => Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, j) => s.get(r, rg.c1 + j)));
    for (let r = first; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) s.set(r, c, undefined);
    rows.forEach((cells, i) => cells.forEach((cell, j) => {
      if (!cell) return;
      const to = first + i;
      s.set(to, rg.c1 + j, cell.f != null ? { ...cell, f: shiftFormula(cell.f, to - keep[i], 0) } : { ...cell });
    }));
  });
  return removed;
}

async function queryDelete(app: App, arg?: { column: number; op: CondOp; value?: string }): Promise<void> {
  const lr = listRegion(app);
  if (!lr) { toast('Select a cell in a list with a header row first.'); return; }
  const s = app.sheet, rg = lr.rg;
  if (!checkWritable(app, [rg])) return;
  let o = arg;
  if (!o) {
    const cols = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, j) => ({ value: String(rg.c1 + j), label: String(s.value(rg.r1, rg.c1 + j) ?? `Column ${colName(rg.c1 + j)}`) }));
    const res = await openDialog({
      title: 'Delete Matching Records',
      help: `Removes the records of ${rangeName(rg)} that match, and closes up the gaps. Undo brings them back.`,
      fields: [
        { key: 'column', label: 'Field', type: 'select', value: String(Math.min(Math.max(s.sel.c, rg.c1), rg.c2)), options: cols },
        { key: 'op', label: 'Condition', type: 'select', value: 'eq', options: (Object.keys(COND_LABELS) as CondOp[]).map((k) => ({ value: k, label: COND_LABELS[k] })) },
        { key: 'value', label: 'Value', type: 'text', value: '', visibleIf: (v) => !['blank', 'nonblank', 'error'].includes(v.op) },
      ],
      onChange: (v, ctl) => {
        const c = +v.column; let n = 0;
        for (let r = rg.r1 + 1; r <= rg.r2; r++) if (matchesCondition(s.value(r, c), v.op, String(v.value ?? ''))) n++;
        const help = ctl.body.querySelector('.ql-help');
        if (help) help.textContent = `${n} of ${rg.r2 - rg.r1} records in ${rangeName(rg)} match. They are removed and the gaps closed up. Undo brings them back.`;
      },
    });
    if (!res) return;
    o = { column: +res.column, op: res.op, value: res.value };
  }
  const n = deleteMatching(app, rg, o.column, o.op, String(o.value ?? ''));
  toast(n ? `${n} record${n === 1 ? '' : 's'} deleted` : 'No records match.');
}

async function categorizeCmd(app: App, arg?: { rules: string; ifNone?: string; values?: boolean }): Promise<void> {
  const s = app.sheet;
  const sel = activeRange(app);
  const lr = listRegion(app);
  const rg = sel.r1 === sel.r2 && sel.c1 === sel.c2 && lr ? { ...lr.rg, c1: sel.c1, c2: sel.c1 } : clipToUsed(s, sel);
  const header = hasHeaderRow(s, lr?.rg ?? rg) && rg.r1 === (lr?.rg.r1 ?? rg.r1);
  const r1 = header ? rg.r1 + 1 : rg.r1;
  const col = rg.c1, out = rg.c2 + 1;
  let o = arg;
  if (!o) {
    const res = await openDialog({
      title: 'Categorize by Rules',
      help: `Gives each row of ${colName(col)}${r1 + 1}:${colName(col)}${rg.r2 + 1} the category of the first keyword it contains, in column ${colName(out)}. The rules are a two-column range: keywords (comma-separated alternatives, * and ? wildcards) and the category.`,
      fields: [
        { key: 'rules', label: 'Rules range (keywords | category)', type: 'text', value: '', placeholder: "Rules!A2:B20" },
        { key: 'ifNone', label: 'Category when nothing matches', type: 'text', value: 'Other' },
        { key: 'values', label: 'Write values instead of live CATEGORIZE formulas', type: 'checkbox', value: false },
      ],
      validate: (v) => { const x = resolveRangeText(app.book, String(v.rules || ''), s); return x && x.range.c2 === x.range.c1 + 1 ? null : 'Enter a two-column range of rules, such as Rules!A2:B20.'; },
    });
    if (!res) return;
    o = { rules: res.rules, ifNone: res.ifNone, values: !!res.values };
  }
  const rr = resolveRangeText(app.book, o.rules, s);
  if (!rr || rr.range.c2 !== rr.range.c1 + 1) { toast('The rules must be a two-column range.'); return; }
  const target: Range = { r1: header ? rg.r1 : r1, c1: out, r2: rg.r2, c2: out };
  let occupied = false;
  s.forEachIn({ ...target, r1: r1 }, (_r, _c, cell) => { if (cell.v != null || cell.f != null) occupied = true; });
  if (occupied && !(await confirmDialog('Replace Contents?', `Column ${colName(out)} already has data. Replace it?`, 'Replace'))) return;
  if (!checkWritable(app, [target])) return;
  const kw = rangeRefText(rr.sheet, { ...rr.range, c2: rr.range.c1 }, true, s), cat = rangeRefText(rr.sheet, { ...rr.range, c1: rr.range.c2 }, true, s);
  const none = JSON.stringify(o.ifNone ?? '');
  app.book.transact('Categorize by Rules', () => {
    if (header && s.value(rg.r1, out) == null) writeOut(s, rg.r1, out, { v: 'Category', style: s.styleOf(rg.r1, col).bold ? { bold: true } : undefined });
    for (let r = r1; r <= rg.r2; r++) writeOut(s, r, out, { f: `CATEGORIZE(${cellName(r, col)},${kw},${cat},${none})` });
  });
  if (o.values) {
    app.engine.recalc(true);
    app.book.transact('Categorize by Rules', () => { for (let r = r1; r <= rg.r2; r++) { const v = s.value(r, out); writeOut(s, r, out, v); } });
  }
  toast(`${rg.r2 - r1 + 1} rows categorized`);
}

/** Validation, number format and style for a typed column. */
export function columnTypeSpec(type: ColumnType, first: { r: number; c: number }, choices: string[] = []): { dv?: Omit<DataValidation, 'id' | 'ranges'>; style: Partial<CellStyle> } {
  const a = cellName(first.r, first.c);
  const stop = (title: string, msg: string) => ({ errorStyle: 'stop' as const, errorTitle: title, errorMessage: msg, allowBlank: true });
  switch (type) {
    case 'text': return { style: { numFmt: '@' } };
    case 'number': return { dv: { type: 'custom', formula1: `ISNUMBER(${a})`, ...stop('Number column', 'Enter a number.') }, style: { numFmt: 'General' } };
    case 'integer': return { dv: { type: 'whole', operator: 'between', formula1: '-1E+15', formula2: '1E+15', ...stop('Whole number column', 'Enter a whole number.') }, style: { numFmt: '0' } };
    case 'currency': return { dv: { type: 'custom', formula1: `ISNUMBER(${a})`, ...stop('Currency column', 'Enter an amount.') }, style: { numFmt: '$#,##0.00' } };
    case 'percent': return { dv: { type: 'custom', formula1: `ISNUMBER(${a})`, ...stop('Percent column', 'Enter a percentage, such as 15%.') }, style: { numFmt: '0%' } };
    case 'date': return { dv: { type: 'custom', formula1: `ISNUMBER(${a})`, ...stop('Date column', 'Enter a date, such as 2024-03-15.') }, style: { numFmt: 'yyyy-mm-dd' } };
    case 'checkbox': return { dv: { type: 'custom', formula1: `ISLOGICAL(${a})`, ...stop('Checkbox column', 'Enter TRUE or FALSE, or press Space to toggle.') }, style: { checkbox: true, hAlign: 'center' } };
    case 'choice': return { dv: { type: 'list', formula1: `"${choices.map((x) => x.replace(/[",]/g, ' ')).join(',')}"`, showDropdown: true, ...stop('Choice column', 'Pick a value from the list.') }, style: {} };
    case 'email': return { dv: { type: 'custom', formula1: `ISEMAIL(${a})`, ...stop('Email column', 'Enter an email address, such as name@example.com.') }, style: {} };
    case 'url': return { dv: { type: 'custom', formula1: `ISURL(${a})`, ...stop('Web address column', 'Enter a web address, such as https://example.com.') }, style: {} };
    case 'rating': return { dv: { type: 'whole', operator: 'between', formula1: '1', formula2: '5', ...stop('Rating column', 'Enter a whole number from 1 to 5.') }, style: { numFmt: '0" ★"', hAlign: 'center' } };
  }
}

/** Applies a column type to the data cells under the header of each selected column. Returns the number of values that do not fit. */
export function applyColumnType(app: App, type: ColumnType): { columns: number; mismatches: number; range: Range } | string {
  const s = app.sheet;
  const sel = activeRange(app);
  const table = s.tables.find((t) => sel.r1 >= t.range.r1 && sel.r1 <= t.range.r2 && sel.c1 >= t.range.c1 && sel.c1 <= t.range.c2);
  let r1: number, r2: number;
  if (table) { r1 = table.range.r1 + (table.opts.header ? 1 : 0); r2 = table.range.r2 - (table.opts.totals ? 1 : 0); }
  else {
    const lr = listRegion(app);
    if (!lr) return 'Select a column of a list (or table) first.';
    r1 = lr.rg.r1 + (hasHeaderRow(s, lr.rg) ? 1 : 0);
    // Typed columns apply to rows added later too.
    r2 = Math.max(lr.rg.r2, r1 + 999);
  }
  const c1 = Math.max(sel.c1, 0), c2 = Math.min(sel.c2, c1 + 50);
  const target: Range = { r1, c1, r2, c2 };
  if (!checkWritable(app, [target])) return '';
  let mismatches = 0;
  app.book.transact(`Column Type: ${COLUMN_TYPES.find((t) => t.id === type)?.label}`, () => {
    app.book.recordSheet(s);
    for (const dv of s.validations) dv.ranges = dv.ranges.flatMap((x) => subtractRange(x, target));
    s.validations = s.validations.filter((dv) => dv.ranges.length);
    for (let c = c1; c <= c2; c++) {
      const col: Range = { r1, c1: c, r2, c2: c };
      const values: Scalar[] = [];
      for (let r = r1; r <= r2; r++) values.push(s.value(r, c));
      const choices = type === 'choice' ? distinctChoices(values) : [];
      const spec = columnTypeSpec(type, { r: r1, c }, choices);
      if (spec.dv && !(type === 'choice' && !choices.length)) s.validations.push({ id: newId('dv'), ranges: [col], ...spec.dv } as DataValidation);
      for (let r = r1; r <= r2; r++) {
        const cur = s.get(r, c);
        const sid = app.book.derivedStyle(s.styleIdOf(r, c), { ...spec.style, ...(type !== 'checkbox' ? { checkbox: undefined } : {}) });
        if (type === 'checkbox' && (cur?.v == null || cur.v === '') && cur?.f == null && r <= Math.max(r1, (clipToUsed(s, col).r2))) s.set(r, c, { s: sid, v: false });
        else if (cur || sid) s.set(r, c, { ...(cur || {}), s: sid || undefined });
      }
      mismatches += typeMismatches(values, type, choices).length;
    }
  });
  return { columns: c2 - c1 + 1, mismatches, range: target };
}

// ============================================================================================
// Commands

const toggle = (app: App, key: 'lotusFormulas' | 'lotusLabels' | 'slashKey', on: string, off: string) => {
  app.prefs[key] = !app.prefs[key];
  app.savePrefs();
  toast(app.prefs[key] ? on : off);
  app.refresh();
};

def([
  // ---------------------------------------------------------------- Tools ▸ Classic
  {
    id: 'tools.slashMenu', label: 'Slash Menu', icon: 'menu_open', tab: T, group: 'Classic', keys: ['Alt-/'], edits: false, noRecord: true,
    tip: 'Every command as a keyboard menu: type the highlighted letters (for example D S A for Data ▸ Sort & Filter ▸ Sort A to Z), Esc to go back',
    keywords: 'slash menu keyboard lotus 123 commands accelerators', origin: 'Lotus 1-2-3 / command menu',
    run: (app) => openSlashMenu(app),
  },
  {
    id: 'tools.slashKey', label: '/ Key Opens Menu', icon: 'keyboard_command_key', tab: T, group: 'Classic', kind: 'toggle', edits: false, noRecord: true,
    tip: 'Pressing / on a selected cell opens the Slash Menu instead of typing a slash (type \'/ to enter a slash as text)',
    keywords: 'slash key lotus transition menu key', origin: 'Lotus 1-2-3 (and the menu-key option spreadsheets offered 1-2-3 users)',
    isActive: (app) => app.prefs.slashKey,
    run: (app) => toggle(app, 'slashKey', '/ now opens the Slash Menu', '/ types a slash again'),
  },
  {
    id: 'tools.lotusEntry', label: '1-2-3 Formula Entry', icon: 'alternate_email', tab: T, group: 'Classic', kind: 'toggle', edits: false, noRecord: true,
    tip: 'Accept Lotus 1-2-3 formulas as you type: @SUM(A1..A5), +A1*@AVG(B1..B3), #AND#, #OR#, #NOT#, with 1-2-3\'s meanings (zero-based offsets, argument order, precedence)',
    keywords: 'lotus 123 @ functions at sign formula entry transition', origin: 'Lotus 1-2-3 @functions; transition formula entry in later spreadsheets',
    isActive: (app) => app.prefs.lotusFormulas,
    run: (app) => toggle(app, 'lotusFormulas', '1-2-3 formulas are accepted', '1-2-3 formula entry is off'),
  },
  {
    id: 'tools.lotusLabels', label: '1-2-3 Label Prefixes', icon: 'format_quote', tab: T, group: 'Classic', kind: 'toggle', edits: false, noRecord: true,
    tip: 'Typing "Total right-aligns a label, ^Title centers it and \\- repeats it across the cell (\' is always left-aligned text)',
    keywords: 'label prefix caret backslash repeat align lotus', origin: 'Lotus 1-2-3 label prefixes',
    isActive: (app) => app.prefs.lotusLabels,
    run: (app) => toggle(app, 'lotusLabels', 'Label prefixes " ^ \\ are on', 'Label prefixes are off'),
  },
  {
    id: 'tools.convertLotus', label: 'Convert 1-2-3 Formulas', icon: 'transform', tab: T, group: 'Classic', edits: true,
    tip: 'Turn text cells that hold Lotus 1-2-3 formulas (pasted or imported as text) into working formulas',
    keywords: 'lotus 123 convert import formulas text @sum', origin: 'Lotus 1-2-3',
    run: (app) => {
      if (!checkWritable(app, selRanges(app))) return;
      const n = convertLotusCells(app, workRanges(app));
      toast(n ? `${n} formula${n === 1 ? '' : 's'} converted` : `No 1-2-3 formulas found. They look like @SUM(A1..A5) or +B2*@AVG(C1..C9); ${LOTUS_FUNCTIONS.length} @functions are understood.`);
    },
  },
  {
    id: 'tools.exportCode', label: 'Export as Code', icon: 'code', tab: T, group: 'Classic', kind: 'split', edits: false, noRecord: true,
    tip: 'The sheet as a Python or JavaScript program: inputs become variables and formulas become code, in calculation order',
    keywords: 'python javascript program code export resolver model', origin: 'Resolver One (every spreadsheet is a Python program)',
    run: (app, arg?: CodeLang) => exportCode(app, arg),
    menu: () => [{ id: 'tools.exportCode', arg: 'python', label: 'As Python…', icon: 'code' }, { id: 'tools.exportCode', arg: 'js', label: 'As JavaScript…', icon: 'javascript' }],
  },
  // ---------------------------------------------------------------- Formulas ▸ Insight
  {
    id: 'formulas.explainValue', label: 'Explain Value', icon: 'account_tree', tab: F, group: 'Insight', edits: false, noRecord: true, popular: true,
    tip: 'How the selected result is built: its formula, the values it reads and their formulas, as an outline you can expand',
    keywords: 'explain why value breakdown trace audit where does this number come from', origin: 'Javelin\'s model views; explainable models in Causal',
    run: (app) => app.ui.togglePane('explain'),
  },
  {
    id: 'formulas.dependencyDiagram', label: 'Dependency Diagram', icon: 'hub', tab: F, group: 'Insight', edits: false, noRecord: true,
    tip: 'A diagram of what feeds the selected cell and what it feeds, up to four steps each way',
    keywords: 'diagram graph dependencies precedents dependents flow', origin: 'Javelin diagram view',
    run: (app) => app.ui.togglePane('depDiagram'),
  },
  {
    id: 'formulas.formulaRules', label: 'Formula Rules', icon: 'rule', tab: F, group: 'Insight', edits: false, noRecord: true,
    tip: 'Each distinct formula on the sheet once, with the ranges it fills, and the cells whose formula breaks the pattern of their neighbors',
    keywords: 'formulas list distinct r1c1 inconsistent audit rules', origin: 'Lotus Improv formula pane; Multiplan R1C1 notation',
    run: (app) => app.ui.togglePane('formulaRules'),
  },
  {
    id: 'formulas.selectInconsistent', label: 'Select Inconsistent Formulas', icon: 'rule_settings', tab: F, group: 'Insight', edits: false, noRecord: true,
    tip: 'Select the formulas that differ from the matching formulas on both sides of them (the odd one out in a copied block)',
    keywords: 'inconsistent formula odd different neighbors audit', origin: 'Multiplan R1C1 notation (copies of a formula read identically)',
    run: (app) => {
      const odd = inconsistentFormulas(app.sheet);
      if (!odd.length) { toast('No inconsistent formulas on this sheet.'); return; }
      const s = app.sheet;
      s.sel = { r: odd[0].r, c: odd[0].c, ranges: odd.map((x) => ({ r1: x.r, c1: x.c, r2: x.r, c2: x.c })) };
      app.grid.scrollIntoView(odd[0].r, odd[0].c);
      app.grid.invalidate();
      toast(`${odd.length} inconsistent formula${odd.length === 1 ? '' : 's'} selected`);
    },
  },
  {
    id: 'formulas.infoWindow', label: 'Info Window', icon: 'info', tab: F, group: 'Insight', edits: false, noRecord: true,
    tip: 'Everything about the active cell in one place: value, formula (also in R1C1), format, protection, names, precedents, dependents, validation and notes',
    keywords: 'info cell information properties show info', origin: 'Excel for Macintosh Info window',
    run: (app) => app.ui.togglePane('info'),
  },
  {
    id: 'formulas.cellHistory', label: 'Cell History', icon: 'history', tab: F, group: 'Insight', edits: false, noRecord: true,
    tip: 'Earlier contents of the active cell in this session, with the change that replaced each, and Restore',
    keywords: 'history previous value edits changes restore cell', origin: 'Cell edit history in online spreadsheets',
    run: (app) => cellHistoryCmd(app),
  },
  // ---------------------------------------------------------------- View
  {
    id: 'view.mapView', label: 'Map View', icon: 'grid_view', tab: V, group: 'Show', kind: 'toggle', edits: false, noRecord: true,
    tip: 'Show what each cell holds instead of its value: " text, # number, + formula, ~ spilled, ! error, ≠ a formula unlike its neighbors',
    keywords: 'map view cell types overview audit labels values formulas', origin: 'Lotus 1-2-3 Release 3 map view',
    isActive: (app) => app.tab.mapView,
    run: (app) => {
      app.tab.mapView = !app.tab.mapView;
      app.grid.invalidate();
      if (app.tab.mapView) toast(Object.values(MAP_STYLE).map((x) => `${x.glyph} ${x.label}`).join(' · '), { ms: 6000 });
    },
  },
  {
    id: 'view.chartSlideShow', label: 'Chart Slide Show', icon: 'slideshow', tab: V, group: 'Workbook Views', edits: false, noRecord: true,
    tip: 'Show every chart in the workbook full screen, one at a time (arrow keys or click to advance, Esc to stop)',
    keywords: 'slide show present charts graphs full screen presentation', origin: 'Quattro Pro slide shows; 1-2-3 /Graph View',
    run: (app) => chartSlideShow(app),
  },
  // ---------------------------------------------------------------- Data
  {
    id: 'data.dataForm', label: 'Data Form', icon: 'dynamic_form', tab: D, group: 'Records', edits: true,
    tip: 'View, edit, add, find and delete the records of a list one at a time, as a form',
    keywords: 'form record entry data form edit rows', origin: 'Lotus Symphony FORM window; 1-2-3 /Range Input',
    run: (app) => dataForm(app),
  },
  {
    id: 'data.columnType', label: 'Column Type', icon: 'view_column', tab: D, group: 'Records', kind: 'menu', edits: true,
    tip: 'Give the selected list columns a type (number, currency, date, checkbox, choice, email, web address, rating) that formats entries and rejects ones that do not fit',
    keywords: 'typed column field type airtable grist validation dropdown checkbox email', origin: 'Airtable and Grist field types',
    run: async (app, arg?: ColumnType) => {
      if (!arg) {
        const res = await openDialog({ title: 'Column Type', help: 'Applies to the data cells of the selected columns, below the header.', fields: [{ key: 't', label: 'Type', type: 'select', value: 'number', options: COLUMN_TYPES.map((t) => ({ value: t.id, label: t.label })) }] });
        if (!res) return;
        arg = res.t as ColumnType;
      }
      const r = applyColumnType(app, arg);
      if (typeof r === 'string') { if (r) toast(r); return; }
      const label = COLUMN_TYPES.find((t) => t.id === arg)?.label;
      if (r.mismatches) { toast(`${label}: ${r.mismatches} existing value${r.mismatches === 1 ? '' : 's'} do not fit and are circled`); app.exec('data.circleInvalid'); }
      else toast(`${r.columns} column${r.columns === 1 ? '' : 's'} set to ${label}`);
    },
    menu: () => COLUMN_TYPES.map((t) => ({ id: 'data.columnType', arg: t.id, label: t.label, icon: t.icon })),
  },
  {
    id: 'data.categorize', label: 'Categorize by Rules', icon: 'label', tab: D, group: 'Records', edits: true,
    tip: 'Fill a category column from keyword rules (the first keyword found in the text decides), with live CATEGORIZE formulas',
    keywords: 'categorize category keywords rules classify tag agenda', origin: 'Lotus Agenda categories',
    run: (app, arg?: { rules: string; ifNone?: string; values?: boolean }) => categorizeCmd(app, arg),
  },
  {
    id: 'data.queryDelete', label: 'Delete Matching Records', icon: 'delete_sweep', tab: D, group: 'Records', edits: true,
    tip: 'Remove the records of a list that meet a condition and close up the gaps',
    keywords: 'delete records rows matching criteria query condition remove', origin: 'Lotus 1-2-3 /Data Query Delete',
    run: (app, arg?: { column: number; op: CondOp; value?: string }) => queryDelete(app, arg),
  },
  {
    id: 'data.matrix', label: 'Matrix', icon: 'grid_4x4', tab: D, group: 'Analyze', kind: 'menu', edits: true,
    tip: 'Invert or multiply matrices, solve linear equations, or find a determinant, writing the result into the sheet',
    keywords: 'matrix invert inverse multiply solve linear equations determinant', origin: 'Lotus 1-2-3 /Data Matrix',
    run: (app, arg?: any) => matrixCmd(app, arg),
    menu: () => [
      { id: 'data.matrix', arg: 'invert', label: 'Invert…', icon: 'flip' },
      { id: 'data.matrix', arg: 'multiply', label: 'Multiply…', icon: 'close' },
      { id: 'data.matrix', arg: 'solve', label: 'Solve Linear Equations…', icon: 'calculate' },
      { id: 'data.matrix', arg: 'determinant', label: 'Determinant…', icon: 'functions' },
    ],
  },
  {
    id: 'data.monteCarlo', label: 'Monte Carlo Simulation', icon: 'casino', tab: D, group: 'Forecast', edits: true,
    tip: 'Give inputs a range ("10 to 20", "normal 100 15"), run the model many times, and get the spread of the results with a chart',
    keywords: 'monte carlo simulation risk uncertainty distribution range scenarios', origin: 'Uncertain inputs in Causal; spreadsheet risk-analysis add-ins',
    run: (app, arg?: McSpec) => monteCarloCmd(app, arg),
  },
  {
    id: 'data.whatIfSliders', label: 'What-If Sliders', icon: 'tune', tab: D, group: 'Forecast', edits: false, noRecord: true,
    tip: 'Sliders for every number the selected result depends on: drag one and watch the result change',
    keywords: 'sliders what if interactive inputs sensitivity live', origin: 'Live model controls in Causal and reactive notebooks',
    run: (app) => { const s = app.sheet; const { r, c } = s.sel; if (app.ui.paneOpen('sliders')) app.ui.closePane(); else app.ui.openPane('sliders', { sheetId: s.id, r, c }); },
  },
]);

