/**
 * What-If Analysis and optimization: Data Table, Goal Seek, Scenario Manager, Forecast Sheet and Solver
 * (two-phase simplex with branch and bound for linear models, a penalty Nelder–Mead search for nonlinear ones).
 * Trial values are substituted into the model without touching undo history; only the final results are written
 * inside a transaction.
 */
import { def } from './registry';
import type { App } from '../app';
import type { Sheet, Workbook } from '../model/workbook';
import { newId } from '../model/workbook';
import type { Cell, Scalar, Scenario, ChartDef } from '../model/types';
import type { Engine } from '../calc/engine';
import { CellError } from '../model/types';
import { cellKey, cellName, parseRange, splitSheetRef, colName, type Range } from '../model/address';
import { resolveRangeText, rangeRefText } from '../model/structure';
import { isDateFormat } from '../model/numfmt';
import { activeRange, clipToUsed, currentRegion, addrText } from './ops';
import { openDialog, toast, alertDialog, setListOptions, type Field } from '../ui/dialog';
import { blocked, addSheetAfter, writeMatrix } from './data';

// --------------------------------------------------------------------------------------------
// Substitution of trial values

export interface CellRef { sheet: Sheet; r: number; c: number }

/** Temporarily replaces cell values (outside undo history) and recalculates; restore() puts everything back. */
export class Substituter {
  private saved = new Map<string, { ref: CellRef; cell: Cell | undefined }>();
  constructor(public book: Workbook, public engine: Engine) {}

  set(ref: CellRef, v: Scalar): void {
    const k = ref.sheet.id + ':' + cellKey(ref.r, ref.c);
    const cur = ref.sheet.get(ref.r, ref.c);
    if (!this.saved.has(k)) this.saved.set(k, { ref, cell: cur ? { ...cur } : undefined });
    const next: Cell = { ...(cur || {}), v };
    delete next.f; delete next.sp; delete next.arr;
    ref.sheet.setComputed(ref.r, ref.c, next);
    this.book.noteChange(ref.sheet, cellKey(ref.r, ref.c));
  }

  recalc(): void { this.engine.recalc(this.book.calc.mode === 'manual'); }

  restore(): void {
    for (const { ref, cell } of this.saved.values()) {
      ref.sheet.setComputed(ref.r, ref.c, cell ? { ...cell } : undefined);
      this.book.noteChange(ref.sheet, cellKey(ref.r, ref.c));
    }
    this.saved.clear();
    this.recalc();
  }
}

function numOf(v: Scalar): number { return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : NaN; }

/** Parses "B1", "Sheet2!$B$1:$B$3, C5" into cells (max `max`). */
export function parseCells(book: Workbook, text: string, home: Sheet, max = 200): CellRef[] | null {
  const out: CellRef[] = [];
  for (const part of text.split(/[,;]/).map((s) => s.trim()).filter(Boolean)) {
    const x = resolveRangeText(book, part, home);
    if (!x) return null;
    for (let r = x.range.r1; r <= x.range.r2; r++) for (let c = x.range.c1; c <= x.range.c2; c++) { out.push({ sheet: x.sheet, r, c }); if (out.length > max) return null; }
  }
  return out.length ? out : null;
}

const refName = (ref: CellRef, home: Sheet) => (ref.sheet === home ? '' : ref.sheet.name + '!') + cellName(ref.r, ref.c, true, true);

// --------------------------------------------------------------------------------------------
// Data Table

/**
 * Computes a one- or two-variable data table over rg (Excel layout: input values down the first column and/or
 * across the first row, formulas in the first row / first column / corner). Returns the body values.
 */
export function computeDataTable(sub: Substituter, sheet: Sheet, rg: Range, rowInput: CellRef | null, colInput: CellRef | null): Scalar[][] {
  const out: Scalar[][] = [];
  const read = (r: number, c: number) => sheet.value(r, c);
  try {
    if (rowInput && colInput) {
      for (let r = rg.r1 + 1; r <= rg.r2; r++) {
        const row: Scalar[] = [];
        for (let c = rg.c1 + 1; c <= rg.c2; c++) {
          sub.set(rowInput, read(rg.r1, c)); sub.set(colInput, read(r, rg.c1)); sub.recalc();
          row.push(read(rg.r1, rg.c1));
        }
        out.push(row);
      }
    } else if (colInput) {
      for (let r = rg.r1 + 1; r <= rg.r2; r++) {
        sub.set(colInput, read(r, rg.c1)); sub.recalc();
        const row: Scalar[] = [];
        for (let c = rg.c1 + 1; c <= rg.c2; c++) row.push(read(rg.r1, c));
        out.push(row);
      }
    } else if (rowInput) {
      const cols: Scalar[][] = [];
      for (let c = rg.c1 + 1; c <= rg.c2; c++) {
        sub.set(rowInput, read(rg.r1, c)); sub.recalc();
        const col: Scalar[] = [];
        for (let r = rg.r1 + 1; r <= rg.r2; r++) col.push(read(r, rg.c1));
        cols.push(col);
      }
      for (let i = 0; i < rg.r2 - rg.r1; i++) out.push(cols.map((col) => col[i]));
    }
  } finally {
    sub.restore();
  }
  return out;
}

async function dataTableDialog(app: App): Promise<void> {
  if (blocked(app)) return;
  const sheet = app.sheet, book = app.book;
  const rg = activeRange(app);
  if (rg.r2 - rg.r1 < 1 || rg.c2 - rg.c1 < 1) { void alertDialog('Data Table', 'Select the whole table first: the input values and formulas plus the empty cells for the results (at least two rows and two columns).'); return; }
  const res = await openDialog({
    title: 'Data Table',
    help: 'Row input cell: the cell the values across the top row are substituted into. Column input cell: the cell the values down the first column are substituted into.',
    fields: [
      { key: 'row', type: 'text', label: 'Row input cell', value: '', placeholder: 'e.g. $B$3' },
      { key: 'col', type: 'text', label: 'Column input cell', value: '', placeholder: 'e.g. $B$4' },
    ],
    validate: (v) => (!String(v.row).trim() && !String(v.col).trim() ? 'Enter a row input cell, a column input cell, or both.' : null),
  });
  if (!res) return;
  const one = (t: string) => { if (!t.trim()) return null; const x = parseCells(book, t, sheet, 1); return x && x.length === 1 ? x[0] : undefined; };
  const rowIn = one(String(res.row)), colIn = one(String(res.col));
  if (rowIn === undefined || colIn === undefined) { toast('Input cell reference is not valid.'); return; }
  for (const x of [rowIn, colIn]) if (x && x.r >= rg.r1 && x.r <= rg.r2 && x.c >= rg.c1 && x.c <= rg.c2 && x.sheet === sheet) { toast('The input cell must be outside the table.'); return; }
  const sub = new Substituter(book, app.engine);
  const body = computeDataTable(sub, sheet, rg, rowIn, colIn);
  book.transact('Data Table', () => {
    body.forEach((row, i) => row.forEach((v, j) => {
      const r = rg.r1 + 1 + i, c = rg.c1 + 1 + j;
      const s = sheet.get(r, c)?.s;
      sheet.set(r, c, v == null ? (s ? { s } : undefined) : { v, ...(s ? { s } : {}) });
    }));
  });
  toast('Data table results were written as values. Run Data Table again after changing the model.');
}

// --------------------------------------------------------------------------------------------
// Goal Seek

export interface SeekResult { x: number; y: number; ok: boolean; iterations: number }

/** Finds x so that f(x) = target (Newton/secant steps, switching to bracketed regula falsi once the sign changes). */
export function goalSeek(f: (x: number) => number, target: number, x0: number, opts: { maxIter?: number; tol?: number } = {}): SeekResult {
  const maxIter = opts.maxIter ?? 100;
  const tol = opts.tol ?? 1e-7 * Math.max(1, Math.abs(target));
  const g = (x: number) => { const y = f(x); return Number.isFinite(y) ? y - target : NaN; };
  let x = Number.isFinite(x0) ? x0 : 0;
  let gx = g(x);
  let best = { x, gx: Number.isFinite(gx) ? gx : Infinity };
  const note = (xx: number, gg: number) => { if (Number.isFinite(gg) && Math.abs(gg) < Math.abs(best.gx)) best = { x: xx, gx: gg }; };
  let it = 0;
  if (!Number.isFinite(gx)) { x = 0; gx = g(x); note(x, gx); }
  for (; it < maxIter; it++) {
    if (Math.abs(gx) <= tol) return { x, y: gx + target, ok: true, iterations: it };
    const hh = 1e-6 * Math.max(1, Math.abs(x));
    const d = (g(x + hh) - gx) / hh;
    let step: number;
    if (!Number.isFinite(d) || Math.abs(d) < 1e-300) step = (x === 0 ? 1 : Math.abs(x)) * (it % 2 ? -1 : 1) * (1 + it);
    else step = -gx / d;
    const lim = 1e3 * (1 + Math.abs(x));
    if (Math.abs(step) > lim) step = Math.sign(step) * lim;
    let xn = x + step, gn = g(xn);
    for (let k = 0; k < 30 && !Number.isFinite(gn); k++) { step /= 2; xn = x + step; gn = g(xn); }
    note(xn, gn);
    if (Number.isFinite(gn) && Math.sign(gn) !== Math.sign(gx) && gn !== 0) {
      // Bracketed: Illinois regula falsi.
      let a = x, fa = gx, b = xn, fb = gn, side = 0;
      for (let k = 0; k < 200; k++) {
        const cx = (a * fb - b * fa) / (fb - fa);
        const fc = g(cx);
        note(cx, fc);
        if (!Number.isFinite(fc)) break;
        if (Math.abs(fc) <= tol || Math.abs(b - a) < 1e-15 * (1 + Math.abs(cx))) return { x: cx, y: fc + target, ok: Math.abs(fc) <= Math.max(tol, 1e-3), iterations: it + k };
        if (Math.sign(fc) === Math.sign(fb)) { b = cx; fb = fc; if (side === -1) fa /= 2; side = -1; }
        else { a = cx; fa = fc; if (side === 1) fb /= 2; side = 1; }
      }
      return { x: best.x, y: best.gx + target, ok: Math.abs(best.gx) <= Math.max(tol, 1e-3), iterations: it };
    }
    if (!Number.isFinite(gn)) break;
    x = xn; gx = gn;
  }
  return { x: best.x, y: best.gx + target, ok: Math.abs(best.gx) <= Math.max(tol, 1e-3), iterations: it };
}

async function goalSeekDialog(app: App): Promise<void> {
  if (blocked(app)) return;
  const sheet = app.sheet, book = app.book;
  const res = await openDialog({
    title: 'Goal Seek',
    fields: [
      { key: 'set', type: 'text', label: 'Set cell', value: cellName(sheet.sel.r, sheet.sel.c, true, true) },
      { key: 'to', type: 'number', label: 'To value', value: '' as any },
      { key: 'by', type: 'text', label: 'By changing cell', value: '' },
    ],
    validate: (v) => (v.to == null || !Number.isFinite(v.to) ? 'Enter a number in To value.' : !String(v.by).trim() ? 'Enter the cell to change.' : null),
  });
  if (!res) return;
  const setC = parseCells(book, String(res.set), sheet, 1), byC = parseCells(book, String(res.by), sheet, 1);
  if (!setC || setC.length !== 1) { toast('Set cell reference is not valid.'); return; }
  if (!byC || byC.length !== 1) { toast('By changing cell reference is not valid.'); return; }
  const S = setC[0], B = byC[0];
  if (S.sheet.get(S.r, S.c)?.f == null) { void alertDialog('Goal Seek', 'Set cell must contain a formula.'); return; }
  const bc = B.sheet.get(B.r, B.c);
  if (bc?.f != null || (bc?.v != null && typeof bc.v !== 'number')) { void alertDialog('Goal Seek', 'By changing cell must contain a number (not a formula).'); return; }
  const sub = new Substituter(book, app.engine);
  const f = (x: number) => { sub.set(B, x); sub.recalc(); return numOf(S.sheet.value(S.r, S.c)); };
  const x0 = numOf(bc?.v ?? 0) || 0;
  let r: SeekResult;
  try { r = goalSeek(f, +res.to, x0, { maxIter: book.calc.maxIterations || 100 }); } finally { sub.restore(); }
  const keep = await openDialog({
    title: 'Goal Seek Status',
    content: (() => { const d = document.createElement('div'); d.className = 'ql-msg'; d.innerText = `Goal seeking with cell ${refName(S, sheet)} ${r.ok ? 'found a solution' : 'may not have found a solution'}.\n\nTarget value: ${res.to}\nCurrent value: ${+r.y.toPrecision(12)}\n${refName(B, sheet)} = ${+r.x.toPrecision(15)}`; return d; })(),
    ok: 'OK', cancel: 'Cancel',
  });
  if (!keep) return;
  book.transact('Goal Seek', () => { const cur = B.sheet.get(B.r, B.c); B.sheet.set(B.r, B.c, { ...(cur?.s ? { s: cur.s } : {}), v: +r.x.toPrecision(15) }); });
}

// --------------------------------------------------------------------------------------------
// Scenarios

export function scenarioCells(book: Workbook, sc: Scenario, home: Sheet): CellRef[] { return parseCells(book, sc.cells, home, 32) || []; }

/** Values of result cells under each scenario (and currently). */
export function scenarioResults(sub: Substituter, book: Workbook, home: Sheet, scenarios: Scenario[], results: CellRef[]): { current: Scalar[]; per: Scalar[][] } {
  const current = results.map((x) => x.sheet.value(x.r, x.c));
  const per: Scalar[][] = [];
  try {
    for (const sc of scenarios) {
      const cells = scenarioCells(book, sc, home);
      cells.forEach((ref, i) => sub.set(ref, sc.values[i] ?? null));
      sub.recalc();
      per.push(results.map((x) => x.sheet.value(x.r, x.c)));
      sub.restore();
    }
  } finally { sub.restore(); }
  return { current, per };
}

async function scenarioEdit(app: App, existing?: Scenario): Promise<Scenario | null> {
  const sheet = app.sheet, book = app.book;
  const res = await openDialog({
    title: existing ? 'Edit Scenario' : 'Add Scenario',
    fields: [
      { key: 'name', type: 'text', label: 'Scenario name', value: existing?.name || '', autofocus: true },
      { key: 'cells', type: 'text', label: 'Changing cells', value: existing?.cells || addrText(clipToUsed(sheet, activeRange(app))).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2'), help: 'Up to 32 cells, separated by commas.' },
      { key: 'comment', type: 'textarea', label: 'Comment', value: existing?.comment ?? `Created by ${app.prefs.author || 'Author'} on ${new Date().toLocaleDateString()}`, rows: 3 },
    ],
    validate: (v) => {
      if (!String(v.name).trim()) return 'Enter a scenario name.';
      if (sheet.scenarios.some((s) => s !== existing && s.name.toLowerCase() === String(v.name).trim().toLowerCase())) return 'A scenario with this name already exists.';
      const cells = parseCells(book, String(v.cells), sheet, 32);
      if (!cells) return 'Changing cells must be a reference to at most 32 cells.';
      if (cells.some((x) => x.sheet !== sheet)) return 'Changing cells must be on this sheet.';
      return null;
    },
  });
  if (!res) return null;
  const cells = parseCells(book, String(res.cells), sheet, 32)!;
  const vals = await openDialog({
    title: 'Scenario Values',
    help: 'Enter values for each of the changing cells.',
    fields: cells.map((x, i) => {
      const prev = existing && existing.cells === res.cells ? existing.values[i] : x.sheet.value(x.r, x.c);
      return { key: 'v' + i, type: 'text' as const, label: `${i + 1}: ${cellName(x.r, x.c, true, true)}`, value: prev == null ? '' : prev instanceof CellError ? prev.code : String(prev) };
    }),
  });
  if (!vals) return null;
  const values: Scalar[] = cells.map((_, i) => { const t = String(vals['v' + i] ?? '').trim(); if (t === '') return null; const n = Number(t); return Number.isFinite(n) ? n : /^(true|false)$/i.test(t) ? /^true$/i.test(t) : t; });
  return { name: String(res.name).trim(), cells: String(res.cells).trim(), values, comment: res.comment || undefined };
}

function showScenario(app: App, sc: Scenario): void {
  const book = app.book, sheet = app.sheet;
  const cells = scenarioCells(book, sc, sheet);
  book.transact(`Show Scenario ${sc.name}`, () => cells.forEach((x, i) => {
    const cur = x.sheet.get(x.r, x.c);
    const v = sc.values[i] ?? null;
    x.sheet.set(x.r, x.c, v == null ? (cur?.s ? { s: cur.s } : undefined) : { v, ...(cur?.s ? { s: cur.s } : {}) });
  }));
}

/** Builds the Scenario Summary report matrix. */
export function scenarioSummaryMatrix(home: Sheet, scenarios: Scenario[], changing: CellRef[], currentChanging: Scalar[], perChanging: Scalar[][], results: CellRef[], cur: { current: Scalar[]; per: Scalar[][] }): Scalar[][] {
  const out: Scalar[][] = [];
  out.push(['Scenario Summary']);
  out.push([null, 'Current Values:', ...scenarios.map((s) => s.name)]);
  out.push(['Changing Cells:']);
  changing.forEach((x, i) => out.push([refName(x, home), currentChanging[i], ...perChanging.map((p) => p[i] ?? null)]));
  if (results.length) {
    out.push(['Result Cells:']);
    results.forEach((x, i) => out.push([refName(x, home), cur.current[i], ...cur.per.map((p) => p[i])]));
  }
  out.push([]);
  out.push(['Notes: The Current Values column shows the changing cells when this report was created. Changing cells for each scenario appear in the columns to the right.']);
  return out;
}

async function scenarioSummary(app: App): Promise<void> {
  const sheet = app.sheet, book = app.book;
  if (!sheet.scenarios.length) { toast('There are no scenarios on this sheet.'); return; }
  const res = await openDialog({
    title: 'Scenario Summary',
    fields: [{ key: 'results', type: 'text', label: 'Result cells', value: cellName(sheet.sel.r, sheet.sel.c, true, true), help: 'Cells whose values the report shows for each scenario (formulas that depend on the changing cells).' }],
  });
  if (!res) return;
  const results = String(res.results).trim() ? parseCells(book, String(res.results), sheet, 100) : [];
  if (results == null) { toast('Result cells reference is not valid.'); return; }
  // Union of changing cells across scenarios.
  const seen = new Map<string, CellRef>();
  for (const sc of sheet.scenarios) for (const x of scenarioCells(book, sc, sheet)) seen.set(`${x.sheet.id}:${x.r}:${x.c}`, x);
  const changing = [...seen.values()];
  const currentChanging = changing.map((x) => x.sheet.value(x.r, x.c));
  const perChanging = sheet.scenarios.map((sc) => { const cells = scenarioCells(book, sc, sheet); return changing.map((x) => { const i = cells.findIndex((y) => y.sheet === x.sheet && y.r === x.r && y.c === x.c); return i >= 0 ? sc.values[i] ?? null : x.sheet.value(x.r, x.c); }); });
  const sub = new Substituter(book, app.engine);
  const cur = scenarioResults(sub, book, sheet, sheet.scenarios, results);
  const m = scenarioSummaryMatrix(sheet, sheet.scenarios, changing, currentChanging, perChanging, results, cur);
  let target!: Sheet;
  book.transact('Scenario Summary', () => {
    target = addSheetAfter(book, 'Scenario Summary');
    writeMatrix(target, 0, 0, m);
    target.set(0, 0, { v: 'Scenario Summary', s: book.styleId({ bold: true, size: 14 }) });
    for (let c = 1; c < 2 + sheet.scenarios.length; c++) { const x = target.get(1, c); if (x) target.set(1, c, { ...x, s: book.styleId({ bold: true, fill: '#d9d9d9', bb: { style: 'thin' } }) }); }
    target.colW.set(0, 110); target.colW.set(1, 100);
  });
  app.exec('sheet.activate', book.sheets.indexOf(target));
}

async function scenarioManager(app: App): Promise<void> {
  const sheet = app.sheet, book = app.book;
  const opts = () => sheet.scenarios.map((s) => ({ value: s.name, label: s.name }));
  const info = (name?: string) => { const s = sheet.scenarios.find((x) => x.name === name); return s ? `Changing cells: ${s.cells}${s.comment ? `\nComment: ${s.comment}` : ''}` : (sheet.scenarios.length ? '' : 'No scenarios defined. Choose Add to add scenarios.'); };
  let infoEl: HTMLElement | null = null;
  const refreshList = (ctl: any, sel?: string) => { setListOptions(ctl, 'list', opts(), sel); if (infoEl) infoEl.textContent = info(sel); };
  await openDialog({
    title: 'Scenario Manager', ok: 'Close', cancel: false, wide: true,
    fields: [
      { key: 'list', type: 'list', label: 'Scenarios', options: opts(), rows: 6, value: sheet.scenarios[0]?.name },
      { key: 'info', type: 'info', label: info(sheet.scenarios[0]?.name) || ' ' },
      { key: 'add', type: 'button', label: 'Add…', half: true, onClick: async (_v, ctl) => {
        if (blocked(app, 'scenarios')) return;
        const sc = await scenarioEdit(app);
        if (!sc) return;
        book.transact('Add Scenario', () => { book.recordSheet(sheet); sheet.scenarios.push(sc); });
        refreshList(ctl, sc.name);
      } },
      { key: 'edit', type: 'button', label: 'Edit…', half: true, onClick: async (v, ctl) => {
        const cur = sheet.scenarios.find((x) => x.name === v.list);
        if (!cur) { toast('Select a scenario.'); return; }
        const sc = await scenarioEdit(app, cur);
        if (!sc) return;
        book.transact('Edit Scenario', () => { book.recordSheet(sheet); const i = sheet.scenarios.findIndex((x) => x.name === cur.name); if (i >= 0) sheet.scenarios[i] = sc; });
        refreshList(ctl, sc.name);
      } },
      { key: 'del', type: 'button', label: 'Delete', half: true, onClick: (v, ctl) => {
        if (!v.list) return;
        book.transact('Delete Scenario', () => { book.recordSheet(sheet); sheet.scenarios = sheet.scenarios.filter((x) => x.name !== v.list); });
        refreshList(ctl, sheet.scenarios[0]?.name);
      } },
      { key: 'show', type: 'button', label: 'Show', half: true, onClick: (v) => { const sc = sheet.scenarios.find((x) => x.name === v.list); if (sc) showScenario(app, sc); else toast('Select a scenario.'); } },
      { key: 'merge', type: 'button', label: 'Merge…', half: true, onClick: async (_v, ctl) => {
        const others = book.sheets.filter((s) => s !== sheet && s.scenarios.length);
        if (!others.length) { toast('No other sheet has scenarios.'); return; }
        const r = await openDialog({ title: 'Merge Scenarios', fields: [{ key: 's', type: 'select', label: 'Merge scenarios from sheet', options: others.map((s) => s.name), value: others[0].name }] });
        if (!r) return;
        const src = book.sheetByName(r.s)!;
        book.transact('Merge Scenarios', () => { book.recordSheet(sheet); for (const sc of src.scenarios) { let name = sc.name; for (let i = 2; sheet.scenarios.some((x) => x.name === name); i++) name = `${sc.name} ${i}`; sheet.scenarios.push({ ...sc, name, values: [...sc.values] }); } });
        refreshList(ctl);
      } },
      { key: 'sum', type: 'button', label: 'Summary…', half: true, onClick: (_v, ctl) => { ctl.close({}); void scenarioSummary(app); } },
    ],
    onOpen: (ctl) => { infoEl = ctl.body.querySelector('.ql-info'); },
    onChange: (v, _ctl, key) => { if (key === 'list' && infoEl) infoEl.textContent = info(v.list); },
  });
}

// --------------------------------------------------------------------------------------------
// Forecast Sheet

export interface ForecastOptions { periods: number; confidence: number; seasonality: number | 'auto' | 'none'; timelineFmt?: string; valuesFmt?: string; name?: string; chart?: 'line' | 'column' }

/**
 * Creates a forecast sheet: Timeline, Values, Forecast (FORECAST.ETS) and confidence bounds
 * (FORECAST.ETS.CONFINT), plus a chart. Call inside a transaction. Returns the new sheet.
 */
export function createForecastSheet(book: Workbook, timeline: number[], values: Scalar[], o: ForecastOptions): Sheet {
  const n = timeline.length;
  if (n < 3) throw new Error('A forecast needs at least three data points.');
  const order = timeline.map((t, i) => i).sort((a, b) => timeline[a] - timeline[b]);
  const tl = order.map((i) => timeline[i]), vs = order.map((i) => values[i]);
  const steps = tl.slice(1).map((t, i) => t - tl[i]).sort((a, b) => a - b);
  const step = steps[Math.floor(steps.length / 2)] || 1;
  if (!(step > 0)) throw new Error('The timeline needs distinct, evenly spaced values.');
  const sheet = addSheetAfter(book, o.name || 'Forecast');
  const bold = book.styleId({ bold: true, wrap: true, vAlign: 'bottom' });
  const tfmt = o.timelineFmt && o.timelineFmt !== 'General' ? book.styleId({ numFmt: o.timelineFmt }) : 0;
  const vfmt = book.styleId({ numFmt: o.valuesFmt && o.valuesFmt !== 'General' ? o.valuesFmt : '#,##0.00' });
  const heads = ['Timeline', 'Values', 'Forecast', 'Lower Confidence Bound', 'Upper Confidence Bound'];
  heads.forEach((t, j) => sheet.set(0, j, { v: t, s: bold }));
  const seas = o.seasonality === 'auto' ? 1 : o.seasonality === 'none' ? 0 : Math.max(0, Math.round(o.seasonality));
  const last = n + 1; // 1-based row of the last historical value
  const T = `$A$2:$A$${last}`, V = `$B$2:$B$${last}`;
  for (let i = 0; i < n; i++) {
    sheet.set(i + 1, 0, { v: tl[i], ...(tfmt ? { s: tfmt } : {}) });
    const v = vs[i];
    sheet.set(i + 1, 1, v == null ? { s: vfmt } : { v, s: vfmt });
  }
  // Join the forecast lines to the last actual value.
  for (let j = 2; j <= 4; j++) sheet.set(n, j, { f: `B${last}`, s: vfmt });
  const conf = Math.min(0.999, Math.max(0.01, o.confidence));
  for (let k = 1; k <= o.periods; k++) {
    const r = n + k, a1 = r + 1;
    sheet.set(r, 0, { v: +(tl[n - 1] + step * k).toPrecision(15), ...(tfmt ? { s: tfmt } : {}) });
    sheet.set(r, 2, { f: `FORECAST.ETS(A${a1},${V},${T},${seas},1)`, s: vfmt });
    sheet.set(r, 3, { f: `C${a1}-FORECAST.ETS.CONFINT(A${a1},${V},${T},${conf},${seas},1)`, s: vfmt });
    sheet.set(r, 4, { f: `C${a1}+FORECAST.ETS.CONFINT(A${a1},${V},${T},${conf},${seas},1)`, s: vfmt });
  }
  sheet.colW.set(0, 90); for (let j = 1; j <= 4; j++) sheet.colW.set(j, 96);
  sheet.rowH.set(0, 34);
  const end = n + o.periods + 1;
  const ref = (c: number) => rangeRefText(sheet, { r1: 1, c1: c, r2: end - 1, c2: c });
  const chart: ChartDef = {
    id: newId('ch'), type: o.chart === 'column' ? 'column' : 'line', title: 'Forecast',
    series: [
      { name: 'Values', nameRef: rangeRefText(sheet, { r1: 0, c1: 1, r2: 0, c2: 1 }), values: ref(1), categories: ref(0), color: '#2f6fb0' },
      { name: 'Forecast', nameRef: rangeRefText(sheet, { r1: 0, c1: 2, r2: 0, c2: 2 }), values: ref(2), categories: ref(0), color: '#e07b28' },
      { name: 'Lower Confidence Bound', nameRef: rangeRefText(sheet, { r1: 0, c1: 3, r2: 0, c2: 3 }), values: ref(3), categories: ref(0), color: '#7a8a99', type: 'line' },
      { name: 'Upper Confidence Bound', nameRef: rangeRefText(sheet, { r1: 0, c1: 4, r2: 0, c2: 4 }), values: ref(4), categories: ref(0), color: '#7a8a99', type: 'line' },
    ],
    at: { r: 1, c: 6, dx: 0, dy: 0, w: 560, h: 320 }, legend: 'bottom', gridlines: true, sourceRange: rangeRefText(sheet, { r1: 0, c1: 0, r2: end - 1, c2: 4 }),
  };
  sheet.charts.push(chart);
  return sheet;
}

async function forecastDialog(app: App): Promise<void> {
  const sheet = app.sheet, book = app.book;
  const sel = activeRange(app);
  let rg = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? currentRegion(sheet, sel.r1, sel.c1) : clipToUsed(sheet, sel);
  const hasHead = typeof sheet.value(rg.r1, rg.c1) === 'string' || typeof sheet.value(rg.r1, rg.c1 + 1) === 'string';
  const first = hasHead ? rg.r1 + 1 : rg.r1;
  const tlDefault = rangeRefText(sheet, { r1: first, c1: rg.c1, r2: rg.r2, c2: rg.c1 }, true, sheet);
  const vDefault = rangeRefText(sheet, { r1: first, c1: rg.c1 + 1, r2: rg.r2, c2: rg.c1 + 1 }, true, sheet);
  const res = await openDialog({
    title: 'Create Forecast Worksheet',
    fields: [
      { key: 'tl', type: 'text', label: 'Timeline range', value: tlDefault, help: 'Dates or numbers with a consistent step.' },
      { key: 'vals', type: 'text', label: 'Values range', value: vDefault },
      { key: 'periods', type: 'number', label: 'Forecast periods', value: Math.max(4, Math.round((rg.r2 - first + 1) / 3)), min: 1, max: 10000, half: true },
      { key: 'conf', type: 'number', label: 'Confidence interval (%)', value: 95, min: 1, max: 99.9, half: true },
      { key: 'season', type: 'radio', label: 'Seasonality', options: [{ value: 'auto', label: 'Detect automatically' }, { value: 'set', label: 'Set manually' }, { value: 'none', label: 'None' }], value: 'auto' },
      { key: 'seasN', type: 'number', label: 'Season length', value: 12, min: 2, max: 8760, visibleIf: (v) => v.season === 'set' },
      { key: 'chart', type: 'select', label: 'Chart', options: [{ value: 'line', label: 'Line chart' }, { value: 'column', label: 'Column chart' }], value: 'line' },
    ],
    validate: (v) => (!(v.periods >= 1) ? 'Enter how many periods to forecast.' : null),
  });
  if (!res) return;
  const T = resolveRangeText(book, String(res.tl), sheet), V = resolveRangeText(book, String(res.vals), sheet);
  if (!T || !V) { toast('The timeline or values range is not valid.'); return; }
  const tl: number[] = [], vs: Scalar[] = [];
  const nT = T.range.r2 - T.range.r1 + 1;
  if (nT !== V.range.r2 - V.range.r1 + 1) { toast('The timeline and values ranges must be the same size.'); return; }
  for (let i = 0; i < nT; i++) {
    const t = T.sheet.value(T.range.r1 + i, T.range.c1);
    if (typeof t !== 'number') { void alertDialog('Forecast', `The timeline must contain only numbers or dates (cell ${cellName(T.range.r1 + i, T.range.c1)}).`); return; }
    tl.push(t); vs.push(V.sheet.value(V.range.r1 + i, V.range.c1));
  }
  const tfmt = T.sheet.styleOf(T.range.r1, T.range.c1).numFmt;
  let target!: Sheet;
  try {
    book.transact('Forecast Sheet', () => {
      target = createForecastSheet(book, tl, vs, { periods: Math.round(res.periods), confidence: res.conf / 100, seasonality: res.season === 'set' ? res.seasN : res.season, timelineFmt: isDateFormat(tfmt) ? tfmt : tfmt, valuesFmt: V.sheet.styleOf(V.range.r1, V.range.c1).numFmt, name: 'Forecast', chart: res.chart });
    });
  } catch (e) { void alertDialog('Forecast', (e as Error).message); return; }
  void rg;
  app.exec('sheet.activate', book.sheets.indexOf(target));
}

// --------------------------------------------------------------------------------------------
// Solver: numeric core

export type Rel = '<=' | '>=' | '=';

export interface LPResult { status: 'optimal' | 'infeasible' | 'unbounded'; x: number[]; value: number }

/** Maximizes c·x subject to A x (rel) b and x >= 0 (two-phase simplex with Bland's rule). */
export function simplex(c: number[], A: number[][], rel: Rel[], b: number[]): LPResult {
  const EPS = 1e-9;
  const m = A.length, n = c.length;
  const rows = A.map((r) => [...r]), rhs = [...b], ops = [...rel];
  for (let i = 0; i < m; i++) if (rhs[i] < 0) { rows[i] = rows[i].map((x) => -x); rhs[i] = -rhs[i]; ops[i] = ops[i] === '<=' ? '>=' : ops[i] === '>=' ? '<=' : '='; }
  const nS = ops.filter((o) => o !== '=').length, nA = ops.filter((o) => o !== '<=').length;
  const W = n + nS + nA;
  const T: number[][] = [];
  const basis: number[] = [];
  const artCols: number[] = [];
  let si = n, ai = n + nS;
  for (let i = 0; i < m; i++) {
    const row = new Array(W + 1).fill(0);
    for (let j = 0; j < n; j++) row[j] = rows[i][j] || 0;
    row[W] = rhs[i];
    if (ops[i] === '<=') { row[si] = 1; basis.push(si++); }
    else if (ops[i] === '>=') { row[si++] = -1; row[ai] = 1; artCols.push(ai); basis.push(ai++); }
    else { row[ai] = 1; artCols.push(ai); basis.push(ai++); }
    T.push(row);
  }
  const isArt = new Set(artCols);
  const pivot = (pr: number, pc: number) => {
    const p = T[pr][pc];
    for (let j = 0; j <= W; j++) T[pr][j] /= p;
    for (let i = 0; i < m; i++) if (i !== pr && Math.abs(T[i][pc]) > EPS) { const f = T[i][pc]; for (let j = 0; j <= W; j++) T[i][j] -= f * T[pr][j]; }
    basis[pr] = pc;
  };
  // obj: reduced-cost row for maximization (enter when obj[j] < -EPS); obj[W] = -value.
  const run = (obj: number[], allowed: (j: number) => boolean): 'optimal' | 'unbounded' => {
    for (let it = 0; it < 50000; it++) {
      let pc = -1;
      for (let j = 0; j < W; j++) if (allowed(j) && obj[j] < -EPS) { pc = j; break; }
      if (pc < 0) return 'optimal';
      let pr = -1, best = Infinity;
      for (let i = 0; i < m; i++) if (T[i][pc] > EPS) {
        const ratio = T[i][W] / T[i][pc];
        if (ratio < best - EPS || (Math.abs(ratio - best) <= EPS && basis[i] < basis[pr])) { best = ratio; pr = i; }
      }
      if (pr < 0) return 'unbounded';
      pivot(pr, pc);
      const f = obj[pc];
      for (let j = 0; j <= W; j++) obj[j] -= f * T[pr][j];
    }
    return 'optimal';
  };
  if (nA) {
    const obj = new Array(W + 1).fill(0);
    for (const a of artCols) obj[a] = 1;
    for (let i = 0; i < m; i++) if (isArt.has(basis[i])) for (let j = 0; j <= W; j++) obj[j] -= T[i][j];
    run(obj, () => true);
    if (-obj[W] > 1e-7 * (1 + Math.max(...rhs.map(Math.abs), 0))) return { status: 'infeasible', x: new Array(n).fill(0), value: NaN };
    for (let i = 0; i < m; i++) if (isArt.has(basis[i])) {
      let pc = -1;
      for (let j = 0; j < n + nS; j++) if (Math.abs(T[i][j]) > EPS) { pc = j; break; }
      if (pc >= 0) pivot(i, pc);
    }
  }
  const obj = new Array(W + 1).fill(0);
  for (let j = 0; j < n; j++) obj[j] = -c[j];
  for (let i = 0; i < m; i++) { const f = obj[basis[i]]; if (Math.abs(f) > 0) for (let j = 0; j <= W; j++) obj[j] -= f * T[i][j]; }
  const st = run(obj, (j) => !isArt.has(j));
  const x = new Array(n).fill(0);
  for (let i = 0; i < m; i++) if (basis[i] < n) x[basis[i]] = T[i][W];
  if (st === 'unbounded') return { status: 'unbounded', x, value: Infinity };
  return { status: 'optimal', x, value: c.reduce((s, ci, j) => s + ci * x[j], 0) };
}

/** Maximizes c·x with A x (rel) b, bounds lo <= x (lo may be -Infinity for free variables) and integer variables. */
export function solveLP(c: number[], A: number[][], rel: Rel[], b: number[], opts: { nonNeg: boolean[]; ints?: number[]; maxNodes?: number } ): LPResult {
  const n = c.length;
  // Free variables are split into x = p - q.
  const free = opts.nonNeg.map((nn) => !nn);
  const map: number[][] = [];
  let k = 0;
  for (let j = 0; j < n; j++) { map.push(free[j] ? [k, k + 1] : [k]); k += free[j] ? 2 : 1; }
  const expand = (row: number[]) => { const out = new Array(k).fill(0); for (let j = 0; j < n; j++) { out[map[j][0]] = row[j]; if (map[j][1] != null) out[map[j][1]] = -row[j]; } return out; };
  const back = (y: number[]) => map.map((m) => y[m[0]] - (m[1] != null ? y[m[1]] : 0));
  const C = expand(c);
  const solveWith = (extraA: number[][], extraRel: Rel[], extraB: number[]): LPResult => {
    const r = simplex(C, [...A, ...extraA].map(expand), [...rel, ...extraRel], [...b, ...extraB]);
    return { ...r, x: back(r.x) };
  };
  const ints = opts.ints || [];
  const root = solveWith([], [], []);
  if (!ints.length || root.status !== 'optimal') return root;
  let best: LPResult | null = null;
  let nodes = 0;
  const maxNodes = opts.maxNodes ?? 5000;
  const stack: { a: number[][]; r: Rel[]; b: number[] }[] = [{ a: [], r: [], b: [] }];
  while (stack.length && nodes++ < maxNodes) {
    const nd = stack.pop()!;
    const res = solveWith(nd.a, nd.r, nd.b);
    if (res.status !== 'optimal') continue;
    if (best && res.value <= best.value + 1e-9) continue;
    const frac = ints.find((j) => Math.abs(res.x[j] - Math.round(res.x[j])) > 1e-6);
    if (frac == null) { best = { ...res, x: res.x.map((v, j) => (ints.includes(j) ? Math.round(v) : v)) }; continue; }
    const unit = new Array(n).fill(0); unit[frac] = 1;
    const v = res.x[frac];
    stack.push({ a: [...nd.a, unit], r: [...nd.r, '>='], b: [...nd.b, Math.ceil(v)] });
    stack.push({ a: [...nd.a, unit], r: [...nd.r, '<='], b: [...nd.b, Math.floor(v)] });
  }
  return best || { status: 'infeasible', x: root.x, value: NaN };
}

/** Nelder–Mead minimization. */
export function nelderMead(f: (x: number[]) => number, x0: number[], opts: { maxIter?: number; tol?: number; step?: number[] } = {}): { x: number[]; fx: number } {
  const n = x0.length;
  const maxIter = opts.maxIter ?? 400 * Math.max(1, n);
  const tol = opts.tol ?? 1e-12;
  const F = (x: number[]) => { const v = f(x); return Number.isFinite(v) ? v : 1e300; };
  let pts = [x0.slice()];
  for (let i = 0; i < n; i++) { const p = x0.slice(); p[i] += opts.step?.[i] ?? (Math.abs(p[i]) > 1e-8 ? 0.1 * Math.abs(p[i]) + 0.05 : 0.5); pts.push(p); }
  let vals = pts.map(F);
  for (let it = 0; it < maxIter; it++) {
    const idx = vals.map((_, i) => i).sort((a, b) => vals[a] - vals[b]);
    pts = idx.map((i) => pts[i]); vals = idx.map((i) => vals[i]);
    const spread = Math.abs(vals[n] - vals[0]);
    const size = Math.max(...pts.slice(1).map((p) => Math.max(...p.map((v, j) => Math.abs(v - pts[0][j])))));
    if (spread <= tol * (1 + Math.abs(vals[0])) && size < 1e-9 * (1 + Math.max(...pts[0].map(Math.abs)))) break;
    const cen = new Array(n).fill(0);
    for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) cen[j] += pts[i][j] / n;
    const at = (t: number) => cen.map((cv, j) => cv + t * (pts[n][j] - cv));
    const xr = at(-1), fr = F(xr);
    if (fr < vals[0]) { const xe = at(-2), fe = F(xe); if (fe < fr) { pts[n] = xe; vals[n] = fe; } else { pts[n] = xr; vals[n] = fr; } }
    else if (fr < vals[n - 1]) { pts[n] = xr; vals[n] = fr; }
    else {
      const xc = fr < vals[n] ? at(-0.5) : at(0.5), fc = F(xc);
      if (fc < Math.min(fr, vals[n])) { pts[n] = xc; vals[n] = fc; }
      else { for (let i = 1; i <= n; i++) { pts[i] = pts[i].map((v, j) => pts[0][j] + 0.5 * (v - pts[0][j])); vals[i] = F(pts[i]); } }
    }
  }
  let bi = 0; for (let i = 1; i <= n; i++) if (vals[i] < vals[bi]) bi = i;
  return { x: pts[bi], fx: vals[bi] };
}

export interface SolverProblem {
  x0: number[];
  f: (x: number[]) => number;
  goal: 'max' | 'min' | 'value';
  target?: number;
  cons: { g: (x: number[]) => number; rel: Rel }[];
  ints?: number[];
  nonNeg: boolean;
  method?: 'auto' | 'simplex' | 'grg';
}

export interface SolverResult { status: 'optimal' | 'infeasible' | 'unbounded' | 'notLinear' | 'stopped'; x: number[]; value: number; linear: boolean; message: string }

/** Numerically checks whether f is affine around x0; returns [f(x0), gradient] or null. */
export function linearize(f: (x: number[]) => number, x0: number[]): { f0: number; grad: number[] } | null {
  const n = x0.length;
  const f0 = f(x0);
  if (!Number.isFinite(f0)) return null;
  const grad = x0.map((_, i) => { const p = x0.slice(); p[i] += 1; return f(p) - f0; });
  if (grad.some((g) => !Number.isFinite(g))) return null;
  let seed = 12345;
  const rnd = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
  for (let t = 0; t < 3; t++) {
    const p = x0.map((v) => v + (rnd() * 6 - 3) * Math.max(1, Math.abs(v)));
    const pred = f0 + grad.reduce((s, g, i) => s + g * (p[i] - x0[i]), 0);
    const act = f(p);
    if (!Number.isFinite(act) || Math.abs(act - pred) > 1e-7 * (1 + Math.abs(act) + Math.abs(pred))) return null;
  }
  void n;
  return { f0, grad };
}

export function solveProblem(p: SolverProblem): SolverResult {
  const n = p.x0.length;
  const ints = p.ints || [];
  const sgn = p.goal === 'min' ? -1 : 1;
  const feasTol = (x: number[]) => p.cons.every((cn) => { const g = cn.g(x); return cn.rel === '<=' ? g <= 1e-6 * (1 + Math.abs(g)) : cn.rel === '>=' ? g >= -1e-6 * (1 + Math.abs(g)) : Math.abs(g) <= 1e-6; }) && (!p.nonNeg || x.every((v) => v >= -1e-9));
  if (p.method !== 'grg') {
    const lf = linearize(p.f, p.x0);
    const lcs = lf ? p.cons.map((cn) => linearize(cn.g, p.x0)) : [];
    if (lf && lcs.every(Boolean)) {
      const A: number[][] = [], rel: Rel[] = [], b: number[] = [];
      for (let i = 0; i < p.cons.length; i++) {
        const L = lcs[i]!;
        A.push(L.grad); rel.push(p.cons[i].rel); b.push(L.grad.reduce((s, g, j) => s + g * p.x0[j], 0) - L.f0);
      }
      let c = lf.grad.map((g) => sgn * g);
      if (p.goal === 'value') {
        A.push(lf.grad); rel.push('='); b.push((p.target ?? 0) - lf.f0 + lf.grad.reduce((s, g, j) => s + g * p.x0[j], 0));
        c = new Array(n).fill(0);
      }
      const r = solveLP(c, A, rel, b, { nonNeg: new Array(n).fill(p.nonNeg), ints });
      if (r.status === 'optimal') {
        const x = r.x.map((v) => (Math.abs(v) < 1e-11 ? 0 : +v.toPrecision(15)));
        return { status: 'optimal', x, value: p.f(x), linear: true, message: 'Solver found a solution. All constraints and optimality conditions are satisfied.' };
      }
      return { status: r.status, x: p.x0, value: p.f(p.x0), linear: true, message: r.status === 'infeasible' ? 'Solver could not find a feasible solution.' : 'The objective cell values do not converge (the problem is unbounded).' };
    }
    if (p.method === 'simplex') return { status: 'notLinear', x: p.x0, value: p.f(p.x0), linear: false, message: 'The linearity conditions required by the Simplex LP method are not satisfied. Choose the GRG Nonlinear method.' };
  }
  // Nonlinear: penalty method with Nelder–Mead restarts.
  const proj = (x: number[]) => (p.nonNeg ? x.map((v) => Math.max(0, v)) : x);
  const obj = (x: number[]) => { const y = p.f(x); if (!Number.isFinite(y)) return NaN; return p.goal === 'value' ? (y - (p.target ?? 0)) ** 2 : -sgn * y; };
  const viol = (x: number[]) => p.cons.reduce((s, cn) => { const g = cn.g(x); if (!Number.isFinite(g)) return s + 1e12; const v = cn.rel === '<=' ? Math.max(0, g) : cn.rel === '>=' ? Math.max(0, -g) : g; return s + v * v; }, 0);
  const scale = Math.max(1, Math.abs(obj(proj(p.x0))) || 1);
  let x = proj(p.x0);
  let fixed: Map<number, number> | null = null;
  const solveStage = (start: number[]) => {
    let cur = start;
    for (const mu of [1, 1e2, 1e4, 1e6, 1e8, 1e10]) {
      const F = (y: number[]) => { let z = proj(y); if (fixed) z = z.map((v, i) => (fixed!.has(i) ? fixed!.get(i)! : v)); const o = obj(z); return Number.isFinite(o) ? o / scale + mu * viol(z) : NaN; };
      for (let rep = 0; rep < 3; rep++) { const r = nelderMead(F, cur, { maxIter: 600 * Math.max(1, n) }); cur = proj(r.x); }
    }
    return fixed ? cur.map((v, i) => (fixed!.has(i) ? fixed!.get(i)! : v)) : cur;
  };
  x = solveStage(x);
  if (ints.length) { fixed = new Map(ints.map((i) => [i, Math.round(x[i])])); x = solveStage(x.map((v, i) => (fixed!.has(i) ? fixed!.get(i)! : v))); }
  x = x.map((v) => +v.toPrecision(12));
  if (!feasTol(x)) return { status: 'infeasible', x, value: p.f(x), linear: false, message: 'Solver could not find a feasible solution.' };
  return { status: 'optimal', x, value: p.f(x), linear: false, message: 'Solver found a solution. All constraints are satisfied and the solution is a local optimum.' };
}

// --------------------------------------------------------------------------------------------
// Solver: sheet model (stored in hidden names like solver_opt, as spreadsheets do)

export interface SolverConstraint { lhs: string; rel: '<=' | '>=' | '=' | 'int' | 'bin' | 'dif'; rhs: string }
export interface SolverModel { objective: string; goal: 'max' | 'min' | 'value'; target: number; vars: string; cons: SolverConstraint[]; nonNeg: boolean; method: 'auto' | 'simplex' | 'grg' }

const REL_CODES: Record<SolverConstraint['rel'], number> = { '<=': 1, '=': 2, '>=': 3, int: 4, bin: 5, dif: 6 };

export function loadSolverModel(book: Workbook, sheet: Sheet): SolverModel {
  const get = (n: string) => book.names.find((x) => x.name.toLowerCase() === n && x.sheet === sheet.name)?.ref.replace(/^=/, '');
  const num = Number(get('solver_num') || 0);
  const cons: SolverConstraint[] = [];
  const relOf = Object.fromEntries(Object.entries(REL_CODES).map(([k, v]) => [String(v), k])) as Record<string, SolverConstraint['rel']>;
  for (let i = 1; i <= num; i++) { const lhs = get(`solver_lhs${i}`), rel = relOf[get(`solver_rel${i}`) || '1'], rhs = get(`solver_rhs${i}`) ?? ''; if (lhs && rel) cons.push({ lhs: stripSheet(lhs, sheet), rel, rhs: rel === 'int' ? 'integer' : rel === 'bin' ? 'binary' : rel === 'dif' ? 'AllDifferent' : stripSheet(rhs, sheet) }); }
  const typ = get('solver_typ');
  const eng = get('solver_eng');
  return {
    objective: stripSheet(get('solver_opt') || '', sheet), goal: typ === '2' ? 'min' : typ === '3' ? 'value' : 'max', target: Number(get('solver_val') || 0),
    vars: stripSheet(get('solver_adj') || '', sheet), cons, nonNeg: get('solver_neg') !== '2', method: eng === '2' ? 'simplex' : eng === '1' ? 'grg' : 'auto',
  };
}

function stripSheet(ref: string, sheet: Sheet): string {
  return ref.split(',').map((p) => { const { sheet: sn, ref: r } = splitSheetRef(p.trim()); return sn && sn.toLowerCase() === sheet.name.toLowerCase() ? r : p.trim(); }).join(',');
}

function qualify(ref: string, sheet: Sheet): string {
  if (!ref) return ref;
  return ref.split(',').map((p) => { const t = p.trim(); if (!t || /!/.test(t) || !parseRange(t.replace(/\$/g, ''))) return t; return `${sheet.name.includes(' ') || /\W/.test(sheet.name) ? `'${sheet.name.replace(/'/g, "''")}'` : sheet.name}!${t}`; }).join(',');
}

export function saveSolverModel(book: Workbook, sheet: Sheet, m: SolverModel): void {
  book.recordBook();
  book.names = book.names.filter((x) => !(x.sheet === sheet.name && /^solver_/i.test(x.name)));
  const put = (name: string, ref: string) => book.names.push({ name, ref: '=' + ref, sheet: sheet.name, hidden: true });
  put('solver_opt', qualify(m.objective, sheet));
  put('solver_typ', m.goal === 'max' ? '1' : m.goal === 'min' ? '2' : '3');
  put('solver_val', String(m.target || 0));
  put('solver_adj', qualify(m.vars, sheet));
  put('solver_num', String(m.cons.length));
  m.cons.forEach((cn, i) => { put(`solver_lhs${i + 1}`, qualify(cn.lhs, sheet)); put(`solver_rel${i + 1}`, String(REL_CODES[cn.rel])); put(`solver_rhs${i + 1}`, ['int', 'bin', 'dif'].includes(cn.rel) ? `"${cn.rhs}"` : qualify(cn.rhs, sheet)); });
  put('solver_neg', m.nonNeg ? '1' : '2');
  put('solver_eng', m.method === 'simplex' ? '2' : m.method === 'grg' ? '1' : '1');
}

/** Builds the numeric problem for a sheet model. */
export function buildSheetProblem(book: Workbook, engine: Engine, sheet: Sheet, m: SolverModel): { prob: SolverProblem; vars: CellRef[]; objective: CellRef; sub: Substituter; consRows: { lhs: CellRef; rel: Rel; rhsText: string; rhs: (x: number[]) => number }[] } | string {
  const objC = parseCells(book, m.objective, sheet, 1);
  if (!objC || objC.length !== 1) return 'The objective cell must be a single cell.';
  const objective = objC[0];
  const vars = parseCells(book, m.vars, sheet, 200);
  if (!vars) return 'The variable cells reference is not valid (at most 200 cells).';
  for (const v of vars) if (v.sheet.get(v.r, v.c)?.f != null) return `Variable cell ${refName(v, sheet)} contains a formula; variable cells must hold values.`;
  if (objective.sheet.get(objective.r, objective.c)?.f == null) return 'The objective cell must contain a formula that depends on the variable cells.';
  const sub = new Substituter(book, engine);
  const setX = (x: number[]) => { vars.forEach((v, i) => sub.set(v, x[i])); sub.recalc(); };
  const valAt = (ref: CellRef) => numOf(ref.sheet.value(ref.r, ref.c));
  const consRows: { lhs: CellRef; rel: Rel; rhsText: string; rhs: (x: number[]) => number }[] = [];
  const ints: number[] = [];
  const extraBin: number[] = [];
  for (const cn of m.cons) {
    const lhs = parseCells(book, cn.lhs, sheet, 1000);
    if (!lhs) return `Constraint reference ${cn.lhs} is not valid.`;
    if (cn.rel === 'int' || cn.rel === 'bin') {
      for (const l of lhs) { const i = vars.findIndex((v) => v.sheet === l.sheet && v.r === l.r && v.c === l.c); if (i < 0) return `Integer and binary constraints apply only to variable cells (${refName(l, sheet)}).`; ints.push(i); if (cn.rel === 'bin') extraBin.push(i); }
      continue;
    }
    if (cn.rel === 'dif') return 'The AllDifferent constraint is not supported.';
    const t = cn.rhs.trim();
    const n = Number(t);
    let rhsFns: ((x: number[]) => number)[];
    if (t !== '' && Number.isFinite(n)) rhsFns = lhs.map(() => () => n);
    else {
      const cells = parseCells(book, t.replace(/^=/, ''), sheet, 1000);
      if (cells && (cells.length === lhs.length || cells.length === 1)) rhsFns = lhs.map((_, i) => () => valAt(cells[cells.length === 1 ? 0 : i]));
      else { const f = t.replace(/^=/, ''); rhsFns = lhs.map(() => () => numOf(engine.evaluateScalar(f, sheet, objective.r, objective.c))); }
    }
    lhs.forEach((l, i) => consRows.push({ lhs: l, rel: cn.rel as Rel, rhsText: t, rhs: rhsFns[i] }));
  }
  const x0 = vars.map((v) => { const x = valAt(v); return Number.isFinite(x) ? x : 0; });
  const cons = consRows.map((cr) => ({ rel: cr.rel, g: (x: number[]) => { setX(x); return valAt(cr.lhs) - cr.rhs(x); } }));
  for (const i of extraBin) { cons.push({ rel: '<=', g: (x: number[]) => x[i] - 1 }); cons.push({ rel: '>=', g: (x: number[]) => x[i] }); }
  const prob: SolverProblem = { x0, f: (x) => { setX(x); return valAt(objective); }, goal: m.goal, target: m.target, cons, ints: [...new Set(ints)], nonNeg: m.nonNeg, method: m.method };
  return { prob, vars, objective, sub, consRows };
}

async function constraintDialog(cur?: SolverConstraint): Promise<SolverConstraint | null> {
  const res = await openDialog({
    title: cur ? 'Change Constraint' : 'Add Constraint',
    fields: [
      { key: 'lhs', type: 'text', label: 'Cell reference', value: cur?.lhs || '', autofocus: true, half: true },
      { key: 'rel', type: 'select', label: 'Relation', options: [{ value: '<=', label: '<=' }, { value: '=', label: '=' }, { value: '>=', label: '>=' }, { value: 'int', label: 'int' }, { value: 'bin', label: 'bin' }], value: cur?.rel || '<=', half: true },
      { key: 'rhs', type: 'text', label: 'Constraint', value: cur?.rhs || '', placeholder: 'number, cell or formula', visibleIf: (v) => !['int', 'bin'].includes(v.rel) },
    ],
    validate: (v) => (!String(v.lhs).trim() ? 'Enter a cell reference.' : !['int', 'bin'].includes(v.rel) && !String(v.rhs).trim() ? 'Enter a constraint value.' : null),
  });
  if (!res) return null;
  return { lhs: String(res.lhs).trim(), rel: res.rel, rhs: res.rel === 'int' ? 'integer' : res.rel === 'bin' ? 'binary' : String(res.rhs).trim() };
}

const consLabel = (c: SolverConstraint) => `${c.lhs} ${c.rel === 'int' ? '= integer' : c.rel === 'bin' ? '= binary' : `${c.rel} ${c.rhs}`}`;

async function solverDialog(app: App): Promise<void> {
  if (blocked(app)) return;
  const sheet = app.sheet, book = app.book;
  const m = loadSolverModel(book, sheet);
  if (!m.objective) m.objective = cellName(sheet.sel.r, sheet.sel.c, true, true);
  const cons = [...m.cons];
  const opts = () => cons.map((c, i) => ({ value: String(i), label: consLabel(c) }));
  const res = await openDialog({
    title: 'Solver Parameters', ok: 'Solve', wide: true,
    fields: [
      { key: 'objective', type: 'text', label: 'Set objective', value: m.objective, half: true },
      { key: 'goal', type: 'radio', label: 'To', options: [{ value: 'max', label: 'Max' }, { value: 'min', label: 'Min' }, { value: 'value', label: 'Value of' }], value: m.goal, half: true },
      { key: 'target', type: 'number', label: 'Value', value: m.target, visibleIf: (v) => v.goal === 'value' },
      { key: 'vars', type: 'text', label: 'By changing variable cells', value: m.vars, placeholder: '$B$2:$B$5' },
      { key: 'cons', type: 'list', label: 'Subject to the constraints', options: opts(), rows: 6 },
      { key: 'add', type: 'button', label: 'Add', half: true, onClick: async (_v, ctl) => { const c = await constraintDialog(); if (c) { cons.push(c); setListOptions(ctl, 'cons', opts()); } } },
      { key: 'chg', type: 'button', label: 'Change', half: true, onClick: async (v, ctl) => { const i = v.cons == null ? -1 : +v.cons; if (i < 0) { toast('Select a constraint.'); return; } const c = await constraintDialog(cons[i]); if (c) { cons[i] = c; setListOptions(ctl, 'cons', opts()); } } },
      { key: 'del', type: 'button', label: 'Delete', half: true, onClick: (v, ctl) => { const i = v.cons == null ? -1 : +v.cons; if (i >= 0) { cons.splice(i, 1); setListOptions(ctl, 'cons', opts()); } } },
      { key: 'reset', type: 'button', label: 'Reset All', half: true, onClick: (_v, ctl) => { cons.splice(0); setListOptions(ctl, 'cons', opts()); ctl.set('vars', ''); } },
      { key: 'nonNeg', type: 'checkbox', label: 'Make unconstrained variables non-negative', value: m.nonNeg },
      { key: 'method', type: 'select', label: 'Solving method', options: [{ value: 'auto', label: 'Automatic (Simplex LP when linear)' }, { value: 'simplex', label: 'Simplex LP' }, { value: 'grg', label: 'GRG Nonlinear' }], value: m.method },
    ],
    validate: (v) => (!String(v.vars).trim() ? 'Enter the variable cells.' : null),
  });
  if (!res) return;
  const model: SolverModel = { objective: String(res.objective).trim(), goal: res.goal, target: Number(res.target) || 0, vars: String(res.vars).trim(), cons, nonNeg: !!res.nonNeg, method: res.method };
  book.transact('Solver Model', () => saveSolverModel(book, sheet, model));
  const built = buildSheetProblem(book, app.engine, sheet, model);
  if (typeof built === 'string') { void alertDialog('Solver', built); return; }
  const { prob, vars, objective, sub, consRows } = built;
  const original = prob.x0.slice();
  const origObj = numOf(objective.sheet.value(objective.r, objective.c));
  let r: SolverResult;
  try { r = solveProblem(prob); } finally { sub.restore(); }
  const ok = r.status === 'optimal';
  const out = await openDialog({
    title: 'Solver Results',
    help: r.message,
    fields: [
      { key: 'keep', type: 'radio', options: [{ value: 'keep', label: 'Keep Solver Solution' }, { value: 'restore', label: 'Restore Original Values' }], value: ok ? 'keep' : 'restore' },
      { key: 'report', type: 'checkbox', label: 'Answer report (new sheet)', value: false },
      { key: 'i', type: 'info', label: `Objective ${model.objective}: ${+r.value.toPrecision(12)}${r.linear ? ' (linear model, Simplex LP)' : ' (nonlinear search)'}` },
    ],
  });
  if (!out) return;
  const keep = out.keep === 'keep';
  let target: Sheet | null = null;
  book.transact('Solver', () => {
    if (keep) vars.forEach((v, i) => { const cur = v.sheet.get(v.r, v.c); v.sheet.set(v.r, v.c, { ...(cur?.s ? { s: cur.s } : {}), v: r.x[i] }); });
    if (out.report) {
      target = addSheetAfter(book, 'Answer Report');
      const rows: Scalar[][] = [
        ['Answer Report'], [`Worksheet: ${sheet.name}`], [`Result: ${r.message}`], [`Engine: ${r.linear ? 'Simplex LP' : 'GRG Nonlinear (penalty search)'}`], [],
        [`Objective Cell (${model.goal === 'max' ? 'Max' : model.goal === 'min' ? 'Min' : `Value of ${model.target}`})`], ['Cell', 'Original Value', 'Final Value'],
        [refName(objective, sheet), origObj, r.value], [],
        ['Variable Cells'], ['Cell', 'Original Value', 'Final Value', 'Integer'],
        ...vars.map((v, i) => [refName(v, sheet), original[i], r.x[i], prob.ints?.includes(i) ? 'Integer' : 'Contin'] as Scalar[]), [],
        ['Constraints'], ['Cell', 'Cell Value', 'Formula', 'Status', 'Slack'],
      ];
      vars.forEach((v, i) => sub.set(v, r.x[i])); sub.recalc();
      for (const cr of consRows) {
        const lv = numOf(cr.lhs.sheet.value(cr.lhs.r, cr.lhs.c)), rv = cr.rhs(r.x);
        const slack = Math.abs(rv - lv);
        rows.push([refName(cr.lhs, sheet), lv, `${refName(cr.lhs, sheet)}${cr.rel}${cr.rhsText}`, slack < 1e-7 * (1 + Math.abs(rv)) ? 'Binding' : 'Not Binding', +slack.toPrecision(12)]);
      }
      sub.restore();
      writeMatrix(target, 0, 0, rows);
      target.set(0, 0, { v: 'Answer Report', s: book.styleId({ bold: true, size: 14 }) });
      target.colW.set(0, 120); target.colW.set(2, 110); target.colW.set(3, 110);
    }
  });
  if (target) app.exec('sheet.activate', book.sheets.indexOf(target));
}

// --------------------------------------------------------------------------------------------
// Commands

const T = 'Data';

def([
  {
    id: 'data.whatIf', label: 'What-If Analysis', icon: 'psychology_alt', tab: T, group: 'Forecast', kind: 'menu', noRecord: true, run: () => undefined,
    menu: () => [{ id: 'data.scenarios', label: 'Scenario Manager…', icon: 'view_list' }, { id: 'data.goalSeek', label: 'Goal Seek…', icon: 'my_location' }, { id: 'data.dataTable', label: 'Data Table…', icon: 'table_view' }],
  },
  { id: 'data.scenarios', label: 'Scenario Manager', icon: 'view_list', tab: T, group: 'Forecast', edits: true, keywords: 'what-if scenario summary', run: (app) => scenarioManager(app) },
  { id: 'data.goalSeek', label: 'Goal Seek', icon: 'my_location', tab: T, group: 'Forecast', edits: true, keywords: 'what-if target solve backwards', run: (app) => goalSeekDialog(app) },
  { id: 'data.dataTable', label: 'Data Table', icon: 'table_view', tab: T, group: 'Forecast', edits: true, keywords: 'what-if sensitivity one variable two variable', run: (app) => dataTableDialog(app) },
  { id: 'data.forecastSheet', label: 'Forecast Sheet', icon: 'timeline', tab: T, group: 'Forecast', edits: true, keywords: 'forecast ets exponential smoothing prediction', run: (app) => forecastDialog(app) },
  { id: 'data.solver', label: 'Solver', icon: 'calculate', tab: T, group: 'Analyze', edits: true, keywords: 'optimize linear programming simplex maximize minimize constraints', run: (app) => solverDialog(app) },
]);

void colName;
