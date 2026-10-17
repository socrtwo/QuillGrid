/**
 * Formula insight helpers, inspired by Multiplan (R1C1 notation), Lotus Improv and Javelin (formulas
 * as named rules apart from the grid), Lotus 1-2-3 Release 3 (Map view) and Javelin's diagram view.
 *
 * - toR1C1: a formula in relative R1C1 notation, so copies of one formula read identically.
 * - formulaRules: the distinct formulas of a sheet, each with the ranges it fills.
 * - inconsistentFormulas: cells whose formula breaks the pattern of their neighbors.
 * - cellKind: what Map View shows for a cell (label, number, formula...).
 * - explainTree / dependencyGraph: where a value comes from, and what it feeds.
 * - cellHistory: earlier contents of a cell, read from the undo history.
 */
import { tokenize, type RefToken } from '../calc/parser';
import { cellKey, keyRow, keyCol, cellName, rangeName, quoteSheet, type Range } from '../model/address';
import type { Sheet, Workbook } from '../model/workbook';
import type { Cell, Scalar } from '../model/types';
import { CellError } from '../model/types';
import type { Engine } from '../calc/engine';
import { formatValue } from '../model/numfmt';

// ---------------------------------------------------------------------------------------------
// R1C1

function part(axis: 'R' | 'C', i: number, abs: boolean, base: number): string {
  if (abs) return `${axis}${i + 1}`;
  const d = i - base;
  return d ? `${axis}[${d}]` : axis;
}

export function r1c1RefText(ref: RefToken, r: number, c: number): string {
  const pre = ref.sheet ? quoteSheet(ref.sheet) + (ref.sheet2 ? ':' + quoteSheet(ref.sheet2) : '') + '!' : '';
  if (ref.invalid) return pre + '#REF!';
  const cell = (rr: number, cc: number, ar: boolean, ac: boolean) => part('R', rr, ar, r) + part('C', cc, ac, c);
  switch (ref.kind) {
    case 'cell': return pre + cell(ref.r1, ref.c1, ref.ar1, ref.ac1);
    case 'cols': { const a = part('C', ref.c1, ref.ac1, c), b = part('C', ref.c2, ref.ac2, c); return pre + (a === b ? a : `${a}:${b}`); }
    case 'rows': { const a = part('R', ref.r1, ref.ar1, r), b = part('R', ref.r2, ref.ar2, r); return pre + (a === b ? a : `${a}:${b}`); }
    default: return `${pre}${cell(ref.r1, ref.c1, ref.ar1, ref.ac1)}:${cell(ref.r2, ref.c2, ref.ar2, ref.ac2)}`;
  }
}

/** Converts an A1 formula (without "=") in cell (r, c) to R1C1 notation. Unparseable text is returned unchanged. */
export function toR1C1(formula: string, r: number, c: number): string {
  let toks;
  try { toks = tokenize(formula); } catch { return formula; }
  let out = '', last = 0;
  for (const t of toks) {
    if (t.type !== 'ref' || !t.ref) continue;
    out += formula.slice(last, t.pos) + r1c1RefText(t.ref, r, c);
    last = t.end;
  }
  return out + formula.slice(last);
}

const r1c1Cache = new WeakMap<Sheet, { version: number; map: Map<number, string> }>();

/** R1C1 text of a formula cell, cached per workbook version. */
export function r1c1Of(sheet: Sheet, r: number, c: number): string | null {
  const cell = sheet.get(r, c);
  if (cell?.f == null) return null;
  let e = r1c1Cache.get(sheet);
  if (!e || e.version !== sheet.book.version) { e = { version: sheet.book.version, map: new Map() }; r1c1Cache.set(sheet, e); }
  const k = cellKey(r, c);
  let s = e.map.get(k);
  if (s === undefined) { s = toR1C1(cell.f.replace(/\s+/g, ''), r, c).toUpperCase(); e.map.set(k, s); }
  return s;
}

// ---------------------------------------------------------------------------------------------
// Formula rules

export interface FormulaRule { r1c1: string; formula: string; first: { r: number; c: number }; count: number; ranges: Range[] }

/** Packs cells into rectangles: vertical runs per column, merged across neighboring columns with the same rows. */
export function packRanges(cells: { r: number; c: number }[]): Range[] {
  const byCol = new Map<number, number[]>();
  for (const { r, c } of cells) { let a = byCol.get(c); if (!a) byCol.set(c, (a = [])); a.push(r); }
  const runs: Range[] = [];
  for (const c of [...byCol.keys()].sort((a, b) => a - b)) {
    const rows = byCol.get(c)!.sort((a, b) => a - b);
    let start = rows[0], prev = rows[0];
    for (let i = 1; i <= rows.length; i++) {
      if (i < rows.length && rows[i] === prev + 1) { prev = rows[i]; continue; }
      runs.push({ r1: start, c1: c, r2: prev, c2: c });
      if (i < rows.length) { start = rows[i]; prev = rows[i]; }
    }
  }
  const out: Range[] = [];
  for (const run of runs) {
    const m = out.find((o) => o.r1 === run.r1 && o.r2 === run.r2 && o.c2 === run.c1 - 1);
    if (m) m.c2 = run.c2; else out.push({ ...run });
  }
  return out;
}

/** The distinct formulas of a sheet (copies of one formula count once), largest groups first. */
export function formulaRules(sheet: Sheet): FormulaRule[] {
  const groups = new Map<string, { r: number; c: number }[]>();
  for (const [k, cell] of sheet.cells) {
    if (cell.f == null) continue;
    const r = keyRow(k), c = keyCol(k);
    const key = r1c1Of(sheet, r, c)!;
    let g = groups.get(key);
    if (!g) groups.set(key, (g = []));
    g.push({ r, c });
  }
  const out: FormulaRule[] = [];
  for (const [key, cells] of groups) {
    cells.sort((a, b) => a.r - b.r || a.c - b.c);
    out.push({ r1c1: key, formula: sheet.get(cells[0].r, cells[0].c)!.f!, first: cells[0], count: cells.length, ranges: packRanges(cells) });
  }
  return out.sort((a, b) => b.count - a.count || a.first.r - b.first.r || a.first.c - b.first.c);
}

/**
 * Formula cells that differ from their neighbors when the neighbors on both sides (above and below,
 * or left and right) hold the same formula as each other: the odd one out in a copied block.
 */
export function isInconsistent(sheet: Sheet, r: number, c: number): boolean {
  const me = r1c1Of(sheet, r, c);
  if (me == null) return false;
  const pairs: [number, number, number, number][] = [[r - 1, c, r + 1, c], [r, c - 1, r, c + 1]];
  for (const [ra, ca, rb, cb] of pairs) {
    if (ra < 0 || ca < 0) continue;
    const a = r1c1Of(sheet, ra, ca), b = r1c1Of(sheet, rb, cb);
    if (a != null && a === b && a !== me) return true;
  }
  return false;
}

export function inconsistentFormulas(sheet: Sheet): { r: number; c: number }[] {
  const out: { r: number; c: number }[] = [];
  for (const [k, cell] of sheet.cells) if (cell.f != null && isInconsistent(sheet, keyRow(k), keyCol(k))) out.push({ r: keyRow(k), c: keyCol(k) });
  return out.sort((a, b) => a.r - b.r || a.c - b.c);
}

// ---------------------------------------------------------------------------------------------
// Map View

export type CellKind = 'empty' | 'label' | 'number' | 'logical' | 'error' | 'formula' | 'spill' | 'odd';

export function cellKind(sheet: Sheet, r: number, c: number, cell: Cell | undefined = sheet.get(r, c)): CellKind {
  if (!cell) return 'empty';
  if (cell.f != null) return isInconsistent(sheet, r, c) ? 'odd' : 'formula';
  if (cell.sp != null) return 'spill';
  const v = cell.v;
  if (v == null || v === '') return 'empty';
  if (typeof v === 'string') return 'label';
  if (typeof v === 'number') return 'number';
  if (typeof v === 'boolean') return 'logical';
  return 'error';
}

/** Glyph and colors for each kind: 1-2-3's map used " for labels, # for numbers and + for formulas. */
export const MAP_STYLE: Record<Exclude<CellKind, 'empty'>, { glyph: string; fill: string; dark: string; label: string }> = {
  label: { glyph: '"', fill: '#dbe8f7', dark: '#24364d', label: 'Label (text)' },
  number: { glyph: '#', fill: '#dcefdc', dark: '#22402a', label: 'Number' },
  logical: { glyph: '?', fill: '#efe3f5', dark: '#3c2a47', label: 'TRUE/FALSE' },
  error: { glyph: '!', fill: '#f6d5d2', dark: '#4d2421', label: 'Error value' },
  formula: { glyph: '+', fill: '#fbe6c8', dark: '#4a3517', label: 'Formula' },
  spill: { glyph: '~', fill: '#fdf1de', dark: '#3b2e1b', label: 'Spilled result' },
  odd: { glyph: '≠', fill: '#f9c2b8', dark: '#5c241b', label: 'Formula unlike its neighbors' },
};

// ---------------------------------------------------------------------------------------------
// Explain a value

export interface ExplainNode {
  sheet: Sheet; r: number; c: number;
  /** Range for a range precedent (summarized, not expanded). */
  range?: Range;
  address: string;
  value: string;
  formula?: string;
  note?: string;
  children: ExplainNode[];
}

export function displayValue(sheet: Sheet, r: number, c: number): string {
  const v = sheet.value(r, c);
  if (v instanceof CellError) return v.code;
  if (v == null) return '(blank)';
  if (typeof v === 'string') return JSON.stringify(v);
  return formatValue(v, sheet.styleOf(r, c).numFmt, { date1904: sheet.book.calc.date1904 }).text || String(v);
}

function addr(home: Sheet, sheet: Sheet, rg: Range): string {
  const a = rg.r1 === rg.r2 && rg.c1 === rg.c2 ? cellName(rg.r1, rg.c1) : rangeName(rg);
  return sheet === home ? a : `${quoteSheet(sheet.name)}!${a}`;
}

function rangeSummary(sheet: Sheet, rg: Range): string {
  let n = 0, nums = 0, sum = 0, text = 0, errs = 0;
  sheet.forEachIn(rg, (_r, _c, cell) => {
    const v = cell.v;
    if (v == null || v === '') return;
    n++;
    if (typeof v === 'number') { nums++; sum += v; } else if (v instanceof CellError) errs++; else if (typeof v === 'string') text++;
  });
  const parts = [`${n} value${n === 1 ? '' : 's'}`];
  if (nums) parts.push(`${nums} number${nums === 1 ? '' : 's'}, sum ${+sum.toPrecision(12)}`);
  if (text) parts.push(`${text} text`);
  if (errs) parts.push(`${errs} error${errs === 1 ? '' : 's'}`);
  return parts.join(' · ');
}

/** Tree of where a cell's value comes from: its formula, then each precedent, recursively (ranges are summarized). */
export function explainTree(engine: Engine, sheet: Sheet, r: number, c: number, maxDepth = 6, maxNodes = 250): ExplainNode {
  const seen = new Set<string>();
  let count = 0;
  const visit = (sh: Sheet, rr: number, cc: number, depth: number): ExplainNode => {
    count++;
    const cell = sh.get(rr, cc);
    const node: ExplainNode = { sheet: sh, r: rr, c: cc, address: addr(sheet, sh, { r1: rr, c1: cc, r2: rr, c2: cc }), value: displayValue(sh, rr, cc), children: [] };
    if (cell?.f == null) { node.note = cell?.sp != null ? 'part of a spilled result' : cell?.v == null ? 'blank' : 'input'; return node; }
    node.formula = '=' + cell.f;
    const id = `${sh.id}:${rr}:${cc}`;
    if (seen.has(id)) { node.note = 'shown above'; return node; }
    seen.add(id);
    if (depth >= maxDepth) { node.note = 'more…'; return node; }
    for (const p of engine.precedentsOf(sh, rr, cc)) {
      if (count >= maxNodes) { node.note = 'more…'; break; }
      if (p.r1 === p.r2 && p.c1 === p.c2) node.children.push(visit(p.sheet, p.r1, p.c1, depth + 1));
      else {
        count++;
        const rg = { r1: p.r1, c1: p.c1, r2: p.r2, c2: p.c2 };
        node.children.push({ sheet: p.sheet, r: p.r1, c: p.c1, range: rg, address: addr(sheet, p.sheet, rg), value: rangeSummary(p.sheet, rg), children: [] });
      }
    }
    return node;
  };
  return visit(sheet, r, c, 0);
}

/** Plain-text outline of an explain tree (for copying and for tests). */
export function explainText(n: ExplainNode, indent = ''): string {
  const head = `${indent}${n.address} = ${n.value}${n.formula ? `   ${n.formula}` : ''}${n.note ? `   (${n.note})` : ''}`;
  return [head, ...n.children.map((ch) => explainText(ch, indent + '  '))].join('\n');
}

// ---------------------------------------------------------------------------------------------
// Dependency diagram

export interface GraphNode { id: string; sheet: Sheet; range: Range; level: number; label: string; value: string; formula?: string }
export interface Graph { nodes: GraphNode[]; edges: { from: string; to: string }[] }

/** Precedents (negative levels) and dependents (positive levels) of a cell, up to `levels` steps each way. */
export function dependencyGraph(engine: Engine, sheet: Sheet, r: number, c: number, levels = 2, maxPerLevel = 24): Graph {
  const nodes = new Map<string, GraphNode>();
  const edges: { from: string; to: string }[] = [];
  const idOf = (sh: Sheet, rg: Range) => `${sh.id}!${rg.r1},${rg.c1},${rg.r2},${rg.c2}`;
  const add = (sh: Sheet, rg: Range, level: number): string | null => {
    const id = idOf(sh, rg);
    if (nodes.has(id)) return id;
    if ([...nodes.values()].filter((n) => n.level === level).length >= maxPerLevel) return null;
    const single = rg.r1 === rg.r2 && rg.c1 === rg.c2;
    const f = single ? sh.get(rg.r1, rg.c1)?.f : undefined;
    nodes.set(id, { id, sheet: sh, range: rg, level, label: addr(sheet, sh, rg), value: single ? displayValue(sh, rg.r1, rg.c1) : rangeSummary(sh, rg), formula: f != null ? '=' + f : undefined });
    return id;
  };
  const root = add(sheet, { r1: r, c1: c, r2: r, c2: c }, 0)!;
  let frontier = [{ id: root, sheet, r, c }];
  for (let l = 1; l <= levels; l++) {
    const next: typeof frontier = [];
    for (const f of frontier) for (const p of engine.precedentsOf(f.sheet, f.r, f.c)) {
      const rg = { r1: p.r1, c1: p.c1, r2: p.r2, c2: p.c2 };
      const fresh = !nodes.has(idOf(p.sheet, rg));
      const id = add(p.sheet, rg, -l);
      if (!id) continue;
      edges.push({ from: id, to: f.id });
      if (fresh && rg.r1 === rg.r2 && rg.c1 === rg.c2) next.push({ id, sheet: p.sheet, r: p.r1, c: p.c1 });
    }
    frontier = next;
  }
  frontier = [{ id: root, sheet, r, c }];
  for (let l = 1; l <= levels; l++) {
    const next: typeof frontier = [];
    for (const f of frontier) for (const d of engine.dependentsOfCell(f.sheet, f.r, f.c)) {
      const rg = { r1: d.r, c1: d.c, r2: d.r, c2: d.c };
      const fresh = !nodes.has(idOf(d.sheet, rg));
      const id = add(d.sheet, rg, l);
      if (!id) continue;
      edges.push({ from: f.id, to: id });
      if (fresh) next.push({ id, sheet: d.sheet, r: d.r, c: d.c });
    }
    frontier = next;
  }
  const seen = new Set<string>();
  return { nodes: [...nodes.values()], edges: edges.filter((e) => { const k = e.from + '>' + e.to; if (seen.has(k) || e.from === e.to) return false; seen.add(k); return true; }) };
}

// ---------------------------------------------------------------------------------------------
// Cell history

export interface HistoryEntry { label: string; before: Cell | undefined; after: Cell | undefined }

const sameCell = (a: Cell | undefined, b: Cell | undefined) => (a?.f ?? null) === (b?.f ?? null) && (a?.f != null || scalarKey(a?.v) === scalarKey(b?.v));
const scalarKey = (v: Scalar | undefined) => (v instanceof CellError ? 'e' + v.code : JSON.stringify(v ?? null));

function beforeImage(t: Workbook['undoStack'][number], sheet: Sheet, k: number): { found: boolean; cell?: Cell } {
  const m = t.cells.get(sheet);
  if (m?.has(k)) return { found: true, cell: m.get(k) };
  const snap = t.sheets.get(sheet.id) ?? t.book?.sheets?.[sheet.id];
  if (snap) {
    const hit = (snap.cells as [number, Cell][]).find(([kk]) => kk === k);
    return { found: true, cell: hit?.[1] };
  }
  return { found: false };
}

/** Earlier contents of a cell in this session, newest first, read from the undo history. */
export function cellHistory(book: Workbook, sheet: Sheet, r: number, c: number): HistoryEntry[] {
  const k = cellKey(r, c);
  const out: HistoryEntry[] = [];
  let after = sheet.get(r, c);
  for (let i = book.undoStack.length - 1; i >= 0; i--) {
    const t = book.undoStack[i];
    const b = beforeImage(t, sheet, k);
    if (!b.found || sameCell(b.cell, after)) continue;
    out.push({ label: t.label, before: b.cell, after });
    after = b.cell;
  }
  return out;
}

/** Short text for a cell's contents (formula, value or blank). */
export function cellContentText(cell: Cell | undefined): string {
  if (!cell || (cell.v == null && cell.f == null)) return '(blank)';
  if (cell.f != null) return '=' + cell.f;
  const v = cell.v;
  if (v instanceof CellError) return v.code;
  return typeof v === 'string' ? v : String(v);
}
