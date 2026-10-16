/**
 * Formulas tab: Insert Function wizard and function library menus, defined names (Name Manager, Define
 * Name, Use in Formula, Paste Name, Create from Selection), formula auditing (trace precedents and
 * dependents, error checking, Evaluate Formula, Watch Window) and calculation options.
 */
import { def, getCommand, type MenuEntry } from './registry';
import type { App } from '../app';
import type { Sheet, Workbook } from '../model/workbook';
import type { Engine } from '../calc/engine';
import { openDialog, toast, showMenu, alertDialog, confirmDialog, type DialogControl } from '../ui/dialog';
import { h, icon, clear } from '../ui/dom';
import { FUNCTIONS, getFn, functionNames, type FnDef } from '../calc/functions';
import type { FnCategory } from '../calc/functions/registry';
import { tokenize, type Token } from '../calc/parser';
import { isRef, isRefs, isMatrix, isLambda, type Value } from '../calc/values';
import { CellError, type Scalar, type DefinedName } from '../model/types';
import { cellName, parseCell, keyRow, keyCol, cellKey, quoteSheet, absRangeName, rangeName, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { resolveRangeText } from '../model/structure';
import { activeRange, activeCell, clipToUsed, setInput } from './ops';
import { formatValue, formatGeneral } from '../model/numfmt';
import { enclosingFunction } from '../grid/editor';

const TAB = 'Formulas';

// ============================================================================================
// Pure helpers (exported for tests)

/** Formula-literal text for a value (used by Evaluate Formula and the argument previews). */
export function literalText(v: Scalar): string {
  if (v == null) return '0';
  if (typeof v === 'number') return Number.isFinite(v) ? String(+v.toPrecision(15)) : '#NUM!';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof CellError) return v.code;
  return `"${v.replace(/"/g, '""')}"`;
}

/** Short display of an evaluated value: scalars, arrays as {a,b;c,d}, references by their values. */
export function valuePreview(v: Value, maxItems = 24): string {
  if (isLambda(v)) return 'LAMBDA';
  if (isRefs(v)) return v.areas.map((a) => valuePreview(a, maxItems)).join(',');
  if (isRef(v)) {
    if (v.r1 === v.r2 && v.c1 === v.c2) return literalText(v.sheet.value(v.r1, v.c1));
    const rows: Scalar[][] = [];
    const r2 = Math.min(v.r2, v.r1 + 50), c2 = Math.min(v.c2, v.c1 + 50);
    for (let r = v.r1; r <= r2; r++) { const row: Scalar[] = []; for (let c = v.c1; c <= c2; c++) row.push(v.sheet.value(r, c)); rows.push(row); }
    return matrixText(rows, maxItems);
  }
  if (isMatrix(v)) return matrixText(v, maxItems);
  return literalText(v as Scalar);
}

function matrixText(m: Scalar[][], maxItems: number): string {
  let n = 0;
  const out: string[] = [];
  for (const row of m) {
    const parts: string[] = [];
    for (const x of row) { if (n++ >= maxItems) break; parts.push(literalText(x)); }
    out.push(parts.join(','));
    if (n >= maxItems) break;
  }
  return '{' + out.join(';') + (n >= maxItems ? '…' : '') + '}';
}

/** Validates a defined name like the spreadsheet does. Returns an error message or null. */
export function validateDefinedName(book: Workbook, name: string, scope: string | null, ignore?: DefinedName | null): string | null {
  if (!name) return 'Type a name.';
  if (name.length > 255) return 'A name can have at most 255 characters.';
  if (/\s/.test(name)) return 'A name cannot contain spaces. Use an underscore (_) or a period (.) instead.';
  if (!/^[A-Za-z_\\À-￿][A-Za-z0-9_.\\?À-￿]*$/.test(name)) return 'A name must begin with a letter, an underscore or a backslash, and can contain only letters, numbers, periods and underscores.';
  if (parseCell(name) || /^(R\d*C\d*|R|C)$/i.test(name)) return `"${name}" looks like a cell reference. Choose a name that is not a reference, such as ${name}_.`;
  if (/^(TRUE|FALSE)$/i.test(name)) return 'TRUE and FALSE cannot be used as names.';
  const low = name.toLowerCase();
  const sc = scope ? scope.toLowerCase() : null;
  if (book.names.some((n) => n !== ignore && n.name.toLowerCase() === low && (n.sheet ? n.sheet.toLowerCase() : null) === sc)) return `The name "${name}" already exists in this scope. Enter a unique name.`;
  if (book.sheets.some((s) => s.tables.some((t) => t.name.toLowerCase() === low))) return `A table is already named "${name}".`;
  return null;
}

/** Turns a label into a valid name (spaces and symbols become underscores; references get a suffix). */
export function sanitizeName(label: string): string {
  let n = label.trim().replace(/[^A-Za-z0-9_.\\À-￿]+/g, '_');
  if (!n) return '';
  if (!/^[A-Za-z_\\À-￿]/.test(n)) n = '_' + n;
  if (parseCell(n) || /^(R\d*C\d*|R|C)$/i.test(n)) n += '_';
  if (/^(TRUE|FALSE)$/i.test(n)) n = '_' + n;
  return n.slice(0, 255);
}

export interface CreateNamesOpts { top?: boolean; left?: boolean; bottom?: boolean; right?: boolean }

/** Names that Create from Selection makes for a range (labels in the chosen edge rows/columns). */
export function namesFromSelection(sheet: Sheet, rg: Range, o: CreateNamesOpts): { name: string; ref: string; label: string }[] {
  const out: { name: string; ref: string; label: string }[] = [];
  const r1 = rg.r1 + (o.top ? 1 : 0), r2 = rg.r2 - (o.bottom ? 1 : 0);
  const c1 = rg.c1 + (o.left ? 1 : 0), c2 = rg.c2 - (o.right ? 1 : 0);
  if (r1 > r2 || c1 > c2) return out;
  const pre = quoteSheet(sheet.name) + '!';
  const add = (labelV: Scalar, body: Range) => {
    if (labelV == null || labelV === '' || labelV instanceof CellError) return;
    const label = String(labelV);
    const name = sanitizeName(label);
    if (!name) return;
    if (out.some((x) => x.name.toLowerCase() === name.toLowerCase())) return;
    out.push({ name, ref: pre + absRangeName(body), label });
  };
  if (o.top) for (let c = c1; c <= c2; c++) add(sheet.value(rg.r1, c), { r1, c1: c, r2, c2: c });
  if (o.left) for (let r = r1; r <= r2; r++) add(sheet.value(r, rg.c1), { r1: r, c1, r2: r, c2 });
  if (o.bottom) for (let c = c1; c <= c2; c++) add(sheet.value(rg.r2, c), { r1, c1: c, r2, c2: c });
  if (o.right) for (let r = r1; r <= r2; r++) add(sheet.value(r, rg.c2), { r1: r, c1, r2: r, c2 });
  return out;
}

/** Guesses which edges hold labels (text in the first row and/or first column). */
export function guessLabelEdges(sheet: Sheet, rg: Range): CreateNamesOpts {
  const isText = (v: Scalar) => typeof v === 'string' && v !== '';
  let topText = 0, leftText = 0;
  for (let c = rg.c1 + 1; c <= rg.c2; c++) if (isText(sheet.value(rg.r1, c))) topText++;
  for (let r = rg.r1 + 1; r <= rg.r2; r++) if (isText(sheet.value(r, rg.c1))) leftText++;
  const top = rg.r2 > rg.r1 && (topText > 0 || isText(sheet.value(rg.r1, rg.c1)) && rg.c1 === rg.c2);
  const left = rg.c2 > rg.c1 && leftText > 0;
  return { top: top || (!left && rg.r2 > rg.r1), left };
}

/** Splits the argument list text of a call (between the parentheses) at top-level commas. */
export function splitArgs(text: string): string[] {
  const out: string[] = [];
  let depth = 0, inStr = false, inQ = false, cur = '';
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inStr) { cur += ch; if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else inStr = false; } continue; }
    if (inQ) { cur += ch; if (ch === "'") inQ = false; continue; }
    if (ch === '"') inStr = true;
    else if (ch === "'") inQ = true;
    else if (ch === '(' || ch === '{' || ch === '[') depth++;
    else if (ch === ')' || ch === '}' || ch === ']') depth--;
    else if (ch === ',' && depth === 0) { out.push(cur.trim()); cur = ''; continue; }
    cur += ch;
  }
  if (cur.trim() || out.length) out.push(cur.trim());
  return out;
}

/** Span of the call to `name` (or of the innermost call) around a caret position in formula text. */
export function callSpanAt(text: string, pos: number, name?: string): { start: number; end: number; name: string; args: string[] } | null {
  let p = pos;
  for (let guard = 0; guard < 50; guard++) {
    const fn = enclosingFunction(text, p);
    if (!fn) return null;
    // find the "(" belonging to this call
    let depth = 0, inStr = false, open = -1;
    for (let i = p - 1; i >= 0; i--) {
      const ch = text[i];
      if (ch === '"') { inStr = !inStr; continue; }
      if (inStr) continue;
      if (ch === ')') depth++;
      else if (ch === '(') { if (depth === 0) { open = i; break; } depth--; }
    }
    if (open < 0) return null;
    const m = /([A-Za-z_][A-Za-z0-9_.]*)\s*$/.exec(text.slice(0, open));
    const start = m ? open - m[0].length : open;
    if (!name || fn.name === name.toUpperCase()) {
      let d = 0, s = false, close = text.length;
      for (let i = open; i < text.length; i++) {
        const ch = text[i];
        if (ch === '"') { s = !s; continue; }
        if (s) continue;
        if (ch === '(') d++;
        else if (ch === ')') { d--; if (d === 0) { close = i + 1; break; } }
      }
      const inner = text.slice(open + 1, close > open && text[close - 1] === ')' ? close - 1 : close);
      return { start, end: close, name: fn.name, args: splitArgs(inner) };
    }
    p = start;
  }
  return null;
}

/** Display names of a function's parameters for `count` arguments (variadic tails grow). */
export function paramNames(fn: FnDef, count: number): string[] {
  const out = fn.params.filter((p) => p !== '...').map((p) => p.replace(/^\[|\]$/g, ''));
  const target = Math.min(count, fn.max);
  if (target <= out.length || !out.length) return out;
  // Repeat the trailing group of numbered parameters (number2 → number3; criteria_range2, criteria2 → …3).
  const num = (p: string) => /^(.*?)(\d+)$/.exec(p);
  const lastM = num(out[out.length - 1]);
  if (!lastM) { while (out.length < target) out.push(`${out[out.length - 1].replace(/\d*$/, '')}${out.length + 1}`); return out; }
  const n = Number(lastM[2]);
  const group: string[] = [];
  for (let i = out.length - 1; i >= 0; i--) { const m = num(out[i]); if (!m || Number(m[2]) !== n) break; group.unshift(m[1]); }
  for (let k = n + 1; out.length < target; k++) for (const g of group) if (out.length < target) out.push(g + k);
  return out;
}

export function isRequired(fn: FnDef, i: number): boolean {
  const base = fn.params.filter((p) => p !== '...');
  if (i < base.length) return !base[i].startsWith('[') && i < fn.min;
  return false;
}

export function syntaxOf(fn: FnDef): string {
  return `${fn.name}(${fn.params.join(', ')})`;
}

// ============================================================================================
// Evaluate Formula stepper

type SNode =
  | { k: 'lit'; s: number; e: number; neg?: boolean }
  | { k: 'missing'; s: number; e: number }
  | { k: 'ref'; s: number; e: number; multi: boolean; name?: boolean }
  | { k: 'fn'; s: number; e: number; name: string; args: SNode[] }
  | { k: 'un'; s: number; e: number; a: SNode }
  | { k: 'pct'; s: number; e: number; a: SNode }
  | { k: 'bin'; s: number; e: number; op: string; a: SNode; b: SNode }
  | { k: 'paren'; s: number; e: number; a: SNode }
  | { k: 'opaque'; s: number; e: number };

const BIN_PREC: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

/** Parses formula text into a span tree (positions into the text) for step-by-step evaluation. */
export function parseSpans(src: string): SNode {
  const toks = tokenize(src);
  let p = 0;
  const skip = () => { while (toks[p].type === 'ws') p++; };
  const peek = (): Token => { let k = p; while (toks[k].type === 'ws') k++; return toks[k]; };
  const next = (): Token => { skip(); return toks[p++]; };
  const expr = (min = 0): SNode => {
    let left = unary();
    for (;;) {
      const t = peek();
      if (t.type !== 'op' || !(t.text in BIN_PREC)) break;
      const prec = BIN_PREC[t.text];
      if (prec < min) break;
      next();
      const right = expr(prec + 1);
      left = { k: 'bin', s: left.s, e: right.e, op: t.text, a: left, b: right };
    }
    return left;
  };
  const unary = (): SNode => {
    const t = peek();
    if (t.type === 'op' && (t.text === '-' || t.text === '+')) {
      next();
      const a = unary();
      if (t.text === '+') return a;
      if (a.k === 'lit' && /^\d|^\./.test(src.slice(a.s, a.e))) return { k: 'lit', s: t.pos, e: a.e, neg: true };
      return { k: 'un', s: t.pos, e: a.e, a };
    }
    if (t.type === 'op' && t.text === '@') { next(); const a = unary(); return { k: 'opaque', s: t.pos, e: a.e }; }
    return postfix();
  };
  const postfix = (): SNode => {
    let n = refExpr();
    for (;;) {
      const pk = peek();
      if (pk.type === 'op' && pk.text === '%') { next(); n = { k: 'pct', s: n.s, e: pk.end, a: n }; continue; }
      const t = toks[p];
      if (t.type === 'op' && t.text === '#') { p++; n = { k: 'opaque', s: n.s, e: t.end }; continue; }
      if (t.type === 'lparen' && (n.k === 'fn' || n.k === 'ref')) { p++; const e = args(); n = { k: 'opaque', s: n.s, e: e.end }; continue; }
      break;
    }
    return n;
  };
  const refExpr = (): SNode => {
    let n = primary();
    for (;;) {
      const t = toks[p];
      if (t.type === 'op' && t.text === ':') { p++; const b = primary(); n = { k: 'ref', s: n.s, e: b.e, multi: true }; continue; }
      break;
    }
    return n;
  };
  const args = (): { list: SNode[]; end: number } => {
    const list: SNode[] = [];
    if (peek().type === 'rparen') { const t = next(); return { list, end: t.end }; }
    for (;;) {
      const t = peek();
      if (t.type === 'comma' || t.type === 'rparen') list.push({ k: 'missing', s: t.pos, e: t.pos });
      else list.push(expr());
      const sep = next();
      if (sep.type === 'rparen') return { list, end: sep.end };
      if (sep.type !== 'comma') throw new Error('Expected , or )');
    }
  };
  const primary = (): SNode => {
    const t = next();
    switch (t.type) {
      case 'num': case 'str': case 'bool': case 'err': return { k: 'lit', s: t.pos, e: t.end };
      case 'ref': {
        if (t.ref?.invalid) return { k: 'lit', s: t.pos, e: t.end };
        const r = t.ref!;
        return { k: 'ref', s: t.pos, e: t.end, multi: !(r.r1 === r.r2 && r.c1 === r.c2) || !!r.sheet2 };
      }
      case 'name': return { k: 'ref', s: t.pos, e: t.end, multi: false, name: true };
      case 'sref': return { k: 'ref', s: t.pos, e: t.end, multi: true };
      case 'func': {
        const lp = next();
        const a = args();
        void lp;
        return { k: 'fn', s: t.pos, e: a.end, name: t.text.toUpperCase().replace(/^_XLFN\./, ''), args: a.list };
      }
      case 'lparen': {
        const first = expr();
        if (peek().type === 'comma') {
          while (peek().type === 'comma') { next(); expr(); }
          const rp = next();
          return { k: 'opaque', s: t.pos, e: rp.end };
        }
        const rp = next();
        return { k: 'paren', s: t.pos, e: rp.end, a: first };
      }
      case 'lbrace': {
        let depth = 1, q = p;
        while (q < toks.length && depth) { if (toks[q].type === 'lbrace') depth++; else if (toks[q].type === 'rbrace') depth--; q++; }
        p = q;
        return { k: 'lit', s: t.pos, e: toks[q - 1].end };
      }
      case 'op':
        if (t.text === '-' || t.text === '+') { const a = unary(); return t.text === '-' ? { k: 'un', s: t.pos, e: a.e, a } : a; }
        throw new Error(`Unexpected "${t.text}"`);
      default: throw new Error(t.type === 'eof' ? 'The formula is incomplete' : `Unexpected "${t.text}"`);
    }
  };
  const root = expr();
  if (peek().type !== 'eof') throw new Error(`Unexpected "${peek().text}"`);
  return root;
}

const LAZY_WHOLE = new Set(['LET', 'LAMBDA', 'MAP', 'REDUCE', 'SCAN', 'MAKEARRAY', 'BYROW', 'BYCOL', 'IFS', 'SWITCH', 'CHOOSE']);

export interface StepPick { s: number; e: number; kind: 'ref' | 'expr'; stepIn?: { sheet: Sheet; r: number; c: number } }

/**
 * Step-by-step formula evaluation. Each step finds the innermost, leftmost reducible part of the
 * formula text (a reference, a function call whose arguments are all constants, or an operator with
 * constant operands), evaluates it in the cell's context, and replaces it with its value.
 */
export class FormulaStepper {
  text: string;
  error: string | null = null;
  steps = 0;
  constructor(public engine: Engine, public sheet: Sheet, public r: number, public c: number, formula?: string) {
    this.text = (formula ?? sheet.get(r, c)?.f ?? '').replace(/^=/, '');
    this.pick();
  }

  private picked: StepPick | null = null;

  get done(): boolean { return !this.picked; }
  get next(): StepPick | null { return this.picked; }

  /** Picks the next reducible part. */
  pick(): StepPick | null {
    this.picked = null;
    if (!this.text) return null;
    let root: SNode;
    try { root = parseSpans(this.text); } catch (e) { this.error = (e as Error).message; return null; }
    const n = this.find(root, null, -1);
    if (n) {
      const kind = n.k === 'ref' ? 'ref' : 'expr';
      const pick: StepPick = { s: n.s, e: n.e, kind };
      if (n.k === 'ref' && !n.multi && !n.name) {
        const tok = tokenize(this.text.slice(n.s, n.e)).find((t) => t.type === 'ref');
        const ref = tok?.ref;
        if (ref && !ref.invalid) {
          const sh = ref.sheet ? this.sheet.book.sheetByName(ref.sheet) : this.sheet;
          if (sh && sh.get(ref.r1, ref.c1)?.f != null) pick.stepIn = { sheet: sh, r: ref.r1, c: ref.c1 };
        }
      }
      this.picked = pick;
    }
    return this.picked;
  }

  private terminal(n: SNode, parent: SNode | null, argIndex: number): boolean {
    if (n.k === 'lit' || n.k === 'missing') return true;
    if (n.k === 'ref') return !this.substitutable(n, parent, argIndex);
    if (n.k === 'opaque' && parent?.k === 'fn') return true;
    return false;
  }

  private substitutable(n: Extract<SNode, { k: 'ref' }>, parent: SNode | null, argIndex: number): boolean {
    if (parent?.k === 'fn') {
      const f = getFn(parent.name);
      if (f?.refs?.includes(argIndex)) return false;
      if (n.multi) return false;
    } else if (n.multi) {
      // A bare range outside a function (=A1:A3 or =A1:A3*2) evaluates to an array.
      return true;
    }
    if (n.name) {
      const v = this.engine.evaluate(this.text.slice(n.s, n.e), this.sheet, this.r, this.c);
      if (isRef(v) && !(v.r1 === v.r2 && v.c1 === v.c2) && parent?.k === 'fn') return false;
      if (isLambda(v)) return false;
    }
    return true;
  }

  private find(n: SNode, parent: SNode | null, argIndex: number): SNode | null {
    switch (n.k) {
      case 'lit': case 'missing': return null;
      case 'ref': return this.substitutable(n, parent, argIndex) ? n : null;
      case 'opaque': return parent?.k === 'fn' ? null : n;
      case 'un': case 'pct': return this.find(n.a, n, 0) || (this.terminal(n.a, n, 0) ? n : null);
      case 'paren': {
        const x = this.find(n.a, n, 0);
        if (x) return x;
        return n.a.k === 'lit' && !n.a.neg ? n : this.terminal(n.a, n, 0) ? n : null;
      }
      case 'bin': return this.find(n.a, n, 0) || this.find(n.b, n, 1) || n;
      case 'fn': {
        if (LAZY_WHOLE.has(n.name)) return n;
        if (n.name === 'IF' && n.args.length) {
          const c0 = this.find(n.args[0], n, 0);
          if (c0) return c0;
          const cond = this.engine.evaluateScalar(this.text.slice(n.args[0].s, n.args[0].e) || 'FALSE', this.sheet, this.r, this.c);
          const truthy = cond === true || (typeof cond === 'number' && cond !== 0) || (typeof cond === 'string' && cond.toUpperCase() === 'TRUE');
          const branch = n.args[truthy ? 1 : 2];
          if (branch) { const x = this.find(branch, n, truthy ? 1 : 2); if (x) return x; }
          return n;
        }
        if (n.name === 'IFERROR' || n.name === 'IFNA' || n.name === 'AND' || n.name === 'OR') {
          const x = n.args.length ? this.find(n.args[0], n, 0) : null;
          if (x) return x;
          if (n.name === 'IFERROR' || n.name === 'IFNA') return n;
        }
        for (let i = 0; i < n.args.length; i++) { const x = this.find(n.args[i], n, i); if (x) return x; }
        return n;
      }
    }
    return null;
  }

  /** Evaluates the picked part and replaces it with its value. */
  step(): void {
    const p = this.picked;
    if (!p) return;
    const part = this.text.slice(p.s, p.e);
    let lit: string;
    const v = this.engine.evaluate(part, this.sheet, this.r, this.c);
    if (isRef(v) && !(v.r1 === v.r2 && v.c1 === v.c2)) {
      // A function returning a multi-cell reference: show the reference so outer functions still see a range.
      const rg = rangeName({ r1: v.r1, c1: v.c1, r2: v.r2, c2: v.c2 });
      lit = v.sheet === this.sheet ? rg : `${quoteSheet(v.sheet.name)}!${rg}`;
      if (lit === part.replace(/\$/g, '')) lit = valuePreview(v, 10000);
    } else if (isRef(v)) lit = literalText(v.sheet.value(v.r1, v.c1));
    else lit = valuePreview(v, 10000);
    if (lit.endsWith('…}')) lit = lit.replace('…}', '}');
    this.text = this.text.slice(0, p.s) + lit + this.text.slice(p.e);
    this.steps++;
    this.pick();
  }

  /** Runs to the end; returns the final text. */
  run(max = 500): string {
    while (!this.done && max-- > 0) this.step();
    return this.text;
  }
}

// ============================================================================================
// Shared UI helpers

function openCommandMenu(app: App, id: string): void {
  const c = getCommand(id);
  if (!c?.menu) return;
  showMenu(c.menu(app), { anchor: app.ui.anchorFor(id), onPick: (e) => { if (e.id) app.exec(e.id, e.arg); } });
}

function protectedCell(app: App, sheet: Sheet, r: number, c: number): boolean {
  return sheet.protection.on && sheet.styleOf(r, c).locked !== false;
}

function guardWrite(app: App, sheet: Sheet, r: number, c: number): boolean {
  if (app.isReadOnly()) { toast('This workbook is read-only.'); return false; }
  if (protectedCell(app, sheet, r, c)) { toast('The cell is on a protected sheet. Unprotect the sheet to change it.'); return false; }
  return true;
}

/** Inserts text into the in-cell editor at the caret (or replacing a span). */
function insertIntoEditor(app: App, text: string, span?: { start: number; end: number } | null, caret?: number): void {
  const ed = app.grid.editor;
  const inp = ed.input;
  const a = span ? span.start : caret ?? inp.selectionStart ?? inp.value.length;
  const b = span ? span.end : caret ?? inp.selectionEnd ?? a;
  inp.focus();
  inp.setRangeText(text, a, b, 'end');
  inp.dispatchEvent(new Event('input', { bubbles: true }));
}

const fmtVal = (sheet: Sheet, v: Scalar) => formatValue(v, undefined, { date1904: sheet.book.calc.date1904 }).text;

// ============================================================================================
// Recently used functions

const RECENT_KEY = 'quilllattice.recentFunctions';
const DEFAULT_RECENT = ['SUM', 'AVERAGE', 'IF', 'HYPERLINK', 'COUNT', 'MAX', 'SIN', 'SUMIF', 'PMT', 'STDEV'];

export function recentFunctions(): string[] {
  try {
    const raw = localStorage.getItem(RECENT_KEY);
    const list = raw ? JSON.parse(raw) : null;
    if (Array.isArray(list) && list.length) return list.filter((x) => typeof x === 'string' && getFn(x));
  } catch { /* storage unavailable */ }
  return DEFAULT_RECENT.filter((x) => getFn(x));
}

export function noteRecentFunction(name: string): void {
  const list = [name.toUpperCase(), ...recentFunctions().filter((x) => x !== name.toUpperCase())].slice(0, 10);
  try { localStorage.setItem(RECENT_KEY, JSON.stringify(list)); } catch { /* ignore */ }
}

// ============================================================================================
// Insert Function wizard

const CATEGORIES: FnCategory[] = ['Financial', 'Date & Time', 'Math & Trig', 'Statistical', 'Lookup & Reference', 'Database', 'Text', 'Logical', 'Information', 'Engineering', 'Cube', 'Compatibility', 'Web', 'Dynamic Array'];

function uniqueFns(): FnDef[] {
  return functionNames().map((n) => getFn(n)!).filter((f) => f && f.name.toUpperCase() === f.name.toUpperCase());
}

function fnsIn(cat: FnCategory): FnDef[] {
  const seen = new Set<string>();
  const out: FnDef[] = [];
  for (const f of FUNCTIONS.values()) if (f.category === cat && !seen.has(f.name)) { seen.add(f.name); out.push(f); }
  return out.sort((a, b) => a.name.localeCompare(b.name));
}

interface InsertArg { name?: string; fromEditor?: boolean; browse?: boolean }

async function insertFunction(app: App, arg: InsertArg = {}): Promise<void> {
  const sheet = app.sheet;
  const { r, c } = activeCell(app);
  const editing = !!(app.grid?.editing) || !!arg.fromEditor;
  const browse = !!arg.browse;
  // Where the result goes
  let caret: number | undefined;
  let span: { start: number; end: number } | null = null;
  let initArgs: string[] = [];
  let existingCell = false;
  if (editing && app.grid?.editing) {
    const inp = app.grid.editor.input;
    caret = inp.selectionStart ?? inp.value.length;
    if (arg.name) {
      const sp = callSpanAt(inp.value, caret, arg.name);
      if (sp) { span = sp; initArgs = sp.args; }
    }
  } else if (!browse) {
    const f = sheet.get(r, c)?.f;
    if (f) {
      const m = /^\s*([A-Za-z_][A-Za-z0-9_.]*)\s*\(/.exec(f);
      if (m) {
        const sp = callSpanAt(f, f.length - (f.trimEnd().endsWith(')') ? 1 : 0), undefined);
        if (sp && sp.start === f.indexOf(m[1]) && sp.end >= f.trimEnd().length && (!arg.name || sp.name === arg.name.toUpperCase())) { arg = { ...arg, name: arg.name || sp.name }; initArgs = sp.args; existingCell = true; }
      }
    }
  }

  let step: 1 | 2 = arg.name && getFn(arg.name) && !browse ? 2 : 1;
  let selected = arg.name ? getFn(arg.name)?.name || '' : recentFunctions()[0] || 'SUM';
  let fnArgs: string[] = initArgs;
  let ctlRef: DialogControl | null = null;
  const content = h('div', { class: 'ql-fx-wizard' });
  const evalArg = (t: string): string => {
    if (!t.trim()) return '';
    try { return valuePreview(app.engine.evaluate(t, sheet, r, c), 12); } catch { return '#VALUE!'; }
  };

  const renderStep1 = () => {
    clear(content);
    const search = h('input', { class: 'ql-input', type: 'search', placeholder: 'Describe what you want to do, or type a function name', 'aria-label': 'Search for a function' }) as HTMLInputElement;
    const cat = h('select', { 'aria-label': 'Category' }, h('option', { value: 'recent' }, 'Most Recently Used'), h('option', { value: 'all' }, 'All'), ...CATEGORIES.filter((k) => fnsIn(k).length).map((k) => h('option', { value: k }, k))) as HTMLSelectElement;
    const list = h('select', { class: 'ql-listbox ql-fx-list', size: '9', 'aria-label': 'Functions' }) as HTMLSelectElement;
    const syntax = h('div', { class: 'ql-fx-syntax' });
    const desc = h('div', { class: 'ql-help ql-fx-desc' });
    const fill = (names: string[]) => {
      clear(list);
      for (const n of names) list.appendChild(h('option', { value: n }, n));
      const keep = names.includes(selected) ? selected : names[0];
      if (keep) { list.value = keep; selected = keep; }
      show();
    };
    const show = () => {
      const f = getFn(selected);
      syntax.textContent = f ? syntaxOf(f) : '';
      desc.textContent = f ? `${f.desc} (${f.category})` : 'No function selected.';
    };
    const byCat = () => {
      const v = cat.value;
      if (v === 'recent') fill(recentFunctions());
      else if (v === 'all') fill(uniqueFns().map((f) => f.name));
      else fill(fnsIn(v as FnCategory).map((f) => f.name));
    };
    const doSearch = () => {
      const q = search.value.trim().toLowerCase();
      if (!q) { byCat(); return; }
      const words = q.split(/\s+/);
      const scored = uniqueFns().map((f) => {
        const hay = `${f.name} ${f.desc}`.toLowerCase();
        let s = 0;
        for (const w of words) { if (f.name.toLowerCase() === w) s += 20; else if (f.name.toLowerCase().startsWith(w)) s += 8; else if (hay.includes(w)) s += 2; else return null; }
        return { f, s };
      }).filter((x): x is { f: FnDef; s: number } => !!x).sort((a, b) => b.s - a.s || a.f.name.localeCompare(b.f.name));
      fill(scored.slice(0, 60).map((x) => x.f.name));
      if (!scored.length) desc.textContent = 'No functions match. Try other words or pick a category.';
    };
    search.addEventListener('input', doSearch);
    search.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); doSearch(); } if (e.key === 'ArrowDown') { e.preventDefault(); list.focus(); } });
    cat.addEventListener('change', () => { search.value = ''; byCat(); });
    list.addEventListener('change', () => { selected = list.value; show(); });
    list.addEventListener('dblclick', () => { if (!browse) goStep2(); });
    list.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); if (!browse) goStep2(); } });
    content.append(
      h('label', { class: 'ql-label' }, 'Search for a function'), h('div', { class: 'ql-row' }, search, h('button', { class: 'ql-btn', type: 'button', onclick: doSearch }, 'Go')),
      h('label', { class: 'ql-label' }, 'Or select a category'), cat,
      h('label', { class: 'ql-label' }, 'Select a function'), list, syntax, desc,
      h('button', { class: 'ql-link', type: 'button', onclick: () => app.exec('formulas.functionHelp', { name: selected }) }, 'Help on this function'),
    );
    if (selected && !recentFunctions().includes(selected)) cat.value = getFn(selected)?.category || 'all';
    byCat();
    setTimeout(() => search.focus(), 40);
  };

  const goStep2 = () => {
    if (!getFn(selected)) { toast('Select a function first.'); return; }
    step = 2;
    if (!fnArgs.length || (arg.name && arg.name.toUpperCase() !== selected)) fnArgs = [];
    renderStep2();
    updateButtons();
  };

  const renderStep2 = () => {
    const f = getFn(selected)!;
    clear(content);
    const base = f.params.filter((x) => x !== '...').length;
    const want = Math.max(Math.min(Math.max(base, 1), f.max), fnArgs.length);
    while (fnArgs.length < want) fnArgs.push('');
    const rows = h('div', { class: 'ql-fx-args' });
    const result = h('div', { class: 'ql-fx-result' });
    const paramHelp = h('div', { class: 'ql-help' });
    const inputs: HTMLInputElement[] = [];
    const update = () => {
      const call = buildCall();
      let res = '';
      try { const v = app.engine.evaluate(call, sheet, r, c); res = isRef(v) && v.r1 === v.r2 && v.c1 === v.c2 ? fmtVal(sheet, v.sheet.value(v.r1, v.c1)) : isMatrix(v) || isRef(v) ? valuePreview(v, 12) : fmtVal(sheet, v as Scalar); } catch { res = '#VALUE!'; }
      result.textContent = `Formula result = ${res}`;
    };
    const build = () => {
      clear(rows);
      inputs.length = 0;
      const names = paramNames(f, fnArgs.length);
      fnArgs.forEach((val, i) => {
        const inp = h('input', { class: 'ql-input', value: val, 'aria-label': names[i] || `argument ${i + 1}`, spellcheck: 'false' }) as HTMLInputElement;
        const out = h('span', { class: 'ql-fx-val' }, val ? '= ' + evalArg(val) : '');
        inp.addEventListener('input', () => {
          fnArgs[i] = inp.value;
          out.textContent = inp.value ? '= ' + evalArg(inp.value) : '';
          // Variadic: grow when the last box gets text.
          if (i === fnArgs.length - 1 && inp.value && fnArgs.length < f.max && (f.params.includes('...') || fnArgs.length < f.params.length)) {
            fnArgs.push('');
            const at = inp.selectionStart;
            build();
            inputs[i].focus(); if (at != null) inputs[i].setSelectionRange(at, at);
          }
          update();
        });
        inp.addEventListener('focus', () => { paramHelp.textContent = `${names[i]}${isRequired(f, i) ? ' (required)' : ' (optional)'}`; });
        inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); finish(); } });
        inputs.push(inp);
        rows.appendChild(h('div', { class: 'ql-fx-arg' }, h('label', { class: 'ql-fx-name' + (isRequired(f, i) ? ' ql-req' : '') }, names[i] || `arg${i + 1}`), inp, out));
      });
    };
    build();
    content.append(h('div', { class: 'ql-fx-syntax' }, syntaxOf(f)), rows, h('div', { class: 'ql-help' }, f.desc), paramHelp, result,
      h('button', { class: 'ql-link', type: 'button', onclick: () => app.exec('formulas.functionHelp', { name: f.name }) }, 'Help on this function'));
    update();
    setTimeout(() => inputs[0]?.focus(), 40);
  };

  const buildCall = (): string => {
    const a = [...fnArgs];
    while (a.length && !a[a.length - 1].trim()) a.pop();
    return `${selected}(${a.join(',')})`;
  };

  const finish = () => {
    if (browse) { ctlRef?.close({}); return; }
    if (step === 1) { goStep2(); return; }
    const call = buildCall();
    noteRecentFunction(selected);
    ctlRef?.close({ call });
  };

  const footerBtns = () => ctlRef?.body.parentElement?.querySelector('.ql-dialog-foot') as HTMLElement | null;
  const updateButtons = () => {
    const foot = footerBtns();
    if (!foot) return;
    const back = foot.querySelector<HTMLElement>('[data-fx="back"]');
    if (back) back.style.display = step === 2 && !arg.name ? '' : 'none';
    const ok = foot.querySelector<HTMLElement>('[data-fx="ok"]');
    if (ok) ok.textContent = browse ? 'Close' : step === 1 ? 'Next' : 'OK';
    const title = ctlRef?.body.parentElement?.querySelector('.ql-dialog-head h2');
    if (title) title.textContent = browse ? 'Function Reference' : step === 1 ? 'Insert Function' : 'Function Arguments';
  };

  const res = await openDialog({
    title: browse ? 'Function Reference' : step === 1 ? 'Insert Function' : 'Function Arguments', wide: true, ok: false, cancel: browse ? false : 'Cancel', content,
    buttons: [
      { label: 'Back', onClick: () => { step = 1; renderStep1(); updateButtons(); } },
      { label: browse ? 'Close' : step === 1 ? 'Next' : 'OK', primary: true, onClick: () => finish() },
    ],
    onOpen: (ctl) => {
      ctlRef = ctl;
      const foot = footerBtns();
      const btns = foot ? [...foot.querySelectorAll<HTMLElement>('button')] : [];
      if (btns[0]) btns[0].dataset.fx = 'back';
      if (btns[1]) btns[1].dataset.fx = 'ok';
      if (step === 1) renderStep1(); else renderStep2();
      updateButtons();
    },
  });
  if (!res || !res.call) {
    if (editing && app.grid.editing) app.grid.editor.focus();
    return;
  }
  const call = res.call as string;
  if (editing && app.grid.editing) {
    const inp = app.grid.editor.input;
    if (span) insertIntoEditor(app, call, span);
    else if (!inp.value || inp.value === '=') insertIntoEditor(app, inp.value === '=' ? call : '=' + call, { start: 0, end: inp.value.length });
    else insertIntoEditor(app, call, null, caret);
    return;
  }
  if (!guardWrite(app, sheet, r, c)) return;
  const f = sheet.get(r, c)?.f;
  const text = existingCell && f ? '=' + call : '=' + call;
  app.book.transact('Insert Function', () => setInput(app, sheet, r, c, text));
}

// ============================================================================================
// Function help

async function functionHelp(app: App, arg?: { name?: string }): Promise<void> {
  let name = arg?.name;
  if (!name && app.grid?.editing) {
    const inp = app.grid.editor.input;
    name = enclosingFunction(inp.value, inp.selectionStart ?? inp.value.length)?.name;
    if (!name) { const m = /([A-Za-z_][A-Za-z0-9_.]*)\s*\(?$/.exec(inp.value.slice(0, inp.selectionStart ?? inp.value.length)); name = m?.[1]; }
  }
  if (!name) { const f = app.sheet.get(app.sheet.sel.r, app.sheet.sel.c)?.f; const m = f && /([A-Za-z_][A-Za-z0-9_.]*)\s*\(/.exec(f); name = m ? m[1] : undefined; }
  const fn = name ? getFn(name) : undefined;
  if (!fn) { toast(name ? `No help is available for "${name}".` : 'Place the cursor in a function, or select a cell with a formula.'); return; }
  const content = h('div', { class: 'ql-msg' },
    h('p', null, h('code', { class: 'ql-fx-syntax' }, syntaxOf(fn))),
    h('p', null, fn.desc),
    h('p', { class: 'ql-help' }, `Category: ${fn.category}. Takes ${fn.min === fn.max ? fn.min : `${fn.min} to ${fn.max > 200 ? 'many' : fn.max}`} argument${fn.max === 1 ? '' : 's'}.`),
    fn.aliases?.length ? h('p', { class: 'ql-help' }, `Also available as ${fn.aliases.join(', ')}.`) : null,
    h('ul', null, ...paramNames(fn, Math.max(fn.params.filter((p) => p !== '...').length, 1)).map((p, i) => h('li', null, h('b', null, p), isRequired(fn, i) ? ' — required' : ' — optional'))));
  await openDialog({ title: `${fn.name} function`, content, ok: 'Close', cancel: false });
}

// ============================================================================================
// Defined names

function nameScopeLabel(n: DefinedName): string { return n.sheet || 'Workbook'; }

function nameValueText(app: App, n: DefinedName): string {
  const home = (n.sheet && app.book.sheetByName(n.sheet)) || app.sheet;
  try {
    const v = app.engine.evaluate(n.ref.replace(/^=/, ''), home, 0, 0);
    return valuePreview(v, 8);
  } catch { return '#REF!'; }
}

function selectionRefText(app: App): string {
  const s = app.sheet;
  const rg = activeRange(app);
  const body = rg.r1 === 0 && rg.r2 >= MAX_ROWS - 1 || rg.c1 === 0 && rg.c2 >= MAX_COLS - 1 ? rangeName(rg).replace(/([A-Z]+|\d+)/g, '$$$1') : absRangeName(rg);
  return `${quoteSheet(s.name)}!${body}`;
}

async function nameDialog(app: App, existing?: DefinedName, preset?: Partial<DefinedName>): Promise<boolean> {
  const book = app.book;
  const s = app.sheet;
  let suggest = '';
  if (!existing && !preset?.name) {
    const rg = activeRange(app);
    const v = s.value(rg.r1, rg.c1);
    const left = rg.c1 > 0 ? s.value(rg.r1, rg.c1 - 1) : null;
    const up = rg.r1 > 0 ? s.value(rg.r1 - 1, rg.c1) : null;
    const pickText = [rg.r1 !== rg.r2 || rg.c1 !== rg.c2 ? v : null, left, up].find((x) => typeof x === 'string' && x.trim());
    if (typeof pickText === 'string') suggest = sanitizeName(pickText);
  }
  const scopes = [{ value: '', label: 'Workbook' }, ...book.sheets.map((x) => ({ value: x.name, label: x.name }))];
  const r = await openDialog({
    title: existing ? 'Edit Name' : 'New Name',
    fields: [
      { key: 'name', label: 'Name', type: 'text', value: existing?.name ?? preset?.name ?? suggest, autofocus: true },
      { key: 'scope', label: 'Scope', type: 'select', options: scopes, value: existing?.sheet ?? preset?.sheet ?? '' },
      { key: 'comment', label: 'Comment', type: 'textarea', rows: 3, value: existing?.comment ?? preset?.comment ?? '' },
      { key: 'ref', label: 'Refers to', type: 'text', value: '=' + (existing?.ref ?? preset?.ref ?? selectionRefText(app)).replace(/^=/, '') },
    ],
    validate: (v) => {
      const err = validateDefinedName(book, String(v.name || '').trim(), v.scope || null, existing);
      if (err) return err;
      const ref = String(v.ref || '').trim().replace(/^=/, '');
      if (!ref) return 'Type what the name refers to (a reference, constant or formula).';
      if (app.engine.compile(ref) instanceof Error) return 'There is a problem with the "Refers to" formula.';
      if (existing?.sheet !== undefined && existing.sheet !== (v.scope || null) && existing && (existing.sheet || '') !== (v.scope || '')) return 'The scope of an existing name cannot be changed. Delete it and create a new one.';
      return null;
    },
  });
  if (!r) return false;
  const name = String(r.name).trim();
  const ref = String(r.ref).trim().replace(/^=/, '');
  book.transact(existing ? 'Edit Name' : 'Define Name', () => {
    book.recordBook();
    const target = existing ? book.names.find((n) => n.name === existing.name && (n.sheet || null) === (existing.sheet || null)) : null;
    const rec: DefinedName = { name, ref, sheet: r.scope || null };
    if (r.comment) rec.comment = String(r.comment);
    if (target) Object.assign(target, rec, r.comment ? {} : { comment: undefined });
    else book.names.push(rec);
  });
  return true;
}

async function nameManager(app: App): Promise<void> {
  const book = app.book;
  let filter = 'all';
  let selectedIdx = -1;
  const content = h('div', { class: 'ql-namemgr' });
  const tableWrap = h('div', { class: 'ql-namemgr-list' });
  const refInput = h('input', { class: 'ql-input', 'aria-label': 'Refers to', spellcheck: 'false' }) as HTMLInputElement;
  const errOf = (n: DefinedName) => /^#|^{?#/.test(nameValueText(app, n));
  const rows = () => {
    type Row = { kind: 'name'; n: DefinedName; i: number } | { kind: 'table'; name: string; ref: string; sheet: string };
    const out: Row[] = book.names.map((n, i) => ({ kind: 'name' as const, n, i })).filter((x) => !x.n.hidden);
    const tables: Row[] = book.sheets.flatMap((s) => s.tables.map((t) => ({ kind: 'table' as const, name: t.name, ref: `${quoteSheet(s.name)}!${absRangeName(t.range)}`, sheet: s.name })));
    const all = [...out, ...tables];
    return all.filter((x) => {
      switch (filter) {
        case 'sheet': return x.kind === 'name' && !!x.n.sheet;
        case 'book': return x.kind === 'table' || !x.n.sheet;
        case 'errors': return x.kind === 'name' && errOf(x.n);
        case 'noErrors': return x.kind === 'table' || !errOf(x.n);
        case 'defined': return x.kind === 'name';
        case 'tables': return x.kind === 'table';
      }
      return true;
    }).sort((a, b) => (a.kind === 'name' ? a.n.name : a.name).localeCompare(b.kind === 'name' ? b.n.name : b.name));
  };
  const render = () => {
    clear(tableWrap);
    const list = rows();
    const tbl = h('table', { class: 'ql-namemgr-table' }, h('thead', null, h('tr', null, ...['Name', 'Value', 'Refers To', 'Scope', 'Comment'].map((x) => h('th', null, x)))));
    const tb = h('tbody');
    for (const x of list) {
      const isSel = x.kind === 'name' && x.i === selectedIdx;
      const tr = h('tr', { class: isSel ? 'ql-on' : '', tabindex: '0' },
        h('td', null, icon(x.kind === 'table' ? 'table' : 'badge'), ' ', x.kind === 'name' ? x.n.name : x.name),
        h('td', null, x.kind === 'name' ? nameValueText(app, x.n) : '{…}'),
        h('td', null, '=' + (x.kind === 'name' ? x.n.ref : x.ref)),
        h('td', null, x.kind === 'name' ? nameScopeLabel(x.n) : 'Workbook'),
        h('td', null, x.kind === 'name' ? x.n.comment || '' : ''));
      const pick = () => { if (x.kind === 'name') { selectedIdx = x.i; refInput.value = '=' + x.n.ref; refInput.disabled = false; } else { selectedIdx = -1; refInput.value = '=' + x.ref; refInput.disabled = true; } render(); };
      tr.addEventListener('click', pick);
      tr.addEventListener('dblclick', () => { if (x.kind === 'name') void edit(); });
      tr.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); pick(); } });
      tb.appendChild(tr);
    }
    if (!list.length) tb.appendChild(h('tr', null, h('td', { colspan: '5', class: 'ql-help' }, book.names.length ? 'No names match the filter.' : 'This workbook has no defined names. Click New to create one.')));
    tbl.appendChild(tb);
    tableWrap.appendChild(tbl);
  };
  const edit = async () => {
    const n = book.names[selectedIdx];
    if (!n) { toast('Select a name first.'); return; }
    if (await nameDialog(app, n)) render();
  };
  const commitRef = () => {
    const n = book.names[selectedIdx];
    if (!n) return;
    const ref = refInput.value.trim().replace(/^=/, '');
    if (!ref || ref === n.ref) return;
    if (app.engine.compile(ref) instanceof Error) { toast('There is a problem with this formula.'); return; }
    if (app.isReadOnly()) { toast('This workbook is read-only.'); return; }
    book.transact('Edit Name', () => { book.recordBook(); book.names[selectedIdx].ref = ref; });
    render();
  };
  refInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); commitRef(); } if (e.key === 'Escape' && book.names[selectedIdx]) { e.stopPropagation(); refInput.value = '=' + book.names[selectedIdx].ref; } });
  const filterSel = h('select', { 'aria-label': 'Filter' }, ...[['all', 'All names'], ['sheet', 'Names scoped to worksheet'], ['book', 'Names scoped to workbook'], ['errors', 'Names with errors'], ['noErrors', 'Names without errors'], ['defined', 'Defined names'], ['tables', 'Table names']].map(([v, l]) => h('option', { value: v }, l))) as HTMLSelectElement;
  filterSel.addEventListener('change', () => { filter = filterSel.value; render(); });
  const ro = () => { if (app.isReadOnly()) { toast('This workbook is read-only.'); return true; } return false; };
  content.append(
    h('div', { class: 'ql-row ql-namemgr-bar' },
      h('button', { class: 'ql-btn', type: 'button', onclick: async () => { if (ro()) return; if (await nameDialog(app)) render(); } }, icon('add'), ' New…'),
      h('button', { class: 'ql-btn', type: 'button', onclick: () => { if (!ro()) void edit(); } }, icon('edit'), ' Edit…'),
      h('button', {
        class: 'ql-btn', type: 'button', onclick: async () => {
          if (ro()) return;
          const n = book.names[selectedIdx];
          if (!n) { toast('Select a name first.'); return; }
          if (!(await confirmDialog('Delete Name', `Delete the name "${n.name}"? Formulas that use it will show #NAME?.`, 'Delete'))) return;
          book.transact('Delete Name', () => { book.recordBook(); book.names.splice(selectedIdx, 1); });
          selectedIdx = -1; refInput.value = ''; render();
        },
      }, icon('delete'), ' Delete'),
      h('span', { class: 'ql-spacer' }), filterSel),
    tableWrap,
    h('label', { class: 'ql-label' }, 'Refers to'),
    h('div', { class: 'ql-row' }, refInput,
      h('button', { class: 'ql-iconbtn', type: 'button', title: 'Apply', 'aria-label': 'Apply', onclick: commitRef }, icon('check')),
      h('button', { class: 'ql-iconbtn', type: 'button', title: 'Cancel', 'aria-label': 'Cancel change', onclick: () => { const n = book.names[selectedIdx]; refInput.value = n ? '=' + n.ref : ''; } }, icon('close')),
      h('button', { class: 'ql-iconbtn', type: 'button', title: 'Use the current selection', 'aria-label': 'Use the current selection', onclick: () => { if (book.names[selectedIdx]) refInput.value = '=' + selectionRefText(app); } }, icon('select_all'))),
  );
  refInput.disabled = true;
  render();
  await openDialog({ title: 'Name Manager', content, wide: true, ok: 'Close', cancel: false });
}

async function pasteName(app: App): Promise<void> {
  const names = app.book.names.filter((n) => !n.hidden && (!n.sheet || n.sheet === app.sheet.name));
  const tables = app.book.sheets.flatMap((s) => s.tables.map((t) => t.name));
  const all = [...names.map((n) => n.name), ...tables].sort((a, b) => a.localeCompare(b));
  const editing = app.grid?.editing;
  const caret = editing ? app.grid.editor.input.selectionStart ?? undefined : undefined;
  if (!all.length) { toast('This workbook has no names. Use Define Name to create one.'); return; }
  let pasteList = false;
  const r = await openDialog({
    title: 'Paste Name', fields: [{ key: 'name', label: 'Paste name', type: 'list', rows: 10, options: all, value: all[0] }],
    buttons: [{ label: 'Paste List', onClick: (_v, ctl) => { pasteList = true; ctl.close({}); } }],
  });
  if (!r) { if (editing) app.grid.editor.focus(); return; }
  if (pasteList) {
    const s = app.sheet;
    const { r: r0, c: c0 } = activeCell(app);
    if (editing) app.grid.commitEdit('none');
    if (app.isReadOnly()) { toast('This workbook is read-only.'); return; }
    const list = names.slice().sort((a, b) => a.name.localeCompare(b.name));
    for (let i = 0; i < list.length; i++) if (protectedCell(app, s, r0 + i, c0) || protectedCell(app, s, r0 + i, c0 + 1)) { toast('Some of the cells are locked on a protected sheet.'); return; }
    let blocked = false;
    for (let i = 0; i < list.length; i++) if (s.get(r0 + i, c0)?.v != null || s.get(r0 + i, c0 + 1)?.v != null) blocked = true;
    if (blocked && !(await confirmDialog('Paste List', 'The list will replace existing cell contents. Continue?', 'Replace'))) return;
    app.book.transact('Paste Names', () => {
      list.forEach((n, i) => { s.patch(r0 + i, c0, { v: n.name, f: undefined }); s.patch(r0 + i, c0 + 1, { v: '=' + n.ref, f: undefined }); });
    });
    return;
  }
  const name = String(r.name || '');
  if (!name) return;
  useName(app, name, caret);
}

function useName(app: App, name: string, caret?: number): void {
  if (app.grid?.editing) {
    const inp = app.grid.editor.input;
    if (!inp.value) insertIntoEditor(app, '=' + name, { start: 0, end: 0 });
    else insertIntoEditor(app, name, null, caret);
    return;
  }
  const { r, c } = activeCell(app);
  if (protectedCell(app, app.sheet, r, c) || app.isReadOnly()) { toast('This cell cannot be edited.'); return; }
  app.grid.startEdit({ mode: 'enter', text: '=' + name });
}

async function createFromSelection(app: App): Promise<void> {
  const s = app.sheet;
  const rg = clipToUsed(s, activeRange(app));
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) { toast('Select a range that includes the labels (for example a table with headings).'); return; }
  const g = guessLabelEdges(s, rg);
  const r = await openDialog({
    title: 'Create Names from Selection', help: 'Create names from values in the:',
    fields: [
      { key: 'top', label: 'Top row', type: 'checkbox', value: !!g.top },
      { key: 'left', label: 'Left column', type: 'checkbox', value: !!g.left },
      { key: 'bottom', label: 'Bottom row', type: 'checkbox', value: false },
      { key: 'right', label: 'Right column', type: 'checkbox', value: false },
    ],
  });
  if (!r) return;
  const list = namesFromSelection(s, rg, r);
  if (!list.length) { toast('No labels were found in the chosen rows or columns.'); return; }
  const book = app.book;
  const replace: string[] = [];
  for (const n of list) if (book.names.some((x) => !x.sheet && x.name.toLowerCase() === n.name.toLowerCase())) replace.push(n.name);
  if (replace.length && !(await confirmDialog('Create Names', `Replace the existing definition of ${replace.slice(0, 5).join(', ')}${replace.length > 5 ? '…' : ''}?`, 'Replace', 'Keep'))) {
    for (const n of replace) list.splice(list.findIndex((x) => x.name === n), 1);
  }
  if (!list.length) return;
  book.transact('Create Names', () => {
    book.recordBook();
    for (const n of list) {
      const ex = book.names.find((x) => !x.sheet && x.name.toLowerCase() === n.name.toLowerCase());
      if (ex) ex.ref = n.ref; else book.names.push({ name: n.name, ref: n.ref, sheet: null });
    }
  });
  toast(`Created ${list.length} name${list.length === 1 ? '' : 's'}.`);
}

// ============================================================================================
// Trace precedents and dependents

type Arrow = App['tab']['arrows'][number] & { kind?: 'p' | 'd' | 'e' };
interface TraceState { key: string; level: number; frontier: { sheet: Sheet; r: number; c: number }[] }
const traceStates = new WeakMap<object, { p?: TraceState; d?: TraceState }>();

function trace(app: App, kind: 'p' | 'd'): void {
  const tab = app.tab;
  const sheet = app.sheet;
  const { r, c } = activeCell(app);
  const key = `${sheet.id}:${r}:${c}`;
  const st = traceStates.get(tab) || {};
  traceStates.set(tab, st);
  let state = st[kind];
  if (!state || state.key !== key || !tab.arrows.some((a) => (a as Arrow).kind === kind)) state = { key, level: 0, frontier: [{ sheet, r, c }] };
  const offSheet: string[] = [];
  const nextFrontier: { sheet: Sheet; r: number; c: number }[] = [];
  let added = 0;
  const has = (a: Arrow) => tab.arrows.some((x) => x.from.sheet === a.from.sheet && x.from.r === a.from.r && x.from.c === a.from.c && x.to.sheet === a.to.sheet && x.to.r === a.to.r && x.to.c === a.to.c);
  for (const cell of state.frontier) {
    if (kind === 'p') {
      if (cell.sheet.get(cell.r, cell.c)?.f == null) continue;
      for (const p of app.engine.precedentsOf(cell.sheet, cell.r, cell.c)) {
        const a: Arrow = { from: { sheet: p.sheet, r: p.r1, c: p.c1 }, to: { sheet: cell.sheet, r: cell.r, c: cell.c }, kind: 'p' };
        if (p.sheet !== cell.sheet) { offSheet.push(`${quoteSheet(p.sheet.name)}!${rangeName(p)}`); continue; }
        if (!has(a)) { tab.arrows.push(a); added++; }
        for (let rr = p.r1; rr <= Math.min(p.r2, p.r1 + 200); rr++) for (let cc = p.c1; cc <= Math.min(p.c2, p.c1 + 50); cc++) if (p.sheet.get(rr, cc)?.f != null) nextFrontier.push({ sheet: p.sheet, r: rr, c: cc });
      }
    } else {
      for (const d of app.engine.dependentsOfCell(cell.sheet, cell.r, cell.c)) {
        const a: Arrow = { from: { sheet: cell.sheet, r: cell.r, c: cell.c }, to: { sheet: d.sheet, r: d.r, c: d.c }, kind: 'd' };
        if (d.sheet !== cell.sheet) { offSheet.push(`${quoteSheet(d.sheet.name)}!${cellName(d.r, d.c)}`); continue; }
        if (!has(a)) { tab.arrows.push(a); added++; }
        nextFrontier.push(d);
      }
    }
  }
  state.level++;
  state.frontier = nextFrontier;
  st[kind] = state;
  app.grid.invalidate();
  app.refresh();
  const what = kind === 'p' ? 'precedents' : 'dependents';
  if (offSheet.length) {
    const uniq = [...new Set(offSheet)];
    toast(`${kind === 'p' ? 'Precedents' : 'Dependents'} on other sheets: ${uniq.slice(0, 4).join(', ')}${uniq.length > 4 ? '…' : ''}`, { action: 'Go To', onAction: () => app.exec('nav.goto', uniq[0]), ms: 6000 });
  } else if (!added) {
    toast(state.level <= 1 ? (kind === 'p' ? 'The active cell has no formula with references to trace.' : 'No formulas refer to the active cell.') : `No more ${what} to trace.`);
  }
}

function removeArrows(app: App, which: 'all' | 'p' | 'd' = 'all'): void {
  const tab = app.tab;
  tab.arrows = which === 'all' ? [] : tab.arrows.filter((a) => (a as Arrow).kind !== which && !((a as Arrow).kind === 'e' && which === 'p'));
  const st = traceStates.get(tab);
  if (st) { if (which !== 'd') delete st.p; if (which !== 'p') delete st.d; }
  app.grid.invalidate();
  app.refresh();
}

function traceError(app: App): void {
  const sheet = app.sheet;
  const { r, c } = activeCell(app);
  const v = sheet.value(r, c);
  if (!(v instanceof CellError)) { toast('The active cell does not contain an error.'); return; }
  const tab = app.tab;
  const seen = new Set<string>();
  let added = 0;
  const visit = (sh: Sheet, rr: number, cc: number, depth: number) => {
    if (depth > 30 || seen.has(`${sh.id}:${rr}:${cc}`)) return;
    seen.add(`${sh.id}:${rr}:${cc}`);
    for (const p of app.engine.precedentsOf(sh, rr, cc)) {
      for (let r2 = p.r1; r2 <= Math.min(p.r2, p.r1 + 500); r2++) for (let c2 = p.c1; c2 <= Math.min(p.c2, p.c1 + 50); c2++) {
        if (p.sheet.value(r2, c2) instanceof CellError) {
          tab.arrows.push({ from: { sheet: p.sheet, r: r2, c: c2 }, to: { sheet: sh, r: rr, c: cc }, kind: 'e' } as Arrow);
          added++;
          visit(p.sheet, r2, c2, depth + 1);
        }
      }
    }
  };
  visit(sheet, r, c, 0);
  app.grid.invalidate();
  app.refresh();
  if (!added) toast(`The error starts in this cell: ${errorDescription(v)}`);
}

// ============================================================================================
// Error checking

export const ERROR_HELP: Record<string, string> = {
  '#DIV/0!': 'Divide by Zero Error: the formula divides by zero or by an empty cell.',
  '#VALUE!': 'Value Error: a value used in the formula is of the wrong type (for example text where a number is expected).',
  '#REF!': 'Invalid Cell Reference Error: the formula refers to a cell that was deleted or is not valid.',
  '#NAME?': 'Invalid Name Error: the formula contains unrecognized text, such as a misspelled function or an undefined name.',
  '#NUM!': 'Invalid Number Error: the formula produces a number that is too large or too small, or a function cannot find a result.',
  '#N/A': 'Value Not Available Error: a lookup did not find the value, or a required value is not available.',
  '#NULL!': 'Null Error: the ranges in the formula do not intersect.',
  '#SPILL!': 'Spill Error: the formula returns several values but the cells it needs are not empty.',
  '#CALC!': 'Calculation Error: the calculation could not complete (for example an empty array).',
  '#CIRC!': 'Circular Reference: the formula refers to its own cell directly or indirectly.',
};
export function errorDescription(e: CellError): string { return ERROR_HELP[e.code] || `${e.code} error.`; }

const ignoredErrors = new WeakMap<Sheet, Set<number>>();

export interface FoundIssue { r: number; c: number; kind: 'error' | 'inconsistent' | 'textNumber'; message: string }

/** Scans a sheet for error values, formulas inconsistent with their neighbors and numbers stored as text. */
export function findIssues(sheet: Sheet, ignored?: Set<number>): FoundIssue[] {
  const out: FoundIssue[] = [];
  const keys = [...sheet.cells.keys()].sort((a, b) => keyRow(a) - keyRow(b) || keyCol(a) - keyCol(b));
  const shape = (r: number, c: number): string | null => {
    const f = sheet.get(r, c)?.f;
    if (f == null) return null;
    try {
      // Relative (R1C1-style) shape: references as offsets from the cell.
      return tokenize(f).map((t) => t.type === 'ref' && t.ref && !t.ref.invalid ? `R[${t.ref.ar1 ? 'a' + t.ref.r1 : t.ref.r1 - r}]C[${t.ref.ac1 ? 'a' + t.ref.c1 : t.ref.c1 - c}]:${t.ref.kind === 'range' ? `R[${t.ref.ar2 ? 'a' + t.ref.r2 : t.ref.r2 - r}]C[${t.ref.ac2 ? 'a' + t.ref.c2 : t.ref.c2 - c}]` : ''}` : t.text.toUpperCase()).join('');
    } catch { return f; }
  };
  for (const k of keys) {
    if (ignored?.has(k)) continue;
    const r = keyRow(k), c = keyCol(k);
    const cell = sheet.cells.get(k)!;
    if (cell.sp != null && cell.f == null) continue;
    if (cell.v instanceof CellError) { out.push({ r, c, kind: 'error', message: errorDescription(cell.v) }); continue; }
    if (cell.f != null) {
      const me = shape(r, c);
      const up = shape(r - 1, c), down = shape(r + 1, c), left = shape(r, c - 1), right = shape(r, c + 1);
      if ((up && down && up === down && me !== up) || (left && right && left === right && me !== left)) out.push({ r, c, kind: 'inconsistent', message: 'Inconsistent Formula: the formula in this cell differs from the formulas in this area of the sheet.' });
      continue;
    }
    if (typeof cell.v === 'string' && /^\s*[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?\s*$/.test(cell.v) && sheet.styleOf(r, c).numFmt !== '@') out.push({ r, c, kind: 'textNumber', message: 'Number Stored as Text: the number in this cell is formatted as text or preceded by an apostrophe.' });
  }
  return out;
}

async function errorChecking(app: App): Promise<void> {
  const sheet = app.sheet;
  const ign = ignoredErrors.get(sheet) || new Set<number>();
  ignoredErrors.set(sheet, ign);
  let issues = findIssues(sheet, ign);
  if (!issues.length) { await alertDialog('Error Checking', 'The error check is complete for the entire sheet.'); return; }
  const { r: ar, c: ac } = activeCell(app);
  let i = Math.max(0, issues.findIndex((x) => x.r > ar || (x.r === ar && x.c >= ac)));
  const content = h('div', { class: 'ql-errcheck' });
  let ctl: DialogControl | null = null;
  const render = () => {
    clear(content);
    if (!issues.length) { content.appendChild(h('p', null, 'The error check is complete for the entire sheet.')); return; }
    const it = issues[i];
    app.grid.select(it.r, it.c);
    const cell = sheet.get(it.r, it.c);
    content.append(
      h('div', { class: 'ql-label' }, `${it.kind === 'error' ? 'Error' : 'Issue'} in cell ${cellName(it.r, it.c)} (${i + 1} of ${issues.length})`),
      cell?.f != null ? h('pre', { class: 'ql-fx-syntax' }, '=' + cell.f) : h('pre', { class: 'ql-fx-syntax' }, String(cell?.v ?? '')),
      h('p', null, it.message),
      h('div', { class: 'ql-row ql-wrap' },
        h('button', { class: 'ql-btn', type: 'button', onclick: () => app.exec('formulas.functionHelp') }, 'Help on this Error'),
        it.kind === 'textNumber'
          ? h('button', { class: 'ql-btn', type: 'button', onclick: () => { if (!guardWrite(app, sheet, it.r, it.c)) return; app.book.transact('Convert to Number', () => sheet.patch(it.r, it.c, { v: Number(String(sheet.value(it.r, it.c)).trim()) })); issues.splice(i, 1); i = Math.min(i, issues.length - 1); render(); } }, 'Convert to Number')
          : h('button', { class: 'ql-btn', type: 'button', onclick: () => { ctl?.close(null); void evaluateFormula(app); } }, 'Show Calculation Steps…'),
        it.kind === 'inconsistent' ? h('button', {
          class: 'ql-btn', type: 'button', onclick: () => {
            const above = sheet.get(it.r - 1, it.c)?.f;
            if (above == null || !guardWrite(app, sheet, it.r, it.c)) return;
            import('../calc/parser').then(({ shiftFormula }) => { app.book.transact('Copy Formula from Above', () => sheet.patch(it.r, it.c, { f: shiftFormula(above, 1, 0) })); issues.splice(i, 1); i = Math.min(i, issues.length - 1); render(); });
          },
        }, 'Copy Formula from Above') : null,
        h('button', { class: 'ql-btn', type: 'button', onclick: () => { ign.add(cellKey(it.r, it.c)); issues.splice(i, 1); i = Math.min(i, Math.max(0, issues.length - 1)); render(); } }, 'Ignore Error'),
        h('button', { class: 'ql-btn', type: 'button', onclick: () => { ctl?.close(null); app.grid.startEdit({ mode: 'edit', fromBar: !!app.prefs.showFormulaBar }); } }, 'Edit in Formula Bar')),
      h('div', { class: 'ql-row' },
        h('button', { class: 'ql-btn', type: 'button', disabled: issues.length < 2, onclick: () => { i = (i - 1 + issues.length) % issues.length; render(); } }, icon('chevron_left'), ' Previous'),
        h('button', { class: 'ql-btn', type: 'button', disabled: issues.length < 2, onclick: () => { i = (i + 1) % issues.length; render(); } }, 'Next ', icon('chevron_right')),
        h('span', { class: 'ql-spacer' }),
        h('button', { class: 'ql-link', type: 'button', onclick: () => { ign.clear(); issues = findIssues(sheet, ign); i = 0; render(); } }, 'Reset Ignored Errors')),
    );
  };
  await openDialog({ title: 'Error Checking', content, wide: true, ok: false, cancel: 'Close', onOpen: (c) => { ctl = c; render(); } });
}

function circularCells(app: App): { sheet: Sheet; r: number; c: number }[] {
  const out: { sheet: Sheet; r: number; c: number }[] = [];
  const seen = new Set<string>();
  for (const fk of app.engine.circular) {
    if (seen.has(fk)) continue;
    seen.add(fk);
    const i = fk.lastIndexOf(':');
    const sh = app.book.sheetById(fk.slice(0, i));
    const k = Number(fk.slice(i + 1));
    if (sh) out.push({ sheet: sh, r: keyRow(k), c: keyCol(k) });
  }
  return out;
}

function goToCell(app: App, sheet: Sheet, r: number, c: number): void {
  const idx = app.book.sheets.indexOf(sheet);
  if (idx >= 0 && idx !== app.book.activeSheet) app.exec('sheet.activate', idx);
  app.grid.select(r, c);
  app.grid.scrollIntoView(r, c);
}

// ============================================================================================
// Evaluate Formula dialog

async function evaluateFormula(app: App): Promise<void> {
  const sheet = app.sheet;
  const { r, c } = activeCell(app);
  const f = sheet.get(r, c)?.f;
  if (f == null) { toast('Select a cell that contains a formula.'); return; }
  const stack: FormulaStepper[] = [new FormulaStepper(app.engine, sheet, r, c)];
  const content = h('div', { class: 'ql-evalf' });
  const btn = (label: string, fn: () => void, primary = false) => h('button', { class: 'ql-btn' + (primary ? ' ql-primary' : ''), type: 'button', onclick: fn }, label);
  const render = () => {
    clear(content);
    const top = stack[stack.length - 1];
    content.appendChild(h('p', { class: 'ql-help' }, 'Click Evaluate to calculate the underlined part of the formula.'));
    stack.forEach((st, depth) => {
      const label = `${quoteSheet(st.sheet.name)}!${cellName(st.r, st.c)}`;
      const box = h('div', { class: 'ql-evalf-box' + (st === top ? ' ql-on' : '') });
      const p = st.next;
      const t = st.text;
      if (st.error) box.append(h('span', { class: 'ql-warn' }, st.error));
      else if (p && st === top) box.append('=', t.slice(0, p.s), h('u', { class: 'ql-evalf-next' }, t.slice(p.s, p.e)), t.slice(p.e));
      else if (!p) box.append(h('i', null, t ? (st.steps ? fmtFinal(st) : '=' + t) : ''));
      else box.append('=' + t);
      content.append(h('div', { class: 'ql-evalf-row', style: `margin-left:${depth * 14}px` }, h('div', { class: 'ql-evalf-ref' }, label), box));
    });
    const p = top.next;
    const canStepIn = !!p?.stepIn && stack.length < 12;
    const canStepOut = stack.length > 1 && top.done;
    content.append(h('div', { class: 'ql-row ql-wrap ql-evalf-btns' },
      top.done && stack.length === 1 ? btn('Restart', () => { stack.splice(0, stack.length, new FormulaStepper(app.engine, sheet, r, c)); render(); }, true) : btn('Evaluate', () => { if (top.done && stack.length > 1) { stack.pop(); stack[stack.length - 1].step(); } else top.step(); render(); }, true),
      h('button', { class: 'ql-btn', type: 'button', disabled: !canStepIn, onclick: () => { const si = p!.stepIn!; stack.push(new FormulaStepper(app.engine, si.sheet, si.r, si.c)); render(); } }, 'Step In'),
      h('button', { class: 'ql-btn', type: 'button', disabled: !(stack.length > 1), onclick: () => { stack.pop(); stack[stack.length - 1].step(); render(); } }, 'Step Out'),
    ));
    if (canStepOut) content.append(h('p', { class: 'ql-help' }, 'The referenced cell has been evaluated. Click Step Out to return to the previous formula.'));
    if (p?.stepIn) content.append(h('p', { class: 'ql-help' }, `The next part refers to ${cellName(p.stepIn.r, p.stepIn.c)}, which contains a formula. Click Step In to evaluate it.`));
  };
  const fmtFinal = (st: FormulaStepper) => {
    const v = app.engine.evaluate(st.text, st.sheet, st.r, st.c);
    const sc = isRef(v) ? v.sheet.value(v.r1, v.c1) : isMatrix(v) ? null : v as Scalar;
    return isMatrix(v) ? valuePreview(v) : formatValue(sc, st.sheet.styleOf(st.r, st.c).numFmt, { date1904: st.sheet.book.calc.date1904 }).text || literalText(sc);
  };
  render();
  await openDialog({ title: 'Evaluate Formula', content, wide: true, ok: false, cancel: 'Close' });
}

// ============================================================================================
// Calculation options

function setCalcMode(app: App, mode: 'auto' | 'manual' | 'autoNoTables'): void {
  const book = app.book;
  if (book.calc.mode === mode) return;
  book.transact('Calculation Options', () => { book.recordBook(); book.calc.mode = mode; });
  if (mode !== 'manual') app.calculate();
  app.ui.status?.(mode === 'manual' ? 'Calculation: Manual (press F9 to calculate)' : 'Calculation: Automatic');
}

async function iterationSettings(app: App): Promise<void> {
  const calc = app.book.calc;
  const r = await openDialog({
    title: 'Iterative Calculation',
    help: 'Iterative calculation lets formulas with circular references converge by recalculating repeatedly.',
    fields: [
      { key: 'iterative', label: 'Enable iterative calculation', type: 'checkbox', value: calc.iterative },
      { key: 'maxIterations', label: 'Maximum iterations', type: 'number', min: 1, max: 32767, step: 1, value: calc.maxIterations, half: true },
      { key: 'maxChange', label: 'Maximum change', type: 'number', min: 0, step: 0.0001, value: calc.maxChange, half: true },
      { key: 'precisionAsDisplayed', label: 'Set precision as displayed', type: 'checkbox', value: calc.precisionAsDisplayed },
      { key: 'recalcBeforeSave', label: 'Recalculate workbook before saving', type: 'checkbox', value: calc.recalcBeforeSave },
    ],
    validate: (v) => (v.maxIterations == null || v.maxIterations < 1 || v.maxIterations > 32767 ? 'Maximum iterations must be between 1 and 32767.' : v.maxChange == null || v.maxChange < 0 ? 'Maximum change must be zero or more.' : null),
  });
  if (!r) return;
  const book = app.book;
  book.transact('Calculation Options', () => {
    book.recordBook();
    Object.assign(book.calc, { iterative: !!r.iterative, maxIterations: Math.round(r.maxIterations), maxChange: Number(r.maxChange), precisionAsDisplayed: !!r.precisionAsDisplayed, recalcBeforeSave: !!r.recalcBeforeSave });
  });
  app.calculate(true);
}

// ============================================================================================
// Command definitions

const catMenu = (cat: FnCategory): MenuEntry[] => {
  const list = fnsIn(cat).map((f) => ({ id: 'formulas.insertFunction', arg: { name: f.name }, label: f.name } as MenuEntry));
  return [...list, { separator: true }, { id: 'formulas.insertFunction', label: 'Insert Function…', icon: 'function' }];
};

const catCommand = (id: string, label: string, ic: string, cat: FnCategory): any => ({
  id, label, icon: ic, tab: TAB, group: 'Function Library', kind: 'menu', edits: false,
  keywords: `${cat} functions`, tip: `Insert a ${cat} function`,
  run: (app: App, arg?: any) => (arg?.name ? app.exec('formulas.insertFunction', arg) : openCommandMenu(app, id)),
  menu: () => catMenu(cat),
});

def([
  {
    id: 'formulas.insertFunction', label: 'Insert Function', icon: 'function', tab: TAB, group: 'Function Library', keys: ['Shift-F3'], edits: false, popular: true,
    tip: 'Search for a function, then enter its arguments with live results', keywords: 'fx wizard arguments',
    run: (app, arg) => insertFunction(app, arg || {}),
  },
  {
    id: 'formulas.functionHelp', label: 'Function Help', icon: 'help', tab: TAB, group: 'Function Library', edits: false, noRecord: true,
    tip: 'Show the syntax and description of the function at the cursor', keywords: 'syntax describe',
    run: (app, arg) => functionHelp(app, arg),
  },
  {
    id: 'formulas.functionBrowser', label: 'Function Reference', icon: 'menu_book', tab: TAB, group: 'Function Library', edits: false, noRecord: true,
    tip: 'Browse every function by category', run: (app) => insertFunction(app, { browse: true }),
  },
  {
    id: 'formulas.recent', label: 'Recently Used', icon: 'history', tab: TAB, group: 'Function Library', kind: 'menu', edits: false,
    run: (app, arg) => (arg?.name ? app.exec('formulas.insertFunction', arg) : openCommandMenu(app, 'formulas.recent')),
    menu: () => [...recentFunctions().map((n) => ({ id: 'formulas.insertFunction', arg: { name: n }, label: n } as MenuEntry)), { separator: true }, { id: 'formulas.insertFunction', label: 'Insert Function…', icon: 'function' }],
  },
  catCommand('formulas.financial', 'Financial', 'account_balance', 'Financial'),
  catCommand('formulas.logical', 'Logical', 'rule', 'Logical'),
  catCommand('formulas.text', 'Text', 'text_fields', 'Text'),
  catCommand('formulas.dateTime', 'Date & Time', 'calendar_month', 'Date & Time'),
  catCommand('formulas.lookup', 'Lookup & Reference', 'manage_search', 'Lookup & Reference'),
  catCommand('formulas.mathTrig', 'Math & Trig', 'calculate', 'Math & Trig'),
  {
    id: 'formulas.more', label: 'More Functions', icon: 'more_horiz', tab: TAB, group: 'Function Library', kind: 'menu', edits: false,
    run: (app, arg) => (arg?.name ? app.exec('formulas.insertFunction', arg) : openCommandMenu(app, 'formulas.more')),
    menu: () => (['Statistical', 'Engineering', 'Cube', 'Information', 'Compatibility', 'Web', 'Database', 'Dynamic Array'] as FnCategory[])
      .filter((k) => fnsIn(k).length)
      .map((k) => ({ label: k, icon: ({ Statistical: 'query_stats', Engineering: 'engineering', Cube: 'view_in_ar', Information: 'info', Compatibility: 'history_edu', Web: 'public', Database: 'database', 'Dynamic Array': 'data_array' } as Record<string, string>)[k], submenu: catMenu(k) })),
  },
  // Defined names
  {
    id: 'formulas.nameManager', label: 'Name Manager', icon: 'badge', tab: TAB, group: 'Defined Names', keys: ['Mod-F3'], edits: false, noRecord: true,
    tip: 'Create, edit, delete and find all the names used in the workbook', run: (app) => nameManager(app),
  },
  {
    id: 'formulas.defineName', label: 'Define Name', icon: 'new_label', tab: TAB, group: 'Defined Names', edits: true,
    tip: 'Name cells so formulas can refer to them by name', keywords: 'named range',
    run: (app, arg) => (arg && arg.name && arg.ref ? defineNameDirect(app, arg) : nameDialog(app)),
  },
  {
    id: 'formulas.useInFormula', label: 'Use in Formula', icon: 'input', tab: TAB, group: 'Defined Names', kind: 'menu', edits: false,
    tip: 'Insert a name into the formula you are editing',
    run: (app, name) => (typeof name === 'string' && name ? useName(app, name) : openCommandMenu(app, 'formulas.useInFormula')),
    menu: (app) => {
      const names = app.book.names.filter((n) => !n.hidden && (!n.sheet || n.sheet === app.sheet.name)).map((n) => n.name).sort((a, b) => a.localeCompare(b));
      return [...(names.length ? names.map((n) => ({ id: 'formulas.useInFormula', arg: n, label: n, icon: 'badge' })) : [{ label: '(No names defined)' }]), { separator: true }, { id: 'formulas.pasteName', label: 'Paste Names…', icon: 'content_paste' }];
    },
  },
  {
    id: 'formulas.pasteName', label: 'Paste Name', icon: 'content_paste_go', tab: TAB, group: 'Defined Names', keys: ['F3'], edits: false,
    tip: 'Paste a name into a formula, or paste a list of all names', run: (app) => pasteName(app),
  },
  {
    id: 'formulas.createFromSelection', label: 'Create from Selection', icon: 'playlist_add', tab: TAB, group: 'Defined Names', keys: ['Mod-Shift-F3'], edits: true,
    tip: 'Name cells automatically from the labels in the selection', run: (app) => createFromSelection(app),
  },
  // Auditing
  {
    id: 'formulas.tracePrecedents', label: 'Trace Precedents', icon: 'start', tab: TAB, group: 'Formula Auditing', noRecord: true, edits: false,
    tip: 'Show arrows from the cells that affect the active cell. Click again to go one level further.', run: (app) => trace(app, 'p'),
  },
  {
    id: 'formulas.traceDependents', label: 'Trace Dependents', icon: 'output', tab: TAB, group: 'Formula Auditing', noRecord: true, edits: false,
    tip: 'Show arrows to the cells affected by the active cell. Click again to go one level further.', run: (app) => trace(app, 'd'),
  },
  {
    id: 'formulas.removeArrows', label: 'Remove Arrows', icon: 'layers_clear', tab: TAB, group: 'Formula Auditing', kind: 'split', noRecord: true, edits: false,
    run: (app, which) => removeArrows(app, which === 'p' || which === 'd' ? which : 'all'),
    menu: () => [{ id: 'formulas.removeArrows', arg: 'all', label: 'Remove Arrows', icon: 'layers_clear' }, { id: 'formulas.removeArrows', arg: 'p', label: 'Remove Precedent Arrows', icon: 'start' }, { id: 'formulas.removeArrows', arg: 'd', label: 'Remove Dependent Arrows', icon: 'output' }],
  },
  {
    id: 'formulas.errorChecking', label: 'Error Checking', icon: 'error_med', tab: TAB, group: 'Formula Auditing', kind: 'split', edits: false, noRecord: true,
    tip: 'Check the sheet for common formula errors',
    run: (app, arg) => (arg === 'trace' ? traceError(app) : errorChecking(app)),
    menu: (app) => {
      const circ = circularCells(app);
      return [
        { id: 'formulas.errorChecking', label: 'Error Checking…', icon: 'error_med' },
        { id: 'formulas.traceError', label: 'Trace Error', icon: 'troubleshoot' },
        { label: 'Circular References', icon: 'sync_problem', submenu: circ.length ? circ.slice(0, 30).map((x) => ({ id: 'formulas.circularRefs', arg: { sheet: x.sheet.name, r: x.r, c: x.c }, label: `${quoteSheet(x.sheet.name)}!${cellName(x.r, x.c)}` })) : [{ label: 'No circular references' }] },
      ];
    },
  },
  { id: 'formulas.traceError', label: 'Trace Error', icon: 'troubleshoot', tab: TAB, group: 'Formula Auditing', noRecord: true, edits: false, tip: 'Show arrows to the cells that cause the error in the active cell', run: (app) => traceError(app) },
  {
    id: 'formulas.circularRefs', label: 'Circular References', icon: 'sync_problem', tab: TAB, group: 'Formula Auditing', kind: 'menu', noRecord: true, edits: false,
    tip: 'List the cells that are part of a circular reference',
    run: (app, arg) => {
      if (arg && arg.sheet != null) { const sh = app.book.sheetByName(arg.sheet); if (sh) goToCell(app, sh, arg.r, arg.c); return; }
      if (!circularCells(app).length) { toast('There are no circular references.'); return; }
      openCommandMenu(app, 'formulas.circularRefs');
    },
    menu: (app) => { const circ = circularCells(app); return circ.length ? circ.slice(0, 50).map((x) => ({ id: 'formulas.circularRefs', arg: { sheet: x.sheet.name, r: x.r, c: x.c }, label: `${quoteSheet(x.sheet.name)}!${cellName(x.r, x.c)}`, icon: 'sync_problem' })) : [{ label: 'No circular references' }]; },
  },
  { id: 'formulas.evaluate', label: 'Evaluate Formula', icon: 'calculate', tab: TAB, group: 'Formula Auditing', noRecord: true, edits: false, tip: 'Step through a formula to see how its result is calculated', run: (app) => evaluateFormula(app) },
  { id: 'formulas.watchWindow', label: 'Watch Window', icon: 'visibility', tab: TAB, group: 'Formula Auditing', noRecord: true, edits: false, tip: 'Watch the values of cells while you work elsewhere', run: (app) => app.ui.togglePane('watch'), isActive: (app) => !!app.ui?.paneOpen?.('watch') },
  {
    id: 'formulas.addWatch', label: 'Add Watch', icon: 'add_task', tab: TAB, group: 'Formula Auditing', noRecord: true, edits: false,
    tip: 'Add the selected cells to the Watch Window',
    run: (app) => {
      const s = app.sheet;
      const tab = app.tab;
      let n = 0;
      for (const rg0 of s.sel.ranges) {
        const rg = clipToUsed(s, rg0);
        for (let r = rg.r1; r <= rg.r2 && n < 200; r++) for (let c = rg.c1; c <= rg.c2 && n < 200; c++) {
          if (rg.r1 !== rg.r2 || rg.c1 !== rg.c2) { if (s.get(r, c)?.f == null) continue; }
          if (!tab.watches.some((w) => w.sheet === s.name && w.r === r && w.c === c)) { tab.watches.push({ sheet: s.name, r, c }); n++; }
        }
      }
      if (!app.ui.paneOpen('watch')) app.ui.openPane('watch');
      app.refresh();
      toast(n ? `Added ${n} watch${n === 1 ? '' : 'es'}.` : 'Those cells are already watched (or contain no formulas).');
    },
  },
  // Calculation
  {
    id: 'formulas.calcOptions', label: 'Calculation Options', icon: 'settings_suggest', tab: TAB, group: 'Calculation', kind: 'menu', edits: true,
    tip: 'Choose when formulas are recalculated',
    run: (app, arg) => {
      if (arg === 'auto' || arg === 'manual' || arg === 'autoNoTables') return setCalcMode(app, arg);
      if (arg === 'iteration') return iterationSettings(app);
      openCommandMenu(app, 'formulas.calcOptions');
    },
    menu: (app) => {
      const m = app.book.calc.mode;
      return [
        { id: 'formulas.calcOptions', arg: 'auto', label: 'Automatic', checked: m === 'auto' },
        { id: 'formulas.calcOptions', arg: 'autoNoTables', label: 'Automatic Except for Data Tables', checked: m === 'autoNoTables' },
        { id: 'formulas.calcOptions', arg: 'manual', label: 'Manual', checked: m === 'manual' },
        { separator: true },
        { id: 'formulas.calcOptions', arg: 'iteration', label: 'Iterative Calculation…', icon: 'loop', checked: app.book.calc.iterative },
      ];
    },
  },
  { id: 'formulas.calcNow', label: 'Calculate Now', icon: 'calculate', tab: TAB, group: 'Calculation', keys: ['F9'], noRecord: false, edits: false, tip: 'Calculate the entire workbook now', run: (app) => { app.calculate(); app.ui.status?.('Calculated'); } },
  { id: 'formulas.calcSheet', label: 'Calculate Sheet', icon: 'functions', tab: TAB, group: 'Calculation', keys: ['Shift-F9'], edits: false, tip: 'Calculate the active sheet now', run: (app) => { app.calculate(); app.ui.status?.('Calculated'); } },
  { id: 'formulas.calcFull', label: 'Calculate Full', icon: 'restart_alt', tab: TAB, group: 'Calculation', keys: ['Mod-Alt-F9'], edits: false, tip: 'Rebuild dependencies and recalculate every formula in all open workbooks', run: (app) => { for (const d of app.docs) { if (d !== app.tab) { d.book.pending.structural = true; d.engine.recalc(true); d.cf.invalidate(); } } app.calculate(true); app.ui.status?.('Full calculation complete'); } },
]);

function defineNameDirect(app: App, arg: { name: string; ref: string; sheet?: string | null; comment?: string }): void {
  const book = app.book;
  const err = validateDefinedName(book, arg.name, arg.sheet || null);
  if (err) { toast(err); return; }
  book.transact('Define Name', () => { book.recordBook(); book.names.push({ name: arg.name, ref: arg.ref.replace(/^=/, ''), sheet: arg.sheet || null, ...(arg.comment ? { comment: arg.comment } : {}) }); });
}

void resolveRangeText; void formatGeneral;
