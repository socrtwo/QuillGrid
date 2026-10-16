/**
 * Formula tokenizer and parser.
 *
 * Grammar (spreadsheet precedence, highest first):
 *   reference operators  ':'  (range), ' ' (intersection), ',' inside parentheses (union)
 *   prefix  -  +  @
 *   postfix %  #  (percent, spill range)
 *   ^
 *   *  /
 *   +  -
 *   &
 *   =  <>  <  >  <=  >=
 *
 * The tokenizer keeps source offsets so references can be rewritten in place (fill, copy, insert/delete).
 */
import { colIndex, MAX_COLS, MAX_ROWS } from '../model/address';
import { errorFromCode, type ErrorCode } from '../model/types';

export type TokType = 'num' | 'str' | 'bool' | 'err' | 'ref' | 'name' | 'func' | 'op' | 'lparen' | 'rparen' | 'lbrace' | 'rbrace' | 'comma' | 'semi' | 'ws' | 'sref' | 'eof';

export interface Token { type: TokType; text: string; pos: number; end: number; ref?: RefToken }

/** Parsed reference token (A1, $A$1:B2, A:A, 3:5, Sheet1!A1, 'My Sheet'!A1:B2, Sheet1:Sheet3!A1). */
export interface RefToken {
  sheet?: string; sheet2?: string;
  r1: number; c1: number; r2: number; c2: number;
  ar1: boolean; ac1: boolean; ar2: boolean; ac2: boolean;
  kind: 'cell' | 'range' | 'cols' | 'rows';
  /** #REF! in place of a deleted reference. */
  invalid?: boolean;
}

export type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: ErrorCode }
  | { t: 'arr'; rows: Node[][] }
  | { t: 'ref'; ref: RefToken }
  | { t: 'name'; name: string; sheet?: string }
  | { t: 'sref'; table: string; spec: string }
  | { t: 'fn'; name: string; args: Node[] }
  | { t: 'call'; callee: Node; args: Node[] }
  | { t: 'un'; op: string; a: Node }
  | { t: 'bin'; op: string; a: Node; b: Node }
  | { t: 'pct'; a: Node }
  | { t: 'spill'; a: Node }
  | { t: 'isect'; a: Node }
  | { t: 'union'; items: Node[] }
  | { t: 'missing' };

export class FormulaError extends Error {
  constructor(msg: string, public pos = 0) { super(msg); }
}

const SHEET_PREFIX = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_À-￿][\w.À-￿]*)(?::(?:'(?:[^']|'')+'|[A-Za-z_À-￿][\w.À-￿]*))?!`;
const CELL = String.raw`\$?[A-Za-z]{1,3}\$?\d{1,7}`;
const RE_RANGE = new RegExp(String.raw`^(${SHEET_PREFIX})?(${CELL}(?::${CELL})?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})(?![\w(.\[])`);
const RE_REFERR = new RegExp(String.raw`^(${SHEET_PREFIX})?#REF!`, 'i');
const RE_NUM = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/;
const RE_ERR = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA|CIRC!)/i;
const RE_IDENT = /^[A-Za-z_\\À-￿][\w.À-￿]*/;

function unquoteSheet(s: string): string {
  return s.startsWith("'") ? s.slice(1, -1).replace(/''/g, "'") : s;
}

export function parseRefText(text: string): RefToken | null {
  const m = RE_RANGE.exec(text);
  if (!m || m[0].length !== text.length) return null;
  return refFromMatch(m);
}

function refFromMatch(m: RegExpExecArray): RefToken | null {
  let sheet: string | undefined, sheet2: string | undefined;
  if (m[1]) {
    const pre = m[1].slice(0, -1);
    // split on ':' outside quotes
    let split = -1, inQ = false;
    for (let i = 0; i < pre.length; i++) { if (pre[i] === "'") inQ = !inQ; else if (pre[i] === ':' && !inQ) { split = i; break; } }
    if (split >= 0) { sheet = unquoteSheet(pre.slice(0, split)); sheet2 = unquoteSheet(pre.slice(split + 1)); }
    else sheet = unquoteSheet(pre);
  }
  const body = m[2];
  const [p, q] = body.split(':');
  const cellRe = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/;
  if (q === undefined) {
    const a = cellRe.exec(p);
    if (!a) return null;
    const c = colIndex(a[2]), r = +a[4] - 1;
    if (c < 0 || c >= MAX_COLS || r < 0 || r >= MAX_ROWS) return null;
    return { sheet, sheet2, r1: r, c1: c, r2: r, c2: c, ac1: !!a[1], ar1: !!a[3], ac2: !!a[1], ar2: !!a[3], kind: 'cell' };
  }
  const a = cellRe.exec(p), b = cellRe.exec(q);
  if (a && b) {
    const c1 = colIndex(a[2]), r1 = +a[4] - 1, c2 = colIndex(b[2]), r2 = +b[4] - 1;
    if ([c1, c2].some((c) => c < 0 || c >= MAX_COLS) || [r1, r2].some((r) => r < 0 || r >= MAX_ROWS)) return null;
    return { sheet, sheet2, r1, c1, r2, c2, ac1: !!a[1], ar1: !!a[3], ac2: !!b[1], ar2: !!b[3], kind: 'range' };
  }
  const ca = /^(\$?)([A-Za-z]{1,3})$/.exec(p), cb = /^(\$?)([A-Za-z]{1,3})$/.exec(q);
  if (ca && cb) {
    const c1 = colIndex(ca[2]), c2 = colIndex(cb[2]);
    if (c1 < 0 || c2 < 0 || c1 >= MAX_COLS || c2 >= MAX_COLS) return null;
    return { sheet, sheet2, r1: 0, c1, r2: MAX_ROWS - 1, c2, ac1: !!ca[1], ar1: true, ac2: !!cb[1], ar2: true, kind: 'cols' };
  }
  const ra = /^(\$?)(\d{1,7})$/.exec(p), rb = /^(\$?)(\d{1,7})$/.exec(q);
  if (ra && rb) {
    const r1 = +ra[2] - 1, r2 = +rb[2] - 1;
    if (r1 < 0 || r2 < 0 || r1 >= MAX_ROWS || r2 >= MAX_ROWS) return null;
    return { sheet, sheet2, r1, c1: 0, r2, c2: MAX_COLS - 1, ar1: !!ra[1], ac1: true, ar2: !!rb[1], ac2: true, kind: 'rows' };
  }
  return null;
}

export function tokenize(src: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  const push = (type: TokType, text: string, extra: Partial<Token> = {}) => { out.push({ type, text, pos: i, end: i + text.length, ...extra }); i += text.length; };
  while (i < src.length) {
    const ch = src[i];
    const rest = src.slice(i);
    if (/\s/.test(ch)) { const m = /^\s+/.exec(rest)!; push('ws', m[0]); continue; }
    if (ch === '"') {
      let j = i + 1, s = '';
      for (;;) {
        if (j >= src.length) throw new FormulaError('Missing closing quotation mark', i);
        if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
        s += src[j++];
      }
      out.push({ type: 'str', text: s, pos: i, end: j + 1 });
      i = j + 1;
      continue;
    }
    // Structured references: Table1[Col], Table1[[#Headers],[Col]], [@Col]
    if (ch === '[') {
      const end = matchBracket(src, i);
      push('sref', src.slice(i, end + 1), {});
      continue;
    }
    let m: RegExpExecArray | null;
    if ((m = RE_REFERR.exec(rest))) {
      const sheet = m[1] ? unquoteSheet(m[1].slice(0, -1).split(':')[0]) : undefined;
      push('ref', m[0], { ref: { sheet, r1: 0, c1: 0, r2: 0, c2: 0, ar1: false, ac1: false, ar2: false, ac2: false, kind: 'cell', invalid: true } });
      continue;
    }
    if ((m = RE_ERR.exec(rest))) { push('err', m[0].toUpperCase()); continue; }
    if ((m = RE_RANGE.exec(rest))) {
      // Avoid treating function names like LOG10( or names like A1B as refs.
      const ref = refFromMatch(m);
      if (ref) { push('ref', m[0], { ref }); continue; }
    }
    if ((m = RE_NUM.exec(rest)) && !(out.length && out[out.length - 1].type === 'name' && out[out.length - 1].end === i)) { push('num', m[0]); continue; }
    if (ch === "'" ) {
      // Quoted sheet name prefix for a name or #REF!: 'Sheet'!Name
      const j = src.indexOf("'!", i + 1);
      if (j > 0) {
        const after = src.slice(j + 2);
        const nm = RE_IDENT.exec(after);
        if (nm) { push('name', src.slice(i, j + 2 + nm[0].length)); continue; }
      }
      throw new FormulaError('Unexpected quote', i);
    }
    if ((m = RE_IDENT.exec(rest))) {
      let id = m[0];
      // Sheet-qualified name: Sheet1!MyName
      if (src[i + id.length] === '!') {
        const nm = RE_IDENT.exec(src.slice(i + id.length + 1));
        if (nm) { push('name', src.slice(i, i + id.length + 1 + nm[0].length)); continue; }
      }
      // Table name followed by a structured reference
      if (src[i + id.length] === '[') {
        const end = matchBracket(src, i + id.length);
        push('sref', src.slice(i, end + 1));
        continue;
      }
      let j = i + id.length;
      while (j < src.length && src[j] === ' ') j++;
      if (src[j] === '(') { push('func', id); continue; }
      const up = id.toUpperCase();
      if (up === 'TRUE' || up === 'FALSE') { push('bool', id); continue; }
      // trailing dot is not part of a name
      if (id.endsWith('.')) id = id.replace(/\.+$/, '');
      push('name', id);
      continue;
    }
    if (ch === '(') { push('lparen', ch); continue; }
    if (ch === ')') { push('rparen', ch); continue; }
    if (ch === '{') { push('lbrace', ch); continue; }
    if (ch === '}') { push('rbrace', ch); continue; }
    if (ch === ',') { push('comma', ch); continue; }
    if (ch === ';') { push('semi', ch); continue; }
    const two = src.slice(i, i + 2);
    if (two === '<=' || two === '>=' || two === '<>') { push('op', two); continue; }
    if ('+-*/^&=<>%:@#!'.includes(ch)) { push('op', ch); continue; }
    throw new FormulaError(`Unexpected character "${ch}"`, i);
  }
  out.push({ type: 'eof', text: '', pos: src.length, end: src.length });
  return out;
}

function matchBracket(src: string, start: number): number {
  let depth = 0;
  for (let j = start; j < src.length; j++) {
    if (src[j] === "'" ) { j++; continue; } // escape char inside structured refs
    if (src[j] === '[') depth++;
    else if (src[j] === ']') { depth--; if (depth === 0) return j; }
  }
  throw new FormulaError('Missing closing bracket', start);
}

// ----------------------------------------------------------------------------------------------
// Parser

const BIN_PREC: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

export function parse(src: string): Node {
  const toks = tokenize(src);
  let p = 0;
  // Whitespace is significant only as the intersection operator between references.
  const peek = (): Token => { let k = p; while (toks[k].type === 'ws') k++; return toks[k]; };
  const next = (): Token => { while (toks[p].type === 'ws') p++; return toks[p++]; };
  const expect = (type: TokType, text?: string): Token => {
    const t = next();
    if (t.type !== type || (text && t.text !== text)) throw new FormulaError(`Expected ${text || type}`, t.pos);
    return t;
  };

  function parseExpr(minPrec = 0, inParens = false): Node {
    let left = parseUnary(inParens);
    for (;;) {
      const t = peek();
      if (t.type !== 'op' || !(t.text in BIN_PREC)) break;
      const prec = BIN_PREC[t.text];
      if (prec < minPrec) break;
      next();
      // ^ is left-associative in spreadsheets
      const right = parseExpr(prec + 1, inParens);
      left = { t: 'bin', op: t.text, a: left, b: right };
    }
    return left;
  }

  function parseUnary(inParens: boolean): Node {
    const t = peek();
    if (t.type === 'op' && (t.text === '-' || t.text === '+')) { next(); const a = parseUnary(inParens); return t.text === '+' ? a : { t: 'un', op: '-', a }; }
    if (t.type === 'op' && t.text === '@') { next(); return { t: 'isect', a: parseUnary(inParens) }; }
    return parsePostfix(inParens);
  }

  function parsePostfix(inParens: boolean): Node {
    let n = parseRefExpr(inParens);
    for (;;) {
      const t = toks[p]; // postfix operators must follow without whitespace except %
      const pk = peek();
      if (pk.type === 'op' && pk.text === '%') { next(); n = { t: 'pct', a: n }; continue; }
      if (t.type === 'op' && t.text === '#') { p++; n = { t: 'spill', a: n }; continue; }
      if (t.type === 'lparen' && (n.t === 'call' || n.t === 'fn' || n.t === 'name')) {
        // LAMBDA(...)(args) or name(...) invocation
        p++;
        const args = parseArgs();
        n = { t: 'call', callee: n, args };
        continue;
      }
      break;
    }
    return n;
  }

  // range ':' and intersection ' ' bind tighter than everything else
  function parseRefExpr(inParens: boolean): Node {
    let n = parsePrimary();
    for (;;) {
      const t = toks[p];
      if (t.type === 'op' && t.text === ':') { p++; const b = parsePrimary(); n = { t: 'bin', op: ':', a: n, b }; continue; }
      if (t.type === 'ws' && isRefLike(n)) {
        // intersection if the next non-ws token starts a reference
        let k = p; while (toks[k].type === 'ws') k++;
        const nt = toks[k];
        if (nt.type === 'ref' || nt.type === 'name' || (nt.type === 'func' && /^(INDEX|OFFSET|INDIRECT|CHOOSE|IF)$/i.test(nt.text)) || nt.type === 'lparen' || nt.type === 'sref') {
          p = k;
          const b = parsePrimary();
          n = { t: 'bin', op: ' ', a: n, b };
          continue;
        }
      }
      break;
    }
    void inParens;
    return n;
  }

  function isRefLike(n: Node): boolean {
    return n.t === 'ref' || n.t === 'name' || n.t === 'sref' || (n.t === 'bin' && (n.op === ':' || n.op === ' ')) || n.t === 'union' || (n.t === 'fn' && /^(INDEX|OFFSET|INDIRECT)$/.test(n.name));
  }

  function parseArgs(): Node[] {
    const args: Node[] = [];
    if (peek().type === 'rparen') { next(); return args; }
    for (;;) {
      const t = peek();
      if (t.type === 'comma' || t.type === 'rparen') args.push({ t: 'missing' });
      else args.push(parseExpr());
      const sep = next();
      if (sep.type === 'rparen') break;
      if (sep.type !== 'comma') throw new FormulaError('Expected , or )', sep.pos);
    }
    return args;
  }

  function parsePrimary(): Node {
    const t = next();
    switch (t.type) {
      case 'num': return { t: 'num', v: parseFloat(t.text) };
      case 'str': return { t: 'str', v: t.text };
      case 'bool': return { t: 'bool', v: t.text.toUpperCase() === 'TRUE' };
      case 'err': return { t: 'err', v: (errorFromCode(t.text)?.code || '#VALUE!') };
      case 'ref': return t.ref!.invalid ? { t: 'err', v: '#REF!' } : { t: 'ref', ref: t.ref! };
      case 'sref': {
        const k = t.text.indexOf('[');
        return { t: 'sref', table: t.text.slice(0, k), spec: t.text.slice(k) };
      }
      case 'name': {
        const bang = t.text.lastIndexOf('!');
        if (bang > 0) return { t: 'name', name: t.text.slice(bang + 1), sheet: unquoteSheet(t.text.slice(0, bang)) };
        return { t: 'name', name: t.text };
      }
      case 'func': {
        expect('lparen');
        let name = t.text.toUpperCase().replace(/^_XLFN\._XLWS\.|^_XLFN\.|^_XLWS\./, '');
        if (name.startsWith('_XLL.')) name = name.slice(5);
        const args = parseArgs();
        return { t: 'fn', name, args };
      }
      case 'lparen': {
        const first = parseExpr(0, true);
        if (peek().type === 'comma') {
          const items = [first];
          while (peek().type === 'comma') { next(); items.push(parseExpr(0, true)); }
          expect('rparen');
          return { t: 'union', items };
        }
        expect('rparen');
        return first;
      }
      case 'lbrace': {
        const rows: Node[][] = [[]];
        for (;;) {
          const v = parseUnary(false);
          rows[rows.length - 1].push(v);
          const sep = next();
          if (sep.type === 'rbrace') break;
          if (sep.type === 'semi') rows.push([]);
          else if (sep.type !== 'comma') throw new FormulaError('Expected , ; or } in array constant', sep.pos);
        }
        if (rows.some((r) => r.length !== rows[0].length)) throw new FormulaError('Array constant rows must be the same length', t.pos);
        return { t: 'arr', rows };
      }
      case 'op':
        if (t.text === '-' || t.text === '+') { const a = parseUnary(false); return t.text === '-' ? { t: 'un', op: '-', a } : a; }
        throw new FormulaError(`Unexpected "${t.text}"`, t.pos);
      case 'eof': throw new FormulaError('The formula is incomplete', t.pos);
      default: throw new FormulaError(`Unexpected "${t.text}"`, t.pos);
    }
  }

  const root = parseExpr();
  const rest = peek();
  if (rest.type !== 'eof') throw new FormulaError(`Unexpected "${rest.text}"`, rest.pos);
  return root;
}

// ----------------------------------------------------------------------------------------------
// Reference rewriting (fill, copy/paste, insert/delete rows and columns, sheet rename)

import { colName } from '../model/address';
import { quoteSheet } from '../model/address';

export function refText(ref: RefToken, withSheet = true): string {
  let pre = '';
  if (withSheet && ref.sheet) pre = quoteSheet(ref.sheet) + (ref.sheet2 ? ':' + quoteSheet(ref.sheet2) : '') + '!';
  if (ref.invalid) return pre + '#REF!';
  const cell = (r: number, c: number, ar: boolean, ac: boolean) => `${ac ? '$' : ''}${colName(c)}${ar ? '$' : ''}${r + 1}`;
  switch (ref.kind) {
    case 'cell': return pre + cell(ref.r1, ref.c1, ref.ar1, ref.ac1);
    case 'cols': return `${pre}${ref.ac1 ? '$' : ''}${colName(ref.c1)}:${ref.ac2 ? '$' : ''}${colName(ref.c2)}`;
    case 'rows': return `${pre}${ref.ar1 ? '$' : ''}${ref.r1 + 1}:${ref.ar2 ? '$' : ''}${ref.r2 + 1}`;
    default: return `${pre}${cell(ref.r1, ref.c1, ref.ar1, ref.ac1)}:${cell(ref.r2, ref.c2, ref.ar2, ref.ac2)}`;
  }
}

/** Rewrites every reference token in a formula. fn returns the new ref (or null to keep, or {invalid}). */
export function mapRefs(formula: string, fn: (ref: RefToken) => RefToken | null): string {
  let toks: Token[];
  try { toks = tokenize(formula); } catch { return formula; }
  let out = '';
  let last = 0;
  for (const t of toks) {
    if (t.type !== 'ref' || !t.ref || t.ref.invalid) continue;
    const nr = fn({ ...t.ref });
    if (!nr) continue;
    out += formula.slice(last, t.pos) + refText(nr, !!nr.sheet);
    last = t.end;
  }
  return out + formula.slice(last);
}

/** Relative-reference adjustment for copy/fill by (dr, dc). Out-of-sheet refs become #REF!. */
export function shiftFormula(formula: string, dr: number, dc: number): string {
  if (!dr && !dc) return formula;
  return mapRefs(formula, (ref) => {
    const r = { ...ref };
    if (r.kind !== 'cols') { if (!r.ar1) r.r1 += dr; if (!r.ar2) r.r2 += dr; }
    if (r.kind !== 'rows') { if (!r.ac1) r.c1 += dc; if (!r.ac2) r.c2 += dc; }
    if (r.r1 < 0 || r.c1 < 0 || r.r2 < 0 || r.c2 < 0 || r.r1 >= MAX_ROWS || r.r2 >= MAX_ROWS || r.c1 >= MAX_COLS || r.c2 >= MAX_COLS) return { ...r, invalid: true };
    return r;
  });
}

/** Names of all functions used (for validation, #NAME? detection and dependency hints). */
export function functionsIn(node: Node, acc = new Set<string>()): Set<string> {
  walk(node, (n) => { if (n.t === 'fn') acc.add(n.name); });
  return acc;
}

export function walk(node: Node, fn: (n: Node) => void): void {
  fn(node);
  switch (node.t) {
    case 'arr': node.rows.forEach((r) => r.forEach((x) => walk(x, fn))); break;
    case 'fn': node.args.forEach((x) => walk(x, fn)); break;
    case 'call': walk(node.callee, fn); node.args.forEach((x) => walk(x, fn)); break;
    case 'un': case 'pct': case 'spill': case 'isect': walk(node.a, fn); break;
    case 'bin': walk(node.a, fn); walk(node.b, fn); break;
    case 'union': node.items.forEach((x) => walk(x, fn)); break;
  }
}
