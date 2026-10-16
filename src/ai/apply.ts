/**
 * Spreadsheet side of the AI features: building context from cells, reading what the model
 * returns (formulas, JSON tables, per-row answers, chart and formatting specs) and turning it into
 * workbook changes. Parsing functions are pure and tested with mocked model output.
 */
import type { App } from '../app';
import type { Sheet } from '../model/workbook';
import type { Engine } from '../calc/engine';
import type { CFRule, ChartDef, ChartType, Scalar, Matrix, CFType, DataValidation } from '../model/types';
import { CellError } from '../model/types';
import { colName, cellName, rangeName, parseRange, type Range } from '../model/address';
import { rangeRefText } from '../model/structure';
import { parse, walk, FormulaError } from '../calc/parser';
import { getFn } from '../calc/functions';
import { newId } from '../model/workbook';
import { cellText, setInput } from '../commands/ops';
import type { ChatRequest, ChatResult, ChatMessage } from './client';

// ============================================================================================
// Context

export interface TableContext { text: string; headers: string[]; header: boolean; rows: number; shown: number; truncated: boolean; range: Range }

function typeOf(sheet: Sheet, rg: Range, c: number, r1: number): string {
  let n = 0, t = 0, d = 0, b = 0;
  for (let r = r1; r <= Math.min(rg.r2, r1 + 200); r++) {
    const v = sheet.value(r, c);
    if (typeof v === 'number') { if (/[dmy]/i.test(sheet.styleOf(r, c).numFmt || '') && !/^\[/.test(sheet.styleOf(r, c).numFmt || '')) d++; else n++; }
    else if (typeof v === 'string' && v !== '') t++;
    else if (typeof v === 'boolean') b++;
  }
  const max = Math.max(n, t, d, b);
  if (!max) return 'empty';
  return max === d ? 'date' : max === n ? 'number' : max === b ? 'logical' : 'text';
}

/**
 * The cells of a range as tab-separated text with row numbers and column letters, for prompts.
 * Long tables keep the first rows (and a few of the last) within maxChars.
 */
export function tableText(sheet: Sheet, rg: Range, opts: { header?: boolean; maxRows?: number; maxChars?: number; formulas?: boolean } = {}): TableContext {
  const header = !!opts.header;
  const maxChars = opts.maxChars ?? 12000;
  const maxRows = opts.maxRows ?? 400;
  const r1 = header ? rg.r1 + 1 : rg.r1;
  const headers: string[] = [];
  const cols: string[] = [];
  for (let c = rg.c1; c <= rg.c2; c++) {
    const name = header ? cellText(sheet, rg.r1, c) : '';
    headers.push(name || colName(c));
    cols.push(`${colName(c)}${name ? ` "${name}"` : ''} (${typeOf(sheet, rg, c, r1)})`);
  }
  const line = (r: number) => {
    const cells: string[] = [String(r + 1)];
    for (let c = rg.c1; c <= rg.c2; c++) {
      const x = sheet.get(r, c);
      let t = cellText(sheet, r, c).replace(/[\t\r\n]+/g, ' ');
      if (opts.formulas && x?.f != null) t = `=${x.f} → ${t}`;
      cells.push(t.length > 300 ? t.slice(0, 300) + '…' : t);
    }
    return cells.join('\t');
  };
  const lines: string[] = [];
  let size = 0, shown = 0;
  const total = rg.r2 - rg.r1 + 1;
  const bodyRows = rg.r2 - r1 + 1;
  const tail = Math.min(5, Math.max(0, bodyRows - 20));
  const headLine = ['Row', ...Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => colName(rg.c1 + i))].join('\t');
  if (header) { const l = line(rg.r1); lines.push(l); size += l.length; }
  let truncated = false;
  for (let r = r1; r <= rg.r2; r++) {
    if (sheet.hiddenRows.has(r) && sheet.filter) continue;
    if (shown >= maxRows || size > maxChars) {
      truncated = true;
      const skipTo = Math.max(r, rg.r2 - tail + 1);
      lines.push(`… (${skipTo - r} rows not shown)`);
      for (let k = skipTo; k <= rg.r2; k++) lines.push(line(k));
      break;
    }
    const l = line(r);
    lines.push(l);
    size += l.length + 1;
    shown++;
  }
  const desc = `Sheet "${sheet.name}", range ${rangeName(rg)}, ${total} rows${header ? ' (row ' + (rg.r1 + 1) + ' is the header row)' : ''}. Columns: ${cols.join(', ')}.`;
  return { text: `${desc}\n${headLine}\n${lines.join('\n')}`, headers, header, rows: bodyRows, shown, truncated, range: rg };
}

/** Context for one cell: its formula, result and number format, the values it refers to, and the surrounding table. */
export function cellContext(sheet: Sheet, r: number, c: number, around?: TableContext): string {
  const x = sheet.get(r, c);
  const parts = [`Cell ${cellName(r, c)} on sheet "${sheet.name}".`];
  if (x?.f != null) parts.push(`Formula: =${x.f}`);
  const v = x?.v ?? null;
  parts.push(`Current result: ${v instanceof CellError ? v.code + (v.detail ? ` (${v.detail})` : '') : v == null ? '(empty)' : JSON.stringify(v)}`);
  const fmt = sheet.styleOf(r, c).numFmt;
  if (fmt && fmt !== 'General') parts.push(`Number format: ${fmt}`);
  if (x?.f) {
    const refs = refsIn(x.f).slice(0, 12);
    const vals: string[] = [];
    for (const ref of refs) {
      const rg = parseRange(ref.replace(/\$/g, ''));
      if (!rg || ref.includes('!')) continue;
      const cells: string[] = [];
      for (let rr = rg.r1; rr <= Math.min(rg.r2, rg.r1 + 5); rr++) for (let cc = rg.c1; cc <= Math.min(rg.c2, rg.c1 + 5); cc++) cells.push(`${cellName(rr, cc)}=${cellText(sheet, rr, cc) || '(empty)'}`);
      vals.push(`${ref}: ${cells.join(', ')}${rg.r2 - rg.r1 > 5 || rg.c2 - rg.c1 > 5 ? ', …' : ''}`);
    }
    if (vals.length) parts.push(`Referenced values:\n${vals.join('\n')}`);
  }
  if (around) parts.push(`\nThe table around it:\n${around.text}`);
  return parts.join('\n');
}

export function refsIn(formula: string): string[] {
  const out: string[] = [];
  for (const m of formula.matchAll(/(?:'[^']+'|[A-Za-z_][\w.]*)?!?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/g)) if (!out.includes(m[0])) out.push(m[0]);
  return out;
}

// ============================================================================================
// Cleaning model output

/** Removes reasoning blocks, chatty preambles, code fences and wrapping quotes. */
export function cleanOutput(raw: string): string {
  let t = (raw || '').replace(/\r\n?/g, '\n').trim();
  t = t.replace(/^<think>[\s\S]*?<\/think>\s*/i, '');
  const fence = /^```[\w-]*\n([\s\S]*?)\n?```$/.exec(t);
  if (fence) t = fence[1].trim();
  const lines = t.split('\n');
  if (lines.length > 1 && /^(sure|certainly|of course|here(?:'s| is| are)|below is|okay|ok)\b.{0,120}[:：]\s*$/i.test(lines[0])) t = lines.slice(1).join('\n').trim();
  if (/^["“].*["”]$/s.test(t) && !/["“”]/.test(t.slice(1, -1))) t = t.slice(1, -1);
  return t;
}

/** Balanced-parenthesis end of a formula that starts at s[0] (stops before trailing prose). */
function formulaEnd(s: string): number {
  let depth = 0, inStr = false, inQuote = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (inStr) { if (ch === '"') { if (s[i + 1] === '"') i++; else inStr = false; } continue; }
    if (inQuote) { if (ch === "'") inQuote = false; continue; }
    if (ch === '"') inStr = true;
    else if (ch === "'") inQuote = true;
    else if (ch === '(' || ch === '{') depth++;
    else if (ch === ')' || ch === '}') depth--;
    else if (depth === 0 && /\s/.test(ch)) {
      // A space at depth 0 ends the formula when prose follows (" - sums …", " which …").
      const rest = s.slice(i).trim();
      if (/^[–—:]/.test(rest) || /^-\s+[a-z]{2,}\b(?![\d(!:.])/.test(rest) || (/^\(?[a-z]{3,}\b(?![\d(!:.])/i.test(rest) && !/^[A-Z]{1,3}\d/.test(rest))) return i;
    }
  }
  return s.length;
}

/**
 * Finds the formula in a model answer: a line starting with "=", inside a code fence or inline
 * code, or a bare function call. Returns the formula (with "=") and the remaining explanation.
 */
export function extractFormula(raw: string): { formula: string; explanation: string } | null {
  let t = (raw || '').replace(/\r\n?/g, '\n').replace(/^<think>[\s\S]*?<\/think>\s*/i, '').trim();
  let explanation = '';
  const fence = /```[\w-]*\n?([\s\S]*?)```/.exec(t);
  if (fence) {
    const inner = fence[1].trim().split('\n').find((l) => l.trim());
    if (inner) { explanation = (t.slice(0, fence.index) + t.slice(fence.index + fence[0].length)).trim(); t = inner.trim(); }
  }
  const lines = t.split('\n');
  for (let i = 0; i < lines.length; i++) {
    let l = lines[i].trim().replace(/^(?:formula|answer|result)\s*[:：]\s*/i, '').replace(/^[-*>]\s+/, '');
    const inline = /`(=[^`]+)`/.exec(l);
    if (inline) l = inline[1];
    l = l.replace(/^`+|`+$/g, '').trim();
    let f: string | null = null;
    if (l.startsWith('=') && l.length > 1) f = l;
    else if (/^[A-Z][A-Z0-9.]*\(/.test(l)) f = '=' + l;
    if (!f) continue;
    f = f.slice(0, formulaEnd(f)).trim().replace(/[.;,]$/, '').replace(/^==/, '=');
    const rest = [explanation, ...lines.slice(i + 1)].join('\n').replace(/^\s*(explanation|why)\s*[:：]\s*/im, '').trim();
    return { formula: f, explanation: rest };
  }
  return null;
}

export interface FormulaCheck { ok: boolean; error?: string; value?: Scalar; matrix?: Matrix; resultError?: string }

/** Checks a formula with the parser (syntax, known functions, argument counts) and evaluates it at (r, c). */
export function validateFormula(engine: Engine, sheet: Sheet, r: number, c: number, formula: string): FormulaCheck {
  const f = formula.trim().replace(/^=/, '');
  if (!f) return { ok: false, error: 'The formula is empty.' };
  let node;
  try { node = parse(f); } catch (e) { return { ok: false, error: `Syntax error: ${e instanceof FormulaError ? e.message : String(e)}` }; }
  const names = new Set(sheet.book.names.map((n) => n.name.toUpperCase()));
  const problems: string[] = [];
  walk(node, (n) => {
    if (n.t !== 'fn') return;
    const def = getFn(n.name);
    if (!def) { if (!names.has(n.name.toUpperCase())) problems.push(`Unknown function ${n.name.toUpperCase()}`); return; }
    if (n.args.length < def.min || n.args.length > def.max) problems.push(`${def.name} takes ${def.min === def.max ? def.min : `${def.min} to ${def.max === Infinity ? 'many' : def.max}`} arguments (got ${n.args.length})`);
  });
  if (problems.length) return { ok: false, error: problems.join('; ') };
  let matrix: Matrix;
  try { matrix = engine.evaluateMatrix(f, sheet, r, c); } catch (e) { return { ok: false, error: `Could not evaluate: ${(e as Error).message}` }; }
  const value = matrix[0]?.[0] ?? null;
  const res: FormulaCheck = { ok: true, value, matrix };
  if (value instanceof CellError) res.resultError = value.code;
  return res;
}

/** Parses JSON from a model answer, tolerating code fences, prose around it and trailing commas. */
export function parseJsonLoose(raw: string): any {
  let t = (raw || '').replace(/^<think>[\s\S]*?<\/think>\s*/i, '').trim();
  const fence = /```(?:json)?\s*\n?([\s\S]*?)```/.exec(t);
  if (fence) t = fence[1].trim();
  const a = t.indexOf('['), o = t.indexOf('{');
  const start = a < 0 ? o : o < 0 ? a : Math.min(a, o);
  if (start < 0) return undefined;
  const close = t[start] === '[' ? ']' : '}';
  const end = t.lastIndexOf(close);
  if (end <= start) return undefined;
  const body = t.slice(start, end + 1);
  for (const candidate of [body, body.replace(/,\s*([}\]])/g, '$1'), body.replace(/[“”]/g, '"').replace(/,\s*([}\]])/g, '$1')]) {
    try { return JSON.parse(candidate); } catch { /* try next */ }
  }
  return undefined;
}

function jsonScalar(v: any): Scalar {
  if (v == null) return null;
  if (typeof v === 'number' || typeof v === 'boolean' || typeof v === 'string') return v;
  if (Array.isArray(v)) return v.map((x) => (x && typeof x === 'object' ? JSON.stringify(x) : String(x))).join(', ');
  return JSON.stringify(v);
}

/**
 * Reads a table from a model answer: JSON (array of objects → header row + rows; array of arrays;
 * {columns, rows}; an object holding an array), a Markdown table, or tab/comma-separated lines.
 */
export function parseTable(raw: string): Scalar[][] | null {
  const j = parseJsonLoose(raw);
  const fromJson = (x: any): Scalar[][] | null => {
    if (Array.isArray(x)) {
      if (!x.length) return null;
      if (x.every((r) => Array.isArray(r))) return x.map((r: any[]) => r.map(jsonScalar));
      if (x.every((r) => r && typeof r === 'object')) {
        const keys: string[] = [];
        for (const r of x) for (const k of Object.keys(r)) if (!keys.includes(k)) keys.push(k);
        return [keys, ...x.map((r: any) => keys.map((k) => jsonScalar(r[k])))];
      }
      return [['Value'], ...x.map((v) => [jsonScalar(v)])];
    }
    if (x && typeof x === 'object') {
      if (Array.isArray(x.columns) && Array.isArray(x.rows)) return [x.columns.map(String), ...x.rows.map((r: any) => (Array.isArray(r) ? r.map(jsonScalar) : x.columns.map((k: string) => jsonScalar(r[k]))))];
      for (const k of ['rows', 'data', 'items', 'table', 'records', 'result']) if (Array.isArray(x[k])) return fromJson(x[k]);
      const arr = Object.values(x).find((v) => Array.isArray(v));
      if (arr) return fromJson(arr);
    }
    return null;
  };
  if (j !== undefined) { const t = fromJson(j); if (t) return t; }
  const text = cleanOutput(raw);
  const lines = text.split('\n').map((l) => l.trim()).filter(Boolean);
  const md = lines.filter((l) => l.startsWith('|') || (l.split('|').length > 2));
  if (md.length >= 2) {
    const rows = md.filter((l) => !/^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?$/.test(l)).map((l) => l.replace(/^\|/, '').replace(/\|$/, '').split('|').map((c) => c.trim().replace(/^\*\*(.*)\*\*$/, '$1')));
    if (rows.length >= 2) return rows;
  }
  if (lines.length >= 2 && lines.every((l) => l.includes('\t'))) return lines.map((l) => l.split('\t'));
  const commas = lines.map((l) => l.split(',').length);
  if (lines.length >= 2 && commas[0] > 1 && commas.every((n) => n === commas[0])) return lines.map((l) => l.split(',').map((c) => c.trim().replace(/^"(.*)"$/, '$1')));
  return null;
}

// ============================================================================================
// Row-by-row jobs

export type RowAnswer = string | Record<string, string> | null;

export function rowBatchPrompt(instruction: string, items: string[], opts: { keys?: string[]; examples?: { input: string; output: string }[]; context?: string } = {}): string {
  const shape = opts.keys?.length ? `an array of exactly ${items.length} objects, each with the keys ${opts.keys.map((k) => JSON.stringify(k)).join(', ')} (string values)` : `an array of exactly ${items.length} strings`;
  const parts = [`Task: ${instruction}`];
  if (opts.context) parts.push(opts.context);
  if (opts.examples?.length) parts.push('Examples of the expected result:\n' + opts.examples.map((e) => `Input: ${e.input}\nOutput: ${e.output}`).join('\n\n'));
  parts.push(`Apply the task to each numbered item below. Respond with JSON only: ${shape}, in the same order as the items.`);
  parts.push('Items:\n' + items.map((t, i) => `${i + 1}. ${t.replace(/\n+/g, ' ⏎ ')}`).join('\n'));
  return parts.join('\n\n');
}

/** Reads per-item answers (a JSON array, or numbered lines). Returns null when the count does not match. */
export function parseRowAnswers(raw: string, n: number, keys?: string[]): RowAnswer[] | null {
  const j = parseJsonLoose(raw);
  let arr: any[] | null = Array.isArray(j) ? j : j && typeof j === 'object' ? (Object.values(j).find((v) => Array.isArray(v)) as any[] | undefined) || null : null;
  if (arr && arr.length === n) {
    return arr.map((x) => {
      if (keys?.length) {
        if (x && typeof x === 'object' && !Array.isArray(x)) {
          const lower = new Map(Object.entries(x).map(([k, v]) => [k.toLowerCase(), v]));
          return Object.fromEntries(keys.map((k) => [k, x[k] != null ? String(x[k]) : lower.has(k.toLowerCase()) && lower.get(k.toLowerCase()) != null ? String(lower.get(k.toLowerCase())) : '']));
        }
        if (Array.isArray(x)) return Object.fromEntries(keys.map((k, i) => [k, x[i] == null ? '' : String(x[i])]));
        return Object.fromEntries(keys.map((k, i) => [k, i === 0 && x != null ? String(x) : '']));
      }
      if (x == null) return null;
      if (typeof x === 'object' && !Array.isArray(x)) { const vals = Object.values(x); return vals.length === 1 ? String(vals[0] ?? '') : JSON.stringify(x); }
      return Array.isArray(x) ? x.join(', ') : String(x);
    });
  }
  if (arr) return null;
  if (keys?.length) return null;
  // Numbered lines: "1. answer"
  const out: RowAnswer[] = new Array(n).fill(null);
  let found = 0;
  for (const line of cleanOutput(raw).split('\n')) {
    const m = /^\s*(\d+)[.):\]]\s*(.*)$/.exec(line);
    if (!m) continue;
    const i = +m[1] - 1;
    if (i >= 0 && i < n && out[i] == null) { out[i] = m[2].trim().replace(/^"(.*)"$/, '$1'); found++; }
  }
  if (n === 1 && !found) { const t = cleanOutput(raw); return t ? [t] : null; }
  return found === n ? out : null;
}

export function batchItems(items: string[], maxItems = 25, maxChars = 6000): number[][] {
  const batches: number[][] = [];
  let cur: number[] = [], size = 0;
  items.forEach((t, i) => {
    if (cur.length && (cur.length >= maxItems || size + t.length > maxChars)) { batches.push(cur); cur = []; size = 0; }
    cur.push(i); size += t.length;
  });
  if (cur.length) batches.push(cur);
  return batches;
}

export type ChatFn = (req: ChatRequest) => Promise<ChatResult>;

export interface RowJob {
  chat: ChatFn;
  items: string[];
  instruction: string;
  system: string;
  keys?: string[];
  examples?: { input: string; output: string }[];
  context?: string;
  tier?: 'fast' | 'quality';
  temperature?: number;
  signal?: AbortSignal;
  batchSize?: number;
  onProgress?: (done: number, total: number) => void;
}

/** Runs an instruction over many items in batches. A batch whose answer does not line up is retried one item at a time. */
export async function runRowJob(job: RowJob): Promise<RowAnswer[]> {
  const results: RowAnswer[] = new Array(job.items.length).fill(null);
  const batches = batchItems(job.items, job.batchSize ?? 20, 6000);
  let done = 0;
  const ask = async (idx: number[]) => {
    const msgs: ChatMessage[] = [{ role: 'system', content: job.system }, { role: 'user', content: rowBatchPrompt(job.instruction, idx.map((i) => job.items[i]), { keys: job.keys, examples: job.examples, context: job.context }) }];
    const r = await job.chat({ messages: msgs, tier: job.tier ?? 'fast', temperature: job.temperature ?? 0, signal: job.signal, maxTokens: Math.min(8000, 200 + idx.reduce((a, i) => a + Math.ceil(job.items[i].length / 2) + 40, 0) * (job.keys?.length || 1)) });
    return r.truncated ? null : parseRowAnswers(r.text, idx.length, job.keys);
  };
  for (const b of batches) {
    if (job.signal?.aborted) break;
    let ans = await ask(b);
    if (!ans && b.length > 1) {
      ans = [];
      for (const i of b) {
        if (job.signal?.aborted) break;
        const one = await ask([i]);
        ans.push(one ? one[0] : null);
      }
    }
    (ans || []).forEach((a, k) => { results[b[k]] = a; });
    done += b.length;
    job.onProgress?.(done, job.items.length);
  }
  return results;
}

// ============================================================================================
// Specs to workbook objects

const CHART_TYPES: ChartType[] = ['column', 'stackedColumn', 'percentColumn', 'bar', 'stackedBar', 'line', 'lineMarkers', 'area', 'stackedArea', 'pie', 'doughnut', 'scatter', 'scatterLines', 'radar', 'bubble'];

function findCol(headers: string[], name: any, rg: Range): number {
  if (name == null) return -1;
  const s = String(name).trim().toLowerCase();
  let i = headers.findIndex((x) => x.trim().toLowerCase() === s);
  if (i < 0 && /^[a-z]{1,3}$/i.test(s)) { const c = headers.findIndex((_x, k) => colName(rg.c1 + k).toLowerCase() === s); i = c; }
  if (i < 0) i = headers.findIndex((x) => x.toLowerCase().includes(s) || s.includes(x.toLowerCase()));
  return i;
}

/** Builds a chart from {type, title, x, y[]} with header names of the data region. */
export function chartFromSpec(sheet: Sheet, rg: Range, headers: string[], header: boolean, spec: any): ChartDef | null {
  if (!spec || typeof spec !== 'object') return null;
  const type = (CHART_TYPES.includes(spec.type) ? spec.type : 'column') as ChartType;
  const r1 = header ? rg.r1 + 1 : rg.r1;
  if (r1 > rg.r2) return null;
  const xi = findCol(headers, spec.x, rg);
  const ys = (Array.isArray(spec.y) ? spec.y : spec.y != null ? [spec.y] : []).map((y: any) => findCol(headers, y, rg)).filter((i: number) => i >= 0 && i !== xi);
  if (!ys.length) {
    for (let i = 0; i < headers.length; i++) if (i !== xi && typeof sheet.value(r1, rg.c1 + i) === 'number') ys.push(i);
  }
  if (!ys.length) return null;
  const col = (i: number) => ({ r1, c1: rg.c1 + i, r2: rg.r2, c2: rg.c1 + i });
  const series = ys.slice(0, type === 'pie' || type === 'doughnut' ? 1 : 12).map((i: number) => ({
    name: headers[i], ...(header ? { nameRef: rangeRefText(sheet, { r1: rg.r1, c1: rg.c1 + i, r2: rg.r1, c2: rg.c1 + i }) } : {}),
    values: rangeRefText(sheet, col(i)), ...(xi >= 0 ? { categories: rangeRefText(sheet, col(xi)) } : {}),
  }));
  return {
    id: newId('ch'), type, title: typeof spec.title === 'string' ? spec.title.slice(0, 120) : headers[ys[0]], series,
    at: { r: rg.r1, c: rg.c2 + 2, dx: 0, dy: 0, w: 480, h: 288 }, legend: series.length > 1 || type === 'pie' || type === 'doughnut' ? 'bottom' : 'none',
    gridlines: !['pie', 'doughnut'].includes(type), sourceRange: rangeRefText(sheet, rg),
  };
}

const CF_TYPES: CFType[] = ['cellIs', 'expression', 'colorScale', 'dataBar', 'iconSet', 'top10', 'aboveAverage', 'duplicateValues', 'uniqueValues', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors', 'timePeriod'];
const OPS = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'];
const PERIODS = ['today', 'yesterday', 'tomorrow', 'last7Days', 'thisWeek', 'lastWeek', 'nextWeek', 'thisMonth', 'lastMonth', 'nextMonth'];
const ICONSETS = ['3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3Signs', '4Arrows', '5Arrows', '3Symbols', '4Rating', '5Rating', '5Quarters'];
const color = (c: any) => (typeof c === 'string' && /^#[0-9a-f]{3,8}$/i.test(c.trim()) ? c.trim() : undefined);

/** Validated conditional-formatting rules from model JSON. Rules with unknown types or bad ranges are dropped. */
export function cfRulesFromSpec(spec: any, defaultRanges: Range[]): CFRule[] {
  const list = Array.isArray(spec) ? spec : spec && Array.isArray(spec.rules) ? spec.rules : spec ? [spec] : [];
  const out: CFRule[] = [];
  for (const x of list) {
    if (!x || typeof x !== 'object' || !CF_TYPES.includes(x.type)) continue;
    let ranges = defaultRanges;
    if (typeof x.range === 'string' && x.range.trim()) {
      const parsed = x.range.split(/[,\s]+/).filter(Boolean).map((t: string) => parseRange(t.replace(/\$/g, '').replace(/^.*!/, ''))).filter(Boolean) as Range[];
      if (parsed.length) ranges = parsed;
    }
    const rule: CFRule = { id: newId('cf'), type: x.type, ranges };
    if (x.type === 'cellIs') { rule.operator = OPS.includes(x.operator) ? x.operator : 'greaterThan'; }
    const fs = (Array.isArray(x.formulas) ? x.formulas : x.formula != null ? [x.formula] : x.value != null ? [x.value] : []).map((f: any) => String(f).replace(/^=/, '').trim()).filter(Boolean).slice(0, 2);
    if (x.type === 'cellIs' || x.type === 'expression') {
      if (!fs.length) continue;
      if (fs.some((f: string) => { try { parse(f); return false; } catch { return true; } })) continue;
      rule.formulas = fs;
    }
    if (['containsText', 'notContainsText', 'beginsWith', 'endsWith'].includes(x.type)) { if (typeof x.text !== 'string' || !x.text) continue; rule.text = x.text; }
    if (x.type === 'top10') { rule.rank = Math.max(1, Math.min(1000, Number(x.rank) || 10)); rule.percent = !!x.percent; rule.bottom = !!x.bottom; }
    if (x.type === 'aboveAverage') { rule.above = x.above !== false; }
    if (x.type === 'timePeriod') rule.timePeriod = PERIODS.includes(x.timePeriod) ? x.timePeriod : 'today';
    if (x.type === 'colorScale') { const cs = (Array.isArray(x.colors) ? x.colors.map(color).filter(Boolean) : []) as string[]; rule.colors = cs.length >= 2 ? cs.slice(0, 3) : ['#f8696b', '#ffeb84', '#63be7b']; }
    if (x.type === 'dataBar') rule.barColor = color(x.barColor) || '#638ec6';
    if (x.type === 'iconSet') rule.iconSet = ICONSETS.includes(x.iconSet) ? x.iconSet : '3TrafficLights1';
    if (!['colorScale', 'dataBar', 'iconSet'].includes(x.type)) {
      const st = x.style && typeof x.style === 'object' ? x.style : {};
      const style: Record<string, any> = {};
      if (color(st.fill)) style.fill = color(st.fill);
      if (color(st.color)) style.color = color(st.color);
      if (st.bold) style.bold = true;
      if (st.italic) style.italic = true;
      if (st.strike) style.strike = true;
      rule.style = Object.keys(style).length ? style : { fill: '#ffc7ce', color: '#9c0006' };
    }
    out.push(rule);
  }
  return out;
}

/** A list validation for the ranges from {items, title, message}. */
export function validationFromSpec(spec: any, ranges: Range[]): DataValidation | null {
  const items: string[] = (Array.isArray(spec) ? spec : Array.isArray(spec?.items) ? spec.items : []).map((x: any) => String(x).replace(/[,"]/g, ' ').trim()).filter(Boolean);
  const uniq = [...new Set(items)].slice(0, 100);
  if (!uniq.length) return null;
  const formula = `"${uniq.join(',')}"`;
  if (formula.length > 255) return null;
  return {
    id: newId('dv'), ranges, type: 'list', formula1: formula, allowBlank: true, showDropdown: true, errorStyle: 'stop',
    ...(typeof spec?.title === 'string' && spec.title ? { inputTitle: spec.title.slice(0, 32) } : {}),
    ...(typeof spec?.message === 'string' && spec.message ? { inputMessage: spec.message.slice(0, 255) } : {}),
  };
}

// ============================================================================================
// Writing

/** Writes a table at (r0, c0): text is typed in like user input (numbers, dates, "=" formulas). The first row is bold when header. */
export function writeTable(app: App, sheet: Sheet, r0: number, c0: number, rows: Scalar[][], header = true): Range {
  const book = sheet.book;
  const width = Math.max(1, ...rows.map((r) => r.length));
  rows.forEach((row, i) => row.forEach((v, j) => {
    const r = r0 + i, c = c0 + j;
    if (v == null || v === '') { const cur = sheet.get(r, c); sheet.set(r, c, cur?.s ? { s: cur.s } : undefined); return; }
    if (typeof v === 'string') setInput(app, sheet, r, c, v);
    else sheet.set(r, c, { ...(sheet.get(r, c)?.s ? { s: sheet.get(r, c)!.s } : {}), v });
    if (header && i === 0) { const cur = sheet.get(r, c)!; sheet.set(r, c, { ...cur, s: book.derivedStyle(cur.s ?? 0, { bold: true }) }); }
  }));
  return { r1: r0, c1: c0, r2: r0 + rows.length - 1, c2: c0 + width - 1 };
}

/** Plain value for an AI answer written into a cell: numbers stay numbers when numeric is set. */
export function answerValue(text: string, numeric = false): Scalar {
  const t = text.trim();
  if (numeric || /^-?\d+(\.\d+)?$/.test(t)) { const n = Number(t.replace(/,/g, '')); if (t !== '' && Number.isFinite(n)) return n; }
  return t;
}
