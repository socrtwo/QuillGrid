/**
 * Spelling for cell text.
 *
 * Spreadsheets do not mark misspellings as you type; spelling is checked on demand (Review ▸
 * Spelling, F7), walking the text constants of the selection or the sheet. Words are checked in a
 * background worker (spell.worker.ts: nspell with the SCOWL English dictionaries, or Hunspell for a
 * downloaded dictionary of another language, see dictionaries.ts).
 *
 * The walk itself (SpellSession) is independent of the worker and the UI: it takes any
 * SpellChecker and reads and writes cell text through the sheet, so it can be tested directly.
 */
import type { Sheet } from '../model/workbook';
import type { Cell } from '../model/types';
import { keyRow, keyCol, type Range } from '../model/address';
import { installedDictionaries, dictionaryFor, loadDictionary } from './dictionaries';

// --------------------------------------------------------------------------------------------
// Worker protocol (same messages as Anne's Pen's proofing engine)

let worker: Worker | null = null;
let workerFailures = 0;
let reqId = 0;
const pending = new Map<number, (v: any) => void>();
/** The answer to every request a failed (or missing) worker leaves unanswered. */
const FAILED = { bad: [], suggestions: [], failed: true };
/** Downloaded dictionaries already sent to the current worker. */
const inWorker = new Set<string>();

function getWorker(): Worker | null {
  if (worker) return worker;
  if (workerFailures >= 3 || typeof Worker === 'undefined') return null;
  try {
    const w = new Worker(new URL('./spell.worker.ts', import.meta.url), { type: 'module' });
    w.onmessage = (e) => { const cb = pending.get(e.data.id); if (cb) { pending.delete(e.data.id); cb(e.data); } };
    w.onerror = () => {
      // A crashed worker never answers: settle every waiting request so the walk finishes.
      workerFailures++;
      if (worker === w) { worker = null; inWorker.clear(); }
      try { w.terminate(); } catch { /* */ }
      for (const cb of pending.values()) cb(FAILED);
      pending.clear();
    };
    worker = w;
  } catch {
    worker = null;
  }
  return worker;
}

function call(msg: Record<string, any>): Promise<any> {
  const w = getWorker();
  if (!w) return Promise.resolve(FAILED);
  const id = ++reqId;
  return new Promise((res) => { pending.set(id, res); w.postMessage({ ...msg, id }); });
}

/** Send a downloaded dictionary to the worker once. Returns false if it is not available. */
async function ensureLoaded(id: string): Promise<boolean> {
  if (inWorker.has(id) && worker) return true;
  const d = await loadDictionary(id);
  if (!d) return false;
  const r = await call({ type: 'load', dialect: 'dict:' + id, aff: d.aff, dic: d.dic });
  if (r.failed || r.error) return false;
  inWorker.add(id);
  return true;
}

/** Forget a dictionary in the worker (after it is removed). */
export async function unloadDictionary(id: string): Promise<void> {
  inWorker.delete(id);
  if (worker) await call({ type: 'unload', dialect: 'dict:' + id });
}

// --------------------------------------------------------------------------------------------
// Dialects and languages

export type SpellDialect = 'any' | 'document' | 'en-US' | 'en-GB' | 'en-CA' | 'en-AU';

export const SPELL_DIALECTS: { value: SpellDialect; label: string }[] = [
  { value: 'any', label: 'English — all varieties (accept any standard spelling)' },
  { value: 'document', label: 'English — follow the workbook language' },
  { value: 'en-US', label: 'English (United States)' },
  { value: 'en-GB', label: 'English (United Kingdom, including Oxford -ize spelling)' },
  { value: 'en-CA', label: 'English (Canada)' },
  { value: 'en-AU', label: 'English (Australia)' },
];

/** The English dictionary set for a preference and a language tag (BCP 47). */
export function dialectFor(pref: SpellDialect | undefined, lang: string | undefined): 'any' | 'en-US' | 'en-GB' | 'en-CA' | 'en-AU' {
  if (pref && pref !== 'document') return pref;
  const l = (lang || '').toLowerCase().replace(/_/g, '-').replace(/^en-[a-z]{4}(?=-|$)/, 'en');
  if (!l || !isEnglish(l)) return 'any';
  if (/^en-(us|as|gu|mp|pr|um|vi|ph)\b/.test(l)) return 'en-US';
  if (/^en-ca\b/.test(l)) return 'en-CA';
  if (/^en-(au|nz)\b/.test(l)) return 'en-AU';
  if (l === 'en') return 'any';
  return 'en-GB';
}

/** True for an English language tag ("en", "en-GB", "en-IN"...) or no tag. */
export function isEnglish(lang: string | null | undefined): boolean {
  return !lang || /^en([-_]|$)/i.test(lang);
}

/**
 * The dictionary key the worker checks words with for a workbook language: an English variety,
 * 'dict:<id>' for a downloaded dictionary, or null when no dictionary covers the language.
 */
export async function dictionaryKey(lang: string | undefined, pref: SpellDialect | undefined): Promise<string | null> {
  if (isEnglish(lang) && !/^en[-_]za\b/i.test(lang || '')) return dialectFor(pref, lang);
  const installed = await installedDictionaries();
  const id = dictionaryFor(lang, installed);
  if (id) return 'dict:' + id;
  return isEnglish(lang) ? dialectFor(pref, lang) : null;
}

// --------------------------------------------------------------------------------------------
// Checkers

export interface SpellChecker {
  /** The words (of those given) that are not in the dictionary. */
  check(words: string[], dict: string): Promise<Set<string>>;
  suggest(word: string, dict: string): Promise<string[]>;
  add?(word: string): Promise<void>;
}

/** The background worker checker. */
export const workerChecker: SpellChecker = {
  async check(words, dict) {
    if (dict.startsWith('dict:') && !(await ensureLoaded(dict.slice(5)))) throw new Error('The dictionary for this language could not be loaded.');
    const out = new Set<string>();
    for (let i = 0; i < words.length; i += 2000) {
      const r = await call({ type: 'check', words: words.slice(i, i + 2000), dialect: dict });
      if (r.failed) throw new Error('The spelling checker could not start on this device.');
      if (r.error) throw new Error(r.error);
      for (const w of r.bad || []) out.add(w);
    }
    return out;
  },
  async suggest(word, dict) { return (await call({ type: 'suggest', word, dialect: dict })).suggestions || []; },
  async add(word) { await call({ type: 'add', word }); },
};

export async function suggest(word: string, dialect?: string): Promise<string[]> { return workerChecker.suggest(word, dialect || 'any'); }
export async function addWord(word: string): Promise<void> { await workerChecker.add!(word); }

// --------------------------------------------------------------------------------------------
// Words in text

const WORD_RE = /[\p{L}][\p{L}\p{M}'’]*[\p{L}\p{M}]|[\p{L}]/gu;
const URL_RE = /\b(?:https?:\/\/|www\.)\S+|\S+@\S+\.\w+|\b[\w-]+(?:\.[\w-]+)+\/\S*/gi;

export interface SpellOptions {
  ignoreUpper: boolean;
  ignoreDigits: boolean;
  ignoreUrls: boolean;
}
export const DEFAULT_SPELL_OPTIONS: SpellOptions = { ignoreUpper: true, ignoreDigits: true, ignoreUrls: true };

export interface WordAt { word: string; start: number; end: number }

/** The checkable words of a text (offsets into the text). */
export function wordsIn(text: string, opts: SpellOptions = DEFAULT_SPELL_OPTIONS): WordAt[] {
  const skip: [number, number][] = [];
  if (opts.ignoreUrls) { URL_RE.lastIndex = 0; let u: RegExpExecArray | null; while ((u = URL_RE.exec(text))) skip.push([u.index, u.index + u[0].length]); }
  const out: WordAt[] = [];
  WORD_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = WORD_RE.exec(text))) {
    const start = m.index, end = start + m[0].length;
    if (skip.some(([a, b]) => start >= a && start < b)) continue;
    // Words glued to digits ("A4", "2nd", "Q3x") are codes, not words.
    const before = text[start - 1] || '', after = text[end] || '';
    if (opts.ignoreDigits && (/\d/.test(before) || /\d/.test(after))) continue;
    let w = m[0].replace(/’/g, "'");
    let s = start, e = end;
    if (/'s$/i.test(w)) { w = w.slice(0, -2); e -= 2; }
    while (w.startsWith("'")) { w = w.slice(1); s++; }
    while (w.endsWith("'")) { w = w.slice(0, -1); e--; }
    if (w.length < 2) continue;
    if (opts.ignoreUpper && /^[\p{Lu}]{2,}$/u.test(w)) continue;
    out.push({ word: w, start: s, end: e });
  }
  return out;
}

// --------------------------------------------------------------------------------------------
// Cells to check

export interface SpellTarget { sheet: Sheet; r: number; c: number }

/** Text of a checkable cell (a text constant, not a formula), or null. */
export function cellSpellText(cell: Cell | undefined): string | null {
  if (!cell || cell.f != null || cell.sp != null) return null;
  return typeof cell.v === 'string' && /\p{L}/u.test(cell.v) ? cell.v : null;
}

/**
 * Text cells of a sheet in reading order (row by row), limited to the ranges when given, starting
 * at (fromR, fromC) when given (cells before it come last when wrap is true, else are left out).
 */
export function spellTargets(sheet: Sheet, ranges: Range[] | null, from?: { r: number; c: number }, wrap = true): SpellTarget[] {
  const keys: number[] = [];
  const seen = new Set<number>();
  for (const [k, cell] of sheet.cells) {
    if (cellSpellText(cell) == null) continue;
    const r = keyRow(k), c = keyCol(k);
    if (ranges && !ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2)) continue;
    if (seen.has(k)) continue;
    seen.add(k);
    keys.push(k);
  }
  keys.sort((a, b) => a - b);
  let list = keys.map((k) => ({ sheet, r: keyRow(k), c: keyCol(k) }));
  if (from) {
    const i = list.findIndex((t) => t.r > from.r || (t.r === from.r && t.c >= from.c));
    if (i > 0) list = wrap ? [...list.slice(i), ...list.slice(0, i)] : list.slice(i);
    else if (i < 0 && !wrap) list = [];
  }
  return list;
}

/** Replaces text[start, end) in a cell, keeping rich text runs when the change lies inside one run. */
export function replaceInCell(cell: Cell, start: number, end: number, text: string): Partial<Cell> {
  const v = String(cell.v ?? '');
  const next = v.slice(0, start) + text + v.slice(end);
  if (!cell.rich) return { v: next };
  let at = 0;
  const runs = cell.rich.map((r) => ({ ...r }));
  for (const run of runs) {
    const a = at, b = at + run.t.length;
    if (start >= a && end <= b) {
      run.t = run.t.slice(0, start - a) + text + run.t.slice(end - a);
      return { v: next, rich: runs };
    }
    at = b;
  }
  return { v: next, rich: undefined };
}

// --------------------------------------------------------------------------------------------
// The spelling walk

export interface SpellIssue { target: SpellTarget; word: string; start: number; end: number; text: string }

export interface SpellSessionOptions {
  dict: string;
  /** Words the user ignores (lowercase). Shared with preferences: Ignore All adds to it. */
  ignored: Set<string>;
  /** Personal dictionary (lowercase). Add to Dictionary adds to it. */
  dictionary: Set<string>;
  options?: SpellOptions;
  /** Writes a change into a cell (inside a transaction); returns false when the cell cannot be changed. */
  write: (t: SpellTarget, patch: Partial<Cell>) => boolean;
}

/**
 * Walks the targets one misspelling at a time. next() finds the next issue (applying Change All
 * replacements on the way); the action methods resolve the current issue and move on.
 */
export class SpellSession {
  private i = 0;
  private pos = 0;
  private bad = new Set<string>();
  private known = new Set<string>();
  private changeAllMap = new Map<string, string>();
  current: SpellIssue | null = null;
  changed = 0;

  constructor(public targets: SpellTarget[], private checker: SpellChecker, private o: SpellSessionOptions) {}

  private text(t: SpellTarget): string | null { return cellSpellText(t.sheet.get(t.r, t.c)); }

  private async prime(words: string[]): Promise<void> {
    const fresh = [...new Set(words.filter((w) => !this.known.has(w)))];
    if (!fresh.length) return;
    const bad = await this.checker.check(fresh, this.o.dict);
    for (const w of fresh) { this.known.add(w); if (bad.has(w)) this.bad.add(w); }
  }

  /** Checks every word of the remaining targets in one batch (faster than cell by cell). */
  async start(): Promise<void> {
    const words: string[] = [];
    for (const t of this.targets) { const tx = this.text(t); if (tx) for (const w of wordsIn(tx, this.o.options)) words.push(w.word); }
    await this.prime(words);
  }

  private skip(w: string): boolean {
    const l = w.toLowerCase();
    return this.o.ignored.has(l) || this.o.dictionary.has(l);
  }

  /** Adds targets to the end of the walk (another sheet, or the start of this one). */
  extend(targets: SpellTarget[]): void { this.targets.push(...targets); }

  async next(): Promise<SpellIssue | null> {
    while (this.i < this.targets.length) {
      const t = this.targets[this.i];
      const text = this.text(t);
      if (text != null) {
        const words = wordsIn(text, this.o.options).filter((w) => w.start >= this.pos);
        await this.prime(words.map((w) => w.word));
        let restart = false;
        for (const w of words) {
          if (!this.bad.has(w.word) || this.skip(w.word)) continue;
          const auto = this.changeAllMap.get(w.word);
          if (auto != null) {
            if (this.apply(t, w.start, w.end, auto)) { this.pos = w.start + auto.length; restart = true; break; }
            continue;
          }
          this.current = { target: t, word: w.word, start: w.start, end: w.end, text };
          return this.current;
        }
        if (restart) continue;
      }
      this.i++;
      this.pos = 0;
    }
    this.current = null;
    return null;
  }

  private apply(t: SpellTarget, start: number, end: number, rep: string): boolean {
    const cell = t.sheet.get(t.r, t.c);
    if (!cell) return false;
    const ok = this.o.write(t, replaceInCell(cell, start, end, rep));
    if (ok) this.changed++;
    return ok;
  }

  ignoreOnce(): void { if (this.current) this.pos = this.current.end; }
  ignoreAll(): void { if (this.current) { this.o.ignored.add(this.current.word.toLowerCase()); this.pos = this.current.end; } }
  async addToDictionary(): Promise<void> {
    if (!this.current) return;
    this.o.dictionary.add(this.current.word.toLowerCase());
    await this.checker.add?.(this.current.word);
    this.pos = this.current.end;
  }
  change(rep: string): boolean {
    const c = this.current;
    if (!c) return false;
    if (!this.apply(c.target, c.start, c.end, rep)) { this.pos = c.end; return false; }
    this.pos = c.start + rep.length;
    return true;
  }
  changeAll(rep: string): boolean {
    if (!this.current) return false;
    this.changeAllMap.set(this.current.word, rep);
    return this.change(rep);
  }
  /** Moves past the current cell (after an edit typed into the Not in Dictionary box). */
  replaceCellText(text: string): boolean {
    const c = this.current;
    if (!c) return false;
    const cell = c.target.sheet.get(c.target.r, c.target.c);
    if (!cell) return false;
    const ok = this.o.write(c.target, { v: text, rich: undefined });
    if (ok) this.changed++;
    this.pos = 0;
    return ok;
  }
}
