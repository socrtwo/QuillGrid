/**
 * Review tab: spelling, dictionaries, thesaurus, workbook statistics, accessibility, smart lookup,
 * translate, threaded comments, notes, sheet/workbook protection, allow edit ranges, mark as final,
 * and ink visibility.
 */
import { def, type MenuEntry } from './registry';
import type { App } from '../app';
import type { Workbook, Sheet } from '../model/workbook';
import type { CellComment, Cell, AllowEditRange } from '../model/types';
import { cellKey, keyRow, keyCol, cellName, rangeName, type Range } from '../model/address';
import { resolveRangeText, renameSheet, validSheetName, rangeRefText } from '../model/structure';
import { selRanges, activeCell, clipToUsed, cellText } from './ops';
import { openDialog, toast, confirmDialog, promptDialog, alertDialog, choiceDialog, type Field } from '../ui/dialog';
import { h, icon, clear } from '../ui/dom';
import { registerPane, typingIn } from '../ui/panes';
import {
  SpellSession, spellTargets, workerChecker, dictionaryKey, SPELL_DIALECTS, DEFAULT_SPELL_OPTIONS, isEnglish,
  type SpellOptions, type SpellTarget, type SpellIssue, type SpellDialect,
} from '../proof/proofing';

const RV = 'Review';

// --------------------------------------------------------------------------------------------
// Passwords (stored as "ALGORITHM:salt:spinCount:hash", the spreadsheet file format's scheme)

const toB64 = (b: Uint8Array) => { let s = ''; for (const x of b) s += String.fromCharCode(x); return btoa(s); };
const fromB64 = (s: string) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
const utf16le = (s: string) => { const out = new Uint8Array(s.length * 2); for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); out[i * 2] = c & 0xff; out[i * 2 + 1] = c >> 8; } return out; };
const concat = (a: Uint8Array, b: Uint8Array) => { const o = new Uint8Array(a.length + b.length); o.set(a); o.set(b, a.length); return o; };

/** Default iteration count for new password hashes (the spreadsheet format's default). */
export const PASSWORD_SPIN = 100000;

const K256 = new Uint32Array([0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5, 0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174, 0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da, 0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967, 0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85, 0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070, 0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3, 0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2]);
const W256 = new Uint32Array(64);

/** SHA-256 (synchronous, so the many hashing rounds of a password stay fast). */
export function sha256(data: Uint8Array): Uint8Array {
  const H = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
  const len = data.length;
  const padded = new Uint8Array(((len + 9 + 63) >> 6) << 6);
  padded.set(data);
  padded[len] = 0x80;
  const bits = len * 8;
  padded[padded.length - 5] = Math.floor(bits / 0x100000000) & 0xff;
  for (let i = 0; i < 4; i++) padded[padded.length - 1 - i] = (bits >>> (8 * i)) & 0xff;
  const W = W256;
  for (let off = 0; off < padded.length; off += 64) {
    for (let i = 0; i < 16; i++) { const j = off + i * 4; W[i] = (padded[j] << 24) | (padded[j + 1] << 16) | (padded[j + 2] << 8) | padded[j + 3]; }
    for (let i = 16; i < 64; i++) {
      const a = W[i - 15], b = W[i - 2];
      const s0 = ((a >>> 7) | (a << 25)) ^ ((a >>> 18) | (a << 14)) ^ (a >>> 3);
      const s1 = ((b >>> 17) | (b << 15)) ^ ((b >>> 19) | (b << 13)) ^ (b >>> 10);
      W[i] = (W[i - 16] + s0 + W[i - 7] + s1) | 0;
    }
    let a = H[0], b = H[1], c = H[2], d = H[3], e = H[4], f = H[5], g = H[6], hh = H[7];
    for (let i = 0; i < 64; i++) {
      const S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
      const t1 = (hh + S1 + ((e & f) ^ (~e & g)) + K256[i] + W[i]) | 0;
      const S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
      const t2 = (S0 + ((a & b) ^ (a & c) ^ (b & c))) | 0;
      hh = g; g = f; f = e; e = (d + t1) | 0; d = c; c = b; b = a; a = (t1 + t2) | 0;
    }
    H[0] = (H[0] + a) | 0; H[1] = (H[1] + b) | 0; H[2] = (H[2] + c) | 0; H[3] = (H[3] + d) | 0; H[4] = (H[4] + e) | 0; H[5] = (H[5] + f) | 0; H[6] = (H[6] + g) | 0; H[7] = (H[7] + hh) | 0;
  }
  const out = new Uint8Array(32);
  for (let i = 0; i < 8; i++) { out[i * 4] = H[i] >>> 24; out[i * 4 + 1] = (H[i] >>> 16) & 0xff; out[i * 4 + 2] = (H[i] >>> 8) & 0xff; out[i * 4 + 3] = H[i] & 0xff; }
  return out;
}

async function digest(alg: string, data: Uint8Array): Promise<Uint8Array> {
  if (alg === 'SHA-256') return sha256(data);
  const name = alg === 'SHA-512' ? 'SHA-512' : alg === 'SHA-384' ? 'SHA-384' : alg === 'SHA-1' ? 'SHA-1' : '';
  if (!name || !globalThis.crypto?.subtle) throw new Error(`Unsupported password algorithm ${alg}`);
  return new Uint8Array(await crypto.subtle.digest(name, data as BufferSource));
}

/** Iterated, salted hash: H0 = H(salt + password as UTF-16LE), Hn = H(Hn-1 + n as 32-bit little-endian). */
async function iteratedHash(alg: string, pw: string, salt: Uint8Array, spin: number): Promise<Uint8Array> {
  let hv = await digest(alg, concat(salt, utf16le(pw)));
  const it = new Uint8Array(4);
  const sync = alg === 'SHA-256';
  for (let i = 0; i < spin; i++) {
    it[0] = i & 0xff; it[1] = (i >>> 8) & 0xff; it[2] = (i >>> 16) & 0xff; it[3] = (i >>> 24) & 0xff;
    hv = sync ? sha256(concat(hv, it)) : await digest(alg, concat(hv, it));
  }
  return hv;
}

/** Hashes a password with SHA-256 and a random salt. The password itself is never stored. */
export async function hashPassword(pw: string, spin = PASSWORD_SPIN, salt?: Uint8Array): Promise<string> {
  const s = salt || globalThis.crypto.getRandomValues(new Uint8Array(16));
  const hv = await iteratedHash('SHA-256', pw, s, spin);
  return `SHA-256:${toB64(s)}:${spin}:${toB64(hv)}`;
}

/** The 16-bit legacy password verifier older spreadsheet files use (4 hex digits). */
export function legacyPasswordHash(pw: string): string {
  let hv = 0;
  for (let i = pw.length - 1; i >= 0; i--) { hv = ((hv >> 14) & 1) | ((hv << 1) & 0x7fff); hv ^= pw.charCodeAt(i) & 0xff; }
  hv = ((hv >> 14) & 1) | ((hv << 1) & 0x7fff);
  hv ^= pw.length;
  hv ^= 0xce4b;
  return hv.toString(16).toUpperCase().padStart(4, '0');
}

/** True when the password matches a stored hash (any supported scheme). An empty hash accepts anything. */
export async function verifyPassword(pw: string, stored: string | undefined): Promise<boolean> {
  if (!stored) return true;
  if (/^[0-9A-F]{4}$/i.test(stored)) return legacyPasswordHash(pw) === stored.toUpperCase();
  const parts = stored.split(':');
  if (parts.length !== 4) return false;
  const [alg, salt, spin, hv] = parts;
  try {
    const got = await iteratedHash(alg.toUpperCase().replace(/^SHA(\d)/, 'SHA-$1'), pw, fromB64(salt), Math.max(0, Math.min(10_000_000, parseInt(spin, 10) || 0)));
    return toB64(got) === hv;
  } catch { return false; }
}

// --------------------------------------------------------------------------------------------
// Helpers

function objectsLocked(app: App): boolean {
  const s = app.sheet;
  if (s.protection.on && !s.protection.allow.includes('objects')) {
    toast('Comments and notes on a protected sheet cannot be changed. Unprotect the sheet first (Review ▸ Unprotect Sheet).');
    return true;
  }
  return false;
}

function target(app: App, arg?: { r?: number; c?: number }): { r: number; c: number } {
  return arg && typeof arg.r === 'number' && typeof arg.c === 'number' ? { r: arg.r, c: arg.c } : activeCell(app);
}

/** Sets or removes the comment/note at a cell as one undoable step. */
export function setComment(book: Workbook, s: Sheet, r: number, c: number, cm: CellComment | undefined, label: string): void {
  book.transact(label, () => {
    book.recordSheet(s);
    const k = cellKey(r, c);
    if (cm) s.comments.set(k, cm); else s.comments.delete(k);
  });
}

async function askText(title: string, label: string, value = '', ok = 'OK'): Promise<string | null> {
  const r = await openDialog({ title, ok, fields: [{ key: 't', label, type: 'textarea', rows: 5, value, autofocus: true }] });
  if (!r) return null;
  const t = String(r.t || '').trim();
  return t || null;
}

function showComment(app: App, r: number, c: number): void {
  app.grid.select(r, c);
  app.grid.scrollIntoView(r, c);
  try { app.grid.objects.hoverComment(r, c); } catch { /* */ }
}

/** Comments (or notes) of the workbook in reading order: sheet by sheet, row by row. */
function commentList(book: Workbook, notes: boolean): { si: number; k: number }[] {
  const out: { si: number; k: number }[] = [];
  book.sheets.forEach((s, si) => {
    if (s.visibility !== 'visible') return;
    for (const [k, cm] of s.comments) if (!!cm.note === notes) out.push({ si, k });
  });
  return out.sort((a, b) => a.si - b.si || keyRow(a.k) - keyRow(b.k) || keyCol(a.k) - keyCol(b.k));
}

function stepComment(app: App, notes: boolean, dir: 1 | -1): void {
  const list = commentList(app.book, notes);
  if (!list.length) { toast(notes ? 'There are no notes in this workbook.' : 'There are no comments in this workbook.'); return; }
  const cur = { si: app.book.activeSheet, k: cellKey(app.sheet.sel.r, app.sheet.sel.c) };
  const cmp = (a: { si: number; k: number }) => a.si - cur.si || keyRow(a.k) - keyRow(cur.k) || keyCol(a.k) - keyCol(cur.k);
  const next = dir > 0 ? list.find((x) => cmp(x) > 0) || list[0] : [...list].reverse().find((x) => cmp(x) < 0) || list[list.length - 1];
  if (next.si !== app.book.activeSheet) app.exec('sheet.activate', next.si);
  showComment(app, keyRow(next.k), keyCol(next.k));
  if (!notes && !app.ui.paneOpen('comments')) app.ui.openPane('comments');
}

/** Text of the active cell for lookups (or the first selected text). */
function lookupText(app: App): string {
  const s = app.sheet;
  const { r, c } = activeCell(app);
  let t = cellText(s, r, c).trim();
  if (!t) {
    for (const rg of selRanges(app)) {
      s.forEachIn(clipToUsed(s, rg), (rr, cc) => { if (!t) t = cellText(s, rr, cc).trim(); });
      if (t) break;
    }
  }
  return t;
}

/** Writes new text to a cell when the sheet allows it (used by spelling and the thesaurus). */
function writeCellText(app: App, s: Sheet, r: number, c: number, patch: Partial<Cell>, label: string): boolean {
  if (app.isReadOnly()) { toast('This workbook is read-only.'); return false; }
  if (s.protection.on && s.styleOf(r, c).locked !== false && !s.editRanges.some((er) => coversCell(app.book, s, er, r, c))) { app.exec('review.protectionNotice'); return false; }
  app.book.transact(label, () => s.patch(r, c, patch));
  return true;
}

function coversCell(book: Workbook, s: Sheet, er: AllowEditRange, r: number, c: number): boolean {
  for (const part of er.ref.split(',')) {
    const x = resolveRangeText(book, part.trim().replace(/^=/, ''), s);
    if (x && x.sheet === s && r >= x.range.r1 && r <= x.range.r2 && c >= x.range.c1 && c <= x.range.c2) return true;
  }
  return false;
}

// --------------------------------------------------------------------------------------------
// Workbook statistics

export interface SheetStats { name: string; cells: number; formulas: number; lastCell: string; tables: number; charts: number; pictures: number; pivots: number; comments: number; notes: number }
export interface WorkbookStats { sheets: number; cells: number; formulas: number; tables: number; charts: number; pictures: number; pivots: number; comments: number; notes: number; names: number; perSheet: SheetStats[] }

export function workbookStats(book: Workbook): WorkbookStats {
  const perSheet = book.sheets.map((s): SheetStats => {
    let cells = 0, formulas = 0, lr = -1, lc = -1;
    for (const [k, cell] of s.cells) {
      if (cell.f != null) formulas++;
      if (cell.f != null || (cell.v != null && cell.v !== '')) { cells++; lr = Math.max(lr, keyRow(k)); lc = Math.max(lc, keyCol(k)); }
    }
    const notes = [...s.comments.values()].filter((c) => c.note).length;
    return {
      name: s.name, cells, formulas, lastCell: lr >= 0 ? cellName(lr, lc) : '—', tables: s.tables.length, charts: s.charts.length,
      pictures: s.drawings.filter((d) => d.kind === 'image').length, pivots: s.pivots.length, comments: s.comments.size - notes, notes,
    };
  });
  const sum = (k: keyof SheetStats) => perSheet.reduce((a, x) => a + (x[k] as number), 0);
  return { sheets: book.sheets.length, cells: sum('cells'), formulas: sum('formulas'), tables: sum('tables'), charts: sum('charts'), pictures: sum('pictures'), pivots: sum('pivots'), comments: sum('comments'), notes: sum('notes'), names: book.names.filter((n) => !n.hidden).length, perSheet };
}

async function showStats(app: App): Promise<void> {
  const st = workbookStats(app.book);
  const cur = st.perSheet[app.book.activeSheet];
  const row = (label: string, a: string | number, b: string | number) => h('tr', null, h('td', null, label), h('td', null, String(a)), h('td', null, String(b)));
  const table = h('table', { class: 'ql-stats' },
    h('tr', null, h('th', null, ''), h('th', null, 'Current sheet'), h('th', null, 'Workbook')),
    row('Sheets', '', st.sheets),
    row('End of sheet (last cell)', cur.lastCell, ''),
    row('Cells with data', cur.cells.toLocaleString(), st.cells.toLocaleString()),
    row('Formulas', cur.formulas.toLocaleString(), st.formulas.toLocaleString()),
    row('Tables', cur.tables, st.tables),
    row('Charts', cur.charts, st.charts),
    row('Pictures', cur.pictures, st.pictures),
    row('Pivot tables', cur.pivots, st.pivots),
    row('Comments', cur.comments, st.comments),
    row('Notes', cur.notes, st.notes),
    row('Defined names', '', st.names));
  const per = h('table', { class: 'ql-stats' }, h('tr', null, h('th', null, 'Sheet'), h('th', null, 'Last cell'), h('th', null, 'Cells'), h('th', null, 'Formulas')),
    ...st.perSheet.map((p) => h('tr', null, h('td', null, p.name), h('td', null, p.lastCell), h('td', null, p.cells.toLocaleString()), h('td', null, p.formulas.toLocaleString()))));
  await openDialog({ title: 'Workbook Statistics', content: h('div', null, table, st.sheets > 1 ? h('h3', { class: 'ql-field-heading' }, 'By sheet') : null, st.sheets > 1 ? per : null), ok: 'OK', cancel: false });
}

// --------------------------------------------------------------------------------------------
// Accessibility

export type A11yFix = 'alt' | 'unmerge' | 'rename' | 'contrast' | 'linkText' | 'tableHeader';
export interface A11yIssue { kind: 'error' | 'warning' | 'tip'; rule: string; title: string; why: string; sheet: Sheet; where: string; r?: number; c?: number; objId?: string; range?: Range; fix?: A11yFix }

function lum(hex: string): number {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) return 1;
  const n = parseInt(m[1], 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => { const s = v / 255; return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4; });
  return 0.2126 * ch[0] + 0.7152 * ch[1] + 0.0722 * ch[2];
}
/** WCAG contrast ratio between two #rrggbb colors. */
export function contrastRatio(a: string, b: string): number {
  const la = lum(a), lb = lum(b);
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

const VAGUE_LINK = /^(click here|click|here|link|this link|more|read more|learn more|this|go|url|website|details)$/i;

/** Accessibility problems in a workbook. */
export function accessibilityIssues(book: Workbook): A11yIssue[] {
  const out: A11yIssue[] = [];
  const LIMIT = 200;
  const defaults = book.sheets.filter((s) => /^Sheet\d+$/i.test(s.name));
  for (const s of book.sheets) {
    if (s.visibility === 'veryHidden') continue;
    for (const ch of s.charts) if (!ch.alt?.trim()) out.push({ kind: 'error', rule: 'alt', title: `Missing alternative text: ${ch.title || ch.type + ' chart'}`, why: 'People who use screen readers need a description of the chart.', sheet: s, where: `${s.name}!${cellName(ch.at.r, ch.at.c)}`, objId: ch.id, r: ch.at.r, c: ch.at.c, fix: 'alt' });
    for (const d of s.drawings) {
      if (d.kind === 'ink') continue;
      if ((d.kind === 'image' || d.kind === 'shape' || d.kind === 'icon' || d.kind === 'signature') && !d.alt?.trim() && !(d.kind === 'shape' && d.text?.trim())) {
        out.push({ kind: 'error', rule: 'alt', title: `Missing alternative text: ${d.kind === 'image' ? 'picture' : d.kind}`, why: 'Describe pictures and shapes so screen readers can announce them.', sheet: s, where: `${s.name}!${cellName(d.at.r, d.at.c)}`, objId: d.id, r: d.at.r, c: d.at.c, fix: 'alt' });
      }
    }
    for (const tb of s.tables) {
      if (!tb.opts.header) out.push({ kind: 'error', rule: 'tableHeader', title: `Table without a header row: ${tb.name}`, why: 'A header row tells screen readers what each column holds.', sheet: s, where: `${s.name}!${rangeName(tb.range)}`, r: tb.range.r1, c: tb.range.c1, range: tb.range, fix: 'tableHeader' });
      const first = tb.range.r1 + (tb.opts.header ? 1 : 0), last = tb.range.r2 - (tb.opts.totals ? 1 : 0);
      for (let r = first; r <= last; r++) {
        let blank = true;
        for (let c = tb.range.c1; c <= tb.range.c2 && blank; c++) { const x = s.get(r, c); if (x && ((x.v != null && x.v !== '') || x.f != null)) blank = false; }
        if (blank) { out.push({ kind: 'warning', rule: 'blankRow', title: `Blank row in table ${tb.name}`, why: 'Blank rows can make screen readers think the table ended.', sheet: s, where: `${s.name}!${r + 1}:${r + 1}`, r, c: tb.range.c1 }); break; }
      }
    }
    for (const m of s.merges) {
      if (out.length > LIMIT * 4) break;
      if (m.r2 === m.r1 && m.c2 === m.c1) continue;
      out.push({ kind: 'warning', rule: 'merged', title: `Merged cells ${rangeName(m)}`, why: 'Merged cells make it hard to navigate a sheet with a screen reader.', sheet: s, where: `${s.name}!${rangeName(m)}`, r: m.r1, c: m.c1, range: m, fix: 'unmerge' });
    }
    let contrast = 0, links = 0;
    for (const [k, cell] of s.cells) {
      if (cell.v == null || cell.v === '') continue;
      const r = keyRow(k), c = keyCol(k);
      const st = book.styles[cell.s ?? 0] || {};
      if (st.fill && contrast < LIMIT) {
        const fg = /^#[0-9a-f]{6}$/i.test(st.color || '') ? st.color! : '#000000';
        const bg = /^#[0-9a-f]{6}$/i.test(st.fill) ? st.fill : null;
        if (bg && contrastRatio(fg, bg) < 4.5) { contrast++; out.push({ kind: 'warning', rule: 'contrast', title: `Hard-to-read text contrast in ${cellName(r, c)}`, why: `The text and fill colors have a contrast ratio of ${contrastRatio(fg, bg).toFixed(1)}:1 (4.5:1 or more is easier to read).`, sheet: s, where: `${s.name}!${cellName(r, c)}`, r, c, fix: 'contrast' }); }
      }
      if (cell.link && links < LIMIT) {
        const text = String(cell.v).trim();
        if (VAGUE_LINK.test(text) || (text === cell.link && text.length > 30)) { links++; out.push({ kind: 'warning', rule: 'link', title: `Unclear hyperlink text in ${cellName(r, c)}`, why: 'Link text should say where the link goes, not "click here" or a long address.', sheet: s, where: `${s.name}!${cellName(r, c)}`, r, c, fix: 'linkText' }); }
      }
    }
    if (defaults.length > 1 && defaults.includes(s) && s.usedRange()) out.push({ kind: 'tip', rule: 'sheetName', title: `Default sheet name: ${s.name}`, why: 'Give sheets names that describe their contents so they are easier to find.', sheet: s, where: s.name, fix: 'rename' });
    if (s.visibility === 'hidden') out.push({ kind: 'tip', rule: 'hidden', title: `Hidden sheet: ${s.name}`, why: 'Content on hidden sheets is easy to miss. Unhide it or delete it if it is not needed.', sheet: s, where: s.name });
  }
  const order = { error: 0, warning: 1, tip: 2 };
  return out.sort((a, b) => order[a.kind] - order[b.kind]);
}

async function fixIssue(app: App, is: A11yIssue): Promise<void> {
  const book = app.book;
  const s = is.sheet;
  if (app.isReadOnly()) { toast('This workbook is read-only.'); return; }
  switch (is.fix) {
    case 'alt': {
      const ch = s.charts.find((x) => x.id === is.objId), d = s.drawings.find((x) => x.id === is.objId);
      const text = await askText('Alternative Text', 'How would you describe this object and its context to someone who is blind or has low vision?', ch?.alt || d?.alt || '', 'Save');
      if (!text) return;
      book.transact('Alt Text', () => { book.recordSheet(s); if (ch) ch.alt = text; if (d) d.alt = text; });
      break;
    }
    case 'unmerge':
      book.transact('Unmerge Cells', () => { book.recordSheet(s); s.merges = s.merges.filter((m) => !(is.range && m.r1 === is.range.r1 && m.c1 === is.range.c1 && m.r2 === is.range.r2 && m.c2 === is.range.c2)); });
      break;
    case 'rename': {
      if (book.structureLocked) { toast('The workbook structure is protected.'); return; }
      const name = await promptDialog('Rename Sheet', 'Sheet name', s.name);
      if (!name || name === s.name) return;
      const err = validSheetName(name);
      if (err) { toast(err); return; }
      if (book.sheets.some((x) => x !== s && x.name.toLowerCase() === name.toLowerCase())) { toast('That name is already taken.'); return; }
      book.transact('Rename Sheet', () => renameSheet(book, s, name));
      break;
    }
    case 'contrast': {
      const cell = s.get(is.r!, is.c!);
      if (!cell) return;
      const st = book.styles[cell.s ?? 0] || {};
      const color = contrastRatio('#000000', st.fill || '#ffffff') >= contrastRatio('#ffffff', st.fill || '#ffffff') ? '#000000' : '#ffffff';
      book.transact('Text Color', () => s.patch(is.r!, is.c!, { s: book.derivedStyle(cell.s ?? 0, { color }) }));
      break;
    }
    case 'linkText': {
      const cell = s.get(is.r!, is.c!);
      if (!cell) return;
      const text = await promptDialog('Hyperlink Text', `Text to display for ${cell.link}`, String(cell.v ?? ''));
      if (!text) return;
      book.transact('Hyperlink Text', () => s.patch(is.r!, is.c!, { v: text, rich: undefined }));
      break;
    }
    case 'tableHeader': {
      const tb = s.tables.find((t) => is.range && t.range.r1 === is.range.r1 && t.range.c1 === is.range.c1);
      if (!tb) return;
      if (tb.range.r1 === 0) { toast('Insert a row above the table first, then turn on its header row.'); return; }
      book.transact('Header Row', () => {
        book.recordSheet(s);
        tb.range = { ...tb.range, r1: tb.range.r1 - 1 };
        tb.opts = { ...tb.opts, header: true };
        tb.columns.forEach((col, i) => { if (!s.get(tb.range.r1, tb.range.c1 + i)?.v) s.set(tb.range.r1, tb.range.c1 + i, { v: col.name || `Column${i + 1}` }); });
      });
      break;
    }
    default: return;
  }
  toast('Fixed');
}

function goToIssue(app: App, is: A11yIssue): void {
  const si = app.book.sheets.indexOf(is.sheet);
  if (si >= 0 && si !== app.book.activeSheet) app.exec('sheet.activate', si);
  if (is.objId) { app.grid.objects.select(is.objId); if (is.r != null) app.grid.scrollIntoView(is.r, is.c || 0); return; }
  if (is.range) { app.grid.select(is.range.r1, is.range.c1, { range: is.range } as any); app.grid.scrollIntoView(is.range.r1, is.range.c1); return; }
  if (is.r != null) { app.grid.select(is.r, is.c || 0); app.grid.scrollIntoView(is.r, is.c || 0); }
}

registerPane({
  id: 'accessibility', title: 'Accessibility',
  render(app, el) { el.dataset.sig = ''; this.update!(app, el, {}); },
  update(app, el) {
    const sig = `${app.tab.id}:${app.book.version}:${app.book.sheets.length}`;
    if (el.dataset.sig === sig || typingIn(el)) return;
    el.dataset.sig = sig;
    clear(el);
    const issues = accessibilityIssues(app.book);
    const n = (k: string) => issues.filter((i) => i.kind === k).length;
    el.appendChild(h('p', { class: 'ql-help' }, issues.length ? `Errors ${n('error')} · Warnings ${n('warning')} · Tips ${n('tip')}` : '✓ No accessibility issues found. People with disabilities should not have trouble reading this workbook.'));
    el.appendChild(h('div', { class: 'ql-row' }, h('button', { class: 'ql-btn', onclick: () => { el.dataset.sig = ''; this.update!(app, el, {}); } }, icon('refresh'), ' Recheck')));
    for (const is of issues.slice(0, 300)) {
      el.appendChild(h('div', { class: `ql-a11y ql-a11y-${is.kind}` },
        h('b', null, is.title),
        h('span', { class: 'ql-help' }, `${is.where} — ${is.why}`),
        h('div', { class: 'ql-row' },
          h('button', { class: 'ql-btn ql-sm', onclick: () => goToIssue(app, is) }, 'Go to'),
          is.fix ? h('button', { class: 'ql-btn ql-sm', onclick: async () => { goToIssue(app, is); await fixIssue(app, is); el.dataset.sig = ''; } }, ({ alt: 'Add Description…', unmerge: 'Unmerge', rename: 'Rename Sheet…', contrast: 'Use Readable Text Color', linkText: 'Edit Link Text…', tableHeader: 'Add Header Row' } as Record<A11yFix, string>)[is.fix]) : null)));
    }
  },
});

// --------------------------------------------------------------------------------------------
// Thesaurus

function replaceWord(text: string, word: string, rep: string): string {
  const i = text.toLowerCase().indexOf(word.toLowerCase());
  if (i < 0) return rep;
  const orig = text.slice(i, i + word.length);
  const cased = orig[0] && orig[0] === orig[0].toUpperCase() && orig[0] !== orig[0].toLowerCase() ? rep[0].toUpperCase() + rep.slice(1) : rep;
  return text.slice(0, i) + cased + text.slice(i + word.length);
}

registerPane({
  id: 'thesaurus', title: 'Thesaurus',
  render(app, el, opts) {
    clear(el);
    const input = h('input', { class: 'ql-input', type: 'search', placeholder: 'Look up a word…', 'aria-label': 'Word to look up', value: opts.word || '' }) as HTMLInputElement;
    const out = h('div', { class: 'ql-thesaurus' });
    const look = async () => {
      const w = input.value.trim();
      clear(out);
      if (!w) { out.append(h('p', { class: 'ql-help' }, 'Type a word, or select a cell with a word, to find synonyms.')); return; }
      out.append(h('p', { class: 'ql-help' }, 'Looking up…'));
      try {
        const [syn, rel, ant] = await Promise.all([
          fetch(`https://api.datamuse.com/words?rel_syn=${encodeURIComponent(w)}&max=40`).then((r) => r.json()),
          fetch(`https://api.datamuse.com/words?ml=${encodeURIComponent(w)}&max=20`).then((r) => r.json()),
          fetch(`https://api.datamuse.com/words?rel_ant=${encodeURIComponent(w)}&max=15`).then((r) => r.json()),
        ]);
        clear(out);
        const group = (title: string, words: { word: string }[]) => {
          if (!words.length) return;
          out.append(h('h3', { class: 'ql-pane-h' }, title));
          out.append(h('div', { class: 'ql-chips' }, ...words.map((x) => h('span', { class: 'ql-thes-item' },
            h('button', { class: 'ql-chipbtn', title: `Look up "${x.word}"`, onclick: () => { input.value = x.word; void look(); } }, x.word),
            h('button', { class: 'ql-iconbtn ql-sm', title: `Insert "${x.word}"`, 'aria-label': `Insert ${x.word}`, onclick: () => app.exec('review.thesaurusInsert', { word: w, rep: x.word }) }, icon('input'))))));
        };
        const synW = (syn as { word: string }[]).slice(0, 40);
        group('Synonyms', synW);
        group('Related', (rel as { word: string }[]).filter((x) => !synW.some((s) => s.word === x.word) && x.word !== w).slice(0, 20));
        group('Antonyms', ant as { word: string }[]);
        if (!out.childElementCount) out.append(h('p', { class: 'ql-help' }, `No results for "${w}".`));
        out.append(h('p', { class: 'ql-help ql-small' }, 'Results from the Datamuse service. Insert replaces the word in the active cell.'));
      } catch {
        clear(out);
        out.append(h('p', { class: 'ql-help' }, 'The thesaurus needs an internet connection.'));
      }
    };
    input.addEventListener('keydown', (e) => { if (e.key === 'Enter') void look(); });
    el.append(h('div', { class: 'ql-row' }, input, h('button', { class: 'ql-btn', onclick: () => void look() }, icon('search'))), out);
    void look();
  },
  onOpen(app, el, opts) { const inp = el.querySelector<HTMLInputElement>('input'); if (inp && opts.word && inp.value !== opts.word) { inp.value = opts.word; inp.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' })); } void app; },
});

// --------------------------------------------------------------------------------------------
// Spelling

const OPTS_KEY = 'quilllattice.spellopts.v1';
export function loadSpellOptions(): SpellOptions {
  try { return { ...DEFAULT_SPELL_OPTIONS, ...JSON.parse(localStorage.getItem(OPTS_KEY) || '{}') }; } catch { return { ...DEFAULT_SPELL_OPTIONS }; }
}
function saveSpellOptions(o: SpellOptions): void { try { localStorage.setItem(OPTS_KEY, JSON.stringify(o)); } catch { /* */ } }

async function spellingOptions(app: App): Promise<boolean> {
  const o = loadSpellOptions();
  const r = await openDialog({
    title: 'Spelling Options',
    fields: [
      { key: 'dialect', label: 'English dictionary', type: 'select', options: SPELL_DIALECTS, value: app.prefs.spellDialect || 'any' },
      { key: 'language', label: 'Spelling language (language tag)', type: 'text', value: app.prefs.language || 'en-US', help: 'For example en-US, fr-FR or de-DE. Other languages need a downloaded dictionary (Review ▸ Spelling Dictionaries).' },
      { key: 'ignoreUpper', label: 'Ignore words in UPPERCASE', type: 'checkbox', value: o.ignoreUpper },
      { key: 'ignoreDigits', label: 'Ignore words that contain numbers', type: 'checkbox', value: o.ignoreDigits },
      { key: 'ignoreUrls', label: 'Ignore internet and file addresses', type: 'checkbox', value: o.ignoreUrls },
      { key: 'dictionary', label: 'Custom dictionary (one word per line)', type: 'textarea', rows: 5, value: app.prefs.dictionary.join('\n') },
      { key: 'ignored', label: 'Ignored words (one per line)', type: 'textarea', rows: 3, value: app.prefs.ignoredWords.join('\n') },
    ],
    buttons: [{ label: 'Dictionaries…', onClick: () => void app.exec('review.dictionaries') }],
  });
  if (!r) return false;
  app.prefs.spellDialect = r.dialect as SpellDialect;
  app.prefs.language = String(r.language || 'en-US').trim() || 'en-US';
  app.prefs.dictionary = String(r.dictionary || '').split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
  app.prefs.ignoredWords = String(r.ignored || '').split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
  app.savePrefs();
  saveSpellOptions({ ignoreUpper: !!r.ignoreUpper, ignoreDigits: !!r.ignoreDigits, ignoreUrls: !!r.ignoreUrls });
  return true;
}

/** Cells the spelling check walks: the selection when it has more than one cell, else the sheet from the active cell. */
function spellScope(app: App): { ranges: Range[] | null; from: { r: number; c: number } | null } {
  const s = app.sheet;
  const ranges = selRanges(app);
  const multi = ranges.length > 1 || ranges.some((rg) => rg.r2 > rg.r1 || rg.c2 > rg.c1);
  return multi ? { ranges: ranges.map((rg) => clipToUsed(s, rg)), from: null } : { ranges: null, from: activeCell(app) };
}

async function runSpelling(app: App): Promise<void> {
  if (app.grid.editing) app.grid.commitEdit();
  let dict = await dictionaryKey(app.prefs.language, app.prefs.spellDialect);
  if (!dict) {
    if (await confirmDialog('Spelling', `There is no spelling dictionary for "${app.prefs.language}" on this device. Download one now?`, 'Dictionaries…')) await app.exec('review.dictionaries');
    return;
  }
  const book = app.book;
  const startSheet = app.sheet;
  const scope = spellScope(app);
  const startedAt = scope.from;
  const opts = loadSpellOptions();
  const write = (t: SpellTarget, patch: Partial<Cell>) => writeCellText(app, t.sheet, t.r, t.c, patch, 'Spelling');
  const session = new SpellSession(spellTargets(startSheet, scope.ranges, startedAt || undefined, false), workerChecker, {
    dict, options: opts, write,
    ignored: new Set(app.prefs.ignoredWords.map((w) => w.toLowerCase())),
    dictionary: new Set(app.prefs.dictionary.map((w) => w.toLowerCase())),
  });
  let wrapped = !startedAt || (startedAt.r === 0 && startedAt.c === 0);
  let otherSheetsAsked = !!scope.ranges || book.sheets.length < 2;

  // Dialog UI
  const word = h('input', { class: 'ql-input', 'aria-label': 'Not in dictionary' }) as HTMLInputElement;
  const context = h('div', { class: 'ql-spell-context', 'aria-live': 'polite' });
  const where = h('div', { class: 'ql-help' });
  const sugg = h('select', { class: 'ql-listbox', size: '6', 'aria-label': 'Suggestions' }) as HTMLSelectElement;
  const langSel = h('select', { 'aria-label': 'Dictionary language' }) as HTMLSelectElement;
  for (const d of SPELL_DIALECTS) langSel.appendChild(h('option', { value: d.value }, d.label));
  if (!isEnglish(app.prefs.language)) langSel.appendChild(h('option', { value: 'lang' }, `Spelling language: ${app.prefs.language}`));
  langSel.value = isEnglish(app.prefs.language) ? app.prefs.spellDialect || 'any' : 'lang';
  const status = h('div', { class: 'ql-help' });
  let busy = false;
  let closeDialog: (() => void) | null = null;
  let finished = false;

  const btn = (label: string, fn: () => unknown, primary = false) => {
    const b = h('button', { class: 'ql-btn' + (primary ? ' ql-primary' : ''), type: 'button' }, label);
    b.addEventListener('click', async () => { if (busy || !session.current) return; busy = true; try { await fn(); } finally { busy = false; } await advance(); });
    return b;
  };
  const replacement = () => (word.value.trim() !== session.current?.word ? word.value.trim() : sugg.value || '');
  const changeOne = (all: boolean) => {
    const rep = replacement();
    if (!rep) { toast('Choose a suggestion or type the correct spelling.'); return; }
    if (all) session.changeAll(rep); else session.change(rep);
  };
  const buttons = h('div', { class: 'ql-spell-buttons' },
    btn('Ignore Once', () => session.ignoreOnce()),
    btn('Ignore All', () => { const w = session.current!.word; session.ignoreAll(); if (!app.prefs.ignoredWords.includes(w)) { app.prefs.ignoredWords.push(w); app.savePrefs(); } }),
    btn('Add to Dictionary', async () => { const w = session.current!.word; await session.addToDictionary(); if (!app.prefs.dictionary.includes(w)) { app.prefs.dictionary.push(w); app.savePrefs(); } }),
    btn('Change', () => changeOne(false), true),
    btn('Change All', () => changeOne(true)),
    btn('AutoCorrect', () => { const rep = replacement(); if (!rep) { toast('Choose a suggestion first.'); return; } app.prefs.replacements[session.current!.word] = rep; app.savePrefs(); session.change(rep); toast(`AutoCorrect will change "${session.current?.word || ''}" to "${rep}" as you type.`); }),
  );
  sugg.addEventListener('dblclick', () => (buttons.children[3] as HTMLButtonElement).click());
  const optionsBtn = h('button', { class: 'ql-btn', type: 'button' }, 'Options…');
  optionsBtn.addEventListener('click', async () => {
    if (await spellingOptions(app)) {
      const nd = await dictionaryKey(app.prefs.language, app.prefs.spellDialect);
      if (nd && nd !== dict) toast('The new dictionary is used the next time you check spelling.');
    }
  });
  langSel.addEventListener('change', async () => {
    if (langSel.value === 'lang') return;
    app.prefs.spellDialect = langSel.value as SpellDialect;
    if (!isEnglish(app.prefs.language)) app.prefs.language = 'en-US';
    app.savePrefs();
    toast('Dictionary changed. Close and check spelling again to use it for the whole sheet.');
  });

  const content = h('div', { class: 'ql-spell' },
    where,
    h('label', { class: 'ql-label' }, 'Not in Dictionary:'), word, context,
    h('label', { class: 'ql-label' }, 'Suggestions:'), sugg,
    buttons,
    h('div', { class: 'ql-row' }, h('label', { class: 'ql-label' }, 'Dictionary language: '), langSel, optionsBtn),
    status);

  const showIssue = async (is: SpellIssue) => {
    const si = book.sheets.indexOf(is.target.sheet);
    if (si >= 0 && si !== book.activeSheet) app.exec('sheet.activate', si);
    app.grid.select(is.target.r, is.target.c);
    app.grid.scrollIntoView(is.target.r, is.target.c);
    where.textContent = `${is.target.sheet.name}!${cellName(is.target.r, is.target.c)}`;
    word.value = is.word;
    clear(context);
    context.append(is.text.slice(Math.max(0, is.start - 60), is.start), h('mark', null, is.text.slice(is.start, is.end)), is.text.slice(is.end, is.end + 60));
    clear(sugg);
    sugg.appendChild(h('option', { value: '', disabled: true }, '(looking…)'));
    let list: string[] = [];
    try { list = await workerChecker.suggest(is.word, dict!); } catch { /* */ }
    if (session.current !== is) return;
    clear(sugg);
    if (!list.length) sugg.appendChild(h('option', { value: '', disabled: true }, '(no suggestions)'));
    for (const sgg of list) sugg.appendChild(h('option', { value: sgg }, sgg));
    if (list.length) sugg.selectedIndex = 0;
  };

  const advance = async (): Promise<void> => {
    status.textContent = 'Checking…';
    for (;;) {
      let is: SpellIssue | null;
      try { is = await session.next(); } catch (e) { status.textContent = ''; finished = true; closeDialog?.(); toast(`Spelling: ${(e as Error).message}`); return; }
      if (is) { status.textContent = session.changed ? `${session.changed} change${session.changed > 1 ? 's' : ''} made` : ''; await showIssue(is); return; }
      if (!wrapped && startedAt) {
        wrapped = true;
        const before = spellTargets(startSheet, null).filter((t) => t.r < startedAt.r || (t.r === startedAt.r && t.c < startedAt.c));
        if (before.length && (await confirmDialog('Spelling', 'Do you want to continue checking at the beginning of the sheet?', 'Yes', 'No'))) { session.extend(before); continue; }
      }
      if (!otherSheetsAsked) {
        otherSheetsAsked = true;
        const others = book.sheets.filter((s) => s !== startSheet && s.visibility === 'visible');
        const more = others.flatMap((s) => spellTargets(s, null));
        if (more.length && (await confirmDialog('Spelling', 'The active sheet is done. Check the other sheets too?', 'Yes', 'No'))) { session.extend(more); continue; }
      }
      finished = true;
      closeDialog?.();
      const n = session.changed;
      await alertDialog('Spelling', `Spelling check complete.${n ? ` ${n} change${n > 1 ? 's were' : ' was'} made.` : " You're good to go!"}`);
      return;
    }
  };

  try { await session.start(); } catch (e) { toast(`Spelling: ${(e as Error).message}`); return; }
  const first = await session.next().catch((e) => { toast(`Spelling: ${(e as Error).message}`); return undefined; });
  if (first === undefined) return;
  if (!first) {
    // Nothing from the active cell on; let advance() handle wrapping and other sheets.
    await advance();
    if (finished) return;
  }
  const dlg = openDialog({ title: `Spelling: ${langSel.selectedOptions[0]?.textContent?.replace(/ —.*$/, '') || 'English'}`, content, ok: false, cancel: 'Close', onOpen: (ctl) => { closeDialog = () => ctl.close(); } });
  if (session.current) await showIssue(session.current);
  await dlg;
  if (!finished && session.changed) toast(`${session.changed} spelling change${session.changed > 1 ? 's' : ''} made.`);
}

// --------------------------------------------------------------------------------------------
// Protection

export const PROTECT_OPTIONS: { key: string; label: string }[] = [
  { key: 'selectLocked', label: 'Select locked cells' }, { key: 'selectUnlocked', label: 'Select unlocked cells' },
  { key: 'formatCells', label: 'Format cells' }, { key: 'formatColumns', label: 'Format columns' }, { key: 'formatRows', label: 'Format rows' },
  { key: 'insertColumns', label: 'Insert columns' }, { key: 'insertRows', label: 'Insert rows' }, { key: 'insertHyperlinks', label: 'Insert hyperlinks' },
  { key: 'deleteColumns', label: 'Delete columns' }, { key: 'deleteRows', label: 'Delete rows' }, { key: 'sort', label: 'Sort' },
  { key: 'autoFilter', label: 'Use AutoFilter' }, { key: 'pivotTables', label: 'Use pivot tables and pivot charts' }, { key: 'objects', label: 'Edit objects (charts, pictures, comments)' }, { key: 'scenarios', label: 'Edit scenarios' },
];

async function askNewPassword(title: string, pw: string): Promise<string | null | false> {
  if (!pw) return null;
  const again = await promptDialog(title, 'Reenter the password to proceed. Passwords cannot be recovered if they are lost or forgotten; keep a list of passwords and their sheets in a safe place.', '', { type: 'password' });
  if (again == null) return false;
  if (again !== pw) { await alertDialog(title, 'The confirmation password is not identical.'); return false; }
  return hashPassword(pw);
}

async function protectSheet(app: App): Promise<void> {
  const s = app.sheet;
  const book = app.book;
  if (s.protection.on) {
    if (s.protection.passwordHash) {
      const pw = await promptDialog('Unprotect Sheet', 'Password', '', { type: 'password' });
      if (pw == null) return;
      if (!(await verifyPassword(pw, s.protection.passwordHash))) { await alertDialog('Unprotect Sheet', 'The password you supplied is not correct. Verify that the CAPS LOCK key is off and be sure to use the correct capitalization.'); return; }
    }
    book.transact('Unprotect Sheet', () => { book.recordSheet(s); s.protection = { on: false, allow: s.protection.allow }; });
    toast('Sheet unprotected');
    return;
  }
  const allow = new Set(s.protection.allow?.length ? s.protection.allow : ['selectLocked', 'selectUnlocked']);
  const fields: Field[] = [
    { key: 'on', label: 'Protect worksheet and contents of locked cells', type: 'checkbox', value: true },
    { key: 'pw', label: 'Password to unprotect sheet (optional)', type: 'password', value: '' },
    { key: 'h', label: 'Allow all users of this worksheet to:', type: 'heading' },
    ...PROTECT_OPTIONS.map((o) => ({ key: 'a_' + o.key, label: o.label, type: 'checkbox' as const, value: allow.has(o.key), half: true })),
  ];
  const r = await openDialog({ title: 'Protect Sheet', fields, help: 'Locked cells (all cells unless you unlock them in Format Cells ▸ Protection) cannot be changed while the sheet is protected.' });
  if (!r || !r.on) return;
  const hash = await askNewPassword('Confirm Password', String(r.pw || ''));
  if (hash === false) return;
  const list = PROTECT_OPTIONS.filter((o) => r['a_' + o.key]).map((o) => o.key);
  book.transact('Protect Sheet', () => { book.recordSheet(s); s.protection = { on: true, allow: list, ...(hash ? { passwordHash: hash } : {}) }; });
  toast('Sheet protected');
}

async function protectWorkbook(app: App): Promise<void> {
  const book = app.book;
  if (book.structureLocked) {
    if (book.structurePassword) {
      const pw = await promptDialog('Unprotect Workbook', 'Password', '', { type: 'password' });
      if (pw == null) return;
      if (!(await verifyPassword(pw, book.structurePassword))) { await alertDialog('Unprotect Workbook', 'The password you supplied is not correct.'); return; }
    }
    book.transact('Unprotect Workbook', () => { book.recordBook(); book.structureLocked = false; book.structurePassword = undefined; });
    toast('Workbook structure unprotected');
    return;
  }
  const r = await openDialog({
    title: 'Protect Structure and Windows',
    help: 'Protecting the structure prevents adding, deleting, renaming, moving, copying, hiding and unhiding sheets.',
    fields: [{ key: 'structure', label: 'Structure', type: 'checkbox', value: true }, { key: 'pw', label: 'Password (optional)', type: 'password', value: '' }],
  });
  if (!r || !r.structure) return;
  const hash = await askNewPassword('Confirm Password', String(r.pw || ''));
  if (hash === false) return;
  book.transact('Protect Workbook', () => { book.recordBook(); book.structureLocked = true; book.structurePassword = hash || undefined; });
  toast('Workbook structure protected');
}

async function editRangeDialog(app: App, er?: AllowEditRange): Promise<AllowEditRange | null> {
  const s = app.sheet;
  const r = await openDialog({
    title: er ? 'Modify Range' : 'New Range',
    fields: [
      { key: 'title', label: 'Title', type: 'text', value: er?.title || `Range${s.editRanges.length + 1}` },
      { key: 'ref', label: 'Refers to cells', type: 'text', value: er?.ref || selRanges(app).map((rg) => rangeRefText(s, clipToUsed(s, rg), true, s)).join(',') },
      { key: 'pw', label: er?.passwordHash ? 'Range password (leave blank to keep the current one)' : 'Range password (optional)', type: 'password', value: '' },
      ...(er?.passwordHash ? [{ key: 'clearPw', label: 'Remove the password', type: 'checkbox' as const, value: false }] : []),
    ],
    validate: (v) => {
      if (!String(v.title || '').trim()) return 'Type a title for the range.';
      if (s.editRanges.some((x) => x !== er && x.title.toLowerCase() === String(v.title).trim().toLowerCase())) return 'That title is already used.';
      for (const part of String(v.ref || '').split(',')) { const x = resolveRangeText(app.book, part.trim().replace(/^=/, ''), s); if (!x || x.sheet !== s) return `"${part}" is not a range on this sheet.`; }
      return null;
    },
  });
  if (!r) return null;
  let passwordHash = r.clearPw ? undefined : er?.passwordHash;
  if (r.pw) { const hsh = await askNewPassword('Confirm Password', String(r.pw)); if (hsh === false) return null; passwordHash = hsh || undefined; }
  return { title: String(r.title).trim(), ref: String(r.ref).trim().replace(/^=/, ''), ...(passwordHash ? { passwordHash } : {}) };
}

async function allowEditRanges(app: App): Promise<void> {
  const s = app.sheet;
  const book = app.book;
  for (;;) {
    if (s.protection.on) { toast('Unprotect the sheet to change its editable ranges.'); }
    const opts = s.editRanges.map((er) => ({ value: er.title, label: `${er.title} — ${er.ref}${er.passwordHash ? ' (password)' : ''}` }));
    let action: string | null = null;
    const r = await openDialog({
      title: 'Allow Users to Edit Ranges',
      help: 'Ranges unlocked by a password when the sheet is protected (cells in a range without a password stay editable for everyone).',
      fields: [{ key: 'sel', label: 'Ranges unlocked by a password when sheet is protected', type: 'list', rows: 6, options: opts.length ? opts : [{ value: '', label: '(none)' }] }],
      ok: false, cancel: 'Close',
      buttons: [
        { label: 'New…', onClick: (_v, ctl) => { action = 'new'; ctl.close({}); } },
        { label: 'Modify…', onClick: (v, ctl) => { action = 'modify:' + (v.sel || ''); ctl.close(v); } },
        { label: 'Delete', onClick: (v, ctl) => { action = 'delete:' + (v.sel || ''); ctl.close(v); } },
        { label: s.protection.on ? 'Unprotect Sheet…' : 'Protect Sheet…', onClick: (_v, ctl) => { action = 'protect'; ctl.close({}); } },
      ],
    });
    if (!r || !action) return;
    const act: string = action;
    if (act === 'protect') { await app.exec('review.protectSheet'); return; }
    if (s.protection.on) { toast('Unprotect the sheet first.'); continue; }
    if (act === 'new') {
      const er = await editRangeDialog(app);
      if (er) book.transact('Allow Edit Range', () => { book.recordSheet(s); s.editRanges = [...s.editRanges, er]; });
    } else {
      const [kind, title] = [act.slice(0, act.indexOf(':')), act.slice(act.indexOf(':') + 1)];
      const cur = s.editRanges.find((x) => x.title === title);
      if (!cur) { toast('Select a range first.'); continue; }
      if (kind === 'delete') book.transact('Delete Edit Range', () => { book.recordSheet(s); s.editRanges = s.editRanges.filter((x) => x !== cur); });
      else {
        const er = await editRangeDialog(app, cur);
        if (er) book.transact('Modify Edit Range', () => { book.recordSheet(s); s.editRanges = s.editRanges.map((x) => (x.title === cur.title ? er : x)); });
      }
    }
  }
}

/** Whether a locked cell on a protected sheet may be edited through an Allow Edit Ranges range. */
function canEditRange(app: App, arg: { r: number; c: number }): boolean {
  const s = app.sheet;
  const ranges = s.editRanges.filter((er) => coversCell(app.book, s, er, arg.r, arg.c));
  if (!ranges.length) { app.exec('review.protectionNotice'); return false; }
  const unlocked: Set<string> = ((app.tab as any).unlockedRanges ||= new Set<string>());
  if (ranges.some((er) => !er.passwordHash || unlocked.has(`${s.id}:${er.title}`))) return true;
  const er = ranges[0];
  void (async () => {
    const pw = await promptDialog('Unlock Range', `The cell you are trying to change is password protected. Enter the password for "${er.title}" to change it.`, '', { type: 'password' });
    if (pw == null) return;
    if (await verifyPassword(pw, er.passwordHash)) { unlocked.add(`${s.id}:${er.title}`); toast(`"${er.title}" is unlocked until you close the workbook. Type your change again.`); }
    else await alertDialog('Unlock Range', 'The password you supplied is not correct.');
  })();
  return false;
}

// --------------------------------------------------------------------------------------------
// Commands

function langName(l: string): string { try { return new Intl.DisplayNames(['en'], { type: 'language' }).of(l) || l; } catch { return l; } }

const LANGS = ['es', 'fr', 'de', 'it', 'pt', 'nl', 'sv', 'da', 'no', 'fi', 'pl', 'cs', 'ru', 'uk', 'el', 'tr', 'ar', 'he', 'hi', 'zh-CN', 'zh-TW', 'ja', 'ko', 'vi', 'id', 'la', 'en'];

const NOTES_MENU = (app: App): MenuEntry[] => {
  const { r, c } = activeCell(app);
  const cm = app.sheet.comments.get(cellKey(r, c));
  return [
    cm?.note ? { id: 'review.editNote', label: 'Edit Note', icon: 'edit_note' } : { id: 'review.newNote', label: 'New Note', icon: 'sticky_note_2' },
    { id: 'review.prevNote', label: 'Previous Note', icon: 'arrow_upward' },
    { id: 'review.nextNote', label: 'Next Note', icon: 'arrow_downward' },
    { id: 'review.showNote', label: 'Show/Hide Note', icon: 'visibility' },
    { id: 'review.showAllNotes', label: 'Show All Notes', icon: 'library_books', checked: !!app.grid?.objects?.showAllComments },
    { id: 'review.deleteNote', label: 'Delete Note', icon: 'delete' },
    { separator: true },
    { id: 'review.convertNotes', label: 'Convert to Comments', icon: 'forum' },
  ];
};

def([
  // Proofing
  { id: 'review.spelling', label: 'Spelling', icon: 'spellcheck', tab: RV, group: 'Proofing', keys: ['F7'], popular: true, edits: true, tip: 'Check the spelling of text in the selection or the sheet', run: (app) => runSpelling(app) },
  { id: 'review.spellingOptions', label: 'Spelling Options…', icon: 'tune', tab: RV, group: 'Proofing', noRecord: true, run: (app) => spellingOptions(app) },
  {
    id: 'review.dictionaries', label: 'Spelling Dictionaries…', icon: 'download', tab: RV, group: 'Proofing', noRecord: true,
    tip: 'Download free spelling dictionaries for 90+ languages and choose the spelling language',
    run: async (app) => { const { openDictionaryWizard } = await import('../ui/dictWizard'); await openDictionaryWizard(app); },
  },
  {
    id: 'review.language', label: 'Spelling Language', icon: 'translate', tab: RV, group: 'Language', kind: 'menu', noRecord: true,
    run: (app, v) => {
      if (v === 'more') return app.exec('review.dictionaries');
      if (typeof v !== 'string') return spellingOptions(app);
      app.prefs.spellDialect = v as SpellDialect;
      if (!isEnglish(app.prefs.language)) app.prefs.language = v === 'any' || v === 'document' ? 'en-US' : v;
      app.savePrefs();
      toast(`Spelling: ${SPELL_DIALECTS.find((d) => d.value === v)?.label || v}`);
    },
    menu: (app) => [...SPELL_DIALECTS.map((d) => ({ id: 'review.language', arg: d.value, label: d.label, checked: isEnglish(app.prefs.language) && app.prefs.spellDialect === d.value })), { separator: true }, { id: 'review.language', arg: 'more', label: 'Other Languages…', icon: 'download' }],
  },
  {
    id: 'review.thesaurus', label: 'Thesaurus', icon: 'menu_book', tab: RV, group: 'Proofing', keys: ['Shift-F7'], popular: true, noRecord: true,
    tip: 'Find synonyms for the word in the active cell (online)',
    run: (app) => {
      const t = lookupText(app);
      const words = t.split(/\s+/).filter(Boolean);
      app.ui.openPane('thesaurus', { word: words.length <= 3 ? t : words[0].replace(/[^\p{L}\p{M}'-]/gu, '') });
    },
  },
  {
    id: 'review.thesaurusInsert', label: 'Insert Synonym', icon: 'input', tab: RV, group: 'Proofing', internal: true, edits: true,
    run: (app, arg: { word: string; rep: string }) => {
      if (!arg?.rep) return;
      const s = app.sheet;
      const { r, c } = activeCell(app);
      const cell = s.get(r, c);
      if (cell?.f != null) { toast('The active cell has a formula. Select a cell with text.'); return; }
      const text = typeof cell?.v === 'string' ? cell.v : '';
      writeCellText(app, s, r, c, { v: text ? replaceWord(text, arg.word, arg.rep) : arg.rep, rich: undefined }, 'Insert Synonym');
    },
  },
  { id: 'review.workbookStats', label: 'Workbook Statistics', icon: 'analytics', tab: RV, group: 'Proofing', noRecord: true, tip: 'Counts of sheets, cells, formulas, tables, charts and more', run: (app) => showStats(app) },
  // Accessibility, insights, language
  { id: 'review.accessibility', label: 'Check Accessibility', icon: 'accessibility_new', tab: RV, group: 'Accessibility', noRecord: true, popular: true, run: (app) => app.ui.togglePane('accessibility'), isActive: (app) => app.ui?.paneOpen?.('accessibility') },
  {
    id: 'review.smartLookup', label: 'Smart Lookup', icon: 'travel_explore', tab: RV, group: 'Insights', noRecord: true, tip: 'Search the web for the text in the active cell',
    run: (app) => { const t = lookupText(app); if (!t) { toast('Select a cell with text to look up.'); return; } app.followLink(`https://duckduckgo.com/?q=${encodeURIComponent(t.slice(0, 300))}`); },
  },
  {
    id: 'review.translate', label: 'Translate', icon: 'g_translate', tab: RV, group: 'Language', noRecord: true, tip: 'Translate the selected cells with an online translator',
    run: async (app) => {
      const s = app.sheet;
      const parts: string[] = [];
      for (const rg of selRanges(app)) s.forEachIn(clipToUsed(s, rg), (r, c) => { const t = cellText(s, r, c).trim(); if (t) parts.push(t); });
      const text = parts.join('\n').slice(0, 5000);
      if (!text) { toast('Select cells with text to translate.'); return; }
      const r = await openDialog({ title: 'Translate', help: 'Opens an online translator in your browser. The selected text is sent to that service.', fields: [{ key: 'to', label: 'Translate to', type: 'select', options: LANGS.map((l) => ({ value: l, label: langName(l) })), value: (navigator.language || 'en').split('-')[0] === 'en' ? 'es' : (navigator.language || 'es').split('-')[0] }] });
      if (r) app.followLink(`https://translate.google.com/?sl=auto&tl=${encodeURIComponent(r.to)}&text=${encodeURIComponent(text)}&op=translate`);
    },
  },
  // Comments
  {
    id: 'review.newComment', label: 'New Comment', icon: 'add_comment', tab: RV, group: 'Comments', keys: ['Mod-Alt-m'], popular: true, edits: true,
    run: async (app, arg) => {
      if (objectsLocked(app)) return;
      const s = app.sheet;
      const { r, c } = target(app, arg);
      const cur = s.comments.get(cellKey(r, c));
      if (cur && !cur.note) return app.exec('review.replyComment', { r, c });
      if (cur?.note && !(await confirmDialog('New Comment', 'This cell has a note. A cell can have a note or a comment thread. Convert the note to a comment?', 'Convert'))) return;
      const text = typeof arg?.text === 'string' ? arg.text : await askText(`New Comment — ${cellName(r, c)}`, 'Start a conversation', '', 'Post');
      if (!text) return;
      const replies = cur?.note ? [{ author: app.prefs.author, text, time: Date.now() }] : [];
      setComment(app.book, s, r, c, cur?.note ? { author: cur.author, text: cur.text, time: cur.time, replies } : { author: app.prefs.author, text, time: Date.now(), replies: [] }, 'New Comment');
      if (!app.ui.paneOpen('comments')) app.ui.openPane('comments');
    },
  },
  {
    id: 'review.editComment', label: 'Edit Comment', icon: 'edit_note', tab: RV, group: 'Comments', edits: true,
    run: async (app, arg) => {
      if (objectsLocked(app)) return;
      const s = app.sheet;
      const { r, c } = target(app, arg);
      const cur = s.comments.get(cellKey(r, c));
      if (!cur || cur.note) return app.exec('review.newComment', { r, c });
      const text = typeof arg?.text === 'string' ? arg.text : await askText(`Edit Comment — ${cellName(r, c)}`, 'Comment', cur.text, 'Save');
      if (!text || text === cur.text) return;
      setComment(app.book, s, r, c, { ...cur, text, time: Date.now() }, 'Edit Comment');
    },
  },
  {
    id: 'review.replyComment', label: 'Reply to Comment', icon: 'reply', tab: RV, group: 'Comments', edits: true,
    run: async (app, arg) => {
      if (objectsLocked(app)) return;
      const s = app.sheet;
      const { r, c } = target(app, arg);
      const cur = s.comments.get(cellKey(r, c));
      if (!cur || cur.note) { toast('There is no comment thread in this cell.'); return; }
      const text = typeof arg?.text === 'string' && arg.text.trim() ? arg.text.trim() : await askText(`Reply — ${cellName(r, c)}`, `${cur.author}: ${cur.text.slice(0, 120)}`, '', 'Reply');
      if (!text) return;
      setComment(app.book, s, r, c, { ...cur, resolved: false, replies: [...(cur.replies || []), { author: app.prefs.author, text, time: Date.now() }] }, 'Reply');
    },
  },
  {
    id: 'review.resolveComment', label: 'Resolve Thread', icon: 'task_alt', tab: RV, group: 'Comments', edits: true,
    run: (app, arg) => {
      if (objectsLocked(app)) return;
      const s = app.sheet;
      const { r, c } = target(app, arg);
      const cur = s.comments.get(cellKey(r, c));
      if (!cur || cur.note) { toast('There is no comment thread in this cell.'); return; }
      setComment(app.book, s, r, c, { ...cur, resolved: !cur.resolved }, cur.resolved ? 'Reopen Thread' : 'Resolve Thread');
    },
  },
  {
    id: 'review.deleteComment', label: 'Delete', icon: 'delete_sweep', tab: RV, group: 'Comments', edits: true, tip: 'Delete the comments and notes in the selected cells',
    run: (app, arg) => {
      if (objectsLocked(app)) return;
      const s = app.sheet;
      const book = app.book;
      const keys: number[] = [];
      if (arg && typeof arg.r === 'number') keys.push(cellKey(arg.r, arg.c));
      else for (const k of s.comments.keys()) { const r = keyRow(k), c = keyCol(k); if (selRanges(app).some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2)) keys.push(k); }
      const found = keys.filter((k) => s.comments.has(k));
      if (!found.length) { toast('There are no comments or notes in the selection.'); return; }
      book.transact(found.length > 1 ? 'Delete Comments' : 'Delete Comment', () => { book.recordSheet(s); for (const k of found) s.comments.delete(k); });
    },
  },
  {
    id: 'review.deleteAllComments', label: 'Delete All Comments in Sheet', icon: 'delete_forever', tab: RV, group: 'Comments', edits: true,
    run: async (app) => {
      if (objectsLocked(app)) return;
      const s = app.sheet;
      const n = [...s.comments.values()].filter((c) => !c.note).length;
      if (!n) { toast('There are no comments on this sheet.'); return; }
      if (!(await confirmDialog('Delete All Comments', `Delete all ${n} comment thread${n > 1 ? 's' : ''} on this sheet? Notes are kept.`, 'Delete'))) return;
      app.book.transact('Delete All Comments', () => { app.book.recordSheet(s); for (const [k, cm] of [...s.comments]) if (!cm.note) s.comments.delete(k); });
    },
  },
  { id: 'review.prevComment', label: 'Previous Comment', icon: 'arrow_upward', tab: RV, group: 'Comments', noRecord: true, run: (app) => stepComment(app, false, -1) },
  { id: 'review.nextComment', label: 'Next Comment', icon: 'arrow_downward', tab: RV, group: 'Comments', noRecord: true, run: (app) => stepComment(app, false, 1) },
  { id: 'review.showComments', label: 'Show Comments', icon: 'forum', tab: RV, group: 'Comments', kind: 'toggle', noRecord: true, run: (app) => app.ui.togglePane('comments'), isActive: (app) => !!app.ui?.paneOpen?.('comments') },
  // Notes
  { id: 'review.notes', label: 'Notes', icon: 'sticky_note_2', tab: RV, group: 'Notes', kind: 'menu', noRecord: true, run: (app) => app.exec('review.newNote'), menu: NOTES_MENU },
  {
    id: 'review.newNote', label: 'New Note', icon: 'sticky_note_2', tab: RV, group: 'Notes', keys: ['Shift-F2'], edits: true,
    run: async (app, arg) => {
      if (objectsLocked(app)) return;
      const s = app.sheet;
      const { r, c } = target(app, arg);
      const cur = s.comments.get(cellKey(r, c));
      if (cur && !cur.note) { toast('This cell has a comment thread. A cell can have a comment or a note.'); return; }
      if (cur?.note) return app.exec('review.editNote', { r, c, text: arg?.text });
      const text = typeof arg?.text === 'string' ? arg.text : await askText(`New Note — ${cellName(r, c)}`, `${app.prefs.author}:`, '', 'Save');
      if (!text) return;
      setComment(app.book, s, r, c, { author: app.prefs.author, text, time: Date.now(), note: true }, 'New Note');
    },
  },
  {
    id: 'review.editNote', label: 'Edit Note', icon: 'edit_note', tab: RV, group: 'Notes', edits: true,
    run: async (app, arg) => {
      if (objectsLocked(app)) return;
      const s = app.sheet;
      const { r, c } = target(app, arg);
      const cur = s.comments.get(cellKey(r, c));
      if (!cur?.note) return app.exec('review.newNote', { r, c, text: arg?.text });
      const text = typeof arg?.text === 'string' ? arg.text : await askText(`Edit Note — ${cellName(r, c)}`, `${cur.author}:`, cur.text, 'Save');
      if (text == null) return;
      setComment(app.book, s, r, c, { ...cur, text, time: Date.now() }, 'Edit Note');
    },
  },
  {
    id: 'review.deleteNote', label: 'Delete Note', icon: 'delete', tab: RV, group: 'Notes', edits: true,
    run: (app, arg) => {
      if (objectsLocked(app)) return;
      const s = app.sheet;
      const { r, c } = target(app, arg);
      if (!s.comments.get(cellKey(r, c))?.note) { toast('There is no note in this cell.'); return; }
      setComment(app.book, s, r, c, undefined, 'Delete Note');
    },
  },
  { id: 'review.prevNote', label: 'Previous Note', icon: 'arrow_upward', tab: RV, group: 'Notes', noRecord: true, run: (app) => stepComment(app, true, -1) },
  { id: 'review.nextNote', label: 'Next Note', icon: 'arrow_downward', tab: RV, group: 'Notes', noRecord: true, run: (app) => stepComment(app, true, 1) },
  {
    id: 'review.showNote', label: 'Show/Hide Note', icon: 'visibility', tab: RV, group: 'Notes', noRecord: true,
    run: (app) => {
      const { r, c } = activeCell(app);
      const cm = app.sheet.comments.get(cellKey(r, c));
      if (!cm) { toast('There is no note in this cell.'); return; }
      const key = `${app.sheet.id}:${r}:${c}`;
      const tab = app.tab as any;
      if (tab.shownNote === key) { tab.shownNote = null; app.grid.objects.hoverComment(-1, -1); }
      else { tab.shownNote = key; app.grid.objects.hoverComment(r, c); }
    },
  },
  {
    id: 'review.showAllNotes', label: 'Show All Notes', icon: 'library_books', tab: RV, group: 'Notes', kind: 'toggle', noRecord: true,
    run: (app) => { const o = app.grid.objects; o.showAllComments = !o.showAllComments; o.position(); },
    isActive: (app) => !!app.grid?.objects?.showAllComments,
  },
  {
    id: 'review.convertNotes', label: 'Convert Notes to Comments', icon: 'forum', tab: RV, group: 'Notes', edits: true,
    run: async (app) => {
      const book = app.book;
      const sheets = book.sheets.filter((s) => [...s.comments.values()].some((c) => c.note));
      if (!sheets.length) { toast('There are no notes to convert.'); return; }
      if (sheets.some((s) => s.protection.on && !s.protection.allow.includes('objects'))) { toast('Some notes are on protected sheets. Unprotect them first.'); return; }
      const n = sheets.reduce((a, s) => a + [...s.comments.values()].filter((c) => c.note).length, 0);
      if (!(await confirmDialog('Convert Notes', `Convert all ${n} note${n > 1 ? 's' : ''} in this workbook to comment threads? People can then reply to them.`, 'Convert All Notes'))) return;
      book.transact('Convert Notes', () => {
        for (const s of sheets) {
          book.recordSheet(s);
          for (const [k, cm] of [...s.comments]) if (cm.note) s.comments.set(k, { author: cm.author, text: cm.text, time: cm.time, replies: [] });
        }
      });
      toast(`${n} note${n > 1 ? 's' : ''} converted`);
    },
  },
  // Protection
  { id: 'review.protectSheet', label: 'Protect Sheet', icon: 'lock', tab: RV, group: 'Protect', edits: false, popular: true, run: (app) => (app.isReadOnly() ? toast('This workbook is read-only.') : protectSheet(app)), isActive: (app) => app.sheet.protection.on },
  { id: 'review.protectWorkbook', label: 'Protect Workbook', icon: 'shield_lock', tab: RV, group: 'Protect', edits: false, run: (app) => (app.isReadOnly() ? toast('This workbook is read-only.') : protectWorkbook(app)), isActive: (app) => app.book.structureLocked },
  { id: 'review.allowEditRanges', label: 'Allow Edit Ranges', icon: 'edit_square', tab: RV, group: 'Protect', edits: true, tip: 'Let people edit chosen ranges of a protected sheet, optionally with a password', run: (app) => allowEditRanges(app) },
  { id: 'review.canEditRange', label: 'Can Edit Protected Cell', tab: RV, internal: true, noRecord: true, edits: false, run: (app, arg) => canEditRange(app, arg) },
  {
    id: 'review.markFinal', label: 'Mark as Final', icon: 'verified', tab: RV, group: 'Protect', kind: 'toggle', edits: false, noRecord: true,
    tip: 'Let readers know the workbook is final and make it read-only',
    run: async (app) => {
      const book = app.book;
      if (book.finalMarked) {
        if (!(await confirmDialog('Mark as Final', 'An author has marked this workbook as final to discourage editing. Edit it anyway?', 'Edit Anyway'))) return;
        book.finalMarked = false;
        toast('The workbook is no longer marked as final.');
      } else {
        if (!(await confirmDialog('Mark as Final', 'This workbook will be marked as final and become read-only. Save it to keep the mark.', 'Mark as Final'))) return;
        book.finalMarked = true;
        toast('Marked as final. Editing commands are turned off.');
      }
      app.bookChanged(app.tab);
    },
    isActive: (app) => app.book.finalMarked,
  },
  {
    id: 'review.protectionNotice', label: 'Protection Notice', icon: 'lock', tab: RV, internal: true, noRecord: true, edits: false,
    run: (app) => {
      if (app.book.finalMarked) { toast('An author has marked this workbook as final.', { action: 'Edit Anyway', onAction: () => app.exec('review.markFinal'), ms: 5000 }); return; }
      if (app.tab.readOnly) { toast('This workbook is open read-only. Save a copy to make changes.'); return; }
      toast("The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet (Review ▸ Unprotect Sheet). You might be asked for a password.", { ms: 5000 });
    },
  },
  // Ink
  {
    id: 'review.hideInk', label: 'Hide Ink', icon: 'ink_eraser', tab: RV, group: 'Ink', kind: 'toggle', noRecord: true, tip: 'Hide or show ink drawings on the sheets (hidden ink does not print)',
    run: (app) => { const tab = app.tab as any; tab.inkHidden = !tab.inkHidden; tab.el?.classList.toggle('ql-hide-ink', tab.inkHidden); if (!app.sheet.drawings.some((d) => d.kind === 'ink') && tab.inkHidden) toast('There is no ink on this sheet. Ink you draw stays hidden until you turn this off.'); },
    isActive: (app) => !!(app.tab as any)?.inkHidden,
  },
]);

