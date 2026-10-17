/**
 * Quattro and Quattro Pro.
 *
 *   WQ1   Quattro / Quattro Pro 1-4 for DOS     Lotus WK1 records (BOF 0x5120); read by lotus.ts
 *   WQ2   Quattro Pro 5 for DOS                 BOF 0x5121; read by lotus.ts
 *   WB1   Quattro Pro 1-5 for Windows           BOF 0x1001-0x1003, 2-byte body
 *   WB2, WB3                                     same record layout
 *   QPW   Quattro Pro 7 and later               compound file; values found by scanning
 *
 * Quattro Pro for Windows notebooks keep each page between a begin-page record (0xCA, page number)
 * and an end-page record (0xCB), with the page name in 0xCC. Cells carry column, page, row and a
 * style number. A formula stores Lotus-style code followed by a table of references: code 0x01 and
 * 0x02 take the next table entry, which is a cell (6 bytes) or, when bit 12 of its first word is set,
 * a block (10 bytes). Relative coordinates are signed offsets, flagged in the row word
 * (0x2000 row, 0x4000 column, 0x8000 page).
 */
import { Workbook, Sheet } from '../../model/workbook';
import { CellError, type Cell, type CellStyle } from '../../model/types';
import { cellKey, colName, MAX_COLS, MAX_ROWS } from '../../model/address';
import { lotusLabel, translateLotus } from '../../heritage/lotus';
import { u16, i16, f64, cstr, codepage, walkRecords, charsToPx, type Decoder } from './bin';
import { importLotus, lotusVariant, type LegacyOptions } from './lotus';
import { LOTUS_FUNCS, unknownFunctions } from './lotusfmla';
import { isOle2, readCfb } from '../recover/doc97';

/** Quattro functions with their own codes (shared ones use the 1-2-3 codes). */
const QUATTRO_FUNCS: Record<number, [string, number]> = {
  0x7a: ['MEMAVAIL', 0], 0x7b: ['MEMEMSAVAIL', 0], 0x7c: ['FILEEXISTS', 1], 0x7d: ['CURVALUE', 2],
  0x80: ['HEX', 1], 0x81: ['NUM', 1], 0x82: ['SUMPRODUCT', 2],
};

export function importQuattro(data: ArrayBuffer | Uint8Array, name = '', warnings: string[] = [], opts: LegacyOptions = {}): Workbook {
  const b = data instanceof Uint8Array ? data : new Uint8Array(data);
  const v = lotusVariant(b);
  if (v === 'wq1' || v === 'wq2') return importLotus(b, name, warnings, opts);
  if (isOle2(b)) return importQpw(b, name, warnings);
  if (u16(b, 0) === 0 && u16(b, 2) === 2 && u16(b, 4) >= 0x1001 && u16(b, 4) <= 0x1007) return importWb(b, name, warnings, opts);
  throw new Error('This is not a Quattro Pro file.');
}

// ------------------------------------------------------------------------------------------
// Quattro Pro for Windows (WB1, WB2, WB3)

interface Ref { file: number; c: number; s: number; r: number; flags: number }

function wbFormula(code: Uint8Array, refs: Uint8Array, at: { r: number; c: number; s: number }, sheetNames: string[], dec: Decoder): string {
  const st: { s: string; atom: boolean }[] = [];
  const pop = () => { const x = st.pop(); if (!x) throw new Error('formula stack underflow'); return x; };
  const wrap = (n: { s: string; atom: boolean }) => (n.atom ? n.s : `(${n.s})`);
  let rp = 0;
  const cellText = (x: Ref, force: boolean): string => {
    const rowRel = !!(x.flags & 0x2000), colRel = !!(x.flags & 0x4000), shRel = !!(x.flags & 0x8000);
    const c = colRel ? at.c + ((x.c << 24) >> 24) : x.c;
    let r = x.r & 0x1fff;
    if (rowRel) { if (r & 0x1000) r -= 0x2000; r += at.r; }
    const s = shRel ? at.s + ((x.s << 24) >> 24) : x.s;
    if (c < 0 || r < 0 || s < 0) throw new Error('reference out of range');
    const prefix = s !== at.s || force ? `${shRel ? '' : '$'}${colName(s)}:` : '';
    return `${prefix}${colRel ? '' : '$'}${colName(c)}${rowRel ? '' : '$'}${r + 1}`;
  };
  const nextRef = (): string => {
    if (rp + 6 > refs.length) throw new Error('missing reference');
    const file = u16(refs, rp);
    if (file & 0x0002 && !(file & 0x1000)) throw new Error('link to another notebook');
    if (file & 0x1000) {
      if (rp + 10 > refs.length) throw new Error('missing reference');
      const a: Ref = { file, c: refs[rp + 2], s: refs[rp + 3], r: u16(refs, rp + 4), flags: u16(refs, rp + 4) & 0xe000 };
      const b: Ref = { file, c: refs[rp + 6], s: refs[rp + 7], r: u16(refs, rp + 8), flags: u16(refs, rp + 8) & 0xe000 };
      rp += 10;
      const cross = a.s !== b.s;
      return `${cellText(a, cross)}..${cellText(b, cross)}`;
    }
    const x: Ref = { file, c: refs[rp + 2], s: refs[rp + 3], r: u16(refs, rp + 4), flags: u16(refs, rp + 4) & 0xe000 };
    rp += 6;
    return cellText(x, false);
  };
  const funcs = { ...LOTUS_FUNCS, ...QUATTRO_FUNCS };
  const OPS: Record<number, [string, number]> = {
    0x08: ['-', 1], 0x09: ['+', 2], 0x0a: ['-', 2], 0x0b: ['*', 2], 0x0c: ['/', 2], 0x0d: ['^', 2], 0x0e: ['=', 2], 0x0f: ['<>', 2],
    0x10: ['<=', 2], 0x11: ['>=', 2], 0x12: ['<', 2], 0x13: ['>', 2], 0x14: ['#AND#', 2], 0x15: ['#OR#', 2], 0x16: ['#NOT#', 1], 0x17: ['+', 1], 0x18: ['&', 2],
  };
  let i = 0;
  while (i < code.length) {
    const op = code[i++];
    if (op === 0x03) break;
    if (op === 0x00) { const v = f64(code, i); i += 8; if (!Number.isFinite(v)) throw new Error('bad constant'); st.push({ s: String(Math.abs(v)), atom: true }); if (v < 0) st[st.length - 1] = { s: `-${Math.abs(v)}`, atom: false }; continue; }
    if (op === 0x01 || op === 0x02) { st.push({ s: nextRef(), atom: true }); continue; }
    if (op === 0x04) { const x = pop(); st.push({ s: `(${x.s})`, atom: true }); continue; }
    if (op === 0x05) { const v = i16(code, i); i += 2; st.push({ s: String(v), atom: v >= 0 }); continue; }
    if (op === 0x06) { let e = i; while (e < code.length && code[e]) e++; st.push({ s: `"${cstr(code, i, dec).replace(/"/g, '""')}"`, atom: true }); i = e + 1; continue; }
    const o = OPS[op];
    if (o) {
      if (o[1] === 1) { const a = pop(); st.push({ s: `${o[0]}${wrap(a)}`, atom: false }); }
      else { const bb = pop(), a = pop(); st.push({ s: `${wrap(a)}${o[0]}${wrap(bb)}`, atom: false }); }
      continue;
    }
    const fn = funcs[op];
    if (!fn) throw new Error(`function code 0x${op.toString(16)} is not known`);
    let n = fn[1];
    if (n < 0) n = code[i++];
    if (st.length < n) throw new Error('formula stack underflow');
    const args = st.splice(st.length - n, n).map((x) => x.s);
    st.push({ s: fn[0] === '@' ? `@@(${args[0]})` : args.length || !['PI', 'TRUE', 'FALSE', 'NA', 'ERR', 'NOW', 'RAND'].includes(fn[0]) ? `@${fn[0]}(${args.join(',')})` : `@${fn[0]}`, atom: true });
  }
  if (st.length !== 1) throw new Error('unbalanced formula');
  return translateLotus(st[0].s, { sheetNames });
}

function importWb(b: Uint8Array, name: string, warnings: string[], opts: LegacyOptions): Workbook {
  const dec = codepage(opts.codepage ?? 1252);
  const book = new Workbook(false);
  const { recs, truncated, skipped } = walkRecords(b, { lenient: !!opts.lenient, stop: (op) => op === 0x0001, plausible: (op, len) => op < 0x0500 && len < 0x8000 });
  if (truncated && !opts.lenient) throw new Error('The notebook file is truncated.');
  if (skipped || truncated) warnings.push('Part of the notebook (pictures or charts) could not be read.');
  // Pass 1: page names.
  const names: string[] = [];
  let page = -1;
  for (const r of recs) {
    if (r.op === 0xca && r.data.length >= 1) page = r.data[0];
    else if (r.op === 0xcc && page >= 0) names[page] = cstr(r.data, 0, dec, 64);
  }
  const sheets: Sheet[] = [];
  const sheetNames = (n: number) => { const out: string[] = []; for (let i = 0; i < n; i++) out.push(names[i] || `Page${i + 1}`); return out; };
  const sheet = (i: number) => {
    while (sheets.length <= i) { const s = new Sheet(book, names[sheets.length] || `Page${sheets.length + 1}`); s.defaultColW = charsToPx(9); sheets.push(s); }
    return sheets[i];
  };
  let maxPage = 0;
  for (const r of recs) if (r.op >= 0x0c && r.op <= 0x10 && r.data.length >= 4) maxPage = Math.max(maxPage, r.data[1]);
  for (let i = 0; i <= maxPage; i++) sheet(i);
  const all = sheetNames(Math.max(maxPage + 1, names.length));
  const styles = new Map<string, number>();
  const sid = (st: CellStyle) => { const k = JSON.stringify(st); let id = styles.get(k); if (id == null) { id = Object.keys(st).length ? book.styleId(st) : 0; styles.set(k, id); } return id; };
  let bad = 0;
  let pending: { sh: Sheet; r: number; c: number } | null = null;
  for (const rec of recs) {
    const d = rec.data;
    if (rec.op < 0x0c || rec.op > 0x10 || d.length < 6) {
      if (rec.op === 0x33 && pending && d.length >= 6) { const cur = pending.sh.get(pending.r, pending.c); if (cur) cur.v = cstr(d, 6, dec); pending = null; }
      continue;
    }
    const c = d[0], s = d[1], r = u16(d, 2);
    if (r >= MAX_ROWS || c >= MAX_COLS) continue;
    const sh = sheet(s);
    try {
      switch (rec.op) {
        case 0x0d: sh.setComputed(r, c, { v: i16(d, 6) }); break;
        case 0x0e: sh.setComputed(r, c, { v: f64(d, 6) }); break;
        case 0x0f: {
          const raw = cstr(d, 6, dec);
          const lab = lotusLabel(raw);
          const cell: Cell = { v: lab ? lab.text : raw.length === 1 && "'\"^\\".includes(raw) ? '' : raw };
          if (lab && lab.align !== 'left') { const st = sid({ hAlign: lab.align }); if (st) cell.s = st; }
          sh.setComputed(r, c, cell);
          break;
        }
        case 0x10: {
          const v = f64(d, 6);
          const total = u16(d, 16), clen = u16(d, 18);
          const code = d.subarray(20, 20 + clen), refs = d.subarray(20 + clen, 20 + total);
          const cell: Cell = { v: Number.isFinite(v) ? v : null };
          try {
            const f = wbFormula(code, refs, { r, c, s }, all, dec);
            const unk = unknownFunctions(f);
            if (unk.length) throw new Error(`unknown function ${unk.join(', ')}`);
            cell.f = f;
          } catch (e: any) {
            bad++;
            sh.comments.set(cellKey(r, c), { author: 'Import', text: `Quattro Pro formula, not translated (${e?.message || e}).`, time: Date.now(), note: true });
          }
          if (cell.v == null && !cell.f) cell.v = new CellError('#VALUE!');
          sh.setComputed(r, c, cell);
          pending = cell.v == null ? { sh, r, c } : null;
          break;
        }
      }
    } catch (e) {
      if (!opts.lenient) throw e;
    }
  }
  for (const sh of sheets) for (const cell of sh.cells.values()) if (cell.f != null && cell.v == null) cell.v = new CellError('#VALUE!');
  if (bad) warnings.push(`${bad} formula${bad > 1 ? 's' : ''} could not be translated and keep their last values (see the cell notes).`);
  if (!sheets.length) sheet(0);
  book.sheets.push(...sheets);
  book.activeSheet = 0;
  if (name) book.props.title = name.replace(/\.[^.]+$/, '');
  return book;
}

// ------------------------------------------------------------------------------------------
// Quattro Pro 7+ (QPW): the record layout of the "NativeContent_MAIN" stream is not public.

function importQpw(b: Uint8Array, name: string, warnings: string[]): Workbook {
  // Strings live in a notebook-wide table (record 0x0407); numbers are packed in column blocks
  // whose layout is unknown, so only the text is recovered, one string per row.
  const cfb = readCfb(b);
  const s = cfb?.read?.('NativeContent_MAIN') || cfb?.read?.('PerfectOffice_MAIN') || null;
  if (!s) throw new Error('The Quattro Pro notebook stream was not found.');
  const book = new Workbook(false);
  const sh = new Sheet(book, 'Text');
  book.sheets.push(sh);
  const dec = codepage(1252);
  const { recs } = walkRecords(s, { lenient: true });
  let row = 0;
  for (const r of recs) {
    if (r.op !== 0x0407) continue;
    // A count, then length-prefixed strings.
    let p = 4;
    const d = r.data;
    while (p < d.length && row < 100000) {
      const n = u16(d, p);
      if (!n || p + 2 + n > d.length) break;
      const t = dec(d.subarray(p + 2, p + 2 + n)).replace(/\0+$/, '');
      if (t) sh.setComputed(row++, 0, { v: t });
      p += 2 + n;
    }
  }
  warnings.push('Quattro Pro 7 and later notebooks (.qpw) are read for their text only; numbers and formulas are not recovered. Save the notebook as .wb3 or .xls in Quattro Pro to keep everything.');
  if (name) book.props.title = name.replace(/\.[^.]+$/, '');
  return book;
}
