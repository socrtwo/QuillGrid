/**
 * Apple Numbers documents (.numbers), read-only.
 *
 * Numbers '09 (and Numbers 1.x) saves an XML document, index.xml or index.xml.gz, in a ZIP or a
 * package folder: each sheet ("workspace") holds tables (sf:tabular-model) whose grid lists the
 * cells in order (sf:t text, sf:n number, sf:d date, sf:b boolean, sf:f formula with its cached
 * result, sf:g empty, sf:s spanned, sf:pm pop-up menu, sf:cb checkbox), with sf:ct encoding a
 * jump in the cell position. Read: values, formulas (converted from Numbers syntax, including
 * references to other tables), column widths, row heights, merged cells and the basic cell
 * formats (number and date formats, bold, italic, font, size, color, alignment, fill).
 *
 * Numbers 2013 and later saves iWork archives (Index/*.iwa, or Index.zip inside the package):
 * Snappy-compressed protocol buffer messages (iwa.ts). Read: sheets and tables, cell values
 * (numbers, text, rich text, dates, booleans, durations, cached formula results), column widths
 * and row heights, and date formats. Formulas are stored as expression trees in a separate
 * calculation engine and are opened as their last values.
 *
 * Each table becomes a sheet named after the table (Numbers sheets can hold several tables),
 * which is also how other spreadsheet programs open these files. When the package holds only a
 * preview picture, the picture is opened with a note saying so.
 */
import JSZip from 'jszip';
import { Workbook, Sheet, newId } from '../model/workbook';
import { CellError, errorFromCode, type Cell, type CellStyle } from '../model/types';
import { cellKey, MAX_COLS, MAX_ROWS } from '../model/address';
import { parse } from '../calc/parser';
import { parseXml, kids, kid, desc, attr, decodeText, bytesToBase64, imageSize } from './index';
import { gunzip, isGzip } from './gzip';
import {
  unframeIwa, indexArchives, pbDecode, pbMsg, pbMsgs, pbVarint, pbString, pbRef, pbRefs, pbBytes, pbFloat, type IwaObject, type PbMessage,
} from './iwa';

/** Seconds since 2001-01-01 (the Apple reference date) -> date serial. */
const EPOCH_2001 = 36892;
const appleSecondsToSerial = (s: number) => EPOCH_2001 + s / 86400;

const uniqueName = (book: Workbook, name: string) => {
  let nm = (name || 'Table').replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Table';
  let i = 2;
  while (book.sheetByName(nm)) nm = `${(name || 'Table').slice(0, 27)} (${i++})`;
  return nm;
};

/** ICU date pattern (Numbers) -> spreadsheet format code. */
export function icuToFormat(p: string): string {
  let out = '';
  for (let i = 0; i < p.length;) {
    const ch = p[i];
    if (ch === "'") { const j = p.indexOf("'", i + 1); const lit = p.slice(i + 1, j < 0 ? undefined : j); out += lit ? `"${lit}"` : "'"; i = j < 0 ? p.length : j + 1; continue; }
    let n = 1;
    while (p[i + n] === ch) n++;
    switch (ch) {
      case 'y': out += n === 2 ? 'yy' : 'yyyy'; break;
      case 'M': case 'L': out += n >= 4 ? 'mmmm' : n === 3 ? 'mmm' : n === 2 ? 'mm' : 'm'; break;
      case 'd': out += n >= 2 ? 'dd' : 'd'; break;
      case 'E': case 'e': out += n >= 4 ? 'dddd' : 'ddd'; break;
      case 'H': case 'h': case 'k': case 'K': out += n >= 2 ? 'hh' : 'h'; break;
      case 'm': out += n >= 2 ? 'mm' : 'm'; break;
      case 's': out += n >= 2 ? 'ss' : 's'; break;
      case 'S': out += '.' + '0'.repeat(Math.min(3, n)); if (out.endsWith('..' + '0'.repeat(Math.min(3, n)))) out = out.replace(/\.\./, '.'); break;
      case 'a': out += 'AM/PM'; break;
      default: out += /[A-Za-z]/.test(ch) ? '' : ch.repeat(n);
    }
    i += n;
  }
  return out.trim() || 'yyyy-mm-dd';
}

const CURRENCY: Record<string, string> = { USD: '$', EUR: '€', GBP: '£', JPY: '¥', CNY: '¥', INR: '₹', KRW: '₩', CHF: 'CHF', CAD: '$', AUD: '$' };

// --------------------------------------------------------------------------------------------
// Entry

export async function importNumbers(data: ArrayBuffer, warnings: string[] = []): Promise<Workbook> {
  const u8 = new Uint8Array(data);
  if (u8[0] === 0x3c || (u8[0] === 0xef && u8[1] === 0xbb)) return readNumbers09(decodeText(u8), null, warnings);
  if (isGzip(u8)) return readNumbers09(decodeText(await gunzip(u8)), null, warnings);
  const zip = await JSZip.loadAsync(data);
  const files = new Map<string, JSZip.JSZipObject>();
  zip.forEach((p, f) => { if (!f.dir) files.set(p, f); });
  const find = (re: RegExp) => [...files.keys()].find((k) => re.test(k));
  // Numbers '09
  const xmlName = find(/(^|\/)index\.xml(\.gz)?$/i);
  if (xmlName) {
    let bytes = await files.get(xmlName)!.async('uint8array');
    if (isGzip(bytes)) bytes = await gunzip(bytes);
    return readNumbers09(decodeText(bytes), zip, warnings);
  }
  // Numbers 2013+: Index/*.iwa, possibly inside Index.zip
  let iwa = [...files.keys()].filter((k) => /(^|\/)Index\/.*\.iwa$/i.test(k)).map((k) => ({ name: k, file: files.get(k)! }));
  const inner = find(/(^|\/)Index\.zip$/i);
  if (!iwa.length && inner) {
    const z2 = await JSZip.loadAsync(await files.get(inner)!.async('uint8array'));
    iwa = [];
    z2.forEach((p, f) => { if (!f.dir && /\.iwa$/i.test(p)) iwa.push({ name: p, file: f }); });
  }
  if (iwa.length) {
    const objects = new Map<number, IwaObject>();
    let bad = 0;
    for (const { name, file } of iwa) {
      try { indexArchives(unframeIwa(await file.async('uint8array')), name, objects); } catch { bad++; }
    }
    if (!objects.size) throw new Error('The Numbers document could not be read; it may be protected with a password.');
    if (bad) warnings.push(`${bad} part${bad > 1 ? 's' : ''} of the document could not be read.`);
    return readIwa(objects, warnings);
  }
  // Only a preview
  const preview = find(/(^|\/)(QuickLook\/Thumbnail\.(jpg|png)|preview\.(jpg|png)|preview-web\.jpg)$/i);
  if (preview) return previewOnly(await files.get(preview)!.async('uint8array'), preview, warnings);
  throw new Error('This Numbers document has no readable tables.');
}

function previewOnly(bytes: Uint8Array, name: string, warnings: string[]): Workbook {
  const book = new Workbook(false);
  const sheet = new Sheet(book, 'Preview');
  book.sheets.push(sheet);
  const sz = imageSize(bytes) || { w: 640, h: 480 };
  const mime = /png$/i.test(name) ? 'image/png' : 'image/jpeg';
  sheet.drawings.push({ id: newId('img'), kind: 'image', at: { r: 0, c: 0, dx: 0, dy: 0, w: sz.w, h: sz.h }, src: `data:${mime};base64,${bytesToBase64(bytes)}`, alt: 'Preview of the Numbers document' });
  warnings.push('This Numbers document contains only a preview picture; its tables could not be read. The preview was opened instead.');
  book.clearHistory();
  return book;
}

// --------------------------------------------------------------------------------------------
// Numbers '09 XML

interface StyleInfo { style: CellStyle; numFmt?: string; dateFmt?: string }

function readNumbers09(text: string, zip: JSZip | null, warnings: string[]): Workbook {
  const doc = parseXml(text, 'Numbers document');
  const root = doc.documentElement;
  const book = new Workbook(false);
  // Styles by sfa:ID and by sf:ident (parents are referenced by ident).
  const byId = new Map<string, Element>(), byIdent = new Map<string, Element>();
  for (const st of desc(root, 'cell-style')) {
    const id = attr(st, 'ID'); if (id) byId.set(id, st);
    const ident = attr(st, 'ident'); if (ident) byIdent.set(ident, st);
  }
  const resolved = new Map<string, StyleInfo>();
  const resolve = (el: Element | undefined, depth = 0): StyleInfo => {
    if (!el || depth > 10) return { style: {} };
    const id = attr(el, 'ID') || '';
    if (id && resolved.has(id)) return resolved.get(id)!;
    const parent = attr(el, 'parent-ident');
    const base = parent ? resolve(byIdent.get(parent), depth + 1) : { style: {} };
    const info: StyleInfo = { style: { ...base.style }, numFmt: base.numFmt, dateFmt: base.dateFmt };
    const pm = kid(el, 'property-map');
    for (const p of kids(pm)) readStyleProp(p, info);
    if (id) resolved.set(id, info);
    return info;
  };
  const sid = new Map<string, number>();
  const styleFor = (ref: string | null, kind: 'number' | 'date' | 'other'): number => {
    if (!ref) return 0;
    const key = ref + '|' + kind;
    let s = sid.get(key);
    if (s == null) {
      const info = resolve(byId.get(ref));
      const st: CellStyle = { ...info.style };
      if (kind === 'date') st.numFmt = info.dateFmt || 'yyyy-mm-dd';
      else if (kind === 'number' && info.numFmt && info.numFmt !== 'General') st.numFmt = info.numFmt;
      s = book.styleId(st);
      sid.set(key, s);
    }
    return s;
  };

  // Table ids -> sheet names (for cross-table references).
  const tableNames = new Map<string, NumTable>();
  const workspaces = desc(root, 'workspace');
  const plan: { ws: Element; models: Element[] }[] = [];
  const reserve = new Set<string>();
  for (const ws of workspaces) {
    const models = desc(ws, 'tabular-model').filter((m) => hasAncestor(m, 'tabular-info'));
    for (const m of models) {
      let nm = attr(m, 'name') || 'Table';
      let i = 2;
      while (reserve.has(nm.toLowerCase())) nm = `${attr(m, 'name') || 'Table'} (${i++})`;
      reserve.add(nm.toLowerCase());
      const grid = kid(m, 'grid');
      const rows = +(attr(grid, 'numrows') || 0), hr = +(attr(m, 'num-header-rows') || 0), fr = +(attr(m, 'num-footer-rows') || 0);
      const id = attr(m, 'id'); if (id) tableNames.set(id.toUpperCase(), { name: nm, r1: hr, r2: Math.max(hr, rows - fr - 1) });
    }
    plan.push({ ws, models });
  }
  let badFormulas = 0;
  for (const { ws, models } of plan) {
    if (!models.length) {
      const sheet = new Sheet(book, uniqueName(book, attr(ws, 'workspace-name') || 'Sheet'));
      book.sheets.push(sheet);
      continue;
    }
    for (const m of models) {
      const self = tableNames.get((attr(m, 'id') || '').toUpperCase()) || { name: attr(m, 'name') || 'Table', r1: 0, r2: 0 };
      const sheet = new Sheet(book, uniqueName(book, self.name));
      book.sheets.push(sheet);
      badFormulas += readTable09(sheet, m, styleFor, tableNames, self);
      const hr = +(attr(m, 'num-header-rows') || 0), hc = +(attr(m, 'num-header-columns') || 0);
      if (attr(m, 'header-rows-frozen') === 'true' || attr(m, 'header-columns-frozen') === 'true') sheet.freeze = { rows: attr(m, 'header-rows-frozen') === 'true' ? hr : 0, cols: attr(m, 'header-columns-frozen') === 'true' ? hc : 0 };
    }
  }
  if (desc(root, 'chart-info').length) warnings.push('Charts in Numbers documents are not opened.');
  if (badFormulas) warnings.push('Some Numbers formulas could not be converted; their last values were kept.');
  if (!book.sheets.length) book.sheets.push(new Sheet(book, 'Sheet1'));
  void zip;
  book.clearHistory();
  return book;
}

function hasAncestor(el: Element, local: string): boolean {
  for (let p = el.parentElement; p; p = p.parentElement) if (p.localName === local) return true;
  return false;
}

function colorOf(el: Element | null): string | undefined {
  const c = el && (el.localName === 'color' ? el : desc(el, 'color')[0]);
  if (!c) return undefined;
  const r = +(attr(c, 'r') ?? NaN), g = +(attr(c, 'g') ?? NaN), b = +(attr(c, 'b') ?? NaN), a = +(attr(c, 'a') ?? 1);
  if (![r, g, b].every(Number.isFinite)) { const w = +(attr(c, 'w') ?? NaN); if (!Number.isFinite(w)) return undefined; return '#' + Math.round(w * 255).toString(16).padStart(2, '0').repeat(3); }
  // Blend translucent colors over white, as they appear on screen.
  const mix = (x: number) => Math.round((x * a + (1 - a)) * 255);
  return '#' + [r, g, b].map((x) => mix(x).toString(16).padStart(2, '0')).join('');
}

function readStyleProp(p: Element, info: StyleInfo): void {
  const num = (e: Element) => { const n = desc(e, 'number')[0]; return n ? +(attr(n, 'number') || 0) : NaN; };
  switch (p.localName) {
    case 'SFTCellStylePropertyNumberFormat': {
      const nf = kids(p)[0];
      const fs = attr(nf, 'format-string');
      if (fs != null) {
        const sym = CURRENCY[attr(nf, 'format-currency-code') || ''] || attr(nf, 'format-currency-code') || '$';
        info.numFmt = fs.replace(/¤/g, sym.length === 1 ? sym : `"${sym} "`).replace(/#,##0\.###/g, '#,##0.###') || undefined;
        if (/^#,##0\.###;-#,##0\.###$/.test(info.numFmt || '')) info.numFmt = undefined; // "Automatic"
      }
      break;
    }
    case 'SFTCellStylePropertyDateTimeFormat': {
      const df = desc(p, 'date-format')[0];
      const f = attr(df, 'fmt');
      if (f) info.dateFmt = icuToFormat(f);
      break;
    }
    case 'SFTCellStylePropertyFill': case 'fill': {
      const c = colorOf(p);
      if (c) info.style.fill = c;
      break;
    }
    case 'SFTCellStylePropertyParagraphStyle': case 'SFTCellStylePropertyLayoutStyle': {
      for (const pm of desc(p, 'property-map')) for (const q of kids(pm)) readStyleProp(q, info);
      break;
    }
    case 'bold': if (num(p) === 1) info.style.bold = true; else if (num(p) === 0) delete info.style.bold; break;
    case 'italic': if (num(p) === 1) info.style.italic = true; break;
    case 'underline': if (num(p) >= 1) info.style.underline = 'single'; break;
    case 'strikethru': if (num(p) >= 1) info.style.strike = true; break;
    case 'fontSize': { const n = num(p); if (n > 0) info.style.size = n; break; }
    case 'fontName': { const s = attr(desc(p, 'string')[0], 'string'); if (s) info.style.font = s.replace(/-(Bold|Italic|BoldItalic|Regular|Roman|Medium|Light)$/i, '').replace(/MT$/, ''); break; }
    case 'fontColor': { const c = colorOf(p); if (c && c !== '#000000') info.style.color = c; break; }
    case 'alignment': { const n = num(p); const m: Record<number, CellStyle['hAlign']> = { 0: 'left', 1: 'right', 2: 'center', 3: 'justify' }; if (m[n]) info.style.hAlign = m[n]; else if (n === 4) delete info.style.hAlign; break; }
    case 'SFTCellTextWrapProperty': if (num(p) === 1) info.style.wrap = true; break;
  }
}

/** A Numbers table as seen from formulas: its sheet name and the rows of its body (0-based). */
export interface NumTable { name: string; r1: number; r2: number }

/**
 * Converts a Numbers formula to spreadsheet syntax. References to other tables
 * ("SFTGlobalID_<id> :: B2") become sheet references, and a bare column ("B", "$D"), which in
 * Numbers means the body cells of that column, becomes the matching row range.
 */
export function numbersFormula(fs: string, tables: Map<string, NumTable>, self: NumTable): string | null {
  let f = fs.trim().replace(/^=/, '');
  if (/__CHART\(/.test(f)) return null;
  const quote = (n: string) => (/^[A-Za-z_][\w.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
  const col = (t: NumTable, d: string, c: string) => `${d}${c}$${t.r1 + 1}:${d}${c}$${t.r2 + 1}`;
  const seg = (s: string) => s
    .replace(/SFTGlobalID_([0-9A-Fa-f]+)\s*::\s*(\$?)([A-Z]{1,3})(\$?\d+)?(?![\w(])/g, (_m, id: string, d: string, c: string, row: string | undefined) => {
      const t = tables.get(id.toUpperCase());
      if (!t) return '#REF!';
      return quote(t.name) + '!' + (row ? `${d}${c}${row}` : col(t, d, c));
    })
    .replace(/(?<![\w.$:!'"])(\$?)([A-Z]{1,3})(?![\w(:!$])/g, (m, d: string, c: string) => (/^(TRUE|FALSE)$/.test(c) ? m : col(self, d, c)));
  let out = '', i = 0, start = 0;
  while (i < f.length) {
    if (f[i] === '"') {
      out += seg(f.slice(start, i));
      let j = i + 1;
      while (j < f.length && !(f[j] === '"' && f[j + 1] !== '"')) j += f[j] === '"' ? 2 : 1;
      out += f.slice(i, j + 1);
      i = start = j + 1;
      continue;
    }
    i++;
  }
  f = out + seg(f.slice(start));
  f = f.replace(/\s*([+\-*/^&=<>,])\s*/g, '$1').replace(/\(\s+/g, '(').replace(/\s+\)/g, ')');
  try { parse(f); } catch { return null; }
  return f;
}

function readTable09(sheet: Sheet, m: Element, styleFor: (ref: string | null, kind: 'number' | 'date' | 'other') => number, tableNames: Map<string, NumTable>, self: NumTable): number {
  const grid = kid(m, 'grid');
  if (!grid) return 0;
  kids(kid(grid, 'columns'), 'grid-column').forEach((c, i) => { const w = +(attr(c, 'width') || 0); if (w) sheet.colW.set(i, Math.round((w * 4) / 3)); });
  kids(kid(grid, 'rows'), 'grid-row').forEach((r, i) => { const h = +(attr(r, 'height') || 0); if (h && Math.abs(Math.round((h * 4) / 3) - sheet.defaultRowH) > 1) sheet.rowH.set(i, Math.round((h * 4) / 3)); });
  const ds = kid(grid, 'datasource');
  if (!ds) return 0;
  // Pop-up menu choices are referenced by id.
  const choices = new Map<string, Element>();
  for (const mc of desc(ds, 'menu-choices')) for (const c of kids(mc)) { const id = attr(c, 'ID'); if (id) choices.set(id, c); }
  let r = 0, c = -1, first = true, bad = 0;
  for (const el of kids(ds)) {
    const ct = attr(el, 'ct');
    if (ct != null) {
      const n = +ct;
      const dr = Math.floor((n + 128) / 256), dc = n - dr * 256;
      if (first) { r = dr; c = dc; } else { r += dr; c += dc; }
    } else c = first ? 0 : c + 1;
    first = false;
    if (r >= MAX_ROWS || c >= MAX_COLS || c < 0) continue;
    const cell = cell09(el, choices, styleFor, tableNames, self);
    if (cell === 'bad') { bad++; continue; }
    const cs = +(attr(el, 'col-span') || 1), rs = +(attr(el, 'row-span') || 1);
    if (cs > 1 || rs > 1) sheet.merges.push({ r1: r, c1: c, r2: r + rs - 1, c2: c + cs - 1 });
    if (cell && (cell.v != null || cell.f != null || cell.s)) sheet.cells.set(cellKey(r, c), cell);
  }
  return bad;
}

function value09(el: Element, choices: Map<string, Element>): { v: Cell['v']; kind: 'number' | 'date' | 'other' } {
  switch (el.localName) {
    case 't': case 'rs': {
      const ct = kid(el, 'ct');
      let s = ct ? attr(ct, 's') : null;
      if (s == null) s = desc(el, 'p').map((p) => p.textContent || '').join('\n') || desc(el, 'text-body')[0]?.textContent || '';
      return { v: s, kind: 'other' };
    }
    case 'n': case 'rn': { const n = Number(attr(el, 'v')); return { v: Number.isFinite(n) ? n : null, kind: 'number' }; }
    case 'd': case 'rd': { const s = Number(attr(el, 'cell-date')); return { v: Number.isFinite(s) ? appleSecondsToSerial(s) : null, kind: 'date' }; }
    case 'b': case 'rb': case 'cb': { const v = attr(el, 'v'); return { v: v === '1' || v === 'true', kind: 'other' }; }
    case 'du': case 'rdu': { const s = Number(attr(el, 'du') ?? attr(el, 'v')); return { v: Number.isFinite(s) ? s / 86400 : null, kind: 'number' }; }
    case 're': return { v: errorFromCode(attr(el, 'v') || '') || new CellError('#VALUE!'), kind: 'other' };
    case 'pm': {
      const ref = desc(el, 'proxied-cell-ref')[0];
      const target = ref ? choices.get(attr(ref, 'IDREF') || '') : undefined;
      return target ? value09(target, choices) : { v: null, kind: 'other' };
    }
  }
  return { v: null, kind: 'other' };
}

function cell09(el: Element, choices: Map<string, Element>, styleFor: (ref: string | null, kind: 'number' | 'date' | 'other') => number, tableNames: Map<string, NumTable>, self: NumTable): Cell | null | 'bad' {
  const n = el.localName;
  if (n === 'g' || n === 's' || n === 'e') { const s = styleFor(attr(el, 's'), 'other'); return s ? { s } : null; }
  const cell: Cell = {};
  let kind: 'number' | 'date' | 'other' = 'other';
  if (n === 'f') {
    const fo = kid(el, 'fo');
    const res = kids(kid(el, 'r'))[0];
    if (res) { const x = value09(res, choices); if (x.v != null) cell.v = x.v; kind = x.kind; }
    const fs = attr(fo, 'fs');
    if (fs) { const f = numbersFormula(fs, tableNames, self); if (f) cell.f = f; else if (!/__CHART/.test(fs)) { if (cell.v == null) return 'bad'; } }
  } else {
    const x = value09(el, choices);
    if (x.v != null && x.v !== '') cell.v = x.v;
    kind = x.kind;
  }
  const s = styleFor(attr(el, 's'), kind);
  if (s) cell.s = s;
  return cell;
}

// --------------------------------------------------------------------------------------------
// Numbers 2013+ (iWork archives)

/** Message type ids of the iWork archives used here. */
const T = { Document: 1, Sheet: 2, TableInfo: 6000, TableModel: 6001, Tile: 6002, DataList: 6005, HeaderBucket: 6006 } as const;

function readDecimal128(b: Uint8Array, p: number): number {
  const exp = (((b[p + 15] & 0x7f) << 7) | (b[p + 14] >> 1)) - 0x1820;
  let mant = b[p + 14] & 1;
  for (let j = 13; j >= 0; j--) mant = mant * 256 + b[p + j];
  const v = mant * Math.pow(10, exp);
  return b[p + 15] & 0x80 ? -v : v;
}

const popcnt = (x: number) => { let n = 0; for (let v = x >>> 0; v; v &= v - 1) n++; return n; };

interface IwaCell { type: number; num?: number; date?: number; str?: number; rich?: number; fmt?: number }

/** Decodes one cell storage record (pre-2016 version 4, and the version 5 layout). */
export function decodeIwaCell(buf: Uint8Array): IwaCell | null {
  if (buf.length < 8) return null;
  const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const ver = buf[0], type = buf[1];
  const out: IwaCell = { type };
  if (ver >= 5) {
    const flags = dv.getUint32(8, true);
    let p = 12;
    if (flags & 0x1) { out.num = readDecimal128(buf, p); p += 16; }
    if (flags & 0x2) { out.num = dv.getFloat64(p, true); p += 8; }
    if (flags & 0x4) { out.date = dv.getFloat64(p, true); p += 8; }
    if (flags & 0x8) { out.str = dv.getUint32(p, true); p += 4; }
    if (flags & 0x10) { out.rich = dv.getUint32(p, true); p += 4; }
    // Then style ids (0x20 cell, 0x40 text, 0x80 conditional, 0x100 rule), 0x200 formula,
    // 0x400 control, 0x800 formula error, 0x1000 suggestion, then format ids from 0x2000.
    for (const bit of [0x20, 0x40, 0x80, 0x100, 0x200, 0x400, 0x800, 0x1000]) if (flags & bit) p += 4;
    if (flags & 0x2000) { out.fmt = dv.getUint32(p, true); }
    return out;
  }
  const flags = dv.getUint32(4, true);
  let p = (ver > 1 ? 12 : 8) + popcnt(flags & (ver > 1 ? 0x0d8e : 0x018e)) * 4;
  if (flags & 0x200) { out.rich = dv.getUint32(p, true); p += 4; }
  p += popcnt(flags & (ver > 1 ? 0x3000 : 0x1000)) * 4;
  if (flags & 0x10) { out.str = dv.getUint32(p, true); p += 4; }
  if (flags & 0x20) { out.num = dv.getFloat64(p, true); p += 8; }
  if (flags & 0x40) { out.date = dv.getFloat64(p, true); p += 8; }
  return out;
}

function readIwa(objects: Map<number, IwaObject>, warnings: string[]): Workbook {
  const book = new Workbook(false);
  const get = (id: number | undefined, type?: number): PbMessage | undefined => {
    if (id == null) return undefined;
    const o = objects.get(id);
    if (!o || (type != null && o.type !== type)) return undefined;
    return pbDecode(o.data);
  };
  const docObj = objects.get(1)?.type === T.Document ? objects.get(1) : [...objects.values()].find((o) => o.type === T.Document);
  const doc = docObj ? pbDecode(docObj.data) : undefined;
  const sheetIds = doc ? pbRefs(doc, 1) : [...objects.values()].filter((o) => o.type === T.Sheet).map((o) => o.id);
  let tables = 0;
  for (const sid of sheetIds) {
    const sh = get(sid, T.Sheet);
    if (!sh) continue;
    const sheetName = pbString(sh, 1) || 'Sheet';
    const infos = pbRefs(sh, 2).map((id) => ({ id, o: objects.get(id) })).filter((x) => x.o?.type === T.TableInfo);
    if (!infos.length) { book.sheets.push(new Sheet(book, uniqueName(book, sheetName))); continue; }
    for (const { id } of infos) {
      const info = get(id, T.TableInfo)!;
      const model = get(pbRef(info, 2), T.TableModel);
      if (!model) continue;
      const sheet = new Sheet(book, uniqueName(book, pbString(model, 8) || sheetName));
      book.sheets.push(sheet);
      readIwaTable(sheet, model, get, warnings);
      tables++;
    }
  }
  if (!tables && !book.sheets.length) throw new Error('No tables were found in this Numbers document.');
  if (!book.sheets.length) book.sheets.push(new Sheet(book, 'Sheet1'));
  warnings.push('Formulas in Numbers documents were opened as their last calculated values.');
  book.clearHistory();
  return book;
}

function readIwaTable(sheet: Sheet, model: PbMessage, get: (id: number | undefined, type?: number) => PbMessage | undefined, warnings: string[]): void {
  const nRows = pbVarint(model, 6) || 0, nCols = pbVarint(model, 7) || 0;
  const store = pbMsg(model, 4);
  if (!store) return;
  // Shared strings and rich text (TableDataList entries: key -> string or rich payload).
  const listOf = (ref: number | undefined) => {
    const m = new Map<number, PbMessage>();
    const dl = get(ref, T.DataList);
    for (const e of pbMsgs(dl, 3)) m.set(pbVarint(e, 1) ?? -1, e);
    return m;
  };
  const strings = listOf(pbRef(store, 4));
  const rich = listOf(pbRef(store, 17));
  const formats = listOf(pbRef(store, 11));
  const richText = (key: number): string | null => {
    const e = rich.get(key);
    const payload = get(pbRef(e, 9)) || get(pbRef(e, 4));
    const storage = payload ? get(pbRef(payload, 1)) : undefined;
    const t = storage ? (pbMsgs(storage, 3).map((x) => pbString(x, 1) ?? '').join('') || pbString(storage, 3)) : null;
    return t ? t.replace(/\u2029/g, '\n').replace(/\n$/, '') : null;
  };
  // Column widths and row heights (points).
  for (const b of [get(pbRef(store, 2))]) for (const h of pbMsgs(b, 2)) {
    const i = pbVarint(h, 1) ?? -1, size = pbFloat(h, 2) || 0;
    if (i >= 0 && i < MAX_COLS && size > 0) sheet.colW.set(i, Math.round((size * 4) / 3));
    if ((pbVarint(h, 3) ?? 0) & 1) sheet.hiddenCols.add(i);
  }
  for (const bid of pbRefs(pbMsg(store, 1), 2)) for (const h of pbMsgs(get(bid), 2)) {
    const i = pbVarint(h, 1) ?? -1, size = pbFloat(h, 2) || 0;
    if (i >= 0 && i < MAX_ROWS && size > 0 && Math.abs(Math.round((size * 4) / 3) - sheet.defaultRowH) > 1) sheet.rowH.set(i, Math.round((size * 4) / 3));
    if ((pbVarint(h, 3) ?? 0) & 1) sheet.hiddenRows.add(i);
  }
  const book = sheet.book;
  const dateStyles = new Map<string, number>();
  const dateStyle = (fmtKey?: number) => {
    const f = fmtKey != null ? pbString(pbMsg(formats.get(fmtKey), 6), 14) : undefined;
    const code = f ? icuToFormat(f) : 'yyyy-mm-dd';
    let s = dateStyles.get(code);
    if (s == null) { s = book.styleId({ numFmt: code }); dateStyles.set(code, s); }
    return s;
  };
  const tiles = pbMsgs(pbMsg(store, 3), 1);
  let unknown = 0;
  for (const t of tiles) {
    const tile = get(pbRef(t, 2), T.Tile);
    if (!tile) continue;
    const base = (pbVarint(t, 1) || 0) * 256;
    for (const ri of pbMsgs(tile, 5)) {
      const r = base + (pbVarint(ri, 1) || 0);
      if (r >= MAX_ROWS || (nRows && r >= nRows)) continue;
      const wide = (pbVarint(ri, 8) || 0) > 0;
      let storage = pbBytes(ri, 6), offsets = pbBytes(ri, 7);
      if (!storage || !offsets) { storage = pbBytes(ri, 3); offsets = pbBytes(ri, 4); }
      if (!storage || !offsets) continue;
      const odv = new DataView(offsets.buffer, offsets.byteOffset, offsets.byteLength);
      const list: [number, number][] = [];
      for (let c = 0; c * 2 + 1 < offsets.length; c++) { const o = odv.getUint16(c * 2, true); if (o !== 0xffff) list.push([c, o * (wide ? 4 : 1)]); }
      list.forEach(([c, o], i) => {
        if (c >= MAX_COLS || (nCols && c >= nCols)) return;
        const end = i + 1 < list.length ? list[i + 1][1] : storage!.length;
        const ic = decodeIwaCell(storage!.subarray(o, end));
        if (!ic) return;
        const cell: Cell = {};
        switch (ic.type) {
          case 0: return;
          case 2: case 10: if (ic.num != null) cell.v = ic.num; break;
          case 3: { const e = ic.str != null ? strings.get(ic.str) : undefined; const s = pbString(e, 3); if (s != null) cell.v = s; break; }
          case 5: if (ic.date != null) { cell.v = appleSecondsToSerial(ic.date); cell.s = dateStyle(ic.fmt); } break;
          case 6: cell.v = !!ic.num; break;
          case 7: if (ic.num != null) { cell.v = ic.num / 86400; cell.s = book.styleId({ numFmt: '[h]:mm:ss' }); } break;
          case 8: cell.v = new CellError('#VALUE!'); break;
          case 9: { const s = ic.rich != null ? richText(ic.rich) : null; if (s != null) cell.v = s; break; }
          default: unknown++;
        }
        if (cell.v != null && cell.v !== '') sheet.cells.set(cellKey(r, c), cell);
      });
    }
  }
  if (unknown && !warnings.includes('Some Numbers cells of an unknown type were skipped.')) warnings.push('Some Numbers cells of an unknown type were skipped.');
}
