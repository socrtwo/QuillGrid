/**
 * KSpread (KOffice 1.x) native documents (.ksp), the format of KDE's spreadsheet before it
 * moved to OpenDocument (KOffice 2 / Calligra Sheets).
 *
 * The document is a KoStore: a ZIP (KOffice 1.3 and later, mimetype application/x-kspread) or a
 * gzip-compressed tar archive (KOffice 1.0-1.2) holding maindoc.xml and documentinfo.xml; a bare
 * maindoc.xml is accepted too. maindoc.xml is <spreadsheet> <map> <table name=...> with
 * <column>, <row> and <cell row column> elements (1-based); each cell has an optional <format>
 * (alignment, colors, font, borders, number format and precision, spans) and a <text> with a
 * dataType (Num, Str, Bool, Date, Time) or a formula. Formulas may store references relative to
 * the cell ("#-1#0#": column offset, row offset) or absolute ("$2$3#"), with ';' between
 * arguments; both are converted to A1 syntax.
 *
 * Fidelity: values, formulas, basic formatting, column widths, row heights, merges, comments,
 * hidden sheets and the active sheet. Charts and embedded objects are not opened.
 */
import JSZip from 'jszip';
import { Workbook, Sheet } from '../model/workbook';
import { type Cell, type CellStyle, type Border } from '../model/types';
import { cellKey, colName, MAX_COLS, MAX_ROWS } from '../model/address';
import { partsToSerial, timeToSerial } from '../model/dates';
import { parseXml, kids, kid, desc, attr, decodeText, normColor } from './index';
import { gunzip, isGzip, untar } from './gzip';

/** Converts KSpread's stored formula to A1 syntax for the cell at (r, c), zero-based. */
export function kspreadFormula(text: string, r: number, c: number): string {
  let f = text.trim().replace(/^=/, '');
  let out = '', i = 0, start = 0;
  const seg = (s: string) => s
    .replace(/([#$])(-?\d+)([#$])(-?\d+)#/g, (_m, ca: string, cn: string, ra: string, rn: string) => {
      const col = ca === '$' ? +cn - 1 : c + +cn, row = ra === '$' ? +rn - 1 : r + +rn;
      if (col < 0 || row < 0 || col >= MAX_COLS || row >= MAX_ROWS) return '#REF!';
      return `${ca === '$' ? '$' : ''}${colName(col)}${ra === '$' ? '$' : ''}${row + 1}`;
    })
    .replace(/;/g, ',');
  while (i < f.length) {
    if (f[i] === '"') {
      out += seg(f.slice(start, i));
      let j = i + 1;
      while (j < f.length && f[j] !== '"') j++;
      out += f.slice(i, j + 1);
      i = start = j + 1;
      continue;
    }
    i++;
  }
  f = out + seg(f.slice(start));
  return f;
}

const H_ALIGN: Record<string, CellStyle['hAlign']> = { 1: 'left', 2: 'center', 3: 'right' };
const V_ALIGN: Record<string, CellStyle['vAlign']> = { 1: 'top', 2: 'middle' };

function penBorder(el: Element | null): Border | undefined {
  const pen = el && kid(el, 'pen');
  if (!pen) return undefined;
  const w = +(attr(pen, 'width') || 0), st = +(attr(pen, 'style') || 0);
  if (!st) return undefined;
  const color = normColor(attr(pen, 'color')) || undefined;
  const style: Border['style'] = st === 2 ? 'dashed' : st === 3 ? 'dotted' : st >= 4 ? 'dashDot' : w >= 3 ? 'thick' : w >= 2 ? 'medium' : 'thin';
  return { style, ...(color && color !== '#000000' ? { color } : {}) };
}

/** KSpread number format ids (Format::FormatType) -> spreadsheet format codes. */
function numberFormat(type: number, precision: number, prefix: string, postfix: string): string | undefined {
  const dec = precision > 0 ? '.' + '0'.repeat(Math.min(precision, 15)) : precision === 0 ? '' : '';
  const wrap = (code: string) => (prefix ? `"${prefix}"` : '') + code + (postfix ? `"${postfix}"` : '');
  if (type === 1) return wrap('0' + (dec || ''));
  if (type === 10) return wrap('#,##0' + (precision >= 0 ? dec : '.00'));
  if (type === 25) return '0' + (dec || '') + '%';
  if (type === 30) return '0' + (dec || '.00') + 'E+00';
  if (type >= 35 && type < 50) return type === 36 ? 'd mmmm yyyy' : 'yyyy-mm-dd';
  if (type >= 50 && type < 70) return type === 51 ? 'h:mm:ss AM/PM' : type === 52 ? 'h:mm AM/PM' : 'h:mm:ss';
  if (type >= 70 && type < 80) return '# ?/?';
  if (precision >= 0 && dec) return wrap('0' + dec);
  if (prefix || postfix) return wrap('General');
  return undefined;
}

function readFormat(fmt: Element | null): { style: CellStyle; colspan: number; rowspan: number } {
  const s: CellStyle = {};
  if (!fmt) return { style: s, colspan: 0, rowspan: 0 };
  const ha = H_ALIGN[attr(fmt, 'align') || '']; if (ha) s.hAlign = ha;
  const va = V_ALIGN[attr(fmt, 'alignY') || '']; if (va) s.vAlign = va;
  const bg = normColor(attr(fmt, 'bgcolor')); if (bg && bg !== '#ffffff') s.fill = bg;
  const fg = normColor(attr(fmt, 'textColor') || attr(fmt, 'textcolor')); if (fg && fg !== '#000000') s.color = fg;
  if (attr(fmt, 'multirow') === 'yes' || attr(fmt, 'multirow') === '1') s.wrap = true;
  const ind = +(attr(fmt, 'indent') || 0); if (ind > 0) s.indent = Math.round(ind / 10);
  const ang = +(attr(fmt, 'angle') || 0); if (ang) s.rotate = Math.max(-90, Math.min(90, -ang));
  if (attr(fmt, 'verticaltext') === 'yes') s.rotate = 255;
  const font = kid(fmt, 'font');
  if (font) {
    const fam = attr(font, 'family'); if (fam) s.font = fam;
    const size = +(attr(font, 'size') || 0); if (size && size !== 10) s.size = size;
    if (+(attr(font, 'weight') || 50) >= 63 || attr(font, 'bold') === 'yes') s.bold = true;
    if (attr(font, 'italic') === '1' || attr(font, 'italic') === 'yes') s.italic = true;
    if (attr(font, 'underline') === '1' || attr(font, 'underline') === 'yes') s.underline = 'single';
    if (attr(font, 'strikeout') === '1' || attr(font, 'strikeout') === 'yes') s.strike = true;
  }
  const bl = penBorder(kid(fmt, 'left-border')), br = penBorder(kid(fmt, 'right-border')), bt = penBorder(kid(fmt, 'top-border')), bb = penBorder(kid(fmt, 'bottom-border'));
  if (bl) s.bl = bl; if (br) s.br = br; if (bt) s.bt = bt; if (bb) s.bb = bb;
  const ftype = +(attr(fmt, 'format') || 0), prec = +(attr(fmt, 'precision') ?? -1);
  const nf = numberFormat(ftype, prec, attr(fmt, 'prefix') || '', attr(fmt, 'postfix') || '');
  if (nf) s.numFmt = nf;
  if (attr(fmt, 'dontprinttext') === 'yes') s.hidden = true;
  return { style: s, colspan: +(attr(fmt, 'colspan') || 0), rowspan: +(attr(fmt, 'rowspan') || 0) };
}

function typedValue(textEl: Element): Cell['v'] {
  const t = textEl.textContent || '';
  switch (attr(textEl, 'dataType')) {
    case 'Num': { const n = Number(t); return Number.isFinite(n) ? n : t; }
    case 'Bool': return /^(true|1|yes)$/i.test(t.trim());
    case 'Date': {
      const m = /^(\d{1,4})[-/.](\d{1,2})[-/.](\d{1,4})/.exec(t.trim());
      if (!m) return t;
      const [y, mo, d] = m[1].length === 4 ? [+m[1], +m[2], +m[3]] : [+m[3], +m[2], +m[1]];
      return partsToSerial(y, mo, d);
    }
    case 'Time': { const m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(t.trim()); return m ? timeToSerial(+m[1], +m[2], +(m[3] || 0)) : t; }
    case 'Str': return t;
  }
  const n = Number(t);
  return t.trim() && Number.isFinite(n) ? n : t;
}

export async function importKspread(data: ArrayBuffer, warnings: string[] = []): Promise<Workbook> {
  let u8: Uint8Array = new Uint8Array(data);
  let main: string | null = null, info: string | null = null;
  if (u8[0] === 0x50 && u8[1] === 0x4b) {
    const zip = await JSZip.loadAsync(data);
    const f = zip.file('maindoc.xml') || zip.file(/maindoc\.xml$/i)[0];
    if (!f) throw new Error('maindoc.xml is missing from this KSpread document.');
    main = decodeText(await f.async('uint8array'));
    const di = zip.file('documentinfo.xml');
    if (di) info = decodeText(await di.async('uint8array'));
  } else {
    if (isGzip(u8)) u8 = await gunzip(u8);
    if (u8.length > 512 && String.fromCharCode(...u8.subarray(257, 262)) === 'ustar') {
      const files = untar(u8);
      const m = [...files.keys()].find((k) => /(^|\/)maindoc\.xml$/i.test(k));
      if (!m) throw new Error('maindoc.xml is missing from this KSpread document.');
      main = decodeText(files.get(m)!);
      const d = [...files.keys()].find((k) => /(^|\/)documentinfo\.xml$/i.test(k));
      if (d) info = decodeText(files.get(d)!);
    } else main = decodeText(u8);
  }
  return readKspreadXml(main, info, warnings);
}

export function readKspreadXml(text: string, info: string | null, warnings: string[] = []): Workbook {
  const doc = parseXml(text.replace(/<!DOCTYPE[^>]*>/, ''), 'KSpread document');
  const root = doc.documentElement;
  if (root.localName !== 'spreadsheet') throw new Error('This is not a KSpread document.');
  const book = new Workbook(false);
  const map = kid(root, 'map') || desc(root, 'map')[0];
  for (const t of kids(map, 'table')) {
    let name = attr(t, 'name') || `Sheet${book.sheets.length + 1}`;
    while (book.sheetByName(name)) name += '_';
    const sheet = new Sheet(book, name);
    book.sheets.push(sheet);
    if (attr(t, 'hide') === '1' || attr(t, 'hide') === 'true') sheet.visibility = 'hidden';
    if (attr(t, 'grid') === '0') sheet.showGrid = false;
    if (attr(t, 'hidezero') === '1') sheet.showZeros = false;
    if (attr(t, 'showFormula') === '1') sheet.showFormulas = true;
    if (attr(t, 'layoutDirection') === 'rtl') sheet.rtl = true;
    for (const col of kids(t, 'column')) {
      const c = +(attr(col, 'column') || 0) - 1;
      if (c < 0 || c >= MAX_COLS) continue;
      const w = +(attr(col, 'width') || 0);
      if (w > 0) sheet.colW.set(c, Math.round((w * 4) / 3));
      if (attr(col, 'hide') === '1') sheet.hiddenCols.add(c);
      const st = readFormat(kid(col, 'format')).style;
      if (Object.keys(st).length) sheet.colStyle.set(c, book.styleId(st));
    }
    for (const row of kids(t, 'row')) {
      const r = +(attr(row, 'row') || 0) - 1;
      if (r < 0 || r >= MAX_ROWS) continue;
      const h = +(attr(row, 'height') || 0);
      if (h > 0 && Math.abs(Math.round((h * 4) / 3) - sheet.defaultRowH) > 1) sheet.rowH.set(r, Math.round((h * 4) / 3));
      if (attr(row, 'hide') === '1') sheet.hiddenRows.add(r);
      const st = readFormat(kid(row, 'format')).style;
      if (Object.keys(st).length) sheet.rowStyle.set(r, book.styleId(st));
    }
    for (const ce of kids(t, 'cell')) {
      const r = +(attr(ce, 'row') || 0) - 1, c = +(attr(ce, 'column') || 0) - 1;
      if (r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS) continue;
      const { style, colspan, rowspan } = readFormat(kid(ce, 'format'));
      const cell: Cell = {};
      const te = kid(ce, 'text');
      if (te) {
        const raw = te.textContent || '';
        if (raw.startsWith('=')) {
          cell.f = kspreadFormula(raw, r, c);
          const res = attr(te, 'outStr');
          if (res != null) { const n = Number(res); cell.v = res.trim() && Number.isFinite(n) ? n : res; }
        } else {
          const v = typedValue(te);
          if (v != null && v !== '') cell.v = v;
          if (attr(te, 'dataType') === 'Date' && !style.numFmt) style.numFmt = 'yyyy-mm-dd';
          if (attr(te, 'dataType') === 'Time' && !style.numFmt) style.numFmt = 'h:mm:ss';
          const link = desc(te, 'a')[0];
          if (link && attr(link, 'href')) { cell.link = attr(link, 'href')!; cell.v = link.textContent || cell.v; }
        }
      }
      const s = book.styleId(style);
      if (s) cell.s = s;
      if (colspan > 0 || rowspan > 0) sheet.merges.push({ r1: r, c1: c, r2: r + Math.max(0, rowspan), c2: c + Math.max(0, colspan) });
      const cm = kid(ce, 'comment');
      if (cm && cm.textContent) sheet.comments.set(cellKey(r, c), { author: '', text: cm.textContent, time: 0, note: true });
      if (cell.v != null || cell.f != null || cell.s) sheet.cells.set(cellKey(r, c), cell);
    }
    if (kids(t, 'object').length || desc(t, 'chart').length) warnings.push('Charts and embedded objects in KSpread files are not opened.');
  }
  if (!book.sheets.length) book.sheets.push(new Sheet(book, 'Sheet1'));
  const active = attr(map, 'activeTable');
  if (active) { const i = book.sheets.findIndex((s) => s.name === active); if (i >= 0) book.activeSheet = i; }
  if (info) {
    try {
      const d = parseXml(info, 'documentinfo.xml');
      const t = (n: string) => (desc(d, n)[0]?.textContent || '').trim();
      book.props.title = t('title'); book.props.author = t('full-name') || t('creator'); book.props.subject = t('subject'); book.props.comments = t('abstract');
      book.props.company = t('company');
    } catch { /* optional */ }
  }
  book.clearHistory();
  return book;
}
