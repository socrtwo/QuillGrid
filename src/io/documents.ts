/**
 * Opening word-processor documents, presentations and e-mail in a spreadsheet.
 *
 * The content is first turned into a simple block model (paragraphs and tables, grouped by slide
 * for presentations), then into sheets:
 *
 *   documents      every table becomes a sheet ("Table 1", ...); the running text goes to column A
 *                  of a "Text" sheet, one paragraph per row
 *   presentations  one sheet per slide: its text in column A, its tables below the text
 *   e-mail         see email.ts
 *
 * Table cells keep their spans (merged cells), bold text and numbers (converted to values).
 *
 * Supported here: DOCX/DOCM/DOTX and Word 2003 XML, ODT/OTT/FODT and SXW, RTF, Word 97-2003 and
 * Word 6/95 (DOC, text with tab-separated rows read as tables), Word for DOS and Windows Write,
 * WordPerfect, Works and other word processors (text only), EPUB, PPTX/PPSX/POTX, ODP/FODP and
 * SXI, PowerPoint 97-2003 (slide text), PDF (text of uncompressed and Flate-compressed pages).
 */
import JSZip from 'jszip';
import { Workbook, Sheet } from '../model/workbook';
import type { Cell, CellStyle } from '../model/types';
import { MAX_COLS, MAX_ROWS } from '../model/address';
import { convertField } from './csv';
import { parseXml, kids, attr, decodeText } from './index';
import { readCfb, extractDoc97, isOle2 } from './recover/doc97';
import { stringsScan, cp1252 } from './recover/salvage';
import { immortalInflate } from './recover/inflate';

// --------------------------------------------------------------------------------------------
// Block model

export interface TCell { text: string; colspan?: number; rowspan?: number; bold?: boolean; covered?: boolean }
export type Block =
  | { kind: 'para'; text: string; bold?: boolean; heading?: boolean }
  | { kind: 'table'; rows: TCell[][] };
export interface Slide { title?: string; blocks: Block[]; notes?: string }

const clean = (s: string) => s.replace(/ /g, ' ').replace(/[ \t]+\n/g, '\n').replace(/\s+$/, '');

/** Puts a cell's text in a sheet: numbers, dates and percentages become values. */
function putText(sh: Sheet, r: number, c: number, text: string, style: CellStyle | null, sid: (st: CellStyle) => number, convert = true): void {
  if (r >= MAX_ROWS || c >= MAX_COLS) return;
  const t = text.length > 32767 ? text.slice(0, 32767) : text;
  const cell: Cell = {};
  const st: CellStyle = { ...(style || {}) };
  if (convert && t && !t.includes('\n') && t.length < 40) {
    const conv = convertField(t, {});
    if (conv.f == null && conv.v != null && typeof conv.v !== 'string') {
      cell.v = conv.v;
      if (conv.fmt && conv.fmt !== 'General') st.numFmt = conv.fmt;
    } else cell.v = t;
  } else cell.v = t;
  if (t.includes('\n')) st.wrap = true;
  const s = sid(st);
  if (s) cell.s = s;
  if (cell.v === '' && !cell.s) return;
  sh.setComputed(r, c, cell);
}

function styler(book: Workbook): (st: CellStyle) => number {
  const cache = new Map<string, number>();
  return (st) => {
    if (!Object.keys(st).length) return 0;
    const k = JSON.stringify(st);
    let id = cache.get(k);
    if (id == null) { id = book.styleId(st); cache.set(k, id); }
    return id;
  };
}

/** Writes a table at (r0, c0); returns the number of rows used. */
function putTable(sh: Sheet, rows: TCell[][], r0: number, c0: number, sid: (st: CellStyle) => number): number {
  // Grid placement with row and column spans.
  const taken = new Set<string>();
  let maxC = 0;
  // Covered cells (spanned positions some formats list explicitly) are skipped; occupancy decides.
  rows.forEach((row, ri) => {
    let c = 0;
    for (const cell of row) {
      if (cell.covered) continue;
      while (taken.has(`${ri},${c}`)) c++;
      const cs = Math.max(1, cell.colspan || 1), rs = Math.max(1, cell.rowspan || 1);
      const style: CellStyle = { vAlign: 'top', bt: { style: 'thin' }, bb: { style: 'thin' }, bl: { style: 'thin' }, br: { style: 'thin' } };
      if (cell.bold || (ri === 0 && rows.length > 1 && row.filter((x) => !x.covered).every((x) => x.bold))) style.bold = true;
      putText(sh, r0 + ri, c0 + c, clean(cell.text), style, sid);
      if (cs > 1 || rs > 1) sh.merges.push({ r1: r0 + ri, c1: c0 + c, r2: r0 + ri + rs - 1, c2: c0 + c + cs - 1 });
      for (let a = 0; a < rs; a++) for (let b = 0; b < cs; b++) taken.add(`${ri + a},${c + b}`);
      c += cs;
      maxC = Math.max(maxC, c);
    }
  });
  // Width from the longest line in each column.
  for (let c = 0; c < maxC; c++) {
    let w = 4;
    for (const row of rows) { const t = row[c]?.text || ''; for (const line of t.split('\n')) w = Math.max(w, Math.min(60, line.length)); }
    const px = Math.round(w * 7 + 10);
    if ((sh.colW.get(c0 + c) || 0) < px) sh.colW.set(c0 + c, px);
  }
  return rows.length;
}

/** Documents: the text on a "Text" sheet, each table on its own sheet. */
export function blocksToBook(blocks: Block[], title: string): Workbook {
  const book = new Workbook(false);
  const sid = styler(book);
  const paras = blocks.filter((b): b is Extract<Block, { kind: 'para' }> => b.kind === 'para');
  const tables = blocks.filter((b): b is Extract<Block, { kind: 'table' }> => b.kind === 'table' && b.rows.length > 0);
  if (paras.some((p) => p.text.trim()) || !tables.length) {
    const sh = new Sheet(book, 'Text');
    sh.colW.set(0, 640);
    let r = 0;
    for (const p of paras) {
      if (r >= MAX_ROWS) break;
      const t = clean(p.text);
      if (!t && r === 0) continue;
      if (t) putText(sh, r, 0, t, p.heading || p.bold ? { bold: true, ...(p.heading ? { size: 13 } : {}) } : null, sid, false);
      r++;
    }
    book.sheets.push(sh);
  }
  tables.forEach((t, i) => {
    const sh = new Sheet(book, `Table ${i + 1}`);
    putTable(sh, t.rows, 0, 0, sid);
    book.sheets.push(sh);
  });
  if (!book.sheets.length) book.sheets.push(new Sheet(book, 'Text'));
  book.props.title = title;
  book.activeSheet = 0;
  return book;
}

/** Presentations: one sheet per slide. */
export function slidesToBook(slides: Slide[], title: string): Workbook {
  const book = new Workbook(false);
  const sid = styler(book);
  slides.forEach((sl, i) => {
    const sh = new Sheet(book, `Slide ${i + 1}`);
    sh.colW.set(0, 480);
    let r = 0;
    if (sl.title) { putText(sh, r++, 0, clean(sl.title), { bold: true, size: 14 }, sid, false); }
    for (const b of sl.blocks) {
      if (r >= MAX_ROWS) break;
      if (b.kind === 'para') {
        const t = clean(b.text);
        if (!t) continue;
        putText(sh, r++, 0, t, b.bold ? { bold: true } : null, sid, false);
      } else {
        if (r) r++;
        r += putTable(sh, b.rows, r, 0, sid) + 1;
      }
    }
    if (sl.notes && sl.notes.trim()) {
      r++;
      putText(sh, r++, 0, 'Notes', { bold: true, italic: true }, sid, false);
      for (const line of clean(sl.notes).split('\n')) if (line.trim()) putText(sh, r++, 0, line, { italic: true }, sid, false);
    }
    book.sheets.push(sh);
  });
  if (!book.sheets.length) book.sheets.push(new Sheet(book, 'Slide 1'));
  book.props.title = title;
  return book;
}

/** Plain text: one paragraph per row; runs of tab-separated lines become tables. */
export function textToBlocks(text: string): Block[] {
  const lines = text.replace(/\r\n?/g, '\n').replace(/\f/g, '\n').split('\n');
  const out: Block[] = [];
  let i = 0;
  while (i < lines.length) {
    if (lines[i].includes('\t')) {
      let j = i;
      while (j < lines.length && lines[j].includes('\t')) j++;
      if (j - i >= 2) {
        out.push({ kind: 'table', rows: lines.slice(i, j).map((l) => l.split('\t').map((t) => ({ text: t.trim() }))) });
        i = j;
        continue;
      }
    }
    out.push({ kind: 'para', text: lines[i] });
    i++;
  }
  return out;
}

// --------------------------------------------------------------------------------------------
// Entry point

export async function importDocument(id: string, data: ArrayBuffer, name: string, warnings: string[]): Promise<Workbook> {
  const title = name.replace(/^.*[\\/]/, '').replace(/\.[^.]+$/, '');
  const u8 = new Uint8Array(data);
  switch (id) {
    case 'eml': case 'mbox': case 'msg': return (await import('./email')).importEmail(id, data, name, warnings);
    case 'docx': return blocksToBook(await docxBlocks(u8, warnings), title);
    case 'odt': return blocksToBook(await odtBlocks(u8), title);
    case 'rtf': return blocksToBook(rtfBlocks(decodeLatin(u8)), title);
    case 'doc': return blocksToBook(docBlocks(u8, warnings), title);
    case 'wordtext': return blocksToBook(await otherWordBlocks(u8, name, warnings), title);
    case 'pptx': return slidesToBook(await pptxSlides(u8), title);
    case 'odp': return slidesToBook(await odpSlides(u8), title);
    case 'ppt': return slidesToBook(pptSlides(u8, warnings), title);
  }
  throw new Error(`Cannot import ${id} documents.`);
}

const decodeLatin = (u8: Uint8Array) => { let s = ''; for (let i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]); return s; };

async function zipText(zip: JSZip, path: string): Promise<string | null> {
  const f = zip.file(path) || zip.file(new RegExp(`^${path.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, 'i'))[0];
  if (!f) return null;
  return decodeText(await f.async('uint8array'));
}

// --------------------------------------------------------------------------------------------
// Word (DOCX, Word 2003 XML)

function wordText(el: Element): string {
  let s = '';
  const walk = (n: Element) => {
    for (const c of Array.from(n.children)) {
      const ln = c.localName;
      if (ln === 't') s += c.textContent || '';
      else if (ln === 'tab') s += '\t';
      else if (ln === 'br' || ln === 'cr') s += '\n';
      else if (ln === 'noBreakHyphen') s += '-';
      else if (ln === 'sym') s += String.fromCharCode(parseInt(attr(c, 'char') || '20', 16) & 0xffff);
      else if (ln === 'delText' || ln === 'instrText' || ln === 'rPr' || ln === 'pPr' || ln === 'fldData') continue;
      else if (ln === 'p') { if (s && !s.endsWith('\n')) s += '\n'; walk(c); }
      else walk(c);
    }
  };
  walk(el);
  return s;
}

function wordBold(p: Element): boolean {
  const runs = Array.from(p.getElementsByTagNameNS('*', 'r')).filter((r) => (r.textContent || '').trim());
  if (!runs.length) return false;
  return runs.every((r) => { const b = Array.from(r.children).find((c) => c.localName === 'rPr')?.getElementsByTagNameNS('*', 'b')[0]; return !!b && attr(b, 'val') !== '0' && attr(b, 'val') !== 'false'; });
}

function wordBlocksFrom(body: Element): Block[] {
  const out: Block[] = [];
  const visit = (el: Element) => {
    for (const c of Array.from(el.children)) {
      const ln = c.localName;
      if (ln === 'p') {
        const style = c.getElementsByTagNameNS('*', 'pStyle')[0];
        const heading = !!style && /heading|title|berschrift|titre/i.test(attr(style, 'val') || '');
        out.push({ kind: 'para', text: wordText(c), bold: wordBold(c), heading });
      } else if (ln === 'tbl') out.push({ kind: 'table', rows: wordTable(c) });
      else if (ln === 'sdt' || ln === 'sdtContent' || ln === 'customXml' || ln === 'smartTag' || ln === 'ins' || ln === 'body' || ln === 'sect' || ln === 'sub-section' || ln === 'section') visit(c);
    }
  };
  visit(body);
  return out;
}

function wordTable(tbl: Element): TCell[][] {
  const rows: TCell[][] = [];
  const vm: Map<number, TCell> = new Map(); // column -> cell that started a vertical merge
  for (const tr of kids(tbl, 'tr')) {
    const row: TCell[] = [];
    let col = 0;
    const cells = Array.from(tr.children).flatMap((c) => (c.localName === 'tc' ? [c] : c.localName === 'sdt' ? Array.from(c.getElementsByTagNameNS('*', 'tc')) : []));
    for (const tc of cells) {
      const pr = Array.from(tc.children).find((x) => x.localName === 'tcPr');
      const span = Number(attr(pr?.getElementsByTagNameNS('*', 'gridSpan')[0], 'val') || 1) || 1;
      const vmEl = pr?.getElementsByTagNameNS('*', 'vMerge')[0];
      const vmVal = vmEl ? attr(vmEl, 'val') || 'continue' : null;
      const parts: string[] = [];
      for (const p of Array.from(tc.children)) {
        if (p.localName === 'p') parts.push(wordText(p));
        else if (p.localName === 'tbl') parts.push(wordTable(p).map((r) => r.map((x) => x.text).join('\t')).join('\n'));
      }
      const text = parts.join('\n').replace(/\n+$/, '');
      if (vmVal === 'continue' && vm.has(col)) {
        const top = vm.get(col)!;
        top.rowspan = (top.rowspan || 1) + 1;
        row.push({ text: '', covered: true, colspan: span });
        for (let k = 1; k < span; k++) row.push({ text: '', covered: true });
      } else {
        const cell: TCell = { text, ...(span > 1 ? { colspan: span } : {}), bold: tc.getElementsByTagNameNS('*', 'p').length > 0 && Array.from(tc.getElementsByTagNameNS('*', 'p')).every(wordBold) && !!text.trim() };
        row.push(cell);
        for (let k = 1; k < span; k++) row.push({ text: '', covered: true });
        if (vmVal === 'restart') vm.set(col, cell); else vm.delete(col);
      }
      col += span;
    }
    rows.push(row);
  }
  return rows;
}

async function docxBlocks(u8: Uint8Array, warnings: string[]): Promise<Block[]> {
  if (isOle2(u8)) throw new Error('This Word document is encrypted with a password; remove the password in Word and open it again.');
  if (u8[0] === 0x3c || u8[0] === 0xef || u8[0] === 0xff) {
    const doc = parseXml(decodeText(u8), 'Word XML');
    const body = doc.getElementsByTagNameNS('*', 'body')[0];
    return body ? wordBlocksFrom(body) : [];
  }
  const zip = await JSZip.loadAsync(u8);
  const main = await zipText(zip, 'word/document.xml');
  if (main == null) throw new Error('The document part (word/document.xml) is missing.');
  const doc = parseXml(main, 'word/document.xml');
  const body = doc.getElementsByTagNameNS('*', 'body')[0];
  const blocks = body ? wordBlocksFrom(body) : [];
  // Footnotes and endnotes follow the text.
  for (const part of ['word/footnotes.xml', 'word/endnotes.xml']) {
    const t = await zipText(zip, part);
    if (!t) continue;
    const d = parseXml(t, part);
    const notes = Array.from(d.getElementsByTagNameNS('*', part.includes('foot') ? 'footnote' : 'endnote')).filter((n) => !/separator|continuation/i.test(attr(n, 'type') || ''));
    if (notes.length) {
      blocks.push({ kind: 'para', text: '' }, { kind: 'para', text: part.includes('foot') ? 'Footnotes' : 'Endnotes', bold: true });
      for (const n of notes) blocks.push({ kind: 'para', text: wordText(n) });
    }
  }
  void warnings;
  return blocks;
}

// --------------------------------------------------------------------------------------------
// OpenDocument (ODT, FODT, SXW) and presentations (ODP, FODP, SXI)

function odfText(el: Element): string {
  let s = '';
  const walk = (n: Node) => {
    if (n.nodeType === 3) { s += (n.nodeValue || ''); return; }
    if (n.nodeType !== 1) return;
    const e = n as Element, ln = e.localName;
    if (ln === 's') { s += ' '.repeat(Number(attr(e, 'c') || 1) || 1); return; }
    if (ln === 'tab') { s += '\t'; return; }
    if (ln === 'line-break') { s += '\n'; return; }
    if (ln === 'note' || ln === 'annotation' || ln === 'tracked-changes' || ln === 'soft-page-break') return;
    if ((ln === 'p' || ln === 'h') && s && !s.endsWith('\n')) s += '\n';
    e.childNodes.forEach(walk);
  };
  walk(el);
  return s;
}

function odfTable(t: Element): TCell[][] {
  const rows: TCell[][] = [];
  const rowEls: Element[] = [];
  const collect = (el: Element) => {
    for (const c of Array.from(el.children)) {
      if (c.localName === 'table-row') { const n = Math.min(1000, Number(attr(c, 'number-rows-repeated') || 1) || 1); for (let k = 0; k < n; k++) rowEls.push(c); }
      else if (/^table-(header-)?rows$|^table-row-group$/.test(c.localName)) collect(c);
    }
  };
  collect(t);
  for (const tr of rowEls) {
    const row: TCell[] = [];
    for (const c of Array.from(tr.children)) {
      const n = Math.min(256, Number(attr(c, 'number-columns-repeated') || 1) || 1);
      for (let k = 0; k < n; k++) {
        if (c.localName === 'covered-table-cell') row.push({ text: '', covered: true });
        else if (c.localName === 'table-cell') {
          const cs = Number(attr(c, 'number-columns-spanned') || 1) || 1, rs = Number(attr(c, 'number-rows-spanned') || 1) || 1;
          const parts = Array.from(c.children).map((x) => (x.localName === 'table' ? odfTable(x).map((r) => r.map((y) => y.text).join('\t')).join('\n') : odfText(x)));
          row.push({ text: parts.join('\n'), ...(cs > 1 ? { colspan: cs } : {}), ...(rs > 1 ? { rowspan: rs } : {}) });
        }
      }
    }
    // Trailing empty repeated cells are padding.
    while (row.length && !row[row.length - 1].text && !row[row.length - 1].covered) row.pop();
    rows.push(row);
  }
  while (rows.length && !rows[rows.length - 1].some((c) => c.text)) rows.pop();
  return rows;
}

function odfBlocks(el: Element, out: Block[]): void {
  for (const c of Array.from(el.children)) {
    const ln = c.localName;
    if (ln === 'p' || ln === 'h') out.push({ kind: 'para', text: odfText(c), heading: ln === 'h' });
    else if (ln === 'table') out.push({ kind: 'table', rows: odfTable(c) });
    else if (ln === 'list' || ln === 'list-item' || ln === 'list-header' || ln === 'section' || ln === 'text' || ln === 'body' || ln === 'index-body'
      || /-index$/.test(ln) || ln === 'frame' || ln === 'text-box') odfBlocks(c, out);
  }
}

async function odfContent(u8: Uint8Array): Promise<Document> {
  if (u8[0] !== 0x50) return parseXml(decodeText(u8), 'flat OpenDocument');
  const zip = await JSZip.loadAsync(u8);
  const t = await zipText(zip, 'content.xml');
  if (t == null) throw new Error('The document content (content.xml) is missing.');
  return parseXml(t, 'content.xml');
}

async function odtBlocks(u8: Uint8Array): Promise<Block[]> {
  const doc = await odfContent(u8);
  const body = doc.getElementsByTagNameNS('*', 'body')[0];
  const text = body ? Array.from(body.children).find((c) => c.localName === 'text') || body : doc.documentElement;
  const out: Block[] = [];
  odfBlocks(text, out);
  return out;
}

async function odpSlides(u8: Uint8Array): Promise<Slide[]> {
  const doc = await odfContent(u8);
  const pages = Array.from(doc.getElementsByTagNameNS('*', 'page')).filter((p) => p.parentElement?.localName === 'presentation');
  return pages.map((pg) => {
    const blocks: Block[] = [];
    let title: string | undefined;
    let notes = '';
    const visit = (el: Element) => {
      for (const c of Array.from(el.children)) {
        const ln = c.localName;
        if (ln === 'notes') {
          // The notes page also holds a picture of the slide and page-number fields; keep the notes frame.
          const frames = Array.from(c.getElementsByTagNameNS('*', 'frame')).filter((f) => attr(f, 'class') === 'notes');
          notes = frames.flatMap((f) => Array.from(f.getElementsByTagNameNS('*', 'p')).map(odfText)).join('\n');
          continue;
        }
        if (ln === 'table') { blocks.push({ kind: 'table', rows: odfTable(c) }); continue; }
        if (ln === 'frame' && /title/.test(attr(c, 'class') || '') && title == null) { title = Array.from(c.getElementsByTagNameNS('*', 'p')).map(odfText).join('\n'); continue; }
        if (ln === 'p' || ln === 'h') { blocks.push({ kind: 'para', text: odfText(c) }); continue; }
        visit(c);
      }
    };
    visit(pg);
    return { title, blocks, notes };
  });
}

// --------------------------------------------------------------------------------------------
// PowerPoint (PPTX)

function drawingText(p: Element): string {
  let s = '';
  for (const c of Array.from(p.getElementsByTagNameNS('*', '*'))) {
    if (c.localName === 't' && c.parentElement?.localName === 'r' || c.localName === 't' && c.parentElement?.localName === 'fld') s += c.textContent || '';
    else if (c.localName === 'br') s += '\n';
  }
  return s;
}

function dmlTable(tbl: Element): TCell[][] {
  return kids(tbl, 'tr').map((tr) => kids(tr, 'tc').map((tc) => {
    const gs = Number(attr(tc, 'gridSpan') || 1) || 1, rs = Number(attr(tc, 'rowSpan') || 1) || 1;
    const covered = attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1';
    const text = Array.from(tc.getElementsByTagNameNS('*', 'p')).map(drawingText).join('\n');
    return covered ? { text: '', covered: true } : { text, ...(gs > 1 ? { colspan: gs } : {}), ...(rs > 1 ? { rowspan: rs } : {}) };
  }));
}

async function pptxSlides(u8: Uint8Array): Promise<Slide[]> {
  const zip = await JSZip.loadAsync(u8);
  const pres = await zipText(zip, 'ppt/presentation.xml');
  const rels = await zipText(zip, 'ppt/_rels/presentation.xml.rels');
  let paths: string[] = [];
  if (pres && rels) {
    const rd = parseXml(rels, 'presentation rels');
    const map = new Map(Array.from(rd.getElementsByTagNameNS('*', 'Relationship')).map((r) => [attr(r, 'Id'), attr(r, 'Target') || '']));
    const pd = parseXml(pres, 'presentation.xml');
    paths = Array.from(pd.getElementsByTagNameNS('*', 'sldId')).map((s) => {
      const id = Array.from(s.attributes).find((a) => a.localName === 'id' && (a.prefix === 'r' || /relationships/.test(a.namespaceURI || '')))?.value;
      const t = id ? map.get(id) || '' : '';
      return t.startsWith('/') ? t.slice(1) : `ppt/${t.replace(/^\.\//, '')}`;
    }).filter(Boolean);
  }
  if (!paths.length) {
    paths = Object.keys(zip.files).filter((p) => /^ppt\/slides\/slide\d+\.xml$/i.test(p)).sort((a, b) => Number(/(\d+)\.xml/.exec(a)![1]) - Number(/(\d+)\.xml/.exec(b)![1]));
  }
  const slides: Slide[] = [];
  for (const p of paths) {
    const t = await zipText(zip, p);
    if (!t) continue;
    const doc = parseXml(t, p);
    const blocks: Block[] = [];
    let title: string | undefined;
    const tree = doc.getElementsByTagNameNS('*', 'spTree')[0];
    const visit = (el: Element) => {
      for (const c of Array.from(el.children)) {
        const ln = c.localName;
        if (ln === 'sp') {
          const ph = c.getElementsByTagNameNS('*', 'ph')[0];
          const isTitle = !!ph && /title/i.test(attr(ph, 'type') || '');
          const paras = Array.from(c.getElementsByTagNameNS('*', 'txBody')[0]?.children || []).filter((x) => x.localName === 'p').map(drawingText);
          if (isTitle && title == null) title = paras.join('\n');
          else for (const q of paras) blocks.push({ kind: 'para', text: q });
        } else if (ln === 'graphicFrame') {
          const tbl = c.getElementsByTagNameNS('*', 'tbl')[0];
          if (tbl) blocks.push({ kind: 'table', rows: dmlTable(tbl) });
        } else if (ln === 'grpSp') visit(c);
      }
    };
    if (tree) visit(tree);
    // Speaker notes
    let notes = '';
    const relsText = await zipText(zip, p.replace(/([^/]+)$/, '_rels/$1.rels'));
    if (relsText) {
      const nr = Array.from(parseXml(relsText, 'slide rels').getElementsByTagNameNS('*', 'Relationship')).find((r) => /notesSlide/.test(attr(r, 'Type') || ''));
      if (nr) {
        const np = new URL(attr(nr, 'Target') || '', `http://x/${p}`).pathname.slice(1);
        const nt = await zipText(zip, np);
        if (nt) {
          const nd = parseXml(nt, np);
          notes = Array.from(nd.getElementsByTagNameNS('*', 'sp')).filter((s) => /body/.test(attr(s.getElementsByTagNameNS('*', 'ph')[0], 'type') || ''))
            .flatMap((s) => Array.from(s.getElementsByTagNameNS('*', 'p')).map(drawingText)).join('\n');
        }
      }
    }
    slides.push({ title, blocks, notes });
  }
  return slides;
}

// --------------------------------------------------------------------------------------------
// PowerPoint 97-2003 (PPT): slide text from the SlideListWithText records

function pptSlides(u8: Uint8Array, warnings: string[]): Slide[] {
  const cfb = readCfb(u8);
  const s = cfb?.read?.('PowerPoint Document') || cfb?.streams.get('PowerPoint Document');
  if (!s) throw new Error('The presentation stream was not found.');
  const u16 = (o: number) => s[o] | (s[o + 1] << 8);
  const u32 = (o: number) => (s[o] | (s[o + 1] << 8) | (s[o + 2] << 16) | (s[o + 3] << 24)) >>> 0;
  const slides: Slide[] = [];
  let cur: Slide | null = null;
  let textType = -1;
  const loose: string[] = [];
  const walk = (start: number, end: number, inSlwt: boolean, depth: number) => {
    let p = start;
    while (p + 8 <= end && depth < 20) {
      const verInst = u16(p), type = u16(p + 2), len = u32(p + 4);
      const body = p + 8, next = body + len;
      if (next > end || next < body) break;
      const container = (verInst & 0xf) === 0xf;
      const inst = verInst >> 4;
      if (container) {
        walk(body, next, inSlwt || (type === 0x0ff0 && inst === 0), depth + 1);
      } else if (type === 0x03f3 && inSlwt) { cur = { blocks: [] }; slides.push(cur); }
      else if (type === 0x0f9f) textType = u32(body);
      else if (type === 0x0fa0 || type === 0x0fa8) {
        let t = '';
        if (type === 0x0fa0) for (let k = body; k + 1 < next; k += 2) t += String.fromCharCode(u16(k));
        else for (let k = body; k < next; k++) t += cp1252(s[k]);
        t = t.replace(/\r/g, '\n').replace(/\v/g, '\n');
        if (inSlwt && cur) {
          const sl: Slide = cur;
          if ((textType === 0 || textType === 6) && sl.title == null) sl.title = t;
          else for (const line of t.split('\n')) sl.blocks.push({ kind: 'para', text: line });
        } else if (!inSlwt) loose.push(t);
      }
      p = next;
    }
  };
  walk(0, s.length, false, 0);
  // Programs other than PowerPoint often keep slide text only in the slides' drawings.
  const hasText = (sl: Slide) => !!sl.title || sl.blocks.some((b) => b.kind === 'para' && b.text.trim());
  if (!slides.some(hasText)) {
    const fromDrawings: Slide[] = [];
    const scan = (start: number, end: number, slide: Slide | null, depth: number) => {
      let p = start;
      while (p + 8 <= end && depth < 24) {
        const verInst = u16(p), type = u16(p + 2), len = u32(p + 4);
        const body = p + 8, next = body + len;
        if (next > end || next < body) break;
        if ((verInst & 0xf) === 0xf) {
          if (type === 0x03ee) { const sl: Slide = { blocks: [] }; fromDrawings.push(sl); scan(body, next, sl, depth + 1); }
          else if (type !== 0x0ff0 && type !== 0x03f8 && type !== 0x0ff2) scan(body, next, slide, depth + 1); // skip slide lists, masters and handouts
        } else if (slide && (type === 0x0fa0 || type === 0x0fa8)) {
          let t = '';
          if (type === 0x0fa0) for (let k = body; k + 1 < next; k += 2) t += String.fromCharCode(u16(k));
          else for (let k = body; k < next; k++) t += cp1252(s[k]);
          const lines = t.replace(/[\r\v]/g, '\n').split('\n');
          if (slide.title == null && !slide.blocks.length) slide.title = lines.shift();
          for (const line of lines) slide.blocks.push({ kind: 'para', text: line });
        }
        p = next;
      }
    };
    scan(0, s.length, null, 0);
    if (fromDrawings.some(hasText)) { slides.length = 0; slides.push(...fromDrawings); }
  }
  if (!slides.length) {
    warnings.push('The slide list was not found; the presentation\'s text is shown on one sheet.');
    return [{ blocks: loose.flatMap((t) => t.split('\n')).map((text) => ({ kind: 'para' as const, text })) }];
  }
  warnings.push('PowerPoint 97-2003 slides are read for the text of their placeholders; text boxes and tables drawn on slides may be missing.');
  return slides;
}

// --------------------------------------------------------------------------------------------
// RTF: paragraphs and tables (\trowd ... \cell ... \row)

export function rtfBlocks(src: string): Block[] {
  const out: Block[] = [];
  let para = '', cell = '', row: TCell[] = [], table: TCell[][] | null = null, intbl = false;
  let codepage = 1252, uc = 1, skipChars = 0;
  const stack: { skip: boolean; uc: number }[] = [];
  let skip = false;
  const SKIP = new Set(['fonttbl', 'colortbl', 'stylesheet', 'info', 'pict', 'header', 'footer', 'headerl', 'headerr', 'headerf', 'footerl', 'footerr', 'footerf',
    'listtable', 'listoverridetable', 'rsidtbl', 'generator', 'xmlnstbl', 'themedata', 'colorschememapping', 'latentstyles', 'datastore', 'object', 'fldinst',
    'bkmkstart', 'bkmkend', 'revtbl', 'pgdsctbl', 'footnote', 'annotation', 'nonshppict', 'filetbl', 'mmathPr', 'wgrffmtfilter', 'xe', 'tc']);
  const dec1252 = (b: number) => (codepage === 1252 ? cp1252(b) : new TextDecoder(`windows-${codepage}`, { fatal: false }).decode(new Uint8Array([b])));
  const emit = (t: string) => { if (skip) return; if (skipChars > 0) { const n = Math.min(skipChars, t.length); skipChars -= n; t = t.slice(n); } if (intbl) cell += t; else para += t; };
  const endPara = () => {
    if (intbl) { cell += '\n'; return; }
    if (table) { out.push({ kind: 'table', rows: table }); table = null; }
    out.push({ kind: 'para', text: para });
    para = '';
  };
  let i = 0;
  const n = src.length;
  while (i < n) {
    const ch = src[i];
    if (ch === '{') { stack.push({ skip, uc }); i++; if (src.startsWith('\\*', i)) { skip = true; i += 2; } continue; }
    if (ch === '}') { const st = stack.pop(); if (st) { skip = st.skip; uc = st.uc; } i++; continue; }
    if (ch === '\r' || ch === '\n') { i++; continue; }
    if (ch !== '\\') { emit(ch); i++; continue; }
    const nx = src[i + 1];
    if (nx === '\\' || nx === '{' || nx === '}') { emit(nx); i += 2; continue; }
    if (nx === "'") { const b = parseInt(src.substr(i + 2, 2), 16); if (!Number.isNaN(b)) emit(dec1252(b)); i += 4; continue; }
    if (nx === '~') { emit(' '); i += 2; continue; }
    if (nx === '-' || nx === ':') { i += 2; continue; }
    if (nx === '_') { emit('-'); i += 2; continue; }
    if (nx === '\n' || nx === '\r') { endPara(); i += 2; continue; }
    const m = /^\\([a-zA-Z]+)(-?\d+)? ?/.exec(src.slice(i, i + 40));
    if (!m) { i++; continue; }
    i += m[0].length;
    const w = m[1], arg = m[2] != null ? Number(m[2]) : null;
    if (SKIP.has(w)) { skip = true; continue; }
    switch (w) {
      case 'ansicpg': if (arg) codepage = arg; break;
      case 'uc': uc = arg ?? 1; break;
      case 'u': { let c = arg ?? 0; if (c < 0) c += 65536; emit(String.fromCharCode(c)); skipChars = uc; break; }
      case 'par': case 'sect': endPara(); break;
      case 'line': emit('\n'); break;
      case 'tab': emit('\t'); break;
      case 'emdash': emit('—'); break; case 'endash': emit('–'); break;
      case 'lquote': emit('‘'); break; case 'rquote': emit('’'); break; case 'ldblquote': emit('“'); break; case 'rdblquote': emit('”'); break;
      case 'bullet': emit('•'); break;
      case 'intbl': intbl = true; break;
      case 'pard': intbl = false; break;
      case 'trowd': if (para.trim()) { out.push({ kind: 'para', text: para }); } para = ''; intbl = true; break;
      case 'cell': row.push({ text: cell.replace(/\n+$/, '') }); cell = ''; break;
      case 'nestcell': cell += '\t'; break;
      case 'row': (table ??= []).push(row); row = []; intbl = false; break;
    }
  }
  if (table) out.push({ kind: 'table', rows: table });
  if (para) out.push({ kind: 'para', text: para });
  return out;
}

// --------------------------------------------------------------------------------------------
// Word 97-2003 and older word processors

function docBlocks(u8: Uint8Array, warnings: string[]): Block[] {
  if (isOle2(u8)) {
    const r = extractDoc97(u8);
    if (!r) throw new Error('The Word document could not be read.');
    if (r.encrypted) throw new Error('This Word document is encrypted with a password; remove the password in Word and open it again.');
    const blocks = textToBlocks(r.main);
    if (r.footnotes) blocks.push({ kind: 'para', text: '' }, { kind: 'para', text: 'Footnotes', bold: true }, ...textToBlocks(r.footnotes));
    warnings.push('Word 97-2003 documents are read for their text; table rows become sheet rows and tab-separated lines are treated as tables.');
    return blocks;
  }
  return writeBlocks(u8);
}

/** Word for DOS and Windows Write: the text runs from byte 128 to the end of text (fcMac). */
function writeBlocks(u8: Uint8Array): Block[] {
  const fcMac = (u8[14] | (u8[15] << 8) | (u8[16] << 16) | (u8[17] << 24)) >>> 0;
  const end = fcMac > 128 && fcMac <= u8.length ? fcMac : u8.length;
  let s = '';
  for (let i = 128; i < end; i++) { const b = u8[i]; s += b === 0x0d ? '' : b < 0x20 && b !== 0x0a && b !== 0x09 ? '' : cp1252(b); }
  return textToBlocks(s);
}

/** WordPerfect 5 and 6: printable text of the document area, with hard returns as line ends. */
function wordPerfectText(u8: Uint8Array): string {
  const start = Math.min(u8.length, (u8[4] | (u8[5] << 8) | (u8[6] << 16) | (u8[7] << 24)) >>> 0 || 16);
  let s = '';
  for (let i = start; i < u8.length;) {
    const b = u8[i];
    if (b >= 0x20 && b < 0x7f) { s += String.fromCharCode(b); i++; continue; }
    if (b === 0x0a || b === 0x0d || b === 0xcc || b === 0xcf) { s += '\n'; i++; continue; }
    if (b === 0x09 || b === 0xc1) { s += '\t'; i++; continue; }
    if (b >= 0xd0 && b <= 0xfe) {
      // Variable-length function: code, subcode, length, data, length, code.
      const len = u8[i + 2] | (u8[i + 3] << 8);
      if (len > 0 && i + 4 + len <= u8.length && u8[i + 3 + len] === b) { if (b === 0xd0 || b === 0xd1) s += '\n'; i += 4 + len; continue; }
    }
    if (b === 0xc0 || b === 0xf0) { i += 4; s += '?'; continue; }
    if (b >= 0x80 && b < 0xc0) { s += ' '; i++; continue; }
    i++;
  }
  return s.replace(/[ \t]*\n/g, '\n').replace(/\n{3,}/g, '\n\n');
}

async function otherWordBlocks(u8: Uint8Array, name: string, warnings: string[]): Promise<Block[]> {
  const head = decodeLatin(u8.subarray(0, 512));
  if (head.startsWith('%PDF-')) { warnings.push('PDF files are read for their text only; tables are recognized from tab-separated lines.'); return textToBlocks(pdfText(u8)); }
  if (head.startsWith('\xffWPC')) { warnings.push('WordPerfect documents are read for their text only.'); return textToBlocks(wordPerfectText(u8)); }
  if (/^\{\\rtf/.test(head)) return rtfBlocks(decodeLatin(u8));
  if ((u8[0] === 0x31 || u8[0] === 0x32) && u8[1] === 0xbe) return writeBlocks(u8);
  if (/<w:wordDocument\b|progid="Word\.Document"/i.test(head)) return docxBlocks(u8, warnings);
  if (u8[0] === 0x50 && u8[1] === 0x4b) {
    const zip = await JSZip.loadAsync(u8);
    if (zip.file('content.xml')) return odtBlocks(u8);
    // EPUB: XHTML chapters in reading order.
    const opfPath = Object.keys(zip.files).find((p) => p.toLowerCase().endsWith('.opf'));
    if (opfPath) {
      const opf = parseXml((await zipText(zip, opfPath)) || '', 'package.opf');
      const items = new Map(Array.from(opf.getElementsByTagNameNS('*', 'item')).map((it) => [attr(it, 'id'), attr(it, 'href') || '']));
      const base = opfPath.includes('/') ? opfPath.slice(0, opfPath.lastIndexOf('/') + 1) : '';
      const blocks: Block[] = [];
      for (const ref of Array.from(opf.getElementsByTagNameNS('*', 'itemref'))) {
        const href = items.get(attr(ref, 'idref'));
        if (!href) continue;
        const t = await zipText(zip, decodeURIComponent(base + href));
        if (t) blocks.push(...htmlBlocks(t));
      }
      return blocks;
    }
    const idx = (await zipText(zip, 'index.xml')) || '';
    if (idx) return textToBlocks(idx.replace(/<sf:br\/>/g, '\n').replace(/<\/sf:p>/g, '\n').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>'));
    warnings.push('Only loose text could be recovered from this document.');
  }
  if (/<abiword\b/.test(head)) {
    const d = parseXml(decodeText(u8), 'AbiWord');
    return Array.from(d.getElementsByTagNameNS('*', 'p')).map((p) => ({ kind: 'para' as const, text: p.textContent || '' }));
  }
  if (isOle2(u8)) {
    const cfb = readCfb(u8);
    const parts: string[] = [];
    for (const p of cfb?.paths || []) { const st = cfb?.read?.(p); if (st && st.length > 64) parts.push(stringsScan(st, 4)); }
    warnings.push(`${name.replace(/^.*[\\/]/, '')} is read for its readable text only.`);
    return textToBlocks(parts.join('\n'));
  }
  warnings.push(`${name.replace(/^.*[\\/]/, '')} is read for its readable text only.`);
  return textToBlocks(stringsScan(u8, 4));
}

function htmlBlocks(html: string): Block[] {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const out: Block[] = [];
  const visit = (el: Element) => {
    for (const c of Array.from(el.children)) {
      const tag = c.tagName.toUpperCase();
      if (tag === 'TABLE') {
        out.push({ kind: 'table', rows: Array.from((c as HTMLTableElement).rows || []).map((tr) => Array.from(tr.cells).map((td) => ({ text: (td.textContent || '').trim(), colspan: td.colSpan > 1 ? td.colSpan : undefined, rowspan: td.rowSpan > 1 ? td.rowSpan : undefined, bold: td.tagName === 'TH' }))) });
      } else if (/^(P|H[1-6]|LI|PRE|BLOCKQUOTE|DT|DD|CAPTION)$/.test(tag)) out.push({ kind: 'para', text: (c.textContent || '').replace(/\s+/g, ' ').trim(), heading: /^H/.test(tag) });
      else if (tag !== 'SCRIPT' && tag !== 'STYLE') visit(c);
    }
  };
  visit(doc.body || doc.documentElement);
  return out;
}

// --------------------------------------------------------------------------------------------
// PDF: text operators of each content stream

export function pdfText(u8: Uint8Array): string {
  const src = decodeLatin(u8);
  const out: string[] = [];
  const re = /stream\r?\n/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(src))) {
    const dictStart = src.lastIndexOf('<<', m.index);
    const dict = src.slice(dictStart, m.index);
    const start = m.index + m[0].length;
    const end = src.indexOf('endstream', start);
    if (end < 0) break;
    if (/\/(Image|XObject\s*\/Subtype\s*\/Image|FontFile|Length1)/.test(dict) && !/\/Contents/.test(dict)) { re.lastIndex = end; continue; }
    let data = u8.subarray(start, end);
    if (/\/FlateDecode/.test(dict)) {
      try { data = immortalInflate(data.subarray(2), {}).data; } catch { re.lastIndex = end; continue; }
    } else if (/\/Filter/.test(dict)) { re.lastIndex = end; continue; }
    const content = decodeLatin(data);
    if (!/\bBT\b/.test(content)) { re.lastIndex = end; continue; }
    out.push(pdfContentText(content));
    re.lastIndex = end;
  }
  return out.join('\n\f\n');
}

function pdfContentText(c: string): string {
  let s = '';
  const str = (lit: string) => lit.replace(/\\(\d{1,3}|.)/g, (_x, e: string) => (/^\d/.test(e) ? String.fromCharCode(parseInt(e, 8)) : e === 'n' ? '\n' : e === 'r' ? '' : e === 't' ? '\t' : e));
  const tok = /\((?:\\.|[^\\)])*\)|<[0-9A-Fa-f\s]*>|\[|\]|-?\d*\.?\d+|\/[^\s/<>()[\]]+|[A-Za-z'"*]+/g;
  let m: RegExpExecArray | null;
  let lastNum = 0;
  while ((m = tok.exec(c))) {
    const t = m[0];
    if (t[0] === '(') { s += str(t.slice(1, -1)); continue; }
    if (t[0] === '<' && t.length > 2) { const hex = t.slice(1, -1).replace(/\s/g, ''); let x = ''; for (let k = 0; k + 1 < hex.length; k += 2) x += String.fromCharCode(parseInt(hex.substr(k, 2), 16)); s += x.replace(/\0/g, ''); continue; }
    if (/^-?\d*\.?\d+$/.test(t)) { lastNum = Number(t); if (lastNum < -200) s += ' '; continue; }
    if (t === 'Td' || t === 'TD' || t === 'T*' || t === "'" || t === '"') { if (!s.endsWith('\n')) s += '\n'; continue; }
    if (t === 'ET') { if (!s.endsWith('\n')) s += '\n'; continue; }
  }
  return s;
}
