/**
 * HTML table import/export and Markdown table export.
 */
import { Workbook, Sheet } from '../model/workbook';
import { CellError, type Cell, type CellStyle, type Border } from '../model/types';
import { cellKey, keyRow, keyCol, type Range } from '../model/address';
import { formatValue } from '../model/numfmt';
import { convertField } from './csv';
import { normColor } from './index';

// --------------------------------------------------------------------------------------------
// Import

export interface HtmlImportOptions { allTables?: boolean }

function cssBorder(v: string | undefined): Border | undefined {
  if (!v || /\bnone\b|^0(px)?\b|hidden/.test(v)) return undefined;
  const w = parseFloat(/([\d.]+)px/.exec(v)?.[1] || /([\d.]+)pt/.exec(v)?.[1] || '1');
  const style: Border['style'] = /double/.test(v) ? 'double' : /dashed/.test(v) ? 'dashed' : /dotted/.test(v) ? 'dotted' : w >= 3 ? 'thick' : w >= 2 ? 'medium' : 'thin';
  const color = normColor(/(#[0-9a-f]{3,6}|rgba?\([^)]*\)|\b[a-z]+\b)\s*$/i.exec(v.trim())?.[1] || '') || undefined;
  return { style, ...(color ? { color } : {}) };
}

function styleFromElement(el: HTMLElement): CellStyle {
  const st: CellStyle = {};
  const s = el.style;
  const fw = s.fontWeight;
  if (fw === 'bold' || +fw >= 600 || el.tagName === 'TH') st.bold = true;
  if (s.fontStyle === 'italic') st.italic = true;
  if (/underline/.test(s.textDecoration || '')) st.underline = 'single';
  if (/line-through/.test(s.textDecoration || '')) st.strike = true;
  const color = normColor(s.color);
  if (color && color !== '#000000') st.color = color;
  const bg = normColor(s.backgroundColor || el.getAttribute('bgcolor') || '');
  if (bg && bg !== '#ffffff') st.fill = bg;
  const ta = (s.textAlign || el.getAttribute('align') || '').toLowerCase();
  if (ta === 'center' || ta === 'right' || ta === 'left' || ta === 'justify') st.hAlign = ta;
  const va = (s.verticalAlign || el.getAttribute('valign') || '').toLowerCase();
  if (va === 'top' || va === 'bottom') st.vAlign = va; else if (va === 'middle') st.vAlign = 'middle';
  const fs = s.fontSize;
  if (fs) { const m = /([\d.]+)(px|pt)/.exec(fs); if (m) st.size = Math.round((m[2] === 'px' ? +m[1] * 0.75 : +m[1]) * 2) / 2; }
  if (s.fontFamily) st.font = s.fontFamily.split(',')[0].replace(/["']/g, '').trim();
  if (s.whiteSpace === 'normal' || s.whiteSpace === 'pre-wrap') st.wrap = true;
  const all = cssBorder(s.border);
  const bt = cssBorder(s.borderTop) || all, bb = cssBorder(s.borderBottom) || all, bl = cssBorder(s.borderLeft) || all, br = cssBorder(s.borderRight) || all;
  if (bt) st.bt = bt; if (bb) st.bb = bb; if (bl) st.bl = bl; if (br) st.br = br;
  // Inline wrappers that cover the whole cell
  const only = el.children.length === 1 ? el.children[0] : null;
  if (only && only.textContent === el.textContent) {
    if (/^(B|STRONG)$/.test(only.tagName)) st.bold = true;
    if (/^(I|EM)$/.test(only.tagName)) st.italic = true;
    if (only.tagName === 'U') st.underline = 'single';
  }
  const fmt = el.getAttribute('x:num') != null ? null : (el.style.getPropertyValue('mso-number-format') || '').replace(/^["']|["']$/g, '').replace(/\\(.)/g, '$1');
  if (fmt && fmt !== 'General') st.numFmt = fmt === '\\@' || fmt === '@' ? '@' : fmt;
  return st;
}

function cellText(el: HTMLElement): string {
  let out = '';
  const walk = (n: Node) => {
    if (n.nodeType === 3) { out += (n.nodeValue || '').replace(/[ \t\r\n]+/g, ' '); return; }
    if (n.nodeType !== 1) return;
    const tag = (n as Element).tagName;
    if (tag === 'BR') { out += '\n'; return; }
    if (tag === 'STYLE' || tag === 'SCRIPT') return;
    const block = /^(P|DIV|LI|TR|H[1-6])$/.test(tag);
    if (block && out && !out.endsWith('\n')) out += '\n';
    n.childNodes.forEach(walk);
    if (block && !out.endsWith('\n')) out += '\n';
  };
  el.childNodes.forEach(walk);
  return out.replace(/ *\n */g, '\n').replace(/^\s+|\s+$/g, '');
}

export function importHtml(text: string, warnings: string[] = [], opts: HtmlImportOptions = { allTables: true }): Workbook {
  const doc = new DOMParser().parseFromString(text, 'text/html');
  const book = new Workbook(false);
  let tables = Array.from(doc.querySelectorAll('table')).filter((t) => !t.parentElement?.closest('table'));
  if (!opts.allTables) tables = tables.slice(0, 1);
  const styleCache = new Map<string, number>();
  const sid = (st: CellStyle) => { const k = JSON.stringify(st); let s = styleCache.get(k); if (s == null) { s = book.styleId(st); styleCache.set(k, s); } return s; };
  tables.forEach((table, ti) => {
    const cap = table.querySelector('caption')?.textContent?.trim();
    let name = (cap || (tables.length > 1 ? `Table${ti + 1}` : 'Sheet1')).replace(/[\\/?*[\]:]/g, '_').slice(0, 31);
    while (book.sheetByName(name)) name = name.slice(0, 28) + `_${ti + 1}`;
    const sheet = new Sheet(book, name);
    book.sheets.push(sheet);
    const occupied = new Set<number>();
    const rows = Array.from(table.rows);
    rows.forEach((tr, r) => {
      let c = 0;
      const h = parseFloat(tr.style.height || tr.getAttribute('height') || '');
      if (h > 0) sheet.rowH.set(r, Math.round(/pt/.test(tr.style.height) ? (h * 4) / 3 : h));
      for (const td of Array.from(tr.cells)) {
        while (occupied.has(cellKey(r, c))) c++;
        const rs = Math.max(1, Math.min(1000, parseInt(td.getAttribute('rowspan') || '1', 10) || 1));
        const cs = Math.max(1, Math.min(1000, parseInt(td.getAttribute('colspan') || '1', 10) || 1));
        for (let i = 0; i < rs; i++) for (let j = 0; j < cs; j++) occupied.add(cellKey(r + i, c + j));
        if (rs > 1 || cs > 1) sheet.merges.push({ r1: r, c1: c, r2: r + rs - 1, c2: c + cs - 1 });
        const st = styleFromElement(td);
        const txt = cellText(td);
        const xnum = td.getAttribute('x:num');
        let conv: { v: Cell['v']; f?: string; fmt?: string };
        if (xnum != null && xnum !== '' && Number.isFinite(Number(xnum))) conv = { v: Number(xnum) };
        else if (st.numFmt === '@' || td.getAttribute('x:str') != null) conv = { v: txt || null };
        else conv = txt.includes('\n') ? { v: txt } : convertField(txt);
        if (conv.f) conv = { v: txt };
        if (conv.fmt && !st.numFmt) st.numFmt = conv.fmt;
        const cell: Cell = {};
        if (conv.v != null && conv.v !== '') cell.v = conv.v;
        if (txt.includes('\n')) st.wrap = true;
        const s = sid(st);
        if (s) cell.s = s;
        const a = td.querySelector('a[href]');
        const href = a?.getAttribute('href');
        if (href && !/^\s*javascript:/i.test(href)) cell.link = href;
        if (cell.v != null || cell.s || cell.link) sheet.cells.set(cellKey(r, c), cell);
        const w = parseFloat(td.style.width || td.getAttribute('width') || '');
        if (w > 0 && cs === 1 && !/%/.test(td.style.width || td.getAttribute('width') || '') && !sheet.colW.has(c)) sheet.colW.set(c, Math.round(/pt/.test(td.style.width) ? (w * 4) / 3 : w));
        c += cs;
      }
    });
    // <col width>
    Array.from(table.querySelectorAll('col')).forEach((col, i) => {
      const w = parseFloat((col as HTMLElement).style.width || col.getAttribute('width') || '');
      if (w > 0) sheet.colW.set(i, Math.round(w));
    });
  });
  if (!book.sheets.length) {
    // No table: put text lines in column A.
    const sheet = new Sheet(book, 'Sheet1');
    book.sheets.push(sheet);
    const body = doc.body ? cellText(doc.body) : '';
    body.split('\n').filter((l) => l.trim()).forEach((l, i) => sheet.cells.set(cellKey(i, 0), { v: l.trim() }));
    if (body.trim()) warnings.push('The page has no tables; its text was placed in column A.');
  }
  book.clearHistory();
  return book;
}

// --------------------------------------------------------------------------------------------
// Export

const escHtml = (s: string) => s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function borderCss(b: Border): string {
  const w = b.style === 'thick' ? 3 : b.style === 'medium' || b.style === 'mediumDashed' ? 2 : 1;
  const st = b.style === 'double' ? 'double' : /dash/i.test(b.style) ? 'dashed' : b.style === 'dotted' || b.style === 'hair' ? 'dotted' : 'solid';
  return `${b.style === 'double' ? 3 : w}px ${st} ${b.color || '#000000'}`;
}

export function styleCss(st: CellStyle, isNumber = false): string {
  const c: string[] = [];
  if (st.font) c.push(`font-family:'${st.font.replace(/'/g, '')}'`);
  if (st.size) c.push(`font-size:${st.size}pt`);
  if (st.bold) c.push('font-weight:bold');
  if (st.italic) c.push('font-style:italic');
  const deco = [st.underline ? 'underline' : '', st.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) c.push(`text-decoration:${deco}`);
  if (st.color) c.push(`color:${st.color}`);
  if (st.fill) c.push(`background-color:${st.fill}`);
  const ha = st.hAlign && st.hAlign !== 'general' ? (st.hAlign === 'centerContinuous' ? 'center' : st.hAlign === 'distributed' || st.hAlign === 'fill' ? 'justify' : st.hAlign) : isNumber ? 'right' : '';
  if (ha) c.push(`text-align:${ha}`);
  if (st.vAlign) c.push(`vertical-align:${st.vAlign === 'distributed' || st.vAlign === 'justify' ? 'middle' : st.vAlign}`);
  if (st.wrap) c.push('white-space:pre-wrap'); else c.push('white-space:pre');
  if (st.indent) c.push(`padding-left:${st.indent * 9}px`);
  if (st.bt) c.push(`border-top:${borderCss(st.bt)}`);
  if (st.bb) c.push(`border-bottom:${borderCss(st.bb)}`);
  if (st.bl) c.push(`border-left:${borderCss(st.bl)}`);
  if (st.br) c.push(`border-right:${borderCss(st.br)}`);
  return c.join(';');
}

function displayText(sheet: Sheet, r: number, c: number): { text: string; color?: string; isNumber?: boolean } {
  const cell = sheet.get(r, c);
  const v = cell?.v ?? null;
  if (v == null) return { text: '' };
  if (v instanceof CellError) return { text: v.code };
  const st = sheet.styleOf(r, c);
  const f = formatValue(v, st.numFmt, { date1904: sheet.book.calc.date1904 });
  return { text: f.text, color: f.color, isNumber: f.isNumber };
}

function sheetTable(sheet: Sheet, range?: Range | null): string {
  const rg = range || sheet.usedRange();
  if (!rg) return '<table></table>';
  const covered = new Set<number>();
  const spans = new Map<number, Range>();
  for (const m of sheet.merges) {
    if (m.r1 < rg.r1 || m.c1 < rg.c1 || m.r1 > rg.r2 || m.c1 > rg.c2) continue;
    spans.set(cellKey(m.r1, m.c1), m);
    for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(cellKey(r, c));
  }
  const out: string[] = ['<table style="border-collapse:collapse">', '<colgroup>'];
  for (let c = rg.c1; c <= rg.c2; c++) if (!sheet.hiddenCols.has(c)) out.push(`<col style="width:${sheet.colWidth(c)}px">`);
  out.push('</colgroup>');
  for (let r = rg.r1; r <= rg.r2; r++) {
    if (sheet.hiddenRows.has(r)) continue;
    const h = sheet.rowH.get(r);
    out.push(h ? `<tr style="height:${h}px">` : '<tr>');
    for (let c = rg.c1; c <= rg.c2; c++) {
      if (sheet.hiddenCols.has(c)) continue;
      const k = cellKey(r, c);
      if (covered.has(k)) continue;
      const m = spans.get(k);
      const span = m ? `${m.r2 > m.r1 ? ` rowspan="${Math.min(m.r2, rg.r2) - m.r1 + 1}"` : ''}${m.c2 > m.c1 ? ` colspan="${Math.min(m.c2, rg.c2) - m.c1 + 1}"` : ''}` : '';
      const d = displayText(sheet, r, c);
      const st = { ...sheet.styleOf(r, c) };
      if (d.color) st.color = d.color;
      const css = styleCss(st, d.isNumber);
      const cell = sheet.get(r, c);
      let body = escHtml(d.text).replace(/\n/g, '<br>');
      if (cell?.link && !/^\s*javascript:/i.test(cell.link)) body = `<a href="${escHtml(cell.link)}">${body}</a>`;
      out.push(`<td${span}${css ? ` style="${escHtml(css)}"` : ''}>${body}</td>`);
    }
    out.push('</tr>');
  }
  out.push('</table>');
  return out.join('\n');
}

export function exportHtml(book: Workbook, opts: { sheetIndex?: number; range?: Range } = {}): Blob {
  const sheets = opts.sheetIndex != null ? [book.sheets[opts.sheetIndex]].filter(Boolean) : book.sheets.filter((s) => s.visibility === 'visible');
  const title = book.props.title || (sheets.length === 1 ? sheets[0].name : 'Workbook');
  const body = sheets.map((s) => (sheets.length > 1 ? `<h2>${escHtml(s.name)}</h2>\n` : '') + sheetTable(s, opts.range)).join('\n');
  const html = `<!DOCTYPE html>\n<html>\n<head>\n<meta charset="utf-8">\n<meta name="generator" content="QuillLattice">\n<title>${escHtml(title)}</title>\n<style>body{font-family:Calibri,Carlito,Arial,sans-serif;font-size:11pt}table{border-collapse:collapse;margin-bottom:1.5em}td{padding:1px 4px;vertical-align:bottom;overflow:hidden}</style>\n</head>\n<body>\n${body}\n</body>\n</html>\n`;
  return new Blob([html], { type: 'text/html;charset=utf-8' });
}

// --------------------------------------------------------------------------------------------
// Markdown

export function sheetToMarkdown(sheet: Sheet, range?: Range | null): string {
  const rg = range || sheet.usedRange(false);
  if (!rg) return '';
  const rows: string[][] = [];
  const numeric: number[] = new Array(rg.c2 - rg.c1 + 1).fill(0);
  for (let r = rg.r1; r <= rg.r2; r++) {
    if (sheet.hiddenRows.has(r)) continue;
    const row: string[] = [];
    for (let c = rg.c1; c <= rg.c2; c++) {
      if (sheet.hiddenCols.has(c)) continue;
      const d = displayText(sheet, r, c);
      if (d.isNumber && r > rg.r1) numeric[c - rg.c1]++;
      row.push(d.text.replace(/\\/g, '\\\\').replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>'));
    }
    rows.push(row);
  }
  if (!rows.length) return '';
  const cols = rows[0].length;
  const visibleCols = [];
  for (let c = rg.c1; c <= rg.c2; c++) if (!sheet.hiddenCols.has(c)) visibleCols.push(c - rg.c1);
  const dataRows = rows.length - 1;
  const sep = visibleCols.map((ci) => (dataRows && numeric[ci] > dataRows / 2 ? '---:' : '---'));
  const line = (r: string[]) => '| ' + Array.from({ length: cols }, (_, i) => r[i] ?? '').join(' | ') + ' |';
  return [line(rows[0]), '| ' + sep.join(' | ') + ' |', ...rows.slice(1).map(line)].join('\n') + '\n';
}

export function exportMarkdown(book: Workbook, sheetIndex = 0, range?: Range): Blob {
  const sheet = book.sheets[Math.max(0, Math.min(sheetIndex, book.sheets.length - 1))];
  return new Blob([sheetToMarkdown(sheet, range)], { type: 'text/markdown;charset=utf-8' });
}

export { keyRow, keyCol };
