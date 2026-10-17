/**
 * E-mail in a spreadsheet.
 *
 *   EML / EMLX / MHT   one "Message" sheet: header fields (From, To, Cc, Date, Subject,
 *                      Attachments) in columns A-B, then the body, one line per row
 *   MBOX               one "Messages" sheet: a row per message (Date, From, To, Subject,
 *                      Attachments, Body)
 *   MSG (Outlook)      as EML, read from the compound file's MAPI properties
 *
 * In every case, delimited-text attachments (CSV, TSV, tab-separated TXT) become sheets of their
 * own, and spreadsheet attachments (XLSX, XLS, ODS, ...) are opened and their sheets added. Tables
 * in an HTML body are added as sheets too.
 */
import { Workbook, Sheet } from '../model/workbook';
import type { Cell, CellStyle } from '../model/types';
import { MAX_ROWS } from '../model/address';
import { parseCsv, detectDelimiter, fillSheet } from './csv';
import { readCfb } from './recover/doc97';
import { cp1252 } from './recover/salvage';

// --------------------------------------------------------------------------------------------
// MIME

export interface MimePart { headers: Record<string, string>; body: Uint8Array; parts: MimePart[] }
export interface Attachment { name: string; mime: string; data: Uint8Array }
export interface Message {
  from: string; to: string; cc: string; date: Date | null; subject: string;
  text: string; html: string; attachments: Attachment[];
}

const latin1 = (u8: Uint8Array) => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, Array.from(u8.subarray(i, i + 0x8000))); return s; };
const bytesOf = (s: string) => { const u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i) & 0xff; return u; };

function decodeCharset(u8: Uint8Array, charset: string): string {
  const cs = (charset || 'utf-8').toLowerCase().replace(/^"|"$/g, '');
  if (cs === 'us-ascii' || cs === 'iso-8859-1' || cs === 'latin1' || cs === 'windows-1252' || cs === 'cp1252') { let s = ''; for (const b of u8) s += cp1252(b); return s; }
  try { return new TextDecoder(cs, { fatal: false }).decode(u8); } catch { return new TextDecoder('utf-8').decode(u8); }
}

/** RFC 2047 encoded words: =?charset?B?...?= and =?charset?Q?...?=. */
export function decodeWords(s: string): string {
  return s.replace(/\?=\s+=\?/g, '?==?').replace(/=\?([^?]+)\?([BbQq])\?([^?]*)\?=/g, (_m, cs: string, enc: string, txt: string) => {
    let bytes: Uint8Array;
    if (enc.toUpperCase() === 'B') { try { bytes = bytesOf(atob(txt.replace(/\s/g, ''))); } catch { return txt; } }
    else bytes = bytesOf(txt.replace(/_/g, ' ').replace(/=([0-9A-Fa-f]{2})/g, (_x, h: string) => String.fromCharCode(parseInt(h, 16))));
    return decodeCharset(bytes, cs.replace(/\*.*$/, ''));
  });
}

function parseHeaders(text: string): Record<string, string> {
  const out: Record<string, string> = {};
  const unfolded = text.replace(/\r?\n[ \t]+/g, ' ');
  for (const line of unfolded.split(/\r?\n/)) {
    const m = /^([^:\s]+):\s*(.*)$/.exec(line);
    if (!m) continue;
    const k = m[1].toLowerCase();
    out[k] = out[k] != null && (k === 'received') ? out[k] : (out[k] != null ? `${out[k]}, ${m[2]}` : m[2]);
  }
  return out;
}

/** A header parameter (boundary, charset, name, filename), with RFC 2231 continuations and charsets. */
function param(h: string | undefined, name: string): string {
  if (!h) return '';
  const parts: [number, string, boolean][] = [];
  const re = new RegExp(`(?:^|;)\\s*${name}(\\*(\\d+))?(\\*)?\\s*=\\s*("((?:[^"\\\\]|\\\\.)*)"|[^;]*)`, 'gi');
  let m: RegExpExecArray | null;
  while ((m = re.exec(h))) parts.push([m[2] ? +m[2] : 0, m[5] != null ? m[5].replace(/\\(.)/g, '$1') : m[4].trim(), !!m[3]]);
  if (!parts.length) return '';
  parts.sort((a, b) => a[0] - b[0]);
  let charset = '';
  const raw = parts.map(([, v, ext], i) => {
    if (!ext) return v;
    if (i === 0) { const mm = /^([^']*)'[^']*'(.*)$/.exec(v); if (mm) { charset = mm[1]; v = mm[2]; } }
    return v.replace(/%([0-9A-Fa-f]{2})/g, (_x, hh: string) => String.fromCharCode(parseInt(hh, 16)));
  }).join('');
  return decodeWords(charset ? decodeCharset(bytesOf(raw), charset) : raw);
}

function decodeBody(body: Uint8Array, cte: string): Uint8Array {
  const enc = (cte || '').toLowerCase().trim();
  if (enc === 'base64') {
    const t = latin1(body).replace(/[^A-Za-z0-9+/=]/g, '');
    try { return bytesOf(atob(t.slice(0, t.length - (t.length % 4)))); } catch { return body; }
  }
  if (enc === 'quoted-printable') return bytesOf(latin1(body).replace(/=\r?\n/g, '').replace(/=([0-9A-Fa-f]{2})/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16))));
  return body;
}

export function parseMime(raw: Uint8Array, depth = 0): MimePart {
  const s = latin1(raw);
  const m = /\r?\n\r?\n/.exec(s);
  const headEnd = m ? m.index : s.length;
  const headers = parseHeaders(s.slice(0, headEnd));
  const bodyStart = m ? m.index + m[0].length : s.length;
  const part: MimePart = { headers, body: raw.subarray(bodyStart), parts: [] };
  const ct = headers['content-type'] || '';
  if (/^multipart\//i.test(ct) && depth < 12) {
    const boundary = param(ct, 'boundary');
    if (boundary) {
      const body = s.slice(bodyStart);
      const delim = `--${boundary}`;
      const pieces = body.split(new RegExp(`(?:^|\\r?\\n)${delim.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}`));
      for (const p of pieces.slice(1)) {
        if (p.startsWith('--')) break;
        const content = p.replace(/^[ \t]*\r?\n/, '');
        part.parts.push(parseMime(bytesOf(content), depth + 1));
      }
    }
  }
  return part;
}

function messageFromMime(root: MimePart): Message {
  const h = root.headers;
  const msg: Message = {
    from: decodeWords(h.from || ''), to: decodeWords(h.to || ''), cc: decodeWords(h.cc || ''), subject: decodeWords(h.subject || ''),
    date: h.date && !Number.isNaN(Date.parse(h.date.replace(/\s*\(.*\)\s*$/, ''))) ? new Date(Date.parse(h.date.replace(/\s*\(.*\)\s*$/, ''))) : null,
    text: '', html: '', attachments: [],
  };
  const walk = (p: MimePart) => {
    const ct = (p.headers['content-type'] || 'text/plain').toLowerCase();
    if (p.parts.length) { p.parts.forEach(walk); return; }
    const disp = p.headers['content-disposition'] || '';
    const fname = param(disp, 'filename') || param(p.headers['content-type'], 'name');
    const data = decodeBody(p.body, p.headers['content-transfer-encoding']);
    const isAttach = /^attachment/i.test(disp) || (!!fname && !/^text\/(plain|html)/.test(ct));
    if (!isAttach && ct.startsWith('text/plain') && !msg.text) { msg.text = decodeCharset(data, param(p.headers['content-type'], 'charset')); return; }
    if (!isAttach && ct.startsWith('text/html') && !msg.html) { msg.html = decodeCharset(data, param(p.headers['content-type'], 'charset')); return; }
    if (ct.startsWith('message/rfc822')) { const inner = messageFromMime(parseMime(data)); msg.attachments.push(...inner.attachments); if (!fname) return; }
    msg.attachments.push({ name: fname || `attachment ${msg.attachments.length + 1}`, mime: ct.split(';')[0].trim(), data });
  };
  walk(root);
  return msg;
}

export function parseEml(raw: Uint8Array): Message {
  let u8 = raw;
  // Apple Mail .emlx: a byte count line, then the message.
  const first = latin1(raw.subarray(0, 16));
  const m = /^(\d+)\s*\r?\n/.exec(first);
  if (m) u8 = raw.subarray(m[0].length, m[0].length + Number(m[1]));
  return messageFromMime(parseMime(u8));
}

export function splitMbox(raw: Uint8Array): Uint8Array[] {
  const s = latin1(raw);
  const out: Uint8Array[] = [];
  const re = /(?:^|\n)From [^\n]*\n/g;
  const starts: number[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(s))) starts.push(m.index + (m[0].startsWith('\n') ? 1 : 0) + m[0].trimStart().length);
  for (let k = 0; k < starts.length; k++) {
    const end = k + 1 < starts.length ? s.lastIndexOf('From ', starts[k + 1] - 1) : s.length;
    out.push(bytesOf(s.slice(starts[k], end).replace(/^>(>*From )/gm, '$1')));
  }
  return out;
}

// --------------------------------------------------------------------------------------------
// Outlook .msg (MAPI properties in a compound file)

export function parseMsg(raw: Uint8Array): Message {
  const cfb = readCfb(raw);
  if (!cfb?.read) throw new Error('The Outlook message could not be read.');
  const paths = cfb.paths || [];
  const props = (prefix: string) => {
    const pfx = prefix ? `${prefix}/` : '';
    const str = (id: number): string => {
      const tag = id.toString(16).toUpperCase().padStart(4, '0');
      const u = cfb.read!(`${pfx}__substg1.0_${tag}001F`);
      if (u) { let t = ''; for (let i = 0; i + 1 < u.length; i += 2) t += String.fromCharCode(u[i] | (u[i + 1] << 8)); return t.replace(/\0+$/, ''); }
      const a = cfb.read!(`${pfx}__substg1.0_${tag}001E`);
      if (a) return decodeCharset(a, 'windows-1252').replace(/\0+$/, '');
      return '';
    };
    const bin = (id: number) => cfb.read!(`${pfx}__substg1.0_${id.toString(16).toUpperCase().padStart(4, '0')}0102`);
    return { str, bin };
  };
  const top = props('');
  const msg: Message = { from: '', to: '', cc: '', subject: '', date: null, text: '', html: '', attachments: [] };
  const smtp = (x: string) => (/^\/O=/i.test(x) ? '' : x); // Exchange (X.500) addresses mean nothing outside the organization
  const senderName = top.str(0x0c1a), senderMail = smtp(top.str(0x5d01)) || smtp(top.str(0x0c1f)) || smtp(top.str(0x0065));
  msg.from = senderName && senderMail && senderName !== senderMail ? `${senderName} <${senderMail}>` : senderName || senderMail;
  msg.to = top.str(0x0e04);
  msg.cc = top.str(0x0e03);
  msg.subject = top.str(0x0037);
  msg.text = top.str(0x1000);
  const html = top.bin(0x1013);
  msg.html = html ? decodeCharset(html, 'utf-8') : top.str(0x1013);
  // Dates: client submit time (0x0039) or delivery time (0x0E06), FILETIME in the fixed-size property stream.
  const pst = cfb.read('__properties_version1.0');
  if (pst) {
    for (let o = 32; o + 16 <= pst.length; o += 16) {
      const tag = (pst[o] | (pst[o + 1] << 8) | (pst[o + 2] << 16) | (pst[o + 3] << 24)) >>> 0;
      const id = tag >>> 16, type = tag & 0xffff;
      if (type === 0x0040 && (id === 0x0039 || id === 0x0e06) && !msg.date) {
        const lo = (pst[o + 8] | (pst[o + 9] << 8) | (pst[o + 10] << 16) | (pst[o + 11] << 24)) >>> 0;
        const hi = (pst[o + 12] | (pst[o + 13] << 8) | (pst[o + 14] << 16) | (pst[o + 15] << 24)) >>> 0;
        const ms = (hi * 4294967296 + lo) / 10000 - 11644473600000;
        if (ms > 0) msg.date = new Date(ms);
      }
    }
  }
  if (!msg.date) {
    const hdr = top.str(0x007d);
    const d = /^Date:\s*(.+)$/im.exec(hdr)?.[1];
    if (d && !Number.isNaN(Date.parse(d))) msg.date = new Date(Date.parse(d));
  }
  const storages = [...new Set(paths.map((p) => /^(__attach_version1\.0_#[0-9A-Fa-f]{8})\//.exec(p)?.[1]).filter((x): x is string => !!x))];
  for (const st of storages) {
    const a = props(st);
    const data = a.bin(0x3701);
    if (!data) continue;
    msg.attachments.push({ name: a.str(0x3707) || a.str(0x3704) || 'attachment', mime: a.str(0x370e) || 'application/octet-stream', data });
  }
  return msg;
}

// --------------------------------------------------------------------------------------------
// Workbook

function styler(book: Workbook): (st: CellStyle) => number {
  const cache = new Map<string, number>();
  return (st) => { const k = JSON.stringify(st); let id = cache.get(k); if (id == null) { id = book.styleId(st); cache.set(k, id); } return id; };
}

const serial = (d: Date) => (d.getTime() - d.getTimezoneOffset() * 60000) / 86400000 + 25569;

function htmlToText(html: string): string {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  doc.querySelectorAll('script,style').forEach((e) => e.remove());
  doc.querySelectorAll('br').forEach((e) => e.replaceWith('\n'));
  doc.querySelectorAll('p,div,tr,li,h1,h2,h3,h4,h5,h6').forEach((e) => e.append('\n'));
  doc.querySelectorAll('td,th').forEach((e) => e.append('\t'));
  return (doc.body?.textContent || '').replace(/ /g, ' ').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

function uniqueName(book: Workbook, base: string): string {
  let nm = base.replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Sheet';
  let k = 2;
  while (book.sheets.some((s) => s.name.toLowerCase() === nm.toLowerCase())) nm = `${base.slice(0, 26)} (${k++})`;
  return nm;
}

/** Copies a sheet of another workbook, translating its style numbers. */
export function adoptSheet(target: Workbook, src: Sheet, name: string): Sheet {
  const sh = new Sheet(target, uniqueName(target, name));
  const map = new Map<number, number>();
  const sid = (s: number | undefined) => {
    if (!s) return undefined;
    let id = map.get(s);
    if (id == null) { id = target.styleId(src.book.styles[s] || {}); map.set(s, id); }
    return id || undefined;
  };
  for (const [k, c] of src.cells) { const n: Cell = { ...c }; const s = sid(c.s); if (s) n.s = s; else delete n.s; sh.cells.set(k, n); }
  for (const [c, w] of src.colW) sh.colW.set(c, w);
  for (const [r, h] of src.rowH) sh.rowH.set(r, h);
  sh.merges = src.merges.map((m) => ({ ...m }));
  for (const [k, cm] of src.comments) sh.comments.set(k, { ...cm });
  sh.freeze = { ...src.freeze };
  sh.defaultColW = src.defaultColW;
  target.sheets.push(sh);
  return sh;
}

async function addAttachments(book: Workbook, atts: Attachment[], warnings: string[]): Promise<void> {
  for (const a of atts) {
    const ext = /\.([A-Za-z0-9]+)$/.exec(a.name)?.[1]?.toLowerCase() || '';
    const base = a.name.replace(/\.[^.]+$/, '');
    try {
      if (['csv', 'tsv', 'tab', 'txt'].includes(ext) || /text\/(csv|tab-separated-values)/.test(a.mime)) {
        const text = decodeCharset(a.data, 'utf-8');
        const delim = ext === 'tsv' || ext === 'tab' ? '\t' : ext === 'csv' ? detectDelimiter(text) : text.includes('\t') ? '\t' : '';
        if (!delim) continue;
        const sh = new Sheet(book, uniqueName(book, base));
        fillSheet(sh, parseCsv(text, { delimiter: delim }), { delimiter: delim });
        book.sheets.push(sh);
      } else if (['xlsx', 'xlsm', 'xls', 'ods', 'fods', 'xlsb', 'wk1', 'wks', 'wk3', 'wk4', '123', 'dbf', 'wb1', 'wb2', 'wb3', 'wq1', 'slk', 'dif'].includes(ext)) {
        const { importFile } = await import('./index');
        const r = await importFile(a.data.slice().buffer, a.name);
        for (const s of r.book.sheets) adoptSheet(book, s, r.book.sheets.length > 1 ? `${base} ${s.name}` : base);
      }
    } catch (e: any) {
      warnings.push(`The attachment "${a.name}" could not be opened (${e?.message || e}).`);
    }
  }
}

async function addHtmlTables(book: Workbook, html: string): Promise<void> {
  if (!/<table/i.test(html)) return;
  const { importHtml } = await import('./html');
  const tb = importHtml(html, [], { allTables: true });
  tb.sheets.forEach((s, i) => { if (s.cells.size > 1) adoptSheet(book, s, `Body table ${i + 1}`); });
}

export async function messageBook(msg: Message, title: string, warnings: string[]): Promise<Workbook> {
  const book = new Workbook(false);
  const sid = styler(book);
  const sh = new Sheet(book, 'Message');
  book.sheets.push(sh);
  sh.colW.set(0, 96); sh.colW.set(1, 560);
  const label = sid({ bold: true, vAlign: 'top' });
  const fields: [string, Cell['v'], CellStyle?][] = [
    ['From', msg.from], ['To', msg.to], ['Cc', msg.cc],
    ['Date', msg.date ? serial(msg.date) : '', msg.date ? { numFmt: 'yyyy-mm-dd hh:mm', hAlign: 'left' } : undefined],
    ['Subject', msg.subject], ['Attachments', msg.attachments.map((a) => a.name).join(', ')],
  ];
  let r = 0;
  for (const [k, v, st] of fields) {
    if (v === '' || v == null) continue;
    sh.setComputed(r, 0, { v: k, s: label });
    const cell: Cell = { v };
    if (st) cell.s = sid(st);
    sh.setComputed(r, 1, cell);
    r++;
  }
  r++;
  sh.setComputed(r++, 0, { v: 'Body', s: label });
  const body = msg.text || (msg.html ? htmlToText(msg.html) : '');
  for (const line of body.replace(/\r\n?/g, '\n').split('\n')) {
    if (r >= MAX_ROWS) break;
    if (line) sh.setComputed(r, 0, { v: line.length > 32767 ? line.slice(0, 32767) : line });
    r++;
  }
  if (msg.html) await addHtmlTables(book, msg.html);
  await addAttachments(book, msg.attachments, warnings);
  book.props.title = msg.subject || title;
  book.activeSheet = 0;
  return book;
}

export async function importEmail(id: string, data: ArrayBuffer, name: string, warnings: string[]): Promise<Workbook> {
  const u8 = new Uint8Array(data);
  const title = name.replace(/^.*[\\/]/, '').replace(/\.[^.]+$/, '');
  if (id === 'msg') return messageBook(parseMsg(u8), title, warnings);
  if (id === 'eml') return messageBook(parseEml(u8), title, warnings);
  // MBOX: a row per message.
  const msgs = splitMbox(u8).map((m) => { try { return parseEml(m); } catch { return null; } }).filter((m): m is Message => !!m);
  if (!msgs.length) throw new Error('No messages were found in this mailbox.');
  const book = new Workbook(false);
  const sid = styler(book);
  const sh = new Sheet(book, 'Messages');
  book.sheets.push(sh);
  const head = ['Date', 'From', 'To', 'Subject', 'Attachments', 'Body'];
  const hs = sid({ bold: true, bb: { style: 'thin' } });
  head.forEach((h, c) => sh.setComputed(0, c, { v: h, s: hs }));
  [130, 200, 200, 280, 160, 480].forEach((w, c) => sh.colW.set(c, w));
  const dateStyle = sid({ numFmt: 'yyyy-mm-dd hh:mm', hAlign: 'left', vAlign: 'top' });
  const top = sid({ vAlign: 'top' });
  msgs.forEach((m, i) => {
    const r = i + 1;
    if (r >= MAX_ROWS) return;
    if (m.date) sh.setComputed(r, 0, { v: serial(m.date), s: dateStyle });
    const vals = [m.from, m.to, m.subject, m.attachments.map((a) => a.name).join(', '), (m.text || (m.html ? htmlToText(m.html) : '')).slice(0, 32767)];
    vals.forEach((v, k) => { if (v) sh.setComputed(r, k + 1, { v, s: top }); });
  });
  sh.freeze = { rows: 1, cols: 0 };
  sh.filter = { range: { r1: 0, c1: 0, r2: msgs.length, c2: head.length - 1 }, columns: {} };
  for (const m of msgs) await addAttachments(book, m.attachments, warnings);
  book.props.title = title;
  book.activeSheet = 0;
  return book;
}
