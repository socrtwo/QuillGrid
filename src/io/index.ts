/**
 * File format registry, sniffing and the import/export API, plus small helpers shared by the
 * converter modules (XML escaping and DOM access, text decoding, ZIP directory listing, base64).
 *
 * Converter modules are loaded lazily so the app only pays for the formats it actually opens.
 */
import type { Workbook } from '../model/workbook';
import type { Range } from '../model/address';

// --------------------------------------------------------------------------------------------
// Registry

export interface FormatDef {
  id: string;
  label: string;
  ext: string[];
  mime: string;
  canImport: boolean;
  canExport: boolean;
  binary: boolean;
  note?: string;
}

export interface CsvOptions {
  delimiter?: string;
  quote?: string;
  encoding?: string;
  headerRows?: number;
  decimal?: '.' | ',';
  thousands?: string;
  dateOrder?: 'mdy' | 'dmy' | 'ymd';
  detectTypes?: boolean;
  columnTypes?: ('general' | 'text' | 'date' | 'skip')[];
  bom?: boolean;
  lineEnding?: '\n' | '\r\n';
}

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

export const FORMATS: FormatDef[] = [
  { id: 'xlsx', label: 'Excel Workbook (.xlsx)', ext: ['xlsx'], mime: XLSX_MIME, canImport: true, canExport: true, binary: true },
  { id: 'xlsm', label: 'Excel Macro-Enabled Workbook (.xlsm)', ext: ['xlsm'], mime: 'application/vnd.ms-excel.sheet.macroEnabled.12', canImport: true, canExport: true, binary: true, note: 'VBA projects are kept as-is but not run' },
  { id: 'xltx', label: 'Excel Template (.xltx)', ext: ['xltx', 'xltm'], mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template', canImport: true, canExport: true, binary: true },
  { id: 'ods', label: 'OpenDocument Spreadsheet (.ods)', ext: ['ods', 'ots'], mime: 'application/vnd.oasis.opendocument.spreadsheet', canImport: true, canExport: true, binary: true },
  { id: 'fods', label: 'Flat OpenDocument Spreadsheet (.fods)', ext: ['fods'], mime: 'application/vnd.oasis.opendocument.spreadsheet-flat-xml', canImport: true, canExport: true, binary: false },
  { id: 'csv', label: 'CSV (Comma delimited) (.csv)', ext: ['csv'], mime: 'text/csv', canImport: true, canExport: true, binary: false },
  { id: 'tsv', label: 'Tab-separated values (.tsv)', ext: ['tsv', 'tab'], mime: 'text/tab-separated-values', canImport: true, canExport: true, binary: false },
  { id: 'txt', label: 'Text (Tab delimited) (.txt)', ext: ['txt', 'text', 'prn'], mime: 'text/plain', canImport: true, canExport: true, binary: false },
  { id: 'html', label: 'Web Page (.html)', ext: ['html', 'htm', 'xhtml'], mime: 'text/html', canImport: true, canExport: true, binary: false },
  { id: 'json', label: 'QuillLattice JSON (.json)', ext: ['json'], mime: 'application/json', canImport: true, canExport: true, binary: false },
  { id: 'qlw', label: 'QuillLattice Workbook (.qlw)', ext: ['qlw'], mime: 'application/x-quilllattice', canImport: true, canExport: true, binary: true },
  { id: 'qlwe', label: 'Encrypted QuillLattice Workbook (.qlwe)', ext: ['qlwe'], mime: 'application/x-quilllattice-encrypted', canImport: true, canExport: true, binary: true, note: 'AES-256 with a password' },
  { id: 'xls', label: 'Excel 97-2003 Workbook (.xls)', ext: ['xls', 'xlt', 'xlw'], mime: 'application/vnd.ms-excel', canImport: true, canExport: false, binary: true, note: 'Opens values, formulas and basic formatting; save as .xlsx to keep it' },
  { id: 'sylk', label: 'SYLK (.slk)', ext: ['slk', 'sylk'], mime: 'application/x-sylk', canImport: true, canExport: false, binary: false },
  { id: 'dif', label: 'Data Interchange Format (.dif)', ext: ['dif'], mime: 'application/x-dif', canImport: true, canExport: false, binary: false },
  { id: 'xml', label: 'XML Spreadsheet 2003 (.xml)', ext: ['xml'], mime: 'application/xml', canImport: true, canExport: false, binary: false },
  { id: 'md', label: 'Markdown table (.md)', ext: ['md', 'markdown'], mime: 'text/markdown', canImport: false, canExport: true, binary: false },
];

export function formatForName(name: string): FormatDef | undefined {
  const lower = name.toLowerCase();
  const ext = lower.includes('.') ? lower.slice(lower.lastIndexOf('.') + 1) : lower;
  return FORMATS.find((f) => f.ext.includes(ext));
}

export function formatById(id: string): FormatDef | undefined {
  return FORMATS.find((f) => f.id === id);
}

// --------------------------------------------------------------------------------------------
// Sniffing

const QLWE_MAGIC = 'QFDE1';

export function sniffFormat(data: ArrayBuffer, name: string): FormatDef | undefined {
  const u8 = new Uint8Array(data);
  const byName = formatForName(name);
  if (u8.length >= 4 && u8[0] === 0x50 && u8[1] === 0x4b && (u8[2] === 3 || u8[2] === 5 || u8[2] === 7)) {
    const head = latin1(u8.subarray(0, 120));
    if (head.includes('mimetypeapplication/vnd.oasis.opendocument.spreadsheet')) return formatById('ods');
    const names = zipNames(u8).map((n) => n.toLowerCase());
    if (names.includes('workbook.json')) return formatById('qlw');
    if (names.some((n) => n.startsWith('xl/') || n === '[content_types].xml')) {
      if (byName && ['xlsx', 'xlsm', 'xltx'].includes(byName.id)) return byName;
      if (names.includes('xl/vbaproject.bin')) return formatById('xlsm');
      return formatById('xlsx');
    }
    if (names.includes('content.xml')) return formatById('ods');
    // Damaged archive: trust the extension, else assume xlsx so recovery can try.
    return byName && byName.binary ? byName : formatById('xlsx');
  }
  if (u8.length >= 8 && u8[0] === 0xd0 && u8[1] === 0xcf && u8[2] === 0x11 && u8[3] === 0xe0) return formatById('xls');
  // BIFF2-5 bare streams
  if (u8.length >= 4 && (u8[0] === 0x09 || u8[0] === 0x08) && u8[1] === 0x08 && byName?.id === 'xls') return byName;
  if (latin1(u8.subarray(0, 5)) === QLWE_MAGIC) return formatById('qlwe');
  const head = decodeText(data.slice(0, Math.min(data.byteLength, 4096)));
  const t = head.replace(/^﻿/, '').trimStart();
  if (t.startsWith('ID;')) return formatById('sylk');
  if (/^TABLE\r?\n\s*0\s*,\s*1\s*\r?\n/.test(t)) return formatById('dif');
  if (t.startsWith('{') && /"format"\s*:\s*"quilllattice"|"qlw"\s*:/.test(head)) return formatById('json');
  if (t.startsWith('<')) {
    if (/urn:schemas-microsoft-com:office:spreadsheet/.test(head)) {
      // An HTML page saved by Excel also carries this namespace.
      if (/<html/i.test(head) && !/<\?mso-application\s+progid="Excel\.Sheet"/i.test(head)) return formatById('html');
      return formatById('xml');
    }
    if (/office:document/.test(head)) return formatById('fods');
    if (/<(!doctype html|html|table|body|head|meta)\b/i.test(t)) return formatById('html');
    if (byName && byName.canImport) return byName;
    return formatById('html');
  }
  if (byName && byName.canImport) return byName;
  if (t.startsWith('{') || t.startsWith('[')) return formatById('json');
  // Text heuristics
  const lines = t.split(/\r?\n/).slice(0, 20).filter((l) => l.length);
  const tabs = lines.filter((l) => l.includes('\t')).length;
  if (tabs && tabs >= lines.length / 2) return formatById('tsv');
  return formatById('csv');
}

// --------------------------------------------------------------------------------------------
// Import / export

export interface ImportResult { book: Workbook; format: string; warnings: string[]; recovery?: any; encrypted?: boolean }

/** Error thrown when a password is needed (or was wrong). `needsPassword` lets the UI prompt. */
export class PasswordError extends Error {
  needsPassword = true;
  constructor(msg: string, public wrong = false) { super(msg); }
}

const ZIP_FORMATS = ['xlsx', 'xlsm', 'xltx', 'ods', 'qlw'];

export async function importFile(data: ArrayBuffer, name: string, opts: { repair?: boolean; password?: string; csv?: CsvOptions } = {}): Promise<ImportResult> {
  const fmt = sniffFormat(data, name);
  const warnings: string[] = [];
  const recover = async (cause: unknown): Promise<ImportResult> => {
    let report: any = null;
    try {
      const mod: any = await import('./recover');
      report = await mod.recoverWorkbook(data, name, { forceRepair: !!opts.repair });
    } catch (e) {
      if (cause) throw new Error(friendlyOpenError(errMsg(cause)));
      throw e;
    }
    if (!report || !report.book) {
      if (cause) throw new Error(friendlyOpenError(errMsg(cause)));
      throw new Error('No recoverable content was found in this file.');
    }
    const w: string[] = [];
    if (cause) w.push(`Normal open failed: ${friendlyOpenError(errMsg(cause))}`);
    if (Array.isArray(report.warnings)) w.push(...report.warnings);
    if (Array.isArray(report.notes)) w.push(...report.notes);
    report.book.clearHistory?.();
    return { book: report.book, format: report.format || fmt?.id || 'xlsx', warnings: w, recovery: report };
  };

  if (opts.repair) return recover(null);
  if (!fmt) throw new Error(`Unsupported file type: ${name}`);
  try {
    const book = await importAs(fmt.id, data, name, opts, warnings);
    book.clearHistory();
    return { book, format: fmt.id, warnings, encrypted: fmt.id === 'qlwe' || undefined };
  } catch (e) {
    if (e instanceof PasswordError) throw e;
    if (fmt.id === 'xls' || ZIP_FORMATS.includes(fmt.id)) {
      if (fmt.id === 'xls' && isEncryptedOoxml(new Uint8Array(data))) throw new PasswordError('This workbook is encrypted with a password by Excel. Remove the password in Excel (File ▸ Info ▸ Protect Workbook) and open it again.');
      return recover(e);
    }
    throw new Error(friendlyOpenError(errMsg(e)));
  }
}

async function importAs(id: string, data: ArrayBuffer, name: string, opts: { password?: string; csv?: CsvOptions }, warnings: string[]): Promise<Workbook> {
  switch (id) {
    case 'xlsx': case 'xlsm': case 'xltx': return (await import('./xlsx-import')).importXlsx(data, name, warnings);
    case 'ods': return (await import('./ods')).importOds(data, warnings);
    case 'fods': return (await import('./ods')).importFods(decodeText(data), warnings);
    case 'csv': case 'tsv': case 'txt': {
      const csv = await import('./csv');
      const o: CsvOptions = { ...(opts.csv || {}) };
      if (!o.delimiter && id === 'tsv') o.delimiter = '\t';
      return csv.importCsv(data, name, o);
    }
    case 'html': return (await import('./html')).importHtml(decodeText(data), warnings);
    case 'json': return (await import('./native')).importJson(decodeText(data));
    case 'qlw': return (await import('./native')).importQlw(data);
    case 'qlwe': {
      if (!opts.password) throw new PasswordError('This workbook is protected with a password.');
      return (await import('./native')).importQlwe(data, opts.password);
    }
    case 'xls': {
      const mod: any = await import('./recover/xls');
      const book = await mod.importXls(data, name);
      if (!book) throw new Error('The .xls file could not be read.');
      return book;
    }
    case 'sylk': return (await import('./other')).importSylk(decodeText(data));
    case 'dif': return (await import('./other')).importDif(decodeText(data));
    case 'xml': return (await import('./other')).importXml2003(decodeText(data), warnings);
  }
  throw new Error(`Cannot open ${formatById(id)?.label || id} files.`);
}

export async function exportFile(book: Workbook, formatId: string, opts: { sheetIndex?: number; password?: string; csv?: CsvOptions; selectionOnly?: Range } = {}): Promise<Blob> {
  const fmt = formatById(formatId);
  if (!fmt || !fmt.canExport) throw new Error(`Cannot save as ${formatId}`);
  const sheetIndex = opts.sheetIndex ?? book.activeSheet;
  switch (fmt.id) {
    case 'xlsx': case 'xlsm': case 'xltx': return (await import('./xlsx-export')).exportXlsx(book, { kind: fmt.id as 'xlsx' });
    case 'ods': return (await import('./ods')).exportOds(book);
    case 'fods': return (await import('./ods')).exportFods(book);
    case 'csv': case 'tsv': case 'txt': {
      const o: CsvOptions = { ...(opts.csv || {}) };
      if (!o.delimiter) o.delimiter = fmt.id === 'csv' ? ',' : '\t';
      return (await import('./csv')).exportCsv(book, sheetIndex, o, opts.selectionOnly);
    }
    case 'html': return (await import('./html')).exportHtml(book, { sheetIndex: opts.sheetIndex, range: opts.selectionOnly });
    case 'md': return (await import('./html')).exportMarkdown(book, sheetIndex, opts.selectionOnly);
    case 'json': return (await import('./native')).exportJson(book);
    case 'qlw': return (await import('./native')).exportQlw(book);
    case 'qlwe': {
      if (!opts.password) throw new PasswordError('A password is required to save an encrypted workbook.');
      return (await import('./native')).exportQlwe(book, opts.password);
    }
  }
  throw new Error(`Cannot save as ${formatId}`);
}

export { parseCsv, detectDelimiter } from './csv';

/** Translates low-level ZIP/XML/inflate errors into plain language. */
export function friendlyOpenError(msg: string): string {
  if (/size mismatch|invalid (distance|code|block|stored)|too far back|incorrect header|unexpected end|crc32|corrupt/i.test(msg)) return 'compressed data inside the file is damaged.';
  if (/end of central directory|is this a zip|zip file/i.test(msg)) return 'the file\'s table of contents (ZIP central directory) is missing or damaged.';
  if (/workbook\.xml|workbook part/i.test(msg)) return 'the main workbook part could not be read.';
  if (/xml parse error|parsererror|not well-formed/i.test(msg)) return 'part of the file contains malformed XML.';
  if (/password|decrypt|operation-specific reason|OperationError/i.test(msg)) return 'the password is incorrect or the file is damaged.';
  return msg;
}

function errMsg(e: unknown): string { return e instanceof Error ? e.message : String(e); }

function isEncryptedOoxml(u8: Uint8Array): boolean {
  const pat = 'EncryptedPackage';
  const lim = Math.min(u8.length, 1 << 20);
  outer: for (let i = 0; i + pat.length * 2 < lim; i++) {
    for (let j = 0; j < pat.length; j++) if (u8[i + j * 2] !== pat.charCodeAt(j) || u8[i + j * 2 + 1] !== 0) continue outer;
    return true;
  }
  return false;
}

// --------------------------------------------------------------------------------------------
// Shared helpers: text

function latin1(u8: Uint8Array): string {
  let s = '';
  for (let i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]);
  return s;
}

/** Decodes bytes as text: BOM detection (UTF-8/16), then UTF-8, falling back to Windows-1252. */
export function decodeText(data: ArrayBuffer | Uint8Array, encoding?: string): string {
  const u8 = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (u8[0] === 0xff && u8[1] === 0xfe) return new TextDecoder('utf-16le').decode(u8.subarray(2));
  if (u8[0] === 0xfe && u8[1] === 0xff) return new TextDecoder('utf-16be').decode(u8.subarray(2));
  if (u8[0] === 0xef && u8[1] === 0xbb && u8[2] === 0xbf) return new TextDecoder('utf-8').decode(u8.subarray(3));
  if (encoding && !/^utf-?8$/i.test(encoding)) {
    try { return new TextDecoder(encoding).decode(u8); } catch { /* unsupported label: fall through */ }
  }
  // BOM-less UTF-16 (common for "Unicode text" exports): many zero bytes at odd/even positions.
  if (!encoding && u8.length >= 4) {
    let zo = 0, ze = 0;
    const n = Math.min(u8.length, 2000);
    for (let i = 0; i < n; i++) if (u8[i] === 0) { if (i & 1) zo++; else ze++; }
    if (zo > n / 4 && ze < n / 50) return new TextDecoder('utf-16le').decode(u8);
    if (ze > n / 4 && zo < n / 50) return new TextDecoder('utf-16be').decode(u8);
  }
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(u8);
  } catch {
    return new TextDecoder('windows-1252').decode(u8);
  }
}

// --------------------------------------------------------------------------------------------
// Shared helpers: XML

// eslint-disable-next-line no-control-regex
const INVALID_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

export function stripInvalidXml(s: string): string { return s.replace(INVALID_XML, ''); }

/** Escapes text for XML content or attribute values and strips characters invalid in XML 1.0. */
export function esc(s: unknown): string {
  return stripInvalidXml(String(s ?? ''))
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

export const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const ENT: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
/** Decodes XML entities and character references. */
export function unescXml(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_m, e: string) => {
    if (e[0] === '#') {
      const cp = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      try { return String.fromCodePoint(cp); } catch { return ''; }
    }
    return ENT[e];
  });
}

/** Parses an XML string; throws on malformed XML. */
export function parseXml(text: string, what = 'XML'): Document {
  const doc = new DOMParser().parseFromString(text.replace(/^﻿/, ''), 'application/xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  if (err && (doc.documentElement.localName === 'parsererror' || err.parentNode === doc.documentElement)) {
    throw new Error(`XML parse error in ${what}: ${(err.textContent || '').slice(0, 200)}`);
  }
  return doc;
}

/** Direct element children with the given local name (namespace/prefix agnostic). */
export function kids(el: Element | Document | null | undefined, local?: string): Element[] {
  if (!el) return [];
  const out: Element[] = [];
  const parent = (el as Document).documentElement && (el as Document).nodeType === 9 ? [(el as Document).documentElement] : null;
  const list = parent || Array.from((el as Element).children || []);
  for (const c of list) if (!local || c.localName === local) out.push(c);
  return out;
}

export function kid(el: Element | Document | null | undefined, local: string): Element | null {
  if (!el) return null;
  if ((el as Document).nodeType === 9) { const d = (el as Document).documentElement; return d && d.localName === local ? d : null; }
  const ch = (el as Element).children;
  for (let i = 0; i < ch.length; i++) if (ch[i].localName === local) return ch[i];
  return null;
}

/** Child at a path of local names: path(el, 'a', 'b') = first b inside first a. */
export function path(el: Element | null | undefined, ...names: string[]): Element | null {
  let cur: Element | null = el || null;
  for (const n of names) { if (!cur) return null; cur = kid(cur, n); }
  return cur;
}

/** All descendants with the given local name. */
export function desc(el: Element | Document | null | undefined, local: string): Element[] {
  if (!el) return [];
  return Array.from(el.getElementsByTagNameNS('*', local));
}

/** Attribute by name; a bare local name also matches prefixed attributes (r:id matches 'id'). */
export function attr(el: Element | null | undefined, name: string): string | null {
  if (!el) return null;
  const v = el.getAttribute(name);
  if (v != null) return v;
  const at = el.attributes;
  for (let i = 0; i < at.length; i++) if (at[i].localName === name) return at[i].value;
  return null;
}

/** r:id style relationship attribute. */
export function relId(el: Element | null | undefined, name = 'id'): string | null {
  if (!el) return null;
  const at = el.attributes;
  for (let i = 0; i < at.length; i++) {
    const a = at[i];
    if (a.localName === name && (a.prefix || a.namespaceURI)) return a.value;
  }
  return el.getAttribute('r:' + name);
}

export function numAttr(el: Element | null | undefined, name: string): number | null {
  const v = attr(el, name);
  if (v == null || v === '') return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

export function boolAttr(el: Element | null | undefined, name: string, dflt = false): boolean {
  const v = attr(el, name);
  if (v == null) return dflt;
  return v === '1' || v.toLowerCase() === 'true' || v.toLowerCase() === 'on';
}

// --------------------------------------------------------------------------------------------
// Shared helpers: binary

/** Lists entry names in a ZIP archive from the central directory (local headers as a fallback). */
export function zipNames(u8: Uint8Array): string[] {
  const names: string[] = [];
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  const dec = new TextDecoder('utf-8');
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
    if (u8[i] === 0x50 && u8[i + 1] === 0x4b && u8[i + 2] === 5 && u8[i + 3] === 6) { eocd = i; break; }
  }
  if (eocd >= 0) {
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    for (let k = 0; k < count && p + 46 <= u8.length; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
      names.push(dec.decode(u8.subarray(p + 46, p + 46 + nl)));
      p += 46 + nl + el + cl;
    }
    if (names.length) return names;
  }
  // Scan local file headers.
  for (let p = 0; p + 30 <= u8.length;) {
    if (dv.getUint32(p, true) !== 0x04034b50) { p++; continue; }
    const nl = dv.getUint16(p + 26, true), el = dv.getUint16(p + 28, true), cs = dv.getUint32(p + 18, true);
    names.push(dec.decode(u8.subarray(p + 30, p + 30 + nl)));
    p += 30 + nl + el + (cs > 0 && cs < u8.length ? cs : 0);
    if (names.length > 5000) break;
  }
  return names;
}

export function bytesToBase64(bytes: Uint8Array): string {
  let s = '';
  const CH = 0x8000;
  for (let i = 0; i < bytes.length; i += CH) s += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + CH)));
  return btoa(s);
}

export function base64ToBytes(b64: string): Uint8Array {
  const s = atob(b64.replace(/\s+/g, ''));
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}

const MIME_EXT: Record<string, string> = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp',
  tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf', ico: 'image/x-icon',
};
export function mimeForExt(ext: string): string { return MIME_EXT[ext.toLowerCase()] || 'application/octet-stream'; }
export function extForMime(mime: string): string {
  const m = mime.toLowerCase();
  if (m === 'image/jpeg') return 'jpeg';
  for (const [e, t] of Object.entries(MIME_EXT)) if (t === m) return e;
  return 'png';
}

export function parseDataUrl(src: string): { mime: string; bytes: Uint8Array } | null {
  const m = /^data:([^;,]+)?((?:;[^;,]*)*?)(;base64)?,(.*)$/s.exec(src);
  if (!m) return null;
  const mime = m[1] || 'application/octet-stream';
  if (m[3]) return { mime, bytes: base64ToBytes(m[4]) };
  return { mime, bytes: new TextEncoder().encode(decodeURIComponent(m[4])) };
}

/** Size in px of a PNG/JPEG/GIF image, when it can be read from the header. */
export function imageSize(b: Uint8Array): { w: number; h: number } | null {
  if (b[0] === 0x89 && b[1] === 0x50 && b.length > 24) return { w: (b[16] << 24 | b[17] << 16 | b[18] << 8 | b[19]) >>> 0, h: (b[20] << 24 | b[21] << 16 | b[22] << 8 | b[23]) >>> 0 };
  if (b[0] === 0x47 && b[1] === 0x49 && b.length > 10) return { w: b[6] | (b[7] << 8), h: b[8] | (b[9] << 8) };
  if (b[0] === 0xff && b[1] === 0xd8) {
    let p = 2;
    while (p + 9 < b.length) {
      if (b[p] !== 0xff) { p++; continue; }
      const mk = b[p + 1];
      const len = (b[p + 2] << 8) | b[p + 3];
      if (mk >= 0xc0 && mk <= 0xcf && mk !== 0xc4 && mk !== 0xc8 && mk !== 0xcc) return { h: (b[p + 5] << 8) | b[p + 6], w: (b[p + 7] << 8) | b[p + 8] };
      p += 2 + len;
    }
  }
  return null;
}

// --------------------------------------------------------------------------------------------
// Shared helpers: formulas

/** Removes Excel's future-function prefixes (_xlfn., _xlws., _xlpm.) outside string literals. */
export function stripXlfn(f: string): string {
  if (f.indexOf('_xl') < 0) return f;
  return mapOutsideStrings(f, (s) => s.replace(/_xlfn\._xlws\.|_xlfn\.|_xlws\.|_xlpm\./gi, ''));
}

/** Applies fn to the parts of a formula that are not inside "string literals" or 'quoted names'. */
export function mapOutsideStrings(f: string, fn: (s: string) => string): string {
  let out = '', i = 0, start = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"' || ch === "'") {
      out += fn(f.slice(start, i));
      let j = i + 1;
      while (j < f.length) { if (f[j] === ch) { if (f[j + 1] === ch) { j += 2; continue; } break; } j++; }
      out += f.slice(i, j + 1);
      i = start = j + 1;
      continue;
    }
    i++;
  }
  return out + fn(f.slice(start));
}

// --------------------------------------------------------------------------------------------
// Shared helpers: colors

/** Normalizes a CSS-ish color to '#rrggbb' (or null). */
export function normColor(c: string | null | undefined): string | null {
  if (!c) return null;
  const s = c.trim().toLowerCase();
  let m = /^#?([0-9a-f]{6})$/.exec(s);
  if (m) return '#' + m[1];
  m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(s);
  if (m) return '#' + m[1] + m[1] + m[2] + m[2] + m[3] + m[3];
  m = /^#?([0-9a-f]{2})([0-9a-f]{6})$/.exec(s); // AARRGGBB
  if (m) return '#' + m[2];
  m = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)/.exec(s);
  if (m) return '#' + [m[1], m[2], m[3]].map((x) => Math.min(255, +x).toString(16).padStart(2, '0')).join('');
  const named = CSS_NAMED[s];
  return named ? '#' + named : null;
}

const CSS_NAMED: Record<string, string> = {
  black: '000000', white: 'ffffff', red: 'ff0000', green: '008000', lime: '00ff00', blue: '0000ff', yellow: 'ffff00', cyan: '00ffff', aqua: '00ffff',
  magenta: 'ff00ff', fuchsia: 'ff00ff', gray: '808080', grey: '808080', silver: 'c0c0c0', maroon: '800000', olive: '808000', purple: '800080',
  teal: '008080', navy: '000080', orange: 'ffa500', pink: 'ffc0cb', brown: 'a52a2a', gold: 'ffd700', lightgray: 'd3d3d3', lightgrey: 'd3d3d3',
  darkgray: 'a9a9a9', darkgrey: 'a9a9a9', darkred: '8b0000', darkblue: '00008b', darkgreen: '006400', lightblue: 'add8e6', lightgreen: '90ee90',
  lightyellow: 'ffffe0', whitesmoke: 'f5f5f5', gainsboro: 'dcdcdc',
};
