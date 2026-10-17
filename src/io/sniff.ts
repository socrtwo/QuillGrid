/**
 * Content-based file identification.
 *
 * Every rule looks at the bytes (signatures, record structure, ZIP and compound-file directories,
 * text patterns) and returns a confidence between 0 and 1. The file extension is only a tie-breaker
 * between rules that score about the same, and the fallback when no rule recognizes the content.
 * When the content and the extension disagree, the content wins and the result carries a short
 * note for the person opening the file.
 *
 * Rules are data: other modules add their own with registerSniffer().
 */
import { lotusVariant, LOTUS_VARIANT_LABEL } from './legacy/lotus';
import { isOle2, readCfb } from './recover/doc97';

export type SniffCategory = 'spreadsheet' | 'document' | 'presentation' | 'email' | 'unknown';

export interface SniffResult {
  /** Importer format id (see FORMATS in index.ts); the file can be opened when that format can import. */
  format: string;
  /** What the file is, e.g. "Lotus 1-2-3 Release 2 worksheet (WK1)". */
  label: string;
  /** 0..1. */
  confidence: number;
  category: SniffCategory;
  /** Short message for the person opening the file (content and extension disagree, no extension, ...). */
  note?: string;
  /** Recognized from the extension only. */
  byExtension?: boolean;
  /** The rule that matched. */
  rule: string;
}

export interface SniffContext {
  bytes: Uint8Array;
  name: string;
  /** Lower-case extension without the dot ('' when none). */
  ext: string;
  /** Latin-1 view of the first 4 KB. */
  head: string;
  /** Text of the first 64 KB (BOM-aware UTF-8/UTF-16, else Windows-1252). */
  text(): string;
  /** ZIP entry names in lower case, or null when the file is not a ZIP archive. */
  zip(): string[] | null;
  /** Stream paths of an OLE2 compound file in lower case, or null. */
  cfb(): string[] | null;
  /** Reads a compound-file stream by path (case-insensitive). */
  cfbRead(path: string): Uint8Array | null;
}

export interface SniffRule {
  id: string;
  /** Importer format id. */
  format: string;
  label: string;
  category: SniffCategory;
  /** Extensions (lower case, no dot) this kind of file normally has. */
  ext: string[];
  /** Confidence 0..1, or an object that refines the label. */
  detect(ctx: SniffContext): number | { confidence: number; label?: string; format?: string } | null;
}

const RULES: SniffRule[] = [];

/** Adds (or replaces, by id) a detection rule. */
export function registerSniffer(rule: SniffRule): void {
  const i = RULES.findIndex((r) => r.id === rule.id);
  if (i >= 0) RULES[i] = rule; else RULES.push(rule);
}

export function sniffRules(): readonly SniffRule[] { return RULES; }

// ------------------------------------------------------------------------------------------
// Context

function latin1(u8: Uint8Array): string {
  let s = '';
  for (let i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]);
  return s;
}

function decodeHead(u8: Uint8Array): string {
  if (u8[0] === 0xff && u8[1] === 0xfe) return new TextDecoder('utf-16le').decode(u8.subarray(2));
  if (u8[0] === 0xfe && u8[1] === 0xff) return new TextDecoder('utf-16be').decode(u8.subarray(2));
  if (u8[0] === 0xef && u8[1] === 0xbb && u8[2] === 0xbf) return new TextDecoder('utf-8').decode(u8.subarray(3));
  try { return new TextDecoder('utf-8', { fatal: true }).decode(u8); }
  catch {
    // A multi-byte sequence cut at the 64 KB boundary is not an error.
    try { return new TextDecoder('utf-8', { fatal: true }).decode(u8.subarray(0, Math.max(0, u8.length - 3))); }
    catch { return new TextDecoder('windows-1252').decode(u8); }
  }
}

export function zipEntryNames(u8: Uint8Array): string[] {
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
  for (let p = 0; p + 30 <= u8.length;) {
    if (dv.getUint32(p, true) !== 0x04034b50) { p++; continue; }
    const nl = dv.getUint16(p + 26, true), el = dv.getUint16(p + 28, true), cs = dv.getUint32(p + 18, true);
    names.push(dec.decode(u8.subarray(p + 30, p + 30 + nl)));
    p += 30 + nl + el + (cs > 0 && cs < u8.length ? cs : 0);
    if (names.length > 5000) break;
  }
  return names;
}

export function extOf(name: string): string {
  const base = name.replace(/^.*[\\/]/, '');
  const i = base.lastIndexOf('.');
  return i > 0 && i < base.length - 1 ? base.slice(i + 1).toLowerCase() : '';
}

function makeContext(bytes: Uint8Array, name: string): SniffContext {
  let text: string | null = null, zip: string[] | null | undefined, cfbNames: string[] | null | undefined;
  let cfbObj: ReturnType<typeof readCfb> | undefined;
  const cfb = () => {
    if (cfbObj === undefined) { try { cfbObj = isOle2(bytes) ? readCfb(bytes) : null; } catch { cfbObj = null; } }
    return cfbObj;
  };
  return {
    bytes, name, ext: extOf(name),
    head: latin1(bytes.subarray(0, 4096)),
    text: () => (text ??= decodeHead(bytes.subarray(0, 65536))),
    zip: () => {
      if (zip === undefined) zip = bytes.length > 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && (bytes[2] === 3 || bytes[2] === 5 || bytes[2] === 7) ? zipEntryNames(bytes).map((n) => n.toLowerCase()) : null;
      return zip;
    },
    cfb: () => {
      if (cfbNames === undefined) { const c = cfb(); cfbNames = c ? (c.paths || [...c.streams.keys()]).map((n) => n.toLowerCase()) : null; }
      return cfbNames;
    },
    cfbRead: (path) => { const c = cfb(); return c?.read?.(path) ?? null; },
  };
}

// ------------------------------------------------------------------------------------------
// Sniffing

/** Extensions that say nothing about the content. */
const NEUTRAL_EXT = new Set(['', 'dat', 'bin', 'tmp', 'bak', 'old', 'file', 'download', 'part', 'crdownload', 'data', 'export', 'out']);

/** Identifies a file from its content, using the extension only to break ties. */
export function sniff(data: ArrayBuffer | Uint8Array, name = ''): SniffResult {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const ctx = makeContext(bytes, name);
  const scored: { rule: SniffRule; c: number; label: string; format: string }[] = [];
  for (const rule of RULES) {
    let r: ReturnType<SniffRule['detect']>;
    try { r = rule.detect(ctx); } catch { r = 0; }
    if (!r) continue;
    const c = typeof r === 'number' ? r : r.confidence;
    if (c <= 0) continue;
    scored.push({ rule, c: Math.min(1, c), label: (typeof r === 'object' && r.label) || rule.label, format: (typeof r === 'object' && r.format) || rule.format });
  }
  const ext = ctx.ext;
  const byExt = RULES.filter((r) => r.ext.includes(ext));
  scored.sort((a, b) => b.c - a.c);
  let best = scored[0];
  if (best) {
    // Tie-break: among near-equal candidates prefer one whose extensions include this one.
    const near = scored.filter((s) => best!.c - s.c <= 0.05);
    const extHit = near.find((s) => s.rule.ext.includes(ext));
    if (extHit) best = extHit;
  }
  if (!best || best.c < 0.25) {
    const er = byExt[0];
    if (er) return { format: er.format, label: er.label, confidence: 0.2, category: er.category, byExtension: true, rule: er.id, note: bytes.length ? undefined : 'The file is empty.' };
    if (best) return result(best, ctx);
    return { format: 'unknown', label: 'Unrecognized data', confidence: 0, category: 'unknown', rule: 'none' };
  }
  return result(best, ctx);
}

function result(best: { rule: SniffRule; c: number; label: string; format: string }, ctx: SniffContext): SniffResult {
  // A rule that refines its format (an OpenDocument package that is a text, not a spreadsheet) takes that format's category.
  const category = best.format === best.rule.format ? best.rule.category : (RULES.find((r) => r.format === best.format)?.category ?? best.rule.category);
  const out: SniffResult = { format: best.format, label: best.label, confidence: best.c, category, rule: best.rule.id };
  const ext = ctx.ext;
  if (!ext && best.c >= 0.5) out.note = `The file has no extension; its contents show it is a ${best.label}.`;
  else if (!NEUTRAL_EXT.has(ext) && !best.rule.ext.includes(ext) && best.c >= 0.5) {
    const claimed = RULES.find((r) => r.ext.includes(ext));
    // Plain-text kinds are interchangeable; only mention real disagreements.
    const textish = (id?: string) => !!id && ['csv', 'tsv', 'txt', 'text-table'].includes(id);
    if (claimed && !(textish(claimed.id) && textish(best.rule.id)) && claimed.format !== best.format) {
      out.note = `"${ctx.name.replace(/^.*[\\/]/, '')}" is named .${ext} (${claimed.label}) but contains a ${best.label}; it was opened by its contents.`;
    }
  }
  return out;
}

// ------------------------------------------------------------------------------------------
// Helpers for rules

const u16 = (b: Uint8Array, o: number) => (o + 2 <= b.length ? b[o] | (b[o + 1] << 8) : -1);
const u32 = (b: Uint8Array, o: number) => (o + 4 <= b.length ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0 : -1);
const startsWith = (b: Uint8Array, sig: number[], at = 0) => sig.every((x, i) => b[at + i] === x);
const zipHas = (ctx: SniffContext, test: (n: string) => boolean) => !!ctx.zip()?.some(test);
/** The ODF/OOo "mimetype" entry is stored first and uncompressed. */
const zipMime = (ctx: SniffContext): string => {
  const m = /^PK\x03\x04[\s\S]{26}mimetype([\x20-\x7e]{5,120})/.exec(ctx.head);
  return m ? m[1].replace(/PK[\s\S]*$/, '') : '';
};
const cfbHas = (ctx: SniffContext, test: (n: string) => boolean) => !!ctx.cfb()?.some(test);

/** RFC 822 header lines at the start of a text. */
function headerFields(t: string): string[] {
  const lines = t.replace(/^﻿/, '').split(/\r?\n/, 40);
  const out: string[] = [];
  for (const l of lines) {
    if (!l) break;
    if (/^[ \t]/.test(l)) continue;
    const m = /^([A-Za-z][A-Za-z0-9-]{1,40}):/.exec(l);
    if (!m) return out.length >= 2 ? out : [];
    out.push(m[1].toLowerCase());
  }
  return out;
}

/** BIFF BOF at offset 0 of a stream: version 2 (0x0009), 3 (0x0209), 4 (0x0409) or 5/8 (0x0809). */
function biffVersion(b: Uint8Array): { ver: number; dt: number } | null {
  const id = u16(b, 0), len = u16(b, 2);
  if (len < 4 || len > 20) return null;
  if (id === 0x0009) return { ver: 2, dt: u16(b, 6) };
  if (id === 0x0209) return { ver: 3, dt: u16(b, 6) };
  if (id === 0x0409) return { ver: 4, dt: u16(b, 6) };
  if (id === 0x0809) { const v = u16(b, 4); return { ver: v === 0x0600 ? 8 : 5, dt: u16(b, 6) }; }
  return null;
}

/** dBASE header: version byte, plausible YYMMDD, header and record sizes consistent with the length. */
function dbfScore(b: Uint8Array): number {
  if (b.length < 33) return 0;
  const v = b[0];
  if (![0x02, 0x03, 0x04, 0x05, 0x07, 0x30, 0x31, 0x32, 0x43, 0x63, 0x83, 0x8b, 0x8e, 0xcb, 0xf5, 0xfb].includes(v)) return 0;
  if (v === 0x02) {
    // dBASE II: record count (u16), date, record length; field descriptors from byte 8.
    const n = u16(b, 1), rl = u16(b, 6);
    if (rl < 1 || rl > 1000 || b[3] > 12 || b[4] > 31) return 0;
    return 521 + n * rl <= b.length + 1 && b.length <= 521 + n * rl + 2 ? 0.7 : 0.3;
  }
  const mm = b[2], dd = b[3];
  if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return 0;
  const n = u32(b, 4), hl = u16(b, 8), rl = u16(b, 10);
  if (hl < 33 || rl < 1 || hl > b.length) return 0;
  if ((hl - 1) % 32 !== 0 && (hl - 264 - 1) % 32 !== 0 && v !== 0x30 && v !== 0x31 && v !== 0x32) return 0.2;
  const expect = hl + n * rl;
  if (b.length >= expect && b.length <= expect + 2) return 0.9;
  return b[hl - 1] === 0x0d || b[hl - 264 - 1] === 0x0d ? 0.6 : 0.25;
}

/** VisiCalc saves cells as text lines such as ">A1:/FR"Total" and ends with window and global commands. */
function visicalcScore(t: string): number {
  const lines = t.split(/\r\n?|\n/).filter((l) => l.length).slice(0, 50);
  if (lines.length < 2) return 0;
  const cells = lines.filter((l) => /^>[A-Z]{1,2}\d{1,3}:/.test(l)).length;
  if (!cells) return 0;
  return cells >= Math.min(3, lines.length) * 0.6 ? 0.85 : 0.4;
}

// ------------------------------------------------------------------------------------------
// The table

const S = 'spreadsheet', D = 'document', P = 'presentation', E = 'email';

const BUILTIN: SniffRule[] = [
  // ---- Office Open XML and other ZIP packages
  { id: 'xlsx', format: 'xlsx', label: 'Excel workbook (XLSX)', category: S, ext: ['xlsx', 'xlam'], detect: (c) => {
    if (!zipHas(c, (n) => n.startsWith('xl/'))) return 0;
    if (zipHas(c, (n) => n === 'xl/workbook.bin')) return { confidence: 0.95, format: 'xlsb', label: 'Excel binary workbook (XLSB)' };
    if (zipHas(c, (n) => n === 'xl/vbaproject.bin')) return { confidence: 0.95, format: 'xlsm', label: 'Excel macro-enabled workbook (XLSM)' };
    return 0.95;
  } },
  { id: 'xlsm', format: 'xlsm', label: 'Excel macro-enabled workbook (XLSM)', category: S, ext: ['xlsm'], detect: (c) => (zipHas(c, (n) => n.startsWith('xl/')) ? 0.9 : 0) },
  { id: 'xltx', format: 'xltx', label: 'Excel template (XLTX)', category: S, ext: ['xltx', 'xltm'], detect: (c) => (zipHas(c, (n) => n.startsWith('xl/')) ? 0.9 : 0) },
  { id: 'xlsb', format: 'xlsb', label: 'Excel binary workbook (XLSB)', category: S, ext: ['xlsb'], detect: (c) => (zipHas(c, (n) => n === 'xl/workbook.bin') ? 0.96 : 0) },
  { id: 'docx', format: 'docx', label: 'Word document (DOCX)', category: D, ext: ['docx', 'docm', 'dotx', 'dotm'], detect: (c) => (zipHas(c, (n) => n === 'word/document.xml' || /^word\/document\d*\.xml$/.test(n)) ? 0.95 : 0) },
  { id: 'pptx', format: 'pptx', label: 'PowerPoint presentation (PPTX)', category: P, ext: ['pptx', 'pptm', 'potx', 'potm', 'ppsx', 'ppsm'], detect: (c) => (zipHas(c, (n) => n === 'ppt/presentation.xml') ? 0.95 : 0) },
  { id: 'vsdx', format: 'vsdx', label: 'Visio drawing (VSDX)', category: 'unknown', ext: ['vsdx'], detect: (c) => (zipHas(c, (n) => n.startsWith('visio/')) ? 0.9 : 0) },
  { id: 'odf', format: 'ods', label: 'OpenDocument spreadsheet (ODS)', category: S, ext: ['ods', 'ots'], detect: (c) => {
    const m = zipMime(c);
    const names = m ? null : c.zip();
    const hasContent = !!names?.includes('content.xml');
    if (/opendocument\.spreadsheet/.test(m)) return 0.97;
    if (/opendocument\.text/.test(m)) return { confidence: 0.97, format: 'odt', label: 'OpenDocument text (ODT)' };
    if (/opendocument\.presentation/.test(m)) return { confidence: 0.97, format: 'odp', label: 'OpenDocument presentation (ODP)' };
    if (/opendocument\.graphics/.test(m)) return { confidence: 0.9, format: 'odg', label: 'OpenDocument drawing (ODG)' };
    if (/sun\.xml\.calc/.test(m)) return { confidence: 0.95, format: 'sxc', label: 'StarOffice/OpenOffice.org 1 spreadsheet (SXC)' };
    if (/sun\.xml\.writer/.test(m)) return { confidence: 0.95, format: 'sxw', label: 'StarOffice/OpenOffice.org 1 text (SXW)' };
    if (/sun\.xml\.impress/.test(m)) return { confidence: 0.95, format: 'sxi', label: 'StarOffice/OpenOffice.org 1 presentation (SXI)' };
    if (/application\/x-kspread/.test(m)) return { confidence: 0.9, format: 'ksp', label: 'KSpread spreadsheet' };
    if (/epub\+zip/.test(m)) return { confidence: 0.95, format: 'epub', label: 'EPUB e-book' };
    if (hasContent && !names!.some((n) => n.startsWith('xl/') || n.startsWith('word/') || n.startsWith('ppt/'))) return 0.5;
    return 0;
  } },
  { id: 'odt', format: 'odt', label: 'OpenDocument text (ODT)', category: D, ext: ['odt', 'ott'], detect: () => 0 },
  { id: 'odp', format: 'odp', label: 'OpenDocument presentation (ODP)', category: P, ext: ['odp', 'otp'], detect: () => 0 },
  { id: 'sxc', format: 'sxc', label: 'StarOffice/OpenOffice.org 1 spreadsheet (SXC)', category: S, ext: ['sxc', 'stc'], detect: () => 0 },
  { id: 'sxw', format: 'sxw', label: 'StarOffice/OpenOffice.org 1 text (SXW)', category: D, ext: ['sxw', 'stw'], detect: () => 0 },
  { id: 'sxi', format: 'sxi', label: 'StarOffice/OpenOffice.org 1 presentation (SXI)', category: P, ext: ['sxi', 'sti'], detect: () => 0 },
  { id: 'epub', format: 'epub', label: 'EPUB e-book', category: D, ext: ['epub'], detect: () => 0 },
  { id: 'qgw', format: 'qgw', label: 'QuillGrid workbook', category: S, ext: ['qgw', 'qlw'], detect: (c) => (zipHas(c, (n) => n === 'workbook.json') ? 0.97 : 0) },
  { id: 'iwork', format: 'numbers', label: 'Apple Numbers spreadsheet', category: S, ext: ['numbers'], detect: (c) => {
    const n = c.zip();
    if (!n) return 0;
    const iwa = n.some((x) => x.startsWith('index/') && x.endsWith('.iwa')) || n.includes('index.zip');
    const old = n.includes('index.xml') || n.includes('index.xml.gz') || n.includes('index.apxl') || n.includes('index.apxl.gz');
    if (!iwa && !old) return 0;
    if (n.some((x) => /^index\/(calculationengine|tables\/)/.test(x)) || (old && c.ext === 'numbers')) return 0.9;
    if (n.some((x) => /^index\/(slide|masterslide)/.test(x)) || n.includes('index.apxl') || n.includes('index.apxl.gz')) return { confidence: 0.9, format: 'key', label: 'Apple Keynote presentation' };
    return { confidence: 0.85, format: 'pages', label: 'Apple Pages document' };
  } },
  { id: 'pages', format: 'pages', label: 'Apple Pages document', category: D, ext: ['pages'], detect: () => 0 },
  { id: 'key', format: 'key', label: 'Apple Keynote presentation', category: P, ext: ['key', 'kth'], detect: () => 0 },

  // ---- OLE2 compound files
  { id: 'xls', format: 'xls', label: 'Excel 97-2003 workbook (XLS)', category: S, ext: ['xls', 'xlt', 'xla', 'et', 'ett'], detect: (c) => {
    const n = c.cfb();
    if (!n) return 0;
    if (n.includes('workbook')) {
      if (c.ext === 'xlr') return { confidence: 0.9, format: 'xls', label: 'Microsoft Works spreadsheet (XLR)' };
      if (c.ext === 'et' || c.ext === 'ett') return { confidence: 0.92, label: 'WPS Spreadsheets workbook (ET)' };
      return 0.95;
    }
    if (n.includes('book')) return { confidence: 0.93, label: 'Excel 5.0/95 workbook (XLS, BIFF5)' };
    if (n.includes('encryptedpackage')) {
      if (['docx', 'docm'].includes(c.ext)) return { confidence: 0.9, format: 'docx', label: 'Encrypted Word document' };
      if (['pptx', 'pptm'].includes(c.ext)) return { confidence: 0.9, format: 'pptx', label: 'Encrypted PowerPoint presentation' };
      return { confidence: 0.85, format: 'xlsx', label: 'Encrypted Excel workbook' };
    }
    // Works 6-9 (.xlr) and other programs keep a BIFF stream under another name.
    for (const p of n) {
      if (/^(wksssworkbook|workbook\d*|book\d*)$/.test(p)) {
        const s = c.cfbRead(p);
        if (s && biffVersion(s)) return { confidence: 0.85, label: p.startsWith('wks') ? 'Microsoft Works spreadsheet (XLR)' : 'Excel workbook (XLS)' };
      }
    }
    return 0;
  } },
  { id: 'xlr', format: 'xls', label: 'Microsoft Works spreadsheet (XLR)', category: S, ext: ['xlr'], detect: () => 0 },
  { id: 'qpw', format: 'quattro', label: 'Quattro Pro notebook (QPW)', category: S, ext: ['qpw'], detect: (c) => (cfbHas(c, (n) => n === 'nativecontent_main' || n === 'perfectoffice_main') ? 0.92 : 0) },
  { id: 'doc', format: 'doc', label: 'Word 97-2003 document (DOC)', category: D, ext: ['doc', 'dot', 'wiz'], detect: (c) => {
    if (cfbHas(c, (n) => n === 'worddocument')) return 0.93;
    // Word for Windows 1/2 and Word for DOS/Write files are not compound files.
    const b = c.bytes;
    if ((b[0] === 0xdb || b[0] === 0x9b) && b[1] === 0xa5) return { confidence: 0.8, label: 'Word for Windows 1/2 document (DOC)' };
    if ((b[0] === 0x31 || b[0] === 0x32) && b[1] === 0xbe && b[2] === 0 && b[3] === 0 && b[4] === 0 && b[5] === 0xab) {
      return c.ext === 'wri' ? { confidence: 0.85, format: 'wordtext', label: 'Windows Write document (WRI)' } : { confidence: 0.75, label: 'Word for DOS/Windows 1 document (DOC)' };
    }
    return 0;
  } },
  { id: 'wri', format: 'wordtext', label: 'Windows Write document (WRI)', category: D, ext: ['wri'], detect: () => 0 },
  { id: 'ppt', format: 'ppt', label: 'PowerPoint 97-2003 presentation (PPT)', category: P, ext: ['ppt', 'pps', 'pot'], detect: (c) => (cfbHas(c, (n) => n === 'powerpoint document') ? 0.93 : 0) },
  { id: 'msg', format: 'msg', label: 'Outlook message (MSG)', category: E, ext: ['msg', 'oft'], detect: (c) => (cfbHas(c, (n) => n.startsWith('__substg1.0_') || n === '__properties_version1.0') ? 0.93 : 0) },
  { id: 'works-wps', format: 'wordtext', label: 'Microsoft Works word processor document (WPS)', category: D, ext: ['wps'], detect: (c) => {
    if (cfbHas(c, (n) => n === 'mn0' || (n === 'contents' && c.ext === 'wps'))) return 0.85;
    // Works 2/3 for DOS and Windows word processor files.
    const b = c.bytes;
    if (c.ext === 'wps' && ((b[0] === 0x01 && b[1] === 0xfe) || (b[0] === 0xff && b[1] === 0xff))) return 0.4;
    return 0;
  } },
  { id: 'works-wdb', format: 'works-db', label: 'Microsoft Works database (WDB)', category: S, ext: ['wdb'], detect: (c) => {
    if (u16(c.bytes, 0) === 0x5420) return 0.85;
    if (c.ext === 'wdb' && c.cfb()) return 0.6;
    return 0;
  } },
  { id: 'starcalc', format: 'sdc', label: 'StarCalc spreadsheet (SDC)', category: S, ext: ['sdc', 'vor'], detect: (c) => (cfbHas(c, (n) => n === 'starcalcdocument') ? 0.93 : 0) },
  { id: 'starwriter', format: 'wordtext', label: 'StarWriter document (SDW)', category: D, ext: ['sdw'], detect: (c) => (cfbHas(c, (n) => n === 'starwriterdocument') ? 0.93 : 0) },
  { id: 'starimpress', format: 'sdd', label: 'StarImpress presentation (SDD)', category: P, ext: ['sdd'], detect: (c) => (cfbHas(c, (n) => n === 'starimpressdocument' || n === 'stardrawdocument3') ? 0.9 : 0) },
  { id: 'hwp', format: 'wordtext', label: 'Hangul word processor document (HWP)', category: D, ext: ['hwp'], detect: (c) => (c.head.startsWith('HWP Document File') || (cfbHas(c, (n) => n === 'fileheader') && cfbHas(c, (n) => n.startsWith('bodytext/'))) ? 0.85 : 0) },
  { id: 'mdb', format: 'mdb', label: 'Microsoft Access database', category: 'unknown', ext: ['mdb', 'accdb'], detect: (c) => (/^\x00\x01\x00\x00Standard (Jet|ACE) DB/.test(c.head) ? 0.95 : 0) },

  // ---- Excel BIFF2-5 bare streams (BIFF2-4 were single files, BIFF4W a workbook)
  { id: 'biff', format: 'xls', label: 'Excel worksheet', category: S, ext: ['xls', 'xlw', 'xlc', 'xlm'], detect: (c) => {
    const v = biffVersion(c.bytes);
    if (!v) return 0;
    if (v.ver >= 5) return { confidence: 0.6, label: `Excel ${v.ver === 8 ? '97-2003' : '5.0/95'} workbook stream (BIFF${v.ver})` };
    if (v.ver === 4 && v.dt === 0x0100) return { confidence: 0.92, label: 'Excel 4.0 workbook (XLW, BIFF4W)' };
    const kind = v.dt === 0x0040 ? 'macro sheet' : v.dt === 0x0020 ? 'chart' : 'worksheet';
    return { confidence: 0.9, label: `Excel ${v.ver === 2 ? '2.x' : v.ver === 3 ? '3.0' : '4.0'} ${kind} (BIFF${v.ver})` };
  } },

  // ---- Lotus-format records
  { id: 'lotus', format: 'lotus', label: 'Lotus 1-2-3 worksheet', category: S, ext: ['wk1', 'wks', 'wk3', 'wk4', '123', 'wk2', 'wrk', 'wr1', 'wku', 'wt4', '12m'], detect: (c) => {
    const v = lotusVariant(c.bytes);
    if (!v) return 0;
    if (v === 'works-dos' || v === 'works-win') return { confidence: 0.92, format: 'works', label: LOTUS_VARIANT_LABEL[v] };
    if (v === 'wq1' || v === 'wq2') return { confidence: 0.92, format: 'quattro', label: LOTUS_VARIANT_LABEL[v] };
    if (v === 'wk1' && /^wr/.test(c.ext)) return { confidence: 0.92, label: 'Symphony worksheet (WR1)' };
    return { confidence: 0.92, label: LOTUS_VARIANT_LABEL[v] };
  } },
  { id: 'works', format: 'works', label: 'Microsoft Works spreadsheet (WKS)', category: S, ext: ['wks'], detect: () => 0 },
  { id: 'lotus-fmt', format: 'lotus-fmt', label: 'Lotus 1-2-3 formatting file (FMT/FM3)', category: S, ext: ['fmt', 'fm3', 'all'], detect: (c) => {
    const op = u16(c.bytes, 0), len = u16(c.bytes, 2), v = u16(c.bytes, 4);
    return op === 0 && (len === 2 || len === 26) && (v >= 0x8000 && v <= 0x8007) ? 0.8 : 0;
  } },
  { id: 'quattro', format: 'quattro', label: 'Quattro Pro spreadsheet', category: S, ext: ['wq1', 'wq2', 'wb1', 'wb2', 'wb3', 'wkq', 'wkp'], detect: (c) => {
    const b = c.bytes;
    if (u16(b, 0) !== 0 || u16(b, 2) !== 2) return 0;
    const v = u16(b, 4);
    if (v === 0x5120) return { confidence: 0.92, label: 'Quattro / Quattro Pro for DOS spreadsheet (WQ1)' };
    if (v === 0x5121) return { confidence: 0.92, label: 'Quattro Pro 5 for DOS spreadsheet (WQ2)' };
    if (v >= 0x1001 && v <= 0x1007) return { confidence: 0.9, label: `Quattro Pro for Windows notebook (WB${v === 0x1001 ? 1 : v === 0x1002 ? 2 : 3})` };
    return 0;
  } },

  // ---- other spreadsheet programs
  { id: 'dbf', format: 'dbf', label: 'dBASE / FoxPro / Clipper table (DBF)', category: S, ext: ['dbf', 'dbase'], detect: (c) => {
    const s = dbfScore(c.bytes);
    if (!s) return 0;
    const v = c.bytes[0];
    const label = v === 0x02 ? 'dBASE II table (DBF)' : v === 0x30 || v === 0x31 || v === 0x32 ? 'Visual FoxPro table (DBF)' : v === 0xf5 || v === 0xfb ? 'FoxPro table (DBF)' : v === 0x04 || v === 0x8b || v === 0x7b ? 'dBASE IV table (DBF)' : v === 0x05 ? 'dBASE 5 table (DBF)' : 'dBASE III table (DBF)';
    return { confidence: s, label };
  } },
  { id: 'visicalc', format: 'visicalc', label: 'VisiCalc worksheet (VC)', category: S, ext: ['vc', 'vcf'], detect: (c) => (c.bytes[0] === 0x3e ? visicalcScore(c.text()) : 0) },
  { id: 'cwk', format: 'cwk', label: 'AppleWorks/ClarisWorks document (CWK)', category: S, ext: ['cwk', 'cws', 'cwd'], detect: (c) => (c.head.slice(4, 8) === 'BOBO' ? 0.85 : 0) },
  { id: 'multiplan', format: 'multiplan', label: 'Multiplan worksheet', category: S, ext: ['mp', 'mpl'], detect: (c) => {
    const b = c.bytes;
    if (b[0] === 0x0c && b[1] === 0xed) return { confidence: 0.75, label: 'Multiplan for DOS worksheet' };
    if (/^ID;PMP/.test(c.head)) return { confidence: 0.9, format: 'sylk', label: 'Multiplan worksheet (SYLK)' };
    return 0;
  } },
  { id: 'supercalc', format: 'supercalc', label: 'SuperCalc worksheet', category: S, ext: ['cal', 'sc5', 'sc4'], detect: (c) => (/^Supercalc/i.test(c.head) || /SuperCalc/.test(c.head.slice(0, 64)) ? 0.7 : 0) },
  { id: 'gnumeric', format: 'gnumeric', label: 'Gnumeric workbook', category: S, ext: ['gnumeric'], detect: (c) => {
    if (c.bytes[0] === 0x1f && c.bytes[1] === 0x8b) return c.ext === 'gnumeric' ? 0.8 : 0.1;
    return /<gnm:Workbook/.test(c.head) ? 0.95 : 0;
  } },
  { id: 'wingz', format: 'wingz', label: 'Wingz worksheet', category: S, ext: ['wkz'], detect: () => 0 },
  { id: 'planmaker', format: 'planmaker', label: 'PlanMaker workbook', category: S, ext: ['pmd', 'pmdx', 'pmv'], detect: () => 0 },
  { id: 'uof', format: 'uof', label: 'Uniform Office Format document', category: S, ext: ['uos', 'uof'], detect: (c) => (/uof:UOF/.test(c.head) ? 0.8 : 0) },

  // ---- word processors
  { id: 'rtf', format: 'rtf', label: 'Rich Text Format document (RTF)', category: D, ext: ['rtf'], detect: (c) => (c.head.startsWith('{\\rtf') ? 0.95 : 0) },
  { id: 'wpd', format: 'wordtext', label: 'WordPerfect document (WPD)', category: D, ext: ['wpd', 'wp', 'wp5', 'wp6', 'wpt'], detect: (c) => {
    if (c.head.startsWith('\xffWPC')) {
      const prod = c.bytes[8], type = c.bytes[9];
      if (prod === 1 && type === 10) return 0.95;
      if (type === 10 || type === 0x0a) return 0.9;
      // Other WordPerfect Corporation products share the signature (Presentations, Quattro Pro DOS graphs, ...).
      return { confidence: 0.6, label: 'WordPerfect Corporation file' };
    }
    return 0;
  } },
  { id: 'abw', format: 'wordtext', label: 'AbiWord document (ABW)', category: D, ext: ['abw', 'zabw'], detect: (c) => (/<abiword\b/.test(c.head) ? 0.9 : 0) },
  { id: 'fodt', format: 'odt', label: 'Flat OpenDocument text (FODT)', category: D, ext: ['fodt'], detect: (c) => (/<office:document\b/.test(c.head) && /office:mimetype="application\/vnd\.oasis\.opendocument\.text"/.test(c.text().slice(0, 8000)) ? 0.95 : 0) },
  { id: 'fodp', format: 'odp', label: 'Flat OpenDocument presentation (FODP)', category: P, ext: ['fodp'], detect: (c) => (/<office:document\b/.test(c.head) && /office:mimetype="application\/vnd\.oasis\.opendocument\.presentation"/.test(c.text().slice(0, 8000)) ? 0.95 : 0) },
  { id: 'wordml', format: 'wordtext', label: 'Word 2003 XML document', category: D, ext: [], detect: (c) => (/<\?mso-application\s+progid="Word\.Document"/i.test(c.head) || /<w:wordDocument\b/.test(c.head) ? 0.9 : 0) },
  { id: 'pdf', format: 'pdf', label: 'PDF document', category: D, ext: ['pdf'], detect: (c) => (c.head.startsWith('%PDF-') ? 0.97 : c.head.slice(0, 1024).includes('%PDF-') ? 0.6 : 0) },
  { id: 'lwp', format: 'wordtext', label: 'Lotus Word Pro document (LWP)', category: D, ext: ['lwp'], detect: (c) => (c.head.slice(0, 64).includes('WordPro') ? 0.85 : 0) },
  { id: 'wordtext', format: 'wordtext', label: 'Word processor document', category: D, ext: ['wps', 'sdw', 'hwp', 'wp4', 'ws', 'ws7', 'wsd', 'mcw', 'cwd'], detect: () => 0 },
  { id: 'md', format: 'txt', label: 'Markdown document', category: D, ext: ['md', 'markdown'], detect: () => 0 },

  // ---- e-mail
  { id: 'eml', format: 'eml', label: 'E-mail message (EML)', category: E, ext: ['eml', 'emlx', 'mht', 'mhtml', 'nws'], detect: (c) => {
    let t = c.text();
    if (/^\d+\s*\r?\n/.test(t)) t = t.replace(/^\d+\s*\r?\n/, ''); // emlx byte count line
    const f = headerFields(t);
    if (f.length < 2) return 0;
    const mail = ['from', 'to', 'subject', 'date', 'received', 'return-path', 'message-id', 'mime-version', 'delivered-to', 'x-mailer', 'reply-to', 'cc', 'content-type'];
    const hits = f.filter((x) => mail.includes(x)).length;
    if (hits < 2) return 0;
    if (/^content-type:\s*multipart\/related/im.test(t.slice(0, 4000)) && !f.includes('from') && !f.includes('received')) return { confidence: 0.85, label: 'Web archive (MHT)' };
    return Math.min(0.92, 0.55 + hits * 0.08);
  } },
  { id: 'mbox', format: 'mbox', label: 'Mailbox (MBOX)', category: E, ext: ['mbox', 'mbx', 'mbs', 'mail'], detect: (c) => {
    const t = c.text();
    if (!/^From \S+.*\d{4}\s*\r?\n/.test(t) && !/^From \S+ +\w{3} \w{3}/.test(t)) return 0;
    return headerFields(t.replace(/^From .*\r?\n/, '')).length >= 2 ? 0.92 : 0.6;
  } },
  { id: 'pst', format: 'pst', label: 'Outlook data file (PST/OST)', category: E, ext: ['pst', 'ost'], detect: (c) => (c.head.startsWith('!BDN') ? 0.97 : 0) },
  { id: 'dbx', format: 'dbx', label: 'Outlook Express mailbox (DBX)', category: E, ext: ['dbx'], detect: (c) => (startsWith(c.bytes, [0xcf, 0xad, 0x12, 0xfe]) ? 0.9 : 0) },
  { id: 'tnef', format: 'tnef', label: 'Outlook rich-text attachment (winmail.dat)', category: E, ext: ['tnef'], detect: (c) => (startsWith(c.bytes, [0x78, 0x9f, 0x3e, 0x22]) ? 0.9 : 0) },

  // ---- text formats this app already reads
  { id: 'qgwe', format: 'qgwe', label: 'Encrypted QuillGrid workbook', category: S, ext: ['qgwe', 'qlwe'], detect: (c) => (c.head.startsWith('QFDE1') ? 0.98 : 0) },
  { id: 'sylk', format: 'sylk', label: 'SYLK file (SLK)', category: S, ext: ['slk', 'sylk'], detect: (c) => (/^﻿?ID;/.test(c.text()) ? 0.92 : 0) },
  { id: 'dif', format: 'dif', label: 'Data Interchange Format (DIF)', category: S, ext: ['dif'], detect: (c) => (/^﻿?TABLE\r?\n\s*0\s*,\s*1\s*\r?\n/.test(c.text().trimStart()) ? 0.92 : 0) },
  { id: 'xml2003', format: 'xml', label: 'Excel 2003 XML spreadsheet', category: S, ext: ['xml'], detect: (c) => {
    const t = c.text();
    if (!/urn:schemas-microsoft-com:office:spreadsheet/.test(t.slice(0, 8000))) return 0;
    if (/<html/i.test(t.slice(0, 4000)) && !/<\?mso-application\s+progid="Excel\.Sheet"/i.test(t)) return { confidence: 0.85, format: 'html', label: 'Web page (HTML) saved by Excel' };
    return 0.92;
  } },
  { id: 'fods', format: 'fods', label: 'Flat OpenDocument spreadsheet (FODS)', category: S, ext: ['fods'], detect: (c) => {
    const t = c.text().slice(0, 8000);
    if (!/<office:document\b/.test(t)) return 0;
    if (/opendocument\.text"/.test(t) || /opendocument\.presentation"/.test(t)) return 0;
    return 0.9;
  } },
  { id: 'json', format: 'json', label: 'QuillGrid JSON workbook', category: S, ext: ['json'], detect: (c) => (/^\s*\{/.test(c.text()) && /"format"\s*:\s*"quillgrid"|"qgw"\s*:/.test(c.text().slice(0, 4096)) ? 0.9 : 0) },
  { id: 'html', format: 'html', label: 'Web page (HTML)', category: S, ext: ['html', 'htm', 'xhtml', 'shtml'], detect: (c) => (/^\s*(<\?xml[^>]*>\s*)?(<!--[\s\S]*?-->\s*)*<(!doctype html|html|table|body|head|meta)\b/i.test(c.text().replace(/^﻿/, '')) ? 0.8 : 0) },
  { id: 'csv', format: 'csv', label: 'Comma-separated values (CSV)', category: S, ext: ['csv'], detect: (c) => textTable(c, ',') },
  { id: 'tsv', format: 'tsv', label: 'Tab-separated values (TSV)', category: S, ext: ['tsv', 'tab'], detect: (c) => textTable(c, '\t') },
  { id: 'txt', format: 'txt', label: 'Text file', category: S, ext: ['txt', 'text', 'prn', 'asc'], detect: (c) => (isText(c) ? 0.26 : 0) },
];

function isText(c: SniffContext): boolean {
  const b = c.bytes.subarray(0, 8192);
  if (!b.length) return false;
  if ((b[0] === 0xff && b[1] === 0xfe) || (b[0] === 0xfe && b[1] === 0xff)) return true;
  let bad = 0;
  for (const x of b) if (x < 0x09 || (x > 0x0d && x < 0x20 && x !== 0x1a && x !== 0x1b)) bad++;
  return bad <= b.length * 0.01;
}

function textTable(c: SniffContext, sep: string): number {
  if (!isText(c)) return 0;
  const lines = c.text().split(/\r\n?|\n/).filter((l) => l.length).slice(0, 30);
  if (!lines.length) return 0;
  const counts = lines.map((l) => l.split(sep).length - 1);
  const withSep = counts.filter((n) => n > 0).length;
  if (!withSep) return 0;
  const same = counts.filter((n) => n === counts[0]).length;
  const base = sep === '\t' ? 0.45 : 0.4;
  return withSep >= lines.length * 0.6 ? base + (same / lines.length) * 0.1 : 0.27;
}

for (const r of BUILTIN) registerSniffer(r);
