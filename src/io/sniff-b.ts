/**
 * Content detection for the modern, XML and text spreadsheet formats added in the "B" set:
 * Excel binary workbooks, OpenOffice.org 1.x Calc, Gnumeric, KSpread, Apple Numbers, UOF,
 * SQLite databases, JSON / NDJSON records, fixed-width text, MHTML web archives, gzip
 * containers and DIF variants.
 *
 * Every detector is a small synchronous function of the first bytes (and the ZIP entry names,
 * when the file is a ZIP archive), so it can run on every open without loading the converter
 * modules. SNIFF_RULES_B lists them in priority order for the general content sniffer; sniffB()
 * runs them directly.
 */

export interface SniffInput {
  /** The whole file. */
  u8: Uint8Array;
  /** File name as given (may be empty). */
  name: string;
  /** Lower-case extension without the dot ('' when none). */
  ext: string;
  /** The first 4 KB decoded as text (UTF-8/UTF-16 aware, else Latin-1); BOM removed. */
  head: string;
  /** Lower-case ZIP entry names when the file is a ZIP archive, else null. */
  zip: string[] | null;
}

export interface SniffRule {
  /** Format id in the FORMATS registry. */
  id: string;
  /** Short description of the signature, for documentation and diagnostics. */
  signature: string;
  detect: (x: SniffInput) => boolean;
}

const latin1 = (u8: Uint8Array, a = 0, b = u8.length): string => {
  let s = '';
  for (let i = a; i < Math.min(b, u8.length); i++) s += String.fromCharCode(u8[i]);
  return s;
};

// --------------------------------------------------------------------------------------------
// Binary and package formats

/** Excel binary workbook (.xlsb): an OOXML package whose workbook part is xl/workbook.bin. */
export function isXlsb(x: SniffInput): boolean {
  return !!x.zip && x.zip.some((n) => /(^|\/)workbook\.bin$/.test(n));
}

/** StarOffice / OpenOffice.org 1.x Calc (.sxc/.stc): ZIP whose mimetype is application/vnd.sun.xml.calc. */
export function isSxc(x: SniffInput): boolean {
  if (!x.zip) return false;
  const head = latin1(x.u8, 0, 200);
  if (/mimetypeapplication\/vnd\.sun\.xml\.calc/.test(head)) return true;
  return x.zip.includes('content.xml') && !/opendocument/.test(head) && (x.ext === 'sxc' || x.ext === 'stc');
}

/** KSpread / KOffice / Calligra Sheets native (.ksp): KoStore ZIP or tar.gz with maindoc.xml, or bare XML. */
export function isKspread(x: SniffInput): boolean {
  if (x.zip) {
    const head = latin1(x.u8, 0, 200);
    if (/mimetypeapplication\/(x-kspread|vnd\.kde\.kspread)/.test(head)) return true;
    return x.zip.includes('maindoc.xml') && (x.ext === 'ksp' || x.ext === 'kspd');
  }
  if (isGzipBytes(x.u8)) return x.ext === 'ksp';
  return /^\s*(<\?xml[^>]*>\s*)?(<!DOCTYPE\s+spreadsheet\b|<spreadsheet\b[^>]*\b(mime="application\/x-kspread"|editor="KSpread"))/i.test(x.head);
}

/** Apple Numbers: '09 (index.xml[.gz]) or 2013+ (Index/*.iwa, or Index.zip inside the package). */
export function isNumbers(x: SniffInput): boolean {
  if (x.zip) {
    const z = x.zip;
    if (z.some((n) => /^(.*\/)?index\/(document|tables\/.*)\.iwa$/.test(n)) || z.some((n) => /(^|\/)index\.zip$/.test(n))) return true;
    if (z.some((n) => /(^|\/)index\.xml(\.gz)?$/.test(n)) && (x.ext === 'numbers' || z.some((n) => /buildversionhistory\.plist$|quicklook\//.test(n)))) return true;
    return x.ext === 'numbers' && z.some((n) => /quicklook\/(thumbnail|preview)\./.test(n));
  }
  return /xmlns:ls="http:\/\/developer\.apple\.com\/namespaces\/ls"/.test(x.head) && /<ls:document\b/.test(x.head);
}

/** SQLite 3 database file. */
export function isSqlite(x: SniffInput): boolean {
  return latin1(x.u8, 0, 16) === 'SQLite format 3\0';
}

export const isGzipBytes = (u8: Uint8Array): boolean => u8.length > 18 && u8[0] === 0x1f && u8[1] === 0x8b && u8[2] === 8;

/** Gnumeric: gzipped or plain XML in the Gnumeric namespace. */
export function isGnumeric(x: SniffInput): boolean {
  if (isGzipBytes(x.u8)) return x.ext === 'gnumeric' || x.ext === 'gnm';
  return /<(gnm|gmr):Workbook\b/.test(x.head) || /xmlns(:\w+)?="http:\/\/www\.gnumeric\.org\/v\d+\.dtd"/.test(x.head) || /xmlns(:\w+)?="http:\/\/www\.gnome\.org\/gnumeric\//.test(x.head);
}

/** Any other gzip file: decompressed and identified again. */
export function isGzipContainer(x: SniffInput): boolean {
  return isGzipBytes(x.u8);
}

// --------------------------------------------------------------------------------------------
// XML and text formats

/** Uniform Office Format spreadsheet (UOF, GB/T 20916; .uos). */
export function isUof(x: SniffInput): boolean {
  return /<uof:UOF\b/.test(x.head) && (/vnd\.uof\.spreadsheet/.test(x.head) || /uof-spreadsheet/.test(x.head));
}

/** MIME HTML web archive (.mht/.mhtml), as saved by Excel's "Single File Web Page". */
export function isMhtml(x: SniffInput): boolean {
  const h = x.head.slice(0, 2048);
  return /^(?:[\w-]+:[^\n]*\r?\n(?:[ \t][^\n]*\r?\n)*)*?MIME-Version:/im.test(h) && /Content-Type:\s*multipart\/related/i.test(h);
}

/** DIF (Data Interchange Format) in any of its common header variants. */
export function isDif(x: SniffInput): boolean {
  return /^\s*TABLE\s*\r?\n\s*0\s*,\s*\d+\s*\r?\n/i.test(x.head) && /\bVECTORS\b|\bTUPLES\b|\bDATA\b/i.test(x.head);
}

/** Newline-delimited JSON (.ndjson/.jsonl): several lines, each a JSON object or array. */
export function isNdjson(x: SniffInput): boolean {
  if (x.ext === 'ndjson' || x.ext === 'jsonl') return true;
  const lines = x.head.split(/\r?\n/).filter((l) => l.trim());
  if (lines.length < 2) return false;
  const full = lines.slice(0, x.head.length >= 4096 ? -1 : undefined).slice(0, 20);
  if (full.length < 2) return false;
  return full.every((l) => { const t = l.trim(); if (!/^[[{]/.test(t) || !/[\]}]$/.test(t)) return false; try { JSON.parse(t); return true; } catch { return false; } });
}

/** A JSON document whose top level is an array or an object (records, tables). */
export function isJsonData(x: SniffInput): boolean {
  const t = x.head.trimStart();
  if (t.startsWith('[')) return /^\[\s*([[{"\]]|-?\d|true\b|false\b|null\b)/.test(t);
  if (t.startsWith('{')) return /^\{\s*("|\})/.test(t);
  return false;
}

/**
 * Fixed-width text (.prn "Formatted Text (Space delimited)" and similar reports): no field
 * delimiter is used consistently, and runs of spaces line up into at least two columns.
 */
export function isFixedWidth(x: SniffInput): boolean {
  if (x.zip || x.u8.some((b, i) => i < 2048 && b === 0)) return false;
  const lines = x.head.split(/\r?\n/).slice(0, 40);
  if (x.head.length >= 4096) lines.pop();
  const body = lines.filter((l) => l.trim());
  if (body.length < 3) return false;
  if (/^\s*[<{[]/.test(x.head) || /^ID;/.test(x.head) || isDif(x)) return false;
  for (const d of ['\t', ',', ';', '|']) {
    const counts = body.map((l) => l.split(d).length - 1);
    if (counts.filter((n) => n > 0 && n === counts[0]).length >= body.length * 0.8 && counts[0] > 0) return false;
  }
  return fixedWidthColumns(body).length >= 2;
}

/**
 * Start positions of fixed-width columns: a column starts where every line has a space (or has
 * ended) in the position before and at least one line has text. Requires each boundary gap to be
 * blank in all lines, so ordinary prose (where spaces wander) yields a single column.
 */
export function fixedWidthColumns(lines: string[]): number[] {
  const body = lines.filter((l) => l.trim());
  if (!body.length) return [];
  const width = Math.max(...body.map((l) => l.length));
  if (width > 2000) return [];
  const blank = new Array<boolean>(width).fill(true);
  for (const l of body) for (let i = 0; i < l.length; i++) if (l[i] !== ' ') blank[i] = false;
  const starts: number[] = [];
  for (let i = 0; i < width; i++) if (!blank[i] && (i === 0 || blank[i - 1])) starts.push(i);
  // Each column must hold text in most lines, or it is a stray word.
  const ok = starts.filter((s, k) => {
    const e = k + 1 < starts.length ? starts[k + 1] : width;
    return body.filter((l) => l.slice(s, e).trim()).length >= Math.max(2, body.length * 0.5);
  });
  return ok.length >= 2 ? ok : [];
}

// --------------------------------------------------------------------------------------------
// Rules

/** Detection rules, most specific first. */
export const SNIFF_RULES_B: SniffRule[] = [
  { id: 'xlsb', signature: 'ZIP package containing xl/workbook.bin', detect: isXlsb },
  { id: 'sxc', signature: 'ZIP with mimetype application/vnd.sun.xml.calc', detect: isSxc },
  { id: 'ksp', signature: 'KoStore ZIP (mimetype application/x-kspread) or maindoc.xml; <spreadsheet> XML', detect: isKspread },
  { id: 'numbers', signature: 'ZIP with Index/*.iwa, Index.zip or index.xml[.gz]; <ls:document> XML', detect: isNumbers },
  { id: 'sqlite', signature: '"SQLite format 3\\0" at offset 0', detect: isSqlite },
  { id: 'gnumeric', signature: 'gzip + .gnumeric/.gnm, or XML in the Gnumeric namespace', detect: isGnumeric },
  { id: 'gz', signature: 'gzip magic 1F 8B 08 (decompressed and identified again)', detect: isGzipContainer },
  { id: 'uos', signature: '<uof:UOF> root with the spreadsheet MIME type', detect: isUof },
  { id: 'mht', signature: 'MIME-Version header with multipart/related', detect: isMhtml },
  { id: 'dif', signature: 'TABLE / 0,n header followed by VECTORS, TUPLES or DATA', detect: isDif },
  { id: 'ndjson', signature: 'several lines that are each a JSON object or array', detect: isNdjson },
  { id: 'json', signature: 'top-level JSON array or object', detect: isJsonData },
  { id: 'prn', signature: 'text whose blank columns line up in every line', detect: isFixedWidth },
];

function headText(u8: Uint8Array): string {
  const n = Math.min(u8.length, 4096);
  const s = u8.subarray(0, n);
  try {
    if (s[0] === 0xff && s[1] === 0xfe) return new TextDecoder('utf-16le').decode(s.subarray(2));
    if (s[0] === 0xfe && s[1] === 0xff) return new TextDecoder('utf-16be').decode(s.subarray(2));
    if (s[0] === 0xef && s[1] === 0xbb && s[2] === 0xbf) return new TextDecoder('utf-8').decode(s.subarray(3));
    return new TextDecoder('utf-8', { fatal: false }).decode(s);
  } catch { return latin1(s); }
}

/** Builds the detector input; zipNames is passed in so the caller's ZIP listing is reused. */
export function sniffInput(u8: Uint8Array, name: string, zipNames: (u8: Uint8Array) => string[]): SniffInput {
  const lower = name.toLowerCase().replace(/^.*[\\/]/, '');
  const ext = lower.includes('.') ? lower.slice(lower.lastIndexOf('.') + 1) : '';
  const isZip = u8.length >= 4 && u8[0] === 0x50 && u8[1] === 0x4b && (u8[2] === 3 || u8[2] === 5 || u8[2] === 7);
  return { u8, name, ext, head: headText(u8), zip: isZip ? zipNames(u8).map((n) => n.toLowerCase()) : null };
}

/**
 * Runs the B rules. Binary/package rules always apply; the text rules apply only when the
 * extension does not already name a text format they would steal from (a .csv stays CSV).
 */
export function sniffB(u8: Uint8Array, name: string, zipNames: (u8: Uint8Array) => string[]): string | null {
  const x = sniffInput(u8, name, zipNames);
  for (const r of SNIFF_RULES_B) {
    if (r.id === 'prn' && ['csv', 'tsv', 'tab', 'slk', 'sylk', 'dif', 'html', 'htm', 'xml', 'json'].includes(x.ext)) continue;
    if ((r.id === 'json' || r.id === 'ndjson') && ['csv', 'tsv', 'tab', 'txt'].includes(x.ext) && !x.head.trimStart().startsWith('[')) continue;
    try { if (r.detect(x)) return r.id; } catch { /* a detector must never break opening */ }
  }
  return null;
}
