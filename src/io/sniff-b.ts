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

/**
 * What a detector sees. The shape matches the general sniffer's context (sniff.ts), so these
 * rules can be registered there unchanged: bytes, name, lower-case extension, decoded text of
 * the first 64 KB and the ZIP entry names.
 */
export interface SniffInput {
  bytes: Uint8Array;
  name: string;
  /** Lower-case extension without the dot ('' when none). */
  ext: string;
  /** Text of the start of the file (BOM-aware UTF-8/UTF-16, else Latin-1). */
  text(): string;
  /** Lower-case ZIP entry names, or null when the file is not a ZIP archive. */
  zip(): string[] | null;
}

export interface SniffRule {
  id: string;
  /** Importer format id in the FORMATS registry. */
  format: string;
  /** What the file is. */
  label: string;
  category: 'spreadsheet';
  /** Extensions (lower case, no dot) this kind of file normally has. */
  ext: string[];
  /** Short description of the signature, for documentation and diagnostics. */
  signature: string;
  /** Confidence 0..1 when the content matches, else null. */
  detect: (x: SniffInput) => number | null;
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
  const z = x.zip();
  return !!z && z.some((n) => /(^|\/)workbook\.bin$/.test(n));
}

/** StarOffice / OpenOffice.org 1.x Calc (.sxc/.stc): ZIP whose mimetype is application/vnd.sun.xml.calc. */
export function isSxc(x: SniffInput): boolean {
  const z = x.zip();
  if (!z) return false;
  const head = latin1(x.bytes, 0, 200);
  if (/mimetypeapplication\/vnd\.sun\.xml\.calc/.test(head)) return true;
  return z.includes('content.xml') && !/opendocument/.test(head) && (x.ext === 'sxc' || x.ext === 'stc');
}

/** KSpread / KOffice / Calligra Sheets native (.ksp): KoStore ZIP or tar.gz with maindoc.xml, or bare XML. */
export function isKspread(x: SniffInput): boolean {
  const z = x.zip();
  if (z) {
    const head = latin1(x.bytes, 0, 200);
    if (/mimetypeapplication\/(x-kspread|vnd\.kde\.kspread)/.test(head)) return true;
    return z.includes('maindoc.xml') && (x.ext === 'ksp' || x.ext === 'kspd');
  }
  if (isGzipBytes(x.bytes)) return x.ext === 'ksp';
  return /^\s*(<\?xml[^>]*>\s*)?(<!DOCTYPE\s+spreadsheet\b|<spreadsheet\b[^>]*\b(mime="application\/x-kspread"|editor="KSpread"))/i.test(x.text().slice(0, 4096));
}

/** Apple Numbers: '09 (index.xml[.gz]) or 2013+ (Index/*.iwa, or an inner Index.zip). */
export function isNumbers(x: SniffInput): boolean {
  const z = x.zip();
  if (z) {
    if (z.some((n) => /^(.*\/)?index\/(document|tables\/.*)\.iwa$/.test(n)) || z.some((n) => /(^|\/)index\.zip$/.test(n))) return true;
    if (z.some((n) => /(^|\/)index\.xml(\.gz)?$/.test(n)) && (x.ext === 'numbers' || z.some((n) => /buildversionhistory\.plist$|quicklook\//.test(n)))) return true;
    return x.ext === 'numbers' && z.some((n) => /quicklook\/(thumbnail|preview)\.|^preview\./.test(n));
  }
  const head = x.text().slice(0, 4096);
  return /xmlns:ls="http:\/\/developer\.apple\.com\/namespaces\/ls"/.test(head) && /<ls:document\b/.test(head);
}

/** SQLite 3 database file. */
export function isSqlite(x: SniffInput): boolean {
  return latin1(x.bytes, 0, 16) === 'SQLite format 3\0';
}

export const isGzipBytes = (u8: Uint8Array): boolean => u8.length > 18 && u8[0] === 0x1f && u8[1] === 0x8b && u8[2] === 8;

/** Gnumeric: gzipped (by extension) or plain XML in the Gnumeric namespace. */
export function isGnumeric(x: SniffInput): boolean {
  if (isGzipBytes(x.bytes)) return x.ext === 'gnumeric' || x.ext === 'gnm';
  const head = x.text().slice(0, 4096);
  return /<(gnm|gmr):Workbook\b/.test(head) || /xmlns(:\w+)?="http:\/\/www\.gnumeric\.org\/v\d+\.dtd"/.test(head) || /xmlns(:\w+)?="http:\/\/www\.gnome\.org\/gnumeric\//.test(head);
}

/** Any other gzip file: decompressed and identified again. */
export function isGzipContainer(x: SniffInput): boolean {
  return isGzipBytes(x.bytes);
}

// --------------------------------------------------------------------------------------------
// XML and text formats

/** Uniform Office Format spreadsheet (UOF, GB/T 20916; .uos). */
export function isUof(x: SniffInput): boolean {
  const head = x.text().slice(0, 4096);
  return /<uof:UOF\b/.test(head) && /vnd\.uof\.spreadsheet|uof-spreadsheet/.test(head);
}

/** MIME HTML web archive (.mht/.mhtml), as saved by Excel's "Single File Web Page". */
export function isMhtml(x: SniffInput): boolean {
  const h = x.text().slice(0, 2048);
  return /^(?:[\w-]+:[^\n]*\r?\n(?:[ \t][^\n]*\r?\n)*)*?MIME-Version:/im.test(h) && /Content-Type:\s*multipart\/related/i.test(h);
}

/** DIF (Data Interchange Format) in any of its common header variants. */
export function isDif(x: SniffInput): boolean {
  const head = x.text().slice(0, 4096);
  return /^\s*TABLE\s*\r?\n\s*0\s*,\s*\d+\s*\r?\n/i.test(head) && /\bVECTORS\b|\bTUPLES\b|\bDATA\b/i.test(head);
}

/** Newline-delimited JSON (.ndjson/.jsonl): several lines, each a JSON object or array. */
export function isNdjson(x: SniffInput): boolean {
  if (x.ext === 'ndjson' || x.ext === 'jsonl') return true;
  const t = x.text();
  const lines = t.split(/\r?\n/).filter((l) => l.trim());
  if (lines.length < 2) return false;
  const full = lines.slice(0, x.bytes.length > t.length ? -1 : undefined).slice(0, 20);
  if (full.length < 2) return false;
  return full.every((l) => { const s = l.trim(); if (!/^[[{]/.test(s) || !/[\]}]$/.test(s)) return false; try { JSON.parse(s); return true; } catch { return false; } });
}

/** A JSON document whose top level is an array or an object (records, tables). */
export function isJsonData(x: SniffInput): boolean {
  const t = x.text().slice(0, 4096).trimStart();
  if (t.startsWith('[')) return /^\[\s*([[{"\]]|-?\d|true\b|false\b|null\b)/.test(t);
  if (t.startsWith('{')) return /^\{\s*("|\})/.test(t);
  return false;
}

/**
 * Fixed-width text (.prn "Formatted Text (Space delimited)" and similar reports): no field
 * delimiter is used consistently, and runs of spaces line up into at least two columns.
 */
export function isFixedWidth(x: SniffInput): boolean {
  if (x.zip() || x.bytes.subarray(0, 2048).some((b) => b === 0)) return false;
  const head = x.text().slice(0, 4096);
  const lines = head.split(/\r?\n/).slice(0, 40);
  if (head.length >= 4096) lines.pop();
  const body = lines.filter((l) => l.trim());
  if (body.length < 3) return false;
  if (/^\s*[<{[]/.test(head) || /^ID;/.test(head) || isDif(x)) return false;
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

const rule = (id: string, label: string, ext: string[], confidence: number, signature: string, test: (x: SniffInput) => boolean, format = id): SniffRule =>
  ({ id: `b.${id}`, format, label, category: 'spreadsheet', ext, signature, detect: (x) => (test(x) ? confidence : null) });

/** Detection rules, most specific first (format ids are those of the FORMATS registry). */
export const SNIFF_RULES_B: SniffRule[] = [
  rule('xlsb', 'Excel Binary Workbook', ['xlsb'], 0.97, 'ZIP package containing xl/workbook.bin', isXlsb),
  rule('sxc', 'OpenOffice.org 1.x spreadsheet', ['sxc', 'stc'], 0.95, 'ZIP with mimetype application/vnd.sun.xml.calc', isSxc),
  rule('ksp', 'KSpread spreadsheet', ['ksp'], 0.9, 'KoStore ZIP (mimetype application/x-kspread) or maindoc.xml; <spreadsheet> XML', isKspread),
  rule('numbers', 'Apple Numbers spreadsheet', ['numbers'], 0.95, "ZIP with Index/*.iwa, an inner Index.zip or index.xml[.gz]; <ls:document> XML", isNumbers),
  rule('sqlite', 'SQLite database', ['sqlite', 'sqlite3', 'db', 'db3'], 1, '"SQLite format 3\\0" at offset 0', isSqlite),
  rule('gnumeric', 'Gnumeric spreadsheet', ['gnumeric', 'gnm'], 0.95, 'gzip with a .gnumeric/.gnm name, or XML in the Gnumeric namespace', isGnumeric),
  rule('gz', 'gzip-compressed file', ['gz', 'gzip'], 0.5, 'gzip magic 1F 8B 08 (decompressed and identified again)', isGzipContainer),
  rule('uos', 'Uniform Office Format spreadsheet', ['uos', 'uof'], 0.95, '<uof:UOF> root with the spreadsheet MIME type', isUof),
  rule('mht', 'Single File Web Page (MHTML)', ['mht', 'mhtml'], 0.85, 'MIME-Version header with multipart/related', isMhtml),
  rule('dif', 'Data Interchange Format', ['dif'], 0.9, 'TABLE / 0,n header followed by VECTORS, TUPLES or DATA', isDif),
  rule('ndjson', 'JSON Lines', ['ndjson', 'jsonl'], 0.7, 'several lines that are each a JSON object or array', isNdjson),
  rule('json', 'JSON data', ['json'], 0.6, 'top-level JSON array or object', isJsonData),
  rule('prn', 'fixed-width text', ['prn'], 0.3, 'text whose blank columns line up in every line', isFixedWidth),
];

function textOf(u8: Uint8Array): string {
  const n = Math.min(u8.length, 64 * 1024);
  const s = u8.subarray(0, n);
  try {
    if (s[0] === 0xff && s[1] === 0xfe) return new TextDecoder('utf-16le').decode(s.subarray(2));
    if (s[0] === 0xfe && s[1] === 0xff) return new TextDecoder('utf-16be').decode(s.subarray(2));
    if (s[0] === 0xef && s[1] === 0xbb && s[2] === 0xbf) return new TextDecoder('utf-8').decode(s.subarray(3));
    return new TextDecoder('utf-8', { fatal: false }).decode(s);
  } catch { return latin1(s); }
}

/** Builds the detector input; zipNames lists a ZIP's entries (the caller's helper is reused). */
export function sniffInput(u8: Uint8Array, name: string, zipNames: (u8: Uint8Array) => string[]): SniffInput {
  const lower = name.toLowerCase().replace(/^.*[\\/]/, '');
  const ext = lower.includes('.') ? lower.slice(lower.lastIndexOf('.') + 1) : '';
  const isZip = u8.length >= 4 && u8[0] === 0x50 && u8[1] === 0x4b && (u8[2] === 3 || u8[2] === 5 || u8[2] === 7);
  let text: string | null = null, zip: string[] | null | undefined;
  return {
    bytes: u8, name, ext,
    text: () => (text ??= textOf(u8)),
    zip: () => (zip === undefined ? (zip = isZip ? zipNames(u8).map((n) => n.toLowerCase()) : null) : zip),
  };
}

/**
 * Runs the B rules and returns the format id of the first match. Binary and package rules
 * always apply; the text rules do not take a file whose extension names a text format they
 * would misread (a .csv stays CSV).
 */
export function sniffB(u8: Uint8Array, name: string, zipNames: (u8: Uint8Array) => string[]): string | null {
  const x = sniffInput(u8, name, zipNames);
  for (const r of SNIFF_RULES_B) {
    if (r.format === 'prn' && ['csv', 'tsv', 'tab', 'slk', 'sylk', 'dif', 'html', 'htm', 'xml', 'json'].includes(x.ext)) continue;
    if ((r.format === 'json' || r.format === 'ndjson') && ['csv', 'tsv', 'tab', 'txt'].includes(x.ext) && !x.text().trimStart().startsWith('[')) continue;
    try { if (r.detect(x) != null) return r.format; } catch { /* a detector must never break opening */ }
  }
  return null;
}
