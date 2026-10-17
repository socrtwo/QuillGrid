/**
 * SpreadsheetML (.xlsx / .xlsm / .xltx) reader.
 *
 * Package parts are read with JSZip. Small parts (workbook, styles, theme, drawings, charts,
 * tables, comments) go through DOMParser with namespace-agnostic helpers, so files that use
 * prefixed elements (x:row, x:c) read the same as Excel's own. The <sheetData> section, which
 * can hold millions of cells, is scanned with a regex/indexOf cell scanner instead of a DOM.
 */
import JSZip from 'jszip';
import { Workbook, Sheet, newId } from '../model/workbook';
import {
  CellError, errorFromCode,
  type Cell, type CellStyle, type Border, type CFRule, type CFType, type DataValidation, type FilterCriteria, type TableDef,
  type ChartDef, type ChartSeries, type ChartType, type DrawingObject, type Sparkline, type CellComment, type DefinedName,
} from '../model/types';
import { cellKey, parseRange, colIndex, splitSheetRef, MAX_COLS, MAX_ROWS, type Range } from '../model/address';
import { BUILTIN_FORMATS } from '../model/numfmt';
import { partsToSerial } from '../model/dates';
import { shiftFormula } from '../calc/parser';
import {
  parseXml, kids, kid, path, desc, attr, relId, numAttr, boolAttr, unescXml, stripXlfn, bytesToBase64, mimeForExt, decodeText,
} from './index';

// --------------------------------------------------------------------------------------------
// Package access

export interface Rel { id: string; type: string; target: string; external: boolean }

export class Pkg {
  files = new Map<string, JSZip.JSZipObject>();
  constructor(public zip: JSZip) {
    zip.forEach((p, f) => { if (!f.dir) this.files.set(normPath(p), f); });
  }
  has(p: string): boolean { return this.files.has(normPath(p)); }
  async bytes(p: string): Promise<Uint8Array | null> {
    const f = this.files.get(normPath(p));
    return f ? f.async('uint8array') : null;
  }
  async text(p: string): Promise<string | null> {
    const b = await this.bytes(p);
    return b ? decodeText(b) : null;
  }
  async xml(p: string): Promise<Document | null> {
    const t = await this.text(p);
    return t == null ? null : parseXml(t, p);
  }
  async rels(p: string): Promise<Rel[]> {
    const i = p.lastIndexOf('/');
    const relPath = (i >= 0 ? p.slice(0, i + 1) : '') + '_rels/' + p.slice(i + 1) + '.rels';
    let doc: Document | null = null;
    try { doc = await this.xml(relPath); } catch { return []; }
    if (!doc) return [];
    return desc(doc, 'Relationship').map((r) => {
      const external = attr(r, 'TargetMode') === 'External';
      const tgt = attr(r, 'Target') || '';
      return { id: attr(r, 'Id') || '', type: attr(r, 'Type') || '', target: external ? tgt : resolvePath(p, tgt), external };
    });
  }
}

function normPath(p: string): string { return p.replace(/\\/g, '/').replace(/^\/+/, '').toLowerCase(); }

export function resolvePath(base: string, target: string): string {
  let t = target.replace(/\\/g, '/');
  try { t = decodeURIComponent(t); } catch { /* keep as is */ }
  if (t.startsWith('/')) return t.slice(1);
  const parts = base.split('/');
  parts.pop();
  for (const seg of t.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg !== '') parts.push(seg);
  }
  return parts.join('/');
}

const relOf = (rels: Rel[], type: string) => rels.find((r) => r.type.endsWith('/' + type));
const relById = (rels: Rel[], id: string | null) => (id ? rels.find((r) => r.id === id) : undefined);

// --------------------------------------------------------------------------------------------
// Colors

const INDEXED = [
  '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
  '#000000', '#ffffff', '#ff0000', '#00ff00', '#0000ff', '#ffff00', '#ff00ff', '#00ffff',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#c0c0c0', '#808080',
  '#9999ff', '#993366', '#ffffcc', '#ccffff', '#660066', '#ff8080', '#0066cc', '#ccccff',
  '#000080', '#ff00ff', '#ffff00', '#00ffff', '#800080', '#800000', '#008080', '#0000ff',
  '#00ccff', '#ccffff', '#ccffcc', '#ffff99', '#99ccff', '#ff99cc', '#cc99ff', '#ffcc99',
  '#3366ff', '#33cccc', '#99cc00', '#ffcc00', '#ff9900', '#ff6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
];
export { INDEXED as INDEXED_COLORS };

function rgbToHsl(hex: string): [number, number, number] {
  const n = parseInt(hex.slice(1), 16);
  const r = ((n >> 16) & 255) / 255, g = ((n >> 8) & 255) / 255, b = (n & 255) / 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
  const l = (mx + mn) / 2;
  if (mx === mn) return [0, 0, l];
  const d = mx - mn;
  const s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
  let h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
  h /= 6;
  return [h, s, l];
}

function hslToRgb(h: number, s: number, l: number): string {
  const f = (p: number, q: number, t: number) => {
    if (t < 0) t += 1; if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let r: number, g: number, b: number;
  if (s === 0) r = g = b = l;
  else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    r = f(p, q, h + 1 / 3); g = f(p, q, h); b = f(p, q, h - 1 / 3);
  }
  return '#' + [r, g, b].map((x) => Math.round(Math.max(0, Math.min(1, x)) * 255).toString(16).padStart(2, '0')).join('');
}

export function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [h, s, l] = rgbToHsl(hex);
  const nl = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  return hslToRgb(h, s, nl);
}

interface ColorCtx { theme: string[]; indexed: string[] }

/** Resolves an OOXML CT_Color given an attribute accessor. */
function resolveColor(a: (k: string) => string | null, ctx: ColorCtx): string | undefined {
  if (a('auto') === '1' || a('auto') === 'true') return undefined;
  let hex: string | undefined;
  const rgb = a('rgb');
  const theme = a('theme'), idx = a('indexed');
  if (rgb && /^[0-9a-f]{6,8}$/i.test(rgb)) hex = '#' + rgb.slice(-6).toLowerCase();
  else if (theme != null) {
    const t = +theme;
    hex = ctx.theme[[1, 0, 3, 2][t] ?? t];
  } else if (idx != null) {
    const i = +idx;
    if (i === 64) return undefined;
    hex = i === 65 ? '#ffffff' : ctx.indexed[i];
  }
  const tint = a('tint');
  if (hex && tint) hex = applyTint(hex, +tint);
  return hex;
}

const elColor = (el: Element | null, ctx: ColorCtx) => (el ? resolveColor((k) => attr(el, k), ctx) : undefined);

// --------------------------------------------------------------------------------------------
// Fonts and rich text

interface Prop { n: string; a: (k: string) => string | null }
type FontProps = Partial<Pick<CellStyle, 'font' | 'size' | 'bold' | 'italic' | 'underline' | 'strike' | 'vertAlign' | 'color'>>;

function fontFromProps(props: Prop[], ctx: ColorCtx): FontProps {
  const f: FontProps = {};
  const on = (p: Prop) => { const v = p.a('val'); return v == null || !(v === '0' || v === 'false'); };
  for (const p of props) {
    switch (p.n) {
      case 'b': if (on(p)) f.bold = true; break;
      case 'i': if (on(p)) f.italic = true; break;
      case 'strike': if (on(p)) f.strike = true; break;
      case 'u': { const v = p.a('val') || 'single'; if (v !== 'none') f.underline = v as CellStyle['underline']; break; }
      case 'vertAlign': { const v = p.a('val'); if (v === 'superscript' || v === 'subscript') f.vertAlign = v; break; }
      case 'sz': { const v = Number(p.a('val')); if (v > 0) f.size = v; break; }
      case 'color': { const c = resolveColor(p.a, ctx); if (c) f.color = c; break; }
      case 'name': case 'rFont': { const v = p.a('val'); if (v) f.font = v; break; }
    }
  }
  return f;
}

const domProps = (el: Element | null): Prop[] => (el ? kids(el).map((c) => ({ n: c.localName, a: (k: string) => attr(c, k) })) : []);

const ATTR_RE = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
function attrMap(s: string): Record<string, string> {
  const o: Record<string, string> = {};
  ATTR_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = ATTR_RE.exec(s))) { const k = m[1].includes(':') ? m[1].slice(m[1].indexOf(':') + 1) : m[1]; o[k] = unescXml(m[2] ?? m[3]); }
  return o;
}

function regexProps(xml: string): Prop[] {
  const out: Prop[] = [];
  const re = /<(?:[\w.-]+:)?([\w]+)\b([^>]*?)\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) { const a = attrMap(m[2]); out.push({ n: m[1], a: (k) => a[k] ?? null }); }
  return out;
}

/** Decodes OOXML _xHHHH_ escapes. */
export function decodeX(s: string): string {
  return s.indexOf('_x') < 0 ? s : s.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h) => String.fromCharCode(parseInt(h, 16)));
}

const T_RE = /<(?:[\w.-]+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:[\w.-]+:)?t>/g;
const R_RE = /<(?:[\w.-]+:)?r>([\s\S]*?)<\/(?:[\w.-]+:)?r>/g;
const RPR_RE = /<(?:[\w.-]+:)?rPr>([\s\S]*?)<\/(?:[\w.-]+:)?rPr>/;
const RPH_RE = /<(?:[\w.-]+:)?rPh\b[\s\S]*?<\/(?:[\w.-]+:)?rPh>/g;

interface SString { t: string; rich?: Cell['rich'] }

/** Parses the body of an <si> or <is> element (plain or rich text). */
function parseStringItem(inner: string, ctx: ColorCtx): SString {
  if (inner.indexOf('rPh') >= 0) inner = inner.replace(RPH_RE, '');
  if (/<(?:[\w.-]+:)?r>/.test(inner)) {
    const rich: NonNullable<Cell['rich']> = [];
    let t = '';
    R_RE.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = R_RE.exec(inner))) {
      const body = m[1];
      let txt = '';
      T_RE.lastIndex = 0;
      let tm: RegExpExecArray | null;
      while ((tm = T_RE.exec(body))) txt += decodeX(unescXml(tm[1]));
      const pm = RPR_RE.exec(body);
      const s = pm ? fontFromProps(regexProps(pm[1]), ctx) : undefined;
      rich.push(s && Object.keys(s).length ? { t: txt, s } : { t: txt });
      t += txt;
    }
    return rich.length > 1 || rich[0]?.s ? { t, rich } : { t };
  }
  let t = '';
  T_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = T_RE.exec(inner))) t += decodeX(unescXml(m[1]));
  return { t };
}

function parseSST(xml: string, ctx: ColorCtx): SString[] {
  const out: SString[] = [];
  const re = /<(?:[\w.-]+:)?si(?:\s[^>]*)?(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?si>)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) out.push(m[1] ? parseStringItem(m[1], ctx) : { t: '' });
  return out;
}

// --------------------------------------------------------------------------------------------
// Styles

const EXTRA_FORMATS: Record<number, string> = {
  5: '"$"#,##0_);\\("$"#,##0\\)', 6: '"$"#,##0_);[Red]\\("$"#,##0\\)', 7: '"$"#,##0.00_);\\("$"#,##0.00\\)', 8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  27: 'm/d/yyyy', 28: 'm/d/yyyy', 29: 'm/d/yyyy', 30: 'm/d/yy', 31: 'yyyy-mm-dd', 32: 'h:mm', 33: 'h:mm:ss', 34: 'h:mm', 35: 'h:mm:ss', 36: 'm/d/yyyy',
  50: 'm/d/yyyy', 51: 'm/d/yyyy', 52: 'yyyy-mm', 53: 'm/d/yyyy', 54: 'm/d/yyyy', 55: 'h:mm', 56: 'h:mm:ss', 57: 'm/d/yyyy', 58: 'm/d/yyyy', 59: '0', 60: '0.00', 61: '#,##0', 62: '#,##0.00',
};

const BORDER_STYLES: Record<string, Border['style']> = {
  thin: 'thin', medium: 'medium', thick: 'thick', dashed: 'dashed', dotted: 'dotted', double: 'double', hair: 'hair',
  mediumDashed: 'mediumDashed', dashDot: 'dashDot', mediumDashDot: 'dashDot', dashDotDot: 'dashDot', mediumDashDotDot: 'mediumDashed', slantDashDot: 'dashDot',
};

interface StyleTables {
  numFmts: Map<number, string>;
  fonts: FontProps[];
  fills: Partial<CellStyle>[];
  borders: Partial<CellStyle>[];
  xfMap: number[];
  dxfs: Partial<CellStyle>[];
}

function parseFill(f: Element, ctx: ColorCtx, dxf = false): Partial<CellStyle> {
  const out: Partial<CellStyle> = {};
  const pf = kid(f, 'patternFill');
  if (pf) {
    const type = attr(pf, 'patternType') || (dxf ? 'solid' : 'none');
    const fg = elColor(kid(pf, 'fgColor'), ctx), bg = elColor(kid(pf, 'bgColor'), ctx);
    if (type === 'none') return out;
    if (type === 'solid') { const c = dxf ? bg ?? fg : fg ?? bg; if (c) out.fill = c; }
    else { if (bg) out.fill = bg; out.pattern = { type, ...(fg ? { color: fg } : {}) }; }
    return out;
  }
  const gf = kid(f, 'gradientFill');
  if (gf) {
    const stops = kids(gf, 'stop').map((s) => elColor(kid(s, 'color'), ctx)).filter((c): c is string => !!c);
    if (stops.length) {
      out.gradient = { from: stops[0], to: stops[stops.length - 1], ...(numAttr(gf, 'degree') ? { angle: numAttr(gf, 'degree')! } : {}) };
      out.fill = stops[0];
    }
  }
  return out;
}

function parseBorder(b: Element, ctx: ColorCtx): Partial<CellStyle> {
  const out: Partial<CellStyle> = {};
  const side = (name: string): Border | undefined => {
    const el = kid(b, name);
    const st = el && attr(el, 'style');
    if (!el || !st || st === 'none') return undefined;
    const color = elColor(kid(el, 'color'), ctx);
    return { style: BORDER_STYLES[st] || 'thin', ...(color ? { color } : {}) };
  };
  const l = side('left') || side('start'), r = side('right') || side('end'), t = side('top'), bt = side('bottom'), d = side('diagonal');
  if (l) out.bl = l; if (r) out.br = r; if (t) out.bt = t; if (bt) out.bb = bt;
  if (d) { if (boolAttr(b, 'diagonalUp')) out.diagUp = d; if (boolAttr(b, 'diagonalDown')) out.diagDown = d; }
  return out;
}

function parseAlignment(al: Element | null, st: CellStyle): void {
  if (!al) return;
  const h = attr(al, 'horizontal');
  if (h && h !== 'general') st.hAlign = h as CellStyle['hAlign'];
  const v = attr(al, 'vertical');
  if (v && v !== 'bottom') st.vAlign = v === 'center' ? 'middle' : (v as CellStyle['vAlign']);
  if (boolAttr(al, 'wrapText')) st.wrap = true;
  if (boolAttr(al, 'shrinkToFit')) st.shrink = true;
  const ind = numAttr(al, 'indent'); if (ind) st.indent = ind;
  const rot = numAttr(al, 'textRotation');
  if (rot) st.rotate = rot === 255 ? 255 : rot > 90 ? -(rot - 90) : rot;
  if (attr(al, 'readingOrder') === '2') st.rtl = true;
}

function parseStyles(doc: Document | null, book: Workbook, ctx: ColorCtx): StyleTables {
  const t: StyleTables = { numFmts: new Map(), fonts: [], fills: [], borders: [], xfMap: [0], dxfs: [] };
  if (!doc) return t;
  const root = doc.documentElement;
  const custom = kids(path(root, 'colors', 'indexedColors'), 'rgbColor').map((c) => '#' + (attr(c, 'rgb') || '000000').slice(-6).toLowerCase());
  if (custom.length) ctx.indexed = custom.concat(INDEXED.slice(custom.length));
  for (const nf of kids(kid(root, 'numFmts'), 'numFmt')) {
    const id = numAttr(nf, 'numFmtId'), code = attr(nf, 'formatCode');
    if (id != null && code != null) {
      t.numFmts.set(id, code);
      if (id >= 164 && !book.customFormats.includes(code)) book.customFormats.push(code);
    }
  }
  t.fonts = kids(kid(root, 'fonts'), 'font').map((f) => fontFromProps(domProps(f), ctx));
  t.fills = kids(kid(root, 'fills'), 'fill').map((f, i) => {
    const r = parseFill(f, ctx);
    return i === 1 && r.pattern?.type === 'gray125' ? {} : r;
  });
  t.borders = kids(kid(root, 'borders'), 'border').map((b) => parseBorder(b, ctx));
  const def = t.fonts[0] || {};
  const fmtFor = (id: number | null) => (id == null || id === 0 ? undefined : t.numFmts.get(id) ?? BUILTIN_FORMATS[id] ?? EXTRA_FORMATS[id]);
  const xfStyle = (xf: Element): CellStyle => {
    const st: CellStyle = {};
    const fmt = fmtFor(numAttr(xf, 'numFmtId'));
    if (fmt && fmt !== 'General') st.numFmt = fmt;
    const font = t.fonts[numAttr(xf, 'fontId') ?? 0];
    if (font) {
      if (font.font && font.font !== def.font) st.font = font.font;
      if (font.size && font.size !== def.size) st.size = font.size;
      if (font.bold) st.bold = true; if (font.italic) st.italic = true; if (font.strike) st.strike = true;
      if (font.underline) st.underline = font.underline; if (font.vertAlign) st.vertAlign = font.vertAlign;
      if (font.color && font.color !== def.color) st.color = font.color;
    }
    Object.assign(st, t.fills[numAttr(xf, 'fillId') ?? 0] || {});
    Object.assign(st, t.borders[numAttr(xf, 'borderId') ?? 0] || {});
    parseAlignment(kid(xf, 'alignment'), st);
    const pr = kid(xf, 'protection');
    if (pr) { if (attr(pr, 'locked') === '0' || attr(pr, 'locked') === 'false') st.locked = false; if (boolAttr(pr, 'hidden')) st.hidden = true; }
    return st;
  };
  t.xfMap = kids(kid(root, 'cellXfs'), 'xf').map((xf) => book.styleId(xfStyle(xf)));
  if (!t.xfMap.length) t.xfMap = [0];
  const styleXfs = kids(kid(root, 'cellStyleXfs'), 'xf');
  for (const cs of kids(kid(root, 'cellStyles'), 'cellStyle')) {
    const name = attr(cs, 'name');
    const xf = styleXfs[numAttr(cs, 'xfId') ?? -1];
    if (!name || !xf) continue;
    book.cellStyles.push({ name, style: xfStyle(xf), ...(attr(cs, 'builtinId') != null ? { builtin: true } : {}) });
  }
  t.dxfs = kids(kid(root, 'dxfs'), 'dxf').map((d) => {
    const st: Partial<CellStyle> = {};
    const f = kid(d, 'font');
    if (f) Object.assign(st, fontFromProps(domProps(f), ctx));
    const fill = kid(d, 'fill');
    if (fill) Object.assign(st, parseFill(fill, ctx, true));
    const b = kid(d, 'border');
    if (b) Object.assign(st, parseBorder(b, ctx));
    const nf = kid(d, 'numFmt');
    if (nf && attr(nf, 'formatCode')) st.numFmt = attr(nf, 'formatCode')!;
    parseAlignment(kid(d, 'alignment'), st as CellStyle);
    return st;
  });
  return t;
}

// --------------------------------------------------------------------------------------------
// Theme

export function parseTheme(doc: Document | null, book: Workbook): string[] {
  const colors = [...book.theme.colors];
  if (!doc) return colors;
  const scheme = desc(doc, 'clrScheme')[0];
  const order = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  if (scheme) {
    order.forEach((n, i) => {
      const el = kid(scheme, n);
      const c = el && kids(el)[0];
      if (!c) return;
      const v = c.localName === 'sysClr' ? attr(c, 'lastClr') : attr(c, 'val');
      if (v && /^[0-9a-f]{6}$/i.test(v)) colors[i] = '#' + v.toLowerCase();
    });
    book.theme.name = attr(desc(doc, 'theme')[0], 'name') || attr(scheme, 'name') || book.theme.name;
  }
  const major = attr(path(desc(doc, 'majorFont')[0], 'latin'), 'typeface');
  const minor = attr(path(desc(doc, 'minorFont')[0], 'latin'), 'typeface');
  book.theme.colors = colors;
  if (major) book.theme.headingFont = major;
  if (minor) book.theme.bodyFont = minor;
  return colors;
}

// --------------------------------------------------------------------------------------------
// Workbook

interface Ctx extends ColorCtx {
  pkg: Pkg;
  book: Workbook;
  sst: SString[];
  styles: StyleTables;
  date1904: boolean;
  warnings: string[];
  persons: Map<string, string>;
}

const warnOnce = (ctx: Ctx, msg: string) => { if (!ctx.warnings.includes(msg)) ctx.warnings.push(msg); };

export async function importXlsx(data: ArrayBuffer, _name = '', warnings: string[] = []): Promise<Workbook> {
  const zip = await JSZip.loadAsync(data);
  const pkg = new Pkg(zip);
  const book = new Workbook(false);
  const rootRels = await pkg.rels('');
  let wbPath = relOf(rootRels, 'officeDocument')?.target || 'xl/workbook.xml';
  if (!pkg.has(wbPath)) wbPath = [...pkg.files.keys()].find((k) => /(^|\/)workbook\.xml$/.test(k)) || wbPath;
  const wbDoc = await pkg.xml(wbPath);
  if (!wbDoc) throw new Error('workbook.xml is missing or unreadable');
  const wbRoot = wbDoc.documentElement;
  const wbRels = await pkg.rels(wbPath);

  const ctx: Ctx = { pkg, book, sst: [], styles: null as unknown as StyleTables, date1904: false, warnings, persons: new Map(), theme: [...book.theme.colors], indexed: [...INDEXED] };

  const tryXml = async (p: string | undefined, what: string): Promise<Document | null> => {
    if (!p) return null;
    try { return await pkg.xml(p); } catch (e) { warnOnce(ctx, `The ${what} part is damaged and was skipped.`); return null; }
  };

  ctx.theme = parseTheme(await tryXml(relOf(wbRels, 'theme')?.target, 'theme'), book);
  ctx.styles = parseStyles(await tryXml(relOf(wbRels, 'styles')?.target, 'styles'), book, ctx);
  const sstRel = relOf(wbRels, 'sharedStrings');
  if (sstRel) { const t = await pkg.text(sstRel.target); if (t) ctx.sst = parseSST(t, ctx); }
  const personRel = wbRels.find((r) => /\/person$/.test(r.type));
  if (personRel) {
    const pd = await tryXml(personRel.target, 'persons');
    for (const p of desc(pd, 'person')) ctx.persons.set(attr(p, 'id') || '', attr(p, 'displayName') || '');
  }

  // Workbook settings
  const wbPr = kid(wbRoot, 'workbookPr');
  ctx.date1904 = boolAttr(wbPr, 'date1904');
  book.calc.date1904 = ctx.date1904;
  const calcPr = kid(wbRoot, 'calcPr');
  if (calcPr) {
    const mode = attr(calcPr, 'calcMode');
    book.calc.mode = mode === 'manual' ? 'manual' : mode === 'autoNoTable' ? 'autoNoTables' : 'auto';
    book.calc.iterative = boolAttr(calcPr, 'iterate');
    if (numAttr(calcPr, 'iterateCount') != null) book.calc.maxIterations = numAttr(calcPr, 'iterateCount')!;
    if (numAttr(calcPr, 'iterateDelta') != null) book.calc.maxChange = numAttr(calcPr, 'iterateDelta')!;
    if (attr(calcPr, 'fullPrecision') === '0') book.calc.precisionAsDisplayed = true;
    if (attr(calcPr, 'calcOnSave') === '0') book.calc.recalcBeforeSave = false;
  }
  const prot = kid(wbRoot, 'workbookProtection');
  if (prot && boolAttr(prot, 'lockStructure')) {
    book.structureLocked = true;
    const pw = attr(prot, 'workbookPassword');
    if (pw) book.structurePassword = pw;
  }
  if (desc(wbRoot, 'fileSharing')[0] && boolAttr(desc(wbRoot, 'fileSharing')[0], 'readOnlyRecommended')) book.readOnlyRecommended = true;

  // Sheets
  const sheetEls = kids(kid(wbRoot, 'sheets'), 'sheet');
  const used = new Set<string>();
  for (const se of sheetEls) {
    let name = attr(se, 'name') || `Sheet${book.sheets.length + 1}`;
    while (used.has(name.toLowerCase())) name += '_';
    used.add(name.toLowerCase());
    const sheet = new Sheet(book, name);
    const state = attr(se, 'state');
    if (state === 'hidden' || state === 'veryHidden') sheet.visibility = state;
    book.sheets.push(sheet);
    const rel = relById(wbRels, relId(se));
    if (!rel || !pkg.has(rel.target)) { warnOnce(ctx, `Sheet "${name}" has no data part; it was opened empty.`); continue; }
    try {
      await parseSheet(ctx, sheet, rel.target);
    } catch (e) {
      warnOnce(ctx, `Sheet "${name}" could not be read completely: ${e instanceof Error ? e.message : e}`);
    }
  }
  if (!book.sheets.length) book.sheets.push(new Sheet(book, 'Sheet1'));

  // Defined names
  for (const dn of kids(kid(wbRoot, 'definedNames'), 'definedName')) {
    const name = attr(dn, 'name') || '';
    const ref = stripXlfn((dn.textContent || '').trim().replace(/^=/, ''));
    const lsi = numAttr(dn, 'localSheetId');
    const sheet = lsi != null ? book.sheets[lsi] : undefined;
    if (/^_xlnm\./i.test(name)) {
      const kind = name.slice(6).toLowerCase();
      if (!sheet) continue;
      if (kind === 'print_area') sheet.print.area = stripSheetRefs(ref);
      else if (kind === 'print_titles') {
        for (const part of splitTopLevel(ref)) {
          const r = stripSheetRefs(part);
          if (/^\$?\d+:\$?\d+$/.test(r)) sheet.print.titleRows = r.replace(/\$/g, '');
          else if (/^\$?[A-Z]+:\$?[A-Z]+$/i.test(r)) sheet.print.titleCols = r.replace(/\$/g, '');
        }
      }
      continue;
    }
    if (name.startsWith('_xlfn.')) continue;
    const d: DefinedName = { name, ref, sheet: sheet ? sheet.name : null };
    if (attr(dn, 'comment')) d.comment = attr(dn, 'comment')!;
    if (boolAttr(dn, 'hidden')) d.hidden = true;
    book.names.push(d);
  }

  const bv = path(wbRoot, 'bookViews', 'workbookView');
  const at = numAttr(bv, 'activeTab');
  if (at != null && at < book.sheets.length) book.activeSheet = at;

  // VBA passthrough
  const vba = wbRels.find((r) => /vbaProject$/.test(r.type));
  if (vba) {
    const b = await pkg.bytes(vba.target);
    if (b) { (book as any).vbaProject = b; warnOnce(ctx, 'This workbook contains VBA macros. They are kept when saved as .xlsm but are not run.'); }
  }
  if (wbRels.some((r) => /pivotCacheDefinition$/.test(r.type))) warnOnce(ctx, 'PivotTables were opened as static values.');
  if (wbRels.some((r) => /externalLink$/.test(r.type))) warnOnce(ctx, 'Links to other workbooks show their last saved values.');

  await parseDocProps(pkg, rootRels, book);
  book.clearHistory();
  return book;
}

/** "Sheet1!$A$1:$B$2,Sheet1!$D$1" -> "$A$1:$B$2,$D$1" */
function stripSheetRefs(ref: string): string {
  return splitTopLevel(ref).map((p) => splitSheetRef(p).ref).join(',');
}

function splitTopLevel(s: string): string[] {
  const out: string[] = [];
  let cur = '', q = false, depth = 0;
  for (const ch of s) {
    if (ch === "'") q = !q;
    if (!q && ch === '(') depth++;
    if (!q && ch === ')') depth--;
    if (!q && depth === 0 && ch === ',') { out.push(cur); cur = ''; continue; }
    cur += ch;
  }
  if (cur) out.push(cur);
  return out;
}

export async function parseDocProps(pkg: Pkg, rootRels: Rel[], book: Workbook): Promise<void> {
  const p = book.props;
  try {
    const core = await pkg.xml(relOf(rootRels, 'core-properties')?.target || 'docProps/core.xml');
    if (core) {
      const txt = (n: string) => (desc(core, n)[0]?.textContent || '').trim();
      p.title = txt('title'); p.subject = txt('subject'); p.author = txt('creator'); p.keywords = txt('keywords');
      p.comments = txt('description'); p.lastModifiedBy = txt('lastModifiedBy'); p.category = txt('category'); p.status = txt('contentStatus');
      const c = Date.parse(txt('created')), m = Date.parse(txt('modified'));
      if (Number.isFinite(c)) p.created = c;
      if (Number.isFinite(m)) p.modified = m;
    }
    const app = await pkg.xml(relOf(rootRels, 'extended-properties')?.target || 'docProps/app.xml');
    if (app) {
      p.company = (desc(app, 'Company')[0]?.textContent || '').trim();
      p.manager = (desc(app, 'Manager')[0]?.textContent || '').trim();
    }
    const cust = await pkg.xml(relOf(rootRels, 'custom-properties')?.target || 'docProps/custom.xml');
    if (cust) for (const pr of desc(cust, 'property')) { const n = attr(pr, 'name'); if (n) p.custom[n] = (pr.textContent || '').trim(); }
  } catch { /* properties are optional */ }
}

// --------------------------------------------------------------------------------------------
// Sheet

const A_R = /(?:^|\s)r\s*=\s*(?:"([^"]*)"|'([^']*)')/;
const A_T = /(?:^|\s)t\s*=\s*(?:"([^"]*)"|'([^']*)')/;
const A_S = /(?:^|\s)s\s*=\s*(?:"([^"]*)"|'([^']*)')/;
const A_CM = /(?:^|\s)cm\s*=/;
const getA = (s: string, re: RegExp): string | null => { const m = re.exec(s); return m ? (m[1] ?? m[2]) : null; };
const F_RE = /<(?:[\w.-]+:)?f(\s[^>]*?)?(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?f>)/;
const V_RE = /<(?:[\w.-]+:)?v(?:\s[^>]*)?>([\s\S]*?)<\/(?:[\w.-]+:)?v>/;
const IS_RE = /<(?:[\w.-]+:)?is>([\s\S]*?)<\/(?:[\w.-]+:)?is>/;

/** Column width in characters -> px (Calibri 11, max digit width 7). */
export function widthToPx(w: number): number { return Math.trunc(((256 * w + Math.trunc(128 / 7)) / 256) * 7); }
export const ptToPx = (pt: number) => Math.round((pt * 4) / 3);

function decodeA1(ref: string): { r: number; c: number } | null {
  let i = 0, c = 0;
  if (ref.charCodeAt(0) === 36) i++;
  for (; i < ref.length; i++) {
    const ch = ref.charCodeAt(i);
    if (ch >= 65 && ch <= 90) c = c * 26 + ch - 64;
    else if (ch >= 97 && ch <= 122) c = c * 26 + ch - 96;
    else break;
  }
  if (ref.charCodeAt(i) === 36) i++;
  const r = parseInt(ref.slice(i), 10);
  if (!c || !(r > 0)) return null;
  return { r: r - 1, c: c - 1 };
}

function isoToSerial(s: string, date1904: boolean): number | null {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(s.trim());
  if (m) {
    const d = partsToSerial(+m[1], +m[2], +m[3], date1904);
    return d + ((+(m[4] || 0)) * 3600 + (+(m[5] || 0)) * 60 + (+(m[6] || 0))) / 86400;
  }
  const t = /^T?(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?/.exec(s.trim());
  if (t) return ((+t[1]) * 3600 + (+t[2]) * 60 + (+(t[3] || 0))) / 86400;
  return null;
}

async function parseSheet(ctx: Ctx, sheet: Sheet, partPath: string): Promise<void> {
  const xml = await ctx.pkg.text(partPath);
  if (xml == null) return;
  const rels = await ctx.pkg.rels(partPath);
  // Split off sheetData and scan it; parse the rest as a DOM.
  let rest = xml;
  const sm = /<([A-Za-z_][\w.-]*:)?sheetData\b[^>]*?(\/?)>/.exec(xml);
  let dataXml = '';
  if (sm) {
    if (sm[2] === '/') rest = xml.slice(0, sm.index) + xml.slice(sm.index + sm[0].length);
    else {
      const close = `</${sm[1] || ''}sheetData>`;
      let end = xml.indexOf(close, sm.index);
      if (end < 0) { end = xml.length; warnOnce(ctx, `Sheet "${sheet.name}" was cut short; the rows that remain were opened.`); }
      dataXml = xml.slice(sm.index + sm[0].length, end);
      rest = xml.slice(0, sm.index) + (end < xml.length ? xml.slice(end + close.length) : `</${xmlRootName(xml)}>`);
    }
  }
  let doc: Document | null = null;
  try { doc = parseXml(rest, partPath); } catch { warnOnce(ctx, `Some settings of sheet "${sheet.name}" could not be read.`); }
  const root = doc?.documentElement || null;
  const isChartSheet = root?.localName === 'chartsheet';

  if (root) parseSheetSettings(ctx, sheet, root);
  if (dataXml) scanSheetData(ctx, sheet, dataXml);
  if (!root) return;
  parseSheetParts(ctx, sheet, root, rels);

  // Related parts
  const drawingId = relId(kid(root, 'drawing'));
  const drawingRel = relById(rels, drawingId);
  if (drawingRel) {
    try { await parseDrawing(ctx, sheet, drawingRel.target, isChartSheet); } catch (e) { warnOnce(ctx, `Pictures or charts on sheet "${sheet.name}" could not be read.`); }
  }
  for (const rel of rels) {
    try {
      if (rel.type.endsWith('/comments')) await parseComments(ctx, sheet, rel.target);
      else if (/\/threadedComment$/.test(rel.type)) await parseThreaded(ctx, sheet, rel.target);
      else if (rel.type.endsWith('/table')) await parseTable(ctx, sheet, rel.target);
    } catch { warnOnce(ctx, `A comment or table part on sheet "${sheet.name}" is damaged and was skipped.`); }
  }
  // Hyperlinks
  for (const h of kids(kid(root, 'hyperlinks'), 'hyperlink')) {
    const rg = parseRange(attr(h, 'ref') || '');
    if (!rg) continue;
    const rel = relById(rels, relId(h));
    let link = rel ? rel.target : '';
    const loc = attr(h, 'location');
    if (loc) link = (link ? link : '') + '#' + loc;
    if (!link) continue;
    const tip = attr(h, 'tooltip');
    for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 1000); r++) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 100); c++) {
      const k = cellKey(r, c);
      sheet.cells.set(k, { ...(sheet.cells.get(k) || {}), link, ...(tip ? { linkTip: tip } : {}) });
    }
  }
}

function xmlRootName(xml: string): string {
  const m = /<([A-Za-z_][\w.:-]*)[\s>]/.exec(xml.replace(/<\?[\s\S]*?\?>/g, '').replace(/<!--[\s\S]*?-->/g, ''));
  return m ? m[1] : 'worksheet';
}

function parseSheetSettings(ctx: Ctx, sheet: Sheet, root: Element): void {
  const pr = kid(root, 'sheetPr');
  if (pr) {
    const tc = elColor(kid(pr, 'tabColor'), ctx);
    if (tc) sheet.tabColor = tc;
    if (attr(pr, 'codeName')) sheet.codeName = attr(pr, 'codeName')!;
    const ps = kid(pr, 'pageSetUpPr');
    if (ps && boolAttr(ps, 'fitToPage')) { sheet.print.fitWide = 1; sheet.print.fitTall = 1; }
  }
  const sv = path(root, 'sheetViews', 'sheetView');
  if (sv) {
    if (attr(sv, 'showGridLines') === '0' || attr(sv, 'showGridLines') === 'false') sheet.showGrid = false;
    if (attr(sv, 'showRowColHeaders') === '0') sheet.showHeadings = false;
    if (attr(sv, 'showZeros') === '0') sheet.showZeros = false;
    if (boolAttr(sv, 'showFormulas')) sheet.showFormulas = true;
    if (boolAttr(sv, 'rightToLeft')) sheet.rtl = true;
    const z = numAttr(sv, 'zoomScale');
    if (z && z >= 10 && z <= 400) sheet.zoom = z / 100;
    if (boolAttr(sv, 'tabSelected')) ctx.book.activeSheet = ctx.book.sheets.indexOf(sheet);
    const tl = decodeA1(attr(sv, 'topLeftCell') || '');
    if (tl) sheet.scroll = { r: tl.r, c: tl.c };
    const pane = kid(sv, 'pane');
    if (pane) {
      const st = attr(pane, 'state');
      const xs = numAttr(pane, 'xSplit') || 0, ys = numAttr(pane, 'ySplit') || 0;
      if (st === 'frozen' || st === 'frozenSplit') sheet.freeze = { rows: Math.round(ys), cols: Math.round(xs) };
      else if (xs || ys) sheet.split = { x: Math.round((xs / 20) * 4 / 3), y: Math.round((ys / 20) * 4 / 3) };
    }
    const sel = kids(sv, 'selection').pop();
    if (sel) {
      const ac = decodeA1(attr(sel, 'activeCell') || '');
      const ranges = (attr(sel, 'sqref') || '').split(/\s+/).map((x) => parseRange(x)).filter((x): x is Range => !!x);
      if (ac) sheet.sel = { r: ac.r, c: ac.c, ranges: ranges.length ? ranges : [{ r1: ac.r, c1: ac.c, r2: ac.r, c2: ac.c }] };
    }
  }
  const fp = kid(root, 'sheetFormatPr');
  if (fp) {
    const dh = numAttr(fp, 'defaultRowHeight');
    if (dh) sheet.defaultRowH = ptToPx(dh);
    const dw = numAttr(fp, 'defaultColWidth');
    const bw = numAttr(fp, 'baseColWidth');
    if (dw) sheet.defaultColW = widthToPx(dw);
    else if (bw != null && bw !== 8) sheet.defaultColW = widthToPx(bw) + 5;
  }
  for (const col of kids(kid(root, 'cols'), 'col')) {
    const min = (numAttr(col, 'min') || 1) - 1, max = Math.min((numAttr(col, 'max') || min + 1) - 1, MAX_COLS - 1);
    const w = numAttr(col, 'width');
    const hidden = boolAttr(col, 'hidden');
    const lvl = numAttr(col, 'outlineLevel');
    const style = numAttr(col, 'style');
    const sid = style ? ctx.styles.xfMap[style] || 0 : 0;
    const px = w != null && !(hidden && w === 0) ? widthToPx(w) : null;
    for (let c = min; c <= max; c++) {
      if (px != null && (px !== sheet.defaultColW || boolAttr(col, 'customWidth'))) sheet.colW.set(c, px);
      if (hidden) sheet.hiddenCols.add(c);
      if (lvl) sheet.colLevel.set(c, lvl);
      if (sid) sheet.colStyle.set(c, sid);
    }
  }
}

function scanSheetData(ctx: Ctx, sheet: Sheet, xml: string): void {
  const { sst, styles, date1904 } = ctx;
  const xfMap = styles.xfMap;
  const cells = sheet.cells;
  const defH = sheet.defaultRowH;
  const re = /<([A-Za-z_][\w.-]*:)?(row|c)(?=[\s/>])([^>]*?)(\/?)>/g;
  let row = -1, col = -1;
  const shared = new Map<string, { f: string; r: number; c: number }>();
  let arrays: { rg: Range; key: number }[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml))) {
    const pre = m[1] || '', at = m[3];
    if (m[2] === 'row') {
      const rA = getA(at, A_R);
      row = rA ? parseInt(rA, 10) - 1 : row + 1;
      col = -1;
      if (row < 0 || row >= MAX_ROWS) { row = Math.max(0, Math.min(row, MAX_ROWS - 1)); }
      if (at.length > 8) {
        const a = attrMap(at);
        if (a.ht != null) {
          const px = ptToPx(+a.ht);
          const ch = a.customHeight;
          if (ch === '1' || ch === 'true' || (ch == null && px !== defH)) { if (px > 0) sheet.rowH.set(row, px); }
        }
        if (a.hidden === '1' || a.hidden === 'true') sheet.hiddenRows.add(row);
        if (a.outlineLevel && +a.outlineLevel > 0) sheet.rowLevel.set(row, +a.outlineLevel);
        if (a.s && (a.customFormat === '1' || a.customFormat === 'true')) { const sid = xfMap[+a.s] || 0; if (sid) sheet.rowStyle.set(row, sid); }
      }
      continue;
    }
    // <c>
    let body = '';
    if (m[4] !== '/') {
      const close = `</${pre}c>`;
      const e = xml.indexOf(close, re.lastIndex);
      if (e < 0) break;
      body = xml.slice(re.lastIndex, e);
      re.lastIndex = e + close.length;
    }
    const ref = getA(at, A_R);
    const a1 = ref ? decodeA1(ref) : null;
    if (a1) { if (a1.r !== row && row >= 0 && a1.r > row) row = a1.r; else if (row < 0) row = a1.r; col = a1.c; }
    else { col++; if (row < 0) row = 0; }
    if (col >= MAX_COLS) continue;
    const r = a1 ? a1.r : row, c = col;
    const t = getA(at, A_T);
    const sAttr = getA(at, A_S);
    const cell: Cell = {};
    if (sAttr) { const sid = xfMap[+sAttr] || 0; if (sid) cell.s = sid; }
    if (body) {
      // value
      let v: Cell['v'] = null;
      if (t === 'inlineStr') {
        const im = IS_RE.exec(body);
        if (im) { const ss = parseStringItem(im[1], ctx); v = ss.t; if (ss.rich) cell.rich = ss.rich; }
      } else {
        const vm = V_RE.exec(body);
        if (vm) {
          const raw = vm[1];
          switch (t) {
            case 's': { const ss = sst[+raw]; if (ss) { v = ss.t; if (ss.rich) cell.rich = ss.rich; } else v = ''; break; }
            case 'str': v = decodeX(unescXml(raw)); break;
            case 'b': v = raw.trim() === '1' || raw.trim().toLowerCase() === 'true'; break;
            case 'e': v = errorFromCode(unescXml(raw).trim()) || new CellError('#VALUE!'); break;
            case 'd': v = isoToSerial(raw, date1904); break;
            default: { const n = Number(raw); v = raw.trim() === '' ? null : Number.isFinite(n) ? n : decodeX(unescXml(raw)); }
          }
        }
      }
      if (v != null) cell.v = v;
      // formula
      if (body.indexOf('f') >= 0) {
        const fm = F_RE.exec(body);
        if (fm) {
          const fa = fm[1] ? attrMap(fm[1]) : {};
          const text = fm[2] != null ? stripXlfn(unescXml(fm[2]).replace(/^=/, '')) : '';
          const ft = fa.t || 'normal';
          if (ft === 'shared' && fa.si != null) {
            if (text) { shared.set(fa.si, { f: text, r, c }); cell.f = text; }
            else { const master = shared.get(fa.si); if (master) cell.f = shiftFormula(master.f, r - master.r, c - master.c); }
          } else if (ft === 'array') {
            if (text) cell.f = text;
            const rg = fa.ref ? parseRange(fa.ref) : null;
            const dynamic = A_CM.test(at);
            if (rg && (rg.r2 > rg.r1 || rg.c2 > rg.c1)) {
              arrays.push({ rg, key: cellKey(r, c) });
              if (!dynamic) cell.arr = rg;
            } else if (rg && !dynamic) cell.arr = rg;
          } else if (ft === 'dataTable') {
            // What-if data tables keep their values.
          } else if (text) cell.f = text;
        }
      }
    }
    if (arrays.length && cell.f == null) {
      for (const a of arrays) {
        const g = a.rg;
        if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) { cell.sp = a.key; break; }
      }
      if (arrays.length > 16) arrays = arrays.filter((a) => a.rg.r2 >= r);
    }
    if (cell.v == null && cell.f == null && !cell.s && cell.sp == null) continue;
    cells.set(cellKey(r, c), cell);
  }
}

function parseSqref(s: string | null): Range[] {
  return (s || '').trim().split(/\s+/).map((x) => parseRange(x)).filter((x): x is Range => !!x);
}

const txt = (el: Element | null) => (el ? el.textContent || '' : '');

function parseSheetParts(ctx: Ctx, sheet: Sheet, root: Element, _rels: Rel[]): void {
  // Merges
  for (const mc of kids(kid(root, 'mergeCells'), 'mergeCell')) {
    const rg = parseRange(attr(mc, 'ref') || '');
    if (rg && (rg.r2 > rg.r1 || rg.c2 > rg.c1)) sheet.merges.push(rg);
  }
  // Protection
  const sp = kid(root, 'sheetProtection');
  if (sp && boolAttr(sp, 'sheet')) {
    const allow: string[] = [];
    if (attr(sp, 'selectLockedCells') !== '1') allow.push('selectLocked');
    if (attr(sp, 'selectUnlockedCells') !== '1') allow.push('selectUnlocked');
    for (const k of PROTECT_ACTIONS) if (attr(sp, k) === '0' || attr(sp, k) === 'false') allow.push(k);
    sheet.protection = { on: true, allow };
    const pw = attr(sp, 'password');
    const hv = attr(sp, 'hashValue');
    if (hv) sheet.protection.passwordHash = `${attr(sp, 'algorithmName') || 'SHA-512'}:${attr(sp, 'saltValue') || ''}:${attr(sp, 'spinCount') || '100000'}:${hv}`;
    else if (pw) sheet.protection.passwordHash = pw;
  }
  // Auto filter
  const af = kid(root, 'autoFilter');
  if (af) { const f = parseAutoFilter(af); if (f) sheet.filter = f; }
  // Conditional formatting
  let prio = 0;
  for (const cf of kids(root, 'conditionalFormatting')) {
    const ranges = parseSqref(attr(cf, 'sqref'));
    if (!ranges.length) continue;
    for (const rule of kids(cf, 'cfRule')) {
      const r = parseCfRule(ctx, rule, ranges);
      const extId = desc(rule, 'id').find((e) => e.parentElement?.localName === 'ext')?.textContent?.trim();
      if (r && extId) (r as { extId?: string }).extId = extId;
      if (r) { r.priority = r.priority ?? ++prio; sheet.cf.push(r); }
    }
  }
  sheet.cf.sort((a, b) => (a.priority ?? 0) - (b.priority ?? 0));
  // Data validation
  for (const dv of kids(kid(root, 'dataValidations'), 'dataValidation')) {
    const v = parseDv(dv, parseSqref(attr(dv, 'sqref')), txt(kid(dv, 'formula1')), txt(kid(dv, 'formula2')));
    if (v) sheet.validations.push(v);
  }
  // Print
  const po = kid(root, 'printOptions');
  if (po) {
    sheet.print.gridlines = boolAttr(po, 'gridLines');
    sheet.print.headings = boolAttr(po, 'headings');
    sheet.print.centerH = boolAttr(po, 'horizontalCentered');
    sheet.print.centerV = boolAttr(po, 'verticalCentered');
  }
  const pm = kid(root, 'pageMargins');
  if (pm) {
    const g = (k: string, d: number) => numAttr(pm, k) ?? d;
    const m = sheet.print.margins;
    sheet.print.margins = { top: g('top', m.top), bottom: g('bottom', m.bottom), left: g('left', m.left), right: g('right', m.right), header: g('header', m.header), footer: g('footer', m.footer) };
  }
  const ps = kid(root, 'pageSetup');
  if (ps) {
    const pz = numAttr(ps, 'paperSize');
    if (pz && PAPER[pz]) sheet.print.paper = PAPER[pz];
    if (attr(ps, 'orientation') === 'landscape') sheet.print.orientation = 'landscape';
    const sc = numAttr(ps, 'scale'); if (sc) sheet.print.scale = sc;
    if (sheet.print.fitWide != null) {
      const fw = numAttr(ps, 'fitToWidth'), ft = numAttr(ps, 'fitToHeight');
      sheet.print.fitWide = fw ?? 1; sheet.print.fitTall = ft ?? 1;
    }
    if (attr(ps, 'pageOrder') === 'overThenDown') sheet.print.order = 'overThenDown';
    if (boolAttr(ps, 'blackAndWhite')) sheet.print.blackAndWhite = true;
    if (boolAttr(ps, 'draft')) sheet.print.draft = true;
    if (boolAttr(ps, 'useFirstPageNumber') && numAttr(ps, 'firstPageNumber') != null) sheet.print.firstPage = numAttr(ps, 'firstPageNumber')!;
  }
  const hf = kid(root, 'headerFooter');
  if (hf) {
    sheet.print.header = splitHeaderFooter(txt(kid(hf, 'oddHeader')));
    sheet.print.footer = splitHeaderFooter(txt(kid(hf, 'oddFooter')));
    if (boolAttr(hf, 'differentOddEven')) {
      sheet.print.differentOddEven = true;
      if (kid(hf, 'evenHeader')) sheet.print.evenHeader = splitHeaderFooter(txt(kid(hf, 'evenHeader')));
      if (kid(hf, 'evenFooter')) sheet.print.evenFooter = splitHeaderFooter(txt(kid(hf, 'evenFooter')));
    }
    if (boolAttr(hf, 'differentFirst')) {
      sheet.print.differentFirst = true;
      if (kid(hf, 'firstHeader')) sheet.print.firstHeader = splitHeaderFooter(txt(kid(hf, 'firstHeader')));
      if (kid(hf, 'firstFooter')) sheet.print.firstFooter = splitHeaderFooter(txt(kid(hf, 'firstFooter')));
    }
    if (attr(hf, 'scaleWithDoc') === '0' || attr(hf, 'scaleWithDoc') === 'false') sheet.print.scaleWithDoc = false;
    if (attr(hf, 'alignWithMargins') === '0' || attr(hf, 'alignWithMargins') === 'false') sheet.print.alignWithMargins = false;
  }
  sheet.rowBreaks = kids(kid(root, 'rowBreaks'), 'brk').map((b) => numAttr(b, 'id') || 0).filter((x) => x > 0);
  sheet.colBreaks = kids(kid(root, 'colBreaks'), 'brk').map((b) => numAttr(b, 'id') || 0).filter((x) => x > 0);
  // Extensions: x14 data validations, sparklines
  for (const ext of kids(kid(root, 'extLst'), 'ext')) {
    // Excel 2010 conditional formats (data bar details, icon sets, rules that refer to other sheets).
    for (const cfx of desc(ext, 'conditionalFormatting')) {
      const ranges = parseSqref(txt(kid(cfx, 'sqref')).trim());
      if (!ranges.length) continue;
      for (const rule of kids(cfx, 'cfRule')) {
        const id = attr(rule, 'id');
        const existing = id ? sheet.cf.find((x) => (x as { extId?: string }).extId === id) : undefined;
        const dbx = kid(rule, 'dataBar');
        if (existing) {
          if (dbx && attr(dbx, 'gradient') === '0') existing.gradientBar = false;
          delete (existing as { extId?: string }).extId;
          continue;
        }
        const r = parseCfRule(ctx, rule, ranges);
        if (!r) continue;
        const fs = kids(rule, 'f').map((f) => stripXlfn((f.textContent || '').trim()));
        if (fs.length && !r.formulas) r.formulas = fs;
        if (dbx) { r.stops = kids(dbx, 'cfvo').map((v) => ({ type: (attr(v, 'type') || 'min').replace(/^auto/, '').toLowerCase() as 'min' })); r.barColor = elColor(kid(dbx, 'fillColor'), ctx) || r.barColor || '#638ec6'; r.gradientBar = attr(dbx, 'gradient') !== '0'; }
        const dxf = kid(rule, 'dxf');
        if (dxf && !r.style) {
          const st: Partial<CellStyle> = {};
          const f = kid(dxf, 'font'); if (f) Object.assign(st, fontFromProps(domProps(f), ctx));
          const fill = kid(dxf, 'fill'); if (fill) Object.assign(st, parseFill(fill, ctx, true));
          r.style = st;
        }
        r.priority = r.priority ?? sheet.cf.length + 1;
        sheet.cf.push(r);
      }
    }
    for (const r of sheet.cf) delete (r as { extId?: string }).extId;
    for (const dv of desc(ext, 'dataValidation')) {
      const f1 = kid(dv, 'formula1'), f2 = kid(dv, 'formula2');
      const v = parseDv(dv, parseSqref(txt(kid(dv, 'sqref'))), f1 ? txt(desc(f1, 'f')[0] || f1) : '', f2 ? txt(desc(f2, 'f')[0] || f2) : '');
      if (v) sheet.validations.push(v);
    }
    for (const g of desc(ext, 'sparklineGroup')) {
      const typ = attr(g, 'type');
      const type: Sparkline['type'] = typ === 'column' ? 'column' : typ === 'stacked' ? 'winloss' : 'line';
      const color = elColor(kid(g, 'colorSeries'), ctx);
      for (const s of desc(g, 'sparkline')) {
        const data = stripXlfn(txt(kid(s, 'f')).trim());
        const cell = decodeA1(txt(kid(s, 'sqref')).trim());
        if (!cell || !data) continue;
        sheet.sparklines.push({
          id: newId('spk'), cell, data, type, ...(color ? { color } : {}),
          ...(boolAttr(g, 'markers') ? { markers: true } : {}), ...(boolAttr(g, 'high') ? { high: true } : {}),
          ...(boolAttr(g, 'low') ? { low: true } : {}), ...(boolAttr(g, 'negative') ? { negative: true } : {}),
        });
      }
    }
  }
}

export const PROTECT_ACTIONS = ['formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows', 'sort', 'autoFilter', 'pivotTables', 'objects', 'scenarios'];

export const PAPER: Record<number, string> = { 1: 'Letter', 2: 'Letter', 3: 'Tabloid', 4: 'Ledger', 5: 'Legal', 6: 'Statement', 7: 'Executive', 8: 'A3', 9: 'A4', 10: 'A4', 11: 'A5', 12: 'B4', 13: 'B5', 14: 'Folio', 20: 'Envelope #10', 27: 'Envelope DL', 28: 'Envelope C5', 34: 'Envelope B5', 37: 'Envelope Monarch', 66: 'A2', 70: 'A6' };

const HF_CODES: [string, string][] = [['&P', '&[Page]'], ['&N', '&[Pages]'], ['&D', '&[Date]'], ['&T', '&[Time]'], ['&Z', '&[Path]'], ['&F', '&[File]'], ['&A', '&[Tab]'], ['&G', '&[Picture]']];

/** Splits an Excel header/footer string into left/center/right sections and maps field codes. */
export function splitHeaderFooter(s: string): { left: string; center: string; right: string } {
  const out = { left: '', center: '', right: '' };
  let cur: keyof typeof out = 'center';
  let i = 0;
  while (i < s.length) {
    if (s[i] === '&' && i + 1 < s.length) {
      const n = s[i + 1];
      if (n === '&') { out[cur] += '&&'; i += 2; continue; }
      if (n === 'L' || n === 'C' || n === 'R') { cur = n === 'L' ? 'left' : n === 'C' ? 'center' : 'right'; i += 2; continue; }
      const code = HF_CODES.find(([c]) => c[1] === n);
      if (code) { out[cur] += code[1]; i += 2; continue; }
    }
    out[cur] += s[i++];
  }
  return out;
}

function parseAutoFilter(af: Element): Sheet['filter'] {
  const range = parseRange(attr(af, 'ref') || '');
  if (!range) return null;
  const columns: Record<number, FilterCriteria> = {};
  for (const fc of kids(af, 'filterColumn')) {
    const id = numAttr(fc, 'colId') ?? 0;
    const crit: FilterCriteria = {};
    const fs = kid(fc, 'filters');
    if (fs) {
      crit.values = kids(fs, 'filter').map((f) => attr(f, 'val') || '');
      const dg = kids(fs, 'dateGroupItem');
      for (const d of dg) crit.values.push([attr(d, 'year'), attr(d, 'month'), attr(d, 'day')].filter(Boolean).join('-'));
      if (boolAttr(fs, 'blank')) crit.blanks = true;
    }
    const cf = kid(fc, 'customFilters');
    if (cf) {
      crit.custom = kids(cf, 'customFilter').map((f) => ({ op: attr(f, 'operator') || 'equal', value: attr(f, 'val') || '' }));
      if (boolAttr(cf, 'and')) crit.and = true;
    }
    const t10 = kid(fc, 'top10');
    if (t10) crit.top10 = { top: attr(t10, 'top') !== '0', percent: boolAttr(t10, 'percent'), n: numAttr(t10, 'val') ?? 10 };
    const dyn = kid(fc, 'dynamicFilter');
    if (dyn) { const t = attr(dyn, 'type'); if (t === 'aboveAverage' || t === 'belowAverage') crit.dynamic = t; }
    columns[id] = crit;
  }
  const f: NonNullable<Sheet['filter']> = { range, columns };
  const sc = desc(af, 'sortCondition')[0];
  if (sc) { const rg = parseRange(attr(sc, 'ref') || ''); if (rg) f.sort = { col: rg.c1 - range.c1, desc: boolAttr(sc, 'descending') }; }
  return f;
}

const CF_TYPES: CFType[] = ['cellIs', 'expression', 'colorScale', 'dataBar', 'iconSet', 'top10', 'aboveAverage', 'duplicateValues', 'uniqueValues', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors', 'timePeriod'];

function cfvoStops(el: Element): NonNullable<CFRule['stops']> {
  return kids(el, 'cfvo').map((v) => {
    let type = attr(v, 'type') || 'min';
    if (type === 'autoMin') type = 'min';
    if (type === 'autoMax') type = 'max';
    const raw = attr(v, 'val') ?? txt(desc(v, 'f')[0] || null);
    const n = raw !== '' && raw != null ? Number(raw) : NaN;
    const stop: NonNullable<CFRule['stops']>[number] = { type: type as 'min' };
    if (raw != null && raw !== '' && type !== 'min' && type !== 'max') stop.value = Number.isFinite(n) ? n : stripXlfn(raw);
    return stop;
  });
}

function parseCfRule(ctx: Ctx, rule: Element, ranges: Range[]): CFRule | null {
  const type = attr(rule, 'type') as CFType;
  if (!CF_TYPES.includes(type)) return null;
  const r: CFRule = { id: newId('cf'), ranges, type };
  const pr = numAttr(rule, 'priority'); if (pr != null) r.priority = pr;
  const op = attr(rule, 'operator');
  if (op && type === 'cellIs') r.operator = (op === 'notEqual' || CELLIS_OPS.includes(op) ? op : 'equal') as CFRule['operator'];
  const fs = kids(rule, 'formula').map((f) => stripXlfn((f.textContent || '').trim()));
  if (fs.length) r.formulas = fs;
  if (attr(rule, 'text') && /Text|With/.test(type)) r.text = attr(rule, 'text')!;
  if (boolAttr(rule, 'stopIfTrue')) r.stopIfTrue = true;
  if (type === 'top10') { r.rank = numAttr(rule, 'rank') ?? 10; if (boolAttr(rule, 'percent')) r.percent = true; if (boolAttr(rule, 'bottom')) r.bottom = true; }
  if (type === 'aboveAverage') {
    r.above = attr(rule, 'aboveAverage') !== '0';
    if (boolAttr(rule, 'equalAverage')) r.equalAverage = true;
    const sd = numAttr(rule, 'stdDev'); if (sd) r.stdDev = sd;
  }
  if (type === 'timePeriod') r.timePeriod = (attr(rule, 'timePeriod') || 'today') as CFRule['timePeriod'];
  const dxf = numAttr(rule, 'dxfId');
  if (dxf != null && ctx.styles.dxfs[dxf]) r.style = { ...ctx.styles.dxfs[dxf] };
  const cs = kid(rule, 'colorScale');
  if (cs) {
    r.stops = cfvoStops(cs);
    r.colors = kids(cs, 'color').map((c) => elColor(c, ctx) || '#000000');
  }
  const db = kid(rule, 'dataBar');
  if (db) {
    r.stops = cfvoStops(db);
    r.barColor = elColor(kid(db, 'color'), ctx) || '#638ec6';
    r.gradientBar = true;
    if (attr(db, 'showValue') === '0') r.showValue = false;
  }
  const is = kid(rule, 'iconSet');
  if (is) {
    r.stops = cfvoStops(is);
    r.iconSet = attr(is, 'iconSet') || '3TrafficLights1';
    if (boolAttr(is, 'reverse')) r.reverseIcons = true;
    if (attr(is, 'showValue') === '0') r.showValue = false;
  }
  return r;
}
const CELLIS_OPS = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'];

function parseDv(dv: Element, ranges: Range[], f1: string, f2: string): DataValidation | null {
  if (!ranges.length) return null;
  const t = attr(dv, 'type') || 'none';
  const v: DataValidation = { id: newId('dv'), ranges, type: (t === 'none' ? 'any' : t) as DataValidation['type'] };
  const op = attr(dv, 'operator');
  if (v.type !== 'any' && v.type !== 'list' && v.type !== 'custom') v.operator = (op || 'between') as DataValidation['operator'];
  if (f1.trim()) v.formula1 = stripXlfn(f1.trim());
  if (f2.trim() && v.type !== 'list' && v.type !== 'custom' && (v.operator === 'between' || v.operator === 'notBetween')) v.formula2 = stripXlfn(f2.trim());
  v.allowBlank = boolAttr(dv, 'allowBlank');
  if (v.type === 'list') v.showDropdown = !boolAttr(dv, 'showDropDown');
  if (boolAttr(dv, 'showInputMessage') || attr(dv, 'prompt')) {
    if (attr(dv, 'promptTitle')) v.inputTitle = attr(dv, 'promptTitle')!;
    if (attr(dv, 'prompt')) v.inputMessage = attr(dv, 'prompt')!;
  }
  const es = attr(dv, 'errorStyle');
  if (es === 'warning' || es === 'information') v.errorStyle = es; else v.errorStyle = 'stop';
  if (attr(dv, 'errorTitle')) v.errorTitle = attr(dv, 'errorTitle')!;
  if (attr(dv, 'error')) v.errorMessage = attr(dv, 'error')!;
  return v;
}

// --------------------------------------------------------------------------------------------
// Comments

async function parseComments(ctx: Ctx, sheet: Sheet, p: string): Promise<void> {
  const doc = await ctx.pkg.xml(p);
  if (!doc) return;
  const authors = kids(desc(doc, 'authors')[0], 'author').map((a) => a.textContent || '');
  for (const c of desc(doc, 'comment')) {
    const a1 = decodeA1(attr(c, 'ref') || '');
    if (!a1) continue;
    const k = cellKey(a1.r, a1.c);
    if (sheet.comments.has(k)) continue; // threaded comment already there
    const author = authors[numAttr(c, 'authorId') ?? 0] || '';
    const textEl = kid(c, 'text');
    let text = textEl ? desc(textEl, 't').filter((t) => t.parentElement?.localName !== 'rPh').map((t) => t.textContent || '').join('') : '';
    if (author && text.startsWith(author + ':')) text = text.slice(author.length + 1).replace(/^\r?\n/, '');
    sheet.comments.set(k, { author, text: decodeX(text), time: 0, note: true });
  }
}

async function parseThreaded(ctx: Ctx, sheet: Sheet, p: string): Promise<void> {
  const doc = await ctx.pkg.xml(p);
  if (!doc) return;
  const byId = new Map<string, CellComment>();
  for (const tc of desc(doc, 'threadedComment')) {
    const a1 = decodeA1(attr(tc, 'ref') || '');
    if (!a1) continue;
    const author = ctx.persons.get(attr(tc, 'personId') || '') || '';
    const time = Date.parse(attr(tc, 'dT') || '') || 0;
    const text = txt(kid(tc, 'text'));
    const parent = attr(tc, 'parentId');
    if (parent && byId.has(parent)) {
      const root = byId.get(parent)!;
      (root.replies ||= []).push({ author, text, time });
      continue;
    }
    const cm: CellComment = { author, text, time };
    if (boolAttr(tc, 'done')) cm.resolved = true;
    byId.set(attr(tc, 'id') || newId('tc'), cm);
    sheet.comments.set(cellKey(a1.r, a1.c), cm);
  }
}

// --------------------------------------------------------------------------------------------
// Tables

async function parseTable(ctx: Ctx, sheet: Sheet, p: string): Promise<void> {
  const doc = await ctx.pkg.xml(p);
  if (!doc) return;
  const t = doc.documentElement;
  const range = parseRange(attr(t, 'ref') || '');
  if (!range) return;
  const header = attr(t, 'headerRowCount') !== '0';
  const totals = (numAttr(t, 'totalsRowCount') || 0) > 0 || boolAttr(t, 'totalsRowShown') && (numAttr(t, 'totalsRowCount') || 0) > 0;
  const si = kid(t, 'tableStyleInfo');
  const columns = kids(kid(t, 'tableColumns'), 'tableColumn').map((c) => {
    const col: TableDef['columns'][number] = { name: decodeX(attr(c, 'name') || '') };
    const fn = attr(c, 'totalsRowFunction');
    if (fn && fn !== 'none') col.totalFn = fn;
    const tf = kid(c, 'totalsRowFormula'); if (tf) col.totalFormula = stripXlfn(txt(tf).trim());
    const cf = kid(c, 'calculatedColumnFormula'); if (cf) col.calcFormula = stripXlfn(txt(cf).trim());
    return col;
  });
  const def: TableDef = {
    id: newId('tbl'), name: attr(t, 'displayName') || attr(t, 'name') || `Table${sheet.tables.length + 1}`, range, columns,
    style: attr(si, 'name') || 'TableStyleMedium2',
    opts: {
      header, totals, banded: si ? boolAttr(si, 'showRowStripes') : true, bandedCols: boolAttr(si, 'showColumnStripes'),
      firstCol: boolAttr(si, 'showFirstColumn'), lastCol: boolAttr(si, 'showLastColumn'), filterButton: !!kid(t, 'autoFilter'),
    },
  };
  sheet.tables.push(def);
}

// --------------------------------------------------------------------------------------------
// Drawings: pictures, shapes, charts

const EMU_PX = 9525;

function anchorPoint(el: Element | null): { r: number; c: number; dx: number; dy: number } | null {
  if (!el) return null;
  const n = (k: string) => Number(txt(kid(el, k)).trim()) || 0;
  return { c: n('col'), dx: n('colOff') / EMU_PX, r: n('row'), dy: n('rowOff') / EMU_PX };
}

function spanPx(sheet: Sheet, from: number, to: number, cols: boolean): number {
  let s = 0;
  const lim = Math.min(to, from + 2000);
  for (let i = from; i < lim; i++) s += cols ? sheet.colWidth(i) : sheet.rowHeight(i);
  return s;
}

function pxToCell(sheet: Sheet, x: number, y: number): { r: number; c: number; dx: number; dy: number } {
  let c = 0, r = 0;
  while (c < MAX_COLS - 1 && x >= sheet.colWidth(c) && c < 5000) { x -= sheet.colWidth(c); c++; }
  while (r < MAX_ROWS - 1 && y >= sheet.rowHeight(r) && r < 100000) { y -= sheet.rowHeight(r); r++; }
  return { r, c, dx: x, dy: y };
}

/** Picks the content element of an anchor, resolving mc:AlternateContent. */
function anchorContent(anchor: Element): Element | null {
  for (const ch of kids(anchor)) {
    if (['pic', 'sp', 'graphicFrame', 'grpSp', 'cxnSp'].includes(ch.localName)) return ch;
    if (ch.localName === 'AlternateContent') {
      const choice = kid(ch, 'Choice');
      const inner = choice && kids(choice).find((x) => ['pic', 'sp', 'graphicFrame', 'grpSp', 'cxnSp'].includes(x.localName));
      if (inner && !(inner.localName === 'graphicFrame' && !desc(inner, 'chart').some((c) => c.namespaceURI?.includes('drawingml/2006/chart')))) return inner;
      const fb = kid(ch, 'Fallback');
      const f = fb && kids(fb).find((x) => ['pic', 'sp', 'graphicFrame', 'grpSp', 'cxnSp'].includes(x.localName));
      if (f) return f;
      if (inner) return inner;
    }
  }
  return null;
}

async function parseDrawing(ctx: Ctx, sheet: Sheet, p: string, chartSheet: boolean): Promise<void> {
  const doc = await ctx.pkg.xml(p);
  if (!doc) return;
  const rels = await ctx.pkg.rels(p);
  for (const anchor of kids(doc.documentElement)) {
    const kind = anchor.localName;
    if (kind !== 'twoCellAnchor' && kind !== 'oneCellAnchor' && kind !== 'absoluteAnchor') continue;
    const content = anchorContent(anchor);
    if (!content) continue;
    let at: ChartDef['at'];
    const xfrmExt = desc(content, 'ext').find((e) => e.parentElement?.localName === 'xfrm');
    const extW = (numAttr(xfrmExt, 'cx') || 0) / EMU_PX, extH = (numAttr(xfrmExt, 'cy') || 0) / EMU_PX;
    if (kind === 'absoluteAnchor') {
      const pos = kid(anchor, 'pos'), ext = kid(anchor, 'ext');
      const x = (numAttr(pos, 'x') || 0) / EMU_PX, y = (numAttr(pos, 'y') || 0) / EMU_PX;
      const pc = pxToCell(sheet, x, y);
      at = { ...pc, w: (numAttr(ext, 'cx') || 0) / EMU_PX || extW || 480, h: (numAttr(ext, 'cy') || 0) / EMU_PX || extH || 288 };
      if (chartSheet) at = { r: 0, c: 0, dx: 0, dy: 0, w: at.w || 800, h: at.h || 500 };
    } else {
      const from = anchorPoint(kid(anchor, 'from'));
      if (!from) continue;
      let w = 0, h = 0;
      if (kind === 'twoCellAnchor') {
        const to = anchorPoint(kid(anchor, 'to'));
        if (to) {
          w = spanPx(sheet, from.c, to.c, true) - from.dx + to.dx;
          h = spanPx(sheet, from.r, to.r, false) - from.dy + to.dy;
        }
      } else {
        const ext = kid(anchor, 'ext');
        w = (numAttr(ext, 'cx') || 0) / EMU_PX; h = (numAttr(ext, 'cy') || 0) / EMU_PX;
      }
      if (!(w > 0)) w = extW || 100;
      if (!(h > 0)) h = extH || 100;
      at = { ...from, w, h };
    }
    at = { r: at.r, c: at.c, dx: Math.round(at.dx), dy: Math.round(at.dy), w: Math.round(at.w), h: Math.round(at.h) };
    const nv = desc(content, 'cNvPr')[0];
    const alt = attr(nv, 'descr') || undefined;
    const xfrm = desc(content, 'xfrm')[0];
    const rot = (numAttr(xfrm, 'rot') || 0) / 60000;
    switch (content.localName) {
      case 'pic': {
        const blip = desc(content, 'blip')[0];
        const rel = relById(rels, relId(blip, 'embed'));
        if (!rel) break;
        const bytes = await ctx.pkg.bytes(rel.target);
        if (!bytes) break;
        const ext = rel.target.slice(rel.target.lastIndexOf('.') + 1);
        const d: DrawingObject = { id: newId('img'), kind: 'image', at, src: `data:${mimeForExt(ext)};base64,${bytesToBase64(bytes)}` };
        if (alt) d.alt = alt;
        if (rot) d.rotate = rot;
        const lock = desc(content, 'picLocks')[0];
        if (lock && boolAttr(lock, 'noChangeAspect')) d.lockAspect = true;
        sheet.drawings.push(d);
        break;
      }
      case 'sp': case 'cxnSp': {
        const isText = boolAttr(desc(content, 'cNvSpPr')[0], 'txBox');
        const geom = desc(content, 'prstGeom')[0];
        const spPr = kid(content, 'spPr');
        const d: DrawingObject = { id: newId('shp'), kind: isText ? 'textbox' : 'shape', at, shape: attr(geom, 'prst') || (content.localName === 'cxnSp' ? 'line' : 'rect') };
        const fillEl = kid(spPr, 'solidFill');
        const fill = fillEl ? drawingColor(fillEl, ctx) : undefined;
        if (fill) d.fill = fill;
        if (kid(spPr, 'noFill')) d.fill = 'none';
        const ln = kid(spPr, 'ln');
        if (ln) {
          const lc = kid(ln, 'solidFill');
          if (lc) d.stroke = drawingColor(lc, ctx);
          if (kid(ln, 'noFill')) d.stroke = 'none';
          const w = numAttr(ln, 'w'); if (w) d.strokeWidth = Math.round((w / EMU_PX) * 100) / 100;
        }
        const tb = kid(content, 'txBody');
        if (tb) {
          d.text = kids(tb, 'p').map((pp) => desc(pp, 't').map((t) => t.textContent || '').join('')).join('\n');
          const rpr = desc(tb, 'rPr')[0] || desc(tb, 'defRPr')[0];
          if (rpr) {
            const font: Partial<CellStyle> = {};
            const sz = numAttr(rpr, 'sz'); if (sz) font.size = sz / 100;
            if (boolAttr(rpr, 'b')) font.bold = true;
            if (boolAttr(rpr, 'i')) font.italic = true;
            const lat = attr(kid(rpr, 'latin'), 'typeface'); if (lat && !lat.startsWith('+')) font.font = lat;
            const sf = kid(rpr, 'solidFill'); if (sf) { const c = drawingColor(sf, ctx); if (c) font.color = c; }
            if (Object.keys(font).length) d.font = font;
          }
        }
        if (rot) d.rotate = rot;
        if (alt) d.alt = alt;
        sheet.drawings.push(d);
        break;
      }
      case 'graphicFrame': {
        const ch = desc(content, 'chart').find((c) => relId(c));
        const rel = relById(rels, relId(ch || null));
        if (!rel) { warnOnce(ctx, 'Some embedded objects (SmartArt, newer chart types) could not be shown.'); break; }
        const cdoc = await ctx.pkg.xml(rel.target).catch(() => null);
        const chart = cdoc ? parseChart(cdoc, ctx) : null;
        if (chart) sheet.charts.push({ id: newId('ch'), at, ...chart } as ChartDef);
        break;
      }
      case 'grpSp':
        warnOnce(ctx, 'Grouped shapes were simplified.');
        for (const pic of desc(content, 'pic')) {
          const blip = desc(pic, 'blip')[0];
          const rel = relById(rels, relId(blip, 'embed'));
          const bytes = rel ? await ctx.pkg.bytes(rel.target) : null;
          if (!rel || !bytes) continue;
          const ext = rel.target.slice(rel.target.lastIndexOf('.') + 1);
          sheet.drawings.push({ id: newId('img'), kind: 'image', at, src: `data:${mimeForExt(ext)};base64,${bytesToBase64(bytes)}` });
        }
        break;
    }
  }
}

/** a:solidFill (srgbClr / schemeClr / sysClr / prstClr) -> hex. */
function drawingColor(fill: Element, ctx: ColorCtx): string | undefined {
  const c = kids(fill)[0];
  if (!c) return undefined;
  let hex: string | undefined;
  const v = attr(c, 'val');
  if (c.localName === 'srgbClr' && v) hex = '#' + v.toLowerCase();
  else if (c.localName === 'sysClr') hex = '#' + (attr(c, 'lastClr') || '000000').toLowerCase();
  else if (c.localName === 'schemeClr' && v) {
    const map: Record<string, number> = { dk1: 0, lt1: 1, dk2: 2, lt2: 3, tx1: 0, bg1: 1, tx2: 2, bg2: 3, accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9, hlink: 10, folHlink: 11 };
    hex = ctx.theme[map[v] ?? 4];
  } else if (c.localName === 'prstClr' && v) hex = PRST[v] || '#000000';
  if (!hex) return undefined;
  const lm = numAttr(kid(c, 'lumMod'), 'val'), lo = numAttr(kid(c, 'lumOff'), 'val');
  if (lm != null || lo != null) {
    const [h, s, l] = rgbToHsl(hex);
    hex = hslToRgb(h, s, Math.max(0, Math.min(1, l * ((lm ?? 100000) / 100000) + (lo ?? 0) / 100000)));
  }
  const sh = numAttr(kid(c, 'shade'), 'val'); if (sh != null) hex = applyTint(hex, -(1 - sh / 100000));
  const ti = numAttr(kid(c, 'tint'), 'val'); if (ti != null) hex = applyTint(hex, 1 - ti / 100000);
  return hex;
}
const PRST: Record<string, string> = { black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', blue: '#0000ff', yellow: '#ffff00', gray: '#808080' };

// --------------------------------------------------------------------------------------------
// Charts

const richText = (el: Element | null) => (el ? kids(el, 'tx').concat(el.localName === 'tx' ? [el] : []).flatMap((tx) => desc(tx, 'p').map((p) => desc(p, 't').map((t) => t.textContent || '').join(''))).join('\n') : '');

function chartTitle(el: Element | null): string | undefined {
  if (!el) return undefined;
  const tx = kid(el, 'tx');
  if (!tx) return undefined;
  const rich = kid(tx, 'rich');
  if (rich) return desc(rich, 'p').map((p) => desc(p, 't').map((t) => t.textContent || '').join('')).join('\n') || undefined;
  const f = desc(tx, 'f')[0];
  if (f) return (desc(tx, 'v')[0]?.textContent || f.textContent || '') || undefined;
  return undefined;
}
void richText;

function refOf(el: Element | null): string | undefined {
  if (!el) return undefined;
  const f = desc(el, 'f')[0];
  return f ? stripXlfn((f.textContent || '').trim().replace(/^=/, '')) || undefined : undefined;
}

function litValues(el: Element | null): string | undefined {
  if (!el) return undefined;
  const pts = desc(el, 'pt').map((p) => txt(kid(p, 'v')));
  return pts.length ? '{' + pts.map((v) => (Number.isFinite(Number(v)) && v !== '' ? v : `"${v.replace(/"/g, '""')}"`)).join(',') + '}' : undefined;
}

function chartTypeOf(el: Element): ChartType | null {
  const n = el.localName;
  const grouping = attr(kid(el, 'grouping'), 'val') || 'standard';
  switch (n) {
    case 'barChart': case 'bar3DChart': {
      const bar = attr(kid(el, 'barDir'), 'val') === 'bar';
      if (grouping === 'stacked') return bar ? 'stackedBar' : 'stackedColumn';
      if (grouping === 'percentStacked') return bar ? 'percentBar' : 'percentColumn';
      return bar ? 'bar' : 'column';
    }
    case 'lineChart': case 'line3DChart': {
      if (grouping === 'stacked' || grouping === 'percentStacked') return 'stackedLine';
      const chartMarker = attr(kid(el, 'marker'), 'val');
      const sers = kids(el, 'ser');
      const anyMarkers = sers.some((s) => { const sym = attr(path(s, 'marker', 'symbol'), 'val'); return sym ? sym !== 'none' : chartMarker === '1' || chartMarker === 'true'; });
      return anyMarkers ? 'lineMarkers' : 'line';
    }
    case 'pieChart': case 'pie3DChart': case 'ofPieChart': return 'pie';
    case 'doughnutChart': return 'doughnut';
    case 'areaChart': case 'area3DChart':
      return grouping === 'stacked' ? 'stackedArea' : grouping === 'percentStacked' ? 'percentArea' : 'area';
    case 'scatterChart': {
      const st = attr(kid(el, 'scatterStyle'), 'val') || 'marker';
      const sers = kids(el, 'ser');
      const noLines = st === 'marker' || sers.every((s) => !!path(s, 'spPr', 'ln', 'noFill'));
      return noLines ? 'scatter' : 'scatterLines';
    }
    case 'radarChart': return attr(kid(el, 'radarStyle'), 'val') === 'filled' ? 'filledRadar' : 'radar';
    case 'bubbleChart': return 'bubble';
    case 'stockChart': return 'stock';
    case 'surfaceChart': case 'surface3DChart': return 'column';
  }
  return null;
}

function parseChart(doc: Document, ctx: Ctx): Partial<ChartDef> | null {
  const space = doc.documentElement;
  const chart = kid(space, 'chart');
  if (!chart) return null;
  const plot = kid(chart, 'plotArea');
  if (!plot) return null;
  const groups = kids(plot).filter((e) => /Chart$/.test(e.localName));
  const series: ChartSeries[] = [];
  let main: ChartType | null = null;
  let dataLabels = false;
  for (const g of groups) {
    const t = chartTypeOf(g);
    if (!t) continue;
    main = main || t;
    if (desc(g, 'dLbls').some((d) => attr(kid(d, 'showVal'), 'val') === '1')) dataLabels = true;
    for (const s of kids(g, 'ser')) {
      const tx = kid(s, 'tx');
      const vals = kid(s, 'val') || kid(s, 'yVal');
      const cats = kid(s, 'cat') || kid(s, 'xVal');
      const ser: ChartSeries = { values: refOf(vals) || litValues(vals) || '' };
      const nameRef = refOf(tx);
      if (nameRef && /^"(?:[^"]|"")*"$/.test(nameRef)) ser.name = nameRef.slice(1, -1).replace(/""/g, '"');
      else if (nameRef) { ser.nameRef = nameRef; const cached = desc(tx!, 'v')[0]?.textContent; if (cached) ser.name = cached; }
      else if (tx) { const v = desc(tx, 'v')[0]?.textContent || ''; if (v) ser.name = v; }
      const cr = refOf(cats) || litValues(cats);
      if (cr) ser.categories = cr;
      const fill = path(s, 'spPr', 'solidFill') || path(s, 'spPr', 'ln', 'solidFill');
      const col = fill ? drawingColor(fill, ctx) : undefined;
      if (col) ser.color = col;
      if (groups.length > 1) ser.type = t;
      series.push(ser);
    }
  }
  if (!main) { warnOnce(ctx, 'Some charts use types that could not be shown.'); return null; }
  const def: Partial<ChartDef> = { type: groups.length > 1 && new Set(series.map((s) => s.type)).size > 1 ? 'combo' : main, series };
  if (def.type !== 'combo') for (const s of series) delete s.type;
  const title = chartTitle(kid(chart, 'title'));
  if (title) def.title = title;
  else if (kid(chart, 'title') && attr(kid(chart, 'autoTitleDeleted'), 'val') !== '1' && series.length === 1 && series[0].name) def.title = series[0].name;
  const legend = kid(chart, 'legend');
  if (!legend) def.legend = 'none';
  else {
    const pos = attr(kid(legend, 'legendPos'), 'val') || 'r';
    def.legend = ({ r: 'right', b: 'bottom', t: 'top', l: 'left', tr: 'right' } as Record<string, ChartDef['legend']>)[pos] || 'right';
  }
  if (dataLabels) def.dataLabels = true;
  const valAxes = kids(plot, 'valAx');
  const catAx = kid(plot, 'catAx') || kid(plot, 'dateAx');
  let yAx: Element | null = valAxes[0] || null, xAx: Element | null = catAx;
  if (!catAx && valAxes.length > 1) {
    xAx = valAxes.find((a) => ['b', 't'].includes(attr(kid(a, 'axPos'), 'val') || '')) || valAxes[0];
    yAx = valAxes.find((a) => a !== xAx) || null;
  }
  if (yAx) {
    def.gridlines = !!kid(yAx, 'majorGridlines');
    const yt = chartTitle(kid(yAx, 'title')); if (yt) def.yTitle = yt;
    const sc = kid(yAx, 'scaling');
    const mn = numAttr(kid(sc, 'min'), 'val'), mx = numAttr(kid(sc, 'max'), 'val');
    if (mn != null) def.yMin = mn;
    if (mx != null) def.yMax = mx;
    if (kid(sc, 'logBase')) def.logScale = true;
  } else def.gridlines = false;
  if (xAx) { const xt = chartTitle(kid(xAx, 'title')); if (xt) def.xTitle = xt; }
  const style = numAttr(kid(space, 'style'), 'val');
  if (style) def.style = style;
  const tl = desc(plot, 'trendline')[0];
  if (tl) {
    const tt = attr(kid(tl, 'trendlineType'), 'val');
    def.trendline = tt === 'exp' ? 'exponential' : tt === 'poly' ? 'polynomial' : tt === 'movingAvg' ? 'movingAverage' : 'linear';
  }
  return def;
}

export { colIndex };

// --------------------------------------------------------------------------------------------
// Shared with the Excel binary workbook reader (xlsb.ts), whose drawings, charts, threaded
// comments and theme are XML parts like in .xlsx.

/** Reads a sheet's drawing part (pictures, shapes, charts) into the sheet. */
export async function readSheetDrawing(pkg: Pkg, sheet: Sheet, drawingPath: string, theme: string[], warnings: string[], persons = new Map<string, string>()): Promise<void> {
  const ctx: Ctx = { pkg, book: sheet.book, sst: [], styles: { numFmts: new Map(), fonts: [], fills: [], borders: [], xfMap: [0], dxfs: [] }, date1904: sheet.book.calc.date1904, warnings, persons, theme, indexed: [...INDEXED] };
  await parseDrawing(ctx, sheet, drawingPath, false);
}

/** Reads a threaded comments part into the sheet (persons: person id -> display name). */
export async function readThreadedComments(pkg: Pkg, sheet: Sheet, partPath: string, persons: Map<string, string>, warnings: string[]): Promise<void> {
  const ctx: Ctx = { pkg, book: sheet.book, sst: [], styles: { numFmts: new Map(), fonts: [], fills: [], borders: [], xfMap: [0], dxfs: [] }, date1904: false, warnings, persons, theme: sheet.book.theme.colors, indexed: [...INDEXED] };
  await parseThreaded(ctx, sheet, partPath);
}
