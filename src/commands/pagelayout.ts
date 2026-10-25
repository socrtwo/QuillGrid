/**
 * Page Layout tab: themes (colors, fonts, effects), page setup (margins, orientation, size, print
 * area, breaks, background, print titles), Scale to Fit, sheet print options, and the Page Setup
 * dialog (Page, Margins, Header/Footer, Sheet).
 */
import { def, type MenuEntry } from './registry';
import type { App } from '../app';
import type { Sheet, ThemeDef } from '../model/workbook';
import { DEFAULT_THEME } from '../model/workbook';
import type { PrintSettings } from '../model/types';
import { absRangeName, rangeName } from '../model/address';
import { resolveRangeText } from '../model/structure';
import { selRanges, activeCell, clipToUsed, isWholeCols, isWholeRows } from './ops';
import { openDialog, toast, choiceDialog, type Field, type DialogControl } from '../ui/dialog';
import { h, pickFile, readFileAsDataURL } from '../ui/dom';
import { FONT_LIST } from '../fonts';
import { PAPER_SIZES, MARGIN_PRESETS, computePageBreaks, layoutPages, parseTitleSpan, splitAreaText, expandHeaderFooter, type MarginPreset } from '../print/pages';

// --------------------------------------------------------------------------------------------
// Themes

/** Built-in themes: 12 colors (dark 1, light 1, dark 2, light 2, accents 1–6, hyperlink, followed hyperlink) and fonts. */
export const THEMES: ThemeDef[] = [
  DEFAULT_THEME,
  { name: 'Harbor', colors: ['#000000', '#ffffff', '#17324d', '#e6eef5', '#1f6f8b', '#99a8b2', '#e6a400', '#3c8d5a', '#b5523b', '#6b5b95', '#1f6f8b', '#6b5b95'], headingFont: 'Open Sans', bodyFont: 'Open Sans' },
  { name: 'Meadow', colors: ['#000000', '#ffffff', '#2e3b1f', '#eef3e2', '#5b8c2a', '#a3b81c', '#e0a526', '#2f8f83', '#7a5c3e', '#c75d2c', '#2f7d32', '#7a5c3e'], headingFont: 'Lato', bodyFont: 'Lato' },
  { name: 'Ember', colors: ['#000000', '#ffffff', '#3a1c14', '#f7ece6', '#c0392b', '#e67e22', '#f1c40f', '#8e44ad', '#2c3e50', '#16a085', '#c0392b', '#8e44ad'], headingFont: 'Merriweather', bodyFont: 'Open Sans' },
  { name: 'Slate', colors: ['#000000', '#ffffff', '#2f3640', '#e8eaed', '#4b6584', '#778ca3', '#a5b1c2', '#3867d6', '#20bf6b', '#fa8231', '#3867d6', '#8854d0'], headingFont: 'Arimo', bodyFont: 'Arimo' },
  { name: 'Orchard', colors: ['#000000', '#ffffff', '#3d2b1f', '#f6efe6', '#b5452f', '#d98c1f', '#6a8d3a', '#8c5b9e', '#3f7fa6', '#c9a227', '#3f7fa6', '#8c5b9e'], headingFont: 'Libre Baskerville', bodyFont: 'Lato' },
  { name: 'Glacier', colors: ['#000000', '#ffffff', '#163a5f', '#eaf4fb', '#2d9cdb', '#56ccf2', '#6fcf97', '#9b51e0', '#f2994a', '#4f6d7a', '#2d9cdb', '#9b51e0'], headingFont: 'Carlito', bodyFont: 'Carlito' },
  { name: 'Dusk', colors: ['#000000', '#ffffff', '#2b2140', '#efeaf6', '#6c4ab6', '#b24d8f', '#e5735c', '#f2b134', '#3b9c9c', '#5d6d7e', '#6c4ab6', '#b24d8f'], headingFont: 'Gelasio', bodyFont: 'Open Sans' },
  { name: 'Terracotta', colors: ['#000000', '#ffffff', '#4a2c20', '#f8eee8', '#c8643b', '#d9a05b', '#7d9c6b', '#5d7fa3', '#a3586b', '#857f72', '#c8643b', '#a3586b'], headingFont: 'Caladea', bodyFont: 'Carlito' },
  { name: 'Ledger', colors: ['#000000', '#ffffff', '#1e2a24', '#eef1ec', '#2e6b4f', '#8c9a3a', '#c2862b', '#5a6f8f', '#a14a3b', '#6f6a5e', '#2e6b4f', '#5a6f8f'], headingFont: 'Tinos', bodyFont: 'Arimo' },
  { name: 'Sapphire', colors: ['#000000', '#ffffff', '#0f2350', '#e7ecf8', '#1d4ed8', '#0ea5e9', '#10b981', '#f59e0b', '#ef4444', '#64748b', '#1d4ed8', '#7c3aed'], headingFont: 'Open Sans', bodyFont: 'Carlito' },
  { name: 'Parchment', colors: ['#000000', '#ffffff', '#3b3226', '#f4ecdc', '#8b5e34', '#a68a64', '#6b705c', '#b5838d', '#457b9d', '#bc6c25', '#457b9d', '#b5838d'], headingFont: 'EB Garamond', bodyFont: 'EB Garamond' },
  { name: 'Monochrome', colors: ['#000000', '#ffffff', '#262626', '#f2f2f2', '#404040', '#595959', '#737373', '#8c8c8c', '#a6a6a6', '#bfbfbf', '#0563c1', '#954f72'], headingFont: 'Arimo', bodyFont: 'Arimo' },
];

export const FONT_SCHEMES: { name: string; heading: string; body: string }[] = [
  { name: 'Anne\'s Office', heading: 'Carlito', body: 'Carlito' },
  { name: 'Classic Serif', heading: 'Caladea', body: 'Carlito' },
  { name: 'Grotesque', heading: 'Arimo', body: 'Arimo' },
  { name: 'Newsprint', heading: 'Tinos', body: 'Arimo' },
  { name: 'Humanist', heading: 'Lato', body: 'Lato' },
  { name: 'Friendly', heading: 'Open Sans', body: 'Open Sans' },
  { name: 'Editorial', heading: 'Merriweather', body: 'Open Sans' },
  { name: 'Bookish', heading: 'Libre Baskerville', body: 'Lato' },
  { name: 'Modern Serif', heading: 'Gelasio', body: 'Lato' },
  { name: 'Old Style', heading: 'EB Garamond', body: 'EB Garamond' },
  { name: 'Report', heading: 'Tinos', body: 'Tinos' },
];

export const THEME_EFFECTS = ['Flat', 'Subtle Shadow', 'Soft Edges', 'Glossy', 'Matte', 'Frosted', 'Inset', 'Bold Outline'];

const COLOR_LABELS = ['Text/Background – Dark 1', 'Text/Background – Light 1', 'Text/Background – Dark 2', 'Text/Background – Light 2', 'Accent 1', 'Accent 2', 'Accent 3', 'Accent 4', 'Accent 5', 'Accent 6', 'Hyperlink', 'Followed Hyperlink'];

/** Applies a theme (or part of one) to the workbook; the default cell font follows the body font. */
export function applyTheme(app: App, patch: Partial<ThemeDef>, label = 'Theme'): void {
  const book = app.book;
  book.transact(label, () => {
    book.recordBook();
    const old = book.theme;
    const next: ThemeDef = { ...old, ...patch, colors: [...(patch.colors || old.colors)] };
    const base = book.styles[0] || {};
    if (patch.bodyFont && (!base.font || base.font === old.bodyFont)) book.styles[0] = { ...base, font: patch.bodyFont };
    book.theme = next;
  });
  app.grid.invalidate();
}

const swatches = (colors: string[]) => colors.slice(4, 10).map((c) => `linear-gradient(${c},${c})`).join(',');
const swatchCss = (colors: string[]) => `background-image:${swatches(colors)};background-size:${100 / 6}% 6px;background-position:${[0, 1, 2, 3, 4, 5].map((i) => `${i * 20}% 100%`).join(',')};background-repeat:no-repeat;padding-bottom:8px;`;

// --------------------------------------------------------------------------------------------
// Helpers

function refreshBreaks(app: App): void {
  const g = app.grid as any;
  if (!g) return;
  if (app.tab.viewMode !== 'normal' || g.showPageBreaks) g.pageBreaks = computePageBreaks(app.sheet);
  app.grid.invalidate();
}

/** Changes the print settings of the active sheet as one undoable step. */
export function setPrint(app: App, label: string, fn: (p: PrintSettings, s: Sheet) => void, sheet: Sheet = app.sheet): void {
  app.book.transact(label, () => { app.book.recordSheet(sheet); fn(sheet.print, sheet); });
  refreshBreaks(app);
}

const marginText = (m: PrintSettings['margins']) => `Top ${m.top}"  Bottom ${m.bottom}"  Left ${m.left}"  Right ${m.right}"`;

/** The selection as print-area text (whole rows/columns clipped to the used range). */
function selectionAreaText(app: App): string {
  const s = app.sheet;
  return selRanges(app).map((rg) => absRangeName(isWholeCols(rg) || isWholeRows(rg) ? clipToUsed(s, rg) : rg)).join(',');
}

export const HF_PRESETS: { l?: string; c?: string; r?: string }[] = [
  {},
  { c: 'Page &[Page]' },
  { c: 'Page &[Page] of &[Pages]' },
  { c: '&[Tab]' },
  { c: '&[File]' },
  { c: '&[Path]' },
  { c: '&[Tab], Page &[Page]' },
  { c: '&[File], Page &[Page]' },
  { c: '&[File]', r: '&[Tab]' },
  { l: 'Confidential', c: '&[Date]', r: 'Page &[Page]' },
  { l: 'Prepared by %AUTHOR% &[Date]', r: 'Page &[Page]' },
  { l: '&[Tab]', c: 'Page &[Page] of &[Pages]', r: '&[Date]' },
];

const HF_CODES: { code: string; label: string; icon: string }[] = [
  { code: '&[Page]', label: 'Page Number', icon: 'tag' }, { code: '&[Pages]', label: 'Number of Pages', icon: 'filter_9_plus' },
  { code: '&[Date]', label: 'Date', icon: 'calendar_today' }, { code: '&[Time]', label: 'Time', icon: 'schedule' },
  { code: '&[Path]&[File]', label: 'File Path', icon: 'folder' }, { code: '&[File]', label: 'File Name', icon: 'description' },
  { code: '&[Tab]', label: 'Sheet Name', icon: 'tab' }, { code: '&B', label: 'Bold on/off', icon: 'format_bold' }, { code: '&I', label: 'Italic on/off', icon: 'format_italic' },
];

// --------------------------------------------------------------------------------------------
// Page Setup dialog

export type SetupSection = 'page' | 'margins' | 'header' | 'sheet';

export async function pageSetupDialog(app: App, section: SetupSection = 'page'): Promise<void> {
  const s = app.sheet;
  const p = s.print;
  const ctx = { page: 1, pages: Math.max(1, layoutPages(s).pages.length), date: new Date(), file: app.tab.fileName || app.tab.title, path: app.tab.fileName || app.tab.title, sheet: s.name };
  const presetLabel = (pr: { l?: string; c?: string; r?: string }) => [pr.l, pr.c, pr.r].filter(Boolean).map((x) => expandHeaderFooter(x!.replace('%AUTHOR%', app.prefs.author), ctx)).join(', ') || '(none)';
  const presetOpts = HF_PRESETS.map((pr, i) => ({ value: String(i), label: presetLabel(pr) }));
  const sec = (v: string) => (vals: Record<string, any>) => vals.section === v;
  const hfFields = (prefix: string, title: string, src: { left: string; center: string; right: string } | undefined, visible: (v: Record<string, any>) => boolean): Field[] => [
    { key: prefix + 'h', label: title, type: 'heading', visibleIf: visible },
    { key: prefix + 'l', label: 'Left section', type: 'text', value: src?.left || '', visibleIf: visible },
    { key: prefix + 'c', label: 'Center section', type: 'text', value: src?.center || '', visibleIf: visible },
    { key: prefix + 'r', label: 'Right section', type: 'text', value: src?.right || '', visibleIf: visible },
  ];
  const isHf = sec('header');
  const fields: Field[] = [
    { key: 'section', type: 'radio', options: [{ value: 'page', label: 'Page' }, { value: 'margins', label: 'Margins' }, { value: 'header', label: 'Header/Footer' }, { value: 'sheet', label: 'Sheet' }], value: section },
    // Page
    { key: 'orientation', label: 'Orientation', type: 'radio', options: [{ value: 'portrait', label: 'Portrait' }, { value: 'landscape', label: 'Landscape' }], value: p.orientation, visibleIf: sec('page') },
    { key: 'scaleMode', label: 'Scaling', type: 'radio', options: [{ value: 'adjust', label: 'Adjust to % of normal size' }, { value: 'fit', label: 'Fit to pages' }], value: p.fitWide || p.fitTall ? 'fit' : 'adjust', visibleIf: sec('page') },
    { key: 'scale', label: 'Adjust to', type: 'number', min: 10, max: 400, suffix: '% normal size', value: p.scale || 100, visibleIf: (v) => v.section === 'page' && v.scaleMode === 'adjust' },
    { key: 'fitWide', label: 'Pages wide', type: 'number', min: 1, max: 32767, value: p.fitWide || '', placeholder: 'Automatic', half: true, visibleIf: (v) => v.section === 'page' && v.scaleMode === 'fit' },
    { key: 'fitTall', label: 'Pages tall', type: 'number', min: 1, max: 32767, value: p.fitTall || '', placeholder: 'Automatic', half: true, visibleIf: (v) => v.section === 'page' && v.scaleMode === 'fit' },
    { key: 'paper', label: 'Paper size', type: 'select', options: Object.keys(PAPER_SIZES).map((k) => ({ value: k, label: `${k} (${PAPER_SIZES[k][0]}" × ${PAPER_SIZES[k][1]}")` })), value: p.paper, visibleIf: sec('page') },
    { key: 'firstPage', label: 'First page number', type: 'text', value: p.firstPage != null ? String(p.firstPage) : 'Auto', visibleIf: sec('page') },
    // Margins
    { key: 'top', label: 'Top', type: 'number', min: 0, max: 20, step: 0.05, suffix: 'in', value: p.margins.top, half: true, visibleIf: sec('margins') },
    { key: 'bottom', label: 'Bottom', type: 'number', min: 0, max: 20, step: 0.05, suffix: 'in', value: p.margins.bottom, half: true, visibleIf: sec('margins') },
    { key: 'left', label: 'Left', type: 'number', min: 0, max: 20, step: 0.05, suffix: 'in', value: p.margins.left, half: true, visibleIf: sec('margins') },
    { key: 'right', label: 'Right', type: 'number', min: 0, max: 20, step: 0.05, suffix: 'in', value: p.margins.right, half: true, visibleIf: sec('margins') },
    { key: 'header', label: 'Header', type: 'number', min: 0, max: 20, step: 0.05, suffix: 'in', value: p.margins.header, half: true, visibleIf: sec('margins') },
    { key: 'footer', label: 'Footer', type: 'number', min: 0, max: 20, step: 0.05, suffix: 'in', value: p.margins.footer, half: true, visibleIf: sec('margins') },
    { key: 'centerH', label: 'Center on page horizontally', type: 'checkbox', value: p.centerH, visibleIf: sec('margins') },
    { key: 'centerV', label: 'Center on page vertically', type: 'checkbox', value: p.centerV, visibleIf: sec('margins') },
    // Header/Footer
    { key: 'headerPreset', label: 'Header', type: 'select', options: [{ value: '', label: '(choose a preset)' }, ...presetOpts], value: '', visibleIf: isHf },
    { key: 'footerPreset', label: 'Footer', type: 'select', options: [{ value: '', label: '(choose a preset)' }, ...presetOpts], value: '', visibleIf: isHf },
    { key: 'codes', type: 'info', html: '<div class="ql-hf-codes" role="toolbar" aria-label="Insert header and footer codes"></div>', visibleIf: isHf },
    ...hfFields('h', 'Custom header', p.header, isHf),
    ...hfFields('f', 'Custom footer', p.footer, isHf),
    { key: 'differentOddEven', label: 'Different odd and even pages', type: 'checkbox', value: !!p.differentOddEven, visibleIf: isHf },
    { key: 'differentFirst', label: 'Different first page', type: 'checkbox', value: !!p.differentFirst, visibleIf: isHf },
    ...hfFields('eh', 'Even page header', p.evenHeader, (v) => isHf(v) && v.differentOddEven),
    ...hfFields('ef', 'Even page footer', p.evenFooter, (v) => isHf(v) && v.differentOddEven),
    ...hfFields('ph', 'First page header', p.firstHeader, (v) => isHf(v) && v.differentFirst),
    ...hfFields('pf', 'First page footer', p.firstFooter, (v) => isHf(v) && v.differentFirst),
    { key: 'scaleWithDoc', label: 'Scale with document', type: 'checkbox', value: p.scaleWithDoc !== false, half: true, visibleIf: isHf },
    { key: 'alignWithMargins', label: 'Align with page margins', type: 'checkbox', value: p.alignWithMargins !== false, half: true, visibleIf: isHf },
    // Sheet
    { key: 'area', label: 'Print area', type: 'text', value: p.area || '', placeholder: 'e.g. A1:F40 (blank prints the used range)', visibleIf: sec('sheet') },
    { key: 'titleRows', label: 'Rows to repeat at top', type: 'text', value: p.titleRows || '', placeholder: 'e.g. $1:$2', visibleIf: sec('sheet') },
    { key: 'titleCols', label: 'Columns to repeat at left', type: 'text', value: p.titleCols || '', placeholder: 'e.g. $A:$A', visibleIf: sec('sheet') },
    { key: 'gridlines', label: 'Gridlines', type: 'checkbox', value: p.gridlines, half: true, visibleIf: sec('sheet') },
    { key: 'headings', label: 'Row and column headings', type: 'checkbox', value: p.headings, half: true, visibleIf: sec('sheet') },
    { key: 'blackAndWhite', label: 'Black and white', type: 'checkbox', value: !!p.blackAndWhite, half: true, visibleIf: sec('sheet') },
    { key: 'draft', label: 'Draft quality', type: 'checkbox', value: !!p.draft, half: true, visibleIf: sec('sheet') },
    { key: 'comments', label: 'Comments and notes', type: 'select', options: [{ value: 'none', label: '(None)' }, { value: 'atEnd', label: 'At end of sheet' }], value: p.comments === 'atEnd' ? 'atEnd' : 'none', half: true, visibleIf: sec('sheet') },
    { key: 'errors', label: 'Cell errors as', type: 'select', options: [{ value: 'displayed', label: 'displayed' }, { value: 'blank', label: '<blank>' }, { value: 'dashes', label: '--' }, { value: 'na', label: '#N/A' }], value: p.errors || 'displayed', half: true, visibleIf: sec('sheet') },
    { key: 'order', label: 'Page order', type: 'radio', options: [{ value: 'downThenOver', label: 'Down, then over' }, { value: 'overThenDown', label: 'Over, then down' }], value: p.order, visibleIf: sec('sheet') },
  ];
  let lastHf: HTMLInputElement | null = null;
  const setPreset = (ctl: DialogControl, prefix: 'h' | 'f', i: string) => {
    const pr = HF_PRESETS[+i];
    if (!pr) return;
    const fix = (x?: string) => (x || '').replace('%AUTHOR%', app.prefs.author);
    ctl.set(prefix + 'l', fix(pr.l)); ctl.set(prefix + 'c', fix(pr.c)); ctl.set(prefix + 'r', fix(pr.r));
  };
  const r = await openDialog({
    title: 'Page Setup', wide: true, fields,
    buttons: [
      { label: 'Print Preview', onClick: (_v, ctl) => ctl.close({ ...ctl.values(), __then: 'preview' }) },
      { label: 'Print…', onClick: (_v, ctl) => ctl.close({ ...ctl.values(), __then: 'print' }) },
    ],
    onOpen: (ctl) => {
      ctl.body.addEventListener('focusin', (e) => {
        const t = e.target as HTMLInputElement;
        if (t.tagName === 'INPUT' && /^f_(h|f|eh|ef|ph|pf)[lcr]_/.test(t.id)) lastHf = t;
      });
      const bar = ctl.body.querySelector<HTMLElement>('.ql-hf-codes');
      if (bar) {
        bar.appendChild(h('span', { class: 'ql-help' }, 'Insert: '));
        for (const c of HF_CODES) {
          const b = h('button', { class: 'ql-iconbtn', type: 'button', title: c.label, 'aria-label': `Insert ${c.label}` }, h('span', { class: 'ql-icon', 'aria-hidden': 'true' }, c.icon));
          b.addEventListener('mousedown', (e) => e.preventDefault());
          b.addEventListener('click', () => {
            const inp = lastHf || ctl.body.querySelector<HTMLInputElement>('input[id^="f_hc_"]');
            if (!inp) return;
            const a = inp.selectionStart ?? inp.value.length, z = inp.selectionEnd ?? a;
            inp.value = inp.value.slice(0, a) + c.code + inp.value.slice(z);
            inp.setSelectionRange(a + c.code.length, a + c.code.length);
            inp.focus();
          });
          bar.appendChild(b);
        }
      }
    },
    onChange: (v, ctl, key) => {
      if (key === 'headerPreset' && v.headerPreset !== '') setPreset(ctl, 'h', v.headerPreset);
      if (key === 'footerPreset' && v.footerPreset !== '') setPreset(ctl, 'f', v.footerPreset);
    },
    validate: (v) => {
      for (const part of splitAreaText(v.area || '')) if (!resolveRangeText(app.book, part, s)) return `The print area "${part}" is not a valid reference.`;
      if (v.titleRows && !parseTitleSpan(v.titleRows, 'row')) return 'Rows to repeat must be whole rows, like $1:$2.';
      if (v.titleCols && !parseTitleSpan(v.titleCols, 'col')) return 'Columns to repeat must be whole columns, like $A:$B.';
      if (v.scaleMode === 'adjust' && !(v.scale >= 10 && v.scale <= 400)) return 'The scale must be between 10% and 400%.';
      const fp = String(v.firstPage || '').trim();
      if (fp && !/^auto$/i.test(fp) && !/^-?\d+$/.test(fp)) return 'The first page number must be a whole number or "Auto".';
      const [w, hh] = PAPER_SIZES[v.paper] || PAPER_SIZES.Letter;
      const [pw, ph] = v.orientation === 'landscape' ? [hh, w] : [w, hh];
      if ((v.left || 0) + (v.right || 0) >= pw || (v.top || 0) + (v.bottom || 0) >= ph) return 'The margins do not fit the paper size.';
      return null;
    },
  });
  if (!r) return;
  const sec3 = (pfx: string) => ({ left: r[pfx + 'l'] || '', center: r[pfx + 'c'] || '', right: r[pfx + 'r'] || '' });
  setPrint(app, 'Page Setup', (pp) => {
    pp.orientation = r.orientation;
    if (r.scaleMode === 'fit') { pp.fitWide = r.fitWide > 0 ? Math.round(r.fitWide) : undefined; pp.fitTall = r.fitTall > 0 ? Math.round(r.fitTall) : undefined; if (!pp.fitWide && !pp.fitTall) pp.fitWide = pp.fitTall = 1; }
    else { pp.fitWide = undefined; pp.fitTall = undefined; pp.scale = Math.round(r.scale); }
    pp.paper = r.paper;
    const fp = String(r.firstPage || '').trim();
    pp.firstPage = /^-?\d+$/.test(fp) ? parseInt(fp, 10) : undefined;
    pp.margins = { top: r.top ?? pp.margins.top, bottom: r.bottom ?? pp.margins.bottom, left: r.left ?? pp.margins.left, right: r.right ?? pp.margins.right, header: r.header ?? pp.margins.header, footer: r.footer ?? pp.margins.footer };
    pp.centerH = !!r.centerH; pp.centerV = !!r.centerV;
    pp.header = sec3('h'); pp.footer = sec3('f');
    pp.differentOddEven = !!r.differentOddEven || undefined; pp.differentFirst = !!r.differentFirst || undefined;
    pp.evenHeader = r.differentOddEven ? sec3('eh') : pp.evenHeader; pp.evenFooter = r.differentOddEven ? sec3('ef') : pp.evenFooter;
    pp.firstHeader = r.differentFirst ? sec3('ph') : pp.firstHeader; pp.firstFooter = r.differentFirst ? sec3('pf') : pp.firstFooter;
    pp.scaleWithDoc = r.scaleWithDoc ? undefined : false; pp.alignWithMargins = r.alignWithMargins ? undefined : false;
    pp.area = (r.area || '').trim() || undefined;
    pp.titleRows = (r.titleRows || '').trim() || undefined;
    pp.titleCols = (r.titleCols || '').trim() || undefined;
    pp.gridlines = !!r.gridlines; pp.headings = !!r.headings; pp.blackAndWhite = !!r.blackAndWhite || undefined; pp.draft = !!r.draft || undefined;
    pp.comments = r.comments === 'atEnd' ? 'atEnd' : undefined;
    pp.errors = r.errors && r.errors !== 'displayed' ? r.errors : undefined;
    pp.order = r.order;
  });
  if (r.__then === 'preview') await app.exec('file.printPreview');
  else if (r.__then === 'print') await app.exec('file.print');
}

// --------------------------------------------------------------------------------------------
// Theme dialogs

async function customizeColors(app: App): Promise<void> {
  const t = app.book.theme;
  const fields: Field[] = [
    ...COLOR_LABELS.map((label, i) => ({ key: 'c' + i, label, type: 'color' as const, value: t.colors[i] || '#000000', half: true })),
    { key: 'name', label: 'Name', type: 'text', value: THEMES.some((x) => x.name === t.name) ? 'Custom 1' : t.name },
  ];
  const r = await openDialog({
    title: 'Create New Theme Colors', fields, preview: true, ok: 'Save',
    buttons: [{ label: 'Reset', onClick: (_v, ctl) => { t.colors.forEach((c, i) => ctl.set('c' + i, c)); } }],
    onChange: (v, ctl) => {
      const cols = COLOR_LABELS.map((_l, i) => v['c' + i]);
      ctl.setPreview(`<div style="display:flex;gap:4px;align-items:flex-end;height:60px">${cols.slice(4, 10).map((c, i) => `<div style="flex:1;height:${30 + i * 5}px;background:${c}"></div>`).join('')}</div><p style="margin:6px 0 0;color:${cols[0]};background:${cols[1]}">Text <a style="color:${cols[10]}">Hyperlink</a> <a style="color:${cols[11]}">Followed hyperlink</a></p>`);
    },
  });
  if (!r) return;
  applyTheme(app, { colors: COLOR_LABELS.map((_l, i) => String(r['c' + i])), name: String(r.name || 'Custom') }, 'Theme Colors');
}

async function customizeFonts(app: App): Promise<void> {
  const t = app.book.theme;
  const r = await openDialog({
    title: 'Create New Theme Fonts', preview: true, ok: 'Save',
    fields: [
      { key: 'heading', label: 'Heading font', type: 'combo', options: FONT_LIST, value: t.headingFont },
      { key: 'body', label: 'Body font', type: 'combo', options: FONT_LIST, value: t.bodyFont },
    ],
    onChange: (v, ctl) => ctl.setPreview(`<div style="font-family:'${String(v.heading).replace(/['"<>;]/g, '')}';font-size:18px">Heading</div><div style="font-family:'${String(v.body).replace(/['"<>;]/g, '')}'">Body text body text body text.</div>`),
  });
  if (!r || !r.heading || !r.body) return;
  applyTheme(app, { headingFont: r.heading, bodyFont: r.body }, 'Theme Fonts');
}

// --------------------------------------------------------------------------------------------
// Commands

const PL = 'Page Layout';
const FIT_MENU = (id: string, cur: number | undefined): MenuEntry[] => [
  { id, arg: 0, label: 'Automatic', checked: !cur },
  ...[1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => ({ id, arg: n, label: `${n} page${n > 1 ? 's' : ''}`, checked: cur === n })),
  { separator: true }, { id: 'layout.pageSetup', arg: 'page', label: 'More Pages…', icon: 'more_horiz' },
];

def([
  // Themes
  {
    id: 'layout.themes', label: 'Themes', icon: 'palette', tab: PL, group: 'Themes', kind: 'menu', edits: true, popular: true,
    tip: 'Change the colors and fonts of the whole workbook at once',
    run: async (app, name) => {
      if (name === 'reset') return applyTheme(app, { ...DEFAULT_THEME, effects: undefined }, 'Reset Theme');
      if (typeof name !== 'string') {
        const r = await openDialog({ title: 'Themes', fields: [{ key: 't', label: 'Theme', type: 'list', rows: 10, options: THEMES.map((t) => ({ value: t.name, label: `${t.name} — ${t.headingFont} / ${t.bodyFont}` })), value: app.book.theme.name }] });
        if (!r?.t) return;
        name = r.t;
      }
      const t = THEMES.find((x) => x.name === name);
      if (!t) { toast(`Unknown theme: ${name}`); return; }
      applyTheme(app, { ...t, colors: [...t.colors] }, `Theme: ${t.name}`);
    },
    menu: (app) => [
      { heading: 'Built-In' },
      ...THEMES.map((t) => ({ id: 'layout.themes', arg: t.name, label: `${t.name}  ·  ${t.headingFont}${t.bodyFont !== t.headingFont ? ' / ' + t.bodyFont : ''}`, checked: app.book.theme.name === t.name, preview: `font-family:"${t.headingFont}";${swatchCss(t.colors)}` })),
      { separator: true },
      { id: 'layout.themes', arg: 'reset', label: 'Reset to Theme from Template', icon: 'restart_alt' },
    ],
  },
  {
    id: 'layout.themeColors', label: 'Colors', icon: 'format_color_fill', tab: PL, group: 'Themes', kind: 'menu', edits: true,
    tip: 'Choose the theme color palette',
    run: (app, name) => {
      if (name === 'custom' || name === undefined) return customizeColors(app);
      const t = THEMES.find((x) => x.name === name);
      if (t) applyTheme(app, { colors: [...t.colors], name: t.name === app.book.theme.name ? t.name : `${app.book.theme.name.replace(/ \(.*\)$/, '')} (${t.name} colors)` }, 'Theme Colors');
    },
    menu: (app) => [
      ...THEMES.map((t) => ({ id: 'layout.themeColors', arg: t.name, label: t.name, checked: app.book.theme.colors.join() === t.colors.join(), preview: swatchCss(t.colors) })),
      { separator: true }, { id: 'layout.themeColors', arg: 'custom', label: 'Customize Colors…', icon: 'tune' },
    ],
  },
  {
    id: 'layout.themeFonts', label: 'Fonts', icon: 'text_fields', tab: PL, group: 'Themes', kind: 'menu', edits: true,
    tip: 'Choose the heading and body fonts; cells using the default font follow the body font',
    run: (app, name) => {
      if (name === 'custom' || name === undefined) return customizeFonts(app);
      const f = FONT_SCHEMES.find((x) => x.name === name);
      if (f) applyTheme(app, { headingFont: f.heading, bodyFont: f.body }, 'Theme Fonts');
    },
    menu: (app) => [
      ...FONT_SCHEMES.map((f) => ({ id: 'layout.themeFonts', arg: f.name, label: `${f.name}: ${f.heading}${f.body !== f.heading ? ' / ' + f.body : ''}`, checked: app.book.theme.headingFont === f.heading && app.book.theme.bodyFont === f.body, preview: `font-family:"${f.heading}"` })),
      { separator: true }, { id: 'layout.themeFonts', arg: 'custom', label: 'Customize Fonts…', icon: 'tune' },
    ],
  },
  {
    id: 'layout.themeEffects', label: 'Effects', icon: 'auto_awesome', tab: PL, group: 'Themes', kind: 'menu', edits: true,
    tip: 'Theme effects are stored with the workbook theme (and kept in saved files); charts and shapes here keep their own look',
    run: (app, name) => { if (typeof name === 'string' && THEME_EFFECTS.includes(name)) { applyTheme(app, { effects: name }, 'Theme Effects'); toast(`Theme effects: ${name}. Effects are stored with the theme; charts and shapes keep their own formatting.`); } },
    menu: (app) => [
      { heading: 'Stored with the theme; charts and shapes keep their own look' },
      ...THEME_EFFECTS.map((e) => ({ id: 'layout.themeEffects', arg: e, label: e, checked: (app.book.theme.effects || 'Flat') === e })),
    ],
  },
  // Page Setup group
  {
    id: 'layout.margins', label: 'Margins', icon: 'border_outer', tab: PL, group: 'Page Setup', kind: 'menu', edits: true,
    run: (app, arg) => {
      if (arg === 'custom' || arg === undefined) return pageSetupDialog(app, 'margins');
      if (arg === 'last') { const last = (app.tab as any).lastCustomMargins; if (last) setPrint(app, 'Margins', (p) => { p.margins = { ...last }; }); return; }
      const pr = MARGIN_PRESETS[arg as MarginPreset];
      if (!pr) return;
      const cur = app.sheet.print.margins;
      if (!Object.values(MARGIN_PRESETS).some((x) => JSON.stringify(x.m) === JSON.stringify(cur))) (app.tab as any).lastCustomMargins = { ...cur };
      setPrint(app, 'Margins', (p) => { p.margins = { ...pr.m }; });
    },
    menu: (app) => {
      const cur = JSON.stringify(app.sheet.print.margins);
      const last = (app.tab as any).lastCustomMargins;
      return [
        ...(last ? [{ id: 'layout.margins', arg: 'last', label: `Last Custom Setting — ${marginText(last)}`, icon: 'history' }] : []),
        ...(Object.keys(MARGIN_PRESETS) as MarginPreset[]).map((k) => ({ id: 'layout.margins', arg: k, label: `${MARGIN_PRESETS[k].label} — ${marginText(MARGIN_PRESETS[k].m)}`, checked: JSON.stringify(MARGIN_PRESETS[k].m) === cur })),
        { separator: true }, { id: 'layout.margins', arg: 'custom', label: 'Custom Margins…', icon: 'tune' },
      ];
    },
  },
  {
    id: 'layout.orientation', label: 'Orientation', icon: 'crop_rotate', tab: PL, group: 'Page Setup', kind: 'menu', edits: true,
    run: (app, arg) => {
      const o = arg === 'portrait' || arg === 'landscape' ? arg : app.sheet.print.orientation === 'portrait' ? 'landscape' : 'portrait';
      setPrint(app, 'Orientation', (p) => { p.orientation = o; });
    },
    menu: (app) => [{ id: 'layout.orientation', arg: 'portrait', label: 'Portrait', icon: 'crop_portrait', checked: app.sheet.print.orientation === 'portrait' }, { id: 'layout.orientation', arg: 'landscape', label: 'Landscape', icon: 'crop_landscape', checked: app.sheet.print.orientation === 'landscape' }],
  },
  {
    id: 'layout.size', label: 'Size', icon: 'description', tab: PL, group: 'Page Setup', kind: 'menu', edits: true,
    run: (app, arg) => {
      if (typeof arg !== 'string' || !PAPER_SIZES[arg]) return pageSetupDialog(app, 'page');
      setPrint(app, 'Paper Size', (p) => { p.paper = arg; });
    },
    menu: (app) => [
      ...Object.entries(PAPER_SIZES).map(([k, [w, hh]]) => ({ id: 'layout.size', arg: k, label: `${k} — ${w}" × ${hh}"`, checked: app.sheet.print.paper === k })),
      { separator: true }, { id: 'layout.size', arg: 'more', label: 'More Paper Sizes…', icon: 'more_horiz' },
    ],
  },
  {
    id: 'layout.printArea', label: 'Print Area', icon: 'print_connect', tab: PL, group: 'Page Setup', kind: 'menu', edits: true,
    run: (app, arg) => {
      const sel = selectionAreaText(app);
      if (arg === 'clear') { setPrint(app, 'Clear Print Area', (p) => { p.area = undefined; }); toast('Print area cleared'); return; }
      if (arg === 'add' && app.sheet.print.area) { setPrint(app, 'Add to Print Area', (p) => { p.area = `${p.area},${sel}`; }); toast(`Print area: ${app.sheet.print.area}`); return; }
      setPrint(app, 'Set Print Area', (p) => { p.area = sel; });
      toast(`Print area set to ${sel}`);
    },
    menu: (app) => [
      { id: 'layout.printArea', arg: 'set', label: 'Set Print Area', icon: 'select_all' },
      ...(app.sheet.print.area ? [{ id: 'layout.printArea', arg: 'add', label: 'Add to Print Area', icon: 'add_box' }, { id: 'layout.printArea', arg: 'clear', label: 'Clear Print Area', icon: 'clear' }] : []),
    ],
  },
  {
    id: 'layout.breaks', label: 'Breaks', icon: 'insert_page_break', tab: PL, group: 'Page Setup', kind: 'menu', edits: true,
    run: (app, arg) => {
      const s = app.sheet;
      const { r, c } = activeCell(app);
      const rg = selRanges(app)[selRanges(app).length - 1];
      const rowsOnly = isWholeRows(rg), colsOnly = isWholeCols(rg);
      if (arg === 'reset') {
        if (!s.rowBreaks.length && !s.colBreaks.length) { toast('There are no manual page breaks.'); return; }
        setPrint(app, 'Reset All Page Breaks', (_p, sh) => { sh.rowBreaks = []; sh.colBreaks = []; });
        return;
      }
      if (arg === 'remove') {
        const hasR = s.rowBreaks.includes(r) && !colsOnly, hasC = s.colBreaks.includes(c) && !rowsOnly;
        if (!hasR && !hasC) { toast('There is no manual page break at the active cell.'); return; }
        setPrint(app, 'Remove Page Break', (_p, sh) => { if (hasR) sh.rowBreaks = sh.rowBreaks.filter((x) => x !== r); if (hasC) sh.colBreaks = sh.colBreaks.filter((x) => x !== c); });
        return;
      }
      // Insert above and to the left of the active cell (only one direction for whole rows/columns).
      if (r === 0 && c === 0) { toast('Select a cell below or to the right of where the new page should start.'); return; }
      setPrint(app, 'Insert Page Break', (_p, sh) => {
        if (r > 0 && !colsOnly && !sh.rowBreaks.includes(r)) sh.rowBreaks = [...sh.rowBreaks, r].sort((a, b) => a - b);
        if (c > 0 && !rowsOnly && !sh.colBreaks.includes(c)) sh.colBreaks = [...sh.colBreaks, c].sort((a, b) => a - b);
      });
      if (app.sheet.print.fitWide || app.sheet.print.fitTall) toast('Manual page breaks are ignored while Scale to Fit is set to a number of pages.');
    },
    menu: () => [
      { id: 'layout.breaks', arg: 'insert', label: 'Insert Page Break', icon: 'insert_page_break' },
      { id: 'layout.breaks', arg: 'remove', label: 'Remove Page Break', icon: 'remove' },
      { id: 'layout.breaks', arg: 'reset', label: 'Reset All Page Breaks', icon: 'restart_alt' },
    ],
  },
  {
    id: 'layout.background', label: 'Background', icon: 'wallpaper', tab: PL, group: 'Page Setup', edits: true,
    tip: 'Pick a picture to show behind the sheet. It appears in Print Preview and printouts; the grid on screen stays plain.',
    run: async (app, arg) => {
      const s = app.sheet;
      if (arg === 'delete' || (s.background && arg !== 'pick')) {
        const ch = arg === 'delete' ? 1 : await choiceDialog('Sheet Background', 'This sheet has a background picture.', ['Replace…', 'Delete Background']);
        if (ch < 0) return;
        if (ch === 1) { setPrint(app, 'Delete Background', (_p, sh) => { sh.background = null; }); toast('Background deleted'); return; }
      }
      const [f] = await pickFile('image/*');
      if (!f) return;
      if (f.size > 8 * 1024 * 1024) { toast('That picture is too large (over 8 MB).'); return; }
      const url = await readFileAsDataURL(f);
      if (!/^data:image\//.test(url)) { toast('That file is not a picture.'); return; }
      setPrint(app, 'Background', (_p, sh) => { sh.background = url; }, s);
      toast('Background set. It shows in Print Preview and when printing; the grid on screen stays plain.', { ms: 5000 });
    },
  },
  { id: 'layout.printTitles', label: 'Print Titles', icon: 'title', tab: PL, group: 'Page Setup', edits: true, tip: 'Rows and columns to repeat on every printed page', run: (app) => pageSetupDialog(app, 'sheet') },
  {
    id: 'layout.pageSetup', label: 'Page Setup…', icon: 'settings_applications', tab: PL, group: 'Page Setup', edits: true,
    run: (app, arg) => pageSetupDialog(app, (['page', 'margins', 'header', 'sheet'].includes(arg) ? arg : arg === 'headerFooter' ? 'header' : 'page') as SetupSection),
  },
  { id: 'layout.headerFooter', label: 'Header & Footer…', icon: 'call_to_action', tab: PL, group: 'Page Setup', edits: true, run: (app) => pageSetupDialog(app, 'header') },
  // Scale to Fit
  {
    id: 'layout.fitWidth', label: 'Width', icon: 'width', tab: PL, group: 'Scale to Fit', kind: 'combo', edits: true,
    value: (app) => (app.sheet.print.fitWide ? `${app.sheet.print.fitWide} page${app.sheet.print.fitWide > 1 ? 's' : ''}` : 'Automatic'),
    run: (app, n) => {
      const v = Math.max(0, Math.round(Number(n) || 0));
      setPrint(app, 'Scale to Fit Width', (p) => { p.fitWide = v || undefined; if (!p.fitWide && !p.fitTall) p.scale = 100; });
    },
    menu: (app) => FIT_MENU('layout.fitWidth', app.sheet.print.fitWide),
  },
  {
    id: 'layout.fitHeight', label: 'Height', icon: 'height', tab: PL, group: 'Scale to Fit', kind: 'combo', edits: true,
    value: (app) => (app.sheet.print.fitTall ? `${app.sheet.print.fitTall} page${app.sheet.print.fitTall > 1 ? 's' : ''}` : 'Automatic'),
    run: (app, n) => {
      const v = Math.max(0, Math.round(Number(n) || 0));
      setPrint(app, 'Scale to Fit Height', (p) => { p.fitTall = v || undefined; if (!p.fitWide && !p.fitTall) p.scale = 100; });
    },
    menu: (app) => FIT_MENU('layout.fitHeight', app.sheet.print.fitTall),
  },
  {
    id: 'layout.scale', label: 'Scale', icon: 'aspect_ratio', tab: PL, group: 'Scale to Fit', kind: 'size', edits: true,
    tip: 'Print at a percentage of normal size (turns off fitting to a number of pages)',
    value: (app) => { const p = app.sheet.print; return `${p.fitWide || p.fitTall ? layoutPages(app.sheet).scale : p.scale || 100}%`; },
    run: (app, v) => {
      const n = Math.round(parseFloat(String(v)));
      if (!(n >= 10 && n <= 400)) { toast('The scale must be between 10% and 400%.'); return; }
      setPrint(app, 'Scale', (p) => { p.scale = n; p.fitWide = undefined; p.fitTall = undefined; });
    },
    menu: () => [25, 50, 75, 85, 90, 100, 110, 125, 150, 200].map((n) => ({ id: 'layout.scale', arg: n, label: `${n}%` })),
  },
  {
    id: 'layout.scaling', label: 'Scaling', icon: 'fit_screen', tab: PL, group: 'Scale to Fit', kind: 'menu', edits: true,
    tip: 'No scaling, fit the sheet on one page, or fit all columns or all rows on one page',
    run: (app, arg) => {
      const map: Record<string, [number | undefined, number | undefined]> = { none: [undefined, undefined], fitSheet: [1, 1], fitCols: [1, undefined], fitRows: [undefined, 1] };
      if (arg === 'custom' || !(arg in map)) return pageSetupDialog(app, 'page');
      setPrint(app, 'Scaling', (p) => { [p.fitWide, p.fitTall] = map[arg]; if (arg === 'none') p.scale = 100; });
    },
    menu: (app) => {
      const p = app.sheet.print;
      return [
        { id: 'layout.scaling', arg: 'none', label: 'No Scaling', checked: !p.fitWide && !p.fitTall && (p.scale || 100) === 100 },
        { id: 'layout.scaling', arg: 'fitSheet', label: 'Fit Sheet on One Page', checked: p.fitWide === 1 && p.fitTall === 1 },
        { id: 'layout.scaling', arg: 'fitCols', label: 'Fit All Columns on One Page', checked: p.fitWide === 1 && !p.fitTall },
        { id: 'layout.scaling', arg: 'fitRows', label: 'Fit All Rows on One Page', checked: !p.fitWide && p.fitTall === 1 },
        { separator: true }, { id: 'layout.scaling', arg: 'custom', label: 'Custom Scaling Options…', icon: 'tune' },
      ];
    },
  },
  // Sheet Options
  { id: 'layout.printGridlines', label: 'Print Gridlines', icon: 'grid_on', tab: PL, group: 'Sheet Options', kind: 'toggle', edits: true, run: (app) => setPrint(app, 'Print Gridlines', (p) => { p.gridlines = !p.gridlines; }), isActive: (app) => app.sheet.print.gridlines },
  { id: 'layout.printHeadings', label: 'Print Headings', icon: 'view_headline', tab: PL, group: 'Sheet Options', kind: 'toggle', edits: true, run: (app) => setPrint(app, 'Print Headings', (p) => { p.headings = !p.headings; }), isActive: (app) => app.sheet.print.headings },
  { id: 'layout.blackAndWhite', label: 'Print in Black and White', icon: 'contrast', tab: PL, group: 'Sheet Options', kind: 'toggle', edits: true, run: (app) => setPrint(app, 'Black and White', (p) => { p.blackAndWhite = !p.blackAndWhite || undefined; }), isActive: (app) => !!app.sheet.print.blackAndWhite },
  {
    id: 'layout.centerOnPage', label: 'Center on Page', icon: 'center_focus_strong', tab: PL, group: 'Sheet Options', kind: 'toggle', edits: true,
    run: (app) => setPrint(app, 'Center on Page', (p) => { const on = !(p.centerH && p.centerV); p.centerH = on; p.centerV = on; }), isActive: (app) => app.sheet.print.centerH && app.sheet.print.centerV,
  },
]);

/** Text of the print area for display (used by tests and the Page Setup field). */
export function printAreaLabel(sheet: Sheet): string { return sheet.print.area || (sheet.usedRange() ? rangeName(sheet.usedRange()!) : ''); }
