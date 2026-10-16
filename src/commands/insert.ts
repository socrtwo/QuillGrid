/**
 * Insert tab: pictures, screenshots, shapes, icons, diagrams, text boxes, Decorative Text, signature
 * lines, embedded files, equations, symbols, emoji, in-cell checkboxes, links, header & footer,
 * sparklines, slicers and timelines. (Tables live in table.ts and charts in charts.ts.)
 */
import { def, type MenuEntry } from './registry';
import type { App } from '../app';
import type { DrawingObject, Sparkline, Cell, FilterCriteria } from '../model/types';
import type { Sheet } from '../model/workbook';
import { newId } from '../model/workbook';
import { type Range, cellName, parseCell, rangeName, quoteSheet, keyRow, keyCol } from '../model/address';
import { resolveRangeText, rangeRefText } from '../model/structure';
import { isDateFormat } from '../model/numfmt';
import { serialToDate, dateToSerial } from '../model/dates';
import { activeRange, clipToUsed, cellText, selRanges } from './ops';
import { openDialog, toast, showMenu, colorMenu, promptDialog } from '../ui/dialog';
import { h, pickFile, readFileAsDataURL, icon } from '../ui/dom';
import { SHAPES, SHAPE_GROUPS, shapePreviewSvg, WORDART_STYLES, DIAGRAMS, diagramSvg, renderShapeSvg, equationRuns } from '../charts/shapes';
import { addDrawing, galleryDialog, viewAnchor, cellAnchor, canEditObjects, findObject, mutateObject, selectedDrawing } from './objects';
import { svgToPng } from './charts';
import { tableAt, tableDataRows } from './table';
import { applyFilters } from './dataops';
import { isScriptUrl, isSafeExternalUrl, escHtml } from '../util/sanitize';

const accent = (app: App, i = 0) => app.book.theme.colors[4 + i] || '#2f6fb0';

// --------------------------------------------------------------------------------------------
// Pictures

async function imageSize(src: string): Promise<{ w: number; h: number }> {
  return new Promise((res) => {
    const img = new Image();
    img.onload = () => res({ w: img.naturalWidth || 200, h: img.naturalHeight || 150 });
    img.onerror = () => res({ w: 200, h: 150 });
    img.src = src;
  });
}

/** Converts a picture file to a safe data URL (SVG pictures are rasterized). */
export async function pictureDataUrl(file: Blob): Promise<string | null> {
  let src = await readFileAsDataURL(file);
  if (!/^data:image\//i.test(src)) return null;
  if (/^data:image\/svg/i.test(src)) {
    const text = await file.text();
    const sz = await imageSize(src);
    const png = await svgToPng(text, sz.w, sz.h, 1);
    if (!png) return null;
    src = await readFileAsDataURL(png);
  }
  return src;
}

export async function insertPictureFile(app: App, file: Blob, name = 'Picture'): Promise<DrawingObject | null> {
  if (!canEditObjects(app)) return null;
  const src = await pictureDataUrl(file);
  if (!src) { toast('That file is not a picture QuillLattice can show.'); return null; }
  const sz = await imageSize(src);
  const k = Math.min(1, 640 / Math.max(sz.w, sz.h));
  return addDrawing(app, { kind: 'image', at: { ...cellAnchor(app), w: Math.round(sz.w * k), h: Math.round(sz.h * k) }, src, alt: name.replace(/\.[^.]+$/, ''), lockAspect: true }, 'Insert Picture');
}

async function pickPictures(): Promise<File[]> {
  try {
    const { isNative, openDocument } = await import('../platform/native');
    if (isNative()) {
      const f = await openDocument('image/*', ['image/png', 'image/jpeg', 'image/gif', 'image/webp', 'image/bmp', 'image/svg+xml']);
      return f ? [new File([f.data], f.name, { type: f.mime || 'image/png' })] : [];
    }
  } catch { /* fall back to the browser picker */ }
  return pickFile('image/*', true);
}

async function captureScreen(): Promise<Blob | null> {
  const md = (navigator as any).mediaDevices;
  if (!md?.getDisplayMedia) return null;
  let stream: MediaStream | null = null;
  try {
    stream = await md.getDisplayMedia({ video: true, audio: false });
    const video = document.createElement('video');
    video.srcObject = stream;
    video.muted = true;
    await video.play();
    await new Promise((r) => setTimeout(r, 250));
    const cv = document.createElement('canvas');
    cv.width = video.videoWidth; cv.height = video.videoHeight;
    cv.getContext('2d')!.drawImage(video, 0, 0);
    return await new Promise<Blob | null>((res) => cv.toBlob((b) => res(b), 'image/png'));
  } catch { return null; } finally { stream?.getTracks().forEach((t) => t.stop()); }
}

// --------------------------------------------------------------------------------------------
// Icons (Material Symbols names, rendered with the bundled icon font)

export const ICONS: [string, string[]][] = [
  ['Arrows', ['arrow_forward', 'arrow_back', 'arrow_upward', 'arrow_downward', 'north_east', 'south_east', 'trending_up', 'trending_down', 'sync', 'autorenew', 'swap_horiz', 'double_arrow', 'subdirectory_arrow_right', 'redo', 'undo']],
  ['Business', ['business', 'work', 'store', 'storefront', 'account_balance', 'payments', 'savings', 'paid', 'request_quote', 'receipt_long', 'shopping_cart', 'inventory_2', 'local_shipping', 'handshake', 'monitoring']],
  ['Charts', ['bar_chart', 'show_chart', 'pie_chart', 'donut_large', 'scatter_plot', 'insights', 'analytics', 'leaderboard', 'query_stats', 'timeline', 'stacked_line_chart', 'area_chart', 'bubble_chart', 'candlestick_chart', 'table_chart']],
  ['Communication', ['mail', 'call', 'chat', 'forum', 'campaign', 'notifications', 'contact_mail', 'contact_phone', 'send', 'share', 'rss_feed', 'alternate_email', 'sms', 'voicemail', 'support_agent']],
  ['People', ['person', 'group', 'groups', 'face', 'person_add', 'diversity_3', 'family_restroom', 'badge', 'emoji_people', 'accessibility_new', 'school', 'engineering', 'psychology', 'self_improvement', 'sports_handball']],
  ['Interface', ['home', 'search', 'settings', 'menu', 'close', 'check', 'check_circle', 'cancel', 'add_circle', 'remove_circle', 'info', 'help', 'warning', 'error', 'star', 'favorite', 'thumb_up', 'thumb_down', 'visibility', 'lock', 'lock_open', 'key', 'bookmark', 'flag', 'push_pin']],
  ['Files & Office', ['description', 'folder', 'folder_open', 'article', 'attach_file', 'print', 'save', 'calendar_month', 'event', 'schedule', 'alarm', 'task_alt', 'checklist', 'assignment', 'content_paste', 'edit', 'draw', 'calculate', 'functions', 'grid_on']],
  ['Technology', ['computer', 'laptop', 'smartphone', 'tablet', 'cloud', 'cloud_upload', 'cloud_download', 'wifi', 'bluetooth', 'database', 'dns', 'memory', 'router', 'code', 'terminal', 'bug_report', 'security', 'smart_toy', 'battery_full', 'power']],
  ['Places & Travel', ['public', 'map', 'location_on', 'flight', 'directions_car', 'directions_bus', 'train', 'directions_bike', 'hotel', 'restaurant', 'local_cafe', 'apartment', 'factory', 'local_hospital', 'park']],
  ['Nature & Weather', ['wb_sunny', 'dark_mode', 'cloud_queue', 'thunderstorm', 'ac_unit', 'water_drop', 'eco', 'forest', 'local_florist', 'pets', 'bolt', 'whatshot', 'spa', 'grass', 'landscape']],
  ['Signs & Symbols', ['lightbulb', 'emoji_objects', 'rocket_launch', 'emoji_events', 'military_tech', 'verified', 'new_releases', 'priority_high', 'do_not_disturb_on', 'recycling', 'health_and_safety', 'gavel', 'balance', 'target', 'extension']],
];

// --------------------------------------------------------------------------------------------
// Symbols and emoji

const SYMBOL_BLOCKS: [string, number, number][] = [
  ['Latin-1 Supplement', 0x00a1, 0x00ff], ['Latin Extended-A', 0x0100, 0x017f], ['Greek and Coptic', 0x0391, 0x03c9], ['Cyrillic', 0x0410, 0x044f],
  ['General Punctuation', 0x2010, 0x205e], ['Superscripts and Subscripts', 0x2070, 0x209c], ['Currency Symbols', 0x20a0, 0x20c0], ['Letterlike Symbols', 0x2100, 0x214f],
  ['Number Forms', 0x2150, 0x218b], ['Arrows', 0x2190, 0x21ff], ['Mathematical Operators', 0x2200, 0x22ff], ['Miscellaneous Technical', 0x2300, 0x23ff],
  ['Box Drawing', 0x2500, 0x257f], ['Block Elements', 0x2580, 0x259f], ['Geometric Shapes', 0x25a0, 0x25ff], ['Miscellaneous Symbols', 0x2600, 0x26ff],
  ['Dingbats', 0x2700, 0x27bf],
];
const SPECIAL_CHARS: [string, string][] = [
  ['—', 'Em Dash'], ['–', 'En Dash'], ['\u00a0', 'Nonbreaking Space'], ['©', 'Copyright'], ['®', 'Registered'], ['™', 'Trademark'], ['§', 'Section'], ['¶', 'Paragraph'],
  ['…', 'Ellipsis'], ['‘', 'Single Opening Quote'], ['’', 'Single Closing Quote'], ['“', 'Double Opening Quote'], ['”', 'Double Closing Quote'], ['•', 'Bullet'], ['°', 'Degree'],
  ['±', 'Plus-Minus'], ['×', 'Multiplication'], ['÷', 'Division'], ['≠', 'Not Equal'], ['≤', 'Less or Equal'], ['≥', 'Greater or Equal'], ['∞', 'Infinity'], ['€', 'Euro'], ['£', 'Pound'], ['¥', 'Yen'],
  ['½', 'One Half'], ['¼', 'One Quarter'], ['¾', 'Three Quarters'], ['→', 'Right Arrow'], ['✓', 'Check Mark'], ['✗', 'Ballot X'], ['★', 'Star'],
];
const RECENT_KEY = 'quilllattice.recentSymbols';
function recentSymbols(): string[] { try { const v = JSON.parse(localStorage.getItem(RECENT_KEY) || '[]'); return Array.isArray(v) ? v.filter((x) => typeof x === 'string').slice(0, 20) : []; } catch { return []; } }
function rememberSymbol(ch: string): void { try { const r = [ch, ...recentSymbols().filter((x) => x !== ch)].slice(0, 20); localStorage.setItem(RECENT_KEY, JSON.stringify(r)); } catch { /* storage off */ } }

/** Inserts text at the caret of the cell editor (starting an edit of the active cell when needed). */
export function insertTextIntoCell(app: App, text: string): void {
  const g = app.grid;
  if (!g) return;
  if (!g.editing) {
    if (app.isReadOnly()) { toast('This workbook is read-only.'); return; }
    g.startEdit({ mode: 'edit' });
    if (!g.editing) return;
  }
  const inp = g.editor.input;
  const s = inp.selectionStart ?? inp.value.length, e = inp.selectionEnd ?? inp.value.length;
  inp.setRangeText(text, s, e, 'end');
  inp.dispatchEvent(new Event('input', { bubbles: true }));
  inp.focus();
}

const EMOJI = '😀😃😄😁😆😅🤣😂🙂😉😊😇🥰😍🤩😘😋😛😜🤪🤗🤔🤨😐😑😶😏😒🙄😬😌😔😴😷🤒🤕🥳😎🤓🧐😕😟🙁😮😯😲😳🥺😢😭😱😖😞😓😩😫😤😡👍👎👏🙌🤝🙏💪👀👋✋👌✌️🤞👈👉👆👇❤️🧡💛💚💙💜🖤🤍💔💯✨⭐🌟🔥✅❌⚠️❓❗💡📌📎📝📚📖🔖📅📆⏰⌛📈📉📊💼📁📂🗂️🖊️✏️🎉🎊🎁🎈🏆🥇🥈🥉⚽🏀🎵🎧🎬📷💻🖥️📱☎️📧✉️🏠🏢🏫🏥🌍🌎🌏🌱🌳🌵🌷🌹🌻🍀🍁☀️⛅🌧️❄️🌈☔🍎🍊🍋🍌🍉🍇🍓🍒🍕🍔🍟☕🍵🐶🐱🐭🐰🦊🐻🐼🐨🐯🦁🐮🐷🐸🐵🐔🐧🐦🦉🐝🦋🐢🐙🐬🐳🚗🚕🚌🚑🚒✈️🚀⚓🛠️⚙️🔧🔨🔑🔒🔓🛡️⚖️🧪🔬🔭💊';

// --------------------------------------------------------------------------------------------
// Equations

const EQUATIONS: [string, string][] = [
  ['Area of a circle', 'A = \\pi r^2'], ['Pythagorean theorem', 'a^2 + b^2 = c^2'], ['Quadratic formula', 'x = (-b \\pm \\sqrt(b^2 - 4ac)) / 2a'],
  ['Binomial theorem', '(x + a)^n = \\sum_(k=0)^n (n k) x^k a^(n-k)'], ['Compound interest', 'A = P(1 + r/n)^(nt)'], ['Standard deviation', '\\sigma = \\sqrt((1/N) \\sum (x_i - \\mu)^2)'],
  ['Euler\'s identity', 'e^(i\\pi) + 1 = 0'], ['Mean', 'x̄ = (1/n) \\sum_(i=1)^n x_i'], ['Present value', 'PV = FV / (1 + r)^n'], ['Slope', 'm = (y_2 - y_1) / (x_2 - x_1)'],
];

async function equationDialog(app: App, arg: any): Promise<void> {
  const o = arg && typeof arg === 'object' && arg.id ? findObject(app, arg.id) : null;
  const cur = o?.kind === 'drawing' ? o.def : null;
  const r = await openDialog({
    title: cur ? 'Edit Equation' : 'Insert Equation', wide: true, preview: true, previewStyle: 'min-height:80px;display:flex;align-items:center;justify-content:center;background:#fff',
    help: 'Type in linear format: x^2 for superscripts, x_1 for subscripts, (…) to group, and \\alpha, \\sum, \\sqrt, \\pi, \\le, \\infty and similar words for symbols.',
    fields: [
      { key: 'tpl', label: 'Built-in equations', type: 'select', options: [{ value: '', label: '(choose one to start from)' }, ...EQUATIONS.map(([n, e]) => ({ value: e, label: n }))], value: '' },
      { key: 'eq', label: 'Equation', type: 'textarea', rows: 3, value: cur?.text || '', autofocus: true },
      { key: 'size', label: 'Font size', type: 'select', options: ['12', '14', '16', '18', '20', '24', '28', '36'], value: String(cur?.font?.size || 18), half: true },
      { key: 'color', label: 'Color', type: 'color', value: cur?.font?.color || '#000000', half: true },
    ],
    onChange: (v, ctl, key) => {
      if (key === 'tpl' && v.tpl) ctl.set('eq', v.tpl);
      const eq = key === 'tpl' && v.tpl ? v.tpl : v.eq;
      const tmp: DrawingObject = { id: 'eqpv', kind: 'textbox', shape: 'equation', at: { r: 0, c: 0, dx: 0, dy: 0, w: 420, h: 70 }, text: eq, font: { size: Number(v.size), color: v.color }, fill: 'none', stroke: 'none' };
      ctl.setPreview(renderShapeSvg(tmp));
    },
    validate: (v) => (String(v.eq || '').trim() ? null : 'Type an equation.'),
  });
  if (!r) return;
  const text = String(r.eq).trim();
  const len = equationRuns(text).reduce((a, x) => a + x.t.length * (x.k === 'n' ? 1 : 0.65), 0);
  const size = Number(r.size) || 18;
  const w = Math.max(80, Math.min(900, len * size * 0.75 + 30)), hh = Math.max(36, size * 2.6);
  if (cur && o) { mutateObject(app, o, 'Edit Equation', (d: DrawingObject) => { d.text = text; d.font = { ...(d.font || {}), size, color: r.color }; d.at = { ...d.at, w, h: hh }; }); return; }
  addDrawing(app, { kind: 'textbox', shape: 'equation', text, font: { size, color: r.color }, fill: 'none', stroke: 'none', at: { ...viewAnchor(app, w, hh), w, h: hh }, alt: `Equation: ${text}` }, 'Insert Equation');
}

// --------------------------------------------------------------------------------------------
// Diagrams

async function diagramDialog(app: App, arg: any): Promise<void> {
  const o = arg && typeof arg === 'object' && arg.id ? findObject(app, arg.id) : null;
  const cur = o?.kind === 'drawing' && o.def.shape?.startsWith('diagram:') ? o.def : null;
  const colors = [0, 1, 2, 3, 4, 5].map((i) => accent(app, i));
  const kind0 = cur ? cur.shape!.slice(8) : 'process';
  const sample: Record<string, string> = { hierarchy: 'Director\n  Manager A\n    Team 1\n    Team 2\n  Manager B\n    Team 3', timeline: 'Q1: Plan\nQ2: Build\nQ3: Test\nQ4: Launch', radial: 'Goal\nPeople\nProcess\nTools\nData', matrix: 'Strengths\nWeaknesses\nOpportunities\nThreats', venn: 'Skills\nPassion\nMarket' };
  const r = await openDialog({
    title: cur ? 'Edit Diagram' : 'Insert Diagram', wide: true, preview: true, previewStyle: 'display:flex;justify-content:center;background:#fff;overflow:auto',
    fields: [
      { key: 'kind', label: 'Layout', type: 'select', options: DIAGRAMS.map((d) => ({ value: d.kind, label: `${d.group}: ${d.label}` })), value: kind0 },
      { key: 'text', label: 'Text (one item per line)', type: 'textarea', rows: 7, value: cur?.text || sample[kind0] || 'Plan\nDo\nCheck\nAct', autofocus: true },
      { key: 'help', type: 'info', label: '' },
      { key: 'tc', label: 'Text color', type: 'color', value: cur?.font?.color || '#ffffff', half: true },
    ],
    onChange: (v, ctl, key) => {
      if (key === 'kind' && !cur && sample[v.kind] && Object.values(sample).concat('Plan\nDo\nCheck\nAct').includes(v.text)) ctl.set('text', sample[v.kind]);
      const t = ctl.get('text');
      const d = DIAGRAMS.find((x) => x.kind === v.kind);
      const info = ctl.body.querySelector('.ql-info');
      if (info) info.textContent = d?.help || '';
      ctl.setPreview(diagramSvg(v.kind, t, 420, 230, colors, v.tc));
    },
    validate: (v) => (String(v.text || '').trim() ? null : 'Type at least one item.'),
  });
  if (!r) return;
  if (cur && o) { mutateObject(app, o, 'Edit Diagram', (d: DrawingObject) => { d.shape = 'diagram:' + r.kind; d.text = r.text; d.font = { ...(d.font || {}), color: r.tc }; }); return; }
  addDrawing(app, { kind: 'shape', shape: 'diagram:' + r.kind, text: r.text, font: { color: r.tc }, fx: { colors }, at: { ...viewAnchor(app, 480, 260), w: 480, h: 260 }, alt: `Diagram: ${String(r.text).split('\n').map((x: string) => x.trim()).filter(Boolean).join(', ')}` }, 'Insert Diagram');
}

// --------------------------------------------------------------------------------------------
// Links

interface LinkParts { kind: 'web' | 'place' | 'email'; url: string; sheet: string; cell: string; email: string; subject: string }

export function parseLink(link: string | undefined, home: string): LinkParts {
  const out: LinkParts = { kind: 'web', url: '', sheet: home, cell: 'A1', email: '', subject: '' };
  if (!link) return out;
  if (link.startsWith('#')) {
    out.kind = 'place';
    const body = link.slice(1);
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(body);
    if (m) { out.sheet = (m[1] ?? m[2]).replace(/''/g, "'"); out.cell = m[3]; } else out.cell = body;
    return out;
  }
  if (/^mailto:/i.test(link)) {
    out.kind = 'email';
    const [addr, q] = link.slice(7).split('?');
    out.email = decodeURIComponent(addr || '');
    const sm = /(?:^|&)subject=([^&]*)/i.exec(q || '');
    if (sm) out.subject = decodeURIComponent(sm[1]);
    return out;
  }
  out.url = link;
  return out;
}

export function buildLink(p: LinkParts): string | null {
  if (p.kind === 'place') { const cell = p.cell.trim() || 'A1'; return '#' + (parseCell(cell.replace(/\$/g, '').split(':')[0]) ? `${quoteSheet(p.sheet)}!${cell}` : cell); }
  if (p.kind === 'email') { if (!p.email.trim()) return null; return 'mailto:' + p.email.trim() + (p.subject ? '?subject=' + encodeURIComponent(p.subject) : ''); }
  let u = p.url.trim();
  if (!u) return null;
  if (!/^[a-z][a-z0-9+.-]*:/i.test(u)) u = /^[\w-]+(\.[\w-]+)+/.test(u) ? 'https://' + u : u;
  return u;
}

async function linkDialog(app: App): Promise<void> {
  const s = app.sheet;
  const { r, c } = s.sel;
  const cell = s.get(r, c);
  const cur = parseLink(cell?.link, s.name);
  const names = app.book.names.filter((n) => !n.hidden).map((n) => n.name);
  const res = await openDialog({
    title: cell?.link ? 'Edit Hyperlink' : 'Insert Hyperlink',
    fields: [
      { key: 'kind', label: 'Link to', type: 'radio', options: [{ value: 'web', label: 'Existing file or web page' }, { value: 'place', label: 'Place in this workbook' }, { value: 'email', label: 'Email address' }], value: cur.kind },
      { key: 'text', label: 'Text to display', type: 'text', value: cellText(s, r, c) },
      { key: 'url', label: 'Address', type: 'text', value: cur.url, placeholder: 'https://example.com', visibleIf: (v) => v.kind === 'web', autofocus: true },
      { key: 'sheet', label: 'Sheet', type: 'select', options: [...app.book.sheets.filter((x) => x.visibility !== 'veryHidden').map((x) => x.name), ...(names.length ? ['(Defined name)'] : [])], value: cur.sheet, visibleIf: (v) => v.kind === 'place', half: true },
      { key: 'cell', label: 'Cell reference or defined name', type: 'combo', value: cur.cell, options: names, visibleIf: (v) => v.kind === 'place', half: true },
      { key: 'email', label: 'Email address', type: 'text', value: cur.email, placeholder: 'name@example.com', visibleIf: (v) => v.kind === 'email' },
      { key: 'subject', label: 'Subject', type: 'text', value: cur.subject, visibleIf: (v) => v.kind === 'email' },
      { key: 'tip', label: 'ScreenTip (shown when pointing at the link)', type: 'text', value: cell?.linkTip || '' },
    ],
    buttons: cell?.link ? [{ label: 'Remove Link', onClick: (_v, ctl) => ctl.close({ __remove: true }) }] : [],
    validate: (v) => {
      const link = buildLink({ ...cur, kind: v.kind, url: v.url || '', sheet: v.sheet === '(Defined name)' ? s.name : v.sheet, cell: v.cell || '', email: v.email || '', subject: v.subject || '' });
      if (!link) return v.kind === 'email' ? 'Enter an email address.' : 'Enter an address.';
      if (isScriptUrl(link)) return 'That address isn\'t allowed.';
      if (v.kind === 'place' && v.sheet !== '(Defined name)' && !parseCell(String(v.cell).replace(/\$/g, '').split(':')[0]) && !names.some((n) => n.toLowerCase() === String(v.cell).toLowerCase())) return 'Enter a cell reference such as A1, or a defined name.';
      return null;
    },
  });
  if (!res) return;
  if (res.__remove) { app.exec('insert.removeLink'); return; }
  const link = buildLink({ ...cur, kind: res.kind, url: res.url || '', sheet: res.sheet === '(Defined name)' ? s.name : res.sheet, cell: res.cell || '', email: res.email || '', subject: res.subject || '' })!;
  const display = String(res.text ?? '');
  const linkStyle = { color: '#0563c1', underline: 'single' as const };
  app.book.transact(cell?.link ? 'Edit Hyperlink' : 'Insert Hyperlink', () => {
    const rg = clipToUsed(s, activeRange(app));
    const single = rg.r1 === rg.r2 && rg.c1 === rg.c2;
    for (let rr = rg.r1; rr <= Math.min(rg.r2, rg.r1 + 9999); rr++) for (let cc = rg.c1; cc <= rg.c2; cc++) {
      const cur2 = s.get(rr, cc);
      if (s.protection.on && s.styleOf(rr, cc).locked !== false) continue;
      const next: Cell = { ...(cur2 || {}), link };
      if (res.tip) next.linkTip = String(res.tip); else delete next.linkTip;
      if ((single || cur2?.v == null) && cur2?.f == null) {
        const text = display || (cur2?.v != null && !single ? String(cur2.v) : link.replace(/^mailto:/i, '').replace(/^#/, ''));
        if (single || cur2?.v == null) { next.v = text; delete next.rich; }
      }
      next.s = app.book.derivedStyle(cur2?.s ?? 0, linkStyle);
      s.set(rr, cc, next);
    }
  });
}

// --------------------------------------------------------------------------------------------
// Sparklines

function sparklinesInSelection(app: App): Sparkline[] {
  const s = app.sheet;
  return s.sparklines.filter((sp) => selRanges(app).some((rg) => sp.cell.r >= rg.r1 && sp.cell.r <= rg.r2 && sp.cell.c >= rg.c1 && sp.cell.c <= rg.c2));
}

/** One sparkline per location cell: each gets the matching row (or column) of the data range. */
export function buildSparklines(sheet: Sheet, data: { sheet: Sheet; range: Range }, loc: Range, type: Sparkline['type']): Sparkline[] {
  const cells: { r: number; c: number }[] = [];
  for (let r = loc.r1; r <= loc.r2; r++) for (let c = loc.c1; c <= loc.c2; c++) cells.push({ r, c });
  const d = data.range;
  const rows = d.r2 - d.r1 + 1, cols = d.c2 - d.c1 + 1;
  const ref = (rg: Range) => rangeRefText(data.sheet, rg, true, sheet);
  if (cells.length === 1) return [{ id: newId('sp'), cell: cells[0], data: ref(d), type }];
  const vertical = loc.c1 === loc.c2;
  if (vertical && cells.length === rows) return cells.map((cell, i) => ({ id: newId('sp'), cell, data: ref({ r1: d.r1 + i, c1: d.c1, r2: d.r1 + i, c2: d.c2 }), type }));
  if (!vertical && cells.length === cols) return cells.map((cell, i) => ({ id: newId('sp'), cell, data: ref({ r1: d.r1, c1: d.c1 + i, r2: d.r2, c2: d.c1 + i }), type }));
  if (vertical && cells.length === cols) return cells.map((cell, i) => ({ id: newId('sp'), cell, data: ref({ r1: d.r1, c1: d.c1 + i, r2: d.r2, c2: d.c1 + i }), type }));
  if (!vertical && cells.length === rows) return cells.map((cell, i) => ({ id: newId('sp'), cell, data: ref({ r1: d.r1 + i, c1: d.c1, r2: d.r1 + i, c2: d.c2 }), type }));
  throw new Error('The location range must have one cell for each row or column of the data range.');
}

async function sparklineDialog(app: App, type: Sparkline['type']): Promise<void> {
  const s = app.sheet;
  if (s.protection.on) { toast('The sheet is protected.'); return; }
  const sel = clipToUsed(s, activeRange(app));
  // Data = the selection when it holds numbers; location = the empty column to its right.
  let hasNums = false;
  s.forEachIn(sel, (_r, _c, cell) => { if (typeof cell.v === 'number') hasNums = true; });
  const dataGuess = hasNums ? rangeName(sel) : '';
  const locGuess = hasNums ? rangeName({ r1: sel.r1, c1: sel.c2 + 1, r2: sel.r2, c2: sel.c2 + 1 }) : rangeName(sel);
  const label = type === 'line' ? 'Line' : type === 'column' ? 'Column' : 'Win/Loss';
  const r = await openDialog({
    title: `Create ${label} Sparklines`,
    fields: [
      { key: 'data', label: 'Data range', type: 'text', value: dataGuess, placeholder: 'B2:M10', autofocus: !dataGuess },
      { key: 'loc', label: 'Location range', type: 'text', value: locGuess, placeholder: 'N2:N10' },
    ],
    validate: (v) => {
      if (!resolveRangeText(app.book, String(v.data).replace(/^=/, '').replace(/\$/g, ''), s)) return 'The data range isn\'t valid.';
      const loc = resolveRangeText(app.book, String(v.loc).replace(/^=/, '').replace(/\$/g, ''), s);
      if (!loc || loc.sheet !== s) return 'The location must be a range on this sheet.';
      if (loc.range.r1 !== loc.range.r2 && loc.range.c1 !== loc.range.c2) return 'The location must be a single row or column.';
      return null;
    },
  });
  if (!r) return;
  const data = resolveRangeText(app.book, String(r.data).replace(/^=/, '').replace(/\$/g, ''), s)!;
  const loc = resolveRangeText(app.book, String(r.loc).replace(/^=/, '').replace(/\$/g, ''), s)!.range;
  let list: Sparkline[];
  try { list = buildSparklines(s, data, loc, type); } catch (e) { toast((e as Error).message); return; }
  const color = accent(app, 0);
  app.book.transact(`Insert ${label} Sparklines`, () => {
    app.book.recordSheet(s);
    s.sparklines = s.sparklines.filter((sp) => !list.some((n) => n.cell.r === sp.cell.r && n.cell.c === sp.cell.c));
    for (const sp of list) s.sparklines.push({ ...sp, color, ...(type === 'winloss' ? { negative: true } : {}) });
  });
}

// --------------------------------------------------------------------------------------------
// Slicers and timelines

interface SlicerState { tableId: string; table: string; column: string; items: string[]; selected: string[]; timeline?: boolean; months?: number[]; from?: number; to?: number; cols?: number }

function columnItems(sheet: Sheet, t: import('../model/types').TableDef, col: number): string[] {
  const d = tableDataRows(t);
  const c = t.range.c1 + col;
  const seen = new Map<string, number | string>();
  for (let r = d.r1; r <= d.r2; r++) { const txt = cellText(sheet, r, c); if (!seen.has(txt)) seen.set(txt, (sheet.value(r, c) as any) ?? ''); }
  return [...seen.entries()].sort((a, b) => (typeof a[1] === 'number' && typeof b[1] === 'number' ? a[1] - b[1] : String(a[0]).localeCompare(String(b[0]), undefined, { numeric: true }))).map(([k]) => k);
}

function monthStarts(sheet: Sheet, t: import('../model/types').TableDef, col: number): number[] {
  const d = tableDataRows(t);
  const c = t.range.c1 + col;
  const set = new Set<number>();
  const d1904 = sheet.book.calc.date1904;
  for (let r = d.r1; r <= d.r2; r++) {
    const v = sheet.value(r, c);
    if (typeof v !== 'number') continue;
    const dt = serialToDate(v, d1904);
    set.add(dateToSerial(new Date(Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth(), 1)), d1904));
  }
  return [...set].sort((a, b) => a - b);
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function monthLabel(serial: number, d1904: boolean): string { const d = serialToDate(serial, d1904); return `${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}`; }

function findTableForSlicer(app: App, st: Partial<SlicerState>): { sheet: Sheet; t: import('../model/types').TableDef } | null {
  for (const sh of app.book.sheets) for (const t of sh.tables) if (t.id === st.tableId || t.name === st.table) return { sheet: sh, t };
  return null;
}

/** Applies a slicer's selection to its table's filter and hides/shows rows. */
export function applySlicer(sheet: Sheet, t: import('../model/types').TableDef, col: number, st: SlicerState): void {
  sheet.book.recordSheet(sheet);
  const cols = { ...(t.filter?.columns || {}) };
  let crit: FilterCriteria | null = null;
  if (st.timeline && st.from != null && st.to != null && st.months?.length) {
    const d1904 = sheet.book.calc.date1904;
    const end = serialToDate(st.to, d1904);
    const next = dateToSerial(new Date(Date.UTC(end.getUTCFullYear(), end.getUTCMonth() + 1, 1)), d1904);
    const allSel = st.from <= st.months[0] && st.to >= st.months[st.months.length - 1];
    if (!allSel) crit = { custom: [{ op: 'ge', value: String(st.from) }, { op: 'lt', value: String(next) }], and: true };
  } else if (st.selected.length && st.selected.length < st.items.length) {
    crit = { values: st.selected.filter((x) => x !== ''), blanks: st.selected.includes('') };
  }
  if (crit) cols[col] = crit; else delete cols[col];
  t.filter = Object.keys(cols).length || t.filter?.sort ? { ...(t.filter || {}), columns: cols } : undefined;
  if (!t.filter) for (let r = t.range.r1; r <= t.range.r2; r++) sheet.hiddenRows.delete(r);
  applyFilters(sheet);
}

async function insertSlicer(app: App, timeline: boolean): Promise<void> {
  const s = app.sheet;
  if (!canEditObjects(app)) return;
  const t = tableAt(s, s.sel.r, s.sel.c) || (s.tables.length === 1 ? s.tables[0] : null);
  if (!t) { toast(timeline ? 'Select a cell in a table that has a date column, then insert a timeline.' : 'Select a cell in a table first (Insert ▸ Table), then insert a slicer.'); return; }
  const cols = t.columns.map((c, i) => ({ i, name: c.name }));
  const d = tableDataRows(t);
  const dateCols = cols.filter(({ i }) => { const st = s.styleOf(d.r1, t.range.c1 + i); return typeof s.value(d.r1, t.range.c1 + i) === 'number' && isDateFormat(st.numFmt); });
  if (timeline && !dateCols.length) { toast('This table has no date column for a timeline.'); return; }
  const r = await openDialog({
    title: timeline ? 'Insert Timeline' : 'Insert Slicers',
    help: timeline ? 'Choose the date column to filter by period.' : 'Choose the columns to filter with buttons. One slicer is created for each.',
    fields: [{ key: 'cols', label: t.name, type: 'list', multiple: !timeline, rows: Math.min(10, Math.max(4, cols.length)), options: (timeline ? dateCols : cols).map((c) => ({ value: String(c.i), label: c.name })), value: timeline ? String(dateCols[0].i) : [] }],
    validate: (v) => (!v.cols || (Array.isArray(v.cols) && !v.cols.length) ? 'Choose a column.' : null),
  });
  if (!r) return;
  const chosen = (Array.isArray(r.cols) ? r.cols : [r.cols]).map(Number);
  const d1904 = app.book.calc.date1904;
  let first: string | null = null;
  app.book.transact(timeline ? 'Insert Timeline' : 'Insert Slicer', () => {
    app.book.recordSheet(s);
    chosen.forEach((ci, k) => {
      const name = t.columns[ci].name;
      const st: SlicerState = { tableId: t.id, table: t.name, column: name, items: [], selected: [] };
      let w = 190, hh: number;
      if (timeline) {
        st.timeline = true;
        st.months = monthStarts(s, t, ci);
        st.items = st.months.map((m) => monthLabel(m, d1904));
        st.from = st.months[0]; st.to = st.months[st.months.length - 1];
        st.selected = [...st.items];
        st.cols = Math.min(6, Math.max(1, st.items.length));
        w = Math.max(260, st.cols * 82 + 16);
        hh = 30 + 8 + Math.ceil(st.items.length / st.cols) * 30 + 8;
      } else {
        st.items = columnItems(s, t, ci);
        hh = Math.min(300, 30 + 8 + st.items.length * 30 + 6);
      }
      const at = { r: t.range.r1 + k * 2, c: t.range.c2 + 2 + k * 3, dx: 0, dy: 0, w, h: Math.max(80, hh) };
      const id = newId('slc');
      first = first || id;
      s.drawings.push({ id, kind: 'shape', shape: 'slicer', text: timeline ? `${name} (timeline)` : name, fill: accent(app, 0), at, fx: { ...st }, alt: `${timeline ? 'Timeline' : 'Slicer'} for ${t.name}[${name}]` });
    });
  });
  if (first) app.grid?.objects.select(first);
  toast('Double-click a slicer to choose what it shows.');
}

async function slicerPick(app: App, arg: any): Promise<void> {
  const o = findObject(app, arg?.id);
  if (!o || o.kind !== 'drawing' || o.def.shape !== 'slicer') { toast('Select a slicer first.'); return; }
  const st = { ...(o.def.fx || {}) } as SlicerState;
  const found = findTableForSlicer(app, st);
  if (!found) { toast('The table this slicer filters no longer exists.'); return; }
  const { sheet, t } = found;
  const ci = t.columns.findIndex((c) => c.name === st.column);
  if (ci < 0) { toast(`The column "${st.column}" is no longer in ${t.name}.`); return; }
  const d1904 = app.book.calc.date1904;
  if (st.timeline) {
    st.months = monthStarts(sheet, t, ci);
    st.items = st.months.map((m) => monthLabel(m, d1904));
    if (!st.months.length) { toast('The date column has no dates.'); return; }
    const opts = st.months.map((m, i) => ({ value: String(m), label: st.items[i] }));
    const r = await openDialog({
      title: `Timeline: ${st.column}`,
      fields: [
        { key: 'from', label: 'From', type: 'select', options: opts, value: String(st.from ?? st.months[0]), half: true },
        { key: 'to', label: 'To', type: 'select', options: opts, value: String(st.to ?? st.months[st.months.length - 1]), half: true },
      ],
      buttons: [{ label: 'Clear Filter', onClick: (_v, ctl) => ctl.close({ from: String(st.months![0]), to: String(st.months![st.months!.length - 1]) }) }],
    });
    if (!r) return;
    let from = Number(r.from), to = Number(r.to);
    if (from > to) [from, to] = [to, from];
    st.from = from; st.to = to;
    st.selected = st.months.filter((m) => m >= from && m <= to).map((m) => monthLabel(m, d1904));
  } else {
    st.items = columnItems(sheet, t, ci);
    const sel = st.selected?.length ? st.selected.filter((x) => st.items.includes(x)) : [...st.items];
    const r = await openDialog({
      title: `Slicer: ${st.column}`, help: 'Choose the values to show (Ctrl or Shift for several).',
      fields: [{ key: 'v', label: '', type: 'list', multiple: true, rows: Math.min(14, Math.max(5, st.items.length)), options: st.items.map((x) => ({ value: x, label: x === '' ? '(blank)' : x })), value: sel }],
      buttons: [{ label: 'Clear Filter', onClick: (_v, ctl) => ctl.close({ v: [...st.items] }) }],
    });
    if (!r) return;
    const v: string[] = Array.isArray(r.v) ? r.v : [];
    st.selected = v.length === st.items.length || !v.length ? [] : v;
  }
  app.book.transact('Slicer', () => {
    app.book.recordSheet(o.sheet);
    o.def.fx = { ...st };
    applySlicer(sheet, t, ci, st);
  });
}

// --------------------------------------------------------------------------------------------
// Checkboxes

function toggleCheckbox(app: App, arg: { r: number; c: number }): boolean {
  const s = app.sheet;
  const r = arg?.r ?? s.sel.r, c = arg?.c ?? s.sel.c;
  const st = s.styleOf(r, c);
  if (!st.checkbox) return false;
  if (s.protection.on && st.locked !== false) { app.exec('review.protectionNotice'); return false; }
  const cell = s.get(r, c);
  if (cell?.f != null) { toast('This checkbox is set by a formula.'); return false; }
  const v = cell?.v;
  app.book.transact('Toggle Checkbox', () => s.set(r, c, { ...(cell || {}), v: !(v === true) }));
  return true;
}

// --------------------------------------------------------------------------------------------
// Commands

const SHAPE_SIZE = (id: string): [number, number] => (/^(line|straightConnector1|doubleArrowLine)$/.test(id) ? [160, 24] : /Connector$/.test(id) ? [160, 80] : /^(left|right)(Brace|Bracket)$/.test(id) ? [30, 120] : [140, 100]);

function insertShape(app: App, id: string): DrawingObject | null {
  const [w, hh] = SHAPE_SIZE(id);
  const line = /^(line|straightConnector1|doubleArrowLine|elbowConnector|curvedConnector|arc|bracketPair|bracePair|leftBracket|rightBracket|leftBrace|rightBrace)$/.test(id);
  return addDrawing(app, { kind: 'shape', shape: id, at: { ...viewAnchor(app, w, hh), w, h: hh }, fill: line ? 'none' : accent(app, 0), stroke: line ? accent(app, 0) : app.book.theme.colors[2] || '#1f3b57', strokeWidth: line ? 2 : 1.5 }, 'Insert Shape');
}

def([
  { id: 'insert.pivotTable', label: 'PivotTable', icon: 'pivot_table_chart', tab: 'Insert', group: 'Tables', edits: true, popular: true, keywords: 'pivot summarize', run: (app, arg) => app.exec('pivot.create', arg) },
  { id: 'insert.recommendedPivot', label: 'Recommended PivotTables', icon: 'auto_awesome', tab: 'Insert', group: 'Tables', edits: true, keywords: 'pivot suggest', run: (app, arg) => app.exec('pivot.recommended', arg) },
  {
    id: 'insert.picture', label: 'Pictures', icon: 'image', tab: 'Insert', group: 'Illustrations', kind: 'split', edits: true, popular: true, keywords: 'image photo picture insert',
    run: async (app, arg) => {
      if (arg === 'screenshot') return app.exec('insert.screenshot');
      if (!canEditObjects(app)) return;
      const files = await pickPictures();
      for (const f of files) await insertPictureFile(app, f, f.name);
    },
    menu: () => [{ id: 'insert.picture', label: 'This Device…', icon: 'folder_open' }, { id: 'insert.picture', arg: 'screenshot', label: 'Screen Clipping…', icon: 'screenshot_monitor' }],
  },
  { id: 'insert.pictureFile', label: 'Insert Picture File', icon: 'image', tab: 'Insert', group: 'Illustrations', internal: true, edits: true, run: (app, file: Blob) => (file instanceof Blob ? insertPictureFile(app, file, (file as File).name || 'Picture') : null) },
  {
    id: 'insert.screenshot', label: 'Screenshot', icon: 'screenshot_monitor', tab: 'Insert', group: 'Illustrations', edits: true, keywords: 'screen capture clipping',
    run: async (app) => {
      if (!canEditObjects(app)) return;
      const blob = (navigator as any).mediaDevices?.getDisplayMedia ? await captureScreen() : null;
      if (blob) { await insertPictureFile(app, blob, 'Screenshot'); return; }
      toast('Choose a screenshot picture to insert.');
      const files = await pickPictures();
      for (const f of files) await insertPictureFile(app, f, f.name);
    },
  },
  {
    id: 'insert.shapes', label: 'Shapes', icon: 'category', tab: 'Insert', group: 'Illustrations', kind: 'split', edits: true, popular: true, keywords: 'shape rectangle oval arrow star callout flowchart line',
    run: async (app, arg) => {
      if (!canEditObjects(app)) return;
      const id = typeof arg === 'string' ? arg : await galleryDialog('Shapes', SHAPES.map((s) => ({ value: s.id, label: s.label, group: s.group, html: shapePreviewSvg(s.id, 34) })), { size: 'small', search: true, ok: 'Insert' });
      if (id) insertShape(app, id);
    },
    menu: () => [
      ...SHAPE_GROUPS.map((g) => ({ label: g, icon: 'category', submenu: SHAPES.filter((s) => s.group === g).map((s) => ({ id: 'insert.shapes', arg: s.id, label: s.label })) })),
      { separator: true }, { id: 'insert.shapes', label: 'All Shapes…', icon: 'apps' },
    ],
  },
  {
    id: 'insert.icons', label: 'Icons', icon: 'emoji_symbols', tab: 'Insert', group: 'Illustrations', edits: true, keywords: 'icon symbol pictogram',
    run: async (app, arg) => {
      const existing = arg && typeof arg === 'object' && arg.id ? selectedDrawing(app, arg) : null;
      const name = typeof arg === 'string' ? arg : await galleryDialog(existing ? 'Change Icon' : 'Insert Icons', ICONS.flatMap(([g, list]) => list.map((n) => ({ value: n, label: n.replace(/_/g, ' '), group: g, html: `<span class="ql-icon ql-objgal-icon">${escHtml(n)}</span>` }))), { size: 'small', search: true, ok: existing ? 'OK' : 'Insert', current: existing?.text });
      if (!name) return;
      if (existing) { const o = findObject(app, existing.id)!; mutateObject(app, o, 'Change Icon', (d: DrawingObject) => { d.text = name; d.alt = name.replace(/_/g, ' '); }); return; }
      addDrawing(app, { kind: 'icon', shape: 'icon', text: name, fill: accent(app, 0), at: { ...viewAnchor(app, 96, 96), w: 96, h: 96 }, lockAspect: true, alt: name.replace(/_/g, ' ') }, 'Insert Icon');
    },
  },
  { id: 'insert.diagram', label: 'Diagram', icon: 'account_tree', tab: 'Insert', group: 'Illustrations', edits: true, keywords: 'diagram process cycle hierarchy org chart list matrix pyramid', run: (app, arg) => diagramDialog(app, arg) },
  {
    id: 'insert.textBox', label: 'Text Box', icon: 'text_fields', tab: 'Insert', group: 'Text', edits: true, popular: true, keywords: 'text box label note',
    run: (app) => { const d = addDrawing(app, { kind: 'textbox', shape: 'rect', text: '', at: { ...viewAnchor(app, 200, 80), w: 200, h: 80 }, fill: '#ffffff', stroke: '#7f7f7f', strokeWidth: 0.75, font: { size: 11 } }, 'Insert Text Box'); if (d) app.exec('drawing.editText', { id: d.id }); },
  },
  {
    id: 'insert.wordArt', label: 'Decorative Text', icon: 'format_shapes', tab: 'Insert', group: 'Text', kind: 'split', edits: true, keywords: 'wordart fancy styled text title',
    run: async (app, arg) => {
      if (!canEditObjects(app)) return;
      const sample = (id: string) => { const st = WORDART_STYLES.find((x) => x.id === id)!; return renderShapeSvg({ id: 'pv' + id, kind: 'wordart', text: 'Abc', at: { r: 0, c: 0, dx: 0, dy: 0, w: 110, h: 54 }, font: { color: st.fill, bold: st.bold !== false, ...(st.font ? { font: st.font } : {}) }, fx: { textOutline: st.outline, textOutlineW: st.outlineW, shadow: st.shadow, glow: st.glow } }); };
      const id = typeof arg === 'string' ? arg : await galleryDialog('Decorative Text', WORDART_STYLES.map((s) => ({ value: s.id, label: s.label, html: sample(s.id) })), { size: 'small', ok: 'Insert' });
      if (!id) return;
      const st = WORDART_STYLES.find((x) => x.id === id) || WORDART_STYLES[0];
      const d = addDrawing(app, { kind: 'wordart', text: 'Your text here', at: { ...viewAnchor(app, 360, 90), w: 360, h: 90 }, font: { color: st.fill, bold: st.bold !== false, ...(st.font ? { font: st.font } : {}) }, fx: Object.fromEntries(Object.entries({ textOutline: st.outline, textOutlineW: st.outlineW, shadow: st.shadow, glow: st.glow }).filter(([, v]) => v != null)) }, 'Insert Decorative Text');
      if (d) app.exec('drawing.editText', { id: d.id });
    },
    menu: () => WORDART_STYLES.map((s) => ({ id: 'insert.wordArt', arg: s.id, label: s.label, icon: 'format_shapes' })),
  },
  {
    id: 'insert.signatureLine', label: 'Signature Line', icon: 'signature', tab: 'Insert', group: 'Text', edits: true, keywords: 'sign signature',
    run: async (app) => {
      if (!canEditObjects(app)) return;
      const r = await openDialog({
        title: 'Signature Setup',
        fields: [
          { key: 'name', label: 'Suggested signer (for example, John Doe)', type: 'text', autofocus: true },
          { key: 'title', label: 'Suggested signer\'s title', type: 'text' },
          { key: 'email', label: 'Suggested signer\'s email address', type: 'text' },
          { key: 'instr', label: 'Instructions to the signer', type: 'textarea', rows: 2, value: 'Before signing this document, verify that the content you are signing is correct.' },
          { key: 'date', label: 'Show sign date in signature line', type: 'checkbox', value: true },
        ],
      });
      if (!r) return;
      const w = 320, hh = 120;
      const lines = [r.name, r.title, r.email].filter(Boolean).map((x: string) => escHtml(x));
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${hh}" viewBox="0 0 ${w} ${hh}" font-family="Carlito, Arial, sans-serif"><rect width="${w}" height="${hh}" fill="#ffffff"/><text x="14" y="56" font-size="34" fill="#404040">✕</text>${r.date ? `<text x="${w - 12}" y="24" font-size="12" text-anchor="end" fill="#7f7f7f">Date: ____________</text>` : ''}<line x1="12" y1="66" x2="${w - 12}" y2="66" stroke="#404040" stroke-width="1.5"/>${lines.map((l: string, i: number) => `<text x="14" y="${84 + i * 15}" font-size="12.5" fill="#262626">${l}</text>`).join('')}</svg>`;
      const png = await svgToPng(svg, w, hh, 2);
      const src = png ? await readFileAsDataURL(png) : 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
      addDrawing(app, { kind: 'signature', src, at: { ...viewAnchor(app, w, hh), w, h: hh }, alt: `Signature line${r.name ? ' for ' + r.name : ''}`, lockAspect: true, fx: { signer: r.name || '', signerTitle: r.title || '', signerEmail: r.email || '', instructions: r.instr || '' } }, 'Insert Signature Line');
    },
  },
  {
    id: 'insert.object', label: 'Object', icon: 'attach_file', tab: 'Insert', group: 'Text', edits: true, keywords: 'embed file attachment object',
    run: async (app) => {
      if (!canEditObjects(app)) return;
      const files = await pickFile('*/*');
      const f = files[0];
      if (!f) return;
      if (f.size > 15 * 1024 * 1024) { toast('Files larger than 15 MB can\'t be embedded.'); return; }
      const src = await readFileAsDataURL(f);
      const d = addDrawing(app, { kind: 'shape', shape: 'embed', text: f.name, src, fill: accent(app, 0), at: { ...cellAnchor(app), w: 96, h: 112 }, alt: `Embedded file: ${f.name}`, fx: { mime: f.type || 'application/octet-stream', size: f.size } }, 'Insert Object');
      if (d) toast('Double-click the icon to save a copy of the file.');
    },
  },
  { id: 'insert.equation', label: 'Equation', icon: 'functions', tab: 'Insert', group: 'Symbols', edits: true, keywords: 'math formula equation', run: (app, arg) => equationDialog(app, arg) },
  {
    id: 'insert.symbol', label: 'Symbol', icon: 'special_character', tab: 'Insert', group: 'Symbols', kind: 'split', edits: true, popular: true, keywords: 'symbol special character unicode',
    run: async (app, arg) => {
      if (typeof arg === 'string' && arg) { insertTextIntoCell(app, arg); rememberSymbol(arg); return; }
      const grid = h('div', { class: 'ql-symbol-grid', role: 'listbox', 'aria-label': 'Symbols' });
      const info = h('div', { class: 'ql-help' }, 'Choose a symbol.');
      let chosen = '';
      const pickBtn = (ch: string, title: string) => {
        const b = h('button', { class: 'ql-sym-btn', type: 'button', title }, ch === '\u00a0' ? '␣' : ch);
        b.addEventListener('click', () => { chosen = ch; info.textContent = title; grid.querySelectorAll('.ql-sel').forEach((x) => x.classList.remove('ql-sel')); b.classList.add('ql-sel'); });
        b.addEventListener('dblclick', () => { chosen = ch; (grid.closest('.ql-dialog')?.querySelector('.ql-primary') as HTMLElement | null)?.click(); });
        return b;
      };
      const fill = (key: string) => {
        grid.replaceChildren();
        if (key === 'recent') { const rs = recentSymbols(); if (!rs.length) grid.appendChild(h('div', { class: 'ql-help' }, 'No recently used symbols yet.')); rs.forEach((ch) => grid.appendChild(pickBtn(ch, `U+${ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}`))); return; }
        if (key === 'special') { SPECIAL_CHARS.forEach(([ch, n]) => grid.appendChild(pickBtn(ch, n))); return; }
        const b = SYMBOL_BLOCKS[Number(key)];
        for (let cp = b[1]; cp <= b[2]; cp++) { const ch = String.fromCodePoint(cp); grid.appendChild(pickBtn(ch, `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`)); }
      };
      fill(recentSymbols().length ? 'recent' : '0');
      const r = await openDialog({
        title: 'Symbol', wide: true, ok: 'Insert', content: h('div', null, grid, info),
        fields: [
          { key: 'set', label: 'Subset', type: 'select', options: [{ value: 'recent', label: 'Recently used symbols' }, { value: 'special', label: 'Special characters' }, ...SYMBOL_BLOCKS.map((b, i) => ({ value: String(i), label: b[0] }))], value: recentSymbols().length ? 'recent' : '0' },
          { key: 'code', label: 'Character code (hex), or type a character', type: 'text', placeholder: '00A9' },
        ],
        onChange: (v, _c, key) => { if (key === 'set') fill(v.set); },
      });
      if (!r) return;
      let ch = chosen;
      const code = String(r.code || '').trim();
      if (code) ch = /^(u\+)?[0-9a-f]{2,6}$/i.test(code) ? (() => { try { return String.fromCodePoint(parseInt(code.replace(/^u\+/i, ''), 16)); } catch { return ''; } })() : code;
      if (!ch) return;
      rememberSymbol(ch);
      insertTextIntoCell(app, ch);
    },
    menu: () => [
      ...(recentSymbols().length ? recentSymbols() : ['€', '£', '¥', '©', '®', '™', '±', '≠', '≤', '≥', '÷', '×', '∞', 'µ', 'α', 'β', 'π', 'Ω', '∑', '°']).slice(0, 20).map((c) => ({ id: 'insert.symbol', arg: c, label: c })),
      { separator: true }, { id: 'insert.symbol', label: 'More Symbols…', icon: 'special_character' }, { id: 'insert.emoji', label: 'Emoji…', icon: 'mood' },
    ],
  },
  {
    id: 'insert.emoji', label: 'Emoji', icon: 'mood', tab: 'Insert', group: 'Symbols', edits: true, keys: ['Mod-.'], keywords: 'emoji smiley face',
    run: (app, arg) => {
      if (typeof arg === 'string' && arg) { app.exec('insert.symbol', arg); return; }
      const list = typeof (Intl as any).Segmenter === 'function' ? [...new (Intl as any).Segmenter(undefined, { granularity: 'grapheme' }).segment(EMOJI)].map((s: any) => s.segment) : Array.from(EMOJI);
      showMenu(list.map((e: string) => ({ id: 'insert.emoji', arg: e, label: e })), { onPick: (e) => app.exec('insert.symbol', e.arg), anchor: app.ui.anchorFor('insert.emoji'), grid: true, className: 'ql-emoji-menu', title: 'Emoji' });
    },
  },
  {
    id: 'insert.checkbox', label: 'Checkbox', icon: 'check_box', tab: 'Insert', group: 'Symbols', kind: 'split', edits: true, keywords: 'checkbox tick check box boolean',
    run: (app, arg) => {
      const s = app.sheet;
      if (s.protection.on) { app.exec('review.protectionNotice'); return; }
      const remove = arg === 'remove';
      app.book.transact(remove ? 'Remove Checkboxes' : 'Insert Checkboxes', () => {
        for (const rg0 of selRanges(app)) {
          const rg = clipToUsed(s, rg0, remove ? 0 : 0);
          const rr2 = rg0.r1 === rg0.r2 ? rg0.r2 : Math.max(rg.r2, Math.min(rg0.r2, rg0.r1 + 999));
          const cc2 = rg0.c1 === rg0.c2 ? rg0.c2 : Math.max(rg.c2, Math.min(rg0.c2, rg0.c1 + 99));
          for (let r = rg0.r1; r <= rr2; r++) for (let c = rg0.c1; c <= cc2; c++) {
            const cur = s.get(r, c);
            const sid = app.book.derivedStyle(s.styleIdOf(r, c), { checkbox: remove ? undefined : true, hAlign: remove ? undefined : 'center' });
            const next: Cell = { ...(cur || {}), s: sid || undefined };
            if (!remove && (cur?.v == null || cur.v === '') && cur?.f == null) next.v = false;
            if (!next.s) delete next.s;
            s.set(r, c, next);
          }
        }
      });
    },
    menu: () => [{ id: 'insert.checkbox', label: 'Insert Checkboxes', icon: 'check_box' }, { id: 'insert.checkbox', arg: 'remove', label: 'Remove Checkboxes (keep values)', icon: 'check_box_outline_blank' }],
  },
  { id: 'cell.toggleCheckbox', label: 'Toggle Checkbox', icon: 'check_box', tab: 'Insert', group: 'Symbols', internal: true, edits: true, run: (app, arg) => toggleCheckbox(app, arg) },
  {
    id: 'insert.link', label: 'Link', icon: 'link', tab: 'Insert', group: 'Links', kind: 'split', keys: ['Mod-k'], edits: true, popular: true, keywords: 'hyperlink url web address email',
    run: (app) => linkDialog(app),
    menu: (app) => { const has = !!app.sheet.get(app.sheet.sel.r, app.sheet.sel.c)?.link; return [{ id: 'insert.link', label: has ? 'Edit Link…' : 'Insert Link…', icon: 'link' }, ...(has ? [{ id: 'insert.openLink', label: 'Open Link', icon: 'open_in_new' }, { id: 'insert.copyLink', label: 'Copy Link', icon: 'content_copy' }, { id: 'insert.removeLink', label: 'Remove Link', icon: 'link_off' }] : [])] as MenuEntry[]; },
  },
  {
    id: 'insert.openLink', label: 'Open Link', icon: 'open_in_new', tab: 'Insert', group: 'Links', noRecord: true,
    run: (app) => { const l = app.sheet.get(app.sheet.sel.r, app.sheet.sel.c)?.link; if (!l) { toast('The active cell has no link.'); return; } app.followLink(l); },
  },
  {
    id: 'insert.removeLink', label: 'Remove Link', icon: 'link_off', tab: 'Insert', group: 'Links', edits: true, keywords: 'remove hyperlink unlink',
    run: (app) => {
      const s = app.sheet;
      let n = 0;
      app.book.transact('Remove Hyperlinks', () => {
        for (const rg of selRanges(app)) s.forEachIn(clipToUsed(s, rg), (r, c, cell) => {
          if (!cell.link) return;
          const { link: _l, linkTip: _t, ...rest } = cell; void _l; void _t;
          const st = app.book.styles[cell.s ?? 0] || {};
          if (st.color === '#0563c1' && st.underline === 'single') rest.s = app.book.derivedStyle(cell.s ?? 0, { color: undefined, underline: undefined }) || undefined;
          if (!rest.s) delete rest.s;
          s.set(r, c, rest);
          n++;
        });
      });
      if (!n) toast('There are no links in the selection.');
    },
  },
  {
    id: 'insert.copyLink', label: 'Copy Link', icon: 'content_copy', tab: 'Insert', group: 'Links', noRecord: true,
    run: async (app) => {
      const l = app.sheet.get(app.sheet.sel.r, app.sheet.sel.c)?.link;
      if (!l) { toast('The active cell has no link.'); return; }
      const { writeClipboardText } = await import('../platform/native');
      const ok = await writeClipboardText(l).catch(() => false);
      toast(ok ? 'Link copied' : 'The link could not be copied.');
    },
  },
  {
    id: 'insert.linkToCell', label: 'Link to Cell', icon: 'link', tab: 'Insert', group: 'Links', internal: true, edits: true,
    run: (app, arg: { sheet?: string; cell: string }) => {
      const s = app.sheet;
      const link = '#' + (arg.sheet ? `${quoteSheet(arg.sheet)}!` : '') + arg.cell;
      app.book.transact('Insert Hyperlink', () => { const cur = s.get(s.sel.r, s.sel.c); s.set(s.sel.r, s.sel.c, { ...(cur || {}), link, v: cur?.v ?? arg.cell }); });
    },
  },
  {
    id: 'insert.headerFooter', label: 'Header & Footer', icon: 'call_to_action', tab: 'Insert', group: 'Text', edits: true, keywords: 'header footer page number print',
    run: (app) => { app.exec('view.pageLayout'); app.exec('layout.pageSetup', 'header'); },
  },
  { id: 'insert.sparkLine', label: 'Line Sparkline', icon: 'show_chart', tab: 'Insert', group: 'Sparklines', edits: true, keywords: 'sparkline mini chart in cell', run: (app) => sparklineDialog(app, 'line') },
  { id: 'insert.sparkColumn', label: 'Column Sparkline', icon: 'bar_chart', tab: 'Insert', group: 'Sparklines', edits: true, keywords: 'sparkline mini chart in cell', run: (app) => sparklineDialog(app, 'column') },
  { id: 'insert.sparkWinLoss', label: 'Win/Loss Sparkline', icon: 'align_vertical_center', tab: 'Insert', group: 'Sparklines', edits: true, keywords: 'sparkline win loss', run: (app) => sparklineDialog(app, 'winloss') },
  {
    id: 'sparkline.edit', label: 'Edit Sparklines', icon: 'edit', tab: 'Sparkline', group: 'Sparkline', edits: true, isEnabled: (app) => sparklinesInSelection(app).length > 0,
    run: async (app) => {
      const list = sparklinesInSelection(app);
      if (!list.length) { toast('Select cells that contain sparklines.'); return; }
      const sp = list[0];
      const r = await openDialog({
        title: 'Edit Sparklines', help: list.length > 1 ? `Changes apply to ${list.length} sparklines.` : undefined,
        fields: [
          { key: 'type', label: 'Type', type: 'radio', options: [{ value: 'line', label: 'Line' }, { value: 'column', label: 'Column' }, { value: 'winloss', label: 'Win/Loss' }], value: sp.type },
          ...(list.length === 1 ? [{ key: 'data', label: 'Data range', type: 'text' as const, value: sp.data }] : []),
          { key: 'color', label: 'Sparkline color', type: 'color', value: sp.color || accent(app, 0), half: true },
          { key: 'markers', label: 'Markers', type: 'checkbox', value: !!sp.markers, half: true },
          { key: 'high', label: 'High point', type: 'checkbox', value: !!sp.high, half: true },
          { key: 'low', label: 'Low point', type: 'checkbox', value: !!sp.low, half: true },
          { key: 'negative', label: 'Negative points', type: 'checkbox', value: !!sp.negative, half: true },
        ],
        validate: (v) => (v.data != null && !resolveRangeText(app.book, String(v.data).replace(/^=/, '').replace(/\$/g, ''), app.sheet) ? 'The data range isn\'t valid.' : null),
      });
      if (!r) return;
      app.book.transact('Edit Sparklines', () => {
        app.book.recordSheet(app.sheet);
        for (const x of sparklinesInSelection(app)) {
          x.type = r.type; x.color = r.color; x.markers = !!r.markers || undefined; x.high = !!r.high || undefined; x.low = !!r.low || undefined; x.negative = !!r.negative || undefined;
          if (r.data != null) x.data = String(r.data).replace(/^=/, '');
        }
      });
    },
  },
  {
    id: 'sparkline.clear', label: 'Clear Sparklines', icon: 'clear', tab: 'Sparkline', group: 'Group', edits: true, isEnabled: (app) => sparklinesInSelection(app).length > 0,
    run: (app) => {
      const list = sparklinesInSelection(app);
      if (!list.length) { toast('Select cells that contain sparklines.'); return; }
      app.book.transact('Clear Sparklines', () => { app.book.recordSheet(app.sheet); app.sheet.sparklines = app.sheet.sparklines.filter((x) => !list.includes(x)); });
    },
  },
  {
    id: 'sparkline.color', label: 'Sparkline Color', icon: 'format_color_fill', tab: 'Sparkline', group: 'Style', edits: true, isEnabled: (app) => sparklinesInSelection(app).length > 0,
    run: async (app, arg) => {
      if (!sparklinesInSelection(app).length) { toast('Select cells that contain sparklines.'); return; }
      const c = typeof arg === 'string' ? arg : await colorMenu(app.ui.anchorFor('sparkline.color'), { title: 'Sparkline Color' });
      if (!c) return;
      app.book.transact('Sparkline Color', () => { app.book.recordSheet(app.sheet); for (const x of sparklinesInSelection(app)) x.color = c; });
    },
  },
  {
    id: 'sparkline.type', label: 'Sparkline Type', icon: 'swap_horiz', tab: 'Sparkline', group: 'Type', kind: 'menu', edits: true, isEnabled: (app) => sparklinesInSelection(app).length > 0,
    run: (app, arg) => { const t = (['line', 'column', 'winloss'].includes(arg) ? arg : 'line') as Sparkline['type']; app.book.transact('Sparkline Type', () => { app.book.recordSheet(app.sheet); for (const x of sparklinesInSelection(app)) x.type = t; }); },
    menu: () => [{ id: 'sparkline.type', arg: 'line', label: 'Line', icon: 'show_chart' }, { id: 'sparkline.type', arg: 'column', label: 'Column', icon: 'bar_chart' }, { id: 'sparkline.type', arg: 'winloss', label: 'Win/Loss', icon: 'align_vertical_center' }],
  },
  { id: 'insert.slicer', label: 'Slicer', icon: 'filter_alt', tab: 'Insert', group: 'Filters', edits: true, keywords: 'slicer filter buttons', run: (app) => insertSlicer(app, false) },
  { id: 'insert.timeline', label: 'Timeline', icon: 'date_range', tab: 'Insert', group: 'Filters', edits: true, keywords: 'timeline date filter period', run: (app) => insertSlicer(app, true) },
  { id: 'insert.slicerPick', label: 'Choose Slicer Items', icon: 'filter_alt', tab: 'Insert', group: 'Filters', internal: true, edits: true, run: (app, arg) => slicerPick(app, arg ?? { id: app.tab.grid?.objects.selected }) },
  {
    id: 'insert.slicerClear', label: 'Clear Slicer Filter', icon: 'filter_alt_off', tab: 'Insert', group: 'Filters', edits: true,
    run: (app, arg) => {
      const o = findObject(app, arg?.id);
      if (!o || o.kind !== 'drawing' || o.def.shape !== 'slicer') { toast('Select a slicer first.'); return; }
      const st = { ...(o.def.fx || {}) } as SlicerState;
      const found = findTableForSlicer(app, st);
      if (!found) return;
      const ci = found.t.columns.findIndex((c) => c.name === st.column);
      st.selected = []; if (st.months?.length) { st.from = st.months[0]; st.to = st.months[st.months.length - 1]; }
      app.book.transact('Clear Slicer', () => { app.book.recordSheet(o.sheet); o.def.fx = { ...st }; if (ci >= 0) applySlicer(found.sheet, found.t, ci, st); });
    },
  },
]);

void icon; void promptDialog; void cellName; void keyRow; void keyCol;
