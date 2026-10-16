/**
 * Side panes (task panes). On phones they appear as bottom sheets. Command modules register their own
 * panes with registerPane(); the core ones (Navigator, Clipboard, Comments, Selection, Watch Window)
 * are defined here.
 */
import type { App } from '../app';
import { h, icon, clear } from './dom';
import { toast } from './dialog';
import { cellName, keyRow, keyCol, rangeName, MAX_COLS } from '../model/address';
import { formatValue } from '../model/numfmt';

export interface PaneDef {
  id: string;
  title: string;
  render: (app: App, el: HTMLElement, opts: any) => void;
  update?: (app: App, el: HTMLElement, opts: any) => void;
  onOpen?: (app: App, el: HTMLElement, opts: any) => void;
  onClose?: (app: App) => void;
}

export const PANES: Record<string, PaneDef> = {};

export function registerPane(def: PaneDef): void { PANES[def.id] = def; }

/** True while the user is typing in a field of the pane (redrawing would destroy the field). */
export const typingIn = (el: HTMLElement) => { const a = document.activeElement as HTMLElement | null; return !!a && el.contains(a) && (a.tagName === 'INPUT' || a.tagName === 'TEXTAREA' || a.tagName === 'SELECT'); };

const sigOf = (el: HTMLElement, sig: string) => { if (el.dataset.sig === sig) return false; el.dataset.sig = sig; return true; };

registerPane({
  id: 'navigator', title: 'Navigator',
  render(app, el) { this.update!(app, el, {}); },
  update(app, el) {
    const book = app.book;
    const sig = JSON.stringify([book.version, book.sheets.map((s) => s.name), book.names.length, app.tab.id]);
    if (!sigOf(el, sig) || typingIn(el)) return;
    clear(el);
    const section = (title: string, ic: string, items: { label: string; sub?: string; go: () => void }[]) => {
      el.appendChild(h('h3', { class: 'ql-pane-h' }, icon(ic), ` ${title} (${items.length})`));
      const list = h('div', { class: 'ql-nav-list' });
      for (const it of items) list.appendChild(h('button', { class: 'ql-nav-item', onclick: it.go }, h('span', null, it.label), it.sub ? h('span', { class: 'ql-help' }, it.sub) : null));
      if (!items.length) list.appendChild(h('div', { class: 'ql-help' }, 'None'));
      el.appendChild(list);
    };
    section('Sheets', 'tab', book.sheets.map((s, i) => ({ label: s.name, sub: s.visibility !== 'visible' ? 'hidden' : s.usedRange() ? rangeName(s.usedRange()!) : 'empty', go: () => app.exec('sheet.activate', i) })));
    section('Names', 'badge', book.names.map((n) => ({ label: n.name, sub: n.ref, go: () => app.exec('nav.goto', n.name) })));
    section('Tables', 'table', book.sheets.flatMap((s) => s.tables.map((t) => ({ label: t.name, sub: `${s.name}!${rangeName(t.range)}`, go: () => app.exec('nav.goto', t.name) }))));
    section('Charts and objects', 'bar_chart', book.sheets.flatMap((s, i) => [...s.charts.map((c) => ({ label: c.title || `${c.type} chart`, sub: s.name, go: () => { app.exec('sheet.activate', i); app.grid.objects.select(c.id); app.grid.scrollIntoView(c.at.r, c.at.c); } })), ...s.drawings.map((d) => ({ label: d.alt || d.text || d.kind, sub: s.name, go: () => { app.exec('sheet.activate', i); app.grid.objects.select(d.id); app.grid.scrollIntoView(d.at.r, d.at.c); } }))]));
    section('Comments and notes', 'forum', book.sheets.flatMap((s, i) => [...s.comments].map(([k, c]) => ({ label: `${s.name}!${cellName(keyRow(k), keyCol(k))}`, sub: c.text.slice(0, 60), go: () => { app.exec('sheet.activate', i); app.grid.select(keyRow(k), keyCol(k)); } }))));
  },
});

registerPane({
  id: 'clipboard', title: 'Clipboard',
  render(app, el) { this.update!(app, el, {}); },
  update(app, el) {
    if (!sigOf(el, String(app.clipboard.length) + ':' + (app.clipboard[0]?.time || 0))) return;
    clear(el);
    el.appendChild(h('div', { class: 'ql-row' }, h('button', { class: 'ql-btn', onclick: () => { app.clipboard = []; app.refresh(); } }, 'Clear All')));
    if (!app.clipboard.length) el.appendChild(h('p', { class: 'ql-help' }, 'Cut or copy cells to collect items here, then click an item to paste it.'));
    for (const it of app.clipboard) {
      el.appendChild(h('button', { class: 'ql-clip-item', title: 'Paste', onclick: () => app.exec('edit.pasteText', it.text) }, h('pre', null, it.text.slice(0, 300)), h('div', { class: 'ql-help' }, new Date(it.time).toLocaleTimeString())));
    }
  },
});

registerPane({
  id: 'comments', title: 'Comments',
  render(app, el) { this.update!(app, el, {}); },
  update(app, el) {
    const s = app.sheet;
    const sig = JSON.stringify([s.id, [...s.comments].map(([k, c]) => [k, c.text, c.resolved, c.replies?.length])]);
    if (!sigOf(el, sig) || typingIn(el)) return;
    clear(el);
    el.appendChild(h('div', { class: 'ql-row' }, h('button', { class: 'ql-btn', onclick: () => app.exec('review.newComment') }, icon('add_comment'), ' New')));
    const items = [...s.comments].filter(([, c]) => !c.note).sort((a, b) => a[0] - b[0]);
    if (!items.length) el.appendChild(h('p', { class: 'ql-help' }, 'No comments on this sheet.'));
    for (const [k, c] of items) {
      const r = keyRow(k), col = keyCol(k);
      const reply = h('input', { class: 'ql-input', placeholder: 'Reply…' }) as HTMLInputElement;
      reply.addEventListener('keydown', (e) => { if (e.key === 'Enter' && reply.value.trim()) { app.exec('review.replyComment', { r, c: col, text: reply.value.trim() }); reply.value = ''; } });
      el.appendChild(h('div', { class: 'ql-comment-card' + (c.resolved ? ' ql-resolved' : '') },
        h('button', { class: 'ql-link', onclick: () => app.grid.select(r, col) }, cellName(r, col)),
        h('div', null, h('b', null, c.author), ' ', h('span', { class: 'ql-help' }, new Date(c.time).toLocaleString())),
        h('div', { class: 'ql-comment-text' }, c.text),
        ...(c.replies || []).map((rp) => h('div', { class: 'ql-comment-reply' }, h('b', null, rp.author), ' ', rp.text)),
        reply,
        h('div', { class: 'ql-row' },
          h('button', { class: 'ql-btn ql-sm', onclick: () => app.exec('review.resolveComment', { r, c: col }) }, c.resolved ? 'Reopen' : 'Resolve'),
          h('button', { class: 'ql-btn ql-sm', onclick: () => app.exec('review.deleteComment', { r, c: col }) }, 'Delete'))));
    }
  },
});

registerPane({
  id: 'selection', title: 'Selection',
  render(app, el) { this.update!(app, el, {}); },
  update(app, el) {
    const s = app.sheet;
    const g = app.grid;
    const sig = JSON.stringify([s.id, s.charts.map((c) => [c.id, c.title]), s.drawings.map((d) => [d.id, d.alt, d.text]), g.objects.selected]);
    if (!sigOf(el, sig)) return;
    clear(el);
    const objs = [...s.charts.map((c) => ({ id: c.id, label: c.title || `Chart (${c.type})`, ic: 'bar_chart' })), ...s.drawings.map((d) => ({ id: d.id, label: d.alt || d.text || d.kind, ic: d.kind === 'image' ? 'image' : 'shapes' }))];
    if (!objs.length) el.appendChild(h('p', { class: 'ql-help' }, 'This sheet has no charts, pictures or shapes.'));
    for (const o of objs) el.appendChild(h('button', { class: 'ql-nav-item' + (g.objects.selected === o.id ? ' ql-on' : ''), onclick: () => g.objects.select(o.id) }, icon(o.ic), ' ', o.label));
  },
});

registerPane({
  id: 'watch', title: 'Watch Window',
  render(app, el) { this.update!(app, el, {}); },
  update(app, el) {
    const tab = app.tab;
    const vals = tab.watches.map((w) => { const sh = app.book.sheetByName(w.sheet); return sh ? formatValue(sh.value(w.r, w.c), sh.styleOf(w.r, w.c).numFmt).text + '|' + (sh.get(w.r, w.c)?.f || '') : '#REF!'; });
    if (!sigOf(el, JSON.stringify([tab.watches, vals])) || typingIn(el)) return;
    clear(el);
    el.appendChild(h('div', { class: 'ql-row' },
      h('button', { class: 'ql-btn', onclick: () => { const s = app.sheet; const { r, c } = s.sel; if (!tab.watches.some((w) => w.sheet === s.name && w.r === r && w.c === c)) tab.watches.push({ sheet: s.name, r, c }); app.refresh(); } }, icon('add'), ' Add Watch'),
      h('button', { class: 'ql-btn', onclick: () => { tab.watches = []; app.refresh(); } }, 'Delete All')));
    const table = h('table', { class: 'ql-stats' }, h('tr', null, h('th', null, 'Sheet'), h('th', null, 'Cell'), h('th', null, 'Value'), h('th', null, 'Formula'), h('th', null, '')));
    tab.watches.forEach((w, i) => {
      const [v, f] = vals[i].split('|');
      table.appendChild(h('tr', null, h('td', null, w.sheet), h('td', null, h('button', { class: 'ql-link', onclick: () => app.exec('nav.goto', `'${w.sheet}'!${cellName(w.r, w.c)}`) }, cellName(w.r, w.c))), h('td', null, v), h('td', null, f ? '=' + f : ''), h('td', null, h('button', { class: 'ql-iconbtn ql-sm', 'aria-label': 'Remove', onclick: () => { tab.watches.splice(i, 1); app.refresh(); } }, icon('close')))));
    });
    el.appendChild(table);
  },
});

void toast; void MAX_COLS;
