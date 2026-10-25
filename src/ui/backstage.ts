/** File backstage: Home, New, Open, Info, Save As, Export, Print, Share, Options. */
import type { App } from '../app';
import { h, icon, clear } from './dom';
import { TEMPLATES } from '../templates';
import { FORMATS } from '../io';
import { kvKeys, kvGet, kvDel } from '../util/store';
import { Workbook } from '../model/workbook';

const PAGES: [string, string, string][] = [
  ['home', 'Home', 'home'], ['new', 'New', 'note_add'], ['open', 'Open', 'folder_open'], ['info', 'Info', 'info'],
  ['save', 'Save', 'save'], ['saveas', 'Save As', 'save_as'], ['export', 'Export', 'ios_share'], ['print', 'Print', 'print'],
  ['share', 'Share', 'share'], ['close', 'Close', 'close'], ['options', 'Options', 'settings'],
];

export function renderBackstage(app: App, root: HTMLElement, page: string): void {
  clear(root);
  const nav = h('nav', { class: 'ql-bs-nav' }, h('button', { class: 'ql-bs-back', onclick: () => app.ui.hideBackstage(), 'aria-label': 'Back to workbook' }, icon('arrow_back'), ' Back'));
  for (const [id, label, ic] of PAGES) {
    nav.appendChild(h('button', { class: 'ql-bs-item' + (id === page ? ' ql-on' : ''), onclick: () => {
      if (id === 'save') { app.ui.hideBackstage(); app.exec('file.save'); return; }
      if (id === 'close') { app.exec('file.close'); return; }
      if (id === 'options') { app.exec('file.options'); return; }
      renderBackstage(app, root, id);
    } }, icon(ic), ' ', label));
  }
  const body = h('div', { class: 'ql-bs-body' });
  root.append(nav, body);
  const card = (title: string, sub: string, ic: string, onclick: () => void) => h('button', { class: 'ql-bs-card', onclick }, icon(ic, 'ql-bs-card-icon'), h('div', null, h('b', null, title), h('div', { class: 'ql-help' }, sub)));
  const book = app.book;
  switch (page) {
    case 'home':
    case 'new': {
      const hr = new Date().getHours();
      body.appendChild(h('h1', null, page === 'home' ? 'Good ' + (hr < 12 ? 'morning' : hr < 18 ? 'afternoon' : 'evening') : 'New'));
      const grid = h('div', { class: 'ql-bs-templates' });
      const cats = [...new Set(TEMPLATES.map((t) => t.category))];
      for (const c of page === 'home' ? cats.slice(0, 2) : cats) {
        for (const t of TEMPLATES.filter((x) => x.category === c)) grid.appendChild(h('button', { class: 'ql-template', onclick: () => { app.ui.hideBackstage(); app.exec('file.new', t.id); } }, h('div', { class: 'ql-template-thumb' }, icon(t.icon || 'grid_on')), h('b', null, t.name), h('div', { class: 'ql-help' }, t.description)));
      }
      body.appendChild(grid);
      if (page === 'new') {
        const mine = h('div', { class: 'ql-bs-templates' });
        body.append(h('h2', null, 'My Templates'), mine);
        kvKeys('tpl:').then(async (keys) => {
          if (!keys.length) mine.appendChild(h('p', { class: 'ql-help' }, 'Use File ▸ Save As ▸ Template to add your own.'));
          for (const k of keys) {
            const t = await kvGet(k);
            mine.appendChild(h('div', { class: 'ql-template' }, h('button', { class: 'ql-link', onclick: () => { app.ui.hideBackstage(); app.createTab(Workbook.fromJSON(t.json), { title: t.name }); } }, h('b', null, t.name)), h('button', { class: 'ql-iconbtn ql-sm', title: 'Delete template', onclick: async () => { await kvDel(k); renderBackstage(app, root, 'new'); } }, icon('delete'))));
          }
        });
      }
      if (page === 'home') { body.appendChild(h('h2', null, 'Recent')); recentList(app, body); }
      break;
    }
    case 'open': {
      body.append(h('h1', null, 'Open'),
        card('Browse…', 'Open a file from this device, Drive, or any document provider', 'folder_open', () => app.exec('file.open')),
        card('Encrypted workbook…', 'Open a password-protected .qgwe file', 'lock_open', () => app.exec('file.openEncrypted')),
        card('Open and Repair…', 'Rebuild a damaged .xlsx, .xls or .ods and show what was recovered', 'healing', () => app.exec('file.openRepair')),
        card('Recover Data from Any File…', 'Pull readable values and text out of any file, however damaged', 'text_snippet', () => app.exec('file.recoverText')),
        card('Text Import Wizard…', 'Import delimited or fixed-width text with column types', 'input', () => app.exec('data.fromText')),
        card('Recover unsaved workbooks', 'Drafts saved automatically before a crash or close', 'restore_page', () => app.exec('file.recover')),
        h('h2', null, 'Recent'));
      recentList(app, body, true);
      break;
    }
    case 'info': {
      const tab = app.tab;
      let cells = 0, formulas = 0;
      for (const s of book.sheets) for (const c of s.cells.values()) { if (c.v != null || c.f != null) cells++; if (c.f != null) formulas++; }
      body.append(h('h1', null, 'Info'), h('h2', null, tab.displayName),
        h('div', { class: 'ql-bs-cards' },
          card('Protect Workbook', book.structureLocked ? 'Structure is protected' : 'Control what changes people can make', 'shield_lock', () => app.exec('review.protectWorkbook')),
          card('Mark as Final', book.finalMarked ? 'Workbook is marked final' : 'Make the workbook read-only', 'verified', () => app.exec('review.markFinal')),
          card('Encrypt with Password', 'Save an AES-256 encrypted copy', 'enhanced_encryption', () => app.exec('file.encrypt')),
          card('Inspect Workbook', 'Find and remove comments, hidden sheets, personal information', 'policy', () => app.exec('file.inspect')),
          card('Check Accessibility', 'Find content people with disabilities might have trouble reading', 'accessibility_new', () => { app.ui.hideBackstage(); app.exec('review.accessibility'); }),
          card('Version History', 'Browse and restore snapshots', 'history', () => app.exec('file.versions')),
          card('Revert to Saved', 'Discard changes since the last save', 'settings_backup_restore', () => app.exec('file.revert'))),
        h('h3', null, 'Properties'),
        h('table', { class: 'ql-stats' },
          ...[['Title', book.props.title || '—'], ['Author', book.props.author || app.prefs.author], ['Sheets', String(book.sheets.length)], ['Cells with data', String(cells)], ['Formulas', String(formulas)], ['Format', tab.format.toUpperCase()], ['Location', tab.uri ? 'Device storage' : 'Not saved yet'], ['Created', new Date(book.props.created || tab.created).toLocaleString()], ['Last saved', tab.saved ? new Date(tab.saved).toLocaleString() : '—']].map(([k, v]) => h('tr', null, h('td', null, k), h('td', null, v)))),
        h('button', { class: 'ql-btn', onclick: () => app.exec('file.properties') }, 'All Properties…'));
      break;
    }
    case 'saveas':
    case 'export': {
      body.append(h('h1', null, page === 'saveas' ? 'Save As' : 'Export'), h('p', { class: 'ql-help' }, page === 'saveas' ? 'Choose a format. The system file dialog lets you pick the folder (device, SD card, Drive, Dropbox…).' : 'Save a copy in another format; the current workbook keeps its file.'));
      const grid = h('div', { class: 'ql-bs-cards' });
      for (const f of FORMATS.filter((x) => x.canExport)) grid.appendChild(card(f.label, `.${f.ext[0]}`, 'description', () => (page === 'saveas' ? app.exec('file.saveAsFormat', f.id) : app.exec('file.exportAs', f.id))));
      grid.appendChild(card('PDF', 'Through the system print service ("Save as PDF")', 'picture_as_pdf', () => app.exec('file.exportPdf')));
      grid.appendChild(card('Encrypted Anne\'s Grid (.qgwe)', 'AES-256 password protected', 'enhanced_encryption', () => app.exec('file.encrypt')));
      grid.appendChild(card('Template', 'Reuse this workbook as a starting point', 'bookmark_add', () => app.exec('file.saveTemplate')));
      body.appendChild(grid);
      break;
    }
    case 'print': {
      const p = app.sheet.print;
      body.append(h('h1', null, 'Print'),
        card('Print…', 'Opens the system print dialog (printers, Save as PDF)', 'print', () => app.exec('file.print')),
        card('Print Preview', 'See the pages as they will print', 'preview', () => app.exec('file.printPreview')),
        card('Page Setup…', `${p.paper}, ${p.orientation}, scale ${p.fitWide || p.fitTall ? 'fit to page' : p.scale + '%'}`, 'settings_applications', () => app.exec('layout.pageSetup')),
        card('Export as PDF', 'Choose "Save as PDF" in the print dialog', 'picture_as_pdf', () => app.exec('file.exportPdf')));
      break;
    }
    case 'share': {
      body.append(h('h1', null, 'Share'),
        card('Share a copy…', 'Send as .xlsx, .ods, .csv, PDF…', 'share', () => app.exec('file.share')),
        card('Email', 'Attach to an email', 'mail', () => app.exec('file.email')),
        card('Copy sheet as Markdown table', 'Paste into notes, chats or documentation', 'markdown', () => app.exec('file.copyMarkdown')));
      break;
    }
  }
}

function recentList(app: App, body: HTMLElement, full = false) {
  const list = h('div', { class: 'ql-recent' });
  const rec = app.prefs.recent;
  if (!rec.length) list.appendChild(h('p', { class: 'ql-help' }, 'Workbooks you open or save appear here.'));
  for (const r of rec.slice(0, full ? 40 : 8)) {
    list.appendChild(h('div', { class: 'ql-recent-row' },
      h('button', { class: 'ql-recent-open', onclick: () => app.exec('file.openRecent', r.key) }, icon('grid_on'), h('div', null, h('b', null, r.name), h('div', { class: 'ql-help' }, new Date(r.when).toLocaleString()))),
      h('button', { class: 'ql-iconbtn', title: r.pinned ? 'Unpin' : 'Pin', onclick: () => { r.pinned = !r.pinned; app.savePrefs(); renderBackstage(app, body.parentElement!, full ? 'open' : 'home'); } }, icon(r.pinned ? 'keep_off' : 'keep')),
      h('button', { class: 'ql-iconbtn', title: 'Remove from list', onclick: () => { app.prefs.recent = app.prefs.recent.filter((x) => x !== r); app.savePrefs(); renderBackstage(app, body.parentElement!, full ? 'open' : 'home'); } }, icon('close'))));
  }
  if (full && rec.length) list.appendChild(h('button', { class: 'ql-btn', onclick: () => app.exec('file.clearRecent') }, 'Clear unpinned'));
  body.appendChild(list);
}
