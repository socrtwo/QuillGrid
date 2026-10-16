/** View tab: workbook views, show/hide, zoom, freeze panes, split, windows, appearance, panes. */
import { def, getCommand } from './registry';
import { addAllQatChoices } from '../prefs';
import type { App, ViewMode } from '../app';
import { activeRange, activeCell } from './ops';
import { openDialog, toast } from '../ui/dialog';
import { computePageBreaks } from '../print/pages';

function setView(app: App, m: ViewMode): void {
  app.tab.viewMode = m;
  const g = app.grid as any;
  g.pageBreaks = m === 'normal' && !g.showPageBreaks ? undefined : computePageBreaks(app.sheet);
  app.grid.invalidate();
  app.refresh();
}

function setSheetFlag(app: App, key: 'showGrid' | 'showHeadings' | 'showFormulas' | 'showZeros' | 'rtl', label: string): void {
  const s = app.sheet;
  app.book.transact(label, () => { app.book.recordSheet(s); (s as any)[key] = !s[key]; });
  app.grid.layout();
}

def([
  { id: 'view.normal', label: 'Normal', icon: 'grid_on', tab: 'View', group: 'Workbook Views', kind: 'toggle', noRecord: true, run: (app) => setView(app, 'normal'), isActive: (app) => app.tab.viewMode === 'normal' },
  { id: 'view.pageBreak', label: 'Page Break Preview', icon: 'splitscreen', tab: 'View', group: 'Workbook Views', kind: 'toggle', noRecord: true, run: (app) => setView(app, 'pageBreak'), isActive: (app) => app.tab.viewMode === 'pageBreak' },
  { id: 'view.pageLayout', label: 'Page Layout', icon: 'description', tab: 'View', group: 'Workbook Views', kind: 'toggle', noRecord: true, run: (app) => setView(app, 'pageLayout'), isActive: (app) => app.tab.viewMode === 'pageLayout' },
  { id: 'view.showPageBreaks', label: 'Show Page Breaks', icon: 'insert_page_break', tab: 'View', group: 'Workbook Views', kind: 'toggle', noRecord: true, run: (app) => { const g = app.grid as any; g.showPageBreaks = !g.showPageBreaks; setView(app, app.tab.viewMode); }, isActive: (app) => !!(app.grid as any).showPageBreaks },
  {
    id: 'view.customViews', label: 'Custom Views…', icon: 'view_carousel', tab: 'View', group: 'Workbook Views', noRecord: true,
    run: async (app) => {
      const key = 'customViews';
      const views: any[] = ((app.book as any)[key] ||= []);
      const r = await openDialog({ title: 'Custom Views', fields: [{ key: 'v', label: 'Views', type: 'list', rows: 8, options: views.map((v) => v.name) }, { key: 'name', label: 'Name for a new view', type: 'text', value: '' }], ok: 'Show', buttons: [{ label: 'Add', onClick: (vals, ctl) => { if (!vals.name) return; const s = app.sheet; views.push({ name: vals.name, sheet: s.name, sel: JSON.parse(JSON.stringify(s.sel)), zoom: s.zoom, hiddenRows: [...s.hiddenRows], hiddenCols: [...s.hiddenCols], filter: s.filter ? JSON.parse(JSON.stringify(s.filter)) : null, print: JSON.parse(JSON.stringify(s.print)) }); ctl.close(); toast(`View "${vals.name}" added`); } }, { label: 'Delete', onClick: (vals, ctl) => { const n = Array.isArray(vals.v) ? vals.v[0] : vals.v; const i = views.findIndex((x) => x.name === n); if (i >= 0) views.splice(i, 1); ctl.close(); } }] });
      const name = r && (Array.isArray(r.v) ? r.v[0] : r.v);
      const v = views.find((x) => x.name === name);
      if (!v) return;
      const idx = app.book.sheets.findIndex((s) => s.name === v.sheet);
      if (idx < 0) return;
      app.exec('sheet.activate', idx);
      const s = app.sheet;
      app.book.transact('Show View', () => { app.book.recordSheet(s); s.hiddenRows = new Set(v.hiddenRows); s.hiddenCols = new Set(v.hiddenCols); s.filter = v.filter; s.print = v.print; s.zoom = v.zoom; s.sel = v.sel; });
    },
  },
  { id: 'view.formulaBar', label: 'Formula Bar', icon: 'function', tab: 'View', group: 'Show', kind: 'toggle', noRecord: true, run: (app) => { app.prefs.showFormulaBar = !app.prefs.showFormulaBar; app.savePrefs(); app.refresh(); setTimeout(() => app.grid.layout(), 50); }, isActive: (app) => app.prefs.showFormulaBar },
  { id: 'view.gridlines', label: 'Gridlines', icon: 'grid_4x4', tab: 'View', group: 'Show', kind: 'toggle', run: (app) => setSheetFlag(app, 'showGrid', 'Gridlines'), isActive: (app) => app.sheet.showGrid },
  { id: 'view.headings', label: 'Headings', icon: 'view_headline', tab: 'View', group: 'Show', kind: 'toggle', run: (app) => setSheetFlag(app, 'showHeadings', 'Headings'), isActive: (app) => app.sheet.showHeadings },
  { id: 'view.ruler', label: 'Ruler', icon: 'straighten', tab: 'View', group: 'Show', kind: 'toggle', noRecord: true, run: (app) => { toast('The ruler is shown in Page Layout view.'); setView(app, 'pageLayout'); }, isActive: (app) => app.tab.viewMode === 'pageLayout' },
  { id: 'view.showFormulas', label: 'Show Formulas', icon: 'functions', tab: 'Formulas', group: 'Formula Auditing', kind: 'toggle', keys: ['Mod-`'], run: (app) => setSheetFlag(app, 'showFormulas', 'Show Formulas'), isActive: (app) => app.sheet.showFormulas },
  { id: 'view.showZeros', label: 'Show Zero Values', icon: 'exposure_zero', tab: 'View', group: 'Show', kind: 'toggle', run: (app) => setSheetFlag(app, 'showZeros', 'Show Zeros'), isActive: (app) => app.sheet.showZeros },
  { id: 'view.rightToLeft', label: 'Sheet Right-to-Left', icon: 'format_textdirection_r_to_l', tab: 'Page Layout', group: 'Sheet Options', kind: 'toggle', run: (app) => setSheetFlag(app, 'rtl', 'Sheet Right-to-Left'), isActive: (app) => app.sheet.rtl },
  {
    id: 'view.gridColor', label: 'Gridline Color…', icon: 'border_color', tab: 'View', group: 'Show',
    run: async (app) => { const { colorMenu } = await import('../ui/dialog'); const c = await colorMenu(app.ui.anchorFor('view.gridlines'), { auto: 'Automatic' }); if (c === undefined) return; const s = app.sheet; app.book.transact('Gridline Color', () => { app.book.recordSheet(s); s.gridColor = c; }); },
  },
  // Zoom
  {
    id: 'view.zoom', label: 'Zoom…', icon: 'zoom_in', tab: 'View', group: 'Zoom', noRecord: true,
    run: async (app) => {
      const cur = Math.round(app.sheet.zoom * 100);
      const r = await openDialog({ title: 'Zoom', fields: [{ key: 'z', label: 'Magnification', type: 'radio', options: ['200', '100', '75', '50', '25', 'fit', 'custom'].map((v) => ({ value: v, label: v === 'fit' ? 'Fit selection' : v === 'custom' ? 'Custom' : `${v}%` })), value: ['200', '100', '75', '50', '25'].includes(String(cur)) ? String(cur) : 'custom' }, { key: 'custom', label: 'Custom (%)', type: 'number', min: 10, max: 400, value: cur }] });
      if (!r) return;
      if (r.z === 'fit') return app.exec('view.zoomSelection');
      app.grid.setZoom((r.z === 'custom' ? r.custom : Number(r.z)) / 100);
    },
  },
  { id: 'view.zoom100', label: '100%', icon: 'crop_free', tab: 'View', group: 'Zoom', noRecord: true, run: (app) => app.grid.setZoom(1) },
  { id: 'view.zoomIn', label: 'Zoom In', icon: 'zoom_in', tab: 'View', group: 'Zoom', keys: ['Mod-Alt-='], noRecord: true, run: (app) => app.grid.setZoom(Math.min(4, Math.round((app.sheet.zoom + 0.1) * 10) / 10)) },
  { id: 'view.zoomOut', label: 'Zoom Out', icon: 'zoom_out', tab: 'View', group: 'Zoom', keys: ['Mod-Alt--'], noRecord: true, run: (app) => app.grid.setZoom(Math.max(0.1, Math.round((app.sheet.zoom - 0.1) * 10) / 10)) },
  {
    id: 'view.zoomSelection', label: 'Zoom to Selection', icon: 'zoom_out_map', tab: 'View', group: 'Zoom', noRecord: true,
    run: (app) => {
      const g = app.grid;
      const rg = activeRange(app);
      const z0 = app.sheet.zoom;
      const rc = g.rangeRect({ ...rg, r2: Math.min(rg.r2, rg.r1 + 500), c2: Math.min(rg.c2, rg.c1 + 100) });
      const z = Math.max(0.1, Math.min(4, Math.min((g.width - g.headerW - 20) / (rc.w / z0), (g.height - g.headerH - 20) / (rc.h / z0))));
      g.setZoom(z);
      g.scrollIntoView(rg.r1, rg.c1);
    },
  },
  // Window
  {
    id: 'view.freezePanes', label: 'Freeze Panes', icon: 'view_compact', tab: 'View', group: 'Window', kind: 'menu', run: (app, mode) => freeze(app, mode),
    menu: (app) => [app.sheet.freeze.rows || app.sheet.freeze.cols ? { id: 'view.freezePanes', arg: 'off', label: 'Unfreeze Panes', icon: 'view_comfy' } : { id: 'view.freezePanes', arg: 'panes', label: 'Freeze Panes', icon: 'view_compact' }, { id: 'view.freezePanes', arg: 'row', label: 'Freeze Top Row', icon: 'border_top' }, { id: 'view.freezePanes', arg: 'col', label: 'Freeze First Column', icon: 'border_left' }],
    isActive: (app) => !!(app.sheet.freeze.rows || app.sheet.freeze.cols),
  },
  { id: 'view.freezeTopRow', label: 'Freeze Top Row', icon: 'border_top', tab: 'View', group: 'Window', run: (app) => freeze(app, 'row') },
  { id: 'view.freezeFirstCol', label: 'Freeze First Column', icon: 'border_left', tab: 'View', group: 'Window', run: (app) => freeze(app, 'col') },
  { id: 'view.unfreeze', label: 'Unfreeze Panes', icon: 'view_comfy', tab: 'View', group: 'Window', run: (app) => freeze(app, 'off') },
  {
    id: 'view.split', label: 'Split', icon: 'vertical_split', tab: 'View', group: 'Window', kind: 'toggle',
    tip: 'Divide the window at the active cell so different parts of the sheet stay in view (shown as frozen panes)',
    run: (app) => freeze(app, app.sheet.freeze.rows || app.sheet.freeze.cols ? 'off' : 'panes'), isActive: (app) => !!(app.sheet.freeze.rows || app.sheet.freeze.cols),
  },
  { id: 'view.newWindow', label: 'New Window', icon: 'open_in_new', tab: 'View', group: 'Window', noRecord: true, run: async (app) => { const { Workbook } = await import('../model/workbook'); const json = app.book.toJSON(); const t = app.createTab(Workbook.fromJSON(json), { title: `${app.tab.title}:2`, fileName: '' }); t.dirty = true; toast('A copy of the workbook opened in a new tab. Changes are not linked; save it under a new name.'); } },
  { id: 'view.arrangeAll', label: 'Arrange All', icon: 'grid_view', tab: 'View', group: 'Window', noRecord: true, run: (app) => app.exec('view.switchWindows') },
  { id: 'view.hideWindow', label: 'Hide Window', icon: 'visibility_off', tab: 'View', group: 'Window', noRecord: true, run: (app) => { if (app.docs.length < 2) { toast('Open another workbook first.'); return; } (app.tab as any).hidden = true; const next = app.docs.find((d) => !(d as any).hidden); if (next) app.activate(next); } },
  { id: 'view.unhideWindow', label: 'Unhide Window…', icon: 'visibility', tab: 'View', group: 'Window', noRecord: true, run: (app) => { const hidden = app.docs.filter((d) => (d as any).hidden); if (!hidden.length) { toast('No hidden windows.'); return; } for (const d of hidden) (d as any).hidden = false; app.activate(hidden[0]); } },
  {
    id: 'view.switchWindows', label: 'Switch Windows', icon: 'tab', tab: 'View', group: 'Window', kind: 'menu', noRecord: true,
    run: (app, id) => { const d = app.docs.find((x) => x.id === id); if (d) app.activate(d); },
    menu: (app) => app.docs.map((d) => ({ id: 'view.switchWindows', arg: d.id, label: d.displayName + (d.dirty ? ' •' : ''), checked: d === app.tab })),
  },
  { id: 'view.nextWindow', label: 'Next Workbook', icon: 'arrow_forward', tab: 'View', group: 'Window', keys: ['Mod-F6'], noRecord: true, run: (app) => { const i = app.docs.indexOf(app.tab); app.activate(app.docs[(i + 1) % app.docs.length]); } },
  // Appearance and panes
  { id: 'view.darkSheet', label: 'Dark Cells', icon: 'dark_mode', tab: 'View', group: 'Appearance', kind: 'toggle', noRecord: true, tip: 'Show cells dark when the app theme is dark', run: (app) => { app.prefs.darkSheet = !app.prefs.darkSheet; if (app.prefs.darkSheet && app.prefs.theme === 'light') app.prefs.theme = 'dark'; app.savePrefs(); app.ui.applyPrefs(); }, isActive: (app) => app.prefs.darkSheet },
  { id: 'view.fullScreen', label: 'Full Screen', icon: 'fullscreen', tab: 'View', group: 'Appearance', kind: 'toggle', keys: ['Mod-Shift-F1'], noRecord: true, run: (app) => { document.body.classList.toggle('ql-fullscreen'); try { if (document.fullscreenElement) void document.exitFullscreen(); else if (document.body.classList.contains('ql-fullscreen')) void document.documentElement.requestFullscreen?.(); } catch { /* */ } setTimeout(() => app.grid.layout(), 100); }, isActive: () => document.body.classList.contains('ql-fullscreen') },
  { id: 'view.navigator', label: 'Navigator', icon: 'list_alt', tab: 'View', group: 'Appearance', noRecord: true, origin: 'LibreOffice Calc Navigator', run: (app) => app.ui.togglePane('navigator') },
  { id: 'view.clipboardPane', label: 'Clipboard Pane', icon: 'content_paste_search', tab: 'Home', group: 'Clipboard', noRecord: true, run: (app) => app.ui.togglePane('clipboard') },
  { id: 'view.selectionPane', label: 'Selection Pane', icon: 'layers', tab: 'Page Layout', group: 'Arrange', noRecord: true, run: (app) => app.ui.togglePane('selection') },
  { id: 'view.cmdSearch', label: 'Search Commands', icon: 'search', tab: 'View', group: 'Appearance', keys: ['Alt-q'], popular: true, noRecord: true, run: (app) => app.ui.commandPalette() },
  { id: 'view.touchMode', label: 'Touch/Mouse Mode', icon: 'touch_app', tab: 'View', group: 'Appearance', kind: 'toggle', noRecord: true, tip: 'Add space between commands for use with fingers (on) or a mouse (off)', run: (app) => { app.prefs.touchMode = !app.prefs.touchMode; app.savePrefs(); app.ui.applyPrefs(); setTimeout(() => app.grid.layout(), 50); }, isActive: (app) => app.prefs.touchMode },
  { id: 'view.showQat', label: 'Show Quick Access Bar', icon: 'bolt', tab: 'View', group: 'Appearance', kind: 'toggle', noRecord: true, run: (app) => { app.prefs.qatHidden = !app.prefs.qatHidden; app.savePrefs(); app.ui.applyPrefs(); }, isActive: (app) => !app.prefs.qatHidden },
  { id: 'view.qatAtBottom', label: 'Quick Access Bar at the Bottom', icon: 'swap_vert', tab: 'View', group: 'Appearance', kind: 'toggle', noRecord: true, tip: 'Show the Quick Access Bar at the bottom of the window, near the on-screen keyboard', run: (app) => { app.prefs.qatPosition = app.prefs.qatPosition === 'bottom' ? 'top' : 'bottom'; app.savePrefs(); app.ui.applyPrefs(); setTimeout(() => app.grid.layout(), 50); }, isActive: (app) => app.prefs.qatPosition === 'bottom' },
  { id: 'view.qatAddAll', label: 'Add All Quick Access Choices (A–Z)', icon: 'sort_by_alpha', tab: 'View', group: 'Appearance', noRecord: true, tip: 'Put every choice from the Quick Access Bar menu on the bar: Save and Print first, the rest in alphabetical order', run: (app) => { app.prefs.qat = addAllQatChoices(app.prefs.qat, (id) => getCommand(id)?.label); app.prefs.qatHidden = false; app.savePrefs(); app.ui.applyPrefs(); toast('Added all Quick Access choices'); } },
  { id: 'view.collapseRibbon', label: 'Collapse the Ribbon', icon: 'expand_less', tab: 'View', group: 'Appearance', keys: ['Mod-F1'], noRecord: true, run: (app) => { app.prefs.ribbon = app.prefs.ribbon === 'full' ? 'tabs' : 'full'; app.savePrefs(); app.ui.applyPrefs(); setTimeout(() => app.grid.layout(), 50); } },
  { id: 'view.sheetViewMenu', label: 'Sheet View', icon: 'visibility', tab: 'View', group: 'Sheet View', kind: 'menu', noRecord: true, run: () => undefined, menu: () => [{ id: 'view.customViews', label: 'Custom Views…', icon: 'view_carousel' }, { id: 'view.showPageBreaks', label: 'Show Page Breaks', icon: 'insert_page_break' }, { id: 'view.showZeros', label: 'Show Zero Values', icon: 'exposure_zero' }, { id: 'view.gridColor', label: 'Gridline Color…', icon: 'border_color' }] },
]);

function freeze(app: App, mode: string | undefined): void {
  const s = app.sheet;
  const { r, c } = activeCell(app);
  const g = app.grid;
  const top = Math.max(0, g.rowAt(g.headerH + g.frozenH)), left = Math.max(0, g.colAt(g.headerW + g.frozenW));
  let next = { rows: 0, cols: 0 };
  if (mode === 'row') next = { rows: Math.max(1, top + 1) - 0 > 1 ? top + 1 : 1, cols: 0 };
  else if (mode === 'col') next = { rows: 0, cols: Math.max(1, left + 1) > 1 ? left + 1 : 1 };
  else if (mode === 'off') next = { rows: 0, cols: 0 };
  else next = s.freeze.rows || s.freeze.cols ? { rows: 0, cols: 0 } : { rows: r, cols: c };
  if (mode === 'row') next.rows = 1 + (s.scroll.r > 0 ? 0 : 0);
  if (mode === 'col') next.cols = 1;
  app.book.transact('Freeze Panes', () => { app.book.recordSheet(s); s.freeze = next; });
  g.scroller.scrollTop = 0; g.scroller.scrollLeft = 0;
  g.layout();
}
