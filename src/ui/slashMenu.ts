/**
 * The slash menu: a keyboard-driven command menu in the manner of Lotus 1-2-3's "/" menu. The whole
 * ribbon (and File) becomes a hierarchy of one-line menus: type the highlighted letter of an item
 * to open it, or move with the arrow keys and press Enter; Escape or Backspace goes back one level.
 * The line under the menu shows the next level's items, or the command's description, as 1-2-3 did.
 */
import type { App } from '../app';
import { getCommand, type MenuEntry } from '../commands/registry';
import { tabsFor, BACKSTAGE_ITEMS } from './ribbonLayout';
import { assignKeys } from '../heritage/models';
import { h, clear } from './dom';
import { toast } from './dialog';

export interface SlashItem { label: string; tip?: string; id?: string; arg?: any; children?: () => SlashItem[]; disabled?: boolean }

const clean = (s: string) => s.replace(/…$/, '').trim();

function fromEntries(app: App, entries: MenuEntry[]): SlashItem[] {
  const out: SlashItem[] = [];
  for (const e of entries) {
    if (e.separator || e.heading) continue;
    if (e.submenu) { const sub = e.submenu; out.push({ label: clean(e.label || ''), children: () => fromEntries(app, sub) }); continue; }
    if (!e.id) continue;
    const c = getCommand(e.id);
    if (!c) continue;
    out.push({ label: clean(e.label || c.label), id: e.id, arg: e.arg, tip: c.tip, disabled: !!c.isEnabled && !c.isEnabled(app) });
  }
  return out;
}

/** Turns a command into a menu item: commands with option menus become submenus (the command itself first). */
export function commandItem(app: App, id: string): SlashItem | null {
  const c = getCommand(id);
  if (!c || c.internal) return null;
  const disabled = !!c.isEnabled && !c.isEnabled(app);
  const label = clean(c.label) + (c.kind === 'toggle' && c.isActive ? (c.isActive(app) ? ' (on)' : ' (off)') : '');
  if (c.menu && (c.kind === 'menu' || c.kind === 'split' || c.kind === 'gallery')) {
    const menu = c.menu;
    return {
      label, tip: c.tip, disabled,
      children: () => {
        let entries: MenuEntry[] = [];
        try { entries = menu(app); } catch { entries = []; }
        const items = fromEntries(app, entries);
        // A split button runs itself too; a plain menu command may run without an argument.
        if (c.kind === 'split' && !items.some((x) => x.id === id && x.arg === undefined)) items.unshift({ label: clean(c.label), id, tip: c.tip });
        return items.length ? items : [{ label: clean(c.label), id, tip: c.tip }];
      },
    };
  }
  return { label, id, tip: c.tip, disabled };
}

/** Top level: File, then each ribbon tab (with the user's customizations), then Quit. */
export function slashTree(app: App): SlashItem[] {
  const file: SlashItem = { label: 'File', tip: 'New, open, save, print, export, options', children: () => BACKSTAGE_ITEMS.map((id) => commandItem(app, id)).filter((x): x is SlashItem => !!x) };
  const tabs = tabsFor(app).filter((t) => !t.contextual || t.contextual(app)).map((t): SlashItem => ({
    label: t.label,
    tip: t.groups.map((g) => g.label).join(' · '),
    children: () => t.groups.map((g): SlashItem => ({
      label: g.label,
      tip: g.items.map((id) => getCommand(id)?.label).filter(Boolean).join(' · '),
      children: () => {
        const ids = g.launcher && !g.items.includes(g.launcher) ? [...g.items, g.launcher] : g.items;
        return ids.map((id) => commandItem(app, id)).filter((x): x is SlashItem => !!x);
      },
    })),
  }));
  return [file, ...tabs.filter((t) => t.label !== 'Help'), ...tabs.filter((t) => t.label === 'Help'), { label: 'Quit', tip: 'Leave the menu' }];
}

/** Top-level tabs that keep their first letter when two start alike (Data before Draw, File before Formulas). */
const ROOT_PRIORITY = ['File', 'Home', 'Insert', 'Formulas', 'Data', 'View', 'Page Layout', 'Review', 'Tools', 'AI', 'Developer', 'Help', 'Quit'];

/** Accelerator letters for a menu level (the main menu favors the most used tabs). */
export function slashKeys(items: SlashItem[]): string[] {
  const labels = items.map((x) => x.label);
  const prio = ROOT_PRIORITY.map((l) => labels.indexOf(l)).filter((i) => i >= 0);
  return assignKeys(labels, prio.length >= 3 ? prio : undefined);
}

let open: HTMLElement | null = null;

export function openSlashMenu(app: App): void {
  if (open) return;
  if (app.grid?.editing) app.grid.commitEdit();
  const stack: { title: string; items: SlashItem[]; keys: string[]; sel: number }[] = [];
  const push = (title: string, items: SlashItem[]) => stack.push({ title, items, keys: slashKeys(items), sel: 0 });
  push('', slashTree(app));

  const overlay = h('div', { class: 'ql-overlay ql-palette-overlay ql-slash-overlay' });
  const box = h('div', { class: 'ql-slash', role: 'menu', tabindex: '-1', 'aria-label': 'Slash menu' });
  const pathEl = h('div', { class: 'ql-slash-path' });
  const itemsEl = h('div', { class: 'ql-slash-items' });
  const descEl = h('div', { class: 'ql-slash-desc' });
  box.append(pathEl, itemsEl, descEl, h('div', { class: 'ql-slash-hint' }, 'Type a highlighted letter or use ← → and Enter · Esc goes back'));
  overlay.appendChild(box);

  const close = () => {
    if (!open) return;
    overlay.remove();
    open = null;
    document.removeEventListener('keydown', onKey, true);
    app.focus();
  };

  const choose = (i: number) => {
    const lvl = stack[stack.length - 1];
    const it = lvl.items[i];
    if (!it) return;
    if (it.label === 'Quit' && !it.id && !it.children) { close(); return; }
    if (it.disabled) { toast(`${it.label} is not available here.`); return; }
    if (it.children) {
      const kids = it.children();
      if (!kids.length) { toast('Nothing to choose here.'); return; }
      push(it.label, kids);
      render();
      return;
    }
    if (it.id) { const { id, arg } = it; close(); app.exec(id, arg); }
  };

  const render = () => {
    const lvl = stack[stack.length - 1];
    clear(pathEl); clear(itemsEl); clear(descEl);
    pathEl.append(h('span', { class: 'ql-slash-mode' }, 'MENU'), ' ', stack.slice(1).map((x) => x.title).join(' ▸ ') || 'Main menu');
    lvl.items.forEach((it, i) => {
      const k = lvl.keys[i];
      const pos = k ? it.label.toUpperCase().indexOf(k) : -1;
      const label = pos >= 0 ? [it.label.slice(0, pos), h('u', null, it.label[pos]), it.label.slice(pos + 1)] : [it.label, k ? h('span', { class: 'ql-slash-k' }, ` (${k})`) : null];
      itemsEl.appendChild(h('button', {
        class: 'ql-slash-item' + (i === lvl.sel ? ' ql-sel' : '') + (it.disabled ? ' ql-disabled' : ''),
        role: 'menuitem', 'aria-keyshortcuts': k || undefined,
        onmouseenter: () => { lvl.sel = i; renderDesc(); markSel(); },
        onclick: () => { lvl.sel = i; choose(i); },
      }, ...label, it.children ? h('span', { class: 'ql-slash-more' }, ' ▸') : null));
    });
    renderDesc();
  };
  const markSel = () => { const lvl = stack[stack.length - 1]; [...itemsEl.children].forEach((el, i) => el.classList.toggle('ql-sel', i === lvl.sel)); (itemsEl.children[lvl.sel] as HTMLElement | undefined)?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' }); };
  const renderDesc = () => {
    const lvl = stack[stack.length - 1];
    const it = lvl.items[lvl.sel];
    clear(descEl);
    if (!it) return;
    let text = it.tip || '';
    if (!text && it.children) { try { text = it.children().map((x) => x.label).join('   '); } catch { text = ''; } }
    descEl.textContent = text || (it.id ? getCommand(it.id)?.label || '' : '');
  };

  const onKey = (e: KeyboardEvent) => {
    if (!open) return;
    const lvl = stack[stack.length - 1];
    const n = lvl.items.length;
    let handled = true;
    switch (e.key) {
      case 'ArrowRight': case 'ArrowDown': case 'Tab': lvl.sel = (lvl.sel + (e.shiftKey && e.key === 'Tab' ? n - 1 : 1)) % n; markSel(); renderDesc(); break;
      case 'ArrowLeft': case 'ArrowUp': lvl.sel = (lvl.sel + n - 1) % n; markSel(); renderDesc(); break;
      case 'Home': lvl.sel = 0; markSel(); renderDesc(); break;
      case 'End': lvl.sel = n - 1; markSel(); renderDesc(); break;
      case 'Enter': case ' ': choose(lvl.sel); break;
      case 'Escape': case 'Backspace': if (stack.length > 1) { stack.pop(); render(); } else close(); break;
      default: {
        const ch = e.key.length === 1 ? e.key.toUpperCase() : '';
        const i = ch && !e.ctrlKey && !e.metaKey && !e.altKey ? lvl.keys.indexOf(ch) : -1;
        if (i >= 0) { lvl.sel = i; markSel(); choose(i); } else handled = e.key.length === 1;
      }
    }
    if (handled) { e.preventDefault(); e.stopImmediatePropagation(); }
  };

  overlay.addEventListener('mousedown', (e) => { if (e.target === overlay) close(); });
  document.body.appendChild(overlay);
  open = overlay;
  document.addEventListener('keydown', onKey, true);
  render();
  box.focus();
}

export function slashMenuOpen(): boolean { return !!open; }
