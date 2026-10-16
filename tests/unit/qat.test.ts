import { describe, it, expect } from 'vitest';
import '../../src/commands/index';
import { getCommand } from '../../src/commands/registry';
import { QAT_STANDARD, QAT_CHOICES, DEFAULT_QAT, addAllQatChoices } from '../../src/prefs';

const label = (id: string) => getCommand(id)?.label;

describe('Quick Access Bar choices', () => {
  it('every drop-down choice is a registered command', () => {
    for (const id of QAT_CHOICES) expect(getCommand(id), id).toBeTruthy();
  });

  it('covers the desktop drop-down list', () => {
    for (const id of ['file.autosave', 'file.new', 'file.open', 'file.save', 'file.email', 'file.quickPrint', 'file.print', 'file.printPreview', 'review.spelling', 'edit.undo', 'edit.redo', 'data.sortAZ', 'data.sortZA', 'view.touchMode']) {
      expect(QAT_STANDARD).toContain(id);
    }
  });

  it('Add All puts Save and Print first, then the rest alphabetically, then other existing items', () => {
    const out = addAllQatChoices([...DEFAULT_QAT, '|', 'fmt.bold'], label);
    expect(out.slice(0, 2)).toEqual(['file.save', 'file.print']);
    const rest = out.slice(2, QAT_CHOICES.length);
    const labels = rest.map((id) => label(id)!);
    expect(labels).toEqual([...labels].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' })));
    expect(new Set(out.filter((x) => x !== '|')).size).toBe(out.filter((x) => x !== '|').length);
    expect(out.slice(QAT_CHOICES.length)).toEqual(['|', 'fmt.bold']);
  });

  it('is idempotent', () => {
    const once = addAllQatChoices(DEFAULT_QAT, label);
    expect(addAllQatChoices(once, label)).toEqual(once);
  });
});
