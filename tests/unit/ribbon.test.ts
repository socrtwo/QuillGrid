import { describe, it, expect } from 'vitest';
import '../../src/calc/functions';
import '../../src/commands';
import { DEFAULT_TABS, BACKSTAGE_ITEMS } from '../../src/ui/ribbonLayout';
import { getCommand, allCommands } from '../../src/commands/registry';

describe('ribbon', () => {
  it('every ribbon and backstage item is a registered command', () => {
    const missing: string[] = [];
    for (const t of DEFAULT_TABS) for (const g of t.groups) { for (const id of g.items) if (!getCommand(id)) missing.push(`${t.id}/${g.id}/${id}`); if (g.launcher && !getCommand(g.launcher)) missing.push(`${t.id}/${g.id}/launcher:${g.launcher}`); }
    for (const id of BACKSTAGE_ITEMS) if (!getCommand(id)) missing.push(`backstage/${id}`);
    expect(missing).toEqual([]);
  });
  it('commands have labels, icons and tabs', () => {
    const bad = allCommands().filter((c) => !c.label || !c.tab || !c.icon).map((c) => c.id);
    expect(bad).toEqual([]);
    expect(allCommands().length).toBeGreaterThan(600);
  });
});
