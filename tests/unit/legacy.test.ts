import { describe, it, expect, vi } from 'vitest';

describe('carry-over from the former name', () => {
  it('copies quilllattice.* settings to quillgrid.* once, without overwriting', async () => {
    localStorage.clear();
    localStorage.setItem('quilllattice.prefs.v1', '{"firstRun":false}'); // rename:keep
    localStorage.setItem('quilllattice.ai.v1', '{"active":"device"}'); // rename:keep
    localStorage.setItem('quillgrid.ai.v1', '{"active":"groq"}');
    vi.resetModules();
    await import('../../src/util/legacy');
    expect(localStorage.getItem('quillgrid.prefs.v1')).toBe('{"firstRun":false}');
    expect(localStorage.getItem('quillgrid.ai.v1')).toBe('{"active":"groq"}');
    expect(localStorage.getItem('quilllattice.prefs.v1')).toBe('{"firstRun":false}'); // rename:keep — old entries are kept
    // Runs once: a later old-name value is not copied again.
    localStorage.setItem('quilllattice.ink', 'x'); // rename:keep
    vi.resetModules();
    await import('../../src/util/legacy');
    expect(localStorage.getItem('quillgrid.ink')).toBeNull();
  });
});
