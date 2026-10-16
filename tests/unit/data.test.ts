import { describe, it, expect } from 'vitest';
import { inferFlashFill, runFlashFill } from '../../src/commands/dataops';

describe('flash fill quick', () => {
  it('first name', () => {
    const p = inferFlashFill([{ src: ['John Smith'], out: 'John' }])!;
    expect(runFlashFill(p, ['Mary Jones'])).toBe('Mary');
  });
  it('last, first', () => {
    const p = inferFlashFill([{ src: ['John Smith'], out: 'Smith, John' }])!;
    expect(runFlashFill(p, ['Mary Jones'])).toBe('Jones, Mary');
  });
  it('initials', () => {
    const p = inferFlashFill([{ src: ['john smith'], out: 'J.S.' }, { src: ['ann lee'], out: 'A.L.' }])!;
    expect(runFlashFill(p, ['bob ray'])).toBe('B.R.');
  });
  it('email', () => {
    const p = inferFlashFill([{ src: ['John', 'Smith'], out: 'john.smith@example.com' }])!;
    expect(runFlashFill(p, ['Mary', 'Jones'])).toBe('mary.jones@example.com');
  });
  it('domain', () => {
    const p = inferFlashFill([{ src: ['a.b@foo.com'], out: 'foo.com' }, { src: ['x@bar.org'], out: 'bar.org' }])!;
    expect(runFlashFill(p, ['q@zed.net'])).toBe('zed.net');
  });
  it('upper', () => {
    const p = inferFlashFill([{ src: ['hello world'], out: 'HELLO WORLD' }])!;
    expect(runFlashFill(p, ['abc def'])).toBe('ABC DEF');
  });
});
