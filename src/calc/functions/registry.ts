/**
 * Worksheet function registry.
 *
 * Each function declares its arity and which arguments accept arrays/references. Arguments in scalar
 * positions are "lifted": when given a multi-cell range or array, the function is applied element-wise
 * and returns an array (dynamic-array semantics). Implementations receive evaluated values and may
 * throw via `fail()` (see ../values) to return an error.
 */
import type { Ctx, Value } from '../values';
import type { Node } from '../parser';

export type FnCategory = 'Math & Trig' | 'Statistical' | 'Text' | 'Logical' | 'Lookup & Reference' | 'Date & Time' | 'Financial' | 'Information' | 'Engineering' | 'Database' | 'Web' | 'Compatibility' | 'Cube' | 'Dynamic Array';

export interface FnDef {
  name: string;
  min: number;
  max: number;
  category: FnCategory;
  /** Short description for the Insert Function dialog and formula autocomplete. */
  desc: string;
  /** Argument names, e.g. ['number1', '[number2]', '...']. */
  params: string[];
  /** Argument indices that accept arrays or references without lifting. 'all' = every argument. */
  arrays?: number[] | 'all';
  /** Argument indices that must be (and stay) references, e.g. OFFSET(reference, ...). Implies arrays. */
  refs?: number[];
  /** Lazy functions receive unevaluated nodes and evaluate them through ctx.evalNode (IF, LET, LAMBDA...). */
  lazy?: (nodes: Node[], ctx: Ctx) => Value;
  impl?: (args: Value[], ctx: Ctx) => Value;
  volatile?: boolean;
  /** Function may return a reference (INDEX, OFFSET, INDIRECT, CHOOSE...). */
  returnsRef?: boolean;
  /** Other names for compatibility (e.g. STDEV -> STDEV.S). */
  aliases?: string[];
}

export const FUNCTIONS = new Map<string, FnDef>();

export function defFn(d: FnDef | FnDef[]): void {
  for (const f of Array.isArray(d) ? d : [d]) {
    const name = f.name.toUpperCase();
    if (FUNCTIONS.has(name)) console.warn(`Duplicate function ${name}`);
    FUNCTIONS.set(name, f);
    for (const a of f.aliases || []) FUNCTIONS.set(a.toUpperCase(), f);
  }
}

export function getFn(name: string): FnDef | undefined { return FUNCTIONS.get(name.toUpperCase()); }

export function isArrayArg(f: FnDef, i: number): boolean {
  if (f.arrays === 'all') return true;
  if (f.arrays?.includes(i)) return true;
  if (f.refs?.includes(i)) return true;
  // Variadic tail: if the last declared array index is the variadic marker, apply to the rest.
  if (Array.isArray(f.arrays) && f.arrays.includes(-1) && i >= Math.max(...f.arrays.filter((x) => x >= 0), 0)) return true;
  return false;
}

/** Unique function names (aliases excluded), sorted. */
export function functionNames(): string[] {
  const seen = new Set<string>();
  for (const f of FUNCTIONS.values()) seen.add(f.name.toUpperCase());
  return [...seen].sort();
}
