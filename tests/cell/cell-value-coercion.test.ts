// Tests for value-level type guards + cellValueAsString / cellValueAsNumber
// coercion helpers.

import { describe, expect, it } from 'vitest';
import {
  cellValueAsNumber,
  cellValueAsString,
  isDurationValue,
  isErrorValue,
  isFormulaValue,
  isRichTextValue,
  makeDurationValue,
  makeErrorValue,
} from '../../src/cell/cell.js';
import { makeRichText } from '../../src/cell/rich-text.js';

describe('value type guards', () => {
  it('isFormulaValue narrows the formula variant', () => {
    const v = { kind: 'formula' as const, t: 'normal' as const, formula: 'A1+1' };
    expect(isFormulaValue(v)).toBe(true);
    expect(isFormulaValue('plain')).toBe(false);
    expect(isFormulaValue(42)).toBe(false);
    expect(isFormulaValue(null)).toBe(false);
  });

  it('isRichTextValue / isErrorValue / isDurationValue distinguish their kinds', () => {
    expect(isRichTextValue({ kind: 'rich-text', runs: makeRichText([{ text: 'hi' }]) })).toBe(true);
    expect(isErrorValue(makeErrorValue('#NAME?'))).toBe(true);
    expect(isDurationValue(makeDurationValue(1500))).toBe(true);
    expect(isErrorValue({ kind: 'rich-text', runs: makeRichText([]) })).toBe(false);
  });
});

describe('cellValueAsString', () => {
  it('null → empty string', () => {
    expect(cellValueAsString(null)).toBe('');
  });

  it('numbers / booleans / strings pass through via String', () => {
    expect(cellValueAsString(42)).toBe('42');
    expect(cellValueAsString(true)).toBe('true');
    expect(cellValueAsString('hello')).toBe('hello');
  });

  it('Date → ISO string', () => {
    expect(cellValueAsString(new Date('2024-01-02T03:04:05Z'))).toBe('2024-01-02T03:04:05.000Z');
  });

  it('rich text concatenates run text', () => {
    const rt = makeRichText([{ text: 'Hello ' }, { text: 'world', font: { b: true } }]);
    expect(cellValueAsString({ kind: 'rich-text', runs: rt })).toBe('Hello world');
  });

  it('formula with cached value uses the cache; uncached → empty', () => {
    expect(
      cellValueAsString({ kind: 'formula', t: 'normal', formula: 'A1+1', cachedValue: 7 }),
    ).toBe('7');
    expect(cellValueAsString({ kind: 'formula', t: 'normal', formula: 'A1+1' })).toBe('');
  });

  it('error → token, duration → "<ms> ms"', () => {
    expect(cellValueAsString(makeErrorValue('#DIV/0!'))).toBe('#DIV/0!');
    expect(cellValueAsString(makeDurationValue(2500))).toBe('2500 ms');
  });
});

describe('cellValueAsNumber', () => {
  it('numbers / booleans pass through with bool → 0/1', () => {
    expect(cellValueAsNumber(42)).toBe(42);
    expect(cellValueAsNumber(false)).toBe(0);
    expect(cellValueAsNumber(true)).toBe(1);
  });

  it('numeric strings parse; non-numeric strings → undefined', () => {
    expect(cellValueAsNumber('3.14')).toBeCloseTo(3.14);
    expect(cellValueAsNumber('hello')).toBeUndefined();
    expect(cellValueAsNumber('')).toBeUndefined();
  });

  it('formulas only return their numeric cached value', () => {
    expect(
      cellValueAsNumber({ kind: 'formula', t: 'normal', formula: 'A1+1', cachedValue: 7 }),
    ).toBe(7);
    expect(
      cellValueAsNumber({ kind: 'formula', t: 'normal', formula: '"hi"', cachedValue: 'hi' }),
    ).toBeUndefined();
  });

  it('joins rich-text runs, then parses', () => {
    // The docblock has always promised this; the branch was missing, so a
    // number typed into a cell that carries per-run formatting read as
    // undefined. `cellValueAsString` already joins the runs, so reading the
    // same cell numerically has to see the same text.
    expect(cellValueAsNumber({ kind: 'rich-text', runs: [{ text: '42' }] })).toBe(42);
    expect(
      cellValueAsNumber({ kind: 'rich-text', runs: [{ text: '1', font: { b: true } }, { text: '2.5' }] }),
    ).toBe(12.5);
    expect(cellValueAsNumber({ kind: 'rich-text', runs: [{ text: 'hello' }] })).toBeUndefined();
    expect(cellValueAsNumber({ kind: 'rich-text', runs: [] })).toBeUndefined();
    expect(cellValueAsNumber({ kind: 'rich-text', runs: [{ text: '' }] })).toBeUndefined();
  });

  it('null / Date / errors / durations → undefined', () => {
    expect(cellValueAsNumber(null)).toBeUndefined();
    expect(cellValueAsNumber(new Date(0))).toBeUndefined();
    expect(cellValueAsNumber(makeErrorValue('#REF!'))).toBeUndefined();
    expect(cellValueAsNumber(makeDurationValue(100))).toBeUndefined();
  });

  it('NaN / Infinity → undefined', () => {
    expect(cellValueAsNumber(Number.NaN)).toBeUndefined();
    expect(cellValueAsNumber(Number.POSITIVE_INFINITY)).toBeUndefined();
  });
});
