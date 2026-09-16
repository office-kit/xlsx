import { describe, expect, it } from 'vitest';
import {
  bindValue,
  type Cell,
  getCoordinate,
  isEmptyCell,
  isFormulaCell,
  isRichTextCell,
  makeArrayFormula,
  makeCell,
  makeDataTableFormula,
  makeDurationValue,
  makeErrorValue,
  makeFormula,
  makeSharedFormula,
  setArrayFormula,
  setCellValue,
  setDataTableFormula,
  setFormula,
  setSharedFormula,
} from '../../src/cell/cell.js';
import { makeRichText, makeTextRun } from '../../src/cell/rich-text.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';

describe('makeCell / getCoordinate', () => {
  it('builds a Cell with default styleId 0 and null value', () => {
    const c = makeCell(1, 1);
    expect(c.row).toBe(1);
    expect(c.col).toBe(1);
    expect(c.value).toBeNull();
    expect(c.styleId).toBe(0);
  });

  it('getCoordinate returns "A1" for (1,1)', () => {
    expect(getCoordinate(makeCell(1, 1))).toBe('A1');
    expect(getCoordinate(makeCell(1, 27))).toBe('AA1');
    expect(getCoordinate(makeCell(1048576, 16384))).toBe('XFD1048576');
  });

  it('rejects out-of-range row / col', () => {
    expect(() => makeCell(0, 1)).toThrowError(OpenXmlSchemaError);
    expect(() => makeCell(1, 0)).toThrowError(OpenXmlSchemaError);
    expect(() => makeCell(1, 16385)).toThrowError(OpenXmlSchemaError);
    expect(() => makeCell(1048577, 1)).toThrowError(OpenXmlSchemaError);
    expect(() => makeCell(1.5, 1)).toThrowError(OpenXmlSchemaError);
  });

  it('Cell stays mutable for hot-path performance', () => {
    const c = makeCell(1, 1);
    c.value = 42;
    c.styleId = 5;
    expect(c.value).toBe(42);
    expect(c.styleId).toBe(5);
  });
});

describe('setCellValue', () => {
  it('passes plain primitives through verbatim', () => {
    const c = makeCell(1, 1);
    setCellValue(c, 42);
    expect(c.value).toBe(42);
    setCellValue(c, 'hi');
    expect(c.value).toBe('hi');
    setCellValue(c, true);
    expect(c.value).toBe(true);
    setCellValue(c, null);
    expect(c.value).toBeNull();
  });
});

describe('bindValue (type-inferring setter)', () => {
  it('"=…" strings become formulas', () => {
    const c = makeCell(1, 1);
    bindValue(c, '=SUM(A1:A10)');
    expect(isFormulaCell(c)).toBe(true);
    const v = c.value as { kind: 'formula'; formula: string; t: string };
    expect(v.formula).toBe('SUM(A1:A10)');
    expect(v.t).toBe('normal');
  });

  it('error tokens become error values', () => {
    const c = makeCell(1, 1);
    bindValue(c, '#DIV/0!');
    expect(c.value).toEqual({ kind: 'error', code: '#DIV/0!' });
  });

  it('non-special strings stay strings', () => {
    const c = makeCell(1, 1);
    bindValue(c, 'hello world');
    expect(c.value).toBe('hello world');
  });

  it('numbers / booleans / Dates / null pass through', () => {
    const c = makeCell(1, 1);
    bindValue(c, 3.14);
    expect(c.value).toBe(3.14);
    bindValue(c, false);
    expect(c.value).toBe(false);
    const d = new Date(0);
    bindValue(c, d);
    expect(c.value).toBe(d);
    bindValue(c, null);
    expect(c.value).toBeNull();
  });
});

describe('formula setters', () => {
  it('setFormula sets a normal formula with optional cached value', () => {
    const c = makeCell(1, 1);
    setFormula(c, 'A1+B1', { cachedValue: 42 });
    expect(c.value).toEqual({ kind: 'formula', t: 'normal', formula: 'A1+B1', cachedValue: 42 });
  });

  it('setArrayFormula carries the ref range', () => {
    const c = makeCell(1, 1);
    setArrayFormula(c, 'A1:B2', 'A1*B1');
    expect(c.value).toEqual({ kind: 'formula', t: 'array', formula: 'A1*B1', ref: 'A1:B2' });
  });

  it('every setter assigns the value its constructor builds', () => {
    const c = makeCell(1, 1);

    setFormula(c, 'A1+B1', { cachedValue: 42 });
    expect(c.value).toEqual(makeFormula('A1+B1', { cachedValue: 42 }));

    setArrayFormula(c, 'A1:B2', 'A1*B1');
    expect(c.value).toEqual(makeArrayFormula('A1:B2', 'A1*B1'));

    setSharedFormula(c, 0, 'A1+1', 'A1:A10');
    expect(c.value).toEqual(makeSharedFormula(0, 'A1+1', 'A1:A10'));

    setDataTableFormula(c, 'TABLE(B1,C1)', { ref: 'A1:A3', r1: '$B$1' });
    expect(c.value).toEqual(makeDataTableFormula('TABLE(B1,C1)', { ref: 'A1:A3', r1: '$B$1' }));
  });

  it('setSharedFormula validates si and accepts optional formula / ref', () => {
    const c = makeCell(1, 1);
    setSharedFormula(c, 0, 'A1+1', 'A1:A10');
    expect(c.value).toEqual({ kind: 'formula', t: 'shared', formula: 'A1+1', si: 0, ref: 'A1:A10' });

    setSharedFormula(c, 1);
    expect(c.value).toEqual({ kind: 'formula', t: 'shared', formula: '', si: 1 });
    expect(() => setSharedFormula(c, -1)).toThrowError(OpenXmlSchemaError);
    expect(() => setSharedFormula(c, 1.5)).toThrowError(OpenXmlSchemaError);
  });
});

describe('value-shape helpers', () => {
  it('makeErrorValue rejects unknown codes', () => {
    expect(makeErrorValue('#REF!')).toEqual({ kind: 'error', code: '#REF!' });
    // biome-ignore lint/suspicious/noExplicitAny: bad input on purpose
    expect(() => makeErrorValue('#NOPE!' as any)).toThrowError(OpenXmlSchemaError);
  });

  it('makeDurationValue rejects non-finite ms', () => {
    expect(makeDurationValue(1000)).toEqual({ kind: 'duration', ms: 1000 });
    expect(() => makeDurationValue(Number.NaN)).toThrowError(OpenXmlSchemaError);
  });

  it('isFormulaCell / isRichTextCell / isEmptyCell discriminate', () => {
    const c = makeCell(1, 1);
    expect(isEmptyCell(c)).toBe(true);

    setFormula(c, 'A1');
    expect(isFormulaCell(c)).toBe(true);
    expect(isRichTextCell(c)).toBe(false);

    c.value = { kind: 'rich-text', runs: makeRichText([makeTextRun('hi')]) };
    expect(isRichTextCell(c)).toBe(true);
    expect(isFormulaCell(c)).toBe(false);

    c.value = 42;
    expect(isFormulaCell(c)).toBe(false);
    expect(isRichTextCell(c)).toBe(false);
    expect(isEmptyCell(c)).toBe(false);
  });
});

describe('Cell mutation patterns', () => {
  it('worksheet write loop sets value + style without spread / freeze cost', () => {
    const cells: Cell[] = [];
    for (let i = 1; i <= 100; i++) {
      const c = makeCell(i, 1);
      c.value = i;
      c.styleId = i % 5;
      cells.push(c);
    }
    const first = cells[0];
    if (first === undefined) throw new Error('expected at least one cell');
    expect(first.value).toBe(1);
    expect(cells[99]?.value).toBe(100);
    // Mutating in place is the canonical hot-path pattern.
    first.value = 999;
    expect(first.value).toBe(999);
  });
});
