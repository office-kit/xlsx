import { describe, expect, test } from 'vitest';
import { fitGeneralNumber, generalText } from './display.ts';

describe('General number display', () => {
  // Read from Excel 16 for Mac in a 40-character-wide column.
  test.each([
    [Math.PI, '3.141592654'],
    [1 / 3, '0.333333333'],
    [-1 / 3, '-0.333333333'],
    [-Math.PI, '-3.141592654'],
    [12345678901, '12345678901'],
    [123456789012, '1.23457E+11'],
    [-12345678901, '-12345678901'],
    [(2 / 3) * 1000, '666.6666667'],
    [0.0001, '0.0001'],
    [0.0000001, '0.0000001'],
    [1 / 7 / 10000, '1.42857E-05'],
    [1 / 7 / 100000, '1.42857E-06'],
    [1e15, '1E+15'],
    [1234567.891234, '1234567.891'],
    [0.1 + 0.2, '0.3'],
    [1e100, '1E+100'],
    [-1e-20, '-1E-20'],
    [99999999999, '99999999999'],
    [99999999999.5, '1E+11'],
    [(1 / 3) * 1e10, '3333333333'],
  ])('%s shows as %s', (n, text) => {
    expect(generalText(n)).toBe(text);
  });

  test('a narrow column drops to scientific without trailing zeros', () => {
    const measure = (s: string) => s.length * 7;
    expect(fitGeneralNumber(1e15, 5 * 7, measure)).toBe('1E+15');
    expect(fitGeneralNumber(123456789, 7 * 7, measure)).toBe('1.2E+08');
  });
});
