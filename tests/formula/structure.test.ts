import { describe, expect, it } from 'vitest';
import { adjustFormulaForRowEdit, type RowEdit } from '../../src/formula/structure.js';

const insert = (at: number, count = 1, sheet = 'S'): RowEdit => ({ sheet, at, count });
const remove = (at: number, count = 1, sheet = 'S'): RowEdit => ({ sheet, at, count: -count });

describe('adjustFormulaForRowEdit — insertion', () => {
  it.each([
    ['A5', 'A6'],
    ['$A$5', '$A$6'],
    ['A$5+$A5', 'A$6+$A6'],
    ['A4', 'A4'],
    ['SUM(E10:E13)', 'SUM(E11:E14)'],
    ['SUM(E5:E10)', 'SUM(E6:E11)'],
    ['SUM(E1:E4)', 'SUM(E1:E4)'],
    ['SUM($E$1:$E$5)', 'SUM($E$1:$E$6)'],
    ['SUM(E9:E2)', 'SUM(E10:E2)'],
    ['SUM(5:7)', 'SUM(6:8)'],
    ['SUM($3:$5)', 'SUM($3:$6)'],
    ['SUM(A:A)', 'SUM(A:A)'],
    ['SUM(A5:A1048576)', 'SUM(A6:A1048576)'],
  ])('%s → %s (insert 1 row at 5)', (input, output) => {
    expect(adjustFormulaForRowEdit(input, 'S', insert(5))).toBe(output);
  });

  it('grows a range that straddles the insertion by the inserted count', () => {
    expect(adjustFormulaForRowEdit('SUM(E10:E13)', 'S', insert(12, 3))).toBe('SUM(E10:E16)');
  });

  it('pushes a reference off the grid to #REF!', () => {
    expect(adjustFormulaForRowEdit('A1048576+A1', 'S', insert(2))).toBe('#REF!+A1');
  });

  it('follows the sheet named in the reference, in any case and quoting', () => {
    expect(adjustFormulaForRowEdit("'My S'!A5+S!A5+s!A5+Other!A5", 'Other', insert(5, 1, 'My S'))).toBe(
      "'My S'!A6+S!A5+s!A5+Other!A5",
    );
    expect(adjustFormulaForRowEdit('S!A5+s!A5+A5', 'Other', insert(5))).toBe('S!A6+s!A6+A5');
    expect(adjustFormulaForRowEdit("'It''s'!A5", 'Other', insert(5, 1, "It's"))).toBe("'It''s'!A6");
  });

  it('leaves 3-D, external, structured and name references alone', () => {
    for (const f of ['SUM(S:T!A5)', "SUM('S:T'!A5)", '[1]S!A5', 'Table1[Col]', 'MyName', 'XFE5', 'A2000000']) {
      expect(adjustFormulaForRowEdit(f, 'S', insert(1))).toBe(f);
    }
  });

  it('keeps text literals and spacing untouched', () => {
    expect(adjustFormulaForRowEdit('IF(A5 > 1, "A5",  A6 )', 'S', insert(5))).toBe('IF(A6 > 1, "A5",  A7 )');
  });

  it('returns the very same string when nothing moves', () => {
    const f = 'SUM(A1:A3)';
    expect(adjustFormulaForRowEdit(f, 'S', insert(10))).toBe(f);
    expect(adjustFormulaForRowEdit('Other!A5', 'T', insert(1))).toBe('Other!A5');
  });

  it('ignores unprefixed references when there is no host sheet', () => {
    expect(adjustFormulaForRowEdit('S!$A$5:$B$9,A5', undefined, insert(5))).toBe('S!$A$6:$B$10,A5');
  });
});

describe('adjustFormulaForRowEdit — deletion', () => {
  it.each([
    ['A4', 'A4'],
    ['A5', '#REF!'],
    ['A6', '#REF!'],
    ['A7', 'A5'],
    ['S!A6', 'S!#REF!'],
    ['SUM(A1:A10)', 'SUM(A1:A8)'],
    ['SUM(A5:A6)', 'SUM(#REF!)'],
    ['SUM(A6:A10)', 'SUM(A5:A8)'],
    ['SUM(A1:A5)', 'SUM(A1:A4)'],
    ['SUM(4:6)', 'SUM(4:4)'],
    ['SUM(A:A)', 'SUM(A:A)'],
  ])('%s → %s (delete rows 5..6)', (input, output) => {
    expect(adjustFormulaForRowEdit(input, 'S', remove(5, 2))).toBe(output);
  });
});
