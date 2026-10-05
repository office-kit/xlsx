import { describe, expect, test } from 'vitest';
import {
  adjustFormulaForMove,
  adjustFormulaForStructure,
  CalcParseError,
  formulaReferences,
  fromStorageFormula,
  parseFormula,
  deleteSheetInFormula,
  renameSheetInFormula,
  toggleReferenceAt,
  toStorageFormula,
  translateFormula,
} from './index.ts';

describe('parseFormula', () => {
  test('unary minus binds tighter than ^, and ^ is left-associative', () => {
    expect(parseFormula('-2^2')).toMatchObject({ type: 'binary', op: '^', left: { type: 'unary' } });
    expect(parseFormula('2^3^2')).toMatchObject({ type: 'binary', op: '^', left: { type: 'binary', op: '^' } });
  });

  test('reads references in every A1 form', () => {
    for (const f of ['A1', '$A$1', 'A1:B2', 'A:A', '$1:$3', 'Sheet1!A1', "'My Sheet'!A1:B2", 'Sheet1:Sheet3!A1', 'A1#', '@A1:A3', 'Table1[[#Headers],[Col 1]]']) {
      expect(() => parseFormula(f), f).not.toThrow();
    }
  });

  test('space is intersection and a parenthesised comma is union', () => {
    expect(parseFormula('A1:B5 B2:C3')).toMatchObject({ type: 'binary', op: ' ' });
    expect(parseFormula('SUM((A1,B2))')).toMatchObject({ type: 'call', args: [{ type: 'binary', op: ',' }] });
  });

  test('strips storage prefixes from function names', () => {
    expect(parseFormula('_xlfn._xlws.SORT(A1:A3)')).toMatchObject({ type: 'call', name: 'SORT' });
  });

  test('reports the position of a syntax error', () => {
    try {
      parseFormula('SUM(1,,');
      expect.unreachable();
    } catch (e) {
      if (!(e instanceof CalcParseError)) throw e;
      expect(e.position).toBe(7);
    }
  });
});

describe('formulaReferences / toggleReferenceAt', () => {
  test('lists reference spans with their sheet', () => {
    const refs = formulaReferences("=SUM(A1:B2,'Q 1'!C3)+x");
    expect(refs.map((r) => [r.text, r.sheet, r.range])).toEqual([
      ['A1:B2', undefined, { r1: 1, c1: 1, r2: 2, c2: 2 }],
      ["'Q 1'!C3", 'Q 1', { r1: 3, c1: 3, r2: 3, c2: 3 }],
    ]);
  });

  test('cycles $ flags of the reference under the caret', () => {
    let state = { text: '=A1+B2', caret: 2 };
    const seen: string[] = [];
    for (let i = 0; i < 4; i++) {
      state = toggleReferenceAt(state.text, state.caret);
      seen.push(state.text);
    }
    expect(seen).toEqual(['=$A$1+B2', '=A$1+B2', '=$A1+B2', '=A1+B2']);
  });
});

describe('translateFormula', () => {
  test('moves relative parts only', () => {
    expect(translateFormula('A1+$B$2+C$3+$D4', 2, 1)).toBe('B3+$B$2+D$3+$D6');
  });

  test('off-grid references become #REF! and keep their sheet', () => {
    expect(translateFormula('Sheet2!A1*2', -1, 0)).toBe('Sheet2!#REF!*2');
  });

  test('leaves strings and function names alone', () => {
    expect(translateFormula('LOG10(A1)&"B2"', 1, 0)).toBe('LOG10(A2)&"B2"');
  });
});

describe('adjustFormulaForStructure', () => {
  const insertRows = { sheet: 'S', axis: 'row', at: 3, count: 2 } as const;

  test('shifts references at or after the insertion', () => {
    expect(adjustFormulaForStructure('A2+A3+SUM(A1:A5)', 'S', insertRows)).toBe('A2+A5+SUM(A1:A7)');
  });

  test('only touches the edited sheet; unqualified means the formula sheet', () => {
    expect(adjustFormulaForStructure('A3+T!A3', 'T', insertRows)).toBe('A3+T!A3');
    expect(adjustFormulaForStructure('A3+S!A3', 'T', insertRows)).toBe('A3+S!A5');
  });

  test('deleted references become #REF! and ranges shrink', () => {
    const del = { sheet: 'S', axis: 'row', at: 2, count: -2 } as const;
    expect(adjustFormulaForStructure('A2+A4+SUM(A1:A5)', 'S', del)).toBe('#REF!+A2+SUM(A1:A3)');
  });

  test('a band shifts only references wholly inside it', () => {
    const edit = { sheet: 'S', axis: 'row', at: 2, count: 1, band: { from: 1, to: 2 } } as const;
    expect(adjustFormulaForStructure('A2+B2+C2+A2:B3+A2:C3', 'S', edit)).toBe('A3+B3+C2+A3:B4+A2:C3');
  });
});

describe('adjustFormulaForMove', () => {
  const move = { sheet: 'S', range: { r1: 1, c1: 1, r2: 2, c2: 2 }, toSheet: 'S', dRow: 0, dCol: 3 } as const;

  test('references wholly inside the moved block follow it', () => {
    expect(adjustFormulaForMove('A1+SUM(A1:B2)', 'S', move)).toBe('D1+SUM(D1:E2)');
  });

  test('partially overlapping references stay', () => {
    expect(adjustFormulaForMove('SUM(A1:C3)', 'S', move)).toBe('SUM(A1:C3)');
  });

  test('references into the overwritten destination become #REF!', () => {
    expect(adjustFormulaForMove('E2+F2', 'S', move)).toBe('#REF!+F2');
  });

  test('a move to another sheet qualifies the reference', () => {
    expect(adjustFormulaForMove('A1', 'S', { ...move, toSheet: 'T 2', dCol: 0 })).toBe("'T 2'!A1");
  });
});

test('deleteSheetInFormula turns references to the sheet into #REF!', () => {
  const order = ['A', 'Gone', 'C', 'D'];
  expect(deleteSheetInFormula("Gone!A1+'gone'!B2:C3+C!A1+SUM(Gone!X)", 'Gone', order)).toBe('#REF!+#REF!+C!A1+SUM(#REF!)');
  expect(deleteSheetInFormula('"Gone!A1"&A1', 'Gone', order)).toBe('"Gone!A1"&A1');
});

test('deleteSheetInFormula shrinks a 3-D reference ending on the sheet', () => {
  const order = ['Jan', 'Gone', 'Mar', 'Apr'];
  expect(deleteSheetInFormula('SUM(Gone:Apr!A1)', 'Gone', order)).toBe('SUM(Mar:Apr!A1)');
  expect(deleteSheetInFormula('SUM(Jan:Gone!A1)', 'Gone', order)).toBe('SUM(Jan!A1)');
  expect(deleteSheetInFormula('SUM(Jan:Mar!A1)', 'Gone', order)).toBe('SUM(Jan:Mar!A1)');
});

test('renameSheetInFormula quotes when needed', () => {
  expect(renameSheetInFormula('Old!A1+old!B1+Other!A1', 'Old', 'New Name')).toBe("'New Name'!A1+'New Name'!B1+Other!A1");
});

describe('storage prefixes', () => {
  test('adds _xlfn / _xlws / _xlpm the way Excel writes them', () => {
    expect(toStorageFormula('LET(x,1,SORT(XLOOKUP(x,A:A,B:B)))+SUM(1)')).toBe('_xlfn.LET(_xlpm.x,1,_xlfn._xlws.SORT(_xlfn.XLOOKUP(_xlpm.x,A:A,B:B)))+SUM(1)');
  });

  test('round-trips', () => {
    const f = 'MAP(A1:A3,LAMBDA(v,v*2))+IFS(TRUE,1)';
    expect(fromStorageFormula(toStorageFormula(f))).toBe(f);
  });
});

// Each pair was typed into Excel for Mac and read back from the cell's formula.
describe('entry normalisation', () => {
  test.each([
    ['SUM(A10:A3)', 'SUM(A3:A10)'],
    ['SUM(B3:A1)', 'SUM(A1:B3)'],
    ['SUM(3:1)', 'SUM(1:3)'],
    ['SUM(C:A)', 'SUM(A:C)'],
    ['$b$2:a1', 'A1:$B$2'],
    ['$A3:B$1', '$A$1:B3'],
    ['B$1:$A3', '$A$1:B3'],
    ['A1:A1', 'A1:A1'],
    ["'S'!A1", 'S!A1'],
    ["'My Sheet'!a1", "'My Sheet'!A1"],
    ['if(true,1,2)', 'IF(TRUE,1,2)'],
    ['1+false', '1+FALSE'],
    ['1E3', '1000'],
    ['1.50', '1.5'],
    ['.5', '0.5'],
    ['0005', '5'],
    ['1e-5', '0.00001'],
    ['1.0E+3', '1000'],
    ['0.1E1', '1'],
    ['1E20', '100000000000000000000'],
    ['1E21', '1E+21'],
    ['1.5E300', '1.5E+300'],
    ['1E-19', '0.0000000000000000001'],
    ['1E-20', '1E-20'],
    ['1.23456789E-12', '1.23456789E-12'],
    ['123456789012345678', '123456789012345000'],
    ['12345678901234567890123', '1.23456789012345E+22'],
    ['0.1234567890123456789', '0.123456789012345'],
    ['"a"&"b"', '"a"&"b"'],
  ])('%s → %s', (typed, stored) => {
    expect(toStorageFormula(typed)).toBe(stored);
  });
});

test('a typed sheet name takes the real sheet name\'s case', () => {
  expect(toStorageFormula("sheet1!a1+'my data'!b2+'SHEET1'!C3", ['Sheet1', 'My Data'])).toBe("Sheet1!A1+'My Data'!B2+Sheet1!C3");
  // An unknown sheet stays as typed (Excel asks for the file instead).
  expect(toStorageFormula('other!a1', ['Sheet1'])).toBe('other!A1');
});
