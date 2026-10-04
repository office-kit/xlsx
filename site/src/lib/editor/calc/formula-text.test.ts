import { describe, expect, test } from 'vitest';
import {
  adjustFormulaForMove,
  adjustFormulaForStructure,
  CalcParseError,
  formulaReferences,
  fromStorageFormula,
  parseFormula,
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
