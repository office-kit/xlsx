import { describe, expect, it } from 'vitest';
import {
  CELL_REF_RE,
  COL_RANGE_RE,
  makeTranslator,
  ROW_RANGE_RE,
  stripWsName,
  TranslatorError,
  translateCol,
  translateFormula,
  translateRange,
  translateRow,
  translatorRender,
} from '../../src/formula/translate.js';
import { MAX_ROW } from '../../src/utils/coordinate.js';

// --- regex tests (mirrors openpyxl test_translate.test_*_re) ---------------

describe('range / cell regexes', () => {
  it.each([
    ['1:1', ['1', '1']],
    ['1234:5678', ['1234', '5678']],
    ['$1234:78910', ['$1234', '78910']],
    ['$12321:$23432', ['$12321', '$23432']],
    ['112233:$445566', ['112233', '$445566']],
    ['A:A', null],
    ['$ABC:AZZ', null],
    ['$DEF:$FOV', null],
    ['HA:$JA', null],
    ['named1', null],
    ['A15', null],
    ['$AB303', null],
    ['YY$101', null],
    ['$ZZ$99', null],
    ['B2:C3', null],
    ['$ATV25:$BBC35', null],
    ['WWW$918:WWW$919', null],
    ['$III$305:$IIT$503', null],
  ] as ReadonlyArray<readonly [string, ReadonlyArray<string> | null]>)('ROW_RANGE_RE %s → %j', (s, expected) => {
    const m = ROW_RANGE_RE.exec(s);
    if (expected === null) expect(m).toBeNull();
    else expect([m?.[1], m?.[2]]).toEqual([expected[0], expected[1]]);
  });

  it.each([
    ['1:1', null],
    ['A:A', ['A', 'A']],
    ['$ABC:AZZ', ['$ABC', 'AZZ']],
    ['$DEF:$FOV', ['$DEF', '$FOV']],
    ['HA:$JA', ['HA', '$JA']],
    ['A15', null],
    ['B2:C3', null],
    ['$ATV25:$BBC35', null],
  ] as ReadonlyArray<readonly [string, ReadonlyArray<string> | null]>)('COL_RANGE_RE %s → %j', (s, expected) => {
    const m = COL_RANGE_RE.exec(s);
    if (expected === null) expect(m).toBeNull();
    else expect([m?.[1], m?.[2]]).toEqual([expected[0], expected[1]]);
  });

  it.each([
    ['1:1', null],
    ['A:A', null],
    ['A15', ['A', '15']],
    ['$AB303', ['$AB', '303']],
    ['YY$101', ['YY', '$101']],
    ['$ZZ$99', ['$ZZ', '$99']],
    ['B2:C3', null],
  ] as ReadonlyArray<readonly [string, ReadonlyArray<string> | null]>)('CELL_REF_RE %s → %j', (s, expected) => {
    const m = CELL_REF_RE.exec(s);
    if (expected === null) expect(m).toBeNull();
    else expect([m?.[1], m?.[2]]).toEqual([expected[0], expected[1]]);
  });
});

// --- translateRow / translateCol (mirrors openpyxl test_translate_row/col) -

describe('translateRow', () => {
  it.each([
    ['1', 1, '2'],
    ['$222333', 1, '$222333'],
    ['1048576', -100, '1048476'],
    ['$1012023', -100, '$1012023'],
    ['101', 0, '101'],
    ['$101', 0, '$101'],
    ['$12', -15, '$12'],
  ])('%s + %s → %s', (input, delta, expected) => {
    expect(translateRow(input, delta as number)).toBe(expected);
  });

  it('throws when shift would underflow', () => {
    expect(() => translateRow('12', -15)).toThrowError(TranslatorError);
  });

  it('throws when shift would overflow past the last row', () => {
    // translateCol has always raised at both ends, because there are no column
    // letters past XFD to hand back. This direction used to produce A1048581,
    // a reference no spreadsheet can resolve.
    expect(() => translateRow('1048576', 1)).toThrowError(TranslatorError);
    expect(() => translateRow('1', MAX_ROW)).toThrowError(TranslatorError);
  });

  it('still allows a shift that lands on the last row', () => {
    expect(translateRow('1048575', 1)).toBe('1048576');
    expect(translateRow('$1048576', 5)).toBe('$1048576');
  });
});

describe('translateCol', () => {
  it.each([
    ['A', 1, 'B'],
    ['XED', 26, 'XFD'],
    ['$XED', 26, '$XED'],
    ['WWW', -52, 'WUW'],
    ['$WWW', -52, '$WWW'],
    ['ABC', 0, 'ABC'],
    ['$ABC', 0, '$ABC'],
    ['$AA', -100, '$AA'],
  ])('%s + %s → %s', (input, delta, expected) => {
    expect(translateCol(input, delta as number)).toBe(expected);
  });

  it('throws when shift would underflow', () => {
    expect(() => translateCol('AA', -100)).toThrowError(TranslatorError);
  });

  it('throws when shift would overflow past the last column', () => {
    expect(() => translateCol('XFD', 1)).toThrowError(TranslatorError);
  });
});

describe('translateFormula against the grid edges', () => {
  it('refuses to shift a reference off the bottom', () => {
    expect(() => translateFormula('=A1048576', 'A1', { rowDelta: 5 })).toThrowError(TranslatorError);
  });

  it('refuses to shift a row range off the bottom', () => {
    expect(() => translateFormula('=SUM(1048570:1048576)', 'A1', { rowDelta: 10 })).toThrowError(
      TranslatorError,
    );
  });

  it('leaves an absolute row anchored, however far the destination is', () => {
    expect(translateFormula('=A$1', 'A1', { rowDelta: MAX_ROW - 1 })).toBe('=A$1');
  });
});

// --- stripWsName ----------------------------------------------------------

describe('stripWsName', () => {
  it.each([
    ['A$3', ['', 'A$3']],
    ['Pipeline!B$4:B$138', ['Pipeline!', 'B$4:B$138']],
    ["'Summary slices'!$C$3", ["'Summary slices'!", '$C$3']],
    ["'Lions! Tigers! Bears!'!$OM$1", ["'Lions! Tigers! Bears!'!", '$OM$1']],
    ['named_range_1', ['', 'named_range_1']],
    ['Sheet-2!named_range_2', ['Sheet-2!', 'named_range_2']],
  ] as ReadonlyArray<readonly [string, readonly [string, string]]>)('%s', (input, expected) => {
    expect(stripWsName(input)).toEqual(expected);
  });
});

// --- translateRange (mirrors openpyxl test_translate_range) ---------------

describe('translateRange', () => {
  it.each([
    ['1:1', 2, 1, '3:3'],
    ['$1234:78910', 1, 10, '$1234:78911'],
    ['$12321:$23432', 3, 5, '$12321:$23432'],
    ['112233:$445566', -3, -20, '112230:$445566'],
    ['987:999', 0, 12, '987:999'],
    ['A:A', 0, 1, 'B:B'],
    ['$ABC:AZZ', 1, 3, '$ABC:BAC'],
    ['$DEF:$FOV', 25, 25, '$DEF:$FOV'],
    ['HA:$JA', -5, -15, 'GL:$JA'],
    ['named1', -33, 33, 'named1'],
    ['A15', -3, 4, 'E12'],
    ['$AB303', 3, 2, '$AB306'],
    ['YY$101', 4, 2, 'ZA$101'],
    ['$ZZ$99', 5, 2, '$ZZ$99'],
    ['B2:C3', 4, 3, 'E6:F7'],
    ['$ATV25:$BBC35', 5, 3, '$ATV30:$BBC40'],
    ['WWW$918:WWW$919', 5, 4, 'WXA$918:WXA$919'],
    ['$III$305:$IIT$503', 25, 35, '$III$305:$IIT$503'],
  ])('%s + (%s,%s) → %s', (input, rd, cd, expected) => {
    expect(translateRange(input, rd as number, cd as number)).toBe(expected);
  });

  it('row-range fall-off raises', () => {
    expect(() => translateRange('1:5', -2, 3)).toThrowError(TranslatorError);
  });
});

// --- translateFormula end-to-end ------------------------------------------

describe('translateFormula', () => {
  it.each([
    ['=IF(A$3<40%,"",INDEX(Pipeline!B$4:B$138,#REF!))', 'A1', 'B2', '=IF(B$3<40%,"",INDEX(Pipeline!C$4:C$138,#REF!))'],
    ["='Summary slices'!$C$3", 'A1', 'B2', "='Summary slices'!$C$3"],
    ['=-MAX(Pipeline!AA4:AA138)', 'A1', 'B2', '=-MAX(Pipeline!AB5:AB139)'],
    [
      '=TEXT(-\'External Ref\'!K7/DENOMINATOR,"$#,##0""M""")',
      'A1',
      'B2',
      '=TEXT(-\'External Ref\'!L8/DENOMINATOR,"$#,##0""M""")',
    ],
    ["=ROWS('Sh 1'!$1:3)+COLUMNS('Sh 2'!$A:C)", 'A1', 'B2', "=ROWS('Sh 1'!$1:4)+COLUMNS('Sh 2'!$A:D)"],
    ['Just text', 'A1', 'B2', 'Just text'],
    ['123.456', 'A1', 'B2', '123.456'],
    ['31/12/1999', 'A1', 'B2', '31/12/1999'],
    ['', 'A1', 'B2', ''],
  ])('translate %s from %s → %s', (formula, origin, dest, expected) => {
    expect(translateFormula(formula, origin, { dest })).toBe(expected);
  });

  it('explicit row/col delta', () => {
    expect(translateFormula("='Summary slices'!C3", 'A1', { rowDelta: 2, colDelta: 3 })).toBe("='Summary slices'!F5");
  });
});

// --- makeTranslator -------------------------------------------------------

describe('makeTranslator', () => {
  it.each([
    ['A1', 1, 1],
    ['AA1', 1, 27],
    ['AA1001', 1001, 27],
    ['XFD111', 111, 16384],
  ])('parses origin %s → row %s col %s', (origin, row, col) => {
    const t = makeTranslator('=formula', origin);
    expect(t.row).toBe(row);
    expect(t.col).toBe(col);
    expect(t.formula).toBe('=formula');
  });

  it('translatorRender returns the formula string', () => {
    const t = makeTranslator('=A1+B2', 'A1');
    expect(translatorRender(t)).toBe('=A1+B2');
  });
});
