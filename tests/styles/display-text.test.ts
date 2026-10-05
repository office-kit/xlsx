// Table-driven checks for `getCellDisplayText`: a format code and a value in,
// the text Excel prints out. Excel's own output is the expectation throughout;
// rows where Excel's answer is not recoverable from the cell (column-width
// dependent notation, `*` fills) are marked where they appear.

import { describe, expect, it } from 'vitest';
import type { CellValue } from '../../src/cell/cell.js';
import { makeDurationValue, makeErrorValue, makeFormula } from '../../src/cell/cell.js';
import { makeRichText, makeTextRun } from '../../src/cell/rich-text.js';
import { getCellDisplayText, setCellNumberFormat } from '../../src/styles/index.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/index.js';
import { setCell } from '../../src/worksheet/index.js';

const display = (code: string, value: CellValue, opts?: { date1904?: boolean }): string => {
  const wb = createWorkbook({ date1904: opts?.date1904 ?? false });
  const ws = addWorksheet(wb, 'Sheet1');
  const cell = setCell(ws, 1, 1, value);
  setCellNumberFormat(wb, cell, code);
  return getCellDisplayText(wb, cell);
};

/** 2024-03-14, a Thursday, under the Windows epoch. */
const MARCH_14_2024 = 45_365;

describe('getCellDisplayText: the five readings cellValueAsString gets wrong', () => {
  it.each([
    ['0.0%', 0.5, '50.0%'],
    ['#,##0.00', 1_234_567.891, '1,234,567.89'],
    ['General', 1.1 + 2.2, '3.3'],
    ['yyyy-mm-dd', MARCH_14_2024, '2024-03-14'],
    ['General', true, 'TRUE'],
  ] satisfies Array<[string, CellValue, string]>)('%s + %j => %j', (code, value, expected) => {
    expect(display(code, value)).toBe(expected);
  });
});

describe('getCellDisplayText: General', () => {
  it.each([
    [0, '0'],
    [-0.5, '-0.5'],
    [1_234_567.891, '1234567.891'],
    // 15 significant digits is Excel's stored precision, and rounding to it is
    // what removes the binary noise.
    [0.1 + 0.2, '0.3'],
    [1 / 3, '0.333333333333333'],
    [123_456_789_012, '123456789012'],
    [1e21, '1E+21'],
    [1e-7, '1E-07'],
  ])('General + %j => %j', (value, expected) => {
    expect(display('General', value)).toBe(expected);
  });
});

describe('getCellDisplayText: numeric codes', () => {
  it.each([
    ['0', 2.5, '3'],
    ['0', -2.5, '-3'],
    // Excel rounds the 15-digit decimal it shows, so 1.005 goes up even though
    // the nearest double sits below it.
    ['0.00', 1.005, '1.01'],
    ['0.00', 2.675, '2.68'],
    ['0.000', 0.0005, '0.001'],
    ['#,##0', 1_234_567, '1,234,567'],
    ['#,##0', 5, '5'],
    ['000,000', 5, '000,005'],
    ['#,##0.00', -1_234.5, '-1,234.50'],
    ['$#,##0.00', -1_234.5, '-$1,234.50'],
    ['\\$0.00', 12.3, '$12.30'],
    ['"total: "0', 7, 'total: 7'],
    ['0"kg"', 12, '12kg'],
    ['#.##', 0.5, '.5'],
    // No integer placeholder at all: Excel still prints the integer digits,
    // and prints nothing where the integer part is zero.
    ['.00', 123.456, '123.46'],
    ['.00', 0.5, '.50'],
    ['0.##', 1.5, '1.5'],
    // Excel keeps the point even when no fraction digit is shown (checked in Excel).
    ['0.##', 1, '1.'],
    ['0.', 1.5, '2.'],
    ['#,##0.', 1_234.5, '1,235.'],
    ['#.', 0, '.'],
    ['0."x"', 12, '12.x'],
    ['0.;(0.)', -3.2, '(3.)'],
    ['0.%', 0.25, '25.%'],
    ['0.E+0', 123, '1.E+2'],
    ['0.??', 1, '1.  '],
    ['#', 0, ''],
    ['000-00-0000', 123_456_789, '123-45-6789'],
    ['0%', 0.1234, '12%'],
    ['0.00%', -0.1234, '-12.34%'],
    // A trailing comma divides by a thousand, and two of them by a million.
    ['#,##0,, "M"', 1_234_567_890, '1,235 M'],
    ['[Red]#,##0.00', 99.5, '99.50'],
    ['[$€-407]#,##0.00', 1_234.5, '€1,234.50'],
  ] satisfies Array<[string, CellValue, string]>)('%s + %j => %j', (code, value, expected) => {
    expect(display(code, value)).toBe(expected);
  });
});

describe('getCellDisplayText: scientific notation', () => {
  it.each([
    ['0.00E+00', 1_234.5, '1.23E+03'],
    ['0.00E+00', 0.5, '5.00E-01'],
    ['0.00E+00', 0, '0.00E+00'],
    ['0.00E-00', 0.001, '1.00E-03'],
    // Rounding the mantissa past its width steps the exponent instead.
    ['0.0E+00', 9.99, '1.0E+01'],
    // `##0.0E+0` asks for up to three integer digits: engineering notation,
    // exponent in steps of three.
    ['##0.0E+0', 12_345, '12.3E+3'],
    ['##0.0E+0', 0.0123, '12.3E-3'],
    ['##0.0E+0', 0.00123, '1.2E-3'],
    ['##0.0E+0', 9.99, '10.0E+0'],
  ])('%s + %j => %j', (code, value, expected) => {
    expect(display(code, value)).toBe(expected);
  });
});

describe('getCellDisplayText: fractions', () => {
  it.each([
    ['# ?/?', 2.75, '2 3/4'],
    // `??` pads so the slash lines up down the column.
    ['# ??/??', 2.75, '2  3/4 '],
    // The closest fraction under the denominator cap, which for 0.7 is a
    // semiconvergent (5/7) rather than a continued-fraction convergent (2/3).
    ['?/?', 0.7, '5/7'],
    ['# ?/?', 1 / 3, ' 1/3'],
    // An exact whole number blanks the fraction.
    ['# ?/?', 2, '2    '],
    // A spelled-out denominator is used as given, and goes blank with the rest
    // of the fraction when the value is a whole number.
    ['# ?/16', 2.3125, '2 5/16'],
    ['# ?/16', 3, '3     '],
    // `?` pads the denominator on the right so the slash lines up. A `0` pads
    // on the left: a zero on the right would multiply the denominator by ten.
    ['# ??/00', 2.5, '2  1/02'],
    ['?/00', 0.5, '1/02'],
  ])('%s + %j => %j', (code, value, expected) => {
    expect(display(code, value)).toBe(expected);
  });
});

describe('getCellDisplayText: multi-section codes', () => {
  it.each([
    // Two sections split at zero; the negative branch supplies its own sign.
    ['0.00;(0.00)', -1_234.5, '(1234.50)'],
    ['0.00;(0.00)', 1_234.5, '1234.50'],
    ['0.00;(0.00)', 0, '0.00'],
    // Three sections add a dedicated zero branch.
    ['0.00;(0.00);"zero"', 0, 'zero'],
    ['0.00;(0.00);"zero"', -1, '(1.00)'],
    ['[Red]#,##0.00;[Blue](#,##0.00)', -99.5, '(99.50)'],
    // An empty section hides the value.
    ['#,##0;', -5, ''],
    ['', 42, ''],
    // A lone section takes the sign in front of the whole rendering.
    ['0 ;(0)', -3, '(3)'],
  ] satisfies Array<[string, CellValue, string]>)('%s + %j => %j', (code, value, expected) => {
    expect(display(code, value)).toBe(expected);
  });

  it('a value that rounds to zero prints without a sign', () => {
    expect(display('0.00', -0.0004)).toBe('0.00');
  });

  it('accounting codes render all four of their sections', () => {
    const accounting = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';
    expect(display(accounting, 1_234.5)).toBe(' $1,234.50 ');
    expect(display(accounting, -1_234.5)).toBe(' $(1,234.50)');
    expect(display(accounting, 0)).toBe(' $-   ');
    expect(display(accounting, 'n/a')).toBe(' n/a ');
  });
});

describe('getCellDisplayText: text sections', () => {
  it('the fourth section formats a string value', () => {
    expect(display('0.00;(0.00);"zero";"[" @ "]"', 'hi')).toBe('[ hi ]');
  });

  it('a lone text layout formats a string value', () => {
    expect(display('@" pcs"', '12')).toBe('12 pcs');
    expect(display('@', 'text')).toBe('text');
  });

  it('a numeric code leaves a string alone', () => {
    expect(display('#,##0.00', 'not a number')).toBe('not a number');
    expect(display('0.00;(0.00)', 'still text')).toBe('still text');
  });

  it('a number in a text-formatted cell falls back to General', () => {
    expect(display('@', 123.5)).toBe('123.5');
  });
});

describe('getCellDisplayText: dates and times', () => {
  it.each([
    ['mm-dd-yy', MARCH_14_2024, '03-14-24'],
    ['d-mmm-yy', MARCH_14_2024, '14-Mar-24'],
    ['d-mmm', MARCH_14_2024, '14-Mar'],
    ['mmm-yy', MARCH_14_2024, 'Mar-24'],
    ['mmmm d, yyyy', MARCH_14_2024, 'March 14, 2024'],
    ['mmmmm', MARCH_14_2024, 'M'],
    ['ddd', MARCH_14_2024, 'Thu'],
    ['dddd', MARCH_14_2024, 'Thursday'],
    ['yyyy-mm-dd hh:mm:ss', 45_365.7375, '2024-03-14 17:42:00'],
    ['m/d/yy h:mm', 45_365.5, '3/14/24 12:00'],
    ['h:mm AM/PM', 45_365.5, '12:00 PM'],
    ['h:mm:ss AM/PM', 45_365.02, '12:28:48 AM'],
    ['h:mm:ss A/P', 45_365.6, '2:24:00 P'],
    // Excel prints the meridiem in the case the code spells it.
    ['h:mm am/pm', 45_365.5, '12:00 pm'],
    ['h:mm a/p', 45_365.5, '12:00 p'],
    // `m` is a minute next to an hour or a second, a month everywhere else.
    ['mm:ss', 0.5 + 61 / 86_400, '01:01'],
    ['h"h" mm"m"', 45_365.5, '12h 00m'],
    ['mmss.0', 0.5 + 61.4 / 86_400, '0101.4'],
  ])('%s + %j => %j', (code, value, expected) => {
    expect(display(code, value)).toBe(expected);
  });

  it.each([
    ['[h]:mm:ss', 1.5, '36:00:00'],
    ['[mm]:ss', 0.5, '720:00'],
    ['[h]:mm', 0.0625, '1:30'],
  ])('elapsed %s + %j => %j', (code, value, expected) => {
    expect(display(code, value)).toBe(expected);
  });

  it('a duration value renders through an elapsed code', () => {
    expect(display('[h]:mm', makeDurationValue(5_400_000))).toBe('1:30');
  });

  it('an elapsed span keeps its sign, unlike a calendar date', () => {
    expect(display('[h]:mm', -0.0625)).toBe('-1:30');
    expect(display('[h]:mm', makeDurationValue(-5_400_000))).toBe('-1:30');
  });

  it('renders each fractional-second group at its own width', () => {
    // The serial is rounded once, at the widest group; the narrower group
    // rounds those digits rather than reading the serial again.
    expect(display('[ss].000" "ss.0', 43_201.234 / 86_400)).toBe('43201.234 01.2');
  });

  it('a Date value renders through a date code', () => {
    expect(display('yyyy-mm-dd', new Date(Date.UTC(2024, 2, 14)))).toBe('2024-03-14');
  });

  it('a Date value under a non-date code shows the serial, the way the file does', () => {
    // The worksheet stores a serial and the style decides what it means, so a
    // Date in a General cell reads as a number in Excel too.
    expect(display('General', new Date(Date.UTC(2024, 2, 14)))).toBe('45365');
  });

  it('reads serials under the workbook epoch', () => {
    expect(display('yyyy-mm-dd', MARCH_14_2024, { date1904: false })).toBe('2024-03-14');
    expect(display('yyyy-mm-dd', MARCH_14_2024, { date1904: true })).toBe('2028-03-15');
  });

  it('a negative serial has no date reading', () => {
    // Excel fills the cell with `#` characters, and how many depends on the
    // column width, so there is nothing faithful to print.
    expect(display('yyyy-mm-dd', -5)).toBe('-5');
  });
});

describe('getCellDisplayText: value kinds a number format does not touch', () => {
  it.each([
    [true, 'TRUE'],
    [false, 'FALSE'],
  ])('a boolean prints as %j => %j whatever the code says', (value, expected) => {
    expect(display('0.00', value)).toBe(expected);
    expect(display('yyyy-mm-dd', value)).toBe(expected);
  });

  it('an error cell prints its token', () => {
    expect(display('0.00', makeErrorValue('#REF!'))).toBe('#REF!');
  });

  it('rich text prints its runs joined', () => {
    const runs = makeRichText([makeTextRun('bold '), makeTextRun('and plain')]);
    expect(display('0.00', { kind: 'rich-text', runs })).toBe('bold and plain');
    expect(display('@" (note)"', { kind: 'rich-text', runs })).toBe('bold and plain (note)');
  });

  it('an empty cell prints nothing', () => {
    expect(display('0.00', null)).toBe('');
  });
});

describe('getCellDisplayText: formula cells', () => {
  it('renders the cached value through the format', () => {
    expect(display('0.0%', makeFormula('A1/B1', { cachedValue: 0.5 }))).toBe('50.0%');
    expect(display('yyyy-mm-dd', makeFormula('TODAY()', { cachedValue: MARCH_14_2024 }))).toBe('2024-03-14');
  });

  it('prints nothing when the producer cached no value', () => {
    expect(display('0.00', makeFormula('A1+B1'))).toBe('');
  });

  it('prints an error result as its token', () => {
    expect(display('0.00', makeFormula('A1/0', { cachedValue: '#DIV/0!', cachedValueType: 'error' }))).toBe('#DIV/0!');
  });
});

// Each row was read back from Excel for Mac with the code applied to the value.
describe('getCellDisplayText: comparison sections', () => {
  it.each([
    ['[<1000]0;[<1000000]0.0,"K";0.0,,"M"', 12345, '12.3K'],
    ['[<1000]0;[<1000000]0.0,"K";0.0,,"M"', 1500000, '1.5M'],
    ['[<1000]0;[<1000000]0.0,"K"', 500, '500'],
    ['[<1000]0;[<1000000]0.0,"K"', -5, '-5'],
    ['[>10]"big";[<0]"neg";"small"', 5, 'small'],
    ['[=1]"one";"other"', 1, 'one'],
    ['[>=0]0;[<0]0', -5, '5'],
    ['[<0]0;0', -5, '5'],
    ['[>100]0;0', -5, '-5'],
    ['[>100]0;[<-100]0;0', -5, '-5'],
    ['[>100]0;[<-100]0;0', -500, '500'],
    ['[>10]0;0;"z"', 5, 'z'],
    ['[=1]0;[=2]0;0', -3, '-3'],
    ['[Red][<0]0.0;0', -2, '2.0'],
  ] satisfies Array<[string, CellValue, string]>)('%s on %s shows %s', (code, value, expected) => {
    expect(display(code, value)).toBe(expected);
  });

  it('a value no section takes has no reading and degrades', () => {
    expect(display('[<1000]0;[<1000000]0.0,"K"', 2000000)).toBe('2000000');
  });
});

describe('getCellDisplayText: the 1900 system\'s phantom days', () => {
  it('shows serial 0 as 1/0/1900 and serial 60 as 2/29/1900, like Excel', () => {
    expect(display('m/d/yyyy', 0)).toBe('1/0/1900');
    expect(display('m/d/yyyy', 60)).toBe('2/29/1900');
    expect(display('m/d/yyyy', 59)).toBe('2/28/1900');
    expect(display('m/d/yyyy', 61)).toBe('3/1/1900');
    expect(display('dddd', 60)).toBe('Wednesday');
  });
});

describe('getCellDisplayText: codes outside the supported set', () => {
  it.each([
    // Era and calendar tokens need a calendar this renderer does not carry.
    ['ggge"年"m"月"d"日"', MARCH_14_2024, '45365'],
    ['[DBNum1]0', 42, '42'],
    // More than the four sections Excel defines.
    ['0;0;0;0;0', 1.5, '1.5'],
    // Unterminated quote and bracket.
    ['0.00"unterminated', 1.5, '1.5'],
    ['[Red0.00', 1.5, '1.5'],
    // Two numeric layouts spliced into one section: there is no single number
    // to lay out, so neither layout gets applied.
    ['0.00" ("0.00")"', 1.5, '1.5'],
  ] satisfies Array<[string, CellValue, string]>)('%s degrades to the plain coercion', (code, value, expected) => {
    expect(display(code, value)).toBe(expected);
  });

  it('degrades a Date to its ISO form, not to a serial', () => {
    const date = new Date(Date.UTC(2024, 2, 14));
    expect(display('[DBNum1]0', date)).toBe(date.toISOString());
  });

  it('a serial past the range a Date covers degrades to the number', () => {
    expect(display('yyyy-mm-dd', 1e9)).toBe('1000000000');
  });

  it('an invalid Date and a non-finite span degrade rather than throw', () => {
    // Mapping a whole sheet to text cannot lose the sheet over one bad value.
    expect(display('yyyy-mm-dd', new Date(Number.NaN))).toBe('Invalid Date');
    expect(display('[h]:mm', { kind: 'duration', ms: Number.NaN })).toBe('NaN ms');
  });
});

describe('getCellDisplayText: the styleId drives the format', () => {
  it('a cell with the default style reads as General', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    expect(getCellDisplayText(wb, setCell(ws, 1, 1, 1.1 + 2.2))).toBe('3.3');
  });

  it('two cells in one workbook keep their own formats', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    const percent = setCell(ws, 1, 1, 0.5);
    const date = setCell(ws, 1, 2, MARCH_14_2024);
    setCellNumberFormat(wb, percent, '0.0%');
    setCellNumberFormat(wb, date, 'yyyy-mm-dd');
    expect(getCellDisplayText(wb, percent)).toBe('50.0%');
    expect(getCellDisplayText(wb, date)).toBe('2024-03-14');
  });
});

it('renders literal-only numeric and text sections without appending the value', () => {
  expect(display('"red"', 12)).toBe('red');
  expect(display('0;0;0;"hidden"', 'hello')).toBe('hidden');
});
it('reads fixed fraction denominators containing zero', () => {
  expect(display('# ?/10', 0.3)).toBe(' 3/10');
});
it('falls back when percentage scaling overflows', () => {
  expect(display('0%', 1e308)).toBe('1e+308');
});
it('keeps a fixed denominator separate from a following literal', () => {
  expect(display('# ?/10" kg"', 0.3)).toBe(' 3/10 kg');
});
