import { expect, it } from 'vitest';
import { parseInput } from './input.ts';

// Each expectation was read back from Excel for Mac after typing the same text.
it.each([
  ['１２３', 123, undefined],
  ['１２．５', 12.5, undefined],
  ['－５', -5, undefined],
  ['＋５', 5, undefined],
  ['（５）', -5, undefined],
  ['　１２　', 12, undefined],
  ['１，０００', 1000, '#,##0'],
  ['５０％', 0.5, '0%'],
  ['１Ｅ３', 1000, '0.00E+00'],
  ['ＴＲＵＥ', true, undefined],
])('%s is read as its half-width value', (typed, value, format) => {
  const parsed = parseInput(typed);
  expect(parsed.value).toBe(value);
  expect(parsed.impliedFormat).toBe(format);
});

it.each(['ＡＢＣ', '１２３ａｂｃ'])('%s stays full-width text', (typed) => {
  expect(parseInput(typed).value).toBe(typed);
});

it('a full-width formula is converted outside its string literals', () => {
  expect(parseInput('＝ｓｕｍ（Ａ１：Ａ２）').value).toMatchObject({ kind: 'formula', formula: 'SUM(A1:A2)' });
  expect(parseInput('＝"ａｂ"＆"ｃ"').value).toMatchObject({ kind: 'formula', formula: '"ａｂ"&"ｃ"' });
});
