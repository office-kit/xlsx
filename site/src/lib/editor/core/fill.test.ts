import type { CellValue } from '@office-kit/xlsx/cell';
import { expect, test } from 'vitest';
import { extendSeries, type FillContext } from './fill.ts';

const TIME_STYLE = 1;
const ctx: FillContext = {
  translate: (f) => f,
  isDate: (styleId) => styleId === TIME_STYLE,
  isTime: (styleId) => styleId === TIME_STYLE,
};
const fill = (values: CellValue[], count: number, styleId = 0) =>
  extendSeries(values.map((value) => ({ value, styleId })), count, 1, 'row', ctx).map((s) => s.value);

// Rows from Microsoft's "Fill data automatically in worksheet cells" table.
test('AutoFill continues the series Excel documents', () => {
  expect(fill(['Qtr3'], 3)).toEqual(['Qtr4', 'Qtr1', 'Qtr2']);
  expect(fill(['Quarter3'], 2)).toEqual(['Quarter4', 'Quarter1']);
  expect(fill(['1st Period'], 2)).toEqual(['2nd Period', '3rd Period']);
  expect(fill(['text1', 'textA'], 3)).toEqual(['text2', 'textA', 'text3']);
  expect(fill(['Product 1'], 2)).toEqual(['Product 2', 'Product 3']);
});

test('a lone time counts up by an hour', () => {
  const nine = 9 / 24;
  const out = fill([nine], 3, TIME_STYLE) as number[];
  expect(out.map((v) => Math.round(v * 24))).toEqual([10, 11, 12]);
});
