// The workbook shown on the landing page. The page renders this exact code
// next to the sheet it produces, and the "Download the .xlsx" button runs it
// in the visitor's browser — so what is on screen is always what the library
// does.

import { makeFormula } from '@office-kit/xlsx/cell';
import {
  makeBorder,
  makeColor,
  makeFont,
  makePatternFill,
  makeSide,
  registerCellStyle,
} from '@office-kit/xlsx/styles';
import { addWorksheet, createWorkbook, type Workbook } from '@office-kit/xlsx/workbook';
import { appendRow, appendRows, setColumnWidths, setFreezePanes } from '@office-kit/xlsx/worksheet';

export function buildHeroWorkbook(): Workbook {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Q3 sales');

  const white = makeColor({ rgb: 'FFFFFFFF' });
  const green = makeColor({ rgb: 'FF168A4F' });
  const header = registerCellStyle(wb, {
    font: makeFont({ bold: true, color: white }),
    fill: makePatternFill({ patternType: 'solid', fgColor: green }),
  });
  appendRow(ws, ['Region', 'Units', 'Revenue', 'Growth'], {
    styleIds: [header, header, header, header],
  });
  setFreezePanes(ws, 'A2');

  const regions = [
    ['North', 1_240, 186_000, 0.124],
    ['South', 980, 139_650, 0.081],
    ['East', 1_515, 242_400, 0.193],
    ['West', 760, 102_600, -0.036],
  ] as const;
  const units = registerCellStyle(wb, { numberFormat: '#,##0' });
  const money = registerCellStyle(wb, { numberFormat: '"$"#,##0' });
  const percent = registerCellStyle(wb, { numberFormat: '0.0%' });
  appendRows(ws, regions, { styleIds: [undefined, units, money, percent] });

  // The library never evaluates formulas, so cache the sums it can work out:
  // viewers that do not calculate then show a number instead of a blank.
  const sum = (col: 1 | 2): number => regions.reduce((n, region) => n + region[col], 0);
  const total = {
    font: makeFont({ bold: true }),
    border: makeBorder({ top: makeSide({ style: 'thin' }) }),
  };
  appendRow(
    ws,
    [
      'Total',
      makeFormula('SUM(B2:B5)', { cachedValue: sum(1) }),
      makeFormula('SUM(C2:C5)', { cachedValue: sum(2) }),
    ],
    {
      styleIds: [
        registerCellStyle(wb, total),
        registerCellStyle(wb, { ...total, numberFormat: '#,##0' }),
        registerCellStyle(wb, { ...total, numberFormat: '"$"#,##0' }),
        registerCellStyle(wb, total),
      ],
    },
  );

  setColumnWidths(ws, [14, 10, 14, 10]);
  return wb;
}
