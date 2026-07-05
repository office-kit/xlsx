// Apply number formats: currency, percentage, and a date-time.

import { saveWorkbook } from '@office-kit/xlsx/io';
import { toFile } from '@office-kit/xlsx/node';
import {
  FORMAT_DATE_DATETIME,
  setCellAsCurrency,
  setCellAsPercent,
  setCellNumberFormat,
} from '@office-kit/xlsx/styles';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { setCell } from '@office-kit/xlsx/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Numbers');

setCellAsCurrency(wb, setCell(ws, 1, 1, 12_400), { symbol: '$' });
setCellAsPercent(wb, setCell(ws, 1, 2, 0.187), 1);

const dateCell = setCell(ws, 1, 3, new Date('2026-05-08T09:30:00Z'));
setCellNumberFormat(wb, dateCell, FORMAT_DATE_DATETIME);

await saveWorkbook(wb, toFile('numbers.xlsx'));
