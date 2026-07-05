// Color-scale rule: red for low values, yellow for the middle, green
// for high. Excel's classic 3-color heat-map.

import { saveWorkbook } from '@office-kit/xlsx/io';
import { toFile } from '@office-kit/xlsx/node';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import {
  addConditionalFormatting,
  makeCfRule,
  makeConditionalFormatting,
  setCell,
} from '@office-kit/xlsx/worksheet';

const wb = createWorkbook();
const ws = addWorksheet(wb, 'Heat');

for (let r = 1; r <= 10; r++) setCell(ws, r, 1, Math.round(Math.random() * 100));

addConditionalFormatting(
  ws,
  makeConditionalFormatting({
    sqref: 'A1:A10',
    rules: [
      makeCfRule({
        type: 'colorScale',
        priority: 1,
        formulas: [],
        innerXml:
          '<colorScale>' +
          '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' +
          '<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/>' +
          '</colorScale>',
      }),
    ],
  }),
);

await saveWorkbook(wb, toFile('heatmap.xlsx'));
