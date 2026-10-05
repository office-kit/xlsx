import { makeCfRule, makeConditionalFormatting, makeDataValidation, setCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { chartRefs, insertChart } from './charts.ts';
import { EditorController } from './controller.svelte.ts';

function withChart(): EditorController {
  const ctl = new EditorController();
  const ws = ctl.doc.ws;
  setCell(ws, 1, 1, 'Item');
  setCell(ws, 1, 2, 'Qty');
  for (let r = 2; r <= 5; r++) {
    setCell(ws, r, 1, `i${r}`);
    setCell(ws, r, 2, r);
  }
  insertChart(ctl.doc, 'columnClustered', { r1: 1, c1: 1, r2: 5, c2: 2 });
  return ctl;
}

const refs = (ctl: EditorController) => {
  const item = ctl.doc.ws.drawing?.items[0];
  return item?.content.kind === 'chart' ? chartRefs(item.content.chart) : undefined;
};

it('a chart keeps plotting its cells when rows and columns are inserted', () => {
  const ctl = withChart();
  ctl.selectRange({ r1: 3, c1: 1, r2: 3, c2: 16384 });
  A.insertLines(ctl, 'row');
  expect(refs(ctl)).toEqual({ values: ['Sheet1!$B$2:$B$6'], others: ['Sheet1!$A$2:$A$6', 'Sheet1!$B$1'] });
  ctl.selectRange({ r1: 1, c1: 1, r2: 1048576, c2: 1 });
  A.insertLines(ctl, 'col');
  expect(refs(ctl)).toEqual({ values: ['Sheet1!$C$2:$C$6'], others: ['Sheet1!$B$2:$B$6', 'Sheet1!$C$1'] });
  ctl.doc.undo();
  ctl.doc.undo();
  expect(refs(ctl)).toEqual({ values: ['Sheet1!$B$2:$B$5'], others: ['Sheet1!$A$2:$A$5', 'Sheet1!$B$1'] });
});

it('renaming the sheet re-points charts, conditional formatting and validation, and undo restores them', () => {
  const ctl = withChart();
  const ws = ctl.doc.ws;
  ws.conditionalFormatting.push(makeConditionalFormatting({ sqref: 'C1', rules: [makeCfRule({ type: 'expression', priority: 1, formulas: ['Sheet1!$A$1>3'] })] }));
  ws.dataValidations.push(makeDataValidation({ sqref: 'D1', type: 'list', formula1: 'Sheet1!$A$2:$A$5' }));
  expect(A.renameSheetAt(ctl, 0, 'Data')).toBeUndefined();
  expect(refs(ctl)).toEqual({ values: ['Data!$B$2:$B$5'], others: ['Data!$A$2:$A$5', 'Data!$B$1'] });
  expect(ws.conditionalFormatting[0]?.rules[0]?.formulas).toEqual(['Data!$A$1>3']);
  expect(ws.dataValidations[0]?.formula1).toBe('Data!$A$2:$A$5');
  ctl.doc.undo();
  expect(refs(ctl)?.values).toEqual(['Sheet1!$B$2:$B$5']);
  expect(ctl.doc.ws.dataValidations[0]?.formula1).toBe('Sheet1!$A$2:$A$5');
});
