import { addChartsheet, addWorksheet, setSheetState } from '@office-kit/xlsx/workbook';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';

const titles = (ctl: EditorController) => ctl.doc.wb.sheets.map((s) => s.sheet.title);
const active = (ctl: EditorController) => ctl.doc.wb.sheets[ctl.doc.activeSheetIndex]?.sheet.title;

it('a chart sheet can be deleted next to the only worksheet', () => {
  const ctl = new EditorController();
  addChartsheet(ctl.doc.wb, 'Chart1');
  ctl.showSheet(1);
  A.deleteSheet(ctl, 1);
  expect(ctl.dialog).toBeNull();
  expect(titles(ctl)).toEqual(['Sheet1']);
  expect(ctl.shownChartsheet).toBeNull();
});

it('deleting the active sheet moves to the next visible worksheet, skipping hidden ones', () => {
  const ctl = new EditorController();
  addWorksheet(ctl.doc.wb, 'Hidden');
  addWorksheet(ctl.doc.wb, 'Sheet3');
  setSheetState(ctl.doc.wb, 'Hidden', 'hidden');
  A.deleteSheet(ctl, 0);
  expect(active(ctl)).toBe('Sheet3');
});

it('deleting another sheet keeps the active one', () => {
  const ctl = new EditorController();
  addWorksheet(ctl.doc.wb, 'Sheet2');
  addWorksheet(ctl.doc.wb, 'Sheet3');
  ctl.showSheet(2);
  A.deleteSheet(ctl, 0);
  expect(active(ctl)).toBe('Sheet3');
});

it('the last visible worksheet still cannot be deleted', () => {
  const ctl = new EditorController();
  addChartsheet(ctl.doc.wb, 'Chart1');
  A.deleteSheet(ctl, 0);
  expect(titles(ctl)).toEqual(['Sheet1', 'Chart1']);
  expect(ctl.dialog).toMatchObject({ kind: 'alert' });
});
