import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { setFreezePanes } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import { SpreadsheetEditor } from './editor.svelte.ts';

it('opens a frozen sheet on the active pane selection, as Excel does', () => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'S');
  setFreezePanes(ws, 'B2');
  const view = ws.views[0];
  if (!view) throw new Error('no view');
  view.selections = [
    { pane: 'topRight', activeCell: 'B1', sqref: 'B1' },
    { pane: 'bottomLeft', activeCell: 'A2', sqref: 'A2' },
    { pane: 'bottomRight', activeCell: 'D10', sqref: 'D10' },
  ];
  expect(new SpreadsheetEditor(wb).active).toEqual({ row: 10, col: 4 });
});
