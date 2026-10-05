// Guards the editor's large-sheet behaviour: a single edit on a big sheet must
// cost time proportional to what it touches, not to the sheet size.

import { makeFormula } from '@office-kit/xlsx/cell';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { getCell, setCell } from '@office-kit/xlsx/worksheet';
import { expect, test } from 'vitest';
import { usedRange } from './cells.ts';
import { EditorController } from './controller.svelte.ts';
import { SpreadsheetEditor } from './editor.svelte.ts';

const ROWS = 100_000;

function bigWorkbook() {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Data');
  for (let r = 1; r <= ROWS; r++) {
    setCell(ws, r, 1, r);
    setCell(ws, r, 2, `item ${r}`);
    setCell(ws, r, 3, r * 1.5);
    setCell(ws, r, 4, makeFormula(`A${r}*C${r}`));
  }
  setCell(ws, 1, 6, makeFormula(`SUM(D1:D${ROWS})`));
  return wb;
}

test('editing one cell of a 400k-cell sheet stays interactive', () => {
  const t0 = performance.now();
  const doc = new SpreadsheetEditor(bigWorkbook());
  const ctl = new EditorController(doc);
  const loaded = performance.now() - t0;

  const ws = doc.ws;
  expect(usedRange(ws)).toEqual({ r1: 1, c1: 1, r2: ROWS, c2: 6 });

  ctl.selectCell({ row: 50_000, col: 1 });
  const t1 = performance.now();
  ctl.startEdit('0');
  ctl.commitEdit();
  const edit = performance.now() - t1;
  const d = getCell(ws, 50_000, 4)?.value;
  expect(d && typeof d === 'object' && 'cachedValue' in d ? d.cachedValue : undefined).toBe(0);

  const t2 = performance.now();
  doc.undo();
  const undo = performance.now() - t2;

  // Generous bounds for CI machines; locally these are a few ms for edit/undo.
  expect(loaded).toBeLessThan(15_000);
  expect(edit).toBeLessThan(500);
  expect(undo).toBeLessThan(500);
});
