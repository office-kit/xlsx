// A script that loads a file, edits it and saves it must give the same bytes
// when run again on the same input, and a saved file reloaded and re-saved
// must not drift. Both need the zip mtime and the document timestamps pinned
// (see deterministic-bytes.test.ts); everything else has to be stable on its own.

import { describe, expect, it } from 'vitest';
import { makeFormula } from '../../src/cell/cell.js';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { makePerson } from '../../src/workbook/persons.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { makeThreadedComment } from '../../src/worksheet/threaded-comments.js';
import { setCell } from '../../src/worksheet/worksheet.js';

const STAMP = new Date(Date.UTC(2026, 0, 2, 3, 4, 0));

async function source(): Promise<Uint8Array> {
  const wb = createWorkbook();
  wb.properties = { creator: 'test', created: STAMP.toISOString(), modified: STAMP.toISOString() };
  const ws = addWorksheet(wb, 'Data');
  setCell(ws, 1, 1, 'a');
  setCell(ws, 1, 2, 1);
  setCell(ws, 2, 1, makeFormula('_xlfn.XLOOKUP("a",A1:A1,B1:B1)'));
  const person = makePerson({ displayName: 'Reviewer' });
  (wb.persons ??= []).push(person);
  (ws.threadedComments ??= []).push(makeThreadedComment({ ref: 'B1', personId: person.id, text: 'check', created: STAMP }));
  return workbookToBytes(wb, { mtime: STAMP });
}

/** The "user script": load, change one cell, save. */
async function editScript(input: Uint8Array): Promise<Uint8Array> {
  const wb = await loadWorkbook(fromBuffer(input));
  const ref = wb.sheets[0];
  if (ref?.kind !== 'worksheet') return expect.unreachable('expected a worksheet');
  setCell(ref.sheet, 3, 1, makeFormula('SUM(B1:B2)'));
  return workbookToBytes(wb, { mtime: STAMP });
}

describe('load → edit → save', () => {
  it('running the same edit script twice on one input gives identical bytes', async () => {
    const input = await source();
    expect(await editScript(input)).toEqual(await editScript(input));
  });

  it('reloading and re-saving an edited file does not change it', async () => {
    const once = await editScript(await source());
    const wb = await loadWorkbook(fromBuffer(once));
    expect(await workbookToBytes(wb, { mtime: STAMP })).toEqual(once);
  });
});
