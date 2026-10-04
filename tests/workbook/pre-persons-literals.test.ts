// Workbooks and worksheets written as object literals before `persons`,
// `threadedComments` and `sparklineGroups` existed must still type-check and
// must survive every mutation, save and reload with those fields absent.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { makeStylesheet } from '../../src/styles/stylesheet.js';
import { addWorksheet, duplicateSheet, type Workbook } from '../../src/workbook/workbook.js';
import { getCell, makeWorksheet, setCell, type Worksheet } from '../../src/worksheet/worksheet.js';

const STAMP = new Date(Date.UTC(2026, 0, 2, 3, 4, 0));

/** The shape a 0.23 caller could write by hand: `authors`, no `persons`. */
function legacyWorkbook(): Workbook {
  return { sheets: [], activeSheetIndex: 0, styles: makeStylesheet(), date1904: false, authors: [], definedNames: [] };
}

/** A worksheet with the two newer list fields removed, as an older literal would have it. */
function legacySheet(title: string): Worksheet {
  const { threadedComments: _threads, sparklineGroups: _groups, ...rest } = makeWorksheet(title);
  return rest;
}

describe('workbooks without the newer optional lists', () => {
  it('edit, duplicate, save, reload and re-save', async () => {
    const wb = legacyWorkbook();
    const data = legacySheet('Data');
    wb.sheets.push({ kind: 'worksheet', sheet: data, sheetId: 1, state: 'visible' });
    setCell(data, 1, 1, 'kept');
    const added = addWorksheet(wb, 'Added');
    delete added.threadedComments;
    delete added.sparklineGroups;
    setCell(added, 2, 2, 42);
    duplicateSheet(wb, 'Added', 'Copy');

    const bytes = await workbookToBytes(wb, { mtime: STAMP });
    const back = await loadWorkbook(fromBuffer(bytes));
    expect(back.sheets.map((s) => s.sheet.title)).toEqual(['Data', 'Added', 'Copy']);
    const first = back.sheets[0];
    if (first?.kind !== 'worksheet') return expect.unreachable('expected a worksheet');
    expect(getCell(first.sheet, 1, 1)?.value).toBe('kept');
    const copy = back.sheets[2];
    if (copy?.kind !== 'worksheet') return expect.unreachable('expected a worksheet');
    expect(getCell(copy.sheet, 2, 2)?.value).toBe(42);
    expect(back.persons ?? []).toEqual([]);
    expect(await workbookToBytes(back, { mtime: STAMP })).toEqual(bytes);
  });
});
