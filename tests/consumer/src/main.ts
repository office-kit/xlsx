// Runtime half of the packaging gate, and the reason this fixture emits with
// plain `tsc` instead of a bundler: an unbundled `node dist/main.js` is where
// Node's own ESM resolver gets a vote. TypeScript accepting an import proves
// nothing about Node accepting it, so the two have to agree here or a
// consumer ships and then fails with ERR_MODULE_NOT_FOUND.
//
// Static imports touch all fifteen published subpaths so a typo in the
// `exports` map fails here rather than in a user's first hour.

import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { cellValueAsNumber, makeFormula, makeRichText } from '@office-kit/xlsx/cell';
import { makeBarChart } from '@office-kit/xlsx/chart';
import { makeChartsheet } from '@office-kit/xlsx/chartsheet';
import { makeOneCellAnchor } from '@office-kit/xlsx/drawing';
import { loadWorkbook, saveWorkbook, workbookToBytes } from '@office-kit/xlsx/io';
import { fromBuffer, fromFile, toFile } from '@office-kit/xlsx/node';
import { makeCoreProperties } from '@office-kit/xlsx/packaging';
import { defineSchema } from '@office-kit/xlsx/schema';
import { createWriteOnlyWorkbook, loadWorkbookStream } from '@office-kit/xlsx/streaming';
import { makeAlignment } from '@office-kit/xlsx/styles';
import { inferCellType, OpenXmlError } from '@office-kit/xlsx/utils';
import { addWorksheet, createWorkbook } from '@office-kit/xlsx/workbook';
import { getCell, setCell } from '@office-kit/xlsx/worksheet';
import { parseXml } from '@office-kit/xlsx/xml';
import { DEFAULT_DECOMPRESSION_LIMITS, openZip } from '@office-kit/xlsx/zip';

import { ok } from './assert.js';

const subpathEntries: ReadonlyArray<readonly [string, unknown]> = [
  ['cell', makeRichText],
  ['chart', makeBarChart],
  ['chartsheet', makeChartsheet],
  ['drawing', makeOneCellAnchor],
  ['io', loadWorkbook],
  ['node', fromFile],
  ['packaging', makeCoreProperties],
  ['schema', defineSchema],
  ['streaming', loadWorkbookStream],
  ['streaming (write-only)', createWriteOnlyWorkbook],
  ['styles', makeAlignment],
  ['utils', inferCellType],
  ['utils (errors)', OpenXmlError],
  ['workbook', createWorkbook],
  ['worksheet', setCell],
  ['xml', parseXml],
  ['zip', openZip],
];

for (const [subpath, exported] of subpathEntries) {
  ok(typeof exported === 'function', `${subpath} export is not callable at runtime`);
}
ok(cellValueAsNumber({ kind: 'rich-text', runs: makeRichText([{ text: '1' }, { text: '2.5', font: { b: true } }]) }) === 12.5, 'published numeric coercion lost rich text');
ok(DEFAULT_DECOMPRESSION_LIMITS.maxTotalUncompressedBytes > 0, 'zip limits did not load');

const scratch = mkdtempSync(join(tmpdir(), 'xlsx-consumer-'));
try {
  const wb = createWorkbook();
  const sheet = addWorksheet(wb, 'Consumer');
  setCell(sheet, 1, 1, 'round-trip');
  setCell(sheet, 2, 1, 42);

  const reloaded = await loadWorkbook(fromBuffer(await workbookToBytes(wb)));
  const first = reloaded.sheets[0];
  ok(first?.kind === 'worksheet', 'in-memory round-trip lost the worksheet');
  ok(getCell(first.sheet, 1, 1)?.value === 'round-trip', 'in-memory round-trip lost A1');

  const path = join(scratch, 'consumer.xlsx');
  await saveWorkbook(wb, toFile(path));
  const fromDisk = await loadWorkbook(fromFile(path));
  const diskSheet = fromDisk.sheets[0];
  ok(diskSheet?.kind === 'worksheet', 'disk round-trip lost the worksheet');
  ok(getCell(diskSheet.sheet, 2, 1)?.value === 42, 'disk round-trip lost A2');

  // An edit script run twice on the same input must write the same bytes, and
  // re-saving its output must not drift. Only the public subpaths are used,
  // with the zip mtime and document dates pinned (the wall clock differs).
  const stamp = new Date(Date.UTC(2026, 0, 2, 3, 4, 0));
  wb.properties = { creator: 'consumer', created: stamp.toISOString(), modified: stamp.toISOString() };
  const input = await workbookToBytes(wb, { mtime: stamp });
  const editScript = async (bytes: Uint8Array): Promise<Uint8Array> => {
    const book = await loadWorkbook(fromBuffer(bytes));
    const ws = book.sheets[0];
    ok(ws?.kind === 'worksheet', 'edit script lost the worksheet');
    setCell(ws.sheet, 3, 1, makeFormula('_xlfn.XLOOKUP(42,A2:A2,A1:A1)'));
    return workbookToBytes(book, { mtime: stamp });
  };
  const runA = await editScript(input);
  const runB = await editScript(input);
  ok(runA.length === runB.length && runA.every((b, i) => b === runB[i]), 'running the edit script twice gave different bytes');
  const resaved = await workbookToBytes(await loadWorkbook(fromBuffer(runA)), { mtime: stamp });
  ok(resaved.length === runA.length && resaved.every((b, i) => b === runA[i]), 're-saving the edited file changed it');

  console.log(`ok: ${subpathEntries.length} subpaths loaded, round-tripped via memory and ${path}`);
} finally {
  rmSync(scratch, { recursive: true, force: true });
}
