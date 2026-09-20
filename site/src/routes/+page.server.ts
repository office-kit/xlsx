import { fromArrayBuffer, loadWorkbook, workbookToBytes } from '@office-kit/xlsx/io';
import basicReadWriteSource from '$lib/examples/basic-read-write.ts?raw';
import { buildHeroWorkbook } from '$lib/examples/hero-workbook';
import heroWorkbookSource from '$lib/examples/hero-workbook.ts?raw';
import streamingWriteSource from '$lib/examples/streaming-write.ts?raw';
import { highlight } from '$lib/server/highlight';
import { readWorkbookGrids } from '$lib/sheet-grid';
import type { PageServerLoad } from './$types';

// Each example opens with a comment addressed to maintainers; visitors only
// need the code, so a snippet starts where the code does.
const from = (source: string, marker: string): string => source.slice(source.indexOf(marker));

// The sheet is drawn as a window onto the grid, a little larger than the data,
// so the empty cells around the table read as a spreadsheet and not a table.
const VIEWPORT = { minRows: 11, minCols: 6, maxRows: 11, maxCols: 6 };

export const load: PageServerLoad = async () => {
  // The grid is read from the saved bytes, not from the object the example
  // returned, so the picture on the page has been through a full write and
  // read. Doing it at prerender time puts it in the static HTML: it needs no
  // client JS and cannot drift from the code beside it.
  const bytes = await workbookToBytes(buildHeroWorkbook());
  const sheets = readWorkbookGrids(await loadWorkbook(fromArrayBuffer(bytes)), VIEWPORT);

  const [heroCode, roundTripCode, streamingCode] = await Promise.all([
    highlight(from(heroWorkbookSource, 'export function'), 'ts'),
    highlight(from(basicReadWriteSource, 'import '), 'ts'),
    highlight(from(streamingWriteSource, 'import '), 'ts'),
  ]);

  return { sheets, heroCode, roundTripCode, streamingCode };
};
