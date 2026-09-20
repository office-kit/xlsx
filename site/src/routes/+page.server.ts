import { isFormulaValue } from '@office-kit/xlsx/cell';
import { fromArrayBuffer, loadWorkbook, workbookToBytes } from '@office-kit/xlsx/io';
import {
  getCellBorder,
  getCellDisplayText,
  getCellFill,
  getCellFont,
  type Color,
} from '@office-kit/xlsx/styles';
import { columnLetterFromIndex } from '@office-kit/xlsx/utils';
import { getSheetByIndex } from '@office-kit/xlsx/workbook';
import { getCell, getCellExtent, getColumnDimension } from '@office-kit/xlsx/worksheet';
import basicReadWriteSource from '$lib/examples/basic-read-write.ts?raw';
import { buildHeroWorkbook } from '$lib/examples/hero-workbook';
import heroWorkbookSource from '$lib/examples/hero-workbook.ts?raw';
import streamingWriteSource from '$lib/examples/streaming-write.ts?raw';
import { highlight } from '$lib/server/highlight';
import type { PageServerLoad } from './$types';

export type GridCell = {
  text: string;
  /** Present on a formula cell; the landing page shows it in the formula bar. */
  formula?: string;
  numeric: boolean;
  bold: boolean;
  /** CSS colours, present only when the cell's style sets one. */
  color?: string;
  fill?: string;
  ruleAbove: boolean;
};

export type Grid = {
  sheetName: string;
  /** `width` is in CSS pixels. */
  columns: Array<{ letter: string; width: number }>;
  rows: Array<{ number: number; cells: GridCell[] }>;
};

// Each example opens with a comment addressed to maintainers; visitors only
// need the code, so a snippet starts where the code does.
const from = (source: string, marker: string): string => source.slice(source.indexOf(marker));

// Excel's column width when a sheet sets none, in character units.
const DEFAULT_COLUMN_WIDTH = 8.43;

// Excel turns a width in characters into pixels as `chars * digit + padding`,
// where a digit of the default font (Calibri 11) is 7px wide.
const DIGIT_WIDTH_PX = 7;
const CELL_PADDING_PX = 5;

// The sheet is drawn as a window onto the grid, a little larger than the data,
// so the empty cells around the table read as a spreadsheet and not a table.
const VIEWPORT_ROWS = 11;
const VIEWPORT_COLS = 6;

// The model stores colours as "AARRGGBB"; CSS wants "#RRGGBB".
const css = (color: Color | undefined): string | undefined =>
  color?.rgb === undefined ? undefined : `#${color.rgb.slice(2)}`;

export const load: PageServerLoad = async () => {
  // The grid is read from the saved bytes, not from the object the example
  // returned, so the picture on the page has been through a full write and
  // read. Doing it at prerender time puts it in the static HTML: it needs no
  // client JS and cannot drift from the code beside it.
  const bytes = await workbookToBytes(buildHeroWorkbook());
  const wb = await loadWorkbook(fromArrayBuffer(bytes));
  const ws = getSheetByIndex(wb, 0);
  const extent = ws && getCellExtent(ws);
  if (!ws || !extent) throw new Error('hero-workbook.ts produced no cells');

  const columnIndices = Array.from(
    { length: Math.max(extent.maxCol, VIEWPORT_COLS) },
    (_, i) => i + 1,
  );
  const grid: Grid = {
    sheetName: ws.title,
    columns: columnIndices.map((col) => ({
      letter: columnLetterFromIndex(col),
      width: Math.round(
        (getColumnDimension(ws, col)?.width ?? DEFAULT_COLUMN_WIDTH) * DIGIT_WIDTH_PX +
          CELL_PADDING_PX,
      ),
    })),
    rows: Array.from({ length: Math.max(extent.maxRow, VIEWPORT_ROWS) }, (_, i) => {
      const row = i + 1;
      return {
        number: row,
        cells: columnIndices.map((col): GridCell => {
          const cell = getCell(ws, row, col);
          if (!cell) return { text: '', numeric: false, bold: false, ruleAbove: false };
          const formula = isFormulaValue(cell.value) ? cell.value : undefined;
          const font = getCellFont(wb, cell);
          const fill = getCellFill(wb, cell);
          return {
            text: getCellDisplayText(wb, cell),
            formula: formula && `=${formula.formula}`,
            numeric: typeof (formula ? formula.cachedValue : cell.value) === 'number',
            bold: font.bold === true,
            color: css(font.color),
            fill: fill.kind === 'pattern' && fill.patternType === 'solid' ? css(fill.fgColor) : undefined,
            ruleAbove: getCellBorder(wb, cell).top?.style !== undefined,
          };
        }),
      };
    }),
  };

  const [heroCode, roundTripCode, streamingCode] = await Promise.all([
    highlight(from(heroWorkbookSource, 'export function'), 'ts'),
    highlight(from(basicReadWriteSource, 'import '), 'ts'),
    highlight(from(streamingWriteSource, 'import '), 'ts'),
  ]);

  return { grid, heroCode, roundTripCode, streamingCode };
};
