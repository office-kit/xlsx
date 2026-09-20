// Turns a loaded worksheet into plain data that SheetGrid.svelte can draw.
// It runs at prerender time for the landing page and in the browser for the
// playground and the REPL, so it depends on nothing but the library.

import { isFormulaValue } from '@office-kit/xlsx/cell';
import {
  getCellAlignment,
  getCellBorder,
  getCellDisplayText,
  getCellFill,
  getCellFont,
  type Color,
  type Side,
} from '@office-kit/xlsx/styles';
import { columnLetterFromIndex, coordinateToTuple } from '@office-kit/xlsx/utils';
import { iterWorksheets, type Workbook } from '@office-kit/xlsx/workbook';
import {
  getCell,
  getCellExtent,
  getCellExtentRef,
  getColumnDimension,
  getFreezePanes,
  getMergedCells,
  type Worksheet,
} from '@office-kit/xlsx/worksheet';

export type GridEdge = { width: number; color?: string };

export type GridCell = {
  text: string;
  /** Present on a formula cell, with its leading `=`. */
  formula?: string;
  align: 'left' | 'center' | 'right';
  bold: boolean;
  italic: boolean;
  /** CSS colours, present only when the cell's style sets one. */
  color?: string;
  fill?: string;
  edges: { top?: GridEdge; right?: GridEdge; bottom?: GridEdge; left?: GridEdge };
  /** Greater than 1 on the top-left cell of a merged range. */
  rowSpan: number;
  colSpan: number;
  /** Covered by a merged range whose top-left cell is elsewhere; not drawn. */
  covered: boolean;
};

export type SheetGridData = {
  name: string;
  /** `width` is in CSS pixels. */
  columns: Array<{ letter: string; width: number }>;
  rows: Array<{ number: number; cells: GridCell[] }>;
  /** Rows and columns held in place by freeze panes, counted from the top left. */
  frozenRows: number;
  frozenCols: number;
  /** The sheet's used range, such as `A1:F18`, or `''` when it has no cells. */
  usedRange: string;
  /** The used range is larger than the `maxRows` x `maxCols` that were read. */
  clipped: boolean;
};

export type SheetGridOptions = {
  /** A DOM table of a whole sheet would freeze the tab, so reading stops here. */
  maxRows: number;
  maxCols: number;
  /** Pads a small sheet with empty cells so it reads as a spreadsheet, not a table. */
  minRows?: number;
  minCols?: number;
};

// Excel's column width when a sheet sets none, in character units.
const DEFAULT_COLUMN_WIDTH = 8.43;

// Excel turns a width in characters into pixels as `chars * digit + padding`,
// where a digit of the default font (Calibri 11) is 7px wide.
const DIGIT_WIDTH_PX = 7;
const CELL_PADDING_PX = 5;

const EDGE_WIDTH_PX: Partial<Record<NonNullable<Side['style']>, number>> = {
  medium: 2,
  mediumDashed: 2,
  mediumDashDot: 2,
  mediumDashDotDot: 2,
  double: 3,
  thick: 3,
};
const THIN_EDGE_PX = 1;

// The model stores colours as "AARRGGBB"; CSS wants "#RRGGBB". Theme and
// indexed colours need the workbook's palette to resolve, so they fall back to
// the grid's default ink rather than to a guess.
const css = (color: Color | undefined): string | undefined =>
  color?.rgb === undefined ? undefined : `#${color.rgb.slice(2)}`;

const edge = (side: Side | undefined): GridEdge | undefined => {
  if (side?.style === undefined || side.style === 'none') return undefined;
  const color = css(side.color);
  return { width: EDGE_WIDTH_PX[side.style] ?? THIN_EDGE_PX, ...(color ? { color } : {}) };
};

const EMPTY_CELL: GridCell = {
  text: '',
  align: 'left',
  bold: false,
  italic: false,
  edges: {},
  rowSpan: 1,
  colSpan: 1,
  covered: false,
};

function readCell(wb: Workbook, ws: Worksheet, row: number, col: number): GridCell {
  const cell = getCell(ws, row, col);
  if (!cell) return EMPTY_CELL;
  const formula = isFormulaValue(cell.value) ? cell.value : undefined;
  const value = formula ? formula.cachedValue : cell.value;
  const font = getCellFont(wb, cell);
  const fill = getCellFill(wb, cell);
  const border = getCellBorder(wb, cell);
  const horizontal = getCellAlignment(wb, cell).horizontal;
  const color = css(font.color);
  const background = fill.kind === 'pattern' && fill.patternType === 'solid' ? css(fill.fgColor) : undefined;
  const edges: GridCell['edges'] = {};
  for (const name of ['top', 'right', 'bottom', 'left'] as const) {
    const drawn = edge(border[name]);
    if (drawn) edges[name] = drawn;
  }
  return {
    text: getCellDisplayText(wb, cell),
    ...(formula ? { formula: `=${formula.formula}` } : {}),
    // Excel right-aligns numbers and dates unless the cell says otherwise.
    align:
      horizontal === 'center' || horizontal === 'right' || horizontal === 'left'
        ? horizontal
        : typeof value === 'number' || value instanceof Date
          ? 'right'
          : 'left',
    bold: font.bold === true,
    italic: font.italic === true,
    ...(color ? { color } : {}),
    ...(background ? { fill: background } : {}),
    edges,
    rowSpan: 1,
    colSpan: 1,
    covered: false,
  };
}

export function readSheetGrid(wb: Workbook, ws: Worksheet, options: SheetGridOptions): SheetGridData {
  const extent = getCellExtent(ws);
  const usedRows = extent?.maxRow ?? 0;
  const usedCols = extent?.maxCol ?? 0;
  const rowCount = Math.max(Math.min(usedRows, options.maxRows), options.minRows ?? 0);
  const colCount = Math.max(Math.min(usedCols, options.maxCols), options.minCols ?? 0);

  const rows = Array.from({ length: rowCount }, (_, r) => ({
    number: r + 1,
    cells: Array.from({ length: colCount }, (__, c) => readCell(wb, ws, r + 1, c + 1)),
  }));

  for (const range of getMergedCells(ws)) {
    const anchor = rows[range.minRow - 1]?.cells[range.minCol - 1];
    if (!anchor) continue;
    const lastRow = Math.min(range.maxRow, rowCount);
    const lastCol = Math.min(range.maxCol, colCount);
    // Cells are shared with EMPTY_CELL until written to, so replace, never mutate.
    const replace = (row: number, col: number, next: GridCell): void => {
      const target = rows[row - 1];
      if (target) target.cells[col - 1] = next;
    };
    for (let row = range.minRow; row <= lastRow; row++) {
      for (let col = range.minCol; col <= lastCol; col++) {
        replace(row, col, { ...EMPTY_CELL, covered: true });
      }
    }
    replace(range.minRow, range.minCol, {
      ...anchor,
      rowSpan: lastRow - range.minRow + 1,
      colSpan: lastCol - range.minCol + 1,
    });
  }

  const frozen = getFreezePanes(ws);
  const topLeft = frozen === undefined ? undefined : coordinateToTuple(frozen);

  return {
    name: ws.title,
    columns: Array.from({ length: colCount }, (_, c) => ({
      letter: columnLetterFromIndex(c + 1),
      width: Math.round(
        (getColumnDimension(ws, c + 1)?.width ?? DEFAULT_COLUMN_WIDTH) * DIGIT_WIDTH_PX + CELL_PADDING_PX,
      ),
    })),
    rows,
    frozenRows: topLeft ? topLeft.row - 1 : 0,
    frozenCols: topLeft ? topLeft.col - 1 : 0,
    usedRange: getCellExtentRef(ws) ?? '',
    clipped: usedRows > rowCount || usedCols > colCount,
  };
}

export function readWorkbookGrids(wb: Workbook, options: SheetGridOptions): SheetGridData[] {
  return [...iterWorksheets(wb)].map((ws) => readSheetGrid(wb, ws, options));
}
