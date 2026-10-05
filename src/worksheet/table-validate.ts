// Cross-checks between a TableDefinition and the worksheet cells under it.
//
// Excel treats a table that disagrees with its sheet as a damaged file and
// repairs it by dropping the table, so the mistake reaches the user as missing
// filters and a dead structured reference in the delivered file, a long way
// from the call that built it. Every path that attaches a table to a worksheet
// runs these checks so the throw lands at that call instead.
//
// The load and save paths deliberately do not run them: loadWorkbook accepts a
// mismatched table from an input file, and re-checking on save would make
// saveWorkbook throw on a workbook it had just read.

import { assertExcelName } from '../utils/excel-name.js';
import { type CellValue, isErrorValue, isFormulaValue, isRichTextValue } from '../cell/cell.js';
import { richTextToString } from '../cell/rich-text.js';
import { rangeBoundaries, tupleToCoordinate } from '../utils/coordinate.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import type { TableDefinition } from './table.js';
import type { Worksheet } from './worksheet.js';

// A table ref carries two plain corners. `rangeBoundaries` also accepts the
// whole-column ("A:C") and whole-row ("1:5") forms, which fill the missing axis
// to the sheet maximum and would reach the width check as a 16384-column table,
// and `$` markers, which Excel never writes into `table/@ref`.
const TWO_CORNER_REF_RE = /^[A-Za-z]{1,3}[1-9][0-9]*(:[A-Za-z]{1,3}[1-9][0-9]*)?$/;

/** What a header cell offers up for the comparison against its column name. */
type HeaderCell =
  | { kind: 'text'; text: string }
  | { kind: 'empty' }
  /** Holds something Excel would not keep as header text; `held` describes it. */
  | { kind: 'nonText'; held: string };

/**
 * Excel stores table header text as a string, so only a string, rich text or a
 * formula caching a string can be paired with a column name. Reading a number
 * or a date as its text form would mean guessing at the number format the
 * header carries, and a wrong guess either blesses a file Excel repairs or
 * rejects one it accepts.
 */
const readHeaderCell = (value: CellValue): HeaderCell => {
  if (value === null) return { kind: 'empty' };
  if (typeof value === 'string') return value === '' ? { kind: 'empty' } : { kind: 'text', text: value };
  if (isRichTextValue(value)) {
    const text = richTextToString(value.runs);
    return text === '' ? { kind: 'empty' } : { kind: 'text', text };
  }
  if (isFormulaValue(value)) {
    const cached = value.cachedValue;
    if (cached === undefined) return { kind: 'nonText', held: 'holds a formula with no cached value' };
    if (typeof cached !== 'string') return { kind: 'nonText', held: 'holds a formula whose cached value is not text' };
    return cached === '' ? { kind: 'empty' } : { kind: 'text', text: cached };
  }
  if (typeof value === 'number') return { kind: 'nonText', held: `holds the number ${value}` };
  if (typeof value === 'boolean') return { kind: 'nonText', held: `holds the boolean ${value}` };
  if (value instanceof Date) return { kind: 'nonText', held: 'holds a date' };
  if (isErrorValue(value)) return { kind: 'nonText', held: `holds the error ${value.code}` };
  // Duration is the only CellValue variant left.
  return { kind: 'nonText', held: 'holds a duration' };
};

const TEXT_REMEDY = 'Write the header row before adding the table, or pass headerRowCount: 0 for a header-less table.';
const NON_TEXT_REMEDY = 'Excel stores header text as a string, so the header cell has to hold text.';

/**
 * Reject a table whose declared geometry or column names disagree with the
 * sheet underneath it.
 */
export const validateTableAgainstSheet = (ws: Worksheet, table: TableDefinition): void => {
  assertExcelName(table.displayName, 'table displayName');
  const where = `table "${table.displayName}"`;
  const ref = table.ref.trim();
  if (!TWO_CORNER_REF_RE.test(ref)) {
    throw new OpenXmlSchemaError(
      `${where}: ref "${table.ref}" is not a plain two-corner cell range like "A1:C4"`,
    );
  }
  const bounds = rangeBoundaries(ref);
  const { columns } = table;
  const width = bounds.maxCol - bounds.minCol + 1;
  if (columns.length !== width) {
    throw new OpenXmlSchemaError(
      `${where}: ref "${table.ref}" spans ${width} column(s) but ${columns.length} column(s) were supplied`,
    );
  }
  const headerRows = table.headerRowCount ?? 1;
  const totalsRows = table.totalsRowCount ?? 0;
  for (const field of ['headerRowCount', 'totalsRowCount'] as const) {
    const count = table[field];
    if (count !== undefined && (!Number.isInteger(count) || count < 0 || count > 0xffff_ffff)) {
      throw new OpenXmlSchemaError(`${where}: ${field} must be an unsigned 32-bit integer, got ${count}`);
    }
  }
  const height = bounds.maxRow - bounds.minRow + 1;
  if (height < headerRows + totalsRows) {
    throw new OpenXmlSchemaError(
      `${where}: ref "${table.ref}" is ${height} row(s) tall, which cannot contain` +
        ` ${headerRows} header row(s) and ${totalsRows} totals row(s)`,
    );
  }
  // Structured references are case-insensitive, so two names differing only in
  // case are ambiguous and Excel drops the table rather than pick one.
  const seen = new Set<string>();
  for (const [i, column] of columns.entries()) {
    if (column.name === '') {
      throw new OpenXmlSchemaError(`${where}: column ${i + 1} has no name; every table column needs one`);
    }
    const key = column.name.toLowerCase();
    if (seen.has(key)) {
      throw new OpenXmlSchemaError(
        `${where}: column ${i + 1} repeats the name "${column.name}"; table column names have to be unique`,
      );
    }
    seen.add(key);
  }
  if (headerRows === 0) return;
  const headerRow = ws.rows.get(bounds.minRow);
  for (const [i, column] of columns.entries()) {
    const col = bounds.minCol + i;
    const header = readHeaderCell(headerRow?.get(col)?.value ?? null);
    if (header.kind === 'text' && header.text === column.name) continue;
    const held = header.kind === 'text' ? `holds "${header.text}"` : header.kind === 'empty' ? 'is empty' : header.held;
    throw new OpenXmlSchemaError(
      `${where}: header cell ${tupleToCoordinate(col, bounds.minRow)} ${held}` +
        ` but column ${i + 1} is named "${column.name}".` +
        ` ${header.kind === 'nonText' ? NON_TEXT_REMEDY : TEXT_REMEDY}`,
    );
  }
};
