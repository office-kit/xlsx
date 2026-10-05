// Format-aware readers for a cell: the text Excel prints in it, and the date a
// date-formatted serial stands for. Both need the workbook stylesheet, which is
// why they live here rather than next to the value coercions in `src/cell/`.

import type { Cell, CellValue } from '../cell/cell.js';
import { cellValueAsString, isDurationValue, isErrorValue, isFormulaValue, isRichTextValue } from '../cell/cell.js';
import { richTextToString } from '../cell/rich-text.js';
import { dateToExcel, durationToExcel, excelToDate, type ExcelEpoch } from '../utils/datetime.js';
import type { Workbook } from '../workbook/workbook.js';
import { getCellNumberFormat } from './cell-style.js';
import {
  hasCalendarDate,
  parseFormatCode,
  renderNumericValue,
  renderTextValue,
  type ParsedFormat,
} from './format-code.js';

const epochOf = (wb: Workbook): ExcelEpoch => (wb.date1904 ? 'mac' : 'windows');

const renderValue = (format: ParsedFormat, value: CellValue, epoch: ExcelEpoch): string | undefined => {
  if (isRichTextValue(value)) return renderTextValue(format, richTextToString(value.runs));
  if (typeof value === 'string') return renderTextValue(format, value);
  if (value instanceof Date) return renderNumericValue(format, dateToExcel(value, { epoch }), epoch);
  // A span that is not a finite number of milliseconds has no serial to put
  // through the format, and `durationToExcel` throws on one.
  if (isDurationValue(value)) {
    return Number.isFinite(value.ms) ? renderNumericValue(format, durationToExcel(value.ms), epoch) : undefined;
  }
  if (typeof value === 'number') return renderNumericValue(format, value, epoch);
  return undefined;
};

const displayText = (code: string, value: CellValue, epoch: ExcelEpoch): string => {
  if (value === null) return '';
  // A number format never applies to a boolean or an error token: Excel prints
  // the literal `TRUE` / `FALSE` and the error code whatever the format says.
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (isErrorValue(value)) return value.code;
  // An invalid Date has no serial either, and the `cellValueAsString` fallback
  // would throw on it rather than degrade, so it reads as its own string form.
  if (value instanceof Date && Number.isNaN(value.getTime())) return String(value);
  if (isFormulaValue(value)) {
    if (value.cachedValue === undefined) return '';
    if (value.cachedValueType === 'error') return String(value.cachedValue);
    return displayText(code, value.cachedValue, epoch);
  }
  const format = parseFormatCode(code);
  if (format === undefined) return cellValueAsString(value);
  return renderValue(format, value, epoch) ?? cellValueAsString(value);
};

/**
 * The text Excel prints in `c`: the cell's value put through the number format
 * its style points at.
 *
 * This is not a shorter `cellValueAsString`. That one takes a `CellValue` and
 * answers "what is this value in JavaScript terms", so it never sees the
 * stylesheet and cannot know that `0.5` is a percentage or that `45365` is a
 * date. This one takes the workbook and the cell, resolves the cell's `numFmt`,
 * and answers "what would somebody looking at this file in Excel read". Pick
 * `cellValueAsString` when you want the value and this when you want the text:
 * writing a sheet out as CSV, rendering it as HTML, diffing two workbooks by
 * what they show, or logging a cell for a human. Choosing wrong fails quietly,
 * which is why both say so.
 *
 * Covered:
 *
 * - `General`, at the 15 significant digits Excel stores. This is what turns
 *   `=1.1+2.2` into `3.3` rather than `3.3000000000000003`.
 * - Every code in `BUILTIN_FORMATS`, including the accounting codes,
 *   scientific notation and the two fraction codes. The one exception is that
 *   catalogue's entry for numFmtId 44, whose stored form is missing the `;`
 *   section separators Excel's own accounting code carries, so it reads as a
 *   text layout.
 * - Custom numeric codes: `0` `#` `?` digit placeholders, thousands grouping,
 *   trailing commas that divide by a thousand, `%`, `E+`/`E-` (including the
 *   engineering step `##0.0E+0`), `?/?` fractions with a placeholder or a
 *   spelled-out denominator, quoted and `\`-escaped literals, `_x` width
 *   reservations, currency symbols, and colour codes (which change nothing
 *   about the text).
 * - The positive / negative / zero / text sections of a multi-section code. A
 *   lone section takes the minus sign in front of it; a value that rounds to
 *   zero prints without one.
 * - Dates and times: `y` `m` `d` `h` `s` runs, `mmm` / `mmmm` month names,
 *   `ddd` / `dddd` weekday names, `AM/PM` and `A/P`, fractional seconds, and
 *   the elapsed forms `[h]` `[mm]` `[ss]`.
 * - Comparison sections (`[<1000]0;[<1000000]0.0,"K";0.0,,"M"`), picked the
 *   way Excel does; a value no section takes has no reading.
 * - The 1900 system's phantom days: serial 0 prints as 1/0/1900 and 60 as
 *   2/29/1900, as in Excel.
 * - Booleans as `TRUE` / `FALSE`, error cells as their token, rich text as its
 *   runs joined, a formula cell from the value Excel cached for it, and an
 *   empty cell as `''`.
 *
 * Not covered, and the cell falls back to `cellValueAsString` so the reading is
 * a plain coercion rather than a guess:
 *
 * - A section splicing two numeric layouts together (`0.00" ("0.00")"`).
 * - Calendar and numbering modifiers: era tokens (`g`, `e`, `b`) and bracket
 *   groups such as `[DBNum1]` that replace the digits themselves.
 * - More than four sections, or an unterminated `"` or `[`.
 *
 * Three more places where the text is a judgement call, because Excel's own
 * answer is not recoverable from the cell:
 *
 * - `*x` repeats `x` until the column is full, so nothing is printed for it.
 * - Month and weekday names come out in English; a `[$-409]`-style locale tag
 *   is read for its currency symbol and otherwise ignored.
 * - How wide the column is decides how many digits `General` shows and when it
 *   gives up and switches to scientific notation. Every stored digit is kept.
 *
 * A `Date` cell value is converted to its serial and run through the format, so
 * it reads as a date only if the cell carries a date format. That is the same
 * rule the file itself follows: the worksheet stores a serial and the style
 * decides what it means. Use `setCellAsDate` on the write side.
 */
export function getCellDisplayText(wb: Workbook, c: Cell): string {
  return displayText(getCellNumberFormat(wb, c), c.value, epochOf(wb));
}

/**
 * The `Date` a date-formatted cell stands for, or `undefined` when the cell has
 * no date reading.
 *
 * Excel stores a date as a plain number of days since the workbook epoch, so
 * nothing in a loaded cell's value distinguishes `45365` the date from `45365`
 * the count: the cell's number format is the only evidence, and
 * `cellValueAsDate` deliberately does not consult it. This composes the three
 * steps that reading one otherwise takes, `getCellNumberFormat` then the
 * format's own reading then `excelToDate` with the epoch from `wb.date1904`,
 * into the call that reading a foreign workbook actually needs.
 *
 * - A `Date` value passes through, formatted or not.
 * - A number needs a format that names a day: a year, month or day part.
 *   Without one it is a count, and the answer is `undefined`.
 * - A formula cell is read from the value Excel cached for it.
 * - A time-of-day format (`h:mm`) names a moment inside a day but not which
 *   day, and an elapsed-time format (`[h]:mm`) measures a span rather than
 *   naming a moment, so neither has a date reading. Neither does a `duration`
 *   value: its `ms` is already the span.
 * - The format is read by the parser {@link getCellDisplayText} renders with,
 *   so the two agree on what any one cell is. A code outside the subset that
 *   one documents has no date reading here either.
 *
 * The returned `Date` is built in UTC, which is how `excelToDate` and
 * `dateToExcel` read every serial in this library. Use `getUTCFullYear` and
 * friends, or the local-time accessors will shift the day for anybody east or
 * west of UTC.
 */
export function getCellDate(wb: Workbook, c: Cell): Date | undefined {
  const value = isFormulaValue(c.value) ? c.value.cachedValue : c.value;
  if (value instanceof Date) return value;
  if (typeof value !== 'number' || !Number.isFinite(value)) return undefined;
  const format = parseFormatCode(getCellNumberFormat(wb, c));
  if (format === undefined || !hasCalendarDate(format, value)) return undefined;
  const date = excelToDate(value, { epoch: epochOf(wb) });
  // A serial far outside the range a `Date` covers has no reading to hand back.
  return Number.isNaN(date.getTime()) ? undefined : date;
}
