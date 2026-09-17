// Payload of a `t="d"` cell, shared by the two worksheet readers so that one
// file is answered the same way whether the caller loads a workbook or streams
// it.
//
// `d` is an ISO 29500 strict cell type (ECMA-376 §18.18.11 ST_CellType): the
// `<v>` holds an ISO 8601 date, time or duration instead of a serial number.
// ECMA-376 §18.17.4 admits the date (B.1.1 / B.2.1), time (B.1.2 / B.2.2) and
// datetime (B.1.3 / B.2.3) profiles, and Excel writes an ISO duration for an
// elapsed-time cell, which openpyxl's reader accepts too.

import { cellLabel, quoteCellText } from './cell-text.js';
import { OpenXmlSchemaError } from './exceptions.js';

const MS_PER_SECOND = 1_000;
const MS_PER_MINUTE = 60_000;
const MS_PER_HOUR = 3_600_000;

/** `YYYY-MM-DD`, optionally `T`-joined to a time, optionally zone-qualified. */
const ISO_DATE =
  /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?)?(Z|[+-]\d{2}:\d{2})?$/;
/** `HH:MM[:SS[.fff]]` with no date. Excel stores a time of day as a fraction of a day. */
const ISO_TIME = /^(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?$/;
/** `PT[nH][nM][n[.fff]S]`, the form Excel writes for an elapsed-time cell. */
const ISO_DURATION = /^PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d{1,3})?)S)?$/;

const intOr = (raw: string | undefined, fallback: number): number =>
  raw === undefined ? fallback : Number.parseInt(raw, 10);

/** `.5` and `.500` are both 500ms. */
const fractionMs = (raw: string | undefined): number =>
  raw === undefined ? 0 : Number.parseInt(raw.padEnd(3, '0'), 10);

/** Minutes a `±HH:MM` suffix puts the wall clock ahead of UTC. `Z` and an absent suffix are 0. */
const offsetMinutes = (raw: string | undefined): number => {
  if (raw === undefined || raw === 'Z') return 0;
  const sign = raw.startsWith('-') ? -1 : 1;
  return sign * (intOr(raw.slice(1, 3), 0) * 60 + intOr(raw.slice(4, 6), 0));
};

/**
 * Value of a `t="d"` cell's `<v>`: a `Date`, a duration for the date-less
 * forms, or `null` when the element is absent or blank, which is an empty cell
 * the way it is under `t="n"`.
 *
 * A value with no zone suffix is read as UTC. Excel writes none, and every
 * `Date` in this model is read and written in UTC, so reading a naive value as
 * local time would shift it by the reader's own timezone offset and again on
 * save. `new Date(string)` cannot be used for the same reason: ECMA-262 reads
 * `2024-03-14` as UTC midnight but `2024-03-14T00:00:00` as local midnight.
 */
export function parseCellDate(
  raw: string | undefined,
  sheet: string,
  col: number,
  row: number,
): Date | { kind: 'duration'; ms: number } | null {
  if (raw === undefined || raw.trim() === '') return null;
  const text = raw.trim();

  const date = ISO_DATE.exec(text);
  if (date) {
    const year = intOr(date[1], Number.NaN);
    const month = intOr(date[2], Number.NaN);
    const day = intOr(date[3], Number.NaN);
    const hour = intOr(date[4], 0);
    const minute = intOr(date[5], 0);
    const second = intOr(date[6], 0);
    const wall = new Date(Date.UTC(year, month - 1, day, hour, minute, second, fractionMs(date[7])));
    // Date.UTC rolls a component past its range into the next one, so a
    // non-existent date like 2024-02-30 arrives here as March 1. Comparing the
    // components back is what rejects it.
    const rolled =
      wall.getUTCFullYear() !== year ||
      wall.getUTCMonth() !== month - 1 ||
      wall.getUTCDate() !== day ||
      wall.getUTCHours() !== hour ||
      wall.getUTCMinutes() !== minute ||
      wall.getUTCSeconds() !== second;
    if (!rolled) return new Date(wall.getTime() - offsetMinutes(date[8]) * MS_PER_MINUTE);
  }

  const time = ISO_TIME.exec(text);
  if (time) {
    const hour = intOr(time[1], 0);
    const minute = intOr(time[2], 0);
    const second = intOr(time[3], 0);
    if (hour <= 23 && minute <= 59 && second <= 59) {
      return {
        kind: 'duration',
        ms: hour * MS_PER_HOUR + minute * MS_PER_MINUTE + second * MS_PER_SECOND + fractionMs(time[4]),
      };
    }
  }

  const duration = ISO_DURATION.exec(text);
  if (duration && (duration[1] !== undefined || duration[2] !== undefined || duration[3] !== undefined)) {
    const seconds = duration[3] === undefined ? 0 : Number.parseFloat(duration[3]);
    const ms =
      intOr(duration[1], 0) * MS_PER_HOUR + intOr(duration[2], 0) * MS_PER_MINUTE + seconds * MS_PER_SECOND;
    return { kind: 'duration', ms: Math.round(ms) };
  }

  const at = cellLabel(sheet, col, row);
  throw new OpenXmlSchemaError(
    `worksheet: <v>${quoteCellText(text)}</v> at ${at} is not an ISO 8601 date, time or duration`,
  );
}
