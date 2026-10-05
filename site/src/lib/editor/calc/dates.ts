// Date serials. Excel's 1900 system counts 1900-01-01 as day 1 and keeps
// Lotus 1-2-3's phantom 1900-02-29 as day 60, so every serial before 61 is
// one day off a real calendar. The 1904 system (old Mac workbooks) counts
// 1904-01-01 as day 0 and has no phantom day.

const DAY_MS = 86_400_000;
const EPOCH_1899_12_30 = Date.UTC(1899, 11, 30);
const EPOCH_1904 = Date.UTC(1904, 0, 1);
/** 1900-03-01, the first serial on which the 1900 system matches the real calendar. */
const FIRST_TRUE_SERIAL = 61;
const PHANTOM_LEAP_DAY = 60;
/** Serial of 1904-01-01 in the 1900 system. */
export const DATE1904_OFFSET = 1462;
/** 9999-12-31, the last date Excel accepts. */
export const MAX_DATE_SERIAL = 2_958_465;

export interface Ymd {
  readonly y: number;
  readonly m: number;
  readonly d: number;
}

const daysToSerial1900 = (days: number): number => (days < FIRST_TRUE_SERIAL ? days - 1 : days);

/**
 * DATE(y, m, d) with Excel's overflow rules: months past 12 roll into later
 * years and days past the month's end roll forward — counted in Excel's
 * calendar, so DATE(1900,3,0) is the phantom 1900-02-29.
 */
export function serialFromYmd(y: number, m: number, d: number, date1904: boolean): number {
  const months = y * 12 + (m - 1);
  const year = Math.floor(months / 12);
  const month = months - year * 12;
  const first = Date.UTC(year, month, 1);
  // Date.UTC maps years 0–99 to 1900–1999; setUTCFullYear does not.
  const firstDate = new Date(first);
  firstDate.setUTCFullYear(year);
  const days = Math.round((firstDate.getTime() - (date1904 ? EPOCH_1904 : EPOCH_1899_12_30)) / DAY_MS);
  return (date1904 ? days : daysToSerial1900(days)) + d - 1;
}

export function ymdFromSerial(serial: number, date1904: boolean): Ymd {
  const s = Math.floor(serial);
  if (!date1904) {
    if (s === 0) return { y: 1900, m: 1, d: 0 };
    if (s === PHANTOM_LEAP_DAY) return { y: 1900, m: 2, d: 29 };
  }
  const days = date1904 ? s : s < PHANTOM_LEAP_DAY ? s + 1 : s;
  const date = new Date((date1904 ? EPOCH_1904 : EPOCH_1899_12_30) + days * DAY_MS);
  return { y: date.getUTCFullYear(), m: date.getUTCMonth() + 1, d: date.getUTCDate() };
}

/** 0 = Sunday … 6 = Saturday, on Excel's calendar (so serial 1 is a Sunday). */
export function dayOfWeek(serial: number, date1904: boolean): number {
  const s = Math.floor(serial) + (date1904 ? DATE1904_OFFSET : 0);
  return (((s - 1) % 7) + 7) % 7;
}

export function isLeapYear(y: number): boolean {
  return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
}

export function daysInMonth(y: number, m: number): number {
  if (m === 2) return isLeapYear(y) ? 29 : 28;
  return new Date(Date.UTC(2001, m, 0)).getUTCDate();
}

/** Serial for a wall-clock moment, read in the local time zone like Excel's NOW(). */
export function serialFromDate(date: Date, date1904: boolean): number {
  const localMs = date.getTime() - date.getTimezoneOffset() * 60_000;
  const serial = (localMs - EPOCH_1899_12_30) / DAY_MS;
  return date1904 ? serial - DATE1904_OFFSET : serial;
}

/** Split a time fraction into h / m / s, rounding to the nearest second like Excel's HOUR / MINUTE / SECOND. */
export function hmsFromSerial(serial: number): { h: number; mi: number; s: number } {
  const fraction = serial - Math.floor(serial);
  let total = Math.round(fraction * 86_400);
  if (total >= 86_400) total = 0;
  return { h: Math.floor(total / 3600), mi: Math.floor((total % 3600) / 60), s: total % 60 };
}

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

const monthFromName = (name: string): number | undefined => {
  const lower = name.toLowerCase();
  if (lower.length < 3) return undefined;
  const idx = MONTHS.findIndex((m) => lower.startsWith(m));
  if (idx < 0) return undefined;
  const full = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'][idx] ?? '';
  return full.startsWith(lower) || lower === 'sept' ? idx + 1 : undefined;
};

const TIME = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?\s*(am|pm|a|p)?$/i;
const TIME_HOUR_ONLY = /^(\d{1,2})\s*(am|pm|a|p)$/i;

const parseTime = (text: string): number | undefined => {
  const t = TIME.exec(text);
  let h: number;
  let mi = 0;
  let s = 0;
  let ampm: string | undefined;
  if (t !== null) {
    h = Number(t[1]);
    mi = Number(t[2]);
    s = t[3] !== undefined ? Number(t[3]) : 0;
    ampm = t[4];
  } else {
    const ho = TIME_HOUR_ONLY.exec(text);
    if (ho === null) return undefined;
    h = Number(ho[1]);
    ampm = ho[2];
  }
  if (ampm !== undefined) {
    if (h < 1 || h > 12) return undefined;
    const pm = ampm.toLowerCase().startsWith('p');
    h = (h % 12) + (pm ? 12 : 0);
  }
  if (mi >= 60 || s >= 60) return undefined;
  return (h * 3600 + mi * 60 + s) / 86_400;
};

const twoDigitYear = (y: number, digits: number): number => (digits <= 2 ? (y < 30 ? 2000 + y : 1900 + y) : y);

const validYmd = (y: number, m: number, d: number): boolean => {
  if (m < 1 || m > 12 || d < 1) return false;
  // The phantom 1900-02-29 is a date Excel accepts.
  if (y === 1900 && m === 2) return d <= 29;
  return d <= daysInMonth(y, m);
};

const parseDatePart = (text: string, currentYear: number): Ymd | undefined => {
  let m: RegExpExecArray | null;
  // 2024-01-15, 2024/1/15
  if ((m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(text)) !== null) {
    return { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) };
  }
  // 1/15/2024, 1-15-24 (US order, as Excel en-US reads it)
  if ((m = /^(\d{1,2})[-/](\d{1,2})[-/](\d{1,4})$/.exec(text)) !== null) {
    const yText = m[3] ?? '';
    return { y: twoDigitYear(Number(yText), yText.length), m: Number(m[1]), d: Number(m[2]) };
  }
  // 1/15 → this year
  if ((m = /^(\d{1,2})[-/](\d{1,2})$/.exec(text)) !== null) {
    return { y: currentYear, m: Number(m[1]), d: Number(m[2]) };
  }
  // 15-Jan-2024, 15 January 2024, 15-Jan
  if ((m = /^(\d{1,2})[-\s/]([A-Za-z]+)\.?(?:[-\s/,]+(\d{1,4}))?$/.exec(text)) !== null) {
    const month = monthFromName(m[2] ?? '');
    if (month === undefined) return undefined;
    const yText = m[3];
    return { y: yText === undefined ? currentYear : twoDigitYear(Number(yText), yText.length), m: month, d: Number(m[1]) };
  }
  // Jan 15, 2024 / January 15 2024 / Jan 15
  if ((m = /^([A-Za-z]+)\.?[-\s/]+(\d{1,2})(?:(?:,\s*|[-\s/]+)(\d{1,4}))?$/.exec(text)) !== null) {
    const month = monthFromName(m[1] ?? '');
    if (month === undefined) return undefined;
    const yText = m[3];
    const dayOrYear = Number(m[2]);
    if (yText === undefined && (m[2] ?? '').length > 2) return { y: dayOrYear, m: month, d: 1 };
    return { y: yText === undefined ? currentYear : twoDigitYear(Number(yText), yText.length), m: month, d: dayOrYear };
  }
  // Jan-2024, January 2024 → first of the month
  if ((m = /^([A-Za-z]+)\.?[-\s/]+(\d{4})$/.exec(text)) !== null) {
    const month = monthFromName(m[1] ?? '');
    if (month === undefined) return undefined;
    return { y: Number(m[2]), m: month, d: 1 };
  }
  return undefined;
};

/**
 * Read a date and/or time typed as text, the way Excel's DATEVALUE / VALUE
 * do in an en-US locale. Returns the date serial (0 when only a time was
 * given) and the time fraction, or undefined when the text is not a date.
 */
export function parseDateTimeText(
  input: string,
  date1904: boolean,
  currentYear: number,
): { date: number; time: number; hasDate: boolean } | undefined {
  const text = input.trim();
  if (text === '') return undefined;
  const time = parseTime(text);
  if (time !== undefined) return { date: 0, time, hasDate: false };
  // Date followed by a time: split at the last space that leaves a valid time.
  const parts = text.split(/\s+/);
  for (let cut = parts.length; cut >= 1; cut--) {
    const datePart = parts.slice(0, cut).join(' ');
    const timePart = parts.slice(cut).join(' ');
    const t = timePart === '' ? 0 : parseTime(timePart);
    if (t === undefined) continue;
    const ymd = parseDatePart(datePart, currentYear);
    if (ymd === undefined) continue;
    if (!validYmd(ymd.y, ymd.m, ymd.d) || ymd.y > 9999) return undefined;
    const serial = serialFromYmd(ymd.y, ymd.m, ymd.d, date1904);
    if (serial < (date1904 ? 0 : 1)) return undefined;
    return { date: serial, time: t, hasDate: true };
  }
  return undefined;
}
