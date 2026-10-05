// Date & Time functions. Serials follow the workbook's date system; see
// dates.ts for the 1900 leap-year quirk.

import {
  dayOfWeek,
  daysInMonth,
  hmsFromSerial,
  isLeapYear,
  MAX_DATE_SERIAL,
  parseDateTimeText,
  serialFromYmd,
  type Ymd,
  ymdFromSerial,
} from '../dates.ts';
import type { FnContext, FunctionSpec } from '../function-spec.ts';
import { type CalcError, type CalcValue, ERRORS, isError } from '../types.ts';
import { collectNumbers, num, nums, optBool, scalar, str, trunc } from './helpers.ts';

const C = 'Date & Time' as const;

/** A date argument as a whole serial, rejecting what Excel rejects. */
const dateArg = (v: CalcValue | undefined, ctx: FnContext): number | CalcError => {
  const n = num(v, ctx);
  if (isError(n)) return n;
  // 2958465.99 is still 12/31/9999; 2958466 is past the last day.
  if (n < 0 || n >= MAX_DATE_SERIAL + 1) return ERRORS.NUM;
  return n;
};

const ymdOf = (serial: number, ctx: FnContext): Ymd => ymdFromSerial(serial, ctx.host.date1904);

const serialOf = (y: number, m: number, d: number, ctx: FnContext): number => serialFromYmd(y, m, d, ctx.host.date1904);

const part = (name: string, description: string, fn: (serial: number, ctx: FnContext) => number): FunctionSpec => ({
  name,
  category: C,
  syntax: `${name}(serial_number)`,
  description,
  minArgs: 1,
  maxArgs: 1,
  impl: (args, ctx) => {
    const s = dateArg(args[0], ctx);
    return isError(s) ? s : fn(s, ctx);
  },
});

/** Monday = 0 … Sunday = 6 for a weekend-mask string. */
const mondayIndex = (serial: number, ctx: FnContext): number => (dayOfWeek(serial, ctx.host.date1904) + 6) % 7;

const WEEKEND_CODES: Readonly<Record<number, readonly number[]>> = {
  1: [5, 6],
  2: [6, 0],
  3: [0, 1],
  4: [1, 2],
  5: [2, 3],
  6: [3, 4],
  7: [4, 5],
  11: [6],
  12: [0],
  13: [1],
  14: [2],
  15: [3],
  16: [4],
  17: [5],
};

/** Seven booleans, Monday first: true where the day is a weekend day. */
const weekendMask = (v: CalcValue | undefined, ctx: FnContext): boolean[] | CalcError => {
  const raw = scalar(v);
  if (raw === null) return [false, false, false, false, false, true, true];
  if (typeof raw === 'string') {
    if (!/^[01]{7}$/.test(raw)) return ERRORS.VALUE;
    return [...raw].map((ch) => ch === '1');
  }
  const n = num(raw, ctx);
  if (isError(n)) return n;
  const days = WEEKEND_CODES[trunc(n)];
  if (days === undefined) return ERRORS.NUM;
  return Array.from({ length: 7 }, (_, i) => days.includes(i));
};

const holidaySet = (v: CalcValue | undefined, ctx: FnContext): Set<number> | CalcError => {
  if (v === undefined || v === null) return new Set();
  const values = collectNumbers([v], ctx);
  if (isError(values)) return values;
  return new Set(values.map(Math.floor));
};

const networkDays = (start: number, end: number, weekend: readonly boolean[], holidays: ReadonlySet<number>, ctx: FnContext): number => {
  const sign = start <= end ? 1 : -1;
  const from = Math.min(start, end);
  const to = Math.max(start, end);
  const workPerWeek = weekend.filter((w) => !w).length;
  const span = to - from + 1;
  const weeks = Math.floor(span / 7);
  let count = weeks * workPerWeek;
  for (let d = from + weeks * 7; d <= to; d++) if (!weekend[mondayIndex(d, ctx)]) count++;
  for (const h of holidays) if (h >= from && h <= to && !weekend[mondayIndex(h, ctx)]) count--;
  return sign * count;
};

const workday = (start: number, days: number, weekend: readonly boolean[], holidays: ReadonlySet<number>, ctx: FnContext): number | CalcError => {
  if (weekend.every(Boolean)) return ERRORS.VALUE;
  const step = days >= 0 ? 1 : -1;
  let d = start;
  let left = Math.abs(days);
  while (left > 0) {
    d += step;
    if (d < 0 || d > MAX_DATE_SERIAL) return ERRORS.NUM;
    if (!weekend[mondayIndex(d, ctx)] && !holidays.has(d)) left--;
  }
  return d;
};

const isLastOfFeb = (ymd: Ymd): boolean => ymd.m === 2 && ymd.d === daysInMonth(ymd.y, 2);

const days360 = (a: Ymd, b: Ymd, european: boolean): number => {
  let d1 = a.d;
  let d2 = b.d;
  let m2 = b.m;
  let y2 = b.y;
  if (european) {
    if (d1 === 31) d1 = 30;
    if (d2 === 31) d2 = 30;
  } else {
    // What Excel does, which is not quite the NASD rule its documentation
    // describes: an end date on the last day of February is left alone.
    if (d1 === 31 || isLastOfFeb(a)) d1 = 30;
    if (d2 === 31) {
      if (d1 < 30) {
        d2 = 1;
        m2 += 1;
        if (m2 > 12) {
          m2 = 1;
          y2 += 1;
        }
      } else {
        d2 = 30;
      }
    }
  }
  return (y2 - a.y) * 360 + (m2 - a.m) * 30 + (d2 - d1);
};

const yearFrac = (s: number, e: number, basis: number, ctx: FnContext): number | CalcError => {
  const [start, end] = s <= e ? [s, e] : [e, s];
  const a = ymdOf(start, ctx);
  const b = ymdOf(end, ctx);
  switch (basis) {
    case 0: {
      let d1 = a.d;
      let d2 = b.d;
      if (isLastOfFeb(a)) {
        d1 = 30;
        if (isLastOfFeb(b)) d2 = 30;
      }
      if (d1 === 31) d1 = 30;
      if (d2 === 31 && d1 === 30) d2 = 30;
      return ((b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1)) / 360;
    }
    case 1: {
      const days = end - start;
      const withinYear = a.y === b.y || (a.y + 1 === b.y && (a.m > b.m || (a.m === b.m && a.d >= b.d)));
      if (withinYear) {
        let yearLen: number;
        if (a.y === b.y) {
          yearLen = isLeapYear(a.y) ? 366 : 365;
        } else {
          const feb29 = (isLeapYear(a.y) && a.m <= 2) || (isLeapYear(b.y) && (b.m > 2 || (b.m === 2 && b.d === 29)));
          yearLen = feb29 ? 366 : 365;
        }
        return days / yearLen;
      }
      const years = b.y - a.y + 1;
      const total = serialOf(b.y + 1, 1, 1, ctx) - serialOf(a.y, 1, 1, ctx);
      return days / (total / years);
    }
    case 2:
      return (end - start) / 360;
    case 3:
      return (end - start) / 365;
    case 4:
      return days360(a, b, true) / 360;
    default:
      return ERRORS.NUM;
  }
};

const isoWeek = (serial: number, ctx: FnContext): number => {
  const { y, m, d } = ymdOf(serial, ctx);
  const date = new Date(Date.UTC(y, m - 1, d));
  const day = (date.getUTCDay() + 6) % 7;
  date.setUTCDate(date.getUTCDate() - day + 3);
  const firstThursday = new Date(Date.UTC(date.getUTCFullYear(), 0, 4));
  return 1 + Math.round(((date.getTime() - firstThursday.getTime()) / 86_400_000 - 3 + ((firstThursday.getUTCDay() + 6) % 7)) / 7);
};

export const DATETIME_FUNCTIONS: FunctionSpec[] = [
  {
    name: 'DATE',
    category: C,
    syntax: 'DATE(year, month, day)',
    description: 'Returns the serial number of a particular date.',
    minArgs: 3,
    maxArgs: 3,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined, undefined]);
      if (isError(v)) return v;
      let y = trunc(v[0] ?? 0);
      const m = trunc(v[1] ?? 0);
      const d = trunc(v[2] ?? 0);
      if (y < 0 || y >= 10000) return ERRORS.NUM;
      if (y < 1900) y += 1900;
      const serial = serialOf(y, m, d, ctx);
      return serial < 0 || serial > MAX_DATE_SERIAL ? ERRORS.NUM : serial;
    },
  },
  {
    name: 'TIME',
    category: C,
    syntax: 'TIME(hour, minute, second)',
    description: 'Returns the serial number of a particular time.',
    minArgs: 3,
    maxArgs: 3,
    impl: (args, ctx) => {
      const v = nums(args, ctx, [undefined, undefined, undefined]);
      if (isError(v)) return v;
      const total = trunc(v[0] ?? 0) * 3600 + trunc(v[1] ?? 0) * 60 + trunc(v[2] ?? 0);
      if (total < 0 || trunc(v[0] ?? 0) > 32767) return ERRORS.NUM;
      return (total % 86_400) / 86_400;
    },
  },
  {
    name: 'TODAY',
    category: C,
    syntax: 'TODAY()',
    description: 'Returns the serial number of the current date.',
    minArgs: 0,
    maxArgs: 0,
    volatile: true,
    impl: (_args, ctx) => Math.floor(ctx.host.now()),
  },
  {
    name: 'NOW',
    category: C,
    syntax: 'NOW()',
    description: 'Returns the serial number of the current date and time.',
    minArgs: 0,
    maxArgs: 0,
    volatile: true,
    impl: (_args, ctx) => ctx.host.now(),
  },
  part('YEAR', 'Converts a serial number to a year.', (s, ctx) => ymdOf(s, ctx).y),
  part('MONTH', 'Converts a serial number to a month.', (s, ctx) => ymdOf(s, ctx).m),
  part('DAY', 'Converts a serial number to a day of the month.', (s, ctx) => ymdOf(s, ctx).d),
  part('HOUR', 'Converts a serial number to an hour.', (s) => hmsFromSerial(s).h),
  part('MINUTE', 'Converts a serial number to a minute.', (s) => hmsFromSerial(s).mi),
  part('SECOND', 'Converts a serial number to a second.', (s) => hmsFromSerial(s).s),
  {
    name: 'WEEKDAY',
    category: C,
    syntax: 'WEEKDAY(serial_number, [return_type])',
    description: 'Converts a serial number to a day of the week.',
    minArgs: 1,
    maxArgs: 2,
    impl: (args, ctx) => {
      const s = dateArg(args[0], ctx);
      if (isError(s)) return s;
      const t = args[1] === undefined ? 1 : num(args[1], ctx);
      if (isError(t)) return t;
      const type = trunc(t);
      const dow = dayOfWeek(s, ctx.host.date1904);
      if (type === 1) return dow + 1;
      if (type === 2) return ((dow + 6) % 7) + 1;
      if (type === 3) return (dow + 6) % 7;
      if (type >= 11 && type <= 17) {
        const start = (type - 10) % 7;
        return ((dow - start + 7) % 7) + 1;
      }
      return ERRORS.NUM;
    },
  },
  {
    name: 'WEEKNUM',
    category: C,
    syntax: 'WEEKNUM(serial_number, [return_type])',
    description: 'Converts a serial number to a number representing where the week falls numerically within a year.',
    minArgs: 1,
    maxArgs: 2,
    impl: (args, ctx) => {
      const s = dateArg(args[0], ctx);
      if (isError(s)) return s;
      const t = args[1] === undefined ? 1 : num(args[1], ctx);
      if (isError(t)) return t;
      const type = trunc(t);
      if (type === 21) return isoWeek(s, ctx);
      const startDow = type === 1 || type === 17 ? 0 : type === 2 || type === 11 ? 1 : type >= 12 && type <= 16 ? type - 10 : -1;
      if (startDow < 0) return ERRORS.NUM;
      const { y } = ymdOf(s, ctx);
      const jan1 = serialOf(y, 1, 1, ctx);
      const offset = (dayOfWeek(jan1, ctx.host.date1904) - startDow + 7) % 7;
      return Math.floor((Math.floor(s) - jan1 + offset) / 7) + 1;
    },
  },
  part('ISOWEEKNUM', 'Returns the ISO week number of the year for a given date.', isoWeek),
  {
    name: 'EDATE',
    category: C,
    syntax: 'EDATE(start_date, months)',
    description: 'Returns the serial number of the date that is the indicated number of months before or after the start date.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const s = dateArg(args[0], ctx);
      if (isError(s)) return s;
      const m = num(args[1], ctx);
      if (isError(m)) return m;
      const { y, m: month, d } = ymdOf(s, ctx);
      const total = y * 12 + month - 1 + trunc(m);
      const ny = Math.floor(total / 12);
      const nm = total - ny * 12 + 1;
      const r = serialOf(ny, nm, Math.min(d, daysInMonth(ny, nm)), ctx);
      return r < 0 || r > MAX_DATE_SERIAL ? ERRORS.NUM : r;
    },
  },
  {
    name: 'EOMONTH',
    category: C,
    syntax: 'EOMONTH(start_date, months)',
    description: 'Returns the serial number of the last day of the month before or after a specified number of months.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const s = dateArg(args[0], ctx);
      if (isError(s)) return s;
      const m = num(args[1], ctx);
      if (isError(m)) return m;
      const { y, m: month } = ymdOf(s, ctx);
      const r = serialOf(y, month + trunc(m) + 1, 0, ctx);
      return r < 0 || r > MAX_DATE_SERIAL ? ERRORS.NUM : r;
    },
  },
  {
    name: 'DATEDIF',
    category: C,
    syntax: 'DATEDIF(start_date, end_date, unit)',
    description: 'Calculates the number of days, months, or years between two dates.',
    minArgs: 3,
    maxArgs: 3,
    impl: (args, ctx) => {
      const s = dateArg(args[0], ctx);
      if (isError(s)) return s;
      const e = dateArg(args[1], ctx);
      if (isError(e)) return e;
      const unit = str(args[2]);
      if (isError(unit)) return unit;
      const start = Math.floor(s);
      const end = Math.floor(e);
      if (start > end) return ERRORS.NUM;
      const a = ymdOf(start, ctx);
      const b = ymdOf(end, ctx);
      const months = (b.y - a.y) * 12 + (b.m - a.m) - (b.d < a.d ? 1 : 0);
      switch (unit.toUpperCase()) {
        case 'Y':
          return Math.floor(months / 12);
        case 'M':
          return months;
        case 'D':
          return end - start;
        case 'MD': {
          if (b.d >= a.d) return b.d - a.d;
          const pm = b.m === 1 ? 12 : b.m - 1;
          const py = b.m === 1 ? b.y - 1 : b.y;
          return daysInMonth(py, pm) - a.d + b.d;
        }
        case 'YM':
          return months % 12;
        case 'YD': {
          let anchor = serialOf(b.y, a.m, a.d, ctx);
          if (anchor > end) anchor = serialOf(b.y - 1, a.m, a.d, ctx);
          return end - anchor;
        }
        default:
          return ERRORS.NUM;
      }
    },
  },
  {
    name: 'DAYS',
    category: C,
    syntax: 'DAYS(end_date, start_date)',
    description: 'Returns the number of days between two dates.',
    minArgs: 2,
    maxArgs: 2,
    impl: (args, ctx) => {
      const e = dateArg(args[0], ctx);
      if (isError(e)) return e;
      const s = dateArg(args[1], ctx);
      if (isError(s)) return s;
      return Math.floor(e) - Math.floor(s);
    },
  },
  {
    name: 'DAYS360',
    category: C,
    syntax: 'DAYS360(start_date, end_date, [method])',
    description: 'Calculates the number of days between two dates based on a 360-day year.',
    minArgs: 2,
    maxArgs: 3,
    impl: (args, ctx) => {
      const s = dateArg(args[0], ctx);
      if (isError(s)) return s;
      const e = dateArg(args[1], ctx);
      if (isError(e)) return e;
      const european = optBool(args[2], false);
      if (isError(european)) return european;
      return days360(ymdOf(s, ctx), ymdOf(e, ctx), european);
    },
  },
  {
    name: 'NETWORKDAYS',
    category: C,
    syntax: 'NETWORKDAYS(start_date, end_date, [holidays])',
    description: 'Returns the number of whole workdays between two dates.',
    minArgs: 2,
    maxArgs: 3,
    args: ['scalar', 'scalar', 'ref'],
    impl: (args, ctx) => {
      const s = dateArg(args[0], ctx);
      if (isError(s)) return s;
      const e = dateArg(args[1], ctx);
      if (isError(e)) return e;
      const h = holidaySet(args[2], ctx);
      if (isError(h)) return h;
      return networkDays(Math.floor(s), Math.floor(e), [false, false, false, false, false, true, true], h, ctx);
    },
  },
  {
    name: 'NETWORKDAYS.INTL',
    category: C,
    syntax: 'NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])',
    description: 'Returns the number of whole workdays between two dates using parameters to indicate which days are weekend days.',
    minArgs: 2,
    maxArgs: 4,
    args: ['scalar', 'scalar', 'scalar', 'ref'],
    impl: (args, ctx) => {
      const s = dateArg(args[0], ctx);
      if (isError(s)) return s;
      const e = dateArg(args[1], ctx);
      if (isError(e)) return e;
      const mask = weekendMask(args[2], ctx);
      if (isError(mask)) return mask;
      const h = holidaySet(args[3], ctx);
      if (isError(h)) return h;
      return networkDays(Math.floor(s), Math.floor(e), mask, h, ctx);
    },
  },
  {
    name: 'WORKDAY',
    category: C,
    syntax: 'WORKDAY(start_date, days, [holidays])',
    description: 'Returns the serial number of the date before or after a specified number of workdays.',
    minArgs: 2,
    maxArgs: 3,
    args: ['scalar', 'scalar', 'ref'],
    impl: (args, ctx) => {
      const s = dateArg(args[0], ctx);
      if (isError(s)) return s;
      const d = num(args[1], ctx);
      if (isError(d)) return d;
      const h = holidaySet(args[2], ctx);
      if (isError(h)) return h;
      return workday(Math.floor(s), trunc(d), [false, false, false, false, false, true, true], h, ctx);
    },
  },
  {
    name: 'WORKDAY.INTL',
    category: C,
    syntax: 'WORKDAY.INTL(start_date, days, [weekend], [holidays])',
    description: 'Returns the serial number of the date before or after a specified number of workdays with custom weekend parameters.',
    minArgs: 2,
    maxArgs: 4,
    args: ['scalar', 'scalar', 'scalar', 'ref'],
    impl: (args, ctx) => {
      const s = dateArg(args[0], ctx);
      if (isError(s)) return s;
      const d = num(args[1], ctx);
      if (isError(d)) return d;
      const mask = weekendMask(args[2], ctx);
      if (isError(mask)) return mask;
      const h = holidaySet(args[3], ctx);
      if (isError(h)) return h;
      return workday(Math.floor(s), trunc(d), mask, h, ctx);
    },
  },
  {
    name: 'YEARFRAC',
    category: C,
    syntax: 'YEARFRAC(start_date, end_date, [basis])',
    description: 'Returns the year fraction representing the number of whole days between start_date and end_date.',
    minArgs: 2,
    maxArgs: 3,
    impl: (args, ctx) => {
      const s = dateArg(args[0], ctx);
      if (isError(s)) return s;
      const e = dateArg(args[1], ctx);
      if (isError(e)) return e;
      const b = args[2] === undefined ? 0 : num(args[2], ctx);
      if (isError(b)) return b;
      return yearFrac(Math.floor(s), Math.floor(e), trunc(b), ctx);
    },
  },
  {
    name: 'DATEVALUE',
    category: C,
    syntax: 'DATEVALUE(date_text)',
    description: 'Converts a date in the form of text to a serial number.',
    minArgs: 1,
    maxArgs: 1,
    impl: (args, ctx) => {
      const t = str(args[0]);
      if (isError(t)) return t;
      const dt = parseDateTimeText(t, ctx.host.date1904, new Date().getFullYear());
      return dt === undefined || !dt.hasDate ? ERRORS.VALUE : dt.date;
    },
  },
  {
    name: 'TIMEVALUE',
    category: C,
    syntax: 'TIMEVALUE(time_text)',
    description: 'Converts a time in the form of text to a serial number.',
    minArgs: 1,
    maxArgs: 1,
    impl: (args, ctx) => {
      const t = str(args[0]);
      if (isError(t)) return t;
      const dt = parseDateTimeText(t, ctx.host.date1904, new Date().getFullYear());
      // Hours past 24 wrap: TIMEVALUE("25:00") is 1:00.
      return dt === undefined ? ERRORS.VALUE : dt.time % 1;
    },
  },
];
