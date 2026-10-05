// The Number Format drop-down on the Home tab and the categories of the
// Format Cells ▸ Number tab.

import type { MessageKey } from '../i18n/i18n.svelte.ts';

export interface NumberFormatPreset {
  readonly id: string;
  readonly label: MessageKey;
  readonly code: (order: 'mdy' | 'ymd') => string;
}

export const NUMBER_FORMAT_PRESETS: readonly NumberFormatPreset[] = [
  { id: 'general', label: 'nfGeneral', code: () => 'General' },
  { id: 'number', label: 'nfNumber', code: () => '0.00' },
  { id: 'currency', label: 'nfCurrency', code: (o) => (o === 'ymd' ? '"¥"#,##0;[Red]"¥"\\-#,##0' : '"$"#,##0.00') },
  { id: 'accounting', label: 'nfAccounting', code: (o) => (o === 'ymd' ? '_ "¥"* #,##0_ ;_ "¥"* \\-#,##0_ ;_ "¥"* "-"_ ;_ @_ ' : '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)') },
  { id: 'shortDate', label: 'nfShortDate', code: (o) => (o === 'ymd' ? 'yyyy/m/d' : 'm/d/yyyy') },
  { id: 'longDate', label: 'nfLongDate', code: (o) => (o === 'ymd' ? 'yyyy"年"m"月"d"日"' : '[$-x-sysdate]dddd, mmmm dd, yyyy') },
  { id: 'time', label: 'nfTime', code: (o) => (o === 'ymd' ? 'h:mm:ss' : '[$-x-systime]h:mm:ss AM/PM') },
  { id: 'percentage', label: 'nfPercentage', code: () => '0.00%' },
  { id: 'fraction', label: 'nfFraction', code: () => '# ?/?' },
  { id: 'scientific', label: 'nfScientific', code: () => '0.00E+00' },
  { id: 'text', label: 'nfText', code: () => '@' },
];

/** Which Home-tab category a format code belongs to (for the drop-down's current value). */
export function numberFormatCategory(code: string): string {
  if (code === 'General') return 'general';
  if (code === '@') return 'text';
  if (/E[+-]/i.test(code)) return 'scientific';
  if (code.includes('%')) return 'percentage';
  if (code.includes('?/')) return 'fraction';
  if (/_\(?["$¥€£]*\*|\* #/.test(code)) return 'accounting';
  if (/[$¥€£]/.test(code)) return 'currency';
  const stripped = code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
  if (/[dmy]/i.test(stripped) && !/^[hms:.\sAPM/]+$/i.test(stripped)) return /dddd|mmmm|年/.test(code) ? 'longDate' : 'shortDate';
  if (/[hs]/i.test(stripped)) return 'time';
  if (/[0#]/.test(code)) return 'number';
  return 'custom';
}

/** Categories and built-in codes for Format Cells ▸ Number. */
export const FORMAT_CELLS_CATEGORIES = [
  'general',
  'number',
  'currency',
  'accounting',
  'date',
  'time',
  'percentage',
  'fraction',
  'scientific',
  'text',
  'special',
  'custom',
] as const;

export type FormatCategory = (typeof FORMAT_CELLS_CATEGORIES)[number];

export const DATE_CODES = ['m/d/yyyy', 'yyyy/m/d', 'yyyy-mm-dd', 'd-mmm', 'd-mmm-yy', 'dd-mmm-yy', 'mmm-yy', 'mmmm-yy', 'mmmm d, yyyy', 'dddd, mmmm d, yyyy', 'm/d/yy h:mm', 'yyyy"年"m"月"d"日"', 'm"月"d"日"', 'ggge"年"m"月"d"日"'];
export const TIME_CODES = ['h:mm', 'h:mm:ss', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'mm:ss', 'mm:ss.0', '[h]:mm:ss', 'h"時"mm"分"', 'h"時"mm"分"ss"秒"'];
export const FRACTION_CODES = ['# ?/?', '# ??/??', '# ???/???', '# ?/2', '# ?/4', '# ?/8', '# ??/16', '# ?/10', '# ??/100'];
export const SPECIAL_CODES = ['00000', '00000-0000', '[<=9999999]###-####;(###) ###-####', '000-00-0000', '[$-ja-JP]ggge"年"m"月"d"日"', '[DBNum1][$-ja-JP]General', '[DBNum3][$-ja-JP]General'];
export const CUSTOM_CODES = [
  'General',
  '0',
  '0.00',
  '#,##0',
  '#,##0.00',
  '#,##0_);(#,##0)',
  '#,##0_);[Red](#,##0)',
  '#,##0.00_);(#,##0.00)',
  '#,##0.00_);[Red](#,##0.00)',
  '"$"#,##0_);("$"#,##0)',
  '"$"#,##0.00_);[Red]("$"#,##0.00)',
  '0%',
  '0.00%',
  '0.00E+00',
  '##0.0E+0',
  '# ?/?',
  '# ??/??',
  'm/d/yyyy',
  'd-mmm-yy',
  'd-mmm',
  'mmm-yy',
  'h:mm AM/PM',
  'h:mm:ss AM/PM',
  'h:mm',
  'h:mm:ss',
  'm/d/yyyy h:mm',
  'mm:ss',
  'mm:ss.0',
  '@',
  '[h]:mm:ss',
  '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
  '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
];

/** Build a Number/Currency/Percentage code from the Format Cells options. */
export function buildNumberCode(opts: { category: 'number' | 'currency' | 'accounting' | 'percentage' | 'scientific'; decimals: number; thousands?: boolean; symbol?: string; negative?: 0 | 1 | 2 | 3 }): string {
  const dec = opts.decimals > 0 ? `.${'0'.repeat(opts.decimals)}` : '';
  switch (opts.category) {
    case 'percentage':
      return `0${dec}%`;
    case 'scientific':
      return `0${dec}E+00`;
    case 'accounting': {
      const sym = opts.symbol ? `"${opts.symbol}"` : '';
      return `_(${sym}* #,##0${dec}_);_(${sym}* \\(#,##0${dec}\\);_(${sym}* "-"${'?'.repeat(opts.decimals)}_);_(@_)`;
    }
    case 'currency':
    case 'number': {
      const sym = opts.category === 'currency' && opts.symbol ? `"${opts.symbol}"` : '';
      const base = `${sym}${opts.category === 'currency' || opts.thousands ? '#,##0' : '0'}${dec}`;
      switch (opts.negative ?? 0) {
        case 1:
          return `${base};[Red]${base}`;
        case 2:
          return `${base}_);(${base})`;
        case 3:
          return `${base}_);[Red](${base})`;
        default:
          return opts.category === 'currency' ? `${base};-${base}` : base;
      }
    }
  }
}
