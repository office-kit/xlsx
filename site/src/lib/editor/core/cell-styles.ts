// The Cell Styles gallery (Home ▸ Styles). Each preset is applied as a direct
// format patch, which is how the styles render; Excel additionally records the
// named style, which this editor does not need for display.

import { makeColor, type Fill } from '@office-kit/xlsx/styles';
import type { MessageKey } from '../i18n/i18n.svelte.ts';
import { format } from './actions.ts';
import type { EditorController } from './controller.svelte.ts';
import type { StylePatch } from './format.ts';

interface Preset {
  readonly id: string;
  readonly label: MessageKey;
  readonly preview: { fill?: string; color?: string; bold?: boolean; border?: string };
  readonly patch: StylePatch;
}

const solid = (rgb: string): Fill => ({ kind: 'pattern', patternType: 'solid', fgColor: makeColor({ rgb }) });

export const CELL_STYLE_PRESETS: readonly Preset[] = [
  { id: 'normal', label: 'csNormal', preview: {}, patch: { fill: null, font: { bold: undefined, italic: undefined, color: makeColor({ theme: 1 }) }, numFmt: 'General', border: () => ({}) } },
  { id: 'bad', label: 'csBad', preview: { fill: '#FFC7CE', color: '#9C0006' }, patch: { fill: solid('FFC7CE'), font: { color: makeColor({ rgb: '9C0006' }) } } },
  { id: 'good', label: 'csGood', preview: { fill: '#C6EFCE', color: '#006100' }, patch: { fill: solid('C6EFCE'), font: { color: makeColor({ rgb: '006100' }) } } },
  { id: 'neutral', label: 'csNeutral', preview: { fill: '#FFEB9C', color: '#9C5700' }, patch: { fill: solid('FFEB9C'), font: { color: makeColor({ rgb: '9C5700' }) } } },
  { id: 'calculation', label: 'csCalculation', preview: { fill: '#F2F2F2', color: '#FA7D00', bold: true }, patch: { fill: solid('F2F2F2'), font: { bold: true, color: makeColor({ rgb: 'FA7D00' }) } } },
  { id: 'check', label: 'csCheckCell', preview: { fill: '#A5A5A5', color: '#FFFFFF', bold: true }, patch: { fill: solid('A5A5A5'), font: { bold: true, color: makeColor({ rgb: 'FFFFFF' }) } } },
  { id: 'input', label: 'csInput', preview: { fill: '#FFCC99', color: '#3F3F76' }, patch: { fill: solid('FFCC99'), font: { color: makeColor({ rgb: '3F3F76' }) } } },
  { id: 'output', label: 'csOutput', preview: { fill: '#F2F2F2', color: '#3F3F3F', bold: true }, patch: { fill: solid('F2F2F2'), font: { bold: true, color: makeColor({ rgb: '3F3F3F' }) } } },
  { id: 'note', label: 'csNote', preview: { fill: '#FFFFCC' }, patch: { fill: solid('FFFFCC') } },
  { id: 'warning', label: 'csWarning', preview: { color: '#FF0000' }, patch: { font: { color: makeColor({ rgb: 'FF0000' }) } } },
  { id: 'title', label: 'csTitle', preview: { color: '#44546A', bold: true }, patch: { font: { size: 18, bold: true, color: makeColor({ theme: 3 }) } } },
  { id: 'heading1', label: 'csHeading1', preview: { color: '#44546A', bold: true, border: '3px solid #4472C4' }, patch: { font: { size: 15, bold: true, color: makeColor({ theme: 3 }) }, border: (b, e) => (e.bottom ? { ...b, bottom: { style: 'thick', color: makeColor({ theme: 4 }) } } : b) } },
  { id: 'heading2', label: 'csHeading2', preview: { color: '#44546A', bold: true, border: '3px solid #A9C4EB' }, patch: { font: { size: 13, bold: true, color: makeColor({ theme: 3 }) }, border: (b, e) => (e.bottom ? { ...b, bottom: { style: 'thick', color: makeColor({ theme: 4, tint: 0.5 }) } } : b) } },
  { id: 'total', label: 'csTotal', preview: { bold: true, border: '3px double #4472C4' }, patch: { font: { bold: true }, border: (b, e) => ({ ...b, ...(e.top ? { top: { style: 'thin', color: makeColor({ theme: 4 }) } } : {}), ...(e.bottom ? { bottom: { style: 'double', color: makeColor({ theme: 4 }) } } : {}) }) } },
  { id: 'accent1', label: 'csAccent1', preview: { fill: '#4472C4', color: '#FFFFFF' }, patch: { fill: { kind: 'pattern', patternType: 'solid', fgColor: makeColor({ theme: 4 }) }, font: { color: makeColor({ theme: 0 }) } } },
  { id: 'accent2', label: 'csAccent2', preview: { fill: '#ED7D31', color: '#FFFFFF' }, patch: { fill: { kind: 'pattern', patternType: 'solid', fgColor: makeColor({ theme: 5 }) }, font: { color: makeColor({ theme: 0 }) } } },
  { id: 'comma', label: 'csComma', preview: {}, patch: { numFmt: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)' } },
  { id: 'currency', label: 'csCurrency', preview: {}, patch: { numFmt: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)' } },
  { id: 'percent', label: 'csPercent', preview: {}, patch: { numFmt: '0%' } },
];

export function applyCellStylePreset(ctl: EditorController, id: string): void {
  const p = CELL_STYLE_PRESETS.find((x) => x.id === id);
  if (p) format(ctl, p.patch, 'Cell Style');
}
