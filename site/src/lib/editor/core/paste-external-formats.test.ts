import { getCell } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import { msoClassFormats, msoNumberFormat, pasteFromEvent, unescapeMsoFormat } from './clipboard.ts';
import { EditorController } from './controller.svelte.ts';

const pasteText = (ctl: EditorController, text: string) =>
  pasteFromEvent(ctl, { clipboardData: { getData: (type: string) => (type === 'text/plain' ? text : '') } } as unknown as ClipboardEvent);

it('plain-text paste gives dates, percentages and grouped numbers the format typing would', () => {
  const ctl = new EditorController();
  pasteText(ctl, '2024/3/1\t12.5%\t1,234\tplain\r\n');
  const ws = ctl.doc.ws;
  const fmt = (col: number) => ctl.doc.styles.get(getCell(ws, 1, col)?.styleId ?? 0).numFmt;
  expect(getCell(ws, 1, 1)?.value).toBe(45352);
  expect(fmt(1)).toBe('m/d/yy');
  expect(getCell(ws, 1, 2)?.value).toBe(0.125);
  expect(fmt(2)).toBe('0.0%');
  expect(getCell(ws, 1, 3)?.value).toBe(1234);
  expect(fmt(3)).toBe('#,##0');
  expect(fmt(4)).toBe('General');
});

// The style block is what Excel for Mac put on the clipboard for a copied range.
it("reads Excel's CSS-escaped number formats per class", () => {
  const css = `.xl63 {mso-number-format:"yyyy\\/m\\/d";}
.xl65 {mso-number-format:"\\0022¥\\0022\\#\\,\\#\\#0\\.00";}
.xl66 {mso-number-format:"\\@";}
.xl67 {font-weight:700; background:yellow;}
.xl68 {mso-number-format:"0\\;\\[Red\\]\\\\\\(0\\\\\\)";}`;
  expect(Object.fromEntries(msoClassFormats(css))).toEqual({
    xl63: 'yyyy/m/d',
    xl65: '"¥"#,##0.00',
    xl66: '@',
    xl68: '0;[Red]\\(0\\)',
  });
  expect(msoNumberFormat('mso-number-format:General;')).toBeUndefined();
  expect(unescapeMsoFormat('\\0022A\\0022')).toBe('"A"');
});
