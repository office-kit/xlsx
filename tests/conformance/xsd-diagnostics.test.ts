import { expect, it } from 'vitest';
import { parseXsdDiagnostics } from './xsd-diagnostics.js';
const detail = "Element 'workbookPr', attribute 'dateCompatibility': The attribute is not allowed.";
it('replays Windows CRLF diagnostics without misclassifying successful batch members', () => {
  const files = new Map([['C:\\Temp\\3.xml', 'xl/workbook.xml'], ['C:\\Temp\\5.xml', 'xl/styles.xml']]);
  expect(parseXsdDiagnostics(`${detail}\r\nC:\\Temp\\3.xml fails to validate\r\nC:\\Temp\\5.xml validates\r\n`, files)).toEqual([
    { part: 'xl/workbook.xml', message: detail },
    { part: 'xl/workbook.xml', message: 'C:\\Temp\\3.xml fails to validate' },
  ]);
});
it('retains prefixed diagnostics and resolves alternate path separators', () => {
  const files = new Map([['C:\\Temp\\3.xml', 'xl/workbook.xml']]);
  const message = `C:/Temp/3.xml:2: ${detail}`;
  expect(parseXsdDiagnostics(`${message}\nC:/Temp/3.xml fails to validate\n`, files)).toEqual([
    { part: 'xl/workbook.xml', message },
    { part: 'xl/workbook.xml', message: 'C:/Temp/3.xml fails to validate' },
  ]);
});
it('keeps unlocated diagnostics inconclusive rather than discarding them', () => {
  expect(parseXsdDiagnostics(`unexpected schema warning\r\n/tmp/1.xml validates\r\n`, new Map([['/tmp/1.xml', 'xl/styles.xml']]))).toEqual([
    { part: '<runner>', message: 'unexpected schema warning' },
  ]);
});
it('attributes successive unprefixed errors to their respective failed summaries', () => {
  expect(parseXsdDiagnostics('first error\n/tmp/1.xml fails to validate\nsecond error\n/tmp/2.xml fails to validate\n', new Map([['/tmp/1.xml', 'first.xml'], ['/tmp/2.xml', 'second.xml']]))).toEqual([
    { part: 'first.xml', message: 'first error' }, { part: 'first.xml', message: '/tmp/1.xml fails to validate' },
    { part: 'second.xml', message: 'second error' }, { part: 'second.xml', message: '/tmp/2.xml fails to validate' },
  ]);
});
it('does not invent an association for unknown files or empty diagnostics', () => {
  expect(parseXsdDiagnostics('error\n/tmp/unknown.xml fails to validate\n', new Map())).toEqual([
    { part: '<runner>', message: 'error' }, { part: '<runner>', message: '/tmp/unknown.xml fails to validate' },
  ]);
  expect(parseXsdDiagnostics('', new Map())).toEqual([]);
});
