import fc from 'fast-check';
import { strFromU8, unzipSync } from 'fflate';
import { expect, it } from 'vitest';
import { workbookToBytes } from '../../src/io/save.js';
import { addDefinedName } from '../../src/workbook/defined-names.js';
import { addWorksheet, createWorkbook, duplicateSheet, moveSheet, removeSheet, renameSheet, setActiveSheet, swapSheets } from '../../src/workbook/workbook.js';
import { addExcelTable } from '../../src/worksheet/table.js';
import { setCell, setComment, setHyperlink } from '../../src/worksheet/worksheet.js';
import { required } from './required.js';
import { validateXlsx } from './validate.js';
import { attribute, children, parseDocument, textContent } from './xml-tree.js';

it('keeps local names and active sheet attached through move, swap, insert and removal', () => {
  const wb = createWorkbook();
  for (const name of ['A', 'B', 'C']) addWorksheet(wb, name);
  addDefinedName(wb, { name: 'Local', value: '7', scope: 1 });
  setActiveSheet(wb, 'B');
  moveSheet(wb, 'B', 0);
  expect(wb.definedNames[0]?.scope).toBe(0);
  swapSheets(wb, 'B', 'C');
  expect(wb.definedNames[0]?.scope).toBe(2);
  duplicateSheet(wb, 'A', 'Copy', { index: 0 });
  expect(wb.definedNames[0]?.scope).toBe(3);
  expect(wb.sheets[wb.activeSheetIndex]?.sheet.title).toBe('B');
  removeSheet(wb, 'Copy');
  expect(wb.definedNames[0]?.scope).toBe(2);
  removeSheet(wb, 'B');
  expect(wb.definedNames).toEqual([]);
});

it('generated editing sequences preserve independently tracked names, cells and table uniqueness', async () => {
  await fc.assert(fc.asyncProperty(fc.array(fc.record({
    kind: fc.constantFrom('move', 'swap', 'rename', 'copy', 'remove'),
    first: fc.nat(20), second: fc.nat(20),
  }), { minLength: 1, maxLength: 10 }), async operations => {
    const wb = createWorkbook();
    const expected: Array<{ title: string; value: number }> = [];
    for (let i = 0; i < 3; i++) {
      const title = `S${i}`;
      const ws = addWorksheet(wb, title);
      expected.push({ title, value: i + 1 });
      setCell(ws, 1, 1, 'Value'); setCell(ws, 2, 1, i + 1);
      setCell(ws, 3, 1, { kind: 'formula', t: 'normal', formula: 'A2+1', cachedValue: i + 2 });
      setComment(ws, { ref: 'A1', author: 'QA', text: `note${i}` });
      setHyperlink(ws, 'A1', { target: `https://example.com/${i}` });
      addExcelTable(wb, ws, { name: `Table${i}`, ref: 'A1:A2', columns: ['Value'] });
      addDefinedName(wb, { name: 'Local', value: String(i + 1), scope: i });
    }
    addDefinedName(wb, { name: 'Global', value: '99' });
    let active = required(expected[1]); setActiveSheet(wb, active.title);
    let serial = 0;
    for (const operation of operations) {
      const a = operation.first % expected.length; const b = operation.second % expected.length;
      const source = required(expected[a]);
      if (operation.kind === 'move') {
        moveSheet(wb, source.title, b); expected.splice(a, 1); expected.splice(b, 0, source);
      } else if (operation.kind === 'swap') {
        swapSheets(wb, source.title, required(expected[b]).title);
        const other = required(expected[b]); expected[a] = other; expected[b] = source;
      } else if (operation.kind === 'rename') {
        const title = `Renamed ${serial++}`; renameSheet(wb, source.title, title); source.title = title;
      } else if (operation.kind === 'copy') {
        const title = `Copy ${serial++}`; duplicateSheet(wb, source.title, title, { index: b });
        expected.splice(b, 0, { title, value: source.value });
      } else if (expected.length > 1) {
        removeSheet(wb, source.title); expected.splice(a, 1);
        if (active === source) active = required(expected[Math.min(a, expected.length - 1)]);
      }
      expect(wb.sheets[wb.activeSheetIndex]?.sheet.title).toBe(active.title);
      const parts = unzipSync(await workbookToBytes(wb));
      const root = parseDocument(strFromU8(required(parts['xl/workbook.xml'])));
      expect(children(required(children(root, 'sheets')[0]), 'sheet').map(n => attribute(n, 'name'))).toEqual(expected.map(s => s.title));
      const names = children(required(children(root, 'definedNames')[0]), 'definedName');
      expect(names.filter(n => attribute(n, 'name') === 'Local').map(n => [Number(attribute(n, 'localSheetId')), textContent(n)]).sort((x, y) => Number(x[0]) - Number(y[0]))).toEqual(expected.map((s, index) => [index, String(s.value)]));
      expect(names.find(n => attribute(n, 'name') === 'Global')?.attributes.some(a2 => a2.local === 'localSheetId')).toBe(false);
      for (const [index, sheet] of expected.entries()) {
        const xml = parseDocument(strFromU8(required(parts[`xl/worksheets/sheet${index + 1}.xml`])));
        const cells = children(required(children(xml, 'sheetData')[0]), 'row').flatMap(row => children(row, 'c'));
        const number = required(cells.find(cell => attribute(cell, 'r') === 'A2'));
        expect(textContent(required(children(number, 'v')[0]))).toBe(String(sheet.value));
        const formula = required(cells.find(cell => attribute(cell, 'r') === 'A3'));
        expect(textContent(required(children(formula, 'f')[0]))).toBe('A2+1');
        expect(textContent(required(children(formula, 'v')[0]))).toBe(String(sheet.value + 1));
      }
      const tables = Object.entries(parts).filter(([path]) => /^xl\/tables\/table\d+\.xml$/.test(path)).map(([, bytes]) => parseDocument(strFromU8(bytes)));
      for (const attr of ['id', 'name', 'displayName']) expect(new Set(tables.map(t => attribute(t, attr))).size).toBe(expected.length);
      expect((await validateXlsx(await workbookToBytes(wb), { skipXsd: true })).issues).toEqual([]);
    }
    expect((await validateXlsx(await workbookToBytes(wb))).issues).toEqual([]);
  }), { seed: 376194, numRuns: 30 });
}, 60_000);
