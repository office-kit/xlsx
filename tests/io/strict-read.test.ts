import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { getCell } from '../../src/worksheet/worksheet.js';
import { validateXlsx } from '../conformance/validate.js';

for (const file of ['sample.strict.xlsx', 'SimpleStrict.xlsx']) {
  it(`reads genuine ${file} and saves schema-valid Transitional XML`, async () => {
    const bytes = readFileSync(new URL(`./fixtures/strict/${file}`, import.meta.url));
    const wb = await loadWorkbook(fromBuffer(bytes));
    expect(wb.sheets.length).toBeGreaterThan(0);
    if (file === 'sample.strict.xlsx') {
      const first = wb.sheets[0];
      if (first?.kind !== 'worksheet') throw new Error('expected worksheet');
      expect(getCell(first.sheet, 1, 1)?.value).toBe('Lorem');
      expect(getCell(first.sheet, 10, 2)?.value).toMatchObject({ kind: 'formula', formula: 'SUM(B1:B9)', cachedValue: 4995 });
    }
    if (file === 'SimpleStrict.xlsx') {
      const sheet = wb.sheets[1];
      expect(sheet?.kind).toBe('worksheet');
      if (sheet?.kind === 'worksheet') expect(getCell(sheet.sheet, 4, 1)?.value).toBe(32874);
    }
    const stream = await loadWorkbookStream(fromBuffer(bytes));
    try {
      expect(stream.sheetNames).toEqual(wb.sheets.map((s) => s.sheet.title));
      for (const name of stream.sheetNames) {
        for await (const row of stream.openWorksheet(name).iterRows()) {
          const ref = wb.sheets.find((s) => s.sheet.title === name);
          if (ref?.kind !== 'worksheet') throw new Error('expected worksheet');
          for (const cell of row) {
            const value = getCell(ref.sheet, cell.row, cell.col)?.value ?? null;
            const expected = typeof value === 'object' && value !== null && !(value instanceof Date)
              ? value.kind === 'formula' ? value.cachedValue ?? null
                : value.kind === 'rich-text' ? value.runs.map((r) => r.text).join('') : value
              : value;
            expect(cell.value).toEqual(expected);
          }
        }
      }
    } finally { await stream.close(); }
    const saved = await workbookToBytes(wb);
    const reloaded = await loadWorkbook(fromBuffer(saved));
    expect(reloaded.styles).toEqual(wb.styles);
    for (const [path, data] of wb.passthrough ?? []) {
      if (path.endsWith('.bin')) expect(reloaded.passthrough?.get(path)).toEqual(data);
    }
    for (let i = 0; i < wb.sheets.length; i++) {
      const before = wb.sheets[i];
      const after = reloaded.sheets[i];
      if (before?.kind === 'worksheet' && after?.kind === 'worksheet') {
        for (const [row, cells] of before.sheet.rows) {
          for (const [col, cell] of cells) expect(getCell(after.sheet, row, col)?.value).toEqual(cell.value);
        }
      }
    }
    const result = await validateXlsx(saved);
    expect(result.issues).toEqual([]);
  });
}
