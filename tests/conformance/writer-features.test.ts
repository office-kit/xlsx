// Writer feature survey. Exercises every major user-facing writer API
// once and pipes the result through validateXlsx, so any feature whose
// output drifts off-schema fails CI rather than waiting for an end-user
// to discover it in Excel.
//
// Each `it` is intentionally narrow: when something breaks, the failing
// scenario name should point straight at the API responsible.

import { describe, expect, it } from 'vitest';
import { setFormula } from '../../src/cell/cell.js';
import { workbookToBytes } from '../../src/io/save.js';
import { makeColor } from '../../src/styles/colors.js';
import { setCellBackgroundColor, setCellFont } from '../../src/styles/cell-style.js';
import { makeFont } from '../../src/styles/fonts.js';
import { addDefinedName } from '../../src/workbook/defined-names.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { makeAutoFilter } from '../../src/worksheet/auto-filter.js';
import { parseMultiCellRange } from '../../src/worksheet/cell-range.js';
import {
  addCellIsRule,
  addColorScaleRule,
  addFormulaRule,
} from '../../src/worksheet/conditional-formatting.js';
import { makeDataValidation } from '../../src/worksheet/data-validations.js';
import {
  addInternalHyperlink,
  addUrlHyperlink,
} from '../../src/worksheet/hyperlinks.js';
import {
  makeHeaderFooter,
  makePageMargins,
  makePageSetup,
  makePrintOptions,
} from '../../src/worksheet/page-setup.js';
import { makeSheetProtection } from '../../src/worksheet/protection.js';
import { addExcelTable } from '../../src/worksheet/table.js';
import {
  addDataValidation,
  appendRow,
  appendRows,
  setFreezePanes,
  freezeRows,
  groupColumns,
  groupRows,
  hideColumn,
  hideRow,
  mergeCells,
  setAutoFilter,
  ensureCell,
  setCell,
  setColumnWidth,
  setComment,
  setDefaultRowHeight,
  setRowHeight,
  setSheetTabColor,
  type Worksheet,
} from '../../src/worksheet/worksheet.js';
import { validateXlsx } from './validate.js';

const dump = (issues: { tier: string; part: string; message: string }[]): string =>
  issues.map((i) => `[${i.tier}] ${i.part}: ${i.message}`).join('\n');

const expectClean = async (wb: ReturnType<typeof createWorkbook>): Promise<void> => {
  const bytes = await workbookToBytes(wb);
  const result = await validateXlsx(bytes);
  expect(result.issues, dump(result.issues)).toEqual([]);
};

const ws = (s: ReturnType<typeof addWorksheet>): Worksheet => s as Worksheet;

describe('conformance: writer feature survey', () => {
  describe('cell values', () => {
    it('numeric, string, boolean, date, and inline rich text', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'Vals'));
      setCell(w, 1, 1, 42);
      setCell(w, 2, 1, 'plain');
      setCell(w, 3, 1, true);
      setCell(w, 4, 1, new Date(Date.UTC(2024, 0, 15)));
      setCell(w, 5, 1, 0.5);
      setCell(w, 6, 1, -1.234e-10);
      setCell(w, 7, 1, '');
      await expectClean(wb);
    });

    it('formulas (normal + cached value)', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'F'));
      setCell(w, 1, 1, 10);
      setCell(w, 1, 2, 20);
      const f = ensureCell(w, 1, 3);
      setFormula(f, 'A1+B1', { cachedValue: 30 });
      const f2 = ensureCell(w, 2, 3);
      setFormula(f2, 'SUM(A1:B1)', { cachedValue: 30 });
      await expectClean(wb);
    });

    it('appendRow and appendRows', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'A'));
      appendRow(w, ['name', 'qty', 'price']);
      appendRows(w, [
        ['apple', 3, 1.5],
        ['banana', 5, 0.5],
        ['cherry', 12, 0.1],
      ]);
      await expectClean(wb);
    });
  });

  describe('styles', () => {
    it('font + background fill', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'S'));
      const c = setCell(w, 1, 1, 'styled');
      setCellFont(wb, c, makeFont({ bold: true, italic: true, size: 14, name: 'Arial' }));
      const c2 = setCell(w, 2, 1, 'filled');
      setCellBackgroundColor(wb, c2, makeColor({ rgb: 'FFFFFF00' }));
      await expectClean(wb);
    });
  });

  describe('sheet structure', () => {
    it('merge, freeze, group, dimensions, hide', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'St'));
      setCell(w, 1, 1, 'header');
      mergeCells(w, 'A1:D1');
      freezeRows(w, 1);
      setRowHeight(w, 2, 24);
      setColumnWidth(w, 1, 18);
      setDefaultRowHeight(w, 15);
      hideRow(w, 5);
      hideColumn(w, 5);
      groupRows(w, 6, 8);
      groupColumns(w, 6, 8);
      setSheetTabColor(w, '7F00FF');
      await expectClean(wb);
    });

    it('setFreezePanes both axes', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'F'));
      setCell(w, 1, 1, 'tl');
      setFreezePanes(w, { rows: 1, cols: 1 });
      await expectClean(wb);
    });
  });

  describe('defined names', () => {
    it('workbook-scope and sheet-scope', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'N'));
      setCell(w, 1, 1, 1);
      addDefinedName(wb, { name: 'Pi', value: '3.14159265' });
      addDefinedName(wb, { name: 'LocalRange', value: "'N'!$A$1:$A$10", scope: 0 });
      await expectClean(wb);
    });
  });

  describe('tables and filtering', () => {
    it('Excel table over a typed range', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'T'));
      appendRow(w, ['name', 'qty']);
      appendRow(w, ['apple', 3]);
      appendRow(w, ['banana', 5]);
      addExcelTable(wb, w, {
        name: 'Fruit',
        ref: 'A1:B3',
        columns: ['name', 'qty'],
        headerRowCount: 1,
      });
      await expectClean(wb);
    });

    it('autoFilter range', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'F'));
      appendRow(w, ['col1', 'col2']);
      appendRow(w, ['a', 1]);
      appendRow(w, ['b', 2]);
      setAutoFilter(w, makeAutoFilter({ ref: 'A1:B3' }));
      await expectClean(wb);
    });
  });

  describe('data validation', () => {
    it('list-type dropdown', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'D'));
      addDataValidation(
        w,
        makeDataValidation({
          type: 'list',
          sqref: parseMultiCellRange('A1:A10'),
          formula1: '"red,green,blue"',
          allowBlank: true,
          showInputMessage: true,
          showErrorMessage: true,
        }),
      );
      await expectClean(wb);
    });

    it('whole-number range', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'D'));
      addDataValidation(
        w,
        makeDataValidation({
          type: 'whole',
          sqref: parseMultiCellRange('B1:B5'),
          formula1: '0',
          formula2: '100',
          operator: 'between',
        }),
      );
      await expectClean(wb);
    });
  });

  describe('conditional formatting', () => {
    it('cellIs, formula, color scale rules combined', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'CF'));
      for (let r = 1; r <= 10; r++) setCell(w, r, 1, r * 5);
      addCellIsRule(w, 'A1:A10', { operator: 'greaterThan', formula1: '20', dxfId: 0 });
      addFormulaRule(w, 'A1:A10', { formula: '=ISEVEN(A1)', dxfId: 0 });
      addColorScaleRule(w, 'A1:A10', {
        cfvos: [
          { type: 'min' },
          { type: 'percentile', val: '50' },
          { type: 'max' },
        ],
        colors: ['FFFF0000', 'FFFFFF00', 'FF00FF00'],
      });
      await expectClean(wb);
    });
  });

  describe('page setup', () => {
    it('orientation, margins, header/footer, print options', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'P'));
      setCell(w, 1, 1, 'page-setup');
      w.pageSetup = makePageSetup({ orientation: 'landscape', paperSize: 9, fitToWidth: 1, fitToHeight: 1 });
      w.pageMargins = makePageMargins({ left: 0.75, right: 0.75, top: 1, bottom: 1, header: 0.3, footer: 0.3 });
      w.headerFooter = makeHeaderFooter({
        oddHeader: '&L&"Arial"&Bxlsx-kit',
        oddFooter: '&CPage &P of &N',
      });
      w.printOptions = makePrintOptions({ horizontalCentered: true, gridLines: true });
      await expectClean(wb);
    });
  });

  describe('protection', () => {
    it('sheet protection with locked structure', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'P'));
      setCell(w, 1, 1, 'locked');
      w.sheetProtection = makeSheetProtection({ sheet: true, formatColumns: true, formatRows: true });
      await expectClean(wb);
    });
  });

  describe('hyperlinks and comments', () => {
    it('URL + internal hyperlink + cell comment', async () => {
      const wb = createWorkbook();
      const w = ws(addWorksheet(wb, 'H'));
      const w2 = ws(addWorksheet(wb, 'Other'));
      setCell(w2, 1, 1, 'jump target');
      setCell(w, 1, 1, 'click me');
      addUrlHyperlink(w, 'A1', 'https://example.com', { tooltip: 'Open' });
      setCell(w, 2, 1, 'jump');
      addInternalHyperlink(w, 'A2', 'Other!A1', { display: 'Go' });
      setCell(w, 3, 1, 'noted');
      setComment(w, { ref: 'A3', author: 'tester', text: 'a note' });
      await expectClean(wb);
    });
  });

  describe('combination', () => {
    it('multi-sheet workbook with mixed features in each', async () => {
      const wb = createWorkbook();
      const a = ws(addWorksheet(wb, 'Data'));
      appendRow(a, ['region', 'sales']);
      appendRow(a, ['EMEA', 100]);
      appendRow(a, ['APAC', 250]);
      appendRow(a, ['AMER', 175]);
      setAutoFilter(a, makeAutoFilter({ ref: 'A1:B4' }));

      const b = ws(addWorksheet(wb, 'Calc'));
      setCell(b, 1, 1, 'total');
      const tot = ensureCell(b, 1, 2);
      setFormula(tot, 'SUM(Data!B2:B4)', { cachedValue: 525 });

      const c = ws(addWorksheet(wb, 'Report'));
      const header = setCell(c, 1, 1, 'styled header');
      setCellFont(wb, header, makeFont({ bold: true, color: makeColor({ rgb: 'FFFFFFFF' }) }));
      setCellBackgroundColor(wb, header, makeColor({ rgb: 'FF003366' }));
      mergeCells(c, 'A1:E1');
      freezeRows(c, 1);

      addDefinedName(wb, { name: 'TotalCell', value: "'Calc'!$B$1" });

      await expectClean(wb);
    });
  });
});
