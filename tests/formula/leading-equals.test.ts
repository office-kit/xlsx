// A leading `=` is how a spreadsheet UI spells a formula and is exactly what
// OOXML forbids in the stored text (ECMA-376 §18.3.1.40). Excel reports a file
// whose `<f>` carries one as unreadable content, so these assertions run on the
// serialised XML as well as on the in-memory value: `FormulaValue` and
// `DataValidation` are public interfaces, so a caller can build one as a
// literal and reach the writer without passing a constructor.

import { describe, expect, it } from 'vitest';
import {
  bindValue,
  type FormulaValue,
  makeArrayFormula,
  makeCell,
  makeDataTableFormula,
  makeFormula,
  makeSharedFormula,
  setArrayFormula,
  setCellValue,
  setDataTableFormula,
  setFormula,
  setSharedFormula,
} from '../../src/cell/cell.js';
import { inferCellType } from '../../src/utils/inference.js';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { makeStylesheet } from '../../src/styles/stylesheet.js';
import { makeSharedStrings } from '../../src/workbook/shared-strings.js';
import { addDefinedName, makeDefinedName } from '../../src/workbook/defined-names.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { makeCfRule, makeConditionalFormatting } from '../../src/worksheet/conditional-formatting.js';
import { makeDataValidation } from '../../src/worksheet/data-validations.js';
import { parseWorksheetXml } from '../../src/worksheet/reader.js';
import { worksheetToBytes } from '../../src/worksheet/writer.js';
import {
  addConditionalFormatting,
  addDataValidation,
  ensureCell,
  getCell,
  makeWorksheet,
  setCell,
  type Worksheet,
} from '../../src/worksheet/worksheet.js';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const sheetXml = (rows: string, afterSheetData = ''): string =>
  `<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheetData>${rows}</sheetData>${afterSheetData}</worksheet>`;

const sheetText = (ws: Worksheet): string =>
  new TextDecoder().decode(worksheetToBytes(ws, { sharedStrings: makeSharedStrings(), styles: makeStylesheet() }));

const formulaAt = (ws: Worksheet, row: number, col: number): FormulaValue => {
  const value = getCell(ws, row, col)?.value;
  if (value === null || typeof value !== 'object' || !('kind' in value) || value.kind !== 'formula') {
    throw new Error(`no formula at row ${row}, col ${col}`);
  }
  return value;
};

describe('formula values normalise a leading =', () => {
  it('strips it in every value constructor', () => {
    expect(makeFormula('=SUM(A1:A3)').formula).toBe('SUM(A1:A3)');
    expect(makeArrayFormula('A1:A3', '=TRANSPOSE(B1:D1)').formula).toBe('TRANSPOSE(B1:D1)');
    expect(makeSharedFormula(0, '=A1*2').formula).toBe('A1*2');
    expect(makeDataTableFormula('=TABLE(B1,C1)', { ref: 'A1:A3' }).formula).toBe('TABLE(B1,C1)');
  });

  it('strips it in every in-place setter', () => {
    const ws = makeWorksheet('Sheet1');

    setFormula(ensureCell(ws, 1, 1), '=A2+B2');
    setArrayFormula(ensureCell(ws, 2, 1), 'A2:A4', '=TRANSPOSE(B1:D1)');
    setSharedFormula(ensureCell(ws, 3, 1), 0, '=A1*2');
    setDataTableFormula(ensureCell(ws, 4, 1), '=TABLE(B1,C1)', { ref: 'A4:A6' });

    expect(formulaAt(ws, 1, 1).formula).toBe('A2+B2');
    expect(formulaAt(ws, 2, 1).formula).toBe('TRANSPOSE(B1:D1)');
    expect(formulaAt(ws, 3, 1).formula).toBe('A1*2');
    expect(formulaAt(ws, 4, 1).formula).toBe('TABLE(B1,C1)');
  });

  it('strips the whitespace a paste leaves around the =', () => {
    expect(makeFormula(' =SUM(A1:A3) ').formula).toBe('SUM(A1:A3)');
    expect(makeFormula('= SUM(A1:A3)').formula).toBe('SUM(A1:A3)');
  });

  it('rejects text that normalises to nothing, rather than emitting <f/>', () => {
    expect(() => makeFormula('=')).toThrow(OpenXmlSchemaError);
    expect(() => makeFormula('')).toThrow(OpenXmlSchemaError);
    expect(() => makeFormula('  ')).toThrow(OpenXmlSchemaError);
    expect(() => makeArrayFormula('A1:A3', '=')).toThrow(OpenXmlSchemaError);
    // A shared-formula reference cell legitimately carries no text of its own,
    // and Excel writes `<f t="dataTable">` with none either.
    expect(makeSharedFormula(1).formula).toBe('');
    expect(makeDataTableFormula('', { ref: 'A1:A3' }).formula).toBe('');
  });

  it('rejects a second = rather than stripping that one too', () => {
    // `'==A1'` is invalid in Excel's formula bar as well, so there is nothing
    // to recover: stripping again would store `A1`, a formula the caller never
    // wrote, and keeping it would store the `=` this whole module removes.
    const cell = makeCell(1, 1);
    expect(() => makeFormula('==A1')).toThrow(OpenXmlSchemaError);
    expect(() => makeArrayFormula('A1:A3', '= =TRANSPOSE(B1:D1)')).toThrow(OpenXmlSchemaError);
    expect(() => makeSharedFormula(0, '==A1*2')).toThrow(OpenXmlSchemaError);
    expect(() => makeDataTableFormula('==TABLE(B1,C1)', { ref: 'A1:A3' })).toThrow(OpenXmlSchemaError);
    expect(() => setFormula(cell, '==A1')).toThrow(OpenXmlSchemaError);
    expect(() => setArrayFormula(cell, 'A1:A3', '==A1')).toThrow(OpenXmlSchemaError);
    expect(() => setSharedFormula(cell, 0, '==A1')).toThrow(OpenXmlSchemaError);
    expect(() => setDataTableFormula(cell, '==TABLE(B1,C1)', { ref: 'A1:A3' })).toThrow(OpenXmlSchemaError);
    // `bindValue` routes any string starting with `=` to the formula path, so
    // it reports the same problem instead of landing `<f>=A1</f>` on the cell.
    expect(() => bindValue(cell, '==A1')).toThrow(OpenXmlSchemaError);
    expect(() => bindValue(cell, '= ')).toThrow(OpenXmlSchemaError);
  });

  it('binds a lone = as text, as Excel stores it', () => {
    const cell = makeCell(1, 1);
    bindValue(cell, '=');
    expect(cell.value).toBe('=');
    expect(inferCellType('=')).toBe('s');
  });

  it('names the call in the rejection', () => {
    expect(() => makeFormula('==A1')).toThrow(/^makeFormula: /);
    expect(() => makeSharedFormula(0, '==A1')).toThrow(/^makeSharedFormula: /);
    expect(() => makeDataTableFormula('==TABLE(B1,C1)', { ref: 'A1:A3' })).toThrow(/^makeDataTableFormula: /);
  });

  it('bounds the text it quotes back, which is a whole formula', () => {
    const long = `==${'A'.repeat(500)}`;
    let message = '';
    try {
      makeFormula(long);
    } catch (e) {
      message = (e as Error).message;
    }
    expect(message).toContain('(502 chars)');
    expect(message).not.toContain('A'.repeat(50));
  });

  it('freezes the value it returns', () => {
    expect(Object.isFrozen(makeFormula('A1'))).toBe(true);
    expect(Object.isFrozen(makeArrayFormula('A1:A2', 'A1'))).toBe(true);
    expect(Object.isFrozen(makeSharedFormula(0, 'A1'))).toBe(true);
    expect(Object.isFrozen(makeDataTableFormula('TABLE(B1,)', { ref: 'A1:A3' }))).toBe(true);
  });
});

describe('<f> never carries a leading =', () => {
  it('drops it from a hand-built FormulaValue that skipped the constructors', () => {
    const ws = makeWorksheet('Sheet1');
    setCell(ws, 1, 1, { kind: 'formula', t: 'normal', formula: '=SUM(B1:B3)' });
    setCellValue(ensureCell(ws, 2, 1), { kind: 'formula', t: 'array', formula: '=ROW(A1:A2)', ref: 'A2:A3' });

    const xml = sheetText(ws);
    expect(xml).toContain('<f>SUM(B1:B3)</f>');
    expect(xml).toContain('<f t="array" ref="A2:A3">ROW(A1:A2)</f>');
    expect(xml).not.toContain('<f>=');
  });

  it.each(['normal', 'array'] as const)('rejects empty hand-built %s formulas on save', (t) => {
    for (const formula of ['', '  ', '=', ' = ']) {
      const ws = makeWorksheet('Sheet1');
      setCell(ws, 1, 1, { kind: 'formula', t, formula, ...(t === 'array' ? { ref: 'A1:A2' } : {}) });
      expect(() => sheetText(ws)).toThrow(OpenXmlSchemaError);
    }
  });

  it('preserves empty shared references and data-table formulas on save', () => {
    const ws = makeWorksheet('Sheet1');
    setCell(ws, 1, 1, makeSharedFormula(0));
    setCell(ws, 2, 1, makeDataTableFormula('', { ref: 'A2:A3' }));
    expect(sheetText(ws)).toContain('<f t="shared" si="0"/>');
    expect(sheetText(ws)).toContain('<f t="dataTable" ref="A2:A3"/>');
  });

  it('rejects a hand-built value carrying a second =, naming the cell', () => {
    const ws = makeWorksheet('Sheet1');
    setCell(ws, 7, 2, { kind: 'formula', t: 'normal', formula: '==A1' });
    expect(() => sheetText(ws)).toThrow(/<f> at B7/);
  });

  it('drops it on the way back out of a file that had one', () => {
    const ws = parseWorksheetXml(sheetXml('<row r="1"><c r="A1"><f>=SUM(B1:B2)</f></c></row>'), 'Sheet1', {
      sharedStrings: [],
    });
    expect(formulaAt(ws, 1, 1).formula).toBe('SUM(B1:B2)');
    expect(sheetText(ws)).toContain('<f>SUM(B1:B2)</f>');
  });

  it('repairs a file whose <f> carries a second =, rather than failing the load', () => {
    // One malformed `<f>` must not cost the whole workbook: the file is
    // already outside the spec, and every reading of the expression is the
    // same once the prefix is gone.
    const ws = parseWorksheetXml(sheetXml('<row r="1"><c r="A1"><f>==SUM(B1:B2)</f></c></row>'), 'Sheet1', {
      sharedStrings: [],
    });
    expect(formulaAt(ws, 1, 1).formula).toBe('SUM(B1:B2)');
    expect(sheetText(ws)).toContain('<f>SUM(B1:B2)</f>');
  });

  it('rejects a file whose <f> holds nothing but the =', () => {
    expect(() =>
      parseWorksheetXml(sheetXml('<row r="1"><c r="A1"><f>=</f></c></row>'), 'Sheet1', { sharedStrings: [] }),
    ).toThrow(OpenXmlSchemaError);
  });

  it('shifts the references of a shared formula whose origin text had one', () => {
    const ws = parseWorksheetXml(
      sheetXml(
        '<row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">=A1*2</f></c></row>' +
          '<row r="3"><c r="A3"><f t="shared" si="0"/></c></row>',
      ),
      'Sheet1',
      { sharedStrings: [] },
    );
    expect(formulaAt(ws, 2, 1).formula).toBe('A1*2');
    expect(formulaAt(ws, 3, 1).formula).toBe('A2*2');
    expect(sheetText(ws)).not.toContain('si="0">=');
  });
});

describe('the other elements that carry OOXML formula text', () => {
  it('normalises dataValidation formulas, constructed and hand-built', () => {
    const dv = makeDataValidation({ type: 'custom', sqref: 'A1', formula1: '=$A$1>0' });
    expect(dv.formula1).toBe('$A$1>0');

    const ws = makeWorksheet('Sheet1');
    addDataValidation(ws, dv);
    addDataValidation(ws, { ...dv, formula1: '=$A$2>0', formula2: '=$A$3' });

    const xml = sheetText(ws);
    expect(xml).toContain('<formula1>$A$1&gt;0</formula1>');
    expect(xml).toContain('<formula1>$A$2&gt;0</formula1><formula2>$A$3</formula2>');
    expect(xml).not.toContain('<formula1>=');
  });

  it('normalises conditional-formatting formulas, constructed and hand-built', () => {
    const rule = makeCfRule({ type: 'expression', priority: 1, formulas: ['=$A$1>0'] });
    expect(rule.formulas[0]).toBe('$A$1>0');

    const ws = makeWorksheet('Sheet1');
    addConditionalFormatting(ws, makeConditionalFormatting({ sqref: 'A1:A5', rules: [rule] }));
    addConditionalFormatting(
      ws,
      makeConditionalFormatting({
        sqref: 'B1:B5',
        rules: [{ ...rule, priority: 2, formulas: ['=$B$1>0'] }],
      }),
    );

    const xml = sheetText(ws);
    expect(xml).toContain('<formula>$A$1&gt;0</formula>');
    expect(xml).toContain('<formula>$B$1&gt;0</formula>');
    expect(xml).not.toContain('<formula>=');
  });

  it('normalises defined-name values, constructed and hand-built', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    expect(addDefinedName(wb, { name: 'Constructed', value: '=Sheet1!$A$1' }).value).toBe('Sheet1!$A$1');
    wb.definedNames.push({ name: 'HandBuilt', value: '=Sheet1!$A$2' });

    const { unzipSync } = await import('fflate');
    const xml = new TextDecoder().decode(unzipSync(await workbookToBytes(wb))['xl/workbook.xml']);
    expect(xml).toContain('<definedName name="Constructed">Sheet1!$A$1</definedName>');
    expect(xml).toContain('<definedName name="HandBuilt">Sheet1!$A$2</definedName>');
  });
});

describe('a second = in the other elements that carry formula text', () => {
  it('is rejected by each constructor, naming where the text came from', () => {
    expect(() => makeDefinedName({ name: 'Doubled', value: '==Sheet1!$A$1' })).toThrow(
      /^makeDefinedName "Doubled": /,
    );
    expect(() => makeCfRule({ type: 'expression', priority: 3, formulas: ['==$A$1>0'] })).toThrow(
      /^makeCfRule at priority 3: /,
    );
    expect(() => makeDataValidation({ type: 'custom', sqref: 'A1', formula1: '==$A$1>0' })).toThrow(
      /^makeDataValidation formula1: /,
    );
    expect(() =>
      makeDataValidation({ type: 'whole', operator: 'between', sqref: 'A1', formula1: '$A$2', formula2: '==$A$3' }),
    ).toThrow(/^makeDataValidation formula2: /);
  });

  it('is rejected on save when a hand-built value carries it, naming the element', async () => {
    const dvSheet = makeWorksheet('Sheet1');
    const dv = makeDataValidation({ type: 'custom', sqref: 'A1', formula1: '$A$1>0' });
    addDataValidation(dvSheet, { ...dv, formula1: '==$A$1>0' });
    expect(() => sheetText(dvSheet)).toThrow(/<formula1> at A1/);

    const cfSheet = makeWorksheet('Sheet1');
    const rule = makeCfRule({ type: 'expression', priority: 4, formulas: [] });
    addConditionalFormatting(
      cfSheet,
      makeConditionalFormatting({ sqref: 'A1:A5', rules: [{ ...rule, formulas: ['==$A$1>0'] }] }),
    );
    expect(() => sheetText(cfSheet)).toThrow(/<formula> at priority 4/);

    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    wb.definedNames.push({ name: 'HandBuilt', value: '==Sheet1!$A$2' });
    await expect(workbookToBytes(wb)).rejects.toThrow(/<definedName> "HandBuilt"/);
  });

  it('is repaired, not refused, when it comes out of a file', () => {
    const ws = parseWorksheetXml(
      sheetXml(
        '',
        '<conditionalFormatting sqref="A1:A5"><cfRule type="expression" priority="1">' +
          '<formula>==$A$1&gt;0</formula></cfRule></conditionalFormatting>' +
          '<dataValidations count="1"><dataValidation type="custom" sqref="A1">' +
          '<formula1>==$A$1&gt;0</formula1></dataValidation></dataValidations>',
      ),
      'Sheet1',
      { sharedStrings: [] },
    );
    expect(ws.conditionalFormatting[0]?.rules[0]?.formulas[0]).toBe('$A$1>0');
    expect(ws.dataValidations[0]?.formula1).toBe('$A$1>0');

    const xml = sheetText(ws);
    expect(xml).toContain('<formula>$A$1&gt;0</formula>');
    expect(xml).toContain('<formula1>$A$1&gt;0</formula1>');
  });

  it('is repaired in a <definedName> a producer wrote wrong', async () => {
    const { unzipSync, zipSync } = await import('fflate');
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    wb.definedNames.push(makeDefinedName({ name: 'Total', value: 'Sheet1!$A$1' }));

    const archive = unzipSync(await workbookToBytes(wb));
    const part = 'xl/workbook.xml';
    const entry = archive[part];
    if (!entry) throw new Error(`no ${part} in the package`);
    const patched = new TextDecoder().decode(entry).replace('>Sheet1!$A$1<', () => '>==Sheet1!$A$1<');
    expect(patched).toContain('>==Sheet1!$A$1</definedName>');
    archive[part] = new TextEncoder().encode(patched);

    const reloaded = await loadWorkbook(fromBuffer(zipSync(archive)));
    expect(reloaded.definedNames[0]?.value).toBe('Sheet1!$A$1');
  });
});

describe('a formula written with a leading = survives a real save and load', () => {
  it('reads back the text Excel stores', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    setCell(ws, 1, 1, 12);
    setCell(ws, 2, 1, makeFormula('=A1*2', { cachedValue: 24 }));

    const reloaded = await loadWorkbook(fromBuffer(await workbookToBytes(wb)));
    const sheet = reloaded.sheets[0];
    if (sheet?.kind !== 'worksheet') throw new Error('expected a worksheet');
    expect(formulaAt(sheet.sheet, 2, 1).formula).toBe('A1*2');
  });
});
