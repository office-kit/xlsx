import { describe, expect, it } from 'vitest';
import { fromTree, toTree } from '../../src/schema/serialize.js';
import {
  BUILTIN_FORMATS,
  BUILTIN_FORMATS_MAX_SIZE,
  builtinFormatCode,
  builtinFormatId,
  classifyDateFormat,
  FORMAT_DATE_DATETIME,
  FORMAT_DATE_TIMEDELTA,
  FORMAT_DATE_YYYYMMDD2,
  FORMAT_GENERAL,
  FORMAT_TEXT,
  isBuiltinFormat,
  isDateFormat,
  isTimedeltaFormat,
  makeNumberFormat,
} from '../../src/styles/numbers.js';
import { NumberFormatSchema } from '../../src/styles/numbers.schema.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { parseXml } from '../../src/xml/parser.js';
import { serializeXml } from '../../src/xml/serializer.js';

describe('BUILTIN_FORMATS catalogue', () => {
  // Cross-checked against openpyxl/openpyxl/styles/numbers.py: the
  // catalogue has 36 entries (IDs 0–22, 37–49 with gaps).
  it('has the 36 openpyxl built-ins', () => {
    expect(Object.keys(BUILTIN_FORMATS).length).toBe(36);
  });

  it('matches well-known IDs', () => {
    expect(BUILTIN_FORMATS[0]).toBe('General');
    expect(BUILTIN_FORMATS[1]).toBe('0');
    expect(BUILTIN_FORMATS[9]).toBe('0%');
    expect(BUILTIN_FORMATS[14]).toBe('mm-dd-yy');
    expect(BUILTIN_FORMATS[22]).toBe('m/d/yy h:mm');
    expect(BUILTIN_FORMATS[46]).toBe('[h]:mm:ss');
    expect(BUILTIN_FORMATS[49]).toBe('@');
  });

  it('reserves 164 as the user-defined floor', () => {
    expect(BUILTIN_FORMATS_MAX_SIZE).toBe(164);
  });

  it('exposes named ergonomics constants', () => {
    expect(FORMAT_GENERAL).toBe('General');
    expect(FORMAT_TEXT).toBe('@');
    expect(FORMAT_DATE_DATETIME).toBe('yyyy-mm-dd h:mm:ss');
    expect(FORMAT_DATE_TIMEDELTA).toBe('[hh]:mm:ss');
    expect(FORMAT_DATE_YYYYMMDD2).toBe('yyyy-mm-dd');
  });
});

describe('builtinFormatCode / builtinFormatId / isBuiltinFormat', () => {
  it('round-trip lookups through ID and code', () => {
    expect(builtinFormatCode(9)).toBe('0%');
    expect(builtinFormatId('0%')).toBe(9);
    expect(builtinFormatCode(builtinFormatId('mm-dd-yy') ?? -1)).toBe('mm-dd-yy');
  });

  it('unknown lookups return undefined', () => {
    expect(builtinFormatCode(999)).toBeUndefined();
    expect(builtinFormatCode(23)).toBeUndefined(); // hole in the catalogue
    expect(builtinFormatId('not-a-real-format')).toBeUndefined();
  });

  it('isBuiltinFormat checks code membership', () => {
    expect(isBuiltinFormat('General')).toBe(true);
    expect(isBuiltinFormat('@')).toBe(true);
    expect(isBuiltinFormat('yyyy-mm-dd h:mm:ss')).toBe(false);
  });
});

describe('isDateFormat heuristic', () => {
  // openpyxl's reference test cases.
  it.each([
    ['mm-dd-yy', true],
    ['yyyy-mm-dd', true],
    ['h:mm:ss', true],
    ['[h]:mm:ss', true],
    ['0', false],
    ['0.00', false],
    ['"$"#,##0.00', false],
    ['General', false],
    ['#,##0', false],
    ['@', false],
    ['m/d/yy h:mm', true],
    [null, false],
    [undefined, false],
    // openpyxl-style cell with literal text containing date tokens
    // shouldn't be classified as date.
    ['"date" 0', false],
    // With locale tag — still date.
    ['[$-409]m/d/yyyy', true],
    // Japanese era and weekday names.
    ['[$-411]gge', true],
    ['aaaa', true],
    ['general', false],
  ])('isDateFormat(%j) === %s', (code, expected) => {
    expect(isDateFormat(code as string | null | undefined)).toBe(expected);
  });
});

describe('isTimedeltaFormat heuristic', () => {
  it.each([
    ['[h]:mm:ss', true],
    ['[hh]:mm:ss', true],
    ['[mm]:ss', true],
    ['[ss].0', true],
    ['mm:ss', false],
    ['h:mm:ss', false],
    ['0', false],
    [null, false],
  ])('isTimedeltaFormat(%j) === %s', (code, expected) => {
    expect(isTimedeltaFormat(code as string | null | undefined)).toBe(expected);
  });
});

describe('classifyDateFormat', () => {
  it.each([
    ['yyyy-mm-dd', 'date'],
    ['mm-dd-yy', 'date'],
    ['h:mm:ss', 'time'],
    ['mm:ss', 'time'],
    ['[$-411]gge', 'date'],
    ['aaa', 'date'],
    ['h:mm AM/PM', 'time'],
    ['yyyy-mm-dd h:mm:ss', 'datetime'],
    ['m/d/yy h:mm', 'datetime'],
    ['0%', undefined],
    [null, undefined],
  ])('classifyDateFormat(%j) === %s', (code, expected) => {
    expect(classifyDateFormat(code as string | null | undefined)).toBe(expected);
  });
});

describe('NumberFormat value + schema', () => {
  it('makeNumberFormat freezes the result and validates inputs', () => {
    const f = makeNumberFormat({ numFmtId: 200, formatCode: '0.0000' });
    expect(Object.isFrozen(f)).toBe(true);
    expect(f.numFmtId).toBe(200);
    expect(f.formatCode).toBe('0.0000');
  });

  it('rejects negative or non-integer numFmtId', () => {
    expect(() => makeNumberFormat({ numFmtId: -1, formatCode: '0' })).toThrowError(OpenXmlSchemaError);
    expect(() => makeNumberFormat({ numFmtId: 1.5, formatCode: '0' })).toThrowError(OpenXmlSchemaError);
  });

  it('round-trips via the schema', () => {
    const f = makeNumberFormat({ numFmtId: 165, formatCode: 'yyyy-mm-dd' });
    const back = fromTree(parseXml(serializeXml(toTree(f, NumberFormatSchema))), NumberFormatSchema);
    expect(back).toEqual(f);
  });

  it('emits the expected XML attributes', () => {
    const xml = new TextDecoder().decode(
      serializeXml(toTree(makeNumberFormat({ numFmtId: 165, formatCode: '0.0' }), NumberFormatSchema)),
    );
    expect(xml).toContain('<numFmt');
    expect(xml).toContain('numFmtId="165"');
    expect(xml).toContain('formatCode="0.0"');
  });
});
