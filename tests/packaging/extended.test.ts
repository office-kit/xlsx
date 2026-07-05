import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { extendedPropsFromBytes, extendedPropsToBytes, makeExtendedProperties } from '../../src/packaging/extended';
import { openZip } from '../../src/zip/reader';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');

describe('extendedProperties — basic', () => {
  it('makeExtendedProperties yields an empty object', () => {
    expect(makeExtendedProperties()).toEqual({});
  });

  it('round-trips simple flat fields', () => {
    const p = {
      application: '@office-kit/xlsx',
      appVersion: '0.0.0',
      docSecurity: 0,
      scaleCrop: false,
      company: 'Acme',
      linksUpToDate: false,
      sharedDoc: false,
      hyperlinksChanged: false,
    };
    const back = extendedPropsFromBytes(extendedPropsToBytes(p));
    expect(back).toEqual(p);
  });

  it('emits booleans as 1/0 but reads back true/false (loose accept on parse)', () => {
    const xml = new TextDecoder().decode(extendedPropsToBytes({ scaleCrop: true }));
    expect(xml).toContain('<ScaleCrop>1</ScaleCrop>');
    // openpyxl's output uses 'false'/'true'; we still parse it loosely.
    const looseInput = new TextEncoder().encode(
      '<?xml version="1.0" encoding="UTF-8"?>' +
        `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">` +
        '<ScaleCrop>false</ScaleCrop></Properties>',
    );
    expect(extendedPropsFromBytes(looseInput)).toEqual({ scaleCrop: false });
  });
});

describe('extendedProperties — openpyxl genuine/empty.xlsx round-trip', () => {
  it('parses the simple flat fields out of docProps/app.xml', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const p = extendedPropsFromBytes(zip.read('docProps/app.xml'));

    // Cross-checked against the live XML inspected this turn.
    expect(p.application).toBe('Microsoft Excel');
    expect(p.docSecurity).toBe(0);
    expect(p.scaleCrop).toBe(false);
    expect(p.company).toBe('IT-Services');
    expect(p.linksUpToDate).toBe(false);
    expect(p.sharedDoc).toBe(false);
    expect(p.hyperlinksChanged).toBe(false);
    expect(p.appVersion).toBe('12.0000');
  });

  it('preserves HeadingPairs / TitlesOfParts as raw subtrees', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const p = extendedPropsFromBytes(zip.read('docProps/app.xml'));
    expect(p.headingPairs).toBeDefined();
    expect(p.titlesOfParts).toBeDefined();
    // Sanity: TitlesOfParts wraps a vt:vector with size="3".
    const vector = p.titlesOfParts?.children[0];
    expect(vector?.attrs['size']).toBe('3');
  });

  it('parsing → re-serialising → re-parsing yields equal objects', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const original = zip.read('docProps/app.xml');
    const a = extendedPropsFromBytes(original);
    const re = extendedPropsToBytes(a);
    const b = extendedPropsFromBytes(re);
    expect(b).toEqual(a);
  });
});
