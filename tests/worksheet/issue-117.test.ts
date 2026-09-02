// Regression for https://github.com/office-kit/xlsx/issues/117 — `setHyperlink`
// accepted any string as `target` and the writer emitted it verbatim into the
// worksheet rels. A target that is not a legal `xsd:anyURI` produced a package
// Excel refuses ("We found a problem with some content in …") and drops
// content on repair, with no error or warning at save time.
//
// The reporter hit four variants in one real export — a trailing space, a
// space after the scheme, and two umlaut hostnames — where one bad row poisons
// an export of thousands.

import { describe, expect, it } from 'vitest';
import { OpenXmlSchemaError } from '../../src/utils/exceptions';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { addUrlHyperlink, makeHyperlink } from '../../src/worksheet/hyperlinks';
import { parseWorksheetXml } from '../../src/worksheet/reader';
import { setCellByCoord, setHyperlink } from '../../src/worksheet/index';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const sheet = () => addWorksheet(createWorkbook(), 'Tabelle1');

describe('issue #117 — an invalid hyperlink target is rejected where it can be fixed', () => {
  it.each([
    ['a space after the scheme', 'https:// www.example.com'],
    ['a trailing space', 'https://www.example.com '],
    ['a tab', 'https://www.example.com\tx'],
    ['a newline', 'https://www.example.com\nx'],
  ])('rejects %s', (_label, target) => {
    const ws = sheet();
    setCellByCoord(ws, 'A1', 'Firma');
    expect(() => setHyperlink(ws, 'A1', { target })).toThrowError(OpenXmlSchemaError);
    expect(ws.hyperlinks).toEqual([]);
  });

  it('rejects a non-ASCII host and points at punycode', () => {
    const ws = sheet();
    expect(() => setHyperlink(ws, 'A1', { target: 'https://www.übleis.at' })).toThrowError(/punycode/);
  });

  it('rejects an empty target', () => {
    const ws = sheet();
    expect(() => setHyperlink(ws, 'A1', { target: '' })).toThrowError(/target is empty/);
  });

  it('accepts the punycode form of the same host', () => {
    const ws = sheet();
    const hl = setHyperlink(ws, 'A1', { target: 'https://www.xn--bleis-jva.at' });
    expect(hl.target).toBe('https://www.xn--bleis-jva.at');
  });

  it('accepts non-ASCII outside the host, and percent-encoded paths', () => {
    const ws = sheet();
    expect(() => addUrlHyperlink(ws, 'A1', 'https://example.com/Größe.pdf')).not.toThrow();
    expect(() => addUrlHyperlink(ws, 'A2', '../docs/report%20v2.pdf')).not.toThrow();
  });

  it('still accepts an in-workbook jump with no target at all', () => {
    const ws = sheet();
    const hl = setHyperlink(ws, 'A1', { location: "'Tabelle 2'!A1" });
    expect(hl.location).toBe("'Tabelle 2'!A1");
  });

  it('rejects the same targets through makeHyperlink', () => {
    expect(() => makeHyperlink({ ref: 'A1', target: 'https://www.example.com ' })).toThrowError(
      OpenXmlSchemaError,
    );
  });

  it('still loads a file whose rels already carry a broken target', async () => {
    // The check is on authoring, not on reading: a package that already has
    // the bad link has to open so it can be inspected and repaired.
    const ws = parseWorksheetXml(
      `<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheetData/>` +
        '<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks></worksheet>',
      'Tabelle1',
      {
        sharedStrings: [],
        rels: {
          rels: [
            {
              id: 'rId1',
              type: `${REL_NS}/hyperlink`,
              target: 'https:// www.example.com',
              targetMode: 'External',
            },
          ],
        },
      },
    );
    expect(ws.hyperlinks[0]?.target).toBe('https:// www.example.com');
  });
});
