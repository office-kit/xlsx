// Phase 7 acceptance: VBA / pivot / activeX / customXml passthrough. The goal
// is "openpyxl が壊さない xlsx は @office-kit/xlsx も壊さない" — these tests pin that contract.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import {
  addWorksheet,
  createWorkbook,
  listCustomXmlParts,
} from '../../src/workbook/workbook';
import { OpenXmlNotImplementedError } from '../../src/utils/exceptions';
import { openZip } from '../../src/zip/reader';

const td = new TextDecoder();

describe('Encrypted xlsx detection', () => {
  it('throws OpenXmlNotImplementedError for a CFB compound document', async () => {
    // Synthesise the OLE/CFB magic bytes — Excel's encrypted-document wrapper
    // starts with this 8-byte signature.
    const cfb = new Uint8Array(512);
    cfb.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0);
    await expect(openZip(fromBuffer(cfb))).rejects.toThrowError(OpenXmlNotImplementedError);
    await expect(openZip(fromBuffer(cfb))).rejects.toThrowError(/Encrypted xlsx/i);
  });
});

describe('VBA project round-trip', () => {
  it('preserves vbaProject.bin byte-identical and promotes to xlsm content type', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    // Synthetic vbaProject bytes — the real file would be a CFB compound
    // document; what we care about is byte preservation.
    wb.vbaProject = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00, 0x01, 0x02]);

    const bytes = await workbookToBytes(wb);
    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);

    expect(entries['xl/vbaProject.bin']).toBeDefined();
    expect(entries['xl/vbaProject.bin']).toEqual(wb.vbaProject);

    const ct = td.decode(entries['[Content_Types].xml']);
    expect(ct).toContain('macroEnabled.main+xml');
    expect(ct).toContain('Extension="bin"');
    expect(ct).toContain('vnd.ms-office.vbaProject');

    const wbRels = td.decode(entries['xl/_rels/workbook.xml.rels']);
    expect(wbRels).toContain('relationships/vbaProject');
    expect(wbRels).toContain('vbaProject.bin');

    // Round-trip via load.
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.vbaProject).toEqual(wb.vbaProject);
  });

  it('also preserves vbaProjectSignature.bin when present', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    wb.vbaProject = new Uint8Array([1, 2, 3, 4]);
    wb.vbaSignature = new Uint8Array([5, 6, 7, 8]);
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.vbaSignature).toEqual(wb.vbaSignature);
  });
});

describe('Custom XML pass-through', () => {
  it('round-trips a customXml/itemN.xml entry verbatim with its content type', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    wb.passthrough = new Map();
    wb.passthrough.set(
      'customXml/item1.xml',
      new TextEncoder().encode('<?xml version="1.0"?><root><field>value</field></root>'),
    );
    wb.passthroughContentTypes = new Map();
    wb.passthroughContentTypes.set('customXml/item1.xml', 'application/xml');

    const bytes = await workbookToBytes(wb);
    const { unzipSync } = await import('fflate');
    const entries = unzipSync(bytes);
    expect(entries['customXml/item1.xml']).toEqual(wb.passthrough.get('customXml/item1.xml'));

    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const parts = listCustomXmlParts(wb2);
    expect(parts.length).toBe(1);
    expect(parts[0]?.path).toBe('customXml/item1.xml');
    expect(parts[0]?.content).toEqual(wb.passthrough.get('customXml/item1.xml'));
  });
});

describe('ActiveX + ctrlProps + embeddings pass-through', () => {
  it('round-trips xl/activeX, xl/ctrlProps, xl/embeddings entries verbatim', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    wb.passthrough = new Map();
    wb.passthroughContentTypes = new Map();

    const cases: Array<{ path: string; ct: string }> = [
      { path: 'xl/activeX/activeX1.xml', ct: 'application/vnd.ms-office.activeX+xml' },
      { path: 'xl/ctrlProps/ctrlProp1.xml', ct: 'application/vnd.ms-excel.controlproperties+xml' },
      { path: 'xl/embeddings/oleObject1.bin', ct: 'application/vnd.openxmlformats-officedocument.oleObject' },
      { path: 'customUI/customUI.xml', ct: 'application/xml' },
    ];
    for (let i = 0; i < cases.length; i++) {
      const c = cases[i];
      if (!c) continue;
      wb.passthrough.set(c.path, new Uint8Array([0xff, i]));
      wb.passthroughContentTypes.set(c.path, c.ct);
    }

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));

    expect(wb2.passthrough?.size).toBe(cases.length);
    for (const c of cases) {
      expect(wb2.passthrough?.get(c.path)).toEqual(wb.passthrough.get(c.path));
      expect(wb2.passthroughContentTypes?.get(c.path)).toBe(c.ct);
    }
  });
});

describe('Pivot table pass-through', () => {
  it('round-trips xl/pivotCache and xl/pivotTables entries as bytes', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    wb.passthrough = new Map();
    wb.passthroughContentTypes = new Map();
    const defXml = new TextEncoder().encode(
      '<?xml version="1.0"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="0"/>',
    );
    const recsXml = new TextEncoder().encode(
      '<?xml version="1.0"?><pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>',
    );
    const pivotXml = new TextEncoder().encode(
      '<?xml version="1.0"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P1"/>',
    );
    wb.passthrough.set('xl/pivotCache/pivotCacheDefinition1.xml', defXml);
    wb.passthrough.set('xl/pivotCache/pivotCacheRecords1.xml', recsXml);
    wb.passthrough.set('xl/pivotTables/pivotTable1.xml', pivotXml);
    wb.passthroughContentTypes.set(
      'xl/pivotCache/pivotCacheDefinition1.xml',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
    );
    wb.passthroughContentTypes.set(
      'xl/pivotCache/pivotCacheRecords1.xml',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
    );
    wb.passthroughContentTypes.set(
      'xl/pivotTables/pivotTable1.xml',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
    );

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));

    expect(wb2.passthrough?.get('xl/pivotCache/pivotCacheDefinition1.xml')).toEqual(defXml);
    expect(wb2.passthrough?.get('xl/pivotCache/pivotCacheRecords1.xml')).toEqual(recsXml);
    expect(wb2.passthrough?.get('xl/pivotTables/pivotTable1.xml')).toEqual(pivotXml);
  });
});

describe('Comment VML is not captured as pass-through', () => {
  it('leaves xl/drawings/vmlDrawingN.vml outside the passthrough bucket', async () => {
    // Build a workbook with no comments / passthrough — round-trip should
    // produce no passthrough entries.
    const wb = createWorkbook();
    addWorksheet(wb, 'Sheet1');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.passthrough).toBeUndefined();
  });
});

describe('Modern Excel parts pass-through', () => {
  it(
    'round-trips externalLinks / richData / threadedComments / timelines / workbookCache verbatim',
    async () => {
      const wb = createWorkbook();
      addWorksheet(wb, 'Sheet1');
      wb.passthrough = new Map();
      wb.passthroughContentTypes = new Map();

      const cases: Array<{ path: string; ct: string }> = [
        // External workbook links — formula refs to other workbooks.
        {
          path: 'xl/externalLinks/externalLink1.xml',
          ct: 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
        },
        {
          path: 'xl/externalLinks/_rels/externalLink1.xml.rels',
          ct: 'application/vnd.openxmlformats-package.relationships+xml',
        },
        // Rich data types (Excel 365 stocks / geography).
        {
          path: 'xl/richData/rdRichValueTypes.xml',
          ct: 'application/vnd.ms-excel.rdrichvaluetypes+xml',
        },
        {
          path: 'xl/richData/rdRichValue.xml',
          ct: 'application/vnd.ms-excel.rdrichvalue+xml',
        },
        // Modern threaded comments (separate from legacy comments).
        {
          path: 'xl/threadedComments/threadedComment1.xml',
          ct: 'application/vnd.ms-excel.threadedcomments+xml',
        },
        // Pivot timeline filter.
        {
          path: 'xl/timelineCaches/timelineCache1.xml',
          ct: 'application/vnd.ms-excel.timelineCacheDefinition+xml',
        },
        {
          path: 'xl/timelines/timeline1.xml',
          ct: 'application/vnd.ms-excel.Timeline+xml',
        },
        // Power Query metadata.
        {
          path: 'xl/workbookCache/cache.xml',
          ct: 'application/xml',
        },
      ];
      for (let i = 0; i < cases.length; i++) {
        const c = cases[i];
        if (!c) continue;
        wb.passthrough.set(c.path, new Uint8Array([0xab, i]));
        wb.passthroughContentTypes.set(c.path, c.ct);
      }

      const bytes = await workbookToBytes(wb);
      const wb2 = await loadWorkbook(fromBuffer(bytes));

      for (const c of cases) {
        expect(wb2.passthrough?.get(c.path), `byte mismatch on ${c.path}`).toEqual(
          wb.passthrough.get(c.path),
        );
        expect(wb2.passthroughContentTypes?.get(c.path), `ct mismatch on ${c.path}`).toBe(c.ct);
      }
      expect(wb2.passthrough?.size).toBe(cases.length);
    },
  );
});
