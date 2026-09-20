// Synthetic namespace variants supplement the genuine Excel fixtures in strict-read.test.ts.

import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { loadWorkbookStream } from '../../src/streaming/read-only.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import { makeBarChart, makeBarSeries, makeChartSpace } from '../../src/chart/chart.js';
import { addChartAt } from '../../src/drawing/drawing.js';
import { validateXlsx } from '../conformance/validate.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';
import { openZip } from '../../src/zip/reader.js';
import { createZipWriter } from '../../src/zip/writer.js';

/**
 * Transitional namespace to its strict counterpart. Two of the pairs rename
 * the local part as well (`extended-properties`, `custom-properties`), which
 * is why this is a table rather than a prefix swap.
 */
const STRICT_NAMESPACES: ReadonlyArray<readonly [string, string]> = [
  ...['chart', 'spreadsheetDrawing', 'chartDrawing', 'picture'].map((part): readonly [string, string] => [
    `http://schemas.openxmlformats.org/drawingml/2006/${part}`,
    `http://purl.oclc.org/ooxml/drawingml/${part}`,
  ]),
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'http://purl.oclc.org/ooxml/officeDocument/relationships',
  ],
  [
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'http://purl.oclc.org/ooxml/spreadsheetml/main',
  ],
  [
    'http://schemas.openxmlformats.org/drawingml/2006/main',
    'http://purl.oclc.org/ooxml/drawingml/main',
  ],
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'http://purl.oclc.org/ooxml/officeDocument/extendedProperties',
  ],
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties',
    'http://purl.oclc.org/ooxml/officeDocument/customProperties',
  ],
  [
    'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
    'http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes',
  ],
];

/** Namespaces the package layer keeps in both variants (ECMA-376 part 2). */
const OPC_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OPC_CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

/** Dublin Core, which docProps/core.xml carries in both variants. */
const DCORE_NS = 'http://purl.org/dc/elements/1.1/';

const WORKSHEET_PART = 'xl/worksheets/sheet1.xml';
const WORKBOOK_PART = 'xl/workbook.xml';

/** Rewrite the namespace URIs of every XML part `appliesTo` accepts. */
const rezipWithNamespaces = async (
  bytes: Uint8Array,
  pairs: ReadonlyArray<readonly [string, string]>,
  appliesTo: (path: string) => boolean = () => true,
): Promise<Uint8Array> => {
  const archive = await openZip(fromBuffer(bytes));
  const sink = toBuffer();
  const writer = createZipWriter(sink);
  const decoder = new TextDecoder('utf-8');
  const encoder = new TextEncoder();
  try {
    for (const path of archive.list()) {
      const payload = archive.read(path);
      const isXml = path.endsWith('.xml') || path.endsWith('.rels');
      if (!isXml || !appliesTo(path)) {
        await writer.addEntry(path, payload);
        continue;
      }
      let text = decoder.decode(payload);
      for (const [from, to] of pairs) text = text.split(from).join(to);
      await writer.addEntry(path, encoder.encode(text));
    }
    await writer.finalize();
  } catch (cause) {
    await writer.abort(cause);
    throw cause;
  } finally {
    archive.close();
  }
  return sink.result();
};

/** Build once: each call is a save, an inflate, a rewrite and a deflate. */
const memo = <T>(build: () => Promise<T>): (() => Promise<T>) => {
  let pending: Promise<T> | undefined;
  return () => (pending ??= build());
};

const transitionalPackage = memo(async (): Promise<Uint8Array> => {
  const wb = createWorkbook();
  // core.xml pins the Dublin Core URIs the detector must ignore.
  wb.properties = { creator: 'iso-strict fixture' };
  const ws = addWorksheet(wb, 'Data');
  setCell(ws, 1, 1, 'header');
  setCell(ws, 2, 1, 42);
  return workbookToBytes(wb);
});

const strictPackage = memo(async (): Promise<Uint8Array> =>
  rezipWithNamespaces(await transitionalPackage(), STRICT_NAMESPACES),
);

const expectReadable = async (bytes: Uint8Array): Promise<void> => {
  const wb = await loadWorkbook(fromBuffer(bytes));
  expect(wb.sheets.map((s) => s.sheet.title)).toEqual(['Data']);
  const first = wb.sheets[0];
  if (first?.kind !== 'worksheet') throw new Error('expected worksheet');
  expect(getCell(first.sheet, 1, 1)?.value).toBe('header');
  expect(getCell(first.sheet, 2, 1)?.value).toBe(42);
  const streamed = await loadWorkbookStream(fromBuffer(bytes));
  try {
    for (const options of [{}, { minRow: 2, maxRow: 2 }]) {
      const rows = [];
      for await (const row of streamed.openWorksheet('Data').iterRows(options)) rows.push(row);
      expect(rows.length).toBe(options.minRow === 2 ? 1 : 2);
    }
  } finally { await streamed.close(); }
};

const partText = async (bytes: Uint8Array, path: string): Promise<string> => {
  const archive = await openZip(fromBuffer(bytes));
  try {
    return new TextDecoder().decode(archive.read(path));
  } finally {
    archive.close();
  }
};

describe('loadWorkbook on an ISO 29500 strict package', () => {
  it('reads Strict package relationships and parts', async () => {
    await expectReadable(await strictPackage());
  });

  it('reads Strict parts behind Transitional package relationships', async () => {
    const partsOnly = STRICT_NAMESPACES.filter(([from]) => !from.endsWith('/relationships'));
    await expectReadable(await rezipWithNamespaces(await transitionalPackage(), partsOnly));
  });

  for (const part of [WORKBOOK_PART, WORKSHEET_PART]) {
    it(`reads a Strict ${part} in a mixed package`, async () => {
      await expectReadable(await rezipWithNamespaces(await transitionalPackage(), STRICT_NAMESPACES, (path) => path === part));
    });
  }

  it('keeps the OPC package namespaces, so detection cannot key off them', async () => {
    // Strict rewrites the markup namespaces and leaves ECMA-376 part 2 alone:
    // the container still parses, which is why the rel *type* is what gives it
    // away in the root rels.
    const strict = await strictPackage();
    expect(await partText(strict, '_rels/.rels')).toContain(OPC_RELS_NS);
    expect(await partText(strict, '[Content_Types].xml')).toContain(OPC_CONTENT_TYPES_NS);
  });
});

describe('loadWorkbook on a transitional package', () => {
  it('loads it, Dublin Core purl.org URIs and all', async () => {
    // Guards against a detector written as "any purl URI": docProps/core.xml
    // is purl.org/dc in both variants.
    const bytes = await transitionalPackage();
    expect(await partText(bytes, 'docProps/core.xml')).toContain(DCORE_NS);
    const wb = await loadWorkbook(fromBuffer(bytes));
    expect(wb.sheets.map((s) => s.sheet.title)).toEqual(['Data']);
    const streamed = await loadWorkbookStream(fromBuffer(bytes));
    expect(streamed.sheetNames).toEqual(['Data']);
    await streamed.close();
  });

  it('rejects a workbook root in an unrelated namespace rather than reading it as sheetless', async () => {
    const wrongNs = await rezipWithNamespaces(
      await transitionalPackage(),
      [
        [
          'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
          'http://schemas.openxmlformats.org/spreadsheetml/2099/main',
        ],
      ],
      (path) => path === WORKBOOK_PART,
    );
    await expect(loadWorkbook(fromBuffer(wrongNs))).rejects.toBeInstanceOf(OpenXmlSchemaError);
    await expect(loadWorkbook(fromBuffer(wrongNs))).rejects.toThrow(/expected workbook/);
    await expect(loadWorkbookStream(fromBuffer(wrongNs))).rejects.toThrow(/expected workbook/);
  });
});

it('round-trips synthetic Strict chart and drawing namespaces', async () => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Data');
  setCell(ws, 1, 1, 42);
  addChartAt(ws, 'D3', { space: makeChartSpace({ title: 'Values', plotArea: {
    chart: makeBarChart({ series: [makeBarSeries({ idx: 0, val: { ref: 'Data!$A$1', cache: [42] } })] }),
  } }) });
  const strict = await rezipWithNamespaces(await workbookToBytes(wb), STRICT_NAMESPACES);
  const loaded = await loadWorkbook(fromBuffer(strict));
  const ref = loaded.sheets[0];
  if (ref?.kind !== 'worksheet') throw new Error('expected worksheet');
  expect(ref.sheet.drawing?.items[0]?.content).toMatchObject({ kind: 'chart', chart: { space: { title: { text: 'Values' } } } });
  const saved = await workbookToBytes(loaded);
  expect((await validateXlsx(saved)).issues).toEqual([]);
  expect(await partText(saved, 'xl/charts/chart1.xml')).not.toContain('http://purl.oclc.org/ooxml/');
});
