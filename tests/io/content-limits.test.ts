// `contentLimits` bounds the quantity that decides what a load costs, which is
// cells, not inflated bytes. A service accepting uploads from strangers has to
// be able to say "nothing over N cells" and have the load stop there rather
// than after the model is built.

import { describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer, toBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import { loadWorkbookStream, type ReadOnlyWorksheet } from '../../src/streaming/read-only.js';
import { OpenXmlContentLimitError, OpenXmlError } from '../../src/utils/exceptions.js';
import { createWriteOnlyWorkbook } from '../../src/streaming/write-only.js';
import { addWorksheet, createWorkbook, getSheet } from '../../src/workbook/workbook.js';
import { getCell, setCell } from '../../src/worksheet/worksheet.js';

/** `rows` x `cols` of numbers on one sheet named Data. */
const archive = async (rows: number, cols: number): Promise<Uint8Array> => {
  const sink = toBuffer();
  const wb = await createWriteOnlyWorkbook(sink);
  const ws = await wb.addWorksheet('Data');
  for (let r = 0; r < rows; r++) {
    const row = new Array<number>(cols);
    for (let c = 0; c < cols; c++) row[c] = r * cols + c;
    await ws.appendRow(row);
  }
  await ws.close();
  await wb.finalize();
  return sink.result();
};

/** Two sheets of `rows` x 1, so a cap can be shown to cover the workbook. */
const twoSheetArchive = async (rows: number): Promise<Uint8Array> => {
  const wb = createWorkbook();
  for (const title of ['First', 'Second']) {
    const ws = addWorksheet(wb, title);
    for (let r = 1; r <= rows; r++) setCell(ws, r, 1, r);
  }
  return workbookToBytes(wb);
};

const drain = async (ws: ReadOnlyWorksheet, minRow?: number): Promise<unknown[][]> => {
  const rows: unknown[][] = [];
  for await (const row of ws.iterValues(minRow === undefined ? {} : { minRow })) rows.push(row);
  return rows;
};

describe('loadWorkbook with no contentLimits', () => {
  it('reads the workbook, as it always did', async () => {
    const wb = await loadWorkbook(fromBuffer(await archive(20, 4)));
    const ws = getSheet(wb, 'Data');
    if (ws === undefined) throw new Error('load produced no worksheet');
    expect(getCell(ws, 20, 4)?.value).toBe(79);
  });
});

describe('loadWorkbook with a cell cap', () => {
  it('reads a workbook that fits', async () => {
    const bytes = await archive(10, 4);
    const wb = await loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 40 } });
    expect(getSheet(wb, 'Data')).toBeDefined();
  });

  it('refuses one that does not, naming the cap and the cell that reached it', async () => {
    const bytes = await archive(10, 4);
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 39 } })).rejects.toThrow(
      OpenXmlContentLimitError,
    );
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 39 } })).rejects.toThrow(
      'worksheet: reading Data!D10 passes contentLimits.maxCells of 39',
    );
  });

  it('counts across every worksheet, not per sheet', async () => {
    const bytes = await twoSheetArchive(30);
    // 30 cells on each of two sheets: a 40-cell cap fits the first sheet and
    // has to stop partway through the second.
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 40 } })).rejects.toThrow(
      'worksheet: reading Second!A11 passes contentLimits.maxCells of 40',
    );
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 60 } })).resolves.toBeDefined();
  });

  it('stops before the cells past the cap are modelled', async () => {
    // 200k cells would take seconds and hundreds of MB to model. A cap of 10
    // has to refuse in the time it takes to read ten cells, which is what
    // makes the option worth having.
    const bytes = await archive(20_000, 10);
    const started = performance.now();
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 10 } })).rejects.toThrow(
      OpenXmlContentLimitError,
    );
    const elapsed = performance.now() - started;
    // Generous by two orders of magnitude against modelling all 200k cells;
    // this asserts the shape of the cost, not a machine's speed.
    expect(elapsed).toBeLessThan(2_000);
  });
});

describe('loadWorkbook with a row cap', () => {
  it('reads a workbook that fits and refuses one that does not', async () => {
    const bytes = await archive(10, 2);
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxRows: 10 } })).resolves.toBeDefined();
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxRows: 9 } })).rejects.toThrow(
      'worksheet: reading row 10 of Data passes contentLimits.maxRows of 9',
    );
  });
});

describe('a cap that cannot mean anything', () => {
  it('is rejected where the caller passes it', async () => {
    const bytes = await archive(2, 2);
    for (const bad of [0, -1, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
      await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: bad } })).rejects.toThrow(
        OpenXmlError,
      );
    }
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxCells: 0 } })).rejects.toThrow(
      'contentLimits.maxCells must be a positive integer; got 0',
    );
    await expect(loadWorkbook(fromBuffer(bytes), { contentLimits: { maxRows: -3 } })).rejects.toThrow(
      'contentLimits.maxRows must be a positive integer; got -3',
    );
  });
});

describe('loadWorkbookStream with a cell cap', () => {
  it('yields the rows that fit and then rejects the iterator', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await archive(10, 2)), {
      contentLimits: { maxCells: 4 },
    });
    try {
      await expect(drain(wb.openWorksheet('Data'))).rejects.toThrow(
        'worksheet: reading Data!A3 passes contentLimits.maxCells of 4',
      );
    } finally {
      await wb.close();
    }
  });

  it('applies the cap on the indexed band-query path too', async () => {
    // minRow > 1 replays from the row-offset index instead of streaming the
    // part, which is a second route into the row iterator.
    const wb = await loadWorkbookStream(fromBuffer(await archive(10, 2)), {
      contentLimits: { maxCells: 4 },
    });
    try {
      await expect(drain(wb.openWorksheet('Data'), 5)).rejects.toThrow(OpenXmlContentLimitError);
    } finally {
      await wb.close();
    }
  });

  it('counts per traversal, so a second pass over the same sheet still runs', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await archive(4, 2)), {
      contentLimits: { maxCells: 8 },
    });
    try {
      const ws = wb.openWorksheet('Data');
      expect(await drain(ws)).toHaveLength(4);
      expect(await drain(ws)).toHaveLength(4);
    } finally {
      await wb.close();
    }
  });

  it('reads everything when no cap is given', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await archive(10, 2)));
    try {
      expect(await drain(wb.openWorksheet('Data'))).toHaveLength(10);
    } finally {
      await wb.close();
    }
  });
});

describe('loadWorkbookStream with a row cap', () => {
  it('rejects the iterator at the row past the cap', async () => {
    const wb = await loadWorkbookStream(fromBuffer(await archive(10, 2)), {
      contentLimits: { maxRows: 3 },
    });
    try {
      await expect(drain(wb.openWorksheet('Data'))).rejects.toThrow(
        'worksheet: reading row 4 of Data passes contentLimits.maxRows of 3',
      );
    } finally {
      await wb.close();
    }
  });
});
