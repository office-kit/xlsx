import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node.js';
import { loadWorkbook } from '../../src/io/load.js';
import { workbookToBytes } from '../../src/io/save.js';
import { makePerson } from '../../src/workbook/persons.js';
import { addWorksheet, createWorkbook, duplicateSheet } from '../../src/workbook/workbook.js';
import { makeLegacyComment } from '../../src/worksheet/comments.js';
import { makeThreadedComment } from '../../src/worksheet/threaded-comments.js';
import {
  parsePersonsXml,
  parseThreadedCommentsXml,
  personsToBytes,
  threadedCommentsToBytes,
  threadPlaceholders,
} from '../../src/worksheet/threaded-comments-xml.js';
import type { Worksheet } from '../../src/worksheet/worksheet.js';

const decode = (b: Uint8Array | undefined): string => new TextDecoder().decode(b);

const firstWorksheet = (wb: Awaited<ReturnType<typeof loadWorkbook>>): Worksheet => {
  const ref = wb.sheets[0];
  if (ref?.kind !== 'worksheet') throw new Error('expected a worksheet');
  return ref.sheet;
};

const buildThreadWorkbook = () => {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Sheet1');
  const alice = makePerson({ displayName: 'Alice' });
  const bob = makePerson({ displayName: 'Bob', userId: 'bob@example.com', providerId: 'AD' });
  (wb.persons ??= []).push(alice, bob);
  const root = makeThreadedComment({ ref: 'B2', personId: alice.id, text: 'Is this right?' });
  root.done = true;
  const reply = makeThreadedComment({ ref: 'B2', personId: bob.id, text: 'Yes & <checked>', parentId: root.id });
  (ws.threadedComments ??= []).push(root, reply);
  return { wb, ws, alice, bob, root, reply };
};

describe('makeThreadedComment / makePerson', () => {
  it('assigns braced upper-case GUIDs and an Excel-style timestamp', () => {
    const p = makePerson({ displayName: 'Alice' });
    expect(p.id).toMatch(/^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/);
    expect(p).toMatchObject({ displayName: 'Alice', userId: 'Alice', providerId: 'None' });
    const c = makeThreadedComment({ ref: 'A1', personId: p.id, text: 'hi', created: new Date('2024-05-01T09:30:00Z') });
    expect(c.created).toBe('2024-05-01T09:30:00.00');
    expect(c.id).not.toBe(p.id);
    expect(c.parentId).toBeUndefined();
  });
});

describe('threaded comments XML', () => {
  it('round-trips roots, replies, resolved state and mentions', () => {
    const { root, reply, bob } = buildThreadWorkbook();
    reply.mentions = [{ personId: bob.id, mentionId: '{M1}', startIndex: 0, length: 3 }];
    const parsed = parseThreadedCommentsXml(threadedCommentsToBytes([root, reply]));
    expect(parsed).toEqual([root, reply]);
  });

  it('round-trips the person list', () => {
    const { wb } = buildThreadWorkbook();
    expect(parsePersonsXml(personsToBytes(wb.persons ?? []))).toEqual(wb.persons);
  });

  it('builds one Excel-style legacy placeholder per thread', () => {
    const { root, reply } = buildThreadWorkbook();
    const [placeholder, ...rest] = threadPlaceholders([root, reply]);
    expect(rest).toEqual([]);
    expect(placeholder?.author).toBe(`tc=${root.id}`);
    expect(placeholder?.uid).toBe(root.id);
    expect(placeholder?.text.startsWith('[Threaded comment]\n\n')).toBe(true);
    expect(placeholder?.text.endsWith('Comment:\n    Is this right?\nReply:\n    Yes & <checked>')).toBe(true);
  });
});

describe('threaded comments in a saved workbook', () => {
  it('writes the parts, rels and content types Excel expects', async () => {
    const { wb, root } = buildThreadWorkbook();
    const entries = unzipSync(await workbookToBytes(wb));
    expect(Object.keys(entries)).toEqual(
      expect.arrayContaining([
        'xl/threadedComments/threadedComment1.xml',
        'xl/persons/person.xml',
        'xl/comments1.xml',
        'xl/drawings/vmlDrawing1.vml',
      ]),
    );
    const ct = decode(entries['[Content_Types].xml']);
    expect(ct).toContain('PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"');
    expect(ct).toContain('PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"');
    expect(decode(entries['xl/_rels/workbook.xml.rels'])).toContain(
      'Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"',
    );
    expect(decode(entries['xl/worksheets/_rels/sheet1.xml.rels'])).toContain(
      'Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"',
    );
    expect(decode(entries['xl/worksheets/sheet1.xml'])).toMatch(/<legacyDrawing r:id="rId\d+"\/>/);
    const comments = decode(entries['xl/comments1.xml']);
    expect(comments).toContain(`<author>tc=${root.id}</author>`);
    expect(comments).toContain(`xr:uid="${root.id}"`);
    expect(decode(entries['xl/drawings/vmlDrawing1.vml'])).toContain('<x:Row>1</x:Row><x:Column>1</x:Column>');
  });

  it('round-trips threads and persons without turning placeholders into notes', async () => {
    const { wb, root, reply } = buildThreadWorkbook();
    firstWorksheet(wb).legacyComments.push(makeLegacyComment({ ref: 'D4', author: 'Carol', text: 'a plain note' }));
    const back = await loadWorkbook(fromBuffer(await workbookToBytes(wb)));
    const ws = firstWorksheet(back);
    expect(ws.threadedComments).toEqual([root, reply]);
    expect(ws.legacyComments).toEqual([{ ref: 'D4', author: 'Carol', text: 'a plain note' }]);
    expect(back.persons).toEqual(wb.persons);
    expect(back.passthrough?.size ?? 0).toBe(0);
    expect(back.workbookRelsExtras ?? []).toEqual([]);
    expect(ws.relsExtras ?? []).toEqual([]);

    // A second pass must not duplicate parts or placeholders.
    const entries = unzipSync(await workbookToBytes(back));
    expect(Object.keys(entries).filter((p) => p.startsWith('xl/threadedComments/'))).toEqual([
      'xl/threadedComments/threadedComment1.xml',
    ]);
    expect(decode(entries['xl/comments1.xml']).match(/<comment /g)?.length).toBe(2);
  });

  it('gives a duplicated sheet its own thread ids', () => {
    const { wb, root } = buildThreadWorkbook();
    const copy = duplicateSheet(wb, 'Sheet1', 'Sheet2');
    const [copyRoot, copyReply] = copy.threadedComments ?? [];
    expect(copyRoot?.id).not.toBe(root.id);
    expect(copyReply?.parentId).toBe(copyRoot?.id);
    expect(copyRoot?.text).toBe(root.text);
  });

  it('gives a duplicated sheet its own mention ids and leaves the source alone', () => {
    const { wb, ws, bob, root } = buildThreadWorkbook();
    root.mentions = [{ personId: bob.id, mentionId: '{11111111-2222-3333-4444-555555555555}', startIndex: 0, length: 3 }];
    const before = structuredClone(ws.threadedComments);
    const copy = duplicateSheet(wb, 'Sheet1', 'Sheet2');
    const copied = copy.threadedComments?.[0]?.mentions?.[0];
    expect(copied?.mentionId).not.toBe('{11111111-2222-3333-4444-555555555555}');
    expect(copied?.personId).toBe(bob.id);
    expect(ws.threadedComments).toEqual(before);
  });
});

describe('an Excel-authored workbook with threaded comments', () => {
  // Saved by Excel for Mac: a two-comment thread on B2, a resolved thread on
  // D5, and a plain note on F7.
  const load = () => loadWorkbook(fromBuffer(readFileSync(resolve(__dirname, '../fixtures/threaded-comments/excel-mac.xlsx'))));

  it('models threads, persons and the remaining note', async () => {
    const wb = await load();
    const ws = firstWorksheet(wb);
    expect(wb.persons?.map((p) => p.displayName).sort()).toEqual(['Alice', 'Bob']);
    expect(ws.threadedComments?.map((c) => [c.ref, c.text, c.parentId !== undefined, c.done === true])).toEqual([
      ['B2', 'Is this right?', false, false],
      ['B2', 'Yes & <checked>', true, false],
      ['D5', 'Resolved one', false, true],
    ]);
    expect(ws.legacyComments.map((c) => c.ref)).toEqual(['F7']);
    expect(wb.passthrough?.size ?? 0).toBe(0);
  });

  it('survives load → save → load unchanged', async () => {
    const wb = await load();
    const back = await loadWorkbook(fromBuffer(await workbookToBytes(wb)));
    expect(firstWorksheet(back).threadedComments).toEqual(firstWorksheet(wb).threadedComments);
    expect(firstWorksheet(back).legacyComments).toEqual(firstWorksheet(wb).legacyComments);
    expect(back.persons).toEqual(wb.persons);
  });
});
