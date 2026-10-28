import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { assert, describe, expect, it } from 'vitest';
import { loadWorkbook } from '../../src/io/load.js';
import { fromBuffer } from '../../src/io/node.js';
import { workbookToBytes } from '../../src/io/save.js';
import type { Workbook } from '../../src/workbook/workbook.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import type { Worksheet } from '../../src/worksheet/worksheet.js';
import { editCommentAuthor, editCommentText, renameCommentAuthor, setComment, setComments } from '../../src/worksheet/worksheet.js';
import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';
import { type LegacyComment, makeLegacyComment } from '../../src/worksheet/comments.js';
import { SHEET_MAIN_NS } from '../../src/xml/namespaces.js';
import { validateXlsx } from '../conformance/validate.js';

// The shapes Excel 12 wrote in Apache POI's SimpleWithComments.xlsx (#307):
// single-quoted styles, whitespace around the anchor numbers, and an
// x:Visible on the note that is shown.
const BOLD = '<rPr><b/><sz val="8"/><color indexed="81"/><rFont val="Tahoma"/></rPr>';
const PLAIN = '<rPr><sz val="8"/><color indexed="81"/><rFont val="Tahoma"/></rPr>';
const note = (ref: string, body: string) =>
  `<comment ref="${ref}" authorId="0"><text><r>${BOLD}<t>Yegor Kozlov:</t></r><r>${PLAIN}<t xml:space="preserve">\n${body}</t></r></text></comment>`;
const COMMENTS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<comments xmlns="${SHEET_MAIN_NS}"><authors><author>Yegor Kozlov</author></authors><commentList>` +
  `${note('B1', 'first cell')}${note('B2', 'second cell')}${note('B3:C4', 'third cell')}</commentList></comments>`;
const shape = (anchor: string, row: number, extra = '', objectType = 'Note') =>
  `<v:shape id="_x0000_s1025" type="#_x0000_t202" style='position:absolute;\n margin-left:107.25pt;width:96pt;height:55.5pt;z-index:1;\n visibility:hidden' fillcolor="infoBackground [80]" o:insetmode="auto">` +
  `<v:textbox style='mso-direction-alt:auto'>\n <div style='text-align:left'></div>\n</v:textbox>` +
  `<x:ClientData ObjectType="${objectType}">\n <x:MoveWithCells/>\n <x:SizeWithCells/>\n <x:Anchor>\n    ${anchor}</x:Anchor>\n` +
  ` <x:AutoFill>False</x:AutoFill>\n <x:Row>${row}</x:Row>\n <x:Column>1</x:Column>${extra}\n</x:ClientData></v:shape>`;
const vml = (shapes: string) =>
  '<xml xmlns:v="urn:schemas-microsoft-com:vml"\n xmlns:o="urn:schemas-microsoft-com:office:office"\n xmlns:x="urn:schemas-microsoft-com:office:excel">\n' +
  `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>${shapes}</xml>`;
const EXCEL_VML = vml(
  shape('2, 15, 0, 2, 4, 15, 4, 8', 0) + shape('2, 15, 0, 7, 4, 15, 4, 13', 1) + shape('2, 15, 1, 7, 4, 15, 5, 13', 2, '\n <x:Visible/>'),
);

async function fixture(vmlXml: string = EXCEL_VML): Promise<Uint8Array> {
  const wb = createWorkbook();
  const ws = addWorksheet(wb, 'Sheet1');
  setComments(ws, [
    { ref: 'B1', author: 'x', text: 'x' },
    { ref: 'B2', author: 'x', text: 'x' },
    { ref: 'B3:C4', author: 'x', text: 'x' },
  ]);
  const parts = unzipSync(await workbookToBytes(wb));
  assert(parts['xl/comments1.xml'] && parts['xl/drawings/vmlDrawing1.vml'], 'expected comment parts');
  parts['xl/comments1.xml'] = strToU8(COMMENTS_XML);
  parts['xl/drawings/vmlDrawing1.vml'] = strToU8(vmlXml);
  return zipSync(parts);
}

const firstSheet = (wb: Workbook): Worksheet => {
  const ref = wb.sheets[0];
  assert(ref?.kind === 'worksheet', 'expected a worksheet');
  return ref.sheet;
};

const part = (bytes: Uint8Array, name: string): string => {
  const p = unzipSync(bytes)[name];
  assert(p, `expected ${name}`);
  return strFromU8(p);
};

const runs = (body: string) => [
  { text: 'Yegor Kozlov:', font: { b: true, sz: 8, color: { indexed: 81 }, name: 'Tahoma' } },
  { text: `\n${body}`, font: { sz: 8, color: { indexed: 81 }, name: 'Tahoma' } },
];

describe('issue #307 — notes keep their rich text and their box', () => {
  it('loads runs, anchor and visibility', async () => {
    const ws = firstSheet(await loadWorkbook(fromBuffer(await fixture())));
    expect(ws.legacyComments).toEqual<LegacyComment[]>([
      {
        ref: 'B1',
        author: 'Yegor Kozlov',
        text: 'Yegor Kozlov:\nfirst cell',
        richText: runs('first cell'),
        anchor: { fromCol: 2, fromColOffset: 15, fromRow: 0, fromRowOffset: 2, toCol: 4, toColOffset: 15, toRow: 4, toRowOffset: 8 },
      },
      {
        ref: 'B2',
        author: 'Yegor Kozlov',
        text: 'Yegor Kozlov:\nsecond cell',
        richText: runs('second cell'),
        anchor: { fromCol: 2, fromColOffset: 15, fromRow: 0, fromRowOffset: 7, toCol: 4, toColOffset: 15, toRow: 4, toRowOffset: 13 },
      },
      {
        ref: 'B3:C4',
        author: 'Yegor Kozlov',
        text: 'Yegor Kozlov:\nthird cell',
        richText: runs('third cell'),
        visible: true,
        anchor: { fromCol: 2, fromColOffset: 15, fromRow: 1, fromRowOffset: 7, toCol: 4, toColOffset: 15, toRow: 5, toRowOffset: 13 },
      },
    ]);
  });

  it('saves them back, and a second load sees the same notes', async () => {
    const wb = await loadWorkbook(fromBuffer(await fixture()));
    const saved = await workbookToBytes(wb);
    expect((await validateXlsx(saved)).issues).toEqual([]);
    const comments = part(saved, 'xl/comments1.xml');
    expect(comments).toContain(`<text><r>${BOLD.replace('<b/><sz val="8"/><color indexed="81"/><rFont val="Tahoma"/>', '<rFont val="Tahoma"/><b/><color indexed="81"/><sz val="8"/>')}<t>Yegor Kozlov:</t></r>`);
    const drawing = part(saved, 'xl/drawings/vmlDrawing1.vml');
    expect(drawing.match(/<x:Anchor>[^<]*<\/x:Anchor>/g)).toEqual([
      '<x:Anchor>2, 15, 0, 2, 4, 15, 4, 8</x:Anchor>',
      '<x:Anchor>2, 15, 0, 7, 4, 15, 4, 13</x:Anchor>',
      '<x:Anchor>2, 15, 1, 7, 4, 15, 5, 13</x:Anchor>',
    ]);
    expect(drawing.match(/<x:Visible\/>/g)).toHaveLength(1);
    expect(drawing.match(/visibility:visible/g)).toHaveLength(1);
    expect(firstSheet(await loadWorkbook(fromBuffer(saved))).legacyComments).toEqual(firstSheet(wb).legacyComments);
  });

  it('writes plain text once the text no longer matches its runs', async () => {
    const wb = await loadWorkbook(fromBuffer(await fixture()));
    const ws = firstSheet(wb);
    const [b1, b2] = ws.legacyComments;
    assert(b1 && b2, 'expected notes');
    b1.text = 'rewritten';
    expect(editCommentText(ws, 'B2', 'edited')).toBe(true);
    expect(ws.legacyComments[1]).toEqual({ ref: 'B2', author: 'Yegor Kozlov', text: 'edited', anchor: b2.anchor });
    expect(editCommentAuthor(ws, 'B3:C4', 'Someone')).toBe(true);
    expect(ws.legacyComments[2]).toMatchObject({ author: 'Someone', visible: true, richText: runs('third cell') });
    const back = firstSheet(await loadWorkbook(fromBuffer(await workbookToBytes(wb)))).legacyComments;
    expect(back.map((c) => [c.text, c.richText !== undefined, c.anchor !== undefined, c.visible])).toEqual([
      ['rewritten', false, true, undefined],
      ['edited', false, true, undefined],
      ['Yegor Kozlov:\nthird cell', true, true, true],
    ]);
  });

  it('renames the author of a note whose text was edited directly', async () => {
    const ws = firstSheet(await loadWorkbook(fromBuffer(await fixture())));
    const b1 = ws.legacyComments[0];
    assert(b1, 'expected a note');
    b1.text = 'changed';
    expect(renameCommentAuthor(ws, 'Yegor Kozlov', 'Z')).toBe(3);
    expect(ws.legacyComments.map((c) => [c.author, c.richText !== undefined])).toEqual([
      ['Z', false],
      ['Z', true],
      ['Z', true],
    ]);
    const b2 = ws.legacyComments[1];
    assert(b2, 'expected a note');
    b2.text = 'also changed';
    expect(editCommentAuthor(ws, 'B2', 'Q')).toBe(true);
    expect(ws.legacyComments[1]).not.toHaveProperty('richText');
  });

  it('writes notes built from scratch with a box', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    const anchor = { fromCol: 1, fromColOffset: 0, fromRow: 0, fromRowOffset: 0, toCol: 3, toColOffset: 0, toRow: 4, toRowOffset: 0 };
    setComment(ws, { ref: 'A1', author: 'me', text: 'hi', visible: true, anchor });
    const saved = await workbookToBytes(wb);
    expect((await validateXlsx(saved)).issues).toEqual([]);
    expect(firstSheet(await loadWorkbook(fromBuffer(saved))).legacyComments).toEqual([
      { ref: 'A1', author: 'me', text: 'hi', visible: true, anchor },
    ]);
  });

  it('reads a VML boolean x:Visible with False content as hidden', async () => {
    const ws = firstSheet(await loadWorkbook(fromBuffer(await fixture(vml(shape('1, 0, 0, 0, 3, 0, 4, 0', 0, '<x:Visible>False</x:Visible>'))))));
    expect(ws.legacyComments[0]).toMatchObject({ visible: false });
    expect(ws.legacyComments[1]).not.toHaveProperty('anchor');
  });

  it('matches shapes to notes by cell, not by order, and ignores other objects', async () => {
    const shuffled = vml(
      shape('9, 0, 9, 0, 9, 0, 9, 0', 0, '', 'Checkbox') +
        shape('2, 15, 1, 7, 4, 15, 5, 13', 2, '<x:Visible/>') +
        shape('2, 15, 0, 2, 4, 15, 4, 8', 0),
    );
    const ws = firstSheet(await loadWorkbook(fromBuffer(await fixture(shuffled))));
    expect(ws.legacyComments.map((c) => [c.ref, c.anchor?.fromRowOffset, c.visible])).toEqual([
      ['B1', 2, undefined],
      ['B2', undefined, undefined],
      ['B3:C4', 7, true],
    ]);
  });

  it('refuses runs that do not join to the text, and anchors that are not non-negative integers', async () => {
    expect(() => makeLegacyComment({ ref: 'A1', author: 'a', text: 'plain', richText: [{ text: 'bold', font: { b: true } }] })).toThrow(
      OpenXmlSchemaError,
    );
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Sheet1');
    const anchor = { fromCol: 1, fromColOffset: 0, fromRow: 0, fromRowOffset: 0, toCol: 3, toColOffset: 0, toRow: 4, toRowOffset: 1.5 };
    setComment(ws, { ref: 'A1', author: 'me', text: 'hi', anchor });
    await expect(workbookToBytes(wb)).rejects.toThrow(OpenXmlSchemaError);
  });

  it.each([
    ['VML that is not well-formed XML', vml(shape('1, 0, 0, 0, 3, 0, 4, 0', 0).replace("<div style='text-align:left'></div>", '<div>a<br>b</div>'))],
    ['an anchor with the wrong number of values', vml(shape('1, 0, 0, 0, 3, 0, 4', 0))],
    ['an anchor with a negative value', vml(shape('1, 0, 0, 0, 3, 0, 4, -1', 0))],
    ['an anchor too large to hold exactly', vml(shape('1, 0, 0, 0, 3, 0, 4, 99999999999999999999', 0))],
  ])('still loads %s, without a box', async (_label, vmlXml) => {
    const ws = firstSheet(await loadWorkbook(fromBuffer(await fixture(vmlXml))));
    expect(ws.legacyComments).toHaveLength(3);
    expect(ws.legacyComments[0]).not.toHaveProperty('anchor');
    expect(ws.legacyComments[0]?.richText).toEqual(runs('first cell'));
  });
});
