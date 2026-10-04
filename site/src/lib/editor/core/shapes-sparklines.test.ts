import { describe, expect, test } from 'vitest';
import { EditorController } from './controller.svelte.ts';
import { insertShape, reorderDrawing, setShapeFill, setShapeText, shapeAt, shapeText, withText } from './shapes.ts';
import { clearSparklines, createSparklines, pairSparklines, setSparklineFlag, setSparklineType, sparklineValues } from './sparklines.ts';

function type(ctl: EditorController, row: number, col: number, text: string): void {
  ctl.selectCell({ row, col });
  ctl.startEdit(text);
  expect(ctl.commitEdit()).toBe(true);
}

describe('shapes', () => {
  test('a drag draws a shape and a click drops the default size; each is one undo step', () => {
    const ctl = new EditorController();
    const doc = ctl.doc;
    const i = insertShape(doc, doc, { kind: 'shape', prst: 'ellipse' }, { x: 100, y: 50 }, { x: 40, y: 10 });
    const j = insertShape(doc, doc, { kind: 'textBox' }, { x: 300, y: 50 }, { x: 300, y: 50 });
    expect([i, j]).toEqual([0, 1]);
    const ellipse = shapeAt(doc, 0);
    expect(ellipse?.spPr.geometry).toEqual({ kind: 'preset', prst: 'ellipse' });
    expect(shapeAt(doc, 1)?.textBox).toBe(true);
    doc.undo();
    expect(doc.ws.drawing?.items).toHaveLength(1);
  });

  test('a line drawn up and to the left keeps its direction as flips', () => {
    const ctl = new EditorController();
    insertShape(ctl.doc, ctl.doc, { kind: 'shape', prst: 'straightConnector1' }, { x: 200, y: 200 }, { x: 100, y: 120 });
    const line = shapeAt(ctl.doc, 0);
    expect(line?.connector).toBe(true);
    expect(line?.spPr.xfrm).toEqual({ flipH: true, flipV: true });
    expect(line?.spPr.ln?.tailEnd).toEqual({ type: 'triangle' });
  });

  test('text, fill and z-order edits', () => {
    const ctl = new EditorController();
    const doc = ctl.doc;
    insertShape(doc, doc, { kind: 'shape', prst: 'rect' }, { x: 0, y: 0 }, { x: 50, y: 50 });
    insertShape(doc, doc, { kind: 'shape', prst: 'star5' }, { x: 0, y: 0 }, { x: 50, y: 50 });
    setShapeText(doc, 0, 'one\ntwo');
    expect(shapeText(shapeAt(doc, 0)?.txBody)).toBe('one\ntwo');
    // Each paragraph keeps the shape's centring.
    expect(shapeAt(doc, 0)?.txBody?.paragraphs.map((p) => p.pPr?.algn)).toEqual(['ctr', 'ctr']);
    setShapeFill(doc, 0, '#FF0000');
    expect(shapeAt(doc, 0)?.spPr.fill).toEqual({ kind: 'solidFill', color: { base: { kind: 'srgb', value: 'FF0000' }, mods: [] } });
    expect(reorderDrawing(doc, 0, 1)).toBe(1);
    expect(shapeAt(doc, 1)?.spPr.geometry).toEqual({ kind: 'preset', prst: 'rect' });
    doc.undo();
    expect(shapeAt(doc, 0)?.spPr.geometry).toEqual({ kind: 'preset', prst: 'rect' });
  });

  test('withText keeps the first run font', () => {
    const body = withText({ bodyPr: {}, paragraphs: [{ runs: [{ kind: 'r', rPr: { sz: 1800, b: true }, t: 'old' }] }] }, 'new');
    expect(body.paragraphs[0]?.runs).toEqual([{ kind: 'r', rPr: { sz: 1800, b: true }, t: 'new' }]);
  });
});

describe('sparklines', () => {
  test('pairs each location cell with the data row or column lined up with it', () => {
    expect(pairSparklines('A1:E3', 'F1:F3', 'Sheet1')).toEqual([
      { formula: 'Sheet1!A1:E1', location: 'F1' },
      { formula: 'Sheet1!A2:E2', location: 'F2' },
      { formula: 'Sheet1!A3:E3', location: 'F3' },
    ]);
    expect(pairSparklines('A1:B5', 'A7:B7', 'Data')).toEqual([
      { formula: 'Data!A1:A5', location: 'A7' },
      { formula: 'Data!B1:B5', location: 'B7' },
    ]);
    expect(pairSparklines("'My Sheet'!A1:D1", 'E1', 'Sheet1')).toEqual([{ formula: "'My Sheet'!A1:D1", location: 'E1' }]);
    expect(pairSparklines('A1:E3', 'F1:G2', 'Sheet1')).toBe('sparklineBadLocation');
    expect(pairSparklines('A1:E3', 'F1:F2', 'Sheet1')).toBe('sparklineSizeMismatch');
    expect(pairSparklines('nonsense', 'F1', 'Sheet1')).toBe('sparklineBadData');
  });

  test('create, edit and clear through undoable steps', () => {
    const ctl = new EditorController();
    const doc = ctl.doc;
    [3, 'x', -2, 8].forEach((v, i) => type(ctl, 1, i + 1, String(v)));
    expect(createSparklines(doc, 'stacked', 'A1:D1', 'E1')).toBeUndefined();
    expect(sparklineValues(doc.calc, doc.ws.title, 'A1:D1')).toEqual([3, null, -2, 8]);
    expect(ctl.sparklines.get('1:5')?.group.type).toBe('stacked');
    setSparklineType(doc, 0, 'line');
    setSparklineFlag(doc, 0, 'markers', true);
    expect(doc.ws.sparklineGroups?.[0]).toMatchObject({ markers: true });
    expect(doc.ws.sparklineGroups?.[0]?.type).toBeUndefined();
    // A new group over the same cell replaces the old sparkline.
    createSparklines(doc, 'column', 'A1:D1', 'E1');
    expect(doc.ws.sparklineGroups?.map((g) => g.type)).toEqual(['column']);
    clearSparklines(doc, [{ r1: 1, c1: 5, r2: 1, c2: 5 }]);
    expect(doc.ws.sparklineGroups).toEqual([]);
    doc.undo();
    expect(doc.ws.sparklineGroups).toHaveLength(1);
  });
});
