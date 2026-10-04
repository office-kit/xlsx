import { describe, expect, it } from 'vitest';
import { applyPageForm, pageForm } from '../dialogs/page-setup.ts';
import { EditorController } from './controller.svelte.ts';
import { buildHeader, headerFooterFor, hfPresets, hfSegments, parseHeader, presetLabel, renderHf } from './header-footer.ts';
import { headerFooterCss, printHtml, printTitles } from './print.ts';

const FIELDS = { date: '2026/10/4', time: '9:30', file: 'Book1.xlsx', path: '/x/', sheet: 'Sheet1' };

describe('headerFooterFor', () => {
  it('picks the first-page and even-page variants when enabled', () => {
    const hf = { differentFirst: true, differentOddEven: true, firstHeader: 'F', evenHeader: 'E', oddHeader: 'O' };
    expect([1, 2, 3].map((p) => headerFooterFor(hf, p, 'header'))).toEqual(['F', 'E', 'O']);
    expect(headerFooterFor({ oddFooter: 'X' }, 2, 'footer')).toBe('X');
  });
});

describe('header / footer codes', () => {
  it('splits sections and keeps literal ampersands', () => {
    expect(parseHeader('&LA&&B&CPage &P&RZ')).toEqual({ left: 'A&&B', center: 'Page &P', right: 'Z' });
    expect(parseHeader('plain')).toEqual({ left: '', center: 'plain', right: '' });
    expect(buildHeader({ left: '', center: '&A', right: '' })).toBe('&C&A');
    expect(buildHeader({ left: '', center: '', right: '' })).toBeUndefined();
  });

  it('fills fields, numbers pages and drops formatting codes', () => {
    expect(renderHf('Page &P of &N', FIELDS, 2, 5)).toBe('Page 2 of 5');
    expect(renderHf('&D &T &F &A &Z&F', FIELDS, 1, 1)).toBe('2026/10/4 9:30 Book1.xlsx Sheet1 /x/Book1.xlsx');
    expect(renderHf('&"Arial,Bold"&12&KFF0000&B&IHi &&', FIELDS, 1, 1)).toBe('Hi &');
    expect(renderHf('&P+2', FIELDS, 1, 1)).toBe('3');
    expect(hfSegments('p&P/&N', FIELDS)).toEqual(['p', { counter: 'page', offset: 0 }, '/', { counter: 'pages', offset: 0 }]);
  });

  it('labels presets as Excel lists them', () => {
    const presets = hfPresets({ page: 'Page &P', pageOf: 'Page &P of &N', confidential: 'Confidential' });
    expect(presets[0]).toBe('');
    expect(presets.map((p) => presetLabel(p, FIELDS))).toContain('Confidential, 2026/10/4, Page 1');
    expect(presetLabel('&C&A&R&F', FIELDS)).toBe('Sheet1, Book1.xlsx');
  });
});

describe('page setup header / footer', () => {
  it('round-trips first and even page sections and the option flags', () => {
    const ctl = new EditorController();
    const form = pageForm(ctl);
    form.headers.odd.header.center = '&A';
    form.headers.first.footer.right = 'First';
    form.headers.even.header.left = 'Even';
    form.differentFirst = true;
    form.scaleWithDoc = false;
    expect(applyPageForm(ctl, form)).toBeUndefined();
    // Even sections are dropped while "different odd and even" is off.
    expect(ctl.doc.ws.headerFooter).toEqual({ differentFirst: true, scaleWithDoc: false, oddHeader: '&C&A', firstFooter: '&RFirst' });
    const back = pageForm(ctl);
    expect(back.headers.first.footer.right).toBe('First');
    expect(back.alignWithMargins).toBe(true);
    ctl.doc.undo();
    expect(ctl.doc.ws.headerFooter ?? {}).toEqual({});
  });
});

describe('printing', () => {
  it('puts header / footer sections into page-margin boxes', () => {
    const css = headerFooterCss({ oddHeader: '&LQ "1"&CPage &P of &N', differentFirst: true, firstFooter: '&R</style>' }, FIELDS);
    expect(css).toContain('@top-left { content: "Q \\22 1\\22 "');
    expect(css).toContain('@top-center { content: "Page " counter(page) " of " counter(pages)');
    expect(css).toContain('@page :first');
    expect(css).toContain('@bottom-right { content: "\\3c /style>"');
    expect(css).not.toContain('</style>');
  });

  it('repeats title rows in <thead> and prepends title columns', () => {
    const ctl = new EditorController();
    for (const [row, col, text] of [[1, 1, 'Key'], [1, 3, 'Head'], [2, 3, 'x'], [3, 3, 'y']] as const) {
      ctl.selectCell({ row, col });
      ctl.startEdit(text);
      ctl.commitEdit();
    }
    const form = pageForm(ctl);
    form.printArea = 'C1:C3';
    form.titleRows = '$1:$1';
    form.titleCols = '$A:$A';
    expect(applyPageForm(ctl, form)).toBeUndefined();
    expect(printTitles(ctl)).toEqual({ rows: [1, 1], cols: [1, 1] });
    const html = printHtml(ctl);
    const thead = html.slice(html.indexOf('<thead>'), html.indexOf('</thead>'));
    expect(thead).toContain('Key');
    expect(thead).toContain('Head');
    const tbody = html.slice(html.indexOf('<tbody>'));
    expect(tbody).not.toContain('Head');
    expect(tbody.match(/<tr/g)).toHaveLength(2);
  });
});
