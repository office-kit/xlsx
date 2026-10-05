// Printing: lay the print area (or used range) out as an HTML table in a
// hidden frame and hand it to the browser's print dialog, honouring the
// sheet's page setup where CSS can express it (orientation, margins,
// gridlines, headings, centring, scale, print titles, header / footer).

import { getCellAt, usedRange } from './cells.ts';
import { colLetter, MAX_COL, MAX_ROW, parseRangeAddress, type Range } from './address.ts';
import type { EditorController } from './controller.svelte.ts';
import { hfSegments, parseHeader, type HfFields } from './header-footer.ts';
import type { HeaderFooter } from '@office-kit/xlsx/worksheet';
import type { RenderStyle, StrokeStyle } from './render-style.ts';
import { displayCell } from '../grid/display.ts';

const ESCAPES: Readonly<Record<string, string>> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };
const escapeHtml = (s: string) => s.replace(/[&<>"]/g, (c) => ESCAPES[c] ?? c);

function stroke(s: StrokeStyle | null): string {
  if (!s) return 'none';
  return `${s.double ? 3 : s.width}px ${s.double ? 'double' : s.dash.length ? 'dashed' : 'solid'} ${s.color}`;
}

function cellCss(st: RenderStyle, gridlines: boolean): string {
  const parts = [
    `font:${st.italic ? 'italic ' : ''}${st.bold ? 'bold ' : ''}${st.fontPx}px ${st.fontFamily}`,
    `color:${st.color}`,
    `text-align:${st.hAlign === 'general' ? 'inherit' : st.hAlign === 'centerContinuous' ? 'center' : st.hAlign === 'fill' || st.hAlign === 'distributed' ? 'justify' : st.hAlign}`,
    `vertical-align:${st.vAlign === 'center' ? 'middle' : st.vAlign === 'top' ? 'top' : 'bottom'}`,
    `white-space:${st.wrap ? 'pre-wrap' : 'pre'}`,
  ];
  if (st.fill) parts.push(`background:${st.fill}`);
  if (st.underline !== 'none' || st.strike) parts.push(`text-decoration:${st.underline !== 'none' ? 'underline' : ''} ${st.strike ? 'line-through' : ''}`);
  if (st.indent) parts.push(`padding-left:${st.indent * 9 + 2}px`);
  const edge = (s: StrokeStyle | null) => (s ? stroke(s) : gridlines ? '1px solid #d4d4d4' : 'none');
  parts.push(`border-top:${edge(st.top)}`, `border-right:${edge(st.right)}`, `border-bottom:${edge(st.bottom)}`, `border-left:${edge(st.left)}`);
  return parts.join(';');
}

function namedRanges(ctl: EditorController, name: string): Range[] {
  const doc = ctl.doc;
  const dn = doc.wb.definedNames.find((d) => d.name === name && d.scope === doc.activeSheetIndex);
  return (dn?.value ?? '').split(',').flatMap((part) => {
    const p = parseRangeAddress(part.trim());
    return p ? [p.range] : [];
  });
}

function printRanges(ctl: EditorController): Range[] {
  const ranges = namedRanges(ctl, '_xlnm.Print_Area');
  if (ranges.length) return ranges;
  const used = usedRange(ctl.doc.ws);
  return used ? [used] : [];
}

/** Rows and columns from _xlnm.Print_Titles: whole-row and whole-column ranges to repeat on every page. */
const PRINT_HASHES = '########';

export function printTitles(ctl: EditorController): { rows?: [number, number]; cols?: [number, number] } {
  const out: { rows?: [number, number]; cols?: [number, number] } = {};
  for (const r of namedRanges(ctl, '_xlnm.Print_Titles')) {
    if (r.c1 === 1 && r.c2 >= MAX_COL) out.rows = [r.r1, r.r2];
    else if (r.r1 === 1 && r.r2 >= MAX_ROW) out.cols = [r.c1, r.c2];
  }
  return out;
}

// CSS strings sit inside the document's <style>, so `<` is escaped too: a
// header reading "</style>" must not end the stylesheet.
const cssString = (s: string) => `"${s.replace(/[\\"<\n]/g, (c) => (c === '\n' ? '\\A ' : `\\${c.charCodeAt(0).toString(16)} `))}"`;

function sectionCss(section: string, fields: HfFields): string {
  // CSS counters cannot do arithmetic, so &P+n offsets print the plain page number.
  return hfSegments(section, fields)
    .map((seg) => (typeof seg === 'string' ? cssString(seg) : `counter(${seg.counter})`))
    .join(' ');
}

const BOXES = { header: ['top-left', 'top-center', 'top-right'], footer: ['bottom-left', 'bottom-center', 'bottom-right'] } as const;

function marginBoxes(header: string | undefined, footer: string | undefined, fields: HfFields): string {
  const out: string[] = [];
  for (const [kind, text] of [['header', header], ['footer', footer]] as const) {
    const parts = parseHeader(text);
    const [left, center, right] = BOXES[kind];
    for (const [box, section] of [[left, parts.left], [center, parts.center], [right, parts.right]] as const) {
      out.push(`@${box} { content: ${section ? sectionCss(section, fields) : 'none'}; font: 10pt sans-serif; }`);
    }
  }
  return out.join(' ');
}

/**
 * Header and footer as CSS page-margin boxes, so the print engine fills in
 * page numbers. First and even pages map to the `:first` and `:left` page
 * selectors (page 1 is a right-hand page).
 */
export function headerFooterCss(hf: HeaderFooter | undefined, fields: HfFields): string {
  if (!hf) return '';
  const rules = [`@page { ${marginBoxes(hf.oddHeader, hf.oddFooter, fields)} }`];
  if (hf.differentOddEven) rules.push(`@page :left { ${marginBoxes(hf.evenHeader, hf.evenFooter, fields)} }`);
  if (hf.differentFirst) rules.push(`@page :first { ${marginBoxes(hf.firstHeader, hf.firstFooter, fields)} }`);
  return rules.join('\n');
}

/** Build the printable HTML document for the active sheet. */
export function printHtml(ctl: EditorController): string {
  const doc = ctl.doc;
  const ws = doc.ws;
  const opts = ws.printOptions ?? {};
  const ps = ws.pageSetup ?? {};
  const m = ws.pageMargins ?? { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };
  const gridlines = opts.gridLines === true;
  const headings = opts.headings === true;
  const merges = doc.merges;
  const titles = printTitles(ctl);
  const tables: string[] = [];

  const rowHtml = (row: number, visibleCols: readonly number[]) => {
    const tds: string[] = [];
    for (const col of visibleCols) {
      const merge = merges.at(row, col);
      if (merge && (merge.r1 !== row || merge.c1 !== col)) continue;
      const cell = getCellAt(ws, row, col);
      const style = doc.styles.get(cell?.styleId ?? ctl.defaultStyleAt(row, col));
      const shown = cell && cell.value !== null ? displayCell(doc.wb, cell, style.numFmt) : undefined;
      let css = cellCss(style, gridlines);
      if (shown?.color) css += `;color:${shown.color}`;
      if (style.hAlign === 'general' && shown?.kind === 'number') css += ';text-align:right';
      if (style.hAlign === 'general' && (shown?.kind === 'bool' || shown?.kind === 'error')) css += ';text-align:center';
      const span = merge ? ` rowspan="${merge.r2 - merge.r1 + 1}" colspan="${visibleCols.filter((c) => c >= merge.c1 && c <= merge.c2).length}"` : '';
      // A date with no reading prints as #s; the page has no column width to fill exactly.
      const text = shown?.hashes ? PRINT_HASHES : (shown?.text ?? '');
      tds.push(`<td style="${escapeHtml(css)}"${span}>${escapeHtml(text)}</td>`);
    }
    const head = headings ? `<th>${row}</th>` : '';
    return `<tr style="height:${doc.rows.sizeOf(row)}px">${head}${tds.join('')}</tr>`;
  };

  for (const r of printRanges(ctl)) {
    // Title columns print to the left of every page's columns; a browser
    // repeats a table's <thead> on each printed page, which is where the
    // title rows go (taken out of the body when the range contains them).
    const visibleCols: number[] = [];
    if (titles.cols) for (let c = titles.cols[0]; c <= titles.cols[1]; c++) if (c < r.c1 && !doc.cols.isHidden(c)) visibleCols.push(c);
    for (let c = r.c1; c <= r.c2; c++) if (!doc.cols.isHidden(c)) visibleCols.push(c);
    const isTitleRow = (row: number) => titles.rows !== undefined && row >= titles.rows[0] && row <= titles.rows[1];
    const head: string[] = [];
    if (headings) head.push(`<tr><th></th>${visibleCols.map((c) => `<th>${colLetter(c)}</th>`).join('')}</tr>`);
    if (titles.rows) for (let row = titles.rows[0]; row <= titles.rows[1]; row++) if (!doc.rows.isHidden(row)) head.push(rowHtml(row, visibleCols));
    const body: string[] = [];
    for (let row = r.r1; row <= r.r2; row++) if (!doc.rows.isHidden(row) && !isTitleRow(row)) body.push(rowHtml(row, visibleCols));
    const colgroup = (headings ? '<col style="width:32px">' : '') + visibleCols.map((c) => `<col style="width:${doc.cols.sizeOf(c)}px">`).join('');
    tables.push(`<table><colgroup>${colgroup}</colgroup><thead>${head.join('')}</thead><tbody>${body.join('')}</tbody></table>`);
  }

  const now = new Date();
  const fields: HfFields = { date: now.toLocaleDateString(), time: now.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }), file: doc.fileName, path: '', sheet: ws.title };

  const scale = (ps.scale ?? 100) / 100;
  const orientation = ps.orientation === 'landscape' ? 'landscape' : 'portrait';
  const center = [opts.horizontalCentered ? 'margin-left:auto;margin-right:auto' : '', opts.verticalCentered ? 'margin-top:auto;margin-bottom:auto' : ''].filter(Boolean).join(';');
  return `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(ws.title)}</title><style>
@page { size: ${orientation}; margin: ${m.top}in ${m.right}in ${m.bottom}in ${m.left}in; }
${headerFooterCss(ws.headerFooter, fields)}
body { margin: 0; zoom: ${scale}; }
table { border-collapse: collapse; table-layout: fixed; ${center}; page-break-after: always; }
td { overflow: hidden; padding: 0 2px; font-size: 11pt; }
th { background: #f3f3f3; border: 1px solid #c8c8c8; font: 10px sans-serif; color: #444; }
</style></head><body>${tables.join('')}</body></html>`;
}

export function printWorkbook(ctl: EditorController): void {
  const frame = document.createElement('iframe');
  frame.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0';
  document.body.append(frame);
  const w = frame.contentWindow;
  if (!w) {
    frame.remove();
    return;
  }
  w.document.open();
  w.document.write(printHtml(ctl));
  w.document.close();
  w.addEventListener('afterprint', () => frame.remove());
  w.focus();
  w.print();
}
