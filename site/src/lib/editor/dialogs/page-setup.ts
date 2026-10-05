// Page Setup dialog model: the worksheet's print settings as editable form
// fields, and writing them back (plus the Print_Area / Print_Titles names) as
// one undo step.

import type { CellCommentMode, HeaderFooter, PageOrder, PageOrientation } from '@office-kit/xlsx/worksheet';
import { isWholeColumns, isWholeRows, parseRangeAddress, quoteSheetName, rangeAddress } from '../core/address.ts';
import type { EditorController } from '../core/controller.svelte.ts';
import { buildHeader, parseHeader, type HeaderParts } from '../core/header-footer.ts';
import type { MessageKey } from '../i18n/i18n.svelte.ts';

const PRINT_AREA = '_xlnm.Print_Area';
const PRINT_TITLES = '_xlnm.Print_Titles';

/** Paper sizes Excel offers first, by their OOXML paperSize code (ECMA-376 §18.3.1.63). */
export const PAPER_SIZES: ReadonlyArray<{ code: number; label: string }> = [
  { code: 1, label: 'US Letter (8.5 × 11 in)' },
  { code: 5, label: 'US Legal (8.5 × 14 in)' },
  { code: 7, label: 'Executive (7.25 × 10.5 in)' },
  { code: 3, label: 'Tabloid (11 × 17 in)' },
  { code: 8, label: 'A3 (297 × 420 mm)' },
  { code: 9, label: 'A4 (210 × 297 mm)' },
  { code: 11, label: 'A5 (148 × 210 mm)' },
  { code: 12, label: 'B4 (JIS) (257 × 364 mm)' },
  { code: 13, label: 'B5 (JIS) (182 × 257 mm)' },
  { code: 20, label: 'Envelope #10' },
  { code: 27, label: 'Envelope DL' },
];

/** Excel's Normal margins, used when the sheet has none stored. */
const NORMAL_MARGINS = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 };

export type HfPage = 'odd' | 'first' | 'even';

export interface PageForm {
  orientation: Exclude<PageOrientation, 'default'>;
  fit: boolean;
  scale: number;
  /** Pages wide / tall; '' is "automatic" (no limit). */
  fitWide: string;
  fitTall: string;
  paperSize: number;
  firstPage: string;
  margins: { top: number; bottom: number; left: number; right: number; header: number; footer: number };
  centerH: boolean;
  centerV: boolean;
  /** Header / footer per page kind; first and even apply only when their flag is on. */
  headers: Record<HfPage, { header: HeaderParts; footer: HeaderParts }>;
  differentFirst: boolean;
  differentOddEven: boolean;
  scaleWithDoc: boolean;
  alignWithMargins: boolean;
  printArea: string;
  titleRows: string;
  titleCols: string;
  gridLines: boolean;
  headings: boolean;
  blackAndWhite: boolean;
  draft: boolean;
  comments: CellCommentMode;
  pageOrder: PageOrder;
}

function stripSheet(ref: string): string {
  const bang = ref.lastIndexOf('!');
  return bang >= 0 ? ref.slice(bang + 1) : ref;
}

export function pageForm(ctl: EditorController): PageForm {
  const doc = ctl.doc;
  const ws = doc.ws;
  const ps = ws.pageSetup ?? {};
  const scope = doc.activeSheetIndex;
  const named = (name: string) => doc.wb.definedNames.find((d) => d.name === name && d.scope === scope)?.value;
  const titles = (named(PRINT_TITLES) ?? '').split(',').map(stripSheet).filter((s) => s !== '');
  const m = ws.pageMargins ?? NORMAL_MARGINS;
  const hf = ws.headerFooter ?? {};
  return {
    orientation: ps.orientation === 'landscape' ? 'landscape' : 'portrait',
    fit: ws.sheetProperties?.pageSetUpPr?.fitToPage === true,
    scale: ps.scale ?? 100,
    fitWide: ps.fitToWidth === 0 ? '' : String(ps.fitToWidth ?? 1),
    fitTall: ps.fitToHeight === 0 ? '' : String(ps.fitToHeight ?? 1),
    paperSize: ps.paperSize ?? (ctl.dateOrder() === 'ymd' ? 9 : 1),
    firstPage: ps.useFirstPageNumber && ps.firstPageNumber !== undefined ? String(ps.firstPageNumber) : '',
    margins: { top: m.top, bottom: m.bottom, left: m.left, right: m.right, header: m.header, footer: m.footer },
    centerH: ws.printOptions?.horizontalCentered === true,
    centerV: ws.printOptions?.verticalCentered === true,
    headers: {
      odd: { header: parseHeader(hf.oddHeader), footer: parseHeader(hf.oddFooter) },
      first: { header: parseHeader(hf.firstHeader), footer: parseHeader(hf.firstFooter) },
      even: { header: parseHeader(hf.evenHeader), footer: parseHeader(hf.evenFooter) },
    },
    differentFirst: hf.differentFirst === true,
    differentOddEven: hf.differentOddEven === true,
    scaleWithDoc: hf.scaleWithDoc !== false,
    alignWithMargins: hf.alignWithMargins !== false,
    printArea: (named(PRINT_AREA) ?? '').split(',').map(stripSheet).join(','),
    titleRows: titles.find((s) => /^\$?\d+:\$?\d+$/.test(s)) ?? '',
    titleCols: titles.find((s) => /^\$?[A-Za-z]+:\$?[A-Za-z]+$/.test(s)) ?? '',
    gridLines: ws.printOptions?.gridLines === true,
    headings: ws.printOptions?.headings === true,
    blackAndWhite: ps.blackAndWhite === true,
    draft: ps.draft === true,
    comments: ps.cellComments ?? 'none',
    pageOrder: ps.pageOrder ?? 'downThenOver',
  };
}

const FIT_MAX = 32767;

function whole(text: string, min: number, max: number): number | undefined {
  const n = Number(text.trim());
  return text.trim() !== '' && Number.isInteger(n) && n >= min && n <= max ? n : undefined;
}

/** Validate and write the form. Returns the message to show when a field is invalid. */
export function applyPageForm(ctl: EditorController, f: PageForm): MessageKey | undefined {
  const doc = ctl.doc;
  const ws = doc.ws;
  const scope = doc.activeSheetIndex;
  const sheetPrefix = `${quoteSheetName(ws.title)}!`;

  if (!f.fit && (!Number.isInteger(f.scale) || f.scale < 10 || f.scale > 400)) return 'dlgPsScaleRange';
  const wide = f.fitWide.trim() === '' ? 0 : whole(f.fitWide, 1, FIT_MAX);
  const tall = f.fitTall.trim() === '' ? 0 : whole(f.fitTall, 1, FIT_MAX);
  if (f.fit && (wide === undefined || tall === undefined)) return 'dlgPsFitRange';
  const firstPage = f.firstPage.trim() === '' ? undefined : whole(f.firstPage, -32765, 32767);
  if (f.firstPage.trim() !== '' && firstPage === undefined) return 'dlgPsFirstPage';
  if (Object.values(f.margins).some((v) => !Number.isFinite(v) || v < 0 || v > 49)) return 'dlgPsMarginRange';

  const areas: string[] = [];
  for (const part of f.printArea.split(',').map((s) => s.trim()).filter((s) => s !== '')) {
    const parsed = parseRangeAddress(part.replace(/^=/, ''));
    if (!parsed) return 'invalidReference';
    areas.push(sheetPrefix + rangeAddress(parsed.range, true));
  }
  const titles: string[] = [];
  if (f.titleRows.trim()) {
    const parsed = parseRangeAddress(f.titleRows.trim());
    if (!parsed || !isWholeRows(parsed.range)) return 'dlgPsTitleRows';
    titles.push(sheetPrefix + rangeAddress(parsed.range, true));
  }
  if (f.titleCols.trim()) {
    const parsed = parseRangeAddress(f.titleCols.trim());
    if (!parsed || !isWholeColumns(parsed.range)) return 'dlgPsTitleCols';
    titles.push(sheetPrefix + rangeAddress(parsed.range, true));
  }

  doc.transact('Page Setup', (tx) => {
    tx.sheet(ws, 'pageSetup', 'pageMargins', 'printOptions', 'headerFooter', 'sheetProperties');
    tx.workbook('definedNames');
    const ps = { ...ws.pageSetup };
    ps.orientation = f.orientation;
    ps.paperSize = f.paperSize;
    if (f.fit) {
      ps.fitToWidth = wide ?? 1;
      ps.fitToHeight = tall ?? 1;
    } else ps.scale = f.scale;
    if (firstPage === undefined) {
      delete ps.firstPageNumber;
      delete ps.useFirstPageNumber;
    } else {
      ps.firstPageNumber = firstPage;
      ps.useFirstPageNumber = true;
    }
    ps.blackAndWhite = f.blackAndWhite;
    ps.draft = f.draft;
    ps.cellComments = f.comments;
    ps.pageOrder = f.pageOrder;
    ws.pageSetup = ps;
    ws.sheetProperties = { ...ws.sheetProperties, pageSetUpPr: { ...ws.sheetProperties?.pageSetUpPr, fitToPage: f.fit } };
    ws.pageMargins = { ...f.margins };
    ws.printOptions = { ...ws.printOptions, horizontalCentered: f.centerH, verticalCentered: f.centerV, gridLines: f.gridLines, headings: f.headings };
    ws.headerFooter = headerFooterOf(f);
    const others = doc.wb.definedNames.filter((d) => !(d.scope === scope && (d.name === PRINT_AREA || d.name === PRINT_TITLES)));
    doc.wb.definedNames = [
      ...others,
      ...(areas.length > 0 ? [{ name: PRINT_AREA, value: areas.join(','), scope }] : []),
      ...(titles.length > 0 ? [{ name: PRINT_TITLES, value: titles.join(','), scope }] : []),
    ];
  });
  return undefined;
}

function headerFooterOf(f: PageForm): HeaderFooter {
  const hf: HeaderFooter = {};
  // The flags are written only when they differ from Excel's defaults, so an
  // untouched sheet keeps an empty <headerFooter>.
  if (f.differentFirst) hf.differentFirst = true;
  if (f.differentOddEven) hf.differentOddEven = true;
  if (!f.scaleWithDoc) hf.scaleWithDoc = false;
  if (!f.alignWithMargins) hf.alignWithMargins = false;
  const put = (page: HfPage, header: 'oddHeader' | 'firstHeader' | 'evenHeader', footer: 'oddFooter' | 'firstFooter' | 'evenFooter') => {
    const h = buildHeader(f.headers[page].header);
    const ft = buildHeader(f.headers[page].footer);
    if (h) hf[header] = h;
    if (ft) hf[footer] = ft;
  };
  put('odd', 'oddHeader', 'oddFooter');
  if (f.differentFirst) put('first', 'firstHeader', 'firstFooter');
  if (f.differentOddEven) put('even', 'evenHeader', 'evenFooter');
  return hf;
}
