import type { ApiModule } from './types';

export type SectionInput = {
  name: string;
  module: ApiModule;
  sourceFile: string;
};

export type SectionDef = {
  id: string;
  title: string;
  description: string;
  match: (input: SectionInput) => boolean;
};

const startsWith =
  (...prefixes: string[]) =>
  (file: string): boolean =>
    prefixes.some((p) => file.startsWith(p));

const fileEquals =
  (...files: string[]) =>
  (file: string): boolean =>
    files.includes(file);

// Order matters: the first matching section wins. Place narrower rules
// (specific filenames) before broader ones (whole directory).
export const SECTIONS: SectionDef[] = [
  {
    id: 'loading',
    title: 'Loading & saving',
    description: 'Open a workbook from any source, write it back to any sink.',
    match: ({ sourceFile }) => startsWith('src/io/')(sourceFile),
  },
  {
    id: 'streaming-io',
    title: 'Streaming readers & writers',
    description:
      'createWriteOnlyWorkbook, loadWorkbookStream, and the iter-based row APIs.',
    match: ({ sourceFile }) => startsWith('src/streaming/')(sourceFile),
  },
  {
    id: 'node-helpers',
    title: 'Node fs helpers',
    description: 'fromFile / toFile / fromBuffer / fromReadable / toWritable.',
    match: ({ module }) => module === 'node',
  },
  {
    id: 'errors',
    title: 'Errors',
    description: 'OpenXmlError and its subclasses, plus shared error options.',
    match: ({ sourceFile }) => sourceFile === 'src/utils/exceptions.ts',
  },
  {
    id: 'coordinates',
    title: 'Coordinates & ranges',
    description: 'A1 ↔ row/col conversion, range parsing, EMU and date helpers.',
    match: ({ sourceFile }) =>
      fileEquals(
        'src/utils/coordinate.ts',
        'src/utils/datetime.ts',
        'src/utils/units.ts',
      )(sourceFile) || sourceFile === 'src/worksheet/cell-range.ts',
  },
  {
    id: 'inference',
    title: 'Cell value inference',
    description: 'inferCellType, escape helpers, error code constants.',
    match: ({ sourceFile }) =>
      fileEquals(
        'src/utils/inference.ts',
        'src/utils/escape.ts',
        'src/utils/css.ts',
      )(sourceFile),
  },
  {
    id: 'cells',
    title: 'Cells & values',
    description: 'Cell shape, formula helpers, inline rich-text composition.',
    match: ({ sourceFile }) => startsWith('src/cell/')(sourceFile),
  },
  {
    id: 'tables',
    title: 'Tables & autoFilter',
    description: 'Excel Tables, table styles, and column-level autoFilter.',
    match: ({ sourceFile }) =>
      fileEquals(
        'src/worksheet/table.ts',
        'src/worksheet/auto-filter.ts',
      )(sourceFile),
  },
  {
    id: 'validation',
    title: 'Validation & formatting',
    description: 'Data validation rules and conditional formatting.',
    match: ({ sourceFile }) =>
      fileEquals(
        'src/worksheet/data-validations.ts',
        'src/worksheet/conditional-formatting.ts',
      )(sourceFile),
  },
  {
    id: 'comments-hyperlinks',
    title: 'Comments & hyperlinks',
    description: 'Legacy comments, threaded comments, and hyperlinks.',
    match: ({ sourceFile }) =>
      fileEquals(
        'src/worksheet/comments.ts',
        'src/worksheet/threaded-comments.ts',
        'src/worksheet/hyperlinks.ts',
      )(sourceFile),
  },
  {
    id: 'worksheet',
    title: 'Worksheet',
    description:
      'Worksheet shape and the cell / row / column helpers built on top of it.',
    match: ({ sourceFile }) => startsWith('src/worksheet/')(sourceFile),
  },
  {
    id: 'styles',
    title: 'Styles',
    description: 'Font, Fill, Border, Alignment, NumberFormat, Stylesheet.',
    match: ({ sourceFile }) => startsWith('src/styles/')(sourceFile),
  },
  {
    id: 'drawings',
    title: 'Drawings & images',
    description: 'Picture / chart anchors, image format detection.',
    match: ({ sourceFile }) => startsWith('src/drawing/')(sourceFile),
  },
  {
    id: 'charts',
    title: 'Charts',
    description: 'Legacy c: chart kinds plus the cx: chartex family.',
    match: ({ sourceFile }) => startsWith('src/chart/')(sourceFile),
  },
  {
    id: 'chartsheets',
    title: 'Chartsheets',
    description: 'Standalone chartsheet objects and their views.',
    match: ({ sourceFile }) => startsWith('src/chartsheet/')(sourceFile),
  },
  {
    id: 'workbook',
    title: 'Workbook',
    description:
      'Workbook root model: sheets, defined names, properties, calc settings.',
    match: ({ sourceFile }) => startsWith('src/workbook/')(sourceFile),
  },
  {
    id: 'formulas',
    title: 'Formulas',
    description: 'Formula tokens, evaluator scaffolding, name parsing.',
    match: ({ sourceFile }) => startsWith('src/formula/')(sourceFile),
  },
  {
    id: 'low-level',
    title: 'Low-level (XML / ZIP / schema / packaging)',
    description:
      'XML reader / writer, OPC packaging, schema, ZIP. Reach in only when wiring custom parts.',
    match: ({ sourceFile }) =>
      startsWith('src/xml/', 'src/zip/', 'src/schema/', 'src/packaging/')(sourceFile),
  },
  {
    id: 'misc',
    title: 'Other utilities',
    description: 'Anything else exported from the public surface.',
    match: () => true,
  },
];

export function classify(input: SectionInput): string {
  for (const s of SECTIONS) {
    if (s.match(input)) return s.id;
  }
  return 'misc';
}
