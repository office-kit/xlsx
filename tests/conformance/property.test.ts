// Property-based conformance: fast-check generates arbitrary workbook
// shapes and we use validateXlsx as the oracle. Any generated workbook
// that round-trips through workbookToBytes must produce a package that
// passes all three tiers; if not, fast-check shrinks the counterexample
// to the minimal failing case.
//
// Surfaces this hits: cell coordinate encoding (A1 → r="…"), shared-string
// dedup, mergeCells overlap detection, defined-name escaping, sheet-id /
// r:id pairing, dimension calculation, freeze-pane serialisation, and the
// stylesheet pool's behaviour under arbitrary cell-style assignment.

import fc from 'fast-check';
import { describe, it } from 'vitest';
import { workbookToBytes } from '../../src/io/save.js';
import { setCellBackgroundColor, setCellFont } from '../../src/styles/cell-style.js';
import { makeColor } from '../../src/styles/colors.js';
import { makeFont } from '../../src/styles/fonts.js';
import { addDefinedName } from '../../src/workbook/defined-names.js';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook.js';
import {
  setFreezePanes,
  hideColumn,
  hideRow,
  mergeCells,
  setCell,
  setColumnWidth,
  setRowHeight,
  type Worksheet,
} from '../../src/worksheet/worksheet.js';
import { validateXlsx } from './validate.js';

// ---------------------------------------------------------------------------
// Generators
// ---------------------------------------------------------------------------

// ASCII printable text only — Excel allows much more, but the property test
// targets coordinate / structure invariants, not Unicode escaping (covered
// elsewhere). We do include the canonical XML-special characters.
const cellString = fc.string({ minLength: 0, maxLength: 12 }).filter(
  (s) => /^[ -~]*$/.test(s) && !s.includes(''),
);

// Numbers Excel accepts: finite IEEE-754 doubles. Skip NaN / ±Infinity since
// xlsx does not encode them as numeric cells.
const cellNumber = fc.double({
  noNaN: true,
  noDefaultInfinity: true,
  min: -1e15,
  max: 1e15,
});

const cellValue = fc.oneof(
  cellNumber,
  cellString,
  fc.boolean(),
  fc.integer({ min: -1_000_000, max: 1_000_000 }),
);

interface CellSpec {
  row: number;
  col: number;
  value: number | string | boolean;
  style?: { bold?: boolean; size?: number; bg?: string } | undefined;
}

const styleSpec = fc.record(
  {
    bold: fc.boolean(),
    size: fc.integer({ min: 8, max: 18 }),
    bg: fc.constantFrom('FFFFFF00', 'FFFF0000', 'FF00FF00', 'FF0000FF'),
  },
  { requiredKeys: [] },
);

const cellSpec: fc.Arbitrary<CellSpec> = fc.record(
  {
    row: fc.integer({ min: 1, max: 50 }),
    col: fc.integer({ min: 1, max: 26 }),
    value: cellValue,
    style: fc.option(styleSpec, { nil: undefined }),
  },
  { requiredKeys: ['row', 'col', 'value'] },
);

interface MergeSpec {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

const mergeSpec: fc.Arbitrary<MergeSpec> = fc
  .tuple(
    fc.integer({ min: 51, max: 60 }),
    fc.integer({ min: 1, max: 5 }),
    fc.integer({ min: 0, max: 4 }),
    fc.integer({ min: 0, max: 4 }),
  )
  .map(([r1, c1, dr, dc]) => ({ r1, c1, r2: r1 + dr, c2: c1 + dc }));

interface SheetSpec {
  cells: CellSpec[];
  merges: MergeSpec[];
  freezeRows?: number | undefined;
  freezeCols?: number | undefined;
  hiddenRow?: number | undefined;
  hiddenCol?: number | undefined;
  rowHeight?: { row: number; h: number } | undefined;
  colWidth?: { col: number; w: number } | undefined;
}

const sheetSpec: fc.Arbitrary<SheetSpec> = fc.record(
  {
    cells: fc.array(cellSpec, { minLength: 0, maxLength: 30 }),
    merges: fc.array(mergeSpec, { minLength: 0, maxLength: 4 }),
    freezeRows: fc.option(fc.integer({ min: 1, max: 3 }), { nil: undefined }),
    freezeCols: fc.option(fc.integer({ min: 1, max: 3 }), { nil: undefined }),
    hiddenRow: fc.option(fc.integer({ min: 100, max: 110 }), { nil: undefined }),
    hiddenCol: fc.option(fc.integer({ min: 10, max: 14 }), { nil: undefined }),
    rowHeight: fc.option(
      fc.record({ row: fc.integer({ min: 1, max: 50 }), h: fc.integer({ min: 6, max: 100 }) }),
      { nil: undefined },
    ),
    colWidth: fc.option(
      fc.record({ col: fc.integer({ min: 1, max: 20 }), w: fc.integer({ min: 4, max: 80 }) }),
      { nil: undefined },
    ),
  },
  { requiredKeys: ['cells', 'merges'] },
);

interface WorkbookSpec {
  sheets: SheetSpec[];
  definedNames: Array<{ name: string; value: string }>;
}

const definedNameSpec = fc.record({
  name: fc.constantFrom('Alpha', 'Beta', 'Gamma', 'Delta', 'Epsilon'),
  value: fc.constantFrom('1', '2.5', "'Sheet1'!$A$1", "'Sheet1'!$A$1:$B$2"),
});

const workbookSpec: fc.Arbitrary<WorkbookSpec> = fc.record({
  sheets: fc.array(sheetSpec, { minLength: 1, maxLength: 4 }),
  definedNames: fc.array(definedNameSpec, { minLength: 0, maxLength: 3 }),
});

// ---------------------------------------------------------------------------
// Materialisation: WorkbookSpec → real workbook bytes, with collision filters
// so the spec stays "valid by construction" (mergeCell overlap and duplicate
// cell coordinates would be Tier-C / OPC violations the writer cannot emit).
// ---------------------------------------------------------------------------

const overlaps = (a: MergeSpec, b: MergeSpec): boolean =>
  a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;

const dedupeMerges = (m: MergeSpec[]): MergeSpec[] => {
  const kept: MergeSpec[] = [];
  for (const candidate of m) {
    if (kept.some((k) => overlaps(k, candidate))) continue;
    kept.push(candidate);
  }
  return kept;
};

const dedupeCells = (cells: CellSpec[]): CellSpec[] => {
  const seen = new Set<string>();
  const out: CellSpec[] = [];
  for (const c of cells) {
    const k = `${c.row},${c.col}`;
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(c);
  }
  return out;
};

const dedupeNames = (
  names: ReadonlyArray<{ name: string; value: string }>,
): Array<{ name: string; value: string }> => {
  const seen = new Set<string>();
  const out: Array<{ name: string; value: string }> = [];
  for (const n of names) {
    if (seen.has(n.name)) continue;
    seen.add(n.name);
    out.push(n);
  }
  return out;
};

const buildWorkbook = (spec: WorkbookSpec) => {
  const wb = createWorkbook();
  spec.sheets.forEach((sheet, i) => {
    const ws = addWorksheet(wb, `Sheet${i + 1}`) as Worksheet;
    for (const cell of dedupeCells(sheet.cells)) {
      const c = setCell(ws, cell.row, cell.col, cell.value);
      if (cell.style?.bold !== undefined || cell.style?.size !== undefined) {
        setCellFont(
          wb,
          c,
          makeFont({
            ...(cell.style.bold !== undefined ? { bold: cell.style.bold } : {}),
            ...(cell.style.size !== undefined ? { size: cell.style.size } : {}),
          }),
        );
      }
      if (cell.style?.bg) {
        setCellBackgroundColor(wb, c, makeColor({ rgb: cell.style.bg }));
      }
    }
    for (const m of dedupeMerges(sheet.merges)) {
      mergeCells(ws, `${col(m.c1)}${m.r1}:${col(m.c2)}${m.r2}`);
    }
    if (sheet.freezeRows && sheet.freezeCols)
      setFreezePanes(ws, { rows: sheet.freezeRows, cols: sheet.freezeCols });
    if (sheet.hiddenRow) hideRow(ws, sheet.hiddenRow);
    if (sheet.hiddenCol) hideColumn(ws, sheet.hiddenCol);
    if (sheet.rowHeight) setRowHeight(ws, sheet.rowHeight.row, sheet.rowHeight.h);
    if (sheet.colWidth) setColumnWidth(ws, sheet.colWidth.col, sheet.colWidth.w);
  });
  for (const dn of dedupeNames(spec.definedNames)) {
    addDefinedName(wb, dn);
  }
  return wb;
};

const col = (c: number): string => {
  let n = c;
  let out = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    out = String.fromCharCode(65 + r) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
};

// ---------------------------------------------------------------------------
// Property
// ---------------------------------------------------------------------------

describe('conformance: property-based oracle', () => {
  it('every well-formed WorkbookSpec produces a schema-clean xlsx', async () => {
    // CI runs more iterations than local — see vitest.config.ts; the env hook
    // here keeps local turnaround fast while still letting CI exercise depth.
    const numRuns = process.env['CI'] ? 50 : 25;
    await fc.assert(
      fc.asyncProperty(workbookSpec, async (spec) => {
        const wb = buildWorkbook(spec);
        const bytes = await workbookToBytes(wb);
        const result = await validateXlsx(bytes);
        if (!result.ok) {
          // fast-check displays this on shrink output, so we leave the rich
          // diagnostic in the assertion message rather than throwing plain.
          throw new Error(
            `validation failed:\n${result.issues
              .map((i) => `  [${i.tier}] ${i.part}: ${i.message}`)
              .join('\n')}`,
          );
        }
      }),
      { numRuns, verbose: false },
    );
  }, 120_000);
});
