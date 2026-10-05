// Excel Table object (xl/tables/tableN.xml).
//
// Tables ride on top of a worksheet range, give it a name + structured column
// references, and own their own AutoFilter. Each table sits in a separate part
// — the worksheet only carries a `<tableParts>` block pointing at the
// workbook-rels rId. Stage-1 covers the table shell + columns + styleInfo +
// autoFilter + sortState; xml extlst is reserved for later.

import type { Workbook } from '../workbook/workbook.js';
import type { AutoFilter } from './auto-filter.js';
import type { SortState } from './sort-state.js';
import { addTable, type Worksheet } from './worksheet.js';

export interface TableColumn {
  /** 1-based column id (per-table). */
  id: number;
  /** Header name. */
  name: string;
  /** Totals-row aggregation function. */
  totalsRowFunction?: 'sum' | 'min' | 'max' | 'count' | 'countNums' | 'average' | 'stdDev' | 'var' | 'custom';
  /** Override label for the totals row. */
  totalsRowLabel?: string;
  /** Custom totals-row formula text. */
  totalsRowFormula?: string;
  /** Calculated-column formula. */
  calculatedColumnFormula?: string;
}

export interface TableStyleInfo {
  /** Built-in style name (TableStyleMedium2, etc) or custom. */
  name?: string;
  showFirstColumn?: boolean;
  showLastColumn?: boolean;
  showRowStripes?: boolean;
  showColumnStripes?: boolean;
}

export interface TableDefinition {
  /** Workbook-unique id (`<table id="N">`). */
  id: number;
  /** Workbook-unique displayName — Excel surfaces this in formulas. */
  displayName: string;
  /** Optional friendly name; usually matches `displayName`. */
  name?: string;
  /** Range covered by the table, e.g. "A1:E10". */
  ref: string;
  /** Number of header rows. Defaults to 1; 0 means a header-less table. */
  headerRowCount?: number;
  /** Number of totals rows. */
  totalsRowCount?: number;
  /** Whether the totals row is currently visible. */
  totalsRowShown?: boolean;
  styleInfo?: TableStyleInfo;
  columns: TableColumn[];
  autoFilter?: AutoFilter;
  /** The table's last applied sort. */
  sortState?: SortState;
  /** Worksheet-rels rId — populated on read; the writer assigns its own. */
  rId?: string;
}

export function makeTableColumn(opts: { id: number; name: string }): TableColumn {
  return { id: opts.id, name: opts.name };
}

export function makeTableDefinition(opts: {
  id: number;
  displayName: string;
  ref: string;
  name?: string;
  columns?: TableColumn[];
  headerRowCount?: number;
  totalsRowCount?: number;
  totalsRowShown?: boolean;
  styleInfo?: TableStyleInfo;
  autoFilter?: AutoFilter;
}): TableDefinition {
  return {
    id: opts.id,
    displayName: opts.displayName,
    ref: opts.ref,
    columns: opts.columns ?? [],
    ...(opts.name !== undefined ? { name: opts.name } : {}),
    ...(opts.headerRowCount !== undefined ? { headerRowCount: opts.headerRowCount } : {}),
    ...(opts.totalsRowCount !== undefined ? { totalsRowCount: opts.totalsRowCount } : {}),
    ...(opts.totalsRowShown !== undefined ? { totalsRowShown: opts.totalsRowShown } : {}),
    ...(opts.styleInfo ? { styleInfo: opts.styleInfo } : {}),
    ...(opts.autoFilter ? { autoFilter: opts.autoFilter } : {}),
  };
}

const nextTableId = (wb: Workbook): number => {
  let max = 0;
  for (const ref of wb.sheets) {
    if (ref.kind !== 'worksheet') continue;
    for (const t of ref.sheet.tables) {
      if (t.id > max) max = t.id;
    }
  }
  return max + 1;
};

/**
 * High-level wrapper that builds a TableDefinition + pushes it onto `ws.tables`
 * in one call. Auto-assigns the workbook-unique `id`, derives `displayName`
 * from the supplied `name`, and constructs `TableColumn` records (1-based ids)
 * from a string-array shorthand.
 *
 * Goes through {@link addTable}, so the definition is checked against the sheet
 * before it lands. Pass `headerRowCount: 0` for a genuinely header-less table.
 */
export const addExcelTable = (
  wb: Workbook,
  ws: Worksheet,
  opts: {
    name: string;
    ref: string;
    columns: ReadonlyArray<string | TableColumn>;
    style?: string;
    styleInfo?: TableStyleInfo;
    headerRowCount?: number;
    totalsRowCount?: number;
    totalsRowShown?: boolean;
    autoFilter?: AutoFilter;
    displayName?: string;
  },
): TableDefinition => {
  const cols: TableColumn[] = opts.columns.map((c, i): TableColumn =>
    typeof c === 'string' ? { id: i + 1, name: c } : c,
  );
  const styleInfo: TableStyleInfo | undefined =
    opts.styleInfo ??
    (opts.style !== undefined ? { name: opts.style, showRowStripes: true, showColumnStripes: false } : undefined);
  const def: TableDefinition = makeTableDefinition({
    id: nextTableId(wb),
    displayName: opts.displayName ?? opts.name,
    name: opts.name,
    ref: opts.ref,
    columns: cols,
    ...(opts.headerRowCount !== undefined ? { headerRowCount: opts.headerRowCount } : {}),
    ...(opts.totalsRowCount !== undefined ? { totalsRowCount: opts.totalsRowCount } : {}),
    ...(opts.totalsRowShown !== undefined ? { totalsRowShown: opts.totalsRowShown } : {}),
    ...(styleInfo ? { styleInfo } : {}),
    ...(opts.autoFilter ? { autoFilter: opts.autoFilter } : {}),
  });
  return addTable(ws, def);
};
