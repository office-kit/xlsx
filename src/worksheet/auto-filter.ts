// AutoFilter.
//
// Value lists, custom conditions, Top 10 and dynamic filters are modelled;
// any other criterion (colour, icon) is kept as its XML so it survives a
// round-trip.

import { OpenXmlSchemaError } from '../utils/exceptions.js';
import type { SortState } from './sort-state.js';

/** One condition of a custom filter; `val` may hold `*` / `?` wildcards. */
export interface CustomFilterCondition {
  /** Comparison against `val`. Defaults to `equal`. */
  operator?: 'equal' | 'lessThan' | 'lessThanOrEqual' | 'notEqual' | 'greaterThanOrEqual' | 'greaterThan';
  val: string;
}

export type FilterColumn =
  | {
      kind: 'filters';
      colId: number;
      /** Discrete values that pass the filter. Stored as strings to match the wire format. */
      values: string[];
      /** Whether blanks are visible. */
      blank?: boolean;
    }
  | {
      /** Text / Number Filters ▸ Custom: one or two conditions. */
      kind: 'custom';
      colId: number;
      /** Both conditions must hold (`true`) or either one (default). */
      and?: boolean;
      conditions: CustomFilterCondition[];
    }
  | {
      /** Top 10: the `val` largest (or smallest) items, or percent. */
      kind: 'top10';
      colId: number;
      /** Smallest instead of largest when `false`. Defaults to `true`. */
      top?: boolean;
      percent?: boolean;
      val: number;
      /** The cut-off value Excel computed when it applied the filter. */
      filterVal?: number;
    }
  | {
      /** Dynamic filter such as `aboveAverage`, `today` or `Q1`; `val` holds the average it compared with. */
      kind: 'dynamic';
      colId: number;
      type: string;
      val?: number;
      maxVal?: number;
    }
  | {
      /** A criterion not modelled here (colour, icon), kept as the `<filterColumn>`'s inner XML. */
      kind: 'raw';
      colId: number;
      xml: string;
    };

export interface AutoFilter {
  /** Excel range the filter covers (`"A1:E100"`). */
  ref: string;
  filterColumns: FilterColumn[];
  /** The sort applied from the filter buttons; Excel nests it in `<autoFilter>`. */
  sortState?: SortState;
}

export function makeAutoFilter(opts: { ref: string; filterColumns?: FilterColumn[] }): AutoFilter {
  return { ref: opts.ref, filterColumns: opts.filterColumns ?? [] };
}

export function makeFilterColumn(opts: {
  colId: number;
  values: ReadonlyArray<string>;
  blank?: boolean;
}): Extract<FilterColumn, { kind: 'filters' }> {
  return {
    kind: 'filters',
    colId: opts.colId,
    values: [...opts.values],
    ...(opts.blank !== undefined ? { blank: opts.blank } : {}),
  };
}

// ---- Worksheet ergonomic builders ---------------------------------------

import type { Worksheet } from './worksheet.js';

/** Add an AutoFilter dropdown header strip to the given range. */
export const addAutoFilter = (ws: Worksheet, ref: string): AutoFilter => {
  ws.autoFilter = makeAutoFilter({ ref });
  return ws.autoFilter;
};

/**
 * Add a value-list dropdown filter to a column inside the existing AutoFilter
 * range. `colId` is 0-based relative to the AutoFilter left edge.
 */
export const addAutoFilterColumn = (
  ws: Worksheet,
  colId: number,
  values: ReadonlyArray<string>,
  opts: { blank?: boolean } = {},
): Extract<FilterColumn, { kind: 'filters' }> => {
  if (!ws.autoFilter) {
    throw new OpenXmlSchemaError('addAutoFilterColumn: call addAutoFilter(ws, ref) first');
  }
  const fc = makeFilterColumn({ colId, values, ...(opts.blank !== undefined ? { blank: opts.blank } : {}) });
  ws.autoFilter.filterColumns.push(fc);
  return fc;
};

/** Drop the worksheet's AutoFilter entirely. */
export const removeAutoFilter = (ws: Worksheet): void => {
  delete (ws as { autoFilter?: AutoFilter }).autoFilter;
};
