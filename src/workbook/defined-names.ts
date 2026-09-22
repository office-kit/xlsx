// Workbook-level defined names. Models the OOXML schema's `<definedName>` element.
//
// A defined name binds an identifier to a formula-style value (`'Sheet
// 1'!$A$1:$B$10`, `SUM(A:A)`, etc). Workbook-scope names omit `localSheetId`;
// sheet-scope names use the 0-based sheet index. Excel reserves a handful of
// names with the `_xlnm.` prefix for built-in uses (Print_Area, Print_Titles,
// Sheet_Title, etc) — those round-trip here as plain DefinedName entries since
// the value semantics are the same.

import { OpenXmlSchemaError } from '../utils/exceptions.js';
import { normalizeFormulaText } from '../utils/formula-text.js';

export interface DefinedName {
  /** Identifier — `_xlnm.Print_Area` for built-ins, otherwise user-chosen. */
  name: string;
  /** The formula expression the name points at. */
  value: string;
  /** 0-based sheet index for sheet-scope names; undefined → workbook-scope. */
  scope?: number;
  /** Hidden from the Name Manager when true. */
  hidden?: boolean;
  /** Optional human-readable description. */
  comment?: string;
}

export function makeDefinedName(opts: Partial<DefinedName> & { name: string; value: string }): DefinedName {
  return {
    name: opts.name,
    value: normalizeFormulaText(opts.value),
    ...(opts.scope !== undefined ? { scope: opts.scope } : {}),
    ...(opts.hidden !== undefined ? { hidden: opts.hidden } : {}),
    ...(opts.comment !== undefined ? { comment: opts.comment } : {}),
  };
}

// ---- Workbook ergonomic helpers -----------------------------------------

import { type CellRangeBoundaries, parseSheetRange } from '../utils/coordinate.js';
import type { Worksheet } from '../worksheet/worksheet.js';
import { getRangeAddress } from '../worksheet/worksheet.js';
import type { Workbook } from './workbook.js';

/**
 * One parsed leg of a defined name's value. Defined-name values can be
 * comma-separated multi-range expressions (e.g. `_xlnm.Print_Titles` sets
 * `Sheet!$1:$1,Sheet!$A:$A`); this represents one such leg.
 */
export interface DefinedNameTarget {
  sheet: string;
  range: string;
  bounds: CellRangeBoundaries;
}

/**
 * Add a workbook-scope or sheet-scope defined name. If a defined name with the
 * same `name` (and `scope`) already exists, it's replaced — Excel allows one
 * workbook-scope and one per-sheet-scope name, but not two with the same scope.
 * Returns the resulting `DefinedName`.
 */
export const addDefinedName = (
  wb: Workbook,
  opts: Partial<DefinedName> & { name: string; value: string },
): DefinedName => {
  const dn = makeDefinedName(opts);
  // Replace any existing entry with the same name + scope.
  const idx = wb.definedNames.findIndex((d) => d.name === dn.name && d.scope === dn.scope);
  if (idx >= 0) {
    wb.definedNames[idx] = dn;
  } else {
    wb.definedNames.push(dn);
  }
  return dn;
};

/**
 * High-level: register a defined name pointing at a worksheet range. Combines
 * {@link getRangeAddress} (sheet-qualified, properly quoted) with {@link
 * addDefinedName}, so the caller doesn't have to assemble the formula string by
 * hand.
 *
 * Pass `opts.localToSheet: true` to scope the name to the worksheet (instead of
 * the workbook). Re-using the same `name` + scope replaces the previous entry
 * (Excel's per-scope-uniqueness rule).
 *
 * Throws when `localToSheet: true` is set but the worksheet isn't on
 * `wb.sheets` — that would be a stale Worksheet reference.
 */
export const addDefinedNameForRange = (
  wb: Workbook,
  name: string,
  ws: Worksheet,
  range: string,
  opts: { localToSheet?: boolean; hidden?: boolean; comment?: string } = {},
): DefinedName => {
  const value = getRangeAddress(ws, range);
  let scope: number | undefined;
  if (opts.localToSheet) {
    const idx = wb.sheets.findIndex((s) => s.sheet === ws);
    if (idx < 0) {
      throw new OpenXmlSchemaError(
        `addDefinedNameForRange: worksheet "${ws.title}" is not registered on this workbook`,
      );
    }
    scope = idx;
  }
  return addDefinedName(wb, {
    name,
    value,
    ...(scope !== undefined ? { scope } : {}),
    ...(opts.hidden !== undefined ? { hidden: opts.hidden } : {}),
    ...(opts.comment !== undefined ? { comment: opts.comment } : {}),
  });
};

/** Look up a defined name by identifier and (optional) sheet scope. */
export const getDefinedName = (
  wb: Workbook,
  name: string,
  scope?: number,
): DefinedName | undefined => wb.definedNames.find((d) => d.name === name && d.scope === scope);

/**
 * Resolve a defined name's `value` into one or more {@link DefinedNameTarget}s.
 * Comma-separated values (e.g. `_xlnm.Print_Titles` typically sets
 * `Sheet!$1:$1,Sheet!$A:$A`) yield one entry per leg; a plain `Sheet!A1:B5`
 * yields a single-element array.
 *
 * Returns `undefined` when the name doesn't exist; throws when the value can't
 * be parsed (e.g. a constant or a non-range formula — defined names are
 * sometimes used for things like `=42` or `=SUM(A:A)` which aren't ranges).
 */
export const getDefinedNameTarget = (
  wb: Workbook,
  name: string,
  scope?: number,
): DefinedNameTarget[] | undefined => {
  const dn = getDefinedName(wb, name, scope);
  if (!dn) return undefined;
  // Defined-name values use `,` as the leg separator. Sheet titles can
  // themselves contain commas inside `'...'` quotes — split on commas that
  // aren't inside an unbalanced single-quoted segment.
  const legs: string[] = [];
  let current = '';
  let inQuote = false;
  for (let i = 0; i < dn.value.length; i++) {
    const c = dn.value[i];
    if (c === "'") {
      // Doubled `''` inside a quoted run is the escape for a literal apostrophe
      // — skip the second one without flipping the state.
      if (inQuote && dn.value[i + 1] === "'") {
        current += "''";
        i++;
        continue;
      }
      inQuote = !inQuote;
      current += c;
      continue;
    }
    if (c === ',' && !inQuote) {
      legs.push(current);
      current = '';
      continue;
    }
    current += c;
  }
  if (current.length > 0) legs.push(current);
  return legs.map((leg) => parseSheetRange(leg));
};

/**
 * Remove a defined name by identifier + scope. Returns true if any entry was
 * removed.
 */
export const removeDefinedName = (wb: Workbook, name: string, scope?: number): boolean => {
  const idx = wb.definedNames.findIndex((d) => d.name === name && d.scope === scope);
  if (idx < 0) return false;
  wb.definedNames.splice(idx, 1);
  return true;
};

/**
 * Every defined name. Pass `{ scope }` to narrow to workbook-scope
 * (`scope: 'workbook'`) or one specific sheet (`scope: 0`); omit the option, or
 * pass `'all'`, to list every name. `'workbook'` is the only spelling that
 * narrows to the unscoped ones: the option is `number | 'workbook' | 'all'`, so
 * there is no `undefined` to pass for them.
 *
 * A narrowed call filters, so it hands back a fresh array. Listing all hands
 * back a read-only view of the live one, which `addDefinedName` and
 * `renameDefinedName` are visible through and which `removeDefinedNames`
 * replaces outright.
 */
export const listDefinedNames = (
  wb: Workbook,
  opts: { scope?: number | 'workbook' | 'all' } = {},
): ReadonlyArray<DefinedName> => {
  const scope = opts.scope ?? 'all';
  if (scope === 'all') return wb.definedNames;
  if (scope === 'workbook') return wb.definedNames.filter((d) => d.scope === undefined);
  return wb.definedNames.filter((d) => d.scope === scope);
};

/**
 * Bulk-remove every defined name matching `predicate`. Returns the count
 * removed. Mirrors {@link removeDataValidations} on worksheets.
 */
export const removeDefinedNames = (
  wb: Workbook,
  predicate: (d: DefinedName) => boolean,
): number => {
  const before = wb.definedNames.length;
  wb.definedNames = wb.definedNames.filter((d) => !predicate(d));
  return before - wb.definedNames.length;
};

/**
 * Rename a defined name, scoped or workbook-scope. Returns `true` when an entry
 * was renamed. Throws when `newName` is already taken with the same scope
 * (Excel forbids duplicates within a scope).
 */
export const renameDefinedName = (
  wb: Workbook,
  oldName: string,
  newName: string,
  scope?: number,
): boolean => {
  const idx = wb.definedNames.findIndex((d) => d.name === oldName && d.scope === scope);
  if (idx < 0) return false;
  const conflict = wb.definedNames.findIndex((d, i) => i !== idx && d.name === newName && d.scope === scope);
  if (conflict >= 0) {
    throw new OpenXmlSchemaError(`renameDefinedName: "${newName}" is already in use at the same scope`);
  }
  const existing = wb.definedNames[idx];
  if (!existing) return false;
  wb.definedNames[idx] = { ...existing, name: newName };
  return true;
};

/**
 * Read-only snapshot of every `_xlnm.Print_Area` defined name. Each entry is
 * the raw DefinedName carrying `scope` (sheet index) and `value` (the
 * print-area expression like `'Sheet1'!$A$1:$D$10`).
 */
export const listPrintAreas = (wb: Workbook): ReadonlyArray<DefinedName> =>
  wb.definedNames.filter((d) => d.name === '_xlnm.Print_Area');

/**
 * Read-only snapshot of every `_xlnm.Print_Titles` defined name. Each entry's
 * `value` is the title-row / title-col expression Excel re-uses on every
 * printed page.
 */
export const listPrintTitles = (wb: Workbook): ReadonlyArray<DefinedName> =>
  wb.definedNames.filter((d) => d.name === '_xlnm.Print_Titles');

/**
 * Define the print-area for a given sheet. Excel uses the built-in
 * `_xlnm.Print_Area` defined name with sheet scope.
 */
export const setPrintArea = (wb: Workbook, sheetIndex: number, ref: string): DefinedName => {
  return addDefinedName(wb, {
    name: '_xlnm.Print_Area',
    value: ref,
    scope: sheetIndex,
  });
};

/**
 * Define print-title rows / columns on a sheet. Excel uses the
 * `_xlnm.Print_Titles` defined name. Pass `rows` ("$1:$1") to repeat row 1 on
 * every printed page; `cols` ("$A:$A") to repeat column A.
 */
export const setPrintTitles = (
  wb: Workbook,
  sheetIndex: number,
  opts: { rows?: string; cols?: string; sheetName: string },
): DefinedName => {
  const parts: string[] = [];
  // The wire form is "Sheet!$1:$1,Sheet!$A:$A"; both refs share the sheet
  // prefix.
  if (opts.cols !== undefined) parts.push(`'${opts.sheetName}'!${opts.cols}`);
  if (opts.rows !== undefined) parts.push(`'${opts.sheetName}'!${opts.rows}`);
  if (parts.length === 0) {
    throw new OpenXmlSchemaError('setPrintTitles: at least one of rows or cols must be set');
  }
  return addDefinedName(wb, {
    name: '_xlnm.Print_Titles',
    value: parts.join(','),
    scope: sheetIndex,
  });
};
