// Runs REPL code against the library, entirely in the browser. The REPL page
// imports this module dynamically, so the library is fetched when the page
// mounts and stays out of the route's initial JavaScript.

import * as cell from '@office-kit/xlsx/cell';
import { fromArrayBuffer, loadWorkbook, workbookToBytes } from '@office-kit/xlsx/io';
import * as styles from '@office-kit/xlsx/styles';
import * as utils from '@office-kit/xlsx/utils';
import * as workbook from '@office-kit/xlsx/workbook';
import * as worksheet from '@office-kit/xlsx/worksheet';
import { readWorkbookGrids, type SheetGridData } from '$lib/sheet-grid';

// A DOM table of a whole sheet would freeze the tab, so the preview stops at
// the maximums. The minimums pad a small sheet out to a pane's worth of empty
// cells, the way a spreadsheet looks before anyone types in it.
export const PREVIEW_LIMIT = { maxRows: 200, maxCols: 40, minRows: 30, minCols: 8 };

const subpaths = { cell, styles, utils, workbook, worksheet };

// User code sees every export as a bare name. The package has no root import
// because a name's subpath is part of its identity, and two subpaths may use
// one name for different functions (`isMergedCell` is in both /cell and
// /worksheet). Such a name cannot be bare, so it becomes a stub that says where
// to find the real ones.
function buildScope(): Map<string, unknown> {
  const owners = new Map<string, string[]>();
  for (const [subpath, exports] of Object.entries(subpaths)) {
    for (const name of Object.keys(exports)) {
      owners.set(name, [...(owners.get(name) ?? []), subpath]);
    }
  }
  const scope = new Map<string, unknown>([['xlsx', subpaths]]);
  for (const exports of Object.values(subpaths)) {
    for (const [name, value] of Object.entries(exports)) {
      const homes = owners.get(name) ?? [];
      scope.set(
        name,
        homes.length === 1
          ? value
          : () => {
              const choices = homes.map((home) => `xlsx.${home}.${name}(…)`).join(' or ');
              throw new Error(`${name} is exported by more than one subpath. Say which one: ${choices}`);
            },
      );
    }
  }
  return scope;
}

const scope = buildScope();
const scopeNames = [...scope.keys()];
const scopeValues = [...scope.values()];

// `new Function` wraps its source in a two-line header, and the wrapper below
// adds two more lines before the user's code, so a stack frame's line number
// is off by four.
const WRAPPER_LINES = 4;
const USER_FRAME = /(?:<anonymous>|Function):(\d+):\d+/;

export class ReplError extends Error {
  /** 1-based line in the user's code, when the engine reported one. */
  readonly line: number | undefined;
  /** The code did not parse. Engines give such an error no line; the editor can. */
  readonly syntax: boolean;

  constructor(message: string, options: { line?: number | undefined; syntax?: boolean } = {}) {
    super(message);
    this.name = 'ReplError';
    this.line = options.line;
    this.syntax = options.syntax ?? false;
  }
}

const IMPORT_LINE = /^\s*import\s/m;

// A stack trace through the bundled library is noise to someone writing a few
// lines of code. Keep the message, find the line it came from, and add a hint
// for the two mistakes that are otherwise baffling: a lookup that matched
// nothing, and an `import` pasted in from a real module.
function explain(err: unknown, code: string): ReplError {
  if (!(err instanceof Error)) return new ReplError(String(err));
  const frame = USER_FRAME.exec(err.stack ?? '');
  const line = frame ? Number(frame[1]) - WRAPPER_LINES : 0;
  const syntax = err instanceof SyntaxError;
  let hint = '';
  if (syntax && IMPORT_LINE.test(code)) {
    hint = '\nImports do not work here, and are not needed: delete the import lines, every function is already in scope.';
  } else if (/of (null|undefined)/.test(err.message)) {
    hint = '\nSomething on that line is undefined. getSheet, getCell, and the other get… functions return undefined when nothing matches, so check the name or coordinates you passed.';
  }
  return new ReplError(`${err.message}${hint}`, { line: line > 0 ? line : undefined, syntax });
}

export type ReplResult = {
  /** The saved workbook: what the preview was read from, and what Download hands over. */
  bytes: Uint8Array;
  sheets: SheetGridData[];
};

export async function runCode(code: string): Promise<ReplResult> {
  const wb = workbook.createWorkbook();
  try {
    // Wrapped in an async function so the code may use `await`. Evaluating the
    // visitor's own code in their own tab is the whole point of the page.
    // eslint-disable-next-line no-new-func
    const fn = new Function(
      ...scopeNames,
      'wb',
      `'use strict';\nreturn (async () => {\n${code}\n})();`,
    );
    await fn(...scopeValues, wb);
  } catch (err) {
    throw explain(err, code);
  }

  // The preview is read from the saved bytes, not from `wb`, so what is on
  // screen has been through a full write and read.
  try {
    const bytes = await workbookToBytes(wb);
    const sheets = readWorkbookGrids(await loadWorkbook(fromArrayBuffer(bytes)), PREVIEW_LIMIT);
    return { bytes, sheets };
  } catch (err) {
    throw new ReplError(`The workbook could not be saved: ${err instanceof Error ? err.message : String(err)}`);
  }
}
