// The calculation engine: owns the dependency graph over a live Workbook,
// orders evaluation, writes results back into each FormulaValue's cache and
// manages dynamic-array spills.
//
// Ordering. A recalculation marks a set of formulas dirty (everything, or the
// transitive dependents of edited cells plus volatile formulas) and evaluates
// them in Kahn topological order computed from the reverse-dependency index —
// an explicit queue, so a 100 000-long chain never touches the JS call stack.
// Formulas that read cells dynamically (INDIRECT, OFFSET) are volatile; if one
// of them reads a formula that is still dirty, that formula is computed on
// demand. Whatever Kahn cannot order sits on or behind a cycle: Tarjan's
// algorithm finds the strongly connected components, their cells get 0 (what
// Excel shows with iteration off) and are reported in `circularRefs`, and the
// rest is ordered again.
//
// Spills. A formula whose result is a multi-cell array becomes the anchor of a
// spill: its FormulaValue turns into `t: 'array'` with `ref` covering the
// result and `cachedValue` holding the top-left element, and the other
// elements are written into the neighbouring cells as plain values — which is
// also how Excel stores a dynamic array in a file. The engine remembers which
// cells each anchor owns; it clears them when the anchor recalculates to a
// different shape and treats them as editable-by-the-user (the anchor then
// shows #SPILL!) when an update touches one. If any non-empty cell it does not
// own lies in the way, the anchor shows #SPILL! and remembers the area it
// wanted, so clearing the obstacle re-spills it.

import type { Cell, CellValue, FormulaValue } from '@office-kit/xlsx/cell';
import { richTextToString } from '@office-kit/xlsx/cell';
import { getCellNumberFormat } from '@office-kit/xlsx/styles';
import { dateToExcel } from '@office-kit/xlsx/utils';
import type { Workbook } from '@office-kit/xlsx/workbook';
import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { deleteCell, setCell } from '@office-kit/xlsx/worksheet';
import { cellAddress } from './address.ts';

/** Whether `ref` (`A1:D20`, `A1`) covers `row`. */
const refSpansRow = (ref: string, row: number): boolean => {
  const rows = ref.match(/\d+/g)?.map(Number) ?? [];
  const [r1 = 0, r2 = r1] = rows;
  return row >= Math.min(r1, r2) && row <= Math.max(r1, r2);
};
import type { AstNode, StructuredSpec } from './ast.ts';
import { serialFromDate } from './dates.ts';
import { Evaluator, type Frame, resolveRect, topLeft } from './evaluator.ts';
import { FUNCTIONS } from './functions/index.ts';
import type { EvalHost } from './host.ts';
import { CalcParseError } from './lexer.ts';
import { parseFormula } from './parser.ts';
import { RangeIndex, type Rect } from './graph.ts';
import { translateFormula } from './translate.ts';
import {
  type Area,
  type CalcArray,
  type CalcError,
  type CalcScalar,
  type CellRef,
  errorFromCode,
  ERRORS,
  isArray,
  isError,
  makeArray,
  MAX_COL,
  MAX_ROW,
} from './types.ts';

export interface CalcEngineOptions {
  /** Clock for NOW / TODAY. Defaults to the system clock. */
  readonly now?: () => Date;
  /** Source of RAND / RANDBETWEEN / RANDARRAY. Defaults to Math.random. */
  readonly random?: () => number;
}

const enum State {
  Clean,
  Dirty,
  Computing,
}

interface SheetInfo {
  readonly id: number;
  readonly ws: Worksheet;
  readonly name: string;
  /** Index in `wb.sheets`, which is what a defined name's `scope` counts. */
  readonly bookIndex: number;
  readonly cellReaders: Map<number, Set<FormulaNode>>;
  readonly rangeReaders: RangeIndex<FormulaNode>;
  maxRow: number;
  maxCol: number;
}

interface Dependency {
  readonly sheet: SheetInfo;
  readonly rect: Rect;
}

interface FormulaNode {
  readonly key: number;
  readonly sheet: SheetInfo;
  readonly row: number;
  readonly col: number;
  readonly ast: AstNode;
  /** Offset from the shared-formula master whose tree this node reuses. */
  readonly dRow: number;
  readonly dCol: number;
  /** `t` the formula was loaded with; shared formulas never spill. */
  readonly loadedKind: FormulaValue['t'];
  readonly deps: Dependency[];
  readonly volatile: boolean;
  readonly hasSubtotal: boolean;
  state: State;
  /** Area this anchor currently spills into (anchor included). */
  spill: Rect | undefined;
  /** Area a blocked anchor wanted (#SPILL!). */
  wanted: Rect | undefined;
}

// Cell keys: row * 2^14 + col fits in 34 bits; the sheet id goes above that.
const COL_SPAN = 16_384;
const SHEET_SPAN = 2 ** 34;
const cellKey = (row: number, col: number): number => row * COL_SPAN + col;

/** References this large are read only as far as the sheet's populated extent. */
const TRIM_THRESHOLD = 1 << 20;
const MAX_SPILL_PASSES = 8;
const VOLATILE_EXTRA = new Set(['CELL', 'INFO']);

const contains = (r: Rect, row: number, col: number): boolean => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;

export class CalcEngine implements EvalHost {
  private readonly evaluator: Evaluator;
  private built = false;
  private sheets: SheetInfo[] = [];
  private sheetsByName = new Map<string, SheetInfo>();
  private nodes = new Map<number, FormulaNode>();
  private volatileNodes = new Set<FormulaNode>();
  private spillOwner = new Map<number, FormulaNode>();
  private blocked = new Set<FormulaNode>();
  private names = new Map<string, AstNode | null>();
  private areaCache = new Map<string, CalcArray>();
  private parsedCache = new Map<string, AstNode>();
  /** Cells whose displayed value changed during the current pass. */
  private changed = new Map<number, CellRef>();
  /** Spill cells that appeared during a pass; their readers need another pass. */
  private lateCells: Array<{ sheet: SheetInfo; row: number; col: number }> = [];
  private circular: CellRef[] = [];

  constructor(
    private readonly wb: Workbook,
    private readonly options: CalcEngineOptions = {},
  ) {
    this.evaluator = new Evaluator(this, FUNCTIONS);
  }

  /** Circular references found by the last calculation. */
  get circularRefs(): readonly CellRef[] {
    return this.circular;
  }

  get date1904(): boolean {
    return this.wb.date1904;
  }

  // ---- public API ------------------------------------------------------------

  /** Rebuild the dependency graph and evaluate every formula. */
  recalculateAll(): void {
    this.build();
    this.runPass(new Set(this.nodes.values()));
  }

  /**
   * Cells at these addresses were edited or cleared. Recompute only their
   * transitive dependents plus volatile formulas, in dependency order, and
   * return the cells whose displayed value changed.
   */
  update(changed: Iterable<CellRef>): CellRef[] {
    if (!this.built) {
      this.build();
      this.changed.clear();
      return this.runPass(new Set(this.nodes.values()));
    }
    const roots = new Set<FormulaNode>();
    const seeds: Array<{ sheet: SheetInfo; rect: Rect }> = [];
    this.changed.clear();
    for (const ref of changed) {
      const info = this.sheetInfo(ref.sheet);
      if (info === undefined) continue;
      const key = info.id * SHEET_SPAN + cellKey(ref.row, ref.col);
      this.extend(info, ref.row, ref.col);
      const owner = this.spillOwner.get(key);
      if (owner !== undefined && owner.key !== key) {
        // The user typed into (or cleared) a spilled cell: it is theirs now.
        this.spillOwner.delete(key);
        roots.add(owner);
      }
      for (const anchor of this.blocked) if (anchor.sheet === info && anchor.wanted !== undefined && contains(anchor.wanted, ref.row, ref.col)) roots.add(anchor);
      const old = this.nodes.get(key);
      if (old !== undefined) {
        roots.delete(old);
        this.removeNode(old, seeds);
      }
      const cell = info.ws.rows.get(ref.row)?.get(ref.col);
      if (cell !== undefined && isFormula(cell.value)) {
        const node = this.createNode(info, cell, cell.value);
        if (node !== undefined) roots.add(node);
      }
      seeds.push({ sheet: info, rect: { r1: ref.row, c1: ref.col, r2: ref.row, c2: ref.col } });
    }
    const dirty = this.closure(roots, seeds);
    return this.runPass(dirty);
  }

  /** Rows were hidden, shown or filtered: SUBTOTAL / AGGREGATE results depend on that. */
  recalculateSubtotals(): CellRef[] {
    if (!this.built) return [];
    const roots = [...this.nodes.values()].filter((n) => n.hasSubtotal);
    if (roots.length === 0) return [];
    this.changed.clear();
    return this.runPass(this.closure(roots, []));
  }

  /** Rows / columns / sheets / defined names / tables changed shape: rebuild lazily on next use. */
  invalidateAll(): void {
    this.built = false;
    this.areaCache.clear();
    this.parsedCache.clear();
    this.names.clear();
  }

  /**
   * Evaluate `formula` (no leading `=`) as if it sat at (sheet, row, col).
   * Arrays collapse to their top-left value; see {@link evaluateArray}.
   * Text that does not parse evaluates to #NAME?, which is how Excel treats
   * a conditional-format rule it cannot read.
   */
  evaluate(formula: string, sheet: string, row: number, col: number): CalcScalar {
    const v = topLeft(this.evaluateArray(formula, sheet, row, col));
    // A formula never yields blank: =A1 over an empty A1 is 0, as in a cell.
    return v === null ? 0 : v;
  }

  /** Like {@link evaluate}, keeping a multi-cell result as an array. */
  evaluateArray(formula: string, sheet: string, row: number, col: number): CalcScalar | CalcArray {
    this.ensureBuilt();
    const info = this.sheetInfo(sheet);
    if (info === undefined) return ERRORS.REF;
    let ast = this.parsedCache.get(formula);
    if (ast === undefined) {
      try {
        ast = parseFormula(formula);
      } catch (e) {
        if (e instanceof CalcParseError) return ERRORS.NAME;
        throw e;
      }
      if (this.parsedCache.size > 1000) this.parsedCache.clear();
      this.parsedCache.set(formula, ast);
    }
    return this.evaluator.evaluateFormula(ast, { sheet: info.name, row, col, dRow: 0, dCol: 0, scope: undefined, depth: 0 });
  }

  // ---- EvalHost --------------------------------------------------------------

  sheetName(name: string): string | undefined {
    return this.sheetInfo(name)?.name;
  }

  sheetsBetween(first: string, last: string): string[] | undefined {
    const a = this.sheetInfo(first);
    const b = this.sheetInfo(last);
    if (a === undefined || b === undefined) return undefined;
    const i = this.sheets.indexOf(a);
    const j = this.sheets.indexOf(b);
    return this.sheets.slice(Math.min(i, j), Math.max(i, j) + 1).map((s) => s.name);
  }

  cellValue(sheet: string, row: number, col: number): CalcScalar {
    const info = this.sheetInfo(sheet);
    if (info === undefined) return ERRORS.REF;
    const cell = info.ws.rows.get(row)?.get(col);
    return cell === undefined ? null : this.readCell(info, cell);
  }

  areaValues(area: Area): CalcArray {
    const info = this.sheetInfo(area.sheet);
    if (info === undefined) return makeArray(1, 1, [ERRORS.REF]);
    let { r2, c2 } = area;
    const { r1, c1 } = area;
    if ((r2 - r1 + 1) * (c2 - c1 + 1) >= TRIM_THRESHOLD) {
      r2 = Math.max(r1, Math.min(r2, info.maxRow));
      c2 = Math.max(c1, Math.min(c2, info.maxCol));
    }
    const key = `${info.id}:${r1}:${c1}:${r2}:${c2}`;
    const cached = this.areaCache.get(key);
    if (cached !== undefined) return cached;
    const rows = r2 - r1 + 1;
    const cols = c2 - c1 + 1;
    const data = new Array<CalcScalar>(rows * cols).fill(null);
    const fillRow = (r: number, rowMap: Map<number, Cell>): void => {
      const base = (r - r1) * cols;
      if (rowMap.size < cols) {
        for (const [c, cell] of rowMap) if (c >= c1 && c <= c2) data[base + c - c1] = this.readCell(info, cell);
      } else {
        for (let c = c1; c <= c2; c++) {
          const cell = rowMap.get(c);
          if (cell !== undefined) data[base + c - c1] = this.readCell(info, cell);
        }
      }
    };
    if (rows > info.ws.rows.size * 2) {
      for (const [r, rowMap] of info.ws.rows) if (r >= r1 && r <= r2) fillRow(r, rowMap);
    } else {
      for (let r = r1; r <= r2; r++) {
        const rowMap = info.ws.rows.get(r);
        if (rowMap !== undefined) fillRow(r, rowMap);
      }
    }
    const result = makeArray(rows, cols, data);
    this.areaCache.set(key, result);
    return result;
  }

  definedName(name: string, sheet: string): AstNode | undefined {
    const info = this.sheetInfo(sheet);
    const upper = name.toUpperCase();
    const cacheKey = `${info?.bookIndex ?? -1}|${upper}`;
    const cached = this.names.get(cacheKey);
    if (cached !== undefined) return cached ?? undefined;
    const matches = this.wb.definedNames.filter((d) => d.name.toUpperCase() === upper);
    const dn = matches.find((d) => d.scope !== undefined && d.scope === info?.bookIndex) ?? matches.find((d) => d.scope === undefined);
    let ast: AstNode | null = null;
    if (dn !== undefined) {
      try {
        ast = parseFormula(dn.value);
      } catch (e) {
        if (!(e instanceof CalcParseError)) throw e;
        ast = { type: 'error', error: ERRORS.NAME };
      }
    }
    this.names.set(cacheKey, ast);
    return ast ?? undefined;
  }

  structuredArea(table: string | undefined, spec: StructuredSpec, sheet: string, row: number, col: number): Area | CalcError {
    const found = this.findTable(table, sheet, row, col);
    if (found === undefined) return ERRORS.REF;
    const { info, def, bounds } = found;
    const headerRows = def.headerRowCount ?? 1;
    const totalRows = def.totalsRowCount ?? 0;
    const dataR1 = bounds.r1 + headerRows;
    const dataR2 = bounds.r2 - totalRows;
    const columnIndex = (name: string): number => def.columns.findIndex((c) => c.name.toLowerCase() === name.trim().toLowerCase());
    let c1 = bounds.c1;
    let c2 = bounds.c2;
    if (spec.col1 !== undefined) {
      const i = columnIndex(spec.col1);
      if (i < 0) return ERRORS.REF;
      const j = spec.col2 === undefined ? i : columnIndex(spec.col2);
      if (j < 0) return ERRORS.REF;
      c1 = bounds.c1 + Math.min(i, j);
      c2 = bounds.c1 + Math.max(i, j);
    }
    const parts: Array<[number, number]> = [];
    const specials = spec.specials.length === 0 ? ['#Data'] : spec.specials;
    for (const s of specials) {
      switch (s) {
        case '#All':
          parts.push([bounds.r1, bounds.r2]);
          break;
        case '#Data':
          parts.push([dataR1, dataR2]);
          break;
        case '#Headers':
          if (headerRows === 0) return ERRORS.REF;
          parts.push([bounds.r1, dataR1 - 1]);
          break;
        case '#Totals':
          if (totalRows === 0) return ERRORS.REF;
          parts.push([dataR2 + 1, bounds.r2]);
          break;
        case '#This Row':
          if (info.name !== sheet || row < dataR1 || row > dataR2) return ERRORS.VALUE;
          parts.push([row, row]);
          break;
      }
    }
    const r1 = Math.min(...parts.map((p) => p[0]));
    const r2 = Math.max(...parts.map((p) => p[1]));
    if (r2 < r1) return ERRORS.REF;
    return { sheet: info.name, r1, c1, r2, c2 };
  }

  spillArea(sheet: string, row: number, col: number): Area | undefined {
    const info = this.sheetInfo(sheet);
    if (info === undefined) return undefined;
    const node = this.nodes.get(info.id * SHEET_SPAN + cellKey(row, col));
    if (node === undefined) return undefined;
    if (node.state === State.Dirty) this.computeNode(node);
    const r = node.spill ?? { r1: row, c1: col, r2: row, c2: col };
    return { sheet: info.name, ...r };
  }

  formulaText(sheet: string, row: number, col: number): string | undefined {
    const info = this.sheetInfo(sheet);
    const cell = info?.ws.rows.get(row)?.get(col);
    if (info === undefined || cell === undefined || !isFormula(cell.value)) return undefined;
    if (cell.value.formula !== '') return cell.value.formula;
    const node = this.nodes.get(info.id * SHEET_SPAN + cellKey(row, col));
    const master = node === undefined ? undefined : info.ws.rows.get(row - node.dRow)?.get(col - node.dCol);
    if (node === undefined || master === undefined || !isFormula(master.value)) return undefined;
    return translateFormula(master.value.formula, node.dRow, node.dCol);
  }

  isSubtotalCell(sheet: string, row: number, col: number): boolean {
    const info = this.sheetInfo(sheet);
    if (info === undefined) return false;
    return this.nodes.get(info.id * SHEET_SPAN + cellKey(row, col))?.hasSubtotal === true;
  }

  isRowHidden(sheet: string, row: number): boolean {
    return this.sheetInfo(sheet)?.ws.rowDimensions.get(row)?.hidden === true;
  }

  isRowFiltered(sheet: string, row: number): boolean {
    const ws = this.sheetInfo(sheet)?.ws;
    if (ws?.rowDimensions.get(row)?.hidden !== true) return false;
    const filters = [ws.autoFilter, ...ws.tables.map((t) => t.autoFilter)];
    return filters.some((af) => af !== undefined && af.filterColumns.length > 0 && refSpansRow(af.ref, row));
  }

  numberFormat(sheet: string, row: number, col: number): string {
    const cell = this.sheetInfo(sheet)?.ws.rows.get(row)?.get(col);
    return cell === undefined ? 'General' : getCellNumberFormat(this.wb, cell);
  }

  now(): number {
    return serialFromDate(this.options.now?.() ?? new Date(), this.wb.date1904);
  }

  random(): number {
    return (this.options.random ?? Math.random)();
  }

  // ---- graph construction ----------------------------------------------------

  private ensureBuilt(): void {
    if (!this.built) this.recalculateAll();
  }

  private sheetInfo(name: string): SheetInfo | undefined {
    return this.sheetsByName.get(name) ?? this.sheetsByName.get(name.toLowerCase());
  }

  private build(): void {
    this.sheets = [];
    this.sheetsByName = new Map();
    this.nodes = new Map();
    this.volatileNodes = new Set();
    this.spillOwner = new Map();
    this.blocked = new Set();
    this.names = new Map();
    this.areaCache = new Map();
    this.wb.sheets.forEach((ref, bookIndex) => {
      if (ref.kind !== 'worksheet') return;
      const info: SheetInfo = {
        id: this.sheets.length,
        ws: ref.sheet,
        name: ref.sheet.title,
        bookIndex,
        cellReaders: new Map(),
        rangeReaders: new RangeIndex(),
        maxRow: 1,
        maxCol: 1,
      };
      for (const [r, rowMap] of ref.sheet.rows) {
        if (r > info.maxRow) info.maxRow = r;
        for (const c of rowMap.keys()) if (c > info.maxCol) info.maxCol = c;
      }
      this.sheets.push(info);
      this.sheetsByName.set(info.name, info);
      this.sheetsByName.set(info.name.toLowerCase(), info);
    });
    this.built = true;
    for (const info of this.sheets) {
      for (const rowMap of info.ws.rows.values()) {
        for (const cell of rowMap.values()) if (isFormula(cell.value)) this.createNode(info, cell, cell.value);
      }
    }
  }

  private createNode(info: SheetInfo, cell: Cell, fv: FormulaValue): FormulaNode | undefined {
    // What-if data tables need Excel's table engine; their cached results stay as loaded.
    if (fv.t === 'dataTable') return undefined;
    let ast: AstNode;
    let dRow = 0;
    let dCol = 0;
    if (fv.t === 'shared' && fv.formula === '') {
      const master = this.sharedMaster(info, fv.si, cell.row, cell.col);
      if (master === undefined) return undefined;
      ast = master.ast;
      dRow = cell.row - master.row;
      dCol = cell.col - master.col;
    } else {
      ast = this.parseCached(fv.formula);
    }
    const frame = { sheet: info.name, row: cell.row, col: cell.col, dRow, dCol };
    const deps: Dependency[] = [];
    const flags = { volatile: false, subtotal: false };
    this.collectDependencies(ast, frame, deps, flags, 0);
    const node: FormulaNode = {
      key: info.id * SHEET_SPAN + cellKey(cell.row, cell.col),
      sheet: info,
      row: cell.row,
      col: cell.col,
      ast,
      dRow,
      dCol,
      loadedKind: fv.t,
      deps,
      volatile: flags.volatile,
      hasSubtotal: flags.subtotal,
      state: State.Clean,
      spill: undefined,
      wanted: undefined,
    };
    this.nodes.set(node.key, node);
    if (node.volatile) this.volatileNodes.add(node);
    for (const d of deps) {
      if (d.rect.r1 === d.rect.r2 && d.rect.c1 === d.rect.c2) {
        const k = cellKey(d.rect.r1, d.rect.c1);
        let set = d.sheet.cellReaders.get(k);
        if (set === undefined) {
          set = new Set();
          d.sheet.cellReaders.set(k, set);
        }
        set.add(node);
      } else {
        d.sheet.rangeReaders.add(d.rect, node);
      }
    }
    // A loaded array formula already owns the cells of its `ref`.
    if (fv.t === 'array' && fv.ref !== undefined) {
      const rect = parseA1Range(fv.ref);
      if (rect !== undefined && rect.r1 === cell.row && rect.c1 === cell.col) {
        node.spill = rect;
        for (let r = rect.r1; r <= rect.r2; r++) {
          for (let c = rect.c1; c <= rect.c2; c++) if (r !== cell.row || c !== cell.col) this.spillOwner.set(info.id * SHEET_SPAN + cellKey(r, c), node);
        }
      }
    }
    return node;
  }

  private sharedMaster(info: SheetInfo, si: number | undefined, row: number, col: number): { ast: AstNode; row: number; col: number } | undefined {
    if (si === undefined) return undefined;
    // The master is the group's cell carrying text and a ref covering the follower.
    for (const rowMap of info.ws.rows.values()) {
      for (const cell of rowMap.values()) {
        const v = cell.value;
        if (!isFormula(v) || v.t !== 'shared' || v.si !== si || v.formula === '') continue;
        const rect = v.ref === undefined ? undefined : parseA1Range(v.ref);
        if (rect === undefined || contains(rect, row, col)) return { ast: this.parseCached(v.formula), row: cell.row, col: cell.col };
      }
    }
    return undefined;
  }

  private parseCached(text: string): AstNode {
    let ast = this.parsedCache.get(`=${text}`);
    if (ast === undefined) {
      try {
        ast = parseFormula(text);
      } catch (e) {
        if (!(e instanceof CalcParseError)) throw e;
        ast = { type: 'error', error: ERRORS.NAME };
      }
      this.parsedCache.set(`=${text}`, ast);
    }
    return ast;
  }

  /** Static precedents of a tree. Over-approximation is harmless; misses are not, so dynamic readers are volatile. */
  private collectDependencies(
    node: AstNode,
    frame: { sheet: string; row: number; col: number; dRow: number; dCol: number },
    out: Dependency[],
    flags: { volatile: boolean; subtotal: boolean },
    depth: number,
  ): void {
    switch (node.type) {
      case 'ref': {
        const rect = resolveRect(node.area, frame.dRow, frame.dCol);
        for (const s of this.prefixSheets(node.prefix, frame.sheet)) out.push({ sheet: s, rect });
        return;
      }
      case 'name': {
        if (depth > 32) return;
        const sheet = node.prefix === undefined ? frame.sheet : (this.sheetInfo(node.prefix.sheet)?.name ?? frame.sheet);
        const body = this.definedName(node.name, sheet);
        if (body !== undefined) this.collectDependencies(body, { ...frame, dRow: frame.row - 1, dCol: frame.col - 1 }, out, flags, depth + 1);
        return;
      }
      case 'structured': {
        const area = this.structuredArea(node.table, node.spec, frame.sheet, frame.row, frame.col);
        const info = isError(area) ? undefined : this.sheetInfo(area.sheet);
        if (!isError(area) && info !== undefined) out.push({ sheet: info, rect: area });
        return;
      }
      case 'unary':
      case 'postfix':
        this.collectDependencies(node.operand, frame, out, flags, depth);
        return;
      case 'binary':
        this.collectDependencies(node.left, frame, out, flags, depth);
        this.collectDependencies(node.right, frame, out, flags, depth);
        return;
      case 'call': {
        const spec = FUNCTIONS.get(node.name);
        if (spec?.volatile === true || VOLATILE_EXTRA.has(node.name)) flags.volatile = true;
        if (node.name === 'SUBTOTAL' || node.name === 'AGGREGATE') flags.subtotal = true;
        if (spec === undefined) {
          // A call to a name: a LAMBDA held in a defined name.
          this.collectDependencies({ type: 'name', name: node.name, prefix: undefined }, frame, out, flags, depth);
        }
        for (const a of node.args) this.collectDependencies(a, frame, out, flags, depth);
        return;
      }
      case 'invoke':
        this.collectDependencies(node.callee, frame, out, flags, depth);
        for (const a of node.args) this.collectDependencies(a, frame, out, flags, depth);
    }
  }

  private prefixSheets(prefix: { sheet: string; sheet2?: string; external?: string } | undefined, current: string): SheetInfo[] {
    if (prefix === undefined) {
      const info = this.sheetInfo(current);
      return info === undefined ? [] : [info];
    }
    if (prefix.external !== undefined) return [];
    if (prefix.sheet2 !== undefined) {
      return (this.sheetsBetween(prefix.sheet, prefix.sheet2) ?? []).flatMap((n) => this.sheetInfo(n) ?? []);
    }
    const info = this.sheetInfo(prefix.sheet);
    return info === undefined ? [] : [info];
  }

  private removeNode(node: FormulaNode, seeds: Array<{ sheet: SheetInfo; rect: Rect }>): void {
    this.nodes.delete(node.key);
    this.volatileNodes.delete(node);
    this.blocked.delete(node);
    for (const d of node.deps) {
      if (d.rect.r1 === d.rect.r2 && d.rect.c1 === d.rect.c2) {
        d.sheet.cellReaders.get(cellKey(d.rect.r1, d.rect.c1))?.delete(node);
      } else {
        d.sheet.rangeReaders.remove(d.rect, node);
      }
    }
    if (node.spill !== undefined) {
      this.releaseSpill(node, undefined);
      seeds.push({ sheet: node.sheet, rect: node.spill });
      node.spill = undefined;
    }
  }

  // ---- dirty sets and ordering -------------------------------------------------

  /** Formulas reading any cell of `rect` on `sheet`. */
  private readersOf(sheet: SheetInfo, rect: Rect, out: Set<FormulaNode>): void {
    const area = (rect.r2 - rect.r1 + 1) * (rect.c2 - rect.c1 + 1);
    if (area <= sheet.cellReaders.size) {
      for (let r = rect.r1; r <= rect.r2; r++) {
        for (let c = rect.c1; c <= rect.c2; c++) {
          const set = sheet.cellReaders.get(cellKey(r, c));
          if (set !== undefined) for (const n of set) out.add(n);
        }
      }
    } else {
      for (const [k, set] of sheet.cellReaders) {
        const r = Math.floor(k / COL_SPAN);
        const c = k - r * COL_SPAN;
        if (contains(rect, r, c)) for (const n of set) out.add(n);
      }
    }
    sheet.rangeReaders.query(rect, out);
  }

  private dependentsOf(node: FormulaNode): Set<FormulaNode> {
    const out = new Set<FormulaNode>();
    this.readersOf(node.sheet, node.spill ?? { r1: node.row, c1: node.col, r2: node.row, c2: node.col }, out);
    return out;
  }

  private closure(roots: Iterable<FormulaNode>, seeds: ReadonlyArray<{ sheet: SheetInfo; rect: Rect }>, withVolatile = true): Set<FormulaNode> {
    const dirty = new Set<FormulaNode>();
    const queue: FormulaNode[] = [];
    const push = (n: FormulaNode): void => {
      if (dirty.has(n)) return;
      dirty.add(n);
      queue.push(n);
    };
    for (const n of roots) push(n);
    if (withVolatile) for (const n of this.volatileNodes) push(n);
    for (const s of seeds) {
      const readers = new Set<FormulaNode>();
      this.readersOf(s.sheet, s.rect, readers);
      for (const n of readers) push(n);
    }
    while (queue.length > 0) {
      const n = queue.pop();
      if (n === undefined) break;
      for (const d of this.dependentsOf(n)) push(d);
    }
    return dirty;
  }

  private runPass(initial: Set<FormulaNode>): CellRef[] {
    this.circular = [];
    let dirty = initial;
    for (let pass = 0; pass < MAX_SPILL_PASSES && dirty.size > 0; pass++) {
      this.lateCells = [];
      this.areaCache.clear();
      this.evaluateSet(dirty);
      if (this.lateCells.length === 0) break;
      const seeds = this.lateCells.map((c) => ({ sheet: c.sheet, rect: { r1: c.row, c1: c.col, r2: c.row, c2: c.col } }));
      dirty = this.closure([], seeds, false);
    }
    this.areaCache.clear();
    const result = [...this.changed.values()];
    this.changed.clear();
    return result;
  }

  private evaluateSet(dirty: Set<FormulaNode>): void {
    for (const n of dirty) n.state = State.Dirty;
    let remaining = this.kahn(dirty);
    if (remaining.length === 0) return;

    // Whatever is left sits on or behind a cycle.
    this.areaCache.clear();
    const pending = new Set(remaining);
    for (const component of this.stronglyConnected(pending)) {
      const first = component[0];
      const cyclic = component.length > 1 || (first !== undefined && this.dependentsOf(first).has(first));
      if (!cyclic) continue;
      for (const n of component) {
        n.state = State.Clean;
        pending.delete(n);
        this.writeScalar(n, 0);
        this.circular.push({ sheet: n.sheet.name, row: n.row, col: n.col });
      }
    }
    this.areaCache.clear();
    remaining = this.kahn(pending);
    // Anything still left reads a cycle through a dynamic reference; compute it directly.
    for (const n of remaining) if (n.state === State.Dirty) this.computeNode(n);
  }

  /** Evaluate `set` in topological order; returns the nodes that could not be ordered. */
  private kahn(set: Set<FormulaNode>): FormulaNode[] {
    const indegree = new Map<FormulaNode, number>();
    const edges = new Map<FormulaNode, FormulaNode[]>();
    for (const n of set) indegree.set(n, 0);
    for (const n of set) {
      const out: FormulaNode[] = [];
      for (const d of this.dependentsOf(n)) {
        if (!set.has(d)) continue;
        out.push(d);
        indegree.set(d, (indegree.get(d) ?? 0) + 1);
      }
      edges.set(n, out);
    }
    const queue: FormulaNode[] = [];
    for (const [n, deg] of indegree) if (deg === 0) queue.push(n);
    let done = 0;
    while (queue.length > 0) {
      const n = queue.pop();
      if (n === undefined) break;
      done++;
      if (n.state === State.Dirty) this.computeNode(n);
      for (const d of edges.get(n) ?? []) {
        const deg = (indegree.get(d) ?? 0) - 1;
        indegree.set(d, deg);
        if (deg === 0) queue.push(d);
      }
    }
    if (done === set.size) return [];
    return [...set].filter((n) => (indegree.get(n) ?? 0) > 0);
  }

  /** Tarjan's SCC, iterative so long cycles cannot overflow the stack. */
  private stronglyConnected(set: Set<FormulaNode>): FormulaNode[][] {
    const index = new Map<FormulaNode, number>();
    const low = new Map<FormulaNode, number>();
    const onStack = new Set<FormulaNode>();
    const stack: FormulaNode[] = [];
    const components: FormulaNode[][] = [];
    let counter = 0;
    const succ = (n: FormulaNode): FormulaNode[] => [...this.dependentsOf(n)].filter((d) => set.has(d));
    const work: Array<{ node: FormulaNode; next: FormulaNode[]; i: number }> = [];
    const enter = (n: FormulaNode): void => {
      index.set(n, counter);
      low.set(n, counter);
      counter++;
      stack.push(n);
      onStack.add(n);
      work.push({ node: n, next: succ(n), i: 0 });
    };
    for (const root of set) {
      if (index.has(root)) continue;
      enter(root);
      while (work.length > 0) {
        const frame = work[work.length - 1];
        if (frame === undefined) break;
        if (frame.i < frame.next.length) {
          const w = frame.next[frame.i++];
          if (w === undefined) continue;
          if (!index.has(w)) enter(w);
          else if (onStack.has(w)) low.set(frame.node, Math.min(low.get(frame.node) ?? 0, index.get(w) ?? 0));
          continue;
        }
        work.pop();
        const parent = work[work.length - 1];
        if (parent !== undefined) low.set(parent.node, Math.min(low.get(parent.node) ?? 0, low.get(frame.node) ?? 0));
        if (low.get(frame.node) === index.get(frame.node)) {
          const component: FormulaNode[] = [];
          for (;;) {
            const w = stack.pop();
            if (w === undefined) break;
            onStack.delete(w);
            component.push(w);
            if (w === frame.node) break;
          }
          components.push(component);
        }
      }
    }
    return components;
  }

  // ---- evaluation and write-back -------------------------------------------------

  private computeNode(node: FormulaNode): void {
    node.state = State.Computing;
    const frame: Frame = { sheet: node.sheet.name, row: node.row, col: node.col, dRow: node.dRow, dCol: node.dCol, scope: undefined, depth: 0 };
    const result = this.evaluator.evaluateFormula(node.ast, frame);
    node.state = State.Clean;
    if (isArray(result) && result.data.length > 1 && node.loadedKind !== 'shared') this.writeSpill(node, result);
    else this.writeScalar(node, topLeft(result));
  }

  private readCell(info: SheetInfo, cell: Cell): CalcScalar {
    const v = cell.value;
    if (isFormula(v)) {
      const node = this.nodes.get(info.id * SHEET_SPAN + cellKey(cell.row, cell.col));
      if (node !== undefined && node.state !== State.Clean) {
        // A cycle reached through a dynamic reference reads as 0, like Excel.
        if (node.state === State.Computing) return 0;
        this.computeNode(node);
        this.areaCache.clear();
        return this.readCell(info, info.ws.rows.get(cell.row)?.get(cell.col) ?? cell);
      }
      if (v.cachedValue === undefined) return null;
      if (v.cachedValueType === 'error') return errorFromCode(String(v.cachedValue)) ?? ERRORS.VALUE;
      return v.cachedValue;
    }
    return modelToScalar(v, this.wb.date1904);
  }

  private writeScalar(node: FormulaNode, value: CalcScalar): void {
    if (node.spill !== undefined) {
      this.releaseSpill(node, undefined);
      node.spill = undefined;
    }
    if (node.wanted !== undefined && !(isError(value) && value.code === '#SPILL!')) {
      node.wanted = undefined;
      this.blocked.delete(node);
    }
    // A loaded (CSE) array formula stays one, sized to its own cell.
    this.setFormulaResult(node, value, node.loadedKind, node.loadedKind === 'array' ? cellAddress(node.row, node.col) : undefined);
  }

  private writeSpill(node: FormulaNode, value: CalcArray): void {
    const rect: Rect = { r1: node.row, c1: node.col, r2: node.row + value.rows - 1, c2: node.col + value.cols - 1 };
    const ws = node.sheet.ws;
    let blocked = rect.r2 > MAX_ROW || rect.c2 > MAX_COL;
    for (let r = rect.r1; r <= rect.r2 && !blocked; r++) {
      const rowMap = ws.rows.get(r);
      if (rowMap === undefined) continue;
      for (let c = rect.c1; c <= rect.c2; c++) {
        if (r === node.row && c === node.col) continue;
        const cell = rowMap.get(c);
        if (cell === undefined || cell.value === null) continue;
        if (this.spillOwner.get(node.sheet.id * SHEET_SPAN + cellKey(r, c)) !== node) {
          blocked = true;
          break;
        }
      }
    }
    if (blocked) {
      this.writeScalar(node, ERRORS.SPILL);
      node.wanted = rect;
      this.blocked.add(node);
      return;
    }
    node.wanted = undefined;
    this.blocked.delete(node);
    if (node.spill !== undefined) this.releaseSpill(node, rect);
    for (let r = rect.r1; r <= rect.r2; r++) {
      for (let c = rect.c1; c <= rect.c2; c++) {
        if (r === node.row && c === node.col) continue;
        const key = node.sheet.id * SHEET_SPAN + cellKey(r, c);
        const v = value.data[(r - rect.r1) * value.cols + (c - rect.c1)] ?? null;
        const next = scalarToModel(v);
        const cell = ws.rows.get(r)?.get(c);
        const wasOwned = this.spillOwner.get(key) === node;
        if (cell === undefined || !sameModelValue(cell.value, next)) {
          setCell(ws, r, c, next);
          this.markChanged(node.sheet, r, c);
          // Readers of a cell outside the previous spill were not ordered after this anchor.
          if (!wasOwned) this.lateCells.push({ sheet: node.sheet, row: r, col: c });
        }
        this.spillOwner.set(key, node);
      }
    }
    if (rect.r2 > node.sheet.maxRow) node.sheet.maxRow = rect.r2;
    if (rect.c2 > node.sheet.maxCol) node.sheet.maxCol = rect.c2;
    node.spill = rect;
    this.areaCache.clear();
    this.setFormulaResult(node, value.data[0] ?? null, 'array', `${cellAddress(rect.r1, rect.c1)}:${cellAddress(rect.r2, rect.c2)}`);
  }

  /** Clear the cells `node` owns outside `keep`. */
  private releaseSpill(node: FormulaNode, keep: Rect | undefined): void {
    const spill = node.spill;
    if (spill === undefined) return;
    const ws = node.sheet.ws;
    for (let r = spill.r1; r <= spill.r2; r++) {
      for (let c = spill.c1; c <= spill.c2; c++) {
        if ((r === node.row && c === node.col) || (keep !== undefined && contains(keep, r, c))) continue;
        const key = node.sheet.id * SHEET_SPAN + cellKey(r, c);
        if (this.spillOwner.get(key) !== node) continue;
        this.spillOwner.delete(key);
        const cell = ws.rows.get(r)?.get(c);
        if (cell === undefined) continue;
        if (cell.styleId === 0 && cell.hyperlinkId === undefined && cell.commentId === undefined) deleteCell(ws, r, c);
        else cell.value = null;
        this.markChanged(node.sheet, r, c);
        this.lateCells.push({ sheet: node.sheet, row: r, col: c });
      }
    }
    this.areaCache.clear();
  }

  private setFormulaResult(node: FormulaNode, value: CalcScalar, t: FormulaValue['t'], ref: string | undefined): void {
    const cell = node.sheet.ws.rows.get(node.row)?.get(node.col);
    if (cell === undefined || !isFormula(cell.value)) return;
    const old = cell.value;
    // A top-level blank shows as 0 in Excel (`=A1` with A1 empty).
    const v = value === null ? 0 : value;
    const cached = isError(v) ? v.code : v;
    const errorType = isError(v);
    const next: FormulaValue = Object.freeze({
      kind: 'formula',
      formula: old.formula,
      t,
      cachedValue: cached,
      ...(errorType ? { cachedValueType: 'error' as const } : {}),
      ...(t === 'shared' && old.ref !== undefined ? { ref: old.ref } : ref !== undefined ? { ref } : {}),
      ...(old.si !== undefined ? { si: old.si } : {}),
      ...(old.aca !== undefined ? { aca: old.aca } : {}),
      ...(old.ca !== undefined ? { ca: old.ca } : {}),
    });
    const changed = old.cachedValue !== next.cachedValue || old.cachedValueType !== next.cachedValueType;
    cell.value = next;
    if (changed) this.markChanged(node.sheet, node.row, node.col);
  }

  private markChanged(sheet: SheetInfo, row: number, col: number): void {
    this.changed.set(sheet.id * SHEET_SPAN + cellKey(row, col), { sheet: sheet.name, row, col });
  }

  private extend(info: SheetInfo, row: number, col: number): void {
    if (row > info.maxRow) info.maxRow = row;
    if (col > info.maxCol) info.maxCol = col;
    this.areaCache.clear();
  }

  private findTable(name: string | undefined, sheet: string, row: number, col: number) {
    const own = this.sheetInfo(sheet);
    const candidates = name !== undefined ? this.sheets : own === undefined ? [] : [own];
    for (const info of candidates) {
      for (const def of info.ws.tables) {
        const bounds = parseA1Range(def.ref);
        if (bounds === undefined) continue;
        if (name !== undefined ? def.displayName.toLowerCase() === name.toLowerCase() : contains(bounds, row, col)) return { info, def, bounds };
      }
    }
    return undefined;
  }
}

// ---- model value conversion ------------------------------------------------------

function isFormula(v: CellValue): v is FormulaValue {
  return typeof v === 'object' && v !== null && !(v instanceof Date) && v.kind === 'formula';
}

function modelToScalar(v: CellValue, date1904: boolean): CalcScalar {
  if (v === null || typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return v;
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? ERRORS.VALUE : dateToExcel(v, { epoch: date1904 ? 'mac' : 'windows' });
  switch (v.kind) {
    case 'duration':
      return v.ms / 86_400_000;
    case 'error':
      return errorFromCode(v.code) ?? ERRORS.VALUE;
    case 'rich-text':
      return richTextToString(v.runs);
    case 'formula':
      return null;
  }
}

function scalarToModel(v: CalcScalar): CellValue {
  if (v === null) return 0;
  if (isError(v)) return { kind: 'error', code: v.code };
  return v;
}

function sameModelValue(a: CellValue, b: CellValue): boolean {
  if (a === b) return true;
  return typeof a === 'object' && a !== null && typeof b === 'object' && b !== null && !(a instanceof Date) && !(b instanceof Date) && a.kind === 'error' && b.kind === 'error' && a.code === b.code;
}

const A1_RANGE = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/;

function parseA1Range(ref: string): Rect | undefined {
  const m = A1_RANGE.exec(ref.trim());
  if (m === null) return undefined;
  const col = (s: string): number => {
    let n = 0;
    for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64;
    return n;
  };
  const c1 = col(m[1] ?? '');
  const r1 = Number(m[2]);
  const c2 = m[3] === undefined ? c1 : col(m[3]);
  const r2 = m[4] === undefined ? r1 : Number(m[4]);
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}
