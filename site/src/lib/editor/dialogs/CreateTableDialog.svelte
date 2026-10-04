<script lang="ts">
  import { makeCell } from '@office-kit/xlsx/cell';
  import { getCellDisplayText } from '@office-kit/xlsx/styles';
  import { addExcelTable } from '@office-kit/xlsx/worksheet';
  import { MAX_ROW, parseRangeAddress, rangeAddress, rangesIntersect, type Range } from '../core/address.ts';
  import { getCellAt } from '../core/cells.ts';
  import { getEditor } from '../core/context.ts';
  import { autoFilterRange } from '../core/filter.ts';
  import { currentRegion } from '../core/navigation.ts';
  import { currentRange, selectRange } from '../core/selection.ts';
  import { structuralEdit } from '../core/structure.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  /** Excel's default table style. */
  const DEFAULT_STYLE = 'TableStyleMedium2';

  const ctl = getEditor();
  const doc = ctl.doc;
  const ws = doc.ws;
  const sel = currentRange(doc.selection);
  const initialRange = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? currentRegion(ws, doc.selection.active) : sel;
  let refText = $state(`=${rangeAddress(initialRange, true)}`);
  // Excel proposes headers when the first row is all text over more data.
  const firstRowText = Array.from({ length: initialRange.c2 - initialRange.c1 + 1 }, (_, i) => getCellAt(ws, initialRange.r1, initialRange.c1 + i)?.value).every((v) => typeof v === 'string' && v !== '');
  let hasHeaders = $state(initialRange.r2 > initialRange.r1 && firstRowText);
  let notice = $state<string | null>(null);

  function tableRange(text: string): Range | undefined {
    const parsed = parseRangeAddress(text.trim().replace(/^=/, ''));
    if (!parsed || (parsed.sheet !== undefined && parsed.sheet.toLowerCase() !== ws.title.toLowerCase())) return undefined;
    const r = parsed.range;
    // Whole columns shrink to the used rows, as Excel does.
    if (r.r2 === MAX_ROW) {
      let last = r.r1;
      for (const row of ws.rows.keys()) if (row > last) last = row;
      return { ...r, r2: last };
    }
    return r;
  }

  function tableRangeOf(ref: string): Range | undefined {
    return parseRangeAddress(ref)?.range;
  }

  function uniqueTableName(): string {
    const taken = new Set<string>();
    for (const s of doc.wb.sheets) if (s.kind === 'worksheet') for (const tb of s.sheet.tables) taken.add(tb.displayName.toLowerCase());
    for (const dn of doc.wb.definedNames) taken.add(dn.name.toLowerCase());
    let n = 1;
    while (taken.has(`table${n}`)) n++;
    return `Table${n}`;
  }

  /** Header texts: blanks become ColumnN and repeats get a numeric suffix, as Excel names them. */
  function headerNames(range: Range, fromCells: boolean): string[] {
    const seen = new Set<string>();
    const out: string[] = [];
    for (let c = range.c1; c <= range.c2; c++) {
      const cell = fromCells ? getCellAt(ws, range.r1, c) : undefined;
      let name = cell && cell.value !== null ? getCellDisplayText(doc.wb, cell).trim() : '';
      if (!name) name = `Column${c - range.c1 + 1}`;
      let unique = name;
      for (let k = 2; seen.has(unique.toLowerCase()); k++) unique = `${name}${k}`;
      seen.add(unique.toLowerCase());
      out.push(unique);
    }
    return out;
  }

  function onok(): boolean {
    const range = tableRange(refText);
    if (!range) {
      notice = t('invalidReference');
      return false;
    }
    const overlapsTable = ws.tables.some((tb) => {
      const r = tableRangeOf(tb.ref);
      return r !== undefined && rangesIntersect(r, range);
    });
    if (overlapsTable) {
      notice = t('dlgTableOverlap');
      return false;
    }
    if (ws.mergedCells.some((m) => rangesIntersect({ r1: m.minRow, c1: m.minCol, r2: m.maxRow, c2: m.maxCol }, range))) {
      notice = t('dlgTableMerged');
      return false;
    }
    if (!hasHeaders && range.r2 >= MAX_ROW) {
      notice = t('dlgTableNoRoom');
      return false;
    }
    const final: Range = hasHeaders ? range : { ...range, r2: range.r2 + 1 };
    const names = headerNames(range, hasHeaders);
    const ref = rangeAddress(final);
    doc.transact('Create Table', (tx) => {
      tx.structural = true;
      if (hasHeaders) {
        tx.cells(ws, { ...final, r2: final.r1 });
        tx.sheet(ws, 'tables', 'autoFilter');
      } else {
        // A header row is inserted above the data, shifting only the table's columns down.
        structuralEdit(tx, doc.wb, ws, { axis: 'row', at: range.r1, count: 1, band: { from: range.c1, to: range.c2 } });
      }
      names.forEach((name, i) => {
        const col = final.c1 + i;
        const cell = getCellAt(ws, final.r1, col);
        if (cell) cell.value = name;
        else {
          let rowMap = ws.rows.get(final.r1);
          if (!rowMap) {
            rowMap = new Map();
            ws.rows.set(final.r1, rowMap);
          }
          rowMap.set(col, makeCell(final.r1, col, name, ctl.defaultStyleAt(final.r1, col)));
        }
      });
      // A table owns its filter; a sheet AutoFilter over the same cells would conflict.
      const af = autoFilterRange(ws);
      if (af && rangesIntersect(af, final)) delete ws.autoFilter;
      addExcelTable(doc.wb, ws, { name: uniqueTableName(), ref, columns: names, style: DEFAULT_STYLE, autoFilter: { ref, filterColumns: [] } });
    });
    doc.setSelection(selectRange(final, doc.selection.active));
    return true;
  }
</script>

<Dialog title={t('dlgCreateTable')} width={340} {onok} bind:notice>
  <label class="lbl" for="ct-range">{t('dlgTableDataWhere')}</label>
  <input id="ct-range" class="xl-input range" bind:value={refText} spellcheck="false" />
  <label class="check"><input type="checkbox" bind:checked={hasHeaders} />{t('dlgTableHasHeaders')}</label>
</Dialog>

<style>
  .range {
    width: 100%;
    font-family: var(--xl-mono);
    margin: 4px 0 6px;
  }
  .lbl {
    display: block;
  }
</style>
