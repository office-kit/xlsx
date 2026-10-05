<script lang="ts">
  import { untrack } from 'svelte';
  // Column Width, Row Height and Standard Width share one form: a single
  // number in Excel's units (characters for widths, points for heights).
  import { getColumnDimension, getRowDimension } from '@office-kit/xlsx/worksheet';
  import { MAX_COL, MAX_ROW } from '../core/address.ts';
  import { setColumnWidths, setRowHeights } from '../core/commands.ts';
  import { getEditor } from '../core/context.ts';
  import { defaultColPx, defaultRowPx, displayColWidth, pxToColWidth, pxToPt, storedColWidth } from '../core/metrics.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  let { props }: { props: Record<string, unknown> | undefined } = $props();
  // DialogHost re-creates a dialog for every opening, so its arguments are read once.
  const args = untrack(() => props) ?? {};

  const MAX_WIDTH = 255;
  const MAX_HEIGHT = 409;
  /** Most lines one dialog run resizes individually (a whole-column selection spans a million rows). */
  const LINE_LIMIT = 100_000;

  const ctl = getEditor();
  const doc = ctl.doc;
  const ws = doc.ws;
  const kind = args['kind'] === 'rowHeight' ? 'rowHeight' : args['kind'] === 'standardWidth' ? 'standardWidth' : 'columnWidth';
  const active = doc.selection.active;
  const defaultWidth = displayColWidth(ws.defaultColumnWidth ?? pxToColWidth(defaultColPx(ws)));
  const initial =
    kind === 'rowHeight'
      ? (getRowDimension(ws, active.row)?.height ?? pxToPt(defaultRowPx(ws)))
      : kind === 'standardWidth'
        ? defaultWidth
        : (() => {
            const width = getColumnDimension(ws, active.col)?.width;
            return width === undefined ? defaultWidth : displayColWidth(width);
          })();
  let text = $state(String(Math.round(initial * 100) / 100));
  let notice = $state<string | null>(null);

  function lines(axis: 'row' | 'col'): number[] {
    const out = new Set<number>();
    for (const r of doc.selection.ranges) {
      const [lo, hi] = axis === 'row' ? [r.r1, r.r2] : [r.c1, r.c2];
      for (let i = lo; i <= hi && out.size < LINE_LIMIT; i++) out.add(i);
    }
    return [...out];
  }

  function onok(): boolean {
    const n = Number(text.trim());
    const max = kind === 'rowHeight' ? MAX_HEIGHT : MAX_WIDTH;
    if (text.trim() === '' || !Number.isFinite(n) || n < 0 || n > max) {
      notice = t(kind === 'rowHeight' ? 'dlgRowHeightRange' : 'dlgColumnWidthRange');
      return false;
    }
    if (kind === 'standardWidth') {
      doc.transact('Standard Width', (tx) => {
        tx.sheet(ws, 'defaultColumnWidth');
        ws.defaultColumnWidth = storedColWidth(n);
      });
    } else if (kind === 'columnWidth') {
      setColumnWidths(doc, lines('col'), storedColWidth(n));
    } else if (doc.selection.ranges.some((r) => r.r1 === 1 && r.r2 === MAX_ROW) && doc.selection.ranges.some((r) => r.c1 === 1 && r.c2 === MAX_COL)) {
      // The whole sheet: set the default height rather than a million rows.
      doc.transact('Row Height', (tx) => {
        tx.sheet(ws, 'defaultRowHeight', 'rowDimensions');
        ws.defaultRowHeight = n;
        for (const dim of ws.rowDimensions.values()) {
          delete dim.height;
          delete dim.customHeight;
        }
      });
    } else {
      setRowHeights(doc, lines('row'), n);
    }
    return true;
  }
</script>

<Dialog title={t(kind === 'rowHeight' ? 'dlgRowHeight' : kind === 'standardWidth' ? 'dlgStandardWidth' : 'dlgColumnWidth')} width={280} {onok} bind:notice>
  <div class="row">
    <label for="size-v">{t(kind === 'rowHeight' ? 'dlgRowHeightLabel' : kind === 'standardWidth' ? 'dlgStandardWidthLabel' : 'dlgColumnWidthLabel')}</label>
    <input id="size-v" class="xl-input grow" bind:value={text} inputmode="decimal" />
  </div>
</Dialog>
