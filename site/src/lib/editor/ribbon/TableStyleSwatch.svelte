<script lang="ts">
  // A miniature table drawn with a built-in table style, as in Excel's
  // Table Styles gallery: the current style options shape the preview.
  import type { TableDefinition } from '@office-kit/xlsx/worksheet';
  import type { StrokeStyle } from '../core/render-style.ts';
  import { tableLooks, type TableOptions } from '../core/table-style.ts';
  import type { ThemePalette } from '../core/theme.ts';

  let { style, palette, opts, header = true, total = false }: { style: string | undefined; palette: ThemePalette; opts: TableOptions; header?: boolean; total?: boolean } = $props();

  const ROWS = 5;
  const COLS = 5;

  const cells = $derived.by(() => {
    const def: TableDefinition = {
      id: 1,
      displayName: 'Preview',
      ref: `A1:E${ROWS}`,
      columns: [],
      ...(header ? {} : { headerRowCount: 0 }),
      ...(total ? { totalsRowCount: 1 } : {}),
      styleInfo: { ...(style ? { name: style } : {}), showRowStripes: opts.rowStripes, showColumnStripes: opts.colStripes, showFirstColumn: opts.firstCol, showLastColumn: opts.lastCol },
    };
    const look = tableLooks([def], palette);
    return Array.from({ length: ROWS * COLS }, (_, i) => look?.(Math.floor(i / COLS) + 1, (i % COLS) + 1));
  });

  const edge = (s: StrokeStyle | undefined) => (s ? `${s.double ? 2 : Math.min(s.width, 2)}px ${s.double ? 'double' : 'solid'} ${s.color}` : '0.5px solid #e6e6e6');
</script>

<span class="swatch" aria-hidden="true">
  {#each cells as look, i (i)}
    <span
      class="cell"
      style:background={look?.fill ?? '#fff'}
      style:border-top={edge(look?.top)}
      style:border-bottom={edge(look?.bottom)}
      style:border-left={edge(look?.left)}
      style:border-right={edge(look?.right)}
    >
      <span class="text" style:background={look?.color ?? '#000'} style:height={look?.bold ? '2px' : '1px'}></span>
    </span>
  {/each}
</span>

<style>
  .swatch {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    width: 52px;
    height: 38px;
    background: #fff;
  }
  .cell {
    display: flex;
    align-items: center;
    justify-content: center;
    box-sizing: border-box;
  }
  .text {
    width: 55%;
    opacity: 0.75;
  }
</style>
