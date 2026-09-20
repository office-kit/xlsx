<script lang="ts">
  import { tick } from 'svelte';
  import type { GridCell, GridEdge, SheetGridData } from '$lib/sheet-grid';

  type Props = {
    sheets: SheetGridData[];
  };

  const { sheets }: Props = $props();

  type Position = { row: number; col: number };

  let activeIndex = $state(0);
  let picked = $state<Position | undefined>();
  // eslint-disable-next-line prefer-const -- reassigned by `bind:this` in template
  let body = $state<HTMLTableSectionElement | undefined>();

  // The sheets are replaced on every REPL run, so the tab that was open may be gone.
  const sheet = $derived(sheets[Math.min(activeIndex, sheets.length - 1)]);

  const cellAt = (grid: SheetGridData, at: Position): GridCell | undefined =>
    grid.rows[at.row]?.cells[at.col];

  // Until the reader picks a cell, the formula bar describes the first formula,
  // which is the cell with the most to say.
  const selected = $derived.by((): Position | undefined => {
    if (!sheet) return undefined;
    if (picked && cellAt(sheet, picked)?.covered === false) return picked;
    for (const [row, { cells }] of sheet.rows.entries()) {
      const col = cells.findIndex((cell) => cell.formula !== undefined);
      if (col >= 0) return { row, col };
    }
    return cellAt(sheet, { row: 0, col: 0 }) ? { row: 0, col: 0 } : undefined;
  });
  const selectedCell = $derived(sheet && selected ? cellAt(sheet, selected) : undefined);
  const selectedRef = $derived(
    sheet && selected ? `${sheet.columns[selected.col]?.letter}${sheet.rows[selected.row]?.number}` : '',
  );

  function openSheet(index: number): void {
    activeIndex = index;
    picked = undefined;
  }

  const positionOf = (target: EventTarget | null): Position | undefined => {
    const td = target instanceof Element ? target.closest<HTMLElement>('td[data-row]') : null;
    return td ? { row: Number(td.dataset.row), col: Number(td.dataset.col) } : undefined;
  };

  const STEPS: Record<string, Position> = {
    ArrowUp: { row: -1, col: 0 },
    ArrowDown: { row: 1, col: 0 },
    ArrowLeft: { row: 0, col: -1 },
    ArrowRight: { row: 0, col: 1 },
  };

  async function onKeydown(event: KeyboardEvent): Promise<void> {
    const step = STEPS[event.key];
    const from = positionOf(event.target);
    if (!step || !from || !sheet) return;
    // Walk past cells a merge covers; stop at the edge of the grid.
    let next = { row: from.row + step.row, col: from.col + step.col };
    while (cellAt(sheet, next)?.covered) {
      next = { row: next.row + step.row, col: next.col + step.col };
    }
    if (!cellAt(sheet, next)) return;
    event.preventDefault();
    picked = next;
    await tick();
    body?.querySelector<HTMLElement>(`td[data-row="${next.row}"][data-col="${next.col}"]`)?.focus();
  }

  // Inset shadows draw a cell's own borders without disturbing the grid lines
  // or the size of the cell.
  function edgeShadows(edges: GridCell['edges']): string | undefined {
    const shadow = (x: number, y: number, e: GridEdge): string =>
      `inset ${x * e.width}px ${y * e.width}px 0 ${e.color ?? 'var(--sheet-ink)'}`;
    const parts = [
      edges.top && shadow(0, 1, edges.top),
      edges.right && shadow(-1, 0, edges.right),
      edges.bottom && shadow(0, -1, edges.bottom),
      edges.left && shadow(1, 0, edges.left),
    ].filter((part) => part !== undefined);
    return parts.length > 0 ? parts.join(', ') : undefined;
  }
</script>

{#if sheet}
  <div class="sheet">
    <div class="formula-bar">
      <span class="name-box">{selectedRef}</span>
      <span class="fx" aria-hidden="true">fx</span>
      <span class="formula">{selectedCell?.formula ?? selectedCell?.text ?? ''}</span>
    </div>
    <!-- A scrollable region must be focusable, or keyboard users cannot scroll it. -->
    <!-- svelte-ignore a11y_no_noninteractive_tabindex -->
    <div class="sheet-scroll" tabindex="0" role="region" aria-label="Cells of {sheet.name}">
      <table role="grid">
        <colgroup>
          <col class="gutter-col" />
          {#each sheet.columns as column (column.letter)}
            <col class="data-col" style:--col-w="{column.width}px" />
          {/each}
        </colgroup>
        <thead>
          <tr>
            <td></td>
            {#each sheet.columns as column (column.letter)}
              <th scope="col">{column.letter}</th>
            {/each}
          </tr>
        </thead>
        <!-- One listener pair for the whole grid instead of one per cell: a
             previewed sheet can hold thousands of cells. -->
        <!-- svelte-ignore a11y_click_events_have_key_events, a11y_no_noninteractive_element_interactions -->
        <tbody
          bind:this={body}
          onclick={(event) => (picked = positionOf(event.target) ?? picked)}
          onkeydown={onKeydown}
        >
          {#each sheet.rows as row, r (row.number)}
            <tr class:frozen-row={r + 1 === sheet.frozenRows}>
              <th scope="row">{row.number}</th>
              {#each row.cells as cell, c (c)}
                {#if !cell.covered}
                  {@const isSelected = selected?.row === r && selected.col === c}
                  <td
                    role="gridcell"
                    data-row={r}
                    data-col={c}
                    tabindex={isSelected ? 0 : -1}
                    aria-selected={isSelected}
                    rowspan={cell.rowSpan}
                    colspan={cell.colSpan}
                    class:bold={cell.bold}
                    class:italic={cell.italic}
                    class:frozen-col={c + cell.colSpan === sheet.frozenCols}
                    style:text-align={cell.align}
                    style:color={cell.color}
                    style:background-color={cell.fill}
                    style:box-shadow={edgeShadows(cell.edges)}
                  >
                    {cell.text}
                  </td>
                {/if}
              {/each}
            </tr>
          {/each}
        </tbody>
      </table>
    </div>
    <div class="sheet-tabs" role="tablist" aria-label="Worksheets">
      {#each sheets as tab, index (index)}
        <button
          type="button"
          role="tab"
          aria-selected={tab === sheet}
          onclick={() => openSheet(index)}
        >
          {tab.name}
        </button>
      {/each}
    </div>
  </div>
{/if}

<style>
  /* A worksheet carries its own fills and font colours, chosen against white
   * paper, so like a slide or a page it stays light in both themes. */
  .sheet {
    --sheet-paper: #ffffff;
    --sheet-chrome: #f3f4f6;
    --sheet-rule: #e2e4e9;
    --sheet-rule-strong: #c9cdd6;
    --sheet-ink: #15171c;
    --sheet-ink-2: #5b616e;
    --sheet-select: #168a4f;
    --sheet-row-h: 28px;
    --sheet-gutter-w: 38px;
    --sheet-scale: 1;

    display: flex;
    flex-direction: column;
    height: 100%;
    min-height: 0;
    background: var(--sheet-paper);
    color: var(--sheet-ink);
    font-size: 0.84rem;
    line-height: 1.2;
    overflow: hidden;
  }

  .formula-bar {
    flex: none;
    display: flex;
    align-items: center;
    gap: 0.6rem;
    padding: 0.45rem 0.6rem;
    border-bottom: 1px solid var(--sheet-rule-strong);
    background: var(--sheet-chrome);
    font-family: var(--mono);
    font-size: 0.78rem;
  }

  .name-box,
  .formula {
    min-height: 1.85rem;
    padding: 0.35rem 0.55rem;
    border: 1px solid var(--sheet-rule-strong);
    border-radius: 3px;
    background: var(--sheet-paper);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  .name-box {
    flex: none;
    min-width: 3.25rem;
  }

  .formula {
    flex: 1;
    min-width: 0;
  }

  .fx {
    color: var(--sheet-ink-2);
    font-style: italic;
  }

  .sheet-scroll {
    flex: 1;
    min-height: 0;
    max-height: var(--sheet-max-height, none);
    overflow: auto;
  }

  table {
    width: max-content;
    min-width: 100%;
    margin: 0;
    table-layout: fixed;
    border-collapse: separate;
    border-spacing: 0;
    font-size: inherit;
    font-variant-numeric: tabular-nums;
  }

  .gutter-col {
    width: var(--sheet-gutter-w);
  }

  .data-col {
    width: calc(var(--col-w) * var(--sheet-scale));
  }

  th,
  td {
    height: var(--sheet-row-h);
    padding: 0 0.45rem;
    border: none;
    border-right: 1px solid var(--sheet-rule);
    border-bottom: 1px solid var(--sheet-rule);
    vertical-align: middle;
    white-space: nowrap;
    overflow: hidden;
  }

  th,
  thead td {
    background: var(--sheet-chrome);
    border-color: var(--sheet-rule-strong);
    color: var(--sheet-ink-2);
    font-size: 0.74rem;
    font-weight: 500;
    text-align: center;
  }

  /* Headers stay in view while the cells scroll under them, as in Excel. */
  thead th,
  thead td {
    position: sticky;
    top: 0;
    z-index: 2;
  }

  tbody th {
    position: sticky;
    left: 0;
    z-index: 1;
  }

  thead td {
    left: 0;
    z-index: 3;
  }

  td {
    cursor: cell;
  }

  td.bold {
    font-weight: 700;
  }

  td.italic {
    font-style: italic;
  }

  /* Freeze panes: Excel marks the frozen edge with a darker rule. */
  .frozen-row > * {
    border-bottom-color: var(--sheet-ink-2);
  }

  td.frozen-col {
    border-right-color: var(--sheet-ink-2);
  }

  td[aria-selected='true'] {
    outline: 2px solid var(--sheet-select);
    outline-offset: -2px;
  }

  td:focus-visible {
    outline: 2px solid var(--sheet-select);
    outline-offset: -2px;
    border-radius: 0;
  }

  .sheet-tabs {
    flex: none;
    display: flex;
    gap: 0.15rem;
    padding: 0 0.6rem;
    border-top: 1px solid var(--sheet-rule-strong);
    background: var(--sheet-chrome);
    overflow-x: auto;
  }

  .sheet-tabs button {
    flex: none;
    min-height: 32px;
    margin-top: -1px;
    padding: 0 0.9rem;
    border: 1px solid transparent;
    border-bottom: 2px solid transparent;
    background: none;
    color: var(--sheet-ink-2);
    font-family: var(--sans);
    font-size: 0.78rem;
    font-weight: 550;
    cursor: pointer;
  }

  .sheet-tabs button:hover {
    color: var(--sheet-ink);
  }

  .sheet-tabs button[aria-selected='true'] {
    border-color: var(--sheet-rule-strong);
    border-top-color: var(--sheet-paper);
    border-bottom-color: var(--sheet-select);
    background: var(--sheet-paper);
    color: var(--sheet-ink);
  }

  .sheet :focus-visible {
    outline-color: var(--sheet-select);
  }

  .sheet-tabs button:focus-visible {
    outline-offset: -2px;
  }

  @media (max-width: 520px) {
    /* Scaled down so a four-column table fits a phone without scrolling. */
    .sheet {
      --sheet-scale: 0.8;
      --sheet-gutter-w: 30px;
      font-size: 0.78rem;
    }

    th,
    td {
      padding: 0 0.35rem;
    }
  }
</style>
