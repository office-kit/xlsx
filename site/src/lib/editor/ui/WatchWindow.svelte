<script lang="ts">
  // Formulas ▸ Watch Window: a floating list of cells whose values stay in
  // view while you work elsewhere. Watches are a per-session view aid, as in
  // Excel, which doesn't save them in the file.
  import { fromStorageFormula } from '../calc/index.ts';
  import { cellAddress, MAX_COL, MAX_ROW, quoteSheetName } from '../core/address.ts';
  import { getCellAt } from '../core/cells.ts';
  import { getEditor } from '../core/context.ts';
  import { currentRange } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const ctl = getEditor();
  const doc = ctl.doc;
  let selected = $state<number | null>(null);
  let pos = $state({ x: 0, y: 0 });
  let drag: { dx: number; dy: number } | null = null;

  /** Excel caps a single Add Watch at a block this size. */
  const MAX_ADD = 1000;

  const rows = $derived.by(() => {
    void doc.version;
    // One pass over the names, not one per watch.
    const nameAt = new Map<string, string>();
    for (const d of doc.wb.definedNames) {
      const key = d.value.replace(/^=/, '').replaceAll('$', '').toUpperCase();
      if (!nameAt.has(key)) nameAt.set(key, d.name);
    }
    return ctl.watches.map((w) => {
      const ref = doc.wb.sheets.find((s) => s.sheet === w.sheet);
      const cell = ref?.kind === 'worksheet' ? getCellAt(ref.sheet, w.row, w.col) : undefined;
      const v = cell?.value;
      const formula = v !== null && v !== undefined && typeof v === 'object' && !(v instanceof Date) && v.kind === 'formula' ? `=${fromStorageFormula(v.formula)}` : '';
      const name = nameAt.get(`${quoteSheetName(w.sheet.title)}!${cellAddress(w.row, w.col)}`.toUpperCase()) ?? '';
      return { sheet: w.sheet.title, name, cell: cellAddress(w.row, w.col, true), value: cell ? ctl.displayText(cell) : '', formula };
    });
  });

  function addWatch() {
    const r = currentRange(doc.selection);
    const r2 = Math.min(r.r2, MAX_ROW);
    const c2 = Math.min(r.c2, MAX_COL);
    const next = [...ctl.watches];
    const have = new Set(next.filter((w) => w.sheet === doc.ws).map((w) => `${w.row}:${w.col}`));
    let added = 0;
    // A whole-column selection is a million cells: stop at the cap rather than walk it.
    for (let row = r.r1; row <= r2 && added < MAX_ADD; row++) {
      for (let col = r.c1; col <= c2 && added < MAX_ADD; col++) {
        const key = `${row}:${col}`;
        if (have.has(key)) continue;
        have.add(key);
        next.push({ sheet: doc.ws, row, col });
        added++;
      }
    }
    ctl.watches = next;
  }

  function deleteWatch() {
    if (selected === null) return;
    ctl.watches = ctl.watches.filter((_, i) => i !== selected);
    selected = null;
  }

  function go(i: number) {
    const w = ctl.watches[i];
    if (!w) return;
    const index = doc.wb.sheets.findIndex((s) => s.sheet === w.sheet);
    if (index >= 0 && index !== doc.activeSheetIndex) doc.activateSheet(index);
    ctl.selectCell({ row: w.row, col: w.col });
  }
</script>

<svelte:window
  onpointermove={(e) => {
    if (drag) pos = { x: e.clientX - drag.dx, y: e.clientY - drag.dy };
  }}
  onpointerup={() => (drag = null)}
/>

<div class="watch" role="dialog" aria-label={t('watchWindow')} style:transform="translate({pos.x}px, {pos.y}px)">
  <div
    class="title"
    role="presentation"
    onpointerdown={(e) => {
      drag = { dx: e.clientX - pos.x, dy: e.clientY - pos.y };
    }}
  >
    {t('watchWindow')}
    <button class="close" aria-label={t('dlgClose')} onclick={() => (ctl.watchWindow = false)}>×</button>
  </div>
  <div class="bar">
    <button class="xl-btn" onclick={addWatch}>{t('addWatch')}</button>
    <button class="xl-btn" disabled={selected === null} onclick={deleteWatch}>{t('deleteWatch')}</button>
  </div>
  <div class="table" role="grid">
    <div class="head" role="row">
      <span role="columnheader">{t('wwSheet')}</span><span role="columnheader">{t('wwName')}</span><span role="columnheader">{t('wwCell')}</span><span role="columnheader">{t('wwValue')}</span><span role="columnheader">{t('wwFormula')}</span>
    </div>
    {#each rows as r, i (i)}
      <button class="line" role="row" aria-selected={selected === i} onclick={() => (selected = i)} ondblclick={() => go(i)}>
        <span>{r.sheet}</span><span>{r.name}</span><span>{r.cell}</span><span>{r.value}</span><span>{r.formula}</span>
      </button>
    {/each}
  </div>
</div>

<style>
  .watch {
    position: fixed;
    right: 24px;
    bottom: 60px;
    z-index: 70;
    width: 520px;
    background: var(--xl-bg);
    border: 1px solid var(--xl-border);
    border-radius: 6px;
    box-shadow: 0 8px 24px rgb(0 0 0 / 0.2);
    font-size: 12px;
  }
  .title {
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding: 6px 10px;
    font-weight: 600;
    cursor: move;
    border-bottom: 1px solid var(--xl-border);
  }
  .close {
    border: 0;
    background: transparent;
    font-size: 16px;
    cursor: pointer;
  }
  .bar {
    display: flex;
    gap: 6px;
    padding: 6px 10px;
  }
  .table {
    max-height: 180px;
    overflow: auto;
    margin: 0 10px 10px;
    border: 1px solid var(--xl-border);
    background: #fff;
  }
  .head,
  .line {
    display: grid;
    grid-template-columns: 1fr 1fr 70px 1fr 1.4fr;
    gap: 6px;
    padding: 2px 6px;
    width: 100%;
    text-align: left;
  }
  .head {
    position: sticky;
    top: 0;
    background: #f3f3f3;
    font-weight: 600;
  }
  .line {
    border: 0;
    background: transparent;
    font: inherit;
  }
  .line span {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .line[aria-selected='true'] {
    background: #cfe3f7;
  }
</style>
