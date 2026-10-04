<script lang="ts">
  // Mode indicator, selection statistics (Average / Count / Sum …), circular
  // reference warning and the zoom slider.
  import { getEditor } from '../core/context.ts';
  import { selectionStats } from '../core/stats.ts';
  import { setSheetViewMode } from '../core/actions.ts';
  import { cellAddress } from '../core/address.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Icon from './Icon.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;

  const mode = $derived(ctl.edit ? (ctl.edit.mode === 'point' ? t('modePoint') : ctl.edit.mode === 'edit' ? t('modeEdit') : t('modeEnter')) : t('modeReady'));

  // Statistics are recomputed after the selection settles so dragging a
  // selection over a huge sheet stays smooth.
  let stats = $state<ReturnType<typeof selectionStats> | null>(null);
  $effect(() => {
    void doc.version;
    const sel = doc.selection;
    const ws = doc.ws;
    const id = setTimeout(() => (stats = selectionStats(ws, sel)), 60);
    return () => clearTimeout(id);
  });

  const fmt = (n: number) => (Number.isInteger(n) ? n.toLocaleString() : Number(n.toPrecision(10)).toLocaleString(undefined, { maximumFractionDigits: 10 }));

  const circular = $derived.by(() => {
    void doc.version;
    const c = doc.calc.circularRefs[0];
    return c ? `${c.sheet === doc.ws.title ? '' : c.sheet + '!'}${cellAddress(c.row, c.col)}` : null;
  });

  $effect(() => {
    if (!ctl.toast) return;
    const id = setTimeout(() => (ctl.toast = null), 3500);
    return () => clearTimeout(id);
  });

  const zoomPct = $derived(Math.round(doc.zoom * 100));
  const VIEWS: Array<['normal' | 'pageLayout' | 'pageBreakPreview', string, MessageKey]> = [
    ['normal', 'sheet', 'viewNormal'],
    ['pageLayout', 'page-layout', 'viewPageLayout'],
    ['pageBreakPreview', 'page', 'viewPageBreak'],
  ];

  // Right-click on the bar picks the aggregates, as in Excel; remembered per browser.
  type Agg = 'average' | 'count' | 'numCount' | 'min' | 'max' | 'sum';
  const AGGS: Array<[Agg, MessageKey]> = [
    ['average', 'statAverage'],
    ['count', 'statCount'],
    ['numCount', 'statNumCount'],
    ['min', 'statMin'],
    ['max', 'statMax'],
    ['sum', 'statSum'],
  ];
  const STORE_KEY = 'xlsx-editor.status-aggregates';
  function loadAggs(): Set<Agg> {
    try {
      const raw = localStorage.getItem(STORE_KEY);
      if (raw) return new Set(raw.split(',').filter((a): a is Agg => AGGS.some(([k]) => k === a)));
    } catch {
      // Storage can be unavailable (private mode); fall back to Excel's defaults.
    }
    return new Set<Agg>(['average', 'count', 'sum']);
  }
  let shown = $state(loadAggs());
  let menu = $state<{ x: number; y: number } | null>(null);
  function toggleAgg(a: Agg) {
    const next = new Set(shown);
    if (next.has(a)) next.delete(a);
    else next.add(a);
    shown = next;
    try {
      localStorage.setItem(STORE_KEY, [...next].join(','));
    } catch {
      // Not persisting is fine; the choice still applies to this session.
    }
  }
  function aggValue(a: Agg, s: NonNullable<typeof stats>): string | undefined {
    switch (a) {
      case 'average':
        return s.numCount > 0 ? fmt(s.sum / s.numCount) : undefined;
      case 'count':
        return s.count.toLocaleString();
      case 'numCount':
        return s.numCount.toLocaleString();
      case 'min':
        return s.numCount > 0 ? fmt(s.min) : undefined;
      case 'max':
        return s.numCount > 0 ? fmt(s.max) : undefined;
      case 'sum':
        return s.numCount > 0 ? fmt(s.sum) : undefined;
    }
  }
</script>

<div
  class="status"
  role="group"
  aria-label="Status bar"
  oncontextmenu={(e) => {
    e.preventDefault();
    menu = { x: e.clientX, y: e.clientY };
  }}
>
  <span class="mode">{mode}</span>
  {#if circular}<span class="warn">{t('circularReferences')}: {circular}</span>{/if}
  {#if ctl.clipboard}<span class="hint">{t('selectDestination')}</span>{/if}
  {#if ctl.toast}<span class="toast" role="status">{t(ctl.toast)}</span>{/if}
  <span class="spacer"></span>
  {#if stats && stats.count > 1}
    {#each AGGS as [key, label] (key)}
      {@const value = shown.has(key) ? aggValue(key, stats) : undefined}
      {#if value !== undefined}<span>{t(label)}: {value}</span>{/if}
    {/each}
  {/if}
  {#if menu}
    <div class="xl-menu agg-menu" style:left="{menu.x}px" style:bottom="{window.innerHeight - menu.y}px" role="menu">
      {#each AGGS as [key, label] (key)}
        <button class="xl-menu-item" role="menuitemcheckbox" aria-checked={shown.has(key)} onclick={() => toggleAgg(key)}>
          <span class="check">{shown.has(key) ? '✓' : ''}</span>{t(label)}
        </button>
      {/each}
    </div>
    <button class="agg-backdrop" aria-label="Close" onclick={() => (menu = null)}></button>
  {/if}
  <span class="views">
    {#each VIEWS as [mode, icon, label] (mode)}
      <button class="xl-btn" title={t(label)} aria-label={t(label)} aria-pressed={(ctl.sheetView?.view ?? 'normal') === mode} onclick={() => setSheetViewMode(ctl, mode)}><Icon name={icon} size={14} /></button>
    {/each}
  </span>
  <span class="zoom">
    <button class="xl-btn" title={t('zoomOut')} onclick={() => doc.setZoom(Math.round(doc.zoom * 10 - 1) / 10)}><Icon name="minus" size={12} /></button>
    <input type="range" min="10" max="400" step="5" value={zoomPct} aria-label={t('zoom')} oninput={(e) => doc.setZoom(Number((e.currentTarget as HTMLInputElement).value) / 100)} />
    <button class="xl-btn" title={t('zoomIn')} onclick={() => doc.setZoom(Math.round(doc.zoom * 10 + 1) / 10)}><Icon name="plus" size={12} /></button>
    <button class="xl-btn pct" onclick={() => ctl.openDialog('zoom')}>{zoomPct}%</button>
  </span>
</div>

<style>
  .status {
    display: flex;
    align-items: center;
    gap: 16px;
    height: 24px;
    padding: 0 8px;
    background: var(--xl-bg);
    border-top: 1px solid var(--xl-border);
    font-size: 11.5px;
    color: var(--xl-text-2);
    flex-shrink: 0;
    white-space: nowrap;
    overflow: hidden;
  }
  .spacer {
    flex: 1;
  }
  .warn {
    color: var(--xl-danger);
  }
  .toast {
    color: var(--xl-text);
    background: #fff8d6;
    padding: 1px 8px;
    border-radius: 3px;
  }
  .views {
    display: flex;
    gap: 1px;
  }
  .views [aria-pressed='true'] {
    background: var(--xl-hover, #e1e1e1);
  }
  .zoom {
    display: flex;
    align-items: center;
    gap: 2px;
  }
  .zoom input {
    width: 100px;
    accent-color: var(--xl-accent);
  }
  .pct {
    min-width: 46px;
  }
  .agg-menu {
    position: fixed;
    z-index: 61;
    min-width: 180px;
  }
  .agg-menu .check {
    display: inline-block;
    width: 14px;
  }
  .agg-backdrop {
    position: fixed;
    inset: 0;
    z-index: 60;
    border: 0;
    background: transparent;
  }
</style>
