<script lang="ts">
  import type { Side } from '@office-kit/xlsx/styles';
  import { makeColor, makeSide } from '@office-kit/xlsx/styles';
  import { getEditor } from '../../core/context.ts';
  import { resolveColor } from '../../core/theme.ts';
  import { t, type MessageKey } from '../../i18n/i18n.svelte.ts';
  import ColorPicker from '../ColorPicker.svelte';
  import { EDGES, LINE_STYLES, strokeOf, type BorderState, type Edge } from '../format-cells.ts';

  let { model = $bindable() }: { model: BorderState } = $props();

  const ctl = getEditor();
  const EDGE_LABELS: Record<Edge, MessageKey> = {
    top: 'borderTop',
    insideH: 'dlgBorderInsideH',
    bottom: 'borderBottom',
    left: 'borderLeft',
    insideV: 'dlgBorderInsideV',
    right: 'borderRight',
    diagUp: 'dlgBorderDiagUp',
    diagDown: 'dlgBorderDiagDown',
  };
  // Preview geometry: a 2×2 block of cells inside a 200×120 box.
  const X0 = 20;
  const X1 = 180;
  const Y0 = 12;
  const Y1 = 108;
  const LINES: Record<Edge, [number, number, number, number]> = {
    top: [X0, Y0, X1, Y0],
    insideH: [X0, (Y0 + Y1) / 2, X1, (Y0 + Y1) / 2],
    bottom: [X0, Y1, X1, Y1],
    left: [X0, Y0, X0, Y1],
    insideV: [(X0 + X1) / 2, Y0, (X0 + X1) / 2, Y1],
    right: [X1, Y0, X1, Y1],
    diagUp: [X0, Y1, X1, Y0],
    diagDown: [X0, Y0, X1, Y1],
  };

  function currentSide(): Side {
    return makeSide({ style: model.lineStyle, ...(model.lineColor ? { color: makeColor({ rgb: model.lineColor }) } : {}) });
  }

  function set(edge: Edge, side: Side | null): void {
    model.edges[edge] = side;
    if (!model.touched.includes(edge)) model.touched.push(edge);
  }

  function toggle(edge: Edge): void {
    set(edge, model.edges[edge] ? null : currentSide());
  }

  function preset(kind: 'none' | 'outline' | 'inside'): void {
    if (kind === 'none') for (const e of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'] as const) set(e, null);
    else if (kind === 'outline') for (const e of ['top', 'bottom', 'left', 'right'] as const) set(e, currentSide());
    else for (const e of ['insideH', 'insideV'] as const) set(e, currentSide());
  }

  function strokeColor(side: Side): string {
    return resolveColor(side.color, ctl.doc.styles.palette, '#000000');
  }
</script>

{#snippet edgeButton(edge: Edge)}
  <button class="xl-btn outlined edge" aria-pressed={model.edges[edge] !== null} title={t(EDGE_LABELS[edge])} aria-label={t(EDGE_LABELS[edge])} onclick={() => toggle(edge)}>
    <svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true">
      <rect x="2" y="2" width="12" height="12" fill="none" stroke="#bbb" stroke-dasharray="1 1" />
      {#if edge === 'top'}<line x1="2" y1="2" x2="14" y2="2" stroke="#000" stroke-width="2" />
      {:else if edge === 'insideH'}<line x1="2" y1="8" x2="14" y2="8" stroke="#000" stroke-width="2" />
      {:else if edge === 'bottom'}<line x1="2" y1="14" x2="14" y2="14" stroke="#000" stroke-width="2" />
      {:else if edge === 'left'}<line x1="2" y1="2" x2="2" y2="14" stroke="#000" stroke-width="2" />
      {:else if edge === 'insideV'}<line x1="8" y1="2" x2="8" y2="14" stroke="#000" stroke-width="2" />
      {:else if edge === 'right'}<line x1="14" y1="2" x2="14" y2="14" stroke="#000" stroke-width="2" />
      {:else if edge === 'diagUp'}<line x1="2" y1="14" x2="14" y2="2" stroke="#000" stroke-width="1.5" />
      {:else}<line x1="2" y1="2" x2="14" y2="14" stroke="#000" stroke-width="1.5" />{/if}
    </svg>
  </button>
{/snippet}

<div class="border">
  <div class="main">
    <div class="presets">
      <button class="xl-btn outlined" onclick={() => preset('none')}>{t('dlgNone')}</button>
      <button class="xl-btn outlined" onclick={() => preset('outline')}>{t('dlgOutline')}</button>
      <button class="xl-btn outlined" onclick={() => preset('inside')}>{t('dlgInside')}</button>
    </div>
    <div class="lbl">{t('borders')}</div>
    <div class="board">
      <div class="side-buttons">
        {@render edgeButton('top')}
        {@render edgeButton('insideH')}
        {@render edgeButton('bottom')}
      </div>
      <svg class="preview" viewBox="0 0 200 120" role="img" aria-label={t('dlgPreview')}>
        <rect x={X0} y={Y0} width={X1 - X0} height={Y1 - Y0} fill="#fff" />
        {#each [0, 1] as r (r)}
          {#each [0, 1] as c (c)}
            <text x={X0 + 40 + c * 80} y={Y0 + 28 + r * 48} font-size="11" text-anchor="middle" fill="#666">{t('dlgText')}</text>
          {/each}
        {/each}
        {#each EDGES as edge (edge)}
          {@const [x1, y1, x2, y2] = LINES[edge]}
          {@const side = model.edges[edge]}
          {#if side}
            {@const s = strokeOf(side.style)}
            {#if s.double}
              {@const dx = y1 === y2 ? 0 : 1.5}
              {@const dy = y1 === y2 ? 1.5 : 0}
              <line x1={x1 - dx} y1={y1 - dy} x2={x2 - dx} y2={y2 - dy} stroke={strokeColor(side)} stroke-width="1" />
              <line x1={x1 + dx} y1={y1 + dy} x2={x2 + dx} y2={y2 + dy} stroke={strokeColor(side)} stroke-width="1" />
            {:else}
              <line {x1} {y1} {x2} {y2} stroke={strokeColor(side)} stroke-width={s.width} stroke-dasharray={s.dash || undefined} />
            {/if}
          {/if}
        {/each}
      </svg>
    </div>
    <div class="bottom-buttons">
      {@render edgeButton('diagUp')}
      {@render edgeButton('left')}
      {@render edgeButton('insideV')}
      {@render edgeButton('right')}
      {@render edgeButton('diagDown')}
    </div>
    <p class="hint">{t('dlgBorderHint')}</p>
  </div>
  <div class="line">
    <fieldset>
      <legend>{t('dlgLine')}</legend>
      <div class="lbl">{t('dlgStyle')}</div>
      <div class="styles" role="listbox" aria-label={t('dlgStyle')}>
        {#each LINE_STYLES as style (style)}
          {@const s = strokeOf(style)}
          <button role="option" aria-selected={model.lineStyle === style} aria-label={style} onclick={() => (model.lineStyle = style)}>
            <svg viewBox="0 0 60 8" width="60" height="8" aria-hidden="true">
              {#if s.double}
                <line x1="2" y1="2.5" x2="58" y2="2.5" stroke="currentColor" />
                <line x1="2" y1="5.5" x2="58" y2="5.5" stroke="currentColor" />
              {:else}
                <line x1="2" y1="4" x2="58" y2="4" stroke="currentColor" stroke-width={s.width} stroke-dasharray={s.dash || undefined} />
              {/if}
            </svg>
          </button>
        {/each}
      </div>
      <div class="lbl">{t('dlgColor')}</div>
      <ColorPicker bind:value={model.lineColor} noneLabel={t('automatic')} label={t('dlgColor')} />
    </fieldset>
  </div>
</div>

<style>
  .border {
    display: flex;
    gap: 12px;
  }
  .main {
    flex: 1;
  }
  .presets {
    display: flex;
    gap: 8px;
    margin-bottom: 6px;
  }
  .board {
    display: flex;
    gap: 6px;
    align-items: stretch;
  }
  .side-buttons {
    display: flex;
    flex-direction: column;
    justify-content: space-between;
  }
  .bottom-buttons {
    display: flex;
    justify-content: space-between;
    margin: 4px 0 0 36px;
    width: 220px;
  }
  .edge {
    min-height: 26px;
    padding: 3px;
  }
  .preview {
    width: 220px;
    height: 132px;
    background: #fafafa;
    border: 1px solid var(--xl-border);
  }
  .line {
    width: 170px;
  }
  .styles {
    display: grid;
    grid-template-columns: 1fr 1fr;
    border: 1px solid var(--xl-border-strong);
    background: #fff;
    border-radius: 3px;
    margin-bottom: 6px;
  }
  .styles button {
    border: 0;
    background: transparent;
    padding: 4px;
    color: #000;
    cursor: default;
  }
  .styles button[aria-selected='true'] {
    background: var(--xl-selected);
    outline: 1px solid var(--xl-accent);
  }
</style>
