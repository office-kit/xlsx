<script lang="ts">
  // Format, the second contextual tab of a selected chart: fill and outline of
  // the chart area, and its size.
  import { editChart, setChartFill, setChartOutline } from '../core/chart-edit.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import ColorGrid from '../ui/ColorGrid.svelte';
  import MenuButton from '../ui/MenuButton.svelte';
  import DrawingSizeGroup from './DrawingSizeGroup.svelte';
  import Group from './Group.svelte';

  let { index }: { index: number } = $props();

  const doc = getEditor().doc;

  /** Excel's Shape Outline ▸ Weight choices, in points. */
  const WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];

  let outlineColor = $state<string>('595959');

  function currentOutlineWidth(): number {
    void doc.version;
    const item = doc.ws.drawing?.items[index];
    const chart = item?.content.kind === 'chart' ? item.content.chart : undefined;
    const w = (chart?.space?.spPr ?? chart?.cxSpace?.spPr)?.ln?.w;
    return w === undefined ? 0.75 : w / 12_700;
  }
</script>

<Group label={t('chGroupShapeStyles')}>
  <div class="col">
    <MenuButton icon="fill" label={t('chShapeFill')} title={t('chShapeFill')}>
      {#snippet menu(close)}
        <ColorGrid palette={doc.styles.palette} noneLabel={t('noFill')} onpick={(rgb) => { editChart(doc, index, 'Shape Fill', (c) => setChartFill(c, rgb)); close(); }} />
      {/snippet}
    </MenuButton>
    <MenuButton icon="borders" label={t('chShapeOutline')} title={t('chShapeOutline')}>
      {#snippet menu(close)}
        <ColorGrid
          palette={doc.styles.palette}
          noneLabel={t('chNoOutline')}
          onpick={(rgb) => {
            if (rgb) outlineColor = rgb;
            editChart(doc, index, 'Shape Outline', (c) => setChartOutline(c, rgb, currentOutlineWidth()));
            close();
          }}
        />
        <div class="xl-menu-sep"></div>
        <div class="head">{t('chWeight')}</div>
        {#each WEIGHTS as pt (pt)}
          <button class="xl-menu-item" onclick={() => { editChart(doc, index, 'Shape Outline', (c) => setChartOutline(c, outlineColor, pt)); close(); }}>
            <span class="line" style:height="{Math.max(1, pt * 1.33)}px"></span>{pt} pt
          </button>
        {/each}
      {/snippet}
    </MenuButton>
  </div>
</Group>

<DrawingSizeGroup {index} />

<style>
  .head {
    padding: 4px 10px 2px;
    font-size: 11.5px;
    font-weight: 600;
    color: var(--xl-text-2);
  }
  .line {
    display: inline-block;
    width: 60px;
    background: var(--xl-text);
  }
</style>
