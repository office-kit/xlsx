<script lang="ts">
  // Sparkline, the contextual tab of a cell holding a sparkline. Every
  // command edits the whole group, as Excel's tab does; Clear removes the
  // sparklines in the selected cells.
  import { getEditor } from '../core/context.ts';
  import { clearSparklines, setSparklineColor, setSparklineFlag, setSparklineType, SPARKLINE_FLAGS, type SparklineColorSlot } from '../core/sparklines.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import ColorGrid from '../ui/ColorGrid.svelte';
  import Icon from '../ui/Icon.svelte';
  import MenuButton from '../ui/MenuButton.svelte';
  import Group from './Group.svelte';

  let { groupIndex }: { groupIndex: number } = $props();

  const ctl = getEditor();
  const doc = ctl.doc;

  // Edits mutate the group in place, so read its fields here: a derived that
  // returned the group object itself would compare equal and not update.
  const view = $derived.by(() => {
    void doc.version;
    const g = doc.ws.sparklineGroups?.[groupIndex];
    return { type: g?.type ?? 'line', flags: new Set(SPARKLINE_FLAGS.filter((f) => g?.[f] === true)) };
  });
  const type = $derived(view.type);

  const TYPES = [
    { type: 'line', label: 'spkLine', icon: 'chart-line' },
    { type: 'column', label: 'spkColumn', icon: 'chart' },
    { type: 'stacked', label: 'spkWinLoss', icon: 'chart' },
  ] as const;

  const FLAG_LABELS: Record<(typeof SPARKLINE_FLAGS)[number], MessageKey> = {
    high: 'spkHigh',
    low: 'spkLow',
    negative: 'spkNegative',
    first: 'spkFirst',
    last: 'spkLast',
    markers: 'spkMarkers',
  };

  const MARKER_SLOTS: ReadonlyArray<{ slot: SparklineColorSlot; label: MessageKey }> = [
    { slot: 'colorNegative', label: 'spkNegative' },
    { slot: 'colorMarkers', label: 'spkMarkers' },
    { slot: 'colorHigh', label: 'spkHigh' },
    { slot: 'colorLow', label: 'spkLow' },
    { slot: 'colorFirst', label: 'spkFirst' },
    { slot: 'colorLast', label: 'spkLast' },
  ];

  let markerSlot = $state<SparklineColorSlot>('colorMarkers');
</script>

<Group label={t('spkGroupType')}>
  {#each TYPES as ty (ty.type)}
    <button class="xl-btn big" class:on={type === ty.type} aria-pressed={type === ty.type} onclick={() => setSparklineType(doc, groupIndex, ty.type)}>
      <Icon name={ty.icon} size={24} /><span>{t(ty.label)}</span>
    </button>
  {/each}
</Group>

<Group label={t('spkGroupShow')}>
  <div class="flags">
    {#each SPARKLINE_FLAGS as flag (flag)}
      <label class="check">
        <input type="checkbox" checked={view.flags.has(flag)} disabled={flag === 'markers' && type !== 'line'} onchange={(e) => setSparklineFlag(doc, groupIndex, flag, e.currentTarget.checked)} />
        {t(FLAG_LABELS[flag])}
      </label>
    {/each}
  </div>
</Group>

<Group label={t('spkGroupStyle')}>
  <div class="stack">
    <MenuButton icon="font-color" label={t('spkColor')} title={t('spkColor')}>
      {#snippet menu(close)}
        <ColorGrid palette={doc.styles.palette} noneLabel={t('spkColor')} onpick={(rgb) => { close(); if (rgb) setSparklineColor(doc, groupIndex, 'colorSeries', rgb); }} />
      {/snippet}
    </MenuButton>
    <MenuButton icon="fill" label={t('spkMarkerColor')} title={t('spkMarkerColor')}>
      {#snippet menu(close)}
        <div class="slots">
          {#each MARKER_SLOTS as m (m.slot)}
            <button class="xl-menu-item" class:sel={markerSlot === m.slot} onclick={() => (markerSlot = m.slot)}>{t(m.label)}</button>
          {/each}
        </div>
        <div class="xl-menu-sep"></div>
        <ColorGrid palette={doc.styles.palette} noneLabel={t(MARKER_SLOTS.find((m) => m.slot === markerSlot)?.label ?? 'spkMarkers')} onpick={(rgb) => { close(); if (rgb) setSparklineColor(doc, groupIndex, markerSlot, rgb); }} />
      {/snippet}
    </MenuButton>
  </div>
</Group>

<Group label={t('spkGroupGroup')}>
  <button class="xl-btn big" onclick={() => clearSparklines(doc, doc.selection.ranges)}><Icon name="eraser" size={24} /><span>{t('spkClear')}</span></button>
</Group>

<style>
  .big {
    flex-direction: column;
    height: 64px;
    min-width: 52px;
    font-size: 11px;
  }
  .on {
    background: var(--xl-hover);
    outline: 1px solid var(--xl-accent);
  }
  .flags {
    display: grid;
    grid-template-columns: auto auto;
    gap: 2px 12px;
    font-size: 11.5px;
  }
  .check {
    display: flex;
    align-items: center;
    gap: 4px;
    white-space: nowrap;
  }
  .stack {
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: flex-start;
    gap: 2px;
    height: 100%;
  }
  .slots {
    display: flex;
    flex-wrap: wrap;
    max-width: 240px;
  }
  .sel {
    font-weight: 600;
    color: var(--xl-accent);
  }
</style>
