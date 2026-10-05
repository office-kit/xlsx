<script lang="ts">
  import { untrack } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { getCellAt } from '../core/cells.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';
  import { applyFormatCells, FormatSampler, initialState, numberCode, sampleValue } from './format-cells.ts';
  import AlignmentTab from './format-cells/AlignmentTab.svelte';
  import BorderTab from './format-cells/BorderTab.svelte';
  import FillTab from './format-cells/FillTab.svelte';
  import FontTab from './format-cells/FontTab.svelte';
  import NumberTab from './format-cells/NumberTab.svelte';
  import ProtectionTab from './format-cells/ProtectionTab.svelte';

  let { props }: { props: Record<string, unknown> | undefined } = $props();
  // DialogHost re-creates a dialog for every opening, so its arguments are read once.
  const args = untrack(() => props) ?? {};

  const TABS = [
    { id: 'number', label: 'dlgTabNumber' },
    { id: 'alignment', label: 'dlgTabAlignment' },
    { id: 'font', label: 'dlgTabFont' },
    { id: 'border', label: 'dlgTabBorder' },
    { id: 'fill', label: 'dlgTabFill' },
    { id: 'protection', label: 'dlgTabProtection' },
  ] as const satisfies ReadonlyArray<{ id: string; label: MessageKey }>;
  type TabId = (typeof TABS)[number]['id'];

  const ctl = getEditor();
  const start = initialState(ctl);
  const fmt = $state(structuredClone(start));
  const requested = args['tab'];
  let tab = $state<TabId>(TABS.find((x) => x.id === requested)?.id ?? 'number');
  let notice = $state<string | null>(null);

  const sampler = new FormatSampler(ctl.doc.wb.date1904);
  const active = ctl.doc.selection.active;
  const sample = sampleValue(getCellAt(ctl.doc.ws, active.row, active.col));

  function onok(): boolean {
    if (sampler.format(numberCode(fmt.number), 1234.5) === undefined) {
      tab = 'number';
      notice = t('dlgBadNumberFormat');
      return false;
    }
    applyFormatCells(ctl, start, fmt);
    return true;
  }

  function onTabKey(e: KeyboardEvent, i: number): void {
    if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
    const next = TABS[(i + (e.key === 'ArrowRight' ? 1 : TABS.length - 1)) % TABS.length];
    if (!next) return;
    tab = next.id;
    (e.currentTarget as HTMLElement).parentElement?.querySelector<HTMLElement>(`[data-tab="${next.id}"]`)?.focus();
  }
</script>

<Dialog title={t('dlgFormatCells')} width={560} {onok} bind:notice>
  <div class="seg" role="tablist">
    {#each TABS as item, i (item.id)}
      <button
        role="tab"
        data-tab={item.id}
        aria-selected={tab === item.id}
        tabindex={tab === item.id ? 0 : -1}
        onclick={() => (tab = item.id)}
        onkeydown={(e) => onTabKey(e, i)}>{t(item.label)}</button
      >
    {/each}
  </div>
  <div class="panel" role="tabpanel">
    {#if tab === 'number'}
      <NumberTab bind:model={fmt.number} {sampler} {sample} />
    {:else if tab === 'alignment'}
      <AlignmentTab bind:model={fmt.align} />
    {:else if tab === 'font'}
      <FontTab bind:model={fmt.font} />
    {:else if tab === 'border'}
      <BorderTab bind:model={fmt.border} />
    {:else if tab === 'fill'}
      <FillTab bind:model={fmt.fill} />
    {:else}
      <ProtectionTab bind:model={fmt.protection} />
    {/if}
  </div>
</Dialog>

<style>
  .panel {
    min-height: 330px;
  }
</style>
