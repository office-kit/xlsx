<script lang="ts">
  import * as A from '../core/actions.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import MenuButton from '../ui/MenuButton.svelte';
  import Group from './Group.svelte';

  const ctl = getEditor();
  const setup = $derived.by(() => {
    void ctl.doc.version;
    return ctl.doc.ws.pageSetup;
  });
  const print = $derived.by(() => {
    void ctl.doc.version;
    return ctl.doc.ws.printOptions;
  });
</script>

<Group label={t('groupPageSetup')}>
  <MenuButton large icon="margins" label={t('margins')}>
    {#snippet menu(close)}
      <button class="xl-menu-item" onclick={() => { A.setMargins(ctl, 'normal'); close(); }}>{t('marginsNormal')}</button>
      <button class="xl-menu-item" onclick={() => { A.setMargins(ctl, 'wide'); close(); }}>{t('marginsWide')}</button>
      <button class="xl-menu-item" onclick={() => { A.setMargins(ctl, 'narrow'); close(); }}>{t('marginsNarrow')}</button>
      <div class="xl-menu-sep"></div>
      <button class="xl-menu-item" onclick={() => { ctl.openDialog('pageSetup', { tab: 'margins' }); close(); }}>{t('customMargins')}</button>
    {/snippet}
  </MenuButton>
  <MenuButton large icon={setup?.orientation === 'landscape' ? 'landscape' : 'portrait'} label={t('orientation')}>
    {#snippet menu(close)}
      <button class="xl-menu-item" onclick={() => { A.setOrientation(ctl, 'portrait'); close(); }}><Icon name="portrait" size={16} />{t('portrait')}</button>
      <button class="xl-menu-item" onclick={() => { A.setOrientation(ctl, 'landscape'); close(); }}><Icon name="landscape" size={16} />{t('landscape')}</button>
    {/snippet}
  </MenuButton>
  <MenuButton large icon="page" label={t('size')}>
    {#snippet menu(close)}
      {#each [[1, 'Letter'], [5, 'Legal'], [9, 'A4'], [8, 'A3'], [11, 'A5'], [13, 'B5 (JIS)']] as const as [id, name] (id)}
        <button class="xl-menu-item" onclick={() => { A.setPaperSize(ctl, id); close(); }}>{setup?.paperSize === id ? '✓ ' : ''}{name}</button>
      {/each}
    {/snippet}
  </MenuButton>
  <MenuButton large icon="print" label={t('printArea')}>
    {#snippet menu(close)}
      <button class="xl-menu-item" onclick={() => { A.setPrintArea(ctl, true); close(); }}>{t('setPrintArea')}</button>
      <button class="xl-menu-item" onclick={() => { A.setPrintArea(ctl, false); close(); }}>{t('clearPrintArea')}</button>
    {/snippet}
  </MenuButton>
  <MenuButton large icon="page" label={t('breaks')}>
    {#snippet menu(close)}
      <button class="xl-menu-item" onclick={() => { A.insertPageBreak(ctl); close(); }}>{t('insertPageBreak')}</button>
      <button class="xl-menu-item" onclick={() => { A.removePageBreak(ctl); close(); }}>{t('removePageBreak')}</button>
      <button class="xl-menu-item" onclick={() => { A.resetPageBreaks(ctl); close(); }}>{t('resetPageBreaks')}</button>
    {/snippet}
  </MenuButton>
  <button class="xl-btn big" onclick={() => ctl.openDialog('pageSetup', { tab: 'sheet' })}><Icon name="headings" size={24} /><span>{t('printTitles')}</span></button>
</Group>

<Group label={t('groupScaleToFit')}>
  <div class="col">
    <label class="field">{t('width')}<select class="xl-select" value={String(setup?.fitToWidth ?? 'auto')} onchange={(e) => A.setFitTo(ctl, 'width', (e.currentTarget as HTMLSelectElement).value)}><option value="auto">{t('automatic')}</option>{#each [1, 2, 3, 4] as n (n)}<option value={String(n)}>{t('nPages', { n })}</option>{/each}</select></label>
    <label class="field">{t('height')}<select class="xl-select" value={String(setup?.fitToHeight ?? 'auto')} onchange={(e) => A.setFitTo(ctl, 'height', (e.currentTarget as HTMLSelectElement).value)}><option value="auto">{t('automatic')}</option>{#each [1, 2, 3, 4] as n (n)}<option value={String(n)}>{t('nPages', { n })}</option>{/each}</select></label>
    <label class="field">{t('scale')}<input class="xl-input num" type="number" min="10" max="400" value={setup?.scale ?? 100} onchange={(e) => A.setPrintScale(ctl, Number((e.currentTarget as HTMLInputElement).value))} />%</label>
  </div>
</Group>

<Group label={t('groupSheetOptions')}>
  <div class="grid2">
    <span></span><span class="hdr">{t('gridlines')}</span><span class="hdr">{t('headings')}</span>
    <span>{t('view')}</span>
    <input type="checkbox" checked={ctl.showGridlines} onchange={(e) => A.setViewFlag(ctl, 'showGridLines', (e.currentTarget as HTMLInputElement).checked)} />
    <input type="checkbox" checked={ctl.showHeaders} onchange={(e) => A.setViewFlag(ctl, 'showRowColHeaders', (e.currentTarget as HTMLInputElement).checked)} />
    <span>{t('print')}</span>
    <input type="checkbox" checked={print?.gridLines === true} onchange={(e) => A.setPrintOption(ctl, 'gridLines', (e.currentTarget as HTMLInputElement).checked)} />
    <input type="checkbox" checked={print?.headings === true} onchange={(e) => A.setPrintOption(ctl, 'headings', (e.currentTarget as HTMLInputElement).checked)} />
  </div>
</Group>

<style>
  .col {
    display: flex;
    flex-direction: column;
    gap: 2px;
  }
  .field {
    display: flex;
    align-items: center;
    gap: 6px;
    justify-content: space-between;
  }
  .field select {
    width: 100px;
  }
  .num {
    width: 60px;
  }
  .big {
    flex-direction: column;
    height: 64px;
    min-width: 52px;
    font-size: 11px;
  }
  .grid2 {
    display: grid;
    grid-template-columns: auto auto auto;
    gap: 2px 10px;
    align-items: center;
  }
  .hdr {
    font-weight: 600;
  }
</style>
