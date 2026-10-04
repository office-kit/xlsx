<script lang="ts">
  import * as A from '../core/actions.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import MenuButton from '../ui/MenuButton.svelte';
  import Group from './Group.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;
  const view = $derived(ctl.sheetView);
  // Excel's Focus Cell colour choices.
  const FOCUS_COLORS = ['#F5C342', '#8BC34A', '#4FC3F7', '#F48FB1', '#B39DDB'] as const;
  const pageLayout = $derived(view?.view === 'pageLayout' || view?.view === 'pageBreakPreview');
</script>

<Group label={t('groupWorkbookViews')}>
  <button class="xl-btn big" aria-pressed={!pageLayout} onclick={() => A.setSheetViewMode(ctl, 'normal')}><Icon name="sheet" size={24} /><span>{t('viewNormal')}</span></button>
  <button class="xl-btn big" aria-pressed={view?.view === 'pageLayout'} onclick={() => A.setSheetViewMode(ctl, 'pageLayout')}><Icon name="page-layout" size={24} /><span>{t('viewPageLayout')}</span></button>
  <button class="xl-btn big" aria-pressed={view?.view === 'pageBreakPreview'} onclick={() => A.setSheetViewMode(ctl, 'pageBreakPreview')}><Icon name="page" size={24} /><span>{t('viewPageBreak')}</span></button>
</Group>

<Group label={t('groupShow')}>
  <button class="xl-btn big" aria-pressed={ctl.navigationPane} onclick={() => (ctl.navigationPane = !ctl.navigationPane)}><Icon name="navigation" size={24} /><span>{t('navigation')}</span></button>
  <div class="col">
    <label class="check"><input type="checkbox" checked={ctl.showGridlines} onchange={(e) => A.setViewFlag(ctl, 'showGridLines', (e.currentTarget as HTMLInputElement).checked)} />{t('gridlines')}</label>
    <label class="check"><input type="checkbox" checked={ctl.showFormulaBar} onchange={(e) => (ctl.showFormulaBar = (e.currentTarget as HTMLInputElement).checked)} />{t('formulaBar')}</label>
    <label class="check"><input type="checkbox" checked={ctl.showHeaders} onchange={(e) => A.setViewFlag(ctl, 'showRowColHeaders', (e.currentTarget as HTMLInputElement).checked)} />{t('headings')}</label>
  </div>
</Group>

<Group label={t('groupFocusCell')}>
  <MenuButton large split icon="crosshair" label={t('focusCell')} onmain={() => (ctl.focusCell = ctl.focusCell ? null : FOCUS_COLORS[0])}>
    {#snippet menu(close)}
      {#each FOCUS_COLORS as color (color)}
        <button class="xl-menu-item" aria-pressed={ctl.focusCell === color} onclick={() => { ctl.focusCell = color; close(); }}>
          <span class="swatch" style:background={color}></span>{color === ctl.focusCell ? '✓ ' : ''}{color}
        </button>
      {/each}
    {/snippet}
  </MenuButton>
</Group>

<Group label={t('groupZoom')}>
  <button class="xl-btn big" onclick={() => ctl.openDialog('zoom')}><Icon name="zoom-in" size={24} /><span>{t('zoom')}</span></button>
  <button class="xl-btn big" onclick={() => doc.setZoom(1)}><span class="hundred">100</span><span>100%</span></button>
  <button class="xl-btn big" onclick={() => A.zoomToSelection(ctl)}><Icon name="grid" size={24} /><span>{t('zoomToSelection')}</span></button>
</Group>

<Group label={t('groupWindow')}>
  <MenuButton large icon="freeze" label={t('freezePanes')}>
    {#snippet menu(close)}
      <button class="xl-menu-item" onclick={() => { A.freezePanes(ctl, 'panes'); close(); }}>{ctl.frozen.rows + ctl.frozen.cols > 0 ? t('unfreezePanes') : t('freezePanes')}</button>
      <button class="xl-menu-item" onclick={() => { A.freezePanes(ctl, 'topRow'); close(); }}>{t('freezeTopRow')}</button>
      <button class="xl-menu-item" onclick={() => { A.freezePanes(ctl, 'firstColumn'); close(); }}>{t('freezeFirstColumn')}</button>
    {/snippet}
  </MenuButton>
  <div class="col">
    <button class="xl-btn" onclick={() => ctl.openDialog('unhideSheet')}>{t('unhideSheetEllipsis')}</button>
    <label class="check"><input type="checkbox" checked={ctl.showFormulas} onchange={(e) => A.setViewFlag(ctl, 'showFormulas', (e.currentTarget as HTMLInputElement).checked)} />{t('showFormulas')}</label>
    <label class="check"><input type="checkbox" checked={view?.showZeros !== false} onchange={(e) => A.setViewFlag(ctl, 'showZeros', (e.currentTarget as HTMLInputElement).checked)} />{t('showZeros')}</label>
  </div>
</Group>

<style>
  .swatch {
    display: inline-block;
    width: 12px;
    height: 12px;
    margin-right: 6px;
    border: 1px solid #999;
    vertical-align: -2px;
  }
  .col {
    display: flex;
    flex-direction: column;
    gap: 2px;
  }
  .check {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 1px 4px;
    white-space: nowrap;
  }
  .big {
    flex-direction: column;
    height: 64px;
    min-width: 52px;
    font-size: 11px;
  }
  .hundred {
    font-weight: 700;
    font-size: 13px;
  }
</style>
