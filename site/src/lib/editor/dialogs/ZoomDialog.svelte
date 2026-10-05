<script lang="ts">
  import { zoomToSelection } from '../core/actions.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const PRESETS = [200, 100, 75, 50, 25];
  const MIN_ZOOM = 10;
  const MAX_ZOOM = 400;

  const ctl = getEditor();
  const current = Math.round(ctl.doc.zoom * 100);
  let choice = $state<number | 'fit' | 'custom'>(PRESETS.includes(current) ? current : 'custom');
  let custom = $state(String(current));
  let notice = $state<string | null>(null);

  function onok(): boolean {
    if (choice === 'fit') {
      zoomToSelection(ctl);
      return true;
    }
    const pct = choice === 'custom' ? Number(custom.trim()) : choice;
    if (!Number.isInteger(pct) || pct < MIN_ZOOM || pct > MAX_ZOOM) {
      notice = t('dlgZoomRange');
      return false;
    }
    ctl.doc.setZoom(pct / 100);
    return true;
  }
</script>

<Dialog title={t('zoom')} width={240} {onok} bind:notice>
  <fieldset>
    <legend>{t('dlgMagnification')}</legend>
    {#each PRESETS as p (p)}
      <label class="check"><input type="radio" name="zoom" value={p} bind:group={choice} />{p}%</label>
    {/each}
    <label class="check"><input type="radio" name="zoom" value="fit" bind:group={choice} />{t('dlgFitSelection')}</label>
    <div class="row">
      <label class="check"><input type="radio" name="zoom" value="custom" bind:group={choice} />{t('dlgCustom')}</label>
      <input class="xl-input pct" bind:value={custom} inputmode="numeric" aria-label={t('dlgCustom')} onfocus={() => (choice = 'custom')} />%
    </div>
  </fieldset>
</Dialog>

<style>
  .pct {
    width: 56px;
  }
</style>
