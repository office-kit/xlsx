<script lang="ts">
  import { untrack } from 'svelte';
  import { renameSheetAt } from '../core/actions.ts';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  let { props }: { props: Record<string, unknown> | undefined } = $props();
  // DialogHost re-creates a dialog for every opening, so its arguments are read once.
  const args = untrack(() => props) ?? {};

  const ctl = getEditor();
  const index = typeof args['index'] === 'number' ? args['index'] : ctl.doc.activeSheetIndex;
  let name = $state(ctl.doc.wb.sheets[index]?.sheet.title ?? '');
  let notice = $state<string | null>(null);

  function onok(): boolean {
    const err = renameSheetAt(ctl, index, name);
    if (err) {
      notice = t(err);
      return false;
    }
    return true;
  }
</script>

<Dialog title={t('renameSheet')} width={320} {onok} bind:notice>
  <div class="row">
    <label for="rs-name">{t('dlgSheetName')}</label>
    <input id="rs-name" class="xl-input grow" bind:value={name} maxlength="31" spellcheck="false" />
  </div>
</Dialog>
