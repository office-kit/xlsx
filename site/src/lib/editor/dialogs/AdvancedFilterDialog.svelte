<script lang="ts">
  import { parseRangeAddress, rangeAddress, type Range } from '../core/address.ts';
  import { runAdvancedFilter } from '../core/advanced-filter.ts';
  import { getEditor } from '../core/context.ts';
  import { dataRange } from '../core/data.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';
  import RefInput from './RefInput.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;

  let copy = $state(false);
  let listText = $state(rangeAddress(dataRange(ctl), true));
  let criteriaText = $state('');
  let copyText = $state('');
  let unique = $state(false);
  let notice = $state<string | null>(null);

  /** A range on the active sheet: Advanced Filter works within one sheet here. */
  function rangeOn(text: string): Range | undefined {
    const parsed = parseRangeAddress(text.trim().replace(/^=/, ''));
    if (!parsed || (parsed.sheet !== undefined && parsed.sheet.toLowerCase() !== doc.ws.title.toLowerCase())) return undefined;
    return parsed.range;
  }

  function onok(): boolean {
    const list = rangeOn(listText);
    const criteria = criteriaText.trim() === '' ? undefined : rangeOn(criteriaText);
    const copyTo = copy ? rangeOn(copyText) : undefined;
    if (!list || (criteriaText.trim() !== '' && !criteria) || (copy && !copyTo)) {
      notice = t('invalidReference');
      return false;
    }
    const err = runAdvancedFilter(ctl, { list, criteria, copyTo, unique });
    if (err) {
      notice = t(err);
      return false;
    }
    return true;
  }
</script>

<Dialog title={t('dtAdvancedFilter')} width={360} modal={false} {onok} bind:notice>
  <fieldset>
    <legend>{t('dtAfAction')}</legend>
    <label class="check"><input type="radio" name="af-action" value={false} bind:group={copy} />{t('dtAfInPlace')}</label>
    <label class="check"><input type="radio" name="af-action" value={true} bind:group={copy} />{t('dtAfCopy')}</label>
  </fieldset>
  <div class="grid">
    <label for="af-list">{t('dtAfListRange')}</label>
    <RefInput id="af-list" bind:value={listText} autofocus />
    <label for="af-criteria">{t('dtAfCriteriaRange')}</label>
    <RefInput id="af-criteria" bind:value={criteriaText} />
    <label for="af-copy" class:dim={!copy}>{t('dtAfCopyTo')}</label>
    <RefInput id="af-copy" bind:value={copyText} />
  </div>
  <label class="check"><input type="checkbox" bind:checked={unique} />{t('dtAfUnique')}</label>
</Dialog>

<style>
  .grid {
    display: grid;
    grid-template-columns: auto 1fr;
    gap: 6px 8px;
    align-items: center;
    margin: 8px 0;
  }
  .dim {
    color: var(--xl-text-2);
  }
</style>
