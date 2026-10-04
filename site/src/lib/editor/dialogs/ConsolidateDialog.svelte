<script lang="ts">
  import type { DataConsolidateFunction } from '@office-kit/xlsx/worksheet';
  import { parseRangeAddress } from '../core/address.ts';
  import { CONSOLIDATE_FUNCTIONS, parseSource, runConsolidate, sourceText } from '../core/consolidate.ts';
  import { getEditor } from '../core/context.ts';
  import { selectRange } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';
  import RefInput from './RefInput.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;
  // The sheet remembers the last consolidation, so the dialog reopens with it.
  const saved = doc.ws.dataConsolidate;

  let fn = $state<DataConsolidateFunction>(saved?.function ?? 'sum');
  let reference = $state('');
  const sources = $state<string[]>(
    (saved?.dataRefs ?? []).flatMap((r) => {
      const parsed = r.ref ? parseRangeAddress(r.ref) : undefined;
      if (!parsed) return [];
      return [sourceText({ sheet: r.sheet ?? parsed.sheet ?? doc.ws.title, range: parsed.range })];
    }),
  );
  let selected = $state(-1);
  let topLabels = $state(saved?.topLabels ?? false);
  let leftLabels = $state(saved?.leftLabels ?? false);
  let notice = $state<string | null>(null);

  function add(): void {
    const src = parseSource(ctl, reference);
    if (!src) {
      notice = t('invalidReference');
      return;
    }
    const text = sourceText(src);
    if (!sources.includes(text)) sources.push(text);
    selected = sources.indexOf(text);
  }

  function remove(): void {
    if (selected < 0) return;
    sources.splice(selected, 1);
    selected = Math.min(selected, sources.length - 1);
  }

  function onok(): boolean {
    // A reference typed but not added still counts, as in Excel.
    const typed = reference.trim() === '' ? undefined : parseSource(ctl, reference);
    if (reference.trim() !== '' && !typed) {
      notice = t('invalidReference');
      return false;
    }
    const list = typed && !sources.includes(sourceText(typed)) ? [...sources, sourceText(typed)] : sources;
    const parsed = list.flatMap((s) => parseSource(ctl, s) ?? []);
    if (parsed.length === 0) {
      notice = t('dtConsNoSources');
      return false;
    }
    const target = runConsolidate(ctl, parsed, { function: fn, topLabels, leftLabels });
    doc.setSelection(selectRange(target, doc.selection.active));
    return true;
  }
</script>

<Dialog title={t('dtConsolidate')} width={420} modal={false} {onok} bind:notice>
  <label class="lbl" for="cons-fn">{t('dtConsFunction')}</label>
  <select id="cons-fn" class="xl-select wide" bind:value={fn}>
    {#each CONSOLIDATE_FUNCTIONS as f (f.fn)}<option value={f.fn}>{t(f.label)}</option>{/each}
  </select>
  <label class="lbl" for="cons-ref">{t('dtConsReference')}</label>
  <div class="row">
    <RefInput id="cons-ref" bind:value={reference} withSheet autofocus />
    <button class="xl-btn outlined" onclick={add}>{t('dtConsAdd')}</button>
  </div>
  <span class="lbl">{t('dtConsAllReferences')}</span>
  <div class="row top">
    <div class="list refs" role="listbox" aria-label={t('dtConsAllReferences')}>
      {#each sources as s, i (s)}
        <button role="option" aria-selected={i === selected} onclick={() => {
          selected = i;
          reference = s;
        }}>{s}</button>
      {/each}
    </div>
    <button class="xl-btn outlined" disabled={selected < 0} onclick={remove}>{t('dtConsDelete')}</button>
  </div>
  <fieldset>
    <legend>{t('dtConsUseLabels')}</legend>
    <label class="check"><input type="checkbox" bind:checked={topLabels} />{t('dtConsTopRow')}</label>
    <label class="check"><input type="checkbox" bind:checked={leftLabels} />{t('dtConsLeftColumn')}</label>
  </fieldset>
</Dialog>

<style>
  .lbl {
    display: block;
    margin: 6px 0 3px;
  }
  .wide {
    width: 100%;
  }
  .top {
    align-items: flex-start;
  }
  .refs {
    flex: 1;
    height: 96px;
    overflow-y: auto;
    font-family: var(--xl-mono);
  }
</style>
