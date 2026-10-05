<script lang="ts">
  import { untrack } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import { buildHeader, hfPresets, parseHeader, presetLabel, type HeaderParts } from '../core/header-footer.ts';
  import Dialog from './Dialog.svelte';
  import { applyPageForm, PAPER_SIZES, pageForm, type HfPage } from './page-setup.ts';

  let { props }: { props: Record<string, unknown> | undefined } = $props();
  // DialogHost re-creates a dialog for every opening, so its arguments are read once.
  const args = untrack(() => props) ?? {};

  const TABS = [
    { id: 'page', label: 'dlgPsPage' },
    { id: 'margins', label: 'margins' },
    { id: 'headerFooter', label: 'dlgPsHeaderFooter' },
    { id: 'sheet', label: 'dlgSheet' },
  ] as const satisfies ReadonlyArray<{ id: string; label: MessageKey }>;
  type TabId = (typeof TABS)[number]['id'];
  const MARGINS = [
    ['top', 'dlgMarginTop'],
    ['bottom', 'dlgMarginBottom'],
    ['left', 'dlgMarginLeft'],
    ['right', 'dlgMarginRight'],
    ['header', 'dlgMarginHeader'],
    ['footer', 'dlgMarginFooter'],
  ] as const satisfies ReadonlyArray<[string, MessageKey]>;

  const ctl = getEditor();
  const form = $state(pageForm(ctl));
  let tab = $state<TabId>(TABS.find((x) => x.id === args['tab'])?.id ?? 'page');
  let notice = $state<string | null>(null);
  const paperSizes = PAPER_SIZES.some((p) => p.code === form.paperSize) ? PAPER_SIZES : [...PAPER_SIZES, { code: form.paperSize, label: `#${form.paperSize}` }];

  let hfPage = $state<HfPage>('odd');
  const hfPages = $derived<HfPage[]>(['odd', ...(form.differentFirst ? ['first' as const] : []), ...(form.differentOddEven ? ['even' as const] : [])]);
  const HF_PAGE_LABELS: Record<HfPage, MessageKey> = { odd: 'dtHfOddPage', first: 'dtHfFirstPage', even: 'dtHfEvenPage' };
  // Unticking "different first page" while on that page falls back to the odd pages.
  const shownPage = $derived(hfPages.includes(hfPage) ? hfPage : 'odd');
  const presets = hfPresets({ page: t('dtHfPage'), pageOf: t('dtHfPageOf'), confidential: t('dtHfConfidential') });
  const now = new Date();
  const fields = { date: now.toLocaleDateString(), time: now.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }), file: ctl.doc.fileName, path: '', sheet: ctl.doc.ws.title };

  function presetOf(parts: HeaderParts): string {
    return buildHeader(parts) ?? '';
  }

  function onok(): boolean {
    const err = applyPageForm(ctl, form);
    if (err) {
      notice = t(err);
      return false;
    }
    return true;
  }
</script>

<Dialog title={t('dlgPageSetup')} width={520} {onok} bind:notice>
  <div class="seg" role="tablist">
    {#each TABS as item (item.id)}
      <button role="tab" aria-selected={tab === item.id} onclick={() => (tab = item.id)}>{t(item.label)}</button>
    {/each}
  </div>
  <div class="panel" role="tabpanel">
    {#if tab === 'page'}
      <fieldset>
        <legend>{t('orientation')}</legend>
        <label class="check inline"><input type="radio" name="ps-or" value="portrait" bind:group={form.orientation} />{t('portrait')}</label>
        <label class="check inline"><input type="radio" name="ps-or" value="landscape" bind:group={form.orientation} />{t('landscape')}</label>
      </fieldset>
      <fieldset>
        <legend>{t('dlgPsScaling')}</legend>
        <div class="row">
          <label class="check"><input type="radio" name="ps-fit" value={false} bind:group={form.fit} />{t('dlgPsAdjustTo')}</label>
          <input class="xl-input" type="number" min="10" max="400" bind:value={form.scale} disabled={form.fit} aria-label={t('scale')} />
          <span>{t('dlgPsNormalSize')}</span>
        </div>
        <div class="row">
          <label class="check"><input type="radio" name="ps-fit" value={true} bind:group={form.fit} />{t('dlgPsFitTo')}</label>
          <input class="xl-input small" bind:value={form.fitWide} disabled={!form.fit} placeholder={t('dlgAutomaticShort')} aria-label={t('dlgPsPagesWide')} />
          <span>{t('dlgPsPagesWide')}</span>
          <input class="xl-input small" bind:value={form.fitTall} disabled={!form.fit} placeholder={t('dlgAutomaticShort')} aria-label={t('dlgPsPagesTall')} />
          <span>{t('dlgPsPagesTall')}</span>
        </div>
      </fieldset>
      <div class="row">
        <label for="ps-paper">{t('dlgPsPaperSize')}</label>
        <select id="ps-paper" class="xl-select grow" bind:value={form.paperSize}>
          {#each paperSizes as p (p.code)}<option value={p.code}>{p.label}</option>{/each}
        </select>
      </div>
      <div class="row">
        <label for="ps-first">{t('dlgPsFirstPageNumber')}</label>
        <input id="ps-first" class="xl-input" bind:value={form.firstPage} placeholder={t('dlgAutomaticShort')} />
      </div>
    {:else if tab === 'margins'}
      <div class="margins">
        {#each MARGINS as [key, label] (key)}
          <div class="row">
            <label for="ps-m-{key}">{t(label)}</label>
            <input id="ps-m-{key}" class="xl-input" type="number" min="0" step="0.05" bind:value={form.margins[key]} />
            <span class="hint">{t('dlgInches')}</span>
          </div>
        {/each}
      </div>
      <fieldset>
        <legend>{t('dlgPsCenterOnPage')}</legend>
        <label class="check inline"><input type="checkbox" bind:checked={form.centerH} />{t('dlgHorizontally')}</label>
        <label class="check inline"><input type="checkbox" bind:checked={form.centerV} />{t('dlgVertically')}</label>
      </fieldset>
    {:else if tab === 'headerFooter'}
      {#if hfPages.length > 1}
        <div class="seg small-seg" role="tablist">
          {#each hfPages as id (id)}
            <button role="tab" aria-selected={shownPage === id} onclick={() => (hfPage = id)}>{t(HF_PAGE_LABELS[id])}</button>
          {/each}
        </div>
      {/if}
      {#each [['header', 'dlgPsHeader', 'dtHfHeaderPreset'], ['footer', 'dlgPsFooter', 'dtHfFooterPreset']] as const as [part, label, presetKey] (part)}
        {@const parts = form.headers[shownPage][part]}
        {@const current = presetOf(parts)}
        <fieldset>
          <legend>{t(label)}</legend>
          <div class="row">
            <label for="ps-{part}-preset">{t(presetKey)}</label>
            <select
              id="ps-{part}-preset"
              class="xl-select grow"
              value={current}
              onchange={(e) => (form.headers[shownPage][part] = parseHeader(e.currentTarget.value))}
            >
              {#each presets as codes (codes)}
                <option value={codes}>{codes === '' ? t('dtHfNone') : presetLabel(codes, fields)}</option>
              {/each}
              {#if !presets.includes(current)}
                <option value={current}>{t('dtHfCustom')}</option>
              {/if}
            </select>
          </div>
          <div class="hf">
            <input class="xl-input" bind:value={parts.left} aria-label="{t(label)} — {t('dlgLeftSection')}" placeholder={t('dlgLeftSection')} />
            <input class="xl-input" bind:value={parts.center} aria-label="{t(label)} — {t('dlgCenterSection')}" placeholder={t('dlgCenterSection')} />
            <input class="xl-input" bind:value={parts.right} aria-label="{t(label)} — {t('dlgRightSection')}" placeholder={t('dlgRightSection')} />
          </div>
        </fieldset>
      {/each}
      <p class="hint">{t('dlgPsCodesHint')}</p>
      <div class="two">
        <label class="check"><input type="checkbox" bind:checked={form.differentOddEven} />{t('dtHfDifferentOddEven')}</label>
        <label class="check"><input type="checkbox" bind:checked={form.differentFirst} />{t('dtHfDifferentFirst')}</label>
        <label class="check"><input type="checkbox" bind:checked={form.scaleWithDoc} />{t('dtHfScaleWithDoc')}</label>
        <label class="check"><input type="checkbox" bind:checked={form.alignWithMargins} />{t('dtHfAlignMargins')}</label>
      </div>
    {:else}
      <div class="row">
        <label for="ps-area">{t('printArea')}</label>
        <input id="ps-area" class="xl-input grow mono" bind:value={form.printArea} spellcheck="false" />
      </div>
      <fieldset>
        <legend>{t('printTitles')}</legend>
        <div class="row">
          <label for="ps-rows">{t('dlgPsRowsToRepeat')}</label>
          <input id="ps-rows" class="xl-input grow mono" bind:value={form.titleRows} placeholder="$1:$1" spellcheck="false" />
        </div>
        <div class="row">
          <label for="ps-cols">{t('dlgPsColumnsToRepeat')}</label>
          <input id="ps-cols" class="xl-input grow mono" bind:value={form.titleCols} placeholder="$A:$A" spellcheck="false" />
        </div>
      </fieldset>
      <fieldset class="two">
        <legend>{t('print')}</legend>
        <label class="check"><input type="checkbox" bind:checked={form.gridLines} />{t('gridlines')}</label>
        <label class="check"><input type="checkbox" bind:checked={form.headings} />{t('dlgPsRowColHeadings')}</label>
        <label class="check"><input type="checkbox" bind:checked={form.blackAndWhite} />{t('dlgPsBlackWhite')}</label>
        <label class="check"><input type="checkbox" bind:checked={form.draft} />{t('dlgPsDraft')}</label>
      </fieldset>
      <div class="row">
        <label for="ps-comments">{t('dlgPgComments')}</label>
        <select id="ps-comments" class="xl-select" bind:value={form.comments}>
          <option value="none">{t('dlgNone')}</option>
          <option value="atEnd">{t('dlgPsAtEnd')}</option>
          <option value="asDisplayed">{t('dlgPsAsDisplayed')}</option>
        </select>
      </div>
      <fieldset>
        <legend>{t('dlgPsPageOrder')}</legend>
        <label class="check"><input type="radio" name="ps-order" value="downThenOver" bind:group={form.pageOrder} />{t('dlgPsDownThenOver')}</label>
        <label class="check"><input type="radio" name="ps-order" value="overThenDown" bind:group={form.pageOrder} />{t('dlgPsOverThenDown')}</label>
      </fieldset>
    {/if}
  </div>
</Dialog>

<style>
  .panel {
    min-height: 280px;
  }
  .inline {
    display: inline-flex;
    margin-right: 18px;
  }
  .small {
    width: 52px;
  }
  .margins {
    display: grid;
    grid-template-columns: 1fr 1fr;
    column-gap: 16px;
  }
  .hf {
    display: grid;
    grid-template-columns: 1fr 1fr 1fr;
    gap: 6px;
  }
  .small-seg {
    margin-bottom: 6px;
  }
  .two {
    display: grid;
    grid-template-columns: 1fr 1fr;
  }
  .mono {
    font-family: var(--xl-mono);
  }
</style>
