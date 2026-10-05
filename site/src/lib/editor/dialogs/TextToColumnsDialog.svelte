<script lang="ts">
  import { getCellDisplayText } from '@office-kit/xlsx/styles';
  import { cellAddress, MAX_ROW, parseRangeAddress } from '../core/address.ts';
  import { getCellAt, usedRange } from '../core/cells.ts';
  import { getEditor } from '../core/context.ts';
  import { currentRange } from '../core/selection.ts';
  import { parseBreaks, splitText, textToColumns, type SplitOptions } from '../core/text-to-columns.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const PREVIEW_ROWS = 8;

  const ctl = getEditor();
  const doc = ctl.doc;
  const sel = currentRange(doc.selection);
  // Whole columns convert only the rows that hold data.
  const range = sel.r2 === MAX_ROW ? { ...sel, r2: Math.max(sel.r1, usedRange(doc.ws)?.r2 ?? sel.r1) } : sel;
  const singleColumn = doc.selection.ranges.length === 1 && sel.c1 === sel.c2;

  let kind = $state<'delimited' | 'fixed'>('delimited');
  let tab = $state(true);
  let semicolon = $state(false);
  let comma = $state(false);
  let space = $state(false);
  let otherOn = $state(false);
  let other = $state('');
  let consecutiveAsOne = $state(false);
  let qualifier = $state<'"' | "'" | 'none'>('"');
  let breaksText = $state('');
  let destination = $state(cellAddress(range.r1, range.c1, true));
  let notice = $state<string | null>(null);

  const samples = (() => {
    const out: string[] = [];
    for (let r = range.r1; r <= range.r2 && out.length < PREVIEW_ROWS; r++) {
      const cell = getCellAt(doc.ws, r, range.c1);
      out.push(cell && cell.value !== null ? getCellDisplayText(doc.wb, cell) : '');
    }
    return out;
  })();

  const options = $derived.by((): SplitOptions | undefined => {
    if (kind === 'fixed') {
      const breaks = parseBreaks(breaksText);
      return breaks ? { kind: 'fixed', breaks } : undefined;
    }
    return { kind: 'delimited', tab, semicolon, comma, space, other: otherOn ? other : '', consecutiveAsOne, qualifier: qualifier === 'none' ? null : qualifier };
  });
  const preview = $derived(options ? samples.map((s) => splitText(s, options)) : []);
  const previewWidth = $derived(Math.max(1, ...preview.map((p) => p.length)));

  function onok(): boolean {
    if (!singleColumn) {
      notice = t('dlgTtcOneColumn');
      return false;
    }
    if (!options) {
      notice = t('dlgTtcBadBreaks');
      return false;
    }
    const dest = parseRangeAddress(destination.trim().replace(/^=/, ''));
    if (!dest || (dest.sheet !== undefined && dest.sheet.toLowerCase() !== doc.ws.title.toLowerCase())) {
      notice = t('invalidReference');
      return false;
    }
    textToColumns(ctl, range, options, { row: dest.range.r1, col: dest.range.c1 });
    return true;
  }
</script>

<Dialog title={t('dlgTextToColumns')} width={520} {onok} okLabel={t('dlgFinish')} bind:notice>
  <fieldset>
    <legend>{t('dlgTtcDataType')}</legend>
    <label class="check"><input type="radio" name="ttc-kind" value="delimited" bind:group={kind} />{t('dlgTtcDelimited')}</label>
    <label class="check"><input type="radio" name="ttc-kind" value="fixed" bind:group={kind} />{t('dlgTtcFixedWidth')}</label>
  </fieldset>
  {#if kind === 'delimited'}
    <fieldset>
      <legend>{t('dlgTtcDelimiters')}</legend>
      <div class="delims">
        <label class="check"><input type="checkbox" bind:checked={tab} />{t('dlgTtcTab')}</label>
        <label class="check"><input type="checkbox" bind:checked={semicolon} />{t('dlgTtcSemicolon')}</label>
        <label class="check"><input type="checkbox" bind:checked={comma} />{t('dlgTtcComma')}</label>
        <label class="check"><input type="checkbox" bind:checked={space} />{t('dlgTtcSpace')}</label>
        <label class="check"><input type="checkbox" bind:checked={otherOn} />{t('dlgTtcOther')}<input class="xl-input other" maxlength="1" bind:value={other} oninput={() => (otherOn = other !== '')} aria-label={t('dlgTtcOther')} /></label>
      </div>
      <div class="row">
        <label class="check"><input type="checkbox" bind:checked={consecutiveAsOne} />{t('dlgTtcConsecutive')}</label>
        <label for="ttc-q" class="qlabel">{t('dlgTtcQualifier')}</label>
        <select id="ttc-q" class="xl-select" bind:value={qualifier}>
          <option value={'"'}>"</option>
          <option value="'">'</option>
          <option value="none">{t('dlgNone')}</option>
        </select>
      </div>
    </fieldset>
  {:else}
    <div class="row">
      <label for="ttc-breaks">{t('dlgTtcBreaks')}</label>
      <input id="ttc-breaks" class="xl-input grow" bind:value={breaksText} placeholder="5, 12" spellcheck="false" />
    </div>
    <p class="hint">{t('dlgTtcBreaksHint')}</p>
  {/if}
  <div class="row">
    <label for="ttc-dest">{t('dlgDestination')}</label>
    <input id="ttc-dest" class="xl-input grow mono" bind:value={destination} spellcheck="false" />
  </div>
  <div class="lbl">{t('dlgDataPreview')}</div>
  <div class="preview list">
    <table>
      <tbody>
        {#each preview as cells, r (r)}
          <tr>
            {#each Array.from({ length: previewWidth }, (_, i) => i) as c (c)}<td>{cells[c] ?? ''}</td>{/each}
          </tr>
        {/each}
      </tbody>
    </table>
  </div>
</Dialog>

<style>
  .delims {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
  }
  .other {
    width: 28px;
    margin-left: 6px;
    min-height: 20px;
    padding: 1px 4px;
  }
  .qlabel {
    margin-left: auto;
  }
  .preview {
    height: 130px;
  }
  .preview table {
    border-collapse: collapse;
    font-family: var(--xl-mono);
    font-size: 11px;
  }
  .preview td {
    border-right: 1px solid var(--xl-border);
    padding: 1px 6px;
    white-space: pre;
  }
  .mono {
    font-family: var(--xl-mono);
  }
</style>
