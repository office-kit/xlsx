<script lang="ts">
  import { makeCell } from '@office-kit/xlsx/cell';
  import { getCellDisplayText, makeColor } from '@office-kit/xlsx/styles';
  import { makeHyperlink, setHyperlink } from '@office-kit/xlsx/worksheet';
  import { OpenXmlError } from '@office-kit/xlsx/utils';
  import { cellAddress, parseRangeAddress, quoteSheetName, rangeAddress } from '../core/address.ts';
  import { getCellAt } from '../core/cells.ts';
  import { clearRanges } from '../core/commands.ts';
  import { getEditor } from '../core/context.ts';
  import { applyStyle } from '../core/format.ts';
  import { currentRange } from '../core/selection.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  type Mode = 'web' | 'place' | 'email';

  /** Excel's Hyperlink cell style colour. */
  const LINK_COLOR = '0563C1';

  const ctl = getEditor();
  const doc = ctl.doc;
  const ws = doc.ws;
  const range = currentRange(doc.selection);
  const anchor = { row: range.r1, col: range.c1 };
  const anchorAddress = cellAddress(anchor.row, anchor.col);
  const existing = ws.hyperlinks.find((h) => h.ref.replaceAll('$', '').split(':')[0] === anchorAddress);
  const anchorCell = getCellAt(ws, anchor.row, anchor.col);
  const initialText = anchorCell && anchorCell.value !== null ? getCellDisplayText(doc.wb, anchorCell) : '';
  const worksheets = doc.wb.sheets.filter((s) => s.kind === 'worksheet').map((s) => s.sheet.title);
  const names = doc.wb.definedNames.filter((d) => !d.hidden && !d.name.startsWith('_xlnm.')).map((d) => d.name);

  function initialMode(): Mode {
    if (existing?.target?.startsWith('mailto:')) return 'email';
    if (existing && existing.target === undefined && existing.location !== undefined) return 'place';
    return 'web';
  }

  const placeParts = (() => {
    const loc = existing?.location ?? '';
    const parsed = parseRangeAddress(loc);
    if (parsed?.sheet !== undefined) return { sheet: parsed.sheet, cell: rangeAddress(parsed.range) };
    return { sheet: names.includes(loc) ? '' : (worksheets[doc.activeSheetIndex] ?? worksheets[0] ?? ''), cell: names.includes(loc) ? loc : 'A1' };
  })();
  const mail = existing?.target?.startsWith('mailto:') ? existing.target.slice(7).split('?') : undefined;

  let mode = $state<Mode>(initialMode());
  let display = $state(existing?.display ?? initialText);
  let tooltip = $state(existing?.tooltip ?? '');
  let address = $state(existing?.target && !existing.target.startsWith('mailto:') ? existing.target : '');
  let sheet = $state(placeParts.sheet);
  let cellRef = $state(placeParts.cell);
  let email = $state(mail?.[0] ?? '');
  let subject = $state(mail?.[1]?.startsWith('subject=') ? decodeURIComponent(mail[1].slice(8)) : '');
  let notice = $state<string | null>(null);

  function link(): { target: string } | { location: string } | undefined {
    switch (mode) {
      case 'web': {
        const a = address.trim();
        if (!a) return undefined;
        // Excel adds the scheme to a bare host name.
        return { target: /^[a-z][a-z0-9+.-]*:/i.test(a) || a.startsWith('/') || a.startsWith('.') ? a : `https://${a}` };
      }
      case 'email': {
        const e = email.trim();
        if (!e) return undefined;
        return { target: `mailto:${e}${subject.trim() ? `?subject=${encodeURIComponent(subject.trim())}` : ''}` };
      }
      case 'place': {
        const ref = cellRef.trim();
        if (names.includes(ref)) return { location: ref };
        if (!parseRangeAddress(ref)) return undefined;
        return { location: `${quoteSheetName(sheet)}!${ref.toUpperCase()}` };
      }
    }
  }

  function onok(): boolean {
    const target = link();
    if (!target) {
      notice = t(mode === 'place' ? 'invalidReference' : 'dlgLinkAddressRequired');
      return false;
    }
    const ref = rangeAddress(range);
    const text = display.trim() === '' ? ('target' in target ? target.target.replace(/^mailto:/, '') : target.location) : display;
    const opts = { ...target, ...(tooltip.trim() ? { tooltip: tooltip.trim() } : {}) };
    try {
      makeHyperlink({ ref, ...opts });
    } catch (e) {
      // The library rejects targets Excel would refuse (spaces, control characters, non-ASCII hosts).
      if (!(e instanceof OpenXmlError)) throw e;
      notice = t('dlgLinkInvalid');
      return false;
    }
    doc.transact(existing ? 'Edit Hyperlink' : 'Insert Hyperlink', (tx) => {
      tx.sheet(ws, 'hyperlinks');
      tx.cells(ws, range);
      setHyperlink(ws, ref, opts);
      if (text !== initialText || anchorCell?.value === null || anchorCell === undefined) {
        const cell = getCellAt(ws, anchor.row, anchor.col);
        if (cell) cell.value = text;
        else {
          let rowMap = ws.rows.get(anchor.row);
          if (!rowMap) {
            rowMap = new Map();
            ws.rows.set(anchor.row, rowMap);
          }
          rowMap.set(anchor.col, makeCell(anchor.row, anchor.col, text, ctl.defaultStyleAt(anchor.row, anchor.col)));
        }
      }
      if (!existing) applyStyle(doc.wb, ws, range, { font: { color: makeColor({ rgb: LINK_COLOR }), underline: 'single' } });
    });
    return true;
  }

  function removeLink(): void {
    clearRanges(doc, doc.selection.ranges, 'hyperlinks');
    ctl.closeDialog();
  }
</script>

<Dialog title={t(existing ? 'dlgEditHyperlink' : 'dlgInsertHyperlink')} width={480} {onok} bind:notice>
  {#snippet footerStart()}
    {#if existing}<button class="xl-btn outlined" onclick={removeLink}>{t('removeHyperlink')}</button>{/if}
  {/snippet}
  <div class="seg" role="tablist">
    <button role="tab" aria-selected={mode === 'web'} onclick={() => (mode = 'web')}>{t('dlgLinkWeb')}</button>
    <button role="tab" aria-selected={mode === 'place'} onclick={() => (mode = 'place')}>{t('dlgLinkPlace')}</button>
    <button role="tab" aria-selected={mode === 'email'} onclick={() => (mode = 'email')}>{t('dlgLinkEmail')}</button>
  </div>
  <div class="row">
    <label for="hl-text">{t('dlgTextToDisplay')}</label>
    <input id="hl-text" class="xl-input grow" bind:value={display} />
  </div>
  {#if mode === 'web'}
    <div class="row">
      <label for="hl-addr">{t('dlgAddress')}</label>
      <input id="hl-addr" class="xl-input grow" bind:value={address} placeholder="https://" spellcheck="false" data-autofocus />
    </div>
  {:else if mode === 'email'}
    <div class="row">
      <label for="hl-mail">{t('dlgEmailAddress')}</label>
      <input id="hl-mail" class="xl-input grow" type="email" bind:value={email} spellcheck="false" data-autofocus />
    </div>
    <div class="row">
      <label for="hl-subj">{t('dlgSubject')}</label>
      <input id="hl-subj" class="xl-input grow" bind:value={subject} />
    </div>
  {:else}
    <div class="row">
      <label for="hl-ref">{t('dlgCellReference')}</label>
      <input id="hl-ref" class="xl-input grow" bind:value={cellRef} spellcheck="false" data-autofocus />
    </div>
    <div class="lbl">{t('dlgPlaceInDocument')}</div>
    <div class="list places" role="listbox" aria-label={t('dlgPlaceInDocument')}>
      {#each worksheets as title (title)}
        <button role="option" aria-selected={sheet === title && !names.includes(cellRef)} onclick={() => { sheet = title; if (names.includes(cellRef)) cellRef = 'A1'; }}>{title}</button>
      {/each}
      {#each names as n (n)}
        <button role="option" class="name" aria-selected={cellRef === n} onclick={() => (cellRef = n)}>{n}</button>
      {/each}
    </div>
  {/if}
  <div class="row">
    <label for="hl-tip">{t('dlgScreenTip')}</label>
    <input id="hl-tip" class="xl-input grow" bind:value={tooltip} />
  </div>
</Dialog>

<style>
  .places {
    height: 110px;
    margin-bottom: 6px;
  }
  .name {
    font-style: italic;
  }
</style>
