<script lang="ts">
  // Name Box + Cancel/Enter/Insert Function buttons + the formula input. The
  // input mirrors the in-cell editor: typing here starts (or continues) the
  // same edit session, so both stay in sync like Excel's.
  import { getEditor } from '../core/context.ts';
  import { cellAddress, rangeAddress, rangeOf } from '../core/address.ts';
  import { currentRange } from '../core/selection.ts';
  import { editTextFor } from '../core/input.ts';
  import { isDateFormat, getCellDisplayText } from '@office-kit/xlsx/styles';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from './Icon.svelte';
  import { goToReference, validateName } from '../core/names.ts';

  const ctl = getEditor();
  const doc = ctl.doc;
  let input: HTMLTextAreaElement;
  let nameBox: HTMLInputElement;
  let nameText = $state('');
  let nameFocused = $state(false);
  let namesOpen = $state(false);
  let namesMenu = $state<HTMLDivElement>();

  // Like any menu, the defined-names list closes on Escape or a click elsewhere.
  $effect(() => {
    if (!namesOpen) return;
    const down = (e: PointerEvent) => {
      if (!(e.target instanceof Node) || namesMenu?.contains(e.target) || (e.target instanceof Element && e.target.closest('.nb-arrow'))) return;
      namesOpen = false;
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') namesOpen = false;
    };
    window.addEventListener('pointerdown', down, true);
    window.addEventListener('keydown', key, true);
    return () => {
      window.removeEventListener('pointerdown', down, true);
      window.removeEventListener('keydown', key, true);
    };
  });

  const address = $derived.by(() => {
    const sel = doc.selection;
    const e = ctl.edit;
    if (e?.point) {
      const r = doc.selection.ranges[0];
      return r ? rangeAddress(r) : '';
    }
    const r = currentRange(sel);
    // While dragging out a range Excel shows its size ("3R x 2C"); afterwards the active cell.
    if (sel.ranges.length === 1) {
      const name = definedNameFor(r);
      if (name) return name;
    }
    return cellAddress(sel.active.row, sel.active.col);
  });

  function definedNameFor(r: { r1: number; c1: number; r2: number; c2: number }): string | undefined {
    const want = `${doc.ws.title}!${rangeAddress(r, true)}`.toLowerCase();
    for (const dn of doc.wb.definedNames) {
      if (dn.name.startsWith('_xlnm.')) continue;
      const v = dn.value.replaceAll("'", '').toLowerCase();
      if (v === want) return dn.name;
    }
    return undefined;
  }

  $effect(() => {
    if (!nameFocused) nameText = address;
  });

  const barText = $derived.by(() => {
    void doc.version;
    const e = ctl.edit;
    if (e) return e.text;
    const { row, col } = doc.selection.active;
    const cell = ctl.cell(row, col);
    if (!cell) return '';
    const style = doc.styles.get(cell.styleId);
    return editTextFor(cell.value, getCellDisplayText(doc.wb, cell), isDateFormat(style.numFmt), { dateOrder: ctl.dateOrder(), date1904: doc.wb.date1904 });
  });

  $effect(() => {
    const text = barText;
    if (input && document.activeElement !== input) input.value = text;
  });

  $effect(() => {
    // Keep the caret in sync when the edit is driven from the bar.
    const e = ctl.edit;
    if (e && e.source === 'bar' && input && document.activeElement === input) {
      if (input.value !== e.text) input.value = e.text;
      if (input.selectionStart !== e.selStart || input.selectionEnd !== e.selEnd) input.setSelectionRange(e.selStart, e.selEnd);
    }
  });

  function onFocus() {
    if (!ctl.edit) {
      // Loading the cell's content starts in Edit mode (arrows move the caret), like F2.
      ctl.startEdit(undefined, 'bar');
    } else ctl.edit.source = 'bar';
  }

  function onInput() {
    if (!ctl.edit) ctl.startEdit(input.value, 'bar');
    ctl.setEditText(input.value, input.selectionStart, input.selectionEnd);
    if (ctl.edit) ctl.edit.source = 'bar';
  }

  function onKey(e: KeyboardEvent) {
    if (e.isComposing) return;
    if (ctl.handleCompletionKey(e)) {
      const ed = ctl.edit;
      if (ed && input.value !== ed.text) {
        input.value = ed.text;
        input.setSelectionRange(ed.selStart, ed.selEnd);
      }
      return;
    }
    if (e.key === 'Enter' && !e.altKey) {
      e.preventDefault();
      if (ctl.commitEdit({ fillSelection: e.ctrlKey || e.metaKey })) {
        input.blur();
        if (!(e.ctrlKey || e.metaKey)) ctl.advance(e.shiftKey ? 'up' : 'down');
      }
    } else if (e.key === 'Enter' && e.altKey) {
      e.preventDefault();
      const s = input.selectionStart;
      ctl.setEditText(input.value.slice(0, s) + '\n' + input.value.slice(input.selectionEnd), s + 1);
      input.value = ctl.edit?.text ?? input.value;
      input.setSelectionRange(s + 1, s + 1);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      ctl.cancelEdit();
      input.blur();
    } else if (e.key === 'Tab') {
      e.preventDefault();
      if (ctl.commitEdit()) {
        input.blur();
        ctl.advance(e.shiftKey ? 'left' : 'right');
      }
    } else if (e.key === 'F4') {
      e.preventDefault();
      ctl.toggleAbsolute();
    }
  }

  function onSelect() {
    const e = ctl.edit;
    if (e && e.source === 'bar') {
      e.selStart = input.selectionStart;
      e.selEnd = input.selectionEnd;
      e.point = null;
    }
  }

  function goName() {
    const text = nameText.trim();
    nameFocused = false;
    nameBox.blur();
    ctl.gridFocusRequest++;
    if (!text) return;
    if (!goToReference(ctl, text)) {
      if (validateName(text) === undefined) {
        // Typing a new name into the Name Box defines it for the selection.
        ctl.defineNameForSelection(text);
      } else ctl.toast = 'invalidReference';
    }
  }

  const names = $derived.by(() => {
    void doc.version;
    return doc.wb.definedNames.filter((d) => !d.name.startsWith('_xlnm.') && !d.hidden).map((d) => d.name);
  });
</script>

<div class="fbar">
  <div class="namebox">
    <input
      bind:this={nameBox}
      class="name-input"
      aria-label={t('nameBox')}
      bind:value={nameText}
      onfocus={() => { nameFocused = true; nameBox.select(); }}
      onblur={() => (nameFocused = false)}
      onkeydown={(e) => { if (e.key === 'Enter') { e.preventDefault(); goName(); } else if (e.key === 'Escape') { nameFocused = false; nameText = address; nameBox.blur(); ctl.gridFocusRequest++; } }}
    />
    <button class="nb-arrow" aria-label={t('definedNames')} onclick={() => (namesOpen = !namesOpen)} onmousedown={(e) => e.preventDefault()}><Icon name="chevron-down" size={11} /></button>
    {#if namesOpen}
      <div class="xl-menu names" role="menu" bind:this={namesMenu}>
        {#each names as n (n)}
          <button class="xl-menu-item" onclick={() => { goToReference(ctl, n); namesOpen = false; ctl.gridFocusRequest++; }}>{n}</button>
        {:else}
          <div class="empty">{t('noNames')}</div>
        {/each}
      </div>
    {/if}
  </div>
  <div class="buttons">
    <button class="xl-btn" title={t('cancel')} disabled={!ctl.edit} onclick={() => ctl.cancelEdit()} onmousedown={(e) => e.preventDefault()}><Icon name="close" size={14} /></button>
    <button class="xl-btn" title={t('enter')} disabled={!ctl.edit} onclick={() => ctl.commitEdit()} onmousedown={(e) => e.preventDefault()}><Icon name="check" size={14} /></button>
    <button class="xl-btn fx" title={t('insertFunction')} onclick={() => ctl.openDialog('insertFunction')} onmousedown={(e) => e.preventDefault()}><i>fx</i></button>
  </div>
  <textarea
    bind:this={input}
    class="formula"
    class:expanded={ctl.formulaBarExpanded}
    spellcheck="false"
    aria-label={t('formulaBar')}
    rows="1"
    onfocus={onFocus}
    oninput={onInput}
    onkeydown={onKey}
    onselect={onSelect}
    onkeyup={onSelect}
    onclick={onSelect}
  ></textarea>
  <button class="xl-btn expand" title={t('expandFormulaBar')} onclick={() => (ctl.formulaBarExpanded = !ctl.formulaBarExpanded)} onmousedown={(e) => e.preventDefault()}>
    <Icon name={ctl.formulaBarExpanded ? 'chevron-up' : 'chevron-down'} size={12} />
  </button>
</div>

<style>
  .fbar {
    display: flex;
    align-items: stretch;
    gap: 4px;
    padding: 4px 6px;
    background: #fff;
    border-bottom: 1px solid var(--xl-border);
    flex-shrink: 0;
  }
  .namebox {
    position: relative;
    display: flex;
    width: 110px;
    border: 1px solid var(--xl-border-strong);
    border-radius: 3px;
    height: 24px;
  }
  .name-input {
    flex: 1;
    min-width: 0;
    border: 0;
    font: inherit;
    padding: 0 6px;
    outline: none;
    background: transparent;
  }
  .nb-arrow {
    border: 0;
    background: transparent;
    cursor: pointer;
    padding: 0 3px;
  }
  .names {
    position: absolute;
    top: 26px;
    left: 0;
    min-width: 160px;
  }
  .empty {
    padding: 6px 12px;
    color: var(--xl-text-3);
  }
  .buttons {
    display: flex;
    align-items: center;
    border-right: 1px solid var(--xl-border);
    padding-right: 4px;
  }
  .fx {
    font-family: Georgia, serif;
    font-size: 13px;
  }
  .formula {
    flex: 1;
    resize: none;
    border: 1px solid var(--xl-border-strong);
    border-radius: 3px;
    font: 13px var(--xl-font);
    padding: 3px 6px;
    height: 24px;
    line-height: 16px;
    overflow: hidden;
    white-space: pre-wrap;
    outline: none;
  }
  .formula:focus {
    border-color: var(--xl-accent);
  }
  .formula.expanded {
    height: 84px;
    overflow: auto;
  }
  .expand {
    align-self: flex-start;
  }
</style>
