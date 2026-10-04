<script lang="ts">
  // Insert ▸ Symbol: pick a character from a Unicode block (or Excel's
  // Special Characters list) and type it into the active cell at the caret.
  import { getEditor } from '../core/context.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  // The Unicode blocks Excel's Subset list offers most often, as [first, last] code points.
  const SUBSETS: ReadonlyArray<[MessageKey, number, number]> = [
    ['symLatin1', 0x00a1, 0x00ff],
    ['symLatinExtA', 0x0100, 0x017f],
    ['symGreek', 0x0391, 0x03c9],
    ['symCyrillic', 0x0410, 0x044f],
    ['symPunctuation', 0x2010, 0x205e],
    ['symCurrency', 0x20a0, 0x20c0],
    ['symLetterlike', 0x2100, 0x214f],
    ['symNumberForms', 0x2150, 0x218b],
    ['symArrows', 0x2190, 0x21ff],
    ['symMath', 0x2200, 0x22ff],
    ['symTechnical', 0x2300, 0x23ff],
    ['symEnclosed', 0x2460, 0x24ff],
    ['symBox', 0x2500, 0x257f],
    ['symShapes', 0x25a0, 0x25ff],
    ['symMisc', 0x2600, 0x26ff],
    ['symDingbats', 0x2700, 0x27bf],
    ['symCjkSymbols', 0x3000, 0x303f],
  ];
  const SPECIAL: ReadonlyArray<[string, MessageKey]> = [
    ['—', 'symEmDash'],
    ['–', 'symEnDash'],
    ['©', 'symCopyright'],
    ['®', 'symRegistered'],
    ['™', 'symTrademark'],
    ['§', 'symSection'],
    ['¶', 'symParagraph'],
    ['…', 'symEllipsis'],
    ['‘', 'symSingleOpen'],
    ['’', 'symSingleClose'],
    ['“', 'symDoubleOpen'],
    ['”', 'symDoubleClose'],
    [' ', 'symNbsp'],
  ];

  const ctl = getEditor();
  let tab = $state<'symbols' | 'special'>('symbols');
  let subset = $state(0);
  let chosen = $state<string | null>(null);
  const chars = $derived.by(() => {
    const [, from, to] = SUBSETS[subset] ?? SUBSETS[0] ?? ['symLatin1', 0xa1, 0xff];
    const out: string[] = [];
    for (let cp = from; cp <= to; cp++) out.push(String.fromCodePoint(cp));
    return out;
  });

  function insert(ch: string) {
    // Typing into a cell that isn't being edited starts an edit, as when Excel inserts a symbol.
    if (!ctl.edit) ctl.startEdit(undefined, 'cell');
    const e = ctl.edit;
    if (!e) return;
    ctl.setEditText(e.text.slice(0, e.selStart) + ch + e.text.slice(e.selEnd), e.selStart + ch.length);
  }

  function onok(): boolean {
    if (chosen) insert(chosen);
    return true;
  }
</script>

<Dialog title={t('symbol')} width={520} {onok} okLabel={t('symInsert')} okDisabled={!chosen}>
  <div class="tabs" role="tablist">
    <button role="tab" class="xl-btn" aria-selected={tab === 'symbols'} onclick={() => (tab = 'symbols')}>{t('symSymbols')}</button>
    <button role="tab" class="xl-btn" aria-selected={tab === 'special'} onclick={() => (tab = 'special')}>{t('symSpecial')}</button>
  </div>
  {#if tab === 'symbols'}
    <div class="row">
      <label for="sym-subset">{t('symSubset')}</label>
      <select id="sym-subset" class="xl-select" bind:value={subset}>
        {#each SUBSETS as [label], i (label)}<option value={i}>{t(label)}</option>{/each}
      </select>
    </div>
    <div class="chars" role="listbox" aria-label={t('symbol')}>
      {#each chars as ch (ch)}
        <button
          class="ch"
          role="option"
          aria-selected={chosen === ch}
          title={`U+${(ch.codePointAt(0) ?? 0).toString(16).toUpperCase().padStart(4, '0')}`}
          onclick={() => (chosen = ch)}
          ondblclick={() => {
            chosen = ch;
            insert(ch);
            ctl.closeDialog();
          }}>{ch}</button
        >
      {/each}
    </div>
    <div class="hint">{chosen ? `U+${(chosen.codePointAt(0) ?? 0).toString(16).toUpperCase().padStart(4, '0')}` : ''}</div>
  {:else}
    <div class="special" role="listbox" aria-label={t('symSpecial')}>
      {#each SPECIAL as [ch, label] (label)}
        <button class="sp" role="option" aria-selected={chosen === ch} onclick={() => (chosen = ch)}>
          <span class="glyph">{ch === ' ' ? '°' : ch}</span>{t(label)}
        </button>
      {/each}
    </div>
  {/if}
</Dialog>

<style>
  .tabs {
    display: flex;
    gap: 4px;
    margin-bottom: 8px;
  }
  .tabs [aria-selected='true'] {
    background: var(--xl-hover, #e1e1e1);
  }
  .chars {
    display: grid;
    grid-template-columns: repeat(16, 1fr);
    max-height: 220px;
    overflow-y: auto;
    margin-top: 6px;
    border: 1px solid var(--xl-border);
    background: #fff;
  }
  .ch {
    height: 28px;
    border: 0;
    border-right: 1px solid #eee;
    border-bottom: 1px solid #eee;
    background: #fff;
    font-size: 16px;
    padding: 0;
  }
  .ch[aria-selected='true'],
  .sp[aria-selected='true'] {
    background: var(--xl-accent);
    color: #fff;
  }
  .special {
    display: flex;
    flex-direction: column;
    max-height: 240px;
    overflow-y: auto;
    border: 1px solid var(--xl-border);
    background: #fff;
  }
  .sp {
    display: flex;
    gap: 12px;
    border: 0;
    background: transparent;
    padding: 3px 8px;
    text-align: left;
  }
  .glyph {
    width: 24px;
    text-align: center;
  }
</style>
