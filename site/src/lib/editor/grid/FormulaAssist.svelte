<script lang="ts">
  // Formula AutoComplete list and the argument ScreenTip, placed under the
  // cell being edited (Excel puts both just below the edit).
  import { getEditor } from '../core/context.ts';
  import { argumentHintAt, currentArgIndex, syntaxParts } from '../core/formula-assist.ts';

  const ctl = getEditor();
  const geo = $derived(ctl.geometry);

  const anchor = $derived.by(() => {
    const e = ctl.edit;
    if (!e || e.sheetIndex !== ctl.doc.activeSheetIndex) return undefined;
    const r = geo.rectOf({ r1: e.row, c1: e.col, r2: e.row, c2: e.col });
    return { x: Math.max(geo.headerW, r.x), y: Math.min(geo.height - 40, r.y + r.h + 2) };
  });

  const hint = $derived.by(() => {
    const e = ctl.edit;
    if (!e || ctl.completion) return undefined;
    const h = argumentHintAt(e.text, e.selEnd);
    if (!h) return undefined;
    const { head, args } = syntaxParts(h.fn.syntax);
    return { head, args, current: currentArgIndex(args, h.argIndex) };
  });
</script>

{#if anchor && ctl.completion}
  {@const items = ctl.completion.items}
  {@const active = items[ctl.completionIndex]}
  <div class="assist" style:left="{anchor.x}px" style:top="{anchor.y}px">
    <div class="list" role="listbox" aria-label="AutoComplete">
      {#each items as item, i (item.name)}
        <button
          class="item"
          class:selected={i === ctl.completionIndex}
          role="option"
          aria-selected={i === ctl.completionIndex}
          tabindex="-1"
          onpointerdown={(ev) => {
            // Keep focus (and the caret) in the editor.
            ev.preventDefault();
            ev.stopPropagation();
          }}
          onclick={() => (ctl.completionIndex = i)}
          ondblclick={() => ctl.acceptCompletion(i)}
        >
          <span class="icon" aria-hidden="true">{item.kind === 'function' ? 'fx' : '⊞'}</span>{item.name}
        </button>
      {/each}
    </div>
    {#if active?.detail}<div class="detail">{active.detail}</div>{/if}
  </div>
{:else if anchor && hint}
  <div class="tip" style:left="{anchor.x}px" style:top="{anchor.y}px" role="tooltip">
    {hint.head}{#each hint.args as arg, i (i)}{#if i > 0},&nbsp;{/if}{#if i === hint.current}<b>{arg}</b>{:else}{arg}{/if}{/each})
  </div>
{/if}

<style>
  .assist {
    position: absolute;
    z-index: 5;
    display: flex;
    align-items: flex-start;
    gap: 4px;
    pointer-events: none;
  }
  .list {
    pointer-events: auto;
    min-width: 180px;
    max-height: 264px;
    overflow-y: auto;
    background: #fff;
    border: 1px solid #b5b5b5;
    box-shadow: 0 2px 8px rgb(0 0 0 / 0.18);
    padding: 2px 0;
  }
  .item {
    display: flex;
    align-items: center;
    gap: 6px;
    width: 100%;
    border: 0;
    background: transparent;
    padding: 1px 8px 1px 4px;
    font: 12px/18px var(--xl-font, system-ui);
    text-align: left;
    color: #222;
  }
  .item.selected {
    background: #cfe3f7;
  }
  .icon {
    width: 16px;
    font-size: 10px;
    font-style: italic;
    color: #555;
    text-align: center;
  }
  .detail,
  .tip {
    max-width: 320px;
    background: #fffffe;
    border: 1px solid #b5b5b5;
    box-shadow: 0 2px 6px rgb(0 0 0 / 0.15);
    padding: 3px 6px;
    font-size: 11.5px;
    color: #333;
  }
  .tip {
    position: absolute;
    z-index: 5;
    white-space: nowrap;
    pointer-events: none;
  }
</style>
