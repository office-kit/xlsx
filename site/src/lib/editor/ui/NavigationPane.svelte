<script lang="ts">
  // View ▸ Navigation: sheets and the elements on them; clicking one goes there.
  import { getEditor } from '../core/context.ts';
  import { navigationTree, type NavElement, type NavSheet } from '../core/navigation.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Icon from './Icon.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;

  const sheets = $derived.by(() => {
    void doc.version;
    return navigationTree(doc.wb);
  });
  let collapsed = $state(new Set<number>());

  const KIND_LABEL: Record<NavElement['kind'], MessageKey> = {
    table: 'navTable',
    name: 'navName',
    pivot: 'navPivot',
    chart: 'navChart',
    picture: 'navPicture',
    shape: 'navShape',
  };

  function label(e: NavElement): string {
    return e.label ?? `${t(KIND_LABEL[e.kind])} ${e.ordinal}`;
  }

  function goSheet(s: NavSheet) {
    if (s.hidden) return;
    if (s.index !== doc.activeSheetIndex) doc.activateSheet(s.index);
  }

  function go(s: NavSheet, e: NavElement) {
    if (s.hidden) return;
    goSheet(s);
    if (e.target.kind === 'range') ctl.selectRange(e.target.range);
    else ctl.selectedDrawing = e.target.index;
  }

  function toggle(index: number) {
    const next = new Set(collapsed);
    if (!next.delete(index)) next.add(index);
    collapsed = next;
  }
</script>

<aside class="pane" aria-label={t('navigation')}>
  <header>
    <span class="title">{t('navigation')}</span>
    <button class="xl-btn icon" title={t('dlgClose')} aria-label={t('dlgClose')} onclick={() => (ctl.navigationPane = false)}><Icon name="close" size={14} /></button>
  </header>
  <ul class="tree" role="tree">
    {#each sheets as s (s.index)}
      <li role="treeitem" aria-selected={s.index === doc.activeSheetIndex} aria-expanded={s.elements.length > 0 ? !collapsed.has(s.index) : undefined}>
        <div class="row" class:active={s.index === doc.activeSheetIndex} class:hidden={s.hidden}>
          {#if s.elements.length > 0}
            <button class="twisty" aria-label={s.title} onclick={() => toggle(s.index)}>{collapsed.has(s.index) ? '▸' : '▾'}</button>
          {:else}
            <span class="twisty"></span>
          {/if}
          <button class="name" disabled={s.hidden} onclick={() => goSheet(s)}>{s.title}{s.hidden ? ` (${t('navHidden')})` : ''}</button>
        </div>
        {#if s.elements.length > 0 && !collapsed.has(s.index)}
          <ul role="group">
            {#each s.elements as e, i (i)}
              <li role="treeitem" aria-selected={false}>
                <button class="name element" disabled={s.hidden} onclick={() => go(s, e)}>
                  <span class="kind">{t(KIND_LABEL[e.kind])}</span>{label(e)}
                </button>
              </li>
            {/each}
          </ul>
        {/if}
      </li>
    {/each}
  </ul>
</aside>

<style>
  .pane {
    display: flex;
    flex-direction: column;
    width: 260px;
    flex: none;
    min-height: 0;
    background: var(--xl-panel);
    border-left: 1px solid var(--xl-border);
  }
  header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 6px 10px 2px;
  }
  .title {
    font-size: 15px;
    font-weight: 600;
  }
  .tree,
  .tree ul {
    list-style: none;
    margin: 0;
    padding: 0;
  }
  .tree {
    overflow: auto;
    padding: 4px 6px 10px;
    font-size: 12px;
  }
  .tree ul {
    padding-left: 22px;
  }
  .row {
    display: flex;
    align-items: center;
  }
  .row.active > .name {
    font-weight: 600;
  }
  .row.hidden {
    opacity: 0.6;
  }
  .twisty {
    width: 18px;
    flex: none;
    border: 0;
    background: transparent;
    cursor: pointer;
    font: inherit;
  }
  .name {
    flex: 1;
    min-width: 0;
    border: 0;
    background: transparent;
    padding: 3px 4px;
    text-align: left;
    font: inherit;
    cursor: pointer;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    border-radius: 3px;
  }
  .name:hover:not(:disabled) {
    background: var(--xl-hover, #e8f0e9);
  }
  .kind {
    display: inline-block;
    min-width: 64px;
    color: #666;
  }
</style>
