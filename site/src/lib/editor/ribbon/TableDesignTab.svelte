<script lang="ts">
  // Table (Table Design), the contextual tab of a cell inside an Excel table:
  // name and size, Convert to Range, the style options and the style gallery.
  import { getEditor } from '../core/context.ts';
  import { selectRange } from '../core/selection.ts';
  import { BUILT_IN_TABLE_STYLES, hasBuiltInStyle, tableOptions } from '../core/table-style.ts';
  import {
    convertToRange,
    headerRows,
    renameTable,
    setFilterButton,
    setHeaderRow,
    setStyleOption,
    setTableStyle,
    setTotalRow,
    styleOption,
    tableRange,
    totalRows,
    type StyleOption,
  } from '../core/tables.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import MenuButton from '../ui/MenuButton.svelte';
  import Group from './Group.svelte';
  import TableStyleSwatch from './TableStyleSwatch.svelte';

  let { tableIndex }: { tableIndex: number } = $props();

  const ctl = getEditor();
  const doc = ctl.doc;

  const def = $derived.by(() => {
    void doc.version;
    return doc.ws.tables[tableIndex];
  });
  const palette = $derived(doc.styles.palette);
  const opts = $derived.by(() => {
    void doc.version;
    return def ? tableOptions(def) : { rowStripes: true, colStripes: false, firstCol: false, lastCol: false };
  });
  const flags = $derived.by(() => {
    void doc.version;
    return {
      header: def ? headerRows(def) > 0 : false,
      total: def ? totalRows(def) > 0 : false,
      filter: def?.autoFilter !== undefined,
    };
  });

  let nameText = $state('');
  $effect(() => {
    nameText = def?.displayName ?? '';
  });

  function commitName(): void {
    if (!def || nameText === def.displayName) return;
    const err = renameTable(ctl, def, nameText.trim());
    if (err) {
      ctl.openDialog('alert', { message: err });
      nameText = def.displayName;
    }
  }

  const GALLERY: ReadonlyArray<{ label: MessageKey; styles: readonly string[] }> = [
    { label: 'dtStyleLight', styles: BUILT_IN_TABLE_STYLES.filter((s) => s.startsWith('TableStyleLight')) },
    { label: 'dtStyleMedium', styles: BUILT_IN_TABLE_STYLES.filter((s) => s.startsWith('TableStyleMedium')) },
    { label: 'dtStyleDark', styles: BUILT_IN_TABLE_STYLES.filter((s) => s.startsWith('TableStyleDark')) },
  ];
  /** The ribbon shows the run of seven the current style belongs to, as Excel scrolls its gallery to it. */
  const strip = $derived.by(() => {
    const current = def?.styleInfo?.name;
    const i = current ? BUILT_IN_TABLE_STYLES.indexOf(current) : -1;
    const start = i < 0 ? 21 : i - (Number(/\d+$/.exec(current ?? '')?.[0] ?? 1) - 1) % 7;
    return BUILT_IN_TABLE_STYLES.slice(start, start + 7);
  });

  function styleTitle(name: string): string {
    const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name);
    if (!m?.[1] || !m[2]) return name;
    const family: MessageKey = m[1] === 'Light' ? 'dtStyleLight' : m[1] === 'Medium' ? 'dtStyleMedium' : 'dtStyleDark';
    return `${t(family)} ${m[2]}`;
  }

  function removeDuplicates(): void {
    if (!def) return;
    const range = tableRange(def);
    if (!range) return;
    doc.setSelection(selectRange({ ...range, r2: range.r2 - totalRows(def) }, doc.selection.active));
    ctl.openDialog('removeDuplicates');
  }

  function convert(): void {
    const target = def;
    if (!target) return;
    ctl.openDialog('alert', { message: 'dtConvertConfirm', onConfirm: () => convertToRange(ctl, target) });
  }
</script>

{#snippet option(key: StyleOption, label: MessageKey)}
  {#if def}
    <label class="check"><input type="checkbox" checked={styleOption(def, key)} onchange={(e) => setStyleOption(ctl, def, key, e.currentTarget.checked)} />{t(label)}</label>
  {/if}
{/snippet}

{#if def}
  <Group label={t('dtGroupProperties')}>
    <div class="col">
      <label class="name">
        <span>{t('dtTableName')}</span>
        <input
          class="xl-input"
          bind:value={nameText}
          spellcheck="false"
          onkeydown={(e) => {
            e.stopPropagation();
            if (e.key === 'Enter') commitName();
            else if (e.key === 'Escape') nameText = def?.displayName ?? '';
          }}
          onblur={commitName}
        />
      </label>
      <button class="xl-btn" onclick={() => ctl.openDialog('resizeTable', { table: tableIndex })}><Icon name="table" size={16} />{t('dtResizeTable')}</button>
    </div>
  </Group>

  <Group label={t('dtGroupTools')}>
    <div class="col">
      <button class="xl-btn" onclick={removeDuplicates}><Icon name="duplicates" size={16} />{t('removeDuplicates')}</button>
      <button class="xl-btn" onclick={convert}><Icon name="grid" size={16} />{t('dtConvertToRange')}</button>
    </div>
  </Group>

  <Group label={t('dtGroupStyleOptions')}>
    <div class="options">
      <label class="check"><input type="checkbox" checked={flags.header} onchange={(e) => setHeaderRow(ctl, def, e.currentTarget.checked)} />{t('dtHeaderRow')}</label>
      {@render option('firstCol', 'dtFirstColumn')}
      <label class="check"><input type="checkbox" checked={flags.filter} disabled={!flags.header} onchange={(e) => setFilterButton(ctl, def, e.currentTarget.checked)} />{t('dtFilterButton')}</label>
      <label class="check"><input type="checkbox" checked={flags.total} onchange={(e) => setTotalRow(ctl, def, e.currentTarget.checked, t('dtTotalLabel'))} />{t('dtTotalRow')}</label>
      {@render option('lastCol', 'dtLastColumn')}
      <span></span>
      {@render option('bandedRows', 'dtBandedRows')}
      {@render option('bandedCols', 'dtBandedColumns')}
    </div>
  </Group>

  <Group label={t('dtGroupStyles')}>
    <div class="strip">
      {#each strip as name (name)}
        <button class="swatch-btn" class:on={def.styleInfo?.name === name} title={styleTitle(name)} aria-label={styleTitle(name)} aria-pressed={def.styleInfo?.name === name} onclick={() => setTableStyle(ctl, def, name)}>
          <TableStyleSwatch style={name} {palette} {opts} header={flags.header} total={flags.total} />
        </button>
      {/each}
    </div>
    <MenuButton icon="chevron-down" title={t('dtGroupStyles')} align="right">
      {#snippet menu(close)}
        <div class="gallery">
          {#if def.styleInfo?.name && !hasBuiltInStyle(def.styleInfo.name)}
            <div class="section">{t('dtCustomStyle')}: {def.styleInfo.name}</div>
          {/if}
          {#each GALLERY as section (section.label)}
            <div class="section">{t(section.label)}</div>
            <div class="grid">
              {#if section.label === 'dtStyleLight'}
                <button class="swatch-btn" class:on={!def.styleInfo?.name} title={t('dtStyleNone')} aria-label={t('dtStyleNone')} onclick={() => { setTableStyle(ctl, def, undefined); close(); }}>
                  <TableStyleSwatch style={undefined} {palette} {opts} header={flags.header} total={flags.total} />
                </button>
              {/if}
              {#each section.styles as name (name)}
                <button class="swatch-btn" class:on={def.styleInfo?.name === name} title={styleTitle(name)} aria-label={styleTitle(name)} onclick={() => { setTableStyle(ctl, def, name); close(); }}>
                  <TableStyleSwatch style={name} {palette} {opts} header={flags.header} total={flags.total} />
                </button>
              {/each}
            </div>
          {/each}
          <div class="xl-menu-sep"></div>
          <button class="xl-menu-item" onclick={() => { setTableStyle(ctl, def, undefined); close(); }}>{t('dtClearStyle')}</button>
        </div>
      {/snippet}
    </MenuButton>
  </Group>
{/if}

<style>
  .col {
    display: flex;
    flex-direction: column;
    gap: 4px;
    align-items: flex-start;
  }
  .name {
    display: flex;
    flex-direction: column;
    font-size: 11px;
    gap: 2px;
  }
  .name input {
    width: 120px;
  }
  .options {
    display: grid;
    grid-template-columns: repeat(3, auto);
    column-gap: 12px;
    row-gap: 2px;
    font-size: 11.5px;
  }
  .strip {
    display: flex;
    gap: 3px;
  }
  .swatch-btn {
    padding: 2px;
    border: 1px solid transparent;
    border-radius: 3px;
    background: transparent;
    cursor: pointer;
    line-height: 0;
  }
  .swatch-btn:hover {
    border-color: var(--xl-border);
  }
  .swatch-btn.on {
    border-color: var(--xl-accent);
    box-shadow: 0 0 0 1px var(--xl-accent);
  }
  .gallery {
    padding: 6px;
    max-height: 70vh;
    overflow-y: auto;
    width: 470px;
  }
  .section {
    font-size: 11px;
    font-weight: 600;
    color: var(--xl-text-2);
    padding: 6px 2px 4px;
  }
  .grid {
    display: grid;
    grid-template-columns: repeat(7, auto);
    gap: 4px;
  }
</style>
