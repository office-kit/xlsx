<script lang="ts">
  import * as A from '../core/actions.ts';
  import { getEditor } from '../core/context.ts';
  import { cut as cutSel, copy as copySel, paste as pasteSel } from '../core/actions.ts';
  import { quickSort } from '../core/data.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Icon from '../ui/Icon.svelte';
  import MenuButton from '../ui/MenuButton.svelte';
  import ColorGrid from '../ui/ColorGrid.svelte';
  import ComboBox, { type ComboItem } from '../ui/ComboBox.svelte';
  import { FONT_NAMES } from '../core/fonts.ts';
  import { cssFontFamily } from '../core/render-style.ts';
  import Group from './Group.svelte';
  import { formatPainter } from '../core/format-painter.svelte.ts';
  import { NUMBER_FORMAT_PRESETS, numberFormatCategory } from '../core/number-formats.ts';
  import { CELL_STYLE_PRESETS, applyCellStylePreset } from '../core/cell-styles.ts';
  import { findNext } from '../core/find.ts';
  import { currentRange } from '../core/selection.ts';
  import { deleteTableRows, tableAt } from '../core/tables.ts';
  import type { BorderPreset } from '../core/format.ts';

  const ctl = getEditor();
  const doc = ctl.doc;

  const font = $derived.by(() => {
    void doc.version;
    void doc.selection;
    return A.activeFont(ctl);
  });
  const style = $derived.by(() => {
    void doc.version;
    void doc.selection;
    return A.activeStyle(ctl);
  });

  const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

  // Excel's font list: the theme's heading and body fonts, then every font.
  const fontItems: ComboItem[] = $derived.by(() => {
    const { major, minor } = doc.styles.themeFonts;
    return [
      { heading: t('themeFonts') },
      { value: major, label: `${major} ${t('fontHeadings')}`, font: cssFontFamily(major) },
      { value: minor, label: `${minor} ${t('fontBody')}`, font: cssFontFamily(minor) },
      { heading: t('allFonts') },
      ...FONT_NAMES.map((name) => ({ value: name, font: cssFontFamily(name) })),
    ];
  });
  const sizeItems: ComboItem[] = SIZES.map((n) => ({ value: String(n) }));

  let fontName = $state('');
  let fontSize = $state('');
  $effect(() => {
    fontName = font.name ?? 'Calibri';
    fontSize = String(font.size ?? 11);
  });

  let lastFontColor = $state('FF0000');
  let lastFillColor = $state('FFFF00');
  let lastBorder = $state<BorderPreset>('bottom');

  const BORDERS: Array<{ id: BorderPreset; key: Parameters<typeof t>[0] }> = [
    { id: 'bottom', key: 'borderBottom' },
    { id: 'top', key: 'borderTop' },
    { id: 'left', key: 'borderLeft' },
    { id: 'right', key: 'borderRight' },
    { id: 'none', key: 'borderNone' },
    { id: 'all', key: 'borderAll' },
    { id: 'outside', key: 'borderOutside' },
    { id: 'thickOutside', key: 'borderThickOutside' },
    { id: 'bottomDouble', key: 'borderBottomDouble' },
    { id: 'thickBottom', key: 'borderThickBottom' },
    { id: 'topBottom', key: 'borderTopBottom' },
    { id: 'topThickBottom', key: 'borderTopThickBottom' },
    { id: 'topDoubleBottom', key: 'borderTopDoubleBottom' },
  ];

  const numberCategory = $derived(numberFormatCategory(style.numFmt));
</script>

<Group label={t('groupClipboard')}>
  <div class="row">
    <MenuButton large split icon="paste" label={t('paste')} onmain={() => pasteSel(ctl)}>
      {#snippet menu(close)}
        <button class="xl-menu-item" onclick={() => { pasteSel(ctl); close(); }}>{t('paste')}</button>
        <button class="xl-menu-item" onclick={() => { pasteSel(ctl, 'formulas'); close(); }}>{t('pasteFormulas')}</button>
        <button class="xl-menu-item" onclick={() => { pasteSel(ctl, 'values'); close(); }}>{t('pasteValues')}</button>
        <button class="xl-menu-item" onclick={() => { pasteSel(ctl, 'valuesAndFormats'); close(); }}>{t('pasteValuesNumberFormats')}</button>
        <button class="xl-menu-item" onclick={() => { pasteSel(ctl, 'formats'); close(); }}>{t('pasteFormatting')}</button>
        <button class="xl-menu-item" onclick={() => { pasteSel(ctl, 'noBorders'); close(); }}>{t('pasteNoBorders')}</button>
        <button class="xl-menu-item" onclick={() => { pasteSel(ctl, 'columnWidths'); close(); }}>{t('pasteColumnWidths')}</button>
        <button class="xl-menu-item" onclick={() => { pasteSel(ctl, 'transpose'); close(); }}>{t('pasteTranspose')}</button>
        <div class="xl-menu-sep"></div>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('pasteSpecial'); close(); }}>{t('pasteSpecialEllipsis')}</button>
      {/snippet}
    </MenuButton>
    <div class="col">
      <button class="xl-btn" title={t('cut')} onclick={() => cutSel(ctl)}><Icon name="cut" /></button>
      <button class="xl-btn" title={t('copy')} onclick={() => copySel(ctl)}><Icon name="copy" /></button>
      <button class="xl-btn" title={t('formatPainter')} aria-pressed={formatPainter.active} onclick={() => formatPainter.toggle(ctl, false)} ondblclick={() => formatPainter.toggle(ctl, true)}><Icon name="brush" /></button>
    </div>
  </div>
</Group>

<Group label={t('groupFont')}>
  <div class="col">
    <div class="row">
      <ComboBox
        bind:value={fontName}
        items={fontItems}
        label={t('fontName')}
        width={118}
        oncommit={(name) => {
          A.setFontName(ctl, name);
          ctl.gridFocusRequest++;
        }}
      />
      <span class="size">
        <ComboBox
          bind:value={fontSize}
          items={sizeItems}
          label={t('fontSize')}
          width={52}
          oncommit={(text) => {
            const n = Number(text);
            if (n > 0 && n <= 409) A.setFontSize(ctl, n);
            else fontSize = String(font.size ?? 11);
            ctl.gridFocusRequest++;
          }}
        />
      </span>
      <button class="xl-btn" title={t('increaseFontSize')} onclick={() => A.stepFontSize(ctl, 1)}><Icon name="font-grow" /></button>
      <button class="xl-btn" title={t('decreaseFontSize')} onclick={() => A.stepFontSize(ctl, -1)}><Icon name="font-shrink" /></button>
    </div>
    <div class="row">
      <button class="xl-btn" title={t('bold')} aria-pressed={!!font.bold} onclick={() => A.toggleBold(ctl)}><Icon name="bold" /></button>
      <button class="xl-btn" title={t('italic')} aria-pressed={!!font.italic} onclick={() => A.toggleItalic(ctl)}><Icon name="italic" /></button>
      <MenuButton split icon="underline" title={t('underline')} onmain={() => A.toggleUnderline(ctl)}>
        {#snippet menu(close)}
          <button class="xl-menu-item" onclick={() => { A.toggleUnderline(ctl, 'single'); close(); }}>{t('underline')}</button>
          <button class="xl-menu-item" onclick={() => { A.toggleUnderline(ctl, 'double'); close(); }}>{t('doubleUnderline')}</button>
        {/snippet}
      </MenuButton>
      <MenuButton split icon="borders" title={t('borders')} onmain={() => A.applyBorder(ctl, lastBorder)}>
        {#snippet menu(close)}
          {#each BORDERS as b (b.id)}
            <button class="xl-menu-item" onclick={() => { lastBorder = b.id; A.applyBorder(ctl, b.id); close(); }}>{t(b.key)}</button>
          {/each}
          <div class="xl-menu-sep"></div>
          <button class="xl-menu-item" onclick={() => { ctl.openDialog('formatCells', { tab: 'border' }); close(); }}>{t('moreBorders')}</button>
        {/snippet}
      </MenuButton>
      <MenuButton split icon="fill" title={t('fillColor')} onmain={() => A.setFillColor(ctl, lastFillColor)}>
        {#snippet menu(close)}
          <ColorGrid palette={doc.styles.palette} noneLabel={t('noFill')} onpick={(rgb) => { if (rgb) lastFillColor = rgb; A.setFillColor(ctl, rgb); close(); }} />
        {/snippet}
      </MenuButton>
      <span class="swatch fill" style:background="#{lastFillColor}"></span>
      <MenuButton split icon="font-color" title={t('fontColor')} onmain={() => A.setFontColor(ctl, lastFontColor)}>
        {#snippet menu(close)}
          <ColorGrid palette={doc.styles.palette} noneLabel={t('automatic')} onpick={(rgb) => { if (rgb) lastFontColor = rgb; A.setFontColor(ctl, rgb); close(); }} />
        {/snippet}
      </MenuButton>
      <span class="swatch" style:background="#{lastFontColor}"></span>
    </div>
  </div>
</Group>

<Group label={t('groupAlignment')}>
  <div class="col">
    <div class="row">
      <button class="xl-btn" title={t('alignTop')} aria-pressed={style.vAlign === 'top'} onclick={() => A.setVAlign(ctl, 'top')}><Icon name="align-top" /></button>
      <button class="xl-btn" title={t('alignMiddle')} aria-pressed={style.vAlign === 'center'} onclick={() => A.setVAlign(ctl, 'center')}><Icon name="align-middle" /></button>
      <button class="xl-btn" title={t('alignBottom')} aria-pressed={style.vAlign === 'bottom'} onclick={() => A.setVAlign(ctl, 'bottom')}><Icon name="align-bottom" /></button>
      <MenuButton icon="orientation" title={t('orientation')}>
        {#snippet menu(close)}
          <button class="xl-menu-item" onclick={() => { A.setRotation(ctl, 45); close(); }}>{t('angleCounterclockwise')}</button>
          <button class="xl-menu-item" onclick={() => { A.setRotation(ctl, 135); close(); }}>{t('angleClockwise')}</button>
          <button class="xl-menu-item" onclick={() => { A.setRotation(ctl, 255); close(); }}>{t('verticalText')}</button>
          <button class="xl-menu-item" onclick={() => { A.setRotation(ctl, 90); close(); }}>{t('rotateUp')}</button>
          <button class="xl-menu-item" onclick={() => { A.setRotation(ctl, 180); close(); }}>{t('rotateDown')}</button>
          <button class="xl-menu-item" onclick={() => { A.setRotation(ctl, 0); close(); }}>{t('noRotation')}</button>
          <div class="xl-menu-sep"></div>
          <button class="xl-menu-item" onclick={() => { ctl.openDialog('formatCells', { tab: 'alignment' }); close(); }}>{t('formatCellAlignment')}</button>
        {/snippet}
      </MenuButton>
      <button class="xl-btn" title={t('wrapText')} aria-pressed={style.wrap} onclick={() => A.toggleWrap(ctl)}><Icon name="wrap" /></button>
      <MenuButton title={t('wrapText')}>
        {#snippet menu(close)}
          <button class="xl-menu-item" role="menuitemcheckbox" aria-checked={style.wrap} onclick={() => { A.toggleWrap(ctl); close(); }}>{t('wrapText')}</button>
          <button class="xl-menu-item" role="menuitemcheckbox" aria-checked={style.shrink} onclick={() => { A.toggleShrink(ctl); close(); }}>{t('shrinkTextToFit')}</button>
        {/snippet}
      </MenuButton>
    </div>
    <div class="row">
      <button class="xl-btn" title={t('alignLeft')} aria-pressed={style.hAlign === 'left'} onclick={() => A.setHAlign(ctl, 'left')}><Icon name="align-left" /></button>
      <button class="xl-btn" title={t('center')} aria-pressed={style.hAlign === 'center'} onclick={() => A.setHAlign(ctl, 'center')}><Icon name="align-center" /></button>
      <button class="xl-btn" title={t('alignRight')} aria-pressed={style.hAlign === 'right'} onclick={() => A.setHAlign(ctl, 'right')}><Icon name="align-right" /></button>
      <button class="xl-btn" title={t('decreaseIndent')} onclick={() => A.stepIndent(ctl, -1)}><Icon name="indent-out" /></button>
      <button class="xl-btn" title={t('increaseIndent')} onclick={() => A.stepIndent(ctl, 1)}><Icon name="indent-in" /></button>
      <MenuButton split icon="merge" title={t('mergeCenter')} onmain={() => A.merge(ctl, 'mergeCenter')}>
        {#snippet menu(close)}
          <button class="xl-menu-item" onclick={() => { A.merge(ctl, 'mergeCenter'); close(); }}>{t('mergeCenter')}</button>
          <button class="xl-menu-item" onclick={() => { A.merge(ctl, 'mergeAcross'); close(); }}>{t('mergeAcross')}</button>
          <button class="xl-menu-item" onclick={() => { A.merge(ctl, 'merge'); close(); }}>{t('mergeCells')}</button>
          <button class="xl-menu-item" onclick={() => { A.merge(ctl, 'unmerge'); close(); }}>{t('unmergeCells')}</button>
        {/snippet}
      </MenuButton>
    </div>
  </div>
</Group>

<Group label={t('groupNumber')}>
  <div class="col">
    <div class="row">
      <select class="xl-select numfmt" aria-label={t('numberFormat')} value={numberCategory} onchange={(e) => {
        const id = (e.currentTarget as HTMLSelectElement).value;
        const preset = NUMBER_FORMAT_PRESETS.find((p) => p.id === id);
        if (id === 'more') ctl.openDialog('formatCells', { tab: 'number' });
        else if (preset) A.setNumberFormat(ctl, preset.code(ctl.dateOrder()));
        // Otherwise focus stays on the <select> and the next typed letter picks another format.
        ctl.gridFocusRequest++;
      }}>
        {#each NUMBER_FORMAT_PRESETS as p (p.id)}<option value={p.id}>{t(p.label)}</option>{/each}
        {#if numberCategory === 'custom'}<option value="custom">{t('nfCustom')}</option>{/if}
        <option value="more">{t('moreNumberFormats')}</option>
      </select>
    </div>
    <div class="row">
      <MenuButton split icon="currency" title={t('accountingFormat')} onmain={() => A.setNumberFormat(ctl, ctl.dateOrder() === 'ymd' ? '_ "¥"* #,##0_ ;_ "¥"* \\-#,##0_ ;_ "¥"* "-"_ ;_ @_ ' : '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)')}>
        {#snippet menu(close)}
          <button class="xl-menu-item" onclick={() => { A.setNumberFormat(ctl, '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'); close(); }}>$ English (United States)</button>
          <button class="xl-menu-item" onclick={() => { A.setNumberFormat(ctl, '_ "¥"* #,##0_ ;_ "¥"* \\-#,##0_ ;_ "¥"* "-"_ ;_ @_ '); close(); }}>¥ 日本語</button>
          <button class="xl-menu-item" onclick={() => { A.setNumberFormat(ctl, '_-* #,##0.00\\ "€"_-;\\-* #,##0.00\\ "€"_-;_-* "-"??\\ "€"_-;_-@_-'); close(); }}>€ Euro</button>
          <button class="xl-menu-item" onclick={() => { A.setNumberFormat(ctl, '_-"£"* #,##0.00_-;\\-"£"* #,##0.00_-;_-"£"* "-"??_-;_-@_-'); close(); }}>£ English (United Kingdom)</button>
          <div class="xl-menu-sep"></div>
          <button class="xl-menu-item" onclick={() => { ctl.openDialog('formatCells', { tab: 'number' }); close(); }}>{t('moreAccountingFormats')}</button>
        {/snippet}
      </MenuButton>
      <button class="xl-btn" title={t('percentStyle')} onclick={() => A.setNumberFormat(ctl, '0%')}><Icon name="percent" /></button>
      <button class="xl-btn" title={t('commaStyle')} onclick={() => A.setNumberFormat(ctl, '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)')}><Icon name="comma" /></button>
      <button class="xl-btn" title={t('increaseDecimal')} onclick={() => A.stepDecimals(ctl, 1)}><Icon name="dec-inc" /></button>
      <button class="xl-btn" title={t('decreaseDecimal')} onclick={() => A.stepDecimals(ctl, -1)}><Icon name="dec-dec" /></button>
    </div>
  </div>
</Group>

<Group label={t('groupStyles')}>
  <div class="col">
    <MenuButton icon="cond-format" label={t('conditionalFormatting')}>
      {#snippet menu(close)}
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('conditionalFormatting', { preset: 'highlight' }); close(); }}>{t('cfHighlightRules')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('conditionalFormatting', { preset: 'topBottom' }); close(); }}>{t('cfTopBottomRules')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('conditionalFormatting', { preset: 'dataBar' }); close(); }}>{t('cfDataBars')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('conditionalFormatting', { preset: 'colorScale' }); close(); }}>{t('cfColorScales')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('conditionalFormatting', { preset: 'iconSet' }); close(); }}>{t('cfIconSets')}</button>
        <div class="xl-menu-sep"></div>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('conditionalFormatting', { preset: 'new' }); close(); }}>{t('cfNewRule')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('conditionalFormatting', { preset: 'clear' }); close(); }}>{t('cfClearRules')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('conditionalFormatting', { preset: 'manage' }); close(); }}>{t('cfManageRules')}</button>
      {/snippet}
    </MenuButton>
    <button class="xl-btn" onclick={() => ctl.openDialog('createTable')}><Icon name="table" /><span>{t('formatAsTable')}</span></button>
    <MenuButton icon="cell-styles" label={t('cellStyles')}>
      {#snippet menu(close)}
        <div class="styles">
          {#each CELL_STYLE_PRESETS as p (p.id)}
            <button class="style-chip" style:background={p.preview.fill} style:color={p.preview.color} style:font-weight={p.preview.bold ? 700 : 400} style:border-bottom={p.preview.border ?? 'none'} onclick={() => { applyCellStylePreset(ctl, p.id); close(); }}>{t(p.label)}</button>
          {/each}
        </div>
      {/snippet}
    </MenuButton>
  </div>
</Group>

<Group label={t('groupCells')}>
  <div class="col">
    <MenuButton split icon="insert" label={t('insert')} onmain={() => A.insertCellsSmart(ctl)}>
      {#snippet menu(close)}
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('insertCells'); close(); }}>{t('insertCellsEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => { A.insertLines(ctl, 'row'); close(); }}>{t('insertSheetRows')}</button>
        <button class="xl-menu-item" onclick={() => { A.insertLines(ctl, 'col'); close(); }}>{t('insertSheetColumns')}</button>
        <button class="xl-menu-item" onclick={() => { A.insertSheet(ctl); close(); }}>{t('insertSheet')}</button>
      {/snippet}
    </MenuButton>
    <MenuButton split icon="delete" label={t('delete')} onmain={() => A.deleteCellsSmart(ctl)}>
      {#snippet menu(close)}
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('deleteCells'); close(); }}>{t('deleteCellsEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => { A.deleteLines(ctl, 'row'); close(); }}>{t('deleteSheetRows')}</button>
        <button class="xl-menu-item" onclick={() => { A.deleteLines(ctl, 'col'); close(); }}>{t('deleteSheetColumns')}</button>
        {#if tableAt(doc.ws, doc.selection.active.row, doc.selection.active.col)}
          <button class="xl-menu-item" onclick={() => { deleteTableRows(ctl, currentRange(doc.selection)); close(); }}>{t('deleteTableRows')}</button>
        {/if}
        <button class="xl-menu-item" onclick={() => { A.deleteSheet(ctl, doc.activeSheetIndex); close(); }}>{t('deleteSheet')}</button>
      {/snippet}
    </MenuButton>
    <MenuButton icon="format" label={t('format')}>
      {#snippet menu(close)}
        <div class="menu-head">{t('cellSize')}</div>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('rowHeight'); close(); }}>{t('rowHeightEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => { A.autofitSelection(ctl, 'row'); close(); }}>{t('autofitRowHeight')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('columnWidth'); close(); }}>{t('columnWidthEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => { A.autofitSelection(ctl, 'col'); close(); }}>{t('autofitColumnWidth')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('standardWidth'); close(); }}>{t('defaultWidthEllipsis')}</button>
        <div class="menu-head">{t('visibility')}</div>
        <button class="xl-menu-item" onclick={() => { A.hideLines(ctl, 'row', true); close(); }}>{t('hideRows')}</button>
        <button class="xl-menu-item" onclick={() => { A.hideLines(ctl, 'col', true); close(); }}>{t('hideColumns')}</button>
        <button class="xl-menu-item" onclick={() => { A.setSheetHidden(ctl, doc.activeSheetIndex, true); close(); }}>{t('hideSheet')}</button>
        <button class="xl-menu-item" onclick={() => { A.hideLines(ctl, 'row', false); close(); }}>{t('unhideRows')}</button>
        <button class="xl-menu-item" onclick={() => { A.hideLines(ctl, 'col', false); close(); }}>{t('unhideColumns')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('unhideSheet'); close(); }}>{t('unhideSheetEllipsis')}</button>
        <div class="menu-head">{t('organizeSheets')}</div>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('renameSheet', { index: doc.activeSheetIndex }); close(); }}>{t('renameSheet')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('moveCopySheet', { index: doc.activeSheetIndex }); close(); }}>{t('moveCopySheetEllipsis')}</button>
        <div class="menu-head">{t('protection')}</div>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('protectSheet'); close(); }}>{t('protectSheetEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => { A.toggleLocked(ctl); close(); }}>{t('lockCell')}</button>
        <div class="xl-menu-sep"></div>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('formatCells'); close(); }}>{t('formatCellsEllipsis')}</button>
      {/snippet}
    </MenuButton>
  </div>
</Group>

<Group label={t('groupEditing')}>
  <div class="col">
    <MenuButton split icon="sum" title={t('autoSum')} onmain={() => A.autoSum(ctl)}>
      {#snippet menu(close)}
        {#each [['SUM', 'fnSum'], ['AVERAGE', 'fnAverage'], ['COUNT', 'fnCount'], ['MAX', 'fnMax'], ['MIN', 'fnMin']] as const as [fn, key] (fn)}
          <button class="xl-menu-item" onclick={() => { A.autoSum(ctl, fn); close(); }}>{t(key)}</button>
        {/each}
        <div class="xl-menu-sep"></div>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('insertFunction'); close(); }}>{t('moreFunctions')}</button>
      {/snippet}
    </MenuButton>
    <MenuButton icon="fill-down" title={t('fill')}>
      {#snippet menu(close)}
        <button class="xl-menu-item" onclick={() => { A.fillFrom(ctl, 'down'); close(); }}>{t('fillDown')}</button>
        <button class="xl-menu-item" onclick={() => { A.fillFrom(ctl, 'right'); close(); }}>{t('fillRight')}</button>
        <button class="xl-menu-item" onclick={() => { A.fillFrom(ctl, 'up'); close(); }}>{t('fillUp')}</button>
        <button class="xl-menu-item" onclick={() => { A.fillFrom(ctl, 'left'); close(); }}>{t('fillLeft')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('series'); close(); }}>{t('seriesEllipsis')}</button>
      {/snippet}
    </MenuButton>
    <MenuButton icon="eraser" title={t('clear')}>
      {#snippet menu(close)}
        <button class="xl-menu-item" onclick={() => { A.clear(ctl, 'all'); close(); }}>{t('clearAll')}</button>
        <button class="xl-menu-item" onclick={() => { A.clear(ctl, 'formats'); close(); }}>{t('clearFormats')}</button>
        <button class="xl-menu-item" onclick={() => { A.clear(ctl, 'contents'); close(); }}>{t('clearContents')}</button>
        <button class="xl-menu-item" onclick={() => { A.clear(ctl, 'comments'); close(); }}>{t('clearComments')}</button>
        <button class="xl-menu-item" onclick={() => { A.clear(ctl, 'hyperlinks'); close(); }}>{t('clearHyperlinks')}</button>
        <button class="xl-menu-item" onclick={() => { A.clear(ctl, 'removeHyperlinks'); close(); }}>{t('removeHyperlinks')}</button>
      {/snippet}
    </MenuButton>
  </div>
  <div class="row">
    <MenuButton large icon="sort" label={t('sortFilter')}>
      {#snippet menu(close)}
        <button class="xl-menu-item" onclick={() => { quickSort(ctl, false); close(); }}><Icon name="sort-asc" size={16} />{t('sortAZ')}</button>
        <button class="xl-menu-item" onclick={() => { quickSort(ctl, true); close(); }}><Icon name="sort-desc" size={16} />{t('sortZA')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('sort'); close(); }}>{t('customSortEllipsis')}</button>
        <div class="xl-menu-sep"></div>
        <button class="xl-menu-item" onclick={() => { ctl.toggleFilter(); close(); }}><Icon name="filter" size={16} />{t('filter')}</button>
      {/snippet}
    </MenuButton>
    <MenuButton large icon="search" label={t('findSelect')} align="right">
      {#snippet menu(close)}
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('find'); close(); }}>{t('findEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => { findNext(ctl, 1); close(); }}>{t('findNext')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('replace'); close(); }}>{t('replaceEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('goto'); close(); }}>{t('gotoEllipsis')}</button>
        <button class="xl-menu-item" onclick={() => { ctl.openDialog('gotoSpecial'); close(); }}>{t('gotoSpecialEllipsis')}</button>
      {/snippet}
    </MenuButton>
  </div>
</Group>

<style>
  .row {
    display: flex;
    align-items: center;
    gap: 1px;
  }
  .size {
    margin-left: 2px;
  }
  .numfmt {
    width: 150px;
  }
  .swatch {
    width: 14px;
    height: 3px;
    margin-left: -24px;
    margin-right: 10px;
    margin-top: 16px;
    pointer-events: none;
  }
  .menu-head {
    font-weight: 600;
    font-size: 11px;
    color: var(--xl-text-2);
    padding: 6px 10px 2px;
    background: #f7f7f7;
  }
  .styles {
    display: grid;
    grid-template-columns: repeat(4, 110px);
    gap: 4px;
    padding: 6px;
  }
  .style-chip {
    border: 1px solid var(--xl-border);
    padding: 6px 4px;
    font: inherit;
    font-size: 11.5px;
    text-align: left;
    cursor: pointer;
  }
  .style-chip:hover {
    outline: 2px solid #f5a623;
  }
</style>
