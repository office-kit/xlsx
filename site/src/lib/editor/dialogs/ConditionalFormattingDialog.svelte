<script lang="ts">
  // Home ▸ Conditional Formatting. One dialog serves every entry of the menu:
  // the Highlight Cells / Top-Bottom quick rules (value + preset format), the
  // Data Bar / Color Scale / Icon Set galleries (one click applies), New
  // Formatting Rule, Clear Rules and the Rules Manager. The manager edits a
  // draft list and commits it on OK as a single undo step, like Excel's.
  import { untrack } from 'svelte';
  import type { TimePeriod } from '@office-kit/xlsx/worksheet';
  import { cellAddress, rangeAddress, rangesIntersect, type Range } from '../core/address.ts';
  import {
    addRule,
    AVERAGE_CHOICES,
    cfFormatOf,
    clearRulesIn,
    clearSheetRules,
    defaultRuleForm,
    formFromRule,
    formUsesFormat,
    ICON_SETS,
    listRules,
    operandFormula,
    operandText,
    parseAppliesTo,
    parseVisualRule,
    replaceRules,
    ruleFromForm,
    withCompanionFormula,
    type AverageChoice,
    type CfFormat,
    type ContainsKind,
    type DraftRule,
    type NewRule,
    type RuleForm,
    type RuleFormKind,
    type ValuesStyle,
  } from '../core/conditional-format.ts';
  import { getEditor } from '../core/context.ts';
  import { resolveColor } from '../core/theme.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import { makeCfRule } from '@office-kit/xlsx/worksheet';
  import ColorPicker from './ColorPicker.svelte';
  import Dialog from './Dialog.svelte';

  let { props }: { props: Record<string, unknown> | undefined } = $props();

  const ctl = getEditor();
  const doc = ctl.doc;
  const palette = doc.styles.palette;
  // DialogHost remounts the dialog for every open, so the props are read once.
  const preset = untrack(() => String(props?.['preset'] ?? 'new'));
  const ranges: readonly Range[] = doc.selection.ranges;
  const first = ranges[0] ?? { r1: 1, c1: 1, r2: 1, c2: 1 };
  const a1 = cellAddress(first.r1, first.c1);

  let notice = $state<string | null>(null);

  // ---- preset formats -----------------------------------------------------------

  const PRESET_FORMATS: ReadonlyArray<{ key: MessageKey; format: CfFormat }> = [
    { key: 'cfFmtLightRedDarkRed', format: { fill: '#FFC7CE', color: '#9C0006' } },
    { key: 'cfFmtYellowDarkYellow', format: { fill: '#FFEB9C', color: '#9C5700' } },
    { key: 'cfFmtGreenDarkGreen', format: { fill: '#C6EFCE', color: '#006100' } },
    { key: 'cfFmtLightRed', format: { fill: '#FFC7CE' } },
    { key: 'cfFmtRedText', format: { color: '#9C0006' } },
  ];
  const CUSTOM = PRESET_FORMATS.length;

  let presetIndex = $state(0);
  let custom = $state<CfFormat>({});
  const chosenFormat = $derived(presetIndex === CUSTOM ? custom : (PRESET_FORMATS[presetIndex]?.format ?? {}));

  // ---- quick rules (Highlight Cells / Top-Bottom) ------------------------------------

  type Quick = 'greater' | 'less' | 'between' | 'equal' | 'text' | 'date' | 'duplicate' | 'top' | 'topPct' | 'bottom' | 'bottomPct' | 'above' | 'below';

  const QUICK_LABEL: Record<Quick, MessageKey> = {
    greater: 'cfGreaterThan',
    less: 'cfLessThan',
    between: 'cfBetween',
    equal: 'cfEqualTo',
    text: 'cfTextContains',
    date: 'cfDateOccurring',
    duplicate: 'cfDuplicateValues',
    top: 'cfTop10Items',
    topPct: 'cfTop10Percent',
    bottom: 'cfBottom10Items',
    bottomPct: 'cfBottom10Percent',
    above: 'cfAboveAverage',
    below: 'cfBelowAverage',
  };
  const QUICK_PROMPT: Record<Quick, MessageKey> = {
    greater: 'cfPromptGreater',
    less: 'cfPromptLess',
    between: 'cfPromptBetween',
    equal: 'cfPromptEqual',
    text: 'cfPromptText',
    date: 'cfPromptDate',
    duplicate: 'cfPromptDuplicate',
    top: 'cfPromptTop',
    topPct: 'cfPromptTop',
    bottom: 'cfPromptBottom',
    bottomPct: 'cfPromptBottom',
    above: 'cfPromptAbove',
    below: 'cfPromptBelow',
  };
  const HIGHLIGHT: readonly Quick[] = ['greater', 'less', 'between', 'equal', 'text', 'date', 'duplicate'];
  const TOP_BOTTOM: readonly Quick[] = ['top', 'topPct', 'bottom', 'bottomPct', 'above', 'below'];

  const PERIODS: ReadonlyArray<[TimePeriod, MessageKey]> = [
    ['yesterday', 'cfPeriodYesterday'],
    ['today', 'cfPeriodToday'],
    ['tomorrow', 'cfPeriodTomorrow'],
    ['last7Days', 'cfPeriodLast7Days'],
    ['lastWeek', 'cfPeriodLastWeek'],
    ['thisWeek', 'cfPeriodThisWeek'],
    ['nextWeek', 'cfPeriodNextWeek'],
    ['lastMonth', 'cfPeriodLastMonth'],
    ['thisMonth', 'cfPeriodThisMonth'],
    ['nextMonth', 'cfPeriodNextMonth'],
  ];

  // Excel pre-fills the value box from the active cell.
  function activeText(): string {
    const v = doc.calc.cellValue(doc.ws.title, doc.selection.active.row, doc.selection.active.col);
    return typeof v === 'number' || typeof v === 'string' ? String(v) : '';
  }

  let quick = $state<Quick>(preset === 'topBottom' ? 'top' : 'greater');
  let q1 = $state(activeText());
  let q2 = $state('');
  let qRank = $state(10);
  let qPeriod = $state<TimePeriod>('yesterday');
  let qUnique = $state(false);

  function quickRule(): NewRule | undefined {
    switch (quick) {
      case 'greater':
      case 'less':
      case 'equal':
        if (q1.trim() === '') return undefined;
        return { type: 'cellIs', operator: quick === 'greater' ? 'greaterThan' : quick === 'less' ? 'lessThan' : 'equal', formulas: [operandFormula(q1)] };
      case 'between':
        if (q1.trim() === '' || q2.trim() === '') return undefined;
        return { type: 'cellIs', operator: 'between', formulas: [operandFormula(q1), operandFormula(q2)] };
      case 'text':
        return withCompanionFormula({ type: 'containsText', operator: 'containsText', text: q1, formulas: [] }, a1);
      case 'date':
        return withCompanionFormula({ type: 'timePeriod', timePeriod: qPeriod, formulas: [] }, a1);
      case 'duplicate':
        return { type: qUnique ? 'uniqueValues' : 'duplicateValues', formulas: [] };
      case 'top':
      case 'topPct':
      case 'bottom':
      case 'bottomPct':
        return { type: 'top10', rank: Math.max(1, Math.round(qRank)), formulas: [], ...(quick.startsWith('bottom') ? { bottom: true } : {}), ...(quick.endsWith('Pct') ? { percent: true } : {}) };
      case 'above':
        return { type: 'aboveAverage', formulas: [] };
      case 'below':
        return { type: 'aboveAverage', aboveAverage: false, formulas: [] };
    }
  }

  function applyQuick(): boolean {
    const rule = quickRule();
    if (!rule) {
      notice = t('cfNeedValue');
      return false;
    }
    addRule(doc, ranges, rule, chosenFormat);
    return true;
  }

  // ---- galleries -------------------------------------------------------------------

  const BAR_COLORS: ReadonlyArray<[string, MessageKey]> = [
    ['638EC6', 'cfBarBlue'],
    ['63C384', 'cfBarGreen'],
    ['FF555A', 'cfBarRed'],
    ['FFB628', 'cfBarOrange'],
    ['008AEF', 'cfBarLightBlue'],
    ['D6007B', 'cfBarPurple'],
  ];

  // Colours listed lowest → highest value; Excel names a scale highest-first.
  const SCALES: ReadonlyArray<[MessageKey, readonly string[]]> = [
    ['cfScaleGYR', ['F8696B', 'FFEB84', '63BE7B']],
    ['cfScaleRYG', ['63BE7B', 'FFEB84', 'F8696B']],
    ['cfScaleGWR', ['F8696B', 'FCFCFF', '63BE7B']],
    ['cfScaleRWG', ['63BE7B', 'FCFCFF', 'F8696B']],
    ['cfScaleBWR', ['F8696B', 'FCFCFF', '5A8AC6']],
    ['cfScaleRWB', ['5A8AC6', 'FCFCFF', 'F8696B']],
    ['cfScaleWR', ['F8696B', 'FCFCFF']],
    ['cfScaleRW', ['FCFCFF', 'F8696B']],
    ['cfScaleGW', ['FCFCFF', '63BE7B']],
    ['cfScaleWG', ['63BE7B', 'FCFCFF']],
    ['cfScaleGY', ['FFEF9C', '63BE7B']],
    ['cfScaleYG', ['63BE7B', 'FFEF9C']],
  ];

  const ICON_GROUPS: ReadonlyArray<[MessageKey, readonly string[]]> = [
    ['cfDirectional', ['3Arrows', '3ArrowsGray', '4Arrows', '4ArrowsGray', '5Arrows', '5ArrowsGray']],
    ['cfShapes', ['3TrafficLights1', '3TrafficLights2', '3Signs', '4TrafficLights', '4RedToBlack']],
    ['cfIndicators', ['3Symbols', '3Symbols2', '3Flags']],
    ['cfRatings', ['4Rating', '5Quarters', '5Rating']],
  ];

  function applyForm(patch: Partial<RuleForm>): void {
    const form = { ...defaultRuleForm(), kind: 'values' as const, ...patch };
    addRule(doc, ranges, ruleFromForm(form, a1));
    ctl.closeDialog();
  }

  function applyScale(colors: readonly string[]): void {
    const [lo = 'FFFFFF', mid = 'FFFFFF', hi = 'FFFFFF'] = colors;
    if (colors.length === 3) applyForm({ style: 'scale3', min: { type: 'min', val: '', color: lo }, mid: { type: 'percentile', val: '50', color: mid }, max: { type: 'max', val: '', color: hi } });
    else applyForm({ style: 'scale2', min: { type: 'min', val: '', color: lo }, max: { type: 'max', val: '', color: mid } });
  }

  // ---- rule editor (New / Edit Formatting Rule) ---------------------------------------

  const RULE_KINDS: ReadonlyArray<[RuleFormKind, MessageKey]> = [
    ['values', 'cfTypeValues'],
    ['contains', 'cfTypeContains'],
    ['topBottom', 'cfTypeTopBottom'],
    ['average', 'cfTypeAverage'],
    ['unique', 'cfTypeUnique'],
    ['formula', 'cfTypeFormula'],
  ];
  const STYLES: ReadonlyArray<[ValuesStyle, MessageKey]> = [
    ['scale2', 'cf2ColorScale'],
    ['scale3', 'cf3ColorScale'],
    ['dataBar', 'cfDataBarStyle'],
    ['iconSet', 'cfIconSetsStyle'],
  ];
  const CONTAINS: ReadonlyArray<[ContainsKind, MessageKey]> = [
    ['cellValue', 'cfCellValue'],
    ['text', 'cfSpecificText'],
    ['date', 'cfDatesOccurring'],
    ['blanks', 'cfBlanks'],
    ['noBlanks', 'cfNoBlanks'],
    ['errors', 'cfErrors'],
    ['noErrors', 'cfNoErrors'],
  ];
  const OPERATORS: ReadonlyArray<[RuleForm['operator'], MessageKey]> = [
    ['between', 'cfOpBetween'],
    ['notBetween', 'cfOpNotBetween'],
    ['equal', 'cfOpEqual'],
    ['notEqual', 'cfOpNotEqual'],
    ['greaterThan', 'cfOpGreater'],
    ['lessThan', 'cfOpLess'],
    ['greaterThanOrEqual', 'cfOpGreaterEqual'],
    ['lessThanOrEqual', 'cfOpLessEqual'],
  ];
  const TEXT_OPS: ReadonlyArray<[RuleForm['textOp'], MessageKey]> = [
    ['containsText', 'cfTextContaining'],
    ['notContains', 'cfTextNotContaining'],
    ['beginsWith', 'cfTextBeginsWith'],
    ['endsWith', 'cfTextEndsWith'],
  ];
  const CFVO_TYPES: ReadonlyArray<[RuleForm['min']['type'], MessageKey]> = [
    ['min', 'cfLowestValue'],
    ['max', 'cfHighestValue'],
    ['num', 'cfNumber'],
    ['percent', 'cfPercent'],
    ['percentile', 'cfPercentile'],
    ['formula', 'cfFormula'],
  ];

  function averageLabel(a: AverageChoice): string {
    const m = /^std(\d)(Above|Below)$/.exec(a);
    if (m) return t(m[2] === 'Above' ? 'cfAvgStdAbove' : 'cfAvgStdBelow', { n: m[1] ?? '1' });
    const keys: Record<string, MessageKey> = { above: 'cfAvgAbove', below: 'cfAvgBelow', equalAbove: 'cfAvgEqualAbove', equalBelow: 'cfAvgEqualBelow' };
    return t(keys[a] ?? 'cfAvgAbove');
  }

  let form = $state<RuleForm>(defaultRuleForm());
  let formFormat = $state<CfFormat>({});
  /** Manager draft index being edited, 'new' for a new rule from the manager, null outside the manager's editor. */
  let editing = $state<number | 'new' | null>(null);

  function formValid(): boolean {
    if (form.kind === 'formula') return form.formula.trim().replace(/^=/, '') !== '';
    if (form.kind === 'contains' && form.contains === 'cellValue') {
      const two = form.operator === 'between' || form.operator === 'notBetween';
      return form.value1.trim() !== '' && (!two || form.value2.trim() !== '');
    }
    return true;
  }

  function applyNew(): boolean {
    if (!formValid()) {
      notice = t('cfNeedValue');
      return false;
    }
    addRule(doc, ranges, ruleFromForm(form, a1), formUsesFormat(form) ? formFormat : undefined);
    return true;
  }

  // ---- Rules Manager -----------------------------------------------------------------------

  interface Row {
    appliesTo: string;
    rule: DraftRule['rule'];
    format?: CfFormat;
  }

  let rows = $state<Row[]>(listRules(doc.ws).map((e) => ({ appliesTo: `=${e.sqref.replaceAll(' ', ',')}`, rule: structuredClone(e.rule) })));
  let scope = $state<'selection' | 'sheet'>('selection');
  let selectedRow = $state<number | null>(null);
  let managerDirty = $state(false);

  const visibleRows = $derived(
    rows
      .map((row, index) => ({ row, index }))
      .filter(({ row }) => scope === 'sheet' || (parseAppliesTo(row.appliesTo) ?? []).some((r) => ranges.some((s) => rangesIntersect(r, s)))),
  );

  function move(delta: -1 | 1): void {
    const i = selectedRow;
    if (i === null) return;
    // Move past rows hidden by the scope filter so the visible order changes.
    const visible = visibleRows.map((v) => v.index);
    const pos = visible.indexOf(i);
    const target = visible[pos + delta];
    if (target === undefined) return;
    const next = rows.slice();
    const [item] = next.splice(i, 1);
    if (!item) return;
    next.splice(target, 0, item);
    rows = next;
    selectedRow = target;
    managerDirty = true;
  }

  function deleteRow(): void {
    if (selectedRow === null) return;
    rows = rows.filter((_, i) => i !== selectedRow);
    selectedRow = null;
    managerDirty = true;
  }

  function startEdit(index: number | 'new'): void {
    if (index === 'new') {
      form = defaultRuleForm();
      formFormat = {};
    } else {
      const row = rows[index];
      if (!row) return;
      form = formFromRule(row.rule, palette);
      formFormat = row.format ?? cfFormatOf(doc.wb, row.rule.dxfId, palette);
    }
    editing = index;
  }

  function finishEdit(): boolean {
    if (!formValid()) {
      notice = t('cfNeedValue');
      return false;
    }
    const target = editing;
    const base = target === 'new' || target === null ? undefined : rows[target];
    const appliesTo = base?.appliesTo ?? `=${ranges.map((r) => rangeAddress(r, true)).join(',')}`;
    const topLeft = parseAppliesTo(appliesTo)?.[0];
    const rule = makeCfRule({ ...ruleFromForm(form, topLeft ? cellAddress(topLeft.r1, topLeft.c1) : a1), priority: 0, ...(base?.rule.stopIfTrue ? { stopIfTrue: true } : {}) });
    const row: Row = { appliesTo, rule, ...(formUsesFormat(form) ? { format: formFormat } : {}) };
    if (target === 'new' || target === null) {
      rows = [row, ...rows];
      selectedRow = 0;
    } else rows = rows.map((r, i) => (i === target ? row : r));
    editing = null;
    managerDirty = true;
    return false;
  }

  function applyManager(): boolean {
    if (!managerDirty) return true;
    const drafts: DraftRule[] = [];
    for (const row of rows) {
      const parsed = parseAppliesTo(row.appliesTo);
      if (!parsed) {
        notice = t('cfInvalidRange', { ref: row.appliesTo });
        return false;
      }
      drafts.push({ ranges: parsed, rule: $state.snapshot(row.rule), ...(row.format ? { format: $state.snapshot(row.format) } : {}) });
    }
    replaceRules(doc, drafts);
    managerDirty = false;
    return true;
  }

  function describe(row: Row): string {
    const r = row.rule;
    const opKey = OPERATORS.find(([op]) => op === r.operator)?.[1];
    switch (r.type) {
      case 'cellIs': {
        const op = opKey ? t(opKey) : (r.operator ?? '');
        const a = operandText(r.formulas[0] ?? '');
        return r.formulas.length > 1 ? t('cfDescCellBetween', { op, a, b: operandText(r.formulas[1] ?? '') }) : t('cfDescCellValue', { op, a });
      }
      case 'expression':
        return t('cfDescFormula', { f: r.formulas[0] ?? '' });
      case 'colorScale':
        return t('cfDescColorScale');
      case 'dataBar':
        return t('cfDescDataBar');
      case 'iconSet':
        return t('cfDescIconSet');
      case 'top10': {
        const n = r.rank ?? 10;
        return t(r.bottom ? (r.percent ? 'cfDescBottomPercent' : 'cfDescBottom') : r.percent ? 'cfDescTopPercent' : 'cfDescTop', { n });
      }
      case 'aboveAverage':
        return t(r.aboveAverage === false ? 'cfDescBelow' : 'cfDescAbove');
      case 'containsText':
        return t('cfDescContains', { t: r.text ?? '' });
      case 'notContainsText':
        return t('cfDescNotContains', { t: r.text ?? '' });
      case 'beginsWith':
        return t('cfDescBegins', { t: r.text ?? '' });
      case 'endsWith':
        return t('cfDescEnds', { t: r.text ?? '' });
      case 'containsBlanks':
        return t('cfBlanks');
      case 'notContainsBlanks':
        return t('cfNoBlanks');
      case 'containsErrors':
        return t('cfErrors');
      case 'notContainsErrors':
        return t('cfNoErrors');
      case 'timePeriod':
        return t(PERIODS.find(([p]) => p === r.timePeriod)?.[1] ?? 'cfDatesOccurring');
      case 'duplicateValues':
        return t('cfDescDuplicate');
      case 'uniqueValues':
        return t('cfDescUnique');
    }
  }

  function rowFormat(row: Row): CfFormat {
    return row.format ?? cfFormatOf(doc.wb, row.rule.dxfId, palette);
  }

  // ---- shared bits ------------------------------------------------------------------------------

  const noHash = (c: string | undefined): string | null => (c ? c.replace('#', '') : null);
  const withHash = (c: string | null): string | undefined => (c ? `#${c}` : undefined);

  type FormatPatch = { [K in keyof CfFormat]?: CfFormat[K] | undefined };

  function setFormat(target: 'custom' | 'form', patch: FormatPatch): void {
    const m: FormatPatch = { ...(target === 'custom' ? custom : formFormat), ...patch };
    // Cleared colours and unticked boxes leave the DXF rather than being stored as "off".
    const clean: CfFormat = {
      ...(m.fill ? { fill: m.fill } : {}),
      ...(m.color ? { color: m.color } : {}),
      ...(m.bold ? { bold: true } : {}),
      ...(m.italic ? { italic: true } : {}),
      ...(m.underline ? { underline: true } : {}),
      ...(m.strike ? { strike: true } : {}),
    };
    if (target === 'custom') custom = clean;
    else formFormat = clean;
  }

  const title = $derived.by((): string => {
    if (preset === 'manage') return editing === null ? t('cfTitleManager') : t(editing === 'new' ? 'cfTitleNewRule' : 'cfTitleEditRule');
    if (preset === 'new') return t('cfTitleNewRule');
    if (preset === 'clear') return t('cfTitleClear');
    if (preset === 'dataBar') return t('cfDataBars');
    if (preset === 'colorScale') return t('cfColorScales');
    if (preset === 'iconSet') return t('cfIconSets');
    return t(QUICK_LABEL[quick]).replace('…', '');
  });

  function onok(): boolean | void {
    if (preset === 'highlight' || preset === 'topBottom') return applyQuick();
    if (preset === 'new') return applyNew();
    if (preset === 'manage') return editing === null ? applyManager() : finishEdit();
  }

  const sampleStyle = (f: CfFormat): string =>
    [
      f.fill ? `background:${f.fill}` : '',
      f.color ? `color:${f.color}` : '',
      f.bold ? 'font-weight:700' : '',
      f.italic ? 'font-style:italic' : '',
      [f.underline ? 'underline' : '', f.strike ? 'line-through' : ''].filter(Boolean).join(' ') ? `text-decoration:${[f.underline ? 'underline' : '', f.strike ? 'line-through' : ''].filter(Boolean).join(' ')}` : '',
    ]
      .filter(Boolean)
      .join(';');

  const width = preset === 'manage' ? 680 : preset === 'new' ? 520 : preset === 'highlight' || preset === 'topBottom' ? 520 : 420;
</script>

{#snippet formatEditor(target: 'custom' | 'form', f: CfFormat)}
  <fieldset>
    <div class="row">
      <span class="lbl">{t('cfFillColor')}</span>
      <ColorPicker value={noHash(f.fill)} noneLabel={t('noColor')} label={t('cfFillColor')} onchange={(rgb) => setFormat(target, { fill: withHash(rgb) })} />
    </div>
    <div class="row">
      <span class="lbl">{t('cfFontColor')}</span>
      <ColorPicker value={noHash(f.color)} noneLabel={t('automatic')} label={t('cfFontColor')} onchange={(rgb) => setFormat(target, { color: withHash(rgb) })} />
    </div>
    <div class="row">
      <label class="check"><input type="checkbox" checked={f.bold === true} onchange={(e) => setFormat(target, { bold: e.currentTarget.checked })} />{t('cfBold')}</label>
      <label class="check"><input type="checkbox" checked={f.italic === true} onchange={(e) => setFormat(target, { italic: e.currentTarget.checked })} />{t('cfItalic')}</label>
      <label class="check"><input type="checkbox" checked={f.underline === true} onchange={(e) => setFormat(target, { underline: e.currentTarget.checked })} />{t('cfUnderline')}</label>
      <label class="check"><input type="checkbox" checked={f.strike === true} onchange={(e) => setFormat(target, { strike: e.currentTarget.checked })} />{t('cfStrikethrough')}</label>
    </div>
    <div class="row">
      <span class="lbl">{t('cfPreview')}</span>
      <span class="sample grow" style={sampleStyle(f)}>{Object.keys(f).length === 0 ? t('cfNoFormatSet') : t('cfPreviewSample')}</span>
    </div>
  </fieldset>
{/snippet}

{#snippet cfvoEditor(label: MessageKey, stop: RuleForm['min'], withColor: boolean)}
  <div class="stop">
    <div class="stop-title">{t(label)}</div>
    <select class="xl-select" bind:value={stop.type} aria-label={t('cfType')}>
      {#each CFVO_TYPES as [v, k] (v)}
        <option value={v}>{t(k)}</option>
      {/each}
    </select>
    <input class="xl-input" bind:value={stop.val} disabled={stop.type === 'min' || stop.type === 'max'} aria-label={t('cfValue')} />
    {#if withColor}
      <ColorPicker value={stop.color} noneLabel={t('automatic')} label={t('cfColor')} onchange={(rgb) => (stop.color = rgb ?? 'FFFFFF')} />
    {/if}
  </div>
{/snippet}

{#snippet ruleEditor()}
  <div class="lbl-top">{t('cfSelectRuleType')}</div>
  <div class="list kinds" role="listbox" aria-label={t('cfSelectRuleType')}>
    {#each RULE_KINDS as [k, key] (k)}
      <button role="option" aria-selected={form.kind === k} onclick={() => (form.kind = k)}>► {t(key)}</button>
    {/each}
  </div>
  <div class="lbl-top">{t('cfEditDescription')}</div>
  <fieldset>
    {#if form.kind === 'values'}
      <div class="row">
        <span class="lbl">{t('cfFormatStyle')}</span>
        <select class="xl-select" bind:value={form.style}>
          {#each STYLES as [v, k] (v)}
            <option value={v}>{t(k)}</option>
          {/each}
        </select>
      </div>
      {#if form.style === 'scale2' || form.style === 'scale3'}
        <div class="stops">
          {@render cfvoEditor('cfMinimum', form.min, true)}
          {#if form.style === 'scale3'}{@render cfvoEditor('cfMidpoint', form.mid, true)}{/if}
          {@render cfvoEditor('cfMaximum', form.max, true)}
        </div>
        <div class="scale-preview" style:background="linear-gradient(90deg, #{form.min.color}{form.style === 'scale3' ? `, #${form.mid.color}` : ''}, #{form.max.color})"></div>
      {:else if form.style === 'dataBar'}
        <div class="stops">
          {@render cfvoEditor('cfMinimum', form.min, false)}
          {@render cfvoEditor('cfMaximum', form.max, false)}
        </div>
        <div class="row">
          <span class="lbl">{t('cfColor')}</span>
          <ColorPicker value={form.barColor} noneLabel={t('automatic')} label={t('cfColor')} onchange={(rgb) => (form.barColor = rgb ?? '638EC6')} />
          <label class="check"><input type="checkbox" bind:checked={form.barOnly} />{t('cfShowBarOnly')}</label>
        </div>
      {:else}
        <div class="row">
          <span class="lbl">{t('cfIconStyle')}</span>
          <select class="xl-select" bind:value={form.iconSet}>
            {#each ICON_GROUPS as [, sets] (sets[0])}
              {#each sets as s (s)}
                <option value={s}>{(ICON_SETS[s] ?? []).map((i) => i.glyph).join(' ')}</option>
              {/each}
            {/each}
          </select>
          <label class="check"><input type="checkbox" bind:checked={form.reverse} />{t('cfReverseIconOrder')}</label>
          <label class="check"><input type="checkbox" bind:checked={form.iconOnly} />{t('cfShowIconOnly')}</label>
        </div>
      {/if}
    {:else if form.kind === 'contains'}
      <div class="row wrap">
        <select class="xl-select" bind:value={form.contains}>
          {#each CONTAINS as [v, k] (v)}
            <option value={v}>{t(k)}</option>
          {/each}
        </select>
        {#if form.contains === 'cellValue'}
          <select class="xl-select" bind:value={form.operator}>
            {#each OPERATORS as [v, k] (v)}
              <option value={v}>{t(k)}</option>
            {/each}
          </select>
          <input class="xl-input" bind:value={form.value1} aria-label={t('cfValue')} />
          {#if form.operator === 'between' || form.operator === 'notBetween'}
            <span>{t('cfAnd')}</span>
            <input class="xl-input" bind:value={form.value2} aria-label={t('cfValue')} />
          {/if}
        {:else if form.contains === 'text'}
          <select class="xl-select" bind:value={form.textOp}>
            {#each TEXT_OPS as [v, k] (v)}
              <option value={v}>{t(k)}</option>
            {/each}
          </select>
          <input class="xl-input grow" bind:value={form.text} aria-label={t('cfValue')} />
        {:else if form.contains === 'date'}
          <select class="xl-select" bind:value={form.period}>
            {#each PERIODS as [v, k] (v)}
              <option value={v}>{t(k)}</option>
            {/each}
          </select>
        {/if}
      </div>
    {:else if form.kind === 'topBottom'}
      <div class="row">
        <select class="xl-select" value={form.bottom ? 'bottom' : 'top'} onchange={(e) => (form.bottom = e.currentTarget.value === 'bottom')}>
          <option value="top">{t('cfTop')}</option>
          <option value="bottom">{t('cfBottom')}</option>
        </select>
        <input class="xl-input" type="number" min="1" max={form.percent ? 100 : 1000} bind:value={form.rank} />
        <label class="check"><input type="checkbox" bind:checked={form.percent} />{t('cfPercentOfRange')}</label>
      </div>
    {:else if form.kind === 'average'}
      <div class="row">
        <select class="xl-select" bind:value={form.average}>
          {#each AVERAGE_CHOICES as a (a)}
            <option value={a}>{averageLabel(a)}</option>
          {/each}
        </select>
        <span>{t('cfTheAverage')}</span>
      </div>
    {:else if form.kind === 'unique'}
      <div class="row">
        <select class="xl-select" value={form.unique ? 'unique' : 'duplicate'} onchange={(e) => (form.unique = e.currentTarget.value === 'unique')}>
          <option value="duplicate">{t('cfDuplicate')}</option>
          <option value="unique">{t('cfUnique')}</option>
        </select>
        <span>{t('cfValuesInRange')}</span>
      </div>
    {:else}
      <div class="lbl-top">{t('cfFormulaPrompt')}</div>
      <input class="xl-input formula" bind:value={form.formula} spellcheck="false" placeholder="=$A1>0" />
    {/if}
  </fieldset>
  {#if formUsesFormat(form)}
    {@render formatEditor('form', formFormat)}
  {/if}
{/snippet}

{#snippet editorFooter()}
  <button class="xl-btn outlined" onclick={() => (editing = null)}>{t('dlgCancel')}</button>
  <button class="xl-btn primary" onclick={finishEdit}>{t('dlgOk')}</button>
{/snippet}

{#snippet galleryFooter()}
  <button class="xl-btn outlined" onclick={() => ctl.closeDialog()}>{t('dlgCancel')}</button>
{/snippet}

{#if preset === 'dataBar' || preset === 'colorScale' || preset === 'iconSet' || preset === 'clear'}
  <Dialog {title} {width} footer={galleryFooter} bind:notice>
    {#if preset === 'dataBar'}
      <div class="section">{t('cfGradientFill')}</div>
      <div class="gallery">
        {#each BAR_COLORS as [hex, key] (hex)}
          <button class="tile" title={t(key)} aria-label={t(key)} onclick={() => applyForm({ style: 'dataBar', barColor: hex, min: { type: 'min', val: '', color: hex }, max: { type: 'max', val: '', color: hex } })}>
            {#each [0.9, 0.6, 0.35] as w (w)}
              <span class="bar" style:width="{w * 100}%" style:background="linear-gradient(90deg, #{hex}, #ffffff)" style:border-color="#{hex}"></span>
            {/each}
          </button>
        {/each}
      </div>
    {:else if preset === 'colorScale'}
      <div class="gallery">
        {#each SCALES as [key, colors] (key)}
          <button class="tile scale" title={t(key)} aria-label={t(key)} onclick={() => applyScale(colors)}>
            {#each [0, 1, 2] as row (row)}
              <span class="scale-row" style:background="linear-gradient(90deg, {[...colors].reverse().map((c) => `#${c}`).join(', ')})" style:opacity={1 - row * 0.15}></span>
            {/each}
          </button>
        {/each}
      </div>
    {:else if preset === 'iconSet'}
      {#each ICON_GROUPS as [group, sets] (group)}
        <div class="section">{t(group)}</div>
        <div class="icon-grid">
          {#each sets as s (s)}
            <button class="icon-row" title={s} aria-label={s} onclick={() => applyForm({ style: 'iconSet', iconSet: s })}>
              {#each ICON_SETS[s] ?? [] as ic, i (i)}
                <span style:color={ic.color}>{ic.glyph}</span>
              {/each}
            </button>
          {/each}
        </div>
      {/each}
    {:else}
      <div class="clear">
        <button class="xl-menu-item" onclick={() => { clearRulesIn(doc, ranges); ctl.closeDialog(); }}>{t('cfClearSelected')}</button>
        <button class="xl-menu-item" onclick={() => { clearSheetRules(doc); ctl.closeDialog(); }}>{t('cfClearSheet')}</button>
      </div>
    {/if}
  </Dialog>
{:else if preset === 'highlight' || preset === 'topBottom'}
  <Dialog {title} {width} {onok} bind:notice>
    <div class="quick">
      <div class="list choices" role="listbox" aria-label={t(preset === 'highlight' ? 'cfHighlightRules' : 'cfTopBottomRules')}>
        {#each preset === 'highlight' ? HIGHLIGHT : TOP_BOTTOM as q (q)}
          <button role="option" aria-selected={quick === q} onclick={() => (quick = q)}>{t(QUICK_LABEL[q])}</button>
        {/each}
      </div>
      <div class="grow">
        <div class="lbl-top">{t(QUICK_PROMPT[quick])}</div>
        <div class="row wrap">
          {#if quick === 'greater' || quick === 'less' || quick === 'equal' || quick === 'text'}
            <input class="xl-input grow" bind:value={q1} spellcheck="false" aria-label={t('cfValue')} />
          {:else if quick === 'between'}
            <input class="xl-input grow" bind:value={q1} spellcheck="false" aria-label={t('cfValue')} />
            <span>{t('cfAnd')}</span>
            <input class="xl-input grow" bind:value={q2} spellcheck="false" aria-label={t('cfValue')} />
          {:else if quick === 'date'}
            <select class="xl-select grow" bind:value={qPeriod}>
              {#each PERIODS as [v, k] (v)}
                <option value={v}>{t(k)}</option>
              {/each}
            </select>
          {:else if quick === 'duplicate'}
            <select class="xl-select" value={qUnique ? 'unique' : 'duplicate'} onchange={(e) => (qUnique = e.currentTarget.value === 'unique')}>
              <option value="duplicate">{t('cfDuplicate')}</option>
              <option value="unique">{t('cfUnique')}</option>
            </select>
          {:else if quick === 'top' || quick === 'topPct' || quick === 'bottom' || quick === 'bottomPct'}
            <input class="xl-input" type="number" min="1" max={quick.endsWith('Pct') ? 100 : 1000} bind:value={qRank} aria-label={t('cfValue')} />
            <span>{quick.endsWith('Pct') ? '%' : ''}</span>
          {:else}
            <span>{t('cfForSelectedRange')}</span>
          {/if}
        </div>
        <div class="row">
          <span>{t('cfWith')}</span>
          <select class="xl-select grow" bind:value={presetIndex}>
            {#each PRESET_FORMATS as p, i (p.key)}
              <option value={i}>{t(p.key)}</option>
            {/each}
            <option value={CUSTOM}>{t('cfFmtCustom')}</option>
          </select>
        </div>
        {#if presetIndex === CUSTOM}
          {@render formatEditor('custom', custom)}
        {:else}
          <div class="row"><span class="sample grow" style={sampleStyle(chosenFormat)}>{t('cfPreviewSample')}</span></div>
        {/if}
      </div>
    </div>
  </Dialog>
{:else if preset === 'manage'}
  <!-- The rule editor opens inside the manager's frame so Cancel returns to the draft list. -->
  <Dialog {title} {width} {onok} bind:notice {...editing === null ? {} : { footer: editorFooter }}>
    {#if editing !== null}
      {@render ruleEditor()}
    {:else}
    <div class="row">
      <span class="lbl">{t('cfShowRulesFor')}</span>
      <select class="xl-select" bind:value={scope}>
        <option value="selection">{t('cfCurrentSelection')}</option>
        <option value="sheet">{t('cfThisWorksheet')}</option>
      </select>
    </div>
    <div class="row toolbar">
      <button class="xl-btn outlined" onclick={() => startEdit('new')}>{t('cfNewRuleButton')}</button>
      <button class="xl-btn outlined" disabled={selectedRow === null} onclick={() => selectedRow !== null && startEdit(selectedRow)}>{t('cfEditRuleButton')}</button>
      <button class="xl-btn outlined" disabled={selectedRow === null} onclick={deleteRow}>{t('cfDeleteRule')}</button>
      <button class="xl-btn outlined" disabled={selectedRow === null} onclick={() => move(-1)} aria-label={t('cfMoveUp')} title={t('cfMoveUp')}>▲</button>
      <button class="xl-btn outlined" disabled={selectedRow === null} onclick={() => move(1)} aria-label={t('cfMoveDown')} title={t('cfMoveDown')}>▼</button>
    </div>
    <div class="table" role="grid">
      <div class="thead" role="row">
        <span role="columnheader">{t('cfColRule')}</span>
        <span role="columnheader">{t('cfColFormat')}</span>
        <span role="columnheader">{t('cfColAppliesTo')}</span>
        <span role="columnheader">{t('cfColStopIfTrue')}</span>
      </div>
      {#each visibleRows as { row, index } (index)}
        {@const visual = parseVisualRule(row.rule)}
        <div class="tr" role="row" tabindex="-1" aria-selected={selectedRow === index} onpointerdown={() => (selectedRow = index)} ondblclick={() => startEdit(index)}>
          <span class="desc" role="gridcell">{describe(row)}</span>
          <span role="gridcell">
            {#if visual?.kind === 'colorScale'}
              <span class="sample swatch" style:background="linear-gradient(90deg, {visual.colors.map((c) => resolveColor(c, palette, '#FFFFFF')).join(', ')})"></span>
            {:else if visual?.kind === 'dataBar'}
              <span class="sample swatch"><span class="bar small" style:background={resolveColor(visual.color, palette, '#638EC6')}></span></span>
            {:else if visual?.kind === 'iconSet'}
              <span class="sample swatch">{#each ICON_SETS[visual.iconSet] ?? [] as ic, i (i)}<span style:color={ic.color}>{ic.glyph}</span>{/each}</span>
            {:else}
              <span class="sample swatch" style={sampleStyle(rowFormat(row))}>{t('cfPreviewSample')}</span>
            {/if}
          </span>
          <span role="gridcell"><input class="xl-input applies" value={row.appliesTo} spellcheck="false" onchange={(e) => { row.appliesTo = e.currentTarget.value; managerDirty = true; }} /></span>
          <span role="gridcell" class="center">
            {#if visual === undefined}
              <input type="checkbox" checked={row.rule.stopIfTrue === true} onchange={(e) => { row.rule.stopIfTrue = e.currentTarget.checked; managerDirty = true; }} aria-label={t('cfColStopIfTrue')} />
            {/if}
          </span>
        </div>
      {:else}
        <div class="empty">{t('cfNoRules')}</div>
      {/each}
    </div>
    {/if}
  </Dialog>
{:else}
  <Dialog {title} {width} {onok} bind:notice>
    {@render ruleEditor()}
  </Dialog>
{/if}

<style>
  .section {
    font-weight: 600;
    margin: 8px 0 4px;
  }
  .gallery {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 8px;
  }
  .tile {
    display: flex;
    flex-direction: column;
    gap: 3px;
    height: 48px;
    padding: 6px;
    border: 1px solid var(--xl-border-strong);
    border-radius: 4px;
    background: #fff;
    cursor: pointer;
  }
  .tile:hover {
    border-color: var(--xl-accent);
  }
  .bar {
    display: block;
    height: 9px;
    border: 1px solid;
  }
  .bar.small {
    width: 60%;
    height: 100%;
    border: 0;
  }
  .scale-row {
    display: block;
    flex: 1;
  }
  .icon-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 6px;
  }
  .icon-row {
    display: flex;
    justify-content: center;
    gap: 6px;
    padding: 4px;
    border: 1px solid var(--xl-border-strong);
    border-radius: 4px;
    background: #fff;
    font-size: 15px;
    cursor: pointer;
  }
  .icon-row:hover {
    border-color: var(--xl-accent);
  }
  .clear {
    display: flex;
    flex-direction: column;
  }
  .quick {
    display: flex;
    gap: 12px;
  }
  .choices {
    width: 170px;
    flex: none;
  }
  .kinds {
    height: 128px;
  }
  .lbl-top {
    margin: 6px 0 4px;
  }
  .wrap {
    flex-wrap: wrap;
  }
  .sample {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    min-height: 24px;
    padding: 0 8px;
    border: 1px solid var(--xl-border-strong);
    background: #fff;
  }
  .swatch {
    width: 110px;
    height: 20px;
    overflow: hidden;
    justify-content: flex-start;
    gap: 3px;
  }
  .stops {
    display: flex;
    gap: 8px;
  }
  .stop {
    display: flex;
    flex-direction: column;
    gap: 4px;
    flex: 1;
    min-width: 0;
  }
  .stop-title {
    font-weight: 600;
  }
  .scale-preview {
    height: 18px;
    margin: 8px 0 2px;
    border: 1px solid var(--xl-border-strong);
  }
  .formula {
    width: 100%;
    font-family: var(--xl-mono);
  }
  .toolbar {
    flex-wrap: wrap;
  }
  .table {
    border: 1px solid var(--xl-border-strong);
    background: #fff;
    border-radius: 3px;
    height: 220px;
    overflow: auto;
  }
  .thead,
  .tr {
    display: grid;
    grid-template-columns: 1.4fr 120px 1fr 90px;
    gap: 6px;
    align-items: center;
    padding: 3px 6px;
  }
  .thead {
    position: sticky;
    top: 0;
    background: #f3f3f3;
    font-weight: 600;
    border-bottom: 1px solid var(--xl-border);
  }
  .tr[aria-selected='true'] {
    background: var(--xl-selected);
  }
  .desc {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .applies {
    width: 100%;
  }
  .center {
    text-align: center;
  }
  .empty {
    padding: 12px;
    color: var(--xl-text-2);
  }
</style>
