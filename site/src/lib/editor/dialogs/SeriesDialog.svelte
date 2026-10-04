<script lang="ts">
  import { isDateFormat } from '@office-kit/xlsx/styles';
  import { getCellAt } from '../core/cells.ts';
  import { getEditor } from '../core/context.ts';
  import { parseInput } from '../core/input.ts';
  import { currentRange } from '../core/selection.ts';
  import { fillSeries, type DateUnit, type SeriesType } from '../core/series.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const TYPES: ReadonlyArray<[SeriesType, MessageKey]> = [
    ['linear', 'dlgSeriesLinear'],
    ['growth', 'dlgSeriesGrowth'],
    ['date', 'dlgSeriesDate'],
    ['autofill', 'dlgSeriesAutoFill'],
  ];
  const UNITS: ReadonlyArray<[DateUnit, MessageKey]> = [
    ['day', 'dlgUnitDay'],
    ['weekday', 'dlgUnitWeekday'],
    ['month', 'dlgUnitMonth'],
    ['year', 'dlgUnitYear'],
  ];

  const ctl = getEditor();
  const doc = ctl.doc;
  const range = currentRange(doc.selection);
  const first = getCellAt(doc.ws, range.r1, range.c1);
  const startsWithDate = first !== undefined && isDateFormat(doc.styles.get(first.styleId).numFmt);
  // Excel runs the series along the selection's longer side.
  let byRows = $state(range.c2 - range.c1 > range.r2 - range.r1);
  let type = $state<SeriesType>(startsWithDate ? 'date' : 'linear');
  let unit = $state<DateUnit>('day');
  let step = $state('1');
  let stop = $state('');
  let notice = $state<string | null>(null);

  function onok(): boolean {
    const stepValue = Number(step.trim());
    if (step.trim() === '' || !Number.isFinite(stepValue)) {
      notice = t('dlgSeriesBadStep');
      return false;
    }
    let stopValue: number | undefined;
    if (stop.trim() !== '') {
      const parsed = parseInput(stop.trim(), { dateOrder: ctl.dateOrder(), date1904: doc.wb.date1904 }).value;
      if (typeof parsed !== 'number') {
        notice = t('dlgSeriesBadStop');
        return false;
      }
      stopValue = parsed;
    }
    fillSeries(ctl, range, byRows, { type, step: stepValue, dateUnit: unit, ...(stopValue !== undefined ? { stop: stopValue } : {}) });
    return true;
  }
</script>

<Dialog title={t('dlgSeries')} width={400} {onok} bind:notice>
  <div class="cols">
    <fieldset>
      <legend>{t('dlgSeriesIn')}</legend>
      <label class="check"><input type="radio" name="sr-in" value={true} bind:group={byRows} />{t('dlgRows')}</label>
      <label class="check"><input type="radio" name="sr-in" value={false} bind:group={byRows} />{t('dlgColumns')}</label>
    </fieldset>
    <fieldset>
      <legend>{t('dlgType')}</legend>
      {#each TYPES as [value, label] (value)}
        <label class="check"><input type="radio" name="sr-type" {value} bind:group={type} />{t(label)}</label>
      {/each}
    </fieldset>
    <fieldset>
      <legend>{t('dlgDateUnit')}</legend>
      {#each UNITS as [value, label] (value)}
        <label class="check"><input type="radio" name="sr-unit" {value} bind:group={unit} disabled={type !== 'date'} />{t(label)}</label>
      {/each}
    </fieldset>
  </div>
  <div class="row">
    <label for="sr-step">{t('dlgStepValue')}</label>
    <input id="sr-step" class="xl-input" bind:value={step} disabled={type === 'autofill'} inputmode="decimal" />
    <label for="sr-stop" class="stop">{t('dlgStopValue')}</label>
    <input id="sr-stop" class="xl-input" bind:value={stop} disabled={type === 'autofill'} />
  </div>
</Dialog>

<style>
  .cols {
    display: grid;
    grid-template-columns: 1fr 1fr 1fr;
    gap: 8px;
  }
  .cols fieldset {
    margin: 0;
  }
  .row input {
    width: 80px;
  }
  .stop {
    margin-left: 12px;
  }
</style>
