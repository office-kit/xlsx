<script lang="ts">
  // Formulas ▸ Evaluate Formula: step through the active cell's formula, the
  // part evaluated next underlined, as Excel's dialog does.
  import { CalcParseError } from '../calc/index.ts';
  import { cellAddress, quoteSheetName } from '../core/address.ts';
  import { getEditor } from '../core/context.ts';
  import { startEvaluation, stepEvaluation, type EvaluationState } from '../core/evaluate-formula.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const ctl = getEditor();
  const doc = ctl.doc;
  const sheet = doc.ws.title;
  const { row, col } = doc.selection.active;
  const value = ctl.cell(row, col)?.value;
  const formula = value !== null && typeof value === 'object' && !(value instanceof Date) && value.kind === 'formula' ? value.formula : undefined;

  function begin(): EvaluationState | undefined {
    if (formula === undefined) return undefined;
    try {
      return startEvaluation(formula);
    } catch (e) {
      // The stored formula doesn't parse (an unsupported construct): nothing to step through.
      if (e instanceof CalcParseError) return undefined;
      throw e;
    }
  }

  let evaluation = $state.raw(begin());
  const reference = `${quoteSheetName(sheet)}!${cellAddress(row, col, true)}`;

  function step() {
    if (evaluation) evaluation = stepEvaluation(evaluation, (f) => doc.calc.evaluateArray(f, sheet, row, col));
  }
</script>

<Dialog title={t('evaluateFormulaTitle')} width={460}>
  <div class="grid">
    <div class="lbl">{t('dlgEvalReference')}</div>
    <div class="lbl">{t('dlgEvalEvaluation')}</div>
    <div class="ref">{reference}</div>
    <div class="box" aria-live="polite">
      {#if evaluation}
        ={evaluation.text.slice(0, evaluation.start)}<u>{evaluation.text.slice(evaluation.start, evaluation.end)}</u>{evaluation.text.slice(evaluation.end)}
      {:else}
        <span class="hint">{t('dlgEvalNoFormula')}</span>
      {/if}
    </div>
  </div>
  <p class="hint">{t('dlgEvalHint')}</p>
  {#snippet footer()}
    <button class="xl-btn outlined" disabled={!evaluation || evaluation.done} onclick={step}>{t('dlgEvalEvaluate')}</button>
    <button class="xl-btn outlined" disabled={!evaluation} onclick={() => (evaluation = begin())}>{t('dlgEvalRestart')}</button>
    <span class="spacer"></span>
    <button class="xl-btn primary" onclick={() => ctl.closeDialog()}>{t('dlgClose')}</button>
  {/snippet}
</Dialog>

<style>
  .grid {
    display: grid;
    grid-template-columns: 110px 1fr;
    gap: 4px 10px;
  }
  .ref {
    padding: 4px 0;
  }
  .box {
    min-height: 70px;
    max-height: 160px;
    overflow: auto;
    padding: 4px 6px;
    border: 1px solid var(--xl-border);
    background: #fff;
    font-family: Menlo, Consolas, monospace;
    font-size: 12px;
    word-break: break-all;
  }
  .spacer {
    flex: 1;
  }
</style>
