<script lang="ts">
  // Excel's data-validation error alert. Stop offers Retry/Cancel; Warning asks
  // "Continue?" (Yes keeps the entry, No returns to editing); Information
  // offers OK (keep) / Cancel (drop).
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  const ctl = getEditor();
  const p = $derived(ctl.validationPrompt);
  let first = $state<HTMLButtonElement>();
  $effect(() => {
    if (p) first?.focus();
  });
</script>

{#if p}
  <div class="backdrop">
    <div class="box" role="alertdialog" aria-modal="true" aria-labelledby="xl-dv-title" tabindex="-1" onkeydown={(e) => {
      e.stopPropagation();
      if (e.key === 'Escape') ctl.resolveValidation(p.style === 'stop' ? 'retry' : 'cancel');
    }}>
      <div class="head" id="xl-dv-title">{p.title ?? 'Microsoft Excel'}</div>
      <div class="body">
        <div class="icon {p.style}" aria-hidden="true">{p.style === 'stop' ? '×' : p.style === 'warning' ? '!' : 'i'}</div>
        <div>
          <p>{p.message ?? t('dvDefaultError')}</p>
          {#if p.style === 'warning'}<p>{t('dvContinue')}</p>{/if}
        </div>
      </div>
      <div class="footer">
        {#if p.style === 'stop'}
          <button class="xl-btn primary" bind:this={first} onclick={() => ctl.resolveValidation('retry')}>{t('dvRetry')}</button>
          <button class="xl-btn" onclick={() => ctl.resolveValidation('cancel')}>{t('cancel')}</button>
        {:else if p.style === 'warning'}
          <button class="xl-btn primary" bind:this={first} onclick={() => ctl.resolveValidation('commit')}>{t('dvYes')}</button>
          <button class="xl-btn" onclick={() => ctl.resolveValidation('retry')}>{t('dvNo')}</button>
          <button class="xl-btn" onclick={() => ctl.resolveValidation('cancel')}>{t('cancel')}</button>
        {:else}
          <button class="xl-btn primary" bind:this={first} onclick={() => ctl.resolveValidation('commit')}>{t('ok')}</button>
          <button class="xl-btn" onclick={() => ctl.resolveValidation('cancel')}>{t('cancel')}</button>
        {/if}
      </div>
    </div>
  </div>
{/if}

<style>
  .backdrop {
    position: fixed;
    inset: 0;
    z-index: 90;
    display: grid;
    place-items: center;
    background: rgb(0 0 0 / 0.12);
  }
  .box {
    width: 380px;
    background: var(--xl-bg);
    border: 1px solid var(--xl-border);
    border-radius: 8px;
    box-shadow: 0 10px 30px rgb(0 0 0 / 0.25);
    padding: 14px 16px 12px;
    font-size: 12.5px;
  }
  .head {
    font-weight: 600;
    margin-bottom: 10px;
  }
  .body {
    display: flex;
    gap: 12px;
  }
  .body p {
    margin: 2px 0 6px;
    white-space: pre-line;
    word-break: break-word;
  }
  .icon {
    flex: none;
    width: 32px;
    height: 32px;
    border-radius: 50%;
    display: grid;
    place-items: center;
    color: #fff;
    font-weight: 700;
    font-size: 18px;
  }
  .icon.stop {
    background: #d13438;
  }
  .icon.warning {
    background: #e3a21a;
  }
  .icon.information {
    background: var(--xl-accent);
  }
  .footer {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 10px;
  }
</style>
