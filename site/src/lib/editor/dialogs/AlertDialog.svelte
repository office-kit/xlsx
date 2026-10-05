<script lang="ts">
  import { en } from '../i18n/en.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  let { props }: { props: Record<string, unknown> | undefined } = $props();

  function isMessageKey(v: unknown): v is MessageKey {
    return typeof v === 'string' && Object.hasOwn(en, v);
  }

  const message = $derived(isMessageKey(props?.['message']) ? t(props['message']) : String(props?.['message'] ?? ''));
  const detail = $derived(typeof props?.['detail'] === 'string' ? props['detail'] : '');
  // A question (Yes / No) instead of a notice when the opener passes what Yes does.
  const confirm = $derived(typeof props?.['onConfirm'] === 'function' ? props['onConfirm'] : undefined);
</script>

<Dialog
  title={t('dlgAlertTitle')}
  width={340}
  showCancel={confirm !== undefined}
  {...confirm ? { okLabel: t('dvYes'), cancelLabel: t('dvNo') } : {}}
  onok={() => {
    confirm?.();
  }}
>
  <div class="alert">
    <div class="icon" aria-hidden="true">!</div>
    <div>
      <p class="msg">{message}</p>
      {#if detail}<p class="detail">{detail}</p>{/if}
    </div>
  </div>
</Dialog>

<style>
  .alert {
    display: flex;
    gap: 12px;
    align-items: flex-start;
  }
  .icon {
    flex: none;
    width: 36px;
    height: 36px;
    border-radius: 8px;
    background: var(--xl-accent);
    color: #fff;
    font-weight: 700;
    font-size: 20px;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .msg {
    margin: 2px 0 6px;
    font-weight: 600;
    line-height: 1.4;
  }
  .detail {
    margin: 0;
    color: var(--xl-text-2);
    white-space: pre-line;
    word-break: break-word;
  }
</style>
