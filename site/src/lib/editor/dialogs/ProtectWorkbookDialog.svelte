<script lang="ts">
  // Review ▸ Protect Workbook: lock the structure (sheets can't be added,
  // deleted, moved, renamed or hidden), optionally behind a password.
  import { getEditor } from '../core/context.ts';
  import { protectionHash } from '../core/sheet-password.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  const ctl = getEditor();
  let password = $state('');
  let confirm = $state('');
  let structure = $state(true);
  let busy = $state(false);
  let notice = $state<string | null>(null);

  function onok(): boolean {
    if (password && password !== confirm) {
      notice = t('dlgPrPasswordMismatch');
      return false;
    }
    busy = true;
    const own = ctl.dialog;
    // 100,000 SHA-512 rounds: protect once hashing finishes, unless the dialog was cancelled meanwhile.
    void (async () => {
      const hash = password ? await protectionHash(password) : undefined;
      if (ctl.dialog !== own) return;
      const wb = ctl.doc.wb;
      ctl.doc.transact('Protect Workbook', (tx) => {
        tx.workbook('workbookProtection');
        wb.workbookProtection = {
          lockStructure: structure,
          ...(hash
            ? { workbookAlgorithmName: hash.algorithmName, workbookHashValue: hash.hashValue, workbookSaltValue: hash.saltValue, workbookSpinCount: hash.spinCount }
            : {}),
        };
      });
      ctl.closeDialog();
    })();
    return false;
  }
</script>

<Dialog title={t('protectWorkbook')} width={360} {onok} okDisabled={busy} bind:notice>
  <label class="check"><input type="checkbox" bind:checked={structure} />{t('dlgPwStructure')}</label>
  <div class="row">
    <label for="pw-pw">{t('dlgPrPassword')}</label>
    <input id="pw-pw" class="xl-input grow" type="password" bind:value={password} autocomplete="new-password" />
  </div>
  <div class="row">
    <label for="pw-pw2">{t('dlgPrConfirm')}</label>
    <input id="pw-pw2" class="xl-input grow" type="password" bind:value={confirm} autocomplete="new-password" disabled={!password} />
  </div>
  {#if busy}<p class="hint">{t('dlgPrHashing')}</p>{/if}
</Dialog>
