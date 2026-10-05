<script lang="ts">
  // The password prompt for Unprotect Sheet / Unprotect Workbook. The
  // password is checked against the stored SHA-512 hash; a wrong one keeps
  // the protection, as in Excel.
  import { getEditor } from '../core/context.ts';
  import { hashSheetPassword } from '../core/sheet-password.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  let { props }: { props: Record<string, unknown> | undefined } = $props();

  const ctl = getEditor();
  const target = props?.['target'] === 'workbook' ? 'workbook' : 'sheet';
  let password = $state('');
  let busy = $state(false);
  let notice = $state<string | null>(null);

  function stored(): { salt: string; hash: string; spin: number } | undefined {
    if (target === 'workbook') {
      const p = ctl.doc.wb.workbookProtection;
      return p?.workbookHashValue && p.workbookSaltValue ? { salt: p.workbookSaltValue, hash: p.workbookHashValue, spin: p.workbookSpinCount ?? 0 } : undefined;
    }
    const p = ctl.doc.ws.sheetProtection;
    return p?.hashValue && p.saltValue ? { salt: p.saltValue, hash: p.hashValue, spin: p.spinCount ?? 0 } : undefined;
  }

  function onok(): boolean {
    const s = stored();
    busy = true;
    const own = ctl.dialog;
    const ws = ctl.doc.ws;
    const wb = ctl.doc.wb;
    void (async () => {
      const ok = !s || (await hashSheetPassword(password, s.salt, s.spin)) === s.hash;
      if (ctl.dialog !== own) return;
      busy = false;
      if (!ok) {
        notice = t('dlgUnprotectWrong');
        return;
      }
      if (target === 'workbook') {
        ctl.doc.transact('Unprotect Workbook', (tx) => {
          tx.workbook('workbookProtection');
          delete wb.workbookProtection;
        });
      } else {
        ctl.doc.transact('Unprotect Sheet', (tx) => {
          tx.sheet(ws, 'sheetProtection');
          delete ws.sheetProtection;
        });
      }
      ctl.closeDialog();
    })();
    return false;
  }
</script>

<Dialog title={target === 'workbook' ? t('unprotectWorkbook') : t('unprotectSheet')} width={320} {onok} okDisabled={busy} bind:notice>
  <div class="row">
    <label for="up-pw">{t('dlgPrPassword')}</label>
    <input id="up-pw" class="xl-input grow" type="password" bind:value={password} autocomplete="current-password" />
  </div>
</Dialog>
