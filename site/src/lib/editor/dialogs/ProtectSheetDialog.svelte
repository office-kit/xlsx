<script lang="ts">
  import { makeSheetProtection, type SheetProtection } from '@office-kit/xlsx/worksheet';
  import { getEditor } from '../core/context.ts';
  import { protectionHash } from '../core/sheet-password.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  type Flag = Exclude<keyof SheetProtection, 'sheet' | 'saltValue' | 'spinCount' | 'algorithmName' | 'hashValue'>;

  // Excel's allow-list, in its order. Each attribute is true when the action is *locked*.
  const ALLOW: ReadonlyArray<[Flag, MessageKey]> = [
    ['selectLockedCells', 'dlgPrSelectLocked'],
    ['selectUnlockedCells', 'dlgPrSelectUnlocked'],
    ['formatCells', 'dlgPrFormatCells'],
    ['formatColumns', 'dlgPrFormatColumns'],
    ['formatRows', 'dlgPrFormatRows'],
    ['insertColumns', 'dlgPrInsertColumns'],
    ['insertRows', 'dlgPrInsertRows'],
    ['insertHyperlinks', 'dlgPrInsertHyperlinks'],
    ['deleteColumns', 'dlgPrDeleteColumns'],
    ['deleteRows', 'dlgPrDeleteRows'],
    ['sort', 'sort'],
    ['autoFilter', 'dlgPrAutoFilter'],
    ['pivotTables', 'dlgPrPivotTables'],
    ['objects', 'dlgPrObjects'],
    ['scenarios', 'dlgPrScenarios'],
  ];

  const ctl = getEditor();
  let password = $state('');
  let confirm = $state('');
  const allowed = $state<Record<Flag, boolean>>({
    selectLockedCells: true,
    selectUnlockedCells: true,
    formatCells: false,
    formatColumns: false,
    formatRows: false,
    insertColumns: false,
    insertRows: false,
    insertHyperlinks: false,
    deleteColumns: false,
    deleteRows: false,
    sort: false,
    autoFilter: false,
    pivotTables: false,
    objects: false,
    scenarios: false,
  });
  let busy = $state(false);
  let notice = $state<string | null>(null);

  function onok(): boolean {
    if (password && password !== confirm) {
      notice = t('dlgPrPasswordMismatch');
      return false;
    }
    busy = true;
    const own = ctl.dialog;
    const ws = ctl.doc.ws;
    // Hashing runs 100,000 SHA-512 rounds, so the sheet is protected once it
    // finishes — unless the dialog was cancelled meanwhile.
    void (async () => {
      const hash = password ? await protectionHash(password) : undefined;
      if (ctl.dialog !== own) return;
      const flags: SheetProtection = { sheet: true };
      for (const [flag] of ALLOW) flags[flag] = !allowed[flag];
      ctl.doc.transact('Protect Sheet', (tx) => {
        tx.sheet(ws, 'sheetProtection');
        ws.sheetProtection = makeSheetProtection({ ...flags, ...hash });
      });
      ctl.closeDialog();
    })();
    return false;
  }
</script>

<Dialog title={t('protectSheet')} width={380} {onok} okDisabled={busy} bind:notice>
  <div class="row">
    <label for="pr-pw">{t('dlgPrPassword')}</label>
    <input id="pr-pw" class="xl-input grow" type="password" bind:value={password} autocomplete="new-password" />
  </div>
  <div class="row">
    <label for="pr-pw2">{t('dlgPrConfirm')}</label>
    <input id="pr-pw2" class="xl-input grow" type="password" bind:value={confirm} autocomplete="new-password" disabled={!password} />
  </div>
  <div class="lbl">{t('dlgPrAllowAll')}</div>
  <div class="list allow">
    {#each ALLOW as [flag, label] (flag)}
      <label class="check"><input type="checkbox" bind:checked={allowed[flag]} />{t(label)}</label>
    {/each}
  </div>
  {#if busy}<p class="hint">{t('dlgPrHashing')}</p>{/if}
</Dialog>

<style>
  .allow {
    height: 200px;
    padding: 2px 6px;
  }
</style>
