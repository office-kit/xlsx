<script lang="ts">
  import { untrack } from 'svelte';
  import { deleteLines, insertLines, shiftCells } from '../core/actions.ts';
  import { getEditor } from '../core/context.ts';
  import { t, type MessageKey } from '../i18n/i18n.svelte.ts';
  import Dialog from './Dialog.svelte';

  let { props }: { props: Record<string, unknown> | undefined } = $props();
  // DialogHost re-creates a dialog for every opening, so its arguments are read once.
  const args = untrack(() => props) ?? {};

  type Choice = 'right' | 'down' | 'left' | 'up' | 'row' | 'col';

  const ctl = getEditor();
  const inserting = args['kind'] !== 'deleteCells';
  const OPTIONS: ReadonlyArray<[Choice, MessageKey]> = inserting
    ? [
        ['right', 'dlgShiftRight'],
        ['down', 'dlgShiftDown'],
        ['row', 'dlgEntireRow'],
        ['col', 'dlgEntireColumn'],
      ]
    : [
        ['left', 'dlgShiftLeft'],
        ['up', 'dlgShiftUp'],
        ['row', 'dlgEntireRow'],
        ['col', 'dlgEntireColumn'],
      ];
  let choice = $state<Choice>(inserting ? 'down' : 'up');

  function onok(): void {
    if (choice === 'row' || choice === 'col') (inserting ? insertLines : deleteLines)(ctl, choice);
    else shiftCells(ctl, choice);
  }
</script>

<Dialog title={t(inserting ? 'dlgInsertCells' : 'dlgDeleteCells')} width={260} {onok}>
  <div role="radiogroup" aria-label={t(inserting ? 'dlgInsertCells' : 'dlgDeleteCells')}>
    {#each OPTIONS as [value, label] (value)}
      <label class="check"><input type="radio" name="idc" {value} bind:group={choice} />{t(label)}</label>
    {/each}
  </div>
</Dialog>
